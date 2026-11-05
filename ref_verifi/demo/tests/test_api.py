import csv
import io
import json
import threading
import unittest
from urllib.error import HTTPError
from urllib.parse import urlencode
from urllib.request import Request, build_opener, ProxyHandler

from analysis import load_rows
from server import make_server


class ApiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # 0번 포트는 운영체제가 빈 포트를 선택. 다른 개발 서버를 건드리지 않음.
        cls.server = make_server(port=0)
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()
        cls.base = f"http://127.0.0.1:{cls.server.server_port}"
        cls.opener = build_opener(ProxyHandler({}))

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.thread.join(timeout=5)

    def request(self, route, method="GET", body=None, content_type="application/json"):
        request = Request(self.base + route, data=body, method=method, headers={"Content-Type": content_type})
        try:
            response = self.opener.open(request, timeout=5)
        except HTTPError as exc:
            response = exc
        with response:
            return response.status, response.read(), response.headers

    def test_health(self):
        status, body, _ = self.request("/health")
        self.assertEqual(status, 200)
        self.assertEqual(json.loads(body), {"status": "ok"})

    def test_total_is_correct_not_just_http_200(self):
        status, body, headers = self.request("/api/summary")
        self.assertEqual(status, 200)
        self.assertIn("application/json", headers["Content-Type"])
        data = json.loads(body)
        self.assertEqual((data["record_count"], data["total_amount"]), (5, 35000))
        self.assertEqual({x["department"]: x["amount"] for x in data["by_department"]}, {"생산": 20000, "품질": 5000, "영업": 10000})

    def test_quality_filter(self):
        status, body, _ = self.request("/api/summary?" + urlencode({"department": "품질"}))
        self.assertEqual(status, 200)
        self.assertEqual(json.loads(body)["total_amount"], 5000)

    def test_unknown_department_returns_empty_data(self):
        status, body, _ = self.request("/api/summary?" + urlencode({"department": "없는부서"}))
        self.assertEqual(status, 200)
        data = json.loads(body)
        self.assertEqual((data["record_count"], data["total_amount"], data["by_department"]), (0, 0, []))

    def test_bad_query(self):
        status, body, _ = self.request("/api/summary?unexpected=yes")
        self.assertEqual(status, 400)
        self.assertEqual(json.loads(body)["error"]["code"], "BAD_QUERY")

    def test_missing_route(self):
        self.assertEqual(self.request("/does-not-exist")[0], 404)

    def test_wrong_method(self):
        self.assertEqual(self.request("/health", "POST", b"{}")[0], 405)

    def test_preview_valid_input(self):
        payload = json.dumps({"rows": load_rows()}, ensure_ascii=False).encode("utf-8")
        status, body, _ = self.request("/api/preview", "POST", payload)
        self.assertEqual(status, 200)
        self.assertEqual(json.loads(body)["total_amount"], 35000)

    def test_preview_invalid_data_is_422(self):
        rows = load_rows()
        rows[0]["amount"] = "-10"
        status, body, _ = self.request("/api/preview", "POST", json.dumps({"rows": rows}).encode())
        self.assertEqual(status, 422)
        self.assertEqual(json.loads(body)["error"]["code"], "INVALID_DATA")

    def test_preview_broken_json_is_400(self):
        status, body, _ = self.request("/api/preview", "POST", b"{broken")
        self.assertEqual(status, 400)
        self.assertEqual(json.loads(body)["error"]["code"], "BAD_JSON")

    def test_preview_missing_rows_is_422(self):
        self.assertEqual(self.request("/api/preview", "POST", b"{}")[0], 422)

    def test_preview_wrong_content_type_is_415(self):
        self.assertEqual(self.request("/api/preview", "POST", b"{}", "text/plain")[0], 415)

    def test_csv_has_bom_and_correct_values(self):
        status, body, headers = self.request("/api/export.csv")
        self.assertEqual(status, 200)
        self.assertTrue(body.startswith(b"\xef\xbb\xbf"))
        self.assertIn("attachment", headers["Content-Disposition"])
        rows = list(csv.DictReader(io.StringIO(body.decode("utf-8-sig"))))
        self.assertEqual({row["department"]: int(row["amount"]) for row in rows}, {"생산": 20000, "품질": 5000, "영업": 10000})
        self.assertEqual(sum(int(row["count"]) for row in rows), 5)

    def test_csv_uses_selected_filter(self):
        status, body, _ = self.request("/api/export.csv?" + urlencode({"department": "품질"}))
        self.assertEqual(status, 200)
        rows = list(csv.DictReader(io.StringIO(body.decode("utf-8-sig"))))
        self.assertEqual(rows, [{"department": "품질", "count": "1", "amount": "5000"}])

    def test_csv_empty_result_keeps_headers(self):
        status, body, _ = self.request("/api/export.csv?" + urlencode({"department": "없는부서"}))
        self.assertEqual(status, 200)
        self.assertEqual(list(csv.reader(io.StringIO(body.decode("utf-8-sig")))), [["department", "count", "amount"]])

    def test_home_html_is_served(self):
        status, body, headers = self.request("/")
        self.assertEqual(status, 200)
        self.assertIn("text/html", headers["Content-Type"])
        self.assertIn("부서별 구매금액 집계", body.decode("utf-8"))
        # 이 검사는 HTML 응답만 확인. JS 실행·버튼 클릭 검증은 별도 브라우저에서.
