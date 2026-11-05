"""검증 실습용 표준 라이브러리 서버. 운영용 서버나 FastAPI 구현은 아님."""
from __future__ import annotations

import argparse
import csv
import io
import json
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlsplit

from analysis import DATA_PATH, DataValidationError, load_rows, summarize

HTML_PATH = Path(__file__).resolve().parent / "index.html"


def make_server(host: str = "127.0.0.1", port: int = 0, data_path: Path = DATA_PATH):
    class Handler(BaseHTTPRequestHandler):
        def log_message(self, format, *args):
            # 자동 테스트에서는 조용히, 직접 실행 시에는 --verbose로 로그 확인.
            if self.server.verbose:
                super().log_message(format, *args)

        def send_bytes(self, status, body, content_type, attachment=False):
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            if attachment:
                self.send_header("Content-Disposition", 'attachment; filename="summary.csv"')
            self.end_headers()
            self.wfile.write(body)

        def send_json(self, status, value):
            body = json.dumps(value, ensure_ascii=False, allow_nan=False).encode("utf-8")
            self.send_bytes(status, body, "application/json; charset=utf-8")

        def error(self, status, code, message):
            self.send_json(status, {"error": {"code": code, "message": message}})

        def get_filter(self, query):
            params = parse_qs(query, keep_blank_values=True)
            if set(params) - {"department"}:
                raise ValueError("지원하는 조회 조건은 department뿐입니다.")
            values = params.get("department", [""])
            if len(values) != 1 or len(values[0]) > 80:
                raise ValueError("department는 80자 이하의 값 하나만 허용합니다.")
            return values[0].strip() or None

        def do_GET(self):
            url = urlsplit(self.path)
            if url.path == "/health":
                return self.send_json(200, {"status": "ok"})
            if url.path in ("/", "/index.html"):
                return self.send_bytes(200, HTML_PATH.read_bytes(), "text/html; charset=utf-8")
            if url.path == "/favicon.ico":
                return self.send_bytes(204, b"", "image/x-icon")
            if url.path not in ("/api/summary", "/api/export.csv"):
                return self.error(404, "NOT_FOUND", "존재하지 않는 경로입니다.")
            try:
                department = self.get_filter(url.query)
            except ValueError as exc:
                return self.error(400, "BAD_QUERY", str(exc))
            try:
                result = summarize(load_rows(data_path), department)
            except (OSError, DataValidationError):
                return self.error(500, "SOURCE_DATA_ERROR", "원본 데이터 검증에 실패했습니다. 서버의 데이터 파일을 확인하세요.")
            if url.path == "/api/summary":
                return self.send_json(200, result)
            output = io.StringIO(newline="")
            writer = csv.DictWriter(output, fieldnames=["department", "count", "amount"], lineterminator="\r\n")
            writer.writeheader()
            writer.writerows(result["by_department"])
            return self.send_bytes(200, output.getvalue().encode("utf-8-sig"), "text/csv; charset=utf-8", attachment=True)

        def do_POST(self):
            route = urlsplit(self.path).path
            if route != "/api/preview":
                return self.error(405 if route in ("/health", "/api/summary", "/api/export.csv") else 404,
                                  "METHOD_OR_PATH", "요청 메서드 또는 경로가 맞지 않습니다.")
            if self.headers.get("Content-Type", "").split(";", 1)[0].strip() != "application/json":
                return self.error(415, "CONTENT_TYPE", "application/json을 사용하세요.")
            try:
                length = int(self.headers.get("Content-Length", "0"))
            except ValueError:
                return self.error(400, "BAD_LENGTH", "Content-Length가 잘못되었습니다.")
            if not 0 < length <= 1_000_000:
                return self.error(413 if length > 1_000_000 else 400, "BODY_SIZE", "JSON 본문은 1~1000000바이트여야 합니다.")
            try:
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
            except (UnicodeDecodeError, json.JSONDecodeError):
                return self.error(400, "BAD_JSON", "JSON 형식이 잘못되었습니다.")
            try:
                if not isinstance(payload, dict) or "rows" not in payload:
                    raise DataValidationError("rows 항목이 필요합니다.")
                result = summarize(payload["rows"])
            except DataValidationError as exc:
                return self.error(422, "INVALID_DATA", str(exc))
            return self.send_json(200, result)

    server = ThreadingHTTPServer((host, port), Handler)
    server.verbose = False
    return server


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()
    server = make_server(args.host, args.port)
    server.verbose = args.verbose
    print(f"Binding: {args.host}:{server.server_port}", flush=True)
    print(f"Local browser: http://127.0.0.1:{server.server_port}/", flush=True)
    print("Web IDE: use the platform's preview for this port. Stop: Ctrl+C.", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
