import copy
import json
import unittest
from pathlib import Path

from analysis import DataValidationError, load_rows, summarize, validate_rows

EXPECTED_PATH = Path(__file__).resolve().parents[1] / "data" / "expected_summary.json"


class AnalysisTests(unittest.TestCase):
    def setUp(self):
        self.rows = load_rows()

    def test_fixture_has_six_unique_requests(self):
        self.assertEqual(len(self.rows), 6)
        self.assertEqual(len({row["request_id"] for row in self.rows}), 6)

    def test_result_matches_hand_calculated_answer(self):
        expected = json.loads(EXPECTED_PATH.read_text(encoding="utf-8-sig"))
        self.assertEqual(summarize(self.rows), expected)

    def test_cancelled_amount_does_not_change_result(self):
        before = summarize(self.rows)
        altered = copy.deepcopy(self.rows)
        altered[3]["amount"] = "999999"
        self.assertEqual(summarize(altered), before)

    def test_quality_filter(self):
        result = summarize(self.rows, "품질")
        self.assertEqual((result["record_count"], result["total_amount"]), (1, 5000))
        self.assertEqual(result["by_department"], [{"department": "품질", "count": 1, "amount": 5000}])

    def test_unknown_department_is_empty(self):
        result = summarize(self.rows, "없는부서")
        self.assertEqual((result["record_count"], result["total_amount"], result["by_department"]), (0, 0, []))

    def test_empty_input(self):
        self.assertEqual(summarize([])["total_amount"], 0)
        self.assertEqual(summarize([])["record_count"], 0)

    def test_zero_amount_is_valid_and_counted(self):
        self.rows[0]["amount"] = 0
        result = summarize(self.rows)
        self.assertEqual((result["record_count"], result["total_amount"]), (5, 23000))

    def test_duplicate_identifier_is_rejected(self):
        with self.assertRaisesRegex(DataValidationError, "중복"):
            summarize(self.rows + [copy.deepcopy(self.rows[0])])

    def test_missing_required_column_is_rejected(self):
        del self.rows[0]["department"]
        with self.assertRaises(DataValidationError):
            summarize(self.rows)

    def test_blank_department_is_rejected(self):
        self.rows[0]["department"] = "  "
        with self.assertRaises(DataValidationError):
            summarize(self.rows)

    def test_invalid_amounts_are_rejected(self):
        for value in (-1, "-1", "1.5", 1.5, True, None, "", "1,000", "NaN"):
            with self.subTest(amount=value):
                rows = copy.deepcopy(self.rows)
                rows[0]["amount"] = value
                with self.assertRaises(DataValidationError):
                    summarize(rows)

    def test_unknown_status_is_rejected(self):
        self.rows[0]["status"] = "maybe"
        with self.assertRaises(DataValidationError):
            summarize(self.rows)

    def test_input_is_not_modified(self):
        before = copy.deepcopy(self.rows)
        summarize(self.rows)
        self.assertEqual(self.rows, before)

    def test_order_does_not_change_result(self):
        self.assertEqual(summarize(self.rows[::-1]), summarize(self.rows))

    def test_identifier_leading_zero_is_preserved(self):
        self.rows[0]["request_id"] = "0001"
        self.assertEqual(validate_rows(self.rows)[0]["request_id"], "0001")

    def test_adding_approved_request_changes_total_by_its_amount(self):
        result = summarize(self.rows + [{"request_id": "R007", "department": "생산", "amount": "4000", "status": "approved"}])
        self.assertEqual((result["record_count"], result["total_amount"]), (6, 39000))
