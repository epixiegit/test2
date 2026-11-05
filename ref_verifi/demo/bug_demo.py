"""의도적으로 틀린 코드. 기본 테스트에는 포함하지 않음."""
import unittest
from analysis import load_rows


def buggy_total(rows):
    # 버그: 취소 건을 걸러내지 않았다.
    return sum(int(row["amount"]) for row in rows)


class CancelledAmountBug(unittest.TestCase):
    def test_cancelled_request_must_be_excluded(self):
        self.assertEqual(buggy_total(load_rows()), 35000,
                         "취소된 9,000원을 포함해 44,000원이 나옵니다. 기대값 35,000원을 바꾸지 말고 필터를 확인하세요.")
