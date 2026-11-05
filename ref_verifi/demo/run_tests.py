"""표준 라이브러리만 사용하는 재실행 가능한 검증 진입점."""
from __future__ import annotations
import argparse
import contextlib
import io
import json
import sys
import unittest
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent
KIT = ROOT.parent
sys.dont_write_bytecode = True


class RecordedResult(unittest.TextTestResult):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.records = []

    def addSuccess(self, test):
        super().addSuccess(test)
        self.records.append({"test": test.id(), "status": "PASS"})

    def addFailure(self, test, err):
        super().addFailure(test, err)
        self.records.append({"test": test.id(), "status": "FAIL", "detail": self._exc_info_to_string(err, test)})

    def addError(self, test, err):
        super().addError(test, err)
        self.records.append({"test": test.id(), "status": "ERROR", "detail": self._exc_info_to_string(err, test)})

    def addSkip(self, test, reason):
        super().addSkip(test, reason)
        self.records.append({"test": test.id(), "status": "SKIP", "detail": reason})

    def addSubTest(self, test, subtest, err):
        super().addSubTest(test, subtest, err)
        if err is not None:
            status = "FAIL" if issubclass(err[0], test.failureException) else "ERROR"
            self.records.append({"test": subtest.id(), "status": status,
                                 "detail": self._exc_info_to_string(err, test)})


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--group", choices=["all", "analysis", "api"], default="all")
    parser.add_argument("--bug-demo", action="store_true", help="취소 건을 포함하는 의도적인 실패 예제")
    parser.add_argument("--report", action="store_true", help="실행 기록을 ref_verifi/.runs 아래 새 폴더에 저장")
    args = parser.parse_args()
    started = datetime.now(timezone.utc)
    if args.bug_demo:
        import bug_demo
        suite = unittest.defaultTestLoader.loadTestsFromModule(bug_demo)
    else:
        pattern = "test_*.py" if args.group == "all" else f"test_{args.group}.py"
        suite = unittest.defaultTestLoader.discover(str(ROOT / "tests"), pattern=pattern)
    log = io.StringIO()
    with contextlib.redirect_stdout(log), contextlib.redirect_stderr(log):
        result = unittest.TextTestRunner(stream=log, verbosity=2, resultclass=RecordedResult).run(suite)
    print(log.getvalue(), end="")
    # unittest는 skip을 실패로 보지 않는다. 이 실습의 완료 판정은 더 엄격하게 구분.
    complete = result.testsRun > 0 and not result.skipped and not result.expectedFailures
    ok = result.wasSuccessful() and complete
    status = "PASS" if ok else ("INCOMPLETE" if result.wasSuccessful() else "FAIL")
    if args.report:
        folder = KIT / ".runs" / (started.strftime("%Y%m%dT%H%M%S_%fZ") + ("_bug" if args.bug_demo else "_" + args.group))
        folder.mkdir(parents=True, exist_ok=False)
        (folder / "test_log.txt").write_text(log.getvalue(), encoding="utf-8-sig")
        report = {"started_utc": started.isoformat(), "finished_utc": datetime.now(timezone.utc).isoformat(),
                  "python": sys.executable, "python_version": sys.version, "group": "intentional_bug" if args.bug_demo else args.group,
                  "status": status, "tests_run": result.testsRun,
                  "failures": len(result.failures), "errors": len(result.errors), "skipped": len(result.skipped),
                  "browser_tested": False, "tests": result.records}
        (folder / "summary.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"Evidence: {folder}")
    if args.bug_demo:
        print("이 명령은 의도적인 실패 예제입니다. 정상 재현이면 FAIL 1개와 종료코드 1이 나옵니다.")
    if status == "INCOMPLETE":
        print("INCOMPLETE: 실행하지 못한 검사 또는 검사 0개. 통과로 판정하지 않습니다.")
    return 0 if ok else (2 if status == "INCOMPLETE" else 1)


if __name__ == "__main__":
    raise SystemExit(main())
