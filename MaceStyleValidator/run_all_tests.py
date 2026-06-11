"""Run the offline rule test suite. Exit non-zero if anything fails.

  python3 run_all_tests.py

Covers:
  1. Engine coverage   (test_rule_coverage.py)  — each implemented check detects
     and, where applicable, fixes a real violation.
  2. Tracker coverage  (test_tracker_rules.py)  — every rule in the testers' set
     is handled deterministically or by AI, or is a documented gap.

Neither needs SharePoint, Azure, or network. To also lint live rule data:
  python3 dump_style_rules.py && python3 rule_doctor.py rules_snapshot.json
"""
import sys

import test_rule_coverage
import test_tracker_rules


def main():
    results = []
    print("=" * 60)
    print("1/2  Engine coverage")
    print("=" * 60)
    results.append(("engine coverage", test_rule_coverage.run()))

    print("\n" + "=" * 60)
    print("2/2  Tracker coverage")
    print("=" * 60)
    results.append(("tracker coverage", test_tracker_rules.run()))

    print("\n" + "=" * 60)
    failed = [name for name, code in results if code != 0]
    for name, code in results:
        print(f"  {'PASS' if code == 0 else 'FAIL'}  {name}")
    print("=" * 60)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
