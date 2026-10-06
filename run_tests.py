"""
Runs all test sets from the test plan and prints PASS or FAIL.
Start with:  python run_tests.py

The expected result of every test is in test_data/expected.json.
"""
import io
import json
import os

import openpyxl

from logic import (DEFAULT_SETTINGS, check_plan, improve, kpis, read_distances, read_plan,
                   read_timetable, to_excel)

FOLDER = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(FOLDER, "data")
TEST_DATA = os.path.join(FOLDER, "test_data")


def file_path(name):
    """Finds a file in test_data/ or data/."""
    if "(original)" in name:
        return os.path.join(DATA, "Bus_Planning.xlsx")
    if os.path.exists(os.path.join(TEST_DATA, name)):
        return os.path.join(TEST_DATA, name)
    return os.path.join(DATA, name)


def found(issues, check, severity, row, bus):
    """Is there an issue with this check (and row / bus, if given)?"""
    for issue in issues:
        if issue["check"] != check or issue["severity"] != severity:
            continue
        if row is not None and issue["row"] != row:
            continue
        if bus is not None and issue["bus"] != bus:
            continue
        return True
    return False


def run_test(test):
    """Runs one test. Returns a list of problems (empty list = PASS)."""
    settings = dict(DEFAULT_SETTINGS)
    settings["use_file_energy"] = test["energy_source"] == "file"
    if test.get("start_soc"):
        settings["start_soc"] = test["start_soc"]

    distances = read_distances(os.path.join(DATA, "DistanceMatrix.xlsx"))
    timetable = read_timetable(file_path(test["timetable"]))

    # GEN-2: the downloaded Excel file must have the same columns as the client's file
    if test["id"] == "GEN-2":
        ours = openpyxl.load_workbook(io.BytesIO(to_excel(improve(timetable, distances, settings)))).active
        client = openpyxl.load_workbook(os.path.join(DATA, "Bus_Planning.xlsx")).active
        our_header = [cell.value for cell in ours[1]]
        client_header = [cell.value for cell in client[1]]
        if our_header == client_header and ours.title == client.title:
            return []
        return ["export format differs from the client's file"]

    # GEN-1 checks the plan made by the tool, the other tests read a test file
    if test["id"] == "GEN-1":
        plan = improve(timetable, distances, settings)
        issues = []
    else:
        plan, issues = read_plan(file_path(test["file"]), distances)

    timeline = None
    if plan is not None:
        timeline, more_issues = check_plan(plan, timetable, distances, settings)
        issues = issues + more_issues

    problems = []
    for check, severity, row, bus in test["must"]:
        if not found(issues, check, severity, row, bus):
            problems.append(f"expected {severity} {check} (row {row}, bus {bus}) not found")
    for check in test["must_not"]:
        if found(issues, check, "error", None, None):
            problems.append(f"unexpected error {check}")
    if test["id"] == "GEN-1" and kpis(timeline)["Buses used"] >= 20:
        problems.append("generated plan does not use fewer than 20 buses")
    for name, value in test["kpi"].items():
        if kpis(timeline)[name] != value:
            problems.append(f"{name} = {kpis(timeline)[name]}, expected {value}")
    if test.get("energy_check"):
        e = test["energy_check"]
        row = timeline[(timeline["bus"] == e["bus"]) & (timeline["start"] == e["start"])]
        energy = row["energy"].iloc[0]
        if abs(energy - e["value"]) > 0.000001:
            problems.append(f"energy {energy:.4f}, expected {e['value']}")
    return problems


if __name__ == "__main__":
    with open(os.path.join(TEST_DATA, "expected.json")) as f:
        tests = json.load(f)

    passed = 0
    for test in tests:
        problems = run_test(test)
        if len(problems) == 0:
            passed += 1
            print(f"PASS  {test['id']:<8} {test['what']}")
        else:
            print(f"FAIL  {test['id']:<8} {test['what']}  -> {problems}")
    print(f"\n{passed} of {len(tests)} tests passed")
