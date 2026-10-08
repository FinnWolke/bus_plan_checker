# Runs all test sets from the test plan (expected results in test_data/expected.json). Start with: python run_tests.py
import io
import json
import os
import openpyxl
from logic import DEFAULT_SETTINGS, check_plan, improve, kpis, read_distances, read_plan, read_timetable, to_excel

FOLDER = os.path.dirname(os.path.abspath(__file__))

# find a file in test_data/ or data/ ("(original)" = the client's bus plan)
def path(name):
    if "(original)" in name:
        name = "Bus_Planning.xlsx"
    test_file = os.path.join(FOLDER, "test_data", name)
    if os.path.exists(test_file):
        return test_file
    return os.path.join(FOLDER, "data", name)

# is there an issue with this check (and row / bus, if given)?
def found(issues, check, severity="error", row=None, bus=None):
    for i in issues:
        if i["check"] != check or i["severity"] != severity:
            continue
        if row is not None and i["row"] != row:
            continue
        if bus is not None and i["bus"] != bus:
            continue
        return True
    return False

# run one test; returns a list of problems (empty list = PASS)
def run_test(test):
    settings = DEFAULT_SETTINGS.copy()
    settings["use_file_energy"] = test["energy_source"] == "file"
    if test.get("min_idle"):
        settings["min_idle"] = test["min_idle"]
    if test.get("start_soc"):
        settings["start_soc"] = test["start_soc"]
    distances = read_distances(path("DistanceMatrix.xlsx"))
    timetable = read_timetable(path(test["timetable"]))

    # GEN-2: the downloaded Excel must have the same sheet name and columns as the client's file
    if test["id"] == "GEN-2":
        ours = openpyxl.load_workbook(io.BytesIO(to_excel(improve(timetable, distances, settings)))).active
        client = openpyxl.load_workbook(path("Bus_Planning.xlsx")).active
        our_columns = [c.value for c in ours[1]]
        client_columns = [c.value for c in client[1]]
        if our_columns == client_columns and ours.title == client.title:
            return []
        return ["wrong format"]

    # GEN tests check the generated plan, all other tests check a file
    if test["file"] == "Generated in the tool":
        plan = improve(timetable, distances, settings)
        issues = []
    else:
        plan, issues = read_plan(path(test["file"]), distances)
    timeline = None
    if plan is not None:
        timeline, more = check_plan(plan, timetable, distances, settings)
        issues = issues + more

    problems = []
    for check, severity, row, bus in test["must"]:
        if not found(issues, check, severity, row, bus):
            problems.append(f"expected {severity} {check} (row {row}, bus {bus})")
    for check in test["must_not"]:
        if found(issues, check):
            problems.append(f"unexpected error {check}")
    for name, value in test["kpi"].items():
        if kpis(timeline)[name] != value:
            problems.append(f"{name} = {kpis(timeline)[name]}, expected {value}")
    if test["id"] == "GEN-1" and kpis(timeline)["Buses used"] >= 20:
        problems.append("not fewer than 20 buses")
    if test.get("energy_check"):
        e = test["energy_check"]
        row = timeline[(timeline["bus"] == e["bus"]) & (timeline["start"] == e["start"])]
        energy = row["energy"].iloc[0]
        if abs(energy - e["value"]) > 0.000001:
            problems.append(f"energy {energy:.4f}, expected {e['value']}")
    return problems

# run all tests and print the result
if __name__ == "__main__":
    tests = json.load(open(os.path.join(FOLDER, "test_data", "expected.json")))
    passed = 0
    for test in tests:
        problems = run_test(test)
        if len(problems) == 0:
            passed = passed + 1
            print(f"PASS  {test['id']:<8} {test['what']}")
        else:
            print(f"FAIL  {test['id']:<8} {test['what']}  -> {problems}")
    print(f"\n{passed} of {len(tests)} tests passed")
