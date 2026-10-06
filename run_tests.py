# Runs all test sets from the test plan (expected results in test_data/expected.json). Start with: python run_tests.py
import io, json, os
import openpyxl
from logic import DEFAULT_SETTINGS, check_plan, improve, kpis, read_distances, read_plan, read_timetable, to_excel

FOLDER = os.path.dirname(os.path.abspath(__file__))

# find a file in test_data/ or data/ ("(original)" = the client's bus plan)
def path(name):
    if "(original)" in name:
        name = "Bus_Planning.xlsx"
    test_file = os.path.join(FOLDER, "test_data", name)
    return test_file if os.path.exists(test_file) else os.path.join(FOLDER, "data", name)

# is there an issue with this check (and row / bus, if given)?
def found(issues, check, severity="error", row=None, bus=None):
    return any(i["check"] == check and i["severity"] == severity and row in (None, i["row"]) and bus in (None, i["bus"])
               for i in issues)

# run one test; returns a list of problems (empty list = PASS)
def run_test(test):
    settings = dict(DEFAULT_SETTINGS, use_file_energy=test["energy_source"] == "file", start_soc=test.get("start_soc") or 0.9)
    distances, timetable = read_distances(path("DistanceMatrix.xlsx")), read_timetable(path(test["timetable"]))
    if test["id"] == "GEN-2":   # downloaded Excel must have the same sheet and columns as the client's file
        ours = openpyxl.load_workbook(io.BytesIO(to_excel(improve(timetable, distances, settings)))).active
        client = openpyxl.load_workbook(path("Bus_Planning.xlsx")).active
        return [] if [c.value for c in ours[1]] == [c.value for c in client[1]] and ours.title == client.title else ["wrong format"]
    plan, issues = (improve(timetable, distances, settings), []) if test["id"] == "GEN-1" else read_plan(path(test["file"]), distances)
    timeline = None
    if plan is not None:
        timeline, more = check_plan(plan, timetable, distances, settings)
        issues += more
    problems = [f"expected {s} {c} (row {r}, bus {b})" for c, s, r, b in test["must"] if not found(issues, c, s, r, b)]
    problems += [f"unexpected error {c}" for c in test["must_not"] if found(issues, c)]
    problems += [f"{k} = {kpis(timeline)[k]}, expected {v}" for k, v in test["kpi"].items() if kpis(timeline)[k] != v]
    if test["id"] == "GEN-1" and kpis(timeline)["Buses used"] >= 20:
        problems.append("not fewer than 20 buses")
    if test.get("energy_check"):
        e = test["energy_check"]
        energy = timeline[(timeline["bus"] == e["bus"]) & (timeline["start"] == e["start"])]["energy"].iloc[0]
        if abs(energy - e["value"]) > 0.000001:
            problems.append(f"energy {energy:.4f}, expected {e['value']}")
    return problems

# run all tests and print the result
if __name__ == "__main__":
    tests = json.load(open(os.path.join(FOLDER, "test_data", "expected.json")))
    passed = 0
    for test in tests:
        problems = run_test(test)
        passed += not problems
        print(f"{'PASS' if not problems else 'FAIL'}  {test['id']:<8} {test['what']}" + (f"  -> {problems}" if problems else ""))
    print(f"\n{passed} of {len(tests)} tests passed")
