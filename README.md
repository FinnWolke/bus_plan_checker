# Bus Plan Checker (simple version) – Project 5, Team 7

| File | What it does |
|---|---|
| `logic.py` | All calculations: read files + data checks, battery/SOC per bus, feasibility checks, KPIs, improved plan, Excel export |
| `app.py` | The Streamlit screen: Overview, Feasibility, Gantt chart, Improved plan, Compare plans |
| `run_tests.py` | Runs all test sets in `test_data/` and prints PASS/FAIL |
| `data/` | Example files from Canvas (used by run_tests.py; upload them in the app) |

Run:
```
pip install -r requirements.txt
python -m streamlit run app.py
python run_tests.py
```
