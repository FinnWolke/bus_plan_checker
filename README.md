# Bus Plan Checker – Project 5, Team 7

Online: https://busplan-team7.streamlit.app

| File | What it does |
|---|---|
| `logic.py` | All calculations: reading the files + data checks, battery per bus, feasibility checks FC1–FC8, KPIs, improved plan, Excel export |
| `app.py` | The Streamlit screen: Overview, Feasibility, Gantt chart, Improved plan, Compare plans |
| `run_tests.py` | Runs all test sets in `test_data/` and prints PASS/FAIL |
| `data/` | Example files from Canvas |

Run on your own computer:
```
pip install -r requirements.txt
python -m streamlit run app.py
python run_tests.py
```
