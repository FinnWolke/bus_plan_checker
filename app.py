"""Bus Plan Checker - the screen. Start with:  streamlit run app.py"""
import pandas as pd
import plotly.express as px
import streamlit as st

from logic import (Settings, check_plan, improve, is_feasible, kpis, read_distances, read_plan,
                   read_timetable, to_excel)

CHECKS = {
    "FC1": "SOC never below 10%", "FC2": "Charging at least 15 minutes", "FC3": "Never charged above 90%",
    "FC4": "No overlap, service trips in timetable, no line on material trips", "FC5": "Bus starts where it ended",
    "FC6": "Charging only at the garage", "FC7": "Trip not faster than the distance matrix allows",
    "FC8": "Every timetable trip driven exactly once",
}
COLORS = {"service trip": "#2a6fdb", "material trip": "#f08c00", "idle": "#c9ced6", "charging": "#2f9e44"}

st.set_page_config(page_title="Bus Plan Checker", layout="wide")
st.title("Bus Plan Checker – lines 400 & 401")

# ----------------------------------------------------------------------------- step 1: upload the 3 Excel files
st.write("Upload the three Excel files. The checks start as soon as all three are uploaded.")
col1, col2, col3 = st.columns(3)
plan_file = col1.file_uploader("1. Bus plan", type="xlsx")
timetable_file = col2.file_uploader("2. Timetable", type="xlsx")
distance_file = col3.file_uploader("3. Distance matrix", type="xlsx")

with st.sidebar:
    st.header("Assumptions")
    s = Settings(soh=st.slider("State of Health (%)", 85, 95, 85) / 100,
                 kwh_per_km=st.slider("Consumption while driving (kWh/km)", 0.7, 2.5, 1.2))
    st.caption(f"Usable battery: 300 kWh × {s.soh:.0%} = {s.capacity:.0f} kWh")

if not (plan_file and timetable_file and distance_file):
    st.info("Waiting for all three files …")
    st.stop()

# ----------------------------------------------------------------------------- read + check
try:
    dist = read_distances(distance_file)
    timetable = read_timetable(timetable_file)
except Exception as error:
    st.error(f"❌ {error}")
    st.stop()

plan, issues = read_plan(plan_file, dist)
if plan is None:
    st.error("❌ The bus plan cannot be read. Fix these data errors and upload again.")
    st.dataframe(pd.DataFrame(issues), hide_index=True)
    st.stop()
timeline, feasibility_issues = check_plan(plan, timetable, dist, s)
issues += feasibility_issues

errors = [i for i in issues if i["severity"] == "error"]
if errors:
    st.error(f"❌ **Not feasible** – {len(errors)} error(s) in: {', '.join(sorted({i['check'] for i in errors}))}")
else:
    st.success("✅ **The bus plan is feasible.**")


def gantt(timeline, issues, title=""):
    """Gantt chart: one row per bus, coloured by activity, red ✕ where an error starts."""
    day = pd.Timestamp("2026-01-01")
    df = timeline.assign(Start=day + pd.to_timedelta(timeline.start, unit="m"),
                         End=day + pd.to_timedelta(timeline.end, unit="m"), Bus="Bus " + timeline.bus.astype(str))
    fig = px.timeline(df, x_start="Start", x_end="End", y="Bus", color="activity", color_discrete_map=COLORS,
                      hover_data={"start_loc": True, "end_loc": True, "line": True, "soc": True, "Bus": False},
                      title=title, height=120 + 30 * timeline.bus.nunique())
    marks = [i for i in issues if i["severity"] == "error" and i["bus"] is not None and i["time"]]
    if marks:
        times = [day + pd.Timedelta(minutes=int(i["time"][:2]) * 60 + int(i["time"][3:]) + (1440 if i["time"] < "03" else 0))
                 for i in marks]
        fig.add_scatter(x=times, y=[f"Bus {i['bus']}" for i in marks], mode="markers", name="error",
                        marker=dict(symbol="x", size=11, color="red"), hovertext=[i["message"] for i in marks])
    fig.update_yaxes(categoryorder="array", categoryarray=[f"Bus {b}" for b in sorted(timeline.bus.unique(), reverse=True)])
    fig.update_xaxes(tickformat="%H:%M")
    return fig


def show_kpis(k):
    for chunk in range(0, len(k), 5):
        for col, (name, value) in zip(st.columns(5), list(k.items())[chunk:chunk + 5]):
            col.metric(name, value)


tab_overview, tab_feasibility, tab_gantt, tab_improve, tab_compare = st.tabs(
    ["Overview", "Feasibility", "Gantt chart", "Improved plan", "Compare plans"])

# ----------------------------------------------------------------------------- tabs
with tab_overview:
    show_kpis(kpis(timeline))
    per_bus = timeline.groupby("bus").agg(start=("start", "min"), end=("end", "max"), lowest_soc=("soc", "min"))
    per_bus["start"] = per_bus.start.map(lambda m: f"{m // 60 % 24:02d}:{m % 60:02d}")
    per_bus["end"] = per_bus.end.map(lambda m: f"{m // 60 % 24:02d}:{m % 60:02d}")
    per_bus["errors"] = [", ".join(sorted({i["check"] for i in errors if i["bus"] == b})) or "–" for b in per_bus.index]
    st.dataframe(per_bus, width="stretch")

with tab_feasibility:
    st.subheader("1. Is the data correct?")
    data_issues = [i for i in issues if i["check"].startswith("DQ")]
    if data_issues:
        st.dataframe(pd.DataFrame(data_issues), hide_index=True, width="stretch")
    else:
        st.success("No data problems found.")
    st.subheader("2. Is the bus plan feasible?")
    st.dataframe(pd.DataFrame([{"check": c, "rule": rule,
                                "result": "❌ FAIL" if any(i["check"] == c for i in errors) else "✅ PASS",
                                "errors": sum(i["check"] == c for i in errors)} for c, rule in CHECKS.items()]),
                 hide_index=True, width="stretch")
    st.subheader("3. Where are the errors?")
    plan_errors = [i for i in errors if i["check"].startswith("FC")]
    if plan_errors:
        st.dataframe(pd.DataFrame(plan_errors), hide_index=True, width="stretch")
    else:
        st.success("No feasibility errors.")

with tab_gantt:
    st.plotly_chart(gantt(timeline, issues), width="stretch")

with tab_improve:
    if st.button("Generate improved plan", type="primary"):
        st.session_state["improved"] = improve(timetable, dist, s)
    if "improved" in st.session_state:
        better = st.session_state["improved"]
        better_timeline, better_issues = check_plan(better, timetable, dist, s)
        if is_feasible(better_issues):
            st.success(f"✅ Feasible plan with {better_timeline.bus.nunique()} buses.")
        else:
            st.error("The generated plan is not feasible with these assumptions.")
        show_kpis(kpis(better_timeline))
        st.plotly_chart(gantt(better_timeline, better_issues), width="stretch")
        st.download_button("Download improved plan (Excel)", to_excel(better), "Bus_Planning_improved.xlsx")

with tab_compare:
    other_file = st.file_uploader("Plan B (leave empty to use the improved plan)", type="xlsx")
    if other_file:
        other_plan, other_issues = read_plan(other_file, dist)
        if other_plan is None:
            st.error("Plan B cannot be read.")
            st.stop()
    elif "improved" in st.session_state:
        other_plan, other_issues = st.session_state["improved"], []
    else:
        st.info("Upload plan B or generate an improved plan first.")
        st.stop()
    other_timeline, more = check_plan(other_plan, timetable, dist, s)
    other_issues += more
    a, b = kpis(timeline), kpis(other_timeline)
    rows = [{"KPI": "Feasible", "Plan A": "no" if errors else "yes",
             "Plan B": "yes" if is_feasible(other_issues) else "no", "Difference (B − A)": ""}]
    rows += [{"KPI": k, "Plan A": a[k], "Plan B": b[k], "Difference (B − A)": round(b[k] - a[k], 2)} for k in a]
    st.dataframe(pd.DataFrame(rows).astype(str), hide_index=True, width="stretch")
    left, right = st.columns(2)
    left.plotly_chart(gantt(timeline, issues, "Plan A"), width="stretch")
    right.plotly_chart(gantt(other_timeline, other_issues, "Plan B"), width="stretch")
