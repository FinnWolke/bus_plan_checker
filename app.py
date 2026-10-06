# Bus Plan Checker - the Streamlit screen. Start with:  python -m streamlit run app.py
import pandas as pd
import plotly.express as px
import streamlit as st
from logic import (DEFAULT_SETTINGS, check_plan, improve, is_feasible, kpis, minutes_to_time,
                   read_distances, read_plan, read_timetable, to_excel)

CHECKS = {"FC1": "SOC never below 10%", "FC2": "Charging at least 15 minutes", "FC3": "Never charged above 90%",
          "FC4": "No overlap, service trips in timetable, no line on material trips", "FC5": "Bus starts where it ended",
          "FC6": "Charging only at the garage", "FC7": "Trip not faster than the distance matrix allows",
          "FC8": "Every timetable trip driven exactly once"}
COLORS = {"service trip": "#2a6fdb", "material trip": "#f08c00", "idle": "#c9ced6", "charging": "#2f9e44"}
DAY = pd.Timestamp("2026-01-01")

# Gantt chart: one row per bus, coloured by activity, red X where an error is
def gantt_chart(timeline, issues, title=""):
    df = timeline.assign(Start=DAY + pd.to_timedelta(timeline["start"], unit="m"),
                         End=DAY + pd.to_timedelta(timeline["end"], unit="m"), Bus="Bus " + timeline["bus"].astype(str))
    fig = px.timeline(df, x_start="Start", x_end="End", y="Bus", color="activity", color_discrete_map=COLORS, title=title,
                      hover_data=["start_loc", "end_loc", "line", "soc"], height=120 + 30 * df["bus"].nunique())
    marks = [i for i in issues if i["severity"] == "error" and i["bus"] is not None and i["row"] is not None]
    if marks:
        errors = df.merge(pd.DataFrame(marks)[["bus", "row", "message"]], on=["bus", "row"])
        fig.add_scatter(x=errors["Start"], y=errors["Bus"], mode="markers", name="error", hovertext=errors["message"],
                        marker=dict(symbol="x", size=11, color="red"))
    fig.update_yaxes(categoryorder="array", categoryarray=[f"Bus {b}" for b in sorted(df["bus"].unique(), reverse=True)])
    fig.update_xaxes(tickformat="%H:%M")
    return fig

# KPIs as big numbers, 5 per row
def show_kpis(values):
    for name, column in zip(values, st.columns(5) + st.columns(5)):
        column.metric(name, values[name])

# Page: title, 3 upload boxes and 2 sliders
st.set_page_config(page_title="Bus Plan Checker", layout="wide")
st.title("Bus Plan Checker – lines 400 & 401")
st.write("Upload the three Excel files. The checks start as soon as all three are uploaded.")
col1, col2, col3 = st.columns(3)
plan_file = col1.file_uploader("1. Bus plan", type="xlsx")
timetable_file = col2.file_uploader("2. Timetable", type="xlsx")
distance_file = col3.file_uploader("3. Distance matrix", type="xlsx")
st.sidebar.header("Assumptions")
settings = dict(DEFAULT_SETTINGS, soh=st.sidebar.slider("State of Health (%)", 85, 95, 85) / 100,
                kwh_per_km=st.sidebar.slider("Consumption while driving (kWh/km)", 0.7, 2.5, 1.2))
st.sidebar.caption(f"Usable battery: 300 kWh × {settings['soh']:.0%} = {300 * settings['soh']:.0f} kWh")
if not (plan_file and timetable_file and distance_file):
    st.info("Waiting for all three files …")
    st.stop()

# Read the files and check the plan
try:
    distances, timetable = read_distances(distance_file), read_timetable(timetable_file)
except Exception as error:
    st.error(f"❌ {error}")
    st.stop()
plan, issues = read_plan(plan_file, distances)
if plan is None:
    st.error("❌ The bus plan cannot be read. Fix these data errors and upload again.")
    st.dataframe(pd.DataFrame(issues), hide_index=True)
    st.stop()
timeline, plan_issues = check_plan(plan, timetable, distances, settings)
issues += plan_issues
errors = [i for i in issues if i["severity"] == "error"]
if errors:
    st.error(f"❌ **Not feasible** – {len(errors)} error(s) in: {', '.join(sorted({i['check'] for i in errors}))}")
else:
    st.success("✅ **The bus plan is feasible.**")
tab1, tab2, tab3, tab4, tab5 = st.tabs(["Overview", "Feasibility", "Gantt chart", "Improved plan", "Compare plans"])

# Tab 1: KPIs and a table per bus
with tab1:
    show_kpis(kpis(timeline))
    st.dataframe(pd.DataFrame([{"Bus": bus, "Start": minutes_to_time(rows["start"].min()), "End": minutes_to_time(rows["end"].max()),
                                "Lowest SOC (%)": rows["soc"].min(),
                                "Errors": ", ".join(sorted({i["check"] for i in errors if i["bus"] == bus})) or "-"}
                               for bus, rows in timeline.groupby("bus")]), hide_index=True, width="stretch")

# Tab 2: is the data correct, is the plan feasible, where are the errors
with tab2:
    st.subheader("1. Is the data correct?")
    data_issues = [i for i in issues if i["check"].startswith("DQ")]
    if data_issues:
        st.dataframe(pd.DataFrame(data_issues), hide_index=True, width="stretch")
    else:
        st.success("No data problems found.")
    st.subheader("2. Is the bus plan feasible?")
    counts = {c: sum(i["check"] == c for i in errors) for c in CHECKS}
    st.dataframe(pd.DataFrame([{"Check": c, "Rule": CHECKS[c], "Result": "❌ FAIL" if counts[c] else "✅ PASS", "Errors": counts[c]}
                               for c in CHECKS]), hide_index=True, width="stretch")
    st.subheader("3. Where are the errors?")
    plan_errors = [i for i in errors if i["check"].startswith("FC")]
    if plan_errors:
        st.dataframe(pd.DataFrame(plan_errors), hide_index=True, width="stretch")
    else:
        st.success("No feasibility errors.")

# Tab 3: Gantt chart
with tab3:
    st.plotly_chart(gantt_chart(timeline, issues), width="stretch")

# Tab 4: make, show and download the improved plan
with tab4:
    if st.button("Generate improved plan", type="primary"):
        st.session_state["improved"] = improve(timetable, distances, settings)
    if "improved" in st.session_state:
        new_plan = st.session_state["improved"]
        new_timeline, new_issues = check_plan(new_plan, timetable, distances, settings)
        if is_feasible(new_issues):
            st.success(f"✅ Feasible plan with {new_timeline['bus'].nunique()} buses.")
        else:
            st.error("The generated plan is not feasible with these assumptions.")
        show_kpis(kpis(new_timeline))
        st.plotly_chart(gantt_chart(new_timeline, new_issues), width="stretch")
        st.download_button("Download improved plan (Excel)", to_excel(new_plan), "Bus_Planning_improved.xlsx")

# Tab 5: compare the uploaded plan (A) with another file or the improved plan (B)
with tab5:
    other_file = st.file_uploader("Plan B (leave empty to use the improved plan)", type="xlsx")
    plan_b, issues_b = read_plan(other_file, distances) if other_file else (st.session_state.get("improved"), [])
    if plan_b is None:
        st.info("Upload a readable plan B or generate an improved plan first.")
    else:
        timeline_b, more = check_plan(plan_b, timetable, distances, settings)
        issues_b += more
        a, b = kpis(timeline), kpis(timeline_b)
        table = [{"KPI": "Feasible", "Plan A": "no" if errors else "yes", "Plan B": "yes" if is_feasible(issues_b) else "no",
                  "Difference (B - A)": ""}]
        table += [{"KPI": k, "Plan A": a[k], "Plan B": b[k], "Difference (B - A)": round(b[k] - a[k], 2)} for k in a]
        st.dataframe(pd.DataFrame(table).astype(str), hide_index=True, width="stretch")
        left, right = st.columns(2)
        left.plotly_chart(gantt_chart(timeline, issues, "Plan A"), width="stretch")
        right.plotly_chart(gantt_chart(timeline_b, issues_b, "Plan B"), width="stretch")
