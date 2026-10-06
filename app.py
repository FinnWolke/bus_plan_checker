"""
Bus Plan Checker - the Streamlit screen.
Start with:  python -m streamlit run app.py
"""
import pandas as pd
import plotly.express as px
import streamlit as st

from logic import (DEFAULT_SETTINGS, check_plan, improve, is_feasible, kpis, minutes_to_time,
                   read_distances, read_plan, read_timetable, to_excel)

# If you start this file with the Run button (python app.py), start Streamlit for you.
if __name__ == "__main__":
    from streamlit.runtime import exists
    if not exists():
        import subprocess
        import sys
        subprocess.run([sys.executable, "-m", "streamlit", "run", __file__])
        sys.exit()

CHECKS = {
    "FC1": "SOC never below 10%",
    "FC2": "Charging at least 15 minutes",
    "FC3": "Never charged above 90%",
    "FC4": "No overlap, service trips in timetable, no line on material trips",
    "FC5": "Bus starts where it ended",
    "FC6": "Charging only at the garage",
    "FC7": "Trip not faster than the distance matrix allows",
    "FC8": "Every timetable trip driven exactly once",
}
COLORS = {"service trip": "#2a6fdb", "material trip": "#f08c00", "idle": "#c9ced6", "charging": "#2f9e44"}


# ---------------------------------------------------------------------------
# Functions for the screen
# ---------------------------------------------------------------------------
def gantt_chart(timeline, issues, title=""):
    """One row per bus, coloured by activity. A red X marks where an error is."""
    day = pd.Timestamp("2026-01-01")
    df = timeline.copy()
    df["Start"] = day + pd.to_timedelta(df["start"], unit="m")
    df["End"] = day + pd.to_timedelta(df["end"], unit="m")
    df["Bus"] = "Bus " + df["bus"].astype(str)

    fig = px.timeline(df, x_start="Start", x_end="End", y="Bus", color="activity",
                      color_discrete_map=COLORS, title=title,
                      hover_data=["start_loc", "end_loc", "line", "soc"],
                      height=120 + 30 * df["bus"].nunique())

    # Red X at the start of every activity with an error
    error_times = []
    error_buses = []
    error_texts = []
    for issue in issues:
        if issue["severity"] == "error" and issue["bus"] is not None and issue["row"] is not None:
            match = timeline[(timeline["bus"] == issue["bus"]) & (timeline["row"] == issue["row"])]
            if len(match) > 0:
                error_times.append(day + pd.Timedelta(minutes=int(match["start"].iloc[0])))
                error_buses.append(f"Bus {issue['bus']}")
                error_texts.append(issue["message"])
    if error_times:
        fig.add_scatter(x=error_times, y=error_buses, mode="markers", name="error",
                        marker=dict(symbol="x", size=11, color="red"), hovertext=error_texts)

    bus_order = [f"Bus {b}" for b in sorted(df["bus"].unique(), reverse=True)]
    fig.update_yaxes(categoryorder="array", categoryarray=bus_order)
    fig.update_xaxes(tickformat="%H:%M")
    return fig


def show_kpis(kpi_values):
    """Shows the KPIs as big numbers, 5 per row."""
    names = list(kpi_values.keys())
    for start in range(0, len(names), 5):
        columns = st.columns(5)
        for column, name in zip(columns, names[start:start + 5]):
            column.metric(name, kpi_values[name])


# ---------------------------------------------------------------------------
# Page: title, upload boxes and sliders
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Bus Plan Checker", layout="wide")
st.title("Bus Plan Checker – lines 400 & 401")
st.write("Upload the three Excel files. The checks start as soon as all three are uploaded.")

col1, col2, col3 = st.columns(3)
plan_file = col1.file_uploader("1. Bus plan", type="xlsx")
timetable_file = col2.file_uploader("2. Timetable", type="xlsx")
distance_file = col3.file_uploader("3. Distance matrix", type="xlsx")

st.sidebar.header("Assumptions")
settings = dict(DEFAULT_SETTINGS)
settings["soh"] = st.sidebar.slider("State of Health (%)", 85, 95, 85) / 100
settings["kwh_per_km"] = st.sidebar.slider("Consumption while driving (kWh/km)", 0.7, 2.5, 1.2)
st.sidebar.caption(f"Usable battery: 300 kWh × {settings['soh']:.0%} = {300 * settings['soh']:.0f} kWh")

if plan_file is None or timetable_file is None or distance_file is None:
    st.info("Waiting for all three files …")
    st.stop()

# ---------------------------------------------------------------------------
# Read the files and check the plan
# ---------------------------------------------------------------------------
try:
    distances = read_distances(distance_file)
    timetable = read_timetable(timetable_file)
except Exception as error:
    st.error(f"❌ {error}")
    st.stop()

plan, issues = read_plan(plan_file, distances)
if plan is None:
    st.error("❌ The bus plan cannot be read. Fix these data errors and upload again.")
    st.dataframe(pd.DataFrame(issues), hide_index=True)
    st.stop()

timeline, plan_issues = check_plan(plan, timetable, distances, settings)
issues = issues + plan_issues
errors = [issue for issue in issues if issue["severity"] == "error"]

if len(errors) == 0:
    st.success("✅ **The bus plan is feasible.**")
else:
    failed_checks = sorted(set(issue["check"] for issue in errors))
    st.error(f"❌ **Not feasible** – {len(errors)} error(s) in: {', '.join(failed_checks)}")

tab1, tab2, tab3, tab4, tab5 = st.tabs(["Overview", "Feasibility", "Gantt chart", "Improved plan", "Compare plans"])

# ---------------------------------------------------------------------------
# Tab 1: Overview
# ---------------------------------------------------------------------------
with tab1:
    show_kpis(kpis(timeline))

    rows = []
    for bus in sorted(timeline["bus"].unique()):
        bus_timeline = timeline[timeline["bus"] == bus]
        bus_errors = sorted(set(issue["check"] for issue in errors if issue["bus"] == bus))
        rows.append({"Bus": bus,
                     "Start": minutes_to_time(bus_timeline["start"].min()),
                     "End": minutes_to_time(bus_timeline["end"].max()),
                     "Lowest SOC (%)": bus_timeline["soc"].min(),
                     "Errors": ", ".join(bus_errors) if bus_errors else "-"})
    st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")

# ---------------------------------------------------------------------------
# Tab 2: Feasibility
# ---------------------------------------------------------------------------
with tab2:
    st.subheader("1. Is the data correct?")
    data_issues = [issue for issue in issues if issue["check"].startswith("DQ")]
    if data_issues:
        st.dataframe(pd.DataFrame(data_issues), hide_index=True, width="stretch")
    else:
        st.success("No data problems found.")

    st.subheader("2. Is the bus plan feasible?")
    summary = []
    for check, rule in CHECKS.items():
        count = len([issue for issue in errors if issue["check"] == check])
        result = "❌ FAIL" if count > 0 else "✅ PASS"
        summary.append({"Check": check, "Rule": rule, "Result": result, "Errors": count})
    st.dataframe(pd.DataFrame(summary), hide_index=True, width="stretch")

    st.subheader("3. Where are the errors?")
    plan_errors = [issue for issue in errors if issue["check"].startswith("FC")]
    if plan_errors:
        st.dataframe(pd.DataFrame(plan_errors), hide_index=True, width="stretch")
    else:
        st.success("No feasibility errors.")

# ---------------------------------------------------------------------------
# Tab 3: Gantt chart
# ---------------------------------------------------------------------------
with tab3:
    st.plotly_chart(gantt_chart(timeline, issues), width="stretch")

# ---------------------------------------------------------------------------
# Tab 4: Improved plan
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# Tab 5: Compare plans
# ---------------------------------------------------------------------------
with tab5:
    other_file = st.file_uploader("Plan B (leave empty to use the improved plan)", type="xlsx")

    plan_b = None
    issues_b = []
    if other_file is not None:
        plan_b, issues_b = read_plan(other_file, distances)
        if plan_b is None:
            st.error("Plan B cannot be read.")
    elif "improved" in st.session_state:
        plan_b = st.session_state["improved"]
    else:
        st.info("Upload plan B or generate an improved plan first.")

    if plan_b is not None:
        timeline_b, more_issues = check_plan(plan_b, timetable, distances, settings)
        issues_b = issues_b + more_issues
        kpi_a = kpis(timeline)
        kpi_b = kpis(timeline_b)

        table = [{"KPI": "Feasible",
                  "Plan A": "yes" if len(errors) == 0 else "no",
                  "Plan B": "yes" if is_feasible(issues_b) else "no",
                  "Difference (B - A)": ""}]
        for name in kpi_a:
            table.append({"KPI": name, "Plan A": kpi_a[name], "Plan B": kpi_b[name],
                          "Difference (B - A)": round(kpi_b[name] - kpi_a[name], 2)})
        st.dataframe(pd.DataFrame(table).astype(str), hide_index=True, width="stretch")

        left, right = st.columns(2)
        left.plotly_chart(gantt_chart(timeline, issues, "Plan A"), width="stretch")
        right.plotly_chart(gantt_chart(timeline_b, issues_b, "Plan B"), width="stretch")
