# Bus Plan Checker - the Streamlit screen. Start with:  python -m streamlit run app.py
import pandas as pd
import plotly.express as px
import streamlit as st
from logic import (DEFAULT_SETTINGS, check_plan, improve, is_feasible, kpis, minutes_to_time,
                   read_distances, read_plan, read_timetable, to_excel)

# the feasibility checks with a short explanation
CHECKS = {"FC1": "SOC never below 10%",
          "FC2": "Charging at least 15 minutes",
          "FC3": "Never charged above 90%",
          "FC4": "No overlap (rows in time order), service trips in timetable, no line on material trips",
          "FC5": "Bus starts where it ended",
          "FC6": "Charging only at the garage",
          "FC7": "Trip not faster than the distance matrix allows",
          "FC8": "Every timetable trip driven exactly once",
          "FC9": "Enough idle time between two trips (minimum from the sidebar)"}
# the data checks with a short explanation
DATA_CHECKS = {"DQ-ACT": "Unknown activity",
               "DQ-LOC": "Unknown location",
               "DQ-TIME": "Invalid time",
               "DQ-DUR": "End time before start time",
               "DQ-BUS": "Invalid bus number",
               "DQ-NUM": "Energy is not a number",
               "DQ-LINE": "Line is not 400 or 401",
               "DQ-ROUTE": "Route is not in the distance matrix",
               "DQ-ENERGY": "Energy in the file differs from our calculation (km × kWh/km, standing still 5 kW)"}
COLORS = {"service trip": "#2a6fdb", "material trip": "#f08c00", "idle": "#c9ced6", "charging": "#2f9e44"}
DAY = pd.Timestamp("2026-01-01")

# Gantt chart: one row per bus, coloured by activity, red X where an error is
def gantt_chart(timeline, issues, title=""):
    df = timeline.copy()
    df["Start"] = DAY + pd.to_timedelta(df["start"], unit="m")
    df["End"] = DAY + pd.to_timedelta(df["end"], unit="m")
    df["Bus"] = "Bus " + df["bus"].astype(str)
    fig = px.timeline(df, x_start="Start", x_end="End", y="Bus", color="activity", color_discrete_map=COLORS, title=title,
                      hover_data=["start_loc", "end_loc", "line", "soc"], height=120 + 30 * df["bus"].nunique())
    # red X on every row with an error
    marks = []
    for i in issues:
        if i["severity"] == "error" and i["bus"] is not None and i["row"] is not None:
            marks.append({"bus": i["bus"], "row": i["row"], "message": i["message"]})
    if len(marks) > 0:
        errors = df.merge(pd.DataFrame(marks), on=["bus", "row"])
        fig.add_scatter(x=errors["Start"], y=errors["Bus"], mode="markers", name="error", hovertext=errors["message"],
                        marker=dict(symbol="x", size=11, color="red"))
    # bus 1 at the top
    bus_order = []
    for b in sorted(df["bus"].unique(), reverse=True):
        bus_order.append(f"Bus {b}")
    fig.update_yaxes(categoryorder="array", categoryarray=bus_order)
    fig.update_xaxes(tickformat="%H:%M")
    return fig

# KPIs as big numbers, 5 per row
def show_kpis(values):
    columns = st.columns(5) + st.columns(5)
    number = 0
    for name in values:
        columns[number].metric(name, values[name])
        number = number + 1

# Page: title, 3 upload boxes and 2 sliders
st.set_page_config(page_title="Bus Plan Checker", layout="wide")
st.title("Bus Plan Checker – lines 400 & 401")
st.write("Upload the three Excel files. The checks start as soon as all three are uploaded.")
col1, col2, col3 = st.columns(3)
plan_file = col1.file_uploader("1. Bus plan", type="xlsx")
timetable_file = col2.file_uploader("2. Timetable", type="xlsx")
distance_file = col3.file_uploader("3. Distance matrix", type="xlsx")

st.sidebar.header("Assumptions")
settings = DEFAULT_SETTINGS.copy()
settings["soh"] = st.sidebar.slider("State of Health (%)", 85, 95, 85) / 100
settings["kwh_per_km"] = st.sidebar.slider("Consumption while driving (kWh/km)", 0.7, 2.5, 1.2)
settings["min_idle"] = st.sidebar.slider("Minimum idle time between trips (min)", 0, 15, 0,
                                         help="Buffer in case a bus is late. Used in the checks (FC9) and in the improved plan.")
st.sidebar.caption(f"Usable battery: 300 kWh × {settings['soh']:.0%} = {300 * settings['soh']:.0f} kWh")
st.sidebar.caption("Every bus starts the day with a full battery (100%).")

if not plan_file or not timetable_file or not distance_file:
    st.info("Waiting for all three files …")
    st.stop()

# Read the files and check the plan
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

# only the real errors (no warnings)
errors = []
for i in issues:
    if i["severity"] == "error":
        errors.append(i)
if len(errors) > 0:
    failed = []
    for i in errors:
        if i["check"] not in failed:
            failed.append(i["check"])
    st.error(f"❌ **Not feasible** – {len(errors)} error(s) in: {', '.join(sorted(failed))}")
else:
    st.success("✅ **The bus plan is feasible.**")

tab1, tab2, tab3, tab4, tab5 = st.tabs(["Overview", "Feasibility", "Gantt chart", "Improved plan", "Compare plans"])

# Tab 1: KPIs and a table per bus
with tab1:
    show_kpis(kpis(timeline))
    bus_table = []
    for bus, rows in timeline.groupby("bus"):
        bus_errors = []
        for i in errors:
            if i["bus"] == bus and i["check"] not in bus_errors:
                bus_errors.append(i["check"])
        bus_table.append({"Bus": bus, "Start": minutes_to_time(rows["start"].min()), "End": minutes_to_time(rows["end"].max()),
                          "Lowest SOC (%)": rows["soc"].min(), "Errors": ", ".join(sorted(bus_errors)) or "-"})
    st.dataframe(pd.DataFrame(bus_table), hide_index=True, width="stretch")

# Tab 2: is the data correct, is the plan feasible, where are the errors
with tab2:
    st.subheader("1. Is the data correct?")
    data_problems = 0
    for check in DATA_CHECKS:   # one coloured line per problem type (red = row is skipped), details in a drop-down
        found = []
        for i in issues:
            if i["check"] == check:
                found.append(i)
        if len(found) == 0:
            continue
        data_problems = data_problems + len(found)
        text = f"**{check} ({len(found)}×):** {DATA_CHECKS[check]}"
        if found[0]["severity"] == "error":
            st.error(text)
        else:
            st.warning(text)
        with st.expander(f"Show where ({check})"):
            details = pd.DataFrame(found)
            details = details[["bus", "row", "time", "message"]]
            details = details.fillna("-")
            details.columns = ["bus", "excel row", "time", "what is wrong"]
            st.dataframe(details, hide_index=True, width="stretch")
    if data_problems == 0:
        st.success("No data problems found.")

    st.subheader("2. Is the bus plan feasible?")
    check_table = []
    for check in CHECKS:
        count = 0
        for i in errors:
            if i["check"] == check:
                count = count + 1
        result = "❌ FAIL" if count > 0 else "✅ PASS"
        check_table.append({"Check": check, "Rule": CHECKS[check], "Result": result, "Errors": count})
    st.dataframe(pd.DataFrame(check_table), hide_index=True, width="stretch")

    st.subheader("3. Where are the errors?")
    plan_errors = []
    for i in errors:
        if i["check"].startswith("FC"):
            plan_errors.append(i)
    if len(plan_errors) > 0:
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
    if other_file:
        plan_b, issues_b = read_plan(other_file, distances)
    else:
        plan_b = st.session_state.get("improved")
        issues_b = []
    if plan_b is None:
        st.info("Upload a readable plan B or generate an improved plan first.")
    else:
        timeline_b, more = check_plan(plan_b, timetable, distances, settings)
        issues_b = issues_b + more
        kpis_a = kpis(timeline)
        kpis_b = kpis(timeline_b)
        feasible_a = "yes" if len(errors) == 0 else "no"
        feasible_b = "yes" if is_feasible(issues_b) else "no"
        compare = [{"KPI": "Feasible", "Plan A": feasible_a, "Plan B": feasible_b, "Difference (B - A)": ""}]
        for k in kpis_a:
            compare.append({"KPI": k, "Plan A": kpis_a[k], "Plan B": kpis_b[k], "Difference (B - A)": round(kpis_b[k] - kpis_a[k], 2)})
        st.dataframe(pd.DataFrame(compare).astype(str), hide_index=True, width="stretch")
        left, right = st.columns(2)
        left.plotly_chart(gantt_chart(timeline, issues, "Plan A"), width="stretch")
        right.plotly_chart(gantt_chart(timeline_b, issues_b, "Plan B"), width="stretch")
