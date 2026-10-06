"""Bus Plan Checker - all calculations (no Streamlit in this file).

1. read_distances / read_timetable / read_plan : read the Excel files + data checks (DQ-...)
2. check_plan                                  : simulate the battery per bus + feasibility checks (FC1-FC8)
3. kpis                                        : key performance indicators
4. improve                                     : make an improved (feasible) bus plan
5. to_excel                                    : save a plan in the client's Excel format
"""
import io
import math
from dataclasses import dataclass

import pandas as pd

GARAGE = "ehvgar"
ACTIVITIES = ["service trip", "material trip", "idle", "charging"]
PLAN_COLUMNS = ["start location", "end location", "start time", "end time",
                "activity", "line", "energy consumption", "bus"]


@dataclass
class Settings:
    """Assumptions from the assignment. Percentages are fractions of 300 kWh x SOH."""
    soh: float = 0.85             # state of health
    kwh_per_km: float = 1.2       # consumption while driving
    start_soc: float = 0.90       # battery level at the start of the day
    min_soc: float = 0.10         # FC1: never below 10%
    max_soc: float = 0.90         # FC3: never charged above 90%
    min_charge: int = 15          # FC2: charge at least 15 minutes
    idle_kw: float = 5            # consumption while standing still
    fast_kw: float = 450          # charging power up to 90%
    slow_kw: float = 60           # charging power above 90%
    use_file_energy: bool = False  # True = use the 'energy consumption' column instead of the model

    @property
    def capacity(self):
        return 300 * self.soh


# ----------------------------------------------------------------------------- small helpers
def to_minutes(value):
    """'06:04' or '06:04:00' -> 364 minutes. Times before 03:00 belong to the end of the day (+24h)."""
    parts = str(value).strip().split(":")
    if len(parts) not in (2, 3) or not all(p.isdigit() for p in parts):
        return None
    hours, minutes = int(parts[0]), int(parts[1])
    if hours > 23 or minutes > 59:
        return None
    total = hours * 60 + minutes
    return total + 24 * 60 if total < 180 else total


def hhmm(minutes):
    return f"{int(minutes) // 60 % 24:02d}:{int(minutes) % 60:02d}"


def number(value):
    """float(value), or None if the cell is empty or not a number."""
    try:
        x = float(value)
        return None if math.isnan(x) else x
    except (TypeError, ValueError):
        return None


def problem(check, message, severity="error", bus=None, row=None, time=""):
    return {"severity": severity, "check": check, "bus": bus, "row": row, "time": time, "message": message}


def route(dist, start, end, line=None):
    """(min minutes, max minutes, km) between two locations; line=None for a material trip."""
    return dist.get((start, end, line)) or next((v for (s, e, _), v in dist.items() if (s, e) == (start, end)), None)


# ----------------------------------------------------------------------------- 1. reading + data checks
def read_distances(file):
    dm = pd.read_excel(file)
    missing = {"start", "end", "min_travel_time", "max_travel_time", "distance_m", "line"} - set(dm.columns)
    if missing:
        raise ValueError(f"Distance matrix: missing column(s) {', '.join(sorted(missing))}")
    return {(r.start, r.end, None if pd.isna(r.line) else int(r.line)):
            (int(r.min_travel_time), int(r.max_travel_time), r.distance_m / 1000) for r in dm.itertuples()}


def read_timetable(file):
    tt = pd.read_excel(file, dtype=object)
    missing = {"start", "departure_time", "end", "line"} - set(tt.columns)
    if missing:
        raise ValueError(f"Timetable: missing column(s) {', '.join(sorted(missing))}")
    tt["dep"] = tt["departure_time"].map(to_minutes)
    bad = tt[tt["dep"].isna()]
    if len(bad):
        raise ValueError(f"Timetable: invalid departure time on row(s) {', '.join(str(i + 2) for i in bad.index)}")
    tt["line"] = tt["line"].astype(int)
    return tt[["start", "end", "dep", "line"]].sort_values("dep").reset_index(drop=True)


def read_plan(file, dist):
    """Returns (plan, issues). Rows with data errors are reported and left out of the plan."""
    df = pd.read_excel(file, dtype=object)
    df.columns = [str(c).strip().lower() for c in df.columns]
    missing = [c for c in PLAN_COLUMNS if c not in df.columns]
    if missing:
        return None, [problem("DQ-COL", f"Missing column(s): {', '.join(missing)}.")]
    if df.empty:
        return None, [problem("DQ-EMPTY", "The bus plan contains no data (only a header).")]

    locations = {key[0] for key in dist} | {key[1] for key in dist}
    lines = {key[2] for key in dist if key[2]}
    issues, rows = [], []
    for i, r in df.iterrows():
        row, errors = i + 2, []
        activity = str(r["activity"]).strip().lower()
        start, end = to_minutes(r["start time"]), to_minutes(r["end time"])
        line, energy, bus = number(r["line"]), number(r["energy consumption"]), number(r["bus"])
        if activity not in ACTIVITIES:
            errors.append(("DQ-ACT", f"Unknown activity '{r['activity']}'."))
        if r["start location"] not in locations or r["end location"] not in locations:
            errors.append(("DQ-LOC", f"Unknown location '{r['start location']}' / '{r['end location']}'."))
        if start is None or end is None:
            errors.append(("DQ-TIME", f"Invalid time '{r['start time']}' / '{r['end time']}'."))
        elif end < start:
            errors.append(("DQ-DUR", "End time is before start time."))
        if bus is None:
            errors.append(("DQ-BUS", f"Invalid bus number '{r['bus']}'."))
        if energy is None:
            errors.append(("DQ-NUM", f"Energy consumption '{r['energy consumption']}' is not a number."))
        if activity == "service trip" and line not in lines:
            errors.append(("DQ-LINE", f"Line '{r['line']}' is not one of {sorted(lines)}."))
        if activity in ("service trip", "material trip") and not errors and \
                route(dist, r["start location"], r["end location"], line if activity == "service trip" else None) is None:
            errors.append(("DQ-ROUTE", f"No distance known from {r['start location']} to {r['end location']}."))
        for check, message in errors:
            issues.append(problem(check, message, bus=None if bus is None else int(bus), row=row))
        if not errors:
            rows.append({"row": row, "bus": int(bus), "start_loc": r["start location"], "end_loc": r["end location"],
                         "start": start, "end": end, "activity": activity,
                         "line": None if line is None else int(line), "energy": energy})
    if not rows:
        return None, issues + [problem("DQ-EMPTY", "No usable rows in the bus plan.")]
    plan = pd.DataFrame(rows)
    for bus, acts in plan.groupby("bus"):
        times = list(zip(acts["start"], acts["end"]))
        if times != sorted(times):
            issues.append(problem("DQ-ORDER", "Rows are not in time order (the tool sorts them).", "warning", bus=bus))
    plan["line"] = pd.Series([None if pd.isna(x) else int(x) for x in plan["line"]], index=plan.index, dtype=object)
    return plan.sort_values(["bus", "start", "end"]).reset_index(drop=True), issues


# ----------------------------------------------------------------------------- 2. battery + feasibility checks
def charge(soc, minutes, s):
    """kWh gained when charging: 450 kW up to 90%, 60 kW above 90%."""
    fast_minutes = min(minutes, max(0, s.max_soc * s.capacity - soc) / s.fast_kw * 60)
    gained = fast_minutes * s.fast_kw / 60 + (minutes - fast_minutes) * s.slow_kw / 60
    return min(gained, s.capacity - soc)


def model_energy(a, soc, dist, s):
    """Energy (kWh) of one activity according to our model; negative = charged."""
    minutes = a.end - a.start
    if a.activity == "idle":
        return s.idle_kw * minutes / 60
    if a.activity == "charging":
        return -charge(soc, minutes, s)
    return route(dist, a.start_loc, a.end_loc, a.line if a.activity == "service trip" else None)[2] * s.kwh_per_km


def check_plan(plan, timetable, dist, s):
    """Simulate the battery of every bus. Returns (timeline with SOC per activity, list of issues)."""
    issues, timeline = [], []
    timetable_trips = {(t.start, t.dep, t.end, t.line) for t in timetable.itertuples()}
    driven = {}                                         # timetable trip -> [(bus, row), ...]

    for bus, acts in plan.groupby("bus"):
        soc, prev, too_low = s.start_soc * s.capacity, None, None
        for a in acts.itertuples():
            where = {"bus": bus, "row": a.row, "time": hhmm(a.start)}
            if prev is not None:
                if a.start < prev.end:
                    issues.append(problem("FC4", f"Overlaps with the previous activity (ends {hhmm(prev.end)}).", **where))
                if a.start_loc != prev.end_loc:
                    issues.append(problem("FC5", f"Bus is at {prev.end_loc} but this activity starts at {a.start_loc}.", **where))
                soc -= max(0, a.start - prev.end) * s.idle_kw / 60     # gap = standing still

            energy = model_energy(a, soc, dist, s)
            if abs(a.energy - energy) > 0.05:
                issues.append(problem("DQ-ENERGY", f"Energy in file {a.energy:.2f} kWh, model {energy:.2f} kWh.", "warning", **where))
            if s.use_file_energy:
                energy = a.energy
            soc -= energy
            pct = round(soc / s.capacity * 100, 2)
            minutes = a.end - a.start

            if a.activity == "charging":
                if minutes < s.min_charge:
                    issues.append(problem("FC2", f"Charging session of {minutes} min (minimum {s.min_charge}).", **where))
                if pct > s.max_soc * 100:
                    issues.append(problem("FC3", f"Battery charged to {pct:.2f}% (maximum {s.max_soc * 100:.0f}%).", **where))
                if a.start_loc != GARAGE:
                    issues.append(problem("FC6", f"Charging at {a.start_loc}; only possible at {GARAGE}.", **where))
            if a.activity == "service trip":
                key = (a.start_loc, a.start, a.end_loc, a.line)
                if key in timetable_trips:
                    driven.setdefault(key, []).append((bus, a.row))
                else:
                    issues.append(problem("FC4", f"Service trip line {a.line} at {hhmm(a.start)} is not in the timetable.", **where))
            if a.activity == "material trip" and a.line is not None:
                issues.append(problem("FC4", f"Material trip has line number {a.line}.", **where))
            if a.activity in ("service trip", "material trip"):
                fastest = route(dist, a.start_loc, a.end_loc, a.line if a.activity == "service trip" else None)[0]
                if minutes < fastest:
                    issues.append(problem("FC7", f"Trip takes {minutes} min, minimum travel time is {fastest} min.", **where))
            if pct < s.min_soc * 100 and too_low is None:
                too_low = (where, pct)

            km = route(dist, a.start_loc, a.end_loc, a.line if a.activity == "service trip" else None)[2] \
                if a.activity in ("service trip", "material trip") else 0
            timeline.append({"bus": bus, "row": a.row, "activity": a.activity, "start_loc": a.start_loc, "end_loc": a.end_loc,
                             "start": a.start, "end": a.end, "duration": minutes, "line": a.line, "km": km,
                             "energy": energy, "soc": pct})
            prev = a
        if too_low:
            where, pct = too_low
            issues.append(problem("FC1", f"SOC drops to {pct:.2f}% (minimum {s.min_soc * 100:.0f}%).", **where))

    for t in timetable.itertuples():
        buses = driven.get((t.start, t.dep, t.end, t.line), [])
        if len(buses) != 1:
            text = "is not driven by any bus" if not buses else f"is driven {len(buses)} times"
            bus, row = buses[1] if len(buses) > 1 else (None, None)
            issues.append(problem("FC8", f"Timetable trip {t.start}->{t.end} line {t.line} at {hhmm(t.dep)} {text}.",
                                  bus=bus, row=row, time=hhmm(t.dep)))
    return pd.DataFrame(timeline), issues


def is_feasible(issues):
    return not any(i["severity"] == "error" for i in issues)


# ----------------------------------------------------------------------------- 3. KPIs
def kpis(timeline):
    t = timeline
    material, charging = t[t.activity == "material trip"], t[t.activity == "charging"]
    return {
        "Buses used": int(t.bus.nunique()),
        "Service trips": int((t.activity == "service trip").sum()),
        "Material trips": len(material),
        "Material trip time (min)": int(material.duration.sum()),
        "Material trip distance (km)": round(float(material.km.sum()), 2),
        "Charging sessions": len(charging),
        "Total charging time (min)": int(charging.duration.sum()),
        "Idle time (min)": int(t[t.activity == "idle"].duration.sum()),
        "Energy used (kWh)": round(float(t.energy[t.energy > 0].sum()), 1),
        "Lowest SOC (%)": float(t.soc.min()),
    }


# ----------------------------------------------------------------------------- 4. improved plan
def act(start_loc, end_loc, start, end, activity, energy, line=None):
    return {"start_loc": start_loc, "end_loc": end_loc, "start": start, "end": end,
            "activity": activity, "line": line, "energy": energy}


def deadhead(start, end, dist, s):
    """(minutes, kWh) of a material trip; nothing if the bus is already there."""
    if start == end:
        return 0, 0.0
    _, minutes, km = route(dist, start, end)
    return minutes, km * s.kwh_per_km


def drive_direct(bus, trip, dist, s):
    """Option 1: (wait) + (material trip) + service trip."""
    minutes, energy = deadhead(bus["loc"], trip["start"], dist, s)
    wait = trip["dep"] - minutes - bus["free"]
    if wait < 0:
        return None                                        # bus cannot be there in time
    acts = []
    if wait:
        acts.append(act(bus["loc"], bus["loc"], bus["free"], bus["free"] + wait, "idle", s.idle_kw * wait / 60))
    if minutes:
        acts.append(act(bus["loc"], trip["start"], trip["dep"] - minutes, trip["dep"], "material trip", energy))
    return finish(bus, trip, acts, cost=wait + 2 * minutes, dist=dist, s=s)


def via_garage(bus, trip, dist, s):
    """Option 2: material trip to the garage, charge (>= 15 min, max 90%), material trip to the trip."""
    to_garage, e1 = deadhead(bus["loc"], GARAGE, dist, s)
    from_garage, e2 = deadhead(GARAGE, trip["start"], dist, s)
    arrive, soc = bus["free"] + to_garage, bus["soc"] - e1
    window = trip["dep"] - from_garage - arrive
    minutes = min(window, math.floor((s.max_soc * s.capacity - soc) / s.fast_kw * 60))
    if minutes < s.min_charge:
        return None
    acts = []
    if to_garage:
        acts.append(act(bus["loc"], GARAGE, bus["free"], arrive, "material trip", e1))
    acts.append(act(GARAGE, GARAGE, arrive, arrive + minutes, "charging", -minutes * s.fast_kw / 60))
    if window > minutes:
        acts.append(act(GARAGE, GARAGE, arrive + minutes, arrive + window, "idle", s.idle_kw * (window - minutes) / 60))
    if from_garage:
        acts.append(act(GARAGE, trip["start"], trip["dep"] - from_garage, trip["dep"], "material trip", e2))
    return finish(bus, trip, acts, cost=(window - minutes) + 2 * (to_garage + from_garage) + 10, dist=dist, s=s)


def finish(bus, trip, acts, cost, dist, s):
    """Add the service trip; reject the option if the bus could not reach the garage with >= 10% afterwards."""
    acts.append(act(trip["start"], trip["end"], trip["dep"], trip["arr"], "service trip", trip["energy"], trip["line"]))
    soc = bus["soc"] - sum(a["energy"] for a in acts)
    if soc - deadhead(trip["end"], GARAGE, dist, s)[1] < s.min_soc * s.capacity:
        return None
    return {"acts": acts, "soc": soc, "cost": cost}


def improve(timetable, dist, s):
    """Make an improved plan: give every trip (in order of departure) to the bus that can do it
    with the lowest cost. A bus only goes charging when it cannot do the next trip otherwise.
    A new bus is only used when no existing bus can do the trip."""
    buses = []
    for t in timetable.itertuples():
        _, minutes, km = route(dist, t.start, t.end, t.line)
        trip = {"start": t.start, "end": t.end, "dep": t.dep, "arr": t.dep + minutes, "line": t.line,
                "energy": km * s.kwh_per_km}
        options = []
        for bus in buses:
            option = drive_direct(bus, trip, dist, s) or via_garage(bus, trip, dist, s)
            if option:
                options.append((option["cost"], bus["id"], bus, option))
        if options:
            _, _, bus, option = min(options, key=lambda o: (o[0], o[1]))
        else:                                              # no bus available: a new bus leaves the garage
            bus = {"id": len(buses) + 1, "loc": GARAGE, "soc": s.start_soc * s.capacity, "acts": [],
                   "free": trip["dep"] - deadhead(GARAGE, trip["start"], dist, s)[0]}
            option = drive_direct(bus, trip, dist, s)
            buses.append(bus)
        bus["acts"] += option["acts"]
        bus.update(loc=trip["end"], free=trip["arr"], soc=option["soc"])

    rows = []
    for bus in buses:                                      # end of the day: back to the garage
        minutes, energy = deadhead(bus["loc"], GARAGE, dist, s)
        if minutes:
            bus["acts"].append(act(bus["loc"], GARAGE, bus["free"], bus["free"] + minutes, "material trip", energy))
        rows += [{**a, "bus": bus["id"]} for a in bus["acts"]]
    plan = pd.DataFrame(rows)
    plan["row"] = range(2, len(plan) + 2)
    plan["line"] = pd.Series([None if pd.isna(x) else int(x) for x in plan["line"]], index=plan.index, dtype=object)
    return plan


# ----------------------------------------------------------------------------- 5. export
def to_excel(plan):
    """Plan -> Excel bytes with exactly the columns of the client's bus plan."""
    out = pd.DataFrame({
        "start location": plan.start_loc, "end location": plan.end_loc,
        "start time": plan.start.map(lambda m: hhmm(m) + ":00"), "end time": plan.end.map(lambda m: hhmm(m) + ":00"),
        "activity": plan.activity, "line": plan.line, "energy consumption": plan.energy.astype(float).round(4),
        "bus": plan.bus})
    buffer = io.BytesIO()
    out.to_excel(buffer, index=False, sheet_name="Sheet1")
    return buffer.getvalue()
