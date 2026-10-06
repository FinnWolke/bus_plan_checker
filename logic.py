"""Bus Plan Checker - all calculations: reading + data checks, feasibility checks, KPIs,
improved plan and Excel export."""
import io
import math

import pandas as pd

# Fixed values from the assignment
BATTERY_KWH = 300
MIN_SOC = 0.10           # battery never below 10%
MAX_SOC = 0.90           # never charged above 90%
MIN_CHARGE_TIME = 15     # charge at least 15 minutes
IDLE_KW = 5              # energy use when standing still
FAST_KW = 450            # charging speed up to 90%
SLOW_KW = 60             # charging speed above 90%
GARAGE = "ehvgar"

# Values that can be changed in the app (or in the tests)
DEFAULT_SETTINGS = {"soh": 0.85, "kwh_per_km": 1.2, "start_soc": 0.90, "use_file_energy": False}

ACTIVITIES = ["service trip", "material trip", "idle", "charging"]
PLAN_COLUMNS = ["start location", "end location", "start time", "end time",
                "activity", "line", "energy consumption", "bus"]


# ----------------------------------------------------------------- helpers
def time_to_minutes(value):
    """'06:04' or '06:04:00' -> 364. Times before 03:00 count as after midnight (+24h)."""
    parts = str(value).strip().split(":")
    if len(parts) not in (2, 3) or not all(p.isdigit() for p in parts):
        return None
    hours, minutes = int(parts[0]), int(parts[1])
    if hours > 23 or minutes > 59:
        return None
    total = hours * 60 + minutes
    return total + 1440 if total < 180 else total


def minutes_to_time(minutes):
    return f"{int(minutes) // 60 % 24:02d}:{int(minutes) % 60:02d}"


def to_number(value):
    try:
        number = float(value)
        return None if math.isnan(number) else number
    except (TypeError, ValueError):
        return None


def make_issue(check, message, severity="error", bus=None, row=None, time=""):
    return {"severity": severity, "check": check, "bus": bus, "row": row, "time": time, "message": message}


def find_route(distances, start, end, line=None):
    """(min minutes, max minutes, km). Material trips use line=None."""
    return distances.get((start, end, line)) or distances.get((start, end, None))


def capacity(settings):
    return BATTERY_KWH * settings["soh"]


def fix_line_column(plan):
    """pandas turns None into NaN and 401 into 401.0; we want None or a whole number."""
    plan["line"] = pd.Series([None if pd.isna(x) else int(x) for x in plan["line"]], index=plan.index, dtype=object)


# ----------------------------------------------------------------- 1. read files + data checks
def check_columns(df, columns, name):
    missing = [c for c in columns if c not in df.columns]
    if missing:
        raise ValueError(f"{name}: missing column(s) {', '.join(missing)}")


def read_distances(file):
    df = pd.read_excel(file)
    check_columns(df, ["start", "end", "min_travel_time", "max_travel_time", "distance_m", "line"], "Distance matrix")
    distances = {}
    for _, r in df.iterrows():
        line = None if pd.isna(r["line"]) else int(r["line"])
        distances[(r["start"], r["end"], line)] = (int(r["min_travel_time"]), int(r["max_travel_time"]), r["distance_m"] / 1000)
    return distances


def read_timetable(file):
    df = pd.read_excel(file, dtype=object)
    check_columns(df, ["start", "departure_time", "end", "line"], "Timetable")
    df["dep"] = df["departure_time"].apply(time_to_minutes)
    if df["dep"].isna().any():
        raise ValueError("Timetable: some departure times are not valid")
    df["line"] = df["line"].astype(int)
    return df[["start", "end", "dep", "line"]].sort_values("dep").reset_index(drop=True)


def read_plan(file, distances):
    """Returns (plan, issues). Rows with a data error are reported and left out."""
    df = pd.read_excel(file, dtype=object)
    df.columns = [str(c).strip().lower() for c in df.columns]
    missing = [c for c in PLAN_COLUMNS if c not in df.columns]
    if missing:
        return None, [make_issue("DQ-COL", "Missing column(s): " + ", ".join(missing) + ".")]
    if len(df) == 0:
        return None, [make_issue("DQ-EMPTY", "The bus plan contains no data (only a header).")]

    locations = {key[0] for key in distances} | {key[1] for key in distances}
    lines = {key[2] for key in distances if key[2] is not None}
    issues, rows = [], []

    for i, r in df.iterrows():
        activity = str(r["activity"]).strip().lower()
        start, end = time_to_minutes(r["start time"]), time_to_minutes(r["end time"])
        line, energy, bus = to_number(r["line"]), to_number(r["energy consumption"]), to_number(r["bus"])
        errors = []
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
        if not errors and activity in ("service trip", "material trip") and \
                find_route(distances, r["start location"], r["end location"], line if activity == "service trip" else None) is None:
            errors.append(("DQ-ROUTE", f"No distance known from {r['start location']} to {r['end location']}."))

        for check, message in errors:
            issues.append(make_issue(check, message, bus=None if bus is None else int(bus), row=i + 2))
        if not errors:
            rows.append({"row": i + 2, "bus": int(bus), "start_loc": r["start location"], "end_loc": r["end location"],
                         "start": start, "end": end, "activity": activity, "line": line, "energy": energy})

    if not rows:
        return None, issues + [make_issue("DQ-EMPTY", "No usable rows in the bus plan.")]
    plan = pd.DataFrame(rows)
    fix_line_column(plan)
    for bus, bus_rows in plan.groupby("bus"):
        times = list(zip(bus_rows["start"], bus_rows["end"]))
        if times != sorted(times):
            issues.append(make_issue("DQ-ORDER", "Rows are not in time order (the tool sorts them).", "warning", bus=bus))
    return plan.sort_values(["bus", "start", "end"]).reset_index(drop=True), issues


# ----------------------------------------------------------------- 2. battery + feasibility checks
def charge_energy(soc, minutes, settings):
    """kWh charged: 450 kW up to 90%, 60 kW above 90%, never above 100%."""
    fast_minutes = min(minutes, max(0, MAX_SOC * capacity(settings) - soc) / FAST_KW * 60)
    charged = fast_minutes * FAST_KW / 60 + (minutes - fast_minutes) * SLOW_KW / 60
    return min(charged, capacity(settings) - soc)


def trip_km(a, distances):
    if a["activity"] == "service trip":
        return find_route(distances, a["start_loc"], a["end_loc"], a["line"])[2]
    if a["activity"] == "material trip":
        return find_route(distances, a["start_loc"], a["end_loc"])[2]
    return 0


def check_plan(plan, timetable, distances, settings=DEFAULT_SETTINGS):
    """Follows the battery of every bus and checks FC1-FC8. Returns (timeline, issues)."""
    full = capacity(settings)
    issues, timeline, driven = [], [], {}
    timetable_trips = set(zip(timetable["start"], timetable["dep"], timetable["end"], timetable["line"]))

    for bus, bus_rows in plan.groupby("bus"):
        soc, previous, too_low = settings["start_soc"] * full, None, None
        for _, a in bus_rows.iterrows():
            minutes = a["end"] - a["start"]
            time = minutes_to_time(a["start"])

            def error(check, message, severity="error"):
                issues.append(make_issue(check, message, severity, bus=bus, row=a["row"], time=time))

            if previous is not None:
                if a["start"] < previous["end"]:
                    error("FC4", f"Overlaps with the previous activity (ends {minutes_to_time(previous['end'])}).")
                if a["start_loc"] != previous["end_loc"]:
                    error("FC5", f"Bus is at {previous['end_loc']} but this activity starts at {a['start_loc']}.")
                soc -= max(0, a["start"] - previous["end"]) * IDLE_KW / 60   # waiting = standing still

            # energy according to our model
            if a["activity"] == "idle":
                energy = IDLE_KW * minutes / 60
            elif a["activity"] == "charging":
                energy = -charge_energy(soc, minutes, settings)
            else:
                energy = trip_km(a, distances) * settings["kwh_per_km"]
            if abs(a["energy"] - energy) > 0.05:
                error("DQ-ENERGY", f"Energy in file {a['energy']:.2f} kWh, model {energy:.2f} kWh.", "warning")
            if settings["use_file_energy"]:
                energy = a["energy"]
            soc -= energy
            soc_percent = round(soc / full * 100, 2)

            if a["activity"] == "charging":
                if minutes < MIN_CHARGE_TIME:
                    error("FC2", f"Charging session of {minutes} min (minimum {MIN_CHARGE_TIME}).")
                if soc_percent > MAX_SOC * 100:
                    error("FC3", f"Battery charged to {soc_percent:.2f}% (maximum 90%).")
                if a["start_loc"] != GARAGE:
                    error("FC6", f"Charging at {a['start_loc']}; only possible at {GARAGE}.")
            if a["activity"] == "service trip":
                trip = (a["start_loc"], a["start"], a["end_loc"], a["line"])
                if trip in timetable_trips:
                    driven.setdefault(trip, []).append((bus, a["row"]))
                else:
                    error("FC4", f"Service trip line {a['line']} at {time} is not in the timetable.")
            if a["activity"] == "material trip" and a["line"] is not None:
                error("FC4", f"Material trip has line number {a['line']}.")
            if a["activity"] in ("service trip", "material trip"):
                fastest = find_route(distances, a["start_loc"], a["end_loc"],
                                     a["line"] if a["activity"] == "service trip" else None)[0]
                if minutes < fastest:
                    error("FC7", f"Trip takes {minutes} min, minimum travel time is {fastest} min.")
            if soc_percent < MIN_SOC * 100 and too_low is None:
                too_low = make_issue("FC1", f"SOC drops to {soc_percent:.2f}% (minimum 10%).", bus=bus, row=a["row"], time=time)

            timeline.append({"bus": bus, "row": a["row"], "activity": a["activity"], "start_loc": a["start_loc"],
                             "end_loc": a["end_loc"], "start": a["start"], "end": a["end"], "duration": minutes,
                             "line": a["line"], "km": trip_km(a, distances), "energy": energy, "soc": soc_percent})
            previous = a
        if too_low:
            issues.append(too_low)

    # FC8: every timetable trip exactly once
    for trip in sorted(timetable_trips, key=lambda t: t[1]):
        start, dep, end, line = trip
        buses = driven.get(trip, [])
        if len(buses) == 0:
            issues.append(make_issue("FC8", f"Timetable trip {start}->{end} line {line} at {minutes_to_time(dep)} "
                                            "is not driven by any bus.", time=minutes_to_time(dep)))
        elif len(buses) > 1:
            issues.append(make_issue("FC8", f"Timetable trip {start}->{end} line {line} at {minutes_to_time(dep)} "
                                            f"is driven {len(buses)} times.", bus=buses[1][0], row=buses[1][1],
                                     time=minutes_to_time(dep)))
    return pd.DataFrame(timeline), issues


def is_feasible(issues):
    return all(issue["severity"] != "error" for issue in issues)


# ----------------------------------------------------------------- 3. KPIs
def kpis(timeline):
    material = timeline[timeline["activity"] == "material trip"]
    charging = timeline[timeline["activity"] == "charging"]
    return {
        "Buses used": int(timeline["bus"].nunique()),
        "Service trips": int((timeline["activity"] == "service trip").sum()),
        "Material trips": len(material),
        "Material trip time (min)": int(material["duration"].sum()),
        "Material trip distance (km)": round(float(material["km"].sum()), 2),
        "Charging sessions": len(charging),
        "Total charging time (min)": int(charging["duration"].sum()),
        "Idle time (min)": int(timeline[timeline["activity"] == "idle"]["duration"].sum()),
        "Energy used (kWh)": round(float(timeline[timeline["energy"] > 0]["energy"].sum()), 1),
        "Lowest SOC (%)": float(timeline["soc"].min()),
    }


# ----------------------------------------------------------------- 4. improved plan (greedy)
# Go through the timetable from early to late and give every trip to the bus that can do it
# with the least waiting. A bus only charges when it cannot do the trip otherwise.
# If no bus can do the trip, a new bus leaves the garage.

def act(start_loc, end_loc, start, end, activity, energy, line=None):
    return {"start_loc": start_loc, "end_loc": end_loc, "start": start, "end": end,
            "activity": activity, "line": line, "energy": energy}


def drive(start, end, distances, settings):
    """(minutes, kWh) of a material trip; zero if the bus is already there."""
    if start == end:
        return 0, 0
    _, minutes, km = find_route(distances, start, end)
    return minutes, km * settings["kwh_per_km"]


def try_trip(bus, trip, distances, settings, charge):
    """Can this bus do the trip (with or without charging first)? Returns an option or None."""
    acts, t, loc, cost = [], bus["free_at"], bus["location"], 0
    if charge:
        to_garage, e1 = drive(loc, GARAGE, distances, settings)
        if to_garage:
            acts.append(act(loc, GARAGE, t, t + to_garage, "material trip", e1))
        t, loc = t + to_garage, GARAGE
        soc = bus["soc"] - e1
        time_left = trip["dep"] - drive(GARAGE, trip["start"], distances, settings)[0] - t
        minutes = min(time_left, math.floor((MAX_SOC * capacity(settings) - soc) / FAST_KW * 60))
        if minutes < MIN_CHARGE_TIME:
            return None
        acts.append(act(GARAGE, GARAGE, t, t + minutes, "charging", -minutes * FAST_KW / 60))
        t, cost = t + minutes, 2 * to_garage + 10

    drive_time, drive_energy = drive(loc, trip["start"], distances, settings)
    wait = trip["dep"] - drive_time - t
    if wait < 0:
        return None
    if wait > 0:
        acts.append(act(loc, loc, t, t + wait, "idle", IDLE_KW * wait / 60))
    if drive_time > 0:
        acts.append(act(loc, trip["start"], trip["dep"] - drive_time, trip["dep"], "material trip", drive_energy))
    acts.append(act(trip["start"], trip["end"], trip["dep"], trip["arr"], "service trip", trip["energy"], trip["line"]))

    soc = bus["soc"] - sum(a["energy"] for a in acts)
    if soc - drive(trip["end"], GARAGE, distances, settings)[1] < MIN_SOC * capacity(settings):
        return None   # afterwards the bus could not reach the garage with 10%
    return {"acts": acts, "soc": soc, "cost": cost + wait + 2 * drive_time}


def improve(timetable, distances, settings=DEFAULT_SETTINGS):
    buses = []
    for _, t in timetable.iterrows():
        _, minutes, km = find_route(distances, t["start"], t["end"], t["line"])
        trip = {"start": t["start"], "end": t["end"], "dep": t["dep"], "arr": t["dep"] + minutes,
                "line": t["line"], "energy": km * settings["kwh_per_km"]}

        best_bus, best = None, None
        for bus in buses:
            option = try_trip(bus, trip, distances, settings, charge=False) or \
                     try_trip(bus, trip, distances, settings, charge=True)
            if option and (best is None or option["cost"] < best["cost"]):
                best_bus, best = bus, option
        if best is None:   # new bus from the garage
            best_bus = {"number": len(buses) + 1, "location": GARAGE, "acts": [],
                        "free_at": trip["dep"] - drive(GARAGE, trip["start"], distances, settings)[0],
                        "soc": settings["start_soc"] * capacity(settings)}
            best = try_trip(best_bus, trip, distances, settings, charge=False)
            buses.append(best_bus)
        best_bus["acts"] += best["acts"]
        best_bus.update(location=trip["end"], free_at=trip["arr"], soc=best["soc"])

    rows = []
    for bus in buses:   # end of the day: back to the garage
        minutes, energy = drive(bus["location"], GARAGE, distances, settings)
        if minutes:
            bus["acts"].append(act(bus["location"], GARAGE, bus["free_at"], bus["free_at"] + minutes, "material trip", energy))
        rows += [dict(a, bus=bus["number"]) for a in bus["acts"]]
    plan = pd.DataFrame(rows)
    plan["row"] = range(2, len(plan) + 2)
    fix_line_column(plan)
    return plan


# ----------------------------------------------------------------- 5. export to Excel (client format)
def to_excel(plan):
    table = pd.DataFrame({
        "start location": plan["start_loc"], "end location": plan["end_loc"],
        "start time": [minutes_to_time(m) + ":00" for m in plan["start"]],
        "end time": [minutes_to_time(m) + ":00" for m in plan["end"]],
        "activity": plan["activity"], "line": plan["line"],
        "energy consumption": plan["energy"].astype(float).round(4), "bus": plan["bus"]})
    buffer = io.BytesIO()
    table.to_excel(buffer, index=False, sheet_name="Sheet1")
    return buffer.getvalue()
