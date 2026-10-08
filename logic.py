# Bus Plan Checker - all calculations: data checks, feasibility checks, KPIs, improved plan, Excel export
import io
import math
import pandas as pd

# Fixed values from the assignment
BATTERY_KWH, MIN_SOC, MAX_SOC = 300, 0.10, 0.90   # battery size, never below 10%, never charged above 90%
MIN_CHARGE_TIME = 15                              # charge at least 15 minutes
IDLE_KW, FAST_KW, SLOW_KW = 5, 450, 60            # standing still, charging up to 90%, charging above 90%
GARAGE = "ehvgar"                                 # only place with a charger
DEFAULT_SETTINGS = {"soh": 0.85, "kwh_per_km": 1.2, "start_soc": 0.90, "use_file_energy": False}
ACTIVITIES = ["service trip", "material trip", "idle", "charging"]
PLAN_COLUMNS = ["start location", "end location", "start time", "end time", "activity", "line", "energy consumption", "bus"]

# '06:04' or '06:04:00' -> 364 minutes; times before 03:00 count as after midnight (+24h)
def time_to_minutes(value):
    parts = str(value).strip().split(":")
    if len(parts) not in (2, 3) or not all(p.isdigit() for p in parts) or int(parts[0]) > 23 or int(parts[1]) > 59:
        return None
    total = int(parts[0]) * 60 + int(parts[1])
    return total + 1440 if total < 180 else total

# 364 -> '06:04'
def minutes_to_time(minutes):
    return f"{int(minutes) // 60 % 24:02d}:{int(minutes) % 60:02d}"

# cell -> number, or None if empty / not a number
def to_number(value):
    try:
        return None if math.isnan(float(value)) else float(value)
    except (TypeError, ValueError):
        return None

# one problem found by the tool
def make_issue(check, message, severity="error", bus=None, row=None, time=""):
    return {"severity": severity, "check": check, "bus": bus, "row": row, "time": time, "message": message}

# (min minutes, max minutes, km) between two locations; material trips use line=None
def find_route(distances, start, end, line=None):
    return distances.get((start, end, line)) or distances.get((start, end, None))

# usable battery in kWh
def capacity(settings):
    return BATTERY_KWH * settings["soh"]

# pandas turns None into NaN and 401 into 401.0; we want None or a whole number
def fix_line_column(plan):
    plan["line"] = pd.Series([None if pd.isna(x) else int(x) for x in plan["line"]], index=plan.index, dtype=object)

# stop with a clear message if a column is missing
def check_columns(df, columns, name):
    missing = [c for c in columns if c not in df.columns]
    if missing:
        raise ValueError(f"{name}: missing column(s) {', '.join(missing)}")

# ---------------------------------------------------------------- 1. read files + data checks
# distance matrix -> {(start, end, line): (min time, max time, km)}
def read_distances(file):
    df = pd.read_excel(file)
    check_columns(df, ["start", "end", "min_travel_time", "max_travel_time", "distance_m", "line"], "Distance matrix")
    return {(r["start"], r["end"], None if pd.isna(r["line"]) else int(r["line"])):
            (int(r["min_travel_time"]), int(r["max_travel_time"]), r["distance_m"] / 1000) for _, r in df.iterrows()}

# timetable -> table with start, end, dep (minutes) and line
def read_timetable(file):
    df = pd.read_excel(file, dtype=object)
    check_columns(df, ["start", "departure_time", "end", "line"], "Timetable")
    df["dep"] = df["departure_time"].apply(time_to_minutes)
    if df["dep"].isna().any():
        raise ValueError("Timetable: some departure times are not valid")
    df["line"] = df["line"].astype(int)
    return df[["start", "end", "dep", "line"]].sort_values("dep").reset_index(drop=True)

# bus plan -> (plan, issues); rows with a data error are reported and left out
def read_plan(file, distances):
    df = pd.read_excel(file, dtype=object)
    df.columns = [str(c).strip().lower() for c in df.columns]
    missing = [c for c in PLAN_COLUMNS if c not in df.columns]
    if missing:
        return None, [make_issue("DQ-COL", "Missing column(s): " + ", ".join(missing) + ".")]
    if len(df) == 0:
        return None, [make_issue("DQ-EMPTY", "The bus plan contains no data (only a header).")]
    locations = {k[0] for k in distances} | {k[1] for k in distances}
    lines = {k[2] for k in distances if k[2] is not None}
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
        if not errors and activity in ("service trip", "material trip") and find_route(
                distances, r["start location"], r["end location"], line if activity == "service trip" else None) is None:
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
    # rows of a bus that are not in time order in the file (the tool sorts them afterwards)
    for bus, bus_rows in plan.groupby("bus"):
        previous = None
        for _, r in bus_rows.iterrows():
            if previous is not None and (r["start"], r["end"]) < (previous["start"], previous["end"]):
                issues.append(make_issue("DQ-ORDER", f"Row {r['row']} ({minutes_to_time(r['start'])}-{minutes_to_time(r['end'])}) "
                                         f"should come before row {previous['row']} ({minutes_to_time(previous['start'])}-"
                                         f"{minutes_to_time(previous['end'])}).", "warning", bus=bus, row=r["row"],
                                         time=minutes_to_time(r["start"])))
                break
            previous = r
    return plan.sort_values(["bus", "start", "end"]).reset_index(drop=True), issues

# ---------------------------------------------------------------- 2. battery + feasibility checks
# kWh charged: 450 kW up to 90%, 60 kW above 90%, never above 100%
def charge_energy(soc, minutes, settings):
    fast_minutes = min(minutes, max(0, MAX_SOC * capacity(settings) - soc) / FAST_KW * 60)
    return min(fast_minutes * FAST_KW / 60 + (minutes - fast_minutes) * SLOW_KW / 60, capacity(settings) - soc)

# (shortest time, km) of a trip; (0, 0) for idle and charging
def trip_info(a, distances):
    if a["activity"] not in ("service trip", "material trip"):
        return 0, 0
    route = find_route(distances, a["start_loc"], a["end_loc"], a["line"] if a["activity"] == "service trip" else None)
    return route[0], route[2]

# follow the battery of every bus and check FC1-FC8 -> (timeline, issues)
def check_plan(plan, timetable, distances, settings=DEFAULT_SETTINGS):
    full = capacity(settings)
    issues, timeline, driven = [], [], {}
    timetable_trips = set(zip(timetable["start"], timetable["dep"], timetable["end"], timetable["line"]))
    for bus, bus_rows in plan.groupby("bus"):
        soc, previous, too_low = settings["start_soc"] * full, None, None
        for _, a in bus_rows.iterrows():
            minutes, time = a["end"] - a["start"], minutes_to_time(a["start"])
            fastest, km = trip_info(a, distances)
            def error(check, message, severity="error"):
                issues.append(make_issue(check, message, severity, bus=bus, row=a["row"], time=time))
            # FC4 overlap and FC5 location, compared with the previous activity; a gap counts as standing still
            if previous is not None:
                if a["start"] < previous["end"]:
                    error("FC4", f"Overlaps with the previous activity (ends {minutes_to_time(previous['end'])}).")
                if a["start_loc"] != previous["end_loc"]:
                    error("FC5", f"Bus is at {previous['end_loc']} but this activity starts at {a['start_loc']}.")
                soc -= max(0, a["start"] - previous["end"]) * IDLE_KW / 60
            # energy according to our model, compared with the file
            if a["activity"] == "idle":
                energy, how = IDLE_KW * minutes / 60, f"{minutes} min × {IDLE_KW} kW"
            elif a["activity"] == "charging":
                energy, how = -charge_energy(soc, minutes, settings), f"{minutes} min charging"
            else:
                energy, how = km * settings["kwh_per_km"], f"{km:.2f} km × {settings['kwh_per_km']} kWh/km"
            if abs(a["energy"] - energy) > 0.05:
                error("DQ-ENERGY", f"{a['activity'].capitalize()}: file says {a['energy']:.2f} kWh, "
                                   f"should be {energy:.2f} kWh ({how}).", "warning")
            soc -= a["energy"] if settings["use_file_energy"] else energy
            soc_percent = round(soc / full * 100, 2)
            # FC2, FC3, FC6: charging rules
            if a["activity"] == "charging":
                if minutes < MIN_CHARGE_TIME:
                    error("FC2", f"Charging session of {minutes} min (minimum {MIN_CHARGE_TIME}).")
                if soc_percent > MAX_SOC * 100:
                    error("FC3", f"Battery charged to {soc_percent:.2f}% (maximum 90%).")
                if a["start_loc"] != GARAGE:
                    error("FC6", f"Charging at {a['start_loc']}; only possible at {GARAGE}.")
            # FC4: service trip must be in the timetable, material trip has no line
            trip = (a["start_loc"], a["start"], a["end_loc"], a["line"])
            if a["activity"] == "service trip" and trip in timetable_trips:
                driven.setdefault(trip, []).append((bus, a["row"]))
            elif a["activity"] == "service trip":
                error("FC4", f"Service trip line {a['line']} at {time} is not in the timetable.")
            if a["activity"] == "material trip" and a["line"] is not None:
                error("FC4", f"Material trip has line number {a['line']}.")
            # FC7: not faster than the shortest travel time; FC1: remember first time below 10%
            if minutes < fastest:
                error("FC7", f"Trip takes {minutes} min, minimum travel time is {fastest} min.")
            if soc_percent < MIN_SOC * 100 and too_low is None:
                too_low = make_issue("FC1", f"SOC drops to {soc_percent:.2f}% (minimum 10%).", bus=bus, row=a["row"], time=time)
            timeline.append({"bus": bus, "row": a["row"], "activity": a["activity"], "start_loc": a["start_loc"],
                             "end_loc": a["end_loc"], "start": a["start"], "end": a["end"], "duration": minutes,
                             "line": a["line"], "km": km, "energy": energy, "soc": soc_percent})
            previous = a
        if too_low:
            issues.append(too_low)
    # FC8: every timetable trip driven exactly once
    for trip in sorted(timetable_trips, key=lambda t: t[1]):
        buses, text = driven.get(trip, []), f"Timetable trip {trip[0]}->{trip[2]} line {trip[3]} at {minutes_to_time(trip[1])}"
        if len(buses) != 1:
            bus, row = buses[1] if buses else (None, None)
            issues.append(make_issue("FC8", text + (f" is driven {len(buses)} times." if buses else " is not driven by any bus."),
                                     bus=bus, row=row, time=minutes_to_time(trip[1])))
    return pd.DataFrame(timeline), issues

# feasible = no errors (warnings are allowed)
def is_feasible(issues):
    return all(issue["severity"] != "error" for issue in issues)

# ---------------------------------------------------------------- 3. KPIs
def kpis(timeline):
    material, charging = timeline[timeline["activity"] == "material trip"], timeline[timeline["activity"] == "charging"]
    return {"Buses used": int(timeline["bus"].nunique()),
            "Service trips": int((timeline["activity"] == "service trip").sum()),
            "Material trips": len(material),
            "Material trip time (min)": int(material["duration"].sum()),
            "Material trip distance (km)": round(float(material["km"].sum()), 2),
            "Charging sessions": len(charging),
            "Total charging time (min)": int(charging["duration"].sum()),
            "Idle time (min)": int(timeline[timeline["activity"] == "idle"]["duration"].sum()),
            "Energy used (kWh)": round(float(timeline[timeline["energy"] > 0]["energy"].sum()), 1),
            "Lowest SOC (%)": float(timeline["soc"].min())}

# ---------------------------------------------------------------- 4. improved plan (greedy)
# one activity of the new plan
def act(start_loc, end_loc, start, end, activity, energy, line=None):
    return {"start_loc": start_loc, "end_loc": end_loc, "start": start, "end": end, "activity": activity, "line": line, "energy": energy}

# (minutes, kWh) of a material trip; zero if the bus is already there
def drive(start, end, distances, settings):
    if start == end:
        return 0, 0
    _, minutes, km = find_route(distances, start, end)
    return minutes, km * settings["kwh_per_km"]

# can this bus do the trip (first charging if charge=True)? returns activities, new SOC and cost, or None
def try_trip(bus, trip, distances, settings, charge):
    acts, t, loc, soc, cost = [], bus["free_at"], bus["location"], bus["soc"], 0
    if charge:
        minutes_to_garage, energy = drive(loc, GARAGE, distances, settings)
        if minutes_to_garage:
            acts.append(act(loc, GARAGE, t, t + minutes_to_garage, "material trip", energy))
        t, loc, soc, cost = t + minutes_to_garage, GARAGE, soc - energy, 2 * minutes_to_garage + 10
        time_left = trip["dep"] - drive(GARAGE, trip["start"], distances, settings)[0] - t
        minutes = min(time_left, math.floor((MAX_SOC * capacity(settings) - soc) / FAST_KW * 60))
        if minutes < MIN_CHARGE_TIME:
            return None
        acts.append(act(GARAGE, GARAGE, t, t + minutes, "charging", -minutes * FAST_KW / 60))
        t += minutes
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

# every trip (early to late) goes to the bus with the lowest cost; charge only if needed; else a new bus
def improve(timetable, distances, settings=DEFAULT_SETTINGS):
    buses = []
    for _, t in timetable.iterrows():
        _, minutes, km = find_route(distances, t["start"], t["end"], t["line"])
        trip = {"start": t["start"], "end": t["end"], "dep": t["dep"], "arr": t["dep"] + minutes,
                "line": t["line"], "energy": km * settings["kwh_per_km"]}
        best_bus, best = None, None
        for bus in buses:
            option = try_trip(bus, trip, distances, settings, False) or try_trip(bus, trip, distances, settings, True)
            if option and (best is None or option["cost"] < best["cost"]):
                best_bus, best = bus, option
        if best is None:
            best_bus = {"number": len(buses) + 1, "location": GARAGE, "acts": [], "soc": settings["start_soc"] * capacity(settings),
                        "free_at": trip["dep"] - drive(GARAGE, trip["start"], distances, settings)[0]}
            best = try_trip(best_bus, trip, distances, settings, False)
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

# ---------------------------------------------------------------- 5. export in the client's Excel format
def to_excel(plan):
    table = pd.DataFrame({"start location": plan["start_loc"], "end location": plan["end_loc"],
                          "start time": [minutes_to_time(m) + ":00" for m in plan["start"]],
                          "end time": [minutes_to_time(m) + ":00" for m in plan["end"]],
                          "activity": plan["activity"], "line": plan["line"],
                          "energy consumption": plan["energy"].astype(float).round(4), "bus": plan["bus"]})
    buffer = io.BytesIO()
    table.to_excel(buffer, index=False, sheet_name="Sheet1")
    return buffer.getvalue()
