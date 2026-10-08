# Bus Plan Checker - all calculations: data checks, feasibility checks, KPIs, improved plan, Excel export
import io
import math
import pandas as pd

# Fixed values from the assignment
BATTERY_KWH = 300        # battery size of a new bus
MIN_SOC = 0.10           # battery never below 10%
MAX_SOC = 0.90           # never charge above 90%
MIN_CHARGE_TIME = 15     # charge at least 15 minutes
IDLE_KW = 5              # energy use while standing still
FAST_KW = 450            # charging speed up to 90%
SLOW_KW = 60             # charging speed above 90%
GARAGE = "ehvgar"        # only place with a charger
DEFAULT_SETTINGS = {"soh": 0.85, "kwh_per_km": 1.2, "start_soc": 1.0, "use_file_energy": False,
                    "min_idle": 0}   # min_idle = minutes standing still between two service trips (buffer for delays)
ACTIVITIES = ["service trip", "material trip", "idle", "charging"]
PLAN_COLUMNS = ["start location", "end location", "start time", "end time", "activity", "line", "energy consumption", "bus"]

# '06:04' or '06:04:00' -> 364 minutes; times before 03:00 count as after midnight (+24h)
def time_to_minutes(value):
    parts = str(value).strip().split(":")
    if len(parts) != 2 and len(parts) != 3:
        return None
    for p in parts:
        if not p.isdigit():
            return None
    hours = int(parts[0])
    minutes = int(parts[1])
    if hours > 23 or minutes > 59:
        return None
    total = hours * 60 + minutes
    if total < 180:
        total = total + 1440
    return total

# 364 -> '06:04'
def minutes_to_time(minutes):
    hours = int(minutes) // 60 % 24
    rest = int(minutes) % 60
    return f"{hours:02d}:{rest:02d}"

# cell -> number, or None if empty / not a number
def to_number(value):
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if math.isnan(number):
        return None
    return number

# one problem found by the tool
def make_issue(check, message, severity="error", bus=None, row=None, time=""):
    return {"severity": severity, "check": check, "bus": bus, "row": row, "time": time, "message": message}

# (min minutes, max minutes, km) between two locations; material trips use line=None
def find_route(distances, start, end, line=None):
    if (start, end, line) in distances:
        return distances[(start, end, line)]
    return distances.get((start, end, None))

# usable battery in kWh
def capacity(settings):
    return BATTERY_KWH * settings["soh"]

# pandas turns None into NaN and 401 into 401.0; we want None or a whole number
def fix_line_column(plan):
    lines = []
    for x in plan["line"]:
        if pd.isna(x):
            lines.append(None)
        else:
            lines.append(int(x))
    plan["line"] = pd.Series(lines, index=plan.index, dtype=object)

# stop with a clear message if a column is missing
def check_columns(df, columns, name):
    for c in columns:
        if c not in df.columns:
            raise ValueError(f"{name}: missing column {c}")

# ---------------------------------------------------------------- 1. read files + data checks
# distance matrix -> {(start, end, line): (min time, max time, km)}
def read_distances(file):
    df = pd.read_excel(file)
    check_columns(df, ["start", "end", "min_travel_time", "max_travel_time", "distance_m", "line"], "Distance matrix")
    distances = {}
    for _, r in df.iterrows():
        line = None
        if not pd.isna(r["line"]):
            line = int(r["line"])
        distances[(r["start"], r["end"], line)] = (int(r["min_travel_time"]), int(r["max_travel_time"]), r["distance_m"] / 1000)
    return distances

# timetable -> table with start, end, dep (minutes) and line, sorted on departure time
# every row must be complete; with the distance matrix we also check that the route exists
def read_timetable(file, distances=None):
    df = pd.read_excel(file, dtype=object)
    check_columns(df, ["start", "departure_time", "end", "line"], "Timetable")
    if len(df) == 0:
        raise ValueError("Timetable: the file contains no trips")
    for i, r in df.iterrows():
        row = i + 2   # row 1 is the header
        if time_to_minutes(r["departure_time"]) is None:
            raise ValueError(f"Timetable row {row}: departure time '{r['departure_time']}' is not valid")
        if pd.isna(r["start"]) or pd.isna(r["end"]) or to_number(r["line"]) is None:
            raise ValueError(f"Timetable row {row}: start, end or line is empty or not valid")
        if distances is not None and find_route(distances, r["start"], r["end"], int(to_number(r["line"]))) is None:
            raise ValueError(f"Timetable row {row}: no route from {r['start']} to {r['end']} (line {r['line']}) "
                             f"in the distance matrix")
    df["dep"] = df["departure_time"].apply(time_to_minutes)
    df["line"] = df["line"].astype(float).astype(int)
    df = df.sort_values("dep").reset_index(drop=True)
    return df[["start", "end", "dep", "line"]]

# bus plan -> (plan, issues); rows with a data error are reported and left out
def read_plan(file, distances):
    df = pd.read_excel(file, dtype=object)
    df.columns = [str(c).strip().lower() for c in df.columns]
    missing = []
    for c in PLAN_COLUMNS:
        if c not in df.columns:
            missing.append(c)
    if missing:
        return None, [make_issue("DQ-COL", "Missing column(s): " + ", ".join(missing) + ".")]

    # all known locations and lines from the distance matrix
    locations = set()
    lines = set()
    for start, end, line in distances:
        locations.add(start)
        locations.add(end)
        if line is not None:
            lines.add(line)

    issues = []
    rows = []
    for i, r in df.iterrows():
        excel_row = i + 2   # row 1 is the header
        activity = str(r["activity"]).strip().lower()
        start = time_to_minutes(r["start time"])
        end = time_to_minutes(r["end time"])
        line = to_number(r["line"])
        energy = to_number(r["energy consumption"])
        bus = to_number(r["bus"])

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
        if len(errors) == 0 and "trip" in activity:
            route_line = line if activity == "service trip" else None
            if find_route(distances, r["start location"], r["end location"], route_line) is None:
                errors.append(("DQ-ROUTE", f"No distance known from {r['start location']} to {r['end location']}."))

        if len(errors) > 0:
            for check, message in errors:
                bus_number = None if bus is None else int(bus)
                issues.append(make_issue(check, message, bus=bus_number, row=excel_row))
        else:
            rows.append({"row": excel_row, "bus": int(bus), "start_loc": r["start location"], "end_loc": r["end location"],
                         "start": start, "end": end, "activity": activity, "line": line, "energy": energy})

    if len(rows) == 0:
        issues.append(make_issue("DQ-EMPTY", "No usable rows in the bus plan."))
        return None, issues
    plan = pd.DataFrame(rows)
    fix_line_column(plan)
    return plan, issues   # rows stay in the order of the file: the plan is checked exactly as it was given

# ---------------------------------------------------------------- 2. battery + feasibility checks
# kWh charged: 450 kW up to 90%, 60 kW above 90%, never above 100%
def charge_energy(soc, minutes, settings):
    full = capacity(settings)
    fast_minutes = (MAX_SOC * full - soc) / FAST_KW * 60   # minutes needed to reach 90%
    fast_minutes = max(0, min(minutes, fast_minutes))
    slow_minutes = minutes - fast_minutes
    energy = fast_minutes * FAST_KW / 60 + slow_minutes * SLOW_KW / 60
    return min(energy, full - soc)

# follow the battery of every bus and check FC1-FC8 -> (timeline, issues)
def check_plan(plan, timetable, distances, settings=DEFAULT_SETTINGS):
    full = capacity(settings)
    issues = []
    timeline = []
    driven = {}   # timetable trip -> list of (bus, row) that drive it
    timetable_trips = set(zip(timetable["start"], timetable["dep"], timetable["end"], timetable["line"]))
    service_trips = 0        # number of service trips in the bus plan
    not_in_timetable = 0     # service trips of the bus plan that are not in the timetable

    for bus, bus_rows in plan.groupby("bus"):
        soc = settings["start_soc"] * full
        previous = None
        too_low = None
        idle_since_trip = None   # minutes standing still since the last service trip (None = no trip yet)
        for _, a in bus_rows.iterrows():
            minutes = a["end"] - a["start"]
            time = minutes_to_time(a["start"])
            row = a["row"]

            # shortest travel time and km (only for trips)
            fastest, km = 0, 0
            if "trip" in a["activity"]:
                route = find_route(distances, a["start_loc"], a["end_loc"], a["line"])
                fastest, km = route[0], route[2]

            # FC4 overlap and FC5 location, compared with the previous row; a gap counts as standing still
            if previous is not None:
                if a["start"] < previous["end"]:
                    issues.append(make_issue("FC4", f"Starts at {time}, before the previous row (row {previous['row']}) "
                                             f"ends at {minutes_to_time(previous['end'])}: overlap or rows in the wrong order.",
                                             bus=bus, row=row, time=time))
                if a["start_loc"] != previous["end_loc"]:
                    issues.append(make_issue("FC5", f"Bus is at {previous['end_loc']} but this activity starts at {a['start_loc']}.",
                                             bus=bus, row=row, time=time))
                gap = max(0, a["start"] - previous["end"])
                soc = soc - gap * IDLE_KW / 60
                if idle_since_trip is not None:
                    idle_since_trip = idle_since_trip + gap

            # energy according to our model, compared with the file
            if a["activity"] == "idle":
                energy = IDLE_KW * minutes / 60
                how = f"{minutes} min × {IDLE_KW} kW"
            elif a["activity"] == "charging":
                energy = -charge_energy(soc, minutes, settings)
                how = f"{minutes} min charging"
            else:
                energy = km * settings["kwh_per_km"]
                how = f"{km:.2f} km × {settings['kwh_per_km']} kWh/km"
            if abs(a["energy"] - energy) > 0.05:
                issues.append(make_issue("DQ-ENERGY", f"{a['activity'].capitalize()}: file says {a['energy']:.2f} kWh, "
                                         f"should be {energy:.2f} kWh ({how}).", "warning", bus=bus, row=row, time=time))
            if settings["use_file_energy"]:
                soc = soc - a["energy"]
            else:
                soc = soc - energy
            soc_percent = round(soc / full * 100, 2)

            # FC2, FC3, FC6: charging rules
            if a["activity"] == "charging":
                if minutes < MIN_CHARGE_TIME:
                    issues.append(make_issue("FC2", f"Charging session of {minutes} min (minimum {MIN_CHARGE_TIME}).",
                                             bus=bus, row=row, time=time))
                if soc_percent > MAX_SOC * 100:
                    issues.append(make_issue("FC3", f"Battery charged to {soc_percent:.2f}% (maximum 90%).",
                                             bus=bus, row=row, time=time))
                if a["start_loc"] != GARAGE:
                    issues.append(make_issue("FC6", f"Charging at {a['start_loc']}; only possible at {GARAGE}.",
                                             bus=bus, row=row, time=time))

            # FC4: service trip must be in the timetable, material trip has no line
            trip = (a["start_loc"], a["start"], a["end_loc"], a["line"])
            if a["activity"] == "service trip":
                service_trips = service_trips + 1
                if trip in timetable_trips:
                    if trip not in driven:
                        driven[trip] = []
                    driven[trip].append((bus, row))
                else:
                    not_in_timetable = not_in_timetable + 1
                    issues.append(make_issue("FC4", f"Service trip line {a['line']} at {time} is not in the timetable.",
                                             bus=bus, row=row, time=time))
            if a["activity"] == "material trip" and a["line"] is not None:
                issues.append(make_issue("FC4", f"Material trip has line number {a['line']}.", bus=bus, row=row, time=time))

            # FC9: enough idle time (buffer) between two service trips
            if a["activity"] == "idle" and idle_since_trip is not None:
                idle_since_trip = idle_since_trip + minutes
            if a["activity"] == "service trip":
                if idle_since_trip is not None and idle_since_trip < settings["min_idle"]:
                    issues.append(make_issue("FC9", f"Only {idle_since_trip} min idle before this trip (minimum {settings['min_idle']}).",
                                             bus=bus, row=row, time=time))
                idle_since_trip = 0

            # FC7: not faster than the shortest travel time
            if minutes < fastest:
                issues.append(make_issue("FC7", f"Trip takes {minutes} min, minimum travel time is {fastest} min.",
                                         bus=bus, row=row, time=time))
            # FC1: remember the first time the battery is below 10%
            if soc_percent < MIN_SOC * 100 and too_low is None:
                too_low = make_issue("FC1", f"SOC drops to {soc_percent:.2f}% (minimum 10%).", bus=bus, row=row, time=time)

            timeline.append({"bus": bus, "row": row, "activity": a["activity"], "start_loc": a["start_loc"],
                             "end_loc": a["end_loc"], "start": a["start"], "end": a["end"], "duration": minutes,
                             "line": a["line"], "km": km, "energy": energy, "soc": soc_percent})
            previous = a
        if too_low is not None:
            issues.append(too_low)

    # FC8: every timetable trip driven exactly once
    not_once = 0
    for _, t in timetable.iterrows():
        trip = (t["start"], t["dep"], t["end"], t["line"])
        buses = driven.get(trip, [])
        if len(buses) != 1:
            time = minutes_to_time(t["dep"])
            if len(buses) == 0:
                bus, row = None, None
            else:
                bus, row = buses[1]   # the second bus that drives it
            issues.append(make_issue("FC8", f"Trip {t['start']}->{t['end']} line {t['line']} at {time} is driven {len(buses)} times.",
                                     bus=bus, row=row, time=time))
            not_once = not_once + 1

    # DQ-MATCH: do the rows of the timetable and the service trips of the bus plan belong together?
    if not_in_timetable > 0 or not_once > 0:
        issues.append(make_issue("DQ-MATCH", f"{not_in_timetable} of the {service_trips} service trips in the bus plan are not "
                                 f"in the timetable, and {not_once} of the {len(timetable)} timetable trips are not driven "
                                 f"exactly once. Check that the timetable belongs to this bus plan (details: FC4 and FC8).",
                                 "warning"))
    return pd.DataFrame(timeline), issues

# feasible = no errors (warnings are allowed)
def is_feasible(issues):
    for issue in issues:
        if issue["severity"] == "error":
            return False
    return True

# ---------------------------------------------------------------- 3. KPIs
def kpis(timeline):
    service = timeline[timeline["activity"] == "service trip"]
    material = timeline[timeline["activity"] == "material trip"]
    charging = timeline[timeline["activity"] == "charging"]
    idle = timeline[timeline["activity"] == "idle"]
    used = timeline[timeline["energy"] > 0]
    return {"Buses used": int(timeline["bus"].nunique()),
            "Service trips": len(service),
            "Material trips": len(material),
            "Material trip time (min)": int(material["duration"].sum()),
            "Material trip distance (km)": round(float(material["km"].sum()), 2),
            "Charging sessions": len(charging),
            "Total charging time (min)": int(charging["duration"].sum()),
            "Idle time (min)": int(idle["duration"].sum()),
            "Energy used (kWh)": round(float(used["energy"].sum()), 1),
            "Lowest SOC (%)": float(timeline["soc"].min())}

# ---------------------------------------------------------------- 4. improved plan (greedy)
# one activity of the new plan
def act(start_loc, end_loc, start, end, activity, energy, line=None):
    return {"start_loc": start_loc, "end_loc": end_loc, "start": start, "end": end, "activity": activity, "line": line, "energy": energy}

# (minutes, kWh) of a material trip; zero if the bus is already there
def drive(start, end, distances, settings):
    if start == end:
        return 0, 0
    route = find_route(distances, start, end)
    minutes = route[1]
    km = route[2]
    return minutes, km * settings["kwh_per_km"]

# can this bus do the trip (first charging if charge=True)? returns activities, new SOC and cost, or None
def try_trip(bus, trip, distances, settings, charge):
    acts = []
    t = bus["free_at"]
    loc = bus["location"]
    soc = bus["soc"]
    cost = 0
    if charge:
        # drive to the garage
        minutes_to_garage, energy = drive(loc, GARAGE, distances, settings)
        if minutes_to_garage > 0:
            acts.append(act(loc, GARAGE, t, t + minutes_to_garage, "material trip", energy))
        t = t + minutes_to_garage
        loc = GARAGE
        soc = soc - energy
        cost = 2 * minutes_to_garage + 10   # penalty: charging costs time and driving
        # charge as long as possible, but not above 90%
        minutes_to_start = drive(GARAGE, trip["start"], distances, settings)[0]
        time_left = trip["dep"] - minutes_to_start - t - settings["min_idle"]   # keep time for the buffer
        minutes_to_90 = math.floor((MAX_SOC * capacity(settings) - soc) / FAST_KW * 60)
        minutes = min(time_left, minutes_to_90)
        if minutes < MIN_CHARGE_TIME:
            return None
        acts.append(act(GARAGE, GARAGE, t, t + minutes, "charging", -minutes * FAST_KW / 60))
        t = t + minutes
    # wait, drive to the start of the trip and do the trip
    drive_time, drive_energy = drive(loc, trip["start"], distances, settings)
    wait = trip["dep"] - drive_time - t
    if wait < settings["min_idle"]:
        return None   # too late, or not enough buffer before the trip
    if wait > 0:
        acts.append(act(loc, loc, t, t + wait, "idle", IDLE_KW * wait / 60))
    if drive_time > 0:
        acts.append(act(loc, trip["start"], trip["dep"] - drive_time, trip["dep"], "material trip", drive_energy))
    acts.append(act(trip["start"], trip["end"], trip["dep"], trip["arr"], "service trip", trip["energy"], trip["line"]))
    # new battery level; afterwards the bus must still reach the garage with 10%
    soc = bus["soc"]
    for a in acts:
        soc = soc - a["energy"]
    energy_to_garage = drive(trip["end"], GARAGE, distances, settings)[1]
    if soc - energy_to_garage < MIN_SOC * capacity(settings):
        return None
    return {"acts": acts, "soc": soc, "cost": cost + wait + 2 * drive_time}

# every trip (early to late) goes to the bus with the lowest cost; charge only if needed; else a new bus
def improve(timetable, distances, settings=DEFAULT_SETTINGS):
    buses = []
    for _, t in timetable.iterrows():
        route = find_route(distances, t["start"], t["end"], t["line"])
        trip = {"start": t["start"], "end": t["end"], "dep": t["dep"], "arr": t["dep"] + route[1],
                "line": t["line"], "energy": route[2] * settings["kwh_per_km"]}
        # find the cheapest bus that can do this trip
        best_bus = None
        best = None
        for bus in buses:
            option = try_trip(bus, trip, distances, settings, False)
            if option is None:
                option = try_trip(bus, trip, distances, settings, True)
            if option is not None and (best is None or option["cost"] < best["cost"]):
                best_bus = bus
                best = option
        # no bus can do it: add a new bus that starts at the garage
        if best is None:
            minutes_from_garage = drive(GARAGE, trip["start"], distances, settings)[0]
            best_bus = {"number": len(buses) + 1, "location": GARAGE, "acts": [],
                        "soc": settings["start_soc"] * capacity(settings), "free_at": trip["dep"] - minutes_from_garage - settings["min_idle"]}
            best = try_trip(best_bus, trip, distances, settings, False)
            buses.append(best_bus)
        best_bus["acts"] = best_bus["acts"] + best["acts"]
        best_bus["location"] = trip["end"]
        best_bus["free_at"] = trip["arr"]
        best_bus["soc"] = best["soc"]

    # end of the day: every bus drives back to the garage
    rows = []
    for bus in buses:
        minutes, energy = drive(bus["location"], GARAGE, distances, settings)
        if minutes > 0:
            bus["acts"].append(act(bus["location"], GARAGE, bus["free_at"], bus["free_at"] + minutes, "material trip", energy))
        for a in bus["acts"]:
            a["bus"] = bus["number"]
            rows.append(a)
    plan = pd.DataFrame(rows)
    plan["row"] = range(2, len(plan) + 2)
    fix_line_column(plan)
    return plan

# ---------------------------------------------------------------- 5. export in the client's Excel format
def to_excel(plan):
    table = pd.DataFrame()
    table["start location"] = plan["start_loc"]
    table["end location"] = plan["end_loc"]
    table["start time"] = [minutes_to_time(m) + ":00" for m in plan["start"]]
    table["end time"] = [minutes_to_time(m) + ":00" for m in plan["end"]]
    table["activity"] = plan["activity"]
    table["line"] = plan["line"]
    table["energy consumption"] = plan["energy"].astype(float).round(4)
    table["bus"] = plan["bus"]
    buffer = io.BytesIO()
    table.to_excel(buffer, index=False, sheet_name="Sheet1")
    return buffer.getvalue()
