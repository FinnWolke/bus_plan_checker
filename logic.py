"""
Bus Plan Checker - all calculations.

Contents:
  1. Reading the Excel files + data checks (DQ)
  2. Battery simulation + feasibility checks (FC1 - FC8)
  3. KPIs
  4. Making an improved bus plan
  5. Saving a plan to Excel
"""
import io
import math

import pandas as pd

# ---------------------------------------------------------------------------
# Fixed values from the assignment
# ---------------------------------------------------------------------------
BATTERY_KWH = 300        # original battery capacity
MIN_SOC = 0.10           # battery may never go below 10%
MAX_SOC = 0.90           # battery may not be charged above 90% during the day
MIN_CHARGE_TIME = 15     # charging takes at least 15 minutes
IDLE_KW = 5              # energy use when the bus stands still (kW)
FAST_CHARGE_KW = 450     # charging speed up to 90%
SLOW_CHARGE_KW = 60      # charging speed above 90%
GARAGE = "ehvgar"        # the only place with a charger

# Values that can be changed in the app (or in the tests)
DEFAULT_SETTINGS = {
    "soh": 0.85,               # State of Health
    "kwh_per_km": 1.2,         # energy use while driving
    "start_soc": 0.90,         # battery level at the start of the day
    "use_file_energy": False,  # True = use the 'energy consumption' column of the file
}

ACTIVITIES = ["service trip", "material trip", "idle", "charging"]
PLAN_COLUMNS = ["start location", "end location", "start time", "end time",
                "activity", "line", "energy consumption", "bus"]


# ---------------------------------------------------------------------------
# Small helper functions
# ---------------------------------------------------------------------------
def time_to_minutes(value):
    """'06:04' or '06:04:00' -> 364. Returns None if it is not a valid time.
    Times before 03:00 belong to the end of the day, so we add 24 hours."""
    parts = str(value).strip().split(":")
    if len(parts) < 2 or len(parts) > 3:
        return None
    for part in parts:
        if not part.isdigit():
            return None
    hours = int(parts[0])
    minutes = int(parts[1])
    if hours > 23 or minutes > 59:
        return None
    total = hours * 60 + minutes
    if total < 3 * 60:
        total = total + 24 * 60
    return total


def minutes_to_time(minutes):
    """364 -> '06:04'"""
    minutes = int(minutes)
    return f"{(minutes // 60) % 24:02d}:{minutes % 60:02d}"


def to_number(value):
    """Turns a cell into a number. Returns None if the cell is empty or not a number."""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if math.isnan(number):
        return None
    return number


def make_issue(check, message, severity="error", bus=None, row=None, time=""):
    """One problem found by the tool."""
    return {"severity": severity, "check": check, "bus": bus, "row": row, "time": time, "message": message}


def find_route(distances, start, end, line=None):
    """Returns (min minutes, max minutes, km) between two locations.
    For a material trip line is None."""
    if (start, end, line) in distances:
        return distances[(start, end, line)]
    if (start, end, None) in distances:
        return distances[(start, end, None)]
    return None


def fix_line_column(plan):
    """pandas turns None into NaN and 401 into 401.0. We want None or a whole number."""
    clean = [None if pd.isna(x) else int(x) for x in plan["line"]]
    plan["line"] = pd.Series(clean, index=plan.index, dtype=object)


def capacity(settings):
    """Usable battery capacity in kWh (300 kWh x State of Health)."""
    return BATTERY_KWH * settings["soh"]


# ---------------------------------------------------------------------------
# 1. Reading the files + data checks
# ---------------------------------------------------------------------------
def read_distances(file):
    """Distance matrix -> dictionary {(start, end, line): (min time, max time, km)}"""
    df = pd.read_excel(file)
    for column in ["start", "end", "min_travel_time", "max_travel_time", "distance_m", "line"]:
        if column not in df.columns:
            raise ValueError(f"Distance matrix: column '{column}' is missing")

    distances = {}
    for _, row in df.iterrows():
        if pd.isna(row["line"]):
            line = None
        else:
            line = int(row["line"])
        distances[(row["start"], row["end"], line)] = (
            int(row["min_travel_time"]), int(row["max_travel_time"]), row["distance_m"] / 1000)
    return distances


def read_timetable(file):
    """Timetable -> table with start, end, dep (minutes) and line."""
    df = pd.read_excel(file, dtype=object)
    for column in ["start", "departure_time", "end", "line"]:
        if column not in df.columns:
            raise ValueError(f"Timetable: column '{column}' is missing")

    df["dep"] = df["departure_time"].apply(time_to_minutes)
    if df["dep"].isna().any():
        raise ValueError("Timetable: some departure times are not valid")
    df["line"] = df["line"].astype(int)
    return df[["start", "end", "dep", "line"]].sort_values("dep").reset_index(drop=True)


def read_plan(file, distances):
    """Reads the bus plan and checks the data.
    Returns (plan, issues). Rows with an error are reported and left out."""
    df = pd.read_excel(file, dtype=object)
    df.columns = [str(c).strip().lower() for c in df.columns]

    # Are all columns there?
    missing = [c for c in PLAN_COLUMNS if c not in df.columns]
    if missing:
        return None, [make_issue("DQ-COL", "Missing column(s): " + ", ".join(missing) + ".")]
    if len(df) == 0:
        return None, [make_issue("DQ-EMPTY", "The bus plan contains no data (only a header).")]

    locations = set()
    lines = set()
    for (start, end, line) in distances:
        locations.add(start)
        locations.add(end)
        if line is not None:
            lines.add(line)

    issues = []
    good_rows = []
    for i, r in df.iterrows():
        excel_row = i + 2  # row 1 is the header in Excel
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
        if len(errors) == 0 and activity in ["service trip", "material trip"]:
            route_line = line if activity == "service trip" else None
            if find_route(distances, r["start location"], r["end location"], route_line) is None:
                errors.append(("DQ-ROUTE", f"No distance known from {r['start location']} to {r['end location']}."))

        if len(errors) > 0:
            for check, message in errors:
                issues.append(make_issue(check, message, bus=None if bus is None else int(bus), row=excel_row))
        else:
            good_rows.append({
                "row": excel_row, "bus": int(bus),
                "start_loc": r["start location"], "end_loc": r["end location"],
                "start": start, "end": end, "activity": activity,
                "line": None if line is None else int(line), "energy": energy})

    if len(good_rows) == 0:
        issues.append(make_issue("DQ-EMPTY", "No usable rows in the bus plan."))
        return None, issues

    plan = pd.DataFrame(good_rows)
    fix_line_column(plan)

    # Are the rows of every bus in time order?
    for bus in plan["bus"].unique():
        rows = plan[plan["bus"] == bus]
        times = list(zip(rows["start"], rows["end"]))
        if times != sorted(times):
            issues.append(make_issue("DQ-ORDER", "Rows are not in time order (the tool sorts them).", "warning", bus=bus))

    plan = plan.sort_values(["bus", "start", "end"]).reset_index(drop=True)
    return plan, issues


# ---------------------------------------------------------------------------
# 2. Battery simulation + feasibility checks
# ---------------------------------------------------------------------------
def charge_energy(soc_kwh, minutes, settings):
    """kWh that is charged in a number of minutes: 450 kW up to 90%, 60 kW above 90%."""
    room_until_90 = max(0, MAX_SOC * capacity(settings) - soc_kwh)
    fast_minutes = min(minutes, room_until_90 / FAST_CHARGE_KW * 60)
    slow_minutes = minutes - fast_minutes
    charged = fast_minutes * FAST_CHARGE_KW / 60 + slow_minutes * SLOW_CHARGE_KW / 60
    return min(charged, capacity(settings) - soc_kwh)   # battery cannot go above 100%


def activity_energy(activity, soc_kwh, distances, settings):
    """Energy use (kWh) of one activity according to our model. Negative = charging."""
    minutes = activity["end"] - activity["start"]
    if activity["activity"] == "idle":
        return IDLE_KW * minutes / 60
    if activity["activity"] == "charging":
        return -charge_energy(soc_kwh, minutes, settings)
    km = trip_km(activity, distances)
    return km * settings["kwh_per_km"]


def trip_km(activity, distances):
    """Distance of a trip in km (0 for idle and charging)."""
    if activity["activity"] == "service trip":
        return find_route(distances, activity["start_loc"], activity["end_loc"], activity["line"])[2]
    if activity["activity"] == "material trip":
        return find_route(distances, activity["start_loc"], activity["end_loc"])[2]
    return 0


def check_plan(plan, timetable, distances, settings=None):
    """Follows the battery of every bus through the day and checks all rules.
    Returns (timeline, issues). The timeline has the SOC after every activity."""
    if settings is None:
        settings = DEFAULT_SETTINGS
    full = capacity(settings)
    issues = []
    timeline = []

    timetable_trips = set()
    for _, t in timetable.iterrows():
        timetable_trips.add((t["start"], t["dep"], t["end"], t["line"]))
    driven = {}  # timetable trip -> list of (bus, row) that drive it

    for bus in sorted(plan["bus"].unique()):
        soc = settings["start_soc"] * full
        previous = None
        first_too_low = None

        for _, a in plan[plan["bus"] == bus].iterrows():
            row = a["row"]
            time = minutes_to_time(a["start"])
            minutes = a["end"] - a["start"]

            # FC4 overlap and FC5 location, compared with the previous activity
            if previous is not None:
                if a["start"] < previous["end"]:
                    issues.append(make_issue("FC4", f"Overlaps with the previous activity (ends {minutes_to_time(previous['end'])}).",
                                             bus=bus, row=row, time=time))
                if a["start_loc"] != previous["end_loc"]:
                    issues.append(make_issue("FC5", f"Bus is at {previous['end_loc']} but this activity starts at {a['start_loc']}.",
                                             bus=bus, row=row, time=time))
                gap = a["start"] - previous["end"]
                if gap > 0:
                    soc = soc - IDLE_KW * gap / 60   # waiting time without a row = standing still

            # Energy: our model, compared with the value in the file
            energy = activity_energy(a, soc, distances, settings)
            if abs(a["energy"] - energy) > 0.05:
                issues.append(make_issue("DQ-ENERGY", f"Energy in file {a['energy']:.2f} kWh, model {energy:.2f} kWh.",
                                         "warning", bus=bus, row=row, time=time))
            if settings["use_file_energy"]:
                energy = a["energy"]
            soc = soc - energy
            soc_percent = round(soc / full * 100, 2)

            # FC2, FC3, FC6: rules for charging
            if a["activity"] == "charging":
                if minutes < MIN_CHARGE_TIME:
                    issues.append(make_issue("FC2", f"Charging session of {minutes} min (minimum {MIN_CHARGE_TIME}).",
                                             bus=bus, row=row, time=time))
                if soc_percent > MAX_SOC * 100:
                    issues.append(make_issue("FC3", f"Battery charged to {soc_percent:.2f}% (maximum {MAX_SOC * 100:.0f}%).",
                                             bus=bus, row=row, time=time))
                if a["start_loc"] != GARAGE:
                    issues.append(make_issue("FC6", f"Charging at {a['start_loc']}; only possible at {GARAGE}.",
                                             bus=bus, row=row, time=time))

            # FC4: service trips must be in the timetable, material trips have no line
            if a["activity"] == "service trip":
                trip = (a["start_loc"], a["start"], a["end_loc"], a["line"])
                if trip in timetable_trips:
                    driven.setdefault(trip, []).append((bus, row))
                else:
                    issues.append(make_issue("FC4", f"Service trip line {a['line']} at {time} is not in the timetable.",
                                             bus=bus, row=row, time=time))
            if a["activity"] == "material trip" and a["line"] is not None:
                issues.append(make_issue("FC4", f"Material trip has line number {a['line']}.", bus=bus, row=row, time=time))

            # FC7: a trip cannot be faster than the shortest travel time
            if a["activity"] in ["service trip", "material trip"]:
                route_line = a["line"] if a["activity"] == "service trip" else None
                fastest = find_route(distances, a["start_loc"], a["end_loc"], route_line)[0]
                if minutes < fastest:
                    issues.append(make_issue("FC7", f"Trip takes {minutes} min, minimum travel time is {fastest} min.",
                                             bus=bus, row=row, time=time))

            # FC1: remember the first moment the battery is too low
            if soc_percent < MIN_SOC * 100 and first_too_low is None:
                first_too_low = (row, time, soc_percent)

            timeline.append({"bus": bus, "row": row, "activity": a["activity"],
                             "start_loc": a["start_loc"], "end_loc": a["end_loc"],
                             "start": a["start"], "end": a["end"], "duration": minutes,
                             "line": a["line"], "km": trip_km(a, distances),
                             "energy": energy, "soc": soc_percent})
            previous = a

        if first_too_low is not None:
            row, time, soc_percent = first_too_low
            issues.append(make_issue("FC1", f"SOC drops to {soc_percent:.2f}% (minimum {MIN_SOC * 100:.0f}%).",
                                     bus=bus, row=row, time=time))

    # FC8: every timetable trip must be driven exactly once
    for _, t in timetable.iterrows():
        trip = (t["start"], t["dep"], t["end"], t["line"])
        buses = driven.get(trip, [])
        if len(buses) == 0:
            issues.append(make_issue("FC8", f"Timetable trip {t['start']}->{t['end']} line {t['line']} at "
                                            f"{minutes_to_time(t['dep'])} is not driven by any bus.",
                                     time=minutes_to_time(t["dep"])))
        elif len(buses) > 1:
            issues.append(make_issue("FC8", f"Timetable trip {t['start']}->{t['end']} line {t['line']} at "
                                            f"{minutes_to_time(t['dep'])} is driven {len(buses)} times.",
                                     bus=buses[1][0], row=buses[1][1], time=minutes_to_time(t["dep"])))

    return pd.DataFrame(timeline), issues


def is_feasible(issues):
    """A plan is feasible when there are no errors (warnings are allowed)."""
    for issue in issues:
        if issue["severity"] == "error":
            return False
    return True


# ---------------------------------------------------------------------------
# 3. KPIs
# ---------------------------------------------------------------------------
def kpis(timeline):
    material = timeline[timeline["activity"] == "material trip"]
    charging = timeline[timeline["activity"] == "charging"]
    idle = timeline[timeline["activity"] == "idle"]
    service = timeline[timeline["activity"] == "service trip"]
    return {
        "Buses used": int(timeline["bus"].nunique()),
        "Service trips": len(service),
        "Material trips": len(material),
        "Material trip time (min)": int(material["duration"].sum()),
        "Material trip distance (km)": round(float(material["km"].sum()), 2),
        "Charging sessions": len(charging),
        "Total charging time (min)": int(charging["duration"].sum()),
        "Idle time (min)": int(idle["duration"].sum()),
        "Energy used (kWh)": round(float(timeline[timeline["energy"] > 0]["energy"].sum()), 1),
        "Lowest SOC (%)": float(timeline["soc"].min()),
    }


# ---------------------------------------------------------------------------
# 4. Improved bus plan (greedy method)
# ---------------------------------------------------------------------------
# Idea: go through the timetable trips from early to late. Give every trip to
# the bus that can do it with the least waiting. A bus only goes charging when
# it cannot do the next trip otherwise. If no bus can do the trip, a new bus
# leaves the garage.

def new_activity(start_loc, end_loc, start, end, activity, energy, line=None):
    return {"start_loc": start_loc, "end_loc": end_loc, "start": start, "end": end,
            "activity": activity, "line": line, "energy": energy}


def material_trip(start, end, distances, settings):
    """(minutes, kWh) of a material trip. Zero if the bus is already there."""
    if start == end:
        return 0, 0
    min_time, max_time, km = find_route(distances, start, end)
    return max_time, km * settings["kwh_per_km"]


def enough_battery(soc_after_trip, trip, distances, settings):
    """After the trip the bus must still be able to reach the garage with at least 10%."""
    _, energy_to_garage = material_trip(trip["end"], GARAGE, distances, settings)
    return soc_after_trip - energy_to_garage >= MIN_SOC * capacity(settings)


def option_direct(bus, trip, distances, settings):
    """Option 1: the bus (waits and) drives to the start of the trip and does the trip."""
    drive_time, drive_energy = material_trip(bus["location"], trip["start"], distances, settings)
    wait = trip["dep"] - drive_time - bus["free_at"]
    if wait < 0:
        return None   # the bus is too late

    activities = []
    if wait > 0:
        activities.append(new_activity(bus["location"], bus["location"], bus["free_at"], bus["free_at"] + wait,
                                       "idle", IDLE_KW * wait / 60))
    if drive_time > 0:
        activities.append(new_activity(bus["location"], trip["start"], trip["dep"] - drive_time, trip["dep"],
                                       "material trip", drive_energy))
    activities.append(new_activity(trip["start"], trip["end"], trip["dep"], trip["arr"],
                                   "service trip", trip["energy"], trip["line"]))

    soc = bus["soc"] - sum(a["energy"] for a in activities)
    if not enough_battery(soc, trip, distances, settings):
        return None
    return {"activities": activities, "soc": soc, "cost": wait + 2 * drive_time}


def option_charge(bus, trip, distances, settings):
    """Option 2: the bus goes to the garage, charges (15+ min, max 90%) and then does the trip."""
    to_garage_time, to_garage_energy = material_trip(bus["location"], GARAGE, distances, settings)
    from_garage_time, from_garage_energy = material_trip(GARAGE, trip["start"], distances, settings)

    arrive = bus["free_at"] + to_garage_time
    soc = bus["soc"] - to_garage_energy
    time_available = trip["dep"] - from_garage_time - arrive
    minutes_until_90 = math.floor((MAX_SOC * capacity(settings) - soc) / FAST_CHARGE_KW * 60)
    charge_minutes = min(time_available, minutes_until_90)
    if charge_minutes < MIN_CHARGE_TIME:
        return None   # not enough time, or the battery is already too full

    activities = []
    if to_garage_time > 0:
        activities.append(new_activity(bus["location"], GARAGE, bus["free_at"], arrive,
                                       "material trip", to_garage_energy))
    activities.append(new_activity(GARAGE, GARAGE, arrive, arrive + charge_minutes,
                                   "charging", -charge_minutes * FAST_CHARGE_KW / 60))
    wait = time_available - charge_minutes
    if wait > 0:
        activities.append(new_activity(GARAGE, GARAGE, arrive + charge_minutes, arrive + time_available,
                                       "idle", IDLE_KW * wait / 60))
    if from_garage_time > 0:
        activities.append(new_activity(GARAGE, trip["start"], trip["dep"] - from_garage_time, trip["dep"],
                                       "material trip", from_garage_energy))
    activities.append(new_activity(trip["start"], trip["end"], trip["dep"], trip["arr"],
                                   "service trip", trip["energy"], trip["line"]))

    soc = bus["soc"] - sum(a["energy"] for a in activities)
    if not enough_battery(soc, trip, distances, settings):
        return None
    return {"activities": activities, "soc": soc, "cost": wait + 2 * (to_garage_time + from_garage_time) + 10}


def improve(timetable, distances, settings=None):
    """Makes an improved bus plan from the timetable."""
    if settings is None:
        settings = DEFAULT_SETTINGS
    buses = []

    for _, t in timetable.iterrows():
        min_time, max_time, km = find_route(distances, t["start"], t["end"], t["line"])
        trip = {"start": t["start"], "end": t["end"], "dep": t["dep"], "arr": t["dep"] + max_time,
                "line": t["line"], "energy": km * settings["kwh_per_km"]}

        # Find the best existing bus for this trip
        best_bus = None
        best_option = None
        for bus in buses:
            option = option_direct(bus, trip, distances, settings)
            if option is None:
                option = option_charge(bus, trip, distances, settings)
            if option is not None and (best_option is None or option["cost"] < best_option["cost"]):
                best_bus = bus
                best_option = option

        # No bus can do it: a new bus leaves the garage
        if best_bus is None:
            drive_time, _ = material_trip(GARAGE, trip["start"], distances, settings)
            best_bus = {"number": len(buses) + 1, "location": GARAGE, "free_at": trip["dep"] - drive_time,
                        "soc": settings["start_soc"] * capacity(settings), "activities": []}
            best_option = option_direct(best_bus, trip, distances, settings)
            buses.append(best_bus)

        best_bus["activities"] += best_option["activities"]
        best_bus["location"] = trip["end"]
        best_bus["free_at"] = trip["arr"]
        best_bus["soc"] = best_option["soc"]

    # End of the day: every bus drives back to the garage
    rows = []
    for bus in buses:
        drive_time, drive_energy = material_trip(bus["location"], GARAGE, distances, settings)
        if drive_time > 0:
            bus["activities"].append(new_activity(bus["location"], GARAGE, bus["free_at"], bus["free_at"] + drive_time,
                                                  "material trip", drive_energy))
        for a in bus["activities"]:
            a["bus"] = bus["number"]
            rows.append(a)

    plan = pd.DataFrame(rows)
    plan["row"] = range(2, len(plan) + 2)
    fix_line_column(plan)
    return plan


# ---------------------------------------------------------------------------
# 5. Save a plan in the client's Excel format
# ---------------------------------------------------------------------------
def to_excel(plan):
    table = pd.DataFrame({
        "start location": plan["start_loc"],
        "end location": plan["end_loc"],
        "start time": [minutes_to_time(m) + ":00" for m in plan["start"]],
        "end time": [minutes_to_time(m) + ":00" for m in plan["end"]],
        "activity": plan["activity"],
        "line": plan["line"],
        "energy consumption": plan["energy"].astype(float).round(4),
        "bus": plan["bus"],
    })
    buffer = io.BytesIO()
    table.to_excel(buffer, index=False, sheet_name="Sheet1")
    return buffer.getvalue()
