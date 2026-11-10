from flask import Flask, request, jsonify, send_file
from pathlib import Path
from datetime import datetime, timezone, timedelta
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from io import BytesIO
from threading import RLock
import os
import uuid
import json
import hashlib

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("PVUNL_DATA_DIR", str(APP_DIR)))
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_FILE = DATA_DIR / "PVUNL_Truck_Movement.xlsx"
ADMIN_ACCESS_CODE = os.environ.get("PVUNL_ADMIN_CODE", "PVUNL2026")
TZ = timezone(timedelta(hours=5, minutes=30), name="IST")
LOCK = RLock()
OPERATOR_CONFIG_FILE = DATA_DIR / "operator_access.json"

def password_hash(value):
    return hashlib.sha256(value.encode("utf-8")).hexdigest()

def get_operator_password_hash():
    if OPERATOR_CONFIG_FILE.exists():
        try:
            saved = json.loads(OPERATOR_CONFIG_FILE.read_text(encoding="utf-8"))
            if saved.get("password_hash"):
                return saved["password_hash"]
        except Exception:
            app.logger.exception("Could not read operator password configuration")
    initial_password = os.environ.get("PVUNL_OPERATOR_CODE", "PVUNL2028")
    return password_hash(initial_password)

def save_operator_password(new_password):
    payload = {"password_hash": password_hash(new_password)}
    temp_file = OPERATOR_CONFIG_FILE.with_suffix(".tmp")
    temp_file.write_text(json.dumps(payload), encoding="utf-8")
    temp_file.replace(OPERATOR_CONFIG_FILE)

app = Flask(__name__, static_folder=None)

TRUCK_HEADERS = ["Truck_ID", "Truck_Number", "Agency", "Truck_Type", "Driver_Name", "Driver_Mobile", "Registered_At"]
MOVEMENT_HEADERS = ["Movement_ID", "Truck_ID", "Truck_Number", "Agency",
                    "Entry_Date", "Entry_Time", "Entry_Gate",
                    "Exit_Date", "Exit_Time", "Exit_Gate", "TAT_Minutes", "Status"]

def now_local():
    return datetime.now(TZ)

def ensure_database():
    if DB_FILE.exists():
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Trucks"
    ws.append(TRUCK_HEADERS)
    wm = wb.create_sheet("Movements")
    wm.append(MOVEMENT_HEADERS)
    for sheet in (ws, wm):
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="173B60")
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
    wb.save(DB_FILE)

def load_db():
    ensure_database()
    return load_workbook(DB_FILE)

def save_db(wb):
    wb.save(DB_FILE)

def rows_as_dict(ws):
    headers = [c.value for c in ws[1]]
    return [dict(zip(headers, row)) for row in ws.iter_rows(min_row=2, values_only=True)
            if any(v is not None for v in row)]

def get_role():
    supplied = request.headers.get("X-Access-Code", "")
    if supplied and supplied == ADMIN_ACCESS_CODE:
        return "admin"
    if supplied and password_hash(supplied) == get_operator_password_hash():
        return "operator"
    return None

def authorized(required_role=None):
    role = get_role()
    if role is None:
        return False
    if required_role == "admin" and role != "admin":
        return False
    return True

def auth_error():
    return jsonify({"error": "Invalid access code or insufficient permissions"}), 401

@app.post("/api/login")
def login_api():
    data = request.get_json(silent=True) or {}
    supplied = str(data.get("access_code", "")).strip()
    if supplied == ADMIN_ACCESS_CODE:
        return jsonify({"ok": True, "role": "admin"})
    if password_hash(supplied) == get_operator_password_hash():
        return jsonify({"ok": True, "role": "operator"})
    return jsonify({"error": "Invalid access code"}), 401

@app.get("/")
def home():
    return send_file(APP_DIR / "index.html")

@app.get("/manifest.webmanifest")
def pwa_manifest():
    return send_file(APP_DIR / "manifest.webmanifest", mimetype="application/manifest+json")


@app.get("/service-worker.js")
def pwa_service_worker():
    response = send_file(APP_DIR / "service-worker.js", mimetype="application/javascript")
    response.headers["Service-Worker-Allowed"] = "/"
    response.headers["Cache-Control"] = "no-cache"
    return response


@app.get("/icons/pvunl-app.svg")
def pwa_app_icon():
    return send_file(APP_DIR / "pvunl-app.svg", mimetype="image/svg+xml")


@app.get("/api/health")
def health():
    try:
        ensure_database()
        return jsonify({"ok": True, "database": str(DB_FILE.name)})
    except Exception as exc:
        app.logger.exception("Health check failed")
        return jsonify({"ok": False, "error": str(exc)}), 500


def parse_movement_datetime(date_value, time_value):
    """Parse a stored local scan date/time into a naive datetime."""
    if not date_value or not time_value:
        return None
    if isinstance(date_value, datetime):
        date_part = date_value.strftime("%Y-%m-%d")
    else:
        date_part = str(date_value).strip()[:10]
    if isinstance(time_value, datetime):
        time_part = time_value.strftime("%H:%M:%S")
    else:
        time_part = str(time_value).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(f"{date_part} {time_part}", fmt)
        except (TypeError, ValueError):
            continue
    return None

@app.get("/api/dashboard")
def dashboard():
    if not authorized("admin"):
        return auth_error()
    with LOCK:
        wb = load_db()
        trucks = rows_as_dict(wb["Trucks"])
        movements = rows_as_dict(wb["Movements"])
    current = now_local()
    today = current.strftime("%Y-%m-%d")
    open_trips = [m for m in movements if str(m.get("Status", "")).upper() == "OUTSIDE"]
    out_today = [m for m in movements if str(m.get("Exit_Date") or "") == today]
    returned_today = [m for m in movements if str(m.get("Entry_Date") or "") == today and m.get("Entry_Time")]
    completed_today = [m for m in returned_today if m.get("TAT_Minutes") is not None]
    tat_values = [float(m["TAT_Minutes"]) for m in completed_today]
    avg_tat_today = round(sum(tat_values) / len(tat_values), 1) if tat_values else 0
    open_rows = []
    for m in open_trips:
        age = None
        try:
            started = datetime.strptime(f"{m.get('Exit_Date')} {m.get('Exit_Time')}", "%Y-%m-%d %H:%M:%S").replace(tzinfo=TZ)
            age = max(0, round((current - started).total_seconds() / 60, 1))
        except Exception:
            pass
        open_rows.append({"Truck_Number": m.get("Truck_Number") or "", "Agency": m.get("Agency") or "",
                          "OUT_Date": m.get("Exit_Date") or "", "OUT_Time": m.get("Exit_Time") or "",
                          "OUT_Gate": m.get("Exit_Gate") or "", "Age_Minutes": age})
    open_rows.sort(key=lambda x: x.get("Age_Minutes") if x.get("Age_Minutes") is not None else -1, reverse=True)
    recent = sorted(movements, key=lambda x: f"{x.get('Exit_Date') or x.get('Entry_Date') or ''} {x.get('Exit_Time') or x.get('Entry_Time') or ''}", reverse=True)[:30]
    recent_movements = [{"Movement_ID": m.get("Movement_ID") or "", "Truck_Number": m.get("Truck_Number") or "",
        "Agency": m.get("Agency") or "", "Entry_Date": m.get("Entry_Date") or "", "Entry_Time": m.get("Entry_Time") or "",
        "Entry_Gate": m.get("Entry_Gate") or "", "Exit_Date": m.get("Exit_Date") or "", "Exit_Time": m.get("Exit_Time") or "",
        "Exit_Gate": m.get("Exit_Gate") or "", "TAT_Minutes": m.get("TAT_Minutes"), "Status": m.get("Status") or ""} for m in recent]
    daily = []
    for offset in range(6, -1, -1):
        day = (current - timedelta(days=offset)).strftime("%Y-%m-%d")
        daily.append({"date": day, "out": sum(1 for m in movements if str(m.get("Exit_Date") or "") == day),
                      "in": sum(1 for m in movements if str(m.get("Entry_Date") or "") == day and m.get("Entry_Time"))})

    # Last 3 calendar days: daily and shift-wise inside/outside TAT and completed ash-mound trips.
    # Shift C is attributed to the date on which the shift started (22:00-06:00).
    report_days = [(current - timedelta(days=offset)).strftime("%Y-%m-%d") for offset in (2, 1, 0)]
    shift_names = ("A", "B", "C")

    def shift_for(stamp):
        hour = stamp.hour
        if 6 <= hour < 14: return "A"
        if 14 <= hour < 22: return "B"
        return "C"

    def shift_date_for(stamp):
        # 00:00-05:59 belongs to the previous day's night shift.
        if stamp.hour < 6:
            return (stamp.date() - timedelta(days=1)).strftime("%Y-%m-%d")
        return stamp.strftime("%Y-%m-%d")

    tat_buckets = {}
    ash_daily = {day: 0 for day in report_days}
    ash_shift = {day: {shift: 0 for shift in shift_names} for day in report_days}
    for day in report_days:
        tat_buckets[day] = {
            shift: {"inside_sum": 0.0, "inside_count": 0, "outside_sum": 0.0, "outside_count": 0}
            for shift in shift_names
        }

    by_truck = {}
    for movement in movements:
        truck_key = str(movement.get("Truck_ID") or movement.get("Truck_Number") or "")
        if truck_key:
            by_truck.setdefault(truck_key, []).append(movement)

    for truck_movements in by_truck.values():
        truck_movements.sort(key=lambda m: (
            parse_movement_datetime(m.get("Exit_Date"), m.get("Exit_Time")) or
            parse_movement_datetime(m.get("Entry_Date"), m.get("Entry_Time")) or
            datetime.min
        ))
        previous_return = None
        for movement in truck_movements:
            out_dt = parse_movement_datetime(movement.get("Exit_Date"), movement.get("Exit_Time"))
            in_dt = parse_movement_datetime(movement.get("Entry_Date"), movement.get("Entry_Time"))

            # One OUT followed by its subsequent IN is one completed trip to Main Ash Mound.
            if out_dt and in_dt and in_dt >= out_dt:
                trip_day = shift_date_for(out_dt)
                trip_shift = shift_for(out_dt)
                if trip_day in ash_daily:
                    ash_daily[trip_day] += 1
                    ash_shift[trip_day][trip_shift] += 1

                # Outside TAT is attributed to the return IN day and shift.
                outside_day = shift_date_for(in_dt)
                outside_shift = shift_for(in_dt)
                if outside_day in tat_buckets:
                    bucket = tat_buckets[outside_day][outside_shift]
                    bucket["outside_sum"] += (in_dt - out_dt).total_seconds() / 60
                    bucket["outside_count"] += 1

            # Inside TAT is attributed to the next OUT day and shift.
            if previous_return and out_dt and out_dt >= previous_return:
                inside_day = shift_date_for(out_dt)
                inside_shift = shift_for(out_dt)
                if inside_day in tat_buckets:
                    bucket = tat_buckets[inside_day][inside_shift]
                    bucket["inside_sum"] += (out_dt - previous_return).total_seconds() / 60
                    bucket["inside_count"] += 1

            if in_dt:
                previous_return = in_dt

    tat_3day = []
    tat_shift_3day = []
    ash_mound_3day = []
    for day in report_days:
        day_buckets = tat_buckets[day]
        inside_count = sum(v["inside_count"] for v in day_buckets.values())
        outside_count = sum(v["outside_count"] for v in day_buckets.values())
        inside_sum = sum(v["inside_sum"] for v in day_buckets.values())
        outside_sum = sum(v["outside_sum"] for v in day_buckets.values())
        tat_3day.append({
            "date": day,
            "avg_inside_minutes": round(inside_sum / inside_count, 1) if inside_count else None,
            "inside_count": inside_count,
            "avg_outside_minutes": round(outside_sum / outside_count, 1) if outside_count else None,
            "outside_count": outside_count
        })
        ash_mound_3day.append({"date": day, "trips": ash_daily[day]})
        for shift in shift_names:
            bucket = day_buckets[shift]
            tat_shift_3day.append({
                "date": day, "shift": shift,
                "avg_inside_minutes": round(bucket["inside_sum"] / bucket["inside_count"], 1) if bucket["inside_count"] else None,
                "inside_count": bucket["inside_count"],
                "avg_outside_minutes": round(bucket["outside_sum"] / bucket["outside_count"], 1) if bucket["outside_count"] else None,
                "outside_count": bucket["outside_count"]
            })
    ash_mound_shift_3day = [
        {"date": day, "shift": shift, "trips": ash_shift[day][shift]}
        for day in report_days for shift in shift_names
    ]

    # Truck-wise lifetime performance, calculated from the complete movement history.
    truck_perf = {}
    for truck in trucks:
        key = str(truck.get("Truck_ID") or "")
        if key:
            truck_perf[key] = {"truck_number": truck.get("Truck_Number") or "", "agency": truck.get("Agency") or "", "completed_trips": 0, "inside_sum": 0.0, "inside_count": 0, "outside_sum": 0.0, "outside_count": 0, "longest_outside": 0.0, "last_movement": None, "current_status": "INSIDE PLANT", "open_since": None}
    for movement in movements:
        key = str(movement.get("Truck_ID") or "")
        if key not in truck_perf: continue
        item = truck_perf[key]
        out_dt = parse_movement_datetime(movement.get("Exit_Date"), movement.get("Exit_Time"))
        in_dt = parse_movement_datetime(movement.get("Entry_Date"), movement.get("Entry_Time"))
        if out_dt and in_dt and in_dt >= out_dt:
            outside_minutes = (in_dt - out_dt).total_seconds() / 60
            item["outside_sum"] += outside_minutes; item["outside_count"] += 1; item["completed_trips"] += 1
            item["longest_outside"] = max(item["longest_outside"], outside_minutes)
        for stamp in (out_dt, in_dt):
            if stamp and (item["last_movement"] is None or stamp > item["last_movement"]): item["last_movement"] = stamp
    for key, item in truck_perf.items():
        truck_movements = [m for m in movements if str(m.get("Truck_ID") or "") == key]
        truck_movements.sort(key=lambda m: parse_movement_datetime(m.get("Exit_Date"), m.get("Exit_Time")) or parse_movement_datetime(m.get("Entry_Date"), m.get("Entry_Time")) or datetime.min)
        previous_return = None; latest_open = None
        for movement in truck_movements:
            out_dt = parse_movement_datetime(movement.get("Exit_Date"), movement.get("Exit_Time"))
            in_dt = parse_movement_datetime(movement.get("Entry_Date"), movement.get("Entry_Time"))
            if previous_return and out_dt and out_dt >= previous_return:
                item["inside_sum"] += (out_dt - previous_return).total_seconds() / 60; item["inside_count"] += 1
            if in_dt: previous_return = in_dt
            if str(movement.get("Status") or "").upper() == "OUTSIDE" and out_dt: latest_open = out_dt
        if latest_open:
            item["current_status"] = "OUTSIDE PLANT"; item["open_since"] = latest_open.strftime("%Y-%m-%d %H:%M:%S")
        item["avg_inside_minutes"] = round(item["inside_sum"] / item["inside_count"], 1) if item["inside_count"] else None
        item["avg_outside_minutes"] = round(item["outside_sum"] / item["outside_count"], 1) if item["outside_count"] else None
        item["longest_outside_minutes"] = round(item["longest_outside"], 1) if item["outside_count"] else None
        item["last_movement"] = item["last_movement"].strftime("%Y-%m-%d %H:%M:%S") if item["last_movement"] else ""
        for field in ("inside_sum", "outside_sum", "longest_outside"): item.pop(field, None)
    truck_performance = sorted(truck_perf.values(), key=lambda x: x["truck_number"])

    # Per-truck lifetime metrics, based on completed IN/OUT cycles.
    truck_perf = {}
    for truck in trucks:
        key = str(truck.get("Truck_ID") or "")
        if key:
            truck_perf[key] = {"truck_number": truck.get("Truck_Number") or "", "agency": truck.get("Agency") or "", "completed_trips": 0, "inside_sum": 0.0, "inside_count": 0, "outside_sum": 0.0, "outside_count": 0, "longest_outside": 0.0, "current_status": "INSIDE PLANT", "open_since": ""}
    for key, item in truck_perf.items():
        truck_movements = [m for m in movements if str(m.get("Truck_ID") or "") == key]
        truck_movements.sort(key=lambda m: parse_movement_datetime(m.get("Exit_Date"), m.get("Exit_Time")) or parse_movement_datetime(m.get("Entry_Date"), m.get("Entry_Time")) or datetime.min)
        previous_return = None
        for movement in truck_movements:
            out_dt = parse_movement_datetime(movement.get("Exit_Date"), movement.get("Exit_Time"))
            in_dt = parse_movement_datetime(movement.get("Entry_Date"), movement.get("Entry_Time"))
            if out_dt and in_dt and in_dt >= out_dt:
                mins = (in_dt - out_dt).total_seconds() / 60
                item["outside_sum"] += mins
                item["outside_count"] += 1
                item["completed_trips"] += 1
                item["longest_outside"] = max(item["longest_outside"], mins)
            if previous_return and out_dt and out_dt >= previous_return:
                item["inside_sum"] += (out_dt - previous_return).total_seconds() / 60
                item["inside_count"] += 1
            if in_dt:
                previous_return = in_dt
            if str(movement.get("Status") or "").upper() == "OUTSIDE" and out_dt:
                item["current_status"] = "OUTSIDE PLANT"
                item["open_since"] = out_dt.strftime("%Y-%m-%d %H:%M:%S")
            elif in_dt:
                item["current_status"] = "INSIDE PLANT"
                item["open_since"] = ""
        item["avg_inside_minutes"] = round(item["inside_sum"] / item["inside_count"], 1) if item["inside_count"] else None
        item["avg_outside_minutes"] = round(item["outside_sum"] / item["outside_count"], 1) if item["outside_count"] else None
        item["longest_outside_minutes"] = round(item["longest_outside"], 1) if item["outside_count"] else None
        for field in ("inside_sum", "inside_count", "outside_sum", "outside_count", "longest_outside"):
            if field in ("inside_count", "outside_count"):
                continue
            item.pop(field, None)
    truck_performance = sorted(truck_perf.values(), key=lambda x: x["truck_number"])

    gate_summary = []
    for gate in ("Gate 3", "Gate 4"):
        gate_summary.append({"gate": gate,
            "out": sum(1 for m in out_today if str(m.get("Exit_Gate") or "") == gate),
            "in": sum(1 for m in returned_today if str(m.get("Entry_Gate") or "") == gate)})
    agency_map = {}
    for m in movements:
        agency = str(m.get("Agency") or "Not specified")
        agency_map[agency] = agency_map.get(agency, 0) + 1
    agency_summary = [{"agency": k, "trips": v} for k, v in sorted(agency_map.items(), key=lambda x: x[1], reverse=True)[:8]]
    return jsonify({"registered_trucks": len(trucks), "open_trips": len(open_trips),
        "inside_plant": max(0, len(trucks) - len(open_trips)),
        "outside_fleet_pct": round(len(open_trips) / len(trucks) * 100, 1) if trucks else 0,
        "out_today": len(out_today), "returned_today": len(returned_today), "avg_tat_today": avg_tat_today,
        "completed_today": len(completed_today), "open_truck_rows": open_rows, "daily_trend": daily,
        "gate_summary": gate_summary, "agency_summary": agency_summary, "tat_trend": tat_3day, "tat_shift_trend": tat_shift_3day,
        "ash_mound_daily": ash_mound_3day, "ash_mound_shift": ash_mound_shift_3day,
        "recent_movements": recent_movements, "truck_performance": truck_performance})


@app.get("/api/reports")
def reports_api():
    if not authorized("admin"):
        return auth_error()
    today = now_local().strftime("%Y-%m-%d")
    default_from = (now_local() - timedelta(days=6)).strftime("%Y-%m-%d")
    date_from = str(request.args.get("from", default_from)).strip()
    date_to = str(request.args.get("to", today)).strip()
    try:
        start = datetime.strptime(date_from, "%Y-%m-%d").date()
        end = datetime.strptime(date_to, "%Y-%m-%d").date()
        if end < start or (end - start).days > 366:
            raise ValueError()
    except ValueError:
        return jsonify({"error": "Select a valid date range of up to 367 days."}), 400
    with LOCK:
        wb = load_db()
        movements = rows_as_dict(wb["Movements"])
    selected = []
    hours = [{"hour": hour, "out": 0, "in": 0} for hour in range(24)]
    gates = {gate: {"gate": gate, "out": 0, "in": 0} for gate in ("Gate 3", "Gate 4")}
    for movement in movements:
        out_date = str(movement.get("Exit_Date") or "")[:10]
        in_date = str(movement.get("Entry_Date") or "")[:10]
        out_match = bool(out_date and date_from <= out_date <= date_to)
        in_match = bool(in_date and date_from <= in_date <= date_to)
        if out_match:
            try:
                hour = int(str(movement.get("Exit_Time") or "00:00").split(":")[0])
                if 0 <= hour < 24: hours[hour]["out"] += 1
            except (TypeError, ValueError): pass
            gate = str(movement.get("Exit_Gate") or "")
            if gate in gates: gates[gate]["out"] += 1
        if in_match and movement.get("Entry_Time"):
            try:
                hour = int(str(movement.get("Entry_Time") or "00:00").split(":")[0])
                if 0 <= hour < 24: hours[hour]["in"] += 1
            except (TypeError, ValueError): pass
            gate = str(movement.get("Entry_Gate") or "")
            if gate in gates: gates[gate]["in"] += 1
        if out_match or in_match:
            selected.append({
                "Movement_ID": str(movement.get("Movement_ID") or ""),
                "Truck_Number": str(movement.get("Truck_Number") or ""),
                "Agency": str(movement.get("Agency") or ""),
                "Exit_Date": out_date, "Exit_Time": str(movement.get("Exit_Time") or ""),
                "Exit_Gate": str(movement.get("Exit_Gate") or ""),
                "Entry_Date": in_date, "Entry_Time": str(movement.get("Entry_Time") or ""),
                "Entry_Gate": str(movement.get("Entry_Gate") or ""),
                "TAT_Minutes": movement.get("TAT_Minutes"),
                "Status": str(movement.get("Status") or "")
            })
    selected.sort(key=lambda row: (row["Exit_Date"] or row["Entry_Date"], row["Exit_Time"] or row["Entry_Time"]), reverse=True)
    completed = [row for row in selected if row["TAT_Minutes"] not in (None, "")]
    tat_values = [float(row["TAT_Minutes"]) for row in completed]
    return jsonify({
        "date_from": date_from, "date_to": date_to, "total_records": len(selected),
        "completed_trips": len(completed),
        "open_trips": sum(1 for row in selected if row["Status"].upper() == "OUTSIDE"),
        "avg_tat_minutes": round(sum(tat_values) / len(tat_values), 1) if tat_values else None,
        "max_tat_minutes": round(max(tat_values), 1) if tat_values else None,
        "gate_summary": list(gates.values()), "hourly_summary": hours, "rows": selected
    })


@app.route("/api/trucks", methods=["GET", "POST"])
def trucks_api():
    if not authorized("admin"):
        return auth_error()
    with LOCK:
        wb = load_db()
        ws = wb["Trucks"]
        if request.method == "GET":
            trucks = rows_as_dict(ws)
            return jsonify({"trucks": trucks})
        data = request.get_json(silent=True) or {}
        truck_number = str(data.get("truck_number", "")).strip().upper()
        if not truck_number:
            return jsonify({"error": "Truck registration number is required"}), 400
        existing = rows_as_dict(ws)
        if any(str(t.get("Truck_Number", "")).upper() == truck_number for t in existing):
            return jsonify({"error": "This truck number is already registered"}), 409
        truck_id = "PVUNL-" + uuid.uuid4().hex[:10].upper()
        stamp = now_local().strftime("%Y-%m-%d %H:%M:%S")
        ws.append([truck_id, truck_number, str(data.get("agency", "")).strip(),
                   str(data.get("truck_type", "")).strip(), str(data.get("driver_name", "")).strip(),
                   str(data.get("driver_mobile", "")).strip(), stamp])
        save_db(wb)
    return jsonify({"message": "Truck registered", "truck_id": truck_id})

@app.get("/api/truck-lookup/<truck_id>")
def truck_lookup_api(truck_id):
    """Return minimal truck details for an authenticated QR deep link."""
    if not authorized():
        return auth_error()
    with LOCK:
        wb = load_db()
        trucks = rows_as_dict(wb["Trucks"])
        truck = next((t for t in trucks if str(t.get("Truck_ID", "")) == truck_id), None)
    if not truck:
        return jsonify({"error": "Truck ID not found. Please contact the administrator."}), 404
    return jsonify({"truck_id": truck_id, "truck_number": truck.get("Truck_Number") or "",
                    "agency": truck.get("Agency") or ""})


@app.post("/api/scan")
def scan_api():
    if not authorized():
        return auth_error()
    data = request.get_json(silent=True) or {}
    truck_id = str(data.get("truck_id", "")).strip()
    action = str(data.get("action", "")).lower().strip()
    gate = str(data.get("gate", "")).strip()
    if gate not in ("Gate 3", "Gate 4"):
        return jsonify({"error": "Please select Gate 3 or Gate 4"}), 400
    if action not in ("entry", "exit"):
        return jsonify({"error": "Select a valid movement action (IN or OUT)"}), 400
    with LOCK:
        wb = load_db()
        trucks = rows_as_dict(wb["Trucks"])
        truck = next((t for t in trucks if str(t.get("Truck_ID", "")) == truck_id), None)
        if not truck:
            return jsonify({"error": "Truck ID not found. Register the truck first."}), 404
        ws = wb["Movements"]
        movements = rows_as_dict(ws)
        active_out = next((m for m in reversed(movements) if str(m.get("Truck_ID", "")) == truck_id and str(m.get("Status", "")).upper() == "OUTSIDE"), None)
        current = now_local()
        date_s, time_s = current.strftime("%Y-%m-%d"), current.strftime("%H:%M:%S")
        if action == "exit":
            if active_out:
                return jsonify({"error": "This truck is already recorded OUTSIDE. Record IN before another OUT."}), 409
            movement_id = "MOV-" + uuid.uuid4().hex[:12].upper()
            ws.append([movement_id, truck_id, truck.get("Truck_Number"), truck.get("Agency"), "", "", "", date_s, time_s, gate, None, "OUTSIDE"])
            save_db(wb)
            return jsonify({"message": f"OUT recorded for {truck.get('Truck_Number')} at {time_s}. Awaiting IN."})
        if not active_out:
            return jsonify({"error": "No open OUT record found for this truck. Record OUT before IN."}), 409
        try:
            out_dt = datetime.strptime(f"{active_out.get('Exit_Date')} {active_out.get('Exit_Time')}", "%Y-%m-%d %H:%M:%S")
            tat = max(0, round((current.replace(tzinfo=None) - out_dt).total_seconds() / 60, 1))
        except Exception:
            tat = None
        headers = [c.value for c in ws[1]]
        row_num = None
        for idx in range(2, ws.max_row + 1):
            if str(ws.cell(idx, headers.index("Movement_ID") + 1).value) == str(active_out.get("Movement_ID")):
                row_num = idx
                break
        if row_num is None:
            return jsonify({"error": "Movement record could not be located"}), 500
        updates = {"Entry_Date": date_s, "Entry_Time": time_s, "Entry_Gate": gate, "TAT_Minutes": tat, "Status": "INSIDE"}
        for key, value in updates.items():
            ws.cell(row_num, headers.index(key) + 1).value = value
        save_db(wb)
        return jsonify({"message": f"IN recorded for {truck.get('Truck_Number')} at {time_s}", "tat_minutes": tat})

@app.post("/api/operator-password")
def change_operator_password_api():
    if not authorized("admin"):
        return auth_error()
    data = request.get_json(silent=True) or {}
    new_password = str(data.get("new_password", "")).strip()
    if len(new_password) < 8:
        return jsonify({"error": "Operator password must contain at least 8 characters"}), 400
    if new_password == ADMIN_ACCESS_CODE:
        return jsonify({"error": "Operator password must be different from administrator password"}), 400
    with LOCK:
        save_operator_password(new_password)
    return jsonify({"message": "Gate Operator password updated successfully."})

@app.delete("/api/movements/<movement_id>")
def delete_movement_api(movement_id):
    """Delete one movement record. Administrator only."""
    if not authorized("admin"):
        return auth_error()
    with LOCK:
        wb = load_db()
        ws = wb["Movements"]
        headers = [c.value for c in ws[1]]
        id_col = headers.index("Movement_ID") + 1
        row_num = next((r for r in range(2, ws.max_row + 1)
                        if str(ws.cell(r, id_col).value or "") == movement_id), None)
        if row_num is None:
            return jsonify({"error": "Movement record not found"}), 404
        ws.delete_rows(row_num, 1)
        save_db(wb)
    return jsonify({"message": "Movement record deleted"})

@app.post("/api/reset-movements")
def reset_movements_api():
    """Delete all movement history but retain registered trucks. Administrator only."""
    if not authorized("admin"):
        return auth_error()
    with LOCK:
        wb = load_db()
        ws = wb["Movements"]
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
        save_db(wb)
    return jsonify({"message": "All movement records have been deleted. Truck master has been retained."})

@app.post("/api/reset-all")
def reset_all_api():
    """Delete all movement history and all registered trucks. Administrator only."""
    if not authorized("admin"):
        return auth_error()
    with LOCK:
        wb = load_db()
        for sheet_name, headers in (("Trucks", TRUCK_HEADERS), ("Movements", MOVEMENT_HEADERS)):
            ws = wb[sheet_name]
            if ws.max_row > 1:
                ws.delete_rows(2, ws.max_row - 1)
        save_db(wb)
    return jsonify({"message": "All truck registrations and movement records have been deleted."})

@app.get("/api/export")
def export_api():
    if not authorized("admin"):
        return auth_error()
    with LOCK:
        ensure_database()
        data = BytesIO(DB_FILE.read_bytes())
    data.seek(0)
    return send_file(data, as_attachment=True, download_name="PVUNL_Truck_Movement.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    ensure_database()
    print("PVUNL Truck Movement Management is starting...")
    print("Open http://127.0.0.1:5000")
    print("Administrator access code: PVUNL2026")
    print("Gate operator access code: PVUNL-OPS-2026 (change before operational use)")
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")), debug=False)
