from flask import Flask, request, jsonify, send_file
from pathlib import Path
from datetime import datetime, timezone, timedelta
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from io import BytesIO
from threading import RLock
import os
import uuid

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("PVUNL_DATA_DIR", str(APP_DIR)))
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_FILE = DATA_DIR / "PVUNL_Truck_Movement.xlsx"
ADMIN_ACCESS_CODE = os.environ.get("PVUNL_ADMIN_CODE", "PVUNL2026")
OPERATOR_ACCESS_CODE = os.environ.get("PVUNL_OPERATOR_CODE", "PVUNL-OPS-2026")
TZ = timezone(timedelta(hours=5, minutes=30), name="IST")
LOCK = RLock()

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
    if supplied and supplied == OPERATOR_ACCESS_CODE:
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
    if supplied == OPERATOR_ACCESS_CODE:
        return jsonify({"ok": True, "role": "operator"})
    return jsonify({"error": "Invalid access code"}), 401

@app.get("/")
def home():
    return send_file(APP_DIR / "index.html")

@app.get("/api/health")
def health():
    try:
        ensure_database()
        return jsonify({"ok": True, "database": str(DB_FILE.name)})
    except Exception as exc:
        app.logger.exception("Health check failed")
        return jsonify({"ok": False, "error": str(exc)}), 500

@app.get("/api/dashboard")
def dashboard():
    if not authorized("admin"):
        return auth_error()
    with LOCK:
        wb = load_db()
        trucks = rows_as_dict(wb["Trucks"])
        movements = rows_as_dict(wb["Movements"])
    today = now_local().strftime("%Y-%m-%d")
    active = [m for m in movements if str(m.get("Status", "")).upper() == "INSIDE"]
    exited_today = [m for m in movements if m.get("Exit_Date") == today]
    completed = [m for m in movements if m.get("TAT_Minutes") is not None]
    avg_tat = round(sum(float(m["TAT_Minutes"]) for m in completed) / len(completed), 1) if completed else 0
    inside_trucks = [{
        "Truck_Number": m.get("Truck_Number") or "",
        "Agency": m.get("Agency") or "",
        "Entry_Time": m.get("Entry_Time") or "",
        "Entry_Gate": m.get("Entry_Gate") or "",
        "Status": m.get("Status") or ""
    } for m in active]
    recent = sorted(movements, key=lambda x: str(x.get("Movement_ID", "")), reverse=True)[:30]
    recent_movements = [{
        "Movement_ID": m.get("Movement_ID") or "",
        "Truck_Number": m.get("Truck_Number") or "",
        "Agency": m.get("Agency") or "",
        "Entry_Date": m.get("Entry_Date") or "",
        "Entry_Time": m.get("Entry_Time") or "",
        "Entry_Gate": m.get("Entry_Gate") or "",
        "Exit_Date": m.get("Exit_Date") or "",
        "Exit_Time": m.get("Exit_Time") or "",
        "Exit_Gate": m.get("Exit_Gate") or "",
        "TAT_Minutes": m.get("TAT_Minutes"),
        "Status": m.get("Status") or ""
    } for m in recent]
    return jsonify({"registered_trucks": len(trucks), "inside": len(active),
                    "exited_today": len(exited_today), "avg_tat_minutes": avg_tat,
                    "inside_trucks": inside_trucks, "recent_movements": recent_movements})

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

@app.post("/api/scan")
def scan_api():
    if not authorized():
        return auth_error()
    data = request.get_json(silent=True) or {}
    truck_id = str(data.get("truck_id", "")).strip()
    action = str(data.get("action", "")).lower().strip()
    gate = str(data.get("gate", "Main Gate")).strip() or "Main Gate"
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
        active = next((m for m in reversed(movements)
                       if str(m.get("Truck_ID", "")) == truck_id and str(m.get("Status", "")).upper() == "INSIDE"), None)
        current = now_local()
        date_s, time_s = current.strftime("%Y-%m-%d"), current.strftime("%H:%M:%S")
        if action == "entry":
            if active:
                return jsonify({"error": "This truck is already recorded INSIDE the plant."}), 409
            movement_id = "MOV-" + uuid.uuid4().hex[:12].upper()
            ws.append([movement_id, truck_id, truck.get("Truck_Number"), truck.get("Agency"),
                       date_s, time_s, gate, "", "", "", None, "INSIDE"])
            save_db(wb)
            return jsonify({"message": f"IN recorded for {truck.get('Truck_Number')} at {time_s}"})
        if not active:
            return jsonify({"error": "No open entry found for this truck. Record IN before OUT."}), 409
        try:
            entry_dt = datetime.strptime(f"{active.get('Entry_Date')} {active.get('Entry_Time')}", "%Y-%m-%d %H:%M:%S")
            tat = max(0, round((current.replace(tzinfo=None) - entry_dt).total_seconds() / 60, 1))
        except Exception:
            tat = None
        # Update the active movement row in place.
        headers = [c.value for c in ws[1]]
        row_num = None
        for idx in range(2, ws.max_row + 1):
            if str(ws.cell(idx, headers.index("Movement_ID") + 1).value) == str(active.get("Movement_ID")):
                row_num = idx
                break
        if row_num is None:
            return jsonify({"error": "Movement record could not be located"}), 500
        updates = {"Exit_Date": date_s, "Exit_Time": time_s, "Exit_Gate": gate,
                   "TAT_Minutes": tat, "Status": "EXITED"}
        for key, value in updates.items():
            ws.cell(row_num, headers.index(key) + 1).value = value
        save_db(wb)
    return jsonify({"message": f"OUT recorded for {truck.get('Truck_Number')} at {time_s}", "tat_minutes": tat})


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
