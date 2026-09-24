import os
import shutil
import secrets
import logging
import re
import calendar
from datetime import datetime, timezone, date, time
from pathlib import Path
from typing import List, Optional
from fastapi import FastAPI, Request, UploadFile, File, Form, Depends, HTTPException, status, Query, Response
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from starlette.background import BackgroundTask
from sqlalchemy.orm import Session

from .config import *
from .schemas import *
from .database import engine, Base, get_db, SessionLocal
from .models import User, Employee, SwipeLog, ProcessingJob, AttendanceRecord
from .security import authenticate, create_access_token, decode_token, seed_default_users, is_authorized_for_manual_punchin
from .services.excel_service import read_swipes, update_attendance, read_attendance_employees

logger = logging.getLogger("attendance_backend")

# Create DB tables if not existing and seed default accounts
try:
    Base.metadata.create_all(bind=engine)
    # Ensure is_active column exists in employees table for existing databases
    try:
        with engine.begin() as conn:
            try:
                conn.execute(text("ALTER TABLE employees ADD COLUMN IF NOT EXISTS is_active BOOLEAN DEFAULT TRUE"))
            except Exception:
                try:
                    conn.execute(text("ALTER TABLE employees ADD COLUMN is_active BOOLEAN DEFAULT 1"))
                except Exception:
                    pass
            try:
                conn.execute(text("UPDATE employees SET is_active = TRUE WHERE is_active IS NULL"))
            except Exception:
                try:
                    conn.execute(text("UPDATE employees SET is_active = 1 WHERE is_active IS NULL"))
                except Exception:
                    pass
    except Exception as migr_err:
        logger.warning(f"Migration check: {migr_err}")
    db_init_session = SessionLocal()
    seed_default_users(db_init_session)
    db_init_session.close()
    logger.info("Database initialized and verified successfully.")
except Exception as err:
    logger.error(f"Error during DB initialization: {err}", exc_info=True)

app = FastAPI(title=APP_NAME)
app.add_middleware(
    CORSMiddleware,
    allow_origins=CORS_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"]
)

@app.middleware("http")
async def add_security_headers(request: Request, call_next):
    response = await call_next(request)
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["X-XSS-Protection"] = "1; mode=block"
    response.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    return response

bearer = HTTPBearer(auto_error=False)
import tempfile
default_temp = Path(tempfile.gettempdir()) / "attendance_punchin"
BASE = Path(os.getenv("TEMP_DIR", str(default_temp)))
BASE.mkdir(parents=True, exist_ok=True)

def auth(c: HTTPAuthorizationCredentials = Depends(bearer)):
    if not c:
        raise HTTPException(401, "Authentication required.")
    try:
        return decode_token(c.credentials)
    except Exception:
        raise HTTPException(401, "Invalid or expired session.")

async def save_upload(upload, dest):
    size = 0
    with open(dest, "wb") as f:
        while True:
            chunk = await upload.read(1024 * 1024)
            if not chunk:
                break
            size += len(chunk)
            if size > MAX_UPLOAD_MB * 1024 * 1024:
                raise HTTPException(413, f"File exceeds {MAX_UPLOAD_MB} MB.")
            f.write(chunk)

@app.get("/health", response_model=HealthResponse)
def health():
    return {"status": "ok", "service": APP_NAME}

@app.post("/auth/login", response_model=LoginResponse)
def login(p: LoginRequest, db: Session = Depends(get_db)):
    user_info = authenticate(p.username, p.password, db=db)
    if not user_info:
        raise HTTPException(401, "Invalid username or password.")
    token = create_access_token(user_info["username"], role=user_info["role"])
    return {
        "access_token": token,
        "token_type": "bearer",
        "expires_in_minutes": ACCESS_TOKEN_MINUTES,
        "username": user_info["username"],
        "role": user_info["role"]
    }

def persist_job_data(
    db: Session, job_token: str, username: str,
    att_filename: str, swipe_filename: str,
    valid: int, invalid: int, dates: int,
    swipe_data: dict, employee_names: dict
):
    """Persist a ProcessingJob with AttendanceRecord rows for ALL employees.
    Employees with a swipe match → PRESENT with first_punch time.
    Employees in attendance master but no swipe → ABSENT with first_punch=None.
    """
    try:
        job_record = ProcessingJob(
            job_uuid=job_token,
            username=username,
            attendance_filename=att_filename,
            swipe_filename=swipe_filename,
            valid_swipes_count=valid,
            invalid_swipes_count=invalid,
            date_count=dates,
            status="COMPLETED",
            created_at=datetime.now(timezone.utc)
        )
        db.add(job_record)

        # Collect all unique dates seen across all swiped employees
        all_dates: set = set()
        for date_dict in swipe_data.values():
            all_dates.update(date_dict.keys())

        # Upsert every employee found in the attendance master
        all_emp_codes = set(employee_names.keys()) | set(swipe_data.keys())
        for emp_code in all_emp_codes:
            emp_name = employee_names.get(emp_code, "")
            emp = db.query(Employee).filter(Employee.employee_code == emp_code).first()
            if not emp:
                emp = Employee(
                    employee_code=emp_code,
                    employee_name=emp_name or emp_code
                )
                db.add(emp)
            elif emp_name and not emp_name.startswith("Employee "):
                # Update name if we now have the real one
                emp.employee_name = emp_name

            # Store AttendanceRecord for every date in this batch
            for d_obj in all_dates:
                dt_obj = swipe_data.get(emp_code, {}).get(d_obj)
                if dt_obj:
                    # Swipe log
                    db.add(SwipeLog(
                        job_uuid=job_token,
                        employee_code=emp_code,
                        swipe_time=dt_obj,
                        swipe_type="IN"
                    ))
                db.add(AttendanceRecord(
                    job_uuid=job_token,
                    employee_code=emp_code,
                    record_date=d_obj,
                    first_punch=dt_obj.strftime("%H:%M:%S") if dt_obj else None,
                    total_swipes=1 if dt_obj else 0,
                    status="PRESENT" if dt_obj else "ABSENT"
                ))
        db.commit()
    except Exception as e:
        db.rollback()
        logger.error(f"Failed to persist job details to DB: {e}", exc_info=True)

@app.post("/process")
async def process(
    attendance_file: UploadFile = File(...),
    swipe_file: UploadFile = File(...),
    attendance_password: str = Form(""),
    swipe_password: str = Form(""),
    user_payload: dict = Depends(auth),
    db: Session = Depends(get_db)
):
    if Path(attendance_file.filename or "").suffix.lower() not in {".xlsx", ".xlsm"}:
        raise HTTPException(400, "Attendance file must be .xlsx or .xlsm.")
    if Path(swipe_file.filename or "").suffix.lower() not in {".xlsx", ".xlsm"}:
        raise HTTPException(400, "Swipe file must be .xlsx or .xlsm.")
    
    job_token = secrets.token_hex(16)
    job = BASE / job_token
    job.mkdir()
    ap = job / ("attendance" + Path(attendance_file.filename).suffix.lower())
    sp = job / ("swipes" + Path(swipe_file.filename).suffix.lower())
    op = job / "Updated_Attendance_Punchin.xlsx"
    
    try:
        await save_upload(attendance_file, ap)
        await save_upload(swipe_file, sp)

        # Read employee names from attendance master BEFORE processing
        employee_names = {}
        try:
            employee_names = read_attendance_employees(str(ap), attendance_password)
        except Exception as name_err:
            logger.warning(f"Could not read employee names from attendance file: {name_err}")

        data, valid, invalid, swipe_names = read_swipes(str(sp), swipe_password)
        dates = update_attendance(str(ap), attendance_password, data, str(op))

        # Merge: attendance names take priority; swipe names fill gaps
        merged_names = {**swipe_names, **{k: v for k, v in employee_names.items() if v}}

        # Persist to DB with real employee names and ALL employees (present + absent)
        username = user_payload.get("sub", "coordinator")
        persist_job_data(db, job_token, username, attendance_file.filename, swipe_file.filename, valid, invalid, dates, data, merged_names)

        resp = FileResponse(
            str(op),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="Updated_Attendance_Punchin.xlsx"
        )
        resp.headers["X-Valid-Swipes"] = str(valid)
        resp.headers["X-Invalid-Swipes"] = str(invalid)
        resp.headers["X-Date-Count"] = str(dates)
        resp.headers["Access-Control-Expose-Headers"] = "X-Valid-Swipes, X-Invalid-Swipes, X-Date-Count"
        resp.background = BackgroundTask(lambda: shutil.rmtree(job, ignore_errors=True))
        return resp
    except HTTPException:
        shutil.rmtree(job, ignore_errors=True)
        raise
    except ValueError as e:
        shutil.rmtree(job, ignore_errors=True)
        raise HTTPException(400, str(e))
    except Exception as e:
        shutil.rmtree(job, ignore_errors=True)
        raise HTTPException(500, f"Unexpected processing error: {str(e)}")

@app.get("/jobs")
def get_jobs(db: Session = Depends(get_db), user_payload: dict = Depends(auth)):
    jobs = db.query(ProcessingJob).order_by(ProcessingJob.created_at.desc()).limit(50).all()
    return [{
        "job_uuid": j.job_uuid,
        "username": j.username,
        "attendance_filename": j.attendance_filename,
        "swipe_filename": j.swipe_filename,
        "valid_swipes_count": j.valid_swipes_count,
        "invalid_swipes_count": j.invalid_swipes_count,
        "date_count": j.date_count,
        "status": j.status,
        "created_at": j.created_at.isoformat() if j.created_at else None
    } for j in jobs]


def get_cycle_label(dt) -> str:
    """Return 20th-to-20th attendance cycle label, e.g. '20 Aug → 20 Sep 2026'."""
    if isinstance(dt, str):
        try:
            dt = datetime.fromisoformat(dt.replace("Z", "+00:00"))
        except Exception:
            dt = datetime.now()
    elif not dt or not hasattr(dt, "day"):
        dt = datetime.now()
    day, month, year = dt.day, dt.month, dt.year
    if day < 20:
        # Before 20th: cycle started on 20th of previous month
        if month == 1:
            start_month, start_year = 12, year - 1
        else:
            start_month, start_year = month - 1, year
        end_month, end_year = month, year
    else:
        # On or after 20th: cycle starts this month
        start_month, start_year = month, year
        if month == 12:
            end_month, end_year = 1, year + 1
        else:
            end_month, end_year = month + 1, year
    start_dt = datetime(start_year, start_month, 20)
    end_dt = datetime(end_year, end_month, 20)
    return f"20 {start_dt.strftime('%b')} → 20 {end_dt.strftime('%b')} {end_year}"


@app.get("/reports")
def list_reports(db: Session = Depends(get_db), user_payload: dict = Depends(auth)):
    """Return all processed report batches with 20th-to-20th cycle labels."""
    jobs = db.query(ProcessingJob).order_by(ProcessingJob.created_at.desc()).limit(100).all()
    result = []
    for j in jobs:
        dt = j.created_at or datetime.now(timezone.utc)
        result.append({
            "job_uuid": j.job_uuid,
            "cycle_label": get_cycle_label(dt),
            "username": j.username,
            "attendance_filename": j.attendance_filename,
            "swipe_filename": j.swipe_filename,
            "valid_swipes_count": j.valid_swipes_count,
            "invalid_swipes_count": j.invalid_swipes_count,
            "date_count": j.date_count,
            "status": j.status,
            "created_at": j.created_at.isoformat() if j.created_at else None
        })
    return result


@app.get("/reports/{job_uuid}")
def get_report_detail(job_uuid: str, db: Session = Depends(get_db), user_payload: dict = Depends(auth)):
    """Return full analytics for a specific processed report, including daily punch-in times."""
    job = db.query(ProcessingJob).filter(ProcessingJob.job_uuid == job_uuid).first()
    if not job:
        raise HTTPException(404, f"Report '{job_uuid}' not found.")

    records = db.query(AttendanceRecord).filter(AttendanceRecord.job_uuid == job_uuid).all()

    daily_counts: dict = {}
    emp_punch_count: dict = {}
    employees: set = set()
    employee_daily_records: list = []  # per-employee per-date detail rows

    for rec in records:
        employees.add(rec.employee_code)
        date_label = rec.record_date.strftime("%d-%b") if rec.record_date else "Unknown"

        if date_label not in daily_counts:
            daily_counts[date_label] = {"punches": 0, "total": 0, "times": [], "first_punch": None, "last_punch": None}
        daily_counts[date_label]["total"] += 1
        emp_punch_count.setdefault(rec.employee_code, 0)

        if rec.first_punch:
            daily_counts[date_label]["punches"] += 1
            daily_counts[date_label]["times"].append(rec.first_punch)
            emp_punch_count[rec.employee_code] += 1
            fp = rec.first_punch
            curr_first = daily_counts[date_label]["first_punch"]
            curr_last = daily_counts[date_label]["last_punch"]
            if curr_first is None or fp < curr_first:
                daily_counts[date_label]["first_punch"] = fp
            if curr_last is None or fp > curr_last:
                daily_counts[date_label]["last_punch"] = fp

        # Collect per-employee row
        employee_daily_records.append({
            "date": date_label,
            "employee_code": rec.employee_code,
            "punch_in_time": rec.first_punch or "—",
            "status": "Present" if rec.first_punch else "Absent"
        })

    # Sort dates chronologically
    def _sort_key(d_label):
        try:
            return datetime.strptime(d_label, "%d-%b")
        except Exception:
            return datetime.min

    sorted_daily = {k: daily_counts[k] for k in sorted(daily_counts, key=_sort_key)}

    # Sort employee_daily_records by date then employee_code
    employee_daily_records.sort(key=lambda x: (_sort_key(x["date"]), x["employee_code"]))

    # Build month_counts derived from daily
    month_counts: dict = {}
    for d_label, d_info in sorted_daily.items():
        parts = d_label.split("-")
        m_key = parts[1].capitalize() if len(parts) >= 2 else d_label
        month_counts.setdefault(m_key, {"punches": 0, "total": 0})
        month_counts[m_key]["punches"] += d_info["punches"]
        month_counts[m_key]["total"] += d_info["total"]
    for m in month_counts.values():
        m["rate"] = round((m["punches"] / m["total"] * 100), 1) if m["total"] > 0 else 0.0

    # Fetch employee names — use whatever is stored, no filtering
    emp_names: dict = {}
    emp_objs = db.query(Employee).filter(Employee.employee_code.in_(list(employees))).all()
    for emp_obj in emp_objs:
        if emp_obj.employee_name:
            # Prefer real names; fall back to stored code-as-name rather than showing nothing
            stored = emp_obj.employee_name.strip()
            # Skip auto-generated dummy "Employee XXXX" placeholder names
            if stored and not (stored.startswith("Employee ") and stored[9:].isdigit()):
                emp_names[emp_obj.employee_code] = stored

    # Inject employee names into employee_daily_records
    for row in employee_daily_records:
        row["employee_name"] = emp_names.get(row["employee_code"], "")

    num_employees = len(employees)
    total_dates = len(sorted_daily)
    total_punches = sum(d["punches"] for d in sorted_daily.values())
    total_possible = sum(d["total"] for d in sorted_daily.values())
    completion_rate = round((total_punches / total_possible * 100), 1) if total_possible > 0 else 0.0

    # Build Excel grid matrix: each row is an employee, columns are dates
    date_keys = list(sorted_daily.keys())
    emp_matrix_map: dict = {}
    for code in employees:
        emp_matrix_map[code] = {
            "Employee Code": code,
            "Employee Name": emp_names.get(code, ""),
            "Present Days": emp_punch_count.get(code, 0),
            "Attendance Rate": f"{round((emp_punch_count.get(code, 0) / total_dates * 100), 1)}%" if total_dates > 0 else "0%"
        }
        for dk in date_keys:
            emp_matrix_map[code][dk] = "—"

    for r in employee_daily_records:
        c = r["employee_code"]
        d = r["date"]
        if c in emp_matrix_map and d in date_keys:
            emp_matrix_map[c][d] = r["punch_in_time"] if r["punch_in_time"] != "—" else "ABSENT"

    excel_grid_rows = list(emp_matrix_map.values())
    excel_grid_rows.sort(key=lambda x: (x["Employee Name"] or x["Employee Code"]).lower())

    dt = job.created_at or datetime.now(timezone.utc)
    return {
        "job_uuid": job_uuid,
        "cycle_label": get_cycle_label(dt),
        "attendance_filename": job.attendance_filename,
        "swipe_filename": job.swipe_filename,
        "total_employees": num_employees,
        "total_dates": total_dates,
        "total_possible": total_possible,
        "total_punches": total_punches,
        "completion_rate": completion_rate,
        "daily_counts": sorted_daily,
        "month_counts": month_counts,
        "emp_punch_count": emp_punch_count,
        "emp_names": emp_names,
        "employee_daily_records": employee_daily_records,
        "excel_date_columns": date_keys,
        "excel_grid_rows": excel_grid_rows,
        "is_swipe_log": False,
        "records": []
    }


@app.get("/analytics/dashboard")
def get_dashboard_analytics(
    financial_year: Optional[str] = Query(None, description="Financial year filter, e.g. 'FY 2026-27' or 'all'"),
    month: Optional[str] = Query(None, description="Month filter, e.g. 'July 2026', 'Jul', or 'all'"),
    job_uuid: Optional[str] = Query(None, description="Optional filter by job_uuid"),
    db: Session = Depends(get_db),
    user_payload: dict = Depends(auth)
):
    """Return aggregated analytics organized financial-year-wise and month-wise."""
    def _compute_fy(d: date) -> str:
        if d.month >= 4:
            return f"FY {d.year}-{str(d.year + 1)[-2:]}"
        else:
            return f"FY {d.year - 1}-{str(d.year)[-2:]}"

    all_dates = [r[0] for r in db.query(AttendanceRecord.record_date).distinct().all() if r[0]]
    if not all_dates:
        cur_fy = _compute_fy(date.today())
        return {
            "financial_years": [cur_fy],
            "selected_financial_year": cur_fy,
            "available_months": [],
            "selected_month": "All Months (Full FY)",
            "total_employees": 0,
            "total_dates": 0,
            "total_possible": 0,
            "total_punches": 0,
            "completion_rate": 0.0,
            "month_summary": [],
            "daily_counts": {},
            "emp_names": {},
            "employee_daily_records": [],
            "top_performers": [],
            "missing_punch_employees": []
        }

    # Extract unique financial years
    fy_map = {}
    for d in all_dates:
        fy = _compute_fy(d)
        fy_map.setdefault(fy, []).append(d)

    sorted_fys = sorted(list(fy_map.keys()), reverse=True)

    # Determine selected financial year
    if financial_year and financial_year in sorted_fys:
        active_fy = financial_year
    elif financial_year == "all":
        active_fy = "all"
    else:
        active_fy = sorted_fys[0]

    base_query = db.query(AttendanceRecord)
    if job_uuid:
        base_query = base_query.filter(AttendanceRecord.job_uuid == job_uuid)

    if active_fy != "all":
        m = re.search(r'(\d{4})', active_fy)
        start_year = int(m.group(1)) if m else 2026
        fy_start = date(start_year, 4, 1)
        fy_end = date(start_year + 1, 3, 31)
        fy_records_query = base_query.filter(
            AttendanceRecord.record_date >= fy_start,
            AttendanceRecord.record_date <= fy_end
        )
    else:
        fy_records_query = base_query

    # Available months in this FY: order by Indian FY sequence (Apr -> Mar)
    fy_dates = [d for d in all_dates if (active_fy == "all" or _compute_fy(d) == active_fy)]
    unique_months_tuples = sorted(list(set((d.year, d.month) for d in fy_dates)), key=lambda x: (x[0], (x[1] - 4) % 12))
    available_months = [date(y, m, 1).strftime("%B %Y") for (y, m) in unique_months_tuples]

    # Month filtering
    selected_month = month or "All Months (Full FY)"
    query = fy_records_query
    if selected_month and selected_month not in ("All Months (Full FY)", "all", ""):
        matched_tuple = None
        for (y, m) in unique_months_tuples:
            m_full = date(y, m, 1).strftime("%B %Y")
            m_short = date(y, m, 1).strftime("%b %Y")
            m_name_only = date(y, m, 1).strftime("%B")
            if selected_month.lower() in (m_full.lower(), m_short.lower(), m_name_only.lower()):
                matched_tuple = (y, m)
                break
        if matched_tuple:
            y, m = matched_tuple
            last_day = calendar.monthrange(y, m)[1]
            m_start = date(y, m, 1)
            m_end = date(y, m, last_day)
            query = query.filter(AttendanceRecord.record_date >= m_start, AttendanceRecord.record_date <= m_end)

    records = query.all()

    # Aggregate metrics
    daily_counts: dict = {}
    emp_punch_count: dict = {}
    employees: set = set()
    employee_daily_records: list = []

    for rec in records:
        employees.add(rec.employee_code)
        d_val = rec.record_date
        d_key = d_val.isoformat() if d_val else "unknown"
        d_display = d_val.strftime("%d-%b") if d_val else "—"

        if d_key not in daily_counts:
            daily_counts[d_key] = {
                "date_iso": d_key,
                "date_display": d_display,
                "day_name": d_val.strftime("%a") if d_val else "—",
                "month_name": d_val.strftime("%B %Y") if d_val else "—",
                "punches": 0,
                "total": 0,
                "times": [],
                "first_punch": None,
                "last_punch": None
            }
        daily_counts[d_key]["total"] += 1
        emp_punch_count.setdefault(rec.employee_code, 0)

        if rec.first_punch:
            daily_counts[d_key]["punches"] += 1
            daily_counts[d_key]["times"].append(rec.first_punch)
            emp_punch_count[rec.employee_code] += 1
            fp = rec.first_punch
            c_first = daily_counts[d_key]["first_punch"]
            c_last = daily_counts[d_key]["last_punch"]
            if c_first is None or fp < c_first:
                daily_counts[d_key]["first_punch"] = fp
            if c_last is None or fp > c_last:
                daily_counts[d_key]["last_punch"] = fp

        employee_daily_records.append({
            "date": d_key,
            "date_display": d_display,
            "employee_code": rec.employee_code,
            "punch_in_time": rec.first_punch or "—",
            "status": "Present" if rec.first_punch else "Absent"
        })

    sorted_daily = {k: daily_counts[k] for k in sorted(daily_counts.keys())}

    emp_names: dict = {}
    emp_objs = db.query(Employee).filter(Employee.employee_code.in_(list(employees))).all()
    for e in emp_objs:
        if e.employee_name:
            stored = e.employee_name.strip()
            if stored and not (stored.startswith("Employee ") and stored[9:].isdigit()):
                emp_names[e.employee_code] = stored

    for r in employee_daily_records:
        r["employee_name"] = emp_names.get(r["employee_code"], "")

    employee_daily_records.sort(key=lambda x: (x["date"], x["employee_code"]))

    total_employees = len(employees)
    total_dates = len(sorted_daily)
    total_punches = sum(d["punches"] for d in sorted_daily.values())
    total_possible = sum(d["total"] for d in sorted_daily.values())
    completion_rate = round((total_punches / total_possible * 100), 1) if total_possible > 0 else 0.0

    # Month-wise breakdown summary for the financial year
    fy_records = fy_records_query.all()
    month_summary_map: dict = {}
    for r in fy_records:
        if not r.record_date:
            continue
        m_key = r.record_date.strftime("%B %Y")
        if m_key not in month_summary_map:
            month_summary_map[m_key] = {
                "month_name": m_key,
                "month_short": r.record_date.strftime("%b"),
                "year": r.record_date.year,
                "month_num": r.record_date.month,
                "dates": set(),
                "punches": 0,
                "total": 0
            }
        month_summary_map[m_key]["dates"].add(r.record_date)
        month_summary_map[m_key]["total"] += 1
        if r.first_punch:
            month_summary_map[m_key]["punches"] += 1

    month_summary = []
    for m_info in sorted(month_summary_map.values(), key=lambda x: (x["year"], (x["month_num"] - 4) % 12)):
        tot = m_info["total"]
        punc = m_info["punches"]
        month_summary.append({
            "month_name": m_info["month_name"],
            "month_short": m_info["month_short"],
            "dates_count": len(m_info["dates"]),
            "punches": punc,
            "total": tot,
            "completion_rate": round(punc / tot * 100, 1) if tot > 0 else 0.0
        })

    top_performers = []
    missing_punch_employees = []
    sorted_emps = sorted(emp_punch_count.items(), key=lambda x: x[1], reverse=True)
    for code, count in sorted_emps[:10]:
        top_performers.append({
            "Employee Code": code,
            "Employee Name": emp_names.get(code, "—"),
            "Punch-Ins Recorded": count,
            "Attendance Rate": f"{round(count / total_dates * 100, 1)}%" if total_dates > 0 else "0%"
        })

    for code, count in sorted_emps:
        if count < total_dates:
            missing_punch_employees.append({
                "Employee Code": code,
                "Employee Name": emp_names.get(code, "—"),
                "Punches Found": count,
                "Missing Days": total_dates - count,
                "Attendance Rate": f"{round(count / total_dates * 100, 1)}%" if total_dates > 0 else "0%"
            })
    missing_punch_employees.sort(key=lambda x: x["Missing Days"], reverse=True)

    return {
        "financial_years": sorted_fys,
        "selected_financial_year": active_fy,
        "available_months": available_months,
        "selected_month": selected_month,
        "total_employees": total_employees,
        "total_dates": total_dates,
        "total_possible": total_possible,
        "total_punches": total_punches,
        "completion_rate": completion_rate,
        "month_summary": month_summary,
        "daily_counts": sorted_daily,
        "emp_names": emp_names,
        "employee_daily_records": employee_daily_records,
        "top_performers": top_performers,
        "missing_punch_employees": missing_punch_employees[:15]
    }


@app.get("/reports/{job_uuid}/download")
def download_report_excel(job_uuid: str, db: Session = Depends(get_db), user_payload: dict = Depends(auth)):
    """Generate and download the full updated attendance Excel file for a specific processed batch."""
    job = db.query(ProcessingJob).filter(ProcessingJob.job_uuid == job_uuid).first()
    if not job:
        raise HTTPException(404, f"Report '{job_uuid}' not found.")

    records = db.query(AttendanceRecord).filter(AttendanceRecord.job_uuid == job_uuid).all()
    if not records:
        raise HTTPException(404, f"No records found for report '{job_uuid}'.")

    dates = sorted(list(set(r.record_date for r in records)))
    emp_codes = sorted(list(set(r.employee_code for r in records)))

    emp_objs = db.query(Employee).filter(Employee.employee_code.in_(emp_codes)).all()
    emp_map = {e.employee_code: e for e in emp_objs}

    data_map = {}
    for r in records:
        data_map.setdefault(r.employee_code, {})[r.record_date] = r

    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from io import BytesIO

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Updated Attendance"

    header_fill = PatternFill(start_color="0077B6", end_color="0077B6", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    absent_fill = PatternFill(start_color="E5B80B", end_color="E5B80B", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    headers = ["Employee Code", "Employee Name", "Area"]
    for d in dates:
        headers.append(d.strftime("%d-%b") + " Punchin Time")
    ws.append(headers)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for code in emp_codes:
        emp = emp_map.get(code)
        emp_name = emp.employee_name if emp else ""
        emp_area = emp.area if emp and emp.area else ""
        row = [code, emp_name, emp_area]
        for d in dates:
            rec = data_map.get(code, {}).get(d)
            fp = rec.first_punch if rec and rec.first_punch else ""
            row.append(fp)
        ws.append(row)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border
            if cell.column > 3:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if not cell.value:
                    cell.fill = absent_fill

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 13)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    filename = f"Updated_Attendance_Punchin_{job_uuid[:8]}.xlsx"
    return Response(
        content=buf.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Access-Control-Expose-Headers": "Content-Disposition"
        }
    )



@app.post("/admin/update-emp-names")
async def update_emp_names(
    attendance_file: UploadFile = File(...),
    attendance_password: str = Form(""),
    db: Session = Depends(get_db),
    user_payload: dict = Depends(auth)
):
    """Backfill real employee names into the Employee table from an uploaded attendance or swipe Excel.
    Use this when existing records show dummy names (Employee 1031) instead of real names.
    """
    job = BASE / secrets.token_hex(8)
    job.mkdir(parents=True, exist_ok=True)
    ap = job / ("attendance" + Path(attendance_file.filename or "file.xlsx").suffix.lower())
    try:
        await save_upload(attendance_file, ap)
        # Try reading names from attendance master first
        emp_map = {}
        try:
            emp_map = read_attendance_employees(str(ap), attendance_password)
        except Exception:
            pass
        # If nothing found, try reading as swipe log (which also has Employee Name)
        if not emp_map:
            try:
                _, _, _, swipe_names = read_swipes(str(ap), attendance_password)
                emp_map = swipe_names
            except Exception:
                pass

        updated = 0
        for emp_code, emp_name in emp_map.items():
            if not emp_name:
                continue
            emp_obj = db.query(Employee).filter(Employee.employee_code == emp_code).first()
            if emp_obj:
                emp_obj.employee_name = emp_name
                updated += 1
            else:
                db.add(Employee(employee_code=emp_code, employee_name=emp_name))
                updated += 1
        db.commit()
        return {"status": "ok", "employees_updated": updated, "total_found": len(emp_map)}
    except Exception as e:
        db.rollback()
        raise HTTPException(500, f"Failed to update employee names: {e}")
    finally:
        shutil.rmtree(job, ignore_errors=True)


# ─── Employee CRUD Endpoints ──────────────────────────────────────────

@app.get("/employees", response_model=List[EmployeeResponse])
def get_employees(
    search: Optional[str] = Query(None, description="Search by code or name"),
    active_only: Optional[bool] = Query(None, description="Filter active status only"),
    skip: int = Query(0, ge=0),
    limit: int = Query(500, ge=1, le=1000),
    db: Session = Depends(get_db),
    user_payload: dict = Depends(auth)
):
    """Retrieve all employees with optional search filtering by name or employee code."""
    query = db.query(Employee)
    if active_only is not None:
        query = query.filter(Employee.is_active == active_only)
    if search:
        term = f"%{search.strip()}%"
        query = query.filter(
            (Employee.employee_name.ilike(term)) | (Employee.employee_code.ilike(term))
        )
    return query.order_by(Employee.employee_name.asc()).offset(skip).limit(limit).all()


@app.get("/employees/{employee_id}", response_model=EmployeeResponse)
def get_employee(
    employee_id: int,
    db: Session = Depends(get_db),
    user_payload: dict = Depends(auth)
):
    """Retrieve single employee details by ID."""
    emp = db.query(Employee).filter(Employee.id == employee_id).first()
    if not emp:
        raise HTTPException(status_code=404, detail="Employee not found.")
    return emp


@app.post("/employees", response_model=EmployeeResponse, status_code=status.HTTP_201_CREATED)
def create_employee(
    payload: EmployeeCreate,
    db: Session = Depends(get_db),
    user_payload: dict = Depends(auth)
):
    """Create a new employee with employee code and employee name."""
    code = payload.employee_code.strip()
    name = payload.employee_name.strip()
    if not code:
        raise HTTPException(status_code=400, detail="Employee code cannot be empty.")
    if not name:
        raise HTTPException(status_code=400, detail="Employee name cannot be empty.")

    # Check for duplicate employee code
    existing = db.query(Employee).filter(Employee.employee_code == code).first()
    if existing:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=f"Employee code '{code}' already exists (assigned to {existing.employee_name})."
        )

    new_emp = Employee(
        employee_code=code,
        employee_name=name,
        area=payload.area.strip() if payload.area and payload.area.strip() else None,
        attendance_mode=payload.attendance_mode.strip() if payload.attendance_mode and payload.attendance_mode.strip() else None,
        created_at=datetime.now(timezone.utc)
    )
    db.add(new_emp)
    db.commit()
    db.refresh(new_emp)
    return new_emp


@app.put("/employees/{employee_id}", response_model=EmployeeResponse)
def update_employee(
    employee_id: int,
    payload: EmployeeUpdate,
    db: Session = Depends(get_db),
    user_payload: dict = Depends(auth)
):
    """Update employee details (name, code, area, attendance mode)."""
    emp = db.query(Employee).filter(Employee.id == employee_id).first()
    if not emp:
        raise HTTPException(status_code=404, detail="Employee not found.")

    # If updating code, ensure uniqueness
    if payload.employee_code is not None:
        new_code = payload.employee_code.strip()
        if not new_code:
            raise HTTPException(status_code=400, detail="Employee code cannot be empty.")
        if new_code != emp.employee_code:
            conflict = db.query(Employee).filter(
                Employee.employee_code == new_code,
                Employee.id != employee_id
            ).first()
            if conflict:
                raise HTTPException(
                    status_code=status.HTTP_409_CONFLICT,
                    detail=f"Employee code '{new_code}' already exists (assigned to {conflict.employee_name})."
                )
            old_code = emp.employee_code
            emp.employee_code = new_code
            # Cascade code change to historical records
            db.query(AttendanceRecord).filter(AttendanceRecord.employee_code == old_code).update(
                {AttendanceRecord.employee_code: new_code}
            )
            db.query(SwipeLog).filter(SwipeLog.employee_code == old_code).update(
                {SwipeLog.employee_code: new_code}
            )

    if payload.employee_name is not None:
        new_name = payload.employee_name.strip()
        if not new_name:
            raise HTTPException(status_code=400, detail="Employee name cannot be empty.")
        emp.employee_name = new_name

    if payload.area is not None:
        emp.area = payload.area.strip() if payload.area and payload.area.strip() else None

    if payload.attendance_mode is not None:
        emp.attendance_mode = payload.attendance_mode.strip() if payload.attendance_mode and payload.attendance_mode.strip() else None

    if payload.is_active is not None:
        emp.is_active = payload.is_active

    db.commit()
    db.refresh(emp)
    return emp


@app.delete("/employees/{employee_id}")
def delete_employee(
    employee_id: int,
    db: Session = Depends(get_db),
    user_payload: dict = Depends(auth)
):
    """Disable an employee instead of permanently deleting from the master database."""
    emp = db.query(Employee).filter(Employee.id == employee_id).first()
    if not emp:
        raise HTTPException(status_code=404, detail="Employee not found.")

    emp.is_active = False
    db.commit()
    db.refresh(emp)
    return {
        "status": "ok",
        "message": f"Employee '{emp.employee_name}' ({emp.employee_code}) disabled successfully.",
        "employee_id": employee_id,
        "employee_code": emp.employee_code,
        "is_active": False
    }


@app.post("/employees/{employee_id}/toggle-status")
@app.patch("/employees/{employee_id}/toggle-status")
def toggle_employee_status(
    employee_id: int,
    db: Session = Depends(get_db),
    user_payload: dict = Depends(auth)
):
    """Toggle an employee's active / disabled status."""
    emp = db.query(Employee).filter(Employee.id == employee_id).first()
    if not emp:
        raise HTTPException(status_code=404, detail="Employee not found.")

    curr = getattr(emp, "is_active", True)
    if curr is None:
        curr = True
    emp.is_active = not curr
    db.commit()
    db.refresh(emp)
    action = "enabled" if emp.is_active else "disabled"
    return {
        "status": "ok",
        "message": f"Employee '{emp.employee_name}' ({emp.employee_code}) {action} successfully.",
        "employee_id": emp.id,
        "employee_code": emp.employee_code,
        "is_active": emp.is_active
    }


# ─── Manual Punch-in Endpoints ────────────────────────────────────────

@app.get("/attendance/check", response_model=AttendanceCheckResponse)
def check_attendance_record(
    employee_code: str = Query(..., min_length=1, description="Employee code"),
    date_val: date = Query(..., alias="date", description="Date to check (YYYY-MM-DD)"),
    db: Session = Depends(get_db),
    user_payload: dict = Depends(auth)
):
    """Check if an attendance record already exists for the specified employee and date."""
    role = user_payload.get("role", "")
    if not is_authorized_for_manual_punchin(role):
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Access denied. Only Coordinator and Dean-Faculty roles are authorized."
        )

    code = employee_code.strip()
    emp = db.query(Employee).filter(Employee.employee_code == code).first()
    emp_name = emp.employee_name if emp else None

    rec = db.query(AttendanceRecord).filter(
        AttendanceRecord.employee_code == code,
        AttendanceRecord.record_date == date_val
    ).first()

    if rec:
        return AttendanceCheckResponse(
            exists=True,
            employee_code=code,
            employee_name=emp_name,
            date=str(date_val),
            first_punch=rec.first_punch,
            status=rec.status
        )
    return AttendanceCheckResponse(
        exists=False,
        employee_code=code,
        employee_name=emp_name,
        date=str(date_val),
        first_punch=None,
        status=None
    )


@app.post("/attendance/manual-punch-in", response_model=ManualPunchInResponse)
@app.post("/attendance/manual-punchin", response_model=ManualPunchInResponse)
@app.post("/api/attendance/manual-punchin", response_model=ManualPunchInResponse)
def manual_punch_in(
    payload: ManualPunchInRequest,
    db: Session = Depends(get_db),
    user_payload: dict = Depends(auth)
):
    """Manually add or update a punch-in record for an employee on a given date.
    Authorized roles: Coordinator and Dean-Faculty.
    """
    role = user_payload.get("role", "")
    if not is_authorized_for_manual_punchin(role):
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Access denied. Only Coordinator and Dean-Faculty roles are authorized to perform manual punch-in."
        )

    # 1. Validate employee existence
    code = payload.employee_code.strip()
    if not code:
        raise HTTPException(status_code=400, detail="Employee code cannot be empty.")

    emp = db.query(Employee).filter(Employee.employee_code == code).first()
    if not emp:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f"Employee with code '{code}' does not exist in master records."
        )

    # 2. Validate and parse punch-in time
    time_raw = payload.punch_in_time.strip()
    parsed_time = None
    for fmt in ("%H:%M:%S", "%H:%M", "%I:%M:%S %p", "%I:%M %p"):
        try:
            parsed_time = datetime.strptime(time_raw, fmt).time()
            break
        except ValueError:
            pass

    if not parsed_time:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Invalid punch-in time '{time_raw}'. Expected format: HH:MM:SS or HH:MM."
        )

    formatted_time = parsed_time.strftime("%H:%M:%S")
    punch_date = payload.date
    combined_dt = datetime.combine(punch_date, parsed_time)
    username = user_payload.get("sub", "coordinator")

    try:
        # 3. Check for existing attendance records
        existing_records = db.query(AttendanceRecord).filter(
            AttendanceRecord.employee_code == code,
            AttendanceRecord.record_date == punch_date
        ).all()

        action = "updated" if existing_records else "created"

        if existing_records:
            is_same = all(r.first_punch == formatted_time for r in existing_records)
            for rec in existing_records:
                rec.first_punch = formatted_time
                rec.status = "PRESENT"
                if not rec.total_swipes or rec.total_swipes < 1:
                    rec.total_swipes = 1
            message = "Punch-in updated successfully" if not is_same else f"Punch-in already recorded as {formatted_time}"
        else:
            new_record = AttendanceRecord(
                job_uuid="MANUAL",
                employee_code=code,
                record_date=punch_date,
                first_punch=formatted_time,
                last_punch=None,
                total_swipes=1,
                status="PRESENT",
                created_at=datetime.now(timezone.utc)
            )
            db.add(new_record)
            message = "Manual punch-in added successfully"

        # 4. Record swipe log audit entry
        swipe_log = SwipeLog(
            job_uuid=existing_records[0].job_uuid if existing_records and existing_records[0].job_uuid else "MANUAL",
            employee_code=code,
            swipe_time=combined_dt,
            swipe_type="IN",
            area=f"Manual by {username}",
            attendance_mode="Manual",
            created_at=datetime.now(timezone.utc)
        )
        db.add(swipe_log)

        db.commit()

        return ManualPunchInResponse(
            success=True,
            message=message,
            action=action,
            record=ManualPunchInRecord(
                employee_code=code,
                employee_name=emp.employee_name,
                date=str(punch_date),
                punch_in_time=formatted_time,
                status="PRESENT",
                action=action
            )
        )
    except HTTPException:
        db.rollback()
        raise
    except Exception as err:
        db.rollback()
        logger.error(f"Failed to persist manual punch-in for {code} on {punch_date}: {err}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Database error during manual punch-in: {str(err)}")


