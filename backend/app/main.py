import os
import shutil
import secrets
import logging
from datetime import datetime, timezone
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, Form, Depends, HTTPException, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from starlette.background import BackgroundTask
from sqlalchemy.orm import Session

from .config import *
from .schemas import *
from .database import engine, Base, get_db
from .models import User, Employee, SwipeLog, ProcessingJob, AttendanceRecord
from .security import authenticate, create_access_token, decode_token, seed_default_users
from .services.excel_service import read_swipes, update_attendance, read_attendance_employees

logger = logging.getLogger("attendance_backend")

# Create DB tables if not existing and seed default accounts
try:
    Base.metadata.create_all(bind=engine)
    db_init_session = Session(bind=engine)
    seed_default_users(db_init_session)
    db_init_session.close()
except Exception as err:
    logger.error(f"Error during DB initialization: {err}")

app = FastAPI(title=APP_NAME)
app.add_middleware(
    CORSMiddleware,
    allow_origins=CORS_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"]
)

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
            status="COMPLETED"
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
        logger.error(f"Failed to persist job details to DB: {e}")

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


def get_cycle_label(dt: datetime) -> str:
    """Return 20th-to-20th attendance cycle label, e.g. '20 Aug → 20 Sep 2026'."""
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
        "is_swipe_log": False,
        "records": []
    }


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
