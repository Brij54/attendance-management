from datetime import datetime, timezone
from sqlalchemy import Column, Integer, String, Float, DateTime, Date, ForeignKey, Text
from sqlalchemy.orm import relationship
from .database import Base

class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    username = Column(String(50), unique=True, index=True, nullable=False)
    password_hash = Column(String(128), nullable=False)
    role = Column(String(50), nullable=False, default="coordinator")
    created_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))

class Employee(Base):
    __tablename__ = "employees"

    id = Column(Integer, primary_key=True, index=True)
    employee_code = Column(String(50), unique=True, index=True, nullable=False)
    employee_name = Column(String(150), nullable=False)
    area = Column(String(100), nullable=True)
    attendance_mode = Column(String(50), nullable=True)
    created_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))

class SwipeLog(Base):
    __tablename__ = "swipe_logs"

    id = Column(Integer, primary_key=True, index=True)
    job_uuid = Column(String(50), index=True, nullable=True)
    employee_code = Column(String(50), index=True, nullable=False)
    swipe_time = Column(DateTime, index=True, nullable=False)
    swipe_type = Column(String(20), nullable=True)
    area = Column(String(100), nullable=True)
    attendance_mode = Column(String(50), nullable=True)
    latitude = Column(Float, nullable=True)
    longitude = Column(Float, nullable=True)
    created_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))

class ProcessingJob(Base):
    __tablename__ = "processing_jobs"

    id = Column(Integer, primary_key=True, index=True)
    job_uuid = Column(String(50), unique=True, index=True, nullable=False)
    username = Column(String(50), index=True, nullable=False)
    attendance_filename = Column(String(255), nullable=True)
    swipe_filename = Column(String(255), nullable=True)
    valid_swipes_count = Column(Integer, default=0)
    invalid_swipes_count = Column(Integer, default=0)
    date_count = Column(Integer, default=0)
    status = Column(String(50), default="COMPLETED")
    created_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))

class AttendanceRecord(Base):
    __tablename__ = "attendance_records"

    id = Column(Integer, primary_key=True, index=True)
    job_uuid = Column(String(50), index=True, nullable=True)
    employee_code = Column(String(50), index=True, nullable=False)
    record_date = Column(Date, index=True, nullable=False)
    first_punch = Column(String(20), nullable=True)
    last_punch = Column(String(20), nullable=True)
    total_swipes = Column(Integer, default=0)
    status = Column(String(50), nullable=True)
    created_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))
