from datetime import datetime, date as dt_date
from typing import Optional
from pydantic import BaseModel, Field

class LoginRequest(BaseModel):
    username: str = Field(min_length=1, max_length=100)
    password: str = Field(min_length=1, max_length=200)

class LoginResponse(BaseModel):
    access_token: str
    token_type: str = "bearer"
    expires_in_minutes: int
    username: str
    role: str

class HealthResponse(BaseModel):
    status: str
    service: str

class EmployeeBase(BaseModel):
    employee_code: str = Field(min_length=1, max_length=50, description="Unique code of the employee")
    employee_name: str = Field(min_length=1, max_length=150, description="Full name of the employee")
    area: Optional[str] = Field(None, max_length=100, description="Department or location area")
    attendance_mode: Optional[str] = Field(None, max_length=50, description="Attendance mode e.g. Bio/Card/Face")
    is_active: Optional[bool] = Field(default=True, description="Whether employee is active or disabled")

class EmployeeCreate(EmployeeBase):
    pass

class EmployeeUpdate(BaseModel):
    employee_name: Optional[str] = Field(None, min_length=1, max_length=150)
    employee_code: Optional[str] = Field(None, min_length=1, max_length=50)
    area: Optional[str] = Field(None, max_length=100)
    attendance_mode: Optional[str] = Field(None, max_length=50)
    is_active: Optional[bool] = Field(None, description="Set active or disabled status")

class EmployeeResponse(EmployeeBase):
    id: int
    created_at: Optional[datetime] = None

    class Config:
        from_attributes = True
        orm_mode = True

class ManualPunchInRequest(BaseModel):
    employee_code: str = Field(..., alias="employeeCode", min_length=1, description="Employee code")
    date: dt_date = Field(..., description="Attendance date in YYYY-MM-DD format")
    punch_in_time: str = Field(..., alias="punchInTime", min_length=1, description="Punch-in time in HH:mm:ss or HH:mm format")

    class Config:
        populate_by_name = True

class ManualPunchInRecord(BaseModel):
    employee_code: str
    employee_name: Optional[str] = None
    date: str
    punch_in_time: str
    status: str
    action: Optional[str] = None

class ManualPunchInResponse(BaseModel):
    success: bool
    message: str
    action: Optional[str] = None
    record: Optional[ManualPunchInRecord] = None

class AttendanceCheckResponse(BaseModel):
    exists: bool
    employee_code: str
    employee_name: Optional[str] = None
    date: str
    first_punch: Optional[str] = None
    status: Optional[str] = None

