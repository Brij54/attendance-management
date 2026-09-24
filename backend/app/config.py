import os
from pathlib import Path
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parents[2]
load_dotenv(BASE_DIR / ".env.production")
load_dotenv(Path(__file__).resolve().parents[1] / ".env.production")
load_dotenv(BASE_DIR / ".env")
load_dotenv(Path(__file__).resolve().parents[1] / ".env")
load_dotenv()

APP_NAME = os.getenv(
    "APP_NAME",
    "Attendance Punch-In Automation API"
)

JWT_SECRET = os.getenv(
    "JWT_SECRET",
    "CHANGE_THIS_SECRET_IN_PRODUCTION"
)

JWT_ALGORITHM = "HS256"

ACCESS_TOKEN_MINUTES = int(
    os.getenv("ACCESS_TOKEN_MINUTES", "480")
)

COORDINATOR_USERNAME = os.getenv(
    "COORDINATOR_USERNAME",
    "coordinator"
)

COORDINATOR_PASSWORD_HASH = os.getenv(
    "COORDINATOR_PASSWORD_HASH",
    "ceb6c970658f31504a901b89dcd3e461"
)

COORDINATOR_PASSWORD = os.getenv(
    "COORDINATOR_PASSWORD",
    "test@123"
)

DEAN_FACULTY_USERNAME = os.getenv(
    "DEAN_FACULTY_USERNAME",
    "dean-faculty"
)

DEAN_FACULTY_PASSWORD_HASH = os.getenv(
    "DEAN_FACULTY_PASSWORD_HASH",
    "760a8bc0dd64752b7a905eea4a972576"
)

DEAN_FACULTY_PASSWORD = os.getenv(
    "DEAN_FACULTY_PASSWORD",
    "dean@123"
)


MAX_UPLOAD_MB = int(
    os.getenv("MAX_UPLOAD_MB", "50")
)

CORS_ORIGINS = [
    x.strip()
    for x in os.getenv(
        "CORS_ORIGINS",
        "http://localhost:8502"
    ).split(",")
    if x.strip()
]

DATABASE_URL = os.getenv(
    "DATABASE_URL",
    "postgresql://postgres:postgres_password@localhost:5432/attendance_db"
)