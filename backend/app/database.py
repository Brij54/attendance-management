import logging
from pathlib import Path
from sqlalchemy import create_engine, text
from sqlalchemy.orm import sessionmaker, declarative_base
from .config import DATABASE_URL

logger = logging.getLogger("attendance_backend")

Base = declarative_base()

def init_engine():
    """Initialize database engine with fallback support:
    1. Primary DATABASE_URL (PostgreSQL in Docker/Production)
    2. Localhost PostgreSQL fallback if host was 'postgres'
    3. SQLite local file fallback so the service never loses data
    """
    # 1. Try primary configured DATABASE_URL
    try:
        logger.info(f"Connecting to database: {DATABASE_URL.split('@')[-1] if '@' in DATABASE_URL else DATABASE_URL}")
        eng = create_engine(
            DATABASE_URL,
            pool_pre_ping=True,
            connect_args={"connect_timeout": 5} if "postgresql" in DATABASE_URL else {}
        )
        with eng.connect() as conn:
            conn.execute(text("SELECT 1"))
        logger.info("Successfully connected to primary database.")
        return eng
    except Exception as e:
        logger.warning(f"Primary database connection failed: {e}")

    # 2. If host was 'postgres' (e.g. running outside Docker container), try localhost
    if "@postgres:" in DATABASE_URL:
        local_url = DATABASE_URL.replace("@postgres:", "@localhost:")
        try:
            logger.info("Attempting connection to local PostgreSQL on localhost:5432...")
            eng = create_engine(local_url, pool_pre_ping=True, connect_args={"connect_timeout": 3})
            with eng.connect() as conn:
                conn.execute(text("SELECT 1"))
            logger.info("Successfully connected to local PostgreSQL on localhost:5432.")
            return eng
        except Exception as e:
            logger.warning(f"Localhost PostgreSQL connection failed: {e}")

    # 3. Fallback to local SQLite database
    sqlite_file = Path(__file__).resolve().parents[2] / "attendance_db.sqlite3"
    sqlite_url = f"sqlite:///{sqlite_file}"
    logger.warning(f"Falling back to local SQLite database: {sqlite_url}")
    return create_engine(sqlite_url, connect_args={"check_same_thread": False})

engine = init_engine()
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

