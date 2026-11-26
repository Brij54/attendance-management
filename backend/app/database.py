import logging
import time
from sqlalchemy import create_engine, text
from sqlalchemy.orm import sessionmaker, declarative_base

from .config import DATABASE_URL

logger = logging.getLogger("attendance_backend")

Base = declarative_base()


def init_engine():
    """
    Initialize PostgreSQL database engine.

    Production configuration:
    - Connection health checking
    - Automatic connection recycling
    - Connection pooling
    - Connection timeout
    - Transient startup retry handling
    - Automatic driver detection (psycopg2 vs psycopg3)
    - No silent SQLite fallback
    """
    global DATABASE_URL

    if not DATABASE_URL:
        raise RuntimeError("DATABASE_URL is not configured")

    # SQLAlchemy 2.0+ defaults generic 'postgresql://' to 'psycopg' (Psycopg 3).
    # If psycopg (v3) is not installed, route to psycopg2 (psycopg2-binary).
    try:
        import psycopg  # noqa: F401
    except ImportError:
        if DATABASE_URL.startswith("postgresql+psycopg://"):
            DATABASE_URL = DATABASE_URL.replace("postgresql+psycopg://", "postgresql+psycopg2://", 1)
        elif DATABASE_URL.startswith("postgres://"):
            DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql+psycopg2://", 1)
        elif DATABASE_URL.startswith("postgresql://"):
            DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+psycopg2://", 1)
    else:
        if DATABASE_URL.startswith("postgres://"):
            DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

    if not DATABASE_URL.startswith(
        ("postgresql://", "postgresql+psycopg2://", "postgresql+psycopg://")
    ):
        raise RuntimeError(
            "Production requires PostgreSQL. "
            "SQLite fallback is disabled."
        )

    logger.info("Initializing PostgreSQL database connection pool")

    eng = create_engine(
        DATABASE_URL,
        pool_pre_ping=True,
        pool_recycle=1800,
        pool_size=10,
        max_overflow=5,
        pool_timeout=30,
        connect_args={
            "connect_timeout": 5
        }
    )

    # Health check with brief retry to guard against container boot lag
    max_retries = 3
    for attempt in range(1, max_retries + 1):
        try:
            with eng.connect() as conn:
                conn.execute(text("SELECT 1"))
            logger.info("Successfully connected to PostgreSQL")
            return eng
        except Exception:
            if attempt < max_retries:
                logger.warning(
                    f"Database connection attempt {attempt}/{max_retries} failed; retrying in 2s..."
                )
                time.sleep(2)
            else:
                eng.dispose()
                logger.exception("PostgreSQL connection initialization failed after retries")
                raise


engine = init_engine()

SessionLocal = sessionmaker(
    autocommit=False,
    autoflush=False,
    bind=engine
)


def get_db():
    db = SessionLocal()
    try:
        yield db
    except Exception:
        db.rollback()
        raise
    finally:
        db.close()
