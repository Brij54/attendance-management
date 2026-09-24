import hashlib
from datetime import datetime, timedelta, timezone
from jose import jwt
from sqlalchemy.orm import Session
from .config import (
    JWT_SECRET, JWT_ALGORITHM, ACCESS_TOKEN_MINUTES,
    COORDINATOR_USERNAME, COORDINATOR_PASSWORD_HASH, COORDINATOR_PASSWORD,
    DEAN_FACULTY_USERNAME, DEAN_FACULTY_PASSWORD_HASH, DEAN_FACULTY_PASSWORD
)

def hash_md5(val: str) -> str:
    return hashlib.md5(val.encode("utf-8")).hexdigest()

def hash_sha256(val: str) -> str:
    return hashlib.sha256(val.encode("utf-8")).hexdigest()

def verify_password(plain_password: str, stored_hash_or_plain: str) -> bool:
    if not plain_password or not stored_hash_or_plain:
        return False
    md5_input = hash_md5(plain_password)
    sha256_input = hash_sha256(plain_password)
    stored_clean = stored_hash_or_plain.strip().lower()
    
    return (
        md5_input == stored_clean or
        sha256_input == stored_clean or
        plain_password == stored_hash_or_plain
    )

def create_access_token(username: str, role: str = "coordinator"):
    exp = datetime.now(timezone.utc) + timedelta(minutes=ACCESS_TOKEN_MINUTES)
    return jwt.encode({"sub": username, "role": role, "exp": exp}, JWT_SECRET, algorithm=JWT_ALGORITHM)

def authenticate(username: str, password: str, db: Session = None):
    u_clean = username.strip().lower()
    
    # 1. Try DB lookup if session available
    if db is not None:
        try:
            from .models import User
            db_user = db.query(User).filter(User.username.ilike(u_clean)).first()
            if db_user:
                if verify_password(password, db_user.password_hash):
                    return {"username": db_user.username, "role": db_user.role}
        except Exception:
            pass

    # 2. Fallback check for Coordinator account from ENV
    if u_clean in (COORDINATOR_USERNAME.lower(), "coordinator"):
        if verify_password(password, COORDINATOR_PASSWORD_HASH) or verify_password(password, COORDINATOR_PASSWORD):
            return {"username": COORDINATOR_USERNAME, "role": "coordinator"}
            
    # 3. Fallback check for Dean-Faculty account from ENV
    if u_clean in (DEAN_FACULTY_USERNAME.lower(), "dean-faculty", "dean", "faculty"):
        if verify_password(password, DEAN_FACULTY_PASSWORD_HASH) or verify_password(password, DEAN_FACULTY_PASSWORD):
            return {"username": DEAN_FACULTY_USERNAME, "role": "dean-faculty"}

    return None

def seed_default_users(db: Session):
    try:
        from .models import User
        users_to_seed = [
            (COORDINATOR_USERNAME, COORDINATOR_PASSWORD_HASH, "coordinator"),
            (DEAN_FACULTY_USERNAME, DEAN_FACULTY_PASSWORD_HASH, "dean-faculty")
        ]
        for uname, phash, role in users_to_seed:
            existing = db.query(User).filter(User.username.ilike(uname)).first()
            if not existing:
                db.add(User(username=uname, password_hash=phash, role=role))
        db.commit()
    except Exception:
        db.rollback()

def decode_token(token: str):
    payload = jwt.decode(token, JWT_SECRET, algorithms=[JWT_ALGORITHM])
    if not payload.get("sub"):
        raise ValueError("Invalid token")
    return payload

ALLOWED_MANUAL_PUNCHIN_ROLES = {"coordinator", "dean-faculty"}

def is_authorized_for_manual_punchin(role: str) -> bool:
    """Check whether a user role is permitted to perform manual punch-in.
    Permitted roles: 'coordinator' and 'dean-faculty' (case-insensitive, supporting 'dean_faculty').
    """
    if not role:
        return False
    clean_role = role.strip().lower().replace("_", "-")
    return clean_role in ALLOWED_MANUAL_PUNCHIN_ROLES
