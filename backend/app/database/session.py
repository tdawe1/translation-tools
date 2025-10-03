from sqlalchemy.orm import Session
from .database import get_session_local

def get_db() -> Session:
    """
    Dependency to get database session
    """
    SessionLocal = get_session_local()
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()