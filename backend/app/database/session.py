from sqlalchemy.orm import Session
from .database import SessionLocal

def get_db() -> Session:
    """
    Dependency to get database session
    """
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()