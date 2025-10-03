#!/usr/bin/env python3
"""Debug script to test user creation step by step"""

import os
import sys
import logging
import traceback

# Set up environment
os.environ['PYTEST_RUNNING'] = '1'
os.environ['DEBUG'] = 'true'
os.environ['SECRET_KEY'] = 'test-secret-key-for-pytest-testing-only-32-chars-long'
os.environ['OPENAI_API_KEY'] = 'mock-sk-for-testing'
os.environ['DATABASE_URL'] = 'sqlite:///:memory:'
os.environ['LOG_LEVEL'] = 'WARNING'
os.environ['UPLOAD_DIR'] = 'test_uploads'
os.environ['OUTPUT_DIR'] = 'test_outputs'

# Add path
sys.path.insert(0, '/home/thomas/translation-tools/translations-pptx-pipeline/backend')

# Enable debug logging
logging.basicConfig(level=logging.DEBUG)

try:
    from app.database.database import get_engine, Base, User as DBUser
    from app.services.auth_service import AuthService
    from app.models.auth import UserCreate
    from sqlalchemy.orm import sessionmaker

    print("=== Database Setup ===")
    engine = get_engine()
    print(f"Engine URL: {engine.url}")

    # Create tables
    Base.metadata.create_all(bind=engine)
    print("Tables created")

    # Create session
    SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
    db = SessionLocal()
    print("Session created")

    # Test user creation
    print("\n=== User Creation Test ===")
    auth_service = AuthService()
    user_data = UserCreate(
        email='test@example.com',
        password='testpassword123!',
        full_name='Test User'
    )

    try:
        print("Creating user...")
        user = auth_service.create_user(db, user_data)
        print(f"User created successfully: {user.id}")
        db.commit()
        print("Changes committed")

        # Verify user exists
        count = db.query(DBUser).count()
        print(f"Users in database: {count}")

    except Exception as e:
        print(f"Error creating user: {type(e).__name__}: {str(e)}")
        print("Full traceback:")
        traceback.print_exc()
        db.rollback()

    finally:
        db.close()

except Exception as e:
    print(f"Import error: {type(e).__name__}: {str(e)}")
    traceback.print_exc()