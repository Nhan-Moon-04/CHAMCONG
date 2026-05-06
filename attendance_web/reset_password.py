#!/usr/bin/env python
"""Script to reset admin password in database"""
import os
import sys
from pathlib import Path

# Add parent dir to path
sys.path.insert(0, str(Path(__file__).parent))

from werkzeug.security import generate_password_hash
from app.database import db
from app.models import AppUser
from app import create_app

app = create_app()

with app.app_context():
    admin = AppUser.query.filter_by(username="admin").first()
    if admin:
        admin.password_hash = generate_password_hash("1234567")
        db.session.commit()
        print("✅ Password reset thành công!")
        print(f"Username: admin")
        print(f"Password: 1234567")
    else:
        print("❌ Không tìm thấy tài khoản admin")
        sys.exit(1)
