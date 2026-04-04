from werkzeug.security import generate_password_hash

from ..database import db
from ..models import AppUser
from .audit import log_action


def ensure_default_admin_user(username, password, actor="system-seed"):
    admin_username = (username or "admin").strip().lower() or "admin"
    admin_password = str(password or "123456")

    existing = AppUser.query.filter_by(username=admin_username).first()
    if existing:
        should_commit = False

        if not existing.password_hash:
            existing.password_hash = generate_password_hash(admin_password)
            should_commit = True

        if not existing.is_admin:
            existing.is_admin = True
            should_commit = True

        if not existing.is_active:
            existing.is_active = True
            should_commit = True

        if should_commit:
            db.session.commit()

        return existing

    user = AppUser(
        username=admin_username,
        full_name="System Admin",
        password_hash=generate_password_hash(admin_password),
        is_admin=True,
        is_active=True,
    )
    db.session.add(user)
    db.session.flush()

    log_action(
        "app_users",
        str(user.id),
        "INSERT",
        changed_by=actor,
        after_data=user.to_dict(),
        notes="Seed tai khoan admin mac dinh",
    )
    db.session.commit()
    return user
