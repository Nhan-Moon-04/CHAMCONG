from flask import current_app, flash, redirect, render_template, request, session, url_for
from sqlalchemy import func
from werkzeug.security import check_password_hash, generate_password_hash

from .database import db
from .models import AppUser
from .services.audit import log_action


def _normalize_username(value):
    return (value or "").strip().lower()


def _set_auth_session(user):
    session["is_authenticated"] = True
    session["user_id"] = user.id
    session["username"] = user.username
    session["display_name"] = user.full_name or user.username
    session["is_admin"] = bool(user.is_admin)


def _user_audit_payload(user):
    return {
        "id": user.id,
        "username": user.username,
        "full_name": user.full_name,
        "is_admin": bool(user.is_admin),
        "is_active": bool(user.is_active),
    }


def register_user_routes(app):
    def _require_admin():
        if session.get("is_admin"):
            return None

        flash("Bạn không có quyền truy cập chức năng này", "error")
        return redirect(url_for("dashboard"))

    def _active_admin_count():
        return AppUser.query.filter_by(is_admin=True, is_active=True).count()

    @app.route("/account", methods=["GET", "POST"])
    def account_profile():
        user_id = session.get("user_id")
        user = AppUser.query.get_or_404(user_id)

        if request.method == "POST":
            action = (request.form.get("action") or "change_password").strip().lower()
            if action != "change_password":
                flash("Thao tác tài khoản không hợp lệ", "error")
                return redirect(url_for("account_profile"))

            current_password = request.form.get("current_password") or ""
            new_password = request.form.get("new_password") or ""
            confirm_password = request.form.get("confirm_password") or ""

            if not check_password_hash(user.password_hash, current_password):
                flash("Mật khẩu hiện tại không đúng", "error")
                return redirect(url_for("account_profile"))

            if len(new_password) < 4:
                flash("Mật khẩu mới tối thiểu 4 ký tự", "error")
                return redirect(url_for("account_profile"))

            if new_password != confirm_password:
                flash("Xác nhận mật khẩu mới chưa khớp", "error")
                return redirect(url_for("account_profile"))

            actor = session.get("username") or user.username or "admin"
            before_data = _user_audit_payload(user)

            user.password_hash = generate_password_hash(new_password)
            db.session.flush()

            log_action(
                "app_users",
                str(user.id),
                "UPDATE",
                changed_by=actor,
                before_data=before_data,
                after_data=_user_audit_payload(user),
                notes="User tự đổi mật khẩu",
            )
            db.session.commit()
            _set_auth_session(user)

            flash("Đã đổi mật khẩu", "success")
            return redirect(url_for("account_profile"))

        return render_template(
            "account_profile.html",
            title="Tài khoản",
            user=user,
        )

    @app.route("/users", methods=["GET"])
    def users():
        blocked = _require_admin()
        if blocked:
            return blocked

        rows = AppUser.query.order_by(AppUser.username.asc()).all()
        active_users = [row for row in rows if row.is_active]
        admin_users = [row for row in rows if row.is_admin]

        return render_template(
            "users.html",
            title="Quản lý user",
            users=rows,
            total_users=len(rows),
            active_users=len(active_users),
            admin_users=len(admin_users),
            current_user_id=session.get("user_id"),
            enable_ot_after_6pm_meal=current_app.config.get("ENABLE_OT_AFTER_6PM_MEAL", False),
        )

    @app.route("/users/toggle_ot_after_6pm", methods=["POST"])
    def toggle_ot_after_6pm():
        blocked = _require_admin()
        if blocked:
            return blocked

        current_value = current_app.config.get("ENABLE_OT_AFTER_6PM_MEAL", False)
        new_value = not bool(current_value)
        current_app.config["ENABLE_OT_AFTER_6PM_MEAL"] = new_value

        actor = session.get("username") or "admin"
        log_action(
            "app_config",
            "ENABLE_OT_AFTER_6PM_MEAL",
            "UPDATE",
            changed_by=actor,
            before_data={"enabled": current_value},
            after_data={"enabled": new_value},
            notes="Toggle OT-after-6pm meal feature via Users UI",
        )

        flash(f"Tính tiền ăn OT sau 18:00: {'Bật' if new_value else 'Tắt'}", "success")
        return redirect(url_for("users"))

    @app.route("/users/new", methods=["GET", "POST"])
    def create_user():
        blocked = _require_admin()
        if blocked:
            return blocked

        form_values = {
            "username": "",
            "full_name": "",
            "is_admin": False,
            "is_active": True,
        }

        if request.method == "POST":
            actor = session.get("username") or "admin"
            username = _normalize_username(request.form.get("username"))
            full_name = (request.form.get("full_name") or "").strip() or None
            password = request.form.get("password") or ""
            is_admin = request.form.get("is_admin") == "on"
            is_active = request.form.get("is_active") == "on"

            form_values.update(
                {
                    "username": username,
                    "full_name": full_name or "",
                    "is_admin": is_admin,
                    "is_active": is_active,
                }
            )

            if not username:
                flash("Cần nhập tên đăng nhập", "error")
                return render_template(
                    "user_form.html",
                    title="Thêm user",
                    form_title="Thêm user mới",
                    form_subtitle="Tạo tài khoản đăng nhập mới cho hệ thống.",
                    submit_label="Tạo user",
                    back_url=url_for("users"),
                    form_values=form_values,
                    is_edit=False,
                )

            if len(password) < 4:
                flash("Mật khẩu tối thiểu 4 ký tự", "error")
                return render_template(
                    "user_form.html",
                    title="Thêm user",
                    form_title="Thêm user mới",
                    form_subtitle="Tạo tài khoản đăng nhập mới cho hệ thống.",
                    submit_label="Tạo user",
                    back_url=url_for("users"),
                    form_values=form_values,
                    is_edit=False,
                )

            existing = AppUser.query.filter(func.lower(AppUser.username) == username).first()
            if existing:
                flash("Tên đăng nhập đã tồn tại", "error")
                return render_template(
                    "user_form.html",
                    title="Thêm user",
                    form_title="Thêm user mới",
                    form_subtitle="Tạo tài khoản đăng nhập mới cho hệ thống.",
                    submit_label="Tạo user",
                    back_url=url_for("users"),
                    form_values=form_values,
                    is_edit=False,
                )

            user = AppUser(
                username=username,
                full_name=full_name,
                password_hash=generate_password_hash(password),
                is_admin=is_admin,
                is_active=is_active,
            )
            db.session.add(user)
            db.session.flush()

            log_action(
                "app_users",
                str(user.id),
                "INSERT",
                changed_by=actor,
                after_data=_user_audit_payload(user),
            )
            db.session.commit()
            flash("Đã thêm user", "success")
            return redirect(url_for("users"))

        return render_template(
            "user_form.html",
            title="Thêm user",
            form_title="Thêm user mới",
            form_subtitle="Tạo tài khoản đăng nhập mới cho hệ thống.",
            submit_label="Tạo user",
            back_url=url_for("users"),
            form_values=form_values,
            is_edit=False,
        )

    @app.route("/users/<int:user_id>/edit", methods=["GET", "POST"])
    def edit_user(user_id):
        blocked = _require_admin()
        if blocked:
            return blocked

        user = AppUser.query.get_or_404(user_id)
        form_values = {
            "username": user.username,
            "full_name": user.full_name or "",
            "is_admin": bool(user.is_admin),
            "is_active": bool(user.is_active),
        }

        if request.method == "POST":
            actor = session.get("username") or "admin"
            username = _normalize_username(request.form.get("username"))
            full_name = (request.form.get("full_name") or "").strip() or None
            password = request.form.get("password") or ""
            is_admin = request.form.get("is_admin") == "on"
            is_active = request.form.get("is_active") == "on"

            form_values.update(
                {
                    "username": username,
                    "full_name": full_name or "",
                    "is_admin": is_admin,
                    "is_active": is_active,
                }
            )

            if not username:
                flash("Cần nhập tên đăng nhập", "error")
                return render_template(
                    "user_form.html",
                    title="Sửa user",
                    form_title="Cập nhật user",
                    form_subtitle="Cập nhật quyền và thông tin tài khoản.",
                    submit_label="Lưu cập nhật",
                    back_url=url_for("users"),
                    form_values=form_values,
                    is_edit=True,
                    user=user,
                )

            if password and len(password) < 4:
                flash("Mật khẩu mới tối thiểu 4 ký tự", "error")
                return render_template(
                    "user_form.html",
                    title="Sửa user",
                    form_title="Cập nhật user",
                    form_subtitle="Cập nhật quyền và thông tin tài khoản.",
                    submit_label="Lưu cập nhật",
                    back_url=url_for("users"),
                    form_values=form_values,
                    is_edit=True,
                    user=user,
                )

            duplicate = (
                AppUser.query.filter(func.lower(AppUser.username) == username, AppUser.id != user.id)
                .first()
            )
            if duplicate:
                flash("Tên đăng nhập đã tồn tại", "error")
                return render_template(
                    "user_form.html",
                    title="Sửa user",
                    form_title="Cập nhật user",
                    form_subtitle="Cập nhật quyền và thông tin tài khoản.",
                    submit_label="Lưu cập nhật",
                    back_url=url_for("users"),
                    form_values=form_values,
                    is_edit=True,
                    user=user,
                )

            current_user_id = session.get("user_id")
            is_last_active_admin = user.is_admin and user.is_active and _active_admin_count() <= 1

            if is_last_active_admin and (not is_admin or not is_active):
                flash("Không thể hạ quyền user admin cuối cùng", "error")
                return render_template(
                    "user_form.html",
                    title="Sửa user",
                    form_title="Cập nhật user",
                    form_subtitle="Cập nhật quyền và thông tin tài khoản.",
                    submit_label="Lưu cập nhật",
                    back_url=url_for("users"),
                    form_values=form_values,
                    is_edit=True,
                    user=user,
                )

            if user.id == current_user_id and (not is_admin or not is_active):
                flash("Không thể tự khóa hoặc bỏ quyền admin của tài khoản đang đăng nhập", "error")
                return render_template(
                    "user_form.html",
                    title="Sửa user",
                    form_title="Cập nhật user",
                    form_subtitle="Cập nhật quyền và thông tin tài khoản.",
                    submit_label="Lưu cập nhật",
                    back_url=url_for("users"),
                    form_values=form_values,
                    is_edit=True,
                    user=user,
                )

            before_data = _user_audit_payload(user)

            user.username = username
            user.full_name = full_name
            user.is_admin = is_admin
            user.is_active = is_active
            if password:
                user.password_hash = generate_password_hash(password)

            db.session.flush()
            log_action(
                "app_users",
                str(user.id),
                "UPDATE",
                changed_by=actor,
                before_data=before_data,
                after_data=_user_audit_payload(user),
            )
            db.session.commit()

            if user.id == current_user_id:
                _set_auth_session(user)

            flash("Đã cập nhật user", "success")
            return redirect(url_for("users"))

        return render_template(
            "user_form.html",
            title="Sửa user",
            form_title="Cập nhật user",
            form_subtitle="Cập nhật quyền và thông tin tài khoản.",
            submit_label="Lưu cập nhật",
            back_url=url_for("users"),
            form_values=form_values,
            is_edit=True,
            user=user,
        )

    @app.route("/users/<int:user_id>/delete", methods=["POST"])
    def delete_user(user_id):
        blocked = _require_admin()
        if blocked:
            return blocked

        user = AppUser.query.get_or_404(user_id)
        current_user_id = session.get("user_id")
        if user.id == current_user_id:
            flash("Không thể xóa tài khoản đang đăng nhập", "error")
            return redirect(url_for("users"))

        if user.is_admin and user.is_active and _active_admin_count() <= 1:
            flash("Không thể xóa user admin cuối cùng", "error")
            return redirect(url_for("users"))

        actor = session.get("username") or "admin"
        before_data = _user_audit_payload(user)
        db.session.delete(user)

        log_action(
            "app_users",
            str(user.id),
            "DELETE",
            changed_by=actor,
            before_data=before_data,
        )
        db.session.commit()
        flash("Đã xóa user", "success")
        return redirect(url_for("users"))
