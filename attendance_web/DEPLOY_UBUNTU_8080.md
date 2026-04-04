# Deploy attendance_web len Ubuntu (port 8080)

Tai lieu nay huong dan deploy ban production cho project Flask + PostgreSQL tren Ubuntu, expose app o cong `8080`.

## 1) Yeu cau

- Ubuntu 22.04 hoac 24.04
- Tai khoan co quyen sudo
- Da copy source code `attendance_web` len server (vi du: `/opt/attendance_web`)

## 2) Cai package he thong

```bash
sudo apt update
sudo apt install -y python3 python3-venv python3-pip postgresql postgresql-contrib libpq-dev
```

## 3) Tao database PostgreSQL

```bash
sudo -u postgres psql
```

Trong psql:

```sql
CREATE USER attendance_user WITH PASSWORD 'StrongPass123!';
CREATE DATABASE attendance_db OWNER attendance_user;
GRANT ALL PRIVILEGES ON DATABASE attendance_db TO attendance_user;
\c attendance_db
GRANT USAGE, CREATE ON SCHEMA public TO attendance_user;
ALTER SCHEMA public OWNER TO attendance_user;
\q
```

## 4) Chuan bi source code

```bash
sudo mkdir -p /opt/attendance_web
sudo chown -R $USER:$USER /opt/attendance_web
# copy source vao /opt/attendance_web (git clone hoac rsync)
cd /opt/attendance_web
```

Tao venv va cai thu vien:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
pip install gunicorn
```

## 5) Tao file .env

Tao file `/opt/attendance_web/.env`:

```env
SECRET_KEY=change-me-to-a-long-random-string
APP_NAME=HIEP LOI Workforce
LOGIN_USERNAME=admin
LOGIN_PASSWORD=123456

DATABASE_URL=postgresql://attendance_user:StrongPass123!@127.0.0.1:5432/attendance_db
BACKUP_TARGET_DIR=/opt/attendance_web/backups
BACKUP_RETENTION_DAYS=30
ENABLE_BACKUP_SCHEDULER=1
APP_TIMEZONE=Asia/Ho_Chi_Minh
```

Tao thu muc backup:

```bash
mkdir -p /opt/attendance_web/backups
```

## 6) Chay thu app bang gunicorn o cong 8080

```bash
cd /opt/attendance_web
source .venv/bin/activate
gunicorn --bind 0.0.0.0:8080 --workers 3 --timeout 120 run:app
```

Test nhanh:

```bash
curl -I http://127.0.0.1:8080
```

Neu OK, nhan `Ctrl+C` de dung va tao service.

## 7) Tao systemd service (tu chay cung may)

Tao user he thong (1 lan):

```bash
sudo useradd --system --create-home --shell /bin/bash attendance
sudo chown -R attendance:attendance /opt/attendance_web
```

Tao file `/etc/systemd/system/attendance-web.service`:

```ini
[Unit]
Description=Attendance Web (Flask + Gunicorn)
After=network.target postgresql.service

[Service]
User=attendance
Group=attendance
WorkingDirectory=/opt/attendance_web
EnvironmentFile=/opt/attendance_web/.env
ExecStart=/opt/attendance_web/.venv/bin/gunicorn --bind 0.0.0.0:8080 --workers 3 --timeout 120 run:app
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
```

Reload va start:

```bash
sudo systemctl daemon-reload
sudo systemctl enable attendance-web
sudo systemctl restart attendance-web
sudo systemctl status attendance-web --no-pager
```

## 8) Mo firewall cong 8080

```bash
sudo ufw allow 8080/tcp
sudo ufw reload
sudo ufw status
```

## 9) Truy cap he thong

- Trong LAN: `http://<server-ip>:8080`
- Vi du: `http://192.168.1.158:8080`

Dang nhap bang tai khoan trong `.env` (`LOGIN_USERNAME` / `LOGIN_PASSWORD`).

## 10) Lenh van hanh thuong dung

Xem log realtime:

```bash
sudo journalctl -u attendance-web -f
```

Restart service:

```bash
sudo systemctl restart attendance-web
```

Kiem tra cong dang lang nghe:

```bash
ss -lntp | grep 8080
```

## 11) Cap nhat code (quy trinh nhanh)

```bash
cd /opt/attendance_web
sudo -u attendance -H bash -lc 'cd /opt/attendance_web && source .venv/bin/activate && pip install -r requirements.txt'
sudo systemctl restart attendance-web
sudo systemctl status attendance-web --no-pager
```

## 12) Loi thuong gap

1. Service crash do sai ENV
- Kiem tra file `.env` co du `DATABASE_URL`, `SECRET_KEY`, `BACKUP_TARGET_DIR`.
- Xem log: `sudo journalctl -u attendance-web -n 200 --no-pager`

2. Khong ket noi duoc PostgreSQL
- Kiem tra DB dang chay: `sudo systemctl status postgresql`
- Thu ket noi bang user app:
  ```bash
  psql "postgresql://attendance_user:StrongPass123!@127.0.0.1:5432/attendance_db" -c "SELECT 1;"
  ```

3. Vao duoc localhost nhung may khac khong vao duoc
- Kiem tra UFW da mo `8080/tcp`
- Kiem tra mang LAN/router co chan cong hay khong
