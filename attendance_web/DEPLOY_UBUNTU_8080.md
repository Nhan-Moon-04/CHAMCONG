# Deploy Ubuntu 24 - Port 8080 - Auto Start

Huong dan nay dung cho source:
- https://github.com/Nhan-Moon-04/CHAMCONG

## 1) Cai Docker va Git

Chay tren Ubuntu 24:

```bash
sudo apt update
sudo apt install -y docker.io docker-compose-v2 git
sudo systemctl enable --now docker
sudo usermod -aG docker $USER
newgrp docker
```

Kiem tra:

```bash
docker --version
docker compose version
```

## 2) Clone source

```bash
sudo mkdir -p /opt
cd /opt
sudo git clone https://github.com/Nhan-Moon-04/CHAMCONG.git
sudo chown -R $USER:$USER /opt/CHAMCONG
cd /opt/CHAMCONG/attendance_web
```

## 3) Tao file env

```bash
cp .env.example .env
nano .env
```

Cap nhat cac bien quan trong:
- SECRET_KEY=0989057191
- LOGIN_USERNAME=admin
- LOGIN_PASSWORD=123456
- APP_TIMEZONE=Asia/Ho_Chi_Minh
- BACKUP_TARGET_DIR=/app/backups

Neu ban muon doi tai khoan DB thi sua them DATABASE_URL.

## 3.1) PostgreSQL o dau?

Ban KHONG can cai PostgreSQL tren host Ubuntu khi deploy theo cach nay.

PostgreSQL da nam trong docker-compose voi service `db` (image `postgres:16`).
Khi chay `docker compose up -d --build`, he thong se len ca:
- container DB: attendance_db
- container Web: attendance_web

Kiem tra DB container:

```bash
docker compose ps db
docker compose logs -f db
```

## 4) Chay web o port 8080

Trong file docker-compose.yml, sua mapping port tu:
- "5000:5000"
thanh:
- "8080:5000"

Lenh sua nhanh:

```bash
sed -i 's/"5000:5000"/"8080:5000"/' docker-compose.yml
```

## 5) Build va run

```bash
docker compose up -d --build
```

Lenh tren se build web image va chay dong thoi ca `db` + `web`.

Kiem tra trang thai:

```bash
docker compose ps
docker compose logs -f web
```

Test local tren server:

```bash
curl -I http://127.0.0.1:8080/login
```

Truy cap tu may ngoai:
- http://IP_SERVER:8080/login

## 6) Mo firewall port 8080 (neu dung UFW)

```bash
sudo ufw allow 8080/tcp
sudo ufw reload
sudo ufw status
```

## 7) Auto start sau reboot

Compose file da co:
- restart: unless-stopped

Chi can dam bao docker service auto start:

```bash
sudo systemctl enable docker
sudo systemctl status docker
```

Sau khi reboot, containers se tu len lai.

## 8) Update code khi da deploy

```bash
cd /opt/CHAMCONG
git pull
cd attendance_web
docker compose up -d --build
```

## 9) Lenh quan ly nhanh

Dung app:

```bash
docker compose stop
```

Chay lai:

```bash
docker compose start
```

Khoi dong lai:

```bash
docker compose restart
```

Xem log web:

```bash
docker compose logs -f web
```

## 10) Luu y backup

App co scheduler backup luc 17:00 moi ngay. Backup duoc ghi vao thu muc:
- attendance_web/backups

Neu log bao loi "pg_dump: not found" thi can them postgresql-client vao Dockerfile web image.
