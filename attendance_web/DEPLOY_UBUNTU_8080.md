# REBUILD TU DAU UBUNTU 24

Muc tieu:
- Xoa sach toan bo repo CHAMCONG cu.
- Xoa sach Docker cu (containers, images, volumes, cache, package docker).
- Cai lai tu dau va deploy lai du an CHAMCONG o port 8080.

Source:
- https://github.com/Nhan-Moon-04/CHAMCONG

## 0) Canh bao

Huong dan nay co buoc XOA SACH DU LIEU Docker va source code cu.

Neu can giu file backup cua app, copy ra ngoai truoc khi xoa.

## 1) Backup tuy chon truoc khi wipe

```bash
sudo mkdir -p /root/chamcong_backup
if [ -d /opt/CHAMCONG/attendance_web/backups ]; then sudo cp -a /opt/CHAMCONG/attendance_web/backups /root/chamcong_backup/; fi
ls -la /root/chamcong_backup
```

## 2) Stop stack cu neu con ton tai

```bash
if [ -f /opt/CHAMCONG/attendance_web/docker-compose.yml ]; then cd /opt/CHAMCONG/attendance_web; sudo docker compose down --remove-orphans; fi
sudo docker rm -f attendance_web attendance_db 2>/dev/null || true
sudo docker volume rm attendance_web_postgres_data 2>/dev/null || true
sudo docker network rm attendance_web_default 2>/dev/null || true
```

## 3) Xoa repo CHAMCONG cu

```bash
sudo rm -rf /opt/CHAMCONG
```

## 4) Go Docker cu hoan toan (wipe)

```bash
sudo systemctl stop docker docker.socket 2>/dev/null || true
sudo apt purge -y docker.io docker-compose-v2 docker-buildx-plugin containerd runc
sudo apt autoremove -y --purge
sudo rm -rf /var/lib/docker /var/lib/containerd /etc/docker
```

Neu muon xoa sach hon nua (toan bo cache apt):

```bash
sudo apt clean
```

## 5) Cai lai Docker va Git

```bash
sudo apt update
sudo apt install -y docker.io docker-compose-v2 git
sudo systemctl enable --now docker
sudo usermod -aG docker $USER
newgrp docker
docker --version
docker compose version
```

## 6) Clone source lai tu dau

```bash
sudo mkdir -p /opt
cd /opt
sudo git clone https://github.com/Nhan-Moon-04/CHAMCONG.git
sudo chown -R $USER:$USER /opt/CHAMCONG
cd /opt/CHAMCONG/attendance_web
```

## 7) Tao file .env va dat bien dung

```bash
    cp .env.example .env
nano .env
```

Dat it nhat cac bien sau (quan trong nhat la password phai dong bo):

```dotenv
SECRET_KEY=0989057191
APP_NAME=HIEP LOI
WEB_PORT=8080
LOGIN_USERNAME=admin
LOGIN_PASSWORD=123456

POSTGRES_DB=attendance_db
POSTGRES_USER=postgres
POSTGRES_PASSWORD=postgres
DATABASE_URL=postgresql+psycopg://postgres:postgres@db:5432/attendance_db

BACKUP_TARGET_DIR=/app/backups
BACKUP_RETENTION_DAYS=30
ENABLE_BACKUP_SCHEDULER=1
APP_TIMEZONE=Asia/Ho_Chi_Minh
```

Luu y:
- POSTGRES_PASSWORD trong .env phai giong password trong DATABASE_URL.

## 8) Chay app o port 8080

Khong sua file docker-compose.yml truc tiep. Chi can dat trong `.env`:
- `WEB_PORT=8080`

Kiem tra compose da nhan dung port:

```bash
docker compose config | grep -n "8080:5000"
```

## 9) Build va run lai hoan toan

```bash
docker compose up -d --build
docker compose ps
docker compose logs -f db
docker compose logs -f web
```

Test nhanh tren server:

```bash
curl -I http://127.0.0.1:8080/login
```

## 10) Mo firewall port 8080 (neu dung UFW)

```bash
sudo ufw allow 8080/tcp
sudo ufw reload
sudo ufw status
```

## 11) Kiem tra loi password DB neu web restart loop

Kiem tra bien dang su dung thuc te:

```bash
docker compose config | grep -E "POSTGRES_DB|POSTGRES_USER|POSTGRES_PASSWORD|DATABASE_URL"
```

Neu van bao password authentication failed, doi password user postgres trong DB cho khop:

```bash
docker compose exec -u postgres db psql -d postgres -c "ALTER USER postgres WITH PASSWORD 'postgres';"
docker compose restart web
docker compose logs -f web
```

## 12) Lenh quan ly nhanh sau khi deploy

```bash
docker compose stop
docker compose start
docker compose restart
docker compose logs -f web
docker compose logs -f db
```

## 13) Auto start sau reboot

Compose service da co restart unless-stopped.

Chi can dam bao docker daemon auto start:

```bash
sudo systemctl enable docker
sudo systemctl status docker
```
