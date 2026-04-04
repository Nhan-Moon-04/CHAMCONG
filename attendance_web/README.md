# Web Cham Cong (Flask + PostgreSQL)

Ung dung nay duoc dung de:
- Import file cham cong CSV/XLSX (bao gom file mau 17-31.csv).
- Quan ly bang ma ca lam, bang luong theo thang, lich lam, tang ca, ngay OFF/le.
- Tu dong tinh bang chi tiet cham cong theo quy tac da mo ta.
- Theo doi audit log cho moi thay doi.
- Backup PostgreSQL luc 17:00 moi ngay vao thu muc local (OneDrive co the dong bo thu muc nay).
- Tu dong xoa ban backup cu hon 30 ngay.

## 1) Chuan bi

1. Cai PostgreSQL 16+.
2. Tao database:
   - `attendance_db`
3. Tao file `.env` tu `.env.example` va cap nhat:
   - `DATABASE_URL`
   - `BACKUP_TARGET_DIR` (tro den thu muc trong OneDrive neu muon dong bo len cloud)

## 2) Chay tren Windows

1. Mo terminal tai thu muc `attendance_web`.
2. Tao virtual env:
   - `python -m venv .venv`
   - `.venv\\Scripts\\activate`
3. Cai thu vien:
   - `pip install -r requirements.txt`
4. Dat bien moi truong trong PowerShell (neu chay local PostgreSQL thi host la `localhost`):
   - `$env:DATABASE_URL="postgresql://attendance_user:123456@localhost:5432/attendance_db"`
   - `$env:BACKUP_TARGET_DIR="D:/OneDrive/attendance-backups"`
   - `$env:APP_TIMEZONE="Asia/Ho_Chi_Minh"`
5. Chay app:
   - `python run.py`
6. Mo trinh duyet:
   - `http://127.0.0.1:5000`

Luu y quan trong:
- PowerShell KHONG dung cu phap `KEY=value` nhu Linux; phai dung `$env:KEY="value"`.
- App da tu dong nhan ca `postgresql://...` va doi sang driver `psycopg`.
- Host `db` chi dung khi chay trong Docker Compose. Chay local thi dung `localhost`.

## 3) Chay bang Docker (Windows/Ubuntu)

1. Mo terminal tai thu muc `attendance_web`.
2. Chay:
   - `docker compose up -d --build`
3. Truy cap:
   - `http://127.0.0.1:5000`

## 4) Luong nghiep vu da co

1. `Bang ma ca lam`:
   - Da seed san cac ma: `X`, `XVP`, `N4`, `XT`, `X3`, `X4`, `N`, `S`, `C`, `P`, `L`.
   - Co cot `tien an` tren moi ma ca.
   - Co thao tac `Sua` mo ra trang rieng `/shifts/<id>/edit` de cap nhat de nhin.
   - Co thao tac `Xoa` tren tung dong ma ca.
   - Neu ma ca dang duoc dung trong lich lam hoac dang la ca mac dinh cua nhan vien thi he thong se chan xoa.
2. `Bang luong`:
   - Luu theo `month_key` (`YYYY-MM`) cho tung nhan vien.
   - Co `hinh thuc nhan tien`, `he so luong`.
3. `Bang lich lam`:
   - Lien ket nhan vien + ngay + ma ca.
   - Co `nghi 1-2 gio` (`absence_hours`).
   - Co tang ca tham chieu qua bang `overtime_entries`.
   - Co the import file Excel tren trang Schedules theo dinh dang viec.xlsx.
   - Header nhan vien: `ID <ma> - <ten>`, cot B la `Ngay`, o du lieu la ma ca.
   - Co tuy chon xoa lich cu cua thang truoc khi import.
4. `Ngay OFF`:
   - Chu nhat OFF mac dinh.
   - Ngay le luu trong bang holidays; neu khong xep ca khac thi tinh `L`.
5. `Phep nam`:
   - Tu dong cap nhat: `P = -1`, `S = -0.5`, `C = -0.5`.
6. `Bang chi tiet`:
   - Cot theo dung bo cot yeu cau (STT, Ma NV, Ho Ten, Ngay, Ten Ca, Gio Vao, Gio Ra, Gio Thuc, Chenh Lech, Tang Ca, Tong Gio, Status Code, So Gio, Luong Theo Ngay, Ghi Chu, Tien An Theo Ngay).
7. `Import xu ly sai thang`:
   - Co checkbox `Xoa du lieu cu cua thang truoc khi import`.
   - Dung khi import lai thang (vd: thang 03) de du lieu cu cua thang bi xoa truoc khi nap file moi.
   - Co nut `Xoa batch nay` trong lich su import de xoa du lieu cua lan import vua nap.

## 5) Backup OneDrive

- Dat `BACKUP_TARGET_DIR` den thu muc duoc OneDrive dong bo, vi du:
  - `D:/OneDrive/attendance-backups`
- Scheduler se backup luc 17:00 moi ngay.
- Sau moi lan backup, he thong xoa file `.dump` cu hon 30 ngay.

## 6) Audit log

- Trang `Audit log` ghi nhan insert/update/import/backup.
- Co nut `Xuat CSV audit` de trich xuat lich su thay doi.

## 7) Luu y

- App dung `pg_dump` cho backup, can dam bao lenh `pg_dump` co trong PATH.
- Neu import file tho khong dung encoding, app se tu fallback qua `cp1258` va `latin1`.

## 8) Loi thuong gap va cach sua

1. Loi PowerShell khong nhan `KEY=value`:
   - Dung cu phap:
   - `$env:DATABASE_URL="postgresql://attendance_user:123456@localhost:5432/attendance_db"`
   - `$env:BACKUP_TARGET_DIR="D:/OneDrive/attendance-backups"`

2. Loi `No module named 'psycopg2'`:
   - App da tu dong chuyen `postgresql://` sang `postgresql+psycopg://`.
   - Chi can giu URL PostgreSQL dung va cai requirements la du.

3. Loi `permission denied for schema public`:
   - Dang nhap PostgreSQL bang tai khoan admin (`postgres`) va cap quyen:
   - `ALTER DATABASE attendance_db OWNER TO attendance_user;`
   - `GRANT ALL PRIVILEGES ON DATABASE attendance_db TO attendance_user;`
   - `GRANT USAGE, CREATE ON SCHEMA public TO attendance_user;`
   - `ALTER SCHEMA public OWNER TO attendance_user;`

4. Host `db` hay `localhost`:
   - Chay app tren may host: dung `localhost`.
   - Chay app trong Docker Compose: dung `db`.
