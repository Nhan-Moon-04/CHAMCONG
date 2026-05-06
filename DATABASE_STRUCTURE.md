# Cấu Trúc Cơ Sở Dữ Liệu - Hệ Thống Chấm Công

## 📊 Tổng Quan Hệ Thống

Hệ thống gồm **2 phần chính**:
1. **VBA (Excel)** - Xử lý dữ liệu cũ từ các file .xls, .bas
2. **Python (Flask + PostgreSQL)** - Ứng dụng web đầy đủ với cơ sở dữ liệu

---

## 🔴 PHẦN VBA (Excel) - Dữ Liệu Cũ

### File VBA trong workspace:
```
📁 d:\CODE\CHAMCONG\
  ├─ CHAM CONG DONE.bas           [VBA - Xử lý chấm công]
  ├─ ModChamCong (2).bas          [VBA - Module chấm công cũ]
  ├─ SHEET CHÂM CÔNG DONE.bas     [VBA - Sheet định dạng cũ]
  │
  └─ 📊 File dữ liệu Excel:
     ├─ CONG DOAN 2026.xls        [Dữ liệu đoàn công đoàn]
     ├─ TIEN COM THANG 2 2026 -.xls [Dữ liệu tiền cơm tháng 2/2026]
     └─ draf.xlsx                  [Bản nháp]
```

**Chức năng VBA:**
- ❌ KHÔNG còn sử dụng trong hệ thống web mới
- ✓ Chỉ dùng để tham khảo công thức cũ hoặc import dữ liệu lịch sử
- ✓ Có thể import dữ liệu từ các file .xls cũ vào Python

---

## 🔵 PHẦN PYTHON (Flask + PostgreSQL) - Hệ Thống Mới

### Cấu trúc thư mục:
```
📁 attendance_web/
  ├─ 📄 run.py                [Entry point - Khởi động Flask]
  ├─ 📄 requirements.txt       [Dependencies Python]
  ├─ 📄 docker-compose.yml     [Docker config]
  │
  ├─ 📁 app/                  [ỨNG DỤNG FLASK CHÍNH]
  │  ├─ __init__.py
  │  ├─ config.py             [Cấu hình ứng dụng]
  │  ├─ database.py           [Khởi tạo SQLAlchemy]
  │  ├─ models.py             [ĐỊNH NGHĨA TẤT CẢ TABLE]
  │  ├─ routes.py             [API/Endpoint chính]
  │  ├─ union_routes.py        [API đoàn công đoàn]
  │  │
  │  ├─ 📁 services/          [BỎ BIỆT LÝ TRÍ KINH DOANH]
  │  │  ├─ attendance.py       [🔵 Xử lý chấm công]
  │  │  ├─ salary_calculator.py [🔵 Tính lương]
  │  │  ├─ backup.py           [🔵 Backup DB]
  │  │  ├─ importer.py         [🔵 Import dữ liệu]
  │  │  ├─ schedule_importer.py [🔵 Import lịch làm việc]
  │  │  ├─ salary_importer.py   [🔵 Import lương]
  │  │  ├─ salary_meal_export.py [🔵 Export tiền cơm]
  │  │  ├─ audit.py            [🔵 Ghi log thay đổi]
  │  │  ├─ users.py            [🔵 Quản lý user]
  │  │  ├─ nu_shift.py         [🔵 Xử lý ca nữ]
  │  │  └─ salary_overview_export.py [🔵 Export tổng lương]
  │  │
  │  ├─ 📁 static/            [CSS, JS, Image]
  │  │  └─ css/
  │  │     └─ style.css
  │  │
  │  └─ 📁 templates/         [HTML Jinja2]
  │     ├─ base.html          [Layout cơ bản]
  │     ├─ login.html
  │     ├─ dashboard.html
  │     ├─ employees.html
  │     ├─ shifts.html
  │     ├─ schedules.html
  │     ├─ salaries.html
  │     ├─ salary_meal_period.html
  │     ├─ salary_overview.html
  │     ├─ holidays.html
  │     ├─ audit.html
  │     ├─ advances.html
  │     ├─ union_*.html       [Union/đoàn công đoàn]
  │     └─ imports.html
  │
  ├─ 📁 tests/               [UNIT TEST]
  │  ├─ test_salary_calculator.py
  │  ├─ test_meal_count.py
  │  ├─ repro_salary_overview.py
  │  └─ repro_employee_detail.py
  │
  ├─ 📁 uploads/            [Folder upload files]
  │  └─ attendance_imports/  [File import chấm công]
  │
  ├─ 📁 backups/            [Backup database]
  ├─ 📁 instance/           [Flask instance files]
  └─ 📁 tests/
```

---

## 📋 CÁC TABLE TRONG CƠ SỞ DỮ LIỆU (PostgreSQL)

### 👥 1. QUẢN LÝ NHÂN VIÊN

#### `app_users` - Tài khoản người dùng
```sql
CREATE TABLE app_users (
  id SERIAL PRIMARY KEY,
  username VARCHAR(64) UNIQUE NOT NULL,
  password_hash VARCHAR(255) NOT NULL,
  full_name VARCHAR(120),
  is_admin BOOLEAN DEFAULT FALSE,
  is_active BOOLEAN DEFAULT TRUE,
  created_at DATETIME,
  updated_at DATETIME
);
```
**🔵 Xử lý bởi Python:** 
- `routes.py` - Login, quản lý tài khoản
- `services/users.py` - CRUD user

---

#### `employees` - Danh sách nhân viên
```sql
CREATE TABLE employees (
  id SERIAL PRIMARY KEY,
  employee_code VARCHAR(32) UNIQUE NOT NULL,
  full_name VARCHAR(120) NOT NULL,
  gender VARCHAR(16),
  hometown VARCHAR(120),
  birth_year INTEGER,
  default_shift_code VARCHAR(16) FOREIGN KEY,
  is_active BOOLEAN DEFAULT TRUE,
  created_at DATETIME,
  updated_at DATETIME
);
```
**🔵 Xử lý bởi Python:**
- `routes.py` - CRUD nhân viên
- `services/importer.py` - Import từ file .xls (từ VBA)

---

### 📅 2. QUẢN LÝ CA LÀM VIỆC

#### `shift_templates` - Khuôn mẫu ca làm việc
```sql
CREATE TABLE shift_templates (
  id SERIAL PRIMARY KEY,
  code VARCHAR(16) UNIQUE NOT NULL,
  name VARCHAR(120) NOT NULL,
  start_time TIME,
  end_time TIME,
  break_minutes INTEGER,
  standard_hours DECIMAL(5,2) DEFAULT 8,
  default_overtime_hours DECIMAL(5,2),
  meal_allowance DECIMAL(12,2),
  meal_count INTEGER,
  is_night_shift BOOLEAN,
  is_leave_code BOOLEAN,
  is_paid_leave BOOLEAN,
  notes TEXT,
  created_at DATETIME,
  updated_at DATETIME
);
```
**🔵 Xử lý bởi Python:**
- `routes.py` - CRUD shift template
- `services/salary_calculator.py` - Tính lương theo ca
- `services/nu_shift.py` - Xử lý ca nữ đặc biệt

---

#### `work_schedules` - Lịch làm việc cụ thể
```sql
CREATE TABLE work_schedules (
  id SERIAL PRIMARY KEY,
  employee_id INTEGER FOREIGN KEY NOT NULL,
  work_date DATE NOT NULL,
  month_key VARCHAR(7) NOT NULL,
  shift_id INTEGER FOREIGN KEY NOT NULL,
  absence_hours DECIMAL(5,2),
  notes VARCHAR(255),
  created_at DATETIME,
  updated_at DATETIME,
  UNIQUE(employee_id, work_date)
);
```
**🔵 Xử lý bởi Python:**
- `routes.py` - Xem/sửa lịch
- `services/schedule_importer.py` - Import lịch từ file

---

### ⏰ 3. QUẢN LÝ CHẤM CÔNG

#### `attendance_logs` - Raw log chấm công (tất cả check-in/out)
```sql
CREATE TABLE attendance_logs (
  id SERIAL PRIMARY KEY,
  employee_code VARCHAR(32) NOT NULL,
  employee_name VARCHAR(120),
  department VARCHAR(120),
  event_time DATETIME NOT NULL,
  source_file VARCHAR(255),
  import_batch VARCHAR(36),
  created_at DATETIME,
  updated_at DATETIME
);
```
**🔵 Xử lý bởi Python:**
- `services/importer.py` - Import từ file chấm công (từ VBA hoặc thiết bị)
- `routes.py` - Xem raw logs

---

#### `attendance_daily` - Tổng hợp chấm công theo ngày
```sql
CREATE TABLE attendance_daily (
  id SERIAL PRIMARY KEY,
  employee_id INTEGER FOREIGN KEY NOT NULL,
  work_date DATE NOT NULL,
  first_check_in DATETIME,
  last_check_out DATETIME,
  total_hours DECIMAL(6,2),
  import_batch VARCHAR(36),
  created_at DATETIME,
  updated_at DATETIME,
  UNIQUE(employee_id, work_date)
);
```
**🔵 Xử lý bởi Python:**
- `services/attendance.py` - Tính từ attendance_logs
- `routes.py` - Xem dữ liệu daily

---

#### `attendance_details` - Chi tiết chấm công + tính công (QUAN TRỌNG)
```sql
CREATE TABLE attendance_details (
  id SERIAL PRIMARY KEY,
  employee_id INTEGER FOREIGN KEY NOT NULL,
  work_date DATE NOT NULL,
  month_key VARCHAR(7) NOT NULL,
  shift_code VARCHAR(16),
  shift_name VARCHAR(120),
  check_in DATETIME,
  check_out DATETIME,
  standard_hours DECIMAL(6,2),
  actual_work_hours DECIMAL(6,2),
  deviation_hours DECIMAL(6,2),
  overtime_hours DECIMAL(6,2),
  total_span_hours DECIMAL(6,2),
  status_code VARCHAR(16),        -- FULL, HALF, LATE, ABSENT...
  paid_hours DECIMAL(6,2),
  daily_wage DECIMAL(12,2),
  meal_allowance_daily DECIMAL(12,2),
  notes VARCHAR(255),
  created_at DATETIME,
  updated_at DATETIME,
  UNIQUE(employee_id, work_date)
);
```
**🔵 Xử lý bởi Python:**
- `services/attendance.py` - Tính chi tiết công từ shifts
- `routes.py` - API /details

---

### 💰 4. QUẢN LÝ LƯƠNG

#### `monthly_salaries` - Mức lương tháng
```sql
CREATE TABLE monthly_salaries (
  id SERIAL PRIMARY KEY,
  employee_id INTEGER FOREIGN KEY NOT NULL,
  month_key VARCHAR(7) NOT NULL,
  base_daily_wage DECIMAL(12,2),
  pay_method VARCHAR(32),         -- cash, bank...
  salary_coefficient DECIMAL(10,4) DEFAULT 1.0,
  created_at DATETIME,
  updated_at DATETIME,
  UNIQUE(employee_id, month_key)
);
```
**🔵 Xử lý bởi Python:**
- `services/salary_importer.py` - Import lương từ file
- `routes.py` - CRUD lương
- `services/salary_calculator.py` - Tính lương

---

#### `payroll_payment_statuses` - Trạng thái thanh toán
```sql
CREATE TABLE payroll_payment_statuses (
  id SERIAL PRIMARY KEY,
  employee_id INTEGER FOREIGN KEY NOT NULL,
  month_key VARCHAR(7) NOT NULL,
  salary_received BOOLEAN,         -- Nhận lương chưa
  meal_period_1_received BOOLEAN,  -- Nhận tiền cơm kỳ 1
  meal_period_2_received BOOLEAN,  -- Nhận tiền cơm kỳ 2
  updated_by VARCHAR(64),
  created_at DATETIME,
  updated_at DATETIME,
  UNIQUE(employee_id, month_key)
);
```
**🔵 Xử lý bởi Python:**
- `routes.py` - Cập nhật trạng thái
- `services/salary_meal_export.py` - Export tiền cơm

---

#### `advance_payments` - Tạm ứng
```sql
CREATE TABLE advance_payments (
  id SERIAL PRIMARY KEY,
  employee_id INTEGER FOREIGN KEY NOT NULL,
  advance_date DATE NOT NULL,
  month_key VARCHAR(7) NOT NULL,
  amount DECIMAL(12,2),
  input_mode VARCHAR(16),          -- amount, days
  payment_method VARCHAR(32),      -- cash, bank
  advance_days DECIMAL(6,2),
  notes VARCHAR(255),
  created_at DATETIME,
  updated_at DATETIME
);
```
**🔵 Xử lý bởi Python:**
- `routes.py` - CRUD tạm ứng

---

#### `monthly_workday_configs` - Config ngày làm việc tháng
```sql
CREATE TABLE monthly_workday_configs (
  id SERIAL PRIMARY KEY,
  month_key VARCHAR(7) UNIQUE NOT NULL,
  company_work_days DECIMAL(6,2) DEFAULT 26,
  notes VARCHAR(255),
  created_at DATETIME,
  updated_at DATETIME
);
```
**🔵 Xử lý bởi Python:**
- `routes.py` - Cài đặt ngày làm việc

---

### 📈 5. QUẢN LÝ OT VÀ PHÉP

#### `overtime_entries` - Chi tiết OT
```sql
CREATE TABLE overtime_entries (
  id SERIAL PRIMARY KEY,
  schedule_id INTEGER FOREIGN KEY UNIQUE NOT NULL,
  hours DECIMAL(5,2),
  reason VARCHAR(255),
  created_at DATETIME,
  updated_at DATETIME
);
```
**🔵 Xử lý bởi Python:**
- `services/salary_calculator.py` - Tính OT

---

#### `leave_balances` - Số ngày phép còn lại
```sql
CREATE TABLE leave_balances (
  id SERIAL PRIMARY KEY,
  employee_id INTEGER FOREIGN KEY NOT NULL,
  year INTEGER NOT NULL,
  total_days DECIMAL(5,2) DEFAULT 12,
  used_days DECIMAL(5,2) DEFAULT 0,
  created_at DATETIME,
  updated_at DATETIME,
  UNIQUE(employee_id, year)
);
```
**🔵 Xử lý bởi Python:**
- `routes.py` - CRUD phép

---

#### `payroll_leave_snapshots` - Snapshot phép năm theo tháng
```sql
CREATE TABLE payroll_leave_snapshots (
  id SERIAL PRIMARY KEY,
  employee_id INTEGER FOREIGN KEY NOT NULL,
  month_key VARCHAR(7) NOT NULL,
  year INTEGER NOT NULL,
  entitled_days DECIMAL(6,2),
  bonus_entitled_days DECIMAL(6,2),
  used_days DECIMAL(6,2),
  remaining_days DECIMAL(6,2),
  work_days DECIMAL(6,2),
  sick_leave_days DECIMAL(6,2),
  monthly_breakdown JSON,
  source_file VARCHAR(255),
  created_at DATETIME,
  updated_at DATETIME,
  UNIQUE(employee_id, month_key)
);
```
**🔵 Xử lý bởi Python:**
- `services/salary_importer.py` - Import sheet `phep nam`

---

#### `payroll_slips` - Phiếu lương chi tiết từng người theo tháng
```sql
CREATE TABLE payroll_slips (
  id SERIAL PRIMARY KEY,
  employee_id INTEGER FOREIGN KEY NOT NULL,
  month_key VARCHAR(7) NOT NULL,
  payroll_group VARCHAR(32),
  attendance_days DECIMAL(6,2),
  leave_used_days DECIMAL(6,2),
  leave_remaining_days DECIMAL(6,2),
  salary_by_attendance DECIMAL(12,2),
  overtime_weekday_hours DECIMAL(6,2),
  overtime_sunday_hours DECIMAL(6,2),
  overtime_pay DECIMAL(12,2),
  role_allowance DECIMAL(12,2),
  child_allowance DECIMAL(12,2),
  transport_phone_allowance DECIMAL(12,2),
  meal_allowance DECIMAL(12,2),
  attendance_allowance DECIMAL(12,2),
  gross_total DECIMAL(12,2),
  social_insurance_deduction DECIMAL(12,2),
  union_fee_deduction DECIMAL(12,2),
  pit_tax_deduction DECIMAL(12,2),
  advance_deduction DECIMAL(12,2),
  net_income DECIMAL(12,2),
  source_file VARCHAR(255),
  extra_data JSON,
  created_at DATETIME,
  updated_at DATETIME,
  UNIQUE(employee_id, month_key)
);
```
**🔵 Xử lý bởi Python:**
- `services/salary_importer.py` - Import sheet `PHIẾU LƯƠNG`

---

#### `payroll_insurance_contributions` - BHXH/BHYT/BHTN theo tháng
```sql
CREATE TABLE payroll_insurance_contributions (
  id SERIAL PRIMARY KEY,
  employee_id INTEGER FOREIGN KEY NOT NULL,
  month_key VARCHAR(7) NOT NULL,
  insured_salary DECIMAL(12,2),
  employer_bhxh DECIMAL(12,2),
  employee_bhxh DECIMAL(12,2),
  employer_bhyt DECIMAL(12,2),
  employee_bhyt DECIMAL(12,2),
  employer_bhtn DECIMAL(12,2),
  employee_bhtn DECIMAL(12,2),
  employer_accident DECIMAL(12,2),
  employer_total DECIMAL(12,2),
  employee_total DECIMAL(12,2),
  union_fund DECIMAL(12,2),
  source_file VARCHAR(255),
  created_at DATETIME,
  updated_at DATETIME,
  UNIQUE(employee_id, month_key)
);
```
**🔵 Xử lý bởi Python:**
- `services/salary_importer.py` - Import sheet `BAO HIEM`

---

#### `payroll_tax_contributions` - Thuế TNCN theo tháng
```sql
CREATE TABLE payroll_tax_contributions (
  id SERIAL PRIMARY KEY,
  employee_id INTEGER FOREIGN KEY NOT NULL,
  month_key VARCHAR(7) NOT NULL,
  taxable_income DECIMAL(12,2),
  pit_tax DECIMAL(12,2),
  self_deduction DECIMAL(12,2),
  dependent_deduction DECIMAL(12,2),
  insurance_deduction DECIMAL(12,2),
  total_deduction DECIMAL(12,2),
  source_file VARCHAR(255),
  notes VARCHAR(255),
  created_at DATETIME,
  updated_at DATETIME,
  UNIQUE(employee_id, month_key)
);
```
**🔵 Xử lý bởi Python:**
- `services/salary_importer.py` - Tách từ sheet `PHIẾU LƯƠNG` (và mở rộng cho `CT THUE TNCN` khi có dữ liệu)

---

#### `holidays` - Ngày lễ
```sql
CREATE TABLE holidays (
  id SERIAL PRIMARY KEY,
  holiday_date DATE UNIQUE NOT NULL,
  name VARCHAR(120) NOT NULL,
  is_paid BOOLEAN DEFAULT TRUE,
  notes VARCHAR(255),
  created_at DATETIME,
  updated_at DATETIME
);
```
**🔵 Xử lý bởi Python:**
- `routes.py` - CRUD ngày lễ

---

### 📊 6. ĐOÀN CÔNG ĐOÀN (UNION)

#### Dữ liệu Union (từ file .xls)
- Import từ: `CONG DOAN 2026.xls` (VBA -> Python)
- **🔵 Xử lý bởi Python:** `union_routes.py`

---

### 📝 7. AUDIT & LOG

#### `audit_logs` - Ghi nhật ký thay đổi
```sql
CREATE TABLE audit_logs (
  id SERIAL PRIMARY KEY,
  table_name VARCHAR(64) NOT NULL,
  record_id VARCHAR(64) NOT NULL,
  action VARCHAR(32) NOT NULL,      -- CREATE, UPDATE, DELETE
  changed_by VARCHAR(64),           -- User thay đổi
  changed_at DATETIME,
  before_data JSON,                 -- Dữ liệu trước
  after_data JSON,                  -- Dữ liệu sau
  notes VARCHAR(255),
  created_at DATETIME
);
```
**🔵 Xử lý bởi Python:**
- `services/audit.py` - Ghi log tất cả thay đổi
- `routes.py` - Xem audit logs

---

## 🔄 LUỒNG DỮ LIỆU: VBA → Python

```
VBA/Excel (Nguồn cũ)
      ↓
   .xls files
      ↓
Python importer.py
      ↓
PostgreSQL Database
      ↓
Flask Web App
      ↓
HTML Templates
```

### Import Flow:
1. **VBA để dữ liệu:** `CONG DOAN 2026.xls`, `TIEN COM THANG 2 2026 -.xls`, etc.
2. **Python đọc và transform:** `services/importer.py`, `services/salary_importer.py`
3. **Lưu vào PostgreSQL:** ORM SQLAlchemy
4. **Hiển thị trên web:** Flask routes + HTML templates

---

## 🎯 TÓM TẮT XỬ LÝ VBA vs PYTHON

| Chức năng | VBA (Cũ) | Python (Mới) |
|-----------|----------|------------|
| **Quản lý nhân viên** | ❌ | 🔵 routes.py + services/users.py |
| **Quản lý ca shift** | ❌ | 🔵 routes.py + services/nu_shift.py |
| **Chấm công** | ❌ Xử lý cũ | 🔵 services/attendance.py |
| **Lương** | ❌ Công thức cũ | 🔵 services/salary_calculator.py |
| **Import dữ liệu** | ✓ Tạo .xls | 🔵 services/importer.py (đọc .xls) |
| **Tiền cơm** | ✓ File .xls | 🔵 services/salary_meal_export.py |
| **Đoàn công đoàn** | ✓ File .xls | 🔵 union_routes.py + union_*.html |
| **Backup DB** | ❌ | 🔵 services/backup.py |
| **Web UI** | ❌ | 🔵 Flask + Jinja2 + CSS |
| **Audit Log** | ❌ | 🔵 services/audit.py |

---

## 📌 NOTES

- **PostgreSQL** được chọn thay vì SQLite để scale tốt hơn
- **ORM SQLAlchemy** quản lý schema + relationships
- **Naming Convention** tự động (fk_, ix_, uq_, etc.)
- **Timestampmixin** tự động thêm created_at, updated_at
- **Import từ VBA:** Dữ liệu cũ được chuyển qua Python importer
- **Audit Trail:** Mọi thay đổi được ghi lại trong audit_logs

