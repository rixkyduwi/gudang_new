# Gudang New (Inventory & Sales Admin)

A Flask-based warehouse inventory management and sales tracking app.

## ⚙️ Python Version
- Direkomendasikan: **Python 3.10**
- Minimum: **Python 3.8+** (boleh 3.9, 3.10, 3.11; 3.10 disarankan karena stabil dan dukungan library lebih luas)

## 📦 Requirement
- Flask
- flask_cors
- flask_mysqldb
- flask_security_too
- flask_sqlalchemy
- flask_jwt_extended
- Flask-Bcrypt
- flask_compress
- flask_assets
- pillow
- pandas
- datetime
- openpyxl
- python-dateutil
- reportlab
- cryptography
- num2words
- waitress
- flask_pjax
- pyOpenSSL

Semua dapat diinstal via:
```bash
pip install -r requirements.txt
```

## 🗄️ Database
- MySQL database: `gudang_new`
- Default connection:
  - host: `localhost`
  - user: `root`
  - password: `` (kosong)
- config di `app/__init__.py`:
  - `MYSQL_HOST`, `MYSQL_USER`, `MYSQL_PASSWORD`, `MYSQL_DB`
  - `SQLALCHEMY_DATABASE_URI` diatur ke `mysql://root@localhost/gudang_new`

## 🚀 Menjalankan Server
```bash
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
python app.py
```
Server berjalan di `http://0.0.0.0:5000`.

### Opsi run
- Local (localhost): ganti `app.run` di `app.py` ke `host='127.0.0.1'`
- HTTPS adhoc: tambahkan `ssl_context='adhoc'` (untuk testing)

## 🔐 Autentikasi
- JWT (via cookie) di `app/login.py`
- Token diset dengan:
  - `create_access_token`, `create_refresh_token`
  - `set_access_cookies`, `set_refresh_cookies`
- Endpoint:
  - `GET /login` (tampilan login)
  - `POST /login` (login API)
  - `POST /logout`
  - `POST /refresh`
- Gunakan `@jwt_required()` untuk endpoint terproteksi.

## 🧩 Struktur Fitur
- `app/admin_master.py`:
  - Entity CRUD: `products`, `suppliers`, `customers`, `salespersons`, `senders`
  - Route base: `/admin/<entity>`
  - Create/Update/Delete dengan `POST`, `PUT`, `DELETE`
  - Soft-delete products (set `is_active=0`)
- `app/admin_sales.py`:
  - Sales assignment dan target
  - Endpoint: `/admin/salespersons/assign`, `/admin/salespersons/<id>/customers`, `/admin/sales_targets`

## 🧱 Static & Template
- HTML di `app/templates/admin/` untuk dashboard, login, master list, penerimaan/pengeluaran
- Assets di `app/static/` (css/js/images)

## 🛠️ Setup Tambahan
1. Buat database MySQL manual:
```sql
CREATE DATABASE gudang_new;
```
2. Import schema/tabel jika ada file SQL (tidak disertakan di repo ini)
3. Tambah user `users` dan hash password sesuai `flask_bcrypt`

## 🧪 T s
- Proyek ini belum menyediakan test suite otomatis dalam repo ini.

## 📌 Catatan
- Pastikan `SECRET_KEY` dan `SECURITY_PASSWORD_SALT` pada `app/__init__.py` di-update ke nilai aman di production.
- Hati-hati dengan penggunaan `session` + JWT; idealnya pilih satu arsitektur.
