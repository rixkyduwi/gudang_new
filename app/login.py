from . import app, bcrypt, User
from flask import request, render_template, redirect, url_for, jsonify, session, flash
#Singkat & to the point:
@app.route('/')
def homepahe():
    return redirect(url_for('login'))
    #return render_template('index.html')
    
#
#Yang kamu simpan di session['jwt_token'] itu access token (dari create_access_token). Kamu belum punya refresh token.
#
#Menaruh JWT di server session bikin arsitektur “setengah-setengah”: kamu sudah punya session (stateful) dan JWT (seharusnya stateless). Pilih salah satu:
#
#Pakai session saja (tanpa JWT), atau
#
#Full JWT via cookie HttpOnly + refresh token + CSRF di cookie (recommended untuk JWT di cookie).
#
#
#
#Di bawah ini contoh refactor rapi pakai Flask-JWT-Extended dengan:
#
#Access token (TTL pendek/sedang)
#
#Refresh token (TTL lebih lama)
#
#Simpan dua-duanya di cookie HttpOnly
#
#CSRF protect aktif untuk request mutasi
#
#Tanpa menyimpan JWT ke session (hapus itu)
#
#
#
#---
#
#Konfigurasi yang disarankan
#
# config
from datetime import timedelta

app.config.update(
    JWT_SECRET_KEY="qwdu92y17dqsu81",
    JWT_TOKEN_LOCATION=["cookies"],                 # JWT di cookie
    JWT_COOKIE_SECURE=True,                         # True di production (HTTPS)
    JWT_COOKIE_SAMESITE="Lax",                      # atau "Strict" jika cocok
    JWT_COOKIE_CSRF_PROTECT=True,                   # aktifkan CSRF untuk cookie-JWT
    JWT_ACCESS_TOKEN_EXPIRES=timedelta(minutes=30), # contoh
    JWT_REFRESH_TOKEN_EXPIRES=timedelta(days=7),
)
#
#> Saat JWT_COOKIE_CSRF_PROTECT=True, lib akan set dua cookie CSRF:
#csrf_access_token dan csrf_refresh_token.
#Untuk POST/PUT/PATCH/DELETE, kirim header X-CSRF-TOKEN berisi nilai dari cookie yang relevan.
#
#
#
#
#---
#
#Login: set cookie access & refresh (tanpa simpan ke session)
#
from flask import jsonify, request, render_template, redirect, url_for

from flask_jwt_extended import (
    create_access_token, create_refresh_token,
    set_access_cookies, set_refresh_cookies, jwt_required,
    get_jwt_identity,verify_jwt_in_request
)
def check_jwt_and_redirect():
    try:
        verify_jwt_in_request()
        return redirect(url_for('dashboard'))  # ganti 'dashboard' dengan endpoint dashboard Anda
    except Exception as e:
        return None
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'GET':
        redirect_to_dashboard = check_jwt_and_redirect()
        if redirect_to_dashboard:
            return redirect_to_dashboard
        return render_template('admin/login.html')

    data = request.get_json(silent=True) or {}
    username = data.get('username')
    password = data.get('password')

    user = User.query.filter_by(username=username).first()
    if not user:
        return jsonify(success=False, message="Username salah"), 401
    if not bcrypt.check_password_hash(user.password_bcrypt, password):
        return jsonify(success=False, message="Password salah"), 401

    access_token = create_access_token(identity=username)
    refresh_token = create_refresh_token(identity=username)

    resp = jsonify(success=True)
    set_access_cookies(resp, access_token)    # -> set cookie HttpOnly + csrf_access_token
    set_refresh_cookies(resp, refresh_token)  # -> set cookie HttpOnly + csrf_refresh_token
    session['username'] = username
    return resp, 200

#> Perhatikan: tidak ada session['jwt_token'].
#Kalau kamu butuh nama user di UI, kirim via API terpisah (protected) atau simpan display name non-sensitif di session biasa (opsional), tapi bukan JWT.
#
#
#
#
#---
#
#Protected route (cukup pakai decorator)
#
#from flask_jwt_extended import verify_jwt_in_request
#
#@app.route('/dashboard')
#@jwt_required()  # otomatis validasi access token dari cookie
#def dashboard():
#    user = get_jwt_identity()
#    return render_template('admin/dashboard.html', user=user)

#> Hapus decorator custom login_required + decode_token.
#decode_token tidak memverifikasi signature/expiry seperti @jwt_required()/verify_jwt_in_request() lakukan.
#
#
#
#
#---
#
#Refresh endpoint (rotasi access token)
#
from flask_jwt_extended import jwt_required, get_jwt_identity, set_access_cookies

@app.route('/refresh', methods=['POST'])
@jwt_required(refresh=True)  # minta refresh token dari cookie
def refresh():
    user = get_jwt_identity()
    new_access = create_access_token(identity=user)
    resp = jsonify(ok=True)
    set_access_cookies(resp, new_access)  # rotate access token
    return resp

#Front-end flow:
#
#Jika API balas 401 (expired) → panggil /refresh (POST + kirim header X-CSRF-TOKEN dari cookie csrf_refresh_token) → ulangi request.
#
#
#
#---
#
#Logout (bersihkan cookies)

from flask_jwt_extended import unset_jwt_cookies

@app.route('/logout', methods=['POST'])
def logout():
    resp = jsonify(message='Logout berhasil')
    unset_jwt_cookies(resp)  # hapus access+refresh cookies & csrf cookies
    flash('Logout berhasil')
    return resp, 200

#Kalau kamu tetap ingin redirect ke /login, lakukan redirect di client setelah menerima 200.
#unset_jwt_cookies bekerja di response yang kamu kirim (bukan di redirect final yang berbeda domain/URL tanpa membawa header set-cookie).
#
#
#---
#
#CSRF di AJAX (penting)
#
#Karena JWT di cookie, untuk POST/PUT/PATCH/DELETE kamu harus kirim header X-CSRF-TOKEN dari cookie:
#
#// contoh fetch util
#function getCookie(name) {
#  return document.cookie.split('; ').find(r => r.startsWith(name + '='))?.split('=')[1];
#}
#
##// untuk request yang memakai ACCESS token:
#const csrf = getCookie('csrf_access_token');
#
#fetch('/api/produk', {
#  method: 'POST',
#  headers: { 'Content-Type': 'application/json', 'X-CSRF-TOKEN': csrf },
#  body: JSON.stringify({ nama: 'Teh Botol' })
#});

#Untuk /refresh pakai csrf_refresh_token sebagai header X-CSRF-TOKEN.
#
#
#---
#
#Handler error JWT (opsional tapi rapi)

from flask_jwt_extended import JWTManager
jwt = JWTManager(app)

@jwt.unauthorized_loader
def unauthorized(reason):
    print("Unauthorized 401")
    return redirect(url_for('login'))

@jwt.expired_token_loader
def expired(jwt_header, jwt_payload):
    print("Token expired 401")
    return redirect(url_for('login'))

@jwt.invalid_token_loader
def invalid(reason):
    print("Invalid token 422")
    return redirect(url_for('login'))


#---
#
#Jawaban spesifik pertanyaanmu
#
#1. session['jwt_token'] sebaiknya tidak dipakai. Kalau kamu pilih JWT di cookie, simpan via set_access_cookies & set_refresh_cookies.
#
#
#2. Access vs Refresh:
#
#Access: dipakai tiap request, TTL lebih pendek.
#
#Refresh: hanya untuk endpoint /refresh, TTL lebih panjang.
#
#
#
#3. Expired sama atau tidak? Tidak perlu sama. Contoh sehat: Access 15–30 menit, Refresh 7–14 hari.
#
#
#
#Kalau mau, kasih tahu endpoint mana saja yang perlu kamu proteksi (tambah/edit/hapus). Aku bisa bikinin blueprint kecil CRUD (create/update/delete) yang sudah siap pakai: protected by @jwt_required(), kirim X-CSRF-TOKEN, plus contoh front-end fetch-nya.


# @app.route('/bikin_akun', methods=['GET', 'POST'])
# def register():
#     if request.method == 'POST':
#         # Tidak perlu jwt_required()

#         username = request.form.get('username')
#         password = request.form.get('password')

#         if not username or not password:
#             return jsonify({"msg": "Username dan password wajib diisi"}), 400

#         if User.query.filter_by(username=username).first():
#             return jsonify({"msg": "Username sudah digunakan"}), 400

#         hashed_password = bcrypt.generate_password_hash(password).decode('utf-8')

#         user = User(username=username, password=hashed_password, active=True)
#         db.session.add(user)
#         db.session.commit()

#         return jsonify({"msg": "Akun berhasil dibuat"}), 201

#     # Jika GET request, tampilkan form HTML atau pesan
#     return render_template('admin/register.html')

#         # Logout setelah registrasi berhasil
#         response = jsonify({'message': 'Logout berhasil'})
#         unset_jwt_cookies(response)
#         session.pop('jwt_token', None)
#         session.pop('username', None)
#         flash('Sukses Logout')
#         return redirect(url_for('login', msg='Registration Successful'))

#     return render_template('admin/register.html')
