from app import app

if __name__ == "__main__":
    # 1. STANDAR (Hanya akses lokal/localhost) 
    # akses: http://127.0.0.1:5000
    # app.run(host="127.0.0.1", port=5000, debug=True)

    # 2. SATU JARINGAN (Bisa diakses HP/Laptop lain via IP Server)
    # Gunakan host="0.0.0.0" agar Flask mendengarkan semua koneksi jaringan.
    # Contoh akses: http://192.168.1.15:5000
    app.run(host="0.0.0.0", port=5000, debug=True)

    # 3. SATU JARINGAN + HTTPS (Adhoc)
    # Menambahkan ssl_context='adhoc' untuk mengaktifkan HTTPS.
    # Catatan: Browser akan memunculkan peringatan "Your connection is not private".
    # Anda harus klik "Advanced" -> "Proceed" untuk masuk.
    # app.run(host="0.0.0.0", port=5000, debug=True, ssl_context='adhoc')