# Import library bawaan Python
import io,  os, textwrap, locale, uuid, calendar
# Import library pihak ketiga
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment
from io import BytesIO
from flask import render_template, request, jsonify, g, send_file,  Response
from PIL import Image
from dateutil.relativedelta import relativedelta
from flask_jwt_extended import jwt_required
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
# Fungsi untuk mengubah angka menjadi teks (terbilang)
from num2words import num2words
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation

# Import dari aplikasi lokal
from . import app, mysql, render_pjax

# Middleware untuk membuka koneksi database sebelum setiap request
@app.before_request
def before_request():
    g.con = mysql.connection.cursor()

# Middleware untuk menutup koneksi database setelah setiap request
@app.teardown_request
def teardown_request(exception):
    if hasattr(g, 'con'):
        g.con.close()
from flask import g, request
import time

@app.before_request
def before_request():
    g.start_time = time.time()

@app.after_request
def after_request(response):
    if hasattr(g, 'start_time'):
        elapsed = time.time() - g.start_time
        app.logger.info("%s %s selesai dalam %.2f detik",
                        request.method, request.path, elapsed)
    return response


# Fungsi untuk mengelola gambar (upload, edit, delete)
def do_image(do, table, id):
    try:
        if do == "delete":
            filename = get_image_filename(table, id)
            delete_image(filename)
            return True

        # Upload gambar baru
        file = request.files['gambar']
        if file is None or file.filename == '':
            return "default.jpg"
        else:
            filename = get_image_filename(table, id)
            delete_image(filename)
            return resize_and_save_image(file, table, id)

    except KeyError:
        # Jika kunci 'gambar' tidak ada dalam request.files
        if do == "edit" and table == "galeri":
            return True
        reset = request.form.get('reset', 'false')
        if reset == "true":
            g.con.execute(f"UPDATE {table} SET gambar = %s WHERE id = %s", ("default.jpg", id))
            g.con.connection.commit()
        return "default.jpg"

    except FileNotFoundError:
        pass  # Jika file tidak ditemukan, abaikan

    except Exception as e:
        print(str(e))
        return str(e)

# Fungsi untuk mengubah ukuran dan menyimpan gambar
def resize_and_save_image(file, table=None, id=None):
    img = Image.open(file).convert('RGB').resize((600, 300))
    random_name = uuid.uuid4().hex + ".jpg"
    destination = os.path.join(app.config['UPLOAD_FOLDER'], random_name)
    img.save(destination)

    if table and id:
        g.con.execute(f"UPDATE {table} SET gambar = %s WHERE id = %s", (random_name, id))
        g.con.connection.commit()
        return True
    else:
        return random_name

# Fungsi untuk mendapatkan nama file gambar dari database
def get_image_filename(table, id):
    g.con.execute(f"SELECT gambar FROM {table} WHERE id = %s", (id,))
    result = g.con.fetchone()
    if result == "default.jpg":
        return None
    return result[0] if result else None

# Fungsi untuk menghapus file gambar dari server
def delete_image(filename):
    if filename == "default.jpg":
        return True
    if filename:
        image_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if os.path.exists(image_path):
            os.remove(image_path)

# Fungsi untuk mengambil data dari database dan mengubahnya menjadi format dictionary
def fetch(query, params=None):
    print(query)
    print(params)
    g.con.execute(query, params or ())
    data = g.con.fetchall()
    column_names = [desc[0] for desc in g.con.description]
    return [dict(zip(column_names, row)) for row in data]

# Fungsi untuk mengambil daftar tahun dari database
def fetch_years(pilih):
    if pilih == "barang_masuk":
        query = "SELECT YEAR(tglfaktur) AS tahun FROM barang_masuk GROUP BY tahun"
    elif pilih == "barang_keluar":
        query = "SELECT YEAR(tglfaktur) AS tahun FROM barang_keluar GROUP BY tahun"
    else:
        query = ""
    g.con.execute(query)
    data_thn = g.con.fetchall()
    return [{'tahun': str(sistem[0])} for sistem in data_thn]

def fetch_sum(query, params=None):
    g.con.execute(query, params or ())
    row = g.con.fetchone()
    return row[0] if row and row[0] is not None else 0
# form json    
def req(key):
    return request.json.get(key)

# update stok gudang
def update_stok_gudang(id_barang, jmlpermintaan, operasi):
    jml = int(jmlpermintaan)

    # Ambil stok gudang dan limit sekali query
    g.con.execute("""
        SELECT bg.sisa_gudang, b.stoklimit
        FROM barang b
        LEFT JOIN barang_gudang bg ON bg.id_barang = b.id
        WHERE b.id = %s
    """, (id_barang,))
    row = g.con.fetchone()
    
    print("=" * 40)
    print(f"DEBUG update_stok_gudang")
    print(f"id_barang     : {id_barang}")
    print(f"jmlpermintaan : {jmlpermintaan}")
    print(f"operasi       : {operasi}")
    print(f"sisa_gudang   : {row[0] if row else 'None'}")
    print(f"stoklimit     : {row[1] if row else 'None'}")

    if not row:
        return jsonify({"error": "Data barang tidak ditemukan"}), 404

    sisa_gudang, stoklimit = row if row[0] is not None else (0, row[1])

    if operasi == "kurangi":
        new_qty = sisa_gudang - jml
    else:
        new_qty = sisa_gudang + jml

    if new_qty < 0:
        return jsonify({"error": "Stok Kurang"}), 404

    ket = "Stok Aman" if new_qty > stoklimit else "Stok Tidak Aman"

    print(f"keterangan    : {ket}")
    print("=" * 40)
    # Insert jika belum ada di barang_gudang
    g.con.execute("""
        INSERT INTO barang_gudang (id_barang, sisa_gudang, keterangan)
        VALUES (%s, %s, %s)
        ON DUPLICATE KEY UPDATE
            sisa_gudang = VALUES(sisa_gudang),
            keterangan = VALUES(keterangan)
    """, (id_barang, new_qty, ket))
    g.con.connection.commit()

    return True

#function di render_template
@app.template_filter('format_currency')
def format_currency(value):
    locale.setlocale(locale.LC_ALL, 'id_ID.UTF-8')
    return locale.currency(value, grouping=True, symbol='Rp')

@app.template_filter('clean_currency')
def clean_currency(value):
    """Remove 'Rp', dots, and spaces, then convert to float"""
    cleaned_value = value.replace('Rp', '').replace('RP', '').replace('.', '').replace(',', '.').strip()
    return float(cleaned_value)

@app.template_filter('floatformat')
def floatformat(value, precision=2):
    try:
        return f"{float(value):.{precision}f}"
    except (ValueError, TypeError):
        return value

@app.template_filter('format_rp')
def format_rp(value):
    value =f"{int(value):,}".replace(",", ".")
    value = f"Rp. {value},00"
    return value

@app.template_filter('format_rupiah')
def format_rupiah(value):
    try:
        # Ubah ke Decimal dan pastikan dua digit di belakang koma
        value = Decimal(value).quantize(Decimal("0.01"))

        # Pecah jadi bagian depan dan koma
        bagian_utama, bagian_koma = str(value).split('.')

        # Format bagian utama dengan titik pemisah ribuan
        bagian_utama = f"{int(bagian_utama):,}".replace(",", ".")

        # Gabungkan dan pakai koma sebagai desimal (gaya Indonesia)
        return f"{bagian_utama},{bagian_koma}"
    except (InvalidOperation, ValueError, TypeError):
        return value

@app.template_filter('format_date')
def format_date(value):
    if not value:
        return ""

    # Kalau sudah datetime/date object
    if isinstance(value, (datetime, date)):
        return value.strftime("%d/%m/%Y")

    # Kalau string, coba parse dengan beberapa format
    if isinstance(value, str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
            try:
                dt = datetime.strptime(value, fmt)
                return dt.strftime("%d/%m/%Y")
            except ValueError:
                continue

    # Fallback kalau gak dikenali
    return str(value)

list_bulan = [{"value":"1","nama_bulan":"Januari"},{"value":"2","nama_bulan":"Februari"},{"value":"3","nama_bulan":"Maret"},
                {"value":"4","nama_bulan":"April"},{"value":"5","nama_bulan":"Mei"},{"value":"6","nama_bulan":"Juni"},{"value":"7","nama_bulan":"Juli"},
                {"value":"8","nama_bulan":"Agustus"},{"value":"9","nama_bulan":"September"},{"value":"10","nama_bulan":"Oktober"},
                {"value":"11","nama_bulan":"November"},{"value":"12","nama_bulan":"Desember"}] 

def time_zone_wib():
    return datetime.utcnow() + timedelta(hours=7)               
@app.route('/admin/dashboard')
def dashboard():
    today = time_zone_wib()
    hari_ini = today.strftime("%d/%m/%Y")
    tanggal_db = today.strftime("%Y-%m-%d")
    bulan_ini = today.month   # lebih aman
    tahun_ini = today.year
    # untuk bulan berjalan
    last = calendar.monthrange(tahun_ini, bulan_ini)[1]
    start_month = date(tahun_ini, bulan_ini, 1)
    end_month   = date(tahun_ini, bulan_ini, last) + timedelta(days=1)
    print(start_month)
    print(end_month)
    
    # untuk hari ini
    start_day = date.fromisoformat(tanggal_db)          # yyyy-mm-dd
    end_day   = start_day + timedelta(days=1)
    print( start_day)
    print(end_day)

    # target sales + realtime sales dalam 1 query
    target_sales = fetch("""
        SELECT ts.id, s.nama_sales, ts.target, s.id AS id_sales,
               COALESCE(SUM(dbk.harga_total), 0) AS data_realtime
        FROM target_sales ts
        INNER JOIN sales s ON s.id = ts.id_sales
        LEFT JOIN detail_sales ds ON ds.id_sales = s.id
        LEFT JOIN barang_keluar bk 
               ON bk.id_sales = ds.id
              AND bk.tglfaktur >= %s AND bk.tglfaktur < %s
        LEFT JOIN detail_barang_keluar dbk 
               ON dbk.id_barang_keluar = bk.id
        WHERE ts.tahun = %s AND ts.bulan = %s
        GROUP BY ts.id, s.nama_sales, ts.target, s.id
    """, (start_month, end_month, tahun_ini, bulan_ini))
    # inkaso (penerimaan pembayaran hari ini)
    print('inkaso')
    print(tanggal_db)
    g.con.execute("""
        SELECT COALESCE(SUM(dbk.harga_total), 0)
        FROM detail_barang_keluar dbk
        INNER JOIN barang_keluar bk ON bk.id = dbk.id_barang_keluar
        WHERE bk.tanggal_pembayaran >= %s AND bk.tanggal_pembayaran < %s
    AND bk.lunas_tidak = 'Lunas'
    """, (start_day, end_day))
    inkaso = g.con.fetchone()[0]

    # performa penjualan bulan ini
    g.con.execute("""
        SELECT COALESCE(SUM(dbk.harga_total), 0)
        FROM detail_barang_keluar dbk
        INNER JOIN barang_keluar bk ON bk.id = dbk.id_barang_keluar
        WHERE bk.tglfaktur >= %s AND bk.tglfaktur < %s
    """, (start_month, end_month))
    performa_penjualan = g.con.fetchone()[0] or 0

    # performa belanja bulan ini
    g.con.execute("""
        SELECT COALESCE(SUM(dbm.harga_total), 0)
        FROM detail_barang_masuk dbm
        INNER JOIN barang_masuk bm ON bm.id = dbm.id_barang_masuk
        WHERE bm.tglfaktur >= %s AND bm.tglfaktur < %s
    """, (start_month, end_month,))
    performa_belanja = g.con.fetchone()[0] or 0
    print(performa_penjualan)
    print(performa_belanja)
    revenue = performa_penjualan - performa_belanja
    print(revenue)

    return render_pjax(
        'admin/dashboard.html',
        tahun_ini=tahun_ini,
        target_sales=target_sales,
        inkaso=inkaso,
        hari_ini=hari_ini,
        revenue=revenue,
        performa_penjualan=performa_penjualan,
        performa_belanja=performa_belanja
    )
# Halaman penerimaan barang
@app.route('/admin/penerimaan-tambah', methods=['GET'])
def adminpenerimaantambah():
    # Ambil data tambahan
    data_barang = fetch("SELECT id, kode_barang, qty, nama_barang FROM barang ORDER BY id")
    data_pabrik = fetch("SELECT id, nama_supplier, alamat, tlp FROM pabrik ORDER BY id")
    thn = fetch_years("barang_masuk")

    return render_template("admin/tambah-penerimaan.html", tahun=thn, data_pabrik=data_pabrik, data_barang=data_barang,
    tanggal=time_zone_wib().date())

@app.route('/admin/penerimaan')
def adminpenerimaan():
    tahun = request.args.get('tahun')
    bulan = request.args.get('bulan')
    tanggal = request.args.get('tanggal')
    nama_principle = request.args.get('nama_principle', type=str)
    page = request.args.get('page', default=1, type=int)
    per_page = request.args.get('per_page', default=10, type=int)

    filters, params = [], []

    clause, prms = build_date_range(
        year=int(tahun) if tahun else None,
        month=int(bulan) if bulan else None,
        day=int(tanggal) if tanggal else None,
        alias="bm",              # <- alias sesuai query kamu
        col="tglfaktur"
    )
    if clause:
        filters.append(clause)
        params.extend(prms)

    if nama_principle:
        filters.append("p.nama_supplier = %s")
        params.append(nama_principle)

    where_clause = "WHERE " + " AND ".join(filters) if filters else ""
    base_query = f"""
        SELECT bm.id, bm.tglfaktur, bm.nofaktur, p.nama_supplier,
               SUM(dbm.harga_total) AS performa_belanja
        FROM barang_masuk bm
        JOIN pabrik p ON p.id = bm.id_supplier
        JOIN detail_barang_masuk dbm ON dbm.id_barang_masuk = bm.id
        {where_clause}
        GROUP BY bm.id
        ORDER BY bm.id DESC
        LIMIT %s OFFSET %s
    """
    offset = (page - 1) * per_page
    params.extend([per_page, offset])
    info_list = fetch(base_query, params)
    print(info_list)
    if info_list != []:
        data_detail_masuk_all = fetch("""
            SELECT dbm.*, b.nama_barang, b.kode_barang, b.qty
            FROM detail_barang_masuk dbm
            JOIN barang b ON b.id = dbm.id_barang
            WHERE dbm.id_barang_masuk IN %s
        """, (tuple([i['id'] for i in info_list]),))
        for i in info_list:
            detail = []
            for j in data_detail_masuk_all:
                if i['id'] == j['id_barang_masuk']:
                    detail.append(j)
            i['detail'] = detail 
    count_query = f"""
        SELECT COUNT(DISTINCT bm.id)
        FROM barang_masuk bm
        JOIN pabrik p ON p.id = bm.id_supplier
        {where_clause}
    """
    g.con.execute(count_query, params[:-2] if filters else ())
    total_records = g.con.fetchone()[0]

    total_pages = (total_records + per_page - 1) // per_page
    has_next = page < total_pages
    has_prev = page > 1
    page_range = range(max(1, page - 3), min(total_pages + 1, page + 3))

    return render_pjax(
        "admin/penerimaan.html",
        
        info_list=info_list,
        tahun=fetch_years("barang_masuk"),  # bisa diisi pakai fetch_years kalau perlu
        data_pabrik=fetch("SELECT id, nama_supplier, alamat, tlp FROM pabrik ORDER BY id"),
        data_barang=fetch("SELECT id, kode_barang, nama_barang, qty, stoklimit FROM barang ORDER BY id"),
        tanggal=datetime.utcnow().date(),
        page=page,
        per_page=per_page,
        total_pages=total_pages,
        total_records=total_records,
        has_next=has_next,
        has_prev=has_prev,
        page_range=page_range
    )
@app.route('/admin/penerimaan/tambah/id', methods=['POST'])
@jwt_required()
def tambah_id_penerimaan_new():
    data = request.json
    performabelanja = Decimal('0.00')
    print(data)
    for i in data['items']:
        # Validasi harga_total harus sesuai dengan harga_satuan * jml_menerima
        try:            
            try:
                harga_satuan = Decimal(i['harga_satuan'])
                harga_total = Decimal(i['harga_total'])
            except InvalidOperation:
                return jsonify({"error": "Format harga tidak valid"}), 400
            jml_menerima = int(i['jml_menerima'])
            harga_satuan = Decimal(i['harga_satuan'])
            harga_total_calculated = jml_menerima * harga_satuan
            if Decimal(i['harga_total']) != harga_total_calculated:
                return jsonify({"error": f"Harga total tidak sesuai. Seharusnya {harga_total_calculated}"}), 400
        except ValueError:
            return jsonify({"error": "Jumlah menerima dan harga satuan harus berupa angka"}), 400
        performabelanja += Decimal(i['harga_total'])

    g.con.execute("SELECT id FROM pabrik WHERE id = %s", (data['id_supplier'],))
    supplier = g.con.fetchone()
    if not supplier:
        return jsonify({"error": "Supplier not found"}), 404
    g.con.execute("""
        INSERT INTO barang_masuk (tglfaktur, nofaktur, id_supplier)
        VALUES (%s, %s, %s)
    """, (data['tglfaktur'], data['nofaktur'], data['id_supplier']))
    
    new_id = g.con.lastrowid
    print(data['items'])
    #cek barang
    for i in data['items']:
        print(i)
        print(i['id_barang'])
        g.con.execute("SELECT sisa_gudang FROM barang_gudang WHERE id_barang = %s", (i['id_barang'],))
        result = g.con.fetchone()
        g.con.execute("SELECT stoklimit FROM barang WHERE id = %s", (i['id_barang'],))
        stoklimit = g.con.fetchone()
        g.con.execute("""INSERT INTO detail_barang_masuk (id_barang_masuk, id_barang, jml_menerima, harga_satuan, harga_total,pembayaran,jatohtempo,lunastidak)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s) """,
        (new_id, i['id_barang'], i['jml_menerima'], i['harga_satuan'],  i['harga_total'], i['pembayaran'], i['jthtempo'], "Tidak Lunas"))
        
        if result:
            print(result[0])
            new_qty = result[0] + int(i['jml_menerima'])
            print(new_qty)
            if new_qty <= stoklimit[0]:
                ket = "Stok Tidak Aman"
            else:
                ket = "Stok Aman"
                g.con.execute("""
                    UPDATE barang_gudang SET sisa_gudang = %s, keterangan = %s WHERE id_barang = %s
                """, (new_qty, ket, i['id_barang']))
        else:
            new_qty = int(i['jml_menerima'])
            if new_qty <= stoklimit[0]:
                ket = "Stok Tidak Aman"
            else :
                ket = "Stok Aman"
            g.con.execute("""INSERT INTO barang_gudang(id_barang, sisa_gudang, keterangan) VALUES (%s,%s,%s) """,
        ( i['id_barang'], new_qty, ket))
        
    g.con.connection.commit()
    return jsonify({"msg": "SUKSES"})

@app.route('/admin/penerimaan/edit/<int:id>', methods=['GET'])
def edit_id_penerimaan_new(id):
    # Ambil data barang_masuk
    g.con.execute("""
        SELECT barang_masuk.id, barang_masuk.id_supplier, barang_masuk.tglfaktur, barang_masuk.nofaktur, 
        pabrik.nama_supplier, pabrik.alamat, pabrik.tlp
        FROM barang_masuk
        INNER JOIN detail_barang_masuk dbm ON dbm.id_barang_masuk = barang_masuk.id
        INNER JOIN pabrik ON pabrik.id = barang_masuk.id_supplier
        WHERE barang_masuk.id = %s
    """, (id,))
    barang_masuk = g.con.fetchone()
    # Ambil detail_barang_masuk
    g.con.execute("""
        SELECT dbm.id, dbm.id_barang_masuk, dbm.id_barang, dbm.jml_menerima, dbm.harga_satuan,
        dbm.harga_total, dbm.jatohtempo, dbm.tglpembayaran, dbm.pembayaran,
        dbm.keterangan, dbm.lunastidak, barang.nama_barang, barang.kode_barang, barang.qty
        FROM detail_barang_masuk dbm
        INNER JOIN barang ON barang.id = dbm.id_barang
        WHERE dbm.id_barang_masuk = %s
    """, (id,))
    detail_barang_masuk = g.con.fetchall()
    for i in detail_barang_masuk:
        print(i[6])

    # Fetch data_barang for the template
    
    data_barang = fetch("SELECT id, kode_barang, qty, nama_barang FROM barang ORDER BY id")
    data_pabrik = fetch("SELECT id, nama_supplier, alamat, tlp FROM pabrik ORDER BY id")
    thn = fetch_years("barang_masuk")

    return render_template(
        "admin/edit_penerimaan.html",
        barang_masuk=barang_masuk,
        detail_barang_masuk=detail_barang_masuk,
        data_barang=data_barang,
        tahun=thn, data_pabrik=data_pabrik,
        tanggal=time_zone_wib().date()
    )

@app.route('/admin/penerimaan/edit/<int:id_barang_masuk>', methods=['PUT'])
@jwt_required()
def edit_id_penerimaan_action(id_barang_masuk):
    # PUT method
    form_data = request.json
    print(form_data)
    id_suplier = form_data['id_supplier']
    g.con.execute("""
        SELECT id, nama_supplier, alamat, tlp FROM pabrik WHERE id = %s
    """, (id_suplier,))
    supplier = g.con.fetchone()
    if not supplier:
        return jsonify({"error": "Supplier not found"}), 404
    performabelanja=0
    for item in form_data['items']:
        print(item.get('harga_total','0'))
        harga_total = int(item.get('harga_total', '0'))
        performabelanja += harga_total

    # if str(form_data.get('pajak', '')).lower() == 'yes':
        # pajak_rate = decimal('11') / decimal('100')
        # pajak_amount = pajak_rate * performabelanja
        # performabelanja += pajak_amount  # sekarang ppn ditambahkan dengan benar
    # Update nofaktur di tabel barang_masuk
    g.con.execute("""
        UPDATE barang_masuk 
        SET nofaktur = %s, tglfaktur = %s, id_supplier = %s
        WHERE id = %s
    """, (
        form_data['nofaktur'],
        form_data['tglfaktur'],
        id_suplier,
        int(id_barang_masuk)
    ))
    print(form_data['nofaktur'])
    # Update detail_barang_masuk sesuai data user
    items = form_data['items']
    # hapus stok lama dari gudang
    for i in items:
        g.con.execute("SELECT jml_menerima FROM detail_barang_masuk WHERE id_barang_masuk = %s AND id_barang = %s", (int(id_barang_masuk), i['id_barang']))
        result = g.con.fetchone()
        if result:
            update_stok_gudang(i['id_barang'],result[0],'kurangi')
    # Hapus semua detail_barang_masuk yang terkait dengan id_barang_masuk ini
    g.con.execute("DELETE FROM detail_barang_masuk WHERE id_barang_masuk = %s", (int(id_barang_masuk),))

    # Insert data baru dari user
    fields = [
        'id_barang', 'jml_menerima', 'harga_satuan', 'harga_total', 'pembayaran', 'jatohtempo'
    ]
    for i in items:
        query = f"""INSERT INTO detail_barang_masuk 
            (id_barang_masuk, {', '.join(fields)}) 
            VALUES (%s, {', '.join(['%s'] * len(fields))})"""
        
        values = (int(id_barang_masuk),) + tuple(i.get(field) for field in fields)
        print(query, values)
        g.con.execute(query, values)
        # masukan stok baru
        update_stok_gudang(i['id_barang'],i['jml_menerima'],'tambah')
        
    g.con.connection.commit()
    return jsonify({"msg": "SUKSES"})

@app.route('/admin/penerimaan/hapus/id', methods=['DELETE'])
@jwt_required()
def hapus_id_penerimaan():
    id = request.json.get('id')
    g.con.execute("SELECT id, id_barang, jml_menerima  FROM detail_barang_masuk WHERE id_barang_masuk = %s", (int(id),))
    data = g.con.fetchall()
    if not data:
        return jsonify({"error": "Data tidak ditemukan"}), 404
    for i in data:
        print(i)
        _id, id_barang, jumlah = i
        # Kurangi jumlah di barang_gudang
        g.con.execute("SELECT sisa_gudang FROM barang_gudang WHERE id_barang = %s", (id_barang,))
        gudang = g.con.fetchone()
        g.con.execute("SELECT stoklimit FROM barang WHERE id = %s", (id_barang,))
        stoklimit = g.con.fetchone()
        if gudang:
            new_jumlah = max(0, gudang[0] - jumlah)
            if int(new_jumlah) <= stoklimit[0]:
                ket = "Stok Tidak Aman"
            else :
                ket = "Stok Aman"
            g.con.execute("UPDATE barang_gudang SET sisa_gudang = %s, keterangan = %s WHERE id_barang = %s", (new_jumlah, ket, id_barang))
            
        # Hapus dari detail_barang_masuk
        g.con.execute("DELETE FROM detail_barang_masuk WHERE id = %s", (int(_id),))
        g.con.connection.commit()
    # Hapus dari barang_masuk
    g.con.execute("DELETE FROM barang_masuk WHERE id = %s", (int(id),))
    g.con.connection.commit()
    return jsonify({"msg": "BERHASIL DIHAPUS"})
    
@app.route('/admin/penyimpanan')
def adminpenyimpanan():
    # Param lama tetap didukung biar nggak breaking
    product_id = request.args.get('product_id') or request.args.get('id_barang')
    q          = request.args.get('q')  # optional: cari kode/nama
    low_only   = request.args.get('low_only', default=0, type=int)  # 1 = hanya stok tidak aman
    page       = request.args.get('page', default=1, type=int)
    per_page   = request.args.get('per_page', default=10, type=int)

    # Base FROM: ambil dari view v_product_stock + join products buat meta (code, name, unit, stock_min)
    base_from = """
        FROM v_product_stock s
        JOIN products p ON p.id = s.product_id
    """

    # Filter dinamis
    filters = []
    params  = []
    if product_id:
        filters.append("s.product_id = %s")
        params.append(product_id)

    if q:
        filters.append("(p.code LIKE %s OR p.name LIKE %s)")
        like = f"%{q}%"
        params.extend([like, like])

    # low_only: stok on hand <= stock_min
    if low_only == 1:
        filters.append("(s.qty_on_hand <= p.stock_min)")

    where_clause = "WHERE " + " AND ".join(filters) if filters else ""

    # Count total row (untuk pagination)
    count_sql = f"SELECT COUNT(*) {base_from} {where_clause}"
    g.con.execute(count_sql, params)
    total_records = g.con.fetchone()[0]

    # Paging
    total_pages = (total_records + per_page - 1) // per_page
    offset = (page - 1) * per_page

    # Data list
    list_sql = f"""
        SELECT
          s.product_id,
          p.code       AS kode_barang,
          p.name       AS nama_barang,
          p.unit,
          p.stock_min  AS stok_limit,
          s.qty_on_hand AS sisa_gudang,
          CASE
            WHEN s.qty_on_hand <= p.stock_min THEN 'Stok Tidak Aman'
            ELSE 'Aman'
          END AS status
        {base_from}
        {where_clause}
        ORDER BY s.product_id ASC
        LIMIT %s OFFSET %s
    """
    list_params = params + [per_page, offset]
    info_list = fetch(list_sql, list_params)
    print(info_list)
    # Dropdown daftar barang (untuk filter select)
    nama_barang_list = fetch("""
        SELECT p.id, p.code AS kode_barang, p.name AS nama_barang
        FROM products p
        ORDER BY p.name ASC
    """)

    has_next = page < total_pages
    has_prev = page > 1
    page_range = range(max(1, page - 3), min(total_pages + 1, page + 3))

    return render_pjax(
        "admin/penyimpanan.html",
        info_list=info_list,
        nama_barang=nama_barang_list,
        page=page,
        per_page=per_page,
        total_pages=total_pages,
        total_records=total_records,
        has_next=has_next,
        has_prev=has_prev,
        page_range=page_range,
        q=q,
        product_id=product_id,
        low_only=low_only
    )
    
@app.route('/admin/penyimpanan', methods=['PUT'])
@jwt_required()
def edit_penyimpanan():
    """
    Payload JSON:
    {
      "product_id": 123,
      "new_qty": 50,         # stok final yang diinginkan
      "note": "Stock opname Sept"
    }
    """
    data = request.json or {}
    print(data)
    product_id = data.get('product_id')
    new_qty    = data.get('new_qty')
    note       = data.get('note') or 'Manual adjustment'

    if not product_id or new_qty is None:
        return jsonify({"error": "product_id dan new_qty wajib"}), 400

    try:
        # 1) Ambil stok saat ini (qty_on_hand) + batas aman (stock_min)
        sql_now = """
            SELECT s.qty_on_hand, p.stock_min, p.name
            FROM v_product_stock s
            JOIN products p ON p.id = s.product_id
            WHERE s.product_id = %s
        """
        g.con.execute(sql_now, (product_id,))
        row = g.con.fetchone()
        if not row:
            return jsonify({"error": "Produk tidak ditemukan"}), 404

        current_qty = int(row[0] or 0)
        stock_min   = int(row[1] or 0)
        product_name= row[2]

        # 2) Hitung delta
        new_qty = int(new_qty)
        delta = new_qty - current_qty
        if delta == 0:
            status = "Stok Tidak Aman" if new_qty <= stock_min else "Aman"
            return jsonify({
                "msg": "Tidak ada perubahan stok",
                "product_id": product_id,
                "product_name": product_name,
                "qty_on_hand": new_qty,
                "status": status
            })

        # 3) Insert stock_moves (ADJUSTMENT)
        if delta > 0:
            g.con.execute(
                "INSERT INTO stock_moves (product_id, ref_type, qty_in,  note) VALUES (%s,'ADJUSTMENT',%s,%s)",
                (product_id, delta, note)
            )
        else:
            g.con.execute(
                "INSERT INTO stock_moves (product_id, ref_type, qty_out, note) VALUES (%s,'ADJUSTMENT',%s,%s)",
                (product_id, abs(delta), note)
            )

        g.con.connection.commit()

        # 4) Ambil stok terbaru setelah insert (opsional)
        g.con.execute(sql_now, (product_id,))
        row2 = g.con.fetchone()
        new_now = int(row2[0] or 0)
        status = "Stok Tidak Aman" if new_now <= stock_min else "Aman"

        return jsonify({
            "msg": "SUKSES",
            "product_id": product_id,
            "product_name": product_name,
            "qty_on_hand": new_now,
            "status": status,
            "delta": delta
        })
    except Exception as e:
        g.con.connection.rollback()
        return jsonify({"error": str(e)}), 500

@app.route('/admin/penyimpanan/hapus/id', methods=['DELETE'])
@jwt_required()
def hapus_id_penyimpanan():
    try:
        id = request.json.get('id')
        # Hapus dari barang_gudang
        g.con.execute("DELETE FROM barang_gudang WHERE id = %s", (int(id),))
        g.con.connection.commit()

        return jsonify({"msg": "BERHASIL DIHAPUS"})
    except Exception as e:
        print(e)
        return jsonify({"error": str(e)}), 500

@app.route('/admin/pengeluaran-tambah')
def tambahpengeluaran():
    data_barang = fetch("""SELECT barang.id, barang.nama_barang,barang.qty, barang_gudang.sisa_gudang, barang.stoklimit 
    FROM barang INNER JOIN barang_gudang on barang_gudang.id_barang = barang.id ORDER BY barang_gudang.id""")
    print(data_barang)
    data_sales = fetch("""SELECT * FROM detail_sales 
    INNER JOIN sales ON sales.id = detail_sales.id_sales 
    INNER JOIN outlet ON outlet.id = detail_sales.id_outlet ORDER BY detail_sales.id""")
    nama_sales = fetch("""SELECT DISTINCT nama_sales FROM sales""")
    data_outlet = fetch("""SELECT outlet.nama_outlet, outlet.alamat_outlet FROM detail_sales 
    INNER JOIN sales ON sales.id = detail_sales.id_sales 
    INNER JOIN outlet ON outlet.id = detail_sales.id_outlet ORDER BY detail_sales.id""")
    tanggal = time_zone_wib()
    
    # Hitung Jatuh Tempo tambah 1 bulan
    jth_tempo = tanggal + relativedelta(months=1)
    jth_tempo = jth_tempo.strftime("%Y-%m-%d")
    # Hitung nofaktur berdasarkan jumlah barang_keluar pada tanggal yang sama
    tanggal = tanggal.strftime('%Y-%m-%d')
    g.con.execute("""
        SELECT RIGHT(nomerfaktur, 3) AS last_3
        FROM barang_keluar WHERE tglfaktur = %s 
        ORDER BY nomerfaktur DESC LIMIT 1
    """, (tanggal,))
    row = g.con.fetchone()
    last_3 = int(row[0]) if row and row[0] else 0
    print(last_3)

    nofaktur = f"{tanggal.replace('-', '')}{last_3 + 1:03d}" # Gabungkan tanggal dengan nomor urut (001, 002, ...)
    print(nofaktur)
    return render_template("admin/tambah-pengeluaran.html", data_sales = data_sales, data_outlet = data_outlet, nama_sales=nama_sales,  
                           data_barang = data_barang, tanggal = tanggal, nofaktur = nofaktur, jth_tempo = jth_tempo)

@app.route('/api/ceknofaktur/<path:tanggal>')
def ceknofaktur(tanggal):
    dt_tanggal = None
    print(tanggal)
    # Coba parse dengan 2 format umum
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            dt_tanggal = datetime.strptime(tanggal, fmt)
            break
        except ValueError:
            continue

    if not dt_tanggal:
        return {"error": "Format tanggal tidak dikenali (pakai YYYY-MM-DD atau DD/MM/YYYY)"}, 400

    # Hitung Jatuh Tempo tambah 1 bulan
    jth_tempo = (dt_tanggal + relativedelta(months=1)).strftime("%d/%m/%Y")

    # Query DB pakai format ISO standar
    tgl_str = dt_tanggal.strftime('%Y-%m-%d')
    g.con.execute("""
        SELECT RIGHT(nomerfaktur, 3) AS last_3
        FROM barang_keluar 
        WHERE tglfaktur = %s 
        ORDER BY nomerfaktur DESC LIMIT 1
    """, (tgl_str,))
    row = g.con.fetchone()
    last_3 = int(row[0]) if row and row[0] else 0

    # Gabungkan tanggal (yyyymmdd) + nomor urut
    nofaktur = f"{dt_tanggal.strftime('%Y%m%d')}{last_3 + 1:03d}"

    return jsonify({"jatuh_tempo": jth_tempo, "nofaktur": nofaktur})
@app.route('/admin/pengeluaran')
def adminpengeluaran():
    tahun        = request.args.get('tahun')
    bulan        = request.args.get('bulan')
    tanggal      = request.args.get('tanggal')
    nama_sales_q = request.args.get('nama_sales')
    nama_outlet_q= request.args.get('nama_outlet')
    page         = request.args.get('page', default=1, type=int)
    per_page     = request.args.get('per_page', default=10, type=int)
   
    # Build filter
    filters, params = [], []

    # Catatan: selalu qualify kolom tanggal dengan alias tabel 'bk'
    clause, prms = build_date_range(
        year=int(tahun) if tahun else None,
        month=int(bulan) if bulan else None,
        day=int(tanggal) if tanggal else None,
        alias="bk",              # <- alias sesuai query kamu
        col="tglfaktur"
    )
    if clause:
        filters.append(clause)
        params.extend(prms)
    if nama_sales_q:
        filters.append("s.nama_sales = %s")
        params.append(nama_sales_q)
    if nama_outlet_q:
        filters.append("o.nama_outlet = %s")
        params.append(nama_outlet_q)

    where_clause = (" WHERE " + " AND ".join(filters)) if filters else ""

    # Data utama (per faktur) + agregasi performa_sales
    base_query = f"""
        SELECT
            bk.id,
            bk.nomerfaktur,
            bk.tglfaktur,
            bk.jatuhtempo,
            bk.cashtempo,
            bk.pajak,
            YEAR(bk.tglfaktur)  AS tahun,
            s.nama_sales,
            o.nama_outlet,
            COALESCE(SUM(dbk.harga_total), 0) AS performa_sales
        FROM barang_keluar bk
        JOIN detail_sales ds ON ds.id = bk.id_sales
        JOIN sales        s  ON s.id  = ds.id_sales
        JOIN outlet       o  ON o.id  = ds.id_outlet
        LEFT JOIN detail_barang_keluar dbk ON dbk.id_barang_keluar = bk.id
        {where_clause}
        GROUP BY bk.id, bk.nomerfaktur, bk.tglfaktur, s.nama_sales, o.nama_outlet
        ORDER BY bk.id DESC
    """

    # Pagination
    offset = (page - 1) * per_page
    paginated_query = base_query + " LIMIT %s OFFSET %s"
    params_with_pagination = params + [per_page, offset]
    print(paginated_query)
    print(params_with_pagination)
    info_list = fetch(paginated_query, params_with_pagination)

    # Count total faktur (DISTINCT bk.id)
    count_query = f"""
        SELECT COUNT(*) FROM (
            SELECT DISTINCT bk.id
            FROM barang_keluar bk
            JOIN detail_sales ds ON ds.id = bk.id_sales
            JOIN sales        s  ON s.id  = ds.id_sales
            JOIN outlet       o  ON o.id  = ds.id_outlet
            {where_clause}
        ) x
    """
    g.con.execute(count_query, params)
    total_records = g.con.fetchone()[0]

    # Ambil semua detail item untuk faktur di halaman ini — 1 query (hindari N+1)
    detail_map = {}
    if info_list:
        page_ids = [row['id'] for row in info_list]
        # Buat placeholder (%s, %s, ...) aman
        placeholders = ", ".join(["%s"] * len(page_ids))
        detail_sql = f"""
            SELECT
                dbk.id_barang_keluar,
                dbk.jmlpermintaan,
                dbk.harga_satuan,
                dbk.diskon,
                dbk.harga_total,
                COALESCE(dbk.batch,'-'),
                COALESCE(dbk.ed,'-'),
                b.kode_barang, b.nama_barang, b.qty, b.stoklimit
            FROM detail_barang_keluar dbk
            JOIN barang b ON b.id = dbk.id_barang
            WHERE dbk.id_barang_keluar IN ({placeholders})
            ORDER BY dbk.id
        """
        details = fetch(detail_sql, page_ids)
        for d in details:
            detail_map.setdefault(d['id_barang_keluar'], []).append(d)

    # Tempel detail ke setiap faktur
    for row in info_list:
        row['detail'] = detail_map.get(row['id'], [])
    
    # Dropdown/helper data (boleh di-cache)
    # Lookup tambahan untuk filter UI
    nama_outlet = fetch("SELECT nama_outlet FROM outlet GROUP BY nama_outlet")
    nama_sales  = fetch("SELECT nama_sales  FROM sales  GROUP BY nama_sales")
    
    # Tahun list (pakai helpermu yang memang butuh query)
    thn = fetch_years("barang_keluar")

    # Pagination info
    total_pages = (total_records + per_page - 1) // per_page
    has_next = page < total_pages
    has_prev = page > 1
    page_range = range(max(1, page - 3), min(total_pages + 1, page + 3))

    # Render
    return render_pjax(
        "admin/pengeluaran.html",
        info_list=info_list,
        tahun=thn,
        nama_outlet=nama_outlet,
        nama_sales=nama_sales,
        page=page,
        per_page=per_page,
        total_pages=total_pages,
        total_records=total_records,
        has_next=has_next,
        has_prev=has_prev,
        page_range=page_range
    )

@app.route('/api/barangkeluar')
def api_barangkeluar():
    draw = request.args.get('draw', type=int)
    start = request.args.get('start', type=int)
    length = request.args.get('length', type=int)
    search_value = request.args.get('search[value]', type=str)
    query = """
        SELECT barang_keluar.id, barang_keluar.nomerfaktur, barang_keluar.tglfaktur,
               sales.nama_sales, outlet.nama_outlet
        FROM barang_keluar
        INNER JOIN detail_sales ON detail_sales.id = barang_keluar.id_sales
        INNER JOIN sales ON sales.id = detail_sales.id_sales
        INNER JOIN outlet ON outlet.id = detail_sales.id_outlet
    """

    params = []

    if search_value:
        query += " WHERE nomerfaktur LIKE %s OR sales.nama_sales LIKE %s"
        like_value = f"%{search_value}%"
        params.extend([like_value, like_value])

    query += " ORDER BY barang_keluar.id DESC LIMIT %s OFFSET %s"
    params.extend([length, start])

    data = fetch(query, params)


    total_query = "SELECT COUNT(*) FROM barang_keluar"
    g.con.execute(total_query)
    total_records = g.con.fetchone()[0]
    return jsonify({
        'draw': draw,
        'recordsTotal': total_records,
        'recordsFiltered': total_records,  # Bisa disesuaikan kalau pakai filter
        'data': data
    })

def to_decimal(val):
    try:
        # pastikan val string dulu
        val = str(val).replace('.', '').replace(',', '.')
        return Decimal(val)
    except (InvalidOperation, TypeError, ValueError):
        return Decimal('0.00')
@app.route('/admin/pengeluaran/tambah/id', methods=['POST'])
@jwt_required()
def tambah_id_pengeluaran_new():
    data = request.json
    print(data)
    performasales = Decimal('0.00')  # Awal Pastikan ini Decimal, bukan float!
    # Calculate Jatuh Tempo and No Faktur
    tanggal = time_zone_wib()

    nofaktur = data["nofaktur"]# Gabungkan tanggal dengan nomor urut (001, 002, ...)

    for item in data['items']:
        performasales += Decimal(item['harga_total'])
    print(performasales)
    # if str(data.get('pajak', '')).lower() == 'yes':
        # pajak_rate = decimal('11') / decimal('100')
        # pajak_amount = pajak_rate * performasales
        # performasales += pajak_amount  # sekarang ppn ditambahkan dengan benar

    g.con.execute("""
        SELECT detail_sales.id FROM detail_sales 
        INNER JOIN sales ON sales.id = detail_sales.id_sales 
        INNER JOIN outlet ON outlet.id = detail_sales.id_outlet 
        WHERE sales.nama_sales = %s AND outlet.nama_outlet = %s
    """, (data['nama_sales'], data['nama_outlet']))
    id_sales = g.con.fetchone()
    if not id_sales:
        return jsonify({"error": "Sales not found"}), 404
    g.con.execute("""
    INSERT INTO barang_keluar (tglfaktur, jatuhtempo, nomerfaktur, id_sales,cashtempo, pajak)
    VALUES (%s, %s, %s, %s, %s, %s)""", 
    (data['tglfaktur'], data['jthtempo'], data['nofaktur'], id_sales[0], data['pembayaran'], data['pajak']))
    new_id = g.con.lastrowid
    print(data['items'])
    #cek barang
    for i in data['items']:
        print(i)
        print(i['nama_barang'])
        print(i['id_barang'])
        print(i['harga_satuan'])
        print(i['hpp'])
        g.con.execute("SELECT sisa_gudang FROM barang_gudang WHERE id_barang = %s", (i['id_barang'],))
        result = g.con.fetchone()
        g.con.execute("SELECT stoklimit FROM barang WHERE id = %s", (i['id_barang'],))
        stoklimit = g.con.fetchone()
        if result:
            g.con.execute("""
            INSERT INTO detail_barang_keluar (id_barang_keluar, id_barang, jmlpermintaan, harga_satuan, diskon, harga_total, 
            cn, hpp, totalhpp, profit, batch, ed, lunas_or_no ) 
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) """,
            (new_id, i['id_barang'], i['jmlpermintaan'], i['harga_satuan'], i['diskon'], i['harga_total'], i['cn'], i['hpp'], i['totalhpp'], i['profit'], 
            i['batch'],i['ed'], "Tidak Lunas" ))
            print(result[0])
            new_qty = result[0] - int(i['jmlpermintaan'])
            if new_qty < 0 :
                return jsonify({"error": "Stok Kurang"}), 404
            elif new_qty <= stoklimit[0]:
                ket = "Stok Tidak Aman"
            else :
                ket = "Stok Aman"
                g.con.execute("""
                    UPDATE barang_gudang SET sisa_gudang = %s, keterangan = %s WHERE id_barang = %s
                """, (new_qty, ket, i['id_barang']))
        else:
            return jsonify({"error": "Barang tidak ditemukan di gudang"}), 404
    
    g.con.connection.commit()
    return jsonify({"msg": "SUKSES"})
def awal(pdf,width,height,barang_keluar,ada_diskon,ada_batch,ada_ed, cash_tempo):
        # Set margin awal
        x_margin = 0.75 * cm
        y = height - 1 * cm
        y2 = y
        # Header
        pdf.setFont("Helvetica-Bold", 12)
        pdf.drawString(x_margin, y - 0 * cm, "PT. BREGAS SAPORETE MEDIKALINDO")
        pdf.setFont("Helvetica", 8 )
        pdf.drawString(x_margin, y - 0.6 * cm, "Jl. Pati Perum Vero Permai No.39 Margadana, Kota Tegal")
        pdf.drawString(x_margin, y - 1 * cm, "Telp: 02833170260 | email : bregassaporetemedikalindo@gmail.com ") 
        pdf.drawString(x_margin, y - 1.4 * cm, "No Rek: 0101-01-003113-30-5 An PT Bregas Saporete Medikalindo(BRI)")
        pdf.drawString(x_margin, y - 1.8 * cm, "Izin CDAKB: PB-UMKU 912010045181900010001")
        y -= 2.4 * cm
        x_margin_9 = x_margin + 9.6 * cm   
        # Garis pemisah
        pdf.line(x_margin, y, width - x_margin, y)
        y -= 0.5 * cm

        # Info Faktur dan "Kepada Yth" (samping-sampingan)
        tanggal = barang_keluar[0].strftime('%d/%m/%Y')
        jth_tempo = barang_keluar[1].strftime('%d/%m/%Y')

        pdf.setFont("Helvetica-Bold", 12)
        pdf.drawString(x_margin_9, y2, "FAKTUR")
        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(x_margin_9 + 4.2 * cm, y2, "Kepada Yth:")

        # Data Faktur di sebelah kiri
        pdf.setFont("Helvetica", 8)
        y2 -= 0.2 * cm
        pdf.drawString(x_margin_9, y2 - 0.4 * cm, "No Faktur")
        pdf.drawString(x_margin_9 + 1.9 * cm, y2 - 0.4 * cm, ": "+barang_keluar[2])
        pdf.drawString(x_margin_9, y2 - 0.8 * cm, "Inv Date")
        pdf.drawString(x_margin_9 + 1.9 * cm, y2 - 0.8 * cm, ": "+tanggal)
        pdf.drawString(x_margin_9, y2 - 1.2 * cm, "Sales")
        print(barang_keluar[3])
        pdf.drawString(x_margin_9 + 1.9 * cm, y2 - 1.2 * cm, ": "+barang_keluar[3])
        pdf.drawString(x_margin_9, y2 - 1.6 * cm, "Jatuh Tempo")
        pdf.drawString(x_margin_9 + 1.9 * cm, y2 - 1.6 * cm, ": "+jth_tempo)
        pdf.drawString(x_margin_9, y2 - 2 * cm, "Cash / Tempo")
        pdf.drawString(x_margin_9 + 1.9 * cm, y2 - 2 * cm, ": "+cash_tempo)
        
        # Data 'Kepada Yth' dan Alamat di sebelah kanan
        alamat = barang_keluar[5]
        alamat_lines = textwrap.wrap(alamat, width=50)
        pdf.drawString(x_margin_9 + 4.2 * cm, y2 - 0.4 * cm, barang_keluar[4])  # Nama outlet
        y3 = y2 
        for i, line in enumerate(alamat_lines[:11]):  # Limit to 6 lines
            y3 = y2 - (0.8 + i * 0.4) * cm
            pdf.drawString(x_margin_9 + 4.2 * cm, y2 - (0.8 + i * 0.4) * cm, line)  # Alamat outlet
        pdf.drawString(x_margin_9 + 4.2 * cm, y3 - 0.4 * cm, f"NPWP: {barang_keluar[6]}")  # Npwp outlet
        # Tabel barang
        pdf.setFont("Helvetica-Bold", 9)
        headers = ["No", "Nama Barang", "Qty"]
        header_positions = [x_margin, x_margin + 0.6 * cm, x_margin + 9 * cm]
        
        if ada_batch == True :
            header_positions[2] =  x_margin + 7 * cm
            headers.append("Batch")
            if ada_ed == True :
                header_positions[2] =  x_margin + 6 * cm
                header_positions.append(x_margin + 7.5 * cm)
            else:
                header_positions.append(x_margin + 8 * cm)  
        if ada_ed == True:
            headers.append("ED")
            if ada_batch == True :
                header_positions.append(x_margin + 9.5 * cm)  
            else:
                header_positions[2] =  x_margin + 7 * cm
                header_positions.append(x_margin + 8 * cm)  
        headers.append("Harga")
        header_positions.append(x_margin + 11 * cm) 
        if ada_diskon == True:
            headers.append("Diskon")
            header_positions.append(x_margin + 14.5 * cm)
        headers.append("Total")
        header_positions.append(x_margin + 16 * cm)
        for pos, header in zip(header_positions, headers):
            pdf.drawString(pos, y, header)
        y -= 0.3 * cm
        pdf.line(x_margin, y, width - x_margin, y)
        y -= 0.5 * cm

        pdf.setFont("Helvetica", 9)
        return y, x_margin,width, header_positions
    
    # footer total
def hitung_total(pdf, y, x_margin, width, jumlah_total, pajak, page_num=1, total_pages=1, ada_diskon=False ):
    if page_num == total_pages:
        pdf.line(x_margin, y, width - x_margin, y)
        y -= 0.4 * cm
        if pajak == "pajak":
            pdf.setFont("Helvetica", 8)
            if ada_diskon == True:
                pdf.drawString(x_margin + 15.975 * cm, y, f"Rp ")
                pdf.drawString(x_margin + 15.975 * cm, y - 0.4 * cm, f"Rp ")  
                pdf.drawString(x_margin + 15.975 * cm, y - 0.8 * cm, f"Rp ")   
            else:
                pdf.drawString(x_margin + 16 * cm, y, f"Rp ")
                pdf.drawString(x_margin + 16 * cm, y - 0.4 * cm, f"Rp ")  
                pdf.drawString(x_margin + 16 * cm, y - 0.8 * cm, f"Rp ")     
            pdf.drawString(x_margin + 14 * cm, y, f"Total: ")
            pdf.drawRightString(x_margin + 19.6 * cm, y, f"{format_rupiah(jumlah_total)}")
            pajak_val = Decimal('0.11') * jumlah_total
            pajak_val = pajak_val.quantize(Decimal('1'), rounding=ROUND_HALF_UP) # pembulatan
            pdf.drawString(x_margin + 14 * cm, y - 0.4 * cm, f"PPN 11%: ")  
            pdf.drawRightString(x_margin + 19.6 * cm, y - 0.4 * cm, f"{format_rupiah(pajak_val)}")
            grand_total = jumlah_total + pajak_val
            terbilang_rupiah = num2words(int(grand_total), lang='id')
            terbilang_rupiah = terbilang_rupiah.title() 
            terbilang_rupiah += " Rupiah"
            pdf.setFont("Helvetica-Bold", 8)
            y3 = y
            if len(terbilang_rupiah) >= 90:
                terbilang_rupiah_lines = textwrap.wrap(terbilang_rupiah, width=75)
                for i, line in enumerate(terbilang_rupiah_lines[:3]):
                    y3 -=  0.4 * cm
                    pdf.drawString(x_margin ,y - (i * 0.4) * cm, line)
            else:
                pdf.drawString(x_margin, y , terbilang_rupiah)
            pdf.drawString(x_margin + 14 * cm, y - 0.8 * cm, f"Grand Total: ")
            
            pdf.drawRightString(x_margin + 19.6 * cm, y - 0.8 * cm, f"{format_rupiah(grand_total)}")

            # Footer
            pdf.drawString(x_margin, y3 - 0.5 * cm, "Penerima")
            pdf.drawString(x_margin + 4.5 * cm, y3 - 0.5 * cm, "Gudang")
            pdf.drawString(x_margin + 8.5 * cm, y3 - 0.5 * cm, "Expedisi")
        else:
            pdf.setFont("Helvetica-Bold", 10)
            terbilang_rupiah = num2words(jumlah_total, lang='id')
            terbilang_rupiah = terbilang_rupiah.title() 
            terbilang_rupiah += " Rupiah"
            y3 = y
            if len(terbilang_rupiah) >= 90:
                terbilang_rupiah_lines = textwrap.wrap(terbilang_rupiah, width=75)
                for i, line in enumerate(terbilang_rupiah_lines[:3]):
                    y3 -=  0.4 * cm
                    pdf.drawString(x_margin ,y - (i * 0.4) * cm, line)
            else:
                pdf.drawString(x_margin, y , terbilang_rupiah)
            if ada_diskon == True:
                pdf.drawString(x_margin + 14.9 * cm, y, f"Total: Rp ")
            else:
                pdf.drawString(x_margin + 14.92 * cm, y, f"Total: Rp ")
            pdf.drawRightString(x_margin + 19.6 * cm, y, f"{format_rupiah(jumlah_total)}")
            y3 -= 0.5 * cm

            # Footer
            pdf.drawString(x_margin, y3, "Penerima")
            pdf.drawString(x_margin + 4.5 * cm, y3, "Gudang")
            pdf.drawString(x_margin + 8.5 * cm, y3, "Expedisi")
    else:
        pdf.line(x_margin, y, width - x_margin, y)
        y -= 0.4 * cm
        if pajak == "pajak":
            pdf.setFont("Helvetica-Bold", 8)
            y3 = y

            # Footer
            pdf.drawString(x_margin, y3 - 0.5 * cm, "Penerima")
            pdf.drawString(x_margin + 4.5 * cm, y3 - 0.5 * cm, "Gudang")
            pdf.drawString(x_margin + 8.5 * cm, y3 - 0.5 * cm, "Expedisi")
        else:
            pdf.setFont("Helvetica-Bold", 10)
            y3 = y
            y3 -= 0.5 * cm

            # Footer
            pdf.drawString(x_margin, y3, "Penerima")
            pdf.drawString(x_margin + 4.5 * cm, y3, "Gudang")
            pdf.drawString(x_margin + 8.5 * cm, y3, "Expedisi")

def export_pdf(buffer, pajak):
    id = request.args.get("id")
    g.con.execute("""SELECT tglfaktur, jatuhtempo, nomerfaktur, sales.nama_sales, outlet.nama_outlet, 
    outlet.alamat_outlet, outlet.npwp, barang_keluar.cashtempo FROM barang_keluar 
    INNER JOIN detail_sales on detail_sales.id = barang_keluar.id_sales 
    INNER JOIN sales on sales.id = detail_sales.id_sales 
    INNER JOIN outlet on outlet.id = detail_sales.id_outlet WHERE barang_keluar.id = %s """, (id,))
    barang_keluar = g.con.fetchone()
    g.con.execute("SELECT barang.nama_barang, jmlpermintaan, barang.qty, harga_satuan, harga_total, diskon, batch, ed " \
    "FROM detail_barang_keluar INNER JOIN barang on barang.id = detail_barang_keluar.id_barang WHERE id_barang_keluar = %s ", (id,))
    data_db = g.con.fetchall()
    data = []
    jumlah_total = 0
    ada_diskon = False
    ada_batch = False
    ada_ed = False
    for i in range(len(data_db)):
        item = {"No":i+1, "Nama Barang":data_db[i][0],"Quantity":str(data_db[i][1])+" "+data_db[i][2],
                "Harga Satuan": format_currency(data_db[i][3]),"Total":format_currency(data_db[i][4])}
        data.append(item)
        jumlah_total += data_db[i][4]
        if data_db[i][5] and data_db[i][5] != '' and data_db[i][5] != 0:
            ada_diskon = True
        if data_db[i][6] and data_db[i][6] != '' and data_db[i][6] != 0:
            ada_batch = True
        if data_db[i][7] and data_db[i][7] != '' and data_db[i][7] != 0:
            ada_ed = True

    # Pagination logic
    max_rows_per_page = 12
    total_rows = max(len(data_db), max_rows_per_page)
    total_pages = (total_rows + max_rows_per_page - 1) // max_rows_per_page

    pdf = canvas.Canvas(buffer, pagesize=(21.6 * cm, 14.5 * cm))
    width, height = (21.6 * cm, 14.5 * cm)
    cash_tempo = barang_keluar[7]
    y, x_margin, width, header_positions = awal(
                pdf, width, height, barang_keluar, ada_diskon, ada_batch, ada_ed, cash_tempo
            )

    jumlah_total = 0
    row_idx = 0
    page_num = 1

    for i in range(total_rows):
        if (i > 0 and i % max_rows_per_page == 0):
            hitung_total(pdf, y, x_margin, width, jumlah_total, pajak, page_num, total_pages,ada_diskon)
            pdf.showPage()
            page_num += 1
            y, x_margin, width, header_positions = awal(
                pdf, width, height, barang_keluar, ada_diskon, ada_batch, ada_ed, cash_tempo
            )

        if i < len(data_db):
            nama_barang = data_db[i][0]
            jml = data_db[i][1]
            satuan = data_db[i][2]
            harga_satuan = data_db[i][3]
            total = data_db[i][4]
            diskon = data_db[i][5] if ada_diskon else None
            batch = data_db[i][6] if ada_batch else None
            ed = data_db[i][7] if ada_ed else None

            jumlah_total += total
            qty = f"{jml} {satuan}"

            values = [str(i + 1), nama_barang, qty]

            if ada_batch:
                values.append(str(batch) if batch else "-")
            if ada_ed:
                values.append(str(ed) if ed else "-")

            values.append(f"satuan {format_rupiah(harga_satuan)}")

            if ada_diskon:
                values.append(f"{diskon}%")
            
            values.append(f"total {format_rupiah(total)}")

        else:
            # Isi kosong untuk baris tidak terpakai
            values = [str(i + 1), "-", "-"]
            if ada_batch:
                values.append("-")
            if ada_ed:
                values.append("-")
            values.append("-")  # harga
            if ada_diskon:
                values.append("-")
            values.append("-")  # total
        # Gambar data ke PDF
        for pos, value in zip(header_positions, values):
            if value.startswith("satuan"):
                pdf.drawString(pos, y, "Rp ")
                pdf.drawRightString(pos + (3.2 * cm), y, value[7:].strip())
            elif value.startswith("total"):
                pdf.drawString(pos, y, "Rp ")
                pdf.drawRightString(x_margin + 19.6 * cm, y, value[6:].strip())
            else:
                pdf.drawString(pos, y, value)
        y -= 0.5 * cm

    print(jumlah_total)
    # Total
    hitung_total(pdf, y, x_margin,width, jumlah_total, pajak, page_num, total_pages, ada_diskon)
    #return Response(generate(), mimetype='application/pdf')
    pdf.save()

@app.route('/admin/pengeluaran/print', methods=['GET'])
def print_pdf():
    buffer = BytesIO()
    export_pdf(buffer,"")
    buffer.seek(0)
    print(f"PDF size: {len(buffer.getvalue())} bytes")  # Debug ukuran file
    def generate():
        yield buffer.read()
    return send_file(buffer,as_attachment=True,download_name="faktur.pdf",mimetype='application/pdf')

@app.route('/admin/pengeluaran/print_pajak', methods=['GET'])
def print_pdf_pajak():
    buffer = BytesIO()
    export_pdf(buffer,"pajak")
    buffer.seek(0)
    print(f"PDF size: {len(buffer.getvalue())} bytes")  # Debug ukuran file
    def generate():
        yield buffer.read()
    return send_file(buffer,as_attachment=True,download_name="faktur+pajak.pdf",mimetype='application/pdf')

@app.route('/admin/pengeluaran/edit/<int:id>', methods=['GET'])
def edit_id_pengeluaran_new(id):
    # Ambil data barang_keluar
    g.con.execute("""
        SELECT barang_keluar.id, barang_keluar.id_sales, barang_keluar.tglfaktur, barang_keluar.jatuhtempo, barang_keluar.nomerfaktur, 
        barang_keluar.cashtempo, barang_keluar.pajak, sales.nama_sales, outlet.nama_outlet, outlet.alamat_outlet
        FROM barang_keluar
        INNER JOIN detail_sales ON detail_sales.id = barang_keluar.id_sales
        INNER JOIN sales ON sales.id = detail_sales.id_sales
        INNER JOIN outlet ON outlet.id = detail_sales.id_outlet
        WHERE barang_keluar.id = %s
    """, (id,))
    barang_keluar = g.con.fetchone()
    print(barang_keluar[2])
    print(barang_keluar[3])
    # Ambil detail_barang_keluar
    g.con.execute("""
        SELECT dbk.id, dbk.id_barang_keluar, dbk.id_barang, dbk.jmlpermintaan,dbk.harga_satuan,dbk.diskon,
        dbk.harga_total,dbk.cn,dbk.hpp,dbk.totalhpp,dbk.profit,dbk.tanggal_pembayaran,
        dbk.metode_pembayaran,dbk.lunas_or_no, barang.nama_barang, barang.kode_barang, dbk.batch, dbk.ed
        FROM detail_barang_keluar dbk
        INNER JOIN barang ON barang.id = dbk.id_barang
        WHERE dbk.id_barang_keluar = %s
    """, (id,))
    detail_barang_keluar = g.con.fetchall()
    # Ambil data sales dan outlet untuk dropdown (jika diperlukan)
    data_sales = fetch("""SELECT * FROM detail_sales 
    INNER JOIN sales ON sales.id = detail_sales.id_sales 
    INNER JOIN outlet ON outlet.id = detail_sales.id_outlet ORDER BY detail_sales.id""")
    
    nama_sales = fetch(""" SELECT DISTINCT nama_sales FROM sales """)
    data_outlet = fetch("""
        SELECT outlet.id, outlet.nama_outlet, outlet.alamat_outlet FROM outlet
    """)
    # Fetch data_barang for the template
    data_barang = fetch("""SELECT barang.id, barang.nama_barang,barang.qty, barang_gudang.sisa_gudang, barang.stoklimit 
    FROM barang INNER JOIN barang_gudang on barang_gudang.id_barang = barang.id ORDER BY barang_gudang.id""")
    
    return render_template(
        "admin/edit_pengeluaran.html",
        barang_keluar=barang_keluar,
        detail_barang_keluar=detail_barang_keluar,
        data_sales=data_sales,
        nama_sales = nama_sales,
        data_outlet=data_outlet,
        data_barang=data_barang
    )

@app.route('/admin/pengeluaran/edit/<int:id>', methods=['PUT'])
@jwt_required()
def edit_id_pengeluaran_action(id):
    # PUT method
    form_data = request.json
    print(form_data)
    id_barang_keluar = id
    nama_sales = form_data['nama_sales']
    nama_outlet = form_data['nama_outlet']
    g.con.execute("""
        SELECT detail_sales.id FROM detail_sales 
        INNER JOIN sales ON sales.id = detail_sales.id_sales 
        INNER JOIN outlet ON outlet.id = detail_sales.id_outlet 
        WHERE sales.nama_sales = %s AND outlet.nama_outlet = %s
    """, (nama_sales, nama_outlet))
    id_sales = g.con.fetchone()
    if not id_sales:
        return jsonify({"error": "Sales not found"}), 404
    performasales=0
    for item in form_data['items']:
        print(item.get('harga_total', '0'))
        harga_total = Decimal(str(item.get('harga_total', '0')))
        performasales += harga_total

    # if str(form_data.get('pajak', '')).lower() == 'yes':
        # pajak_rate = decimal('11') / decimal('100')
        # pajak_amount = pajak_rate * performasales
        # performasales += pajak_amount  # sekarang ppn ditambahkan dengan benar
    # Update nomerfaktur di tabel barang_keluar
    g.con.execute("""
        UPDATE barang_keluar 
        SET nomerfaktur = %s, tglfaktur = %s, jatuhtempo = %s, id_sales = %s, cashtempo = %s, pajak = %s
        WHERE id = %s
    """, (
        form_data['nofaktur'],
        form_data['tglfaktur'],
        form_data['jthtempo'],
        id_sales[0],
        form_data['pembayaran'],
        form_data.get('pajak', ''),
        int(id_barang_keluar)
    ))
    print(form_data['nofaktur'])
    # Update detail_barang_keluar sesuai data user
    items = form_data['items']
    #hapus stok lama dari gudang
    for i in items:
        g.con.execute("SELECT jmlpermintaan FROM detail_barang_keluar WHERE id_barang_keluar = %s AND id_barang = %s", (int(id_barang_keluar), i['id_barang']))
        result = g.con.fetchone()
        if result:
            update_stok_gudang(i['id_barang'],result[0],'kurangi')
    # Hapus semua detail_barang_keluar yang terkait dengan id_barang_keluar ini
    g.con.execute("DELETE FROM detail_barang_keluar WHERE id_barang_keluar = %s", (int(id_barang_keluar),))

    g.con.connection.commit()
    # Insert data baru dari user
    fields = [
        'id_barang', 'jmlpermintaan', 'harga_satuan', 'diskon', 'harga_total',
        'cn', 'hpp', 'totalhpp', 'profit', 'batch','ed'
    ]
    for i in items:
        query = f"""INSERT INTO detail_barang_keluar 
            (id_barang_keluar, {', '.join(fields)}) 
            VALUES (%s, {', '.join(['%s'] * len(fields))})"""
        
        values = (int(id_barang_keluar),) + tuple(i.get(field) for field in fields)
        print(query, values)
        g.con.execute(query, values)
        #masukan stok baru
        update_stok_gudang(i['id_barang'],i['jmlpermintaan'],'tambah')
        
    g.con.connection.commit()
    return jsonify({"msg": "SUKSES"})
@app.route('/admin/pengeluaran/hapus/id', methods=['DELETE'])
@jwt_required()
def hapus_pengeluaran():
    try:
        id = request.json.get('id')
        print(id)
        g.con.execute("SELECT id_barang, jmlpermintaan, id FROM detail_barang_keluar WHERE id_barang_keluar = %s", 
        (int(id),))
        data = g.con.fetchall()
        print(data)
        if not data:
            return jsonify({"error": "Data tidak ditemukan"}), 404

        for i in data:
            # Kurangi jumlah di barang_gudang
            g.con.execute("SELECT sisa_gudang FROM barang_gudang WHERE id_barang = %s", (i[0],))
            gudang = g.con.fetchone()
            g.con.execute("SELECT stoklimit FROM barang WHERE id = %s", (i[0],))
            stoklimit = g.con.fetchone()
            if gudang:
                new_jumlah = max(0, gudang[0] + i[1])
                if int(new_jumlah) <= stoklimit[0]:
                    ket = "Stok Tidak Aman"
                else :
                    ket = "Stok Aman"
                g.con.execute("UPDATE barang_gudang SET sisa_gudang = %s, keterangan = %s WHERE id_barang = %s"
                              , (new_jumlah, ket, i[0]))
            # Hapus dari detail_barang_masuk
            g.con.execute("DELETE FROM detail_barang_keluar WHERE id = %s", (i[2],))
        # Hapus dari barang_masuk
        g.con.execute("DELETE FROM barang_keluar WHERE id = %s", (int(id),))
        g.con.connection.commit()
        return jsonify({"msg": "SUKSES"})
    except Exception as e:
        print(str(e))
        return jsonify({"error": str(e)})

@app.route('/admin/keuangan')
def adminkeuangan():
    # Ambil parameter dari request
    tahun = request.args.get('tahun')
    bulan = request.args.get('bulan')
    tanggal = request.args.get('tanggal')
    nama_sales = request.args.get('nama_sales','')
    nama_outlet = request.args.get('nama_outlet','')
    page = request.args.get('page', default=1, type=int)
    per_page = request.args.get('per_page', default=10, type=int)

    
    # Fetch data
    data_sales = fetch("""
        SELECT * FROM detail_sales 
        INNER JOIN sales ON sales.id = detail_sales.id_sales 
        INNER JOIN outlet ON outlet.id = detail_sales.id_outlet 
        ORDER BY detail_sales.id DESC 
    """)
    data_outlet = fetch("""
        SELECT outlet.nama_outlet, outlet.alamat_outlet FROM detail_sales 
        INNER JOIN sales ON sales.id = detail_sales.id_sales 
        INNER JOIN outlet ON outlet.id = detail_sales.id_outlet 
        ORDER BY detail_sales.id DESC
    """)

    # Build query dengan parameter binding
    query = """
        SELECT bk.id AS id_barang_keluar, bk.tglfaktur, bk.nomerfaktur, 
               bk.cashtempo, s.nama_sales, o.nama_outlet, o.alamat_outlet,
               bk.tanggal_pembayaran, bk.keterangan_pembayaran, bk.lunas_tidak
        FROM barang_keluar bk
        INNER JOIN detail_sales ds ON ds.id = bk.id_sales
        INNER JOIN sales s ON s.id = ds.id_sales
        INNER JOIN outlet o ON o.id = ds.id_outlet
    """

    filters, params, count_params = [], [], []
    clause, prms = build_date_range(
            year=int(tahun) if tahun else None,
            month=int(bulan) if bulan else None,
            day=int(tanggal) if tanggal else None,
            alias="bk",              # <- alias sesuai query kamu
            col="tglfaktur"
        )
    if clause:
        filters.append(clause)
        params.extend(prms)
        count_params.extend(prms)
    if nama_sales:
        filters.append("s.nama_sales = %s")
        params.append(nama_sales)
        count_params.append(nama_sales)
    if nama_outlet:
        filters.append("o.nama_outlet = %s")
        params.append(nama_outlet)
        count_params.append(nama_outlet)

    if filters:
        where_clause = " WHERE " + " AND ".join(filters)
        query += where_clause
    else:
        where_clause = ""

    query += " ORDER BY bk.id DESC "

    # Pagination
    offset = (page - 1) * per_page
    query += "LIMIT %s OFFSET %s"
    params.extend([per_page, offset])

    # Hitung total record
    count_query = f"""
        SELECT COUNT(*) FROM barang_keluar bk 
        INNER JOIN detail_sales ds ON ds.id = bk.id_sales
        INNER JOIN sales s ON s.id = ds.id_sales
        INNER JOIN outlet o ON o.id = ds.id_outlet
        {where_clause}
    """
    print(count_query)
    print(count_params)

    g.con.execute(count_query, count_params)
    total_records = g.con.fetchone()[0]
    barang_keluar = fetch(query, params)  # hasil = list of dict
    info_list = []
    if barang_keluar != []:
        data_detail_keluar = fetch("""
        SELECT dbk.id, dbk.id_barang_keluar, barang.nama_barang, dbk.harga_total, 
            dbk.tanggal_pembayaran, dbk.metode_pembayaran, dbk.lunas_or_no
        FROM detail_barang_keluar dbk
        INNER JOIN barang ON barang.id = dbk.id_barang 
        WHERE dbk.id_barang_keluar IN %s
        ORDER BY dbk.id
        """, (tuple([i['id_barang_keluar'] for i in barang_keluar]),))
        for faktur in barang_keluar:
            details = []
            performa_sales =0
            for d in data_detail_keluar:
                if d["id_barang_keluar"] == faktur["id_barang_keluar"]:   # pakai dict style

                    # Cari pembayaran yang match
                    performa_sales += int(d["harga_total"])
                    details.append({
                        "id": d["id"],
                        "nama_barang": d["nama_barang"],
                        "harga_total": d["harga_total"],
                        "lunas_or_no": d["lunas_or_no"]
                    })

            info_list.append({
                "id_barang_keluar": faktur["id_barang_keluar"],
                "tglfaktur": faktur["tglfaktur"],
                "nomerfaktur": faktur["nomerfaktur"],
                "nama_sales": faktur["nama_sales"],
                "nama_outlet": faktur["nama_outlet"],
                "cashtempo": faktur["cashtempo"],
                "keterangan_pembayaran": faktur["keterangan_pembayaran"] or "-",
                "tanggal_pembayaran": faktur["tanggal_pembayaran"] or "",
                "lunas_tidak": faktur["lunas_tidak"] or "-",
                "performa_sales" : performa_sales,
                "detail_items": details
            })

    # Data barang (opsional tambahan)
    data_barang = fetch("SELECT id, kode_barang, nama_barang, qty, stoklimit FROM barang ORDER BY id")

    thn = fetch_years("barang_keluar")
    data_pabrik = fetch("SELECT id, nama_supplier, alamat, tlp FROM pabrik ORDER BY id")
    
    nama_outlet = fetch(""" SELECT DISTINCT nama_outlet FROM outlet  """)
    nama_sales = fetch(""" SELECT DISTINCT nama_sales FROM sales """)
    tanggal_pengeluaran = fetch("SELECT tglfaktur FROM barang_keluar GROUP BY tglfaktur ORDER BY id")

    # Calculate Jatuh Tempo and No Faktur
    tanggal_now = time_zone_wib()
    jth_tempo = (tanggal_now + relativedelta(months=1)).strftime("%Y-%m-%d")
    tanggal_str = tanggal_now.strftime('%Y-%m-%d')
    g.con.execute("SELECT COUNT(*) FROM barang_keluar WHERE tglfaktur = %s", (tanggal_str,))
    count = g.con.fetchone()[0]
    nofaktur = f"{tanggal_str.replace('-', '')}{int(count) + 1:03d}"

    # Pagination info
    total_pages = (total_records + per_page - 1) // per_page
    has_next = page < total_pages
    has_prev = page > 1
    page_range = range(max(1, page - 3), min(total_pages + 1, page + 3))

    # Render template
    return render_pjax(
        "admin/keuangan.html",
        
        info_list=info_list,
        tahun=thn,
        data_sales=data_sales,
        data_outlet=data_outlet,
        nama_outlet=nama_outlet,
        tanggal_pengeluaran=tanggal_pengeluaran,
        data_barang=data_barang,
        data_pabrik=data_pabrik,
        tanggal=tanggal_str,
        nofaktur=nofaktur,
        jth_tempo=jth_tempo,
        nama_sales=nama_sales,
        page=page,
        per_page=per_page,
        total_pages=total_pages,
        total_records=total_records,
        has_next=has_next,
        has_prev=has_prev,
        page_range=page_range
    )

@app.route('/admin/edit-keuangan/<id>', methods=['GET'])
@jwt_required()
def edit_keuangan_new(id):
    print(id)
    return render_template("admin/edit-keuangan.html")
    
@app.route('/admin/keuangan/edit/id', methods=['PUT'])
@jwt_required()
def edit_id_keuangan_new():
    form_data = request.json
    print(form_data)
    id_barang_keluar = form_data['id_barang_keluar']
    nomerfaktur = form_data['nomerfaktur']
    print(nomerfaktur)
    tanggal_pembayaran = form_data['tanggal_pembayaran']
    keterangan_pembayaran = form_data['keterangan_pembayaran']
    lunas_tidak = form_data['lunas_tidak']
    g.con.execute("UPDATE barang_keluar SET nomerfaktur = %s, tanggal_pembayaran = %s, keterangan_pembayaran = %s, lunas_tidak = %s WHERE id = %s", 
                  (nomerfaktur, tanggal_pembayaran, keterangan_pembayaran, lunas_tidak, int(id_barang_keluar),))
    g.con.connection.commit()
    return jsonify({"msg": "SUKSES"})
@app.route('/admin/administrasi')
def adminadministrasi():
    tahun = request.args.get('tahun')
    bulan = request.args.get('bulan')
    tanggal = request.args.get('tanggal')
    page = request.args.get('page', default=1, type=int)
    per_page = request.args.get('per_page', default=10, type=int)
    query = """ SELECT bm.id, bm.tglfaktur, bm.nofaktur, bm.tanggal_pembayaran, 
    bm.keterangan_pembayaran, bm.lunas_tidak, SUM(dbm.harga_total) AS performa_belanja, p.nama_supplier, p.alamat, p.tlp
    from barang_masuk bm
    JOIN detail_barang_masuk dbm ON dbm.id_barang_masuk = bm.id
    join pabrik p on p.id = bm.id_supplier
    """
    filters, params, count_params = [], [], []
    clause, prms = build_date_range(
        year=int(tahun) if tahun else None,
        month=int(bulan) if bulan else None,
        day=int(tanggal) if tanggal else None,
        alias="bm",              # <- alias sesuai query kamu
        col="tglfaktur"
    )
    if clause:
        filters.append(clause)
        params.extend(prms)
        count_params.extend(prms)
    nama_principle = request.args.get('nama_principle',type=str)
    print(nama_principle)
    if nama_principle:
        filters.append(' p.nama_supplier = %s ')
        params.append(nama_principle)
        count_params.append(nama_principle)
    if filters:
        where_clause = " WHERE " + " AND ".join(filters)
        query += where_clause
    else:
        where_clause = ""
    
    # Count query
    count_query = f""" SELECT COUNT(*) FROM barang_masuk bm
    JOIN detail_barang_masuk dbm ON dbm.id_barang_masuk = bm.id
    join pabrik p on p.id = bm.id_supplier
    { where_clause}
    GROUP BY bm.id
    ORDER BY bm.id DESC """
    print(count_query)
    g.con.execute(count_query, count_params)
    total_records = g.con.fetchone()[0]
    
    # Pagination logic
    offset = (page - 1) * per_page
    query += " GROUP BY bm.id ORDER BY bm.id DESC LIMIT %s OFFSET %s"
    params.extend([per_page, offset])

    # get all data with pagination 
    print(query)
    info_list = fetch(query, params)

    data_pabrik = fetch("SELECT id, nama_supplier, alamat, tlp FROM pabrik ORDER BY id")
    
    thn = fetch_years("barang_masuk")
    # Pagination info
    total_pages = (total_records + per_page - 1) // per_page
    has_next = page < total_pages
    has_prev = page > 1
    page_range = range(max(1, page - 3), min(total_pages + 1, page + 3))
    return render_pjax(
        "admin/administrasi.html",
        data_pabrik=data_pabrik,
        info_list=info_list,
        tahun=thn,
        page=page,
        per_page=per_page,
        total_pages=total_pages,
        total_records=total_records,
        has_next=has_next,
        has_prev=has_prev,
        page_range=page_range
    )

@app.route('/admin/administrasi/edit/id', methods=['PUT'])
@jwt_required()
def edit_administrasi():
    form_data = request.json
    fields = [ 'tanggal_pembayaran','keterangan_pembayaran', 'lunas_tidak']
    print(form_data)
    try:
        # Siapkan query UPDATE
        query = f"UPDATE barang_masuk SET {', '.join([f'{field} = %s' for field in fields])} WHERE id = %s"
        values = tuple(form_data.get(field) for field in fields) + (int(form_data.get('id_barang_masuk')),)
        print(values)
        g.con.execute(query, values)
        g.con.connection.commit()
        return jsonify({"msg": "SUKSES"})
    except Exception as e:
        print(str(e))
        return jsonify({"error": str(e)})

@app.route('/admin/administrasi/hapus/id', methods=['DELETE'])
@jwt_required()
def hapus_administrasi():
    form_data = request.get_json()
    try:
        id = form_data['id']
        print(id) 
        g.con.execute("DELETE FROM barang_masuk WHERE id = %s", (id,))
        
        data = fetch("SELECT id, id_barang, jml_menerima  FROM detail_barang_masuk WHERE id_barang_masuk = %s", 
        (id,))
        print(data)
        if not data:
            return jsonify({"error": "Data tidak ditemukan"}), 404
        for i in data:
            print(i)
            _id, id_barang, jumlah = i['id'], i['id_barang'], i['jml_menerima']
            print(_id)
            # Kurangi jumlah di barang_gudang
            g.con.execute("SELECT sisa_gudang FROM barang_gudang WHERE id_barang = %s", (id_barang,))
            sisa_gudang = g.con.fetchone()[0]
            g.con.execute("SELECT stoklimit FROM barang WHERE id = %s", (id_barang,))
            stoklimit = g.con.fetchone()[0]

            if sisa_gudang:
                new_jumlah = max(0, int(sisa_gudang) - jumlah)
                if int(new_jumlah) <= int(stoklimit):
                    ket = "Stok Tidak Aman"
                else :
                    ket = "Stok Aman"
                g.con.execute("UPDATE barang_gudang SET sisa_gudang = %s, keterangan = %s WHERE id_barang = %s", 
                (new_jumlah, ket, id_barang))
        
        g.con.execute("DELETE FROM detail_barang_masuk WHERE id_barang_masuk = %s", (id,))
        g.con.connection.commit()

        return jsonify({"msg": "SUKSES"})
    except Exception as e:
        print(str(e))
        return jsonify({"error": str(e)})
        
@app.route('/admin/barang')
def adminbarang():
    search_id_barang = request.args.get('id_barang')
    page = request.args.get('page', default=1, type=int)
    per_page = request.args.get('per_page', default=10, type=int)
    query = "SELECT id, code as kode_barang, name as nama_barang, unit as qty, stock_min as stoklimit FROM products"
    filters = []
    params = []
    count_params = []

    if search_id_barang:
        filters.append("id = %s")
        params.append(search_id_barang)
        count_params.append(search_id_barang)

    # Tambahkan filter ke query
    if filters:
        where_clause = " WHERE " + " AND ".join(filters)
        query += where_clause
    else:
        where_clause = ""

    # Pagination logic
    # offset = (page - 1) * per_page
    # query += " LIMIT %s OFFSET %s"
    # params.extend([per_page, offset])

    # Query untuk menghitung total records
    count_query = "SELECT COUNT(*) FROM products" + where_clause

    # Eksekusi count query
    g.con.execute(count_query, count_params)
    total_records = g.con.fetchone()[0]

    # Eksekusi query utama
    info_list = fetch(query, params)
    nama_barang = fetch("SELECT MIN(id) AS id, MIN(code) AS kode_barang, name AS nama_barang FROM products GROUP BY nama_barang;")
    total_pages = (total_records + per_page - 1) //per_page
    has_next = page < total_pages
    has_prev = page > 1
    page_range = range(max(1, page - 3), min(total_pages + 1, page + 3))
    
    return render_pjax("admin/barang.html",  info_list=info_list, nama_barang = nama_barang,
        page=page,
        per_page=per_page,
        total_pages=total_pages,
        total_records=total_records,
        has_next=has_next,
        has_prev=has_prev,
        page_range=page_range)
@app.route('/admin/barang', methods=['POST'])
@jwt_required()
def tambah_id_barang_new():
    data = request.get_json() or {}

    # Back-compat mapping
    code       = data.get('code')        or data.get('kode_barang')
    name       = data.get('name')        or data.get('nama_barang')
    unit       = data.get('unit')        or data.get('satuan') or 'pcs'
    stock_min  = data.get('stock_min')   or data.get('stoklimit') or 0
    openingqty = data.get('opening_qty') or data.get('qty') or 0

    if not code or not name:
        return jsonify({"error": "code dan name wajib"}), 400

    try:
        # Cek unik code
        g.con.execute("SELECT id FROM products WHERE code=%s", (code,))
        if g.con.fetchone():
            return jsonify({"error": "Kode barang sudah ada"}), 409

        # Insert product
        g.con.execute(
            "INSERT INTO products (code, name, unit, stock_min) VALUES (%s,%s,%s,%s)",
            (code, name, unit, int(stock_min))
        )
        product_id = g.con.lastrowid

        # Opening stock (optional)
        try:
            openingqty = int(openingqty)
        except Exception:
            openingqty = 0

        if openingqty > 0:
            g.con.execute(
                "INSERT INTO stock_moves (product_id, ref_type, qty_in, note) VALUES (%s,'ADJUSTMENT',%s,%s)",
                (product_id, openingqty, 'Opening balance')
            )

        g.con.connection.commit()
        return jsonify({"msg": "SUKSES", "product_id": product_id})
    except Exception as e:
        g.con.connection.rollback()
        return jsonify({"error": str(e)}), 500

@app.route('/admin/barang', methods=['PUT'])
@jwt_required()
def edit_barang():
    data = request.get_json() or {}
    product_id = data.get('id') or data.get('product_id')
    if not product_id:
        return jsonify({"error": "id/product_id wajib"}), 400

    # Back-compat mapping
    code      = data.get('code')      or data.get('kode_barang')
    name      = data.get('name')      or data.get('nama_barang')
    unit      = data.get('unit')      or data.get('satuan')
    stock_min = data.get('stock_min') or data.get('stoklimit')

    # Target stok (opsional)
    target_qty = data.get('new_qty')
    if target_qty is None:
        target_qty = data.get('qty')  # kompatibel payload lama

    try:
        # Ambil product & stok sekarang
        g.con.execute("SELECT id, code FROM products WHERE id=%s", (product_id,))
        prod = g.con.fetchone()
        if not prod:
            return jsonify({"error": "Produk tidak ditemukan"}), 404

        # Uniqueness code (kalau mau ganti code)
        if code:
            g.con.execute("SELECT id FROM products WHERE code=%s AND id<>%s", (code, product_id))
            if g.con.fetchone():
                return jsonify({"error": "Kode barang sudah dipakai produk lain"}), 409

        # Build UPDATE dinamis
        sets, vals = [], []
        if code is not None:      sets.append("code=%s");      vals.append(code)
        if name is not None:      sets.append("name=%s");      vals.append(name)
        if unit is not None:      sets.append("unit=%s");      vals.append(unit)
        if stock_min is not None: sets.append("stock_min=%s"); vals.append(int(stock_min))

        if sets:
            sql_upd = "UPDATE products SET " + ", ".join(sets) + " WHERE id=%s"
            vals.append(product_id)
            g.con.execute(sql_upd, tuple(vals))

        # Adjustment stok bila diminta
        if target_qty is not None:
            # stok saat ini
            g.con.execute("""
                SELECT COALESCE(s.qty_on_hand,0), p.stock_min
                FROM products p
                LEFT JOIN v_product_stock s ON s.product_id=p.id
                WHERE p.id=%s
            """, (product_id,))
            cur = g.con.fetchone()
            current_qty = int(cur[0] or 0)
            stock_min   = int(cur[1] or 0)

            try:
                target_qty = int(target_qty)
            except Exception:
                return jsonify({"error": "new_qty/qty harus numerik"}), 400

            delta = target_qty - current_qty
            if delta != 0:
                if delta > 0:
                    g.con.execute(
                        "INSERT INTO stock_moves (product_id, ref_type, qty_in,  note) VALUES (%s,'ADJUSTMENT',%s,%s)",
                        (product_id, delta, 'Manual adjustment via edit barang')
                    )
                else:
                    g.con.execute(
                        "INSERT INTO stock_moves (product_id, ref_type, qty_out, note) VALUES (%s,'ADJUSTMENT',%s,%s)",
                        (product_id, abs(delta), 'Manual adjustment via edit barang')
                    )

        g.con.connection.commit()
        return jsonify({"msg": "SUKSES"})
    except Exception as e:
        g.con.connection.rollback()
        return jsonify({"error": str(e)}), 500

@app.route('/admin/barang/<int:product_id>', methods=['DELETE'])
@jwt_required()
def archive_product(product_id):
    # Tolak kalau produk masih punya transaksi aktif
    g.con.execute("""
      SELECT 
        (SELECT COUNT(*) FROM purchase_items WHERE product_id=%s) +
        (SELECT COUNT(*) FROM sales_items    WHERE product_id=%s) +
        (SELECT COUNT(*) FROM stock_moves    WHERE product_id=%s)
    """, (product_id, product_id, product_id))
    used = g.con.fetchone()[0]

    if used > 0:
        # Soft-delete
        g.con.execute("UPDATE products SET is_active=0, archived_at=NOW() WHERE id=%s", (product_id,))
        g.con.connection.commit()
        return jsonify({"msg": "Produk diarsipkan (soft-delete)"}), 200
    else:
        # Boleh hard-delete kalau benar-benar belum dipakai
        g.con.execute("DELETE FROM products WHERE id=%s", (product_id,))
        g.con.connection.commit()
        return jsonify({"msg": "Produk dihapus"}), 200
 
@app.route('/admin/target_sales')
def admintarget_sales():
    id_sales = request.args.get('id_sales')
    tahun = request.args.get('tahun')
    bulan = request.args.get('bulan')

    query = """SELECT target_sales.id, sales.nama_sales, target_sales.target, target_sales.bulan, 
    target_sales.tahun, target_sales.id_sales
    FROM target_sales INNER JOIN sales on sales.id = target_sales.id_sales """
    filters = []
    if tahun:
        filters.append(f"tahun = {tahun}")
    if bulan:
        filters.append(f"bulan = {bulan}")
    if id_sales:
        filters.append(f"id_sales = '{id_sales}'")
    if filters:
        query += " WHERE " + " AND ".join(filters)
    query += " ORDER BY sales.nama_sales asc"
    info_list = fetch(query)
    for i in info_list:
        bulan_dicari = str(int(i['bulan']))  # normalisasi ke format '5', '12', dst
        nama_bulan = next((b['nama_bulan'] for b in list_bulan if b['value'] == bulan_dicari), None)
        i['bulan'] = nama_bulan
    nama_sales = fetch("SELECT id, nama_sales FROM sales GROUP BY nama_sales")
    tahun = fetch("select tahun FROM target_sales GROUP BY tahun desc")
    return render_pjax("admin/target_sales.html",  
    info_list=info_list, nama_sales = nama_sales, tahun=tahun)

@app.route('/admin/target_sales/tambah', methods=['POST'])
@jwt_required()
def tambah_id_target_sales_new():
    form_data = request.get_json()
    fields = ['id_sales','target','tahun','bulan']
    print(form_data)
    query = f"INSERT INTO target_sales ({', '.join(fields)}) VALUES ({', '.join(['%s'] * len(fields))})"

    try:
        values = tuple(form_data[field] for field in fields)
        g.con.execute(query, values)
        g.con.connection.commit()
        return jsonify({"msg": "SUKSES"})
    except Exception as e:
        print("Error saat insert:", str(e))
        return jsonify({"error": str(e)}), 500
    
@app.route('/admin/target_sales/edit/id', methods=['PUT'])
@jwt_required()
def edit_target_sales():
    form_data = request.json
    fields = ['id_sales','target','tahun','bulan']
    
    try:
        id_target_sales = form_data['id'] 

        # Siapkan query UPDATE
        query = f"UPDATE target_sales SET {', '.join([f'{field} = %s' for field in fields])} WHERE id = %s"
        values = tuple(form_data.get(field) for field in fields) + (id_target_sales,)
        print(query)
        print(values)
        g.con.execute(query, values)
        g.con.connection.commit()
        return jsonify({"msg": "SUKSES"})
    except Exception as e:
        print(str(e))
        return jsonify({"error": str(e)})

@app.route('/admin/target_sales/hapus/id', methods=['DELETE'])
@jwt_required()
def hapus_target_sales():
    form_data = request.get_json()
    try:
        id_target_sales = form_data['id']

        query = "DELETE FROM target_sales WHERE id = %s"
        g.con.execute(query, (id_target_sales,))
        g.con.connection.commit()

        return jsonify({"msg": "SUKSES"})
    except Exception as e:
        print(str(e))
        return jsonify({"error": str(e)})
    
@app.route('/admin/pabrik')
def adminpabrik():
    nama_supplier = request.args.get('nama_supplier')
    page = request.args.get('page', default=1, type=int)
    per_page = request.args.get('per_page', default=10, type=int)
    query = "SELECT id, nama_supplier, alamat, tlp FROM pabrik"
    filters = []
    params = []
    count_params = []

    # Tambahkan filter nama_supplier jika tersedia
    if nama_supplier:
        filters.append("nama_supplier = %s")
        params.append(nama_supplier)
        count_params.append(nama_supplier)

    # Tambahkan WHERE jika ada filter
    if filters:
        where_clause = " WHERE " + " AND ".join(filters)
        query += where_clause
    else:
        where_clause = ""

    #Pagination logic
    # offset = (page - 1) * per_page
    # query += " LIMIT %s OFFSET %s"
    # params.extend([per_page, offset])

    # Hitung total record
    count_query = "SELECT COUNT(*) FROM pabrik" + where_clause

    g.con.execute(count_query, count_params)
    total_records = g.con.fetchone()[0]

    # Eksekusi query utama
    info_list = fetch(query, params)
    nama_supplier = fetch("SELECT nama_supplier FROM pabrik GROUP BY nama_supplier")
    total_pages = (total_records + per_page - 1) // per_page
    has_next = page < total_pages
    has_prev = page > 1
    page_range = range(max(1, page - 3), min(total_pages + 1, page + 3))

    return render_pjax("admin/pabrik.html",info_list=info_list, nama_supplier = nama_supplier,                      
        page=page,
        per_page=per_page,
        total_pages=total_pages,
        total_records=total_records,
        has_next=has_next,
        has_prev=has_prev,
        page_range=page_range
        )

@app.route('/admin/pabrik/tambah', methods=['POST'])
@jwt_required()
def tambah_id_pabrik_new():
    form_data = request.get_json()
    fields = ['nama_supplier', 'alamat', 'tlp']
    print(form_data)

    query = f"INSERT INTO pabrik ({', '.join(fields)}) VALUES ({', '.join(['%s'] * len(fields))})"

    try:
        values = tuple(form_data[field] for field in fields)
        g.con.execute(query, values)
        g.con.connection.commit()
        return jsonify({"msg": "SUKSES"})
    except Exception as e:
        print("Error saat insert:", str(e))
        return jsonify({"error": str(e)}), 500
    
@app.route('/admin/pabrik/edit/id', methods=['PUT'])
@jwt_required()
def edit_pabrik():
    form_data = request.json
    fields = ['nama_supplier', 'alamat', 'tlp']
    
    try:
        id_pabrik = form_data['id'] 

        # Siapkan query UPDATE
        query = f"UPDATE pabrik SET {', '.join([f'{field} = %s' for field in fields])} WHERE id = %s"
        values = tuple(form_data.get(field) for field in fields) + (id_pabrik,)
        print(query)
        print(values)
        g.con.execute(query, values)
        g.con.connection.commit()
        return jsonify({"msg": "SUKSES"})
    except Exception as e:
        print(str(e))
        return jsonify({"error": str(e)})

@app.route('/admin/pabrik/hapus/id', methods=['DELETE'])
@jwt_required()
def hapus_pabrik():
    form_data = request.get_json()
    try:
        id_pabrik = form_data['id']

        query = "DELETE FROM pabrik WHERE id = %s"
        g.con.execute(query, (id_pabrik,))
        g.con.connection.commit()

        return jsonify({"msg": "SUKSES"})
    except Exception as e:
        print(str(e))
        return jsonify({"error": str(e)})

@app.route('/admin/sales')
def adminsales():
    # Ambil parameter dari request
    nama_sales = request.args.get('nama_sales')
    nama_outlet = request.args.get('nama_outlet')
    alamat_outlet = request.args.get('alamat_outlet')
    # Query utama untuk mengambil data sales
    query = """
        SELECT detail_sales.id as id, detail_sales.id_sales, detail_sales.id_outlet, 
               sales.nama_sales, outlet.nama_outlet, outlet.alamat_outlet, outlet.npwp
        FROM detail_sales
        INNER JOIN sales ON sales.id = detail_sales.id_sales
        INNER JOIN outlet ON outlet.id = detail_sales.id_outlet
    """

    # Buat filter dan parameter
    filters = []
    params = []

    if nama_sales:
        filters.append("sales.nama_sales = %s")
        params.append(nama_sales)
    if nama_outlet:
        filters.append("outlet.nama_outlet = %s")
        params.append(nama_outlet)
    if alamat_outlet:
        filters.append("outlet.alamat_outlet = %s")
        params.append(alamat_outlet)

    # Tambahkan filter ke query jika ada
    if filters:
        query += " WHERE " + " AND ".join(filters)

    query += " ORDER BY sales.nama_sales ASC"

    # Eksekusi query dengan parameter binding
    info_list = fetch(query, params)

    # Ambil data untuk dropdown filter
    nama_sales = fetch(""" SELECT DISTINCT nama_sales FROM sales """)
    nama_outlet = fetch("""
        SELECT  DISTINCT nama_outlet FROM outlet
    """)
    alamat_outlet = fetch("""
        SELECT DISTINCT alamat_outlet FROM outlet 
    """)
    # Render template dengan data
    return render_pjax(
        "admin/sales.html",
        
        info_list=info_list,
        nama_sales=nama_sales,
        nama_outlet=nama_outlet,
        alamat_outlet=alamat_outlet
    )

@app.route('/admin/sales/tambah', methods=['POST'])
@jwt_required()
def tambah_id_sales_new():
    form_data = request.get_json()
    print(form_data)
    g.con.execute("""select id, nama_sales FROM sales WHERE nama_sales = %s """,(form_data['nama_sales'],))
    result = g.con.fetchone()
    if result:
        id_sales = result[0]
    else:
        g.con.execute(""" insert into sales(nama_sales) values (%s)""",(form_data['nama_sales'],))
        id_sales = g.con.lastrowid
        g.con.connection.commit()
        
    g.con.execute("""select id, nama_outlet FROM outlet WHERE nama_outlet= %s """,(form_data['nama_outlet'],))
    result = g.con.fetchone()
    if result:
        id_outlet = result[0]
    else:
        g.con.execute(""" insert into outlet(nama_outlet, alamat_outlet, npwp) values (%s,%s,%s)""",
        (form_data['nama_outlet'],form_data['alamat_outlet'],form_data['npwp']))
        id_outlet = g.con.lastrowid
        g.con.connection.commit()
    try:
        g.con.execute(""" insert into detail_sales(id_sales,id_outlet) values(%s,%s) """,(id_sales, id_outlet))
        g.con.connection.commit()
        return jsonify({"msg": "SUKSES"})
    except Exception as e:
        print("Error saat insert:", str(e))
        return jsonify({"error": str(e)}), 500

@app.route('/admin/sales/edit/id', methods=['PUT'])
@jwt_required()
def edit_sales():
    form_data = request.json

    try:
        id_detail_sales = form_data.get('id')
        print(id_detail_sales)
        # Ambil data detail_sales saat ini
        g.con.execute("SELECT id_sales, id_outlet FROM detail_sales WHERE id = %s", (id_detail_sales,))
        detail = g.con.fetchone()
        if not detail:
            return jsonify({"error": "Detail sales tidak ditemukan"}), 404

        current_id_sales = detail[0]
        current_id_outlet = detail[1]

        # ==== HANDLE NAMA SALES ====
        g.con.execute("SELECT nama_sales FROM sales WHERE id = %s", (current_id_sales,))
        current_sales = g.con.fetchone()
        if not current_sales:
            return jsonify({"error": "Sales tidak ditemukan"}), 404

        nama_sales_baru = form_data.get('nama_sales')
        if nama_sales_baru != current_sales[0]:
            g.con.execute("SELECT id FROM sales WHERE nama_sales = %s", (nama_sales_baru,))
            existing_sales = g.con.fetchone()
            if existing_sales:
                new_id_sales = existing_sales[0]
            else:
                g.con.execute("INSERT INTO sales (nama_sales) VALUES (%s)", (nama_sales_baru,))
                new_id_sales = g.con.lastrowid
                g.con.connection.commit()
        else:
            new_id_sales = current_id_sales  # Tidak berubah

        # ==== HANDLE OUTLET ====
        g.con.execute("SELECT nama_outlet, alamat_outlet, npwp FROM outlet WHERE id = %s", (current_id_outlet,))
        current_outlet = g.con.fetchone()
        if not current_outlet:
            return jsonify({"error": "Outlet tidak ditemukan"}), 404

        nama_outlet_baru = form_data.get('nama_outlet')
        alamat_outlet_baru = form_data.get('alamat_outlet')
        npwp_baru = form_data.get('npwp')

        if (nama_outlet_baru != current_outlet[0]) or (alamat_outlet_baru != current_outlet[1] ) or (npwp_baru != current_outlet[2]):
            g.con.execute(
                "SELECT id FROM outlet WHERE nama_outlet = %s AND alamat_outlet = %s AND npwp = %s",
                (nama_outlet_baru, alamat_outlet_baru, npwp_baru)
            )
            existing_outlet = g.con.fetchone()
            if existing_outlet:
                new_id_outlet = existing_outlet[0]
            else:
                g.con.execute(
                    "INSERT INTO outlet (nama_outlet, alamat_outlet, npwp) VALUES (%s, %s, %s)",
                    (nama_outlet_baru, alamat_outlet_baru, npwp_baru)
                )
                new_id_outlet = g.con.lastrowid
                g.con.connection.commit()
        else:
            new_id_outlet = current_id_outlet  # Tidak berubah

        # ==== UPDATE detail_sales ====
        g.con.execute(
            "UPDATE detail_sales SET id_sales = %s, id_outlet = %s WHERE id = %s",
            (new_id_sales, new_id_outlet, id_detail_sales)
        )
        g.con.connection.commit()

        return jsonify({"msg": "SUKSES"}), 200

    except Exception as e:
        print("Error:", str(e))
        return jsonify({"error": str(e)}), 500

@app.route('/admin/sales/hapus/id', methods=['DELETE'])
@jwt_required()
def hapus_sales():
    form_data = request.get_json()
    try:
        id_detail_sales = form_data['id']
        print(id_detail_sales)
        # Hapus data dari tabel detail_sales
        
        g.con.execute("SELECT id, id_sales, id_outlet FROM detail_sales WHERE id = %s", (id_detail_sales,))
        detail_sales = g.con.fetchone()
        print(detail_sales)
        g.con.execute("DELETE FROM detail_sales WHERE id = %s", (id_detail_sales,))
        g.con.execute("SELECT id_sales FROM detail_sales WHERE id_sales = %s", (detail_sales[1],))
        result= g.con.fetchall()
        print(result)
        if result:
            pass
        else:
            g.con.execute("DELETE FROM sales WHERE id = %s", (detail_sales[1],))
            g.con.connection.commit()
        g.con.execute("SELECT id_outlet FROM detail_sales WHERE id_outlet  = %s", (detail_sales[2],))
        result= g.con.fetchall()
        print(result)
        if result:
            pass
        else:
            g.con.execute("DELETE FROM outlet WHERE id = %s", (detail_sales[2],))
            g.con.connection.commit()

        g.con.connection.commit()
        return jsonify({"msg": "SUKSES"})
    except Exception as e:
        print(str(e))
        return jsonify({"error": str(e)}), 500
def query_performa(tahun: int | None, bulan: int | None, sales_name: str | None, lunas_tidak: str | None):
    where, params = [], []
    # Pastikan build_date_range pakai kolom header (bk.tglfaktur)
    date_sql, date_params = build_date_range(tahun, bulan, None, alias="bk", col="tglfaktur")
    
    if date_sql:
        where.append(date_sql)
        params.extend(date_params)

    if sales_name:
        # case-insensitive + trim
        where.append("UPPER(TRIM(s.nama_sales)) = UPPER(TRIM(%s))")
        params.append(sales_name)
    # Normalisasi status
    st = (lunas_tidak or '').strip().casefold()
    print(st)
    if st == 'lunas':
        nilai_case = """
            CASE
              WHEN (bk.lunas_tidak = 'Lunas')
              THEN SUM(dbk.harga_total)
            END
        """
    elif st in ('tidak lunas', 'tdk lunas', 'belum lunas'):
        nilai_case = """
            CASE
              WHEN (bk.lunas_tidak != 'Lunas'  )
              THEN SUM(dbk.harga_total)
            END
        """
    else:
        # Tanpa filter: total nilai (lunas + tidak lunas)
        nilai_case = """
            SUM(dbk.harga_total)
        """

    sql = f"""
    SELECT
      x.id_sales, x.nama_sales, x.nama_outlet, x.tahun, x.bulan,
      SUM(x.nilai) AS nilai
    FROM (
      SELECT
        bk.id AS id_faktur,
        ds.id AS id_sales,
        s.nama_sales,
        o.nama_outlet,
        YEAR(bk.tglfaktur) AS tahun,
        MONTH(bk.tglfaktur) AS bulan,
        {nilai_case} AS nilai
      FROM barang_keluar bk
      JOIN detail_sales ds ON ds.id = bk.id_sales
      JOIN sales s ON s.id = ds.id_sales
      JOIN outlet o ON o.id = ds.id_outlet
      LEFT JOIN detail_barang_keluar dbk ON dbk.id_barang_keluar = bk.id
          
      {"WHERE " + " AND ".join(where) if where else ""}
      GROUP BY
        bk.id, ds.id, s.nama_sales, o.nama_outlet,
        YEAR(bk.tglfaktur), MONTH(bk.tglfaktur),
        bk.lunas_tidak
    ) x

    GROUP BY x.id_sales, x.nama_sales, x.nama_outlet, x.tahun, x.bulan
    ORDER BY x.tahun DESC, x.bulan ASC, x.nama_outlet, x.nama_sales;
    """
    return fetch(sql, params)  # hasil: list of dict {id_sales, nama_sales, nama_outlet, tahun, bulan, nilai}
@app.route('/admin/performa_sales')
def adminperformasales():
    tahun = request.args.get('tahun', type=int)
    bulan = request.args.get('bulan', type=int)
    lunas_tidak = request.args.get('lunas_tidak', type=str)
    print(lunas_tidak)
    selected_sales = request.args.get('sales')

    # Ambil data agregat yang sama dengan export
    rows = query_performa(tahun, bulan, selected_sales, lunas_tidak)

    # Daftar tahun untuk filter UI
    list_tahun = fetch("""
        SELECT DISTINCT YEAR(tglfaktur) AS tahun
        FROM barang_keluar
        ORDER BY tahun DESC
    """)
    thn = fetch("""
        SELECT DISTINCT YEAR(tglfaktur) AS tahun
        FROM barang_keluar
        {where}
        ORDER BY tahun DESC
    """.format(where="WHERE YEAR(tglfaktur)=%s" if tahun else ""), ([tahun] if tahun else []))

    # Daftar nama sales untuk dropdown
    nama_sales = fetch("""
        SELECT DISTINCT nama_sales FROM sales
    """)

    # Bangun pivot per (tahun, id_sales, outlet, sales)
    months = [bulan] if bulan else list(range(1, 13))
    idx = {}  # key: (tahun, id_sales, outlet, sales)
    for r in rows:
        key = (r['tahun'], r['id_sales'], r['nama_outlet'], r['nama_sales'])
        if key not in idx:
            idx[key] = {
                "tahun": r['tahun'],
                "id_sales": r['id_sales'],
                "nama_outlet": r['nama_outlet'],
                "nama_sales": r['nama_sales'],
                **{f"M{i}": 0 for i in months},
                "total_sales": 0
            }
        val = int(r['nilai'] or 0)
        if r['bulan'] in months:
            idx[key][f"M{r['bulan']}"] += val
            idx[key]["total_sales"] += val

    data_fix = list(idx.values())

    # Footer per tahun hanya dari data yang tampil
    footer_totals = {}
    for t in sorted({d["tahun"] for d in data_fix}, reverse=True):
        totals = {f"M{i}": 0 for i in months}
        totals["total_sales"] = 0
        for d in data_fix:
            if d["tahun"] != t:
                continue
            for i in months:
                totals[f"M{i}"] += d[f"M{i}"]
                totals["total_sales"] += d[f"M{i}"]
        footer_totals[t] = totals
    
    return render_pjax(
        "admin/performa_sales.html",
        
        data_fix=data_fix,
        list_tahun=list_tahun,
        tahun=thn,
        nama_sales=nama_sales,
        footer_totals=footer_totals,
        filter_bulan=bulan,
        months=months
    )
from flask import request, send_file
import io
from datetime import date
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

def parse_amount(x):
    s = str(x or "").replace("Rp", "").replace("rp", "").replace(" ", "").replace(".", "").replace(",", "")
    return int(s) if s.isdigit() else None

from typing import Tuple, List, Optional
from datetime import date, timedelta
from typing import Optional, Tuple, List

def build_date_range(
    year: Optional[int],
    month: Optional[int],
    day: Optional[int] = None,
    alias: str = "bm",        # alias tabel di query
    col: str = "tglfaktur"    # nama kolom tanggal
) -> Tuple[Optional[str], List[date]]:
    """
    Balikkan (clause_sql, params) untuk digunakan di WHERE query:
    "{alias}.{col} >= %s AND {alias}.{col} < %s"

    Contoh:
    - build_date_range(2025, None)           -> tahun penuh 2025
    - build_date_range(None, 5)              -> bulan Mei (semua tahun)
    - build_date_range(None, None, 1)        -> tanggal 1 (semua bulan/tahun)
    - build_date_range(2025, 5)              -> Mei 2025
    - build_date_range(2025, 5, 10)          -> 10 Mei 2025
    """

    clause, params = None, []

    # Tahun + Bulan + Hari
    if year and month and day:
        start = date(year, month, day)
        end = start + timedelta(days=1)
        clause = f"{alias}.{col} >= %s AND {alias}.{col} < %s"
        params = [start, end]

    # Tahun + Bulan
    elif year and month:
        from calendar import monthrange
        last = monthrange(year, month)[1]
        start = date(year, month, 1)
        end = date(year, month, last) + timedelta(days=1)
        clause = f"{alias}.{col} >= %s AND {alias}.{col} < %s"
        params = [start, end]

    # Tahun saja
    elif year:
        start = date(year, 1, 1)
        end = date(year + 1, 1, 1)
        clause = f"{alias}.{col} >= %s AND {alias}.{col} < %s"
        params = [start, end]

    # Bulan saja (semua tahun)
    elif month:
        clause = f"MONTH({alias}.{col}) = %s"
        params = [month]

    # Hari saja (semua bulan/tahun)
    elif day:
        clause = f"DAY({alias}.{col}) = %s"
        params = [day]

    return clause, params
from flask import request, send_file, current_app
import io, time


@app.route('/admin/performa_sales/export_excell')
def adminperformasalesprint():
    import time
    t0 = time.perf_counter()
    log = current_app.logger

    th = request.args.get('tahun', type=int)
    bln = request.args.get('bulan', type=int)
    lunas_tidak = request.args.get('lunas_tidak', type=str)
    print(lunas_tidak)
    nama_sales = request.args.get('sales')
    raw_angka = request.args.get('angka')
    min_amount = parse_amount(raw_angka)

    months = [bln] if bln else list(range(1, 13))

    # 1) Ambil data agregat (kolom: id_sales, nama_sales, nama_outlet, tahun, bulan, nilai)
    rows = query_performa(th, bln, nama_sales, lunas_tidak)
    t1 = time.perf_counter()
    log.debug("[EXPORT] fetched=%d rows, keys=%s", len(rows), list(rows[0].keys()) if rows else None)

    # 2) Pivot ke baris Excel (punya M{bulan} & total_sales)
    pivot = {}
    for r in rows:
        key = (r["tahun"], r["id_sales"], r["nama_outlet"], r["nama_sales"])
        if key not in pivot:
            pivot[key] = {
                "tahun": r["tahun"],
                "id_sales": r["id_sales"],
                "nama_outlet": r["nama_outlet"],
                "nama_sales": r["nama_sales"],
                **{f"M{i}": 0 for i in months},
                "total_sales": 0
            }
        v = int(r["nilai"] or 0)
        if r["bulan"] in months:
            pivot[key][f"M{r['bulan']}"] += v
            pivot[key]["total_sales"] += v

    data = list(pivot.values())

    # 3) Filter nominal (opsional)
    if min_amount is not None:
        data = [d for d in data if (d[f"M{bln}"] if bln else d["total_sales"]) >= min_amount]

    # 4) Workbook + header
    wb = Workbook()
    ws = wb.active
    ws.title = "Laporan Performa Sales"

    headers = ["Tahun", "Nama Outlet", "Nama Sales"] + [f"Bulan {m}" for m in months] + ["Jumlah Total"]
    ws["A1"] = "LAPORAN PERFORMA SALES"
    ws["A1"].font = Font(size=14, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.merge_cells(f"A1:{get_column_letter(len(headers))}1")

    ws.append(headers)
    for c in range(1, len(headers) + 1):
        ws.cell(row=2, column=c).font = Font(bold=True)

    # 5) Tulis data dari 'data' (BUKAN 'rows')
    totals = defaultdict(lambda: {"total": 0, **{f"M{i}": 0 for i in months}})
    for item in data:
        vals = [item["tahun"], item["nama_outlet"], item["nama_sales"]]
        month_vals = [int(item.get(f"M{i}", 0) or 0) for i in months]
        total_val = int(item.get("total_sales", 0) or 0)
        ws.append(vals + month_vals + [total_val])

        # akumulasi footer per tahun
        for i, v in zip(months, month_vals):
            totals[item["tahun"]][f"M{i}"] += v
            totals[item["tahun"]]["total"] += v

    # 6) Footer JUMLAH per tahun
    for thn in sorted(totals.keys(), reverse=True):
        row = [thn, "", "JUMLAH:"] + [totals[thn][f"M{i}"] for i in months] + [totals[thn]["total"]]
        ws.append(row)
        for c in range(1, len(headers) + 1):
            ws.cell(row=ws.max_row, column=c).font = Font(bold=True)

    # 7) Format angka Rp
    num_cols = range(4, 4 + len(months) + 1)
    for col in num_cols:
        for rr in range(3, ws.max_row + 1):
            cell = ws.cell(row=rr, column=col)
            cell.number_format = '"Rp" #,##0'
            cell.alignment = Alignment(horizontal="right")

    # 8) Freeze, filter, width
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}{ws.max_row}"
    for i in range(1, len(headers) + 1):
        col = get_column_letter(i)
        maxlen = max(len(str(ws.cell(row=r, column=i).value or "")) for r in range(1, ws.max_row + 1))
        ws.column_dimensions[col].width = min(50, max(12, maxlen + 2))

    # 9) Kirim file
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    t2 = time.perf_counter()
    log.info("[EXPORT] wrote excel rows=%d cols=%d fetch=%.3fs build=%.3fs total=%.3fs",
             ws.max_row - 2, len(headers), t1 - t0, t2 - t1, t2 - t0)

    return send_file(bio, as_attachment=True,
                     download_name="laporan_performa_sales.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
@app.route('/admin/annual_report', methods=['GET'])
def annual_report():
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    # Header dan logo
    logo_path = "static/image/Logo.png"
    try:
        c.drawImage(logo_path, 1.5 * cm, height - 4 * cm, width=4 * cm, preserveAspectRatio=True)
    except:
        pass  # Jika logo tidak ditemukan, lanjutkan saja

    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(width / 2, height - 2 * cm, "LAPORAN TAHUNAN PERUSAHAAN ")

    c.setFont("Helvetica", 12)
    c.drawCentredString(width / 2, height - 3 * cm, f"Tahun: {time_zone_wib().year}")

    y = height - 5.5 * cm
    c.setFont("Helvetica-Bold", 14)
    c.drawString(2 * cm, y, "1. RINGKASAN KEUANGAN")
    y -= 1 * cm

    # ringkasan
    tahun_ini = str(time_zone_wib().year)

    g.con.execute("""
    SELECT
        b.nama_barang,
        COALESCE(SUM(bm.jml_menerima), 0) AS total_pembelian,
        COALESCE(SUM(bk.jmlpermintaan), 0) AS total_penjualan,
        (COALESCE(SUM(bm.jml_menerima), 0) - COALESCE(SUM(bk.jmlpermintaan), 0)) AS stok_akhir,
        COALESCE(SUM(bk.profit), 0) AS total_laba
    FROM
        barang b
    LEFT JOIN detail_barang_masuk bm ON b.id = bm.id_barang
    LEFT JOIN detail_barang_keluar bk ON b.id = bk.id_barang
    LEFT JOIN barang_keluar dbk ON bk.id_barang_keluar = dbk.id
    WHERE 
        YEAR(dbk.tglfaktur) = %s
    GROUP BY
        b.id, b.nama_barang;""",(tahun_ini,))
    ringkasan = g.con.fetchall()
    c.setFont("Helvetica-Bold", 12)
    c.drawString(2 * cm, y, "Detail Barang:")
    y -= 0.7 * cm
    c.setFont("Helvetica", 9)

    for row in ringkasan:
        nama_barang = row[0]  # b.nama_barang
        total_pembelian = row[1]
        total_penjualan = row[2]
        stok_akhir = row[3]
        total_laba = row[4]

        if y < 3 * cm:
            c.showPage()
            y = height - 3 * cm

        c.drawString(
            2 * cm, y,
            f"""{nama_barang[:30]:30} | Beli: {total_pembelian} 
            | Jual: {total_penjualan} | Stok: {stok_akhir} | Laba: Rp {total_laba:,.2f}"""
        )
        y -= 0.5 * cm


    c.setFont("Helvetica", 10)
    # Angka langsung dari query
    total_pendapatan = sum([row[2] for row in ringkasan])  # total_penjualan
    total_pengeluaran = sum([row[1] for row in ringkasan]) # total_pembelian
    laba_bersih = sum([row[4] for row in ringkasan])       # total_laba


    c.drawString(2 * cm, y, f"Total Pendapatan: Rp {total_pendapatan:,.2f}")
    y -= 0.5 * cm
    c.drawString(2 * cm, y, f"Total Pengeluaran: Rp {total_pengeluaran:,.2f}")
    y -= 0.5 * cm
    c.drawString(2 * cm, y, f"Laba Bersih: Rp {laba_bersih:,.2f}")
    y -= 1 * cm


    c.setFont("Helvetica-Bold", 14)
    c.drawString(2 * cm, y, "2. CATATAN KEUANGAN")
    y -= 1 * cm
    c.setFont("Helvetica", 10)
    notes = [
        "• Semua data keuangan berdasarkan laporan sistem selama 12 bulan terakhir.",
        "• Pembayaran yang belum lunas dicatat dalam laporan piutang.",
        "• Seluruh transaksi telah diaudit internal."
    ]
    for note in notes:
        c.drawString(2 * cm, y, note)
        y -= 0.5 * cm

    # Footer
    c.setFont("Helvetica", 8)
    c.drawString(2 * cm, 2 * cm, f"Dibuat pada {time_zone_wib().strftime('%d-%m-%Y')}")

    c.showPage()
    c.save()
    buffer.seek(0)

    return send_file(buffer, as_attachment=True, download_name="laporan_tahunan.pdf", mimetype='application/pdf')
from flask import send_file
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from io import BytesIO

@app.route('/admin/latest_penerimaan', methods=['GET'])
def latest_penerimaan_excell():
    wb = Workbook()
    ws = wb.active
    ws.title = "Laporan Bulanan"

    bulan_sekarang = time_zone_wib().month
    tahun_ini = time_zone_wib().year
    current_row = 1

    # Judul umum
    ws.merge_cells(f'A{current_row}:K{current_row}')
    ws[f'A{current_row}'] = "LAPORAN BULANAN PERUSAHAAN"
    ws[f'A{current_row}'].font = Font(size=14, bold=True)
    ws[f'A{current_row}'].alignment = Alignment(horizontal="center")
    current_row += 2

    for bulan in range(1, bulan_sekarang + 1):
        nama_bulan = next((b['nama_bulan'] for b in list_bulan if b['value'] == str(bulan)), "Bulan Tidak Diketahui")
        print(nama_bulan)
        # Ambil data dari database
        clause, prms = build_date_range(tahun_ini, bulan, alias='bm', col='tglfaktur')
        sql = f"""
        SELECT
          bm.tglfaktur, bm.nofaktur, s.nama_supplier AS supplier,
          b.nama_barang, b.kode_barang, b.qty,
          dbm.jml_menerima AS jumlah_penerimaan,
          dbm.harga_satuan, dbm.harga_total,
          dbm.lunastidak, bm.lunas_tidak
        FROM detail_barang_masuk dbm
        JOIN barang_masuk bm ON bm.id = dbm.id_barang_masuk
        JOIN pabrik s ON s.id = bm.id_supplier
        JOIN barang b ON b.id = dbm.id_barang
        WHERE {clause}
        """
        hasil = fetch(sql, prms)
        
        if not hasil:
            continue  # Lewati jika tidak ada data bulan ini

        # Judul bulan
        ws.merge_cells(f'A{current_row}:K{current_row}')
        ws[f'A{current_row}'] = f"LAPORAN PEMBELIAN - {nama_bulan} {tahun_ini}"
        ws[f'A{current_row}'].alignment = Alignment(horizontal="center")
        ws[f'A{current_row}'].font = Font(bold=True)
        current_row += 1

        # Header
        headers = [
            "TANGGAL FAKTUR",
            "NO FAKTUR",
            "SUPPLIER",
            "NAMA BARANG",
            "KODE BARANG",
            "QUANTITY",
            "JUMLAH PENERIMAAN",
            "HARGA SATUAN",
            "HARGA TOTAL",
            "LUNAS / TIDAK",
            "PERFORMA BELANJA"
        ]
        ws.append(headers)
        for cell in ws[current_row]:
            cell.font = Font(bold=True)
        current_row += 1

        performa_belanja = 0
        # Tulis data
        baris_awal = current_row
        for row in hasil:
            def norm(v): return v.strip().title() if isinstance(v, str) else None
            valid = {'Lunas', 'Tidak Lunas'}
            v1, v2 = norm(row.get('lunas_tidak')), norm(row.get('lunastidak'))
            lunas_tidak = next((v for v in (v1, v2) if v in valid), 'Tidak Lunas')
            print("yang dipilih:" + lunas_tidak)
            if lunas_tidak == "Lunas":
                # Hitung performa
                performa_belanja += row['harga_total']
            ws.append([
                row['tglfaktur'].strftime("%Y-%m-%d") if row['tglfaktur'] else '',
                row['nofaktur'],
                row['supplier'],
                row['nama_barang'],
                row['kode_barang'],
                row['qty'],
                row['jumlah_penerimaan'],
                row['harga_satuan'],
                row['harga_total'],
                lunas_tidak,
                ''  # Placeholder kolom performa
            ])
            current_row += 1

        # Isi dan merge performa
        baris_akhir = current_row - 1
        ws.merge_cells(f'K{baris_awal}:K{baris_akhir}')
        ws[f'K{baris_awal}'] = performa_belanja
        ws[f'K{baris_awal}'].alignment = Alignment(horizontal='center', vertical='center')

        current_row += 2  # Spasi antar bulan

    # Otomatis lebar kolom
    for i, col in enumerate(ws.columns, 1):
        max_length = 0
        col_letter = get_column_letter(i)
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2

    # Output ke Excel
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="REPORT_PENERIMAAN.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/admin/latest_pengeluaran', methods=['GET'])
def latest_pengeluaran_excell():
    wb = Workbook()
    ws = wb.active
    ws.title = "Laporan Bulanan"

    bulan_sekarang = time_zone_wib().month
    tahun_ini = time_zone_wib().year
    current_row = 1

    # ===== helper untuk hitung lebar kolom saat append =====
    # total kolom = 12 (lihat headers)
    COLS = 12
    col_max = [0] * COLS
    def upd_width(vals):
        for i, v in enumerate(vals[:COLS]):
            if v is None:
                continue
            l = len(str(v))
            if l > col_max[i]:
                col_max[i] = l

    # ===== judul umum =====
    ws.merge_cells(f'A{current_row}:L{current_row}')
    ws[f'A{current_row}'] = "LAPORAN BULANAN PERUSAHAAN"
    ws[f'A{current_row}'].font = Font(size=14, bold=True)
    ws[f'A{current_row}'].alignment = Alignment(horizontal="center")
    upd_width(["LAPORAN BULANAN PERUSAHAAN"])  # opsional, supaya kolom A tidak terlalu sempit
    current_row += 2

    # helper normalisasi status lunas
    valid_status = {'Lunas', 'Tidak Lunas'}
    def norm(v): 
        return v.strip().title() if isinstance(v, str) else None

    for bulan in range(1, bulan_sekarang + 1):
        # nama bulan
        bulan_str = str(bulan)
        nama_bulan = next((b['nama_bulan'] for b in list_bulan if b['value'] == bulan_str), "Bulan Tidak Diketahui")

        # rentang tanggal [start, end)
        last = calendar.monthrange(tahun_ini, bulan)[1]
        start = date(tahun_ini, bulan, 1)
        end   = date(tahun_ini, bulan, last) + timedelta(days=1)

        # ambil data (pakai rentang tanggal biar index kepakai)
        hasil = fetch("""
            SELECT
                bk.tglfaktur,
                bk.jatuhtempo,
                bk.nomerfaktur,
                s.nama_sales,
                o.nama_outlet,
                b.nama_barang,
                b.qty,
                dbk.jmlpermintaan,
                dbk.harga_satuan,
                dbk.harga_total,
                dbk.lunas_or_no,
                bk.lunas_tidak
            FROM detail_barang_keluar dbk
            JOIN barang_keluar bk ON bk.id = dbk.id_barang_keluar
            JOIN detail_sales ds ON ds.id = bk.id_sales
            JOIN sales s ON s.id = ds.id_sales
            JOIN outlet o ON o.id = ds.id_outlet
            JOIN barang b ON b.id = dbk.id_barang
            WHERE bk.tglfaktur >= %s AND bk.tglfaktur < %s
        """, (start, end))

        if not hasil:
            continue

        # judul bulan
        ws.merge_cells(f'A{current_row}:L{current_row}')
        ws[f'A{current_row}'] = f"{nama_bulan} - {tahun_ini}"
        ws[f'A{current_row}'].alignment = Alignment(horizontal="center")
        ws[f'A{current_row}'].font = Font(bold=True)
        upd_width([f"{nama_bulan} - {tahun_ini}"])
        current_row += 1

        # header tabel
        headers = [
            "TANGGAL FAKTUR",
            "JATUH TEMPO",
            "NOMOR FAKTUR",
            "SALES",
            "NAMA OUTLET",
            "NAMA BARANG",
            "QUANTITY",
            "JUMLAH PERMINTAAN",
            "HARGA SATUAN",
            "HARGA TOTAL",
            "LUNAS / TIDAK LUNAS",
            "PERFORMA SALES"
        ]
        ws.append(headers)
        for cell in ws[current_row]:
            cell.font = Font(bold=True)
        upd_width(headers)
        current_row += 1

        performa_sales = 0
        baris_awal = current_row

        # tulis data
        for row in hasil:
            v1, v2 = norm(row.get('lunas_tidak')), norm(row.get('lunas_or_no'))
            lunas_tidak = next((v for v in (v1, v2) if v in valid_status), 'Tidak Lunas')
            if lunas_tidak == "Lunas":
                performa_sales += row['harga_total']

            row_vals = [
                row['tglfaktur'].strftime("%Y-%m-%d") if row['tglfaktur'] else '',
                row['jatuhtempo'].strftime("%Y-%m-%d") if row['jatuhtempo'] else '',
                row['nomerfaktur'],
                row['nama_sales'],
                row['nama_outlet'],
                row['nama_barang'],
                row['qty'],
                row['jmlpermintaan'],
                row['harga_satuan'],
                row['harga_total'],
                lunas_tidak,
                ''  # placeholder performa
            ]
            ws.append(row_vals)
            upd_width(row_vals)
            current_row += 1

        # merge & isi kolom performa sales
        baris_akhir = current_row - 1
        ws.merge_cells(f'L{baris_awal}:L{baris_akhir}')
        ws[f'L{baris_awal}'] = performa_sales
        ws[f'L{baris_awal}'].alignment = Alignment(horizontal='center', vertical='center')
        upd_width(['', '', '', '', '', '', '', '', '', '', '', performa_sales])

        current_row += 2  # spasi antar bulan

    # set lebar kolom sekali, tanpa loop seluruh cells
    for i, mx in enumerate(col_max, start=1):
        ws.column_dimensions[get_column_letter(i)].width = (mx or 0) + 2

    # output excel
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="REPORT_PENGELUARAN.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route("/admin/Report_Laporan")
def report_laporan():
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    # --- Header dan Logo ---
    logo_path = "static/image/Logo.png"
    try:
        c.drawImage(logo_path, 1.5 * cm, height - 4 * cm, width=4 * cm, preserveAspectRatio=True)
    except:
        pass  # Kalau logo nggak ada, skip

    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(width / 2, height - 2 * cm, "REPORT LAPORAN BY SYSTEM")

    hari_ini = time_zone_wib().strftime("%d-%m-%Y")
    c.setFont("Helvetica", 12)
    c.drawCentredString(width / 2, height - 2.8 * cm, hari_ini)

    # --- Posisi awal tabel ---
    y_pos = height - 5 * cm
    c.setFont("Helvetica", 11)

    # --- Header tabel ---
    # c.setFont("Helvetica-Bold", 11)
    # c.drawString(col1_x, y_pos, "No")
    # c.drawString(col2_x, y_pos, "Keterangan")
    # c.drawString(col3_x, y_pos, "Nilai")
    # y_pos -= 0.6 * cm
    # c.setFont("Helvetica", 11)

    # --- Query & data ---
    today = time_zone_wib()
    tanggal_db = today.strftime("%Y-%m-%d")
    bulan_ini = today.month
    tahun_ini = today.year
    print(tanggal_db)
    # untuk hari ini
    start_day = date.fromisoformat(tanggal_db)          # yyyy-mm-dd
    end_day   = start_day + timedelta(days=1)
    # Data list
    data = []
    # INKASO HARIAN
    g.con.execute("""
            SELECT COALESCE(SUM(dbk.harga_total), 0) 
            FROM detail_barang_keluar dbk 
            INNER JOIN barang_keluar bk ON bk.id = dbk.id_barang_keluar 
            WHERE bk.tanggal_pembayaran >= %s AND bk.tanggal_pembayaran < %s
            AND bk.lunas_tidak = 'Lunas'
        """, (start_day, end_day))
    inkaso = g.con.fetchone()[0]
    print(inkaso)
    data.append(("INKASO - HARIAN", format_rp(inkaso)))

    # SALES HARIAN
    g.con.execute("""
        SELECT COALESCE(SUM(harga_total),0)
        FROM detail_barang_keluar
        INNER JOIN barang_keluar bk ON bk.id = detail_barang_keluar.id_barang_keluar
        WHERE DATE(bk.tglfaktur)=%s
    """, (tanggal_db,))
    sales_harian = g.con.fetchone()[0]
    data.append(("SALES - HARIAN", format_rp(sales_harian)))

    # PERFORMA SALES TAHUNAN
    performa_sales_tahunan = 0

    # PERFORMA BELANJA TAHUNAN
    performa_belanja_tahunan = 0

    # CASHFLOW TAHUNAN
    cashflow = 0

    # TERBAYAR / LUNAS - KEUANGAN (Tahunan)
    lunas_keuangan = 0

    # TERBAYAR / LUNAS - ADMINISTRASI (Tahunan)
    lunas_admin = 0

    # HUTANG / TIDAK LUNAS - KEUANGAN (Tahunan)
    hutang_keuangan = 0

    # PIUTANG / TIDAK LUNAS - ADMINISTRASI (Tahunan)
    piutang_admin = 0
    bulan_sekarang = time_zone_wib().month
    tahun_ini = time_zone_wib().year

    for bulan in range(1, bulan_sekarang + 1):
        # Ambil data dari database
        hasil = fetch("""
            SELECT
                bk.tglfaktur,
                bk.jatuhtempo,
                bk.nomerfaktur,
                s.nama_sales,
                o.nama_outlet,
                b.nama_barang,
                b.qty,
                dbk.jmlpermintaan,
                dbk.harga_satuan,
                dbk.harga_total,
                dbk.lunas_or_no,
                bk.lunas_tidak
            FROM
                detail_barang_keluar dbk
            INNER JOIN barang_keluar bk on bk.id = dbk.id_barang_keluar
            INNER JOIN detail_sales ds on ds.id = bk.id_sales
            INNER JOIN sales s on s.id = ds.id_sales
            INNER JOIN outlet o on o.id = ds.id_outlet
            INNER JOIN barang b on b.id = dbk.id_barang     
            WHERE
                YEAR(bk.tglfaktur) = %s AND MONTH(bk.tglfaktur) = %s;
        """, (tahun_ini, bulan))
        
        performa_sales = 0
        # Tulis data per row
        for idx, row in enumerate(hasil, start=1):
            def norm(v): return v.strip().title() if isinstance(v, str) else None
            valid = {'Lunas', 'Tidak Lunas'}
            v1, v2 = norm(row.get('lunas_tidak')), norm(row.get('lunas_or_no'))
            lunas_tidak = next((v for v in (v1, v2) if v in valid), 'Tidak Lunas')
            performa_sales += row['harga_total']
            if lunas_tidak == "Lunas":
                lunas_keuangan += row['harga_total']
            else:
                hutang_keuangan += row['harga_total']

        performa_sales_tahunan += performa_sales
        # Ambil data dari database
        hasil = fetch("""
            SELECT
                bm.tglfaktur,
                bm.nofaktur,
                s.nama_supplier AS supplier,
                b.nama_barang,
                b.kode_barang,
                b.qty,
                dbm.jml_menerima AS jumlah_penerimaan,
                dbm.harga_satuan,
                dbm.harga_total,
                dbm.lunastidak,
                bm.lunas_tidak
            FROM
                detail_barang_masuk dbm 
            INNER JOIN barang_masuk bm on bm.id = dbm.id_barang_masuk 
            INNER JOIN pabrik s on s.id = bm.id_supplier
            INNER JOIN barang b on b.id = dbm.id_barang
            WHERE
                YEAR(bm.tglfaktur) = %s AND MONTH(bm.tglfaktur) = %s;
        """, (tahun_ini, bulan))
        
        performa_belanja = 0
        # Tulis data per row
        for idx, row in enumerate(hasil, start=1):
            def norm(v): return v.strip().title() if isinstance(v, str) else None
            valid = {'Lunas', 'Tidak Lunas'}
            v1, v2 = norm(row.get('lunas_tidak')), norm(row.get('lunastidak'))
            lunas_tidak = next((v for v in (v1, v2) if v in valid), 'Tidak Lunas')
            performa_belanja += row['harga_total']
            if lunas_tidak == "Lunas":
                lunas_admin += row['harga_total']
            else:
                piutang_admin += row['harga_total']

        performa_belanja_tahunan += performa_belanja
        
        if bulan == bulan_sekarang:
            # PERFORMA SALES BULANAN
            performa_sales_bulanan = performa_sales
            data.append(("PERFORMA SALES - BULANAN", format_rp(performa_sales_bulanan)))

            # PERFORMA BELANJA BULANAN
            performa_belanja_bulanan = performa_belanja
            data.append(("PERFORMA BELANJA - BULANAN", format_rp(performa_belanja_bulanan)))

            # CASHFLOW BULANAN
            cashflow = performa_sales_bulanan - performa_belanja_bulanan
            data.append(("CASHFLOW - BULANAN", format_rp(cashflow)))

    # PERFORMA SALES TAHUNAN
    data.append(("PERFORMA SALES - TAHUNAN", format_rp(performa_sales_tahunan)))

    # PERFORMA BELANJA TAHUNAN
    data.append(("PERFORMA BELANJA - TAHUNAN", format_rp(performa_belanja_tahunan)))

    # CASHFLOW TAHUNAN
    cashflow = performa_sales_tahunan - performa_belanja_tahunan
    data.append(("CASHFLOW - TAHUNAN", format_rp(cashflow) ))

    # TERBAYAR / LUNAS - KEUANGAN (Tahunan)
    data.append(("LUNAS - KEUANGAN", format_rp(lunas_keuangan)))

    # TERBAYAR / LUNAS - ADMINISTRASI (Tahunan)
    data.append(("LUNAS - ADMINISTRASI", format_rp(lunas_admin)))

    # HUTANG / TIDAK LUNAS - KEUANGAN (Tahunan)
    data.append(("TIDAK LUNAS - KEUANGAN", format_rp(hutang_keuangan)))

    # PIUTANG / TIDAK LUNAS - ADMINISTRASI (Tahunan)
    data.append(("TIDAK LUNAS - ADMINISTRASI", format_rp(piutang_admin)))

    # --- Tulis data ke tabel PDF ---
    # data yang sebelumnya dipakai di loop
    tabel_data = []  # header
    for i, (ket, nilai) in enumerate(data, start=1):
        tabel_data.append([str(i), ket, str(nilai)])

    # buat tabel
    table = Table(tabel_data, colWidths=[50, 300, 150])  # sesuaikan lebar kolom
    table.setStyle(TableStyle([
        ("TEXTCOLOR", (0,0), (-1,0), colors.black),
        ("ALIGN", (0,0), (0,-1), "CENTER"),   # kolom No rata tengah
        ("ALIGN", (1,0), (1,-1), "RIGHT"),     # kolom Keterangan rata kiri
        ("ALIGN", (2,0), (2,-1), "RIGHT"),    # kolom Nilai rata kanan
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), 10),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),  # border tabel
    ]))

    # render ke canvas
    table.wrapOn(c, width, height)
    table.drawOn(c, 2*cm, y_pos - len(tabel_data)*0.6*cm)

    c.save()

    buffer.seek(0)
    return send_file(
        buffer,
        as_attachment=True,
        download_name="REPORT_LAPORAN.pdf",
        mimetype="application/pdf"
    )
