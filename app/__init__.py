from flask import Flask,jsonify,request,render_template,send_from_directory,abort
from flask_sqlalchemy import SQLAlchemy
from flask_mysqldb import MySQL
from flask_jwt_extended import JWTManager
from flask_security import UserMixin, RoleMixin
from flask_bcrypt import Bcrypt
from flask_pjax import PJAX
from datetime import timedelta
import os, uuid

app = Flask(__name__)
pjax = PJAX(app)
app.config['MYSQL_HOST'] = 'localhost'
app.config['MYSQL_USER'] = 'root'
app.config['MYSQL_PASSWORD'] = ''
app.config['MYSQL_DB'] = 'gudang_new'
project_directory = os.path.abspath(os.path.dirname(__file__))
upload_folder = os.path.join(project_directory, 'static', 'image')
upload_nota = os.path.join(project_directory, 'static', 'nota')
app.config['UPLOAD_FOLDER'] = upload_folder 
app.config['UPLOAD_NOTA'] = upload_nota
app.config['SQLALCHEMY_DATABASE_URI'] = 'mysql://root@localhost/gudang_new'
app.config['SECRET_KEY'] = 'bukan rahasia'
app.config['SECURITY_PASSWORD_HASH'] = 'bcrypt'
app.config['SECURITY_PASSWORD_SALT'] = b'asahdjhwquoyo192382qo'
# Nonaktifkan rute login bawaan
app.config['SECURITY_LOGIN_URL'] = None
app.config['SECURITY_LOGOUT_URL'] = '/logout'  
app.config['TEMPLATES_AUTO_RELOAD'] = True
ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

db = SQLAlchemy(app)
bcrypt = Bcrypt(app)

class User(db.Model, UserMixin):
    id = db.Column(db.Integer(), primary_key=True)
    username = db.Column(db.String(255), unique=True)
    password_bcrypt = db.Column(db.String(255))
    password = db.Column(db.String(255))
    active = db.Column(db.Boolean())
    

jwt_manager = JWTManager(app)
mysql = MySQL()
mysql.init_app(app)

# allow CORS biar api yang dibuat bisa dipake website lain
from flask_cors import CORS
CORS(app)
# Import rute dari modul-modul Anda

@app.route('/sitemap.xml')
def sitemap():
    # Logika untuk menghasilkan sitemap.xml
    # Misalnya, jika sitemap.xml adalah file statis, Anda bisa mengembalikan file secara langsung
    return send_from_directory(app.static_folder, 'sitemap.xml')

@app.route('/robots.txt')
def robots():
    # Logika untuk menghasilkan robots.txt
    return """
    User-agent: *
    Disallow: /private/
    Disallow: /cgi-bin/
    Disallow: /images/
    Disallow: /pages/thankyou.html
    """
@app.route('/.well-known/appspecific/com.chrome.devtools.json')
def devtools_json():
    return jsonify({
        "workspace": {
            "uuid": uuid.uuid4(),  # UUID v4 statis
            "root": project_directory  # statis project root folder
        }
    })
# Fungsi untuk menangani kesalahan 404
@app.errorhandler(404)
def page_not_found(error):
    # Cek apakah klien meminta JSON
    if request.accept_mimetypes.accept_json and not request.accept_mimetypes.accept_html:
        # Jika klien meminta JSON, kirim respons dalam format JSON
        response = jsonify({'error': 'Not found'})
        response.status_code = 404
        return response
    # Jika tidak, kirim respons dalam format HTML
    return render_template('404.html'), 404
@app.errorhandler(401)
def unauthorized(e):
    return redirect(url_for('login'))

# Route untuk halaman yang tidak ada
@app.route('/invalid')
def invalid():
    # Menggunakan abort untuk memicu kesalahan 404
     abort(404)
    

def render_pjax(template, pjax_block='pjax_content', **kwargs):
    print(request.headers)
    if "X-PJAX" in request.headers or request.args.get('_pjax'):
        app.update_template_context(kwargs)
        template = app.jinja_env.get_template(template)
        block = template.blocks[pjax_block]
        context = template.new_context(kwargs)
        return u"".join(block(context))
    else:
        return render_template(template, **kwargs)
# Inisialisasi Flask-Compress untuk mengompres respons
from flask_compress import Compress

Compress(app)


from .admin_master import admin_master_bp
app.register_blueprint(admin_master_bp)
from .admin_sales import admin_sales_bp
app.register_blueprint(admin_sales_bp)
from . import login, api_admin

