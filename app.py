from flask import Flask, render_template, request, url_for, redirect, g, flash, jsonify, send_file, session, flash, get_flashed_messages, render_template_string
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, current_user, login_required
import pyodbc
import os
from functools import wraps
import logging
from logging.handlers import RotatingFileHandler
import urllib.parse

DANH_SACH_TO_TRUONG = [
    {"NT1":{
        "11S01":"7721",
        "11S03":"2540",
        "11S05":"398",
        "11S07":"7146",
        "11S09":"115",
        "11S11":"2340",
        "11S13":"3943",
        "12S01":"233",
        "12S03":"385",
        "12S05":"1163",
        "12S07":"6318",
        "12S09":"12756",
        "12S11":"1192",
    },
    "NT2":{
        "21S01":"262",
        "21S03":"828",
        "21S05":"4727",
        "21S07":"152",
        "21S09":"4531",
        "21S11":"2565",
        "21S13":"3494",
        "22S01":"83",
        "22S03":"1162",
        "22S05":"1152",
        "22S07":"1657",
        "22S09":"376",
        "22S11":"4952",
        "22S13":"3590",
        "23S01":"4726",
        "23S07":"4882",
        "23S09":"2576",
        "25S09":"669",
        "25S11":"4706"
    }}
]

used_db = r"Driver={SQL Server};Server=172.16.60.100;Database=HR;UID=huynguyen;PWD=Namthuan@123;"

params = urllib.parse.quote_plus(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=172.16.60.100;"
    "DATABASE=HR;"
    "UID=huynguyen;"
    "PWD=Namthuan@123;"
)

app = Flask("incentive_system")
app.config["SQLALCHEMY_DATABASE_URI"] = f"mssql+pyodbc:///?odbc_connect={params}"
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config["SECRET_KEY"] = "incentive_system"

db = SQLAlchemy(app)

handler = RotatingFileHandler('app.log', maxBytes=10000, backupCount=1, encoding='utf-8')
handler.setLevel(logging.INFO)
formatter = logging.Formatter(
    '%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d]'
)
handler.setFormatter(formatter)
app.logger.addHandler(handler)
app.logger.setLevel(logging.INFO)

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

class Nhanvien(UserMixin, db.Model):
    __tablename__ = 'Nhanvien'
    id = db.Column(db.Integer, primary_key=True)
    macongty = db.Column(db.String(10), nullable=False)
    masothe = db.Column(db.Integer, nullable=False)
    hoten = db.Column(db.Unicode(50), nullable=False)
    phongban = db.Column(db.String(10), nullable=False)
    capbac = db.Column(db.String(10), nullable=False)
    phanquyen = db.Column(db.String(10), nullable=False)
    matkhau = db.Column(db.String(10), nullable=False)

    def __repr__(self):
        return f"<User {self.hoten}>"

def connect_db():
    conn = pyodbc.connect(r'DRIVER={SQL Server};SERVER=172.16.60.100;DATABASE=HR;UID=huynguyen;PWD=Namthuan@123')
    return conn

def close_db(conn):
    conn.close()
    
def execute_query(conn, query):
    cursor = conn.cursor()
    cursor.execute(query)
    return cursor

def get_line(masothe,macongty):
    conn = connect_db()
    cursor = execute_query(conn, f"select Line from Danh_sach_CBCNV where MST='{masothe}' and Factory='{macongty}'")
    result = cursor.fetchone()
    close_db(conn)
    return result[0]

def get_all_styles(ngay:None, chuyen:None):
    conn = connect_db() 
    if ngay:
        cursor = execute_query(conn, f"SELECT Distinct STYLE FROM [INCENTIVE].[dbo].[SL_CA_NHAN] WHERE NGAY='{ngay}' AND CHUYEN='{chuyen}'")
        result = cursor.fetchall()
        close_db(conn)
        return [style[0] for style in result]
    else:
        return []
    
def lay_danhsach_sanluong(ngay, chuyen, style):
    conn = connect_db()
    if style:
        cursor = execute_query(conn, f"SELECT * FROM [INCENTIVE].[dbo].[SL_CA_NHAN] WHERE NGAY='{ngay}' AND CHUYEN='{chuyen}' AND STYLE='{style}'") 
    else:
        return []
    result = cursor.fetchall()
    close_db(conn)
    return list(result)

def capnhat_sanluong(mst,hoten,chuyen,ngay,style,macongdoan,sanluong):
    conn = connect_db()
    execute_query(conn, f"INSERT INTO [INCENTIVE].[dbo].[SL_CA_NHAN] VALUES('{mst}', N'{hoten}', '{chuyen}', '{ngay}', '{style}', '{macongdoan}', '{sanluong}',NULL)")
    try:
        conn.commit()
        close_db(conn)
        return True
    except Exception as e:
        print(e)
        return False

def lay_tencongdoan(thongtin):
    macongdoan = thongtin.split("_")[0]
    style = thongtin.split("_")[1]
    conn = connect_db()
    cursor = execute_query(conn, f"SELECT TEN_CONG_DOAN FROM [INCENTIVE].[dbo].[SAM_SEW] WHERE STYLE='{style}' AND MA_CONG_DOAN='{macongdoan}'")
    result = cursor.fetchone()
    close_db(conn)
    return result[0]
 
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if current_user.is_anonymous:
            return redirect(url_for('login', next=request.url))
        return f(*args, **kwargs)
    return decorated_function

@app.before_request
def before_request():
    if current_user.is_authenticated:
        g.notice = {"line":get_line(current_user.masothe, current_user.macongty)}
    else:
        g.notice = {"line":None}
        
@app.context_processor
def inject_notice():
    return dict(notice=getattr(g, 'notice', {}))  
    
@login_manager.user_loader
def load_user(user_id):
    return Nhanvien.query.get(int(user_id))

@app.errorhandler(404)
def page_not_found(e):
    return render_template_string("Trang không tìm thấy, vui lòng <a href='/'>quay lại</a> trang chủ"), 404

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        try:
            macongty = request.form['macongty']
            masothe = request.form['masothe']
            matkhau = request.form['matkhau']
            user = Nhanvien.query.filter_by(masothe=masothe, macongty=macongty).first()    
            if user and user.matkhau == matkhau:
                if login_user(user):    
                    app.logger.info(f"Nguoi dung {current_user.masothe} o {current_user.macongty} vua  dang nhap !!!")
                    return redirect(url_for('home'))
            return redirect(url_for("login"))
        except Exception as e:
            app.logger.error(f'Nguoi dung {masothe} o {macongty} dang nhap that bai: {e} !!!')
            return redirect(url_for("login"))
    else:
        danhsachcongty = ["NT1","NT2"]
        return render_template("login.html", danhsachcongty=danhsachcongty)

@app.route("/logout", methods=["GET","POST"])
def logout():
    try:
        app.logger.info(f"Nguoi dung {current_user.masothe} o {current_user.macongty} vua  dang xuat !!!")
        logout_user()
    except Exception as e:
        app.logger.error(f'Không thế đăng xuất {e} !!!')
    return redirect("/")

@app.route("/", methods=['GET','POST'])
@login_required
def home():
    if request.method == "GET":
        ngay = request.args.get("ngay")   
        chuyen = request.args.get("chuyen")
        style = request.args.get("style")
        styles = get_all_styles(ngay, chuyen)
        if not style and styles:
            style = styles[0]
        danhsach = lay_danhsach_sanluong(ngay, chuyen, style)
        return render_template("home.html",styles=styles,danhsach=danhsach)
    
@app.route("/nhapsanluongcanhan", methods=["POST"])
def nhapsanluongcanhan():
    if request.method == "POST":
        ngay = request.form.get("ngay")   
        chuyen = request.form.get("chuyen")
        style = request.form.get("style")
        mst = request.form.get("mst")
        hoten = request.form.get("hoten")
        macongdoan = request.form.get("macongdoan")
        sanluong = request.form.get("sanluong")
        capnhat_sanluong(mst,hoten,chuyen,ngay,style,macongdoan,sanluong)   
        return redirect(f"/?chuyen={chuyen}&ngay={ngay}&style={style}")

@app.route('/xemtencongdoan', methods=["GET","POST"])
def xemtencongdoan():
    thongtin = request.args.get("thongtin")
    return jsonify(lay_tencongdoan(thongtin))
    
if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=80)