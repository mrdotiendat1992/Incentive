from flask import Flask, render_template, request, url_for, redirect, g, flash, jsonify, send_file, session, flash, get_flashed_messages, render_template_string
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, current_user, login_required
import pyodbc
import os
from functools import wraps
import logging
from logging.handlers import RotatingFileHandler
import urllib.parse

# DANH_SACH_TO_TRUONG = {
#     "11S01":"7721",
#     "11S03":"2540",
#     "11S05":"398",
#     "11S07":"7146",
#     "11S09":"115",
#     "11S11":"2340",
#     "11S13":"3943",
#     "12S01":"233",
#     "12S03":"385",
#     "12S05":"1163",
#     "12S07":"6318",
#     "12S09":"12756",
#     "12S11":"1192",
#     "21S01":"262",
#     "21S03":"828",
#     "21S05":"4727",
#     "21S07":"152",
#     "21S09":"4531",
#     "21S11":"2565",
#     "21S13":"3494",
#     "22S01":"83",
#     "22S03":"1162",
#     "22S05":"1152",
#     "22S07":"1657",
#     "22S09":"376",
#     "22S11":"4952",
#     "22S13":"3590",
#     "23S01":"4726",
#     "23S07":"4882",
#     "23S09":"2576",
#     "25S09":"669",
#     "25S11":"4706"
# }

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
    try:
        conn = connect_db()
        cursor = execute_query(conn, f"select Line from Danh_sach_CBCNV where The_cham_cong='{masothe}' and Factory='{macongty}'")
        result = cursor.fetchone()
        close_db(conn)
        return result[0]
    except:
        return ""
    
def get_all_styles(ngay, chuyen):
    try:
        if ngay:
            conn = connect_db()
            cursor = execute_query(conn, f"SELECT Distinct STYLE FROM [INCENTIVE].[dbo].[SL_CA_NHAN] WHERE NGAY='{ngay}' AND CHUYEN='{chuyen}'")
            result = cursor.fetchall()
            close_db(conn)
            return [style[0] for style in result]
        else:
            return []
    except:
        return []

def lay_danhsach_congnhan_trongchuyen(chuyen):
    conn = connect_db()
    query = f"SELECT * FROM [INCENTIVE].[dbo].[DS_CN_MAY_THEO_HC_CHUYEN] WHERE Chuyen='{chuyen}'"
    cursor = execute_query(conn, query) 
    result = cursor.fetchall()
    close_db(conn)
    return list(result)
    
def lay_danhsach_chuyen_hotro(chuyen):
    conn = connect_db()
    query = f"SELECT * FROM [INCENTIVE].[dbo].[DS_CHUYEN_MAY] WHERE LINE LIKE '{chuyen[0]}%' ORDER BY LINE"
    cursor = execute_query(conn, query) 
    result = cursor.fetchall()
    close_db(conn)
    return [line[0] for line in result]
  
def lay_danhsach_sanluong(ngay, chuyen, style,mst,hoten,macongdoan):
    conn = connect_db()
    query = f"SELECT * FROM [INCENTIVE].[dbo].[SL_CA_NHAN] WHERE 1=1 "
    if ngay:
        query += f"AND NGAY='{ngay}' "
    else:
        return []
    if chuyen:
        query += f"AND CHUYEN='{chuyen}' "
    if style:
        query += f"AND STYLE='{style}' "
    else:
        return []
    if mst:
        query += f"AND MST='{mst}' "
    if hoten:
        query += f"AND HO_TEN=N'{hoten}' "
    if macongdoan:
        query += f"AND MA_CONG_DOAN='{macongdoan}' "
    query += "ORDER BY CAST(MST as INT) ASC, MA_CONG_DOAN ASC"
    cursor = execute_query(conn, query) 
    result = cursor.fetchall()
    close_db(conn)
    return list(result)

def capnhat_sanluong(mst,hoten,chuyen,ngay,style,macongdoan,sanluong):
    conn = connect_db()
    query = f"INSERT INTO [INCENTIVE].[dbo].[SL_CA_NHAN] (MST,HO_TEN,CHUYEN,NGAY,STYLE,MA_CONG_DOAN,SL_CA_NHAN) VALUES('{mst}', N'{hoten}', '{chuyen}', '{ngay}', '{style}', '{macongdoan}', '{sanluong}')"
    print(query)
    execute_query(conn, query)
    try:
        conn.commit()
        close_db(conn)
        return True
    except Exception as e:
        print(e)
        return False

def xoa_sanluong(id):
    conn = connect_db()
    query = f"DELETE FROM [INCENTIVE].[dbo].[SL_CA_NHAN] WHERE ID='{id}'"
    execute_query(conn, query)
    try:
        conn.commit()
        close_db(conn)
        return True
    except Exception as e:
        print(e)
        return False

def lay_tencongdoan(thongtin):
    try:
        macongdoan = thongtin.split("_")[0]
        style = thongtin.split("_")[1]
        conn = connect_db()
        cursor = execute_query(conn, f"SELECT TEN_CONG_DOAN FROM [INCENTIVE].[dbo].[SAM_SEW] WHERE STYLE='{style}' AND MA_CONG_DOAN='{macongdoan}'")
        result = cursor.fetchone()
        close_db(conn)
        return result[0]
    except:
        return None

def them_nguoi_di_hotro(nhamay,chuyen,mst,hoten,chucdanh,chuyendihotro,ngaydieuchuyendi,giodieuchuyendi,sogiohotro):
    try:
        conn = connect_db()
        query = f"insert into [INCENTIVE].[dbo].[CN_MAY_DI_HO_TRO] values ('{nhamay}','{mst}',N'{hoten}',N'{chucdanh}','{chuyen}','{chuyendihotro}','{ngaydieuchuyendi}','{giodieuchuyendi}','{sogiohotro}')"
        print(query)
        execute_query(conn, query)
        conn.commit()
        close_db(conn)
        return True
    except:
        return False

def lay_danhsach_di_hotro(chuyen):
    try:
        conn = connect_db()
        query = f"SELECT * FROM [INCENTIVE].[dbo].[CN_MAY_DI_HO_TRO] WHERE CHUYEN='{chuyen}'"
        cursor = execute_query(conn, query)
        result = cursor.fetchall()
        close_db(conn)
        return list(result)
    except:
        return []
 
def lay_danhsach_tnc_chua_lenchuyen(chuyen):
    try:
        conn = connect_db()
        query = f"SELECT * FROM [INCENTIVE].[dbo].[CN_TNC_CHUA_NGOI_CHUYEN] WHERE CHUYEN LIKE '{chuyen[0]}TNC%' ORDER BY CAST(MST as INT) ASC"
        cursor = execute_query(conn, query)
        result = cursor.fetchall()
        close_db(conn)
        return list(result)
    except:
        return []
    
def nhan_tnc_len_chuyen(id,chuyen):
    try:
        conn = connect_db()
        query = f"update [INCENTIVE].[dbo].[CN_TNC_NGOI_CHUYEN] SET CHUYEN_NGOI_LV='{chuyen}' WHERE ID='{id}'"
        execute_query(conn, query)
        conn.commit()
        close_db(conn)
        return True
    except:
        return False
    
def capnhat_sogio_hotro(id,sogio):
    try:
        conn = connect_db()
        query = f"update [INCENTIVE].[dbo].[CN_MAY_DI_HO_TRO] SET SO_GIO_HO_TRO='{sogio}' WHERE ID='{id}'"
        print(query)
        execute_query(conn, query)
        conn.commit()
        close_db(conn)
        return True
    except:
        return False
    
def laytongsanluongtheocongdoan(ngay,chuyen,style):
    try:
        conn = connect_db()
        query = f"select MA_CONG_DOAN,QTY from [INCENTIVE].[dbo].[TONG_SL_CONG_DOAN] where NGAY='{ngay}' and CHUYEN='{chuyen}' and STYLE='{style}' group by MA_CONG_DOAN,QTY"
        rows = execute_query(conn, query).fetchall()
        close_db(conn)
        result=[]
        for row in rows:
            if row[0]:
                result.append([row[0],row[1]])
        return result
    except:
        return []
    
def laydachsachtotruong():
    try:
        conn = connect_db()
        query = f"SELECT * FROM [INCENTIVE].[dbo].[DS_TO_TRUONG]"
        cursor = execute_query(conn, query)
        result = {"NT1":{
                "Tổ trưởng": {},
                "IE" : []
            },"NT2":{
                "Tổ trưởng": {},
                "IE" : []
            }}
        rows = cursor.fetchall()
        for row in rows:
            # print(row)
            if (row[2][2]=="S" and row[2][1].isdigit()):
                if (row[2] in result[row[0]]["Tổ trưởng"]):
                    if row[1] not in result[row[0]]["Tổ trưởng"][row[2]]:
                        result[row[0]]["Tổ trưởng"][row[2]].append(int(row[1]))
                else:
                    result[row[0]]["Tổ trưởng"][row[2]] = [int(row[1])]
            else:
                result[row[0]]["IE"].append(int(row[1]))
        # print(result)
        close_db(conn)
        return result
    except Exception as e:
        print(e)
        return []

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if current_user.is_anonymous:
            return redirect(url_for('login', next=request.url))
        return f(*args, **kwargs)
    return decorated_function

@app.before_request
def before_request():
    try:
        if current_user.is_authenticated:
            g.notice = {"line":get_line(current_user.masothe, current_user.macongty)}
        else:
            g.notice = {"line":None}
    except:
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
                    print(f"Nguoi dung {current_user.masothe} o {current_user.macongty} vua  dang nhap !!!")
                    return redirect(url_for('home'))
            return redirect(url_for("login"))
        except Exception as e:
            app.logger.error(f'Nguoi dung {masothe} o {macongty} dang nhap that bai: {e} !!!')
            return redirect(url_for("login"))
    else:
        danhsachcongty = ["NT1","NT2"]
        return render_template("login.html", danhsachcongty=danhsachcongty)

@app.route("/logout", methods=["GET","POST"])
@login_required
def logout():
    try:
        print(f"Nguoi dung {current_user.masothe} o {current_user.macongty} vua  dang xuat !!!")
        logout_user()
    except Exception as e:
        app.logger.error(f'Không thế đăng xuất {e} !!!')
    return redirect("/")

@app.route("/", methods=['GET','POST'])
@login_required
def home():
    if request.method == "GET":
        ngay = request.args.get("ngay")   
        chuyen = g.notice['line']
        style = request.args.get("style")
        mst = request.args.get("mst")
        hoten = request.args.get("hoten")
        macongdoan = request.args.get("search_macongdoan")
        styles = get_all_styles(ngay, chuyen)
        danhsach_congnhan_hotro = lay_danhsach_congnhan_trongchuyen(chuyen)
        danhsach_chuyen = lay_danhsach_chuyen_hotro(chuyen)
        danhsach_sanluong = lay_danhsach_sanluong(ngay, chuyen, style,mst,hoten,macongdoan)
        danhsach_tnc = lay_danhsach_tnc_chua_lenchuyen(chuyen)
        danhsach_di_hotro = lay_danhsach_di_hotro(chuyen)
        return render_template("home.html",styles=styles,danhsach_sanluong=danhsach_sanluong,
                               danhsach_congnhan_hotro=danhsach_congnhan_hotro,
                               danhsach_chuyen=danhsach_chuyen,danhsach_tnc=danhsach_tnc,
                               danhsach_di_hotro=danhsach_di_hotro)
    
@app.route("/nhapsanluongcanhan", methods=["POST"])
@login_required
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
    
@app.route("/nhapsanluong_congdoanmoi", methods=["POST"])
@login_required
def nhapsanluong_congdoanmoi():
    if request.method == "POST":
        ngay = request.form.get("ngay")   
        chuyen = request.form.get("chuyen")
        style = request.form.get("style")
        mst = request.form.get("mst")
        hoten = request.form.get("hoten")
        macongdoan = request.form.get("macongdoan")
        sanluong = request.form.get("sanluong")
        capnhat_sanluong(mst,hoten,chuyen,ngay,style,macongdoan,sanluong)   
        return redirect(f"/?chuyen={chuyen}&ngay={ngay}&style={style}&mst={mst}")

@app.route('/xemtencongdoan', methods=["GET","POST"])
@login_required
def xemtencongdoan():
    thongtin = request.args.get("thongtin")
    tencongdoan = lay_tencongdoan(thongtin)
    if tencongdoan:
        return jsonify(tencongdoan)
    else:
        return jsonify("Không tìm thấy")
    
@app.route("/themnguoidihotro", methods=["POST"])
@login_required
def themnguoidihotro():
    if request.method == "POST":
        chuyen = request.form.get("line_dieuchuyendi")
        mst = request.form.get("nguoiduocdieuchuyen").split("_")[0]
        hoten = request.form.get("nguoiduocdieuchuyen").split("_")[1]
        chucdanh = 'Công nhân may công nghiệp'
        chuyendihotro = request.form.get("chuyenhotro")
        ngaydieuchuyendi = request.form.get("ngaydieuchuyendi")
        giodieuchuyendi = request.form.get("giodieuchuyendi")
        sogiohotro = request.form.get("sogiohotro")
        nhamay = 'NT1' if chuyen[0]=="1" else 'NT2'
        them_nguoi_di_hotro(nhamay,chuyen,mst,hoten,chucdanh,chuyendihotro,ngaydieuchuyendi,giodieuchuyendi,sogiohotro)
        ngay = request.form.get("ngay")   
        chuyen = g.notice['line']
        style = request.form.get("style")
        return redirect(f"/?chuyen={chuyen}&ngay={ngay}&style={style}")

@app.route("/nhantnclenchuyen", methods=["POST"])
@login_required
def nhantnclenchuyen():
    if request.method == "POST":
        id = request.form.get("id_tnc")
        chuyen = request.form.get("chuyen_nhan_tnc")
        nhan_tnc_len_chuyen(id,chuyen)
        ngay = request.form.get("ngay")   
        chuyen = g.notice['line']
        style = request.form.get("style")
        return redirect(f"/?chuyen={chuyen}&ngay={ngay}&style={style}")

@app.route("/laytongsanluongtheocongdoan", methods=["POST"])
def laytongsanluong():
    if request.method == "POST":
        ngay = request.args.get("ngay")
        chuyen = request.args.get("chuyen")
        style = request.args.get("style")
        data = laytongsanluongtheocongdoan(ngay,chuyen,style)
        return jsonify(data)
    
@app.route("/capnhatsogiohotro", methods=["POST"])
def capnhatsogiohotro():
    if request.method == "POST":
        id = request.form.get("id_hotro")
        sogio = request.form.get("sogio")
        capnhat_sogio_hotro(id,sogio)
        ngay = request.form.get("ngay")   
        chuyen = g.notice['line']
        style = request.form.get("style")
        return redirect(f"/?chuyen={chuyen}&ngay={ngay}&style={style}")
    
@app.route("/danhsach_totruong", methods=["POST"])
def danhsach_totruong():
    if request.method == "POST":
        danhsach = laydachsachtotruong()
        return jsonify(danhsach)

@app.route("/xoasanluongcanhan", methods=["POST"])
def xoasanluongcanhan():
    if request.method == "POST":
        id = request.form.get("id_xoasanluong")
        print(id)
        ngay = request.form.get("ngay")   
        chuyen = g.notice['line']
        style = request.form.get("style")
        mst = request.form.get("mst")
        xoa_sanluong(id)
        return redirect(f"/?chuyen={chuyen}&ngay={ngay}&style={style}&mst={mst}")
    
if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=80)