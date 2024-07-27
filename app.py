from flask import Flask, render_template, request, url_for, redirect, g, flash, jsonify, send_file, session, flash, get_flashed_messages, render_template_string,make_response
from flask_sqlalchemy import SQLAlchemy
from flask_paginate import Pagination, get_page_parameter
from flask_login import LoginManager, UserMixin, login_user, logout_user, current_user, login_required
import pyodbc
import datetime
from functools import wraps
import logging
from logging.handlers import RotatingFileHandler
import urllib.parse
from pandas import DataFrame,read_excel,ExcelWriter
from openpyxl import load_workbook
import os
import time
from io import BytesIO

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

def chuyen_so_thanh_sotien(so):
    return "{:,.0f}".format(so)

def chinh_do_rong_cot(file_excel):
    try:
        # Mở tệp Excel để chỉnh độ rộng cột
        wb = load_workbook(file_excel)
        ws = wb.active

        # Chỉnh độ rộng cột theo độ rộng dữ liệu
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter  # Lấy tên cột (ví dụ: 'A', 'B')
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = max_length + 2
            ws.column_dimensions[column_letter].width = adjusted_width

        # Lưu lại tệp Excel đã chỉnh sửa
        wb.save(file_excel)
        wb.close()
        time.sleep(1)
        return True
    except Exception as e:
        app.logger.error(f"Loi khi dieu chinh do rong cot file excel: {e}")
        return False
    
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
        query = f"select CHUYEN from [INCENTIVE].[dbo].[DS_TO_TRUONG] where MST='{masothe}' and NHA_MAY='{macongty}'"
        # app.logger.info(query)
        cursor = execute_query(conn, query)
        rows = cursor.fetchall()
        result = [row[0] for row in rows]
        # app.logger.info(result)
        close_db(conn)
        return result
    except:
        return []
    
def get_all_styles(ngay, chuyen):
    try:
        if ngay and chuyen:
            conn = connect_db()
            query = f"SELECT Distinct STYLE FROM [INCENTIVE].[dbo].[SL_CA_NHAN] WHERE NGAY='{ngay}' AND CHUYEN='{chuyen}'"
            # app.logger.info(query)
            cursor = execute_query(conn, query)
            result = cursor.fetchall()
            close_db(conn)
            return [style[0] for style in result]
        else:
            return []
    except:
        return []

def laydulieuthuongmaychitiet():
    conn = connect_db()
    query = f"SELECT * FROM [INCENTIVE].[dbo].[THUONG_CN_MAY_HANG_NGAY_CHI_TIET] order by Ngay desc, Chuyen Asc, Cast(MaSoThe as int) Asc"
    cursor = execute_query(conn, query) 
    result = cursor.fetchall()
    close_db(conn)
    return result

def lay_danhsach_congnhan_trongchuyen(chuyen):
    conn = connect_db()
    query = f"SELECT * FROM [INCENTIVE].[dbo].[DS_CN_MAY_THEO_HC_CHUYEN] WHERE Chuyen='{chuyen}'"
    cursor = execute_query(conn, query) 
    result = cursor.fetchall()
    close_db(conn)
    return list(result)
    
def lay_danhsach_chuyen_hotro(chuyen):
    try:
        if chuyen:
            conn = connect_db()
            query = f"SELECT * FROM [INCENTIVE].[dbo].[DS_CHUYEN_MAY] WHERE LINE LIKE '{chuyen[0]}%' ORDER BY LINE"
            cursor = execute_query(conn, query) 
            result = cursor.fetchall()
            close_db(conn)
            return [line[0] for line in result]
        else:
            return []
    except:
        return []
  
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
        query += f"AND MA_CONG_DOAN ='{macongdoan}' "
    query += "ORDER BY CAST(MST as INT) ASC, MA_CONG_DOAN ASC"
    cursor = execute_query(conn, query) 
    result = cursor.fetchall()
    close_db(conn)
    return list(result)

def capnhat_sanluong(mst,hoten,chuyen,ngay,style,macongdoan,sanluong):
    conn = connect_db()
    query = f"INSERT INTO [INCENTIVE].[dbo].[SL_CA_NHAN] (MST,HO_TEN,CHUYEN,NGAY,STYLE,MA_CONG_DOAN,SL_CA_NHAN) VALUES('{mst}', N'{hoten}', '{chuyen}', '{ngay}', '{style}', '{macongdoan}', '{sanluong}')"
    # app.logger.info(query)
    execute_query(conn, query)
    try:
        conn.commit()
        close_db(conn)
        return True
    except Exception as e:
        app.logger.info(e)
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
        app.logger.info(e)
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
        if giodieuchuyendi:
            query = f"insert into [INCENTIVE].[dbo].[CN_MAY_DI_HO_TRO] values ('{nhamay}','{mst}',N'{hoten}',N'{chucdanh}','{chuyen}','{chuyendihotro}','{ngaydieuchuyendi}','{giodieuchuyendi}','{sogiohotro}')"
        else:
            query = f"insert into [INCENTIVE].[dbo].[CN_MAY_DI_HO_TRO] values ('{nhamay}','{mst}',N'{hoten}',N'{chucdanh}','{chuyen}','{chuyendihotro}','{ngaydieuchuyendi}',NULL,'{sogiohotro}')"
        # app.logger.info(query)
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
        # app.logger.info(query)
        execute_query(conn, query)
        conn.commit()
        close_db(conn)
        return True
    except:
        return False
    
def laytongsanluongtheocongdoan(ngay,chuyen,style,macongdoan):
    try:
        conn = connect_db()
        if macongdoan:
            query = f"select MA_CONG_DOAN,QTY from [INCENTIVE].[dbo].[TONG_SL_CONG_DOAN] where NGAY='{ngay}' and CHUYEN='{chuyen}' and STYLE='{style}' and MA_CONG_DOAN='{macongdoan}' group by MA_CONG_DOAN,QTY"
        else:
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
    
def lay_sanluong_tong_theochuyen(ngay, chuyen, style):
    try:
        if ngay and chuyen and style:
            conn = connect_db()
            query = f"select QTY from [INCENTIVE].[dbo].[SL_NGAY_CHUYEN_STYLE ] where NGAY='{ngay}' and CHUYEN='{chuyen}' and GR_STYLE='{style}'"
            # app.logger.info(query)
            result = execute_query(conn, query).fetchone()
            close_db(conn)
            return result[0]
        else:
            return 0
    except:
        return 0

def lay_baocao_thuong_congnhan_may(macongty,mst,ngay,chuyen):
    try:
        conn = connect_db()
        query = f"SELECT MST,HO_TEN,CHUYEN,NGAY,SAH,SCP,SO_GIO,Eff_CA_NHAN,THUONG_CA_NHAN FROM [INCENTIVE].[dbo].[INCENTIVE_CN_MAY_HANG_NGAY] WHERE 1=1" 
        if macongty:
            query += f" AND CHUYEN LIKE '{macongty}%'"
        if mst:
            query += f" AND MST='{mst}'"
        if ngay:
            query += f" AND NGAY='{ngay}'"
        if chuyen:
            query += f" AND CHUYEN LIKE '%{chuyen}%'"
        query += " ORDER BY NGAY DESC, CHUYEN ASC"
        # app.logger.info(query)
        rows = execute_query(conn, query).fetchall()
        close_db(conn)
        return rows
    except:
        return []
    
def lay_baocao_thuong_congnhan_nhommay(macongty,ngay,chuyen,style):
    try:
        conn = connect_db()
        query = f"SELECT Workdate,Line,Sah,Total_hours,Eff,Style,TRANG_THAI,CHUYEN_MOI,OQL,GR_INCENTIVE,GR_INCENTIVE_TOPUP1,GR_INCENTIVE_TOPUP2,TONG_THUONG FROM [INCENTIVE].[dbo].[THUONG_NHOM_MAY_HANG_NGAY] WHERE 1=1"
        if macongty:
            query += f" AND Line LIKE '{macongty}%'"
        if ngay:
            query += f" AND Workdate='{ngay}'"
        if chuyen:
            query += f" AND Line LIKE '%{chuyen}%'" 
        if style:
            query += f" AND Style LIKE '%{style}%'"
        query += "ORDER BY Workdate DESC, Line ASC"
        rows = execute_query(conn, query).fetchall()
        close_db(conn)
        return rows
    except:
        return []
    
def lay_baocao_sogio_lamviec(macongty,mst,ngay,chuyen):
    try:
        conn = connect_db()
        query = f"SELECT MST,HO_TEN,CHUYEN,NGAY,SO_GIO,CHUC_VU FROM [INCENTIVE].[dbo].[TGLV_1] WHERE 1=1" 
        if macongty:
            query += f" AND CHUYEN LIKE '{macongty}%'"
        if mst:
            query += f" AND MST='{mst}'"
        if ngay:
            query += f" AND NGAY='{ngay}'"
        if chuyen:
            query += f" AND CHUYEN LIKE '%{chuyen}%'"
        query += " ORDER BY NGAY DESC, CHUYEN ASC"
        # app.logger.info(query)
        rows = execute_query(conn, query).fetchall()
        close_db(conn)
        return rows
    except:
        return []
    
def lay_baocao_sanluong_canhan(macongty,mst,ngay,chuyen):
    try:
        conn = connect_db()
        query = f"SELECT * FROM [INCENTIVE].[dbo].[SL_CA_NHAN_1] WHERE 1=1" 
        if macongty:
            query += f" AND CHUYEN LIKE '{macongty}%'"
        if mst:
            query += f" AND MST='{mst}'"
        if ngay:
            query += f" AND NGAY='{ngay}'"
        if chuyen:
            query += f" AND CHUYEN LIKE '%{chuyen}%'"
        query += " ORDER BY NGAY DESC, CHUYEN ASC, MST ASC"
        # app.logger.info(query)
        rows = execute_query(conn, query).fetchall()
        close_db(conn)
        return rows
    except:
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
            lines = get_line(current_user.masothe, current_user.macongty)
            # app.logger.info(lines)
            if lines:
                if len(lines) == 1:
                    g.notice = {"line":lines,"role":"tt"}
                elif len(lines) > 1:
                    g.notice = {"line":lines,"role":"tk"}
            else:
                g.notice = {"line":[],"role":""}
        else:
            g.notice = {"line":[],"role":""}
    except:
        g.notice = {"line":[],"role":""}
        
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
@login_required
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
        chuyen = request.args.get('chuyen')
        style = request.args.get("style")
        mst = request.args.get("mst")
        hoten = request.args.get("hoten")
        macongdoan = request.args.get("search_macongdoan")
        styles = get_all_styles(ngay, chuyen)
        sanluongtong = lay_sanluong_tong_theochuyen(ngay, chuyen, style)
        if g.notice['role'] == 'tt':
            chuyen = g.notice['line'][0]
        danhsach_congnhan_hotro = lay_danhsach_congnhan_trongchuyen(chuyen)
        danhsach_chuyen = lay_danhsach_chuyen_hotro(chuyen)
        danhsach_sanluong = lay_danhsach_sanluong(ngay, chuyen, style,mst,hoten,macongdoan)
        danhsach_tnc = lay_danhsach_tnc_chua_lenchuyen(chuyen)
        danhsach_di_hotro = lay_danhsach_di_hotro(chuyen)
        return render_template("home.html",styles=styles,danhsach_sanluong=danhsach_sanluong,
                               danhsach_congnhan_hotro=danhsach_congnhan_hotro,
                               danhsach_chuyen=danhsach_chuyen,danhsach_tnc=danhsach_tnc,
                               danhsach_di_hotro=danhsach_di_hotro,sanluongtong=sanluongtong)
    elif request.method == "POST":
        try:
            ngay = request.form.get("ngay")   
            chuyen = request.form.get('chuyen')
            style = request.form.get("style")
            mst = request.form.get("mst")
            hoten = request.form.get("hoten")
            macongdoan = request.form.get("search_macongdoan")
            danhsach_sanluong = lay_danhsach_sanluong(ngay, chuyen, style,mst,hoten,macongdoan)
            data = [{
                "Mã số thẻ": int(row[0]),
                "Họ tên": row[1],
                "Chuyền": row[2],
                "Ngày": datetime.datetime.strptime(row[3], "%Y-%m-%d").strftime("%d/%m/%Y"),
                "Style": row[4],
                "Mã công đoạn": int(row[5]) if row[5] else 0,
                "Sản lượng": int(row[6]) if row[6] else 0,
            } for row in danhsach_sanluong]
            df = DataFrame(data)
            output = BytesIO()
            with ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)

            # Điều chỉnh độ rộng cột
            output.seek(0)
            workbook = load_workbook(output)
            sheet = workbook.active

            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                sheet.column_dimensions[column_letter].width = adjusted_width

            output = BytesIO()
            workbook.save(output)
            output.seek(0)
            time_stamp = datetime.datetime.now().strftime("%d%m%Y%H%M%S")
            # Trả file về cho client
            response = make_response(output.read())
            response.headers['Content-Disposition'] = f'attachment; filename={ngay}_{chuyen}_{style}_{time_stamp}.xlsx'
            response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            return response    
        except Exception as e:
            app.logger.error(f'Không thế tạo bảng {e} !!!')
            return redirect("/")    
    
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
        return redirect(f"/?chuyen={chuyen}&ngay={ngay}&style={style}&search_macongdoan={macongdoan}")
    
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
        sogiohotro = request.form.get("sogiohotro").replace(",", ".")
        nhamay = 'NT1' if chuyen[0]=="1" else 'NT2'
        them_nguoi_di_hotro(nhamay,chuyen,mst,hoten,chucdanh,chuyendihotro,ngaydieuchuyendi,giodieuchuyendi,sogiohotro)
        ngay = request.form.get("ngay")   
        chuyen = request.args.get('chuyen')
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
        chuyen = request.args.get('chuyen')
        style = request.form.get("style")
        return redirect(f"/?chuyen={chuyen}&ngay={ngay}&style={style}")

@app.route("/laytongsanluongtheocongdoan", methods=["POST"])
@login_required
def laytongsanluong():
    if request.method == "POST":
        ngay = request.args.get("ngay")
        chuyen = request.args.get("chuyen")
        style = request.args.get("style")
        macongdoan = request.args.get("macongdoan")
        data = laytongsanluongtheocongdoan(ngay,chuyen,style,macongdoan)
        return jsonify(data)
    
@app.route("/capnhatsogiohotro", methods=["POST"])
@login_required
def capnhatsogiohotro():
    if request.method == "POST":
        id = request.form.get("id_hotro")
        sogio = request.form.get("sogio")
        capnhat_sogio_hotro(id,sogio)
        ngay = request.form.get("ngay")   
        chuyen = request.args.get('chuyen')
        style = request.form.get("style")
        return redirect(f"/?chuyen={chuyen}&ngay={ngay}&style={style}")

@app.route("/xoasanluongcanhan", methods=["POST"])
@login_required
def xoasanluongcanhan():
    if request.method == "POST":
        id = request.form.get("id_xoasanluong")
        ngay = request.form.get("ngay")   
        chuyen = request.form.get('chuyen')
        style = request.form.get("style")
        xoa_sanluong(id)
        return redirect(f"/?chuyen={chuyen}&ngay={ngay}&style={style}")
        
@app.route("/taidulieulen", methods=["POST"])
def taidulieulen():
    if request.method == "POST":
        try:
            file = request.files["file"]
            if not file:
                app.logger.info("No file")
                return redirect("/")
            thoigian = datetime.datetime.now().strftime("%d%m%Y%H%M%S")
            filepath = f"tailen/data_{thoigian}.xlsx"
            file.save(filepath)
            data = read_excel(filepath).to_dict(orient="records")
            for row in data:
                capnhat_sanluong(
                    row["Mã số thẻ"],
                    row["Họ tên"],
                    row["Chuyền"],
                    datetime.datetime.strptime(row["Ngày"],"%d/%m/%Y").strftime("%Y-%m-%d"),
                    row["Style"],
                    row["Mã công đoạn"],
                    row["Sản lượng"]
                )
            ngay = request.form.get("ngay")   
            chuyen = request.args.get('chuyen')
            style = request.form.get("style")
            return redirect(f"/?chuyen={chuyen}&ngay={ngay}&style={style}")
        except Exception as e:
            app.logger.info(e)
            ngay = request.form.get("ngay")   
            chuyen = request.args.get('chuyen')
            style = request.form.get("style")
            return redirect(f"/?chuyen={chuyen}&ngay={ngay}&style={style}")

@app.route("/baocao_thuong_may", methods=["GET","POST"])
def baocao_may():
    if request.method == "GET":
        try:
            macongty = request.args.get("macongty")
            mst = request.args.get("mst")
            ngay = request.args.get("ngay")
            chuyen = request.args.get("chuyen")
            danhsach = lay_baocao_thuong_congnhan_may(macongty,mst,ngay,chuyen)
            page = request.args.get(get_page_parameter(), type=int, default=1)
            per_page = 10
            total = len(danhsach)
            start = (page - 1) * per_page
            end = start + per_page
            paginated_rows = danhsach[start:end]
            pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
            return render_template("baocao_thuong_may.html", danhsach=paginated_rows,pagination=pagination)
        except Exception as e:
            app.logger.error(e)
            return render_template("baocao_thuong_may.html", danhsach=[])
    elif request.method == "POST":
        try:
            macongty = request.form.get("macongty")
            mst = request.form.get("mst")
            ngay = request.form.get("ngay")
            chuyen = request.form.get("chuyen")
            danhsach = lay_baocao_thuong_congnhan_may(macongty,mst,ngay,chuyen)
            data = []
            for row in danhsach:
                data.append({
                    "Mã số thẻ": row[0],
                    "Họ tên":row[1],
                    "Chuyền":row[2],
                    "Ngày":datetime.datetime.strptime(row[3],"%Y-%m-%d").strftime("%d/%m/%Y"),
                    "SAH":round(row[4],2) if row[4] else "",
                    "SCP":row[5],
                    "Số giờ":row[6],
                    "Hiệu suất":f"{round(row[7]*100)} %" if row[7] else "",
                    "Thưởng": chuyen_so_thanh_sotien(row[8]) if row[8] else ""
                })
            df = DataFrame(data)
            output = BytesIO()
            with ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)

            # Điều chỉnh độ rộng cột
            output.seek(0)
            workbook = load_workbook(output)
            sheet = workbook.active

            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                sheet.column_dimensions[column_letter].width = adjusted_width

            output = BytesIO()
            workbook.save(output)
            output.seek(0)
            time_stamp = datetime.datetime.now().strftime("%d%m%Y%H%M%S")
            # Trả file về cho client
            response = make_response(output.read())
            response.headers['Content-Disposition'] = f'attachment; filename=baocaothuongmay_{time_stamp}.xlsx'
            response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            return response  
        except Exception as e:
            app.logger.error(e)
            return redirect("/baocao_thuong_may")
            
@app.route("/baocao_thuong_nhommay", methods=["GET","POST"])
def baocao_nhommay():
    if request.method == "GET":
        try:
            macongty = request.args.get("macongty")
            ngay = request.args.get("ngay")
            chuyen = request.args.get("chuyen")
            style = request.args.get("style")
            danhsach = lay_baocao_thuong_congnhan_nhommay(macongty,ngay,chuyen,style)
            page = request.args.get(get_page_parameter(), type=int, default=1)
            per_page = 10
            total = len(danhsach)
            start = (page - 1) * per_page
            end = start + per_page
            paginated_rows = danhsach[start:end]
            pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
            return render_template("baocao_thuong_nhommay.html", danhsach=paginated_rows,pagination=pagination)
        except Exception as e:
            app.logger.error(e)
            return render_template("baocao_thuong_may.html", danhsach=[])
    elif request.method == "POST":
        try:
            macongty = request.form.get("macongty")
            ngay = request.form.get("ngay")
            chuyen = request.form.get("chuyen")
            style = request.form.get("style")
            danhsach = lay_baocao_thuong_congnhan_nhommay(macongty,ngay,chuyen,style)
            data = []
            for row in danhsach:
                data.append({
                    "Ngày": datetime.datetime.strptime(row[0],"%Y-%m-%d").strftime("%d/%m/%Y"),
                    "Chuyền":row[1],
                    "SAH":round(row[2],2) if row[2] else "",
                    "Số giờ":row[3],
                    "Hiệu suất":f"{round(row[4]*100)} %" if row[4] else "",
                    "Style": row[5],
                    "Trạng thái đơn hàng": row[6],
                    "Chuyền mới": row[7],
                    "OQL": row[8],
                    "Thưởng nhóm": chuyen_so_thanh_sotien(row[9]) if row[9] else "",
                    "Thưởng 1": chuyen_so_thanh_sotien(row[10]) if row[10] else "",
                    "Thưởng 2": chuyen_so_thanh_sotien(row[11]) if row[11] else "",
                    "Tổng thưởng": chuyen_so_thanh_sotien(row[12]) if row[12] else ""
                })
            df = DataFrame(data)
            output = BytesIO()
            with ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)

            # Điều chỉnh độ rộng cột
            output.seek(0)
            workbook = load_workbook(output)
            sheet = workbook.active

            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                sheet.column_dimensions[column_letter].width = adjusted_width

            output = BytesIO()
            workbook.save(output)
            output.seek(0)
            time_stamp = datetime.datetime.now().strftime("%d%m%Y%H%M%S")
            # Trả file về cho client
            response = make_response(output.read())
            response.headers['Content-Disposition'] = f'attachment; filename=baocaothuongnhommay_{time_stamp}.xlsx'
            response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            return response  
        except Exception as e:
            app.logger.error(e)
            return redirect("/baocao_thuong_nhommay")
        
@app.route("/baocao_sogio_lamviec", methods=["GET", "POST"])
def baocao_sogio_lamviec():
    if request.method == "GET":
        try:
            macongty = request.args.get("macongty")
            mst = request.args.get("mst")
            ngay = request.args.get("ngay")
            chuyen = request.args.get("chuyen")
            danhsach = lay_baocao_sogio_lamviec(macongty,mst,ngay,chuyen)
            page = request.args.get(get_page_parameter(), type=int, default=1)
            per_page = 10
            total = len(danhsach)
            start = (page - 1) * per_page
            end = start + per_page
            paginated_rows = danhsach[start:end]
            pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
            return render_template("baocao_sogio_lamviec.html", danhsach=paginated_rows,pagination=pagination)
        except Exception as e:
            app.logger.error(e)
            return render_template("baocao_sogio_lamviec.html", danhsach=[])
    elif request.method == "POST":
        try:
            macongty = request.form.get("macongty")
            mst = request.form.get("mst")
            ngay = request.form.get("ngay")
            chuyen = request.form.get("chuyen")
            danhsach = lay_baocao_sogio_lamviec(macongty,mst,ngay,chuyen)
            data = []
            for row in danhsach:
                data.append({
                    "Mã số thẻ": row[0],
                    "Họ tên":row[1],
                    "Chuyền":row[2],
                    "Ngày": datetime.datetime.strptime(row[3],"%Y-%m-%d").strftime("%d/%m/%Y"),
                    "Số giờ":row[4],
                    "Chức danh" : row[5],
                })
            df = DataFrame(data)
            output = BytesIO()
            with ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)

            # Điều chỉnh độ rộng cột
            output.seek(0)
            workbook = load_workbook(output)
            sheet = workbook.active

            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                sheet.column_dimensions[column_letter].width = adjusted_width

            output = BytesIO()
            workbook.save(output)
            output.seek(0)
            time_stamp = datetime.datetime.now().strftime("%d%m%Y%H%M%S")
            # Trả file về cho client
            response = make_response(output.read())
            response.headers['Content-Disposition'] = f'attachment; filename=baocaosogiolamviec_{time_stamp}.xlsx'
            response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            return response  
        except Exception as e:
            app.logger.error(e)
            return redirect("/baocao_sogio_lamviec")
    
@app.route("/baocao_sanluong_canhan", methods=["GET","POST"])
def baocao_sanluong_canhan():
    if request.method == "GET":
        try:
            macongty = request.args.get("macongty")
            mst = request.args.get("mst")
            ngay = request.args.get("ngay")
            chuyen = request.args.get("chuyen")
            danhsach = lay_baocao_sanluong_canhan(macongty,mst,ngay,chuyen)
            page = request.args.get(get_page_parameter(), type=int, default=1)
            per_page = 10
            total = len(danhsach)
            start = (page - 1) * per_page
            end = start + per_page
            paginated_rows = danhsach[start:end]
            pagination = Pagination(page=page, per_page=per_page, total=total, css_framework='bootstrap4')
            return render_template("baocao_sanluong_canhan.html", danhsach=paginated_rows,pagination=pagination)
        except Exception as e:
            app.logger.error(e)
            return render_template("baocao_sanluong_canhan.html", danhsach=[])
    elif request.method == "POST":
        try:
            macongty = request.form.get("macongty")
            mst = request.form.get("mst")
            ngay = request.form.get("ngay")
            chuyen = request.form.get("chuyen")
            danhsach = lay_baocao_sanluong_canhan(macongty,mst,ngay,chuyen)
            data = []
            for row in danhsach:
                data.append({
                    "Mã số thẻ": row[0],
                    "Họ tên": row[1],
                    "Chuyền":row[2],
                    "Ngày": datetime.datetime.strptime(row[3],"%Y-%m-%d").strftime("%d/%m/%Y"),
                    "Style":row[4],
                    "Mã công đoạn" : row[5],
                    "Sản lượng": row[6]
                })
            df = DataFrame(data)
            output = BytesIO()
            with ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)

            # Điều chỉnh độ rộng cột
            output.seek(0)
            workbook = load_workbook(output)
            sheet = workbook.active

            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                sheet.column_dimensions[column_letter].width = adjusted_width

            output = BytesIO()
            workbook.save(output)
            output.seek(0)
            time_stamp = datetime.datetime.now().strftime("%d%m%Y%H%M%S")
            # Trả file về cho client
            response = make_response(output.read())
            response.headers['Content-Disposition'] = f'attachment; filename=baocaosanluongcanhan_{time_stamp}.xlsx'
            response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            return response  
        except Exception as e:
            app.logger.error(e)
            return redirect("/baocao_sanluong_canhan")
@app.route("/taithuongchitiet", methods=["GET","POST"])
def taithuongchitiet():
    try:
        danhsach = laydulieuthuongmaychitiet()
        data = []
        for row in danhsach:
            data.append({
                "Ngày": datetime.datetime.strptime(row[0],"%Y-%m-%d").strftime("%d/%m/%Y"),
                "Mã số thẻ": row[1],
                "Họ tên": row[2],
                "Chuyền": row[3],
                "SAH nhóm": round(row[4],1) if row[4] else "",
                "Thòi gian làm việc nhóm": round(row[5]) if row[5] else "",
                "Hiệu suất nhóm": f"{round(row[6]*100)} %" if row[6] else "",
                "Style": row[7] if row[7] else "",
                "Chuyền mới": row[8] if row[8] else "",
                "Ngày vào chuyền": row[9] if row[9] else "",
                "Trạng thái": row[10] if row[10] else "",
                "UI": row[11] if row[11] else "",
                "BE": row[12] if row[12] else "",
                "BE TOPUP 1": row[13] if row[13] else "",
                "TOPUP 1": row[14] if row[14] else "",
                "BE TOPUP 2": row[15] if row[1] else "",
                "TOPUP 2": row[16] if row[16] else "",
                "OQL": f"{row[17]*100} %" if row[17] else "",
                "AQL": row[18] if row[18] else "",
                "Group incentive": round(row[19]) if row[19] else "",
                "Group incentive topup 1": round(row[20]) if row[20] else "",
                "Group incentive topup 2": round(row[21]) if row[21] else "",
                "Tổng thưởng": round(row[22]) if row[22] else "",
                "SAH": round(row[23],1) if row[23] else "",
                "Thời gian làm việc": row[24] if row[24] else "",
                "Hiệu suất": f"{round(row[25]*100)} %" if row[25] else "",
                "SCP": row[26] if row[26] else "",
                "Hệ số SCP": row[27] if row[27] else "",
                "Hệ số thưởng cá nhân": round(row[28],1) if row[28] else "",
                "Hệ số thưởng nhóm": round(row[29],1) if row[29] else "",
                "Thưởng cá nhân": round(row[30]) if row[30] else ""
            })
        df = DataFrame(data)
        output = BytesIO()
        with ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        # Điều chỉnh độ rộng cột
        output.seek(0)
        workbook = load_workbook(output)
        sheet = workbook.active

        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column_letter].width = adjusted_width

        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        time_stamp = datetime.datetime.now().strftime("%d%m%Y%H%M%S")
        # Trả file về cho client
        response = make_response(output.read())
        response.headers['Content-Disposition'] = f'attachment; filename=thuongmaychitiet_{time_stamp}.xlsx'
        response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        return response  
    except Exception as e:
        app.logger.error(f"Loi lay thuong may chi tiet: {e}")
        return redirect("/")

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=80)