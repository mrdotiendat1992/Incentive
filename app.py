from flask import Flask, render_template, request, url_for, redirect, g, flash, jsonify, send_file, session, flash, get_flashed_messages, render_template_string
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, current_user, login_required
from flask_paginate import Pagination, get_page_parameter
import pyodbc
import openpyxl
import pandas as pd
from datetime import datetime, timedelta
import os
from functools import wraps
import logging
from logging.handlers import RotatingFileHandler
import numpy as np
import urllib.parse
from waitress import serve
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
                    app.logger.info(f"Nguoi dung {masothe} o {macongty} vua  dang nhap thanh cong !!!")
                    return redirect(url_for('home'))
            return redirect(url_for("login"))
        except Exception as e:
            app.logger.error(f'Nguoi dung {masothe} o {macongty} dang nhap that bai: {e} !!!')
            return redirect(url_for("login"))
    return render_template("login.html")

@app.route("/logout", methods=["POST"])
def logout():
    try:
        app.logger.info(f"Nguoi dung {current_user.masothe} o {current_user.macongty} vua  dang xuat !!!")
        logout_user()
    except Exception as e:
        app.logger.error(f'Không thế đăng xuất {e} !!!')
        flash(f'Không thế đăng xuất {e} !!!')
    return redirect("/")

@app.route("/", methods=['GET','POST'])
@login_required
def home():
    if request.method == "GET":
        flash(f"Xin chào {current_user.hoten} !!!")
        return render_template("home.html")
    
if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=83)
    