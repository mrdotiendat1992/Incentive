from flask import render_template, request, Blueprint, redirect
from flask_login import current_user
from flask_paginate import Pagination, get_page_parameter
import pandas as pd
from helper.utils import *
from helper.nhap_excel import get_data, get_excel_from_table, upload_excel_to_db
from datetime import datetime

donhang = Blueprint('trang_thai_don_hang', __name__)

SIZE = 20
@donhang.route("/trang_thai_don_hang", methods=["GET"])
def trang_thai_don_hang():
    try:    
        if "IED" not in current_user.phongban:
            return redirect("/")

        ngay = request.args.get("ngay")
        chuyen = request.args.get("chuyen")
        style = request.args.get("style")
        page = request.args.get(get_page_parameter(), type=int, default=1)
        filters = {
            "ngay": {
                "type": "equal",
                "value": ngay
            },
            "chuyen": {
                "type": "approximately",
                "value": chuyen
            },
            "style": {
                "type": "approximately",
                "value": style
            }
        }
        data, total = get_data(filters, page, SIZE, "[INCENTIVE].[dbo].[TRANG_THAI_DON_HANG]", "NGAY DESC").values()
        for row in data:
            row_list = list(row)
            row_list[0] = datetime.strftime(datetime.strptime(row_list[0], "%Y-%m-%d"), "%d/%m/%Y")
            data[data.index(row)] = tuple(row_list)
        pagination = Pagination(page=page, per_page=SIZE, total=total, css_framework='bootstrap4')
        return render_template("trang_thai_don_hang.html", danhsach=data, pagination=pagination)
    except Exception as e:
        print(e)
        return render_template("trang_thai_don_hang.html")
        
@donhang.route("/trang_thai_don_hang/excel", methods=["GET"])
def get_excel():
    return get_excel_from_table("INCENTIVE", "TRANG_THAI_DON_HANG", "trang_thai_don_hang")
    
@donhang.route("/trang_thai_don_hang/upload_excel", methods=["POST"])
def upload_excel():
    try:
        if 'file' not in request.files:
            return None
        
        file = request.files["file"]
        upload_excel_to_db("INCENTIVE", "TRANG_THAI_DON_HANG", file)
        return redirect("/trang_thai_don_hang")
    except Exception as e:
        print(e)
        return None

@donhang.route("/trang_thai_don_hang/filter", methods=["POST"])
def filter():
    try:
        ngay = request.form.get("ngay")
        chuyen = request.form.get("chuyen")
        style = request.form.get("style")
        return redirect(f"/trang_thai_don_hang?ngay={ngay}&chuyen={chuyen}&style={style}")
    except Exception as e:
        print(e)
        return None