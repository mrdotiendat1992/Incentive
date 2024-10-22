from flask import render_template, request, Blueprint, make_response, redirect
from flask_paginate import Pagination, get_page_parameter
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment
from io import BytesIO
from datetime import datetime
import pandas as pd
from helper.utils import *
from helper.nhap_excel import get_data, get_excel_from_table, upload_excel_to_db

tnc = Blueprint('hieusuat_tnc', __name__)

SIZE = 20

@tnc.route("/hieusuat_tnc", methods=["GET"])
def danhsach_totruong():
    try:    
        ngay = request.args.get("ngay")
        chuyen = request.args.get("chuyen")
        mst = request.args.get("mst")
        page = request.args.get(get_page_parameter(), type=int, default=1)
        filters = {
            "ngay": ngay,
            "chuyen": chuyen,
            "mst": mst
        }
        
        data, total = get_data(filters, page, SIZE, "[INCENTIVE].[dbo].[HIEU_SUAT_CN_TNC]").values()
        for row in data:
            row_list = list(row)
            row_list[4] = datetime.strftime(datetime.strptime(row_list[4], "%Y-%m-%d"), "%d/%m/%Y")
            data[data.index(row)] = tuple(row_list)

        pagination = Pagination(page=page, per_page=SIZE, total=total, css_framework='bootstrap4')
        return render_template("hieusuat_tnc.html", danhsach=data, pagination=pagination)
    except Exception as e:
        print(e)
        return render_template("hieusuat_tnc.html")
        
@tnc.route("/hieusuat_tnc/excel", methods=["GET"])
def get_excel():
    return get_excel_from_table("INCENTIVE", "HIEU_SUAT_CN_TNC", "hieusuat_tnc")
    
@tnc.route("/hieusuat_tnc/upload_excel", methods=["POST"])
def upload_excel():
    try:
        if 'file' not in request.files:
            return None
        
        file = request.files["file"]
        upload_excel_to_db("INCENTIVE", "HIEU_SUAT_CN_TNC", file)
        return redirect("/hieusuat_tnc")
    except Exception as e:
        print(e)
        return None
    
@tnc.route("/hieusuat_tnc/filter", methods=["POST"])
def filter():
    try:
        mst = request.form.get("mst")
        ngay = request.form.get("ngay")
        chuyen = request.form.get("chuyen")
        return redirect(f"/hieusuat_tnc?mst={mst}&ngay={ngay}&chuyen={chuyen}")
    except Exception as e:
        print(e)
        return None