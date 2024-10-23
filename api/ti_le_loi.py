from flask import render_template, request, Blueprint, redirect
from flask_login import current_user
from flask_paginate import Pagination, get_page_parameter
import pandas as pd
from helper.utils import *
from helper.nhap_excel import get_data, get_excel_from_table, upload_excel_to_db
from datetime import datetime

loi = Blueprint('ti_le_loi', __name__)

SIZE = 20
@loi.route("/ti_le_loi", methods=["GET"])
def ti_le_loi():
    try:    
        if "QAD" not in current_user.phongban:
            return redirect("/")

        chuyen = request.args.get("chuyen")
        ngay = request.args.get("ngay")
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
        }
        data, total = get_data(filters, page, SIZE, "[INCENTIVE].[dbo].[TI_LE_LOI]", "NGAY DESC").values()
        for row in data:
            row_list = list(row)
            row_list[1] = datetime.strftime(datetime.strptime(row_list[1], "%Y-%m-%d"), "%d/%m/%Y")
            data[data.index(row)] = tuple(row_list)
        pagination = Pagination(page=page, per_page=SIZE, total=total, css_framework='bootstrap4')
        return render_template("ti_le_loi.html", danhsach=data, pagination=pagination)
    except Exception as e:
        print(e)
        return render_template("ti_le_loi.html")
        
@loi.route("/ti_le_loi/excel", methods=["GET"])
def get_excel():
    return get_excel_from_table("INCENTIVE", "TI_LE_LOI", "ti_le_loi", ["ngay"])
    
@loi.route("/ti_le_loi/upload_excel", methods=["POST"])
def upload_excel():
    try:
        if 'file' not in request.files:
            return None
        
        file = request.files["file"]
        upload_excel_to_db("INCENTIVE", "TI_LE_LOI", file)
        return redirect("/ti_le_loi")
    except Exception as e:
        print(e)
        return None

@loi.route("/ti_le_loi/filter", methods=["POST"])
def filter():
    try:
        chuyen = request.form.get("chuyen")
        ngay = request.form.get("ngay")
        return redirect(f"/ti_le_loi?chuyen={chuyen}&ngay={ngay}")
    except Exception as e:
        print(e)
        return None