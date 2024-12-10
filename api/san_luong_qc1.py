from flask import render_template, request, Blueprint, redirect
from flask_login import current_user
from flask_paginate import Pagination, get_page_parameter
import pandas as pd
from helper.utils import *
from helper.nhap_excel import get_data, get_excel_from_table, upload_excel_to_db
from datetime import datetime

qc1 = Blueprint('san_luong_qc1', __name__)

SIZE = 20
@qc1.route("/san_luong_qc1", methods=["GET"])
def san_luong_qc1():
    try:
        if "IED" not in current_user.phongban:
            return redirect("/")

        ngay = request.args.get("ngay")
        mst = request.args.get("mst")
        page = request.args.get(get_page_parameter(), type=int, default=1)
        filters = {
            "nha_may": {
                "type": "equal",
                "value": current_user.macongty
            },
            "mst": {
                "type": "equal",
                "value": mst
            },
            "ngay": {
                "type": "equal",
                "value": ngay
            }
        }
        data, total = get_data(filters, page, SIZE, "[INCENTIVE].[dbo].[SL_CN_QC1_NHAP_TAY]", "NGAY DESC").values()
        for row in data:
            row_list = list(row)
            for i in range(len(row_list)):
                if row_list[i] is None:
                    row_list[i] = ""
            row_list[2] = datetime.strftime(datetime.strptime(row_list[2], "%Y-%m-%d"), "%d/%m/%Y") if row_list[2] else ""
            data[data.index(row)] = tuple(row_list)
        pagination = Pagination(page=page, per_page=SIZE, total=total, css_framework='bootstrap4')
        return render_template("san_luong_qc1.html", danhsach=data, pagination=pagination)
    except Exception as e:
        print(e)
        return render_template("san_luong_qc1.html")
        
@qc1.route("/san_luong_qc1/excel", methods=["GET"])
def get_excel():
    filters = {
        "nha_may": {
            "type": "equal",
            "value": current_user.macongty
        },
        "ngay": {
            "type": "equal",
            "value": request.args.get("ngay")
        },
        "mst": {
            "type": "equal",
            "value": request.args.get("mst")
        }
    }
    return get_excel_from_table("INCENTIVE", "SL_CN_QC1_NHAP_TAY", "san_luong_qc1", filters)
    
@qc1.route("/san_luong_qc1/upload_excel", methods=["POST"])
def upload_excel():
    try:
        if 'file' not in request.files:
            return None
        
        file = request.files["file"]
        upload_excel_to_db("INCENTIVE", "SL_CN_QC1_NHAP_TAY", file)
        return redirect("/san_luong_qc1")
    except Exception as e:
        print(e)
        return None

@qc1.route("/san_luong_qc1/filter", methods=["POST"])
def filter():
    try:
        mst = request.form.get("mst")
        ngay = request.form.get("ngay")
        return redirect(f"/san_luong_qc1?mst={mst}&ngay={ngay}")
    except Exception as e:
        print(e)
        return None