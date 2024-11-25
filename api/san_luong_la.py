from flask import render_template, request, Blueprint, redirect
from flask_login import current_user
from flask_paginate import Pagination, get_page_parameter
import pandas as pd
from helper.utils import *
from helper.nhap_excel import get_data, get_excel_from_table, upload_excel_to_db
from datetime import datetime

la = Blueprint('san_luong_la', __name__)

SIZE = 20
@la.route("/san_luong_la", methods=["GET"])
def san_luong_la():
    try:
        if "IED" not in current_user.phongban:
            return redirect("/")

        ngay = request.args.get("ngay")
        mst = request.args.get("mst")
        chuyen = request.args.get("chuyen")
        page = request.args.get(get_page_parameter(), type=int, default=1)
        filters = {
            "chuyen": {
                "type": "approximately",
                "value": chuyen
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
        data, total = get_data(filters, page, SIZE, "[INCENTIVE].[dbo].[SL_CN_LA_NHAP_TAY]", "NGAY").values()
        for row in data:
            row_list = list(row)
            for i in range(len(row_list)):
                if row_list[i] is None:
                    row_list[i] = ""
            row_list[1] = datetime.strftime(datetime.strptime(row_list[1], "%Y-%m-%d"), "%d/%m/%Y")
            data[data.index(row)] = tuple(row_list)
        pagination = Pagination(page=page, per_page=SIZE, total=total, css_framework='bootstrap4')
        return render_template("san_luong_la.html", danhsach=data, pagination=pagination)
    except Exception as e:
        print(e)
        return render_template("san_luong_la.html")
        
@la.route("/san_luong_la/excel", methods=["GET"])
def get_excel():
    filters = {
        "chuyen": {
            "type": "approximately",
            "value": request.args.get("chuyen")
        },
        "ngay": {
            "type": "approximately",
            "value": request.args.get("ngay")
        },
        "mst": {
            "type": "equal",
            "value": request.args.get("mst")
        }
    }
    return get_excel_from_table("INCENTIVE", "SL_CN_LA_NHAP_TAY", "san_luong_la", filters)
    
@la.route("/san_luong_la/upload_excel", methods=["POST"])
def upload_excel():
    try:
        if 'file' not in request.files:
            return None
        
        file = request.files["file"]
        upload_excel_to_db("INCENTIVE", "SL_CN_LA_NHAP_TAY", file)
        return redirect("/san_luong_la")
    except Exception as e:
        print(e)
        return None

@la.route("/san_luong_la/filter", methods=["POST"])
def filter():
    try:
        mst = request.form.get("mst")
        ngay = request.form.get("ngay")
        chuyen = request.form.get("chuyen")
        return redirect(f"/san_luong_la?mst={mst}&ngay={ngay}&chuyen={chuyen}")
    except Exception as e:
        print(e)
        return None