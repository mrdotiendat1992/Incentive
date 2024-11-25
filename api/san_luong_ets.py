from flask import render_template, request, Blueprint, redirect
from flask_login import current_user
from flask_paginate import Pagination, get_page_parameter
import pandas as pd
from helper.utils import *
from helper.nhap_excel import get_data, get_excel_from_table, upload_excel_to_db
from datetime import datetime

ets = Blueprint('san_luong_ets', __name__)

SIZE = 20
@ets.route("/san_luong_ets", methods=["GET"])
def san_luong_ets():
    try:
        if "IED" not in current_user.phongban:
            return redirect("/")

        ngay = request.args.get("ngay")
        chuyen = request.args.get("chuyen")
        style = request.args.get("style")
        page = request.args.get(get_page_parameter(), type=int, default=1)
        filters = {
            "workdate": {
                "type": "equal",
                "value": ngay
            },
            "line": {
                "type": "approximately",
                "value": chuyen
            },
            "style_a": {
                "type": "approximately",
                "value": style
            }
        }
        data, total = get_data(filters, page, SIZE, "[DW].[dbo].[ETS_Qty_NHAP_TAY]", "WORKDATE").values()
        for row in data:
            row_list = list(row)
            for i in range(len(row_list)):
                if row_list[i] is None:
                    row_list[i] = ""
            row_list[0] = datetime.strftime(datetime.strptime(row_list[0], "%Y-%m-%d"), "%d/%m/%Y")
            data[data.index(row)] = tuple(row_list)
            
        pagination = Pagination(page=page, per_page=SIZE, total=total, css_framework='bootstrap4')
        return render_template("san_luong_ets.html", danhsach=data, pagination=pagination)
    except Exception as e:
        print(e)
        return render_template("san_luong_ets.html")
        
@ets.route("/san_luong_ets/excel", methods=["GET"])
def get_excel():
    filters = {
        "workdate": {
            "type": "equal",
            "value": request.args.get("ngay")
        },
        "line": {
            "type": "approximately",
            "value": request.args.get("chuyen")
        },
        "style_a": {
            "type": "approximately",
            "value": request.args.get("style")
        }
    }
    return get_excel_from_table("DW", "ETS_Qty_NHAP_TAY", "san_luong_ets", filters)
    
@ets.route("/san_luong_ets/upload_excel", methods=["POST"])
def upload_excel():
    try:
        if 'file' not in request.files:
            return None
        
        file = request.files["file"]
        upload_excel_to_db("DW", "ETS_Qty_NHAP_TAY", file)
        return redirect("/san_luong_ets")
    except Exception as e:
        print(e)
        return None

@ets.route("/san_luong_cat/filter", methods=["POST"])
def filter():
    try:
        ngay = request.args.get("ngay")
        chuyen = request.args.get("chuyen")
        style = request.args.get("style")
        return redirect(f"/san_luong_cat?chuyen={chuyen}&ngay={ngay}&style={style}")
    except Exception as e:
        print(e)
        return None