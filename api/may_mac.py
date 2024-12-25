from flask import render_template, request, Blueprint, redirect
from flask_login import current_user
from flask_paginate import Pagination, get_page_parameter
import pandas as pd
from helper.utils import *
from helper.nhap_excel import get_data, get_excel_from_table, upload_excel_to_db
from datetime import datetime

maymac = Blueprint('cong_doan_may_mac', __name__)

SIZE = 20
@maymac.route("/cong_doan_may_mac", methods=["GET"])
def cong_doan_may_mac():
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
            "style": {
              "type": "approximately",
              "value": style
            }
        }
        data, total = get_data(filters, page, SIZE, "[DW].[dbo].[CONG_DOAN_MAY_MAC]", "WORKDATE DESC").values()
        for row in data:
            row_list = list(row)
            for i in range(len(row_list)):
                if row_list[i] is None:
                    row_list[i] = ""
            row_list[0] = datetime.strftime(datetime.strptime(row_list[0], "%Y-%m-%d"), "%d/%m/%Y") if row_list[0] else ""
            data[data.index(row)] = tuple(row_list)
            
        pagination = Pagination(page=page, per_page=SIZE, total=total, css_framework='bootstrap4')
        return render_template("cong_doan_may_mac.html", danhsach=data, pagination=pagination)
    except Exception as e:
        print(e)
        return render_template("cong_doan_may_mac.html")
        
@maymac.route("/cong_doan_may_mac/excel", methods=["GET"])
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
    print(filters)
    return get_excel_from_table("DW", "CONG_DOAN_MAY_MAC", "cong_doan_may_mac", filters)
    
@maymac.route("/cong_doan_may_mac/upload_excel", methods=["POST"])
def upload_excel():
    try:
        if 'file' not in request.files:
            return None
        
        file = request.files["file"]
        upload_excel_to_db("DW", "CONG_DOAN_MAY_MAC", file)
        return redirect("/cong_doan_may_mac")
    except Exception as e:
        print(e)
        return None

@maymac.route("/cong_doan_may_mac/filter", methods=["POST"])
def filter():
    try:
        ngay = request.form.get("ngay")
        chuyen = request.form.get("chuyen")
        style = request.form.get("style")
        print(ngay)
        return redirect(f"/cong_doan_may_mac?chuyen={chuyen}&ngay={ngay}&style={style}")
    except Exception as e:
        print(e)
        return None