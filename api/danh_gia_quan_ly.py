from flask import render_template, request, Blueprint, redirect
from flask_login import current_user
from flask_paginate import Pagination, get_page_parameter
import pandas as pd
from helper.utils import *
from helper.nhap_excel import get_data, get_excel_from_table, upload_excel_to_db
from datetime import datetime

quanly = Blueprint('danh_gia_quan_ly', __name__)

SIZE = 20
@quanly.route("/danh_gia_quan_ly", methods=["GET"])
def danh_gia_quan_ly():
    try:    
        if "IED" not in current_user.phongban:
            return redirect("/")
        
        ngay = request.args.get("ngay")
        nhamay = request.args.get("nhamay")
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
        data, total = get_data(filters, page, SIZE, "[INCENTIVE].[dbo].[DANH_GIA_QUAN_LY]", "NGAY DESC").values()
        for row in data:
            row_list = list(row)
            row_list[2] = datetime.strftime(datetime.strptime(row_list[2], "%Y-%m-%d"), "%d/%m/%Y")
            data[data.index(row)] = tuple(row_list)
        pagination = Pagination(page=page, per_page=SIZE, total=total, css_framework='bootstrap4')
        return render_template("danh_gia_quan_ly.html", danhsach=data, pagination=pagination)
    except Exception as e:
        print(e)
        return render_template("danh_gia_quan_ly.html")
        
@quanly.route("/danh_gia_quan_ly/excel", methods=["GET"])
def get_excel():
    filters = {
        "nha_may": {
            "type": "equal",
            "value": current_user.macongty
        },
        "mst": {
            "type": "equal",
            "value": request.args.get("mst")
        },
        "ngay": {
            "type": "equal",
            "value": request.args.get("ngay")
        }
    }
    return get_excel_from_table("INCENTIVE", "DANH_GIA_QUAN_LY", "danh_gia_quan_ly", filters, ["ngay"])
    
@quanly.route("/danh_gia_quan_ly/upload_excel", methods=["POST"])
def upload_excel():
    try:
        if 'file' not in request.files:
            return None
        
        file = request.files["file"]
        upload_excel_to_db("INCENTIVE", "DANH_GIA_QUAN_LY", file)
        return redirect("/danh_gia_quan_ly")
    except Exception as e:
        print(e)
        return None

@quanly.route("/danh_gia_quan_ly/filter", methods=["POST"])
def filter():
    try:
        ngay = request.form.get("ngay")
        mst = request.form.get("mst")
        return redirect(f"/danh_gia_quan_ly?ngay={ngay}&mst={mst}")
    except Exception as e:
        print(e)
        return None