from flask import render_template, request, Blueprint, redirect
from flask_login import current_user
from flask_paginate import Pagination, get_page_parameter
import pandas as pd
from helper.utils import *
from helper.nhap_excel import get_data, get_excel_from_table, upload_excel_to_db
from datetime import datetime

tienthuong = Blueprint('tru_tien_thuong', __name__)

SIZE = 20
@tienthuong.route("/tru_tien_thuong", methods=["GET"])
def tru_tien_thuong():
    try:    
        if "IED" not in current_user.phongban:
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
        data, total = get_data(filters, page, SIZE, "[INCENTIVE].[dbo].[TRU_TIEN_THUONG_NHOM_MAY]", "NGAY DESC").values()
        for row in data:
            row_list = list(row)
            row_list[1] = datetime.strftime(datetime.strptime(row_list[1], "%Y-%m-%d"), "%d/%m/%Y")
            data[data.index(row)] = tuple(row_list)
        pagination = Pagination(page=page, per_page=SIZE, total=total, css_framework='bootstrap4')
        return render_template("tru_tien_thuong.html", danhsach=data, pagination=pagination)
    except Exception as e:
        print(e)
        return render_template("tru_tien_thuong.html")
        
@tienthuong.route("/tru_tien_thuong/excel", methods=["GET"])
def get_excel():
    filters = {
        "ngay": {
            "type": "equal",
            "value": request.args.get("ngay")
        },
        "chuyen": {
            "type": "approximately",
            "value": request.args.get("chuyen")
        },    
    }
    return get_excel_from_table("INCENTIVE", "TRU_TIEN_THUONG_NHOM_MAY", "tru_tien_thuong", filters, ["ngay"])
    
@tienthuong.route("/tru_tien_thuong/upload_excel", methods=["POST"])
def upload_excel():
    try:
        if 'file' not in request.files:
            return None
        
        file = request.files["file"]
        upload_excel_to_db("INCENTIVE", "TRU_TIEN_THUONG_NHOM_MAY", file)
        return redirect("/tru_tien_thuong")
    except Exception as e:
        print(e)
        return None

@tienthuong.route("/tru_tien_thuong/filter", methods=["POST"])
def filter():
    try:
        chuyen = request.form.get("chuyen")
        ngay = request.form.get("ngay")
        return redirect(f"/tru_tien_thuong?chuyen={chuyen}&ngay={ngay}")
    except Exception as e:
        print(e)
        return None