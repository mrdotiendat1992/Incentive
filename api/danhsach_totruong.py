from flask import render_template, request, Blueprint, redirect
from flask_paginate import Pagination, get_page_parameter
import pandas as pd
from helper.utils import *
from helper.nhap_excel import get_data, get_excel_from_table, upload_excel_to_db

totruong = Blueprint('danhsach_totruong', __name__)

SIZE = 20
@totruong.route("/danhsach_totruong", methods=["GET"])
def danhsach_totruong():
    try:    
        nhamay = request.args.get("nhamay")
        chuyen = request.args.get("chuyen")
        mst = request.args.get("mst")
        page = request.args.get(get_page_parameter(), type=int, default=1)
        filters = {
            "nha_may": nhamay,
            "chuyen": chuyen,
            "mst": mst
        }
        data, total = get_data(filters, page, SIZE, "[INCENTIVE].[dbo].[DS_TO_TRUONG]").values()
        pagination = Pagination(page=page, per_page=SIZE, total=total, css_framework='bootstrap4')
        return render_template("danhsach_totruong.html", danhsach=data, pagination=pagination)
    except Exception as e:
        print(e)
        return render_template("danhsach_totruong.html")
        
@totruong.route("/danhsach_totruong/excel", methods=["GET"])
def get_excel():
    return get_excel_from_table("INCENTIVE", "DS_TO_TRUONG", "danhsach_totruong")
    
@totruong.route("/danhsach_totruong/upload_excel", methods=["POST"])
def upload_excel():
    try:
        if 'file' not in request.files:
            return None
        
        file = request.files["file"]
        upload_excel_to_db("INCENTIVE", "DS_TO_TRUONG", file)
        return redirect("/danhsach_totruong")
    except Exception as e:
        print(e)
        return None

@totruong.route("/danhsach_totruong/filter", methods=["POST"])
def filter():
    try:
        mst = request.form.get("mst")
        nhamay = request.form.get("nhamay")
        chuyen = request.form.get("chuyen")
        return redirect(f"/danhsach_totruong?mst={mst}&nhamay={nhamay}&chuyen={chuyen}")
    except Exception as e:
        print(e)
        return None