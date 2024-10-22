from flask import render_template, request, Blueprint, redirect
from flask_login import current_user
from flask_paginate import Pagination, get_page_parameter
import pandas as pd
from helper.utils import *
from helper.nhap_excel import get_data, get_excel_from_table, upload_excel_to_db
from datetime import datetime

scp = Blueprint('scp_canhan', __name__)

SIZE = 20
@scp.route("/scp_canhan", methods=["GET"])
def scp_canhan():
    try:
        if "IED" not in current_user.phongban:
            return redirect("/")

        nhamay = request.args.get("nhamay")
        tungay = request.args.get("tungay")
        denngay = request.args.get("denngay")
        mst = request.args.get("mst")
        page = request.args.get(get_page_parameter(), type=int, default=1)
        filters = {
            "nha_may": {
                "type": "equal",
                "value": nhamay
            },
            "tu_ngay": {
                "type": "gte",
                "value": tungay
            },
            "den_ngay": {
                "type": "lte",
                "value": denngay
            },
            "mst": {
                "type": "equal",
                "value": mst
            }
        }
        data, total = get_data(filters, page, SIZE, "[INCENTIVE].[dbo].[SCP_CA_NHAN]", "TU_NGAY DESC").values()
        for row in data:
            row_list = list(row)
            row_list[3] = datetime.strftime(datetime.strptime(row_list[3], "%Y-%m-%d"), "%d/%m/%Y")
            row_list[4] = datetime.strftime(datetime.strptime(row_list[4], "%Y-%m-%d"), "%d/%m/%Y")
            data[data.index(row)] = tuple(row_list)
        pagination = Pagination(page=page, per_page=SIZE, total=total, css_framework='bootstrap4')
        return render_template("scp_canhan.html", danhsach=data, pagination=pagination)
    except Exception as e:
        print(e)
        return render_template("scp_canhan.html")
        
@scp.route("/scp_canhan/excel", methods=["GET"])
def get_excel():
    return get_excel_from_table("INCENTIVE", "SCP_CA_NHAN", "scp_canhan")
    
@scp.route("/scp_canhan/upload_excel", methods=["POST"])
def upload_excel():
    try:
        if 'file' not in request.files:
            return None
        
        file = request.files["file"]
        upload_excel_to_db("INCENTIVE", "SCP_CA_NHAN", file)
        return redirect("/scp_canhan")
    except Exception as e:
        print(e)
        return None

@scp.route("/scp_canhan/filter", methods=["POST"])
def filter():
    try:
        mst = request.form.get("mst")
        nhamay = request.form.get("nhamay")
        tungay = request.form.get("tungay")
        denngay = request.form.get("denngay")
        return redirect(f"/scp_canhan?mst={mst}&nhamay={nhamay}&tungay={tungay}&denngay={denngay}")
    except Exception as e:
        print(e)
        return None