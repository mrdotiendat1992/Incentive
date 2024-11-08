from flask import render_template, request, Blueprint, redirect
from flask_login import current_user
from flask_paginate import Pagination, get_page_parameter
import pandas as pd
from helper.utils import *
from helper.nhap_excel import get_data, get_excel_from_table, upload_excel_to_db
from datetime import datetime

samsew = Blueprint('sam_sew', __name__)

SIZE = 20
@samsew.route("/sam_sew", methods=["GET"])
def sam_sew():
    try:
        if "IED" not in current_user.phongban:
            return redirect("/")

        style = request.args.get("style")
        phienban = request.args.get("phienban")
        page = request.args.get(get_page_parameter(), type=int, default=1)
        filters = {
            "style": {
                "type": "approximately",
                "value": style
            },
            "phien_ban": {
                "type": "equal",
                "value": phienban
            }
        }
        data, total = get_data(filters, page, SIZE, "[INCENTIVE].[dbo].[SAM_SEW]", "STYLE").values()
        for row in data:
            row_list = list(row)
            row_list[6] = datetime.strftime(datetime.strptime(row_list[6], "%Y-%m-%d"), "%d/%m/%Y")
            row_list[7] = datetime.strftime(datetime.strptime(row_list[7], "%Y-%m-%d"), "%d/%m/%Y")
            data[data.index(row)] = tuple(row_list)
        pagination = Pagination(page=page, per_page=SIZE, total=total, css_framework='bootstrap4')
        return render_template("sam_sew.html", danhsach=data, pagination=pagination)
    except Exception as e:
        print(e)
        return render_template("sam_sew.html")
        
@samsew.route("/sam_sew/excel", methods=["GET"])
def get_excel():
    filters = {
        "style": {
            "type": "approximately",
            "value": request.args.get("style")
        },
        "phien_ban": {
            "type": "equal",
            "value": request.args.get("phienban")
        }
    }
    return get_excel_from_table("INCENTIVE", "SAM_SEW", "sam_sew", filters)
    
@samsew.route("/sam_sew/upload_excel", methods=["POST"])
def upload_excel():
    try:
        if 'file' not in request.files:
            return None
        
        file = request.files["file"]
        upload_excel_to_db("INCENTIVE", "SAM_SEW", file)
        return redirect("/sam_sew")
    except Exception as e:
        print(e)
        return None

@samsew.route("/sam_sew/filter", methods=["POST"])
def filter():
    try:
        style = request.form.get("style")
        phienban = request.form.get("phienban")
        return redirect(f"/sam_sew?style={style}&phienban={phienban}")
    except Exception as e:
        print(e)
        return None