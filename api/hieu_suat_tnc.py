from helper.utils import connect_db, execute_query, close_db
from flask import render_template, request, Blueprint, redirect
from flask_login import current_user
from flask_paginate import Pagination, get_page_parameter
from datetime import datetime
from helper.utils import *
from helper.nhap_excel import get_data, get_excel_from_table, upload_excel_to_db, getDataTNC

tnc = Blueprint('hieusuat_tnc', __name__)

SIZE = 20

@tnc.route("/hieusuat_tnc", methods=["GET"])
def danhsach_totruong():
    try:    
        if "IED" not in current_user.phongban:
            return redirect("/")
        
        ngay = request.args.get("ngay")
        chuyen = request.args.get("chuyen")
        mst = request.args.get("mst")
        page = request.args.get(get_page_parameter(), type=int, default=1)
        filters = {
            "nha_may": {
                "type": "equal",
                "value": current_user.macongty
            },
            "ngay": {
                "type": "equal",
                "value": ngay
            },
            "chuyen": {
                "type": "approximately",
                "value": chuyen
            },
            "mst": {
                "type": "equal",
                "value": mst
            }
        }
        
        data, total = get_data(filters, page, SIZE, "[INCENTIVE].[dbo].[HIEU_SUAT_CN_TNC]", "NGAY DESC").values()
        for row in data:
            row_list = list(row)
            for i in range(len(row_list)):
                if row_list[i] is None:
                    row_list[i] = ""
            row_list[4] = datetime.strftime(datetime.strptime(row_list[4], "%Y-%m-%d"), "%d/%m/%Y")
            data[data.index(row)] = tuple(row_list)

        lines, tnc = getDataTNC().values()
        pagination = Pagination(page=page, per_page=SIZE, total=total, css_framework='bootstrap4')
        return render_template("hieusuat_tnc.html", danhsach=data, lines=lines, tnc=tnc, pagination=pagination)
    except Exception as e:
        print(e)
        return render_template("hieusuat_tnc.html")
        
@tnc.route("/hieusuat_tnc/excel", methods=["GET"])
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
        "chuyen": {
            "type": "approximately",
            "value": request.args.get("chuyen")
        },
        "mst": {
            "type": "equal",
            "value": request.args.get("mst")
        }
    }
    return get_excel_from_table("INCENTIVE", "HIEU_SUAT_CN_TNC", "hieusuat_tnc", filters, ["ngay"])
    
@tnc.route("/hieusuat_tnc/upload_excel", methods=["POST"])
def upload_excel():
    try:
        if 'file' not in request.files:
            return None
        
        file = request.files["file"]
        upload_excel_to_db("INCENTIVE", "HIEU_SUAT_CN_TNC", file)
        return redirect("/hieusuat_tnc")
    except Exception as e:
        print(e)
        return None
    
@tnc.route("/hieusuat_tnc/filter", methods=["POST"])
def filter():
    try:
        mst = request.form.get("mst")
        ngay = request.form.get("ngay")
        chuyen = request.form.get("chuyen")
        return redirect(f"/hieusuat_tnc?mst={mst}&ngay={ngay}&chuyen={chuyen}")
    except Exception as e:
        print(e)
        return None
    
@tnc.route("/update_line_tnc", methods=["POST"])
def update_line_tnc():
    try:
        data = request.json
        mst = data["mst"]
        chuyen = data["chuyen"]
        conn = connect_db()
        queryUpdate = f"UPDATE HR.dbo.Danh_sach_CBCNV SET Ghi_chu = '{chuyen}' WHERE The_cham_cong = {mst} AND Factory = '{current_user.macongty}'"
        execute_query(conn, queryUpdate)
        conn.commit()
        close_db(conn)
        return {"message": "Cập nhật thành công"}
    except Exception as e:
        print(e)
        return {"message": "Cập nhật thất bại"}, 400
