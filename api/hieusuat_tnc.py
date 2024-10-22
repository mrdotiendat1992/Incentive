from flask import render_template, request, Blueprint, make_response, redirect
from flask_paginate import Pagination, get_page_parameter
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment
from io import BytesIO
from datetime import datetime
import pandas as pd
from utils import *

tnc = Blueprint('hieusuat_tnc', __name__)

SIZE = 20

def getListTNC(filters, page):
    conn = connect_db()
    offset = (page - 1) * SIZE + 1
    query = f"SELECT *, ROW_NUMBER() OVER (ORDER BY NGAY DESC) AS RowNum FROM [INCENTIVE].[dbo].[HIEU_SUAT_CN_TNC]"

    conditions = []
    for key, value in filters.items():
        if value:
            if key != "chuyen":
                conditions.append(f"{key} = '{value}'")
            else:
                conditions.append(f"{key} LIKE '%{value}%'")

    if conditions:
        query += " WHERE " + " AND ".join(conditions)

    last_query = f"WITH TEMP AS ({query}) SELECT * FROM TEMP WHERE RowNum BETWEEN {offset} AND {offset + SIZE - 1}"

    cursor = execute_query(conn, last_query)
    rows = cursor.fetchall()
    for row in rows:
        row_list = list(row)
        row_list[4] = datetime.strftime(datetime.strptime(row_list[4], "%Y-%m-%d"), "%d/%m/%Y")
        rows[rows.index(row)] = tuple(row_list)

    count_query = f"SELECT COUNT(*) FROM [INCENTIVE].[dbo].[HIEU_SUAT_CN_TNC]"
    if conditions:
        count_query += " WHERE " + " AND ".join(conditions)
    cursor2 = execute_query(conn, count_query)
    total = cursor2.fetchall()[0][0]
    close_db(conn)
    return {
        "data": rows,
        "total": total
    }

@tnc.route("/hieusuat_tnc", methods=["GET"])
def danhsach_totruong():
    try:    
        ngay = request.args.get("ngay")
        chuyen = request.args.get("chuyen")
        mst = request.args.get("mst")
        page = request.args.get(get_page_parameter(), type=int, default=1)
        filters = {
            "ngay": ngay,
            "chuyen": chuyen,
            "mst": mst
        }
        print(page)
        data, total = getListTNC(filters, page).values()
        pagination = Pagination(page=page, per_page=SIZE, total=total, css_framework='bootstrap4')
        return render_template("hieusuat_tnc.html", danhsach=data, pagination=pagination)
    except Exception as e:
        print(e)
        return render_template("hieusuat_tnc.html")
        
@tnc.route("/hieusuat_tnc/excel", methods=["GET"])
def get_excel():
    try:
        conn = connect_db()
        query_header = """SELECT COLUMN_NAME
                FROM [INCENTIVE].INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_NAME = 'HIEU_SUAT_CN_TNC';"""
        cursor = execute_query(conn, query_header)
        rows = cursor.fetchall()
        headers = [row[0] for row in rows]
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        
        ws.append(headers)

        thin_border = Border(left=Side(border_style="thin"),
                    right=Side(border_style="thin"),
                    top=Side(border_style="thin"),
                    bottom=Side(border_style="thin"))
        
        for row in ws.iter_rows(min_row=1, max_row=1, max_col=len(headers)):
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')
            
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column].width = max(adjusted_width,7)

        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        
        current_time = datetime.now().strftime("%d/%m/%Y_%H_%M_%S")
        close_db(conn)
        
        response = make_response(buffer.read())
        response.headers.set('Content-Disposition', f'attachment; filename="hieusuat_tnc{current_time}.xlsx"')
        response.headers.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        return response    
    except Exception as e:
        print(e)
        return None
    
@tnc.route("/hieusuat_tnc/upload_excel", methods=["POST"])
def upload_excel():
    try:
        if 'file' not in request.files:
            return None
        
        file = request.files["file"]
        df = pd.read_excel(file)
        data_tuples = [tuple(row) for row in df.itertuples(index=False, name=None)]

        conn = connect_db()
        execute_query_data(conn, "INSERT INTO [INCENTIVE].[dbo].[HIEU_SUAT_CN_TNC] VALUES (?, ?, ?, ?, ?, ?)", data_tuples)
        conn.commit()
        close_db(conn)
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