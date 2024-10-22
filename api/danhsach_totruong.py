from flask import render_template, request, Blueprint, make_response, redirect
from flask_paginate import Pagination, get_page_parameter
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment
from utils import *
from io import BytesIO
from datetime import datetime

totruong = Blueprint('danhsach_totruong', __name__)

SIZE = 20

def getListTotruong(filters, page):
    conn = connect_db()
    offset = (page - 1) * SIZE + 1
    query = f"SELECT *, ROW_NUMBER() OVER (ORDER BY NHA_MAY) AS RowNum FROM [INCENTIVE].[dbo].[DS_TO_TRUONG]"

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

    count_query = f"SELECT COUNT(*) FROM [INCENTIVE].[dbo].[DS_TO_TRUONG]"
    if conditions:
        count_query += " WHERE " + " AND ".join(conditions)
    cursor2 = execute_query(conn, count_query)
    total = cursor2.fetchall()[0][0]
    close_db(conn)
    return {
        "data": rows,
        "total": total
    }

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
        print(page)
        data, total = getListTotruong(filters, page).values()
        pagination = Pagination(page=page, per_page=SIZE, total=total, css_framework='bootstrap4')
        return render_template("danhsach_totruong.html", danhsach=data, pagination=pagination)
    except Exception as e:
        print(e)
        return render_template("danhsach_totruong.html")
        
@totruong.route("/dannhsach_totruong/excel", methods=["GET"])
def get_excel():
    try:
        conn = connect_db()
        query_header = """SELECT COLUMN_NAME
                FROM [INCENTIVE].INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_NAME = 'DS_TO_TRUONG';"""
        cursor = execute_query(conn, query_header)
        rows = cursor.fetchall()
        headers = [row[0] for row in rows]

        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        
        ws.append(headers)

        cursor = execute_query(conn, "SELECT * FROM [INCENTIVE].[dbo].[DS_TO_TRUONG]")
        data = cursor.fetchall()
        print(data)
        for row in data:
            ws.append(list(row))

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