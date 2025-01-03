from flask import Flask
from flask_sqlalchemy import SQLAlchemy
import urllib.parse
import logging
from logging.handlers import RotatingFileHandler
from api import *

params = urllib.parse.quote_plus(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=172.16.60.100;"
    "DATABASE=HR;"
    
    "UID=huynguyen;"
    "PWD=Namthuan@123;"
)

app = Flask("incentive_system")
app.config["SQLALCHEMY_DATABASE_URI"] = f"mssql+pyodbc:///?odbc_connect={params}"
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config["SECRET_KEY"] = "incentive_system"

app.register_blueprint(totruong)
app.register_blueprint(tnc)
app.register_blueprint(scp)
app.register_blueprint(samsew)
app.register_blueprint(donhang)
app.register_blueprint(quanly)
app.register_blueprint(loi)
app.register_blueprint(poly)
app.register_blueprint(cat)
app.register_blueprint(la)
app.register_blueprint(qc1)
app.register_blueprint(qc2)
app.register_blueprint(ets)
app.register_blueprint(maymac)

db = SQLAlchemy(app)
handler = RotatingFileHandler('app.log', maxBytes=10000, backupCount=1, encoding='utf-8')
handler.setLevel(logging.INFO)
formatter = logging.Formatter(
    '%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d]'
)

handler.setFormatter(formatter)
app.logger.addHandler(handler)
app.logger.setLevel(logging.INFO)

