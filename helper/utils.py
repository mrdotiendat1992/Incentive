import pyodbc
import os
from dotenv import load_dotenv

load_dotenv("../.env")

database_server = os.environ.get("DATABASE_SERVER")
database_name = os.environ.get("DATABASE_NAME")
database_user = os.environ.get("DATABASE_USER")
database_password = os.environ.get("DATABASE_PASSWORD")

def connect_db():
    conn = pyodbc.connect(f'DRIVER={{SQL Server}};SERVER={database_server};DATABASE={database_name};UID={database_user};PWD={database_password}')
    return conn

def close_db(conn):
    conn.close()
    
def execute_query(conn, query):
    cursor = conn.cursor()
    cursor.execute(query)
    return cursor

def execute_query_data(conn, query, data):
    cursor = conn.cursor()
    cursor.executemany(query, data)
    return cursor