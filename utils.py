import pyodbc

def connect_db():
    conn = pyodbc.connect(r'DRIVER={SQL Server};SERVER=172.16.60.100;DATABASE=HR;UID=huynguyen;PWD=Namthuan@123')
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