# database.py
import pyodbc
from config import DATABASE_CONFIG

def get_connection():
    conn = pyodbc.connect(
        f"DRIVER={DATABASE_CONFIG['driver']};"
        f"SERVER={DATABASE_CONFIG['server']};"
        f"DATABASE={DATABASE_CONFIG['database']};"
        f"UID={DATABASE_CONFIG['username']};"
        f"PWD={DATABASE_CONFIG['password']}"
    )
    return conn

def fetch_query(query):
    conn = get_connection()
    df = pd.read_sql(query, conn)
    conn.close()
    return df
