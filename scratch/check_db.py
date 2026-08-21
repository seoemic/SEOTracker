import sys
import os
sys.path.append(os.getcwd())
import mysql.connector
from auditor import DB_CONFIG

def check():
    try:
        cnx = mysql.connector.connect(**DB_CONFIG)
        cur = cnx.cursor()
        
        print("--- TABLE: marcas ---")
        cur.execute("DESCRIBE marcas")
        for row in cur.fetchall():
            print(row)
            
        print("\n--- TABLE: historico ---")
        cur.execute("DESCRIBE historico")
        for row in cur.fetchall():
            print(row)
            
        cur.close()
        cnx.close()
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    check()
