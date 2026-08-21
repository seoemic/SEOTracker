import smtplib
import os
import sys
sys.path.append(os.getcwd())
from auditor import SMTP_CONFIG

def test_smtp():
    print(f"Probando conexión SMTP...")
    print(f"Servidor: {SMTP_CONFIG['servidor']}:{SMTP_CONFIG['puerto']}")
    print(f"Emisor: {SMTP_CONFIG['email_emisor']}")
    
    try:
        server = smtplib.SMTP(SMTP_CONFIG['servidor'], SMTP_CONFIG['puerto'], timeout=10)
        print("1. Conexión establecida.")
        server.starttls()
        print("2. STARTTLS exitoso.")
        server.login(SMTP_CONFIG['email_emisor'], SMTP_CONFIG['app_password'])
        print("3. Login exitoso.")
        server.quit()
        print("--- TEST EXITOSO ---")
    except Exception as e:
        print(f"--- TEST FALLIDO ---")
        print(f"Error: {e}")

if __name__ == "__main__":
    test_smtp()
