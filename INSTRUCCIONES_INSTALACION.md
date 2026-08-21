# Guía de Instalación — SEO Auditor

Para que el proyecto funcione correctamente en un nuevo equipo, sigue estos pasos:

## 1. Requisitos de Software
Asegúrate de tener instalado:
*   **Python 3.10 o superior.**
*   **MySQL Server** (ej: usando XAMPP, WAMP o una instalación independiente).

## 2. Instalación de Dependencias
Abre una terminal en la carpeta del proyecto y ejecuta:
```bash
pip install -r requirements.txt
```

## 3. Configuración de Base de Datos
1.  Asegúrate de que tu servidor MySQL esté encendido.
2.  Ejecuta el script de configuración inicial:
    ```bash
    python setup_db.py
    ```
    *El script te pedirá el usuario y contraseña de tu MySQL (por defecto suele ser root y sin contraseña).*

## 4. Archivos de Configuración Críticos
Para que el sistema pueda rastrear y enviar correos, el proyecto **debe incluir** estos archivos (pídelos al desarrollador si no los tienes):
*   `credentials.json`: Las credenciales de la Service Account de Google Cloud para acceder a Google Sheets.
*   **Variables de Entorno (Opcional):** El sistema usa por defecto `smtp.gmail.com`. Si se desea cambiar el emisor del correo, debe editarse en `app.py` o configurarse como variables de entorno (`SMTP_EMAIL`, `SMTP_PASS`).

## 5. Ejecución
Para iniciar el servidor web:
```bash
python app.py
```
Luego, abre tu navegador en: `http://localhost:5000`

---
**Usuario Administrador por defecto:**
*   **Email:** admin@seoauditor.com
*   **Password:** admin123
