"""
app.py — Punto de entrada de la aplicación web SEOTracker.

Aquí se configura Flask, las rutas de autenticación, el dashboard principal,
el CRUD de marcas/células/usuarios, y la lógica de reportes.
Todo el "trabajo pesado" de auditoría lo delega al módulo auditor.py.
"""
from flask import (Flask, render_template, request, redirect, url_for,
                   flash, send_file, jsonify, abort)
from flask_login import (LoginManager, UserMixin, login_user, logout_user,
                         login_required, current_user)
from werkzeug.security import generate_password_hash, check_password_hash
import mysql.connector
import threading
import datetime
import os
import string
import random
import auditor

# Guardamos el progreso de las auditorías en curso para que el frontend
# pueda consultar el estado via polling (marca_id → estado del job).
audit_status = {}

# Lo mismo pero para los rastreos manuales del "Modo Lista".
# Cada job tiene un UUID propio para que el usuario pueda rastrear su solicitud.
rastreo_jobs = {}

app = Flask(__name__)
app.secret_key = 'seo_auditor_secret_2026'  # ⚠ Cámbiala por algo seguro antes de subir a producción

# Configuramos Flask-Login para manejar sesiones de usuario.
# Si alguien intenta acceder sin estar logueado, lo mandamos al login.
login_manager = LoginManager(app)
login_manager.login_view = 'login'
login_manager.login_message = 'Por favor inicia sesión para continuar.'

# Le pasamos la configuración de BD y SMTP al módulo auditor.
# Usa variables de entorno si existen; si no, valores por defecto para desarrollo local.
auditor.DB_CONFIG = {
    'host': os.getenv('DB_HOST', '127.0.0.1'),
    'user': os.getenv('DB_USER', 'root'),
    'password': os.getenv('DB_PASS', ''),
    'database': 'seo_auditor'
}
auditor.SMTP_CONFIG = {
    'email_emisor': os.getenv('SMTP_EMAIL', 'fabianperez@estrategiamic.com'),
    'app_password':  os.getenv('SMTP_PASS',  'etpr gevb fcvt nynw'),
    'servidor': os.getenv('SMTP_SERVER', 'smtp.gmail.com'),
    'puerto': int(os.getenv('SMTP_PORT', 587))
}

# Modelo de usuario para Flask-Login. Envuelve la fila de la BD en un objeto
# con propiedades útiles como is_admin y is_lider.
class User(UserMixin):
    def __init__(self, row):
        self.id       = row['id']
        self.username = row['username']
        self.email    = row['email']
        self.rol      = row['rol']
        self.rol_empresa = row.get('rol_empresa', '')
        self.activo   = row['activo']

    def get_id(self): return str(self.id)
    @property
    def is_admin(self): return self.rol == 'admin'
    @property
    def is_lider(self): return self.rol == 'lider'

def crear_usuario_automatico(email, rol):
    if not email or not email.strip(): return
    cnx = auditor.get_conn()
    cur = cnx.cursor(dictionary=True)
    try:
        cur.execute("SELECT id FROM usuarios WHERE email = %s", (email.strip(),))
        if cur.fetchone():
            return  # Este correo ya tiene cuenta, no hacemos nada
        
        username = email.strip().split('@')[0]
        # Si ya hay alguien con ese username, le agregamos caracteres al azar para hacerlo único
        cur.execute("SELECT id FROM usuarios WHERE username = %s", (username,))
        while cur.fetchone():
            username += "".join(random.choices(string.ascii_lowercase + string.digits, k=3))
            cur.execute("SELECT id FROM usuarios WHERE username = %s", (username,))
        
        password_plana = "".join(random.choices(string.ascii_letters + string.digits + "!@#$%^&*", k=10))
        hashed = generate_password_hash(password_plana)
        
        cur.execute(
            "INSERT INTO usuarios (username, email, password_hash, rol, rol_empresa) VALUES (%s, %s, %s, %s, %s)",
            (username, email.strip(), hashed, rol, 'Interno')
        )
        cnx.commit()
        
        # Le mandamos un correo con sus credenciales para que pueda entrar
        auditor.enviar_correo_credenciales(email.strip(), rol, password_plana)
        
    except Exception as e:
        print(f"Error creando usuario automatico: {e}")
    finally:
        cur.close()
        cnx.close()

@login_manager.user_loader
def load_user(user_id):
    try:
        cnx = auditor.get_conn()
        cur = cnx.cursor(dictionary=True)
        cur.execute("SELECT * FROM usuarios WHERE id=%s AND activo=1", (user_id,))
        row = cur.fetchone()
        cur.close(); cnx.close()
        return User(row) if row else None
    except Exception:
        return None

# Decoradores de permisos: restringen el acceso a ciertas rutas
# según el rol del usuario que está logueado.
def lider_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if not current_user.is_authenticated or current_user.rol != 'lider':
            abort(403)
        return f(*args, **kwargs)
    return decorated

def pm_or_lider_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if not current_user.is_authenticated or current_user.rol not in ('lider', 'pm'):
            abort(403)
        return f(*args, **kwargs)
    return decorated

def verificar_acceso_marca(m):
    if current_user.is_lider: return
    if current_user.email not in (m.get('implementador_email'), m.get('gerente_email'), m.get('pm_email')):
        abort(403)

# =====================================================================
# AUTENTICACIÓN — Login, logout y perfil del usuario
# =====================================================================
@app.route('/', methods=['GET','POST'])
@app.route('/login', methods=['GET','POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        email = request.form.get('email','').strip()
        password = request.form.get('password','')
        try:
            cnx = auditor.get_conn()
            cur = cnx.cursor(dictionary=True)
            cur.execute("SELECT * FROM usuarios WHERE email=%s AND activo=1", (email,))
            row = cur.fetchone()
            cur.close(); cnx.close()
            if row and check_password_hash(row['password_hash'], password):
                login_user(User(row))
                return redirect(url_for('dashboard'))
            flash('Correo o contraseña incorrectos.', 'error')
        except Exception as e:
            flash(f'Error de conexión: {e}', 'error')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

@app.route('/perfil', methods=['GET', 'POST'])
@login_required
def perfil():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        if not username:
            flash('El nombre de usuario no puede estar vacío.', 'error')
            return redirect(url_for('perfil'))
        try:
            cnx = auditor.get_conn()
            cur = cnx.cursor()
            if password:
                hashed = generate_password_hash(password)
                cur.execute("UPDATE usuarios SET username=%s, password_hash=%s WHERE id=%s", (username, hashed, current_user.id))
            else:
                cur.execute("UPDATE usuarios SET username=%s WHERE id=%s", (username, current_user.id))
            cnx.commit()
            cur.close()
            cnx.close()
            flash('Perfil actualizado exitosamente.', 'success')
            
            # No necesitamos actualizar el objeto en memoria manualmente:
            # Flask-Login lo recarga desde la BD en la siguiente petición.
            
            return redirect(url_for('perfil'))
        except Exception as e:
            flash(f'Error al actualizar el perfil: {e}', 'error')
    return render_template('perfil.html')


# =====================================================================
# DASHBOARD — Vista principal con resumen de todas las marcas
# =====================================================================
@app.route('/dashboard')
@login_required
def dashboard():
    try:
        email_filtro = None if current_user.is_lider else current_user.email
        marcas = auditor.obtener_marcas(email_filtro=email_filtro)
        data = []
        ahora = datetime.datetime.now()
        
        for m in marcas:
            # --- Check de expiración de headers ---
            if m.get('headers_personalizados') and m.get('fecha_headers'):
                # Usar isinstance para evitar errores si fecha_headers es date
                fecha_h = m['fecha_headers']
                if not isinstance(fecha_h, datetime.datetime):
                    fecha_h = datetime.datetime.combine(fecha_h, datetime.time())
                dias_headers = (ahora - fecha_h).days
                if dias_headers >= 90 and not m.get('alerta_headers_enviada'):
                    # Disparar alerta en background para no bloquear
                    threading.Thread(target=auditor.enviar_alerta_headers, args=(m,), daemon=True).start()
                    # Actualizar bd a enviada
                    try:
                        cx = auditor.get_conn()
                        cr = cx.cursor()
                        cr.execute("UPDATE marcas SET alerta_headers_enviada=1 WHERE id=%s", (m['id'],))
                        cx.commit()
                        cr.close()
                        cx.close()
                        m['alerta_headers_enviada'] = 1
                    except Exception as ex:
                        print(f"Error al marcar alerta enviada: {ex}")
                m['headers_expirados'] = dias_headers >= 90
            else:
                m['headers_expirados'] = False
            # --- Fin check ---

            ult = auditor.obtener_ultimo_historico(m['id'])
            proxima = None
            if ult and isinstance(ult.get('fecha_ejecucion'), (datetime.datetime, datetime.date)):
                proxima = ult['fecha_ejecucion'] + datetime.timedelta(days=15)
            data.append({**m, 'ultimo': ult, 'proxima': proxima})
        return render_template('dashboard.html', marcas=data)
    except Exception as e:
        flash(f"Error cargando el dashboard: {e}", "error")
        return render_template('dashboard.html', marcas=[])

@app.route('/auditar/<int:marca_id>', methods=['POST'])
@login_required
def forzar_auditoria(marca_id):
    if current_user.rol not in ('lider', 'implementador'):
        return jsonify({'error': 'No tienes permisos para auditar en vivo.'}), 403
    m = auditor.obtener_marca(marca_id)
    if not m:
        return jsonify({'error': 'Marca no encontrada'}), 404
    try:
        verificar_acceso_marca(m)
    except:
        return jsonify({'error': 'No tienes acceso a esta marca.'}), 403

    # Si ya hay una auditoría en marcha para esta marca, no lanzamos otra
    if audit_status.get(marca_id, {}).get('status') == 'running':
        return jsonify({'ok': True, 'message': f'La auditoría de {m["marca"]} ya está en progreso.'})

    audit_status[marca_id] = {'status': 'running', 'message': 'Auditoría iniciada...'}

    def _run_audit(marca_db, mid):
        try:
            auditor.auditar_marca(marca_db)
            audit_status[mid] = {'status': 'done', 'message': f'Auditoría de {marca_db["marca"]} completada.'}
        except Exception as e:
            audit_status[mid] = {'status': 'error', 'message': f'Error: {e}'}

    threading.Thread(target=_run_audit, args=(m, marca_id), daemon=True).start()
    return jsonify({'ok': True, 'message': f'Auditoría de {m["marca"]} iniciada en segundo plano.'})

@app.route('/audit_status/<int:marca_id>')
@login_required
def get_audit_status(marca_id):
    """El frontend consulta este endpoint cada pocos segundos para saber cómo va la auditoría."""
    status = audit_status.get(marca_id, {'status': 'idle', 'message': 'Sin auditoría activa.'})
    return jsonify(status)

@app.route('/historico/<int:marca_id>')
@login_required
def historico(marca_id):
    m = auditor.obtener_marca(marca_id)
    if not m: abort(404)
    verificar_acceso_marca(m)
    hist_raw = auditor.obtener_historico_marca(marca_id, limit=40)
    
    if not m.get('responsive', True):
        agrupado = {}
        for h in hist_raw:
            rev = h['numero_revision']
            if rev not in agrupado:
                agrupado[rev] = {
                    'numero_revision': rev,
                    'fecha_ejecucion': h['fecha_ejecucion'],
                    'mobile': None,
                    'desktop': None,
                    'es_agrupado': True
                }
            ruta = h.get('ruta_reporte_excel', '').lower()
            if 'mobile' in ruta:
                agrupado[rev]['mobile'] = h
            elif 'desktop' in ruta:
                agrupado[rev]['desktop'] = h
            else:
                agrupado[rev]['desktop'] = h
        hist = list(agrupado.values())
        hist.sort(key=lambda x: x['numero_revision'], reverse=True)
    else:
        hist = hist_raw

    return render_template('historico.html', marca=m, historico=hist)

@app.route('/test_email/<int:marca_id>', methods=['POST'])
@login_required
@lider_required
def test_email(marca_id):
    m = auditor.obtener_marca(marca_id)
    if not m: return jsonify({'error': 'Marca no encontrada'}), 404
    
    # Enviamos un correo de prueba (sin reporte adjunto) para verificar que SMTP funciona
    success = auditor.enviar_correo(
        marca_db = m,
        path_excel = "", # Sin excel para el test
        stats = {'total':0,'redireccion':0,'no_creadas':0,'meta_titulo':0,'meta_descripcion':0,'h1':0,'desc_corta':0,'desc_larga':0},
        historico_anterior = None,
        num_revision = 0,
        fecha_str = "TEST DE CONEXIÓN"
    )
    if success:
        return jsonify({'ok': True, 'message': f'Correo de prueba enviado a la Célula: {m["implementador_email"]}, {m["gerente_email"]}'})
    else:
        return jsonify({'ok': False, 'error': 'Error de autenticación SMTP. Revisa el correo y la contraseña de aplicación en el servidor.'})

@app.route('/diagnostico/<int:marca_id>')
@login_required
@lider_required
def diagnostico(marca_id):
    """Herramienta de diagnóstico: prueba la conexión con Google Sheets y muestra
    un resumen rápido (columnas, total de URLs, primeras 5 URLs) para verificar
    que el documento está bien configurado."""
    m = auditor.obtener_marca(marca_id)
    if not m:
        return jsonify({'error': 'Marca no encontrada'}), 404

    import os
    resultado = {
        'marca': m['marca'],
        'doc_contenido': m['doc_contenido'],
        'credentials_encontradas': os.path.exists(auditor.PATH_CREDENTIALS_SHEETS),
        'status': 'error',
        'detalle': '',
        'columnas': [],
        'total_urls': 0,
        'primeras_urls': []
    }

    try:
        df = auditor.descargar_hoja_sheets(m['doc_contenido'])
        if df.empty:
            resultado['detalle'] = 'El DataFrame está vacío. Verifica credenciales y permisos del Sheets.'
        else:
            resultado['status'] = 'ok'
            resultado['columnas'] = df.columns.tolist()
            resultado['total_urls'] = int(df['URL'].dropna().nunique()) if 'URL' in df.columns else 0
            resultado['primeras_urls'] = df['URL'].dropna().head(5).tolist() if 'URL' in df.columns else []
            resultado['detalle'] = f'{len(df)} filas encontradas en el documento.'
    except Exception as e:
        resultado['detalle'] = str(e)

    # Renderizamos el resultado como una página HTML sencilla (no usa template)
    html = f"""
    <!DOCTYPE html><html><head><title>Diagnóstico — {resultado['marca']}</title>
    <link rel="stylesheet" href="/static/css/style.css">
    </head><body style="padding:30px;background:#0a0a14;color:#e0e0e0;font-family:sans-serif;">
    <a href="/dashboard" style="color:#4fc3f7;text-decoration:none;">← Dashboard</a>
    <h2 style="color:#4fc3f7;margin:20px 0 10px;">Diagnóstico: {resultado['marca']}</h2>
    <p><strong>URL del documento:</strong> <a href="{resultado['doc_contenido']}" target="_blank" style="color:#4fc3f7;">{resultado['doc_contenido'][:80]}...</a></p>
    <p><strong>credentials.json:</strong> {'✅ Encontrado' if resultado['credentials_encontradas'] else '❌ NO encontrado — asegúrate de colocar el archivo en la carpeta del proyecto'}</p>
    <hr style="border-color:#252540;margin:20px 0;">
    <p><strong>Estado:</strong> {'✅ Conexión exitosa' if resultado['status']=='ok' else '❌ Error de conexión'}</p>
    <p><strong>Detalle:</strong> {resultado['detalle']}</p>
    {'<p><strong>Total URLs únicas:</strong> ' + str(resultado['total_urls']) + '</p>' if resultado['status']=='ok' else ''}
    {'<p><strong>Columnas encontradas:</strong> ' + ', '.join(resultado['columnas']) + '</p>' if resultado['columnas'] else ''}
    {'<p><strong>Primeras 5 URLs:</strong></p><ul>' + ''.join(f'<li style="font-size:13px;color:#aaa;">{u}</li>' for u in resultado['primeras_urls']) + '</ul>' if resultado['primeras_urls'] else ''}
    </body></html>
    """
    return html

@app.route('/descargar/<path:filename>')
@login_required
def descargar(filename):
    path = os.path.abspath(filename)
    if not os.path.exists(path):
        abort(404)
    return send_file(path, as_attachment=True)


# RASTREO POR LISTA — El usuario pega URLs manualmente y las auditamos

@app.route('/rastreo_lista', methods=['GET', 'POST'])
@login_required
def rastreo_lista():
    if request.method == 'POST':
        data = request.get_json()
        marca_id = data.get('marca_id')
        urls_str = data.get('urls', '')

        m = auditor.obtener_marca(marca_id)
        if not m:
            return jsonify({'ok': False, 'error': 'Marca no encontrada'})

        urls_lista = [u.strip() for u in urls_str.split('\n') if u.strip()]
        if not urls_lista:
            return jsonify({'ok': False, 'error': 'No se proporcionaron URLs válidas'})

        # Creamos un ID único para este job y lo lanzamos en un hilo aparte
        # para no bloquear la respuesta al usuario
        import uuid
        job_id = str(uuid.uuid4())
        rastreo_jobs[job_id] = {'status': 'running', 'result': None, 'error': None}

        def _run_rastreo(marca_db, urls, jid):
            try:
                res = auditor.auditar_lista_manual(marca_db, urls)
                if res.get('ok'):
                    rastreo_jobs[jid] = {'status': 'done', 'result': res, 'error': None}
                else:
                    rastreo_jobs[jid] = {'status': 'error', 'result': None, 'error': res.get('error', 'Error desconocido')}
            except Exception as e:
                rastreo_jobs[jid] = {'status': 'error', 'result': None, 'error': str(e)}

        threading.Thread(target=_run_rastreo, args=(m, urls_lista, job_id), daemon=True).start()
        return jsonify({'ok': True, 'job_id': job_id})

    # Si es GET, mostramos el formulario con la lista de marcas disponibles
    marcas = auditor.obtener_marcas()
    return render_template('rastreo_lista.html', marcas=marcas)


@app.route('/rastreo_lista/estado/<job_id>')
@login_required
def rastreo_lista_estado(job_id):
    """El frontend consulta este endpoint para saber si el rastreo manual ya terminó."""
    job = rastreo_jobs.get(job_id)
    if not job:
        return jsonify({'status': 'not_found', 'error': 'Job no encontrado o expirado'}), 404
    return jsonify(job)

@app.route('/rastreo_lista/buscar_por_dominio')
@login_required
def buscar_marca_por_dominio():
    dominio = request.args.get('dominio', '').strip().lower()
    if not dominio:
        return jsonify({'ok': False})
    try:
        cnx = auditor.get_conn()
        cur = cnx.cursor(dictionary=True)
        cur.execute("""
            SELECT m.id, m.marca, m.pais FROM marcas m
            WHERE LOWER(m.link_marca) LIKE %s
            LIMIT 1
        """, (f'%{dominio}%',))
        row = cur.fetchone()
        cur.close(); cnx.close()
        if row:
            return jsonify({'ok': True, 'id': row['id'], 'nombre': f"{row['marca']} ({row['pais']})"})
        return jsonify({'ok': False})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})

@app.route('/rastreo_lista/correos_disponibles/<int:marca_id>')
@login_required
def correos_disponibles(marca_id):
    correos = auditor.obtener_correos_disponibles(marca_id)
    return jsonify({'ok': True, 'correos': correos})

@app.route('/rastreo_lista/enviar_correo', methods=['POST'])
@login_required
def rastreo_lista_enviar_correo():
    data = request.get_json()
    marca_id = data.get('marca_id')
    excel_path = data.get('excel_path')
    stats = data.get('stats')
    destinatarios = data.get('destinatarios', [])  # lista de emails seleccionados

    excel_path_desktop = data.get('excel_path_desktop')
    stats_desktop = data.get('stats_desktop')

    m = auditor.obtener_marca(marca_id)
    if not m:
        return jsonify({'ok': False, 'error': 'Marca no encontrada'})
    if not destinatarios:
        return jsonify({'ok': False, 'error': 'Selecciona al menos un destinatario.'})
    
    success = auditor.enviar_correo_lista(
        m, excel_path, stats, destinatarios,
        path_excel_desktop=excel_path_desktop,
        stats_desktop=stats_desktop
    )
    if success:
        return jsonify({'ok': True, 'message': f'Correo enviado a {len(destinatarios)} destinatario(s).'})
    else:
        return jsonify({'ok': False, 'error': 'Hubo un problema al enviar el correo.'})


# MARCAS — Alta, edición, eliminación y listado de marcas

@app.route('/marcas')
@login_required
def marcas():
    email_filtro = None if current_user.is_lider else current_user.email
    lista = auditor.obtener_marcas(email_filtro=email_filtro)
    return render_template('marcas.html', marcas=lista)

@app.route('/marcas/nueva', methods=['GET','POST'])
@login_required
@pm_or_lider_required
def marca_nueva():
    celulas = auditor.obtener_celulas()
    if request.method == 'POST':
        f = request.form
        try:
            cnx = auditor.get_conn()
            cur = cnx.cursor()
            cur.execute(
                "INSERT INTO marcas (celula_id,marca,pais,entorno,responsive,doc_contenido,"
                "link_marca,clases_h1,clases_desc_corta,clases_desc_larga,palabra_excluir,headers_personalizados) "
                "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                (f['celula_id'], f['marca'], f['pais'], f['entorno'],
                 1 if f.get('responsive') else 0,
                 f['doc_contenido'], f['link_marca'],
                 f['clases_h1'], f['clases_desc_corta'], f['clases_desc_larga'],
                 f.get('palabra_excluir', ''), f.get('headers_personalizados', ''))
            )
            cnx.commit(); cur.close(); cnx.close()
            flash('Marca creada correctamente.', 'success')
            return redirect(url_for('marcas'))
        except Exception as e:
            flash(f'Error: {e}', 'error')
    return render_template('marca_form.html', marca=None, celulas=celulas, accion='Nueva')

@app.route('/marcas/editar/<int:marca_id>', methods=['GET','POST'])
@login_required
def marca_editar(marca_id):
    m = auditor.obtener_marca(marca_id)
    if not m: abort(404)
    verificar_acceso_marca(m)
    celulas = auditor.obtener_celulas()
    if request.method == 'POST':
        f = request.form
        try:
            nuevos_headers = f.get('headers_personalizados', '')
            viejos_headers = m.get('headers_personalizados', '') or ''
            
            cnx = auditor.get_conn()
            cur = cnx.cursor()
            
            if nuevos_headers != viejos_headers and nuevos_headers.strip():
                # Actualizar también fecha_headers y resetear alerta
                cur.execute(
                    "UPDATE marcas SET celula_id=%s,marca=%s,pais=%s,entorno=%s,responsive=%s,"
                    "doc_contenido=%s,link_marca=%s,clases_h1=%s,clases_desc_corta=%s,"
                    "clases_desc_larga=%s,palabra_excluir=%s,headers_personalizados=%s,"
                    "fecha_headers=CURRENT_TIMESTAMP, alerta_headers_enviada=0 WHERE id=%s",
                    (f['celula_id'], f['marca'], f['pais'], f['entorno'],
                     1 if f.get('responsive') else 0,
                     f['doc_contenido'], f['link_marca'],
                     f['clases_h1'], f['clases_desc_corta'], f['clases_desc_larga'],
                     f.get('palabra_excluir', ''), nuevos_headers, marca_id)
                )
            else:
                cur.execute(
                    "UPDATE marcas SET celula_id=%s,marca=%s,pais=%s,entorno=%s,responsive=%s,"
                    "doc_contenido=%s,link_marca=%s,clases_h1=%s,clases_desc_corta=%s,"
                    "clases_desc_larga=%s,palabra_excluir=%s,headers_personalizados=%s WHERE id=%s",
                    (f['celula_id'], f['marca'], f['pais'], f['entorno'],
                     1 if f.get('responsive') else 0,
                     f['doc_contenido'], f['link_marca'],
                     f['clases_h1'], f['clases_desc_corta'], f['clases_desc_larga'],
                     f.get('palabra_excluir', ''), nuevos_headers, marca_id)
                )
            cnx.commit(); cur.close(); cnx.close()
            flash('Marca actualizada correctamente.', 'success')
            return redirect(url_for('marcas'))
        except Exception as e:
            flash(f'Error: {e}', 'error')
    return render_template('marca_form.html', marca=m, celulas=celulas, accion='Editar')

@app.route('/auditar_manual/<int:marca_id>', methods=['GET','POST'])
@login_required
def auditar_manual(marca_id):
    if current_user.rol not in ('lider', 'implementador', 'pm', 'gestor_cuenta'):
        abort(403)
    m = auditor.obtener_marca(marca_id)
    if not m: abort(404)
    verificar_acceso_marca(m)
    
    if request.method == 'POST':
        urls_str = request.form.get('urls', '')
        urls = [u.strip() for u in urls_str.replace(',', '\n').split('\n') if u.strip()]
        
        if not urls:
            flash("Ingresa al menos una URL.", "warning")
        else:
            try:
                res = auditor.auditar_lista_manual(m, urls)
                if res:
                    flash(f"Auditoría manual de {len(urls)} URLs completada.", "success")
                else:
                    flash("No se pudo completar la auditoría manual.", "error")
            except Exception as e:
                flash(f"Error: {e}", "error")
            return redirect(url_for('auditar_manual', marca_id=marca_id))

    return render_template('rastreo_lista.html', marca=m)


# CÉLULAS — Cada célula agrupa un implementador, gerente y PM

@app.route('/celulas')
@login_required
def celulas():
    lista = auditor.obtener_celulas()
    return render_template('celulas.html', celulas=lista)

@app.route('/celulas/nueva', methods=['GET','POST'])
@login_required
@pm_or_lider_required
def celula_nueva():
    if request.method == 'POST':
        f = request.form
        try:
            cnx = auditor.get_conn()
            cur = cnx.cursor()
            cur.execute(
                "INSERT INTO celulas (nombre, implementador_email, gerente_email, pm_email) "
                "VALUES (%s,%s,%s,%s)",
                (f['nombre'], f['implementador_email'], f['gerente_email'], f['pm_email'])
            )
            cnx.commit(); cur.close(); cnx.close()
            
            # Si los correos asignados no tienen cuenta en el sistema, se las creamos automáticamente
            crear_usuario_automatico(f['implementador_email'], 'implementador')
            crear_usuario_automatico(f['gerente_email'], 'gestor_cuenta')
            crear_usuario_automatico(f['pm_email'], 'pm')
            
            flash('Célula creada correctamente.', 'success')
            return redirect(url_for('celulas'))
        except Exception as e:
            flash(f'Error: {e}', 'error')
    
    implementadores = auditor.obtener_usuarios_por_rol('implementador')
    gestores = auditor.obtener_usuarios_por_rol('gestor_cuenta')
    pms = auditor.obtener_usuarios_por_rol('pm')
    
    return render_template('celula_form.html', celula=None, accion='Nueva', 
                           implementadores=implementadores, gestores=gestores, pms=pms)

@app.route('/celulas/editar/<int:celula_id>', methods=['GET','POST'])
@login_required
@pm_or_lider_required
def celula_editar(celula_id):
    c = auditor.obtener_celula(celula_id)
    if not c: abort(404)
    if request.method == 'POST':
        f = request.form
        try:
            cnx = auditor.get_conn()
            cur = cnx.cursor()
            cur.execute(
                "UPDATE celulas SET nombre=%s, implementador_email=%s, gerente_email=%s, pm_email=%s WHERE id=%s",
                (f['nombre'], f['implementador_email'], f['gerente_email'], f['pm_email'], celula_id)
            )
            cnx.commit(); cur.close(); cnx.close()

            # Igual que al crear: nos aseguramos de que cada correo tenga su usuario
            crear_usuario_automatico(f['implementador_email'], 'implementador')
            crear_usuario_automatico(f['gerente_email'], 'gestor_cuenta')
            crear_usuario_automatico(f['pm_email'], 'pm')

            flash('Célula actualizada correctamente.', 'success')
            return redirect(url_for('celulas'))
        except Exception as e:
            flash(f'Error: {e}', 'error')
    
    implementadores = auditor.obtener_usuarios_por_rol('implementador')
    gestores = auditor.obtener_usuarios_por_rol('gestor_cuenta')
    pms = auditor.obtener_usuarios_por_rol('pm')
    
    return render_template('celula_form.html', celula=c, accion='Editar',
                           implementadores=implementadores, gestores=gestores, pms=pms)

@app.route('/celulas/eliminar/<int:celula_id>', methods=['POST'])
@login_required
@lider_required
def celula_eliminar(celula_id):
    try:
        cnx = auditor.get_conn(); cur = cnx.cursor()
        cur.execute("DELETE FROM celulas WHERE id=%s", (celula_id,))
        cnx.commit(); cur.close(); cnx.close()
        flash('Célula eliminada.', 'success')
    except Exception as e:
        flash(f'Error: {e}', 'error')
    return redirect(url_for('celulas'))

@app.route('/marcas/eliminar/<int:marca_id>', methods=['POST'])
@login_required
@lider_required
def marca_eliminar(marca_id):
    try:
        cnx = auditor.get_conn()
        cur = cnx.cursor()
        cur.execute("DELETE FROM marcas WHERE id=%s", (marca_id,))
        cnx.commit(); cur.close(); cnx.close()
        flash('Marca eliminada.', 'success')
    except Exception as e:
        flash(f'Error: {e}', 'error')
    return redirect(url_for('marcas'))


# USUARIOS — Solo los líderes pueden gestionar cuentas de usuario

@app.route('/usuarios')
@login_required
@lider_required
def usuarios():
    cnx = auditor.get_conn()
    cur = cnx.cursor(dictionary=True)
    cur.execute("SELECT id,username,email,rol,rol_empresa,activo,creado_en FROM usuarios ORDER BY id")
    lista = cur.fetchall()
    cur.close(); cnx.close()
    return render_template('usuarios.html', usuarios=lista)

@app.route('/usuarios/nuevo', methods=['GET','POST'])
@login_required
@lider_required
def usuario_nuevo():
    if request.method == 'POST':
        f = request.form
        try:
            password_plana = "".join(random.choices(string.ascii_letters + string.digits + "!@#$%^&*", k=10))
            
            cnx = auditor.get_conn()
            cur = cnx.cursor()
            cur.execute(
                "INSERT INTO usuarios (username,email,password_hash,rol,rol_empresa) VALUES (%s,%s,%s,%s,%s)",
                (f['username'], f['email'].strip(),
                 generate_password_hash(password_plana), f['rol'], f.get('rol_empresa', ''))
            )
            cnx.commit(); cur.close(); cnx.close()
            
            try:
                auditor.enviar_correo_credenciales(f['email'].strip(), f['rol'], password_plana)
                flash('Usuario creado y credenciales enviadas por correo.', 'success')
            except Exception as e_mail:
                flash(f'Usuario creado, pero hubo un error al enviar el correo: {e_mail}', 'warning')
                
            return redirect(url_for('usuarios'))
        except Exception as e:
            flash(f'Error: {e}', 'error')
    return render_template('usuario_form.html', usuario=None, accion='Nuevo')

@app.route('/usuarios/editar/<int:uid>', methods=['GET','POST'])
@login_required
@lider_required
def usuario_editar(uid):
    cnx = auditor.get_conn()
    cur = cnx.cursor(dictionary=True)
    cur.execute("SELECT * FROM usuarios WHERE id=%s", (uid,))
    u = cur.fetchone(); cur.close(); cnx.close()
    if not u: abort(404)
    if request.method == 'POST':
        f = request.form
        new_pass = f.get('password','').strip()
        try:
            cnx = auditor.get_conn(); cur = cnx.cursor()
            if new_pass:
                cur.execute("UPDATE usuarios SET username=%s,email=%s,password_hash=%s,rol=%s,rol_empresa=%s,activo=%s WHERE id=%s",
                    (f['username'],f['email'],generate_password_hash(new_pass),f['rol'], f.get('rol_empresa', ''),
                     1 if f.get('activo') else 0, uid))
            else:
                cur.execute("UPDATE usuarios SET username=%s,email=%s,rol=%s,rol_empresa=%s,activo=%s WHERE id=%s",
                    (f['username'],f['email'],f['rol'], f.get('rol_empresa', ''),
                     1 if f.get('activo') else 0, uid))
            cnx.commit(); cur.close(); cnx.close()
            flash('Usuario actualizado.', 'success')
            return redirect(url_for('usuarios'))
        except Exception as e:
            flash(f'Error: {e}', 'error')
    return render_template('usuario_form.html', usuario=u, accion='Editar')

@app.route('/usuarios/eliminar/<int:uid>', methods=['POST'])
@login_required
@lider_required
def usuario_eliminar(uid):
    if uid == current_user.id:
        flash('No puedes eliminarte a ti mismo.', 'error')
        return redirect(url_for('usuarios'))
    try:
        cnx = auditor.get_conn(); cur = cnx.cursor()
        cur.execute("DELETE FROM usuarios WHERE id=%s", (uid,))
        cnx.commit(); cur.close(); cnx.close()
        flash('Usuario eliminado.', 'success')
    except Exception as e:
        flash(f'Error: {e}', 'error')
    return redirect(url_for('usuarios'))

# CONFIGURACIÓN GLOBAL — Correos de líderes y credenciales SMTP

@app.route('/configuracion', methods=['GET', 'POST'])
@login_required
@lider_required
def configuracion():
    if request.method == 'POST':
        f = request.form
        try:
            cnx = auditor.get_conn()
            cur = cnx.cursor()
            query = "INSERT INTO configuracion (clave, valor) VALUES (%s, %s) ON DUPLICATE KEY UPDATE valor = VALUES(valor)"
            cur.executemany(query, [
                ('lider_implementacion', f['lider_implementacion']),
                ('lider_seo', f['lider_seo']),
                ('lider_operaciones', f['lider_operaciones']),
                ('smtp_email', f['smtp_email']),
                ('smtp_password', f['smtp_password'])
            ])
            cnx.commit()
            cur.close()
            cnx.close()
            flash('Configuración global actualizada.', 'success')
            return redirect(url_for('configuracion'))
        except Exception as e:
            flash(f'Error guardando configuración: {e}', 'error')

    # Si es GET, cargamos la configuración actual para mostrarla en el formulario
    try:
        cnx = auditor.get_conn()
        cur = cnx.cursor(dictionary=True)
        cur.execute("SELECT clave, valor FROM configuracion")
        rows = cur.fetchall()
        cur.close()
        cnx.close()
        conf = {row['clave']: row['valor'] for row in rows}
    except Exception as e:
        flash(f'Error cargando configuración: {e}', 'error')
        conf = {}
        
    return render_template('config.html', config=conf)

# REPORTES — Dashboard analítico con KPIs, gráficas y tabla resumen

@app.route('/reportes')
@login_required
def reportes():
    try:
        cnx = auditor.get_conn()
        cur = cnx.cursor(dictionary=True)

        # Primero sacamos los números grandes para las tarjetas KPI del encabezado
        cur.execute("SELECT COUNT(*) as total FROM historico")
        total_auditorias = cur.fetchone()['total']

        cur.execute("SELECT COUNT(*) as total FROM marcas")
        total_marcas = cur.fetchone()['total']

        cur.execute("""
            SELECT COALESCE(SUM(urls_redireccion + urls_no_creadas + urls_meta_titulo +
                   urls_meta_descripcion + urls_h1 + urls_desc_corta + urls_desc_larga), 0) as total
            FROM historico
        """)
        total_errores = cur.fetchone()['total']

        cur.execute("""
            SELECT COUNT(DISTINCT marca_id) as total FROM historico
            WHERE fecha_ejecucion >= DATE_SUB(NOW(), INTERVAL 30 DAY)
        """)
        marcas_activas = cur.fetchone()['total']

        # Datos agrupados por marca para la gráfica de barras y la tabla resumen
        cur.execute("""
            SELECT m.id, m.marca, COUNT(h.id) as total_auditorias,
                   COALESCE(SUM(h.urls_redireccion + h.urls_no_creadas + h.urls_meta_titulo +
                       h.urls_meta_descripcion + h.urls_h1 + h.urls_desc_corta + h.urls_desc_larga), 0) as total_errores,
                   COALESCE(SUM(h.urls_contenido_arreglado), 0) as total_arreglados,
                   COALESCE(SUM(h.urls_contenido_cambiado), 0) as total_cambiados,
                   MAX(h.numero_revision) as ultima_revision,
                   MAX(h.fecha_ejecucion) as ultima_fecha
            FROM marcas m
            LEFT JOIN historico h ON h.marca_id = m.id
            GROUP BY m.id, m.marca
            ORDER BY total_auditorias DESC
        """)
        resumen_marcas = cur.fetchall()

        # Proporción de cada tipo de error para el gráfico de dona
        cur.execute("""
            SELECT
                COALESCE(SUM(urls_redireccion), 0)      as redireccion,
                COALESCE(SUM(urls_no_creadas), 0)       as no_creadas,
                COALESCE(SUM(urls_meta_titulo), 0)      as meta_titulo,
                COALESCE(SUM(urls_meta_descripcion), 0) as meta_descripcion,
                COALESCE(SUM(urls_h1), 0)               as h1,
                COALESCE(SUM(urls_desc_corta), 0)       as desc_corta,
                COALESCE(SUM(urls_desc_larga), 0)       as desc_larga
            FROM historico
        """)
        distribucion = cur.fetchone()

        # Lista simple de marcas para el dropdown del gráfico de evolución temporal
        cur.execute("SELECT id, marca FROM marcas ORDER BY marca")
        marcas_lista = cur.fetchall()

        cur.close(); cnx.close()

        # MySQL devuelve Decimals y datetimes que JSON no entiende, así que los convertimos
        for m in resumen_marcas:
            m['total_errores'] = int(m['total_errores']) if m['total_errores'] is not None else 0
            m['total_arreglados'] = int(m['total_arreglados']) if m['total_arreglados'] is not None else 0
            m['total_cambiados'] = int(m['total_cambiados']) if m['total_cambiados'] is not None else 0
            if m['ultima_fecha'] and hasattr(m['ultima_fecha'], 'isoformat'):
                m['ultima_fecha'] = m['ultima_fecha'].isoformat()
                
        if distribucion:
            for k in distribucion:
                distribucion[k] = int(distribucion[k]) if distribucion[k] is not None else 0

        return render_template('reportes.html',
            total_auditorias=total_auditorias,
            total_marcas=total_marcas,
            total_errores=int(total_errores),
            marcas_activas=marcas_activas,
            resumen_marcas=resumen_marcas,
            distribucion=distribucion,
            marcas_lista=marcas_lista
        )
    except Exception as e:
        flash(f"Error cargando reportes: {e}", "error")
        return render_template('reportes.html',
            total_auditorias=0, total_marcas=0,
            total_errores=0, marcas_activas=0,
            resumen_marcas=[], distribucion=None, marcas_lista=[]
        )


@app.route('/reportes/api/datos')
@login_required
def reportes_api_datos():
    """API interna que devuelve los datos históricos en JSON para que
    Chart.js dibuje la gráfica de evolución de errores."""
    marca_id = request.args.get('marca_id', type=int)
    try:
        cnx = auditor.get_conn()
        cur = cnx.cursor(dictionary=True)

        if marca_id:
            cur.execute("""
                SELECT fecha_ejecucion, numero_revision,
                       urls_procesadas, urls_redireccion, urls_no_creadas,
                       urls_meta_titulo, urls_meta_descripcion,
                       urls_h1, urls_desc_corta, urls_desc_larga,
                       urls_contenido_arreglado, urls_contenido_cambiado
                FROM historico
                WHERE marca_id = %s
                ORDER BY fecha_ejecucion ASC
            """, (marca_id,))
        else:
            # Sin filtro de marca: agrupamos todos los errores por fecha
            cur.execute("""
                SELECT DATE(fecha_ejecucion) as fecha_ejecucion,
                       SUM(urls_redireccion + urls_no_creadas + urls_meta_titulo +
                           urls_meta_descripcion + urls_h1 + urls_desc_corta + urls_desc_larga) as total_errores,
                       SUM(urls_contenido_arreglado) as urls_contenido_arreglado,
                       SUM(urls_contenido_cambiado) as urls_contenido_cambiado,
                       COUNT(*) as auditorias
                FROM historico
                GROUP BY DATE(fecha_ejecucion)
                ORDER BY fecha_ejecucion ASC
            """)

        rows = cur.fetchall()
        cur.close(); cnx.close()

        # Convertimos las fechas a string legible para el eje X de la gráfica
        for r in rows:
            if 'fecha_ejecucion' in r and hasattr(r['fecha_ejecucion'], 'strftime'):
                r['fecha_ejecucion'] = r['fecha_ejecucion'].strftime('%d/%m/%Y')

        return jsonify({'ok': True, 'datos': rows})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})



# ARRANQUE — Migraciones automáticas de BD y lanzamiento del servidor

if __name__ == '__main__':
    # Al arrancar, verificamos si la BD necesita columnas nuevas y las agregamos
    try:
        cnx = auditor.get_conn()
        cur = cnx.cursor()
        cur.execute("SHOW COLUMNS FROM historico LIKE 'urls_contenido_arreglado'")
        if not cur.fetchone():
            print("[Migration] Agregando columnas de contenidos arreglados y cambiados...")
            cur.execute("ALTER TABLE historico ADD COLUMN urls_contenido_arreglado INT DEFAULT 0")
            cur.execute("ALTER TABLE historico ADD COLUMN urls_contenido_cambiado INT DEFAULT 0")
            cnx.commit()
            print("[Migration] Columnas creadas con éxito. Iniciando backfill...")
            auditor.backfill_historico_cambios()
        cur.close(); cnx.close()
    except Exception as e:
        print(f"[Migration] Error al migrar base de datos: {e}")

    auditor.iniciar_scheduler()
    app.run(host='0.0.0.0', port=5000, debug=False)
