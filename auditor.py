"""
auditor.py — rastreo, comparación y reportes de SEOTracker.
"""
# --- stdlib ---
import re
import os
import difflib
import datetime
import time
import random
import unicodedata
import smtplib
import threading
import json
from collections import defaultdict

# --- third-party ---
import pandas as pd
import schedule
import mysql.connector
import requests
import gspread
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import PatternFill, Font
from oauth2client.service_account import ServiceAccountCredentials

# --- config (se sobreescribe desde app.py) ---
DB_CONFIG = {
    'host': '127.0.0.1',
    'user': 'root',
    'password': '',
    'database': 'seo_auditor'
}

SMTP_CONFIG = {
    'email_emisor': 'fabianperez@estrategiamic.com',
    'app_password': 'etpr gevb fcvt nynw',
    'servidor': 'smtp.gmail.com',
    'puerto': 587
}

PATH_CREDENTIALS_SHEETS = 'credentials.json'
MESES_ES = {
    1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
    5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
    9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
}

# --- limpieza de texto ---

def limpiar_html(texto):
    if pd.isna(texto) or str(texto).strip().lower() in ('nan', ''):
        return 'Vacío'
    if isinstance(texto, str):
        # normalizar unicode primero (NFC vs NFD, \xa0, etc)
        texto = unicodedata.normalize('NFKC', texto)
        texto = texto.replace('\n', ' ').replace('\r', ' ').replace('\xa0', ' ')
        # eliminar caracteres invisibles que pueden causar falsas diferencias
        for ch in ('\u200b', '\u200c', '\u200d', '\u00ad', '\ufeff', '\u2060', '\u202f', '\u2009', '\u2002', '\u2003'):
            texto = texto.replace(ch, '')
        soup = BeautifulSoup(texto, 'html.parser')
        limpio = soup.get_text(separator=' ', strip=True)
        limpio = re.sub(r'(?i)\b(FAQs\s+H2|H[23])[:\s\n\r]*', ' ', limpio)
        limpio = re.sub(r'\s+([.,:;!?\)])', r'\1', limpio)
        limpio = re.sub(r'\s+', ' ', limpio).strip()
        return limpio if limpio and limpio.lower() != 'nan' else 'Vacío'
    clean = unicodedata.normalize('NFKC', str(texto)).replace('\xa0', ' ').strip()
    clean = re.sub(r'\s+', ' ', clean).strip()
    return 'Vacío' if not clean or clean.lower() == 'nan' else clean


def normalizar(serie):
    # los mismos invisibles que en limpiar_html, con \u00a0 convertido a espacio
    _INVISIBLE = (
        '\u200b',  # zero-width space
        '\u200c',  # zero-width non-joiner
        '\u200d',  # zero-width joiner
        '\u00ad',  # soft hyphen
        '\ufeff',  # BOM / zero-width no-break space
        '\u2060',  # word joiner
        '\u00a0',  # non-breaking space (→ convertir a espacio normal)
        '\u202f',  # narrow no-break space
        '\u2009',  # thin space
        '\u2002',  # en space
        '\u2003',  # em space
    )

    def _norm(x):
        if pd.isna(x): return ""
        s = unicodedata.normalize('NFKC', str(x))
        for ch in _INVISIBLE:
            s = s.replace(ch, ' ' if ch == '\u00a0' else '')
        return s

    return serie.apply(_norm).str.replace(r'\s+', ' ', regex=True).str.strip().str.lower()


def extraer_h2h3_html(soup):
    tags = soup.find_all(['h2', 'h3'])
    out = [f"[{t.name.upper()}] {t.get_text(strip=True)}" for t in tags if t.get_text(strip=True)]
    return " | ".join(out) if out else "Ninguno"


def _normalizar_para_comparar(texto):
    """Quita acentos, pasa a minúsculas y colapsa espacios para comparar textos."""
    if not texto:
        return ""
    s = unicodedata.normalize('NFKD', str(texto)).encode('ASCII', 'ignore').decode('utf-8')
    return re.sub(r'\s+', ' ', s).strip().lower()


def detectar_h1_duplicado_en_desc_larga(h1_vivo, desc_larga_vivo, umbral=0.95):
    """
    Revisa si algún H2 dentro de la descripción larga (rastreada en vivo)
    es muy similar al H1 de la URL.
    Retorna 'H1 duplicado en descripción larga' si se detecta similitud,
    o '' (cadena vacía) si no hay problema.
    El umbral de similitud es 0.82 (82%) usando SequenceMatcher.
    """
    if not h1_vivo or not desc_larga_vivo:
        return ''
    from difflib import SequenceMatcher
    from bs4 import BeautifulSoup as _BS

    h1_norm = _normalizar_para_comparar(h1_vivo)
    if not h1_norm:
        return ''

    # intentar parsear la desc larga como HTML para extraer H2s reales
    try:
        soup_dl = _BS(str(desc_larga_vivo), 'html.parser')
        h2_tags = [t.get_text(strip=True) for t in soup_dl.find_all('h2') if t.get_text(strip=True)]
    except Exception:
        h2_tags = []

    # si no encontró H2s con HTML, buscar patrones [H2] en texto plano (formato de CONTENIDO_SF_H2_H3)
    if not h2_tags:
        for parte in str(desc_larga_vivo).split('|'):
            parte = parte.strip()
            if parte.upper().startswith('[H2]'):
                h2_tags.append(re.sub(r'^\[H2\]\s*', '', parte, flags=re.IGNORECASE).strip())

    for h2_texto in h2_tags:
        h2_norm = _normalizar_para_comparar(h2_texto)
        if not h2_norm:
            continue
        ratio = SequenceMatcher(None, h1_norm, h2_norm).ratio()
        # también detectar cuando el H1 aparece como prefijo exacto dentro del H2
        h1_es_prefijo = h2_norm.startswith(h1_norm) and len(h1_norm) >= 10
        if ratio >= umbral or h1_es_prefijo:
            return 'H1 duplicado en descripción larga'

    return ''


# --- db helpers ---

_COLUMNS_CHECKED = False


def get_conn():
    global _COLUMNS_CHECKED
    cnx = mysql.connector.connect(**DB_CONFIG)
    if not _COLUMNS_CHECKED:
        try:
            cur = cnx.cursor()
            cur.execute("SHOW COLUMNS FROM marcas LIKE 'headers_personalizados'")
            if not cur.fetchone():
                cur.execute("ALTER TABLE marcas ADD COLUMN headers_personalizados TEXT DEFAULT NULL")
                cnx.commit()

            cur.execute("SHOW COLUMNS FROM marcas LIKE 'fecha_headers'")
            if not cur.fetchone():
                cur.execute("ALTER TABLE marcas ADD COLUMN fecha_headers DATETIME DEFAULT NULL")
                cnx.commit()
                # inicializar headers existentes a la fecha actual para no disparar alertas de inmediato
                cur.execute("UPDATE marcas SET fecha_headers = CURRENT_TIMESTAMP WHERE headers_personalizados IS NOT NULL AND headers_personalizados != '' AND fecha_headers IS NULL")
                cnx.commit()

            cur.execute("SHOW COLUMNS FROM marcas LIKE 'alerta_headers_enviada'")
            if not cur.fetchone():
                cur.execute("ALTER TABLE marcas ADD COLUMN alerta_headers_enviada BOOLEAN DEFAULT FALSE")
                cnx.commit()

            cur.execute("SHOW COLUMNS FROM usuarios LIKE 'rol_empresa'")
            if not cur.fetchone():
                cur.execute("ALTER TABLE usuarios ADD COLUMN rol_empresa VARCHAR(100) DEFAULT NULL")
                cnx.commit()

            cur.close()
            _COLUMNS_CHECKED = True
        except Exception as e:
            print(f"Error verificando nuevas columnas: {e}")
    return cnx


def obtener_marcas(email_filtro=None):
    cnx = get_conn()
    cur = cnx.cursor(dictionary=True)
    query = """
        SELECT m.*, c.nombre as celula, c.implementador_email, c.gerente_email, c.pm_email 
        FROM marcas m 
        JOIN celulas c ON m.celula_id = c.id
    """
    params = ()
    if email_filtro:
        query += " WHERE %s IN (c.implementador_email, c.gerente_email, c.pm_email)"
        params = (email_filtro,)
    cur.execute(query, params)
    rows = cur.fetchall()
    cur.close(); cnx.close()
    return rows


def obtener_marca(marca_id):
    cnx = get_conn()
    cur = cnx.cursor(dictionary=True)
    cur.execute("""
        SELECT m.*, c.nombre as celula, c.implementador_email, c.gerente_email, c.pm_email 
        FROM marcas m 
        JOIN celulas c ON m.celula_id = c.id
        WHERE m.id = %s
    """, (marca_id,))
    row = cur.fetchone()
    cur.close(); cnx.close()
    return row


def obtener_celulas():
    cnx = get_conn()
    cur = cnx.cursor(dictionary=True)
    cur.execute("SELECT * FROM celulas ORDER BY nombre")
    rows = cur.fetchall()
    cur.close(); cnx.close()
    return rows


def obtener_celula(celula_id):
    cnx = get_conn()
    cur = cnx.cursor(dictionary=True)
    cur.execute("SELECT * FROM celulas WHERE id = %s", (celula_id,))
    row = cur.fetchone()
    cur.close(); cnx.close()
    return row


def obtener_usuarios_por_rol(rol):
    cnx = get_conn()
    cur = cnx.cursor(dictionary=True)
    cur.execute("SELECT * FROM usuarios WHERE rol = %s AND activo = 1 ORDER BY email", (rol,))
    rows = cur.fetchall()
    cur.close(); cnx.close()
    return rows


def obtener_configuracion():
    cnx = get_conn()
    cur = cnx.cursor(dictionary=True)
    try:
        cur.execute("SELECT clave, valor FROM configuracion")
        rows = cur.fetchall()
        cur.close(); cnx.close()
        return {r['clave']: r['valor'] for r in rows}
    except:
        cur.close(); cnx.close()
        return {}


def obtener_config_smtp():
    conf = obtener_configuracion()
    return {
        'email_emisor': conf.get('smtp_email') or SMTP_CONFIG['email_emisor'],
        'app_password': conf.get('smtp_password') or SMTP_CONFIG['app_password'],
        'servidor': SMTP_CONFIG['servidor'],
        'puerto': SMTP_CONFIG['puerto']
    }


def obtener_ultimo_historico(marca_id):
    cnx = get_conn()
    cur = cnx.cursor(dictionary=True)
    cur.execute(
        "SELECT * FROM historico WHERE marca_id = %s ORDER BY fecha_ejecucion DESC LIMIT 1",
        (marca_id,)
    )
    row = cur.fetchone()
    cur.close(); cnx.close()
    return row


def obtener_ultimos_historicos_no_responsive(marca_id):
    cnx = get_conn()
    cur = cnx.cursor(dictionary=True)

    cur.execute("""
        SELECT * FROM historico 
        WHERE marca_id = %s AND (ruta_reporte_excel LIKE '%Mobile%' OR ruta_reporte_excel LIKE '%mobile%')
        ORDER BY fecha_ejecucion DESC LIMIT 1
    """, (marca_id,))
    mobile = cur.fetchone()

    cur.execute("""
        SELECT * FROM historico 
        WHERE marca_id = %s AND (ruta_reporte_excel LIKE '%Desktop%' OR ruta_reporte_excel LIKE '%desktop%')
        ORDER BY fecha_ejecucion DESC LIMIT 1
    """, (marca_id,))
    desktop = cur.fetchone()

    cur.close(); cnx.close()
    return mobile, desktop


def siguiente_revision(marca_id):
    cnx = get_conn()
    cur = cnx.cursor(dictionary=True)
    cur.execute("SELECT * FROM historico WHERE marca_id=%s ORDER BY fecha_ejecucion DESC LIMIT 1", (marca_id,))
    row = cur.fetchone()
    cur.close(); cnx.close()
    if row:
        errores = (row['urls_redireccion'] + row['urls_no_creadas'] +
                   row['urls_meta_titulo'] + row['urls_meta_descripcion'] +
                   row['urls_h1'] + row['urls_desc_corta'] + row['urls_desc_larga'])
        if errores > 0:
            return row['numero_revision'] + 1
        else:
            return 1
    return 1


def guardar_historico(marca_id, stats, ruta_excel, num_rev=None, arreglados=0, cambiados=0):
    if num_rev is None:
        num_rev = siguiente_revision(marca_id)
    cnx = get_conn()
    cur = cnx.cursor()
    cur.execute(
        "INSERT INTO historico (marca_id, numero_revision, fecha_ejecucion, urls_procesadas, "
        "urls_redireccion, urls_no_creadas, urls_meta_titulo, urls_meta_descripcion, "
        "urls_h1, urls_desc_corta, urls_desc_larga, ruta_reporte_excel, urls_contenido_arreglado, urls_contenido_cambiado) "
        "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
        (marca_id, num_rev, datetime.datetime.now(),
         stats['total'], stats['redireccion'], stats['no_creadas'],
         stats['meta_titulo'], stats['meta_descripcion'],
         stats['h1'], stats['desc_corta'], stats['desc_larga'], ruta_excel,
         arreglados, cambiados)
    )
    cnx.commit(); cur.close(); cnx.close()
    return num_rev


def calcular_cambios_contenido(path_prev, path_curr, errores_prev, errores_curr):
    arreglados = max(0, errores_prev - errores_curr)
    if not path_prev or not path_curr:
        return arreglados, 0
    if not os.path.exists(path_prev) or not os.path.exists(path_curr):
        return arreglados, 0
    try:
        df_prev = pd.read_excel(path_prev)
        df_curr = pd.read_excel(path_curr)

        url_col_prev = 'URL' if 'URL' in df_prev.columns else ('Dirección' if 'Dirección' in df_prev.columns else None)
        url_col_curr = 'URL' if 'URL' in df_curr.columns else ('Dirección' if 'Dirección' in df_curr.columns else None)

        if not url_col_prev or not url_col_curr:
            return arreglados, 0

        live_cols = [
            'Meta Titulo (Rastreo en Vivo)',
            'Meta Descripción (Rastreo en Vivo)',
            'H1 (Rastreo en Vivo)',
            'Descripción Corta (Rastreo en Vivo)',
            'Descripción Larga (Rastreo en Vivo)'
        ]

        cols_prev = [url_col_prev] + [c for c in live_cols if c in df_prev.columns]
        cols_curr = [url_col_curr] + [c for c in live_cols if c in df_curr.columns]

        df_p = df_prev[cols_prev].copy().rename(columns={url_col_prev: 'url_join'})
        df_c = df_curr[cols_curr].copy().rename(columns={url_col_curr: 'url_join'})

        merged = pd.merge(df_p, df_c, on='url_join', suffixes=('_prev', '_curr'))

        cambiados = 0
        for _, row in merged.iterrows():
            changed = False
            for col in live_cols:
                col_prev = f"{col}_prev"
                col_curr = f"{col}_curr"
                if col_prev in merged.columns and col_curr in merged.columns:
                    val_prev = str(row[col_prev]).strip().lower()
                    val_curr = str(row[col_curr]).strip().lower()
                    if val_prev != val_curr:
                        changed = True
                        break
            if changed:
                cambiados += 1

        return arreglados, cambiados
    except Exception as e:
        print(f"Error calculando cambios de contenido: {e}")
        return arreglados, 0


def backfill_historico_cambios():
    try:
        cnx = get_conn()
        cur = cnx.cursor(dictionary=True)
        cur.execute("SELECT id, marca_id, ruta_reporte_excel, urls_redireccion, urls_no_creadas, urls_meta_titulo, urls_meta_descripcion, urls_h1, urls_desc_corta, urls_desc_larga FROM historico ORDER BY fecha_ejecucion ASC")
        rows = cur.fetchall()
        cur.close(); cnx.close()

        marca_hist = defaultdict(list)
        for r in rows:
            marca_hist[r['marca_id']].append(r)

        cnx = get_conn()
        cur = cnx.cursor()
        for marca_id, hist_list in marca_hist.items():
            for i in range(1, len(hist_list)):
                curr = hist_list[i]
                prev = hist_list[i - 1]

                errores_prev = (prev['urls_redireccion'] + prev['urls_no_creadas'] +
                                prev['urls_meta_titulo'] + prev['urls_meta_descripcion'] +
                                prev['urls_h1'] + prev['urls_desc_corta'] + prev['urls_desc_larga'])
                errores_curr = (curr['urls_redireccion'] + curr['urls_no_creadas'] +
                                curr['urls_meta_titulo'] + curr['urls_meta_descripcion'] +
                                curr['urls_h1'] + curr['urls_desc_corta'] + curr['urls_desc_larga'])

                arr, camb = calcular_cambios_contenido(
                    prev['ruta_reporte_excel'], curr['ruta_reporte_excel'],
                    errores_prev, errores_curr
                )

                cur.execute("""
                    UPDATE historico 
                    SET urls_contenido_arreglado = %s, urls_contenido_cambiado = %s 
                    WHERE id = %s
                """, (arr, camb, curr['id']))
        cnx.commit()
        cur.close(); cnx.close()
        print("Backfill de cambios de contenidos de históricos finalizado con éxito.")
    except Exception as e:
        print(f"Error en backfill de históricos: {e}")


def obtener_historico_marca(marca_id, limit=10):
    cnx = get_conn()
    cur = cnx.cursor(dictionary=True)
    cur.execute(
        "SELECT * FROM historico WHERE marca_id=%s ORDER BY fecha_ejecucion DESC LIMIT %s",
        (marca_id, limit)
    )
    rows = cur.fetchall()
    cur.close(); cnx.close()
    return rows


def obtener_correos_disponibles(marca_id):
    correos = []
    seen = set()

    def add(label, valor):
        if not valor:
            return
        for e in valor.replace(';', ',').split(','):
            e = e.strip()
            if e and e not in seen:
                seen.add(e)
                correos.append({'label': label, 'email': e})

    try:
        marca = obtener_marca(marca_id)
        if marca:
            add(f"Implementador ({marca.get('celula', '')})", marca.get('implementador_email'))
            add(f"Gerente ({marca.get('celula', '')})", marca.get('gerente_email'))
            add(f"PM ({marca.get('celula', '')})", marca.get('pm_email'))
    except Exception:
        pass

    try:
        conf = obtener_configuracion()
        add('Líder de Implementación', conf.get('lider_implementacion'))
        add('Líder SEO', conf.get('lider_seo'))
        add('Líder de Operaciones', conf.get('lider_operaciones'))
    except Exception:
        pass

    try:
        cnx = get_conn()
        cur = cnx.cursor(dictionary=True)
        cur.execute("SELECT username, email FROM usuarios WHERE activo=1 AND email != '' ORDER BY username")
        for u in cur.fetchall():
            add(f"Usuario: {u['username']}", u.get('email'))
        cur.close(); cnx.close()
    except Exception:
        pass

    return correos


# --- sheets ---

def descargar_hoja_sheets(url_sheet, fila_encabezado=2):
    if not os.path.exists(PATH_CREDENTIALS_SHEETS):
        print("ERROR: credentials.json no encontrado")
        return pd.DataFrame()
    try:
        scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(PATH_CREDENTIALS_SHEETS, scope)
        client = gspread.authorize(creds)
        m = re.search(r'/d/([a-zA-Z0-9-_]+)', url_sheet)
        if not m:
            print("URL de Sheets inválida")
            return pd.DataFrame()
        sheet = client.open_by_key(m.group(1)).sheet1

        all_values = sheet.get_all_values()

        if len(all_values) < fila_encabezado:
            print("La hoja no tiene suficientes filas")
            return pd.DataFrame()

        headers = all_values[fila_encabezado - 1]
        data_rows = all_values[fila_encabezado:]
        headers = [str(h).strip() for h in headers]
        df = pd.DataFrame(data_rows, columns=headers)
        df = df.loc[:, ~df.columns.duplicated()].copy()
        df = df.replace('', pd.NA).dropna(how='all').fillna('')

        # si hay columna "Reoptimización X", sus valores sobreescriben la columna X original
        for col in list(df.columns):
            col_str = str(col).strip()
            if col_str.lower().startswith('reoptimización ') or col_str.lower().startswith('reoptimizacion '):
                prefijo_len = len('reoptimización ') if col_str.lower().startswith('reoptimización ') else len('reoptimizacion ')
                col_orig_name = col_str[prefijo_len:].strip()
                col_orig_match = next((c for c in df.columns if str(c).strip().lower() == col_orig_name.lower()), None)
                if col_orig_match:
                    reopt_series = df[col].replace(r'^\s*$', pd.NA, regex=True).replace('', pd.NA)
                    df[col_orig_match] = reopt_series.combine_first(df[col_orig_match].replace(r'^\s*$', pd.NA, regex=True)).fillna('')

        return df
    except Exception as e:
        print(f"Error Sheets: {e}")
        return pd.DataFrame()


# --- rastreo ---

_UA_POOL = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:123.0) Gecko/20100101 Firefox/123.0',
]

UA_MOBILE = 'Mozilla/5.0 (Linux; Android 10; K) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Mobile Safari/537.36'
UA_DESKTOP = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'


def get_by_selectors(soup, selector_str):
    if not selector_str:
        return ""
    for sel in [s.strip() for s in selector_str.split(',') if s.strip()]:
        try:
            elem = soup.select_one(sel)
            if elem:
                txt = elem.get_text(separator=' ', strip=True)
                if txt:
                    return txt
        except Exception:
            continue
    return ""


def parse_headers_personalizados(headers_str):
    if not headers_str or not str(headers_str).strip():
        return {}
    headers_str = str(headers_str).strip()
    if headers_str.startswith('{') and headers_str.endswith('}'):
        try:
            parsed = json.loads(headers_str)
            if isinstance(parsed, dict):
                return {str(k).strip(): str(v).strip() for k, v in parsed.items()}
        except Exception:
            pass
    out = {}
    for line in headers_str.splitlines():
        line = line.strip()
        if not line or line.startswith('#'):
            continue
        if ':' in line:
            key, val = line.split(':', 1)
            out[key.strip()] = val.strip()
    return out


def fetch_url(url, marca_db, user_agent=None, session=None):
    result = {
        'Dirección': url, 'Código de respuesta': None, 'Código_Final': None, 'Código_Original': None,
        'Título 1': "", 'Meta Description 1': "",
        'H1-1': "", 'Descripción corta 1': "", 'Descripción larga 1': "",
        'CONTENIDO_SF_H2_H3': "Ninguno", 'Status': 'Error', 'Error': None
    }

    ua = user_agent if user_agent else random.choice(_UA_POOL)
    req_headers = {
        'User-Agent': ua,
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
        'Accept-Language': 'es-ES,es;q=0.9,en-US;q=0.8,en;q=0.7',
        'Sec-Ch-Ua': '"Chromium";v="122", "Not(A:Brand";v="24", "Google Chrome";v="122"',
        'Sec-Ch-Ua-Mobile': '?0',
        'Sec-Ch-Ua-Platform': '"Windows"',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'none',
        'Upgrade-Insecure-Requests': '1'
    }

    custom_headers = parse_headers_personalizados(marca_db.get('headers_personalizados', ''))
    if custom_headers:
        req_headers.update(custom_headers)

    http_client = session if session else requests

    try:
        # pausa por URL para evitar rate limits (429)
        time.sleep(random.uniform(2.5, 4.5))
        resp = http_client.get(url, headers=req_headers, timeout=15)

        def process_resp(r):
            final = r.status_code
            orig = r.history[0].status_code if r.history else final
            if r.history:
                texto = f"{orig} -> {r.url}"
            else:
                texto = final
            return final, orig, texto

        final_st, orig_st, text_st = process_resp(resp)
        result['Código_Final'] = final_st
        result['Código_Original'] = orig_st
        result['Código de respuesta'] = text_st

        # 429: un único reintento forzando requests limpio (sin session)
        if final_st == 429:
            time.sleep(5.0)
            resp = requests.get(url, headers=req_headers, timeout=15)
            final_st, orig_st, text_st = process_resp(resp)
            result['Código_Final'] = final_st
            result['Código_Original'] = orig_st
            result['Código de respuesta'] = text_st

        if final_st != 200:
            result['Error'] = f'HTTP {final_st}'
            return result

        soup = BeautifulSoup(resp.text, 'html.parser')
        result['Título 1'] = soup.title.string.strip() if soup.title and soup.title.string else ""
        mt = soup.find('meta', attrs={'name': 'description'}) or soup.find('meta', attrs={'property': 'og:description'})
        if mt:
            result['Meta Description 1'] = mt.get('content', '').strip()
        result['H1-1']               = get_by_selectors(soup, marca_db.get('clases_h1', ''))
        result['Descripción corta 1'] = get_by_selectors(soup, marca_db.get('clases_desc_corta', ''))
        result['Descripción larga 1'] = get_by_selectors(soup, marca_db.get('clases_desc_larga', ''))
        result['CONTENIDO_SF_H2_H3'] = extraer_h2h3_html(soup)
        result['Status'] = 'OK'
        return result

    except Exception as e:
        result['Error'] = str(e)
        return result


# --- formato excel ---

def _rich_diff(texto_a, texto_b, rojo_font, lado='rastreo'):
    """
    Genera CellRichText resaltando en rojo las diferencias.
    lado='doc'     → resalta lo que está en texto_a (Doc) pero no en texto_b (Rastreo)
    lado='rastreo' → resalta lo que está en texto_b (Rastreo) pero no en texto_a (Doc)
    """
    try:
        # tokenizar conservando espacios para reconstruir el texto fielmente
        wa = re.split(r'(\s+)', str(texto_a))
        wb = re.split(r'(\s+)', str(texto_b))
        matcher = difflib.SequenceMatcher(None, wa, wb)
        partes = []
        tiene_rojo = False

        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if lado == 'doc':
                # recorremos los tokens del Doc (wa)
                frag = ''.join(wa[i1:i2])
                if not frag:
                    continue
                if tag == 'equal':
                    partes.append(frag)
                elif tag in ('replace', 'delete'):
                    # lo que no está en el Rastreo → rojo
                    partes.append(TextBlock(rojo_font, frag))
                    tiene_rojo = True
                # 'insert' → token solo en el Rastreo, no aparece en el Doc → ignorar
            else:
                # lado='rastreo': recorremos los tokens del Rastreo (wb)
                frag = ''.join(wb[j1:j2])
                if not frag:
                    continue
                if tag == 'equal':
                    partes.append(frag)
                elif tag in ('replace', 'insert'):
                    # lo que no coincide con el Doc → rojo
                    partes.append(TextBlock(rojo_font, frag))
                    tiene_rojo = True
                # 'delete' → token solo en el Doc, no aparece en el Rastreo → ignorar

        return CellRichText(*partes) if tiene_rojo and partes else None
    except Exception:
        return None


def aplicar_formato_excel(path_excel):
    try:
        wb = load_workbook(path_excel)
        ws = wb.active
        headers = [c.value for c in ws[1]]

        def _norm_col(txt):
            if txt is None: return ""
            s = unicodedata.normalize('NFKD', str(txt)).encode('ASCII', 'ignore').decode('utf-8')
            return re.sub(r'\s+', ' ', s).strip().lower()

        headers_norm = [_norm_col(h) for h in headers]

        def idx(name):
            n = _norm_col(name)
            return headers_norm.index(n) + 1 if n in headers_norm else -1

        rojo_fill = PatternFill(start_color='FFB6C1', end_color='FFB6C1', fill_type='solid')
        rojo_font = InlineFont(color='00FF0000')

        cols_coincide = [idx(c) for c in [
            'META_TITULO_COINCIDE', 'META_DESCRIPCION_COINCIDE',
            'H1_COINCIDE', 'DESCRIPCION_CORTA_COINCIDE', 'DESCRIPCION_LARGA_COINCIDE'
        ] if idx(c) != -1]

        pares = [
            (idx('Meta Título (Hoja de Contenido)'),        idx('Meta Titulo (Rastreo en Vivo)'),        idx('META_TITULO_COINCIDE')),
            (idx('Meta Descripción (Hoja de Contenido)'),   idx('Meta Descripción (Rastreo en Vivo)'),   idx('META_DESCRIPCION_COINCIDE')),
            (idx('H1 (Hoja de Contenido)'),                 idx('H1 (Rastreo en Vivo)'),                 idx('H1_COINCIDE')),
            (idx('Descripción Corta (Hoja de Contenido)'),  idx('Descripción Corta (Rastreo en Vivo)'),  idx('DESCRIPCION_CORTA_COINCIDE')),
            (idx('Descripción Larga (Hoja de Contenido)'),  idx('Descripción Larga (Rastreo en Vivo)'),  idx('DESCRIPCION_LARGA_COINCIDE')),
        ]

        rojo_fuerte_fill = PatternFill(start_color='FF4444', end_color='FF4444', fill_type='solid')
        blanco_font      = Font(color='FFFFFFFF', bold=True)
        col_h1_dup = idx('H1 DUPLICADO EN DESC. LARGA')

        for row in range(2, ws.max_row + 1):
            # fondo rojo en columnas COINCIDE que no dicen "Sí"
            for ci in cols_coincide:
                celda = ws.cell(row=row, column=ci)
                if str(celda.value).strip().lower() != 'sí':
                    celda.fill = rojo_fill

            # rojo intenso en la celda de H1 duplicado si tiene texto
            if col_h1_dup != -1:
                celda_dup = ws.cell(row=row, column=col_h1_dup)
                if str(celda_dup.value).strip():
                    celda_dup.fill = rojo_fuerte_fill
                    celda_dup.font = blanco_font

            for cc, cr, cq in pares:
                if -1 in (cc, cr, cq):
                    continue

                # la columna COINCIDE es la fuente de verdad
                val_coincide = str(ws.cell(row=row, column=cq).value).strip().lower()
                if val_coincide == 'sí':
                    continue

                val_esperado = ws.cell(row=row, column=cc).value
                val_rastreo  = ws.cell(row=row, column=cr).value

                # normalizar para el diff visual (eliminar caracteres invisibles)
                def _n(v):
                    if v is None: return ""
                    s = unicodedata.normalize('NFKC', str(v))
                    for ch in ('\u200b', '\u200c', '\u200d', '\u00ad', '\ufeff',
                               '\u2060', '\u202f', '\u2009', '\u2002', '\u2003'):
                        s = s.replace(ch, '')
                    s = s.replace('\u00a0', ' ')
                    return re.sub(r'\s+', ' ', s).strip()

                n_esperado = _n(val_esperado)
                n_rastreo  = _n(val_rastreo)

                if n_esperado == n_rastreo:
                    continue

                # columna Doc: resaltar lo que falta (está en Doc pero no en Rastreo)
                rich_doc = _rich_diff(n_esperado, n_rastreo, rojo_font, lado='doc')
                if rich_doc:
                    ws.cell(row=row, column=cc).value = rich_doc

                # columna Rastreo: resaltar lo que sobra/difiere
                rich_ras = _rich_diff(n_esperado, n_rastreo, rojo_font, lado='rastreo')
                if rich_ras:
                    ws.cell(row=row, column=cr).value = rich_ras

        wb.save(path_excel)
    except Exception as e:
        print(f"Error formato excel: {e}")


# --- correos ---

def enviar_correo_credenciales(destinatario, rol, password_plana):
    smtp_conf = obtener_config_smtp()
    msg = MIMEMultipart()
    msg['From'] = smtp_conf['email_emisor']
    msg['To'] = destinatario
    msg['Subject'] = "Tus accesos para SEOTracker"

    html = f"""
    <html>
      <body style='font-family: Arial, sans-serif; color: #333; line-height: 1.6;'>
        <h2>Bienvenido a SEOTracker</h2>
        <p>Se ha creado una cuenta para ti con el rol de <strong>{rol}</strong>.</p>
        <p>Tus credenciales de acceso son:</p>
        <ul>
          <li><strong>Usuario / Correo:</strong> {destinatario}</li>
          <li><strong>Contraseña:</strong> {password_plana}</li>
        </ul>
        <p>Te recomendamos cambiar tu contraseña una vez inicies sesión.</p>
        <br>
        <p>Atentamente,<br>El equipo de SEOTracker</p>
      </body>
    </html>
    """
    msg.attach(MIMEText(html, 'html'))

    try:
        print(f"Enviando credenciales a: {destinatario}...")
        server = smtplib.SMTP(smtp_conf['servidor'], smtp_conf['puerto'])
        server.starttls()
        server.login(smtp_conf['email_emisor'], smtp_conf['app_password'])
        server.sendmail(smtp_conf['email_emisor'], destinatario, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        print(f"Error enviando credenciales a {destinatario}: {e}")
        return False


def enviar_alerta_headers(marca_db):
    """
    Envía una alerta al implementador y al líder de implementación
    informando que los headers personalizados han expirado (llevan más de 90 días).
    """
    conf = obtener_configuracion()
    smtp_conf = obtener_config_smtp()

    destinatarios_set = set()
    if marca_db.get('implementador_email'):
        destinatarios_set.update([e.strip() for e in marca_db['implementador_email'].replace(';', ',').split(',') if e.strip()])
    if conf.get('lider_implementacion'):
        destinatarios_set.update([e.strip() for e in conf['lider_implementacion'].replace(';', ',').split(',') if e.strip()])

    if not destinatarios_set:
        return False

    destinatario_str = ", ".join(destinatarios_set)

    msg = MIMEMultipart()
    msg['From'] = smtp_conf['email_emisor']
    msg['To'] = destinatario_str
    msg['Subject'] = f"¡Alerta! Cabeceras HTTP expiradas para la marca {marca_db['marca']}"

    html = f"""
    <html>
      <body style='font-family: Arial, sans-serif; color: #333; line-height: 1.6;'>
        <h2>Alerta de Expiración de Cabeceras HTTP</h2>
        <p>Hola,</p>
        <p>Las cabeceras HTTP personalizadas (por ejemplo, tokens de sesión) para la marca <strong>{marca_db['marca']}</strong> han superado los 90 días de antigüedad.</p>
        <p>Es muy probable que hayan expirado o estén por expirar y las auditorías de esta marca comiencen a fallar (errores de redirección o bloqueo).</p>
        <p>Por favor, ingresa a <strong>SEOTracker</strong>, genera unas nuevas cabeceras para el entorno y actualízalas en la configuración de la marca.</p>
        <br>
        <p>Atentamente,<br>El equipo de SEOTracker</p>
      </body>
    </html>
    """
    msg.attach(MIMEText(html, 'html'))

    try:
        print(f"Enviando alerta de headers expirados a: {destinatario_str}...")
        server = smtplib.SMTP(smtp_conf['servidor'], smtp_conf['puerto'])
        server.starttls()
        server.login(smtp_conf['email_emisor'], smtp_conf['app_password'])
        server.sendmail(smtp_conf['email_emisor'], list(destinatarios_set), msg.as_string())
        server.quit()
        return True
    except Exception as e:
        print(f"Error enviando alerta de headers para {marca_db['marca']}: {e}")
        return False


def _delta_icon(prev, actual):
    if prev is None:
        return '-'
    try:
        diff = int(actual) - int(prev)
    except Exception:
        return str(actual)
    if diff > 0:
        return f"<span style='color:green'>+{diff} ↑</span>"
    elif diff < 0:
        return f"<span style='color:red'>{diff} ↓</span>"
    else:
        return f"<span style='color:gray'>0 ↔</span>"


def enviar_correo(marca_db, path_excel, stats, historico_anterior, num_revision, fecha_str):
    conf = obtener_configuracion()
    destinatarios_set = set()
    siempre = [
        marca_db.get('implementador_email'),
        marca_db.get('gerente_email'),
        marca_db.get('pm_email'),
        conf.get('lider_implementacion')
    ]
    for email in siempre:
        if email:
            destinatarios_set.update([e.strip() for e in email.replace(';', ',').split(',') if e.strip()])

    if num_revision >= 3 and conf.get('lider_seo'):
        destinatarios_set.update([e.strip() for e in conf.get('lider_seo').replace(';', ',').split(',') if e.strip()])

    if num_revision >= 4 and conf.get('lider_operaciones'):
        destinatarios_set.update([e.strip() for e in conf.get('lider_operaciones').replace(';', ',').split(',') if e.strip()])

    destinatario_str = ", ".join(destinatarios_set)

    try:
        msg = MIMEMultipart('alternative')
        smtp_conf = obtener_config_smtp()
        msg['From'] = smtp_conf['email_emisor']
        msg['To'] = destinatario_str
        msg['Subject'] = f"{fecha_str} - {marca_db.get('marca', '')} - Revisión {num_revision}"

        errores_totales = sum(stats.get(k, 0) for k in ['redireccion', 'no_creadas', 'meta_titulo', 'meta_descripcion', 'h1', 'desc_corta', 'desc_larga'])

        ant = historico_anterior
        if ant:
            ant_fecha = ant['fecha_ejecucion'].strftime(f"%d de {MESES_ES[ant['fecha_ejecucion'].month]} {ant['fecha_ejecucion'].year}")
            ant_reporte_txt = f"<p style='color:#888;font-size:13px;'>Reporte anterior: Revisión # {ant['numero_revision']} del {ant_fecha}.</p>"
        else:
            ant_fecha = "-"
            ant_reporte_txt = ""

        revision_color = '#2ecc71' if num_revision == 1 else ('#f1c40f' if num_revision <= 3 else '#e74c3c')

        if errores_totales == 0:
            body_html = f"""
            <!DOCTYPE html>
            <html>
            <body style='background:#0f0f1a;color:#e0e0e0;font-family:Arial,sans-serif;padding:30px;'>
              <div style='max-width:680px;margin:auto;background:#1a1a2e;border-radius:12px;padding:30px;text-align:center;'>
                <h2 style='color:#2ecc71;margin-bottom:4px;'>Auditoría SEO Exitosa</h2>
                <h3 style='color:#e0e0e0;margin-top:0;font-weight:400;'>{marca_db.get('marca', '')} — Revisión {num_revision}</h3>
                <p style='color:#aaa;margin-top:8px;'>Fecha de esta revisión: {fecha_str}</p>
                <div style='margin-top:20px;padding:20px;background:#16213e;border-radius:8px;'>
                  <h1 style='color:#4fc3f7;font-size:48px;margin:0;'>100%</h1>
                  <p style='font-size:18px;color:#e0e0e0;margin-top:10px;'>Todas las URLs coinciden perfectamente con el documento de contenidos.</p>
                  <p style='color:#aaa;'>Se procesaron exitosamente {stats['total']} URLs sin encontrar ninguna discrepancia ni error.</p>
                </div>
              </div>
            </body>
            </html>
            """
        else:
            rows = [
                ("URLs procesadas", stats['total'], ant['urls_procesadas'] if ant else None),
                ("URLs con redirección", stats['redireccion'], ant['urls_redireccion'] if ant else None),
                ("URLs no creadas", stats['no_creadas'], ant['urls_no_creadas'] if ant else None),
                ("Meta título no coincide", stats['meta_titulo'], ant['urls_meta_titulo'] if ant else None),
                ("Meta descripción no coinc.", stats['meta_descripcion'], ant['urls_meta_descripcion'] if ant else None),
                ("H1 no coincide", stats['h1'], ant['urls_h1'] if ant else None),
                ("Desc. corta no coincide", stats['desc_corta'], ant['urls_desc_corta'] if ant else None),
                ("Desc. larga no coincide", stats['desc_larga'], ant['urls_desc_larga'] if ant else None)
            ]

            table_rows_html = ""
            for label, actual, prev in rows:
                delta = _delta_icon(prev, actual)
                prev_str = str(prev) if prev is not None else "–"
                table_rows_html += f"""
                <tr>
                  <td style='padding:8px 12px;border-bottom:1px solid #2a2a3a;'>{label}</td>
                  <td style='padding:8px 12px;border-bottom:1px solid #2a2a3a;text-align:center;font-weight:bold;'>{actual}</td>
                  <td style='padding:8px 12px;border-bottom:1px solid #2a2a3a;text-align:center;color:#888;'>{prev_str}</td>
                  <td style='padding:8px 12px;border-bottom:1px solid #2a2a3a;'>{delta}</td>
                </tr>"""

            body_html = f"""
            <!DOCTYPE html>
            <html>
            <body style='background:#0f0f1a;color:#e0e0e0;font-family:Arial,sans-serif;padding:30px;'>
              <div style='max-width:680px;margin:auto;background:#1a1a2e;border-radius:12px;padding:30px;'>
                <h2 style='color:#4fc3f7;margin-bottom:4px;'>Reporte de Auditoría SEO</h2>
                <h3 style='color:#e0e0e0;margin-top:0;font-weight:400;'>{marca_db.get('marca', '')} — Revisión {num_revision}</h3>
                <p style='color:#aaa;margin-top:8px;'>Fecha de esta revisión: {fecha_str}</p>
                <p style='color:#aaa;margin-top:2px;'>Fecha de revisión anterior: {ant_fecha if ant else "-"}</p>
                <p style='color:#aaa;margin-top:8px;'>Se procesaron {stats['total']} URLs.</p>
                <p style='font-weight:bold;'>Marca: <span style='color:{revision_color}'>{marca_db.get('marca', '')}</span></p>
                <p style='font-weight:bold;'>Número de revisión: <span style='color:{revision_color}'>{num_revision}</span></p>
                <p style='line-height:1.7;'>Durante la comparación de contenidos entre el rastreo y el archivo de contenidos se evidenció lo siguiente:</p>
                <table style='width:100%;border-collapse:collapse;background:#12122a;border-radius:8px;overflow:hidden;'>
                
                  <thead>
                    <tr style='background:#16213e;'>
                      <th style='padding:6px 12px;'></th>
                      <th style='padding:6px 12px;text-align:center;color:#aaa;font-size:12px;font-weight:normal;'>{fecha_str}</th>
                      <th style='padding:6px 12px;text-align:center;color:#aaa;font-size:12px;font-weight:normal;'>{ant_fecha if ant else "-"}</th>
                      <th style='padding:6px 12px;'></th>
                    </tr>
                    <tr style='background:#16213e;'>
                      <th style='padding:10px 12px;text-align:left;color:#4fc3f7;'>Parametro de Revisión</th>
                      <th style='padding:10px 12px;text-align:center;color:#4fc3f7;'>Revisión actual # {num_revision}</th>
                      <th style='padding:10px 12px;text-align:center;color:#4fc3f7;'>Revisión anterior # {ant['numero_revision'] if ant else "-"}</th>
                      <th style='padding:10px 12px;text-align:left;color:#4fc3f7;'>Comparación de Revisiones</th>
                    </tr>
                  </thead>
                  <tbody>
                    {table_rows_html}
                  </tbody>
                </table>
                {ant_reporte_txt}
                <p style='margin-top:24px;color:#aaa;font-size:13px;'>Adjunto el archivo Excel con el detalle completo. Las celdas en rojo indican las diferencias encontradas.</p>
                <div style='margin-top:20px;padding:12px;background:#0a3d62;border-radius:8px;font-size:12px;color:#7fb3d3;'>Generado automáticamente por el Sistema SEOTracker.</div>
              </div>
            </body>
            </html>
            """

        msg.attach(MIMEText(body_html, 'html', 'utf-8'))

        if errores_totales > 0:
            if os.path.exists(path_excel):
                with open(path_excel, 'rb') as f:
                    part = MIMEApplication(f.read(), Name=os.path.basename(path_excel))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(path_excel)}"'
                msg.attach(part)

            if historico_anterior and 'ruta_reporte_excel' in historico_anterior:
                ant_path = os.path.abspath(historico_anterior['ruta_reporte_excel'])
                if os.path.exists(ant_path) and ant_path != os.path.abspath(path_excel):
                    with open(ant_path, 'rb') as f:
                        part_ant = MIMEApplication(f.read(), Name=f"ANTERIOR - {os.path.basename(ant_path)}")
                    part_ant['Content-Disposition'] = f'attachment; filename="ANTERIOR_{os.path.basename(ant_path)}"'
                    msg.attach(part_ant)

        server = smtplib.SMTP(smtp_conf['servidor'], smtp_conf['puerto'])
        server.starttls()
        server.login(smtp_conf['email_emisor'], smtp_conf['app_password'])
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        print(f"❌ Error enviando correo a {destinatario_str}: {e}")
        return False


def enviar_correo_no_responsive(marca_db, path_excel_mobile, path_excel_desktop, stats_mobile, stats_desktop, hist_ant_mobile, hist_ant_desktop, num_rev_mobile, num_rev_desktop, fecha_str):
    conf = obtener_configuracion()
    destinatarios_set = set()
    siempre = [
        marca_db.get('implementador_email'),
        marca_db.get('gerente_email'),
        marca_db.get('pm_email'),
        conf.get('lider_implementacion')
    ]
    for email in siempre:
        if email:
            destinatarios_set.update([e.strip() for e in email.replace(';', ',').split(',') if e.strip()])

    max_rev = max(num_rev_mobile, num_rev_desktop)
    if max_rev >= 3 and conf.get('lider_seo'):
        destinatarios_set.update([e.strip() for e in conf.get('lider_seo').replace(';', ',').split(',') if e.strip()])

    if max_rev >= 4 and conf.get('lider_operaciones'):
        destinatarios_set.update([e.strip() for e in conf.get('lider_operaciones').replace(';', ',').split(',') if e.strip()])

    destinatario_str = ", ".join(destinatarios_set)

    try:
        msg = MIMEMultipart('alternative')
        smtp_conf = obtener_config_smtp()
        msg['From'] = smtp_conf['email_emisor']
        msg['To'] = destinatario_str
        msg['Subject'] = f"{fecha_str} - {marca_db.get('marca', '')} - Revisión (Mobile: {num_rev_mobile} / Desk: {num_rev_desktop})"

        def _armar_filas_tabla(stats, ant):
            rows = [
                ("URLs procesadas", stats['total'], ant['urls_procesadas'] if ant else None),
                ("URLs con redirección", stats['redireccion'], ant['urls_redireccion'] if ant else None),
                ("URLs no creadas", stats['no_creadas'], ant['urls_no_creadas'] if ant else None),
                ("Meta título no coincide", stats['meta_titulo'], ant['urls_meta_titulo'] if ant else None),
                ("Meta descripción no coinc.", stats['meta_descripcion'], ant['urls_meta_descripcion'] if ant else None),
                ("H1 no coincide", stats['h1'], ant['urls_h1'] if ant else None),
                ("Desc. corta no coincide", stats['desc_corta'], ant['urls_desc_corta'] if ant else None),
                ("Desc. larga no coincide", stats['desc_larga'], ant['urls_desc_larga'] if ant else None)
            ]
            table_rows_html = ""
            for label, actual, prev in rows:
                delta = _delta_icon(prev, actual)
                prev_str = str(prev) if prev is not None else "–"
                table_rows_html += f"""
                <tr>
                  <td style='padding:8px 12px;border-bottom:1px solid #2a2a3a;text-align:left;'>{label}</td>
                  <td style='padding:8px 12px;border-bottom:1px solid #2a2a3a;text-align:center;font-weight:bold;'>{actual}</td>
                  <td style='padding:8px 12px;border-bottom:1px solid #2a2a3a;text-align:center;color:#888;'>{prev_str}</td>
                  <td style='padding:8px 12px;border-bottom:1px solid #2a2a3a;text-align:left;'>{delta}</td>
                </tr>"""
            return table_rows_html

        rows_mobile_html  = _armar_filas_tabla(stats_mobile, hist_ant_mobile)
        rows_desktop_html = _armar_filas_tabla(stats_desktop, hist_ant_desktop)

        revision_color_m = '#2ecc71' if num_rev_mobile == 1 else ('#f1c40f' if num_rev_mobile <= 3 else '#e74c3c')
        revision_color_d = '#2ecc71' if num_rev_desktop == 1 else ('#f1c40f' if num_rev_desktop <= 3 else '#e74c3c')

        ant_fecha_m = hist_ant_mobile['fecha_ejecucion'].strftime("%d/%m/%Y") if hist_ant_mobile else "-"
        ant_fecha_d = hist_ant_desktop['fecha_ejecucion'].strftime("%d/%m/%Y") if hist_ant_desktop else "-"

        body_html = f"""
        <!DOCTYPE html>
        <html>
        <body style='background:#0f0f1a;color:#e0e0e0;font-family:Arial,sans-serif;padding:30px;'>
          <div style='max-width:750px;margin:auto;background:#1a1a2e;border-radius:12px;padding:30px;'>
            <h2 style='color:#4fc3f7;margin-bottom:4px;'>Reporte de Auditoría SEO Consolidado (No Responsive)</h2>
            <h3 style='color:#e0e0e0;margin-top:0;font-weight:400;'>{marca_db.get('marca', '')}</h3>
            <p style='color:#aaa;margin-top:8px;'>Fecha de esta revisión: {fecha_str}</p>
            
            <p style='line-height:1.7;'>Este sitio está configurado como <strong>no responsive</strong>, por lo que se realizaron dos auditorías independientes: una simulación Mobile y otra Desktop.</p>
            
            <!-- SECCIÓN MOBILE -->
            <div style='margin-top:30px; border-left: 4px solid #3498db; padding-left: 15px; margin-bottom: 25px;'>
              <h3 style='color:#4fc3f7; margin-bottom: 5px;'>1. Resultados Móvil (Mobile)</h3>
              <p style='margin: 2px 0;'>Número de revisión: <span style='color:{revision_color_m}; font-weight:bold;'>{num_rev_mobile}</span></p>
              <p style='margin: 2px 0; color:#aaa; font-size:13px;'>Fecha revisión anterior: {ant_fecha_m}</p>
              <table style='width:100%;border-collapse:collapse;background:#12122a;border-radius:8px;overflow:hidden;margin-top:10px;'>
                <thead>
                  <tr style='background:#16213e;color:#4fc3f7;'>
                    <th style='padding:10px 12px;text-align:left;'>Parámetro (Mobile)</th>
                    <th style='padding:10px 12px;text-align:center;'>Actual #{num_rev_mobile}</th>
                    <th style='padding:10px 12px;text-align:center;'>Anterior</th>
                    <th style='padding:10px 12px;text-align:left;'>Comparación</th>
                  </tr>
                </thead>
                <tbody>
                  {rows_mobile_html}
                </tbody>
              </table>
            </div>

            <!-- SECCIÓN DESKTOP -->
            <div style='margin-top:30px; border-left: 4px solid #3498db; padding-left: 15px; margin-bottom: 25px;'>
              <h3 style='color:#4fc3f7; margin-bottom: 5px;'>2. Resultados Escritorio (Desktop)</h3>
              <p style='margin: 2px 0;'>Número de revisión: <span style='color:{revision_color_d}; font-weight:bold;'>{num_rev_desktop}</span></p>
              <p style='margin: 2px 0; color:#aaa; font-size:13px;'>Fecha revisión anterior: {ant_fecha_d}</p>
              <table style='width:100%;border-collapse:collapse;background:#12122a;border-radius:8px;overflow:hidden;margin-top:10px;'>
                <thead>
                  <tr style='background:#16213e;color:#4fc3f7;'>
                    <th style='padding:10px 12px;text-align:left;'>Parámetro (Desktop)</th>
                    <th style='padding:10px 12px;text-align:center;'>Actual #{num_rev_desktop}</th>
                    <th style='padding:10px 12px;text-align:center;'>Anterior</th>
                    <th style='padding:10px 12px;text-align:left;'>Comparación</th>
                  </tr>
                </thead>
                <tbody>
                  {rows_desktop_html}
                </tbody>
              </table>
            </div>

            <p style='margin-top:24px;color:#aaa;font-size:13px;'>Adjunto los archivos Excel correspondientes. Las celdas en rojo indican las diferencias encontradas con respecto al documento de contenidos.</p>
            <div style='margin-top:20px;padding:12px;background:#0a3d62;border-radius:8px;font-size:12px;color:#7fb3d3;'>Generado automáticamente por el Sistema SEOTracker.</div>
          </div>
        </body>
        </html>
        """

        msg.attach(MIMEText(body_html, 'html', 'utf-8'))

        if os.path.exists(path_excel_mobile):
            with open(path_excel_mobile, 'rb') as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(path_excel_mobile))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(path_excel_mobile)}"'
            msg.attach(part)

        if os.path.exists(path_excel_desktop):
            with open(path_excel_desktop, 'rb') as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(path_excel_desktop))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(path_excel_desktop)}"'
            msg.attach(part)

        if hist_ant_mobile and 'ruta_reporte_excel' in hist_ant_mobile:
            ant_path = os.path.abspath(hist_ant_mobile['ruta_reporte_excel'])
            if os.path.exists(ant_path) and ant_path != os.path.abspath(path_excel_mobile):
                with open(ant_path, 'rb') as f:
                    part = MIMEApplication(f.read(), Name=f"ANTERIOR_MOBILE - {os.path.basename(ant_path)}")
                part['Content-Disposition'] = f'attachment; filename="ANTERIOR_MOBILE_{os.path.basename(ant_path)}"'
                msg.attach(part)

        if hist_ant_desktop and 'ruta_reporte_excel' in hist_ant_desktop:
            ant_path = os.path.abspath(hist_ant_desktop['ruta_reporte_excel'])
            if os.path.exists(ant_path) and ant_path != os.path.abspath(path_excel_desktop):
                with open(ant_path, 'rb') as f:
                    part = MIMEApplication(f.read(), Name=f"ANTERIOR_DESKTOP - {os.path.basename(ant_path)}")
                part['Content-Disposition'] = f'attachment; filename="ANTERIOR_DESKTOP_{os.path.basename(ant_path)}"'
                msg.attach(part)

        server = smtplib.SMTP(smtp_conf['servidor'], smtp_conf['puerto'])
        server.starttls()
        server.login(smtp_conf['email_emisor'], smtp_conf['app_password'])
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        print(f"❌ Error enviando correo no responsive a {destinatario_str}: {e}")
        return False


# --- auditoría ---

def _ejecutar_auditoria_core(marca_db, user_agent=None, sufijo_tipo="", num_rev=None, progress_callback=None, df_contenido=None):
    marca_id  = marca_db['id']
    marca_nom = marca_db['marca']

    if df_contenido is None:
        df_contenido = descargar_hoja_sheets(marca_db['doc_contenido'])
    else:
        df_contenido = df_contenido.copy()

    if df_contenido.empty or 'URL' not in df_contenido.columns:
        print(f"[{marca_nom}] No se pudo obtener el documento de contenido.")
        return None

    urls  = df_contenido['URL'].dropna().astype(str).str.strip().unique().tolist()
    total = len(urls)

    results = []
    session = requests.Session()
    with ThreadPoolExecutor(max_workers=1) as ex:
        futs = {ex.submit(fetch_url, url, marca_db, user_agent, session): url for url in urls}
        done = 0
        for fut in futs:
            results.append(fut.result())
            done += 1
            if progress_callback:
                progress_callback(done, total, fut.result().get('Dirección', ''))

    df_live = pd.DataFrame(results)
    df_contenido['URL_clean'] = df_contenido['URL'].astype(str).str.strip()
    df_live['Dir_clean']      = df_live['Dirección'].astype(str).str.strip()
    df = pd.merge(df_contenido, df_live, left_on='URL_clean', right_on='Dir_clean', how='left')

    MAPA_COLUMNAS = {
        'meta titulo':        'Meta Título',
        'meta título':        'Meta Título',
        'meta-titulo':        'Meta Título',
        'meta-título':        'Meta Título',
        'meta title':         'Meta Título',
        'titulo':             'Meta Título',
        'title':              'Meta Título',
        'meta descripcion':   'Meta Descripción',
        'meta descripción':   'Meta Descripción',
        'meta description':   'Meta Descripción',
        'meta desc':          'Meta Descripción',
        'descripcion':        'Meta Descripción',
        'h1':                 'H1',
        'titulo h1':          'H1',
        'h1 titulo':          'H1',
        'h1 título':          'H1',
        'h2':                 'H2',
        'h2:':                'H2',
        'titulo h2':          'H2',
        'h2 titulo':          'H2',
        'h2 título':          'H2',
        'subtitulo':          'H2',
        'subtítulo':          'H2',
        'h3':                 'H3',
        'h3:':                'H3',
        'titulo h3':          'H3',
        'h3 titulo':          'H3',
        'h3 título':          'H3',
        'descripcion corta':  'Descripción Corta',
        'descripción corta':  'Descripción Corta',
        'short description':  'Descripción Corta',
        'desc corta':         'Descripción Corta',
        'descripcion larga':  'Descripción Larga',
        'descripción larga':  'Descripción Larga',
        'long description':   'Descripción Larga',
        'desc larga':         'Descripción Larga',
        'url':                'URL',
    }
    renombrar = {}
    for col_real in df.columns:
        norm = str(col_real).strip().lower()
        if norm in MAPA_COLUMNAS and col_real != MAPA_COLUMNAS[norm]:
            renombrar[col_real] = MAPA_COLUMNAS[norm]

    if renombrar:
        df = df.rename(columns=renombrar)

    df = df.loc[:, ~df.columns.duplicated()].copy()

    if 'URL' in df.columns:
        df['URL'] = df['URL'].astype(str).str.strip()
        df = df[df['URL'] != '']
        df = df[df['URL'].str.lower() != 'nan']
    else:
        print(f"[{marca_nom}] ⚠ No se encontró la columna 'URL'. Es posible que el mapeo haya fallado.")

    if df.empty:
        print(f"[{marca_nom}] El documento quedó vacío tras filtrar URLs.")
        return None

    for col in ['Meta Título', 'Meta Descripción', 'H1', 'H2', 'H3', 'Descripción Corta', 'Descripción Larga']:
        if col in df.columns:
            df[col] = df[col].apply(limpiar_html)

    for col in ['Título 1', 'Meta Description 1', 'H1-1', 'Descripción corta 1', 'Descripción larga 1']:
        if col in df.columns:
            df[col] = df[col].apply(limpiar_html)

    # normalizar separador de título (guiones → pipe) para comparación consistente
    for col in ['Meta Título', 'Título 1']:
        if col in df.columns:
            df[col] = df[col].str.replace('-', '|', regex=False)

    palabra = marca_db.get('palabra_excluir', '').strip()
    if palabra and 'Título 1' in df.columns:
        regex_excluir = re.compile(re.escape(palabra), re.IGNORECASE)
        df['Título 1'] = df['Título 1'].apply(lambda x: regex_excluir.sub('', str(x)).strip())

    def coinc(a, b):
        if a not in df.columns:
            return pd.Series(['Sin datos'] * len(df))
        if b not in df.columns:
            return pd.Series(['Sin datos'] * len(df))

        # Código_Final si existe; si no, parsear Código de respuesta
        def get_code(row_idx):
            if 'Código_Final' in df.columns:
                v = df['Código_Final'].iloc[row_idx]
                if pd.notna(v) and v is not None:
                    try: return int(float(v))
                    except: pass
            v2 = df['Código de respuesta'].iloc[row_idx]
            if pd.notna(v2):
                s = str(v2).split('->')[0].strip()
                if s.replace('.', '').isdigit():
                    return int(float(s))
            return 0

        serie_a = normalizar(df[a])
        serie_b = normalizar(df[b])
        matches = (serie_a == serie_b)

        res = []
        for i in range(len(df)):
            code = get_code(i)
            # code==0 → no se pudo leer el código, comparar de todas formas
            # code==200 → OK; cualquier otro código HTTP real → error
            if code not in (0, 200):
                res.append(f"No (HTTP {code})")
            else:
                val_a_norm = unicodedata.normalize('NFKC', str(df[a].iloc[i])).strip().lower()
                val_b_norm = unicodedata.normalize('NFKC', str(df[b].iloc[i])).strip().lower()

                is_a_empty = val_a_norm in ('', 'nan', 'vacío', 'vacio', 'none')
                is_b_empty = val_b_norm in ('', 'nan', 'vacío', 'vacio', 'none')

                if is_a_empty and is_b_empty:
                    res.append('Sí')
                elif is_a_empty:
                    res.append("No (No en documento)")
                else:
                    res.append('Sí' if matches.iloc[i] else 'No')

        return pd.Series(res)

    df['meta_titulo_coincide']        = coinc('Meta Título',       'Título 1')
    df['meta_descripcion_coincide']   = coinc('Meta Descripción',  'Meta Description 1')
    df['h1_coincide']                 = coinc('H1',                'H1-1')
    df['desc_corta_coincide']         = coinc('Descripción Corta', 'Descripción corta 1')
    df['desc_larga_coincide']         = coinc('Descripción Larga', 'Descripción larga 1')

    df['h1_dup_desc_larga'] = df.apply(
        lambda r: detectar_h1_duplicado_en_desc_larga(
            r.get('H1-1', ''),
            r.get('Descripción larga 1', '')
        ), axis=1
    )

    def no_count(col): return df[col].astype(str).str.startswith('No').sum()

    codigo_orig  = df['Código_Original'].fillna(0).astype(float).astype(int) if 'Código_Original' in df.columns else pd.Series([0] * len(df))
    codigo_final = df['Código_Final'].fillna(0).astype(float).astype(int) if 'Código_Final' in df.columns else pd.Series([0] * len(df))

    stats = {
        'total':            total,
        'redireccion':      int(((codigo_orig >= 300) & (codigo_orig <= 399)).sum()),
        'no_creadas':       int(((codigo_final == 404) | df['Código de respuesta'].astype(str).str.strip().isin(['', 'nan'])).sum()),
        'meta_titulo':      int(no_count('meta_titulo_coincide')),
        'meta_descripcion': int(no_count('meta_descripcion_coincide')),
        'h1':               int(no_count('h1_coincide')),
        'desc_corta':       int(no_count('desc_corta_coincide')),
        'desc_larga':       int(no_count('desc_larga_coincide')),
    }

    nuevo_nombre_columnas = {
        'Meta Título': 'Meta Título (Hoja de Contenido)',
        'Título 1': 'Meta Titulo (Rastreo en Vivo)',
        'Meta Descripción': 'Meta Descripción (Hoja de Contenido)',
        'Meta Description 1': 'Meta Descripción (Rastreo en Vivo)',
        'H1': 'H1 (Hoja de Contenido)',
        'H1-1': 'H1 (Rastreo en Vivo)',
        'H2': 'H2 (Hoja de Contenido)',
        'H3': 'H3 (Hoja de Contenido)',
        'Descripción Corta': 'Descripción Corta (Hoja de Contenido)',
        'Descripción corta 1': 'Descripción Corta (Rastreo en Vivo)',
        'Descripción Larga': 'Descripción Larga (Hoja de Contenido)',
        'Descripción larga 1': 'Descripción Larga (Rastreo en Vivo)',
        'meta_titulo_coincide': 'META_TITULO_COINCIDE',
        'meta_descripcion_coincide': 'META_DESCRIPCION_COINCIDE',
        'h1_coincide': 'H1_COINCIDE',
        'desc_corta_coincide': 'DESCRIPCION_CORTA_COINCIDE',
        'desc_larga_coincide': 'DESCRIPCION_LARGA_COINCIDE',
        'CONTENIDO_SF_H2_H3': 'Títulos H2-H3 (Rastreo en Vivo)',
        'h1_dup_desc_larga': 'H1 DUPLICADO EN DESC. LARGA'
    }

    columnas_out = ['URL', 'Código de respuesta',
        'Meta Título', 'meta_titulo_coincide', 'Título 1',
        'Meta Descripción', 'meta_descripcion_coincide', 'Meta Description 1',
        'H1', 'h1_coincide', 'H1-1',
        'H2', 'H3',
        'Descripción Corta', 'desc_corta_coincide', 'Descripción corta 1',
        'Descripción Larga', 'desc_larga_coincide', 'Descripción larga 1',
        'CONTENIDO_SF_H2_H3', 'h1_dup_desc_larga']

    df_final = df[[c for c in columnas_out if c in df.columns]].rename(columns=nuevo_nombre_columnas)

    ahora     = datetime.datetime.now()
    fecha_str = f"{ahora.day} {MESES_ES[ahora.month]} {ahora.year}"
    if num_rev is None:
        num_rev = siguiente_revision(marca_id)

    sufijo = f" - {sufijo_tipo}" if sufijo_tipo else ""
    nombre_archivo = f"{fecha_str} - {marca_nom}{sufijo} - Revision {num_rev}.xlsx"
    os.makedirs("Reportes", exist_ok=True)
    path_excel = os.path.abspath(os.path.join("Reportes", nombre_archivo))

    df_final.to_excel(path_excel, index=False)
    aplicar_formato_excel(path_excel)

    if sufijo_tipo:
        historico_ant = None
        cnx = get_conn()
        cur = cnx.cursor(dictionary=True)
        patron = f"%{sufijo_tipo}%"
        cur.execute("""
            SELECT * FROM historico 
            WHERE marca_id = %s AND (ruta_reporte_excel LIKE %s OR ruta_reporte_excel LIKE %s)
            ORDER BY fecha_ejecucion DESC LIMIT 1
        """, (marca_id, patron, patron.lower()))
        historico_ant = cur.fetchone()
        cur.close(); cnx.close()
    else:
        historico_ant = obtener_ultimo_historico(marca_id)

    arreglados = 0
    cambiados  = 0
    if historico_ant:
        errores_prev = (historico_ant.get('urls_redireccion', 0) + historico_ant.get('urls_no_creadas', 0) +
                        historico_ant.get('urls_meta_titulo', 0) + historico_ant.get('urls_meta_descripcion', 0) +
                        historico_ant.get('urls_h1', 0) + historico_ant.get('urls_desc_corta', 0) +
                        historico_ant.get('urls_desc_larga', 0))
        errores_curr = (stats['redireccion'] + stats['no_creadas'] +
                        stats['meta_titulo'] + stats['meta_descripcion'] +
                        stats['h1'] + stats['desc_corta'] + stats['desc_larga'])

        arreglados, cambiados = calcular_cambios_contenido(
            historico_ant.get('ruta_reporte_excel'),
            path_excel,
            errores_prev,
            errores_curr
        )

    guardar_historico(marca_id, stats, path_excel, num_rev=num_rev, arreglados=arreglados, cambiados=cambiados)

    return {
        'stats': stats,
        'path_excel': path_excel,
        'num_rev': num_rev,
        'historico_ant': historico_ant,
        'fecha_str': fecha_str
    }


def auditar_marca(marca_db, progress_callback=None):
    marca_id      = marca_db['id']
    marca_nom     = marca_db['marca']
    is_responsive = marca_db.get('responsive', True)

    if is_responsive:
        res = _ejecutar_auditoria_core(
            marca_db,
            user_agent=None,
            sufijo_tipo="",
            progress_callback=progress_callback
        )
        if not res:
            return None

        enviar_correo(
            marca_db           = marca_db,
            path_excel         = res['path_excel'],
            stats              = res['stats'],
            historico_anterior = res['historico_ant'],
            num_revision       = res['num_rev'],
            fecha_str          = res['fecha_str']
        )

        return res['path_excel']
    else:
        num_rev = siguiente_revision(marca_id)

        df_contenido_shared = descargar_hoja_sheets(marca_db['doc_contenido'])
        if df_contenido_shared.empty or 'URL' not in df_contenido_shared.columns:
            return None

        res_mobile = _ejecutar_auditoria_core(
            marca_db,
            user_agent=UA_MOBILE,
            sufijo_tipo="Mobile",
            num_rev=num_rev,
            progress_callback=progress_callback,
            df_contenido=df_contenido_shared
        )

        res_desktop = _ejecutar_auditoria_core(
            marca_db,
            user_agent=UA_DESKTOP,
            sufijo_tipo="Desktop",
            num_rev=num_rev,
            progress_callback=progress_callback,
            df_contenido=df_contenido_shared
        )

        if not res_mobile or not res_desktop:
            return None

        enviar_correo_no_responsive(
            marca_db            = marca_db,
            path_excel_mobile   = res_mobile['path_excel'],
            path_excel_desktop  = res_desktop['path_excel'],
            stats_mobile        = res_mobile['stats'],
            stats_desktop       = res_desktop['stats'],
            hist_ant_mobile     = res_mobile['historico_ant'],
            hist_ant_desktop    = res_desktop['historico_ant'],
            num_rev_mobile      = res_mobile['num_rev'],
            num_rev_desktop     = res_desktop['num_rev'],
            fecha_str           = res_mobile['fecha_str']
        )

        return res_mobile['path_excel']


def _auditar_lista_un_modo(marca_db, urls_manuales, df_filtrado, user_agent, sufijo_archivo):
    marca_nom = marca_db['marca']

    results = []
    session = requests.Session()
    with ThreadPoolExecutor(max_workers=1) as ex:
        futs = {ex.submit(fetch_url, url, marca_db, user_agent, session): url for url in urls_manuales}
        for fut in futs:
            results.append(fut.result())

    df_live = pd.DataFrame(results)
    df_live['Dir_clean'] = df_live['Dirección'].astype(str).str.strip().str.lower()

    df = pd.merge(df_live, df_filtrado.copy(), left_on='Dir_clean', right_on='URL_norm', how='left')

    MAPA_COLUMNAS = {
        'meta titulo': 'Meta Título', 'meta título': 'Meta Título', 'meta-titulo': 'Meta Título',
        'meta-título': 'Meta Título', 'meta title': 'Meta Título',
        'meta descripcion': 'Meta Descripción', 'meta descripción': 'Meta Descripción',
        'meta description': 'Meta Descripción',
        'h1': 'H1', 'h2': 'H2', 'h3': 'H3',
        'descripcion corta': 'Descripción Corta', 'short description': 'Descripción Corta',
        'descripcion larga': 'Descripción Larga', 'long description': 'Descripción Larga',
        'url': 'URL'
    }
    renombrar = {}
    for col_real in df.columns:
        norm = str(col_real).strip().lower()
        if norm in MAPA_COLUMNAS and col_real != MAPA_COLUMNAS[norm]:
            renombrar[col_real] = MAPA_COLUMNAS[norm]
    if renombrar:
        df = df.rename(columns=renombrar)

    df = df.loc[:, ~df.columns.duplicated()].copy()
    for col in ['Meta Título', 'Meta Descripción', 'H1', 'H2', 'H3', 'Descripción Corta', 'Descripción Larga',
                'Título 1', 'Meta Description 1', 'H1-1', 'Descripción corta 1', 'Descripción larga 1']:
        if col in df.columns:
            df[col] = df[col].apply(limpiar_html)

    palabra = marca_db.get('palabra_excluir', '').strip()
    if palabra and 'Título 1' in df.columns:
        regex_excluir = re.compile(re.escape(palabra), re.IGNORECASE)
        df['Título 1'] = df['Título 1'].apply(lambda x: regex_excluir.sub('', str(x)).strip())

    def coinc(a, b):
        if a not in df.columns or b not in df.columns:
            return pd.Series(['Sin datos'] * len(df))

        def get_code(row_idx):
            if 'Código_Final' in df.columns:
                v = df['Código_Final'].iloc[row_idx]
                if pd.notna(v) and v is not None:
                    try: return int(float(v))
                    except: pass
            v2 = df['Código de respuesta'].iloc[row_idx]
            if pd.notna(v2):
                s = str(v2).split('->')[0].strip()
                if s.replace('.', '').isdigit():
                    return int(float(s))
            return 0

        serie_a = normalizar(df[a])
        serie_b = normalizar(df[b])
        matches = (serie_a == serie_b)
        res = []
        for i in range(len(df)):
            code = get_code(i)
            val_a = str(df[a].iloc[i]).strip().lower()
            val_b = str(df[b].iloc[i]).strip().lower()

            is_a_empty = val_a in ('', 'nan', 'vacío', 'vacio', 'none')
            is_b_empty = val_b in ('', 'nan', 'vacío', 'vacio', 'none')

            if code not in (0, 200):
                res.append(f"No (HTTP {code})")
            elif is_a_empty and is_b_empty:
                res.append('Sí')
            elif is_a_empty:
                res.append("No (No en documento)")
            else:
                res.append('Sí' if matches.iloc[i] else 'No')
        return pd.Series(res)

    df['meta_titulo_coincide']      = coinc('Meta Título',       'Título 1')
    df['meta_descripcion_coincide'] = coinc('Meta Descripción',  'Meta Description 1')
    df['h1_coincide']               = coinc('H1',                'H1-1')
    df['desc_corta_coincide']       = coinc('Descripción Corta', 'Descripción corta 1')
    df['desc_larga_coincide']       = coinc('Descripción Larga', 'Descripción larga 1')

    df['h1_dup_desc_larga'] = df.apply(
        lambda r: detectar_h1_duplicado_en_desc_larga(
            r.get('H1-1', ''),
            r.get('Descripción larga 1', '')
        ), axis=1
    )

    stats = {
        'total': len(urls_manuales),
        'meta_titulo':      int(df['meta_titulo_coincide'].astype(str).str.startswith('No').sum()),
        'meta_descripcion': int(df['meta_descripcion_coincide'].astype(str).str.startswith('No').sum()),
        'h1':               int(df['h1_coincide'].astype(str).str.startswith('No').sum()),
        'desc_corta':       int(df['desc_corta_coincide'].astype(str).str.startswith('No').sum()),
        'desc_larga':       int(df['desc_larga_coincide'].astype(str).str.startswith('No').sum()),
    }

    ahora = datetime.datetime.now()
    nombre_archivo = f"LISTA_{ahora.strftime('%Y%m%d_%H%M%S')}_{marca_nom}_{sufijo_archivo}.xlsx"
    os.makedirs("Reportes", exist_ok=True)
    path_excel = os.path.abspath(os.path.join("Reportes", nombre_archivo))

    nuevo_nombre_columnas = {
        'Meta Título': 'Meta Título (Hoja de Contenido)',        'Título 1': 'Meta Titulo (Rastreo en Vivo)',
        'Meta Descripción': 'Meta Descripción (Hoja de Contenido)', 'Meta Description 1': 'Meta Descripción (Rastreo en Vivo)',
        'H1': 'H1 (Hoja de Contenido)',                          'H1-1': 'H1 (Rastreo en Vivo)',
        'H2': 'H2 (Hoja de Contenido)',
        'H3': 'H3 (Hoja de Contenido)',
        'Descripción Corta': 'Descripción Corta (Hoja de Contenido)', 'Descripción corta 1': 'Descripción Corta (Rastreo en Vivo)',
        'Descripción Larga': 'Descripción Larga (Hoja de Contenido)', 'Descripción larga 1': 'Descripción Larga (Rastreo en Vivo)',
        'meta_titulo_coincide': 'META_TITULO_COINCIDE',
        'meta_descripcion_coincide': 'META_DESCRIPCION_COINCIDE',
        'h1_coincide': 'H1_COINCIDE',
        'desc_corta_coincide': 'DESCRIPCION_CORTA_COINCIDE',
        'desc_larga_coincide': 'DESCRIPCION_LARGA_COINCIDE',
        'CONTENIDO_SF_H2_H3': 'Títulos H2-H3 (Rastreo en Vivo)',
        'h1_dup_desc_larga': 'H1 DUPLICADO EN DESC. LARGA'
    }
    cols_out = ['Dirección', 'Código de respuesta',
                'Meta Título', 'meta_titulo_coincide', 'Título 1',
                'Meta Descripción', 'meta_descripcion_coincide', 'Meta Description 1',
                'H1', 'h1_coincide', 'H1-1',
                'H2', 'H3',
                'Descripción Corta', 'desc_corta_coincide', 'Descripción corta 1',
                'Descripción Larga', 'desc_larga_coincide', 'Descripción larga 1',
                'CONTENIDO_SF_H2_H3', 'h1_dup_desc_larga']
    df_final = df[[c for c in cols_out if c in df.columns]].rename(columns=nuevo_nombre_columnas)
    df_final.to_excel(path_excel, index=False)
    aplicar_formato_excel(path_excel)

    results_list = []
    for _, row in df.iterrows():
        results_list.append({
            'url':                         row.get('Dirección', ''),
            'status_code':                 row.get('Código de respuesta', ''),
            'meta_titulo_coincide':        row.get('meta_titulo_coincide', 'Sin datos'),
            'meta_titulo_esperado':        row.get('Meta Título', ''),
            'meta_titulo_encontrado':      row.get('Título 1', ''),
            'meta_descripcion_coincide':   row.get('meta_descripcion_coincide', 'Sin datos'),
            'meta_descripcion_esperado':   row.get('Meta Descripción', ''),
            'meta_descripcion_encontrado': row.get('Meta Description 1', ''),
            'h1_coincide':                 row.get('h1_coincide', 'Sin datos'),
            'h1_esperado':                 row.get('H1', ''),
            'h1_encontrado':               row.get('H1-1', ''),
            'desc_corta_coincide':         row.get('desc_corta_coincide', 'Sin datos'),
            'desc_corta_esperado':         row.get('Descripción Corta', ''),
            'desc_corta_encontrado':       row.get('Descripción corta 1', ''),
            'desc_larga_coincide':         row.get('desc_larga_coincide', 'Sin datos'),
            'desc_larga_esperado':         row.get('Descripción Larga', ''),
            'desc_larga_encontrado':       row.get('Descripción larga 1', '')
        })

    return {
        'ok': True,
        'stats': stats,
        'results': results_list,
        'excel_path': f"Reportes/{nombre_archivo}"
    }


def auditar_lista_manual(marca_db, urls_manuales):
    urls_manuales = [u.strip() for u in urls_manuales if u.strip()]

    df_contenido = descargar_hoja_sheets(marca_db['doc_contenido'])
    if df_contenido.empty or 'URL' not in df_contenido.columns:
        return {'ok': False, 'error': 'No se pudo obtener el documento de contenido o no tiene columna URL.'}

    df_contenido['URL_norm'] = df_contenido['URL'].astype(str).str.strip().str.lower()
    urls_norm = [u.lower() for u in urls_manuales]
    df_filtrado = df_contenido[df_contenido['URL_norm'].isin(urls_norm)].copy()

    es_responsive = bool(marca_db.get('responsive', True))

    if es_responsive:
        res = _auditar_lista_un_modo(marca_db, urls_manuales, df_filtrado,
                                     user_agent=None, sufijo_archivo='Responsive')
        return res
    else:
        resultado_mobile  = [None]
        resultado_desktop = [None]

        def _run_mobile():
            resultado_mobile[0] = _auditar_lista_un_modo(
                marca_db, urls_manuales, df_filtrado,
                user_agent=UA_MOBILE, sufijo_archivo='Mobile'
            )

        def _run_desktop():
            resultado_desktop[0] = _auditar_lista_un_modo(
                marca_db, urls_manuales, df_filtrado,
                user_agent=UA_DESKTOP, sufijo_archivo='Desktop'
            )

        t_mobile  = threading.Thread(target=_run_mobile)
        t_desktop = threading.Thread(target=_run_desktop)
        t_mobile.start()
        t_desktop.start()
        t_mobile.join()
        t_desktop.join()

        res_m = resultado_mobile[0]
        res_d = resultado_desktop[0]

        if not res_m or not res_m.get('ok'):
            return {'ok': False, 'error': f"Error en rastreo Mobile: {res_m.get('error', 'desconocido') if res_m else 'sin resultado'}"}
        if not res_d or not res_d.get('ok'):
            return {'ok': False, 'error': f"Error en rastreo Desktop: {res_d.get('error', 'desconocido') if res_d else 'sin resultado'}"}

        return {
            'ok': True,
            'no_responsive': True,
            'mobile': {
                'stats':      res_m['stats'],
                'results':    res_m['results'],
                'excel_path': res_m['excel_path'],
            },
            'desktop': {
                'stats':      res_d['stats'],
                'results':    res_d['results'],
                'excel_path': res_d['excel_path'],
            }
        }


def enviar_correo_lista(marca_db, path_excel, stats, destinatarios=None, path_excel_desktop=None, stats_desktop=None):
    if destinatarios:
        destinatario_str = ", ".join([e.strip() for e in destinatarios if e.strip()])
    else:
        conf = obtener_configuracion()
        destinatarios_set = set()
        siempre = [
            marca_db.get('implementador_email'),
            marca_db.get('gerente_email'),
            marca_db.get('pm_email'),
            conf.get('lider_implementacion'),
            conf.get('lider_seo'),
            conf.get('lider_operaciones')
        ]
        for email in siempre:
            if email:
                destinatarios_set.update([e.strip() for e in email.replace(';', ',').split(',') if e.strip()])
        destinatario_str = ", ".join(destinatarios_set)

    es_dual = bool(path_excel_desktop and stats_desktop)

    try:
        msg = MIMEMultipart('alternative')
        smtp_conf = obtener_config_smtp()
        msg['From'] = smtp_conf['email_emisor']
        msg['To'] = destinatario_str
        fecha_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")

        if es_dual:
            msg['Subject'] = f"{fecha_str} - Rastreo Manual Consolidado (No Responsive) - {marca_db.get('marca', '')}"
            body_html = f"""
            <!DOCTYPE html>
            <html>
            <body style='background:#0f0f1a;color:#e0e0e0;font-family:Arial,sans-serif;padding:30px;'>
              <div style='max-width:720px;margin:auto;background:#1a1a2e;border-radius:12px;padding:30px;'>
                <h2 style='color:#4fc3f7;margin-bottom:4px;'>Reporte de Rastreo Manual (No Responsive)</h2>
                <h3 style='color:#e0e0e0;margin-top:0;font-weight:400;'>{marca_db.get('marca', '')}</h3>
                <p style='color:#aaa;margin-top:8px;'>Fecha: {fecha_str}</p>
                <p style='color:#aaa;'>Este sitio no es responsive; se ejecutaron dos rastreos en paralelo (Mobile y Desktop).</p>
                
                <div style='margin-top:20px; border-left: 4px solid #818cf8; padding-left: 15px; margin-bottom: 20px;'>
                  <h3 style='color:#818cf8; margin-bottom: 4px;'>1. Resultados Mobile ({stats.get('total', 0)} URLs)</h3>
                  <ul style='line-height:1.6;'>
                    <li>Meta Título: {stats.get('meta_titulo', 0)} errores</li>
                    <li>Meta Descripción: {stats.get('meta_descripcion', 0)} errores</li>
                    <li>H1: {stats.get('h1', 0)} errores</li>
                    <li>Descripción Corta: {stats.get('desc_corta', 0)} errores</li>
                    <li>Descripción Larga: {stats.get('desc_larga', 0)} errores</li>
                  </ul>
                </div>

                <div style='margin-top:20px; border-left: 4px solid #34d399; padding-left: 15px;'>
                  <h3 style='color:#34d399; margin-bottom: 4px;'>2. Resultados Desktop ({stats_desktop.get('total', 0)} URLs)</h3>
                  <ul style='line-height:1.6;'>
                    <li>Meta Título: {stats_desktop.get('meta_titulo', 0)} errores</li>
                    <li>Meta Descripción: {stats_desktop.get('meta_descripcion', 0)} errores</li>
                    <li>H1: {stats_desktop.get('h1', 0)} errores</li>
                    <li>Descripción Corta: {stats_desktop.get('desc_corta', 0)} errores</li>
                    <li>Descripción Larga: {stats_desktop.get('desc_larga', 0)} errores</li>
                  </ul>
                </div>

                <p style='margin-top:24px;color:#aaa;font-size:13px;'>Se adjuntan ambos archivos Excel con el detalle completo de cada versión.</p>
                <div style='margin-top:20px;padding:12px;background:#0a3d62;border-radius:8px;font-size:12px;color:#7fb3d3;'>Generado automáticamente por el Sistema SEOTracker.</div>
              </div>
            </body>
            </html>
            """
        else:
            msg['Subject'] = f"{fecha_str} - Rastreo Manual (Lista) - {marca_db.get('marca', '')}"
            errores_totales = sum(stats.get(k, 0) for k in ['meta_titulo', 'meta_descripcion', 'h1', 'desc_corta', 'desc_larga'])

            if errores_totales == 0:
                body_html = f"""
                <!DOCTYPE html>
                <html>
                <body style='background:#0f0f1a;color:#e0e0e0;font-family:Arial,sans-serif;padding:30px;'>
                  <div style='max-width:680px;margin:auto;background:#1a1a2e;border-radius:12px;padding:30px;text-align:center;'>
                    <h2 style='color:#2ecc71;margin-bottom:4px;'>Rastreo Manual Exitoso</h2>
                    <h3 style='color:#e0e0e0;margin-top:0;font-weight:400;'>{marca_db.get('marca', '')}</h3>
                    <p style='color:#aaa;margin-top:8px;'>Fecha: {fecha_str}</p>
                    <div style='margin-top:20px;padding:20px;background:#16213e;border-radius:8px;'>
                      <h1 style='color:#4fc3f7;font-size:48px;margin:0;'>100%</h1>
                      <p style='font-size:18px;color:#e0e0e0;margin-top:10px;'>Todas las URLs coinciden perfectamente.</p>
                      <p style='color:#aaa;'>Se procesaron {stats.get('total', 0)} URLs sin encontrar errores.</p>
                    </div>
                  </div>
                </body>
                </html>
                """
            else:
                body_html = f"""
                <!DOCTYPE html>
                <html>
                <body style='background:#0f0f1a;color:#e0e0e0;font-family:Arial,sans-serif;padding:30px;'>
                  <div style='max-width:680px;margin:auto;background:#1a1a2e;border-radius:12px;padding:30px;'>
                    <h2 style='color:#4fc3f7;margin-bottom:4px;'>Reporte de Rastreo Manual (Lista)</h2>
                    <h3 style='color:#e0e0e0;margin-top:0;font-weight:400;'>{marca_db.get('marca', '')}</h3>
                    <p style='color:#aaa;margin-top:8px;'>Fecha: {fecha_str}</p>
                    <p style='color:#aaa;margin-top:8px;'>Se procesaron {stats.get('total', 0)} URLs.</p>
                    <p>Errores detectados:</p>
                    <ul>
                      <li>Meta Título: {stats.get('meta_titulo', 0)}</li>
                      <li>Meta Descripción: {stats.get('meta_descripcion', 0)}</li>
                      <li>H1: {stats.get('h1', 0)}</li>
                      <li>Descripción Corta: {stats.get('desc_corta', 0)}</li>
                      <li>Descripción Larga: {stats.get('desc_larga', 0)}</li>
                    </ul>
                    <p style='margin-top:24px;color:#aaa;font-size:13px;'>Se adjunta el Excel con los resultados detallados.</p>
                  </div>
                </body>
                </html>
                """

        msg.attach(MIMEText(body_html, 'html', 'utf-8'))

        if path_excel:
            path_abs = os.path.abspath(path_excel)
            if os.path.exists(path_abs):
                with open(path_abs, 'rb') as f:
                    part = MIMEApplication(f.read(), Name=os.path.basename(path_abs))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(path_abs)}"'
                msg.attach(part)

        if path_excel_desktop:
            path_abs_d = os.path.abspath(path_excel_desktop)
            if os.path.exists(path_abs_d):
                with open(path_abs_d, 'rb') as f:
                    part_d = MIMEApplication(f.read(), Name=os.path.basename(path_abs_d))
                part_d['Content-Disposition'] = f'attachment; filename="{os.path.basename(path_abs_d)}"'
                msg.attach(part_d)

        server = smtplib.SMTP(smtp_conf['servidor'], smtp_conf['puerto'])
        server.starttls()
        server.login(smtp_conf['email_emisor'], smtp_conf['app_password'])
        server.send_message(msg)
        server.quit()
        print(f"Correo enviado en modo lista a {destinatario_str}")
        return True
    except Exception as e:
        print(f"Error enviando correo de lista manual: {e}")
        return False


# --- scheduler ---

def _trabajo_programado():
    """Revisa todas las marcas y lanza auditoría para las que llevan 15+ días sin rastrear."""
    print("Revisando auditorias programadas...", flush=True)
    for m in obtener_marcas():
        try:
            marca_nombre = str(m.get('marca', '')).encode('ascii', 'ignore').decode('ascii')
        except Exception:
            marca_nombre = "marca_desconocida"
        ult = obtener_ultimo_historico(m['id'])
        if ult is None:
            print(f"  >> {marca_nombre}: nunca auditada. Iniciando...", flush=True)
            threading.Thread(target=auditar_marca, args=(m,), daemon=True).start()
        else:
            dias = (datetime.datetime.now() - ult['fecha_ejecucion']).days
            if dias >= 15:
                print(f"  >> {marca_nombre}: {dias} dias desde ultima auditoria. Iniciando...", flush=True)
                threading.Thread(target=auditar_marca, args=(m,), daemon=True).start()
            else:
                print(f"  OK {marca_nombre}: ultima auditoria hace {dias} dia(s). Sin accion.", flush=True)


def iniciar_scheduler():
    # revisión inmediata al arrancar para no esperar al primer cron de las 2 AM
    def _arranque():
        time.sleep(5)  # dejar que Flask termine de levantar
        print("[Scheduler] Verificando auditorias pendientes al arranque...", flush=True)
        _trabajo_programado()

    threading.Thread(target=_arranque, daemon=True).start()

    schedule.every().day.at("02:00").do(_trabajo_programado)

    def _loop():
        while True:
            schedule.run_pending()
            time.sleep(60)

    t = threading.Thread(target=_loop, daemon=True)
    t.start()
    print("Scheduler iniciado: verificacion inmediata + auditorias diarias a las 02:00 AM.", flush=True)
