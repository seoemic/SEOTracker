import mysql.connector
from mysql.connector import errorcode
from werkzeug.security import generate_password_hash

DB_NAME = 'seo_auditor'

TABLES = {}

# Tabla que almacena a los miembros del equipo y sus roles
TABLES['usuarios'] = (
    "CREATE TABLE `usuarios` ("
    "  `id` int(11) NOT NULL AUTO_INCREMENT,"
    "  `username` varchar(100) NOT NULL UNIQUE,"
    "  `email` varchar(255) NOT NULL UNIQUE,"
    "  `password_hash` varchar(512) NOT NULL,"
    "  `rol` ENUM('lider','pm','gestor_cuenta','implementador') NOT NULL DEFAULT 'lider',"
    "  `activo` boolean DEFAULT True,"
    "  `creado_en` datetime DEFAULT CURRENT_TIMESTAMP,"
    "  PRIMARY KEY (`id`)"
    ") ENGINE=InnoDB"
)

# Tabla que agrupa a los miembros del equipo por células operativas
TABLES['celulas'] = (
    "CREATE TABLE `celulas` ("
    "  `id` int(11) NOT NULL AUTO_INCREMENT,"
    "  `nombre` varchar(150) NOT NULL,"
    "  `implementador_email` varchar(250) NOT NULL,"
    "  `gerente_email` varchar(250) NOT NULL,"
    "  `pm_email` varchar(250) NOT NULL,"
    "  PRIMARY KEY (`id`)"
    ") ENGINE=InnoDB"
)

# Tabla que contiene los datos técnicos y configuraciones de cada marca auditada
TABLES['marcas'] = (
    "CREATE TABLE `marcas` ("
    "  `id` int(11) NOT NULL AUTO_INCREMENT,"
    "  `celula_id` int(11) NOT NULL,"
    "  `marca` varchar(150) NOT NULL,"
    "  `pais` varchar(50) NOT NULL,"
    "  `entorno` varchar(50) NOT NULL,"
    "  `responsive` boolean DEFAULT True,"
    "  `doc_contenido` varchar(255) NOT NULL,"
    "  `link_marca` varchar(255) NOT NULL,"
    "  `clases_h1` varchar(500) NOT NULL,"
    "  `clases_desc_corta` text NOT NULL,"
    "  `clases_desc_larga` text NOT NULL,"
    "  `palabra_excluir` varchar(255) DEFAULT '',"
    "  `headers_personalizados` text DEFAULT NULL,"
    "  `fecha_headers` datetime DEFAULT NULL,"
    "  `alerta_headers_enviada` boolean DEFAULT False,"
    "  PRIMARY KEY (`id`),"
    "  CONSTRAINT `fk_celula` FOREIGN KEY (`celula_id`) "
    "     REFERENCES `celulas` (`id`) ON DELETE CASCADE"
    ") ENGINE=InnoDB"
)

# Tabla para almacenar parámetros globales de la aplicación
TABLES['configuracion'] = (
    "CREATE TABLE `configuracion` ("
    "  `clave` varchar(50) NOT NULL,"
    "  `valor` varchar(255) NOT NULL,"
    "  PRIMARY KEY (`clave`)"
    ") ENGINE=InnoDB"
)

# Registro histórico de cada ejecución de auditoría y sus resultados
TABLES['historico'] = (
    "CREATE TABLE `historico` ("
    "  `id` int(11) NOT NULL AUTO_INCREMENT,"
    "  `marca_id` int(11) NOT NULL,"
    "  `numero_revision` int(11) NOT NULL DEFAULT 1,"
    "  `fecha_ejecucion` datetime NOT NULL,"
    "  `urls_procesadas` int(11) NOT NULL DEFAULT 0,"
    "  `urls_redireccion` int(11) NOT NULL DEFAULT 0,"
    "  `urls_no_creadas` int(11) NOT NULL DEFAULT 0,"
    "  `urls_meta_titulo` int(11) NOT NULL DEFAULT 0,"
    "  `urls_meta_descripcion` int(11) NOT NULL DEFAULT 0,"
    "  `urls_h1` int(11) NOT NULL DEFAULT 0,"
    "  `urls_desc_corta` int(11) NOT NULL DEFAULT 0,"
    "  `urls_desc_larga` int(11) NOT NULL DEFAULT 0,"
    "  `ruta_reporte_excel` varchar(500) NOT NULL,"
    "  `urls_contenido_arreglado` int(11) DEFAULT 0,"
    "  `urls_contenido_cambiado` int(11) DEFAULT 0,"
    "  PRIMARY KEY (`id`),"
    "  CONSTRAINT `fk_marca` FOREIGN KEY (`marca_id`) "
    "     REFERENCES `marcas` (`id`) ON DELETE CASCADE"
    ") ENGINE=InnoDB"
)

def create_database(cursor):
    try:
        cursor.execute(f"CREATE DATABASE {DB_NAME} DEFAULT CHARACTER SET 'utf8'")
    except mysql.connector.Error as err:
        print(f"Error creando la base de datos: {err}\n")
        exit(1)

# Script principal para inicializar el entorno de base de datos
def setup_db(user, password, host="127.0.0.1"):
    try:
        cnx = mysql.connector.connect(user=user, password=password, host=host)
        cursor = cnx.cursor()
    except mysql.connector.Error as err:
        if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
            print("Hubo un problema con tu usuario o contraseña.")
        else:
            print(err)
        return

    try:
        cursor.execute(f"USE {DB_NAME}")
    except mysql.connector.Error as err:
        if err.errno == errorcode.ER_BAD_DB_ERROR:
            create_database(cursor)
            print(f"Base de datos {DB_NAME} creada exitosamente.")
            cnx.database = DB_NAME
        else:
            print(err)
            return

    # Iterar sobre las definiciones de las tablas para crearlas si no existen
    for table_name in TABLES:
        table_description = TABLES[table_name]
        try:
            print(f"Creando tabla {table_name}: ", end='')
            cursor.execute(table_description)
        except mysql.connector.Error as err:
            if err.errno == errorcode.ER_TABLE_EXISTS_ERROR:
                print("Ya existe.")
            else:
                print(err.msg)
        else:
            print("OK")

    # Insertar valores de configuración iniciales para el flujo de trabajo
    try:
        config_defaults = [
            ('lider_implementacion', 'lider_impl@empresa.com'),
            ('lider_seo', 'lider_seo@empresa.com'),
            ('lider_operaciones', 'lider_ops@empresa.com')
        ]
        cursor.executemany(
            "INSERT IGNORE INTO configuracion (clave, valor) VALUES (%s, %s)",
            config_defaults
        )
        cnx.commit()
    except Exception as e:
        print(f"Error insertando configuración: {e}")

    # Crear el usuario administrador predeterminado para el acceso inicial
    try:
        hashed = generate_password_hash("admin123")
        cursor.execute(
            "INSERT IGNORE INTO usuarios (username, email, password_hash, rol) VALUES (%s, %s, %s, %s)",
            ("admin", "admin@seoauditor.com", hashed, "lider")
        )
        cnx.commit()
        print("Usuario lider creado: admin / admin123 (cambia la contraseña al iniciar sesión)")
    except Exception as e:
        print(f"Lider ya existente o error: {e}")

    # Insertar célula por defecto
    try:
        cursor.execute(
            "INSERT IGNORE INTO celulas (nombre, implementador_email, gerente_email, pm_email) "
            "VALUES (%s, %s, %s, %s)",
            ("Célula General", "implementador@empresa.com", "gerente@empresa.com", "pm@empresa.com")
        )
        cnx.commit()
    except Exception as e:
        print(f"Error insertando célula: {e}")

    cursor.close()
    cnx.close()
    print("\n¡Configuración Base de Datos completada!")

if __name__ == '__main__':
    print("--- Configuración Inicial de Base de Datos ---")
    u = input("Usuario MySQL (ej. root): ") or "root"
    p = input("Contraseña MySQL (Enter si vacía): ")
    h = input("Host (Enter para 127.0.0.1): ") or "127.0.0.1"
    setup_db(user=u, password=p, host=h)
