import mysql.connector

DB_CONFIG = {
    'host': '127.0.0.1',
    'user': 'root',
    'password': '',
    'database': 'seo_auditor'
}

def migrate():
    try:
        cnx = mysql.connector.connect(**DB_CONFIG)
        cur = cnx.cursor()
        
        # 1. Modificar tabla marcas
        print("Migrando tabla 'marcas'...")
        try:
            # Primero añadimos las nuevas para no perder datos si falla a la mitad
            cur.execute("ALTER TABLE marcas ADD COLUMN implementador VARCHAR(250) NOT NULL DEFAULT 'vacio@vacio.com'")
            cur.execute("ALTER TABLE marcas ADD COLUMN gerente_cuenta VARCHAR(250) NOT NULL DEFAULT 'vacio@vacio.com'")
            print("Nuevas columnas añadidas.")
        except Exception as e:
            print(f"Nota: {e}")

        try:
            # Ahora eliminamos la antigua
            cur.execute("ALTER TABLE marcas DROP COLUMN correo_notificacion")
            print("Columna 'correo_notificacion' eliminada.")
        except Exception as e:
            print(f"Nota (quizás ya fue eliminada): {e}")

        # 2. Crear tabla configuracion
        print("Creando tabla 'configuracion'...")
        cur.execute("""
            CREATE TABLE IF NOT EXISTS configuracion (
                clave VARCHAR(50) PRIMARY KEY,
                valor VARCHAR(255) NOT NULL
            ) ENGINE=InnoDB
        """)
        
        # Insertar valores por defecto
        config_defaults = [
            ('lider_implementacion', 'lider_impl@empresa.com'),
            ('lider_seo', 'lider_seo@empresa.com'),
            ('lider_operaciones', 'lider_ops@empresa.com')
        ]
        cur.executemany(
            "INSERT IGNORE INTO configuracion (clave, valor) VALUES (%s, %s)",
            config_defaults
        )
        print("Tabla 'configuracion' creada y con valores iniciales.")

        cnx.commit()
        cur.close()
        cnx.close()
        print("Migración completada con éxito.")

    except Exception as e:
        print(f"Error en la migración: {e}")

if __name__ == '__main__':
    migrate()
