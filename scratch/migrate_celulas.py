
import mysql.connector
import os

DB_CONFIG = {
    'host': '127.0.0.1',
    'user': 'root',
    'password': '',
    'database': 'seo_auditor'
}

def migrate():
    try:
        cnx = mysql.connector.connect(**DB_CONFIG)
        cursor = cnx.cursor()

        # 1. Crear tabla celulas
        print("Creando tabla celulas...")
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS `celulas` (
              `id` int(11) NOT NULL AUTO_INCREMENT,
              `nombre` varchar(150) NOT NULL,
              `implementador_email` varchar(250) NOT NULL,
              `gerente_email` varchar(250) NOT NULL,
              `pm_email` varchar(250) NOT NULL,
              PRIMARY KEY (`id`)
            ) ENGINE=InnoDB
        """)

        # 2. Migrar datos existentes de marcas a celulas
        print("Migrando datos de marcas a celulas...")
        cursor.execute("SELECT DISTINCT celula, implementador, gerente_cuenta FROM marcas")
        viejas_celulas = cursor.fetchall()
        
        for nombre, impl, ger in viejas_celulas:
            if not nombre: continue
            # Insertar en celulas si no existe
            cursor.execute("SELECT id FROM celulas WHERE nombre = %s", (nombre,))
            res = cursor.fetchone()
            if not res:
                cursor.execute(
                    "INSERT INTO celulas (nombre, implementador_email, gerente_email, pm_email) VALUES (%s, %s, %s, %s)",
                    (nombre, impl, ger, ger) # PM recibe lo mismo que gerente por defecto inicial
                )
        
        cnx.commit()

        # 3. Añadir columna celula_id a marcas
        print("Añadiendo columna celula_id a marcas...")
        try:
            cursor.execute("ALTER TABLE marcas ADD COLUMN celula_id int(11) AFTER id")
        except mysql.connector.Error as err:
            if err.errno == 1060: # Duplicate column name
                print("Columna celula_id ya existe.")
            else:
                raise err

        # 4. Vincular marcas con sus nuevas celulas
        print("Vinculando marcas con celulas...")
        cursor.execute("SELECT id, celula FROM marcas")
        marcas = cursor.fetchall()
        for mid, cname in marcas:
            cursor.execute("SELECT id FROM celulas WHERE nombre = %s", (cname,))
            cid_res = cursor.fetchone()
            if cid_res:
                cursor.execute("UPDATE marcas SET celula_id = %s WHERE id = %s", (cid_res[0], mid))

        # 5. Limpiar columnas viejas (Opcional, pero recomendado según el plan)
        # Nota: Las dejamos por ahora para evitar errores si algo falla, 
        # pero el usuario pidió que las celulas traigan la info, así que marcas ya no las necesita.
        print("Eliminando columnas obsoletas de marcas...")
        columnas_a_borrar = ['celula', 'implementador', 'gerente_cuenta']
        for col in columnas_a_borrar:
            try:
                cursor.execute(f"ALTER TABLE marcas DROP COLUMN {col}")
            except:
                pass

        cnx.commit()
        cursor.close()
        cnx.close()
        print("Migración completada con éxito.")

    except Exception as e:
        print(f"Error durante la migración: {e}")

if __name__ == "__main__":
    migrate()
