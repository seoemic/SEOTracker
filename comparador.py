import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinter as tk 
from bs4 import BeautifulSoup
import os
from PIL import Image, ImageTk
import re

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

def extraer_h2h3_html(texto_original):
    if isinstance(texto_original, str):
        soup = BeautifulSoup(texto_original, 'html.parser')
        encabezados = soup.find_all(['h2', 'h3'])
        
        output = []
        for tag in encabezados:
            texto_limpio = tag.get_text(strip=True)
            if texto_limpio:
                output.append(f"[{tag.name.upper()}] {texto_limpio}")
        
        return " | ".join(output) if output else "Ninguno"
    return "Ninguno"


def extraer_h2h3(texto_original):
    h2h3_pattern = r'(H2:|H3:)(.*?)(\n|\r|H2:|H3:|$)'
    
    if isinstance(texto_original, str):
        
        texto_limpio_para_regex = texto_original 
        
        matches = re.findall(h2h3_pattern, texto_limpio_para_regex, re.IGNORECASE | re.DOTALL)
        
        output = []
        for match in matches:
            marcador = match[0].replace(':', '').strip()
            contenido = match[1].strip()
            
            contenido = re.sub(r'\s+([.,:;])', r'\1', contenido)
            contenido = re.sub(r' +', ' ', contenido)
            
            if contenido:
                output.append(f"[{marcador.upper()}] {contenido}")
        
        return " | ".join(output) if output else "Ninguno"
    return "Ninguno"

def comparar_hojas(path_screaming, path_contenido, palabra):
    try:
        df_screaming = pd.read_excel(path_screaming)
        df_contenido = pd.read_excel(path_contenido)

        df_contenido['URL'] = df_contenido['URL'].astype(str).str.strip() 
        
        df_screaming['Dirección'] = df_screaming['Dirección'].astype(str).str.strip()

        df_comparacion = pd.merge(df_contenido, df_screaming,
                                  left_on='URL', right_on='Dirección',
                                  how='left',
                                  suffixes=('_contenido', '_screaming'))

        def limpiar_html(texto):
            if isinstance(texto, str):
                texto = texto.replace('\n', ' ').replace('\r', ' ')
                soup = BeautifulSoup(texto, 'html.parser')
                limpio = soup.get_text(separator=' ', strip=True)
                texto = re.sub(r'\s+([.,:;])', r'\1', texto)
                texto = re.sub(r' +', ' ', texto)
                limpio = limpio.replace('H2:', '').replace('H3:', '').strip()
                limpio = re.sub(r' +', ' ', limpio).strip()
                return limpio
            return str(texto).strip()

        df_comparacion['DL_ORIGINAL'] = df_comparacion['Descripción Larga'].astype(str)
        df_comparacion['DL_SF'] = df_comparacion['Descripción larga 1'].astype(str)

        df_comparacion['Descripcion Larga Rastreo Contiene H2 o H3'] = df_comparacion['DL_SF'].str.contains(
            r'<h2 |<h3 ', 
            case=False, 
            regex=True
        ).map({True: 'Contiene H2 o H3 en el Rastreo', False: 'Limpia'})

        
        df_comparacion['Descripcion Larga Contiene H2 o H3'] = df_comparacion['DL_ORIGINAL'].str.contains(
            r'H2:|H3:', 
            case=False,
            regex=True
        ).map({True: 'Contiene H2 o H3 en la Hoja de contenido', False: 'Limpia'})

        df_comparacion['CONTENIDO_H2_H3'] = df_comparacion['DL_ORIGINAL'].apply(extraer_h2h3)
        df_comparacion['CONTENIDO_SF_H2_H3'] = df_comparacion['DL_SF'].apply(extraer_h2h3_html)

        palabra_excluir = palabra.strip()
        
        def limpiar_metatitulo(titulo):
            titulo_limpio = limpiar_html(titulo)
            if palabra_excluir:
                titulo_limpio = titulo_limpio.replace(palabra_excluir, '').strip()
                titulo_limpio = re.sub(r' +', ' ', titulo_limpio)
            return titulo_limpio

        df_comparacion['Título 1'] = df_comparacion['Título 1'].apply(limpiar_metatitulo)

        df_comparacion['Meta Titulo'] = df_comparacion['Meta Titulo'].apply(limpiar_html)
        df_comparacion['Título 1'] = df_comparacion['Título 1'].apply(limpiar_html)
        df_comparacion['Meta Descripción'] = df_comparacion['Meta Descripción'].apply(limpiar_html)

        df_comparacion['Meta Titulo'] = df_comparacion['Meta Titulo'].str.replace('-', '|', regex=False)
        df_comparacion['Título 1'] = df_comparacion['Título 1'].str.replace('-', '|', regex=False)
        
        df_comparacion['Descripción Corta'] = df_comparacion['Descripción Corta'].apply(limpiar_html)
        df_comparacion['Descripción Larga'] = df_comparacion['Descripción Larga'].apply(limpiar_html)

        
        df_comparacion['Meta Description 1'] = df_comparacion['Meta Description 1'].apply(limpiar_html)
        df_comparacion['H1-1'] = df_comparacion['H1-1'].apply(limpiar_html)
        
        df_comparacion['Descripción corta 1'] = df_comparacion['Descripción corta 1'].apply(limpiar_html)
        df_comparacion['Descripción larga 1'] = df_comparacion['Descripción larga 1'].apply(limpiar_html)
        df_comparacion['Descripción larga 1'] = df_comparacion['Descripción larga 1'].str.replace(r' +', ' ', regex=True)
        df_comparacion['Descripción larga 1'] = df_comparacion['Descripción larga 1'].str.replace(' ,', ',', regex=False)
        df_comparacion['Descripción larga 1'] = df_comparacion['Descripción larga 1'].str.replace(' .', '.', regex=False)
        df_comparacion['Descripción larga 1'] = df_comparacion['Descripción larga 1'].str.replace(' !', '!', regex=False)

        # cols_sf = ['Título 1', 'Meta Description 1', 'H1-1', 'Descripción corta 1', 'Descripción larga 1' ]
        
        # for col in cols_sf:
        #      df_comparacion[col] = df_comparacion[col].replace('nan', 'URL NO ENCONTRADA EN RASTREO')

        def clasificar_codigo(codigo):
            
            if str(codigo) == '':
                return 'NO ENCONTRADA EN RASTREO'
            
            try:
               
                codigo_int = int(float(codigo)) 
            except (ValueError, TypeError):
                return 'NO ENCONTRADA EN RASTREO' 

            if codigo_int == 200:
                return 'OK'
            elif 300 <= codigo_int <= 399:
                return 'REDIRECCIÓN'
            else:
                
                return 'OTRO'
        
        
        df_comparacion['CLASIFICACION_RESPUESTA'] = df_comparacion['Código de respuesta'].apply(clasificar_codigo)


        df_comparacion['meta_titulo_coincide'] = (
            df_comparacion['Meta Titulo'].str.lower() == df_comparacion['Título 1'].str.lower()
        ).map({True: 'Sí', False: 'No'})
        
        df_comparacion['meta_descripcion_coincide'] = (
            df_comparacion['Meta Descripción'].str.lower() == df_comparacion['Meta Description 1'].str.lower()
        ).map({True: 'Sí', False: 'No'})

        df_comparacion['h1_coincide'] = (
            df_comparacion['H1'].str.lower() == df_comparacion['H1-1'].str.lower()
        ).map({True: 'Sí', False: 'No'})

        df_comparacion['descripcion_corta_coincide'] = (
            df_comparacion['Descripción Corta'].str.lower() == df_comparacion['Descripción corta 1'].str.lower()
        ).map({True: 'Sí', False: 'No'})

        df_comparacion['descripcion_larga_coincide'] = (
            df_comparacion['Descripción Larga'].str.lower() == df_comparacion['Descripción larga 1'].str.lower()
        ).map({True: 'Sí', False: 'No'})


        columnas_finales = [
            'URL',

            'Código de respuesta',

            'CLASIFICACION_RESPUESTA',
            
            'Meta Titulo',  'meta_titulo_coincide', 'Título 1',

            'Meta Descripción',  'meta_descripcion_coincide', 'Meta Description 1',

            'H1', 'h1_coincide', 'H1-1', 
     
            'Descripción Corta',  'descripcion_corta_coincide', 'Descripción corta 1',
        
            'Descripción Larga',  'descripcion_larga_coincide', 'Descripción larga 1',

            'Descripcion Larga Contiene H2 o H3',

            'Descripcion Larga Rastreo Contiene H2 o H3',

            'CONTENIDO_H2_H3',

            'CONTENIDO_SF_H2_H3',
        ]
        
        df_comparacion.dropna(subset=['URL'], inplace=True)
        
        # Renombre 
        nuevo_nombre_columnas = {
            'Meta Titulo': 'Meta Titulo (Hoja de Contenido)',
            'Título 1': 'Meta Titulo (Rastreo Screaming Frog)',
            'Meta Descripción': 'Meta Descripción (Hoja de Contenido)',
            'Meta Description 1': 'Meta Descripción (Rastreo Screaming Frog)',
            'H1': 'H1 (Hoja de Contenido)',
            'H1-1': 'H1 (Rastreo Screaming Frog)',
            'Descripción Corta': 'Descripción Corta (Hoja de Contenido)',
            'Descripción corta 1': 'Descripción Corta (Rastreo Screaming Frog)',
            'Descripción Larga': 'Descripción Larga (Hoja de Contenido)',
            'Descripción larga 1': 'Descripción Larga (Rastreo Screaming Frog)',
            
            'meta_titulo_coincide': 'META_TITULO_COINCIDE',
            'meta_descripcion_coincide': 'META_DESCRIPCION_COINCIDE',
            'h1_coincide': 'H1_COINCIDE',
            'descripcion_corta_coincide': 'DESCRIPCION_CORTA_COINCIDE',
            'descripcion_larga_coincide': 'DESCRIPCION_LARGA_COINCIDE',
            'CONTENIDO_H2_H3': 'Titulos en Descripcion Larga HOJA DE CONTENIDO', 
            'CONTENIDO_SF_H2_H3': 'TItulos en Descripcion Larga RASTREO (SF)',
        }
        
        df_resultado = df_comparacion[columnas_finales].rename(columns=nuevo_nombre_columnas)

        return df_resultado

    except FileNotFoundError:
        messagebox.showerror("Error", "No se pudo encontrar uno de los archivos seleccionados.")
        return None
    except KeyError as e:
        
        messagebox.showerror("Error de Columna", f"Falta una columna necesaria en una de las hojas: {e}. Por favor, revisa que los nombres de las columnas coincidan EXACTAMENTE con el código.")
        return None
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error inesperado: {e}")
        return None


class ComparadorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Comparador SEO de Hojas de Cálculo")
        self.geometry("600x500") 
        
        self.path_screaming = None
        self.path_contenido = None

        def resource_path(relative_path):
            """Obtiene la ruta absoluta del recurso para PyInstaller."""
            try:
                # Determina si está corriendo en un entorno PyInstaller
                base_path = sys._MEIPASS
            except Exception:
                # Si corre como script normal, usa la ruta del directorio actual
                base_path = os.path.abspath(".")
            return os.path.join(base_path, relative_path)
        
        icono_file = 'img/MIC.ico' 
        logo_file = 'img/Logo_MIC.png'

        try:
           
            icono_path = resource_path(icono_file)
            if os.path.exists(icono_path):
                 
                 self.iconbitmap(icono_path)
            else:
              
                 print(f"Advertencia: Icono no encontrado en {icono_path}. Usando icono predeterminado.")
                 pass

        except Exception as e:
            print(f"No se pudo cargar el icono: {e}")

        
        try:
           
            logo_path = resource_path(logo_file)
            
           
            logo_img = Image.open(logo_path).resize((150, 80))
            self.logo_tk = ctk.CTkImage(light_image=logo_img, dark_image=logo_img, size=(150, 50))

            
            self.logo_label = ctk.CTkLabel(self, text="", image=self.logo_tk)
            self.logo_label.pack(pady=(20, 0))

        except Exception as e:
            
            print(f"Advertencia: Logo no encontrado en {logo_path}. Usando solo título.")
            self.label_title = ctk.CTkLabel(self, text="Comparador SEO de Contenidos", font=ctk.CTkFont(size=20, weight="bold"))
            self.label_title.pack(pady=(20, 10))
           
            self.pady_value = 10 


       
        self.pady_value = 20 if hasattr(self, 'logo_label') else 10 

        
        self.label_title = ctk.CTkLabel(self, text="Comparador SEO de Contenidos", font=ctk.CTkFont(size=20, weight="bold"))
        self.label_title.pack(pady=(20, 10))

     
        frame_screaming = ctk.CTkFrame(self)
        frame_screaming.pack(pady=10, padx=20, fill="x")
        self.label_screaming = ctk.CTkLabel(frame_screaming, text="Hoja de Screaming Frog (SF):")
        self.label_screaming.pack(side="left", padx=(10, 5), pady=5)
        self.btn_screaming = ctk.CTkButton(frame_screaming, text="Seleccionar", command=self.seleccionar_screaming, width=100)
        self.btn_screaming.pack(side="right", padx=(5, 10), pady=5)
        self.entry_screaming = ctk.CTkEntry(frame_screaming, width=280)
        self.entry_screaming.pack(side="right", expand=True, fill="x")

        frame_contenido = ctk.CTkFrame(self)
        frame_contenido.pack(pady=10, padx=20, fill="x")
        self.label_contenido = ctk.CTkLabel(frame_contenido, text="Hoja de Contenido (Target):")
        self.label_contenido.pack(side="left", padx=(10, 5), pady=5)
        self.btn_contenido = ctk.CTkButton(frame_contenido, text="Seleccionar", command=self.seleccionar_contenido, width=100)
        self.btn_contenido.pack(side="right", padx=(5, 10), pady=5)
        self.entry_contenido = ctk.CTkEntry(frame_contenido, width=280)
        self.entry_contenido.pack(side="right", expand=True, fill="x")

        frame_exclusion = ctk.CTkFrame(self)
        frame_exclusion.pack(pady=20, padx=20, fill="x")
        self.label_excluir = ctk.CTkLabel(frame_exclusion, text="Palabra a excluir del Meta Título (Opcional):")
        self.label_excluir.pack(side="left", padx=(10, 5), pady=5)
        self.entry_excluir = ctk.CTkEntry(frame_exclusion, width=200, placeholder_text="Ej: | Nombre de Marca |")
        self.entry_excluir.pack(side="right", expand=True, fill="x", padx=(5, 10))


        self.btn_comparar = ctk.CTkButton(self, text="Comparar y Exportar", command=self.ejecutar_comparacion, font=ctk.CTkFont(weight="bold"))
        self.btn_comparar.pack(pady=20, padx=20, fill="x")

        modo_uso_texto = (
            "\n"
            "⚠️ MODO DE USO:\n"
            "1. Ambos archivos deben ser .xlsx.\n"
            "2. La columna 'URL' debe existir en la hoja de Contenido.\n"
            "3. La columna 'Dirección' debe existir en la hoja de Rastreo (Screaming Frog).\n"
            "4. Celdas en rojo = NO COINCIDE.\n"
            "\n"
            "IMPORTANTE: Los nombres de las columnas no pueden tener filtros, estar fijadas ni en \n negrita, por lo que se recomienda borrar la primera fila y escribir nuevamente los nombres,\n si no coinciden la plataforma te dira cual es la columna que no encontro con su respectivo nombre"
        )
        self.label_instrucciones = ctk.CTkLabel(
            self, 
            text=modo_uso_texto, 
            text_color="#FF8C00", # Naranja
            font=ctk.CTkFont(size=11), 
            justify=tk.LEFT 
        )
        self.label_instrucciones.pack(pady=(0, 10), padx=20, fill="x")


    def seleccionar_screaming(self):
        self.path_screaming = filedialog.askopenfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos de Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
        )
        if self.path_screaming:
            self.entry_screaming.delete(0, ctk.END)
            self.entry_screaming.insert(0, self.path_screaming)

    def seleccionar_contenido(self):
        self.path_contenido = filedialog.askopenfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos de Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
        )
        if self.path_contenido:
            self.entry_contenido.delete(0, ctk.END)
            self.entry_contenido.insert(0, self.path_contenido)

    def ejecutar_comparacion(self):
        if not self.path_screaming or not self.path_contenido:
            messagebox.showwarning("Advertencia", "Por favor, selecciona ambos archivos antes de continuar.")
            return

        palabra = self.entry_excluir.get()

        df_resultado = comparar_hojas(self.path_screaming, self.path_contenido, palabra)

        if df_resultado is not None:
            output_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Archivos de Excel", "*.xlsx")],
                title="Guardar resultados de la comparación"
            )
            
            if output_path:
                try:
                    def resaltar_diferencias(fila):
                        colores = [''] * len(fila)
                        if fila.iloc[4] != 'Sí':
                            colores[4] = 'background-color: #FFB6C1' 
                        if fila.iloc[7] != 'Sí':
                            colores[7] = 'background-color: #FFB6C1'
                        if fila.iloc[10] != 'Sí':
                            colores[10] = 'background-color: #FFB6C1'
                        if fila.iloc[13] != 'Sí':
                            colores[13] = 'background-color: #FFB6C1'
                        if fila.iloc[16] != 'Sí':
                            colores[16] = 'background-color: #FFB6C1'

                        return colores

                    df_estilizado = df_resultado.style.apply(resaltar_diferencias, axis=1)
                    df_estilizado.to_excel(output_path, engine='openpyxl', index=False)
                    
                    messagebox.showinfo("Éxito", f"Comparación completada y guardada en:\n{output_path}")

                except Exception as e:
                    messagebox.showerror("Error al guardar", f"No se pudo guardar el archivo: {e}")

if __name__ == "__main__":
    app = ComparadorApp()
    app.mainloop()
