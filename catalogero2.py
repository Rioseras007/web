import os
from urllib.parse import quote
from datetime import datetime
import ctypes
import string
import tkinter as tk
from tkinter import messagebox, scrolledtext
from pathlib import Path

# --- FUNCIONES AUXILIARES ---
def unidades_dvd():
    """
    Detecta unidades de CD/DVD en Windows
    """
    drives = []
    # Intenta detectar todas las letras de unidad posibles
    for letra in string.ascii_uppercase:
        ruta = letra + ":\\"
        if os.path.exists(ruta):
            try:
                tipo = ctypes.windll.kernel32.GetDriveTypeW(ruta)
                if tipo == 5:  # 5 = CD-ROM
                    drives.append(ruta)
            except Exception:
                pass
    return drives

def obtener_duracion(ruta_archivo):
    """
    Obtiene la duración de un archivo de audio en formato MM:SS
    usando la API MCI de Windows (ctypes).
    """
    buf = ctypes.create_unicode_buffer(128)
    # Es importante usar comillas dobles alrededor de la ruta para manejar espacios
    # Usamos un alias único o genérico; aquí 'mp3file' reaprovechado
    cmd_open = f'open "{ruta_archivo}" type MPEGVideo alias mp3file'
    cmd_status = "status mp3file length"
    cmd_close = "close mp3file"
    
    # Cerrar por si acaso quedó abierto
    ctypes.windll.winmm.mciSendStringW(cmd_close, None, 0, 0)
    
    ret = ctypes.windll.winmm.mciSendStringW(cmd_open, None, 0, 0)
    if ret != 0:
        return None
        
    res = ctypes.windll.winmm.mciSendStringW(cmd_status, buf, 128, 0)
    ctypes.windll.winmm.mciSendStringW(cmd_close, None, 0, 0)
    
    if res == 0:
        try:
            ms = int(buf.value)
            segundos_totales = ms // 1000
            minutos = segundos_totales // 60
            segundos = segundos_totales % 60
            return f"{minutos:02d}:{segundos:02d}"
        except ValueError:
            return None
    return None

class CatalogeroApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Catálogo DVD")
        self.root.geometry("500x400")
        
        # Etiqueta y campo para Número de Catálogo
        lbl_cat = tk.Label(root, text="Número de Catálogo:")
        lbl_cat.pack(pady=(20, 5))
        
        self.entry_cat = tk.Entry(root, font=("Arial", 12))
        self.entry_cat.insert(0, "AR-VD-001") # Valor por defecto
        self.entry_cat.pack(pady=5)
        
        # Botón para detectar y generar
        btn_gen = tk.Button(root, text="Generar Catálogo", command=self.generar_catalogo, 
                            bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), height=2)
        btn_gen.pack(pady=20, fill='x', padx=40)
        
        # Area de log
        self.log_area = scrolledtext.ScrolledText(root, width=50, height=10, state='disabled')
        self.log_area.pack(pady=10, padx=10, fill='both', expand=True)
        
        self.log("Aplicación iniciada. Listo para generar.")

    def log(self, message):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')
        # Actualizar la interfaz gráfica para que no se congele
        self.root.update()

    def generar_catalogo(self):
        catalogo_num = self.entry_cat.get().strip()
        if not catalogo_num:
            messagebox.showwarning("Aviso", "Por favor ingrese un número de catálogo.")
            return

        self.log("Buscando unidad de DVD...")
        dvds = unidades_dvd()
        
        if not dvds:
            self.log("❌ No se detectó ningún DVD/CD insertado.")
            messagebox.showerror("Error", "No se detectó ningún DVD/CD insertado.")
            return
        
        UNIDAD = dvds[0]
        self.log(f"✅ DVD detectado en {UNIDAD}")
        
        try:
            self.crear_html(UNIDAD, catalogo_num)
        except Exception as e:
            self.log(f"❌ Error al generar: {e}")
            messagebox.showerror("Error", f"Ocurrió un error: {e}")

    def crear_html(self, UNIDAD, CATALOGO_NUM):
        # Nombre de salida basado en la unidad
        SALIDA = f"index_DVD_{UNIDAD[0]}.html"
        # Ruta completa al escritorio para asegurar que se vea
        escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
        ruta_salida = os.path.join(escritorio, SALIDA)
        
        # --- CABECERA HTML ---
        html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Catálogo DVD Musical</title>
<style>
body {{
  background:#2b2b2b;
  color:#f2f2f2;
  font-family:Georgia, serif;
  display:flex;
  justify-content:center;
  margin:0;
  padding:40px;
}}
.cover {{
  background:linear-gradient(145deg,#111,#333);
  width:700px;
  padding:30px;
  border:6px solid #000;
  box-shadow:0 0 40px rgba(0,0,0,.8);
}}
.title {{
  text-align:center;
  font-size:30px;
  letter-spacing:2px;
}}
.subtitle {{
  text-align:center;
  font-size:14px;
  color:#ccc;
  margin-bottom:25px;
}}
.section h2 {{
  color:#ffd700;
  border-bottom:1px solid #777;
}}
ul {{ padding-left:20px; }}
li {{ margin-bottom:6px; }}
a {{ color:#f2f2f2; text-decoration:none; }}
a:hover {{ color:#ffd700; text-decoration:underline; }}
.footer {{
  text-align:center;
  font-size:12px;
  margin-top:25px;
  color:#aaa;
}}

/* CSS PARA IMPRESIÓN COMPACTA + 2 COLUMNAS */
@media print {{
  @page {{
    size: A4;
    margin: 1.5cm;
  }}
  body {{
    background:white;
    color:black;
    padding:0;
    margin:0;
    font-family: "Times New Roman", serif;
  }}
  .cover {{
    background:white;
    color:black;
    box-shadow:none;
    border:none;
    width:100%;
    padding:0;
    margin-bottom: 20px;
  }}
  .title {{ 
    font-size:24px; 
    font-weight: bold;
    text-align: center;
    margin-bottom: 5px;
  }}
  .subtitle {{ 
    font-size:12px; 
    text-align: center;
    color: #444;
    margin-bottom: 20px;
    border-bottom: 1px solid #000;
    padding-bottom: 10px;
  }}
  .section {{
    page-break-inside: avoid; /* Evita que una sección se corte a la mitad */
    margin-bottom: 15px;
  }}
  .section h2 {{ 
    font-size:14px; 
    font-weight: bold;
    border-bottom:2px solid black; 
    margin-top: 0;
    margin-bottom: 5px;
    color: black !important; /* Forza negro para impresión */
  }}
  ul {{
    font-size:10px;
    padding-left: 0;
    margin: 0;
    column-count: 3; 
    column-gap: 15px;
    list-style-type: none;
  }}
  li {{ 
    margin-bottom: 1px;
    line-height: 1.2;
  }}
  a {{ 
    color:black; 
    text-decoration:none; 
  }}
  .footer {{ 
    position: fixed;
    bottom: 0;
    width: 100%;
    border-top: 1px solid #ccc;
    font-size: 8px;
    text-align: center;
    padding-top: 5px;
  }}
  /* Ocultar imágenes en impresión para ahorrar tinta? O las mostramos?
     Si el usuario pidió imprimirlas, las dejamos. Ajustamos tamaño si es necesario. */
  .cover-img {{
    max-width: 60px; /* Más pequeño en impresión */
    border: 1px solid #000;
  }}
}}
/* Estilos generales para pantalla */
.thumb-item {{
    list-style: none;
    margin: 4px 0;
}}
.cover-img {{
    max-width: 80px;
    height: auto;
    border: 1px solid #555;
    padding: 2px;
    background: #000;
    display: block;
}}
</style>
</head>
<body>
<div class="cover">
<div class="title">ARCHIVO MUSICAL</div>
<div class="subtitle">Unidad: {UNIDAD} · Catalogado: {CATALOGO_NUM}</div>
"""
        self.log("Escaneando toda la unidad (calculando duraciones)...")
        self.root.update() # Forzar actualización UI
        
        # Estructura intermedia:
        # { categoria: [ {'type': 'mp3'|'jpg', 'size': int, 'name': str, 'html': str}, ... ] }
        data_por_categoria = {}
        # Para la raíz, usaremos una clave especial None
        data_por_categoria[None] = []

        try:
            total_archivos = 0
            UNIDAD = os.path.abspath(UNIDAD) # Asegurar ruta absoluta de raiz
            for root, dirs, files in os.walk(UNIDAD):
                for archivo in files:
                    ext = archivo.lower()
                    if ext.endswith(".mp3") or ext.endswith(".jpg"):
                        total_archivos += 1
                        if total_archivos % 10 == 0: self.root.update()
                        
                        ruta_absoluta = os.path.abspath(os.path.join(root, archivo))
                        ruta_relativa = os.path.relpath(ruta_absoluta, UNIDAD)
                        partes = ruta_relativa.split(os.sep)
                        
                        # CREAR LINK ABSOLUTO
                        path_link = Path(ruta_absoluta).as_uri()
                        
                        if len(partes) > 1:
                            categoria = partes[0]
                            nombre_base = os.path.join(*partes[1:]).replace("\\", "/")
                        else:
                            categoria = None # Raíz
                            nombre_base = archivo
                        
                        # Inicializar lista si no existe
                        if categoria not in data_por_categoria:
                            data_por_categoria[categoria] = []

                        item_data = {
                            'name': nombre_base,
                            'size': 0,
                            'html': ""
                        }

                        if ext.endswith(".mp3"):
                            item_data['type'] = 'mp3'
                            duracion = obtener_duracion(ruta_absoluta)
                            dur_txt = f" ... {duracion}" if duracion else ""
                            item_data['html'] = f"<li><a href='{path_link}'>{nombre_base}{dur_txt}</a></li>"
                        else: # JPG
                            item_data['type'] = 'jpg'
                            try:
                                item_data['size'] = os.path.getsize(ruta_absoluta)
                            except OSError:
                                item_data['size'] = 0
                            item_data['html'] = f"<li class='thumb-item'><img class='cover-img' src='{path_link}'></li>"
                        
                        data_por_categoria[categoria].append(item_data)

        except OSError as e:
            self.log(f"Error: {e}"); messagebox.showerror("Error", str(e)); return

        # --- PROCESAR Y RENDERIZAR ---
        
        # Función helper para procesar una lista de items
        def procesar_items(lista_items):
            # Separar MP3 y JPG
            mp3s = [x for x in lista_items if x['type'] == 'mp3']
            jpgs = [x for x in lista_items if x['type'] == 'jpg']
            
            # Filtrar JPGs: Ordenar por tamaño descendente y tomar top 2
            jpgs.sort(key=lambda x: x['size'], reverse=True)
            jpgs_top = jpgs[:2]
            
            # Combinar
            todos = mp3s + jpgs_top
            # Ordenar alfabéticamente por nombre
            todos.sort(key=lambda x: x['name'])
            
            # Generar HTML
            return "".join([x['html'] for x in todos])

        cara = 65 # Initialize cara (ASCII A)

        # 1. Raíz
        if data_por_categoria[None]:
            html_bloque = procesar_items(data_por_categoria[None])
            # Si quedó algo (podrían haberse filtrado todas las imgs si solo habia imgs malas? no, top 2 siempre queda)
            if html_bloque:
                html_content += f"<div class='section'><h2>Cara {chr(cara)} · Archivos en Raíz</h2><ul>"
                html_content += html_bloque
                html_content += "</ul></div>"
                cara += 1

        # 2. Categorías
        # Filtrar claves que no sean None y ordenar
        cats = [k for k in data_por_categoria.keys() if k is not None]
        cats.sort()
        
        for cat in cats:
            html_bloque = procesar_items(data_por_categoria[cat])
            if html_bloque:
                html_content += f"<div class='section'><h2>Cara {chr(cara)} · {cat}</h2><ul>"
                html_content += html_bloque
                html_content += "</ul></div>"
                cara += 1

        # --- PIE DE PÁGINA ---
        fecha = datetime.now().strftime("%Y-%m-%d")
        html_content += f"""
<div class="footer">
℗ Alfonso · Archivo musical físico · Catalogado {fecha} · {CATALOGO_NUM}
</div>
</div>
</body>
</html>
"""
        # --- GUARDAR HTML ---
        with open(ruta_salida, "w", encoding="utf-8") as f:
            f.write(html_content)

        self.log(f"✔ Catálogo generado exitosamente en:\n{ruta_salida}")
        messagebox.showinfo("Éxito", f"Catálogo generado:\n{ruta_salida}")

if __name__ == "__main__":
    root = tk.Tk()
    app = CatalogeroApp(root)
    root.mainloop()
