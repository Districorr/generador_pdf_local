# -*- coding: utf-8 -*-

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import DateEntry
import os
import subprocess
import sys
from pathlib import Path
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from datetime import datetime
import sqlite3
import uuid
import qrcode
from PIL import Image
import webbrowser
from urllib.parse import quote # Corregido

# --- Configuración ---
TEMPLATE_FILENAME = "template.docx"
TEMP_DOCX_FILENAME = "temp_output.docx" # Nombre genérico para archivo temporal
DB_FILENAME = "registros_cirugias.db"   # Archivo de base de datos SQLite
QR_TEMP_FILENAME = "temp_qr.png"         # Nombre genérico para QR temporal
# --- ¡¡¡IMPORTANTE: AJUSTA ESTA RUTA SI ES NECESARIO!!! ---
LIBREOFFICE_PATH = "C:/Program Files/LibreOffice/program/soffice.exe" # Ejemplo Windows
MAX_IMAGES_IN_TEMPLATE = 10 # ¿Cuántos placeholders {{imagen_X}} tienes?

# --- Funciones Auxiliares ---

def find_libreoffice():
    """Verifica la ruta configurada y busca en rutas comunes si es necesario."""
    manual_path = Path(LIBREOFFICE_PATH)
    # Ser más permisivo con la comprobación
    if (manual_path.is_file() or manual_path.parent.is_dir()) and manual_path.name.lower().startswith("soffice"):
        print(f"Usando ruta de LibreOffice configurada: {manual_path}")
        return str(manual_path)

    print(f"WARN: Ruta configurada '{LIBREOFFICE_PATH}' no encontrada o inválida. Buscando en rutas comunes...")
    if sys.platform == "win32":
        possible_paths = [ Path("C:/Program Files/LibreOffice/program/soffice.exe"), Path("C:/Program Files (x86)/LibreOffice/program/soffice.exe") ]
    elif sys.platform == "darwin": # macOS
        possible_paths = [ Path("/Applications/LibreOffice.app/Contents/MacOS/soffice") ]
    else: # Linux assumed
        possible_paths = [ Path("/usr/bin/soffice"), Path("/snap/bin/libreoffice.soffice") ]

    for path in possible_paths:
        if path.exists() and path.is_file():
            print(f"LibreOffice encontrado en ruta común: {path}")
            return str(path)

    print(f"ERROR CRITICO: No se encontró LibreOffice. Intentando usar '{LIBREOFFICE_PATH}' (debe estar en PATH o la ruta ser correcta).")
    return LIBREOFFICE_PATH

def show_error(title, message):
    messagebox.showerror(title, message)

def show_info(title, message):
    messagebox.showinfo(title, message)

def convert_to_pdf(docx_path, output_dir):
    """Convierte un archivo DOCX a PDF usando LibreOffice."""
    soffice_cmd = find_libreoffice()
    soffice_path = Path(soffice_cmd)
    # Verificar si el comando es ejecutable o está en PATH
    can_execute = soffice_path.is_file() or soffice_cmd.lower() == "soffice"
    if not can_execute:
         show_error("Error Conversión", f"Error: Ejecutable de LibreOffice no encontrado o no válido en la ruta:\n'{soffice_cmd}'.\n\nAsegúrate de que LibreOffice esté instalado y la constante LIBREOFFICE_PATH en el script sea correcta o 'soffice' si está en el PATH.")
         return False

    try:
        print(f"Intentando convertir '{docx_path}' a PDF en '{output_dir}'...")
        # Comando para convertir (Limpio)
        cmd = [
            soffice_cmd,
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', str(output_dir),
            str(docx_path)
        ]
        print("Ejecutando Comando:", " ".join(cmd))

        result = subprocess.run(cmd, capture_output=True, text=True, check=False, encoding='utf-8', errors='ignore', timeout=90) # Timeout 90s

        print(f"Código de retorno de LibreOffice: {result.returncode}")
        if result.stdout: print("Salida (stdout):\n", result.stdout)
        if result.stderr: print("Salida de Error (stderr):\n", result.stderr)

        # Verificar si el PDF se creó físicamente
        pdf_filename = Path(docx_path).with_suffix('.pdf').name
        expected_pdf_path = Path(output_dir) / pdf_filename

        # Evaluar resultado
        if result.returncode == 0 and expected_pdf_path.exists():
             print(f"Conversión a PDF completada. Archivo: {expected_pdf_path}")
             return True
        else:
             # Construir mensaje de error detallado
             error_msg = f"Error al convertir a PDF con LibreOffice.\n"
             if result.returncode != 0: error_msg += f"Código: {result.returncode}. "
             if not expected_pdf_path.exists(): error_msg += f"Archivo no encontrado: {expected_pdf_path}. "
             if result.stderr: error_msg += f"Error Output: {result.stderr[:250]}..." # Acortar
             elif result.stdout and result.returncode !=0 : error_msg += f"Info Output: {result.stdout[:250]}..."
             show_error("Error de Conversión", error_msg)
             return False

    except subprocess.TimeoutExpired:
        show_error("Error de Conversión", "Error: Timeout (90s) esperando a LibreOffice. ¿Está colgado o el documento es muy grande?")
        return False
    except FileNotFoundError:
        # Si soffice_cmd era solo "soffice" y no está en el PATH
        show_error("Error de Conversión", f"Error: Comando '{soffice_cmd}' no encontrado. Asegúrate de que LibreOffice esté instalado Y que la ruta en LIBREOFFICE_PATH sea correcta O que 'soffice' esté en el PATH del sistema.")
        return False
    except Exception as e:
        show_error("Error de Conversión", f"Error inesperado durante la conversión a PDF:\n{e}")
        return False

# --- Funciones de Base de Datos ---
def init_db():
    """Inicializa la DB. Crea la tabla con todas las columnas si no existe."""
    conn = None; db_path = Path(DB_FILENAME); print(f"DB Init: {db_path.resolve()}")
    if not db_path.parent.exists():
        try: db_path.parent.mkdir(parents=True, exist_ok=True); print(f"Directorio '{db_path.parent}' creado.")
        except OSError as e: print(f"Error Crítico: No crear dir DB: {e}"); show_error("Error Crítico DB", f"No crear dir:\n{db_path.parent}\n{e}"); sys.exit(f"Fatal DB dir: {e}")
    try:
        conn = sqlite3.connect(db_path); cursor = conn.cursor()
        # Crear tabla con estructura final
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS cirugias (
                id TEXT PRIMARY KEY, fecha_generacion TEXT NOT NULL, archivo_pdf TEXT,
                fecha_cirugia TEXT, cliente TEXT, paciente TEXT, medico TEXT, tecnico TEXT,
                tipo_cirugia TEXT, lugar TEXT, observaciones_generales TEXT,
                encargado_preparacion TEXT, encargado_logistica TEXT, coordinador_cx TEXT,
                observaciones_logistica TEXT, unique_id TEXT UNIQUE NOT NULL )
        ''')
        conn.commit(); print(f"Base de datos '{DB_FILENAME}' lista.")
    except sqlite3.Error as e: print(f"Error Crítico DB Init: {e}"); show_error("Error Crítico DB", f"No init DB:\n{e}"); sys.exit(f"Fatal DB: {e}")
    finally:
        if conn: conn.close()

def save_record(record_data):
    """Guarda un registro nuevo en la base de datos."""
    conn = None; required = ['id', 'unique_id', 'fecha_generacion'];
    if not all(f in record_data for f in required): show_error("Error Guardar DB", "Faltan id/unique_id/fecha_gen"); return False
    try:
        conn = sqlite3.connect(DB_FILENAME); cursor = conn.cursor();
        fields=list(record_data.keys()); values=list(record_data.values()); ph=','.join(['?']*len(fields));
        sql = f"INSERT INTO cirugias ({','.join(fields)}) VALUES ({ph})"
        cursor.execute(sql, values); conn.commit(); print(f"Guardado (ID: {record_data.get('id')})"); return True
    except sqlite3.IntegrityError as e:
         if "unique constraint failed: cirugias.unique_id" in str(e).lower(): show_error("Error DB", f"ID único ya existe:\n{record_data.get('unique_id')}")
         elif "not null constraint failed" in str(e).lower(): show_error("Error DB", f"Falta valor NOT NULL:\n{e}")
         else: show_error("Error DB", f"Error integridad:\n{e}")
         return False
    except sqlite3.Error as e: show_error("Error DB", f"Error guardando:\n{e}"); return False
    finally:
        if conn: conn.close()

def get_suggestions(field_name, limit=20):
    """Obtiene valores únicos recientes para un campo."""
    conn = None; suggestions = [];
    try:
        conn = sqlite3.connect(DB_FILENAME); cursor = conn.cursor();
        cursor.execute("PRAGMA table_info(cirugias)"); allowed = [col[1] for col in cursor.fetchall()];
        if field_name not in allowed: return []
        sql = f"SELECT DISTINCT {field_name} FROM cirugias WHERE {field_name} IS NOT NULL AND TRIM({field_name}) != '' ORDER BY fecha_generacion DESC LIMIT ?"
        cursor.execute(sql, (limit,)); suggestions = sorted([str(row[0]) for row in cursor.fetchall()])
    except sqlite3.Error as e: print(f"Error Sugerencias '{field_name}': {e}")
    finally:
        if conn: conn.close()
    return suggestions if suggestions else []

def get_record_by_id(record_id):
    """Obtiene todos los datos de un registro por su ID principal (el de la fila)."""
    conn = None; record = None
    try:
        conn = sqlite3.connect(DB_FILENAME); conn.row_factory = sqlite3.Row; cursor = conn.cursor()
        cursor.execute("SELECT * FROM cirugias WHERE id = ?", (record_id,))
        record = cursor.fetchone()
        if record: return dict(record)
        else: print(f"ID '{record_id}' no encontrado."); return None
    except sqlite3.Error as e: print(f"Error get ID '{record_id}': {e}"); return None
    finally:
        if conn: conn.close()

# --- Función get_counts_by_preparador (CORREGIDA) ---
def get_counts_by_preparador():
    """Obtiene el conteo de registros agrupados por encargado_preparacion."""
    conn = None; counts = []
    try:
        conn = sqlite3.connect(DB_FILENAME); cursor = conn.cursor()
        sql = """ SELECT encargado_preparacion, COUNT(*) as count FROM cirugias WHERE encargado_preparacion IS NOT NULL AND TRIM(encargado_preparacion) != '' GROUP BY encargado_preparacion ORDER BY count DESC """
        cursor.execute(sql); counts = cursor.fetchall()
        print(f"Conteos prep: {counts}")
    except sqlite3.Error as e: print(f"Error conteos prep: {e}")
    finally:
        if conn: conn.close()
    return counts
# --- FIN get_counts_by_preparador ---

# --- Función generate_qr_code (CORREGIDA) ---
def generate_qr_code(data_string, filename):
    """Genera una imagen QR y la guarda."""
    try:
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=6, # Tamaño de los cuadrados
            border=2    # Borde alrededor (mínimo 2)
        )
        qr.add_data(data_string)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        img.save(filename)
        print(f"QR code guardado como '{filename}'")
        return True
    except Exception as e:
        print(f"Error generando QR code: {e}")
        # Corregido: Eliminar texto extraño de la f-string
        show_error("Error QR", f"No se pudo generar el código QR:\n{e}")
        return False
# --- FIN generate_qr_code ---


# --- Clase Principal de la Aplicación ---
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Generador de Reportes PDF v2.10") # Incrementar versión
        self.geometry("850x850")
        ctk.set_appearance_mode("System"); ctk.set_default_color_theme("blue"); init_db()

        # Variables de estado
        self.image_file_paths = []
        self.output_pdf_path_str = ""
        self.last_generated_pdf_path = None
        self.stats_window = None

        # Variables Tkinter
        self.fecha_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d')); self.cliente_var = tk.StringVar(value=""); self.paciente_var = tk.StringVar(value=""); self.medico_var = tk.StringVar(value=""); self.tecnico_var = tk.StringVar(value=""); self.tipo_cirugia_var = tk.StringVar(value=""); self.lugar_var = tk.StringVar(value=""); self.enc_prep_var = tk.StringVar(value=""); self.enc_log_var = tk.StringVar(value=""); self.coord_cx_var = tk.StringVar(value="")

        # --- Crear Widgets ---
        main_frame = ctk.CTkFrame(self); main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Fila 0: Fecha
        ctk.CTkLabel(main_frame, text="Fecha cirugía:").grid(row=0, column=0, padx=10, pady=(10,5), sticky="w"); self.entry_fecha = DateEntry(main_frame, width=15, date_pattern='yyyy-mm-dd', textvariable=self.fecha_var, borderwidth=2, state='readonly', selectbackground='gray50', selectforeground='white', font=('Segoe UI', 11)); self.entry_fecha.grid(row=0, column=1, columnspan=2, padx=10, pady=(10,5), sticky="w")
        # Fila 1: Cliente
        ctk.CTkLabel(main_frame, text="CLIENTE:").grid(row=1, column=0, padx=10, pady=5, sticky="w"); self.combo_cliente = ctk.CTkComboBox(main_frame, values=get_suggestions('cliente'), variable=self.cliente_var, command=None); self.combo_cliente.grid(row=1, column=1, columnspan=2, padx=10, pady=5, sticky="ew")
        # Fila 2: Paciente
        ctk.CTkLabel(main_frame, text="Paciente:").grid(row=2, column=0, padx=10, pady=5, sticky="w"); self.entry_paciente = ctk.CTkEntry(main_frame, textvariable=self.paciente_var); self.entry_paciente.grid(row=2, column=1, columnspan=2, padx=10, pady=5, sticky="ew")
        # Fila 3: Médico
        ctk.CTkLabel(main_frame, text="Médico:").grid(row=3, column=0, padx=10, pady=5, sticky="w"); self.combo_medico = ctk.CTkComboBox(main_frame, values=get_suggestions('medico'), variable=self.medico_var, command=None); self.combo_medico.grid(row=3, column=1, columnspan=2, padx=10, pady=5, sticky="ew")
        # Fila 4: Técnico
        ctk.CTkLabel(main_frame, text="Técnico:").grid(row=4, column=0, padx=10, pady=5, sticky="w"); self.combo_tecnico = ctk.CTkComboBox(main_frame, values=get_suggestions('tecnico'), variable=self.tecnico_var, command=None); self.combo_tecnico.grid(row=4, column=1, columnspan=2, padx=10, pady=5, sticky="ew")
        # Fila 5: Tipo Cirugía
        ctk.CTkLabel(main_frame, text="Tipo cirugía:").grid(row=5, column=0, padx=10, pady=5, sticky="w"); self.combo_tipo_cirugia = ctk.CTkComboBox(main_frame, values=get_suggestions('tipo_cirugia'), variable=self.tipo_cirugia_var, command=None); self.combo_tipo_cirugia.grid(row=5, column=1, columnspan=2, padx=10, pady=5, sticky="ew")
        # Fila 6: Lugar
        ctk.CTkLabel(main_frame, text="Lugar:").grid(row=6, column=0, padx=10, pady=5, sticky="w"); self.combo_lugar = ctk.CTkComboBox(main_frame, values=get_suggestions('lugar'), variable=self.lugar_var, command=None); self.combo_lugar.grid(row=6, column=1, columnspan=2, padx=10, pady=5, sticky="ew")
        # Fila 7: Obs Generales
        ctk.CTkLabel(main_frame, text="Obs. Generales:").grid(row=7, column=0, padx=10, pady=5, sticky="nw"); self.text_obs_gen = ctk.CTkTextbox(main_frame, height=60); self.text_obs_gen.grid(row=7, column=1, columnspan=2, padx=10, pady=5, sticky="ew")
        # Separador 1
        ctk.CTkFrame(main_frame, height=2, fg_color="gray50").grid(row=8, column=0, columnspan=3, padx=10, pady=10, sticky="ew")
        # Fila 9: Enc Preparación
        ctk.CTkLabel(main_frame, text="Enc. Preparación:").grid(row=9, column=0, padx=10, pady=5, sticky="w"); self.combo_enc_prep = ctk.CTkComboBox(main_frame, values=get_suggestions('encargado_preparacion'), variable=self.enc_prep_var, command=None); self.combo_enc_prep.grid(row=9, column=1, columnspan=2, padx=10, pady=5, sticky="ew")
        # Fila 10: Enc Logística
        ctk.CTkLabel(main_frame, text="Enc. Logística:").grid(row=10, column=0, padx=10, pady=5, sticky="w"); self.combo_enc_log = ctk.CTkComboBox(main_frame, values=get_suggestions('encargado_logistica'), variable=self.enc_log_var, command=None); self.combo_enc_log.grid(row=10, column=1, columnspan=2, padx=10, pady=5, sticky="ew")
        # Fila 11: Coord CX
        ctk.CTkLabel(main_frame, text="Coord. CX:").grid(row=11, column=0, padx=10, pady=5, sticky="w"); self.combo_coord_cx = ctk.CTkComboBox(main_frame, values=get_suggestions('coordinador_cx'), variable=self.coord_cx_var, command=None); self.combo_coord_cx.grid(row=11, column=1, columnspan=2, padx=10, pady=5, sticky="ew")
        # Fila 12: Obs Logística
        ctk.CTkLabel(main_frame, text="Obs. Logística:").grid(row=12, column=0, padx=10, pady=5, sticky="nw"); self.text_obs_log = ctk.CTkTextbox(main_frame, height=60); self.text_obs_log.grid(row=12, column=1, columnspan=2, padx=10, pady=5, sticky="ew")
        # Separador 2
        ctk.CTkFrame(main_frame, height=2, fg_color="gray50").grid(row=13, column=0, columnspan=3, padx=10, pady=10, sticky="ew")
        # Fila 14: Imágenes
        ctk.CTkLabel(main_frame, text="Imágenes:").grid(row=14, column=0, padx=10, pady=5, sticky="w"); self.entry_images = ctk.CTkEntry(main_frame, placeholder_text="Ninguna", state="readonly", width=10); self.entry_images.grid(row=14, column=1, padx=(10, 5), pady=5, sticky="ew"); self.button_browse = ctk.CTkButton(main_frame, text="Buscar...", width=80, command=self.browse_images); self.button_browse.grid(row=14, column=2, padx=(0, 10), pady=5, sticky="e")
        # Fila 15: Guardar Como
        ctk.CTkLabel(main_frame, text="Guardar PDF como:").grid(row=15, column=0, padx=10, pady=5, sticky="w"); self.entry_output = ctk.CTkEntry(main_frame, placeholder_text="Seleccione", state="readonly", width=10); self.entry_output.grid(row=15, column=1, padx=(10, 5), pady=5, sticky="ew"); self.button_save_as = ctk.CTkButton(main_frame, text="Guardar Como...", width=120, command=self.select_save_path); self.button_save_as.grid(row=15, column=2, padx=(0, 10), pady=5, sticky="e")
        # Fila 16: Botones Acción
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent"); button_frame.grid(row=16, column=0, columnspan=3, padx=10, pady=15);
        self.button_generate = ctk.CTkButton(button_frame, text="Generar PDF", command=self.generate_pdf); self.button_generate.pack(side="left", padx=5)
        self.button_print = ctk.CTkButton(button_frame, text="Imprimir Último", command=self.print_pdf, state="disabled"); self.button_print.pack(side="left", padx=5)
        self.button_email = ctk.CTkButton(button_frame, text="Email Último", command=self.share_email, state="disabled"); self.button_email.pack(side="left", padx=5)
        self.button_whatsapp = ctk.CTkButton(button_frame, text="WhatsApp Último", command=self.share_whatsapp, state="disabled"); self.button_whatsapp.pack(side="left", padx=5)
        self.button_stats = ctk.CTkButton(button_frame, text="Ver Registros", command=self.show_stats_window); self.button_stats.pack(side="left", padx=5)

        main_frame.columnconfigure(1, weight=1)

    # --- Métodos ---
    # (browse_images, sanitize_filename, select_save_path, generate_pdf, regenerate_pdf_from_data,
    #  update_suggestions, print_pdf, share_email, share_whatsapp, show_stats_window, load_data_into_form)
    # --- Pegar aquí TODAS las definiciones de métodos de la clase App ---
    # --- (Incluyendo generate_pdf y regenerate_pdf_from_data con los finally corregidos) ---
    def browse_images(self):
        files = filedialog.askopenfilenames(title="Seleccionar Imágenes", filetypes=(("Image Files", "*.png *.jpg *.jpeg *.bmp *.gif"), ("All Files", "*.*")))
        if files: self.image_file_paths = list(files); display_text = f"{len(self.image_file_paths)} imágenes"; self.entry_images.configure(state="normal"); self.entry_images.delete(0, "end"); self.entry_images.insert(0, display_text); self.entry_images.configure(state="readonly"); print(f"Imágenes: {self.image_file_paths}")
        else: self.image_file_paths = []; self.entry_images.configure(state="normal"); self.entry_images.delete(0, "end"); self.entry_images.insert(0, "Ninguna"); self.entry_images.configure(state="readonly")

    def sanitize_filename(self, name):
        valid_chars = "-_.() abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"; sanitized = ''.join(c for c in name if c in valid_chars); sanitized = sanitized.replace(' ', '_'); return sanitized.strip('_') or "_"

    def select_save_path(self, suggested_filename=None):
        if not suggested_filename:
            paciente = self.paciente_var.get() or "Paciente"; cliente = self.cliente_var.get() or "Cliente"; fecha = self.fecha_var.get() or datetime.now().strftime('%Y%m%d')
            suggested_filename = f"{self.sanitize_filename(paciente)}-{self.sanitize_filename(cliente)}-{self.sanitize_filename(fecha)}.pdf"
        file_path = filedialog.asksaveasfilename(title="Guardar PDF Como", defaultextension=".pdf", initialfile=suggested_filename, filetypes=(("PDF Files", "*.pdf"),))
        if file_path: self.output_pdf_path_str = file_path; self.entry_output.configure(state="normal"); self.entry_output.delete(0, "end"); self.entry_output.insert(0, self.output_pdf_path_str); self.entry_output.configure(state="readonly"); print(f"Salida: {self.output_pdf_path_str}"); return True
        else: self.output_pdf_path_str = ""; self.entry_output.configure(state="normal"); self.entry_output.delete(0, "end"); self.entry_output.insert(0, "Seleccione"); self.entry_output.configure(state="readonly"); return False

    # --- generate_pdf (CON FINALLY CORREGIDO) ---
    def generate_pdf(self):
        print("Generando PDF..."); self.last_generated_pdf_path = None
        self.button_print.configure(state="disabled"); self.button_email.configure(state="disabled"); self.button_whatsapp.configure(state="disabled")
        context = { 'fecha_cirugia': self.fecha_var.get(), 'cliente': self.cliente_var.get(), 'paciente': self.paciente_var.get(), 'medico': self.medico_var.get(), 'tecnico': self.tecnico_var.get(), 'tipo_cirugia': self.tipo_cirugia_var.get(), 'lugar': self.lugar_var.get(), 'observaciones_generales': self.text_obs_gen.get("1.0", "end-1c").strip(), 'encargado_preparacion': self.enc_prep_var.get(), 'encargado_logistica': self.enc_log_var.get(), 'coordinador_cx': self.coord_cx_var.get(), 'observaciones_logistica': self.text_obs_log.get("1.0", "end-1c").strip() }
        if not self.select_save_path(): show_error("Cancelado", "Seleccione ubicación."); return
        if not context['fecha_cirugia'] or not context['paciente']: show_error("Requerido", "Complete fecha y paciente."); return
        try: datetime.strptime(context['fecha_cirugia'], '%Y-%m-%d')
        except ValueError: show_error("Inválido", "Fecha AAAA-MM-DD."); return
        record_id = str(uuid.uuid4()); unique_id_for_qr = str(uuid.uuid4()); context['unique_id'] = unique_id_for_qr
        qr_data = f"ID_Doc: {unique_id_for_qr} | Pac: {context['paciente']} | Fec: {context['fecha_cirugia']}"; qr_temp_path = Path(QR_TEMP_FILENAME);
        if not generate_qr_code(qr_data, qr_temp_path): return
        doc = None; temp_docx_path = Path(TEMP_DOCX_FILENAME); render_success = False; pdf_conversion_success = False
        try:
            template_path = Path(TEMPLATE_FILENAME);
            if not template_path.exists(): raise FileNotFoundError(f"No '{TEMPLATE_FILENAME}'.")
            doc = DocxTemplate(template_path); context['qr_code_image'] = InlineImage(doc, str(qr_temp_path), width=Mm(25))
            for i, img_path_str in enumerate(self.image_file_paths):
                if i >= MAX_IMAGES_IN_TEMPLATE: print(f"WARN: Máx {MAX_IMAGES_IN_TEMPLATE} imgs."); break
                img_path = Path(img_path_str.strip()); img_key = f'imagen_{i+1}'
                if img_path.exists():
                    try: context[img_key] = InlineImage(doc, str(img_path), width=Mm(160))
                    except Exception as e: print(f"WARN: Proc. img '{img_path.name}' {img_key}: {e}")
                else: print(f"WARN: Ruta img inválida: '{img_path_str}'")
            print("Renderizando..."); doc.render(context); print("Guardando DOCX..."); doc.save(temp_docx_path); render_success = True
            output_pdf_path = Path(self.output_pdf_path_str); output_dir = output_pdf_path.parent
            if convert_to_pdf(temp_docx_path, output_dir):
                gen_pdf = output_dir / temp_docx_path.with_suffix('.pdf').name
                if gen_pdf.exists():
                    try:
                        if gen_pdf.resolve() != output_pdf_path.resolve(): gen_pdf.rename(output_pdf_path)
                        pdf_conversion_success = True; self.last_generated_pdf_path = str(output_pdf_path)
                        show_info("Éxito", f"PDF: {self.last_generated_pdf_path}")
                        self.button_print.configure(state="normal"); self.button_email.configure(state="normal"); self.button_whatsapp.configure(state="normal")
                    except Exception as e: show_error("Error", f"PDF ok, error renombrar:\n{e}")
                else: show_error("Error", f"Conversión OK, no se encontró '{gen_pdf.name}'.")
        except FileNotFoundError as e: show_error("Error Fatal", str(e))
        except Exception as e: show_error("Error Fatal Generación", f"Error inesperado:\n{e}")
        finally: # Correctamente indentado
            print("Ejecutando finally (generación)...")
            if qr_temp_path.exists():
                try: qr_temp_path.unlink(); print("Temp QR eliminado.")
                except Exception as e_qr: print(f"WARN: No eliminar QR: {e_qr}")
            if temp_docx_path.exists():
                try: temp_docx_path.unlink(); print("Temp DOCX eliminado.")
                except Exception as e_docx: print(f"WARN: No eliminar DOCX: {e_docx}")
        if pdf_conversion_success:
            record = context.copy(); record['id'] = record_id; record['unique_id'] = context['unique_id']
            record['fecha_generacion'] = datetime.now().isoformat(timespec='seconds'); record['archivo_pdf'] = self.last_generated_pdf_path
            keys_to_remove = [k for k in record if isinstance(record[k], InlineImage)]
            for key in keys_to_remove: record.pop(key, None)
            if save_record(record): self.update_suggestions()

    # --- Regenerar PDF (CON FINALLY CORREGIDO) ---
    def regenerate_pdf_from_data(self, data):
        print(f"Regenerando PDF para ID_DB: {data.get('id', 'N/A')}, UniqueID: {data.get('unique_id', 'N/A')}")
        self.last_generated_pdf_path = None
        self.button_print.configure(state="disabled"); self.button_email.configure(state="disabled"); self.button_whatsapp.configure(state="disabled")
        context = data.copy()
        original_unique_id = context.get('unique_id'); original_record_id = context.get('id')
        if not original_unique_id or not original_record_id: show_error("Error", "Datos importados sin IDs."); return
        sugg_paciente=self.sanitize_filename(context.get('paciente','P')); sugg_cliente=self.sanitize_filename(context.get('cliente','C')); sugg_fecha=self.sanitize_filename(context.get('fecha_cirugia',''))
        suggested_filename = f"{sugg_paciente}-{sugg_cliente}-{sugg_fecha}_REGENERADO.pdf"
        if not self.select_save_path(suggested_filename=suggested_filename): show_error("Cancelado", "Seleccione ubicación para regenerar."); return
        qr_data = f"ID_Doc: {original_unique_id} | Pac: {context['paciente']} | Fec: {context['fecha_cirugia']}"
        qr_temp_path = Path(f"{QR_TEMP_FILENAME.stem}_regen_{uuid.uuid4().hex[:6]}.png")
        if not generate_qr_code(qr_data, qr_temp_path): return
        doc = None; temp_docx_path = Path(f"{TEMP_DOCX_FILENAME.stem}_regen_{uuid.uuid4().hex[:6]}.docx")
        render_success = False; pdf_conversion_success = False
        try:
            template_path = Path(TEMPLATE_FILENAME);
            if not template_path.exists(): raise FileNotFoundError(f"No '{TEMPLATE_FILENAME}'.")
            doc = DocxTemplate(template_path)
            context['qr_code_image'] = InlineImage(doc, str(qr_temp_path), width=Mm(25))
            context['unique_id'] = original_unique_id
            print("WARN: Regeneración no incluye imágenes originales."); keys_to_remove = [k for k in context if k.startswith('imagen_')]; [context.pop(k, None) for k in keys_to_remove]
            print("Renderizando (regen)..."); doc.render(context); print(f"Guardando DOCX (regen): {temp_docx_path}"); doc.save(temp_docx_path); render_success = True
            output_pdf_path = Path(self.output_pdf_path_str); output_dir = output_pdf_path.parent
            if convert_to_pdf(temp_docx_path, output_dir):
                gen_pdf = output_dir / temp_docx_path.with_suffix('.pdf').name
                if gen_pdf.exists():
                    try:
                        if gen_pdf.resolve() != output_pdf_path.resolve(): gen_pdf.rename(output_pdf_path)
                        pdf_conversion_success = True; self.last_generated_pdf_path = str(output_pdf_path)
                        show_info("Éxito", f"PDF REGENERADO: {self.last_generated_pdf_path}")
                        self.button_print.configure(state="normal"); self.button_email.configure(state="normal"); self.button_whatsapp.configure(state="normal")
                    except Exception as e: show_error("Error", f"PDF ok, error renombrar:\n{e}")
                else: show_error("Error", f"Conversión OK, no se encontró '{gen_pdf.name}'.")
        except FileNotFoundError as e: show_error("Error Fatal", str(e))
        except Exception as e: show_error("Error Fatal Regeneración", f"Error inesperado:\n{e}")
        finally: # Correctamente indentado
            print("Ejecutando finally (regeneración)...")
            if qr_temp_path.exists():
                try: qr_temp_path.unlink(); print("Temp QR (regen) eliminado.")
                except Exception as e: print(f"WARN: No eliminar QR (regen): {e}")
            if temp_docx_path.exists():
                try: temp_docx_path.unlink(); print("Temp DOCX (regen) eliminado.")
                except Exception as e: print(f"WARN: No eliminar DOCX (regen): {e}")

    # --- Otros Métodos (update_suggestions, print_pdf, share_email, share_whatsapp, show_stats_window, load_data_into_form) ---
    # ... (Sin cambios) ...
# --- Pega ESTA función completa en lugar de la existente ---
    def update_suggestions(self):
        """Actualiza las listas de sugerencias en los Combobox."""
        print("Actualizando sugerencias...")
        # Cada configure en su propia línea con indentación correcta
        self.combo_medico.configure(values=get_suggestions('medico'))
        self.combo_cliente.configure(values=get_suggestions('cliente'))
        self.combo_tecnico.configure(values=get_suggestions('tecnico'))
        self.combo_tipo_cirugia.configure(values=get_suggestions('tipo_cirugia'))
        self.combo_lugar.configure(values=get_suggestions('lugar'))
        self.combo_enc_prep.configure(values=get_suggestions('encargado_preparacion'))
        self.combo_enc_log.configure(values=get_suggestions('encargado_logistica'))
        self.combo_coord_cx.configure(values=get_suggestions('coordinador_cx'))
    # --- FIN DE LA FUNCIÓN update_suggestions ---    def print_pdf(self):
        if not self.last_generated_pdf_path or not Path(self.last_generated_pdf_path).exists(): show_error("Error", "PDF no generado."); return; print(f"Imprimiendo: {self.last_generated_pdf_path}");
        try:
            if sys.platform == "win32": os.startfile(self.last_generated_pdf_path, "print")
            elif sys.platform == "darwin": subprocess.run(["lpr", self.last_generated_pdf_path], check=True)
            else: subprocess.run(["lp", self.last_generated_pdf_path], check=True)
            show_info("Imprimir", "Enviado.")
        except Exception as e: show_error("Error", f"No imprimir:\n{e}")
    def share_email(self):
        if not self.last_generated_pdf_path: show_error("Error", "Primero genera PDF."); return
        try: subject = f"Reporte - {self.paciente_var.get()}"; body = f"Adjunto reporte {self.paciente_var.get()} ({self.fecha_var.get()}). Archivo: {self.last_generated_pdf_path} (Adjuntar)"; mailto_url = f"mailto:?subject={quote(subject)}&body={quote(body)}"; webbrowser.open(mailto_url); show_info("Email", "Abriendo cliente...\n¡Adjunta PDF!")
        except Exception as e: show_error("Error", f"No abrir cliente correo:\n{e}")
    def share_whatsapp(self):
         if not self.last_generated_pdf_path: show_error("Error", "Primero genera PDF."); return
         try: phone = ""; message = f"Reporte {self.paciente_var.get()} ({self.fecha_var.get()}). Archivo: {self.last_generated_pdf_path} (Adjuntar)"; whatsapp_url = f"https://wa.me/{phone}?text={quote(message)}"; print(f"Abriendo WA: {whatsapp_url[:100]}..."); webbrowser.open(whatsapp_url); show_info("WhatsApp", "Abriendo WA...\n¡Adjunta PDF!")
         except Exception as e: show_error("Error", f"No abrir WA:\n{e}")
    def show_stats_window(self):
        if not self.stats_window or not self.stats_window.winfo_exists(): self.stats_window = StatsWindow(self); self.stats_window.grab_set()
        else: self.stats_window.focus()
    def load_data_into_form(self, data):
        print("Importando datos...");
        if not data or not isinstance(data, dict): show_error("Error", "Datos inválidos."); return
        self.fecha_var.set(data.get('fecha_cirugia', '')); self.cliente_var.set(data.get('cliente', '')); self.paciente_var.set(data.get('paciente', ''));
        self.medico_var.set(data.get('medico', '')); self.tecnico_var.set(data.get('tecnico', '')); self.tipo_cirugia_var.set(data.get('tipo_cirugia', '')); self.lugar_var.set(data.get('lugar', '')); self.enc_prep_var.set(data.get('encargado_preparacion', '')); self.enc_log_var.set(data.get('encargado_logistica', '')); self.coord_cx_var.set(data.get('coordinador_cx', ''))
        self.text_obs_gen.delete("1.0", "end"); self.text_obs_gen.insert("1.0", data.get('observaciones_generales', ''))
        self.text_obs_log.delete("1.0", "end"); self.text_obs_log.insert("1.0", data.get('observaciones_logistica', ''))
        self.image_file_paths = []; self.entry_images.configure(state="normal"); self.entry_images.delete(0, "end"); self.entry_images.insert(0, "Ninguna (Importado)"); self.entry_images.configure(state="readonly")
        self.output_pdf_path_str = ""; self.entry_output.configure(state="normal"); self.entry_output.delete(0, "end"); self.entry_output.insert(0, "Seleccione"); self.entry_output.configure(state="readonly")
        self.last_generated_pdf_path = None; self.button_print.configure(state="disabled"); self.button_email.configure(state="disabled"); self.button_whatsapp.configure(state="disabled")
        show_info("Importado", f"Datos cargados: {data.get('paciente', 'N/A')}")


# --- Clase StatsWindow ---
# (Sin cambios respecto a la última versión completa)
class StatsWindow(ctk.CTkToplevel):
   def __init__(self, master=None):
       super().__init__(master); self.master_app = master
       self.title("Registros y Estadísticas"); self.geometry("1050x700")
       # (Definición COMPLETA de filtros, tabla, stats_text y botones Importar/Regenerar)
       filter_frame = ctk.CTkFrame(self); filter_frame.pack(pady=10, padx=10, fill="x");
       ctk.CTkLabel(filter_frame, text="Desde:").grid(row=0, column=0, padx=(10,5), pady=5, sticky="w"); self.date_from = DateEntry(filter_frame, width=12, date_pattern='yyyy-mm-dd', state='readonly'); self.date_from.grid(row=0, column=1, padx=5, pady=5); self.date_from.set_date(None)
       ctk.CTkLabel(filter_frame, text="Hasta:").grid(row=0, column=2, padx=(10,5), pady=5, sticky="w"); self.date_to = DateEntry(filter_frame, width=12, date_pattern='yyyy-mm-dd', state='readonly'); self.date_to.grid(row=0, column=3, padx=5, pady=5); self.date_to.set_date(None)
       ctk.CTkLabel(filter_frame, text="Médico:").grid(row=1, column=0, padx=(10,5), pady=5, sticky="w"); self.medico_filter = ctk.CTkComboBox(filter_frame, values=["Todos"] + get_suggestions('medico'), width=180); self.medico_filter.grid(row=1, column=1, padx=5, pady=5); self.medico_filter.set("Todos")
       ctk.CTkLabel(filter_frame, text="Paciente:").grid(row=1, column=2, padx=(10,5), pady=5, sticky="w"); self.paciente_filter = ctk.CTkEntry(filter_frame, placeholder_text="Buscar nombre...", width=180); self.paciente_filter.grid(row=1, column=3, padx=5, pady=5)
       ctk.CTkLabel(filter_frame, text="Cliente:").grid(row=2, column=0, padx=(10,5), pady=5, sticky="w"); self.cliente_filter = ctk.CTkEntry(filter_frame, placeholder_text="Buscar cliente...", width=180); self.cliente_filter.grid(row=2, column=1, padx=5, pady=5)
       ctk.CTkLabel(filter_frame, text="ID QR (exacto):").grid(row=2, column=2, padx=(10,5), pady=5, sticky="w"); self.unique_id_filter = ctk.CTkEntry(filter_frame, placeholder_text="Pegar ID QR...", width=180); self.unique_id_filter.grid(row=2, column=3, padx=5, pady=5)
       search_button = ctk.CTkButton(filter_frame, text="Buscar Registros", command=self.load_stats); search_button.grid(row=3, column=0, columnspan=4, padx=10, pady=10)
       results_frame = ctk.CTkFrame(self); results_frame.pack(pady=10, padx=10, fill="both", expand=True); style = ttk.Style(); style.theme_use("clam")
       columns = ('id', 'fecha_gen', 'fecha_cx', 'paciente', 'medico', 'tipo', 'lugar', 'enc_prep','unique_id_col', 'pdf_path'); self.tree = ttk.Treeview(results_frame, columns=columns, show='headings');
       self.tree.heading('id', text='ID_DB'); self.tree.heading('fecha_gen', text='Generado'); self.tree.heading('fecha_cx', text='Fecha Cx'); self.tree.heading('paciente', text='Paciente'); self.tree.heading('medico', text='Médico'); self.tree.heading('tipo', text='Tipo Cx'); self.tree.heading('lugar', text='Lugar'); self.tree.heading('enc_prep', text='Enc Prep'); self.tree.heading('unique_id_col', text='ID QR'); self.tree.heading('pdf_path', text='Archivo PDF')
       self.tree.column('id', width=0, stretch=tk.NO); self.tree.column('fecha_gen', width=120); self.tree.column('fecha_cx', width=90); self.tree.column('paciente', width=140); self.tree.column('medico', width=140); self.tree.column('tipo', width=120); self.tree.column('lugar', width=120); self.tree.column('enc_prep', width=120); self.tree.column('unique_id_col', width=160); self.tree.column('pdf_path', width=180)
       vsb = ttk.Scrollbar(results_frame, orient="vertical", command=self.tree.yview); hsb = ttk.Scrollbar(results_frame, orient="horizontal", command=self.tree.xview); self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set); vsb.pack(side='right', fill='y'); hsb.pack(side='bottom', fill='x'); self.tree.pack(side='left', fill='both', expand=True)
       self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
       bottom_frame = ctk.CTkFrame(self); bottom_frame.pack(pady=10, padx=10, fill="x");
       stats_label = ctk.CTkLabel(bottom_frame, text="Cajas x Enc:"); stats_label.pack(side="left", padx=5, pady=5); self.stats_text = ctk.CTkTextbox(bottom_frame, height=60, width=300, state="disabled"); self.stats_text.pack(side="left", padx=5, pady=5, fill="x", expand=True)
       self.regenerate_button = ctk.CTkButton(bottom_frame, text="Regenerar PDF", command=self.regenerate_selected_pdf, state="disabled"); self.regenerate_button.pack(side="right", padx=(5,10), pady=10)
       self.import_button = ctk.CTkButton(bottom_frame, text="Importar Datos", command=self.import_selected_record, state="disabled"); self.import_button.pack(side="right", padx=(0,5), pady=10)
       self.load_stats()

   def on_tree_select(self, event=None):
       is_one = len(self.tree.selection()) == 1; state = "normal" if is_one else "disabled"; self.import_button.configure(state=state); self.regenerate_button.configure(state=state)

def load_stats(self):
    print("Cargando stats...")
    # Limpiar tabla
    for item in self.tree.get_children():
        self.tree.delete(item)
    self.stats_text.configure(state="normal")
    self.stats_text.delete("1.0", "end")
    self.stats_text.configure(state="disabled")
    self.import_button.configure(state="disabled")
    self.regenerate_button.configure(state="disabled")

    # Filtros
    d_from = self.date_from.get_date().strftime('%Y-%m-%d') if self.date_from.get_date() else None
    d_to = self.date_to.get_date().strftime('%Y-%m-%d') if self.date_to.get_date() else None
    med = self.medico_filter.get() if self.medico_filter.get() != "Todos" else None
    pac = self.paciente_filter.get().strip()
    cli = self.cliente_filter.get().strip()
    uid = self.unique_id_filter.get().strip()

    # SQL base
    sql = "SELECT id, fecha_generacion, fecha_cirugia, paciente, medico, tipo_cirugia, lugar, encargado_preparacion, unique_id, archivo_pdf FROM cirugias WHERE 1=1"
    p = []

    if d_from:
        sql += " AND date(fecha_generacion) >= date(?)"
        p.append(d_from)
    if d_to:
        sql += " AND date(fecha_generacion) <= date(?)"
        p.append(d_to)
    if med:
        sql += " AND medico LIKE ?"
        p.append(f"%{med}%")
    if pac:
        sql += " AND paciente LIKE ?"
        p.append(f"%{pac}%")
    if cli:
        sql += " AND cliente LIKE ?"
        p.append(f"%{cli}%")
    if uid:
        sql += " AND unique_id = ?"
        p.append(uid)
    sql += " ORDER BY fecha_generacion DESC"

    # Inicializar results
    results = []

    conn = None
    try:
        conn = sqlite3.connect(DB_FILENAME)
        cursor = conn.cursor()
        print(f"SQL ejecutado:\n{sql}\nParámetros: {p}")
        cursor.execute(sql, p)
        results = cursor.fetchall()
    except sqlite3.Error as e:
        print(f"Error cargando stats: {e}")
        show_error("Error", f"Error cargando registros de la base de datos:\n{e}")
    finally:
        if conn:
            conn.close()

    if not results:
        print("No se encontraron resultados.")
        show_info("Resultados", "No se encontraron registros con los filtros seleccionados.")
        return

    # Insertar resultados en la tabla
    for r in results:
        dt_gen = r[1]
        try:
            dt_gen = datetime.fromisoformat(r[1]).strftime('%Y-%m-%d %H:%M')
        except (ValueError, TypeError):
            pass

        display_row = (r[0], dt_gen, r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9])
        self.tree.insert("", tk.END, values=display_row)

    # Cargar estadísticas
    counts = get_counts_by_preparador()
    stats_txt = "\n".join([f"{p or 'N/A'}: {c}" for p, c in counts]) or "No datos."
    self.stats_text.configure(state="normal")
    self.stats_text.delete("1.0", "end")
    self.stats_text.insert("1.0", stats_txt)
    self.stats_text.configure(state="disabled")


def import_selected_record(self):
    sel = self.tree.selection()
    if len(sel) != 1:
        return
    try:
        rid = self.tree.item(sel[0], 'values')[0]
        print(f"Importando: {rid}")
    except IndexError:
        show_error("Error", "No se pudo obtener ID de la fila.")
        return

    if rid and self.master_app:
        data = get_record_by_id(rid)
        if data:
            self.master_app.load_data_into_form(data)
            self.destroy()
        else:
            show_error("Error", f"No datos para ID: {rid}")
    else:
        show_error("Error", "No ID seleccionado o App principal no disponible.")

# --- Ejecutar ---
if __name__ == "__main__":
    # Establecer DPI Awareness (importante en Windows para UI nítida)
    if sys.platform == "win32":
        try:
            from ctypes import windll
            try:
                windll.shcore.SetProcessDpiAwareness(2)  # DPI_AWARENESS_PER_MONITOR_AWARE_V2
                print("DPI Awareness set (V2).")
            except AttributeError:
                try:
                    windll.shcore.SetProcessDpiAwareness(1)  # DPI_AWARENESS_PER_MONITOR_AWARE
                    print("DPI Awareness set (V1).")
                except AttributeError:
                    windll.user32.SetProcessDPIAware()  # Método antiguo
                    print("DPI Awareness set (Legacy).")
        except Exception as e:
            print(f"WARN: No se pudo establecer DPI Awareness: {e}")

    app = App()
    app.mainloop()
    print("App cerrada.")

