# -*- coding: utf-8 -*-

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import DateEntry
import os
import subprocess
import sys
import threading # Para hilos
import time # Para simular trabajo o esperar
import traceback # Para imprimir errores completos
import json # Para leer/escribir config.json
from pathlib import Path
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm, Inches
from datetime import datetime
import sqlite3
import uuid
from PIL import Image as PILImage
import webbrowser
from urllib.parse import quote

# --- Configuración ---
CONFIG_FILENAME = "config.json"
TEMPLATE_FILENAME = "template.docx"
DB_FILENAME = "registros_cirugias.db"
LIBREOFFICE_PATH = "C:/Program Files/LibreOffice/program/soffice.exe" # Ajusta si es necesario
MAX_IMAGES_ALLOWED = 100
CONVERT_TIMEOUT = 300 # 5 minutos

# --- Constantes de Calidad/Layout ---
IMG_QUALITY_HIGH = 90
IMG_QUALITY_MEDIUM = 75
IMG_QUALITY_LOW = 60
# VOLVIENDO A 2 COLUMNAS
LAYOUT_2_PER_ROW_WIDTH_MM = 79  # Ancho para 2 imágenes por fila (Ajustar según márgenes)

# --- Claves para config.json ---
CONFIG_KEY_IMG_DIR = "default_image_dir"
CONFIG_KEY_OUTPUT_DIR = "default_output_dir"

# --- Variables Globales para Comunicación de Errores entre Hilos ---
last_conversion_error = ""
last_db_error = ""

# --- Funciones Auxiliares ---

def load_config():
    """Carga la configuración desde config.json, crea defaults si no existe."""
    config_path = Path(CONFIG_FILENAME)
    defaults = {
        CONFIG_KEY_IMG_DIR: str(Path.home() / "Pictures"),
        CONFIG_KEY_OUTPUT_DIR: str(Path.home() / "Documents")
    }
    if config_path.exists():
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f); defaults.update(config); return defaults
        except (json.JSONDecodeError, IOError) as e: print(f"WARN: Error cargando {CONFIG_FILENAME}: {e}. Usando defaults."); return defaults
    else: print(f"INFO: Archivo {CONFIG_FILENAME} no encontrado. Creando con defaults."); save_config(defaults); return defaults

def save_config(config_data):
    """Guarda la configuración en config.json."""
    config_path = Path(CONFIG_FILENAME)
    try:
        with open(config_path, 'w', encoding='utf-8') as f: json.dump(config_data, f, indent=4)
        print(f"INFO: Configuración guardada en {config_path.resolve()}"); return True
    except IOError as e: print(f"ERROR: No se pudo guardar config en {config_path}: {e}"); return False

def compress_image(img_path, target_width_mm, quality):
    """Comprime y redimensiona imágenes. Devuelve Path del archivo temporal o None si falla."""
    try:
        img_path = Path(img_path); print(f"Comprimiendo: {img_path.name} (Q: {quality}, W: {target_width_mm}mm)")
        img = PILImage.open(img_path)
        if img.mode in ('RGBA', 'P'): img = img.convert('RGB')
        target_width_px = int(target_width_mm / 25.4 * 200)
        if img.width > target_width_px:
            ratio = target_width_px / img.width; target_height_px = int(img.height * ratio)
            print(f"  - Redimensionando a {target_width_px}x{target_height_px}"); img = img.resize((target_width_px, target_height_px), PILImage.LANCZOS)
        temp_suffix = f"_comp_{uuid.uuid4().hex[:8]}.jpg"; temp_path = img_path.with_name(img_path.stem + temp_suffix)
        img.save(temp_path, "JPEG", quality=quality, optimize=True, progressive=True); print(f"  - Guardado temp: {temp_path.name}"); return temp_path
    except FileNotFoundError: print(f"ERROR compress: Imagen no encontrada {img_path}"); return None
    except Exception as e: print(f"ERROR compress: {img_path.name}: {e}"); return None

def find_libreoffice():
    """Busca el ejecutable de LibreOffice."""
    manual_path = Path(LIBREOFFICE_PATH);
    if manual_path.is_file() and manual_path.name.lower().startswith("soffice"): return str(manual_path)
    if manual_path.is_dir():
        for name in ["soffice.exe", "soffice"]:
            if (manual_path / name).exists(): found_path = manual_path / name; print(f"Usando ejecutable en dir config: {found_path}"); return str(found_path)
    print(f"WARN: Ruta '{LIBREOFFICE_PATH}' no válida. Buscando..."); paths_to_check = []
    if sys.platform == "win32": paths_to_check = [Path("C:/Program Files/LibreOffice/program/soffice.exe"), Path("C:/Program Files (x86)/LibreOffice/program/soffice.exe")]; lo_env = os.getenv("LIBREOFFICE_PROGRAM_PATH");
    if lo_env: paths_to_check.insert(0, Path(lo_env) / "soffice.exe")
    elif sys.platform == "darwin": paths_to_check = [Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")]
    else: paths_to_check = [Path("/usr/bin/soffice"), Path("/usr/local/bin/soffice"), Path("/snap/bin/libreoffice.soffice"), Path("/opt/libreoffice/program/soffice")]
    for path in paths_to_check:
        if path.exists() and path.is_file() and path.name.lower().startswith("soffice"): print(f"LO encontrado en ruta común: {path}"); return str(path)
    print("INFO: Intentando buscar 'soffice' en PATH...");
    try:
        cmd = ["where", "soffice"] if sys.platform == "win32" else ["which", "soffice"]
        res = subprocess.run(cmd, capture_output=True, text=True, check=False, timeout=5)
        if res.returncode == 0 and res.stdout: print(f"INFO: 'soffice' encontrado en PATH."); return "soffice"
        else: print(f"INFO: 'soffice' no encontrado en PATH (código: {res.returncode}).")
    except (FileNotFoundError, subprocess.TimeoutExpired, Exception) as e: print(f"INFO: No se pudo buscar en PATH: {type(e).__name__}"); pass
    print(f"ERROR CRITICO: No se encontró 'soffice'. Verifica instalación y LIBREOFFICE_PATH ('{LIBREOFFICE_PATH}')."); return LIBREOFFICE_PATH

def show_error_safe(title, message): messagebox.showerror(title, message)
def show_info_safe(title, message): messagebox.showinfo(title, message)

def convert_to_pdf(docx_path, output_dir, timeout_duration=CONVERT_TIMEOUT):
    """Convierte DOCX a PDF. Devuelve True/False. Guarda error en global."""
    global last_conversion_error; last_conversion_error = ""
    soffice_cmd_path = find_libreoffice(); soffice_executable = Path(soffice_cmd_path); can_execute = (soffice_executable.is_file() or soffice_cmd_path.lower() == "soffice")
    if not can_execute: last_conversion_error = f"Ejecutable LO no encontrado/inválido: '{soffice_cmd_path}'"; print(f"ERROR CONVERT: {last_conversion_error}"); return False
    docx_abs_path = str(Path(docx_path).resolve()); output_dir_abs_path = str(Path(output_dir).resolve())
    try:
        print(f"Convirtiendo '{docx_abs_path}' a PDF en '{output_dir_abs_path}' (Timeout: {timeout_duration}s)...")
        cmd = [ soffice_cmd_path, '--headless', '--convert-to', 'pdf', '--outdir', output_dir_abs_path, docx_abs_path ]; print(f"Comando: {' '.join(cmd)}")
        start_t = time.time(); result = subprocess.run(cmd, capture_output=True, text=True, check=False, encoding='utf-8', errors='ignore', timeout=timeout_duration); end_t = time.time()
        print(f"DEBUG: subprocess.run completado en {end_t - start_t:.2f} segundos."); print(f"LO Return Code: {result.returncode}")
        pdf_filename = Path(docx_path).with_suffix('.pdf').name; expected_pdf_path = Path(output_dir_abs_path) / pdf_filename
        if result.returncode == 0 and expected_pdf_path.exists(): print(f"Conversión PDF OK: {expected_pdf_path}"); return True
        else:
            error_msg = f"Fallo conversión PDF '{Path(docx_path).name}'. ";
            if not expected_pdf_path.exists(): error_msg += f"PDF no encontrado: {expected_pdf_path}. "
            if result.returncode != 0: error_msg += f"LO Code: {result.returncode}. "
            if result.stderr: error_msg += f"LO Err: {result.stderr[:250]}..."
            last_conversion_error = error_msg; print(f"ERROR CONVERT: {error_msg}"); return False
    except subprocess.TimeoutExpired: end_t = time.time(); print(f"DEBUG: subprocess.run TIMEOUT después de {end_t - start_t:.2f} segundos."); last_conversion_error = f"Timeout ({timeout_duration}s) convirtiendo '{Path(docx_path).name}'."; print(f"ERROR CONVERT: {last_conversion_error}"); return False
    except FileNotFoundError: last_conversion_error = f"Comando '{soffice_cmd_path}' no encontrado."; print(f"ERROR CONVERT: {last_conversion_error}"); return False
    except Exception as e: last_conversion_error = f"Error inesperado conversión: {type(e).__name__}: {e}"; print(f"ERROR CONVERT: {last_conversion_error}\n{traceback.format_exc()}"); return False

def init_db():
    """Inicializa la base de datos."""
    conn = None; db_path = Path(DB_FILENAME).resolve(); print(f"DB Init: {db_path}")
    if not db_path.parent.exists():
        try: db_path.parent.mkdir(parents=True, exist_ok=True)
        except OSError as e: print(f"FATAL: No crear dir DB: {e}"); sys.exit()
    try:
        conn = sqlite3.connect(db_path, timeout=10); cursor = conn.cursor()
        cursor.execute("CREATE TABLE IF NOT EXISTS cirugias (id TEXT PRIMARY KEY, fecha_generacion TEXT NOT NULL, archivo_pdf TEXT, fecha_cirugia TEXT, cliente TEXT, paciente TEXT, medico TEXT, tecnico TEXT, tipo_cirugia TEXT, lugar TEXT, observaciones_generales TEXT, encargado_preparacion TEXT, encargado_logistica TEXT, coordinador_cx TEXT, observaciones_logistica TEXT, unique_id TEXT UNIQUE NOT NULL)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_fecha_gen ON cirugias (fecha_generacion);"); cursor.execute("CREATE INDEX IF NOT EXISTS idx_paciente ON cirugias (paciente);"); cursor.execute("CREATE INDEX IF NOT EXISTS idx_medico ON cirugias (medico);"); cursor.execute("CREATE INDEX IF NOT EXISTS idx_unique_id ON cirugias (unique_id);")
        conn.commit(); print("DB OK.")
    except sqlite3.Error as e: print(f"FATAL: DB Init Error: {e}"); sys.exit()
    finally:
        if conn: conn.close()

def save_record(record_data):
    """Guarda registro. Devuelve True/False. Guarda error en global."""
    global last_db_error; last_db_error = ""
    conn = None; required = ['id', 'unique_id', 'fecha_generacion']
    if not all(f in record_data and record_data[f] for f in required):
        missing = [f for f in required if not record_data.get(f)]
        last_db_error = f"Faltan campos DB: {', '.join(missing)}"; print(f"ERROR DB SAVE: {last_db_error}"); return False
    db_path = Path(DB_FILENAME).resolve()
    try:
        conn = sqlite3.connect(db_path, timeout=10); cursor = conn.cursor()
        # Excluir claves que no deben ir a la BD explícitamente
        keys_to_exclude_from_db = ['image_pairs', 'fecha_emision', 'fecha_emision_corta'] # Cambiado a image_pairs
        db_data = {k: v for k, v in record_data.items() if k not in keys_to_exclude_from_db}

        fields = list(db_data.keys()); values = list(db_data.values())
        placeholders = ','.join(['?'] * len(fields)); sql = f"INSERT INTO cirugias ({','.join(fields)}) VALUES ({placeholders})"
        print(f"DEBUG DB SQL: {sql}"); print(f"DEBUG DB Values: {values}")
        cursor.execute(sql, values); conn.commit(); print(f"DB Save OK (ID: {record_data.get('id')})"); return True
    except sqlite3.IntegrityError as e:
        if "unique constraint failed: cirugias.unique_id" in str(e).lower(): last_db_error = f"ID único '{record_data.get('unique_id')}' ya existe."
        else: last_db_error = f"Error integridad DB: {e}"
        print(f"ERROR DB SAVE (Integrity): {last_db_error}"); return False
    except sqlite3.Error as e:
        last_db_error = f"Error general DB: {e}"; print(f"ERROR DB SAVE (General): {last_db_error}"); return False
    finally:
        if conn: conn.close()

def get_suggestions(field_name, limit=30):
    """Obtiene sugerencias para campos."""
    conn = None; suggestions = []; db_path = Path(DB_FILENAME).resolve()
    allowed = ['cliente', 'medico', 'tecnico', 'tipo_cirugia', 'lugar', 'encargado_preparacion', 'encargado_logistica', 'coordinador_cx']
    if field_name not in allowed: return []
    try:
        conn = sqlite3.connect(db_path, timeout=5); cursor = conn.cursor()
        sql = f"SELECT DISTINCT {field_name} FROM cirugias WHERE {field_name} IS NOT NULL AND TRIM({field_name}) != '' ORDER BY fecha_generacion DESC LIMIT ?"
        cursor.execute(sql, (limit,)); suggestions = sorted([str(row[0]) for row in cursor.fetchall() if row[0]])
    except sqlite3.Error as e: print(f"Error sugerencias '{field_name}': {e}")
    finally:
        if conn: conn.close()
    return suggestions

def get_record_by_id(record_id):
    """Obtiene un registro completo por ID."""
    conn = None; record = None; db_path = Path(DB_FILENAME).resolve()
    if not record_id: return None
    try:
        conn = sqlite3.connect(db_path, timeout=5); conn.row_factory = sqlite3.Row; cursor = conn.cursor()
        cursor.execute("SELECT * FROM cirugias WHERE id = ?", (record_id,)); record = cursor.fetchone()
        return dict(record) if record else None
    except sqlite3.Error as e: print(f"Error get ID '{record_id}': {e}"); return None
    finally:
        if conn: conn.close()

def get_counts_by_preparador():
     """Obtiene conteos por preparador."""
     conn = None; counts = []; db_path = Path(DB_FILENAME).resolve()
     try:
         conn = sqlite3.connect(db_path, timeout=5); cursor = conn.cursor()
         sql = "SELECT encargado_preparacion, COUNT(*) as count FROM cirugias WHERE encargado_preparacion IS NOT NULL AND TRIM(encargado_preparacion) != '' GROUP BY encargado_preparacion ORDER BY count DESC, encargado_preparacion ASC"
         cursor.execute(sql); counts = cursor.fetchall()
     except sqlite3.Error as e: print(f"Error conteos prep: {e}")
     finally:
         if conn: conn.close()
     return counts

# --- Clase Ventana de Configuración ---
class SettingsWindow(ctk.CTkToplevel):
    def __init__(self, master=None, current_config=None):
        super().__init__(master)
        self.master_app = master
        self.initial_config = current_config if current_config else load_config()

        self.title("Configuración"); self.geometry("650x250"); self.resizable(False, False); self.transient(master); self.grab_set()

        self.img_dir_var = tk.StringVar(value=self.initial_config.get(CONFIG_KEY_IMG_DIR, ""))
        self.output_dir_var = tk.StringVar(value=self.initial_config.get(CONFIG_KEY_OUTPUT_DIR, ""))

        main_frame = ctk.CTkFrame(self); main_frame.pack(pady=15, padx=15, fill="both", expand=True); main_frame.columnconfigure(1, weight=1)
        ctk.CTkLabel(main_frame, text="Carpeta Imágenes (Default):").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        ctk.CTkEntry(main_frame, textvariable=self.img_dir_var, state="readonly").grid(row=0, column=1, padx=(0, 5), pady=10, sticky="ew")
        ctk.CTkButton(main_frame, text="Buscar...", width=100, command=self.browse_img_dir).grid(row=0, column=2, padx=(0, 10), pady=10)
        ctk.CTkLabel(main_frame, text="Carpeta Guardado PDF (Default):").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        ctk.CTkEntry(main_frame, textvariable=self.output_dir_var, state="readonly").grid(row=1, column=1, padx=(0, 5), pady=10, sticky="ew")
        ctk.CTkButton(main_frame, text="Buscar...", width=100, command=self.browse_output_dir).grid(row=1, column=2, padx=(0, 10), pady=10)
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent"); button_frame.grid(row=2, column=0, columnspan=3, pady=(20, 10))
        ctk.CTkButton(button_frame, text="Guardar Cambios", command=self.save_and_close).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="Cancelar", command=self.close_window, fg_color="grey").pack(side="left", padx=10)

    def browse_img_dir(self):
        initial = self.img_dir_var.get() or Path.home()
        directory = filedialog.askdirectory(title="Seleccionar Carpeta de Imágenes por Defecto", initialdir=initial)
        if directory: self.img_dir_var.set(directory)

    def browse_output_dir(self):
        initial = self.output_dir_var.get() or Path.home()
        directory = filedialog.askdirectory(title="Seleccionar Carpeta de Guardado por Defecto", initialdir=initial)
        if directory: self.output_dir_var.set(directory)

    def save_and_close(self):
        new_config = {CONFIG_KEY_IMG_DIR: self.img_dir_var.get(), CONFIG_KEY_OUTPUT_DIR: self.output_dir_var.get()}
        if save_config(new_config):
            if self.master_app: self.master_app.config = new_config; self.master_app._update_status("Configuración guardada.")
            self.close_window()
        else: show_error("Error Guardar", "No se pudo guardar el archivo de configuración.")

    def close_window(self): self.grab_release(); self.destroy()

# --- Clase Principal App ---
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Generador Reportes PDF v2.15 - 2 Columnas") # Versión actualizada
        self.geometry("900x900")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.config = load_config()
        init_db()

        self.image_file_paths = []; self.output_pdf_path_str = ""; self.last_generated_pdf_path = None
        self.stats_window = None; self.temp_image_files = []
        self.fecha_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d')); self.cliente_var = tk.StringVar(); self.paciente_var = tk.StringVar(); self.medico_var = tk.StringVar(); self.tecnico_var = tk.StringVar(); self.tipo_cirugia_var = tk.StringVar(); self.lugar_var = tk.StringVar(); self.enc_prep_var = tk.StringVar(); self.enc_log_var = tk.StringVar(); self.coord_cx_var = tk.StringVar()
        self.image_quality_var = tk.IntVar(value=IMG_QUALITY_HIGH)
        self.validation_error_message = ""
        self.settings_window = None

        self._create_widgets()
        self.update_suggestions()

    def _create_widgets(self):
        """Crea y organiza todos los widgets de la interfaz."""
        main_frame = ctk.CTkFrame(self); main_frame.pack(pady=15, padx=15, fill="both", expand=True); main_frame.columnconfigure(1, weight=1)
        row_idx = 0

        # --- Campos del Formulario ---
        ctk.CTkLabel(main_frame, text="Fecha Cirugía:").grid(row=row_idx, column=0, padx=10, pady=5, sticky="w")
        self.entry_fecha = DateEntry(main_frame, width=18, date_pattern='yyyy-mm-dd', textvariable=self.fecha_var, state='readonly', locale='es_ES'); self.entry_fecha.grid(row=row_idx, column=1, columnspan=2, padx=10, pady=5, sticky="w"); row_idx += 1
        ctk.CTkLabel(main_frame, text="Cliente:").grid(row=row_idx, column=0, padx=10, pady=5, sticky="w")
        self.combo_cliente = ctk.CTkComboBox(main_frame, variable=self.cliente_var); self.combo_cliente.grid(row=row_idx, column=1, columnspan=2, padx=10, pady=5, sticky="ew"); row_idx += 1
        ctk.CTkLabel(main_frame, text="Paciente:").grid(row=row_idx, column=0, padx=10, pady=5, sticky="w")
        self.entry_paciente = ctk.CTkEntry(main_frame, textvariable=self.paciente_var); self.entry_paciente.grid(row=row_idx, column=1, columnspan=2, padx=10, pady=5, sticky="ew"); row_idx += 1
        ctk.CTkLabel(main_frame, text="Médico:").grid(row=row_idx, column=0, padx=10, pady=5, sticky="w")
        self.combo_medico = ctk.CTkComboBox(main_frame, variable=self.medico_var); self.combo_medico.grid(row=row_idx, column=1, columnspan=2, padx=10, pady=5, sticky="ew"); row_idx += 1
        ctk.CTkLabel(main_frame, text="Técnico:").grid(row=row_idx, column=0, padx=10, pady=5, sticky="w")
        self.combo_tecnico = ctk.CTkComboBox(main_frame, variable=self.tecnico_var); self.combo_tecnico.grid(row=row_idx, column=1, columnspan=2, padx=10, pady=5, sticky="ew"); row_idx += 1
        ctk.CTkLabel(main_frame, text="Tipo Cirugía:").grid(row=row_idx, column=0, padx=10, pady=5, sticky="w")
        self.combo_tipo_cirugia = ctk.CTkComboBox(main_frame, variable=self.tipo_cirugia_var); self.combo_tipo_cirugia.grid(row=row_idx, column=1, columnspan=2, padx=10, pady=5, sticky="ew"); row_idx += 1
        ctk.CTkLabel(main_frame, text="Lugar:").grid(row=row_idx, column=0, padx=10, pady=5, sticky="w")
        self.combo_lugar = ctk.CTkComboBox(main_frame, variable=self.lugar_var); self.combo_lugar.grid(row=row_idx, column=1, columnspan=2, padx=10, pady=5, sticky="ew"); row_idx += 1
        ctk.CTkLabel(main_frame, text="Obs. Generales:").grid(row=row_idx, column=0, padx=10, pady=5, sticky="nw")
        self.text_obs_gen = ctk.CTkTextbox(main_frame, height=70); self.text_obs_gen.grid(row=row_idx, column=1, columnspan=2, padx=10, pady=5, sticky="ew"); row_idx += 1
        ctk.CTkFrame(main_frame, height=2, fg_color="gray50").grid(row=row_idx, column=0, columnspan=3, padx=10, pady=8, sticky="ew"); row_idx += 1
        ctk.CTkLabel(main_frame, text="Enc. Preparación:").grid(row=row_idx, column=0, padx=10, pady=5, sticky="w")
        self.combo_enc_prep = ctk.CTkComboBox(main_frame, variable=self.enc_prep_var); self.combo_enc_prep.grid(row=row_idx, column=1, columnspan=2, padx=10, pady=5, sticky="ew"); row_idx += 1
        ctk.CTkLabel(main_frame, text="Enc. Logística:").grid(row=row_idx, column=0, padx=10, pady=5, sticky="w")
        self.combo_enc_log = ctk.CTkComboBox(main_frame, variable=self.enc_log_var); self.combo_enc_log.grid(row=row_idx, column=1, columnspan=2, padx=10, pady=5, sticky="ew"); row_idx += 1
        ctk.CTkLabel(main_frame, text="Coord. CX:").grid(row=row_idx, column=0, padx=10, pady=5, sticky="w")
        self.combo_coord_cx = ctk.CTkComboBox(main_frame, variable=self.coord_cx_var); self.combo_coord_cx.grid(row=row_idx, column=1, columnspan=2, padx=10, pady=5, sticky="ew"); row_idx += 1
        ctk.CTkLabel(main_frame, text="Obs. Logística:").grid(row=row_idx, column=0, padx=10, pady=5, sticky="nw")
        self.text_obs_log = ctk.CTkTextbox(main_frame, height=70); self.text_obs_log.grid(row=row_idx, column=1, columnspan=2, padx=10, pady=5, sticky="ew"); row_idx += 1
        ctk.CTkFrame(main_frame, height=2, fg_color="gray50").grid(row=row_idx, column=0, columnspan=3, padx=10, pady=8, sticky="ew"); row_idx += 1

        # --- Opciones de Calidad Imagen ---
        options_frame = ctk.CTkFrame(main_frame, fg_color="transparent"); options_frame.grid(row=row_idx, column=0, columnspan=3, padx=5, pady=5, sticky="ew"); row_idx += 1
        ctk.CTkLabel(options_frame, text="Calidad Imagen:").grid(row=0, column=0, padx=(5,2), pady=5, sticky="w")
        q_sub = ctk.CTkFrame(options_frame, fg_color="transparent"); q_sub.grid(row=0, column=1, padx=0, pady=0, sticky="w")
        ctk.CTkRadioButton(q_sub, text="Alta(90)", variable=self.image_quality_var, value=IMG_QUALITY_HIGH).pack(side="left", padx=5)
        ctk.CTkRadioButton(q_sub, text="Media(75)", variable=self.image_quality_var, value=IMG_QUALITY_MEDIUM).pack(side="left", padx=5)
        ctk.CTkRadioButton(q_sub, text="Baja(60)", variable=self.image_quality_var, value=IMG_QUALITY_LOW).pack(side="left", padx=5)
        options_frame.columnconfigure(2, weight=1)

        # --- Selección de Imágenes y Salida ---
        ctk.CTkLabel(main_frame, text="Imágenes:").grid(row=row_idx, column=0, padx=10, pady=5, sticky="w")
        self.entry_images = ctk.CTkEntry(main_frame, placeholder_text="Ninguna seleccionada", state="readonly"); self.entry_images.grid(row=row_idx, column=1, padx=(10, 5), pady=5, sticky="ew")
        self.button_browse = ctk.CTkButton(main_frame, text="Buscar Imágenes...", width=140, command=self.browse_images); self.button_browse.grid(row=row_idx, column=2, padx=(0, 10), pady=5, sticky="e"); row_idx += 1
        ctk.CTkLabel(main_frame, text="Guardar PDF en:").grid(row=row_idx, column=0, padx=10, pady=5, sticky="w")
        self.entry_output = ctk.CTkEntry(main_frame, placeholder_text="Seleccionar ubicación...", state="readonly"); self.entry_output.grid(row=row_idx, column=1, padx=(10, 5), pady=5, sticky="ew")
        self.button_save_as = ctk.CTkButton(main_frame, text="Guardar Como...", width=140, command=self.select_save_path); self.button_save_as.grid(row=row_idx, column=2, padx=(0, 10), pady=5, sticky="e"); row_idx += 1

        # --- Botones de Acción (Organizados) ---
        action_frame = ctk.CTkFrame(main_frame, fg_color="transparent"); action_frame.grid(row=row_idx, column=0, columnspan=3, padx=10, pady=(15, 5), sticky="ew"); row_idx += 1
        main_buttons_frame = ctk.CTkFrame(action_frame, fg_color="transparent"); main_buttons_frame.pack(side="left", fill="x", expand=True)
        self.button_generate = ctk.CTkButton(main_buttons_frame, text="Generar PDF", command=self.start_pdf_generation_thread); self.button_generate.pack(side="left", padx=5, pady=5)
        self.button_clear = ctk.CTkButton(main_buttons_frame, text="Limpiar", command=self.clear_form, fg_color="grey"); self.button_clear.pack(side="left", padx=5, pady=5)
        self.button_settings = ctk.CTkButton(main_buttons_frame, text="Configuración", command=self.open_settings_window, fg_color="#008080"); self.button_settings.pack(side="left", padx=5, pady=5) # Teal color
        post_buttons_frame = ctk.CTkFrame(action_frame, fg_color="transparent"); post_buttons_frame.pack(side="right")
        self.button_print = ctk.CTkButton(post_buttons_frame, text="Imprimir", command=self.print_last_pdf, state="disabled", width=100); self.button_print.pack(side="left", padx=5, pady=5)
        self.button_email = ctk.CTkButton(post_buttons_frame, text="Email", command=self.share_email, state="disabled", width=100); self.button_email.pack(side="left", padx=5, pady=5)
        self.button_stats = ctk.CTkButton(post_buttons_frame, text="Registros", command=self.show_stats_window, width=100); self.button_stats.pack(side="left", padx=5, pady=5)

        # --- Etiqueta de Estado ---
        self.status_label = ctk.CTkLabel(main_frame, text="Listo.", anchor="w", text_color="gray60"); self.status_label.grid(row=row_idx, column=0, columnspan=3, padx=10, pady=(5, 10), sticky="ew"); row_idx += 1

    def open_settings_window(self):
        """Abre la ventana de configuración."""
        if self.settings_window is None or not self.settings_window.winfo_exists():
            self.settings_window = SettingsWindow(self, current_config=self.config); self.settings_window.focus()
        else: self.settings_window.focus()

    def _update_status(self, message, is_error=False):
        """Actualiza etiqueta de estado (seguro para hilos via self.after)."""
        color = "tomato" if is_error else "gray60"
        self.after(0, lambda msg=message, clr=color: self.status_label.configure(text=msg, text_color=clr)); print(f"STATUS: {message}")

    def clear_form(self):
        """Limpia el formulario y resetea variables."""
        if not messagebox.askyesno("Confirmar", "¿Limpiar formulario?"): return
        self._update_status("Limpiando formulario...")
        self.fecha_var.set(datetime.now().strftime('%Y-%m-%d')); self.cliente_var.set(""); self.paciente_var.set(""); self.medico_var.set(""); self.tecnico_var.set(""); self.tipo_cirugia_var.set(""); self.lugar_var.set(""); self.enc_prep_var.set(""); self.enc_log_var.set(""); self.coord_cx_var.set("")
        self.text_obs_gen.delete("1.0", "end"); self.text_obs_log.delete("1.0", "end")
        self.image_file_paths = []; self._cleanup_temp_images()
        self.entry_images.configure(state="normal"); self.entry_images.delete(0, "end"); self.entry_images.insert(0, "Ninguna"); self.entry_images.configure(state="readonly")
        self.output_pdf_path_str = ""; self.entry_output.configure(state="normal"); self.entry_output.delete(0, "end"); self.entry_output.insert(0, "Seleccionar..."); self.entry_output.configure(state="readonly")
        self.last_generated_pdf_path = None; self.button_print.configure(state="disabled"); self.button_email.configure(state="disabled")
        self.image_quality_var.set(IMG_QUALITY_HIGH)
        for combo in [self.combo_cliente, self.combo_medico, self.combo_tecnico, self.combo_tipo_cirugia, self.combo_lugar, self.combo_enc_prep, self.combo_enc_log, self.combo_coord_cx]: combo.set("")
        self._update_status("Formulario limpiado.")

    def browse_images(self):
        """Abre diálogo para seleccionar imágenes, usando dir default."""
        self._cleanup_temp_images(); self.image_file_paths = []
        initial_img_dir = self.config.get(CONFIG_KEY_IMG_DIR) or str(Path.home())
        if not Path(initial_img_dir).is_dir(): print(f"WARN: Dir imágenes default no válido: '{initial_img_dir}'. Usando Home."); initial_img_dir = str(Path.home())
        files = filedialog.askopenfilenames(title=f"Seleccionar Imágenes (Máx {MAX_IMAGES_ALLOWED})", initialdir=initial_img_dir, filetypes=(("Imágenes", "*.png *.jpg *.jpeg *.bmp *.gif"), ("Todos", "*.*")))
        if files:
            self.image_file_paths = list(files)[:MAX_IMAGES_ALLOWED]; count = len(self.image_file_paths)
            disp = f"{count} imagen{'s'[:count^1]} seleccionada{'s'[:count^1]}"
            if len(files) > MAX_IMAGES_ALLOWED: disp += f" (Mostrando {MAX_IMAGES_ALLOWED})"
            self.entry_images.configure(state="normal"); self.entry_images.delete(0, "end"); self.entry_images.insert(0, disp); self.entry_images.configure(state="readonly")
            self._update_status(f"{count} imágenes seleccionadas.")
        else:
            self.image_file_paths = []; self.entry_images.configure(state="normal"); self.entry_images.delete(0, "end"); self.entry_images.insert(0, "Ninguna"); self.entry_images.configure(state="readonly")
            self._update_status("Ninguna imagen seleccionada.")

    def sanitize_filename(self, name):
        """Limpia un nombre de archivo."""
        if not name: return "_"
        valid = "-_.() abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"; san = ''.join(c for c in name if c in valid)
        san = san.replace(' ', '_'); san = san.strip('_.- '); return san[:100] or "_"

    def select_save_path(self, suggested_filename=None):
        """Abre diálogo para guardar PDF, usando dir default. Devuelve True/False."""
        if not suggested_filename:
            p = self.sanitize_filename(self.paciente_var.get()) or "P"; c = self.sanitize_filename(self.cliente_var.get()) or "C"
            try: f = datetime.strptime(self.fecha_var.get(), '%Y-%m-%d').strftime('%Y%m%d')
            except ValueError: f = datetime.now().strftime('%Y%m%d')
            suggested_filename = f"Reporte_{p}_{c}_{f}.pdf"
        initial_out_dir = self.config.get(CONFIG_KEY_OUTPUT_DIR) or str(Path.home())
        if not Path(initial_out_dir).is_dir(): print(f"WARN: Dir salida default no válido: '{initial_out_dir}'. Usando Home."); initial_out_dir = str(Path.home())
        file_path = filedialog.asksaveasfilename(title="Guardar PDF Como", initialdir=initial_out_dir, initialfile=suggested_filename, defaultextension=".pdf", filetypes=(("PDF", "*.pdf"), ("Todos", "*.*")))
        if file_path:
            self.output_pdf_path_str = file_path; self.entry_output.configure(state="normal"); self.entry_output.delete(0, "end"); self.entry_output.insert(0, file_path); self.entry_output.configure(state="readonly")
            self._update_status(f"Guardar en: {Path(file_path).name}"); return True
        else:
            current_path = self.output_pdf_path_str or "Seleccionar..."
            self.entry_output.configure(state="normal"); self.entry_output.delete(0, "end"); self.entry_output.insert(0, current_path); self.entry_output.configure(state="readonly")
            self._update_status("Guardado cancelado."); return False

    def _validate_inputs(self):
        """Valida campos obligatorios antes de generar."""
        errors = []
        if not self.fecha_var.get():
            errors.append("- Fecha cirugía obligatoria.")
        else:
            # --- BLOQUE CORREGIDO ---
            try:
                datetime.strptime(self.fecha_var.get(), '%Y-%m-%d')
            except ValueError:
                errors.append("- Formato fecha inválido (AAAA-MM-DD).")
            # --- FIN BLOQUE CORREGIDO ---

        if not self.paciente_var.get().strip(): errors.append("- Paciente obligatorio.")
        if not self.cliente_var.get().strip(): errors.append("- Cliente obligatorio.")
        if not self.medico_var.get().strip(): errors.append("- Médico obligatorio.")
        if not self.enc_prep_var.get().strip(): errors.append("- Enc. Preparación obligatorio.")
        if not self.output_pdf_path_str: errors.append("- Selecciona dónde guardar PDF.")

        if errors: self.validation_error_message = "\n".join(errors); return False
        self.validation_error_message = ""; return True

    def _cleanup_temp_images(self):
        """Elimina las imágenes temporales comprimidas."""
        print("Limpiando imágenes temporales...")
        deleted_count = 0
        for temp_file in list(self.temp_image_files): # Iterar sobre copia
            if temp_file and temp_file.exists():
                try: temp_file.unlink(); deleted_count += 1; self.temp_image_files.remove(temp_file)
                except Exception as e: print(f"WARN: No eliminar {temp_file.name}: {e}")
            elif temp_file in self.temp_image_files: self.temp_image_files.remove(temp_file)
        print(f"Limpieza: {deleted_count} archivos eliminados.")
        self.temp_image_files = []

    # --- Gestión de Hilos ---
    def start_pdf_generation_thread(self):
        """Inicia la generación en un hilo."""
        if not self._validate_inputs():
            self._update_status("Error: Corrige los datos.", is_error=True)
            self.after(0, lambda: show_error_safe("Datos Inválidos", self.validation_error_message))
            return
        self.button_generate.configure(state="disabled", text="Generando..."); self.button_clear.configure(state="disabled"); self.button_stats.configure(state="disabled"); self.button_browse.configure(state="disabled"); self.button_save_as.configure(state="disabled"); self.button_settings.configure(state="disabled")
        self._update_status("Iniciando generación..."); self.last_generated_pdf_path = None
        self.button_print.configure(state="disabled"); self.button_email.configure(state="disabled")
        global last_conversion_error, last_db_error; last_conversion_error = ""; last_db_error = ""
        quality = self.image_quality_var.get()
        # Layout ahora es fijo a 2, se usa image_pairs
        thread = threading.Thread(target=self.generate_pdf_worker, args=(quality,), daemon=True)
        thread.start()

    def generate_pdf_worker(self, image_quality):
        """Lógica de generación (ejecutada en hilo). Layout fijo a 2."""
        images_per_row = 2 # Fijo a 2 columnas
        print(f"\n--- HILO GENERATE INICIADO (Q:{image_quality}, L:{images_per_row}/fila) ---"); start_time = time.time()
        success = False; render_success = False; conversion_success = False; db_save_success = False
        generated_pdf_final_path = None; temp_docx_path = None
        try:
            self._update_status("Paso 1/5: Recopilando datos..."); time.sleep(0.1)
            context = { 'fecha_cirugia': self.fecha_var.get(), 'cliente': self.cliente_var.get().strip(), 'paciente': self.paciente_var.get().strip(), 'medico': self.medico_var.get().strip(), 'tecnico': self.tecnico_var.get().strip(), 'tipo_cirugia': self.tipo_cirugia_var.get().strip(), 'lugar': self.lugar_var.get().strip(), 'observaciones_generales': self.text_obs_gen.get("1.0", "end-1c").strip(), 'encargado_preparacion': self.enc_prep_var.get().strip(), 'encargado_logistica': self.enc_log_var.get().strip(), 'coordinador_cx': self.coord_cx_var.get().strip(), 'observaciones_logistica': self.text_obs_log.get("1.0", "end-1c").strip(),
                       'image_pairs': [] } # Usaremos image_pairs
            record_id = str(uuid.uuid4()); unique_id_for_qr = str(uuid.uuid4()); context['unique_id'] = unique_id_for_qr
            now = datetime.now(); context['fecha_emision'] = now.strftime('%d/%m/%Y %H:%M:%S'); context['fecha_emision_corta'] = now.strftime('%d/%m/%Y'); print(f"DEBUG: Fecha emisión: {context['fecha_emision']}")
            self._update_status("Paso 1/5: Preparando archivos..."); time.sleep(0.1)
            template_path = Path(TEMPLATE_FILENAME); output_pdf_path_user = Path(self.output_pdf_path_str); temp_docx_path = Path(output_pdf_path_user.parent / f"temp_docx_{record_id[:8]}.docx").resolve()
            if not template_path.exists(): raise FileNotFoundError(f"Plantilla '{TEMPLATE_FILENAME}' no encontrada.")
            doc = DocxTemplate(template_path)
            self._update_status("Paso 2/5: Procesando imágenes..."); time.sleep(0.1)
            processed_images = []; num_images = len(self.image_file_paths); target_width_mm = LAYOUT_2_PER_ROW_WIDTH_MM # Usar ancho para 2 columnas
            print(f"DEBUG: Procesando {num_images} imágenes (target width: {target_width_mm}mm).")
            for i, img_path_str in enumerate(self.image_file_paths):
                 self._update_status(f"Paso 2/5: Procesando imagen {i+1}/{num_images}..."); time.sleep(0.05)
                 img_path = Path(img_path_str.strip())
                 if not img_path.exists(): print(f"WARN: Imagen no encontrada: {img_path}"); continue
                 compressed_path = compress_image(img_path, target_width_mm, image_quality); print(f"DEBUG: compress_image para '{img_path.name}' devolvió: {compressed_path}")
                 if compressed_path:
                     self.temp_image_files.append(compressed_path)
                     try: inline_img = InlineImage(doc, str(compressed_path), width=Mm(target_width_mm)); processed_images.append(inline_img); print(f"DEBUG: InlineImage creado para '{compressed_path.name}'")
                     except Exception as img_add_error: print(f"ERROR Add InlineImage: {compressed_path.name}: {img_add_error}"); print(f"DEBUG: Falló InlineImage para '{compressed_path.name}'")
                 else: print(f"WARN: Falló compresión: {img_path.name}"); print(f"DEBUG: Saltando imagen '{img_path.name}'")
            print(f"DEBUG: Total InlineImage procesados: {len(processed_images)}")
            # Agrupar en pares para la plantilla
            image_pairs = [];
            for i in range(0, len(processed_images), 2): img1 = processed_images[i]; img2 = processed_images[i+1] if (i+1) < len(processed_images) else None; image_pairs.append((img1, img2))
            context['image_pairs'] = image_pairs; print(f"DEBUG: Añadiendo 'image_pairs' con {len(context['image_pairs'])} pares.")
            self._update_status("Paso 3/5: Generando DOCX..."); time.sleep(0.1); print(f"DEBUG: Contexto Keys={list(context.keys())}"); print(f"Renderizando DOCX: {temp_docx_path}..."); doc.render(context); doc.save(temp_docx_path); print("DOCX OK."); render_success = True
            self._update_status("Paso 4/5: Convirtiendo a PDF..."); time.sleep(0.1)
            if convert_to_pdf(temp_docx_path, output_pdf_path_user.parent):
                generated_pdf_temp_name = temp_docx_path.with_suffix('.pdf')
                if generated_pdf_temp_name.exists():
                    try:
                        if output_pdf_path_user.exists(): output_pdf_path_user.unlink()
                        generated_pdf_temp_name.rename(output_pdf_path_user); generated_pdf_final_path = str(output_pdf_path_user); conversion_success = True
                    except Exception as rename_error: global last_conversion_error; last_conversion_error = f"PDF generado ({generated_pdf_temp_name.name}) pero no renombrado: {rename_error}"; generated_pdf_final_path = str(generated_pdf_temp_name)
                else: last_conversion_error = f"Conversión OK, pero PDF no encontrado: {generated_pdf_temp_name}"
            if render_success and conversion_success:
                self._update_status("Paso 5/5: Guardando registro..."); time.sleep(0.1)
                record_data = {k: v for k, v in context.items() if k not in ['image_pairs', 'fecha_emision', 'fecha_emision_corta']} # Excluir image_pairs y fechas
                record_data.update({'id': record_id, 'fecha_generacion': datetime.now().isoformat(timespec='seconds'), 'archivo_pdf': generated_pdf_final_path, 'unique_id': unique_id_for_qr})
                if save_record(record_data): db_save_success = True; self.after(0, self.update_suggestions)
            success = render_success and conversion_success and db_save_success
        except FileNotFoundError as e: print(f"ERROR Worker (FileNotFound): {e}"); self.after(0, lambda e=e: show_error_safe("Error Crítico", f"Archivo no encontrado:\n{e}")); self._update_status(f"Error: {e}", is_error=True)
        except Exception as e: print(f"ERROR Worker (General): {type(e).__name__} - {e}\n{traceback.format_exc()}"); self.after(0, lambda e=e: show_error_safe("Error Crítico Inesperado", f"Ocurrió un error:\n{type(e).__name__}: {e}")); self._update_status(f"Error crítico: {e}", is_error=True)
        finally:
            end_time = time.time(); duration = end_time - start_time; print(f"--- HILO GENERATE FINALIZADO (Dur: {duration:.2f}s, Éxito Render: {render_success}, Éxito Convert: {conversion_success}, Éxito DB: {db_save_success}) ---")
            self.after(0, self._finalize_generation, success, generated_pdf_final_path, duration)
            print("Limpiando archivos temporales (worker)...")
            if temp_docx_path and temp_docx_path.exists():
                try: temp_docx_path.unlink(); print(f"  - DOCX temp eliminado.")
                except Exception as clean_err: print(f"WARN: No eliminar DOCX temp: {clean_err}")
            self._cleanup_temp_images()

    def _finalize_generation(self, success, final_pdf_path, duration):
        """Actualiza GUI al finalizar generación."""
        print("Finalizando en GUI...");
        self.button_generate.configure(state="normal", text="Generar PDF"); self.button_clear.configure(state="normal"); self.button_stats.configure(state="normal"); self.button_browse.configure(state="normal"); self.button_save_as.configure(state="normal"); self.button_settings.configure(state="normal")
        fail_reason = ""
        if not success:
            if last_conversion_error: fail_reason = f"Conversión PDF: {last_conversion_error}"
            elif last_db_error: fail_reason = f"Base de Datos: {last_db_error}"
            else: fail_reason = "Causa desconocida (ver consola)."
        if success:
            final_msg = f"Éxito! PDF en {duration:.1f}s: {Path(final_pdf_path).name}"
            self._update_status(final_msg); self.after(0, lambda p=final_pdf_path: show_info_safe("Generación Completada", f"Éxito! PDF generado:\n{Path(p).name}"))
            self.last_generated_pdf_path = final_pdf_path; self.button_print.configure(state="normal"); self.button_email.configure(state="normal")
        else:
            final_msg = f"Generación Fallida ({duration:.1f}s)."
            error_display = f"La generación del PDF falló.\n\nMotivo: {fail_reason}"
            self.after(0, lambda ed=error_display: show_error_safe("Error de Generación", ed))
            self._update_status(f"Generación fallida ({duration:.1f}s).", is_error=True); self.last_generated_pdf_path = None; self.button_print.configure(state="disabled"); self.button_email.configure(state="disabled")

    # --- Hilos para Regeneración ---
    def start_regenerate_thread(self, data, save_path):
        """Inicia la regeneración en un hilo."""
        self._update_status("Iniciando regeneración...")
        global last_conversion_error; last_conversion_error = ""
        thread = threading.Thread(target=self.regenerate_pdf_worker, args=(data, save_path), daemon=True)
        thread.start()

    def regenerate_pdf_worker(self, data, output_pdf_path_str):
        """Lógica de regeneración (ejecutada en hilo)."""
        print(f"\n--- HILO REGENERATE INICIADO (ID: {data.get('id', 'N/A')}) ---"); start_time = time.time()
        success = False; render_success = False; conversion_success = False
        generated_pdf_final_path = None; temp_docx_path = None
        original_unique_id = data.get('unique_id'); original_record_id = data.get('id')
        try:
            if not original_unique_id or not original_record_id: raise ValueError("Datos de registro inválidos.")
            self._update_status("Regenerando: Preparando datos..."); time.sleep(0.1)
            output_pdf_path = Path(output_pdf_path_str); template_path = Path(TEMPLATE_FILENAME); temp_docx_path = Path(output_pdf_path.parent / f"temp_regen_{original_record_id[:8]}.docx").resolve()
            keys_to_keep = ['fecha_cirugia', 'cliente', 'paciente', 'medico', 'tecnico', 'tipo_cirugia', 'lugar', 'observaciones_generales', 'encargado_preparacion', 'encargado_logistica', 'coordinador_cx', 'observaciones_logistica']
            context = {k: data.get(k, '') for k in keys_to_keep}; context['unique_id'] = original_unique_id
            now = datetime.now(); context['fecha_emision'] = now.strftime('%d/%m/%Y %H:%M:%S'); context['fecha_emision_corta'] = now.strftime('%d/%m/%Y')
            regen_warning = "\n\n--- REGENERADO (SIN IMÁGENES ORIGINALES) ---"; context['observaciones_generales'] += regen_warning
            context['image_pairs'] = [] # Sin imágenes, usar la clave correcta para la plantilla
            if not template_path.exists(): raise FileNotFoundError(f"Plantilla '{TEMPLATE_FILENAME}' no encontrada.")
            doc = DocxTemplate(template_path)
            self._update_status("Regenerando: Creando DOCX..."); time.sleep(0.1); print(f"Renderizando DOCX (regen): {temp_docx_path}..."); doc.render(context); doc.save(temp_docx_path); print("DOCX (regen) OK."); render_success = True
            self._update_status("Regenerando: Convirtiendo a PDF..."); time.sleep(0.1)
            if convert_to_pdf(temp_docx_path, output_pdf_path.parent):
                generated_pdf_temp_name = temp_docx_path.with_suffix('.pdf')
                if generated_pdf_temp_name.exists():
                    try:
                        if output_pdf_path.exists(): output_pdf_path.unlink()
                        generated_pdf_temp_name.rename(output_pdf_path); generated_pdf_final_path = str(output_pdf_path); conversion_success = True
                    except Exception as rename_error: global last_conversion_error; last_conversion_error = f"PDF (regen) generado ({generated_pdf_temp_name.name}) pero no renombrado: {rename_error}"; generated_pdf_final_path = str(generated_pdf_temp_name)
                else: last_conversion_error = f"Conversión (regen) OK, pero PDF no encontrado: {generated_pdf_temp_name}"
            success = render_success and conversion_success
        except FileNotFoundError as e: print(f"ERROR Regen Worker (FileNotFound): {e}"); self.after(0, lambda e=e: show_error_safe("Error Crítico (Regen)", f"Archivo no encontrado:\n{e}")); self._update_status(f"Error Regen: {e}", is_error=True)
        except Exception as e: print(f"ERROR Regen Worker (General): {type(e).__name__} - {e}\n{traceback.format_exc()}"); self.after(0, lambda e=e: show_error_safe("Error Crítico (Regen)", f"Ocurrió un error:\n{type(e).__name__}: {e}")); self._update_status(f"Error crítico Regen: {e}", is_error=True)
        finally:
            end_time = time.time(); duration = end_time - start_time; print(f"--- HILO REGENERATE FINALIZADO (Dur: {duration:.2f}s, Éxito: {success}) ---")
            self.after(0, self._finalize_regeneration, success, generated_pdf_final_path, duration)
            print("Limpiando DOCX temporal (regen)...")
            if temp_docx_path and temp_docx_path.exists():
                try: temp_docx_path.unlink(); print(f"  - DOCX temp (regen) eliminado.")
                except Exception as clean_err: print(f"WARN: No eliminar DOCX temp (regen): {clean_err}")

    def _finalize_regeneration(self, success, final_pdf_path, duration):
        """Actualiza GUI al finalizar regeneración."""
        print("Finalizando regeneración en GUI...")
        fail_reason = ""
        if not success:
             if last_conversion_error: fail_reason = f"Conversión PDF: {last_conversion_error}"
             else: fail_reason = "Causa desconocida (ver consola)."
        if success:
            final_msg = f"Regeneración OK ({duration:.1f}s): {Path(final_pdf_path).name}"
            self._update_status(final_msg); self.after(0, lambda p=final_pdf_path: show_info_safe("Regeneración Completada", f"PDF Regenerado:\n{Path(p).name}\n(Sin imágenes originales)"))
            if self.winfo_exists(): self.last_generated_pdf_path = final_pdf_path; self.button_print.configure(state="normal"); self.button_email.configure(state="normal")
        else:
            final_msg = f"Regeneración Fallida ({duration:.1f}s)."
            error_display = f"Falló la regeneración.\n\nMotivo: {fail_reason}"
            self.after(0, lambda ed=error_display: show_error_safe("Error de Regeneración", ed))
            self._update_status(f"Regeneración fallida ({duration:.1f}s).", is_error=True);
            if self.winfo_exists(): self.last_generated_pdf_path = None; self.button_print.configure(state="disabled"); self.button_email.configure(state="disabled")

    # --- Otros métodos ---
    def update_suggestions(self):
        """Actualiza las listas de sugerencias en los Combobox."""
        print("Actualizando sugerencias...")
        suggestion_map = { self.combo_medico: 'medico', self.combo_cliente: 'cliente', self.combo_tecnico: 'tecnico', self.combo_tipo_cirugia: 'tipo_cirugia', self.combo_lugar: 'lugar', self.combo_enc_prep: 'encargado_preparacion', self.combo_enc_log: 'encargado_logistica', self.combo_coord_cx: 'coordinador_cx' }
        for combo, field in suggestion_map.items():
            current = combo.get(); suggestions = get_suggestions(field); combo.configure(values=suggestions)
            if current in suggestions or (current and current not in combo.cget('values')): combo.set(current)
            else: combo.set("")
        print("Sugerencias OK.")

    def print_last_pdf(self):
        """Abre el último PDF generado para imprimir."""
        if not self.last_generated_pdf_path: show_error("Error", "No hay PDF generado."); return
        pdf_path = Path(self.last_generated_pdf_path)
        if not pdf_path.exists(): show_error("Error", f"Archivo no encontrado:\n{pdf_path}"); self.last_generated_pdf_path=None; self.button_print.configure(state="disabled"); return
        try:
            if sys.platform == "win32": os.startfile(str(pdf_path))
            elif sys.platform == "darwin": subprocess.run(["open", str(pdf_path)], check=True)
            else: subprocess.run(["xdg-open", str(pdf_path)], check=True)
            show_info("Imprimir", f"Abriendo '{pdf_path.name}'.\nUsa la opción imprimir del visor.")
        except Exception as e: show_error("Error", f"No abrir PDF:\n{e}")

    def share_email(self):
        """Abre cliente de correo para compartir último PDF."""
        if not self.last_generated_pdf_path: show_error("Error", "Genera PDF primero."); return
        pdf_path = Path(self.last_generated_pdf_path)
        if not pdf_path.exists(): show_error("Error", f"Archivo no encontrado:\n{pdf_path}"); self.last_generated_pdf_path=None; self.button_email.configure(state="disabled"); return
        try:
            p = self.paciente_var.get() or "N/A"; f = self.fecha_var.get() or "N/A"; subj = f"Reporte Cirugía - {p} ({f})"
            body = f"Adjunto reporte:\nPaciente: {p}\nFecha: {f}\nArchivo: {pdf_path.name}\n\n**ADJUNTA MANUALMENTE:**\n{pdf_path.parent}\n\nSaludos."
            url = f"mailto:?subject={quote(subj)}&body={quote(body)}"; webbrowser.open(url)
            show_info("Email", f"Abriendo cliente correo...\n\n¡¡ ADJUNTA MANUALMENTE !!\nArchivo: {pdf_path.name}\nEn: {pdf_path.parent}")
        except Exception as e: show_error("Error", f"No abrir cliente correo:\n{e}")

    def show_stats_window(self):
        """Muestra la ventana de estadísticas."""
        if self.stats_window and self.stats_window.winfo_exists(): self.stats_window.lift(); self.stats_window.focus()
        else: self.stats_window = StatsWindow(self); self.stats_window.focus()

    def load_data_into_form(self, data):
        """Carga datos de un registro en el formulario."""
        if not data or not isinstance(data, dict): show_error("Error", "Datos inválidos."); return
        if self.paciente_var.get() or self.image_file_paths:
             if not messagebox.askyesno("Confirmar", "Sobrescribir formulario actual?\n(Imágenes NO se cargarán)"): return
        self.clear_form() # Limpiar antes de cargar
        self.fecha_var.set(data.get('fecha_cirugia', '')); self.cliente_var.set(data.get('cliente', '')); self.paciente_var.set(data.get('paciente', '')); self.medico_var.set(data.get('medico','')); self.tecnico_var.set(data.get('tecnico','')); self.tipo_cirugia_var.set(data.get('tipo_cirugia','')); self.lugar_var.set(data.get('lugar','')); self.enc_prep_var.set(data.get('encargado_preparacion','')); self.enc_log_var.set(data.get('encargado_logistica','')); self.coord_cx_var.set(data.get('coordinador_cx',''))
        self.text_obs_gen.insert("1.0", data.get('observaciones_generales', '')); self.text_obs_log.insert("1.0", data.get('observaciones_logistica', ''))
        self.entry_images.configure(state="normal"); self.entry_images.delete(0, "end"); self.entry_images.insert(0, "Ninguna (Datos importados)"); self.entry_images.configure(state="readonly")
        self._update_status(f"Datos importados para {data.get('paciente', 'N/A')}.")
        show_info("Datos Importados", f"Datos cargados para: {data.get('paciente', 'N/A')}")

    def on_closing(self):
        """Acciones al cerrar la ventana principal."""
        print("Cerrando aplicación..."); self._cleanup_temp_images()
        if self.stats_window and self.stats_window.winfo_exists(): self.stats_window.destroy()
        if self.settings_window and self.settings_window.winfo_exists(): self.settings_window.destroy()
        self.destroy()

# --- Clase StatsWindow ---
class StatsWindow(ctk.CTkToplevel):
    """Ventana para ver, filtrar y gestionar registros."""
    def __init__(self, master=None):
        super().__init__(master); self.master_app = master; self.title("Registros Cirugías"); self.geometry("1150x750"); self.resizable(True, True)
        self.transient(master); self.grab_set(); self.protocol("WM_DELETE_WINDOW", self.on_closing)
        container = ctk.CTkFrame(self); container.pack(pady=10, padx=10, fill="both", expand=True)
        self._create_filter_widgets(container); self._create_results_widgets(container); self._create_bottom_widgets(container)
        self.load_stats()

    def _create_filter_widgets(self, parent):
        """Crea los widgets de filtro."""
        filter_frame = ctk.CTkFrame(parent); filter_frame.pack(pady=(0, 10), padx=10, fill="x")
        filter_frame.columnconfigure(1, weight=1); filter_frame.columnconfigure(3, weight=1)
        ctk.CTkLabel(filter_frame, text="Desde:").grid(row=0, column=0, padx=(10,5), pady=5, sticky="w"); self.date_from = DateEntry(filter_frame, width=12, date_pattern='yyyy-mm-dd', state='readonly', locale='es_ES'); self.date_from.grid(row=0, column=1, padx=5, pady=5, sticky="w"); self.date_from.set_date(None)
        ctk.CTkLabel(filter_frame, text="Hasta:").grid(row=0, column=2, padx=(10,5), pady=5, sticky="w"); self.date_to = DateEntry(filter_frame, width=12, date_pattern='yyyy-mm-dd', state='readonly', locale='es_ES'); self.date_to.grid(row=0, column=3, padx=5, pady=5, sticky="w"); self.date_to.set_date(None)
        ctk.CTkButton(filter_frame, text="Limpiar Fechas", width=100, command=self._clear_dates, fg_color="grey").grid(row=0, column=4, padx=(10, 5), pady=5)
        ctk.CTkLabel(filter_frame, text="Médico:").grid(row=1, column=0, padx=(10,5), pady=5, sticky="w"); self.medico_filter = ctk.CTkEntry(filter_frame, placeholder_text="Buscar...", width=200); self.medico_filter.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkLabel(filter_frame, text="Paciente:").grid(row=1, column=2, padx=(10,5), pady=5, sticky="w"); self.paciente_filter = ctk.CTkEntry(filter_frame, placeholder_text="Buscar...", width=200); self.paciente_filter.grid(row=1, column=3, padx=5, pady=5, sticky="ew")
        ctk.CTkLabel(filter_frame, text="Cliente:").grid(row=2, column=0, padx=(10,5), pady=5, sticky="w"); self.cliente_filter = ctk.CTkEntry(filter_frame, placeholder_text="Buscar...", width=200); self.cliente_filter.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkLabel(filter_frame, text="ID QR:").grid(row=2, column=2, padx=(10,5), pady=5, sticky="w"); self.unique_id_filter = ctk.CTkEntry(filter_frame, placeholder_text="ID exacto...", width=200); self.unique_id_filter.grid(row=2, column=3, padx=5, pady=5, sticky="ew")
        b_sub = ctk.CTkFrame(filter_frame, fg_color="transparent"); b_sub.grid(row=3, column=0, columnspan=5, padx=10, pady=10)
        ctk.CTkButton(b_sub, text="Buscar", command=self.load_stats).pack(side="left", padx=5); ctk.CTkButton(b_sub, text="Limpiar Filtros", command=self._clear_filters, fg_color="grey").pack(side="left", padx=5)

    def _clear_dates(self): self.date_from.set_date(None); self.date_to.set_date(None)
    def _clear_filters(self): self._clear_dates(); self.medico_filter.delete(0, "end"); self.paciente_filter.delete(0, "end"); self.cliente_filter.delete(0, "end"); self.unique_id_filter.delete(0, "end"); self.load_stats()

    def _create_results_widgets(self, parent):
        """Crea la tabla Treeview para mostrar resultados."""
        results_frame = ctk.CTkFrame(parent); results_frame.pack(pady=10, padx=10, fill="both", expand=True); style = ttk.Style(self); style.theme_use("clam")
        cols = ('fecha_gen', 'fecha_cx', 'paciente', 'medico', 'tipo', 'lugar', 'enc_prep', 'unique_id_col'); self.tree = ttk.Treeview(results_frame, columns=cols, show='headings')
        h = self.tree.heading; s = self._sort_column; c = self.tree.column # Alias
        h('fecha_gen', text='Generado', command=lambda: s('fecha_gen',False)); c('fecha_gen', width=130, anchor='w')
        h('fecha_cx', text='Fecha Cx', command=lambda: s('fecha_cx',False)); c('fecha_cx', width=90, anchor='c')
        h('paciente', text='Paciente', command=lambda: s('paciente',False)); c('paciente', width=180, anchor='w')
        h('medico', text='Médico', command=lambda: s('medico',False)); c('medico', width=180, anchor='w')
        h('tipo', text='Tipo Cx', command=lambda: s('tipo',False)); c('tipo', width=150, anchor='w')
        h('lugar', text='Lugar', command=lambda: s('lugar',False)); c('lugar', width=150, anchor='w')
        h('enc_prep', text='Enc Prep', command=lambda: s('enc_prep',False)); c('enc_prep', width=130, anchor='w')
        h('unique_id_col', text='ID QR', command=lambda: s('unique_id_col',False)); c('unique_id_col', width=100, anchor='w')
        vsb = ttk.Scrollbar(results_frame, orient="vertical", command=self.tree.yview); hsb = ttk.Scrollbar(results_frame, orient="horizontal", command=self.tree.xview); self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set); vsb.pack(side='right', fill='y'); hsb.pack(side='bottom', fill='x'); self.tree.pack(side='left', fill='both', expand=True)
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select); self.tree.bind('<Double-1>', self.on_double_click); self.tree_data = {}

    def _sort_column(self, col, reverse):
        """Ordena la tabla por columna."""
        try:
            data = [(self.tree.set(item_id, col), item_id) for item_id in self.tree.get_children('')]
            def sort_key(item):
                v = item[0]
                if v is None: return ""
                try:
                    if isinstance(v, str) and ('-' in v or '/' in v):
                        date_part = v.split(' ')[0]
                        if len(date_part) == 10 and date_part.count('-') == 2: return datetime.strptime(date_part, '%Y-%m-%d')
                        return str(v).lower()
                    return float(v)
                except (ValueError, TypeError): return str(v).lower()
            data.sort(key=sort_key, reverse=reverse); [self.tree.move(item_id, '', index) for index, (val, item_id) in enumerate(data)]; self.tree.heading(col, command=lambda: self._sort_column(col, not reverse))
        except Exception as e: print(f"Error sort '{col}': {e}")

    def _create_bottom_widgets(self, parent):
        """Crea widgets inferiores (stats, botones)."""
        bottom_frame = ctk.CTkFrame(parent); bottom_frame.pack(pady=(10, 0), padx=10, fill="x")
        stats_frame = ctk.CTkFrame(bottom_frame); stats_frame.pack(side="left", padx=(0, 10), pady=5, fill="x", expand=True); ctk.CTkLabel(stats_frame, text="Registros por Enc. Preparación:").pack(side="top", anchor="w", padx=5, pady=(5,0)); self.stats_text = ctk.CTkTextbox(stats_frame, height=80, width=350, state="disabled", wrap="word"); self.stats_text.pack(side="bottom", fill="x", expand=True, padx=5, pady=(0,5))
        action_frame = ctk.CTkFrame(bottom_frame, fg_color="transparent"); action_frame.pack(side="right", padx=(10, 0), pady=5)
        self.import_button = ctk.CTkButton(action_frame, text="Importar Datos\n(Doble Clic)", command=self.import_selected_record, state="disabled", width=120); self.import_button.pack(side="top", anchor="e", padx=5, pady=5)
        self.regenerate_button = ctk.CTkButton(action_frame, text="Regenerar PDF", command=self.initiate_regenerate, state="disabled", width=120); self.regenerate_button.pack(side="top", anchor="e", padx=5, pady=5)
        self.open_pdf_button = ctk.CTkButton(action_frame, text="Abrir PDF", command=self.open_selected_pdf, state="disabled", width=120); self.open_pdf_button.pack(side="top", anchor="e", padx=5, pady=5)
        ctk.CTkButton(action_frame, text="Cerrar", command=self.on_closing, width=120, fg_color="grey").pack(side="top", anchor="e", padx=5, pady=(15,5))

    def on_tree_select(self, event=None):
        """Actualiza estado de botones al seleccionar."""
        sel = self.tree.selection(); one = len(sel) == 1; state = "normal" if one else "disabled"; self.import_button.configure(state=state); self.regenerate_button.configure(state=state)
        if one: item_id = sel[0]; data = self.tree_data.get(item_id); pdf_p = Path(data['archivo_pdf']) if data and data.get('archivo_pdf') else None; self.open_pdf_button.configure(state="normal" if pdf_p and pdf_p.exists() else "disabled")
        else: self.open_pdf_button.configure(state="disabled")

    def on_double_click(self, event=None):
        """Importa registro con doble clic."""
        if len(self.tree.selection()) == 1: self.import_selected_record()

    def load_stats(self):
        """Carga registros y estadísticas en la tabla."""
        print("Cargando stats...");
        self.tree.delete(*self.tree.get_children()); self.tree_data = {}
        self.stats_text.configure(state="normal"); self.stats_text.delete("1.0", "end"); self.stats_text.configure(state="disabled")
        self.import_button.configure(state="disabled"); self.regenerate_button.configure(state="disabled"); self.open_pdf_button.configure(state="disabled")

        d_from = None
        try: date_obj = self.date_from.get_date(); d_from = date_obj.strftime('%Y-%m-%d') if date_obj else None
        except Exception as e: print(f"WARN: Error fecha 'Desde': {e}")
        d_to = None
        try: date_obj = self.date_to.get_date(); d_to = date_obj.strftime('%Y-%m-%d') if date_obj else None
        except Exception as e: print(f"WARN: Error fecha 'Hasta': {e}")

        med=self.medico_filter.get().strip() or None; pac=self.paciente_filter.get().strip() or None; cli=self.cliente_filter.get().strip() or None; uid=self.unique_id_filter.get().strip() or None

        sql = "SELECT * FROM cirugias WHERE 1=1"; params = []
        if d_from: sql += " AND date(fecha_generacion) >= date(?)"; params.append(d_from)
        if d_to: sql += " AND date(fecha_generacion) <= date(?)"; params.append(d_to)
        if med: sql += " AND medico LIKE ? COLLATE NOCASE"; params.append(f"%{med}%")
        if pac: sql += " AND paciente LIKE ? COLLATE NOCASE"; params.append(f"%{pac}%")
        if cli: sql += " AND cliente LIKE ? COLLATE NOCASE"; params.append(f"%{cli}%")
        if uid: sql += " AND unique_id = ?"; params.append(uid)
        sql += " ORDER BY fecha_generacion DESC"; conn = None; results = []
        try:
            conn=sqlite3.connect(DB_FILENAME, timeout=10); conn.row_factory=sqlite3.Row; cursor=conn.cursor()
            print(f"Ejecutando SQL: {sql} con params: {params}")
            cursor.execute(sql, params); results=cursor.fetchall(); print(f"{len(results)} regs found.")
        except sqlite3.Error as e:
            print(f"DB load err: {e}")
            if self.master_app and self.master_app.winfo_exists(): self.master_app.after(0, lambda e=e: show_error_safe("Error DB", f"Error consulta:\n{e}"))
        finally:
             if conn: conn.close()

        if results:
             for row in results:
                 rec=dict(row)
                 dt_gen=rec.get('fecha_generacion','')
                 disp_dt=''
                 if dt_gen:
                     try: disp_dt=datetime.fromisoformat(dt_gen).strftime('%Y-%m-%d %H:%M')
                     except (ValueError, TypeError): print(f"WARN: No se pudo formatear fecha_generacion: {dt_gen}"); disp_dt=dt_gen
                 uid_f=rec.get('unique_id',''); disp_uid=uid_f[:8]+'...' if len(uid_f)>8 else uid_f
                 disp_vals=(disp_dt, rec.get('fecha_cirugia',''), rec.get('paciente',''), rec.get('medico',''), rec.get('tipo_cirugia',''), rec.get('lugar',''), rec.get('encargado_preparacion',''), disp_uid)
                 item_id=self.tree.insert("", tk.END, values=disp_vals)
                 self.tree_data[item_id]=rec

        counts=get_counts_by_preparador(); stats_txt="\n".join([f"- {n or 'N/A'}: {c}" for n,c in counts]) if counts else "No datos."; self.stats_text.configure(state="normal"); self.stats_text.delete("1.0", "end"); self.stats_text.insert("1.0", stats_txt); self.stats_text.configure(state="disabled")

    def get_selected_record_data(self):
        """Obtiene datos del registro seleccionado."""
        sel = self.tree.selection();
        if len(sel)!=1: return None
        item_id = sel[0]; data = self.tree_data.get(item_id)
        if not data: print(f"ERR: No data for item {item_id}"); return None
        return data

    def import_selected_record(self):
        """Importa datos al formulario principal."""
        data = self.get_selected_record_data();
        if not data: show_error("Error", "Selecciona registro."); return
        if self.master_app: self.master_app.load_data_into_form(data); self.on_closing()
        else: show_error("Error", "No acceso a ventana principal.")

    def initiate_regenerate(self):
        """Pide path y luego inicia hilo de regeneración."""
        data = self.get_selected_record_data()
        if not data: show_error("Error", "Selecciona registro."); return
        s_p=self.master_app.sanitize_filename(data.get('paciente','P')); s_c=self.master_app.sanitize_filename(data.get('cliente','C')); s_f=self.master_app.sanitize_filename(data.get('fecha_cirugia',''))
        sugg = f"REPORTE_{s_p}_{s_c}_{s_f}_REGENERADO.pdf"
        if not self.master_app.select_save_path(suggested_filename=sugg):
            show_info("Cancelado", "Regeneración cancelada."); return
        if self.master_app: self.master_app.start_regenerate_thread(data, self.master_app.output_pdf_path_str)
        else: show_error("Error", "No acceso a ventana principal.")

    def open_selected_pdf(self):
        """Abre el PDF del registro seleccionado."""
        data = self.get_selected_record_data();
        if not data: show_error("Error", "Selecciona registro."); return; pdf_s = data.get('archivo_pdf')
        if not pdf_s: show_error("Error", "Registro sin PDF."); return; pdf_p = Path(pdf_s)
        if not pdf_p.exists(): show_error("Error", f"Archivo no encontrado:\n{pdf_p}"); self.open_pdf_button.configure(state="disabled"); return
        try:
            if sys.platform == "win32": os.startfile(str(pdf_p))
            elif sys.platform == "darwin": subprocess.run(["open", str(pdf_p)], check=True)
            else: subprocess.run(["xdg-open", str(pdf_p)], check=True)
        except Exception as e: show_error("Error", f"No abrir PDF:\n{e}")

    def on_closing(self):
        """Cierra la ventana de estadísticas."""
        print("Cerrando stats."); self.grab_release(); self.destroy();
        if self.master_app: self.master_app.stats_window = None

# --- Punto de Entrada ---
if __name__ == "__main__":
    # Configuración DPI para Windows
    if sys.platform == "win32":
        try:
            from ctypes import windll
            try:
                windll.shcore.SetProcessDpiAwareness(2) # Per Monitor V2
                print("DPI Awareness: Per Monitor V2.")
            except AttributeError:
                try:
                    windll.shcore.SetProcessDpiAwareness(1) # Per Monitor V1
                    print("DPI Awareness: Per Monitor V1.")
                except AttributeError:
                    try:
                        windll.user32.SetProcessDPIAware() # System Aware
                        print("DPI Awareness: System Aware (Legacy).")
                    except AttributeError:
                         print("WARN: No se pudo establecer DPI Awareness (ningún método disponible).")
        except ImportError:
             print("WARN: Módulo 'ctypes' no encontrado. No se pudo establecer DPI Awareness.")
        except Exception as e:
            print(f"WARN: Error inesperado al establecer DPI Awareness: {e}")

    # Crear e iniciar la aplicación
    app = App()
    app.mainloop()
    print("App cerrada.")