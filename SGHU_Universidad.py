# -*- coding: utf-8 -*-
"""
===========================================================================
Universidad Simon Rodriguez (UNESR) – Sistema de Oferta Academica | v1.0
Framework: Flet  |  Nucleo: Palo Verde  |  Periodo: 2026-2
Integrado con Backend SQLite, Carga Multi-Formato (.xlsx, .docx, .txt)
y Generador de Reportes HTML / PDF Imprimibles.
===========================================================================
"""
import os
import sys
import shutil
import subprocess
import sqlite3
import unicodedata
import webbrowser
import base64
import threading
from datetime import datetime
import flet as ft
from typing import Dict, List, Optional, Tuple

from database import (
    BASE_DIR,
    DIAS_ORDEN,
    init_db,
    get_connection,
    clear_all_data,
    get_or_create_facultad,
    get_or_create_carrera,
    get_or_create_profesor,
    get_or_create_materia,
    registrar_ficha_docente,
    get_or_create_salon,
    obtener_salones,
    obtener_salon,
    obtener_capacidad_salon,
    actualizar_datos_ambiente,
    eliminar_salon,
    obtener_profesores,
    actualizar_profesor,
    obtener_materias_profesor,
    eliminar_profesor,
    siguiente_codigo_ambiente,
    crear_ambiente,
    codigos_ambientes_predeterminados,
    CAPACIDAD_DEFECTO,
    actualizar_ambiente_por_docente_materia,
    actualizar_ambiente_profesor,
    obtener_info_profesores_dict,
    marcar_iniciar_en_blanco,
    esta_en_modo_blanco,
)
from parsers import (
    importar_excel,
    importar_docx,
    importar_txt,
)
from scheduler import generar_horarios_automaticos

# ── LOGO OFICIAL UNESR OPTIMIZADO ──
LOGO_UNESR_PATH = os.path.join(BASE_DIR, "src", "UNESR.png")
LOGO_UNESR_B64 = ""
if os.path.isfile(LOGO_UNESR_PATH):
    try:
        with open(LOGO_UNESR_PATH, "rb") as _f_logo:
            LOGO_UNESR_B64 = base64.b64encode(_f_logo.read()).decode("utf-8")
    except Exception as _ex_logo:
        print(f"[Aviso Logo UNESR]: {_ex_logo}")


def obtener_logo_unesr(tamano: int = 42, tooltip: str = "") -> ft.Control:
    """Retorna el logo oficial de la UNESR optimizado en memoria para máxima velocidad."""
    if LOGO_UNESR_B64:
        return ft.Image(
            src_base64=LOGO_UNESR_B64,
            width=tamano,
            height=tamano,
            fit=ft.ImageFit.CONTAIN,
            tooltip=tooltip if tooltip else None,
        )
    elif os.path.isfile(LOGO_UNESR_PATH):
        return ft.Image(
            src=LOGO_UNESR_PATH,
            width=tamano,
            height=tamano,
            fit=ft.ImageFit.CONTAIN,
            tooltip=tooltip if tooltip else None,
        )
    return ft.Icon(ft.Icons.SCHOOL_ROUNDED, color=BLUE, size=tamano)


def importar_archivo(filepath: str) -> Dict:
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".xlsx":
        res = importar_excel(filepath)
    elif ext == ".docx":
        res = importar_docx(filepath)
    elif ext in (".txt", ".csv"):
        res = importar_txt(filepath)
    else:
        raise ValueError(f"Formato no soportado: {ext}")
    return {"procesados": res.get("insertados", 0), "errores": res.get("errores", [])}

# ═══════════════════════════════════════════════════════
#  HELPERS para border / padding / margin
# ═══════════════════════════════════════════════════════

def _ba(w: float, c: str) -> ft.Border:
    s = ft.BorderSide(w, c)
    return ft.Border(top=s, right=s, bottom=s, left=s)

def _br(w: float, c: str) -> ft.Border:
    return ft.Border(right=ft.BorderSide(w, c))

def _bb(w: float, c: str) -> ft.Border:
    return ft.Border(bottom=ft.BorderSide(w, c))

def _pa(v: float) -> ft.Padding:
    return ft.Padding(left=v, top=v, right=v, bottom=v)

def _ps(h: float = 0, v: float = 0) -> ft.Padding:
    return ft.Padding(left=h, top=v, right=h, bottom=v)

def _po(left=0, top=0, right=0, bottom=0) -> ft.Padding:
    return ft.Padding(left=left, top=top, right=right, bottom=bottom)

def _ma(v: float) -> ft.Margin:
    return ft.Margin(left=v, top=v, right=v, bottom=v)

def _ms(h: float = 0, v: float = 0) -> ft.Margin:
    return ft.Margin(left=h, top=v, right=h, bottom=v)

def _mo(left=0, top=0, right=0, bottom=0) -> ft.Margin:
    return ft.Margin(left=left, top=top, right=right, bottom=bottom)

# ═══════════════════════════════════════════════════════
#  PALETA Y CONFIGURACIÓN (MODO CLARO INSTITUCIONAL POR DEFECTO)
# ═══════════════════════════════════════════════════════

BG_P   = "#F8FAFC"   # Fondo principal descansado a la vista (Slate-50)
BG_C   = "#FFFFFF"   # Tarjetas y paneles elevados (Blanco puro)
BG_SB  = "#FFFFFF"   # Barra lateral (Blanco institucional)
BG_IN  = "#F1F5F9"   # Inputs y encabezados de tabla (Slate-100 suave)
BG_RA  = "#F8FAFC"   # Filas alternas
BOR    = "#E2E8F0"   # Bordes suaves (Slate-200)

TW     = "#0F172A"   # Títulos y texto de alto contraste (Slate-900)
TL     = "#334155"   # Texto de lectura / cuerpo (Slate-700)
TG     = "#64748B"   # Texto atenuado/etiquetas (Slate-500)
WHITE  = "#FFFFFF"   # Texto blanco para botones con fondo sólido (BLUE, PURP)

BLUE   = "#1D4ED8"   # Azul institucional UNESR
PURP   = "#6D28D9"   # Púrpura universitario
GRN    = "#059669"   # Verde esmeralda (0 conflictos)
ORG    = "#D97706"   # Ámbar/Naranja (ambientes)
RED    = "#DC2626"   # Rojo alerta
CYAN   = "#0284C7"   # Cian informativo

def _set_light_palette():
    global BG_P, BG_C, BG_SB, BG_IN, BG_RA, BOR, TW, TL, TG
    BG_P  = "#F8FAFC"
    BG_C  = "#FFFFFF"
    BG_SB = "#FFFFFF"
    BG_IN = "#F1F5F9"
    BG_RA = "#F8FAFC"
    BOR   = "#E2E8F0"
    TW    = "#0F172A"
    TL    = "#334155"
    TG    = "#64748B"

def _set_dark_palette():
    global BG_P, BG_C, BG_SB, BG_IN, BG_RA, BOR, TW, TL, TG
    BG_P  = "#0D1B2A"
    BG_C  = "#152235"
    BG_SB = "#0A1628"
    BG_IN = "#1A2F45"
    BG_RA = "#12203A"
    BOR   = "#1E3A52"
    TW    = "#FFFFFF"
    TL    = "#B0BEC5"
    TG    = "#78909C"

MAX_EST = CAPACIDAD_DEFECTO

DIAS    = ["LUNES","MARTES","MIERCOLES","JUEVES","VIERNES","SABADO"]
FRANJAS = [
    "07:40-09:20","09:20-10:50","10:50-12:10",
    "12:30-01:50","02:10-03:30","03:40-05:00","05:00-06:20",
]

AMBIENTES_BASE: List[str] = codigos_ambientes_predeterminados()
AMBIENTES: List[str] = list(AMBIENTES_BASE)
MAX_AMB = len(AMBIENTES_BASE)


def recargar_ambientes_desde_bd():
    """Sincroniza la lista de ambientes con el inventario guardado en SQLite."""
    global MAX_AMB
    try:
        codigos = [str(s["codigo"]).strip().upper() for s in obtener_salones()]
    except Exception:
        codigos = []
    if not codigos:
        codigos = list(AMBIENTES_BASE)
    vistos = []
    for codigo in codigos:
        if codigo and codigo not in vistos:
            vistos.append(codigo)
    AMBIENTES.clear()
    AMBIENTES.extend(vistos)
    MAX_AMB = len(vistos)


def recargar_docentes_desde_bd():
    """Sincroniza la lista de docentes con los profesores guardados en SQLite."""
    try:
        nombres = [
            str(p["nombre"]).strip().upper()
            for p in obtener_profesores()
            if str(p.get("nombre") or "").strip()
        ]
    except Exception:
        nombres = []
    DOCENTES.clear()
    DOCENTES.extend(sorted(set(nombres)))

PERIODOS = ["2026-2","2026-1","2025-2","2025-1"]

PALETA_CARRERAS = [
    "#FF6B6B", "#4ECDC4", "#45B7D1", "#96CEB4", "#FFEAA7",
    "#DDA0DD", "#FF9F43", "#54A0FF", "#1DD1A1", "#F368E0",
]

CARRERAS_BASE: Dict[str,Dict] = {
    "EDUC_INIC":   {"nombre":"Educacion Inicial",        "mencion":"Sin Mencion",              "color":"#FF6B6B"},
    "EDUC_INTEG":  {"nombre":"Educacion Integral",       "mencion":"Sin Mencion",              "color":"#4ECDC4"},
    "EDUC_MATEM":  {"nombre":"Educ. Mencion Matematica", "mencion":"Docencia en Matematica",   "color":"#45B7D1"},
    "EDUC_INGLES": {"nombre":"Educ. Mencion Ingles",     "mencion":"Docencia en Ingles",       "color":"#96CEB4"},
    "ORGAN_SIS":   {"nombre":"Organizacion y Sistemas",  "mencion":"Sin Mencion",              "color":"#FFEAA7"},
    "RRHH":        {"nombre":"Recursos Humanos",         "mencion":"Sin Mencion",              "color":"#DDA0DD"},
}
CARRERAS: Dict[str,Dict] = {k: dict(v) for k, v in CARRERAS_BASE.items()}

DOCENTES_BASE = [
    "ADA URSULINA BRITO CARDENAS","MARYELIN ALEJANDRA GONZALEZ",
    "REYNA DEL CARMEN HURTADO DE VILLARROEL","FELIPA ECHEZURIA",
    "LILIANA MERCEDES MONASTERIO BLANCO","INGRID JOSEFINA BORGES OROPEZA",
    "EMMA JOSEFINA MANEIRO","LOURDES ELENA MANRIQUE RIOS",
    "JOSE FRANCISCO LUZONDO RIVAS","ANIBAL JOSE GARCIA ESCORIHUELA",
    "GABRIEL SEGUNDO LEAL PENALOZA","EVA CRISTINA REVERAND PEREZ",
    "JORGE LUIS DE MOYA PUERTA","CARLOS JESUS PACHECO ISAVA",
    "CELMIS COROMOTO GUERRERO MANCILLA","THAIS DEL CARMEN GALEA",
    "ANTONIO DE JESUS DELGADO","NELLY MEDINA DE PRADO",
    "JUSETH FELIX GONZALEZ LAREZ","JOSE DE JESUS GONZALEZ RODRIGUEZ",
    "LAURA PATRICIA SOSA HERNANDEZ","CARLOS ALBERTO FIGUEROA RAMOS",
]
DOCENTES: List[str] = list(DOCENTES_BASE)

MATERIAS_BASE: Dict[str,List[Dict]] = {
    "EDUC_INIC":[
        {"nombre":"ESTRATEGIAS COGNITIVAS",        "horas":3,"ciclo":"INTRODUCTORIO"},
        {"nombre":"INICIACION UNIVERSITARIA",       "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"USO DE LAS TIC",                "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"LENGUAJE Y COMUNICACION",       "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"ETICA VALORES AUTOBIOGRAFIA",   "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"PENSAMIENTO ROBINSONIANO",      "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"DIDACTICA I",                   "horas":3,"ciclo":"BASICO"},
        {"nombre":"PSICOLOGIA DEL DESARROLLO",     "horas":3,"ciclo":"BASICO"},
        {"nombre":"PRACTICA DOCENTE I",            "horas":4,"ciclo":"PROFESIONAL"},
    ],
    "EDUC_INTEG":[
        {"nombre":"SEMINARIO DE INICIACION",       "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"MATEMATICA I",                  "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"LENGUAJE Y COMUNICACION",       "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"ETICA VALORES AUTOBIOGRAFIA",   "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"PENSAMIENTO ROBINSONIANO",      "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"PENSAMIENTO ANDRAGOGICO",       "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"GEOGRAFIA DE VENEZUELA",        "horas":3,"ciclo":"BASICO"},
        {"nombre":"DIDACTICA GENERAL",             "horas":3,"ciclo":"BASICO"},
    ],
    "EDUC_MATEM":[
        {"nombre":"CALCULO IV",                    "horas":3,"ciclo":"PROFESIONAL"},
        {"nombre":"GEOMETRIA II",                  "horas":4,"ciclo":"PROFESIONAL"},
        {"nombre":"ALGEBRA III",                   "horas":3,"ciclo":"PROFESIONAL"},
        {"nombre":"PROYECTO I",                    "horas":6,"ciclo":"PROFESIONAL"},
        {"nombre":"PROYECTO II",                   "horas":6,"ciclo":"PROFESIONAL"},
        {"nombre":"SEMINARIO ESPECIAL DE GRADO",   "horas":4,"ciclo":"PROFESIONAL"},
        {"nombre":"PRECALCULO",                    "horas":3,"ciclo":"INTRODUCTORIO"},
        {"nombre":"ALGEBRA LINEAL",                "horas":3,"ciclo":"BASICO"},
    ],
    "EDUC_INGLES":[
        {"nombre":"INICIACION UNIVERSITARIA",      "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"LENGUAJE Y COMUNICACION",       "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"METODOS Y TECNICAS DE ESTUDIO", "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"PRECALCULO",                    "horas":3,"ciclo":"INTRODUCTORIO"},
        {"nombre":"ETICA VALORES AUTOBIOGRAFIA",   "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"ENGLISH I",                     "horas":4,"ciclo":"BASICO"},
        {"nombre":"ENGLISH II",                    "horas":4,"ciclo":"BASICO"},
        {"nombre":"PHONETICS AND PHONOLOGY",       "horas":3,"ciclo":"PROFESIONAL"},
    ],
    "ORGAN_SIS":[
        {"nombre":"INICIACION UNIVERSITARIA",      "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"LENGUAJE Y COMUNICACION",       "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"METODOS Y TECNICAS DE ESTUDIO", "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"PRECALCULO",                    "horas":3,"ciclo":"INTRODUCTORIO"},
        {"nombre":"ETICA VALORES AUTOBIOGRAFIA",   "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"PENSAMIENTO ROBINSONIANO",      "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"SISTEMAS DE INFORMACION",       "horas":3,"ciclo":"BASICO"},
        {"nombre":"ADMINISTRACION I",              "horas":3,"ciclo":"BASICO"},
    ],
    "RRHH":[
        {"nombre":"INICIACION UNIVERSITARIA",      "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"LENGUAJE Y COMUNICACION",       "horas":2,"ciclo":"INTRODUCTORIO"},
        {"nombre":"RELACIONES INDUSTRIALES",       "horas":3,"ciclo":"BASICO"},
        {"nombre":"PSICOLOGIA ORGANIZACIONAL",     "horas":3,"ciclo":"BASICO"},
        {"nombre":"GESTION DEL TALENTO HUMANO",    "horas":3,"ciclo":"PROFESIONAL"},
        {"nombre":"DERECHO LABORAL",               "horas":3,"ciclo":"PROFESIONAL"},
        {"nombre":"PROYECTO COMUNITARIO",          "horas":6,"ciclo":"PROFESIONAL"},
    ],
}
MATERIAS: Dict[str,List[Dict]] = {k: [dict(m) for m in v] for k, v in MATERIAS_BASE.items()}


def _sin_acentos(txt: str) -> str:
    if not txt:
        return ""
    return "".join(
        c for c in unicodedata.normalize("NFD", str(txt))
        if unicodedata.category(c) != "Mn"
    ).upper().strip()


def _mapear_carrera_id(nombre_carrera: str, facultad_nombre: str = "") -> str:
    nc = _sin_acentos(nombre_carrera)
    if "INICIAL" in nc or "INIC" in nc:
        return "EDUC_INIC"
    if "INTEGRAL" in nc or "INTEG" in nc:
        return "EDUC_INTEG"
    if "MATEM" in nc:
        return "EDUC_MATEM"
    if "INGLES" in nc:
        return "EDUC_INGLES"
    if "ORGANIZACION" in nc or "SISTEMA" in nc or "OYS" in nc:
        return "ORGAN_SIS"
    if "RECURSOS HUMANOS" in nc or "RRHH" in nc:
        return "RRHH"
    # Crear clave dinámica si viene otra carrera desde Excel/Word/TXT (ej. Mercadeo, Informática, RMF)
    clave = "".join(ch if ch.isalnum() else "_" for ch in nc)[:14].strip("_") or "GEN"
    if clave not in CARRERAS:
        color = PALETA_CARRERAS[len(CARRERAS) % len(PALETA_CARRERAS)]
        CARRERAS[clave] = {
            "nombre": nombre_carrera.strip().title(),
            "mencion": facultad_nombre.strip().title() if facultad_nombre else "Pregrado",
            "color": color,
        }
    return clave


# ═══════════════════════════════════════════════════════
#  MOTOR DE HORARIOS SIN CONFLICTOS + BACKEND SQLITE
# ═══════════════════════════════════════════════════════

class Motor:
    def __init__(self):
        self._sa: Dict[Tuple,dict] = {}
        self._sd: Dict[Tuple,bool] = {}
        self._u:  set              = set()
        self._lista: List[Dict]    = []

    def limpiar_memoria(self):
        self._sa.clear()
        self._sd.clear()
        self._u.clear()
        self._lista.clear()

    def cambiar_ambiente(self, entry: Dict, nuevo_ambiente: str):
        nuevo_amb_up = nuevo_ambiente.strip().upper()
        viejo_amb = entry.get("ambiente", "")
        d = entry.get("dia", "")
        f = entry.get("franja", "")
        if (d, f, viejo_amb) in self._sa:
            del self._sa[(d, f, viejo_amb)]
        entry["ambiente"] = nuevo_amb_up
        self._sa[(d, f, nuevo_amb_up)] = entry
        self._u.add(nuevo_amb_up)
        if nuevo_amb_up not in AMBIENTES:
            AMBIENTES.append(nuevo_amb_up)
        actualizar_ambiente_por_docente_materia(
            entry.get("docente", ""),
            entry.get("materia", ""),
            entry.get("dia", ""),
            nuevo_amb_up,
        )

    def _slot(self, doc: str, dia_pref: str = "", franja_pref: str = "", amb_pref: str = "") -> Optional[Dict]:
        # Si viene con día/franja/ambiente desde archivo o preferencia, intentar asignarlo
        d_norm = _sin_acentos(dia_pref)
        if d_norm in DIAS and franja_pref:
            a_elegido = amb_pref.strip().upper() if amb_pref else ""
            if a_elegido and a_elegido not in AMBIENTES:
                AMBIENTES.append(a_elegido)
            if a_elegido and ((d_norm, franja_pref, a_elegido) not in self._sa or "VIRTUAL" in a_elegido or "COORD" in a_elegido):
                return {"dia": d_norm, "franja": franja_pref, "ambiente": a_elegido}
            for a in AMBIENTES:
                if (d_norm, franja_pref, a) not in self._sa:
                    return {"dia": d_norm, "franja": franja_pref, "ambiente": a}

        # Asignación automática sin conflictos
        for d in DIAS:
            for f in FRANJAS:
                if (d, f, doc) in self._sd:
                    continue
                for a in AMBIENTES:
                    if (d, f, a) not in self._sa:
                        return {"dia": d, "franja": f, "ambiente": a}
        return None

    def registrar(self, datos: Dict, guardar_bd: bool = True) -> Dict:
        sl = self._slot(
            datos["docente"],
            datos.get("dia_fijo", ""),
            datos.get("franja_fija", ""),
            datos.get("ambiente_fijo", ""),
        )
        if sl is None:
            return {**datos, "_error": "Sin slots disponibles"}
        try:
            hora_int = int(sl["franja"][:2])
        except Exception:
            hora_int = 8
        am = "AM" if hora_int in (7, 8, 9, 10, 11) else "PM"
        e = {
            **datos,
            "dia": sl["dia"],
            "franja": sl["franja"],
            "ambiente": sl["ambiente"],
            "am_pm": am,
            "cupo": MAX_EST,
            "_error": None,
        }
        # Guardar clave única en _sa (si es curso compartido o virtual, usar sufijo para no sobrescribir)
        key_sa = (sl["dia"], sl["franja"], sl["ambiente"])
        if key_sa in self._sa:
            key_sa = (sl["dia"], sl["franja"], f"{sl['ambiente']}#{len(self._lista)}")
        self._sa[key_sa] = e
        self._sd[(sl["dia"], sl["franja"], datos["docente"])] = True
        self._u.add(sl["ambiente"])
        self._lista.append(e)

        # Actualizar catálogos en memoria para dropdowns y vistas
        if datos["docente"] not in DOCENTES:
            DOCENTES.append(datos["docente"])
            DOCENTES.sort()
        cid = datos["carrera_id"]
        MATERIAS.setdefault(cid, [])
        if not any(m["nombre"].upper() == datos["materia"].upper() for m in MATERIAS[cid]):
            MATERIAS[cid].append({
                "nombre": datos["materia"],
                "horas": int(datos.get("horas", 2)),
                "ciclo": datos.get("ciclo", "INTRODUCTORIO"),
            })

        if guardar_bd:
            self._persistir_en_sqlite(e)
        return e

    def _persistir_en_sqlite(self, e: Dict):
        try:
            conn = get_connection()
            cur = conn.cursor()
            ci = CARRERAS.get(e["carrera_id"], {"nombre": e["carrera_id"], "mencion": "General"})
            fac_nombre = "Facultad de Ciencias de la Educación" if "EDUC" in e["carrera_id"] else "Facultad de Ciencias Económicas y Sociales"
            fac_id = get_or_create_facultad(cur, fac_nombre)
            car_id = get_or_create_carrera(cur, ci["nombre"], fac_id)
            prof_id = get_or_create_profesor(cur, e["docente"])
            mat_id = get_or_create_materia(cur, e["materia"], car_id, semestre=str(e.get("ciclo", "INTRODUCTORIO")))
            sal_id = get_or_create_salon(cur, e["ambiente"], capacidad=MAX_EST, edificio="Núcleo Palo Verde")

            partes_f = e["franja"].split("-")
            h_ini = partes_f[0].strip() if len(partes_f) > 0 else "07:40"
            h_fin = partes_f[1].strip() if len(partes_f) > 1 else "09:20"
            dia_bd = {
                "LUNES": "Lunes", "MARTES": "Martes", "MIERCOLES": "Miércoles",
                "JUEVES": "Jueves", "VIERNES": "Viernes", "SABADO": "Sábado",
            }.get(e["dia"], e["dia"].title())

            cur.execute(
                """
                INSERT OR IGNORE INTO disponibilidades (profesor_id, dia, hora_inicio, hora_fin)
                VALUES (?, ?, ?, ?)
                """,
                (prof_id, dia_bd, h_ini, h_fin),
            )
            cur.execute(
                """
                INSERT INTO carga_academica
                (profesor_id, materia_id, seccion, bloques_semanales, salon_preferido, dia_fijo, hora_inicio_fija, hora_fin_fija)
                VALUES (?, ?, ?, 1, ?, ?, ?, ?)
                ON CONFLICT(profesor_id, materia_id, seccion) DO UPDATE SET
                    salon_preferido = excluded.salon_preferido,
                    dia_fijo = excluded.dia_fijo,
                    hora_inicio_fija = excluded.hora_inicio_fija,
                    hora_fin_fija = excluded.hora_fin_fija
                """,
                (prof_id, mat_id, str(e.get("seccion", "10100")), e["ambiente"], dia_bd, h_ini, h_fin),
            )
            cur.execute(
                "SELECT id FROM carga_academica WHERE profesor_id=? AND materia_id=? AND seccion=?",
                (prof_id, mat_id, str(e.get("seccion", "10100"))),
            )
            row_c = cur.fetchone()
            if row_c:
                cur.execute(
                    """
                    INSERT INTO asignaciones_horario
                    (carga_id, profesor_id, materia_id, salon_id, dia, hora_inicio, hora_fin, estado)
                    VALUES (?, ?, ?, ?, ?, ?, ?, 'Confirmado')
                    """,
                    (row_c["id"], prof_id, mat_id, sal_id, dia_bd, h_ini, h_fin),
                )
            conn.commit()
            conn.close()
        except Exception as ex:
            print(f"Aviso SQLite: {ex}")

    def sincronizar_desde_sqlite(self) -> List[Dict]:
        """Carga en el Motor todos los horarios almacenados en SQLite (por ejemplo tras importar Excel/Word/TXT)."""
        self.limpiar_memoria()
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("""
            SELECT
                p.nombre AS docente,
                m.nombre AS materia,
                m.semestre AS ciclo,
                c.nombre AS carrera_nombre,
                f.nombre AS facultad_nombre,
                ca.seccion AS seccion,
                ca.bloques_semanales AS horas,
                s.codigo AS ambiente,
                ah.dia AS dia,
                ah.hora_inicio AS hora_inicio,
                ah.hora_fin AS hora_fin
            FROM asignaciones_horario ah
            JOIN profesores p ON p.id = ah.profesor_id
            JOIN materias m ON m.id = ah.materia_id
            JOIN carreras c ON c.id = m.carrera_id
            JOIN facultades f ON f.id = c.facultad_id
            JOIN carga_academica ca ON ca.id = ah.carga_id
            JOIN salones s ON s.id = ah.salon_id
            ORDER BY p.nombre ASC, c.nombre ASC, ah.dia ASC, ah.hora_inicio ASC
        """)
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()

        for r in rows:
            cid = _mapear_carrera_id(r["carrera_nombre"], r["facultad_nombre"])
            franja = f"{r['hora_inicio']}-{r['hora_fin']}"
            if franja not in FRANJAS:
                FRANJAS.append(franja)
            dia_up = _sin_acentos(r["dia"])
            if dia_up not in DIAS and dia_up:
                DIAS.append(dia_up)
            amb_up = (r["ambiente"] or "P1-01").strip().upper()
            if amb_up not in AMBIENTES:
                AMBIENTES.append(amb_up)

            ciclo_raw = _sin_acentos(r["ciclo"] or "INTRODUCTORIO")
            if "PROF" in ciclo_raw:
                ciclo_limpio = "PROFESIONAL"
            elif "BASI" in ciclo_raw or "GEN" in ciclo_raw:
                ciclo_limpio = "BASICO"
            else:
                ciclo_limpio = "INTRODUCTORIO"

            datos = {
                "carrera_id": cid,
                "docente": r["docente"].strip().upper(),
                "materia": r["materia"].strip().upper(),
                "horas": int(r["horas"] or 2),
                "ciclo": ciclo_limpio,
                "seccion": r["seccion"] or "10100",
                "periodo": "2026-2",
                "dia_fijo": dia_up,
                "franja_fija": franja,
                "ambiente_fijo": amb_up,
            }
            self.registrar(datos, guardar_bd=False)

        DOCENTES.sort()
        return list(self._lista)

    def entradas(self) -> List[Dict]:
        return list(self._lista)

    def usados(self) -> set:
        return set(self._u)

    def stats(self) -> Dict:
        t = self.entradas()
        total_amb = max(MAX_AMB, len(AMBIENTES))
        return {
            "cursos":  len(t),
            "amb_uso": len(self._u),
            "amb_lib": max(0, total_amb - len(self._u)),
            "amb_tot": total_amb,
            "docs":    len({e["docente"] for e in t}),
        }


ROWS_SEED = [
    ("EDUC_INIC",  "ADA URSULINA BRITO CARDENAS",            "ESTRATEGIAS COGNITIVAS",       3,"INTRODUCTORIO",10100),
    ("EDUC_INIC",  "ADA URSULINA BRITO CARDENAS",            "INICIACION UNIVERSITARIA",      2,"INTRODUCTORIO",10100),
    ("EDUC_INIC",  "MARYELIN ALEJANDRA GONZALEZ",            "USO DE LAS TIC",               2,"INTRODUCTORIO",10100),
    ("EDUC_INIC",  "REYNA DEL CARMEN HURTADO DE VILLARROEL", "LENGUAJE Y COMUNICACION",      2,"INTRODUCTORIO",10100),
    ("EDUC_INIC",  "FELIPA ECHEZURIA",                       "ETICA VALORES AUTOBIOGRAFIA",  2,"INTRODUCTORIO",10100),
    ("EDUC_INIC",  "FELIPA ECHEZURIA",                       "PENSAMIENTO ROBINSONIANO",     2,"INTRODUCTORIO",10100),
    ("EDUC_INTEG", "LILIANA MERCEDES MONASTERIO BLANCO",     "SEMINARIO DE INICIACION",      2,"INTRODUCTORIO",10100),
    ("EDUC_INTEG", "INGRID JOSEFINA BORGES OROPEZA",         "MATEMATICA I",                 2,"INTRODUCTORIO",10100),
    ("EDUC_INTEG", "EMMA JOSEFINA MANEIRO",                  "LENGUAJE Y COMUNICACION",      2,"INTRODUCTORIO",10100),
    ("EDUC_INTEG", "LOURDES ELENA MANRIQUE RIOS",            "ETICA VALORES AUTOBIOGRAFIA",  2,"INTRODUCTORIO",10100),
    ("EDUC_INTEG", "JOSE FRANCISCO LUZONDO RIVAS",           "PENSAMIENTO ROBINSONIANO",     2,"INTRODUCTORIO",10100),
    ("EDUC_INTEG", "JOSE FRANCISCO LUZONDO RIVAS",           "PENSAMIENTO ANDRAGOGICO",      2,"INTRODUCTORIO",10100),
    ("EDUC_MATEM", "ANIBAL JOSE GARCIA ESCORIHUELA",         "CALCULO IV",                   3,"PROFESIONAL",  20112),
    ("EDUC_MATEM", "GABRIEL SEGUNDO LEAL PENALOZA",          "GEOMETRIA II",                 4,"PROFESIONAL",  20102),
    ("EDUC_MATEM", "ANIBAL JOSE GARCIA ESCORIHUELA",         "ALGEBRA III",                  3,"PROFESIONAL",  20203),
    ("EDUC_MATEM", "EVA CRISTINA REVERAND PEREZ",            "PROYECTO I",                   6,"PROFESIONAL",  20203),
    ("EDUC_MATEM", "EVA CRISTINA REVERAND PEREZ",            "PROYECTO II",                  6,"PROFESIONAL",  20111),
    ("EDUC_MATEM", "EVA CRISTINA REVERAND PEREZ",            "SEMINARIO ESPECIAL DE GRADO",  4,"PROFESIONAL",  20212),
    ("EDUC_INGLES","JORGE LUIS DE MOYA PUERTA",              "INICIACION UNIVERSITARIA",      2,"INTRODUCTORIO",10200),
    ("EDUC_INGLES","CARLOS JESUS PACHECO ISAVA",             "LENGUAJE Y COMUNICACION",      2,"INTRODUCTORIO",10200),
    ("EDUC_INGLES","ADA URSULINA BRITO CARDENAS",            "METODOS Y TECNICAS DE ESTUDIO",2,"INTRODUCTORIO",10200),
    ("EDUC_INGLES","CELMIS COROMOTO GUERRERO MANCILLA",      "PRECALCULO",                   3,"INTRODUCTORIO",10200),
    ("EDUC_INGLES","LOURDES ELENA MANRIQUE RIOS",            "ETICA VALORES AUTOBIOGRAFIA",  2,"INTRODUCTORIO",10200),
    ("EDUC_INGLES","THAIS DEL CARMEN GALEA",                 "PENSAMIENTO ROBINSONIANO",     2,"INTRODUCTORIO",10200),
    ("ORGAN_SIS",  "ANTONIO DE JESUS DELGADO",               "INICIACION UNIVERSITARIA",      2,"INTRODUCTORIO",20100),
    ("ORGAN_SIS",  "CARLOS JESUS PACHECO ISAVA",             "LENGUAJE Y COMUNICACION",      2,"INTRODUCTORIO",20100),
    ("ORGAN_SIS",  "NELLY MEDINA DE PRADO",                  "METODOS Y TECNICAS DE ESTUDIO",2,"INTRODUCTORIO",20100),
    ("ORGAN_SIS",  "JUSETH FELIX GONZALEZ LAREZ",            "PRECALCULO",                   3,"INTRODUCTORIO",20100),
    ("ORGAN_SIS",  "ANTONIO DE JESUS DELGADO",               "ETICA VALORES AUTOBIOGRAFIA",  2,"INTRODUCTORIO",20100),
    ("ORGAN_SIS",  "JOSE DE JESUS GONZALEZ RODRIGUEZ",       "PENSAMIENTO ROBINSONIANO",     2,"INTRODUCTORIO",20100),
    ("RRHH",       "LAURA PATRICIA SOSA HERNANDEZ",          "INICIACION UNIVERSITARIA",      2,"INTRODUCTORIO",30100),
    ("RRHH",       "CARLOS ALBERTO FIGUEROA RAMOS",          "LENGUAJE Y COMUNICACION",      2,"INTRODUCTORIO",30100),
    ("RRHH",       "GABRIEL SEGUNDO LEAL PENALOZA",          "RELACIONES INDUSTRIALES",      3,"BASICO",        30201),
]


def _seed_data(forzar_base: bool = False) -> Tuple[Motor, List[Dict]]:
    init_db()
    recargar_ambientes_desde_bd()
    m = Motor()
    if not forzar_base:
        existentes = m.sincronizar_desde_sqlite()
        if existentes:
            return m, existentes
        # Por defecto, el sistema se mantiene en blanco listo para registrar cursos nuevos
        marcar_iniciar_en_blanco(True)
        CARRERAS.clear()
        CARRERAS.update({k: dict(v) for k, v in CARRERAS_BASE.items()})
        recargar_ambientes_desde_bd()
        DOCENTES.clear()
        return m, []

    marcar_iniciar_en_blanco(False)
    clear_all_data(keep_salones=True)
    CARRERAS.clear()
    CARRERAS.update({k: dict(v) for k, v in CARRERAS_BASE.items()})
    DOCENTES.clear()
    DOCENTES.extend(sorted(DOCENTES_BASE))
    recargar_ambientes_desde_bd()

    datos = [
        {"carrera_id":r[0],"docente":r[1],"materia":r[2],"horas":r[3], "ciclo":r[4],"seccion":r[5],"periodo":"2026-2"}
        for r in ROWS_SEED
    ]
    entries = [m.registrar(d, guardar_bd=True) for d in datos]
    return m, entries


# ═══════════════════════════════════════════════════════
#  GENERADOR DE REPORTES HTML / PDF IMPRIMIBLES
# ═══════════════════════════════════════════════════════

def generar_reporte_html_imprimible(
    entradas: List[Dict],
    tipo: str = "completo",
    docente_filtro: Optional[str] = None,
) -> str:
    """
    Genera un archivo HTML oficial listo para imprimir o Guardar como PDF:
    - tipo == 'oferta'   : 1) Oferta Académica General por carrera con profesores ordenados.
    - tipo == 'profesor' : 2) Horario tipo calendario semanal (de TODOS los profesores o de 1 SOLO profesor).
    - tipo == 'completo' : Oferta Académica General + Calendarios de Profesores.
    """
    carpeta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reportes_imprimibles")
    os.makedirs(carpeta, exist_ok=True)
    fecha_str = datetime.now().strftime("%d/%m/%Y %H:%M")

    # Profesores ordenados alfabéticamente
    todos_docs = sorted({e["docente"] for e in entradas})
    if docente_filtro and docente_filtro != "TODOS":
        docs_mostrar = [d for d in todos_docs if d == docente_filtro]
    else:
        docs_mostrar = todos_docs

    html = [f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Reporte Oficial Universidad Simón Rodríguez - Oferta Académica y Horarios 2026-2</title>
<style>
    * {{ box-sizing: border-box; }}
    body {{
        font-family: 'Segoe UI', Arial, sans-serif;
        color: #0D1B2A;
        background: #FFFFFF;
        margin: 0;
        padding: 24px;
        font-size: 12px;
    }}
    .no-print {{
        background: #0D1B2A;
        color: #FFFFFF;
        padding: 14px 20px;
        border-radius: 10px;
        margin-bottom: 22px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }}
    .btn-print {{
        background: #00E676;
        color: #0D1B2A;
        border: none;
        padding: 9px 18px;
        font-weight: bold;
        border-radius: 6px;
        cursor: pointer;
        font-size: 13px;
    }}
    .header {{
        text-align: center;
        border-bottom: 3px solid #0D1B2A;
        padding-bottom: 12px;
        margin-bottom: 20px;
    }}
    .header h1 {{ margin: 0; font-size: 20px; color: #0D1B2A; text-transform: uppercase; }}
    .header p {{ margin: 3px 0; color: #455A64; font-size: 12px; }}
    .sec-title {{
        background: #152235;
        color: #FFFFFF;
        padding: 9px 14px;
        font-weight: bold;
        font-size: 14px;
        border-left: 5px solid #2196F3;
        margin: 22px 0 12px 0;
        border-radius: 4px;
    }}
    .car-header {{
        background: #E3F2FD;
        color: #0D47A1;
        padding: 7px 12px;
        font-weight: bold;
        font-size: 13px;
        border-left: 4px solid #2196F3;
        margin-top: 14px;
    }}
    table {{
        width: 100%;
        border-collapse: collapse;
        margin-bottom: 16px;
    }}
    th, td {{
        border: 1px solid #CFD8DC;
        padding: 6px 8px;
        text-align: left;
        vertical-align: top;
    }}
    th {{
        background: #ECEFF1;
        color: #0D1B2A;
        font-weight: bold;
        font-size: 11px;
        text-transform: uppercase;
    }}
    .cal-box {{
        page-break-inside: avoid;
        margin-bottom: 26px;
        border: 1px solid #B0BEC5;
        border-radius: 6px;
        overflow: hidden;
    }}
    .cal-head {{
        background: #0D1B2A;
        color: #FFFFFF;
        padding: 8px 14px;
        font-weight: bold;
        font-size: 13px;
    }}
    .slot-item {{
        background: #F1F8E9;
        border-left: 3px solid #2196F3;
        padding: 4px 6px;
        margin-bottom: 4px;
        font-size: 10.5px;
    }}
    @media print {{
        .no-print {{ display: none !important; }}
        body {{ padding: 0; }}
    }}
</style>
</head>
<body>
    <div class="no-print">
        <div>
            <strong>Universidad Simón Rodríguez (UNESR) - Núcleo Palo Verde | Vista Oficial de Impresión HTML / PDF</strong><br>
            <span style="font-size:11px;color:#B0BEC5;">Usa el botón derecho para imprimir en papel o elegir "Guardar como PDF".</span>
        </div>
        <button class="btn-print" onclick="window.print()">Imprimir / Guardar como PDF</button>
    </div>

    <div class="header">
        <h1>Universidad Nacional Experimental Simón Rodríguez (UNESR) — Núcleo Palo Verde</h1>
        <p><strong>Sistema de Oferta Académica y Horarios por Carrera y Docente — Periodo 2026-2</strong></p>
        <p>Generado el: {fecha_str} &bull; Total asignaciones: {len(entradas)} &bull; Conflictos: 0</p>
    </div>
"""]

    # 1) OFERTA ACADÉMICA GENERAL
    if tipo in ("oferta", "completo"):
        html.append('<div class="sec-title">1) OFERTA ACADÉMICA GENERAL POR CARRERA (PROFESORES ORDENADOS)</div>')
        for kid, cinfo in CARRERAS.items():
            cursos_car = [e for e in entradas if e.get("carrera_id") == kid]
            if not cursos_car:
                continue
            cursos_car.sort(key=lambda x: (x["docente"], x["dia"], x["franja"]))
            html.append(f'<div class="car-header">{cinfo["nombre"]} ({cinfo.get("mencion","")}) — {len(cursos_car)} cursos</div>')
            html.append("""
            <table>
                <thead>
                    <tr>
                        <th style="width:26%;">Docente (Orden Alfabético)</th>
                        <th style="width:26%;">Materia / Curso</th>
                        <th style="width:12%;">Ciclo</th>
                        <th style="width:8%;">Sección</th>
                        <th style="width:10%;">Día</th>
                        <th style="width:10%;">Horario</th>
                        <th style="width:8%;">Ambiente</th>
                    </tr>
                </thead>
                <tbody>
            """)
            for e in cursos_car:
                html.append(f"""
                    <tr>
                        <td><strong>{e['docente']}</strong></td>
                        <td>{e['materia']}</td>
                        <td>{e.get('ciclo','')}</td>
                        <td>{e.get('seccion','')}</td>
                        <td>{e['dia']}</td>
                        <td>{e['franja']}</td>
                        <td><strong>{e['ambiente']}</strong></td>
                    </tr>
                """)
            html.append("</tbody></table>")

    # 2) HORARIO TIPO CALENDARIO DE CADA PROFESOR
    if tipo in ("profesor", "completo"):
        sub_t = (
            f"PROFESOR SELECCIONADO: {docente_filtro}"
            if docente_filtro and docente_filtro != "TODOS"
            else f"TODOS LOS PROFESORES REGISTRADOS ({len(docs_mostrar)})"
        )
        html.append(f'<div class="sec-title">2) HORARIO TIPO CALENDARIO SEMANAL — {sub_t}</div>')
        dias_cols = ["LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO"]

        for doc in docs_mostrar:
            asig_doc = [e for e in entradas if e.get("docente") == doc]
            if not asig_doc:
                continue
            franjas_doc = sorted(
                list({e["franja"] for e in asig_doc} | set(FRANJAS[:7])),
                key=lambda f: FRANJAS.index(f) if f in FRANJAS else 99,
            )
            html.append(f'<div class="cal-box">')
            html.append(f'<div class="cal-head">DOCENTE: {doc} ({len(asig_doc)} bloque(s) asignados)</div>')
            html.append("<table><thead><tr><th style='width:13%;'>HORARIO</th>")
            for d in dias_cols:
                html.append(f"<th style='width:14.5%;'>{d}</th>")
            html.append("</tr></thead><tbody>")

            for fr in franjas_doc:
                # Mostrar fila solo si está en las franjas base o tiene clase
                celdas_fila = {d: [x for x in asig_doc if x["dia"] == d and x["franja"] == fr] for d in dias_cols}
                if fr not in FRANJAS[:7] and not any(celdas_fila.values()):
                    continue
                html.append(f"<tr><td><strong>{fr}</strong></td>")
                for d in dias_cols:
                    items = celdas_fila[d]
                    html.append("<td>")
                    for it in items:
                        cnom = CARRERAS.get(it["carrera_id"], {}).get("nombre", it["carrera_id"])
                        html.append(
                            f"<div class='slot-item'>"
                            f"<strong>{it['materia']}</strong><br>"
                            f"<span>{cnom}</span><br>"
                            f"<span>Salón: <strong>{it['ambiente']}</strong> (Sec. {it.get('seccion','')})</span>"
                            f"</div>"
                        )
                    html.append("</td>")
                html.append("</tr>")
            html.append("</tbody></table></div>")

    html.append("</body></html>")
    nombre_arch = f"reporte_{tipo}_{datetime.now().strftime('%H%M%S')}.html"
    ruta_out = os.path.join(carpeta, nombre_arch)
    with open(ruta_out, "w", encoding="utf-8") as f:
        f.write("\n".join(html))

    webbrowser.open(f"file:///{ruta_out.replace(os.sep, '/')}")
    return ruta_out


# ═══════════════════════════════════════════════════════
#  WIDGET HELPERS
# ═══════════════════════════════════════════════════════

def alpha_col(col: str, op: float = 0.15) -> str:
    """
    Retorna el formato de color con opacidad soportado por Flet ('#RRGGBB,0.15'),
    evitando la corrupción de canales RGB producida por concatenación hex de 8 dígitos.
    """
    if not col:
        return col
    c = str(col).strip()
    if c.startswith("#") and len(c) == 9:
        c = c[:7]
    if c.startswith("#") and len(c) == 7:
        return f"{c},{op}"
    return c

def badge(txt: str, col: str) -> ft.Container:
    return ft.Container(
        content=ft.Text(txt, color=col, size=10, weight=ft.FontWeight.W_700),
        bgcolor=alpha_col(col, 0.18), border_radius=6, padding=_ps(8, 3),
        border=_ba(1, alpha_col(col, 0.35)),
    )

def divl() -> ft.Container:
    return ft.Container(height=1, bgcolor=BOR)

def titulo(txt: str, sub: str = "") -> ft.Column:
    c: List[ft.Control] = [ft.Text(txt, color=TW, size=26, weight=ft.FontWeight.BOLD)]
    if sub:
        c.append(ft.Text(sub, color=TG, size=13))
    return ft.Column(c, spacing=4)

def stat_card(tit: str, val: str, ico, col: str, nota: str) -> ft.Container:
    return ft.Container(
        content=ft.Column([
            ft.Row([
                ft.Container(
                    content=ft.Icon(ico, color=col, size=20),
                    bgcolor=alpha_col(col, 0.18), border_radius=10,
                    padding=10, width=42, height=42,
                ),
                ft.Container(expand=True),
                ft.Container(
                    content=ft.Text(nota, color=col, size=10, weight=ft.FontWeight.W_600),
                    bgcolor=alpha_col(col, 0.18), border_radius=6, padding=_ps(8, 3),
                    border=_ba(1, alpha_col(col, 0.35)),
                )
            ], vertical_alignment=ft.CrossAxisAlignment.CENTER),
            ft.Container(height=12),
            ft.Text(val,  color=TW, size=28, weight=ft.FontWeight.BOLD),
            ft.Text(tit,  color=TL, size=12),
        ], spacing=0),
        bgcolor=BG_C, border_radius=14,
        padding=_pa(18), border=_ba(1, BOR), expand=True,
    )

def campo(lbl: str, **kw) -> ft.TextField:
    params = {
        "label": lbl,
        "bgcolor": BG_IN,
        "border_color": BOR,
        "focused_border_color": BLUE,
        "color": TW,
        "label_style": ft.TextStyle(color=TG),
        "border_radius": 10,
    }
    params.update(kw)
    return ft.TextField(**params)

def ddrop(lbl: str, opts, **kw) -> ft.Dropdown:
    params = {
        "label": lbl,
        "options": opts,
        "bgcolor": BG_IN,
        "border_color": BOR,
        "focused_border_color": BLUE,
        "color": TW,
        "label_style": ft.TextStyle(color=TG),
        "border_radius": 10,
    }
    params.update(kw)
    return ft.Dropdown(**params)

def categorizar_piso_ambiente(codigo: str) -> Tuple[str, str, int]:
    """
    Retorna (piso_categoria, codigo_normalizado, prioridad_orden)
    Pisos: Piso 1, Piso 2, Piso 3, Piso 4, Satelitales, Planta Baja (PB), Virtuales, Laboratorios y Otros
    """
    c = str(codigo).strip().strip("´'`").upper()
    for prefix in ["P1.", "P2.", "P3.", "P4."]:
        if c.startswith(prefix):
            c = prefix[:2] + "-" + c[3:]

    if c.startswith("P1-") or c.startswith("P1") or "PISO 1" in c or c.startswith("AULA 1"):
        return "Piso 1", c, 1
    elif c.startswith("P2-") or c.startswith("P2") or "PISO 2" in c or c.startswith("AULA 2"):
        return "Piso 2", c, 2
    elif c.startswith("P3-") or c.startswith("P3") or "PISO 3" in c or c.startswith("AULA 3") or "MAGNA" in c:
        return "Piso 3", c, 3
    elif c.startswith("P4-") or c.startswith("P4") or "PISO 4" in c or c.startswith("AULA 4"):
        return "Piso 4", c, 4
    elif "SATELITAL" in c:
        return "Satelitales", c, 5
    elif c.startswith("PB") or "PLANTA BAJA" in c:
        return "Planta Baja (PB)", c, 6
    elif "VIRTUAL" in c or "ONLINE" in c:
        return "Virtuales", c, 7
    else:
        return "Laboratorios y Otros", c, 8

def clave_orden_ambiente(codigo: str):
    import re
    return [int(t) if t.isdigit() else t for t in re.split(r'(\d+)', str(codigo).upper())]


def abrir_selector_archivo_os(titulo: str = "Seleccionar archivo", extensiones: Optional[List[str]] = None) -> Optional[str]:
    """
    Abre un diálogo nativo de selección de archivos en Linux (Zenity / Kdialog),
    Windows (PowerShell Forms) o macOS (osascript), sin depender de libtk8.6.so.
    Retorna:
      - str con la ruta si el usuario seleccionó un archivo.
      - "" (cadena vacía) si el usuario canceló la selección voluntariamente.
      - None si ningún diálogo gráfico nativo pudo ejecutarse en el SO.
    """
    if extensiones is None:
        extensiones = [".xlsx", ".docx", ".txt", ".csv"]

    # 1. En Linux: intentar Zenity (estándar en GNOME, XFCE, Canaima, Debian, Arch y Wayland)
    if shutil.which("zenity"):
        try:
            cmd = ["zenity", "--file-selection", f"--title={titulo}"]
            patrones = " ".join("*" + (e if e.startswith(".") else "." + e) for e in extensiones)
            cmd.append(f"--file-filter=Archivos soportados | {patrones}")
            cmd.append("--file-filter=Todos los archivos | *")
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
            if res.returncode == 0 and res.stdout.strip():
                return res.stdout.strip()
            elif res.returncode == 1 or not res.stdout.strip():
                return ""  # Cancelado por el usuario
        except Exception as ex:
            print(f"[Aviso Zenity]: {ex}")

    # 2. En Linux: intentar kdialog (KDE Plasma)
    if shutil.which("kdialog"):
        try:
            patrones = " ".join("*" + (e if e.startswith(".") else "." + e) for e in extensiones)
            cmd = ["kdialog", "--getopenfilename", os.path.expanduser("~"), patrones]
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
            if res.returncode == 0 and res.stdout.strip():
                return res.stdout.strip()
            elif res.returncode == 1:
                return ""
        except Exception as ex:
            print(f"[Aviso Kdialog]: {ex}")

    # 3. En Windows: intentar PowerShell OpenFileDialog nativo
    if sys.platform == "win32":
        try:
            patrones = ";".join("*" + (e if e.startswith(".") else "." + e) for e in extensiones)
            ps_script = (
                '[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null; '
                '$f = New-Object System.Windows.Forms.OpenFileDialog; '
                f'$f.Title = "{titulo}"; '
                f'$f.Filter = "Archivos soportados ({patrones})|{patrones}|Todos los archivos (*.*)|*.*"; '
                'if ($f.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $f.FileName }'
            )
            res = subprocess.run(["powershell", "-NoProfile", "-Command", ps_script], capture_output=True, text=True, timeout=120)
            if res.returncode == 0 and res.stdout.strip():
                return res.stdout.strip()
            elif res.returncode == 0 and not res.stdout.strip():
                return ""
        except Exception as ex:
            print(f"[Aviso Windows OpenFileDialog]: {ex}")

    # 4. En macOS: intentar osascript
    if sys.platform == "darwin":
        try:
            res = subprocess.run(
                ["osascript", "-e", f'POSIX path of (choose file with prompt "{titulo}")'],
                capture_output=True, text=True, timeout=120
            )
            if res.returncode == 0 and res.stdout.strip():
                return res.stdout.strip()
            return ""
        except Exception:
            pass

    # 5. Intentar tkinter protegido contra ImportError / missing libtk
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        ruta = filedialog.askopenfilename(title=titulo)
        root.destroy()
        if ruta:
            return ruta
        return ""
    except Exception:
        pass

    return None


def buscar_archivos_recientes_sistema(extensiones: Optional[Tuple[str, ...]] = None) -> List[str]:
    """Busca archivos recientes en las carpetas comunes del usuario (Descargas, Escritorio, etc.)."""
    if extensiones is None:
        extensiones = (".xlsx", ".docx", ".txt", ".csv")
    carpetas = [
        os.path.expanduser("~/Descargas"),
        os.path.expanduser("~/Downloads"),
        os.path.expanduser("~/Escritorio"),
        os.path.expanduser("~/Desktop"),
        BASE_DIR,
        os.path.join(BASE_DIR, "uploads"),
    ]
    encontrados = []
    for c in carpetas:
        if os.path.isdir(c):
            try:
                for f in os.listdir(c):
                    if any(f.lower().endswith(ext) for ext in extensiones):
                        full_p = os.path.join(c, f)
                        if os.path.isfile(full_p):
                            encontrados.append((os.path.getmtime(full_p), full_p))
            except Exception:
                pass
    encontrados.sort(reverse=True, key=lambda x: x[0])
    vistos = set()
    unicos = []
    for _, p in encontrados:
        if p not in vistos:
            vistos.add(p)
            unicos.append(p)
    return unicos[:8]


# ═══════════════════════════════════════════════════════
#  APP
# ═══════════════════════════════════════════════════════

class App:
    NAV = [
        ("dashboard",   ft.Icons.DASHBOARD_ROUNDED,    "Dashboard"),
        ("nueva_carga", ft.Icons.ADD_CIRCLE_ROUNDED,   "Nueva Carga"),
        ("horarios",    ft.Icons.CALENDAR_MONTH,        "Horarios"),
        ("ambientes",   ft.Icons.MEETING_ROOM_ROUNDED,  "Ambientes"),
        ("docentes",    ft.Icons.PEOPLE_ROUNDED,        "Docentes"),
        ("carreras",    ft.Icons.SCHOOL_ROUNDED,        "Carreras"),
    ]

    def __init__(self, page: ft.Page):
        self.page    = page
        self.motor, self._all = _seed_data()
        self.vista   = "dashboard"
        self.ultimo_mensaje = ""
        self.docente_calendario_sel = DOCENTES[0] if DOCENTES else ""

        # Estados de paginación y filtros para Horarios
        self.horarios_vista_modo = "tabla"
        self.horarios_pag = 1
        self.horarios_rpp = 25
        self.horarios_buscar = ""
        self.horarios_carrera = "TODAS"
        self.horarios_dia = "TODOS"
        self.horarios_docente = "TODOS"

        # Estados de filtros para Ambientes
        self.ambientes_piso_sel = "TODOS"
        self.ambientes_estado_sel = "TODOS"
        self.ambientes_buscar = ""

        # Estado de búsqueda y paginación para Docentes
        self.docentes_buscar = ""
        self.docentes_pag = 1
        self.docentes_rpp = 15

        # Selector de archivos para importación masiva de datos tabulares.
        self.fp_arch = ft.FilePicker(on_result=self._on_fp_arch_result)

        # Estado del menú lateral: expandido (True) o contraído (False)
        self.sb_expandido = True

        self._setup()
        self._build()
        self._nav("dashboard")

    def _setup(self):
        p = self.page
        p.title      = "Universidad Nacional Experimental Simón Rodríguez (UNESR) - Oferta Académica 2026-2"
        p.theme_mode = ft.ThemeMode.LIGHT
        p.bgcolor    = BG_P
        p.padding    = 0
        p.spacing    = 0
        try:
            p.window.icon = "UNESR.png"
        except Exception as ex:
            print(f"[Aviso icono de ventana]: {ex}")
        try:
            p.window.width      = 1360
            p.window.height     = 860
            p.window.min_width  = 960
            p.window.min_height = 650
            p.window.to_front()
        except Exception as ex:
            print(f"[Aviso ventana]: {ex}")

    def _build(self):
        self.page.overlay.append(self.fp_arch)
        self.sb_col = ft.Column(spacing=0, expand=True)
        self.sb     = ft.Container(
            content=self.sb_col,
            width=250 if self.sb_expandido else 72,
            bgcolor=BG_SB,
            border=_br(1, BOR),
            animate=ft.Animation(200, ft.AnimationCurve.EASE_OUT),
        )
        self.mc_col = ft.Column(scroll=ft.ScrollMode.AUTO, spacing=0, expand=True)
        self.mc     = ft.Container(
            content=self.mc_col, expand=True, bgcolor=BG_P,
            padding=_ps(24, 20),
        )
        self.page.add(ft.Row([self.sb, self.mc], spacing=0, expand=True))

    def _toggle_sb(self, e=None):
        """Alterna el menú lateral entre contraído (solo logo UNESR) y expandido."""
        self.sb_expandido = not self.sb_expandido
        self.sb.width = 250 if self.sb_expandido else 72
        self._sb_refresh()
        try:
            if getattr(self.sb, "page", None):
                self.sb.update()
            elif hasattr(self.page, "update"):
                self.page.update()
        except Exception:
            if hasattr(self.page, "update"):
                self.page.update()

    def _nav(self, v: str):
        self.vista = v
        self._sb_refresh()
        self._mc_refresh()
        self.page.update()

    def _abrir_modal(self, dlg: ft.AlertDialog):
        """Abre un diálogo modal asegurando compatibilidad total con Flet 0.28+."""
        try:
            if hasattr(self.page, "open") and callable(self.page.open):
                self.page.open(dlg)
                return
        except Exception as ex:
            print(f"[Aviso page.open]: {ex}")

        # Fallback para Flet clásico
        try:
            if dlg not in self.page.overlay:
                self.page.overlay.append(dlg)
            dlg.open = True
            self.page.update()
        except Exception as ex:
            print(f"[Error al abrir modal]: {ex}")

    def _cerrar_modal(self, dlg: ft.AlertDialog):
        """Cierra un diálogo modal compatible con Flet 0.28+."""
        try:
            if hasattr(self.page, "close") and callable(self.page.close):
                self.page.close(dlg)
                return
        except Exception as ex:
            print(f"[Aviso page.close]: {ex}")

        # Fallback
        try:
            dlg.open = False
            self.page.update()
        except Exception as ex:
            print(f"[Error al cerrar modal]: {ex}")

    def _mostrar_notificacion(self, texto: str, color=BLUE):
        """Muestra una notificación SnackBar compatible con Flet 0.28+."""
        sb = ft.SnackBar(ft.Text(texto, color=WHITE), bgcolor=color)
        try:
            if hasattr(self.page, "open") and callable(self.page.open):
                self.page.open(sb)
                return
        except Exception:
            pass
        try:
            self.page.snack_bar = sb
            sb.open = True
            self.page.update()
        except Exception as ex:
            print(f"[Aviso SnackBar]: {ex}")

    def _on_fp_arch_result(self, e: ft.FilePickerResultEvent):
        if e.files and len(e.files) > 0:
            f = e.files[0]
            if f.path:
                self._procesar_archivo_ruta(f.path)

    # ── ACCIÓN DE CARGA DE ARCHIVOS TABULARES ──
    def _importar_archivo_dialogo(self, e=None):
        try:
            self.fp_arch.pick_files(
                dialog_title="Seleccionar archivo de planificación",
                allowed_extensions=["xlsx", "docx", "txt", "csv"]
            )
        except Exception as ex:
            print(f"[Error al abrir selector de archivos]: {ex}")
            self._mostrar_notificacion(f"No se pudo abrir el selector de archivos: {ex}", color=RED)

    def _mostrar_dialogo_archivo_manual(self, tipo: str = "planificacion"):
        tit_texto = "Importar Archivo de Planificación"
        exts = (".xlsx", ".docx", ".txt", ".csv")

        tf_ruta = campo(
            "Ruta del archivo en su computadora",
            hint_text="/home/usuario/Descargas/archivo.xlsx",
            expand=True,
        )

        def cerrar(ev=None):
            self._cerrar_modal(dlg)

        def procesar(ev=None):
            r = tf_ruta.value.strip()
            if not r:
                tf_ruta.error_text = "Ingrese la ruta del archivo"
                tf_ruta.update()
                return
            if not os.path.isfile(r):
                tf_ruta.error_text = "El archivo no existe en la ruta indicada"
                tf_ruta.update()
                return
            self._cerrar_modal(dlg)
            self._procesar_archivo_ruta(r)

        def explorar_os(ev=None):
            r_sel = abrir_selector_archivo_os(
                titulo=tit_texto,
                extensiones=list(exts)
            )
            if r_sel:
                tf_ruta.value = r_sel
                tf_ruta.error_text = None
                tf_ruta.update()

        recientes = buscar_archivos_recientes_sistema(exts)
        filas_recientes = []
        if recientes:
            filas_recientes.append(ft.Container(height=4))
            filas_recientes.append(ft.Text("Archivos detectados en Descargas / Equipo:", size=11, color=TG, weight=ft.FontWeight.W_600))
            for p_rec in recientes[:4]:
                nom = os.path.basename(p_rec)
                def aplicar_reciente(ev, p_sel=p_rec):
                    tf_ruta.value = p_sel
                    tf_ruta.error_text = None
                    tf_ruta.update()
                filas_recientes.append(
                    ft.Container(
                        content=ft.Row([
                            ft.Icon(ft.Icons.INSERT_DRIVE_FILE_ROUNDED, color=BLUE, size=15),
                            ft.Text(nom, size=11, color=TL, expand=True, max_lines=1, overflow=ft.TextOverflow.ELLIPSIS),
                            ft.TextButton("Usar este", style=ft.ButtonStyle(color=BLUE), on_click=aplicar_reciente),
                        ], spacing=6, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                        bgcolor=BG_IN, border_radius=6, padding=_ps(8, 4), margin=_mo(bottom=2),
                    )
                )

        dlg = ft.AlertDialog(
            modal=True,
            bgcolor=BG_C,
            title=ft.Row([
                ft.Icon(ft.Icons.UPLOAD_FILE_ROUNDED, color=PURP, size=22),
                ft.Text(tit_texto, size=15, weight=ft.FontWeight.BOLD, color=TW),
            ], spacing=8),
            content=ft.Container(
                content=ft.Column([
                    ft.Text("Seleccione un archivo Excel, Word, TXT o CSV:", size=12, color=TL),
                    ft.Container(height=6),
                    ft.Row([
                        tf_ruta,
                        ft.IconButton(
                            icon=ft.Icons.FOLDER_OPEN_ROUNDED,
                            icon_color=WHITE,
                            bgcolor=BLUE,
                            tooltip="Explorar en el sistema",
                            on_click=explorar_os,
                        )
                    ], spacing=6),
                    *filas_recientes,
                ], spacing=8, tight=True),
                width=520,
            ),
            actions=[
                ft.TextButton("Cancelar", on_click=cerrar),
                ft.ElevatedButton("Procesar Archivo", icon=ft.Icons.PLAY_ARROW_ROUNDED, bgcolor=BLUE, color=WHITE, on_click=procesar),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        self._abrir_modal(dlg)

    def _procesar_archivo_ruta(self, ruta: str):
        nom_arch = os.path.basename(ruta)
        self._mostrar_notificacion(f"Procesando archivo '{nom_arch}'... por favor espere.", color=BLUE)

        def worker():
            try:
                res = importar_archivo(ruta)
                insertados = res.get("procesados", 0)
                errores = res.get("errores", [])
                res_sched = generar_horarios_automaticos()
                self._all = self.motor.sincronizar_desde_sqlite()
                if DOCENTES and not self.docente_calendario_sel:
                    self.docente_calendario_sel = DOCENTES[0]
                if insertados > 0:
                    pendientes = res_sched.get("pendientes", 0)
                    self.ultimo_mensaje = (
                        f"Archivo '{nom_arch}' procesado: {insertados} cursos registrados; "
                        f"{pendientes} bloques pendientes por conflictos o falta de disponibilidad."
                    )
                    self._sb_refresh()
                    self._mc_refresh()
                    if pendientes or res_sched.get("mensajes"):
                        detalle = res_sched.get("mensajes", [""])[0]
                        self._mostrar_notificacion(
                            f"Archivo '{nom_arch}' importado: {insertados} cursos; "
                            f"{pendientes} bloques pendientes. {detalle}",
                            color=ORG,
                        )
                    else:
                        self._mostrar_notificacion(
                            f"Archivo '{nom_arch}' importado: {insertados} cursos asignados sin conflictos.",
                            color=GRN,
                        )
                else:
                    det = f" (Errores: {', '.join(errores[:2])})" if errores else ""
                    self._mostrar_notificacion(f"No se pudieron extraer cursos válidos de '{nom_arch}'{det}", color=ORG)
            except Exception as ex:
                self._mostrar_notificacion(f"Error al procesar '{nom_arch}': {ex}", color=RED)

        threading.Thread(target=worker, daemon=True).start()

    def _abrir_dialogo_ambiente(self, entry: Dict, al_actualizar=None):
        docente = entry.get("docente", "")
        materia = entry.get("materia", "")
        dia = entry.get("dia", "")
        franja = entry.get("franja", "")
        amb_actual = entry.get("ambiente", "P1-01")

        dd_amb = ddrop(
            "Seleccionar de la lista de ambientes",
            [ft.dropdown.Option(a) for a in AMBIENTES],
            value=amb_actual if amb_actual in AMBIENTES else (AMBIENTES[0] if AMBIENTES else "P1-01"),
            expand=True,
        )
        tf_custom_amb = campo(
            "O escribir código de aula personalizada",
            hint_text="Ej: P2-10, Aula 104, Lab-02, VIRTUAL-03",
        )

        def cerrar(e=None):
            self._cerrar_modal(dlg)

        def guardar(e=None):
            nuevo = tf_custom_amb.value.strip().upper() or (dd_amb.value or "").strip().upper()
            if not nuevo:
                return
            if nuevo not in AMBIENTES:
                AMBIENTES.append(nuevo)
            self.motor.cambiar_ambiente(entry, nuevo)
            for itm in self._all:
                if itm.get("docente") == docente and itm.get("materia") == materia and itm.get("dia") == dia:
                    itm["ambiente"] = nuevo
            actualizar_ambiente_por_docente_materia(docente, materia, dia, nuevo)
            self._all = self.motor.sincronizar_desde_sqlite()
            self._cerrar_modal(dlg)
            self.ultimo_mensaje = f"Ambiente de '{materia}' ({docente}) modificado a '{nuevo}'."
            self._sb_refresh()
            if al_actualizar:
                al_actualizar()
            else:
                self._mc_refresh()
            self._mostrar_notificacion(f"Ambiente actualizado a '{nuevo}' para {materia}", color=GRN)

        dlg = ft.AlertDialog(
            modal=True,
            bgcolor=BG_C,
            title=ft.Row([
                ft.Icon(ft.Icons.MEETING_ROOM_ROUNDED, color=ORG, size=24),
                ft.Text("Modificar Ambiente Asignado", size=16, weight=ft.FontWeight.BOLD, color=TW),
            ], spacing=8),
            content=ft.Container(
                content=ft.Column([
                    ft.Container(
                        content=ft.Column([
                            ft.Text(f"Docente: {docente}", size=12, weight=ft.FontWeight.BOLD, color=TW),
                            ft.Text(f"Materia: {materia}", size=12, color=BLUE, weight=ft.FontWeight.W_600),
                            ft.Text(f"Horario: {dia} | {franja}", size=11, color=TG),
                            ft.Row([
                                ft.Text("Ambiente actual:", size=11, color=TG),
                                badge(amb_actual, ORG),
                            ], spacing=6),
                        ], spacing=4),
                        bgcolor=BG_IN, border_radius=8, padding=_pa(12), border=_ba(1, BOR),
                    ),
                    ft.Container(height=8),
                    dd_amb,
                    ft.Container(height=4),
                    tf_custom_amb,
                ], spacing=6, tight=True),
                width=420,
            ),
            actions=[
                ft.TextButton("Cancelar", on_click=cerrar),
                ft.ElevatedButton(
                    "Guardar Cambio de Aula",
                    icon=ft.Icons.CHECK_ROUNDED,
                    bgcolor=BLUE,
                    color=WHITE,
                    on_click=guardar,
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )

        self._abrir_modal(dlg)

    def _abrir_dialogo_ambiente_docente_completo(self, docente: str, al_actualizar=None):
        """Permite modificar manualmente el ambiente de TODAS las materias asignadas a un docente."""
        if not docente:
            self._mostrar_notificacion("Por favor seleccione un docente primero.", color=ORG)
            return

        asig_doc = [x for x in self._all if x.get("docente") == docente]
        amb_actual = asig_doc[0].get("ambiente", "P1-01") if asig_doc else "P1-01"

        dd_amb = ddrop(
            "Seleccionar nuevo ambiente para el docente",
            [ft.dropdown.Option(a) for a in AMBIENTES],
            value=amb_actual if amb_actual in AMBIENTES else (AMBIENTES[0] if AMBIENTES else "P1-01"),
            expand=True,
        )
        tf_custom = campo("O escribir código de aula personalizada", hint_text="Ej: P2-09, Lab-01")

        def cerrar(e=None):
            self._cerrar_modal(dlg)

        def guardar(e=None):
            nuevo = tf_custom.value.strip().upper() or (dd_amb.value or "").strip().upper()
            if not nuevo:
                return
            if nuevo not in AMBIENTES:
                AMBIENTES.append(nuevo)
            actualizar_ambiente_profesor(docente, nuevo)
            for itm in self._all:
                if itm.get("docente") == docente:
                    itm["ambiente"] = nuevo
            self._all = self.motor.sincronizar_desde_sqlite()
            self._cerrar_modal(dlg)
            self.ultimo_mensaje = f"Ambiente de {docente} asignado a '{nuevo}' en todas sus materias."
            self._sb_refresh()
            if al_actualizar:
                al_actualizar()
            else:
                self._mc_refresh()
            self._mostrar_notificacion(f"Ambiente de {docente} modificado a '{nuevo}' para todas sus clases.", color=GRN)

        dlg = ft.AlertDialog(
            modal=True,
            bgcolor=BG_C,
            title=ft.Row([
                ft.Icon(ft.Icons.MEETING_ROOM_ROUNDED, color=ORG, size=24),
                ft.Text(f"Asignar / Modificar Ambiente del Docente", size=15, weight=ft.FontWeight.BOLD, color=TW),
            ], spacing=8),
            content=ft.Container(
                content=ft.Column([
                    ft.Text(f"Docente: {docente}", size=13, weight=ft.FontWeight.BOLD, color=BLUE),
                    ft.Text(f"Asignará el aula elegida a las {len(asig_doc)} materias asignadas a este profesor:", size=11, color=TG),
                    ft.Container(height=6),
                    dd_amb,
                    ft.Container(height=4),
                    tf_custom,
                ], spacing=6, tight=True),
                width=450,
            ),
            actions=[
                ft.TextButton("Cancelar", on_click=cerrar),
                ft.ElevatedButton("Guardar y Aplicar a Todas las Materias", icon=ft.Icons.CHECK, bgcolor=BLUE, color=WHITE, on_click=guardar),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        self._abrir_modal(dlg)

    def _abrir_dialogo_nuevo_ambiente(self, al_actualizar=None):
        """Permite crear un nuevo ambiente con numeración global o un código especial."""
        dd_piso = ddrop(
            "Piso / Sector",
            [
                ft.dropdown.Option("1", "Piso 1"),
                ft.dropdown.Option("2", "Piso 2"),
                ft.dropdown.Option("3", "Piso 3"),
                ft.dropdown.Option("4", "Piso 4"),
                ft.dropdown.Option("ESP", "Otro / Especial"),
            ],
            value="1",
            width=190,
        )
        tf_cod = campo(
            "Código del Ambiente / Aula",
            value=siguiente_codigo_ambiente(1),
            hint_text="Ej: P1-38, SATELITAL-03, Lab-Física",
            expand=True,
        )
        tf_cap = campo(
            "Capacidad (estudiantes)",
            value=str(MAX_EST),
            keyboard_type=ft.KeyboardType.NUMBER,
            width=190,
        )
        cb_camillas = ft.Checkbox(
            label="CAMILLAS",
            value=False,
        )
        lbl_ayuda = ft.Text(
            "La numeración continúa el total global de ambientes (P(piso)-(nro)).",
            size=11,
            color=TG,
            italic=True,
        )
        lbl_tipo = ft.Text("", size=11, color=TG, italic=True)

        def cambiar_tipo(e=None):
            es_camillas = bool(cb_camillas.value)
            tf_cap.disabled = es_camillas
            tf_cap.visible = not es_camillas
            if es_camillas:
                lbl_tipo.value = "El ambiente se mostrará como 'Camillas' y no sumará capacidad numérica."
            else:
                lbl_tipo.value = ""
            try:
                tf_cap.update()
                lbl_tipo.update()
            except Exception:
                pass

        cb_camillas.on_change = cambiar_tipo

        def cambiar_piso(e=None):
            piso = dd_piso.value or "1"
            if piso != "ESP":
                tf_cod.value = siguiente_codigo_ambiente(piso)
                tf_cod.disabled = False
                lbl_ayuda.value = f"Siguiente código sugerido para el Piso {piso}: {tf_cod.value}"
            else:
                tf_cod.value = ""
                tf_cod.disabled = False
                lbl_ayuda.value = "Escriba un nombre o código especial para el ambiente."
            try:
                tf_cod.update()
                lbl_ayuda.update()
            except Exception:
                pass

        dd_piso.on_change = cambiar_piso

        def cerrar(e=None):
            self._cerrar_modal(dlg)

        def guardar(e=None):
            cod_val = (tf_cod.value or "").strip().upper()
            if not cod_val:
                tf_cod.error_text = "El código de ambiente es obligatorio"
                tf_cod.update()
                return

            piso = dd_piso.value or "ESP"
            if piso != "ESP":
                prefijo = f"P{piso}-"
                if not cod_val.startswith(prefijo) or not cod_val[len(prefijo):].isdigit():
                    tf_cod.error_text = f"Use la nomenclatura {prefijo}(número), ej: {siguiente_codigo_ambiente(piso)}"
                    tf_cod.update()
                    return

            existentes = {str(s["codigo"]).strip().upper() for s in obtener_salones()}
            if cod_val in existentes:
                tf_cod.error_text = "Ese ambiente ya existe en el sistema"
                tf_cod.update()
                return

            es_camillas = bool(cb_camillas.value)
            if es_camillas:
                capacidad = MAX_EST
            else:
                try:
                    capacidad = int(tf_cap.value or MAX_EST)
                except (TypeError, ValueError):
                    capacidad = MAX_EST
                if capacidad <= 0:
                    tf_cap.error_text = "La capacidad debe ser mayor que cero"
                    tf_cap.update()
                    return

            edificio = f"Piso {piso}" if piso != "ESP" else "Ambiente Especial"
            crear_ambiente(cod_val, capacidad, edificio, es_camillas=es_camillas)
            recargar_ambientes_desde_bd()

            self._cerrar_modal(dlg)
            if es_camillas:
                self.ultimo_mensaje = f"Nuevo ambiente de camillas '{cod_val}' creado exitosamente."
                detalle = "como ambiente de camillas"
            else:
                self.ultimo_mensaje = f"Nuevo ambiente '{cod_val}' creado con capacidad para {capacidad} estudiantes."
                detalle = f"con capacidad {capacidad}"
            self._sb_refresh()
            if al_actualizar:
                al_actualizar()
            else:
                self._refrescar_vista_actual()
            self._mostrar_notificacion(f"Ambiente '{cod_val}' registrado {detalle}.", color=GRN)

        dlg = ft.AlertDialog(
            modal=True,
            bgcolor=BG_C,
            title=ft.Row([
                ft.Icon(ft.Icons.ADD_LOCATION_ALT_ROUNDED, color=BLUE, size=24),
                ft.Text("Registrar Nuevo Ambiente / Aula", size=16, weight=ft.FontWeight.BOLD, color=TW),
            ], spacing=8),
            content=ft.Container(
                content=ft.Column([
                    ft.Text("Seleccione el piso y registre la capacidad del nuevo espacio académico:", size=12, color=TL),
                    ft.Container(height=6),
                    ft.Row([dd_piso, tf_cap], spacing=10),
                    cb_camillas,
                    lbl_tipo,
                    tf_cod,
                    lbl_ayuda,
                ], spacing=8, tight=True),
                width=480,
            ),
            actions=[
                ft.TextButton("Cancelar", on_click=cerrar),
                ft.ElevatedButton("Crear Ambiente", icon=ft.Icons.CHECK, bgcolor=BLUE, color=WHITE, on_click=guardar),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        self._abrir_modal(dlg)

    def _abrir_dialogo_gestionar_aula(self, amb: str, clases: List[Dict], al_actualizar=None):
        """Muestra las clases asignadas al aula para modificar o permite asignar un docente/clase si está libre."""
        def cerrar(e=None):
            self._cerrar_modal(dlg)

        items = []
        datos_amb = obtener_salon(amb) or {}
        cap_actual = int(datos_amb.get("capacidad") or obtener_capacidad_salon(amb, MAX_EST))
        es_camillas_actual = bool(int(datos_amb.get("es_camillas") or 0))
        tf_cap = campo(
            "Capacidad del ambiente (estudiantes)",
            value=str(cap_actual),
            keyboard_type=ft.KeyboardType.NUMBER,
            width=210,
            disabled=es_camillas_actual,
            visible=not es_camillas_actual,
        )
        cb_camillas = ft.Checkbox(
            label="CAMILLAS",
            value=es_camillas_actual,
        )
        lbl_tipo_camillas = ft.Text(
            "Este ambiente se muestra como 'Camillas' y no suma capacidad numérica."
            if es_camillas_actual else "",
            size=11, color=TG, italic=True,
        )

        def cambiar_tipo_camillas(e=None):
            es_camillas = bool(cb_camillas.value)
            tf_cap.disabled = es_camillas
            tf_cap.visible = not es_camillas
            lbl_tipo_camillas.value = (
                "Este ambiente se muestra como 'Camillas' y no suma capacidad numérica."
                if es_camillas else ""
            )
            try:
                tf_cap.update()
                lbl_tipo_camillas.update()
            except Exception:
                pass

        cb_camillas.on_change = cambiar_tipo_camillas

        def guardar_datos_ambiente(e=None):
            es_camillas = bool(cb_camillas.value)
            if es_camillas:
                nueva_cap = cap_actual if cap_actual > 0 else MAX_EST
            else:
                try:
                    nueva_cap = int((tf_cap.value or "").strip())
                except (TypeError, ValueError):
                    tf_cap.error_text = "Ingrese un número válido"
                    tf_cap.update()
                    return
                if nueva_cap <= 0:
                    tf_cap.error_text = "La capacidad debe ser mayor que cero"
                    tf_cap.update()
                    return

            actualizar_datos_ambiente(amb, capacidad=nueva_cap, es_camillas=es_camillas)
            self._cerrar_modal(dlg)
            if es_camillas:
                self.ultimo_mensaje = f"El ambiente {amb} quedó registrado como ambiente de camillas."
            else:
                self.ultimo_mensaje = f"Capacidad del ambiente {amb} actualizada a {nueva_cap} estudiantes."
            self._sb_refresh()
            if al_actualizar:
                al_actualizar()
            else:
                self._refrescar_vista_actual()
            self._mostrar_notificacion(
                f"{amb}: {'Camillas' if es_camillas else f'{nueva_cap} estudiantes'}.",
                color=GRN,
            )

        def abrir_confirmacion_eliminar(e=None):
            self._cerrar_modal(dlg)
            self._dialogo_eliminar_ambiente(amb, clases, al_actualizar=al_actualizar)

        items.append(ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.Icon(
                        ft.Icons.AIRLINE_SEAT_FLAT_ROUNDED if es_camillas_actual else ft.Icons.GROUPS_ROUNDED,
                        color=PURP, size=18,
                    ),
                    ft.Text(
                        "Tipo actual: Camillas" if es_camillas_actual else f"Capacidad actual: {cap_actual} estudiantes",
                        size=12, weight=ft.FontWeight.W_600, color=TW,
                    ),
                ], spacing=8),
                cb_camillas,
                lbl_tipo_camillas,
                ft.Row([
                    tf_cap,
                    ft.ElevatedButton(
                        "Guardar Datos",
                        icon=ft.Icons.SAVE_ROUNDED,
                        bgcolor=PURP,
                        color=WHITE,
                        on_click=guardar_datos_ambiente,
                    ),
                    ft.ElevatedButton(
                        "Eliminar Ambiente",
                        icon=ft.Icons.DELETE_FOREVER_ROUNDED,
                        bgcolor=RED,
                        color=WHITE,
                        on_click=abrir_confirmacion_eliminar,
                    ),
                ], spacing=10, vertical_alignment=ft.CrossAxisAlignment.CENTER, wrap=True),
            ], spacing=8),
            bgcolor=alpha_col(PURP, 0.08),
            border=_ba(1, alpha_col(PURP, 0.35)),
            border_radius=10,
            padding=_pa(12),
        ))
        items.append(ft.Container(height=4))

        if clases:
            items.append(ft.Text(f"Clases asignadas actualmente al aula {amb}:", size=13, weight=ft.FontWeight.W_600, color=TW))
            items.append(ft.Container(height=4))
            for cl in clases:
                btn_reubicar = ft.ElevatedButton(
                    "Reasignar Aula",
                    icon=ft.Icons.EDIT_ROUNDED,
                    bgcolor=ORG,
                    color=WHITE,
                    on_click=lambda _, itm=cl: [self._cerrar_modal(dlg), self._abrir_dialogo_ambiente(itm, al_actualizar=al_actualizar)],
                )
                items.append(ft.Container(
                    content=ft.Row([
                        ft.Column([
                            ft.Text(f"{cl.get('materia', '')}", color=TW, size=12, weight=ft.FontWeight.BOLD),
                            ft.Text(f"Prof: {cl.get('docente', '')} | {cl.get('dia', '')} ({cl.get('franja', '')})", color=TG, size=11),
                        ], spacing=2, expand=True),
                        btn_reubicar,
                    ], spacing=10, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                    bgcolor=BG_IN, border_radius=8, padding=_pa(10), border=_ba(1, BOR),
                ))
        else:
            items.append(ft.Text(f"El aula {amb} se encuentra totalmente disponible.", size=13, color=GRN, weight=ft.FontWeight.W_600))
            items.append(ft.Text("Seleccione un docente para asignarle este ambiente en todas sus clases:", size=11, color=TG))
            docs_con_clases = sorted(list({x.get("docente") for x in self._all if x.get("docente")}))
            dd_doc = ddrop("Seleccionar Docente a Asignar", [ft.dropdown.Option(d) for d in docs_con_clases], expand=True)
            items.append(dd_doc)

            def asignar_doc(e=None):
                doc_sel = dd_doc.value
                if not doc_sel:
                    return
                actualizar_ambiente_profesor(doc_sel, amb)
                for itm in self._all:
                    if itm.get("docente") == doc_sel:
                        itm["ambiente"] = amb
                self._all = self.motor.sincronizar_desde_sqlite()
                self._cerrar_modal(dlg)
                self.ultimo_mensaje = f"Docente {doc_sel} asignado exitosamente al ambiente {amb}."
                self._sb_refresh()
                if al_actualizar:
                    al_actualizar()
                else:
                    self._mc_refresh()
                self._mostrar_notificacion(f"{doc_sel} asignado al aula {amb}.", color=GRN)

            items.append(ft.Container(height=4))
            items.append(ft.ElevatedButton("Asignar Docente al Aula", icon=ft.Icons.CHECK, bgcolor=BLUE, color=WHITE, on_click=asignar_doc))

        dlg = ft.AlertDialog(
            modal=True,
            bgcolor=BG_C,
            title=ft.Row([
                ft.Icon(ft.Icons.MEETING_ROOM_ROUNDED, color=ORG, size=24),
                ft.Text(f"Gestión de Ambiente: {amb}", size=16, weight=ft.FontWeight.BOLD, color=TW),
            ], spacing=8),
            content=ft.Container(
                content=ft.Column(items, spacing=8, tight=True),
                width=480,
            ),
            actions=[
                ft.TextButton("Cerrar", on_click=cerrar),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        self._abrir_modal(dlg)

    def _dialogo_eliminar_ambiente(self, amb: str, clases: List[Dict], al_actualizar=None):
        """Confirma y elimina un ambiente individualmente."""
        def cerrar(e=None):
            self._cerrar_modal(dlg)

        def confirmar(e=None):
            self._cerrar_modal(dlg)
            if not eliminar_salon(amb):
                self._mostrar_notificacion(f"No se pudo eliminar el ambiente {amb}.", color=RED)
                return
            recargar_ambientes_desde_bd()
            self.motor.limpiar_memoria()
            self._all = self.motor.sincronizar_desde_sqlite()
            recargar_docentes_desde_bd()
            self.ultimo_mensaje = f"Ambiente '{amb}' eliminado del sistema."
            self._sb_refresh()
            if al_actualizar:
                al_actualizar()
            else:
                self._refrescar_vista_actual()
            self._mostrar_notificacion(f"Ambiente '{amb}' eliminado.", color=BLUE)

        detalle_clases = (
            f"\n\nEste ambiente tiene {len(clases)} clase(s) asignada(s). "
            "Si continúa, esas asignaciones también serán eliminadas y los docentes quedarán sin ese espacio."
            if clases else ""
        )

        dlg = ft.AlertDialog(
            modal=True,
            bgcolor=BG_C,
            title=ft.Row([
                ft.Icon(ft.Icons.WARNING_ROUNDED, color=RED, size=24),
                ft.Text(f"¿Eliminar el ambiente {amb}?", size=16, weight=ft.FontWeight.BOLD, color=TW),
            ], spacing=8),
            content=ft.Text(
                "Esta acción eliminará el ambiente de forma individual y no se puede deshacer."
                + detalle_clases,
                size=13, color=TL,
            ),
            actions=[
                ft.TextButton("Cancelar", on_click=cerrar),
                ft.ElevatedButton(
                    "Sí, Eliminar Ambiente",
                    icon=ft.Icons.DELETE_FOREVER_ROUNDED,
                    bgcolor=RED,
                    color=WHITE,
                    on_click=confirmar,
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        self._abrir_modal(dlg)

    def _dialogo_vaciar_sistema_en_blanco(self, e=None):
        def cerrar(ev=None):
            self._cerrar_modal(dlg)

        def confirmar_vaciado(ev=None):
            self._cerrar_modal(dlg)
            marcar_iniciar_en_blanco(True)
            clear_all_data(keep_salones=True)
            self._all = []
            self.motor.limpiar_memoria()
            DOCENTES.clear()
            CARRERAS.clear()
            CARRERAS.update({k: dict(v) for k, v in CARRERAS_BASE.items()})
            MATERIAS.clear()
            MATERIAS.update({k: [dict(m) for m in v] for k, v in MATERIAS_BASE.items()})
            recargar_ambientes_desde_bd()
            self.docente_calendario_sel = ""
            self.ultimo_mensaje = "Datos cargados eliminados. El sistema quedó en blanco y listo para nuevos registros."
            self._sb_refresh()
            self._mc_refresh()
            self._mostrar_notificacion("Datos eliminados con éxito. El sistema quedó en blanco.", color=BLUE)

        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Row([
                ft.Icon(ft.Icons.WARNING_ROUNDED, color=RED, size=24),
                ft.Text("¿Eliminar todos los datos cargados?", size=16, weight=ft.FontWeight.BOLD, color=TW),
            ], spacing=8),
            content=ft.Text(
                "Esta acción eliminará todos los cursos, profesores, carreras cargadas y asignaciones actuales. "
                "No se puede deshacer.\n\n"
                "Los ambientes (salones), sus capacidades y la configuración del sistema se mantendrán intactos.",
                size=13, color=TL
            ),
            actions=[
                ft.TextButton("Cancelar", on_click=cerrar),
                ft.ElevatedButton(
                    "Sí, Eliminar Datos",
                    icon=ft.Icons.DELETE_SWEEP_ROUNDED,
                    bgcolor=RED,
                    color=WHITE,
                    on_click=confirmar_vaciado,
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        self._abrir_modal(dlg)

    def _cargar_datos_demo(self, e=None):
        marcar_iniciar_en_blanco(False)
        self.motor, self._all = _seed_data(forzar_base=True)
        if DOCENTES:
            self.docente_calendario_sel = DOCENTES[0]
        self.ultimo_mensaje = "Datos base de la Universidad Simon Rodriguez (33 cursos) cargados correctamente."
        self._sb_refresh()
        self._mc_refresh()
        self.page.snack_bar = ft.SnackBar(
            ft.Text("Datos de demostración cargados exitosamente.", color=WHITE),
            bgcolor=BLUE
        )
        self.page.snack_bar.open = True
        self.page.update()

    def _toggle_theme(self, e=None):
        if self.page.theme_mode == ft.ThemeMode.LIGHT:
            self.page.theme_mode = ft.ThemeMode.DARK
            _set_dark_palette()
        else:
            self.page.theme_mode = ft.ThemeMode.LIGHT
            _set_light_palette()
        self.page.bgcolor = BG_P
        self.sb.bgcolor = BG_SB
        self.sb.border = _br(1, BOR)
        self.mc.bgcolor = BG_P
        self._sb_refresh()
        self._mc_refresh()
        self.page.update()

    def _restaurar_prototipo_base(self, e=None):
        self.motor, self._all = _seed_data(forzar_base=True)
        if DOCENTES:
            self.docente_calendario_sel = DOCENTES[0]
        self.ultimo_mensaje = "Datos base de la Universidad Simon Rodriguez (33 cursos) cargados correctamente."
        self._sb_refresh()
        self._mc_refresh()
        self.page.update()

    def _sb_refresh(self):
        st = self.motor.stats()
        es_claro = getattr(self.page, "theme_mode", ft.ThemeMode.LIGHT) == ft.ThemeMode.LIGHT

        # ── ESTADO CONTRAÍDO (SOLO MUESTRA EL LOGO UNESR) ──
        if not self.sb_expandido:
            logo_contraido = ft.Container(
                content=ft.Column([
                    ft.Container(
                        content=obtener_logo_unesr(tamano=42, tooltip="Universidad Experimental Simón Rodríguez (UNESR)"),
                        bgcolor=WHITE,
                        border=_ba(1, alpha_col(BLUE, 0.4)),
                        border_radius=10,
                        padding=4,
                        width=52,
                        height=52,
                        alignment=ft.alignment.center,
                        shadow=ft.BoxShadow(spread_radius=1, blur_radius=3, color=alpha_col(BLUE, 0.15)),
                    ),
                    ft.Container(height=4),
                    ft.Icon(ft.Icons.KEYBOARD_DOUBLE_ARROW_RIGHT_ROUNDED, color=BLUE, size=18),
                ], spacing=2, alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                padding=_ps(6, 16),
                alignment=ft.alignment.center,
                tooltip="Clic en el logo para expandir menú (UNESR)",
                on_click=self._toggle_sb,
                ink=True,
            )

            btn_tema_col = ft.IconButton(
                icon=ft.Icons.DARK_MODE_ROUNDED if es_claro else ft.Icons.LIGHT_MODE_ROUNDED,
                icon_color=TL,
                icon_size=18,
                tooltip="Modo Oscuro" if es_claro else "Modo Claro",
                on_click=self._toggle_theme,
            )
            btn_eliminar_col = ft.IconButton(
                icon=ft.Icons.DELETE_SWEEP_ROUNDED,
                icon_color=RED,
                icon_size=18,
                tooltip="Eliminar datos cargados",
                on_click=self._dialogo_vaciar_sistema_en_blanco,
            )

            self.sb_col.controls = [
                logo_contraido,
                ft.Container(expand=True),
                ft.Container(
                    content=ft.Column([
                        btn_eliminar_col,
                        btn_tema_col,
                    ], spacing=4, alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                    padding=_ps(6, 12),
                    alignment=ft.alignment.center,
                ),
            ]
            return

        # ── ESTADO EXPANDIDO (LOGO + NOMBRE INSTITUCIONAL + TODAS LAS CATEGORÍAS) ──
        logo_expandido = ft.Container(
            content=obtener_logo_unesr(tamano=38, tooltip="UNESR"),
            bgcolor=WHITE,
            border=_ba(1, alpha_col(BLUE, 0.3)),
            border_radius=10,
            padding=3,
            width=46,
            height=46,
            alignment=ft.alignment.center,
            shadow=ft.BoxShadow(spread_radius=1, blur_radius=2, color=alpha_col(BLUE, 0.12)),
            on_click=self._toggle_sb,
            tooltip="Clic para contraer menú lateral",
            ink=True,
        )

        btn_contraer = ft.IconButton(
            icon=ft.Icons.KEYBOARD_DOUBLE_ARROW_LEFT_ROUNDED,
            icon_color=TG,
            icon_size=18,
            tooltip="Contraer menú lateral",
            on_click=self._toggle_sb,
        )

        header = ft.Container(
            content=ft.Column([
                ft.Row([
                    logo_expandido,
                    ft.Column([
                        ft.Text("UNESR", color=TW, size=15, weight=ft.FontWeight.BOLD),
                        ft.Text("Univ. Experimental", color=TG, size=10, weight=ft.FontWeight.W_600),
                        ft.Text("Simón Rodríguez", color=TG, size=10, weight=ft.FontWeight.W_600),
                        ft.Text("Núcleo Palo Verde", color=BLUE, size=9, weight=ft.FontWeight.W_500),
                    ], spacing=0, expand=True),
                    btn_contraer,
                ], spacing=8, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                ft.Container(height=4),
                ft.Container(
                    content=ft.Row([
                        ft.Icon(ft.Icons.CIRCLE, color=GRN, size=8),
                        ft.Text("Periodo 2026-2", color=GRN, size=11),
                    ], spacing=6),
                    bgcolor=alpha_col(GRN, 0.15), border_radius=20,
                    padding=_ps(10, 4), border=_ba(1, alpha_col(GRN, 0.4)),
                ),
            ], spacing=6),
            padding=_ps(14, 14),
        )

        btns = []
        for vid, ico, lbl in self.NAV:
            act = self.vista == vid
            btns.append(ft.Container(
                content=ft.Row([
                    ft.Icon(ico, color=BLUE if act else TG, size=18),
                    ft.Text(lbl, color=TW if act else TG, size=13,
                            weight=ft.FontWeight.W_600 if act else ft.FontWeight.W_400),
                ], spacing=12),
                bgcolor=alpha_col(BLUE, 0.18) if act else None,
                border=_ba(1, alpha_col(BLUE, 0.5)) if act else None,
                border_radius=10,
                padding=_ps(14, 10),
                margin=_ms(10, 0),
                on_click=lambda e, v=vid: self._nav(v),
                ink=True,
            ))

        btn_tema_sb = ft.Container(
            content=ft.Row([
                ft.Icon(ft.Icons.DARK_MODE_ROUNDED if es_claro else ft.Icons.LIGHT_MODE_ROUNDED, color=TL, size=14),
                ft.Text("Modo Oscuro" if es_claro else "Modo Claro", color=TL, size=11, weight=ft.FontWeight.W_600),
            ], spacing=6, alignment=ft.MainAxisAlignment.CENTER),
            bgcolor=BG_IN,
            border=_ba(1, BOR),
            border_radius=8,
            padding=_ps(8, 5),
            margin=_ms(10, 3),
            on_click=self._toggle_theme,
            ink=True,
        )

        btn_eliminar_sb = ft.Container(
            content=ft.Row([
                ft.Icon(ft.Icons.DELETE_SWEEP_ROUNDED, color=RED, size=14),
                ft.Text("Eliminar datos cargados", color=RED, size=11, weight=ft.FontWeight.W_600),
            ], spacing=6, alignment=ft.MainAxisAlignment.CENTER),
            bgcolor=alpha_col(RED, 0.1),
            border=_ba(1, alpha_col(RED, 0.35)),
            border_radius=8,
            padding=_ps(8, 5),
            margin=_ms(10, 3),
            on_click=self._dialogo_vaciar_sistema_en_blanco,
            ink=True,
            tooltip="Elimina cursos, docentes y asignaciones cargadas",
        )

        def mrow(lbl, val, col):
            return ft.Row([
                ft.Text(lbl, color=TG, size=11, expand=True),
                ft.Container(
                    content=ft.Text(val, color=col, size=11, weight=ft.FontWeight.W_700),
                    bgcolor=alpha_col(col, 0.18), border_radius=6,
                    padding=_ps(7, 2),
                ),
            ])

        bot = ft.Container(content=ft.Column([
            divl(),
            ft.Container(content=ft.Column([
                ft.Text("RESUMEN RAPIDO", color=TG, size=10, weight=ft.FontWeight.W_700),
                ft.Container(height=4),
                mrow("Cursos activos",   str(st["cursos"]),  BLUE),
                mrow("Ambientes en uso", f"{st['amb_uso']}/{st['amb_tot']}", ORG),
                mrow("Conflictos",       "0",                 GRN),
                mrow("Docentes",         str(st["docs"]),    CYAN),
                ft.Container(height=4),
                btn_eliminar_sb,
                btn_tema_sb,
            ], spacing=5), padding=_pa(12)),
        ]))

        self.sb_col.controls = [
            header, divl(), ft.Container(height=6),
            ft.Column(btns, spacing=3),
            ft.Container(expand=True), bot,
        ]

    def _mc_refresh(self):
        fn = {
            "dashboard":   self._vd,
            "nueva_carga": self._vnc,
            "horarios":    self._vh,
            "ambientes":   self._va,
            "docentes":    self._vdoc,
            "carreras":    self._vc,
        }.get(self.vista, self._vd)
        self.mc_col.controls = [fn()]

    def _refrescar_vista_actual(self):
        """Recarga barra lateral y vista actual, forzando el repintado de Flet."""
        self._sb_refresh()
        self._mc_refresh()
        try:
            self.page.update()
        except Exception:
            pass

    # ══════════════════════════════════════════════════
    #  VISTA DASHBOARD
    # ══════════════════════════════════════════════════
    def _vd(self) -> ft.Control:
        st = self.motor.stats()

        cards = ft.Row([
            stat_card("Total Cursos",     str(st["cursos"]),  ft.Icons.BOOK_ROUNDED,          BLUE, "Periodo 2026-2"),
            stat_card("Ambientes en Uso", str(st["amb_uso"]), ft.Icons.MEETING_ROOM_ROUNDED,  ORG,  f"de {st['amb_tot']} totales"),
            stat_card("Docentes Activos", str(st["docs"]),    ft.Icons.PEOPLE_ROUNDED,        PURP, "Periodo 2026-2"),
            stat_card("Conflictos",       "0",                ft.Icons.VERIFIED_ROUNDED,      GRN,  "0 colisiones!"),
        ], spacing=14)

        enc = ft.Container(
            content=ft.Row([
                ft.Container(width=4),
                ft.Text("CARRERA",  color=TG, size=10, weight=ft.FontWeight.W_700, expand=2),
                ft.Text("DOCENTE",  color=TG, size=10, weight=ft.FontWeight.W_700, expand=3),
                ft.Text("MATERIA",  color=TG, size=10, weight=ft.FontWeight.W_700, expand=3),
                ft.Text("DIA",      color=TG, size=10, weight=ft.FontWeight.W_700, width=85),
                ft.Text("HORARIO",  color=TG, size=10, weight=ft.FontWeight.W_700, width=120),
                ft.Text("AMBIENTE", color=TG, size=10, weight=ft.FontWeight.W_700, width=95),
                ft.Text("OK",       color=TG, size=10, weight=ft.FontWeight.W_700, width=35),
            ], spacing=8),
            bgcolor=BG_IN, border_radius=8,
            padding=_ps(12, 10), margin=_mo(bottom=4),
        )

        filas = []
        for i, e in enumerate(reversed(self._all[-10:])):
            ci  = CARRERAS.get(e["carrera_id"], {})
            col = ci.get("color", BLUE)
            bg  = BG_C if i % 2 == 0 else BG_RA
            btn_amb = ft.Container(
                content=ft.Row([
                    ft.Icon(ft.Icons.EDIT_ROUNDED, color=BLUE, size=11),
                    badge(e["ambiente"], BLUE),
                ], spacing=2, alignment=ft.MainAxisAlignment.CENTER),
                tooltip="Modificar ambiente",
                on_click=lambda _, itm=e: self._abrir_dialogo_ambiente(itm, al_actualizar=self._mc_refresh),
                ink=True,
            )
            filas.append(ft.Container(
                content=ft.Row([
                    ft.Container(width=4, height=32, bgcolor=col, border_radius=2),
                    ft.Text(ci.get("nombre","")[:16], color=col, size=11, weight=ft.FontWeight.W_600, expand=2),
                    ft.Text(e["docente"][:26],         color=TL, size=11, expand=3),
                    ft.Text(e["materia"][:26],          color=TL, size=11, expand=3),
                    ft.Container(ft.Text(e["dia"],                   color=TG, size=11), width=85),
                    ft.Container(ft.Text(e["franja"],                color=TG, size=11), width=120),
                    ft.Container(btn_amb, width=95),
                    ft.Container(ft.Icon(ft.Icons.CHECK_CIRCLE_ROUNDED, color=GRN, size=15), width=35),
                ], spacing=8, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                bgcolor=bg, border_radius=8,
                padding=_ps(12, 7), margin=_mo(bottom=2),
            ))

        if not self._all:
            filas.append(ft.Container(
                content=ft.Column([
                    ft.Icon(ft.Icons.INBOX_ROUNDED, color=TG, size=38),
                    ft.Text("No hay cursos registrados en el sistema.", color=TW, size=14, weight=ft.FontWeight.W_600),
                    ft.Text("El sistema está en blanco. Utilice 'Nueva Carga' o importe una planificación Excel, Word, TXT o CSV para comenzar.", color=TG, size=12),
                ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=6),
                padding=_pa(24), alignment=ft.alignment.center,
            ))

        tabla = ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.Text("Carga Academica Reciente", color=TW, size=15, weight=ft.FontWeight.W_600),
                    ft.Container(expand=True),
                    ft.TextButton("Ver todo en Horarios ->", style=ft.ButtonStyle(color=BLUE), on_click=lambda _: self._nav("horarios")),
                ]),
                ft.Container(height=8),
                enc,
                ft.Column(filas, spacing=0),
            ], spacing=0),
            bgcolor=BG_C, border_radius=14, padding=_pa(20), border=_ba(1, BOR),
        )

        cards_dist = []
        for k, c in CARRERAS.items():
            cant = len([e for e in self._all if e.get("carrera_id") == k])
            cards_dist.append(ft.Container(
                content=ft.Row([
                    ft.Container(width=4, height=38, bgcolor=c["color"], border_radius=2),
                    ft.Column([
                        ft.Text(c["nombre"], color=TW, size=12, weight=ft.FontWeight.W_600, max_lines=1, overflow=ft.TextOverflow.ELLIPSIS),
                        ft.Text(c["mencion"], color=TG, size=10, max_lines=1, overflow=ft.TextOverflow.ELLIPSIS),
                    ], spacing=2, expand=True),
                    badge(f"{cant} cursos", c["color"]),
                ], spacing=10, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                bgcolor=BG_IN, border_radius=10, padding=_ps(14, 10),
                border=_ba(1, alpha_col(c['color'], 0.35)), expand=True,
                on_click=lambda _, kk=k: (setattr(self, "horarios_carrera", kk), self._nav("horarios")),
                ink=True,
            ))

        filas_dist = []
        for i in range(0, len(cards_dist), 3):
            grupo = cards_dist[i:i+3]
            while len(grupo) < 3:
                grupo.append(ft.Container(expand=True))
            filas_dist.append(ft.Row(grupo, spacing=12))
            filas_dist.append(ft.Container(height=10))

        mini_car = ft.Column(filas_dist, spacing=0)

        btn_cargar_arch = ft.Container(
            content=ft.Row([
                ft.Icon(ft.Icons.UPLOAD_FILE_ROUNDED, color=WHITE, size=16),
                ft.Text("Cargar Excel/Word/TXT", color=WHITE, size=12, weight=ft.FontWeight.W_600)
            ], alignment=ft.MainAxisAlignment.CENTER, spacing=6),
            bgcolor=PURP, border_radius=10, padding=_ps(14, 11),
            on_click=self._importar_archivo_dialogo, ink=True,
        )

        btn_imprimir_rep = ft.Container(
            content=ft.Row([
                ft.Icon(ft.Icons.PRINT_ROUNDED, color=WHITE, size=16),
                ft.Text("Imprimir PDF/HTML", color=WHITE, size=12, weight=ft.FontWeight.W_700)
            ], alignment=ft.MainAxisAlignment.CENTER, spacing=6),
            bgcolor=GRN, border_radius=10, padding=_ps(14, 11),
            on_click=lambda _: generar_reporte_html_imprimible(self._all, tipo="completo"), ink=True,
        )

        btn_eliminar_datos = ft.Container(
            content=ft.Row([
                ft.Icon(ft.Icons.DELETE_SWEEP_ROUNDED, color=WHITE, size=16),
                ft.Text("Eliminar Datos Cargados", color=WHITE, size=12, weight=ft.FontWeight.W_700)
            ], alignment=ft.MainAxisAlignment.CENTER, spacing=6),
            bgcolor=RED, border_radius=10, padding=_ps(14, 11),
            on_click=self._dialogo_vaciar_sistema_en_blanco, ink=True,
            tooltip="Elimina los cursos, docentes y asignaciones cargadas (pide confirmación)",
        )

        controles_sup = [
            ft.Row([
                titulo("Dashboard", "Sistema de Gestion de Oferta Academica - UNESR (Palo Verde)"),
                ft.Container(expand=True),
                btn_eliminar_datos,
                btn_imprimir_rep,
            ], spacing=14, vertical_alignment=ft.CrossAxisAlignment.CENTER),
        ]

        if self.ultimo_mensaje:
            controles_sup.append(ft.Container(height=12))
            controles_sup.append(ft.Container(
                content=ft.Row([
                    ft.Icon(ft.Icons.CHECK_CIRCLE_ROUNDED, color=GRN, size=18),
                    ft.Text(self.ultimo_mensaje, color=TW, size=12, expand=True),
                ], spacing=10),
                bgcolor=alpha_col(GRN, 0.18), border_radius=10, padding=_ps(14, 10), border=_ba(1, alpha_col(GRN, 0.5)),
            ))

        return ft.Column([
            *controles_sup,
            ft.Container(height=20),
            cards,
            ft.Container(height=20),
            tabla,
            ft.Container(height=20),
            ft.Text("Distribucion por Carrera", color=TW, size=15, weight=ft.FontWeight.W_600),
            ft.Container(height=10),
            mini_car,
            ft.Container(height=10),
        ], spacing=0)

    # ══════════════════════════════════════════════════
    #  VISTA NUEVA CARGA (MANUAL + ARCHIVOS .XLSX / .DOCX / .TXT)
    # ══════════════════════════════════════════════════
    def _vnc_ficha_docente(self) -> ft.Control:
        franjas = [
            ("07:00", "07:45"), ("07:45", "08:30"), ("08:30", "09:15"),
            ("09:20", "10:05"), ("10:05", "10:50"), ("10:50", "11:35"),
            ("11:35", "12:15"), ("12:30", "13:15"), ("13:15", "13:55"),
            ("14:10", "14:55"), ("14:55", "15:35"), ("15:40", "16:25"),
            ("16:25", "17:05"), ("17:10", "17:45"), ("17:45", "18:25"),
        ]
        tf_docente = campo("Nombre y Apellidos del Docente", expand=True)
        tf_cedula = campo("Cédula", width=190)
        tf_correo = campo("Correo electrónico", expand=True)
        dd_condicion = ddrop(
            "Condición Docente",
            [ft.dropdown.Option(v) for v in ("Contratado/a", "Ordinario/a", "Otro")],
            value="Contratado/a",
            width=220,
        )
        dd_carrera = ddrop(
            "Carrera Universitaria",
            [ft.dropdown.Option(c["nombre"]) for c in CARRERAS.values()],
            expand=True,
        )
        dd_ambiente = ddrop(
            "Ambiente (opcional)",
            [ft.dropdown.Option(a) for a in AMBIENTES],
            value=None,
            expand=True,
        )
        tf_materias = ft.TextField(
            label="Materias que imparte (una por línea)",
            multiline=True,
            min_lines=3,
            max_lines=6,
            border_radius=10,
            expand=True,
        )
        celdas = {}
        encabezado_horas = ft.Row(
            [
                ft.Container(width=94),
                *[
                    ft.Container(
                        content=ft.Text(f"{inicio}\n{fin}", size=9, color=TG, text_align=ft.TextAlign.CENTER),
                        width=68,
                        alignment=ft.alignment.center,
                    )
                    for inicio, fin in franjas
                ],
            ],
            spacing=2,
            scroll=ft.ScrollMode.AUTO,
        )
        filas_horario = [encabezado_horas]
        for dia in DIAS:
            checks = []
            for indice, _ in enumerate(franjas):
                check = ft.Checkbox(value=False, scale=0.75)
                celdas[(dia, indice)] = check
                checks.append(ft.Container(content=check, width=68, alignment=ft.alignment.center))
            filas_horario.append(
                ft.Row(
                    [
                        ft.Container(
                            content=ft.Text(dia, size=11, color=TW, weight=ft.FontWeight.W_600),
                            width=94,
                        ),
                        *checks,
                    ],
                    spacing=2,
                    scroll=ft.ScrollMode.AUTO,
                )
            )
        estado = ft.Container(
            content=ft.Text("Complete los datos del docente y seleccione su disponibilidad.", color=TG, size=12),
            bgcolor=alpha_col(BLUE, 0.1),
            border=_ba(1, BOR),
            border_radius=10,
            padding=_pa(12),
        )

        def guardar(e=None):
            materias = [m.strip() for m in (tf_materias.value or "").splitlines() if m.strip()]
            disponibilidades = [
                (DIAS_ORDEN[DIAS.index(dia)], franjas[indice][0], franjas[indice][1])
                for (dia, indice), check in celdas.items()
                if check.value
            ]
            try:
                resultado = registrar_ficha_docente({
                    "docente": tf_docente.value,
                    "cedula": tf_cedula.value,
                    "correo": tf_correo.value,
                    "condicion": dd_condicion.value,
                    "carrera": dd_carrera.value,
                    "ambiente": dd_ambiente.value,
                    "materias": materias,
                    "disponibilidades": disponibilidades,
                })
            except Exception as ex:
                estado.content = ft.Text(f"No se pudo guardar la ficha: {ex}", color=RED, size=12)
                estado.bgcolor = alpha_col(RED, 0.14)
                estado.border = _ba(1, alpha_col(RED, 0.5))
                estado.update()
                return
            if resultado.get("errores"):
                mensaje = " ".join(resultado["errores"])
                estado.content = ft.Text(mensaje, color=RED, size=12)
                estado.bgcolor = alpha_col(RED, 0.14)
                estado.border = _ba(1, alpha_col(RED, 0.5))
                estado.update()
                return

            try:
                generacion = generar_horarios_automaticos()
                self._all = self.motor.sincronizar_desde_sqlite()
            except Exception as ex:
                estado.content = ft.Text(
                    f"La ficha se guardó, pero no se pudieron generar los horarios: {ex}",
                    color=RED,
                    size=12,
                )
                estado.bgcolor = alpha_col(RED, 0.14)
                estado.border = _ba(1, alpha_col(RED, 0.5))
                estado.update()
                return
            asignados = resultado.get("insertados", 0)
            pendientes = generacion.get("pendientes", 0)
            estado.content = ft.Column([
                ft.Text(
                    f"Ficha de {resultado['docente']} guardada. {asignados} materia(s) registradas.",
                    color=GRN,
                    size=12,
                    weight=ft.FontWeight.W_600,
                ),
                ft.Text(
                    f"Bloques pendientes por disponibilidad o conflictos: {pendientes}.",
                    color=ORG if pendientes else TG,
                    size=12,
                ),
            ], spacing=4)
            estado.bgcolor = alpha_col(GRN, 0.14)
            estado.border = _ba(1, alpha_col(GRN, 0.5))
            estado.update()
            self._sb_refresh()
            self._mc_refresh()
            self.page.update()

        btn_guardar = ft.ElevatedButton(
            "Guardar ficha y generar horarios",
            icon=ft.Icons.SAVE_ROUNDED,
            bgcolor=BLUE,
            color=WHITE,
            on_click=guardar,
        )
        btn_importar = ft.OutlinedButton(
            "Importar planificación masiva",
            icon=ft.Icons.UPLOAD_FILE_ROUNDED,
            on_click=self._importar_archivo_dialogo,
        )
        return ft.Column([
            ft.Row([
                titulo("Nueva Carga", "Registre manualmente los datos del docente, sus materias y disponibilidad."),
                ft.Container(expand=True),
                btn_importar,
            ], spacing=12, vertical_alignment=ft.CrossAxisAlignment.CENTER),
            ft.Container(height=16),
            ft.Container(
                content=ft.Column([
                    ft.Text("Datos del docente", color=TW, size=15, weight=ft.FontWeight.W_600),
                    divl(),
                    ft.Row([tf_docente, tf_cedula], spacing=12),
                    ft.Container(height=8),
                    ft.Row([tf_correo, dd_condicion], spacing=12),
                    ft.Container(height=8),
                    ft.Row([dd_carrera, dd_ambiente], spacing=12),
                    ft.Container(height=14),
                    tf_materias,
                    ft.Container(height=18),
                    ft.Text("Disponibilidad semanal", color=TW, size=15, weight=ft.FontWeight.W_600),
                    ft.Text("Marque las franjas en las que el docente puede impartir clases.", color=TG, size=11),
                    ft.Container(height=8),
                    ft.Container(
                        content=ft.Column(filas_horario, spacing=4, scroll=ft.ScrollMode.AUTO),
                        bgcolor=BG_P,
                        border=_ba(1, BOR),
                        border_radius=10,
                        padding=_pa(10),
                    ),
                    ft.Container(height=14),
                    estado,
                    ft.Container(height=12),
                    ft.Row([btn_guardar], alignment=ft.MainAxisAlignment.END),
                ], spacing=6),
                bgcolor=BG_C,
                border_radius=14,
                padding=_pa(20),
                border=_ba(1, BOR),
            ),
        ], spacing=0, scroll=ft.ScrollMode.AUTO)

    def _vnc(self) -> ft.Control:
        return self._vnc_ficha_docente()

    def _vh(self) -> ft.Control:
        dd_fc = ddrop(
            "Filtrar Carrera",
            [ft.dropdown.Option("TODAS", "Todas las Carreras")]
            + [ft.dropdown.Option(k, c["nombre"]) for k, c in CARRERAS.items()],
            value=self.horarios_carrera, width=220,
        )
        dd_fd = ddrop(
            "Filtrar Dia",
            [ft.dropdown.Option("TODOS", "Todos los Dias")]
            + [ft.dropdown.Option(d) for d in DIAS],
            value=self.horarios_dia, width=160,
        )
        dd_fdoc = ddrop(
            "Filtrar Docente",
            [ft.dropdown.Option("TODOS", "Todos los Docentes")]
            + [ft.dropdown.Option(d) for d in sorted(DOCENTES)],
            value=self.horarios_docente, width=250,
        )
        tf_busc = campo(
            "Buscar materia, docente o ambiente...",
            value=self.horarios_buscar,
            prefix_icon=ft.Icons.SEARCH,
            width=280,
        )
        dd_rpp = ddrop(
            "Por Pag.",
            [
                ft.dropdown.Option("25", "25 por pag"),
                ft.dropdown.Option("50", "50 por pag"),
                ft.dropdown.Option("100", "100 por pag"),
            ],
            value=str(self.horarios_rpp),
            width=135,
        )

        col_t = ft.Column(spacing=0)
        nav_paginacion = ft.Row(spacing=8, alignment=ft.MainAxisAlignment.CENTER)

        def cambiar_pag(delta: int):
            self.horarios_pag += delta
            build()

        def ir_pag(p: int):
            self.horarios_pag = p
            build()

        def build(e=None):
            self.horarios_carrera = dd_fc.value or "TODAS"
            self.horarios_dia = dd_fd.value or "TODOS"
            self.horarios_docente = dd_fdoc.value or "TODOS"
            self.horarios_buscar = tf_busc.value or ""
            try:
                self.horarios_rpp = int(dd_rpp.value or 25)
            except Exception:
                self.horarios_rpp = 25

            datos = list(self._all)
            if self.horarios_carrera != "TODAS":
                datos = [x for x in datos if x.get("carrera_id") == self.horarios_carrera]
            if self.horarios_dia != "TODOS":
                datos = [x for x in datos if x.get("dia") == self.horarios_dia]
            if self.horarios_docente != "TODOS":
                datos = [x for x in datos if x.get("docente") == self.horarios_docente]

            q = self.horarios_buscar.strip().lower()
            if q:
                datos = [
                    x for x in datos
                    if q in x.get("materia", "").lower()
                    or q in x.get("docente", "").lower()
                    or q in x.get("ambiente", "").lower()
                    or q in str(x.get("seccion", "")).lower()
                    or q in CARRERAS.get(x.get("carrera_id", ""), {}).get("nombre", "").lower()
                ]

            datos.sort(key=lambda x: (x.get("docente", ""), x.get("carrera_id", ""), x.get("dia", ""), x.get("franja", "")))

            total_regs = len(datos)
            total_pags = max(1, (total_regs + self.horarios_rpp - 1) // self.horarios_rpp)
            if self.horarios_pag > total_pags:
                self.horarios_pag = total_pags
            if self.horarios_pag < 1:
                self.horarios_pag = 1

            inicio = (self.horarios_pag - 1) * self.horarios_rpp
            fin = min(inicio + self.horarios_rpp, total_regs)
            pagina_datos = datos[inicio:fin]

            enc = ft.Container(
                content=ft.Row([
                    ft.Container(width=4),
                    ft.Text("CARRERA",  color=TG, size=10, weight=ft.FontWeight.W_700, expand=2),
                    ft.Text("DOCENTE",  color=TG, size=10, weight=ft.FontWeight.W_700, expand=3),
                    ft.Text("MATERIA",  color=TG, size=10, weight=ft.FontWeight.W_700, expand=3),
                    ft.Text("CICLO",    color=TG, size=10, weight=ft.FontWeight.W_700, width=95),
                    ft.Text("DIA",      color=TG, size=10, weight=ft.FontWeight.W_700, width=85),
                    ft.Text("HORARIO",  color=TG, size=10, weight=ft.FontWeight.W_700, width=120),
                    ft.Text("AMBIENTE", color=TG, size=10, weight=ft.FontWeight.W_700, width=95),
                    ft.Text("OK",       color=TG, size=10, weight=ft.FontWeight.W_700, width=35),
                ], spacing=8),
                bgcolor=BG_IN, border_radius=8,
                padding=_ps(12, 10), margin=_mo(bottom=4),
            )

            filas = []
            for i, entry in enumerate(pagina_datos):
                ci  = CARRERAS.get(entry.get("carrera_id", ""), {})
                col = ci.get("color", BLUE)
                bg  = BG_C if i % 2 == 0 else BG_RA
                btn_amb_h = ft.Container(
                    content=ft.Row([
                        ft.Icon(ft.Icons.EDIT_ROUNDED, color=BLUE, size=11),
                        badge(entry.get("ambiente", "P1-01"), BLUE),
                    ], spacing=2, alignment=ft.MainAxisAlignment.CENTER),
                    tooltip="Modificar ambiente asignado a esta clase",
                    on_click=lambda _, itm=entry: self._abrir_dialogo_ambiente(itm, al_actualizar=build),
                    ink=True,
                )
                filas.append(ft.Container(
                    content=ft.Row([
                        ft.Container(width=4, height=32, bgcolor=col, border_radius=2),
                        ft.Text(ci.get("nombre", "")[:18], color=col, size=11, weight=ft.FontWeight.W_600, expand=2),
                        ft.Text(entry.get("docente", "")[:26], color=TL, size=11, expand=3),
                        ft.Text(entry.get("materia", "")[:26], color=TL, size=11, expand=3),
                        ft.Container(badge(entry.get("ciclo", "INTRO")[:5], PURP), width=95),
                        ft.Container(ft.Text(entry.get("dia", ""), color=TG, size=11), width=85),
                        ft.Container(ft.Text(entry.get("franja", ""), color=TG, size=11), width=120),
                        ft.Container(btn_amb_h, width=95),
                        ft.Container(ft.Icon(ft.Icons.CHECK_CIRCLE_ROUNDED, color=GRN, size=15), width=35),
                    ], spacing=8, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                    bgcolor=bg, border_radius=8,
                    padding=_ps(12, 7), margin=_mo(bottom=2),
                ))

            if not pagina_datos:
                filas.append(ft.Container(
                    content=ft.Text("No se encontraron horarios con los criterios actuales.", color=TG, size=12, italic=True),
                    padding=_pa(20), alignment=ft.alignment.center,
                ))

            btn_prev = ft.IconButton(
                icon=ft.Icons.CHEVRON_LEFT_ROUNDED,
                icon_color=TW if self.horarios_pag > 1 else TG,
                disabled=(self.horarios_pag <= 1),
                on_click=lambda _: cambiar_pag(-1),
                tooltip="Pagina anterior",
            )
            btn_next = ft.IconButton(
                icon=ft.Icons.CHEVRON_RIGHT_ROUNDED,
                icon_color=TW if self.horarios_pag < total_pags else TG,
                disabled=(self.horarios_pag >= total_pags),
                on_click=lambda _: cambiar_pag(1),
                tooltip="Pagina siguiente",
            )
            btn_first = ft.IconButton(
                icon=ft.Icons.FIRST_PAGE_ROUNDED,
                icon_color=TW if self.horarios_pag > 1 else TG,
                disabled=(self.horarios_pag <= 1),
                on_click=lambda _: ir_pag(1),
                tooltip="Primera pagina",
            )
            btn_last = ft.IconButton(
                icon=ft.Icons.LAST_PAGE_ROUNDED,
                icon_color=TW if self.horarios_pag < total_pags else TG,
                disabled=(self.horarios_pag >= total_pags),
                on_click=lambda _: ir_pag(total_pags),
                tooltip="Ultima pagina",
            )

            nav_paginacion.controls = [
                ft.Text(f"Mostrando {inicio+1 if total_regs else 0} - {fin} de {total_regs} cursos", color=TG, size=12),
                ft.Container(width=16),
                btn_first,
                btn_prev,
                ft.Container(
                    content=ft.Text(f"Pag. {self.horarios_pag} / {total_pags}", color=TW, size=12, weight=ft.FontWeight.W_600),
                    bgcolor=alpha_col(BLUE, 0.2), border_radius=8, padding=_ps(10, 5), border=_ba(1, alpha_col(BLUE, 0.5)),
                ),
                btn_next,
                btn_last,
            ]

            col_t.controls = [
                enc,
                ft.Column(filas, spacing=0),
                ft.Container(height=14),
                nav_paginacion,
            ]
            if col_t.page:
                col_t.update()

        dd_fc.on_change = lambda _: ir_pag(1)
        dd_fd.on_change = lambda _: ir_pag(1)
        dd_fdoc.on_change = lambda _: ir_pag(1)
        tf_busc.on_change = lambda _: ir_pag(1)
        dd_rpp.on_change = lambda _: ir_pag(1)
        build()

        btn_print_oferta = ft.Container(
            content=ft.Row([
                ft.Icon(ft.Icons.PRINT_ROUNDED, color=BG_P, size=17),
                ft.Text("Imprimir Oferta General (HTML / PDF)", color=BG_P, size=12, weight=ft.FontWeight.W_700),
            ], spacing=8),
            bgcolor=GRN, border_radius=10, padding=_ps(16, 12),
            on_click=lambda _: generar_reporte_html_imprimible(self._all, tipo="oferta"),
            ink=True,
        )

        banner = ft.Container(
            content=ft.Row([
                ft.Icon(ft.Icons.VERIFIED_ROUNDED, color=GRN, size=22),
                ft.Column([
                    ft.Text("Sistema Sin Conflictos de Horario", color=GRN, size=13, weight=ft.FontWeight.W_700),
                    ft.Text("Todos los ambientes y docentes asignados sin superposicion.", color=TG, size=11),
                ], spacing=2),
                ft.Container(expand=True),
                badge("0 CONFLICTOS", GRN),
            ], spacing=14, vertical_alignment=ft.CrossAxisAlignment.CENTER),
            bgcolor=alpha_col(GRN, 0.15), border_radius=12,
            padding=_pa(16), border=_ba(1, alpha_col(GRN, 0.4)),
        )

        filtros_fila = ft.Row(
            [dd_fc, dd_fd, dd_fdoc, dd_rpp, tf_busc],
            spacing=10, vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )

        return ft.Column([
            ft.Row([
                titulo("Horarios Asignados", "1) Oferta academica general con profesores ordenados alfabeticamente."),
                ft.Container(expand=True),
                btn_print_oferta,
            ], spacing=12, vertical_alignment=ft.CrossAxisAlignment.CENTER),
            ft.Container(height=16),
            banner,
            ft.Container(height=14),
            filtros_fila,
            ft.Container(height=12),
            ft.Container(
                content=col_t,
                bgcolor=BG_C, border_radius=14, padding=_pa(16),
                border=_ba(1, BOR),
            ),
            ft.Container(height=10),
        ], spacing=0)

    # ══════════════════════════════════════════════════
    #  VISTA AMBIENTES (ORDENADOS POR PISOS: P1, P2, P3, P4, PB...)
    # ══════════════════════════════════════════════════
    def _va(self) -> ft.Control:
        usados = self.motor.usados()
        st     = self.motor.stats()
        salones_bd = obtener_salones()
        capacidades = {
            str(s["codigo"]).strip().upper(): int(s.get("capacidad") or MAX_EST)
            for s in salones_bd
        }
        ambientes_camillas = {
            str(s["codigo"]).strip().upper()
            for s in salones_bd
            if int(s.get("es_camillas") or 0)
        }
        # Recopilar todos los ambientes únicos
        todos_ambs = list(dict.fromkeys(AMBIENTES + [e.get("ambiente") for e in self._all if e.get("ambiente")]))
        capacidad_total = sum(
            capacidades.get(amb, MAX_EST)
            for amb in todos_ambs
            if amb not in ambientes_camillas
        )

        # Agrupar por piso
        ambientes_por_piso: Dict[str, List[str]] = {}
        for amb in todos_ambs:
            piso, norm, prio = categorizar_piso_ambiente(amb)
            ambientes_por_piso.setdefault(piso, []).append(amb)

        # Ordenar pisos según su prioridad
        pisos_ordenados = sorted(
            ambientes_por_piso.keys(),
            key=lambda p: categorizar_piso_ambiente(ambientes_por_piso[p][0])[2]
        )

        dd_piso = ddrop(
            "Filtrar Piso / Sector",
            [ft.dropdown.Option("TODOS", "Todos los Pisos y Sectores")]
            + [ft.dropdown.Option(p, f"{p} ({len(ambientes_por_piso[p])} aulas)") for p in pisos_ordenados],
            value=self.ambientes_piso_sel, width=240,
        )

        dd_estado = ddrop(
            "Filtrar Estado",
            [
                ft.dropdown.Option("TODOS", "Todos los Estados"),
                ft.dropdown.Option("LIBRES", "Solo Aulas Libres"),
                ft.dropdown.Option("OCUPADOS", "Solo Aulas Ocupadas"),
            ],
            value=self.ambientes_estado_sel, width=190,
        )

        tf_busc = campo(
            "Buscar ambiente (ej: P1-01, Aula 201)...",
            value=self.ambientes_buscar,
            prefix_icon=ft.Icons.SEARCH,
            width=280,
        )

        contenedor_pisos = ft.Column(spacing=20)

        def render_ambientes(e=None):
            self.ambientes_piso_sel = dd_piso.value or "TODOS"
            self.ambientes_estado_sel = dd_estado.value or "TODOS"
            self.ambientes_buscar = tf_busc.value or ""

            q = self.ambientes_buscar.strip().lower()
            secciones_pisos = []

            for piso in pisos_ordenados:
                if self.ambientes_piso_sel != "TODOS" and self.ambientes_piso_sel != piso:
                    continue

                ambs_de_este_piso = sorted(ambientes_por_piso[piso], key=clave_orden_ambiente)
                
                # Filtrar por búsqueda
                if q:
                    ambs_de_este_piso = [a for a in ambs_de_este_piso if q in a.lower()]

                # Filtrar por estado
                if self.ambientes_estado_sel == "LIBRES":
                    ambs_de_este_piso = [a for a in ambs_de_este_piso if a not in usados]
                elif self.ambientes_estado_sel == "OCUPADOS":
                    ambs_de_este_piso = [a for a in ambs_de_este_piso if a in usados]

                if not ambs_de_este_piso:
                    continue

                cards = []
                for amb in ambs_de_este_piso:
                    en_uso = amb in usados
                    clases = [e for e in self._all if e.get("ambiente") == amb]
                    col    = ORG if en_uso else GRN

                    detalles = []
                    if en_uso and clases:
                        primera_c = clases[0]
                        detalles.append(
                            ft.Text(f"{primera_c.get('docente', '')[:16]}", color=TL, size=9, max_lines=1, overflow=ft.TextOverflow.ELLIPSIS)
                        )
                    else:
                        detalles.append(
                            ft.Text("Disponible para asignacion", color=TG, size=9, italic=True)
                        )

                    cards.append(ft.Container(
                        content=ft.Column([
                            ft.Row([
                                ft.Icon(ft.Icons.MEETING_ROOM_ROUNDED, color=col, size=16),
                                ft.Container(expand=True),
                                badge("OCUPADO" if en_uso else "LIBRE", col),
                            ], vertical_alignment=ft.CrossAxisAlignment.CENTER),
                            ft.Container(height=4),
                            ft.Text(amb, color=TW, size=13, weight=ft.FontWeight.BOLD),
                            ft.Text(
                                f"{len(clases)} clase(s) • Camillas"
                                if amb in ambientes_camillas
                                else f"{len(clases)} clase(s) • Cap: {capacidades.get(amb, MAX_EST)}",
                                color=TL, size=10,
                            ),
                            ft.Container(height=2),
                            ft.Text(
                                clases[0].get('docente', '')[:18] if (en_uso and clases) else "Disponible",
                                color=ORG if en_uso else GRN, size=10, weight=ft.FontWeight.W_500,
                                max_lines=1, overflow=ft.TextOverflow.ELLIPSIS,
                            ),
                            ft.Container(height=4),
                            ft.Container(
                                content=ft.Row([
                                    ft.Icon(ft.Icons.EDIT_ROUNDED, size=11, color=col),
                                    ft.Text("Modificar" if en_uso else "Asignar", size=10, color=col, weight=ft.FontWeight.W_600),
                                ], spacing=3, alignment=ft.MainAxisAlignment.CENTER),
                                bgcolor=alpha_col(col, 0.15),
                                border_radius=6,
                                padding=_ps(6, 3),
                            ),
                        ], spacing=2),
                        bgcolor=BG_IN, border_radius=10, padding=_pa(12),
                        border=_ba(1, alpha_col(col, 0.45)), width=155,
                        on_click=lambda _, a=amb, c=clases: self._abrir_dialogo_gestionar_aula(a, c, al_actualizar=self._refrescar_vista_actual),
                        ink=True,
                        tooltip="Clic para modificar los datos, asignar, reasignar o eliminar este ambiente",
                    ))

                # Header de piso con estadísticas
                tot_p = len(ambientes_por_piso[piso])
                lib_p = len([a for a in ambientes_por_piso[piso] if a not in usados])
                ocu_p = tot_p - lib_p

                header_piso = ft.Container(
                    content=ft.Row([
                        ft.Icon(ft.Icons.LAYERS_ROUNDED, color=BLUE, size=20),
                        ft.Text(piso, color=TW, size=16, weight=ft.FontWeight.BOLD),
                        ft.Container(width=10),
                        ft.Text(f"{tot_p} ambientes", color=TG, size=12),
                        ft.Container(expand=True),
                        badge(f"{lib_p} Libres", GRN),
                        badge(f"{ocu_p} Ocupados", ORG),
                    ], vertical_alignment=ft.CrossAxisAlignment.CENTER),
                    bgcolor=BG_IN, border_radius=10, padding=_ps(14, 10),
                )

                secciones_pisos.append(ft.Container(
                    content=ft.Column([
                        header_piso,
                        ft.Container(height=10),
                        ft.Row(cards, wrap=True, spacing=12, run_spacing=12),
                    ], spacing=0),
                    bgcolor=BG_C, border_radius=14, padding=_pa(16),
                    border=_ba(1, BOR),
                ))

            if not secciones_pisos:
                secciones_pisos.append(ft.Container(
                    content=ft.Text("No se encontraron ambientes con los filtros seleccionados.", color=TG, size=13, italic=True),
                    padding=_pa(30), alignment=ft.alignment.center,
                ))

            contenedor_pisos.controls = secciones_pisos
            if contenedor_pisos.page:
                contenedor_pisos.update()

        dd_piso.on_change = render_ambientes
        dd_estado.on_change = render_ambientes
        tf_busc.on_change = render_ambientes
        render_ambientes()

        stats_ambientes = ft.Row([
            stat_card("Total Ambientes", str(st["amb_tot"]),   ft.Icons.MEETING_ROOM_ROUNDED, BLUE, "Registrados"),
            stat_card("En Uso",           str(st["amb_uso"]),   ft.Icons.LOCK_ROUNDED,          ORG,  "Asignados"),
            stat_card("Disponibles",      str(st["amb_lib"]),   ft.Icons.CHECK_ROUNDED,         GRN,  "Libres"),
            stat_card("Capacidad Total",  str(capacidad_total), ft.Icons.PERSON_ROUNDED,        PURP, "Estudiantes"),
        ], spacing=14)

        btn_nuevo_amb = ft.ElevatedButton(
            "+ Nuevo Ambiente",
            icon=ft.Icons.ADD_LOCATION_ALT_ROUNDED,
            bgcolor=BLUE,
            color=WHITE,
            on_click=lambda _: self._abrir_dialogo_nuevo_ambiente(al_actualizar=self._refrescar_vista_actual),
        )

        return ft.Column([
            ft.Row([
                titulo("Gestion de Ambientes por Piso", f"Distribucion ordenada por plantas de los {st['amb_tot']} ambientes academicos."),
                ft.Container(expand=True),
                btn_nuevo_amb,
            ], spacing=12, vertical_alignment=ft.CrossAxisAlignment.CENTER),
            ft.Container(height=18),
            stats_ambientes,
            ft.Container(height=18),
            ft.Row([dd_piso, dd_estado, tf_busc], spacing=10, wrap=True, vertical_alignment=ft.CrossAxisAlignment.CENTER),
            ft.Container(height=14),
            contenedor_pisos,
            ft.Container(height=10),
        ], spacing=0)

    # ══════════════════════════════════════════════════
    #  VISTA DOCENTES (VISTA PREVIA Y DIRECTORIO SIMÉTRICO)
    # ══════════════════════════════════════════════════
    def _vdoc(self) -> ft.Control:
        docs_ordenados = sorted(DOCENTES)
        if not self.docente_calendario_sel and docs_ordenados:
            self.docente_calendario_sel = docs_ordenados[0]

        dd_cal_doc = ddrop(
            "Seleccionar Profesor",
            [ft.dropdown.Option(d) for d in docs_ordenados],
            value=self.docente_calendario_sel if self.docente_calendario_sel in docs_ordenados else (docs_ordenados[0] if docs_ordenados else None),
            expand=True,
        )

        cont_calendario = ft.Column(spacing=4)
        cont_resumen_doc = ft.Column(spacing=10)
        cont_tarjetas_clases = ft.Column(spacing=8)
        lbl_contador = ft.Text("", color=TW, size=11, weight=ft.FontWeight.W_700)

        def render_docente(e=None):
            doc_actual = dd_cal_doc.value or (docs_ordenados[0] if docs_ordenados else "")
            self.docente_calendario_sel = doc_actual
            asig_p = [x for x in self._all if x.get("docente") == doc_actual]
            dias_cal = ["LUNES", "MARTES", "MIERCOLES", "JUEVES", "VIERNES", "SABADO"]

            # Actualizar contador de posición
            idx = docs_ordenados.index(doc_actual) if doc_actual in docs_ordenados else 0
            lbl_contador.value = f"Docente {idx + 1} de {len(docs_ordenados)}"

            # Estadísticas del docente seleccionado
            tot_cursos = len(asig_p)
            tot_horas = sum(int(x.get("horas", 2)) for x in asig_p)
            amb_usados = sorted(list({x.get("ambiente", "") for x in asig_p if x.get("ambiente")}))
            carreras_doc = sorted(list({CARRERAS.get(x.get("carrera_id", ""), {}).get("nombre", "") for x in asig_p}))

            def mini_badge_info(lbl, val, col, ico):
                return ft.Container(
                    content=ft.Row([
                        ft.Container(
                            content=ft.Icon(ico, color=col, size=17),
                            bgcolor=alpha_col(col, 0.15), border_radius=8,
                            padding=_pa(8),
                        ),
                        ft.Column([
                            ft.Text(lbl, color=TG, size=10, weight=ft.FontWeight.W_600),
                            ft.Text(val, color=TW, size=12, weight=ft.FontWeight.W_700, max_lines=1, overflow=ft.TextOverflow.ELLIPSIS),
                        ], spacing=1, expand=True),
                    ], spacing=10, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                    bgcolor=BG_IN, border_radius=10, padding=_ps(12, 8),
                    border=_ba(1, alpha_col(col, 0.4)), expand=True,
                )

            info_docs = obtener_info_profesores_dict()
            info_actual = info_docs.get(doc_actual.strip().upper(), {})
            cond_doc = info_actual.get("condicion", "Contratado/a")
            ced_doc = info_actual.get("cedula", "")

            cont_resumen_doc.controls = [
                ft.Row([
                    mini_badge_info("CURSOS ASIGNADOS", f"{tot_cursos} materias activas", BLUE, ft.Icons.BOOK_ROUNDED),
                    mini_badge_info("CARGA HORARIA", f"{tot_horas} hrs / semana", CYAN, ft.Icons.SCHEDULE_ROUNDED),
                ], spacing=12),
                ft.Row([
                    mini_badge_info("AMBIENTES EN USO", ", ".join(amb_usados) if amb_usados else "Ninguno", ORG, ft.Icons.MEETING_ROOM_ROUNDED),
                    mini_badge_info("CONDICIÓN DOCENTE", f"{cond_doc}" + (f" ({ced_doc})" if ced_doc else ""), PURP, ft.Icons.BADGE_ROUNDED),
                ], spacing=12),
            ]

            # Construir matriz de horario semanal
            enc_cal = ft.Container(
                content=ft.Row([
                    ft.Text("HORARIO", color=TG, size=10, weight=ft.FontWeight.W_700, width=110),
                    *[ft.Text(d, color=TG, size=10, weight=ft.FontWeight.W_700, expand=True) for d in dias_cal],
                ], spacing=6),
                bgcolor=BG_IN, border_radius=8, padding=_ps(10, 8),
            )
            filas_cal = [enc_cal]
            franjas_mostrar = list(dict.fromkeys(FRANJAS[:7] + [x["franja"] for x in asig_p]))

            for idx_f, fr in enumerate(franjas_mostrar):
                celdas = []
                for d in dias_cal:
                    clases_celda = [x for x in asig_p if x.get("dia") == d and x.get("franja") == fr]
                    if clases_celda:
                        items_c = []
                        for cc in clases_celda:
                            col_c = CARRERAS.get(cc.get("carrera_id", ""), {}).get("color", BLUE)
                            items_c.append(ft.Container(
                                content=ft.Column([
                                    ft.Text(cc.get("materia", "")[:24], color=TW, size=10, weight=ft.FontWeight.W_700, max_lines=1, overflow=ft.TextOverflow.ELLIPSIS),
                                    ft.Text(CARRERAS.get(cc.get("carrera_id", ""), {}).get("nombre", "")[:18], color=col_c, size=9, max_lines=1),
                                    ft.Row([
                                        badge(cc.get("ambiente", ""), BLUE),
                                        badge(f"Sec {cc.get('seccion', '10100')}", PURP),
                                    ], spacing=4),
                                ], spacing=2),
                                bgcolor=alpha_col(col_c, 0.2), border_radius=6, padding=_pa(6), border=_ba(1, alpha_col(col_c, 0.5)),
                            ))
                        celdas.append(ft.Column(items_c, spacing=3, expand=True))
                    else:
                        celdas.append(ft.Container(
                            alignment=ft.alignment.center, expand=True, height=48,
                            content=ft.Text("-", color=alpha_col(TG, 0.35), size=12),
                        ))

                filas_cal.append(ft.Container(
                    content=ft.Row([
                        ft.Container(ft.Text(fr, color=TL, size=10, weight=ft.FontWeight.W_600), width=110),
                        *celdas,
                    ], spacing=6, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                    bgcolor=BG_C if idx_f % 2 == 0 else BG_RA,
                    border_radius=6, padding=_ps(10, 6),
                ))

            cont_calendario.controls = filas_cal

            # Detalle individual de asignaciones
            cards_asig = []
            for entry in asig_p:
                ci = CARRERAS.get(entry.get("carrera_id", ""), {})
                col_car = ci.get("color", BLUE)
                btn_amb_doc = ft.Container(
                    content=ft.Row([
                        ft.Icon(ft.Icons.EDIT_ROUNDED, color=ORG, size=11),
                        badge(entry.get("ambiente", "P1-01"), ORG),
                    ], spacing=2, alignment=ft.MainAxisAlignment.CENTER),
                    tooltip="Modificar ambiente asignado a esta materia",
                    on_click=lambda _, itm=entry: self._abrir_dialogo_ambiente(itm, al_actualizar=render_docente),
                    ink=True,
                )
                cards_asig.append(ft.Container(
                    content=ft.Row([
                        ft.Container(width=4, height=36, bgcolor=col_car, border_radius=2),
                        ft.Column([
                            ft.Text(entry.get("materia", ""), color=TW, size=12, weight=ft.FontWeight.W_700),
                            ft.Text(ci.get("nombre", ""), color=col_car, size=11),
                        ], spacing=2, expand=2),
                        ft.Text(f"{entry.get('horas', 2)} hrs/sem", color=TG, size=11, width=80),
                        ft.Container(badge(entry.get("dia", ""), BLUE), width=90),
                        ft.Container(ft.Text(entry.get("franja", ""), color=TW, size=11, weight=ft.FontWeight.W_600), width=110),
                        ft.Container(btn_amb_doc, width=95),
                        ft.Container(badge(f"Sec {entry.get('seccion', 10100)}", PURP), width=80),
                        ft.Icon(ft.Icons.VERIFIED_ROUNDED, color=GRN, size=18),
                    ], spacing=10, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                    bgcolor=BG_IN, border_radius=8, padding=_ps(14, 8), border=_ba(1, BOR),
                ))

            if not cards_asig:
                cards_asig.append(ft.Container(
                    content=ft.Text("Este docente no tiene materias asignadas actualmente.", color=TG, size=12, italic=True),
                    padding=_pa(16),
                ))

            cont_tarjetas_clases.controls = cards_asig

            if cont_calendario.page:
                cont_calendario.update()
            if cont_resumen_doc.page:
                cont_resumen_doc.update()
            if cont_tarjetas_clases.page:
                cont_tarjetas_clases.update()
            if lbl_contador.page:
                lbl_contador.update()

        def ant_doc(e):
            idx = docs_ordenados.index(self.docente_calendario_sel) if self.docente_calendario_sel in docs_ordenados else 0
            if idx > 0:
                self.docente_calendario_sel = docs_ordenados[idx - 1]
                dd_cal_doc.value = self.docente_calendario_sel
                dd_cal_doc.update()
                render_docente()
                render_docentes_cards()

        def sig_doc(e):
            idx = docs_ordenados.index(self.docente_calendario_sel) if self.docente_calendario_sel in docs_ordenados else 0
            if idx < len(docs_ordenados) - 1:
                self.docente_calendario_sel = docs_ordenados[idx + 1]
                dd_cal_doc.value = self.docente_calendario_sel
                dd_cal_doc.update()
                render_docente()
                render_docentes_cards()

        btn_ant = ft.Container(
            content=ft.Row([ft.Icon(ft.Icons.ARROW_BACK_ROUNDED, color=TL, size=15), ft.Text("Ant.", color=TL, size=11)], spacing=4),
            bgcolor=BG_IN, border_radius=8, padding=_ps(10, 8), border=_ba(1, BOR), on_click=ant_doc, ink=True,
        )
        btn_sig = ft.Container(
            content=ft.Row([ft.Text("Sig.", color=TL, size=11), ft.Icon(ft.Icons.ARROW_FORWARD_ROUNDED, color=TL, size=15)], spacing=4),
            bgcolor=BG_IN, border_radius=8, padding=_ps(10, 8), border=_ba(1, BOR), on_click=sig_doc, ink=True,
        )
        badge_pos = ft.Container(
            content=lbl_contador,
            bgcolor=alpha_col(BLUE, 0.18), border_radius=8, padding=_ps(10, 7),
            border=_ba(1, alpha_col(BLUE, 0.4)),
        )

        dd_cal_doc.on_change = lambda e: (render_docente(), render_docentes_cards())
        render_docente()

        # Botones de impresión superiores compactos
        btn_imp_este_doc = ft.Container(
            content=ft.Row([
                ft.Icon(ft.Icons.PRINT_ROUNDED, color=WHITE, size=16),
                ft.Text("Imprimir Horario Docente", color=WHITE, size=11, weight=ft.FontWeight.W_700),
            ], spacing=6),
            bgcolor=GRN, border_radius=8, padding=_ps(12, 9),
            on_click=lambda _: generar_reporte_html_imprimible(self._all, tipo="profesor", docente_filtro=self.docente_calendario_sel),
            ink=True,
        )

        btn_imp_todos = ft.Container(
            content=ft.Row([
                ft.Icon(ft.Icons.PICTURE_AS_PDF_ROUNDED, color=WHITE, size=15),
                ft.Text(f"Imprimir Todos ({len(docs_ordenados)})", color=WHITE, size=11, weight=ft.FontWeight.W_600),
            ], spacing=6),
            bgcolor=PURP, border_radius=8, padding=_ps(12, 9),
            on_click=lambda _: generar_reporte_html_imprimible(self._all, tipo="profesor", docente_filtro="TODOS"),
            ink=True,
        )

        def click_cambiar_amb_doc(e=None):
            doc_sel = dd_cal_doc.value or self.docente_calendario_sel or (docs_ordenados[0] if docs_ordenados else "")
            if not doc_sel:
                self._mostrar_notificacion("No hay un docente seleccionado o registrado.", color=ORG)
                return
            self._abrir_dialogo_ambiente_docente_completo(doc_sel, al_actualizar=render_docente)

        btn_cambiar_amb_doc = ft.ElevatedButton(
            "Asignar / Cambiar Aula",
            icon=ft.Icons.MEETING_ROOM_ROUNDED,
            bgcolor=ORG,
            color=WHITE,
            on_click=click_cambiar_amb_doc,
        )

        def docente_actual():
            return dd_cal_doc.value or self.docente_calendario_sel or (docs_ordenados[0] if docs_ordenados else "")

        def editar_docente_actual(e=None):
            doc_sel = docente_actual()
            if not doc_sel:
                self._mostrar_notificacion("No hay un docente seleccionado o registrado.", color=ORG)
                return
            self._abrir_dialogo_editar_docente(doc_sel, al_actualizar=self._refrescar_vista_actual)

        def eliminar_docente_actual(e=None):
            doc_sel = docente_actual()
            if not doc_sel:
                self._mostrar_notificacion("No hay un docente seleccionado o registrado.", color=ORG)
                return
            total = sum(1 for x in self._all if x.get("docente") == doc_sel)
            self._dialogo_eliminar_docente(doc_sel, total, al_actualizar=self._refrescar_vista_actual)

        btn_editar_doc = ft.ElevatedButton(
            "Modificar Datos",
            icon=ft.Icons.MANAGE_ACCOUNTS_ROUNDED,
            bgcolor=PURP,
            color=WHITE,
            on_click=editar_docente_actual,
        )
        btn_eliminar_doc = ft.ElevatedButton(
            "Eliminar Docente",
            icon=ft.Icons.DELETE_FOREVER_ROUNDED,
            bgcolor=RED,
            color=WHITE,
            on_click=eliminar_docente_actual,
        )

        barra_seleccion_docente = ft.Container(
            content=ft.Column([
                ft.Row([
                    dd_cal_doc,
                    btn_ant,
                    btn_sig,
                    badge_pos,
                ], spacing=8, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                ft.Row([
                    btn_cambiar_amb_doc,
                    btn_editar_doc,
                    btn_eliminar_doc,
                ], spacing=8, alignment=ft.MainAxisAlignment.END),
            ], spacing=8, horizontal_alignment=ft.CrossAxisAlignment.STRETCH),
            bgcolor=BG_C, border_radius=12, padding=_pa(12), border=_ba(1, BOR),
        )

        panel_matriz_semanal = ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.Icon(ft.Icons.CALENDAR_VIEW_WEEK_ROUNDED, color=BLUE, size=18),
                    ft.Text("Horario Semanal Tipo Calendario", color=TW, size=14, weight=ft.FontWeight.W_700),
                ], spacing=8),
                ft.Container(height=4),
                cont_calendario,
            ], spacing=6),
            bgcolor=BG_C, border_radius=14, padding=_pa(16), border=_ba(1, BOR),
        )

        panel_detalle_cursos = ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.Icon(ft.Icons.LIST_ALT_ROUNDED, color=CYAN, size=18),
                    ft.Text("Detalle de Cargas Asignadas al Docente", color=TW, size=14, weight=ft.FontWeight.W_700),
                ], spacing=8),
                ft.Container(height=4),
                cont_tarjetas_clases,
            ], spacing=6),
            bgcolor=BG_C, border_radius=14, padding=_pa(16), border=_ba(1, BOR),
        )

        # ── DIRECTORIO SIMÉTRICO DE PROFESORES (3 COLUMNAS RESPONSIVAS AL 100%) ──
        tf_busc_doc = campo(
            "Buscar profesor en el directorio...",
            value=self.docentes_buscar,
            prefix_icon=ft.Icons.SEARCH,
            width=260,
        )

        grid_docentes_col = ft.Column(spacing=10)
        nav_docentes = ft.Row(spacing=8, alignment=ft.MainAxisAlignment.CENTER)

        def cambiar_pag_doc(delta: int):
            self.docentes_pag += delta
            render_docentes_cards()

        def ir_pag_doc(p: int):
            self.docentes_pag = p
            render_docentes_cards()

        def render_docentes_cards(e=None):
            self.docentes_buscar = tf_busc_doc.value or ""
            q = self.docentes_buscar.strip().lower()

            docs_filtrados = docs_ordenados
            if q:
                docs_filtrados = [d for d in docs_ordenados if q in d.lower()]

            tot_filtrados = len(docs_filtrados)
            tot_pags_doc = max(1, (tot_filtrados + self.docentes_rpp - 1) // self.docentes_rpp)
            if self.docentes_pag > tot_pags_doc:
                self.docentes_pag = tot_pags_doc
            if self.docentes_pag < 1:
                self.docentes_pag = 1

            ini_doc = (self.docentes_pag - 1) * self.docentes_rpp
            fin_doc = min(ini_doc + self.docentes_rpp, tot_filtrados)
            pagina_docs = docs_filtrados[ini_doc:fin_doc]

            cards = []
            for doc in pagina_docs:
                asig = [entry for entry in self._all if entry.get("docente") == doc]
                nombres = doc.title().split()
                ini = nombres[0][0] if nombres else "?"
                tot_h = sum(int(x.get("horas", 2)) for x in asig)
                amb_u = sorted(list({x.get("ambiente", "") for x in asig if x.get("ambiente")}))
                carr_u = sorted(list({CARRERAS.get(x.get("carrera_id", ""), {}).get("nombre", "") for x in asig}))

                es_seleccionado = (doc == self.docente_calendario_sel)
                col_borde = BLUE if es_seleccionado else BOR
                bg_card = alpha_col(BLUE, 0.16) if es_seleccionado else BG_IN

                def click_doc(ev, d_sel=doc):
                    self.docente_calendario_sel = d_sel
                    dd_cal_doc.value = d_sel
                    dd_cal_doc.update()
                    render_docente()
                    render_docentes_cards()

                cards.append(ft.Container(
                    content=ft.Row([
                        ft.Container(
                            content=ft.Text(ini, color=TW, size=14, weight=ft.FontWeight.BOLD),
                            bgcolor=BLUE if es_seleccionado else alpha_col(BLUE, 0.4),
                            width=36, height=36, border_radius=18,
                            alignment=ft.alignment.center,
                        ),
                        ft.Column([
                            ft.Text(
                                " ".join(nombres[:3]),
                                color=TW if es_seleccionado else TL,
                                size=11, weight=ft.FontWeight.W_700 if es_seleccionado else ft.FontWeight.W_500,
                                max_lines=1, overflow=ft.TextOverflow.ELLIPSIS,
                            ),
                            ft.Text(
                                f"{len(asig)} materias • {tot_h} hrs/sem",
                                color=BLUE if es_seleccionado else TG,
                                size=10,
                            ),
                            ft.Text(
                                (carr_u[0][:16] if carr_u else "Sin carrera") + (" • " + amb_u[0] if amb_u else ""),
                                color=TG, size=9, max_lines=1, overflow=ft.TextOverflow.ELLIPSIS,
                            ),
                        ], spacing=1, expand=True),
                        ft.PopupMenuButton(
                            icon=ft.Icons.MORE_VERT_ROUNDED,
                            icon_color=TG,
                            tooltip="Acciones del docente",
                            items=[
                                ft.PopupMenuItem(
                                    text="Modificar datos",
                                    icon=ft.Icons.EDIT_ROUNDED,
                                    on_click=lambda ev, d=doc: self._abrir_dialogo_editar_docente(
                                        d, al_actualizar=self._refrescar_vista_actual
                                    ),
                                ),
                                ft.PopupMenuItem(
                                    text="Eliminar docente",
                                    icon=ft.Icons.DELETE_FOREVER_ROUNDED,
                                    on_click=lambda ev, d=doc: self._dialogo_eliminar_docente(
                                        d,
                                        sum(1 for x in self._all if x.get("docente") == d),
                                        al_actualizar=self._refrescar_vista_actual,
                                    ),
                                ),
                            ],
                        ),
                        ft.Icon(
                            ft.Icons.CHECK_CIRCLE_ROUNDED if es_seleccionado else ft.Icons.CHEVRON_RIGHT_ROUNDED,
                            color=BLUE if es_seleccionado else alpha_col(TG, 0.5),
                            size=16,
                        ),
                    ], spacing=8, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                    bgcolor=bg_card, border_radius=10, padding=_ps(12, 10),
                    border=_ba(1.5 if es_seleccionado else 1, col_borde),
                    expand=True, height=72,
                    on_click=click_doc, ink=True,
                ))

            filas_doc = []
            for i in range(0, len(cards), 3):
                grupo = cards[i:i+3]
                while len(grupo) < 3:
                    grupo.append(ft.Container(expand=True))
                filas_doc.append(ft.Row(grupo, spacing=12))

            if not cards:
                filas_doc.append(ft.Container(
                    content=ft.Text("No se encontraron profesores con ese criterio de búsqueda.", color=TG, size=12, italic=True),
                    padding=_pa(20), alignment=ft.alignment.center,
                ))

            # Controles de navegación de página del directorio
            btn_prev_d = ft.IconButton(
                icon=ft.Icons.CHEVRON_LEFT_ROUNDED,
                icon_color=TW if self.docentes_pag > 1 else TG,
                disabled=(self.docentes_pag <= 1),
                on_click=lambda _: cambiar_pag_doc(-1),
                tooltip="Página anterior",
            )
            btn_next_d = ft.IconButton(
                icon=ft.Icons.CHEVRON_RIGHT_ROUNDED,
                icon_color=TW if self.docentes_pag < tot_pags_doc else TG,
                disabled=(self.docentes_pag >= tot_pags_doc),
                on_click=lambda _: cambiar_pag_doc(1),
                tooltip="Página siguiente",
            )
            nav_docentes.controls = [
                ft.Text(f"Mostrando {ini_doc+1 if tot_filtrados else 0} - {fin_doc} de {tot_filtrados} profesores", color=TG, size=11),
                ft.Container(width=12),
                btn_prev_d,
                ft.Container(
                    content=ft.Text(f"Pág. {self.docentes_pag} / {tot_pags_doc}", color=TW, size=11, weight=ft.FontWeight.W_600),
                    bgcolor=alpha_col(BLUE, 0.2), border_radius=6, padding=_ps(8, 4), border=_ba(1, alpha_col(BLUE, 0.4)),
                ),
                btn_next_d,
            ]

            grid_docentes_col.controls = [
                *filas_doc,
                ft.Container(height=8),
                nav_docentes,
            ]
            if grid_docentes_col.page:
                grid_docentes_col.update()

        tf_busc_doc.on_change = lambda _: (setattr(self, "docentes_pag", 1), render_docentes_cards())
        render_docentes_cards()

        encabezado_docentes = ft.Row([
            titulo("Gestion de Docentes", "Horarios semanales, seleccion interactiva y reporte PDF."),
            ft.Container(expand=True),
            btn_imp_este_doc,
            btn_imp_todos,
        ], spacing=10, vertical_alignment=ft.CrossAxisAlignment.CENTER)

        return ft.Column([
            encabezado_docentes,
            ft.Container(height=16),
            barra_seleccion_docente,
            ft.Container(height=12),
            cont_resumen_doc,
            ft.Container(height=14),
            panel_matriz_semanal,
            ft.Container(height=14),
            panel_detalle_cursos,
            ft.Container(height=20),
            ft.Row([
                ft.Text(f"Directorio de Profesores ({len(docs_ordenados)}) — Haz clic para ver su horario:", color=TW, size=14, weight=ft.FontWeight.W_600),
                ft.Container(expand=True),
                tf_busc_doc,
            ], spacing=10, vertical_alignment=ft.CrossAxisAlignment.CENTER),
            ft.Container(height=12),
            grid_docentes_col,
            ft.Container(height=10),
        ], spacing=0)

    def _abrir_dialogo_editar_docente(self, docente: str, al_actualizar=None):
        """Permite modificar los datos de un docente o eliminarlo individualmente."""
        info = obtener_info_profesores_dict().get(str(docente).strip().upper(), {})
        materias_actuales = obtener_materias_profesor(docente)
        tf_nombre = campo("Nombre completo", value=docente, expand=True)
        tf_cedula = campo("Cédula", value=info.get("cedula", ""), width=190)
        tf_email = campo("Correo electrónico", value=info.get("email", ""), expand=True)
        tf_materias = campo(
            "Materias que dicta (una materia por línea)",
            value="\n".join(materias_actuales),
            multiline=True,
            min_lines=3,
            max_lines=7,
            expand=True,
        )
        dd_cond = ddrop(
            "Condición",
            [
                ft.dropdown.Option("Contratado/a"),
                ft.dropdown.Option("Fijo/a"),
                ft.dropdown.Option("Jubilado/a"),
                ft.dropdown.Option("Invitado/a"),
            ],
            value=info.get("condicion", "Contratado/a") or "Contratado/a",
            width=210,
        )
        total_clases = sum(1 for x in self._all if x.get("docente") == docente)

        def cerrar(e=None):
            self._cerrar_modal(dlg)

        def guardar(e=None):
            nombre_nuevo = (tf_nombre.value or "").strip()
            if not nombre_nuevo:
                tf_nombre.error_text = "El nombre del docente es obligatorio"
                tf_nombre.update()
                return
            materias_nuevas = [
                materia.strip()
                for materia in (tf_materias.value or "").splitlines()
                if materia.strip()
            ]
            if not materias_nuevas:
                tf_materias.error_text = "Debe indicar al menos una materia."
                tf_materias.update()
                return

            ok = actualizar_profesor(
                docente,
                nombre_nuevo.upper(),
                cedula=(tf_cedula.value or "").strip(),
                email=(tf_email.value or "").strip(),
                condicion=dd_cond.value or "Contratado/a",
                materias=materias_nuevas,
            )
            if not ok:
                tf_nombre.error_text = "Verifique que el nombre del docente no esté repetido."
                tf_materias.error_text = (
                    "No se pudo guardar. Revise que las materias no sean ambiguas y que "
                    "el docente tenga una carrera registrada para las materias nuevas."
                )
                tf_materias.update()
                tf_nombre.update()
                return

            self._cerrar_modal(dlg)
            res_sched = generar_horarios_automaticos()
            self.docente_calendario_sel = nombre_nuevo.upper()
            self.motor.limpiar_memoria()
            self._all = self.motor.sincronizar_desde_sqlite()
            recargar_docentes_desde_bd()
            if self.docente_calendario_sel not in DOCENTES:
                self.docente_calendario_sel = DOCENTES[0] if DOCENTES else ""
            self.ultimo_mensaje = f"Datos del docente '{nombre_nuevo.upper()}' actualizados correctamente."
            self._sb_refresh()
            if al_actualizar:
                al_actualizar()
            else:
                self._refrescar_vista_actual()
            if res_sched.get("pendientes"):
                self._mostrar_notificacion(
                    f"Docente actualizado; {res_sched['pendientes']} bloques quedaron pendientes por conflictos o disponibilidad.",
                    color=ORG,
                )
            else:
                self._mostrar_notificacion(f"Docente '{nombre_nuevo.upper()}' actualizado.", color=GRN)

        def abrir_confirmacion_eliminar(e=None):
            self._cerrar_modal(dlg)
            self._dialogo_eliminar_docente(docente, total_clases, al_actualizar=al_actualizar)

        dlg = ft.AlertDialog(
            modal=True,
            bgcolor=BG_C,
            title=ft.Row([
                ft.Icon(ft.Icons.MANAGE_ACCOUNTS_ROUNDED, color=PURP, size=24),
                ft.Text(f"Modificar Docente: {docente}", size=16, weight=ft.FontWeight.BOLD, color=TW),
            ], spacing=8),
            content=ft.Container(
                content=ft.Column([
                    ft.Text("Edite los datos del docente o elimínelo individualmente del sistema.", size=12, color=TL),
                    ft.Container(height=6),
                    tf_nombre,
                    ft.Row([tf_cedula, dd_cond], spacing=10),
                    tf_email,
                    tf_materias,
                    ft.Container(height=4),
                    ft.Text(
                        f"Este docente tiene {total_clases} clase(s) registrada(s).",
                        size=11, color=TG, italic=True,
                    ),
                ], spacing=8, tight=True),
                width=520,
            ),
            actions=[
                ft.TextButton("Cancelar", on_click=cerrar),
                ft.ElevatedButton(
                    "Eliminar Docente",
                    icon=ft.Icons.DELETE_FOREVER_ROUNDED,
                    bgcolor=RED,
                    color=WHITE,
                    on_click=abrir_confirmacion_eliminar,
                ),
                ft.ElevatedButton(
                    "Guardar Cambios",
                    icon=ft.Icons.SAVE_ROUNDED,
                    bgcolor=PURP,
                    color=WHITE,
                    on_click=guardar,
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        self._abrir_modal(dlg)

    def _dialogo_eliminar_docente(self, docente: str, total_clases: int = 0, al_actualizar=None):
        """Confirma y elimina un docente individualmente."""
        def cerrar(e=None):
            self._cerrar_modal(dlg)

        def confirmar(e=None):
            self._cerrar_modal(dlg)
            if not eliminar_profesor(docente):
                self._mostrar_notificacion(f"No se pudo eliminar al docente {docente}.", color=RED)
                return
            self.motor.limpiar_memoria()
            self._all = self.motor.sincronizar_desde_sqlite()
            recargar_docentes_desde_bd()
            if self.docente_calendario_sel == docente or self.docente_calendario_sel not in DOCENTES:
                self.docente_calendario_sel = DOCENTES[0] if DOCENTES else ""
            self.ultimo_mensaje = f"Docente '{docente}' eliminado del sistema."
            self._sb_refresh()
            if al_actualizar:
                al_actualizar()
            else:
                self._refrescar_vista_actual()
            self._mostrar_notificacion(f"Docente '{docente}' eliminado.", color=BLUE)

        detalle_clases = (
            f"\n\nEste docente tiene {total_clases} clase(s) asignada(s). "
            "Si continúa, también se eliminarán sus cargas y asignaciones de horario."
            if total_clases else ""
        )

        dlg = ft.AlertDialog(
            modal=True,
            bgcolor=BG_C,
            title=ft.Row([
                ft.Icon(ft.Icons.WARNING_ROUNDED, color=RED, size=24),
                ft.Text(f"¿Eliminar al docente {docente}?", size=16, weight=ft.FontWeight.BOLD, color=TW),
            ], spacing=8),
            content=ft.Text(
                "Esta acción eliminará al docente de forma individual y no se puede deshacer."
                + detalle_clases,
                size=13, color=TL,
            ),
            actions=[
                ft.TextButton("Cancelar", on_click=cerrar),
                ft.ElevatedButton(
                    "Sí, Eliminar Docente",
                    icon=ft.Icons.DELETE_FOREVER_ROUNDED,
                    bgcolor=RED,
                    color=WHITE,
                    on_click=confirmar,
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        self._abrir_modal(dlg)

    # ══════════════════════════════════════════════════
    #  VISTA CARRERAS (RESPONSIVE ROW SIMÉTRICO)
    # ══════════════════════════════════════════════════
    def _vc(self) -> ft.Control:
        paneles = []
        for kid, c in CARRERAS.items():
            col    = c["color"]
            asig   = [e for e in self._all if e["carrera_id"]==kid]
            pensum = MATERIAS.get(kid, [])
            if not pensum and asig:
                mats_dict = {}
                for item in asig:
                    mnom = item.get("materia", "")
                    if mnom and mnom not in mats_dict:
                        mats_dict[mnom] = {
                            "nombre": mnom,
                            "horas": item.get("horas", 2),
                            "ciclo": item.get("ciclo", "PROFESIONAL"),
                        }
                pensum = list(mats_dict.values())

            mat_r  = [
                ft.Container(
                    content=ft.Row([
                        ft.Container(width=7, height=7, bgcolor=col, border_radius=4),
                        ft.Text(m["nombre"], color=TW, size=11, expand=True, max_lines=1, overflow=ft.TextOverflow.ELLIPSIS),
                        ft.Text(f"{m['horas']}h", color=TG, size=10),
                        badge(m["ciclo"][:5], col),
                    ], spacing=8),
                    padding=_ps(0, 5),
                    border=_bb(1, BOR),
                )
                for m in pensum[:12]
            ]
            if not mat_r:
                mat_r = [
                    ft.Container(
                        content=ft.Text("Sin materias registradas aún.", color=TG, size=11, italic=True),
                        padding=_pa(10),
                    )
                ]

            paneles.append(ft.Container(
                content=ft.Column([
                    ft.Container(content=ft.Row([
                        ft.Container(width=4, height=48, bgcolor=col, border_radius=2),
                        ft.Column([
                            ft.Text(c["nombre"],  color=TW, size=14, weight=ft.FontWeight.BOLD, max_lines=1, overflow=ft.TextOverflow.ELLIPSIS),
                            ft.Text(c["mencion"], color=TG, size=11, max_lines=1, overflow=ft.TextOverflow.ELLIPSIS),
                        ], spacing=3, expand=True),
                        badge(f"{len(asig)} cursos", col),
                    ], spacing=12, vertical_alignment=ft.CrossAxisAlignment.CENTER), height=52),
                    divl(),
                    ft.Container(height=8),
                    ft.Text("Pensum de Materias", color=TG, size=11, weight=ft.FontWeight.W_600),
                    ft.Container(height=4),
                    *mat_r,
                ], spacing=4),
                bgcolor=BG_C, border_radius=14, padding=_pa(18),
                border=_ba(1, alpha_col(col, 0.35)),
                expand=True,
            ))

        filas_carreras = []
        for i in range(0, len(paneles), 2):
            par = paneles[i:i+2]
            if len(par) == 2:
                filas_carreras.append(ft.Row(par, spacing=16))
            else:
                filas_carreras.append(ft.Row([par[0], ft.Container(expand=True)], spacing=16))
            filas_carreras.append(ft.Container(height=16))

        return ft.Column([
            titulo("Gestion de Carreras", "Carreras registradas y su pensum academico en el sistema."),
            ft.Container(height=20),
            *filas_carreras,
            ft.Container(height=10),
        ], spacing=0)

# ═══════════════════════════════════════════════════════
#  ENTRY POINT / ARRANQUE INFALIBLE
# ═══════════════════════════════════════════════════════
def main(page: ft.Page) -> None:
    App(page)


def _parchar_flet_escritorio_interactivo() -> None:
    """
    Asegura que cuando Flet lance flet.exe en Windows (incluso desde una terminal en segundo plano),
    la ventana nativa se cree siempre en el escritorio interactivo principal (WinSta0\\Default).
    En Linux asegura que no intente modo web si corre en Wayland o falta DISPLAY, y resuelve libmpv.
    """
    if os.name != "nt":
        if not os.environ.get("DISPLAY"):
            os.environ["DISPLAY"] = ":0"
        try:
            import flet_desktop
            orig_locate = flet_desktop.__locate_and_unpack_flet_view
            def _patched_locate(page_url, assets_dir, hidden):
                args, flet_env, pid_file = orig_locate(page_url, assets_dir, hidden)
                app_dir = os.path.dirname(args[0])
                lib_dir = os.path.join(app_dir, "lib")
                if os.path.exists(lib_dir):
                    mpv1 = os.path.join(lib_dir, "libmpv.so.1")
                    if not os.path.exists(mpv1):
                        for candidate in ["/usr/lib/x86_64-linux-gnu/libmpv.so.2", "/usr/lib/x86_64-linux-gnu/libmpv.so"]:
                            if os.path.exists(candidate):
                                try:
                                    os.symlink(candidate, mpv1)
                                except OSError:
                                    pass
                                break
                    current_ld = flet_env.get("LD_LIBRARY_PATH", "")
                    if lib_dir not in current_ld.split(":"):
                        flet_env["LD_LIBRARY_PATH"] = f"{lib_dir}:{current_ld}".strip(":")
                return args, flet_env, pid_file
            flet_desktop.__locate_and_unpack_flet_view = _patched_locate
        except Exception:
            pass
        return
    try:
        import asyncio
        import ctypes
        from ctypes import wintypes
        import flet_desktop

        orig_locate = flet_desktop.__locate_and_unpack_flet_view

        class STARTUPINFOW(ctypes.Structure):
            _fields_ = [
                ("cb", wintypes.DWORD),
                ("lpReserved", wintypes.LPWSTR),
                ("lpDesktop", wintypes.LPWSTR),
                ("lpTitle", wintypes.LPWSTR),
                ("dwX", wintypes.DWORD),
                ("dwY", wintypes.DWORD),
                ("dwXSize", wintypes.DWORD),
                ("dwYSize", wintypes.DWORD),
                ("dwXCountChars", wintypes.DWORD),
                ("dwYCountChars", wintypes.DWORD),
                ("dwFillAttribute", wintypes.DWORD),
                ("dwFlags", wintypes.DWORD),
                ("wShowWindow", wintypes.WORD),
                ("cbReserved2", wintypes.WORD),
                ("lpReserved2", ctypes.POINTER(wintypes.BYTE)),
                ("hStdInput", wintypes.HANDLE),
                ("hStdOutput", wintypes.HANDLE),
                ("hStdError", wintypes.HANDLE),
            ]

        class PROCESS_INFORMATION(ctypes.Structure):
            _fields_ = [
                ("hProcess", wintypes.HANDLE),
                ("hThread", wintypes.HANDLE),
                ("dwProcessId", wintypes.DWORD),
                ("dwThreadId", wintypes.DWORD),
            ]

        class _WinProcWrapper:
            def __init__(self, hproc, pid):
                self._hproc = hproc
                self.pid = pid
                self.returncode = None

            async def wait(self):
                kernel32 = ctypes.windll.kernel32
                while True:
                    res = kernel32.WaitForSingleObject(self._hproc, 250)
                    if res == 0:
                        code = wintypes.DWORD()
                        kernel32.GetExitCodeProcess(self._hproc, ctypes.byref(code))
                        self.returncode = code.value
                        return self.returncode
                    await asyncio.sleep(0.25)

        async def _custom_open_flet_view_async(page_url, assets_dir, hidden):
            args, flet_env, pid_file = orig_locate(page_url, assets_dir, hidden)
            si = STARTUPINFOW()
            si.cb = ctypes.sizeof(si)
            si.lpDesktop = "WinSta0\\Default"
            si.dwFlags = 1
            si.wShowWindow = 1
            pi = PROCESS_INFORMATION()

            cmd_str = " ".join(f'"{a}"' for a in args)
            cmd_buf = ctypes.create_unicode_buffer(cmd_str)
            cwd_str = os.path.dirname(args[0])
            kernel32 = ctypes.windll.kernel32
            ok = kernel32.CreateProcessW(
                None, cmd_buf, None, None, False, 0, None, cwd_str, ctypes.byref(si), ctypes.byref(pi)
            )
            if ok:
                return _WinProcWrapper(pi.hProcess, pi.dwProcessId), pid_file
            return (
                await asyncio.create_subprocess_exec(args[0], *args[1:], env=flet_env),
                pid_file,
            )

        flet_desktop.open_flet_view_async = _custom_open_flet_view_async
    except Exception as ex:
        print(f"Aviso parche escritorio Flet: {ex}")


if __name__ == "__main__":
    _parchar_flet_escritorio_interactivo()
    try:
        assets_dir = os.path.join(BASE_DIR, "src")
        if "--web" in sys.argv or os.environ.get("FLET_VIEW") == "web":
            ft.app(target=main, view=ft.AppView.WEB_BROWSER, assets_dir=assets_dir)
        else:
            ft.app(target=main, assets_dir=assets_dir)
    except Exception as e:
        print(f"Error al iniciar Flet: {e}")
