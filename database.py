import sqlite3
import os
import sys


def obtener_ruta_base():
    """Retorna el directorio base correcto ya sea en script .py o empaquetado (.exe / binario)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = obtener_ruta_base()
DB_PATH = os.path.join(BASE_DIR, "universidad_horarios.db")

DIAS_ORDEN = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
BLOQUES_HORARIOS = [
    ("07:00", "09:00"),
    ("09:00", "11:00"),
    ("11:00", "13:00"),
    ("14:00", "16:00"),
    ("16:00", "18:00"),
    ("18:00", "20:00"),
]


def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.execute("PRAGMA journal_mode = WAL;")
    return conn


def init_db():
    conn = get_connection()
    cursor = conn.cursor()

    cursor.executescript("""
    CREATE TABLE IF NOT EXISTS facultades (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL UNIQUE
    );

    CREATE TABLE IF NOT EXISTS carreras (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL,
        facultad_id INTEGER NOT NULL,
        UNIQUE(nombre, facultad_id),
        FOREIGN KEY (facultad_id) REFERENCES facultades(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS salones (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        codigo TEXT NOT NULL UNIQUE,
        capacidad INTEGER DEFAULT 40,
        edificio TEXT DEFAULT 'Edificio Principal'
    );

    CREATE TABLE IF NOT EXISTS profesores (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL UNIQUE,
        cedula TEXT DEFAULT '',
        email TEXT DEFAULT ''
    );

    CREATE TABLE IF NOT EXISTS disponibilidades (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        profesor_id INTEGER NOT NULL,
        dia TEXT NOT NULL,
        hora_inicio TEXT NOT NULL,
        hora_fin TEXT NOT NULL,
        UNIQUE(profesor_id, dia, hora_inicio, hora_fin),
        FOREIGN KEY (profesor_id) REFERENCES profesores(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS materias (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL,
        codigo TEXT DEFAULT '',
        carrera_id INTEGER NOT NULL,
        semestre TEXT DEFAULT '1er Semestre',
        UNIQUE(nombre, carrera_id),
        FOREIGN KEY (carrera_id) REFERENCES carreras(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS carga_academica (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        profesor_id INTEGER NOT NULL,
        materia_id INTEGER NOT NULL,
        seccion TEXT DEFAULT 'Sección A',
        bloques_semanales INTEGER DEFAULT 1,
        salon_preferido TEXT DEFAULT '',
        dia_fijo TEXT DEFAULT '',
        hora_inicio_fija TEXT DEFAULT '',
        hora_fin_fija TEXT DEFAULT '',
        UNIQUE(profesor_id, materia_id, seccion),
        FOREIGN KEY (profesor_id) REFERENCES profesores(id) ON DELETE CASCADE,
        FOREIGN KEY (materia_id) REFERENCES materias(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS asignaciones_horario (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        carga_id INTEGER NOT NULL,
        profesor_id INTEGER NOT NULL,
        materia_id INTEGER NOT NULL,
        salon_id INTEGER NOT NULL,
        dia TEXT NOT NULL,
        hora_inicio TEXT NOT NULL,
        hora_fin TEXT NOT NULL,
        estado TEXT DEFAULT 'Confirmado',
        FOREIGN KEY (carga_id) REFERENCES carga_academica(id) ON DELETE CASCADE,
        FOREIGN KEY (profesor_id) REFERENCES profesores(id) ON DELETE CASCADE,
        FOREIGN KEY (materia_id) REFERENCES materias(id) ON DELETE CASCADE,
        FOREIGN KEY (salon_id) REFERENCES salones(id) ON DELETE CASCADE
    );

    CREATE INDEX IF NOT EXISTS idx_asignaciones_dia ON asignaciones_horario(dia);
    CREATE INDEX IF NOT EXISTS idx_asignaciones_profesor ON asignaciones_horario(profesor_id);
    CREATE INDEX IF NOT EXISTS idx_asignaciones_salon ON asignaciones_horario(salon_id);
    CREATE INDEX IF NOT EXISTS idx_asignaciones_materia ON asignaciones_horario(materia_id);
    CREATE INDEX IF NOT EXISTS idx_carga_profesor ON carga_academica(profesor_id);
    CREATE INDEX IF NOT EXISTS idx_materias_carrera ON materias(carrera_id);
    """);

    try:
        cursor.execute("ALTER TABLE profesores ADD COLUMN condicion TEXT DEFAULT 'Contratado/a'")
    except Exception:
        pass

    # Insertar salones predeterminados de la UNESR si no hay ninguno
    cursor.execute("SELECT COUNT(*) as total FROM salones")
    if cursor.fetchone()["total"] == 0:
        salones_default = [
            (f"P{p}-{n:02d}", 45, f"Piso {p}")
            for p in range(1, 5) for n in range(1, 8)
        ] + [
            ("VIRTUAL-01", 100, "Campus Virtual"),
            ("VIRTUAL-02", 100, "Campus Virtual"),
        ]
        cursor.executemany(
            "INSERT INTO salones (codigo, capacidad, edificio) VALUES (?, ?, ?)",
            salones_default
        )

    conn.commit()
    conn.close()


def clear_all_data(keep_salones=True):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM asignaciones_horario;")
    cursor.execute("DELETE FROM carga_academica;")
    cursor.execute("DELETE FROM disponibilidades;")
    cursor.execute("DELETE FROM materias;")
    cursor.execute("DELETE FROM profesores;")
    cursor.execute("DELETE FROM carreras;")
    cursor.execute("DELETE FROM facultades;")
    if not keep_salones:
        cursor.execute("DELETE FROM salones;")
    conn.commit()
    conn.close()


def get_or_create_facultad(cursor, nombre):
    nombre = nombre.strip() or "Facultad General"
    cursor.execute("SELECT id FROM facultades WHERE LOWER(nombre) = LOWER(?)", (nombre,))
    row = cursor.fetchone()
    if row:
        return row["id"]
    cursor.execute("INSERT INTO facultades (nombre) VALUES (?)", (nombre,))
    return cursor.lastrowid


def get_or_create_carrera(cursor, nombre, facultad_id):
    nombre = nombre.strip() or "Carrera General"
    cursor.execute(
        "SELECT id FROM carreras WHERE LOWER(nombre) = LOWER(?) AND facultad_id = ?",
        (nombre, facultad_id)
    )
    row = cursor.fetchone()
    if row:
        return row["id"]
    cursor.execute(
        "INSERT INTO carreras (nombre, facultad_id) VALUES (?, ?)",
        (nombre, facultad_id)
    )
    return cursor.lastrowid


def get_or_create_profesor(cursor, nombre, cedula="", email="", condicion="Contratado/a"):
    nombre = nombre.strip()
    cursor.execute("SELECT id FROM profesores WHERE LOWER(nombre) = LOWER(?)", (nombre,))
    row = cursor.fetchone()
    if row:
        if cedula or email or condicion:
            cursor.execute(
                """UPDATE profesores 
                   SET cedula = COALESCE(NULLIF(?, ''), cedula), 
                       email = COALESCE(NULLIF(?, ''), email),
                       condicion = COALESCE(NULLIF(?, ''), condicion)
                   WHERE id = ?""",
                (cedula, email, condicion, row["id"])
            )
        return row["id"]
    cursor.execute(
        "INSERT INTO profesores (nombre, cedula, email, condicion) VALUES (?, ?, ?, ?)",
        (nombre, cedula, email, condicion)
    )
    return cursor.lastrowid


def get_or_create_materia(cursor, nombre, carrera_id, codigo="", semestre="1er Semestre"):
    nombre = nombre.strip()
    cursor.execute(
        "SELECT id FROM materias WHERE LOWER(nombre) = LOWER(?) AND carrera_id = ?",
        (nombre, carrera_id)
    )
    row = cursor.fetchone()
    if row:
        return row["id"]
    cursor.execute(
        "INSERT INTO materias (nombre, codigo, carrera_id, semestre) VALUES (?, ?, ?, ?)",
        (nombre, codigo, carrera_id, semestre)
    )
    return cursor.lastrowid


def get_or_create_salon(cursor, codigo, capacidad=40, edificio="Edificio Principal"):
    codigo = codigo.strip()
    cursor.execute("SELECT id FROM salones WHERE LOWER(codigo) = LOWER(?)", (codigo,))
    row = cursor.fetchone()
    if row:
        return row["id"]
    cursor.execute(
        "INSERT INTO salones (codigo, capacidad, edificio) VALUES (?, ?, ?)",
        (codigo, capacidad, edificio)
    )
    return cursor.lastrowid


def actualizar_ambiente_por_docente_materia(profesor_nombre: str, materia_nombre: str, dia: str, nuevo_codigo_salon: str) -> bool:
    """Actualiza el ambiente (salón) de una clase asignada a un profesor específico de forma robusta e insensible a acentos."""
    import unicodedata, re

    def _norm(s):
        if not s:
            return ""
        s = unicodedata.normalize("NFD", str(s))
        s = "".join(c for c in s if unicodedata.category(c) != "Mn")
        return re.sub(r"[^\w\s]", "", s).strip().lower()

    conn = get_connection()
    cursor = conn.cursor()
    salon_id = get_or_create_salon(cursor, nuevo_codigo_salon.strip().upper())

    cursor.execute("""
        SELECT a.id, a.dia, p.nombre AS prof_nom, m.nombre AS mat_nom, a.carga_id
        FROM asignaciones_horario a
        JOIN profesores p ON a.profesor_id = p.id
        JOIN materias m ON a.materia_id = m.id
    """)
    todas = [dict(r) for r in cursor.fetchall()]
    target_prof = _norm(profesor_nombre)
    target_mat = _norm(materia_nombre)
    target_dia = _norm(dia)

    ids_a_actualizar = []
    cargas_a_actualizar = []
    for fila in todas:
        f_prof = _norm(fila["prof_nom"])
        f_mat = _norm(fila["mat_nom"])
        f_dia = _norm(fila["dia"])

        prof_ok = (target_prof in f_prof or f_prof in target_prof) if target_prof else True
        mat_ok = (target_mat in f_mat or f_mat in target_mat) if target_mat else True
        dia_ok = (target_dia == f_dia) if target_dia else True

        if prof_ok and mat_ok and dia_ok:
            ids_a_actualizar.append(fila["id"])
            if fila.get("carga_id"):
                cargas_a_actualizar.append(fila["carga_id"])

    # Si no hubo coincidencia exacta de día, buscar solo por profesor y materia
    if not ids_a_actualizar and target_prof and target_mat:
        for fila in todas:
            f_prof = _norm(fila["prof_nom"])
            f_mat = _norm(fila["mat_nom"])
            if (target_prof in f_prof or f_prof in target_prof) and (target_mat in f_mat or f_mat in target_mat):
                ids_a_actualizar.append(fila["id"])
                if fila.get("carga_id"):
                    cargas_a_actualizar.append(fila["carga_id"])

    if ids_a_actualizar:
        placeholders = ",".join(["?"] * len(ids_a_actualizar))
        cursor.execute(f"""
            UPDATE asignaciones_horario
            SET salon_id = ?
            WHERE id IN ({placeholders})
        """, [salon_id] + ids_a_actualizar)

    if cargas_a_actualizar:
        placeholders = ",".join(["?"] * len(cargas_a_actualizar))
        cursor.execute(f"""
            UPDATE carga_academica
            SET salon_preferido = ?
            WHERE id IN ({placeholders})
        """, [nuevo_codigo_salon.strip().upper()] + cargas_a_actualizar)

    conn.commit()
    conn.close()
    return len(ids_a_actualizar) > 0


def actualizar_ambiente_profesor(profesor_nombre: str, nuevo_codigo_salon: str) -> int:
    """Actualiza el ambiente de TODAS las clases asignadas a un profesor específico."""
    import unicodedata, re

    def _norm(s):
        if not s:
            return ""
        s = unicodedata.normalize("NFD", str(s))
        s = "".join(c for c in s if unicodedata.category(c) != "Mn")
        return re.sub(r"[^\w\s]", "", s).strip().lower()

    conn = get_connection()
    cursor = conn.cursor()
    salon_id = get_or_create_salon(cursor, nuevo_codigo_salon.strip().upper())

    cursor.execute("""
        SELECT a.id, p.nombre AS prof_nom, a.carga_id
        FROM asignaciones_horario a
        JOIN profesores p ON a.profesor_id = p.id
    """)
    todas = [dict(r) for r in cursor.fetchall()]
    target_prof = _norm(profesor_nombre)

    ids = [fila["id"] for fila in todas if (target_prof in _norm(fila["prof_nom"]) or _norm(fila["prof_nom"]) in target_prof)]
    cargas = [fila["carga_id"] for fila in todas if (target_prof in _norm(fila["prof_nom"]) or _norm(fila["prof_nom"]) in target_prof) and fila.get("carga_id")]

    if ids:
        placeholders = ",".join(["?"] * len(ids))
        cursor.execute(f"UPDATE asignaciones_horario SET salon_id = ? WHERE id IN ({placeholders})", [salon_id] + ids)
    if cargas:
        placeholders = ",".join(["?"] * len(cargas))
        cursor.execute(f"UPDATE carga_academica SET salon_preferido = ? WHERE id IN ({placeholders})", [nuevo_codigo_salon.strip().upper()] + cargas)

    conn.commit()
    conn.close()
    return len(ids)


def obtener_info_profesores_dict() -> dict:
    """Retorna un diccionario {nombre_profesor: {'cedula': ..., 'condicion': ..., 'email': ...}}."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT nombre, cedula, email, COALESCE(condicion, 'Contratado/a') AS condicion FROM profesores")
    res = {}
    for r in cursor.fetchall():
        res[r["nombre"].strip().upper()] = dict(r)
    conn.close()
    return res


def marcar_iniciar_en_blanco(activar: bool = True):
    """Marca o desmarca si el sistema debe iniciar en blanco (sin datos precargados)."""
    flag = os.path.join(BASE_DIR, ".iniciar_en_blanco")
    if activar:
        with open(flag, "w", encoding="utf-8") as f:
            f.write("1")
    else:
        if os.path.exists(flag):
            try:
                os.remove(flag)
            except Exception:
                pass


def esta_en_modo_blanco() -> bool:
    """Verifica si el sistema tiene activa la bandera de iniciar en blanco."""
    flag = os.path.join(BASE_DIR, ".iniciar_en_blanco")
    return os.path.exists(flag)

