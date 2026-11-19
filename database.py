import sqlite3
import os
import sys
import shutil
import re
import unicodedata

try:
    import psycopg
    from psycopg.rows import dict_row
except ImportError:
    psycopg = None
    dict_row = None


DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()


class PostgresRow(dict):
    def __init__(self, values):
        super().__init__(values)
        self._columns = tuple(values)

    def __getitem__(self, key):
        if isinstance(key, int):
            key = self._columns[key]
        return super().__getitem__(key)


def _postgres_sql(sql):
    insert_ignore = bool(
        re.match(r"^\s*INSERT\s+OR\s+IGNORE\s+INTO\b", sql, flags=re.IGNORECASE)
    )
    sql = re.sub(r"\bINTEGER\s+PRIMARY\s+KEY\s+AUTOINCREMENT\b", "SERIAL PRIMARY KEY", sql, flags=re.IGNORECASE)
    sql = re.sub(r"\s+COLLATE\s+NOCASE\b", "", sql, flags=re.IGNORECASE)
    sql = re.sub(r"\?", "%s", sql)
    sql = re.sub(
        r"^\s*INSERT\s+OR\s+IGNORE\s+INTO\b",
        "INSERT INTO",
        sql,
        count=1,
        flags=re.IGNORECASE,
    )
    if insert_ignore:
        sql = re.sub(r";?\s*$", " ON CONFLICT DO NOTHING", sql)
    sql = re.sub(
        r"^(ALTER\s+TABLE\s+\S+\s+ADD\s+COLUMN)\s+",
        r"\1 IF NOT EXISTS ",
        sql,
        flags=re.IGNORECASE,
    )
    return sql


class PostgresCursor:
    def __init__(self, cursor):
        self._cursor = cursor

    def execute(self, sql, params=()):
        self._cursor.execute(_postgres_sql(sql), params)
        return self

    def executemany(self, sql, seq_of_params):
        self._cursor.executemany(_postgres_sql(sql), seq_of_params)
        return self

    def executescript(self, script):
        for statement in script.split(";"):
            if statement.strip():
                self.execute(statement)
        return self

    def fetchone(self):
        row = self._cursor.fetchone()
        return PostgresRow(row) if row is not None else None

    def fetchall(self):
        return [PostgresRow(row) for row in self._cursor.fetchall()]

    @property
    def rowcount(self):
        return self._cursor.rowcount


class PostgresConnection:
    def __init__(self, dsn):
        if psycopg is None:
            raise RuntimeError(
                "PostgreSQL fue configurado pero falta psycopg. "
                "Instale las dependencias de requirements-server.txt."
            )
        self._connection = psycopg.connect(dsn, row_factory=dict_row)

    def cursor(self):
        return PostgresCursor(self._connection.cursor())

    def commit(self):
        self._connection.commit()

    def rollback(self):
        self._connection.rollback()

    def close(self):
        self._connection.close()


def _normalizar_espacios(valor):
    try:
        texto = unicodedata.normalize("NFKC", str(valor or ""))
    except Exception:
        texto = str(valor or "")
    return re.sub(r"\s+", " ", texto).strip()


def validar_profesor_basico(nombre, cedula="", email="", requerir_cedula=False):
    """Valida los datos básicos de un docente y retorna una lista de errores."""
    errores = []
    nombre_norm = _normalizar_espacios(nombre).upper()
    cedula_norm = _normalizar_espacios(cedula).upper().replace(" ", "").replace(".", "")
    email_norm = _normalizar_espacios(email).lower()

    if not nombre_norm:
        errores.append("Indique el nombre y apellido del docente.")
    else:
        caracteres_validos = " -'’.,"
        if any(not (ch.isalpha() or ch in caracteres_validos) for ch in nombre_norm):
            errores.append("El nombre del docente solo admite letras y espacios.")

    if not cedula_norm:
        if requerir_cedula:
            errores.append("Indique la cédula del docente.")
    elif not re.fullmatch(r"(V|E)?-?\d{5,11}", cedula_norm):
        errores.append("La cédula debe contener entre 5 y 11 dígitos (por ejemplo: 12345678 o V-12345678).")

    if email_norm and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email_norm):
        errores.append("El correo electrónico no tiene un formato válido.")

    return errores


def obtener_ruta_base():
    """Retorna una carpeta escribible para los datos persistentes de la aplicación."""
    if not getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(__file__))

    if os.name == "nt":
        data_root = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    else:
        data_root = os.environ.get(
            "XDG_DATA_HOME",
            os.path.join(os.path.expanduser("~"), ".local", "share"),
        )
    data_dir = os.path.join(data_root, "UNESR", "Horarios")
    os.makedirs(data_dir, exist_ok=True)

    legacy_db = os.path.join(os.path.dirname(sys.executable), "universidad_horarios.db")
    user_db = os.path.join(data_dir, "universidad_horarios.db")
    if not os.path.exists(user_db) and os.path.isfile(legacy_db):
        shutil.copy2(legacy_db, user_db)
    return data_dir


BASE_DIR = obtener_ruta_base()
DB_PATH = os.environ.get("DATABASE_PATH", os.path.join(BASE_DIR, "universidad_horarios.db"))

DIAS_ORDEN = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
BLOQUES_HORARIOS = [
    ("07:00", "09:00"),
    ("09:00", "11:00"),
    ("11:00", "13:00"),
    ("14:00", "16:00"),
    ("16:00", "18:00"),
    ("18:00", "20:00"),
]

# Inventario oficial de ambientes por piso. La numeración de los salones es
# global y consecutiva: P1-01..P1-08, P2-09..P2-22, P3-23..P3-27 y P4-28..P4-37.
PISOS_SALONES = {1: 8, 2: 14, 3: 5, 4: 10}
CAPACIDAD_DEFECTO = 45
AMBIENTES_ESPECIALES = [
    ("SATELITAL-01", 45, "Satelital 1"),
    ("SATELITAL-02", 45, "Satelital 2"),
    ("VIRTUAL-01", 100, "Campus Virtual"),
    ("VIRTUAL-02", 100, "Campus Virtual"),
]
SALONES_VERSION = "2"
FACULTAD_PNF = "Programas Nacionales de Formación (PNF)"
CARRERAS_PNF = (
    "PNF Fisioterapia",
    "PNF Veterinaria",
    "PNF Agroecología",
)

# Catálogo oficial del Programa Nacional de Formación en Fisioterapia (PNFF),
# tomado de la "Malla Curricular - Programa Nacional de Formación en
# Fisioterapia (PNFF).pdf". Cada registro es:
#   (nombre de la unidad curricular / materia, código, trayecto-tramo, U.C.)
# El campo "trayecto-tramo" se guarda en materias.semestre porque es el que
# alimenta el selector de "Ciclo" en la carga académica (web y escritorio).
MATERIAS_FISIOTERAPIA = (
    # Trayecto Inicial (Ciclo Introductorio - 12 semanas)
    ("Proyecto Nacional y Nueva Ciudadanía", "PNYNC60T12", "Trayecto Inicial (Ciclo Introductorio)", 2),
    ("Salud Pública y Fisioterapia", "SPF60T12", "Trayecto Inicial (Ciclo Introductorio)", 2),
    ("Lectura y Comprensión", "LYC60T12", "Trayecto Inicial (Ciclo Introductorio)", 2),
    ("Matemática, Química y Física", "MQF96T13", "Trayecto Inicial (Ciclo Introductorio)", 3),
    # Primer Trayecto (Año 1)
    ("Sistemas Óseo y Articular", "SOA114", "Primer Trayecto - Tramo I", 4),
    ("Fundamentos de Fisioterapia", "FF112", "Primer Trayecto - Tramo I", 2),
    ("Diversidad Cultural y Nueva Ciudadanía", "DCNC111", "Primer Trayecto - Tramo I", 1),
    ("Bases del Conocimiento Científico I", "BCCI112", "Primer Trayecto - Tramo I", 2),
    ("Sistema Muscular", "SM124", "Primer Trayecto - Tramo II", 4),
    ("Principios Físicos Aplicados a la Fisioterapia", "PFAF122", "Primer Trayecto - Tramo II", 2),
    ("Herramientas Tecnológicas I", "HT122", "Primer Trayecto - Tramo II", 2),
    ("Bases del Conocimiento Científico II", "BCCII122", "Primer Trayecto - Tramo II", 2),
    ("Sistema Nervioso", "SN134", "Primer Trayecto - Tramo III", 4),
    ("Análisis del Movimiento Humano", "AMH134", "Primer Trayecto - Tramo III", 4),
    ("Medios Físicos I", "MFI133", "Primer Trayecto - Tramo III", 3),
    ("Valoración y Diagnóstico en Fisioterapia I", "VDFI133", "Primer Trayecto - Tramo III", 3),
    ("Atención Comunitaria para Personas con Discapacidad I", "ACPDI132", "Primer Trayecto - Tramo III", 2),
    # En el PDF: código VSP132 (en tabla: BSP133) y U.C. 2 (3 en tabla).
    ("Valoración de la Salud Pública", "VSP132", "Primer Trayecto - Tramo III", 2),
    # Segundo Trayecto (Año 2)
    ("Alteraciones de los Sistemas Óseo, Articular y Muscular I", "ASOAMI212", "Segundo Trayecto - Tramo I", 2),
    ("Medios Físicos II", "MFII213", "Segundo Trayecto - Tramo I", 3),
    ("Valoración y Diagnóstico en Fisioterapia II", "VDFII213", "Segundo Trayecto - Tramo I", 3),
    ("Atención Comunitaria a las Personas con Discapacidad II", "ACPDII212", "Segundo Trayecto - Tramo I", 2),
    ("Fundamentos Estadísticos y Epidemiológicos en Salud I", "FEESI213", "Segundo Trayecto - Tramo I", 3),
    ("Alteraciones de los Sistemas Óseo, Articular y Muscular II", "ASOAMII222", "Segundo Trayecto - Tramo II", 2),
    ("Técnicas de Intervención en Fisioterapia I", "TIFI223", "Segundo Trayecto - Tramo II", 3),
    ("Inglés Instrumental", "II221", "Segundo Trayecto - Tramo II", 1),
    ("Recreación, Educación Física y Deporte en Fisioterapia", "REFDF221", "Segundo Trayecto - Tramo II", 1),
    ("Fundamentos Estadísticos y Epidemiológicos en Salud II", "FEESII223", "Segundo Trayecto - Tramo II", 3),
    ("Alteraciones del Sistema Nervioso I", "ASN232", "Segundo Trayecto - Tramo III", 2),
    ("Técnicas de Intervención en Fisioterapia II", "TIFII233", "Segundo Trayecto - Tramo III", 3),
    ("Prácticas Profesionales I", "PPI2313", "Segundo Trayecto - Tramo III", 13),
    ("Herramientas Tecnológicas II", "HTII232", "Segundo Trayecto - Tramo III", 2),
    ("Fundamentos Estadísticos y Epidemiológicos en Salud III", "FEESIII233", "Segundo Trayecto - Tramo III", 3),
    # Tercer Trayecto (Año 3 - Salida T.S.U. en Fisioterapia)
    ("Alteraciones del Sistema Nervioso II", "ASNII312", "Tercer Trayecto - Tramo I", 2),
    ("Técnicas de Intervención en Fisioterapia III", "TIFIII313", "Tercer Trayecto - Tramo I", 3),
    ("Prácticas Profesionales II", "PPII3113", "Tercer Trayecto - Tramo I", 13),
    ("Políticas Públicas en Salud y Marco Normativo en el Área de la Discapacidad", "PPSMNAD312", "Tercer Trayecto - Tramo I", 2),
    ("Diagnóstico Comunitario y Discapacidad", "DCD313", "Tercer Trayecto - Tramo I", 3),
    ("Sistema Cardiovascular y Respiratorio", "SCR322", "Tercer Trayecto - Tramo II", 2),
    ("Fisiología del Ejercicio", "FE323", "Tercer Trayecto - Tramo II", 3),
    ("Valoración y Diagnóstico en Fisioterapia III", "VDFIII322", "Tercer Trayecto - Tramo II", 2),
    ("Técnicas de Intervención en Fisioterapia IV", "TIFIV323", "Tercer Trayecto - Tramo II", 3),
    ("Psicología del Desarrollo Humano", "PDH322", "Tercer Trayecto - Tramo II", 2),
    ("Planificación Comunitaria y Discapacidad", "PCD323", "Tercer Trayecto - Tramo II", 3),
    ("Alteraciones Cardio-Respiratorias", "ACR332", "Tercer Trayecto - Tramo III", 2),
    ("Técnicas de Intervención en Fisioterapia V", "TIFV332", "Tercer Trayecto - Tramo III", 2),
    ("Prácticas Profesionales III", "PPIII3313", "Tercer Trayecto - Tramo III", 13),
    ("Psicología Social", "PS332", "Tercer Trayecto - Tramo III", 2),
    ("Intervención Comunitaria y Discapacidad", "ICD333", "Tercer Trayecto - Tramo III", 3),
    # Cuarto Trayecto (Año 4 - Salida Licenciado/a en Fisioterapia)
    ("Diagnóstico en Fisioterapia", "DF414", "Cuarto Trayecto - Tramo I", 4),
    ("Sistema Tegumentario y Linfático", "STL412", "Cuarto Trayecto - Tramo I", 2),
    ("Intervención Especializada I", "IEI415", "Cuarto Trayecto - Tramo I", 5),
    ("Ambiente y Salud", "AS411", "Cuarto Trayecto - Tramo I", 1),
    ("Gerencia de la Información en los Sistemas de Salud", "GISS414", "Cuarto Trayecto - Tramo I", 4),
    ("Esfera Urogenital", "EU421", "Cuarto Trayecto - Tramo II", 1),
    ("Sistema Estomatognático", "SE421", "Cuarto Trayecto - Tramo II", 1),
    # En el PDF: código IEII425 (en ficha: IEI415).
    ("Intervención Especializada II", "IEII425", "Cuarto Trayecto - Tramo II", 5),
    ("Accesibilidad", "A422", "Cuarto Trayecto - Tramo II", 2),
    ("Protección, Bioseguridad y Seguridad Social", "PBSS422", "Cuarto Trayecto - Tramo II", 2),
    # En el PDF: código EISD424 (en ficha: GISS414).
    ("Evaluación de Impacto en Salud y Discapacidad", "EISD424", "Cuarto Trayecto - Tramo II", 4),
    ("Fisioterapia Aplicada al Deporte", "FAD433", "Cuarto Trayecto - Tramo III", 3),
    ("Intervención Especializada III", "IEIII435", "Cuarto Trayecto - Tramo III", 5),
    ("Consolidación y Presentación del Proyecto", "CPP434", "Cuarto Trayecto - Tramo III", 4),
)

MATERIAS_FISIOTERAPIA_VERSION = "1"

# Catálogo oficial del PNF en Medicina Veterinaria, tomado de la
# "Malla_Curricular_-_PNF_en_Medicina_Veterinaria.pdf". Cada registro es:
#   (unidad curricular, código, trayecto-tramo, U.C.)
# Notas sobre el PDF: los proyectos I, II, III y V aparecen repetidos en
# ambos tramos de su trayecto (se registran una sola vez), el "Proyecto V"
# del Tramo II de Trayecto V (MVPV025) es una errata de MVPYV025, y las dos
# "Prácticas Profesionales" (MVPP900 y MVPP910) se numeran para respetar la
# unicidad de nombre por carrera.
MATERIAS_VETERINARIA = (
    # Trayecto Inicial
    ("Introducción a la Universidad y al PNF en Medicina Veterinaria", "MVIUP001", "Trayecto Inicial", 0),
    ("Proyecto Nacional Nueva Ciudadanía, Lectura y Comprensión", "MVPLC001", "Trayecto Inicial", 0),
    ("Ciencias Básicas y Tecnología: Matemática, Química, Biología y TIC", "MVCBT001", "Trayecto Inicial", 0),
    # Trayecto I - Tramo I
    ("Proyecto I", "MVPYI005", "Trayecto I - Tramo I", 4),
    ("Anatomía de los Animales Domésticos I", "MVAADI105", "Trayecto I - Tramo I", 5),
    ("Histología Veterinaria", "MVHV200", "Trayecto I - Tramo I", 3),
    ("Sistemas de Producción Animal I", "MVSPAI305", "Trayecto I - Tramo I", 3),
    ("Bioquímica I", "MVBI400", "Trayecto I - Tramo I", 2),
    ("Estadística Descriptiva", "MVED505", "Trayecto I - Tramo I", 2),
    ("Pastos y Forraje", "MVPF605", "Trayecto I - Tramo I", 2),
    # Trayecto I - Tramo II
    ("Anatomía de los Animales Domésticos II", "MVAADII110", "Trayecto I - Tramo II", 5),
    ("Histología y Embriología", "MVHE205", "Trayecto I - Tramo II", 3),
    ("Sistemas de Producción Animal II", "MVSPAII310", "Trayecto I - Tramo II", 3),
    ("Bioquímica II", "MVBII405", "Trayecto I - Tramo II", 2),
    ("Alimentación y Nutrición Animal Monogástricos", "MVANAM705", "Trayecto I - Tramo II", 2),
    ("Inglés Técnico I", "MVITI805", "Trayecto I - Tramo II", 1),
    ("Electiva I: Uso y Aprovechamiento de los Desechos Orgánicos", "MVELIUAD", "Trayecto I - Tramo II", 1),
    # Trayecto II - Tramo I
    ("Proyecto II", "MVPYII010", "Trayecto II - Tramo I", 4),
    ("Fisiología Animal Básica", "MVFAB215", "Trayecto II - Tramo I", 3),
    ("Manejo de Sistemas de Producción Animal I", "MVMSPI320", "Trayecto II - Tramo I", 3),
    ("Alimentación y Nutrición de Poligástricos", "MVANAP710", "Trayecto II - Tramo I", 2),
    ("Microbiología Veterinaria", "MVMV410", "Trayecto II - Tramo I", 3),
    ("Parasitología y Enfermedades Parasitarias", "MVPEP415", "Trayecto II - Tramo I", 3),
    ("Actividad Recreativa y Cultural", "MVARC125", "Trayecto II - Tramo I", 1),
    ("Electiva II: Orientación e Integración Psicosocial", "MVELIIOEP", "Trayecto II - Tramo I", 1),
    # Trayecto II - Tramo II
    ("Fisiología Animal Aplicada", "MVFAB220", "Trayecto II - Tramo II", 3),
    ("Inmunología", "MVI425", "Trayecto II - Tramo II", 2),
    ("Bioestadística", "MVB510", "Trayecto II - Tramo II", 2),
    ("Farmacología y Toxicología", "MVFT420", "Trayecto II - Tramo II", 3),
    ("Manejo de Sistemas de Producción Animal II", "MVMSPII325", "Trayecto II - Tramo II", 3),
    ("Genética", "MVG330", "Trayecto II - Tramo II", 2),
    ("Inglés Técnico II", "MVITII810", "Trayecto II - Tramo II", 1),
    # Trayecto III - Tramo I
    ("Proyecto III", "MVPYIII015", "Trayecto III - Tramo I", 4),
    ("Mejoramiento Genético", "MVMG335", "Trayecto III - Tramo I", 2),
    ("Enfermedades Infectocontagiosas", "MVEI230", "Trayecto III - Tramo I", 2),
    ("Semiología", "MVS720", "Trayecto III - Tramo I", 3),
    ("Anatomía Patológica", "MVAP120", "Trayecto III - Tramo I", 3),
    ("Laboratorio Clínico", "MVLC730", "Trayecto III - Tramo I", 3),
    ("Ética y Deontología Veterinaria", "MVEDV130", "Trayecto III - Tramo I", 1),
    ("Imagenología Animal", "MVIA725", "Trayecto III - Tramo I", 2),
    # Trayecto III - Tramo II
    ("Prácticas Profesionales I", "MVPP900", "Trayecto III - Tramo II", 5),
    ("Salud Pública", "MVSP340", "Trayecto III - Tramo II", 2),
    ("Cirugía I (Instrumentación)", "MVCI740", "Trayecto III - Tramo II", 3),
    ("Introducción a la Clínica Veterinaria", "MVICV750", "Trayecto III - Tramo II", 3),
    ("Reproducción Animal I", "MVRAI250", "Trayecto III - Tramo II", 2),
    ("Enfermedades No Infecciosas", "MVENI350", "Trayecto III - Tramo II", 2),
    # Trayecto IV - Tramo I
    ("Proyecto IV", "MVPYIV020", "Trayecto IV - Tramo I", 4),
    ("Epidemiología", "MVE360", "Trayecto IV - Tramo I", 2),
    ("Higiene de los Alimentos", "MVHA430", "Trayecto IV - Tramo I", 2),
    ("Clínica I", "MVCI755", "Trayecto IV - Tramo I", 3),
    ("Cirugía II", "MVCII745", "Trayecto IV - Tramo I", 3),
    ("Reproducción Animal II", "MVRAII260", "Trayecto IV - Tramo I", 2),
    ("Inglés Técnico III", "MVITIII815", "Trayecto IV - Tramo I", 1),
    ("Electiva III: Fauna Silvestre", "MVELIIIFS", "Trayecto IV - Tramo I", 1),
    # Trayecto IV - Tramo II
    ("Prevención y Riesgo en la Profesión", "MVPRP370", "Trayecto IV - Tramo II", 2),
    ("Clínica II", "MVCII760", "Trayecto IV - Tramo II", 3),
    ("Medicina de Rumiantes", "MVMR380", "Trayecto IV - Tramo II", 3),
    ("Industria de Productos Cárnicos", "MVIPC440", "Trayecto IV - Tramo II", 3),
    ("Gestión Agropecuaria", "MVGA140", "Trayecto IV - Tramo II", 2),
    ("Electiva IV: Organización Comunitaria", "MVELIVOC", "Trayecto IV - Tramo II", 1),
    # Trayecto V - Tramo I
    ("Proyecto V", "MVPYV025", "Trayecto V - Tramo I", 4),
    ("Medicina en Equinos, Caninos y otras Especies", "MVMEC385", "Trayecto V - Tramo I", 3),
    ("Legislación Veterinaria", "MVLV150", "Trayecto V - Tramo I", 2),
    ("Biotecnología en Ciencias Veterinarias", "MVBCV270", "Trayecto V - Tramo I", 3),
    ("Industria de Productos Lácteos", "MVIPL450", "Trayecto V - Tramo I", 3),
    # Trayecto V - Tramo II
    ("Prácticas Profesionales II", "MVPP910", "Trayecto V - Tramo II", 9),
    ("Medicina de Cerdos y Aves", "MVMCA390", "Trayecto V - Tramo II", 3),
    ("Terapias Alternativas en Agroecología", "MVTAA160", "Trayecto V - Tramo II", 2),
)

MATERIAS_VETERINARIA_VERSION = "1"

# Unidades curriculares del PNF en Agroecología (Plan de Estudio 2017),
# tomadas del récord académico de la UNESR Núcleo Palo Verde.
# Se excluyen las filas "(Secundaria)" porque son equivalencias de nivel
# medio (códigos con sufijo -1, -2, ...), no unidades curriculares del PNF.
# El récord muestra 45 de 46 unidades aprobadas (188 U.C.); la restante no
# aparece en el documento. "Acreditable (Deporte)" se repite en varios
# trayectos, por lo que se distingue la del Trayecto III para respetar la
# unicidad de nombre por carrera. El código PNFAGT541 aparece como
# "PNFAGT541- Inactiva" en el récord.
MATERIAS_AGROECOLOGIA = (
    # Trayecto Inicial
    ("Matemática General", "PNFAGTI03", "Trayecto Inicial", 0),
    ("Iniciación a la Institución y al PNF en Agroecología", "PNFAGTI02", "Trayecto Inicial", 0),
    ("Lenguaje y Comunicación", "PNFAGTI04", "Trayecto Inicial", 0),
    ("Proyecto Nacional y Nueva Ciudadanía", "PNFAGTI01", "Trayecto Inicial", 0),
    ("Tecnologías de Información y Comunicación - TIC", "PNFAGTI05", "Trayecto Inicial", 0),
    ("Química General", "PNFAGTI06", "Trayecto Inicial", 0),
    # Trayecto I - Tramo I
    ("Ciencias Bio-Químicas", "PNFAGT109", "Trayecto I - Tramo I", 5),
    ("Ciencias Físico-Matemáticas", "PNFAGT108", "Trayecto I - Tramo I", 5),
    ("Ciencias Biológicas", "PNFAGT110", "Trayecto I - Tramo I", 5),
    ("Ecofilosofía", "PNFAGT113", "Trayecto I - Tramo I", 4),
    ("Introducción a la Agroecología", "PNFAGT111", "Trayecto I - Tramo I", 5),
    ("Proyecto Socioproductivo Integrador I", "PNFAGT107", "Trayecto I - Tramo I", 9),
    ("Agroclimatología", "PNFAGT112", "Trayecto I - Tramo I", 4),
    # Trayecto II - Tramo I
    ("Ecofisiología", "PNFAGT216", "Trayecto II - Tramo I", 5),
    ("Acreditable (Deporte)", "PNFAGT220", "Trayecto II - Tramo I", 3),
    ("Agroecología Funcional", "PNFAGT218", "Trayecto II - Tramo I", 5),
    ("Bioestadística", "PNFAGT215", "Trayecto II - Tramo I", 5),
    ("Proyecto Socioproductivo Integrador II", "PNFAGT214", "Trayecto II - Tramo I", 9),
    ("Ecotecnologías", "PNFAGT217", "Trayecto II - Tramo I", 4),
    ("Electiva", "PNFAGT221", "Trayecto II - Tramo I", 4),
    ("Sociohistoria de la Agricultura en América", "PNFAGT219", "Trayecto II - Tramo I", 4),
    # Trayecto III - Tramo I
    ("Electiva II", "PNFAGT327", "Trayecto III - Tramo I", 4),
    ("Proyecto Socioproductivo Integrador III", "PNFAGT321", "Trayecto III - Tramo I", 9),
    ("Agroecología Animal", "PNFAGT322", "Trayecto III - Tramo I", 4),
    ("Cultivo Agroecológico", "PNFAGT323", "Trayecto III - Tramo I", 5),
    ("Geopolítica Territorial Agroecológica Intertropical", "PNFAGT324", "Trayecto III - Tramo I", 4),
    ("Planificación Agroecológica", "PNFAGT325", "Trayecto III - Tramo I", 5),
    ("Acreditable (Deporte) - Trayecto III", "PNFAGT326", "Trayecto III - Tramo I", 3),
    ("Práctica Profesional", "PNFAGT328", "Trayecto III - Tramo I", 0),
    # Trayecto IV - Tramo I
    ("Electiva III", "PNFAGT436", "Trayecto IV - Tramo I", 4),
    ("Acreditable (Deporte 4)", "PNFAGT435", "Trayecto IV - Tramo I", 3),
    ("Proyecto Socioproductivo Integrador IV", "PNFAGT429", "Trayecto IV - Tramo I", 9),
    ("Ciencias de Cuencas Hidrográficas", "PNFAGT430", "Trayecto IV - Tramo I", 7),
    ("Suelo Agroecológico", "PNFAGT431", "Trayecto IV - Tramo I", 5),
    ("Manejo de Enfermedades y Enemigos Naturales de las Plantas", "PNFAGT432", "Trayecto IV - Tramo I", 8),
    ("Ecoconstrucciones", "PNFAGT433", "Trayecto IV - Tramo I", 5),
    ("Ciencia, Tecnología y Sociedad", "PNFAGT434", "Trayecto IV - Tramo I", 4),
    # Trayecto V - Tramo I
    ("Economía Agroecológica", "PNFAGT538", "Trayecto V - Tramo I", 5),
    ("Acreditable (Deporte 5)", "PNFAGT539", "Trayecto V - Tramo I", 3),
    ("Acreditable 5", "PNFAGT540", "Trayecto V - Tramo I", 3),
    ("Mejoramiento Genético Participativo", "PNFAGT537", "Trayecto V - Tramo I", 5),
    ("Electiva IV", "PNFAGT541", "Trayecto V - Tramo I", 4),
    ("Electiva V", "PNFAGT542", "Trayecto V - Tramo I", 4),
    ("Práctica Profesional II", "PNFAGT543", "Trayecto V - Tramo I", 0),
    ("Proyecto Socioproductivo Integrador V", "PNFAGT536", "Trayecto V - Tramo I", 9),
)

MATERIAS_AGROECOLOGIA_VERSION = "1"


def generar_codigos_pisos():
    """Genera los pares (piso, código) con numeración global consecutiva."""
    resultado = []
    numero = 0
    for piso, cantidad in PISOS_SALONES.items():
        for _ in range(cantidad):
            numero += 1
            resultado.append((piso, f"P{piso}-{numero:02d}"))
    return resultado


def salones_predeterminados():
    """Retorna [(codigo, capacidad, edificio), ...] con el inventario oficial."""
    return [
        (codigo, CAPACIDAD_DEFECTO, f"Piso {piso}")
        for piso, codigo in generar_codigos_pisos()
    ] + list(AMBIENTES_ESPECIALES)


def codigos_ambientes_predeterminados():
    return [codigo for codigo, _, _ in salones_predeterminados()]


def _salones_antiguos_predeterminados():
    """Códigos del inventario anterior (7 salones por piso + virtuales)."""
    return {
        f"P{p}-{n:02d}"
        for p in range(1, 5)
        for n in range(1, 8)
    } | {"VIRTUAL-01", "VIRTUAL-02"}


def _obtener_config(cursor, clave):
    cursor.execute("SELECT valor FROM configuracion WHERE clave = ?", (clave,))
    row = cursor.fetchone()
    return row["valor"] if row else None


def _guardar_config(cursor, clave, valor):
    cursor.execute(
        "INSERT INTO configuracion (clave, valor) VALUES (?, ?) "
        "ON CONFLICT(clave) DO UPDATE SET valor = excluded.valor",
        (clave, str(valor)),
    )


def _asegurar_columna(cursor, tabla, columna, definicion):
    try:
        cursor.execute(f"ALTER TABLE {tabla} ADD COLUMN {columna} {definicion}")
    except sqlite3.OperationalError as exc:
        if "duplicate column name" not in str(exc).casefold():
            raise


def get_connection():
    if DATABASE_URL:
        if not DATABASE_URL.startswith(("postgresql://", "postgres://")):
            raise ValueError("DATABASE_URL debe usar el esquema postgresql:// o postgres://.")
        return PostgresConnection(DATABASE_URL)

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.execute("PRAGMA journal_mode = WAL;")
    return conn


def _migrar_salones(cursor):
    """Crea o actualiza el inventario de ambientes conservando la información existente."""
    if _obtener_config(cursor, "salones_version") == SALONES_VERSION:
        return

    predeterminados = salones_predeterminados()
    existentes = {
        row["codigo"].upper(): dict(row)
        for row in cursor.execute(
            "SELECT id, codigo, capacidad, edificio FROM salones"
        ).fetchall()
    }

    if not existentes:
        cursor.executemany(
            "INSERT INTO salones (codigo, capacidad, edificio) VALUES (?, ?, ?)",
            predeterminados,
        )
        _guardar_config(cursor, "salones_version", SALONES_VERSION)
        return

    antiguos = _salones_antiguos_predeterminados()

    # Verificar si alguno de los salones del inventario anterior está en uso.
    cursor.execute("SELECT COUNT(*) AS total FROM asignaciones_horario")
    hay_asignaciones = cursor.fetchone()["total"] > 0
    cursor.execute(
        "SELECT DISTINCT TRIM(salon_preferido) AS codigo FROM carga_academica "
        "WHERE salon_preferido IS NOT NULL AND TRIM(salon_preferido) <> ''"
    )
    cargas_en_uso = {
        row["codigo"].upper() for row in cursor.fetchall() if row["codigo"]
    }
    en_uso = hay_asignaciones or bool(cargas_en_uso & antiguos)

    if not en_uso:
        # Sin referencias: reemplazar el inventario viejo por el nuevo,
        # conservando la capacidad de los códigos que se mantienen iguales.
        cursor.executemany(
            "DELETE FROM salones WHERE UPPER(codigo) = ?",
            [(codigo,) for codigo in antiguos],
        )
        filas = []
        for codigo, capacidad, edificio in predeterminados:
            anterior = existentes.get(codigo)
            filas.append(
                (codigo, anterior["capacidad"] if anterior else capacidad, edificio)
            )
        cursor.executemany(
            "INSERT OR IGNORE INTO salones (codigo, capacidad, edificio) VALUES (?, ?, ?)",
            filas,
        )
    else:
        # Con referencias: renombrar los códigos antiguos a la nueva numeración
        # global (por posición dentro de cada piso) sin perder asignaciones y
        # luego completar los faltantes.
        nuevos_por_piso = {}
        for piso, codigo in generar_codigos_pisos():
            nuevos_por_piso.setdefault(piso, []).append(codigo)

        for piso in range(1, 5):
            nuevos = nuevos_por_piso.get(piso, [])
            for indice, n in enumerate(range(1, 8)):
                viejo = f"P{piso}-{n:02d}"
                if viejo not in existentes or indice >= len(nuevos):
                    continue
                nuevo = nuevos[indice]
                if viejo == nuevo or nuevo in existentes:
                    continue
                cursor.execute(
                    "UPDATE salones SET codigo = ? WHERE UPPER(codigo) = ? "
                    "AND NOT EXISTS (SELECT 1 FROM salones WHERE UPPER(codigo) = UPPER(?))",
                    (nuevo, viejo, nuevo),
                )
                if cursor.rowcount:
                    cursor.execute(
                        "UPDATE carga_academica SET salon_preferido = ? "
                        "WHERE UPPER(TRIM(salon_preferido)) = ?",
                        (nuevo, viejo),
                    )

        existentes = {
            row["codigo"].upper()
            for row in cursor.execute("SELECT codigo FROM salones").fetchall()
        }
        for codigo, capacidad, edificio in predeterminados:
            if codigo not in existentes:
                cursor.execute(
                    "INSERT INTO salones (codigo, capacidad, edificio) VALUES (?, ?, ?)",
                    (codigo, capacidad, edificio),
                )

    _guardar_config(cursor, "salones_version", SALONES_VERSION)


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
        edificio TEXT DEFAULT 'Edificio Principal',
        es_camillas INTEGER DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS profesores (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL UNIQUE,
        cedula TEXT DEFAULT '',
        email TEXT DEFAULT '',
        mencion TEXT DEFAULT ''
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
        unidad_credito INTEGER DEFAULT 0,
        mencion TEXT DEFAULT '',
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
        compartidos TEXT DEFAULT '',
        UNIQUE(profesor_id, materia_id, seccion),
        FOREIGN KEY (profesor_id) REFERENCES profesores(id) ON DELETE CASCADE,
        FOREIGN KEY (materia_id) REFERENCES materias(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS configuracion (
        clave TEXT PRIMARY KEY,
        valor TEXT
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

    _asegurar_columna(cursor, "profesores", "condicion", "TEXT DEFAULT 'Contratado/a'")
    _asegurar_columna(cursor, "salones", "es_camillas", "INTEGER DEFAULT 0")
    _asegurar_columna(cursor, "profesores", "mencion", "TEXT DEFAULT ''")
    _asegurar_columna(cursor, "materias", "unidad_credito", "INTEGER DEFAULT 0")
    _asegurar_columna(cursor, "materias", "horas_academicas", "INTEGER DEFAULT 0")
    _asegurar_columna(cursor, "materias", "mencion", "TEXT DEFAULT ''")
    _asegurar_columna(cursor, "carga_academica", "compartidos", "TEXT DEFAULT ''")

    # Crear/migrar el inventario oficial de ambientes.
    _migrar_salones(cursor)
    _asegurar_carreras_pnf(cursor)
    # Catálogos académicos de los PNF: mismo comportamiento que las demás
    # carreras al filtrar materias, códigos, ciclo y unidades de crédito.
    asegurar_materias_fisioterapia(cursor)
    asegurar_materias_veterinaria(cursor)
    asegurar_materias_agroecologia(cursor)
    _sincronizar_horas_academicas_catalogo(cursor)

    conn.commit()
    conn.close()


def clear_all_data(keep_salones=True):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM asignaciones_horario;")
    cursor.execute("DELETE FROM carga_academica;")
    cursor.execute("DELETE FROM disponibilidades;")
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
    cursor.execute(
        "INSERT INTO facultades (nombre) VALUES (?) ON CONFLICT(nombre) DO NOTHING",
        (nombre,),
    )
    cursor.execute("SELECT id FROM facultades WHERE LOWER(nombre) = LOWER(?)", (nombre,))
    return cursor.fetchone()["id"]


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
        "INSERT INTO carreras (nombre, facultad_id) VALUES (?, ?) "
        "ON CONFLICT(nombre, facultad_id) DO NOTHING",
        (nombre, facultad_id),
    )
    cursor.execute(
        "SELECT id FROM carreras WHERE LOWER(nombre) = LOWER(?) AND facultad_id = ?",
        (nombre, facultad_id),
    )
    return cursor.fetchone()["id"]


def _asegurar_carreras_pnf(cursor):
    facultad_id = get_or_create_facultad(cursor, FACULTAD_PNF)
    for nombre in CARRERAS_PNF:
        cursor.execute(
            "SELECT id FROM carreras WHERE LOWER(nombre) = LOWER(?) AND facultad_id = ?",
            (nombre, facultad_id),
        )
        if not cursor.fetchone():
            get_or_create_carrera(cursor, nombre, facultad_id)


def _asegurar_catalogo_pnf(nombre_carrera, catalogo, clave_version, version, cursor=None, forzar=False):
    """Carga/actualiza el catálogo de un PNF en su carrera correspondiente.

    Se ejecuta una sola vez por versión (controlada con la tabla
    ``configuracion``) para no recrear la carrera si el usuario la renombró
    ni sobrescribir ajustes posteriores. Con ``forzar=True`` se resincroniza
    el catálogo con el PDF de origen.
    """
    conn = None
    if cursor is None:
        conn = get_connection()
        cursor = conn.cursor()

    try:
        if not forzar and _obtener_config(cursor, clave_version) == version:
            return 0

        facultad_id = get_or_create_facultad(cursor, FACULTAD_PNF)
        carrera_id = get_or_create_carrera(cursor, nombre_carrera, facultad_id)
        insertadas = 0
        for nombre, codigo, semestre, unidad_credito in catalogo:
            cursor.execute(
                "SELECT id FROM materias WHERE LOWER(nombre) = LOWER(?) AND carrera_id = ?",
                (nombre, carrera_id),
            )
            fila = cursor.fetchone()
            if fila:
                cursor.execute(
                    "UPDATE materias SET codigo = ?, semestre = ?, unidad_credito = ? "
                    "WHERE id = ?",
                    (codigo, semestre, unidad_credito, fila["id"]),
                )
            else:
                cursor.execute(
                    "INSERT INTO materias (nombre, codigo, carrera_id, semestre, unidad_credito) "
                    "VALUES (?, ?, ?, ?, ?)",
                    (nombre, codigo, carrera_id, semestre, unidad_credito),
                )
                insertadas += 1

        _guardar_config(cursor, clave_version, version)
        if conn is not None:
            conn.commit()
        return insertadas
    finally:
        if conn is not None:
            conn.close()


def asegurar_materias_fisioterapia(cursor=None, forzar=False):
    """Catálogo del PNF en Fisioterapia (63 unidades curriculares)."""
    return _asegurar_catalogo_pnf(
        "PNF Fisioterapia", MATERIAS_FISIOTERAPIA,
        "materias_fisioterapia_version", MATERIAS_FISIOTERAPIA_VERSION,
        cursor, forzar,
    )


def asegurar_materias_veterinaria(cursor=None, forzar=False):
    """Catálogo del PNF en Medicina Veterinaria (68 unidades curriculares)."""
    return _asegurar_catalogo_pnf(
        "PNF Veterinaria", MATERIAS_VETERINARIA,
        "materias_veterinaria_version", MATERIAS_VETERINARIA_VERSION,
        cursor, forzar,
    )


def asegurar_materias_agroecologia(cursor=None, forzar=False):
    """Catálogo del PNF en Agroecología (45 unidades curriculares)."""
    return _asegurar_catalogo_pnf(
        "PNF Agroecología", MATERIAS_AGROECOLOGIA,
        "materias_agroecologia_version", MATERIAS_AGROECOLOGIA_VERSION,
        cursor, forzar,
    )


def calcular_horas_academicas(
    hora_inicio: str = "",
    hora_fin: str = "",
    unidad_credito: int = 0,
    horas_pensum: int = 0,
    bloques_semanales: int = 1,
) -> int:
    """
    Calcula las horas académicas reales (1 hora académica UNESR = 40 a 45 minutos)
    de acuerdo con la información suministrada (rango horario asignado, horas del plan
    de estudios o unidades de crédito), evitando asumir siempre 1 hora.
    """
    horas_horario = 0
    if hora_inicio and hora_fin and ":" in str(hora_inicio) and ":" in str(hora_fin):
        try:
            hi_p = str(hora_inicio).strip().split(":")
            hf_p = str(hora_fin).strip().split(":")
            ini_m = int(hi_p[0]) * 60 + int(hi_p[1][:2])
            fin_m = int(hf_p[0]) * 60 + int(hf_p[1][:2])
            if fin_m < ini_m and int(hf_p[0]) < 12:
                fin_m += 12 * 60
            dur_min = max(0, fin_m - ini_m)
            if dur_min > 0:
                horas_horario = max(1, int(round(dur_min / 42.5)))
        except Exception:
            horas_horario = 0

    try:
        hp = int(horas_pensum or 0)
    except Exception:
        hp = 0
    try:
        uc = int(unidad_credito or 0)
    except Exception:
        uc = 0
    try:
        bs = int(bloques_semanales or 0)
    except Exception:
        bs = 0

    if horas_horario > 1:
        return horas_horario
    if hp > 0:
        return hp
    if uc > 0:
        return uc
    if bs > 1:
        return bs
    if horas_horario == 1:
        return 1
    return 2


def agrupar_franjas_disponibilidad(disponibilidades):
    """
    Agrupa casillas de disponibilidad de 45 minutos contiguas o adyacentes en el mismo día
    (incluyendo recesos de hasta 15 minutos entre bloques como 09:15->09:20 o 12:15->12:30)
    en bloques continuos (dia, hora_inicio, hora_fin, num_horas, sub_franjas).
    """
    por_dia = {}
    for item in disponibilidades or []:
        if len(item) < 3:
            continue
        dia, ini, fin = str(item[0]).strip(), str(item[1]).strip(), str(item[2]).strip()
        if not dia or not ini or not fin:
            continue
        por_dia.setdefault(dia, []).append((ini, fin))

    bloques_agrupados = []
    orden_d = {d: i for i, d in enumerate(DIAS_ORDEN)}
    for dia in sorted(por_dia.keys(), key=lambda d: orden_d.get(d, 99)):
        franjas_dia = sorted(set(por_dia[dia]), key=lambda x: x[0])
        if not franjas_dia:
            continue
        actual_ini, actual_fin = franjas_dia[0]
        sub_lista = [(actual_ini, actual_fin)]

        def _a_min(h_str):
            p = h_str.split(":")
            return int(p[0]) * 60 + int(p[1][:2])

        for ini, fin in franjas_dia[1:]:
            brecha = _a_min(ini) - _a_min(actual_fin)
            if brecha <= 15:
                if _a_min(fin) > _a_min(actual_fin):
                    actual_fin = fin
                sub_lista.append((ini, fin))
            else:
                horas_b = calcular_horas_academicas(actual_ini, actual_fin, bloques_semanales=len(sub_lista))
                bloques_agrupados.append((dia, actual_ini, actual_fin, max(len(sub_lista), horas_b), list(sub_lista)))
                actual_ini, actual_fin = ini, fin
                sub_lista = [(ini, fin)]

        horas_b = calcular_horas_academicas(actual_ini, actual_fin, bloques_semanales=len(sub_lista))
        bloques_agrupados.append((dia, actual_ini, actual_fin, max(len(sub_lista), horas_b), list(sub_lista)))

    return bloques_agrupados


def clave_carrera_canonica(nombre_carrera: str) -> str:
    """Devuelve una clave canónica para emparejar nombres cortos o largos de carreras y PNF."""
    s = unicodedata.normalize("NFD", str(nombre_carrera or "")).upper()
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    if "FISIOTERAPIA" in s or "PNFF" in s:
        return "PNF_FISIOTERAPIA"
    if "VETERINARIA" in s:
        return "PNF_VETERINARIA"
    if "AGROECOLOGIA" in s:
        return "PNF_AGROECOLOGIA"
    if "INICIAL" in s or "EINI" in s or s.strip() == "EDUC_INIC":
        return "EDUC_INIC"
    if "INTEGRAL" in s or "EINT" in s or s.strip() == "EDUC_INTEG":
        return "EDUC_INTEG"
    if "MATEM" in s or "EDMT" in s or s.strip() == "EDUC_MATEM":
        return "EDUC_MATEM"
    if "INGLES" in s or "EDII" in s or s.strip() == "EDUC_INGLES":
        return "EDUC_INGLES"
    if "RECURSOS MATERIALES" in s or "FINANCIER" in s or "RMF" in s or "ARMF" in s:
        return "RMF"
    if "MERCADEO" in s or "AMER" in s or s.strip() == "MERC":
        return "MERC"
    if "INFORMATICA" in s or "AINF" in s or s.strip() == "INF":
        return "INF"
    if "ORGANIZACION" in s or "SISTEMA" in s or "OYS" in s or "AOYS" in s:
        return "ORGAN_SIS"
    if "RECURSOS HUMANOS" in s or "RRHH" in s or "ARRHH" in s:
        return "RRHH"
    return re.sub(r"[^A-Z0-9]+", "_", s).strip("_")


def _sincronizar_horas_academicas_catalogo(cursor):
    """
    Sincroniza horas_academicas y unidad_credito de todas las materias con el Plan de Estudios 2024
    y las mallas curriculares PNF para que nunca aparezcan con 1 hora por defecto.
    """
    plan_xlsx = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads", "PLAN DE ESTUDIO 2024(1).xlsx")
    if os.path.isfile(plan_xlsx) and _obtener_config(cursor, "plan_2024_horas_v1") != "1":
        try:
            import openpyxl
            wb = openpyxl.load_workbook(plan_xlsx, data_only=True)
            mapa_hoja_carrera = {
                "ARRHH_2024": ("Administración - Administración de Recursos Humanos", "Facultad de Administración"),
                "ARMF 2024":  ("Administración - Recursos Materiales y Financieros", "Facultad de Administración"),
                "AINF_2024":  ("Administración - Informática", "Facultad de Administración"),
                "AMER_2024":  ("Administración - Mercadeo", "Facultad de Administración"),
                "AOYS_2024":  ("Administración - Organización y Sistemas", "Facultad de Administración"),
                "EDMT_2024":  ("Educación - Docencia en Matemáticas", "Facultad de Educación"),
                "EDII_2024":  ("Educación - Docencia en Idioma Inglés", "Facultad de Educación"),
                "EINI_2024":  ("Educación Inicial - Sin Mención", "Facultad de Educación"),
                "EINT_2024":  ("Educación Integral - Sin Mención", "Facultad de Educación"),
            }
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                car_info = mapa_hoja_carrera.get(sheet)
                if not car_info:
                    continue
                car_nom, fac_nom = car_info
                fac_id = get_or_create_facultad(cursor, fac_nom)
                car_id = get_or_create_carrera(cursor, car_nom, fac_id)
                ciclo_actual = "INTRODUCTORIO"
                for row in ws.iter_rows(values_only=True):
                    vals = ["" if c is None else str(c).strip() for c in row]
                    linea = " ".join(v for v in vals if v)
                    m_ciclo = re.search(r"CICLO\s+(INTRODUCTORIO|GENERAL|PROFESIONAL|ELECTIVO)", linea.upper())
                    if m_ciclo:
                        ciclo_actual = m_ciclo.group(1).upper()
                        continue
                    if len(vals) >= 3 and vals[0] and vals[1] and vals[2].isdigit():
                        cod = vals[0].strip().upper()
                        nom = re.sub(r"\s+", " ", vals[1]).strip()
                        uc = int(vals[2])
                        ht = int(vals[3]) if len(vals) > 3 and vals[3].isdigit() else 0
                        hp = int(vals[4]) if len(vals) > 4 and vals[4].isdigit() else 0
                        horas_tot = (ht + hp) if (ht + hp) > 0 else max(2, uc)
                        mat_id = get_or_create_materia(cursor, nom, car_id, codigo=cod, semestre=ciclo_actual)
                        cursor.execute(
                            """UPDATE materias
                               SET codigo = COALESCE(NULLIF(?, ''), codigo),
                                   semestre = COALESCE(NULLIF(?, ''), semestre),
                                   unidad_credito = CASE WHEN ? > 0 THEN ? ELSE unidad_credito END,
                                   horas_academicas = ?
                               WHERE id = ?""",
                            (cod, ciclo_actual, uc, uc, horas_tot, mat_id),
                        )
            _guardar_config(cursor, "plan_2024_horas_v1", "1")
        except Exception as ex:
            print(f"[Aviso Plan Estudios 2024]: {ex}")

    # Asegurar que toda materia con horas_academicas = 0 tome unidad_credito (o mínimo 2 horas académicas)
    cursor.execute(
        """UPDATE materias
           SET horas_academicas = CASE
               WHEN COALESCE(unidad_credito, 0) > 1 THEN unidad_credito
               ELSE 2
           END
           WHERE COALESCE(horas_academicas, 0) <= 0"""
    )

    # Actualizar bloques_semanales en carga_academica según el horario fijo o las horas académicas de la materia
    cursor.execute(
        """SELECT ca.id, ca.dia_fijo, ca.hora_inicio_fija, ca.hora_fin_fija,
                  ca.bloques_semanales, m.unidad_credito, COALESCE(m.horas_academicas, 0) AS horas_academicas
           FROM carga_academica ca
           JOIN materias m ON m.id = ca.materia_id"""
    )
    actualizaciones_ca = []
    for r in cursor.fetchall():
        h_real = calcular_horas_academicas(
            r["hora_inicio_fija"],
            r["hora_fin_fija"],
            r["unidad_credito"],
            r["horas_academicas"],
            r["bloques_semanales"],
        )
        if h_real != r["bloques_semanales"] and (r["bloques_semanales"] or 1) <= 1:
            actualizaciones_ca.append((h_real, r["id"]))
    if actualizaciones_ca:
        cursor.executemany("UPDATE carga_academica SET bloques_semanales = ? WHERE id = ?", actualizaciones_ca)


def obtener_carreras():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        """SELECT c.id, c.nombre, c.facultad_id, f.nombre AS facultad
           FROM carreras c
           JOIN facultades f ON f.id = c.facultad_id
           ORDER BY f.nombre COLLATE NOCASE, c.nombre COLLATE NOCASE"""
    )
    carreras = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return carreras


def obtener_materias_carrera(carrera_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        """SELECT id, nombre, codigo, semestre, unidad_credito,
                  COALESCE(NULLIF(horas_academicas, 0), NULLIF(unidad_credito, 0), 2) AS horas_academicas
           FROM materias
           WHERE carrera_id = ?
           ORDER BY
             CASE
               WHEN LOWER(semestre) LIKE 'introduct%' THEN 1
               WHEN LOWER(semestre) LIKE 'trayecto inicial%' THEN 1
               WHEN LOWER(semestre) LIKE 'general%' THEN 2
               WHEN LOWER(semestre) LIKE 'elect%' THEN 3
               WHEN LOWER(semestre) LIKE 'profes%' THEN 4
               WHEN LOWER(semestre) LIKE 'primer trayecto%' OR LOWER(semestre) LIKE 'trayecto i - tramo%' THEN 5
               WHEN LOWER(semestre) LIKE 'segundo trayecto%' OR LOWER(semestre) LIKE 'trayecto ii - tramo%' THEN 6
               WHEN LOWER(semestre) LIKE 'tercer trayecto%' OR LOWER(semestre) LIKE 'trayecto iii - tramo%' THEN 7
               WHEN LOWER(semestre) LIKE 'cuarto trayecto%' OR LOWER(semestre) LIKE 'trayecto iv - tramo%' THEN 8
               WHEN LOWER(semestre) LIKE 'trayecto v - tramo%' THEN 9
               ELSE 10
             END,
             semestre COLLATE NOCASE,
             codigo COLLATE NOCASE,
             nombre COLLATE NOCASE""",
        (carrera_id,),
    )
    materias = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return materias


def obtener_materias_por_nombre_carrera(nombre_o_clave: str):
    """
    Obtiene todas las materias correspondientes a una carrera o PNF buscando por nombre exacto
    o por su clave canónica (ej. 'Recursos Humanos' -> 'Administración - Administración de Recursos Humanos').
    """
    if not str(nombre_o_clave or "").strip():
        return []
    clave_buscada = clave_carrera_canonica(nombre_o_clave)
    carreras = obtener_carreras()
    ids_coincidentes = [
        c["id"] for c in carreras
        if c["nombre"].strip().casefold() == str(nombre_o_clave).strip().casefold()
        or clave_carrera_canonica(c["nombre"]) == clave_buscada
    ]
    resultado = []
    vistos = set()
    for cid in ids_coincidentes:
        for m in obtener_materias_carrera(cid):
            norm_nom = unicodedata.normalize("NFD", m["nombre"].strip()).upper()
            norm_nom = re.sub(r"\s+", " ", "".join(ch for ch in norm_nom if unicodedata.category(ch) != "Mn"))
            if norm_nom not in vistos:
                vistos.add(norm_nom)
                resultado.append(m)
    return resultado


def obtener_disponibilidades_profesor(nombre_profesor: str):
    """Retorna la lista de tuplas (dia, hora_inicio, hora_fin) de disponibilidad de un docente."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        """SELECT d.dia, d.hora_inicio, d.hora_fin
           FROM disponibilidades d
           JOIN profesores p ON p.id = d.profesor_id
           WHERE LOWER(p.nombre) = LOWER(?)
           ORDER BY d.dia, d.hora_inicio""",
        (str(nombre_profesor or "").strip(),),
    )
    filas = [(r["dia"], r["hora_inicio"], r["hora_fin"]) for r in cursor.fetchall()]
    conn.close()
    return filas


def obtener_profesores_capacitados(carrera_id=None, materia_nombre="", busqueda=""):
    conn = get_connection()
    cursor = conn.cursor()
    query = """
        SELECT DISTINCT
            p.id,
            p.nombre,
            COALESCE(p.cedula, '') AS cedula,
            COALESCE(p.email, '') AS email,
            COALESCE(p.condicion, 'Contratado/a') AS condicion,
            c.id AS carrera_id,
            c.nombre AS carrera,
            m.nombre AS materia
        FROM carga_academica ca
        JOIN profesores p ON p.id = ca.profesor_id
        JOIN materias m ON m.id = ca.materia_id
        JOIN carreras c ON c.id = m.carrera_id
        WHERE 1=1
    """
    params = []
    if carrera_id:
        query += " AND c.id = ?"
        params.append(carrera_id)
    if str(materia_nombre or "").strip():
        query += " AND LOWER(m.nombre) = LOWER(?)"
        params.append(str(materia_nombre).strip())
    if str(busqueda or "").strip():
        termino = f"%{str(busqueda).strip().lower()}%"
        query += " AND (LOWER(p.nombre) LIKE ? OR LOWER(c.nombre) LIKE ? OR LOWER(m.nombre) LIKE ?)"
        params.extend([termino, termino, termino])
    query += " ORDER BY p.nombre COLLATE NOCASE, c.nombre COLLATE NOCASE"
    cursor.execute(query, params)
    profesores = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return profesores


def obtener_facultades():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre FROM facultades ORDER BY nombre COLLATE NOCASE")
    facultades = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return facultades


def guardar_carrera(nombre, facultad, carrera_id=None):
    nombre = _normalizar_espacios(nombre)
    facultad = _normalizar_espacios(facultad)
    if not nombre:
        raise ValueError("Indique el nombre de la carrera o PNF.")
    if not facultad:
        raise ValueError("Indique la facultad o categoría del programa.")

    conn = get_connection()
    try:
        cursor = conn.cursor()
        facultad_id = get_or_create_facultad(cursor, facultad)
        if carrera_id is None:
            cursor.execute(
                """SELECT id FROM carreras
                   WHERE LOWER(nombre) = LOWER(?) AND facultad_id = ?""",
                (nombre, facultad_id),
            )
        else:
            cursor.execute(
                """SELECT id FROM carreras
                   WHERE LOWER(nombre) = LOWER(?) AND facultad_id = ? AND id != ?""",
                (nombre, facultad_id, carrera_id),
            )
        duplicada = cursor.fetchone()
        if duplicada:
            raise ValueError("Ya existe una carrera con ese nombre en esa facultad.")

        if carrera_id is None:
            carrera_id = get_or_create_carrera(cursor, nombre, facultad_id)
        else:
            cursor.execute(
                "UPDATE carreras SET nombre = ?, facultad_id = ? WHERE id = ?",
                (nombre, facultad_id, carrera_id),
            )
            if cursor.rowcount == 0:
                raise ValueError("La carrera seleccionada ya no existe.")
        conn.commit()
        return carrera_id
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def get_or_create_profesor(cursor, nombre, cedula="", email="", condicion="Contratado/a", mencion=""):
    nombre = nombre.strip()
    mencion = str(mencion or "").strip()
    cursor.execute("SELECT id FROM profesores WHERE LOWER(nombre) = LOWER(?)", (nombre,))
    row = cursor.fetchone()
    if not row:
        cursor.execute(
            "INSERT INTO profesores (nombre, cedula, email, condicion, mencion) "
            "VALUES (?, ?, ?, ?, ?) ON CONFLICT(nombre) DO NOTHING",
            (nombre, cedula, email, condicion, mencion),
        )
        cursor.execute("SELECT id FROM profesores WHERE LOWER(nombre) = LOWER(?)", (nombre,))
        row = cursor.fetchone()
    if cedula or email or condicion or mencion:
        cursor.execute(
            """UPDATE profesores
               SET cedula = COALESCE(NULLIF(?, ''), cedula),
                   email = COALESCE(NULLIF(?, ''), email),
                   condicion = COALESCE(NULLIF(?, ''), condicion),
                   mencion = COALESCE(NULLIF(?, ''), mencion)
               WHERE id = ?""",
            (cedula, email, condicion, mencion, row["id"]),
        )
    return row["id"]


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
        "INSERT INTO materias (nombre, codigo, carrera_id, semestre) VALUES (?, ?, ?, ?) "
        "ON CONFLICT(nombre, carrera_id) DO NOTHING",
        (nombre, codigo, carrera_id, semestre),
    )
    cursor.execute(
        "SELECT id FROM materias WHERE LOWER(nombre) = LOWER(?) AND carrera_id = ?",
        (nombre, carrera_id),
    )
    return cursor.fetchone()["id"]


def get_or_create_salon(cursor, codigo, capacidad=CAPACIDAD_DEFECTO, edificio="Edificio Principal"):
    codigo = codigo.strip()
    cursor.execute("SELECT id FROM salones WHERE LOWER(codigo) = LOWER(?)", (codigo,))
    row = cursor.fetchone()
    if row:
        return row["id"]
    cursor.execute(
        "INSERT INTO salones (codigo, capacidad, edificio) VALUES (?, ?, ?) "
        "ON CONFLICT(codigo) DO NOTHING",
        (codigo, capacidad, edificio),
    )
    cursor.execute("SELECT id FROM salones WHERE LOWER(codigo) = LOWER(?)", (codigo,))
    return cursor.fetchone()["id"]


def registrar_ficha_docente(datos):
    """Registra manualmente los datos de un docente y sus materias."""
    docente = _normalizar_espacios(datos.get("docente")).upper()
    cedula = _normalizar_espacios(datos.get("cedula")).upper().replace(" ", "")
    correo = _normalizar_espacios(datos.get("correo")).lower()
    condicion = str(datos.get("condicion") or "").strip()

    disponibilidades_guardadas = datos.get("disponibilidades") or []
    disponibilidades = list(dict.fromkeys(
        tuple(bloque) for bloque in disponibilidades_guardadas
    ))
    materias_input = datos.get("materias") or []

    codigos_existentes = {
        str(codigo or "").strip().upper()
        for codigo in obtener_codigos_materia()
    }

    materias = []
    seen = set()
    for indice, materia_input in enumerate(materias_input, 1):
        if not isinstance(materia_input, dict):
            continue

        materia = _normalizar_espacios(materia_input.get("materia")).upper()
        carrera = _normalizar_espacios(materia_input.get("carrera"))
        mencion = _normalizar_espacios(materia_input.get("mencion"))
        ambiente = _normalizar_espacios(materia_input.get("ambiente")).upper()
        codigo = _normalizar_espacios(materia_input.get("codigo")).upper().replace(" ", "")
        ciclo = _normalizar_espacios(materia_input.get("ciclo")).upper()
        seccion = _normalizar_espacios(materia_input.get("seccion"))
        unidad_credito = _normalizar_espacios(materia_input.get("unidad_credito"))
        compartidos = _normalizar_espacios(materia_input.get("compartidos"))

        if not any([materia, carrera, mencion, ambiente, codigo, ciclo, seccion, unidad_credito, compartidos]):
            continue
        if not materia:
            materias.append({"_error": f"Fila {indice}: indique el nombre de la materia."})
            continue
        if not carrera:
            materias.append({"_error": f"Materia '{materia}': seleccione la carrera."})
            continue
        if codigo and codigo not in codigos_existentes and not codigo.isalnum():
            materias.append({"_error": f"Materia '{materia}': el código solo admite letras y números."})
            continue
        if seccion and not seccion.isdigit():
            materias.append({"_error": f"Materia '{materia}': la sección solo admite números."})
            continue
        if unidad_credito and not unidad_credito.isdigit():
            materias.append({"_error": f"Materia '{materia}': la unidad de crédito solo admite números."})
            continue

        key = (materia, carrera.lower(), mencion.lower(), seccion)
        if key in seen:
            materias.append({"_error": f"Materia '{materia}' sección {seccion or 'única'} duplicada."})
            continue
        seen.add(key)
        materias.append({
            "materia": materia,
            "carrera": carrera,
            "mencion": mencion,
            "ambiente": ambiente,
            "codigo": codigo,
            "ciclo": ciclo,
            "seccion": seccion,
            "unidad_credito": unidad_credito,
            "compartidos": compartidos,
        })

    errores = validar_profesor_basico(docente, cedula, correo, requerir_cedula=False)
    errores.extend(materia["_error"] for materia in materias if "_error" in materia)
    materias_validas = [m for m in materias if "_error" not in m]
    if not materias_validas:
        errores.append("Indique al menos una materia.")
    if not disponibilidades:
        errores.append("Seleccione al menos una franja de disponibilidad.")
    for bloque in disponibilidades:
        if (
            len(bloque) != 3
            or bloque[0] not in DIAS_ORDEN
            or not all(
                isinstance(hora, str)
                and len(hora) == 5
                and hora[2] == ":"
                and hora[:2].isdigit()
                and hora[3:].isdigit()
                and 0 <= int(hora[:2]) <= 23
                and 0 <= int(hora[3:]) <= 59
                for hora in bloque[1:]
            )
            or bloque[1] >= bloque[2]
        ):
            errores.append("La disponibilidad contiene una franja inválida.")
            break
    if errores:
        return {"insertados": 0, "errores": errores}

    conn = get_connection()
    try:
        cursor = conn.cursor()
        facultad_admin = get_or_create_facultad(cursor, "Facultad de Administración")
        facultad_educacion = get_or_create_facultad(cursor, "Facultad de Educación")
        mencies = {item["mencion"] for item in materias_validas if item["mencion"]}
        mencion_unica = next(iter(mencies)) if len(mencies) == 1 else ""
        profesor_id = get_or_create_profesor(
            cursor, docente, cedula, correo,
            condicion=condicion, mencion=mencion_unica,
        )

        for ambiente in {item["ambiente"] for item in materias_validas if item["ambiente"]}:
            get_or_create_salon(cursor, ambiente)

        cursor.execute("DELETE FROM disponibilidades WHERE profesor_id = ?", (profesor_id,))
        cursor.executemany(
            """INSERT OR IGNORE INTO disponibilidades (profesor_id, dia, hora_inicio, hora_fin)
               VALUES (?, ?, ?, ?)""",
            [(profesor_id, dia, inicio, fin) for dia, inicio, fin in disponibilidades],
        )

        carreras_db_todas = [
            dict(r) for r in cursor.execute("SELECT id, nombre, facultad_id FROM carreras ORDER BY id ASC").fetchall()
        ]
        carreras_cache = {}
        ids_cargas_guardadas = []
        for indice, item in enumerate(materias_validas, 1):
            mencion = item["mencion"].strip()
            if mencion.lower() in ("", "sin mencion", "sin mención"):
                mencion = ""
            if mencion and mencion.lower() in item["carrera"].lower():
                carrera_full = item["carrera"]
            else:
                carrera_full = f"{item['carrera']} - {mencion}" if mencion else item["carrera"]
            clave_can = clave_carrera_canonica(item["carrera"])
            carrera_id = carreras_cache.get(clave_can)
            if carrera_id is None:
                for c_row in carreras_db_todas:
                    if (
                        c_row["nombre"].strip().casefold() in (carrera_full.casefold(), item["carrera"].strip().casefold())
                        or clave_carrera_canonica(c_row["nombre"]) == clave_can
                    ):
                        carrera_id = c_row["id"]
                        break
            if carrera_id is None:
                carrera_normalizada = carrera_full.upper()
                pnf_like = carrera_normalizada.startswith("PNF ") or "PROGRAMA NACIONAL DE FORMACION" in carrera_normalizada
                admin_like = any(k in carrera_normalizada for k in (
                    "ADMINISTR", "RECURSOS HUMANOS", "RECURSOS MATERIALES", "FINANCIER", "MERCADEO",
                    "INFORMÁTICA", "INFORMATICA", "ORGANIZACIÓN Y SISTEMAS", "ORGANIZACION Y SISTEMAS", "OYS",
                ))
                if pnf_like:
                    facultad_id = get_or_create_facultad(cursor, FACULTAD_PNF)
                else:
                    facultad_id = facultad_admin if admin_like else facultad_educacion
                carrera_id = get_or_create_carrera(cursor, carrera_full, facultad_id)
            carreras_cache[clave_can] = carrera_id
            codigo_materia = item["codigo"] or f"MAT-{indice:03d}"
            materia_id = get_or_create_materia(
                cursor, item["materia"], carrera_id,
                codigo=codigo_materia, semestre=item["ciclo"],
            )
            uc_val = int(item["unidad_credito"]) if item["unidad_credito"] else 0
            cursor.execute(
                """UPDATE materias
                   SET codigo = COALESCE(NULLIF(?, ''), codigo),
                       semestre = COALESCE(NULLIF(?, ''), semestre),
                       unidad_credito = CASE WHEN ? > 0 THEN ? ELSE unidad_credito END,
                       horas_academicas = CASE
                           WHEN COALESCE(horas_academicas, 0) > 1 THEN horas_academicas
                           WHEN ? > 0 THEN ?
                           ELSE 2
                       END,
                       mencion = ?
                   WHERE id = ?""",
                (
                    item["codigo"],
                    item["ciclo"],
                    uc_val,
                    uc_val,
                    uc_val,
                    uc_val,
                    item["mencion"],
                    materia_id,
                ),
            )
            row_h = cursor.execute(
                "SELECT COALESCE(NULLIF(horas_academicas, 0), NULLIF(unidad_credito, 0), 2) AS h FROM materias WHERE id = ?",
                (materia_id,),
            ).fetchone()
            horas_materia = int(row_h["h"]) if row_h else (uc_val if uc_val > 0 else 2)
            seccion_final = item["seccion"] or str(10100 + indice)
            cursor.execute(
                """INSERT INTO carga_academica
                     (profesor_id, materia_id, seccion, bloques_semanales, salon_preferido,
                      dia_fijo, hora_inicio_fija, hora_fin_fija, compartidos)
                   VALUES (?, ?, ?, ?, ?, '', '', '', ?)
                   ON CONFLICT(profesor_id, materia_id, seccion) DO UPDATE SET
                       bloques_semanales = excluded.bloques_semanales,
                       salon_preferido = excluded.salon_preferido,
                       dia_fijo = '',
                       hora_inicio_fija = '',
                       hora_fin_fija = '',
                       compartidos = excluded.compartidos""",
                (profesor_id, materia_id, seccion_final, horas_materia, item["ambiente"], item["compartidos"]),
            )
            row_ca = cursor.execute(
                "SELECT id FROM carga_academica WHERE profesor_id = ? AND materia_id = ? AND seccion = ?",
                (profesor_id, materia_id, seccion_final),
            ).fetchone()
            if row_ca:
                ids_cargas_guardadas.append(row_ca["id"])

        # Eliminar las materias/cargas anteriores del docente que fueron eliminadas en el formulario
        if ids_cargas_guardadas:
            placeholders = ", ".join("?" for _ in ids_cargas_guardadas)
            cursor.execute(
                f"DELETE FROM asignaciones_horario WHERE profesor_id = ? AND carga_id NOT IN ({placeholders})",
                [profesor_id] + ids_cargas_guardadas,
            )
            cursor.execute(
                f"DELETE FROM carga_academica WHERE profesor_id = ? AND id NOT IN ({placeholders})",
                [profesor_id] + ids_cargas_guardadas,
            )

        conn.commit()
        return {"insertados": len(materias_validas), "docente": docente, "errores": []}
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def obtener_codigos_materia():
    """Retorna los códigos de materia registrados en la base de datos."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT DISTINCT TRIM(codigo) AS codigo FROM materias "
        "WHERE TRIM(COALESCE(codigo, '')) <> '' ORDER BY UPPER(TRIM(codigo)) ASC"
    )
    codigos = [row["codigo"] for row in cursor.fetchall()]
    conn.close()
    return codigos


def obtener_secciones():
    """Retorna las secciones registradas en la base de datos, ordenadas numéricamente."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT DISTINCT TRIM(seccion) AS seccion FROM carga_academica "
        "WHERE TRIM(COALESCE(seccion, '')) <> '' "
        "ORDER BY LENGTH(TRIM(seccion)) ASC, CAST(TRIM(seccion) AS INTEGER) ASC, TRIM(seccion) ASC"
    )
    secciones = [row["seccion"] for row in cursor.fetchall()]
    conn.close()
    return secciones


def obtener_salones():
    """Retorna todos los ambientes registrados con su capacidad, edificio y tipo."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id, codigo, capacidad, edificio, COALESCE(es_camillas, 0) AS es_camillas "
        "FROM salones ORDER BY id ASC"
    )
    salones = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return salones


def obtener_salon(codigo):
    """Retorna el ambiente indicado o None si no existe."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id, codigo, capacidad, edificio, COALESCE(es_camillas, 0) AS es_camillas "
        "FROM salones WHERE UPPER(codigo) = UPPER(?)",
        (str(codigo).strip(),),
    )
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None


def obtener_capacidad_salon(codigo, por_defecto=CAPACIDAD_DEFECTO):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT capacidad FROM salones WHERE UPPER(codigo) = UPPER(?)",
        (str(codigo).strip(),),
    )
    row = cursor.fetchone()
    conn.close()
    return int(row["capacidad"]) if row and row["capacidad"] is not None else por_defecto


def actualizar_capacidad_salon(codigo, capacidad):
    """Actualiza la capacidad de un ambiente (funciona también para los manuales)."""
    return actualizar_datos_ambiente(codigo, capacidad=capacidad)


def actualizar_datos_ambiente(codigo, capacidad=None, es_camillas=None, edificio=None):
    """Actualiza capacidad, tipo de camillas y/o edificio de un ambiente."""
    campos = []
    valores = []

    if capacidad is not None:
        try:
            capacidad = int(capacidad)
        except (TypeError, ValueError):
            return False
        if capacidad <= 0:
            return False
        campos.append("capacidad = ?")
        valores.append(capacidad)

    if es_camillas is not None:
        campos.append("es_camillas = ?")
        valores.append(1 if es_camillas else 0)

    if edificio is not None:
        campos.append("edificio = ?")
        valores.append(str(edificio).strip())

    if not campos:
        return False

    valores.append(str(codigo).strip())
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        f"UPDATE salones SET {', '.join(campos)} WHERE UPPER(codigo) = UPPER(?)",
        valores,
    )
    actualizados = cursor.rowcount
    conn.commit()
    conn.close()
    return actualizados > 0


def eliminar_salon(codigo):
    """Elimina un ambiente individualmente y libera las cargas que lo usaban."""
    codigo = str(codigo).strip().upper()
    if not codigo:
        return False

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM salones WHERE UPPER(codigo) = UPPER(?)", (codigo,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        return False

    cursor.execute(
        "UPDATE carga_academica SET salon_preferido = '' "
        "WHERE UPPER(TRIM(salon_preferido)) = ?",
        (codigo,),
    )
    cursor.execute("DELETE FROM salones WHERE id = ?", (row["id"],))
    conn.commit()
    conn.close()
    return True


def obtener_profesores():
    """Retorna todos los docentes registrados con sus datos personales."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id, nombre, COALESCE(cedula, '') AS cedula, "
        "COALESCE(email, '') AS email, COALESCE(condicion, 'Contratado/a') AS condicion "
        "FROM profesores ORDER BY nombre ASC"
    )
    profesores = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return profesores


def _normalizar_materia(nombre):
    normalizado = unicodedata.normalize("NFD", str(nombre or "").strip().casefold())
    sin_acentos = "".join(
        caracter for caracter in normalizado
        if unicodedata.category(caracter) != "Mn"
    )
    return " ".join(sin_acentos.split())


def obtener_materias_profesor(nombre):
    """Devuelve las materias actualmente asignadas al docente, sin duplicados."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        """
        SELECT DISTINCT m.nombre
        FROM carga_academica ca
        JOIN profesores p ON p.id = ca.profesor_id
        JOIN materias m ON m.id = ca.materia_id
        WHERE LOWER(p.nombre) = LOWER(?)
        ORDER BY m.nombre COLLATE NOCASE
        """,
        (str(nombre).strip(),),
    )
    materias = [row["nombre"] for row in cursor.fetchall()]
    conn.close()
    return materias


def obtener_nombres_materias():
    """Retorna todos los nombres de materia registrados en la base de datos."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT DISTINCT TRIM(nombre) AS nombre FROM materias "
        "WHERE TRIM(COALESCE(nombre, '')) <> '' ORDER BY nombre COLLATE NOCASE ASC"
    )
    nombres = [row["nombre"] for row in cursor.fetchall()]
    conn.close()
    return nombres


def obtener_cargas_docente(nombre):
    """Retorna detalle completo de las cargas académicas asociadas al docente."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        """
        SELECT
            ca.id AS carga_id,
            ca.seccion,
            ca.salon_preferido,
            COALESCE(ca.compartidos, '') AS compartidos,
            ca.bloques_semanales,
            m.nombre AS materia,
            m.codigo,
            COALESCE(m.semestre, '') AS ciclo,
            COALESCE(m.unidad_credito, 0) AS unidad_credito,
            COALESCE(NULLIF(m.horas_academicas, 0), NULLIF(m.unidad_credito, 0), ca.bloques_semanales, 2) AS horas_academicas,
            COALESCE(m.mencion, '') AS mencion,
            c.nombre AS carrera
        FROM carga_academica ca
        JOIN profesores p ON p.id = ca.profesor_id
        JOIN materias m ON m.id = ca.materia_id
        JOIN carreras c ON c.id = m.carrera_id
        WHERE LOWER(p.nombre) = LOWER(?)
        ORDER BY m.nombre COLLATE NOCASE, c.nombre COLLATE NOCASE
        """,
        (str(nombre).strip(),),
    )
    cargas = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return cargas


def obtener_opciones_carga():
    """Opciones académicas disponibles, derivadas de la base de datos."""
    conn = get_connection()
    cursor = conn.cursor()
    codigos = [
        row["codigo"].strip()
        for row in cursor.execute(
            "SELECT DISTINCT TRIM(codigo) AS codigo FROM materias "
            "WHERE TRIM(COALESCE(codigo, '')) <> '' ORDER BY UPPER(TRIM(codigo)) ASC"
        ).fetchall()
    ]
    secciones = [
        row["seccion"].strip()
        for row in cursor.execute(
            "SELECT DISTINCT TRIM(seccion) AS seccion FROM carga_academica "
            "WHERE TRIM(COALESCE(seccion, '')) <> '' "
            "ORDER BY LENGTH(TRIM(seccion)) ASC, CAST(TRIM(seccion) AS INTEGER) ASC, TRIM(seccion) ASC"
        ).fetchall()
    ]
    carreras = [
        row["nombre"].strip()
        for row in cursor.execute(
            "SELECT DISTINCT nombre FROM carreras WHERE TRIM(nombre) <> '' ORDER BY nombre COLLATE NOCASE ASC"
        ).fetchall()
    ]
    mencies = [
        row["mencion"].strip()
        for row in cursor.execute(
            "SELECT DISTINCT TRIM(mencion) AS mencion FROM materias "
            "WHERE TRIM(COALESCE(mencion, '')) <> '' ORDER BY mencion COLLATE NOCASE ASC"
        ).fetchall()
    ]
    ambientes = [
        row["codigo"].strip()
        for row in cursor.execute(
            "SELECT DISTINCT TRIM(salon_preferido) AS codigo FROM carga_academica "
            "WHERE TRIM(COALESCE(salon_preferido, '')) <> '' ORDER BY codigo COLLATE NOCASE ASC"
        ).fetchall()
    ]
    conn.close()
    return {
        "codigos": codigos,
        "secciones": secciones,
        "carreras": carreras,
        "menciones": mencies,
        "ambientes": ambientes,
    }


def actualizar_ficha_docente_completa(nombre_actual, datos):
    """Guarda todos los datos del docente y sus cargas académicas detalle por materia."""
    def normal_texto(valor):
        return str(valor or "").strip()

    docente_actual = normal_texto(nombre_actual)
    nombre_nuevo = _normalizar_espacios(datos.get("nuevo_nombre")).upper()
    cedula = _normalizar_espacios(datos.get("cedula")).upper().replace(" ", "")
    email = _normalizar_espacios(datos.get("email")).lower()
    if not docente_actual or not nombre_nuevo:
        return False, "Indique el nombre y apellido del docente."

    disponibilidades_guardadas = datos.get("disponibilidades", None)
    if disponibilidades_guardadas is not None:
        disponibilidades = list(dict.fromkeys(
            tuple(bloque) for bloque in disponibilidades_guardadas
        ))
    else:
        disponibilidades = None

    materias = list(datos.get("materias") or [])
    errores = validar_profesor_basico(nombre_nuevo, cedula, email, requerir_cedula=False)
    codigos_existentes = {
        str(codigo or "").strip().upper()
        for codigo in obtener_codigos_materia()
    }

    for indice, materia in enumerate(materias, 1):
        codigo = normal_texto(materia.get("codigo")).upper()
        seccion = normal_texto(materia.get("seccion"))
        uc = normal_texto(materia.get("unidad_credito"))
        if codigo and codigo not in codigos_existentes and not codigo.isalnum():
            errores.append(f"Materia {indice}: el código solo admite letras y números.")
        if seccion and not seccion.isdigit():
            errores.append(f"Materia {indice}: la sección solo admite números.")
        if uc and not uc.isdigit():
            errores.append(f"Materia {indice}: la unidad de crédito solo admite números.")
    if errores:
        return False, " ".join(errores)

    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT id FROM profesores WHERE LOWER(nombre) = LOWER(?)",
            (docente_actual,),
        )
        row = cursor.fetchone()
        if not row:
            return False, "No se encontró el docente indicado."
        docente_id = row["id"]

        cursor.execute(
            "SELECT id FROM profesores WHERE LOWER(nombre) = LOWER(?) AND id <> ?",
            (nombre_nuevo, docente_id),
        )
        if cursor.fetchone():
            return False, "Ya existe otro docente con ese nombre."

        cursor.execute(
            "UPDATE profesores SET nombre = ?, cedula = ?, email = ?, condicion = ?, mencion = ? WHERE id = ?",
            (
                nombre_nuevo,
                cedula,
                email,
                normal_texto(datos.get("condicion") or "Contratado/a"),
                normal_texto(datos.get("mencion")),
                docente_id,
            ),
        )

        if disponibilidades is not None:
            cursor.execute("DELETE FROM disponibilidades WHERE profesor_id = ?", (docente_id,))
            cursor.executemany(
                """
                INSERT OR IGNORE INTO disponibilidades (profesor_id, dia, hora_inicio, hora_fin)
                VALUES (?, ?, ?, ?)
                """,
                [
                    (docente_id, bloque[0], bloque[1], bloque[2])
                    for bloque in disponibilidades
                    if len(bloque) == 3 and bloque[0] in DIAS_ORDEN and bloque[1] < bloque[2]
                ],
            )

        cargas_actuales = [
            dict(row_)
            for row_ in cursor.execute(
                "SELECT id, materia_id FROM carga_academica WHERE profesor_id = ?",
                (docente_id,),
            ).fetchall()
        ]
        # Eliminar primero las cargas que el usuario quitó en el formulario para evitar colisiones UNIQUE
        ids_enviados = {
            int(m["carga_id"])
            for m in materias
            if m.get("carga_id") is not None and str(m.get("carga_id")).isdigit()
        }
        ids_eliminar_prev = [fila["id"] for fila in cargas_actuales if fila["id"] not in ids_enviados]
        if ids_eliminar_prev:
            ph_prev = ", ".join("?" for _ in ids_eliminar_prev)
            cursor.execute(
                f"DELETE FROM asignaciones_horario WHERE carga_id IN ({ph_prev})",
                ids_eliminar_prev,
            )
            cursor.execute(
                f"DELETE FROM carga_academica WHERE id IN ({ph_prev}) AND profesor_id = ?",
                ids_eliminar_prev + [docente_id],
            )

        carreras_db_todas = [
            dict(r) for r in cursor.execute("SELECT id, nombre, facultad_id FROM carreras ORDER BY id ASC").fetchall()
        ]
        carreras_cache = {}
        ids_deseados = []

        for indice, materia in enumerate(materias, 1):
            nombre_materia = normal_texto(materia.get("materia"))
            carrera = normal_texto(materia.get("carrera"))
            mencion = normal_texto(materia.get("mencion"))
            if mencion.lower() in ("", "sin mencion", "sin mención"):
                mencion = ""
            if not nombre_materia or not carrera:
                continue

            if mencion and mencion.lower() in carrera.lower():
                carrera_full = carrera
            else:
                carrera_full = f"{carrera} - {mencion}" if mencion else carrera
            clave_can = clave_carrera_canonica(carrera)
            carrera_id = carreras_cache.get(clave_can)
            if carrera_id is None:
                for c_row in carreras_db_todas:
                    if (
                        c_row["nombre"].strip().casefold() in (carrera_full.casefold(), carrera.strip().casefold())
                        or clave_carrera_canonica(c_row["nombre"]) == clave_can
                    ):
                        carrera_id = c_row["id"]
                        break
            if carrera_id is None:
                clave_carrera = _normalizar_materia(carrera_full)
                pnf_like = clave_carrera.startswith("pnf ") or "programa nacional de formacion" in clave_carrera
                admin_like = any(
                    palabra in clave_carrera
                    for palabra in (
                        "administracion", "recursos humanos", "recursos materiales", "financier",
                        "mercadeo", "informatica", "organizacion y sistemas", "oys",
                    )
                )
                if pnf_like:
                    facultad_id = get_or_create_facultad(cursor, FACULTAD_PNF)
                else:
                    facultad = "Facultad de Administración" if admin_like else "Facultad de Educación"
                    facultad_id = get_or_create_facultad(cursor, facultad)
                carrera_id = get_or_create_carrera(cursor, carrera_full, facultad_id)
            carreras_cache[clave_can] = carrera_id
            codigo = normal_texto(materia.get("codigo")).upper()
            ciclo = normal_texto(materia.get("ciclo")).upper()
            unidad_credito = normal_texto(materia.get("unidad_credito"))
            seccion = normal_texto(materia.get("seccion"))
            compartidos = normal_texto(materia.get("compartidos"))
            ambiente = normal_texto(materia.get("ambiente")).upper()
            materia_id = get_or_create_materia(
                cursor,
                nombre_materia,
                carrera_id,
                codigo=codigo,
                semestre=ciclo,
            )
            uc_val = int(unidad_credito) if unidad_credito else 0
            cursor.execute(
                """
                UPDATE materias
                SET codigo = COALESCE(NULLIF(?, ''), codigo),
                    semestre = COALESCE(NULLIF(?, ''), semestre),
                    mencion = COALESCE(NULLIF(?, ''), mencion),
                    unidad_credito = CASE WHEN ? > 0 THEN ? ELSE unidad_credito END,
                    horas_academicas = CASE
                        WHEN COALESCE(horas_academicas, 0) > 1 THEN horas_academicas
                        WHEN ? > 0 THEN ?
                        ELSE 2
                    END
                WHERE id = ?
                """,
                (codigo, ciclo, mencion, uc_val, uc_val, uc_val, uc_val, materia_id),
            )
            row_h = cursor.execute(
                "SELECT COALESCE(NULLIF(horas_academicas, 0), NULLIF(unidad_credito, 0), 2) AS h FROM materias WHERE id = ?",
                (materia_id,),
            ).fetchone()
            horas_materia = int(row_h["h"]) if row_h else (uc_val if uc_val > 0 else 2)

            carga_id = materia.get("carga_id")
            if carga_id:
                if disponibilidades is not None:
                    cursor.execute(
                        """
                        UPDATE carga_academica
                        SET materia_id = ?, seccion = ?, bloques_semanales = ?, salon_preferido = ?,
                            dia_fijo = '', hora_inicio_fija = '', hora_fin_fija = '', compartidos = ?
                        WHERE id = ? AND profesor_id = ?
                        """,
                        (materia_id, seccion, horas_materia, ambiente, compartidos, carga_id, docente_id),
                    )
                else:
                    cursor.execute(
                        """
                        UPDATE carga_academica
                        SET materia_id = ?, seccion = ?, bloques_semanales = ?, salon_preferido = ?, compartidos = ?
                        WHERE id = ? AND profesor_id = ?
                        """,
                        (materia_id, seccion, horas_materia, ambiente, compartidos, carga_id, docente_id),
                    )
                if cursor.rowcount == 0:
                    raise ValueError("No se encontró una carga académica para actualizar.")
            else:
                cursor.execute(
                    """
                    INSERT INTO carga_academica
                    (profesor_id, materia_id, seccion, bloques_semanales, salon_preferido,
                     dia_fijo, hora_inicio_fija, hora_fin_fija, compartidos)
                    VALUES (?, ?, ?, ?, ?, '', '', '', ?)
                    ON CONFLICT(profesor_id, materia_id, seccion) DO UPDATE SET
                        bloques_semanales = excluded.bloques_semanales,
                        salon_preferido = excluded.salon_preferido,
                        dia_fijo = '',
                        hora_inicio_fija = '',
                        hora_fin_fija = '',
                        compartidos = excluded.compartidos
                    """,
                    (docente_id, materia_id, seccion or f"SIN-SECCION-{indice}", horas_materia, ambiente, compartidos),
                )
                row_carga = cursor.execute(
                    """
                    SELECT id FROM carga_academica
                    WHERE profesor_id = ? AND materia_id = ? AND seccion = ?
                    """,
                    (docente_id, materia_id, seccion or f"SIN-SECCION-{indice}"),
                ).fetchone()
                carga_id = row_carga["id"]
            ids_deseados.append(carga_id)

        if ids_deseados:
            placeholders = ", ".join("?" for _ in ids_deseados)
            cursor.execute(
                f"DELETE FROM asignaciones_horario WHERE profesor_id = ? AND carga_id NOT IN ({placeholders})",
                [docente_id] + ids_deseados,
            )
            cursor.execute(
                f"DELETE FROM carga_academica WHERE profesor_id = ? AND id NOT IN ({placeholders})",
                [docente_id] + ids_deseados,
            )
        else:
            cursor.execute("DELETE FROM asignaciones_horario WHERE profesor_id = ?", (docente_id,))
            cursor.execute("DELETE FROM carga_academica WHERE profesor_id = ?", (docente_id,))

        conn.commit()
        return True, ""
    except Exception as ex:
        conn.rollback()
        raise
    finally:
        conn.close()


def _actualizar_materias_profesor(cursor, profesor_id, nombres_materias):
    solicitadas = []
    normalizadas = set()
    for nombre in nombres_materias:
        nombre = str(nombre or "").strip()
        clave = _normalizar_materia(nombre)
        if nombre and clave and clave not in normalizadas:
            normalizadas.add(clave)
            solicitadas.append((nombre, clave))

    cursor.execute(
        """
        SELECT ca.id AS carga_id, ca.materia_id, m.nombre, m.carrera_id, m.semestre
        FROM carga_academica ca
        JOIN materias m ON m.id = ca.materia_id
        WHERE ca.profesor_id = ?
        """,
        (profesor_id,),
    )
    cargas_actuales = [dict(row) for row in cursor.fetchall()]
    por_nombre_actual = {}
    for carga in cargas_actuales:
        por_nombre_actual.setdefault(_normalizar_materia(carga["nombre"]), []).append(carga)

    carreras_actuales = list(dict.fromkeys(carga["carrera_id"] for carga in cargas_actuales))
    ids_materias_deseadas = set()

    for nombre, clave in solicitadas:
        cargas_misma_materia = por_nombre_actual.get(clave, [])
        if cargas_misma_materia:
            ids_materias_deseadas.update(carga["materia_id"] for carga in cargas_misma_materia)
            continue

        cursor.execute("SELECT id, carrera_id, nombre FROM materias")
        coincidencias = [
            dict(row) for row in cursor.fetchall()
            if _normalizar_materia(row["nombre"]) == clave
        ]
        en_carreras_actuales = [
            materia for materia in coincidencias
            if materia["carrera_id"] in carreras_actuales
        ]
        candidatas = en_carreras_actuales or coincidencias
        if len(candidatas) > 1:
            return f"La materia '{nombre}' existe en varias carreras; use el nombre exacto de una materia ya asignada."
        if candidatas:
            ids_materias_deseadas.add(candidatas[0]["id"])
        elif carreras_actuales:
            materia_id = get_or_create_materia(
                cursor,
                nombre,
                carreras_actuales[0],
                semestre=cargas_actuales[0].get("semestre") or "1er Semestre",
            )
            ids_materias_deseadas.add(materia_id)
        else:
            return (
                f"No se puede agregar '{nombre}' porque el docente no tiene una carrera "
                "asociada donde registrar la nueva materia."
            )

    cargas_a_eliminar = [
        carga["carga_id"] for carga in cargas_actuales
        if carga["materia_id"] not in ids_materias_deseadas
    ]
    if cargas_a_eliminar:
        placeholders = ",".join("?" for _ in cargas_a_eliminar)
        cursor.execute(
            f"DELETE FROM asignaciones_horario WHERE carga_id IN ({placeholders})",
            cargas_a_eliminar,
        )
        cursor.execute(
            f"DELETE FROM carga_academica WHERE id IN ({placeholders})",
            cargas_a_eliminar,
        )

    ids_actuales = {carga["materia_id"] for carga in cargas_actuales}
    for materia_id in ids_materias_deseadas - ids_actuales:
        cursor.execute(
            """
            INSERT INTO carga_academica
            (profesor_id, materia_id, seccion, bloques_semanales, salon_preferido,
             dia_fijo, hora_inicio_fija, hora_fin_fija)
            VALUES (?, ?, 'Sección A', 1, '', '', '', '')
            ON CONFLICT(profesor_id, materia_id, seccion) DO NOTHING
            """,
            (profesor_id, materia_id),
        )

    return None


def actualizar_profesor(
    nombre_actual,
    nuevo_nombre,
    cedula="",
    email="",
    condicion="Contratado/a",
    materias=None,
):
    """Actualiza los datos y, opcionalmente, la lista de materias del docente."""
    nombre_actual = str(nombre_actual).strip()
    nuevo_nombre = str(nuevo_nombre).strip()
    if not nombre_actual or not nuevo_nombre:
        return False

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id FROM profesores WHERE LOWER(nombre) = LOWER(?)",
        (nombre_actual,),
    )
    row = cursor.fetchone()
    if not row:
        conn.close()
        return False

    cursor.execute(
        "SELECT id FROM profesores WHERE LOWER(nombre) = LOWER(?) AND id <> ?",
        (nuevo_nombre, row["id"]),
    )
    if cursor.fetchone():
        conn.close()
        return False

    cursor.execute(
        "UPDATE profesores SET nombre = ?, cedula = ?, email = ?, condicion = ? WHERE id = ?",
        (
            nuevo_nombre,
            str(cedula or "").strip(),
            str(email or "").strip(),
            str(condicion or "Contratado/a").strip(),
            row["id"],
        ),
    )
    actualizados = cursor.rowcount
    if materias is not None:
        error_materias = _actualizar_materias_profesor(cursor, row["id"], materias)
        if error_materias:
            conn.rollback()
            conn.close()
            return False
    conn.commit()
    conn.close()
    return actualizados > 0


def eliminar_profesor(nombre):
    """Elimina un docente y, en cascada, sus cargas, disponibilidades y asignaciones."""
    nombre = str(nombre).strip()
    if not nombre:
        return False

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM profesores WHERE LOWER(nombre) = LOWER(?)", (nombre,))
    eliminados = cursor.rowcount
    conn.commit()
    conn.close()
    return eliminados > 0


def siguiente_numero_ambiente():
    """Siguiente número consecutivo global de los ambientes con código P(piso)-(nro)."""
    import re

    mayor = 0
    for salon in obtener_salones():
        coincidencia = re.match(r"^P\d+-(\d+)$", str(salon["codigo"]).strip().upper())
        if coincidencia:
            mayor = max(mayor, int(coincidencia.group(1)))
    return mayor + 1


def siguiente_codigo_ambiente(piso):
    """Sugiere el próximo código P(piso)-(nro) continuando el total global."""
    try:
        numero_piso = int(piso)
    except (TypeError, ValueError):
        numero_piso = 1
    return f"P{numero_piso}-{siguiente_numero_ambiente():02d}"


def crear_ambiente(codigo, capacidad=CAPACIDAD_DEFECTO, edificio="Ambiente Especial", es_camillas=False):
    """Crea o actualiza un ambiente conservando la nomenclatura indicada."""
    codigo = str(codigo).strip().upper()
    if not codigo:
        return None
    try:
        capacidad = int(capacidad)
    except (TypeError, ValueError):
        capacidad = CAPACIDAD_DEFECTO
    if capacidad <= 0:
        capacidad = CAPACIDAD_DEFECTO
    camillas = 1 if es_camillas else 0

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM salones WHERE UPPER(codigo) = UPPER(?)", (codigo,))
    row = cursor.fetchone()
    if row:
        cursor.execute(
            "UPDATE salones SET capacidad = ?, edificio = ?, es_camillas = ? WHERE id = ?",
            (capacidad, edificio, camillas, row["id"]),
        )
    else:
        cursor.execute(
            "INSERT INTO salones (codigo, capacidad, edificio, es_camillas) VALUES (?, ?, ?, ?)",
            (codigo, capacidad, edificio, camillas),
        )
    conn.commit()
    conn.close()
    return codigo


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
