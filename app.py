import os
from datetime import datetime
from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    send_file,
    Response,
    jsonify,
)
from werkzeug.utils import secure_filename

from database import (
    init_db,
    get_connection,
    clear_all_data,
    get_or_create_facultad,
    get_or_create_carrera,
    get_or_create_profesor,
    get_or_create_materia,
    get_or_create_salon,
    guardar_carrera,
    FACULTAD_PNF,
    obtener_facultades,
    obtener_materias_carrera,
    obtener_profesores_capacitados,
    registrar_ficha_docente,
    DIAS_ORDEN,
    BLOQUES_HORARIOS,
)
from parsers import importar_excel, importar_docx, importar_txt, parsear_disponibilidad
from scheduler import (
    autoasignar_ambientes_libres,
    generar_horarios_automaticos,
    validar_cambio_manual,
)
from seed_demo import crear_archivos_ejemplo, cargar_demo_en_bd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get(
    "FLASK_SECRET_KEY",
    "servicio_comunitario_universidad_secret_key",
)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

DIAS_CARGA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"]
FRANJAS_CARGA = [
    ("07:00", "07:45"), ("07:45", "08:30"), ("08:30", "09:15"),
    ("09:20", "10:05"), ("10:05", "10:50"), ("10:50", "11:35"),
    ("11:35", "12:15"), ("12:30", "13:15"), ("13:15", "13:55"),
    ("14:10", "14:55"), ("14:55", "15:35"), ("15:40", "16:25"),
    ("16:25", "17:05"), ("17:10", "17:45"), ("17:45", "18:25"),
]

# Inicializar base de datos y plantillas al arrancar
init_db()
crear_archivos_ejemplo(BASE_DIR)


@app.context_processor
def inject_global_profesores():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre FROM profesores ORDER BY nombre ASC")
    lista_profesores = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return {"profesores_global": lista_profesores}



def obtener_estadisticas():
    conn = get_connection()
    cursor = conn.cursor()
    stats = {}
    for tabla, clave in [
        ("facultades", "facultades"),
        ("carreras", "carreras"),
        ("profesores", "profesores"),
        ("materias", "materias"),
        ("salones", "salones"),
        ("asignaciones_horario", "clases_programadas"),
    ]:
        cursor.execute(f"SELECT COUNT(*) AS total FROM {tabla}")
        stats[clave] = cursor.fetchone()["total"]
    cursor.execute("SELECT COUNT(DISTINCT salon_id) AS total FROM asignaciones_horario")
    stats["ambientes_en_uso"] = cursor.fetchone()["total"]
    cursor.execute("SELECT COUNT(DISTINCT profesor_id) AS total FROM asignaciones_horario")
    stats["profesores_activos"] = cursor.fetchone()["total"]
    cursor.execute(
        """SELECT COUNT(*) AS total
           FROM asignaciones_horario a
           JOIN asignaciones_horario b
             ON a.id < b.id
            AND a.dia = b.dia
            AND a.hora_inicio < b.hora_fin
            AND b.hora_inicio < a.hora_fin
           JOIN salones sa ON sa.id = a.salon_id
           JOIN salones sb ON sb.id = b.salon_id
           WHERE a.profesor_id = b.profesor_id
              OR (
                   UPPER(sa.codigo) = UPPER(sb.codigo)
                   AND UPPER(sa.codigo) NOT LIKE '%VIRTUAL%'
                   AND UPPER(sa.codigo) NOT LIKE '%COORDINAC%'
                   AND UPPER(sa.codigo) NOT LIKE '%ONLINE%'
                   AND UPPER(sa.codigo) NOT LIKE '%CANCHA%'
                   AND UPPER(sa.codigo) NOT LIKE '%AUDITORIO%'
              )"""
    )
    stats["conflictos"] = cursor.fetchone()["total"]
    conn.close()
    return stats


def obtener_conflictos_horarios_detallados():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        """SELECT
               a.id AS asignacion_a,
               b.id AS asignacion_b,
               a.salon_id AS salon_a_id,
               b.salon_id AS salon_b_id,
               a.dia,
               a.hora_inicio,
               a.hora_fin,
               pa.nombre AS profesor_a,
               pb.nombre AS profesor_b,
               ma.nombre AS materia_a,
               mb.nombre AS materia_b,
               sa.codigo AS ambiente_a,
               sb.codigo AS ambiente_b
           FROM asignaciones_horario a
           JOIN asignaciones_horario b ON a.id < b.id
             AND a.dia = b.dia
             AND a.hora_inicio < b.hora_fin
             AND b.hora_inicio < a.hora_fin
           JOIN profesores pa ON pa.id = a.profesor_id
           JOIN profesores pb ON pb.id = b.profesor_id
           JOIN materias ma ON ma.id = a.materia_id
           JOIN materias mb ON mb.id = b.materia_id
           JOIN salones sa ON sa.id = a.salon_id
           JOIN salones sb ON sb.id = b.salon_id
           ORDER BY a.dia, a.hora_inicio, pa.nombre, pb.nombre"""
    )
    conflictos = []
    for row in cursor.fetchall():
        item = dict(row)
        motivos = []
        if item["profesor_a"].upper() == item["profesor_b"].upper():
            motivos.append("docente")
        ambiente_a = str(item["ambiente_a"] or "").upper()
        ambiente_b = str(item["ambiente_b"] or "").upper()
        compartibles = ("VIRTUAL", "COORDINAC", "ONLINE", "CANCHA", "AUDITORIO")
        mismo_ambiente = ambiente_a == ambiente_b
        if mismo_ambiente and not any(clave in ambiente_a for clave in compartibles):
            motivos.append("ambiente")
        if motivos:
            item["motivos"] = " y ".join(motivos)
            conflictos.append(item)
    conn.close()
    return conflictos


def consultar_oferta_academica(facultad_id=None, carrera_id=None, profesor_id=None, salon_id=None, dia=None):
    conn = get_connection()
    cursor = conn.cursor()

    query = """
        SELECT
            a.id AS asignacion_id,
            a.dia,
            a.hora_inicio,
            a.hora_fin,
            a.estado,
            p.id AS profesor_id,
            p.nombre AS profesor_nombre,
            p.cedula AS profesor_cedula,
            m.id AS materia_id,
            m.nombre AS materia_nombre,
            m.semestre,
            ca.seccion,
            c.id AS carrera_id,
            c.nombre AS carrera_nombre,
            f.id AS facultad_id,
            f.nombre AS facultad_nombre,
            s.id AS salon_id,
            s.codigo AS salon_codigo,
            s.edificio AS salon_edificio,
            s.capacidad AS salon_capacidad
        FROM asignaciones_horario a
        JOIN carga_academica ca ON ca.id = a.carga_id
        JOIN profesores p ON p.id = a.profesor_id
        JOIN materias m ON m.id = a.materia_id
        JOIN carreras c ON c.id = m.carrera_id
        JOIN facultades f ON f.id = c.facultad_id
        JOIN salones s ON s.id = a.salon_id
        WHERE 1=1
    """
    params = []
    if facultad_id:
        query += " AND f.id = ?"
        params.append(facultad_id)
    if carrera_id:
        query += " AND c.id = ?"
        params.append(carrera_id)
    if profesor_id:
        query += " AND p.id = ?"
        params.append(profesor_id)
    if salon_id:
        query += " AND s.id = ?"
        params.append(salon_id)
    if dia:
        query += " AND a.dia = ?"
        params.append(dia)

    cursor.execute(query, params)
    filas = [dict(r) for r in cursor.fetchall()]
    conn.close()

    orden_dia = {d: i for i, d in enumerate(DIAS_ORDEN)}
    # Ordenado por Facultad -> Profesor (alfabéticamente) -> Día -> Hora
    filas.sort(
        key=lambda x: (
            x["facultad_nombre"],
            x["profesor_nombre"],
            orden_dia.get(x["dia"], 99),
            x["hora_inicio"],
        )
    )
    return filas


def construir_matriz_calendario(asignaciones):
    """
    Construye una matriz [hora_rango][dia] -> lista de clases asignadas,
    ideal para mostrar el horario tipo calendario de un profesor o de un salón.
    """
    rangos_unicos = sorted({f"{a['hora_inicio']} - {a['hora_fin']}" for a in asignaciones})
    if rangos_unicos:
        rangos_horas = rangos_unicos
    else:
        rangos_horas = [f"{b[0]} - {b[1]}" for b in BLOQUES_HORARIOS]

    matriz = {r: {d: [] for d in DIAS_ORDEN} for r in rangos_horas}
    for a in asignaciones:
        r = f"{a['hora_inicio']} - {a['hora_fin']}"
        d = a["dia"]
        if r in matriz and d in matriz[r]:
            matriz[r][d].append(a)

    return rangos_horas, matriz


@app.route("/logo-unesr.png")
def logo_unesr():
    return send_file(os.path.join(BASE_DIR, "src", "UNESR.png"), mimetype="image/png")


@app.route("/")
def index():
    stats = obtener_estadisticas()
    oferta_reciente = consultar_oferta_academica()[:10]
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        """SELECT c.id, c.nombre, f.nombre AS facultad_nombre,
                  COUNT(DISTINCT m.id) AS materias,
                  COUNT(DISTINCT a.id) AS clases
           FROM carreras c
           JOIN facultades f ON f.id = c.facultad_id
           LEFT JOIN materias m ON m.carrera_id = c.id
           LEFT JOIN asignaciones_horario a ON a.materia_id = m.id
           GROUP BY c.id, c.nombre, f.id, f.nombre
           ORDER BY c.nombre"""
    )
    distribucion_carreras = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return render_template(
        "index.html",
        stats=stats,
        oferta_reciente=oferta_reciente,
        distribucion_carreras=distribucion_carreras,
    )


@app.route("/nueva-carga")
def nueva_carga():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre, facultad_id FROM carreras ORDER BY nombre ASC")
    carreras = [dict(row) for row in cursor.fetchall()]
    cursor.execute("SELECT id, codigo, edificio, capacidad FROM salones ORDER BY codigo ASC")
    salones = [dict(row) for row in cursor.fetchall()]
    cursor.execute("SELECT DISTINCT nombre FROM materias ORDER BY nombre ASC")
    materias = [row["nombre"] for row in cursor.fetchall()]
    cursor.execute("SELECT DISTINCT codigo FROM materias WHERE TRIM(COALESCE(codigo, '')) <> '' ORDER BY codigo ASC")
    codigos = [row["codigo"] for row in cursor.fetchall()]
    cursor.execute("SELECT DISTINCT semestre FROM materias WHERE TRIM(COALESCE(semestre, '')) <> '' ORDER BY semestre ASC")
    ciclos = [row["semestre"] for row in cursor.fetchall()]
    conn.close()
    return render_template(
        "nueva_carga.html",
        carreras=carreras,
        salones=salones,
        materias=materias,
        codigos=codigos,
        ciclos=ciclos,
        dias_carga=DIAS_CARGA,
        franjas_carga=FRANJAS_CARGA,
    )


@app.route("/api/materias-carrera/<int:carrera_id>")
def materias_por_carrera(carrera_id):
    return jsonify(obtener_materias_carrera(carrera_id))


@app.route("/api/profesores-capacitados")
def profesores_capacitados():
    carrera_id_raw = request.args.get("carrera_id", "").strip()
    materia = request.args.get("materia", "").strip()
    busqueda = request.args.get("q", "").strip()
    try:
        carrera_id = int(carrera_id_raw) if carrera_id_raw else None
    except ValueError:
        carrera_id = None
    return jsonify(obtener_profesores_capacitados(carrera_id, materia, busqueda))


@app.route("/cargar-archivo", methods=["POST"])
def cargar_archivo():
    archivo = request.files.get("archivo")
    limpiar_previo = request.form.get("limpiar_previo") == "si"

    if not archivo or archivo.filename == "":
        flash("Por favor selecciona un archivo válido (.xlsx, .docx o .txt).", "warning")
        return redirect(url_for("nueva_carga"))

    ext = os.path.splitext(archivo.filename)[1].lower()
    if ext not in (".xlsx", ".docx", ".txt", ".csv"):
        flash("Formato no soportado. Usa archivos Excel (.xlsx), Word (.docx) o Texto (.txt).", "danger")
        return redirect(url_for("nueva_carga"))

    if limpiar_previo:
        clear_all_data(keep_salones=True)

    filename = secure_filename(archivo.filename)
    filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
    archivo.save(filepath)

    try:
        if ext == ".xlsx":
            res = importar_excel(filepath)
        elif ext == ".docx":
            res = importar_docx(filepath)
        else:
            res = importar_txt(filepath)

        res_sched = generar_horarios_automaticos()

        msg = (
            f"Archivo '{filename}' procesado con éxito: {res['insertados']} cargas académicas leídas y "
            f"{res_sched['asignados']} bloques de horario organizados sin choques."
        )
        flash(msg, "success")

        for err in res.get("errores", [])[:5]:
            flash(err, "warning")
        for m in res_sched.get("mensajes", [])[:5]:
            flash(m, "warning")

    except Exception as e:
        flash(f"Error al procesar el archivo: {str(e)}", "danger")

    return redirect(url_for("oferta_general"))


@app.route("/cargar-demo", methods=["POST"])
def cargar_demo():
    clear_all_data(keep_salones=True)
    info = cargar_demo_en_bd(BASE_DIR)
    flash(
        f"Datos de demostración cargados correctamente: {info['importados']} materias registradas y "
        f"{info['asignados']} bloques de clases asignados automáticamente.",
        "success",
    )
    return redirect(url_for("oferta_general"))


@app.route("/limpiar-datos", methods=["POST"])
def limpiar_datos():
    clear_all_data(keep_salones=True)
    flash("Se eliminó la carga académica y los horarios. Carreras, materias y profesores se conservaron.", "info")
    return redirect(url_for("index"))


@app.route("/agregar-manual", methods=["POST"])
def agregar_manual():
    if request.form.get("tipo_carga") == "ficha_docente":
        campos_materia = {
            "materia": request.form.getlist("materia"),
            "carrera": request.form.getlist("carrera"),
            "mencion": request.form.getlist("mencion"),
            "ambiente": request.form.getlist("ambiente"),
            "codigo": request.form.getlist("codigo"),
            "ciclo": request.form.getlist("ciclo"),
            "seccion": request.form.getlist("seccion"),
            "unidad_credito": request.form.getlist("unidad_credito"),
            "compartidos": request.form.getlist("compartidos"),
        }
        materias = [
            {
                campo: valores[indice] if indice < len(valores) else ""
                for campo, valores in campos_materia.items()
            }
            for indice in range(max((len(valores) for valores in campos_materia.values()), default=0))
        ]
        disponibilidades = []
        for bloque in request.form.getlist("disponibilidad"):
            partes = bloque.split("|")
            if len(partes) == 3:
                disponibilidades.append(tuple(partes))
            else:
                flash("Una de las franjas de disponibilidad tiene un formato inválido.", "danger")
                return redirect(url_for("nueva_carga"))

        datos_docente = {
            "docente": request.form.get("docente", ""),
            "cedula": request.form.get("cedula", ""),
            "correo": request.form.get("correo", ""),
            "condicion": request.form.get("condicion", ""),
            "materias": materias,
            "disponibilidades": disponibilidades,
        }
        try:
            resultado = registrar_ficha_docente(datos_docente)
        except Exception as error:
            flash(f"No se pudo guardar la ficha docente: {error}", "danger")
            return redirect(url_for("nueva_carga"))
        if resultado.get("errores"):
            for error in resultado["errores"]:
                flash(error, "warning")
            return redirect(url_for("nueva_carga"))

        try:
            generacion = generar_horarios_automaticos()
        except Exception as error:
            flash(f"La ficha se guardó, pero no se pudieron generar los horarios: {error}", "danger")
            return redirect(url_for("oferta_general"))
        flash(
            f"Ficha de {resultado['docente']} guardada: {resultado['insertados']} materia(s); "
            f"{generacion['asignados']} bloques asignados y {generacion['pendientes']} pendientes.",
            "success" if not generacion["pendientes"] else "warning",
        )
        return redirect(url_for("oferta_general"))

    facultad_nom = request.form.get("facultad", "").strip()
    carrera_nom = request.form.get("carrera", "").strip()
    profesor_nom = request.form.get("profesor", "").strip()
    materia_nom = request.form.get("materia", "").strip()
    disponibilidad_str = request.form.get("disponibilidad", "").strip()
    salon_pref = request.form.get("salon_preferido", "").strip()
    seccion = request.form.get("seccion", "Sección A").strip() or "Sección A"
    semestre = request.form.get("semestre", "1er Semestre").strip() or "1er Semestre"
    bloques = int(request.form.get("bloques", 1) or 1)

    if not profesor_nom or not materia_nom or not carrera_nom or not facultad_nom:
        flash("Por favor completa Facultad, Carrera, Profesor y Materia.", "warning")
        return redirect(url_for("nueva_carga"))

    conn = get_connection()
    cursor = conn.cursor()

    fac_id = get_or_create_facultad(cursor, facultad_nom)
    car_id = get_or_create_carrera(cursor, carrera_nom, fac_id)
    prof_id = get_or_create_profesor(cursor, profesor_nom)
    mat_id = get_or_create_materia(cursor, materia_nom, car_id, semestre=semestre)

    if salon_pref:
        get_or_create_salon(cursor, salon_pref)

    bloques_disp = parsear_disponibilidad(disponibilidad_str)
    if not bloques_disp:
        bloques_disp = [("Lunes", "07:00", "09:00"), ("Miércoles", "07:00", "09:00")]

    for d, hi, hf in bloques_disp:
        cursor.execute(
            "INSERT OR IGNORE INTO disponibilidades (profesor_id, dia, hora_inicio, hora_fin) VALUES (?, ?, ?, ?)",
            (prof_id, d, hi, hf),
        )

    cursor.execute(
        """
        INSERT INTO carga_academica (profesor_id, materia_id, seccion, bloques_semanales, salon_preferido)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(profesor_id, materia_id, seccion) DO UPDATE SET
            bloques_semanales = excluded.bloques_semanales,
            salon_preferido = excluded.salon_preferido
        """,
        (prof_id, mat_id, seccion, bloques, salon_pref),
    )

    conn.commit()
    conn.close()

    generar_horarios_automaticos()
    flash(f"Se agregó la materia '{materia_nom}' con el Prof. {profesor_nom} y se actualizaron los horarios.", "success")
    return redirect(url_for("oferta_general"))


@app.route("/agregar-salon", methods=["POST"])
def agregar_salon():
    codigo = request.form.get("codigo", "").strip()
    edificio = request.form.get("edificio", "Edificio Principal").strip()
    capacidad = int(request.form.get("capacidad", 40) or 40)
    if codigo:
        conn = get_connection()
        cursor = conn.cursor()
        get_or_create_salon(cursor, codigo, capacidad, edificio)
        conn.commit()
        conn.close()
        generar_horarios_automaticos()
        flash(f"Salón '{codigo}' registrado correctamente.", "success")
    return redirect(url_for("ocupacion_salones"))


@app.route("/reorganizar", methods=["POST"])
def reorganizar():
    res = generar_horarios_automaticos()
    flash(
        f"Horarios reorganizados: {res['asignados']} bloques asignados, {res['pendientes']} pendientes.",
        "success" if res["pendientes"] == 0 else "warning",
    )
    for m in res.get("mensajes", []):
        flash(m, "warning")
    return redirect(request.referrer or url_for("oferta_general"))


@app.route("/autoasignar-ambientes", methods=["POST"])
def autoasignar_ambientes():
    resultado = autoasignar_ambientes_libres()
    if resultado["reasignados"]:
        flash(
            f"Ambientes reasignados automáticamente: {resultado['reasignados']}.",
            "success",
        )
    elif not resultado["conflictos_ambiente"]:
        flash("No se encontraron colisiones de ambiente que requieran reasignación.", "info")
    for mensaje in resultado["mensajes"]:
        flash(mensaje, "warning")
    return redirect(request.referrer or url_for("oferta_general"))


@app.route("/editar-asignacion", methods=["POST"])
def editar_asignacion():
    asignacion_id = int(request.form.get("asignacion_id"))
    salon_id = int(request.form.get("salon_id"))
    dia = request.form.get("dia")
    rango_hora = request.form.get("rango_hora", "07:00-09:00")
    partes = rango_hora.split("-")
    hora_inicio = partes[0].strip()
    hora_fin = partes[1].strip()

    valido, mensaje = validar_cambio_manual(asignacion_id, salon_id, dia, hora_inicio, hora_fin)
    if not valido:
        flash(f"No se pudo cambiar el horario: {mensaje}", "danger")
        return redirect(request.referrer or url_for("oferta_general"))

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        """
        UPDATE asignaciones_horario
        SET salon_id = ?, dia = ?, hora_inicio = ?, hora_fin = ?
        WHERE id = ?
        """,
        (salon_id, dia, hora_inicio, hora_fin, asignacion_id),
    )
    conn.commit()
    conn.close()
    flash("Horario actualizado manualmente sin choques.", "success")
    return redirect(request.referrer or url_for("oferta_general"))


@app.route("/horarios", endpoint="horarios")
@app.route("/oferta-general")
def oferta_general():
    facultad_id = request.args.get("facultad_id", type=int)
    carrera_id = request.args.get("carrera_id", type=int)
    profesor_id = request.args.get("profesor_id", type=int)
    salon_id = request.args.get("salon_id", type=int)
    dia = request.args.get("dia", default="")

    asignaciones = consultar_oferta_academica(
        facultad_id=facultad_id,
        carrera_id=carrera_id,
        profesor_id=profesor_id,
        salon_id=salon_id,
        dia=dia if dia else None,
    )

    # Agrupar por Facultad para visualización organizada
    por_facultad = {}
    for a in asignaciones:
        por_facultad.setdefault(a["facultad_nombre"], []).append(a)

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre FROM facultades ORDER BY nombre ASC")
    facultades = [dict(r) for r in cursor.fetchall()]
    cursor.execute("SELECT id, nombre, facultad_id FROM carreras ORDER BY nombre ASC")
    carreras = [dict(r) for r in cursor.fetchall()]
    cursor.execute("SELECT id, nombre FROM profesores ORDER BY nombre ASC")
    profesores = [dict(r) for r in cursor.fetchall()]
    cursor.execute("SELECT id, codigo, edificio FROM salones ORDER BY codigo ASC")
    salones = [dict(r) for r in cursor.fetchall()]
    conn.close()

    return render_template(
        "oferta_general.html",
        asignaciones=asignaciones,
        por_facultad=por_facultad,
        conflictos_horario=obtener_conflictos_horarios_detallados(),
        facultades=facultades,
        carreras=carreras,
        profesores=profesores,
        salones=salones,
        filtros={
            "facultad_id": facultad_id,
            "carrera_id": carrera_id,
            "profesor_id": profesor_id,
            "salon_id": salon_id,
            "dia": dia,
        },
        dias_orden=DIAS_ORDEN,
        bloques_horarios=BLOQUES_HORARIOS,
    )


@app.route("/docentes", endpoint="docentes")
@app.route("/calendario-profesores")
def calendario_profesores():
    profesor_id = request.args.get("profesor_id", type=int)

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre, cedula, email FROM profesores ORDER BY nombre ASC")
    profesores = [dict(r) for r in cursor.fetchall()]
    cursor.execute(
        """SELECT p.id, COUNT(DISTINCT a.id) AS clases
           FROM profesores p
           LEFT JOIN asignaciones_horario a ON a.profesor_id = p.id
           GROUP BY p.id"""
    )
    clases_por_profesor = {row["id"]: row["clases"] for row in cursor.fetchall()}
    for profesor in profesores:
        profesor["clases"] = clases_por_profesor.get(profesor["id"], 0)

    if not profesor_id and profesores:
        profesor_id = profesores[0]["id"]

    profesor_actual = None
    if profesor_id:
        cursor.execute("SELECT id, nombre, cedula, email FROM profesores WHERE id = ?", (profesor_id,))
        row = cursor.fetchone()
        if row:
            profesor_actual = dict(row)

    cursor.execute(
        "SELECT profesor_id, dia, hora_inicio, hora_fin FROM disponibilidades ORDER BY dia, hora_inicio"
    )
    disponibilidades_por_profesor = {}
    for r in cursor.fetchall():
        disponibilidades_por_profesor.setdefault(r["profesor_id"], set()).add(
            f"{r['dia']}|{r['hora_inicio']}|{r['hora_fin']}"
        )

    conn.close()

    asignaciones = consultar_oferta_academica(profesor_id=profesor_id) if profesor_id else []
    rangos_horas, matriz = construir_matriz_calendario(asignaciones)

    # También construir todos los calendarios si el usuario quiere ver "Todos los profesores"
    mostrar_todos = request.args.get("todos") == "1"
    calendarios_todos = []
    if mostrar_todos:
        for p in profesores:
            asig_p = consultar_oferta_academica(profesor_id=p["id"])
            rh, mat = construir_matriz_calendario(asig_p)
            calendarios_todos.append({
                "profesor": p,
                "asignaciones": asig_p,
                "rangos_horas": rh,
                "matriz": mat,
                "disponibilidad": disponibilidades_por_profesor.get(p["id"], set()),
            })

    disponibilidad_actual = disponibilidades_por_profesor.get(profesor_id, set()) if profesor_id else set()

    return render_template(
        "calendario_profesores.html",
        profesores=profesores,
        profesor_actual=profesor_actual,
        disponibilidad_actual=disponibilidad_actual,
        disponibilidades_por_profesor=disponibilidades_por_profesor,
        dias_carga=DIAS_CARGA,
        franjas_carga=FRANJAS_CARGA,
        asignaciones=asignaciones,
        rangos_horas=rangos_horas,
        matriz=matriz,
        dias_orden=DIAS_ORDEN,
        mostrar_todos=mostrar_todos,
        calendarios_todos=calendarios_todos,
    )


@app.route("/ambientes", endpoint="ambientes")
@app.route("/ocupacion-salones")
def ocupacion_salones():
    salon_id = request.args.get("salon_id", type=int)
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        """SELECT s.id, s.codigo, s.edificio, s.capacidad, s.es_camillas,
                  COUNT(DISTINCT a.id) AS clases
           FROM salones s
           LEFT JOIN asignaciones_horario a ON a.salon_id = s.id
           GROUP BY s.id, s.codigo, s.edificio, s.capacidad, s.es_camillas
           ORDER BY s.edificio ASC, s.codigo ASC"""
    )
    salones = [dict(r) for r in cursor.fetchall()]
    conn.close()
    estadisticas_ambientes = {
        "total": len(salones),
        "en_uso": sum(1 for salon in salones if salon["clases"]),
        "disponibles": sum(1 for salon in salones if not salon["clases"]),
        "capacidad": sum(salon["capacidad"] or 0 for salon in salones),
    }

    if not salon_id and salones:
        salon_id = salones[0]["id"]

    salon_actual = next((s for s in salones if s["id"] == salon_id), None)
    asignaciones = consultar_oferta_academica(salon_id=salon_id) if salon_id else []
    rangos_horas, matriz = construir_matriz_calendario(asignaciones)

    return render_template(
        "ocupacion_salones.html",
        salones=salones,
        salon_actual=salon_actual,
        asignaciones=asignaciones,
        rangos_horas=rangos_horas,
        matriz=matriz,
        dias_orden=DIAS_ORDEN,
        estadisticas_ambientes=estadisticas_ambientes,
    )


@app.route("/carreras")
def carreras():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        """SELECT f.nombre AS facultad_nombre, c.id AS carrera_id,
                  c.nombre AS carrera_nombre, m.id AS materia_id,
                  m.nombre AS materia_nombre, m.codigo, m.semestre
           FROM carreras c
           JOIN facultades f ON f.id = c.facultad_id
           LEFT JOIN materias m ON m.carrera_id = c.id
           ORDER BY f.nombre, c.nombre,
             CASE
               WHEN LOWER(m.semestre) LIKE 'introduct%' THEN 1
               WHEN LOWER(m.semestre) LIKE 'trayecto inicial%' THEN 1
               WHEN LOWER(m.semestre) LIKE 'general%' THEN 2
               WHEN LOWER(m.semestre) LIKE 'elect%' THEN 3
               WHEN LOWER(m.semestre) LIKE 'profes%' THEN 4
               WHEN LOWER(m.semestre) LIKE 'primer trayecto%' OR LOWER(m.semestre) LIKE 'trayecto i %' OR LOWER(m.semestre) LIKE 'trayecto i -%' THEN 5
               WHEN LOWER(m.semestre) LIKE 'segundo trayecto%' OR LOWER(m.semestre) LIKE 'trayecto ii%' THEN 6
               WHEN LOWER(m.semestre) LIKE 'tercer trayecto%' OR LOWER(m.semestre) LIKE 'trayecto iii%' THEN 7
               WHEN LOWER(m.semestre) LIKE 'cuarto trayecto%' OR LOWER(m.semestre) LIKE 'trayecto iv%' THEN 8
               WHEN LOWER(m.semestre) LIKE 'trayecto v%' THEN 9
               ELSE 10
             END,
             m.semestre COLLATE NOCASE,
             m.nombre COLLATE NOCASE"""
    )
    filas = [dict(row) for row in cursor.fetchall()]
    conn.close()

    por_carrera = {}
    for fila in filas:
        carrera = por_carrera.setdefault(
            fila["carrera_id"],
            {
                "id": fila["carrera_id"],
                "nombre": fila["carrera_nombre"],
                "facultad": fila["facultad_nombre"],
                "materias": [],
            },
        )
        if fila["materia_id"] is not None:
            carrera["materias"].append(
                {
                    "nombre": fila["materia_nombre"],
                    "codigo": fila["codigo"],
                    "semestre": fila["semestre"],
                }
            )
    return render_template(
        "carreras.html",
        carreras=list(por_carrera.values()),
        facultad_pnf=FACULTAD_PNF,
        facultades=obtener_facultades(),
    )


@app.route("/carreras/guardar", methods=["POST"])
def guardar_carrera_web():
    carrera_id_raw = request.form.get("carrera_id", "").strip()
    try:
        carrera_id = int(carrera_id_raw) if carrera_id_raw else None
    except ValueError:
        flash("La carrera seleccionada no es válida.", "danger")
        return redirect(url_for("carreras"))
    try:
        guardar_carrera(
            request.form.get("nombre", ""),
            request.form.get("facultad", ""),
            carrera_id,
        )
    except ValueError as error:
        flash(str(error), "warning")
        return redirect(url_for("carreras"))
    except Exception as error:
        flash(f"No se pudo guardar la carrera: {error}", "danger")
        return redirect(url_for("carreras"))
    flash("Carrera guardada correctamente.", "success")
    return redirect(url_for("carreras"))


@app.route("/reporte-imprimible")
def reporte_imprimible():
    """
    Genera el reporte HTML autocontenido listo para imprimir / guardar como PDF,
    o lo descarga directamente como archivo .html si se pasa ?descargar=1.
    Soporta:
    - tipo=general (Oferta Académica General por Facultad)
    - tipo=profesor (Horario Tipo Calendario de un Profesor o de Todos)
    - tipo=completo (Oferta General + Calendarios de todos los Profesores)
    """
    tipo = request.args.get("tipo", "completo")
    facultad_id = request.args.get("facultad_id", type=int)
    profesor_id = request.args.get("profesor_id", type=int)
    descargar = request.args.get("descargar") == "1"
    auto_print = request.args.get("print") == "1"

    asignaciones = consultar_oferta_academica(facultad_id=facultad_id, profesor_id=profesor_id)
    por_facultad = {}
    for a in asignaciones:
        por_facultad.setdefault(a["facultad_nombre"], []).append(a)

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre, cedula, email FROM profesores ORDER BY nombre ASC")
    todos_profesores = [dict(r) for r in cursor.fetchall()]

    if profesor_id:
        profesores = [p for p in todos_profesores if p["id"] == profesor_id]
    else:
        profesores = todos_profesores
    conn.close()

    calendarios_profesores = []
    if tipo in ("profesor", "completo"):
        for p in profesores:
            asig_p = consultar_oferta_academica(facultad_id=facultad_id, profesor_id=p["id"])
            rh, mat = construir_matriz_calendario(asig_p)
            calendarios_profesores.append({
                "profesor": p,
                "asignaciones": asig_p,
                "rangos_horas": rh,
                "matriz": mat,
            })

    fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")

    html_content = render_template(
        "reporte_imprimible.html",
        tipo=tipo,
        por_facultad=por_facultad,
        calendarios_profesores=calendarios_profesores,
        todos_profesores=todos_profesores,
        profesor_id_seleccionado=profesor_id,
        dias_orden=DIAS_ORDEN,
        fecha_actual=fecha_actual,
        auto_print=auto_print,
        es_descarga=descargar,
    )

    if descargar:
        nombre_archivo = f"Reporte_Horarios_Universidad_{tipo}.html"
        return Response(
            html_content,
            mimetype="text/html",
            headers={"Content-Disposition": f"attachment; filename={nombre_archivo}"},
        )

    return html_content


@app.route("/descargar-plantilla/<formato>")
def descargar_plantilla(formato):
    rutas = crear_archivos_ejemplo(BASE_DIR)
    if formato not in rutas:
        flash("Formato de plantilla inválido.", "danger")
        return redirect(url_for("index"))
    return send_file(rutas[formato], as_attachment=True)


if __name__ == "__main__":
    app.run(
        host=os.environ.get("FLASK_HOST", "127.0.0.1"),
        port=int(os.environ.get("FLASK_PORT", "5000")),
        debug=False,
        threaded=True,
    )
