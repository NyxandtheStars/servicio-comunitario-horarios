import os
import re
import unicodedata
import openpyxl
import docx
from database import (
    get_connection,
    get_or_create_facultad,
    get_or_create_carrera,
    get_or_create_profesor,
    get_or_create_materia,
    get_or_create_salon,
    BLOQUES_HORARIOS,
)


MAPA_DIAS = {
    "lunes": "Lunes",
    "lun": "Lunes",
    "martes": "Martes",
    "mar": "Martes",
    "miercoles": "Miércoles",
    "miércoles": "Miércoles",
    "miiercoles": "Miércoles",
    "mier": "Miércoles",
    "mie": "Miércoles",
    "mié": "Miércoles",
    "jueves": "Jueves",
    "jue": "Jueves",
    "viernes": "Viernes",
    "vie": "Viernes",
    "sabado": "Sábado",
    "sábado": "Sábado",
    "sab": "Sábado",
    "sáb": "Sábado",
    "domingo": "Domingo",
    "dom": "Domingo",
}


def normalizar_texto(texto):
    if not texto:
        return ""
    texto = str(texto).strip().lower()
    return "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )


def normalizar_hora(h_str, ampm=""):
    """
    Convierte cadenas como '07:40', '02:10', '03.40', "1'0:50" en formato 24h 'HH:MM'.
    Tiene en cuenta AM/PM o deduce horario vespertino universitario (01:00 a 06:59 -> 13:00 a 18:59).
    """
    h_str = str(h_str).strip().replace(".", ":").replace("'", "")
    numeros = re.findall(r"\d+", h_str)
    if not numeros:
        return "07:00"

    h = int(numeros[0])
    m = int(numeros[1][:2]) if len(numeros) >= 2 else 0

    ampm_clean = normalizar_texto(ampm)
    if "am" in ampm_clean and h == 12:
        h = 0
    elif "pm" in ampm_clean and "am" not in ampm_clean:
        if 1 <= h <= 11:
            h += 12
    elif not ampm_clean and 1 <= h <= 6:
        # En horarios universitarios, las horas 01:xx a 06:xx corresponden a la tarde (13:xx a 18:xx)
        h += 12

    return f"{h:02d}:{m:02d}"


def expandir_rango_a_bloques(h_ini, h_fin):
    """
    Si es un horario exacto (por ejemplo 07:40 a 09:20 o 09:20 a 10:50), lo preserva tal cual.
    Solo divide en bloques estándar de 2h si es un rango amplio genérico (ej. 07:00 a 13:00).
    """
    bloques = []
    for b_ini, b_fin in BLOQUES_HORARIOS:
        if b_ini >= h_ini and b_fin <= h_fin:
            bloques.append((b_ini, b_fin))
    # Si el horario tiene minutos específicos (ej. 07:40 a 09:20, 10:50 a 12:10) o abarca un solo bloque,
    # conservar el rango exacto del archivo para que el calendario refleje la hora real de la universidad.
    if len(bloques) <= 1 or not (h_ini.endswith(":00") and h_fin.endswith(":00")):
        return [(h_ini, h_fin)]
    return bloques


def parsear_disponibilidad(texto_disp, ampm=""):
    """
    Parsea cadenas como:
    - "Lunes 07:00-11:00; Miércoles 09:00-13:00"
    - "MIERCOLES 03:40 a 05:00" (con AM/PM)
    - "MARTES MIERCOLES JUEVES 02:10 a 05:00"
    Devuelve lista de tuplas: [(dia, hora_inicio, hora_fin), ...]
    """
    if not texto_disp:
        return []

    texto_limpio = str(texto_disp).strip()
    if normalizar_texto(texto_limpio) in ("todos", "todo", "disponible", "completo", "tiempo completo"):
        res = []
        for d in ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]:
            for b_ini, b_fin in BLOQUES_HORARIOS[:4]:
                res.append((d, b_ini, b_fin))
        return res

    # Limpiar errores de tipeo comunes en horas como "03.40" o "1'0:50"
    texto_limpio = re.sub(r"(\d)\'(\d)", r"\1\2", texto_limpio)
    texto_limpio = re.sub(r"(\d{1,2})\.(\d{2})", r"\1:\2", texto_limpio)

    resultados = []
    segmentos = re.split(r"[;\n|/]+", texto_limpio)

    patron_dia_str = r"(?:lunes|martes|mi[eéi]+rcoles|m[ií]ercoles|jueves|viernes|s[aá]bado|domingo|lun|mar|mi[eé]|jue|vie|s[aá]b|dom)"

    for seg in segmentos:
        seg = seg.strip()
        if not seg:
            continue

        dias_encontrados = re.findall(patron_dia_str, normalizar_texto(seg), re.IGNORECASE)
        rangos_hora = re.findall(
            r"(\d{1,2}(?::\d{2})?)\s*(?:-|a|hasta|–)\s*(\d{1,2}(?::\d{2})?)",
            seg,
            re.IGNORECASE
        )

        if dias_encontrados and rangos_hora:
            # Si hay la misma cantidad de días y rangos (ej. "MARTES MIERCOLES JUEVES" y 3 rangos), emparejarlos
            if len(dias_encontrados) == len(rangos_hora) and len(dias_encontrados) > 1:
                pares = zip(dias_encontrados, rangos_hora)
                for d_raw, (r_ini, r_fin) in pares:
                    dia_canon = MAPA_DIAS.get(normalizar_texto(d_raw), "Lunes")
                    h_ini = normalizar_hora(r_ini, ampm)
                    h_fin = normalizar_hora(r_fin, ampm)
                    if h_fin <= h_ini and int(h_fin.split(":")[0]) < 12:
                        h_fin = f"{int(h_fin.split(':')[0]) + 12:02d}:{h_fin.split(':')[1]}"
                    for b_ini, b_fin in expandir_rango_a_bloques(h_ini, h_fin):
                        if (dia_canon, b_ini, b_fin) not in resultados:
                            resultados.append((dia_canon, b_ini, b_fin))
            else:
                for d_raw in dias_encontrados:
                    dia_canon = MAPA_DIAS.get(normalizar_texto(d_raw), "Lunes")
                    for r_ini, r_fin in rangos_hora:
                        h_ini = normalizar_hora(r_ini, ampm)
                        h_fin = normalizar_hora(r_fin, ampm)
                        if h_fin <= h_ini and int(h_fin.split(":")[0]) < 12:
                            h_fin = f"{int(h_fin.split(':')[0]) + 12:02d}:{h_fin.split(':')[1]}"
                        for b_ini, b_fin in expandir_rango_a_bloques(h_ini, h_fin):
                            if (dia_canon, b_ini, b_fin) not in resultados:
                                resultados.append((dia_canon, b_ini, b_fin))
    return resultados


def mapear_columnas(encabezados):
    """
    Detecta qué índice de columna corresponde a cada campo académico.
    Soporta tanto el formato de plantilla simple como el formato completo de control de estudios:
    (N°, NÚCLEO, NOMBRES Y APELLIDOS, CARRERA, MENCIÓN, CÓDIGO, NOMBRE DEL CURSO, CICLO, SECCIÓN, AMBIENTE, DÍA, HORARIO, AM/PM).
    """
    mapa = {}
    for idx, col in enumerate(encabezados):
        c = normalizar_texto(col)
        if not c:
            continue

        # Ignorar columnas de observaciones como "COMPARTIDOS CON EL PLAN DE ESTUDIOS..."
        if "compartid" in c or "plan de estudio" in c:
            continue

        if c in ("am/pm", "ampm", "turno"):
            mapa.setdefault("ampm", idx)
        elif c in ("dia", "dias", "dia de clase"):
            mapa.setdefault("dia_col", idx)
        elif c in ("horario", "hora", "horas"):
            mapa.setdefault("horario_col", idx)
        elif any(k in c for k in ["disponibilidad", "horario disponible", "disponible"]):
            mapa.setdefault("disponibilidad", idx)
        elif any(k in c for k in ["nombre del curso", "unidad curricular", "asignatura", "materia", "catedra", "curso"]):
            mapa.setdefault("materia", idx)
        elif any(k in c for k in ["nombres y apellidos", "profesor", "docente", "facilitador", "instructor"]):
            mapa.setdefault("profesor", idx)
        elif c == "carrera" or any(k in c for k in ["programa", "especialidad"]):
            mapa.setdefault("carrera", idx)
        elif "mencion" in c:
            mapa.setdefault("mencion", idx)
        elif any(k in c for k in ["facultad", "decana", "nucleo", "escuela", "sede"]):
            mapa.setdefault("facultad", idx)
        elif any(k in c for k in ["salon", "aula", "laboratorio", "ambiente"]):
            mapa.setdefault("salon", idx)
        elif any(k in c for k in ["seccion", "grupo"]):
            mapa.setdefault("seccion", idx)
        elif any(k in c for k in ["semestre", "trimestre", "nivel", "ciclo", "ano"]):
            mapa.setdefault("semestre", idx)
        elif any(k in c for k in ["codigo", "cod"]):
            mapa.setdefault("codigo_materia", idx)
        elif any(k in c for k in ["bloque", "frecuencia"]):
            mapa.setdefault("bloques", idx)
        elif c in ("cedula", "ci", "documento"):
            mapa.setdefault("cedula", idx)

    return mapa


def procesar_filas_datos(filas, nombre_hoja=""):
    """
    Recibe una lista de filas de una hoja o tabla e inserta en la base de datos.
    Busca automáticamente en qué fila está el encabezado real.
    """
    if not filas:
        return {"insertados": 0, "errores": []}

    # Buscar en las primeras 10 filas cuál es la fila de encabezado
    fila_hdr_idx = 0
    mapa = {}
    tiene_encabezado = False

    for idx_candidato in range(min(10, len(filas))):
        fila_cand = [str(x).strip() if x is not None else "" for x in filas[idx_candidato]]
        mapa_cand = mapear_columnas(fila_cand)
        if "profesor" in mapa_cand and ("materia" in mapa_cand or "carrera" in mapa_cand or "mencion" in mapa_cand):
            fila_hdr_idx = idx_candidato
            mapa = mapa_cand
            tiene_encabezado = True
            break

    inicio_datos = fila_hdr_idx + 1 if tiene_encabezado else 0

    if not tiene_encabezado:
        primera = [str(x).strip() if x is not None else "" for x in filas[0]]
        if len(primera) >= 5:
            mapa = {
                "facultad": 0,
                "carrera": 1,
                "profesor": 2,
                "materia": 3,
                "disponibilidad": 4,
            }
            if len(primera) >= 6:
                mapa["salon"] = 5
            if len(primera) >= 7:
                mapa["seccion"] = 6
        elif len(primera) == 4:
            mapa = {
                "carrera": 0,
                "profesor": 1,
                "materia": 2,
                "disponibilidad": 3,
            }
        else:
            return {
                "insertados": 0,
                "errores": [
                    f"Hoja '{nombre_hoja}': No se pudieron identificar las columnas."
                ]
            }

    conn = get_connection()
    cursor = conn.cursor()
    insertados = 0
    errores = []

    for num_fila, fila in enumerate(filas[inicio_datos:], start=inicio_datos + 1):
        celdas = [str(x).strip() if x is not None else "" for x in fila]
        if not any(celdas):
            continue

        def obtener(campo, defecto=""):
            idx = mapa.get(campo)
            if idx is not None and idx < len(celdas) and celdas[idx] and celdas[idx].lower() != "none":
                return re.sub(r"\s+", " ", celdas[idx]).strip()
            return defecto

        profesor_nom = obtener("profesor", "")
        materia_raw = obtener("materia", "")

        # Ignorar filas repetidas de encabezado o filas vacías de cierre
        if not profesor_nom and not materia_raw:
            continue
        if normalizar_texto(profesor_nom) in ("nombres y apellidos", "profesor", "docente"):
            continue
        if not profesor_nom or not materia_raw:
            errores.append(f"{nombre_hoja + ' - ' if nombre_hoja else ''}Fila {num_fila}: Falta nombre del profesor o materia.")
            continue

        facultad_nom = obtener("facultad", "")
        carrera_nom = obtener("carrera", "")
        mencion_nom = obtener("mencion", "")

        # Construir nombre de Carrera combinando Carrera y Mención si existen
        if carrera_nom and mencion_nom and normalizar_texto(mencion_nom) not in ("sin mencion", normalizar_texto(carrera_nom)):
            carrera_final = f"{carrera_nom} - {mencion_nom}"
        elif carrera_nom:
            carrera_final = carrera_nom
        elif mencion_nom and normalizar_texto(mencion_nom) != "sin mencion":
            carrera_final = mencion_nom
        elif nombre_hoja:
            carrera_final = nombre_hoja
        else:
            carrera_final = "Carrera General"

        if not facultad_nom:
            facultad_nom = f"Núcleo Principal ({nombre_hoja})" if nombre_hoja else "Facultad / Núcleo Universitario"
        elif carrera_nom and normalizar_texto(facultad_nom) == "palo verde":
            facultad_nom = f"Núcleo {facultad_nom.title()} - {carrera_nom.title()}"

        salon_pref = obtener("salon", "")
        seccion = obtener("seccion", "Sección Única")
        semestre = obtener("semestre", "General")
        codigo_mat = obtener("codigo_materia", "")
        cedula = obtener("cedula", "")
        bloques_str = obtener("bloques", "1")

        # Construir disponibilidad/horario a partir de DÍA + HORARIO + AM/PM o columna Disponibilidad
        disp_raw = obtener("disponibilidad", "")
        dia_col = obtener("dia_col", "")
        horario_col = obtener("horario_col", "")
        ampm_col = obtener("ampm", "")

        if dia_col or horario_col:
            disp_combinada = f"{dia_col} {horario_col}".strip()
            bloques_disp = parsear_disponibilidad(disp_combinada, ampm=ampm_col)
        else:
            bloques_disp = parsear_disponibilidad(disp_raw, ampm=ampm_col)

        try:
            bloques_sem = max(1, min(4, int(re.findall(r"\d+", str(bloques_str))[0])))
        except Exception:
            bloques_sem = max(1, len(bloques_disp)) if (dia_col and len(bloques_disp) > 1) else 1

        fac_id = get_or_create_facultad(cursor, facultad_nom)
        carreras_lista = [c.strip() for c in re.split(r"[/|]+", carrera_final) if c.strip()] or [carrera_final]
        prof_id = get_or_create_profesor(cursor, profesor_nom, cedula=cedula)

        if not bloques_disp:
            bloques_disp = [
                ("Lunes", "07:00", "09:00"),
                ("Martes", "07:00", "09:00"),
                ("Miércoles", "07:00", "09:00"),
                ("Jueves", "09:00", "11:00"),
                ("Viernes", "07:00", "09:00"),
            ]

        for dia, h_ini, h_fin in bloques_disp:
            cursor.execute(
                """
                INSERT OR IGNORE INTO disponibilidades (profesor_id, dia, hora_inicio, hora_fin)
                VALUES (?, ?, ?, ?)
                """,
                (prof_id, dia, h_ini, h_fin)
            )

        if salon_pref:
            get_or_create_salon(cursor, salon_pref)

        # Si hay columna DÍA/HORARIO específica de la materia, no separar el nombre del curso por comas
        # (porque hay materias llamadas "ÉTICA, VALORES Y AUTOBIOGRAFÍA" o "ACTIVIDADES DEPORTIVAS, CULTURALES Y RECREATIVAS")
        if "dia_col" in mapa or "codigo_materia" in mapa:
            materias_lista = [materia_raw.strip()]
        else:
            materias_lista = [m.strip() for m in re.split(r"[,;/]+", materia_raw) if m.strip()]

        dia_fijo = bloques_disp[0][0] if (dia_col and bloques_disp) else ""
        h_ini_fija = bloques_disp[0][1] if (dia_col and bloques_disp) else ""
        h_fin_fija = bloques_disp[0][2] if (dia_col and bloques_disp) else ""

        for car_nom in carreras_lista:
            car_id = get_or_create_carrera(cursor, car_nom, fac_id)
            for mat_nom in materias_lista:
                mat_id = get_or_create_materia(cursor, mat_nom, car_id, codigo=codigo_mat, semestre=semestre)
                cursor.execute(
                    """
                    INSERT INTO carga_academica
                    (profesor_id, materia_id, seccion, bloques_semanales, salon_preferido, dia_fijo, hora_inicio_fija, hora_fin_fija)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(profesor_id, materia_id, seccion) DO UPDATE SET
                        bloques_semanales = excluded.bloques_semanales,
                        salon_preferido = CASE WHEN excluded.salon_preferido != '' THEN excluded.salon_preferido ELSE carga_academica.salon_preferido END,
                        dia_fijo = CASE WHEN excluded.dia_fijo != '' THEN excluded.dia_fijo ELSE carga_academica.dia_fijo END,
                        hora_inicio_fija = CASE WHEN excluded.hora_inicio_fija != '' THEN excluded.hora_inicio_fija ELSE carga_academica.hora_inicio_fija END,
                        hora_fin_fija = CASE WHEN excluded.hora_fin_fija != '' THEN excluded.hora_fin_fija ELSE carga_academica.hora_fin_fija END
                    """,
                    (prof_id, mat_id, seccion, bloques_sem, salon_pref, dia_fijo, h_ini_fija, h_fin_fija)
                )
                insertados += 1

    conn.commit()
    conn.close()
    return {"insertados": insertados, "errores": errores}


def importar_excel(filepath):
    """
    Procesa cada hoja del libro Excel de forma independiente para que las diferencias
    de columnas entre pestañas (ej. RRHH, RMF, INTEG, MATEM) se mapeen con exactitud.
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    total_insertados = 0
    todos_errores = []

    for sheet in wb.sheetnames:
        ws = wb[sheet]
        filas = []
        for row in ws.iter_rows(values_only=True):
            if any(cell is not None and str(cell).strip() != "" for cell in row):
                filas.append(list(row))
        if filas:
            res = procesar_filas_datos(filas, nombre_hoja=sheet)
            total_insertados += res["insertados"]
            todos_errores.extend(res["errores"])

    return {"insertados": total_insertados, "errores": todos_errores}


def importar_docx(filepath):
    doc = docx.Document(filepath)
    total_insertados = 0
    todos_errores = []

    if doc.tables:
        for idx, table in enumerate(doc.tables, 1):
            filas = []
            for row in table.rows:
                celdas = [cell.text.strip() for cell in row.cells]
                if any(celdas):
                    filas.append(celdas)
            if filas:
                res = procesar_filas_datos(filas, nombre_hoja=f"Tabla {idx}")
                total_insertados += res["insertados"]
                todos_errores.extend(res["errores"])
        return {"insertados": total_insertados, "errores": todos_errores}

    lineas = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return parsear_lineas_texto(lineas)


def importar_txt(filepath):
    with open(filepath, "r", encoding="utf-8", errors="replace") as f:
        lineas = [line.strip() for line in f if line.strip() and not line.strip().startswith("#")]
    return parsear_lineas_texto(lineas)


def parsear_lineas_texto(lineas):
    if not lineas:
        return {"insertados": 0, "errores": ["El archivo de texto está vacío."]}

    import csv
    from io import StringIO

    primera = lineas[0]
    if "|" in primera:
        delim = "|"
    elif "\t" in primera:
        delim = "\t"
    elif ";" in primera and primera.count(";") >= 3:
        delim = ";"
    else:
        delim = ","

    lector = csv.reader(StringIO("\n".join(lineas)), delimiter=delim)
    todas_filas = [[c.strip() for c in row] for row in lector if any(c.strip() for c in row)]

    if not todas_filas:
        return {"insertados": 0, "errores": ["No se encontraron filas válidas."]}

    # Si un archivo TXT/CSV contiene múltiples tablas concatenadas (con sus propios encabezados),
    # separamos en bloques cada vez que aparece una nueva fila de encabezado.
    bloques_tablas = []
    bloque_actual = []

    for fila in todas_filas:
        mapa_cand = mapear_columnas(fila)
        es_encabezado = (
            "profesor" in mapa_cand
            and ("materia" in mapa_cand or "carrera" in mapa_cand or "mencion" in mapa_cand)
            and any(normalizar_texto(c) in ("nombres y apellidos", "profesor", "docente", "nombre del curso", "materia") for c in fila)
        )
        if es_encabezado and bloque_actual:
            bloques_tablas.append(bloque_actual)
            bloque_actual = [fila]
        else:
            bloque_actual.append(fila)

    if bloque_actual:
        bloques_tablas.append(bloque_actual)

    total_insertados = 0
    todos_errores = []
    for idx, bloque in enumerate(bloques_tablas, 1):
        res = procesar_filas_datos(bloque, nombre_hoja=f"Sección {idx}")
        total_insertados += res["insertados"]
        todos_errores.extend(res["errores"])

    return {"insertados": total_insertados, "errores": todos_errores}


# ═══════════════════════════════════════════════════════════════════════════
