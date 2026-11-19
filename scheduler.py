from database import (
    get_connection,
    DIAS_ORDEN,
    BLOQUES_HORARIOS,
    agrupar_franjas_disponibilidad,
    calcular_horas_academicas,
)


def hay_solapamiento(h_ini1, h_fin1, h_ini2, h_fin2):
    """Verifica si dos rangos de hora HH:MM se solapan."""
    return max(h_ini1, h_ini2) < min(h_fin1, h_fin2)


def es_ambiente_compartible(codigo_salon):
    """
    Ambientes virtuales o coordinaciones permiten múltiples cursos simultáneos sin choque físico.
    """
    c = (codigo_salon or "").upper()
    return any(k in c for k in ["VIRTUAL", "COORDINAC", "ONLINE", "CANCHA", "AUDITORIO", "ASESORIA"])


def _construir_candidatos_disponibilidad(disponibles_raw, num_cargas_prof: int = 1, horas_objetivo: int = 2):
    """
    Convierte las franjas de disponibilidad suministradas por el profesor en bloques de clase listos para asignar:
    - Si el profesor marcó casillas de 45 min contiguas (ej. 07:00-07:45 y 07:45-08:30), las une en un
      bloque continuo (07:00-08:30).
    - Si el profesor marcó una ventana larga en un mismo día (ej. 4 o 6 casillas contiguas) para impartir
      varias materias (num_cargas_prof > 1), divide esa ventana en sub-bloques del tamaño requerido
      (horas_objetivo casillas de 45 min, mínimo 2 casillas = 90 min) para que todas sus materias puedan
      asignarse dentro de su horario disponible sin solaparse.
    """
    if not disponibles_raw:
        return []

    agrupados = agrupar_franjas_disponibilidad(disponibles_raw)
    if not agrupados:
        return [(d, hi, hf) for d, hi, hf in disponibles_raw]

    candidatos = []
    paso_casillas = max(2, min(4, int(horas_objetivo or 2)))

    for dia, win_ini, win_fin, num_horas, sub_lista in agrupados:
        # Si hay más materias del profesor que ventanas continuas y esta ventana tiene varias sub-casillas,
        # generar tanto sub-bloques del tamaño de la materia como la ventana completa.
        if num_cargas_prof > len(agrupados) and len(sub_lista) >= paso_casillas * 2:
            idx = 0
            while idx + paso_casillas <= len(sub_lista):
                sub_ini = sub_lista[idx][0]
                sub_fin = sub_lista[idx + paso_casillas - 1][1]
                if (dia, sub_ini, sub_fin) not in candidatos:
                    candidatos.append((dia, sub_ini, sub_fin))
                idx += paso_casillas
            if idx < len(sub_lista):
                sub_ini = sub_lista[idx][0]
                sub_fin = sub_lista[-1][1]
                if (dia, sub_ini, sub_fin) not in candidatos:
                    candidatos.append((dia, sub_ini, sub_fin))
        elif len(sub_lista) > paso_casillas and num_cargas_prof > 1:
            # Añadir primero un bloque ajustado a las horas académicas de la materia y luego la ventana completa
            sub_ini = sub_lista[0][0]
            sub_fin = sub_lista[paso_casillas - 1][1]
            if (dia, sub_ini, sub_fin) not in candidatos:
                candidatos.append((dia, sub_ini, sub_fin))
            rest_ini = sub_lista[paso_casillas][0]
            rest_fin = sub_lista[-1][1]
            if (dia, rest_ini, rest_fin) not in candidatos:
                candidatos.append((dia, rest_ini, rest_fin))

        if (dia, win_ini, win_fin) not in candidatos:
            candidatos.append((dia, win_ini, win_fin))

    # Agregar también las franjas individuales originales como respaldo final si hiciera falta
    for d, hi, hf in disponibles_raw:
        if (d, hi, hf) not in candidatos:
            candidatos.append((d, hi, hf))

    return candidatos


def _seleccionar_salon_libre(salones, ocupado_salon, dia, h_ini, h_fin, salon_pref_id=None, carrera_nombre=""):
    """
    Elige automáticamente un ambiente libre en (dia, h_ini, h_fin):
    1. Si el docente indicó un ambiente preferido y está libre (o es compartible), lo respeta.
    2. Si no indicó ambiente (asignación automática) o el preferido está ocupado, asigna un aula física
       libre (P1..P4, PB, SATELITAL), priorizando las aulas con menor carga para distribuir equitativamente
       y evitando enviar clases presenciales a VIRTUAL/COORDINACIÓN salvo que así se haya solicitado.
    """
    if salon_pref_id is not None:
        for s in salones:
            if s["id"] == salon_pref_id:
                if es_ambiente_compartible(s["codigo"]):
                    return s["id"]
                choque = any(
                    d == dia and hay_solapamiento(h_ini, h_fin, oi, of)
                    for d, oi, of, _ in ocupado_salon.get(s["id"], [])
                )
                if not choque:
                    return s["id"]
                break

    es_fisio = "FISIOTERAPIA" in (carrera_nombre or "").upper()

    def _prioridad_salon(s):
        cod = (s["codigo"] or "").upper()
        comp = es_ambiente_compartible(cod)
        es_aula_piso = cod.startswith(("P1-", "P2-", "P3-", "P4-", "PB-"))
        camillas_match = 0 if (es_fisio and int(s.get("es_camillas") or 0)) else 1
        uso_actual = len(ocupado_salon.get(s["id"], []))
        return (
            1 if comp else 0,
            camillas_match,
            0 if es_aula_piso else 1,
            uso_actual,
            s["id"],
        )

    candidatos_ordenados = sorted(salones, key=_prioridad_salon)
    for s in candidatos_ordenados:
        sid = s["id"]
        if es_ambiente_compartible(s["codigo"]):
            continue
        choque = any(
            d == dia and hay_solapamiento(h_ini, h_fin, oi, of)
            for d, oi, of, _ in ocupado_salon.get(sid, [])
        )
        if not choque:
            return sid

    # Último recurso si todas las aulas físicas están llenas en esa franja
    for s in candidatos_ordenados:
        if es_ambiente_compartible(s["codigo"]):
            return s["id"]
    return None


def generar_horarios_automaticos():
    """
    Algoritmo inteligente de planificación universitaria:
    - Si la carga académica ya trae día y hora específicos desde el Excel oficial (dia_fijo, hora_inicio_fija),
      prioriza asignarla exactamente en ese día, hora y ambiente indicado (permitiendo cursos compartidos entre
      carreras/menciones del mismo profesor o ambientes virtuales).
    - Si es una carga basada en disponibilidad general, asigna los bloques semanales evitando choques de
      profesor y salón.
    """
    conn = get_connection()
    cursor = conn.cursor()

    # Limpiar asignaciones automáticas previas
    cursor.execute("DELETE FROM asignaciones_horario;")

    # Obtener todos los salones disponibles
    cursor.execute("SELECT id, codigo, capacidad, edificio FROM salones ORDER BY id ASC")
    salones = [dict(r) for r in cursor.fetchall()]
    if not salones:
        conn.close()
        return {"asignados": 0, "pendientes": 0, "mensajes": ["No hay salones registrados en el sistema."]}

    mapa_salones_codigo = {s["codigo"].lower(): s["id"] for s in salones}
    mapa_salones_id_codigo = {s["id"]: s["codigo"] for s in salones}

    # Obtener disponibilidades por profesor
    cursor.execute("SELECT profesor_id, dia, hora_inicio, hora_fin FROM disponibilidades")
    disp_por_profesor = {}
    for r in cursor.fetchall():
        pid = r["profesor_id"]
        disp_por_profesor.setdefault(pid, []).append(
            (r["dia"], r["hora_inicio"], r["hora_fin"])
        )

    orden_dia_idx = {d: i for i, d in enumerate(DIAS_ORDEN)}
    for pid in disp_por_profesor:
        disp_por_profesor[pid].sort(
            key=lambda x: (orden_dia_idx.get(x[0], 99), x[1])
        )

    # Obtener todas las cargas académicas con datos de profesor, materia, carrera y facultad
    cursor.execute("""
        SELECT
            ca.id AS carga_id,
            ca.profesor_id,
            p.nombre AS profesor_nombre,
            ca.materia_id,
            m.nombre AS materia_nombre,
            COALESCE(m.unidad_credito, 0) AS unidad_credito,
            COALESCE(m.horas_academicas, 0) AS horas_academicas,
            c.nombre AS carrera_nombre,
            f.nombre AS facultad_nombre,
            ca.seccion,
            ca.bloques_semanales,
            ca.salon_preferido,
            ca.dia_fijo,
            ca.hora_inicio_fija,
            ca.hora_fin_fija,
            COALESCE(ca.compartidos, '') AS compartidos
        FROM carga_academica ca
        JOIN profesores p ON p.id = ca.profesor_id
        JOIN materias m ON m.id = ca.materia_id
        JOIN carreras c ON c.id = m.carrera_id
        JOIN facultades f ON f.id = c.facultad_id
        ORDER BY p.nombre ASC, f.nombre ASC, c.nombre ASC, m.nombre ASC
    """)
    cargas = [dict(r) for r in cursor.fetchall()]

    conteo_cargas_prof = {}
    for c in cargas:
        conteo_cargas_prof[c["profesor_id"]] = conteo_cargas_prof.get(c["profesor_id"], 0) + 1

    # Priorizar las cargas que tienen horario fijo del archivo oficial, y luego por MRV
    cargas.sort(
        key=lambda c: (
            0 if c["dia_fijo"] else 1,
            len(disp_por_profesor.get(c["profesor_id"], [])),
            c["profesor_nombre"],
        )
    )

    # ocupado_profesor[profesor_id] = lista de (dia, h_ini, h_fin, salon_id)
    ocupado_profesor = {}
    # ocupado_salon[salon_id] = lista de (dia, h_ini, h_fin, profesor_id)
    ocupado_salon = {}

    asignados = 0
    pendientes = 0
    mensajes = []

    for carga in cargas:
        cid = carga["carga_id"]
        pid = carga["profesor_id"]
        mid = carga["materia_id"]
        horas_obj = calcular_horas_academicas(
            carga.get("hora_inicio_fija", ""),
            carga.get("hora_fin_fija", ""),
            carga.get("unidad_credito", 0),
            carga.get("horas_academicas", 0),
            carga.get("bloques_semanales", 1),
        )
        salon_pref_str = (carga["salon_preferido"] or "").strip().lower()
        salon_pref_id = mapa_salones_codigo.get(salon_pref_str)

        # Caso 1: La fila del archivo ya traía su Día y Hora exactos (ej. Excel de oferta de universidad)
        if carga["dia_fijo"] and carga["hora_inicio_fija"] and carga["hora_fin_fija"]:
            dia_f = carga["dia_fijo"]
            hi_f = carga["hora_inicio_fija"]
            hf_f = carga["hora_fin_fija"]
            # Si este bloque fijo es muy largo (ej. 09:20-13:00) y solapa con otros bloques fijos más cortos del mismo profesor en el mismo día,
            # recortarlo hasta el inicio del siguiente bloque fijo del docente para no bloquear sus demás materias.
            otros_fijos_mismo_dia = [
                (c_otro["hora_inicio_fija"], c_otro["hora_fin_fija"])
                for c_otro in cargas
                if c_otro["profesor_id"] == pid
                and c_otro["carga_id"] != cid
                and c_otro.get("dia_fijo") == dia_f
                and c_otro.get("hora_inicio_fija")
                and c_otro.get("hora_fin_fija")
            ]
            inicios_posteriores = [
                oi for oi, of in otros_fijos_mismo_dia
                if str(hi_f) < str(oi) < str(hf_f)
            ]
            if inicios_posteriores:
                hf_f = min(inicios_posteriores)

            prof_choque = any(
                d == dia_f and hay_solapamiento(hi_f, hf_f, oi, of)
                for d, oi, of, _ in ocupado_profesor.get(pid, [])
            )
            salon_elegido = None
            if not prof_choque:
                salon_elegido = _seleccionar_salon_libre(
                    salones,
                    ocupado_salon,
                    dia_f,
                    hi_f,
                    hf_f,
                    salon_pref_id=salon_pref_id,
                    carrera_nombre=carga.get("carrera_nombre", ""),
                )

            if salon_elegido is not None:
                horas_reales = calcular_horas_academicas(
                    hi_f, hf_f, carga.get("unidad_credito", 0), carga.get("horas_academicas", 0), horas_obj
                )
                cursor.execute(
                    """
                    INSERT INTO asignaciones_horario
                    (carga_id, profesor_id, materia_id, salon_id, dia, hora_inicio, hora_fin, estado)
                    VALUES (?, ?, ?, ?, ?, ?, ?, 'Confirmado')
                    """,
                    (cid, pid, mid, salon_elegido, dia_f, hi_f, hf_f)
                )
                if horas_reales != carga.get("bloques_semanales"):
                    cursor.execute("UPDATE carga_academica SET bloques_semanales = ? WHERE id = ?", (horas_reales, cid))
                if not salon_pref_str and salon_elegido in mapa_salones_id_codigo:
                    cursor.execute(
                        "UPDATE carga_academica SET salon_preferido = ? WHERE id = ?",
                        (mapa_salones_id_codigo[salon_elegido], cid),
                    )
                ocupado_profesor.setdefault(pid, []).append((dia_f, hi_f, hf_f, salon_elegido))
                ocupado_salon.setdefault(salon_elegido, []).append((dia_f, hi_f, hf_f, pid))
                asignados += 1
                continue
            # Si hubo choque en el horario fijo pero el profesor tiene otros horarios disponibles en disponibilidades,
            # intentamos asignarlo automáticamente en Caso 2 usando su disponibilidad registrada.

        # Caso 2: Asignación automática de horario y ambiente de acuerdo a los horarios disponibles suministrados por el profesor
        disponibles_raw = disp_por_profesor.get(pid, [])
        if not disponibles_raw:
            pendientes += 1
            mensajes.append(
                f"Sin disponibilidad registrada: '{carga['materia_nombre']}' "
                f"({carga['carrera_nombre']}) con el Prof. {carga['profesor_nombre']} "
                f"no tiene bloques horarios para asignar."
            )
            continue

        candidatos_disp = _construir_candidatos_disponibilidad(
            disponibles_raw,
            num_cargas_prof=conteo_cargas_prof.get(pid, 1),
            horas_objetivo=horas_obj,
        )
        # Si el profesor tiene más materias que franjas distintas suministradas en su día activo,
        # agregar como respaldo los bloques estándar de ese mismo día para no dejar ninguna materia sin horario/ambiente.
        dias_activos_prof = list(
            dict.fromkeys(
                (d[0] if isinstance(d, (tuple, list)) else d.get("dia"))
                for d in disponibles_raw
                if (d[0] if isinstance(d, (tuple, list)) else d.get("dia"))
            )
        )
        for d_act in dias_activos_prof:
            for b_ini, b_fin in BLOQUES_HORARIOS:
                par = (d_act, b_ini, b_fin)
                if par not in candidatos_disp:
                    candidatos_disp.append(par)

        asignado_en_disp = False
        for dia, h_ini, h_fin in candidatos_disp:
            prof_choque = any(
                d == dia and hay_solapamiento(h_ini, h_fin, oi, of)
                for d, oi, of, _ in ocupado_profesor.get(pid, [])
            )
            if prof_choque:
                continue

            salon_elegido = _seleccionar_salon_libre(
                salones,
                ocupado_salon,
                dia,
                h_ini,
                h_fin,
                salon_pref_id=salon_pref_id,
                carrera_nombre=carga.get("carrera_nombre", ""),
            )

            if salon_elegido is not None:
                horas_reales = calcular_horas_academicas(
                    h_ini, h_fin, carga.get("unidad_credito", 0), carga.get("horas_academicas", 0), horas_obj
                )
                cursor.execute(
                    """
                    INSERT INTO asignaciones_horario
                    (carga_id, profesor_id, materia_id, salon_id, dia, hora_inicio, hora_fin, estado)
                    VALUES (?, ?, ?, ?, ?, ?, ?, 'Confirmado')
                    """,
                    (cid, pid, mid, salon_elegido, dia, h_ini, h_fin)
                )
                codigo_asignado = mapa_salones_id_codigo.get(salon_elegido, "")
                cursor.execute(
                    """UPDATE carga_academica
                       SET bloques_semanales = ?,
                           salon_preferido = CASE WHEN TRIM(COALESCE(salon_preferido, '')) = '' THEN ? ELSE salon_preferido END
                       WHERE id = ?""",
                    (horas_reales, codigo_asignado, cid),
                )
                ocupado_profesor.setdefault(pid, []).append((dia, h_ini, h_fin, salon_elegido))
                ocupado_salon.setdefault(salon_elegido, []).append((dia, h_ini, h_fin, pid))
                asignado_en_disp = True
                asignados += 1
                break

        if not asignado_en_disp:
            pendientes += 1
            mensajes.append(
                f"Conflicto de disponibilidad/ambiente: '{carga['materia_nombre']}' "
                f"({carga['carrera_nombre']}) con el Prof. {carga['profesor_nombre']} "
                f"no encontró ambiente libre en las franjas suministradas."
            )

    conn.commit()
    conn.close()
    return {
        "asignados": asignados,
        "pendientes": pendientes,
        "mensajes": mensajes,
    }


def autoasignar_ambientes_libres():
    """Asigna ambientes libres a cargas pendientes según disponibilidad docente y reubica clases con choque."""
    conn_chk = get_connection()
    cur_chk = conn_chk.cursor()
    cur_chk.execute(
        """SELECT COUNT(*) AS sin_asignar
           FROM carga_academica ca
           LEFT JOIN asignaciones_horario ah ON ah.carga_id = ca.id
           WHERE ah.id IS NULL"""
    )
    faltan_cargas = cur_chk.fetchone()["sin_asignar"] > 0
    conn_chk.close()
    if faltan_cargas:
        generar_horarios_automaticos()

    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(
            """SELECT a.id, a.profesor_id, a.salon_id, a.dia, a.hora_inicio, a.hora_fin,
                      s.codigo AS salon_codigo, ca.salon_preferido
               FROM asignaciones_horario a
               JOIN salones s ON s.id = a.salon_id
               LEFT JOIN carga_academica ca ON ca.id = a.carga_id
               ORDER BY a.dia, a.hora_inicio, a.hora_fin, a.id"""
        )
        asignaciones = [dict(row) for row in cursor.fetchall()]
        cursor.execute("SELECT id, codigo FROM salones ORDER BY id")
        salones = [dict(row) for row in cursor.fetchall()]
        if not asignaciones or not salones:
            return {
                "reasignados": 0,
                "conflictos_ambiente": 0,
                "conflictos_docente": 0,
                "mensajes": [],
            }

        mapa_salon_codigo = {salon["id"]: salon["codigo"] for salon in salones}
        for asignacion in asignaciones:
            asignacion["salon_id_final"] = asignacion["salon_id"]

        for _ in range(len(asignaciones) + 1):
            hubo_cambio = False
            por_salon = {}
            for asignacion in asignaciones:
                por_salon.setdefault(asignacion["salon_id_final"], []).append(asignacion)

            for asignacion in asignaciones:
                salon_id_actual = asignacion["salon_id_final"]
                codigo_actual = mapa_salon_codigo.get(salon_id_actual, asignacion["salon_codigo"])
                if es_ambiente_compartible(codigo_actual):
                    continue

                colisionada = any(
                    otra["id"] != asignacion["id"]
                    and otra["id"] < asignacion["id"]
                    and otra["dia"] == asignacion["dia"]
                    and hay_solapamiento(
                        asignacion["hora_inicio"],
                        asignacion["hora_fin"],
                        otra["hora_inicio"],
                        otra["hora_fin"],
                    )
                    for otra in por_salon.get(salon_id_actual, [])
                )
                if not colisionada:
                    continue

                preferido = (asignacion["salon_preferido"] or "").strip().casefold()
                candidatos = sorted(
                    (salon for salon in salones if salon["id"] != salon_id_actual),
                    key=lambda salon: (
                        0 if salon["codigo"].casefold() == preferido and preferido else 1,
                        1 if es_ambiente_compartible(salon["codigo"]) else 0,
                        salon["id"],
                    ),
                )

                nuevo_salon_id = None
                for candidato in candidatos:
                    if es_ambiente_compartible(candidato["codigo"]):
                        libre = True
                    else:
                        libre = not any(
                            otra["id"] != asignacion["id"]
                            and otra["dia"] == asignacion["dia"]
                            and hay_solapamiento(
                                asignacion["hora_inicio"],
                                asignacion["hora_fin"],
                                otra["hora_inicio"],
                                otra["hora_fin"],
                            )
                            for otra in por_salon.get(candidato["id"], [])
                        )
                    if libre:
                        nuevo_salon_id = candidato["id"]
                        break

                if nuevo_salon_id is not None:
                    if asignacion in por_salon.get(salon_id_actual, []):
                        por_salon[salon_id_actual].remove(asignacion)
                    asignacion["salon_id_final"] = nuevo_salon_id
                    por_salon.setdefault(nuevo_salon_id, []).append(asignacion)
                    hubo_cambio = True

            if not hubo_cambio:
                break

        disponibilidad_profesor = {}
        cursor.execute(
            "SELECT profesor_id, dia, hora_inicio, hora_fin FROM disponibilidades"
        )
        for fila in cursor.fetchall():
            disponibilidad_profesor.setdefault(fila["profesor_id"], []).append(
                (fila["dia"], fila["hora_inicio"], fila["hora_fin"])
            )

        horarios_candidatos = [
            (dia, inicio, fin)
            for dia in DIAS_ORDEN
            for inicio, fin in BLOQUES_HORARIOS
        ]

        por_salon_final = {}
        for item in asignaciones:
            por_salon_final.setdefault(item["salon_id_final"], []).append(item)

        ids_horario_modificado = set()
        ids_salon_modificado = set()

        for indice, asignacion in enumerate(asignaciones):
            for otra in asignaciones[indice + 1:]:
                if (
                    asignacion["profesor_id"] != otra["profesor_id"]
                    or asignacion["dia"] != otra["dia"]
                    or not hay_solapamiento(
                        asignacion["hora_inicio"],
                        asignacion["hora_fin"],
                        otra["hora_inicio"],
                        otra["hora_fin"],
                    )
                ):
                    continue

                # Intentar mover `otra`; si no encuentra horario/ambiente libre, intentar mover `asignacion`
                for target, fijada in [(otra, asignacion), (asignacion, otra)]:
                    duracion_actual = (
                        int(target["hora_fin"][:2]) * 60 + int(target["hora_fin"][3:])
                        - int(target["hora_inicio"][:2]) * 60
                        - int(target["hora_inicio"][3:])
                    )
                    disponibles = disponibilidad_profesor.get(target["profesor_id"], [])
                    candidatos_horario = list(disponibles)
                    for h in horarios_candidatos:
                        if h not in candidatos_horario:
                            candidatos_horario.append(h)

                    # Priorizar mismo día y bloques de disponibilidad registrada
                    candidatos_horario.sort(
                        key=lambda x: (
                            0 if x[0] == target["dia"] else 1,
                            0 if x in disponibles else 1,
                        )
                    )

                    nuevo_horario = None
                    for dia, inicio, fin in candidatos_horario:
                        ini_min = int(inicio[:2]) * 60 + int(inicio[3:])
                        fin_min = int(fin[:2]) * 60 + int(fin[3:])
                        fin_candidato = f"{(ini_min + duracion_actual) // 60:02d}:{(ini_min + duracion_actual) % 60:02d}"
                        if fin_min > ini_min and fin_min - ini_min < duracion_actual:
                            continue

                        choque_profesor = any(
                            item["id"] != target["id"]
                            and item["dia"] == dia
                            and hay_solapamiento(inicio, fin_candidato, item["hora_inicio"], item["hora_fin"])
                            for item in asignaciones
                            if item["profesor_id"] == target["profesor_id"]
                        )
                        if choque_profesor:
                            continue

                        # Buscar ambiente libre para este horario
                        salon_elegido_id = None
                        s_act = target["salon_id_final"]
                        cod_act = mapa_salon_codigo.get(s_act, "")
                        if es_ambiente_compartible(cod_act):
                            salon_elegido_id = s_act
                        else:
                            ocupado_act = any(
                                item["id"] != target["id"]
                                and item["dia"] == dia
                                and hay_solapamiento(inicio, fin_candidato, item["hora_inicio"], item["hora_fin"])
                                for item in por_salon_final.get(s_act, [])
                            )
                            if not ocupado_act:
                                salon_elegido_id = s_act

                        # Si el ambiente actual está ocupado en ese horario, probar otros salones registrados
                        if salon_elegido_id is None:
                            for s in salones:
                                sid = s["id"]
                                if sid == s_act:
                                    continue
                                if es_ambiente_compartible(s["codigo"]):
                                    salon_elegido_id = sid
                                    break
                                ocupado_s = any(
                                    item["id"] != target["id"]
                                    and item["dia"] == dia
                                    and hay_solapamiento(inicio, fin_candidato, item["hora_inicio"], item["hora_fin"])
                                    for item in por_salon_final.get(sid, [])
                                )
                                if not ocupado_s:
                                    salon_elegido_id = sid
                                    break

                        if salon_elegido_id is not None:
                            nuevo_horario = (dia, inicio, fin_candidato, salon_elegido_id)
                            break

                    if nuevo_horario:
                        dia_n, ini_n, fin_n, sal_n = nuevo_horario
                        if target in por_salon_final.get(target["salon_id_final"], []):
                            por_salon_final[target["salon_id_final"]].remove(target)
                        target["dia"] = dia_n
                        target["hora_inicio"] = ini_n
                        target["hora_fin"] = fin_n
                        if target["salon_id_final"] != sal_n:
                            target["salon_id_final"] = sal_n
                            ids_salon_modificado.add(target["id"])
                        por_salon_final.setdefault(sal_n, []).append(target)
                        ids_horario_modificado.add(target["id"])
                        break

        # Persistir reubicaciones de ambiente
        cambios_salones = [
            (asignacion["salon_id_final"], asignacion["id"])
            for asignacion in asignaciones
            if asignacion["salon_id_final"] != asignacion["salon_id"]
        ]
        if cambios_salones:
            cursor.executemany(
                "UPDATE asignaciones_horario SET salon_id = ? WHERE id = ?",
                cambios_salones,
            )

        # Persistir reubicaciones de horario
        if ids_horario_modificado:
            cursor.executemany(
                """UPDATE asignaciones_horario
                   SET dia = ?, hora_inicio = ?, hora_fin = ?
                   WHERE id = ?""",
                [
                    (item["dia"], item["hora_inicio"], item["hora_fin"], item["id"])
                    for item in asignaciones
                    if item["id"] in ids_horario_modificado
                ],
            )

        conflictos_ambiente = 0
        conflictos_docente = 0
        for indice, asignacion in enumerate(asignaciones):
            for otra in asignaciones[indice + 1:]:
                if asignacion["dia"] != otra["dia"] or not hay_solapamiento(
                    asignacion["hora_inicio"],
                    asignacion["hora_fin"],
                    otra["hora_inicio"],
                    otra["hora_fin"],
                ):
                    continue
                if asignacion["profesor_id"] == otra["profesor_id"]:
                    conflictos_docente += 1
                mismo_salon = (
                    asignacion["salon_id_final"] == otra["salon_id_final"]
                )
                codigo_salon = mapa_salon_codigo.get(
                    asignacion["salon_id_final"],
                    asignacion["salon_codigo"],
                )
                if (
                    mismo_salon
                    and not es_ambiente_compartible(codigo_salon)
                ):
                    conflictos_ambiente += 1

        conn.commit()
        mensajes = []
        if conflictos_ambiente:
            mensajes.append(
                f"Quedaron {conflictos_ambiente} colisión(es) de ambiente sin resolver: "
                "no hay suficientes ambientes libres en esos horarios."
            )
        if conflictos_docente:
            mensajes.append(
                f"Quedaron {conflictos_docente} colisión(es) de docente; "
                "se conserva el horario porque cambiar de ambiente no las resuelve."
            )
        total_reasignados = len(set([cid for _, cid in cambios_salones] + list(ids_horario_modificado)))
        return {
            "reasignados": total_reasignados,
            "conflictos_ambiente": conflictos_ambiente,
            "conflictos_docente": conflictos_docente,
            "mensajes": mensajes,
        }
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def validar_cambio_manual(asignacion_id, salon_id, dia, hora_inicio, hora_fin):
    """
    Verifica que al editar manualmente un horario no choque ni el profesor ni el salón.
    """
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT profesor_id FROM asignaciones_horario WHERE id = ?", (asignacion_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        return False, "La asignación no existe."

    profesor_id = row["profesor_id"]

    cursor.execute("""
        SELECT a.hora_inicio, a.hora_fin, m.nombre AS materia
        FROM asignaciones_horario a
        JOIN materias m ON m.id = a.materia_id
        WHERE a.profesor_id = ? AND a.dia = ? AND a.id != ?
    """, (profesor_id, dia, asignacion_id))
    for r in cursor.fetchall():
        if hay_solapamiento(hora_inicio, hora_fin, r["hora_inicio"], r["hora_fin"]):
            conn.close()
            return False, f"El profesor ya dicta '{r['materia']}' el {dia} de {r['hora_inicio']} a {r['hora_fin']}."

    cursor.execute("SELECT codigo FROM salones WHERE id = ?", (salon_id,))
    salon_row = cursor.fetchone()
    if salon_row and not es_ambiente_compartible(salon_row["codigo"]):
        cursor.execute("""
            SELECT a.hora_inicio, a.hora_fin, m.nombre AS materia, p.nombre AS profesor
            FROM asignaciones_horario a
            JOIN materias m ON m.id = a.materia_id
            JOIN profesores p ON p.id = a.profesor_id
            WHERE a.salon_id = ? AND a.dia = ? AND a.id != ?
        """, (salon_id, dia, asignacion_id))
        for r in cursor.fetchall():
            if hay_solapamiento(hora_inicio, hora_fin, r["hora_inicio"], r["hora_fin"]):
                conn.close()
                return False, (
                    f"El salón seleccionado ya está ocupado el {dia} de {r['hora_inicio']} a {r['hora_fin']} "
                    f"por '{r['materia']}' (Prof. {r['profesor']})."
                )

    conn.close()
    return True, "OK"
