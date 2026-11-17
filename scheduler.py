from database import get_connection, DIAS_ORDEN


def hay_solapamiento(h_ini1, h_fin1, h_ini2, h_fin2):
    """Verifica si dos rangos de hora HH:MM se solapan."""
    return max(h_ini1, h_ini2) < min(h_fin1, h_fin2)


def es_ambiente_compartible(codigo_salon):
    """
    Ambientes virtuales o coordinaciones permiten múltiples cursos simultáneos sin choque físico.
    """
    c = (codigo_salon or "").upper()
    return any(k in c for k in ["VIRTUAL", "COORDINAC", "ONLINE", "CANCHA", "AUDITORIO"])


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
            c.nombre AS carrera_nombre,
            f.nombre AS facultad_nombre,
            ca.seccion,
            ca.bloques_semanales,
            ca.salon_preferido,
            ca.dia_fijo,
            ca.hora_inicio_fija,
            ca.hora_fin_fija
        FROM carga_academica ca
        JOIN profesores p ON p.id = ca.profesor_id
        JOIN materias m ON m.id = ca.materia_id
        JOIN carreras c ON c.id = m.carrera_id
        JOIN facultades f ON f.id = c.facultad_id
        ORDER BY p.nombre ASC, f.nombre ASC, c.nombre ASC, m.nombre ASC
    """)
    cargas = [dict(r) for r in cursor.fetchall()]

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
        bloques_req = carga["bloques_semanales"] or 1
        salon_pref_str = (carga["salon_preferido"] or "").strip().lower()
        salon_pref_id = mapa_salones_codigo.get(salon_pref_str)

        # Caso 1: La fila del archivo ya traía su Día y Hora exactos (ej. Excel de oferta de universidad)
        if carga["dia_fijo"] and carga["hora_inicio_fija"] and carga["hora_fin_fija"]:
            dia_f = carga["dia_fijo"]
            hi_f = carga["hora_inicio_fija"]
            hf_f = carga["hora_fin_fija"]
            prof_choque = any(
                d == dia_f and hay_solapamiento(hi_f, hf_f, oi, of)
                for d, oi, of, _ in ocupado_profesor.get(pid, [])
            )
            salon_elegido = None
            if not prof_choque:
                salones_candidatos = list(salones)
                if salon_pref_id:
                    salones_candidatos.sort(key=lambda s: 0 if s["id"] == salon_pref_id else 1)

                for salon in salones_candidatos:
                    sid = salon["id"]
                    salon_choque = any(
                        d == dia_f and hay_solapamiento(hi_f, hf_f, oi, of)
                        for d, oi, of, _ in ocupado_salon.get(sid, [])
                    )
                    if es_ambiente_compartible(salon["codigo"]) or not salon_choque:
                        salon_elegido = sid
                        break

            if salon_elegido is not None:
                cursor.execute(
                    """
                    INSERT INTO asignaciones_horario
                    (carga_id, profesor_id, materia_id, salon_id, dia, hora_inicio, hora_fin, estado)
                    VALUES (?, ?, ?, ?, ?, ?, ?, 'Confirmado')
                    """,
                    (cid, pid, mid, salon_elegido, dia_f, hi_f, hf_f)
                )
                ocupado_profesor.setdefault(pid, []).append((dia_f, hi_f, hf_f, salon_elegido))
                ocupado_salon.setdefault(salon_elegido, []).append((dia_f, hi_f, hf_f, pid))
                asignados += 1
            else:
                pendientes += 1
                causa = "docente" if prof_choque else "ambiente"
                mensajes.append(
                    f"Conflicto de {causa}: '{carga['materia_nombre']}' "
                    f"({carga['carrera_nombre']}) con el Prof. {carga['profesor_nombre']} "
                    f"no pudo asignarse en {dia_f} de {hi_f} a {hf_f}."
                )
            continue

        # Caso 2: Asignación automática buscando entre la disponibilidad del profesor
        disponibles_prof = disp_por_profesor.get(pid, [])
        if not disponibles_prof:
            pendientes += bloques_req
            mensajes.append(
                f"Sin disponibilidad registrada: '{carga['materia_nombre']}' "
                f"({carga['carrera_nombre']}) con el Prof. {carga['profesor_nombre']} "
                f"no tiene bloques horarios para asignar."
            )
            continue

        bloques_logrados = 0
        dias_usados_en_materia = set()

        salones_candidatos = list(salones)
        if salon_pref_id:
            salones_candidatos.sort(key=lambda s: 0 if s["id"] == salon_pref_id else 1)

        for pasada in (1, 2):
            if bloques_logrados >= bloques_req:
                break

            for dia, h_ini, h_fin in disponibles_prof:
                if bloques_logrados >= bloques_req:
                    break

                if pasada == 1 and dia in dias_usados_en_materia:
                    continue

                prof_choque = any(
                    d == dia and hay_solapamiento(h_ini, h_fin, oi, of)
                    for d, oi, of, _ in ocupado_profesor.get(pid, [])
                )
                if prof_choque:
                    continue

                salon_elegido = None
                for s in salones_candidatos:
                    sid = s["id"]
                    if es_ambiente_compartible(s["codigo"]):
                        salon_elegido = sid
                        break
                    salon_choque = any(
                        d == dia and hay_solapamiento(h_ini, h_fin, oi, of)
                        for d, oi, of, _ in ocupado_salon.get(sid, [])
                    )
                    if not salon_choque:
                        salon_elegido = sid
                        break

                if salon_elegido is not None:
                    cursor.execute(
                        """
                        INSERT INTO asignaciones_horario
                        (carga_id, profesor_id, materia_id, salon_id, dia, hora_inicio, hora_fin, estado)
                        VALUES (?, ?, ?, ?, ?, ?, ?, 'Confirmado')
                        """,
                        (cid, pid, mid, salon_elegido, dia, h_ini, h_fin)
                    )
                    ocupado_profesor.setdefault(pid, []).append((dia, h_ini, h_fin, salon_elegido))
                    ocupado_salon.setdefault(salon_elegido, []).append((dia, h_ini, h_fin, pid))
                    dias_usados_en_materia.add(dia)
                    bloques_logrados += 1
                    asignados += 1

        if bloques_logrados < bloques_req:
            faltantes = bloques_req - bloques_logrados
            pendientes += faltantes
            mensajes.append(
                f"Conflicto de disponibilidad/salón: A '{carga['materia_nombre']}' "
                f"({carga['carrera_nombre']}) con el Prof. {carga['profesor_nombre']} "
                f"le faltaron {faltantes} bloque(s) por asignar."
            )

    conn.commit()
    conn.close()
    return {
        "asignados": asignados,
        "pendientes": pendientes,
        "mensajes": mensajes,
    }


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
