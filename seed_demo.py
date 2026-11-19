import os
import openpyxl
import docx
from parsers import importar_excel
from scheduler import generar_horarios_automaticos


DATOS_DEMO = [
    [
        "Facultad",
        "Carrera",
        "Profesor",
        "Materia",
        "Disponibilidad",
        "Salón Preferido",
        "Sección",
        "Semestre",
        "Bloques",
    ],
    [
        "Facultad de Ingeniería",
        "Ingeniería de Sistemas",
        "Juan Pérez",
        "Programación I, Estructuras de Datos",
        "Lunes 07:00-11:00; Miércoles 07:00-11:00; Viernes 09:00-11:00",
        "Lab Computación 1",
        "Sección A",
        "2do Semestre",
        "2",
    ],
    [
        "Facultad de Ingeniería",
        "Ingeniería Civil / Ingeniería Industrial",
        "Juan Pérez",
        "Cálculo I",
        "Martes 07:00-11:00; Jueves 07:00-11:00",
        "Aula 101",
        "Sección B",
        "1er Semestre",
        "1",
    ],
    [
        "Facultad de Ingeniería",
        "Ingeniería de Sistemas",
        "María Rodríguez",
        "Bases de Datos, Ingeniería de Software",
        "Martes 09:00-13:00; Jueves 09:00-13:00; Viernes 14:00-18:00",
        "Lab Computación 2",
        "Sección A",
        "5to Semestre",
        "2",
    ],
    [
        "Facultad de Ciencias Económicas y Sociales",
        "Administración de Empresas",
        "Carlos Gómez",
        "Contabilidad General, Microeconomía",
        "Lunes 09:00-13:00; Miércoles 09:00-13:00; Jueves 14:00-16:00",
        "Aula 101",
        "Sección A",
        "2do Semestre",
        "1",
    ],
    [
        "Facultad de Ciencias Económicas y Sociales",
        "Contaduría Pública",
        "Carlos Gómez",
        "Auditoría Financiera",
        "Martes 14:00-18:00; Viernes 07:00-11:00",
        "Aula 102",
        "Sección Única",
        "6to Semestre",
        "1",
    ],
    [
        "Facultad de Ciencias de la Salud",
        "Medicina / Enfermería",
        "Ana Martínez",
        "Anatomía Humana, Fisiología",
        "Lunes 07:00-13:00; Miércoles 14:00-18:00; Viernes 07:00-11:00",
        "Aula Magna 301",
        "Sección A",
        "1er Semestre",
        "1",
    ],
    [
        "Facultad de Humanidades y Educación",
        "Educación Mención Integral",
        "Luis Hernández",
        "Psicología Educativa, Didáctica General",
        "Martes 07:00-11:00; Jueves 07:00-11:00; Sábado 09:00-13:00",
        "Aula 201",
        "Sección A",
        "3er Semestre",
        "1",
    ],
]


def crear_archivos_ejemplo(directorio_base):
    carpeta = os.path.join(directorio_base, "ejemplos")
    os.makedirs(carpeta, exist_ok=True)

    # 1. Crear archivo Excel (.xlsx)
    ruta_xlsx = os.path.join(carpeta, "plantilla_profesores.xlsx")
    if not os.path.exists(ruta_xlsx):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Carga Académica"
        for fila in DATOS_DEMO:
            ws.append(fila)
        wb.save(ruta_xlsx)

    # 2. Crear archivo Word (.docx)
    ruta_docx = os.path.join(carpeta, "plantilla_profesores.docx")
    if not os.path.exists(ruta_docx):
        doc = docx.Document()
        doc.add_heading("Oferta y Disponibilidad de Profesores - Universidad", level=1)
        doc.add_paragraph(
            "Tabla de entrada con los profesores, facultades, carreras, materias y disponibilidad horaria."
        )
        tabla = doc.add_table(rows=1, cols=len(DATOS_DEMO[0]))
        hdr_cells = tabla.rows[0].cells
        for i, encabezado in enumerate(DATOS_DEMO[0]):
            hdr_cells[i].text = encabezado

        for fila in DATOS_DEMO[1:]:
            row_cells = tabla.add_row().cells
            for i, val in enumerate(fila):
                row_cells[i].text = str(val)
        doc.save(ruta_docx)

    # 3. Crear archivo Texto (.txt)
    ruta_txt = os.path.join(carpeta, "plantilla_profesores.txt")
    if not os.path.exists(ruta_txt):
        with open(ruta_txt, "w", encoding="utf-8") as f:
            for fila in DATOS_DEMO:
                f.write(" | ".join(str(x) for x in fila) + "\n")

    return {
        "xlsx": ruta_xlsx,
        "docx": ruta_docx,
        "txt": ruta_txt,
    }


def cargar_demo_en_bd(directorio_base):
    rutas = crear_archivos_ejemplo(directorio_base)
    res_import = importar_excel(rutas["xlsx"])
    res_sched = generar_horarios_automaticos()
    return {
        "importados": res_import["insertados"],
        "asignados": res_sched["asignados"],
        "pendientes": res_sched["pendientes"],
    }
