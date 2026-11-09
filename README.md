# Universidad Simón Rodríguez (UNESR) — Sistema de Oferta Académica y Gestión de Horarios (Núcleo Palo Verde)

Sistema interactivo de escritorio desarrollado en **Python (Flet + SQLite)** para la **Universidad Nacional Experimental Simón Rodríguez (UNESR / USR) — Núcleo Palo Verde**, diseñado para organizar la oferta académica universitaria por carrera/facultad, asignar horarios y ambientes sin conflictos de forma automática, e importar cargas académicas desde archivos **Excel (`.xlsx`)**, **Word (`.docx`)** y **Texto (`.txt` / `.csv`)**.

---

## Características Principales

1. **Interfaz Moderna en Modo Claro / Oscuro Institucional (Flet)**
   - **Nuevo Modo Claro**: Diseño limpio, espacioso y descansado a la vista, conservando los acentos cromáticos institucionales (Azul UNESR, Verde, Púrpura, Naranja). Incluye botón para alternar entre Modo Claro y Oscuro según la preferencia del usuario.
   - **Dashboard**: Estadísticas en tiempo real de cursos activos, ambientes ocupados/disponibles, docentes activos y verificación de `0 Conflictos`.
   - **Nueva Carga**: Registro individual con asignación automática o selección manual de ambiente, día y franja horaria. Importación masiva desde documentos `.xlsx`, `.docx`, `.txt`, `.csv`, `.pdf` e imágenes.
   - **Horarios (Oferta Académica General)**: Consulta filtrable por carrera, día y docente (ordenados alfabéticamente), con botón interactivo de reasignación manual de ambiente por fila y exportación a HTML/PDF imprimible.
   - **Ambientes**: Monitoreo visual de ocupación y capacidad de cada aula y laboratorio (`P1-01` a `P4-07`, `PB`, `VIRTUAL`).
   - **Docentes (Horario Tipo Calendario + Impresión)**:
     - Visualizador interactivo del horario semanal tipo calendario (`Lunes` a `Sábado` por bloque horario) para cada docente.
     - **Reasignación manual de ambiente**: Botón directo en cada tarjeta de materia asignada para cambiar o fijar el salón de clases sin afectar el resto de las asignaciones.
     - Generador de reportes imprimibles HTML/PDF con opción de elegir **todos los profesores registrados** o **un (1) solo profesor** de la lista.
   - **Carreras**: Panel de carreras registradas y su pensum de materias.

2. **Lectura y Escaneo de Hojas de Disponibilidad Docente (PDF e Imágenes / OCR)**
   - Permite cargar fotos de hojas físicas o documentos digitales en formatos **`.pdf`**, **`.jpg`**, **`.jpeg`**, **`.png`**, **`.webp`** y **`.bmp`**.
   - Integra motor OCR (**Tesseract OCR** + **pdftotext/pdftoppm**) que lee el texto de la hoja física, identificando automáticamente:
     - Nombre completo del docente y Cédula de Identidad (`V-XXXXXXX`).
     - Carrera y disponibilidad semanal (días y franjas horarias solicitadas).
     - Asignaturas propuestas a impartir.
   - Diálogo modal interactivo de verificación antes de registrar para revisar y corregir cualquier dato antes de guardarlo en la base de datos.

3. **Gestión en Blanco o con Datos Demo**
   - **Modo en Blanco**: Botón en la barra lateral para vaciar el sistema y dejar la base de datos completamente limpia y lista para registrar los datos reales del nuevo período académico.
   - **Cargar Demo**: Botón para repoblar la base de datos con los datos de prueba y prototipo cuando sea necesario hacer demostraciones o pruebas de verificación.
   - Vistas con estados vacíos amigables cuando no hay registros aún.

4. **Base de Datos SQL Ligera y Persistente (`SQLite`)**
   - Almacenamiento automático en `universidad_horarios.db` ([`database.py`](database.py)) con índices relacionales optimizados.

---

## Compatibilidad Multiplataforma (Windows y Canaima / Debian / Arch)

El sistema está optimizado para funcionar sin fricción en cualquier entorno:

### En Windows
* **Opción rápida:** Hacer doble clic en [`iniciar_sistema.bat`](iniciar_sistema.bat).
* **Opción manual (PowerShell / CMD):**
  ```powershell
  python -m venv .venv
  .\.venv\Scripts\activate
  pip install -r requirements.txt
  python SGHU_Universidad.py
  ```

### En Canaima GNU/Linux (Debian / Arch / Ubuntu)
* **Opción rápida:** Hacer doble clic en [`iniciar_sistema.sh`](iniciar_sistema.sh) o en [`SGHU_Universidad.desktop`](SGHU_Universidad.desktop).
* **Opción manual (Terminal):**
  ```bash
  chmod +x iniciar_sistema.sh
  ./iniciar_sistema.sh
  ```
  *(El script crea el entorno virtual `.venv` e instala las dependencias automáticamente si es necesario).*
* **Nota para Arch Linux:** Si deseas la ventana nativa de escritorio y el sistema avisa de `libmpv.so.1`, instala MPV y crea el enlace simbólico:
  ```bash
  sudo pacman -S mpv
  sudo ln -s /usr/lib/libmpv.so.2 /usr/lib/libmpv.so.1
  ```
  O ejecuta en modo navegador web con:
  ```bash
  ./iniciar_sistema.sh --web
  ```

### Modo Servidor Web en Red Local (UNESR LAN / Wi-Fi)
Si en el departamento o laboratorio hay una mezcla de computadoras (algunas con Canaima y otras con Windows):
1. En **una sola computadora** de la sala, iniciar el servidor web:
   ```bash
   python app.py
   ```
2. Cualquier otra máquina (Canaima, Windows o teléfonos de los profesores) conectada a la misma red o Wi-Fi puede acceder desde su navegador abriendo:
   `http://<IP-DE-LA-MAQUINA>:5000`
3. Todos consultan y gestionan la misma base de datos `universidad_horarios.db` en tiempo real sin tener que instalar nada en sus equipos.
