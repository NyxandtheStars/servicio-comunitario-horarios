# Universidad Simón Rodríguez (UNESR) — Sistema de Oferta Académica y Gestión de Horarios (Núcleo Palo Verde)

Sistema interactivo de escritorio desarrollado en **Python (Flet + SQLite)** para la **Universidad Nacional Experimental Simón Rodríguez (UNESR / USR) — Núcleo Palo Verde**, diseñado para organizar la oferta académica universitaria por carrera/facultad, asignar horarios y ambientes sin conflictos de forma automática, e importar cargas académicas desde archivos **Excel (`.xlsx`)**, **Word (`.docx`)** y **Texto (`.txt` / `.csv`)**.

---

## Características Principales

1. **Interfaz Moderna en Modo Claro / Oscuro Institucional (Flet)**
   - **Nuevo Modo Claro**: Diseño limpio, espacioso y descansado a la vista, conservando los acentos cromáticos institucionales (Azul UNESR, Verde, Púrpura, Naranja). Incluye botón para alternar entre Modo Claro y Oscuro según la preferencia del usuario.
   - **Dashboard**: Estadísticas en tiempo real de cursos activos, ambientes ocupados/disponibles, docentes activos y verificación de `0 Conflictos`.
   - **Nueva Carga**: Formulario manual para registrar los datos del docente, carrera, ambiente opcional, materias y disponibilidad semanal. Los horarios se generan usando las franjas seleccionadas; también se admite importación masiva desde `.xlsx`, `.docx`, `.txt` y `.csv`.
   - **Horarios (Oferta Académica General)**: Consulta filtrable por carrera, día y docente (ordenados alfabéticamente), con botón interactivo de reasignación manual de ambiente por fila y exportación a HTML/PDF imprimible.
   - **Ambientes**: Monitoreo visual de ocupación y capacidad de cada aula y laboratorio (`P1-01` a `P4-37` con numeración global consecutiva, `SATELITAL-01/02` y `VIRTUAL-01/02`). Permite registrar nuevos ambientes, modificar su capacidad, marcarlos como **ambiente de camillas** (sin capacidad numérica) y eliminarlos individualmente.
   - **Docentes (Horario Tipo Calendario + Impresión)**:
     - Visualizador interactivo del horario semanal tipo calendario (`Lunes` a `Sábado` por bloque horario) para cada docente.
     - **Modificación y eliminación individual**: Cada docente puede editar sus datos personales (nombre, cédula, correo y condición) o eliminarse del sistema desde su tarjeta o desde la barra del docente seleccionado.
       - Desde el mismo diálogo se puede editar su lista de materias, una por línea.
     - **Reasignación manual de ambiente**: Botón directo en cada tarjeta de materia asignada para cambiar o fijar el salón de clases sin afectar el resto de las asignaciones.
     - Generador de reportes imprimibles HTML/PDF con opción de elegir **todos los profesores registrados** o **un (1) solo profesor** de la lista.
   - **Carreras**: Panel de carreras registradas y su pensum de materias.

2. **Gestión en Blanco o con Datos Demo**
   - **Eliminar Datos Cargados**: Botón visible en el Dashboard y en la barra lateral que, tras una ventana de confirmación, elimina cursos, docentes, carreras y asignaciones cargadas, dejando intactos los ambientes y sus capacidades.
   - **Modo en Blanco**: El sistema queda completamente limpio y listo para registrar los datos reales del nuevo período académico.
   - **Cargar Demo**: Botón para repoblar la base de datos con los datos de prueba y prototipo cuando sea necesario hacer demostraciones o pruebas de verificación.
   - Vistas con estados vacíos amigables cuando no hay registros aún.

3. **Base de Datos SQL Ligera y Persistente (`SQLite`)**
   - Almacenamiento automático en `universidad_horarios.db` ([`database.py`](database.py)) con índices relacionales optimizados.

---

## Compatibilidad Multiplataforma (Windows y Canaima / Debian / Arch)

El sistema está optimizado para funcionar sin fricción en cualquier entorno:

### En Windows
* **Opción rápida:** Hacer doble clic en [`iniciar_sistema.bat`](iniciar_sistema.bat). El script crea el entorno virtual `.venv` e instala las dependencias automáticamente en la primera ejecución.
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
