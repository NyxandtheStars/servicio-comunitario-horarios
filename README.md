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
   - **Carreras**: Panel para agregar y modificar carreras o PNF y su facultad/categoría. Incluye PNF Fisioterapia, Veterinaria y Agroecología, agrupados en «Programas Nacionales de Formación (PNF)»; al renombrar o reclasificar una carrera se conservan sus materias y horarios.

2. **Gestión en Blanco o con Datos Demo**
   - **Eliminar Carga Académica**: Botón visible en el Dashboard y en la barra lateral que elimina las cargas académicas, disponibilidades y asignaciones actuales. Conserva carreras, materias de los planes de estudio, profesores y ambientes.
   - **Modo en Blanco**: Se conserva el catálogo académico y los profesores registrados; quedan listas para una nueva carga académica las asignaciones, disponibilidades y horarios.
   - **Cargar Demo**: Botón para repoblar la base de datos con los datos de prueba y prototipo cuando sea necesario hacer demostraciones o pruebas de verificación.
   - Vistas con estados vacíos amigables cuando no hay registros aún.

3. **Base de Datos SQL Ligera y Persistente (`SQLite`)**
   - En ejecución desde Python, los datos se almacenan en `universidad_horarios.db` junto al proyecto. En el ejecutable Windows se guardan en `%LOCALAPPDATA%\UNESR\Horarios`, para no requerir permisos de escritura en la carpeta de instalación. Si se encuentra una base anterior junto al `.exe`, se copia a esa ubicación la primera vez sin modificar el original.
   - El registro de errores de arranque del ejecutable se guarda en `%LOCALAPPDATA%\UNESR\Horarios\errores_inicio.log`.

---

## Compatibilidad Multiplataforma (Windows y Canaima / Debian / Arch)

El sistema está optimizado para funcionar sin fricción en cualquier entorno:

### En Windows
* **Opción rápida:** Descarga el artefacto `SGHU_Universidad_Windows` desde GitHub Actions, extrae la carpeta completa y ejecuta [`iniciar_sistema.bat`](iniciar_sistema.bat). El paquete incluye el ejecutable y no instala Python ni dependencias.
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

### Servidor multiusuario con Docker y PostgreSQL

El servidor web Flask permite que varios equipos consulten y carguen horarios en una base PostgreSQL compartida. Se instala en una computadora que permanecerá encendida y conectada a la red institucional; cada profesor usa un navegador. No se debe instalar PostgreSQL en cada equipo docente.

La navegación web está organizada como la aplicación UNESR: **Dashboard**, **Nueva Carga**, **Horarios**, **Ambientes**, **Docentes** y **Carreras**. El Dashboard es un resumen; el registro individual y la importación están en su panel independiente de Nueva Carga. El logo institucional se sirve desde `src/UNESR.png`. La interfaz Flet de escritorio no se modifica con estos cambios.

#### 1. Preparar la computadora servidor

Usa una computadora Linux (por ejemplo, Debian, Ubuntu o Canaima) con acceso administrativo y espacio suficiente para la base, los archivos importados y las copias de seguridad. Instala Docker Engine y el complemento Docker Compose siguiendo la documentación oficial correspondiente a esa distribución; comprueba la instalación:

```bash
docker --version
docker compose version
```

Configura una dirección IP fija o una reserva DHCP para el servidor. Así los profesores conservarán la misma dirección para abrir el sitio.

#### 2. Obtener el proyecto y guardar los secretos

En la computadora servidor, clona este repositorio y entra a su carpeta. Para actualizar una instalación ya existente, primero realiza una copia de seguridad (ver abajo) y luego actualiza el mismo checkout; no borres la carpeta ni los volúmenes Docker.

```bash
git clone https://github.com/mooritaa/servicio-comunitario-horarios.git
cd servicio-comunitario-horarios
```

Crea `.env` a partir del ejemplo y genera dos secretos independientes. Pega los valores generados en `POSTGRES_PASSWORD` y `FLASK_SECRET_KEY`, sin comillas ni espacios; conserva `WEB_PORT=5000` salvo que ese puerto ya esté ocupado. No compartas ni subas `.env` a GitHub.

```bash
cp .env.example .env
openssl rand -hex 32
openssl rand -hex 32
chmod 600 .env
```

Si `openssl` no está disponible, genera cada valor con `python3 -c "import secrets; print(secrets.token_hex(32))"`.

#### 3. Crear la base e iniciar la aplicación

Desde la carpeta del proyecto:

```bash
docker compose config --quiet
docker compose up --build -d
docker compose ps
docker compose logs --tail=100 web db
```

La aplicación y PostgreSQL se inician juntos. PostgreSQL escucha únicamente dentro de la red privada de Docker; solo el puerto web se publica en la máquina servidor. El esquema se crea al iniciar la aplicación. Esta instalación empieza con una base PostgreSQL vacía: no copia ni modifica `universidad_horarios.db` ni `universidad_horarios_backup.db`. No cargues datos demo ni importes archivos reales hasta acordar qué datos deben ingresarse.

Para comprobar desde el propio servidor, abre `http://127.0.0.1:5000`. Desde otro equipo de la misma red, abre `http://<IP-DEL-SERVIDOR>:5000` (o el puerto configurado en `WEB_PORT`). La dirección IP se puede consultar con `hostname -I`.

#### 4. Permitir acceso solo a la red autorizada

La aplicación actualmente no tiene inicio de sesión ni permisos por usuario. Antes de dar la dirección a los profesores, limita el puerto web con el firewall al rango de la red institucional. Por ejemplo, con UFW sustituye `192.168.1.0/24` por el rango real de la institución:

```bash
sudo ufw allow from 192.168.1.0/24 to any port 5000 proto tcp
sudo ufw status
```

No habilites el acceso desde cualquier origen, no reenvíes el puerto del router a Internet y no publiques el servicio directamente en Internet. Si se requiere acceso externo, primero debe configurarse una solución segura de acceso remoto y autenticación.

#### 5. Operación, reinicio y actualización

Los contenedores vuelven a iniciar automáticamente si Docker reinicia. Comandos habituales:

```bash
docker compose ps
docker compose logs -f web
docker compose restart web
docker compose stop
docker compose start
```

Para actualizar el código, haz primero una copia de seguridad, actualiza el checkout y reconstruye la aplicación:

```bash
git pull
docker compose up --build -d
docker compose ps
```

`docker compose down` detiene y elimina los contenedores, pero conserva los volúmenes persistentes. **Nunca uses `docker compose down -v` para una detención o actualización normal**: `-v` elimina los volúmenes de PostgreSQL y de archivos cargados.

#### 6. Copias de seguridad y restauración

Realiza respaldos periódicos y guárdalos fuera de la computadora servidor, en un lugar con acceso restringido. Esto crea un respaldo PostgreSQL en formato comprimido:

```bash
mkdir -p respaldos
docker compose exec -T db pg_dump -U horarios -d horarios -Fc \
  > "respaldos/horarios-$(date +%F).dump"
```

Los archivos importados están en otro volumen y deben respaldarse por separado:

```bash
docker compose cp web:/app/uploads ./respaldos/uploads
```

Verifica que los archivos generados existan y tengan tamaño mayor que cero, y copia la carpeta `respaldos` a un almacenamiento seguro. Un respaldo que no se ha probado no garantiza que pueda restaurarse.

La restauración reemplaza datos existentes. Hazla solo cuando sea necesario, conserva primero una copia del estado actual, y detén la aplicación para evitar escrituras concurrentes. Con un archivo `.dump` creado por el comando anterior:

```bash
docker compose stop web
docker compose exec -T db pg_restore -U horarios -d horarios \
  --clean --if-exists < respaldos/horarios-AAAA-MM-DD.dump
docker compose start web
```

Sustituye el nombre por el respaldo que vas a restaurar. No ejecutes una restauración como parte de una actualización normal.

### Compilar y distribuir el ejecutable para los profesores

El workflow **Compilar Ejecutables Multiplataforma** construye el paquete de escritorio en runners Windows y Linux de GitHub Actions. Un `.exe` de Windows debe compilarse en Windows; Docker en Linux no lo genera. Para iniciar una compilación:

1. Abre el repositorio en GitHub y entra a **Actions**.
2. Selecciona **Compilar Ejecutables Multiplataforma** en la lista de workflows.
3. Pulsa **Run workflow**, elige la rama que contiene los cambios aprobados y confirma.
4. Espera que finalicen los trabajos de compilación y validación. Si uno falla, no distribuyas ese paquete.
5. Abre la ejecución completada y descarga el artefacto **SGHU_Universidad_Windows**.
6. Copia el ZIP a cada equipo Windows, extráelo completo en una carpeta con permisos de escritura del profesor y ejecuta `iniciar_sistema.bat`. No ejecutes el `.exe` desde dentro del ZIP ni copies únicamente el `.exe`: necesita los demás archivos del paquete.

El artefacto de Actions es temporal y puede caducar; para una versión estable se puede crear y publicar un tag `v*`, lo que genera un GitHub Release con los paquetes de Windows y Linux. Descarga el paquete aprobado y conserva la versión distribuida para facilitar soporte.

La distribución Windows se genera como **un único `.exe` con las dependencias incluidas**; no requiere instalar Python ni ejecutar `pip`. Al iniciarlo por primera vez crea su base SQLite y carpetas de datos en `%LOCALAPPDATA%\UNESR\Horarios`. El archivo `.bat` sigue disponible como lanzador compatible, pero no es necesario para abrir el sistema.

El ejecutable de escritorio funciona localmente con SQLite y **no se conecta** al servidor PostgreSQL. Si varios profesores deben compartir los mismos datos, cada uno debe entrar por navegador a la URL del servidor; no se deben intercambiar copias de la base SQLite como mecanismo multiusuario. Para soporte de un fallo al abrir en Windows, ejecuta `iniciar_sistema.bat` y comparte el contenido de `errores_inicio.log` (sin incluir datos personales que no sean necesarios).

### Modo web local sin Docker

En Linux, `./iniciar_sistema.sh --server` inicia el servidor web con SQLite si `DATABASE_URL` no está configurado, o PostgreSQL si se configura una URL válida. Por defecto escucha solo en `127.0.0.1`, por lo que sirve para pruebas en el mismo equipo. Para un servidor multiusuario se recomienda el despliegue Docker documentado arriba.
