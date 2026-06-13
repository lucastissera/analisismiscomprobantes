# Análisis Integral del Contribuyente

Plataforma para análisis y automatización ARCA (Mis Comprobantes, DFE, Nuestra Parte): procesa `.xlsx`/`.csv`, emite informes ajustados y descarga desde AFIP.

## Uso local (desarrollo)

```bash
python -m pip install -r requirements.txt
playwright install chromium
python app.py
```

Abrir en navegador: `http://127.0.0.1:5000` (o el puerto que indique la consola).

### Enlace ARCA (certificado digital)

En el panel izquierdo podés:

1. Subir certificado **.pfx/.p12** o par **.crt + .key** (emitidos en Administración de Certificados Digitales de ARCA).
2. Indicar **CUIT representado** y el **período** (máximo un año).
3. Elegir **emitidos** o **recibidos** y, por defecto, **procesar automáticamente** el Excel descargado.

Requisitos: `CUIT_EN_ARCA_UI=1` y `CUIT_EN_ARCA_PLAYWRIGHT=1` (en local suelen activarse solos). La automatización usa Chromium vía Playwright; los selectores de AFIP pueden requerir ajustes si el sitio cambia.

Alternativa sin certificado: planilla Excel con CUIT, clave fiscal y CUIT representado (misma sección).

**Certificado paso a paso:** ver [docs/CERTIFICADO_ARCA.md](docs/CERTIFICADO_ARCA.md).

### Actualizar el `.exe` en `dist/` (obligatorio tras cada cambio)

```bat
build_windows.bat
```

O dejá corriendo `watch_portable.bat` mientras editás. Los agentes de Cursor deben seguir [AGENTS.md](AGENTS.md).

### Plantillas de imputación en disco (solo modo local / escritorio)

Para probar la misma funcionalidad que el `.exe` sin compilar:

```bash
set ENABLE_LOCAL_PLANTILLAS_IMPUTACION=1
python run_desktop.py
```

(O en PowerShell: `$env:ENABLE_LOCAL_PLANTILLAS_IMPUTACION=1` antes de `python run_desktop.py`.)

Los datos de plantillas se guardan por defecto en la carpeta `data_local_imputaciones/` del proyecto (o en la ruta de `IMPUTACIONES_DATA_DIR` si la definís).

---

## Aplicación Windows (portable)

Es la **misma aplicación** que en la web: misma pantalla y lógica, pero el servidor Flask corre **solo en tu PC** (`127.0.0.1`). **No hace falta Internet** para iniciar sesión ni para procesar los Excel/CSV (solo se usa red si abrís enlaces externos, p. ej. WhatsApp desde el login).

En el **`.exe` empaquetado**, por defecto se intenta abrir la interfaz en una **ventana tipo aplicación** usando **Microsoft Edge** o **Google Chrome** con el modo ``--app=URL`` (sin barra de pestañas como en una navegación normal). Sigue siendo el mismo HTML servido en local; **no hace falta Internet** para procesar archivos. Si no se encuentra Edge ni Chrome en las rutas habituales de Windows, se usa el navegador predeterminado del sistema.

- Para abrir siempre el navegador “completo” (pestaña normal): definí **`DESKTOP_APP_WINDOW=0`** antes de ejecutar el `.exe`.
- En **desarrollo** (`python run_desktop.py` sin congelar), por defecto se usa el navegador del sistema en modo normal; podés forzar ventana app con **`DESKTOP_APP_WINDOW=1`**.

No hace falta instalar Python en la PC destino.

### Qué distribuir

Después de compilar (ver más abajo), hay que entregar **toda la carpeta**, no solo el ejecutable:

| Contenido | Ubicación |
|-----------|-----------|
| Ejecutable | `dist\AnalisisIntegralContribuyente\AnalisisIntegralContribuyente.exe` |
| Librerías y recursos | `dist\AnalisisIntegralContribuyente\_internal\` |

Convéniente: comprimir en **ZIP** la carpeta `AnalisisIntegralContribuyente` completa y que el usuario la descomprima donde quiera (Escritorio, `C:\Programas\`, etc.).

### Cómo usarla en otra PC

1. Descomprimir la carpeta `AnalisisIntegralContribuyente`.
2. Ejecutar `AnalisisIntegralContribuyente.exe` (no se abre la ventana negra de consola: solo Edge/Chrome en modo app o el navegador predeterminado).
3. Se abre la **ventana de la aplicación** (Edge o Chrome en modo app, si están instalados) en `http://127.0.0.1:8765` (puerto por defecto).
4. Para **salir por completo**, usá el enlace del pie de página **«Cerrar aplicación de escritorio»** (cierra el proceso del servidor local).

**Puerto:** si 8765 está ocupado, antes de abrir el `.exe` podés definir la variable de entorno `PORT` (por ejemplo `PORT=8777`).

### Credenciales (login)

Hay tres modos compatibles (portable y servidor web):

1. **Nube (recomendado con varios usuarios):** un único `auth_users.json` en un hosting HTTPS. Cada app lo descarga, lo guarda en caché fuera del sistema y lo actualiza cada ~2 minutos. Configuración: ver **[docs/AUTH_USUARIOS_NUBE.md](docs/AUTH_USUARIOS_NUBE.md)**.
   - Servidor: variables `AUTH_USERS_URL` y opcional `AUTH_USERS_REMOTE_TOKEN` en `.env`.
   - Portable: copiá `auth_remote.example.txt` como `auth_remote.txt` junto al `.exe`, o un `.env` con las mismas variables.

2. **Archivo local externo:** `AUTH_USERS_PATH` apunta a un JSON fuera del proyecto.

3. **Archivo junto al sistema (desarrollo / un solo equipo):** `auth_users.json` en la raíz del repo. El build portable puede copiarlo al `dist\` (no subir claves al repo).

Si no hay URL remota, JSON válido ni `AUTH_ADMIN_USER` / `AUTH_ADMIN_PASSWORD`, el login no aceptará credenciales.

### Plantillas de imputación guardadas (solo en portable / local)

En el ejecutable de Windows, las plantillas se guardan en el perfil del usuario:

`%LOCALAPPDATA%\DepuracionExcelComprobantes\plantillas_imputacion\`

Ahí están el índice `plantillas.json` y los archivos `.xlsx` / `.csv` asociados. No se pierden al actualizar el portable **mientras no borres** esa carpeta de datos del usuario.

### Cómo generar de nuevo el portable

**Build único (recompila y copia claves):**

- **Doble clic o CMD:** `build_windows.bat` — instala dependencias, ejecuta PyInstaller y, si existe `auth_users.json` en la **raíz del repo**, lo copia automáticamente a `dist\AnalisisIntegralContribuyente\auth_users.json` (junto al `.exe`).
- **O a mano:** `python tools/portable_build.py` (desde la raíz del proyecto).

**Vigilancia automática** (mientras `watch_portable.bat` o `python tools/portable_watch.py` esté en marcha, tras unos **3,5 s** sin nuevos guardados se **recompila** el portable y se copian las claves si hay `auth_users.json` en la raíz):

1. `python -m pip install watchdog` (una vez).
2. Ejecutá **`watch_portable.bat`** o `python tools/portable_watch.py`.
3. Guardá cambios en **código, plantillas (`templates/`), `i18n.py`, `app.py`, etc.** — el vigilante ignora `dist/`, `build/` y cachés para no reentrar en bucle. **Cerrá el `.exe` del portable** si Windows bloquea archivos en uso durante el build.

Opciones: `python tools/portable_watch.py --no-initial` (no compila al arrancar), `--solo-claves` (solo vigila `auth_users.json`, sin el resto del repo).

La salida del build queda en `dist\AnalisisIntegralContribuyente\`. El empaquetado se define en `MisComprobantesDesktop.spec`.

**Nota:** el portable incluye **Playwright + Chromium** en la subcarpeta `ms-playwright\` (se instala al compilar con `build_windows.bat`). Distribuí **toda** la carpeta `AnalisisIntegralContribuyente`, no solo el `.exe`. El ejecutable no muestra ventana de consola negra.

---

## Web y portable: un solo código, dos entregas

El portable **no es un fork**: usa el mismo `app.py`, `templates/`, `sumar_imp_total.py`, `i18n.py`, etc. que el despliegue web.

**Regla práctica:** cada cambio en la web que deba verse en Windows requiere **recompilar** el portable (`build_windows.bat`, `python tools/portable_build.py`, o dejar **`watch_portable.bat`** abierto mientras editás para que el `dist\` se actualice solo).

**Si agregás recursos nuevos** (por ejemplo otra carpeta de estáticos, nuevos templates, JSON embebidos), revisá `MisComprobantesDesktop.spec` y la lista `datas=` para incluirlos; si no, el portable puede fallar o quedar sin esos archivos.

**Resumen:** mantené el repo como fuente única; la “replicación” al portable es **recompilar** tras los cambios, más ajustar el `.spec` cuando haga falta.

### Checklist de release (web + portable)

Marcá cada ítem al publicar una versión nueva. El orden sugiere: código listo → web → portable → difusión.

#### Código y pruebas locales

- [ ] Cambios probados con `python app.py` (o `run_desktop.py` con variables de entorno que correspondan).
- [ ] Si tocaste imputaciones / plantillas: probá con `ENABLE_LOCAL_PLANTILLAS_IMPUTACION=1` y `run_desktop.py`.
- [ ] Si agregaste dependencias: actualizá `requirements.txt` y verificá que Render siga pudiendo instalar (y que el portable no necesite hooks extra; si falla el build, revisá warnings de PyInstaller).

#### Web (Render / GitHub)

- [ ] `git status` limpio o con los commits que querés publicar.
- [ ] `git push` al branch que Render despliega (p. ej. `main`).
- [ ] En Render: deploy correcto (build + start sin error).
- [ ] Smoke test en la URL pública: login, procesar un `.xlsx` / `.csv`, descarga del informe, idiomas si los cambiaste.
- [ ] Si usás **CUIT en ARCA** en producción: variables y Playwright/Chromium como en el build command de Render.

#### Portable Windows

- [ ] ¿Nuevos archivos en `templates/`, `static/`, JSON u otros datos? → actualizá `datas` en `MisComprobantesDesktop.spec` si hace falta.
- [ ] Ejecutá `build_windows.bat` o `python tools/portable_build.py` (incluye PyInstaller + copia de `auth_users.json` si existe en la raíz).
- [ ] (Opcional) Con `watch_portable.bat` activo, los guardados en el repo actualizan `dist\…` solos; si no, ejecutá `build_windows.bat` o `python tools/portable_build.py` antes de empaquetar el ZIP.
- [ ] Smoke test en otra carpeta o otra PC: abrir `AnalisisIntegralContribuyente.exe`, login, procesar, plantillas si aplica.
- [ ] Si distribuís credenciales propias: asegurate de que el build haya copiado `auth_users.json` a `dist\…` (automático con `build_windows.bat` / `portable_build.py`) o documentá el uso de `AUTH_USERS_PATH`.
- [ ] Generá el **ZIP** de toda la carpeta `dist\AnalisisIntegralContribuyente\` (no solo el `.exe`).
- [ ] Nombrá el ZIP con versión o fecha (ej. `AnalisisIntegralContribuyente_2026-05-17.zip`) para saber qué build es.

#### Cierre

- [ ] Avisaste a quien corresponda: URL web actualizada y/o enlace/archivo del portable nuevo.
- [ ] (Opcional) Tag o release en GitHub alineado con la versión que quedó en web + portable.

---

## Despliegue web externo en Render

1. Subir estos cambios a GitHub (`git add . && git commit -m "deploy config" && git push`).
2. Entrar a [Render](https://render.com/) con tu cuenta de GitHub.
3. Click en **New +** -> **Web Service**.
4. Elegir el repo `analisismiscomprobantes`.
5. Completar:
   - **Environment**: `Python 3`
   - **Build Command**: `pip install -r requirements.txt && playwright install chromium`
   - **Start Command**: `gunicorn app:app`
6. Crear el servicio y esperar el deploy.
7. Render te va a dar una URL pública (`https://...onrender.com`) para compartir.

## Notas

- El frontend tiene el botón **Subir excel** y acepta `.xlsx` y `.csv`.
- Muestra tabla de sumas y permite descargar el archivo ajustado.
- **CUIT en ARCA** (descarga desde AFIP con Playwright): en el servidor, definí `CUIT_EN_ARCA_PLAYWRIGHT=1` y usá el build command de arriba con Chromium. Sin eso, el formulario valida credenciales y fechas pero no abre el navegador.
