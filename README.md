# analisismiscomprobantes

Analiza Mis Comprobantes ARCA y emite un `.xlsx` ajustado (notas de crédito en negativo, multiplicación por tipo de cambio) desde archivos `.xlsx` o `.csv`.

## Uso local (desarrollo)

```bash
python -m pip install -r requirements.txt
python app.py
```

Abrir en navegador: `http://127.0.0.1:5000` (o el puerto que indique la consola).

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
| Ejecutable | `dist\MisComprobantesAnalisis\MisComprobantesAnalisis.exe` |
| Librerías y recursos | `dist\MisComprobantesAnalisis\_internal\` |

Convéniente: comprimir en **ZIP** la carpeta `MisComprobantesAnalisis` completa y que el usuario la descomprima donde quiera (Escritorio, `C:\Programas\`, etc.).

### Cómo usarla en otra PC

1. Descomprimir la carpeta `MisComprobantesAnalisis`.
2. Ejecutar `MisComprobantesAnalisis.exe`.
3. Se abre una **ventana de la aplicación** (Edge o Chrome en modo app, si están instalados) apuntando a `http://127.0.0.1:8765` (puerto por defecto). Además puede verse una **consola** (ventana negra): dejala abierta mientras usás la app; al cerrarla se detiene el programa.
4. Cerrá la ventana de la interfaz cuando termines (la consola podés cerrarla después para salir del todo).

**Puerto:** si 8765 está ocupado, antes de abrir el `.exe` podés definir la variable de entorno `PORT` (por ejemplo `PORT=8777`).

### Credenciales (login)

El instalable portable **no incluye** tu `auth_users.json` del proyecto (para no subir claves al repo). Incluye solo `auth_users.example.json` con un usuario de ejemplo:

- **Usuario:** `admin`  
- **Contraseña:** `definir_clave_segura_aqui` (texto literal del archivo de ejemplo)

Para usar **tus** usuarios y claves:

1. **Recomendado:** mantené `auth_users.json` en la **raíz del repo**. Al ejecutar `build_windows.bat` o `python tools/portable_build.py` (o al guardar ese archivo con `watch_portable.bat` activo), se copia solo a `dist\MisComprobantesAnalisis\auth_users.json` junto al `.exe`.

2. También podés copiar manualmente ese JSON junto al `.exe` en una máquina ya desplegada, o definir **`AUTH_USERS_PATH`** con la ruta absoluta a tu JSON.

Si no hay ningún JSON válido y tampoco definiste `AUTH_ADMIN_USER` / `AUTH_ADMIN_PASSWORD` en el entorno, el login no aceptará credenciales.

### Plantillas de imputación guardadas (solo en portable / local)

En el ejecutable de Windows, las plantillas se guardan en el perfil del usuario:

`%LOCALAPPDATA%\DepuracionExcelComprobantes\plantillas_imputacion\`

Ahí están el índice `plantillas.json` y los archivos `.xlsx` / `.csv` asociados. No se pierden al actualizar el portable **mientras no borres** esa carpeta de datos del usuario.

### Cómo generar de nuevo el portable

**Build único (recompila y copia claves):**

- **Doble clic o CMD:** `build_windows.bat` — instala dependencias, ejecuta PyInstaller y, si existe `auth_users.json` en la **raíz del repo**, lo copia automáticamente a `dist\MisComprobantesAnalisis\auth_users.json` (junto al `.exe`).
- **O a mano:** `python tools/portable_build.py` (desde la raíz del proyecto).

**Vigilancia automática** (cada vez que guardás `auth_users.json` en la raíz, tras ~2,5 s sin nuevos guardados se recompila el portable y se vuelve a copiar el archivo de claves):

1. `python -m pip install watchdog` (una vez).
2. Ejecutá **`watch_portable.bat`** o `python tools/portable_watch.py`.
3. Editá y guardá `auth_users.json`; no hace falta copiar nada a mano ni lanzar PyInstaller aparte.

Opciones del vigilante: `python tools/portable_watch.py --no-initial` (no compila al arrancar, solo al detectar cambios).

La salida del build queda en `dist\MisComprobantesAnalisis\`. El empaquetado se define en `MisComprobantesDesktop.spec`.

**Nota:** en el `.exe` no se empaqueta Playwright; la sección **CUIT en ARCA** queda deshabilitada en el build de escritorio (comportamiento pensado para evitar dependencias pesadas).

---

## Web y portable: un solo código, dos entregas

El portable **no es un fork**: usa el mismo `app.py`, `templates/`, `sumar_imp_total.py`, `i18n.py`, etc. que el despliegue web.

**Regla práctica:** cada vez que subís cambios pensados para la **web** (HTML, rutas Flask, lógica de procesamiento, textos), si querés que eso exista también en **Windows**, tenés que **volver a ejecutar** el build de PyInstaller en una máquina con el código actualizado y redistribuir la carpeta `dist\MisComprobantesAnalisis\` (o el ZIP).

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
- [ ] (Opcional en desarrollo) Con `watch_portable.bat` activo, cada guardado de `auth_users.json` recompila y sincroniza solo.
- [ ] Smoke test en otra carpeta o otra PC: abrir `MisComprobantesAnalisis.exe`, login, procesar, plantillas si aplica.
- [ ] Si distribuís credenciales propias: asegurate de que el build haya copiado `auth_users.json` a `dist\…` (automático con `build_windows.bat` / `portable_build.py`) o documentá el uso de `AUTH_USERS_PATH`.
- [ ] Generá el **ZIP** de toda la carpeta `dist\MisComprobantesAnalisis\` (no solo el `.exe`).
- [ ] Nombrá el ZIP con versión o fecha (ej. `MisComprobantesAnalisis_2026-05-17.zip`) para saber qué build es.

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
   - **Build Command**: `pip install -r requirements.txt && playwright install chromium && playwright install-deps chromium`
   - **Start Command**: `gunicorn app:app`
6. Crear el servicio y esperar el deploy.
7. Render te va a dar una URL pública (`https://...onrender.com`) para compartir.

## Notas

- El frontend tiene el botón **Subir excel** y acepta `.xlsx` y `.csv`.
- Muestra tabla de sumas y permite descargar el archivo ajustado.
- **CUIT en ARCA** (descarga desde AFIP con Playwright): en el servidor, definí `CUIT_EN_ARCA_PLAYWRIGHT=1` y usá el build command de arriba con Chromium. Sin eso, el formulario valida credenciales y fechas pero no abre el navegador.
