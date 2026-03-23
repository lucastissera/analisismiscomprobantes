# analisismiscomprobantes

Analiza Mis Comprobantes ARCA y emite un `.xlsx` ajustado (notas de credito en negativo, multiplicacion por tipo de cambio) desde archivos `.xlsx` o `.csv`.

## Uso local

```bash
python -m pip install -r requirements.txt
python app.py
```

Abrir en navegador: `http://127.0.0.1:5000`

## Despliegue web externo en Render

1. Subir estos cambios a GitHub (`git add . && git commit -m "deploy config" && git push`).
2. Entrar a [Render](https://render.com/) con tu cuenta de GitHub.
3. Click en **New +** -> **Web Service**.
4. Elegir el repo `analisismiscomprobantes`.
5. Completar:
   - **Environment**: `Python 3`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
6. Crear el servicio y esperar el deploy.
7. Render te va a dar una URL publica (`https://...onrender.com`) para compartir.

## Notas

- El frontend tiene el boton **Subir excel** y acepta `.xlsx` y `.csv`.
- Muestra tabla de sumas y permite descargar el archivo ajustado.
