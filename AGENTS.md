# Instrucciones para agentes (Cursor / IA)

## Build portable (solo cuando lo pida el usuario)

El ejecutable vive en:

`dist/AnalisisIntegralContribuyente/AnalisisIntegralContribuyente.exe`

**No lo recompiles automáticamente** al terminar cambios de código. Compilá **solo** si el usuario lo pide explícitamente.

### Comando

Desde la raíz del proyecto:

```powershell
python tools/portable_build.py
```

O en Windows:

```bat
build_windows.bat
```

### Vigilancia automática (opcional, manual del usuario)

Si el usuario quiere rebuilds mientras edita, puede dejar abierto en una terminal (recompila ~3,5 s después del último guardado):

```bat
watch_portable.bat
```

Eso es independiente del agente: no lo arranques vos salvo que te lo pidan.

### Hooks de Cursor

Los hooks de rebuild automático (`afterFileEdit` / `stop`) están **desactivados** en `.cursor/hooks.json`. La regla `.cursor/rules/rebuild-portable.mdc` también indica compilar solo a pedido.

### Qué no versionar

`dist/` y `build/` están en `.gitignore`; el `.exe` se genera localmente, no se sube a Git.
