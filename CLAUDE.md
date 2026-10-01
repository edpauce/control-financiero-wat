# Instrucciones del agente — control financiero

Trabajas dentro del **framework WAT** (Workflows, Agents, Tools). La IA razona y decide; el
código determinista ejecuta. Esa separación es lo que hace al sistema confiable.

## Cómo operar

1. **Busca primero un tool existente.** La lógica ejecutable vive en `tools/` y `tools/lib/`.
   Crea scripts nuevos solo si no existe nada para la tarea.
2. **Las rutas pasan por `lib/sources.py`.** Nunca rutas fijas con fecha: `latest(clave)` devuelve
   el archivo más reciente por glob.
3. **Toda escritura sobre un maestro pasa por el gate** (`lib/policy.py`). Respalda antes
   (`cierre.respaldar`), declara capacidad, periodo, monto y evidencia, y respeta la decisión.
4. **Sin `--aplicar` no se escribe.** Muestra la simulación y pide confirmación.
5. **Cuando algo falla, lee el error completo**, corrige el tool y vuelve a probar. Si la
   corrección cambia una regla de negocio, pregunta antes.

## Reglas que no se negocian

- Un mes **cerrado** no se regenera: sus correcciones se hacen editando celdas en sitio.
- Los extractos del banco y los reportes de SUNAT son **insumos**: se leen, jamás se escriben.
- Un cobro o un pago solo cambia de estado **con evidencia citada** (constancia, número de
  operación, archivo fuente).
- Lo ambiguo **se propone y se marca**; no se cierra.
- Las credenciales viven en `.env`. El agente no las lee ni las escribe.

## Verificación

```bash
python -m pytest tests/ -q
```

Un cambio en `tools/` no está terminado hasta que los tests pasan.
