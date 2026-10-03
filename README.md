# SISTEMA MAESTRO DE CONTROL DE ASISTENCIA — AUDICONTA PERÚ E.I.R.L.

Sistema profesional de control de asistencia en Excel + VBA (`.xlsm`) construido a
partir de los archivos reales de la empresa (exportación del huellero, cuadro general
de asistencia, reporte de asistencia e informe mensual).

Transforma las marcaciones biométricas en **asistencia diaria → incidencias →
cuadro general → consolidado mensual → informe → histórico**, conservando toda la
trazabilidad y sin destruir información.

---

## 1. Qué contiene este repositorio

| Carpeta / archivo | Descripción |
|---|---|
| `dist/SISTEMA_ASISTENCIA_AUDICONTA.xlsx` | **Libro base** ya formateado: todas las hojas, catálogos, nómina real, parámetros, validaciones y rangos con nombre. Es el punto de partida. |
| `vba/*.bas`, `vba/ThisWorkbook.cls` | **Todo el código VBA** (18 módulos + eventos del libro) listo para importar. |
| `build/build_workbook.py` | Generador del libro base (openpyxl). Permite re-construir el `.xlsx` si hiciera falta. |
| `INSTALACION.md` | Pasos para dejar el sistema funcionando (10 minutos, sin saber VBA). |
| `MANUAL_DE_USO.md` | Manual del usuario final: flujo diario y de cierre de mes. |
| `docs/DECISIONES_TECNICAS.md` | Arquitectura, mapeo con los archivos originales, errores del archivo anterior corregidos y **limitaciones técnicas con su alternativa**. |

> **Importante — límite técnico honesto:** el código VBA de un `.xlsm` se guarda en un
> binario (`vbaProject.bin`) que **solo Excel puede generar**. Por eso el proyecto se
> entrega como libro `.xlsx` + código VBA importable: en Excel se guarda como `.xlsm`,
> se importan los módulos y se ejecuta `INSTALAR` **una sola vez**. Todo queda embebido
> y a partir de ahí el usuario trabaja solo con botones. El detalle está en
> `docs/DECISIONES_TECNICAS.md`.

---

## 2. Arquitectura (principio RAW → PROCESO → CONTROL → REPORTE → CONSOLIDADO → HISTÓRICO)

Hojas del libro:

| Hoja | Rol | Visibilidad |
|---|---|---|
| `01_MENU` | Menú principal con botones + tablero | Visible |
| `CONFIG` | **Configuración maestra** (empresa, horarios, tolerancia, periodo, remuneración) | Visible (protegida) |
| `PERSONAL` | Nómina maestra (independiente del huellero) | Visible |
| `HORARIOS` | Horarios personalizados por trabajador | Visible |
| `CALENDARIO` | Feriados / descansos / eventos | Visible |
| `CATALOGO_ESTADOS` | Códigos, colores, prioridad, nota de remuneración | Muy oculta |
| `DATA_HUELLERO` | **RAW** de marcaciones (append-only, con clave única) | Muy oculta |
| `PROCESO` | Capa de cálculo (1 fila por trabajador × día) | Muy oculta |
| `REPORTE DIARIO` | Reporte oficial del día | Visible |
| `CONSOLIDADO` | Consolidado mensual (base de remuneraciones) | Visible |
| `CUADRO_GENERAL` | Semilla; genera 1 hoja por mes con el formato oficial | Muy oculta |
| `<MES> AAAA` | Cuadro general del mes generado (M/T por día, resumen, leyenda) | Visible |
| `INCIDENCIAS` | Registro y gestión de incidencias | Visible |
| `AUDITORIA` | Toda modificación manual queda registrada | Muy oculta |
| `ALERTAS` | Inconsistencias detectadas | Visible bajo demanda |
| `HISTORICO_MENSUAL` | Meses cerrados (congelados) | Muy oculta |
| `LOG_ERRORES` | Errores controlados | Muy oculta |
| `INFORME` / `PRUEBAS` | Generadas por macro | Visible |

Módulos VBA: `modInstalador`, `modPrincipal`, `modConfiguracion`, `modImportacion`,
`modPersonal`, `modMarcaciones`, `modAsistencia`, `modIncidencias`, `modConsolidado`,
`modCuadroGeneral`, `modReportes`, `modAuditoria`, `modBackup`, `modValidaciones`,
`modErrores`, `modUtilidades`, `modPruebas`, y eventos en `ThisWorkbook`.

---

## 3. Puesta en marcha rápida

1. Abra `dist/SISTEMA_ASISTENCIA_AUDICONTA.xlsx` en Excel y **Guardar como → Libro habilitado para macros (`.xlsm`)**.
2. `Alt + F11` → menú **Archivo → Importar archivo…** → importe todos los `vba/*.bas` y `vba/ThisWorkbook.cls`.
3. `Alt + F8` → ejecute **`INSTALAR`**.
4. Ejecute **`PRUEBAS_UNITARIAS`** para confirmar la lógica (los 16 casos del requisito 58).
5. Listo: trabaje desde el menú.

El detalle con capturas conceptuales está en `INSTALACION.md`.

---

## 4. Reglas críticas garantizadas

- **Tolerancia exacta al segundo** (sin `ROUND`): 07:35:00 PUNTUAL / 07:35:01 TARDANZA; 08:05:00 PUNTUAL / 08:05:01 TARDANZA; 15:05:00 PUNTUAL / 15:05:01 TARDANZA.
- **La nómina NUNCA se reconstruye al importar.** Un trabajador agregado a mano permanece y aparece en reportes/consolidado.
- **Importación append-only con clave única** (ID + fecha/hora + estado + dispositivo): reimportar el mismo archivo no duplica.
- **No se asume FALTA** ante ausencia de marca: queda `PENDIENTE DE REVISIÓN` hasta la decisión del responsable.
- **Incidencia justificada no se convierte en falta.**
- **Meses cerrados quedan congelados** y no se reprocesan.
- **Respaldo automático** antes de importar, cerrar mes y guardar.
- **Auditoría** de toda modificación manual.
- Parámetros (empresa, horarios, tolerancia, códigos, feriados, remuneración) **en hojas de configuración, no en el código**.
