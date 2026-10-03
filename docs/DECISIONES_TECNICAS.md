# DECISIONES TÉCNICAS Y ANÁLISIS DE LOS ARCHIVOS ORIGINALES

Este documento cumple el requisito 65 (metodología) y 64 (indicar limitaciones y
proponer la alternativa más segura).

---

## 1. Análisis de los archivos adjuntos

### 1.1 Exportación del huellero — `01.10.2026 t.xls`
- Columnas: **Número, Nombre, Tiempo, Estado, Dispositivos, Tipo de Registro**.
- `Tiempo` es texto tipo `1/10/2026 08:04:39`.
- **Muchas marcaciones tienen el campo `Estado` VACÍO** (p. ej. la marca de las 15:02).
  → El sistema **infiere** entrada/salida por contexto horario, sin descartar esas marcas.
- Aparecen **nombres distintos al cuadro** (`AARON  PEÑA JUSCAMAYTA` con doble espacio,
  vs. `PEÑA JUSCAMAYTA AARON R.`), y **múltiples marcas muy cercanas**
  (`14` a las 15:03:18 y 15:03:20). → Normalización + clave de tokens + detección de duplicidad.
- El `Número` del huellero se usa como **ID HUELLERO** y se mapea a la nómina.

### 1.2 Cuadro general — `CUADRO GENERAL DE ASISTENCIA.xlsx`
- Una hoja por mes; layout **variable** (según días del mes y día de inicio).
- Estructura: `N | DNI | APELLIDOS Y NOMBRES`, y por cada día **dos columnas `M` y `T`**,
  con cabeceras de día de semana (LUN…DOM) y los domingos combinados como **DESCANSO**.
- **Códigos reales**: `2` (asistió/puntual), `T`, `F`, `PE`, `S`, `FE`, `V`, `C`, `NF`, `DESCANSO`.
- Dos bloques de empresa: **AUDICONTA PERÚ** y **SEG PRO FORCE S.A.C.** (seguridad).
- **CUADRO RESUMEN CONTROL DE ASISTENCIA**: `PUNTUAL | TARDAN. | FALTA | PERMISO | SUSP. | NO FIRMO`,
  con leyenda de códigos y notas de remuneración (TARDE 30% R.D., FALTA 1 R.D., SUSP. 2 R.D.,
  SIN FIRMA 5% R.D., COMISIÓN/VACACIONES/FERIADOS: sí se paga).
- La **nómina autoritativa con DNI** se tomó de aquí (19 trabajadores).

### 1.3 Reporte de asistencia — `REPORTE_ASISTENCIA_AUDICONTA 01.10.2026.xlsx`
Hojas `DATA`, `PERSONAL`, `PARAMETROS`, `REPORTE DIARIO`. Se conservó su apariencia
(cabecera institucional, tabla N / Apellidos / Hora entrada / Hora salida / Minutos
tardanza / Condición / Firma / Observaciones, y el resumen al pie).

### 1.4 Informe mensual — `INFORME MENSUAL N 001-2026 ... .pdf`
Es un **memorándum/cronograma corporativo** con la papelería institucional
(encabezado AUDICONTA, RUC, dirección, correo y **bloque de firmas con DNI**).
De ahí se tomó el estilo de encabezado y firmas para `GENERAR_INFORME`.

---

## 2. Errores del archivo anterior (corregidos)

El `REPORTE_ASISTENCIA` original tenía defectos que este sistema corrige:

1. **Referencias cruzadas rotas/desalineadas** en `REPORTE DIARIO`
   (p. ej. `C22` apuntaba a `PERSONAL!B9`, `C23` a `B9`, `C26` a `B16`…), que mezclaban
   nombres de trabajadores. → Ahora el reporte se llena por VBA desde la capa PROCESO.
2. **Horas de entrada escritas a mano** (`D18 = 08:04:39`) mezcladas con fórmulas.
   → Las horas provienen del procesamiento de DATA_HUELLERO.
3. **Dependencia de `MAXIFS`** (incompatible con versiones antiguas y frágil).
   → La lógica crítica se calcula en VBA; se evita `MAXIFS`.
4. **Nómina con duplicados** (LUCERO VERDE ENRIQUEZ repetida en filas 9 y 10).
   → Nómina depurada; la identidad se basa en ID PERSONAL + DNI + ID HUELLERO.
5. **Tolerancia por minutos, no por segundos.** → Comparación exacta al segundo.
6. **FALTA automática** ante ausencia. → Ahora `PENDIENTE DE REVISIÓN` (lo decide el responsable).

---

## 3. Lógica de tolerancia (requisitos 11, 12, 58)

Se compara en **segundos enteros**, sin `ROUND`:

```
marcaSeg  = Hora*3600 + Minuto*60 + Segundo (de la marca)
oficialSeg= Hora*3600 + Minuto*60 + Segundo (de la hora oficial del día)
corteSeg  = oficialSeg + tolerancia*60
PUNTUAL  si marcaSeg <= corteSeg
TARDANZA si marcaSeg >  corteSeg
```

Resultado garantizado (verificado en `PRUEBAS_UNITARIAS`):
`07:35:00 PUNTUAL / 07:35:01 TARDANZA`, `08:05:00 PUNTUAL / 08:05:01 TARDANZA`,
`15:05:00 PUNTUAL / 15:05:01 TARDANZA`.

La hora oficial del día sale de CONFIG por día de la semana, y puede ser sobrescrita
por **HORARIOS** (horario personalizado > general).

---

## 4. Clasificación de marcaciones (requisitos 13, 14, 15, 59)

Para cada trabajador y fecha:
1. Se recolectan **todas** las marcas del día (sin tocar el RAW).
2. Se ordenan cronológicamente.
3. Frontera mañana/tarde = punto medio entre *salida almuerzo* (13:00) y *entrada tarde* (15:00),
   es decir ~14:00 (configurable).
4. `ENTRADA MAÑANA` = primera de la mañana; `SALIDA ALMUERZO` = última de la mañana.
   `ENTRADA TARDE` = primera de la tarde; `SALIDA FINAL` = última de la tarde.
5. Marcas a **menos de 60 s** → bandera `DUPLICIDAD DE MARCACIÓN` (se conserva todo en RAW).

---

## 5. Garantías de conservación (requisitos 5, 8, 23, 47, 54, 67)

- **Importar es append-only**: se agrega bajo lo existente con **clave única**
  `ID + fecha/hora + estado + dispositivo`; reimportar no duplica.
- **La nómina es independiente del huellero**: `PROCESAR` recorre PERSONAL, no DATA;
  un trabajador agregado a mano siempre aparece.
- **PROCESO** respeta decisiones manuales previas (no las pisa al reprocesar).
- **Meses cerrados** (registrados en HISTORICO) **no se reprocesan** ni se regeneran.
- **Auditoría** de cada cambio manual; **respaldo** antes de importar/cerrar/guardar;
  **LOG_ERRORES** para errores controlados (nunca se cierra Excel ni se borra data).

---

## 6. Limitaciones técnicas y alternativas (requisito 64)

| Tema | Limitación real | Alternativa adoptada |
|---|---|---|
| **Incrustar VBA automáticamente** | El binario `vbaProject.bin` de un `.xlsm` solo lo genera Excel; no puede crearse de forma fiable fuera de Excel. | Se entrega libro `.xlsx` + código VBA importable + macro `INSTALAR`. Setup único de ~10 min documentado. |
| **UserForms (.frm/.frx)** | Los `.frm/.frx` binarios no pueden autorizarse/probarse fuera de Excel sin riesgo de corrupción. | La interacción se resuelve con **menú de botones + hojas de captura con listas desplegables + diálogos guiados** (100% fiables). El registro de incidencias cumple todos los campos pedidos (fecha, trabajador, tipo, observación, documento, usuario, fecha/hora). Si se desean UserForms visuales, pueden añadirse después sin cambiar la lógica. |
| **Layout exacto del cuadro histórico** | Cada mes del archivo original tiene una disposición distinta hecha a mano. | `GENERAR_CUADRO_GENERAL` reproduce fielmente la **gramática** del formato oficial (N/DNI/Nombre, M/T por día, cabeceras de día, DESCANSO combinado, códigos con color, cuadro resumen y leyenda) generándolo por mes/año para 2026–2029 y posteriores. |
| **Logos/imágenes institucionales** | No se dispone de los archivos de logo. | El encabezado se arma con los datos de CONFIG; basta pegar el logo una vez en el encabezado de impresión si se requiere. |

---

## 7. Catálogo de estados (editable sin tocar VBA)

`CATALOGO_ESTADOS`: `CÓDIGO | DESCRIPCIÓN | CUENTA PARA REMUNERACIÓN | COLOR | PRIORIDAD |
¿ES INCIDENCIA? | ¿ES ASISTENCIA? | NOTA REMUNERACIÓN`. La prioridad resuelve conflictos
(DESCANSO/FERIADO > COMISIÓN > DESC. MÉDICO > PERMISO > CLASE > NO PUSO HUELLA > FALTA).

---

## 8. Reconstrucción del libro base

Si se necesita regenerar `dist/SISTEMA_ASISTENCIA_AUDICONTA.xlsx`:

```bash
pip install openpyxl
python build/build_workbook.py
```
