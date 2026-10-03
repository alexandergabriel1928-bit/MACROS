# MANUAL DE USO — SISTEMA DE CONTROL DE ASISTENCIA AUDICONTA

Responsable del control: **VERDE ENRIQUEZ LUCERO BRITNEY** (configurable en CONFIG).

---

## 1. El menú principal (`01_MENU`)

Botones disponibles:

| Botón | Qué hace |
|---|---|
| **IMPORTAR HUELLERO** | Carga la exportación del marcador (XLS/XLSX/CSV). Append-only, sin duplicar. |
| **PROCESAR ASISTENCIA** | Calcula la asistencia del día (fecha de trabajo en CONFIG). |
| **ASISTENCIA DEL DIA** | Rellena y muestra el REPORTE DIARIO de una fecha. |
| **CUADRO GENERAL** | Genera la hoja del mes con el formato oficial (M/T por día). |
| **CONSOLIDADO MENSUAL** | Resumen por trabajador del mes (base de remuneraciones). |
| **INCIDENCIAS** | Abre la hoja para registrar permisos, clases, comisiones, etc. |
| **PERSONAL** | Nómina maestra (altas/bajas/datos). |
| **CONFIGURACION** | Parámetros de empresa, horarios, tolerancia, periodo, remuneración. |
| **HISTORICO** | Meses cerrados (solo lectura). |
| **AUDITORIA** | Registro de modificaciones manuales. |
| **ALERTAS / VALIDAR** | Detecta inconsistencias del día. |
| **GENERAR INFORME** | Informe mensual para Gerencia. |
| **GUARDAR Y RESPALDAR** | Guarda y crea copia de respaldo. |
| **CERRAR MES** | Cierra y congela el periodo; archiva en histórico. |

El **tablero** del menú muestra: personal activo, presentes, puntuales, tardanzas,
faltas, incidencias, pendientes, % puntualidad, % asistencia, periodo y estado.

---

## 2. Flujo diario

```
ABRIR ARCHIVO
   ↓
CONFIG → poner FECHA DE TRABAJO (por defecto HOY)
   ↓
IMPORTAR HUELLERO   (elige el archivo del marcador)
   ↓
PROCESAR ASISTENCIA
   ↓
ALERTAS / VALIDAR   (revisar pendientes e inconsistencias)
   ↓
INCIDENCIAS         (completar permisos/clases/comisiones si los hay)
   ↓
ASISTENCIA DEL DIA  (revisar el reporte)
   ↓
GUARDAR Y RESPALDAR
```

### Detalle
1. **Fecha de trabajo**: en CONFIG, celda *FECHA DE TRABAJO*. Puede consultar días
   anteriores cambiándola; **no** sobrescribe históricos.
2. **Importar**: seleccione el archivo del huellero. El sistema detecta columnas
   automáticamente (Número, Nombre, Tiempo, Estado, Dispositivos, Tipo de Registro),
   agrega solo lo nuevo y avisa de trabajadores no registrados (puede agregarlos).
3. **Procesar**: ordena marcaciones y determina ENTRADA MAÑANA / SALIDA ALMUERZO /
   ENTRADA TARDE / SALIDA FINAL por contexto horario (funciona aunque el ESTADO venga
   vacío). Aplica tolerancia **al segundo** y marca PUNTUAL/TARDANZA.
4. **Ausencias**: quien no marca queda **PENDIENTE DE REVISIÓN** (no FALTA automática).
5. **Incidencias**: en la hoja INCIDENCIAS indique Fecha, Trabajador, Tipo (menú:
   CLASE, PERMISO, NO PUSO HUELLA, DESCANSO MÉDICO, COMISIÓN, etc.), Observación y
   documento. Al procesar/registrar, la condición final se actualiza y queda **auditada**.
   Una incidencia justificada nunca se vuelve falta.

---

## 3. Flujo de cierre de mes

```
REVISAR PENDIENTES (ALERTAS)
   ↓
CONSOLIDADO MENSUAL
   ↓
CUADRO GENERAL
   ↓
GENERAR INFORME
   ↓
CERRAR MES  → resumen + confirmación → respaldo + histórico (congela el mes)
```

Al **CERRAR MES**:
- Se muestra el resumen (puntuales, tardanzas, faltas, permisos, pendientes…).
- Si hay **pendientes**, avisa y permite VOLVER o CERRAR DE TODAS FORMAS.
- Crea respaldo, genera consolidado, guarda en HISTORICO_MENSUAL y marca el periodo
  como **CERRADO** (ya no se reprocesa).

---

## 4. Códigos del cuadro general

| Código | Significado | Color |
|---|---|---|
| `2` | Asistió / Puntual | Verde |
| `T` | Tardanza | Ámbar |
| `F` | Falta | Rojo |
| `PE` | Permiso | Azul claro |
| `S` | Suspensión | Naranja |
| `FE` | Feriado | Amarillo |
| `V` | Vacaciones | Verde claro |
| `C` | Comisión | Celeste |
| `CL` | Clase | Azul |
| `NH` | No puso huella | Gris |
| `DM` | Descanso médico | Salmón |
| `NF` | No firmó | Gris claro |
| `D` | Descanso | Gris |

Todos son editables en **CATALOGO_ESTADOS** (código, descripción, color, prioridad,
si cuenta para remuneración). Cambiar un código **no** requiere tocar el VBA.

---

## 5. Configuración maestra (CONFIG)

Edite solo las celdas gris claro:
- **Empresa**: nombre, RUC, dirección, celulares, correo, área, responsable, gerente, código de formato, versión.
- **Horario mañana por día** (LUN…SÁB; DOMINGO = DESCANSO).
- **Tolerancia** (minutos), **ingreso tarde**, **tolerancia tarde**, **salida**.
- **Periodo**: fecha de trabajo, mes/año activos, estado (ABIERTO/CERRADO).
- **Respaldo**: ruta y límite de registros.
- **Remuneración** (opcional): si APLICAR DESCUENTOS = NO, el sistema solo entrega
  indicadores; si = SI, usa los porcentajes configurados. Ningún porcentaje está en el código.

**Horario personalizado por trabajador**: hoja HORARIOS (prioridad sobre el general).

---

## 6. Buenas prácticas
- Guarde y respalde al terminar cada día.
- Antes de cerrar el mes, resuelva los **PENDIENTES** desde ALERTAS/INCIDENCIAS.
- No borre filas de DATA_HUELLERO, PROCESO, AUDITORIA ni HISTORICO (son la trazabilidad).
- Si necesita corregir una condición, hágalo por **INCIDENCIAS** (queda auditado),
  no editando a mano las hojas técnicas.
