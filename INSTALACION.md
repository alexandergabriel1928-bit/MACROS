# INSTALACIÓN — SISTEMA DE CONTROL DE ASISTENCIA AUDICONTA

Tiempo estimado: **10 minutos**. No necesita saber programar.

---

## Paso 0. Requisitos
- Excel 2019, 2021 o Microsoft 365 (Windows recomendado).
- Permitir macros al abrir el archivo.

---

## Paso 1. Convertir el libro base a macro-habilitado
1. Abra `dist/SISTEMA_ASISTENCIA_AUDICONTA.xlsx`.
2. **Archivo → Guardar como**.
3. En *Tipo*, elija **“Libro de Excel habilitado para macros (*.xlsm)”**.
4. Guarde como `SISTEMA_ASISTENCIA_AUDICONTA.xlsm`.

> ¿Por qué este paso? El código VBA solo puede incrustarse dentro de un `.xlsm` desde
> el propio Excel (ver `docs/DECISIONES_TECNICAS.md`, sección *Limitaciones*).

---

## Paso 2. Importar el código VBA
1. Con el `.xlsm` abierto, pulse **`Alt + F11`** (abre el editor de VBA).
2. Menú **Archivo → Importar archivo…** (o `Ctrl + M`).
3. Importe **uno por uno** todos los archivos de la carpeta `vba/`:
   - Todos los `mod*.bas` (18 módulos).
4. Para los **eventos del libro** (`ThisWorkbook.cls`):
   - Opción A (recomendada): en el editor, abra el objeto **ThisWorkbook** (doble clic en el
     árbol izquierdo) y **pegue** el contenido de `vba/ThisWorkbook.cls` **debajo** de los
     `Attribute...` (copie solo desde `Option Explicit` en adelante).
   - Opción B: Importar `ThisWorkbook.cls` directamente si su Excel lo permite.
5. Verifique que no haya errores: menú **Depuración → Compilar VBAProject**.

> Si al importar aparece *“No se puede obtener acceso al proyecto de VBA…”*, active:
> **Archivo → Opciones → Centro de confianza → Configuración → Configuración de macros →
> Confiar en el acceso al modelo de objetos de proyectos de VBA**.

---

## Paso 3. Instalar (una sola vez)
1. Vuelva a Excel (`Alt + F11` para alternar).
2. **`Alt + F8`** → seleccione **`INSTALAR`** → **Ejecutar**.
   - Crea los botones del menú.
   - Oculta las hojas técnicas.
   - Protege `CONFIG` dejando editables solo los valores.
3. Verá el menú principal con botones.

---

## Paso 4. Probar la lógica (recomendado)
1. **`Alt + F8`** → **`PRUEBAS_UNITARIAS`** → **Ejecutar**.
2. Revise la hoja **PRUEBAS**: deben salir **todas OK** (incluye los casos de
   tolerancia al segundo 07:35:00 / 07:35:01, 08:05:00 / 08:05:01, 15:05:00 / 15:05:01).

---

## Paso 5. Configurar
Abra **CONFIGURACIÓN** (botón del menú) y revise/edite las celdas gris claro:
empresa, horarios por día, tolerancia, hora de la tarde, periodo activo, ruta de
respaldo y (opcional) parámetros de remuneración.

---

## Mantenimiento
- Para ver/editar hojas técnicas: `Alt + F8` → **`MODO_MANTENIMIENTO`**.
- Para volver al modo normal: ejecute **`INSTALAR`** de nuevo.

## Problemas frecuentes
| Síntoma | Solución |
|---|---|
| Los botones no hacen nada | Ejecute `INSTALAR`. Confirme que importó todos los `mod*.bas`. |
| “Error de compilación: Sub o Function no definida” | Falta un módulo por importar. Importe el que falte y **Compilar VBAProject**. |
| No crea respaldo | Guarde primero el `.xlsm` en una carpeta; defina `RUTA DE RESPALDO` en CONFIG. |
| Las macros están deshabilitadas | Barra amarilla al abrir → **Habilitar contenido**. |
