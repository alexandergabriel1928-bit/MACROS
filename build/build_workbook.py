# -*- coding: utf-8 -*-
"""
Constructor del libro base del SISTEMA MAESTRO DE CONTROL DE ASISTENCIA - AUDICONTA PERU E.I.R.L.
--------------------------------------------------------------------------------------------------
Genera dist/SISTEMA_ASISTENCIA_AUDICONTA.xlsx con TODA la estructura, formato,
catalogos, nomina real, parametros, validaciones y rangos con nombre que necesita
el proyecto VBA. El codigo VBA (carpeta /vba) se importa luego y se ejecuta INSTALAR.

Este archivo NO contiene logica de negocio: solo arma el "esqueleto" visual y de datos.
La logica critica vive en VBA (ver /vba).

Requiere: openpyxl
"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                             NamedStyle, Protection)
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter, quote_sheetname

# ----------------------------------------------------------------------------
# PALETA (colores internos del sistema; NO son los de las plantillas oficiales)
# ----------------------------------------------------------------------------
AZUL      = "1F4E79"   # cabeceras institucionales
AZUL2     = "2E75B6"
AZUL_CLARO= "DDEBF7"
GRIS_HEAD = "44546A"
GRIS_ED   = "F2F2F2"   # celdas editables (gris claro)
GRIS_MED  = "D9D9D9"
VERDE     = "C6EFCE"   # puntual
VERDE_TXT = "006100"
AMBAR     = "FFEB9C"   # tardanza
AMBAR_TXT = "9C6500"
ROJO      = "FFC7CE"   # falta
ROJO_TXT  = "9C0006"
GRIS_PEND = "D9D9D9"   # pendiente
AZUL_JUS  = "BDD7EE"   # justificado
BLANCO    = "FFFFFF"
NEGRO     = "000000"
AMARILLO_MENU = "FFD966"

def fill(c): return PatternFill("solid", fgColor=c)
def font(sz=11, b=False, color=NEGRO, name="Calibri", it=False):
    return Font(name=name, size=sz, bold=b, color=color, italic=it)
thin = Side(style="thin", color="BFBFBF")
med  = Side(style="medium", color="808080")
border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
border_box = Border(left=med, right=med, top=med, bottom=med)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left   = Alignment(horizontal="left", vertical="center", wrap_text=True)
leftv  = Alignment(horizontal="left", vertical="top", wrap_text=True)
right  = Alignment(horizontal="right", vertical="center")

wb = Workbook()

# ===========================================================================
#  CONSTANTES DE NOMBRES DE HOJAS
# ===========================================================================
SH_MENU   = "01_MENU"
SH_CONFIG = "CONFIG"
SH_PERS   = "PERSONAL"
SH_CAT    = "CATALOGO_ESTADOS"
SH_HOR    = "HORARIOS"
SH_CAL    = "CALENDARIO"
SH_DATA   = "DATA_HUELLERO"
SH_PROC   = "PROCESO"
SH_RDIA   = "REPORTE DIARIO"
SH_CONS   = "CONSOLIDADO"
SH_CUAD   = "CUADRO_GENERAL"
SH_INC    = "INCIDENCIAS"
SH_AUD    = "AUDITORIA"
SH_ALE    = "ALERTAS"
SH_HIST   = "HISTORICO_MENSUAL"
SH_LOG    = "LOG_ERRORES"

# ===========================================================================
#  DATOS REALES EXTRAIDOS DE LAS PLANTILLAS ADJUNTAS
#  Nomina autoritativa = CUADRO GENERAL (SETIEMBRE 2026) con DNI.
#  NOMBRE EN HUELLERO / ID HUELLERO = exportacion 01.10.2026 (donde coincide).
#  (N, DNI, APELLIDOS Y NOMBRES, NOMBRE_HUELLERO, ID_HUELLERO, CARGO, EMPRESA/AREA)
# ===========================================================================
ROSTER = [
    (1 , "71584311", "PEÑA JUSCAMAYTA AARON ROLLER",       "AARON  PEÑA JUSCAMAYTA",               "3",  "JEFE DE OPERACIONES",  "AUDICONTA PERU E.I.R.L."),
    (2 , "75222100", "TAMBO MANCHACHI IVETH CECILIA",      "IVETH TAMBO MANCHACHI",                "4",  "SECTORISTA CONTABLE",  "AUDICONTA PERU E.I.R.L."),
    (3 , "75217868", "AVELLANEDA AQUINO GABRIEL ALEXANDER","GABRIEL ALEXANDER AVELLANEDA AQUINO",  "5",  "SECTORISTA CONTABLE",  "AUDICONTA PERU E.I.R.L."),
    (4 , "74890941", "QUINCHO NUÑEZ OSCAR JHOEL",          "OSCAR QUINCHO NUÑEZ",                  "6",  "SECTORISTA CONTABLE",  "AUDICONTA PERU E.I.R.L."),
    (5 , "70933779", "VERDE ENRIQUEZ LUCERO BRITNEY",      "LUCERO VERDE ENRIQUEZ",                "7",  "RESPONSABLE DE RR.HH.","AUDICONTA PERU E.I.R.L."),
    (6 , "62021914", "FARFAN ROMAN JULIETTE KRYSTAL",      "JULIETTE KRYSTAL  FARFAN ROMAN",       "15", "ASISTENTE CONTABLE",   "AUDICONTA PERU E.I.R.L."),
    (7 , "60974103", "RUIZ DE LA PEÑA FABRICIO RICARDO",   "FABRICIO RICARDO  RUIZ DE LA PEÑA",    "14", "ASISTENTE CONTABLE",   "AUDICONTA PERU E.I.R.L."),
    (8 , "61314586", "ARNAO PALOMINO FRANKO ALEXANDER",    "FRANKO ALEXANDER  ANCO PALOMINO",      "13", "ASISTENTE CONTABLE",   "AUDICONTA PERU E.I.R.L."),
    (9 , "61351206", "TANZ HUICHO LUIS MIGUEL",            "LUIS MIGUEL TANZ HUICHO",              "",   "ASISTENTE CONTABLE",   "AUDICONTA PERU E.I.R.L."),
    (10, "62403330", "ALDERETE SANTIAGO ANGGIE DAYANA",    "",                                     "",   "ASISTENTE CONTABLE",   "AUDICONTA PERU E.I.R.L."),
    (11, "77135062", "VELIZ ROJAS DINA",                   "DINA  VELIZ ROJAS",                    "16", "ASISTENTE CONTABLE",   "AUDICONTA PERU E.I.R.L."),
    (12, "61209532", "AQUINO QUISPE LUIS MIGUEL",          "",                                     "",   "ASISTENTE CONTABLE",   "AUDICONTA PERU E.I.R.L."),
    (13, "76633122", "RAVICHAGUA VALDIVIA HEYZON BRANDON", "",                                     "",   "ASISTENTE CONTABLE",   "AUDICONTA PERU E.I.R.L."),
    (14, "75217906", "TERACCAYA SAMANIEGO CAMILA SHANTAL", "",                                     "",   "ASISTENTE CONTABLE",   "AUDICONTA PERU E.I.R.L."),
    (15, "61302056", "ROMERO MENDOZA MARI LUZ",            "",                                     "",   "ASISTENTE CONTABLE",   "AUDICONTA PERU E.I.R.L."),
    (16, "75514867", "AQUINO NUÑEZ KEVIN ANTONY",          "KEVIN AQUINO NUÑEZ",                   "",   "AGENTE DE SEGURIDAD",  "SEG PRO FORCE S.A.C."),
    (17, "76320303", "ROMERO MEZA JACZUMY MERCEDES",       "JACZUMY  MEZA ROMERO",                 "11", "AGENTE DE SEGURIDAD",  "SEG PRO FORCE S.A.C."),
    (18, "61426723", "MAGUIÑA MARTINEZ ALONDRA MELANI",    "ALONDRA MELANI  MAGUIÑA MARTINEZ",     "10", "AGENTE DE SEGURIDAD",  "SEG PRO FORCE S.A.C."),
    (19, "61111564", "BARRIOS ZAMBRANO DANIEL ANDERSON",   "DANIEL ANDERSON   BARRIOS ZAMBRANO",   "9",  "AGENTE DE SEGURIDAD",  "SEG PRO FORCE S.A.C."),
]

# CATALOGO DE ESTADOS / CODIGOS (basado en el cuadro oficial y la leyenda)
# (Codigo, Descripcion, CuentaRemuneracion, ColorHex, Prioridad, EsIncidencia, EsAsistencia, NotaRemuneracion)
CATALOGO = [
    ("2",  "ASISTIO / PUNTUAL",      "PAGA",         VERDE,     80, "NO", "SI", "Dia efectivo"),
    ("T",  "TARDANZA",               "DESCUENTA",    AMBAR,     75, "NO", "SI", "Tarde: 30% R.D. (configurable)"),
    ("F",  "FALTA",                  "DESCUENTA",    ROJO,      10, "NO", "NO", "Falta: 1 R.D."),
    ("PE", "PERMISO",                "SEGUN POLIT.", AZUL_JUS,  50, "SI", "NO", "Permiso: segun politica"),
    ("S",  "SUSPENSION",             "DESCUENTA",    "F8CBAD",  40, "SI", "NO", "Suspension: 2 R.D."),
    ("FE", "FERIADO",                "PAGA",         "FFF2CC",  95, "NO", "NO", "Feriado: si se paga R.D."),
    ("V",  "VACACIONES",             "PAGA",         "E2EFDA",  90, "SI", "NO", "Vacaciones: si se paga R.D."),
    ("C",  "COMISION",               "PAGA",         "DDEBF7",  60, "SI", "SI", "Comision: si se paga R.D."),
    ("CL", "CLASE",                  "SEGUN POLIT.", "D9E1F2",  55, "SI", "NO", "Clase / estudios"),
    ("NH", "NO PUSO HUELLA",         "REVISAR",      "D9D9D9",  30, "SI", "NO", "No marco - revisar"),
    ("DM", "DESCANSO MEDICO",        "SEGUN POLIT.", "FCE4D6",  65, "SI", "NO", "Descanso medico"),
    ("NF", "NO FIRMO",               "DESCUENTA",    "F2F2F2",  20, "SI", "NO", "Sin firma: 5% R.D."),
    ("D",  "DESCANSO",               "NO APLICA",    "BFBFBF",  97, "NO", "NO", "Dia de descanso"),
    ("J",  "JUSTIFICADO",            "SEGUN POLIT.", AZUL_JUS,  45, "SI", "NO", "Justificado con documento"),
    ("INC","INCOMPLETO",             "REVISAR",      "FFF2CC",  25, "NO", "NO", "Marcacion incompleta"),
    ("PR", "PENDIENTE DE REVISION",  "REVISAR",      "D9D9D9",  15, "NO", "NO", "Requiere decision del responsable"),
]

# Observaciones permitidas (menu desplegable base) - ampliable desde CONFIG
OBSERVACIONES = ["CLASE", "PERMISO", "NO PUSO HUELLA", "DESCANSO MEDICO", "COMISION",
                 "DESCANSO", "FERIADO", "VACACIONES", "SUSPENSION", "JUSTIFICADO", "OTRO"]
CONDICIONES = ["PUNTUAL","TARDANZA","FALTA","PERMISO","CLASE","NO PUSO HUELLA",
               "DESCANSO MEDICO","COMISION","DESCANSO","FERIADO","JUSTIFICADO",
               "VACACIONES","SUSPENSION","INCOMPLETO","PENDIENTE DE REVISION"]

defined = []  # (name, "'Sheet'!$A$1")
def add_name(name, sheet, cell):
    ref = "{}!{}".format(quote_sheetname(sheet), cell)
    defined.append((name, ref))

# ===========================================================================
#  HOJA 01_MENU
# ===========================================================================
ws = wb.active
ws.title = SH_MENU
ws.sheet_view.showGridLines = False
for col,w in {"A":2,"B":30,"C":30,"D":30,"E":30,"F":2,"G":34}.items():
    ws.column_dimensions[col].width = w
ws.merge_cells("B2:E2")
c = ws["B2"]; c.value = "SISTEMA MAESTRO DE CONTROL DE ASISTENCIA"
c.font = font(20,True,BLANCO); c.alignment = center; c.fill = fill(AZUL)
ws.row_dimensions[2].height = 40
ws.merge_cells("B3:E3")
c = ws["B3"]; c.value = "AUDICONTA PERU E.I.R.L.   -   RUC 20568153414   -   F-RH-01 v01"
c.font = font(11,True,BLANCO); c.alignment = center; c.fill = fill(AZUL2)
ws.merge_cells("B4:E4")
c = ws["B4"]; c.value = ("Los botones del menu se crean automaticamente al ejecutar la macro INSTALAR "
                         "(ver INSTALACION.md). Mientras tanto use Alt+F8 para ejecutar las macros.")
c.font = font(9,it=True,color="808080"); c.alignment = center
ws.row_dimensions[4].height = 26

# Marcadores donde INSTALAR colocara los botones (grilla logica)
MENU_BUTTONS = [
    ("IMPORTAR HUELLERO",      "IMPORTAR_HUELLERO"),
    ("PROCESAR ASISTENCIA",    "PROCESAR_ASISTENCIA"),
    ("ASISTENCIA DEL DIA",     "VER_ASISTENCIA_DIA"),
    ("CUADRO GENERAL",         "GENERAR_CUADRO_GENERAL"),
    ("CONSOLIDADO MENSUAL",    "GENERAR_CONSOLIDADO"),
    ("INCIDENCIAS",            "ABRIR_INCIDENCIAS"),
    ("PERSONAL",               "ABRIR_PERSONAL"),
    ("CONFIGURACION MAESTRA",  "ABRIR_CONFIGURACION"),
    ("HISTORICO",              "ABRIR_HISTORICO"),
    ("AUDITORIA",              "ABRIR_AUDITORIA"),
    ("ALERTAS / VALIDAR",      "GENERAR_ALERTAS"),
    ("GENERAR INFORME",        "GENERAR_INFORME"),
    ("GUARDAR Y RESPALDAR",    "GUARDAR_Y_RESPALDAR"),
    ("CERRAR MES",             "CERRAR_MES"),
]
# tabla de referencia de botones (para INSTALAR y para el usuario)
ws["B6"] = "BOTON"; ws["C6"] = "MACRO (Alt+F8)"
for cc in ("B6","C6"):
    ws[cc].font = font(10,True,BLANCO); ws[cc].fill = fill(GRIS_HEAD); ws[cc].alignment = center
r = 7
for label, macro in MENU_BUTTONS:
    ws.cell(r,2,label).font = font(11,True,AZUL); ws.cell(r,2).fill = fill(AMARILLO_MENU)
    ws.cell(r,2).alignment = left; ws.cell(r,2).border = border_all
    ws.cell(r,3,macro).font = font(10); ws.cell(r,3).alignment = left; ws.cell(r,3).border = border_all
    ws.row_dimensions[r].height = 22
    r += 1
# Mini-dashboard (se llena por VBA)
ws.merge_cells("G6:G6")
ws["G6"] = "TABLERO (se actualiza al procesar)"
ws["G6"].font = font(11,True,BLANCO); ws["G6"].fill = fill(AZUL); ws["G6"].alignment = center
dash = ["PERSONAL ACTIVO","PRESENTES HOY","PUNTUALES","TARDANZAS","FALTAS",
        "INCIDENCIAS","PENDIENTES","% PUNTUALIDAD","% ASISTENCIA","PERIODO ACTIVO","ESTADO PERIODO"]
dr = 7
for label in dash:
    ws.cell(dr,7, label + ":").font = font(10,True); ws.cell(dr,7).alignment = left
    ws.cell(dr,7).border = border_all; ws.cell(dr,7).fill = fill(AZUL_CLARO)
    # el valor se escribe por VBA en columna H (fusionada visualmente). Dejamos H.
    ws.cell(dr,8).border = border_all
    add_name("dash_" + label.replace(" ","_").replace("%","PCT").replace("Ñ","N"),
             SH_MENU, "$H${}".format(dr))
    dr += 1
ws.column_dimensions["H"].width = 14

# ===========================================================================
#  HOJA CONFIG  (CONFIGURACION MAESTRA)  -> rangos con nombre cfg_*
# ===========================================================================
ws = wb.create_sheet(SH_CONFIG)
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 34
ws.column_dimensions["B"].width = 42
ws.column_dimensions["C"].width = 50
def title_row(ws, row, text, span="A{r}:C{r}"):
    rng = span.format(r=row)
    ws.merge_cells(rng)
    c = ws.cell(row,1,text); c.font = font(12,True,BLANCO); c.fill = fill(AZUL); c.alignment = left
    ws.row_dimensions[row].height = 22

def kv(ws, row, label, value, name=None, help_="", numfmt=None, editable=True):
    ws.cell(row,1,label).font = font(10,True); ws.cell(row,1).alignment = left
    cel = ws.cell(row,2,value)
    cel.alignment = left
    if editable:
        cel.fill = fill(GRIS_ED); cel.protection = Protection(locked=False)
    if numfmt: cel.number_format = numfmt
    if help_:
        ws.cell(row,3,help_).font = font(9,it=True,color="808080"); ws.cell(row,3).alignment = left
    if name: add_name(name, SH_CONFIG, "$B${}".format(row))
    return row+1

r = 1
title_row(ws, r, "CONFIGURACION MAESTRA  -  AUDICONTA PERU E.I.R.L."); r += 1
ws.cell(r,1,"Edite solo las celdas gris claro. No modifique etiquetas ni codigo VBA.").font = font(9,it=True,color="9C0006"); r += 2

title_row(ws, r, "1. DATOS DE LA EMPRESA"); r += 1
r = kv(ws, r, "EMPRESA / INSTITUCION:", "AUDICONTA PERU E.I.R.L.", "cfg_Empresa")
r = kv(ws, r, "RUC:", "20568153414", "cfg_RUC")
r = kv(ws, r, "DIRECCION:", "Av. La Rivera N° 944 - La Merced - Chanchamayo", "cfg_Direccion")
r = kv(ws, r, "CELULARES:", "943889886 - 948024186", "cfg_Celulares")
r = kv(ws, r, "CORREO ELECTRONICO:", "audicontaperu.eirl@gmail.com", "cfg_Correo")
r = kv(ws, r, "AREA / SEDE:", "OFICINA PRINCIPAL", "cfg_AreaSede")
r = kv(ws, r, "RESPONSABLE DEL CONTROL:", "VERDE ENRIQUEZ LUCERO BRITNEY", "cfg_Responsable")
r = kv(ws, r, "GERENTE:", "LUIS ENRIQUE ESPINOZA QUISPE", "cfg_Gerente")
r = kv(ws, r, "DNI GERENTE:", "43530814", "cfg_DNIGerente")
r = kv(ws, r, "CODIGO DEL FORMATO:", "F-RH-01", "cfg_CodFormato")
r = kv(ws, r, "VERSION:", "01", "cfg_Version")
r += 1

title_row(ws, r, "2. HORARIO DE INGRESO - MAÑANA (por dia)"); r += 1
dias = [("LUNES","cfg_EntLun", dt.time(7,30)),
        ("MARTES","cfg_EntMar", dt.time(8,0)),
        ("MIERCOLES","cfg_EntMie", dt.time(8,0)),
        ("JUEVES","cfg_EntJue", dt.time(8,0)),
        ("VIERNES","cfg_EntVie", dt.time(8,0)),
        ("SABADO","cfg_EntSab", dt.time(7,30)),
        ("DOMINGO","cfg_EntDom", "DESCANSO")]
for d,nm,val in dias:
    r = kv(ws, r, "INGRESO "+d+":", val, nm,
           help_="Use formato HH:MM:SS. Escriba DESCANSO si no se labora.",
           numfmt=("HH:MM:SS" if isinstance(val, dt.time) else None))
r += 1

title_row(ws, r, "3. TOLERANCIA Y HORARIO DE TARDE / SALIDA"); r += 1
r = kv(ws, r, "TOLERANCIA MAÑANA (MINUTOS):", 5, "cfg_TolMin",
       help_="Ej: 5 -> hasta 07:35:00 PUNTUAL; 07:35:01 TARDANZA (exacto al segundo).")
r = kv(ws, r, "INGRESO TARDE:", dt.time(15,0), "cfg_EntTarde", numfmt="HH:MM:SS")
r = kv(ws, r, "TOLERANCIA TARDE (MINUTOS):", 5, "cfg_TolTardeMin",
       help_="Ej: 5 -> hasta 15:05:00 PUNTUAL; 15:05:01 TARDANZA.")
r = kv(ws, r, "SALIDA ALMUERZO (REFERENCIA):", dt.time(13,0), "cfg_SalAlmuerzo", numfmt="HH:MM:SS")
r = kv(ws, r, "HORA OFICIAL DE SALIDA:", dt.time(17,0), "cfg_SalTarde", numfmt="HH:MM:SS")
r += 1

title_row(ws, r, "4. PERIODO DE TRABAJO"); r += 1
r = kv(ws, r, "FECHA DE TRABAJO:", dt.date(2026,10,1), "cfg_FechaTrabajo",
       help_="Fecha del dia que se procesa/consulta. La macro propone HOY.", numfmt="DD/MM/YYYY")
r = kv(ws, r, "MES ACTIVO (1-12):", 10, "cfg_PeriodoMes")
r = kv(ws, r, "AÑO ACTIVO:", 2026, "cfg_PeriodoAnio")
r = kv(ws, r, "ESTADO DEL PERIODO:", "ABIERTO", "cfg_EstadoPeriodo",
       help_="ABIERTO o CERRADO. Un periodo CERRADO queda congelado.")
r += 1

title_row(ws, r, "5. RESPALDO Y LIMITES"); r += 1
r = kv(ws, r, "RUTA DE RESPALDO:", "", "cfg_RutaBackup",
       help_="Carpeta para backups. Vacio = misma carpeta del archivo \\BACKUPS.")
r = kv(ws, r, "MAX REGISTROS DATA_HUELLERO:", 100000, "cfg_MaxRegistros")
r += 1

title_row(ws, r, "6. REMUNERACION (opcional - solo si se configura)"); r += 1
r = kv(ws, r, "APLICAR DESCUENTOS:", "NO", "cfg_AplicarDesc",
       help_="SI / NO. Si es NO, el sistema solo entrega indicadores, no calcula descuentos.")
r = kv(ws, r, "DIAS DEL MES PARA R.D.:", 30, "cfg_DiasMesRD",
       help_="Divisor para remuneracion diaria (R.D. = Sueldo / dias).")
r = kv(ws, r, "DESCUENTO POR TARDANZA (% R.D.):", 30, "cfg_DescTardanzaPct")
r = kv(ws, r, "DESCUENTO POR FALTA (R.D.):", 1, "cfg_DescFaltaRD")
r = kv(ws, r, "DESCUENTO POR SUSPENSION (R.D.):", 2, "cfg_DescSuspRD")
r = kv(ws, r, "DESCUENTO SIN FIRMA (% R.D.):", 5, "cfg_DescNoFirmaPct")
r += 1
ws.protection.sheet = False  # INSTALAR la protege dejando libres las celdas gris claro

# ===========================================================================
#  HOJA PERSONAL
# ===========================================================================
ws = wb.create_sheet(SH_PERS)
ws.sheet_view.showGridLines = False
headers = ["ID PERSONAL","N°","DNI","APELLIDOS Y NOMBRES","NOMBRE EN HUELLERO",
           "ID HUELLERO","CARGO","AREA / EMPRESA","FECHA INGRESO","FECHA CESE",
           "TIPO DE JORNADA","HORARIO PERSONALIZADO","ACTIVO","OBSERVACIONES"]
widths = [12,5,12,38,38,11,24,26,13,13,16,20,9,30]
ws.merge_cells("A1:N1")
c = ws["A1"]; c.value = "NOMINA MAESTRA DEL PERSONAL (independiente del huellero - NO se reconstruye al importar)"
c.font = font(12,True,BLANCO); c.fill = fill(AZUL); c.alignment = left
ws.row_dimensions[1].height = 22
for i,h in enumerate(headers, start=1):
    col = get_column_letter(i)
    ws.column_dimensions[col].width = widths[i-1]
    cel = ws.cell(2,i,h); cel.font = font(10,True,BLANCO); cel.fill = fill(GRIS_HEAD)
    cel.alignment = center; cel.border = border_all
ws.row_dimensions[2].height = 30
rr = 3
for (n, dni, apn, hname, hid, cargo, area) in ROSTER:
    idp = "P{:03d}".format(n)
    vals = [idp, n, dni, apn, hname, hid, cargo, area,
            dt.date(2026,1,1), "", "ESTANDAR", "", "SI", ""]
    for i,v in enumerate(vals, start=1):
        cel = ws.cell(rr,i,v); cel.border = border_all; cel.alignment = left
        cel.font = font(10)
        if i in (3,6):  # DNI, ID HUELLERO -> texto
            cel.number_format = "@"
        if i in (9,10): cel.number_format = "DD/MM/YYYY"
    rr += 1
LAST_PERS = rr - 1
# filas vacias editables para nuevos trabajadores
for extra in range(40):
    for i in range(1,15):
        ws.cell(rr,i).border = border_all; ws.cell(rr,i).font = font(10)
    ws.cell(rr,3).number_format="@"; ws.cell(rr,6).number_format="@"
    rr += 1
PERS_MAXROW = rr - 1
ws.cell(rr+1,1,"LEYENDA: La nomina es permanente. Un trabajador agregado aqui permanece aunque no "
        "aparezca en el huellero. Use el boton PERSONAL para altas controladas.").font = font(9,it=True,color="808080")
# Validacion ACTIVO (col M) y TIPO JORNADA (col K)
dv_act = DataValidation(type="list", formula1='"SI,NO"', allow_blank=True); ws.add_data_validation(dv_act)
dv_act.add("M3:M{}".format(PERS_MAXROW))
dv_jor = DataValidation(type="list", formula1='"ESTANDAR,PERSONALIZADO,MEDIO TIEMPO"', allow_blank=True); ws.add_data_validation(dv_jor)
dv_jor.add("K3:K{}".format(PERS_MAXROW))
add_name("rng_Personal", SH_PERS, "$A$3:$N${}".format(PERS_MAXROW))
add_name("rng_PersonalUlt", SH_PERS, "$A${}".format(PERS_MAXROW))

# ===========================================================================
#  HOJA CATALOGO_ESTADOS
# ===========================================================================
ws = wb.create_sheet(SH_CAT)
ws.sheet_view.showGridLines = False
cheaders = ["CODIGO","DESCRIPCION","CUENTA PARA REMUNERACION","COLOR (hex)","PRIORIDAD",
            "¿ES INCIDENCIA?","¿ES ASISTENCIA?","NOTA REMUNERACION"]
cw = [10,26,24,12,10,15,15,34]
ws.merge_cells("A1:H1")
c = ws["A1"]; c.value="CATALOGO DE ESTADOS / CODIGOS (editable - sin tocar VBA)"
c.font=font(12,True,BLANCO); c.fill=fill(AZUL); c.alignment=left
for i,h in enumerate(cheaders,1):
    ws.column_dimensions[get_column_letter(i)].width = cw[i-1]
    cel=ws.cell(2,i,h); cel.font=font(10,True,BLANCO); cel.fill=fill(GRIS_HEAD); cel.alignment=center; cel.border=border_all
ws.row_dimensions[2].height=28
rr=3
for (cod,desc,rem,color,prio,inc,asi,nota) in CATALOGO:
    row=[cod,desc,rem,color,prio,inc,asi,nota]
    for i,v in enumerate(row,1):
        cel=ws.cell(rr,i,v); cel.border=border_all; cel.font=font(10); cel.alignment=left
    # pinta la celda CODIGO con su color
    ws.cell(rr,1).fill = fill(color); ws.cell(rr,1).alignment=center; ws.cell(rr,1).font=font(10,True)
    rr+=1
CAT_MAXROW=rr-1
add_name("rng_Catalogo", SH_CAT, "$A$3:$H${}".format(CAT_MAXROW))
# tabla de observaciones permitidas (col J) y condiciones (col L)
ws.cell(2,10,"OBSERVACIONES (menu)").font=font(10,True,BLANCO); ws.cell(2,10).fill=fill(GRIS_HEAD); ws.cell(2,10).alignment=center
ws.column_dimensions["J"].width=22
for i,o in enumerate(OBSERVACIONES, start=3):
    ws.cell(i,10,o).border=border_all; ws.cell(i,10).font=font(10)
add_name("rng_Observaciones", SH_CAT, "$J$3:$J${}".format(2+len(OBSERVACIONES)))
ws.cell(2,12,"CONDICIONES").font=font(10,True,BLANCO); ws.cell(2,12).fill=fill(GRIS_HEAD); ws.cell(2,12).alignment=center
ws.column_dimensions["L"].width=24
for i,o in enumerate(CONDICIONES, start=3):
    ws.cell(i,12,o).border=border_all; ws.cell(i,12).font=font(10)
add_name("rng_Condiciones", SH_CAT, "$L$3:$L${}".format(2+len(CONDICIONES)))

# ===========================================================================
#  HOJA HORARIOS  (horario estandar por dia ya esta en CONFIG; aqui: por trabajador)
# ===========================================================================
ws = wb.create_sheet(SH_HOR)
ws.sheet_view.showGridLines = False
ws.merge_cells("A1:J1")
c=ws["A1"]; c.value="HORARIOS PERSONALIZADOS POR TRABAJADOR (prioridad: PERSONALIZADO > GENERAL)"
c.font=font(12,True,BLANCO); c.fill=fill(AZUL); c.alignment=left
hh = ["ID PERSONAL","APELLIDOS Y NOMBRES","ENT LUN","ENT MAR","ENT MIE","ENT JUE","ENT VIE","ENT SAB","ENT TARDE","OBS"]
hw = [12,38,10,10,10,10,10,10,10,24]
for i,h in enumerate(hh,1):
    ws.column_dimensions[get_column_letter(i)].width=hw[i-1]
    cel=ws.cell(2,i,h); cel.font=font(10,True,BLANCO); cel.fill=fill(GRIS_HEAD); cel.alignment=center; cel.border=border_all
for rr in range(3,63):
    for i in range(1,11):
        ws.cell(rr,i).border=border_all; ws.cell(rr,i).font=font(10)
        if 3<=i<=9: ws.cell(rr,i).number_format="HH:MM:SS"
add_name("rng_Horarios", SH_HOR, "$A$3:$J$62")

# ===========================================================================
#  HOJA CALENDARIO (feriados / descansos / eventos)
# ===========================================================================
ws = wb.create_sheet(SH_CAL)
ws.sheet_view.showGridLines = False
ws.merge_cells("A1:D1")
c=ws["A1"]; c.value="CALENDARIO (feriados, descansos, eventos)"
c.font=font(12,True,BLANCO); c.fill=fill(AZUL); c.alignment=left
for i,h in enumerate(["FECHA","DIA","TIPO","DESCRIPCION"],1):
    ws.column_dimensions[get_column_letter(i)].width=[14,14,16,40][i-1]
    cel=ws.cell(2,i,h); cel.font=font(10,True,BLANCO); cel.fill=fill(GRIS_HEAD); cel.alignment=center; cel.border=border_all
feriados_2026 = [
    (dt.date(2026,1,1),"Año Nuevo"),(dt.date(2026,4,2),"Jueves Santo"),
    (dt.date(2026,4,3),"Viernes Santo"),(dt.date(2026,5,1),"Dia del Trabajo"),
    (dt.date(2026,6,29),"San Pedro y San Pablo"),(dt.date(2026,7,23),"Dia FF.AA."),
    (dt.date(2026,7,28),"Fiestas Patrias"),(dt.date(2026,7,29),"Fiestas Patrias"),
    (dt.date(2026,8,6),"Batalla de Junin"),(dt.date(2026,8,30),"Santa Rosa de Lima"),
    (dt.date(2026,10,8),"Combate de Angamos"),(dt.date(2026,11,1),"Todos los Santos"),
    (dt.date(2026,12,8),"Inmaculada Concepcion"),(dt.date(2026,12,9),"Batalla de Ayacucho"),
    (dt.date(2026,12,25),"Navidad"),
]
rr=3
for f,desc in feriados_2026:
    ws.cell(rr,1,f).number_format="DD/MM/YYYY"; ws.cell(rr,1).border=border_all
    ws.cell(rr,2,'=TEXT(A{},"dddd")'.format(rr)).border=border_all
    ws.cell(rr,3,"FERIADO").border=border_all
    ws.cell(rr,4,desc).border=border_all
    for i in range(1,5): ws.cell(rr,i).font=font(10); ws.cell(rr,i).alignment=left
    rr+=1
for extra in range(60):
    for i in range(1,5):
        ws.cell(rr,i).border=border_all; ws.cell(rr,i).font=font(10)
    ws.cell(rr,1).number_format="DD/MM/YYYY"
    rr+=1
CAL_MAXROW=rr-1
dv_tipo=DataValidation(type="list", formula1='"LABORABLE,DESCANSO,FERIADO,EVENTO ESPECIAL"', allow_blank=True)
ws.add_data_validation(dv_tipo); dv_tipo.add("C3:C{}".format(CAL_MAXROW))
add_name("rng_Calendario", SH_CAL, "$A$3:$D${}".format(CAL_MAXROW))

# ===========================================================================
#  HOJA DATA_HUELLERO (RAW - base de datos historica, append-only)
# ===========================================================================
ws = wb.create_sheet(SH_DATA)
dh = ["ID HUELLERO","NOMBRE ORIGINAL","FECHA Y HORA","ESTADO","DISPOSITIVO",
      "TIPO REGISTRO","FECHA","HORA","CLAVE UNICA","LOTE IMPORTACION",
      "FECHA IMPORTACION","USUARIO IMPORTO","OBSERVACION TECNICA"]
dw = [12,36,20,12,12,12,12,12,42,16,18,26,28]
for i,h in enumerate(dh,1):
    ws.column_dimensions[get_column_letter(i)].width=dw[i-1]
    cel=ws.cell(1,i,h); cel.font=font(10,True,BLANCO); cel.fill=fill(GRIS_HEAD); cel.alignment=center; cel.border=border_all
ws.freeze_panes="A2"
add_name("rng_DataHead", SH_DATA, "$A$1:$M$1")

# ===========================================================================
#  HOJA PROCESO (capa de procesamiento: 1 fila por trabajador x dia)
# ===========================================================================
ws = wb.create_sheet(SH_PROC)
pp = ["FECHA","ID PERSONAL","DNI","APELLIDOS Y NOMBRES","ENT MAÑANA","SAL ALMUERZO",
      "ENT TARDE","SAL FINAL","MIN TARDANZA M","MIN TARDANZA T","CONDICION AUTO",
      "CONDICION FINAL","CODIGO CUADRO","OBSERVACION","ORIGEN","MANUAL?","USUARIO",
      "FECHA/HORA PROC","CLAVE"]
pw=[12,12,12,36,12,12,12,12,13,13,18,18,10,26,12,8,24,18,36]
for i,h in enumerate(pp,1):
    ws.column_dimensions[get_column_letter(i)].width=pw[i-1]
    cel=ws.cell(1,i,h); cel.font=font(10,True,BLANCO); cel.fill=fill(GRIS_HEAD); cel.alignment=center; cel.border=border_all
ws.freeze_panes="A2"
add_name("rng_ProcHead", SH_PROC, "$A$1:$S$1")

# ===========================================================================
#  HOJA REPORTE DIARIO (plantilla oficial reconstruida y corregida)
#  Se conserva la apariencia del reporte adjunto; la macro la rellena por VBA.
# ===========================================================================
ws = wb.create_sheet(SH_RDIA)
ws.sheet_view.showGridLines = False
ws.print_area = "A1:I45"
ws.page_setup.orientation = "portrait"
ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
widths={"A":4,"B":6,"C":40,"D":13,"E":13,"F":12,"G":16,"H":18,"I":28}
for k,v in widths.items(): ws.column_dimensions[k].width=v
def mergeset(ws,rng,val,fnt,al,fl=None,bd=None,nf=None):
    ws.merge_cells(rng); c=ws[rng.split(":")[0]]; c.value=val; c.font=fnt; c.alignment=al
    if fl: c.fill=fl
    if nf: c.number_format=nf
    if bd:
        # aplica borde a todo el rango
        from openpyxl.utils import range_boundaries
        x1,y1,x2,y2=range_boundaries(rng)
        for row in ws.iter_rows(min_row=y1,max_row=y2,min_col=x1,max_col=x2):
            for cc in row: cc.border=bd
    return c
mergeset(ws,"B2:I2","=cfg_Empresa",font(16,True,AZUL),center)
mergeset(ws,"B3:I3",'="RUC: "&cfg_RUC',font(10,True),center)
mergeset(ws,"B4:I4",'=cfg_Direccion&"  -  Cel. "&cfg_Celulares',font(10),center)
mergeset(ws,"B5:I5",'="Correo: "&cfg_Correo',font(10),center)
mergeset(ws,"B7:I7","REPORTE DE ASISTENCIA",font(14,True,BLANCO),center,fill(AZUL))
ws.row_dimensions[7].height=24
ws["B9"]="FECHA DEL REPORTE:"; ws["B9"].font=font(10,True)
mergeset(ws,"C9:D9",'=TEXT(cfg_FechaTrabajo,"DD/MM/YYYY")',font(10),left)
ws["E9"]="AREA / SEDE:"; ws["E9"].font=font(10,True)
mergeset(ws,"F9:I9","=cfg_AreaSede",font(10),left)
ws["B10"]="HORA DE INGRESO:"; ws["B10"].font=font(10,True)
mergeset(ws,"C10:D10",'=TEXT(cfg_EntLun,"HH:MM")&"  (TOL: "&cfg_TolMin&" MIN)"',font(10),left)
ws["E10"]="RESPONSABLE:"; ws["E10"].font=font(10,True)
mergeset(ws,"F10:I10","=cfg_Responsable",font(10),left)
# cabecera tabla
th=["N°","APELLIDOS Y NOMBRES","HORA\nENTRADA","HORA\nSALIDA","MINUTOS\nTARDANZA","CONDICION","FIRMA","OBSERVACIONES"]
for i,h in enumerate(th):
    col=i+2
    cel=ws.cell(12,col,h); cel.font=font(10,True,BLANCO); cel.fill=fill(GRIS_HEAD)
    cel.alignment=center; cel.border=border_all
ws.row_dimensions[12].height=30
# 40 filas de detalle (VBA rellena). Bordes + formato hora.
FILA_INI_RDIA=13
for rr in range(FILA_INI_RDIA, FILA_INI_RDIA+45):
    for col in range(2,10):
        cel=ws.cell(rr,col); cel.border=border_all; cel.font=font(10)
        cel.alignment=center if col in (2,4,5,6,7) else left
        if col in (5,6): cel.number_format="HH:MM:SS"
FILA_FIN_RDIA=FILA_INI_RDIA+44
add_name("rng_RepDiarioIni", SH_RDIA, "$B${}".format(FILA_INI_RDIA))
# bloque resumen al pie
rsum=FILA_FIN_RDIA+2
mergeset(ws,"B{r}:C{r}".format(r=rsum),"RESUMEN DEL DIA",font(11,True,BLANCO),center,fill(AZUL))
mergeset(ws,"E{r}:I{r}".format(r=rsum),"OBSERVACIONES GENERALES",font(11,True,BLANCO),center,fill(AZUL))
labels=[("TOTAL PERSONAL:","dash2_total"),("PUNTUALES:","dash2_punt"),("TARDANZAS:","dash2_tard"),
        ("FALTAS:","dash2_falta"),("PERMISOS:","dash2_perm"),("INCIDENCIAS:","dash2_inc"),
        ("PENDIENTES:","dash2_pend")]
for i,(lab,nm) in enumerate(labels):
    rr=rsum+1+i
    ws.cell(rr,2,lab).font=font(10,True); ws.cell(rr,2).alignment=left
    ws.cell(rr,3).border=border_all
mergeset(ws,"E{a}:I{b}".format(a=rsum+1,b=rsum+7)," ",font(10),leftv,bd=border_all)
# firmas
rf=rsum+9
mergeset(ws,"B{r}:D{r}".format(r=rf),"_____________________________",font(10),center)
mergeset(ws,"F{r}:I{r}".format(r=rf),"_____________________________",font(10),center)
mergeset(ws,"B{r}:D{r}".format(r=rf+1),"=cfg_Responsable",font(9,True),center)
mergeset(ws,"F{r}:I{r}".format(r=rf+1),"=cfg_Gerente",font(9,True),center)
mergeset(ws,"B{r}:D{r}".format(r=rf+2),"RESPONSABLE DEL CONTROL",font(8),center)
mergeset(ws,"F{r}:I{r}".format(r=rf+2),"GERENCIA",font(8),center)
mergeset(ws,"B{r}:I{r}".format(r=rf+4),'="Codigo: "&cfg_CodFormato&"   Version: "&cfg_Version',font(8,it=True,color="808080"),center)

# ===========================================================================
#  HOJA CONSOLIDADO MENSUAL
# ===========================================================================
ws = wb.create_sheet(SH_CONS)
ws.sheet_view.showGridLines=False
ws.print_area="A1:R60"; ws.page_setup.orientation="landscape"
ws.sheet_properties.pageSetUpPr.fitToPage=True; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0
mergeset(ws,"A1:R1","=cfg_Empresa&\"  -  CONSOLIDADO MENSUAL DE ASISTENCIA\"",font(14,True,BLANCO),center,fill(AZUL))
mergeset(ws,"A2:R2",'="PERIODO: "&TEXT(DATE(cfg_PeriodoAnio,cfg_PeriodoMes,1),"MMMM YYYY")&"     RESPONSABLE: "&cfg_Responsable',font(10,True),center)
ch=["N°","DNI","APELLIDOS Y NOMBRES","DIAS PROG.","PUNTUAL","TARDANZA","FALTA","PERMISO",
    "CLASE","NO PUSO HUELLA","DESC. MEDICO","COMISION","FERIADO","DESCANSO","VACACIONES",
    "DIAS EFECT.","OBSERVACIONES","EMPRESA/AREA"]
cw=[4,12,34,9,8,9,7,8,7,10,10,9,8,9,9,9,26,24]
for i,h in enumerate(ch,1):
    ws.column_dimensions[get_column_letter(i)].width=cw[i-1]
    cel=ws.cell(3,i,h); cel.font=font(9,True,BLANCO); cel.fill=fill(GRIS_HEAD); cel.alignment=center; cel.border=border_all
ws.row_dimensions[3].height=34
for rr in range(4,4+len(ROSTER)+6):
    for i in range(1,19):
        ws.cell(rr,i).border=border_all; ws.cell(rr,i).font=font(9)
        ws.cell(rr,i).alignment=center if i not in (3,17,18) else left
add_name("rng_ConsolidadoIni", SH_CONS, "$A$4")

# ===========================================================================
#  HOJA CUADRO_GENERAL (la genera VBA mes a mes; aqui queda una guia)
# ===========================================================================
ws = wb.create_sheet(SH_CUAD)
ws.sheet_view.showGridLines=False
mergeset(ws,"A1:H1","CUADRO GENERAL DE ASISTENCIA (se genera por VBA, 1 hoja por mes)",font(12,True,BLANCO),center,fill(AZUL))
notes = [
 "Esta hoja es la GUIA/semilla. Al pulsar CUADRO GENERAL, la macro crea (o actualiza) una hoja",
 "con el nombre del mes (ej: OCTUBRE 2026) reproduciendo el formato oficial del archivo",
 "'CUADRO GENERAL DE ASISTENCIA': N / DNI / APELLIDOS Y NOMBRES, una columna M y una T por dia,",
 "cabeceras LUN..DOM, DESCANSO combinado los domingos, codigos con color segun CATALOGO_ESTADOS,",
 "bloques por empresa (AUDICONTA / SEG PRO FORCE) y el CUADRO RESUMEN CONTROL DE ASISTENCIA",
 "(PUNTUAL / TARDAN. / FALTA / PERMISO / SUSP. / NO FIRMO) + leyenda de codigos y notas R.D.",
 "",
 "Codigos (CATALOGO_ESTADOS):  2=Asistio/Puntual  T=Tardanza  F=Falta  PE=Permiso  S=Suspension",
 "FE=Feriado  V=Vacaciones  C=Comision  CL=Clase  NH=No puso huella  DM=Desc.medico  NF=No firmo  D=Descanso",
]
for i,t in enumerate(notes, start=3):
    ws.cell(i,1,t).font=font(10 if t else 10, it=True, color="404040"); ws.cell(i,1).alignment=left

# ===========================================================================
#  HOJA INCIDENCIAS (gestion + registro, con menus desplegables)
# ===========================================================================
ws = wb.create_sheet(SH_INC)
ws.sheet_view.showGridLines=False
mergeset(ws,"A1:H1","GESTION DE INCIDENCIAS",font(12,True,BLANCO),center,fill(AZUL))
ih=["FECHA","ID PERSONAL","APELLIDOS Y NOMBRES","TIPO DE INCIDENCIA","OBSERVACION",
    "DOCUMENTO SUSTENTATORIO","USUARIO QUE REGISTRA","FECHA/HORA REGISTRO"]
iw=[13,12,34,20,34,26,26,18]
for i,h in enumerate(ih,1):
    ws.column_dimensions[get_column_letter(i)].width=iw[i-1]
    cel=ws.cell(2,i,h); cel.font=font(10,True,BLANCO); cel.fill=fill(GRIS_HEAD); cel.alignment=center; cel.border=border_all
for rr in range(3,203):
    for i in range(1,9):
        ws.cell(rr,i).border=border_all; ws.cell(rr,i).font=font(10)
        ws.cell(rr,i).alignment=left
    ws.cell(rr,1).number_format="DD/MM/YYYY"
    ws.cell(rr,8).number_format="DD/MM/YYYY HH:MM"
dv_inc=DataValidation(type="list", formula1="=rng_Observaciones", allow_blank=True)
ws.add_data_validation(dv_inc); dv_inc.add("D3:D202")
add_name("rng_Incidencias", SH_INC, "$A$3:$H$202")

# ===========================================================================
#  HOJA AUDITORIA
# ===========================================================================
ws = wb.create_sheet(SH_AUD)
ah=["FECHA/HORA","USUARIO","TRABAJADOR","FECHA ASISTENCIA","CAMPO MODIFICADO",
    "VALOR ANTERIOR","VALOR NUEVO","MOTIVO","ORIGEN"]
aw=[18,26,34,15,22,24,24,34,14]
for i,h in enumerate(ah,1):
    ws.column_dimensions[get_column_letter(i)].width=aw[i-1]
    cel=ws.cell(1,i,h); cel.font=font(10,True,BLANCO); cel.fill=fill(GRIS_HEAD); cel.alignment=center; cel.border=border_all
ws.freeze_panes="A2"
add_name("rng_AuditHead", SH_AUD, "$A$1:$I$1")

# ===========================================================================
#  HOJA ALERTAS
# ===========================================================================
ws = wb.create_sheet(SH_ALE)
ws.sheet_view.showGridLines=False
mergeset(ws,"A1:E1","ALERTAS / INCONSISTENCIAS (se genera al validar)",font(12,True,BLANCO),center,fill(AZUL))
al=["#","TIPO DE ALERTA","DETALLE","REFERENCIA","FECHA DETECCION"]
aw=[5,30,50,24,18]
for i,h in enumerate(al,1):
    ws.column_dimensions[get_column_letter(i)].width=aw[i-1]
    cel=ws.cell(2,i,h); cel.font=font(10,True,BLANCO); cel.fill=fill(GRIS_HEAD); cel.alignment=center; cel.border=border_all
ws.freeze_panes="A3"
add_name("rng_AlertasHead", SH_ALE, "$A$2:$E$2")

# ===========================================================================
#  HOJA HISTORICO_MENSUAL
# ===========================================================================
ws = wb.create_sheet(SH_HIST)
hh=["PERIODO","ID PERSONAL","DNI","APELLIDOS Y NOMBRES","PUNTUAL","TARDANZA","FALTA",
    "PERMISO","CLASE","NO PUSO HUELLA","DESC. MEDICO","COMISION","VACACIONES","FERIADO",
    "DESCANSO","DIAS TRABAJADOS","OBSERVACIONES","FECHA CIERRE","USUARIO CIERRE"]
hw=[12,12,12,34,8,9,7,8,7,12,11,9,10,8,9,10,26,16,24]
for i,h in enumerate(hh,1):
    ws.column_dimensions[get_column_letter(i)].width=hw[i-1]
    cel=ws.cell(1,i,h); cel.font=font(9,True,BLANCO); cel.fill=fill(GRIS_HEAD); cel.alignment=center; cel.border=border_all
ws.freeze_panes="A2"
add_name("rng_HistHead", SH_HIST, "$A$1:$S$1")

# ===========================================================================
#  HOJA LOG_ERRORES
# ===========================================================================
ws = wb.create_sheet(SH_LOG)
lh=["FECHA","HORA","PROCEDIMIENTO","N° ERROR","DESCRIPCION","USUARIO"]
lw=[14,12,30,10,60,26]
for i,h in enumerate(lh,1):
    ws.column_dimensions[get_column_letter(i)].width=lw[i-1]
    cel=ws.cell(1,i,h); cel.font=font(10,True,BLANCO); cel.fill=fill(GRIS_HEAD); cel.alignment=center; cel.border=border_all
ws.freeze_panes="A2"
add_name("rng_LogHead", SH_LOG, "$A$1:$F$1")

# ===========================================================================
#  REGISTRAR RANGOS CON NOMBRE
# ===========================================================================
for nm, ref in defined:
    try:
        wb.defined_names.add(DefinedName(nm, attr_text=ref))
    except Exception as e:
        print("WARN nombre", nm, e)

# Orden de hojas y hoja activa
order = [SH_MENU,SH_CONFIG,SH_PERS,SH_HOR,SH_CAL,SH_CAT,SH_DATA,SH_PROC,
         SH_RDIA,SH_CONS,SH_CUAD,SH_INC,SH_AUD,SH_ALE,SH_HIST,SH_LOG]
wb._sheets.sort(key=lambda s: order.index(s.title) if s.title in order else 99)
wb.active = 0

import os
os.makedirs("dist", exist_ok=True)
out="dist/SISTEMA_ASISTENCIA_AUDICONTA.xlsx"
wb.save(out)
print("OK ->", out)
print("Hojas:", wb.sheetnames)
print("Nombres definidos:", len(defined))
