Attribute VB_Name = "modConfiguracion"
'==============================================================================
' modConfiguracion  -  Lectura centralizada de la CONFIGURACION MAESTRA
' Toda la parametrizacion proviene de rangos con nombre cfg_* (NO del codigo).
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Valor de un parametro por su nombre de rango (cfg_*). Error controlado.
'------------------------------------------------------------------------------
Public Function Cfg(ByVal nombre As String) As Variant
    On Error GoTo fallo
    Cfg = ThisWorkbook.Names(nombre).RefersToRange.Value
    Exit Function
fallo:
    Cfg = ""
End Function

Public Function CfgTxt(ByVal nombre As String) As String
    CfgTxt = Trim$(CStr(Cfg(nombre)))
End Function

Public Function CfgNum(ByVal nombre As String) As Double
    Dim v As Variant
    v = Cfg(nombre)
    If IsNumeric(v) Then CfgNum = CDbl(v) Else CfgNum = 0
End Function

'------------------------------------------------------------------------------
' Fecha de trabajo (propone HOY si esta vacia).
'------------------------------------------------------------------------------
Public Function FechaTrabajo() As Date
    Dim v As Variant
    v = Cfg("cfg_FechaTrabajo")
    If IsDate(v) Then FechaTrabajo = CDate(v) Else FechaTrabajo = Date
End Function

Public Sub SetFechaTrabajo(ByVal f As Date)
    ThisWorkbook.Names("cfg_FechaTrabajo").RefersToRange.Value = f
End Sub

Public Function PeriodoMes() As Integer: PeriodoMes = CInt(CfgNum("cfg_PeriodoMes")): End Function
Public Function PeriodoAnio() As Integer: PeriodoAnio = CInt(CfgNum("cfg_PeriodoAnio")): End Function
Public Function EstadoPeriodo() As String: EstadoPeriodo = UCase$(CfgTxt("cfg_EstadoPeriodo")): End Function

'------------------------------------------------------------------------------
' Hora de ingreso oficial de la manana segun el DIA de la semana.
' Devuelve -1 si el dia es DESCANSO.
'------------------------------------------------------------------------------
Public Function EntradaManianaPorDia(ByVal f As Date) As Double
    Dim nm As String, v As Variant
    Select Case Weekday(f, vbMonday)
        Case 1: nm = "cfg_EntLun"
        Case 2: nm = "cfg_EntMar"
        Case 3: nm = "cfg_EntMie"
        Case 4: nm = "cfg_EntJue"
        Case 5: nm = "cfg_EntVie"
        Case 6: nm = "cfg_EntSab"
        Case 7: nm = "cfg_EntDom"
    End Select
    v = Cfg(nm)
    If VarType(v) = vbString Then
        If InStr(UCase$(CStr(v)), "DESCANSO") > 0 Or Len(Trim$(CStr(v))) = 0 Then
            EntradaManianaPorDia = -1: Exit Function
        End If
    End If
    If IsDate(v) Then
        EntradaManianaPorDia = CDate(v) - Int(CDate(v)) ' solo parte horaria
    Else
        EntradaManianaPorDia = -1
    End If
End Function

'------------------------------------------------------------------------------
' Hora de ingreso EFECTIVA (prioridad: horario PERSONALIZADO > GENERAL).
' Busca en HORARIOS por ID PERSONAL la columna del dia; si no, usa el general.
' Devuelve -1 si descanso.
'------------------------------------------------------------------------------
Public Function EntradaEfectiva(ByVal f As Date, ByVal idPersonal As String) As Double
    Dim ws As Worksheet, ult As Long, r As Long, colDia As Long, v As Variant
    EntradaEfectiva = EntradaManianaPorDia(f)   ' por defecto, general
    If Len(idPersonal) = 0 Then Exit Function
    On Error Resume Next
    Set ws = Hoja(SH_HOR)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Select Case Weekday(f, vbMonday)
        Case 1: colDia = 3   ' ENT LUN
        Case 2: colDia = 4
        Case 3: colDia = 5
        Case 4: colDia = 6
        Case 5: colDia = 7
        Case 6: colDia = 8   ' ENT SAB
        Case 7: Exit Function ' domingo: usa general (descanso)
    End Select
    ult = UltimaFila(ws, 1)
    For r = 3 To ult
        If Trim$(CStr(ws.Cells(r, 1).Value)) = idPersonal Then
            v = ws.Cells(r, colDia).Value
            If IsDate(v) Then EntradaEfectiva = CDate(v) - Int(CDate(v))
            Exit For
        End If
    Next r
End Function

Public Function EntradaTardeEfectiva(ByVal idPersonal As String) As Double
    Dim ws As Worksheet, ult As Long, r As Long, v As Variant
    v = Cfg("cfg_EntTarde")
    If IsDate(v) Then EntradaTardeEfectiva = CDate(v) - Int(CDate(v)) Else EntradaTardeEfectiva = TimeSerial(15, 0, 0)
    On Error Resume Next
    Set ws = Hoja(SH_HOR)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    ult = UltimaFila(ws, 1)
    For r = 3 To ult
        If Trim$(CStr(ws.Cells(r, 1).Value)) = idPersonal Then
            v = ws.Cells(r, 9).Value ' ENT TARDE
            If IsDate(v) Then EntradaTardeEfectiva = CDate(v) - Int(CDate(v))
            Exit For
        End If
    Next r
End Function

Public Function ToleranciaManianaSeg() As Long
    ToleranciaManianaSeg = CLng(CfgNum("cfg_TolMin")) * 60&
End Function
Public Function ToleranciaTardeSeg() As Long
    ToleranciaTardeSeg = CLng(CfgNum("cfg_TolTardeMin")) * 60&
End Function

'------------------------------------------------------------------------------
' CALENDARIO: tipo de dia (LABORABLE / DESCANSO / FERIADO / EVENTO ESPECIAL).
'------------------------------------------------------------------------------
Public Function TipoDeDia(ByVal f As Date) As String
    Dim ws As Worksheet, ult As Long, r As Long
    TipoDeDia = ""
    On Error Resume Next
    Set ws = Hoja(SH_CAL)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    ult = UltimaFila(ws, 1)
    For r = 3 To ult
        If IsDate(ws.Cells(r, 1).Value) Then
            If Int(CDate(ws.Cells(r, 1).Value)) = Int(f) Then
                TipoDeDia = UCase$(Trim$(CStr(ws.Cells(r, 3).Value)))
                Exit Function
            End If
        End If
    Next r
End Function

'------------------------------------------------------------------------------
' CATALOGO_ESTADOS: codigo de cuadro a partir de la CONDICION final.
' Mapea nombres de condicion -> codigo corto del cuadro oficial.
'------------------------------------------------------------------------------
Public Function CodigoCuadro(ByVal condicion As String) As String
    Select Case UCase$(Trim$(condicion))
        Case "PUNTUAL": CodigoCuadro = "2"
        Case "TARDANZA": CodigoCuadro = "T"
        Case "FALTA": CodigoCuadro = "F"
        Case "PERMISO": CodigoCuadro = "PE"
        Case "SUSPENSION", "SUSPENSIÓN": CodigoCuadro = "S"
        Case "FERIADO": CodigoCuadro = "FE"
        Case "VACACIONES": CodigoCuadro = "V"
        Case "COMISION", "COMISIÓN": CodigoCuadro = "C"
        Case "CLASE": CodigoCuadro = "CL"
        Case "NO PUSO HUELLA": CodigoCuadro = "NH"
        Case "DESCANSO MEDICO", "DESCANSO MÉDICO": CodigoCuadro = "DM"
        Case "NO FIRMO", "NO FIRMÓ": CodigoCuadro = "NF"
        Case "DESCANSO": CodigoCuadro = "D"
        Case "JUSTIFICADO": CodigoCuadro = "J"
        Case "INCOMPLETO": CodigoCuadro = "INC"
        Case "PENDIENTE DE REVISION", "PENDIENTE DE REVISIÓN": CodigoCuadro = "PR"
        Case Else: CodigoCuadro = "PR"
    End Select
End Function

'------------------------------------------------------------------------------
' Color de fondo (RGB) de un codigo segun CATALOGO_ESTADOS (col D = hex).
'------------------------------------------------------------------------------
Public Function ColorDeCodigo(ByVal codigo As String) As Long
    Dim ws As Worksheet, ult As Long, r As Long, hx As String
    ColorDeCodigo = RGB(255, 255, 255)
    On Error Resume Next
    Set ws = Hoja(SH_CAT)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    ult = UltimaFila(ws, 1)
    For r = 3 To ult
        If UCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = UCase$(Trim$(codigo)) Then
            hx = Trim$(CStr(ws.Cells(r, 4).Value))
            If Len(hx) = 6 Then
                ColorDeCodigo = RGB(CLng("&H" & Mid$(hx, 1, 2)), _
                                    CLng("&H" & Mid$(hx, 3, 2)), _
                                    CLng("&H" & Mid$(hx, 5, 2)))
            End If
            Exit Function
        End If
    Next r
End Function

'------------------------------------------------------------------------------
' Prioridad de una condicion (para resolver incidencias). Mayor = gana.
'------------------------------------------------------------------------------
Public Function PrioridadCondicion(ByVal condicion As String) As Long
    Dim ws As Worksheet, ult As Long, r As Long, cod As String
    cod = CodigoCuadro(condicion)
    PrioridadCondicion = 0
    On Error Resume Next
    Set ws = Hoja(SH_CAT)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    ult = UltimaFila(ws, 1)
    For r = 3 To ult
        If UCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = cod Then
            If IsNumeric(ws.Cells(r, 5).Value) Then PrioridadCondicion = CLng(ws.Cells(r, 5).Value)
            Exit Function
        End If
    Next r
End Function

'------------------------------------------------------------------------------
' Verifica si el periodo (mes/anio) de una fecha esta CERRADO en HISTORICO.
'------------------------------------------------------------------------------
Public Function PeriodoCerrado(ByVal f As Date) As Boolean
    Dim ws As Worksheet, ult As Long, r As Long, per As String
    per = Format(f, "YYYY-MM")
    PeriodoCerrado = False
    On Error Resume Next
    Set ws = Hoja(SH_HIST)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    ult = UltimaFila(ws, 1)
    For r = 2 To ult
        If Trim$(CStr(ws.Cells(r, 1).Value)) = per Then
            PeriodoCerrado = True: Exit Function
        End If
    Next r
End Function
