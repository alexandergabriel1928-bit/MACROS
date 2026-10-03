Attribute VB_Name = "modUtilidades"
'==============================================================================
' modUtilidades  -  Funciones de apoyo transversales
' SISTEMA MAESTRO DE CONTROL DE ASISTENCIA - AUDICONTA PERU E.I.R.L.
'==============================================================================
Option Explicit

' --- Nombres de hojas (constantes centralizadas) ---
Public Const SH_MENU   As String = "01_MENU"
Public Const SH_CONFIG As String = "CONFIG"
Public Const SH_PERS   As String = "PERSONAL"
Public Const SH_CAT    As String = "CATALOGO_ESTADOS"
Public Const SH_HOR    As String = "HORARIOS"
Public Const SH_CAL    As String = "CALENDARIO"
Public Const SH_DATA   As String = "DATA_HUELLERO"
Public Const SH_PROC   As String = "PROCESO"
Public Const SH_RDIA   As String = "REPORTE DIARIO"
Public Const SH_CONS   As String = "CONSOLIDADO"
Public Const SH_CUAD   As String = "CUADRO_GENERAL"
Public Const SH_INC    As String = "INCIDENCIAS"
Public Const SH_AUD    As String = "AUDITORIA"
Public Const SH_ALE    As String = "ALERTAS"
Public Const SH_HIST   As String = "HISTORICO_MENSUAL"
Public Const SH_LOG    As String = "LOG_ERRORES"

'------------------------------------------------------------------------------
' Devuelve una hoja por nombre; error controlado si no existe.
'------------------------------------------------------------------------------
Public Function Hoja(ByVal nombre As String) As Worksheet
    On Error Resume Next
    Set Hoja = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0
    If Hoja Is Nothing Then
        Err.Raise vbObjectError + 513, "Hoja", "No existe la hoja '" & nombre & "'."
    End If
End Function

Public Function ExisteHoja(ByVal nombre As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0
    ExisteHoja = Not ws Is Nothing
End Function

'------------------------------------------------------------------------------
' Ultima fila con datos de una columna.
'------------------------------------------------------------------------------
Public Function UltimaFila(ByVal ws As Worksheet, Optional ByVal col As Long = 1) As Long
    UltimaFila = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

'------------------------------------------------------------------------------
' Usuario actual del equipo (para trazabilidad).
'------------------------------------------------------------------------------
Public Function UsuarioActual() As String
    Dim u As String
    u = Trim$(Environ$("USERNAME"))
    If Len(u) = 0 Then u = Application.UserName
    If Len(u) = 0 Then u = "USUARIO"
    UsuarioActual = u
End Function

'------------------------------------------------------------------------------
' NORMALIZAR_NOMBRE: mayusculas, sin tildes, sin puntos, espacios colapsados.
'------------------------------------------------------------------------------
Public Function NORMALIZAR_NOMBRE(ByVal s As String) As String
    Dim t As String
    t = UCase$(Trim$(CStr(s)))
    ' quitar tildes y caracteres especiales comunes
    t = Replace(t, "Á", "A"): t = Replace(t, "É", "E"): t = Replace(t, "Í", "I")
    t = Replace(t, "Ó", "O"): t = Replace(t, "Ú", "U"): t = Replace(t, "Ü", "U")
    t = Replace(t, "À", "A"): t = Replace(t, "È", "E"): t = Replace(t, "Ì", "I")
    t = Replace(t, ".", " "): t = Replace(t, ",", " "): t = Replace(t, "-", " ")
    t = Replace(t, vbTab, " ")
    ' colapsar espacios dobles
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    NORMALIZAR_NOMBRE = Trim$(t)
End Function

'------------------------------------------------------------------------------
' Clave por conjunto de tokens ordenados (detecta nombres invertidos).
' "PEÑA JUSCAMAYTA AARON" y "AARON PEÑA JUSCAMAYTA" -> misma clave.
'------------------------------------------------------------------------------
Public Function ClaveTokens(ByVal s As String) As String
    Dim partes() As String, i As Long, j As Long, tmp As String
    s = NORMALIZAR_NOMBRE(s)
    If Len(s) = 0 Then ClaveTokens = "": Exit Function
    partes = Split(s, " ")
    ' bubble sort simple (nombres cortos)
    For i = LBound(partes) To UBound(partes) - 1
        For j = i + 1 To UBound(partes)
            If partes(j) < partes(i) Then
                tmp = partes(i): partes(i) = partes(j): partes(j) = tmp
            End If
        Next j
    Next i
    ClaveTokens = Join(partes, " ")
End Function

'------------------------------------------------------------------------------
' Segundos del dia de un valor fecha/hora (parte horaria).
' Comparacion EXACTA al segundo (sin ROUND) para decidir puntualidad.
'------------------------------------------------------------------------------
Public Function SegundosDelDia(ByVal v As Date) As Long
    SegundosDelDia = Hour(v) * 3600& + Minute(v) * 60& + Second(v)
End Function

'------------------------------------------------------------------------------
' Convierte texto de marcacion del huellero a Date robustamente.
' Acepta "1/10/2026 08:04:39", "01/10/2026 8:04", valores Date nativos, etc.
'------------------------------------------------------------------------------
Public Function ParsearFechaHora(ByVal v As Variant) As Date
    Dim s As String, p() As String, fp() As String, hp() As String
    Dim d As Integer, m As Integer, a As Integer
    Dim hh As Integer, mi As Integer, ss As Integer
    On Error GoTo fallo
    If IsDate(v) And Not VarType(v) = vbString Then
        ParsearFechaHora = CDate(v): Exit Function
    End If
    s = Trim$(CStr(v))
    If Len(s) = 0 Then ParsearFechaHora = 0: Exit Function
    p = Split(s, " ")
    fp = Split(p(0), "/")
    If UBound(fp) < 2 Then fp = Split(p(0), "-")
    d = CInt(fp(0)): m = CInt(fp(1)): a = CInt(fp(2))
    If a < 100 Then a = 2000 + a
    hh = 0: mi = 0: ss = 0
    If UBound(p) >= 1 Then
        hp = Split(p(1), ":")
        hh = CInt(hp(0))
        If UBound(hp) >= 1 Then mi = CInt(hp(1))
        If UBound(hp) >= 2 Then ss = CInt(hp(2))
    End If
    ParsearFechaHora = DateSerial(a, m, d) + TimeSerial(hh, mi, ss)
    Exit Function
fallo:
    If IsDate(v) Then ParsearFechaHora = CDate(v) Else ParsearFechaHora = 0
End Function

'------------------------------------------------------------------------------
' Nombre del dia en espanol (para cabeceras del cuadro).
'------------------------------------------------------------------------------
Public Function DiaCortoES(ByVal f As Date) As String
    Select Case Weekday(f, vbMonday)
        Case 1: DiaCortoES = "LUN"
        Case 2: DiaCortoES = "MAR"
        Case 3: DiaCortoES = "MIE"
        Case 4: DiaCortoES = "JUE"
        Case 5: DiaCortoES = "VIE"
        Case 6: DiaCortoES = "SAB"
        Case 7: DiaCortoES = "DOM"
    End Select
End Function

Public Function MesES(ByVal m As Integer) As String
    Dim a
    a = Array("ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", _
              "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE")
    MesES = a(m - 1)
End Function

'------------------------------------------------------------------------------
' Activa optimizaciones de rendimiento / las restaura.
'------------------------------------------------------------------------------
Public Sub TurboOn()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

Public Sub TurboOff()
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'------------------------------------------------------------------------------
' Barra de progreso simple en la barra de estado.
'------------------------------------------------------------------------------
Public Sub Progreso(ByVal texto As String, Optional ByVal i As Long = 0, Optional ByVal n As Long = 0)
    If n > 0 Then
        Application.StatusBar = texto & " " & i & " de " & n & " ..."
    Else
        Application.StatusBar = texto
    End If
End Sub

Public Sub ProgresoFin()
    Application.StatusBar = False
End Sub
