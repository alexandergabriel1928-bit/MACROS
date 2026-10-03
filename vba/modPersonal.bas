Attribute VB_Name = "modPersonal"
'==============================================================================
' modPersonal  -  Nomina maestra: carga, busqueda y altas controladas
' La nomina es INDEPENDIENTE del huellero: importar NUNCA la reconstruye.
'==============================================================================
Option Explicit

' Columnas de la hoja PERSONAL
Public Const PC_IDPERS As Long = 1
Public Const PC_N As Long = 2
Public Const PC_DNI As Long = 3
Public Const PC_NOMBRE As Long = 4
Public Const PC_NOMHUE As Long = 5
Public Const PC_IDHUE As Long = 6
Public Const PC_CARGO As Long = 7
Public Const PC_AREA As Long = 8
Public Const PC_FING As Long = 9
Public Const PC_FCESE As Long = 10
Public Const PC_JORNADA As Long = 11
Public Const PC_HORPERS As Long = 12
Public Const PC_ACTIVO As Long = 13
Public Const PC_OBS As Long = 14

' Cache en memoria (poblar con CargarPersonal)
Public gPN As Long                  ' cantidad
Public gpFila() As Long
Public gpID() As String
Public gpDNI() As String
Public gpNombre() As String
Public gpNomHue() As String
Public gpIDHue() As String
Public gpArea() As String
Public gpActivo() As Boolean

'------------------------------------------------------------------------------
' Carga la nomina a memoria. Devuelve la cantidad de trabajadores.
'------------------------------------------------------------------------------
Public Function CargarPersonal() As Long
    Dim ws As Worksheet, ult As Long, r As Long, k As Long
    Set ws = Hoja(SH_PERS)
    ult = UltimaFila(ws, PC_NOMBRE)
    ReDim gpFila(1 To Application.Max(1, ult))
    ReDim gpID(1 To Application.Max(1, ult))
    ReDim gpDNI(1 To Application.Max(1, ult))
    ReDim gpNombre(1 To Application.Max(1, ult))
    ReDim gpNomHue(1 To Application.Max(1, ult))
    ReDim gpIDHue(1 To Application.Max(1, ult))
    ReDim gpArea(1 To Application.Max(1, ult))
    ReDim gpActivo(1 To Application.Max(1, ult))
    k = 0
    For r = 3 To ult
        If Len(Trim$(CStr(ws.Cells(r, PC_NOMBRE).Value))) > 0 Then
            k = k + 1
            gpFila(k) = r
            gpID(k) = Trim$(CStr(ws.Cells(r, PC_IDPERS).Value))
            gpDNI(k) = Trim$(CStr(ws.Cells(r, PC_DNI).Value))
            gpNombre(k) = Trim$(CStr(ws.Cells(r, PC_NOMBRE).Value))
            gpNomHue(k) = Trim$(CStr(ws.Cells(r, PC_NOMHUE).Value))
            gpIDHue(k) = Trim$(CStr(ws.Cells(r, PC_IDHUE).Value))
            gpArea(k) = Trim$(CStr(ws.Cells(r, PC_AREA).Value))
            gpActivo(k) = (UCase$(Trim$(CStr(ws.Cells(r, PC_ACTIVO).Value))) <> "NO")
            If Len(gpID(k)) = 0 Then gpID(k) = "P" & Format(k, "000")
        End If
    Next r
    gPN = k
    CargarPersonal = k
End Function

'------------------------------------------------------------------------------
' Busca el indice (1..gPN) del trabajador que corresponde a una marcacion.
' Prioridad: 1) ID huellero  2) nombre huellero normalizado
'            3) clave de tokens (detecta nombres invertidos)
' Devuelve 0 si no se encuentra.
'------------------------------------------------------------------------------
Public Function IndicePorMarcacion(ByVal idHuellero As String, ByVal nombreOrig As String) As Long
    Dim i As Long, nn As String, nt As String
    idHuellero = Trim$(idHuellero)
    nn = NORMALIZAR_NOMBRE(nombreOrig)
    nt = ClaveTokens(nombreOrig)
    ' 1) por ID huellero
    If Len(idHuellero) > 0 Then
        For i = 1 To gPN
            If Len(gpIDHue(i)) > 0 Then
                If gpIDHue(i) = idHuellero Then IndicePorMarcacion = i: Exit Function
            End If
        Next i
    End If
    ' 2) por nombre huellero normalizado
    For i = 1 To gPN
        If Len(gpNomHue(i)) > 0 Then
            If NORMALIZAR_NOMBRE(gpNomHue(i)) = nn Then IndicePorMarcacion = i: Exit Function
        End If
    Next i
    ' 3) por clave de tokens contra nombre huellero y apellidos/nombres
    For i = 1 To gPN
        If Len(gpNomHue(i)) > 0 Then
            If ClaveTokens(gpNomHue(i)) = nt Then IndicePorMarcacion = i: Exit Function
        End If
    Next i
    For i = 1 To gPN
        If ClaveTokens(gpNombre(i)) = nt Then IndicePorMarcacion = i: Exit Function
    Next i
    IndicePorMarcacion = 0
End Function

'------------------------------------------------------------------------------
' Genera el siguiente ID PERSONAL (Pxxx) sin colisiones.
'------------------------------------------------------------------------------
Public Function SiguienteIDPersonal() As String
    Dim ws As Worksheet, ult As Long, r As Long, mx As Long, s As String, n As Long
    Set ws = Hoja(SH_PERS)
    ult = UltimaFila(ws, PC_NOMBRE)
    mx = 0
    For r = 3 To ult
        s = Trim$(CStr(ws.Cells(r, PC_IDPERS).Value))
        If Len(s) >= 2 And UCase$(Left$(s, 1)) = "P" Then
            If IsNumeric(Mid$(s, 2)) Then
                n = CLng(Mid$(s, 2))
                If n > mx Then mx = n
            End If
        End If
    Next r
    SiguienteIDPersonal = "P" & Format(mx + 1, "000")
End Function

'------------------------------------------------------------------------------
' Agrega un trabajador a la nomina (alta controlada). Devuelve el ID PERSONAL.
'------------------------------------------------------------------------------
Public Function AgregarPersonal(ByVal nombre As String, _
                                Optional ByVal dni As String = "", _
                                Optional ByVal nombreHuellero As String = "", _
                                Optional ByVal idHuellero As String = "", _
                                Optional ByVal cargo As String = "", _
                                Optional ByVal area As String = "") As String
    Dim ws As Worksheet, r As Long, idp As String, n As Long
    Set ws = Hoja(SH_PERS)
    DesprotegerHoja ws
    r = UltimaFila(ws, PC_NOMBRE) + 1
    If r < 3 Then r = 3
    n = 0
    If r > 3 Then If IsNumeric(ws.Cells(r - 1, PC_N).Value) Then n = CLng(ws.Cells(r - 1, PC_N).Value)
    idp = SiguienteIDPersonal()
    ws.Cells(r, PC_IDPERS).Value = idp
    ws.Cells(r, PC_N).Value = n + 1
    ws.Cells(r, PC_DNI).Value = dni:  ws.Cells(r, PC_DNI).NumberFormat = "@"
    ws.Cells(r, PC_NOMBRE).Value = UCase$(Trim$(nombre))
    ws.Cells(r, PC_NOMHUE).Value = nombreHuellero
    ws.Cells(r, PC_IDHUE).Value = idHuellero: ws.Cells(r, PC_IDHUE).NumberFormat = "@"
    ws.Cells(r, PC_CARGO).Value = cargo
    ws.Cells(r, PC_AREA).Value = area
    ws.Cells(r, PC_FING).Value = Date: ws.Cells(r, PC_FING).NumberFormat = "DD/MM/YYYY"
    ws.Cells(r, PC_JORNADA).Value = "ESTANDAR"
    ws.Cells(r, PC_ACTIVO).Value = "SI"
    RegistrarAuditoria UCase$(Trim$(nombre)), "", "ALTA PERSONAL", "", idp, _
                       "Alta de trabajador", "MANUAL"
    AgregarPersonal = idp
End Function

Public Sub DesprotegerHoja(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0
End Sub
