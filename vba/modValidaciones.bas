Attribute VB_Name = "modValidaciones"
'==============================================================================
' modValidaciones  -  Validaciones previas y deteccion de inconsistencias
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Validaciones antes de procesar. Si detenerSiCritico=True, devuelve False al
' primer problema critico. Siempre informa por pantalla lo encontrado.
'------------------------------------------------------------------------------
Public Function ValidarAntesDeProcesar(ByVal fecha As Date, _
                                       Optional ByVal detenerSiCritico As Boolean = True) As Boolean
    Dim msg As String, critico As Boolean
    ValidarAntesDeProcesar = True
    If ContarPersonalActivo() = 0 Then
        msg = msg & "- No hay personal ACTIVO en la nomina." & vbCrLf: critico = True
    End If
    If UltimaFila(Hoja(SH_DATA), 1) < 2 Then
        msg = msg & "- No hay marcaciones en DATA_HUELLERO (importe primero)." & vbCrLf
    End If
    If PeriodoCerrado(fecha) Then
        msg = msg & "- El periodo " & Format(fecha, "MMMM YYYY") & " esta CERRADO." & vbCrLf: critico = True
    End If
    If Len(msg) > 0 Then
        MsgBox "Validacion:" & vbCrLf & msg, IIf(critico, vbExclamation, vbInformation), "Validacion"
        If critico And detenerSiCritico Then ValidarAntesDeProcesar = False
    End If
End Function

'------------------------------------------------------------------------------
' GENERAR_ALERTAS: detecta inconsistencias y las lista en la hoja ALERTAS.
'------------------------------------------------------------------------------
Public Sub GENERAR_ALERTAS()
    Dim ws As Worksheet, wp As Worksheet, wd As Worksheet
    Dim r As Long, ult As Long, n As Long, i As Long, k As Long
    Dim dicDNI As Object, dicHue As Object, dni As String, idh As String
    Dim fecha As Date
    On Error GoTo fallo
    Set ws = Hoja(SH_ALE)
    DesprotegerHoja ws
    ws.Range("A3:E10000").ClearContents
    k = 0
    fecha = FechaTrabajo()

    ' DNI / ID huellero duplicados y trabajador no identificado
    n = CargarPersonal()
    Set dicDNI = CreateObject("Scripting.Dictionary")
    Set dicHue = CreateObject("Scripting.Dictionary")
    For i = 1 To n
        dni = gpDNI(i): idh = gpIDHue(i)
        If Len(dni) > 0 Then
            If dicDNI.Exists(dni) Then k = AddAlerta(ws, k, "DNI DUPLICADO", "DNI " & dni & " repetido", gpNombre(i))
            dicDNI(dni) = 1
        End If
        If Len(idh) > 0 Then
            If dicHue.Exists(idh) Then k = AddAlerta(ws, k, "ID HUELLERO DUPLICADO", "ID " & idh & " repetido", gpNombre(i))
            dicHue(idh) = 1
        End If
    Next i

    ' marcaciones no identificadas en DATA del dia
    Set wd = Hoja(SH_DATA)
    ult = UltimaFila(wd, 1)
    For r = 2 To ult
        If IsDate(wd.Cells(r, 3).Value) Then
            If Int(CDate(wd.Cells(r, 3).Value)) = Int(fecha) Then
                If IndicePorMarcacion(Trim$(CStr(wd.Cells(r, 1).Value)), _
                                      Trim$(CStr(wd.Cells(r, 2).Value))) = 0 Then
                    k = AddAlerta(ws, k, "TRABAJADOR NO REGISTRADO", _
                        "ID " & wd.Cells(r, 1).Value & " - " & wd.Cells(r, 2).Value, "DATA fila " & r)
                End If
            End If
        End If
    Next r

    ' PROCESO: entrada sin salida / salida sin entrada / pendientes / duplicidad
    Set wp = Hoja(SH_PROC)
    ult = UltimaFila(wp, QC_CLAVE)
    For r = 2 To ult
        If IsDate(wp.Cells(r, QC_FECHA).Value) Then
            If Int(CDate(wp.Cells(r, QC_FECHA).Value)) = Int(fecha) Then
                If IsDate(wp.Cells(r, QC_ENTM).Value) And Not IsDate(wp.Cells(r, QC_SALF).Value) _
                   And Not IsDate(wp.Cells(r, QC_SALALM).Value) Then
                    k = AddAlerta(ws, k, "ENTRADA SIN SALIDA", wp.Cells(r, QC_NOMBRE).Value, _
                                  Format(fecha, "DD/MM/YYYY"))
                End If
                If UCase$(Trim$(CStr(wp.Cells(r, QC_CFINAL).Value))) = "PENDIENTE DE REVISION" Then
                    k = AddAlerta(ws, k, "PENDIENTE DE JUSTIFICACION", wp.Cells(r, QC_NOMBRE).Value, _
                                  Format(fecha, "DD/MM/YYYY"))
                End If
                If InStr(UCase$(CStr(wp.Cells(r, QC_OBS).Value)), "DUPLICIDAD") > 0 Then
                    k = AddAlerta(ws, k, "DOBLE MARCACION", wp.Cells(r, QC_NOMBRE).Value, _
                                  Format(fecha, "DD/MM/YYYY"))
                End If
            End If
        End If
    Next r

    If k = 0 Then AddAlerta ws, 0, "SIN ALERTAS", "No se detectaron inconsistencias para " & _
                                                  Format(fecha, "DD/MM/YYYY"), ""
    Hoja(SH_ALE).Visible = xlSheetVisible
    Hoja(SH_ALE).Activate
    MsgBox k & " alerta(s) detectada(s) para el " & Format(fecha, "DD/MM/YYYY") & ".", _
           vbInformation, "Alertas"
    Exit Sub
fallo:
    RegistrarError "modValidaciones.GENERAR_ALERTAS", Err.Number, Err.Description
End Sub

Private Function AddAlerta(ws As Worksheet, ByVal k As Long, ByVal tipo As String, _
                           ByVal detalle As String, ByVal refe As String) As Long
    Dim r As Long
    r = 3 + k
    ws.Cells(r, 1).Value = k + 1
    ws.Cells(r, 2).Value = tipo
    ws.Cells(r, 3).Value = detalle
    ws.Cells(r, 4).Value = refe
    ws.Cells(r, 5).Value = Now: ws.Cells(r, 5).NumberFormat = "DD/MM/YYYY HH:MM"
    AddAlerta = k + 1
End Function
