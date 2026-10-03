Attribute VB_Name = "modReportes"
'==============================================================================
' modReportes  -  REPORTE DIARIO (plantilla oficial) e INFORME mensual
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Rellena la hoja REPORTE DIARIO con la asistencia de una fecha.
'------------------------------------------------------------------------------
Public Sub LlenarReporteDiario(ByVal fecha As Date)
    Dim ws As Worksheet, wp As Worksheet, ini As Long, r As Long, i As Long, n As Long
    Dim fila As Long, dest As Long
    Dim cPunt As Long, cTard As Long, cFalta As Long, cPerm As Long, cInc As Long, cPend As Long, cTotal As Long
    Dim cond As String, cod As String
    On Error GoTo fallo
    Set ws = Hoja(SH_RDIA)
    Set wp = Hoja(SH_PROC)
    DesprotegerHoja ws
    SetFechaTrabajo fecha
    ini = ThisWorkbook.Names("rng_RepDiarioIni").RefersToRange.Row

    ' limpiar area de detalle
    ws.Range(ws.Cells(ini, 2), ws.Cells(ini + 44, 9)).ClearContents
    ws.Range(ws.Cells(ini, 2), ws.Cells(ini + 44, 9)).Interior.Pattern = xlNone

    n = CargarPersonal()
    dest = ini
    For i = 1 To n
        If gpActivo(i) Then
            cTotal = cTotal + 1
            fila = FilaProceso(fecha, gpID(i))
            ws.Cells(dest, 2).Value = cTotal
            ws.Cells(dest, 3).Value = gpNombre(i)
            If fila > 0 Then
                If IsDate(wp.Cells(fila, QC_ENTM).Value) Then _
                    ws.Cells(dest, 4).Value = wp.Cells(fila, QC_ENTM).Value: ws.Cells(dest, 4).NumberFormat = "HH:MM:SS"
                If IsDate(wp.Cells(fila, QC_SALF).Value) Then
                    ws.Cells(dest, 5).Value = wp.Cells(fila, QC_SALF).Value
                ElseIf IsDate(wp.Cells(fila, QC_SALALM).Value) Then
                    ws.Cells(dest, 5).Value = wp.Cells(fila, QC_SALALM).Value
                End If
                ws.Cells(dest, 5).NumberFormat = "HH:MM:SS"
                ws.Cells(dest, 6).Value = wp.Cells(fila, QC_MINTM).Value
                cond = Trim$(CStr(wp.Cells(fila, QC_CFINAL).Value))
                ws.Cells(dest, 7).Value = cond
                ws.Cells(dest, 9).Value = wp.Cells(fila, QC_OBS).Value
                PintarCondicion ws.Cells(dest, 7), cond
                ContarCond cond, cPunt, cTard, cFalta, cPerm, cInc, cPend
            Else
                ws.Cells(dest, 7).Value = "PENDIENTE DE REVISION"
                PintarCondicion ws.Cells(dest, 7), "PENDIENTE DE REVISION"
                cPend = cPend + 1
            End If
            dest = dest + 1
        End If
    Next i

    ' resumen al pie (buscar etiquetas en columna B)
    EscribirResumen ws, "TOTAL PERSONAL:", cTotal
    EscribirResumen ws, "PUNTUALES:", cPunt
    EscribirResumen ws, "TARDANZAS:", cTard
    EscribirResumen ws, "FALTAS:", cFalta
    EscribirResumen ws, "PERMISOS:", cPerm
    EscribirResumen ws, "INCIDENCIAS:", cInc
    EscribirResumen ws, "PENDIENTES:", cPend
    Exit Sub
fallo:
    RegistrarError "modReportes.LlenarReporteDiario", Err.Number, Err.Description, False
End Sub

Private Sub ContarCond(ByVal cond As String, ByRef p As Long, ByRef t As Long, _
                       ByRef f As Long, ByRef pe As Long, ByRef inc As Long, ByRef pend As Long)
    Select Case UCase$(cond)
        Case "PUNTUAL": p = p + 1
        Case "TARDANZA": t = t + 1
        Case "FALTA": f = f + 1
        Case "PERMISO": pe = pe + 1
        Case "PENDIENTE DE REVISION", "INCOMPLETO": pend = pend + 1
        Case "DESCANSO", "FERIADO": ' no cuenta
        Case Else: inc = inc + 1   ' clase, comision, dm, vacaciones, etc.
    End Select
End Sub

Private Sub EscribirResumen(ByVal ws As Worksheet, ByVal etiqueta As String, ByVal valor As Long)
    Dim r As Long, ult As Long
    ult = UltimaFila(ws, 2)
    For r = 1 To ult + 2
        If UCase$(Trim$(CStr(ws.Cells(r, 2).Value))) = UCase$(etiqueta) Then
            ws.Cells(r, 3).Value = valor
            Exit Sub
        End If
    Next r
End Sub

Public Sub PintarCondicion(ByVal celda As Range, ByVal cond As String)
    celda.Interior.Color = ColorDeCodigo(CodigoCuadro(cond))
    celda.Font.Bold = True
End Sub

'------------------------------------------------------------------------------
Public Sub VER_ASISTENCIA_DIA()
    Dim s As String, fecha As Date
    On Error GoTo fallo
    s = InputBox("Fecha a consultar (DD/MM/AAAA):", "Asistencia del dia", _
                 Format(FechaTrabajo(), "DD/MM/YYYY"))
    If Len(s) = 0 Then Exit Sub
    If Not IsDate(s) Then MsgBox "Fecha invalida.", vbExclamation: Exit Sub
    fecha = CDate(s)
    LlenarReporteDiario fecha
    Hoja(SH_RDIA).Activate
    Exit Sub
fallo:
    RegistrarError "modReportes.VER_ASISTENCIA_DIA", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' GENERAR_INFORME: informe mensual para Gerencia (encabezado institucional).
'------------------------------------------------------------------------------
Public Sub GENERAR_INFORME()
    Dim ws As Worksheet, r As Long, mes As Integer, anio As Integer
    Dim tot(1 To 10) As Long, nombres, i As Long
    On Error GoTo fallo
    mes = PeriodoMes(): anio = PeriodoAnio()
    ResumenMensual mes, anio, tot

    Const SHINF As String = "INFORME"
    If ExisteHoja(SHINF) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(SHINF).Delete
        Application.DisplayAlerts = True
    End If
    Set ws = ThisWorkbook.Worksheets.Add(After:=Hoja(SH_CONS))
    ws.Name = SHINF
    ws.Cells.Font.Name = "Calibri"

    With ws
        .Columns("A").ColumnWidth = 3
        .Columns("B").ColumnWidth = 30
        .Columns("C").ColumnWidth = 18
        .Columns("D").ColumnWidth = 18
        .Columns("E").ColumnWidth = 18
        TituloInforme ws, 2, UCase$(CfgTxt("cfg_Empresa"))
        CentroInforme ws, 3, "RUC: " & CfgTxt("cfg_RUC"), 10, True
        CentroInforme ws, 4, CfgTxt("cfg_Direccion"), 10, False
        CentroInforme ws, 5, "Cel: " & CfgTxt("cfg_Celulares") & "   -   " & CfgTxt("cfg_Correo"), 9, False

        .Range("B7:E7").Merge
        .Range("B7").Value = "INFORME MENSUAL DE CONTROL DE ASISTENCIA N° " & Format(mes, "000") & "-" & anio
        .Range("B7").Font.Size = 13: .Range("B7").Font.Bold = True
        .Range("B7").HorizontalAlignment = xlCenter

        r = 9
        SubInforme ws, r, "1. DATOS GENERALES": r = r + 1
        LineaInforme ws, r, "Empresa:", CfgTxt("cfg_Empresa"): r = r + 1
        LineaInforme ws, r, "Periodo:", MesES(mes) & " " & anio: r = r + 1
        LineaInforme ws, r, "Area / Sede:", CfgTxt("cfg_AreaSede"): r = r + 1
        LineaInforme ws, r, "Responsable del control:", CfgTxt("cfg_Responsable"): r = r + 2

        SubInforme ws, r, "2. RESUMEN DE ASISTENCIA": r = r + 1
        LineaInforme ws, r, "Total trabajadores:", CStr(tot(1)): r = r + 1
        LineaInforme ws, r, "Puntuales (marcas):", CStr(tot(2)): r = r + 1
        LineaInforme ws, r, "Tardanzas:", CStr(tot(3)): r = r + 1
        LineaInforme ws, r, "Faltas:", CStr(tot(4)): r = r + 1
        LineaInforme ws, r, "Permisos:", CStr(tot(5)): r = r + 1
        LineaInforme ws, r, "Clases:", CStr(tot(6)): r = r + 1
        LineaInforme ws, r, "No puso huella:", CStr(tot(7)): r = r + 1
        LineaInforme ws, r, "Descanso medico:", CStr(tot(8)): r = r + 1
        LineaInforme ws, r, "Comisiones:", CStr(tot(9)): r = r + 1
        LineaInforme ws, r, "Pendientes de revision:", CStr(tot(10)): r = r + 2

        SubInforme ws, r, "3. INCIDENCIAS RELEVANTES": r = r + 1
        r = VolcarIncidenciasMes(ws, r, mes, anio) + 1

        SubInforme ws, r, "4. OBSERVACIONES": r = r + 1
        LineaLibre ws, r, "Situaciones pendientes de justificacion: " & tot(10) & ". " & _
                   "Revisar la hoja ALERTAS para inconsistencias del periodo.": r = r + 2

        SubInforme ws, r, "5. CONCLUSION": r = r + 1
        LineaLibre ws, r, "El comportamiento de asistencia del periodo " & MesES(mes) & " " & anio & _
                   " se resume en " & tot(2) & " marcas puntuales, " & tot(3) & " tardanzas y " & _
                   tot(4) & " faltas. Los detalles por trabajador constan en el CONSOLIDADO MENSUAL " & _
                   "y en el CUADRO GENERAL del periodo.": r = r + 3

        SubInforme ws, r, "6. FIRMA": r = r + 2
        .Cells(r, "B").Value = "______________________________"
        .Cells(r + 1, "B").Value = CfgTxt("cfg_Responsable")
        .Cells(r + 2, "B").Value = "Responsable del Control de Asistencia"
        .Cells(r + 2, "B").Font.Size = 9
        .Cells(r, "D").Value = "______________________________"
        .Cells(r + 1, "D").Value = CfgTxt("cfg_Gerente")
        .Cells(r + 2, "D").Value = "Gerencia"
        .Cells(r + 2, "D").Font.Size = 9

        .PageSetup.Orientation = xlPortrait
        .PageSetup.FitToPagesWide = 1: .PageSetup.FitToPagesTall = 0
    End With
    ws.Activate
    MsgBox "Informe mensual generado (hoja INFORME).", vbInformation, "Informe"
    Exit Sub
fallo:
    RegistrarError "modReportes.GENERAR_INFORME", Err.Number, Err.Description
End Sub

' --- helpers de formato del informe ---
Private Sub TituloInforme(ws As Worksheet, r As Long, t As String)
    ws.Range("B" & r & ":E" & r).Merge
    With ws.Range("B" & r)
        .Value = t: .Font.Size = 16: .Font.Bold = True
        .Font.Color = RGB(31, 78, 121): .HorizontalAlignment = xlCenter
    End With
End Sub
Private Sub CentroInforme(ws As Worksheet, r As Long, t As String, sz As Integer, b As Boolean)
    ws.Range("B" & r & ":E" & r).Merge
    With ws.Range("B" & r)
        .Value = t: .Font.Size = sz: .Font.Bold = b: .HorizontalAlignment = xlCenter
    End With
End Sub
Private Sub SubInforme(ws As Worksheet, r As Long, t As String)
    ws.Range("B" & r & ":E" & r).Merge
    With ws.Range("B" & r)
        .Value = t: .Font.Bold = True: .Font.Size = 11
        .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(68, 84, 106)
        .HorizontalAlignment = xlLeft
    End With
End Sub
Private Sub LineaInforme(ws As Worksheet, r As Long, etq As String, val As String)
    ws.Cells(r, "B").Value = etq: ws.Cells(r, "B").Font.Bold = True
    ws.Range("C" & r & ":E" & r).Merge
    ws.Cells(r, "C").Value = val
End Sub
Private Sub LineaLibre(ws As Worksheet, r As Long, val As String)
    ws.Range("B" & r & ":E" & r).Merge
    With ws.Cells(r, "B")
        .Value = val: .WrapText = True: .VerticalAlignment = xlTop
    End With
    ws.Rows(r).RowHeight = 48
End Sub

'------------------------------------------------------------------------------
' Totales del mes (por condicion final) -> arreglo tot(1..10)
'------------------------------------------------------------------------------
Public Sub ResumenMensual(ByVal mes As Integer, ByVal anio As Integer, ByRef tot() As Long)
    Dim wp As Worksheet, ult As Long, r As Long, f, cond As String, i As Long
    For i = 1 To 10: tot(i) = 0: Next i
    tot(1) = ContarPersonalActivo()
    Set wp = Hoja(SH_PROC)
    ult = UltimaFila(wp, QC_CLAVE)
    For r = 2 To ult
        f = wp.Cells(r, QC_FECHA).Value
        If IsDate(f) Then
            If Month(f) = mes And Year(f) = anio Then
                cond = UCase$(Trim$(CStr(wp.Cells(r, QC_CFINAL).Value)))
                Select Case cond
                    Case "PUNTUAL": tot(2) = tot(2) + 1
                    Case "TARDANZA": tot(3) = tot(3) + 1
                    Case "FALTA": tot(4) = tot(4) + 1
                    Case "PERMISO": tot(5) = tot(5) + 1
                    Case "CLASE": tot(6) = tot(6) + 1
                    Case "NO PUSO HUELLA": tot(7) = tot(7) + 1
                    Case "DESCANSO MEDICO": tot(8) = tot(8) + 1
                    Case "COMISION": tot(9) = tot(9) + 1
                    Case "PENDIENTE DE REVISION", "INCOMPLETO": tot(10) = tot(10) + 1
                End Select
            End If
        End If
    Next r
End Sub

Public Function ContarPersonalActivo() As Long
    Dim n As Long, i As Long
    n = CargarPersonal()
    For i = 1 To n
        If gpActivo(i) Then ContarPersonalActivo = ContarPersonalActivo + 1
    Next i
End Function

Private Function VolcarIncidenciasMes(ws As Worksheet, ByVal r As Long, ByVal mes As Integer, ByVal anio As Integer) As Long
    Dim wi As Worksheet, ult As Long, k As Long, f, cnt As Long
    Set wi = Hoja(SH_INC)
    ult = UltimaFila(wi, 1)
    For k = 3 To ult
        f = wi.Cells(k, 1).Value
        If IsDate(f) Then
            If Month(f) = mes And Year(f) = anio Then
                ws.Range("B" & r & ":E" & r).Merge
                ws.Cells(r, "B").Value = Format(f, "DD/MM") & " - " & wi.Cells(k, 3).Value & _
                    " - " & wi.Cells(k, 4).Value & ": " & wi.Cells(k, 5).Value
                ws.Cells(r, "B").Font.Size = 9
                r = r + 1: cnt = cnt + 1
            End If
        End If
    Next k
    If cnt = 0 Then
        ws.Range("B" & r & ":E" & r).Merge
        ws.Cells(r, "B").Value = "Sin incidencias registradas en el periodo."
        r = r + 1
    End If
    VolcarIncidenciasMes = r
End Function
