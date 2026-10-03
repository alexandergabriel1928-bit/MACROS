Attribute VB_Name = "modCuadroGeneral"
'==============================================================================
' modCuadroGeneral  -  Genera la hoja del mes reproduciendo el formato oficial
'  del 'CUADRO GENERAL DE ASISTENCIA': N / DNI / APELLIDOS Y NOMBRES, una
'  columna M y una T por dia, cabeceras de dia de semana, DESCANSO combinado,
'  codigos con color (CATALOGO_ESTADOS) y el CUADRO RESUMEN de control.
'==============================================================================
Option Explicit

Private Const CG_COLN As Long = 1     ' N
Private Const CG_COLDNI As Long = 2   ' DNI
Private Const CG_COLNOM As Long = 3   ' APELLIDOS Y NOMBRES
Private Const CG_DIA0 As Long = 4     ' primera columna de dia (M del dia 1)

'------------------------------------------------------------------------------
Public Sub GENERAR_CUADRO_GENERAL()
    Dim mes As Integer, anio As Integer, dias As Integer, d As Integer
    Dim ws As Worksheet, nombreHoja As String
    Dim n As Long, i As Long, fila As Long, num As Long
    Dim mCol As Long, tCol As Long, fecha As Date
    Dim entrada As Double, tipoDia As String, codigo As String, pfila As Long
    Dim wp As Worksheet
    On Error GoTo fallo

    mes = PeriodoMes(): anio = PeriodoAnio()
    dias = Day(DateSerial(anio, mes + 1, 0))
    nombreHoja = MesES(mes) & " " & anio
    Set wp = Hoja(SH_PROC)

    ' recrear hoja del mes
    If ExisteHoja(nombreHoja) Then
        If PeriodoCerrado(DateSerial(anio, mes, 1)) Then
            MsgBox "El periodo esta CERRADO. La hoja '" & nombreHoja & "' no se regenera.", _
                   vbExclamation, "Cuadro general"
            Hoja(nombreHoja).Activate: Exit Sub
        End If
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(nombreHoja).Delete
        Application.DisplayAlerts = True
    End If
    Set ws = ThisWorkbook.Worksheets.Add(After:=Hoja(SH_CUAD))
    ws.Name = nombreHoja

    TurboOn
    ws.Cells.Font.Name = "Calibri": ws.Cells.Font.Size = 9

    ' --- encabezado institucional ---
    ws.Range(ws.Cells(1, 1), ws.Cells(1, CG_DIA0 + dias * 2)).Merge
    ws.Cells(1, 1).Value = UCase$(CfgTxt("cfg_Empresa"))
    ws.Cells(1, 1).Font.Size = 13: ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).HorizontalAlignment = xlCenter
    ws.Cells(1, 1).Interior.Color = RGB(31, 78, 121): ws.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    ws.Range(ws.Cells(2, 1), ws.Cells(2, CG_DIA0 + dias * 2)).Merge
    ws.Cells(2, 1).Value = "RUC: " & CfgTxt("cfg_RUC") & "   -   " & CfgTxt("cfg_Direccion")
    ws.Cells(2, 1).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(3, 1), ws.Cells(3, CG_DIA0 + dias * 2)).Merge
    ws.Cells(3, 1).Value = "REPORTE DEL MES DE " & UCase$(MesES(mes)) & " " & anio
    ws.Cells(3, 1).Font.Bold = True: ws.Cells(3, 1).HorizontalAlignment = xlCenter

    ' --- cabeceras de columnas ---
    Dim hR As Long: hR = 5   ' fila weekday
    ws.Cells(hR + 1, CG_COLN).Value = "N"
    ws.Cells(hR + 1, CG_COLDNI).Value = "DNI"
    ws.Cells(hR + 1, CG_COLNOM).Value = "APELLIDOS Y NOMBRES"
    For d = 1 To dias
        fecha = DateSerial(anio, mes, d)
        mCol = CG_DIA0 + (d - 1) * 2
        tCol = mCol + 1
        ' weekday (fila hR, combinado)
        ws.Range(ws.Cells(hR, mCol), ws.Cells(hR, tCol)).Merge
        ws.Cells(hR, mCol).Value = DiaCortoES(fecha)
        ' numero de dia (fila hR+1, combinado)
        ws.Range(ws.Cells(hR + 1, mCol), ws.Cells(hR + 1, tCol)).Merge
        ws.Cells(hR + 1, mCol).Value = d
        ' M / T (fila hR+2)
        ws.Cells(hR + 2, mCol).Value = "M"
        ws.Cells(hR + 2, tCol).Value = "T"
        ws.Columns(mCol).ColumnWidth = 3.2
        ws.Columns(tCol).ColumnWidth = 3.2
    Next d
    ws.Range(ws.Cells(hR, 1), ws.Cells(hR + 2, CG_DIA0 + dias * 2 - 1)).Font.Bold = True
    ws.Range(ws.Cells(hR, 1), ws.Cells(hR + 2, CG_DIA0 + dias * 2 - 1)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(hR, 1), ws.Cells(hR + 2, CG_DIA0 + dias * 2 - 1)).Interior.Color = RGB(68, 84, 106)
    ws.Range(ws.Cells(hR, 1), ws.Cells(hR + 2, CG_DIA0 + dias * 2 - 1)).Font.Color = RGB(255, 255, 255)
    ws.Columns(CG_COLN).ColumnWidth = 4
    ws.Columns(CG_COLDNI).ColumnWidth = 11
    ws.Columns(CG_COLNOM).ColumnWidth = 34

    ' --- filas de trabajadores ---
    Dim bodyR As Long: bodyR = hR + 3
    n = CargarPersonal()
    num = 0
    For i = 1 To n
        If gpActivo(i) Then
            num = num + 1
            fila = bodyR + num - 1
            ws.Cells(fila, CG_COLN).Value = num
            ws.Cells(fila, CG_COLDNI).Value = gpDNI(i): ws.Cells(fila, CG_COLDNI).NumberFormat = "@"
            ws.Cells(fila, CG_COLNOM).Value = gpNombre(i)
            ws.Cells(fila, CG_COLNOM).HorizontalAlignment = xlLeft
            For d = 1 To dias
                fecha = DateSerial(anio, mes, d)
                mCol = CG_DIA0 + (d - 1) * 2
                tCol = mCol + 1
                entrada = EntradaEfectiva(fecha, gpID(i))
                tipoDia = TipoDeDia(fecha)
                pfila = FilaProceso(fecha, gpID(i))
                codigo = ""
                If pfila > 0 Then codigo = Trim$(CStr(wp.Cells(pfila, QC_CODIGO).Value))

                If (entrada = -1 Or tipoDia = "DESCANSO") And (codigo = "" Or codigo = "D") Then
                    ws.Range(ws.Cells(fila, mCol), ws.Cells(fila, tCol)).Merge
                    ws.Cells(fila, mCol).Value = "D E S C A N S O"
                    ws.Cells(fila, mCol).Font.Size = 6
                    ws.Cells(fila, mCol).HorizontalAlignment = xlCenter
                    ws.Cells(fila, mCol).Interior.Color = ColorDeCodigo("D")
                ElseIf codigo = "" Then
                    ' sin dato: dejar vacio
                Else
                    PonerCodigoDia ws, fila, mCol, tCol, codigo
                End If
            Next d
        End If
    Next i
    Dim lastBody As Long: lastBody = bodyR + num - 1

    ' bordes del cuerpo
    With ws.Range(ws.Cells(hR, 1), ws.Cells(lastBody, CG_DIA0 + dias * 2 - 1)).Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(150, 150, 150)
    End With

    ' --- CUADRO RESUMEN CONTROL DE ASISTENCIA ---
    GenerarResumenCuadro ws, lastBody + 3, mes, anio, n, bodyR

    ' --- leyenda de codigos ---
    GenerarLeyenda ws, lastBody + 3, CG_DIA0 + 6

    ws.Rows(bodyR & ":" & lastBody).RowHeight = 14
    ws.Cells(hR, 1).Select
    ActiveWindow.FreezePanes = False
    TurboOff
    ws.Activate
    MsgBox "Cuadro general '" & nombreHoja & "' generado.", vbInformation, "Cuadro general"
    Exit Sub
fallo:
    TurboOff
    RegistrarError "modCuadroGeneral.GENERAR_CUADRO_GENERAL", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
Private Sub PonerCodigoDia(ws As Worksheet, ByVal fila As Long, ByVal mCol As Long, _
                           ByVal tCol As Long, ByVal codigo As String)
    Dim col As Long
    Select Case codigo
        Case "2"
            ws.Cells(fila, mCol).Value = "2": ws.Cells(fila, tCol).Value = "2"
        Case "T"
            ws.Cells(fila, mCol).Value = "T": ws.Cells(fila, tCol).Value = "2"
        Case Else
            ws.Cells(fila, mCol).Value = codigo: ws.Cells(fila, tCol).Value = codigo
    End Select
    ws.Cells(fila, mCol).Interior.Color = ColorDeCodigo(ws.Cells(fila, mCol).Value)
    ws.Cells(fila, tCol).Interior.Color = ColorDeCodigo(ws.Cells(fila, tCol).Value)
    ws.Cells(fila, mCol).HorizontalAlignment = xlCenter
    ws.Cells(fila, tCol).HorizontalAlignment = xlCenter
End Sub

'------------------------------------------------------------------------------
Private Sub GenerarResumenCuadro(ws As Worksheet, ByVal r0 As Long, ByVal mes As Integer, _
                                 ByVal anio As Integer, ByVal n As Long, ByVal bodyR As Long)
    Dim i As Long, num As Long, r As Long
    Dim cP As Long, cT As Long, cF As Long, cPe As Long, cS As Long, cNF As Long
    ws.Range(ws.Cells(r0, 1), ws.Cells(r0, 9)).Merge
    ws.Cells(r0, 1).Value = "CUADRO RESUMEN CONTROL DE ASISTENCIA - " & UCase$(MesES(mes)) & " " & anio
    ws.Cells(r0, 1).Font.Bold = True: ws.Cells(r0, 1).HorizontalAlignment = xlCenter
    ws.Cells(r0, 1).Interior.Color = RGB(31, 78, 121): ws.Cells(r0, 1).Font.Color = RGB(255, 255, 255)
    r = r0 + 1
    Dim hdr: hdr = Array("N", "DNI", "APELLIDOS Y NOMBRES", "PUNTUAL", "TARDAN.", "FALTA", "PERMISO", "SUSP.", "NO FIRMO")
    For i = 0 To 8
        ws.Cells(r, i + 1).Value = hdr(i)
        ws.Cells(r, i + 1).Font.Bold = True: ws.Cells(r, i + 1).HorizontalAlignment = xlCenter
        ws.Cells(r, i + 1).Interior.Color = RGB(68, 84, 106): ws.Cells(r, i + 1).Font.Color = RGB(255, 255, 255)
    Next i
    r = r + 1
    For i = 1 To n
        If gpActivo(i) Then
            num = num + 1
            ContarResumen gpID(i), mes, anio, cP, cT, cF, cPe, cS, cNF
            ws.Cells(r, 1).Value = num
            ws.Cells(r, 2).Value = gpDNI(i): ws.Cells(r, 2).NumberFormat = "@"
            ws.Cells(r, 3).Value = gpNombre(i): ws.Cells(r, 3).HorizontalAlignment = xlLeft
            ws.Cells(r, 4).Value = cP
            ws.Cells(r, 5).Value = cT
            ws.Cells(r, 6).Value = cF
            ws.Cells(r, 7).Value = cPe
            ws.Cells(r, 8).Value = cS
            ws.Cells(r, 9).Value = cNF
            r = r + 1
        End If
    Next i
    With ws.Range(ws.Cells(r0 + 1, 1), ws.Cells(r - 1, 9)).Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(150, 150, 150)
    End With
End Sub

Private Sub ContarResumen(ByVal idp As String, ByVal mes As Integer, ByVal anio As Integer, _
                          ByRef cP As Long, ByRef cT As Long, ByRef cF As Long, _
                          ByRef cPe As Long, ByRef cS As Long, ByRef cNF As Long)
    Dim wp As Worksheet, ult As Long, r As Long, f, cond As String
    cP = 0: cT = 0: cF = 0: cPe = 0: cS = 0: cNF = 0
    Set wp = Hoja(SH_PROC)
    ult = UltimaFila(wp, QC_CLAVE)
    For r = 2 To ult
        If Trim$(CStr(wp.Cells(r, QC_IDPERS).Value)) = idp Then
            f = wp.Cells(r, QC_FECHA).Value
            If IsDate(f) Then
                If Month(f) = mes And Year(f) = anio Then
                    cond = UCase$(Trim$(CStr(wp.Cells(r, QC_CFINAL).Value)))
                    Select Case cond
                        Case "PUNTUAL": cP = cP + 1
                        Case "TARDANZA": cT = cT + 1
                        Case "FALTA": cF = cF + 1
                        Case "PERMISO": cPe = cPe + 1
                        Case "SUSPENSION": cS = cS + 1
                        Case "NO FIRMO", "NO PUSO HUELLA": cNF = cNF + 1
                    End Select
                End If
            End If
        End If
    Next r
End Sub

Private Sub GenerarLeyenda(ws As Worksheet, ByVal r0 As Long, ByVal c0 As Long)
    Dim leg, i As Long, r As Long
    leg = Array("2 = ASISTIO/PUNTUAL", "T = TARDANZA", "F = FALTA", "PE = PERMISO", _
                "S = SUSPENSION", "FE = FERIADO", "V = VACACIONES", "C = COMISION", _
                "CL = CLASE", "NH = NO PUSO HUELLA", "DM = DESCANSO MEDICO", _
                "NF = NO FIRMO", "D = DESCANSO")
    ws.Cells(r0, c0).Value = "LEYENDA DE CODIGOS": ws.Cells(r0, c0).Font.Bold = True
    For i = 0 To UBound(leg)
        r = r0 + 1 + i
        ws.Range(ws.Cells(r, c0), ws.Cells(r, c0 + 3)).Merge
        ws.Cells(r, c0).Value = leg(i)
        ws.Cells(r, c0).Font.Size = 8: ws.Cells(r, c0).HorizontalAlignment = xlLeft
    Next i
End Sub
