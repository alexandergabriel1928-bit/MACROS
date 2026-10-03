Attribute VB_Name = "modConsolidado"
'==============================================================================
' modConsolidado  -  Consolidado mensual (base para remuneraciones)
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
Public Sub GENERAR_CONSOLIDADO()
    Dim ws As Worksheet, wp As Worksheet, mes As Integer, anio As Integer
    Dim n As Long, i As Long, ini As Long, dest As Long
    Dim cPunt As Long, cTard As Long, cFalta As Long, cPerm As Long, cClase As Long
    Dim cNH As Long, cDM As Long, cCom As Long, cFer As Long, cDesc As Long, cVac As Long
    Dim diasProg As Long, diasEfect As Long
    On Error GoTo fallo
    mes = PeriodoMes(): anio = PeriodoAnio()
    Set ws = Hoja(SH_CONS)
    Set wp = Hoja(SH_PROC)
    DesprotegerHoja ws
    ini = ThisWorkbook.Names("rng_ConsolidadoIni").RefersToRange.Row

    ' limpiar
    ws.Range(ws.Cells(ini, 1), ws.Cells(ini + 80, 18)).ClearContents
    n = CargarPersonal()
    dest = ini
    Dim num As Long
    For i = 1 To n
        If gpActivo(i) Then
            num = num + 1
            ContarPorMes wp, gpID(i), mes, anio, cPunt, cTard, cFalta, cPerm, cClase, _
                         cNH, cDM, cCom, cFer, cDesc, cVac, diasProg, diasEfect
            ws.Cells(dest, 1).Value = num
            ws.Cells(dest, 2).Value = gpDNI(i): ws.Cells(dest, 2).NumberFormat = "@"
            ws.Cells(dest, 3).Value = gpNombre(i)
            ws.Cells(dest, 4).Value = diasProg
            ws.Cells(dest, 5).Value = cPunt
            ws.Cells(dest, 6).Value = cTard
            ws.Cells(dest, 7).Value = cFalta
            ws.Cells(dest, 8).Value = cPerm
            ws.Cells(dest, 9).Value = cClase
            ws.Cells(dest, 10).Value = cNH
            ws.Cells(dest, 11).Value = cDM
            ws.Cells(dest, 12).Value = cCom
            ws.Cells(dest, 13).Value = cFer
            ws.Cells(dest, 14).Value = cDesc
            ws.Cells(dest, 15).Value = cVac
            ws.Cells(dest, 16).Value = diasEfect
            ws.Cells(dest, 18).Value = gpArea(i)
            dest = dest + 1
        End If
    Next i
    ws.Activate
    MsgBox "Consolidado del periodo " & MesES(mes) & " " & anio & " generado.", _
           vbInformation, "Consolidado"
    Exit Sub
fallo:
    RegistrarError "modConsolidado.GENERAR_CONSOLIDADO", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Cuenta condiciones del mes para un trabajador.
'------------------------------------------------------------------------------
Public Sub ContarPorMes(ByVal wp As Worksheet, ByVal idp As String, _
                        ByVal mes As Integer, ByVal anio As Integer, _
                        ByRef cPunt As Long, ByRef cTard As Long, ByRef cFalta As Long, _
                        ByRef cPerm As Long, ByRef cClase As Long, ByRef cNH As Long, _
                        ByRef cDM As Long, ByRef cCom As Long, ByRef cFer As Long, _
                        ByRef cDesc As Long, ByRef cVac As Long, _
                        ByRef diasProg As Long, ByRef diasEfect As Long)
    Dim ult As Long, r As Long, f, cond As String
    cPunt = 0: cTard = 0: cFalta = 0: cPerm = 0: cClase = 0: cNH = 0
    cDM = 0: cCom = 0: cFer = 0: cDesc = 0: cVac = 0: diasProg = 0: diasEfect = 0
    ult = UltimaFila(wp, QC_CLAVE)
    For r = 2 To ult
        If Trim$(CStr(wp.Cells(r, QC_IDPERS).Value)) = idp Then
            f = wp.Cells(r, QC_FECHA).Value
            If IsDate(f) Then
                If Month(f) = mes And Year(f) = anio Then
                    cond = UCase$(Trim$(CStr(wp.Cells(r, QC_CFINAL).Value)))
                    Select Case cond
                        Case "PUNTUAL": cPunt = cPunt + 1
                        Case "TARDANZA": cTard = cTard + 1
                        Case "FALTA": cFalta = cFalta + 1
                        Case "PERMISO": cPerm = cPerm + 1
                        Case "CLASE": cClase = cClase + 1
                        Case "NO PUSO HUELLA": cNH = cNH + 1
                        Case "DESCANSO MEDICO": cDM = cDM + 1
                        Case "COMISION": cCom = cCom + 1
                        Case "FERIADO": cFer = cFer + 1
                        Case "DESCANSO": cDesc = cDesc + 1
                        Case "VACACIONES": cVac = cVac + 1
                    End Select
                    If cond <> "DESCANSO" And cond <> "FERIADO" Then diasProg = diasProg + 1
                    If cond = "PUNTUAL" Or cond = "TARDANZA" Or cond = "COMISION" Then diasEfect = diasEfect + 1
                End If
            End If
        End If
    Next r
End Sub
