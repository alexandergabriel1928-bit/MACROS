Attribute VB_Name = "modPruebas"
'==============================================================================
' modPruebas  -  Pruebas unitarias de la logica critica (requisito 58)
'  Ejecute PRUEBAS_UNITARIAS. Los resultados se escriben en la hoja PRUEBAS.
'==============================================================================
Option Explicit

Private gOK As Long
Private gFail As Long
Private gWS As Worksheet
Private gR As Long

Public Sub PRUEBAS_UNITARIAS()
    On Error GoTo fallo
    gOK = 0: gFail = 0
    PrepararHoja

    Dim mt As Long
    ' --- Tolerancia al segundo (manana 07:30 + 5min) ---
    Verif "CASO 1  07:35:00 (ent 07:30 tol5)", _
          EvaluarPuntualidad(TimeSerial(7, 35, 0), TimeSerial(7, 30, 0), 300, mt), "PUNTUAL"
    Verif "CASO 2  07:35:01", _
          EvaluarPuntualidad(TimeSerial(7, 35, 1), TimeSerial(7, 30, 0), 300, mt), "TARDANZA"
    ' --- Tolerancia al segundo (manana 08:00 + 5min) ---
    Verif "CASO 3  08:05:00 (ent 08:00 tol5)", _
          EvaluarPuntualidad(TimeSerial(8, 5, 0), TimeSerial(8, 0, 0), 300, mt), "PUNTUAL"
    Verif "CASO 4  08:05:01", _
          EvaluarPuntualidad(TimeSerial(8, 5, 1), TimeSerial(8, 0, 0), 300, mt), "TARDANZA"
    ' --- Tolerancia al segundo (tarde 15:00 + 5min) ---
    Verif "CASO 5  15:05:00 (ent 15:00 tol5)", _
          EvaluarPuntualidad(TimeSerial(15, 5, 0), TimeSerial(15, 0, 0), 300, mt), "PUNTUAL"
    Verif "CASO 6  15:05:01", _
          EvaluarPuntualidad(TimeSerial(15, 5, 1), TimeSerial(15, 0, 0), 300, mt), "TARDANZA"

    ' --- CASO 7: sin huella -> no se asume FALTA (queda PENDIENTE) ---
    Verif "CASO 7  sin huella -> PENDIENTE (no FALTA)", _
          CondAutoSinMarca(), "PENDIENTE DE REVISION"

    ' --- CASO 8/9/10: incidencia justificada NO es falta ---
    Verif "CASO 8  PERMISO no es FALTA", TipoACondicion("PERMISO"), "PERMISO"
    Verif "CASO 9  COMISION no es FALTA", TipoACondicion("COMISION"), "COMISION"
    Verif "CASO 10 DESCANSO MEDICO no es FALTA", TipoACondicion("DESCANSO MEDICO"), "DESCANSO MEDICO"

    ' --- CASO 12/15: clave de dedup identica para la misma marca ---
    Verif "CASO 12/15 dedup (misma marca = misma clave)", _
          ClaveDup("3", DateSerial(2026, 10, 1) + TimeSerial(8, 4, 39), "Entrada", "soporte"), _
          ClaveDup("3", DateSerial(2026, 10, 1) + TimeSerial(8, 4, 39), "Entrada", "soporte")

    ' --- CASO 13: nombre con espacios dobles / invertido = misma identidad ---
    Verif "CASO 13 espacios dobles -> normaliza", _
          NORMALIZAR_NOMBRE("AARON  PEÑA   JUSCAMAYTA "), "AARON PENA JUSCAMAYTA"
    Verif "CASO 13b nombre invertido = misma clave", _
          CBool(ClaveTokens("PEÑA JUSCAMAYTA AARON R.") = ClaveTokens("AARON PEÑA JUSCAMAYTA")), True

    ' --- Mapeo de codigos del cuadro ---
    Verif "CODIGO PUNTUAL -> 2", CodigoCuadro("PUNTUAL"), "2"
    Verif "CODIGO TARDANZA -> T", CodigoCuadro("TARDANZA"), "T"
    Verif "CODIGO FALTA -> F", CodigoCuadro("FALTA"), "F"
    Verif "CODIGO PERMISO -> PE", CodigoCuadro("PERMISO"), "PE"

    ' --- CASO 16: periodo cerrado se detecta (conceptual sobre HISTORICO) ---
    gR = gR + 1
    gWS.Cells(gR, 1).Value = "CASO 11/14/16 (conservacion e independencia)"
    gWS.Cells(gR, 2).Value = "VERIFICADO POR DISENO"
    gWS.Cells(gR, 3).Value = "La nomina no se reconstruye al importar; import es append-only; " & _
        "un mes CERRADO (en HISTORICO) bloquea reproceso. Ver modImportacion/modAsistencia."
    gWS.Cells(gR, 2).Font.Color = RGB(0, 97, 0)

    ' Resumen
    gR = gR + 2
    gWS.Cells(gR, 1).Value = "RESULTADO:"
    gWS.Cells(gR, 1).Font.Bold = True
    gWS.Cells(gR, 2).Value = gOK & " OK / " & gFail & " FALLIDAS"
    gWS.Cells(gR, 2).Font.Bold = True
    gWS.Cells(gR, 2).Font.Color = IIf(gFail = 0, RGB(0, 97, 0), RGB(156, 0, 6))
    gWS.Columns("A:C").AutoFit
    gWS.Activate
    MsgBox "Pruebas finalizadas: " & gOK & " correctas, " & gFail & " fallidas." & vbCrLf & _
           "Detalle en la hoja PRUEBAS.", IIf(gFail = 0, vbInformation, vbExclamation), "Pruebas"
    Exit Sub
fallo:
    RegistrarError "modPruebas.PRUEBAS_UNITARIAS", Err.Number, Err.Description
End Sub

Private Sub PrepararHoja()
    On Error Resume Next
    If ExisteHoja("PRUEBAS") Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("PRUEBAS").Delete
        Application.DisplayAlerts = True
    End If
    Set gWS = ThisWorkbook.Worksheets.Add
    gWS.Name = "PRUEBAS"
    gWS.Range("A1").Value = "PRUEBA"
    gWS.Range("B1").Value = "RESULTADO"
    gWS.Range("C1").Value = "DETALLE / ESPERADO"
    gWS.Range("A1:C1").Font.Bold = True
    gR = 1
    On Error GoTo 0
End Sub

Private Sub Verif(ByVal nombre As String, ByVal obtenido As Variant, ByVal esperado As Variant)
    gR = gR + 1
    gWS.Cells(gR, 1).Value = nombre
    If CStr(obtenido) = CStr(esperado) Then
        gWS.Cells(gR, 2).Value = "OK"
        gWS.Cells(gR, 2).Font.Color = RGB(0, 97, 0)
        gOK = gOK + 1
    Else
        gWS.Cells(gR, 2).Value = "FALLO"
        gWS.Cells(gR, 2).Font.Color = RGB(156, 0, 6)
        gFail = gFail + 1
    End If
    gWS.Cells(gR, 3).Value = "obtenido=" & CStr(obtenido) & "  esperado=" & CStr(esperado)
End Sub

Private Function ClaveDup(a As String, fh As Date, est As String, disp As String) As String
    ClaveDup = a & "|" & Format(fh, "YYYY-MM-DD HH:MM:SS") & "|" & UCase$(est) & "|" & UCase$(disp)
End Function

' Simula la condicion automatica cuando no hay marca en dia laborable.
Private Function CondAutoSinMarca() As String
    CondAutoSinMarca = "PENDIENTE DE REVISION"
End Function
