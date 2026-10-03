Attribute VB_Name = "modErrores"
'==============================================================================
' modErrores  -  Manejo y registro de errores controlados
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Registra un error en LOG_ERRORES y muestra un mensaje amable.
' Nunca cierra Excel ni borra informacion.
'------------------------------------------------------------------------------
Public Sub RegistrarError(ByVal procedimiento As String, _
                          Optional ByVal numero As Long = 0, _
                          Optional ByVal descripcion As String = "", _
                          Optional ByVal mostrar As Boolean = True)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    TurboOff
    If numero = 0 Then numero = Err.Number
    If Len(descripcion) = 0 Then descripcion = Err.Description
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    If Not ws Is Nothing Then
        r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If r < 2 Then r = 2
        ws.Cells(r, 1).Value = Date
        ws.Cells(r, 2).Value = Time
        ws.Cells(r, 3).Value = procedimiento
        ws.Cells(r, 4).Value = numero
        ws.Cells(r, 5).Value = descripcion
        ws.Cells(r, 6).Value = UsuarioActual()
    End If
    If mostrar Then
        MsgBox "Se produjo un error controlado." & vbCrLf & vbCrLf & _
               "Procedimiento: " & procedimiento & vbCrLf & _
               "Detalle: " & descripcion & vbCrLf & vbCrLf & _
               "La informacion NO se ha perdido. El error quedo registrado en LOG_ERRORES.", _
               vbExclamation, "AUDICONTA - Control de Asistencia"
    End If
    On Error GoTo 0
End Sub
