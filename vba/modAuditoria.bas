Attribute VB_Name = "modAuditoria"
'==============================================================================
' modAuditoria  -  Registro de toda modificacion manual (trazabilidad)
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Agrega una linea a AUDITORIA. Nunca borra el historico.
'------------------------------------------------------------------------------
Public Sub RegistrarAuditoria(ByVal trabajador As String, _
                              ByVal fechaAsistencia As Variant, _
                              ByVal campo As String, _
                              ByVal valorAnterior As String, _
                              ByVal valorNuevo As String, _
                              ByVal motivo As String, _
                              Optional ByVal origen As String = "MANUAL")
    Dim ws As Worksheet, r As Long
    On Error GoTo fallo
    Set ws = Hoja(SH_AUD)
    r = UltimaFila(ws, 1) + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 1).NumberFormat = "DD/MM/YYYY HH:MM:SS"
    ws.Cells(r, 2).Value = UsuarioActual()
    ws.Cells(r, 3).Value = trabajador
    ws.Cells(r, 4).Value = fechaAsistencia
    ws.Cells(r, 5).Value = campo
    ws.Cells(r, 6).Value = valorAnterior
    ws.Cells(r, 7).Value = valorNuevo
    ws.Cells(r, 8).Value = motivo
    ws.Cells(r, 9).Value = origen
    Exit Sub
fallo:
    RegistrarError "modAuditoria.RegistrarAuditoria", Err.Number, Err.Description, False
End Sub
