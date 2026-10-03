VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'==============================================================================
' ThisWorkbook  -  Eventos del libro
'  Al abrir: muestra el MENU, oculta hojas tecnicas, actualiza tablero y
'  detecta cambio de periodo. Al cerrar: avisa si hay cambios sin guardar.
'==============================================================================
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    Application.ScreenUpdating = False
    If ExisteHoja(SH_MENU) Then
        modPrincipal.OcultarTecnicas
        modPrincipal.ActualizarTablero
        modPrincipal.DetectarNuevoPeriodo
        ThisWorkbook.Worksheets(SH_MENU).Activate
    End If
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    If Not ThisWorkbook.Saved Then
        Dim r As VbMsgBoxResult
        r = MsgBox("Existen cambios pendientes de guardar. Desea guardar antes de salir?", _
                   vbQuestion + vbYesNoCancel, "Cerrar")
        Select Case r
            Case vbYes: ThisWorkbook.Save
            Case vbCancel: Cancel = True
        End Select
    End If
    On Error GoTo 0
End Sub
