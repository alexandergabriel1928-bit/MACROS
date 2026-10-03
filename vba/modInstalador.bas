Attribute VB_Name = "modInstalador"
'==============================================================================
' modInstalador  -  Instalacion/activacion del sistema (ejecutar UNA vez)
'  Crea los botones del menu, oculta las hojas tecnicas, protege lo necesario
'  y deja el libro listo para usarse por botones (sin tocar VBA).
'==============================================================================
Option Explicit

Private Type TBoton
    etiqueta As String
    macro As String
End Type

'------------------------------------------------------------------------------
Public Sub INSTALAR()
    On Error GoTo fallo
    If MsgBox("Se instalara/actualizara el SISTEMA DE CONTROL DE ASISTENCIA:" & vbCrLf & _
              "- Botones del menu principal" & vbCrLf & _
              "- Ocultar hojas tecnicas" & vbCrLf & _
              "- Proteger CONFIG y PERSONAL (celdas editables libres)" & vbCrLf & vbCrLf & _
              "Continuar?", vbQuestion + vbYesNo, "Instalar") <> vbYes Then Exit Sub

    ConstruirMenu
    BotonesVolver
    modPrincipal.OcultarTecnicas
    ProtegerHojas
    modPrincipal.ActualizarTablero
    Hoja(SH_MENU).Activate
    MsgBox "Sistema instalado correctamente." & vbCrLf & _
           "Use los botones del menu principal.", vbInformation, "Instalar"
    Exit Sub
fallo:
    RegistrarError "modInstalador.INSTALAR", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Modo mantenimiento: muestra todas las hojas y quita protecciones.
'------------------------------------------------------------------------------
Public Sub MODO_MANTENIMIENTO()
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible
        ws.Unprotect
    Next ws
    MsgBox "Modo mantenimiento: todas las hojas visibles y desprotegidas." & vbCrLf & _
           "Ejecute INSTALAR para volver al modo normal.", vbInformation
End Sub

'------------------------------------------------------------------------------
Private Sub ConstruirMenu()
    Dim ws As Worksheet, botones(1 To 14) As TBoton, i As Long
    Dim col As Long, filaTop As Double, x As Double, y As Double
    Dim anchoB As Double, altoB As Double, sep As Double, porCol As Long
    Set ws = Hoja(SH_MENU)

    ' limpiar botones previos
    EliminarShapes ws

    botones(1).etiqueta = "IMPORTAR HUELLERO": botones(1).macro = "IMPORTAR_HUELLERO"
    botones(2).etiqueta = "PROCESAR ASISTENCIA": botones(2).macro = "PROCESAR_ASISTENCIA"
    botones(3).etiqueta = "ASISTENCIA DEL DIA": botones(3).macro = "VER_ASISTENCIA_DIA"
    botones(4).etiqueta = "CUADRO GENERAL": botones(4).macro = "GENERAR_CUADRO_GENERAL"
    botones(5).etiqueta = "CONSOLIDADO MENSUAL": botones(5).macro = "GENERAR_CONSOLIDADO"
    botones(6).etiqueta = "INCIDENCIAS": botones(6).macro = "ABRIR_INCIDENCIAS"
    botones(7).etiqueta = "PERSONAL": botones(7).macro = "ABRIR_PERSONAL"
    botones(8).etiqueta = "CONFIGURACION": botones(8).macro = "ABRIR_CONFIGURACION"
    botones(9).etiqueta = "HISTORICO": botones(9).macro = "ABRIR_HISTORICO"
    botones(10).etiqueta = "AUDITORIA": botones(10).macro = "ABRIR_AUDITORIA"
    botones(11).etiqueta = "ALERTAS / VALIDAR": botones(11).macro = "GENERAR_ALERTAS"
    botones(12).etiqueta = "GENERAR INFORME": botones(12).macro = "GENERAR_INFORME"
    botones(13).etiqueta = "GUARDAR Y RESPALDAR": botones(13).macro = "GUARDAR_Y_RESPALDAR"
    botones(14).etiqueta = "CERRAR MES": botones(14).macro = "CERRAR_MES"

    anchoB = 150: altoB = 38: sep = 10
    filaTop = ws.Range("B6").Top + 180   ' debajo de la tabla de referencia
    porCol = 7
    For i = 1 To 14
        col = (i - 1) \ porCol
        x = ws.Range("B2").Left + col * (anchoB + sep)
        y = filaTop + ((i - 1) Mod porCol) * (altoB + sep)
        CrearBoton ws, botones(i).etiqueta, botones(i).macro, x, y, anchoB, altoB
    Next i
End Sub

Private Sub CrearBoton(ws As Worksheet, ByVal etiqueta As String, ByVal macro As String, _
                       ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)
    Dim sh As Shape
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    sh.Name = "btn_" & macro
    sh.Fill.ForeColor.RGB = RGB(31, 78, 121)
    sh.Line.ForeColor.RGB = RGB(31, 78, 121)
    With sh.TextFrame2.TextRange
        .Text = etiqueta
        .Font.Size = 11
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    sh.TextFrame2.VerticalAnchor = msoAnchorMiddle
    sh.TextFrame2.HorizontalAnchor = msoAnchorCenter
    sh.OnAction = macro
End Sub

'------------------------------------------------------------------------------
' Boton VOLVER AL MENU en hojas de trabajo.
'------------------------------------------------------------------------------
Private Sub BotonesVolver()
    Dim nombres, nm, ws As Worksheet, sh As Shape
    nombres = Array(SH_CONFIG, SH_PERS, SH_RDIA, SH_CONS, SH_INC, SH_CAL, SH_HOR, SH_ALE)
    For Each nm In nombres
        If ExisteHoja(CStr(nm)) Then
            Set ws = Hoja(CStr(nm))
            EliminarShapesVolver ws
            Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, 5, 2, 110, 22)
            sh.Name = "btn_VOLVER"
            sh.Fill.ForeColor.RGB = RGB(68, 84, 106)
            sh.Line.Visible = msoFalse
            sh.TextFrame2.TextRange.Text = "<< MENU"
            sh.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            sh.TextFrame2.TextRange.Font.Bold = msoTrue
            sh.TextFrame2.VerticalAnchor = msoAnchorMiddle
            sh.TextFrame2.HorizontalAnchor = msoAnchorCenter
            sh.OnAction = "VolverAlMenu"
        End If
    Next nm
End Sub

Private Sub EliminarShapes(ws As Worksheet)
    Dim sh As Shape, i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 4) = "btn_" Then ws.Shapes(i).Delete
    Next i
End Sub
Private Sub EliminarShapesVolver(ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).Name = "btn_VOLVER" Then ws.Shapes(i).Delete
    Next i
End Sub

'------------------------------------------------------------------------------
' Proteccion: CONFIG y PERSONAL con celdas editables libres.
'------------------------------------------------------------------------------
Private Sub ProtegerHojas()
    On Error Resume Next
    Dim ws As Worksheet
    ' CONFIG: desbloquear solo columna B (valores)
    Set ws = Hoja(SH_CONFIG)
    ws.Unprotect
    ws.Cells.Locked = True
    ws.Range("B1:B200").Locked = False
    ws.Protect UserInterfaceOnly:=True, AllowFiltering:=True
    ' PERSONAL: dejar editable toda la tabla de datos
    Set ws = Hoja(SH_PERS)
    ws.Unprotect
    On Error GoTo 0
End Sub
