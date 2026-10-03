Attribute VB_Name = "modImportacion"
'==============================================================================
' modImportacion  -  Importa la exportacion del huellero a DATA_HUELLERO
'  * Append-only: agrega debajo, NUNCA borra ni reemplaza el historico.
'  * Clave unica: evita duplicados (ID + fecha/hora + tipo + dispositivo).
'  * Detecta trabajadores no registrados y ofrece darlos de alta.
'==============================================================================
Option Explicit

' Columnas de DATA_HUELLERO
Private Const DC_IDHUE As Long = 1
Private Const DC_NOMBRE As Long = 2
Private Const DC_FHORA As Long = 3
Private Const DC_ESTADO As Long = 4
Private Const DC_DISPO As Long = 5
Private Const DC_TIPO As Long = 6
Private Const DC_FECHA As Long = 7
Private Const DC_HORA As Long = 8
Private Const DC_CLAVE As Long = 9
Private Const DC_LOTE As Long = 10
Private Const DC_FIMP As Long = 11
Private Const DC_UIMP As Long = 12
Private Const DC_OBS As Long = 13

'------------------------------------------------------------------------------
Public Sub IMPORTAR_HUELLERO()
    Dim rutaArch As String, wbOri As Workbook, wsOri As Worksheet
    Dim wsData As Worksheet, dicClaves As Object, dicNuevos As Object
    Dim colMap As Object
    Dim lote As String, usuario As String
    Dim nTotal As Long, nNuevos As Long, nDup As Long, nNoID As Long, nErr As Long
    Dim r As Long, ultOri As Long, hdrRow As Long, destino As Long
    Dim idHue As String, nombre As String, fhoraTxt As String, estado As String
    Dim dispo As String, tipo As String, fh As Date, clave As String
    Dim msg As String, k As Variant

    On Error GoTo fallo
    usuario = UsuarioActual()

    ' --- Seleccion de archivo ---
    rutaArch = SeleccionarArchivo()
    If Len(rutaArch) = 0 Then Exit Sub

    ' --- Respaldo previo (requisito 39) ---
    CrearRespaldo "import"

    Set wsData = Hoja(SH_DATA)

    ' --- Abrir origen ---
    Application.ScreenUpdating = False
    Set wbOri = Workbooks.Open(Filename:=rutaArch, ReadOnly:=True, Local:=True)
    Set wsOri = wbOri.Worksheets(1)

    ' --- Detectar fila de cabecera y mapa de columnas ---
    Set colMap = DetectarColumnas(wsOri, hdrRow)
    If colMap Is Nothing Then
        wbOri.Close False
        MsgBox "No se reconocio la estructura del archivo. Se esperan columnas: " & _
               "Numero, Nombre, Tiempo, Estado, Dispositivos, Tipo de Registro.", _
               vbExclamation, "Importar"
        Exit Sub
    End If

    ' --- Diccionario de claves existentes (dedup) ---
    Set dicClaves = CreateObject("Scripting.Dictionary")
    Dim ultData As Long
    ultData = UltimaFila(wsData, DC_IDHUE)
    For r = 2 To ultData
        clave = Trim$(CStr(wsData.Cells(r, DC_CLAVE).Value))
        If Len(clave) > 0 Then If Not dicClaves.Exists(clave) Then dicClaves.Add clave, 1
    Next r

    ' --- Cargar nomina para deteccion de nuevos ---
    CargarPersonal
    Set dicNuevos = CreateObject("Scripting.Dictionary")

    lote = "L" & Format(Now, "YYYYMMDD-HHMMSS")
    ultOri = wsOri.Cells(wsOri.Rows.Count, colMap("tiempo")).End(xlUp).Row
    If ultOri <= hdrRow Then ultOri = wsOri.Cells(wsOri.Rows.Count, colMap("nombre")).End(xlUp).Row
    destino = ultData + 1
    If destino < 2 Then destino = 2

    TurboOn
    For r = hdrRow + 1 To ultOri
        fhoraTxt = Trim$(CStr(wsOri.Cells(r, colMap("tiempo")).Value))
        nombre = Trim$(CStr(wsOri.Cells(r, colMap("nombre")).Value))
        If Len(fhoraTxt) = 0 And Len(nombre) = 0 Then GoTo siguiente
        nTotal = nTotal + 1
        idHue = Trim$(CStr(wsOri.Cells(r, colMap("numero")).Value))
        estado = Trim$(CStr(wsOri.Cells(r, colMap("estado")).Value))
        dispo = Trim$(CStr(wsOri.Cells(r, colMap("dispo")).Value))
        tipo = Trim$(CStr(wsOri.Cells(r, colMap("tipo")).Value))

        fh = ParsearFechaHora(wsOri.Cells(r, colMap("tiempo")).Value)
        If fh = 0 Then fh = ParsearFechaHora(fhoraTxt)

        clave = idHue & "|" & Format(fh, "YYYY-MM-DD HH:MM:SS") & "|" & UCase$(estado) & "|" & UCase$(dispo)
        If dicClaves.Exists(clave) Then
            nDup = nDup + 1
            GoTo siguiente
        End If
        dicClaves.Add clave, 1

        ' escribir RAW (sin modificar lo original de la exportacion)
        wsData.Cells(destino, DC_IDHUE).Value = idHue
        wsData.Cells(destino, DC_IDHUE).NumberFormat = "@"
        wsData.Cells(destino, DC_NOMBRE).Value = nombre
        wsData.Cells(destino, DC_FHORA).Value = fh
        wsData.Cells(destino, DC_FHORA).NumberFormat = "DD/MM/YYYY HH:MM:SS"
        wsData.Cells(destino, DC_ESTADO).Value = estado
        wsData.Cells(destino, DC_DISPO).Value = dispo
        wsData.Cells(destino, DC_TIPO).Value = tipo
        wsData.Cells(destino, DC_FECHA).Value = Int(fh)
        wsData.Cells(destino, DC_FECHA).NumberFormat = "DD/MM/YYYY"
        wsData.Cells(destino, DC_HORA).Value = fh - Int(fh)
        wsData.Cells(destino, DC_HORA).NumberFormat = "HH:MM:SS"
        wsData.Cells(destino, DC_CLAVE).Value = clave
        wsData.Cells(destino, DC_LOTE).Value = lote
        wsData.Cells(destino, DC_FIMP).Value = Now
        wsData.Cells(destino, DC_FIMP).NumberFormat = "DD/MM/YYYY HH:MM"
        wsData.Cells(destino, DC_UIMP).Value = usuario
        If Len(estado) = 0 Then wsData.Cells(destino, DC_OBS).Value = "ESTADO VACIO (se infiere por horario)"
        destino = destino + 1
        nNuevos = nNuevos + 1

        ' deteccion de trabajador no registrado
        If IndicePorMarcacion(idHue, nombre) = 0 Then
            nNoID = nNoID + 1
            If Not dicNuevos.Exists(idHue & "|" & nombre) Then
                dicNuevos.Add idHue & "|" & nombre, Array(idHue, nombre, fh)
            End If
        End If
siguiente:
    Next r

    wbOri.Close False
    TurboOff
    ProgresoFin

    ' --- Resumen de importacion ---
    msg = "IMPORTACION FINALIZADA" & vbCrLf & String(32, "-") & vbCrLf & _
          "Archivo: " & Mid$(rutaArch, InStrRev(rutaArch, Application.PathSeparator) + 1) & vbCrLf & _
          "Registros leidos:        " & nTotal & vbCrLf & _
          "Nuevos registros:        " & nNuevos & vbCrLf & _
          "Duplicados omitidos:     " & nDup & vbCrLf & _
          "Marcaciones sin ID valido: " & nNoID & vbCrLf & _
          "Lote: " & lote
    MsgBox msg, vbInformation, "Importar huellero"

    ' --- Ofrecer alta de trabajadores nuevos ---
    If dicNuevos.Count > 0 Then
        GestionarTrabajadoresNuevos dicNuevos
    End If

    modPrincipal.ActualizarTablero
    Exit Sub
fallo:
    On Error Resume Next
    If Not wbOri Is Nothing Then wbOri.Close False
    TurboOff: ProgresoFin
    RegistrarError "modImportacion.IMPORTAR_HUELLERO", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
Private Function SeleccionarArchivo() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Seleccione la exportacion del huellero (XLS / XLSX / CSV)"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Archivos del huellero", "*.xls; *.xlsx; *.xlsm; *.csv"
        .Filters.Add "Todos", "*.*"
        If .Show = -1 Then SeleccionarArchivo = .SelectedItems(1) Else SeleccionarArchivo = ""
    End With
End Function

'------------------------------------------------------------------------------
' Detecta la fila de cabecera y mapea columnas por nombre (flexible).
' Devuelve un Dictionary con claves: numero, nombre, tiempo, estado, dispo, tipo
'------------------------------------------------------------------------------
Private Function DetectarColumnas(ByVal ws As Worksheet, ByRef hdrRow As Long) As Object
    Dim d As Object, r As Long, c As Long, txt As String, found As Boolean
    Set d = CreateObject("Scripting.Dictionary")
    For r = 1 To 10
        Set d = CreateObject("Scripting.Dictionary")
        For c = 1 To 20
            txt = NORMALIZAR_NOMBRE(CStr(ws.Cells(r, c).Value))
            Select Case True
                Case txt = "NUMERO" Or txt = "N" Or txt = "ID" Or txt = "NO"
                    If Not d.Exists("numero") Then d.Add "numero", c
                Case txt = "NOMBRE" Or txt = "NOMBRES"
                    d("nombre") = c
                Case txt = "TIEMPO" Or txt = "FECHA Y HORA" Or txt = "FECHA/HORA" Or txt = "HORA"
                    If Not d.Exists("tiempo") Then d.Add "tiempo", c
                Case txt = "ESTADO"
                    d("estado") = c
                Case InStr(txt, "DISPOSITIV") > 0
                    d("dispo") = c
                Case InStr(txt, "TIPO") > 0 And InStr(txt, "REGISTRO") > 0
                    d("tipo") = c
            End Select
        Next c
        If d.Exists("nombre") And d.Exists("tiempo") Then
            hdrRow = r
            found = True
            Exit For
        End If
    Next r
    If Not found Then Set DetectarColumnas = Nothing: Exit Function
    If Not d.Exists("numero") Then d.Add "numero", 1
    If Not d.Exists("estado") Then d.Add "estado", d("tiempo") + 1
    If Not d.Exists("dispo") Then d.Add "dispo", d("tiempo") + 2
    If Not d.Exists("tipo") Then d.Add "tipo", d("tiempo") + 3
    Set DetectarColumnas = d
End Function

'------------------------------------------------------------------------------
' Preguntar por cada trabajador nuevo: AGREGAR / IGNORAR / REVISAR (cancelar).
'------------------------------------------------------------------------------
Private Sub GestionarTrabajadoresNuevos(ByVal dic As Object)
    Dim k As Variant, info As Variant, resp As VbMsgBoxResult, msg As String
    Dim idHue As String, nombre As String, idp As String, agregados As Long
    For Each k In dic.Keys
        info = dic(k)
        idHue = CStr(info(0)): nombre = CStr(info(1))
        msg = "SE HA DETECTADO UN TRABAJADOR NO REGISTRADO" & vbCrLf & String(36, "-") & vbCrLf & _
              "ID huellero: " & idHue & vbCrLf & _
              "Nombre:      " & nombre & vbCrLf & _
              "Fecha/hora:  " & Format(info(2), "DD/MM/YYYY HH:MM:SS") & vbCrLf & vbCrLf & _
              "SI = Agregar a PERSONAL    NO = Ignorar    CANCELAR = Revisar luego"
        resp = MsgBox(msg, vbYesNoCancel + vbQuestion, "Personal nuevo detectado")
        If resp = vbCancel Then Exit Sub
        If resp = vbYes Then
            idp = AgregarPersonal(nombre, "", nombre, idHue, "", "")
            agregados = agregados + 1
        End If
    Next k
    If agregados > 0 Then
        MsgBox agregados & " trabajador(es) agregado(s) a PERSONAL. " & _
               "Complete DNI/cargo cuando pueda.", vbInformation, "Personal"
    End If
End Sub
