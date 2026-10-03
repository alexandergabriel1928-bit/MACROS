Attribute VB_Name = "modIncidencias"
'==============================================================================
' modIncidencias  -  Registro y aplicacion de incidencias manuales
'  Una incidencia justificada NUNCA se convierte en FALTA automaticamente.
'  Toda aplicacion queda auditada.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Convierte el TIPO DE INCIDENCIA (menu) en una CONDICION estandar.
'------------------------------------------------------------------------------
Public Function TipoACondicion(ByVal tipo As String) As String
    Select Case UCase$(Trim$(tipo))
        Case "CLASE": TipoACondicion = "CLASE"
        Case "PERMISO": TipoACondicion = "PERMISO"
        Case "NO PUSO HUELLA": TipoACondicion = "NO PUSO HUELLA"
        Case "DESCANSO MEDICO", "DESCANSO MÉDICO": TipoACondicion = "DESCANSO MEDICO"
        Case "COMISION", "COMISIÓN": TipoACondicion = "COMISION"
        Case "DESCANSO": TipoACondicion = "DESCANSO"
        Case "FERIADO": TipoACondicion = "FERIADO"
        Case "VACACIONES": TipoACondicion = "VACACIONES"
        Case "SUSPENSION", "SUSPENSIÓN": TipoACondicion = "SUSPENSION"
        Case "FALTA": TipoACondicion = "FALTA"
        Case "JUSTIFICADO", "OTRO": TipoACondicion = "JUSTIFICADO"
        Case Else: TipoACondicion = "JUSTIFICADO"
    End Select
End Function

'------------------------------------------------------------------------------
' Aplica todas las incidencias registradas para una fecha a la hoja PROCESO.
'------------------------------------------------------------------------------
Public Sub AplicarIncidencias(ByVal fecha As Date)
    Dim ws As Worksheet, ult As Long, r As Long
    Dim fI: Dim idp As String, nombre As String, tipo As String, obs As String
    Dim cond As String, fila As Long, prevCond As String
    On Error GoTo fallo
    Set ws = Hoja(SH_INC)
    ult = UltimaFila(ws, 1)
    For r = 3 To ult
        fI = ws.Cells(r, 1).Value
        If IsDate(fI) Then
            If Int(CDate(fI)) = Int(fecha) Then
                idp = Trim$(CStr(ws.Cells(r, 2).Value))
                nombre = Trim$(CStr(ws.Cells(r, 3).Value))
                tipo = Trim$(CStr(ws.Cells(r, 4).Value))
                obs = Trim$(CStr(ws.Cells(r, 5).Value))
                If Len(tipo) = 0 Then GoTo sig
                If Len(idp) = 0 Then idp = ResolverIDPorNombre(nombre)
                If Len(idp) = 0 Then GoTo sig
                cond = TipoACondicion(tipo)
                fila = AsegurarFilaProceso(fecha, idp)
                If fila = 0 Then GoTo sig
                prevCond = Trim$(CStr(ws.Parent.Worksheets(SH_PROC).Cells(fila, QC_CFINAL).Value))
                AplicarCondicionAFila fila, cond, obs, prevCond, fecha
            End If
        End If
sig:
    Next r
    Exit Sub
fallo:
    RegistrarError "modIncidencias.AplicarIncidencias", Err.Number, Err.Description, False
End Sub

Private Sub AplicarCondicionAFila(ByVal fila As Long, ByVal cond As String, _
                                  ByVal obs As String, ByVal prevCond As String, ByVal fecha As Date)
    Dim wp As Worksheet
    Set wp = Hoja(SH_PROC)
    If UCase$(prevCond) = UCase$(cond) And UCase$(Trim$(CStr(wp.Cells(fila, QC_MANUAL).Value))) = "SI" Then Exit Sub
    RegistrarAuditoria CStr(wp.Cells(fila, QC_NOMBRE).Value), fecha, "CONDICION", _
                       prevCond, cond, IIf(Len(obs) > 0, obs, "Incidencia registrada"), "INCIDENCIA"
    wp.Cells(fila, QC_CFINAL).Value = cond
    wp.Cells(fila, QC_CODIGO).Value = CodigoCuadro(cond)
    If Len(obs) > 0 Then wp.Cells(fila, QC_OBS).Value = obs
    wp.Cells(fila, QC_MANUAL).Value = "SI"
    wp.Cells(fila, QC_ORIGEN).Value = "INCIDENCIA"
    wp.Cells(fila, QC_USER).Value = UsuarioActual()
End Sub

Public Function ResolverIDPorNombre(ByVal nombre As String) As String
    Dim i As Long, nt As String
    nt = ClaveTokens(nombre)
    If gPN = 0 Then CargarPersonal
    For i = 1 To gPN
        If ClaveTokens(gpNombre(i)) = nt Then ResolverIDPorNombre = gpID(i): Exit Function
    Next i
    ResolverIDPorNombre = ""
End Function

'------------------------------------------------------------------------------
' Garantiza que exista una fila en PROCESO para (fecha, idp). La crea si falta.
'------------------------------------------------------------------------------
Public Function AsegurarFilaProceso(ByVal fecha As Date, ByVal idp As String) As Long
    Dim wp As Worksheet, fila As Long, i As Long
    fila = FilaProceso(fecha, idp)
    If fila > 0 Then AsegurarFilaProceso = fila: Exit Function
    Set wp = Hoja(SH_PROC)
    If gPN = 0 Then CargarPersonal
    fila = UltimaFila(wp, QC_CLAVE) + 1
    If fila < 2 Then fila = 2
    wp.Cells(fila, QC_FECHA).Value = fecha: wp.Cells(fila, QC_FECHA).NumberFormat = "DD/MM/YYYY"
    wp.Cells(fila, QC_IDPERS).Value = idp
    For i = 1 To gPN
        If gpID(i) = idp Then
            wp.Cells(fila, QC_DNI).Value = gpDNI(i): wp.Cells(fila, QC_DNI).NumberFormat = "@"
            wp.Cells(fila, QC_NOMBRE).Value = gpNombre(i)
            Exit For
        End If
    Next i
    wp.Cells(fila, QC_CAUTO).Value = "PENDIENTE DE REVISION"
    wp.Cells(fila, QC_CFINAL).Value = "PENDIENTE DE REVISION"
    wp.Cells(fila, QC_CLAVE).Value = Format(fecha, "YYYY-MM-DD") & "|" & idp
    AsegurarFilaProceso = fila
End Function

'------------------------------------------------------------------------------
' Registra una incidencia desde el flujo guiado (boton INCIDENCIAS).
'------------------------------------------------------------------------------
Public Sub REGISTRAR_INCIDENCIA_GUIADA()
    Dim ws As Worksheet, r As Long
    Dim sFecha As String, sNombre As String, sTipo As String, sObs As String, sDoc As String
    Dim fecha As Date, idp As String
    On Error GoTo fallo
    sFecha = InputBox("Fecha de la incidencia (DD/MM/AAAA):", "Nueva incidencia", _
                      Format(FechaTrabajo(), "DD/MM/YYYY"))
    If Len(sFecha) = 0 Then Exit Sub
    If Not IsDate(sFecha) Then MsgBox "Fecha invalida.", vbExclamation: Exit Sub
    fecha = CDate(sFecha)
    sNombre = InputBox("Apellidos y nombres del trabajador (o ID PERSONAL):", "Nueva incidencia")
    If Len(sNombre) = 0 Then Exit Sub
    idp = sNombre
    If ResolverIDPorNombre(sNombre) <> "" Then idp = ResolverIDPorNombre(sNombre)
    sTipo = InputBox("Tipo de incidencia:" & vbCrLf & _
                     "CLASE / PERMISO / NO PUSO HUELLA / DESCANSO MEDICO / COMISION /" & vbCrLf & _
                     "DESCANSO / FERIADO / VACACIONES / SUSPENSION / FALTA / JUSTIFICADO", _
                     "Nueva incidencia")
    If Len(sTipo) = 0 Then Exit Sub
    sObs = InputBox("Observacion / justificacion:", "Nueva incidencia")
    sDoc = InputBox("Documento sustentatorio (opcional):", "Nueva incidencia")

    Set ws = Hoja(SH_INC)
    r = UltimaFila(ws, 1) + 1
    If r < 3 Then r = 3
    ws.Cells(r, 1).Value = fecha: ws.Cells(r, 1).NumberFormat = "DD/MM/YYYY"
    ws.Cells(r, 2).Value = IIf(ResolverIDPorNombre(sNombre) <> "", idp, "")
    ws.Cells(r, 3).Value = UCase$(sNombre)
    ws.Cells(r, 4).Value = UCase$(sTipo)
    ws.Cells(r, 5).Value = sObs
    ws.Cells(r, 6).Value = sDoc
    ws.Cells(r, 7).Value = CfgTxt("cfg_Responsable")
    ws.Cells(r, 8).Value = Now: ws.Cells(r, 8).NumberFormat = "DD/MM/YYYY HH:MM"

    ' aplicar de inmediato si ya hay proceso de ese dia
    AplicarIncidencias fecha
    modReportes.LlenarReporteDiario fecha
    modPrincipal.ActualizarTablero
    MsgBox "Incidencia registrada y aplicada.", vbInformation, "Incidencias"
    Exit Sub
fallo:
    RegistrarError "modIncidencias.REGISTRAR_INCIDENCIA_GUIADA", Err.Number, Err.Description
End Sub
