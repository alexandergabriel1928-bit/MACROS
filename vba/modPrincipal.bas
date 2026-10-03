Attribute VB_Name = "modPrincipal"
'==============================================================================
' modPrincipal  -  Acciones del menu, tablero, guardado, cierre de mes y
'                  navegacion. Punto de entrada de los botones.
'==============================================================================
Option Explicit

' Hojas tecnicas que deben ocultarse al usuario final (acceso por boton)
Private Function HojasTecnicas() As Variant
    HojasTecnicas = Array(SH_DATA, SH_PROC, SH_AUD, SH_CAT, SH_LOG, SH_HIST, SH_CUAD)
End Function

'------------------------------------------------------------------------------
' Navegacion basica (muestra la hoja aunque este oculta)
'------------------------------------------------------------------------------
Public Sub IrAHoja(ByVal nombre As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = Hoja(nombre)
    ws.Visible = xlSheetVisible
    ws.Activate
    On Error GoTo 0
End Sub

Public Sub VolverAlMenu()
    IrAHoja SH_MENU
    OcultarTecnicas
End Sub

Public Sub ABRIR_PERSONAL(): IrAHoja SH_PERS: End Sub
Public Sub ABRIR_CONFIGURACION(): IrAHoja SH_CONFIG: End Sub
Public Sub ABRIR_INCIDENCIAS(): IrAHoja SH_INC: End Sub
Public Sub ABRIR_AUDITORIA(): IrAHoja SH_AUD: End Sub
Public Sub ABRIR_HISTORICO(): IrAHoja SH_HIST: End Sub

'------------------------------------------------------------------------------
' Ocultar / mostrar hojas tecnicas
'------------------------------------------------------------------------------
Public Sub OcultarTecnicas()
    Dim v, nm
    On Error Resume Next
    For Each nm In HojasTecnicas()
        Hoja(CStr(nm)).Visible = xlSheetVeryHidden
    Next nm
    On Error GoTo 0
End Sub

Public Sub MostrarTecnicas()
    Dim nm
    On Error Resume Next
    For Each nm In HojasTecnicas()
        Hoja(CStr(nm)).Visible = xlSheetVisible
    Next nm
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Tablero del menu (requisito 49)
'------------------------------------------------------------------------------
Public Sub ActualizarTablero()
    Dim fecha As Date, n As Long, i As Long, fila As Long, wp As Worksheet
    Dim act As Long, pres As Long, punt As Long, tard As Long, falta As Long
    Dim inc As Long, pend As Long, cond As String
    On Error Resume Next
    fecha = FechaTrabajo()
    Set wp = Hoja(SH_PROC)
    n = CargarPersonal()
    For i = 1 To n
        If gpActivo(i) Then
            act = act + 1
            fila = FilaProceso(fecha, gpID(i))
            If fila > 0 Then
                cond = UCase$(Trim$(CStr(wp.Cells(fila, QC_CFINAL).Value)))
                Select Case cond
                    Case "PUNTUAL": punt = punt + 1: pres = pres + 1
                    Case "TARDANZA": tard = tard + 1: pres = pres + 1
                    Case "FALTA": falta = falta + 1
                    Case "PENDIENTE DE REVISION", "INCOMPLETO": pend = pend + 1
                    Case "DESCANSO", "FERIADO"
                    Case Else: inc = inc + 1: pres = pres + 1
                End Select
            End If
        End If
    Next i
    EscribirTablero "dash_PERSONAL_ACTIVO", act
    EscribirTablero "dash_PRESENTES_HOY", pres
    EscribirTablero "dash_PUNTUALES", punt
    EscribirTablero "dash_TARDANZAS", tard
    EscribirTablero "dash_FALTAS", falta
    EscribirTablero "dash_INCIDENCIAS", inc
    EscribirTablero "dash_PENDIENTES", pend
    If act > 0 Then
        EscribirTablero "dash_PCT_PUNTUALIDAD", Format(punt / act, "0%")
        EscribirTablero "dash_PCT_ASISTENCIA", Format(pres / act, "0%")
    End If
    EscribirTablero "dash_PERIODO_ACTIVO", MesES(PeriodoMes()) & " " & PeriodoAnio()
    EscribirTablero "dash_ESTADO_PERIODO", EstadoPeriodo()
    On Error GoTo 0
End Sub

Private Sub EscribirTablero(ByVal nombre As String, ByVal valor As Variant)
    On Error Resume Next
    ThisWorkbook.Names(nombre).RefersToRange.Value = valor
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' GUARDAR Y RESPALDAR (requisito 52)
'------------------------------------------------------------------------------
Public Sub GUARDAR_Y_RESPALDAR()
    On Error GoTo fallo
    If MsgBox("Se guardara el archivo y se creara un respaldo. Continuar?", _
              vbQuestion + vbYesNo, "Guardar y respaldar") <> vbYes Then Exit Sub
    If CrearRespaldo("manual") Then
        MsgBox "Informacion guardada correctamente y respaldo creado.", vbInformation, "Guardar"
    Else
        ThisWorkbook.Save
        MsgBox "Archivo guardado. (No se pudo crear respaldo: revise la ruta en CONFIG o guarde el libro primero).", _
               vbExclamation, "Guardar"
    End If
    Exit Sub
fallo:
    RegistrarError "modPrincipal.GUARDAR_Y_RESPALDAR", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' CERRAR MES (requisitos 23, 24, 45)
'------------------------------------------------------------------------------
Public Sub CERRAR_MES()
    Dim mes As Integer, anio As Integer, tot(1 To 10) As Long, resp As VbMsgBoxResult
    Dim msg As String
    On Error GoTo fallo
    mes = PeriodoMes(): anio = PeriodoAnio()
    If PeriodoCerrado(DateSerial(anio, mes, 1)) Then
        MsgBox "El periodo " & MesES(mes) & " " & anio & " ya esta CERRADO.", vbExclamation
        Exit Sub
    End If
    modReportes.ResumenMensual mes, anio, tot
    msg = "RESUMEN DEL PERIODO " & MesES(mes) & " " & anio & vbCrLf & String(32, "-") & vbCrLf & _
          "Trabajadores:        " & tot(1) & vbCrLf & _
          "Puntuales:           " & tot(2) & vbCrLf & _
          "Tardanzas:           " & tot(3) & vbCrLf & _
          "Faltas:              " & tot(4) & vbCrLf & _
          "Permisos:            " & tot(5) & vbCrLf & _
          "Clases:              " & tot(6) & vbCrLf & _
          "No puso huella:      " & tot(7) & vbCrLf & _
          "Descanso medico:     " & tot(8) & vbCrLf & _
          "Comisiones:          " & tot(9) & vbCrLf & _
          "PENDIENTES:          " & tot(10) & vbCrLf & String(32, "-")
    If tot(10) > 0 Then
        resp = MsgBox(msg & vbCrLf & "EXISTEN REGISTROS PENDIENTES DE REVISION." & vbCrLf & _
                      "VOLVER (No) para resolverlos, o CERRAR DE TODAS FORMAS (Si)?", _
                      vbExclamation + vbYesNo, "Cerrar mes")
        If resp <> vbYes Then Exit Sub
    End If
    resp = MsgBox(msg & vbCrLf & "Desea CERRAR DEFINITIVAMENTE el periodo?" & vbCrLf & _
                  "(El mes quedara congelado y no se podra reprocesar).", _
                  vbQuestion + vbYesNo, "Cerrar mes")
    If resp <> vbYes Then Exit Sub

    ' respaldo + consolidado + cuadro + historico
    CrearRespaldo "cierre"
    modConsolidado.GENERAR_CONSOLIDADO
    GuardarHistorico mes, anio
    ThisWorkbook.Names("cfg_EstadoPeriodo").RefersToRange.Value = "CERRADO"
    ThisWorkbook.Save
    MsgBox "Periodo " & MesES(mes) & " " & anio & " CERRADO y archivado en HISTORICO_MENSUAL.", _
           vbInformation, "Cerrar mes"
    Exit Sub
fallo:
    RegistrarError "modPrincipal.CERRAR_MES", Err.Number, Err.Description
End Sub

Private Sub GuardarHistorico(ByVal mes As Integer, ByVal anio As Integer)
    Dim wh As Worksheet, i As Long, n As Long, r As Long, per As String
    Dim cP As Long, cT As Long, cF As Long, cPe As Long, cCl As Long, cNH As Long
    Dim cDM As Long, cCom As Long, cFer As Long, cDesc As Long, cVac As Long
    Dim diasP As Long, diasE As Long
    Set wh = Hoja(SH_HIST)
    per = Format(DateSerial(anio, mes, 1), "YYYY-MM")
    n = CargarPersonal()
    r = UltimaFila(wh, 1) + 1
    If r < 2 Then r = 2
    For i = 1 To n
        If gpActivo(i) Then
            modConsolidado.ContarPorMes Hoja(SH_PROC), gpID(i), mes, anio, cP, cT, cF, cPe, cCl, _
                                        cNH, cDM, cCom, cFer, cDesc, cVac, diasP, diasE
            wh.Cells(r, 1).Value = per
            wh.Cells(r, 2).Value = gpID(i)
            wh.Cells(r, 3).Value = gpDNI(i): wh.Cells(r, 3).NumberFormat = "@"
            wh.Cells(r, 4).Value = gpNombre(i)
            wh.Cells(r, 5).Value = cP: wh.Cells(r, 6).Value = cT: wh.Cells(r, 7).Value = cF
            wh.Cells(r, 8).Value = cPe: wh.Cells(r, 9).Value = cCl: wh.Cells(r, 10).Value = cNH
            wh.Cells(r, 11).Value = cDM: wh.Cells(r, 12).Value = cCom: wh.Cells(r, 13).Value = cVac
            wh.Cells(r, 14).Value = cFer: wh.Cells(r, 15).Value = cDesc: wh.Cells(r, 16).Value = diasE
            wh.Cells(r, 18).Value = Date: wh.Cells(r, 18).NumberFormat = "DD/MM/YYYY"
            wh.Cells(r, 19).Value = UsuarioActual()
            r = r + 1
        End If
    Next i
End Sub

'------------------------------------------------------------------------------
' Deteccion de nuevo periodo (requisito 46). Llamado desde Workbook_Open.
'------------------------------------------------------------------------------
Public Sub DetectarNuevoPeriodo()
    Dim mActual As Integer, aActual As Integer
    On Error Resume Next
    mActual = Month(Date): aActual = Year(Date)
    If PeriodoMes() <> mActual Or PeriodoAnio() <> aActual Then
        If MsgBox("Se ha detectado un nuevo periodo. Desea activar el periodo de " & _
                  MesES(mActual) & " " & aActual & "?" & vbCrLf & _
                  "(Los meses anteriores se conservan intactos).", _
                  vbQuestion + vbYesNo, "Nuevo periodo") = vbYes Then
            ThisWorkbook.Names("cfg_PeriodoMes").RefersToRange.Value = mActual
            ThisWorkbook.Names("cfg_PeriodoAnio").RefersToRange.Value = aActual
            ThisWorkbook.Names("cfg_EstadoPeriodo").RefersToRange.Value = "ABIERTO"
        End If
    End If
    On Error GoTo 0
End Sub
