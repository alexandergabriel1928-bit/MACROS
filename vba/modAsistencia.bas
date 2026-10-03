Attribute VB_Name = "modAsistencia"
'==============================================================================
' modAsistencia  -  Orquestador del procesamiento de asistencia (1 dia)
'  RAW -> PROCESAMIENTO -> CONTROL.  Nunca modifica DATA_HUELLERO (RAW).
'  Respeta periodos CERRADOS y decisiones MANUALES previas.
'==============================================================================
Option Explicit

' Columnas de PROCESO
Public Const QC_FECHA As Long = 1
Public Const QC_IDPERS As Long = 2
Public Const QC_DNI As Long = 3
Public Const QC_NOMBRE As Long = 4
Public Const QC_ENTM As Long = 5
Public Const QC_SALALM As Long = 6
Public Const QC_ENTT As Long = 7
Public Const QC_SALF As Long = 8
Public Const QC_MINTM As Long = 9
Public Const QC_MINTT As Long = 10
Public Const QC_CAUTO As Long = 11
Public Const QC_CFINAL As Long = 12
Public Const QC_CODIGO As Long = 13
Public Const QC_OBS As Long = 14
Public Const QC_ORIGEN As Long = 15
Public Const QC_MANUAL As Long = 16
Public Const QC_USER As Long = 17
Public Const QC_FPROC As Long = 18
Public Const QC_CLAVE As Long = 19

'------------------------------------------------------------------------------
Public Sub PROCESAR_ASISTENCIA()
    Dim fecha As Date, ws As Worksheet, n As Long, i As Long
    Dim j As TJornada, idp As String, entrada As Double, tipoDia As String
    Dim condAuto As String, minTM As Long, minTT As Long, condTarde As String, dummy As Long
    Dim dicProc As Object, key As String, fila As Long
    Dim esManual As Boolean, condFinalPrev As String, obsPrev As String

    On Error GoTo fallo
    fecha = FechaTrabajo()

    ' periodo cerrado -> bloquear
    If PeriodoCerrado(fecha) Then
        MsgBox "El periodo de " & Format(fecha, "MMMM YYYY") & " esta CERRADO." & vbCrLf & _
               "No se puede reprocesar un mes cerrado.", vbExclamation, "Procesar"
        Exit Sub
    End If

    If MsgBox("Se procesara la asistencia del " & Format(fecha, "DD/MM/YYYY") & "." & vbCrLf & _
              "Esta accion recalcula el dia (sin tocar decisiones manuales). Continuar?", _
              vbQuestion + vbYesNo, "Procesar asistencia") <> vbYes Then Exit Sub

    ' validaciones previas (no detienen, informan)
    modValidaciones.ValidarAntesDeProcesar fecha, False

    Set ws = Hoja(SH_PROC)
    n = CargarPersonal()
    If n = 0 Then MsgBox "No hay personal en la nomina.", vbExclamation: Exit Sub
    CargarMarcacionesDelDia fecha

    ' indexar filas PROCESO existentes por clave
    Set dicProc = CreateObject("Scripting.Dictionary")
    Dim ult As Long, r As Long
    ult = UltimaFila(ws, QC_CLAVE)
    For r = 2 To ult
        key = Trim$(CStr(ws.Cells(r, QC_CLAVE).Value))
        If Len(key) > 0 Then dicProc(key) = r
    Next r

    TurboOn
    For i = 1 To n
        If gpActivo(i) Then
            Progreso "Procesando trabajador", i, n
            idp = gpID(i)
            key = Format(fecha, "YYYY-MM-DD") & "|" & idp
            j = ObtenerJornada(i)
            entrada = EntradaEfectiva(fecha, idp)
            tipoDia = TipoDeDia(fecha)

            ' --- condicion automatica base ---
            minTM = 0: minTT = 0
            If tipoDia = "FERIADO" Then
                condAuto = "FERIADO"
            ElseIf tipoDia = "DESCANSO" Or entrada = -1 Then
                condAuto = "DESCANSO"
            Else
                If Not j.tieneMarca Then
                    condAuto = "PENDIENTE DE REVISION"   ' no asumir FALTA (lo decide el responsable)
                ElseIf Not j.entManianaOK And j.entTardeOK Then
                    condAuto = "INCOMPLETO"
                ElseIf j.entManianaOK Then
                    condAuto = EvaluarPuntualidad(j.entManiana, entrada, ToleranciaManianaSeg(), minTM)
                Else
                    condAuto = "INCOMPLETO"
                End If
                ' tardanza de la tarde (informativa)
                If j.entTardeOK Then
                    condTarde = EvaluarPuntualidad(j.entTarde, EntradaTardeEfectiva(idp), ToleranciaTardeSeg(), minTT)
                End If
            End If

            ' --- ubicar/crear fila y respetar decision manual previa ---
            esManual = False: condFinalPrev = "": obsPrev = ""
            If dicProc.Exists(key) Then
                fila = dicProc(key)
                If UCase$(Trim$(CStr(ws.Cells(fila, QC_MANUAL).Value))) = "SI" Then
                    esManual = True
                    condFinalPrev = Trim$(CStr(ws.Cells(fila, QC_CFINAL).Value))
                    obsPrev = Trim$(CStr(ws.Cells(fila, QC_OBS).Value))
                End If
            Else
                fila = ult + 1: ult = ult + 1
                dicProc(key) = fila
            End If

            ' --- escribir fila ---
            ws.Cells(fila, QC_FECHA).Value = fecha: ws.Cells(fila, QC_FECHA).NumberFormat = "DD/MM/YYYY"
            ws.Cells(fila, QC_IDPERS).Value = idp
            ws.Cells(fila, QC_DNI).Value = gpDNI(i): ws.Cells(fila, QC_DNI).NumberFormat = "@"
            ws.Cells(fila, QC_NOMBRE).Value = gpNombre(i)
            EscribirHora ws.Cells(fila, QC_ENTM), j.entManianaOK, j.entManiana
            EscribirHora ws.Cells(fila, QC_SALALM), j.salAlmuerzoOK, j.salAlmuerzo
            EscribirHora ws.Cells(fila, QC_ENTT), j.entTardeOK, j.entTarde
            EscribirHora ws.Cells(fila, QC_SALF), j.salFinalOK, j.salFinal
            ws.Cells(fila, QC_MINTM).Value = minTM
            ws.Cells(fila, QC_MINTT).Value = minTT
            ws.Cells(fila, QC_CAUTO).Value = condAuto
            If esManual Then
                ws.Cells(fila, QC_CFINAL).Value = condFinalPrev
                ws.Cells(fila, QC_OBS).Value = obsPrev
                ws.Cells(fila, QC_MANUAL).Value = "SI"
            Else
                ws.Cells(fila, QC_CFINAL).Value = condAuto
                ws.Cells(fila, QC_MANUAL).Value = "NO"
            End If
            ws.Cells(fila, QC_CODIGO).Value = CodigoCuadro(CStr(ws.Cells(fila, QC_CFINAL).Value))
            If j.duplicidad Then ws.Cells(fila, QC_OBS).Value = _
                Trim$(ws.Cells(fila, QC_OBS).Value & " [DUPLICIDAD DE MARCACION]")
            ws.Cells(fila, QC_ORIGEN).Value = "HUELLERO"
            ws.Cells(fila, QC_USER).Value = UsuarioActual()
            ws.Cells(fila, QC_FPROC).Value = Now
            ws.Cells(fila, QC_FPROC).NumberFormat = "DD/MM/YYYY HH:MM:SS"
            ws.Cells(fila, QC_CLAVE).Value = key
        End If
    Next i

    ' aplicar incidencias (prioridad) -> marca MANUAL las que correspondan
    modIncidencias.AplicarIncidencias fecha
    TurboOff
    ProgresoFin

    ' refrescar reporte diario + tablero
    modReportes.LlenarReporteDiario fecha
    modPrincipal.ActualizarTablero

    MsgBox "Asistencia del " & Format(fecha, "DD/MM/YYYY") & " procesada correctamente.", _
           vbInformation, "Procesar asistencia"
    Exit Sub
fallo:
    TurboOff: ProgresoFin
    RegistrarError "modAsistencia.PROCESAR_ASISTENCIA", Err.Number, Err.Description
End Sub

Private Sub EscribirHora(ByVal celda As Range, ByVal ok As Boolean, ByVal h As Date)
    If ok Then
        celda.Value = h
        celda.NumberFormat = "HH:MM:SS"
    Else
        celda.ClearContents
    End If
End Sub

'------------------------------------------------------------------------------
' Devuelve la fila de PROCESO para (fecha, idPersonal) o 0 si no existe.
'------------------------------------------------------------------------------
Public Function FilaProceso(ByVal fecha As Date, ByVal idPersonal As String) As Long
    Dim ws As Worksheet, ult As Long, r As Long, key As String
    Set ws = Hoja(SH_PROC)
    key = Format(fecha, "YYYY-MM-DD") & "|" & idPersonal
    ult = UltimaFila(ws, QC_CLAVE)
    For r = 2 To ult
        If Trim$(CStr(ws.Cells(r, QC_CLAVE).Value)) = key Then FilaProceso = r: Exit Function
    Next r
    FilaProceso = 0
End Function
