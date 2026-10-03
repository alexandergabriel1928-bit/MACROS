Attribute VB_Name = "modMarcaciones"
'==============================================================================
' modMarcaciones  -  Clasificacion de marcaciones por trabajador y dia
'  Determina ENTRADA MAÑANA / SALIDA ALMUERZO / ENTRADA TARDE / SALIDA FINAL
'  usando el contexto horario (funciona aunque el campo ESTADO venga vacio).
'==============================================================================
Option Explicit

Public Type TJornada
    tieneMarca As Boolean
    nMarcas As Long
    entManiana As Date
    salAlmuerzo As Date
    entTarde As Date
    salFinal As Date
    entManianaOK As Boolean
    salAlmuerzoOK As Boolean
    entTardeOK As Boolean
    salFinalOK As Boolean
    duplicidad As Boolean
End Type

' Cache de marcaciones del dia procesado
Private mFecha As Date
Private mN As Long
Private mIdx() As Long        ' indice de trabajador (0 si no identificado)
Private mHora() As Date       ' hora del dia (parte horaria)

'------------------------------------------------------------------------------
' Carga a memoria todas las marcaciones de DATA_HUELLERO para una fecha dada,
' resolviendo a que trabajador pertenece cada una. Requiere CargarPersonal antes.
'------------------------------------------------------------------------------
Public Sub CargarMarcacionesDelDia(ByVal fecha As Date)
    Dim ws As Worksheet, ult As Long, r As Long, k As Long
    Dim fh As Variant, idHue As String, nombre As String
    Set ws = Hoja(SH_DATA)
    ult = UltimaFila(ws, 1)
    ReDim mIdx(1 To Application.Max(1, ult))
    ReDim mHora(1 To Application.Max(1, ult))
    k = 0
    mFecha = Int(fecha)
    For r = 2 To ult
        fh = ws.Cells(r, 3).Value   ' FECHA Y HORA
        If IsDate(fh) Then
            If Int(CDate(fh)) = mFecha Then
                idHue = Trim$(CStr(ws.Cells(r, 1).Value))
                nombre = Trim$(CStr(ws.Cells(r, 2).Value))
                k = k + 1
                mIdx(k) = IndicePorMarcacion(idHue, nombre)
                mHora(k) = CDate(fh) - Int(CDate(fh))
            End If
        End If
    Next r
    mN = k
End Sub

'------------------------------------------------------------------------------
' Devuelve la jornada clasificada de un trabajador (indice de la nomina).
'------------------------------------------------------------------------------
Public Function ObtenerJornada(ByVal idx As Long) As TJornada
    Dim j As TJornada, i As Long, nmar As Long
    Dim horas() As Date, c As Long, a As Long, b As Long, tmp As Date
    Dim bMidSeg As Long, salAlmSeg As Long, entTardeSeg As Long
    Dim mañMin As Date, mañMax As Date, tarMin As Date, tarMax As Date
    Dim cMan As Long, cTar As Long
    Dim idp As String

    ' recolectar marcas del trabajador
    ReDim horas(1 To Application.Max(1, mN))
    c = 0
    For i = 1 To mN
        If mIdx(i) = idx Then
            c = c + 1
            horas(c) = mHora(i)
        End If
    Next i
    j.nMarcas = c
    j.tieneMarca = (c > 0)
    If c = 0 Then ObtenerJornada = j: Exit Function

    ' ordenar ascendente
    For a = 1 To c - 1
        For b = a + 1 To c
            If horas(b) < horas(a) Then tmp = horas(a): horas(a) = horas(b): horas(b) = tmp
        Next b
    Next a

    ' duplicidad: marcas a < 60 seg
    For a = 1 To c - 1
        If SegundosDelDia(horas(a + 1)) - SegundosDelDia(horas(a)) < 60 Then j.duplicidad = True
    Next a

    ' frontera manana/tarde (punto medio entre salida almuerzo y entrada tarde)
    idp = gpID(idx)
    If IsDate(Cfg("cfg_SalAlmuerzo")) Then
        salAlmSeg = SegundosDelDia(CDate(Cfg("cfg_SalAlmuerzo")))
    End If
    entTardeSeg = CLng(EntradaTardeEfectiva(idp) * 86400#)
    If salAlmSeg = 0 Then salAlmSeg = 13 * 3600
    If entTardeSeg = 0 Then entTardeSeg = 15 * 3600
    bMidSeg = (salAlmSeg + entTardeSeg) \ 2

    ' separar y tomar min/max de cada bloque
    cMan = 0: cTar = 0
    For a = 1 To c
        If SegundosDelDia(horas(a)) <= bMidSeg Then
            If cMan = 0 Then mañMin = horas(a)
            mañMax = horas(a): cMan = cMan + 1
        Else
            If cTar = 0 Then tarMin = horas(a)
            tarMax = horas(a): cTar = cTar + 1
        End If
    Next a

    If cMan >= 1 Then
        j.entManiana = mañMin: j.entManianaOK = True
        If cMan >= 2 Then j.salAlmuerzo = mañMax: j.salAlmuerzoOK = True
    End If
    If cTar >= 1 Then
        j.entTarde = tarMin: j.entTardeOK = True
        If cTar >= 2 Then
            j.salFinal = tarMax: j.salFinalOK = True
        Else
            ' una sola marca de tarde: tomarla tambien como posible salida final
            j.salFinal = tarMax: j.salFinalOK = (cMan = 0)
        End If
    End If
    ObtenerJornada = j
End Function

'------------------------------------------------------------------------------
' Evalua puntualidad de una hora de entrada contra la hora oficial efectiva.
' Comparacion EXACTA al segundo (sin redondeo).  Devuelve "PUNTUAL"/"TARDANZA".
' minTard (ByRef) = minutos de tardanza (techo), 0 si puntual.
'------------------------------------------------------------------------------
Public Function EvaluarPuntualidad(ByVal horaMarca As Date, _
                                   ByVal horaOficial As Double, _
                                   ByVal tolSeg As Long, _
                                   ByRef minTard As Long) As String
    Dim marcaSeg As Long, ofiSeg As Long, cutoffSeg As Long, retraso As Long
    marcaSeg = SegundosDelDia(horaMarca)
    ofiSeg = CLng(horaOficial * 86400#)
    cutoffSeg = ofiSeg + tolSeg
    If marcaSeg <= cutoffSeg Then
        minTard = 0
        EvaluarPuntualidad = "PUNTUAL"
    Else
        retraso = marcaSeg - ofiSeg
        minTard = retraso \ 60
        If (retraso Mod 60) > 0 Then minTard = minTard + 1
        EvaluarPuntualidad = "TARDANZA"
    End If
End Function
