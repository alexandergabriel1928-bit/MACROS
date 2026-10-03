Attribute VB_Name = "modBackup"
'==============================================================================
' modBackup  -  Respaldo automatico del archivo
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Crea una copia de respaldo del libro actual.
' Ruta: cfg_RutaBackup (o carpeta \BACKUPS junto al archivo si esta vacia).
' Nombre: BACKUP_ASISTENCIA_AAAA-MM-DD_HHMMSS.<ext>
'------------------------------------------------------------------------------
Public Function CrearRespaldo(Optional ByVal motivo As String = "") As Boolean
    Dim ruta As String, carpeta As String, ext As String, destino As String
    On Error GoTo fallo
    If Len(ThisWorkbook.Path) = 0 Then
        ' archivo nunca guardado: no se puede respaldar aun
        CrearRespaldo = False
        Exit Function
    End If
    carpeta = CfgTxt("cfg_RutaBackup")
    If Len(carpeta) = 0 Then carpeta = ThisWorkbook.Path & Application.PathSeparator & "BACKUPS"
    ' asegurar carpeta
    If Dir(carpeta, vbDirectory) = "" Then MkDir carpeta

    ext = Mid$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") + 1)
    If Len(ext) = 0 Or ext = ThisWorkbook.Name Then ext = "xlsm"

    destino = carpeta & Application.PathSeparator & _
              "BACKUP_ASISTENCIA_" & Format(Now, "YYYY-MM-DD_HHMMSS")
    If Len(motivo) > 0 Then destino = destino & "_" & Replace(motivo, " ", "-")
    destino = destino & "." & ext

    ' guardar primero el libro, luego copiar
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    FileCopy ThisWorkbook.FullName, destino
    Application.DisplayAlerts = True
    CrearRespaldo = True
    Exit Function
fallo:
    Application.DisplayAlerts = True
    RegistrarError "modBackup.CrearRespaldo", Err.Number, Err.Description, False
    CrearRespaldo = False
End Function
