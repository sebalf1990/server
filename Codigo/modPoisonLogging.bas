Attribute VB_Name = "modPoisonLogging"
' Argentum 20 Game Server
'
' modPoisonLogging
' Sistema de venenos (TOGGLE26 new_poison_system).
' Escribe un archivo por dia en Logs\Venenos_YYYY-MM-DD.log.
' Activo solo cuando el feature toggle esta encendido.
' Formato CSV-like:
'   timestamp;event;attacker;victim;preset_id;family;damage;stacks;resist;factor
' event en {apply, tick, expire, resist, cure, vial_use}

Option Explicit

Private Const POISON_LOG_HEADER As String = "timestamp;event;attacker;victim;preset_id;family;damage;stacks;resist;factor"

Public Sub LogPoisonEvent(ByVal evento As String, _
                          ByVal attacker As String, _
                          ByVal victim As String, _
                          ByVal presetId As Long, _
                          ByVal family As Long, _
                          ByVal damage As Long, _
                          ByVal stacks As Long, _
                          ByVal resist As Long, _
                          ByVal factor As Single)
    On Error GoTo ErrHandler
    If Not IsFeatureEnabled("new_poison_system") Then Exit Sub
    Dim today As String
    today = Format$(Date, "yyyy-mm-dd")
    Dim fname As String
    fname = App.Path & "\Logs\Venenos_" & today & ".log"
    Dim fnum As Integer
    fnum = FreeFile
    Dim writeHeader As Boolean
    writeHeader = (LenB(dir(fname)) = 0)
    Open fname For Append As #fnum
    If writeHeader Then Print #fnum, POISON_LOG_HEADER
    Print #fnum, Format$(Now, "yyyy-mm-dd hh:nn:ss") & ";" & _
                 evento & ";" & attacker & ";" & victim & ";" & _
                 presetId & ";" & family & ";" & damage & ";" & _
                 stacks & ";" & resist & ";" & factor
    Close #fnum
    Exit Sub
ErrHandler:
    On Error Resume Next
    Close #fnum
End Sub
