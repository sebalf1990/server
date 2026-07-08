Attribute VB_Name = "ModAccountBridge"
' Argentum 20 Game Server
'
'    Copyright (C) 2026 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Marquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit

' ---------------------------------------------------------------------------
' Account bridge: aplica sobre la SQLite local las operaciones de cuenta
' (alta, cambio de password, borrado/restauracion) que la web encola en
' Postgres. El server es el UNICO escritor de su SQLite (Arquitectura B,
' plan 29.001); este modulo solo consume el canal de operaciones (pull) y
' confirma (ack) las que aplico, nunca al reves.
' Contrato del endpoint web (ao20-patreon-commerce, plan 08.001 Fase 1):
'   GET  /api/bridge/account-ops   X-Service-Token: <token>
'        -> 200 text/plain: linea 1 "OPS;<n>", luego una op por linea:
'           create_account: <op_id>;create_account;<email>;<salt>;<password_hash>;<validate_code>
'           set_password:   <op_id>;set_password;<email>;<salt>;<password_hash>
'           set_deleted:    <op_id>;set_deleted;<email>
'           unset_deleted:  <op_id>;unset_deleted;<email>
'   POST /api/bridge/account-ops/ack   X-Service-Token, Content-Type: text/plain
'        body: ids=1,2,3   -> 200 text/plain: OK;<acked_count>
' Fail-silent en toda la cadena: un error de red, de parseo o de SQL nunca
' debe romper el login ni el timer que dispara el poll.
' ---------------------------------------------------------------------------

' Module-level config cache (loaded once on first call)
Private m_ConfigLoaded            As Boolean
Private m_Enabled                 As Boolean
Private m_BaseUrl                 As String
Private m_ServiceToken            As String
Private m_RegisterUrl             As String
Private m_AllowIngameRegistration As Boolean

' Guard de intervalo minimo entre polls (evita llamadas espurias cuando
' Minuto_Timer y el fallback de pull-on-login-attempt coinciden).
Private Const ACCOUNT_BRIDGE_POLL_MIN_INTERVAL_MS As Double = 5000
Private m_LastPollTick  As Long
Private m_HasPolledOnce As Boolean

' ---------------------------------------------------------------------------
' LoadAccountBridgeConfig
' Reads [AccountBridge] section from Server.ini into module-level cache.
' Defaults: bridge OFF, alta in-game OFF (el registro se hace desde la web).
' ---------------------------------------------------------------------------
Private Sub LoadAccountBridgeConfig()
    If m_ConfigLoaded Then Exit Sub
    On Error GoTo LoadAccountBridgeConfig_Err

    Dim ini As clsIniManager
    Set ini = New clsIniManager
    Call ini.Initialize(IniPath & "Server.ini")

    Dim url As String
    url = ini.GetValue("AccountBridge", "BaseUrl")
    If Len(url) = 0 Then url = "http://localhost:4000"
    m_BaseUrl = url

    m_ServiceToken = ini.GetValue("AccountBridge", "ServiceToken")

    Dim registerUrl As String
    registerUrl = ini.GetValue("AccountBridge", "RegisterUrl")
    If Len(registerUrl) = 0 Then registerUrl = "https://ao.muraliarevestimientos.com/registro"
    m_RegisterUrl = registerUrl

    ' Enabled por defecto OFF: el poll nunca corre solo, hay que prenderlo
    ' explicito en Server.ini una vez que la web este configurada.
    m_Enabled = (Trim$(ini.GetValue("AccountBridge", "Enabled")) = "1")

    ' AllowIngameRegistration por defecto OFF: el alta se hace desde la web.
    ' Se puede volver a habilitar como fallback si la web cae.
    m_AllowIngameRegistration = (Trim$(ini.GetValue("AccountBridge", "AllowIngameRegistration")) = "1")

    m_ConfigLoaded = True
    Exit Sub

LoadAccountBridgeConfig_Err:
    ' Si falla la lectura del ini, usamos defaults seguros: bridge apagado y
    ' alta in-game deshabilitada (el login nunca se rompe por esto).
    m_BaseUrl = "http://localhost:4000"
    m_ServiceToken = ""
    m_RegisterUrl = "https://ao.muraliarevestimientos.com/registro"
    m_Enabled = False
    m_AllowIngameRegistration = False
    m_ConfigLoaded = True
End Sub

' ---------------------------------------------------------------------------
' AccountBridge_AllowIngameRegistration  (PUBLIC)
' Usado por Protocol.HandleCreateAccount para decidir si el alta in-game
' sigue habilitada como fallback o si debe redirigir a la web.
' ---------------------------------------------------------------------------
Public Function AccountBridge_AllowIngameRegistration() As Boolean
    Call LoadAccountBridgeConfig
    AccountBridge_AllowIngameRegistration = m_AllowIngameRegistration
End Function

' ---------------------------------------------------------------------------
' AccountBridge_RegisterUrl  (PUBLIC)
' URL de registro que se muestra cuando el alta in-game esta deshabilitada.
' ---------------------------------------------------------------------------
Public Function AccountBridge_RegisterUrl() As String
    Call LoadAccountBridgeConfig
    AccountBridge_RegisterUrl = m_RegisterUrl
End Function

' ---------------------------------------------------------------------------
' AccountBridge_Poll  (PUBLIC entry point)
' Llamado desde Minuto_Timer (cada 1 minuto) y como fallback desde
' Protocol.HandleLoginAccount (pull-on-login-attempt). Guard interno: no
' vuelve a pollear si ya lo hizo hace menos de 5 segundos.
' GET /api/bridge/account-ops, aplica cada operacion sobre la SQLite local
' y hace ack de las aplicadas. On ANY error: logea y sale en silencio, nunca
' rompe el login ni el timer.
' ---------------------------------------------------------------------------
Public Sub AccountBridge_Poll()
    On Error GoTo AccountBridge_Poll_Err

    Call LoadAccountBridgeConfig
    If Not m_Enabled Then Exit Sub

    Dim nowTick As Long
    nowTick = GetTickCountRaw()
    If m_HasPolledOnce Then
        If TicksElapsed(m_LastPollTick, nowTick) < ACCOUNT_BRIDGE_POLL_MIN_INTERVAL_MS Then Exit Sub
    End If
    m_LastPollTick = nowTick
    m_HasPolledOnce = True

    Dim http As MSXML2.ServerXMLHTTP60
    Set http = New MSXML2.ServerXMLHTTP60
    http.setTimeouts 2000, 2000, 3000, 3000

    http.Open "GET", m_BaseUrl & "/api/bridge/account-ops", False
    http.setRequestHeader "X-Service-Token", m_ServiceToken
    http.send

    If http.Status <> 200 Then
        Call LogError("AccountBridge poll GET failed. Status: " & http.Status)
        Exit Sub
    End If

    Dim rawLines() As String
    rawLines = Split(http.responseText, vbLf)
    If UBound(rawLines) < 0 Then Exit Sub

    Dim headerLine As String
    headerLine = Trim$(rawLines(0))
    If Len(headerLine) > 0 Then
        If Right$(headerLine, 1) = vbCr Then headerLine = Left$(headerLine, Len(headerLine) - 1)
    End If

    Dim headerParts() As String
    headerParts = Split(headerLine, ";")
    If UBound(headerParts) < 1 Then Exit Sub
    If UCase$(Trim$(headerParts(0))) <> "OPS" Then Exit Sub

    Dim opCount As Long
    opCount = CLng(Val(headerParts(1)))
    If opCount <= 0 Then Exit Sub

    Dim appliedIds As String
    appliedIds = ""

    Dim i As Long
    For i = 1 To UBound(rawLines)
        Dim oneLine As String
        oneLine = Trim$(rawLines(i))
        If Len(oneLine) > 0 Then
            If Right$(oneLine, 1) = vbCr Then oneLine = Left$(oneLine, Len(oneLine) - 1)
        End If
        If Len(oneLine) > 0 Then
            Dim appliedId As String
            appliedId = ApplyAccountOp(oneLine)
            If Len(appliedId) > 0 Then
                If Len(appliedIds) > 0 Then appliedIds = appliedIds & ","
                appliedIds = appliedIds & appliedId
            End If
        End If
    Next i

    If Len(appliedIds) > 0 Then Call AckAccountOps(appliedIds)

    Exit Sub
AccountBridge_Poll_Err:
    Call LogError("AccountBridge_Poll error: " & Err.Description)
End Sub

' ---------------------------------------------------------------------------
' ApplyAccountOp  (PRIVATE helper)
' Parsea una linea "op_id;tipo;campos..." y aplica la operacion sobre la
' SQLite local. Devuelve el op_id si se aplico (o si es idempotente-aplicada,
' ej. cuenta que ya existe / email que no matchea ninguna fila), o cadena
' vacia si la linea esta malformada (no se ackea, queda en la web para
' diagnostico).
' ---------------------------------------------------------------------------
Private Function ApplyAccountOp(ByVal OpLine As String) As String
    On Error GoTo ApplyAccountOp_Err
    ApplyAccountOp = ""

    Dim parts() As String
    parts = Split(OpLine, ";")
    If UBound(parts) < 2 Then
        Call LogError("AccountBridge op malformada (faltan campos): " & OpLine)
        Exit Function
    End If

    Dim opId   As String
    Dim opType As String
    opId = Trim$(parts(0))
    opType = LCase$(Trim$(parts(1)))

    Select Case opType
        Case "create_account"
            If UBound(parts) < 5 Then
                Call LogError("AccountBridge op create_account malformada: " & OpLine)
                Exit Function
            End If
            Call ApplyCreateAccountOp(opId, Trim$(parts(2)), parts(3), parts(4), parts(5))
            ApplyAccountOp = opId

        Case "set_password"
            If UBound(parts) < 4 Then
                Call LogError("AccountBridge op set_password malformada: " & OpLine)
                Exit Function
            End If
            Call ApplySetPasswordOp(opId, Trim$(parts(2)), parts(3), parts(4))
            ApplyAccountOp = opId

        Case "set_deleted"
            Call ApplySetDeletedOp(opId, Trim$(parts(2)), 1)
            ApplyAccountOp = opId

        Case "unset_deleted"
            Call ApplySetDeletedOp(opId, Trim$(parts(2)), 0)
            ApplyAccountOp = opId

        Case Else
            Call LogError("AccountBridge op con tipo desconocido: " & OpLine)
    End Select

    Exit Function
ApplyAccountOp_Err:
    Call LogError("ApplyAccountOp error: " & Err.Description & " - Line: " & OpLine)
    ApplyAccountOp = ""
End Function

' ---------------------------------------------------------------------------
' ApplyCreateAccountOp  (PRIVATE helper)
' Idempotente: si ya existe una cuenta con ese email, la operacion se
' considera aplicada (se ackea igual) y solo se logea un warning.
' ---------------------------------------------------------------------------
Private Sub ApplyCreateAccountOp(ByVal opId As String, ByVal email As String, ByVal salt As String, ByVal pwHash As String, ByVal validateCode As String)
    On Error GoTo ApplyCreateAccountOp_Err

    Dim existing As ADODB.Recordset
    Set existing = Query("SELECT id FROM account WHERE email = ?;", email)
    If Not (existing Is Nothing) Then
        If Not existing.EOF Then
            Call LogError("AccountBridge create_account op " & opId & ": la cuenta " & email & " ya existe, se trata como aplicada (idempotente).")
            Exit Sub
        End If
    End If

    Call Query("INSERT INTO account (email, password, salt, validate_code, validated) VALUES (?,?,?,?,1);", email, pwHash, salt, validateCode)
    Exit Sub
ApplyCreateAccountOp_Err:
    Call LogError("ApplyCreateAccountOp error: " & Err.Description & " - op " & opId & " email " & email)
End Sub

' ---------------------------------------------------------------------------
' ApplySetPasswordOp  (PRIVATE helper)
' Idempotente: si no existe cuenta con ese email, se considera aplicada (se
' ackea igual, no se re-encola) y solo se logea un warning.
' ---------------------------------------------------------------------------
Private Sub ApplySetPasswordOp(ByVal opId As String, ByVal email As String, ByVal salt As String, ByVal pwHash As String)
    On Error GoTo ApplySetPasswordOp_Err

    Dim existing As ADODB.Recordset
    Set existing = Query("SELECT id FROM account WHERE email = ?;", email)
    If existing Is Nothing Then
        Call LogError("ApplySetPasswordOp: SELECT fallo para email " & email & " - op " & opId)
        Exit Sub
    End If
    If existing.EOF Then
        Call LogError("AccountBridge set_password op " & opId & ": no existe cuenta con email " & email & ", se trata como aplicada.")
        Exit Sub
    End If

    Call Query("UPDATE account SET password = ?, salt = ? WHERE email = ?;", pwHash, salt, email)
    Exit Sub
ApplySetPasswordOp_Err:
    Call LogError("ApplySetPasswordOp error: " & Err.Description & " - op " & opId & " email " & email)
End Sub

' ---------------------------------------------------------------------------
' ApplySetDeletedOp  (PRIVATE helper)
' deletedValue: 1 = set_deleted, 0 = unset_deleted. Idempotente: si no existe
' cuenta con ese email, se considera aplicada y solo se logea un warning.
' ---------------------------------------------------------------------------
Private Sub ApplySetDeletedOp(ByVal opId As String, ByVal email As String, ByVal deletedValue As Integer)
    On Error GoTo ApplySetDeletedOp_Err

    Dim existing As ADODB.Recordset
    Set existing = Query("SELECT id FROM account WHERE email = ?;", email)
    If existing Is Nothing Then
        Call LogError("ApplySetDeletedOp: SELECT fallo para email " & email & " - op " & opId)
        Exit Sub
    End If
    If existing.EOF Then
        Call LogError("AccountBridge set_deleted/unset_deleted op " & opId & ": no existe cuenta con email " & email & ", se trata como aplicada.")
        Exit Sub
    End If

    Call Query("UPDATE account SET deleted = ? WHERE email = ?;", deletedValue, email)
    Exit Sub
ApplySetDeletedOp_Err:
    Call LogError("ApplySetDeletedOp error: " & Err.Description & " - op " & opId & " email " & email)
End Sub

' ---------------------------------------------------------------------------
' AckAccountOps  (PRIVATE helper)
' Confirma en la web los op_id aplicados. Content-Type explicito "text/plain":
' sin ese header express.text no parsea el body (hallazgo del review de
' Fase 1). Fail-silent: si el ack falla, la proxima pasada reaplica
' idempotente (no hay riesgo de duplicar datos).
' ---------------------------------------------------------------------------
Private Sub AckAccountOps(ByVal Ids As String)
    On Error GoTo AckAccountOps_Err

    Dim http As MSXML2.ServerXMLHTTP60
    Set http = New MSXML2.ServerXMLHTTP60
    http.setTimeouts 2000, 2000, 3000, 3000

    http.Open "POST", m_BaseUrl & "/api/bridge/account-ops/ack", False
    http.setRequestHeader "X-Service-Token", m_ServiceToken
    http.setRequestHeader "Content-Type", "text/plain"
    http.send "ids=" & Ids

    If http.Status <> 200 Then
        Call LogError("AccountBridge ack POST failed. Status: " & http.Status & " ids: " & Ids)
    End If

    Exit Sub
AckAccountOps_Err:
    Call LogError("AckAccountOps error: " & Err.Description & " - ids: " & Ids)
End Sub
