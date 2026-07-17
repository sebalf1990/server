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

' ---- MAO bridge (mercado de personajes, plan 16.002): toggle e intervalo propios ----
Private m_MaoEnabled       As Boolean
Private Const MAO_BRIDGE_POLL_MIN_INTERVAL_MS As Double = 5000
Private m_LastMaoPollTick  As Long
Private m_HasPolledMaoOnce As Boolean

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

    ' MAO (mercado de personajes, plan 16.002): OFF por defecto. Se prende en
    ' Server.ini una vez que la web tenga /api/bridge/mao-ops vivo.
    m_MaoEnabled = (Trim$(ini.GetValue("AccountBridge", "MaoEnabled")) = "1")

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
    m_MaoEnabled = False
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

' ===========================================================================
' MAO bridge (Mercado de Personajes, plan 16.002)
' Segundo canal de ops sobre el mismo [AccountBridge] (BaseUrl/ServiceToken),
' con su propio toggle m_MaoEnabled. El server sigue siendo el UNICO escritor
' de su SQLite (Arquitectura B); este canal solo aplica (pull) y confirma (ack).
' Contrato del endpoint web (16.002 F2):
'   GET  /api/bridge/mao-ops   X-Service-Token
'        -> 200 text/plain: linea 1 "MAOPS;<n>", luego una op por linea:
'           <op_id>;lock_character;<seller_account_id>;<char_id>
'           <op_id>;unlock_character;<seller_account_id>;<char_id>
'           <op_id>;transfer_character;<from_account_id>;<to_account_id>;<char_id>;<sale_transaction_id>
'   POST /api/bridge/mao-ops/ack   X-Service-Token, Content-Type: text/plain
'        body: una linea por op -> "<op_id>;ok" | "<op_id>;error;<detail_code>"
'        -> 200 text/plain: OK;<acked_count>
' Verificacion de escritura por READ-BACK de estado (no changes()/RecordsAffected):
' el UPDATE es un compare-and-swap en el WHERE y luego se relee la fila por PK y
' se interpreta su estado final. Verifica la post-condicion real y resuelve la
' idempotencia gratis. El precio (ARS) vive SOLO en la web; el server no escribe
' price_in_mao (decision 16.002 F0).
' ===========================================================================

' ---------------------------------------------------------------------------
' MaoBridge_Poll  (PUBLIC entry point). Llamado desde Minuto_Timer. Guard de
' intervalo propio. Fail-silent en toda la cadena.
' ---------------------------------------------------------------------------
Public Sub MaoBridge_Poll()
    On Error GoTo MaoBridge_Poll_Err

    Call LoadAccountBridgeConfig
    If Not m_MaoEnabled Then Exit Sub

    Dim nowTick As Long
    nowTick = GetTickCountRaw()
    If m_HasPolledMaoOnce Then
        If TicksElapsed(m_LastMaoPollTick, nowTick) < MAO_BRIDGE_POLL_MIN_INTERVAL_MS Then Exit Sub
    End If
    m_LastMaoPollTick = nowTick
    m_HasPolledMaoOnce = True

    Dim http As MSXML2.ServerXMLHTTP60
    Set http = New MSXML2.ServerXMLHTTP60
    http.setTimeouts 2000, 2000, 3000, 3000

    http.Open "GET", m_BaseUrl & "/api/bridge/mao-ops", False
    http.setRequestHeader "X-Service-Token", m_ServiceToken
    http.send

    If http.Status <> 200 Then
        Call LogError("MaoBridge poll GET failed. Status: " & http.Status)
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
    If UCase$(Trim$(headerParts(0))) <> "MAOPS" Then Exit Sub

    Dim opCount As Long
    opCount = CLng(Val(headerParts(1)))
    If opCount <= 0 Then Exit Sub

    Dim ackBody As String
    ackBody = ""

    Dim i As Long
    For i = 1 To UBound(rawLines)
        Dim oneLine As String
        oneLine = Trim$(rawLines(i))
        If Len(oneLine) > 0 Then
            If Right$(oneLine, 1) = vbCr Then oneLine = Left$(oneLine, Len(oneLine) - 1)
        End If
        If Len(oneLine) > 0 Then
            Dim ackLine As String
            ackLine = ApplyMaoOp(oneLine)
            If Len(ackLine) > 0 Then
                If Len(ackBody) > 0 Then ackBody = ackBody & vbLf
                ackBody = ackBody & ackLine
            End If
        End If
    Next i

    If Len(ackBody) > 0 Then Call AckMaoOps(ackBody)

    Exit Sub
MaoBridge_Poll_Err:
    Call LogError("MaoBridge_Poll error: " & Err.Description)
End Sub

' ---------------------------------------------------------------------------
' ApplyMaoOp  (PRIVATE). Parsea "op_id;tipo;campos..." y despacha. Devuelve la
' linea de ack ("<op_id>;ok" | "<op_id>;error;<detail>") o "" si la linea esta
' malformada (no se ackea). OJO: el contrato manda <seller/from>;<char/to>; las
' funciones reciben charId primero, por eso el mapeo posicional explicito.
' ---------------------------------------------------------------------------
Private Function ApplyMaoOp(ByVal OpLine As String) As String
    On Error GoTo ApplyMaoOp_Err
    ApplyMaoOp = ""

    Dim opId As String
    opId = ""

    Dim parts() As String
    parts = Split(OpLine, ";")
    If UBound(parts) < 2 Then
        Call LogError("MaoBridge op malformada (faltan campos): " & OpLine)
        Exit Function
    End If

    Dim opType As String
    opId = Trim$(parts(0))
    opType = LCase$(Trim$(parts(1)))

    Select Case opType
        Case "lock_character"
            If UBound(parts) < 3 Then
                Call LogError("MaoBridge op lock_character malformada: " & OpLine)
                Exit Function
            End If
            ApplyMaoOp = opId & ";" & ApplyMaoLock(CLng(parts(3)), CLng(parts(2)))

        Case "unlock_character"
            If UBound(parts) < 3 Then
                Call LogError("MaoBridge op unlock_character malformada: " & OpLine)
                Exit Function
            End If
            ApplyMaoOp = opId & ";" & ApplyMaoUnlock(CLng(parts(3)), CLng(parts(2)))

        Case "transfer_character"
            If UBound(parts) < 5 Then
                Call LogError("MaoBridge op transfer_character malformada: " & OpLine)
                Exit Function
            End If
            ApplyMaoOp = opId & ";" & ApplyMaoTransfer(CLng(parts(4)), CLng(parts(2)), CLng(parts(3)), Trim$(parts(5)))

        Case Else
            Call LogError("MaoBridge op con tipo desconocido: " & OpLine)
    End Select

    Exit Function
ApplyMaoOp_Err:
    Call LogError("ApplyMaoOp error: " & Err.Description & " - Line: " & OpLine)
    If Len(opId) > 0 Then
        ApplyMaoOp = opId & ";error;unknown_error"
    Else
        ApplyMaoOp = ""
    End If
End Function

' ---------------------------------------------------------------------------
' ApplyMaoLock  (PRIVATE)  -> "ok" | "error;<detail>". CAS: bloquea solo si el
' pj es del vendedor y no estaba bloqueado; relee por PK y verifica el estado
' final (idempotente). Si el pj esta online, lo kickea para que no pueda vaciar
' el inventario mientras esta publicado.
' ---------------------------------------------------------------------------
Private Function ApplyMaoLock(ByVal charId As Long, ByVal sellerAcc As Long) As String
    On Error GoTo ApplyMaoLock_Err

    Call Query("UPDATE user SET is_locked_in_mao = 1 WHERE id = ? AND account_id = ? AND is_locked_in_mao = 0;", charId, sellerAcc)

    Dim rs As ADODB.Recordset
    Set rs = Query("SELECT name, account_id, is_locked_in_mao FROM user WHERE id = ?;", charId)
    If rs Is Nothing Then
        ApplyMaoLock = "error;unknown_error"
        Exit Function
    End If
    If rs.EOF Then
        ApplyMaoLock = "error;char_not_found"
        Exit Function
    End If
    If CLng(rs!account_id) <> sellerAcc Then
        ApplyMaoLock = "error;owner_mismatch"
        Exit Function
    End If
    If Not CBool(rs!is_locked_in_mao) Then
        ApplyMaoLock = "error;unknown_error"
        Exit Function
    End If

    Dim charName As String
    charName = CStr(rs!name)
    Dim uref As t_UserReference
    uref = NameIndex(charName)
    If IsValidUserRef(uref) Then
        Call modNetwork.Kick(UserList(uref.ArrayIndex).ConnectionDetails.ConnID, "Tu personaje fue publicado en el mercado.")
    End If

    ApplyMaoLock = "ok"
    Exit Function
ApplyMaoLock_Err:
    Call LogError("ApplyMaoLock error: " & Err.Description & " - char " & charId & " seller " & sellerAcc)
    ApplyMaoLock = "error;unknown_error"
End Function

' ---------------------------------------------------------------------------
' ApplyMaoUnlock  (PRIVATE)  -> "ok". Idempotente (0 filas = ya desbloqueado o
' no es del seller = ok).
' ---------------------------------------------------------------------------
Private Function ApplyMaoUnlock(ByVal charId As Long, ByVal sellerAcc As Long) As String
    On Error GoTo ApplyMaoUnlock_Err

    Call Query("UPDATE user SET is_locked_in_mao = 0, price_in_mao = 0 WHERE id = ? AND account_id = ?;", charId, sellerAcc)

    ApplyMaoUnlock = "ok"
    Exit Function
ApplyMaoUnlock_Err:
    Call LogError("ApplyMaoUnlock error: " & Err.Description & " - char " & charId & " seller " & sellerAcc)
    ApplyMaoUnlock = "error;unknown_error"
End Function

' ---------------------------------------------------------------------------
' ApplyMaoTransfer  (PRIVATE)  -> "ok" | "error;<detail>". Reparenta el pj al
' comprador de forma atomica y verificada. Valida en orden: existe / ya
' transferido (idempotente) / sigue bloqueado / es del vendedor / no online /
' cuenta destino existe-no borrada-no baneada / cupo / colision de nombre.
' NO toca clan (guild_index) ni matrimonio (spouse). saleTxId solo trazabilidad.
' ---------------------------------------------------------------------------
Private Function ApplyMaoTransfer(ByVal charId As Long, ByVal fromAcc As Long, ByVal toAcc As Long, ByVal saleTxId As String) As String
    On Error GoTo ApplyMaoTransfer_Err

    Dim rs As ADODB.Recordset
    Set rs = Query("SELECT name, account_id, is_locked_in_mao FROM user WHERE id = ?;", charId)
    If rs Is Nothing Then
        ApplyMaoTransfer = "error;unknown_error"
        Exit Function
    End If
    If rs.EOF Then
        ApplyMaoTransfer = "error;char_not_found"
        Exit Function
    End If

    If CLng(rs!account_id) = toAcc And Not CBool(rs!is_locked_in_mao) Then
        ApplyMaoTransfer = "ok"
        Exit Function
    End If
    If Not CBool(rs!is_locked_in_mao) Then
        ApplyMaoTransfer = "error;not_locked"
        Exit Function
    End If
    If CLng(rs!account_id) <> fromAcc Then
        ApplyMaoTransfer = "error;owner_mismatch"
        Exit Function
    End If

    Dim charName As String
    charName = CStr(rs!name)

    Dim uref As t_UserReference
    uref = NameIndex(charName)
    If IsValidUserRef(uref) Then
        ApplyMaoTransfer = "error;char_online"
        Exit Function
    End If

    Dim rsAcc As ADODB.Recordset
    Set rsAcc = Query("SELECT deleted, is_banned FROM account WHERE id = ?;", toAcc)
    If rsAcc Is Nothing Then
        ApplyMaoTransfer = "error;unknown_error"
        Exit Function
    End If
    If rsAcc.EOF Then
        ApplyMaoTransfer = "error;target_deleted"
        Exit Function
    End If
    If CBool(rsAcc!deleted) Then
        ApplyMaoTransfer = "error;target_deleted"
        Exit Function
    End If
    If CBool(rsAcc!is_banned) Then
        ApplyMaoTransfer = "error;target_banned"
        Exit Function
    End If

    Dim cap As Long
    cap = MaxCharacterForTier(GetPatronTierFromAccountID(toAcc))
    If cap > MAX_PERSONAJES Then cap = MAX_PERSONAJES
    If GetPersonajesCountByIDDatabase(toAcc) >= cap Then
        ApplyMaoTransfer = "error;target_over_cap"
        Exit Function
    End If

    Dim rsName As ADODB.Recordset
    Set rsName = Query("SELECT 1 FROM user WHERE account_id = ? AND name = ? COLLATE NOCASE AND id <> ?;", toAcc, charName, charId)
    If rsName Is Nothing Then
        ApplyMaoTransfer = "error;unknown_error"
        Exit Function
    End If
    If Not rsName.EOF Then
        ApplyMaoTransfer = "error;name_conflict"
        Exit Function
    End If

    Call Query("UPDATE user SET account_id = ?, is_locked_in_mao = 0, price_in_mao = 0 WHERE id = ? AND is_locked_in_mao = 1 AND account_id = ?;", toAcc, charId, fromAcc)

    Dim rsAfter As ADODB.Recordset
    Set rsAfter = Query("SELECT account_id, is_locked_in_mao FROM user WHERE id = ?;", charId)
    If rsAfter Is Nothing Then
        ApplyMaoTransfer = "error;unknown_error"
        Exit Function
    End If
    If rsAfter.EOF Then
        ApplyMaoTransfer = "error;unknown_error"
        Exit Function
    End If
    If CLng(rsAfter!account_id) = toAcc And Not CBool(rsAfter!is_locked_in_mao) Then
        ApplyMaoTransfer = "ok"
    Else
        ApplyMaoTransfer = "error;unknown_error"
    End If
    Exit Function
ApplyMaoTransfer_Err:
    Call LogError("ApplyMaoTransfer error: " & Err.Description & " - char " & charId & " from " & fromAcc & " to " & toAcc & " tx " & saleTxId)
    ApplyMaoTransfer = "error;unknown_error"
End Function

' ---------------------------------------------------------------------------
' AckMaoOps  (PRIVATE). Confirma en la web las ops procesadas. Body = una linea
' por op ("<op_id>;ok" | "<op_id>;error;<detail>") separadas por vbLf.
' Content-Type text/plain explicito (sin el, express.text no parsea). Fail-silent.
' ---------------------------------------------------------------------------
Private Sub AckMaoOps(ByVal Body As String)
    On Error GoTo AckMaoOps_Err

    Dim http As MSXML2.ServerXMLHTTP60
    Set http = New MSXML2.ServerXMLHTTP60
    http.setTimeouts 2000, 2000, 3000, 3000

    http.Open "POST", m_BaseUrl & "/api/bridge/mao-ops/ack", False
    http.setRequestHeader "X-Service-Token", m_ServiceToken
    http.setRequestHeader "Content-Type", "text/plain"
    http.send Body

    If http.Status <> 200 Then
        Call LogError("MaoBridge ack POST failed. Status: " & http.Status)
    End If

    Exit Sub
AckMaoOps_Err:
    Call LogError("AckMaoOps error: " & Err.Description)
End Sub
