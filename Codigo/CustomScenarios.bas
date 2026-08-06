Attribute VB_Name = "CustomScenarios"
' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
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
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit

Public Enum e_DamageSourceType
    e_phisical
    e_magic
    e_pet
    e_dot
    e_trap
End Enum

Public Enum e_ScenarioInterfaces
    eInventoryInterface = 1
End Enum

Public Type e_Rank
    PlayerIndex As Integer
    Score As Integer
End Type

Private CustomScenarioList   As New Dictionary
Private ScenarioUpdateList() As IBaseScenario
Private AvailableUpdateSlots As t_IndexHeap
Private ActiveUpdateSlots    As t_IndexHeap
Const InitialUpdateSize = 20

Private Sub InitializeUpdateStacks()
    ReDim ScenarioUpdateList(InitialUpdateSize) As IBaseScenario
    ReDim AvailableUpdateSlots.IndexInfo(InitialUpdateSize)
    ReDim ActiveUpdateSlots.IndexInfo(InitialUpdateSize)
    Dim i As Integer
    For i = 1 To InitialUpdateSize
        AvailableUpdateSlots.IndexInfo(i) = InitialUpdateSize - (i - 1)
    Next i
    AvailableUpdateSlots.currentIndex = InitialUpdateSize
    ActiveUpdateSlots.currentIndex = 0
End Sub

Private Sub IncreaseArraySize(ByVal ExtraSlots As Integer)
    Dim NewSize As Integer
    NewSize = UBound(ScenarioUpdateList) + ExtraSlots
    ReDim Preserve ScenarioUpdateList(NewSize) As IBaseScenario
    ReDim Preserve AvailableUpdateSlots.IndexInfo(NewSize)
    ReDim Preserve ActiveUpdateSlots.IndexInfo(NewSize)
    Dim i As Integer
    For i = 1 To ExtraSlots
        AvailableUpdateSlots.IndexInfo(i) = NewSize - (i - 1)
    Next i
    AvailableUpdateSlots.currentIndex = ExtraSlots
End Sub

Public Function GetMap(ByVal mapIndex As Integer) As IBaseScenario
    On Error GoTo GetMap_Err:
    Set GetMap = Nothing
    If CustomScenarioList.Exists(mapIndex) Then
        Set GetMap = CustomScenarioList.Item(mapIndex)
    End If
    Exit Function
GetMap_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.GetMap " & mapIndex, Erl)
End Function

Public Function SetCustomScenario(ByVal mapIndex As Integer, ByRef Scenario As IBaseScenario) As Boolean
    ' Plan 05.002 ola 5. Dictionary.Add TIRA error 457 si la clave ya existe, y esta Sub
    ' no tenia handler: el error escalaba al Class_Initialize / SetLobbyIndex del escenario
    ' que estaba registrando y le abortaba el resto de la inicializacion. En el CTF eso
    ' deja MapsRegistered = False y UpdateSlot = -1 (ScenarioCaptureTheFlag.cls:116-117),
    ' o sea un lobby abierto que cobra inscripcion y NUNCA tickea.
    ' Ahora falla suave y RUIDOSO, y devuelve si registro o no para que el llamador pueda
    ' abortar. Los 4 escenarios llaman con Call y siguen compilando sin tocarlos.
    On Error GoTo SetCustomScenario_Err:
    If CustomScenarioList.Exists(mapIndex) Then
        Call LogError("CustomScenarios.SetCustomScenario: el mapa " & mapIndex & " ya tiene un escenario registrado. No se pisa: el escenario nuevo se queda sin ese mapa.")
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("[Eventos] El mapa " & mapIndex & " ya tiene un escenario registrado: el evento nuevo no puede usarlo. Revisa el log.", e_FontTypeNames.FONTTYPE_INFO))
        Exit Function
    End If
    Call CustomScenarioList.Add(mapIndex, Scenario)
    SetCustomScenario = True
    Exit Function
SetCustomScenario_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.SetCustomScenario " & mapIndex, Erl)
End Function

Public Sub RemoveCustomScenario(ByVal mapIndex As Integer)
    ' Plan 05.002 ola 5. Dictionary.Remove TIRA error 32811 si la clave no existe, y aca
    ' tampoco habia handler. El dano no era el error: en ScenarioHuntNpc.ClearMap,
    ' ScenarioDeathMatch.ClearMap y ScenarioNavalBoarding.ClearMap esta llamada es la
    ' PRIMERA linea y RemoveUpdateScenario la ULTIMA, asi que el error saltaba al handler
    ' del ClearMap y el slot del tick no se liberaba nunca. Con el mapa jamas registrado
    ' (instancia agotada => MapNumber = -1) eso pasa siempre.
    On Error GoTo RemoveCustomScenario_Err:
    If Not CustomScenarioList.Exists(mapIndex) Then Exit Sub
    Call CustomScenarioList.Remove(mapIndex)
    Exit Sub
RemoveCustomScenario_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.RemoveCustomScenario " & mapIndex, Erl)
End Sub

Public Function AddUpdateScenario(ByRef Scenario As IBaseScenario) As Integer
    On Error GoTo AddUpdateScenario_Err:
    Dim pos As Integer
    If AvailableUpdateSlots.currentIndex = 0 And ActiveUpdateSlots.currentIndex = 0 Then
        Call InitializeUpdateStacks
    End If
    If AvailableUpdateSlots.currentIndex = 0 Then
        Call IncreaseArraySize(InitialUpdateSize)
    End If
    pos = AvailableUpdateSlots.IndexInfo(AvailableUpdateSlots.currentIndex)
    AvailableUpdateSlots.currentIndex = AvailableUpdateSlots.currentIndex - 1
    ActiveUpdateSlots.IndexInfo(ActiveUpdateSlots.currentIndex) = pos
    ActiveUpdateSlots.currentIndex = ActiveUpdateSlots.currentIndex + 1
    Set ScenarioUpdateList(pos) = Scenario
    AddUpdateScenario = pos
    Exit Function
AddUpdateScenario_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.AddUpdateScenario", Erl)
End Function

Public Sub RemoveUpdateScenario(ByVal Index As Integer)
    ' Plan 05.002 ola 5. Endurecido, NO reescrito: el barrido lineal ya hacia esta Sub
    ' idempotente (los pos son unicos y arrancan en 1, asi que un segundo Remove no
    ' encuentra la clave y no re-encola). La doble liberacion que se sospechaba no existe.
    ' Lo que faltaba:
    '  - handler propio: Debug.Assert se compila AFUERA en release, asi que un Index fuera
    '    de rango salia como error 9 sin atrapar y le abortaba el cierre al escenario.
    '  - Exit For: sin el, el barrido sigue leyendo posiciones que ya quedaron atras.
    '  - guarda de overflow al devolver el slot, la misma que InstanceManager.ReleaseInstance
    '    tuvo que aprender a poner (ver el comentario de InstanceManager.bas:42-44).
    On Error GoTo RemoveUpdateScenario_Err:
    If Index < LBound(ScenarioUpdateList) Or Index > UBound(ScenarioUpdateList) Then
        Call LogError("CustomScenarios.RemoveUpdateScenario: slot fuera de rango (" & Index & ")")
        Exit Sub
    End If
    Set ScenarioUpdateList(Index) = Nothing
    Dim i As Integer
    For i = 0 To ActiveUpdateSlots.currentIndex - 1
        If ActiveUpdateSlots.IndexInfo(i) = Index Then
            ActiveUpdateSlots.IndexInfo(i) = ActiveUpdateSlots.IndexInfo(ActiveUpdateSlots.currentIndex - 1)
            ActiveUpdateSlots.currentIndex = ActiveUpdateSlots.currentIndex - 1
            If AvailableUpdateSlots.currentIndex < UBound(AvailableUpdateSlots.IndexInfo) Then
                AvailableUpdateSlots.currentIndex = AvailableUpdateSlots.currentIndex + 1
                AvailableUpdateSlots.IndexInfo(AvailableUpdateSlots.currentIndex) = Index
            Else
                Call LogError("CustomScenarios.RemoveUpdateScenario: el heap de slots disponibles esta lleno, se descarta el slot " & Index)
            End If
            Exit For
        End If
    Next
    Exit Sub
RemoveUpdateScenario_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.RemoveUpdateScenario " & Index, Erl)
End Sub

Public Sub UpdateAll()
    ' Plan 05.002 ola 5. DOS problemas vivian en el mismo For:
    ' 1) Un solo On Error para todo el loop. El primer escenario que tiraba abortaba el
    '    tick de TODOS los siguientes, y como el For siempre arranca en 0 los starveaba
    '    para siempre (UpdateAll entra cada 40 ms, modGameEvents.MaybeRunGameEvents).
    ' 2) Se iteraba ActiveUpdateSlots MIENTRAS los escenarios lo mutan: los 4 llaman
    '    CloseScenario -> RemoveUpdateScenario desde su propio Update, y RemoveUpdateScenario
    '    compacta el heap con swap-contra-el-ultimo. VB6 evalua el limite del For una sola
    '    vez, asi que la posicion que quedo obsoleta se releia y el escenario que se movio
    '    recibia dos Update en el mismo tick.
    ' Por eso: se copia la lista de slots ANTES de recorrerla, y cada Update va en su
    ' propia Sub, que es la unica forma en VB6 de tener un handler por iteracion sin
    ' recurrir a Resume Next (que aca es trampa: si el error cae en la linea del For,
    ' Resume Next reentra al cuerpo del loop sin inicializar).
    On Error GoTo UpdateAll_Err:
    Dim i         As Integer
    Dim SlotCount As Integer
    Dim Slots()   As Integer
    SlotCount = ActiveUpdateSlots.currentIndex
    If SlotCount <= 0 Then Exit Sub
    ReDim Slots(0 To SlotCount - 1) As Integer
    For i = 0 To SlotCount - 1
        Slots(i) = ActiveUpdateSlots.IndexInfo(i)
    Next i
    For i = 0 To SlotCount - 1
        Call UpdateSingleScenario(Slots(i))
    Next i
    Exit Sub
UpdateAll_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.UpdateAll", Erl)
End Sub

Private Sub UpdateSingleScenario(ByVal Slot As Integer)
    ' El chequeo Is Nothing NO es defensivo de mas: es lo que cubre al escenario que se
    ' dio de baja durante este mismo tick, porque RemoveUpdateScenario deja el slot en
    ' Nothing antes de sacarlo del heap.
    On Error GoTo UpdateSingleScenario_Err:
    If Slot < LBound(ScenarioUpdateList) Or Slot > UBound(ScenarioUpdateList) Then Exit Sub
    If ScenarioUpdateList(Slot) Is Nothing Then Exit Sub
    Call ScenarioUpdateList(Slot).Update
    Exit Sub
UpdateSingleScenario_Err:
    ' NO se traga el error: se loguea con el slot para poder identificar al escenario.
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.UpdateAll slot:" & Slot, Erl)
End Sub

Public Sub Reset(ByVal MapNumber As Integer)
    On Error GoTo Reset_Err:
    Dim Scenario As IBaseScenario
    Set Scenario = GetMap(MapNumber)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    Call Scenario.Reset
    Exit Sub
Reset_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.Reset " & MapNumber, Erl)
End Sub

Public Sub PlayerKillNpc(ByVal MapNumber As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal SourceType As e_DamageSourceType, ByVal SourceIndex As Integer)
    On Error GoTo PlayerKillNpc_Err:
    Dim Scenario As IBaseScenario
    Set Scenario = GetMap(MapNumber)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    Call Scenario.PlayerKillNpc(NpcIndex, UserIndex, SourceType, SourceIndex)
    Exit Sub
PlayerKillNpc_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.PlayerKillNpc map:" & MapNumber & " npc: " & NpcIndex & " user: " & UserIndex, Erl)
End Sub

Public Sub NPcKillPlayer(ByVal MapNumber As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal SourceType As e_DamageSourceType, ByVal SourceIndex As Integer)
    On Error GoTo NPcKillPlayer_Err:
    Dim Scenario As IBaseScenario
    Set Scenario = GetMap(MapNumber)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    Call Scenario.NPcKillPlayer(NpcIndex, UserIndex, SourceType, SourceIndex)
    Exit Sub
NPcKillPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.NPcKillPlayer", Erl)
End Sub

Public Sub PlayerKillPlayer(ByVal MapNumber As Integer, _
                            ByVal killerIndex As Integer, _
                            ByVal deadIndex As Integer, _
                            ByVal SourceType As e_DamageSourceType, _
                            ByVal SourceIndex As Integer)
    On Error GoTo PlayerKillPlayer_Err:
    Dim Scenario As IBaseScenario
    Set Scenario = GetMap(MapNumber)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    Call Scenario.PlayerKillPlayer(killerIndex, deadIndex, SourceType, SourceIndex)
    Exit Sub
PlayerKillPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.PlayerKillPlayer", Erl)
End Sub

Public Sub UserDie(ByVal UserIndex As Integer)
    On Error GoTo UserDie_Err:
    Dim Scenario As IBaseScenario
    Set Scenario = GetMap(UserList(UserIndex).pos.Map)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    Call Scenario.UserDie(UserIndex)
    Exit Sub
UserDie_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.UserDie", Erl)
End Sub

Public Sub NpcDie(ByVal NpcIndex As Integer)
    On Error GoTo NpcDie_Err:
    Dim Scenario As IBaseScenario
    Set Scenario = GetMap(NpcList(NpcIndex).pos.Map)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    Call Scenario.NpcDie(NpcIndex)
    Exit Sub
NpcDie_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.NpcDie", Erl)
End Sub

'inform that user want to pickup an item, return false if user can't pick up item at pos
Public Function UserCanPickUpItem(ByVal UserIndex As Integer) As Boolean
    On Error GoTo UserPickUpItem_Err:
    UserCanPickUpItem = True
    Dim Scenario           As IBaseScenario
    Dim InventoryInterface As IInventoryInterface
    Set Scenario = GetMap(UserList(UserIndex).pos.Map)
    If Scenario Is Nothing Then
        Exit Function
    End If
    If Not IsSet(Scenario.GetExtraInterfaces(), eInventoryInterface) Then
        Exit Function
    End If
    Set InventoryInterface = Scenario
    UserCanPickUpItem = InventoryInterface.UserCanPickUpItem(UserIndex)
    Exit Function
UserPickUpItem_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.UserCanPickUpItem", Erl)
End Function

Public Sub UserDidPickupItem(ByVal UserIndex As Integer, ByVal ItemId As Integer)
    On Error GoTo UserDidPickupItem_Err:
    Dim Scenario           As IBaseScenario
    Dim InventoryInterface As IInventoryInterface
    Set Scenario = GetMap(UserList(UserIndex).pos.Map)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    If Not IsSet(Scenario.GetExtraInterfaces(), eInventoryInterface) Then
        Exit Sub
    End If
    Set InventoryInterface = Scenario
    Call InventoryInterface.UserDidPickupItem(UserIndex, ItemId)
    Exit Sub
UserDidPickupItem_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.UserDidPickupItem", Erl)
End Sub

'inform that user want to drop an item
Public Sub UserDropItem(ByVal UserIndex As Integer, ByVal Slot As Integer, ByVal Map As Integer, ByVal TileX As Integer, ByVal TileY As Integer)
    On Error GoTo UserDropItem_Err:
    Dim Scenario           As IBaseScenario
    Dim InventoryInterface As IInventoryInterface
    Set Scenario = GetMap(Map)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    If Not IsSet(Scenario.GetExtraInterfaces(), eInventoryInterface) Then
        Exit Sub
    End If
    Set InventoryInterface = Scenario
    Call InventoryInterface.UserDropItem(UserIndex, Slot, Map, TileX, TileY)
    Exit Sub
UserDropItem_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.UserDropItem", Erl)
End Sub

Public Function UserCanDropItem(ByVal UserIndex As Integer, ByVal Slot As Integer, ByVal Map As Integer, ByVal TileX As Integer, ByVal TileY As Integer) As Boolean
    On Error GoTo UserCanDropItem_Err:
    UserCanDropItem = True
    Dim Scenario           As IBaseScenario
    Dim InventoryInterface As IInventoryInterface
    Set Scenario = GetMap(Map)
    If Scenario Is Nothing Then
        Exit Function
    End If
    If Not IsSet(Scenario.GetExtraInterfaces(), eInventoryInterface) Then
        Exit Function
    End If
    Set InventoryInterface = Scenario
    UserCanDropItem = InventoryInterface.UserCanDropItemAt(UserIndex, Slot, Map, TileX, TileY)
    Exit Function
UserCanDropItem_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.UserCanDropItem", Erl)
End Function

Public Function PrepareNewEvent(ByVal eventType As e_EventType, ByVal LobbyIndex As Integer) As Boolean
    ' Plan 05.002 ola 5. Ahora devuelve True solo si el lobby quedo con un escenario
    ' VIVO: construido Y enganchado al tick. Antes era una Sub que no verificaba nada.
    '  - e_EventType.Generic (0) nunca tuvo Case. El lobby quedaba con Scenario = Nothing
    '    y, como UpdateWaitingForPlayers SOLO se llama desde el Update de los escenarios
    '    (ver los 4 IBaseScenario_Update), ese lobby no arranca al llenarse, no vence por
    '    tiempo y no se cancela solo: se queda en AcceptingPlayers cobrando inscripciones
    '    y ActiveEventType() lo ve activo, bloqueando todo evento publico posterior.
    '  - Los Class_Initialize de los 4 escenarios atrapan sus propios errores, asi que New
    '    SIEMPRE devuelve objeto, este bien construido o no. Un escenario que fallo antes
    '    de AddUpdateScenario no tickea nunca: mismo lobby zombi por otra puerta.
    ' Por eso se verifica la POSTCONDICION (esta en el tick) en vez de confiar en el New.
    On Error GoTo PrepareNewEvent_Err:
    Debug.Assert LobbyIndex >= LBound(LobbyList) And LobbyIndex <= UBound(LobbyList)
    Set LobbyList(LobbyIndex).Scenario = Nothing
    Select Case eventType
        Case e_EventType.NpcHunt
            Set LobbyList(LobbyIndex).Scenario = New ScenarioHunt
        Case e_EventType.DeathMatch
            Set LobbyList(LobbyIndex).Scenario = New ScenarioDeathMatch
        Case e_EventType.NavalBattle
            Set LobbyList(LobbyIndex).Scenario = New ScenarioNavalBoarding
        Case e_EventType.CaptureTheFlag
            ' Plan 05.002 ola 4: el CTF paso a ser un escenario del motor.
            Set LobbyList(LobbyIndex).Scenario = New ScenarioCaptureTheFlag
        Case e_EventType.Generic
            ' Sin escenario A PROPOSITO: es el evento MANUAL del GM (/CREAREVENTO LOBBY), que
            ' el staff arranca y termina a mano. Lo hace avanzar ModLobby.UpdateLobbiesWithoutScenario,
            ' agregada en esta misma ola. Devuelve True: no hay escenario que verificar.
            PrepareNewEvent = True
            Exit Function
        Case Else
            Call LogError("CustomScenarios.PrepareNewEvent: tipo de evento desconocido (" & eventType & ") para el lobby " & LobbyIndex & ".")
            Exit Function
    End Select
    If LobbyList(LobbyIndex).Scenario Is Nothing Then Exit Function
    LobbyList(LobbyIndex).Scenario.SetLobbyIndex (LobbyIndex)
    ' Postcondicion: AddUpdateScenario lo llaman Class_Initialize (Hunt, DeathMatch,
    ' Naval) o SetLobbyIndex (CTF). Si cualquiera de los dos trago un error (clave
    ' duplicada en SetCustomScenario, instancia de mapa agotada, .ini faltante) el objeto
    ' existe pero NO esta en el tick. Sin esta guarda, eso es un lobby zombi.
    If Not IsScenarioInUpdateList(LobbyList(LobbyIndex).Scenario) Then
        Call LogError("CustomScenarios.PrepareNewEvent: el escenario tipo " & eventType & " del lobby " & LobbyIndex & " no quedo registrado en el tick. Se descarta.")
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("[Eventos] El escenario no pudo inicializarse (ver log). El evento no se abre.", e_FontTypeNames.FONTTYPE_INFO))
        Set LobbyList(LobbyIndex).Scenario = Nothing
        Exit Function
    End If
    PrepareNewEvent = True
    Exit Function
PrepareNewEvent_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.PrepareNewEvent", Erl)
End Function

Public Function IsScenarioInUpdateList(ByRef Scenario As IBaseScenario) As Boolean
    ' Los dos lados son IBaseScenario, asi que Is compara el mismo puntero de interfaz.
    On Error GoTo IsScenarioInUpdateList_Err:
    Dim i As Integer
    If Scenario Is Nothing Then Exit Function
    For i = 0 To ActiveUpdateSlots.currentIndex - 1
        If ScenarioUpdateList(ActiveUpdateSlots.IndexInfo(i)) Is Scenario Then
            IsScenarioInUpdateList = True
            Exit Function
        End If
    Next i
    Exit Function
IsScenarioInUpdateList_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.IsScenarioInUpdateList", Erl)
End Function

Public Sub ClearMap(ByVal MapNumber As Integer)
    On Error GoTo ClearMap_Err:
    Dim x As Long
    Dim y As Long
    For y = 1 To 99
        For x = 1 To 99
            If MapData(MapNumber, x, y).ObjInfo.ObjIndex > 0 Then
                If ItemNoEsDeMapa(MapData(MapNumber, x, y).ObjInfo.ObjIndex) Then
                    Call EraseObj(GetMaxInvOBJ(), MapNumber, x, y)
                End If
            End If
        Next x
    Next y
    Exit Sub
ClearMap_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.ClearMap", Erl)
    Resume Next
End Sub

Public Function IsEventActive() As Boolean
    ' Plan 05.002 ola 3: una sola fuente de verdad, derivada de lo que esta vivo.
    IsEventActive = (ModLobby.ActiveEventType() <> ModLobby.NO_ACTIVE_EVENT)
End Function

' Plan 05.002 ola 7: "el mapa donde esta parado este jugador pertenece a un escenario".
' Deliberadamente NO usa numeros de mapa: EsMapaEvento() es una lista estatica de
' MapasEspeciales.dat y los eventos corren en instancias con numero asignado en runtime, asi
' que un gate basado en el queda ciego para 3 de los 4 modos. Tampoco usa MapInfo(...).Seguro:
' los escenarios lo prenden y apagan por fase (1 -> 0 -> 1) y el gate parpadearia con la fase.
' El registro por mapa es como el motor ya rutea muerte, banderas y NPCs.
Public Function IsUserInScenarioMap(ByVal UserIndex As Integer) As Boolean
    On Error GoTo IsUserInScenarioMap_Err
    IsUserInScenarioMap = False
    If UserIndex <= 0 Then Exit Function
    IsUserInScenarioMap = Not (GetMap(UserList(UserIndex).pos.Map) Is Nothing)
    Exit Function
IsUserInScenarioMap_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.IsUserInScenarioMap", Erl)
    IsUserInScenarioMap = False
End Function

' Reto = flag por usuario (el reto no vive en un escenario). Evento = el mapa donde esta
' parado. El handler devuelve False a proposito: esto es una restriccion de juego, no un
' limite de seguridad. Si algo revienta se prefiere dejar trabajar o comerciar de mas antes
' que dejar sin profesiones a todo el servidor. El error queda logueado igual.
Public Function IsUserInMatch(ByVal UserIndex As Integer) As Boolean
    On Error GoTo IsUserInMatch_Err
    IsUserInMatch = False
    If UserIndex <= 0 Then Exit Function
    If UserList(UserIndex).flags.EnReto Then
        IsUserInMatch = True
        Exit Function
    End If
    IsUserInMatch = IsUserInScenarioMap(UserIndex)
    Exit Function
IsUserInMatch_Err:
    Call TraceError(Err.Number, Err.Description, "CustomScenarios.IsUserInMatch", Erl)
    IsUserInMatch = False
End Function

Public Sub UserDisconnected(ByVal UserIndex As Integer)
    Call RegisterDisconnectedUser(UserIndex)
End Sub

Public Sub UserConnected(ByVal UserIndex As Integer)
    Call RegisterReconnectedUser(UserIndex)
End Sub

Public Sub GetNextWaypointForNpc(ByVal NpcIndex As Integer, ByRef PosX As Integer, ByRef PosY As Integer)
    Dim Scenario As IBaseScenario
    Set Scenario = GetMap(NpcList(NpcIndex).pos.Map)
    If Scenario Is Nothing Then
        Exit Sub
    End If
    Call Scenario.GetNextWaypointForNpc(NpcIndex, PosX, PosY)
    Exit Sub
End Sub
