VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScenarioCaptureTheFlag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' ============================================================================
' CAPTURA DE LA BANDERA como escenario del motor (plan 05.002, ola 4).
'
' Reemplaza a clsCaptura.cls, que corria por fuera del sistema de lobby con su
' propio singleton (InstanciaCaptura), su propia coleccion de participantes, su
' propio cobro/refund y su propia maquina de estados. De ahi salian 3 de los 5
' bugs criticos de la auditoria: el For Each que se salteaba la mitad de los
' jugadores al cerrar, el Exit Sub faltante al cancelar por cupo impar, y el
' drop de inventario que ningun otro modo tenia.
'
' QUE CAMBIA: nada de cara al jugador (decision del dueno 2026-08-05). Mismos
' mensajes, misma sala de espera, mismas rondas, mismas validaciones al anotarse.
' QUE CAMBIA POR DEBAJO: participantes, cobro, refund, estado y desconexiones los
' maneja el lobby; aca quedan banderas, rondas y bases.
'
' OJO CON EL TIEMPO. Update() lo llama modGameEvents.MaybeRunGameEvents cada
' GAME_EVENTS_INTERVAL_MS = 40 ms, o sea 25 veces por segundo. clsCaptura corria
' 1 vez por segundo. Toda la logica por segundo vive dentro del
' If UpdateTime(SecondTimer, frametime), que es el patron que ya usa
' ScenarioDeathMatch. Sacar eso de ahi acelera el modo 25x.
' ============================================================================
Option Explicit
Implements IBaseScenario

Private Enum e_CtfRoundState
    EsperandoInicio = 0
    IniciaRonda = 1
    EnJuego = 2
    Terminado = 3
End Enum

Private LobbyIndex              As Integer
Private UpdateSlot              As Integer
Private LastFrameTime           As Long
Private SecondTimer             As t_Timer
Private RoundState              As e_CtfRoundState
Private cantidad_rondas         As Byte
Private ronda_actual            As Byte
Private tiempo_ronda            As Integer
Private contadorBandera(1 To 2) As Integer
Private rondasGanadas(1 To 2)   As Integer
Private colorEquipo(1 To 2)     As String
Private mapaEquipo(1 To 2)      As Integer
Private xEquipo(1 To 2)         As Byte
Private yEquipo(1 To 2)         As Byte
Private objEquipo(1 To 2)       As Integer
Private MapsRegistered          As Boolean
' B3: la muerte NO se puede resolver adentro del hook. CustomScenarios.UserDie corre
' ANTES de UserMod.UserDie en todos los call sites (Modulo_UsUaRiOs:3055 vs :860), asi
' que ahi el jugador todavia tiene Muerto = 0: ResucitarOCurar no revive nada y el warp
' a la sala hace que TirarTodosLosItems evalue el DropItems de la sala de espera. Se
' anota el muerto y se lo procesa en el Update siguiente, cuando la muerte ya es real.
' Es el mismo patron que usa ScenarioDeathMatch con su clsRespawnInfo.
Private Const MAX_MUERTES_PENDIENTES         As Integer = 64
Private MuertePendiente(1 To MAX_MUERTES_PENDIENTES) As t_UserReference
Private MuerteIntentos(1 To MAX_MUERTES_PENDIENTES)  As Integer
Private MuertesPendientes                    As Integer
' Plan 05.002 ola 5: watchdog de evento vacio. El CTF es el unico modo sin MaxTime, asi
' que necesita su propio criterio de terminacion cuando no queda nadie. Ver Update.
Private Const CAPTURA_SEGUNDOS_SIN_JUGADORES As Integer = 120
Private SegundosSinJugadores                 As Integer

Private Sub Class_Initialize()
    colorEquipo(1) = "azul"
    colorEquipo(2) = "rojo"
    mapaEquipo(1) = MAP_TEAM_1
    mapaEquipo(2) = MAP_TEAM_2
    xEquipo(1) = X_BANDERA_1
    xEquipo(2) = X_BANDERA_2
    yEquipo(1) = Y_BANDERA_1
    yEquipo(2) = Y_BANDERA_2
    objEquipo(1) = OBJ_CAPTURA_BANDERA_1
    objEquipo(2) = OBJ_CAPTURA_BANDERA_2
    UpdateSlot = -1
    RoundState = EsperandoInicio
End Sub

' ============================ IBaseScenario ============================

Private Sub IBaseScenario_SetLobbyIndex(ByVal value As Integer)
    On Error GoTo ErrHandler
    LobbyIndex = value
    cantidad_rondas = LobbyList(LobbyIndex).RoundNumber
    If cantidad_rondas < 1 Then cantidad_rondas = 1
    ' Los 3 mapas del CTF son fijos: por ahi llegan muerte y banderas via
    ' CustomScenarios.GetMap(pos.Map).
    Call CustomScenarios.SetCustomScenario(MAP_TEAM_1, Me)
    Call CustomScenarios.SetCustomScenario(MAP_TEAM_2, Me)
    Call CustomScenarios.SetCustomScenario(MAP_SALA_ESPERA, Me)
    MapsRegistered = True
    UpdateSlot = CustomScenarios.AddUpdateScenario(Me)
    LastFrameTime = GetTickCountRaw()
    Call SetTimer(SecondTimer, 1000)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.SetLobbyIndex", Erl)
End Sub

Private Function IBaseScenario_GetScenarioName() As String
    IBaseScenario_GetScenarioName = "Captura de la Bandera"
End Function

Private Sub IBaseScenario_BroadcastOpenScenario()
    On Error GoTo ErrHandler
    'Msg1688: mismo texto que anunciaba clsCaptura.inicializar.
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_INICIADO_CAPTURA_BANDERA_PARTICIPANTES_PRECIO_INSCRIPCION, LobbyList(LobbyIndex).MaxPlayers & "¬" & PonerPuntos(LobbyList(LobbyIndex).InscriptionPrice) & "¬" & LobbyList(LobbyIndex).MinLevel & "¬" & LobbyList(LobbyIndex).MaxLevel, e_FontTypeNames.FONTTYPE_GUILD))
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.BroadcastOpenScenario", Erl)
End Sub

Private Sub IBaseScenario_SendRules(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Call WriteConsoleMsg(UserIndex, "Captura de la Bandera: llevate la bandera enemiga a tu base y aguantala " & CAPTURA_TIEMPO_BANDERA & " segundos para ganar la ronda.", e_FontTypeNames.FONTTYPE_GUILD)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.SendRules", Erl)
End Sub

Private Function IBaseScenario_GetExtraInterfaces() As Long
    IBaseScenario_GetExtraInterfaces = 0
End Function

Private Function IBaseScenario_RequiresSpawn() As Boolean
    ' El CTF tiene mapas fijos: no necesita que el GM fije un punto de spawn.
    IBaseScenario_RequiresSpawn = False
End Function

Private Function IBaseScenario_ValidateUser(ByVal UserIndex As Integer) As Integer
    ' Contrato del motor: devolver el ID DEL MENSAJE de rechazo (>0), o 0 si puede
    ' entrar. El mensaje lo manda el lobby (ModLobby.CanPlayerJoin), no esta Sub.
    ' El lobby ya valida oro, carcel por trigger, nivel y cupos: aca van solo las
    ' condiciones propias del CTF, con los mismos mensajes que usaba clsCaptura.
    On Error GoTo ErrHandler
    IBaseScenario_ValidateUser = 0
    With UserList(UserIndex)
        If MapInfo(.pos.Map).Seguro = 0 Then
            'Msg928
            IBaseScenario_ValidateUser = MSG_SOLAMENTE_PODRAS_INGRESAR_EVENTO_ZONA_SEGURA
        ElseIf .Counters.Pena > 0 Then
            'Msg930
            IBaseScenario_ValidateUser = MSG_NO_PUEDES_ANOTARTE_EVENTO_MIENTRAS_ENCARCELADO
        ElseIf .flags.Muerto = 1 Then
            'Msg931
            IBaseScenario_ValidateUser = MSG_NO_PUEDES_ANOTARTE_EVENTO_ESTANDO_MUERTO
        ElseIf .flags.Comerciando Then
            'Msg932
            IBaseScenario_ValidateUser = MSG_NO_PUEDES_INGRESAR_EVENTO_MIENTRAS_COMERCIANDO
        ElseIf .flags.Montado = 1 Then
            'Msg933
            IBaseScenario_ValidateUser = MSG_NO_PUEDES_INGRESAR_EVENTO_ESTANDO_MONTADO
        ElseIf .flags.Navegando = 1 Then
            'Msg934
            IBaseScenario_ValidateUser = MSG_NO_PUEDES_INGRESAR_EVENTO_MIENTRAS_NAVEGANDO
        End If
    End With
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.ValidateUser", Erl)
    IBaseScenario_ValidateUser = 0
End Function

Private Sub IBaseScenario_WillSummonPlayer(ByVal UserIndex As Integer)
    ' Al invocarlo, el jugador va a la sala de espera, igual que antes.
    On Error GoTo ErrHandler
    Dim SpawnPos As t_WorldPos
    SpawnPos.Map = MAP_SALA_ESPERA
    SpawnPos.x = RandomNumber(MIN_SALA_ESPERA_X, MAX_SALA_ESPERA_X)
    SpawnPos.y = RandomNumber(MIN_SALA_ESPERA_Y, MAX_SALA_ESPERA_Y)
    LobbyList(LobbyIndex).SummonCoordinates = SpawnPos
    UserList(UserIndex).flags.jugando_captura = 1
    UserList(UserIndex).flags.tiene_bandera = 0
    ' Ojo: aca NO se resetean jugando_captura_timer ni jugando_captura_muertes.
    ' WillSummonPlayer tambien corre al reconectar (RegisterReconnectedUserOnLobby ->
    ' SummonPlayer), y resetearlos ahi convertia el relog en un borrado gratis del
    ' castigo por muerte acumulado. El reset por ronda vive en IniciarRondaNueva.
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.WillSummonPlayer", Erl)
End Sub

Private Sub IBaseScenario_UpdateLobbyState(ByVal oldState As e_LobbyState, ByRef NewState As e_LobbyState)
    On Error GoTo ErrHandler
    If oldState = NewState Then Exit Sub
    If NewState = e_LobbyState.InProgress Then
        Call ArrancarEvento
    ElseIf NewState = e_LobbyState.Closed Then
        Call LimpiarEscenario
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.UpdateLobbyState", Erl)
End Sub

Private Sub IBaseScenario_Update()
    On Error GoTo ErrHandler
    Dim nowRaw    As Long
    Dim frametime As Long
    nowRaw = GetTickCountRaw()
    frametime = CLng(TicksElapsed(LastFrameTime, nowRaw))
    LastFrameTime = nowRaw
    ' Cola de muertos: a ritmo del motor (40 ms) para que el warp a la sala se sienta
    ' inmediato, como en el legacy. Ver EncolarMuerte.
    If MuertesPendientes > 0 Then Call ProcesarMuertesPendientes
    ' B1/B4: el escenario es el que hace avanzar su lobby. Sin esta rama el CTF se
    ' queda para siempre en AcceptingPlayers (no arranca al llenarse ni al vencer el
    ' tiempo) y no cierra nunca al terminar. Es lo que hacen los otros 3 escenarios.
    If LobbyList(LobbyIndex).State = e_LobbyState.AcceptingPlayers Then
        Call UpdateWaitingForPlayers(frametime, LobbyList(LobbyIndex))
        Exit Sub
    ElseIf LobbyList(LobbyIndex).State = e_LobbyState.Closed Then
        Call CloseScenario
        Exit Sub
    End If
    ' TODO lo de abajo corre 1 vez por segundo. Update() entra 25 veces por
    ' segundo: sacar logica de aca adentro acelera el modo 25x.
    If Not UpdateTime(SecondTimer, frametime) Then Exit Sub
    ' Plan 05.002 ola 5: watchdog de terminacion. El CTF es el UNICO modo sin reloj:
    ' DeathMatch, Caceria y Abordaje cierran solos al vencer MaxTime, pero aca la ronda
    ' termina SOLO si alguien aguanta la bandera. Si se cae el ultimo jugador, TickEnJuego
    ' itera sobre referencias invalidas y no pasa nada nunca: el lobby queda en InProgress
    ' para siempre, el slot no vuelve al pool y ActiveEventType() sigue diciendo que hay un
    ' evento publico, asi que no se puede abrir ningun otro hasta reiniciar el server.
    ' El guard NO mira RoundState mas alla de Terminado a proposito: si ArrancarEvento se
    ' comio un error, RoundState se queda en EsperandoInicio con el lobby ya en InProgress,
    ' que es exactamente uno de los casos colgados que esto viene a matar.
    If RoundState <> Terminado And LobbyList(LobbyIndex).State = e_LobbyState.InProgress Then
        If JugadoresConectados() = 0 Then
            SegundosSinJugadores = SegundosSinJugadores + 1
            If SegundosSinJugadores >= CAPTURA_SEGUNDOS_SIN_JUGADORES Then
                Call LogError("[Evento] CTF cerrado por watchdog: " & CAPTURA_SEGUNDOS_SIN_JUGADORES & "s sin jugadores conectados.")
                ' CancelLobby refunda la inscripcion (tambien a los offline, via
                ' GiveGoldToPlayer) y deja el lobby en Closed. El cierre real -mapas y slot-
                ' lo hace CloseScenario en el tick siguiente, por la rama Closed de arriba.
                Call ModLobby.CancelLobby(LobbyList(LobbyIndex))
                Exit Sub
            End If
        Else
            SegundosSinJugadores = 0
        End If
    End If
    If RoundState = IniciaRonda Then
        Call TickCuentaRegresivaRonda
    ElseIf RoundState = EnJuego Then
        Call TickEnJuego
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.Update", Erl)
End Sub

Private Sub IBaseScenario_UserDie(ByVal UserIndex)
    ' Aca todavia NO esta muerto (ver el comentario de MAX_MUERTES_PENDIENTES). Lo que
    ' no depende de Muerto se hace ahora; el warp y la resurreccion se encolan.
    On Error GoTo ErrHandler
    If RoundState <> EnJuego Then Exit Sub
    With UserList(UserIndex)
        If .flags.CurrentTeam <= 0 Then Exit Sub
        'Msg1715
        Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_EQUIPO_MUERTO_BATALLA, .name & "¬" & colorEquipo(.flags.CurrentTeam), e_FontTypeNames.FONTTYPE_GUILD))
        Call TirarBanderaDe(UserIndex)
        .flags.jugando_captura_muertes = .flags.jugando_captura_muertes + 1
        .flags.jugando_captura_timer = .flags.jugando_captura_muertes * CAPTURA_TIEMPO_POR_MUERTE + CAPTURA_TIEMPO_MUERTE
        'Msg1846
        Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(MSG_VOLVERAS_BASE_SEGUNDOS, .flags.jugando_captura_timer, e_FontTypeNames.FONTTYPE_GUILD))
    End With
    Call EncolarMuerte(UserIndex)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.UserDie", Erl)
End Sub

Private Function JugadoresConectados() As Integer
    ' Cuenta inscriptos con sesion viva. NO alcanza con RegisteredPlayers: la desconexion no
    ' saca al jugador del lobby (RegisterDisconnectedUserOnLobby solo marca Connected = False
    ' y lo devuelve a su posicion), asi que RegisteredPlayers sigue contando gente que ya no
    ' esta conectada.
    On Error GoTo ErrHandler
    Dim i As Integer
    For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
        If LobbyList(LobbyIndex).Players(i).Connected Then
            If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
                JugadoresConectados = JugadoresConectados + 1
            End If
        End If
    Next i
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.JugadoresConectados", Erl)
End Function

Private Sub EncolarMuerte(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Dim i As Integer
    For i = 1 To MuertesPendientes
        If MuertePendiente(i).ArrayIndex = UserIndex Then Exit Sub
    Next i
    If MuertesPendientes >= MAX_MUERTES_PENDIENTES Then
        ' No deberia pasar (cupo maximo del CTF = 48 y la cola se vacia cada 40 ms),
        ' pero si pasa se resuelve en el acto antes que perder al jugador.
        Call DevolverALaSalaMuerto(UserIndex)
        Exit Sub
    End If
    MuertesPendientes = MuertesPendientes + 1
    Call SetUserRef(MuertePendiente(MuertesPendientes), UserIndex)
    MuerteIntentos(MuertesPendientes) = 0
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.EncolarMuerte", Erl)
End Sub

Private Sub ProcesarMuertesPendientes()
    On Error GoTo ErrHandler
    Dim i  As Integer
    Dim ui As Integer
    ' Hacia atras: la cola se compacta al sacar elementos.
    For i = MuertesPendientes To 1 Step -1
        MuerteIntentos(i) = MuerteIntentos(i) + 1
        If Not IsValidUserRef(MuertePendiente(i)) Then
            Call SacarDeLaCola(i)
        Else
            ui = MuertePendiente(i).ArrayIndex
            If UserList(ui).flags.Muerto = 1 Then
                Call DevolverALaSalaMuerto(ui)
                Call SacarDeLaCola(i)
            ElseIf MuerteIntentos(i) >= 25 Then
                ' Un segundo esperando el Muerto = 1 que nunca llego: lo revivio otra
                ' cosa (un GM, un sacerdote). Se lo suelta sin tocarlo.
                Call SacarDeLaCola(i)
            End If
        End If
    Next i
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.ProcesarMuertesPendientes", Erl)
End Sub

Private Sub SacarDeLaCola(ByVal Index As Integer)
    If Index < MuertesPendientes Then
        MuertePendiente(Index) = MuertePendiente(MuertesPendientes)
        MuerteIntentos(Index) = MuerteIntentos(MuertesPendientes)
    End If
    Call ClearUserRef(MuertePendiente(MuertesPendientes))
    MuerteIntentos(MuertesPendientes) = 0
    MuertesPendientes = MuertesPendientes - 1
End Sub

Private Sub DevolverALaSalaMuerto(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Call EnviarUsuarioSala(UserIndex)
    Call ResucitarOCurar(UserIndex)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.DevolverALaSalaMuerto", Erl)
End Sub

Private Sub IBaseScenario_OnUserDisconnected(ByVal UserIndex As Integer)
    ' El lobby ya marca el slot como desconectado y devuelve al jugador; aca solo
    ' se suelta la bandera para que no quede perdida en el limbo.
    On Error GoTo ErrHandler
    Call TirarBanderaDe(UserIndex)
    UserList(UserIndex).flags.jugando_captura = 0
    UserList(UserIndex).flags.tiene_bandera = 0
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.OnUserDisconnected", Erl)
End Sub

Private Sub IBaseScenario_OnUserReconnect(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    If RoundState = EnJuego Or RoundState = IniciaRonda Then
        UserList(UserIndex).flags.jugando_captura = 1
        Call EnviarUsuarioSala(UserIndex)
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.OnUserReconnect", Erl)
End Sub

Private Sub IBaseScenario_Reset()
    ' Contrato Reset (lo que R4 nunca hizo): soltar los mapas, salir del tick y
    ' dejar los flags de los jugadores limpios. Los jugadores los devuelve el
    ' lobby via RemovePlayerFromEvent (o ya los devolvio: ForceReset vacia
    ' RegisteredPlayers ANTES de llamar aca, por eso LimpiarEscenario no puede
    ' apoyarse en la lista del lobby).
    On Error GoTo ErrHandler
    Call LimpiarEscenario
    Call DesregistrarMapas
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.Reset", Erl)
End Sub

Private Sub IBaseScenario_PlayerKillNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal SourceType As e_DamageSourceType, ByVal SourceIndex As Integer)
End Sub

Private Sub IBaseScenario_NPcKillPlayer(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal SourceType As e_DamageSourceType, ByVal SourceIndex As Integer)
End Sub

Private Sub IBaseScenario_PlayerKillPlayer(ByVal killerIndex As Integer, ByVal deadIndex As Integer, ByVal SourceType As e_DamageSourceType, ByVal SourceIndex As Integer)
End Sub

Private Sub IBaseScenario_NpcDie(ByVal NpcIndex)
End Sub

Private Sub IBaseScenario_GetNextWaypointForNpc(ByVal NpcIndex As Integer, ByRef PosX As Integer, ByRef PosY As Integer)
End Sub

' ======================= Logica propia del CTF =======================

Private Sub ArrancarEvento()
    ' El lobby ya cobro la inscripcion e invoco a todos a la sala de espera.
    ' B2: los equipos los arma el CTF, NO SortTeams. SortTeams solo corre si el lobby
    ' viene con TeamType = eRandom y TeamSize > 1, y ninguna de las dos rutas de
    ' creacion del CTF los manda: Configuracion.ini trae TeamSize=0/TeamType=0 y el
    ' cliente (ProtocolCmdParse:1582) ni siquiera incluye esos campos en el payload.
    ' Resultado con el sorteo delegado: todos en el equipo 1 y la ronda se ganaba sola.
    On Error GoTo ErrHandler
    Dim i    As Integer
    Dim team As Integer
    'Msg1714
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_EVENTO_CAPTURA_BANDERA_INICIADO, vbNullString, e_FontTypeNames.FONTTYPE_GUILD))
    Call SortearDosEquipos
    For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
        If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
            team = LobbyList(LobbyIndex).Players(i).team
            If team < 1 Then team = 1
            If team > 2 Then team = 2
            With UserList(LobbyList(LobbyIndex).Players(i).User.ArrayIndex)
                .flags.CurrentTeam = team
                .flags.tiene_bandera = 0
                Call SendData(SendTarget.ToAll, LobbyList(LobbyIndex).Players(i).User.ArrayIndex, PrepareMessageUpdateFlag(0, .Char.charindex))
                'Msg1845
                Call WriteConsoleMsg(LobbyList(LobbyIndex).Players(i).User.ArrayIndex, PrepareMessageLocaleMsg(MSG_TOCADO_PARTICIPAR_EQUIPO_DEBERAS_PROTEGER_COMPANEROS_EQUIPO, colorEquipo(team), e_FontTypeNames.FONTTYPE_GUILD))
            End With
        End If
    Next i
    ronda_actual = 0
    rondasGanadas(1) = 0
    rondasGanadas(2) = 0
    Call IniciarRondaNueva
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.ArrancarEvento", Erl)
End Sub

Private Sub IniciarRondaNueva()
    On Error GoTo ErrHandler
    Dim i As Integer
    RoundState = IniciaRonda
    ronda_actual = ronda_actual + 1
    tiempo_ronda = CAPTURA_TIEMPO_INICIO_RONDA
    contadorBandera(1) = 0
    contadorBandera(2) = 0
    'Msg1720
    Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_RONDA_DARA_COMIENZO_SEGUNDOS, ronda_actual & "¬" & tiempo_ronda, e_FontTypeNames.FONTTYPE_GUILD))
    For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
        If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
            Dim ui As Integer
            ui = LobbyList(LobbyIndex).Players(i).User.ArrayIndex
            Call SendData(SendTarget.ToAll, ui, PrepareMessageUpdateFlag(0, UserList(ui).Char.charindex))
            UserList(ui).flags.jugando_captura_muertes = 0
            UserList(ui).flags.jugando_captura_timer = 0
            Call TirarBanderaDe(ui)
            Call EnviarUsuarioSala(ui)
        End If
    Next i
    Call RestaurarBandera(1)
    Call RestaurarBandera(2)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.IniciarRondaNueva", Erl)
End Sub

Private Sub TickCuentaRegresivaRonda()
    On Error GoTo ErrHandler
    Dim i As Integer
    tiempo_ronda = tiempo_ronda - 1
    If tiempo_ronda Mod 30 = 0 And tiempo_ronda > 5 Then
        'Msg1710
        Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_SEGUNDOS_COMIENZA_RONDA, tiempo_ronda & "¬" & ronda_actual, e_FontTypeNames.FONTTYPE_GUILD))
    ElseIf tiempo_ronda <= 5 And tiempo_ronda > 0 Then
        'Msg1711
        Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_RONDA_COMIENZA, tiempo_ronda, e_FontTypeNames.FONTTYPE_GUILD))
    ElseIf tiempo_ronda <= 0 Then
        'Msg1712
        Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_YAAA, vbNullString, e_FontTypeNames.FONTTYPE_GUILD))
        For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
            If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
                Call EnviarUsuarioBase(LobbyList(LobbyIndex).Players(i).User.ArrayIndex)
            End If
        Next i
        RoundState = EnJuego
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.TickCuentaRegresivaRonda", Erl)
End Sub

Private Sub TickEnJuego()
    On Error GoTo ErrHandler
    Dim i                     As Integer
    Dim ui                    As Integer
    Dim team                  As Byte
    Dim estaEnBandera(1 To 2) As Boolean
    For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
        If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
            ui = LobbyList(LobbyIndex).Players(i).User.ArrayIndex
            With UserList(ui)
                team = .flags.CurrentTeam
                If team >= 1 And team <= 2 Then
                    If .pos.Map = mapaEquipo(team) Then
                        If Abs(.pos.x - xEquipo(team)) <= 8 And Abs(.pos.y - yEquipo(team)) <= 5 Then
                            If TieneObjetos(objEquipo(IIf(team = 1, 2, 1)), 1, ui) Then
                                contadorBandera(team) = contadorBandera(team) + 1
                                estaEnBandera(team) = True
                                If contadorBandera(team) >= CAPTURA_TIEMPO_BANDERA Then
                                    Call FinalizarRonda(team)
                                    Exit Sub
                                Else
                                    'Msg1713
                                    Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_ATENCION_EQUIPO_TIENE_BANDERA_ENEMIGA_BASE, .name & "¬" & colorEquipo(team) & "¬" & (CAPTURA_TIEMPO_BANDERA - contadorBandera(team)), e_FontTypeNames.FONTTYPE_GUILD))
                                End If
                            End If
                        End If
                    ElseIf .flags.jugando_captura_timer > 0 Then
                        .flags.jugando_captura_timer = .flags.jugando_captura_timer - 1
                        If .flags.jugando_captura_timer = 0 Then
                            Call EnviarUsuarioBase(ui)
                        End If
                    End If
                End If
            End With
        End If
    Next i
    If Not estaEnBandera(1) Then contadorBandera(1) = 0
    If Not estaEnBandera(2) Then contadorBandera(2) = 0
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.TickEnJuego", Erl)
End Sub

Private Sub FinalizarRonda(ByVal team As Byte)
    On Error GoTo ErrHandler
    Dim rondasParaGanar As Integer
    'Msg1725
    Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_EQUIPO_GANADO_RONDA, colorEquipo(team) & "¬" & ronda_actual, e_FontTypeNames.FONTTYPE_GUILD))
    Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))
    contadorBandera(1) = 0
    contadorBandera(2) = 0
    rondasGanadas(team) = rondasGanadas(team) + 1
    rondasParaGanar = (cantidad_rondas \ 2) + 1
    If rondasGanadas(team) >= rondasParaGanar Then
        Call TerminarEvento(CInt(team))
    Else
        Call IniciarRondaNueva
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.FinalizarRonda", Erl)
End Sub

Private Sub TerminarEvento(ByVal team As Integer)
    ' Patron de cierre del plan: 1) marcar terminado, 2) pagar sobre un snapshot,
    ' 3) anunciar, 4) que el lobby libere. Nada de iterar mientras se muta, que
    ' es el bug que se salteaba la mitad de los jugadores.
    On Error GoTo ErrHandler
    Dim i         As Integer
    Dim ui        As Integer
    Dim recompensa As Long
    If RoundState = Terminado Then Exit Sub
    RoundState = Terminado
    If team >= 1 And team <= 2 Then
        recompensa = LobbyList(LobbyIndex).InscriptionPrice * 2
        For i = 0 To LobbyList(LobbyIndex).RegisteredPlayers - 1
            If LobbyList(LobbyIndex).Players(i).team = team Then
                ' GiveGoldToPlayer paga tambien al desconectado (ola 2).
                Call GiveGoldToPlayer(LobbyList(LobbyIndex), i, recompensa)
                If IsValidUserRef(LobbyList(LobbyIndex).Players(i).User) Then
                    ui = LobbyList(LobbyIndex).Players(i).User.ArrayIndex
                    'Msg1847
                    Call WriteConsoleMsg(ui, PrepareMessageLocaleMsg(MSG_HAS_RECUPERADO_INSCRIPCION_GANADO_MONEDAS_ORO_ADICIONALES, PonerPuntos(LobbyList(LobbyIndex).InscriptionPrice), e_FontTypeNames.FONTTYPE_GUILD))
                End If
            End If
        Next i
        'Msg1687
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_EVENTO_CAPTURA_BANDERA_TERMINADO_GANADOR_EQUIPO, colorEquipo(team), e_FontTypeNames.FONTTYPE_GUILD))
    Else
        'Msg1686 (empate)
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_EVENTO_CAPTURA_BANDERA_TERMINADO_EMPATE, vbNullString, e_FontTypeNames.FONTTYPE_GUILD))
    End If
    ' Pasar a Closed no cierra nada por si solo: UpdateLobbyState solo avisa y cambia
    ' el estado. El cierre real (devolver jugadores, soltar mapas, liberar el slot del
    ' pool) lo hace CloseScenario en el Update siguiente, igual que en los otros
    ' escenarios. Sin eso, los jugadores quedaban atrapados y se filtraba un lobby por
    ' cada CTF terminado.
    Call ModLobby.UpdateLobbyState(LobbyList(LobbyIndex), e_LobbyState.Closed)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.TerminarEvento", Erl)
End Sub

Private Sub LimpiarEscenario()
    ' B5: NO se puede iterar la lista del lobby ni mirar CurrentTeam para saber quien
    ' tiene una bandera. Los dos caminos de cierre llegan aca con esa informacion ya
    ' borrada: CancelLobby corre ReturnAllPlayers (que pone CurrentTeam = 0) antes de
    ' avisarle al escenario, y ForceReset vacia RegisteredPlayers antes de llamar Reset.
    ' Con la version anterior el jugador se quedaba con el estandarte enemigo puesto y
    ' ganaba la ronda del CTF siguiente con solo pararse en su base.
    ' Por eso: barrido por flags.jugando_captura, que es el mismo criterio que usa
    ' SendTarget.ToJugadoresCaptura y que nadie mas toca, y se quitan LAS DOS banderas
    ' sin preguntar de que equipo era.
    On Error GoTo ErrHandler
    Dim ui As Integer
    RoundState = Terminado
    MuertesPendientes = 0
    For ui = 1 To LastUser
        If UserList(ui).flags.jugando_captura = 1 Then
            Call SacarBanderasDe(ui)
            UserList(ui).flags.jugando_captura = 0
            UserList(ui).flags.tiene_bandera = 0
            UserList(ui).flags.CurrentTeam = 0
            UserList(ui).flags.jugando_captura_timer = 0
            UserList(ui).flags.jugando_captura_muertes = 0
        End If
    Next ui
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.LimpiarEscenario", Erl)
End Sub

Private Sub SacarBanderasDe(ByVal UserIndex As Integer)
    ' Version de limpieza: quita ambas banderas y no las tira al piso (el evento ya
    ' termino). TirarBanderaDe es la de juego y esa si las deja donde murio.
    On Error GoTo ErrHandler
    Dim actualizar As Boolean
    If TieneObjetos(OBJ_CAPTURA_BANDERA_1, 1, UserIndex) Then
        Call QuitarObjetos(OBJ_CAPTURA_BANDERA_1, 1, UserIndex)
        actualizar = True
    End If
    If TieneObjetos(OBJ_CAPTURA_BANDERA_2, 1, UserIndex) Then
        Call QuitarObjetos(OBJ_CAPTURA_BANDERA_2, 1, UserIndex)
        actualizar = True
    End If
    If actualizar Then
        UserList(UserIndex).flags.tiene_bandera = 0
        Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageUpdateFlag(0, UserList(UserIndex).Char.charindex))
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.SacarBanderasDe", Erl)
End Sub

Private Sub DesregistrarMapas()
    On Error GoTo ErrHandler
    If MapsRegistered Then
        Call CustomScenarios.RemoveCustomScenario(MAP_TEAM_1)
        Call CustomScenarios.RemoveCustomScenario(MAP_TEAM_2)
        Call CustomScenarios.RemoveCustomScenario(MAP_SALA_ESPERA)
        MapsRegistered = False
    End If
    If UpdateSlot >= 0 Then
        Call CustomScenarios.RemoveUpdateScenario(UpdateSlot)
        UpdateSlot = -1
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.DesregistrarMapas", Erl)
End Sub

Private Sub CloseScenario()
    ' Mismo contrato que ScenarioDeathMatch.CloseScenario: devolver, soltar la
    ' referencia, volver el lobby a UnInitilized, desregistrar y liberar el slot.
    ' El orden importa: primero limpiar banderas y flags (mientras el jugador todavia
    ' esta en el mapa del CTF), despues devolverlo.
    On Error GoTo ErrHandler
    Call LimpiarEscenario
    Call ModLobby.ReturnAllPlayers(LobbyList(LobbyIndex))
    Set LobbyList(LobbyIndex).Scenario = Nothing
    Call ModLobby.UpdateLobbyState(LobbyList(LobbyIndex), e_LobbyState.UnInitilized)
    Call DesregistrarMapas
    Call ReleaseLobby(LobbyIndex)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.CloseScenario", Erl)
End Sub

Private Sub SortearDosEquipos()
    ' Dos equipos parejos por nivel, serpentina 1-2-2-1-1-2... como hacia
    ' clsCaptura.iniciarCaptura. Con cupo impar el legacy CANCELABA el evento entero
    ' (y encima le faltaba el Exit Sub, que es uno de los bugs de esta ola); aca se
    ' echa al de menor nivel devolviendole la inscripcion, que es lo que ya hace
    ' SortTeams para los otros modos y no castiga a los otros 9 jugadores.
    On Error GoTo ErrHandler
    Dim i         As Integer
    Dim ui        As Integer
    Dim best      As Integer
    Dim bestLvl   As Integer
    Dim team      As Integer
    Dim direction As Integer
    With LobbyList(LobbyIndex)
        If .RegisteredPlayers Mod 2 = 1 Then
            best = -1
            bestLvl = 32767
            For i = 0 To .RegisteredPlayers - 1
                If IsValidUserRef(.Players(i).User) Then
                    If UserList(.Players(i).User.ArrayIndex).Stats.ELV <= bestLvl Then
                        bestLvl = UserList(.Players(i).User.ArrayIndex).Stats.ELV
                        best = i
                    End If
                Else
                    best = i
                    Exit For
                End If
            Next i
            If best >= 0 Then
                ' KickPlayer devuelve y refunda, pero no le avisa al escenario: el
                ' jugador se quedaba con jugando_captura = 1 y seguia recibiendo los
                ' mensajes del evento (y sin poder entrar a un reto) hasta el cierre.
                If IsValidUserRef(.Players(best).User) Then
                    ui = .Players(best).User.ArrayIndex
                    Call SacarBanderasDe(ui)
                    UserList(ui).flags.jugando_captura = 0
                    UserList(ui).flags.CurrentTeam = 0
                    UserList(ui).flags.jugando_captura_timer = 0
                    UserList(ui).flags.jugando_captura_muertes = 0
                End If
                Call KickPlayer(LobbyList(LobbyIndex), best)
            End If
        End If
        For i = 0 To .RegisteredPlayers - 1
            .Players(i).team = 0
        Next i
        team = 1
        direction = 1
        Do
            best = -1
            bestLvl = -1
            For i = 0 To .RegisteredPlayers - 1
                If .Players(i).team = 0 And IsValidUserRef(.Players(i).User) Then
                    If UserList(.Players(i).User.ArrayIndex).Stats.ELV > bestLvl Then
                        bestLvl = UserList(.Players(i).User.ArrayIndex).Stats.ELV
                        best = i
                    End If
                End If
            Next i
            If best < 0 Then Exit Do
            .Players(best).team = team
            UserList(.Players(best).User.ArrayIndex).flags.CurrentTeam = team
            team = team + direction
            If team > 2 Then
                direction = -1
                team = 2
            ElseIf team < 1 Then
                direction = 1
                team = 1
            End If
        Loop
        ' Los que se desconectaron entre la inscripcion y el arranque no tienen nivel
        ' que consultar; se reparten alternando para no desbalancear el conteo.
        For i = 0 To .RegisteredPlayers - 1
            If .Players(i).team = 0 Then
                .Players(i).team = team
                team = team + direction
                If team > 2 Then
                    direction = -1
                    team = 2
                ElseIf team < 1 Then
                    direction = 1
                    team = 1
                End If
            End If
        Next i
        .TeamSortDone = True
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.SortearDosEquipos", Erl)
End Sub

Private Sub EnviarUsuarioBase(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Dim pos_map As Integer
    Dim pos_X   As Byte
    Dim pos_Y   As Byte
    If UserList(UserIndex).flags.CurrentTeam = 1 Then
        pos_map = MAP_TEAM_1
        pos_X = X_TEAM_1
        pos_Y = Y_TEAM_1
    ElseIf UserList(UserIndex).flags.CurrentTeam = 2 Then
        pos_map = MAP_TEAM_2
        pos_X = X_TEAM_2
        pos_Y = Y_TEAM_2
    Else
        Exit Sub
    End If
    Call FindLegalPos(UserIndex, pos_map, pos_X, pos_Y)
    Call WarpUserChar(UserIndex, pos_map, pos_X, pos_Y, True)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.EnviarUsuarioBase", Erl)
End Sub

Private Sub EnviarUsuarioSala(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Dim pos_X As Byte
    Dim pos_Y As Byte
    pos_X = RandomNumber(MIN_SALA_ESPERA_X, MAX_SALA_ESPERA_X)
    pos_Y = RandomNumber(MIN_SALA_ESPERA_Y, MAX_SALA_ESPERA_Y)
    Call FindLegalPos(UserIndex, MAP_SALA_ESPERA, pos_X, pos_Y)
    Call WarpUserChar(UserIndex, MAP_SALA_ESPERA, pos_X, pos_Y, True)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.EnviarUsuarioSala", Erl)
End Sub

Private Sub RestaurarBandera(ByVal team As Byte)
    On Error GoTo ErrHandler
    Dim MiObj As t_Obj
    Dim MiPos As t_WorldPos
    MiObj.amount = 1
    MiObj.ObjIndex = objEquipo(team)
    MiPos.Map = mapaEquipo(team)
    MiPos.x = xEquipo(team)
    MiPos.y = yEquipo(team)
    If MapData(MiPos.Map, MiPos.x, MiPos.y).ObjInfo.amount <= 0 Then
        Call TirarItemAlPiso(MiPos, MiObj, False)
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.RestaurarBandera", Erl)
End Sub

Private Sub TirarBanderaDe(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Dim team       As Byte
    Dim objBandera As Integer
    With UserList(UserIndex)
        team = .flags.CurrentTeam
        If team = 1 Then
            objBandera = OBJ_CAPTURA_BANDERA_2
        ElseIf team = 2 Then
            objBandera = OBJ_CAPTURA_BANDERA_1
        Else
            Exit Sub
        End If
        If TieneObjetos(objBandera, 1, UserIndex) Then
            .flags.tiene_bandera = 0
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageUpdateFlag(0, .Char.charindex))
            Call QuitarObjetos(objBandera, 1, UserIndex)
            If RoundState = EnJuego Then
                Dim MiObj As t_Obj
                MiObj.amount = 1
                MiObj.ObjIndex = objBandera
                Call TirarItemAlPiso(.pos, MiObj, False)
                'Msg1719
                Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_ATENCION_BANDERA_EQUIPO_ENCUENTRA_1719, colorEquipo(team) & "¬" & .pos.Map & "," & .pos.x & "," & .pos.y, e_FontTypeNames.FONTTYPE_GUILD))
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.TirarBanderaDe", Erl)
End Sub

' =============== API publica para los call sites de banderas ===============
' Las llama InvUsuario via CustomScenarios.GetMap + TypeOf.

Public Function TomaBandera(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
    On Error GoTo ErrHandler
    Dim banderaTeam As Byte
    TomaBandera = True
    If ObjIndex = OBJ_CAPTURA_BANDERA_1 Then
        banderaTeam = 1
    ElseIf ObjIndex = OBJ_CAPTURA_BANDERA_2 Then
        banderaTeam = 2
    Else
        Exit Function
    End If
    With UserList(UserIndex)
        If banderaTeam = .flags.CurrentTeam Then
            If .pos.x = xEquipo(banderaTeam) And .pos.y = yEquipo(banderaTeam) Then
                TomaBandera = False
                Exit Function
            End If
        End If
        If .flags.CurrentTeam <> banderaTeam Then
            .flags.tiene_bandera = banderaTeam
        Else
            .flags.tiene_bandera = 0
        End If
        Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageUpdateFlag(.flags.tiene_bandera, .Char.charindex))
    End With
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.TomaBandera", Erl)
    TomaBandera = True
End Function

Public Sub QuitarBandera(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
    On Error GoTo ErrHandler
    Dim banderaTeam As Byte
    If ObjIndex = OBJ_CAPTURA_BANDERA_1 Then
        banderaTeam = 1
    ElseIf ObjIndex = OBJ_CAPTURA_BANDERA_2 Then
        banderaTeam = 2
    Else
        Exit Sub
    End If
    With UserList(UserIndex)
        If banderaTeam <> .flags.CurrentTeam Then
            'Msg1716
            Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_ATENCION_EQUIPO_TOMADO_BANDERA_EQUIPO, .name & "¬" & colorEquipo(.flags.CurrentTeam) & "¬" & colorEquipo(banderaTeam), e_FontTypeNames.FONTTYPE_GUILD))
            .flags.tiene_bandera = banderaTeam
        Else
            .flags.tiene_bandera = 0
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageUpdateFlag(0, .Char.charindex))
            Call QuitarObjetos(ObjIndex, 1, UserIndex)
            Call RestaurarBandera(banderaTeam)
            'Msg1717
            Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_ATENCION_BANDERA_EQUIPO_SIDO_RECUPERADA, colorEquipo(banderaTeam), e_FontTypeNames.FONTTYPE_GUILD))
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.QuitarBandera", Erl)
End Sub

Public Sub TiraBandera(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
    On Error GoTo ErrHandler
    Dim team As Byte
    If ObjIndex = OBJ_CAPTURA_BANDERA_1 Then
        team = 1
    ElseIf ObjIndex = OBJ_CAPTURA_BANDERA_2 Then
        team = 2
    Else
        Exit Sub
    End If
    contadorBandera(team) = 0
    With UserList(UserIndex)
        .flags.tiene_bandera = 0
        Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageUpdateFlag(0, .Char.charindex))
        'Msg1718
        Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageLocaleMsg(MSG_EVENTOS_ATENCION_BANDERA_EQUIPO_ENCUENTRA, colorEquipo(team) & "¬" & .pos.Map & "," & .pos.x & "," & .pos.y, e_FontTypeNames.FONTTYPE_GUILD))
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ScenarioCaptureTheFlag.TiraBandera", Erl)
End Sub

' NOTA: no hay CancelarPorGm. /cancelarevento entra por HandleCancelarEvento ->
' CancelLobby, que refunda, devuelve y pasa el lobby a Closed; el escenario se entera
' por UpdateLobbyState y se limpia solo.
