Attribute VB_Name = "modElementalBalanceLog"
' Argentum 20 Game Server
'
' modElementalBalanceLog
' Log estructurado de balance del motor elemental (plan 08.003, Ola 3 punto 6).
' Calcado de modPoisonLogging.bas: header escrito una vez, filas separadas por
' ";", costo cero cuando AMBOS toggles estan apagados. NO reemplaza ni toca
' modElementalCombat.ElementalLog (texto libre, Logs\Elemental_YYYY-MM-DD.log):
' este modulo agrega, EN PARALELO, una segunda escritura estructurada.
'
' ============================================================================
' Esquema v2 (plan 10.001, Ola 5, tajada A) -- decisiones y por que
' ============================================================================
' El esquema v1 (informe 2026-09-08, ver historial de este archivo en git) no
' podia responder telemetria de jugadores (Hecho 22 del plan 10.001): "item" y
' "tier" iban siempre en 0, y "attacker"/"victim" son indices de slot que se
' reusan entre jugadores. v2 agrega:
'
'   - char_id / account_id (columnas nuevas, al final): id estable del
'     ATACANTE (UserList(i).Id / .AccountID, Declares.bas:3265,3267). 0 si el
'     atacante es NPC o el slot no es un user logueado.
'   - victim_char_id (columna nueva, al final): id estable de la VICTIMA
'     cuando es un user logueado; 0 si es NPC.
'   - src_item (columna nueva): ObjIndex del aceite/orbe que originó un
'     encantamiento activo (Hecho 43). Solo se conoce en el camino
'     ench/ammoench de ElementalDamageUserVsTarget y en los eventos
'     enchant_weapon/enchant_ammo/charge_spent; 0 en cualquier otro evento.
'   - fight_id (columna nueva): identificador de pelea PvP (decision 12).
'     0 si el evento no es user-vs-user (una pelea son SIEMPRE dos
'     personajes; NPC vs user no arma "pelea" para B4).
'   - schema (columna nueva, al final): version numerica de este esquema
'     (constante ELEMENTAL_BALANCE_LOG_SCHEMA_VERSION), para que un lector
'     que junte varios archivos sepa sin ambiguedad que fila es de que forma.
'   - item / tier: se documentaban SIEMPRE 0; ahora llevan el ObjIndex real y
'     el tier del catalogo (04.003/07.001) donde el call site lo tiene a
'     mano. Ver "Decision item/tier por call site" mas abajo.
'
' Decision item/tier por call site (documentada en la bitacora de la Ola 5,
' plan 10.001, con la razon completa; resumen aca para quien lea el codigo):
'   - ResolveComponentsVsTarget y FireProcs (las dos funciones genericas que
'     el Hecho 22 senala) GANARON parametros nuevos (attackerIndex,
'     attackerType, itemObjIndex, srcItemObjIndex): las 6 llamadas que hace
'     ElementalDamageUserVsTarget/ElementalDamageNpcVsUser YA conocen esos
'     datos en el momento de llamar (arma/municion/orbe/encantamiento
'     equipados, o 0 en el camino NPC). Se opto por PROPAGAR, no por dejar
'     item=0: es la unica forma de que exista B3 (Objetivo B punto 9).
'   - ApplyDotTickResist NO gano attacker/item: solo recibe datos del target
'     (Hecho 22 lo documenta como limitacion de la funcion, no de este
'     modulo); agrandar su firma es rediseno del motor DoT, fuera de esta
'     Ola. Si gana victim_char_id, que SI esta a mano.
'   - ApplyThornsDamage/ResolveThorns/FireSlotThorns SI ganaron un parametro
'     de ObjIndex (el slot de gear que disparo la espina): FireSlotThorns ya
'     lo tenia y no lo pasaba (Hecho documentado en el header v1); es una
'     propagacion de una sola funcion, no un fan-in como el de arriba.
'   - TryUniversalCrit: el atacante (UserIndex) ya era parametro; se agregan
'     char_id/account_id/victim_char_id/fight_id calculados con eso. item
'     sigue en 0 (el critico universal no esta atado a un item, ya
'     documentado en v1).
'
' Tier del catalogo (OBJ9084-9161, docs/catalogo-elemental-jugable.md): solo
' los rangos que ese documento etiqueta explicitamente "T1"/"T2" devuelven
' 1/2. El resto del catalogo (flechas, Aceites Superiores, orbes, amuletos,
' gear de espinas, veneno-aceites, antidotos, pergaminos) NO tiene una T
' explicita en el doc: se devuelve 0 (no disponible) en vez de adivinar.
' Ver ElementalBalanceCatalogTier(). ElementalBalanceInCatalog() es un chequeo
' de RANGO aparte (9084-9161 completo) para filtrar craft/equip por catalogo
' sin depender de si ese item tiene tier asignado.
'
' Gotcha de compatibilidad del parser (Hecho 26, decision tomada en la Ola 5):
' `parsear_log_balance()` (tools/testclient/calibrar_dos_fuentes.py) es
' header-driven pero ESTRICTO en ancho: una fila con mas/menos campos que el
' ultimo header leido se descarta EN SILENCIO. Si v2 escribiera en el MISMO
' archivo que ya tiene el header v1 (escrito antes de este cambio), todas las
' filas v2 se perderian sin aviso porque LogElementalBalance solo escribe
' header cuando el archivo no existe (`Dir(fname)` vacio). Se eligio
' ARCHIVO POR VERSION DE ESQUEMA en vez de forzar un nuevo header en el
' archivo v1 a mitad de dia: separa fisicamente datos de forma distinta,
' hace trivial que Ola 6 sepa que esquema esta leyendo por el nombre del
' archivo, y no depende de que el escritor detecte "hoy ya tiene header
' viejo" (una condicion mas para romperse en silencio si alguien la toca sin
' leer este comentario). Archivo: Logs\ElementalBalance_YYYY-MM-DD_v2.log
' (el v1 sin sufijo, si algun dia se reactivara, seguiria en
' Logs\ElementalBalance_YYYY-MM-DD.log intacto). Probado en
' tools/testclient/tests/test_calibrar_dos_fuentes.py: un texto con header v1
' y otro con header v2 parsean cada uno con sus propias columnas.
'
' Retencion (decision 10, plan 10.001): esta Ola NO implementa rotacion
' automatica ("borrado cuando se pida, sin rotacion automatica"). Comando
' exacto para el operador cuando el dueno pida purgar (PowerShell, parado en
' dev/server; CONSERVAR una copia si todavia no se corrio
' analizar_telemetria.py sobre esos dias):
'   Remove-Item ".\Logs\ElementalBalance_*.log"
' La Ola 7 evalua si las tareas de backup ya existentes en la VM alcanzan
' para cubrir tambien esta carpeta antes de automatizar nada (decision 10,
' inciso a). Hasta entonces el archivo crece sin limite: riesgo aceptado por
' el dueno, no un olvido de esta Ola.
'
' Toggle de jugadores (plan 10.001, Objetivo B punto 4 y Ola 5 punto 4):
' TOGGLE33 "elemental_player_telemetry", independiente de "elemental_balance_log".
' Los eventos "raros" (enchant_weapon, enchant_ammo, craft, buy, equip,
' kill, death) y el unico evento de hot path nuevo (charge_spent) se gatean
' con ElementalPlayerTelemetryEnabled(), NO con ElementalBalanceLogEnabled():
' asi se puede prender la telemetria de jugadores en beta con el log de
' calibracion por golpe (el que pesa por swing, Hecho 24) apagado. Los 6
' eventos v1 (component, dot_tick_resist, proc_dmg_bonus, proc_apply_state,
' thorns, universal_crit) siguen gateados SOLO por elemental_balance_log,
' sin cambios de comportamiento.
'
' IMPORTANTE -- LogElementalBalance() DEJO de autogatear en
' "elemental_balance_log": con dos toggles independientes escribiendo al
' mismo Sub, ese gate unico hubiera apagado los eventos de jugador cada vez
' que el log de calibracion este OFF (justo el escenario que la decision 7
' pide habilitar). El Sub ahora sale temprano solo si LOS DOS toggles estan
' apagados (red de seguridad barata); cada call site sigue siendo responsable
' de chequear SU propio toggle antes de llamar, exactamente como ya hacian
' los 6 call sites v1 (`If ElementalBalanceLogEnabled() Then Call
' LogElementalBalance(...)`).
'
' Esquema de fila v2:
'   ts_ms;event;attacker;attacker_class;victim;victim_type;item;tier;
'   dmg_type;raw;final;resist_pct;pvp;map;char_id;account_id;
'   victim_char_id;src_item;fight_id;schema
'
' Los primeros 14 campos conservan el significado v1 (ver historial git para
' el docstring completo pre-v2); lo que cambia es que item/tier ahora pueden
' venir poblados. Campos nuevos:
'   char_id        = UserList(idx).Id del ATACANTE si es user logueado; 0 si
'                    es NPC o el slot no tiene un user logueado (Hecho 23).
'   account_id     = UserList(idx).AccountID del ATACANTE, mismo criterio.
'   victim_char_id = UserList(idx).Id de la VICTIMA si es user logueado; 0 si
'                    es NPC.
'   src_item       = ObjIndex del aceite/orbe fuente del encantamiento activo
'                    (Hecho 43); 0 si el evento no viene de un encantamiento.
'   fight_id       = ver ElementalBalanceFightId(). 0 si no aplica (no es
'                    user-vs-user, o el evento no participa de una pelea).
'   schema         = ELEMENTAL_BALANCE_LOG_SCHEMA_VERSION (2).
'
' Eventos nuevos de esta Ola (todos detras de elemental_player_telemetry):
'   enchant_weapon / enchant_ammo -> SetEnchantedWeapon/SetEnchantedAmmo
'     (modElementalCombat.bas). raw=cargas otorgadas, final=duracion en ms
'     (-1 = permanente). item=arma/municion encantada, src_item=aceite/orbe.
'   charge_spent -> OnEnchantedWeaponSwing/OnEnchantedAmmoSwing. UNICO evento
'     nuevo en el hot path (un swing/disparo con encantamiento activo por
'     cargas). raw=cargas restantes tras el descuento.
'   craft -> Trabajo.bas (4 oficios), solo si el item construido cae en el
'     catalogo (ElementalBalanceInCatalog). raw=cantidad construida.
'   buy -> Comercio.bas, compra ya consumada (Hecho 38). raw=cantidad,
'     final=precio total pagado. Sin filtro de catalogo (Hecho 38 no lo
'     pide; el filtro se aplica en el lector, Ola 6, y el evento es raro).
'   equip -> InvUsuario.EquiparInvItem, solo catalogo. Sin cantidad/precio.
'   kill / death -> Modulo_UsUaRiOs.DoDamageOrHeal, bloque de muerte
'     (:3112-3117 en la version citada por el plan), ANTES de que
'     CustomScenarios.UserDie/ActStats->UserDie borren el equipo (Hecho 42).
'     Una sola fila por muerte. item = arma equipada del ATACANTE (si es
'     user); src_item = arma equipada de la VICTIMA. Decision documentada en
'     la bitacora: NO se saca una foto completa de los 7 slots de gear de
'     ambos lados en esta tajada (hubiera exigido varias filas o columnas
'     nuevas por slot); item/src_item dan la pieza mas relevante (el arma)
'     de cada lado. raw=final=dano de la aplicacion letal (Abs(amount), ya
'     post-mitigacion). Cierra el fight_id si attacker y victim son ambos
'     user (decision 12: la pelea termina con la muerte de uno).
'
' Campos en 0 documentados por punto de insercion v1 (sin cambios de esta
' Ola salvo lo ya descripto arriba):
'   :dot_tick_resist ApplyDotTickResist -> attacker/attacker_class/item/
'         char_id/account_id/src_item/fight_id en 0: la funcion solo recibe
'         datos del target (targetIsNpc/targetIndex/dmgType/rawDamage), no
'         conoce el origen del DoT. victim_char_id SI se llena (el target es
'         conocido).
'   :component ResolveComponentsVsTarget -> desde esta Ola SI recibe
'         attacker/item/src_item (ver arriba); sigue en 0 solo cuando el
'         call site no los tiene (camino NPC->user).
'   :proc_dmg_bonus / :proc_apply_state FireProcs -> attacker/attacker_class
'         ya estaban disponibles (parametros de FireProcs); item/src_item
'         ahora tambien, mismo criterio que component. raw/final de
'         proc_apply_state siguen en 0 (no tiene un numero de dano propio).
'   :thorns ApplyThornsDamage -> item ahora SI se llena (ObjIndex del slot
'         que disparo la espina, propagado desde FireSlotThorns); raw sigue
'         en 0 (el dano pre-reduccion se pisa en ResolveThorns, no llega
'         aca). OJO roles invertidos: 'attackerIndex' es quien RECIBE el
'         reflejo (victim de este evento) y 'defenderIndex' quien lo origina
'         (attacker de este evento) -- son los nombres del golpe ORIGINAL
'         que genero las espinas.
'   :universal_crit TryUniversalCrit -> item sigue en 0 (el critico
'         universal no esta atado a un item puntual); raw en 0 (el bonus
'         pre-resistencia se pisa en la misma variable antes de este punto).

Option Explicit

Private Const ELEMENTAL_BALANCE_LOG_HEADER As String = "ts_ms;event;attacker;attacker_class;victim;victim_type;item;tier;dmg_type;raw;final;resist_pct;pvp;map;char_id;account_id;victim_char_id;src_item;fight_id;schema"
Private Const ELEMENTAL_BALANCE_LOG_SCHEMA_VERSION As Long = 2

' Catalogo elemental jugable (plan 04.003/07.001): OBJ9084-9161, 78 items.
Private Const ELEMENTAL_CATALOG_MIN As Long = 9084
Private Const ELEMENTAL_CATALOG_MAX As Long = 9161

' Decision 12 (plan 10.001): ventana de silencio que cierra una pelea sin
' muerte de por medio. 30s: corto para no fundir dos peleas distintas,
' largo para tolerar una persecucion o una pausa por pocion.
Private Const ELEMENTAL_FIGHT_SILENCE_MS As Long = 30000
Private Const MAX_ACTIVE_ELEMENTAL_FIGHTS As Long = 256

Private Type t_ElementalFight
    CharA As Long
    CharB As Long
    FightId As Long
    LastTickMs As Long
    Active As Boolean
End Type

Private mFights(1 To MAX_ACTIVE_ELEMENTAL_FIGHTS) As t_ElementalFight
Private mNextFightId As Long

Public Function ElementalBalanceLogEnabled() As Boolean
    ElementalBalanceLogEnabled = IsFeatureEnabled("elemental_balance_log")
End Function

' TOGGLE33 (plan 10.001, decision 7): telemetria de eventos de jugador,
' independiente del log de calibracion por golpe. Ver el bloque "Toggle de
' jugadores" arriba del todo de este archivo.
Public Function ElementalPlayerTelemetryEnabled() As Boolean
    ElementalPlayerTelemetryEnabled = IsFeatureEnabled("elemental_player_telemetry")
End Function

' Arma el identificador "U<idx>"/"N<idx>" para attacker/victim. "0" si idx<=0.
Public Function ElementalBalanceActorId(ByVal isNpc As Boolean, ByVal idx As Integer) As String
    If idx <= 0 Then
        ElementalBalanceActorId = "0"
    ElseIf isNpc Then
        ElementalBalanceActorId = "N" & idx
    Else
        ElementalBalanceActorId = "U" & idx
    End If
End Function

' Nombre de clase (ListaClases) si el actor es user; "0" si es NPC o el
' indice es invalido (con guard de error: un UserIndex desconectado entre el
' momento del golpe y el logueo no debe tirar la fila entera).
Public Function ElementalBalanceActorClass(ByVal isNpc As Boolean, ByVal idx As Integer) As String
    On Error GoTo eh
    If isNpc Or idx <= 0 Then
        ElementalBalanceActorClass = "0"
        Exit Function
    End If
    ElementalBalanceActorClass = ListaClases(UserList(idx).clase)
    Exit Function
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.ElementalBalanceActorClass", Erl)
    ElementalBalanceActorClass = "0"
End Function

' Mapa del actor (target o atacante). 0 si el indice es invalido.
Public Function ElementalBalanceMap(ByVal isNpc As Boolean, ByVal idx As Integer) As Integer
    On Error GoTo eh
    If idx <= 0 Then Exit Function
    If isNpc Then
        ElementalBalanceMap = NpcList(idx).pos.Map
    Else
        ElementalBalanceMap = UserList(idx).pos.Map
    End If
    Exit Function
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.ElementalBalanceMap", Erl)
    ElementalBalanceMap = 0
End Function

' Id estable de personaje (Hecho 23: UserList(idx).Id, NO el UserIndex/slot
' que se reusa entre jugadores). 0 si es NPC, el indice es invalido, o el
' slot no tiene hoy un user logueado (evita atribuir a un slot reciclado).
Public Function ElementalBalanceCharId(ByVal isNpc As Boolean, ByVal idx As Integer) As Long
    On Error GoTo eh
    If isNpc Or idx <= 0 Then Exit Function
    If idx > UBound(UserList) Then Exit Function
    If Not UserList(idx).flags.UserLogged Then Exit Function
    ElementalBalanceCharId = UserList(idx).Id
    Exit Function
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.ElementalBalanceCharId", Erl)
    ElementalBalanceCharId = 0
End Function

' Id estable de cuenta (Hecho 23: UserList(idx).AccountID). Mismo criterio
' que ElementalBalanceCharId.
Public Function ElementalBalanceAccountId(ByVal isNpc As Boolean, ByVal idx As Integer) As Long
    On Error GoTo eh
    If isNpc Or idx <= 0 Then Exit Function
    If idx > UBound(UserList) Then Exit Function
    If Not UserList(idx).flags.UserLogged Then Exit Function
    ElementalBalanceAccountId = UserList(idx).AccountID
    Exit Function
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.ElementalBalanceAccountId", Erl)
    ElementalBalanceAccountId = 0
End Function

' True si el ObjIndex cae en el catalogo elemental jugable (OBJ9084-9161,
' plan 04.003). Usado para filtrar craft/equip "solo catalogo" (plan 10.001,
' Ola 5 punto 3) SIN depender de si ese item tiene un tier explicito: hay
' items del catalogo (orbes, flechas, pergaminos) sin T1/T2 declarado en
' docs/catalogo-elemental-jugable.md pero que igual son del catalogo.
Public Function ElementalBalanceInCatalog(ByVal objIndex As Long) As Boolean
    ElementalBalanceInCatalog = (objIndex >= ELEMENTAL_CATALOG_MIN And objIndex <= ELEMENTAL_CATALOG_MAX)
End Function

' Tier (1/2) SOLO para los rangos que docs/catalogo-elemental-jugable.md
' etiqueta explicitamente "Arma/Anillo/Tunica/Armadura T1" o "T2". El resto
' del catalogo (flechas 9096-9100, Aceites Superiores 9106-9110, orbes
' 9111-9115, gear de espinas 9130-9132, veneno-aceites 9133-9135,
' antidotos 9136-9137, pergaminos 9138-9142/9161) no tiene una T explicita
' en el doc: se devuelve 0 (no disponible) en vez de adivinar. Decision
' documentada en la bitacora de la Ola 5 (plan 10.001).
Public Function ElementalBalanceCatalogTier(ByVal objIndex As Long) As Long
    Select Case objIndex
        Case 9084 To 9089            ' Armas T1
            ElementalBalanceCatalogTier = 1
        Case 9090 To 9095            ' Armas T2
            ElementalBalanceCatalogTier = 2
        Case 9101 To 9105            ' Aceites T1
            ElementalBalanceCatalogTier = 1
        Case 9116 To 9121            ' Anillos T1
            ElementalBalanceCatalogTier = 1
        Case 9122 To 9127            ' Anillos T2
            ElementalBalanceCatalogTier = 2
        Case 9128 To 9129            ' Amuletos T2
            ElementalBalanceCatalogTier = 2
        Case 9143 To 9148            ' Tunicas T1
            ElementalBalanceCatalogTier = 1
        Case 9149 To 9154            ' Armaduras T1
            ElementalBalanceCatalogTier = 1
        Case 9155 To 9157            ' Tunicas T2
            ElementalBalanceCatalogTier = 2
        Case 9158 To 9160            ' Armaduras T2
            ElementalBalanceCatalogTier = 2
        Case Else
            ElementalBalanceCatalogTier = 0
    End Select
End Function

Private Function FindElementalFightSlot(ByVal a As Long, ByVal b As Long) As Long
    Dim i As Long
    For i = 1 To MAX_ACTIVE_ELEMENTAL_FIGHTS
        If mFights(i).Active Then
            If (mFights(i).CharA = a And mFights(i).CharB = b) Or (mFights(i).CharA = b And mFights(i).CharB = a) Then
                FindElementalFightSlot = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function FreeElementalFightSlot() As Long
    Dim i As Long
    For i = 1 To MAX_ACTIVE_ELEMENTAL_FIGHTS
        If Not mFights(i).Active Then
            FreeElementalFightSlot = i
            Exit Function
        End If
    Next i
    ' Tabla llena (raro: 256 pares de PvP simultaneos): reusar el slot mas
    ' viejo por LastTickMs en vez de perder el evento nuevo.
    Dim oldest As Long, oldestTick As Long
    oldest = 1
    oldestTick = mFights(1).LastTickMs
    For i = 2 To MAX_ACTIVE_ELEMENTAL_FIGHTS
        If mFights(i).LastTickMs < oldestTick Then
            oldestTick = mFights(i).LastTickMs
            oldest = i
        End If
    Next i
    FreeElementalFightSlot = oldest
End Function

' Identificador de pelea (decision 12, plan 10.001): par {A,B} de char_id
' ESTABLES (no UserIndex). Arranca en el primer golpe entre los dos, se
' mantiene mientras haya eventos separados por menos de
' ELEMENTAL_FIGHT_SILENCE_MS, y se asigna en el SERVER (no se reconstruye en
' el lector). Devuelve 0 si cualquiera de los dos lados no es un personaje
' identificado (charA<=0 Or charB<=0): B4 es tiempo-hasta-matar PvP, un NPC
' no arma "pelea". closeFight=True cierra el par (llamado desde kill/death,
' decision 12: "cierra con la muerte de uno").
Public Function ElementalBalanceFightId(ByVal charA As Long, ByVal charB As Long, Optional ByVal closeFight As Boolean = False) As Long
    On Error GoTo eh
    If charA <= 0 Or charB <= 0 Then Exit Function
    Dim nowTick As Long
    nowTick = GetTickCountRaw()
    Dim slot As Long
    slot = FindElementalFightSlot(charA, charB)
    If slot > 0 Then
        If DeadlinePassed(nowTick, AddMod32(mFights(slot).LastTickMs, ELEMENTAL_FIGHT_SILENCE_MS)) Then
            ' Silencio > ventana: la pelea anterior ya se cerro sola: esta cuenta como una nueva.
            mFights(slot).Active = False
            slot = 0
        End If
    End If
    If slot = 0 Then
        slot = FreeElementalFightSlot()
        mNextFightId = mNextFightId + 1
        mFights(slot).CharA = charA
        mFights(slot).CharB = charB
        mFights(slot).FightId = mNextFightId
        mFights(slot).Active = True
    End If
    mFights(slot).LastTickMs = nowTick
    ElementalBalanceFightId = mFights(slot).FightId
    If closeFight Then mFights(slot).Active = False
    Exit Function
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.ElementalBalanceFightId", Erl)
End Function

' Nombre del archivo del dia para el esquema v2 (Hecho 26: archivo por
' version de esquema, ver el bloque de comentarios "Gotcha de compatibilidad
' del parser" arriba).
Private Function ElementalBalanceLogFileName() As String
    ElementalBalanceLogFileName = App.Path & "\Logs\ElementalBalance_" & Format$(Date, "yyyy-mm-dd") & "_v2.log"
End Function

Public Sub LogElementalBalance(ByVal evento As String, _
                                ByVal attacker As String, _
                                ByVal attackerClass As String, _
                                ByVal victim As String, _
                                ByVal victimType As String, _
                                ByVal itemObjIndex As Long, _
                                ByVal tier As Long, _
                                ByVal dmgType As Long, _
                                ByVal raw As Long, _
                                ByVal finalDmg As Long, _
                                ByVal mapNumber As Integer, _
                                ByVal charId As Long, _
                                ByVal accountId As Long, _
                                ByVal victimCharId As Long, _
                                ByVal srcItem As Long, _
                                ByVal fightId As Long)
    On Error GoTo ErrHandler
    ' Red de seguridad barata (ver "IMPORTANTE" arriba): el gate PRECISO es
    ' responsabilidad de cada call site (su propio toggle). Esto solo evita
    ' escribir CUALQUIER cosa si los dos toggles estan apagados.
    If Not (ElementalBalanceLogEnabled() Or ElementalPlayerTelemetryEnabled()) Then Exit Sub
    Dim resistPct As Long
    If raw > 0 Then resistPct = (raw - finalDmg) * 100 \ raw
    Dim pvp As Byte
    If victimType = "user" Then pvp = 1
    Dim fname As String
    fname = ElementalBalanceLogFileName()
    Dim fnum As Integer
    fnum = FreeFile
    Dim writeHeader As Boolean
    writeHeader = (LenB(dir(fname)) = 0)
    Open fname For Append As #fnum
    If writeHeader Then Print #fnum, ELEMENTAL_BALANCE_LOG_HEADER
    Print #fnum, GetTickCountRaw() & ";" & _
                 evento & ";" & attacker & ";" & attackerClass & ";" & _
                 victim & ";" & victimType & ";" & itemObjIndex & ";" & tier & ";" & _
                 dmgType & ";" & raw & ";" & finalDmg & ";" & resistPct & ";" & _
                 pvp & ";" & mapNumber & ";" & charId & ";" & accountId & ";" & _
                 victimCharId & ";" & srcItem & ";" & fightId & ";" & ELEMENTAL_BALANCE_LOG_SCHEMA_VERSION
    Close #fnum
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.LogElementalBalance", Erl)
    On Error Resume Next
    Close #fnum
End Sub
