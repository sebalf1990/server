Attribute VB_Name = "modElementalBalanceLog"
' Argentum 20 Game Server
'
' modElementalBalanceLog
' Log estructurado de balance del motor elemental (plan 08.003, Ola 3 punto 6).
' Calcado de modPoisonLogging.bas: header escrito una vez, filas separadas por
' ";", costo cero cuando el toggle esta apagado. NO reemplaza ni toca
' modElementalCombat.ElementalLog (texto libre, Logs\Elemental_YYYY-MM-DD.log):
' este modulo agrega, EN PARALELO, una segunda escritura estructurada en los
' mismos 6 puntos donde ElementalLog ya escribe hoy
' (modElementalCombat.bas:367,403,433,463,571,1229).
'
' Por que un log aparte y no ampliar ElementalLog: el criterio "calibracion de
' dos fuentes" del plan 08.003 (Objetivo punto 5) necesita timestamps con
' resolucion de milisegundos y columnas tipadas para contrastar contra la
' decodificacion de wire de la sonda (tools/testclient); ElementalLog es texto
' libre con timestamp al segundo (informe 2026-09-08-analisis-banco-de-
' balance.md seccion 6.1), insuficiente para correlacionar golpes que ocurren
' cada ~1.2s.
'
' Toggle: IsFeatureEnabled("elemental_balance_log") (feature_toggle.ini,
' clave nueva, independiente de "elemental_system"). Los 6 puntos de
' insercion ya estan dentro de funciones que salen temprano si
' ElementalSystemEnabled()=False, asi que este log nunca escribe con el motor
' elemental apagado aunque la clave nueva este en 1.
'
' Archivo: Logs\ElementalBalance_YYYY-MM-DD.log (uno por dia, mismo esquema de
' nombre que Venenos_*.log y Elemental_*.log).
'
' Esquema de fila (informe 2026-09-08-analisis-banco-de-balance.md seccion 6.4):
'   ts_ms;event;attacker;attacker_class;victim;victim_type;item;tier;
'   dmg_type;raw;final;resist_pct;pvp;map
'
' ts_ms          = milisegundos monotonicos, GetTickCountRaw() (timeGetTime()
'                  de winmm.dll, modElapsedTime.bas). NO es epoch: es un
'                  contador desde el boot del proceso, sirve para medir
'                  intervalos DENTRO de una misma corrida, no entre reinicios.
' event          = nombre corto del punto de insercion (ver tabla debajo).
' attacker       = "U"<UserIndex> o "N"<NpcIndex> del origen; "0" si el punto
'                  no tiene el atacante a mano (ver "Campos en 0" debajo).
' attacker_class = ListaClases(...) si el atacante es user; "0" si es NPC o
'                  desconocido.
' victim         = "U"<UserIndex> o "N"<NpcIndex> del blanco.
' victim_type    = "user" o "npc".
' item           = ObjIndex del arma/municion/orbe/encantamiento origen; 0 si
'                  el punto no recibe ese dato (ver "Campos en 0").
' tier           = SIEMPRE 0: el motor no tiene un campo/tipo de "tier" de
'                  item (verificado por busqueda en Declares.bas); ningun
'                  punto de insercion puede llenarlo.
' dmg_type       = e_ElementalDamageType (o DMG_TYPE_CRIT=15 para el critico
'                  universal) del componente/proc/tick logueado.
' raw            = dano antes de resistencia; 0 si el evento no tiene numero
'                  de dano (proc_apply_state) o si el valor pre-resistencia ya
'                  fue sobreescrito antes de este punto (thorns, universal_
'                  crit -- ver "Campos en 0"; ninguno de los dos es
'                  observable via wire por la sonda de todos modos, solo
'                  importa "final").
' final          = dano despues de resistencia (mismo criterio de 0 que raw).
' resist_pct     = CALCULADO ACA (no recibido como parametro): (raw-final)*100
'                  \ raw si raw>0, 0 si raw<=0. Centralizado para no repetir
'                  la cuenta -y su guard de division por cero- en 6 call
'                  sites distintos.
' pvp            = 1 si victim_type="user", 0 si es "npc". Definicion LITERAL
'                  pedida por el plan 08.003 (no es "es PvP real": un NPC
'                  atacando a un user tambien cuenta pvp=1). Calculado aca a
'                  partir de victimType, no recibido.
' map            = numero de mapa del victim (UserList/NpcList .pos.Map), 0 si
'                  el indice es invalido.
'
' Campos en 0 documentados por punto de insercion (no se cambian firmas de
' funciones existentes para llenarlos -- fuera de alcance de esta ola):
'   :367 ApplyDotTickResist        -> attacker/attacker_class/item en 0: la
'         funcion solo recibe datos del target (targetIsNpc/targetIndex/
'         dmgType/rawDamage), no conoce el origen del DoT.
'   :403 ResolveComponentsVsTarget -> attacker/attacker_class/item en 0:
'         funcion generica reusada por arma/municion/orbe/encantamientos
'         (ElementalDamageUserVsTarget) y por el path NPC->user; no recibe
'         UserIndex atacante ni el ObjIndex fuente.
'   :433 FireProcs (dmgBonus)      -> attacker/attacker_class SI disponibles
'         (parametros attackerIndex/attackerType de FireProcs); item en 0,
'         mismo motivo que :403.
'   :463 FireProcs (applyState)    -> attacker/attacker_class SI disponibles;
'         item en 0; raw/final en 0 porque este evento aplica un estado
'         (EffectsOverTime), no tiene un numero de dano propio.
'   :571 ApplyThornsDamage (thorns) -> item en 0 (FireSlotThorns SI tiene el
'         ObjIndex del slot pero no lo pasa a ResolveThorns/ApplyThornsDamage);
'         raw en 0 (el dano pre-reduccion se calcula y se pisa en
'         ResolveThorns, no llega a ApplyThornsDamage; ver nota "raw" arriba).
'         OJO roles invertidos: en ApplyThornsDamage, 'attackerIndex' es quien
'         RECIBE el reflejo (victim de este evento) y 'defenderIndex' es quien
'         lo origina (attacker de este evento) -- son los nombres del golpe
'         ORIGINAL que genero las espinas, no los de este evento puntual.
'   :1229 TryUniversalCrit (universal_crit) -> item en 0 (el critico universal
'         no esta atado a un item puntual); raw en 0 (el bonus pre-
'         resistencia se pisa en la misma variable antes de este punto).

Option Explicit

Private Const ELEMENTAL_BALANCE_LOG_HEADER As String = "ts_ms;event;attacker;attacker_class;victim;victim_type;item;tier;dmg_type;raw;final;resist_pct;pvp;map"

Public Function ElementalBalanceLogEnabled() As Boolean
    ElementalBalanceLogEnabled = IsFeatureEnabled("elemental_balance_log")
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
                                ByVal mapNumber As Integer)
    On Error GoTo ErrHandler
    If Not ElementalBalanceLogEnabled() Then Exit Sub
    Dim resistPct As Long
    If raw > 0 Then resistPct = (raw - finalDmg) * 100 \ raw
    Dim pvp As Byte
    If victimType = "user" Then pvp = 1
    Dim today As String
    today = Format$(Date, "yyyy-mm-dd")
    Dim fname As String
    fname = App.Path & "\Logs\ElementalBalance_" & today & ".log"
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
                 pvp & ";" & mapNumber
    Close #fnum
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalBalanceLog.LogElementalBalance", Erl)
    On Error Resume Next
    Close #fnum
End Sub
