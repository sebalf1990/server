Attribute VB_Name = "modElementalCombat"
' Argentum 20 Game Server
'
' modElementalCombat
' Sistema de danos elementales unificado (TOGGLE32 elemental_system). Plan 20.002 - Ola 0.
'
' Motor de la CAPA ADITIVA sobre el combate fisico (core intacto). Provee:
'   - Registro data-driven de tipos de dano (DamageTypes.dat) sobre e_ElementalDamageType.
'   - Roll de componentes de dano tipado.
'   - Resistencia generica por clave (2 sabores): resist-a-dano (chance/flat/pct + cap + inmune)
'     y resist-a-efecto (Ola 1).
'   - Motor de procs on-hit: atacante (alGolpear) y defensor (alSerGolpeado), kind dmgBonus/applyState.
'   - Punto de entrada gateado para el camino user->NPC (unico touchpoint de la Ola 0).
'   - Log de efectos (Logs\Elemental_YYYY-MM-DD.log) y helper de inspeccion para GM.
'
' OFF (toggle en 0) = no se ejecuta NADA: combate fisico vanilla identico. No hay fallback legacy.

Option Explicit

' Registro de tipos de dano (la "clave"). Cargado al boot por LoadDamageTypes.
Public DamageTypeReg(1 To MAX_DAMAGE_TYPE_ID) As t_DamageTypeInfo
Public DamageTypeRegLoaded As Boolean

' Crit universal (plan 20.002 Ola 2). Config data-driven desde DamageTypes.dat [INIT].
Public Const DMG_TYPE_CRIT As Long = 15   ' clave de resistencia para crit (fisico+magico)
Private mUniversalCritChance As Single
Private mUniversalCritMult As Single

' Control / cap de slow (plan 20.002 Ola 3). Cacheado al boot (hot path de velocidad).
Public ElementalSlowCapActive As Boolean
Private mElementalSlowFloor As Single

' ============================================================================
' Toggle maestro
' ============================================================================
Public Function ElementalSystemEnabled() As Boolean
    ElementalSystemEnabled = IsFeatureEnabled("elemental_system")
End Function

' ============================================================================
' Registro de tipos de dano (data-driven)
' ============================================================================
Private Sub SetDamageTypeDefault(ByVal t As e_ElementalDamageType, ByVal nm As String, ByVal col As Long)
    If t < 1 Or t > MAX_DAMAGE_TYPE_ID Then Exit Sub
    DamageTypeReg(t).nombre = nm
    DamageTypeReg(t).NumberColor = col
    DamageTypeReg(t).ResistCapPct = 0
    DamageTypeReg(t).DefaultParticle = 0
    DamageTypeReg(t).TagElementalRelacionado = 0
    DamageTypeReg(t).Defined = 1
End Sub

Public Sub InitDamageTypeRegistryDefaults()
    Dim i As Long
    For i = 1 To MAX_DAMAGE_TYPE_ID
        DamageTypeReg(i).nombre = ""
        DamageTypeReg(i).NumberColor = vbWhite
        DamageTypeReg(i).DefaultParticle = 0
        DamageTypeReg(i).ResistCapPct = 0
        DamageTypeReg(i).TagElementalRelacionado = 0
        DamageTypeReg(i).Defined = 0
    Next i
    ' Built-in (sirven aun sin DamageTypes.dat). Luz/Oscuridad reservados al oficial.
    Call SetDamageTypeDefault(eDmgPhysical, "Fisico", vbWhite)
    Call SetDamageTypeDefault(eDmgFire, "Fuego", RGB(255, 140, 0))
    Call SetDamageTypeDefault(eDmgFrost, "Escarcha", RGB(120, 200, 255))
    Call SetDamageTypeDefault(eDmgPoison, "Veneno", RGB(120, 200, 60))
    Call SetDamageTypeDefault(eDmgAcid, "Acido", RGB(170, 230, 40))
    Call SetDamageTypeDefault(eDmgArcane, "Arcano", RGB(200, 120, 255))
    Call SetDamageTypeDefault(eDmgBleed, "Sangrado", RGB(200, 30, 30))
End Sub

' Carga DamageTypes.dat (metadata tuneable). Tolerante: si no existe, usa los defaults.
Public Sub LoadDamageTypes()
    On Error GoTo ErrHandler
    Call InitDamageTypeRegistryDefaults
    DamageTypeRegLoaded = True
    mUniversalCritChance = 5
    mUniversalCritMult = 0.5
    DamageTypeReg(DMG_TYPE_CRIT).nombre = "Critico"
    DamageTypeReg(DMG_TYPE_CRIT).NumberColor = RGB(255, 165, 0)
    ElementalSlowCapActive = ElementalSystemEnabled()
    mElementalSlowFloor = 0.45
    If Not ElementalSystemEnabled() Then Exit Sub
    Dim fname As String
    fname = DatPath & "DamageTypes.dat"
    If LenB(dir(fname)) = 0 Then Exit Sub
    Dim Leer As New clsIniManager
    Call Leer.Initialize(fname)
    If val(Leer.GetValue("INIT", "CritChancePct")) > 0 Then mUniversalCritChance = val(Leer.GetValue("INIT", "CritChancePct"))
    If val(Leer.GetValue("INIT", "CritMult")) > 0 Then mUniversalCritMult = val(Leer.GetValue("INIT", "CritMult"))
    If val(Leer.GetValue("INIT", "SlowFloorPct")) > 0 Then mElementalSlowFloor = val(Leer.GetValue("INIT", "SlowFloorPct"))
    Dim count As Long, i As Long, sect As String, idv As Long
    count = val(Leer.GetValue("INIT", "TypeCount"))
    For i = 1 To count
        sect = "TYPE" & i
        idv = val(Leer.GetValue(sect, "Id"))
        If idv >= 1 And idv <= MAX_DAMAGE_TYPE_ID Then
            With DamageTypeReg(idv)
                If LenB(Leer.GetValue(sect, "Name")) > 0 Then .nombre = Leer.GetValue(sect, "Name")
                .NumberColor = val(Leer.GetValue(sect, "NumberColor"))
                .DefaultParticle = val(Leer.GetValue(sect, "DefaultParticle"))
                .ResistCapPct = val(Leer.GetValue(sect, "ResistCapPct"))
                .TagElementalRelacionado = val(Leer.GetValue(sect, "TagElementalRelacionado"))
                .Defined = 1
            End With
        End If
    Next i
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.LoadDamageTypes", Erl)
End Sub

Public Function DamageTypeName(ByVal t As e_ElementalDamageType) As String
    If t >= 1 And t <= MAX_DAMAGE_TYPE_ID Then
        If LenB(DamageTypeReg(t).nombre) > 0 Then
            DamageTypeName = DamageTypeReg(t).nombre
            Exit Function
        End If
    End If
    DamageTypeName = "Tipo" & CStr(t)
End Function

Public Function DamageTypeColor(ByVal t As e_ElementalDamageType) As Long
    If t >= 1 And t <= MAX_DAMAGE_TYPE_ID Then DamageTypeColor = DamageTypeReg(t).NumberColor
    If DamageTypeColor = 0 Then DamageTypeColor = vbWhite
End Function

Private Function ResistCapForType(ByVal t As e_ElementalDamageType) As Single
    Dim cap As Single
    If t >= 1 And t <= MAX_DAMAGE_TYPE_ID Then cap = DamageTypeReg(t).ResistCapPct
    If cap <= 0 Then cap = ELEMENTAL_RESIST_CAP_DEFAULT
    ResistCapForType = cap
End Function

' ============================================================================
' Roll de dano
' ============================================================================
Public Function RollDamageComponent(ByRef c As t_DamageComponent) As Long
    Dim lo As Long, hi As Long
    lo = c.MinDamage
    hi = c.MaxDamage
    If lo < 0 Then lo = 0
    If hi < lo Then hi = lo
    If hi <= lo Then
        RollDamageComponent = lo
    Else
        RollDamageComponent = RandomNumber(lo, hi)
    End If
End Function

' ============================================================================
' Resistencia: agregacion por entidad + aplicacion con cap
' ============================================================================
Private Sub AddResistEntry(ByRef acc As t_ElementalResist, ByRef e As t_ElementalResist)
    acc.ReduceChancePct = acc.ReduceChancePct + e.ReduceChancePct
    acc.ReduceFlat = acc.ReduceFlat + e.ReduceFlat
    acc.ReducePct = acc.ReducePct + e.ReducePct
    If e.Immune <> 0 Then acc.Immune = 1
    acc.ReduceEffectMagnitudePct = acc.ReduceEffectMagnitudePct + e.ReduceEffectMagnitudePct
    acc.ReduceEffectDurationPct = acc.ReduceEffectDurationPct + e.ReduceEffectDurationPct
    acc.ReduceEffectChancePct = acc.ReduceEffectChancePct + e.ReduceEffectChancePct
End Sub

Private Sub AddSetResist(ByRef acc As t_ElementalResist, ByRef rs As t_ElementalResistSet, ByVal dmgType As e_ElementalDamageType)
    Dim k As Integer
    For k = 1 To rs.Count
        If rs.Resist(k).DamageType = dmgType Then
            Call AddResistEntry(acc, rs.Resist(k))
        End If
    Next k
End Sub

Public Function GetUserElementalResist(ByVal UserIndex As Integer, ByVal dmgType As e_ElementalDamageType) As t_ElementalResist
    Dim acc As t_ElementalResist
    acc.DamageType = dmgType
    If UserIndex > 0 Then
        With UserList(UserIndex).invent
            Call AddSetResist(acc, ObjData(.EquippedArmorObjIndex).ElementalResist, dmgType)
            Call AddSetResist(acc, ObjData(.EquippedHelmetObjIndex).ElementalResist, dmgType)
            Call AddSetResist(acc, ObjData(.EquippedShieldObjIndex).ElementalResist, dmgType)
            Call AddSetResist(acc, ObjData(.EquippedRingAccesoryObjIndex).ElementalResist, dmgType)
            Call AddSetResist(acc, ObjData(.EquippedAmuletAccesoryObjIndex).ElementalResist, dmgType)
            Call AddSetResist(acc, ObjData(.EquippedBackpackObjIndex).ElementalResist, dmgType)
            Call AddSetResist(acc, ObjData(.EquippedSaddleObjIndex).ElementalResist, dmgType)
        End With
    End If
    GetUserElementalResist = acc
End Function

Public Function GetNpcElementalResist(ByVal NpcIndex As Integer, ByVal dmgType As e_ElementalDamageType) As t_ElementalResist
    Dim acc As t_ElementalResist
    acc.DamageType = dmgType
    If NpcIndex > 0 Then
        Dim t As Integer
        t = NpcList(NpcIndex).Numero
        If t > 0 Then
            Call AddSetResist(acc, NpcInfoCache(t).ElementalResist, dmgType)
        End If
    End If
    GetNpcElementalResist = acc
End Function

Private Function GetTargetResist(ByVal targetIsNpc As Boolean, ByVal targetIndex As Integer, ByVal dmgType As e_ElementalDamageType) As t_ElementalResist
    If targetIsNpc Then
        GetTargetResist = GetNpcElementalResist(targetIndex, dmgType)
    Else
        GetTargetResist = GetUserElementalResist(targetIndex, dmgType)
    End If
End Function

' Aplica la cascada de resistencia-a-dano a un numero. Devuelve el dano final (>=0).
Public Function ApplyElementalResist(ByVal RawDamage As Long, ByRef r As t_ElementalResist, ByVal dmgType As e_ElementalDamageType, ByRef outNullified As Boolean) As Long
    outNullified = False
    If RawDamage <= 0 Then Exit Function
    ' Inmunidad: bloqueo absoluto
    If r.Immune <> 0 Then
        outNullified = True
        Exit Function
    End If
    ' Chance de anular el componente entero
    If r.ReduceChancePct > 0 Then
        Dim ch As Long
        ch = r.ReduceChancePct
        If ch > 100 Then ch = 100
        If RandomNumber(1, 100) <= ch Then
            outNullified = True
            Exit Function
        End If
    End If
    ' Cascada: flat -> pct (capeado)
    Dim dmg As Long
    dmg = RawDamage - r.ReduceFlat
    If dmg <= 0 Then Exit Function
    Dim pct As Single
    pct = r.ReducePct
    Dim cap As Single
    cap = ResistCapForType(dmgType)
    If pct > cap Then pct = cap
    If pct < 0 Then pct = 0
    dmg = Int(CDbl(dmg) * (1# - pct))
    If dmg < 0 Then dmg = 0
    ApplyElementalResist = dmg
End Function

' ============================================================================
' Resolucion de componentes y procs
' ============================================================================
Private Function ResolveComponentsVsTarget(ByRef src As t_ElementalSource, ByVal targetIsNpc As Boolean, ByVal targetIndex As Integer, ByVal logCtx As String) As Long
    Dim total As Long, i As Integer
    For i = 1 To src.CompCount
        If src.Comp(i).DamageType >= 1 And src.Comp(i).DamageType <= MAX_DAMAGE_TYPE_ID Then
            Dim raw As Long
            raw = RollDamageComponent(src.Comp(i))
            Dim r As t_ElementalResist
            r = GetTargetResist(targetIsNpc, targetIndex, src.Comp(i).DamageType)
            Dim nullified As Boolean, finalDmg As Long
            finalDmg = ApplyElementalResist(raw, r, src.Comp(i).DamageType, nullified)
            total = total + finalDmg
            Call ElementalLog(logCtx & " comp " & DamageTypeName(src.Comp(i).DamageType) & " raw=" & raw & " final=" & finalDmg & IIf(nullified, " [nullified]", "") & IIf(r.Immune <> 0, " [immune]", ""))
        End If
    Next i
    ResolveComponentsVsTarget = total
End Function

' Dispara los procs de una fuente con el trigger dado contra un target.
' Devuelve el dano extra (kind=dmgBonus) ya reducido por la resistencia del target.
Private Function FireProcs(ByRef src As t_ElementalSource, ByVal trig As e_ProcTrigger, ByVal targetIsNpc As Boolean, ByVal targetIndex As Integer, ByVal attackerIndex As Integer, ByVal attackerType As e_ReferenceType, ByVal logCtx As String) As Long
    Dim total As Long, i As Integer
    For i = 1 To src.ProcCount
        If src.Proc(i).Trigger = trig Then
            Dim ch As Long
            ch = src.Proc(i).ChancePct
            If ch < 0 Then ch = 0
            If ch > 100 Then ch = 100
            If ch >= 100 Or RandomNumber(1, 100) <= ch Then
                Select Case src.Proc(i).Kind
                    Case eProcDamageBonus
                        Dim c As t_DamageComponent
                        c.DamageType = src.Proc(i).DamageType
                        c.MinDamage = src.Proc(i).MinDamage
                        c.MaxDamage = src.Proc(i).MaxDamage
                        Dim raw As Long
                        raw = RollDamageComponent(c)
                        Dim r As t_ElementalResist
                        r = GetTargetResist(targetIsNpc, targetIndex, c.DamageType)
                        Dim nul As Boolean, fd As Long
                        fd = ApplyElementalResist(raw, r, c.DamageType, nul)
                        total = total + fd
                        Call ElementalLog(logCtx & " PROC dmgBonus " & DamageTypeName(c.DamageType) & " final=" & fd)
                    Case eProcApplyState
                        ' Aplica el preset (EotId) respetando inmunidad / resist-a-efecto del tipo del proc.
                        If src.Proc(i).EotId > 0 Then
                            Dim rr As t_ElementalResist
                            rr = GetTargetResist(targetIsNpc, targetIndex, src.Proc(i).DamageType)
                            If rr.Immune <> 0 Then
                                Call ElementalLog(logCtx & " PROC applyState BLOQUEADO (inmune " & DamageTypeName(src.Proc(i).DamageType) & ")")
                            ElseIf rr.ReduceEffectChancePct > 0 And RandomNumber(1, 100) <= rr.ReduceEffectChancePct Then
                                Call ElementalLog(logCtx & " PROC applyState resistido (chance efecto)")
                            Else
                                Dim trt As e_ReferenceType
                                If targetIsNpc Then trt = eNpc Else trt = eUser
                                Call EffectsOverTime.CreateEffect(attackerIndex, attackerType, targetIndex, trt, src.Proc(i).EotId)
                                If attackerType = eUser And LenB(EffectOverTime(src.Proc(i).EotId).ApplyMsg) > 0 Then Call WriteConsoleMsg(attackerIndex, EffectOverTime(src.Proc(i).EotId).ApplyMsg, e_FontTypeNames.FONTTYPE_FIGHT)
                                Call ElementalLog(logCtx & " PROC applyState EotId=" & src.Proc(i).EotId & " aplicado")
                            End If
                        Else
                            Call ElementalLog(logCtx & " PROC applyState sin EotId (ignorado)")
                        End If
                End Select
            End If
        End If
    Next i
    FireProcs = total
End Function

' ============================================================================
' Punto de entrada gateado: camino user -> NPC (unico touchpoint de la Ola 0)
' Devuelve el dano elemental EXTRA a sumar (ya resistido). NO toca el dano fisico.
' ============================================================================
Public Function ElementalDamageUserVsNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal WeaponObjIndex As Integer, ByVal MunitionObjIndex As Integer, ByRef outColor As Long) As Long
    On Error GoTo ErrHandler
    outColor = vbWhite
    If Not ElementalSystemEnabled() Then Exit Function
    If UserIndex <= 0 Or NpcIndex <= 0 Then Exit Function
    Dim total As Long
    Dim ctx As String
    ctx = "U" & UserIndex & "->N" & NpcIndex
    ' Componentes + procs onHit del arma
    If WeaponObjIndex > 0 Then
        total = total + ResolveComponentsVsTarget(ObjData(WeaponObjIndex).Elemental, True, NpcIndex, ctx & " weap")
        total = total + FireProcs(ObjData(WeaponObjIndex).Elemental, eProcOnHit, True, NpcIndex, UserIndex, eUser, ctx & " weap")
    End If
    ' Municion (rango): suma sus componentes/procs solo si el arma es de proyectil
    If MunitionObjIndex > 0 And WeaponObjIndex > 0 Then
        If ObjData(WeaponObjIndex).Proyectil > 0 Then
            total = total + ResolveComponentsVsTarget(ObjData(MunitionObjIndex).Elemental, True, NpcIndex, ctx & " ammo")
            total = total + FireProcs(ObjData(MunitionObjIndex).Elemental, eProcOnHit, True, NpcIndex, UserIndex, eUser, ctx & " ammo")
        End If
    End If
    ' Encantamiento temporal del arma (hechizo Encantar Arma). Aditivo al elemental base.
    With UserList(UserIndex).flags
        If .EnchantWeaponObjIndex > 0 And .EnchantWeaponObjIndex = WeaponObjIndex Then
            If .EnchantWeaponPermanent = 1 Or Not DeadlinePassed(GetTickCountRaw(), .EnchantWeaponDeadline) Then
                total = total + ResolveComponentsVsTarget(.EnchantWeaponSource, True, NpcIndex, ctx & " ench")
                total = total + FireProcs(.EnchantWeaponSource, eProcOnHit, True, NpcIndex, UserIndex, eUser, ctx & " ench")
            End If
        End If
    End With
    ' Procs onDamaged del NPC defensor (thorns/aura): aplican efecto al atacante (user).
    ' El dano dmgBonus de retaliacion se loguea (HP sink al user en ola posterior).
    Dim t As Integer
    t = NpcList(NpcIndex).Numero
    If t > 0 Then
        Dim retal As Long
        retal = FireProcs(NpcInfoCache(t).Elemental, eProcOnDamaged, False, UserIndex, NpcIndex, eNpc, ctx & " thorns")
        If retal > 0 Then Call ElementalLog(ctx & " thorns retaliation=" & retal & " [HP sink Ola1]")
    End If
    ' Color del numero elemental: tipo primario del arma base; si no tiene, del encantamiento.
    If WeaponObjIndex > 0 Then
        If ObjData(WeaponObjIndex).Elemental.CompCount > 0 Then
            outColor = DamageTypeColor(ObjData(WeaponObjIndex).Elemental.Comp(1).DamageType)
        ElseIf UserList(UserIndex).flags.EnchantWeaponObjIndex = WeaponObjIndex And UserList(UserIndex).flags.EnchantWeaponSource.CompCount > 0 Then
            outColor = DamageTypeColor(UserList(UserIndex).flags.EnchantWeaponSource.Comp(1).DamageType)
        End If
    End If
    ElementalDamageUserVsNpc = total
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.ElementalDamageUserVsNpc", Erl)
End Function

' ============================================================================
' Tooling: log de efectos + inspeccion de resistencias
' ============================================================================
Public Sub ElementalLog(ByVal msg As String)
    On Error GoTo ErrHandler
    If Not ElementalSystemEnabled() Then Exit Sub
    Dim fname As String
    fname = App.Path & "\Logs\Elemental_" & Format$(Date, "yyyy-mm-dd") & ".log"
    Dim fnum As Integer
    fnum = FreeFile
    Open fname For Append As #fnum
    Print #fnum, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & msg
    Close #fnum
    Exit Sub
ErrHandler:
    On Error Resume Next
    Close #fnum
End Sub

' Devuelve un texto con la resistencia efectiva del target para cada tipo built-in.
Public Function DescribeEntityResist(ByVal targetIsNpc As Boolean, ByVal targetIndex As Integer) As String
    Dim s As String, t As Long
    Dim r As t_ElementalResist
    For t = eDmgPhysical To eDmgBleed
        r = GetTargetResist(targetIsNpc, targetIndex, t)
        If r.Immune <> 0 Or r.ReduceChancePct <> 0 Or r.ReduceFlat <> 0 Or r.ReducePct <> 0 Then
            s = s & DamageTypeName(t) & ": "
            If r.Immune <> 0 Then
                s = s & "INMUNE"
            Else
                s = s & "chance=" & r.ReduceChancePct & "% flat=" & r.ReduceFlat & " pct=" & Format$(r.ReducePct, "0.00") & " (cap " & Format$(ResistCapForType(t), "0.00") & ")"
            End If
            s = s & "  "
        End If
    Next t
    If LenB(s) = 0 Then s = "sin resistencias elementales"
    DescribeEntityResist = s
End Function

' ============================================================================
' Parsers de datos (esquema centralizado; lo llaman FileIO/MODULO_NPCs)
' ============================================================================
Public Sub ParseElementalSourceFromIni(ByRef Leer As clsIniManager, ByVal sect As String, ByRef src As t_ElementalSource)
    Dim n As Integer, i As Integer, p As String
    src.CompCount = 0
    src.ProcCount = 0
    n = val(Leer.GetValue(sect, "ElemCompCount"))
    If n > MAX_DAMAGE_COMPONENTS Then n = MAX_DAMAGE_COMPONENTS
    If n < 0 Then n = 0
    For i = 1 To n
        p = "ElemComp" & i
        src.Comp(i).DamageType = val(Leer.GetValue(sect, p & "Type"))
        src.Comp(i).MinDamage = val(Leer.GetValue(sect, p & "Min"))
        src.Comp(i).MaxDamage = val(Leer.GetValue(sect, p & "Max"))
    Next i
    src.CompCount = n
    n = val(Leer.GetValue(sect, "ElemProcCount"))
    If n > MAX_ELEMENTAL_PROCS Then n = MAX_ELEMENTAL_PROCS
    If n < 0 Then n = 0
    For i = 1 To n
        p = "ElemProc" & i
        src.Proc(i).ChancePct = val(Leer.GetValue(sect, p & "Chance"))
        src.Proc(i).Kind = val(Leer.GetValue(sect, p & "Kind"))
        src.Proc(i).Trigger = val(Leer.GetValue(sect, p & "Trigger"))
        src.Proc(i).DamageType = val(Leer.GetValue(sect, p & "Type"))
        src.Proc(i).MinDamage = val(Leer.GetValue(sect, p & "Min"))
        src.Proc(i).MaxDamage = val(Leer.GetValue(sect, p & "Max"))
        src.Proc(i).EotId = val(Leer.GetValue(sect, p & "Eot"))
    Next i
    src.ProcCount = n
End Sub

Public Sub ParseElementalResistFromIni(ByRef Leer As clsIniManager, ByVal sect As String, ByRef rs As t_ElementalResistSet)
    Dim n As Integer, i As Integer, p As String
    rs.Count = 0
    n = val(Leer.GetValue(sect, "ElemResistCount"))
    If n > MAX_ELEMENTAL_RESISTS Then n = MAX_ELEMENTAL_RESISTS
    If n < 0 Then n = 0
    For i = 1 To n
        p = "ElemResist" & i
        rs.Resist(i).DamageType = val(Leer.GetValue(sect, p & "Type"))
        rs.Resist(i).ReduceChancePct = val(Leer.GetValue(sect, p & "Chance"))
        rs.Resist(i).ReduceFlat = val(Leer.GetValue(sect, p & "Flat"))
        rs.Resist(i).ReducePct = val(Leer.GetValue(sect, p & "Pct"))
        rs.Resist(i).Immune = val(Leer.GetValue(sect, p & "Immune"))
        rs.Resist(i).ReduceEffectMagnitudePct = val(Leer.GetValue(sect, p & "EffMag"))
        rs.Resist(i).ReduceEffectDurationPct = val(Leer.GetValue(sect, p & "EffDur"))
        rs.Resist(i).ReduceEffectChancePct = val(Leer.GetValue(sect, p & "EffChance"))
    Next i
    rs.Count = n
End Sub

' ============================================================================
' Crit universal (Ola 2): fallback para clases sin firma propia (bandido/asesino).
' Resistencia a crit via clave DMG_TYPE_CRIT (reduce chance y dano; inmune = sin crit).
' ============================================================================
Public Function UniversalCritActive() As Boolean
    UniversalCritActive = ElementalSystemEnabled()
End Function

' Devuelve el dano EXTRA de crit (0 si no critea). baseDamage = dano base del golpe.
Public Function TryUniversalCrit(ByVal UserIndex As Integer, ByVal targetIsNpc As Boolean, ByVal targetIndex As Integer, ByVal baseDamage As Long, ByRef outColor As Long) As Long
    On Error GoTo eh
    outColor = DamageTypeColor(DMG_TYPE_CRIT)
    If Not ElementalSystemEnabled() Then Exit Function
    If baseDamage <= 0 Then Exit Function
    Dim cr As t_ElementalResist
    cr = GetTargetResist(targetIsNpc, targetIndex, DMG_TYPE_CRIT)
    If cr.Immune <> 0 Then Exit Function
    Dim chance As Single
    chance = mUniversalCritChance - cr.ReduceChancePct
    If chance <= 0 Then Exit Function
    If chance > 100 Then chance = 100
    If RandomNumber(1, 100) <= chance Then
        Dim bonus As Long
        bonus = Int(CDbl(baseDamage) * mUniversalCritMult)
        Dim nul As Boolean
        bonus = ApplyElementalResist(bonus, cr, DMG_TYPE_CRIT, nul)
        If bonus < 0 Then bonus = 0
        TryUniversalCrit = bonus
        Call ElementalLog("U" & UserIndex & " universal crit base=" & baseDamage & " bonus=" & bonus)
    End If
    Exit Function
eh:
    Call TraceError(Err.Number, Err.Description, "modElementalCombat.TryUniversalCrit", Erl)
End Function

' ============================================================================
' Cap de control (Ola 3): nunca congelar por slow. Devuelve el multiplicador de
' velocidad (1+mod) con el slow capeado. Toggle OFF = vanilla (floor 0 = max(0,...)).
' El freeze duro sigue siendo por flag Paralizado/Inmovilizado, aparte de esto.
' ============================================================================
Public Function CappedSpeedMult(ByVal moveSpeedMod As Single) As Single
    Dim mult As Single
    mult = 1 + moveSpeedMod
    Dim floorVal As Single
    If ElementalSlowCapActive Then
        floorVal = mElementalSlowFloor
    Else
        floorVal = 0
    End If
    If mult < floorVal Then mult = floorVal
    CappedSpeedMult = mult
End Function
