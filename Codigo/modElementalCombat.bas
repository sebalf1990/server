Attribute VB_Name = "modElementalCombat"
' Argentum 20 Game Server
'
' modElementalCombat
' Sistema de danos elementales unificado (TOGGLE32 elemental_system). Plan 20.002 - Ola 0.
'
' Motor de la CAPA ADITIVA sobre el combate fisico (core intacto). Provee:
'   - Registro data-driven de tipos de dano (DamageTypes.dat) sobre e_DamageType.
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

' ============================================================================
' Toggle maestro
' ============================================================================
Public Function ElementalSystemEnabled() As Boolean
    ElementalSystemEnabled = IsFeatureEnabled("elemental_system")
End Function

' ============================================================================
' Registro de tipos de dano (data-driven)
' ============================================================================
Private Sub SetDamageTypeDefault(ByVal t As e_DamageType, ByVal nm As String, ByVal col As Long)
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
    If Not ElementalSystemEnabled() Then Exit Sub
    Dim fname As String
    fname = DatPath & "DamageTypes.dat"
    If LenB(dir(fname)) = 0 Then Exit Sub
    Dim Leer As New clsIniManager
    Call Leer.Initialize(fname)
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

Public Function DamageTypeName(ByVal t As e_DamageType) As String
    If t >= 1 And t <= MAX_DAMAGE_TYPE_ID Then
        If LenB(DamageTypeReg(t).nombre) > 0 Then
            DamageTypeName = DamageTypeReg(t).nombre
            Exit Function
        End If
    End If
    DamageTypeName = "Tipo" & CStr(t)
End Function

Public Function DamageTypeColor(ByVal t As e_DamageType) As Long
    If t >= 1 And t <= MAX_DAMAGE_TYPE_ID Then DamageTypeColor = DamageTypeReg(t).NumberColor
    If DamageTypeColor = 0 Then DamageTypeColor = vbWhite
End Function

Private Function ResistCapForType(ByVal t As e_DamageType) As Single
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

Private Sub AddSetResist(ByRef acc As t_ElementalResist, ByRef rs As t_ElementalResistSet, ByVal dmgType As e_DamageType)
    Dim k As Integer
    For k = 1 To rs.Count
        If rs.Resist(k).DamageType = dmgType Then
            Call AddResistEntry(acc, rs.Resist(k))
        End If
    Next k
End Sub

Public Function GetUserElementalResist(ByVal UserIndex As Integer, ByVal dmgType As e_DamageType) As t_ElementalResist
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

Public Function GetNpcElementalResist(ByVal NpcIndex As Integer, ByVal dmgType As e_DamageType) As t_ElementalResist
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

Private Function GetTargetResist(ByVal targetIsNpc As Boolean, ByVal targetIndex As Integer, ByVal dmgType As e_DamageType) As t_ElementalResist
    If targetIsNpc Then
        GetTargetResist = GetNpcElementalResist(targetIndex, dmgType)
    Else
        GetTargetResist = GetUserElementalResist(targetIndex, dmgType)
    End If
End Function

' Aplica la cascada de resistencia-a-dano a un numero. Devuelve el dano final (>=0).
Public Function ApplyElementalResist(ByVal RawDamage As Long, ByRef r As t_ElementalResist, ByVal dmgType As e_DamageType, ByRef outNullified As Boolean) As Long
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
Private Function FireProcs(ByRef src As t_ElementalSource, ByVal trig As e_ProcTrigger, ByVal targetIsNpc As Boolean, ByVal targetIndex As Integer, ByVal logCtx As String) As Long
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
                        ' Ola 1: aplicar el preset de efecto (EotId) sobre el motor EOT.
                        Call ElementalLog(logCtx & " PROC applyState EotId=" & src.Proc(i).EotId & " [stub Ola0]")
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
Public Function ElementalDamageUserVsNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal WeaponObjIndex As Integer, ByVal MunitionObjIndex As Integer) As Long
    On Error GoTo ErrHandler
    If Not ElementalSystemEnabled() Then Exit Function
    If UserIndex <= 0 Or NpcIndex <= 0 Then Exit Function
    Dim total As Long
    Dim ctx As String
    ctx = "U" & UserIndex & "->N" & NpcIndex
    ' Componentes + procs onHit del arma
    If WeaponObjIndex > 0 Then
        total = total + ResolveComponentsVsTarget(ObjData(WeaponObjIndex).Elemental, True, NpcIndex, ctx & " weap")
        total = total + FireProcs(ObjData(WeaponObjIndex).Elemental, eProcOnHit, True, NpcIndex, ctx & " weap")
    End If
    ' Municion (rango): suma sus componentes/procs
    If MunitionObjIndex > 0 And WeaponObjIndex > 0 Then
        If ObjData(WeaponObjIndex).Proyectil > 0 Then
            total = total + ResolveComponentsVsTarget(ObjData(MunitionObjIndex).Elemental, True, NpcIndex, ctx & " ammo")
            total = total + FireProcs(ObjData(MunitionObjIndex).Elemental, eProcOnHit, True, NpcIndex, ctx & " ammo")
        End If
    End If
    ' Procs onDamaged del NPC defensor (thorns/aura). Ola 0: fire + log; el HP sink al
    ' atacante se conecta cuando se enganche el camino user-como-target (ola siguiente).
    Dim t As Integer
    t = NpcList(NpcIndex).Numero
    If t > 0 Then
        Dim retal As Long
        retal = FireProcs(NpcInfoCache(t).Elemental, eProcOnDamaged, False, UserIndex, ctx & " thorns")
        If retal > 0 Then Call ElementalLog(ctx & " thorns retaliation=" & retal & " [HP sink Ola1]")
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
