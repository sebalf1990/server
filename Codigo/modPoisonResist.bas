Attribute VB_Name = "modPoisonResist"
' Argentum 20 Game Server
'
' modPoisonResist
' Sistema de venenos (TOGGLE26 new_poison_system).
' Suma de resistencias por familia desde items equipados.
' Familias: 1=Menor, 2=Hemo, 3=Neuro.
'
' Resistencias:
'   - ChancePct: resta lineal al ChanceAplicarPct del agresor (cualquier fuente).
'   - DanoFlat (Min/Max): puntos absolutos restados antes del %, random en rango.
'   - DanoPct (Min/Max): % aplicado despues del flat, random en rango.
'   - Inmunidad: bloqueo total desde cualquier fuente.

Option Explicit

Public Type t_PoisonResist
    ChancePct As Long
    DanoMinFlat As Long
    DanoMaxFlat As Long
    DanoMinPct As Long
    DanoMaxPct As Long
    Inmune As Byte
End Type

' Suma las resistencias de un item al acumulador (segun familia).
Private Sub AddItemResist(ByRef acc As t_PoisonResist, ByVal ObjIndex As Integer, ByVal familia As Byte)
    If ObjIndex <= 0 Then Exit Sub
    Select Case familia
        Case 1
            acc.ChancePct = acc.ChancePct + ObjData(ObjIndex).ResistChanceVenenoMenorPct
            acc.DanoMinFlat = acc.DanoMinFlat + ObjData(ObjIndex).ResistDanoVenenoMenorMinFlat
            acc.DanoMaxFlat = acc.DanoMaxFlat + ObjData(ObjIndex).ResistDanoVenenoMenorMaxFlat
            acc.DanoMinPct = acc.DanoMinPct + ObjData(ObjIndex).ResistDanoVenenoMenorMinPct
            acc.DanoMaxPct = acc.DanoMaxPct + ObjData(ObjIndex).ResistDanoVenenoMenorMaxPct
            If ObjData(ObjIndex).InmunidadVenenoMenor <> 0 Then acc.Inmune = 1
        Case 2
            acc.ChancePct = acc.ChancePct + ObjData(ObjIndex).ResistChanceHemoPct
            acc.DanoMinFlat = acc.DanoMinFlat + ObjData(ObjIndex).ResistDanoHemoMinFlat
            acc.DanoMaxFlat = acc.DanoMaxFlat + ObjData(ObjIndex).ResistDanoHemoMaxFlat
            acc.DanoMinPct = acc.DanoMinPct + ObjData(ObjIndex).ResistDanoHemoMinPct
            acc.DanoMaxPct = acc.DanoMaxPct + ObjData(ObjIndex).ResistDanoHemoMaxPct
            If ObjData(ObjIndex).InmunidadHemo <> 0 Then acc.Inmune = 1
        Case 3
            acc.ChancePct = acc.ChancePct + ObjData(ObjIndex).ResistChanceNeuroPct
            acc.DanoMinFlat = acc.DanoMinFlat + ObjData(ObjIndex).ResistDanoNeuroMinFlat
            acc.DanoMaxFlat = acc.DanoMaxFlat + ObjData(ObjIndex).ResistDanoNeuroMaxFlat
            acc.DanoMinPct = acc.DanoMinPct + ObjData(ObjIndex).ResistDanoNeuroMinPct
            acc.DanoMaxPct = acc.DanoMaxPct + ObjData(ObjIndex).ResistDanoNeuroMaxPct
            If ObjData(ObjIndex).InmunidadNeuro <> 0 Then acc.Inmune = 1
    End Select
End Sub

' Devuelve la resistencia total de un usuario para la familia dada.
' Suma aditiva de los 7 slots equipables relevantes.
Public Function GetUserPoisonResist(ByVal UserIndex As Integer, ByVal familia As Byte) As t_PoisonResist
    Dim r As t_PoisonResist
    If UserIndex <= 0 Then
        GetUserPoisonResist = r
        Exit Function
    End If
    With UserList(UserIndex).invent
        Call AddItemResist(r, .EquippedArmorObjIndex, familia)
        Call AddItemResist(r, .EquippedHelmetObjIndex, familia)
        Call AddItemResist(r, .EquippedShieldObjIndex, familia)
        Call AddItemResist(r, .EquippedRingAccesoryObjIndex, familia)
        Call AddItemResist(r, .EquippedAmuletAccesoryObjIndex, familia)
        Call AddItemResist(r, .EquippedBackpackObjIndex, familia)
        Call AddItemResist(r, .EquippedSaddleObjIndex, familia)
    End With
    GetUserPoisonResist = r
End Function

' Aplica la cascada de reduccion al dano por tick. Devuelve el dano final (>=0).
Public Function ApplyDamageResist(ByVal RawDamage As Long, ByRef r As t_PoisonResist) As Long
    Dim dmg As Long
    dmg = RawDamage
    ' Cascada 1: flat absoluto
    If r.DanoMaxFlat > 0 Then
        Dim flat As Long
        If r.DanoMinFlat >= r.DanoMaxFlat Then
            flat = r.DanoMinFlat
        Else
            flat = RandomNumber(r.DanoMinFlat, r.DanoMaxFlat)
        End If
        dmg = dmg - flat
    End If
    If dmg <= 0 Then
        ApplyDamageResist = 0
        Exit Function
    End If
    ' Cascada 2: porcentaje
    If r.DanoMaxPct > 0 Then
        Dim pct As Long
        If r.DanoMinPct >= r.DanoMaxPct Then
            pct = r.DanoMinPct
        Else
            pct = RandomNumber(r.DanoMinPct, r.DanoMaxPct)
        End If
        If pct >= 100 Then
            dmg = 0
        Else
            dmg = (dmg * (100 - pct)) \ 100
        End If
    End If
    If dmg < 0 Then dmg = 0
    ApplyDamageResist = dmg
End Function

' === Resistencia de NPCs (Fase post-cierre, retomada como diferido) ===
' Devuelve la resistencia de un NPC a la familia dada, leyendo de NpcInfoCache (no del NpcList).
' Modelo simetrico al de items equipables pero sin min-max (NPCs no se acumulan):
'   - ChancePct: chance de safar de quedar envenenado al aplicar.
'   - DanoMinFlat = ResistDano*Flat (resta absoluta deterministica).
'   - DanoMaxFlat = mismo valor (no hay rango aleatorio).
'   - DanoMinPct = DanoMaxPct = ResistDano*Pct (mismo).
'   - Inmune: bloqueo total al aplicar.
Public Function GetNpcPoisonResist(ByVal NpcIndex As Integer, ByVal familia As Byte) As t_PoisonResist
    Dim r As t_PoisonResist
    If NpcIndex <= 0 Then
        GetNpcPoisonResist = r
        Exit Function
    End If
    Dim npcType As Integer
    npcType = NpcList(NpcIndex).Numero
    If npcType <= 0 Then
        GetNpcPoisonResist = r
        Exit Function
    End If
    With NpcInfoCache(npcType)
        ' 1) Resistencia especifica por familia
        Select Case familia
            Case 1
                r.ChancePct = .ResistChanceVenenoMenorPct
                r.DanoMinFlat = .ResistDanoVenenoMenorFlat
                r.DanoMaxFlat = .ResistDanoVenenoMenorFlat
                r.DanoMinPct = .ResistDanoVenenoMenorPct
                r.DanoMaxPct = .ResistDanoVenenoMenorPct
                If .InmunidadVenenoMenor <> 0 Then r.Inmune = 1
            Case 2
                r.ChancePct = .ResistChanceHemoPct
                r.DanoMinFlat = .ResistDanoHemoFlat
                r.DanoMaxFlat = .ResistDanoHemoFlat
                r.DanoMinPct = .ResistDanoHemoPct
                r.DanoMaxPct = .ResistDanoHemoPct
                If .InmunidadHemo <> 0 Then r.Inmune = 1
            Case 3
                r.ChancePct = .ResistChanceNeuroPct
                r.DanoMinFlat = .ResistDanoNeuroFlat
                r.DanoMaxFlat = .ResistDanoNeuroFlat
                r.DanoMinPct = .ResistDanoNeuroPct
                r.DanoMaxPct = .ResistDanoNeuroPct
                If .InmunidadNeuro <> 0 Then r.Inmune = 1
        End Select
        ' 2) Resistencia generica (afecta cualquier familia, se SUMA a la especifica)
        '    - Inmunidad: gana si la especifica o la generica esta activa.
        '    - Chance: se suma (capeo natural via chFin <= 0 en el caller).
        '    - Dano flat/pct: se suman a los rangos especificos (ApplyDamageResist los aplica en cascada).
        If .InmunidadVenenoGenerica <> 0 Then r.Inmune = 1
        r.ChancePct = r.ChancePct + .ResistChanceVenenoGenericoPct
        r.DanoMinFlat = r.DanoMinFlat + .ResistDanoVenenoGenericoFlat
        r.DanoMaxFlat = r.DanoMaxFlat + .ResistDanoVenenoGenericoFlat
        r.DanoMinPct = r.DanoMinPct + .ResistDanoVenenoGenericoPct
        r.DanoMaxPct = r.DanoMaxPct + .ResistDanoVenenoGenericoPct
    End With
    GetNpcPoisonResist = r
End Function
