Attribute VB_Name = "modPoisonResist"
' Argentum 20 Game Server
'
' modPoisonResist
' Sistema de venenos (TOGGLE26 new_poison_system).
' ADAPTER (plan 20.002 Step 4): la resist al EFECTO del veneno (Inmune + ChancePct)
' se lee del motor elemental (tipo Veneno, eDmgPoison). Las 3 familias colapsaron a
' ese unico tipo. La resist al DANO del tick ya la aplica el motor en PerformTick.
' GetUser/NpcPoisonResist traducen esa resist para los call-sites que solo usan
' .Inmune y .ChancePct.

Option Explicit

Public Type t_PoisonResist
    ChancePct As Long
    DanoMinFlat As Long
    DanoMaxFlat As Long
    DanoMinPct As Long
    DanoMaxPct As Long
    Inmune As Byte
End Type

' Devuelve la resistencia total de un usuario para la familia dada.
' Suma aditiva de los 7 slots equipables relevantes.
Public Function GetUserPoisonResist(ByVal UserIndex As Integer, ByVal familia As Byte) As t_PoisonResist
    ' Adapter (Step 4 dimension EFECTO): la resist al efecto del veneno sale del motor elemental (tipo
    ' Veneno). Los call-sites solo usan Inmune + ChancePct (la resist al DAÑO/tick ya migro). familia se
    ' ignora: las 3 familias comparten el tipo Veneno. Ver Engram resist-model.
    Dim r As t_PoisonResist
    If UserIndex <= 0 Then
        GetUserPoisonResist = r
        Exit Function
    End If
    Dim rE As t_ElementalResist
    rE = modElementalCombat.GetUserElementalResist(UserIndex, e_ElementalDamageType.eDmgPoison)
    If rE.Immune <> 0 Or rE.ImmuneEffect <> 0 Then r.Inmune = 1
    r.ChancePct = rE.ReduceEffectChancePct
    GetUserPoisonResist = r
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
    ' Adapter (Step 4 dimension EFECTO): resist al efecto del veneno desde el motor elemental (tipo Veneno).
    Dim r As t_PoisonResist
    If NpcIndex <= 0 Then
        GetNpcPoisonResist = r
        Exit Function
    End If
    Dim rE As t_ElementalResist
    rE = modElementalCombat.GetNpcElementalResist(NpcIndex, e_ElementalDamageType.eDmgPoison)
    If rE.Immune <> 0 Or rE.ImmuneEffect <> 0 Then r.Inmune = 1
    r.ChancePct = rE.ReduceEffectChancePct
    GetNpcPoisonResist = r
End Function
