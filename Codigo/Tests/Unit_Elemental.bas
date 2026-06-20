Attribute VB_Name = "Unit_Elemental"
Option Explicit
#If UNIT_TEST = 1 Then

' ==========================================================================
' Elemental Damage Test Suite (plan 20.002, TOGGLE32 elemental_system)
' Tests the deterministic core of modElementalCombat: resistance cascade
' (immune / chance / flat / pct + cap) and component rolling bounds.
' Chance-based paths use 0 or 100 to stay deterministic.
' ==========================================================================
Public Function test_suite_elemental() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start

    Call InitDamageTypeRegistryDefaults

    Call UnitTesting.RunTest("test_resist_immune", test_resist_immune())
    Call UnitTesting.RunTest("test_resist_flat", test_resist_flat())
    Call UnitTesting.RunTest("test_resist_flat_overkill", test_resist_flat_overkill())
    Call UnitTesting.RunTest("test_resist_pct_half", test_resist_pct_half())
    Call UnitTesting.RunTest("test_resist_pct_capped", test_resist_pct_capped())
    Call UnitTesting.RunTest("test_resist_chance_full", test_resist_chance_full())
    Call UnitTesting.RunTest("test_resist_chance_zero", test_resist_chance_zero())
    Call UnitTesting.RunTest("test_resist_zero_raw", test_resist_zero_raw())
    Call UnitTesting.RunTest("test_roll_fixed", test_roll_fixed())
    Call UnitTesting.RunTest("test_roll_swapped", test_roll_swapped())
    Call UnitTesting.RunTest("test_roll_negative", test_roll_negative())
    Call UnitTesting.RunTest("test_damage_type_name", test_damage_type_name())

    Debug.Print "Elemental suite took " & sw.ElapsedMilliseconds & " ms"
    test_suite_elemental = True
End Function

' Immunity nullifies the whole component and flags outNullified.
Private Function test_resist_immune() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean, d As Long
    r.Immune = 1
    d = ApplyElementalResist(100, r, eDmgFire, nul)
    test_resist_immune = (d = 0 And nul = True)
    Exit Function
e:
    test_resist_immune = False
End Function

' Flat reduction subtracts before pct: 10 - 3 = 7.
Private Function test_resist_flat() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean
    r.ReduceFlat = 3
    test_resist_flat = (ApplyElementalResist(10, r, eDmgFire, nul) = 7)
    Exit Function
e:
    test_resist_flat = False
End Function

' Flat greater than raw clamps to 0.
Private Function test_resist_flat_overkill() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean
    r.ReduceFlat = 10
    test_resist_flat_overkill = (ApplyElementalResist(5, r, eDmgFire, nul) = 0)
    Exit Function
e:
    test_resist_flat_overkill = False
End Function

' 50% pct reduction: 100 -> 50.
Private Function test_resist_pct_half() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean
    r.ReducePct = 0.5
    test_resist_pct_half = (ApplyElementalResist(100, r, eDmgFire, nul) = 50)
    Exit Function
e:
    test_resist_pct_half = False
End Function

' Pct above the default cap (0.75) is clamped: 0.9 -> 0.75, so 100 -> 25.
Private Function test_resist_pct_capped() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean
    r.ReducePct = 0.9
    test_resist_pct_capped = (ApplyElementalResist(100, r, eDmgFire, nul) = 25)
    Exit Function
e:
    test_resist_pct_capped = False
End Function

' ReduceChancePct=100 always nullifies.
Private Function test_resist_chance_full() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean, d As Long
    r.ReduceChancePct = 100
    d = ApplyElementalResist(100, r, eDmgFire, nul)
    test_resist_chance_full = (d = 0 And nul = True)
    Exit Function
e:
    test_resist_chance_full = False
End Function

' ReduceChancePct=0 never nullifies; full damage passes.
Private Function test_resist_chance_zero() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean
    test_resist_chance_zero = (ApplyElementalResist(10, r, eDmgFire, nul) = 10 And nul = False)
    Exit Function
e:
    test_resist_chance_zero = False
End Function

' Zero raw damage stays zero.
Private Function test_resist_zero_raw() As Boolean
    On Error GoTo e
    Dim r As t_ElementalResist, nul As Boolean
    r.ReducePct = 0.5
    test_resist_zero_raw = (ApplyElementalResist(0, r, eDmgFire, nul) = 0)
    Exit Function
e:
    test_resist_zero_raw = False
End Function

' min==max rolls the fixed value.
Private Function test_roll_fixed() As Boolean
    On Error GoTo e
    Dim c As t_DamageComponent
    c.DamageType = eDmgFire: c.MinDamage = 7: c.MaxDamage = 7
    test_roll_fixed = (RollDamageComponent(c) = 7)
    Exit Function
e:
    test_roll_fixed = False
End Function

' min>max is sanitized (hi=lo), returns lo.
Private Function test_roll_swapped() As Boolean
    On Error GoTo e
    Dim c As t_DamageComponent
    c.DamageType = eDmgFire: c.MinDamage = 9: c.MaxDamage = 3
    test_roll_swapped = (RollDamageComponent(c) = 9)
    Exit Function
e:
    test_roll_swapped = False
End Function

' negative bounds clamp to 0.
Private Function test_roll_negative() As Boolean
    On Error GoTo e
    Dim c As t_DamageComponent
    c.DamageType = eDmgFire: c.MinDamage = -5: c.MaxDamage = -1
    test_roll_negative = (RollDamageComponent(c) = 0)
    Exit Function
e:
    test_roll_negative = False
End Function

' Registry default name for fire.
Private Function test_damage_type_name() As Boolean
    On Error GoTo e
    test_damage_type_name = (DamageTypeName(eDmgFire) = "Fuego")
    Exit Function
e:
    test_damage_type_name = False
End Function

#End If
