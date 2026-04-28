Attribute VB_Name = "modProfesiones"
' Argentum 20 Game Server
'
'    Copyright (C) 2026 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
Option Explicit

' Sistema de Profesiones Aprendibles.
' Gateado por el feature toggle "professions_learnable".
' Ver: ia/plans/2026-04-22-sistema-profesiones-aprendibles.md

Public Const PROF_MAX_SLOTS As Byte = 2
Public Const PROF_MIN_ID    As Byte = 17
Public Const PROF_MAX_ID    As Byte = 23

Private Type t_ProfesionConfig
    SkillId            As Integer
    NpcMaestroId       As Integer
    ItemManualId       As Integer
    ItemPocionOlvidoId As Integer
    SkillParaMax       As Integer
    CantidadMin        As Integer
    CantidadMax        As Integer
    ProbBase           As Single
    ProbPorSkill       As Single
    CantidadFija       As Integer
    EsExtraccion       As Boolean
    Cargada            As Boolean
End Type

Private ProfesionConfig(PROF_MIN_ID To PROF_MAX_ID) As t_ProfesionConfig
Private MaxOlvidosPorPersonaje    As Byte
Private MaxProfesionesAprendidas  As Byte

Public Sub LoadProfesionesConfig()
    On Error GoTo LoadProfesionesConfig_Err
    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    Call Lector.Initialize(App.Path & "\profesiones.ini")
    MaxOlvidosPorPersonaje = CByte(val(Lector.GetValue("General", "MaxOlvidosPorPersonaje")))
    MaxProfesionesAprendidas = CByte(val(Lector.GetValue("General", "MaxProfesionesAprendidas")))
    If MaxOlvidosPorPersonaje = 0 Then MaxOlvidosPorPersonaje = 2
    If MaxProfesionesAprendidas = 0 Then MaxProfesionesAprendidas = 2
    Call LoadSeccionProfesion(Lector, "Carpinteria", e_Skill.Carpinteria, False)
    Call LoadSeccionProfesion(Lector, "Herreria", e_Skill.Herreria, False)
    Call LoadSeccionProfesion(Lector, "Sastreria", e_Skill.Sastreria, False)
    Call LoadSeccionProfesion(Lector, "Alquimia", e_Skill.Alquimia, False)
    Call LoadSeccionProfesion(Lector, "Pesca", e_Skill.Pescar, True)
    Call LoadSeccionProfesion(Lector, "Tala", e_Skill.Talar, True)
    Call LoadSeccionProfesion(Lector, "Mineria", e_Skill.Mineria, True)
    Set Lector = Nothing
    Exit Sub
LoadProfesionesConfig_Err:
    Set Lector = Nothing
    Call TraceError(Err.Number, Err.Description, "modProfesiones.LoadProfesionesConfig", Erl)
End Sub

Private Sub LoadSeccionProfesion(ByRef Lector As clsIniManager, ByVal seccion As String, ByVal skillId As Integer, ByVal esExtraccion As Boolean)
    On Error GoTo LoadSeccionProfesion_Err
    With ProfesionConfig(skillId)
        .SkillId = skillId
        .NpcMaestroId = CInt(val(Lector.GetValue(seccion, "NpcMaestroId")))
        .ItemManualId = CInt(val(Lector.GetValue(seccion, "ItemManualId")))
        .ItemPocionOlvidoId = CInt(val(Lector.GetValue(seccion, "ItemPocionOlvidoId")))
        .EsExtraccion = esExtraccion
        If esExtraccion Then
            .SkillParaMax = CInt(val(Lector.GetValue(seccion, "SkillParaMax")))
            .CantidadMin = CInt(val(Lector.GetValue(seccion, "CantidadMin")))
            .CantidadMax = CInt(val(Lector.GetValue(seccion, "CantidadMax")))
            .ProbBase = CSng(val(Lector.GetValue(seccion, "ProbBase")))
            .ProbPorSkill = CSng(val(Lector.GetValue(seccion, "ProbPorSkill")))
            If .SkillParaMax <= 0 Then .SkillParaMax = 100
        Else
            .CantidadFija = CInt(val(Lector.GetValue(seccion, "CantidadFija")))
            If .CantidadFija = 0 Then .CantidadFija = 1
        End If
        .Cargada = True
    End With
    Exit Sub
LoadSeccionProfesion_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.LoadSeccionProfesion", Erl)
End Sub

Public Function NombreProfesion(ByVal ProfesionId As Integer) As String
    If ProfesionId < LBound(SkillsNames) Or ProfesionId > UBound(SkillsNames) Then
        NombreProfesion = "?"
    Else
        NombreProfesion = SkillsNames(ProfesionId)
    End If
End Function

Public Function TieneProfesionAprendida(ByVal UserIndex As Integer, ByVal ProfesionId As Integer) As Boolean
    On Error GoTo TieneProfesionAprendida_Err
    TieneProfesionAprendida = False
    If Not IsFeatureEnabled("professions_learnable") Then
        TieneProfesionAprendida = True
        Exit Function
    End If
    Dim i As Byte
    For i = 1 To PROF_MAX_SLOTS
        If UserList(UserIndex).Professions(i) = ProfesionId Then
            TieneProfesionAprendida = True
            Exit Function
        End If
    Next i
    Exit Function
TieneProfesionAprendida_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.TieneProfesionAprendida", Erl)
End Function

Public Function SlotsOcupados(ByVal UserIndex As Integer) As Byte
    Dim i     As Byte
    Dim count As Byte
    count = 0
    For i = 1 To PROF_MAX_SLOTS
        If UserList(UserIndex).Professions(i) > 0 Then count = count + 1
    Next i
    SlotsOcupados = count
End Function

Public Function PuedeAprenderProfesion(ByVal UserIndex As Integer, ByVal ProfesionId As Integer) As Boolean
    On Error GoTo PuedeAprenderProfesion_Err
    PuedeAprenderProfesion = False
    If Not IsFeatureEnabled("professions_learnable") Then Exit Function
    If ProfesionId < PROF_MIN_ID Or ProfesionId > PROF_MAX_ID Then Exit Function
    If TieneProfesionAprendida(UserIndex, ProfesionId) Then
        Call WriteLocaleMsg(UserIndex, MSG_PROF_YA_APRENDIDA, e_FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(UserIndex).ProfessionForgotCount >= MaxOlvidosPorPersonaje And SlotsOcupados(UserIndex) >= MaxProfesionesAprendidas Then
        Call WriteLocaleMsg(UserIndex, MSG_PROF_SLOT_BLOQUEADO_POR_OLVIDOS, e_FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If SlotsOcupados(UserIndex) >= MaxProfesionesAprendidas Then
        Call WriteLocaleMsg(UserIndex, MSG_PROF_YA_TIENE_2, e_FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    PuedeAprenderProfesion = True
    Exit Function
PuedeAprenderProfesion_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.PuedeAprenderProfesion", Erl)
End Function

Public Function PuedeOlvidarProfesion(ByVal UserIndex As Integer, ByVal ProfesionId As Integer) As Boolean
    On Error GoTo PuedeOlvidarProfesion_Err
    PuedeOlvidarProfesion = False
    If Not IsFeatureEnabled("professions_learnable") Then Exit Function
    If ProfesionId < PROF_MIN_ID Or ProfesionId > PROF_MAX_ID Then Exit Function
    If Not TieneProfesionAprendida(UserIndex, ProfesionId) Then
        Call WriteLocaleMsg(UserIndex, MSG_PROF_NO_APRENDIDA, e_FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UserList(UserIndex).ProfessionForgotCount >= MaxOlvidosPorPersonaje Then
        Call WriteLocaleMsg(UserIndex, MSG_PROF_MAX_OLVIDOS_PERSONAJE, e_FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    PuedeOlvidarProfesion = True
    Exit Function
PuedeOlvidarProfesion_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.PuedeOlvidarProfesion", Erl)
End Function

Public Sub AprenderProfesion(ByVal UserIndex As Integer, ByVal ProfesionId As Integer)
    On Error GoTo AprenderProfesion_Err
    If Not IsFeatureEnabled("professions_learnable") Then Exit Sub
    If Not PuedeAprenderProfesion(UserIndex, ProfesionId) Then Exit Sub
    Dim i As Byte
    For i = 1 To PROF_MAX_SLOTS
        If UserList(UserIndex).Professions(i) = 0 Then
            UserList(UserIndex).Professions(i) = ProfesionId
            Exit For
        End If
    Next i
    Call Execute("INSERT OR REPLACE INTO user_professions (user_id, profession_id, learned_at) VALUES (?, ?, ?);", UserList(UserIndex).Id, CInt(ProfesionId), CLng(GetTickCountRaw() \ 1000))
    Call RefrescarHerramientasInventario(UserIndex, ProfesionId)
    Call WriteLocaleMsg(UserIndex, MSG_PROF_APRENDIDA_OK, e_FontTypeNames.FONTTYPE_INFO, NombreProfesion(ProfesionId))
    Exit Sub
AprenderProfesion_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.AprenderProfesion", Erl)
End Sub

Public Sub OlvidarProfesion(ByVal UserIndex As Integer, ByVal ProfesionId As Integer)
    On Error GoTo OlvidarProfesion_Err
    If Not IsFeatureEnabled("professions_learnable") Then Exit Sub
    If Not PuedeOlvidarProfesion(UserIndex, ProfesionId) Then Exit Sub
    Dim i            As Byte
    Dim puntosDevol  As Byte
    For i = 1 To PROF_MAX_SLOTS
        If UserList(UserIndex).Professions(i) = ProfesionId Then
            UserList(UserIndex).Professions(i) = 0
            Exit For
        End If
    Next i
    puntosDevol = UserList(UserIndex).Stats.UserSkills(ProfesionId)
    UserList(UserIndex).Stats.UserSkills(ProfesionId) = 0
    If puntosDevol > 0 Then
        UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts + puntosDevol
    End If
    UserList(UserIndex).ProfessionForgotCount = UserList(UserIndex).ProfessionForgotCount + 1
    Call Execute("DELETE FROM user_professions WHERE user_id = ? AND profession_id = ?;", UserList(UserIndex).Id, CInt(ProfesionId))
    Call Execute("UPDATE user SET profession_forgot_count = ? WHERE id = ?;", CLng(UserList(UserIndex).ProfessionForgotCount), UserList(UserIndex).Id)
    Call RefrescarHerramientasInventario(UserIndex, ProfesionId)
    Call WriteLocaleMsg(UserIndex, MSG_PROF_OLVIDADA_OK, e_FontTypeNames.FONTTYPE_INFO, NombreProfesion(ProfesionId))
    Exit Sub
OlvidarProfesion_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.OlvidarProfesion", Erl)
End Sub

Public Function CalcularExitoExtraccion(ByVal UserIndex As Integer, ByVal ProfesionId As Integer) As Boolean
    On Error GoTo CalcularExitoExtraccion_Err
    CalcularExitoExtraccion = True
    If ProfesionId < PROF_MIN_ID Or ProfesionId > PROF_MAX_ID Then Exit Function
    If Not ProfesionConfig(ProfesionId).EsExtraccion Then Exit Function
    Dim skillActual As Integer
    Dim prob        As Single
    skillActual = UserList(UserIndex).Stats.UserSkills(ProfesionId)
    With ProfesionConfig(ProfesionId)
        prob = .ProbBase + (.ProbPorSkill * CSng(skillActual))
    End With
    If prob > 100 Then prob = 100
    If prob < 0 Then prob = 0
    CalcularExitoExtraccion = (RandomNumber(1, 100) <= CInt(prob))
    Exit Function
CalcularExitoExtraccion_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.CalcularExitoExtraccion", Erl)
    CalcularExitoExtraccion = True
End Function

Public Function CalcularCantidadExtraccion(ByVal UserIndex As Integer, ByVal ProfesionId As Integer) As Integer
    On Error GoTo CalcularCantidadExtraccion_Err
    CalcularCantidadExtraccion = 1
    If ProfesionId < PROF_MIN_ID Or ProfesionId > PROF_MAX_ID Then Exit Function
    If Not ProfesionConfig(ProfesionId).EsExtraccion Then Exit Function
    Dim skillActual As Integer
    Dim rango       As Integer
    Dim resultado   As Single
    skillActual = UserList(UserIndex).Stats.UserSkills(ProfesionId)
    With ProfesionConfig(ProfesionId)
        If skillActual > .SkillParaMax Then skillActual = .SkillParaMax
        rango = .CantidadMax - .CantidadMin
        If rango < 0 Then rango = 0
        resultado = CSng(.CantidadMin) + (CSng(rango) * CSng(skillActual) / CSng(.SkillParaMax))
    End With
    If resultado < 1 Then resultado = 1
    CalcularCantidadExtraccion = CInt(resultado)
    Exit Function
CalcularCantidadExtraccion_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.CalcularCantidadExtraccion", Erl)
End Function

Public Function CantidadFijaArtesania(ByVal ProfesionId As Integer) As Integer
    If ProfesionId < PROF_MIN_ID Or ProfesionId > PROF_MAX_ID Then
        CantidadFijaArtesania = 1
    Else
        CantidadFijaArtesania = ProfesionConfig(ProfesionId).CantidadFija
        If CantidadFijaArtesania < 1 Then CantidadFijaArtesania = 1
    End If
End Function

Public Function ProfesionDelItem(ByVal ObjIndex As Integer) As Integer
    On Error GoTo ProfesionDelItem_Err
    ProfesionDelItem = 0
    If ObjIndex <= 0 Then Exit Function
    If ObjData(ObjIndex).OBJType <> e_OBJType.otWorkingTools Then Exit Function
    Select Case ObjData(ObjIndex).Subtipo
        Case e_WorkingToolSubType.FishingRod, e_WorkingToolSubType.FishingNet, e_WorkingToolSubType.FishingLine
            ProfesionDelItem = e_Skill.Pescar
        Case e_WorkingToolSubType.AlchemyScissors, e_WorkingToolSubType.AlchemyCauldron
            ProfesionDelItem = e_Skill.Alquimia
        Case e_WorkingToolSubType.CarpentryHacksaw
            ProfesionDelItem = e_Skill.Carpinteria
        Case e_WorkingToolSubType.FellingAxe
            ProfesionDelItem = e_Skill.Talar
        Case e_WorkingToolSubType.SmithHammer
            ProfesionDelItem = e_Skill.Herreria
        Case e_WorkingToolSubType.MinerPickaxe
            ProfesionDelItem = e_Skill.Mineria
        Case e_WorkingToolSubType.TailorSewingbox
            ProfesionDelItem = e_Skill.Sastreria
    End Select
    Exit Function
ProfesionDelItem_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.ProfesionDelItem", Erl)
End Function

Public Sub UsarManualProfesion(ByVal UserIndex As Integer, ByVal Slot As Byte)
    On Error GoTo UsarManualProfesion_Err
    If Not IsFeatureEnabled("professions_learnable") Then
        Call WriteLocaleMsg(UserIndex, MSG_PROF_ITEM_NO_USABLE, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    Dim ObjIndex    As Integer
    Dim ProfesionId As Integer
    ObjIndex = UserList(UserIndex).invent.Object(Slot).ObjIndex
    If ObjIndex <= 0 Then Exit Sub
    ProfesionId = ObjData(ObjIndex).ProfesionId
    If ProfesionId < PROF_MIN_ID Or ProfesionId > PROF_MAX_ID Then Exit Sub
    If Not PuedeAprenderProfesion(UserIndex, ProfesionId) Then Exit Sub
    Call AprenderProfesion(UserIndex, ProfesionId)
    Call QuitarUserInvItem(UserIndex, Slot, 1)
    Call UpdateUserInv(False, UserIndex, Slot)
    Exit Sub
UsarManualProfesion_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.UsarManualProfesion", Erl)
End Sub

Public Sub UsarPocionOlvidoProfesion(ByVal UserIndex As Integer, ByVal Slot As Byte)
    On Error GoTo UsarPocionOlvidoProfesion_Err
    If Not IsFeatureEnabled("professions_learnable") Then
        Call WriteLocaleMsg(UserIndex, MSG_PROF_ITEM_NO_USABLE, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    Dim ObjIndex    As Integer
    Dim ProfesionId As Integer
    ObjIndex = UserList(UserIndex).invent.Object(Slot).ObjIndex
    If ObjIndex <= 0 Then Exit Sub
    ProfesionId = ObjData(ObjIndex).ProfesionId
    If ProfesionId < PROF_MIN_ID Or ProfesionId > PROF_MAX_ID Then Exit Sub
    If Not PuedeOlvidarProfesion(UserIndex, ProfesionId) Then Exit Sub
    Call OlvidarProfesion(UserIndex, ProfesionId)
    Call QuitarUserInvItem(UserIndex, Slot, 1)
    Call UpdateUserInv(False, UserIndex, Slot)
    Exit Sub
UsarPocionOlvidoProfesion_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.UsarPocionOlvidoProfesion", Erl)
End Sub

Private Sub RefrescarHerramientasInventario(ByVal UserIndex As Integer, ByVal ProfesionId As Integer)
    On Error GoTo RefrescarHerramientasInventario_Err
    Dim slot   As Byte
    Dim ObjIdx As Integer
    For slot = 1 To MAX_INVENTORY_SLOTS
        ObjIdx = UserList(UserIndex).invent.Object(slot).ObjIndex
        If ObjIdx > 0 Then
            If ProfesionDelItem(ObjIdx) = ProfesionId Then
                Call WriteChangeInventorySlot(UserIndex, slot)
            End If
        End If
    Next slot
    Exit Sub
RefrescarHerramientasInventario_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.RefrescarHerramientasInventario", Erl)
End Sub

Public Function HandleProfesionChatCommand(ByVal UserIndex As Integer, ByVal chat As String) As Boolean
    On Error GoTo HandleProfesionChatCommand_Err
    HandleProfesionChatCommand = False
    Dim body  As String
    Dim parts() As String
    Dim cmd   As String
    Dim targetName As String
    Dim profArg    As String
    Dim profId     As Integer
    Dim tUserRef   As t_UserReference
    body = mid$(chat, 2)
    body = Trim$(body)
    If LenB(body) = 0 Then Exit Function
    parts = Split(body, " ")
    If UBound(parts) < 2 Then Exit Function
    cmd = UCase$(parts(0))
    If cmd <> "APRENDE" And cmd <> "OLVIDA" Then Exit Function
    HandleProfesionChatCommand = True
    targetName = Replace(parts(1), "+", " ")
    profArg = parts(2)
    profId = ResolverProfesionId(profArg)
    If profId < PROF_MIN_ID Or profId > PROF_MAX_ID Then
        Call WriteConsoleMsg(UserIndex, "Profesion invalida: " & profArg & ". Usa nombre (Carpinteria, Tala, ...) o id (17-23).", e_FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If UCase$(targetName) = "YO" Then
        Call SetUserRef(tUserRef, UserIndex)
    Else
        tUserRef = NameIndex(targetName)
    End If
    If Not IsValidUserRef(tUserRef) Then
        Call WriteConsoleMsg(UserIndex, "Usuario offline: " & targetName, e_FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If cmd = "APRENDE" Then
        Call AprenderProfesion(tUserRef.ArrayIndex, profId)
        Call WriteConsoleMsg(UserIndex, "Aprendio " & NombreProfesion(profId) & " para " & targetName, e_FontTypeNames.FONTTYPE_INFO)
    Else
        Call OlvidarProfesion(tUserRef.ArrayIndex, profId)
        Call WriteConsoleMsg(UserIndex, "Olvido " & NombreProfesion(profId) & " para " & targetName, e_FontTypeNames.FONTTYPE_INFO)
    End If
    Call LogGM(GetUserRealName(UserIndex), "/" & cmd & " " & targetName & " " & NombreProfesion(profId))
    Exit Function
HandleProfesionChatCommand_Err:
    Call TraceError(Err.Number, Err.Description, "modProfesiones.HandleProfesionChatCommand", Erl)
End Function

Private Function ResolverProfesionId(ByVal arg As String) As Integer
    Dim n As Integer
    n = CInt(val(arg))
    If n >= PROF_MIN_ID And n <= PROF_MAX_ID Then
        ResolverProfesionId = n
        Exit Function
    End If
    Select Case UCase$(arg)
        Case "TALA", "TALAR"
            ResolverProfesionId = e_Skill.Talar
        Case "PESCA", "PESCAR"
            ResolverProfesionId = e_Skill.Pescar
        Case "MINERIA"
            ResolverProfesionId = e_Skill.Mineria
        Case "HERRERIA"
            ResolverProfesionId = e_Skill.Herreria
        Case "CARPINTERIA"
            ResolverProfesionId = e_Skill.Carpinteria
        Case "ALQUIMIA"
            ResolverProfesionId = e_Skill.Alquimia
        Case "SASTRERIA"
            ResolverProfesionId = e_Skill.Sastreria
        Case Else
            ResolverProfesionId = 0
    End Select
End Function
