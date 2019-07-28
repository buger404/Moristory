VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FinalPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Private Type Gun
        X As Single
        y As Single
        Class As Integer
        Tag As Variant
        Tag2 As Variant
    End Type
    Private Type MagicItem
        Name As String
        BasicATK As Long
        AP As Long
        RoundCount As Long
        NCount As Long
        Copyed As Boolean
    End Type
    Private Type RoleItem
        DrawX As Long
        DrawY As Long
        TargetX As Long
        TargetY As Long
        Time As Long
        Name As String
        HP As Long
        MaxHP As Long
        MP As Long
        MaxMP As Long
        LastHP As Long
        Magic() As MagicItem
    End Type
    Private Type NoteItem
        Direction As Integer
        Time As Single
        Hand As Long
    End Type
    Private Type NoteList
        Note() As NoteItem
    End Type
    Dim Guns() As Gun
    Dim NL As NoteList, Role() As RoleItem
    Dim HurtTime As Long
    Dim Round As Long, RoundTime As Long, RoundCount As Long
    Dim msg() As String, MsgTime As Long, ATKTotal As Long
    Dim THead As Integer
    Dim ATKBot As Long, NowNote As Long, ATKCount As Long, NowATKCount As Long
    Dim FireTick As Double
    Dim PX As Single, PY As Single
    Dim Crash As New GCrashBox
    Dim FireTime As Single, FireMissed As Long
    Dim SuperMode As Boolean
'======================================
Public Sub SetNote()
    For i = 1 To UBound(NL.Note)
        If NL.Note(i).Time > BGM.position Then NowNote = i: Exit For
    Next
End Sub
Public Sub NewMsg(ByVal Text As String)
    ReDim Preserve msg(UBound(msg) + 1)
    msg(UBound(msg)) = Text
    MsgTime = GetTickCount
End Sub
Public Sub StartAttack()
    Dim LHP As Long, LHP2 As Long
    LHP = Role(3).HP
    
    Dim Skip As Boolean
    Skip = True
    
    Randomize
    Dim Skilled As Boolean
    Dim Start As Long, Ends As Long
    
    For i = 1 To 2
        Skilled = False
        If i = 2 And (Role(0).HP < Role(0).MaxHP * 0.4 And Role(1).HP < Role(1).MaxHP * 0.4 And Role(2).HP < Role(2).MaxHP * 0.4) Then
            If Role(i).Magic(2).NCount = 0 And Role(i).MP >= Role(i).Magic(2).AP Then
                Role(i).Magic(2).NCount = Role(i).Magic(2).NCount + 1
                Role(i).MP = Role(i).MP - Role(i).Magic(2).AP
                
                If Role(i).Magic(2).RoundCount > 3 Then NewMsg Role(i).Name & "开始吟唱<" & Role(i).Magic(2).Name & ">。"
                Skilled = True
            End If
        End If
        If Role(i).MP >= Role(i).MaxMP * 0.3 Then
            Start = 1: Ends = 0
        Else
            Start = 0: Ends = 1
        End If
        
        For S = Start To Ends Step IIf(Start > Ends, -1, 1)
            If Skilled Then Exit For
            If Role(i).Magic(S).NCount = 0 And Role(i).MP >= Role(i).Magic(S).AP Then
                Role(i).Magic(S).NCount = Role(i).Magic(S).NCount + 1
                Role(i).MP = Role(i).MP - Role(i).Magic(S).AP
                
                If Role(i).Magic(S).RoundCount > 3 Then NewMsg Role(i).Name & "开始吟唱<" & Role(i).Magic(S).Name & ">。"
                Skilled = True
            End If
        Next
    Next
    
    For i = 0 To 2
        For S = 0 To UBound(Role(i).Magic)
            If Role(i).Magic(S).NCount > 0 Then Role(i).Magic(S).NCount = Role(i).Magic(S).NCount + 1
            If Role(i).Magic(S).NCount >= Role(i).Magic(S).RoundCount Then
                Role(i).Magic(S).NCount = 0
                
                Skip = False
                
                Page.SwitchChannel Role(i).Magic(S).Name, "Default"
                
                Randomize
                LHP2 = Role(3).HP
                If Role(i).Magic(S).BasicATK > 0 Then
                    Role(3).HP = Role(3).HP - Role(i).Magic(S).BasicATK * (1 + (Rnd * 0.3)) * IIf(SuperMode, 10, 1)
                    NewMsg Role(i).Name & IIf(Role(i).Magic(S).RoundCount = 1, "发动", "吟唱") & "了<" & Role(i).Magic(S).Name & ">，" & Role(3).Name & "受到了" & (LHP2 - Role(3).HP) & "点伤害！"
                ElseIf Role(i).Magic(S).BasicATK = 0 Then
                    For p = 0 To 2
                        Role(p).HP = Role(p).LastHP
                    Next
                    NewMsg Role(i).Name & "使用了<" & Role(i).Magic(S).Name & ">，生命值回到了上一回合的数值！"
                ElseIf Role(i).Magic(S).BasicATK < 0 Then
                    If Role(i).Magic(S).Name = "灵力祭坛" Then
                        For p = 0 To 2
                            Role(p).MP = Role(p).MP - Role(i).Magic(S).BasicATK * (1 + (Rnd * 0.3))
                        Next
                        NewMsg Role(i).Name & "吟唱了<" & Role(i).Magic(S).Name & ">，灵力值回复了！"
                    Else
                        For p = 0 To 2
                            Role(p).HP = Role(p).HP - Role(i).Magic(S).BasicATK * (1 + (Rnd * 0.3))
                        Next
                        NewMsg Role(i).Name & "吟唱了<" & Role(i).Magic(S).Name & ">，生命值回复了！"
                    End If
                End If
            End If
        Next
    Next
    
    ATKTotal = Role(3).HP - LHP
    If ATKTotal <> 0 Then ECore.NewTransform 14, 1000
    
    Round = 3: RoundTime = GetTickCount
    
    If Skip Then RoundTime = GetTickCount - 10000
End Sub
Public Sub Damage(ATK As Long)
    Randomize
    Role(THead).HP = Role(THead).HP - ATK * (1 + Rnd * 0.5)
End Sub
Public Sub UpdateUATK()
    Dim i As Integer, DX As Long, DY As Long

    If ATKBot = 3 Or ATKBot = 4 Or ATKBot = 1 Or ATKBot = 2 Then
        Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(120, 255, 255, 255)
    ElseIf ATKBot <> 10 Then
        Page.DrawImage "DarkSpace.png", 0, 0
        Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(Sin(GetTickCount / 200) * 20 + 126, 0, 0, 0)
    End If
    If ATKBot = 6 Then Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(120, 255, 255, 255)
    If ATKBot = 8 Or ATKBot = 9 Then Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(150, 255, 255, 255)
    If ATKBot = 10 Then Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(150, 255, 255, 255)
    
    Dim RemoveMark As Boolean, CanHurt As Boolean, CrashR As Boolean
    CanHurt = (GetTickCount - HurtTime >= 3000)
    
    Page.DrawImage "head" & THead & ".png", VBA.Round(PX), VBA.Round(PY), alpha:=IIf((CanHurt = False And (GetTickCount Mod 300 < 150)) Or CanHurt, 1, 0), Pos:=posOnCenter
    Crash.Reset
    
    If ATKBot = 8 Or ATKBot = 9 Or ATKBot = 10 Then
        Page.DrawImage "gun04.png", GW / 2, GH / 2, alpha:=1, Pos:=posOnCenter
        If Crash.IsCrash Then Damage 100000
        
        If GW / 2 > PX Then PX = PX + IIf(ATKBot >= 9, 2, 1)
        If GH / 2 > PY Then PY = PY + IIf(ATKBot >= 9, 2, 1)
        If GW / 2 < PX Then PX = PX - IIf(ATKBot >= 9, 2, 1)
        If GH / 2 < PY Then PY = PY - IIf(ATKBot >= 9, 2, 1)
        
        For i = 1 To UBound(Guns)
            With Guns(i)
                .X = .X + Int(GW / 2 - .X) / 30
                .y = .y + Int(GH / 2 - .y) / 30
            End With
        Next
    End If
    
    For i = 1 To UBound(Guns)
        If i > UBound(Guns) Then Exit For
        
        RemoveMark = False
        With Guns(i)
            Select Case Guns(i).Class
                Case 7
                    Page.DrawImage "gun01.png", Int(.X), Int(.y), alpha:=1, Pos:=posOnCenter
                    If CanHurt Then CrashR = Crash.IsCrash
                    RemoveMark = RemoveMark Or (CrashR) And CanHurt
                    If CrashR Then Damage 100
                    If .Tag2 = 3 Or .Tag2 = 4 Or .Tag2 = 10 Then .X = .X - 3
                    If .Tag2 = 4 Or .Tag2 = 5 Or .Tag2 = 6 Then .X = .X + 3
                    If .Tag2 = 6 Or .Tag2 = 7 Or .Tag2 = 8 Then .y = .y + 3
                    If .Tag2 = 8 Or .Tag2 = 9 Or .Tag2 = 10 Then .y = .y - 3
                Case 8
                    Page.DrawImage "gun05.png", Int(.X), Int(.y), alpha:=1, Pos:=posOnCenter
                    If CanHurt Then CrashR = Crash.IsCrash
                    RemoveMark = RemoveMark Or (CrashR) And CanHurt
                    If CrashR Then Damage 100
                    
                    If Abs(GW / 2 - .X) <= 20 And Abs(GH / 2 - .y) <= 20 Then RemoveMark = True
                Case 6
                    Page.DrawImage "gun02.png", Int(.X), Int(.y), alpha:=.Tag2 / 60, Pos:=posOnCenter
                    If .Tag2 < 60 Then .Tag2 = .Tag2 + 1
                    If .Tag2 = 60 Then
                        If CanHurt Then CrashR = Crash.IsCrash
                        RemoveMark = RemoveMark Or (CrashR) And CanHurt
                        If CrashR Then Damage 200
                        If Abs(.X - PX) <= 200 And Abs(.y - PY) <= 200 Then
                            If .X > PX Then PX = PX + 0.5
                            If .y > PY Then PY = PY + 0.5
                            If .X < PX Then PX = PX - 0.5
                            If .y < PY Then PY = PY - 0.5
                        End If
                    End If
                    If ATKBot = 9 Or ATKBot = 10 Then
                        If Abs(GW / 2 - .X) <= 20 And Abs(GH / 2 - .y) <= 20 Then RemoveMark = True
                    End If
                Case 5
                    Page.DrawImage "gun01.png", Int(.X), Int(.y), alpha:=1, Pos:=posOnCenter
                    If CanHurt Then CrashR = Crash.IsCrash
                    RemoveMark = RemoveMark Or (CrashR) And CanHurt
                    If CrashR Then Damage 100
                    
                    .X = .X + Cos(3 / 2 * 3.14 - .Tag2) * 10
                    .y = .y + Sin(3 / 2 * 3.14 - .Tag2) * 10
                    .Tag2 = .Tag2 - 0.06
                    If .Tag2 < -2 Then .Tag2 = 2
                    
                Case 0
                    Page.DrawImage "gun00.png", Int(.X), Int(.y), alpha:=1, Pos:=posOnCenter
                    If CanHurt Then CrashR = Crash.IsCrash
                    RemoveMark = RemoveMark Or (CrashR) And CanHurt
                    If CrashR Then Damage 100
                    
                    .X = .X + Cos(3 / 2 * 3.14 - .Tag2) * 13
                    .y = .y + Sin(3 / 2 * 3.14 - .Tag2) * 10
                    .Tag2 = .Tag2 - 0.03
                    If .Tag2 < -2 Then .Tag2 = 2
                    
                Case 3
                    Page.DrawImage "gun03.png", Int(.X), Int(.y), alpha:=1, Pos:=posOnCenter
                    If CanHurt Then CrashR = Crash.IsCrash
                    RemoveMark = RemoveMark Or (CrashR) And CanHurt
                    If CrashR Then Damage 140
                    
                    .X = .X + Cos(3 / 2 * 3.14 - .Tag2) * 3
                    .y = .y - Sin(3 / 2 * 3.14 - .Tag2) * 10
                    '.Tag2 = .Tag2 - 0.03
                    If .Tag2 < -2 Then .Tag2 = 2
            End Select
        End With
        
        If Guns(i).X < -50 Or Guns(i).y < -50 Or Guns(i).X > GW + 50 Or Guns(i).y > GH + 50 Then RemoveMark = True
        
        If RemoveMark Then
            Guns(i) = Guns(UBound(Guns))
            ReDim Preserve Guns(UBound(Guns) - 1)
            i = i - 1
        End If
    Next
    
    If NowNote <= UBound(NL.Note) Then
        If BGM.position > NL.Note(NowNote).Time Then
            NowNote = NowNote + 1
            Select Case ATKBot
                Case 7
                    FireTick = FireTick + 1
                    If FireTick Mod 5 = 0 Then
                        SE.Play "Sword" & Int(Rnd * 4 + 1) & ".mp3"
                        DX = Int(Rnd * GW)
                        DY = Int(Rnd * GH)
                        For i = 0 To 7
                            ReDim Preserve Guns(UBound(Guns) + 1)
                            With Guns(UBound(Guns))
                                .X = DX
                                .y = DY
                                .Class = 7
                                .Tag2 = i + 3
                            End With
                        Next
                    End If
                Case 0, 1
                    SE.Play "Fire" & Int(Rnd * 2 + 1) & ".mp3"
                    For i = 0 To 6
                        FireTick = FireTick + 1
                        ReDim Preserve Guns(UBound(Guns) + 1)
                        With Guns(UBound(Guns))
                            .X = GW / 2
                            .y = GH
                            .Class = 0
                            .Tag2 = (FireTick Mod 8) / 2 - 2 + (FireTick Mod 128) / 128 * 0.5
                        End With
                    Next
                Case 3
                    SE.Play "Darkness" & Int(Rnd * 4 + 1) & ".mp3"
                    For i = 0 To 6
                        FireTick = FireTick + 1
                        ReDim Preserve Guns(UBound(Guns) + 1)
                        With Guns(UBound(Guns))
                            .Class = 3
                            .Tag = Int(Rnd * 4)
                            .X = IIf(.Tag Mod 2 = 0, 0, GW)
                            .y = IIf((.Tag - 2) = 0, 0, GH)
                            .Tag2 = Abs((FireTick Mod 8) / 2 - 2 + (FireTick Mod 128) / 128 * 0.5) * IIf(.Tag Mod 2 = 0, -1, 1)
                        End With
                    Next
                Case 4, 2
                    SE.Play "Darkness" & Int(Rnd * 4 + 1) & ".mp3"
                    For i = 0 To 3
                        FireTick = FireTick + 1
                        ReDim Preserve Guns(UBound(Guns) + 1)
                        With Guns(UBound(Guns))
                            .Class = 3
                            .Tag = Int(Rnd * 4)
                            .X = IIf(.Tag Mod 2 = 0, 0, GW)
                            .y = IIf((.Tag - 2) = 0, 0, GH)
                            .Tag2 = Abs((FireTick Mod 8) / 2 - 2 + (FireTick Mod 128) / 128 * 0.5) * IIf(.Tag Mod 2 = 0, -1, 1)
                        End With
                    Next
                    SE.Play "Fire" & Int(Rnd * 2 + 1) & ".mp3"
                    For i = 0 To 3
                        FireTick = FireTick + 1
                        ReDim Preserve Guns(UBound(Guns) + 1)
                        With Guns(UBound(Guns))
                            .X = GW / 2
                            .y = GH
                            .Class = 0
                            .Tag2 = (FireTick Mod 8) / 2 - 2 + (FireTick Mod 128) / 128 * 0.5
                        End With
                    Next
                Case 5
                    SE.Play "Fire" & Int(Rnd * 2 + 1) & ".mp3"
                    FireTick = FireTick + 1
                    ReDim Preserve Guns(UBound(Guns) + 1)
                    With Guns(UBound(Guns))
                        .X = (FireTick Mod 20) / 20 * (GW) - 50 + (FireTick Mod 100) / 20 * 10
                        .y = GH - 50 + (FireTick Mod 20) / 20 * 100
                        .Class = 5
                        .Tag2 = 0
                    End With
                Case 6
                    SE.Play "Splash.mp3"
                    FireTick = FireTick + 1
                    ReDim Preserve Guns(UBound(Guns) + 1)
                    With Guns(UBound(Guns))
                        .X = Int(Rnd * GW)
                        .y = Int(Rnd * GH)
                        .Class = 6
                        .Tag2 = 0
                    End With
                Case 9
                    SE.Play "Splash.mp3"
                    FireTick = FireTick + 1
                    ReDim Preserve Guns(UBound(Guns) + 1)
                    With Guns(UBound(Guns))
                        .X = Int(Rnd * GW)
                        .y = Int(Rnd * GH)
                        .Class = 6
                        .Tag2 = 0
                    End With
                Case 8
                    SE.Play "Sword" & Int(Rnd * 4 + 1) & ".mp3"
                    For i = 0 To 1
                        FireTick = FireTick + 1
                        ReDim Preserve Guns(UBound(Guns) + 1)
                        With Guns(UBound(Guns))
                            .Tag = Int(Rnd * 4)
                            .Class = 8
                            Select Case .Tag
                                Case 0: .X = -Int(Rnd * 50): .y = Int(Rnd * GH)
                                Case 1: .X = GW + Int(Rnd * 50): .y = Int(Rnd * GH)
                                Case 2: .X = Int(Rnd * GW): .y = -Int(Rnd * 50)
                                Case Else: .X = Int(Rnd * GW): .y = GH + Int(Rnd * 50)
                            End Select
                        End With
                    Next
                Case 10
                    SE.Play "Sword" & Int(Rnd * 4 + 1) & ".mp3"
                    For i = 0 To 1
                        FireTick = FireTick + 1
                        ReDim Preserve Guns(UBound(Guns) + 1)
                        With Guns(UBound(Guns))
                            .Tag = Int(Rnd * 4)
                            .Class = 8
                            Select Case .Tag
                                Case 0: .X = -Int(Rnd * 50): .y = Int(Rnd * GH)
                                Case 1: .X = GW + Int(Rnd * 50): .y = Int(Rnd * GH)
                                Case 2: .X = Int(Rnd * GW): .y = -Int(Rnd * 50)
                                Case Else: .X = Int(Rnd * GW): .y = GH + Int(Rnd * 50)
                            End Select
                        End With
                    Next
                    SE.Play "Splash.mp3"
                    FireTick = FireTick + 1
                    ReDim Preserve Guns(UBound(Guns) + 1)
                    With Guns(UBound(Guns))
                        .X = Int(Rnd * GW)
                        .y = Int(Rnd * GH)
                        .Class = 6
                        .Tag2 = 0
                    End With
            End Select
        End If
    End If
    
    'OUCH！！！！
    If Crash.CrashResult Then
        HurtTime = GetTickCount
        Randomize
        SE.Play "Damage" & Int(Rnd * 4 + 1) & ".mp3"
    End If
    
    DX = GW / 2: DY = 100
    i = THead
    Page.DrawImage "bar" & i & ".png", DX, DY, alpha:=1, Pos:=posOnCenter
    
    Page.Writes Role(i).Name, DX - GW / 2 + 290, DY - 17, 20, argb(255, 0, 0, 0), align:=StringAlignmentCenter
    
    Page.Writes Role(i).HP & "/" & Role(i).MaxHP, DX - GW / 2 + 400, DY - 15, 18, argb(180, 100, 0, 0)
    DW = EF.GetWidth(Page.GG, Role(i).HP & "/" & Role(i).MaxHP, 18, StringAlignmentNear, FontStyleRegular)
    Page.Writes Role(i).HP & "/" & Role(i).MaxHP, DX - GW / 2 + 400, DY - 15, 18, argb(255, 255, 255, 255), W:=Role(i).HP / Role(i).MaxHP * DW + 1, h:=22
    
    Page.Writes Role(i).MP & "/" & Role(i).MaxMP, DX - GW / 2 + 585, DY - 15, 18, argb(180, 100, 0, 0)
    DW = EF.GetWidth(Page.GG, Role(i).MP & "/" & Role(i).MaxMP, 18, StringAlignmentNear, FontStyleRegular)
    Page.Writes Role(i).MP & "/" & Role(i).MaxMP, DX - GW / 2 + 585, DY - 15, 18, argb(255, 255, 255, 255), W:=Role(i).MP / Role(i).MaxMP * DW + 1, h:=22
    
    If IsKeyPress(VK_LEFT) Then PX = PX - 5
    If IsKeyPress(VK_RIGHT) Then PX = PX + 5
    If IsKeyPress(VK_UP) Then PY = PY - 5
    If IsKeyPress(VK_DOWN) Then PY = PY + 5
    
    If PX <= 16 Then PX = 16
    If PY <= 16 Then PY = 16
    If PX >= GW - 32 Then PX = GW - 32
    If PY >= GH - 32 Then PY = GH - 32
    
    Dim FireOK As Boolean
    Select Case ATKBot
        Case 0: If GetTickCount - RoundTime > 10000 / ATKCount Then FireOK = True
        Case 1: If GetTickCount - RoundTime > 12000 / ATKCount Then FireOK = True
        Case 2: If GetTickCount - RoundTime > 15000 / ATKCount Then FireOK = True
        
        Case 3: If GetTickCount - RoundTime > 12000 / ATKCount Then FireOK = True
        Case 4: If GetTickCount - RoundTime > 15000 / ATKCount Then FireOK = True
        Case 5: If GetTickCount - RoundTime > 15000 / ATKCount Then FireOK = True
        Case 6: If GetTickCount - RoundTime > 15000 / ATKCount Then FireOK = True
        Case 7: If GetTickCount - RoundTime > 15000 / ATKCount Then FireOK = True
        Case 8: If GetTickCount - RoundTime > 15000 / ATKCount Then FireOK = True
        Case 9: If GetTickCount - RoundTime > 15000 / ATKCount Then FireOK = True
        Case 10: If GetTickCount - RoundTime > 15000 / ATKCount Then FireOK = True
    End Select


    If FireOK Then
        
        If FireTick > 400 / ATKCount Then FireTime = FireTime + 10
        
        NowATKCount = NowATKCount + 1
        If NowATKCount < ATKCount Then
            Call UATK
            SE.Play "Decision1.mp3"
            Exit Sub
        End If
        
        FireMissed = FireMissed + FireTick
        
        ReDim Guns(0)
        FireTick = 0
        NowATKCount = 0
        For i = 0 To 2
            Role(i).TargetX = GW / 2
        Next
        Round = 0
        ECore.NewTransform transHighLight, 100
    End If
    
    For i = 0 To 2
        If Role(i).HP < 0 Then
            Call InitRoles
            Round = 9: RoundTime = GetTickCount
            SE.Play "Collapse4.mp3"
            ESave.PutData "FinalBOSSDIE", Val(ESave.GetData("FinalBOSSDIE")) + 1
            BGM.Stops
            Exit Sub
        End If
    Next
    
End Sub
Public Sub UATK()
    Randomize
    ATKBot = Int(Rnd * 11)
    If ATKBot > 10 Then ATKBot = 10
    If NowATKCount = 0 Then
        Select Case ATKBot
            Case 0: PX = GW / 2: PY = GH - 250
            Case 8: PX = GW / 2: PY = GH - 50: SE.Play "Darkness1.mp3"
            Case 9: PX = GW / 2: PY = GH - 50: SE.Play "Darkness1.mp3"
            Case 10: PX = GW / 2: PY = GH - 50: SE.Play "Darkness1.mp3"
            Case Else: PX = GW / 2: PY = GH - 250
        End Select
    End If
    RoundTime = GetTickCount
    ECore.NewTransform transHighLight, 100
End Sub
Public Sub Update()
    If Err.number <> 0 Then ErrCrash Err.number, Err.Description
    On Error Resume Next
    
    '游戏每一帧的过程（绘制过程请放在此处）
    
    If BGM.PlayState = musStopped And BGM.length <> 0 And Round < 9 Then NowNote = 1: BGM.Play
    If BGS.PlayState = musStopped And BGS.length <> 0 Then BGS.Play
    
    If Round = 4 Then Call UpdateUATK: Exit Sub
    
    Page.Clear argb(255, 0, 0, 0) '清空画布
    
    'ATKBot = 10
    
    Dim RPro As Single
    If Round = 9 Then
        RPro = (GetTickCount - RoundTime) / 2000
        If RPro > 1 Then RPro = 1
        RPro = Cubic(RPro, 0, 1, 1, 1)
        Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(255 * RPro, 0, 0, 0)
        If GetTickCount Mod 300 < 150 Then Page.Writes "我们失败了", GW / 2, GH / 2, 36, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
        If GetTickCount - RoundTime >= 3000 Then Round = 10: RoundTime = GetTickCount: ECore.NewTransform transFadeIn, 1000
        Exit Sub
    End If
    If Round = 10 Then
        Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(255, 0, 0, 0)
        Dim DText As String
        Select Case Val(ESave.GetData("FinalBOSSDIE"))
            Case 1: DText = "但是我们的希望仍然紧握在手中"
            Case 2: DText = "但是我们的机会是无限的"
            Case 3: DText = "但是我们一定不能失去信息"
            Case 4: DText = "但是我们不能放弃"
            Case 5: DText = "但是我们必须坚持下去"
            Case Else: DText = "但是希望仍在"
        End Select
        Page.Writes DText, GW / 2, GH / 2, 36, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
        If GetTickCount - RoundTime >= 3000 Then Round = 11: RoundTime = GetTickCount: ECore.NewTransform transHighLight, 100: SE.Play "Load.mp3"
        Exit Sub
    End If
    If Round = 11 Then
        RPro = (GetTickCount - RoundTime) / 3000
        If RPro > 1 Then RPro = 1
        
        Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(255, 255, 255, 255)
        Page.Writes "<时间倒流>", GW / 2, GH / 2, 36, argb(255, 0, 0, 0), align:=StringAlignmentCenter, style:=FontStyleBold
        
        Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(255 * RPro, 255, 255, 255)
        If GetTickCount - RoundTime >= 3000 Then Round = 0: RoundTime = GetTickCount: ECore.NewTransform 14, 3000
        Exit Sub
    End If
    
    If Round <> 2 Then
        Page.DrawImage "DarkSpace.png", 0, 0
        Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(Sin(GetTickCount / 200) * 20 + 126, 0, 0, 0)
        
        Page.DrawImage "km2-normal.png", GW / 2, GH / 2 + Cos((GetTickCount / 250) + i / 5) * 50, alpha:=Cos(GetTickCount / 500) * 0.2 + 0.8, Pos:=posOnCenter
    End If
    
    Dim DW As Long, DX As Long, DY As Long, Pro As Single, LHP As Long
    Dim CanUse As Boolean
    
    If Round = 2 Then
        If (Not SuperMode) And FireMissed >= 3500 And Role(3).HP <= 980000 Then
            DX = GW / 2: DY = 150
            
            Page.Paint 2, GW / 2, DY - 3, 529, 35, argb(120, 50, 202, 185), 90, Pos:=posOnCenter
            Page.DrawImage "magic.png", DX, DY, alpha:=1, Pos:=posOnCenter
            If CheckMouse2 = mMouseUp Then
                ECore.NewTransform 14, 3000
                SE.Play "Thunder9.mp3"
                SuperMode = True
                BGM.rate = 2
                NewMsg Role(0).Name & "发动了<天使之佑>，所有成员的魔法伤害提高10倍！"
            End If
            
            Page.Writes "天使之佑", DX - GW / 2 + 290, DY - 17, 20, IIf(CanUse, argb(255, 0, 0, 0), argb(255, 255, 0, 0)), align:=StringAlignmentCenter
            
            Page.Writes "放大攻击力10倍", DX - GW / 2 + 410, DY - 15, 18, argb(255, 255, 255, 255)
            Page.Writes "0", DX - GW / 2 + 585, DY - 15, 18, argb(255, 255, 255, 255)

            GoTo SkipMagic
        End If
        
        For i = 0 To UBound(Role(0).Magic)
            DX = GW / 2: DY = 150 + 50 * i
            
            Page.Paint 2, GW / 2, DY - 3, 529, 35, argb(120, 255, 255, 255), 90, Pos:=posOnCenter
            
            Page.DrawImage "magic.png", DX, DY, alpha:=1, Pos:=posOnCenter
            
            CanUse = (Role(0).Magic(i).NCount = 0 And Role(0).MP >= Role(0).Magic(i).AP * IIf(Role(0).Magic(i).Copyed, 1.2, 1))
            
            If CheckMouse2 = mMouseUp Then
                If CanUse Then
                    SE.Play "Cursor.mp3"
                    Role(0).TargetY = GH - 150
                    Role(0).TargetX = -GW
                    Role(0).MP = Role(0).MP - Role(0).Magic(i).AP * IIf(Role(0).Magic(i).Copyed, 1.2, 1)
                    Role(0).Magic(i).NCount = Role(0).Magic(i).NCount + 1
                    Role(3).TargetY = 100
                    If Role(0).Magic(i).RoundCount > 3 Then NewMsg Role(0).Name & "开始吟唱<" & Role(0).Magic(i).Name & ">。"
                    
                    Call StartAttack
                    
                Else
                    SE.Play "Buzzer1.mp3"
                End If
            End If
            
            If Role(0).Magic(i).NCount > 0 Then
                Page.Paint 2, GW / 2 - 529 / 2, DY - 3 - 35 / 2, (Role(0).Magic(i).NCount - 1) / (Role(0).Magic(i).RoundCount - 1) * 529, 35, argb(100, 0, 176, 240), 90
            End If
            
            Page.Writes Role(0).Magic(i).Name, DX - GW / 2 + 290, DY - 17, 20, IIf(CanUse, argb(255, 0, 0, 0), argb(255, 255, 0, 0)), align:=StringAlignmentCenter
            
            Page.Writes IIf(Role(0).Magic(i).BasicATK < 0, "0", Role(0).Magic(i).BasicATK), DX - GW / 2 + 410, DY - 15, 18, argb(255, 255, 255, 255)
            Page.Writes Role(0).Magic(i).AP * IIf(Role(0).Magic(i).Copyed, 1.2, 1), DX - GW / 2 + 585, DY - 15, 18, argb(255, 255, 255, 255)
            
            If Not CanUse Then Page.Paint 2, GW / 2, DY - 3, 529, 36, argb(120, 0, 0, 0), 90, Pos:=posOnCenter
            
            If Role(0).Magic(i).NCount > 0 Then
                Page.Writes "[正在吟唱]", DX - GW / 2 + 120, DY - 15, 18, argb(255, 0, 176, 240)
            ElseIf Role(0).Magic(i).Copyed Then
                Page.Writes "[完全分析]", DX - GW / 2 + 120, DY - 15, 18, argb(255, 254, 84, 57)
            End If
        Next
SkipMagic:
    End If

    'Page.DrawImage "focus.png", GW / 2, GH - 150 - 2, alpha:=1, Pos:=posOnCenter
    For i = 0 To 3
        If Role(i).DrawX <> Role(i).TargetX Or Role(i).DrawY <> Role(i).TargetY Then
            If Role(i).Time = 0 Then Role(i).Time = GetTickCount
        End If
        
        If Role(i).Time <> 0 Then
            If GetTickCount - Role(i).Time >= 600 Then
                Role(i).DrawX = Role(i).TargetX: Role(i).DrawY = Role(i).TargetY: Role(i).Time = 0
                DX = Role(i).DrawX: DY = Role(i).DrawY
                If Round = 1 Then Round = 2: ECore.NewTransform transFadeIn, 1000
            Else
                Pro = Cubic((GetTickCount - Role(i).Time) / 600, 0, 1, 1, 1)
                DX = Role(i).DrawX + (Role(i).TargetX - Role(i).DrawX) * Pro: DY = Role(i).DrawY + (Role(i).TargetY - Role(i).DrawY) * Pro
            End If
        Else
            DX = Role(i).DrawX: DY = Role(i).DrawY
        End If
        Page.DrawImage "bar" & i & ".png", DX, DY, alpha:=1, Pos:=posOnCenter
        
        If i = 0 And Round = 0 Then
            If CheckMouse2 = mMouseUp Then
                SE.Play "Cursor.mp3"
                Role(0).TargetY = 100
                Role(1).TargetX = -GW
                Role(2).TargetX = -GW
                Role(3).TargetY = -200
                Round = 1
            End If
        End If
        
        Page.Writes Role(i).Name, DX - GW / 2 + 290, DY - 17, 20, argb(255, 0, 0, 0), align:=StringAlignmentCenter
        
        If i = 3 Then
            Page.Writes Role(i).HP, DX - GW / 2 + 400, DY - 15, 18, argb(255, 255, 255, 255)
            Page.Writes "无法分析", DX - GW / 2 + 585, DY - 15, 18, argb(255, 255, 255, 255)
        Else
            Page.Writes Role(i).HP & "/" & Role(i).MaxHP, DX - GW / 2 + 400, DY - 15, 18, argb(180, 100, 0, 0)
            DW = EF.GetWidth(Page.GG, Role(i).HP & "/" & Role(i).MaxHP, 18, StringAlignmentNear, FontStyleRegular)
            Page.Writes Role(i).HP & "/" & Role(i).MaxHP, DX - GW / 2 + 400, DY - 15, 18, argb(255, 255, 255, 255), W:=Role(i).HP / Role(i).MaxHP * DW + 1, h:=22
            
            Page.Writes Role(i).MP & "/" & Role(i).MaxMP, DX - GW / 2 + 585, DY - 15, 18, argb(180, 100, 0, 0)
            DW = EF.GetWidth(Page.GG, Role(i).MP & "/" & Role(i).MaxMP, 18, StringAlignmentNear, FontStyleRegular)
            Page.Writes Role(i).MP & "/" & Role(i).MaxMP, DX - GW / 2 + 585, DY - 15, 18, argb(255, 255, 255, 255), W:=Role(i).MP / Role(i).MaxMP * DW + 1, h:=22
        End If
    Next
    
    If Round = 3 Then
        Dim TPro As Single
        TPro = (GetTickCount - RoundTime) / 3000
        If TPro > 1 Then TPro = 1
        TPro = Cubic(TPro, 0, 1, 1, 1)
        Page.DrawAnimation "桃花斩", GW / 2, GH / 2 + 100
        Page.DrawAnimation "千叶斩", GW / 2, GH / 2 + 100
        Page.DrawAnimation "圣光矢", GW / 2, GH / 2 + 100
        Page.DrawAnimation "圣羽", GW / 2, GH / 2 + 100
        
        Page.DrawAnimation "青铜炮", GW / 2, GH / 2 + 100
        Page.DrawAnimation "翡翠炮", GW / 2, GH / 2 + 100
        Page.DrawAnimation "青光圣堂", GW / 2, GH / 2 + 100
        Page.DrawAnimation "翡翠导弹", GW / 2, GH / 2 + 100
        Page.DrawAnimation "能力回流", GW / 2, GH / 2 + 320
        Page.DrawAnimation "灵力祭坛", GW / 2, GH / 2 + 320
        Page.DrawAnimation "花之精华", GW / 2, GH / 2 + 320
        
        Page.Writes ATKTotal, GW / 2, GH / 2 + 280 - TPro * 280, 64, argb(255 * TPro, 255, 255, 255), style:=FontStyleBold, align:=StringAlignmentCenter
        If GetTickCount - RoundTime >= 4000 Then
            For i = 0 To 2
                Role(i).MP = Role(i).MP + 100
                If Role(i).HP > Role(i).MaxHP Then Role(i).HP = Role(i).MaxHP
                If Role(i).MP > Role(i).MaxMP Then Role(i).MP = Role(i).MaxMP
                Role(i).LastHP = Role(i).HP
            Next
            THead = Int(Rnd * 3)
            If THead > 2 Then THead = 2
            Round = 4
            
            NowATKCount = 0
            ATKCount = 1
            If SuperMode Then ATKCount = 6 + Int(Rnd * 6)
            Call UATK
            Call SetNote
        End If
    End If

    If UBound(msg) >= 1 Then
        Dim MsgPro As Single
        MsgPro = (GetTickCount - MsgTime) / 1000
        If MsgPro >= 1.5 Then MsgPro = (2 - MsgPro) / 0.5
        If MsgPro <= 1.5 Then MsgPro = Cubic(MsgPro / 1.5, 0, 1, 1, 1)
        If MsgPro < 0 Then MsgPro = 0
        If MsgPro > 1 Then MsgPro = 0
        
        Page.Writes msg(1), GW / 2, 240 - 60 * MsgPro, 18, argb(255 * MsgPro, 255, 255, 255), align:=StringAlignmentCenter
        If MsgPro = 0 Then
            For i = 0 To UBound(msg) - 1
                msg(i) = msg(i + 1)
            Next
            ReDim Preserve msg(UBound(msg) - 1)
            MsgTime = GetTickCount
        End If
    End If

    If Round = 0 Then
        DX = 40
        For i = 0 To 2
            For S = 0 To UBound(Role(i).Magic)
                If Role(i).Magic(S).NCount > 0 Then
                    Page.DrawImage "head" & i & ".png", DX, GH - 230, alpha:=1, Pos:=posOnBottom
                    Page.Paint 1, DX + 14, GH - 230 - 14, 28, 28, argb(255, 255, 255, 255)
                    Page.Writes Role(i).Magic(S).RoundCount - Role(i).Magic(S).NCount, DX + 14 + 14, GH - 230 - 14 + 2, 18, argb(255, 0, 0, 0), align:=StringAlignmentCenter
                    Page.Writes Role(i).Magic(S).Name, DX + 16 + 2, GH - 230 + 20, 16, argb(200, 255, 255, 255), align:=StringAlignmentCenter
                    DX = DX + 60
                End If
            Next
        Next
        If Role(3).HP <= 900000 Then
            NovelPage.RunNovel
            ECore.NewTransform transFallDark, 3000, "NovelPage"
        End If
    End If
    
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub
Public Sub InitRoles()
    ReDim Role(3)
    
    BGM.rate = 1
    'SuperMode = False
    FireTime = 0: FireMissed = 0
    NowATKCount = 0
    If SuperMode Then BGM.rate = 2
    
    Role(0).Name = "青"
    Role(0).MaxHP = 6000
    Role(0).MaxMP = 5000
    
    Role(1).Name = "黑嘴"
    Role(1).MaxHP = 8000
    Role(1).MaxMP = 2000
    
    Role(2).Name = "兮兮"
    Role(2).MaxHP = 3000
    Role(2).MaxMP = 7000
    
    Role(3).Name = "枯梦"
    Role(3).MaxHP = 1000000
    Role(3).MaxMP = 1000000000
    
    RoundCount = 1
    
    For i = 0 To 3
        Role(i).MP = Role(i).MaxMP
        Role(i).HP = Role(i).MaxHP
        Role(i).TargetX = GW / 2
        Role(i).TargetY = GH - 150 + 50 * i
        Role(i).Time = GetTickCount
        Role(i).DrawX = -GW
        Role(i).DrawY = -GH
        Role(i).LastHP = Role(i).HP
    Next
    Role(3).TargetY = 100
    
    i = 0
    ReDim Role(i).Magic(5)
    With Role(i).Magic(0)
        .Name = "青铜炮"
        .BasicATK = 30
        .AP = 30
        .RoundCount = 2
    End With
    With Role(i).Magic(1)
        .Name = "翡翠炮"
        .BasicATK = 60
        .AP = 50
        .RoundCount = 3
    End With
    With Role(i).Magic(2)
        .Name = "青光圣堂"
        .BasicATK = 300
        .AP = 400
        .RoundCount = 5
    End With
    With Role(i).Magic(3)
        .Name = "翡翠导弹"
        .BasicATK = 170
        .AP = 100
        .RoundCount = 3
    End With
    With Role(i).Magic(4)
        .Name = "能力回流"
        .BasicATK = 0
        .AP = 15
        .RoundCount = 2
    End With
    With Role(i).Magic(5)
        .Name = "灵力祭坛"
        .BasicATK = -500
        .AP = 0
        .RoundCount = 4
    End With
    
    i = 1
    ReDim Role(i).Magic(1)
    With Role(i).Magic(0)
        .Name = "圣光矢"
        .BasicATK = 100
        .AP = 90
        .RoundCount = 3
    End With
    With Role(i).Magic(1)
        .Name = "圣羽"
        .BasicATK = 600
        .AP = 400
        .RoundCount = 2
    End With

    i = 2
    ReDim Role(i).Magic(2)
    With Role(i).Magic(0)
        .Name = "千叶斩"
        .BasicATK = 100
        .AP = 90
        .RoundCount = 3
    End With
    With Role(i).Magic(1)
        .Name = "桃花斩"
        .BasicATK = 200
        .AP = 150
        .RoundCount = 2
    End With
    With Role(i).Magic(2)
        .Name = "花之精华"
        .BasicATK = -500
        .AP = 200
        .RoundCount = 2
    End With
    
    For i = 1 To 2
        For S = 0 To UBound(Role(i).Magic)
            ReDim Preserve Role(0).Magic(UBound(Role(0).Magic) + 1)
            Role(0).Magic(UBound(Role(0).Magic)) = Role(i).Magic(S)
            Role(0).Magic(UBound(Role(0).Magic)).Copyed = True
        Next
    Next
End Sub
Public Sub Enter()
    '页面进入事件
    BGM.Create App.Path & "\music\bgm\OST 27 - It's Our Fight.mp3"
    BGM.Play
    BGS.Create App.Path & "\music\bgs\Wind_1_3D_loop.mp3"
    BGS.Play
    Call InitRoles
End Sub
Public Sub Leave()
    '页面离开事件
    BGM.rate = 1
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(id As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.Res.NoLoadCrash = False
    Page.Res.NewImages App.Path & "\assets\finalui"
    Page.Res.NewImages App.Path & "\assets\finalui\ui", 0.5
    '创建页面
    ECore.Add Page, "FinalPage"
    Open App.Path & "\assets\km2finalboss.gun" For Binary As #1
    Get #1, , NL
    Close #1
    
    ReDim msg(0)
    ReDim Guns(0)
    
    Page.Res.HotLoad = True
    Page.Res.NoLoadCrash = True
    Page.Res.NewImages App.Path & "\assets\battle\images", 1.8
    Page.LoadAnimationsFromDir App.Path & "\battle_animation"
    
    Page.CreatePlayAnimation "rpgmv_Wind2", "千叶斩", "Reserve"
    Page.CreatePlayAnimation "rpgmv_SlashSpecial3", "桃花斩", "Reserve"
    Page.CreatePlayAnimation "rpgmv_Recovery4+Recovery5", "花之精华", "Reserve"
    
    Page.CreatePlayAnimation "rpgmv_Holy4", "圣光矢", "Reserve"
    Page.CreatePlayAnimation "rpgmv_Holy2", "圣羽", "Reserve"
    
    Page.CreatePlayAnimation "rpgmv_Laser1", "青铜炮", "Reserve"
    Page.CreatePlayAnimation "rpgmv_Laser2", "翡翠炮", "Reserve"
    Page.CreatePlayAnimation "rpgmv_Meteor+Gun2", "青光圣堂", "Reserve"
    Page.CreatePlayAnimation "rpgmv_Magic2", "翡翠导弹", "Reserve"
    Page.CreatePlayAnimation "rpgmv_Recovery1", "能力回流", "Reserve"
    
    Page.CreatePlayAnimation "rpgmv_Recovery2", "灵力祭坛", "Reserve"

End Sub
