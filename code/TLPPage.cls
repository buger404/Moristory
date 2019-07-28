VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TLPPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Private Type MagicItem
        Name As String
        BasicATK As Long
        AP As Long
        Selected As Boolean
        GetTime As Long
    End Type
    Dim ICard() As MagicItem, UCard() As MagicItem
    Dim MyCard() As MagicItem, YourCard() As MagicItem
    Dim msg() As String, MsgTime As Long, ATKTotal As Long
    Dim HP As Long, MaxHP As Long, MP As Long, MaxMP As Long
    Dim UHP As Long, UMaxHP As Long, UMP As Long, UMaxMP As Long
    Dim Round As Long, RoundTime As Long, LastHP As Long
    Dim SelectIndex As Long
    Dim CanMove As Boolean
'======================================
Public Sub NewMsg(ByVal Text As String)
    ReDim Preserve msg(UBound(msg) + 1)
    msg(UBound(msg)) = Text
    MsgTime = GetTickCount
End Sub
Public Sub Update()
    If Err.number <> 0 Then ErrCrash Err.number, Err.Description
    On Error Resume Next

    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear '清空画布

    Page.DrawImage "GrassMaze.png", 0, 0, alpha:=1
    Page.DrawImage "DemonicWorld.png", 0, 0, alpha:=1
    
    For i = 0 To 10
        Page.DrawImage "tk-normal.png", GW / 2 + Cos((GetTickCount / 250) + i / 5) * 50, GH / 2 + Sin((GetTickCount / 200) + i / 10) * 50, alpha:=Cubic(i / 10, 0, 0, 0, 1), Pos:=posOnCenter
    Next
    
    Dim MPro As Single, DrawX As Long
    Dim pro2 As Single
    
    If Round = 0 And CanMove = False Then
        For i = 1 To UBound(MyCard)
            DrawX = GW / 2 - (UBound(MyCard) * 50 + 100) / 2 + i * 100 / 2 - 50
            MPro = i / UBound(MyCard) * 0.5 + 0.5
            pro2 = Cubic((GetTickCount - MyCard(i).GetTime) / 1000, 0, 1, 1, 1)
            If pro2 > 1 Then pro2 = 1
            If MPro < 0 Then MPro = 0
            If MPro > 1 Then MPro = 1
            If Not MyCard(i).Selected Then
                Page.Paint 2, DrawX, GH - 200 * pro2, 150, 200, argb(IIf(MyCard(i).AP <= MP And SelectIndex < 4, 255, 100), 242 * MPro, 242 * MPro, 242 * MPro), 16
                Page.Paint 0, DrawX, GH - 200 * pro2, 51, 202, argb(IIf(MyCard(i).AP <= MP And SelectIndex < 4, 255, 100), 180 * MPro, 180 * MPro, 180 * MPro)
                Page.Writes MyCard(i).Name, DrawX + 10, GH - 200 * pro2 + 10, 18, IIf(MyCard(i).AP <= MP And SelectIndex < 4, argb(255 * MPro, 32, 32, 32), argb(255 * MPro, 255, 0, 0)), 20, 200
                Page.Writes MyCard(i).AP & vbCrLf & "AP", DrawX + 10, GH - 200 * pro2 + 150, 16, argb(180 * MPro, 32, 32, 32), 100, 80
            Else
                Page.Paint 2, DrawX, GH - 250, 150, 200, argb(255, 50 * MPro, 202 * MPro, 185 * MPro), 16
                Page.Paint 0, DrawX, GH - 250, 51, 202, argb(255, 0, 160 * MPro, 139 * MPro)
                Page.Writes MyCard(i).Name, DrawX + 10, GH - 250 + 10, 18, argb(255 * MPro, 255, 255, 255), 20, 200
                Page.Writes MyCard(i).AP & vbCrLf & "AP", DrawX + 10, GH - 250 + 150, 16, argb(180 * MPro, 255, 255, 255), 100, 80
            End If
            If CheckMouse(DrawX, GH - 200 - IIf(MyCard(i).Selected, 50, 0), 50, 200) = 3 Then
                If MyCard(i).Selected = False Then
                    If MyCard(i).AP <= MP And SelectIndex < 4 Then
                        SelectIndex = SelectIndex + 1: MP = MP - MyCard(i).AP
                        MyCard(i).Selected = True: SE.Play "Cursor.mp3"
                    End If
                Else
                    SelectIndex = SelectIndex - 1: MP = MP + MyCard(i).AP
                    MyCard(i).Selected = False: SE.Play "Cancel.mp3"
                End If
            End If
        Next
    End If
    
    If MP > MaxMP Then MP = MaxMP
    If UMP > UMaxMP Then UMP = UMaxMP
    If HP < 0 Then HP = 0
    If UHP < 0 Then UHP = 0
    
    Dim LHP2 As Long
    
    If Round = 0 Then
        If SelectIndex > 0 Or (CanMove And GetTickCount - RoundTime >= 3000) Then
            If Page.ShowSimpleButton("btn.png", GW / 2 - 194 / 2, GH / 2 - 20, "出牌", argb(255, 255, 255, 255), 18) = 3 Or (CanMove And GetTickCount - RoundTime >= 3000) Then
                Dim LHP As Long
                LHP = UHP
                Round = 1
                For i = 1 To UBound(MyCard)
                    If i > UBound(MyCard) Then Exit For
                    If MyCard(i).Selected Then
                        Page.SwitchChannel MyCard(i).Name, "Default"
                        Randomize
                        LHP2 = UHP
                        If MyCard(i).BasicATK > 0 Then
                            UHP = UHP - MyCard(i).BasicATK * (1 + (Rnd * 0.3))
                            NewMsg "青吟唱了<" & MyCard(i).Name & ">，塔克受到了" & (LHP2 - UHP) & "点伤害！"
                        ElseIf MyCard(i).BasicATK = 0 Then
                            HP = LastHP
                            NewMsg "青使用了<" & MyCard(i).Name & ">，生命值回到了上一回合的数值！"
                        ElseIf MyCard(i).BasicATK < 0 Then
                            MP = MP - MyCard(i).BasicATK * (1 + (Rnd * 0.3))
                            NewMsg "青发动了<" & MyCard(i).Name & ">，灵力值回复了！"
                        End If
                        MyCard(i) = MyCard(UBound(MyCard))
                        ReDim Preserve MyCard(UBound(MyCard) - 1)
                        i = i - 1
                    End If
                Next
                ATKTotal = (UHP - LHP)
                RoundTime = GetTickCount
                If CanMove Then RoundTime = GetTickCount - 5000
                If ATKTotal <> 0 Then ECore.NewTransform 14, 1000
            End If
        End If
    End If
    
    If Round = 1 Then
        Dim TPro As Single
        TPro = (GetTickCount - RoundTime) / 3000
        If TPro > 1 Then TPro = 1
        TPro = Cubic(TPro, 0, 1, 1, 1)
        Page.DrawAnimation "暗黑导弹", GW / 2, GH / 2
        Page.DrawAnimation "黑幕", GW / 2, GH / 2
        Page.DrawAnimation "暗黑撕裂", GW / 2, GH / 2
        Page.DrawAnimation "死亡之钟", GW / 2, GH / 2
        Page.DrawAnimation "青铜炮", GW / 2, GH / 2
        Page.DrawAnimation "翡翠炮", GW / 2, GH / 2
        Page.DrawAnimation "青光圣堂", GW / 2, GH / 2
        Page.DrawAnimation "翡翠导弹", GW / 2, GH / 2
        Page.DrawAnimation "能力回流", GW / 2, GH / 2 + 220
        Page.DrawAnimation "灵力祭坛", GW / 2, GH / 2 + 220
        Page.Writes ATKTotal, GW / 2, GH / 2 + 280 - TPro * 280, 64, argb(255 * TPro, 255, 255, 255), style:=FontStyleBold, align:=StringAlignmentCenter
        If GetTickCount - RoundTime >= 4000 Then
            SelectIndex = 0
            LastHP = HP
            CanMove = False
            Dim Count As Long
            For i = 1 To UBound(YourCard)
                If i > UBound(YourCard) Then Exit For
                If Count = 4 Then Exit For
                If UMP >= YourCard(i).AP Then
                    Page.SwitchChannel YourCard(i).Name, "Default"
                    Randomize
                    LHP2 = HP
                    If YourCard(i).BasicATK > 0 Then
                        HP = HP - YourCard(i).BasicATK * (1 + (Rnd * 0.3))
                        NewMsg "塔克吟唱了<" & YourCard(i).Name & ">，青受到了" & (LHP2 - HP) & "点伤害！"
                    ElseIf YourCard(i).BasicATK = 0 Then
                        HP = HP - MaxHP / 10
                        NewMsg "塔克发动了<" & YourCard(i).Name & ">，青受到猛烈的攻击，造成" & (LHP2 - HP) & "点伤害！"
                    ElseIf YourCard(i).BasicATK < 0 Then
                        UMP = UMP - YourCard(i).BasicATK * (1 + (Rnd * 0.3))
                        NewMsg "塔克发动了<" & YourCard(i).Name & ">，灵力值回复了！"
                    End If
                    UMP = UMP - YourCard(i).AP
                    YourCard(i) = YourCard(UBound(YourCard))
                    ReDim Preserve YourCard(UBound(YourCard) - 1)
                    Count = Count + 1
                    i = i - 1
                End If
            Next
            If Int(Rnd * 5) = 2 Then CanMove = True
            ATKTotal = (HP - LastHP)
            RoundTime = GetTickCount: Round = 2
            If ATKTotal <> 0 Then ECore.NewTransform 14, 1000
        End If
    End If
    
    If Round = 2 Then
        Page.DrawAnimation "暗黑导弹", GW / 2, GH / 2 + 220
        Page.DrawAnimation "黑幕", GW / 2, GH / 2 + 220
        Page.DrawAnimation "暗黑撕裂", GW / 2, GH / 2 + 220
        Page.DrawAnimation "死亡之钟", GW / 2, GH / 2 + 220
        Page.DrawAnimation "青铜炮", GW / 2, GH / 2 + 220
        Page.DrawAnimation "翡翠炮", GW / 2, GH / 2 + 220
        Page.DrawAnimation "青光圣堂", GW / 2, GH / 2 + 220
        Page.DrawAnimation "翡翠导弹", GW / 2, GH / 2 + 220
        Page.DrawAnimation "能力回流", GW / 2, GH / 2
        Page.DrawAnimation "灵力祭坛", GW / 2, GH / 2
        If GetTickCount - RoundTime >= 2500 Then
            RoundTime = GetTickCount: Round = 3
        End If
    End If
    If Round = 3 Then
        If ATKTotal = 0 Then Call GiveCards: Round = 0
        Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(100, 255, 0, 0)
        If GetTickCount Mod 300 < 150 Then Page.Writes ATKTotal, GW / 2, GH / 2 + 200, 64, argb(255, 255, 255, 255), style:=FontStyleBold, align:=StringAlignmentCenter
        If GetTickCount - RoundTime >= 2000 Then
            Call GiveCards: Round = 0: RoundTime = GetTickCount
        End If
    End If
    
    'Page.Paint 0, 0, 0, GW + 1, 95, argb(255, 32, 32, 32)
    'Page.Paint 0, 0, GH - 100, GW + 1, 100 + 1, argb(255, 255, 255, 255)
    
    Page.DrawImage "youbar.png", 0, 0, alpha:=IIf(Round = 0, 0.3, 1)
    Page.DrawImage "mebar.png", GW - 600, GH, alpha:=IIf(Round = 0, 0.3, 1), Pos:=posOnBottom
    
    Page.Writes "塔克", 20, 20, 24, argb(255, 255, 255, 255), style:=FontStyleBold
    Page.Writes "青", GW - 25, GH - 135, 24, argb(255, 255, 255, 255), style:=FontStyleBold, align:=StringAlignmentFar

    Dim BarW As Long, DisplayHP As Long
    BarW = Page.Res.ImgSize("hpbar.png", imgGetWidth)
    
    DisplayHP = UHP
    If Round = 1 Then DisplayHP = UHP - ATKTotal + ATKTotal * (GetTickCount - RoundTime) / 4000
    Page.Writes "HP", 20, 65, 18, argb(200, 255, 255, 255), style:=FontStyleBold
    Page.DrawImage "emptybar.png", 70, 65 + 6, alpha:=0.5
    Page.DrawImage "hpbar.png", 70, 65 + 6, CW:=DisplayHP / UMaxHP * BarW, alpha:=1
    Page.Writes DisplayHP & " / " & UMaxHP, 70 + BarW / 2, 65 + 1, 16, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    
    Page.Writes "AP", 20, 95, 18, argb(200, 255, 255, 255), style:=FontStyleBold
    Page.DrawImage "emptybar.png", 70, 95 + 6, alpha:=0.5
    Page.DrawImage "mpbar.png", 70, 95 + 6, CW:=UMP / UMaxMP * BarW, alpha:=1
    Page.Writes UMP & " / " & UMaxMP, 70 + BarW / 2, 95 + 1, 16, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    
    DisplayHP = HP
    If Round = 2 Then DisplayHP = HP - ATKTotal
    If Round = 3 Then DisplayHP = HP - ATKTotal + ATKTotal * (GetTickCount - RoundTime) / 2000
    Page.Writes "HP", GW - 25, GH - 85 - 6, 18, argb(200, 255, 255, 255), style:=FontStyleBold, align:=StringAlignmentFar
    Page.DrawImage "emptybar.png", GW - BarW - 75, GH - 85, alpha:=0.5
    Page.DrawImage "hpbar.png", GW - BarW - 75, GH - 85, CW:=DisplayHP / MaxHP * BarW, alpha:=1
    Page.Writes DisplayHP & " / " & MaxHP, GW - BarW - 75 + BarW / 2, GH - 85 - 5, 16, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    
    Page.Writes "AP", GW - 25, GH - 55 - 6, 18, argb(200, 255, 255, 255), style:=FontStyleBold, align:=StringAlignmentFar
    Page.DrawImage "emptybar.png", GW - BarW - 75, GH - 55, alpha:=0.5
    Page.DrawImage "mpbar.png", GW - BarW - 75, GH - 55, CW:=MP / MaxMP * BarW, alpha:=1
    Page.Writes MP & " / " & MaxMP, GW - BarW - 75 + BarW / 2, GH - 55 - 5, 16, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    
    If BGM.PlayState = musStopped And BGM.length <> 0 Then BGM.Play
    If BGS.PlayState = musStopped And BGS.length <> 0 Then BGS.Play
    
    If Round = 3 Then
        If HP <= 300 Then
            Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(Cubic((GetTickCount - RoundTime) / 3200, 0, 1, 1, 1) * 255, 255, 0, 0)
            If GetTickCount Mod 300 < 150 Then Page.Writes "我们失败了。", GW / 2, GH / 2 - 24, 48, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
        End If
    End If
    
    If HP <= 300 And Round = 0 Then
        SE.Play "Collapse4.mp3"
        HP = MaxHP: UHP = UMaxHP: MP = MaxMP: UMP = UMaxMP
        LastHP = HP
        ReDim MyCard(0): ReDim YourCard(0): ReDim msg(0)
        Call GiveCards
        BGM.Play
        Round = 0: ECore.NewTransform transFallDark, 2000
    End If
    If UHP <= 300 And Round = 0 Then
        SE.Play "Applause1.mp3"
        NovelPage.RunNovel
        ECore.NewTransform transHighLight, 3000, "NovelPage"
    End If
    
    If UBound(msg) >= 1 Then
        Dim MsgPro As Single
        MsgPro = (GetTickCount - MsgTime) / 500
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
    
    If Round = 0 And CanMove Then
        Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(100, 0, 0, 0)
        If GetTickCount Mod 300 < 150 Then Page.Writes "遭到塔克的恐惧效果，无法出牌。", GW / 2, GH / 2 - 24, 48, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
    End If
End Sub
Public Sub GiveCards()
    Randomize
    Dim index As Long
    
    If Not CanMove Then
        For i = 1 To 3 + Int(Rnd * 3)
            index = Int(Rnd * (UBound(ICard) + 1))
            If index > UBound(ICard) Then index = UBound(ICard)
            If UBound(MyCard) >= 7 Then Exit For
            ReDim Preserve MyCard(UBound(MyCard) + 1)
            MyCard(UBound(MyCard)) = ICard(index)
            MyCard(UBound(MyCard)).GetTime = GetTickCount + 100 * i
        Next
    End If
    
    For i = 1 To 3 + Int(Rnd * 3)
        index = Int(Rnd * (UBound(UCard) + 1))
        If index > UBound(UCard) Then index = UBound(UCard)
        If UBound(YourCard) >= 7 Then Exit For
        ReDim Preserve YourCard(UBound(YourCard) + 1)
        YourCard(UBound(YourCard)) = UCard(index)
    Next
    
    UMP = UMP + 50
    MP = MP + 50
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
    BGM.Create App.Path & "\music\bgm\OST 23 - A Lost World.mp3"
    BGM.Play
    BGS.Create App.Path & "\music\bgs\Wind_1_3D_loop.mp3"
    BGS.Play
End Sub
Public Sub Leave()
    '页面离开事件
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
    Page.Res.HotLoad = True
    Page.Res.NoLoadCrash = True
    Page.Res.NewImages App.Path & "\assets\battle\images", 1.8
    Page.LoadAnimationsFromDir App.Path & "\battle_animation"
    
    Page.CreatePlayAnimation "rpgmv_Darkness1", "暗黑导弹", "Reserve"
    Page.CreatePlayAnimation "rpgmv_Darkness2", "黑幕", "Reserve"
    Page.CreatePlayAnimation "rpgmv_Darkness4", "暗黑撕裂", "Reserve"
    Page.CreatePlayAnimation "rpgmv_StateDeath", "死亡之钟", "Reserve"
    
    Page.CreatePlayAnimation "rpgmv_Laser1", "青铜炮", "Reserve"
    Page.CreatePlayAnimation "rpgmv_Laser2", "翡翠炮", "Reserve"
    Page.CreatePlayAnimation "rpgmv_Meteor+Gun2", "青光圣堂", "Reserve"
    Page.CreatePlayAnimation "rpgmv_Magic2", "翡翠导弹", "Reserve"
    Page.CreatePlayAnimation "rpgmv_Recovery1", "能力回流", "Reserve"
    
    Page.CreatePlayAnimation "rpgmv_Recovery2", "灵力祭坛", "Reserve"
    
    ReDim UCard(4)
    With UCard(0)
        .Name = "暗黑导弹"
        .BasicATK = 36
        .AP = 30
    End With
    With UCard(1)
        .Name = "黑幕"
        .BasicATK = 50
        .AP = 40
    End With
    With UCard(2)
        .Name = "暗黑撕裂"
        .BasicATK = 320
        .AP = 400
    End With
    With UCard(4)
        .Name = "死亡之钟"
        .BasicATK = 0
        .AP = 100
    End With
    With UCard(3)
        .Name = "灵力祭坛"
        .BasicATK = -200
        .AP = 0
    End With
    
    ReDim ICard(5)
    With ICard(0)
        .Name = "青铜炮"
        .BasicATK = 30
        .AP = 30
    End With
    With ICard(1)
        .Name = "翡翠炮"
        .BasicATK = 60
        .AP = 50
    End With
    With ICard(2)
        .Name = "青光圣堂"
        .BasicATK = 300
        .AP = 400
    End With
    With ICard(3)
        .Name = "翡翠导弹"
        .BasicATK = 170
        .AP = 100
    End With
    With ICard(5)
        .Name = "能力回流"
        .BasicATK = 0
        .AP = 15
    End With
    With ICard(4)
        .Name = "灵力祭坛"
        .BasicATK = -200
        .AP = 0
    End With
    
    '创建页面
    ECore.Add Page, "TLPPage"
    MaxHP = 3000: HP = MaxHP: MaxMP = 2000: MP = MaxMP
    UMaxHP = 3300: UHP = UMaxHP: UMaxMP = 1700: UMP = UMaxMP
    LastHP = HP
    ReDim MyCard(0): ReDim YourCard(0): ReDim msg(0)
    Call GiveCards
    
    Page.Res.newImage App.Path & "\assets\fg\DemonicWorld.png", GW, GH
    Page.Res.newImage App.Path & "\assets\bg\GrassMaze.png", GW, GH
    Page.Res.newImage App.Path & "\assets\face\tk-normal.png"
    Page.Res.NewImages App.Path & "\assets\tipui"
    Page.Res.NewImages App.Path & "\assets\tlpui", 0.5
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
