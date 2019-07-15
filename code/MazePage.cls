VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MazePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim Map() As Integer
    Dim PX As Single, PY As Single
    Dim DarkX As Single, DarkY As Single, FollowStep As Long
    Dim FX(5) As Single, FY(5) As Single
    Dim StartX As Long, StartY As Long, EndX As Long, EndY As Long
    Dim XRemain As Single, YRemain As Single
    Dim MapDC As Long, CoverDC As Long
    Dim Direction As Long, MapNames As String
'=====================================
    'Snowstorm
    Dim Hp As Long, MaxHp As Long, LastHp As Long
    Dim MP As Long, MaxMP As Long
    Dim UHP As Long, UMaxHp As Long
    Dim ATKS() As Single, ATKIndex As Long
    Dim UGun() As Fires
    Private Type Fires
        x As Single
        y As Single
        Kind As Long
        Step As Long
        Time As Long
        Hurt As Long
        OldX As Single
        OldY As Single
    End Type
    Dim Crash As New GCrashBox, CanAtk As Boolean
    Dim XLMode As Long, LastXLMode As Long
    Dim DIEMark As Boolean
    Dim MagicTime As Long, ReTime As Long
    Dim Says As String, SayTime As Long
'======================================
Public Sub LoadMap(Optional MapName As String = "")
    Dim W As Long, h As Long, Image As Long, Color As Long

    MapNames = MapName

    Page.Res.newImage App.Path & "\assets\ui\maze" & MapName & ".png"
    Page.Res.newImage App.Path & "\assets\ui\shadow.png", GW, GH
    Image = Page.Res.ResourceImageHandle(Page.Res.GetImage("maze" & MapName & ".png"))
    GdipGetImageWidth Image, W: GdipGetImageHeight Image, h
    
    MapDC = CreateCDC(W * 90, h * 90 + 90)
    CoverDC = CreateCDC(W * 90, h * 90 + 90)
    
    ReDim Map(W + 1, h + 2)
    For y = 0 To h
        For x = 0 To W
            GdipBitmapGetPixel Image, x, y, Color
            If Color = 0 Then Map(x, y + 1) = Page.Res.GetImage("alpha" & MapName & ".png")
            If Color = argb(255, 255, 255, 255) Then Map(x, y + 1) = Page.Res.GetImage("white" & MapName & ".png")
            If Color = argb(255, 153, 204, 255) Then Map(x, y + 1) = Page.Res.GetImage("blue" & MapName & ".png")
            If Color = argb(255, 255, 153, 51) Then Map(x, y + 1) = Page.Res.GetImage("orange" & MapName & ".png")
            If Color = argb(255, 0, 255, 0) Then Map(x, y + 1) = Page.Res.GetImage("green" & MapName & ".png")
            If Color = argb(255, 0, 0, 0) Then Map(x, y + 1) = Page.Res.GetImage("black" & MapName & ".png")
            If Color = argb(255, 0, 255, 255) Then Map(x, y + 1) = Page.Res.GetImage("alpha" & MapName & ".png"): StartX = x: StartY = y + 1
            If Color = argb(255, 255, 0, 0) Then Map(x, y + 1) = Page.Res.GetImage("alpha" & MapName & ".png"): EndX = x: EndY = y + 1
        Next
    Next
    PX = StartX: PY = StartY
    
    Dim ODC As Long
    ODC = Page.CDC
    Page.CDC = MapDC
    For x = 0 To W
        For y = 1 To h + 1
            Page.DrawImage Map(2, 2), (x - 0) * 90, (y - 0) * 90, alpha:=1
            If Page.Res.ImgSize(Map(x, y), imgGetHeight) = 180 Then
                Page.DrawImage Map(x, y), (x - 0) * 90, (y - 0) * 90, cy:=90, ch:=90, alpha:=1
                Page.CDC = CoverDC
                Page.DrawImage Map(x, y), (x - 0) * 90, (y - 0) * 90 - 90, cy:=0, ch:=90, alpha:=1
                Page.CDC = MapDC
            Else
                Page.DrawImage Map(x, y), (x - 0) * 90, (y - 0) * 90, alpha:=1
            End If
        Next
    Next
    Page.CDC = ODC
    
    Select Case MapName
        Case ""
            BGM.Create App.Path & "\music\bgm\OST 4 - Garden Party.mp3"
            BGM.rate = 0.8
            BGM.Play
            BGS.Create App.Path & "\music\bgs\Nightforest.mp3"
            BGS.Play
        Case "snow"
            BGM.Create App.Path & "\music\bgm\OST 15 - Milk Tea.mp3"
            BGM.rate = 0.8
            BGM.Play
            BGS.Create App.Path & "\music\bgs\Nightforest.mp3"
            BGS.Play
            WeatherLayer.ChangeWeather "snow"
        Case "snowstorm"
            BGM.Create App.Path & "\music\bgm\OST 16 - Snowstorm -euphoria-.mp3"
            BGM.rate = 1
            BGM.Play
            BGS.Create App.Path & "\music\bgs\Wind_1_3D_loop.mp3"
            BGS.Play
            WeatherLayer.ChangeWeather "snowstorm"
            UMaxHp = 2000: UHP = UMaxHp
            MaxHp = 4000: Hp = MaxHp: LastHp = MaxHp
            MaxMP = 2000: MP = MaxMP
    End Select
End Sub
Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    If MapNames = "snowstorm" Then Call UpdateSnowStorm: Exit Sub
    
    If MapNames = "" Then
        Page.Clear argb(255, 123, 199, 47) '清空画布
    Else
        Page.Clear argb(255, 242, 242, 255) '清空画布
    End If
    
    Dim W As Long, h As Long, cx As Long, cy As Long
    cx = (PX - 0) * 90 - GW / 2 + 45
    cy = (PY - 0) * 90 - GH / 2 + 45
    
    W = 90 * (UBound(Map, 1) - 1) - IIf(cx >= 0, cx, 0)
    h = 90 * (UBound(Map, 2) - 1) - IIf(cy >= 0, cy, 0)
    
    If W > GW Then W = GW
    If h > GH Then h = GH
    
    Dim pro As Single
    pro = 1 - (Abs(EndX - PX) + Abs(EndY - PY)) / 5
    If pro < 0 Then pro = 0
    If pro > 1 Then pro = 1
    
    Dim Walking As Boolean
    
    If XRemain <> 0 Then
        If Abs(XRemain) < 0.1 Then
            For i = 0 To 4
                FX(i) = FX(i + 1): FY(i) = FY(i + 1)
            Next
            XRemain = 0: PX = Round(PX)
        Else
            PX = PX + IIf(XRemain < 0, -0.05, 0.05)
            XRemain = XRemain - IIf(XRemain < 0, -0.05, 0.05)
            If FollowStep <= 50 Then FollowStep = FollowStep + 1
            If FollowStep >= 50 Then FX(5) = PX: FY(5) = PY
            Walking = True
        End If
    End If
    If YRemain <> 0 Then
        If Abs(YRemain) < 0.1 Then
            For i = 0 To 4
                FX(i) = FX(i + 1): FY(i) = FY(i + 1)
            Next
            YRemain = 0: PY = Round(PY)
        Else
            PY = PY + IIf(YRemain < 0, -0.05, 0.05)
            YRemain = YRemain - IIf(YRemain < 0, -0.05, 0.05)
            If FollowStep <= 50 Then FollowStep = FollowStep + 1
            If FollowStep >= 50 Then FX(5) = PX: FY(5) = PY
            Walking = True
        End If
    End If
    
    PaintDC MapDC, Page.CDC, IIf(cx < 0, -cx, 0), IIf(cy < 0, -cy, 0), IIf(cx >= 0, cx, 0), IIf(cy >= 0, cy, 0), W, h, 1
    If MapNames = "" Then Page.DrawImage "darkssr.png", (DarkX - 0) * 90 - cx, (DarkY - 0) * 90 - cy, alpha:=1 - pro * 4.6
    Page.DrawImage "qing" & IIf(Walking, IIf(GetTickCount Mod 300 < 150, Direction, Direction + 2), Direction + 1) & ".png", GW / 2 - 45, GH / 2 - 45, alpha:=1

    'Page.Paint 1, GW / 2 - 45, GH / 2 - 45, 90, 90, argb(255, 255, 255, 255)
    PaintDC CoverDC, Page.CDC, IIf(cx < 0, -cx, 0), IIf(cy < 0, -cy, 0), IIf(cx >= 0, cx, 0), IIf(cy >= 0, cy, 0), W, h, 1
    
    Page.DrawImage "shadow.png", 0, 0, alpha:=(1 - pro * 1) * IIf(MapNames = "", 1, 0.9)
    Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(pro * 255, 255, 255, 255)
    
    If pro = 1 Then
        BGM.rate = 1
        BGS.Dispose
        NovelPage.RunNovel
        SE.Play "Defense_3D.mp3"
        ECore.NewTransform transHighLight, 3000, "NovelPage"
    End If
    
    Dim HuaIndex As Integer
    HuaIndex = Page.Res.GetImage("black" & MapNames & ".png")
    
    DarkX = IIf(Abs(PX - FX(0)) < 2 And pro = 0, 0, FX(0)): DarkY = IIf(Abs(PY - FY(0)) < 2 And pro = 0, 0, FY(0))
    If pro > 0 Then
        DarkX = PX - 5
        DarkY = PY
    End If
    
    If IsKeyPress(VK_UP) And Map(Round(PX), Round(PY - 1)) <> HuaIndex And XRemain = 0 And YRemain = 0 Then YRemain = -1: Direction = 0
    If IsKeyPress(VK_DOWN) And Map(Round(PX), Round(PY + 1)) <> HuaIndex And XRemain = 0 And YRemain = 0 Then YRemain = 1: Direction = 3
    If IsKeyPress(VK_LEFT) And Map(Round(PX - 1), Round(PY)) <> HuaIndex And XRemain = 0 And YRemain = 0 Then XRemain = -1: Direction = 6
    If IsKeyPress(VK_RIGHT) And Map(Round(PX + 1), Round(PY)) <> HuaIndex And XRemain = 0 And YRemain = 0 Then XRemain = 1: Direction = 9
    
    If BGM.PlayState = musStopped And BGM.length <> 0 Then BGM.Play
    If BGS.PlayState = musStopped And BGS.length <> 0 Then BGS.Play
    
    'snow
    If MapNames = "999" Then
        StretchBlt Page.CDC, IIf(PX > UBound(Map, 1) / 2, GW, 0), IIf(PY < UBound(Map, 2) / 2, GH, 0), IIf(PX > UBound(Map, 1) / 2, -GW, GW), IIf(PY < UBound(Map, 2) / 2, -GH, GH), Page.CDC, 0, 0, GW, GH, vbSrcCopy
    End If
End Sub
Public Sub Attacks(i As Integer)
    LastHp = Hp
    Dim ATK As Long
    Select Case i
        Case 0
            ATK = 15 + Int(Rnd * 6 - 3)
        Case 1
            ATK = 30 + Int(Rnd * 14 - 7)
        Case Is >= 3
            ATK = 10 + Int(Rnd * 6 - 3)
    End Select
    Call NewUGun(11)
    If XLMode = 6 Or XLMode = 8 Then ATK = ATK * 2
    Hp = Hp - ATK
End Sub
Public Sub UpdateSnowStorm()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear argb(255, 242, 242, 255) '清空画布
    
    Dim W As Long, h As Long, cx As Long, cy As Long
    cx = (PX - 0) * 90 - GW / 2 + 45
    cy = (PY - 0) * 90 - GH / 2 + 45
    
    W = 90 * (UBound(Map, 1) - 1) - IIf(cx >= 0, cx, 0)
    h = 90 * (UBound(Map, 2) - 1) - IIf(cy >= 0, cy, 0)
    
    If W > GW Then W = GW
    If h > GH Then h = GH
    
    Dim Walking As Boolean
    
    If XRemain <> 0 Then
        If Abs(XRemain) < 0.1 Then
            For i = 0 To 4
                FX(i) = FX(i + 1): FY(i) = FY(i + 1)
            Next
            XRemain = 0 ': PX = Round(PX)
        Else
            PX = PX + IIf(XRemain < 0, -0.2, 0.2)
            XRemain = XRemain - IIf(XRemain < 0, -0.2, 0.2)
            If FollowStep <= 50 Then FollowStep = FollowStep + 1
            If FollowStep >= 50 Then FX(5) = PX: FY(5) = PY
            Walking = True
        End If
    End If
    If YRemain <> 0 Then
        If Abs(YRemain) < 0.1 Then
            For i = 0 To 4
                FX(i) = FX(i + 1): FY(i) = FY(i + 1)
            Next
            YRemain = 0 ': PY = Round(PY)
        Else
            PY = PY + IIf(YRemain < 0, -0.2, 0.2)
            YRemain = YRemain - IIf(YRemain < 0, -0.2, 0.2)
            If FollowStep <= 50 Then FollowStep = FollowStep + 1
            If FollowStep >= 50 Then FX(5) = PX: FY(5) = PY
            Walking = True
        End If
    End If
    
    PaintDC MapDC, Page.CDC, IIf(cx < 0, -cx, 0), IIf(cy < 0, -cy, 0), IIf(cx >= 0, cx, 0), IIf(cy >= 0, cy, 0), W, h, 1

    Page.DrawImage "qing" & IIf(Walking, IIf(GetTickCount Mod 300 < 150, Direction, Direction + 2), Direction + 1) & ".png", GW / 2 - 45, GH / 2 - 45, alpha:=1
    Crash.Reset
    CanAtk = (Abs(EndX - PX) <= 6 - (1 - (UHP / UMaxHp)) * 3 And Abs(EndY - PY) <= 6 - (1 - (UHP / UMaxHp)) * 3)
    If DIEMark Then CanAtk = False
    
    Select Case UHP / UMaxHp
        Case Is > 0.99
            XLMode = 0
            If XLMode <> LastXLMode Then Says = "那么，欢迎你来到圣光核心检查点。": SayTime = GetTickCount
        Case Is > 0.95
            XLMode = 1
            If XLMode <> LastXLMode Then Says = "在绝望之中哀嚎吧。": SayTime = GetTickCount
        Case Is > 0.9
            XLMode = 8
            If XLMode <> LastXLMode Then Says = "我会全力击溃你。": SayTime = GetTickCount
        Case Is > 0.85
            XLMode = 2
            If XLMode <> LastXLMode Then Says = "比我想象中的厉害嘛。": SayTime = GetTickCount
        Case Is > 0.7
            XLMode = 3
            If XLMode <> LastXLMode Then Says = "来吧，看你打得快还是我恢复得快。": SayTime = GetTickCount
        Case Is > 0.6
            XLMode = 1
            If XLMode <> LastXLMode Then Says = "你应该好好看看你身后。": SayTime = GetTickCount
        Case Is > 0.5
            XLMode = 6
            If XLMode <> LastXLMode Then Says = "你是谁啊？": SayTime = GetTickCount
        Case Is > 0.4
            XLMode = 5
            If XLMode <> LastXLMode Then Says = "你这家伙到底是什么东西啊？": SayTime = GetTickCount
        Case Is >= 300
            XLMode = 7
        Case Is <= 300
            XLMode = 9
            If XLMode <> LastXLMode Then Says = "我。。。": SayTime = GetTickCount
    End Select
    If GetTickCount - SayTime >= 5000 Then
        Randomize
        Dim RI As Integer
        RI = Int(Rnd * 5)
        Select Case XLMode
            Case 0
                Says = "那么，欢迎你来到圣光核心检查点。"
            Case 1
                Select Case RI
                    Case 0: Says = "你的血是什么颜色的呢？"
                    Case 1: Says = "哎呀。尖叫吧。"
                    Case 2: Says = "放心吧，我会让你毫无痛苦地死去的。"
                    Case 3: Says = "你尖叫的声音真是好听。"
                    Case Else: Says = "请原谅我，这也是不可避免的。"
                End Select
            Case 8
                Select Case RI
                    Case 0: Says = "我可不会放水。"
                    Case 1: Says = "这个时候或许防御是最佳选择？"
                    Case 2: Says = "我会使出我的全力。"
                    Case 3: Says = "也请你认真地和我战斗吧。"
                    Case Else: Says = "你可别期望我会放过你。"
                End Select
            Case 2
                Select Case RI
                    Case 0: Says = "看来我小看你了呢。"
                    Case 1: Says = "你是灵魂吧？"
                    Case 2: Says = "真有意思。"
                    Case 3: Says = "真是低估你了。"
                    Case Else: Says = "这么强大的力量，从哪里得来的？"
                End Select
            Case 3
                Select Case RI
                    Case 0: Says = "我可是忍受着你的攻击在这里补血呢。"
                    Case 1: Says = "这最终只是一个赌注罢了。"
                    Case 2: Says = "你的攻击可以打破奇迹吗？"
                    Case 3: Says = "恢复体力也是一种计策。"
                    Case Else: Says = "哎呀。不能安静地让我休息一下吗？"
                End Select
            Case 1
                Select Case RI
                    Case 0: Says = "去死吧。"
                    Case 1: Says = "出其不意地杀死你。"
                    Case 2: Says = "你的死将变得没有理由。"
                    Case 3: Says = "放弃吧。"
                    Case Else: Says = "你可以去死一下吗？"
                End Select
            Case 6
                Select Case RI
                    Case 0: Says = "你谁啊？"
                    Case 1: Says = "你真的是青？"
                    Case 2: Says = "你究竟是什么人？"
                    Case 3: Says = "你到底是谁？"
                    Case Else: Says = "你不是青吧？"
                End Select
            Case 9
                Select Case RI
                    Case 0: Says = "我。。。"
                    Case 1: Says = "。。。"
                    Case 2: Says = "人家。。。"
                    Case 3: Says = "我。。。"
                    Case Else: Says = "。。。"
                End Select
        End Select
        SayTime = GetTickCount
    End If
    
    If XLMode <> LastXLMode Then
        LastXLMode = XLMode
        ECore.NewTransform transHighLight, 100
    End If
    
    Dim RemoveCount As Long, pro As Single, MapW As Long, MapH As Long
    Dim DrawX As Long, DrawY As Long
    MapW = UBound(Map, 1)
    MapH = UBound(Map, 2)
    
    For i = 1 To UBound(UGun)
Start:
        If i > UBound(UGun) - RemoveCount Then Exit For
        If i <= 0 Then Exit For

        Select Case UGun(i).Kind
            Case 0
                pro = 1
                If UGun(i).Step < 8 Then
                    If GetTickCount - UGun(i).Time >= 41 Then UGun(i).Step = UGun(i).Step + 1: UGun(i).Time = GetTickCount
                    If UGun(i).Step = 8 Then SE.Play "Ice4.mp3"
                Else
                    pro = 1 - (GetTickCount - UGun(i).Time) / 2000
                    If GetTickCount - UGun(i).Time >= 2000 Then GoTo Remove
                End If
                DrawX = UGun(i).x * 90 - cx + 45: DrawY = UGun(i).y * 90 - cy + 45
                If DrawX >= 0 And DrawY >= 0 And DrawX <= GW And DrawY <= GH Then
                    Page.DrawImage "Ice3_00" & UGun(i).Step & ".png", DrawX, DrawY, alpha:=pro, Pos:=posOnCenter
                    If pro > 0.6 And UGun(i).Step >= 7 And (Not Crash.CrashResult) Then
                        If Crash.IsCrash Then Attacks (i): GoTo Remove
                    End If
                End If
            Case 1
                If GetTickCount - UGun(i).Time <= 200 Then
                    UGun(i).OldX = Int(PX - UGun(i).x) / 30
                    UGun(i).OldY = Int(PY - UGun(i).y) / 30
                End If
                UGun(i).x = UGun(i).x + UGun(i).OldX
                UGun(i).y = UGun(i).y + UGun(i).OldY
                DrawX = UGun(i).x * 90 - cx: DrawY = UGun(i).y * 90 - cy
                If DrawX >= 0 And DrawY >= 0 And DrawX <= GW And DrawY <= GH Then
                    Page.DrawImage "Ice1_014.png", DrawX, DrawY, alpha:=1
                    If (Not Crash.CrashResult) Then
                        If Crash.IsCrash Then Attacks (i): GoTo Remove
                    End If
                    If UGun(i).x < -5 Or UGun(i).y < -5 Or UGun(i).x > MapW + 5 Or UGun(i).y > MapH + 5 Then GoTo Remove
                End If
            Case 11
                pro = 1
                If UGun(i).Step < 7 Then
                    If GetTickCount - UGun(i).Time >= 41 Then UGun(i).Step = UGun(i).Step + 1: UGun(i).Time = GetTickCount
                Else
                    pro = 1 - (GetTickCount - UGun(i).Time) / 500
                    If GetTickCount - UGun(i).Time >= 500 Then GoTo Remove
                End If
                DrawX = UGun(i).x * 90 - cx + 45: DrawY = UGun(i).y * 90 - cy + 45
                If DrawX >= 0 And DrawY >= 0 And DrawX <= GW And DrawY <= GH Then
                    Page.DrawImage "HitIce_00" & UGun(i).Step & ".png", DrawX, DrawY, alpha:=pro, Pos:=posOnCenter
                End If
            Case 12
                pro = 1
                If UGun(i).Step < 23 Then
                    If GetTickCount - UGun(i).Time >= 20 Then UGun(i).Step = UGun(i).Step + 1: UGun(i).Time = GetTickCount
                Else
                    GoTo Remove
                End If
                DrawX = UGun(i).x * 90 - cx + 45: DrawY = UGun(i).y * 90 - cy + 45
                If DrawX >= 0 And DrawY >= 0 And DrawX <= GW And DrawY <= GH Then
                    Page.DrawImage "Fire3_0" & format(UGun(i).Step + 1, "00") & ".png", DrawX, DrawY, alpha:=1, Pos:=posOnCenter
                End If
            Case 13
                pro = 1
                If UGun(i).Step < 13 Then
                    If GetTickCount - UGun(i).Time >= 20 Then UGun(i).Step = UGun(i).Step + 1: UGun(i).Time = GetTickCount
                Else
                    GoTo Remove
                End If
                DrawX = UGun(i).x * 90 - cx + 45: DrawY = UGun(i).y * 90 - cy + 45
                If DrawX >= 0 And DrawY >= 0 And DrawX <= GW And DrawY <= GH Then
                    Page.DrawImage "Recovery1_0" & format(UGun(i).Step + 1, "00") & ".png", DrawX, DrawY, alpha:=1, Pos:=posOnCenter
                End If
            Case Is >= 3
                UGun(i).x = UGun(i).x + UGun(i).OldX
                UGun(i).y = UGun(i).y + UGun(i).OldY
                DrawX = UGun(i).x * 90 - cx: DrawY = UGun(i).y * 90 - cy
                If DrawX >= 0 And DrawY >= 0 And DrawX <= GW And DrawY <= GH Then
                    Page.DrawImage "Ice1_013.png", DrawX, DrawY, alpha:=1
                    If (Not Crash.CrashResult) Then
                        If Crash.IsCrash Then Attacks (i): GoTo Remove
                    End If
                    If UGun(i).x < -5 Or UGun(i).y < -5 Or UGun(i).x > MapW + 5 Or UGun(i).y > MapH + 5 Then GoTo Remove
                End If
        End Select

        i = i + 1
        GoTo Start
Remove:
        UGun(i) = UGun(UBound(UGun))
        RemoveCount = RemoveCount + 1
        i = i - 1
    Next
    On Error Resume Next
    If RemoveCount <> 0 Then ReDim Preserve UGun(UBound(UGun) - RemoveCount)
    
    PaintDC CoverDC, Page.CDC, IIf(cx < 0, -cx, 0), IIf(cy < 0, -cy, 0), IIf(cx >= 0, cx, 0), IIf(cy >= 0, cy, 0), W, h, 1
    Page.DrawImage "xl" & XLMode & ".png", EndX * 90 - cx, EndY * 90 - cy, alpha:=1
    
    Page.DrawImage "shadow.png", 0, 0, alpha:=0.6 + Cubic((1 - (UHP / UMaxHp)), 0, 0, 0, 1) * 0.4

    If DIEMark = True And Abs(PX - EndX) <= 1 And PY >= EndY And PY - EndY <= 1 And (Not BGM.IsPlaying) And Direction = 0 Then
        BGM.rate = 1
        BGS.Dispose
        NovelPage.RunNovel
        SE.Play "Defense_3D.mp3"
        ECore.NewTransform transHighLight, 3000, "NovelPage"
    End If
    
    Dim HuaIndex As Integer
    HuaIndex = Page.Res.GetImage("black" & MapNames & ".png")
    
    If IsKeyPress(VK_UP) And Map(Round(PX), Round(PY - 1)) <> HuaIndex And XRemain = 0 And YRemain = 0 Then YRemain = -0.2: Direction = 0
    If IsKeyPress(VK_DOWN) And Map(Round(PX), Round(PY + 1)) <> HuaIndex And XRemain = 0 And YRemain = 0 Then YRemain = 0.2: Direction = 3
    If IsKeyPress(VK_LEFT) And Map(Round(PX - 1), Round(PY)) <> HuaIndex And XRemain = 0 And YRemain = 0 Then XRemain = -0.2: Direction = 6
    If IsKeyPress(VK_RIGHT) And Map(Round(PX + 1), Round(PY)) <> HuaIndex And XRemain = 0 And YRemain = 0 Then XRemain = 0.2: Direction = 9
    
    If BGM.PlayState = musStopped And BGM.length <> 0 And DIEMark = False Then BGM.Play: ATKIndex = 1
    If BGS.PlayState = musStopped And BGS.length <> 0 Then BGS.Play
    
    Dim Hot As Long, Map2() As Single
    If BGM.IsPlaying Then
        Map2 = BGM.GetMusicMap
        For i = 0 To UBound(Map2)
            Hot = Hot + Map2(i)
        Next
        If ATKIndex <> 1 And ATKIndex <> UBound(ATKS) Then Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(IIf(Hot / 200000 * 170 <= 170, Hot / 200000 * 170, 170), 0, 0, 0)
    End If
    
    If BGM.position > ATKS(ATKIndex) And ATKIndex < UBound(ATKS) And (Not DIEMark) Then
        ATKIndex = ATKIndex + 1
        Call NewUGun(0)
        If XLMode = 4 Then UHP = UMaxHp
        If XLMode <> 7 Then Call NewUGun(1)
        If XLMode = 3 Or XLMode = 7 Then UHP = UHP + 1
        If XLMode >= 6 Then Call NewUGun(2)
        If XLMode = 1 Or XLMode = 5 Then Call NewUGun(15)
        If XLMode >= 3 Then
            For i = 3 To 10
                Call NewUGun(i)
            Next
        End If
    End If
    
    If CanAtk And MP >= 3 Then Page.DrawAnimation "attackbg", GW - 200, GH - 100
    Page.DrawImage "atk" & IIf(CanAtk And MP >= 3, 1, 0) & ".png", GW - 200, GH - 100, alpha:=1, Pos:=posOnCenter
    Page.Writes "Z", GW - 200 + 25, GH - 100 + 20, 24, argb(255, 255, 255, 255), style:=FontStyleBold
    
    If (Not CanAtk) And MP >= 1 Then Page.DrawAnimation "attackbg", GW - 80, GH - 100
    Page.DrawImage "def" & IIf((Not CanAtk) And MP >= 1, 1, 0) & ".png", GW - 80, GH - 100, alpha:=1, Pos:=posOnCenter
    Page.Writes "X", GW - 80 + 25, GH - 100 + 20, 24, argb(255, 255, 255, 255), style:=FontStyleBold
    
    Page.DrawImage "uhpframe.png", 0, 40, alpha:=1
    Page.DrawImage "uhpbar.png", 20, 122, 0, 0, UHP / UMaxHp * 196, alpha:=1
    Page.Writes "雪狼", 15, 47, 28, argb(255, 255, 255, 255)
    Page.Writes UHP & "/" & UMaxHp, 160, 61, 18, argb(255, 255, 255, 255), 300, 25, StringAlignmentNear, FontStyleItalic
    
    Page.DrawImage "mehpframe.png", GW, 40, alpha:=1, Pos:=posOnRight
    Page.DrawImage "mehpbar.png", GW - 16, 121, 182 - Hp / MaxHp * 182, 0, Hp / MaxHp * 182, alpha:=1, Pos:=posOnRight
    Page.Writes "青", GW - 15, 48, 28, argb(255, 255, 255, 255), align:=StringAlignmentFar
    Page.Writes Hp & "/" & MaxHp, GW - 160 - 80, 61, 18, argb(255, 255, 255, 255), 300, 25, StringAlignmentNear, FontStyleItalic
    
    Page.DrawImage "mpbar.png", GW / 2, GH - 14, cw:=MP / MaxMP * GW, alpha:=1, Pos:=posOnCenter
    Page.Writes "MP     " & MP & "/" & MaxMP, GW / 2, GH - 50, 18, argb(255, 255, 255, 255), 0, 0, StringAlignmentCenter, FontStyleBold
    
    Page.Paint 2, GW / 2, GH - 105, EF.GetWidth(Page.GG, Says, 18, StringAlignmentCenter, FontStyleRegular) + 80, 40, argb(128, 0, 0, 0), 80, Pos:=posOnCenter
    Page.Writes Says, GW / 2, GH - 105 - 14, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    
    Dim ModeText As String, ModeText2 As String
    
    Select Case XLMode
        Case 0: ModeText = "雪狼の正常形态": ModeText2 = "所有能力正常"
        Case 1: ModeText = "雪狼の高速形态": ModeText2 = "当心！雪狼将进行瞬移！"
        Case 2: ModeText = "雪狼の防御形态": ModeText2 = "雪狼受到的伤害将减少25%！"
        Case 3: ModeText = "雪狼の回复形态": ModeText2 = "嘘~现在雪狼需要停下来回复体力了。"
        Case 4: ModeText = "雪狼の再生形态": ModeText2 = "？？？？用修改器可不好哦？？？？"
        Case 5: ModeText = "雪狼の逃生形态": ModeText2 = "现在雪狼跑的很快，并且受到的伤害将减少25%！"
        Case 6: ModeText = "雪狼の强攻形态": ModeText2 = "雪狼可能在一瞬间杀死你，现在雪狼的攻击力是原来的200%！"
        Case 7: ModeText = "雪狼の自救形态": ModeText2 = "雪狼现在只能一边防御一边回复体力了。"
        Case 8: ModeText = "雪狼の攻击形态": ModeText2 = "雪狼现在的攻击力是原来的200%！"
        Case 9: ModeText = "雪狼の濒死形态": ModeText2 = "雪狼快要撑不下去了。"
    End Select
    
    Page.Writes ModeText, GW / 2, 90, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
    Page.Writes ModeText2, GW / 2, 130, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    
    'page.Writes says,gw / 2 - 100,gh - 100,18
    
    If UHP < 300 And DIEMark = False Then
        DIEMark = True
        BGM.position = 60 + 47
        WeatherLayer.ChangeWeather ""
        XLMode = 9
    End If
    
    If Hp < 300 Then
        BGM.Stops
        GameWindow.BackColor = RGB(120, 0, 0)
        MsgBox "真是太失败了。。。我们再来一次吧？", , "雪狼"
        SE.Play "Collapse4.mp3"
        GameWindow.Hide
        Sleep 3000
        Unload GameWindow
        End
    End If
    
    If GetTickCount - MagicTime >= 2000 Then
        If GetTickCount - ReTime >= 250 Then
            MP = MP + 1
            ReTime = GetTickCount
            If MP >= MaxMP Then MP = MaxMP
        End If
    End If
End Sub
Public Sub KeyDown(Code As Integer)
    If MapNames <> "snowstorm" Then Exit Sub
    
    If Code = vbKeyZ And CanAtk And MP >= 3 Then
        Dim ATK As Long
        ATK = 4 + Int(Rnd * 2 - 1)
        If XLMode = 2 Or XLMode = 5 Or XLMode = 7 Then ATK = ATK * 0.75
        UHP = UHP - ATK
        MP = MP - 3
        MagicTime = GetTickCount
        Call NewUGun(12)
    ElseIf Code = vbKeyX And (Not CanAtk) And MP >= 1 Then
        MagicTime = GetTickCount
        Call NewUGun(13)
        MP = MP - 1
        Hp = LastHp
    End If
End Sub
Public Sub NewUGun(ByVal Kind As Long)
    Randomize
    If Kind = 2 Then
        EndX = Int(Rnd * (UBound(Map, 1) - 5) + 2)
        EndY = Int(Rnd * (UBound(Map, 2) - 5) + 2)
        Exit Sub
    End If
    If Kind = 15 Then
        EndX = PX: EndY = PY
        If Direction = 0 Then EndY = PY + 1
        If Direction = 3 Then EndY = PY - 1
        If Direction = 6 Then EndX = PX + 1
        If Direction = 9 Then EndX = PX - 1
        Exit Sub
    End If
    
    ReDim Preserve UGun(UBound(UGun) + 1)
    With UGun(UBound(UGun))
        .Kind = Kind
        .Time = GetTickCount
        Select Case Kind
            Case 0
                If Int(Rnd * 6) = 3 Then
                    .x = PX
                    .y = PY
                Else
                    .x = Int(Rnd * (UBound(Map, 1) - 5) + 2)
                    .y = Int(Rnd * (UBound(Map, 2) - 5) + 2)
                End If
            Case 11
                .x = PX
                .y = PY
            Case 13
                .x = PX
                .y = PY
            Case Else
                .x = EndX
                .y = EndY
                If .Kind = 3 Or .Kind = 4 Or .Kind = 10 Then .OldY = -0.2
                If .Kind = 4 Or .Kind = 5 Or .Kind = 6 Then .OldX = 0.2
                If .Kind = 6 Or .Kind = 7 Or .Kind = 8 Then .OldY = 0.2
                If .Kind = 8 Or .Kind = 9 Or .Kind = 10 Then .OldX = -0.2
        End Select
        Select Case Kind
            Case 1
                SE.Play "Ice1.mp3"
            Case 2
                SE.Play "Ice3.mp3"
            Case 3
                SE.Play "Ice2.mp3"
            Case 11
                SE.Play "Ice_Hit.mp3"
            Case 12
                SE.Play "Fire3.mp3"
            Case 13
                SE.Play "Magic1.mp3"
        End Select
    End With
End Sub
    
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(ID As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
    Says = "那么，欢迎你来到圣光核心检查点。": SayTime = GetTickCount
End Sub
Public Sub Leave()
    '页面离开事件
    WeatherLayer.ChangeWeather ""
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(ID As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    'Page.Res.NoLoadCrash = True
    Page.Res.NewImages App.Path & "\assets\maze"
    Page.Res.HotLoad = True
    'Page.Res.NoLoadCrash = False
    Page.Res.NewImages App.Path & "\assets\xlface"
    Page.Res.NewImages App.Path & "\assets\snowstormui", GW / 1519
    'Page.Res.NoLoadCrash = True
    Page.Res.NewImages App.Path & "\assets\battle\images"
    Page.LoadAnimationsFromDir App.Path & "\battle_animation"
    Page.CreatePlayAnimation "attackbg2", "attackbg", "Default"
    
    '创建页面
    ECore.Add Page, "MazePage"
    ReDim ATKS(278)
    Open App.Path & "\assets\snowstorm.atklist" For Binary As #1
    Get #1, , ATKS
    Close #1
    ATKIndex = 1
    ReDim UGun(0)
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
