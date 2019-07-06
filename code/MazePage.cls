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
'======================================
Public Sub LoadMap()
    Dim w As Long, h As Long, Image As Long, Color As Long

    Page.Res.newImage App.path & "\assets\ui\maze.png"
    Page.Res.newImage App.path & "\assets\ui\shadow.png", GW, GH
    Image = Page.Res.ResourceImageHandle(Page.Res.GetImage("maze.png"))
    GdipGetImageWidth Image, w: GdipGetImageHeight Image, h
    
    MapDC = CreateCDC(w * 90, h * 90 + 90)
    CoverDC = CreateCDC(w * 90, h * 90 + 90)
    
    ReDim Map(w + 1, h + 2)
    For Y = 0 To h
        For X = 0 To w
            GdipBitmapGetPixel Image, X, Y, Color
            If Color = 0 Then Map(X, Y + 1) = Page.Res.GetImage("alpha.png")
            If Color = argb(255, 255, 255, 255) Then Map(X, Y + 1) = Page.Res.GetImage("white.png")
            If Color = argb(255, 153, 204, 255) Then Map(X, Y + 1) = Page.Res.GetImage("blue.png")
            If Color = argb(255, 255, 153, 51) Then Map(X, Y + 1) = Page.Res.GetImage("orange.png")
            If Color = argb(255, 0, 255, 0) Then Map(X, Y + 1) = Page.Res.GetImage("green.png")
            If Color = argb(255, 0, 0, 0) Then Map(X, Y + 1) = Page.Res.GetImage("black.png")
            If Color = argb(255, 0, 255, 255) Then Map(X, Y + 1) = Page.Res.GetImage("alpha.png"): StartX = X: StartY = Y + 1
            If Color = argb(255, 255, 0, 0) Then Map(X, Y + 1) = Page.Res.GetImage("alpha.png"): EndX = X: EndY = Y + 1
        Next
    Next
    PX = StartX: PY = StartY
    
    Dim ODC As Long
    ODC = Page.CDC
    Page.CDC = MapDC
    For X = 0 To w
        For Y = 1 To h + 1
            Page.DrawImage Map(2, 2), (X - 0) * 90, (Y - 0) * 90, Alpha:=1
            If Page.Res.ImgSize(Map(X, Y), imgGetHeight) = 180 Then
                Page.DrawImage Map(X, Y), (X - 0) * 90, (Y - 0) * 90, cy:=90, ch:=90, Alpha:=1
                Page.CDC = CoverDC
                Page.DrawImage Map(X, Y), (X - 0) * 90, (Y - 0) * 90 - 90, cy:=0, ch:=90, Alpha:=1
                Page.CDC = MapDC
            Else
                Page.DrawImage Map(X, Y), (X - 0) * 90, (Y - 0) * 90, Alpha:=1
            End If
        Next
    Next
    Page.CDC = ODC
    
End Sub
Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear argb(255, 123, 199, 47) '清空画布

    Dim w As Long, h As Long, cx As Long, cy As Long
    cx = (PX - 0) * 90 - GW / 2 + 45
    cy = (PY - 0) * 90 - GH / 2 + 45
    
    w = 90 * (UBound(Map, 1) - 1) - IIf(cx >= 0, cx, 0)
    h = 90 * (UBound(Map, 2) - 1) - IIf(cy >= 0, cy, 0)
    
    If w > GW Then w = GW
    If h > GH Then h = GH
    
    Dim pro As Single
    pro = 1 - (Abs(EndX - PX) + Abs(EndY - PY)) / 5
    If pro < 0 Then pro = 0
    If pro > 1 Then pro = 1
    
    PaintDC MapDC, Page.CDC, IIf(cx < 0, -cx, 0), IIf(cy < 0, -cy, 0), IIf(cx >= 0, cx, 0), IIf(cy >= 0, cy, 0), w, h, 1
    Page.DrawImage "darkssr.png", (DarkX - 0) * 90 - cx, (DarkY - 0) * 90 - cy, Alpha:=1 - pro * 4.6
    Page.Paint 1, GW / 2 - 45, GH / 2 - 45, 90, 90, argb(255, 255, 255, 255)
    PaintDC CoverDC, Page.CDC, IIf(cx < 0, -cx, 0), IIf(cy < 0, -cy, 0), IIf(cx >= 0, cx, 0), IIf(cy >= 0, cy, 0), w, h, 1
    
    Page.DrawImage "shadow.png", 0, 0, Alpha:=1 - pro * 1
    Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(pro * 255, 255, 255, 255)
    
    If pro = 1 Then
        BGM.rate = 1
        BGS.Dispose
        NovelPage.RunNovel
        SE.Play "Defense_3D.mp3"
        ECore.NewTransform transHighLight, 3000, "NovelPage"
    End If
    
    Dim HuaIndex As Integer
    HuaIndex = Page.Res.GetImage("black.png")
    
    If XRemain <> 0 Then
        If Abs(XRemain) < 0.1 Then
            For i = 0 To 4
                FX(i) = FX(i + 1): FY(i) = FY(i + 1)
            Next
            XRemain = 0: PX = Round(PX)
        Else
            PX = PX + IIf(XRemain < 0, -0.1, 0.1)
            XRemain = XRemain - IIf(XRemain < 0, -0.1, 0.1)
            If FollowStep <= 50 Then FollowStep = FollowStep + 1
            If FollowStep >= 50 Then FX(5) = PX: FY(5) = PY
        End If
    End If
    If YRemain <> 0 Then
        If Abs(YRemain) < 0.1 Then
            For i = 0 To 4
                FX(i) = FX(i + 1): FY(i) = FY(i + 1)
            Next
            YRemain = 0: PY = Round(PY)
        Else
            PY = PY + IIf(YRemain < 0, -0.1, 0.1)
            YRemain = YRemain - IIf(YRemain < 0, -0.1, 0.1)
            If FollowStep <= 50 Then FollowStep = FollowStep + 1
            If FollowStep >= 50 Then FX(5) = PX: FY(5) = PY
        End If
    End If
    
    DarkX = IIf(Abs(PX - FX(0)) < 2 And pro = 0, 0, FX(0)): DarkY = IIf(Abs(PY - FY(0)) < 2 And pro = 0, 0, FY(0))
    If pro > 0 Then
        DarkX = PX - 5
        DarkY = PY
    End If
    
    If IsKeyPress(VK_UP) And Map(Round(PX), Round(PY - 1)) <> HuaIndex And XRemain = 0 And YRemain = 0 Then YRemain = -1
    If IsKeyPress(VK_DOWN) And Map(Round(PX), Round(PY + 1)) <> HuaIndex And XRemain = 0 And YRemain = 0 Then YRemain = 1
    If IsKeyPress(VK_LEFT) And Map(Round(PX - 1), Round(PY)) <> HuaIndex And XRemain = 0 And YRemain = 0 Then XRemain = -1
    If IsKeyPress(VK_RIGHT) And Map(Round(PX + 1), Round(PY)) <> HuaIndex And XRemain = 0 And YRemain = 0 Then XRemain = 1
    
    If BGM.PlayState = musStopped And BGM.length <> 0 Then BGM.Play
    If BGS.PlayState = musStopped And BGS.length <> 0 Then BGS.Play
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(ID As String, Msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
    BGM.Create App.path & "\music\bgm\OST 4 - Garden Party.mp3"
    BGM.rate = 0.8
    BGM.Play
    BGS.Create App.path & "\music\bgs\Nightforest.mp3"
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
Public Sub AnimationDone(ID As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.Res.NewImages App.path & "\assets\maze"
    '创建页面
    ECore.Add Page, "MazePage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
