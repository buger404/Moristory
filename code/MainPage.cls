VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Private Declare Sub AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal hHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BLENDFUNCTION As Long) ' As Long

    Dim Page As GPage
    Dim ATime1 As Long, AIndex1 As Long, AStep1 As Long
    Dim ATime2 As Long, AIndex2 As Long, AStep2 As Long
    Dim PartName As String
    Dim StarTime As Long, StarDC As Long, StarGG As Long
    Dim LoadTime As Long, NoTime As Long
'======================================

Public Sub Update()
    If Err.number <> 0 Then ErrCrash Err.number, Err.Description
    On Error Resume Next

    '游戏每一帧的过程（绘制过程请放在此处）
    
    'MsgBox 1 / 0
    
    If GetTickCount - StarTime <= 25 Then Exit Sub
    StarTime = GetTickCount
    
    Dim ODC As Long, OGG As Long
    ODC = Page.CDC: OGG = Page.GG
    Page.CDC = StarDC: Page.GG = StarGG
    
    Dim Hot As Long, map() As Single
    If BGM.IsPlaying Then
        map = BGM.GetMusicMap
        For i = 0 To UBound(map)
            Hot = Hot + map(i)
        Next
    End If
    
    If BGM.PlayState = musStopped And BGM.length <> 0 Then BGM.Play
    
    Dim bl As BLENDFUNCTION, blong As Long
    With bl
        .AlphaFormat = &H1
        .BlendFlags = &H0
        .BlendOp = 0
        .SourceConstantAlpha = 255
    End With
    CopyMemory blong, bl, 4
    
    StretchBlt Page.CDC, 15, 0, GW - 30, GH - 20, Page.CDC, 0, 0, GW, GH, SRCCOPY

    Dim brush As Long, bRect As RECT
    
    brush = CreateSolidBrush(RGB(6, 31, 62))

    Dim Rgn As Long, Rgn2 As Long
    Rgn = CreateRectRgn(0, 0, GW, GH)
    Rgn2 = CreateRectRgn(15, 0, GW - 30, GH - 20)
    CombineRgn Rgn, Rgn, Rgn2, RGN_DIFF
    SelectClipRgn Page.CDC, Rgn
    
    With bRect
        .Left = 0
        .top = 0
        .Right = GW
        .Bottom = GH
    End With
    FillRect Page.CDC, bRect, brush
    
    DeleteObject brush
    DeleteObject Rgn
    DeleteObject Rgn2
    
    Dim Img As Long, W As Long, h As Long, index As Integer
    Dim X As Long, y As Long
    Randomize
    For i = 1 To 5
        index = Page.Res.GetImage("star" & Int(Rnd * 5) & ".png")
        Img = Page.Res.ResourceImageHandle(index)
        W = Page.Res.ImgSize(index, imgGetWidth)
        h = Page.Res.ImgSize(index, imgGetHeight)
        
        Select Case Int(Rnd * 2)
            Case 0: X = Int(Rnd * W) - W / 2: y = Int(Rnd * (GH + 400) - 200)
            Case Else: X = GW - (Int(Rnd * W) - W / 2): y = Int(Rnd * (GH + 400) - 200)
        End Select
        Page.DrawImage index, X, y, alpha:=1
    Next
    Dim hi As Single
    hi = Hot / 30000
    If hi > 1 Then hi = 1
    
    Page.Paint 0, GW / 2 - hi * GW / 2, GH - 10, hi * GW, 10, argb(hi * 127 + 127, 255, 255, 255)
    
    SelectClipRgn Page.CDC, 0
    Page.CDC = ODC: Page.GG = OGG
    
    BitBlt Page.CDC, 0, 0, GW, GH, StarDC, 0, 0, vbSrcCopy

    Page.DrawImage "bg.png", 0, 0, alpha:=0.5 + hi * 0.1
    Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(108, 0, 0, 0)

    Dim lPro As Single
    lPro = (GetTickCount Mod 3000) / 1500
    If lPro <= 1 Then
        lPro = Cubic(lPro, 0, 0, 0, 1)
    Else
        lPro = Cubic(2 - lPro, 0, 1, 1, 1)
    End If
    
    Page.DrawImage "logo.png", GW / 2, 300, alpha:=1, Pos:=posOnCenter
    Page.DrawImage "logoblur.png", GW / 2, 300, alpha:=lPro, Pos:=posOnCenter
    
    If GetTickCount Mod 1000 < 500 Then Page.Writes "- 点击开始 -", GW / 2, GH / 2 + 110, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
    Page.Writes "Chapter " & ESave.GetData("PART") & "  " & PartName, GW / 2, GH / 2 - 10, 18, argb(200, 255, 255, 255), align:=StringAlignmentCenter
    
    If CheckMouse(0, 0, GW, GH) = mMouseDown And LoadTime = 0 And GetTickCount - NoTime >= 1000 Then
        Mouse.state = 0
        SE.Play "Load.mp3"
        LoadTime = GetTickCount
    End If

    Dim bPro As Single
    If LoadTime <> 0 Then
        bPro = Cubic((GetTickCount - LoadTime) / 2000, 0, 1, 1, 1)
        Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(255 * bPro, 0, 0, 0)
    End If
    
    If GetTickCount - LoadTime >= 2000 And LoadTime <> 0 Then
        LoadTime = 0
        ECore.NewTransform transFallDark, 1000, "NovelPage"
        NovelPage.ReadSave
        NovelPage.ReadNovel True
        NovelPage.RunNovel
    End If
    
    If ESave.GetData("PART") = "" Then
        NovelPage.ReadSave
        NovelPage.ReadNovel
        ECore.NewTransform transFadeIn, , "NovelPage"
    End If
    
    Page.Writes "MORISTORY POWERED BY EMERALD , MADE BY ERROR404 (QQ 1361778219) .", GW / 2, GH - 43, 16, argb(105, 255, 255, 255), align:=StringAlignmentCenter
    Page.Writes "游戏中所有音乐，音效和部分图片资源，均从网络免费正规渠道获取。", GW / 2, GH - 73, 16, argb(105, 255, 255, 255), align:=StringAlignmentCenter

    Page.Writes "小说版地址：https://shequ.codemao.cn/wiki/novel/cover/47294", GW / 2, GH - 133, 16, argb(105, 255, 255, 255), align:=StringAlignmentCenter
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
    BGM.Create App.Path & "\music\bgm\OST 26 - Reminiscence.mp3"
    BGM.Play
    LoadTime = 0
    NoTime = GetTickCount
    Mouse.state = 0: Mouse.button = 0
    If ESave.GetData("PART") = "" Then Exit Sub
    PartName = GetPartTitle(ESave.GetData("PART"))
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
    Page.Res.NewImages App.Path & "\assets\stars", 0.8
    Page.Res.NewImages App.Path & "\assets\ui", 0.5
    Page.Res.newImage App.Path & "\assets\stars\logo.png", 0.8, Name:="logoblur.png"
    Page.Res.ApplyBlurEffect "logoblur.png", 30, 0

    StarDC = CreateCDC(GW, GH)
    PoolCreateFromHdc StarDC, StarGG
    GdipGraphicsClear StarGG, argb(255, 6, 31, 62)
    
    '创建页面
    ECore.Add Page, "MainPage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
