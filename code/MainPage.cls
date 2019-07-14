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
    Dim Page As GPage
    Dim ATime1 As Long, AIndex1 As Long, AStep1 As Long
    Dim ATime2 As Long, AIndex2 As Long, AStep2 As Long
    Dim PartName As String
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear argb(255, 0, 0, 0) '清空画布

    Dim Hot As Long, Map() As Single
    If BGM.IsPlaying Then
        Map = BGM.GetMusicMap
        For i = 0 To UBound(Map)
            Hot = Hot + Map(i)
        Next
    End If
    
    Dim alpha As Single
    alpha = IIf(Hot / 50000 <= 0.7, Hot / 50000, 0.7)
    Page.DrawImage "bg.png", 0, 0, alpha:=1 - alpha
    'Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(alpha * 255, 255, 255, 255)
    
    Page.DrawImage "logo.png", 50, 50, alpha:=1
    Page.Writes "PART " & ESave.GetData("PART") & "  " & PartName, 85, 133, 18, argb(255, 168, 168, 168)
    Page.Writes "MORISTORY PREVIEW VERSION 710 POWERED BY EMERALD , MADE BY ERROR404 (QQ 1361778219) .", GW / 2, GH - 43, 16, argb(255, 168, 168, 168), align:=StringAlignmentCenter
    Page.Writes "游戏中所有音乐，音效和部分图片资源，均从网络免费正规渠道获取。", GW / 2, GH - 73, 16, argb(200, 255, 255, 255), align:=StringAlignmentCenter
    Dim m As Long
    
    Page.DrawImage "continue" & AIndex1 & ".png", 70, 200, alpha:=1
    m = CheckMouse2
    If m >= mMouseIn And AStep1 = 0 And AIndex1 = 0 Then
        AStep1 = 1: ATime1 = GetTickCount
    ElseIf m <= mMouseOut And AStep1 = 0 And AIndex1 = 8 Then
        AStep1 = -1: ATime1 = GetTickCount
    End If
    If m = mMouseUp Then
        NovelPage.ReadSave
        NovelPage.ReadNovel True
        NovelPage.RunNovel
        ECore.NewTransform transFadeIn, , "NovelPage"
    End If
    If GetTickCount - ATime1 >= 41 And AStep1 <> 0 Then
        AIndex1 = AIndex1 + AStep1
        If AIndex1 = 8 Then AStep1 = 0
        If AIndex1 = 0 Then AStep1 = 0
    End If
    
    Page.DrawImage "timeline" & AIndex2 & ".png", 70, 290, alpha:=1
    If CheckMouse2 >= mMouseIn And AStep2 = 0 And AIndex2 = 0 Then
        AStep2 = 1: ATime2 = GetTickCount
    ElseIf CheckMouse2 <= mMouseOut And AStep2 = 0 And AIndex2 = 10 Then
        AStep2 = -1: ATime2 = GetTickCount
    End If
    If GetTickCount - ATime2 >= 41 And AStep2 <> 0 Then
        AIndex2 = AIndex2 + AStep2
        If AIndex2 = 10 Then AStep2 = 0
        If AIndex2 = 0 Then AStep2 = 0
    End If
    
    If ESave.GetData("PART") = "" Then
        NovelPage.ReadSave
        NovelPage.ReadNovel
        ECore.NewTransform transFadeIn, , "NovelPage"
    End If
    
    If BGM.PlayState = musStopped And BGM.length <> 0 Then BGM.Play
    
    'Exit Sub
    'If App.LogMode <> 0 Then Exit Sub
    
    If Page.ShowColorButton(0, GW / 2 - 50 - 100, GH / 2 - 30, 100, 60, "Snowman", argb(255, 255, 255, 255), argb(255, 0, 176, 240), argb(255, 0, 196, 255), Size:=18) = 3 Then
        ECore.NewTransform transFadeIn, , "SnowmanPage"
    End If
    If Page.ShowColorButton(0, GW / 2 - 50 + 100, GH / 2 - 30, 100, 60, "TicTacToe", argb(255, 255, 255, 255), argb(255, 0, 176, 240), argb(255, 0, 196, 255), Size:=18) = 3 Then
        ECore.NewTransform transFadeIn, , "TicTacToePage"
    End If
    If Page.ShowColorButton(0, GW / 2 - 50 + 100, GH / 2 + 60, 100, 60, "Maze", argb(255, 255, 255, 255), argb(255, 0, 176, 240), argb(255, 0, 196, 255), Size:=18) = 3 Then
        MazePage.LoadMap "snow"
        ECore.NewTransform transFadeIn, , "MazePage"
    End If
    If Page.ShowColorButton(0, GW / 2 - 50 - 100, GH / 2 + 60, 100, 60, "BXBattle", argb(255, 255, 255, 255), argb(255, 0, 176, 240), argb(255, 0, 196, 255), Size:=18) = 3 Then
        ECore.NewTransform transFadeIn, , "BXBattlePage"
    End If
    If Page.ShowColorButton(0, GW / 2 - 50 + 100, GH / 2 + 120, 100, 60, "SNOWSTORM", argb(255, 255, 255, 255), argb(255, 0, 176, 240), argb(255, 0, 196, 255), Size:=18) = 3 Then
        MazePage.LoadMap "snowstorm"
        ECore.NewTransform transFadeIn, , "MazePage"
    End If
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
    Page.Res.NewImages App.Path & "\assets\ui", 0.5
    '创建页面
    ECore.Add Page, "MainPage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
