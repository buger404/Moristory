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
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear '清空画布

    If Page.ShowColorButton(0, GW / 2 - 50 - 100, GH / 2 - 30, 100, 60, IIf(ESave.GetData("HUsh") = "冰棍", "END", "START"), argb(255, 255, 255, 255), argb(255, 0, 176, 240), argb(255, 0, 196, 255), size:=18) = 3 Then
        NovelPage.ReadNovel
        ECore.NewTransform transFadeIn, , "NovelPage"
    End If
    If Page.ShowColorButton(0, GW / 2 - 50 + 100, GH / 2 - 30, 100, 60, "Battle", argb(255, 255, 255, 255), argb(255, 0, 176, 240), argb(255, 0, 196, 255), size:=18) = 3 Then
        ECore.NewTransform transFadeIn, , "BattlePage"
        MazePage.LoadMap
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
    Page.Res.NewImages App.Path & "\assets"
    '创建页面
    ECore.Add Page, "MainPage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
