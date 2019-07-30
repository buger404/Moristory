VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadingPage"
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

    Page.DrawImage "logo.png", GW / 2, GH / 2 - 50, alpha:=1, Pos:=posOnCenter
    
    Dim AText As String
    For i = 1 To 2000 / 200 - 1
        If GetTickCount Mod 2000 > i * 200 Then AText = AText & "︵"
    Next
    
    If GetTickCount Mod 4000 <= 2000 Then
        AText = "(╯‵□′)╯" & AText & "┻━┻"
    Else
        AText = "┬—┬" & AText & "ノ('-'ノ)"
    End If
    
    Page.Writes AText & vbCrLf & vbCrLf & _
                "正在加载第 " & LoadedCount & " 个资源 (" & Int(CrashPro * 100) & "%)", _
                GW / 2, GH / 2 + 50, 18, argb(255, 64, 64, 64), align:=StringAlignmentCenter
                
    Page.Writes Replace(GetCountStr, vbCrLf, " , "), _
                GW / 2, GH - 50, 18, argb(255, 128, 128, 128), align:=StringAlignmentCenter
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
    Page.Res.newImage App.Path & "\assets\logo.png", 0.7
    '创建页面
    ECore.Add Page, "LoadingPage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
