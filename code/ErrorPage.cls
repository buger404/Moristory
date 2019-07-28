VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ErrorPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Public ErrText As String
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear argb(255, 51, 133, 255) '清空画布

    For i = 1 To 10
        Page.DrawImage "bm.png", GW - 200 + Cos(GetTickCount / 200 + i / 10 * 5) * 50, 600 - Sin(GetTickCount / 200 + i / 10 * 5) * 50, alpha:=i / 10, Pos:=posOnCenter
    Next
    
    Dim AText As String
    For i = 1 To 2000 / 200 - 1
        If GetTickCount Mod 2000 > i * 200 Then AText = AText & "︵"
    Next
    
    If GetTickCount Mod 4000 <= 2000 Then
        Page.Writes "(╯‵□′)╯" & AText & "┻━┻", 100, 100, 36, argb(255, 255, 255, 255), style:=FontStyleBold
    Else
        Page.Writes "┬—┬" & AText & "ノ('-'ノ)", 100, 100, 36, argb(255, 255, 255, 255), style:=FontStyleBold
    End If
    
    Page.Writes ErrText, 100, 240, 18, argb(255, 255, 255, 255)

    Page.Writes "请尝试通过以下方式反馈问题：" & vbCrLf & vbCrLf & "① 联系QQ1361778219" & vbCrLf & "② 发送邮件至ris_vb@126.com" & vbCrLf & "③ 在Github页面上提交Issue" & vbCrLf & "④ 顺着网线爬到404家", 100, 370, 18, argb(255, 255, 255, 255)
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
    Page.Res.NewImages App.Path & "\assets\face", 0.7
    '创建页面
    ECore.Add Page, "ErrorPage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
