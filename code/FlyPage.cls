VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlyPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim PX As Single, PY As Long
    Dim Jumptime As Long, JumpMode As Long
    Private Type BlockerInfo
        X As Long
        CY As Long
    End Type
    Dim Blockers() As BlockerInfo
    Dim Crash As New GCrashBox
    Dim Heart As Long, HurtTime As Long
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear '清空画布
    
    Dim CX As Long
    CX = PX Mod (GW * 2)
    Page.DrawImage "Sky.png", -CX, 0, alpha:=1, Direction:=DirNormal
    Page.DrawImage "Sky.png", GW - CX, 0, alpha:=1, Direction:=DirHorizontal
    Page.DrawImage "Sky.png", GW * 2 - CX, 0, alpha:=1, Direction:=DirNormal
    
    If GetTickCount - HurtTime <= 3000 Then
        If GetTickCount Mod 300 < 150 Then Page.DrawImage "yzhurt.png", 150, PY, alpha:=1, Pos:=posOnCenter
    Else
        Page.DrawImage "yz.png", 150, PY, alpha:=1, Pos:=posOnCenter
    End If
    
    Crash.Reset
    
    Dim DrawX As Long
    For i = 0 To UBound(Blockers)
        DrawX = Blockers(i).X - PX + 150
        If DrawX > -20 And DrawX <= GW Then
            Page.Paint 0, DrawX, 0, 20, Blockers(i).CY, argb(100 - 100 * Abs(DrawX - 20) / GW, 0, 0, 0)
            Crash.IsCrash
            Page.Paint 0, DrawX, Blockers(i).CY + (400 - i / 30 * 150), 20, GH - (400 - i / 30 * 150) - Blockers(i).CY, argb(100 - 100 * Abs(DrawX - 20) / GW, 0, 0, 0)
            Crash.IsCrash
        End If
    Next
    
    Dim IsHurt As Boolean
    
    If PY > GH Then Heart = 0
    If Crash.CrashResult Then IsHurt = True
    
    If IsHurt And GetTickCount - HurtTime > 3000 Then
        Heart = Heart - 1
        HurtTime = GetTickCount
        SE.Play "Damage2.mp3"
    End If
    
    If Heart <= 0 Then
        PX = 0: PY = 0
        Heart = 3
        SE.Play "Collapse4.mp3"
        ECore.NewTransform transDarkTo, 1000
    End If
    
    PX = PX + 1 + 2 * (PX / 8888)
    
    If (IsKeyPress(VK_SPACE) Or Mouse.button = 1) And JumpMode = 0 Then
        JumpMode = 1
        Jumptime = GetTickCount
    End If
    
    Dim pro As Single
    If JumpMode = 0 Then
        pro = (GetTickCount - Jumptime) / 500
    Else
        pro = (GetTickCount - Jumptime) / 300
    End If
    
    If pro > 1 Then pro = 1
    
    If JumpMode = 0 Then
        pro = Cubic(pro, 0, 0, 0, 1)
    Else
        pro = Cubic(pro, 0, 1, 1, 1)
    End If
    
    If pro = 1 And JumpMode = 1 Then JumpMode = 0: Jumptime = GetTickCount: SE.Play "Jump1.mp3"
    
    PY = PY - IIf(JumpMode = 0, -3 * pro - 1, 4 * pro)
    
    If PX > 15000 Then
        SE.Play "Applause1.mp3"
        NovelPage.RunNovel
        ECore.NewTransform transHighLight, 3000, "NovelPage"
    End If
    
    If PY < 0 Then PY = 0
    
    For i = 1 To 3
        Page.DrawImage IIf(Heart >= i, "", "un") & "heart.png", 50 + (i - 1) * 100, 50, alpha:=1
    Next
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
    BGM.Create App.Path & "\music\bgm\OST 17 - Good Day with Blues.mp3"
    BGM.Play
    BGS.Dispose
    Jumptime = GetTickCount
    JumpMode = 0
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
    ReDim Blockers(0)

    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.Res.NewImages App.Path & "\assets\flyyz"
    Page.Res.NoLoadDirection = False
    Page.Res.newImage App.Path & "\assets\fg\Sky.png", GW, GH
    '创建页面
    ECore.Add Page, "FlyPage"
    
    Randomize
    ReDim Blockers(30)
    For i = 0 To 30
        Blockers(i).X = 1000 + i * 400
        Blockers(i).CY = Int(Rnd * (GH - (400 - i / 30 * 150)))
    Next
    Heart = 3
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
