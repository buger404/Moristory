VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SnowmanPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim Balls() As Single
    Dim StartTime As Long, BallTime As Long, DrawY As Long
    Dim DeathText As String
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear '清空画布
    
    Dim High As Long, y As Long
    For i = 1 To UBound(Balls)
        High = High + Balls(i)
    Next
    y = High - GH / 2
    If y < 0 Then y = 0
    If y > 4040 Then y = 4040
    DrawY = DrawY + (y - DrawY) / 10
    
    If High >= 4040 + GH / 2 + 50 Then
        SE.Play "Applause1.mp3"
        NovelPage.RunNovel
        ECore.NewTransform transHighLight, 3000, "NovelPage"
    End If
    
    Page.DrawImage "Snowfield.png", 0, DrawY, alpha:=1
    Page.DrawImage "Sky.png", 0, DrawY - 4040, alpha:=1
    
    Page.Writes format(High / 100 * (1 + Cubic(y / 4040, 0, 0, 0, 1) * 10), "0.00") & " m", GW / 2, 20, 24, argb(200, 27, 27, 27), align:=StringAlignmentCenter
    Dim LeftTime As Long
    LeftTime = 40 - Int((GetTickCount - StartTime) / 1000)
    If LeftTime < 0 Then LeftTime = 0
    
    Page.Writes LeftTime & " s", GW / 2, 60, 24, argb(IIf((LeftTime <= 10) And (GetTickCount Mod 300 < 150), 0, 200), IIf((LeftTime <= 10), 255, 27), 27, 27), align:=StringAlignmentCenter
    
    If LeftTime = 0 Then
        SE.Play "Collapse4.mp3"
        StartTime = GetTickCount + 2000
        BGM.Stops
        BGM.Play
        ReDim Balls(1)
        Balls(1) = 400
        DeathText = "超时！"
        'ECore.NewTransform transFallDark, 2000
    End If
    
    Dim High2 As Long
    For i = 1 To UBound(Balls)
        If GH - 40 + DrawY - High2 <= GH + GW Then
            Page.Paint 1, GW / 2 - Int(Balls(i)) / 2, GH - 40 + DrawY - High2, Int(Balls(i)), Int(Balls(i)), argb(255, 242, 242, 255), Pos:=posOnBottom
            Page.Paint 1, GW / 2 - Int(Balls(i)) / 2, GH - 40 + DrawY - High2, Int(Balls(i)), Int(Balls(i)), argb(255, 212, 212, 225), Size:=2, style:=1, Pos:=posOnBottom
        End If
        High2 = High2 + Balls(i)
    Next
    
    Dim m As Integer
    m = CheckMouse(0, 0, GW, GH)
    If m = MButtonState.mMouseDown And BallTime = 0 And GetTickCount - StartTime <= 60000 Then
        BallTime = GetTickCount
        Page.SwitchChannel "Ice", "Default"
    End If
    
    If BallTime <> 0 Then
        Dim Size As Single
        Size = (GetTickCount - BallTime) / (700 / ((1 + Cubic(y / 4040, 0, 1, 1, 1) * 0.6))) * 100
        If Size > GW Then Size = GW
        Page.Paint 1, GW / 2 - Int(Size) / 2, GH - 40 + DrawY - High2, Int(Size), Int(Size), argb(120, 0, 0, 0), Pos:=posOnBottom
        Page.DrawAnimation "Ice", GW / 2, GH - 40 + DrawY - High2 - Int(Size) / 2
        Page.Writes Int(Int(Size) / Balls(UBound(Balls)) * 100) & "%", GW / 2, GH - 40 + DrawY - High2 - Int(Size) / 2 - 12, 24, argb(255, 27, 27, 27), align:=StringAlignmentCenter
        'Page.Paint 1, GW / 2 - Int(Size) / 2, GH - 40 + DrawY - High2, Int(Size), Int(Size), argb(255, 212, 212, 225), Size:=2, style:=1, Pos:=posOnBottom
        If m = MButtonState.mMouseUp Then
            BallTime = 0
            ReDim Preserve Balls(UBound(Balls) + 1)
            Balls(UBound(Balls)) = Size
            SE.Play "Ice4.mp3"
            If Size > Balls(UBound(Balls) - 1) * 1.1 Then
                SE.Play "Collapse4.mp3"
                StartTime = GetTickCount + 2000
                BGM.Stops
                BGM.Play
                ReDim Balls(1)
                Balls(1) = 250
                DeathText = "释放的雪球太大，雪人倒塌了！"
                'ECore.NewTransform transFallDark, 2000
            End If
        End If
    End If
    
    If UBound(Balls) <= 1 Then
        If GetTickCount Mod 300 < 150 Then Page.Writes DeathText, GW / 2, GH / 2 - 18, 36, argb(150, 0, 0, 0), align:=StringAlignmentCenter
    End If
    
    If BGM.PlayState = musStopped And BGM.length <> 0 Then BGM.Play
    If BGS.PlayState = musStopped And BGS.length <> 0 Then BGS.Play
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
    BGM.Create App.Path & "\music\bgm\OST 14 - Snowy.mp3"
    BGM.Play
    StartTime = GetTickCount
    DeathText = "按住鼠标然后释放雪球！"
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
    Page.Res.HotLoad = True
    Page.Res.NoLoadCrash = True
    Page.Res.newImage App.Path & "\assets\bg\Snowfield.png", GW, GH
    Page.Res.newImage App.Path & "\assets\fg\Sky.png", GW, GH
    Page.Res.NewImages App.Path & "\assets\battle\images", 1.8
    Page.LoadAnimation App.Path & "\battle_animation\rpgmv_Ice1.emrani"
    Page.CreatePlayAnimation "rpgmv_Ice1", "Ice", "Default"
    '创建页面
    ECore.Add Page, "SnowmanPage"
    ReDim Balls(1)
    Balls(1) = 250
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
