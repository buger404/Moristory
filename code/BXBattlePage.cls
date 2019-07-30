VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BXBattlePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Private Type BKey
        Pos As Single
        Area As Integer
        Pic As String
    End Type
    Dim Keys() As BKey, NowIndex As Long
    Dim Heart As Long
    Dim HurtTime As Long, LitTime As Long
    Private Type EXCEMark
        Time As Long
        X As Long
        y As Long
    End Type
    Dim Exce() As EXCEMark
    Dim BadMark As Boolean
    Dim EXCED As Boolean
'======================================

Public Sub Update()
    If Err.Number <> 0 Then ErrCrash Err.Number, Err.Description
    On Error Resume Next
    
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear '清空画布
    
    Dim Hot As Long, Map() As Single
    If BGM.IsPlaying Then
        Map = BGM.GetMusicMap
        For i = 0 To UBound(Map)
            Hot = Hot + Map(i)
        Next
    End If
    
    Page.DrawImage "garden.jpg", 0, 0, alpha:=1
    Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(IIf(Hot / 300000 * 20 <= 20, Hot / 300000 * 20, 20) + 50, 255, 255, 255)
    
    If BadMark Then
        Page.DrawImage "bm-sad.png", GW / 2 - 335, 70, alpha:=1
        Page.DrawImage "xx-sad.png", GW / 2 - 15, 150, alpha:=1
    Else
        Page.DrawImage "bm-normal.png", GW / 2 - 335, 70, alpha:=1
        Page.DrawImage "xx-normal.png", GW / 2 - 15, 150, alpha:=1
    End If
    
    Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(20, 0, 0, 0)
    Page.Paint 0, 0, GH * 0.75, GW + 1, GH * 0.25, argb(150, 0, 0, 0)
    
    Page.Writes Int(BGM.position / BGM.length * 100) & "%", GW / 2, 42, 32, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    
    Dim W As Long, Pro As Single
    W = Int(GW / 9)
    
    Page.Paint 0, Int(Mouse.X / W) * W, 0, W, GH + 1, IIf(GetTickCount - HurtTime <= 5000, IIf(GetTickCount Mod 300 < 150, argb(50, 255, 0, 0), argb(255, 255, 0, 0)), argb(50, 255, 255, 255))
    
    Dim Speed As Single
    Speed = 4.5 / (1 + Cubic((BGM.position / BGM.length), 0, 1, 1, 1) * 1.14)
    
    If GetTickCount - LitTime <= 200 Then
        Pro = 1 - (GetTickCount - LitTime) / 200
        Pro = Cubic(Pro, 0, 1, 1, 1)
        Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(128 * Pro, 255, 255, 255)
    End If
    
    Dim NeverDraw As Boolean
    For i = NowIndex To UBound(Keys)
        If Keys(i).Pos = 0 Then Exit For
        If Keys(i).Pos - BGM.position > Speed Then
            If i = NowIndex Then NeverDraw = True
            Exit For
        End If
        Pro = (1 - (Keys(i).Pos - BGM.position) / Speed) * (GH * 0.7 + 90 * 2) - 90 * 1.6
        Page.DrawImage Keys(i).Pic, W * Keys(i).Area + W / 2, Int(Pro), alpha:=IIf(i = NowIndex, 1, 0.6), Pos:=posOnCenter
        If Pro >= GH * 0.75 Then
            If Int(Mouse.X / W) = Keys(i).Area And i = NowIndex Then
                SE.Play "hit.wav"
                NowIndex = i + 1
                'Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(255, 255, 255, 255)
                'LitTime = GetTickCount
                If Abs(BGM.position - Keys(i).Pos) <= 0.2 Then
                    ReDim Preserve Exce(UBound(Exce) + 1)
                    With Exce(UBound(Exce))
                        .X = Mouse.X
                        .y = Mouse.y
                        .Time = GetTickCount
                    End With
                    EXCED = True
                End If
                
                If BGM.IsPlaying Then
                    If Not EXCED And NowIndex >= 16 And BadMark = False Then
                        BGM.rate = 0.5
                        BadMark = True
                    End If
                End If
                
            End If
        End If
        
        If Pro > GH Then
            NowIndex = i + 1
            If GetTickCount - HurtTime >= 5000 Then
                Heart = Heart - 1
                SE.Play "Damage2.mp3"
                Select Case Keys(i).Pic
                    Case "xx_atk.png"
                        Page.SwitchChannel "千叶斩", "Default"
                    Case "bm_atk.png"
                        Page.SwitchChannel "圣光矢", "Default"
                End Select
                HurtTime = GetTickCount
            End If
        End If
    Next
    
    If NeverDraw Then
        If GetTickCount Mod 300 < 150 Then Page.Writes "移动鼠标接住攻击！", GW / 2, GH / 2, 48, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
    End If
    
    For i = 1 To 3
        Page.DrawImage "heart.png", 20 + (i - 1) * 64, 20, alpha:=IIf(Heart >= i, 1, 0.3)
    Next
    If GetTickCount - HurtTime <= 5000 Then Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(IIf(GetTickCount Mod 300 < 150, 20, 0), 255, 0, 0)
    
    Dim ExceStr As String
    ExceStr = "EXCELLENT"
    
    For i = 1 To UBound(Exce)
        If i > UBound(Exce) Then Exit For
        Pro = Cubic((GetTickCount - Exce(i).Time) / 1000, 0, 1, 1, 1)
        If Pro > 1 Then
            Exce(i) = Exce(UBound(Exce))
            ReDim Preserve Exce(UBound(Exce) - 1)
            GoTo SkipThis
        End If
        For S = 1 To Len(ExceStr)
            Page.Writes Mid(ExceStr, S, 1), Exce(i).X + Pro * (S - 4) * 40, Exce(i).y - 9, 18, argb(255 - 255 * Pro, 255, 255, 255), align:=StringAlignmentCenter
        Next
SkipThis:
    Next

    Page.DrawAnimation "千叶斩", GW / 2, GH - 200
    Page.DrawAnimation "圣光矢", GW / 2, GH - 200
    
    If BGM.length = BGM.position Then
        SE.Play "Applause1.mp3"
        NovelPage.RunNovel
        ECore.NewTransform transHighLight, 3000, "NovelPage"
    End If
    
    If Heart <= 0 Then
        SE.Play "Collapse4.mp3"
        NowIndex = 1
        BGM.Stops
        BGM.Play
        HurtTime = 0
        Heart = 3
        ReDim Exce(0)
        ECore.NewTransform transFallDark, 2000
    End If
End Sub
Public Sub KeyDown(Code As Integer)
    Exit Sub
    Dim Area As Integer
    Select Case Code
        Case vbKeyA: Area = 0
        Case vbKeyS: Area = 1
        Case vbKeyD: Area = 2
        Case vbKeyF: Area = 3
        Case vbKeyG: Area = 4
        Case vbKeyH: Area = 5
        Case vbKeyJ: Area = 6
        Case vbKeyK: Area = 7
        Case vbKeyL: Area = 8
    End Select
    
    ReDim Preserve Keys(UBound(Keys) + 1)
    With Keys(UBound(Keys))
        .Pos = BGM.position
        .Area = Area
        .Pic = IIf(Int(Rnd * 2) = 1, "xx_atk.png", "bm_atk.png")
    End With
    
    ECore.NewTransform transHighLight, 50
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
    BGM.Create App.Path & "\music\bgm\OST 13 - Summer.mp3"
    BGM.Play
    BGM.rate = 1
    Heart = 3
    NowIndex = 1
    BadMark = False
End Sub
Public Sub Leave()
    '页面离开事件
    BGM.rate = 1
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
    Page.Res.NoLoadCrash = True
    Page.Res.newImage App.Path & "\assets\face\bm-normal.png"
    Page.Res.newImage App.Path & "\assets\face\xx-normal.png"
    Page.Res.newImage App.Path & "\assets\face\bm-sad.png"
    Page.Res.newImage App.Path & "\assets\face\xx-sad.png"
    Page.Res.newImage App.Path & "\assets\bg\garden.jpg", GW, GH
    Page.Res.NewImages App.Path & "\assets\BXBattle"
    Page.Res.HotLoad = True
    Page.Res.NewImages App.Path & "\assets\battle\images", 1.8
    Page.LoadAnimationsFromDir App.Path & "\battle_animation"
    
    Page.CreatePlayAnimation "rpgmv_Wind2", "千叶斩", "Reserve"
    Page.CreatePlayAnimation "rpgmv_Holy2", "圣光矢", "Reserve"
    
    '创建页面
    ECore.Add Page, "BXBattlePage"
    ReDim Keys(330)
    ReDim Exce(0)
    Open App.Path & "\assets\bxbattlekeys.key" For Binary As #1
    Get #1, , Keys
    Close #1
    NowIndex = 1
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
