VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TicTacToePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim Chess(2, 2) As Long
    Dim Round As Long, RoundTime As Long, RoundText As Long
    Dim NeverChess As Boolean
    Dim XXWon As Boolean, MeWon As Boolean, WSX1 As Long, WSY1 As Long, WSX2 As Long, WSY2 As Long
    Dim XXWonCount As Long
'======================================
Public Function IsWon(Char As Long) As Boolean
    Dim Count As Long
    
    For y = 0 To 2
        Count = 0
        For X = 0 To 2
            If Chess(X, y) >= Char And Chess(X, y) <= Char + 2 Then Count = Count + 1
        Next
        If Count = 3 Then IsWon = True: WSX1 = 0: WSY1 = y: WSX2 = 2: WSY2 = y: Exit Function
    Next
    
    For X = 0 To 2
        Count = 0
        For y = 0 To 2
            If Chess(X, y) >= Char And Chess(X, y) <= Char + 2 Then Count = Count + 1
        Next
        If Count = 3 Then IsWon = True: WSX1 = X: WSY1 = 0: WSX2 = X: WSY2 = 2: Exit Function
    Next
    
    Count = 0
    For X = 0 To 2
        If Chess(X, X) >= Char And Chess(X, X) <= Char + 2 Then Count = Count + 1
    Next
    If Count = 3 Then IsWon = True: WSX1 = 0: WSY1 = 0: WSX2 = 2: WSY2 = 2: Exit Function
    
    Count = 0
    For X = 0 To 2
        If Chess(2 - X, X) >= Char And Chess(2 - X, X) <= Char + 2 Then Count = Count + 1
    Next
    If Count = 3 Then IsWon = True: WSX1 = 2: WSY1 = 0: WSX2 = 0: WSY2 = 2: Exit Function
End Function
Public Sub PlayChessXX()
    Dim Count As Long, AirX As Long, AirY As Long, AirCount As Long
    Dim Air1 As Long, Air2 As Long
    
    AirX = -1: AirY = -1
    
    For y = 0 To 2
        Count = 0: Air1 = -1
        For X = 0 To 2
            If Chess(X, y) >= 10 And Chess(X, y) <= 12 Then
                Count = Count + 1
            ElseIf Chess(X, y) = 0 And Air1 = -1 Then
                Air1 = X: Air2 = y
            End If
        Next
        If Air1 <> -1 And Count > AirCount Then AirX = Air1: AirY = Air2: AirCount = Count
    Next
    
    For X = 0 To 2
        Count = 0: Air1 = -1
        For y = 0 To 2
            If Chess(X, y) >= 10 And Chess(X, y) <= 12 Then
                Count = Count + 1
            ElseIf Chess(X, y) = 0 And Air1 = -1 Then
                Air1 = X: Air2 = y
            End If
        Next
        If Air1 <> -1 And Count > AirCount Then AirX = Air1: AirY = Air2: AirCount = Count
    Next
    
    Count = 0: Air1 = -1
    For X = 0 To 2
        If Chess(X, X) >= 10 And Chess(X, X) <= 12 Then
            Count = Count + 1
        ElseIf Chess(X, X) = 0 And Air1 = -1 Then
            Air1 = X: Air2 = X
        End If
    Next
    If Air1 <> -1 And (Count > AirCount Or NeverChess) Then AirX = Air1: AirY = Air2: AirCount = Count
    
    Count = 0: Air1 = -1
    For X = 0 To 2
        If Chess(2 - X, X) >= 10 And Chess(2 - X, X) <= 12 Then
            Count = Count + 1
        ElseIf Chess(2 - X, X) = 0 And Air1 = -1 Then
            Air1 = 2 - X: Air2 = X
        End If
    Next
    If Air1 <> -1 And (Count > AirCount Or NeverChess) Then AirX = Air1: AirY = Air2: AirCount = Count
    
    If AirX <> -1 And AirY <> -1 Then
        NeverChess = False
        SE.Play "Cursor.mp3"
        Chess(AirX, AirY) = 20 + Int(Rnd * 3): If Chess(AirX, AirY) > 22 Then Chess(AirX, AirY) = 22
        Exit Sub
    Else
        XXWonCount = XXWonCount + 1
        Round = 3
        RoundTime = GetTickCount
    End If
End Sub
Public Sub Update()
    If Err.Number <> 0 Then ErrCrash Err.Number, Err.Description
    On Error Resume Next

    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear '清空画布

    Page.DrawImage "desk.png", -70, -GH * 0.2, alpha:=1
    
    For X = 0 To 2
        For y = 0 To 2
            Page.Paint 0, 370 + X * 128, 180 + y * 128, 129, 129, argb(150, 64, 64, 64), style:=1
            If Chess(X, y) <> 0 Then
                If Chess(X, y) >= 20 Then
                    Page.DrawImage "bm" & (Chess(X, y) Mod 20) & ".png", 370 + X * 128, 180 + y * 128, alpha:=1
                ElseIf Chess(X, y) >= 10 Then
                    Page.DrawImage "circle" & (Chess(X, y) Mod 10) & ".png", 370 + X * 128, 180 + y * 128, alpha:=1
                End If
            End If
            If CheckMouse2 = mMouseUp And Chess(X, y) = 0 And Round = 0 Then
                SE.Play "Cursor.mp3"
                Randomize
                Chess(X, y) = 10 + Int(Rnd * 3)
                Round = 1: RoundTime = GetTickCount: RoundText = Int(Rnd * 5)
                ECore.NewTransform
                If Chess(X, y) > 12 Then Chess(X, y) = 12
            End If
        Next
    Next
    If XXWon Or MeWon Then
        Page.PaintLine 370 + WSX1 * 128 + 64, 180 + WSY1 * 128 + 64, _
                       370 + WSX2 * 128 + 64, 180 + WSY2 * 128 + 64, _
                       argb(255, 255, 0, 0), 10
    End If
    
    If Round = 1 Then
        Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(100, 0, 0, 0)
        Dim Text As String
        Select Case XXWonCount
            Case 0
                Select Case RoundText
                    Case 0: Text = "嗯。。。让人家想一想。"
                    Case 1: Text = "等等，青。。"
                    Case 2: Text = "唔。。。人家。。。"
                    Case 3: Text = "果然人家不擅长这个游戏呢。。"
                    Case 4: Text = "嗯。。。让人家想一想。"
                    Case Else
                End Select
            Case Is <= 3
                Select Case RoundText
                    Case 0: Text = "嗯。。。让人家想一想。"
                    Case 1: Text = "等等，青。。"
                    Case Else: Text = "嗯。。。让人家想一想。"
                End Select
            Case Is <= 5
                Select Case RoundText
                    Case 0: Text = "感觉我有点擅长这个游戏呢。"
                    Case 1: Text = "我可能要赢了哟。"
                    Case Else: Text = "加油，青。"
                End Select
            Case Is > 5
                Select Case RoundText
                    Case 0: Text = "青：我不想玩了。"
                    Case 1: Text = "青：好难玩。"
                    Case Else: Text = "青：兮兮让我赢吧！"
                End Select
        End Select
        Page.Writes Text, GW / 2, GH / 2 - 18, 36, argb(255, 255, 255, 255), align:=StringAlignmentCenter
        If GetTickCount - RoundTime >= 1400 Then
            Round = 0
            ECore.NewTransform
            Call PlayChessXX
        End If
    End If
    
    If Round = 3 Then
        Dim Pro As Single
        Pro = (GetTickCount - RoundTime) / 1000
        Pro = Cubic(Pro, 0, 1, 1, 1)
        If Pro >= 1 Then Pro = 1
        
        Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(Pro * 255 * IIf(XXWon Or MeWon, 0.3, 1), 0, 0, 0)
        If XXWon Or MeWon Then
            Page.Writes IIf(XXWon, "兮兮", "青") & "胜利了！", GW / 2, GH / 2 - 18, 36, argb(255, 255, 255, 255), align:=StringAlignmentCenter
        Else
            Page.Writes "嗯。。。死局了呢。", GW / 2, GH / 2 - 18, 36, argb(255, 255, 255, 255), align:=StringAlignmentCenter
        End If
        
        If GetTickCount - RoundTime >= 2000 Then
            If Not MeWon Then
                ECore.NewTransform transFallDark, 1000
                Round = 0
                Erase Chess
                NeverChess = True
                XXWon = False: MeWon = False
            Else
                NovelPage.RunNovel
                ECore.NewTransform transHighLight, 3000, "NovelPage"
            End If
        End If
    End If
    
    If Round <> 3 Then
        XXWon = IsWon(20)
        If XXWon Then XXWonCount = XXWonCount + 1
        MeWon = IsWon(10)
        If XXWon Or MeWon Then
            SE.Play "Applause1.mp3"
            Round = 3
            RoundTime = GetTickCount
            ECore.NewTransform transHighLight
        End If
    End If
    
    If BGM.PlayState = musStopped And BGM.length <> 0 Then BGM.Play
    If BGS.PlayState = musStopped And BGS.length <> 0 Then BGS.Play
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
    BGM.Create App.Path & "\music\bgm\OST 3 - Sweet Flower.mp3"
    BGM.Play
    XXWonCount = 0
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
    Page.Res.NewImages App.Path & "\assets\TicTacToe"
    Page.Res.newImage App.Path & "\assets\bg\desk.png", GW * 1.4, GH * 1.4
    '创建页面
    ECore.Add Page, "TicTacToePage"
    NeverChess = True
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
