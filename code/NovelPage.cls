VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NovelPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim BG As String, FG As String, BGCOLOR As Long
    Dim CuLine As Long, NLines() As String
    Dim Mode As String, PART As String
    Dim LLineTime As Long
    Dim DisplayText As String
    Dim Title As String
    Private Type Role
        name As String
        Love As Long
        Face As String
    End Type
    Private Type Answer
        Caption As String
        LoveIncrase As Long
        JTimeLine As Long
    End Type
    Private Type Question
        Title As String
        Answers() As Answer
    End Type
    Dim Roles() As Role
    Dim Speaker As String
    Dim SpeakerName As String
    Dim TimeLine As Long
    Dim QuesTime As Long, Ques As Question, BlurDC As Long, ODC As Long
    Dim MsgTime As Long, AddingText As String, MsgTime2 As Long, AddPos As Long
    Dim MsgEvent As String
'======================================
Public Function GetSpeakerName(Speaker As String) As String
    GetSpeakerName = "？？？"
    Select Case Speaker
        Case "aside": GetSpeakerName = "旁白"
        Case "bm": GetSpeakerName = "黑嘴"
        Case "me": GetSpeakerName = "你"
        Case "xx": GetSpeakerName = "兮兮"
        Case "fj": GetSpeakerName = "浮橘"
        Case "s": GetSpeakerName = "世"
        Case "kx1": GetSpeakerName = "卡西"
        Case "kx2": GetSpeakerName = "卡茜"
        Case "km1": GetSpeakerName = "枯梦"
        Case "km2": GetSpeakerName = "枯梦"
        Case "jy": GetSpeakerName = "久悠"
        Case "xl": GetSpeakerName = "雪狼"
        Case "ssr": GetSpeakerName = "莎瑟瑞"
        Case "tk": GetSpeakerName = "塔克"
        Case "dn": GetSpeakerName = "迪娜"
        Case "yz": GetSpeakerName = "芽子"
        Case "bg": GetSpeakerName = "冰棍"
        Case "yy1": GetSpeakerName = "雅月"
        Case "yy2": GetSpeakerName = "雅月"
        Case Else
            If InStr(Speaker, "dark") = 1 Then SpeakerName = "？？？"
    End Select
End Function
Public Function GetStrLine() As String
    GetStrLine = UBound(Split(DisplayText, vbCrLf))
End Function
Public Sub RunNovel()
Exec:
    If CuLine > UBound(NLines) Then Exit Sub

    Dim IsCmd As Boolean, Params() As String
    Dim JumpTimeline As Boolean, AnswerStep As Long
    IsCmd = (InStr(NLines(CuLine), "\") <> 0)
    If IsCmd Then
        Params = Split(NLines(CuLine), "\")
        '获取答案们
        If AnswerStep > 0 Then
            With Ques.Answers(UBound(Ques.Answers) - AnswerStep)
                .Caption = Params(0)
                .JTimeLine = Val(Params(1))
                .LoveIncrase = Val(Params(2))
            End With
            AnswerStep = AnswerStep - 1
            CuLine = CuLine + 1
            QuesTime = GetTickCount
            If AnswerStep = 0 Then Exit Sub
            GoTo Exec
        End If
        '这不是你的时间线的剧情。。。
        If JumpTimeline Then
            If UBound(Params) = 2 Then
                If Params(0) = "timeline" And Params(2) = "end" Then
                    JumpTimeline = False
                    CuLine = CuLine + 1: GoTo Exec
                Else
                    CuLine = CuLine + 1: GoTo Exec
                End If
            Else
                CuLine = CuLine + 1: GoTo Exec
            End If
        End If
        
        Select Case Params(0)
            Case "go"
                Select Case Params(1)
                    Case "maze"
                        CuLine = CuLine + 1
                        ECore.NewTransform transFallDark, 3000, "MazePage"
                        Exit Sub
                End Select
            Case "***death"
                SE.Play "Collapse4.mp3"
                ECore.NewTransform transFallDark, 2000, "MainPage"
                Debug.Print Now, "你在 #" & TimeLine & " 时间线，PART " & PART & "的第" & CuLine & "行死了。"
                Exit Sub
            Case "question"
                AnswerStep = Val(Params(2))
                Ques.Title = Params(1)
                ReDim Ques.Answers(AnswerStep)
                PaintDC Page.CDC, ODC, 0, 0, Alpha:=1
                BlurTo BlurDC, GDC, GameWindow
                Mode = "question"
            Case "timeline"
                If Params(2) = "head" Then
                    If TimeLine <> Val(Params(1)) Then JumpTimeline = True: GoTo Exec
                ElseIf Params(2) = "end" And JumpTimeline = True Then
                    JumpTimeline = False: GoTo Exec
                End If
            Case "face"
                For i = 1 To UBound(Roles)
                    If Roles(i).name = Params(1) Then
                        Roles(i).Face = Params(2): Exit For
                    End If
                Next
                ECore.NewTransform
            Case "role"
                If Params(2) = "add" Then
                    ReDim Preserve Roles(UBound(Roles) + 1)
                    With Roles(UBound(Roles))
                        .name = Params(1)
                        .Face = "normal"
                        .Love = Val(ESave.GetData(.name & "'s love"))
                    End With
                ElseIf Params(2) = "remove" Then
                    For i = 1 To UBound(Roles)
                        If Roles(i).name = Params(1) Then
                            Roles(i) = Roles(UBound(Roles)): Exit For
                        End If
                    Next
                    ReDim Preserve Roles(UBound(Roles) - 1)
                End If
                ECore.NewTransform
            Case "say"
                DisplayText = ""
                Speaker = Params(1)
                SpeakerName = GetSpeakerName(Speaker)
            Case "clear"
                DisplayText = ""
                ECore.NewTransform
            Case "*bgm"
                BGM.Create App.path & "\music\bgm\" & Params(1)
                BGM.Play
            Case "*bgs"
                BGS.Create App.path & "\music\bgs\" & Params(1)
                BGS.Play
            Case "*play"
                SE.Play Params(1)
            Case "*title"
                Title = Params(1)
                GameWindow.Caption = "Moristory DEMO - PART " & PART & " , " & Title
            Case "*bg"
                BGCOLOR = argb(255, 0, 0, 0)
                BG = Params(1)
                ECore.NewTransform
            Case "*fg"
                BGCOLOR = argb(255, 0, 0, 0)
                FG = Params(1)
                ECore.NewTransform
            Case "*mode"
                Mode = Params(1)
            Case "effect"
                Select Case Params(1)
                    Case "light"
                        ECore.NewTransform transHighLight, 600
                    Case "dark"
                        ECore.NewTransform transFallDark, 600
                End Select
            Case "bg"
                Dim temp() As String
                temp = Split(Params(1), ",")
                BGCOLOR = argb(255, temp(0), temp(1), temp(2))
                BG = "": FG = ""
                ECore.NewTransform
            Case "window"
                If Params(1) = "Moristory" Then
                    GameWindow.Caption = "Moristory DEMO - PART " & PART & " , " & Title
                Else
                    GameWindow.Caption = Params(1)
                End If
            Case "*turn"
                ESave.PutData "PART", Params(1)
                ECore.NewTransform transFallDark, 1000
                Call ReadNovel
                Exit Sub
            Case "*event"
                MsgEvent = Params(1)
                If MsgEvent Like "ani+*" Then
                    Page.SwitchChannel MsgEvent, "Default"
                End If
        End Select
    ElseIf Left(NLines(CuLine), 1) <> "#" Then
        If JumpTimeline Then CuLine = CuLine + 1: GoTo Exec
        If NLines(CuLine) = "++" Then
            DisplayText = ""
            If UBound(Roles) = 0 Then
                Speaker = IIf(Speaker = "me", "", "me")
            Else
                Speaker = IIf(Speaker = "me", Roles(1).name, "me")
            End If
            SpeakerName = GetSpeakerName(Speaker)
        Else
            If Mode <> "scroll" Then
                AddPos = 1
                AddingText = NLines(CuLine)
                If GetStrLine = 3 Then
                    DisplayText = Split(DisplayText, vbCrLf)(1) & vbCrLf & _
                                  Split(DisplayText, vbCrLf)(2) & vbCrLf & _
                                  Split(DisplayText, vbCrLf)(3)
                End If
                
                If DisplayText <> "" Then DisplayText = DisplayText & vbCrLf
                MsgTime = GetTickCount: MsgTime2 = GetTickCount
            Else
                DisplayText = DisplayText & NLines(CuLine) & vbCrLf
                LLineTime = GetTickCount
            End If
            CuLine = CuLine + 1
            Exit Sub
        End If
    End If
    
    If Mode = "scroll" Then LLineTime = GetTickCount
    
    CuLine = CuLine + 1
    
    GoTo Exec
End Sub
Public Sub ReadNovel()
    PART = ESave.GetData("PART")
    If PART = "" Then PART = "0"
    TimeLine = Val(ESave.GetData("TIMELINE"))
    If TimeLine = 0 Then TimeLine = 1
    
    Dim temp As String, text As String
    
    Open App.path & "\article\PART " & PART & ".mss" For Input As #1
    Do While Not EOF(1)
        Line Input #1, temp
        text = text & temp & vbCrLf
    Loop
    Close #1
    NLines = Split(text, vbCrLf)
    
    CuLine = 0
    DisplayText = ""
    BG = "": FG = "": BGCOLOR = argb(255, 0, 0, 0)
    ReDim Roles(0)
    BGM.Dispose
    BGS.Dispose
    
    Call RunNovel
    
End Sub
Public Sub DrawQues()
    Dim pro As Single
    pro = (GetTickCount - QuesTime) / 1000
    pro = 1 - pro
    If pro < 0 Then pro = 0
    pro = Cubic(pro, 0, 0, 0, 1)

    PaintDC ODC, Page.CDC, 0, 0, Alpha:=pro * 1
    PaintDC BlurDC, Page.CDC, 0, 0, Alpha:=1 - pro * 1
    Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(130 - pro * 130, 0, 0, 0)
    
    Page.DrawImage "questionpad.png", GW / 2, GH / 2 + 200 + pro * 500, Alpha:=1, pos:=posOnCenter
    Page.Writes SpeakerName, GW / 2 - 208, GH / 2 - 59 + pro * 500, 24, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    Page.Writes Ques.Title, GW / 2 + 15, GH / 2 - 53 + pro * 500, 18, argb(255, 27, 27, 27), align:=StringAlignmentCenter
    
    Dim m As Long, aw As Long, ah As Long
    aw = Page.Res.ImgSize("answerpad.png", imgGetWidth)
    ah = Page.Res.ImgSize("answerpad.png", imgGetHeight)
    
    For i = 0 To UBound(Ques.Answers) - 1
        m = CheckMouse(GW / 2 - aw / 2, GH / 2 - 59 + 170 + i * 70 + pro * 500 - ah / 2, aw, ah)
        Page.DrawImage "answerpad.png", GW / 2, GH / 2 - 59 + 170 + i * 70 + pro * 500, Alpha:=IIf(m <> 0, 1, 0.5), pos:=posOnCenter
        Page.Writes (i + 1) & ". " & Ques.Answers(i).Caption, GW / 2 - 200, GH / 2 - 59 + 170 + i * 70 - 12 + pro * 500, 16, argb(200, 27, 27, 27)
        If m = 3 Then
            Mode = "msg"
            ECore.NewTransform
            If Ques.Answers(i).JTimeLine <> 0 Then TimeLine = Ques.Answers(i).JTimeLine: ESave.PutData "TIMELINE", TimeLine
            For s = 1 To UBound(Roles)
                If Roles(s).name = Speaker Then
                    Roles(s).Love = Roles(s).Love + Ques.Answers(i).LoveIncrase
                    ESave.PutData Roles(s).name & "'s love", Roles(s).Love
                    If Ques.Answers(i).LoveIncrase > 0 Then
                        MsgEvent = "ani+ballon_love"
                    ElseIf Ques.Answers(i).LoveIncrase < 0 Then
                        MsgEvent = "ani+ballon_han"
                    End If
                    Exit For
                End If
            Next
            Call RunNovel
        End If
    Next
End Sub
Public Sub DrawMsg()
    Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(130, 0, 0, 0)
    
    Dim pro As Single

    pro = (GetTickCount - MsgTime2) / 2000
    If pro > 1 Then
        If MsgEvent Like "looptext:*" Then
            pro = pro - Int(pro)
        Else
            pro = 1
        End If
    End If
    pro = Cubic(pro, 0, 1, 1, 1)
    
    Dim X As Long
    X = GW / 2
    For i = 1 To UBound(Roles)
        X = X - (Page.Res.ImgSize(Roles(i).name & "-" & Roles(i).Face & ".png", imgGetWidth) + 50) / 2
    Next
    
    X = X + 30
    
    For i = 1 To UBound(Roles)
        Page.DrawImage Roles(i).name & "-" & Roles(i).Face & ".png", X, GH + 140 + IIf(Page.Res.ImgSize(Roles(i).name & "-" & Roles(i).Face & ".png", imgGetHeight) < 400, -400, 0), Alpha:=IIf(Speaker <> Roles(i).name And Speaker <> "aside", 0.7, 1), pos:=posOnBottom
        X = X + (Page.Res.ImgSize(Roles(i).name & "-" & Roles(i).Face & ".png", imgGetWidth) + 50)
    Next
    
    If AddPos <= Len(AddingText) Then
        If GetTickCount - MsgTime >= 25 Then DisplayText = DisplayText & Mid(AddingText, AddPos, 1): AddPos = AddPos + 1: MsgTime = GetTickCount
    End If
    
    Page.DrawImage IIf(Speaker = "aside", "aside", "") & "dialog.png", GW / 2, GH - 130, Alpha:=1, pos:=posOnCenter
    If CheckMouse2 = mMouseUp Then
        If AddPos > Len(AddingText) Then
            MsgEvent = ""
            Call RunNovel
        Else
            'Skip
            SE.Play "Cursor.mp3"
            DisplayText = DisplayText & Mid(AddingText, AddPos, Len(AddingText) - AddPos)
            AddPos = Len(AddingText)
            ECore.NewTransform
        End If
    End If
    
    Page.Writes DisplayText, GW / 2 - 300, GH - 190, 18, argb(200, 27, 27, 27)
    If Speaker <> "aside" Then Page.Writes SpeakerName, GW / 2 - 5, GH - 78, 24, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    
    If MsgEvent Like "ani+*" Then
        'Page.SwitchChannel MsgEvent, "Default"
        Page.DrawAnimation MsgEvent, GW - 200, GH - 140
    End If
    
    If MsgEvent Like "*text:*" Then
        Page.Writes Split(MsgEvent, "text:")(1), GW / 2, 100 - 90 * pro, 48, argb(pro * 255, 255, 255, 255), align:=StringAlignmentCenter
    End If
End Sub
Public Sub DrawScroll()
    Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(170, 0, 0, 0)
    If CheckMouse(0, 0, GW, GH) = mMouseUp Then
        Call RunNovel: SE.Play "Cursor.mp3"
    End If
    
    Dim Lines() As String, pro As Single
    Lines = Split(DisplayText, vbCrLf)
    pro = (GetTickCount - LLineTime) / 1500
    If pro > 1 Then pro = 1
    pro = Cubic(pro, 0, 1, 1, 1)
    
    Dim Y As Long
    Y = GH / 2 - (UBound(Lines) - 2) * 50 - 50 * pro
    For i = 0 To UBound(Lines) - 1
        Page.Writes Lines(i), GW / 2, Y, 20, argb(IIf(i = UBound(Lines) - 1, pro, 1) * 255, 255, 255, 255), align:=StringAlignmentCenter
        Y = Y + 50
    Next
    
    If GetTickCount - LLineTime >= 5000 Then Call RunNovel
End Sub
Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear BGCOLOR  '清空画布
    If BG <> "" Then Page.DrawImage BG, 0, 0, Alpha:=1
    If FG <> "" Then Page.DrawImage FG, 0, 0, Alpha:=1
    
    
    If Mode = "msg" Then
        Call DrawMsg
    ElseIf Mode = "scroll" Then
        Call DrawScroll
    ElseIf Mode = "question" Then
        Call DrawQues
    End If
    
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
    
    Page.Res.NewImages App.path & "\assets\bg", 1000, 740
    Page.Res.NewImages App.path & "\assets\fg", 1000, 740
    Page.Res.NewImages App.path & "\assets\ui", 0.68
    Page.Res.NewImages App.path & "\assets\face"
    Page.Res.NewImages App.path & "\assets\animation"
    
    Page.LoadAnimationsFromDir App.path & "\animation"
    Page.CreatePlayAnimation "ballon_jing", "ani+ballon_jing", "Default"
    Page.CreatePlayAnimation "ballon_ques", "ani+ballon_ques", "Default"
    Page.CreatePlayAnimation "ballon_love", "ani+ballon_love", "Default"
    Page.CreatePlayAnimation "ballon_angry", "ani+ballon_angry", "Default"
    Page.CreatePlayAnimation "ballon_fan", "ani+ballon_fan", "Default"
    Page.CreatePlayAnimation "ballon_han", "ani+ballon_han", "Default"
    Page.CreatePlayAnimation "ballon_happy", "ani+ballon_happy", "Default"
    Page.CreatePlayAnimation "ballon_silence", "ani+ballon_silence", "Default"
    
    '创建页面
    ECore.Add Page, "NovelPage"
    ReDim NLines(0)
    BlurDC = CreateCDC(GW, GH)
    ODC = CreateCDC(GW, GH)
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub

Private Sub Class_Terminate()
    DeleteObject BlurDC
    DeleteObject ODC
End Sub
