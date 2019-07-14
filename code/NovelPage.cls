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
    Dim CuLine As Long, NLines() As String
    Dim Part As String
    Dim LLineTime As Long
    Dim DisplayText As String
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
    Private Type NovelArg
        Speaker As String
        SpeakerName As String
        MsgEvent As String
        Title As String
        mode As String
        BG As String
        FG As String
        BGCOLOR As Long
        Roles() As Role
        BGMn As String
        BGSn As String
    End Type
    Dim TimeLine As Long
    Dim QuesTime As Long, Ques As Question, BlurDC As Long, ODC As Long
    Dim MsgTime As Long, AddingText As String, MsgTime2 As Long, AddPos As Long
    Dim N As NovelArg
    Dim SaveTime As Long
'======================================
Public Sub SaveNow(Optional Index As String = "")
    ESave.PutData "PART" & Index, Part
    ESave.PutData "TIMELINE" & Index, TimeLine
    ESave.PutData "CURRENTLINE" & Index, CuLine
    ESave.Save
    Open ESave.Path & "\running" & Index & ".dat" For Binary As #1
    Put #1, , N
    Close #1
    
    If Index = "" Then
        SaveTime = GetTickCount
        ECore.NewTransform transFadeIn, 1000
    End If
End Sub
Public Sub RepairSave(Index As String)
    FileCopy ESave.Path & "\running" & Index & ".dat", ESave.Path & "\running.dat"
    ESave.PutData "PART", ESave.GetData("PART" & Index)
    ESave.PutData "TIMELINE", Val(ESave.GetData("TIMELINE" & Index))
    ESave.PutData "CURRENTLINE", Val(ESave.GetData("CURRENTLINE" & Index))
    ESave.Save
End Sub
Public Sub ReadSave()
    Part = ESave.GetData("PART")
    TimeLine = Val(ESave.GetData("TIMELINE"))
    CuLine = Val(ESave.GetData("CURRENTLINE"))
    If Part = "" Then Part = "0"
    If TimeLine = 0 Then TimeLine = 1
    If Dir(ESave.Path & "\running.dat") <> "" Then
        Open ESave.Path & "\running.dat" For Binary As #1
        Get #1, , N
        Close #1
        If N.BGMn = "" Then
            BGM.Stops
        Else
            BGM.Create N.BGMn
            BGM.Play
        End If
        If N.BGSn = "" Then
            BGS.Stops
        Else
            BGS.Create N.BGSn
            BGS.Play
        End If
    End If
End Sub
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
            Case "*save"
                SE.Play "Save.mp3"
                CuLine = CuLine + 1
                Call SaveNow
                CuLine = CuLine - 1
            Case "go"
                Select Case Params(1)
                    Case "maze"
                        CuLine = CuLine + 1
                        MazePage.LoadMap
                        ECore.NewTransform transFallDark, 3000, "MazePage"
                        Exit Sub
                    Case "battle"
                        CuLine = CuLine + 1
                        ECore.NewTransform transFallDark, 3000, "BattlePage"
                        Exit Sub
                    Case "chess"
                        CuLine = CuLine + 1
                        ECore.NewTransform transFallDark, 3000, "TicTacToePage"
                        Exit Sub
                    Case "bxbattle"
                        CuLine = CuLine + 1
                        ECore.NewTransform transFallDark, 3000, "BXBattlePage"
                        Exit Sub
                End Select
            Case "***death"
                SE.Play "Collapse4.mp3"
                Call RepairSave(2)
                Call ReadSave
                Call ReadNovel(True)
                Call RunNovel
                ECore.NewTransform transFallDark, 3000
                Debug.Print Now, "你在 #" & TimeLine & " 时间线，PART " & Part & "的第" & CuLine & "行死了。"
                Exit Sub
            Case "question"
                AnswerStep = Val(Params(2))
                Ques.Title = Params(1)
                ReDim Ques.Answers(AnswerStep)
                PaintDC Page.CDC, ODC, 0, 0, alpha:=1
                BlurTo BlurDC, GDC, GameWindow
                N.mode = "question"
            Case "timeline"
                If Params(2) = "head" Then
                    If TimeLine <> Val(Params(1)) Then JumpTimeline = True: GoTo Exec
                ElseIf Params(2) = "end" And JumpTimeline = True Then
                    JumpTimeline = False: GoTo Exec
                End If
            Case "face"
                For i = 1 To UBound(N.Roles)
                    If N.Roles(i).name = Params(1) Then
                        N.Roles(i).Face = Params(2): Exit For
                    End If
                Next
                ECore.NewTransform
            Case "role"
                If Params(2) = "add" Then
                    ReDim Preserve N.Roles(UBound(N.Roles) + 1)
                    With N.Roles(UBound(N.Roles))
                        .name = Params(1)
                        .Face = "normal"
                        .Love = Val(ESave.GetData(.name & "'s love"))
                    End With
                ElseIf Params(2) = "remove" Then
                    For i = 1 To UBound(N.Roles)
                        If N.Roles(i).name = Params(1) Then
                            N.Roles(i) = N.Roles(UBound(N.Roles)): Exit For
                        End If
                    Next
                    ReDim Preserve N.Roles(UBound(N.Roles) - 1)
                End If
                ECore.NewTransform
            Case "say"
                DisplayText = ""
                N.Speaker = Params(1)
                N.SpeakerName = GetSpeakerName(N.Speaker)
            Case "clear"
                DisplayText = ""
                ECore.NewTransform
            Case "*bgm"
                N.BGMn = App.Path & "\music\bgm\" & Params(1)
                BGM.Create App.Path & "\music\bgm\" & Params(1)
                BGM.Play
            Case "*bgs"
                N.BGSn = App.Path & "\music\bgs\" & Params(1)
                BGS.Create App.Path & "\music\bgs\" & Params(1)
                BGS.Play
            Case "*play"
                SE.Play Params(1)
            Case "*title"
                N.Title = Params(1)
                GameWindow.Caption = "Moristory DEMO - PART " & Part & " , " & N.Title
            Case "*bg"
                N.BGCOLOR = argb(255, 0, 0, 0)
                N.BG = Params(1)
                ECore.NewTransform
            Case "*fg"
                N.BGCOLOR = argb(255, 0, 0, 0)
                N.FG = Params(1)
                ECore.NewTransform
            Case "*mode"
                N.mode = Params(1)
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
                N.BGCOLOR = argb(255, temp(0), temp(1), temp(2))
                N.BG = "": N.FG = ""
                ECore.NewTransform
            Case "window"
                If Params(1) = "Moristory" Then
                    GameWindow.Caption = "Moristory DEMO - PART " & Part & " , " & N.Title
                Else
                    GameWindow.Caption = Params(1)
                End If
            Case "*turn"
                Part = Params(1)
                ECore.NewTransform transFallDark, 1000
                Call ReadNovel
                Call SaveNow
                Exit Sub
            Case "*event"
                N.MsgEvent = Params(1)
                If N.MsgEvent Like "ani+*" Then
                    Page.SwitchChannel N.MsgEvent, "Default"
                End If
        End Select
    ElseIf Left(NLines(CuLine), 1) <> "#" Then
        If JumpTimeline Then CuLine = CuLine + 1: GoTo Exec
        If NLines(CuLine) = "++" Then
            DisplayText = ""
            If UBound(N.Roles) = 0 Then
                N.Speaker = IIf(N.Speaker = "me", "", "me")
            Else
                N.Speaker = IIf(N.Speaker = "me", N.Roles(1).name, "me")
            End If
            N.SpeakerName = GetSpeakerName(N.Speaker)
        Else
            If N.mode <> "scroll" Then
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
    
    If N.mode = "scroll" Then LLineTime = GetTickCount
    
    CuLine = CuLine + 1
    
    GoTo Exec
End Sub
Public Sub ReadNovel(Optional NoRun As Boolean = False)
    
    Dim temp As String, Text As String
    
    Open App.Path & "\article\PART " & Part & ".mss" For Input As #1
    Do While Not EOF(1)
        Line Input #1, temp
        Text = Text & temp & vbCrLf
    Loop
    Close #1
    NLines = Split(Text, vbCrLf)
    
    DisplayText = ""
    
    If Not NoRun Then
        CuLine = 0
        N.BG = "": N.FG = "": N.BGCOLOR = argb(255, 0, 0, 0)
        N.BGMn = "": N.BGSn = ""
        ReDim N.Roles(0)
        BGM.Dispose
        BGS.Dispose
        
        Call RunNovel
    End If
End Sub
Public Sub DrawQues()
    Dim pro As Single
    pro = (GetTickCount - QuesTime) / 1000
    pro = 1 - pro
    If pro < 0 Then pro = 0
    pro = Cubic(pro, 0, 0, 0, 1)

    PaintDC ODC, Page.CDC, 0, 0, alpha:=pro * 1
    PaintDC BlurDC, Page.CDC, 0, 0, alpha:=1 - pro * 1
    Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(130 - pro * 130, 0, 0, 0)
    
    Page.DrawImage "questionpad.png", GW / 2, GH / 2 + 200 + pro * 500, alpha:=1, Pos:=posOnCenter
    Page.Writes N.SpeakerName, GW / 2 - 208, GH / 2 - 59 + pro * 500, 24, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    Page.Writes Ques.Title, GW / 2 + 15, GH / 2 - 53 + pro * 500, 18, argb(255, 27, 27, 27), align:=StringAlignmentCenter
    
    Dim m As Long, aw As Long, ah As Long
    aw = Page.Res.ImgSize("answerpad.png", imgGetWidth)
    ah = Page.Res.ImgSize("answerpad.png", imgGetHeight)
    
    For i = 0 To UBound(Ques.Answers) - 1
        m = CheckMouse(GW / 2 - aw / 2, GH / 2 - 59 + 170 + i * 70 + pro * 500 - ah / 2, aw, ah)
        Page.DrawImage "answerpad.png", GW / 2, GH / 2 - 59 + 170 + i * 70 + pro * 500, alpha:=IIf(m <> 0, 1, 0.5), Pos:=posOnCenter
        Page.Writes (i + 1) & ". " & Ques.Answers(i).Caption, GW / 2 - 200, GH / 2 - 59 + 170 + i * 70 - 12 + pro * 500, 16, argb(200, 27, 27, 27)
        If m = 3 Then
            N.mode = "msg"
            ECore.NewTransform
            If Ques.Answers(i).JTimeLine <> 0 Then
                CuLine = CuLine - UBound(Ques.Answers) - 2
                Call SaveNow(2)
                CuLine = CuLine + UBound(Ques.Answers) + 2
                TimeLine = Ques.Answers(i).JTimeLine: ESave.PutData "TIMELINE", TimeLine
            End If
            
            For s = 1 To UBound(N.Roles)
                If N.Roles(s).name = N.Speaker Then
                    N.Roles(s).Love = N.Roles(s).Love + Ques.Answers(i).LoveIncrase
                    ESave.PutData N.Roles(s).name & "'s love", N.Roles(s).Love
                    If Ques.Answers(i).LoveIncrase > 0 Then
                        N.MsgEvent = "ani+ballon_love"
                    ElseIf Ques.Answers(i).LoveIncrase < 0 Then
                        N.MsgEvent = "ani+ballon_han"
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
        If N.MsgEvent Like "looptext:*" Then
            pro = pro - Int(pro)
        Else
            pro = 1
        End If
    End If
    pro = Cubic(pro, 0, 1, 1, 1)
    
    Dim x As Long
    x = GW / 2
    For i = 1 To UBound(N.Roles)
        x = x - (Page.Res.ImgSize(N.Roles(i).name & "-" & N.Roles(i).Face & ".png", imgGetWidth) + 50) / 2
    Next
    
    x = x + 30
    
    For i = 1 To UBound(N.Roles)
        Page.DrawImage N.Roles(i).name & "-" & N.Roles(i).Face & ".png", x, GH + 140 + IIf(Page.Res.ImgSize(N.Roles(i).name & "-" & N.Roles(i).Face & ".png", imgGetHeight) < 400, -400, 0), alpha:=IIf(N.Speaker <> N.Roles(i).name And Speaker <> "aside", 0.7, 1), Pos:=posOnBottom
        x = x + (Page.Res.ImgSize(N.Roles(i).name & "-" & N.Roles(i).Face & ".png", imgGetWidth) + 50)
    Next
    
    If AddPos <= Len(AddingText) Then
        If GetTickCount - MsgTime >= 25 And AddingText <> "" Then DisplayText = DisplayText & Mid(AddingText, AddPos, 1): AddPos = AddPos + 1: MsgTime = GetTickCount
    End If
    
    Page.DrawImage IIf(N.Speaker = "aside", "aside", "") & "dialog.png", GW / 2, GH - 130, alpha:=1, Pos:=posOnCenter
    If CheckMouse2 = mMouseUp Then
        If AddPos > Len(AddingText) Then
            N.MsgEvent = ""
            Call RunNovel
        Else
            'Skip
            SE.Play "Cursor.mp3"
            If AddingText <> "" Then DisplayText = DisplayText & Mid(AddingText, AddPos, Len(AddingText) - AddPos)
            AddPos = Len(AddingText)
            ECore.NewTransform
        End If
    End If
    
    Page.Writes DisplayText, GW / 2 - 300, GH - 190, 18, argb(200, 27, 27, 27)
    If N.Speaker <> "aside" Then Page.Writes N.SpeakerName, GW / 2 - 5, GH - 78, 24, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    
    If N.MsgEvent Like "ani+*" Then
        'Page.SwitchChannel MsgEvent, "Default"
        Page.DrawAnimation N.MsgEvent, GW - 200, GH - 140
    End If
    
    If N.MsgEvent Like "*text:*" Then
        Page.Writes Split(N.MsgEvent, "text:")(1), GW / 2, 100 - 90 * pro, 48, argb(pro * 255, 255, 255, 255), align:=StringAlignmentCenter
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
    
    Dim y As Long
    y = GH / 2 - (UBound(Lines) - 2) * 50 - 50 * pro
    For i = 0 To UBound(Lines) - 1
        Page.Writes Lines(i), GW / 2, y, 20, argb(IIf(i = UBound(Lines) - 1, pro, 1) * 255, 255, 255, 255), align:=StringAlignmentCenter
        y = y + 50
    Next
    
    If GetTickCount - LLineTime >= 3000 Then Call RunNovel
End Sub
Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear N.BGCOLOR  '清空画布
    If N.BG <> "" Then Page.DrawImage N.BG, 0, 0, alpha:=1
    If N.FG <> "" Then Page.DrawImage N.FG, 0, 0, alpha:=1
    
    If SaveTime <> 0 Then
        Page.Clear argb(255, 27, 27, 27)
        Page.ShowLoading GW / 2 - 64, GH / 2 - 64, 128, 128, 3, argb(255, 0, 176, 240), argb(255, 254, 84, 57), argb(255, 255, 255, 255)
        Page.Writes "你的存档已经保存", GW / 2, GH - 190, 20, argb(255, 242, 242, 242), align:=StringAlignmentCenter
        Page.Writes "Tips：死亡时你将会回到最近一个更改你时间线的分支。", GW / 2, GH - 50, 20, argb(180, 242, 242, 242), align:=StringAlignmentCenter
        If GetTickCount - SaveTime >= 5000 Then
            ECore.NewTransform transFadeIn, 1000
            SaveTime = 0
        End If
        Exit Sub
    End If
    
    If N.mode = "msg" Then
        Call DrawMsg
    ElseIf N.mode = "scroll" Then
        Call DrawScroll
    ElseIf N.mode = "question" Then
        Call DrawQues
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
    
    Page.Res.NewImages App.Path & "\assets\bg", 1000, 740
    Page.Res.NewImages App.Path & "\assets\fg", 1000, 740
    Page.Res.NewImages App.Path & "\assets\ui", 0.68
    Page.Res.NewImages App.Path & "\assets\face"
    Page.Res.NewImages App.Path & "\assets\animation"
    
    Page.LoadAnimationsFromDir App.Path & "\animation"
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
