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
    Dim Historys() As String, HistoryIndex As Long
    Dim SaveTime As Long, SaveSndMark As Boolean
    Dim LastMode As String
'======================================
Public Sub SaveNow(Optional index As String = "")
    ESave.PutData "PART" & index, Part
    ESave.PutData "TIMELINE" & index, TimeLine
    ESave.PutData "CURRENTLINE" & index, CuLine
    ESave.Save
    Open ESave.Path & "\running" & index & ".dat" For Binary As #1
    Put #1, , N
    Close #1
    
    If index = "" Then
        SaveTime = GetTickCount
        ECore.NewTransform transFadeIn, 1000
    End If
End Sub
Public Sub RepairSave(index As String)
    FileCopy ESave.Path & "\running" & index & ".dat", ESave.Path & "\running.dat"
    ESave.PutData "PART", ESave.GetData("PART" & index)
    ESave.PutData "TIMELINE", Val(ESave.GetData("TIMELINE" & index))
    ESave.PutData "CURRENTLINE", Val(ESave.GetData("CURRENTLINE" & index))
    ESave.Save
End Sub
Public Sub ReadSave()
    Part = ESave.GetData("PART")
    TimeLine = Val(ESave.GetData("TIMELINE"))
    CuLine = Val(ESave.GetData("CURRENTLINE"))
    If Part = "" Then Part = "-1"
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
    Else
        ReDim N.Roles(0)
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
            Case "demo"
                MsgBox "DEMO内容到此结束，感谢您的支持！请期待7.25的正式版本~", 64, "Moristory"
                ECore.NewTransform transFallDark, 2000, "MainPage"
                Exit Sub
            Case "*save"
                SE.Play "Save.mp3"
                CuLine = CuLine + 1
                Call SaveNow
                CuLine = CuLine - 1
            Case "*weather"
                WeatherLayer.ChangeWeather Params(1)
            Case "go"
                CuLine = CuLine + 1
                TipPage.ShowTip Params(1)
                ECore.NewTransform transFallDark, 1000, "TipPage"
                Exit Sub
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
                ReDim Preserve Historys(UBound(Historys) + 1)
                Historys(UBound(Historys)) = N.SpeakerName & "：" & vbCrLf
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
            Historys(UBound(Historys)) = Historys(UBound(Historys)) & NLines(CuLine) & vbCrLf
            CuLine = CuLine + 1
            Exit Sub
        End If
    End If
    
    If N.mode = "scroll" Then LLineTime = GetTickCount
    
    CuLine = CuLine + 1
    
    GoTo Exec
End Sub
Public Sub ReadNovel(Optional NoRun As Boolean = False)
    ReDim Historys(0)
    
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
    
    If AddPos <= Len(AddingText) Then
        If GetTickCount - MsgTime >= 25 And AddingText <> "" Then DisplayText = DisplayText & Mid(AddingText, AddPos, 1): AddPos = AddPos + 1: MsgTime = GetTickCount
    End If
    
    Dim x As Long
    x = GW / 2
    For i = 1 To UBound(N.Roles)
        If N.Roles(i).name <> "bg" Then x = x - (Page.Res.ImgSize(N.Roles(i).name & "-" & N.Roles(i).Face & ".png", imgGetWidth) + 50) / 2
    Next
    
    x = x + 30
    
    Dim DrawX As Long, DrawY As Long
    Dim DW As Long, DH As Long, TH As Long
    Dim DiaX As Long, DiaY As Long, RoleX As Long
    For i = 1 To UBound(N.Roles)
        If N.Roles(i).name = "bg" Then
            DrawX = x - (Page.Res.ImgSize(N.Roles(i - 1).name & "-" & N.Roles(i - 1).Face & ".png", imgGetWidth) + 0) + 210
            DrawY = GH - 400
        Else
            DrawX = x
            DrawY = GH + 140 + IIf(N.Roles(i).name = "xl", -250, 0) + IIf(Page.Res.ImgSize(N.Roles(i).name & "-" & N.Roles(i).Face & ".png", imgGetHeight) < 400, -400, 0)
            x = x + (Page.Res.ImgSize(N.Roles(i).name & "-" & N.Roles(i).Face & ".png", imgGetWidth) + 0)
        End If
            Page.DrawImage N.Roles(i).name & "-" & N.Roles(i).Face & ".png", DrawX, DrawY, alpha:=IIf(N.Speaker <> N.Roles(i).name And Speaker <> "aside", 0.7, 1), Pos:=posOnBottom
        If N.Roles(i).name = N.Speaker Then
            RoleX = DrawX
            DW = EF.GetWidth(Page.GG, DisplayText, 17, StringAlignmentNear, FontStyleRegular)
            DH = EF.GetHeight(Page.GG, DisplayText, 17, StringAlignmentNear, FontStyleRegular)
            TH = EF.GetHeight(Page.GG, N.SpeakerName, 18, StringAlignmentNear, FontStyleRegular)
            DiaX = DrawX + (x - DrawX) / 3 * 2
            DiaY = DrawY - Page.Res.ImgSize(N.Roles(i).name & "-" & N.Roles(i).Face & ".png", imgGetHeight) - DH - TH - 20
        End If
    Next
    
    Dim CMode As Boolean, CModeX As Long
    
    If DiaX + DW + 40 + 20 > GW - 20 Then
        CModeX = (GW - DW - 40 - 20 - 20) - DiaX
        CMode = True
        DiaX = GW - DW - 40 - 20 - 20
    End If
    If DiaY <= 20 Then
        CMode = True
        DiaY = 20
    End If
    
    Dim Path As Long, DiaR As Long, DiaB As Long, Radius As Long
    Dim BotX As Long
    DiaR = DiaX + DW + 40 + 20
    DiaB = DiaY + DH + 40 + TH + 8
    Radius = 40
    
    If RoleX <= DiaR Then
        BotX = DiaX + 20 + DW * 0.1 - CModeX
    Else
        BotX = DiaX + 20 + DW * 0.9
    End If
    
    '==================================================================
    '=                                                                =
    '=                               B                                =
    '=       A....╭──────────────────────────────────────────╮.....C        =
    '=            │  ███▋                             │              =
    '=            │  ▂▂▂▂▂▂▂▂▂                     │              =
    '=       K....│  ▂▂▂▂▂▂▂▂▂▂▂▂▂               │.....D        =
    '=            │  ▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂▂  ▼   │              =
    '=       J....╰────────  ────────────────────────────────╯.....E        =
    '=             I H....|/....G    F                                 =
    '=                                                                 =
    '===================================================================

        GdipCreatePath FillModeWinding, Path
    '===========================================================================
    '   A
        GdipAddPathArc Path, DiaX, DiaY, Radius, Radius, 180, 90
    '===========================================================================
    '   B
        GdipAddPathLine Path, DiaX + Radius, DiaY, DiaR - Radius, DiaY
    '===========================================================================
    '   C
        GdipAddPathArc Path, DiaR - Radius, DiaY, Radius, Radius, 270, 90
    '===========================================================================
    '   D
        GdipAddPathLine Path, DiaR, DiaY + Radius, DiaR, DiaB - Radius
    '===========================================================================
    '   E
        GdipAddPathArc Path, DiaR - Radius, DiaB - Radius, Radius, Radius, 0, 90
    '===========================================================================
    '   F
        GdipAddPathLine Path, DiaR - Radius, DiaB, BotX + 20, DiaB
    '===========================================================================
    If RoleX <= DiaR Then
        '===========================================================================
        '   G
            GdipAddPathLine Path, BotX + 20, DiaB, BotX, DiaB + 20
        '===========================================================================
        '   H
            GdipAddPathLine Path, BotX, DiaB + 20, BotX, DiaB
        '===========================================================================
    Else
        '===========================================================================
        '   G
            GdipAddPathLine Path, BotX + 20, DiaB, BotX + 20, DiaB + 20
        '===========================================================================
        '   H
            GdipAddPathLine Path, BotX + 20, DiaB + 20, BotX, DiaB
        '===========================================================================
    End If
    
    '===========================================================================
    '   I
        GdipAddPathLine Path, BotX, DiaB, DiaX + Radius, DiaB
    '===========================================================================
    '   J
        GdipAddPathArc Path, DiaX, DiaB - Radius, Radius, Radius, 90, 90
    '===========================================================================
    '   K
        GdipAddPathLine Path, DiaX, DiaB - Radius, DiaX, DiaY + Radius
        
        GdipClosePathFigure Path
    
    Dim cPro As Single
    cPro = (GetTickCount Mod 600) / 300
    If cPro <= 1 Then
        cPro = Cubic(cPro, 0, 0, 1, 1)
    Else
        cPro = Cubic(2 - cPro, 0, 1, 1, 1)
    End If
    
    If DiaX <> 0 Then
        GdipSetSolidFillColor ECore.pB, argb(IIf(CMode, 200, 255), 255, 255, 255)
        GdipFillPath Page.GG, ECore.pB, Path
        
        Page.Writes N.SpeakerName, DiaX + 20, DiaY + 20, 18, argb(255, 0, 176, 240)
        Page.Writes DisplayText, DiaX + 20, DiaY + 20 + TH + 8, 17, argb(255, 0, 0, 0)
        
        Page.DrawImage "continue.png", DiaR - 30, DiaB - 30 - 10 + 10 * cPro, alpha:=1 - 0.5 * cPro, Pos:=posOnCenter
    End If
    
    GdipDeletePath Path
    
    'Page.DrawImage IIf(N.Speaker = "aside", "aside", "") & "dialog.png", GW / 2, GH - 130, alpha:=1, Pos:=
    Dim BarPic As String, DiaColor As Long, DiaColor2 As Long
    BarPic = "asidebar.png"
    If DiaX = 0 Then
        If N.Speaker = "me" Or N.Speaker = "aside" Then
            DiaColor = argb(255, 32, 32, 32)
            DiaColor2 = argb(255, 212, 212, 212)
            BarPic = N.Speaker & "bar.png"
        Else
            DiaColor = argb(255, 232, 232, 232)
            DiaColor2 = argb(255, 74, 74, 74)
            BarPic = "darkbar.png"
        End If
    End If
    
    Page.Paint 0, 0, GH - 200, GW + 1, 201, IIf(BarPic = "darkbar.png", argb(255, 27, 27, 27), argb(255, 255, 255, 255))
    
    Page.DrawImage BarPic, 0, GH - 200, alpha:=1, Pos:=posOnBottom
    
    If DiaX <> 0 Then Page.Paint 0, 0, GH - 200, GW + 1, 201, argb(150, 0, 0, 0)
    
    Dim Hot As Long, Map() As Single
    If BGM.IsPlaying Then
        Map = BGM.GetMusicMap
        For i = 0 To UBound(Map)
            If i > 8 Then Exit For
            Page.PaintLine GW / 2 - i * 20, GH - 100 - Map(i) / 1500 * 80, GW / 2 - i * 20, GH - 100 + Map(i) / 1500 * 80, DiaColor2, 2
            Page.PaintLine GW / 2 + i * 20, GH - 100 - Map(i) / 1500 * 80, GW / 2 + i * 20, GH - 100 + Map(i) / 1500 * 80, DiaColor2, 2
        Next
    End If
    
    If DiaX = 0 Then
        DH = EF.GetHeight(Page.GG, DisplayText, 18, StringAlignmentCenter, FontStyleRegular)
        Page.Writes DisplayText, GW / 2, GH - 100 - DH / 2, 18, DiaColor, align:=StringAlignmentCenter
        Page.Writes N.SpeakerName, GW / 2, GH - 228, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter
        DW = EF.GetWidth(Page.GG, AddingText, 18, StringAlignmentCenter, FontStyleRegular)
        Page.DrawImage "continue.png", GW / 2 + DW / 2 + 10, GH - 100 + DH / 2 - 10 * cPro, alpha:=1 - 0.5 * cPro, Pos:=posOnBottom
    End If
    
    If CheckMouse(0, 0, GW, GH) = mMouseDown Then
        Mouse.state = 0
        If Mouse.button = 1 Then
            Call NextDialog
        ElseIf Mouse.button = 2 Then
            Call ShowHistory
        End If
    End If
    
    'Page.Writes DisplayText, GW / 2 - 300, GH - 190, 18, argb(200, 27, 27, 27)
    'If N.Speaker <> "aside" Then Page.Writes N.SpeakerName, GW / 2 - 5, GH - 78, 24, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    
    If N.MsgEvent Like "ani+*" Then
        'Page.SwitchChannel MsgEvent, "Default"
        Page.DrawAnimation N.MsgEvent, GW - 200, GH - 140
    End If
    
    If N.MsgEvent Like "*text:*" Then
        Page.Writes Split(N.MsgEvent, "text:")(1), GW / 2, 100 - 90 * pro, 48, argb(pro * 255, 255, 255, 255), align:=StringAlignmentCenter
    End If
    
    Page.DrawImage "xbtn.png", 20, 20, alpha:=1
    If CheckMouse2 >= mMouseIn And Mouse.button = 1 Then Call ShowHistory
    
    Page.DrawImage "escbtn.png", 20, 140, alpha:=1
    If CheckMouse2 >= mMouseIn And Mouse.button = 1 Then
        Mouse.state = 0: Mouse.button = 0
        ECore.NewTransform transDarkReturn, 1000, "MainPage"
    End If
End Sub
Public Sub KeyUp(KeyCode As Integer)
    If (KeyCode = vbKeyZ Or KeyCode = vbKeyReturn) And N.mode = "msg" Then
        NovelPage.NextDialog
    End If
    If (KeyCode = vbKeyEscape) And N.mode = "msg" Then
        ECore.NewTransform transDarkReturn, 1000, "MainPage"
    End If
    If KeyCode = vbKeyX And N.mode <> "his" Then
        NovelPage.ShowHistory
        Exit Sub
    End If
    If (KeyCode = vbKeyX Or KeyCode = vbKeyEscape) And N.mode = "his" Then
        NovelPage.CancelHistory
    End If
End Sub
Public Sub NextDialog()
    If N.mode = "question" Then Exit Sub
    If AddPos >= Len(AddingText) Then
        N.MsgEvent = ""
        Call RunNovel
    Else
        'Skip
        SE.Play "Cursor1.mp3"
        If AddingText <> "" Then DisplayText = DisplayText & Mid(AddingText, AddPos, Len(AddingText) - AddPos)
        AddPos = Len(AddingText)
        'ECore.NewTransform
    End If
End Sub
Public Sub ShowHistory()
    If N.mode = "his" Then Exit Sub
    
    SE.Play "Cancel.mp3"
    LastMode = N.mode
    HistoryIndex = UBound(Historys)
    N.mode = "his"
    QuesTime = GetTickCount
    PaintDC Page.CDC, ODC, 0, 0, alpha:=1
    BlurTo BlurDC, GDC, GameWindow
End Sub
Public Sub DrawHistory()
    Dim pro As Single, pro2 As Single, pro3 As Single
    pro = (GetTickCount - QuesTime) / 2000
    pro = 1 - pro
    If pro < 0 Then pro = 0
    pro = Cubic(pro, 0, 0, 0, 1)

    PaintDC ODC, Page.CDC, 0, 0, alpha:=pro * 1
    PaintDC BlurDC, Page.CDC, 0, 0, alpha:=1 - pro

    pro2 = (1 - pro) / 0.5
    If pro2 > 1 Then pro2 = 1
    
    Page.Paint 2, GW / 2, GH - pro2 * GH / 2, pro2 * 600, pro2 * 400, argb(255 - pro2 * 80, 255, 255, 255), Radius:=300 - pro2 * 300, Pos:=posOnCenter
    
    pro3 = (1 - pro - 0.5) / 0.5
    If pro3 < 0 Then pro3 = 0
    
    Dim DH As Long
    DH = EF.GetHeight(Page.GG, Historys(HistoryIndex), 16, StringAlignmentCenter, FontStyleRegular)
    
    Page.Writes Historys(HistoryIndex), GW / 2, GH / 2 - DH / 2, 16, argb(255 * pro3, 32, 32, 32), align:=StringAlignmentCenter
    
    Page.Paint 2, GW / 2, GH / 2 + 185, 600, 40, argb(255 * pro3, 255, 255, 255), Pos:=posOnCenter

    Page.Writes HistoryIndex & "/" & UBound(Historys), GW / 2, GH / 2 + 173, 16, argb(180 * pro3, 32, 32, 32), align:=StringAlignmentCenter
    
    Page.Paint 2, 0, GH - 50, GW + 1, 50, argb(180 * pro3, 255, 255, 255)
    
    Page.Writes "按Esc或右键退出", GW / 2, GH - 42, 18, argb(180 * pro3, 32, 32, 32), align:=StringAlignmentCenter
    
    If HistoryIndex < UBound(Historys) Then
        If Page.ShowColorButton(1, GW - 100, GH / 2 - 30, 60, 60, ">", argb(255 * pro3, 255, 255, 255), argb(190 * pro3, 50, 202, 185), argb(255 * pro3, 50, 202, 185), Size:=24) = 3 Then
            SE.Play "Cursor1.mp3"
            HistoryIndex = HistoryIndex + 1
        End If
    End If
    If HistoryIndex > 1 Then
        If Page.ShowColorButton(1, 40, GH / 2 - 30, 60, 60, "<", argb(255 * pro3, 255, 255, 255), argb(190 * pro3, 50, 202, 185), argb(255 * pro3, 50, 202, 185), Size:=24) = 3 Then
            SE.Play "Cursor1.mp3"
            HistoryIndex = HistoryIndex - 1
        End If
    End If
    
    If CheckMouse(0, 0, GW, GH) = mMouseDown And Mouse.button = 2 Then
        Mouse.state = 0
        Call CancelHistory
    End If
End Sub
Public Sub CancelHistory()
    If N.mode <> "his" Then Exit Sub
    N.mode = LastMode: SE.Play "Cancel.mp3": ECore.NewTransform
End Sub
Public Sub DrawScroll()
    Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(170, 0, 0, 0)
    If CheckMouse(0, 0, GW, GH) = mMouseUp Then
        Call RunNovel: SE.Play "Cursor1.mp3"
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
    
    If N.mode = "msg" Then
        Call DrawMsg
    ElseIf N.mode = "scroll" Then
        Call DrawScroll
    ElseIf N.mode = "question" Then
        Call DrawQues
    ElseIf N.mode = "his" Then
        Call DrawHistory
    End If
    
    If BGM.PlayState = musStopped And BGM.length <> 0 Then BGM.Play
    If BGS.PlayState = musStopped And BGS.length <> 0 Then BGS.Play
    
    If SaveTime <> 0 Then
        'Page.Clear argb(255, 27, 27, 27)
        Dim pro As Single
        pro = 1
        If GetTickCount - SaveTime <= 1000 Then
            pro = Cubic((GetTickCount - SaveTime) / 1000, 0, 0, 1, 1)
        ElseIf GetTickCount - SaveTime >= 2000 And GetTickCount - SaveTime <= 2500 Then
            pro = Cubic(1 - (GetTickCount - SaveTime - 2000) / 500, 0, 1, 1, 1)
        End If
        
        Page.Paint 2, GW / 2, pro * 160 - 60, 400, 60, argb(pro * 255, 255, 255, 255), 250, Pos:=posOnCenter
        
        If SaveSndMark = False And GetTickCount - SaveTime >= 1500 Then
            SE.Play "SaveDone.wav"
            SaveSndMark = True
        End If
        
        If GetTickCount - SaveTime <= 1500 Then
            Page.ShowLoading GW / 2 - 200 + 20, pro * 160 - 75, 30, 30, 3, argb(pro * 255, 0, 176, 240), argb(pro * 255, 254, 84, 57), argb(pro * 255, 255, 255, 255)
            Page.Writes "正在保存你的存档...", GW / 2, pro * 160 - 60 - 12, 18, argb(pro * 255, 64, 64, 64), align:=StringAlignmentCenter
        Else
            Page.Paint 1, GW / 2 - 200 + 20, pro * 160 - 75, 30, 30, argb(pro * 255, 9, 241, 117)
            Page.Writes "你的存档已经保存", GW / 2, pro * 160 - 60 - 12, 18, argb(pro * 255, 64, 64, 64), align:=StringAlignmentCenter
        End If
        
        If GetTickCount - SaveTime >= 2500 Then
            SaveSndMark = False
            SaveTime = 0
        End If
        'Exit Sub
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
    Page.Res.HotLoad = True
    Page.Res.NoLoadCrash = True
    
    Page.Res.NewImages App.Path & "\assets\bg", 1000, 740
    Page.Res.NewImages App.Path & "\assets\fg", 1000, 740
    Page.Res.NewImages App.Path & "\assets\ui", 0.68
    Page.Res.NewImages App.Path & "\assets\dialog", 1000 / 1520
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
