VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DancePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Private Type NoteItem
        Direction As Integer
        Time As Single
        Hand As Long
    End Type
    Private Type NoteList
        Note() As NoteItem
    End Type
    Private Type EXCEMark
        Time As Long
        Text As String
        X As Long
        y As Long
        R As Byte
        g As Byte
        B As Byte
    End Type
    Dim Exce() As EXCEMark
    Dim NL As NoteList
    Dim NowNote As Long, Heart As Long, HurtTime As Long
    Dim Dir1 As String, Dir2 As String
    Dim BadMark As Boolean
'======================================

Public Sub Update()
    If Err.number <> 0 Then ErrCrash Err.number, Err.Description
    On Error Resume Next

    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear '清空画布

    Page.DrawImage "endforest.jpg", 0, 0, alpha:=1
    Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(100, 0, 0, 0)
    
    Page.DrawImage "yy-dance.png", GW / 2, GH / 2 - 80, alpha:=1, Pos:=posOnCenter
    Page.DrawImage "yy-" & Dir1 & "1.png", GW / 2, GH / 2 - 80, alpha:=1, Pos:=posOnCenter
    Page.DrawImage "yy-" & Dir2 & "2.png", GW / 2, GH / 2 - 80, alpha:=1, Pos:=posOnCenter
    
    Page.Paint 1, 150, GH / 2 + 230, 100, 100, argb(255, 27, 27, 27), Pos:=posOnCenter
    Page.Paint 1, 150, GH / 2 + 230, 80, 80, argb(255, 64, 64, 64), Pos:=posOnCenter
    
    Dim DrawX As Long
    For i = NowNote To UBound(NL.Note)
        If i >= UBound(NL.Note) Then Exit For
        DrawX = 100 + (NL.Note(i).Time - BGM.position) / 4 * (GW + 64) + 64
        Page.Paint 1, DrawX, GH / 2 + 230, 64 + 124 * (NL.Note(i).Time - BGM.position) / 4, 64 + 124 * (NL.Note(i).Time - BGM.position) / 4, IIf(i = NowNote, argb(255, 254, 84, 57), argb(128, 255, 255, 255)), style:=1, Size:=3, Pos:=posOnCenter
        Page.DrawImage "block" & NL.Note(i).Direction & ".png", DrawX, GH / 2 + 230, Pos:=posOnCenter, alpha:=1
        If DrawX > GW Then Exit For
        If DrawX < -32 Then
            NowNote = NowNote + 1
            If GetTickCount - HurtTime > 5000 Then
                Heart = Heart - 1
                HurtTime = GetTickCount
                SE.Play "Damage2.mp3"
            End If
        End If
    Next
    
    For i = 1 To 5
        Page.DrawImage "heart.png", 20 + (i - 1) * 48, 20, alpha:=IIf(Heart >= i, 1, 0.3)
    Next
    If GetTickCount - HurtTime <= 5000 Then Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(IIf(GetTickCount Mod 300 < 150, 20, 0), 255, 0, 0)
    
    Dim ExceStr As String
    ExceStr = "EXCELLENT"
    
    For i = 1 To UBound(Exce)
        If i > UBound(Exce) Then Exit For
        ExceStr = Exce(i).Text
        Pro = Cubic((GetTickCount - Exce(i).Time) / 1000, 0, 1, 1, 1)
        If Pro > 1 Then
            Exce(i) = Exce(UBound(Exce))
            ReDim Preserve Exce(UBound(Exce) - 1)
            GoTo SkipThis
        End If
        For S = 1 To Len(ExceStr)
            Page.Writes Mid(ExceStr, S, 1), Exce(i).X + Pro * (S - Len(ExceStr) / 2) * 40, Exce(i).y - 9, 18, argb(255 - 255 * Pro, Exce(i).R, Exce(i).g, Exce(i).B), align:=StringAlignmentCenter, style:=FontStyleBold
        Next
SkipThis:
    Next
    
    Page.Writes Int(BGM.position / BGM.length * 100) & "%", GW / 2, GH - 60, 32, argb(255, 255, 255, 255), align:=StringAlignmentCenter

    If BGM.length = BGM.position Then
        SE.Play "Applause1.mp3"
        NovelPage.RunNovel
        ECore.NewTransform transHighLight, 3000, "NovelPage"
    End If
    
    If Heart <= 0 Then
        SE.Play "Collapse4.mp3"
        NowNote = 1
        BGM.Stops
        BGM.Play
        BadMark = False
        HurtTime = 0
        Heart = 5
        Dir1 = "down": Dir2 = "down"
        ReDim Exce(0)
        ECore.NewTransform transFallDark, 2000
    End If
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub
Public Sub HitNote()
    Dim Dir As String, Pos As Single
    Pos = BGM.position
    
    With NL.Note(NowNote)
        Select Case .Direction
            Case 0: Dir = "left"
            Case 1: Dir = "right"
            Case 2: Dir = "up"
            Case 3: Dir = "down"
        End Select
        
        If .Hand = 1 Then
            Dir1 = Dir
        Else
            Dir2 = Dir
        End If
    End With
    
    Dim Pitch As Single, Text As String
    Pitch = Abs(NL.Note(NowNote).Time - Pos)
    
    If BadMark Then
        Text = "BAD"
        BadMark = False
    ElseIf Pitch <= 0.1 Then
        Text = "EXCELLENT"
    ElseIf Pitch <= 0.2 Then
        Text = "GOOD"
    End If
    
    If Text <> "" Then
        ReDim Preserve Exce(UBound(Exce) + 1)
        With Exce(UBound(Exce))
            .X = 150
            .y = GH / 2 + 230
            .Time = GetTickCount
            .R = Int(Rnd * 125 + 128)
            .g = Int(Rnd * 125 + 128)
            .B = Int(Rnd * 125 + 128)
            .Text = Text
        End With
    End If
    
    SE.Play "hit.wav"
End Sub
Public Sub Enter()
    '页面进入事件
    ReDim Exce(0)
    BGM.Create App.Path & "\music\bgm\OST 32 - Across The Waters[Short].mp3"
    'BGM.Create App.Path & "\music\bgm\OST 27 - It's Our Fight.mp3"
    BGM.Play
    BGS.Stops
    NowNote = 1: Heart = 5
    BadMark = False
    Dir1 = "down": Dir2 = "down"
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
Public Sub KeyUp(KeyCode As Integer)
    Dim Direction As Long ', Hand As Long
    
    'If KeyCode = vbKeyA Then Direction = 0: Hand = 1
    'If KeyCode = vbKeyD Then Direction = 1: Hand = 1
    'If KeyCode = vbKeyW Then Direction = 2: Hand = 1
    'If KeyCode = vbKeyS Then Direction = 3: Hand = 1
    
    Direction = -1
    If KeyCode = vbKeyLeft Then Direction = 0 ': Hand = 2
    If KeyCode = vbKeyRight Then Direction = 1 ': Hand = 2
    If KeyCode = vbKeyUp Then Direction = 2 ': Hand = 2
    If KeyCode = vbKeyDown Then Direction = 3 ': Hand = 2
    
    'ReDim Preserve NL.Note(UBound(NL.Note) + 1)
    'With NL.Note(UBound(NL.Note))
    '    .Direction = Direction
    '    .Hand = Hand
    '    .Time = BGM.position
    'End With
    
    'Exit Sub
    If Abs(NL.Note(NowNote).Time - BGM.position) <= 0.5 Then
        If Direction = NL.Note(NowNote).Direction Then
            Call HitNote: NowNote = NowNote + 1
        ElseIf Direction <> -1 Then
            BadMark = True
        End If
    End If
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.Res.HotLoad = True
    Page.Res.NewImages App.Path & "\assets\face"
    Page.Res.NewImages App.Path & "\assets\bg", GW, GH
    Page.Res.NewImages App.Path & "\assets\fg", GW, GH
    Page.Res.NewImages App.Path & "\assets\danceui", 64, 64
    
    ReDim NL.Note(0)
    Open App.Path & "\assets\yydance.note" For Binary As #1
    Get #1, , NL
    Close #1
    
    '创建页面
    ECore.Add Page, "DancePage"
    NowNote = 1: Heart = 5
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
