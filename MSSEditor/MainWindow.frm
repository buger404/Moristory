VERSION 5.00
Begin VB.Form MainWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H00303030&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Moristory Script Editor"
   ClientHeight    =   8208
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8208
   ScaleWidth      =   8580
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer StateTimer 
      Interval        =   1000
      Left            =   8040
      Top             =   600
   End
   Begin VB.Timer BackupTimer 
      Interval        =   60000
      Left            =   8040
      Top             =   144
   End
   Begin VB.ComboBox BGList 
      Appearance      =   0  'Flat
      Height          =   336
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   384
      Width           =   2580
   End
   Begin VB.ComboBox FGList 
      Appearance      =   0  'Flat
      Height          =   336
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   864
      Width           =   2580
   End
   Begin VB.ListBox MSSList 
      Appearance      =   0  'Flat
      Height          =   4584
      Left            =   480
      TabIndex        =   15
      Top             =   2928
      Width           =   7692
   End
   Begin VB.CommandButton FastBtn 
      BackColor       =   &H00303030&
      Caption         =   "-"
      Height          =   324
      Left            =   7608
      TabIndex        =   14
      Top             =   2424
      Width           =   516
   End
   Begin VB.CommandButton AddBtn 
      BackColor       =   &H00303030&
      Caption         =   "+"
      Height          =   324
      Left            =   6816
      TabIndex        =   13
      Top             =   2424
      Width           =   540
   End
   Begin VB.ComboBox ModeList 
      Appearance      =   0  'Flat
      Height          =   336
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1848
      Width           =   2580
   End
   Begin VB.TextBox TitleText 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1368
      TabIndex        =   10
      Top             =   1872
      Width           =   2580
   End
   Begin VB.TextBox InputBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   480
      TabIndex        =   8
      Top             =   2448
      Width           =   6060
   End
   Begin VB.ComboBox WeatherList 
      Appearance      =   0  'Flat
      Height          =   336
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1344
      Width           =   2580
   End
   Begin VB.ComboBox SPKList 
      Appearance      =   0  'Flat
      Height          =   336
      Left            =   1368
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1368
      Width           =   2580
   End
   Begin VB.ComboBox BGSList 
      Appearance      =   0  'Flat
      Height          =   336
      Left            =   1368
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   864
      Width           =   2580
   End
   Begin VB.ComboBox BGMList 
      Appearance      =   0  'Flat
      Height          =   336
      Left            =   1368
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   384
      Width           =   2580
   End
   Begin VB.Label SaveState 
      Height          =   348
      Left            =   0
      TabIndex        =   21
      Top             =   7896
      Width           =   300
   End
   Begin VB.Label StateText 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "..."
      ForeColor       =   &H00FFFFFF&
      Height          =   348
      Left            =   0
      TabIndex        =   20
      Top             =   7896
      Width           =   8628
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BG"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008C8080&
      Height          =   276
      Left            =   4320
      TabIndex        =   19
      Top             =   384
      Width           =   288
   End
   Begin VB.Label FGLab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FG"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008C8080&
      Height          =   276
      Left            =   4320
      TabIndex        =   18
      Top             =   864
      Width           =   264
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008C8080&
      Height          =   276
      Left            =   4296
      TabIndex        =   11
      Top             =   1848
      Width           =   588
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008C8080&
      Height          =   276
      Left            =   456
      TabIndex        =   9
      Top             =   1848
      Width           =   432
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Weather"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008C8080&
      Height          =   276
      Left            =   4296
      TabIndex        =   6
      Top             =   1392
      Width           =   816
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SPK"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008C8080&
      Height          =   276
      Left            =   456
      TabIndex        =   4
      Top             =   1368
      Width           =   372
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BGS"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008C8080&
      Height          =   276
      Left            =   456
      TabIndex        =   2
      Top             =   864
      Width           =   408
   End
   Begin VB.Label BGMLab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BGM"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008C8080&
      Height          =   276
      Left            =   456
      TabIndex        =   0
      Top             =   384
      Width           =   492
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu FileOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu FileSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu fcboard 
         Caption         =   "FromClipboard"
      End
   End
   Begin VB.Menu Createbtn 
      Caption         =   "Create"
      Begin VB.Menu newrolebtn 
         Caption         =   "New Role"
      End
      Begin VB.Menu removeroleBtn 
         Caption         =   "Remove Role"
      End
   End
   Begin VB.Menu changebtn 
      Caption         =   "Change"
      Begin VB.Menu facebtn 
         Caption         =   "Face"
      End
   End
   Begin VB.Menu saveposBtn 
      Caption         =   "SavePos"
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public EditIndex As String
Dim SaveMark As Boolean
Private Sub AddBtn_Click()
    MSSList.AddItem "", MSSList.ListIndex + 1
    MSSList.ListIndex = MSSList.ListIndex + 1
End Sub

Private Sub BackupTimer_Timer()
    Call MakeBackup
End Sub

Private Sub BGMList_Click()
    If MSSList.ListIndex = -1 Then Exit Sub
    SaveMark = False
    MSSList.AddItem "*bgm\" & BGMList.List(BGMList.ListIndex), MSSList.ListIndex + 1
    MSSList.ListIndex = MSSList.ListIndex + 1
End Sub
Private Sub BGSList_Click()
    If MSSList.ListIndex = -1 Then Exit Sub
    SaveMark = False
    MSSList.AddItem "*bgs\" & BGSList.List(BGSList.ListIndex), MSSList.ListIndex + 1
    MSSList.ListIndex = MSSList.ListIndex + 1
End Sub

Private Sub FastBtn_Click()
    If MSSList.ListIndex <= 2 Then Exit Sub
    SaveMark = False
    Dim temp As Long
    temp = MSSList.ListIndex
    MSSList.RemoveItem MSSList.ListIndex
    MSSList.ListIndex = temp - 1
End Sub

Private Sub fcboard_Click()
    Dim temp() As String
    temp = Split(Clipboard.GetText, vbCrLf)
    
    MSSList.Clear
    MSSList.AddItem "#这不是你应该看的文件。"
    MSSList.AddItem "*title\"
    MSSList.AddItem "*mode\msg"
    
    MSSList.ListIndex = 2
    
    Dim temp2() As String
    
    For i = 0 To UBound(temp)
        temp(i) = Replace(temp(i), "。。", "□□")
        temp(i) = Replace(temp(i), "□。", "□□")
        If Len(temp(i)) > 29 Then
            temp2 = Split(temp(i), "。")
            For s = 0 To UBound(temp2)
                MSSList.AddItem temp2(s) & "。"
                MSSList.List(MSSList.ListCount - 1) = Replace(MSSList.List(MSSList.ListCount - 1), "□", "。")
            Next
        Else
            MSSList.AddItem temp(i)
            MSSList.List(MSSList.ListCount - 1) = Replace(MSSList.List(MSSList.ListCount - 1), "□", "。")
        End If
    Next
    
    SaveMark = False
End Sub

Private Sub FGList_Click()
    If MSSList.ListIndex = -1 Then Exit Sub
    SaveMark = False
    MSSList.AddItem "*fg\" & FGList.List(FGList.ListIndex), MSSList.ListIndex + 1
    MSSList.ListIndex = MSSList.ListIndex + 1
End Sub
Private Sub BGList_Click()
    If MSSList.ListIndex = -1 Then Exit Sub
    SaveMark = False
    MSSList.AddItem IIf(BGList.ListIndex = BGList.ListCount - 1, "bg\0,0,0", "*bg\" & BGList.List(BGList.ListIndex)), MSSList.ListIndex + 1
    MSSList.ListIndex = MSSList.ListIndex + 1
End Sub
Private Sub MakeBackup()
    If Dir(App.Path & "\..\article\PART " & EditIndex & ".mss") = "" Then Exit Sub
    FileCopy App.Path & "\..\article\PART " & EditIndex & ".mss", _
    App.Path & "\backup\PART " & EditIndex & " - " & Year(Now) & "." & Month(Now) & "." & Day(Now) & "  " & Hour(Now) & "_" & Minute(Now) & "_" & Second(Now) & ".mss"
End Sub

Private Sub FileOpen_Click()
    Dim temp As String
    
    temp = VBA.InputBox("PART几？", "打开", "0")
    
    If Dir(App.Path & "\..\article\PART " & temp & ".mss") <> "" Then
        If MsgBox("这个PART已经存在了，覆盖吗？", vbYesNo) = vbNo Then End
        If MsgBox("真的吗？", vbYesNo) = vbNo Then End
        If MsgBox("你没有手滑要覆盖吗？", vbYesNo) = vbNo Then End
        If MsgBox("真的真的覆盖？", vbYesNo) = vbNo Then End
        EditIndex = temp
        Call MakeBackup
    End If
    
    EditIndex = temp
    Me.Caption = "Moristory Script Editor - PART " & EditIndex
    
    MSSList.Clear
    MSSList.AddItem "#这不是你应该看的文件。"
    MSSList.AddItem "*title\"
    MSSList.AddItem "*mode\msg"
    
    MSSList.ListIndex = 2
    
    SaveMark = False
End Sub

Private Sub FileSave_Click()
    Call MakeBackup
    With MSSList
        Open App.Path & "\..\article\PART " & EditIndex & ".mss" For Output As #1
        For i = 0 To MSSList.ListCount - 1
            Print #1, MSSList.List(i)
        Next
        Close #1
    End With
    SaveMark = True
    'MsgBox "Save OK !!!"
End Sub

Private Sub Form_Load()
    DirInto BGMList, App.Path & "\..\music\bgm\"
    DirInto BGSList, App.Path & "\..\music\bgs\"
    DirInto BGList, App.Path & "\..\assets\bg\"
    BGList.AddItem "RGB Color"
    DirInto FGList, App.Path & "\..\assets\fg\"
    
    With WeatherList
        .AddItem ""
        .AddItem "snow"
        .AddItem "snowstorm"
        .ListIndex = 0
    End With
    
    With ModeList
        .AddItem "msg"
        .AddItem "scroll"
        .ListIndex = 0
    End With
    
    With SPKList
        .AddItem "aside"
        .AddItem "me"
        .ListIndex = 0
    End With
    
    MSSList.AddItem "#这不是你应该看的文件。"
    MSSList.AddItem "*title\"
    MSSList.AddItem "*mode\msg"
    
    MSSList.ListIndex = 2
    
    Dim temp As String
    
    If Dir(App.Path & "\backup", vbDirectory) = "" Then MkDir App.Path & "\backup"
    
    EditIndex = VBA.InputBox("PART几？", "打开", "0")
    Me.Caption = "Moristory Script Editor - PART " & EditIndex
    If Dir(App.Path & "\..\article\PART " & EditIndex & ".mss") <> "" Then
        If MsgBox("这个PART已经存在了，你不会覆盖掉它对不对？", vbYesNo) = vbNo Then End
        If MsgBox("真的吗？", vbYesNo) = vbNo Then End
        If MsgBox("你没有手滑吗？", vbYesNo) = vbNo Then End
        If MsgBox("真的真的？", vbYesNo) = vbNo Then End
        Call MakeBackup
        With MSSList
            .Clear
            Open App.Path & "\..\article\PART " & EditIndex & ".mss" For Input As #1
            Do While Not EOF(1)
                Line Input #1, temp
                .AddItem temp
                If temp Like "*title\*" Then InputBox.Text = Split(temp, "\")(1)
            Loop
            Close #1
            SaveMark = True
        End With
    End If
    
End Sub

Public Sub DirInto(obj As ComboBox, folder As String)
    Dim File As String
    File = Dir(folder)
    obj.AddItem ""
    Do While File <> ""
        obj.AddItem File
        File = Dir()
        DoEvents
    Loop
    obj.ListIndex = 0
End Sub
Private Sub GetRoles()
    With SPKList
        .Clear
        .AddItem "aside"
        .AddItem "me"
    End With
    Dim List() As String, temp() As String
    ReDim List(0)
    For i = 2 To MSSList.ListIndex
        temp = Split(MSSList.List(i), "\")
        If UBound(temp) = 2 Then
            If temp(0) = "role" Then
                If temp(2) = "add" Then
                    ReDim Preserve List(UBound(List) + 1)
                    List(UBound(List)) = temp(1)
                ElseIf temp(2) = "remove" Then
                    For s = 1 To UBound(List)
                        If List(s) = temp(1) Then List(s) = List(UBound(List)): ReDim Preserve List(UBound(List) - 1): Exit For
                    Next
                End If
            End If
        End If
    Next
    
    Dim State As String
    For i = 1 To UBound(List)
        SPKList.AddItem List(i)
    Next
    
    For i = 0 To SPKList.ListCount - 1
        State = State & i & "." & SPKList.List(i) & "  "
    Next
    
    State = State & SPKList.ListCount & ".dark  "
    
    StateText.Caption = State
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If SaveMark = False Then Cancel = 1: MsgBox "你还没保存，沙雕。", 16
End Sub

Private Sub InputBox_Change()
    If MSSList.ListIndex <= 2 Then Exit Sub
    SaveMark = False
    MSSList.List(MSSList.ListIndex) = InputBox.Text
End Sub

Private Sub ModeList_Click()
    If MSSList.ListIndex = -1 Then Exit Sub
    SaveMark = False
    MSSList.List(2) = "*mode\" & ModeList.List(ModeList.ListIndex)
End Sub

Private Sub MSSList_Click()
    If MSSList.ListIndex < 2 Then MSSList.ListIndex = 2
    
    If MSSList.ListIndex <= 2 Then Exit Sub
    
    Call GetRoles
    
    InputBox.Text = MSSList.List(MSSList.ListIndex)
End Sub

Private Sub MSSList_KeyUp(KeyCode As Integer, Shift As Integer)
    If GetAsyncKeyState(VK_RBUTTON) Then
        If MSSList.ListIndex = -1 Then Exit Sub
        SaveMark = False
        If Val(Chr(KeyCode)) = SPKList.ListCount Then
            MSSList.AddItem "say\dark", MSSList.ListIndex + 1
        Else
            MSSList.AddItem "say\" & SPKList.List(Val(Chr(KeyCode))), MSSList.ListIndex + 1
        End If
        MSSList.ListIndex = MSSList.ListIndex + 1
    End If
End Sub

Private Sub MSSList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 4 Then
        Call FastBtn_Click
    End If
End Sub

Private Sub saveposBtn_Click()
    If MSSList.ListIndex = -1 Then Exit Sub
    SaveMark = False
    MSSList.AddItem "*save\", MSSList.ListIndex + 1
    MSSList.ListIndex = MSSList.ListIndex + 1
End Sub

Private Sub SPKList_Click()
    If MSSList.ListIndex = -1 Then Exit Sub
    SaveMark = False
    MSSList.AddItem "say\" & SPKList.List(SPKList.ListIndex), MSSList.ListIndex + 1
    MSSList.ListIndex = MSSList.ListIndex + 1
End Sub

Private Sub StateTimer_Timer()
    SaveState.BackColor = IIf(SaveMark, RGB(0, 176, 240), RGB(255, 0, 0))
End Sub

Private Sub TitleText_Change()
    MSSList.List(1) = "*title\" & TitleText.Text
    SaveMark = False
End Sub

Private Sub WeatherList_Click()
    If MSSList.ListIndex = -1 Then Exit Sub
    SaveMark = False
    MSSList.AddItem "*weather\" & WeatherList.List(WeatherList.ListIndex), MSSList.ListIndex + 1
    MSSList.ListIndex = MSSList.ListIndex + 1
End Sub
