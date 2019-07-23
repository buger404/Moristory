VERSION 5.00
Begin VB.Form MainWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H00303030&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Moristory Script Editor"
   ClientHeight    =   8208
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   12624
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8208
   ScaleWidth      =   12624
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox SEList 
      Appearance      =   0  'Flat
      Height          =   336
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   1824
      Width           =   2892
   End
   Begin VB.Timer StateTimer 
      Interval        =   1000
      Left            =   144
      Top             =   216
   End
   Begin VB.ComboBox BGList 
      Appearance      =   0  'Flat
      Height          =   336
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   384
      Width           =   2892
   End
   Begin VB.ComboBox FGList 
      Appearance      =   0  'Flat
      Height          =   336
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   864
      Width           =   2892
   End
   Begin VB.CommandButton FastBtn 
      BackColor       =   &H00303030&
      Caption         =   "-"
      Height          =   324
      Left            =   7896
      TabIndex        =   14
      Top             =   2424
      Width           =   516
   End
   Begin VB.CommandButton AddBtn 
      BackColor       =   &H00303030&
      Caption         =   "+"
      Height          =   324
      Left            =   7152
      TabIndex        =   13
      Top             =   2424
      Width           =   540
   End
   Begin VB.ComboBox ModeList 
      Appearance      =   0  'Flat
      Height          =   336
      Left            =   7464
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1368
      Width           =   948
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
      Height          =   276
      Left            =   480
      TabIndex        =   8
      Top             =   2448
      Width           =   6468
   End
   Begin VB.ComboBox WeatherList 
      Appearance      =   0  'Flat
      Height          =   336
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1344
      Width           =   1068
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SE"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CEDB1A&
      Height          =   276
      Left            =   4296
      TabIndex        =   22
      Top             =   1872
      Width           =   228
   End
   Begin VB.Label SaveState 
      Height          =   348
      Left            =   0
      TabIndex        =   20
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
      TabIndex        =   19
      Top             =   7896
      Width           =   8844
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      Left            =   6744
      TabIndex        =   11
      Top             =   1392
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
      ForeColor       =   &H00CEDB1A&
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
      ForeColor       =   &H00CEDB1A&
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
   Begin VB.Menu musicMenu 
      Caption         =   "Music"
      Begin VB.Menu bgmMenu 
         Caption         =   "BGM"
         Begin VB.Menu bgmBtn 
            Caption         =   "None"
            Index           =   0
         End
      End
      Begin VB.Menu BGSMenu 
         Caption         =   "BGS"
         Begin VB.Menu bgsBtn 
            Caption         =   "None"
            Index           =   0
         End
      End
      Begin VB.Menu SEMenu 
         Caption         =   "SE"
         Begin VB.Menu seBtn 
            Caption         =   "None"
            Index           =   0
         End
      End
   End
   Begin VB.Menu PictureMenu 
      Caption         =   "Pictures"
      Begin VB.Menu bgMenu 
         Caption         =   "Background"
         Begin VB.Menu bgBtn 
            Caption         =   "None"
            Index           =   0
         End
      End
      Begin VB.Menu ForegroundMenu 
         Caption         =   "Foreground"
         Begin VB.Menu fgBtn 
            Caption         =   "None"
            Index           =   0
         End
      End
   End
   Begin VB.Menu saveposBtn 
      Caption         =   "SavePos"
   End
   Begin VB.Menu musaybtn 
      Caption         =   "Mu-Say"
   End
   Begin VB.Menu roleMenu 
      Caption         =   "Roles"
      Begin VB.Menu roleBtn 
         Caption         =   "dark - 未知人物"
         Index           =   0
      End
   End
   Begin VB.Menu rolectrlMenu 
      Caption         =   "RoleCtrl"
      Begin VB.Menu roleaddMenu 
         Caption         =   "Add"
         Begin VB.Menu roleaddBtn 
            Caption         =   "bm - 黑嘴"
            Index           =   0
         End
         Begin VB.Menu roleaddBtn 
            Caption         =   "xx - 兮兮"
            Index           =   1
         End
         Begin VB.Menu roleaddBtn 
            Caption         =   "xl - 雪狼"
            Index           =   2
         End
         Begin VB.Menu roleaddBtn 
            Caption         =   "km1 - 枯梦"
            Index           =   3
         End
         Begin VB.Menu roleaddBtn 
            Caption         =   "fj - 浮橘"
            Index           =   4
         End
         Begin VB.Menu roleaddBtn 
            Caption         =   "kx1 - 卡西"
            Index           =   5
         End
         Begin VB.Menu roleaddBtn 
            Caption         =   "kx2 - 卡茜"
            Index           =   6
         End
         Begin VB.Menu roleaddBtn 
            Caption         =   "yz - 芽子"
            Index           =   7
         End
         Begin VB.Menu roleaddBtn 
            Caption         =   "jy - 久悠"
            Index           =   8
         End
         Begin VB.Menu roleaddBtn 
            Caption         =   "bg - 冰棍"
            Index           =   9
         End
         Begin VB.Menu roleaddBtn 
            Caption         =   "ssr - 莎瑟瑞"
            Index           =   10
         End
      End
      Begin VB.Menu roleremoveBtn 
         Caption         =   "Remove"
         Begin VB.Menu removeroleBtn 
            Caption         =   "Cancel"
            Index           =   0
         End
      End
      Begin VB.Menu roleClearBtn 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu FaceMenu 
      Caption         =   "Face"
      Begin VB.Menu FaceBtn 
         Caption         =   "normal"
         Index           =   0
      End
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AddBtn_Click()
    MSSList.AddItem "", MSSList.ListIndex + 1
    MSSList.ListIndex = MSSList.ListIndex + 1
End Sub



Private Sub bgBtn_Click(index As Integer)
    SaveMark = False
    If AddMode Then
        MSSList.AddItem IIf(index = BGList.ListCount - 1, "bg\0,0,0", "*bg\" & BGList.List(index)), MSSList.ListIndex + 1
    Else
        MSSList.List(MSSList.ListIndex) = IIf(index = BGList.ListCount - 1, "bg\0,0,0", "*bg\" & BGList.List(index))
    End If
End Sub

Private Sub FaceBtn_Click(index As Integer)
    MSSList.AddItem "face\" & NowRole & "\" & FaceBtn(index).Caption, MSSList.ListIndex + 1
End Sub

Private Sub fgBtn_Click(index As Integer)
    SaveMark = False
    If AddMode Then
        MSSList.AddItem "*fg\" & FGList.List(index), MSSList.ListIndex + 1
    Else
        MSSList.List(MSSList.ListIndex) = "*fg\" & FGList.List(index)
    End If
End Sub

Private Sub bgmBtn_Click(index As Integer)
    SaveMark = False
    If AddMode Then
        MSSList.AddItem "*bgm\" & BGMList.List(index), MSSList.ListIndex + 1
    Else
        MSSList.List(MSSList.ListIndex) = "*bgm\" & BGMList.List(index)
    End If
End Sub
Private Sub bgsBtn_Click(index As Integer)
    SaveMark = False
    If AddMode Then
        MSSList.AddItem "*bgs\" & BGSList.List(index), MSSList.ListIndex + 1
    Else
        MSSList.List(MSSList.ListIndex) = "*bgs\" & BGSList.List(index)
    End If
End Sub

Private Sub removeroleBtn_Click(index As Integer)
    If index = 0 Then Exit Sub
    SaveMark = False
    MSSList.AddItem "role\" & Split(removeroleBtn(index).Caption, " - ")(0) & "\remove", MSSList.ListIndex + 1
End Sub

Private Sub roleaddBtn_Click(index As Integer)
    SaveMark = False
    MSSList.AddItem "role\" & Split(roleaddBtn(index).Caption, " - ")(0) & "\add", MSSList.ListIndex + 1
End Sub

Private Sub roleClearBtn_Click()
    SaveMark = False
    Dim I As Integer
    For I = 1 To removeroleBtn.UBound
        Call removeroleBtn_Click(I)
    Next
End Sub

Private Sub seBtn_Click(index As Integer)
    SaveMark = False
    If AddMode Then
        MSSList.AddItem "*play\" & SEList.List(index), MSSList.ListIndex + 1
    Else
        MSSList.List(MSSList.ListIndex) = "*play\" & SEList.List(index)
    End If
End Sub

Private Sub BGMLab_Click()
    'Music.Play BGMList.List(BGMList.ListIndex)
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
    
    Dim temp2() As String, Spliter As String
    Dim asideMode As Boolean
    asideMode = True
    
    For I = 0 To UBound(temp)
        If InStr(temp(I), "\") = 0 Then
            If Left(temp(I), 1) = "“" Or Left(temp(I), 1) = """" Then
                asideMode = True
            ElseIf asideMode Then
                MSSList.AddItem "say\aside"
                asideMode = False
            End If
        End If
        If Len(temp(I)) > 33 Then
            Spliter = VBA.InputBox("下列句子太长，请使用“|”分句。", "分句", temp(I))
            temp2 = Split(Spliter, "|")
            For S = 0 To UBound(temp2)
                MSSList.AddItem temp2(S)
            Next
        Else
            MSSList.AddItem temp(I)
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
    Dim I As Integer
    With MSSList
        Open App.Path & "\..\article\PART " & EditIndex & ".mss" For Output As #1
        For I = 0 To MSSList.ListCount - 1
            Print #1, MSSList.List(I)
        Next
        Close #1
    End With
    SaveMark = True
    'MsgBox "Save OK !!!"
End Sub

Private Sub Form_Load()
    'BASS_Init -1, 44100, BASS_DEVICE_3D, Me.Hwnd, 0
    'Set Music = New GMusicList

    
    DirInto BGMList, bgmBtn, App.Path & "\..\music\bgm\"
    DirInto BGSList, bgsBtn, App.Path & "\..\music\bgs\"
    DirInto SEList, seBtn, App.Path & "\..\music\se\"
    DirInto BGList, bgBtn, App.Path & "\..\assets\bg\"
    BGList.AddItem "RGB Color"
    DirInto FGList, fgBtn, App.Path & "\..\assets\fg\"
    
    Load bgBtn(bgBtn.UBound + 1)
    With bgBtn(bgBtn.UBound)
        .Caption = "RGB"
        .Visible = True
    End With
    
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
    
    
End Sub

Public Sub DirInto(Obj As ComboBox, MenuObj As Object, Folder As String)
    Dim File As String
    File = Dir(Folder)
    Obj.AddItem ""
    Do While File <> ""
        Obj.AddItem File
        Load MenuObj(MenuObj.UBound + 1)
        With MenuObj(MenuObj.UBound)
            .Caption = File
            .Visible = True
        End With
        File = Dir()
        DoEvents
    Loop
    Obj.ListIndex = 0
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

Private Sub MSSList_MouseUp(button As Integer, Shift As Integer, X As Single, y As Single)
    If button = 4 Then
        Call FastBtn_Click
    End If
End Sub

Private Sub roleBtn_Click(index As Integer)
    If index = 0 Then
        MSSList.List(MSSList.ListIndex) = "say\dark"
    ElseIf index = UBound(RoleList) + 1 Then
        MSSList.List(MSSList.ListIndex) = "say\" & VBA.InputBox("RoleName", "RoleEdit")
    Else
        MSSList.List(MSSList.ListIndex) = "say\" & Split(RoleList(index), " - ")(0)
    End If
End Sub

Private Sub saveposBtn_Click()
    If MSSList.ListIndex = -1 Then Exit Sub
    SaveMark = False
    MSSList.AddItem "*save\", MSSList.ListIndex + 1
    MSSList.ListIndex = MSSList.ListIndex + 1
End Sub

Private Sub SEList_Click()
    If MSSList.ListIndex = -1 Then Exit Sub
    SaveMark = False
    MSSList.AddItem "*play\" & SEList.List(SEList.ListIndex), MSSList.ListIndex + 1
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
