VERSION 5.00
Begin VB.Form EditWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Moristory Script Editor"
   ClientHeight    =   7560
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   11928
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   630
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   994
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer BackupTimer 
      Interval        =   60000
      Left            =   144
      Top             =   144
   End
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   11424
      Top             =   216
   End
End
Attribute VB_Name = "EditWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================
'   页面管理器
    Dim EC As GMan
'==================================================
'   在此处放置你的页面类模块声明
    Dim MainPage As MainPage
'==================================================

Private Sub DrawTimer_Timer()
    '绘制
    EC.Display
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '发送字符输入
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
End Sub
Private Sub BackupTimer_Timer()
    Call MakeBackup
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift Then
        MSSList.AddItem "", MSSList.ListIndex + 1
    End If
    If KeyCode = vbKeyD And Shift Then
        MSSList.RemoveItem MSSList.ListIndex + 1
    End If
End Sub

Private Sub Form_Load()
    '初始化Emerald（在此处可以修改窗口大小哟~）
    StartEmerald Me.Hwnd, 900, 720
    '创建字体
    MakeFont "微软雅黑"
    
    '创建页面管理器
    Set EC = New GMan
    HideLOGO = 1
    DisableLOGO = 1
    '创建存档（可选）
    'Set ESave = New GSaving
    'ESave.Create "emerald.test", "Emerald.test"
    
    '创建音乐列表
    Set Music = New GMusic

    '开始显示
    Me.Show
    DrawTimer.Enabled = True
    
    '在此处初始化你的页面
    '=============================================
    '示例：TestPage.cls
    '     Set TestPage = New TestPage
    '公共部分：Dim TestPage As TestPage
        Set MainPage = New MainPage
    '=============================================

    '设置活动页面
    EC.ActivePage = "MainPage"
    
    MSSList.ListIndex = -1
    
    MainWindow.Hide
    
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
        Call MakeBackup
        With MSSList
            .Clear
            Open App.Path & "\..\article\PART " & EditIndex & ".mss" For Input As #1
            Do While Not EOF(1)
                Line Input #1, temp
                .AddItem temp
                'If temp Like "*title\*" Then InputBox.Text = Split(temp, "\")(1)
            Loop
            Close #1
            SaveMark = True
        End With
    End If

End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, y As Single)
    '发送鼠标信息
    UpdateMouse X, y, 1, button
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, y As Single)
    '发送鼠标信息
    If Mouse.state = 0 Then
        UpdateMouse X, y, 0, button
    Else
        Mouse.X = X: Mouse.y = y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, y As Single)
    '发送鼠标信息
    UpdateMouse X, y, 2, button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If SaveMark = False Then Cancel = 1: MsgBox "你还没保存。", 16: Exit Sub
    '终止绘制
    DrawTimer.Enabled = False
    '释放Emerald资源
    Unload MainWindow
    EndEmerald
End Sub
