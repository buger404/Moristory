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
   StartUpPosition =   2  '��Ļ����
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
'   ҳ�������
    Dim EC As GMan
'==================================================
'   �ڴ˴��������ҳ����ģ������
    Dim MainPage As MainPage
'==================================================

Private Sub DrawTimer_Timer()
    '����
    EC.Display
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '�����ַ�����
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
    '��ʼ��Emerald���ڴ˴������޸Ĵ��ڴ�СӴ~��
    StartEmerald Me.Hwnd, 900, 720
    '��������
    MakeFont "΢���ź�"
    
    '����ҳ�������
    Set EC = New GMan
    HideLOGO = 1
    DisableLOGO = 1
    '�����浵����ѡ��
    'Set ESave = New GSaving
    'ESave.Create "emerald.test", "Emerald.test"
    
    '���������б�
    Set Music = New GMusic

    '��ʼ��ʾ
    Me.Show
    DrawTimer.Enabled = True
    
    '�ڴ˴���ʼ�����ҳ��
    '=============================================
    'ʾ����TestPage.cls
    '     Set TestPage = New TestPage
    '�������֣�Dim TestPage As TestPage
        Set MainPage = New MainPage
    '=============================================

    '���ûҳ��
    EC.ActivePage = "MainPage"
    
    MSSList.ListIndex = -1
    
    MainWindow.Hide
    
    MSSList.AddItem "#�ⲻ����Ӧ�ÿ����ļ���"
    MSSList.AddItem "*title\"
    MSSList.AddItem "*mode\msg"
    
    MSSList.ListIndex = 2
    
    Dim temp As String
    
    If Dir(App.Path & "\backup", vbDirectory) = "" Then MkDir App.Path & "\backup"
    
    EditIndex = VBA.InputBox("PART����", "��", "0")
    Me.Caption = "Moristory Script Editor - PART " & EditIndex
    If Dir(App.Path & "\..\article\PART " & EditIndex & ".mss") <> "" Then
        If MsgBox("���PART�Ѿ������ˣ��㲻�Ḳ�ǵ����Բ��ԣ�", vbYesNo) = vbNo Then End
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
    '���������Ϣ
    UpdateMouse X, y, 1, button
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, y As Single)
    '���������Ϣ
    If Mouse.state = 0 Then
        UpdateMouse X, y, 0, button
    Else
        Mouse.X = X: Mouse.y = y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, y As Single)
    '���������Ϣ
    UpdateMouse X, y, 2, button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If SaveMark = False Then Cancel = 1: MsgBox "�㻹û���档", 16: Exit Sub
    '��ֹ����
    DrawTimer.Enabled = False
    '�ͷ�Emerald��Դ
    Unload MainWindow
    EndEmerald
End Sub
