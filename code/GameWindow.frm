VERSION 5.00
Begin VB.Form GameWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Moristory DEMO"
   ClientHeight    =   6672
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   805
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   9024
      Top             =   264
   End
End
Attribute VB_Name = "GameWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================
'   ҳ�������
    Dim EC As GMan
'==================================================
'   �ڴ˴��������ҳ����ģ������

'==================================================

Private Sub DrawTimer_Timer()
    '����
    EC.Display
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '�����ַ�����
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
End Sub

Private Sub Form_Load()
    '��ʼ��Emerald���ڴ˴������޸Ĵ��ڴ�СӴ~��
    StartEmerald Me.Hwnd, 1000, 740
    '��������
    MakeFont "΢���ź�"
    '����ҳ�������
    Set EC = New GMan
    
    '�����浵����ѡ��
    Set ESave = New GSaving
    ESave.Create "Moristory.TIMELINE", "kj" & Val(Me.Visible) & "Ehsd" & Val(VB.Screen.FontCount <> 0) & "Cdfd" & Right(Left("54B89", 3), 1) & "3fdkg5" & UCase("d") & "gsA6D1F7305BEjAC8738C" & CLng("&HE2") & "kjgds"
    ESave.PutData "PART", "2"
    ESave.PutData "TIMELINE", "1"
    
    '���������б�
    Set MusicList = New GMusicList
    MusicList.Create App.path & "\music"

    '��ʼ��ʾ
    Me.Show
    DrawTimer.Enabled = True
    
    Set BGM = New GMusic
    Set BGS = New GMusic
    Set SE = New GMusicList
    SE.Create App.path & "\music\se"
    BGS.Volume = 0.3
    
    '�ڴ˴���ʼ�����ҳ��
    '=============================================
    'ʾ����TestPage.cls
    '     Set TestPage = New TestPage
    '�������֣�Dim TestPage As TestPage
        Set MainPage = New MainPage
        Set NovelPage = New NovelPage
        Set MazePage = New MazePage
    '=============================================

    '���ûҳ��
    EC.ActivePage = "MainPage"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '���������Ϣ
    UpdateMouse X, Y, 1, Button
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '���������Ϣ
    If Mouse.state = 0 Then
        UpdateMouse X, Y, 0, Button
    Else
        Mouse.X = X: Mouse.Y = Y
    End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '���������Ϣ
    UpdateMouse X, Y, 2, Button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set NovelPage = Nothing
    Me.Hide
    '��ֹ����
    DrawTimer.Enabled = False
    '�ͷ�Emerald��Դ
    EndEmerald
End Sub
