VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ErrorPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   ҳ�������
    Dim Page As GPage
    Public ErrText As String
    Public IgnoreTime As Long
'======================================

Public Sub Update()
    '��Ϸÿһ֡�Ĺ��̣����ƹ�������ڴ˴���
    
    Page.Clear argb(255, 51, 133, 255) '��ջ���

    For i = 1 To 10
        Page.DrawImage "bm.png", GW - 200 + Cos(GetTickCount / 200 + i / 10 * 5) * 50, 600 - Sin(GetTickCount / 200 + i / 10 * 5) * 50, alpha:=i / 10, Pos:=posOnCenter
    Next
    
    Dim AText As String
    For i = 1 To 2000 / 200 - 1
        If GetTickCount Mod 2000 > i * 200 Then AText = AText & "��"
    Next
    
    If GetTickCount Mod 4000 <= 2000 Then
        Page.Writes "(�s�F����)�s" & AText & "�ߩ���", 100, 100, 36, argb(255, 255, 255, 255), style:=FontStyleBold
    Else
        Page.Writes "�С���" & AText & "��('-'��)", 100, 100, 36, argb(255, 255, 255, 255), style:=FontStyleBold
    End If
    
    Page.Writes ErrText, 100, 240, 18, argb(255, 255, 255, 255)

    If InStr(ErrText, "����") <> 1 Then
        Page.Writes "�볢��ͨ�����·�ʽ���ǣ�" & vbCrLf & vbCrLf & "�� �Ͽ���Դ" & vbCrLf & "�� ʹ�Լ����Բ��ܵ�����ײ��" & vbCrLf & "�� ������Githubҳ�����ύIssue" & vbCrLf & "�� ˳����������404����404��������", 100, 370, 18, argb(255, 255, 255, 255)
    Else
        Page.Writes "�볢��ͨ�����·�ʽ�������⣺" & vbCrLf & vbCrLf & "�� ��ϵQQ1361778219" & vbCrLf & "�� �����ʼ���ris_vb@126.com" & vbCrLf & "�� ��Githubҳ�����ύIssue" & vbCrLf & "�� ˳����������404��", 100, 370, 18, argb(255, 255, 255, 255)
    End If
    
    Dim m As Integer
    m = CheckMouse(100, 580, 100, 20)
    Page.Writes "[ ǿ�Ƽ��� ]", 100, 580, 18, IIf(m, argb(255, 255, 255, 255), argb(200, 255, 255, 255)), style:=FontStyleUnderline
    If m = 3 Then
        If Not WeatherLayer Is Nothing Then WeatherLayer.Page.TopPage = True
        ECore.NewTransform transDarkTo, 1000, IIf(LastPage = "", "MainPage", LastPage)
        IgnoreTime = GetTickCount
    End If
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '�������¼�
    'Direction:����Depth�����
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '������Ϣ����
End Sub
Public Sub Enter()
    'ҳ������¼�
End Sub
Public Sub Leave()
    'ҳ���뿪�¼�
End Sub
Public Sub Gotfocus()
    '����ȡ�ý����¼�
End Sub
Public Sub Lostfocus()
    '����ʧȥ�����¼�
End Sub
Public Sub AnimationDone(id As Integer)
    '�������Ž����¼�
    'id�������Ķ���ID
End Sub
Private Sub Class_Initialize()
    '����ҳ�������
    Set Page = New GPage
    Page.Create Me
    '������Ϸ��Դ
    Page.Res.NewImages App.Path & "\assets\face", 0.7
    '����ҳ��
    ECore.Add Page, "ErrorPage"
    
    '===============================================
    '   �����Ҫ���Ӷ�����������ڴ˴�
    
    '===============================================
End Sub