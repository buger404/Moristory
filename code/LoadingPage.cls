VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadingPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   ҳ�������
    Dim Page As GPage
'======================================

Public Sub Update()
    '��Ϸÿһ֡�Ĺ��̣����ƹ�������ڴ˴���
    
    Page.Clear '��ջ���

    Page.DrawImage "logo.png", GW / 2, GH / 2 - 50, alpha:=1, Pos:=posOnCenter
    
    Dim AText As String
    For i = 1 To 2000 / 200 - 1
        If GetTickCount Mod 2000 > i * 200 Then AText = AText & "��"
    Next
    
    If GetTickCount Mod 4000 <= 2000 Then
        AText = "(�s�F����)�s" & AText & "�ߩ���"
    Else
        AText = "�С���" & AText & "��('-'��)"
    End If
    
    Page.Writes AText & vbCrLf & vbCrLf & _
                "���ڼ��ص� " & LoadedCount & " ����Դ (" & Int(CrashPro * 100) & "%)", _
                GW / 2, GH / 2 + 50, 18, argb(255, 64, 64, 64), align:=StringAlignmentCenter
                
    Page.Writes Replace(GetCountStr, vbCrLf, " , "), _
                GW / 2, GH - 50, 18, argb(255, 128, 128, 128), align:=StringAlignmentCenter
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
    Page.Res.newImage App.Path & "\assets\logo.png", 0.7
    '����ҳ��
    ECore.Add Page, "LoadingPage"
    
    '===============================================
    '   �����Ҫ���Ӷ�����������ڴ˴�
    
    '===============================================
End Sub