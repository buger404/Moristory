Attribute VB_Name = "MoriCore"
Public BGM As GMusic, BGS As GMusic
Public SE As GMusicList

Public MainPage As MainPage
Public NovelPage As NovelPage
Public MazePage As MazePage
Public BattlePage As BattlePage
Public TicTacToePage As TicTacToePage
Public BXBattlePage As BXBattlePage
Public SnowmanPage As SnowmanPage
Public WeatherLayer As WeatherLayer
Public TipPage As TipPage
Public FlyPage As FlyPage
Public TLPPage As TLPPage
Public DancePage As DancePage
Public FinalPage As FinalPage
Public EndingPage As EndingPage
Public ErrorPage As ErrorPage
Public LoadingPage As LoadingPage

Public CrashPro As Single
Public LastPage As String

Public Function GetPartTitle(Part As String) As String
    Dim temp As String
    Open App.Path & "\article\PART " & Part & ".mss" For Input As #1
    Do While Not EOF(1)
        Line Input #1, temp
        If temp Like "*title\*" Then GetPartTitle = Split(temp, "\")(1): Close #1: Exit Function
    Loop
    Close #1
End Function

Public Sub ErrCrash(Num As Long, Str As String)
    If GetTickCount - ErrorPage.IgnoreTime <= 5000 Then Exit Sub
    If ECore.ActivePage = "ErrorPage" Then Exit Sub
    
    Dim NewStr As String
    NewStr = Str
    Select Case Num
        Case 0: NewStr = "��������һ�£������Ĵ���"
        Case 5: NewStr = "������404�����ǰѷ����Ĵ���ɾ�ɾ��ˡ�"
        Case 6: NewStr = "shit 404Ϊ��������֭��ʱ��С��������ˡ�"
        Case 7: NewStr = "404û��������๤��ְ��"
        Case 9: NewStr = "404����������ı�Ե��̽...Ŷ�����������ˡ�"
        Case 11: NewStr = "�Բ�404��ѧ�����أ�����0ʲô��..."
        Case 13: NewStr = "404�ղ��ڸ����˽��ܶ����ʱ��˫�����Ҹ�����һ���ơ�"
        Case 28: NewStr = "�չ���404һ������������̫��Ĺ����ˣ�"
        Case 35: NewStr = "������404�����ǰѷ����Ĵ���ɾ�ɾ��ˡ�"
        Case 52: NewStr = "404�����˷�����롣"
        Case 53: NewStr = "404��GPS���˵�����..."
        Case 55: NewStr = "�������ĵ�404��һ���ļ������˹��ϡ�"
        Case 58: NewStr = "404˵������ļ��Ѿ������ˣ���TNTը��ô��"
        Case 70: NewStr = "404�����ڵײ������ʵ��û��Ȩ����ɵĹ�����"
        Case 75: NewStr = "404������һ������ĵ�ַ��"
        Case 76: NewStr = "404���������Ѿ���Ǩ�Ĺ�Ԣ��"
    End Select
    
    LastPage = ECore.ActivePage
    If Num = 404233 Then
        ErrorPage.ErrText = "��Ϸû�����⣬ֻ�Ǳ�����ֹ��" & vbCrLf & NewStr & vbCrLf & "�����ʾ �� �ǳ���Ǹ�����������ǸղŰ�����ʲô��"
    Else
        ErrorPage.ErrText = IIf(ECore.ActivePage = "LoadingPage", "��Ϸ�ƺ�����û����������(" & Num & ")��", "������<" & ECore.ActivePage & ">����ˣ��ʱ����Ϸ�ڲ�������һЩ����(" & Num & ")��") & vbCrLf & NewStr & vbCrLf & "�����ʾ�ǳ���Ǹ�������Գ������¿�����Ϸ���������⡣"
    End If
    
    If Not WeatherLayer Is Nothing Then WeatherLayer.Page.TopPage = False
    ECore.ActivePage = "ErrorPage"
End Sub
