VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeatherLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   ҳ�������
    Dim Page As GPage
    Public Weather As String
    Private Type Pos
        X As Long
        y As Long
        Pic As String
    End Type
    Dim Snow() As Pos
    Public SetMode As Boolean, DMode As Boolean
'======================================
Public Sub SwitchDebug()
    DMode = Not DMode
    ECore.LockPage = IIf(DMode, "WeatherLayer", "")
    SE.Play "Cancel.mp3"
    If DMode Then ECore.NewTransform
End Sub
Public Sub SwitchSetting()
    SetMode = Not SetMode
    ECore.LockPage = IIf(SetMode, "WeatherLayer", "")
    SE.Play "Cancel.mp3"
    ECore.NewTransform
End Sub
Public Sub DrawSetting()
    Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(100, 0, 0, 0)
    
    Page.Paint 2, GW / 2, GH / 2, 600, 400, argb(255, 255, 255, 255), 32, Pos:=posOnCenter
    
    Page.DrawImage "seticon.png", GW / 2, GH / 2 - 200, alpha:=1, Pos:=posOnCenter
    
    Dim data As Variant
    data = BGM.Volume
    Page.Writes "��������", GW / 2 - 300 + 70, GH / 2 - 200 + 100, 18, argb(255, 27, 27, 27)
    Page.ShowSliderBar data, "progress.png", GW / 2 - 300 + 200, GH / 2 - 200 + 108
    If BGM.Volume <> data Then BGM.Volume = data: ESave.PutData "BGMVolume", data
    
    data = BGS.Volume
    Page.Writes "������Ч", GW / 2 - 300 + 70, GH / 2 - 200 + 150, 18, argb(255, 27, 27, 27)
    Page.ShowSliderBar data, "progress.png", GW / 2 - 300 + 200, GH / 2 - 200 + 158
    If BGS.Volume <> data Then BGS.Volume = data: ESave.PutData "BGSVolume", data
    
    data = CursorSndS
    Page.ShowColorCheckBox CursorSndS, GW / 2 - 300 + 70, GH / 2 - 200 + 208, 300, 20, "���������Ч", argb(255, 200, 200, 200), argb(255, 67, 238, 181), argb(255, 27, 27, 27), 18
    If data <> CursorSndS Then ESave.PutData "CursorSndS", CursorSndS

End Sub
Public Sub DrawDebug()
    Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(100, 0, 0, 0)
    
    Page.Paint 2, GW / 2, GH / 2, 600, 500, argb(255, 255, 255, 255), 32, Pos:=posOnCenter
    
    Dim Games(10) As String
    Games(0) = "maze"
    Games(1) = "battle"
    Games(2) = "chess"
    Games(3) = "bxbattle"
    Games(4) = "mazesnow"
    Games(5) = "snowman"
    Games(6) = "snowstorm"
    Games(7) = "novel"
    Games(8) = "fly"
    Games(9) = "mazecore"
    Games(10) = "mazelake"
    
    Dim m As Long
    
    For i = 0 To UBound(Games)
        Page.Paint 0, GW / 2 - 250, GH / 2 - 250 + 32 + i * 34, 500, 35, IIf(i Mod 2 = 0, argb(255, 240, 240, 240), argb(255, 255, 255, 255))
        m = CheckMouse2
        If m = mMouseUp Then
            TipPage.ShowTip Games(i)
            ECore.NewTransform transDarkTo, , "TipPage"
            Call SwitchDebug
        End If
        Page.Writes Games(i), GW / 2 - 250 + 10, GH / 2 - 250 + 32 + 4 + i * 34, 18, argb(255, 27, 27, 27)
        If m > 0 Then Page.Paint 0, GW / 2 - 250, GH / 2 - 250 + 32 + i * 34, 500, 35, argb(30, 0, 176, 240)
    Next
End Sub
Public Sub Update()
    '��Ϸÿһ֡�Ĺ��̣����ƹ�������ڴ˴���
    
    Page.Clear 0 '��ջ���

    If SetMode Then Call DrawSetting
    If DMode Then Call DrawDebug
    
    Select Case Weather
        Case ""
        Case "snow": Call DrawSnow
        Case "snowstorm": Call DrawSnowStorm
        Case "dark": Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(70, 10, 0, 0)
    End Select
    
End Sub
Public Sub RandomSnow(i As Integer)
    With Snow(i)
        .X = Int(Rnd * GW)
        .y = -90 - Int(Rnd * GH)
        .Pic = IIf(Int(Rnd * 2) = 0, "bigsnow.png", "smallsnow.png")
    End With
End Sub
Public Sub ChangeWeather(nWeather As String)
    Dim i As Integer
    Weather = nWeather
    Select Case Weather
        Case "snow"
SnowSummer:
            ReDim Snow(30)
            For i = 0 To 30
                Call RandomSnow(i)
            Next
        Case "snowstorm"
            GoTo SnowSummer
    End Select
End Sub
Public Sub DrawSnow()
    Dim i As Integer
    For i = 0 To UBound(Snow)
        If Snow(i).y > 0 Then Page.DrawImage Snow(i).Pic, Snow(i).X, Snow(i).y, alpha:=1, Pos:=posOnCenter
        Snow(i).y = Snow(i).y + 20 * IIf(i Mod 2 = 0, 1.5, 1)
        If Snow(i).y > GH Then Call RandomSnow(i)
    Next
End Sub
Public Sub DrawSnowStorm()
    'Page.DrawImage "snowscreen.png", 0, 0, alpha:=1
    Dim i As Integer
    For i = 0 To UBound(Snow)
        If Snow(i).y > 0 Then
            Page.DrawImage Snow(i).Pic, Snow(i).X, Snow(i).y, alpha:=1, Pos:=posOnCenter
        End If
        Snow(i).y = Snow(i).y + 25 * IIf(i Mod 2 = 0, 1.5, 1)
        Snow(i).X = Snow(i).X + 10 * IIf(i Mod 2 = 0, 1, 2)
        If Snow(i).y > GH Then Call RandomSnow(i): Snow(i).X = Int(Rnd * GW * 2) - GW
    Next
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '�������¼�
    'Direction:����Depth�����
End Sub
Public Sub AnimationMsg(ID As String, msg As String)
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
Public Sub AnimationDone(ID As Integer)
    '�������Ž����¼�
    'id�������Ķ���ID
End Sub
Private Sub Class_Initialize()
    '����ҳ�������
    Set Page = New GPage
    Page.Create Me
    '������Ϸ��Դ
    Page.TopPage = True
    Page.Res.NewImages App.Path & "\assets\weather", 0.8
    Page.Res.NewImages App.Path & "\assets\setui"
    '����ҳ��
    ECore.Add Page, "WeatherLayer"
    If ESave.GetData("BGMVolume") <> "" Then BGM.Volume = Val(ESave.GetData("BGMVolume"))
    If ESave.GetData("BGSVolume") <> "" Then BGS.Volume = Val(ESave.GetData("BGSVolume"))
    If ESave.GetData("CursorSndS") <> "" Then CursorSndS = Val(ESave.GetData("CursorSndS"))
    
    '===============================================
    '   �����Ҫ���Ӷ�����������ڴ˴�
    
    '===============================================
End Sub