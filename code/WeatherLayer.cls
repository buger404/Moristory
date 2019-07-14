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
'   页面绘制器
    Dim Page As GPage
    Public Weather As String
    Private Type Pos
        X As Long
        Y As Long
        Pic As String
    End Type
    Dim Snow() As Pos
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear 0 '清空画布

    Select Case Weather
        Case ""
        Case "snow": Call DrawSnow
        Case "snowstorm": Call DrawSnowStorm
    End Select
    
End Sub
Public Sub RandomSnow(i As Integer)
    With Snow(i)
        .X = Int(Rnd * GW)
        .Y = -90 - Int(Rnd * GH)
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
        If Snow(i).Y > 0 Then Page.DrawImage Snow(i).Pic, Snow(i).X, Snow(i).Y, alpha:=1, Pos:=posOnCenter
        Snow(i).Y = Snow(i).Y + 20 * IIf(i Mod 2 = 0, 1.5, 1)
        If Snow(i).Y > GH Then Call RandomSnow(i)
    Next
End Sub
Public Sub DrawSnowStorm()
    Page.DrawImage "snowscreen.png", 0, 0, alpha:=1
    Dim i As Integer
    For i = 0 To UBound(Snow)
        If Snow(i).Y > 0 Then
            Page.DrawImage Snow(i).Pic, Snow(i).X, Snow(i).Y, alpha:=1, Pos:=posOnCenter
        End If
        Snow(i).Y = Snow(i).Y + 25 * IIf(i Mod 2 = 0, 1.5, 1)
        Snow(i).X = Snow(i).X + 10 * IIf(i Mod 2 = 0, 1, 2)
        If Snow(i).Y > GH Then Call RandomSnow(i): Snow(i).X = Int(Rnd * GW * 2) - GW
    Next
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
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
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.TopPage = True
    Page.Res.NewImages App.Path & "\assets\weather", 0.8
    Page.Res.newImage App.Path & "\assets\ui\snowscreen.png", GW, GH
    '创建页面
    ECore.Add Page, "WeatherLayer"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
