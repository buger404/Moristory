VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndingPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim Dir1 As Long, Dir2 As Long, DirTime As Long
    Dim EndMark As Boolean, DaTime As Long
'======================================

Public Sub Update()
    If Err.number <> 0 Then ErrCrash Err.number, Err.Description
    On Error Resume Next

    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear argb(255, 128, 128, 128) '清空画布
    
    If BGM.IsPlaying = False Then
        If EndMark = False Then
            EndMark = True
            NovelPage.RunNovel
            ECore.NewTransform transFallDark, 2000, "NovelPage"
        Else
            Page.Clear argb(255, 0, 0, 0) '清空画布
            Select Case GetTickCount - DaTime
                Case Is <= 10000: Page.Writes "黑嘴：别发呆了，结束了。", GW / 2, 350, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter
                Case Is <= 20000: Page.Writes "黑嘴：？？？你咋还不走？", GW / 2, 350, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter
                Case Is <= 30000: Page.Writes "黑嘴：真的没有东西喽？", GW / 2, 350, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter
                Case Is <= 40000: Page.Writes "黑嘴：哼，不理你啦，我去睡觉啦。", GW / 2, 350, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter
                Case Else
                    Dim DPro As Single
                    DPro = (GetTickCount - DaTime - 40000) / 60000
                    If DPro = 1 Then Unload GameWindow: End
                    SetLayeredWindowAttributes GHwnd, 0, (1 - DPro) * 255, LWA_ALPHA
                    Page.Writes "(|3[▓" & IIf(GetTickCount Mod 2000 < 500, "▓", "") & IIf(GetTickCount Mod 2000 < 1000, "▓", "") & IIf(GetTickCount Mod 2000 < 1500, "▓", "") & "▓] ", GW / 2, 350, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter
            End Select
            
            If DaTime = 0 Then
                SetWindowLongA GHwnd, GWL_EXSTYLE, GetWindowLongA(GHwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
                DaTime = GetTickCount
            End If
            Exit Sub
        End If
        Exit Sub
    End If
    
    If EndMark Then
        Page.Clear argb(255, 0, 0, 0) '清空画布
        Page.DrawImage "logo.png", GW / 2, 250, alpha:=1, Pos:=posOnCenter
        Page.Writes "- THE END -", GW / 2, 350, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter
        Page.DrawImage "bm.png", GW / 2, 650 - IIf(GetTickCount Mod 200 < 100, 10, 0), alpha:=1, Pos:=posOnCenter
        Exit Sub
    End If
    
    Page.DrawImage "garden.jpg", 0, 0
    Page.DrawImage "Port.png", 0, 0, alpha:=1
    
    'If Mouse.state = 2 Then Mouse.state = 0: BGM.Play

    If GetTickCount - DirTime >= 1000 / 8 Then
        Dir1 = Int(Rnd * 4): Dir2 = Int(Rnd * 4)
        If Dir1 > 3 Then Dir1 = 3
        If Dir2 > 3 Then Dir2 = 3
        DirTime = GetTickCount
    End If
    
    Dim Dir1S As String, Dir2S As String
    
    Select Case Dir1
        Case 0: Dir1S = "left"
        Case 1: Dir1S = "right"
        Case 2: Dir1S = "up"
        Case 3: Dir1S = "down"
    End Select
    
    Select Case Dir2
        Case 0: Dir2S = "left"
        Case 1: Dir2S = "right"
        Case 2: Dir2S = "up"
        Case 3: Dir2S = "down"
    End Select
    
    
    If BGM.position > 120 Then
        Dim pro3 As Single
        pro3 = (BGM.position - 120) / 2
        If pro3 > 1 Then pro3 = 1
        For i = 1 To 2
            Page.DrawImage "bm.png", GW / 2 - 350 + 700 * (i - 1), -900 + 900 * pro3 + 300 - IIf(GetTickCount Mod 200 < 100, 10, 0), alpha:=1, Pos:=posOnCenter
        Next
    End If
    
    Page.DrawImage "yy-dance.png", GW / 2, GH / 2 - 100, alpha:=1, Pos:=posOnCenter
    Page.DrawImage "yy-" & Dir1S & "1.png", GW / 2, GH / 2 - 100, alpha:=1, Pos:=posOnCenter
    Page.DrawImage "yy-" & Dir2S & "2.png", GW / 2, GH / 2 - 100, alpha:=1, Pos:=posOnCenter
    
    'Page.Writes BGM.position, 20, 20, 18, argb(255, 255, 0, 0)
    
    If BGM.position > 20 Then
        Dim pro1 As Single
        pro1 = (BGM.position - 20) / 3
        If pro1 > 1 Then pro1 = 1
        Dim map() As Single, BarW As Long
        BarW = Int(GW / 21)
        map = BGM.GetMusicMap
        For i = 0 To 20
            Page.Paint 0, i * (BarW), GH, BarW + 1, map(i) / 5000 * GH * pro1, argb(128, 255, 255, 255), Pos:=posOnBottom
        Next
    End If
    
    
    If BGM.position > 5.5 Then
        pro2 = (BGM.position - 5.5) / 1
        If pro2 > 1 Then pro2 = 1
        For i = 0 To 10
            Page.DrawImage "tk-normal.png", (GW / 2 + Cos((GetTickCount / 250) + i / 5) * 50 - 250) * pro2, GH / 2 + Sin((GetTickCount / 200) + i / 10) * 50, alpha:=Cubic(i / 10, 0, 0, 0, 1) * pro2, Pos:=posOnCenter
        Next
    End If
    
    If BGM.position > 85 Then
        pro2 = (BGM.position - 85) / 1
        If pro2 > 1 Then pro2 = 1
        Page.DrawImage "yz-" & IIf(GetTickCount Mod 1000 < 500, "se", "normal") & ".png", GW / 2 - 180, GH + 700 - 850 * pro2, alpha:=1, Pos:=posOnCenter
    End If
    
    If BGM.position > 95 Then
        pro2 = (BGM.position - 95) / 1
        If pro2 > 1 Then pro2 = 1
        Page.DrawImage "ssr-normal.png", -1000 + 1000 * pro2 + GW / 2 - 400, 500, alpha:=1, Pos:=posOnCenter
        Page.DrawImage "km1-normal.png", -1000 + 1000 * pro2 + GW / 2 - 100, 500, alpha:=1, Pos:=posOnCenter
    End If
    
    If BGM.position > 80 Then
        pro2 = (BGM.position - 80) / 1
        If pro2 > 1 Then pro2 = 1
        Page.DrawImage "dn-normal.png", -1000 + 1000 * pro2 + GW / 2 - 250, 150, alpha:=Abs(Cos((GetTickCount / 300))), Pos:=posOnCenter
    End If
    
    If BGM.position > 70 Then
        pro2 = (BGM.position - 70) / 1
        If pro2 > 1 Then pro2 = 1
        Page.DrawImage "kx1-normal.png", GW / 2 + 350, -1600 + 1600 * pro2 + 400 + Sin((GetTickCount / 400) + 1) * 50 + 2 - 2, alpha:=1, Pos:=posOnCenter
        For i = 0 To 10
            Page.DrawImage "kx1-normal.png", GW / 2 + 350, -1600 + 1600 * pro2 + 400 + Sin((GetTickCount / 400) + i / 10) * 50, alpha:=0.1, Pos:=posOnCenter
        Next
        Page.DrawImage "kx2-normal.png", GW / 2 + 250, -1600 + 1600 * pro2 + 400 + Cos((GetTickCount / 400) + 1) * 50 + 2 - 2, alpha:=1, Pos:=posOnCenter
    End If
    
    If BGM.position > 55 Then
        pro2 = (BGM.position - 55) / 1
        If pro2 > 1 Then pro2 = 1
        Page.DrawImage "fj-happy.png", GW - 200 + Cos((GetTickCount / 250) + i / 5) * 50, -300 + 500 * pro2, alpha:=1, Pos:=posOnCenter
    End If
    
    If BGM.position > 41 Then
        pro2 = (BGM.position - 41) / 1
        If pro2 > 1 Then pro2 = 1
        Page.DrawImage "jy-smile.png", GW - 350 * pro2, GH - 200 - IIf(GetTickCount Mod 200 < 100, 10, 0), alpha:=1, Pos:=posOnCenter
    End If
    
    If BGM.position > 11 Then
        pro2 = (BGM.position - 11) / 1
        If pro2 > 1 Then pro2 = 1
        Page.DrawImage "bm-normal.png", GW - 350, GH + 800 - 800 * pro2 - IIf(GetTickCount Mod 200 < 100, 10, 0), alpha:=1, Pos:=posOnCenter
        Page.DrawImage "xx-love.png", GW - 120, GH + 800 - 800 * pro2 - IIf(GetTickCount Mod 200 < 100, 10, 0), alpha:=1, Pos:=posOnCenter
    End If
    
    If BGM.position > 32 Then
        pro2 = (BGM.position - 32) / 2
        If pro2 > 1 Then pro2 = 1
        If GetTickCount Mod 10000 < 1000 Then
            Page.DrawImage "xl-normal.png", -900 + 1100 * pro2, GH - 100 - IIf(GetTickCount Mod 200 < 100, 10, 0), alpha:=1, Pos:=posOnCenter
        Else
            Page.DrawImage "xl-happy" & IIf(GetTickCount Mod 1000 < 500, "f", "") & ".png", -900 + 1100 * pro2, GH - 100 - IIf(GetTickCount Mod 200 < 100, 10, 0), alpha:=1, Pos:=posOnCenter
        End If
        Page.DrawImage "bg-normal.png", -900 + 1100 * pro2, GH - 250 - IIf(GetTickCount Mod 200 < 100, 10, 0), alpha:=1, Pos:=posOnCenter
        Page.DrawImage "s-normal.png", -900 + 1450 * pro2, GH - 40 - IIf(GetTickCount Mod 200 < 100, 10, 0), alpha:=1, Pos:=posOnCenter

        Page.DrawImage "qing-normal.png", -900 + 1350 * pro2, GH + 20 - IIf(GetTickCount Mod 200 < 100, 10, 0), alpha:=1, Pos:=posOnCenter
    End If
    
    If BGM.position < 4 Then
        Page.Writes "制作： Error 404", GW / 2 + 20, GH / 2 + 220, 24, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    End If
    
    If BGM.position < 9 And BGM.position > 5 Then
        Page.Writes "剧本： Error 404" & vbCrLf & "顾问&修改： Wey", GW / 2 + 20, GH / 2 + 180, 24, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    End If
    
    If BGM.position > 10 And BGM.position < 14 Then
        Page.Writes "测试（按首字的拼音排序）" & vbCrLf & "冰棍" & vbCrLf & "南笙" & vbCrLf & "小黄", GW / 2 + 20, GH / 2 + 140, 24, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    End If
    
    If BGM.position > 20 And BGM.position < 30 Then
        Page.Writes "结尾曲" & vbCrLf & "< SWEET CANDY >" & vbCrLf & "专辑：ivy ～アイビー～" & vbCrLf & "歌手：桃井はるこ", GW / 2 + 20, GH / 2 + 140, 24, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    End If
    
    If BGM.position > 31 And BGM.position < 40 Then
        Page.Writes "角色设计" & vbCrLf & "黑嘴 - Error 404 / 小婷 / 耗子" & vbCrLf & "雪狼 - 小婷" & vbCrLf & "其他 - Error 404", 80, 100, 24, argb(255, 255, 255, 255)
    End If
    
    If BGM.position > 41 And BGM.position < 45 Then
        Page.Writes "Soundtrack" & vbCrLf & "(1) Lonely Rainbow by Nian" & vbCrLf & "(2) Battle by SGB" & vbCrLf & "(3) Sweet Flower by しょうゆ", GW - 50, 100, 16, argb(255, 255, 255, 255), align:=StringAlignmentFar
    End If
    
    If BGM.position > 45 And BGM.position < 50 Then
        Page.Writes "Soundtrack" & vbCrLf & "(4) Garden Party by Jazz Piano Bar Academy" & vbCrLf & "(5) Firecrackers by F-777" & vbCrLf & "(6) Mysterious Forest by dave shtorn", GW - 50, 100, 16, argb(255, 255, 255, 255), align:=StringAlignmentFar
    End If
    
    If BGM.position > 50 And BGM.position < 55 Then
        Page.Writes "Soundtrack" & vbCrLf & "(7) Knife by Flannel." & vbCrLf & "(8) Kindness by ahiru" & vbCrLf & "(9) Garden Party by 光宗信吉", GW - 50, 100, 16, argb(255, 255, 255, 255), align:=StringAlignmentFar
    End If
    
    If BGM.position > 55 And BGM.position < 60 Then
        Page.Writes "Soundtrack" & vbCrLf & "(10) Sweet Candy by Uno" & vbCrLf & "(11) Heartbeat by SGB" & vbCrLf & "(12) Comedy by MARiA", GW - 50, 100, 16, argb(255, 255, 255, 255), align:=StringAlignmentFar
    End If
    
    If BGM.position > 60 And BGM.position < 65 Then
        Page.Writes "Soundtrack" & vbCrLf & "(13) Summer by LJY" & vbCrLf & "(14) Snowy by DJ PotAtO" & vbCrLf & "(15) Milk Tea by DJ PotAtO", GW - 50, 100, 16, argb(255, 255, 255, 255), align:=StringAlignmentFar
    End If
    
    If BGM.position > 65 And BGM.position < 70 Then
        Page.Writes "Soundtrack" & vbCrLf & "(16) Snowstorm -euphoria- by V.A." & vbCrLf & "(17) Good Day with Blues by Jazz Piano Bar Academy" & vbCrLf & "(18) Hellfire by SGB", GW - 50, 100, 16, argb(255, 255, 255, 255), align:=StringAlignmentFar
    End If
    
    If BGM.position > 70 And BGM.position < 75 Then
        Page.Writes "Soundtrack" & vbCrLf & "(19) Deep in Your Eyes by Jazz Piano Bar Academy" & vbCrLf & "(20) The World Needs You Now by Steve Jablonsky" & vbCrLf & "(21) Shockwave's Revenge by Steve Jablonsky", GW - 50, 100, 16, argb(255, 255, 255, 255), align:=StringAlignmentFar
    End If
    
    If BGM.position > 75 And BGM.position < 80 Then
        Page.Writes "Soundtrack" & vbCrLf & "(22) Wind by SGB" & vbCrLf & "(23) A Lost World by F-777" & vbCrLf & "(24) Across The Waters by F-777", GW - 50, 100, 16, argb(255, 255, 255, 255), align:=StringAlignmentFar
    End If
    
    If BGM.position > 80 And BGM.position < 85 Then
        Page.Writes "Soundtrack" & vbCrLf & "(25) Windstorm by SGB" & vbCrLf & "(26) Reminiscence by 光宗信吉" & vbCrLf & "(27) It's Our Fight by Steve Jablonsky", GW - 50, 100, 16, argb(255, 255, 255, 255), align:=StringAlignmentFar
    End If
    
    If BGM.position > 85 And BGM.position < 90 Then
        Page.Writes "Soundtrack" & vbCrLf & "(28) Bright Red by 光宗信吉" & vbCrLf & "(29) Deadlocked by F-777" & vbCrLf & "(30) Now and Forever by Jazz Piano Bar Academy", GW - 50, 100, 16, argb(255, 255, 255, 255), align:=StringAlignmentFar
    End If
    
    If BGM.position > 90 And BGM.position < 95 Then
        Page.Writes "Soundtrack" & vbCrLf & "(31) System Split by F-777" & vbCrLf & "(32) Across The Waters[Short] by F-777" & vbCrLf & "(33) SWEET CANDY by 桃井はるこ", GW - 50, 100, 16, argb(255, 255, 255, 255), align:=StringAlignmentFar
    End If
    
    
    If BGM.position > 110 Then
        pro2 = (BGM.position - 110) / 2
        If pro2 > 1 Then pro2 = 1
        For i = 1 To 4
            Page.DrawImage "bm.png", GW / 2 - 200 + 100 * i, 900 - 900 * pro2 + GH - 50 - Cos((GetTickCount / 250) + i / 4) * 20, alpha:=1, Pos:=posOnCenter
        Next
    End If
    

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
    If EndMark Then
        BGM.Create App.Path & "\music\bgm\OST 31 - System Split.mp3"
        BGM.Play
    Else
        BGM.Create App.Path & "\music\bgm\OST 33 - SWEET CANDY.mp3"
        BGM.Play
    End If
    BGS.Stops
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
    Page.Res.NoLoadCrash = True
    Page.Res.NewImages App.Path & "\assets\face", 0.7
    Page.Res.newImage App.Path & "\assets\bg\garden.jpg", GW, GH
    Page.Res.newImage App.Path & "\assets\fg\Port.png", GW, GH
    Page.Res.newImage App.Path & "\assets\stars\logo.png", 0.8
    '创建页面
    ECore.Add Page, "EndingPage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
