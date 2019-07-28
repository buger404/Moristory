VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Private Type RoleItem
        index As Long
        Name As String
        Face As String
    End Type
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear argb(255, 40, 41, 35) '清空画布

    Page.Paint 0, 0, 0, GW + 1, 40, argb(255, 66, 67, 62)
    
    Dim DX As Long, DY As Long, Dh As Long, DW As Long, Menus(5) As String
    DX = 0
    Menus(0) = "File"
    Menus(1) = "Music"
    Menus(2) = "Pictures"
    Menus(3) = "Roles"
    Menus(4) = "Save"
    Menus(5) = "Mu-Say"
    
    Dim I As Integer

    For I = 0 To UBound(Menus)
        DW = EF.GetWidth(Page.GG, Menus(I), 16, StringAlignmentCenter, FontStyleRegular)
        If Page.ShowColorButton(0, DX, 0, DW + 40, 40, Menus(I), argb(255, 242, 242, 242), argb(255, 66, 67, 62), argb(255, 109, 110, 106), Size:=16) = 3 Then
            AddMode = True
            Select Case I
                Case 0: MainWindow.PopupMenu MainWindow.File
                Case 1: MainWindow.PopupMenu MainWindow.musicMenu
                Case 2: MainWindow.PopupMenu MainWindow.PictureMenu
                Case 5: Call MuSay
                Case 4: MSSList.AddItem "*save\", MSSList.ListIndex + 1
                Case 3
                    Call GetRoles
                    For S = 1 To MainWindow.removeroleBtn.UBound
                        Unload MainWindow.removeroleBtn(S)
                    Next
                    For S = 2 To UBound(RoleList)
                        Load MainWindow.removeroleBtn(S - 1)
                        With MainWindow.removeroleBtn(S - 1)
                            .Caption = RoleList(S)
                            .Visible = True
                        End With
                    Next
                    MainWindow.PopupMenu MainWindow.rolectrlMenu
            End Select
        End If
        DX = DX + DW + 40
    Next
    
    Page.Writes "Part " & EditIndex, GW / 2, 70, 28, IIf(SaveMark, argb(255, 0, 183, 195), argb(255, 254, 84, 57)), align:=StringAlignmentCenter
    
    Dim ODC As Long, OGG As Long, NowSpeaker As String, m As Integer, temp() As String
    Dim Role() As RoleItem
    ReDim Role(0)
    ODC = Page.CDC: OGG = Page.GG
    Page.CDC = MSSList.DC: Page.GG = MSSList.GG
    
    Page.Clear 0
    Dim RoleColor As Long
    For I = 0 To MSSList.ListCount - 1
        If InStr(MSSList.List(I), "\") Then
            If InStr(MSSList.List(I), "say\") = 1 Then NowSpeaker = Split(MSSList.List(I), "say\")(1)
            If MSSList.List(I) Like "face\*" Then
                temp = Split(MSSList.List(I), "\")
                For S = 1 To UBound(Role)
                    If Role(S).Name = temp(1) Then Role(S).Face = temp(2): Role(S).index = Page.Res.GetImage(temp(1) & "-" & temp(2) & ".png"): Exit For
                Next
            End If
            If MSSList.List(I) Like "role\*" Then
                temp = Split(MSSList.List(I), "\")
                If temp(2) = "add" Then
                    ReDim Preserve Role(UBound(Role) + 1)
                    Role(UBound(Role)).Name = temp(1)
                    Role(UBound(Role)).index = Page.Res.GetImage(temp(1) & "-normal.png")
                    Role(UBound(Role)).Face = "Normal"
                ElseIf temp(2) = "remove" Then
                    For S = 1 To UBound(Role)
                        If Role(S).Name = temp(1) Then Role(S) = Role(UBound(Role)): Exit For
                    Next
                    ReDim Preserve Role(UBound(Role) - 1)
                End If
            End If
        End If
        DY = I * 30 - MSSList.ScrollY
        If DY > GH - 150 - 40 Then Exit For
        If DY >= -30 And DY <= GH - 150 - 40 Then
            m = CheckMouse(30, DY + 40 + 100, (GW - 60 - 10) / 2, 31)
            If InStr(MSSList.List(I), "\") Then
                temp = Split(MSSList.List(I), "\")
                If MSSList.List(I) Like "*bg\*" Then Page.DrawImage temp(1), 0, DY, 0, GH * 0.7, GW - 60, 30, alpha:=0.5
                If MSSList.List(I) Like "bg\*" Then
                    temp = Split(temp(1), ",")
                    Page.Paint 0, 0, DY, GW - 60, 31, argb(120, temp(0), temp(1), temp(2))
                End If
                If MSSList.List(I) Like "*fg\*" Then Page.DrawImage temp(1), 0, DY, 0, GH * 0.15, GW - 60, 30, alpha:=0.5
                'Page.Paint 0, 0, DY, GW, 31, argb(40, 208, 200, 202)
                If MSSList.List(I) Like "*role\*" Then Page.Paint 0, 0, DY, GW - 60, 31, argb(50, 252, 140, 28)
            Else
                Select Case NowSpeaker
                    Case "": RoleColor = argb(100, 255, 0, 0)
                    Case "aside": RoleColor = argb(30, 208, 200, 202)
                    Case "bm": RoleColor = argb(50, 206, 169, 126)
                    Case "xx": RoleColor = argb(50, 247, 218, 247)
                    Case "me": RoleColor = argb(50, 50, 202, 185)
                    Case "dark": RoleColor = argb(50, 0, 0, 0)
                    Case "xl": RoleColor = argb(50, 230, 248, 255)
                    Case "ssr": RoleColor = argb(50, 135, 202, 196)
                    Case "kx1": RoleColor = argb(50, 21, 128, 255)
                    Case "kx2": RoleColor = argb(50, 255, 142, 145)
                    Case "fj": RoleColor = argb(50, 255, 0, 0)
                    Case "km1": RoleColor = argb(50, 250, 3, 156)
                    Case "jy": RoleColor = argb(50, 205, 145, 146)
                    Case "yz": RoleColor = argb(50, 145, 30, 123)
                    Case "bg": RoleColor = argb(50, 198, 198, 129)
                    
                    Case "tk": RoleColor = argb(50, 128, 0, 0)
                    Case "yy": RoleColor = argb(50, 254, 84, 57)
                    Case "s": RoleColor = argb(50, 65, 165, 238)
                End Select
                Page.Paint 0, 0, DY, GW - 60, 31, RoleColor
                Page.Paint 0, 0, DY, 5, 31, RoleColor
                Page.Paint 0, 0, DY, 5, 31, RoleColor
            End If
            If MSSList.ListIndex = I Then
                Page.Paint 0, 0, DY, GW - 60, 31, argb(20, 255, 255, 255)
                Page.Paint 0, 0, DY, 5, 31, argb(255, 255, 255, 255)
            End If
            If m > 0 Then Page.Paint 0, 0, DY, GW - 60, 31, argb(10, 255, 255, 255)
            Page.Writes MSSList.List(I), 10, DY + 3, 16, argb(255, 222, 222, 222)
            If m = 3 Then
                If Mouse.button = 1 Then
                    If MSSList.ListIndex <> I Then
                        MSSList.ListIndex = I
                    Else
                        If InStr(MSSList.List(I), "\") Then
                            AddMode = False
                            Dim IsCmdPro As Boolean
                            If MSSList.List(I) Like "say\*" Then
                                Call GetRoles
                                For S = 1 To MainWindow.roleBtn.UBound
                                    Unload MainWindow.roleBtn(S)
                                Next
                                For S = 1 To UBound(RoleList)
                                    Load MainWindow.roleBtn(S)
                                    With MainWindow.roleBtn(S)
                                        .Caption = RoleList(S)
                                        .Visible = True
                                    End With
                                Next S
                                Load MainWindow.roleBtn(S)
                                With MainWindow.roleBtn(S)
                                    .Caption = "+ Custom"
                                    .Visible = True
                                End With
                                MainWindow.PopupMenu MainWindow.roleMenu
                                IsCmdPro = True
                            End If
                            If MSSList.List(I) Like "*bgm\*" Then MainWindow.PopupMenu MainWindow.bgmMenu: IsCmdPro = True
                            If MSSList.List(I) Like "*bgs\*" Then MainWindow.PopupMenu MainWindow.BGSMenu: IsCmdPro = True
                            If MSSList.List(I) Like "*play\*" Then MainWindow.PopupMenu MainWindow.SEMenu: IsCmdPro = True
                            If Not IsCmdPro Then MSSList.List(I) = VBA.InputBox("Edit", "Editor", MSSList.List(I))
                        Else
                            MSSList.List(I) = VBA.InputBox("Edit", "Editor", MSSList.List(I))
                            SaveMark = False
                        End If
                    End If
                ElseIf Mouse.button = 2 Then
                    If InStr(MSSList.List(I), "\") Then
                        If (MSSList.List(I) Like "*bgm\*") Or (MSSList.List(I) Like "*bgs\*") Or (MSSList.List(I) Like "*play\*") Then
                            If Music.IsPlaying Then
                                Music.Stops
                            Else
                                If MSSList.List(I) Like "*bgm\*" Then Music.Create App.Path & "\..\music\bgm\" & Split(MSSList.List(I), "\")(1)
                                If MSSList.List(I) Like "*bgs\*" Then Music.Create App.Path & "\..\music\bgs\" & Split(MSSList.List(I), "\")(1)
                                If MSSList.List(I) Like "*play\*" Then Music.Create App.Path & "\..\music\se\" & Split(MSSList.List(I), "\")(1)
                                Music.Play
                            End If
                        End If
                        If MSSList.List(I) Like "*bg\*" Then MainWindow.PopupMenu MainWindow.bgMenu
                        If MSSList.List(I) Like "*fg\*" Then MainWindow.PopupMenu MainWindow.ForegroundMenu
                    Else
                        MSSList.Selected(I) = Not MSSList.Selected(I)
                    End If
                End If
            End If
            
            Page.Paint 1, GW - 60 - 5 - 35, DY + 5, 20, 20, IIf(MSSList.Selected(I), argb(255, 50, 202, 185), argb(40, 255, 255, 255))
            If CheckMouse(GW - 60 - 5 - 35 + 30, DY + 5 + 140, 20, 20) = mMouseUp Then
                MSSList.Selected(I) = Not MSSList.Selected(I)
            End If
            
            Dim index As Long, IW As Long
            DX = 20
            If InStr(MSSList.List(I), "\") = 0 Then
                For S = 1 To UBound(Role)
                    IW = Page.Res.ImgSize(Role(S).index, imgGetWidth)
                    DX = DX + IW
                    Page.DrawImage Role(S).index, GW - 60 - 5 - 35 - DX, DY, CH:=30, alpha:=IIf(NowSpeaker = Role(S).Name, 1, 0.4)
                    Page.Writes Role(S).Face, GW - 60 - 5 - 35 - DX, DY + 5, 18, argb(255, 255, 255, 255), IW, 30, StringAlignmentCenter, FontStyleBold
                    If CheckMouse(GW - 50 - DX, DY + 40 + 100, IW, 30) = 3 Then
                        Dim FaceN As String
                        For p = 1 To MainWindow.FaceBtn.UBound
                            Unload MainWindow.FaceBtn(p)
                        Next
                        FaceN = Dir(App.Path & "\..\assets\face\" & Role(S).Name & "-*.png")
                        Do While FaceN <> ""
                            If Not (FaceN Like "*-normal.png") Then
                                Load MainWindow.FaceBtn(MainWindow.FaceBtn.UBound + 1)
                                With MainWindow.FaceBtn(MainWindow.FaceBtn.UBound)
                                    .Caption = Split(Split(FaceN, ".png")(0), "-")(1)
                                    .Visible = True
                                End With
                            End If
                            FaceN = Dir()
                        Loop
                        NowRole = Role(S).Name
                        MainWindow.PopupMenu MainWindow.FaceMenu
                        SaveMark = False
                    End If
                Next
            End If
        End If
    Next
    
    Dim MaxY As Long
    MaxY = MSSList.ListCount * 30 - (GH - 100 - 40 - 50)
    If MaxY < 0 Then MaxY = 0
    
    If MaxY > 0 Then
        Page.Paint 2, GW - 60 - 5, MSSList.ScrollY / MaxY * (GH - 100 - 40 - 50 - 50), 5, 50, argb(255, 255, 255, 255), 10
        If CheckMouse(30 + (GW - 60 - 10), 40 + 100, 50, GH - 100 - 40 - 50) >= mMouseDown Then
            MSSList.ScrollY = (Mouse.y - 45 - (40 + 100)) / (GH - 100 - 40 - 50 - 50) * MaxY
        End If
    End If
    
    If IsKeyPress(VK_UP) Then MSSList.ScrollY = MSSList.ScrollY - 15
    If IsKeyPress(VK_DOWN) Then MSSList.ScrollY = MSSList.ScrollY + 15
    
    If MSSList.ScrollY < 0 Then MSSList.ScrollY = 0
    If MSSList.ScrollY > MaxY Then MSSList.ScrollY = MaxY
    
    PaintDC MSSList.DC, ODC, 30, 40 + 100, CW:=GW - 60, CH:=GH - 100 - 40 - 50, alpha:=1
    
    Page.CDC = ODC: Page.GG = OGG
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(ID As String, Msg As String)
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
Public Sub AnimationDone(ID As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.Res.NewImages App.Path & "\assets\fg", GW, GH
    Page.Res.NewImages App.Path & "\assets\bg", GW, GH
    Page.Res.NewImages App.Path & "\..\assets\face", 0.2
    '创建页面
    ECore.Add Page, "MainPage"
    Set MSSList = New MListBox
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
