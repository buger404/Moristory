VERSION 5.00
Begin VB.Form GameWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Moristory"
   ClientHeight    =   6672
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9660
   Icon            =   "GameWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   805
   StartUpPosition =   2  '屏幕中心
End
Attribute VB_Name = "GameWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================
'   页面管理器
    Dim EC As GMan
'==================================================
'   在此处放置你的页面类模块声明
    Dim EndMark As Boolean
'==================================================

Private Sub DrawTimer_Timer()
    '绘制
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'BXBattlePage.KeyDown KeyCode
    If ECore.ActivePage = "DancePage" Then
        DancePage.KeyUp KeyCode
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '发送字符输入
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If ECore.ActivePage = "MazePage" Then MazePage.KeyDown KeyCode
    If ECore.ActivePage = "NovelPage" Then
        NovelPage.KeyUp KeyCode
    End If
    
    If KeyCode = vbKeyS Then WeatherLayer.SwitchSetting

    If App.LogMode = 0 And KeyCode = vbKeyF3 Then WeatherLayer.SwitchDebug
End Sub

Private Sub Form_Load()
    '初始化Emerald（在此处可以修改窗口大小哟~）
    StartEmerald Me.Hwnd, 1000, 740
    '创建字体
    MakeFont "微软雅黑"
    '创建页面管理器
    Set EC = New GMan

    '开始显示
    Me.Show
    
    Set BGM = New GMusic
    Set BGS = New GMusic
    Set SE = New GMusicList
    SE.HotLoad = True
    SE.Create App.Path & "\music\se"
    BGS.Volume = 0.3
    
    Set ErrorPage = New ErrorPage
    
    '创建存档（可选）
    Set ESave = New GSaving
    ESave.Create "Moristory.TIMELINE", "kj" & Val(Me.Visible) & "Ehsd" & Val(VB.Screen.FontCount <> 0) & "Cdfd" & Right(Left("54B89", 3), 1) & "3fdkg5" & UCase("d") & "gsA6D1F7305BEjAC8738C" & CLng("&HE2") & "kjgds"

    '在此处初始化你的页面
    '=============================================
    '示例：TestPage.cls
    '     Set TestPage = New TestPage
    '公共部分：Dim TestPage As TestPage
        Set MainPage = New MainPage
        Set NovelPage = New NovelPage
        Set MazePage = New MazePage
        Set BattlePage = New BattlePage
        Set TicTacToePage = New TicTacToePage
        Set BXBattlePage = New BXBattlePage
        Set SnowmanPage = New SnowmanPage
        Set TipPage = New TipPage
        Set FlyPage = New FlyPage
        Set TLPPage = New TLPPage
        Set DancePage = New DancePage
        Set FinalPage = New FinalPage
        Set EndingPage = New EndingPage
        Set WeatherLayer = New WeatherLayer
    '=============================================

    '设置活动页面
    If EC.ActivePage = "" Then EC.ActivePage = "MainPage"
    
    Dim Time As Long, FPSTime As Long, FPS As Long, FPSTick As Long, FPSTarget As Long
    Dim LockFPS1 As Long, LockFPS2 As Long, Changed As Boolean
    FPSTime = GetTickCount: Time = GetTickCount
    '======================================================================
    '   LockFPS1：锁定的目标帧数（正常情况）
    '   LockFPS2：锁定的目标帧数（帧数不够时）
        LockFPS1 = 60: LockFPS2 = 40
    '======================================================================
    
    Do While EndMark = False
        '锁定帧数！
        If Changed = False Then
            If FPSctt > 0 And FPS > 0 And GetTickCount - FPSTime > 0 Then
                If FPS > LockFPS2 / 2 Then
                    '分析可以达到的FPS档次
                    'Me.Caption = "分析帧数：" & 1000 / (FPSctt / FPS)
                    If 1000 / (FPSctt / FPS) < LockFPS1 * 0.8 Then
                        FPSTarget = LockFPS2
                    Else
                        FPSTarget = LockFPS1
                    End If
                End If
                '动态设置间隔
                If FPSTarget > 0 Then FPSTick = (1000 / FPSTarget) / ((((GetTickCount - FPSTime) / FPS) * FPSTarget) / 100): Changed = True
            End If
            If FPSTick < 0 Then FPSTick = 0
        End If
        If GetTickCount - FPSTime >= 1000 Then
            FPSTime = GetTickCount
            FPS = 0
        End If
        If GetTickCount - Time >= FPSTick Then
            Time = GetTickCount: FPS = FPS + 1: Changed = False
            EC.Display
        End If
        DoEvents
    Loop
    
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
    EndMark = True
    Set NovelPage = Nothing
    Me.Hide
    '终止绘制
    'DrawTimer.Enabled = False
    '释放Emerald资源
    EndEmerald
    End
End Sub
