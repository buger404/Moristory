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

Private Sub Form_KeyDown(keycode As Integer, Shift As Integer)
    'BXBattlePage.KeyDown KeyCode
    If ECore.ActivePage = "MazePage" Then MazePage.KeyDown keycode
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '发送字符输入
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
End Sub

Private Sub Form_Load()
    '初始化Emerald（在此处可以修改窗口大小哟~）
    StartEmerald Me.Hwnd, 1000, 740
    '创建字体
    MakeFont "微软雅黑"
    '创建页面管理器
    Set EC = New GMan
    
    '创建存档（可选）
    Set ESave = New GSaving
    ESave.Create "Moristory.TIMELINE", "kj" & Val(Me.Visible) & "Ehsd" & Val(VB.Screen.FontCount <> 0) & "Cdfd" & Right(Left("54B89", 3), 1) & "3fdkg5" & UCase("d") & "gsA6D1F7305BEjAC8738C" & CLng("&HE2") & "kjgds"

    '开始显示
    Me.Show
    
    Set BGM = New GMusic
    Set BGS = New GMusic
    Set SE = New GMusicList
    SE.HotLoad = True
    SE.Create App.Path & "\music\se"
    BGS.Volume = 0.3
    
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
        Set WeatherLayer = New WeatherLayer
    '=============================================

    '设置活动页面
    EC.ActivePage = "MainPage"
    
    Do While EndMark = False
        EC.Display
        DoEvents
    Loop
    
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    '发送鼠标信息
    UpdateMouse x, y, 1, button
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    '发送鼠标信息
    If Mouse.state = 0 Then
        UpdateMouse x, y, 0, button
    Else
        Mouse.x = x: Mouse.y = y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    '发送鼠标信息
    UpdateMouse x, y, 2, button
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
