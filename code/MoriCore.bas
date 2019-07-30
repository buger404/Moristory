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
        Case 0: NewStr = "。。。等一下？哪来的错误？"
        Case 5: NewStr = "该死的404又忘记把废弃的代码删干净了。"
        Case 6: NewStr = "shit 404为变量倒果汁的时候不小心溢出来了。"
        Case 7: NewStr = "404没有做好清洁工的职责。"
        Case 9: NewStr = "404让我在数组的边缘试探...哦豁，掉入悬崖了。"
        Case 11: NewStr = "脑残404数学不过关，除以0什么的..."
        Case 13: NewStr = "404刚才在给别人介绍对象的时候被双方左右各扇了一巴掌。"
        Case 28: NewStr = "罢工！404一次性让我们做太多的工作了！"
        Case 35: NewStr = "该死的404又忘记把废弃的代码删干净了。"
        Case 52: NewStr = "404给错了房间号码。"
        Case 53: NewStr = "404的GPS出了点问题..."
        Case 55: NewStr = "丢三落四的404打开一个文件后忘了关上。"
        Case 58: NewStr = "404说的这个文件已经存在了，用TNT炸掉么？"
        Case 70: NewStr = "404让身在底层的我做实在没有权利完成的工作。"
        Case 75: NewStr = "404给了我一个错误的地址！"
        Case 76: NewStr = "404让我来到已经拆迁的公寓。"
    End Select
    
    LastPage = ECore.ActivePage
    If Num = 404233 Then
        ErrorPage.ErrText = "游戏没有问题，只是被迫终止。" & vbCrLf & NewStr & vbCrLf & "黑嘴表示 不 非常抱歉，您可以忘记刚才按到了什么。"
    Else
        ErrorPage.ErrText = IIf(ECore.ActivePage = "LoadingPage", "游戏似乎根本没有正常启动(" & Num & ")。", "当你在<" & ECore.ActivePage & ">里玩耍的时候，游戏内部发生了一些问题(" & Num & ")。") & vbCrLf & NewStr & vbCrLf & "黑嘴表示非常抱歉，您可以尝试重新开启游戏或反馈此问题。"
    End If
    
    If Not WeatherLayer Is Nothing Then WeatherLayer.Page.TopPage = False
    ECore.ActivePage = "ErrorPage"
End Sub
