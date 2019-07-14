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

Public Function GetPartTitle(Part As String) As String
    Dim temp As String
    Open App.Path & "\article\PART " & Part & ".mss" For Input As #1
    Do While Not EOF(1)
        Line Input #1, temp
        If temp Like "*title\*" Then GetPartTitle = Split(temp, "\")(1): Close #1: Exit Function
    Loop
    Close #1
End Function
