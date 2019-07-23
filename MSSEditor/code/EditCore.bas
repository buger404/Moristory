Attribute VB_Name = "EditCore"
Public MSSList As MListBox
Public Music As GMusic
Public EditIndex As String, SaveMark As Boolean
Public RoleList() As String
Public AddMode As Boolean
Public NowRole As String
Public Sub MuSay()
    Dim SayName As String
    SayName = VBA.InputBox("À≠Àµ¡À’‚–©£ø", "Moristory")
    If SayName = "" Then Exit Sub
    Dim I As Integer
    Do While I <= MSSList.ListCount - 1
        If MSSList.Selected(I) Then
            MSSList.Selected(I) = False
            MSSList.AddItem "say\" & SayName, I + 1
            I = I + 1
        End If
        I = I + 1
    Loop
End Sub
Public Sub MakeBackup()
    If Dir(App.Path & "\..\article\PART " & EditIndex & ".mss") = "" Then Exit Sub
    FileCopy App.Path & "\..\article\PART " & EditIndex & ".mss", _
    App.Path & "\backup\PART " & EditIndex & " - " & year(Now) & "." & Month(Now) & "." & Day(Now) & "  " & Hour(Now) & "_" & Minute(Now) & "_" & Second(Now) & ".mss"
End Sub
Public Sub GetRoles()
    Dim List() As String, temp() As String, I As Integer
    ReDim List(0)
    ReDim Preserve List(UBound(List) + 1)
    List(UBound(List)) = "aside"
    
    For I = 2 To MSSList.ListIndex
        temp = Split(MSSList.List(I), "\")
        If UBound(temp) = 2 Then
            If temp(0) = "role" Then
                If temp(2) = "add" Then
                    ReDim Preserve List(UBound(List) + 1)
                    List(UBound(List)) = temp(1)
                ElseIf temp(2) = "remove" Then
                    For S = 1 To UBound(List)
                        If List(S) = temp(1) Then List(S) = List(UBound(List)): ReDim Preserve List(UBound(List) - 1): Exit For
                    Next
                End If
            End If
        End If
    Next
    
    For I = 1 To UBound(List)
        Select Case List(I)
            Case "": List(I) = List(I) & " - æ¯∂‘ «ƒ„Õ¸º«–¥¡À"
            Case "aside": List(I) = List(I) & " - ≈‘∞◊"
            Case "bm": List(I) = List(I) & " - ∫⁄◊Ï"
            Case "xx": List(I) = List(I) & " - Ÿ‚Ÿ‚"
            Case "me": List(I) = List(I) & " - «‡"
            Case "xl": List(I) = List(I) & " - —©¿«"
            Case "ssr": List(I) = List(I) & " - …Ø…™»"
            Case "kx1": List(I) = List(I) & " - ø®Œ˜"
            Case "kx2": List(I) = List(I) & " - ø®‹Á"
            Case "fj": List(I) = List(I) & " - ∏°ÈŸ"
            Case "km1": List(I) = List(I) & " - ø›√Œ"
            Case "jy": List(I) = List(I) & " - æ√”∆"
            Case "yz": List(I) = List(I) & " - —ø◊”"
            Case "bg": List(I) = List(I) & " - ±˘π˜"
        End Select
    Next
    
    RoleList = List
End Sub
