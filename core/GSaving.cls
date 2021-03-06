VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GSaving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Emerald 相关代码

Private Type SaveData
    Name As String
    value As String
    datacheck As String
End Type
Private Type SaveFile
    data() As SaveData
End Type
Dim MySave As SaveFile
Public AutoSave As Boolean
Public sToken As Boolean, Path As String
Public Key As String
'========================================================
'   Init
    Public Function Create(Name As String, Optional BMKey) As Boolean

        Path = "C:\Users\" & PoolCreateObject("Wscript.Network").Username & "\AppData\Local\" & Name
        AutoSave = True
        
        If Not IsMissing(BMKey) Then Key = BMKey
        
        ReDim MySave.data(0)
        If Dir(Path & "\", vbDirectory) = "" Then
            If Name <> "Emerald.Core" Then
                SysPage.index = 0: SysPage.Key = 0
                Dim Win As Object
                Set Win = StartScreenDialog(600, 400, SysPage)
                Set SysPage.Win = Win
                Do While Not Win.AcceptMark
                    DoEvents
                Loop
                Unload Win
                
                If SysPage.Key = 1 Then MkDir Path & "\": Create = True
            Else
                MkDir Path & "\"
            End If
        Else
            If Dir(Path & "\save") <> "" Then
                On Error GoTo ErrHandle
                Open Path & "\save" For Binary As #1
                Get #1, , MySave
                Close #1
                Call CheckSave
ErrHandle:
                If Err.number <> 0 Then ErrCrash Err.number, Err.Description
            End If
            Create = True
        End If
        
        sToken = Create
        
        If ESave Is Nothing Then Set ESave = Me
    End Function
'========================================================
'   RunTime
    Public Sub CheckSave()
        On Error GoTo ErrBM
    
        Dim msg As String
        For i = 1 To UBound(MySave.data)
            If MySave.data(i).value <> "" Then
                If MySave.data(i).datacheck <> IIf(Key = "", BMEA(MySave.data(i).value & i), BMEA(MySave.data(i).value & i, Key)) Then
                    msg = msg & MySave.data(i).Name & vbCrLf
                    MySave.data(i).value = ""
                End If
            End If
        Next
        
        If msg <> "" Then
            Call Save
            MsgBox "已经发现并删除存档中以下非法数据：" & vbCrLf & msg, 64
        End If
        
ErrBM:     If Err.number <> 0 Then ErrCrash Err.number, Err.Description
    End Sub
'========================================================
'   Manager
    Public Function Count() As Long
        Count = UBound(MySave.data)
    End Function
    Public Function GetName(index As Integer) As String
        GetName = MySave.data(index).Name
    End Function
    Public Sub Save()
        On Error GoTo ErrBM

        If sToken = False Then Exit Sub
        Open Path & "\save" For Binary As #1
        Put #1, , MySave
        Close #1
        
ErrBM:     If Err.number <> 0 Then ErrCrash Err.number, Err.Description
    End Sub
    Public Function GetData(Name As String) As String
        On Error GoTo ErrBM

        For i = 1 To UBound(MySave.data)
            If MySave.data(i).Name = Name Then GetData = MySave.data(i).value: Exit For
        Next
        
ErrBM:     If Err.number <> 0 Then ErrCrash Err.number, Err.Description
    End Function
    Public Sub PutData(Name As String, ByVal value As String)
        On Error GoTo ErrBM

    
        Dim fd As Boolean
        For i = 1 To UBound(MySave.data)
            If MySave.data(i).Name = Name Then
                MySave.data(i).value = value
                MySave.data(i).datacheck = IIf(Key = "", BMEA(value & i), BMEA(value & i, Key))
                fd = True
                Exit For
            End If
        Next
        If fd = False Then
            '新增
            ReDim Preserve MySave.data(UBound(MySave.data) + 1)
            MySave.data(UBound(MySave.data)).Name = Name
            MySave.data(UBound(MySave.data)).value = value
            MySave.data(UBound(MySave.data)).datacheck = IIf(Key = "", BMEA(value & UBound(MySave.data)), BMEA(value & UBound(MySave.data), Key))
        End If
        If AutoSave Then Call Save
        
ErrBM:     If Err.number <> 0 Then ErrCrash Err.number, Err.Description
    End Sub
'========================================================
