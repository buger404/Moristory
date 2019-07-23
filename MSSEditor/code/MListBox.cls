VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type MListItem
    Content As String
    Selected As Boolean
End Type
Dim Items() As MListItem
Public DC As Long, GG As Long
Public ListIndex As Integer
Public ScrollY As Long
Public Property Get ListCount() As Long
    ListCount = UBound(Items)
End Property
Public Sub AddItem(Item As String, Optional pos As Integer = -1)
    SaveMark = False
    ReDim Preserve Items(UBound(Items) + 1)
    Items(UBound(Items)).Content = Item
    If pos <> -1 Then
        For I = UBound(Items) - 1 To pos Step -1
            Items(I + 1) = Items(I)
        Next
        Items(pos).Content = Item
    End If
End Sub
Public Sub RemoveItem(index As Integer)
    SaveMark = False
    For I = index To UBound(Items) - 1
        Items(I) = Items(I + 1)
    Next
    ReDim Preserve Items(UBound(Items) - 1)
End Sub
Public Property Get List(index As Integer) As String
    List = Items(index + 1).Content
End Property
Public Property Let List(index As Integer, Content As String)
    SaveMark = False
    Items(index + 1).Content = Content
End Property
Public Property Get Selected(index As Integer) As Boolean
    Selected = Items(index + 1).Selected
End Property
Public Property Let Selected(index As Integer, Sel As Boolean)
    Items(index + 1).Selected = Sel
End Property
Public Sub Clear()
    ReDim Items(0)
End Sub
Private Sub Class_Initialize()
    ReDim Items(0)
    ListIndex = -1
    DC = CreateCDC(GW, GH)
    PoolCreateFromHdc DC, GG
    GdipSetSmoothingMode GG, SmoothingModeAntiAlias
    GdipSetTextRenderingHint GG, TextRenderingHintAntiAliasGridFit
End Sub
