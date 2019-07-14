VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Emerald 相关代码

Private Declare Sub AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal hHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BLENDFUNCTION As Long) ' As Long
Private Type AniTask
    Start As Long
    during As Long
    custom As Boolean
    func As String
    profunc As Integer
    delay As Long
    mark As Boolean
End Type
Dim GTick As Long
Dim Anis() As AniTask
Dim DrawBox As Object, brush As Long, Path As Long, Pen As Long
Dim ScrollMode As Boolean, ScrollX As Long, ScrollY As Long, ScrollBX As Long, ScrollBY As Long, CRgn As Long
Dim AniCollection() As EAnimation, PlayAni() As EAnimation
Public CDC As Long, GG As Long, OutOfScroll As Boolean
Public Res As GResource
Public TopPage As Boolean
'========================================================
'   Init
    Public Sub Create(dBox As Object)
        ReDim Anis(0)
        ReDim AniCollection(0): ReDim PlayAni(0)
        Set Res = New GResource
        
        GTick = 1000 / 30
        CDC = CreateCDC(GW, GH)
        Res.Add CDC, 0, ""
        PoolCreateFromHdc CDC, GG
        GdipSetSmoothingMode GG, SmoothingModeAntiAlias
        GdipSetTextRenderingHint GG, TextRenderingHintAntiAlias
        
        Set DrawBox = dBox
        
        PoolCreateSolidFill 0, brush
        PoolCreatePen1 0, 1, UnitPixel, Pen
        PoolCreatePath FillModeWinding, Path
    End Sub
    Public Sub Dispose()
        Exit Sub
        PoolDeleteBrush brush
        PoolDeletePen Pen
        GdipDeletePath Path
        Res.Dispose
    End Sub
'========================================================
'   Print
    Public Sub Clear(Optional Color As Long = -1)
        GdipGraphicsClear GG, Color
    End Sub
    Public Sub Update()
        Dim i As Integer
        For i = 1 To UBound(Anis)
            If Not Anis(i).mark Then
                If (GetTickCount - Anis(i).Start - Anis(i).delay) > Anis(i).during Then
                     CallByName DrawBox, "AnimationDone", VbMethod, i
                    Anis(i).mark = True
                End If
            End If
        Next
        
        CallByName DrawBox, "Update", VbMethod
    End Sub
    Public Sub Display(DC As Long)
        If TopPage Then
            PaintDC CDC, DC, alpha:=1
        Else
            BitBlt DC, 0, 0, GW, GH, CDC, 0, 0, vbSrcCopy
        End If
    End Sub
'========================================================
'   RunTime
    Public Property Get SrcClass() As Object
        Set SrcClass = DrawBox
    End Property
    Public Property Get GIFTick() As Long
        GIFTick = GTick
    End Property
    Public Property Let GIFTick(nGIFTick As Long)
        GTick = nGIFTick
    End Property
    Public Sub LoadAnimation(Path As String)
        ReDim Preserve AniCollection(UBound(AniCollection) + 1)
        AniCollection(UBound(AniCollection)) = Animations.LoadAnimation(Path)
        With AniCollection(UBound(AniCollection))
            For i = 1 To UBound(.channel)
                For s = 1 To UBound(.channel(i).frames)
                    For t = 0 To UBound(.channel(i).frames(s).aframes)
                        .channel(i).frames(s).aframes(t).picindex = Res.GetImage(.channel(i).frames(s).aframes(t).Pic)
                    Next
                Next
            Next
        End With
    End Sub
    Public Sub LoadAnimationsFromDir(Path As String)
        Dim f As String
        If Right(Path, 1) <> "\" Then Path = Path & "\"
        
        f = Dir(Path & "*.emrani")
        Do While f <> ""
            Call LoadAnimation(Path & f)
            f = Dir()
        Loop
    End Sub
    Private Function FindAnimation(name As String) As EAnimation
        For i = 1 To UBound(AniCollection)
            If AniCollection(i).name = name Then FindAnimation = AniCollection(i): Exit For
        Next
    End Function
    Private Function GetPlayAnimation(playname As String) As Integer
        For i = 1 To UBound(PlayAni)
            If PlayAni(i).name = playname Then GetPlayAnimation = i: Exit For
        Next
    End Function
    Public Sub CreatePlayAnimation(name As String, playname As String, channel As String)
        ReDim Preserve PlayAni(UBound(PlayAni) + 1)
        PlayAni(UBound(PlayAni)) = FindAnimation(name)
        PlayAni(UBound(PlayAni)).name = playname
        PlayAni(UBound(PlayAni)).tick = GetTickCount
        For i = 1 To UBound(PlayAni(UBound(PlayAni)).channel)
            PlayAni(UBound(PlayAni)).channel(i).CurrentFrame = 1
            If PlayAni(UBound(PlayAni)).channel(i).name = channel Then PlayAni(UBound(PlayAni)).CurrentChannel = i
        Next
    End Sub
    Public Sub DisposePlayAnimation(playname As String)
        PlayAni(GetPlayAnimation(playname)) = PlayAni(UBound(PlayAni))
        ReDim Preserve PlayAni(UBound(PlayAni) - 1)
    End Sub
    Public Sub SwitchChannel(playname As String, channel As String)
        Dim s As Integer
        s = GetPlayAnimation(playname)
        If s = 0 Then Exit Sub
        If PlayAni(s).channel(PlayAni(s).CurrentChannel).frames(PlayAni(s).channel(PlayAni(s).CurrentChannel).CurrentFrame).tick = -1 Then
            PlayAni(s).channel(PlayAni(s).CurrentChannel).frames(PlayAni(s).channel(PlayAni(s).CurrentChannel).CurrentFrame).tick = 0
        End If
        
        For i = 1 To UBound(PlayAni(s).channel)
            If PlayAni(s).channel(i).name = channel Then
                PlayAni(s).channel(i).CurrentFrame = 1
                PlayAni(s).CurrentChannel = i
                PlayAni(s).tick = GetTickCount
                Exit For
            End If
        Next
    End Sub
'========================================================
'   Images
    Public Sub DrawAnimation(playname As String, X As Long, Y As Long, Optional Direction As ImgDirection = DirNormal)
        Dim s As Integer, m As Integer, W As Long, h As Long, last As Boolean
        
        s = GetPlayAnimation(playname)
        If s = 0 Then Exit Sub
        
        With PlayAni(s).channel(PlayAni(s).CurrentChannel).frames(PlayAni(s).channel(PlayAni(s).CurrentChannel).CurrentFrame)
            For i = 0 To UBound(.aframes)
                If .aframes(i).alpha = 1 Then m = 1
                If .aframes(i).alpha <> 1 Then m = 1
                If .aframes(i).Size <> 1 Then m = 0
                
                If m = 0 Then
                    W = Res.ImgSize(.aframes(i).picindex, imgGetWidth) * .aframes(i).Size
                    h = Res.ImgSize(.aframes(i).picindex, imgGetHeight) * .aframes(i).Size
                    DrawImageEx .aframes(i).picindex, X + .aframes(i).X, Y + .aframes(i).Y, W, h, PlayAni(s).position
                Else
                    DrawImage .aframes(i).picindex, X + .aframes(i).X, Y + .aframes(i).Y, Pos:=PlayAni(s).position, alpha:=.aframes(i).alpha, Direction:=Direction
                End If
            Next
            
            Dim DisposeMark As Boolean
            If .tick <> -1 Then
                last = (PlayAni(s).channel(PlayAni(s).CurrentChannel).CurrentFrame = UBound(PlayAni(s).channel(PlayAni(s).CurrentChannel).frames))
                If GetTickCount - PlayAni(s).tick >= .tick Then
                    If .msg = "dispose" Or .disposed Then DisposeMark = True
                    If .msg <> "" Then
                        On Error Resume Next
                        CallByName SrcClass, "AnimationMsg", VbMethod, playname, .msg
                    End If
                    If .tick <> 0 Then
                        If last Then
                            PlayAni(s).channel(PlayAni(s).CurrentChannel).CurrentFrame = 1
                        Else
                            PlayAni(s).channel(PlayAni(s).CurrentChannel).CurrentFrame = PlayAni(s).channel(PlayAni(s).CurrentChannel).CurrentFrame + 1
                        End If
                    End If
                    If .tick = 0 Then .tick = -1
                    PlayAni(s).tick = GetTickCount
                    If Not (EMusic Is Nothing) Then
                        For i = 0 To UBound(PlayAni(s).channel(PlayAni(s).CurrentChannel).frames(PlayAni(s).channel(PlayAni(s).CurrentChannel).CurrentFrame).sounds)
                            EMusic.Play PlayAni(s).channel(PlayAni(s).CurrentChannel).frames(PlayAni(s).channel(PlayAni(s).CurrentChannel).CurrentFrame).sounds(i).snd
                        Next
                    End If
                End If
            End If
        End With
        
        If DisposeMark Then DisposePlayAnimation playname
    End Sub
    Public Sub DrawImageEx(N, X As Long, Y As Long, Optional W, Optional h, Optional Pos As PosAlign = posNormal, Optional angle As Long = 0, Optional animation As Integer = 0)
        Dim index As Integer
        
        Dim OX As Long, OY As Long, ow As Long, oh As Long
        OX = X: OY = Y: ow = W: oh = h
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during)
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, OX, OY, ow, oh, 0, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, OX, OY, ow, oh, 0, pro
                End If

            End If
        End If
        
        If TypeName(N) = "String" Then
            index = Res.GetImage(N)
        Else
            index = N
        End If
        If index = 0 Then Exit Sub
        
        If IsMissing(ow) Then GdipGetImageWidth Res.ResourceImageHandle(index), ow
        If IsMissing(oh) Then GdipGetImageWidth Res.ResourceImageHandle(index), oh
        
        If Pos = 1 Then OX = Int(OX - ow / 2): OY = Int(OY - oh / 2)
        If Pos = 2 Then OX = OX - ow
        If Pos = 3 Then OY = OY - oh
        If Pos = 4 Then OX = OX + ow
        If Pos = 5 Then OY = OY + oh
        
        If ScrollMode Then
            OX = OX + ScrollX + ScrollBX: OY = OY + ScrollY + ScrollBY
            Dim Ret As Long
            GdipIsVisibleRect GG, OX - 1, OY - 1, ow, oh, Ret
            If Not Ret Then OutOfScroll = True: Exit Sub
        End If
        
        
        If angle <> 0 Then
            GdipRotateWorldTransform GG, angle, MatrixOrderAppend
            GdipTranslateWorldTransform GG, OX, OY, MatrixOrderAppend
            GdipDrawImageRect GG, Res.ResourceImageHandle(index), -ow / 2, -oh / 2, ow, oh
            GdipResetWorldTransform GG
        Else
            GdipDrawImageRect GG, Res.ResourceImageHandle(index), OX, OY, ow, oh
        End If
        
        With DrawF
            .X = OX
            .Y = OY
            .Width = ow
            .Height = oh
            .CrashIndex = Res.ResourceCrashIndex(index)
            .Shape = -1
            .WSc = Res.ImgSize(index, imgGetWidth) / .Width
            .HSc = Res.ImgSize(index, imgGetHeight) / .Height
        End With
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 255, 0, 255)
            GdipDrawRectangle GG, Pen, OX, OY, ow - 1, oh - 1
        End If
        
        If Debug_pos Then
            GdipSetPenColor Pen, argb(255, 0, 183, 195)
            GdipDrawLine GG, Pen, X - 5, Y, X + 5, Y
            GdipDrawLine GG, Pen, X, Y - 5, X, Y + 5
        End If
    End Sub
    Public Sub DrawImage(N, X As Long, Y As Long, Optional cx, Optional cy, Optional cw, Optional ch, Optional alpha, Optional Pos As PosAlign = posNormal, Optional animation As Integer = 0, Optional Direction As ImgDirection = DirNormal)
        'If OutOfScroll Then Exit Sub
        
        Dim b As BLENDFUNCTION, index As Integer, bl As Long
        Dim OX As Long, OY As Long, ow As Long, oh As Long
        OX = X: OY = Y: ow = W: oh = h
        Dim ocx As Long, ocy As Long, ocw As Long, och As Long
        
        If Not IsMissing(alpha) Then
            Dim pro As Single
            pro = 1
            If animation <> 0 Then
                If Not Anis(animation).mark Then
                    pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during)
                End If
            End If
            alpha = alpha * pro
            
            If alpha < 0 Then alpha = 0
            If alpha > 1 Then alpha = 1
            With b
                .AlphaFormat = &H1
                .BlendFlags = &H0
                .BlendOp = 0
                .SourceConstantAlpha = Int(alpha * 255)
            End With
            CopyMemory bl, b, 4
        End If
        
        If TypeName(N) = "String" Then
            index = Res.GetImage(N)
        Else
            index = N
        End If
        If index = 0 Then Exit Sub
        
        If IsMissing(cx) Then ocx = 0 Else: ocx = cx
        If IsMissing(cy) Then ocy = 0 Else: ocy = cy
        If IsMissing(cw) Then ocw = Res.ImgSize(index, imgGetWidth) - ocx Else: ocw = cw
        If IsMissing(ch) Then och = Res.ImgSize(index, imgGetHeight) - ocy Else: och = ch
        
        If Pos = 1 Then OX = Int(OX - ocw / 2): OY = Int(OY - och / 2)
        If Pos = 2 Then OX = OX - ocw
        If Pos = 3 Then OY = OY - och
        If Pos = 4 Then OX = OX + ocw
        If Pos = 5 Then OY = OY + och
        
        If ScrollMode Then
            OX = OX + ScrollX + ScrollBX: OY = OY + ScrollY + ScrollBY
            Dim Ret As Long, Crect As RECT
            With Crect
                .Left = OX: .top = OY: .Right = OX + ocw: .Bottom = OY + och
            End With
            If RectVisible(CDC, Crect) <> 1 Then OutOfScroll = True: Exit Sub
        End If
        
        Dim srcDC As Long
        If Res.ResourceGIFCount(index) = 0 Then
            srcDC = Res.ResourceHandle2(index, Direction)
        Else
            If GetTickCount - Res.ResourceGIFTime(index) >= GTick Then Res.GIFTick index
            srcDC = Res.ResourceGIFCurrentHandle(index)
        End If
        
        If IsMissing(alpha) Then
            BitBlt CDC, OX, OY, ocw, och, srcDC, ocx, ocy, vbSrcCopy
        Else
            AlphaBlend CDC, OX, OY, ocw, och, srcDC, ocx, ocy, ocw, och, bl
        End If

        With DrawF
            .X = OX
            .Y = OY
            .Width = ocw
            .Height = och
            .CrashIndex = Res.ResourceCrashIndex(index)
            .WSc = 1
            .HSc = 1
            .Shape = -1
        End With
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 0, 0, 255)
            GdipDrawRectangle GG, Pen, OX, OY, ocw - 1, och - 1
        End If
        
        If Debug_pos Then
            GdipSetPenColor Pen, argb(255, 0, 183, 195)
            GdipDrawLine GG, Pen, X - 5, Y, X + 5, Y
            GdipDrawLine GG, Pen, X, Y - 5, X, Y + 5
        End If
        
    End Sub
'========================================================
'   Writer
    Public Sub Writes(ByVal Text As String, X As Long, Y As Long, Optional Size As Long = 14, Optional Color As Long, Optional W As Long = 0, Optional h As Long = 0, Optional align As StringAlignment = StringAlignmentNear, Optional style As FontStyle = FontStyleRegular, Optional animation As Integer = 0)
        'If OutOfScroll Then Exit Sub
        
        Dim OX As Long, OY As Long
        OX = X: OY = Y
        
        If ScrollMode Then
            OX = OX + ScrollX + ScrollBX: OY = OY + ScrollY + ScrollBY
            Dim Ret As Long
            GdipIsVisibleRect GG, OX, OY, W, h, Ret
            If Ret = 0 Then OutOfScroll = True: Exit Sub
        End If
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single, co(3) As Byte, al As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during)
                CopyMemory co(0), Color, 4
                al = co(3): al = al / 255
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, OX, OY, W, h, al, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, OX, OY, W, h, al, pro
                End If
                Color = argb(al * 255, co(0), co(1), co(2))
            End If
        End If
        
        EF.Writes Text, OX, OY, GG, Color, Size, W, h, align, style
        With DrawF
            .X = OX
            .Y = OY
            .Width = W
            .Height = h
            .CrashIndex = 0
            .Shape = 0
        End With
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 0, 255, 0)
            GdipDrawRectangle GG, Pen, OX, OY, W - 1, h - 1
        End If
        
        If Debug_pos Then
            GdipSetPenColor Pen, argb(255, 0, 183, 195)
            GdipDrawLine GG, Pen, X - 5, Y, X + 5, Y
            GdipDrawLine GG, Pen, X, Y - 5, X, Y + 5
        End If
    End Sub
'========================================================
'   Shape
    Public Sub PaintLine(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Optional Color As Long, Optional Size As Long = 1, Optional animation As Integer = 0)
        If ScrollMode Then
            X = X + ScrollX + ScrollBX: Y = Y + ScrollY + ScrollBY
            Dim Ret As Long
            GdipIsVisibleRect GG, IIf(X1 < X2, X1, X2), IIf(Y1 < Y2, Y1, Y2), IIf(X1 < X2, X2 - X1, X1 - X2), IIf(Y1 < Y2, Y2 - Y1, Y1 - Y2), Ret
            If Not Ret Then OutOfScroll = True: Exit Sub
        End If
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single, co(3) As Byte, al As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during)
                CopyMemory co(0), Color, 4
                al = co(3): al = al / 255
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, X, Y, W, h, al, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, X, Y, W, h, al, pro
                End If
                Color = argb(al * 255, co(0), co(1), co(2))
            End If
        End If
        
        If Color <> 0 Then
            GdipSetPenColor Pen, Color
            GdipSetPenWidth Pen, Size
        End If
        
        GdipDrawLine GG, Pen, X1, Y1, X2, Y2
        
        With DrawF
            .X = IIf(X1 < X2, X1, X2)
            .Y = IIf(Y1 < Y2, Y1, Y2)
            .Width = IIf(X1 < X2, X2 - X1, X1 - X2)
            .Height = IIf(Y1 < Y2, Y2 - Y1, Y1 - Y2)
            .CrashIndex = 0
            .Shape = 0
        End With
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 240, 176, 0)
            GdipDrawRectangle GG, Pen, DrawF.X, DrawF.Y, DrawF.Width + 1, DrawF.Height + 1
        End If
    End Sub
    Public Sub Paint(ByVal Shape As Integer, X As Long, Y As Long, W As Long, h As Long, Optional Color As Long, Optional Radius As Long, Optional Size As Long = 1, Optional style As Integer = 0, Optional Pos As PosAlign = posNormal, Optional animation As Integer = 0)
        'shape:0=rect,1=ellipse,2=rectr
        'style:0=fill,1=border
        'If OutOfScroll Then Exit Sub

        Dim OX As Long, OY As Long, ow As Long, oh As Long
        OX = X: OY = Y: ow = W: oh = h
        
        If ScrollMode Then
            OX = OX + ScrollX + ScrollBX: OY = OY + ScrollY + ScrollBY
            Dim Ret As Long
            GdipIsVisibleRect GG, OX - 1, OY - 1, ow, oh, Ret
            If Not Ret Then OutOfScroll = True: Exit Sub
        End If
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single, co(3) As Byte, al As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during)
                CopyMemory co(0), Color, 4
                al = co(3): al = al / 255
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, OX, OY, ow, oh, al, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, OX, OY, ow, oh, al, pro
                End If
                Color = argb(al * 255, co(0), co(1), co(2))
            End If
        End If
        
        GdipResetPath Path
        
        If Pos = 1 Then OX = Int(OX - ow / 2): OY = Int(OY - oh / 2)
        If Pos = 2 Then OX = OX - ow
        If Pos = 3 Then OY = OY - oh
        If Pos = 4 Then OX = OX + ow
        If Pos = 5 Then OY = OY + oh
        
ReShape:
        If Shape = 0 Then GdipAddPathRectangle Path, OX, OY, ow - 1, oh - 1
        If Shape = 1 Then GdipAddPathEllipse Path, OX, OY, ow - 1, oh - 1
        If Shape = 2 Then
        
            If Radius = 0 Then
                Shape = 0: GoTo ReShape
            End If
            
            If Radius > ow Then Radius = ow
            If Radius > oh Then Radius = oh
            
            GdipAddPathArc Path, OX, OY, Radius, Radius, 180, 90
            GdipAddPathArc Path, OX + ow - Radius, OY, Radius, Radius, 270, 90
            GdipAddPathArc Path, OX + ow - Radius, OY + oh - Radius, Radius, Radius, 0, 90
            GdipAddPathArc Path, OX, OY + oh - Radius, Radius, Radius, 90, 90
            GdipClosePathFigure Path
            
        End If
        
        If Color <> 0 Then
            If style = 0 Then GdipSetSolidFillColor brush, Color
            If style = 1 Then
                GdipSetPenColor Pen, Color
                GdipSetPenWidth Pen, Size
            End If
        End If
        
        If Shape = 0 Then GdipSetSmoothingMode GG, SmoothingModeDefault
        If style = 0 Then GdipFillPath GG, brush, Path
        If style = 1 Then GdipDrawPath GG, Pen, Path
        If Shape = 0 Then GdipSetSmoothingMode GG, SmoothingModeAntiAlias
        
        With DrawF
            .X = OX
            .Y = OY
            .Width = ow
            .Height = oh
            .Shape = Shape
            .CrashIndex = 0
        End With
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 255, 0, 0)
            GdipDrawRectangle GG, Pen, OX, OY, ow - 1, oh - 1
        End If
        
        If Debug_pos Then
            GdipSetPenColor Pen, argb(255, 0, 183, 195)
            GdipDrawLine GG, Pen, X - 5, Y, X + 5, Y
            GdipDrawLine GG, Pen, X, Y - 5, X, Y + 5
        End If
    End Sub
    Public Sub PaintArc(X As Long, Y As Long, W As Long, h As Long, degree As Long, Optional Start As Long = 0, Optional Color As Long, Optional Size As Long = 1, Optional style As Integer = 0, Optional Pos As PosAlign = posNormal, Optional animation As Integer = 0)
        'style:0=fill,1=border,2=sector

        Dim OX As Long, OY As Long, ow As Long, oh As Long
        OX = X: OY = Y: ow = W: oh = h
        
        If ScrollMode Then
            OX = OX + ScrollX + ScrollBX: OY = OY + ScrollY + ScrollBY
            Dim Ret As Long
            GdipIsVisibleRect GG, OX - 1, OY - 1, ow, oh, Ret
            If Not Ret Then OutOfScroll = True: Exit Sub
        End If
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single, co(3) As Byte, al As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during)
                CopyMemory co(0), Color, 4
                al = co(3): al = al / 255
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, OX, OY, ow, oh, al, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, OX, OY, ow, oh, al, pro
                End If
                Color = argb(al * 255, co(0), co(1), co(2))
            End If
        End If
        
        GdipResetPath Path
        
        If Pos = 1 Then OX = Int(X - ow / 2): OY = Int(Y - oh / 2)
        If Pos = 2 Then OX = OX - ow
        If Pos = 3 Then OY = OY - oh
        If Pos = 4 Then OX = OX + ow
        If Pos = 5 Then OY = OY + oh
        
        GdipAddPathArc Path, OX, OY, ow, oh, Start, degree
        If style = 2 Then GdipAddPathLine Path, OX + ow / 2, OY + ow / 2, OX + ow / 2, OY + ow / 2
        If style = 0 Then GdipClosePathFigure Path
        
        If Color <> 0 Then
            If style = 0 Or style = 2 Then GdipSetSolidFillColor brush, Color
            If style = 1 Then
                GdipSetPenColor Pen, Color
                GdipSetPenWidth Pen, Size
            End If
        End If
        
        If style = 0 Or style = 2 Then GdipFillPath GG, brush, Path
        If style = 1 Then GdipDrawPath GG, Pen, Path
        
        With DrawF
            .X = OX
            .Y = OY
            .Width = ow
            .Height = oh
            .CrashIndex = 0
            .Shape = 0
        End With
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 0, 255, 255)
            GdipDrawRectangle GG, Pen, OX, OY, ow - 1, oh - 1
        End If
        
        If Debug_pos Then
            GdipSetPenColor Pen, argb(255, 0, 183, 195)
            GdipDrawLine GG, Pen, X - 5, Y, X + 5, Y
            GdipDrawLine GG, Pen, X, Y - 5, X, Y + 5
        End If
    End Sub
    Public Sub PaintPolygon2(Color As Long, style As Integer, Points(), Optional SmoothMode As Boolean = False)

        GdipResetPath Path
        Dim p() As POINTF
        ReDim p((UBound(Points) - 1) / 2)
        For i = 0 To UBound(Points) Step 2
            p(i / 2).X = Points(i): p(i / 2).Y = Points(i + 1)
        Next
        
        If SmoothMode Then
            GdipAddPathCurve Path, p(0), UBound(p) + 1
        Else
            GdipAddPathPolygon Path, p(0), UBound(p) + 1
        End If
        
        If style = 0 Then GdipSetSolidFillColor brush, Color
        If style = 1 Then
            GdipSetPenColor Pen, Color
            GdipSetPenWidth Pen, Size
        End If
        
        If style = 0 Then GdipFillPath GG, brush, Path
        If style = 1 Then GdipDrawPath GG, Pen, Path
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 0, 176, 240)
            GdipDrawPath GG, Pen, Path
        End If
    End Sub
    Public Sub PaintPolygon(Color As Long, style As Integer, SmoothMode As Boolean, ParamArray Points())

        GdipResetPath Path
        Dim p() As POINTF
        ReDim p((UBound(Points) - 1) / 2)
        For i = 0 To UBound(Points) Step 2
            p(i / 2).X = Points(i): p(i / 2).Y = Points(i + 1)
        Next
        
        If SmoothMode Then
            GdipAddPathCurve Path, p(0), UBound(p) + 1
        Else
            GdipAddPathPolygon Path, p(0), UBound(p) + 1
        End If
        
        If style = 0 Then GdipSetSolidFillColor brush, Color
        If style = 1 Then
            GdipSetPenColor Pen, Color
            GdipSetPenWidth Pen, Size
        End If
        
        If style = 0 Then GdipFillPath GG, brush, Path
        If style = 1 Then GdipDrawPath GG, Pen, Path
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 0, 176, 240)
            GdipDrawPath GG, Pen, Path
        End If
    End Sub
'========================================================
'   Scroll
    Public Sub StartScroll(X As Long, Y As Long, W As Long, h As Long, sx As Long, sy As Long)
        OutOfScroll = False
        ScrollMode = True
        ScrollX = sx: ScrollY = sy: ScrollBX = X: ScrollBY = Y
        GdipSetClipRect GG, X, Y, W, h, CombineModeReplace
        CRgn = CreateRectRgn(X, Y, X + W, Y + h)
        SelectClipRgn CDC, CRgn
    End Sub
    Public Sub EndScroll()
        OutOfScroll = False
        ScrollMode = False
        GdipResetClip GG
        DeleteObject CRgn
        SelectClipRgn CDC, ByVal 0
    End Sub
'========================================================
    'Animation
    Public Function NewAnimation(custom As Boolean, func As String, profunc As Integer, during As Long) As Integer
        ReDim Preserve Anis(UBound(Anis) + 1)
        With Anis(UBound(Anis))
            .custom = custom
            .during = during
            .func = func
            .profunc = profunc
            .mark = True
        End With
        
        NewAnimation = UBound(Anis)
    End Function
    Public Sub StartAnimation(id As Integer, Optional delay As Long = 0)
        Anis(id).Start = GetTickCount
        Anis(id).delay = delay
        Anis(id).mark = False
    End Sub
    Public Function AnimationDone(id As Integer) As Boolean
        AnimationDone = Anis(id).mark
    End Function
'========================================================
    'Control
    Public Function ShowEdit(Text As String, Shape As Integer, X As Long, Y As Long, W As Long, h As Long, TextColor As Long, Color As Long, HoverColor As Long, LineColor As Long, Optional Radius As Long = 0, Optional ShapeStyle As Integer = 0, Optional Size As Long = 14, Optional style As FontStyle = FontStyleRegular) As Integer
        Dim m As Integer, r As RECT
        
        m = CheckMouse(X, Y, W - IIf(TextHandle = VarPtr(Text), h, 0), h)
        Paint Shape, X, Y, W, h, IIf(m <> 0 Or TextHandle = VarPtr(Text), HoverColor, Color), Radius, style:=ShapeStyle
        If m = 3 Then TextHandle = VarPtr(Text)
        
        Dim CtrlPressed As Boolean
        CtrlPressed = (GetAsyncKeyState(vk_ctrl) <> 0)
        
        If VarPtr(Text) = TextHandle And WaitChr <> "" Then
            If WaitChr = Chr(vbKeyReturn) Then
                TextHandle = 0
                m = 4
            ElseIf WaitChr = Chr(vbKeyBack) Then
                If Len(Text) > 0 Then Text = Left(Text, Len(Text) - 1)
            ElseIf Asc(WaitChr) = 22 Then '粘贴
                On Error Resume Next
                Text = Text & Clipboard.GetText
            ElseIf Asc(WaitChr) = 3 Then '复制
                On Error Resume Next
                Clipboard.Clear
                Clipboard.SetText Text
            ElseIf Asc(WaitChr) = 24 Then '剪切
                On Error Resume Next
                Clipboard.Clear
                Clipboard.SetText Text
                Text = ""
            ElseIf Asc(WaitChr) <= 26 And Asc(WaitChr) >= 0 Then '其他Ctrl组合键
                
            Else
                Text = Text & WaitChr
            End If
            WaitChr = ""
        End If
        
        Writes Text, X + Radius / 4, Y + h / 2 - Size / 0.75 / 2 - 1, Size, TextColor, W - Radius / 2 - h, Size / 0.75, StringAlignmentNear, style
        
        If TextHandle = VarPtr(Text) Then
            Dim w2 As Long, pro As Long, alpha As Single
            w2 = EF.GetWidth(GG, Text, Size, StringAlignmentNear, style)
            If w2 > W - Radius / 2 - h Then
                If Len(Text) > 0 Then Text = Left(Text, Len(Text) - 1): VBA.Beep
            End If
            Paint Shape, X, Y, W, h, LineColor, Size:=2, Radius:=Radius, style:=1
            pro = GetTickCount Mod 1000
            If pro <= 700 Then
                alpha = 1 - Cubic(pro / 700, 0, 1, 1, 1)
            Else
                alpha = Cubic((pro - 700) / 300, 0, 1, 1, 1)
            End If
            Dim co(3) As Byte, co2 As Long
            co2 = IIf(m <> 0 Or TextHandle = VarPtr(Text), HoverColor, Color)
            CopyMemory co(0), co2, 4
            If w2 = 0 Then w2 = Size / 4
            If alpha <> 0 Then Paint 0, X + w2 + Radius / 4 - Size / 8, Y + h / 2 - Size / 2, 3, Size, argb(Int(alpha * 255), 255 - co(2), 255 - co(1), 255 - co(0))
            
            Paint Shape, X + W - h, Y, h, h, LineColor, Radius:=Radius
            If CheckMouse2 = mMouseUp Then
                TextHandle = 0
                m = 4
            End If
            
            If Shape = 2 Then
                Writes ">", X + W - h, Y + h / 2 - Size / 0.75 / 2 - 1, Size, TextColor, h + 6, Size / 0.75, StringAlignmentCenter, FontStyleBold
            Else
                Writes ">", X + W - h, Y + h / 2 - Size / 0.75 / 2 - 2, Size, TextColor, h + 3, Size / 0.75, StringAlignmentCenter, FontStyleBold
            End If
        End If
        
        If Mouse.state = 2 And m = 0 Then TextHandle = 0: m = 4
        
        ShowEdit = m
        
    End Function
    Public Sub ShowLoading(X As Long, Y As Long, W As Long, h As Long, Size As Long, color1 As Long, color2 As Long, color3 As Long)
        
        PaintArc X, Y, W, h, 120, (GetTickCount Mod 3000) / 3000 * 360, color3, Size * 0.86, 1
        PaintArc X, Y, W, h, 120, (GetTickCount Mod 1800) / 1800 * 360 + 90, color2, Size * 0.93, 1
        PaintArc X, Y, W, h, 120, (GetTickCount Mod 1200) / 1200 * 360 + 90, color1, Size, 1
        'argb(255, 242, 195, 17),argb(255, 240, 127, 34),argb(255, 232, 76, 61)
    End Sub
    Public Function ShowButton(Pic As String, X As Long, Y As Long, Text As String, Color As Long, Optional Size As Long = 14, Optional style As FontStyle = FontStyleRegular) As Integer
        Dim W As Long, h As Long, i As Integer, m As Integer
        i = Res.GetImage(Pic)
        W = Res.ImgSize(i, imgGetWidth): h = Res.ImgSize(i, imgGetHeight)
        
        m = CheckMouse(X, Y, W, Int(h / 2))
        DrawImage i, X, Y, cy:=IIf(m <> 0, Int(h / 2), 0), ch:=Int(h / 2), alpha:=1
        
        Writes Text, X, Y + h / 2 / 2 - Size / 0.75 / 2, Size, Color, W, Size / 0.75, StringAlignmentCenter, style
        
        ShowButton = m
    End Function
    Public Function ShowSimpleButton(Pic As String, X As Long, Y As Long, Text As String, Color As Long, Optional Size As Long = 14, Optional style As FontStyle = FontStyleRegular) As Integer
        Dim W As Long, h As Long, i As Integer, m As Integer
        i = Res.GetImage(Pic)
        W = Res.ImgSize(i, imgGetWidth): h = Res.ImgSize(i, imgGetHeight)
        
        m = CheckMouse(X, Y, W, h)
        DrawImage i, X, Y, alpha:=IIf(m, 1, 0.8)
        
        Writes Text, X, Y + h / 2 - Size / 0.75 / 2, Size, Color, W, Size / 0.75, StringAlignmentCenter, style
    
        ShowSimpleButton = m
    End Function
    Public Function ShowColorButton(Shape As Integer, X As Long, Y As Long, W As Long, h As Long, Text As String, TextColor As Long, Color As Long, HoverColor As Long, Optional Radius As Long = 0, Optional ShapeStyle As Integer = 0, Optional Size As Long = 14, Optional style As FontStyle = FontStyleRegular) As Integer
        Dim m As Integer
        
        m = CheckMouse(X, Y, W, h)
        Paint Shape, X, Y, W, h, IIf(m, HoverColor, Color), Radius, style:=ShapeStyle
        
        Writes Text, X, Y + h / 2 - Size / 0.75 / 2, Size, TextColor, W, Size / 0.75, StringAlignmentCenter, style
    
        ShowColorButton = m
    End Function
    Public Function ShowCheckBox(value, Pic As String, X As Long, Y As Long, Text As String, Color As Long, Optional Size As Long = 14, Optional style As FontStyle = FontStyleRegular) As Integer
        Dim W As Long, h As Long, i As Integer, m As Integer
        Dim X2 As Long, Y2 As Long
        X2 = X: Y2 = Y
        
        i = Res.GetImage(Pic)
        W = Res.ImgSize(i, imgGetWidth): h = Res.ImgSize(i, imgGetHeight)
        
        m = CheckMouse(X, Y, W, Int(h / 2))
        DrawImage i, X, Y, cy:=IIf(value = True, Int(h / 2), 0), ch:=Int(h / 2), alpha:=1
        If m = 3 Then value = IIf(value = True, False, True)
        
        Writes Text, X2 + W + Size / 0.75 / 2, Y2 + h / 2 / 2 - Size / 0.75 / 2, Size, Color, GW, Size / 0.75, StringAlignmentNear, style
        
        ShowCheckBox = m
    End Function
    Public Function ShowColorCheckBox(value, X As Long, Y As Long, W As Long, h As Long, Text As String, DefaultColor As Long, ChooseColor As Long, ForeColor As Long, Optional Size As Long = 14, Optional style As FontStyle = FontStyleRegular) As Integer
        m = CheckMouse(X, Y, W, h)
        
        Paint 1, X, Y, h, h, IIf(value = True, ChooseColor, DefaultColor)
        If m = 3 Then value = IIf(value = True, False, True)
        
        Writes Text, X + h + Size / 0.75 / 2, Y + h / 2 - Size / 0.75 / 2, Size, ForeColor, GW, Size / 0.75, StringAlignmentNear, style
        
        ShowColorCheckBox = m
    End Function
    Public Function ShowSliderBar(value, Pic As String, X As Long, Y As Long) As Integer
        Dim W As Long, h As Long, i As Integer, m As Integer
        Dim X2 As Long, Y2 As Long
        
        X2 = X: Y2 = Y
        
        i = Res.GetImage(Pic)
        W = Res.ImgSize(i, imgGetWidth): h = Res.ImgSize(i, imgGetHeight)
        
        m = CheckMouse(X, Y, W, Int(h / 2))
        DrawImage i, X, Y, cy:=0, ch:=Int(h / 2), alpha:=1
        DrawImage i, X, Y, cy:=Int(h / 2), cw:=value * W, ch:=Int(h / 2), alpha:=1
        
        If m > 0 And Mouse.button > 0 Then
            value = (Mouse.X - X) / W
            If value < 0 Then value = 0
            If value > 1 Then value = 1
        End If
        
        ShowSliderBar = m
    End Function
    Public Function ShowHScrollBar(value, X As Long, Y As Long, W As Long, h As Long, LineColor As Long, BallColor As Long, Optional Size As Long = 2) As Integer
        Dim m As Integer
        
        m = CheckMouse(X, Y, W, h)
        PaintLine X, Y + h / 2 - Size / 2, X + W, Y + h / 2 - Size / 2, LineColor, Size
        Paint 1, X + value * (W - h), Y, h, h, BallColor
        If m > 0 And Mouse.button > 0 Then
            value = (Mouse.X - X) / (W - h)
            If value < 0 Then value = 0
            If value > 1 Then value = 1
        End If
        
        ShowHScrollBar = m
    End Function
    Public Function ShowVScrollBar(value, X As Long, Y As Long, W As Long, h As Long, LineColor As Long, BallColor As Long, Optional Size As Long = 2) As Integer
        Dim m As Integer
        
        m = CheckMouse(X, Y, W, h)
        PaintLine X + W / 2 - Size / 2, Y, X + W / 2 - Size / 2, Y + h, LineColor, Size
        Paint 1, X, Y + value * (h - W), W, W, BallColor
        
        If m > 0 And Mouse.button > 0 Then
            value = (Mouse.Y - Y) / (h - W)
            If value < 0 Then value = 0
            If value > 1 Then value = 1
        End If
        
        ShowVScrollBar = m
    End Function
'========================================================

