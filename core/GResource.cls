VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As clsid) As Long
Public HotLoad As Boolean
Dim AssetsFolder() As String, FolderParam1() As Variant, FolderParam2() As Variant
Public SearchIndex As Long
Public NoLoadCrash As Boolean, NoLoadDirection As Boolean
Dim Mems() As GMem
Dim Effect As Long, effectw As Long, effecth As Long, effectimg As Long, effectindex As Integer
Dim DisposeMark As Boolean
Private Sub Class_Initialize()
    ReDim Mems(0)
    ReDim AssetsFolder(0): ReDim FolderParam1(0): ReDim FolderParam2(0)
    SearchIndex = 1
    NoLoadDirection = True
    NoLoadCrash = True
    'HotLoad = True
End Sub
Public Sub ClipCircle(N)
    Dim index As Integer
    
    If TypeName(N) = "String" Then
        index = GetImage(N)
    Else
        index = N
    End If

    Dim Path As Long, bmpGraph As Long, BMP As Long, BMP2 As Long, tempGraph As Long
    Dim W As Long, h As Long
    
    GdipGetImageWidth Mems(index).ImgHwnd, W
    GdipGetImageHeight Mems(index).ImgHwnd, h

    GdipCreateBitmapFromScan0 W, h, ByVal 0, PixelFormat32bppARGB, ByVal 0, BMP2
    GdipCreateBitmapFromScan0 W * 2, h * 2, ByVal 0, PixelFormat32bppARGB, ByVal 0, BMP

    GdipGetImageGraphicsContext BMP2, tempGraph
    GdipGetImageGraphicsContext BMP, bmpGraph
    GdipSetSmoothingMode tempGraph, SmoothingModeAntiAlias
    GdipSetSmoothingMode bmpGraph, SmoothingModeAntiAlias
    
    PoolCreatePath FillModeWinding, Path
    GdipAddPathEllipseI Path, 1, 1, W * 2 - 2, h * 2 - 2
    GdipSetClipPath bmpGraph, Path, CombineModeReplace
    
    GdipDrawImageRectI bmpGraph, Mems(index).ImgHwnd, 0, 0, W * 2, h * 2
    GdipResetClip bmpGraph
    
    GdipSetInterpolationMode tempGraph, InterpolationModeHighQualityBilinear
    
    GdipDrawImageRectI tempGraph, BMP, 0, 0, W - 1, h - 1

    GdipDeletePath Path
    PoolDeleteGraphics bmpGraph
    PoolDisposeImage BMP
    PoolDeleteGraphics tempGraph
    PoolDisposeImage Mems(index).ImgHwnd
    
    Mems(index).ImgHwnd = BMP2
    
    Call ReDraw(index)
End Sub
Private Sub ApplyEffect(N, Kind As GdipEffectType)
    Dim index As Integer
    
    If TypeName(N) = "String" Then
        index = GetImage(N)
    Else
        index = N
    End If
    
    effectimg = Mems(index).ImgHwnd
    
    PoolCreateEffect2 Kind, Effect
    
    effectindex = index
End Sub
Private Sub ApplyEffectStep2()
    GdipGetImageWidth effectimg, effectw: GdipGetImageHeight effectimg, effecth
    GdipBitmapApplyEffect effectimg, Effect, NewRectL(0, 0, effectw, effecth), 0, 0, 0
    
    PoolDeleteEffect Effect '垃圾处理
    
    Call ReDraw(effectindex)
End Sub
Public Sub ApplyTintEffect(N, amount As Long, hue As Long)
    Dim e As TintParams
    e.amount = amount
    e.hue = hue
    
    ApplyEffect N, Tint
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplySharpenEffect(N, amount As Single, Radius As Single)
    Dim e As SharpenParams
    e.amount = amount
    e.Radius = Radius

    ApplyEffect N, Sharpen
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyBlurEffect(N, Radius As Long, edge As Integer)
    Dim e As BlurParams
    e.Radius = Radius
    e.expandEdge = edge

    ApplyEffect N, Blur
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyBrightnessContrastEffect(N, brightnessLevel As Long, contrastLevel As Long)
    Dim e As BrightnessContrastParams
    e.brightnessLevel = brightnessLevel
    e.contrastLevel = contrastLevel

    ApplyEffect N, BrightnessContrast
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyColorBalanceEffect(N, cyanRed As Long, magentaGreen As Long, yellowBlue As Long)
    Dim e As ColorBalanceParams
    e.cyanRed = cyanRed
    e.magentaGreen = magentaGreen
    e.yellowBlue = YellowGreen

    ApplyEffect N, ColorBalance
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyColorCurveEffect(N, adjustment As CurveAdjustments, adjustValue As Long, channel As CurveChannel)
    Dim e As ColorCurveParams
    e.adjustment = adjustment
    e.adjustValue = adjustValue
    e.channel = channel

    ApplyEffect N, ColorCurve
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyColorMatrixEffect(N, m() As Single)
    Dim e As ColorMatrix
    For i = 0 To 4
        For S = 0 To 4
            e.m(i, S) = m(i, S)
        Next
    Next
    ApplyEffect N, ColorMatrix
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyHueSaturationLightnessEffect(N, hueLevel As Long, lightnessLevel As Long, saturationLevel As Long)
    Dim e As HueSaturationLightnessParams
    e.hueLevel = hueLevel
    e.lightnessLevel = lightnessLevel
    e.saturationLevel = saturationLevel

    ApplyEffect N, HueSaturationLightness
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyLevelsEffect(N, highlight As Long, midtone As Long, Shadow As Long)
    Dim e As LevelsParams
    e.highlight = highlight
    e.midtone = midtone
    e.Shadow = Shadow

    ApplyEffect N, levels
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Private Sub ReDraw(i As Integer)
    If Mems(i).GIF.Count > 0 Then Exit Sub
    
    Dim g As Long
    PoolCreateFromHdc Mems(i).Hwnd, g
    GdipSetSmoothingMode g, SmoothingModeAntiAlias
    GdipGraphicsClear g, 0
    GdipDrawImageRect g, Mems(i).ImgHwnd, 0, 0, Mems(i).W, Mems(i).h
    PoolDeleteGraphics g
    
    For S = 0 To 3
        PoolCreateFromHdc Mems(i).Imgs(S), g
        GdipSetSmoothingMode g, SmoothingModeAntiAlias
        GdipGraphicsClear g, 0
        GdipDrawImageRect g, Mems(i).ImgHwnd, IIf(S = 1 Or S = 3, Mems(i).W, 0), IIf(S = 2 Or S = 3, Mems(i).h, 0), Mems(i).W * IIf(S = 1 Or S = 3, -1, 1), Mems(i).h * IIf(S = 2 Or S = 3, -1, 1)
        PoolDeleteGraphics g
    Next
    
    Dim tempCo As Long, tCo(3) As Byte
    Dim IW As Long, iH As Long

    ReDim Preserve ColorLists(UBound(ColorLists) + 1)
    ReDim ColorLists(UBound(ColorLists)).IsAlpha(Mems(i).W + 1, Mems(i).h + 1)
    
    DC = Mems(i).Imgs(0)
    GdipGetImageWidth Mems(i).ImgHwnd, IW: GdipGetImageHeight Mems(i).ImgHwnd, iH
    For X = 0 To Mems(i).W - 1
        For y = 0 To Mems(i).h - 1
            '取得像素点颜色
            GdipBitmapGetPixel Mems(i).ImgHwnd, X / (Mems(i).W - 1) * IW, y / (Mems(i).h - 1) * iH, tempCo
            CopyMemory tCo(0), tempCo, 4
            
            If tCo(3) = 255 Then ColorLists(UBound(ColorLists)).IsAlpha(X, y) = True
        Next
        DoEvents
    Next
    
    Mems(i).CrashIndex = UBound(ColorLists)
End Sub
Public Sub Add(Mem As Long, Kind As Integer, Name As String, Optional W, Optional h, Optional Folder As String = "")
    ReDim Preserve Mems(UBound(Mems) + 1)
    With Mems(UBound(Mems))
        .Hwnd = Mem
        .Kind = Kind
        .Name = Name
        .Folder = Folder
        If Not IsMissing(W) Then .W = W
        If Not IsMissing(h) Then .h = h
    End With
End Sub
Public Property Get ImgCount() As Long
    ImgCount = UBound(Mems)
End Property
Public Property Get ImgSize(N, sizeIndex As imgIndex)
    Dim index As Integer
    
    If TypeName(N) = "String" Then
        index = GetImage(N)
    Else
        index = N
    End If

    Select Case sizeIndex
        Case 0
            ImgSize = Mems(index).W
        Case 1
            ImgSize = Mems(index).h
        Case 2
            ImgSize = Mems(index).GIF.Count
    End Select
End Property
Public Sub Dispose()
    If DisposeMark Then Exit Sub
    Exit Sub
    For i = 1 To UBound(Mems)
        If Not Mems(i).Copyed Then
            Select Case Mems(i).Kind
                Case 0 'DC
                    If Mems(i).Hwnd <> 0 Then DeleteObject Mems(i).Hwnd
                    If Mems(i).ImgHwnd <> 0 Then PoolDisposeImage Mems(i).ImgHwnd
                    For S = 0 To 3
                        If Mems(i).Imgs(S) <> 0 Then DeleteObject Mems(i).Imgs(S)
                    Next
                Case 1 'Graphics
                    PoolDeleteGraphics Mems(i).Hwnd
            End Select
            If Mems(i).GIF.Count > 0 Then
                For S = 0 To Mems(i).GIF.Count
                    DeleteObject Mems(i).GIF.frames(S)
                Next
            End If
        End If
    Next
    
    DisposeMark = True
End Sub
Public Property Get ResourceGIFHandle(index As Integer, frame As Integer) As Long
    ResourceGIFHandle = Mems(index).GIF.frames(frame)
End Property
Public Property Get ResourceGIFCurrentHandle(index As Integer) As Long
    ResourceGIFCurrentHandle = Mems(index).GIF.frames(Mems(index).GIF.tick)
End Property
Public Property Get ResourceGIFTick(index As Integer) As Long
    ResourceGIFTick = Mems(index).GIF.tick
End Property
Public Sub GIFTick(index As Integer)
    Mems(index).GIF.tick = Mems(index).GIF.tick + 1
    If Mems(index).GIF.tick > Mems(index).GIF.Count Then Mems(index).GIF.tick = 0
End Sub
Public Property Get ResourceGIFCount(index As Integer) As Long
    ResourceGIFCount = Mems(index).GIF.Count
End Property
Public Property Get ResourceGIFTime(index As Integer) As Long
    ResourceGIFTime = Mems(index).GIF.Time
End Property
Public Property Get ResourceName(index As Integer) As Long
    ResourceName = Mems(index).Name
End Property
Public Property Get ResourceHandle(index As Integer) As Long
    ResourceHandle = Mems(index).Hwnd
End Property
Public Property Get ResourceCrashIndex(index As Integer) As Long
    ResourceCrashIndex = Mems(index).CrashIndex
End Property
Public Property Get ResourceHandle2(index As Integer, Direction As ImgDirection) As Long
    ResourceHandle2 = Mems(index).Imgs(Direction)
End Property
Public Property Get ResourceImageHandle(index As Integer) As Long
    ResourceImageHandle = Mems(index).ImgHwnd
End Property
Public Function GetImage(ByVal Name As String) As Integer
    For i = SearchIndex To UBound(Mems)
        If Mems(i).Kind = 0 Then
            If Mems(i).Name = Name Then GetImage = i: Exit For
        End If
    Next
    
    If HotLoad Then
        If GetImage = 0 Then
            For i = 1 To UBound(AssetsFolder)
                If IsExitAFile(AssetsFolder(i) & Name) Then
                    Dim Mem As GMem
                    Mem = FindAssetsFromTree(AssetsFolder(i) & Name, FolderParam1(i), FolderParam2(i))
                    If Mem.Name <> "" Then
                        ReLoadCount = ReLoadCount + 1
                        Mem.Copyed = True
                        ReDim Preserve Mems(UBound(Mems) + 1)
                        Mems(UBound(Mems)) = Mem
                    Else
                        PreLoadCount = PreLoadCount + 1
                        LoadedCount = LoadedCount + 1
                        newImage AssetsFolder(i) & Name, FolderParam1(i), FolderParam2(i)
                        UpdateAssetsTree AssetsFolder(i), FolderParam1(i), FolderParam2(i), Mems(UBound(Mems))
                    End If
                    GetImage = UBound(Mems)
                    Exit For
                End If
            Next
        End If
    End If
End Function
Public Sub newImage(Path As String, Optional arg1, Optional arg2, Optional Name As String = "")
    'when arg1 is missing ; i.w = i.w
    'when arg2 is missing but arg1 has content ; i.w = i.w * arg1
    'other ; i.w = arg1
    DoEvents
    CrashPro = 1
    
    Dim si As Integer, i As Long, W As Long, h As Long, DC As Long, g As Long
    If IsMissing(arg1) Then si = 0
    If IsMissing(arg2) And (Not IsMissing(arg1)) Then si = 1
    If (Not IsMissing(arg2)) And (Not IsMissing(arg1)) Then si = 2
    
    PoolCreateBitmapFromFile StrPtr(Path), i
    GdipGetImageWidth i, W: GdipGetImageHeight i, h
    
    If si = 1 Then W = Round(W * arg1): h = Round(h * arg1)
    If si = 2 Then W = arg1: h = arg2
    
    Dim gC As clsid, gL As Long, index As Long, temp() As String, CrashPath As Long
    CLSIDFromString StrPtr("{6AEDBD6D-3FB5-418A-83A6-7F45229DC872}"), gC
    GdipImageGetFrameCount i, gC, gL
    
    If gL > 0 Then  '动态GIF支持
        temp = Split(Path, "\")
        Add 0, 0, temp(UBound(temp)), W, h, temp(UBound(temp) - 1)
        With Mems(UBound(Mems))
            .GIF.Count = gL
            ReDim .GIF.frames(gL)
            For index = 0 To gL
                GdipImageSelectActiveFrame i, gC, index
                DC = CreateCDC(W, h)
                PoolCreateFromHdc DC, g
                GdipDrawImageRect g, i, 0, 0, W, h
                PoolDeleteGraphics g
                .GIF.frames(index) = DC
            Next
        End With
        
        Mems(UBound(Mems)).ImgHwnd = i
        Exit Sub
    End If
    
    Dim Ccx As Long, Ccy As Long
    If Path Like "*bm-*.png" Then Ccx = 707: Ccy = 375: W = 355: h = 825
    If Path Like "*kx1-*.png" Then Ccx = 1096: Ccy = 220: W = 389: h = 389
    If Path Like "*kx2-*.png" Then Ccx = 444: Ccy = 281: W = 264: h = 213
    If Path Like "*xl-*.png" Then Ccx = 336: Ccy = 740: W = 522: h = 411
    If Path Like "*xx-*.png" Then Ccx = 1115: Ccy = 463: W = 379: h = 751
    If Path Like "*km1-*.png" Then Ccx = 74: Ccy = 267: W = 245: h = 386
    If Path Like "*fj-*.png" Then Ccx = 977: Ccy = 869: W = 220: h = 176
    If Path Like "*bg-*.png" Then Ccx = 195: Ccy = 879: W = 111: h = 253
    If Path Like "*jy-*.png" Then Ccx = 227: Ccy = 325: W = 247: h = 861
    If Path Like "*yz-*.png" Then Ccx = 101: Ccy = 598: W = 265: h = 568
    If Path Like "*ssr-*.png" Then Ccx = 935: Ccy = 379: W = 353: h = 651
    If Path Like "*tk-*.png" Then Ccx = 1206: Ccy = 472: W = 394: h = 347
    If Path Like "*yy-*.png" Then Ccx = 1305: Ccy = 722: W = 295: h = 452
    If Path Like "*km2-*.png" Then Ccx = 469: Ccy = 297: W = 785: h = 807
    
    If Path Like "*dn-*.png" Then Ccx = 1348: Ccy = 272: W = 252: h = 177
    If Path Like "*qing-*.png" Then Ccx = 449: Ccy = 411: W = 365: h = 789
    If Path Like "*s-ban.png" Then Ccx = 717: Ccy = 313: W = 573: h = 845
    If Path Like "*s-angry.png" Then Ccx = 717: Ccy = 313: W = 573: h = 845
    If Path Like "*s-normal.png" Then Ccx = 871: Ccy = 421: W = 289: h = 729
    If Path Like "*s-serious.png" Then Ccx = 871: Ccy = 421: W = 289: h = 729
    
    If Ccx <> 0 Then
        If (Not IsMissing(arg1)) And IsMissing(arg2) Then Ccx = Ccx * arg1: Ccy = Ccy * arg1: W = W * arg1: h = h * arg1
        Dim MaxW As Long, MaxH As Long
        MaxW = 1600: MaxH = 1200
        If (Not IsMissing(arg1)) And IsMissing(arg2) Then MaxW = MaxW * arg1: MaxH = MaxH * arg1
    End If
    
    temp = Split(Path, "\")
    ReDim Preserve Mems(UBound(Mems) + 1)
    With Mems(UBound(Mems))
        .Name = IIf(Name = "", temp(UBound(temp)), Name)
        .Folder = temp(UBound(temp) - 1)
        .W = W
        .h = h
        .ImgHwnd = i
    End With
    
    For S = 0 To 3
        If NoLoadDirection And S > 0 Then Exit For
        DC = CreateCDC(W, h)
        PoolCreateFromHdc DC, g
        If Ccx <> 0 Then
            GdipDrawImageRect g, i, -Ccx, -Ccy, MaxW, MaxH
        Else
            GdipDrawImageRect g, i, IIf(S = 1 Or S = 3, W, 0), IIf(S = 2 Or S = 3, h, 0), W * IIf(S = 1 Or S = 3, -1, 1), h * IIf(S = 2 Or S = 3, -1, 1)
        End If
        PoolDeleteGraphics g
        Mems(UBound(Mems)).Imgs(S) = DC
    Next
    
    If W > GW * 0.5 Or h > GH * 0.5 Or NoLoadCrash Or Ccx <> 0 Then Exit Sub
    
    Dim tempCo As Long, tCo(3) As Byte
    Dim IW As Long, iH As Long

    ReDim Preserve ColorLists(UBound(ColorLists) + 1)
    ReDim ColorLists(UBound(ColorLists)).IsAlpha(W + 1, h + 1)
    
    DC = Mems(UBound(Mems)).Imgs(0)
    GdipGetImageWidth i, IW: GdipGetImageHeight i, iH
    For X = 0 To W - 1
        For y = 0 To h - 1
            '取得像素点颜色
            GdipBitmapGetPixel i, X / (W - 1) * IW, y / (h - 1) * iH, tempCo
            CopyMemory tCo(0), tempCo, 4

            If tCo(3) = 255 Then ColorLists(UBound(ColorLists)).IsAlpha(X, y) = True
        Next
        CrashPro = X / (W - 1)
        DoEvents
    Next
    CrashPro = 1
    
    Mems(UBound(Mems)).CrashIndex = UBound(ColorLists)
End Sub
Public Sub NewImages(Path As String, Optional arg1, Optional arg2)

    Dim index As Integer, Tree As AssetsTree

    If HotLoad Then
        ReDim Preserve AssetsFolder(UBound(AssetsFolder) + 1)
        ReDim Preserve FolderParam1(UBound(FolderParam1) + 1)
        ReDim Preserve FolderParam2(UBound(FolderParam2) + 1)
        If Right(Path, 1) <> "\" Then Path = Path & "\"
        AssetsFolder(UBound(AssetsFolder)) = Path
        FolderParam1(UBound(FolderParam1)) = arg1
        FolderParam2(UBound(FolderParam2)) = arg2
        
        Tree.Path = Path: ReDim Tree.Files(0)
        AddAssetsTree Tree, arg1, arg2
        Exit Sub
    End If
    
    Dim f As String, C As Long
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    
    Tree.Path = Path: ReDim Tree.Files(0)
    index = FindAssetsTree(Path, arg1, arg2)
    
    f = Dir(Path)
    Do While f <> ""
        C = C + 1
        ReDim Preserve Tree.Files(UBound(Tree.Files) + 1)
        f = Dir()
    Loop
    If index = 0 Then
        PreLoadCount = PreLoadCount + C
    Else
        Tree = AssetsTrees(index)
        For i = 1 To UBound(Tree.Files)
            ReDim Preserve Mems(UBound(Mems) + 1)
            Mems(UBound(Mems)) = Tree.Files(i)
            Mems(UBound(Mems)).Copyed = True
        Next
        ReLoadCount = ReLoadCount + UBound(Tree.Files)
        Exit Sub
    End If
    
    f = Dir(Path)
    i = 1
    Do While f <> ""
        If IsMissing(arg1) Then
            newImage Path & f
        ElseIf IsMissing(arg2) Then
            newImage Path & f, arg1
        Else
            newImage Path & f, arg1, arg2
        End If
        Tree.Files(i) = Mems(UBound(Mems))
        i = i + 1
        f = Dir()
        DoEvents
        LoadedCount = LoadedCount + 1
    Loop
    
    If index = 0 Then AddAssetsTree Tree, arg1, arg2
End Sub
