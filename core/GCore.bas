Attribute VB_Name = "GCore"
'========================================================
'   DPI��Ӧ
    Public Declare Function SetProcessDpiAwareness Lib "SHCORE.DLL" (ByVal DPImodel As Long) As Long
'=========================================================================
    Private Declare Sub AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal hHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BLENDFUNCTION As Long) ' As Long
    Public Type MState
        state As Integer
        button As Integer
        x As Single
        y As Single
    End Type
    Public Enum PlayStateMark
        musStopped = 0
        musPlaying = 1
        musStalled = 2
        musPaused = 3
    End Enum
    Public Enum imgIndex
        imgGetWidth = 0
        imgGetHeight = 1
        imgGetGIFFrameCount = 2
    End Enum
    Public Enum MButtonState
        mMouseOut = 0
        mMouseIn = 1
        mMouseDown = 2
        mMouseUp = 3
    End Enum
    Public Enum PosAlign
        posNormal = 0
        posOnCenter = 1
        posOnLeft = 4
        posOnTop = 5
        posOnRight = 2
        posOnBottom = 3
    End Enum
    Public Enum TranslationKind
        transFadeIn = 0
        transFadeOut = 1
        transToRight = 2
        transToLeft = 3
        transToUp = 4
        transToDown = 5
        transToRightFade = 6
        transToLeftFade = 7
        transToUpFade = 8
        transToDownFade = 9
        transHighLight = 10
        transFallDark = 11
    End Enum
    Public Type GGIF
        time As Long
        frames() As Long
        tick As Long
        Count As Long
    End Type
    Public Type GMem
        GIF As GGIF
        kind As Integer
        Hwnd As Long
        ImgHwnd As Long
        Imgs(3) As Long
        name As String
        Folder As String
        w As Long
        h As Long
        copyed As Boolean
        CrashIndex As Long
    End Type
    Public Type AssetsTree
        files() As GMem
        path As String
        arg1 As Variant
        arg2 As Variant
    End Type
    Public Enum ImgDirection
        DirNormal = 0
        DirHorizontal = 1
        DirVertical = 2
        DirHorizontalVertical = 3
    End Enum
    Public Type GraphicsBound
        x As Long
        y As Long
        Width As Long
        Height As Long
        WSc As Single
        HSc As Single
        CrashIndex As Long
        Shape As Long
        Strings As String
    End Type
    Public Type ColorCollection
        IsAlpha() As Boolean
    End Type
    Public ColorLists() As ColorCollection
    Public ECore As GMan, EF As GFont, EAni As Object, ESave As GSaving, EMusic As GMusicList
    Public GHwnd As Long, GDC As Long, GW As Long, GH As Long
    Public Mouse As MState, DrawF As GraphicsBound
    Public FPS As Long, FPSt As Long, tFPS As Long, FPSct As Long, FPSctt As Long
    Public SysPage As GSysPage
    Public PreLoadCount As Long, LoadedCount As Long, ReLoadCount As Long
    Public FPSWarn As Long
    Public EmeraldInstalled As Boolean
    Public BassInstalled As Boolean
    Public Const Version As Long = 19070404      'hhhhhhhhhhhhhhhhffff
    Public TextHandle As Long, WaitChr As String
    
    Public AssetsTrees() As AssetsTree
    Dim LastKeyUpRet As Boolean
    Dim Wndproc As Long
'================================================================================
    '��ȡINI
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As Long, ByVal lpKeyName As Long, ByVal lpDefault As Long, ByVal lpReturnedString As Long, ByVal nSize As Long, ByVal lpFileName As Long) As Long
'================================================================================
'   ����ʱ
'   ��ȡINI�ļ�
'   <SectionName:��������,KeyName:������,IniFileName:INI�ļ�·��>
    Private Function ReadINI(ByVal SectionName As String, ByVal KeyName As String, ByVal IniFileName As String) As String
        Dim strBuf As String
        strBuf = String(128, 0)
        GetPrivateProfileString StrPtr(SectionName), StrPtr(KeyName), StrPtr(""), StrPtr(strBuf), 128, StrPtr(IniFileName)
        strBuf = Left(strBuf, InStr(strBuf, Chr(0)))
        ReadINI = strBuf
    End Function
    Public Sub OutPutDebug(Str As String)
        Open App.path & "\debug.txt" For Append As #1
        Print #1, Now & "    " & Str
        Close #1
    End Sub
'================================================================================
'   Init
    Public Sub SaveSettings(data As GSaving)
        data.PutData "DebugMode", DebugMode
        data.PutData "DisableLOGO", DisableLOGO
        data.PutData "HideLOGO", HideLOGO
        data.PutData "UpdateCheckInterval", UpdateCheckInterval
        data.PutData "UpdateTimeOut", UpdateTimeOut
    End Sub
    Public Sub GetSettings(Optional SkipDebug As Boolean = False)
        If App.LogMode <> 0 And SkipDebug = False Then Exit Sub
    
        Dim data As New GSaving
        data.Create "Emerald.Core"
        data.AutoSave = True
        
        If data.GetData("DebugMode") = "" Then
            UpdateCheckInterval = 1
            UpdateTimeOut = 2000
            Call SaveSettings(data)
        End If
        
        DebugSwitch.DebugMode = Val(data.GetData("DebugMode"))
        DebugSwitch.DisableLOGO = Val(data.GetData("DisableLOGO"))
        DebugSwitch.HideLOGO = Val(data.GetData("HideLOGO"))
        DebugSwitch.UpdateCheckInterval = Val(data.GetData("UpdateCheckInterval"))
        DebugSwitch.UpdateTimeOut = Val(data.GetData("UpdateTimeOut"))
        
        Set data = Nothing
    End Sub
    Public Sub StartEmerald(Hwnd As Long, w As Long, h As Long)
        ReDim ColorLists(0)
            
        Dim strComputer, objWMIService, colItems, objItem, strOSversion As String
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
        
        For Each objItem In colItems
            strOSversion = objItem.Version
        Next
    
        Select Case Val(Split(strOSversion, ".")(0))
        Case Is <= "5"
            MsgBox "�ǳ���Ǹ��Emerald����֧��������Windows 7���°汾�Ĳ���ϵͳ��" & vbCrLf & vbCrLf & "������з����ṩ֧�֣�����ϵQQ 1361778219��", 48, "Emerald�������ݵĲ���ϵͳ"
            End
        End Select
    
        Call GetSettings
    
        If DebugMode Then
            If App.LogMode <> 0 Then MsgBox "��������ʱδ�ر�Debugģʽ��": End
        End If
        
        ReDim AssetsTrees(0)
        
        InitGDIPlus
        
        GHwnd = Hwnd: GW = w: GH = h
        Dim DPI As Long
        DPI = 1440 / Screen.TwipsPerPixelX
        If (GetWindowLongA(Hwnd, GWL_STYLE) And WS_CAPTION) = WS_CAPTION Then
            SetWindowPos Hwnd, 0, 0, 0, w + 3 * Int(DPI / 96), h + 26 * Int(DPI / 96), SWP_NOMOVE Or SWP_NOZORDER
        Else
            SetWindowPos Hwnd, 0, 0, 0, w - 2 * Int(DPI / 96), h - 2 * Int(DPI / 96), SWP_NOMOVE Or SWP_NOZORDER
        End If
        
        GDC = GetDC(Hwnd)
        If App.LogMode <> 0 Then Wndproc = SetWindowLongA(Hwnd, GWL_WNDPROC, AddressOf Process)
        
        Set EAni = New GAnimation
        Set SysPage = New GSysPage
        
        If Val(GetWinNTVersion) > 6.1 Then               '�����ǰϵͳ�汾����win7
            SetProcessDpiAwareness 2&                    '����APIʹ�������ڸ�DPI����²�ģ��
        End If
        
        If DebugMode Then
            Debuginfo.Show
            Debuginfo.Hide
            DebugWindow.Show
        End If
        
        If App.LogMode = 0 Then Call CheckUpdate
        
        If ReLoadCount > LoadedCount Then Suggest "�ظ����ص���Դ����̫������������ÿ��ҳ�����Դ����һ���ļ��з�����"
        
    End Sub
    Public Sub Suggest(text As String)
        Debug.Print Now, "Emerald�ν��飺" & text
    End Sub
    Public Sub EndEmerald()
        If DebugMode Then
            Unload Debuginfo
            Unload DebugWindow
        End If
        
        If App.LogMode <> 0 Then SetWindowLongA GHwnd, GWL_WNDPROC, Wndproc
        If Not (ECore Is Nothing) Then ECore.Dispose
        If Not (EF Is Nothing) Then EF.Dispose
        TerminateGDIPlus
        If BassInstalled Then BASS_Free
    End Sub
    Public Sub MakeFont(ByVal name As String)
        Set EF = New GFont
        EF.MakeFont name
    End Sub
'========================================================
'   RunTime
    Public Function ToTime(time) As String
        ToTime = Int(time / 60) & ":" & format(time Mod 60, "00")
    End Function
    Public Function Process(ByVal Hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        On Error GoTo sth

        If uMsg = WM_MOUSEWHEEL Then
            Dim Direction As Integer, Strong As Single
            Direction = IIf(wParam < 0, -1, 1): Strong = Abs(wParam / 7864320)
            ECore.Wheel Direction, Strong
        End If
        
last:
        Process = CallWindowProcA(Wndproc, Hwnd, uMsg, wParam, lParam)
sth:

    End Function
'   ȡ�õ�ǰϵͳ��WinNT�汾
    Public Function GetWinNTVersion() As String
        Dim strComputer, objWMIService, colItems, objItem, strOSversion As String
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
        
        For Each objItem In colItems
            strOSversion = objItem.Version
        Next
        
        GetWinNTVersion = Left(strOSversion, 3)
    End Function
    Public Sub BlurTo(DC As Long, srcDC As Long, buffWin As Form, Optional radius As Long = 60)
        Dim i As Long, g As Long, e As Long, b As BlurParams, w As Long, h As Long
        'ճ�������崰��
        buffWin.AutoRedraw = True
        BitBlt buffWin.hdc, 0, 0, GW, GH, srcDC, 0, 0, vbSrcCopy: buffWin.Refresh
        
        '����Bitmap
        GdipCreateBitmapFromHBITMAP buffWin.Image.handle, buffWin.Image.hpal, i
        
        'ģ������
        GdipCreateEffect2 GdipEffectType.Blur, e: b.radius = radius: GdipSetEffectParameters e, b, LenB(b)
        GdipGetImageWidth i, w: GdipGetImageHeight i, h
        GdipBitmapApplyEffect i, e, NewRectL(0, 0, w, h), 0, 0, 0
        
        '��~
        GdipCreateFromHDC DC, g
        GdipDrawImage g, i, 0, 0
        GdipDisposeImage i: GdipDeleteGraphics g: GdipDeleteEffect e '��������
        buffWin.AutoRedraw = False
    End Sub
    Public Sub BlurImg(img As Long, radius As Long)
        Dim b As BlurParams, e As Long, w As Long, h As Long
        
        'ģ������

        GdipCreateEffect2 GdipEffectType.Blur, e: b.radius = radius: GdipSetEffectParameters e, b, LenB(b)
        GdipGetImageWidth img, w: GdipGetImageHeight img, h
        GdipBitmapApplyEffect img, e, NewRectL(0, 0, w, h), 0, 0, 0
        
        '��~
        GdipDeleteEffect e '��������
    End Sub
    Public Function CreateCDC(w As Long, h As Long) As Long
        Dim bm As BITMAPINFOHEADER, DC As Long, DIB As Long
    
        With bm
            .biBitCount = 32
            .biHeight = h
            .biWidth = w
            .biPlanes = 1
            .biSizeImage = (.biWidth * .biBitCount + 31) / 32 * 4 * .biHeight
            .biSize = Len(bm)
        End With
        
        DC = CreateCompatibleDC(GDC)
        DIB = CreateDIBSection(DC, bm, DIB_RGB_COLORS, ByVal 0, 0, 0)
        DeleteObject SelectObject(DC, DIB)
        
        CreateCDC = DC
    End Function
    Public Sub PaintDC(DC As Long, destDC As Long, Optional x As Long = 0, Optional y As Long = 0, Optional cx As Long = 0, Optional cy As Long = 0, Optional cw, Optional ch, Optional alpha)
        Dim b As BLENDFUNCTION, index As Integer, bl As Long
        
        If Not IsMissing(alpha) Then
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
        
        If IsMissing(cw) Then cw = GW - cx
        If IsMissing(ch) Then ch = GH - cy
        
        If IsMissing(alpha) Then
            BitBlt destDC, x, y, cw, ch, DC, cx, cy, vbSrcCopy
        Else
            AlphaBlend destDC, x, y, cw, ch, DC, cx, cy, cw, ch, bl
        End If
    End Sub
    Function Cubic(t As Single, arg0 As Single, arg1 As Single, arg2 As Single, arg3 As Single) As Single
        'Formula:B(t)=P_0(1-t)^3+3P_1t(1-t)^2+3P_2t^2(1-t)+P_3t^3
        'Attention:all the args must in this area (0~1)
        Cubic = (arg0 * ((1 - t) ^ 3)) + (3 * arg1 * t * ((1 - t) ^ 2)) + (3 * arg2 * (t ^ 2) * (1 - t)) + (arg3 * (t ^ 3))
    End Function
'========================================================
'   Mouse
    Public Sub UpdateMouse(x As Single, y As Single, state As Long, button As Integer)
        With Mouse
            .x = x
            .y = y
            .state = state
            .button = button
        End With
    End Sub
    Public Function CheckMouse(x As Long, y As Long, w As Long, h As Long) As MButtonState
        'Return Value:0=none,1=in,2=down,3=up
        If Mouse.x >= x And Mouse.y >= y And Mouse.x <= x + w And Mouse.y <= y + h Then
            CheckMouse = Mouse.state + 1
            If Mouse.state = 2 Then Mouse.state = 0
        End If
    End Function
    Public Function CheckMouse2() As MButtonState
        'Return Value:0=none,1=in,2=down,3=up
        If Mouse.x >= DrawF.x And Mouse.y >= DrawF.y And Mouse.x <= DrawF.x + DrawF.Width And Mouse.y <= DrawF.y + DrawF.Height Then
            CheckMouse2 = Mouse.state + 1
            If DrawF.CrashIndex <> 0 Then
                If ColorLists(DrawF.CrashIndex).IsAlpha((Mouse.x - DrawF.x) * DrawF.WSc, (Mouse.y - DrawF.y) * DrawF.HSc) = False Then CheckMouse2 = mMouseOut: Exit Function
            End If
            If Mouse.state = 2 Then Mouse.state = 0
        End If
    End Function
'========================================================
'   KeyBoard
    Public Function IsKeyPress(Code As Long) As Boolean
        IsKeyPress = (GetAsyncKeyState(Code) < 0)
    End Function
    Public Function IsKeyUp(Code As Long) As Boolean
        Dim t As Boolean
        t = LastKeyUpRet
        LastKeyUpRet = (GetAsyncKeyState(Code) < 0)
        If t = True And LastKeyUpRet = False Then IsKeyUp = True
    End Function
'========================================================
'   Screen Window
    Public Function StartScreenDialog(w As Long, h As Long, ch As Object) As Object
        Set StartScreenDialog = New EmeraldWindow
        StartScreenDialog.NewFocusWindow w, h, ch
        Dim f As Object
        For Each f In VB.Forms
            If TypeName(f) <> "EmeraldWindow" Then f.Enabled = False
        Next
    End Function
'========================================================
'   Update
    Public Sub CheckUpdate()
        On Error Resume Next
        If InternetGetConnectedState(0&, 0&) = 0 Then
            Debug.Print Now, "Emerald��δ�������磬������ȡ����"
            Exit Sub
        End If
        
        Dim data As New GSaving
        data.Create "Emerald.Core"
        data.AutoSave = True
        If Now - CDate(data.GetData("UpdateTime")) >= UpdateCheckInterval Or data.GetData("UpdateAble") = 1 Then
            data.PutData "UpdateTime", Now
            
            Dim xmlHttp As Object, Ret As String, Start As Long
            Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
            xmlHttp.Open "GET", "https://raw.githubusercontent.com/Red-Error404/Emerald/master/Version.txt", True
            xmlHttp.send
                         
            Start = GetTickCount
            Do While xmlHttp.ReadyState <> 4
                If GetTickCount - Start >= UpdateTimeOut Then
                    Debug.Print Now, "Emerald�������³�ʱ��"
                    Exit Sub
                End If
                Sleep 10: DoEvents
            Loop
            Ret = xmlHttp.responseText
            Set xmlHttp = Nothing
            Debug.Print Now, "Emerald�����汾��ϣ����°汾�� " & Val(Ret)
            
            If Val(Ret) > Version And App.LogMode = 0 Then
                data.PutData "UpdateAble", 1
                If MsgBox("����Emerald�����°汾����ϣ������ǰ��������", vbYesNo + 48, "Emerald") = vbNo Then Exit Sub
                
                ShellExecuteA 0, "open", "https://github.com/Red-Error404/Emerald/release", "", "", SW_SHOW
                data.PutData "UpdateAble", 0
            End If
        Else
            Debug.Print Now, "Emerald���ϴμ�����ʱ�� " & CDate(data.GetData("UpdateTime"))
        End If
        
        Set data = Nothing
    End Sub
'========================================================
'   AssetsTree
    Public Function AddAssetsTree(Tree As AssetsTree, arg1 As Variant, arg2 As Variant)
        ReDim Preserve AssetsTrees(UBound(AssetsTrees) + 1)
        AssetsTrees(UBound(AssetsTrees)) = Tree
        AssetsTrees(UBound(AssetsTrees)).arg1 = arg1
        AssetsTrees(UBound(AssetsTrees)).arg2 = arg2
    End Function
    Public Function FindAssetsTree(path As String, arg1 As Variant, arg2 As Variant) As Integer
        On Error Resume Next
        For i = 1 To UBound(AssetsTrees)
            If AssetsTrees(i).path = path And AssetsTrees(i).arg1 = arg1 And AssetsTrees(i).arg2 = arg2 Then
                If Err.Number <> 0 Then
                    Err.Clear
                Else
                    FindAssetsTree = i: Exit For
                End If
            End If
        Next
    End Function
    Public Function GetAssetsTree(path As String) As AssetsTree
        For i = 1 To UBound(AssetsTrees)
            If AssetsTrees(i).path = path Then GetAssetsTree = AssetsTrees(i): Exit For
        Next
    End Function
'========================================================
