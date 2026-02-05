Imports System.IO.MemoryMappedFiles
Imports System.Runtime.InteropServices

Public NotInheritable Class AstoniaProcess : Implements IDisposable

    Public ReadOnly proc As Process

    'Public thumbID As IntPtr

    Public Sub New(Optional process As Process = Nothing)
        proc = process
    End Sub

    Public Sub ResetCache()
        _rcc = Nothing
        _CO = Nothing
        _rcSource = Nothing
    End Sub

    Private _rcc? As Rectangle
    Public ReadOnly Property ClientRect() As Rectangle
        Get
            If _rcc IsNot Nothing AndAlso _rcc <> New Rectangle Then
                Return _rcc
            Else
                'dBug.print("called get ClientRect")
                Dim rcc As New Rectangle
                If proc Is Nothing Then Return New Rectangle
                GetClientRect(Me.MainWindowHandle, rcc)
                If rcc <> New Rectangle Then _rcc = rcc
                Return rcc
            End If
        End Get
    End Property

    Public hideRestart As Boolean = False

    Private TIattached As Boolean = False
    Private Shared AttachedThreads As New HashSet(Of Integer)()

    Public Shared Sub DetachThreadInput(aps As List(Of AstoniaProcess))
        dBug.Print(If(AttachedThreads.Count > 0, $"DetachThreadInput {AttachedThreads.Count}", ""))
        For Each tid In AttachedThreads
            AttachThreadInput(ScalaThreadId, tid, False)
        Next
        For Each ap In aps
            ap.TIattached = False
        Next
        AttachedThreads.Clear()
    End Sub

    Public Sub ThreadInput(attach As Boolean)
        If TIattached = attach Then Exit Sub

        If attach Then
            If AttachThreadInput(ScalaThreadId, Me.MainThreadId, True) Then
                dBug.Print($"ThreadInput Attach {Me.Name}")
                AttachedThreads.Add(Me.MainThreadId)
            End If
        Else
            dBug.Print($"ThreadInput Detach {AttachedThreads.Count}")
            For Each tid In AttachedThreads
                AttachThreadInput(ScalaThreadId, tid, False)
            Next
            AttachThreadInput(ScalaThreadId, Me.MainThreadId, False) 'loop fails when alt or scala is minimized so we explicilty detach main thread
            AttachedThreads.Clear()
        End If
        TIattached = attach
    End Sub

    Private _MainThreadId As Integer?
    Public ReadOnly Property MainThreadId() As Integer?
        Get
            If _MainThreadId Is Nothing Then _MainThreadId = GetWindowThreadProcessId(Me.MainWindowHandle, Nothing)
            Return _MainThreadId
        End Get
    End Property

    Public Shared BiggestWindowWidth As Integer
    Public Shared BiggestWindowHeight As Integer

    Public ReadOnly Property WindowRect() As Rectangle
        Get
            Dim rcw As New Rectangle
            If proc Is Nothing Then Return New Rectangle

            GetWindowRect(Me.MainWindowHandle, rcw)

            If rcw.Left - rcw.Width > BiggestWindowWidth Then
                BiggestWindowWidth = rcw.Left - rcw.Width
            End If

            If rcw.Top - rcw.Height > BiggestWindowHeight Then
                BiggestWindowHeight = rcw.Top - rcw.Height
            End If

            Return rcw
        End Get
    End Property
    Public Sub CloseOrKill()
        If proc Is Nothing Then Exit Sub
        Try
            If proc.HasExited() Then Exit Sub 'exception when proc is elevated
            If Me.isSDL Then
                proc.CloseMainWindow()
            Else 'isLegacy
                SendMessage(proc.MainWindowHandle, WM_KEYDOWN, Keys.F12, &H1UI << 30)
            End If
        Catch ex As System.ComponentModel.Win32Exception
            Try
                proc.Kill()
            Catch killEx As Exception
                dBug.Print($"CloseOrKill Kill failed: {killEx.Message}")
            End Try
        End Try
    End Sub
    Public Function IsMinimized() As Boolean
        If proc Is Nothing Then Return False
        Return IsIconic(Me.MainWindowHandle)
    End Function

    Public Function Restore() As Integer
        Dim ret As Integer = ShowWindow(Me.MainWindowHandle, SW_RESTORE)
        Me.ResetCache()
        Return ret
    End Function
    Public Function Hide() As Integer
        Return ShowWindow(Me.MainWindowHandle, SW_MINIMIZE)
    End Function

    Private _restoreLoc As Point? = Nothing
    Private _wasTopmost As Boolean = False
#Disable Warning IDE0140 ' Object creation can be simplified 'needs to be declared and assigned for parrallel.foreach
    Private Shared ReadOnly _restoreDic As Concurrent.ConcurrentDictionary(Of Integer, AstoniaProcess) = New Concurrent.ConcurrentDictionary(Of Integer, AstoniaProcess)
#Enable Warning IDE0140 ' Object creation can be simplified
    Public Shared Sub RestorePos(Optional keep As Boolean = False)
        Dim behind As IntPtr
        If IPC.SidebarSender IsNot Nothing AndAlso Not IPC.SidebarSender.HasExitedSafe Then
            behind = IPC.SidebarSenderhWnd
        Else
            IPC.SidebarSender = Nothing
            behind = ScalaHandle
        End If
        Parallel.ForEach(_restoreDic.Values,
                         Sub(ap As AstoniaProcess)
                             Try
                                 If ap._restoreLoc IsNot Nothing AndAlso Not ap.HasExited() Then
                                     Dim beh = If(ap._wasTopmost, SWP_HWND.TOPMOST, behind)
                                     SetWindowPos(ap.MainWindowHandle, beh,
                                                  ap._restoreLoc?.X, ap._restoreLoc?.Y, -1, -1,
                                                  SetWindowPosFlags.IgnoreResize Or SetWindowPosFlags.DoNotActivate Or SetWindowPosFlags.DoNotChangeOwnerZOrder)
                                     If Not keep Then ap._restoreLoc = Nothing
                                 End If
                             Catch ex As Exception
                                 dBug.Print($"RestorePos: {ex.Message}")
                             End Try
                         End Sub)
        If Not keep Then _restoreDic.Clear()
    End Sub
    Public Sub RestoreSinglePos(Optional behind As Integer = 0)
        If Me._restoreLoc IsNot Nothing AndAlso Not Me.HasExited() Then
            dBug.Print($"restoresingle {behind}")
            If behind = 0 Then
                SetWindowPos(Me.MainWindowHandle, If(Me._wasTopmost, SWP_HWND.TOPMOST, ScalaHandle),
                             Me._restoreLoc?.X, Me._restoreLoc?.Y, -1, -1,
                             SetWindowPosFlags.IgnoreResize Or SetWindowPosFlags.DoNotActivate Or SetWindowPosFlags.DoNotChangeOwnerZOrder)
            Else
                _restoreDic.TryRemove(Me.Id, Nothing)
                SetWindowPos(Me.MainWindowHandle, behind,
                             Me._restoreLoc?.X, Me._restoreLoc?.Y, -1, -1,
                             SetWindowPosFlags.IgnoreResize Or SetWindowPosFlags.DoNotActivate Or SetWindowPosFlags.DoNotChangeOwnerZOrder)
            End If
        End If
    End Sub

    Public Sub SavePos(Optional pt As Point? = Nothing, Optional overwrite As Boolean = True)
        If pt Is Nothing Then
            Dim rc As Rectangle
            GetWindowRect(Me.MainWindowHandle, rc)
            pt = rc.Location
        End If
        If overwrite OrElse Not _restoreDic.ContainsKey(proc.Id) Then Me._restoreLoc = pt
        If Not _restoreDic.ContainsKey(proc.Id) Then _restoreDic.TryAdd(proc.Id, Me)
        Me._wasTopmost = Me.TopMost()
    End Sub

    Private _CO? As Point
    Public ReadOnly Property ClientOffset() As Point
        Get
            If _CO IsNot Nothing AndAlso _CO <> New Point Then
                Return _CO
            Else
                If proc Is Nothing Then
                    Return New Point
                End If
                Try
                    If Me.HasExited Then
                        Return New Point
                    End If
                    'Dim ptt As New Point
                    'ClientToScreen(Me.MainWindowHandle, ptt)
                    Dim rcW As RECT
                    GetWindowRect(Me.MainWindowHandle, rcW)
                    Dim rcC As RECT
                    GetClientRect(Me.MainWindowHandle, rcC)
                    'Dim extrapixel As Integer = 0
                    'Select Case Me.WindowsScaling
                    '    Case 125
                    '        extrapixel = 1
                    '    Case 175
                    '        extrapixel = 1
                    'End Select
                    '_CO = New Point(ptt.X - rcW.left, ptt.Y - rcW.top + extrapixel)

                    Dim border As Integer = (rcW.right - rcW.left - rcC.right) \ 2

                    _CO = New Point(border, rcW.bottom - rcW.top - rcC.bottom - border)

                    dBug.Print($"co {_CO}")
                    Return _CO
                Catch ex As Exception
                    dBug.Print(Message:=$"ex on CO {ex.Message}")
                    Return New Point
                End Try
            End If
        End Get
    End Property

    Private _rcSource? As Rectangle
    Public Function rcSource(TargetSZ As Size, mode As Integer) As Rectangle

        'Dim mode = My.Settings.ScalingMode
        'dBug.print($"rcSource target {TargetSZ}")
        'If mode = 0 Then
        '    Dim compsz As Size = TargetSZ
        '    If (compsz.Width / ClientRect.Width >= 2) AndAlso
        '       (compsz.Height / ClientRect.Height >= 2) Then
        '        mode = 2
        '    Else
        '        mode = 1
        '    End If
        'End If
        If mode = 1 Then 'blurred
            Return ClientRect 'todo handle non 100% scaling
        ElseIf mode = 2 Then 'pixel note: does not support non 100% windows scaling
            If _rcSource IsNot Nothing AndAlso _rcSource <> New Rectangle Then
                Return _rcSource
            Else
                Dim factor As Single = Me.WindowsScaling / 100
                '125%  +2 +5
                Dim nudge = 0
                If factor > 1 Then
                    nudge = 1
                End If
                _rcSource = New Rectangle(Math.Floor(ClientOffset.X * factor) + nudge,
                                          Math.Floor(ClientOffset.Y * factor),
                                          Math.Floor((ClientRect.Width + ClientOffset.X) * factor - nudge),
                                          Math.Floor((ClientRect.Height + ClientOffset.Y) * factor) - nudge)
                Return _rcSource

                _rcSource = New Rectangle(ClientOffset.X, ClientOffset.Y, ClientRect.Width + ClientOffset.X, ClientRect.Height + ClientOffset.Y)
                Return _rcSource
            End If
        End If

    End Function

    Public Function WindowsScaling() As Integer
        If Me.MainWindowHandle = IntPtr.Zero Then Return 0

        Dim hMon As IntPtr = MonitorFromWindow(Me.MainWindowHandle, MONITORFLAGS.DEFAULTTONEAREST)
        Return GetMonitorScaling(hMon)
    End Function

    Public Function CenterBehind(centerOn As Control, Optional extraSWPFlags As UInteger = 0, Optional force As Boolean = False, Optional fixThumb As Boolean = False) As Boolean
        If Not force AndAlso centerOn.RectangleToScreen(centerOn.Bounds).Contains(Form.MousePosition) Then Return False
        Dim pt As Point = centerOn.PointToScreen(New Point(centerOn.Width / 2, centerOn.Height / 2))
        Return CenterWindowPos(centerOn.FindForm.Handle, pt.X, pt.Y, extraSWPFlags, fixThumb)
    End Function
    Public Function CenterWindowPos(hWndInsertAfter As IntPtr, x As Integer, y As Integer, Optional extraSWPFlags As UInteger = 0, Optional fixThumb As Boolean = False) As Boolean
        Try
            y = y - ClientRect.Height / 2 - ClientOffset.Y ' - My.Settings.offset.Y
            If fixThumb Then
                Dim rc As RECT
                GetWindowRect(ScalaHandle, rc)
                If y <= rc.top Then y = rc.top + 1
            End If
            Return SetWindowPos(Me.MainWindowHandle, hWndInsertAfter,
                                x - ClientRect.Width / 2 - ClientOffset.X,' - My.Settings.offset.X,
                                y,
                                -1, -1,
                                SetWindowPosFlags.IgnoreResize Or extraSWPFlags)
        Catch
            dBug.Print("CenterWindowPos exception")
            Return False
        End Try
    End Function


    Public Shared Narrowing Operator CType(ByVal d As AstoniaProcess) As Process
        Return d?.proc
    End Operator
    Public Shared Widening Operator CType(ByVal b As Process) As AstoniaProcess
        Return GetFromCache(b)
    End Operator

    Public ReadOnly Property Id As Integer
        Get
            Return If(proc?.Id, 0)
        End Get
    End Property

    Public Function Activate() As Boolean
        If proc Is Nothing Then Return True

        AllowSetForegroundWindow(ASFW_ANY)
        Return SetForegroundWindow(Me.MainWindowHandle)

    End Function

    Private _mwhCache As IntPtr = IntPtr.Zero
    Public ReadOnly Property MainWindowHandle() As IntPtr
        Get
            If _mwhCache <> IntPtr.Zero Then Return _mwhCache
            Try
                _mwhCache = If(proc?.MainWindowHandle, IntPtr.Zero)
                Return _mwhCache
            Catch
                dBug.Print("MainWindowHandle exeption")
                Return IntPtr.Zero
            End Try
        End Get
    End Property

    'Private Shared ReadOnly nameCache As Concurrent.ConcurrentDictionary(Of Integer, String) = New Concurrent.ConcurrentDictionary(Of Integer, String)
    'Public ReadOnly Property Name As String
    '    Get
    '        If _proc Is Nothing Then Return "Someone"
    '        Try
    '            _proc?.Refresh()
    '            If _proc?.MainWindowTitle = "" Then
    '                Return nameCache.GetValueOrDefault(_proc.Id, "Someone")
    '            End If
    '            Dim nam = Strings.Left(_proc?.MainWindowTitle, _proc?.MainWindowTitle.IndexOf(" - "))
    '            Return nameCache.AddOrUpdate(_proc.Id, nam, Function() nam)
    '        Catch
    '            Return "Someone"
    '        End Try
    '    End Get
    'End Property

    'Friend ReadOnly EscDownAndUpInput() As INPUT = {
    '               New INPUT With {.type = InputType.INPUT_KEYBOARD,
    '                    .u = New InputUnion With {.ki = New KEYBDINPUT With {.dwFlags = KeyEventF.KeyDown, .wScan = 1, .wVk = Keys.Escape}}
    '               },
    '               New INPUT With {.type = InputType.INPUT_KEYBOARD,
    '                    .u = New InputUnion With {.ki = New KEYBDINPUT With {.dwFlags = KeyEventF.KeyUp, .wScan = 1, .wVk = Keys.Escape}}
    '               }
    '         }
    'Friend ReadOnly CtrlUpInput() As INPUT = {
    '    New INPUT With {.type = InputType.INPUT_KEYBOARD,
    '                      .u = New InputUnion With {.ki = New KEYBDINPUT With {.dwFlags = KeyEventF.KeyUp, .wVk = Keys.ControlKey}}
    '               }
    '    }
    'Friend ReadOnly CtrlDownInput() As INPUT = {
    '    New INPUT With {.type = InputType.INPUT_KEYBOARD,
    '                      .u = New InputUnion With {.ki = New KEYBDINPUT With {.dwFlags = KeyEventF.KeyDown, .wVk = Keys.ControlKey}}
    '               }
    '    }


    Private Shared ReadOnly memCache As New System.Runtime.Caching.MemoryCache("nameCache")
    Private Shared ReadOnly cacheItemPolicy As New System.Runtime.Caching.CacheItemPolicy With {
                    .SlidingExpiration = TimeSpan.FromMinutes(1)} ' Cache for 1 minute with sliding expiration
    Public hasLoggedIn As Boolean = False
    Public loggedInAs As String = String.Empty

    ''' <summary>
    ''' Tracks the last time meaningful input (mouse click) was detected for this client.
    ''' Used to warn users when a client is about to timeout due to inactivity.
    ''' </summary>
    Public LastInputTime As DateTime = DateTime.Now

    ''' <summary>
    ''' Records that input was detected for this client, resetting the inactivity timer.
    ''' </summary>
    Public Sub RecordInput()
        LastInputTime = DateTime.Now
    End Sub
    Dim gti As New GUITHREADINFO() With {.cbSize = CUInt(Marshal.SizeOf(GetType(GUITHREADINFO)))}
    Dim DoNotReplaceHwnd As IntPtr = IntPtr.Zero
    Public ReadOnly Property Name As String
        Get
            If proc Is Nothing Then Return "Someone"
            Try
                proc.Refresh()
                If proc.MainWindowTitle = "" Then
                    Dim nm As String = TryCast(memCache.Get(proc.Id), String)
                    If Not String.IsNullOrEmpty(nm) Then
                        Return nm
#If 0 Then
                        If Me.IsActive() AndAlso Me.isSDL Then

                            'check if sysmenu is opened from taskbar/tbthumb

                            Dim hwnd = GetAncestor(WindowFromPoint(Control.MousePosition), GA_ROOT)
                            Dim clss = GetWindowClass(hwnd)
                            If {"Shell_TrayWnd", "Shell_SecondaryTrayWnd", "XamlExplorerHostIslandWindow", "TaskListThumbnailWnd"}.Contains(clss) Then 'list is possibly incomplete  (win 7?? 8??)
                                DoNotReplaceHwnd = GetWindow(hwnd, GW_ENABLEDPOPUP)
                            End If

                            If GetWindow(hwnd, GW_ENABLEDPOPUP) = DoNotReplaceHwnd Then  'don't close when flag set
                                Return nm
                            Else
                                DoNotReplaceHwnd = IntPtr.Zero 'set flag on menu close to scalahandle to denote invalid state (can't use intptr.zero)
                            End If

                            If FrmMain.Bounds.Contains(Control.MousePosition) Then 'SDL client sysmenu open. close it and open our own or correct menu for whatever button is hovered.

                                'Check if alt is selected andalso check if alt is on activeoverview
                                If Me IsNot FrmMain.AltPP AndAlso
                               Not (My.Settings.gameOnOverview AndAlso FrmMain.pnlOverview.Controls.OfType(Of AButton).Any(Function(ab As AButton)
                                                                                                                               Dim abScreenrect = ab.RectangleToScreen(ab.ClientRectangle)
                                                                                                                               Return (ab.AP Is Me) AndAlso abScreenrect.Contains(Control.MousePosition)
                                                                                                                           End Function)) Then
                                    Return nm
                                End If

                                'double check if clients sysmenu is open
                                If GetGUIThreadInfo(Me.MainThreadId, gti) AndAlso (gti.flags = 20) Then
                                    dBug.Print($"gti.flags {gti.flags}")
                                    PostMessage(Me.MainWindowHandle, WM_CANCELMODE, 0, 0) 'close the SysMenu
                                End If

                                'determine cotrol beneath cursor
                                Dim pt As Point = FrmMain.PointToClient(FrmMain.MousePosition)
                                'Dim pt = Control.MousePosition
                                Dim ctl As Control = FrmMain.GetChildAtPoint(pt)
                                dBug.Print($"Rmb: {ctl?.Name} {pt}")

                                If ctl Is FrmMain.pnlSys Then
                                    'pt = FrmMain.pnlSys.PointToClient(FrmMain.MousePosition)
                                    ctl = FrmMain.pnlSys.GetChildAtPoint(pt)
                                    pt = ctl?.PointToClient(Control.MousePosition)
                                End If

                                'open altmenu or QL when on overview
                                If ctl Is FrmMain.pnlOverview Then
                                    ctl = FrmMain.pnlOverview.GetChildAtPoint(pt)
                                    If ctl Is FrmMain.pnlMessage Then
                                        ctl = FrmMain.pnlOverview.Controls.OfType(Of AButton).FirstOrDefault(Function(c) c.Visible)
                                        dBug.Print($"pnlmessage skip?: {ctl?.Name} {pt}")
                                    End If
                                    pt = ctl.PointToClient(Control.MousePosition)
                                    dBug.Print($"alt sysmenu override: {ctl?.Name} {pt}")

                                    If My.Computer.Keyboard.CtrlKeyDown Then
                                        FrmMain.cmsQuickLaunch.Show(ctl, pt)
                                    Else
                                        FrmMain.cmsAlt.Show(ctl, pt)
                                    End If
                                    Return nm
                                End If

                                'open own sysmenu or open quicklaunch/settings when over specific button
                                Select Case ctl?.Name
                                    Case FrmMain.btnStart.Name, FrmMain.cboAlt.Name
                                        FrmMain.cmsQuickLaunch.Show(ctl, pt)
                                    Case FrmMain.cmbResolution.Name
                                        FrmMain.CmbResolution_MouseUp(FrmMain.cmbResolution, New MouseEventArgs(MouseButtons.Right, 1, pt.X, pt.Y, 0))
                                    Case FrmMain.pnlTitleBar.Name, FrmMain.lblTitle.Name
                                        FrmMain.ShowSysMenu(FrmMain, New MouseEventArgs(MouseButtons.Right, 1, pt.X, pt.Y, 0))
                                End Select

                            End If
                            Return nm
                        End If
#End If
                    End If
                    Return "Someone"
                End If
                Dim nam As String = Strings.Left(proc.MainWindowTitle, proc.MainWindowTitle.IndexOf(" - "))
                memCache.Set(proc.Id, nam, cacheItemPolicy)
                If nam <> "Someone" AndAlso Not String.IsNullOrEmpty(nam) Then
                    loggedIns.TryAdd(Me.Id, Me)
                    loggedInAs = nam
                    If Not hasLoggedIn AndAlso FrmMain.cboAlt.SelectedItem?.Id = Me.Id Then
                        FrmMain.PopDropDown(FrmMain.cboAlt)
                    End If
                    hasLoggedIn = True
                End If
                Return nam
            Catch ex As InvalidOperationException 'process has exited
                Return "Someone"
            Catch ex As Exception
                dBug.Print($"Name exception {ex.Message}")
                Return "Someone"
            End Try
        End Get
    End Property



    Public ReadOnly Property DisplayName As String
        Get
            Dim dpname As String = Me.Name
            If dpname <> "Someone" Then
                Return Me.Name
            Else
                If Not String.IsNullOrEmpty(Me.loggedInAs) Then
                    Return Me.loggedInAs & $" (Someone)"
                Else
                    If Me.Id <> 0 Then
                        If Me.UserName <> "Someone" Then
                            Return Me.UserName & $" (Someone)"
                        End If
                    End If
                End If
            End If
            Return dpname
        End Get
    End Property

    Private _commandLine As String = String.Empty
    Public ReadOnly Property CommandLine As String
        Get
            If Not String.IsNullOrEmpty(_commandLine) Then Return _commandLine

            Try
                Task.Run(Sub()
                             Dim QS As New Management.ManagementObjectSearcher($"Select * from Win32_Process WHERE ProcessID={Me.Id}")
                             Dim objCol = QS.Get()
                             _commandLine = objCol.Cast(Of Management.ManagementObject)().FirstOrDefault()?("commandline")?.ToString()
                             If _commandLine Is Nothing Then _commandLine = String.Empty
                         End Sub).Wait()
            Catch ex As Exception
                _commandLine = String.Empty
            End Try

            Return _commandLine
        End Get
    End Property

    Private _OptionDict As Dictionary(Of String, String)
    Public Function getCMDoption(opt As String) As String
        If _OptionDict Is Nothing Then
            _OptionDict = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim cmdLine = Me.CommandLine
            If Not String.IsNullOrEmpty(cmdLine) Then
                ' Split on dashes, skip the first part (exe path)
                Dim parts = cmdLine.Split("-"c).Skip(1)
                For Each part In parts
                    Dim trimmed = part.Trim()
                    If trimmed.Length > 0 Then
                        Dim key = trimmed.Substring(0, 1).ToLower()
                        Dim val = trimmed.Substring(1).Trim()
                        ' Only add if not already present (first occurrence wins)
                        If Not _OptionDict.ContainsKey(key) Then
                            _OptionDict.Add(key, val)
                        End If
                    End If
                Next
            End If
        End If

        Dim value As String = String.Empty
        _OptionDict.TryGetValue(opt.ToLower, value)
        Return value
    End Function


#If DEBUG Then
    Public Sub DebugListAllArgs()
        dBug.Print($"Current Alt: {Me.Name} ({Me.loggedInAs}) {Me.UserName}", 1)
        For Each kvp As KeyValuePair(Of String, String) In _OptionDict
            Dim val = kvp.Value
            If kvp.Key = "p" Then val = "**REDACTED**"
            dBug.Print($"{kvp.Key} : {val}", 1)
        Next
        dBug.Print("---", 1)
    End Sub
#End If

    Private _userName As String = String.Empty
    Public ReadOnly Property UserName As String
        Get
            If Not String.IsNullOrEmpty(_userName) Then Return _userName

            Dim userArg As String = getCMDoption("u") 'this can be empty on elevation mismatch

            If String.IsNullOrEmpty(userArg) Then
                _userName = String.Empty
                Dim nam = Me.Name
                If nam.Contains(" ") Then
                    nam = nam.Replace("Sir ", "").Replace("Lady ", "")
                    'don't know how to handle this w/o hardcoding all titles
                    If nam.Contains(" ") Then
                        nam = nam.Split(" ")(0)
                    End If
                End If
                Return nam
            End If
            _userName = userArg.Trim().FirstToUpper(True)
            Return _userName
        End Get
    End Property

    Private _serverName As String
    Public ReadOnly Property ServerName() As String
        Get
            If Not String.IsNullOrEmpty(_serverName) Then Return _serverName

            _serverName = getCMDoption("d")?.Trim() 'this can be empty on elevation mismatch

            Return _serverName
        End Get
    End Property


    Public Property TopMost() As Boolean
        Get

            Try
                Return (GetWindowLong(Me.MainWindowHandle, GWL_EXSTYLE) And WindowStylesEx.WS_EX_TOPMOST) = WindowStylesEx.WS_EX_TOPMOST
            Catch ex As Exception
                dBug.Print($"TopMost Get: {ex.Message}")
                Return False
            End Try
        End Get
        Set(value As Boolean)
            Try
                SetWindowPos(Me.MainWindowHandle, If(value, SWP_HWND.TOPMOST, SWP_HWND.NOTOPMOST),
                                                  -1, -1, -1, -1,
                                                  SetWindowPosFlags.IgnoreMove Or SetWindowPosFlags.IgnoreResize Or SetWindowPosFlags.DoNotActivate)
            Catch ex As Exception
                dBug.Print($"TopMost Set: {ex.Message}")
            End Try
        End Set
    End Property

    Public Function IsRunning() As Boolean
        'todo: replace with limited access check
        Try
            Return proc IsNot Nothing AndAlso Not proc.HasExited
        Catch e As Exception
            FrmMain.tmrActive.Enabled = False
            FrmMain.tmrOverview.Enabled = False
            FrmMain.tmrTick.Enabled = False
            FrmMain.SaveLocation()
            My.Settings.Save()
            FrmMain.RestartSelf(True)
            End 'program
        End Try
    End Function

    Public Function IsActive() As Boolean
        If proc Is Nothing Then Return False
        Dim hWnd As IntPtr = GetForegroundWindow()
        Dim ProcessID As UInteger = 0

        GetWindowThreadProcessId(hWnd, ProcessID)

        Return proc?.Id = ProcessID
    End Function

    Public Function IsBelow(hwnd As IntPtr) As Boolean
        Dim curr As IntPtr = Me.MainWindowHandle
        Do While True
            curr = GetWindow(curr, 3)
            If curr = IntPtr.Zero Then Return False
            If curr = hwnd Then Return True
        Loop
        Return False
    End Function

    Public Function IsAbove(hwnd As IntPtr) As Boolean
        Dim curr As IntPtr = Me.MainWindowHandle
        Do While True
            curr = GetWindow(curr, 2)
            If curr = IntPtr.Zero Then Return False
            If curr = hwnd Then Return True
        Loop
        Return False
    End Function

    Public Function MainWindowTitle() As String
        Try
            proc?.Refresh()
            Return proc?.MainWindowTitle
        Catch
            dBug.print("MainWindowTitle exception")
            Return ""
        End Try
    End Function

    Private Shared classCache As String = String.Empty
    Private Shared classCacheSet As New HashSet(Of String)
    Private Shared ReadOnly pipe As Char() = {"|"c}
    Private Shared ReadOnly sysMenClass As String() = {"#32768", "SysShadow"}
    ''' <summary>
    ''' Returns True if WindowClass is in My.settings.className 
    ''' </summary>
    ''' <returns></returns>
    ''' 
    Public Function IsAstoniaClass() As Boolean
        Dim classes = My.Settings.className
        If classCache <> classes Then
            classCacheSet.Clear()
            classCacheSet = New HashSet(Of String)(classes.Split(pipe, StringSplitOptions.RemoveEmptyEntries) _
                                                          .Select(Function(wc) Strings.Trim(wc)).Concat(sysMenClass))
            classCache = classes
        End If
        Return classCacheSet.Contains(Me.WindowClass)
    End Function

    Private _wc As String
    Public ReadOnly Property WindowClass() As String
        Get
            If Not String.IsNullOrEmpty(_wc) Then Return _wc
            If Me.MainWindowHandle <> IntPtr.Zero Then _wc = GetWindowClass(Me.MainWindowHandle)
            Return _wc
        End Get
    End Property

    'Public Overrides Function ToString() As String
    '    If _proc Is Nothing Then Return "Someone"
    '    Return Me.Name()
    'End Function

    Shared ReadOnly nameIconCache As New Concurrent.ConcurrentDictionary(Of Integer, Tuple(Of Icon, String)) 'PID, icon, name
    Shared ReadOnly pathIcnCache As New Concurrent.ConcurrentDictionary(Of String, Icon) 'path, icon
    Public Function GetIcon(Optional invalidateCache As Boolean = False) As Icon
        If invalidateCache Then
            For Each tup In nameIconCache.Values
                DestroyIcon(tup.Item1.Handle)
            Next
            nameIconCache.Clear()
            For Each icn In pathIcnCache.Values
                DestroyIcon(icn.Handle)
            Next
            pathIcnCache.Clear()
        End If
        If proc Is Nothing Then Return Nothing
        Try
            Dim ID As Integer = proc.Id

            Dim IcoNam As New Tuple(Of Icon, String)(Nothing, Nothing)

            If ID > 0 AndAlso nameIconCache.TryGetValue(ID, IcoNam) AndAlso IcoNam.Item2 = Me.UserName Then
                Return IcoNam.Item1
            Else
                Dim tup As Tuple(Of Icon, String) = New Tuple(Of Icon, String)(Nothing, String.Empty)
                If nameIconCache.TryRemove(ID, tup) Then
                    dBug.Print($"evicted {ID} {tup.Item2}") ' do NOT destroyicon(item1) here, it is duplicated in PathIcnCache
                End If
            End If

            Dim path As String = Me.FinalPath()

            If Not String.IsNullOrEmpty(path) Then
                Dim ico As Icon = Nothing
                If pathIcnCache.TryGetValue(path, ico) Then
                    nameIconCache.TryAdd(ID, New Tuple(Of Icon, String)(ico, Me.UserName))
                    Return ico
                Else
                    ico = Icon.ExtractAssociatedIcon(path)
                    If ico IsNot Nothing Then
                        nameIconCache.TryAdd(ID, New Tuple(Of Icon, String)(ico, Me.UserName))
                        pathIcnCache.TryAdd(path, ico)
                    End If
                    Return ico
                End If
            End If

        Catch ex As Exception
            dBug.print($"Error retrieving icon: {ex.Message}")
        End Try

        Return Nothing
    End Function

    Dim pathCache As String = Nothing
    Public Function Path() As String
        If String.IsNullOrEmpty(pathCache) Then pathCache = proc?.Path()
        Return pathCache
    End Function

    Private _FinalPath As String = Nothing
    Public ReadOnly Property FinalPath() As String
        Get
            If String.IsNullOrEmpty(_FinalPath) Then _FinalPath = proc?.FinalPath()
            Return _FinalPath
        End Get
    End Property

    Public Overrides Function Equals(obj As Object) As Boolean
        Dim proc2 As AstoniaProcess = TryCast(obj, AstoniaProcess)
        'dBug.print($"obj {proc2?._proc?.Id} eqals _proc {_proc?.Id}")
        Return proc2?.proc IsNot Nothing AndAlso Me.proc IsNot Nothing AndAlso proc2.proc.Id = Me.proc.Id AndAlso proc2.UserName = Me.UserName
    End Function
    'Public Shared Operator =(left As AstoniaProcess, right As AstoniaProcess) As Boolean
    '    dBug.print("AP ==")
    '    Return left.Equals(right)
    'End Operator
    'Public Shared Operator <>(left As AstoniaProcess, right As AstoniaProcess) As Boolean
    '    dBug.print("AP <>")
    '    Return Not left.Equals(right)
    'End Operator
    Private Shared exeCache As IEnumerable(Of String) = My.Settings.exe.Split(pipe, StringSplitOptions.RemoveEmptyEntries).Select(Function(s) s.Trim).ToList
    Private Shared exeSettingCache As String = My.Settings.exe

    Public Shared Function EnumSomeone(Optional usecache As Boolean = False) As IEnumerable(Of AstoniaProcess)
        If exeSettingCache <> My.Settings.exe Then
            exeCache = My.Settings.exe.Split(pipe, StringSplitOptions.RemoveEmptyEntries).Select(Function(s) s.Trim).ToList
            exeSettingCache = My.Settings.exe
        End If
        Return exeCache.SelectMany(Function(s) Process.GetProcessesByName(s).Select(Function(p)
                                                                                        If usecache Then
                                                                                            Dim ap As AstoniaProcess = GetFromCache(p)
                                                                                            Return ap
                                                                                        Else
                                                                                            Return New AstoniaProcess(p)
                                                                                        End If
                                                                                    End Function)) _
            .Where(Function(ap) ap.Name = "Someone" AndAlso ap.IsAstoniaClass())
    End Function

    Public Shared Function EnumAll() As IEnumerable(Of AstoniaProcess)
        If exeSettingCache <> My.Settings.exe Then
            exeCache = My.Settings.exe.Split(pipe, StringSplitOptions.RemoveEmptyEntries).Select(Function(s) s.Trim).ToList
            exeSettingCache = My.Settings.exe
        End If
        Return exeCache.SelectMany(Function(s) Process.GetProcessesByName(s).Select(Function(p) New AstoniaProcess(p))) _
            .Where(Function(ap) ap.IsAstoniaClass())
    End Function

    Public Shared Function ListProcesses(blacklist As IEnumerable(Of String), useCache As Boolean) As List(Of AstoniaProcess)
        'todo move updating cache to frmSettings
        If exeSettingCache <> My.Settings.exe Then
            exeCache = My.Settings.exe.Split(pipe, StringSplitOptions.RemoveEmptyEntries).Select(Function(s) s.Trim).ToList
            exeSettingCache = My.Settings.exe
        End If

        Return exeCache.SelectMany(Function(s) Process.GetProcessesByName(s).Select(Function(p) If(useCache, GetFromCache(p), New AstoniaProcess(p)))) _
                    .Where(Function(ap)
                               Dim nam = ap.UserName
                               Return Not String.IsNullOrEmpty(nam) AndAlso Not blacklist.Contains(ap.Name) AndAlso
                                     (Not My.Settings.Whitelist OrElse FrmMain.topSortList.Concat(FrmMain.botSortList).Contains(nam)) AndAlso
                                      ap.IsAstoniaClass()
                           End Function).ToList
    End Function
    Public Shared Function Enumerate(Optional useCache As Boolean = False) As IEnumerable(Of AstoniaProcess)
        Return AstoniaProcess.Enumerate({}, useCache)
    End Function
    Public Shared ProcCache As New List(Of AstoniaProcess)
    Public Shared CacheCounter As Integer = 0
    Public Shared loggedIns As New Concurrent.ConcurrentDictionary(Of Integer, AstoniaProcess)()
    'todo: make function to enumerate loggedins for autoclose
    ' move this whole functionality to a different process?
    '  set up ipc to control it's settings? 
    '  how to handle conflicting settings?
    '  drunk. GENTSE FEESTEN

    Private Shared Function GetFromCache(p As Process) As AstoniaProcess

        Return If(ProcCache.Find(Function(ap)
                                     If ap.HasExited Then Return False
                                     Return ap.Id = p.Id
                                 End Function), New AstoniaProcess(p))
    End Function

    Private Shared Function GetFromCache(pid As Integer) As AstoniaProcess
        Return If(ProcCache.Find(Function(ap)
                                     If ap.HasExited Then Return False
                                     Return ap.Id = pid
                                 End Function), New AstoniaProcess(Process.GetProcessById(pid)))
    End Function

    Public Shared Function FromHWnd(hWnd As IntPtr) As AstoniaProcess
        Dim pid As Integer = 0
        GetWindowThreadProcessId(hWnd, pid)
        If pid <> 0 Then
            Try
                Return GetFromCache(pid)
            Catch ex As Exception
                dBug.Print($"GetFromHandle GetFromCache failed for pid {pid}: {ex.Message}")
            End Try
        End If
        Return Nothing
    End Function

    Public Shared Function Enumerate(blacklist As IEnumerable(Of String), Optional useCache As Boolean = False) As IEnumerable(Of AstoniaProcess)
        ', Optional resetCacheFirst As Boolean = False
        'If resetCacheFirst Then
        '    _CacheCounter = 0
        '    _ProcCache.Clear()
        'End If
        If useCache Then
            If CacheCounter = 0 Then
                ProcCache = ListProcesses(blacklist, True)
            End If
            CacheCounter += 1
            If CacheCounter > 5 Then
                CacheCounter = 0
                ProcCache.RemoveAll(Function(ap) ap.HasExited)
            End If
            Return ProcCache
        End If
        Return ListProcesses(blacklist, False)
    End Function

    <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function PrintWindow(ByVal hwnd As IntPtr, ByVal hDC As IntPtr, ByVal nFlags As UInteger) As Boolean : End Function
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function GetClientRect(ByVal hWnd As IntPtr, ByRef lpRect As Rectangle) As Boolean : End Function
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function GetClientRect(ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Boolean : End Function

    Public Function GetClientBitmap() As Bitmap
        If proc Is Nothing Then Return Nothing
        Try
            Dim rcc As Rectangle
            If Not GetClientRect(Me.MainWindowHandle, rcc) Then Return Nothing 'GetClientRect fails if astonia is running fullscreen and is tabbed out

            If rcc.Width = 0 OrElse rcc.Height = 0 Then Return Nothing

            Dim fact As Double = Me.WindowsScaling / 100
            Dim bmp As New Bitmap(CInt(rcc.Width * fact), CInt(rcc.Height * fact))

            Using gBM As Graphics = Graphics.FromImage(bmp)
                Dim hdcBm As IntPtr
                Try
                    hdcBm = gBM.GetHdc
                Catch ex As Exception
                    dBug.print("GetHdc error")
                    Return Nothing
                End Try
                PrintWindow(Me.MainWindowHandle, hdcBm, 1)
                gBM.ReleaseHdc()
            End Using
            Return bmp
        Catch
            Return Nothing
        End Try
    End Function
    Private _isSDL? As Boolean = Nothing
    Public ReadOnly Property isSDL() As Boolean
        Get
            If _isSDL Is Nothing Then
                va.Read(0, shm)
                If shm.pID = Me.Id Then
                    _isSDL = True
                Else
                    _isSDL = False
                End If
                'dBug.print($"isSDL {Me.Name} {_isSDL} {shm.pID} {Me.Id}")
            End If
            Return _isSDL
        End Get
    End Property
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto, Pack:=0)>
    Structure MoacSharedMem
        Dim pID As UInt32
        Dim hp, shield, [end], mana As Byte
        Dim base As UInt64
        Dim key, isprite, offX, offY As Integer
        Dim flags, fsprite As Integer
        Dim swapped As Byte
    End Structure
    Private shm As MoacSharedMem
    Private _map As MemoryMappedFile = Nothing
    Private ReadOnly Property map As MemoryMappedFile
        Get
            If _map Is Nothing Then
                _map = MemoryMappedFile.CreateOrOpen($"MOAC{Me.Id}", Marshal.SizeOf(shm))
            End If
            Return _map
        End Get
    End Property
    Private _va As MemoryMappedViewAccessor = Nothing
    Private ReadOnly Property va As MemoryMappedViewAccessor
        Get
            If _va Is Nothing Then
                _va = map.CreateViewAccessor(0, Marshal.SizeOf(shm), MemoryMappedFileAccess.Read)
            End If
            Return _va
        End Get
    End Property
    Public Property DpiAware As Boolean? = Nothing
    Friend Property RegHighDpiAware As Boolean 'proc.MainModule.FileName has wrong path for junctioned/hardlinked/.... will GetFinalPathNameByHandle do the trick?
        Get
            If Me.DpiAware IsNot Nothing Then Return Me.DpiAware
            dBug.Print($"FinalPath ""{Me.FinalPath}""")
            Dim key As Microsoft.Win32.RegistryKey = Nothing
            Try
                key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(REGISTRY_COMPAT_LAYERS)
                Dim value = key?.GetValue(Me.FinalPath)
                If Me.isSDL Then
                    If value IsNot Nothing AndAlso value.contains("HIGHDPIAWARE") Then
                        Me.DpiAware = True
                        Return True
                    End If
                Else
                    If value IsNot Nothing AndAlso value.contains("HIGHDPIAWARE") Then
                        Me.DpiAware = True
                        Return True
                    End If
                End If
            Catch ex As Exception
                dBug.Print($"Exception RegDPI: {ex.Message}")
                Me.DpiAware = False
                Return False
            Finally
                key?.Close()
            End Try


            Me.DpiAware = False
            Return False
        End Get
        Set(value As Boolean)
            ' Open the key with write access. Create if necessary.
            Dim key As Microsoft.Win32.RegistryKey = Nothing
            Try
                key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(REGISTRY_COMPAT_LAYERS, writable:=True)
                If key Is Nothing Then
                    ' The key doesn't exist, so create it
                    key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(REGISTRY_COMPAT_LAYERS)
                End If

                ' Get the current value for this process path
                Dim currentValue As String = DirectCast(key.GetValue(Me.FinalPath, String.Empty), String).Trim()

                ' Modify the value based on the new setting
                If value Then
                    ' Add the HIGHDPIAWARE flag if it's not already present
                    If Not currentValue.Contains("HIGHDPIAWARE") Then
                        ' If currentValue is empty, set it to "~ HIGHDPIAWARE"
                        If String.IsNullOrEmpty(currentValue) Then
                            currentValue = DPI_AWARE_REMOVE
                        Else
                            ' If currentValue contains other flags, append "HIGHDPIAWARE"
                            currentValue &= DPI_AWARE_FLAG
                        End If
                    End If
                Else
                    ' Remove the HIGHDPIAWARE flag if present
                    If currentValue.Contains("HIGHDPIAWARE") Then
                        ' Remove the flag and clean up spaces
                        currentValue = System.Text.RegularExpressions.Regex.Replace(currentValue, "HIGHDPIAWARE\s*", "").Trim()
                        ' Remove "~" if it's the only character remaining
                        If currentValue = "~" Then currentValue = String.Empty
                    End If
                End If

                ' Update the registry with the new flags
                If String.IsNullOrEmpty(currentValue) Then
                    ' Delete the value if it is empty
                    key.DeleteValue(Me.FinalPath, throwOnMissingValue:=False)
                Else
                    key.SetValue(Me.FinalPath, currentValue)
                End If
            Catch ex As Exception
                ' Handle exceptions (e.g., permission issues) as needed
                dBug.print("SETREG An error occurred: " & ex.Message)
            Finally
                key?.Close()
            End Try

            ' Do not update the cached value. The process needs to be restarted for it to take effect.
            ' Me.DpiAware = value
        End Set
    End Property

    Public Function hasBorder() As Boolean
        Return GetWindowLong(Me.MainWindowHandle, GWL_STYLE) And WindowStyles.WS_BORDER
    End Function

    ReadOnly br As New SolidBrush(Color.FromArgb(&HFFFF0400))
    ReadOnly bo As New SolidBrush(Color.FromArgb(&HFFFF7D29))
    ReadOnly by As New SolidBrush(Color.FromArgb(186, 186, 30))
    ReadOnly bl As New SolidBrush(Color.FromArgb(217, 217, 30))
    ReadOnly bb As New SolidBrush(Color.FromArgb(&HFF297DFF))


    Dim paneOpen As Boolean = False
    Private Shared ReadOnly validColors As Integer() = {&HFFFF0000, &HFFFF0400, &HFFFF7B29, &HFFFF7D29, &HFF297BFF, &HFF297DFF, &HFF000000, &HFF000400, &HFFFFFFFF} 'red, orange, lightblue, black, white (troy,base)

    Public Function GetHealthbar(Optional width As Integer = 75, Optional height As Integer = 15) As Bitmap
        Dim bmp As New Bitmap(width, height)

        'struct sharedmem {
        '    unsigned int pid; 0
        '    Char hp, shield,end, mana; 4 5 6 7
        If Me.isSDL Then
            va.Read(0, shm)

            'Dim bars(4) As Byte
            'va.ReadArray(Of Byte)(4, bars, 0, 4)
            Using g As Graphics = Graphics.FromImage(bmp)

                g.Clear(Color.Black)
                If shm.mana = 255 Then
                    g.FillRectangle(br, New Rectangle(0, 0, shm.hp / 100 * width, height / 5 * 2))
                    g.FillRectangle(bo, New Rectangle(0, height / 5 * 2, shm.shield / 100 * width, height / 5 * 2))
                    g.FillRectangle(If(My.Settings.DarkMode, by, bl), New Rectangle(0, height / 5 * 4, shm.end / 100 * width, height / 5))
                Else
                    Dim itemheight As Integer = height / 15 * 5
                    If My.Settings.ShowEnd Then
                        itemheight = height / 15 * 4
                    End If
                    g.FillRectangle(br, New Rectangle(0, 0, shm.hp / 100 * width, itemheight))
                    g.FillRectangle(bo, New Rectangle(0, itemheight, shm.shield / 100 * width, itemheight))
                    g.FillRectangle(bb, New Rectangle(0, itemheight * 2, shm.mana / 100 * width, itemheight))
                    If My.Settings.ShowEnd Then
                        g.FillRectangle(If(My.Settings.DarkMode, by, bl), New Rectangle(0, itemheight * 3, shm.end / 100 * width, height - (itemheight * 3)))
                    End If
                End If
            End Using
            Return bmp
        End If

        Using g As Graphics = Graphics.FromImage(bmp), grab As Bitmap = GetClientBitmap()

            If grab Is Nothing Then Return Nothing
            If grab.Width = 0 OrElse grab.Height = 0 Then Return Nothing

            Dim rows = 3

            Dim barX As Integer = 388

            Dim BadColorCount As Integer
            Dim blackCount As Integer = 0

            For dy As Integer = 0 To 2
                BadColorCount = 0
                blackCount = 0
                For dx As Integer = 0 To 24
                    Dim currentCol As Integer = grab.GetPixel(barX + dx, 205 + dy).ToArgb
                    'Debug.Print($"current{dy}/{dx}:{currentCol:X8}")
                    If {&HFF000000, &HFF000400}.Contains(currentCol) Then
                        blackCount += 1
                    End If

                    If Not validColors.Contains(currentCol) Then
                        'If dy = 0 Then Debug.Print($"badcolor row{dy} &H{grab.GetPixel((barX + dx), (205 + dy)).ToArgb.ToString("X8")}")
                        BadColorCount += 1
                    End If
                    If dy = 0 Then
                        If BadColorCount > 1 OrElse blackCount = 25 Then
                            'Debug.Print($"Pane open? {BadColorCount} {blackCount}")
                            paneOpen = True
                            blackCount = 0
                            barX += 110
                            Exit For
                        Else
                            paneOpen = False
                        End If
                    End If
                    If BadColorCount >= 5 Then Exit For
                Next
                If BadColorCount >= 5 OrElse blackCount = 25 Then
                    rows -= 1
                    Continue For
                End If
            Next

            g.InterpolationMode = Drawing2D.InterpolationMode.NearestNeighbor
            g.PixelOffsetMode = Drawing2D.PixelOffsetMode.Half
            g.DrawImage(grab,
                        New Rectangle(0, 0, bmp.Width, bmp.Height),
                        New Rectangle(barX, 205, 25, rows),
                        GraphicsUnit.Pixel)
        End Using 'g, grab
        Return bmp
    End Function

    Friend Function GetCurrentDirectory() As String
        Return proc?.GetCurrentDirectory
    End Function

    Private alreadylaunched As Boolean = False
    Private disposedValue As Boolean

    Friend Function restart(Optional close As Boolean = True) As String
        ' Ensure process is restarted only once
        If alreadylaunched Then
            Return ""
        End If
        Dim shortcutlink As String = ""
        Try

            ' Log process details
            dBug.Print($"Restarting process: {proc.ProcessName} {proc.Id}")

            Dim mos As Management.ManagementObject = New Management.ManagementObjectSearcher($"Select * from Win32_Process WHERE ProcessID={proc.Id}").Get()(0)
            If mos Is Nothing Then Return ""
            ' Get the current command-line arguments and executable path
            Dim arguments As String = mos("commandline")
            Dim exepath As String = mos("ExecutablePath")

            'dBug.print($"Original arguments: ""{arguments}""")
            dBug.Print($"Executable Path: ""{exepath}""")

            If arguments = "" Then
                dBug.Print("Access denied! Restart process with elevated permissions.")
                Return ""
            End If

            ' Get the working directory
            Dim workdir As String = proc.GetCurrentDirectory()
            dBug.Print($"Working Directory: {workdir}")

            If Me.isSDL AndAlso workdir.EndsWith("bin") Then
                workdir = workdir.Substring(0, workdir.LastIndexOf("\"))
            End If

            If arguments.Contains(exepath) Then arguments = arguments.Replace(exepath, "")
            If arguments.Contains("""""") Then arguments = arguments.Replace("""""", "").Trim

            Dim mruNotfound As Boolean = False
            shortcutlink = MRU.FindFirstMatchingShortcut(exepath, arguments, workdir)

            If String.IsNullOrEmpty(shortcutlink) Then

                mruNotfound = True

                shortcutlink = IO.Path.Combine(FileIO.SpecialDirectories.Temp, $"ScalA\Restart\{Me.UserName}.lnk")

                If Not IO.Directory.Exists(IO.Path.Combine(FileIO.SpecialDirectories.Temp, $"ScalA\Restart\")) Then
                    IO.Directory.CreateDirectory(IO.Path.Combine(FileIO.SpecialDirectories.Temp, $"ScalA\Restart\"))
                End If

                dBug.Print($"scl: {shortcutlink}")
                ' Create a shortcut to relaunch the process
                Dim oLink As Object
                Try
                    oLink = CreateObject("WScript.Shell").CreateShortcut(shortcutlink)
                    oLink.TargetPath = exepath
                    oLink.Arguments = arguments
                    oLink.WorkingDirectory = workdir
                    oLink.WindowStyle = 1
                    oLink.Save()
                Catch ex As Exception
                    dBug.Print($"Failed to create shortcut: {ex.Message}")
                    Return shortcutlink
                End Try

            End If

            ' Close or kill the current process
            If close Then Me.CloseOrKill()

            ' Start the new process using a batch script to avoid admin rights prompt
            Dim bat As String = "\AsInvoker.bat"
            Dim tmpDir As String = FileIO.SpecialDirectories.Temp & "\ScalA"

            If Not FileIO.FileSystem.DirectoryExists(tmpDir) Then FileIO.FileSystem.CreateDirectory(tmpDir)
            If Not FileIO.FileSystem.FileExists(tmpDir & bat) Then FileIO.FileSystem.WriteAllText(tmpDir & bat, My.Resources.AsInvoker, False)

            Dim pp As New Process With {
            .StartInfo = New ProcessStartInfo With {
                .FileName = tmpDir & bat,
                .Arguments = """" & shortcutlink & """",
                .WindowStyle = ProcessWindowStyle.Hidden,
                .CreateNoWindow = True
                }
            }

            ' Start the new process
            Try
                alreadylaunched = True
                pp.Start()
                MRU.Bump(shortcutlink)
            Catch ex As Exception
                dBug.Print($"Failed to restart process: {ex.Message}")
                Try
                    FileIO.FileSystem.DeleteFile(shortcutlink)
                Catch x As Exception

                End Try
            Finally
                pp.Dispose()
                Try
                    If mruNotfound AndAlso System.IO.File.Exists(shortcutlink) Then
                        dBug.Print("Deleting shortcut")
                        Task.Run(Sub()
                                     Threading.Thread.Sleep(1000)
                                     System.IO.File.Delete(shortcutlink)
                                 End Sub)
                    End If
                Catch ex As Exception
                    dBug.Print($"Shortcut deletion failed: {ex.Message}")
                End Try
            End Try
            Return shortcutlink
        Finally
            alreadylaunched = False
        End Try
    End Function

    Friend Async Function ReOpenAsWindowed() As Task

        If alreadylaunched Then
            Return
        End If

        dBug.print($"runasWindowed: {Me.Name} {Me.Id}")

        If Not IO.Directory.Exists(IO.Path.Combine(FileIO.SpecialDirectories.Temp, "\ScalA\")) Then
            IO.Directory.CreateDirectory(IO.Path.Combine(FileIO.SpecialDirectories.Temp, "\ScalA\"))
        End If

        Dim shortcutlink As String = IO.Path.Combine(FileIO.SpecialDirectories.Temp, "\ScalA\tmp.lnk")

        Dim mos As Management.ManagementObject = New Management.ManagementObjectSearcher($“Select * from Win32_Process WHERE ProcessID={proc.Id}").Get()(0)

        Dim arguments As String = mos("commandline")

        'dBug.print($"arguments:""{arguments}""") 'leaks creds to dehub
        dBug.print($"exePath:""{mos("ExecutablePath")}""")

        If arguments = "" Then
            If CustomMessageBox.Show(FrmMain, "Access denied!" & vbCrLf &
                           "Elevate ScalA to Administrator?",
                               "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error) _
               = DialogResult.Cancel Then Return
            FrmMain.RestartSelf(True)
            End 'program
            Return
        End If
        'dBug.print("cmdline:" & arguments) 'leaks creds
        If arguments.StartsWith("""") Then
            'arguments = arguments.Substring(1) 'skipped with startindex
            arguments = arguments.Substring(arguments.IndexOf("""", 1) + 1)
        Else
            For Each exe As String In My.Settings.exe.Split(pipe, StringSplitOptions.RemoveEmptyEntries)
                If arguments.ToLower.StartsWith(exe.Trim) Then
                    arguments = arguments.Substring(exe.Trim.Length + 4) '+ ".exe".Length)
                End If
            Next
        End If

        Dim newargs As New List(Of String)

        For Each arg As String In arguments.Split("-".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Dim item As String = arg.Trim
            If item.StartsWith("w") OrElse item.StartsWith("h") Then
                Continue For
            End If
            If item.StartsWith("o") AndAlso item.Length > 1 Then
                Dim value As Integer = Val(arg.Substring(1))

                If (value And 8) = 8 Then value -= 8 'remove compact bot flag
                If (value And 256) = 256 Then value -= 256 'remove fullscreen flag

                item = $"o{value}"
            End If
            newargs.Add(item)
        Next

        newargs.Add("w800")
        newargs.Add("h600")

        arguments = Strings.Join(newargs.ToArray, " -")
        'dBug.print($"args {arguments}") 'leaks creds to debug

        Dim exepath As String = ""
        Try
            exepath = mos("ExecutablePath")
        Catch
            dBug.print("exepath exept")
        End Try
        'Dim workdir As String = exepath.Substring(0, exepath.LastIndexOf("\")) 'todo: replace with function found at https://stackoverflow.com/a/23842609/7433250
        Dim workdir = proc.GetCurrentDirectory().ToLower

        dBug.print($"wd {workdir}")

        If workdir.EndsWith("bin") Then
            workdir = workdir.Substring(0, workdir.LastIndexOf("\bin"))
        End If


        Dim oLink As Object
        Try

            oLink = CreateObject("WScript.Shell").CreateShortcut(shortcutlink)

            oLink.TargetPath = exepath
            oLink.Arguments = arguments.Trim()
            oLink.WorkingDirectory = workdir
            oLink.WindowStyle = 1
            oLink.Save()
        Catch ex As Exception
            dBug.print($"oLink exept {exepath} {workdir} {shortcutlink}") 'why is this giving an exception?
            Return
        End Try

        Dim targetName As String = Me.UserName

        'SendMessage(Me.MainWindowHandle, &H100, Keys.F12, IntPtr.Zero)
        Me.CloseOrKill()

        Dim pp As Process

        Dim bat As String = "\noAdmin.bat"
        Dim tmpDir As String = FileIO.SpecialDirectories.Temp & "\ScalA"

        If Not FileIO.FileSystem.DirectoryExists(tmpDir) Then FileIO.FileSystem.CreateDirectory(tmpDir)
        If Not FileIO.FileSystem.FileExists(tmpDir & bat) Then FileIO.FileSystem.WriteAllText(tmpDir & bat, My.Resources.AsInvoker, False)

        pp = New Process With {.StartInfo = New ProcessStartInfo With {.FileName = tmpDir & bat,
                                                                       .Arguments = """" & shortcutlink & """",
                                                                       .WindowStyle = ProcessWindowStyle.Hidden,
                                                                       .CreateNoWindow = True}}
        Try
            alreadylaunched = True
            pp.Start()
        Catch
            dBug.print("pp.start() except")
        Finally
            pp.Dispose()
        End Try



        FrmMain.Cursor = Cursors.WaitCursor
        Dim count As Integer = 0

        While True
            count += 1
            Await Task.Delay(50)
            Dim targetPPs As AstoniaProcess() = AstoniaProcess.Enumerate(FrmMain.blackList).Where(Function(ap) ap.UserName = targetName).ToArray()
            If targetPPs.Length > 0 AndAlso targetPPs(0) IsNot Nothing AndAlso targetPPs(0).Id <> 0 Then
                FrmMain.PopDropDown(FrmMain.cboAlt)
                FrmMain.cboAlt.SelectedItem = targetPPs(0)
                Exit While
            End If
            If count >= 100 Then
                CustomMessageBox.Show(FrmMain, "Windowing failed")
                Exit While
            End If
        End While
        FrmMain.Cursor = Cursors.Default

        If System.IO.File.Exists(shortcutlink) Then
            dBug.print("Deleting shortcut")
            System.IO.File.Delete(shortcutlink)
        End If

        alreadylaunched = False
    End Function

    Private Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects)
                _va?.Dispose()
                _map?.Dispose()
                proc?.Dispose()

            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override finalizer
            ' TODO: set large fields to null
            disposedValue = True
        End If
    End Sub

    ' ' TODO: override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
    ' Protected Overrides Sub Finalize()
    '     ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
    '     Dispose(disposing:=False)
    '     MyBase.Finalize()
    ' End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
    Dim Elevated As Boolean = False
    Friend Function HasExited() As Boolean
        If proc Is Nothing Then Return True
        If Elevated Then Return proc.HasExitedSafe
        Try
            Return proc.HasExited
        Catch ex As Exception
            dBug.print("HasExited Exception")
            Elevated = True
            Return proc.HasExitedSafe
        End Try
    End Function

    Friend Function IsElevated() As Boolean
        If proc Is Nothing Then Return False
        Try
            Dim dummy = proc.HasExited
        Catch ex As Exception
            Elevated = True
        End Try
        Return Elevated
    End Function

    Friend Sub setPriority(priority As Integer)
        If Me.proc Is Nothing Then Exit Sub
        Dim hProcess = OpenProcess(ProcessAccessFlags.SetInformation Or ProcessAccessFlags.QueryInformation, False, CUInt(Me.proc.Id))
        If hProcess = IntPtr.Zero Then Exit Sub
        SetPriorityClass(hProcess, priority)

        CloseHandle(hProcess)
    End Sub
End Class

NotInheritable Class AstoniaProcessSorter
    Implements IComparer(Of String)

    Private ReadOnly topOrder As List(Of String)
    Private ReadOnly botOrder As List(Of String)

    Public Sub New(topOrder As List(Of String), botOrder As List(Of String))
        Me.topOrder = topOrder
        Me.botOrder = botOrder
    End Sub

    Public Function Compare(ap1 As String, ap2 As String) As Integer Implements IComparer(Of String).Compare
        'equal = 0, ap1 > ap2 = 1, ap1 <ap2 = -1

        Dim top1 As Boolean = topOrder.Contains(ap1)
        Dim top2 As Boolean = topOrder.Contains(ap2)

        Dim bot1 As Boolean = botOrder.Contains(ap1)
        Dim bot2 As Boolean = botOrder.Contains(ap2)

        'dBug.print($"comp:{ap1} {ap2}")
        'dBug.print($"top: {top1} {top2}")
        'dBug.print($"bot: {bot1} {bot2}")

        If top1 AndAlso bot2 Then Return -1
        If bot1 AndAlso top2 Then Return 1

        If bot1 AndAlso bot2 Then Return botOrder.IndexOf(ap1) - botOrder.IndexOf(ap2)
        If top1 AndAlso top2 Then Return topOrder.IndexOf(ap1) - topOrder.IndexOf(ap2)

        If bot1 Then Return 1
        If bot2 Then Return -1

        If top1 Then Return -1
        If top2 Then Return 1

        Return Comparer(Of String).Default.Compare(ap1, ap2)

    End Function
End Class
NotInheritable Class AstoniaProcessEqualityComparer
    Implements IEqualityComparer(Of AstoniaProcess)

    Public Overloads Function Equals(ap1 As AstoniaProcess, ap2 As AstoniaProcess) As Boolean Implements IEqualityComparer(Of AstoniaProcess).Equals
        If ap1 Is Nothing AndAlso ap2 Is Nothing Then
            Return True
        ElseIf ap1 Is Nothing Or ap2 Is Nothing Then
            Return False
        ElseIf ap1.Id = ap2.Id Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Overloads Function GetHashCode(ap As AstoniaProcess) As Integer Implements IEqualityComparer(Of AstoniaProcess).GetHashCode
        Return ap.Id.GetHashCode
    End Function

End Class


Module ProcessExtensions

    <System.Runtime.CompilerServices.Extension()>
    Public Function HasExitedSafe(ByVal this As Process) As Boolean
        Dim exitCode As Integer = 0
        Dim processHandle As IntPtr = OpenProcess(ProcessAccessFlags.QueryLimitedInformation, False, this.Id)
        Try
            If processHandle <> IntPtr.Zero AndAlso GetExitCodeProcess(processHandle, exitCode) Then Return exitCode <> 259
        Catch
            CustomMessageBox.Show(FrmMain, "Exception on HasExitedSafe")
        Finally
            CloseHandle(processHandle)
        End Try
        Return True
    End Function
    ''' <summary>
    ''' Returns the executable path of a process by using QueryFullProcessImageName.
    ''' </summary>
    ''' <param name="this"></param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension()>
    Public Function Path(ByVal this As Process) As String
        Dim processPath As String = ""

        Dim processHandle As IntPtr = OpenProcess(ProcessAccessFlags.QueryLimitedInformation, False, this?.Id)
        Try
            If Not processHandle = IntPtr.Zero Then
                Dim buffer As New System.Text.StringBuilder(1024)
                If QueryFullProcessImageName(processHandle, 0, buffer, buffer.Capacity) Then
                    processPath = buffer.ToString()
                End If
            End If
        Finally
            CloseHandle(processHandle)
        End Try

        Return processPath
    End Function

    Public Function ProcessPath(ByVal pid As Integer) As String
        ProcessPath = ""

        Dim processHandle As IntPtr = OpenProcess(ProcessAccessFlags.QueryLimitedInformation, False, pid)
        Try
            If Not processHandle = IntPtr.Zero Then
                Dim buffer As New System.Text.StringBuilder(1024)
                If QueryFullProcessImageName(processHandle, 0, buffer, buffer.Capacity) Then
                    processPath = buffer.ToString()
                End If
            End If
        Finally
            CloseHandle(processHandle)
        End Try

        Return processPath
    End Function

    ''' <summary>
    ''' Returns the executable path of a process by using GetFinalPathNameByHandle.
    ''' </summary>
    ''' <param name="this"></param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension()>
    Public Function FinalPath(ByVal this As Process) As String
        Dim processPath As String = ""
        Dim sa As New SecurityAttributes
        sa.Length = Marshal.SizeOf(sa)
        Dim hFile As IntPtr = CreateFile(this.Path, FileAccess.Read, FileShare.ReadWrite, sa, CreationDisposition.OpenExisting, 0, IntPtr.Zero)


        'Dim processHandle As IntPtr = OpenProcess(ProcessAccessFlags.QueryLimitedInformation, False, this?.Id)
        Try
            If Not hFile = IntPtr.Zero Then
                Dim buffer As New System.Text.StringBuilder(1024)
                If GetFinalPathNameByHandleW(hFile, buffer, buffer.Capacity, 0) > 0 Then
                    processPath = buffer.ToString().Replace("\\?\", String.Empty)
                End If
            Else
                dBug.print("Error opening proc")
            End If
        Finally
            CloseHandle(hFile)
        End Try

        Return processPath
    End Function


    Private classCache As String = String.Empty
    Private classCacheSet As New HashSet(Of String)
    Private ReadOnly pipe As Char() = {"|"c}
    ''' <summary>
    ''' Checks if a processes classname is in pipe separated string
    ''' </summary>
    ''' <param name="pp"></param>
    ''' <param name="classes"></param>
    ''' <returns></returns>
    <System.Runtime.CompilerServices.Extension()>
    Public Function HasClassNameIn(pp As Process, classes As String) As Boolean
        If classCache <> classes Then
            classCacheSet.Clear()
            classCacheSet = New HashSet(Of String)(classes.Split(pipe, StringSplitOptions.RemoveEmptyEntries) _
                                                          .Select(Function(wc) Strings.Trim(wc)))
            classCache = classes
        End If
        Return classCacheSet.Contains(GetWindowClass(pp.MainWindowHandle))
    End Function
    Private exeCache As String = String.Empty
    Private exeCacheSet As New HashSet(Of String)
    <System.Runtime.CompilerServices.Extension()>
    Public Function IsAstonia(pp As Process) As Boolean
        If classCache <> My.Settings.className Then
            classCacheSet.Clear()
            classCacheSet = New HashSet(Of String)(My.Settings.className.Split(pipe, StringSplitOptions.RemoveEmptyEntries) _
                                                          .Select(Function(wc) Strings.Trim(wc)))
            classCache = My.Settings.className
        End If
        If exeCache <> My.Settings.exe Then
            exeCacheSet.Clear()
            exeCacheSet = New HashSet(Of String)(My.Settings.exe.Split(pipe, StringSplitOptions.RemoveEmptyEntries) _
                                                          .Select(Function(x) Strings.Trim(x)))
            exeCache = My.Settings.exe
        End If
        Return classCacheSet.Contains(GetWindowClass(pp.MainWindowHandle)) AndAlso exeCacheSet.Contains(pp.ProcessName)
    End Function

    Public Function isAstonia(h As IntPtr) As Boolean
        If classCache <> My.Settings.className Then
            classCacheSet.Clear()
            classCacheSet = New HashSet(Of String)(My.Settings.className.Split(pipe, StringSplitOptions.RemoveEmptyEntries) _
                                                          .Select(Function(wc) Strings.Trim(wc)))
            classCache = My.Settings.className
        End If
        If exeCache <> My.Settings.exe Then
            exeCacheSet.Clear()
            exeCacheSet = New HashSet(Of String)(My.Settings.exe.Split(pipe, StringSplitOptions.RemoveEmptyEntries) _
                                                          .Select(Function(x) Strings.Trim(x)))
            exeCache = My.Settings.exe
        End If
        Dim pid As Integer
        GetWindowThreadProcessId(h, pid)
        Dim procname As String = IO.Path.GetFileNameWithoutExtension(ProcessPath(pid))
        Return classCacheSet.Contains(GetWindowClass(h)) AndAlso exeCacheSet.Contains(procname)
    End Function

End Module
