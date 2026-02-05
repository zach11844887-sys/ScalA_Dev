Imports System.IO.Compression
Imports System.Net.Http

Partial Public NotInheritable Class FrmMain

    Private _altPP As AstoniaProcess

    Public ReadOnly Property showingSomeone() As Boolean
        Get
            Dim a As Boolean = My.Settings.Whitelist
            Dim b As Boolean = topSortList.Contains("Someone")
            Dim c As Boolean = botSortList.Contains("Someone")
            Dim d As Boolean = blackList.Contains("Someone")

            Dim q As Boolean = (Not d OrElse (a AndAlso (b OrElse c))) AndAlso Not (a AndAlso Not b AndAlso Not c AndAlso Not d)

            dBug.Print($"a {a} b {b} c {c} d {d}    q {q}")

            Return q
        End Get
    End Property


    Public Property AltPP As AstoniaProcess
        Get
            Return _altPP
        End Get
        Set(value As AstoniaProcess)
            If (_altPP Is Nothing AndAlso value IsNot Nothing) OrElse
               (_altPP IsNot Nothing AndAlso value Is Nothing) OrElse
               (_altPP IsNot Nothing AndAlso value IsNot Nothing AndAlso _altPP.Id <> value.Id) Then

                IPC.AddOrUpdateInstance(scalaPID, cboAlt.SelectedIndex = 0, value?.Id, showingSomeone)
                'IPC.getInstances().FirstOrDefault(Function(si) si.pid = scalaPID)
                _altPP = value
            End If
        End Set
    End Property
    'Private WndClass() As String = {"MAINWNDMOAC", "䅍义乗䵄䅏C"}
#Region " Alt Dropdown "
    Friend Sub PopDropDown(sender As ComboBox, Optional bList As Boolean = True)

        Dim current As AstoniaProcess = DirectCast(sender.SelectedItem, AstoniaProcess)
        sender.BeginUpdate()
        updatingCombobox = True

        sender.Items.Clear()
        sender.Items.Add(New AstoniaProcess) 'Someone

        sender.Items.AddRange(AstoniaProcess.Enumerate(If(bList, blackList, New List(Of String))).OrderBy(Function(ap) ap.UserName, apSorter).ToArray)

        If current IsNot Nothing AndAlso sender.Items.Contains(current) Then
            sender.SelectedItem = current
        Else
            sender.SelectedIndex = 0
        End If

        updatingCombobox = False
        sender.EndUpdate()
        sender.DropDownHeight = sender.ItemHeight * sender.Items.Count + 2
    End Sub

    Private Sub CboAlt_DropDown(sender As ComboBox, e As EventArgs) Handles cboAlt.DropDown
        pbZoom.Visible = False
        AButton.ActiveOverview = False
        PopDropDown(sender)
    End Sub
    Private Sub CmbResolution_DropDown(sender As ComboBox, e As EventArgs) Handles cmbResolution.DropDown
        moveBusy = False
        pbZoom.Visible = False
        AButton.ActiveOverview = False
    End Sub

    Private Sub ComboBoxes_DropDownClosed(sender As ComboBox, e As EventArgs) Handles cboAlt.DropDownClosed, cmbResolution.DropDownClosed
        moveBusy = False
        Dim unused = RestoreClicking()
    End Sub

#End Region
    Private Async Function RestoreClicking() As Task(Of Boolean)
        Await Task.Delay(150)
        If cboAlt.DroppedDown OrElse cmbResolution.DroppedDown OrElse cmsQuickLaunch.Visible OrElse cmsAlt.Visible OrElse cmsQuit.Visible OrElse SysMenu.Visible Then Return False
        If Not pnlOverview.Visible Then
            pbZoom.Visible = True
        Else
            AButton.ActiveOverview = My.Settings.gameOnOverview
        End If
        Return True
    End Function

    Private prevItem As New AstoniaProcess()
    Public updatingCombobox As Boolean = False
    Private Async Sub CboAlt_SelectedIndexChanged(sender As ComboBox, e As EventArgs) Handles cboAlt.SelectedIndexChanged

        If updatingCombobox Then Exit Sub
        Try

            CloseOtherDropDowns(cmsQuickLaunch.Items, Nothing)
            cmsQuickLaunch.Close()
            'IPC.AddOrUpdateInstance(scalaPID, sender.SelectedIndex = 0)

            dBug.Print($"CboAlt_SelectedIndexChanged {sender.SelectedIndex}")

            'btnAlt1.Focus()

            'If AltPP IsNot Nothing AndAlso AltPP.Id <> 0 AndAlso AltPP.Equals(CType(that.SelectedItem, AstoniaProcess)) Then
            '    AltPP.Activate()
            '    Exit Sub
            'End If

            'If AltPP.Id = 0 AndAlso that.SelectedIndex = 0 Then
            '    Exit Sub
            'End If
#If DEBUG Then
            TickCounter = 0
#End If

            Detach(False)
            AstoniaProcess.RestorePos()
            AltPP = sender.SelectedItem

            AltPP.setPriority(My.Settings.Priority)

            'IPC.AddOrUpdateInstance(scalaPID, sender.SelectedIndex = 0, AltPP.Id)
            UpdateTitle()
            UpdateDpiWarning()
            UpdateWrapperWarning()

            If sender.SelectedIndex = 0 Then
                If Not My.Settings.gameOnOverview Then
                    Try
                        AllowSetForegroundWindow(ASFW_ANY)
                        AppActivate(scalaPID)
                    Catch ex As Exception
                        dBug.Print($"Failed to activate ScalA on overview: {ex.Message}")
                    End Try
                End If
                'pnlOverview.SuspendLayout()
                pnlOverview.Show()
                AstoniaProcess.CacheCounter = 0
                AstoniaProcess.ProcCache.Clear()
                Dim visbut = UpdateButtonLayout(AstoniaProcess.Enumerate(blackList, True).Count)
                For Each but As AButton In visbut.Where(Function(b) b.Text <> "")
                    but.Image = Nothing
                    but.BackgroundImage = Nothing
                    but.Text = String.Empty
                Next
                'pnlOverview.ResumeLayout()
                tmrOverview.Enabled = True
                tmrTick.Enabled = False
                If prevItem.Id <> 0 Then DwmUnregisterThumbnail(thumb)
                sysTrayIcon.Icon = My.Resources.moa3
                prevItem = DirectCast(sender.SelectedItem, AstoniaProcess)
                PnlEqLock.Visible = False
                AOshowEqLock = False
                Me.TopMost = My.Settings.topmost
                If Not startup AndAlso My.Settings.MaxNormOverview AndAlso Me.WindowState <> FormWindowState.Maximized Then
                    btnMax.PerformClick()
                End If
                Exit Sub
            Else
                pnlOverview.Hide()
                tmrOverview.Enabled = False
                PnlEqLock.Visible = True
            End If

            If Not AltPP?.IsRunning Then
                Dim idx As Integer = sender.SelectedIndex
                sender.Items.RemoveAt(idx)
                sender.SelectedIndex = Math.Min(idx, sender.Items.Count - 1)
                Exit Sub
            End If


            If Not AltPP?.Id = 0 Then
                If AltPP.IsMinimized Then AltPP.Restore()

                'AltPP.ResetCache()

                Dim rcW As Rectangle = AltPP.WindowRect
                rcC = AltPP.ClientRect

                Dim ptC As Point
                ClientToScreen(AltPP.MainWindowHandle, ptC)

                dBug.Print($"rcW:{rcW}")
                dBug.Print($"rcC:{rcC}")
                dBug.Print($"ptC:{ptC}")

                'check if target is running as windowed. if not ask to run it with -w
                If rcC.Width = 0 AndAlso rcC.Height = 0 OrElse
                   rcC.X = ptC.X AndAlso rcC.Y = ptC.Y Then
                    'MessageBox.Show("Client is not running in windowed mode", "Error")
                    dBug.Print("Astonia Not Windowed")
                    Await AltPP.ReOpenAsWindowed()
                    'cboAlt.SelectedIndex = 0
                    Exit Sub
                End If
                AltPP.SavePos(rcW.Location)

                dBug.Print("tmrTick.Enabled")
                tmrTick.Enabled = True

                dBug.Print("AltPPTopMost " & AltPP.TopMost.ToString)
                dBug.Print("SelfTopMost " & scalaProc.IsTopMost.ToString)

                Dim item As AstoniaProcess = DirectCast(sender.SelectedItem, AstoniaProcess)
                If startThumbsDict.ContainsKey(item.Id) Then
                    dBug.Print($"reassignThumb {item.Id} {startThumbsDict(item.Id)} {item.Name}")
                    thumb = startThumbsDict(item.Id)
                End If

                For Each thumbid As IntPtr In startThumbsDict.Values
                    If thumbid = thumb Then Continue For
                    DwmUnregisterThumbnail(thumbid)
                Next
                startThumbsDict.Clear()
                If My.Settings.MaxNormOverview AndAlso WindowState = FormWindowState.Maximized Then
                    btnMax.PerformClick()
                End If
                dBug.Print($"updateThumb pbzoom {pbZoom.Size}")
                AltPP.CenterBehind(pbZoom, SetWindowPosFlags.DoNotActivate, True, True)
                If Not My.Settings.MaxNormOverview AndAlso AnimsEnabled AndAlso rectDic.ContainsKey(item.Id) Then
                    AnimateThumb(rectDic(item.Id), New Rectangle(pbZoom.Left, pbZoom.Top, pbZoom.Right, pbZoom.Bottom))
                Else
                    prevMode = 0
                    UpdateThumb(If(chkDebug.Checked, 128, 255))
                End If
                rectDic.Clear()
                sysTrayIcon.Icon = AltPP?.GetIcon

                Dim ScalAWinScaling = Me.WindowsScaling()

                dBug.Print($"{ScalAWinScaling}% ScalA windows scaling")
                dBug.Print($"{AltPP.WindowsScaling}% altPP windows scaling")

                'Dim failcounter = 0
                'If AltPP.WindowsScaling <> ScalAWinScaling Then 'scala is scaled diffrent than Alt
                '    Const timeout As Integer = 500
                '    Dim sw As Stopwatch = Stopwatch.StartNew()
                '    Do 'looped delay until alt is scaled same
                '        AltPP.CenterBehind(pbZoom, Nothing, True)
                '        Dim rc As RECT
                '        GetWindowRect(AltPP.MainWindowHandle, rc)
                '        SetWindowPos(AltPP.MainWindowHandle, ScalaHandle, rc.left + 1, rc.top + 1, -1, -1, SetWindowPosFlags.IgnoreResize Or SetWindowPosFlags.FrameChanged)
                '        dBug.print($"Scaling Delay {sw.ElapsedMilliseconds}ms {ScalAWinScaling}% vs {AltPP.WindowsScaling}")
                '        Await Task.Delay(16)
                '        If sw.ElapsedMilliseconds > timeout Then
                '            sw.Stop()
                '            dBug.print($"Scaling Delay Timeout! {failcounter}")
                '            AstoniaProcess.RestorePos(True)
                '            Await Task.Delay(16)
                '            sw = Stopwatch.StartNew()
                '            failcounter += 1
                '        End If
                '    Loop Until ScalAWinScaling = AltPP.WindowsScaling OrElse failcounter >= 3
                '    AltPP.ResetCache()
                '    UpdateThumb(If(chkDebug.Checked, 128, 255))
                'End If

                Attach(AltPP, True)

                If My.Settings.topmost Then
                    AltPP.TopMost = True
                End If

                ' Initialize zoom state for SDL2 wrapper
                ZoomStateIPC.UpdateFromFrmMain(Me, If(AltPP?.isSDL, False))

                moveBusy = False
            Else 'AltPP.Id = 0


            End If

            DoEqLock(pbZoom.Size)
            prevItem = sender.SelectedItem



            If sender.SelectedIndex > 0 Then
                pbZoom.Visible = True
                Dim sb As Rectangle = Me.Bounds
                frmOverlay.Bounds = New Rectangle(sb.X, sb.Y + 21, sb.Width, sb.Height - 21)
            End If
        Finally
            If sender.SelectedIndex = 0 Then
                AButton.ActiveOverview = My.Settings.gameOnOverview
                Dim fgw As IntPtr = GetForegroundWindow()
                dBug.Print($"FgHwnd: {fgw}")
                Dim fgt = GetWindowThreadProcessId(fgw, Nothing)

                Try
                    AttachThreadInput(fgt, ScalaThreadId, True)

                    AllowSetForegroundWindow(ASFW_ANY)
                    SetForegroundWindow(GetDesktopWindow())

                    Threading.Thread.Sleep(50)

                    AllowSetForegroundWindow(ASFW_ANY)
                    SetForegroundWindow(ScalaHandle)
                Finally
                    AttachThreadInput(fgt, ScalaThreadId, False)
                End Try
            End If
        End Try
    End Sub
    Public Function WindowsScaling() As Integer
        Dim hMon As IntPtr = MonitorFromWindow(ScalaHandle, MONITORFLAGS.DEFAULTTONEAREST)
        Dim percent As Integer = GetMonitorScaling(hMon)

        FrmSettings.pb100PWarning.Visible = percent <> 100
        Return percent
    End Function

    Public Function GetScaling(hWnd As IntPtr) As Integer
        If hWnd = IntPtr.Zero Then Return 100

        Dim hMon As IntPtr = MonitorFromWindow(hWnd, MONITORFLAGS.DEFAULTTONEAREST)
        Return GetMonitorScaling(hMon)
    End Function

    Public Resizing As Boolean

    Private Function UpdateTitle() As Boolean
        Dim titleSuff As String = String.Empty
        Dim traytooltip As String = "ScalA"
        If AltPP?.IsRunning Then
            Try
                Dim title As String = AltPP.MainWindowTitle
                If String.IsNullOrEmpty(title) Then Return False
                If AltPP.Name = "Someone" Then
                    If Not String.IsNullOrEmpty(AltPP.loggedInAs) Then
                        If AltPP?.loggedInAs <> "Someone" Then title = $"{AltPP.loggedInAs} - {title}"
                    Else
                        If AltPP?.UserName <> "Someone" Then title = $"{AltPP.UserName} - {title}"
                    End If
                End If
                titleSuff &= " - " & title
                traytooltip = title
            Catch e As Exception
                Return False
            End Try
        End If
        Me.Text = "ScalA" & titleSuff
        sysTrayIcon.Text = traytooltip.Cap(63)
        With My.Application.Info.Version
            Dim build = .Build + If(.Revision, 1, 0)
            Dim rev As String = If(.Revision, $"b{ .Revision}", "")
            lblTitle.Text = $"- ScalA v{ .Major}.{ .Minor}.{build}{rev}{titleSuff}"
        End With
        Return True
    End Function

    ''' <summary>
    ''' Updates DPI warning icon visibility based on client's DPI Override status.
    ''' Shows warning when WindowsScaling != 100% and DPI Override is not enabled.
    ''' </summary>
    Private Sub UpdateDpiWarning()
        If AltPP Is Nothing OrElse AltPP.Id = 0 Then
            pbDpiWarning.Visible = False
            Return
        End If

        ' Show warning if client is on non-100% scaled display without DPI Override
        Dim needsWarning As Boolean = AltPP.WindowsScaling <> 100 AndAlso Not AltPP.RegHighDpiAware
        pbDpiWarning.Visible = needsWarning

        ' Position the icon right after the title text
        If needsWarning Then
            pbDpiWarning.Left = lblTitle.Right + 4
        End If
    End Sub

    Private Sub PbDpiWarning_Click(sender As Object, e As MouseEventArgs) Handles pbDpiWarning.MouseUp
        If e.Button <> MouseButtons.Left Then Exit Sub

        ' Enable DPI Override for the current client
        If AltPP IsNot Nothing AndAlso AltPP.Id <> 0 Then
            Try
                AltPP.RegHighDpiAware = True
                dBug.Print($"Enabled DPI Override for {AltPP.FinalPath}")
                UpdateDpiWarning()

                ' Notify user that restart is required
                CustomMessageBox.Show(Me, "DPI Override has been enabled." & vbCrLf & vbCrLf &
                    "The client must be restarted for this change to take effect.",
                    "DPI Override Enabled", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                dBug.Print($"Failed to enable DPI Override: {ex.Message}")
                CustomMessageBox.Show(Me, "Failed to enable DPI Override: " & ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Updates SDL2 wrapper warning icon visibility based on version mismatch detection.
    ''' Shows warning when the DLL flags a version mismatch with ScalA.
    ''' </summary>
    Private Sub UpdateWrapperWarning()
        If AltPP Is Nothing OrElse AltPP.Id = 0 OrElse Not AltPP.isSDL Then
            pbWrapperWarning.Visible = False
            Return
        End If

        ' Show warning if DLL detected version mismatch
        Dim needsWarning As Boolean = ZoomStateIPC.HasVersionMismatch()
        pbWrapperWarning.Visible = needsWarning

        ' Position the icon right after the DPI warning (or title if DPI warning not visible)
        If needsWarning Then
            If pbDpiWarning.Visible Then
                pbWrapperWarning.Left = pbDpiWarning.Right + 4
            Else
                pbWrapperWarning.Left = lblTitle.Right + 4
            End If
        End If
    End Sub

    Private Sub PbWrapperWarning_Click(sender As Object, e As MouseEventArgs) Handles pbWrapperWarning.MouseUp
        If e.Button <> MouseButtons.Left Then Exit Sub

        ' Reinstall wrapper for the current client
        If AltPP IsNot Nothing AndAlso AltPP.Id <> 0 AndAlso AltPP.isSDL Then
            Dim sdl2Dir As String = SDL2WrapperHelper.GetSDL2Directory(AltPP)
            If String.IsNullOrEmpty(sdl2Dir) Then
                CustomMessageBox.Show(Me, "Could not find SDL2.dll location for this game.",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Try
                If SDL2WrapperHelper.InstallWrapper(sdl2Dir) Then
                    dBug.Print($"Reinstalled SDL2 wrapper to {sdl2Dir}")
                    pbWrapperWarning.Visible = False

                    CustomMessageBox.Show(Me, "SDL2 wrapper has been reinstalled." & vbCrLf & vbCrLf &
                        "Restart the game for changes to take effect.",
                        "Wrapper Updated", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    CustomMessageBox.Show(Me, "Failed to reinstall SDL2 wrapper." & vbCrLf &
                        "You may need to run ScalA as administrator.",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Catch ex As Exception
                dBug.Print($"Failed to reinstall wrapper: {ex.Message}")
                CustomMessageBox.Show(Me, "Failed to reinstall SDL2 wrapper: " & ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Public Shared zooms() As Size = GetResolutions()

    Private Sub FrmMain_Load(sender As Form, e As EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = True

        'cleanup of old temp files
        Try
            If IO.Directory.Exists(IO.Path.Combine(FileIO.SpecialDirectories.Temp, "ScalA\")) Then
                IO.File.Delete(IO.Path.Combine(FileIO.SpecialDirectories.Temp, "ScalA\restart.lnk"))
            End If
        Catch ex As Exception
            'fail silently
        End Try
        Try
            If IO.Directory.Exists(IO.Path.Combine(FileIO.SpecialDirectories.Temp, "ScalA\")) Then
                IO.File.Delete(IO.Path.Combine(FileIO.SpecialDirectories.Temp, "\ScalA\tmp.lnk"))
            End If
        Catch ex As Exception
            'fail silently
        End Try
        Try
            If IO.Directory.Exists(IO.Path.Combine(FileIO.SpecialDirectories.Temp, "ScalA\Restart\")) Then
                For Each lnk In IO.Directory.EnumerateFiles(IO.Path.Combine(FileIO.SpecialDirectories.Temp, "ScalA\Restart\"))
                    Try
                        IO.File.Delete(lnk)
                    Catch ex As Exception
                        'fail silently
                    End Try
                Next
            End If
        Catch ex As Exception
            'fail silently
        End Try

        If Environment.OSVersion.Version.Major < 6 AndAlso Environment.OSVersion.Version.Minor < 1 Then
            MessageBox.Show("ScalA requires Windows 7 Or later", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End
        End If


        setScalAPriority(My.Settings.Priority)

        dBug.InitDebug()

        If My.Settings.SingleInstance AndAlso IPC.AlreadyOpen Then
            IPC.RequestActivation = True
            AllowSetForegroundWindow(ASFW_ANY)
            End
        End If
        IPC.AlreadyOpen = True

        Try
            Dim ndpKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full")
            If ndpKey IsNot Nothing Then
                Dim releaseValue As Object = ndpKey.GetValue("Release")
                If releaseValue IsNot Nothing AndAlso Integer.TryParse(releaseValue.ToString(), releaseValue) Then
                    Dim release As Integer = CInt(releaseValue)
                    ' Check if the version is greater than or equal to 378675 (corresponding to .NET Framework 4.5.1)
                    If release >= 378675 Then
                        StructureToPtrSupported = True
                    End If
                End If
            End If
        Catch ex As Exception
            dBug.Print($"Failed to check .NET Framework version: {ex.Message}")
        End Try

        sysTrayIcon.Visible = True

        If New Version(My.Settings.SettingsVersion) < My.Application.Info.Version Then
            My.Settings.Upgrade()
            My.Settings.SettingsVersion = My.Application.Info.Version.ToString
            If Not My.Settings.className.Contains("SDL_app") Then
                My.Settings.className &= " | SDL_app"
            End If
            My.Settings.Save()
            zooms = GetResolutions()
            topSortList = My.Settings.topSort.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries).ToList
            botSortList = My.Settings.botSort.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries).ToList
            blackList = topSortList.Intersect(botSortList).ToList
        End If
        topSortList = topSortList.Except(blackList).ToList
        botSortList = botSortList.Except(blackList).ToList

        apSorter = New AstoniaProcessSorter(topSortList, botSortList)

#If DEBUG Then
        dBug.Print("Top:")
        topSortList.ForEach(Sub(el) dBug.Print(el))
        dBug.Print("Bot:")
        botSortList.ForEach(Sub(el) dBug.Print(el))
        dBug.Print("blacklist:")
        blackList.ForEach(Sub(el) dBug.Print(el))
#End If



        dBug.Print("mangleSysMenu")
        InitSysMenu()

        dBug.Print("topmost " & My.Settings.topmost)
        Me.TopMost = My.Settings.topmost
        Me.chkHideMessage.Checked = My.Settings.hideMessage


        'left big in designer to facilitate editing
        cornerNW.Size = New Size(2, 2)
        cornerNE.Size = New Size(2, 2)
        cornerSW.Size = New Size(2, 2)
        cornerSE.Size = New Size(2, 2)

        cmbResolution.Items.Add($"{My.Settings.resol.Width}x{My.Settings.resol.Height}")
        cmbResolution.Items.AddRange(zooms.Select(Function(ss) ss.Width & "x" & ss.Height).ToArray)
        cmbResolution.SelectedIndex = My.Settings.zoom


        dBug.Print("location " & My.Settings.location.ToString)
        suppressWM_MOVEcwp = True
        Me.Location = My.Settings.location
        suppressWM_MOVEcwp = False

        Dim args() As String = Environment.GetCommandLineArgs()

        cboAlt.BeginUpdate()
        cboAlt.Items.Add(New AstoniaProcess()) 'someone
        cboAlt.SelectedIndex = 0
        Dim APlist As List(Of AstoniaProcess) = AstoniaProcess.Enumerate(blackList).ToList
        For Each ap As AstoniaProcess In APlist
            cboAlt.Items.Add(ap)
            If args.Count > 1 AndAlso ap.UserName = args(1) Then
                dBug.Print($"Selecting '{ap.UserName}'")
                cboAlt.SelectedItem = ap
            End If
        Next
        cboAlt.EndUpdate()

        If cmbResolution.SelectedIndex = 0 Then
            ReZoom(My.Settings.resol)
        End If

        AddAButtons()
        UpdateButtonLayout(APlist.Count)

        If Not My.Settings.AlwaysStartOnOverview AndAlso cboAlt.SelectedIndex = 0 AndAlso args.Count = 1 Then
            dBug.Print("Selecting Default")
            cboAlt.SelectedIndex = If(cboAlt.Items.Count = 2 AndAlso Not My.Settings.gameOnOverview, 1, 0)
        End If

        dBug.Print("updateTitle")
        UpdateTitle()

        Dim progdata As String = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) & "\ScalA"
        If My.Settings.links = "" Then
            My.Settings.links = progdata
        End If
        If Not System.IO.Directory.Exists(progdata) Then
            System.IO.Directory.CreateDirectory(progdata)
            System.IO.Directory.CreateDirectory(progdata & "\Example Folder")
        End If




        'Dim image = LoadImage(IntPtr.Zero, "#106", 1, 16, 16, 0)
        'If image <> IntPtr.Zero Then
        '    Using Ico = Icon.FromHandle(image)
        '        bmShield = Ico.ToBitmap
        '        DestroyIcon(Ico.Handle)
        '    End Using
        'End If
        ''set shield if runing as admin
        'If My.User.IsInRole(ApplicationServices.BuiltInRole.Administrator) Then
        '    btnStart.Image = bmShield
        'End If

        If System.IO.File.Exists(FileIO.SpecialDirectories.Temp & "\ScalA\tmp.lnk") Then
            dBug.Print("Deleting shortcut")
            System.IO.File.Delete(FileIO.SpecialDirectories.Temp & "\ScalA\tmp.lnk")
        End If

        FrmBehind.Show()
        frmOverlay.Show(Me)
        FrmSizeBorder.Show(Me)

        suppressWM_MOVEcwp = True

        If cmbResolution.SelectedIndex = 0 Then DoEqLock(My.Settings.resol)

        cmsQuickLaunch.Renderer = New CustomToolStripRenderer()
        CType(cmsQuickLaunch.Renderer, CustomToolStripRenderer).InitAnimationTimer(cmsQuickLaunch)

        cmsAlt.Renderer = New ToolStripProfessionalRenderer(New CustomColorTable)
        cmsQuit.Renderer = cmsAlt.Renderer
        frmOverlay.cmsRestart.Renderer = cmsAlt.Renderer

        If My.Settings.Theme = 0 Then 'undefined, system, light, dark
            If My.Settings.DarkMode Then
                My.Settings.Theme = 3
            Else
                My.Settings.Theme = 1
            End If
        End If

        Dim darkmode As Boolean = WinUsingDarkTheme()

        If My.Settings.Theme = 3 Then
            darkmode = True
        End If
        If My.Settings.Theme = 2 Then
            darkmode = False
        End If
        ApplyTheme(darkmode)

        dBug.Print($"Anims {AnimsEnabled}")

        tmrOverview.Interval = If(My.Settings.gameOnOverview, 30, 60)

        AddHandler Application.Idle, AddressOf Application_Idle

        If My.Settings.DisableWinKey OrElse My.Settings.OnlyEsc OrElse My.Settings.NoAltTab Then keybHook.Hook()

        clipBoardInfo = GetClipboardFilesAndAction()
        registerClipListener()

        'spawn IPC waiter thread
        Dim waitThread As New Threading.Thread(AddressOf IPC.SelectSemaThread) With {.IsBackground = True}
        waitThread.Start(Me)

        'spawn priority setter thread
        Dim priThread As New Threading.Thread(AddressOf prioritySetter) With {.IsBackground = True, .Priority = Threading.ThreadPriority.Lowest}
        priThread.Start()

        WorkerThread = New Threading.Thread(AddressOf WorkerLoop) With {.IsBackground = True, .Priority = Threading.ThreadPriority.Highest}
        WorkerThread.Start()

        StartCloseErrorDialogThread()

    End Sub

    Private Sub prioritySetter()
        Do
            Threading.Thread.Sleep(500)
            If pnlOverview.IsHandleCreated AndAlso Not pnlOverview.IsDisposed AndAlso My.Settings.gameOnOverview AndAlso Me.Invoke(Function() pnlOverview.Visible) Then
                Threading.Thread.Sleep(25)
                Dim ablist As List(Of AButton) = Me.Invoke(Function() pnlOverview.Controls.OfType(Of AButton).Where(Function(a) a.Visible AndAlso a.AP IsNot Nothing).ToList)
                For Each ab As AButton In ablist
                    Threading.Thread.Sleep(25)
                    If Not ab?.IsDisposed Then ab.AP?.setPriority(My.Settings.Priority)
                Next
            End If
        Loop
    End Sub

    Protected Overrides Sub OnHandleDestroyed(e As EventArgs)
        MyBase.OnHandleDestroyed(e)
    End Sub

    'Public bmShield As Bitmap


    'Public startup As Boolean = True

    Private Sub FrmMain_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        dBug.Print("FrmMain_Shown")
        suppressWM_MOVEcwp = False
        If Not Screen.AllScreens.Any(Function(s) s.WorkingArea.Contains(Me.Location)) Then
            dBug.Print("location out of bounds")
            Dim msWA As Rectangle = Screen.PrimaryScreen.WorkingArea
            Me.Location = New Point(Math.Max(msWA.Left, (msWA.Width - Me.Width) / 2), Math.Max(msWA.Top, (msWA.Height - Me.Height) / 2))
        End If
        startup = False
        If My.Settings.StartMaximized Then
            btnMax.PerformClick()
        End If
        If cboAlt.SelectedIndex > 0 Then
            tmrTick.Start()
            moveBusy = False
        End If
        FrmBehind.Bounds = Me.RectangleToScreen(pbZoom.Bounds)
        If My.Settings.CheckForUpdate Then
            UpdateCheck()
        End If
        If cboAlt.SelectedIndex = 0 AndAlso My.Settings.MaxNormOverview AndAlso Me.WindowState = FormWindowState.Normal Then
            btnMax.PerformClick()
        End If
        FrmSizeBorder.Opacity = If(My.Settings.SizingBorder AndAlso Me.WindowState = FormWindowState.Normal, 0.01, 0)

        IPC.AddOrUpdateInstance(scalaPID, cboAlt.SelectedIndex = 0, If(cboAlt.SelectedIndex = 0, Nothing, cboAlt.SelectedItem?.id), showingSomeone)

        'CheckScreenScalingMode()

        Dim sb As Rectangle = Me.RectangleToScreen(pbZoom.Bounds)
        frmOverlay.Bounds = sb 'New Rectangle(sb.X, sb.Y + 21, sb.Width, sb.Height - 21)
        'SetWindowPos(frmOverlay.Handle, 0, sb.X, sb.Y, sb.Width, sb.Height - 21, SetWindowPosFlags.DoNotActivate Or SetWindowPosFlags.DoNotChangeOwnerZOrder)

    End Sub
    Friend Shared updateToVersion As String = "Error"
    Friend Shared ReadOnly client As HttpClient = New HttpClient() With {.Timeout = TimeSpan.FromMilliseconds(5000)}
    Friend Shared Async Sub UpdateCheck()
        Try
            Using clnt As HttpClient = New HttpClient()
                Using response As HttpResponseMessage = Await clnt.GetAsync("https://github.com/smoorke/ScalA/releases/download/ScalA/version")
                    response.EnsureSuccessStatusCode()
                    Dim responseBody As String = Await response.Content.ReadAsStringAsync()

                    If New Version(responseBody) > My.Application.Info.Version Then
                        FrmMain.pnlUpdate.Visible = True
                    Else
                        FrmMain.pnlUpdate.Visible = False
                    End If
                    updateToVersion = responseBody
                End Using
            End Using
        Catch ex As Exception
            FrmMain.pnlUpdate.Visible = False
            updateToVersion = "Error"
        End Try
    End Sub
    Friend Shared Async Function UpdateDownload(self As Form) As Task
        Try
            Using response As HttpResponseMessage = Await client.GetAsync("https://github.com/smoorke/ScalA/releases/download/ScalA/ScalA.exe")
                response.EnsureSuccessStatusCode()
                Dim responseBody As Byte() = Await response.Content.ReadAsByteArrayAsync()

                If Not FileIO.FileSystem.DirectoryExists(FileIO.SpecialDirectories.Temp & "\ScalA\") Then
                    FileIO.FileSystem.CreateDirectory(FileIO.SpecialDirectories.Temp & "\ScalA\")
                End If

                FileIO.FileSystem.WriteAllBytes(FileIO.SpecialDirectories.Temp & "\ScalA\ScalA.exe", responseBody, False)

            End Using
        Catch e As Exception
            CustomMessageBox.Show(self, "Error" & vbCrLf & e.Message)
        End Try
    End Function
    Friend Shared Async Function LogDownload(self As Form) As Task
        Try
            Using response As HttpResponseMessage = Await client.GetAsync("https://github.com/smoorke/ScalA/releases/download/ScalA/ChangeLog.txt")
                response.EnsureSuccessStatusCode()
                Dim responseBody As Byte() = Await response.Content.ReadAsByteArrayAsync()

                If Not FileIO.FileSystem.DirectoryExists(FileIO.SpecialDirectories.Temp & "\ScalA\") Then
                    FileIO.FileSystem.CreateDirectory(FileIO.SpecialDirectories.Temp & "\ScalA\")
                End If

                FileIO.FileSystem.WriteAllBytes(FileIO.SpecialDirectories.Temp & "\ScalA\ChangeLog.txt", responseBody, False)

            End Using
        Catch e As Exception
            CustomMessageBox.Show(self, "Error" & vbCrLf & e.Message)
        End Try
    End Function
    Friend Async Sub pbUpdateAvailable_Click(sender As PictureBox, e As MouseEventArgs) Handles pbUpdateAvailable.MouseDown

        If e.Button <> MouseButtons.Left Then Exit Sub

        Await LogDownload(Me)

        If UpdateDialog.ShowDialog(Me) <> DialogResult.OK Then
            Exit Sub
        End If

        Dim MePath As String = Environment.GetCommandLineArgs(0)
        Dim sb As New Text.StringBuilder(260)
        Dim len = sb.Capacity
        Try
            Dim drive = Strings.Left(MePath, 2)
            If WNetGetConnection(drive, sb, len) = 0 Then MePath = MePath.Replace(drive, sb.ToString)
        Catch
            dBug.Print($"WNetGetConnection Exception")
        End Try
        SaveLocation()
        My.Settings.Save()
        tmrOverview.Stop()
        tmrTick.Stop()
        Await UpdateDownload(Me)
        AstoniaProcess.RestorePos()
        Try
            If Not FileIO.FileSystem.DirectoryExists(FileIO.SpecialDirectories.Temp & "\ScalA\") Then
                FileIO.FileSystem.CreateDirectory(FileIO.SpecialDirectories.Temp & "\ScalA\")
            End If
            'FileIO.FileSystem.WriteAllBytes(FileIO.SpecialDirectories.Temp & "\ScalA\ScalA_Updater.exe", My.Resources.ScalA_Updater, False)

            Using zipStream As New IO.MemoryStream(My.Resources.ScalA_Updater_Zip)
                Using archive As New ZipArchive(zipStream, ZipArchiveMode.Read)
                    For Each entry As ZipArchiveEntry In archive.Entries
                        Dim fullPath As String = IO.Path.Combine(FileIO.SpecialDirectories.Temp & "\ScalA\", entry.FullName)

                        ' Ensure directory structure exists
                        Dim dir As String = IO.Path.GetDirectoryName(fullPath)
                        If Not IO.Directory.Exists(dir) Then
                            IO.Directory.CreateDirectory(dir)
                        End If

                        ' Skip directories
                        If entry.FullName.EndsWith("/") OrElse entry.FullName.EndsWith("\") Then
                            Continue For
                        End If

                        ' Extract entry to disk (overwrite = True)
                        entry.ExtractToFile(fullPath, True)

                    Next
                End Using
            End Using

            Dim si As New ProcessStartInfo With {
                                       .FileName = FileIO.SpecialDirectories.Temp & "\ScalA\ScalA_Updater.exe",
                                       .Arguments = $"""{MePath}"""
                          }
            If Not IsDirectoryWritable(IO.Path.GetDirectoryName(MePath)) Then si.Verb = "runas"
            If MePath <> Environment.GetCommandLineArgs(0) Then si.Arguments &= $" ""{IO.Directory.GetCurrentDirectory().TrimEnd("\")}"" ""{Environment.GetCommandLineArgs(0)}"""

            Process.Start(si)
            sysTrayIcon.Visible = False
            End
        Catch ex As Exception
            dBug.Print($"Failed to start updater: {ex.Message}")
        End Try
        If cboAlt.SelectedIndex = 0 Then
            tmrOverview.Start()
        Else
            tmrTick.Start()
        End If
    End Sub

    Public Sub ApplyTheme(darkmode As Boolean)
        My.Settings.DarkMode = darkmode
        If darkmode Then
            pnlOverview.BackColor = Color.Gray
            pnlButtons.BackColor = Colors.GreyBackground
            pnlSys.BackColor = Colors.GreyBackground
            Me.BackColor = Colors.GreyBackground
            btnStart.BackColor = Colors.GreyBackground
            pnlTitleBar.BackColor = Colors.GreyBackground
            lblTitle.ForeColor = Colors.LightText
            ChkEqLock.ForeColor = Color.Gray
#If DEBUG Then
            chkDebug.ForeColor = Colors.LightText
#End If
        Else
            pnlOverview.BackColor = Color.FromKnownColor(KnownColor.Control)
            pnlButtons.BackColor = Color.FromKnownColor(KnownColor.Control)
            pnlSys.BackColor = Color.FromKnownColor(KnownColor.Control)
            Me.BackColor = Color.FromKnownColor(KnownColor.Control)
            btnStart.BackColor = Color.FromKnownColor(KnownColor.Control)
            pnlTitleBar.BackColor = Color.FromKnownColor(KnownColor.Control)
            lblTitle.ForeColor = Color.Black
            ChkEqLock.ForeColor = Color.Black
#If DEBUG Then
            chkDebug.ForeColor = Color.Black
#End If
        End If

        Dim bBc As Color = If(darkmode, Color.DarkGray, COLOR_LIGHT_GRAY)
        Task.Run(Sub() Parallel.ForEach(pnlOverview.Controls.OfType(Of AButton), Sub(but) but.BackColor = bBc))
        PnlEqLock.BackColor = bBc

        Dim bFaBc As Color = If(darkmode, Colors.GreyBackground, Color.FromKnownColor(KnownColor.Control))
        For Each but As Button In pnlButtons.Controls.OfType(Of Button)
            but.FlatAppearance.BorderColor = bFaBc
        Next

        cmbResolution.DarkTheme = darkmode
        cboAlt.DarkTheme = darkmode
        btnStart.DarkTheme = darkmode

        dBug.Print($"Environment.OSVersion.Version {Environment.OSVersion.Version}", 1)

        'Theme SysMenu and similar menus
        If Environment.OSVersion.Version.Major = 10 Then
            ' Windows 10 or 11
            Dim build = Environment.OSVersion.Version.Build
            'dBug.Print($"Os.build:{build}")
            If build >= 17763 Then
                SetPreferredAppMode(If(darkmode, 2, 3))
                FlushMenuThemes()
            End If
        End If

    End Sub

    Private Shared Function GetResolutions() As Size()
        Dim reslist As New List(Of Size)
        For Each line As String In My.Settings.resolutions.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Dim parts() As String = line.ToUpper.Split("X")
            dBug.Print(parts(0) & " " & parts(1))
            reslist.Add(New Size(parts(0), parts(1)))
        Next

        Return reslist.ToArray
    End Function



#Region " Move Self "

    Private MovingForm As Boolean
    Private MoveForm_MousePosition As Point
    Private caption_Mousedown As Boolean = False
    Private captionMoveTrigger As Boolean = False
    'Private QLwasOpenCaptDragDelay As Boolean = False
    'Private doubleclicktimer As Stopwatch = Stopwatch.StartNew

    Public Sub MoveForm_MouseDown(sender As Control, e As MouseEventArgs) Handles pnlTitleBar.MouseDown, lblTitle.MouseDown
        'Me.TopMost = True
        'setActive(True)
        'Me.Invalidate(True)

        'If FrmSettings.chkDoAlign.Checked Then
        '    If Me.WindowState <> FormWindowState.Maximized AndAlso e.Button = MouseButtons.Left Then
        '        MovingForm = True
        '        MoveForm_MousePosition = e.Location
        '    End If
        'Else
        dBug.Print("Caption.MouseDown")
        If e.Button = MouseButtons.Left AndAlso e.Clicks = 1 Then ' AndAlso doubleclicktimer.ElapsedMilliseconds > 500 Then
            'dBug.Print($"dctimer {doubleclicktimer.ElapsedMilliseconds}")
            'doubleclicktimer.Restart()
            'AltPP?.ThreadInput(False)

            sender.Capture = False
            tmrTick.Stop()
            caption_Mousedown = True
            If Me.WindowState = FormWindowState.Maximized Then
                captionMoveTrigger = True
                wasMaximized = True
            End If

            'Task.Run(Sub()
            '             Threading.Thread.Sleep(500)
            '             If Not caption_Mousedown Then Exit Sub
            '             Detach(False) 'enable smooth dragging on lecacy clients with sleeps, breaks double click in a weird way hence the delayed task and DCTimer
            '         End Sub)
            Dim msg As Message = Message.Create(ScalaHandle, WM_NCLBUTTONDOWN, New IntPtr(HTCAPTION), IntPtr.Zero)
            Me.WndProc(msg)

            'AltPP?.ThreadInput(True)

            caption_Mousedown = False
            captionMoveTrigger = False
            If Not pnlOverview.Visible Then
                Attach(AltPP, True)
                tmrTick.Start()
            End If
            dBug.Print("movetimer stopped")
            'FrmSizeBorder.Bounds = Me.Bounds
        End If
        'End If
    End Sub

    Public Sub MoveForm_MouseMove(sender As Control, e As MouseEventArgs) Handles _
    pnlTitleBar.MouseMove, lblTitle.MouseMove ' Add more handles here (Example: PictureBox1.MouseMove)
        If MovingForm Then
            Dim newoffset As Point = e.Location - MoveForm_MousePosition
            Me.Location += newoffset
            'If FrmSettings.chkDoAlign.Checked Then
            'FrmSettings.ScalaMoved += newoffset
            'End If
        End If
    End Sub

    Public Sub MoveForm_MouseUp(sender As Control, e As MouseEventArgs) Handles pnlTitleBar.MouseUp, lblTitle.MouseUp
        ' only fires when settings.chkAlign is on
        If e.Button = MouseButtons.Left Then
            dBug.Print("Mouseup")
            MovingForm = False
            If AltPP?.IsRunning Then ' AndAlso Not FrmSettings.chkDoAlign.Checked Then
                AltPP?.CenterBehind(pbZoom, SetWindowPosFlags.DoNotActivate Or SetWindowPosFlags.ASyncWindowPosition)
            End If
        End If
    End Sub

#End Region
    Public Sub SaveLocation()
        If Me.WindowState = FormWindowState.Normal Then
            My.Settings.location = Me.Location
        Else
            My.Settings.location = Me.RestoreBounds.Location
        End If
    End Sub

    Private Sub FrmMain_Closing(sender As Form, e As EventArgs) Handles Me.Closing
        AstoniaProcess.RestorePos()
        SaveLocation()
        tmrActive.Stop()
        Hotkey.UnregHotkey(Me)
    End Sub


    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.Style = cp.Style Or WindowStyles.WS_SYSMENU Or WindowStyles.WS_MINIMIZEBOX
            'cp.ExStyle = cp.ExStyle Or WindowStylesEx.WS_EX_COMPOSITED
            'cp.ClassStyle = cp.ClassStyle Or CS_DROPSHADOW
            Return cp
        End Get
    End Property

    Public Sub CmbResolution_MouseUp(sender As ComboBox, e As MouseEventArgs) Handles cmbResolution.MouseUp
        If e.Button = MouseButtons.Right Then
            UntrapMouse(MouseButtons.Right)
            CloseOtherDropDowns(cmsQuickLaunch.Items, Nothing)
            cmsQuickLaunch.Close()
            If sender.Contains(MousePosition) Then
                FrmSettings.Tag = FrmSettings.tabResolutions
                FrmSettings.Show()
                FrmSettings.WindowState = FormWindowState.Normal
                FrmSettings.Activate()
            End If
        End If
    End Sub

    Public suppressResChange As Boolean = False
    Public Sub CmbResolution_SelectedIndexChanged(sender As ComboBox, e As EventArgs) Handles cmbResolution.SelectedIndexChanged
        moveBusy = False
        If sender.SelectedIndex = 0 Then Exit Sub
        dBug.Print($"cboResolution_SelectedIndexChanged {sender.SelectedIndex}")

        My.Settings.zoom = sender.SelectedIndex
        My.Settings.resol = zooms(sender.SelectedIndex - 1)

        If WindowState = FormWindowState.Maximized Then
            btnMax.PerformClick()
            wasMaximized = False
            Exit Sub
        End If

        sender.Items(0) = $"{My.Settings.resol.Width}x{My.Settings.resol.Height}"
        DoEqLock(My.Settings.resol)

        If suppressResChange Then Exit Sub
        ReZoom(My.Settings.resol)
        AltPP?.CenterBehind(pbZoom)

    End Sub
    Private Sub cmbResolution_MouseWheel(sender As ComboBox, e As MouseEventArgs) Handles cmbResolution.MouseWheel
        If sender.SelectedIndex = 0 AndAlso Not sender.DroppedDown Then
            ' Stop default behavior
            DirectCast(e, HandledMouseEventArgs).Handled = True

            If Me.WindowState = FormWindowState.Maximized Then
                Dim mp As Point = e.Location
                btnMax.PerformClick()
                Cursor.Position = sender.PointToScreen(mp)
                Exit Sub
            End If

            ' Parse current resolution
            Dim currentRes() As String = sender.SelectedItem.ToString().Split("x"c)
            Dim size = New Size(Val(currentRes(0)), Val(currentRes(1)))

            ' Find the closest resolution using zooms array
            Dim closest As Integer = FindClosestResolution(size, e.Delta)

            dBug.Print($"closest{closest}")
            ' Set the index to the found item
            If closest > -1 Then sender.SelectedIndex = closest + 1 '+1 because idx0 is the custom res which isn't stored in zooms
        End If

        dBug.Print($"cmbRes mwh {sender.SelectedItem} {sender.SelectedIndex} {e.Delta}")
    End Sub

    Private Function FindClosestResolution(targetSize As Size, delta As Integer) As Integer

        Dim minDiff As Integer = Integer.MaxValue
        Dim closest As Integer = -1

        For idx As Integer = 0 To zooms.Length - 1
            Dim res As Size = zooms(idx)

            ' Calculate difference (absolute sum of width and height differences)
            Dim diff As Integer = Math.Abs(res.Width - targetSize.Width) + Math.Abs(res.Height - targetSize.Height)

            ' Update closest if a smaller difference is found
            If diff < minDiff Then
                minDiff = diff
                closest = idx
                ' Adjust based on scroll direction 'note zooms needs to be sorted for this
                If delta > 0 AndAlso res.Width >= targetSize.Width AndAlso res.Height >= targetSize.Height AndAlso idx > 0 Then
                    closest = idx - 1 ' Move to a smaller resolution
                ElseIf delta < 0 AndAlso res.Width <= targetSize.Width AndAlso res.Height <= targetSize.Height AndAlso idx < zooms.Length - 1 Then
                    closest = idx + 1 ' Move to a larger resolution
                End If
            End If
        Next

        Return closest
    End Function

    Public Sub ReZoom(newSize As Size)
        dBug.Print($"reZoom {newSize}")
        'Me.SuspendLayout()
        suppressResChange = True
        If Me.WindowState <> FormWindowState.Maximized Then
            Me.Size = New Size(newSize.Width + 2, newSize.Height + pnlTitleBar.Height + 1)
            pbZoom.Left = 1
            pnlOverview.Left = 1
            pbZoom.Size = newSize
            pnlOverview.Size = newSize
            cmbResolution.Enabled = True

            'suppressResChange = True
            cmbResolution.SelectedIndex = My.Settings.zoom

        Else 'FormWindowState.Maximized
            pbZoom.Left = 0
            pnlOverview.Left = 0
            pbZoom.Width = newSize.Width
            pbZoom.Height = newSize.Height - pnlTitleBar.Height
            pnlOverview.Size = pbZoom.Size

            'suppressResChange = True
            cmbResolution.SelectedIndex = 0

        End If

        'If Me.WindowsScaling = 175 Then
        '    pbZoom.Top = pnlTitleBar.Height - 1
        '    pbZoom.Height = newSize.Height + 1
        '    If Me.WindowState = FormWindowState.Maximized Then
        '        pbZoom.Height -= pnlTitleBar.Height
        '    End If
        'End If
        cmbResolution.Items(0) = $"{pbZoom.Size.Width}x{pbZoom.Size.Height}"
        suppressResChange = False
        'Me.ResumeLayout(True)

        If cboAlt.SelectedIndex <> 0 Then
            dBug.Print("updateThumb")
            UpdateThumb(If(chkDebug.Checked, 128, 255))
        End If
        pnlTitleBar.Width = newSize.Width - pnlButtons.Width - pnlSys.Width
        dBug.Print($"rezoom pnlTitleBar.Width {pnlTitleBar.Width}")

        cornerNW.Location = New Point(0, 0)
        cornerNE.Location = New Point(Me.Width - 2, 0)
        cornerSW.Location = New Point(0, Me.Height - 2)
        cornerSE.Location = New Point(Me.Width - 2, Me.Height - 2)

        DoEqLock(newSize)

        If Me.WindowState <> FormWindowState.Maximized AndAlso My.Settings.roundCorners Then

            cornerNW.Visible = True
            cornerNE.Visible = True
            cornerSW.Visible = True
            cornerSE.Visible = True

        Else 'maximized
            cornerNW.Visible = False
            cornerNE.Visible = False
            cornerSW.Visible = False
            cornerSE.Visible = False
        End If

        ' Update zoom state for SDL2 wrapper (viewport size changed)
        ZoomStateIPC.UpdateFromFrmMain(Me, If(AltPP?.isSDL, False))

    End Sub

    Private Sub DoEqLock(newSize As Size)
        If rcC.Width = 0 Then
            PnlEqLock.Location = New Point(138.Map(800, 0, newSize.Width, 0), 25)
            PnlEqLock.Size = New Size(524.Map(0, 800, 0, newSize.Width),
                                       45.Map(0, 600, 0, newSize.Height))
        Else
#If True Then
            'Dim zoom As Size = My.Settings.resol
            Dim zoom As Size = newSize
            If cmbResolution.SelectedIndex > 0 Then
                zoom = zooms(cmbResolution.SelectedIndex - 1)
            End If
            dBug.Print($"DoEqLock zoom {zoom}")
            PnlEqLock.Location = New Point(CType(rcC.Width / 2 - 262.Map(0, 800, 0, rcC.Width), Integer).Map(rcC.Width, 0, zoom.Width, 0), 25)
            Dim excludGearLock As Integer = If(AltPP?.isSDL, 18, 0)
            Dim lockHeight = 45
            If rcC.Height >= 2000 Then
                lockHeight += 120
            ElseIf rcC.Height >= 1500 Then
                lockHeight += 80
            ElseIf rcC.Height >= 1000 Then
                lockHeight += 40
            End If
            PnlEqLock.Size = New Size((524 - excludGearLock).Map(0, 800, 0, rcC.Width).Map(0, rcC.Width, 0, zoom.Width),
                                       lockHeight.Map(0, rcC.Height, 0, zoom.Height))
#Else
            PnlEqLock.Location = New Point(CType(rcC.Width / 2 - 262, Integer).Map(rcC.Width, 0, zooms(cmbResolution.SelectedIndex).Width, 0), 25)
            PnlEqLock.Size = New Size(524.Map(rcC.Width, 0, zooms(cmbResolution.SelectedIndex).Width, 0),
                                       45.Map(0, rcC.Height, 0, zooms(cmbResolution.SelectedIndex).Height))
#End If
        End If
    End Sub

    ''' <summary>
    ''' Fix mousebutton stuck after drag bug
    ''' Note: needs to be run before activating self
    ''' </summary>
    Public Sub UntrapMouse(button As MouseButtons)
        Dim activePID = GetActiveProcessID()
        'dBug.print($"active {activePID} is AltPP.id {activePID = AltPP?.Id}")
        dBug.Print($"Untrap ""{AltPP?.Name}"" ""{PrevMouseAlt?.Name}""")
        Dim ap As AstoniaProcess = If(PrevMouseAlt, AltPP)
        If activePID <> ap?.Id Then Exit Sub 'only when dragged from client
        Try
            'If My.Settings.gameOnOverview OrElse (Not pnlOverview.Visible AndAlso Not pbZoom.Contains(MousePosition)) Then
            dBug.Print($"untrap {ap?.Name} mouse {button}")
            If button = MouseButtons.Right Then PostMessage(ap?.MainWindowHandle, WM_RBUTTONUP, 0, 0)
            If button = MouseButtons.Middle Then PostMessage(ap?.MainWindowHandle, WM_MBUTTONUP, 0, 0)
            'End If
        Catch
        Finally
            MouseButtonStale = MouseButtons.None
        End Try
    End Sub
    ''' <summary>
    ''' wasMaximized is used to determine what state to restore to
    ''' </summary>
    Dim wasMaximized As Boolean = False

    Public moveBusy As Boolean = False
    Dim suppressWM_MOVEcwp As Boolean = False

    Private Sub Cycle(Optional up As Boolean = False)
        cboAlt.DroppedDown = False
        tmrTick.Enabled = False
        'PostMessage(AltPP.MainWindowHandle, WM_RBUTTONUP, 0, 0)'couses look to be sent when cycle hotkey contains ctrl
        'PostMessage(AltPP.MainWindowHandle, WM_MBUTTONUP, 0, 0)'causes alt to attack when hotkey contains ctrl
        PopDropDown(cboAlt)
        AstoniaProcess.RestorePos(True)
        If Me.WindowState = FormWindowState.Minimized Then
            SendMessage(ScalaHandle, WM_SYSCOMMAND, SC_RESTORE, IntPtr.Zero)
        End If
        Dim requestedindex = cboAlt.SelectedIndex + If(up, -1, 1)
        If requestedindex < 1 Then
            requestedindex = cboAlt.Items.Count - 1
        End If
        If requestedindex >= cboAlt.Items.Count Then
            requestedindex = 1
        End If
        If requestedindex >= cboAlt.Items.Count Then
            cboAlt.SelectedIndex = 0
            tmrOverview.Enabled = True
            tmrTick.Enabled = False
            pbZoom.Hide()
            pnlOverview.Show()
            sysTrayIcon.Icon = My.Resources.moa3
            Detach(True)
            Exit Sub
        End If
        cboAlt.SelectedIndex = requestedindex
        Me.Activate()
        Me.BringToFront()
        If requestedindex > 0 Then tmrTick.Start()
        Try
            If AltPP IsNot Nothing Then
                AppActivate(AltPP.Id)
            Else
                AppActivate(scalaPID)
            End If
        Catch ex As Exception
            dBug.Print($"Failed to activate window after resize: {ex.Message}")
        End Try
    End Sub


    'Delegate Sub updateButtonImageDelegate(but As AButton, bm As Bitmap)
    'Private Shared ReadOnly updateButtonImage As New updateButtonImageDelegate(AddressOf UpdateButtonImageMethod)
    'Private Shared Sub UpdateButtonImageMethod(but As AButton, bm As Bitmap)
    '    If but Is Nothing Then Exit Sub
    '    but.Image = bm
    'End Sub
    'Delegate Sub updateButtonBackgroundImageDelegate(but As AButton, bm As Bitmap)
    'Private Shared ReadOnly updateButtonBackgroundImage As New updateButtonImageDelegate(AddressOf UpdateButtonBackgroundImageMethod)
    'Private Shared Sub UpdateButtonBackgroundImageMethod(but As AButton, bm As Bitmap)
    '    If but Is Nothing Then Exit Sub
    '    but.BackgroundImage = bm
    'End Sub
    Private Function GetNextPerfectSquare(num As Integer) As Integer
        Dim nextN As Integer = Math.Floor(Math.Sqrt(num)) + 1
        If nextN > 6 Then nextN = 6
        Return nextN * nextN
    End Function
    Private Sub AddAButtons()
        pnlOverview.SuspendLayout()
        For i As Integer = 1 To 42
            Dim but As New AButton("", 0, 0, 200, 150)

            AddHandler but.Click, AddressOf BtnAlt_Click
            AddHandler but.MouseDown, AddressOf BtnAlt_MouseDown
            AddHandler but.MouseEnter, AddressOf BtnAlt_MouseEnter
            AddHandler but.MouseLeave, AddressOf BtnAlt_MouseLeave
            AddHandler but.MouseUp, AddressOf Various_MouseUp
            pnlOverview.Controls.Add(but)
        Next i
        pnlOverview.ResumeLayout()
    End Sub

    Private Function UpdateButtonLayout2(count As Integer) As List(Of AButton)
        count += If(My.Settings.hideMessage, 0, 1)

        Dim targetAspectRatio As Double = 800 / 600

        ' Calculate the number of columns and rows based on the target aspect ratio and available space
        Dim cols As Integer = Math.Max(2, Math.Ceiling(Math.Sqrt(count * pnlOverview.Width / pnlOverview.Height / targetAspectRatio)))
        Dim rows As Integer = Math.Ceiling(count / cols)

        ' Adjust the number of columns or rows to ensure at least 'count' total items
        While cols * rows < count
            If pnlOverview.Width / (cols + 1) >= pnlOverview.Height / (rows + 1) Then
                cols += 1
            Else
                rows += 1
            End If
        End While

        ' Calculate the size of each button based on the adjusted layout
        Dim buttonWidth As Integer = pnlOverview.Width \ cols
        Dim buttonHeight As Integer = pnlOverview.Height \ rows

        ' Ensure that the calculated button size is not too small
        buttonWidth = Math.Max(1, buttonWidth)
        buttonHeight = Math.Max(1, buttonHeight)

        Dim totalButtons = cols * rows
        Dim i = If(My.Settings.hideMessage, 1, 2)

        Dim visButtons As New List(Of AButton)

        For Each but As AButton In pnlOverview.Controls.OfType(Of AButton).ToList
            If i <= totalButtons Then
                but.Size = New Size(buttonWidth, buttonHeight)
                but.Visible = True
                visButtons.Add(but)
            Else
                but.Visible = False
                but.Text = ""
                If but.AP IsNot Nothing Then
                    DwmUnregisterThumbnail(startThumbsDict.GetValueOrDefault(but.AP.Id, IntPtr.Zero))
                    startThumbsDict.TryRemove(but.AP.Id, Nothing)
                End If
                but.AP = Nothing
                but.pidCache = 0
            End If
            i += 1
        Next

        pnlMessage.Size = New Size(buttonWidth, buttonHeight)
        pbMessage.Size = pnlMessage.Size
        chkHideMessage.Location = New Point(pnlMessage.Width - chkHideMessage.Width, pnlMessage.Height - chkHideMessage.Height)

        Return visButtons
    End Function



    Private Function UpdateButtonLayout(count As Integer) As List(Of AButton)

        'Return UpdateButtonLayout2(count)

        'pnlOverview.SuspendLayout()
        Dim numCols As Integer

        Select Case count + If(My.Settings.hideMessage, 0, 1)
            Case 0 To 4
                numCols = 2
            Case 5 To 9
                numCols = 3
            Case 10 To 16
                numCols = 4
            Case 17 To 25
                numCols = 5
            Case Else
                numCols = 6
        End Select
        Dim numRows As Integer = numCols

        If Me.WindowState = FormWindowState.Maximized OrElse My.Settings.ApplyAlterNormal Then

            Dim numAlts = count + If(My.Settings.hideMessage, 0, 1)

            numCols = 2 + My.Settings.ExtraMaxColRow
            numRows = 2 - If(My.Settings.OneLessRowCol, 1, 0)

            While numAlts > numCols * numRows
                numCols += 1
                numRows += 1
                If numCols >= 7 Then Exit While
            End While

            If pbZoom.Width < pbZoom.Height Then
                Dim swapper As Integer = numCols
                numCols = numRows
                numRows = swapper
            End If
        End If

        Dim newSZ As New Size(pbZoom.Size.Width \ numCols, pbZoom.Size.Height \ numRows)
        Dim widthTooMuch As Boolean = False
        Dim heightTooMuch As Boolean = False

        If newSZ.Width * numCols > pbZoom.Width Then widthTooMuch = True
        If newSZ.Height * numRows > pbZoom.Height Then heightTooMuch = True

        Dim totalButtons = numCols * numRows

        Dim i = If(My.Settings.hideMessage, 1, 2)

        Dim visButtons As New List(Of AButton)

        For Each but As AButton In pnlOverview.Controls.OfType(Of AButton).ToList

            If i <= totalButtons Then
                but.Size = newSZ
                If widthTooMuch AndAlso i Mod numCols = 0 Then but.Width -= 1 'last column
                If heightTooMuch AndAlso i > (numRows - 1) * numRows Then but.Height -= 1 'last row
                but.Visible = True
                visButtons.Add(but)
            Else
                but.Visible = False
                but.Text = ""
                If but.AP IsNot Nothing Then
                    DwmUnregisterThumbnail(startThumbsDict.GetValueOrDefault(but.AP.Id, IntPtr.Zero))
                    startThumbsDict.TryRemove(but.AP.Id, Nothing)
                End If
                but.AP = Nothing
                but.pidCache = 0
            End If
            i += 1
        Next

        pnlMessage.Size = newSZ
        pbMessage.Size = newSZ
        chkHideMessage.Location = New Point(pnlMessage.Width - chkHideMessage.Width, pnlMessage.Height - chkHideMessage.Height)

        'pnlOverview.ResumeLayout(True)

        Return visButtons
    End Function
    Private Sub Buttons_BackColorChanged(sender As Button, e As EventArgs) Handles btnQuit.BackColorChanged, btnMax.BackColorChanged, btnMin.BackColorChanged
        Dim target As Control = sender
        While TypeOf target IsNot Form AndAlso target.BackColor.ToArgb = Color.Transparent.ToArgb
            target = target.Parent
        End While
        sender.FlatAppearance.BorderColor = target.BackColor
    End Sub
    Private Sub BtnQuit_MouseEnter(sender As Button, e As EventArgs) Handles btnQuit.MouseEnter
        cornerNE.BackColor = Color.Red
        sender.BackColor = Color.Red
    End Sub
    Private Sub BtnQuit_MouseLeave(sender As Button, e As EventArgs) Handles btnQuit.MouseLeave
        cornerNE.BackColor = Color.Transparent
        sender.BackColor = Color.Transparent
    End Sub
    Private Sub BtnQuit_MouseDown(sender As Button, e As MouseEventArgs) Handles btnQuit.MouseDown
        cornerNE.BackColor = COLOR_CLOSE_PRESSED
        sender.BackColor = COLOR_CLOSE_PRESSED
    End Sub

    Private Sub BtnQuit_Click(sender As Button, e As EventArgs) Handles btnQuit.Click
        Me.Close()
    End Sub
    Private Sub Various_MouseUp(sender As Control, e As MouseEventArgs) Handles Me.MouseUp, btnQuit.MouseUp, pnlSys.MouseUp, pnlButtons.MouseUp, btnStart.MouseUp, cboAlt.MouseUp, cmbResolution.MouseUp, ChkEqLock.MouseUp
        UntrapMouse(e.Button) 'fix mousebutton stuck
    End Sub

    Private Sub btnstart_MouseDown(sender As Control, e As MouseEventArgs) Handles btnStart.MouseDown
        CloseOtherDropDowns(cmsQuickLaunch.Items, Nothing)
        cmsQuickLaunch.Close()
    End Sub

    Private Sub cbx_DropDown(sender As ComboBox, e As EventArgs) Handles cboAlt.DropDown, cmbResolution.DropDown
        CloseOtherDropDowns(cmsQuickLaunch.Items, Nothing)
        cmsQuickLaunch.Close()
    End Sub

    Private Sub BtnMin_Click(sender As Button, e As EventArgs) Handles btnMin.Click
        dBug.Print("btnMin_Click")
        'suppressWM_MOVEcwp = True
        wasMaximized = (Me.WindowState = FormWindowState.Maximized)
        If Not wasMaximized Then
            restoreLoc = Me.Location
            dBug.Print("restoreLoc " & restoreLoc.ToString)
        End If
        AppActivate(scalaPID)

        If My.Settings.MinMin AndAlso pnlOverview.Visible AndAlso My.Settings.gameOnOverview Then
            Detach(True)
            MinAllActiveOverview()
        ElseIf My.Settings.MinMin AndAlso cboAlt.SelectedIndex <> 0 AndAlso AltPP?.isSDL Then
            AltPP.CenterBehind(pbZoom, SetWindowPosFlags.DoNotActivate)
            AltPP.Hide()
        Else
            dBug.Print("swl parent")
            AstoniaProcess.RestorePos(True)
            Detach(True)
        End If
        Me.WindowState = FormWindowState.Minimized
        dBug.Print($"WS {Me.WindowState}")
        'suppressWM_MOVEcwp = False
    End Sub

    Private Sub BtnHelp_Click(sender As Object, e As EventArgs) Handles btnHelp.Click
        Using helpForm As New frmHelp()
            helpForm.ShowDialog(Me)
        End Using
    End Sub

    Private Sub MinAllActiveOverview()
        For Each but As AButton In pnlOverview.Controls.OfType(Of AButton).Where(Function(b) b.AP IsNot Nothing)
            If Not but.AP.HasExited Then
                If but.AP.isSDL Then
                    but.AP.Hide()
                Else
                    but.AP.RestoreSinglePos()
                End If
            End If
        Next
    End Sub

    Private Sub BtnAlt_Click(sender As AButton, e As EventArgs) ' Handles AButton.click
        CloseOtherDropDowns(cmsQuickLaunch.Items, Nothing)
        cmsQuickLaunch.Close()
        If sender.Text = String.Empty Then
            'show cms
            cmsQuickLaunch.Show(sender, sender.PointToClient(MousePosition))
            Exit Sub
        End If
        If My.Settings.gameOnOverview Then
            AButton.ActiveOverview = True
            Exit Sub
        End If
        SelectAlt(sender.AP)
    End Sub
    Private Sub BtnAlt_MouseDown(sender As AButton, e As MouseEventArgs) ' handles AButton.mousedown
        dBug.Print($"MouseDown {e.Button}")
        CloseOtherDropDowns(cmsQuickLaunch.Items, Nothing)
        cmsQuickLaunch.Close()
        If sender.AP Is Nothing Then Exit Sub
        Select Case e.Button
            Case MouseButtons.XButton1, MouseButtons.XButton2
                sender.Select()
                sender.AP.Activate()
            Case MouseButtons.Left
                If e.Clicks = 2 Then
                    SelectAlt(sender.AP)
                End If
        End Select
    End Sub

    Private Sub SelectAlt(pp As AstoniaProcess)
        If SidebarScalA Is Nothing Then
            If Not cboAlt.Items.Contains(pp) Then
                PopDropDown(cboAlt)
            End If
            cboAlt.SelectedItem = pp
        Else
            IPC.AddToWhitelistOrRemoveFromBL(SidebarScalA.Id, pp.Id)
            IPC.SelectAlt(SidebarScalA.Id, pp.Id)

            'Dim sw As Stopwatch = Stopwatch.StartNew()
            'While IPC.ReadSelectAlt(scalaPID).Item1 <> 0 AndAlso sw.ElapsedMilliseconds <= 150
            '    Threading.Thread.Sleep(50)
            'End While
            'Me.BringToFront()
            Try
                'AppActivate(scalaPID)
                'AppActivate(pp.Id)
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub BtnAlt_MouseEnter(sender As AButton, e As EventArgs) ' Handles AButton.MouseEnter
        If My.Settings.gameOnOverview Then Exit Sub
        If sender.AP Is Nothing Then Exit Sub
        opaDict(sender.AP.Id) = 240
    End Sub
    Private Sub BtnAlt_MouseLeave(sender As AButton, e As EventArgs) ' Handles AButton.MouseLeave
        If sender.AP Is Nothing Then Exit Sub
        opaDict(sender.AP.Id) = If(chkDebug.Checked, 128, 255)
    End Sub

    Private Sub ChkHideMessage_CheckedChanged(sender As CheckBox, e As EventArgs) Handles chkHideMessage.CheckedChanged
        dBug.Print("chkHideMessage " & sender.Checked)
        If sender.Checked Then
            'btnAlt9.Visible = True
            pnlMessage.Visible = False
            My.Settings.hideMessage = True
        Else
            pnlMessage.Visible = True
            My.Settings.hideMessage = False
        End If
    End Sub

    'Private _restoreLoc As Point
    '''' <summary>
    '''' Used to set Scala to the right position when restoring from maximized state
    '''' </summary>
    'Private Property RestoreLoc As Point
    '    Get
    '        Return _restoreLoc
    '    End Get
    '    Set(ByVal value As Point)
    '        _restoreLoc = value
    '        dBug.print($"Set restoreloc to {value}")
    '    End Set
    'End Property

    Dim prevWA As Rectangle
    Dim restoreLoc As Point = Me.Location
    Private Sub BtnMax_Click(sender As Button, e As EventArgs) Handles btnMax.Click
        dBug.Print("btnMax_Click")
        suppressWM_MOVEcwp = True
        '🗖,🗗,⧠
        If Me.WindowState = FormWindowState.Normal Then 'go maximized
            My.Settings.zoom = cmbResolution.SelectedIndex

            If Me.Location <> New Point(-32000, -32000) Then '-32K is minimized window location
                restoreLoc = Me.Location
            End If

#Region "border" 'to enable dragging maximized caption
            Dim scrn As Screen = Screen.FromPoint(restoreLoc + New Point(Me.Width / 2, Me.Height / 2))
            dBug.Print("screen workarea " & scrn.WorkingArea.ToString)
            dBug.Print("screen bounds " & scrn.Bounds.ToString)
            prevWA = scrn.WorkingArea

            Dim leftBorder As Integer = scrn.WorkingArea.Width * My.Settings.MaxBorderLeft / 1000
            Dim topBorder As Integer = scrn.WorkingArea.Height * My.Settings.MaxBorderTop / 1000
            Dim rightborder As Integer = scrn.WorkingArea.Width * My.Settings.MaxBorderRight / 1000
            Dim botBorder As Integer = scrn.WorkingArea.Height * My.Settings.MaxBorderBot / 1000

            'find out where taskbar is and add 1 pixel at that location
            'dirty hack to enable dragging when maximized, TODO: split up frmmain into 2, caption and zoom
            If leftBorder + rightborder + topBorder + botBorder = 0 Then
                If scrn.WorkingArea.Left <> scrn.Bounds.Left Then leftBorder = 1
                If scrn.WorkingArea.Top <> scrn.Bounds.Top Then topBorder = 1
                If scrn.WorkingArea.Right <> scrn.Bounds.Right Then rightborder = 1
                If scrn.WorkingArea.Bottom <> scrn.Bounds.Bottom Then botBorder = 1
            End If
            'if taskbar set to auto hide find where it is hiding
            If leftBorder + rightborder + topBorder + botBorder = 0 Then
                'Dim monifo As New MonitorInfo With {.cbSize = Runtime.InteropServices.Marshal.SizeOf(GetType(MonitorInfo))}
                'GetMonitorInfo(MonitorFromPoint(scrn.Bounds.Location, MONITOR.DEFAULTTONEAREST), monifo)
                Dim pabd As New APPBARDATA With {
                        .cbSize = Runtime.InteropServices.Marshal.SizeOf(GetType(APPBARDATA)),
                        .rc = scrn.Bounds.ToRECT}
                For edge = 0 To 3
                    pabd.uEdge = edge
                    If SHAppBarMessage(ABM.GETAUTOHIDEBAREX, pabd) <> IntPtr.Zero Then
                        Select Case edge
                            Case 0
                                leftBorder = 1
                                dBug.Print("Hidden taskbar left")
                            Case 1
                                topBorder = 1
                                dBug.Print("Hidden taskbar top")
                            Case 2
                                rightborder = 1
                                dBug.Print("Hidden taskbar right")
                            Case 3
                                botBorder = 1
                                dBug.Print("Hidden taskbar bottom")
                        End Select
                        Exit For
                    End If
                Next
            End If
            'if no taskbar present add to bottom
            If leftBorder + rightborder + topBorder + botBorder = 0 Then
                dBug.Print("no taskbar present")
                botBorder = 1
            End If

            dBug.Print($"leftborder {leftBorder}")
            dBug.Print($"topborder {topBorder}")
            dBug.Print($"rightborder {rightborder}")
            dBug.Print($"botborder {botBorder}")

            Me.MaximizedBounds = New Rectangle(scrn.WorkingArea.Left - scrn.Bounds.Left + leftBorder,
                                           scrn.WorkingArea.Top - scrn.Bounds.Top + topBorder,
                                           scrn.WorkingArea.Width - leftBorder - rightborder,
                                           scrn.WorkingArea.Height - topBorder - botBorder)
#End Region
            dBug.Print("new maxbound " & MaximizedBounds.ToString)
            If Me.WindowState = FormWindowState.Normal AndAlso Me.Location <> New Point(-32000, -32000) Then
                restoreLoc = Me.Location 'todo: why is this duplicated?
                dBug.Print("restoreLoc " & restoreLoc.ToString)
            End If
            'ReZoom()
            If Me.Location = New Point(-32000, -32000) Then Me.Location = restoreLoc
            Me.WindowState = FormWindowState.Maximized
            sender.Text = "🗗"
            Me.Invalidate()
            ttMain.SetToolTip(sender, "Restore")
            wasMaximized = True
            FrmSizeBorder.Opacity = 0
        ElseIf Me.WindowState = FormWindowState.Maximized Then 'go normal
            dBug.Print($"restorebounds {RestoreBounds.Location}")
            dBug.Print($"maximizbounds {MaximizedBounds.Location}")
            dBug.Print($"restoreloc    {restoreLoc}")
            dBug.Print($"My.Settings.l {My.Settings.location}")
            Me.Location = restoreLoc
            Me.WindowState = FormWindowState.Normal
            sender.Text = "⧠"
            ttMain.SetToolTip(sender, "Maximize")
            'wasMaximized = False
            ReZoom(My.Settings.resol)
            'wasMaximized = True
            'Me.Location = restoreLoc
            AOshowEqLock = False
            FrmSizeBorder.Opacity = If(chkDebug.Checked, 1, 0.01)
            FrmSizeBorder.Opacity = If(My.Settings.SizingBorder, FrmSizeBorder.Opacity, 0)
        End If
        If cboAlt.SelectedIndex > 0 Then
            Attach(AltPP, True)
            AltPP?.CenterBehind(pbZoom)
        End If
        moveBusy = False
        suppressWM_MOVEcwp = False
        FrmSizeBorder.Bounds = Me.Bounds
    End Sub

    Private Sub BtnStart_Click(sender As Button, e As EventArgs) Handles btnStart.Click
        cboAlt.DroppedDown = False
        tmrTick.Stop()
        Dim prevAlt As AstoniaProcess = AltPP
        dBug.Print($"prevAlt?.Name {prevAlt?.Name}")
        AstoniaProcess.RestorePos(True)
        cboAlt.SelectedIndex = 0
        If prevAlt?.Id <> 0 Then
            pnlOverview.Controls.OfType(Of AButton).FirstOrDefault(Function(ab As AButton) ab.AP IsNot Nothing AndAlso ab.AP.Id = prevAlt.Id)?.Select()
        Else
            pnlOverview.Controls.OfType(Of AButton).First().Select()
        End If
    End Sub

    Public Shared topSortList As List(Of String) = My.Settings.topSort.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries).ToList
    Public Shared botSortList As List(Of String) = My.Settings.botSort.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries).ToList
    Public Shared blackList As List(Of String) = topSortList.Intersect(botSortList).ToList
    'Private EQLockClick As Boolean = False

    Public ReadOnly restoreParent As UInteger = GetWindowLong(Me.Handle, GWL_HWNDPARENT)
    Private prevHWNDParent As IntPtr = restoreParent
    Public Function Attach(ap As AstoniaProcess, Optional activate As Boolean = False) As Long
        Try
            If ap Is Nothing OrElse prevHWNDParent = ap.MainWindowHandle Then Return 0
            dBug.Print($"Attach to: {ap.Name} {ap.Id} activate {activate}")
            prevHWNDParent = ap.MainWindowHandle
            If Not pnlOverview.Visible Then
                Dim rcAst As RECT
                GetWindowRect(ap.MainWindowHandle, rcAst)
                If rcAst.top <= Me.Location.Y Then
                    SetWindowPos(ap.MainWindowHandle, ScalaHandle, Me.Location.X, Me.Location.Y + 2, -1, -1, SetWindowPosFlags.IgnoreResize Or SetWindowPosFlags.DoNotActivate)
                End If
            End If
            Dim ret = SetWindowLong(ScalaHandle, GWL_HWNDPARENT, ap.MainWindowHandle)
            'Dim ScalAThreadId As Integer = GetWindowThreadProcessId(ScalaHandle, Nothing) 'move this to a global since this won't change
            'Dim AstoniaThreadId As Integer = GetWindowThreadProcessId(ap.MainWindowHandle, Nothing) 'move this to astoniaproc
            'ap.ThreadInput(False) 
            AttachThreadInput(ScalaThreadId, ap.MainThreadId, False) 'detach input so ctrl, shift and alt still work when there is an elevation mismatch, also fixes sleepy legacy clients lagging ScalA

            Return ret
        Finally
            If activate Then Task.Run(Sub()
                                          Threading.Thread.Sleep(25) 'needed or we get an activation bug. i cant find exactly where it is tho. maybe it's the attachthread resetting input queues?
                                          AltPP?.Activate()
                                      End Sub)
        End Try
    End Function
#If DEBUG Then
    Private prevDetach As String
#End If
    Public Function Detach(show As Boolean) As Long
        ' Disable zoom state for SDL2 wrapper
        ZoomStateIPC.SetZoomStateEnabled(False)
#If DEBUG Then
        If prevDetach <> AltPP?.UserName Then
            dBug.Print($"Detach from: {AltPP?.UserName} show:{show}")
            prevDetach = AltPP?.UserName
        End If
#End If
        Try
            Return SetWindowLong(ScalaHandle, GWL_HWNDPARENT, restoreParent)
        Finally
            prevHWNDParent = restoreParent
            If show Then
                AllowSetForegroundWindow(ASFW_ANY)
                SetForegroundWindow(GetDesktopWindow)
                SetForegroundWindow(ScalaHandle)

                'SwitchToThisWindow(GetDesktopWindow(), True)
                'SwitchToThisWindow(ScalaHandle, True)
                'Task.Run(Sub()
                '             Threading.Thread.Sleep(100)
                '             AllowSetForegroundWindow(scalaPID)
                '             Invoke(Sub() Activate())
                '             Threading.Thread.Sleep(100)
                '             FlashWindow(ScalaHandle, True) 'show on taskbar
                '             Try
                '                 AppActivate(scalaPID)
                '             Catch
                '             End Try
                '             FlashWindow(ScalaHandle, False) 'stop blink
                '         End Sub)
            End If
        End Try
    End Function


    Private Sub setActive(active As Boolean)
        Dim fcol As Color = Color.FromArgb(&HFF666666)
        If active Then fcol = If(My.Settings.DarkMode, Colors.LightText, SystemColors.ControlText)
        lblTitle.ForeColor = fcol
        btnMax.ForeColor = fcol
        btnMin.ForeColor = fcol
        btnStart.ForeColor = fcol
        cboAlt.ForeColor = fcol
        cmbResolution.ForeColor = fcol
        For Each but As Button In pnlButtons.Controls
            If but.Contains(MousePosition) Then
                If but Is btnQuit Then
                    but.ForeColor = Color.White
                Else
                    but.ForeColor = If(My.Settings.DarkMode, Color.White, SystemColors.ControlText)
                End If
            Else
                but.ForeColor = fcol
            End If
        Next
        cboAlt.ForeColor = fcol
        cmbResolution.ForeColor = fcol
    End Sub

    Private Sub StartCloseErrorDialogThread()
        Dim t As New Threading.Thread(AddressOf CloseErrorDialogLoop)
        t.IsBackground = True
        t.Priority = Threading.ThreadPriority.Lowest
        t.Start()
    End Sub
    Private Sub CloseErrorDialogLoop()
        While True
            Threading.Thread.Sleep(420)
            CloseErrorDialog()
        End While
    End Sub
    Private Sub CloseErrorDialog()
        Try
            Dim errorHwnd = FindWindow("#32770", "error")
            If errorHwnd = IntPtr.Zero Then errorHwnd = FindWindow("#32770", "Application Crashed")
            If errorHwnd <> IntPtr.Zero Then
                If FindWindowEx(errorHwnd, Nothing, "Static", "Copy new->moac failed: 32") <> IntPtr.Zero OrElse
                   FindWindowEx(errorHwnd, Nothing, "Static", "Copy new->moac failed: 5") <> IntPtr.Zero OrElse
                   FindWindowEx(errorHwnd, Nothing, "Static", "Details written to moac.log") <> IntPtr.Zero Then
                    Dim butHandle = FindWindowEx(errorHwnd, Nothing, "Button", "OK")
                    If butHandle <> IntPtr.Zero Then
                        SendMessage(butHandle, BM_CLICK, IntPtr.Zero, IntPtr.Zero)
                        dBug.Print("Error dialog closed")
                    End If
                End If
            End If
        Catch
            dBug.Print("CloseErrorDialog Exception")
        End Try
    End Sub

    Private Sub Title_MouseDoubleClick(sender As Control, e As MouseEventArgs) Handles pnlTitleBar.MouseDoubleClick, lblTitle.MouseDoubleClick
        dBug.Print("title_DoubleClick")
        If e.Button = MouseButtons.Left Then btnMax.PerformClick()
        'FrmSizeBorder.Bounds = Me.Bounds
    End Sub




    Public Sub RestartSelf(Optional asAdmin As Boolean = True)

        Dim procStartInfo As New ProcessStartInfo With {
            .UseShellExecute = True,
            .FileName = Environment.GetCommandLineArgs()(0),
            .Arguments = """" & CType(Me.cboAlt.SelectedItem, AstoniaProcess)?.UserName & """",
            .WindowStyle = ProcessWindowStyle.Normal,
            .Verb = If(asAdmin, "runas", "") 'add this to prompt for elevation
        }

        SaveLocation()
        My.Settings.Save()
        Try
            Process.Start(procStartInfo).WaitForInputIdle()
        Catch e As System.ComponentModel.Win32Exception
            'operation cancelled
        Catch e As InvalidOperationException
            'wait for inputidle is needed
        Catch e As Exception
            Throw e
        End Try
        sysTrayIcon.Visible = False
        sysTrayIcon.Dispose()

    End Sub
    Public Sub UnelevateSelf()
        SaveLocation()
        My.Settings.Save()

        tmrActive.Stop()
        tmrOverview.Stop()
        tmrTick.Stop()

        AstoniaProcess.RestorePos()

        ExecuteProcessUnElevated(Environment.GetCommandLineArgs()(0), "Someone", IO.Directory.GetCurrentDirectory())
        sysTrayIcon.Visible = False
        sysTrayIcon.Dispose()
        End 'program
    End Sub


#If DEBUG Then
    Private Sub ChkDebug_CheckedChanged(sender As CheckBox, e As EventArgs) Handles chkDebug.CheckedChanged
        dBug.Print(Screen.GetWorkingArea(sender).ToString)
        If Not pnlOverview.Visible Then
            UpdateThumb(If(sender.Checked, 122, 255))
        Else
            For Each but As AButton In pnlOverview.Controls.OfType(Of AButton)
                but.Invalidate()
            Next
        End If
        If WindowState <> FormWindowState.Maximized Then
            FrmSizeBorder.Opacity = If(sender.Checked, 1, 0.01)
        End If
        FrmBehind.BackColor = If(sender.Checked, Color.Cyan, Color.Black)
        FrmBehind.Opacity = If(sender.Checked, 1, 0.01)
        FrmSizeBorder.Opacity = If(My.Settings.SizingBorder, FrmSizeBorder.Opacity, 0)
    End Sub
#End If


    Private Sub SysTrayIcon_MouseDoubleClick(sender As NotifyIcon, e As MouseEventArgs) Handles sysTrayIcon.MouseDoubleClick
        dBug.Print("sysTrayIcon_MouseDoubleClick")
        If e.Button = MouseButtons.Right Then Exit Sub
        'If Me.WindowState = FormWindowState.Minimized Then
        '    Me.Location = RestoreLoc
        '    SetWindowLong(ScalaHandle, GWL_HWNDPARENT, AltPP.MainWindowHandle) 'hides scala from taskbar
        '    suppressWM_MOVEcwp = True
        '    Me.WindowState = If(wasMaximized, FormWindowState.Maximized, FormWindowState.Normal)
        '    suppressWM_MOVEcwp = False
        '    ReZoom(zooms(cmbResolution.SelectedIndex)) 'handled in WM_SIZE
        '    AltPP?.CenterBehind(pbZoom)
        '    btnMax.Text = If(wasMaximized, "🗗", "⧠")
        '    ttMain.SetToolTip(btnMax, If(wasMaximized, "Restore", "Maximize"))
        '    If wasMaximized Then btnMax.PerformClick()
        'End If
        If AltPP?.IsMinimized Then AltPP?.Restore()
        If Me.WindowState = FormWindowState.Minimized Then
            Me.WndProc(Message.Create(ScalaHandle, WM_SYSCOMMAND, SC_RESTORE, Nothing))
        End If
        Me.Show()
        'Me.BringToFront() 'doesn't work
        If AltPP IsNot Nothing AndAlso AltPP.Id <> 0 AndAlso AltPP.IsRunning Then
            SetWindowPos(AltPP.MainWindowHandle, SWP_HWND.NOTOPMOST, -1, -1, -1, -1, SetWindowPosFlags.IgnoreMove Or SetWindowPosFlags.IgnoreResize Or SetWindowPosFlags.DoNotActivate)
            Attach(AltPP, True)
        Else
            Me.TopMost = True
            Me.TopMost = My.Settings.topmost
        End If
        'moveBusy = False
    End Sub

    Private Async Sub FrmMain_Click(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        dBug.Print($"me.mousedown {sender.name} {e.Button}")
        MyBase.WndProc(Message.Create(ScalaHandle, WM_CANCELMODE, 0, 0))
        CloseOtherDropDowns(cmsQuickLaunch.Items, Nothing)
        cmsQuickLaunch.Close()
        Await Task.Delay(200)
        dBug.Print($"Me.MouseDown awaited")
        If Not pnlOverview.Visible Then
            pbZoom.Visible = True
            If e.Button <> MouseButtons.None Then
                Attach(AltPP, True)
            End If
        Else
            AButton.ActiveOverview = My.Settings.gameOnOverview
        End If
    End Sub

    'Private Function WM_MM_GetWParam() As Integer
    '    Dim wp As Integer
    '    wp = wp Or If(MouseButtons.HasFlag(MouseButtons.Left), MK_LBUTTON, 0)
    '    wp = wp Or If(MouseButtons.HasFlag(MouseButtons.Right), MK_RBUTTON, 0)
    '    wp = wp Or If(MouseButtons.HasFlag(MouseButtons.Middle), MK_MBUTTON, 0)
    '    wp = wp Or If(MouseButtons.HasFlag(MouseButtons.XButton1), MK_XBUTTON1, 0)
    '    wp = wp Or If(MouseButtons.HasFlag(MouseButtons.XButton2), MK_XBUTTON2, 0)
    '    wp = wp Or If(My.Computer.Keyboard.CtrlKeyDown, MK_CONTROL, 0)
    '    wp = wp Or If(My.Computer.Keyboard.ShiftKeyDown, MK_SHIFT, 0)
    '    Return wp
    'End Function

    'Public Function WM_MOUSEMOVE_CreateWParam() As IntPtr
    '    Dim wp As Integer
    '    wp = wp Or ((MouseButtons >> 20) And &H3)   ' 00000011 (Extract left and right buttons)
    '    wp = wp Or ((ModifierKeys >> 14) And &H300) ' 00001100 (Extract Shift and Ctrl keys)
    '    wp = wp Or ((MouseButtons >> 18) And &H70)  ' 01110000 (Extract middle and X buttons)
    '    Return New IntPtr(wp)
    'End Function

    Private Async Sub JiggerMouse()
        prevWMMMpt = New Point
        Cursor.Position += New Point(-1, -1)
        Await Task.Delay(32)
        prevWMMMpt = New Point
        Cursor.Position += New Point(1, 1)
    End Sub
    Private Sub PnlEqLock_MouseDown(sender As Panel, e As MouseEventArgs) Handles PnlEqLock.MouseDown
        dBug.Print($"pnlEqLock.MouseDown {e.Button}")

        Dim wparam = WM_MOUSEMOVE_CreateWParam()

        Dim rc As Rectangle
        GetClientRect(AltPP.MainWindowHandle, rc)

        Dim mp As Point = sender.PointToClient(MousePosition)

        Dim sx As Integer = rcC.Width / 2 - 262.Map(0, 800, 0, rcC.Width)

        Dim excludGearLock As Integer = If(AltPP?.isSDL, 18, 0)
        Dim dx As Integer = (524 - excludGearLock).Map(0, 800, 0, rcC.Width)

        Dim mx As Integer = mp.X.Map(0, sender.Bounds.Width, sx, dx)

        Dim lockHeight = 45
        If rc.Height >= 2000 Then
            lockHeight += 120
        ElseIf rc.Height >= 1500 Then
            lockHeight += 80
        ElseIf rc.Height >= 1000 Then
            lockHeight += 40
        End If
        Dim my As Integer = mp.Y.Map(0, sender.Bounds.Height, 0, lockHeight)

        dBug.Print($"mx:{mx} my:{my}")



        'Dim isSdlWithSpallIcons
        'todo:: block right clicks when SDL and spellIcons


        If e.Button = MouseButtons.Middle OrElse e.Button = MouseButtons.Right Then
            PnlEqLock.Visible = False
            If Not AltPP?.isSDL Then sender.Capture = False
            Application.DoEvents()
            'Attach(AltPP, True)

            AltPP?.ThreadInput(True)

            'SendMouseInput(If(e.Button = MouseButtons.Right, MouseEventF.RightDown, MouseEventF.MiddleDown))

            SendMessage(AltPP?.MainWindowHandle, If(e.Button = MouseButtons.Right, WM_RBUTTONDOWN, WM_MBUTTONDOWN), wparam, New LParamMap(mx, my))

            ' Track input for inactivity timeout warning
            AltPP?.RecordInput()
        End If


    End Sub
    Private Sub PnlEqLock_MouseUp(sender As Panel, e As MouseEventArgs) Handles PnlEqLock.MouseUp
        dBug.Print($"pnlEqLock.MouseUp {e.Button} lock vis {PnlEqLock.Visible}")
        If (e.Button = MouseButtons.Right OrElse e.Button = MouseButtons.Middle) AndAlso PnlEqLock.Contains(MousePosition) Then
            '    'EQLockClick = True
            PnlEqLock.Visible = False
            'sender.Capture = False
            If e.Button = MouseButtons.Right Then
                SendMouseInput(MouseEventF.RightUp)
            Else
                '        SendMouseInput(MouseEventF.MiddleDown)
                '        Await Task.Delay(50)
                SendMouseInput(MouseEventF.MiddleUp)
            End If
            '    Await Task.Delay(25)
            '    'EQLockClick = False
        End If
        'Await Task.Run(Sub() Attach(AltPP, True))
    End Sub

    Private Sub PnlEqLock_Mousemove(sender As Panel, e As MouseEventArgs) Handles PnlEqLock.MouseMove
        If e.Button = MouseButtons.Right OrElse e.Button = MouseButtons.Middle Then
            sender.Visible = False
        End If
    End Sub

    Private Sub ChkEqLock_CheckedChanged(sender As CheckBox, e As EventArgs) Handles ChkEqLock.CheckedChanged
        'locked 🔒
        'unlocked 🔓
        sender.Text = If(sender.CheckState = CheckState.Unchecked, "🔓", "🔒")

    End Sub
    Private prevMPX As Integer
    Private Sub Caption_MouseMove(sender As Object, e As MouseEventArgs) Handles btnStart.MouseMove, cboAlt.MouseMove, cmbResolution.MouseMove,
                                                                                 pnlTitleBar.MouseMove, lblTitle.MouseMove, pbDpiWarning.MouseMove, pbWrapperWarning.MouseMove,
                                                                                 pnlUpdate.MouseMove, pbUpdateAvailable.MouseMove, ChkEqLock.MouseMove,
                                                                                 pnlSys.MouseMove, btnMin.MouseMove, btnMax.MouseMove, btnQuit.MouseMove
        If cboAlt.SelectedIndex = 0 Then Exit Sub
        If prevMPX = MousePosition.X Then Exit Sub
        prevMPX = MousePosition.X
        If Not Me.RectangleToScreen(New Rectangle(0, 0, Me.Width, pnlTitleBar.Height)).Contains(MousePosition) Then Exit Sub

        'TODO: move follwing code to tmrTick and test sizeborder drag
        Dim ptZ As Point = Me.PointToScreen(pbZoom.Location)

        'dBug.Print("CaptionMouseMove")
        If AltPP Is Nothing Then Exit Sub
        newX = MousePosition.X.Map(ptZ.X, ptZ.X + pbZoom.Width, ptZ.X, ptZ.X + pbZoom.Width - rcC.Width) - If(AltPP?.ClientOffset.X, 0)
        newY = Me.Location.Y

        Dim flags = swpFlags
        If Not AltPP?.IsActive() Then flags.SetFlag(SetWindowPosFlags.DoNotChangeOwnerZOrder)
        If AltPP?.IsBelow(ScalaHandle) Then flags.SetFlag(SetWindowPosFlags.IgnoreZOrder)
        Task.Run(Sub()
                     SetWindowPos(If(AltPP?.MainWindowHandle, IntPtr.Zero), ScalaHandle, newX, newY, -1, -1, flags)
                 End Sub)

    End Sub

End Class