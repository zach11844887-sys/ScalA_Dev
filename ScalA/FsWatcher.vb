Imports System.Collections.Concurrent
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading
Imports ScalA.QL
Module FsWatcher

    Private usingLatestReg As Boolean? = Nothing
    Public Function GetUserChoiceKeyPath(protocol As String) As String
        Dim suff As String = ""
        Dim base = $"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\{protocol}\"
        If usingLatestReg IsNot Nothing Then
            If usingLatestReg Then
                suff = "UserChoiceLatest\ProgId"
            Else
                suff = "UserChoice"
            End If
        End If
        If String.IsNullOrEmpty(suff) Then
            Using latestKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(base & "UserChoiceLatest")
                If latestKey IsNot Nothing Then
                    suff = "UserChoiceLatest\ProgId"
                    usingLatestReg = True
                Else
                    suff = "UserChoice"
                    usingLatestReg = False
                End If
            End Using
        End If
        Return base & suff
    End Function

    Public Class RegistryWatcher
        Implements IDisposable

        Shared ReadOnly regWatchers As New ConcurrentDictionary(Of String, RegistryWatcher)
        Public Shared ReadOnly regWatcherLock As New Object()

        Public Shared Sub Init(proto As String)
            Dim key As String = proto.ToLowerInvariant()
            If regWatchers.ContainsKey(key) Then Exit Sub

            SyncLock regWatcherLock
                If Not regWatchers.ContainsKey(key) Then
                    Dim watcher As New RegistryWatcher(key)
                    AddHandler watcher.RegistryChanged, AddressOf OnRegistryChanged
                    watcher.Start()
                    regWatchers.TryAdd(key, watcher)
                    dBug.Print($"Started watcher for protocol: {key}")
                End If
            End SyncLock
        End Sub

        Public Shared Sub OnRegistryChanged(proto As String)
            dBug.Print($"Registry changed for protocol: {proto}")
            QLIconCache.DefURLicons.TryRemove(proto, Nothing)
            If QLIconCache.DefURLicons.Keys.Contains(proto) Then Throw New Exception

            For Each key In QLIconCache.IconCache.Keys.Where(Function(k) k.EndsWith(".url"))
                QLIconCache.IconCache.TryRemove(key, Nothing)
            Next
            If QLIconCache.IconCache.Keys.Any(Function(k) k.ToLower.EndsWith(".url")) Then Throw New Exception

        End Sub

        Private ReadOnly keyPath As String
        Private ReadOnly protocol As String
        Private watcherThread As Thread
        Private running As Boolean

        Public Event RegistryChanged(protocol As String)

        Public Sub New(protocol As String)
            Me.protocol = protocol.ToLowerInvariant()
            keyPath = GetUserChoiceKeyPath(Me.protocol)
        End Sub

        Public Sub Start()
            If watcherThread IsNot Nothing Then Return
            running = True
            watcherThread = New Thread(AddressOf WatchLoop) With {.IsBackground = True, .Priority = ThreadPriority.Lowest}
            watcherThread.Start()
        End Sub

        Private Sub WatchLoop()
            Using key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(keyPath, False)
                If key Is Nothing Then Return
                Dim handle = key.Handle.DangerousGetHandle()
                While running
                    RegNotifyChangeKeyValue(handle, False, RegChangeNotifyFilter.Value, IntPtr.Zero, False)
                    RaiseEvent RegistryChanged(protocol)
                End While
            End Using
        End Sub

        Private disposedValue As Boolean = False

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' No need to stop or join the thread
                    ' If you had other managed resources, dispose them here
                End If

                disposedValue = True
            End If
        End Sub

        Protected Overrides Sub Finalize()
            Dispose(False)
            MyBase.Finalize()
        End Sub

        <Flags>
        Private Enum RegChangeNotifyFilter As UInteger
            Name = 1
            Attributes = 2
            Value = 4
            Security = 8
        End Enum

        <DllImport("advapi32.dll", SetLastError:=True)>
        Private Shared Function RegNotifyChangeKeyValue(hKey As IntPtr, bWatchSubtree As Boolean, dwNotifyFilter As RegChangeNotifyFilter, hEvent As IntPtr, fAsynchronous As Boolean) As Integer : End Function
    End Class

    Public ReadOnly fsWatchers As New List(Of System.IO.FileSystemWatcher)

    Public ReadOnly ResolvedLinkWatchers As New Concurrent.ConcurrentDictionary(Of String, List(Of IO.FileSystemWatcher))
    Public ReadOnly ResolvedLinkLinks As New Concurrent.ConcurrentDictionary(Of String, String)

    Public Sub addLinkWatcher(pth As String, lnk As String)
        Task.Run(Sub()
                     If Not pth.Contains(My.Settings.links) AndAlso Not ResolvedLinkWatchers.Keys.Any(Function(wl As String) pth.Contains(wl)) Then
                         dBug.Print($"addLinkWatcher {pth}")
                         Dim ws As New List(Of FileSystemWatcher)
                         InitWatchers(pth, ws)
                         If ResolvedLinkWatchers.TryAdd(pth, ws) Then
                             ResolvedLinkLinks.TryAdd(lnk, pth)
                             StartDirectoryPoller()
                         End If
                     End If
                 End Sub)
    End Sub
    Public Sub ResolvedLinkwatchers_Clear()
        For Each ws As FileSystemWatcher In ResolvedLinkWatchers.Values.SelectMany(Function(fws As List(Of IO.FileSystemWatcher)) fws)
            ws.EnableRaisingEvents = False
            ws.Dispose()
        Next
        ResolvedLinkLinks.Clear()
        ResolvedLinkWatchers.Clear()
    End Sub
    Public Sub removeLinkWatcher(pth As String)
        dBug.Print($"removeLinkWatcher ""{pth}""")
        Dim removed As List(Of IO.FileSystemWatcher) = Nothing
        If ResolvedLinkWatchers.TryRemove(pth, removed) Then

            For Each kvp In ResolvedLinkLinks.ToList
                If kvp.Value = pth Then
                    dBug.Print($"ResolvedLinkLinks.TryRemove({kvp.Key})", 1)
                    ResolvedLinkLinks.TryRemove(kvp.Key, Nothing)
                End If
            Next

            For Each watcher In removed
                Try
                    watcher.EnableRaisingEvents = False
                    watcher.Dispose()
                Catch ex As Exception
                    dBug.Print($"Error disposing watcher: {ex.Message}")
                End Try
            Next
        End If
    End Sub
    Public Property PollerThread As Thread
    Public Property PollerRunning As Boolean = False
    Private ReadOnly PollerLock As New Object()
    Public Sub StartDirectoryPoller()
        SyncLock PollerLock
            If PollerRunning Then Exit Sub

            PollerRunning = True
            PollerThread = New Thread(AddressOf DirectoryPollerLoop) With {.IsBackground = True, .Priority = ThreadPriority.Lowest}
            PollerThread.Start()
            dBug.Print("Directory poller started.")
        End SyncLock
    End Sub

    Public Sub StopDirectoryPoller()
        PollerRunning = False
        If PollerThread IsNot Nothing AndAlso PollerThread.IsAlive Then
            PollerThread.Join(500)
        End If
        PollerThread = Nothing
        dBug.Print("Directory poller stopped (external call).")
    End Sub

    Private Sub DirectoryPollerLoop()
        Do While PollerRunning
            Try
                Dim allPaths = ResolvedLinkWatchers.Keys.ToList()
                For Each Pth As String In allPaths
                    ' Remove missing dirs
                    If Not IO.Directory.Exists(Pth) Then
                        dBug.Print($"Directory missing during poll: {Pth}")
                        removeLinkWatcher(Pth)
                        Continue For
                    End If
                    ' Check for redundancy
                    For Each other In allPaths
                        If other.Equals(Pth, StringComparison.OrdinalIgnoreCase) Then Continue For 'skip self

                        If Pth.StartsWith(other.TrimEnd("\"c) & "\", StringComparison.OrdinalIgnoreCase) Then
                            dBug.Print($"Redundant path detected during poll: {Pth} (already covered by {other})")
                            removeLinkWatcher(Pth)
                            Exit For ' Stop checking after removing
                        End If
                    Next
                Next
            Catch ex As Exception
                dBug.Print($"Poller error: {ex.Message}")
            End Try

            If ResolvedLinkWatchers.Count = 0 Then
                dBug.Print("Directory poller exiting: no paths to monitor.")
                Exit Do
            End If

            Thread.Sleep(3210)
        Loop

        PollerRunning = False
        PollerThread = Nothing
        dBug.Print("Directory poller thread has stopped.")
    End Sub
    Public Sub UpdateWatchers(newPath As String)
        dBug.Print("updateWatchers")
        For Each w As System.IO.FileSystemWatcher In fsWatchers
            w.EnableRaisingEvents = False
            w.Path = newPath
            w.EnableRaisingEvents = True
        Next
        ResolvedLinkwatchers_Clear()
    End Sub
    Public Sub InitWatchers(path As String, ByRef watchers As List(Of IO.FileSystemWatcher))
        dBug.Print($"initWatchers {path}")
        For Each w As System.IO.FileSystemWatcher In watchers
            w.EnableRaisingEvents = False
            w.Dispose()
        Next
        watchers.Clear()
        Try

            Dim iniWatcher As New System.IO.FileSystemWatcher(path) With {
                .NotifyFilter = System.IO.NotifyFilters.LastWrite,
                .Filter = "desktop.ini",
                .IncludeSubdirectories = True
            }

            AddHandler iniWatcher.Changed, AddressOf OnChanged
            AddHandler iniWatcher.Error, AddressOf OnWatcherError
            iniWatcher.EnableRaisingEvents = True

            watchers.Add(iniWatcher)



            Dim lnkWatcher As New System.IO.FileSystemWatcher(path) With {
            .NotifyFilter = System.IO.NotifyFilters.LastWrite Or
                            System.IO.NotifyFilters.FileName,
            .Filter = "*.lnk",
            .IncludeSubdirectories = True
        }

            AddHandler lnkWatcher.Renamed, AddressOf OnRenamed
            AddHandler lnkWatcher.Changed, AddressOf OnChanged
            AddHandler lnkWatcher.Deleted, AddressOf onDeleteFile
            AddHandler lnkWatcher.Error, AddressOf OnWatcherError
            lnkWatcher.EnableRaisingEvents = True

            watchers.Add(lnkWatcher)



            Dim urlWatcher As New System.IO.FileSystemWatcher(path) With {
            .NotifyFilter = System.IO.NotifyFilters.LastWrite Or
                            System.IO.NotifyFilters.FileName,
            .Filter = "*.url",
            .IncludeSubdirectories = True
        }

            AddHandler urlWatcher.Renamed, AddressOf OnRenamed
            AddHandler urlWatcher.Changed, AddressOf OnChanged
            AddHandler urlWatcher.Deleted, AddressOf onDeleteFile
            AddHandler urlWatcher.Error, AddressOf OnWatcherError
            urlWatcher.EnableRaisingEvents = True

            watchers.Add(urlWatcher)



            Dim dirWatcher As New System.IO.FileSystemWatcher(path) With {
            .NotifyFilter = System.IO.NotifyFilters.DirectoryName,
            .IncludeSubdirectories = True
        }

            AddHandler dirWatcher.Renamed, AddressOf OnRenamedDir
            AddHandler dirWatcher.Deleted, AddressOf OnDeleteDir
            AddHandler dirWatcher.Error, AddressOf OnWatcherError
            dirWatcher.EnableRaisingEvents = True

            watchers.Add(dirWatcher)

        Catch ex As Exception
            dBug.Print($"iniwatch {ex.Message}")
        End Try

    End Sub

    Private ReadOnly iconCache As ConcurrentDictionary(Of String, Bitmap) = QLIconCache.IconCache

    Private Sub OnDeleteDir(sender As System.IO.FileSystemWatcher, e As System.IO.FileSystemEventArgs)
        dBug.Print($"Delete Dir: {sender.NotifyFilter}")
        dBug.Print($"      Type: {e.ChangeType}")
        dBug.Print($"      Path: {e.FullPath}")

        If e.ChangeType = IO.WatcherChangeTypes.Deleted Then
            For Each key In iconCache.Keys.Where(Function(k As String) k.StartsWith(e.FullPath & "\"))
                iconCache.TryRemove(key, Nothing)
            Next
            For Each kvp In ResolvedLinkWatchers.ToList()
                If String.Equals(kvp.Key, e.FullPath, StringComparison.OrdinalIgnoreCase) Then
                    removeLinkWatcher(kvp.Key)
                End If
            Next
        End If
        InvalidateClipboardIfContains(e.FullPath)
    End Sub
    Private Sub OnRenamedDir(sender As System.IO.FileSystemWatcher, e As System.IO.RenamedEventArgs)
        dBug.Print($"Renamed Dir: {sender.NotifyFilter}")
        dBug.Print($"        Old: {e.OldFullPath}")
        dBug.Print($"        New: {e.FullPath}")

        For Each key In iconCache.Keys.Where(Function(k) k.StartsWith(e.OldFullPath & "\"))
            Dim item As Bitmap = Nothing
            If iconCache.TryRemove(key, item) Then iconCache.TryAdd(key.Replace(e.OldFullPath & "\", e.FullPath & "\"), item)
        Next
    End Sub
    Private Sub onDeleteFile(sender As System.IO.FileSystemWatcher, e As FileSystemEventArgs)
        dBug.Print($"Delete File: {sender.NotifyFilter}")
        dBug.Print($"       Type: {e.ChangeType}")
        dBug.Print($"       Path: {e.FullPath}")

        If e.ChangeType = WatcherChangeTypes.Deleted Then
            Dim res As Boolean = iconCache.TryRemove(e.FullPath, Nothing)
            If res Then Debug.Print($"File {e.FullPath} evicted")

            'remove from resolvedlinkwatchers '
            If e.FullPath.ToLower.EndsWith(".lnk") Then
                dBug.Print("Islink", 1)
                'get linkwatchertarget from cache
                Dim target As String = Nothing
                If ResolvedLinkLinks.TryGetValue(e.FullPath, target) Then
                    'if no other links point there remove the linkwatcher
                    If Not ResolvedLinkLinks.Any(Function(kvp) kvp.Key <> e.FullPath AndAlso kvp.Value = target) Then
                        dBug.Print($"ResolvedLinkLinks.Any")
                        removeLinkWatcher(target)
                    End If
                End If
            End If
        End If
        InvalidateClipboardIfContains(e.FullPath)
    End Sub

    Private Sub OnRenamed(sender As System.IO.FileSystemWatcher, e As System.IO.RenamedEventArgs)
        dBug.Print($"Renamed File: {sender.NotifyFilter}")
        dBug.Print($"         Old: {e.OldFullPath}")
        dBug.Print($"         New: {e.FullPath}")

        Dim item As Bitmap = Nothing
        If iconCache.TryRemove(e.OldFullPath, item) Then iconCache.TryAdd(e.FullPath, item)

    End Sub

    Private Sub OnWatcherError(sender As Object, e As System.IO.ErrorEventArgs)
        Dim ex As Exception = e.GetException()
        dBug.Print($"FileSystemWatcher error: {ex.Message}")
        If TypeOf ex Is InternalBufferOverflowException Then
            dBug.Print("FileSystemWatcher buffer overflow - some events may have been lost")
            'we don't know which. invalidate whole cache
            iconCache.Clear()
            ResolvedLinkwatchers_Clear()
        End If
    End Sub

    Private Sub OnChanged(sender As System.IO.FileSystemWatcher, e As System.IO.FileSystemEventArgs)
        If e.ChangeType = System.IO.WatcherChangeTypes.Changed Then
            dBug.Print(sender.ToString)
            dBug.Print($"Changed: {e.FullPath}")
            If e.FullPath.ToLower.EndsWith("desktop.ini") Then
                iconCache.TryRemove(e.FullPath.Substring(0, e.FullPath.LastIndexOf("\") + 1), Nothing)
            End If
            If hideExt.Contains(System.IO.Path.GetExtension(e.FullPath).ToLower) Then
                iconCache.TryRemove(e.FullPath, Nothing)
            End If
        End If
    End Sub

End Module
