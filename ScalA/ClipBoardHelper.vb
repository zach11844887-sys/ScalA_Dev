Imports System.Runtime.InteropServices

Public Module ClipBoardHelper

    Public clipBoardInfo As ClipboardFileInfo

    Private _ListeningFormHand As IntPtr

    Public Sub registerClipListener()
        If AddClipboardFormatListener(ScalaHandle) Then _ListeningFormHand = ScalaHandle
    End Sub

    Public Sub unregisterClipListener()
        If RemoveClipboardFormatListener(_ListeningFormHand) Then _ListeningFormHand = Nothing
    End Sub


#If DEBUG Then
    ''' <summary>
    ''' do not use, can lead to silent crash
    ''' </summary>
    ''' <param name="file"></param>
    ''' <param name="isCut"></param>
    Public Sub SetFileDropListWithEffect(ByVal file As String, ByVal isCut As Boolean)
        Try

            Dim data As New DataObject()
            Dim files As New Specialized.StringCollection From {file}

            data.SetFileDropList(files)

            If Not isCut Then
                unregisterClipListener()
                InvokeExplorerVerb(file, "Copy", ScalaHandle)

                data.SetData("Shell IDList Array", Clipboard.GetDataObject().GetData("Shell IDList Array", True))

                registerClipListener()
            End If


            data.SetData("Preferred DropEffect", New IO.MemoryStream(BitConverter.GetBytes(If(isCut, DragDropEffects.Move, DragDropEffects.Copy Or DragDropEffects.Link))))

            Clipboard.SetDataObject(data, True)

            dBug.Print("Clipboard set: " & If(isCut, "Cut", "Copy"))
        Catch ex As Exception
            dBug.Print("Clipboard error: " & ex.Message)
        End Try

    End Sub
#End If
    ''' <summary>
    ''' Information about files currently on the clipboard
    ''' </summary>
    Public Structure ClipboardFileInfo
        Public Files As List(Of String)
        Public Action As DragDropEffects
    End Structure

    Public Function GetClipboardFilesAndAction() As ClipboardFileInfo
        Dim result As New ClipboardFileInfo With {
            .Files = New List(Of String),
            .Action = DragDropEffects.None
        }

        result.Files = Clipboard.GetFileDropList().Cast(Of String)().ToList

        Dim dataObj = Clipboard.GetDataObject()
        If dataObj IsNot Nothing Then
            'Dim dropEffectObj = dataObj.GetData(DataFormats.GetFormat(PreferredDropEffect).Name)
            Dim dropEffectObj = dataObj.GetData("Preferred DropEffect")
            'Debug.Print($"""{dropEffectObj?.GetType()}""")
            If TypeOf dropEffectObj Is IO.MemoryStream Then
                Dim mes = DirectCast(dropEffectObj, IO.MemoryStream)
                If mes.Length >= 4 Then
                    Dim bytes(3) As Byte
                    mes.Position = 0
                    mes.Read(bytes, 0, 4)
                    result.Action = BitConverter.ToInt32(bytes, 0)
                End If
            End If
        End If

#If DEBUG Then
        Dim count As Integer = result.Files?.Count
        dBug.Print($"Clipboard Contains: {count} File{If(count = 1, "", "s")}, Action: {result.Action}")
        For Each it As String In result.Files
            dBug.Print($"""{it}""")
        Next
        dBug.Print("---")
#End If

        Return result

    End Function

    Public Sub InvokeExplorerVerb(filePath As String, verb As String, Optional hwnd As IntPtr = Nothing)
        Dim exec As New SHELLEXECUTEINFO()
        exec.cbSize = Marshal.SizeOf(exec)
        exec.fMask = SEE_MASK_INVOKEIDLIST
        exec.hwnd = hwnd
        exec.lpVerb = verb   ' canonical verb: "cut", "copy", "paste", "pastelink", "delete", "properties"
        exec.lpFile = filePath
        exec.nShow = SW_HIDE

        If Not ShellExecuteEx(exec) Then
            Debug.Print(New ComponentModel.Win32Exception(Marshal.GetLastWin32Error()).Message)
        End If
    End Sub


    ''' <summary>
    ''' Clears the clipboard if it contains the specified path.
    ''' Must be called when a file/folder that might be in the clipboard is deleted.
    ''' </summary>
    ''' <param name="deletedPath">Path of the deleted file or folder</param>
    Public Sub InvalidateClipboardIfContains(deletedPath As String)
        If clipBoardInfo.Files Is Nothing OrElse clipBoardInfo.Files.Count = 0 Then
            Exit Sub
        End If

        ' Check if deleted path is in clipboard (exact match or parent folder)
        Dim normalizedDeleted = deletedPath.TrimEnd("\"c).ToLowerInvariant()
        Dim shouldClear = clipBoardInfo.Files.Any(Function(f)
                                                      Dim normalizedClip = f.TrimEnd("\"c).ToLowerInvariant()
                                                      Return normalizedClip = normalizedDeleted OrElse
                                                             normalizedClip.StartsWith(normalizedDeleted & "\")
                                                  End Function)

        If shouldClear Then
            dBug.Print($"Clipboard invalidated: deleted path was in clipboard - {deletedPath}")
            ' Must invoke on UI thread since Clipboard requires STA
            If FrmMain.InvokeRequired Then
                FrmMain.BeginInvoke(Sub()
                                        Clipboard.Clear()
                                        clipBoardInfo = New ClipboardFileInfo With {
                                            .Files = New List(Of String),
                                            .Action = DragDropEffects.None
                                        }
                                    End Sub)
            Else
                Clipboard.Clear()
                clipBoardInfo = New ClipboardFileInfo With {
                    .Files = New List(Of String),
                    .Action = DragDropEffects.None
                }
            End If
        End If
    End Sub

End Module
