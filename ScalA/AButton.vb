Public NotInheritable Class Dummy12
    'Dummy class
End Class
Public NotInheritable Class AButton
    Inherits Button

    ''' <summary>
    ''' Orange color used to warn about imminent inactivity timeout.
    ''' </summary>
    Private Shared ReadOnly COLOR_INACTIVITY_WARNING As Color = Color.FromArgb(&HFFFFA500)

    Public Sub New(ByVal text As String, ByVal left As Integer, ByVal top As Integer, ByVal width As Integer, ByVal height As Integer)
        MyBase.New()

        Me.Text = text
        Me.Location = New Point(left, top)
        Me.Size = New Size(width, height)
        Me.Name = "AButton"

        SetStyle(ControlStyles.AllPaintingInWmPaint Or
                 ControlStyles.Selectable Or
                 ControlStyles.DoubleBuffer Or
                 ControlStyles.ResizeRedraw Or
                 ControlStyles.SupportsTransparentBackColor Or
                 ControlStyles.UserPaint, True)
        Size = New Size(200, 150)
        Me.ImageAlign = ContentAlignment.TopRight
        Me.BackgroundImageLayout = ImageLayout.None
        Me.Padding = New Padding(0)
        Me.Margin = New Padding(0)
        Me.TextAlign = ContentAlignment.TopCenter
        Me.Font = NormalFont
        Me.Visible = False
        Me.BackColor = If(My.Settings.DarkMode, Color.DarkGray, COLOR_LIGHT_GRAY)
    End Sub

    Public Shared NormalFont As Font = New Font("Microsoft Sans Serif", 12, GraphicsUnit.Pixel)
    Public Shared BoldFont As Font = New Font("Microsoft Sans Serif", 13, FontStyle.Bold, GraphicsUnit.Pixel)

    ''' <summary>
    ''' When True, the button will display an orange warning background to indicate
    ''' the client is about to timeout due to inactivity.
    ''' </summary>
    Public Property IsInactivityWarning As Boolean = False

    Public pidCache As Integer
    Public AP As AstoniaProcess

    Private _passthrough As Rectangle
    Public ReadOnly Property ThumbRectangle() As Rectangle
        Get
            Return _passthrough
        End Get
    End Property
    Public ReadOnly Property ThumbRECT() As Rectangle
        Get ' New Rectangle(pnlStartup.Left + but.Left + 3, pnlStartup.Top + but.Top + 21, but.Right - 2, pnlStartup.Top + but.Bottom - 3)
            Return New Rectangle(Me.Parent.Left + Me.Left + 3, Me.Parent.Top + Me.Top + 21, Me.Right - 2, Me.Parent.Top + Me.Bottom - 3)
        End Get
    End Property
    Public ReadOnly Property ThumbContains(screenPt As Point) As Boolean
        Get
            Return Not Me.Disposing AndAlso Me.ThumbRectangle.Contains(Me.PointToClient(screenPt))
        End Get
    End Property

    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        _passthrough = New Rectangle(3, 21, Me.Width - 6, Me.Height - 24)
    End Sub

    Protected Overrides Sub OnMouseEnter(e As EventArgs)
        MyBase.OnMouseEnter(e)
        If IsInactivityWarning Then
            ' Keep warning color but slightly lighter on hover
            Me.BackColor = Color.FromArgb(&HFFFFB347) ' Lighter orange
        ElseIf My.Settings.DarkMode Then
            Me.BackColor = Color.FromArgb(&HFFA2A2A2)
        Else
            Me.BackColor = Color.FromArgb(&HFFE5F1FB)
        End If
    End Sub
    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        MyBase.OnMouseLeave(e)
        If IsInactivityWarning Then
            Me.BackColor = COLOR_INACTIVITY_WARNING
        ElseIf My.Settings.DarkMode Then
            Me.BackColor = If(Me.ContextMenuStrip?.Visible AndAlso Me.ContextMenuStrip Is FrmMain.cmsAlt, Color.FromArgb(&HFFA2A2A2), Color.DarkGray)
        Else
            Me.BackColor = If(Me.ContextMenuStrip?.Visible AndAlso Me.ContextMenuStrip Is FrmMain.cmsAlt, COLOR_HIGHLIGHT_BLUE, COLOR_LIGHT_GRAY)
        End If
    End Sub

    ''' <summary>
    ''' Updates the background color based on current state (inactivity warning, dark mode, etc.)
    ''' </summary>
    Public Sub UpdateBackColor()
        If IsInactivityWarning Then
            Me.BackColor = COLOR_INACTIVITY_WARNING
        ElseIf My.Settings.DarkMode Then
            Me.BackColor = Color.DarkGray
        Else
            Me.BackColor = COLOR_LIGHT_GRAY
        End If
    End Sub

    Public Shared ActiveOverview As Boolean = True
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        If My.Settings.DarkMode Then
            Using p As New Pen(Color.Gray, 2)
                e.Graphics.DrawRectangle(p, New Rectangle(3, 3, Me.Width - 6, Me.Height - 6))
            End Using
            If Me.Focused Then
                Using p As New Pen(Color.FromKnownColor(KnownColor.Highlight), 1)
                    e.Graphics.DrawRectangle(p, New Rectangle(2, 2, Me.Width - 5, Me.Height - 5))
                End Using
            End If
        End If
        If ActiveOverview AndAlso My.Settings.gameOnOverview AndAlso Me.Text <> "" Then
            Using b As New SolidBrush(Color.Magenta)
                e.Graphics.FillRectangle(b, Me.ThumbRectangle)
            End Using
        End If
    End Sub
End Class
