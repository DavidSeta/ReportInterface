' Custom control that draws the caption for each pane. Contains an active 
' state and draws the caption different for each state. Caption is drawn
' with a gradient fill and antialias font.

Imports System.Drawing.Drawing2D
Imports System.ComponentModel

Namespace _Root_Controls
    Friend Class ctlCaptionSingle
        Inherits System.Windows.Forms.UserControl

        ' const values
        Private Class Consts
            Public Const DefaultHeight As Integer = 26
            Public Const DefaultFontName As String = "Tahoma"
            Public Const DefaultFontSize As Integer = 12
            Public Const PosOffset As Integer = 4
        End Class

        ' internal members
        Private m_active As Boolean = False
        Private m_antiAlias As Boolean = True
        Private m_allowActive As Boolean = True
        Private m_text As String = ""

        Private m_colorActiveText As Color = Color.Black
        Private m_colorInactiveText As Color = Color.White

        Private m_colorActiveLow As Color = Color.FromArgb(255, 165, 78)
        Private m_colorActiveHigh As Color = Color.FromArgb(255, 225, 155)
        Private m_colorInactiveLow As Color = Color.FromArgb(3, 55, 145)
        Private m_colorInactiveHigh As Color = Color.FromArgb(90, 135, 215)

        ' gdi objects
        Private m_brushActiveText As SolidBrush
        Private m_brushInactiveText As SolidBrush
        Private m_brushActive As LinearGradientBrush
        Private m_brushInactive As LinearGradientBrush
        Private m_format As StringFormat


        ' public properties

        ' the caption of the control
        <Description("Text displayed in the caption."), Category("Appearance"), DefaultValue("")> _
        Public Property Caption() As String
            Get
                Return m_text
            End Get

            Set(ByVal value As String)
                m_text = value
                Invalidate()
            End Set
        End Property

        Public Overrides Property Text() As String
            Get
                Return Me.Caption
            End Get
            Set(ByVal Value As String)
                Me.Caption = Value
            End Set
        End Property

        ' if the caption is active or not
        <Description("The active state of the caption, draws the caption with different gradient colors."), Category("Appearance"), DefaultValue(False)> _
        Public Property Active() As Boolean
            Get
                Return m_active
            End Get
            Set(ByVal value As Boolean)
                m_active = value
                Invalidate()
            End Set
        End Property

        ' if should maintain an active and inactive state
        <Description("True always uses the inactive state colors, false maintains an active and inactive state."), Category("Appearance"), DefaultValue(True)> _
        Public Property AllowActive() As Boolean
            Get
                Return m_allowActive
            End Get
            Set(ByVal value As Boolean)
                m_allowActive = value
                Invalidate()
            End Set
        End Property

        ' if the caption is active or not
        <Description("If should draw the text as antialiased."), Category("Appearance"), DefaultValue(True)> _
        Public Property AntiAlias() As Boolean
            Get
                Return m_antiAlias
            End Get
            Set(ByVal value As Boolean)
                m_antiAlias = value
                Invalidate()
            End Set
        End Property

#Region " Color Properties "

        <Description("Color of the text when active."), Category("Appearance"), DefaultValue(GetType(Color), "Black")> _
        Public Property ActiveTextColor() As Color
            Get
                Return m_colorActiveText
            End Get
            Set(ByVal Value As Color)
                If Value.Equals(Color.Empty) Then Value = Color.Black
                m_colorActiveText = Value
                m_brushActiveText = New SolidBrush(m_colorActiveText)
                Invalidate()
            End Set
        End Property

        <Description("Color of the text when inactive."), Category("Appearance"), DefaultValue(GetType(Color), "White")> _
        Public Property InactiveTextColor() As Color
            Get
                Return m_colorInactiveText
            End Get
            Set(ByVal Value As Color)
                If Value.Equals(Color.Empty) Then Value = Color.White
                m_colorInactiveText = Value
                m_brushInactiveText = New SolidBrush(m_colorInactiveText)
                Invalidate()
            End Set
        End Property

        <Description("Low color of the active gradient."), Category("Appearance"), DefaultValue(GetType(Color), "255, 165, 78")> _
        Public Property ActiveGradientLowColor() As Color
            Get
                Return m_colorActiveLow
            End Get
            Set(ByVal Value As Color)
                If Value.Equals(Color.Empty) Then Value = Color.FromArgb(255, 165, 78)
                m_colorActiveLow = Value
                CreateGradientBrushes()
                Invalidate()
            End Set
        End Property

        <Description("High color of the active gradient."), Category("Appearance"), DefaultValue(GetType(Color), "255, 225, 155")> _
        Public Property ActiveGradientHighColor() As Color
            Get
                Return m_colorActiveHigh
            End Get
            Set(ByVal Value As Color)
                If Value.Equals(Color.Empty) Then Value = Color.FromArgb(255, 225, 155)
                m_colorActiveHigh = Value
                CreateGradientBrushes()
                Invalidate()
            End Set
        End Property

        <Description("Low color of the inactive gradient."), Category("Appearance"), DefaultValue(GetType(Color), "3, 55, 145")> _
          Public Property InactiveGradientLowColor() As Color
            Get
                Return m_colorInactiveLow
            End Get
            Set(ByVal Value As Color)
                If Value.Equals(Color.Empty) Then Value = Color.FromArgb(3, 55, 145)
                m_colorInactiveLow = Value
                CreateGradientBrushes()
                Invalidate()
            End Set
        End Property

        <Description("High color of the inactive gradient."), Category("Appearance"), DefaultValue(GetType(Color), "90, 135, 215")> _
          Public Property InactiveGradientHighColor() As Color
            Get
                Return m_colorInactiveHigh
            End Get
            Set(ByVal Value As Color)
                If Value.Equals(Color.Empty) Then Value = Color.FromArgb(90, 135, 215)
                m_colorInactiveHigh = Value
                CreateGradientBrushes()
                Invalidate()
            End Set
        End Property

#End Region

        ' internal properties

        ' brush used to draw the caption
        Private ReadOnly Property TextBrush() As SolidBrush
            Get
                Return CType(IIf(m_active AndAlso m_allowActive, _
                 m_brushActiveText, m_brushInactiveText), SolidBrush)
            End Get
        End Property

        ' gradient brush for the background
        Private ReadOnly Property BackBrush() As LinearGradientBrush
            Get
                Return CType(IIf(m_active AndAlso m_allowActive, _
                 m_brushActive, m_brushInactive), LinearGradientBrush)
            End Get
        End Property

        ' ctor
        Public Sub New()
            MyBase.New()

            ' this call is required by the Windows Form Designer
            InitializeComponent()

            ' set double buffer styles
            Me.SetStyle(ControlStyles.DoubleBuffer Or ControlStyles.UserPaint Or _
             ControlStyles.AllPaintingInWmPaint Or ControlStyles.ResizeRedraw, True)

            ' init the height
            Me.Height = Consts.DefaultHeight

            ' format used when drawing the text
            m_format = New StringFormat
            m_format.FormatFlags = StringFormatFlags.NoWrap
            m_format.LineAlignment = StringAlignment.Center
            m_format.Trimming = StringTrimming.EllipsisCharacter

            ' init the font
            Me.Font = New Font(Consts.DefaultFontName, Consts.DefaultFontSize, FontStyle.Bold)

            ' create gdi objects
            Me.ActiveTextColor = m_colorActiveText
            Me.InactiveTextColor = m_colorInactiveText

            ' setting the height above actually does this, but leave
            ' in incase change the code (and forget to init the 
            ' gradient brushes)
            CreateGradientBrushes()
        End Sub

        ' internal methods

        ' the caption needs to be drawn
        Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
            DrawCaption(e.Graphics)
            MyBase.OnPaint(e)
        End Sub

        ' draw the caption
        Private Sub DrawCaption(ByVal g As Graphics)
            ' background
            g.FillRectangle(Me.BackBrush, Me.DisplayRectangle)

            ' caption
            If m_antiAlias Then
                g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
            End If

            ' need a rectangle when want to use ellipsis
            Dim bounds As RectangleF = New RectangleF(Consts.PosOffset, 0, _
             Me.DisplayRectangle.Width - Consts.PosOffset, Me.DisplayRectangle.Height)

            g.DrawString(m_text, Me.Font, Me.TextBrush, bounds, m_format)
        End Sub

        ' clicking on the caption does not give focus,
        ' handle the mouse down event and set focus to self
        Protected Overrides Sub OnMouseDown(ByVal e As MouseEventArgs)
            MyBase.OnMouseDown(e)
            If Me.m_allowActive Then Me.Focus()
        End Sub

        Protected Overrides Sub OnSizeChanged(ByVal e As System.EventArgs)
            MyBase.OnSizeChanged(e)

            ' create the gradient brushes based on the new size
            CreateGradientBrushes()
        End Sub

        Private Sub CreateGradientBrushes()
            ' can only create brushes when have a width and height
            If Me.Width > 0 AndAlso Me.Height > 0 Then
                If Not (m_brushActive Is Nothing) Then m_brushActive.Dispose()
                m_brushActive = New LinearGradientBrush(Me.DisplayRectangle, _
                 m_colorActiveHigh, m_colorActiveLow, LinearGradientMode.Vertical)

                If Not (m_brushInactive Is Nothing) Then m_brushInactive.Dispose()
                m_brushInactive = New LinearGradientBrush(Me.DisplayRectangle, _
                  m_colorInactiveHigh, m_colorInactiveLow, LinearGradientMode.Vertical)
            End If
        End Sub


#Region " Windows Form Designer generated code "

        'UserControl overrides dispose to clean up the component list.
        Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing Then
                If Not (components Is Nothing) Then
                    components.Dispose()
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub

        'Required by the Windows Form Designer
        Private components As System.ComponentModel.IContainer

        'NOTE: The following procedure is required by the Windows Form Designer
        'It can be modified using the Windows Form Designer.  
        'Do not modify it using the code editor.
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
            '
            'PaneCaption
            '
            Me.Name = "PaneCaption"
            Me.Size = New System.Drawing.Size(150, 30)
        End Sub

#End Region

    End Class
End Namespace
