Imports System.ComponentModel

Namespace _Composite_Controls
    Friend Class ctlCaptionDouble
        Inherits System.Windows.Forms.UserControl

#Region " Title Properties "

        <Description("Text displayed in the Title line."), Category("Appearance"), DefaultValue("Title")> _
        Public Property Title() As String
            Get
                Return caTitle.Caption
            End Get

            Set(ByVal value As String)
                caTitle.Caption = value
            End Set
        End Property

        <Description("Low color of the Title active gradient."), Category("Appearance"), DefaultValue(GetType(Color), "255, 165, 78")> _
        Public Property Title_ActiveGradientLowColor() As Color
            Get
                Return caTitle.ActiveGradientLowColor
            End Get
            Set(ByVal Value As Color)
                If Value.Equals(Color.Empty) Then Value = Color.FromArgb(255, 165, 78)
                caTitle.ActiveGradientLowColor = Value
            End Set
        End Property

        <Description("High color of the Title active gradient."), Category("Appearance"), DefaultValue(GetType(Color), "255, 225, 155")> _
        Public Property Title_ActiveGradientHighColor() As Color
            Get
                Return caTitle.ActiveGradientHighColor
            End Get
            Set(ByVal Value As Color)
                If Value.Equals(Color.Empty) Then Value = Color.FromArgb(255, 225, 155)
                caTitle.ActiveGradientHighColor = Value
            End Set
        End Property

        <Description("Low color of the Title inactive gradient."), Category("Appearance"), DefaultValue(GetType(Color), "3, 55, 145")> _
          Public Property Title_InactiveGradientLowColor() As Color
            Get
                Return caTitle.InactiveGradientLowColor
            End Get
            Set(ByVal Value As Color)
                If Value.Equals(Color.Empty) Then Value = Color.FromArgb(3, 55, 145)
                caTitle.InactiveGradientLowColor = Value
            End Set
        End Property

        <Description("High color of the Title inactive gradient."), Category("Appearance"), DefaultValue(GetType(Color), "90, 135, 215")> _
          Public Property Title_InactiveGradientHighColor() As Color
            Get
                Return caTitle.InactiveGradientHighColor
            End Get
            Set(ByVal Value As Color)
                If Value.Equals(Color.Empty) Then Value = Color.FromArgb(90, 135, 215)
                caTitle.InactiveGradientHighColor = Value
            End Set
        End Property

        <Description("Color of the Title text when active."), Category("Appearance"), DefaultValue(GetType(Color), "Black")> _
        Public Property Title_ActiveTextColor() As Color
            Get
                Return caTitle.ActiveTextColor
            End Get
            Set(ByVal Value As Color)
                If Value.Equals(Color.Empty) Then Value = Color.Black
                caTitle.ActiveTextColor = Value
            End Set
        End Property

        <Description("Color of the Title text when inactive."), Category("Appearance"), DefaultValue(GetType(Color), "White")> _
        Public Property Title_InactiveTextColor() As Color
            Get
                Return caTitle.InactiveTextColor
            End Get
            Set(ByVal Value As Color)
                If Value.Equals(Color.Empty) Then Value = Color.White
                caTitle.InactiveTextColor = Value
            End Set
        End Property

#End Region

#Region " Description Properties "

        <Description("Text displayed in the Description line."), Category("Appearance"), DefaultValue("Description")> _
        Public Property Description() As String
            Get
                Return caDescription.Caption
            End Get

            Set(ByVal value As String)
                caDescription.Caption = value
            End Set
        End Property

        <Description("Low color of the Description active gradient."), Category("Appearance"), DefaultValue(GetType(Color), "255, 165, 78")> _
        Public Property Description_ActiveGradientLowColor() As Color
            Get
                Return caDescription.ActiveGradientLowColor
            End Get
            Set(ByVal Value As Color)
                If Value.Equals(Color.Empty) Then Value = Color.FromArgb(255, 165, 78)
                caDescription.ActiveGradientLowColor = Value
            End Set
        End Property

        <Description("High color of the Description active gradient."), Category("Appearance"), DefaultValue(GetType(Color), "255, 225, 155")> _
        Public Property Description_ActiveGradientHighColor() As Color
            Get
                Return caDescription.ActiveGradientHighColor
            End Get
            Set(ByVal Value As Color)
                If Value.Equals(Color.Empty) Then Value = Color.FromArgb(255, 225, 155)
                caDescription.ActiveGradientHighColor = Value
            End Set
        End Property

        <Description("Low color of the Description inactive gradient."), Category("Appearance"), DefaultValue(GetType(Color), "3, 55, 145")> _
          Public Property Description_InactiveGradientLowColor() As Color
            Get
                Return caDescription.InactiveGradientLowColor
            End Get
            Set(ByVal Value As Color)
                If Value.Equals(Color.Empty) Then Value = Color.FromArgb(3, 55, 145)
                caDescription.InactiveGradientLowColor = Value
            End Set
        End Property

        <Description("High color of the Description inactive gradient."), Category("Appearance"), DefaultValue(GetType(Color), "90, 135, 215")> _
          Public Property Description_InactiveGradientHighColor() As Color
            Get
                Return caDescription.InactiveGradientHighColor
            End Get
            Set(ByVal Value As Color)
                If Value.Equals(Color.Empty) Then Value = Color.FromArgb(90, 135, 215)
                caDescription.InactiveGradientHighColor = Value
            End Set
        End Property

        <Description("Color of the Description text when active."), Category("Appearance"), DefaultValue(GetType(Color), "Black")> _
        Public Property Description_ActiveTextColor() As Color
            Get
                Return caDescription.ActiveTextColor
            End Get
            Set(ByVal Value As Color)
                If Value.Equals(Color.Empty) Then Value = Color.Black
                caDescription.ActiveTextColor = Value
            End Set
        End Property

        <Description("Color of the Description text when inactive."), Category("Appearance"), DefaultValue(GetType(Color), "White")> _
        Public Property Description_InactiveTextColor() As Color
            Get
                Return caDescription.InactiveTextColor
            End Get
            Set(ByVal Value As Color)
                If Value.Equals(Color.Empty) Then Value = Color.White
                caDescription.InactiveTextColor = Value
            End Set
        End Property

#End Region

#Region " Windows Form Designer generated code "

        Public Sub New()
            MyBase.New()

            'This call is required by the Windows Form Designer.
            InitializeComponent()

            'Add any initialization after the InitializeComponent() call

        End Sub

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
        Friend WithEvents caTitle As _Root_Controls.ctlCaptionSingle
        Friend WithEvents caDescription As _Root_Controls.ctlCaptionSingle
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
            Me.caTitle = New _Root_Controls.ctlCaptionSingle
            Me.caDescription = New _Root_Controls.ctlCaptionSingle
            Me.SuspendLayout()
            '
            'caTitle
            '
            Me.caTitle.AllowActive = False
            Me.caTitle.Caption = "Title"
            Me.caTitle.Dock = System.Windows.Forms.DockStyle.Top
            Me.caTitle.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold)
            Me.caTitle.Location = New System.Drawing.Point(0, 0)
            Me.caTitle.Name = "caTitle"
            Me.caTitle.Size = New System.Drawing.Size(150, 32)
            Me.caTitle.TabIndex = 0
            '
            'caDescription
            '
            Me.caDescription.Caption = "Description"
            Me.caDescription.Dock = System.Windows.Forms.DockStyle.Top
            Me.caDescription.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.caDescription.InactiveGradientHighColor = System.Drawing.Color.FromArgb(CType(175, Byte), CType(200, Byte), CType(245, Byte))
            Me.caDescription.InactiveGradientLowColor = System.Drawing.Color.FromArgb(CType(205, Byte), CType(225, Byte), CType(255, Byte))
            Me.caDescription.InactiveTextColor = System.Drawing.Color.Black
            Me.caDescription.Location = New System.Drawing.Point(0, 32)
            Me.caDescription.Name = "caDescription"
            Me.caDescription.Size = New System.Drawing.Size(150, 30)
            Me.caDescription.TabIndex = 1
            '
            'Caption_Double
            '
            Me.Controls.Add(Me.caDescription)
            Me.Controls.Add(Me.caTitle)
            Me.Name = "Caption_Double"
            Me.Size = New System.Drawing.Size(150, 56)
            Me.ResumeLayout(False)

        End Sub

#End Region

        Private Sub Caption_Double_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Me.Height = caDescription.Top + caDescription.Height + 1
        End Sub
    End Class
End Namespace
