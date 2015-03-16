<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ctlCientList
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.btnCustom = New System.Windows.Forms.Button
        Me.cmbFilter = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.LV = New System.Windows.Forms.ListView
        Me.CtlCaptionSingle1 = New ReportInterface._Root_Controls.ctlCaptionSingle
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.Panel2.Controls.Add(Me.btnCustom)
        Me.Panel2.Controls.Add(Me.cmbFilter)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 32)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(804, 38)
        Me.Panel2.TabIndex = 4
        '
        'btnCustom
        '
        Me.btnCustom.Location = New System.Drawing.Point(347, 5)
        Me.btnCustom.Name = "btnCustom"
        Me.btnCustom.Size = New System.Drawing.Size(144, 26)
        Me.btnCustom.TabIndex = 2
        Me.btnCustom.Text = "My Custom Selection"
        Me.btnCustom.UseVisualStyleBackColor = True
        '
        'cmbFilter
        '
        Me.cmbFilter.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.cmbFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFilter.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbFilter.FormattingEnabled = True
        Me.cmbFilter.Location = New System.Drawing.Point(107, 7)
        Me.cmbFilter.Name = "cmbFilter"
        Me.cmbFilter.Size = New System.Drawing.Size(135, 23)
        Me.cmbFilter.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(0, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(107, 38)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Filter Clients by: "
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LV
        '
        Me.LV.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LV.Location = New System.Drawing.Point(0, 70)
        Me.LV.Name = "LV"
        Me.LV.Size = New System.Drawing.Size(804, 546)
        Me.LV.TabIndex = 5
        Me.LV.UseCompatibleStateImageBehavior = False
        '
        'CtlCaptionSingle1
        '
        Me.CtlCaptionSingle1.Caption = "Web Service Client List"
        Me.CtlCaptionSingle1.Dock = System.Windows.Forms.DockStyle.Top
        Me.CtlCaptionSingle1.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold)
        Me.CtlCaptionSingle1.Location = New System.Drawing.Point(0, 0)
        Me.CtlCaptionSingle1.Name = "CtlCaptionSingle1"
        Me.CtlCaptionSingle1.Size = New System.Drawing.Size(804, 32)
        Me.CtlCaptionSingle1.TabIndex = 3
        '
        'ctlCientList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.LV)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.CtlCaptionSingle1)
        Me.Name = "ctlCientList"
        Me.Size = New System.Drawing.Size(804, 616)
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CtlCaptionSingle1 As ReportInterface._Root_Controls.ctlCaptionSingle
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents cmbFilter As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LV As System.Windows.Forms.ListView
    Friend WithEvents btnCustom As System.Windows.Forms.Button

End Class
