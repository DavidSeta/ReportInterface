<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Me.components = New System.ComponentModel.Container
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.lblLine = New System.Windows.Forms.Label
        Me.lblStatus = New System.Windows.Forms.Label
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.Button3 = New System.Windows.Forms.Button
        Me.btnCheckForProcessed = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.Label1 = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.lblEndDate = New System.Windows.Forms.Label
        Me.lblStartDate = New System.Windows.Forms.Label
        Me.dtEnd = New System.Windows.Forms.DateTimePicker
        Me.dtStart = New System.Windows.Forms.DateTimePicker
        Me.btnReset = New System.Windows.Forms.Button
        Me.btnProcess = New System.Windows.Forms.Button
        Me.btnDownloadProcess = New System.Windows.Forms.Button
        Me.btnDownload = New System.Windows.Forms.Button
        Me.cmbWebService = New System.Windows.Forms.ComboBox
        Me.ClientList = New ReportInterface.ctlCientList
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Timer1
        '
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.Panel2.Controls.Add(Me.lblLine)
        Me.Panel2.Controls.Add(Me.lblStatus)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 642)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1042, 24)
        Me.Panel2.TabIndex = 11
        '
        'lblLine
        '
        Me.lblLine.Dock = System.Windows.Forms.DockStyle.Right
        Me.lblLine.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLine.Location = New System.Drawing.Point(896, 0)
        Me.lblLine.Name = "lblLine"
        Me.lblLine.Size = New System.Drawing.Size(146, 24)
        Me.lblLine.TabIndex = 1
        Me.lblLine.Text = "Line"
        Me.lblLine.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblStatus
        '
        Me.lblStatus.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus.Location = New System.Drawing.Point(0, 0)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(730, 24)
        Me.lblStatus.TabIndex = 0
        Me.lblStatus.Text = "Status"
        Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.Button3)
        Me.Panel3.Controls.Add(Me.btnCheckForProcessed)
        Me.Panel3.Controls.Add(Me.Button2)
        Me.Panel3.Controls.Add(Me.Button1)
        Me.Panel3.Controls.Add(Me.Label8)
        Me.Panel3.Controls.Add(Me.Label7)
        Me.Panel3.Controls.Add(Me.Label6)
        Me.Panel3.Controls.Add(Me.Label5)
        Me.Panel3.Controls.Add(Me.Label4)
        Me.Panel3.Controls.Add(Me.Label3)
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Controls.Add(Me.Panel4)
        Me.Panel3.Controls.Add(Me.Label1)
        Me.Panel3.Controls.Add(Me.Panel1)
        Me.Panel3.Controls.Add(Me.lblEndDate)
        Me.Panel3.Controls.Add(Me.lblStartDate)
        Me.Panel3.Controls.Add(Me.dtEnd)
        Me.Panel3.Controls.Add(Me.dtStart)
        Me.Panel3.Controls.Add(Me.btnReset)
        Me.Panel3.Controls.Add(Me.btnProcess)
        Me.Panel3.Controls.Add(Me.btnDownloadProcess)
        Me.Panel3.Controls.Add(Me.btnDownload)
        Me.Panel3.Controls.Add(Me.cmbWebService)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(804, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(238, 642)
        Me.Panel3.TabIndex = 13
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(51, 556)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(121, 26)
        Me.Button3.TabIndex = 47
        Me.Button3.Text = "Flex Reset"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'btnCheckForProcessed
        '
        Me.btnCheckForProcessed.Location = New System.Drawing.Point(51, 277)
        Me.btnCheckForProcessed.Name = "btnCheckForProcessed"
        Me.btnCheckForProcessed.Size = New System.Drawing.Size(121, 26)
        Me.btnCheckForProcessed.TabIndex = 46
        Me.btnCheckForProcessed.Text = "Check For Processed"
        Me.btnCheckForProcessed.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(51, 524)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(121, 26)
        Me.Button2.TabIndex = 45
        Me.Button2.Text = "Feb 2011 EOM"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(51, 492)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(121, 26)
        Me.Button1.TabIndex = 44
        Me.Button1.Text = "Reset PPC Bid EOM"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(16, 123)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(28, 13)
        Me.Label8.TabIndex = 43
        Me.Label8.Text = "WS:"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(31, 464)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(187, 15)
        Me.Label7.TabIndex = 42
        Me.Label7.Text = "multiple days in one download."
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(13, 449)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(205, 15)
        Me.Label6.TabIndex = 41
        Me.Label6.Text = "3.  Both MSN and Google can process"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(32, 410)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(194, 15)
        Me.Label5.TabIndex = 40
        Me.Label5.Text = "dates but, will put all data into the end."
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 395)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(187, 15)
        Me.Label4.TabIndex = 39
        Me.Label4.Text = "2. Verizon/ASK will process multiple "
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(31, 363)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(188, 18)
        Me.Label3.TabIndex = 38
        Me.Label3.Text = "EndDate Value."
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(16, 346)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(210, 17)
        Me.Label2.TabIndex = 37
        Me.Label2.Text = "1. Yahoo will only process the"
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.Black
        Me.Panel4.Location = New System.Drawing.Point(9, 309)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(221, 2)
        Me.Panel4.TabIndex = 36
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(13, 323)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 15)
        Me.Label1.TabIndex = 35
        Me.Label1.Text = "Notes:"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Black
        Me.Panel1.Location = New System.Drawing.Point(7, 110)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(221, 2)
        Me.Panel1.TabIndex = 34
        '
        'lblEndDate
        '
        Me.lblEndDate.AutoSize = True
        Me.lblEndDate.Location = New System.Drawing.Point(13, 60)
        Me.lblEndDate.Name = "lblEndDate"
        Me.lblEndDate.Size = New System.Drawing.Size(55, 13)
        Me.lblEndDate.TabIndex = 33
        Me.lblEndDate.Text = "End Date:"
        '
        'lblStartDate
        '
        Me.lblStartDate.AutoSize = True
        Me.lblStartDate.Location = New System.Drawing.Point(13, 9)
        Me.lblStartDate.Name = "lblStartDate"
        Me.lblStartDate.Size = New System.Drawing.Size(58, 13)
        Me.lblStartDate.TabIndex = 32
        Me.lblStartDate.Text = "Start Date:"
        '
        'dtEnd
        '
        Me.dtEnd.Location = New System.Drawing.Point(16, 76)
        Me.dtEnd.Name = "dtEnd"
        Me.dtEnd.Size = New System.Drawing.Size(204, 20)
        Me.dtEnd.TabIndex = 31
        '
        'dtStart
        '
        Me.dtStart.Location = New System.Drawing.Point(16, 25)
        Me.dtStart.Name = "dtStart"
        Me.dtStart.Size = New System.Drawing.Size(204, 20)
        Me.dtStart.TabIndex = 30
        '
        'btnReset
        '
        Me.btnReset.Location = New System.Drawing.Point(51, 247)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(121, 24)
        Me.btnReset.TabIndex = 29
        Me.btnReset.Text = "Reset to Process"
        Me.btnReset.UseVisualStyleBackColor = True
        '
        'btnProcess
        '
        Me.btnProcess.Location = New System.Drawing.Point(51, 217)
        Me.btnProcess.Name = "btnProcess"
        Me.btnProcess.Size = New System.Drawing.Size(121, 24)
        Me.btnProcess.TabIndex = 28
        Me.btnProcess.Text = "Process Only"
        Me.btnProcess.UseVisualStyleBackColor = True
        '
        'btnDownloadProcess
        '
        Me.btnDownloadProcess.Location = New System.Drawing.Point(51, 188)
        Me.btnDownloadProcess.Name = "btnDownloadProcess"
        Me.btnDownloadProcess.Size = New System.Drawing.Size(121, 24)
        Me.btnDownloadProcess.TabIndex = 27
        Me.btnDownloadProcess.Text = "Download && Process"
        Me.btnDownloadProcess.UseVisualStyleBackColor = True
        '
        'btnDownload
        '
        Me.btnDownload.Location = New System.Drawing.Point(51, 158)
        Me.btnDownload.Name = "btnDownload"
        Me.btnDownload.Size = New System.Drawing.Size(121, 24)
        Me.btnDownload.TabIndex = 26
        Me.btnDownload.Text = "Download Only"
        Me.btnDownload.UseVisualStyleBackColor = True
        '
        'cmbWebService
        '
        Me.cmbWebService.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbWebService.FormattingEnabled = True
        Me.cmbWebService.Location = New System.Drawing.Point(51, 120)
        Me.cmbWebService.Name = "cmbWebService"
        Me.cmbWebService.Size = New System.Drawing.Size(121, 21)
        Me.cmbWebService.TabIndex = 25
        '
        'ClientList
        '
        Me.ClientList.Dock = System.Windows.Forms.DockStyle.Left
        Me.ClientList.Location = New System.Drawing.Point(0, 0)
        Me.ClientList.Name = "ClientList"
        Me.ClientList.Size = New System.Drawing.Size(804, 642)
        Me.ClientList.TabIndex = 12
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1042, 666)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.ClientList)
        Me.Controls.Add(Me.Panel2)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Report Data Interface"
        Me.Panel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents ClientList As ReportInterface.ctlCientList
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents lblLine As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lblEndDate As System.Windows.Forms.Label
    Friend WithEvents lblStartDate As System.Windows.Forms.Label
    Friend WithEvents dtEnd As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnReset As System.Windows.Forms.Button
    Friend WithEvents btnProcess As System.Windows.Forms.Button
    Friend WithEvents btnDownloadProcess As System.Windows.Forms.Button
    Friend WithEvents btnDownload As System.Windows.Forms.Button
    Friend WithEvents cmbWebService As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents btnCheckForProcessed As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
End Class
