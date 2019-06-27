<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtLargeInbound = New System.Windows.Forms.TextBox()
        Me.txtLargeStorage = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.txtLabour = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblDB = New System.Windows.Forms.Label()
        Me.lblReports = New System.Windows.Forms.Label()
        Me.txtReports = New System.Windows.Forms.TextBox()
        Me.txtDB = New System.Windows.Forms.TextBox()
        Me.btnDBBrowse = New System.Windows.Forms.Button()
        Me.btnReportBrowse = New System.Windows.Forms.Button()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(19, 91)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(283, 24)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Number of large boxes inbound:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(19, 133)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(295, 24)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Number of large boxes in storage:"
        '
        'txtLargeInbound
        '
        Me.txtLargeInbound.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLargeInbound.Location = New System.Drawing.Point(335, 86)
        Me.txtLargeInbound.Margin = New System.Windows.Forms.Padding(4)
        Me.txtLargeInbound.Name = "txtLargeInbound"
        Me.txtLargeInbound.Size = New System.Drawing.Size(143, 29)
        Me.txtLargeInbound.TabIndex = 0
        '
        'txtLargeStorage
        '
        Me.txtLargeStorage.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLargeStorage.Location = New System.Drawing.Point(335, 130)
        Me.txtLargeStorage.Margin = New System.Windows.Forms.Padding(4)
        Me.txtLargeStorage.Name = "txtLargeStorage"
        Me.txtLargeStorage.Size = New System.Drawing.Size(143, 29)
        Me.txtLargeStorage.TabIndex = 1
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(266, 235)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(213, 63)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Generate Report"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'txtLabour
        '
        Me.txtLabour.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLabour.Location = New System.Drawing.Point(335, 166)
        Me.txtLabour.Margin = New System.Windows.Forms.Padding(4)
        Me.txtLabour.Name = "txtLabour"
        Me.txtLabour.Size = New System.Drawing.Size(143, 29)
        Me.txtLabour.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(19, 170)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(144, 24)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Hours of labour:"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 302)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(492, 25)
        Me.StatusStrip1.TabIndex = 7
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(50, 20)
        Me.ToolStripStatusLabel1.Text = "Ready"
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CustomFormat = "MMMMyyyy"
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(295, 203)
        Me.DateTimePicker1.Margin = New System.Windows.Forms.Padding(4)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(183, 22)
        Me.DateTimePicker1.TabIndex = 3
        Me.DateTimePicker1.Value = New Date(2018, 7, 7, 0, 0, 0, 0)
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(19, 203)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(173, 24)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Invoice Month End:"
        '
        'lblDB
        '
        Me.lblDB.AutoSize = True
        Me.lblDB.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDB.Location = New System.Drawing.Point(19, 9)
        Me.lblDB.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblDB.Name = "lblDB"
        Me.lblDB.Size = New System.Drawing.Size(169, 24)
        Me.lblDB.TabIndex = 4
        Me.lblDB.Text = "Database Location:"
        '
        'lblReports
        '
        Me.lblReports.AutoSize = True
        Me.lblReports.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReports.Location = New System.Drawing.Point(19, 43)
        Me.lblReports.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblReports.Name = "lblReports"
        Me.lblReports.Size = New System.Drawing.Size(240, 24)
        Me.lblReports.TabIndex = 4
        Me.lblReports.Text = "Reports Template Location:"
        '
        'txtReports
        '
        Me.txtReports.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReports.Location = New System.Drawing.Point(267, 43)
        Me.txtReports.Margin = New System.Windows.Forms.Padding(4)
        Me.txtReports.Name = "txtReports"
        Me.txtReports.Size = New System.Drawing.Size(173, 29)
        Me.txtReports.TabIndex = 0
        '
        'txtDB
        '
        Me.txtDB.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDB.Location = New System.Drawing.Point(196, 9)
        Me.txtDB.Margin = New System.Windows.Forms.Padding(4)
        Me.txtDB.Name = "txtDB"
        Me.txtDB.Size = New System.Drawing.Size(244, 29)
        Me.txtDB.TabIndex = 0
        '
        'btnDBBrowse
        '
        Me.btnDBBrowse.Location = New System.Drawing.Point(448, 9)
        Me.btnDBBrowse.Name = "btnDBBrowse"
        Me.btnDBBrowse.Size = New System.Drawing.Size(32, 29)
        Me.btnDBBrowse.TabIndex = 10
        Me.btnDBBrowse.Text = "..."
        Me.btnDBBrowse.UseVisualStyleBackColor = True
        '
        'btnReportBrowse
        '
        Me.btnReportBrowse.Location = New System.Drawing.Point(448, 43)
        Me.btnReportBrowse.Name = "btnReportBrowse"
        Me.btnReportBrowse.Size = New System.Drawing.Size(32, 29)
        Me.btnReportBrowse.TabIndex = 10
        Me.btnReportBrowse.Text = "..."
        Me.btnReportBrowse.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(492, 327)
        Me.Controls.Add(Me.btnReportBrowse)
        Me.Controls.Add(Me.btnDBBrowse)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.txtLabour)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.txtLargeStorage)
        Me.Controls.Add(Me.txtDB)
        Me.Controls.Add(Me.txtReports)
        Me.Controls.Add(Me.txtLargeInbound)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblReports)
        Me.Controls.Add(Me.lblDB)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "Form1"
        Me.Text = "AWS Storage Invoice Generator"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents txtLargeInbound As TextBox
    Friend WithEvents txtLargeStorage As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents txtLabour As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As ToolStripStatusLabel
    Friend WithEvents DateTimePicker1 As DateTimePicker
    Friend WithEvents Label4 As Label
    Friend WithEvents lblDB As Label
    Friend WithEvents lblReports As Label
    Friend WithEvents txtReports As TextBox
    Friend WithEvents txtDB As TextBox
    Friend WithEvents btnDBBrowse As Button
    Friend WithEvents btnReportBrowse As Button
End Class
