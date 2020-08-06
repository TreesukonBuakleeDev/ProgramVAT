<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmProcess
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents pgbProcess As System.Windows.Forms.ProgressBar
	Public WithEvents timProcess As System.Windows.Forms.Timer
	Public WithEvents lblProcess As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProcess))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.pgbProcess = New System.Windows.Forms.ProgressBar
        Me.timProcess = New System.Windows.Forms.Timer(Me.components)
        Me.lblProcess = New System.Windows.Forms.Label
        Me.Timmer = New System.Windows.Forms.Label
        Me.lblInsert = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'pgbProcess
        '
        Me.pgbProcess.Location = New System.Drawing.Point(12, 80)
        Me.pgbProcess.Name = "pgbProcess"
        Me.pgbProcess.Size = New System.Drawing.Size(298, 22)
        Me.pgbProcess.TabIndex = 1
        '
        'timProcess
        '
        Me.timProcess.Enabled = True
        Me.timProcess.Interval = 500
        '
        'lblProcess
        '
        Me.lblProcess.BackColor = System.Drawing.Color.Transparent
        Me.lblProcess.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblProcess.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProcess.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblProcess.Location = New System.Drawing.Point(12, 7)
        Me.lblProcess.Name = "lblProcess"
        Me.lblProcess.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblProcess.Size = New System.Drawing.Size(298, 40)
        Me.lblProcess.TabIndex = 0
        Me.lblProcess.Text = "Waiting"
        Me.lblProcess.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Timmer
        '
        Me.Timmer.AutoSize = True
        Me.Timmer.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Timmer.Location = New System.Drawing.Point(134, 120)
        Me.Timmer.Name = "Timmer"
        Me.Timmer.Size = New System.Drawing.Size(55, 15)
        Me.Timmer.TabIndex = 2
        Me.Timmer.Text = "00:00:00"
        Me.Timmer.Visible = False
        '
        'lblInsert
        '
        Me.lblInsert.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInsert.Location = New System.Drawing.Point(12, 52)
        Me.lblInsert.Name = "lblInsert"
        Me.lblInsert.Size = New System.Drawing.Size(301, 25)
        Me.lblInsert.TabIndex = 3
        '
        'frmProcess
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.AutoSize = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(327, 146)
        Me.Controls.Add(Me.lblInsert)
        Me.Controls.Add(Me.Timmer)
        Me.Controls.Add(Me.pgbProcess)
        Me.Controls.Add(Me.lblProcess)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmProcess"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Process"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents Timmer As System.Windows.Forms.Label
    Friend WithEvents lblInsert As System.Windows.Forms.Label
#End Region
End Class