<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmAdjust
#Region "Windows Form Designer generated code "
    <System.Diagnostics.DebuggerNonUserCode()> Public Sub New(ByRef _FRM As frmMain, ByRef _CMD_TYPE As String)
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        FRM = _FRM
        CMD_TYPE = _CMD_TYPE
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
    Public WithEvents txtRef As System.Windows.Forms.TextBox
    Public WithEvents chkNone As System.Windows.Forms.CheckBox
    Public WithEvents txtComment As System.Windows.Forms.TextBox
    Public WithEvents txtRate As System.Windows.Forms.TextBox
    Public WithEvents cboType As System.Windows.Forms.ComboBox
    Public WithEvents cboLocation As System.Windows.Forms.ComboBox
    Public WithEvents txtName As System.Windows.Forms.TextBox
    Public WithEvents txtNewDoc As System.Windows.Forms.TextBox
    Public WithEvents txtDocNo As System.Windows.Forms.TextBox
    Public WithEvents txtInvNo As System.Windows.Forms.TextBox
    Public WithEvents btnCancel As BaseClass.MBGlassButton
    Public WithEvents btnSave As BaseClass.MBGlassButton
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_9 As System.Windows.Forms.Label
    Public WithEvents lblSource As System.Windows.Forms.Label
    Public WithEvents lblBatch As System.Windows.Forms.Label
    Public WithEvents _lblLabels_8 As System.Windows.Forms.Label
    Public WithEvents _Tax_3 As System.Windows.Forms.Label
    Public WithEvents _Tax_2 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_1 As System.Windows.Forms.Label
    Public WithEvents _Tax_1 As System.Windows.Forms.Label
    Public WithEvents _Tax_0 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_7 As System.Windows.Forms.Label
    Public WithEvents _Tax_6 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_5 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_4 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_3 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_2 As System.Windows.Forms.Label
    Public WithEvents _lblLabels_0 As System.Windows.Forms.Label
    Public WithEvents Tax As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Public WithEvents lblLabels As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAdjust))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.BT_Previous = New BaseClass.MBGlassButton
        Me.BT_Next = New BaseClass.MBGlassButton
        Me.txtRef = New System.Windows.Forms.TextBox
        Me.chkNone = New System.Windows.Forms.CheckBox
        Me.txtComment = New System.Windows.Forms.TextBox
        Me.txtRate = New System.Windows.Forms.TextBox
        Me.cboType = New System.Windows.Forms.ComboBox
        Me.cboLocation = New System.Windows.Forms.ComboBox
        Me.txtName = New System.Windows.Forms.TextBox
        Me.txtNewDoc = New System.Windows.Forms.TextBox
        Me.txtDocNo = New System.Windows.Forms.TextBox
        Me.txtInvNo = New System.Windows.Forms.TextBox
        Me.btnCancel = New BaseClass.MBGlassButton
        Me.btnSave = New BaseClass.MBGlassButton
        Me.Label1 = New System.Windows.Forms.Label
        Me._lblLabels_9 = New System.Windows.Forms.Label
        Me.lblSource = New System.Windows.Forms.Label
        Me.lblBatch = New System.Windows.Forms.Label
        Me._lblLabels_8 = New System.Windows.Forms.Label
        Me._Tax_3 = New System.Windows.Forms.Label
        Me._Tax_2 = New System.Windows.Forms.Label
        Me._lblLabels_1 = New System.Windows.Forms.Label
        Me._Tax_1 = New System.Windows.Forms.Label
        Me._Tax_0 = New System.Windows.Forms.Label
        Me._lblLabels_7 = New System.Windows.Forms.Label
        Me._Tax_6 = New System.Windows.Forms.Label
        Me._lblLabels_5 = New System.Windows.Forms.Label
        Me._lblLabels_4 = New System.Windows.Forms.Label
        Me._lblLabels_3 = New System.Windows.Forms.Label
        Me._lblLabels_2 = New System.Windows.Forms.Label
        Me._lblLabels_0 = New System.Windows.Forms.Label
        Me.Tax = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.lblLabels = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtRunno = New System.Windows.Forms.TextBox
        Me.TxtTrans = New System.Windows.Forms.DateTimePicker
        Me.txtDate = New System.Windows.Forms.DateTimePicker
        Me.Lb_TranNo = New System.Windows.Forms.Label
        Me.Tx_tranNo = New System.Windows.Forms.TextBox
        Me.Lb_Cif = New System.Windows.Forms.Label
        Me.Lb_taxCIF = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label15 = New System.Windows.Forms.Label
        Me.txtExchang = New System.Windows.Forms.NumericUpDown
        Me.txtSourceCurr = New System.Windows.Forms.NumericUpDown
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtCurr = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Cb_Verify = New System.Windows.Forms.CheckBox
        Me.txtAmount = New System.Windows.Forms.NumericUpDown
        Me.Tx_TaxCIF = New System.Windows.Forms.NumericUpDown
        Me.Tx_Cif = New System.Windows.Forms.NumericUpDown
        Me.txtTax = New System.Windows.Forms.NumericUpDown
        Me.TxBranch = New System.Windows.Forms.TextBox
        Me.TxTaxID = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        CType(Me.Tax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.txtExchang, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSourceCurr, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtAmount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Tx_TaxCIF, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Tx_Cif, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtTax, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'BT_Previous
        '
        Me.BT_Previous.Arrow = BaseClass.MBGlassButton.MB_Arrow.None
        Me.BT_Previous.BackColor = System.Drawing.Color.CornflowerBlue
        Me.BT_Previous.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.BT_Previous.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.BT_Previous.FlatAppearance.BorderSize = 0
        Me.BT_Previous.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.BT_Previous.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BT_Previous.GroupPosition = BaseClass.MBGlassButton.MB_GroupPos.None
        Me.BT_Previous.ImageSize = New System.Drawing.Size(24, 24)
        Me.BT_Previous.Location = New System.Drawing.Point(369, 12)
        Me.BT_Previous.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.BT_Previous.Name = "BT_Previous"
        Me.BT_Previous.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.BT_Previous.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.BT_Previous.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.BT_Previous.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.BT_Previous.Radius = 20
        Me.BT_Previous.ShowBase = BaseClass.MBGlassButton.MB_ShowBase.Yes
        Me.BT_Previous.Size = New System.Drawing.Size(50, 29)
        Me.BT_Previous.SplitButton = BaseClass.MBGlassButton.MB_SplitButton.No
        Me.BT_Previous.SplitLocation = BaseClass.MBGlassButton.MB_SplitLocation.Bottom
        Me.BT_Previous.TabIndex = 37
        Me.BT_Previous.Text = "<<"
        Me.ToolTip1.SetToolTip(Me.BT_Previous, "Previous")
        Me.BT_Previous.UseVisualStyleBackColor = False
        '
        'BT_Next
        '
        Me.BT_Next.Arrow = BaseClass.MBGlassButton.MB_Arrow.None
        Me.BT_Next.BackColor = System.Drawing.Color.Gainsboro
        Me.BT_Next.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.BT_Next.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.BT_Next.FlatAppearance.BorderSize = 0
        Me.BT_Next.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.BT_Next.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BT_Next.GroupPosition = BaseClass.MBGlassButton.MB_GroupPos.None
        Me.BT_Next.ImageSize = New System.Drawing.Size(24, 24)
        Me.BT_Next.Location = New System.Drawing.Point(425, 12)
        Me.BT_Next.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.BT_Next.Name = "BT_Next"
        Me.BT_Next.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.BT_Next.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.BT_Next.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.BT_Next.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.BT_Next.Radius = 20
        Me.BT_Next.ShowBase = BaseClass.MBGlassButton.MB_ShowBase.Yes
        Me.BT_Next.Size = New System.Drawing.Size(50, 29)
        Me.BT_Next.SplitButton = BaseClass.MBGlassButton.MB_SplitButton.No
        Me.BT_Next.SplitLocation = BaseClass.MBGlassButton.MB_SplitLocation.Bottom
        Me.BT_Next.TabIndex = 37
        Me.BT_Next.Text = ">>"
        Me.ToolTip1.SetToolTip(Me.BT_Next, "Next")
        Me.BT_Next.UseVisualStyleBackColor = False
        '
        'txtRef
        '
        Me.txtRef.AcceptsReturn = True
        Me.txtRef.BackColor = System.Drawing.SystemColors.Window
        Me.txtRef.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRef.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRef.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRef.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtRef.Location = New System.Drawing.Point(310, 119)
        Me.txtRef.MaxLength = 250
        Me.txtRef.Name = "txtRef"
        Me.txtRef.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRef.Size = New System.Drawing.Size(167, 21)
        Me.txtRef.TabIndex = 6
        '
        'chkNone
        '
        Me.chkNone.BackColor = System.Drawing.SystemColors.Control
        Me.chkNone.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkNone.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkNone.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkNone.Location = New System.Drawing.Point(250, 264)
        Me.chkNone.Name = "chkNone"
        Me.chkNone.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkNone.Size = New System.Drawing.Size(103, 19)
        Me.chkNone.TabIndex = 13
        Me.chkNone.Text = "None Rate"
        Me.chkNone.UseVisualStyleBackColor = False
        '
        'txtComment
        '
        Me.txtComment.AcceptsReturn = True
        Me.txtComment.BackColor = System.Drawing.Color.Snow
        Me.txtComment.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtComment.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtComment.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtComment.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtComment.Location = New System.Drawing.Point(112, 8)
        Me.txtComment.MaxLength = 1000
        Me.txtComment.Multiline = True
        Me.txtComment.Name = "txtComment"
        Me.txtComment.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtComment.Size = New System.Drawing.Size(365, 75)
        Me.txtComment.TabIndex = 0
        '
        'txtRate
        '
        Me.txtRate.AcceptsReturn = True
        Me.txtRate.BackColor = System.Drawing.SystemColors.Window
        Me.txtRate.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRate.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRate.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRate.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtRate.Location = New System.Drawing.Point(112, 261)
        Me.txtRate.MaxLength = 0
        Me.txtRate.Name = "txtRate"
        Me.txtRate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRate.Size = New System.Drawing.Size(129, 21)
        Me.txtRate.TabIndex = 12
        Me.txtRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cboType
        '
        Me.cboType.BackColor = System.Drawing.SystemColors.Window
        Me.cboType.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboType.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboType.Location = New System.Drawing.Point(348, 176)
        Me.cboType.Name = "cboType"
        Me.cboType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboType.Size = New System.Drawing.Size(129, 23)
        Me.cboType.TabIndex = 9
        '
        'cboLocation
        '
        Me.cboLocation.BackColor = System.Drawing.SystemColors.Window
        Me.cboLocation.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboLocation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLocation.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboLocation.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboLocation.Location = New System.Drawing.Point(112, 174)
        Me.cboLocation.Name = "cboLocation"
        Me.cboLocation.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboLocation.Size = New System.Drawing.Size(129, 23)
        Me.cboLocation.TabIndex = 8
        '
        'txtName
        '
        Me.txtName.AcceptsReturn = True
        Me.txtName.BackColor = System.Drawing.SystemColors.Window
        Me.txtName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtName.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtName.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtName.Location = New System.Drawing.Point(112, 146)
        Me.txtName.MaxLength = 200
        Me.txtName.Name = "txtName"
        Me.txtName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtName.Size = New System.Drawing.Size(365, 21)
        Me.txtName.TabIndex = 7
        '
        'txtNewDoc
        '
        Me.txtNewDoc.AcceptsReturn = True
        Me.txtNewDoc.BackColor = System.Drawing.SystemColors.Window
        Me.txtNewDoc.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNewDoc.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNewDoc.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNewDoc.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtNewDoc.Location = New System.Drawing.Point(112, 119)
        Me.txtNewDoc.MaxLength = 60
        Me.txtNewDoc.Name = "txtNewDoc"
        Me.txtNewDoc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNewDoc.Size = New System.Drawing.Size(129, 21)
        Me.txtNewDoc.TabIndex = 5
        '
        'txtDocNo
        '
        Me.txtDocNo.AcceptsReturn = True
        Me.txtDocNo.BackColor = System.Drawing.SystemColors.Window
        Me.txtDocNo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDocNo.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDocNo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDocNo.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtDocNo.Location = New System.Drawing.Point(112, 93)
        Me.txtDocNo.MaxLength = 255
        Me.txtDocNo.Name = "txtDocNo"
        Me.txtDocNo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDocNo.Size = New System.Drawing.Size(222, 21)
        Me.txtDocNo.TabIndex = 4
        '
        'txtInvNo
        '
        Me.txtInvNo.AcceptsReturn = True
        Me.txtInvNo.BackColor = System.Drawing.SystemColors.Window
        Me.txtInvNo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtInvNo.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtInvNo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInvNo.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtInvNo.Location = New System.Drawing.Point(112, 66)
        Me.txtInvNo.MaxLength = 255
        Me.txtInvNo.Name = "txtInvNo"
        Me.txtInvNo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtInvNo.Size = New System.Drawing.Size(222, 21)
        Me.txtInvNo.TabIndex = 3
        '
        'btnCancel
        '
        Me.btnCancel.Arrow = BaseClass.MBGlassButton.MB_Arrow.None
        Me.btnCancel.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnCancel.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnCancel.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.FlatAppearance.BorderSize = 0
        Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCancel.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCancel.GroupPosition = BaseClass.MBGlassButton.MB_GroupPos.None
        Me.btnCancel.Image = CType(resources.GetObject("btnCancel.Image"), System.Drawing.Image)
        Me.btnCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCancel.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnCancel.Location = New System.Drawing.Point(386, 111)
        Me.btnCancel.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnCancel.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnCancel.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnCancel.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnCancel.ShowBase = BaseClass.MBGlassButton.MB_ShowBase.Yes
        Me.btnCancel.Size = New System.Drawing.Size(87, 26)
        Me.btnCancel.SplitButton = BaseClass.MBGlassButton.MB_SplitButton.No
        Me.btnCancel.SplitLocation = BaseClass.MBGlassButton.MB_SplitLocation.Bottom
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "      Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'btnSave
        '
        Me.btnSave.Arrow = BaseClass.MBGlassButton.MB_Arrow.None
        Me.btnSave.BackColor = System.Drawing.SystemColors.Control
        Me.btnSave.BaseColor = System.Drawing.Color.FromArgb(CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer), CType(CType(211, Byte), Integer))
        Me.btnSave.BaseStrokeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnSave.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnSave.FlatAppearance.BorderSize = 0
        Me.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSave.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnSave.GroupPosition = BaseClass.MBGlassButton.MB_GroupPos.None
        Me.btnSave.Image = Global.ProgramVAT.My.Resources.Resources.Save
        Me.btnSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnSave.ImageSize = New System.Drawing.Size(24, 24)
        Me.btnSave.Location = New System.Drawing.Point(12, 111)
        Me.btnSave.MenuListPosition = New System.Drawing.Point(0, 0)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.OnColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(214, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.btnSave.OnStrokeColor = System.Drawing.Color.FromArgb(CType(CType(196, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(118, Byte), Integer))
        Me.btnSave.PressColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnSave.PressStrokeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnSave.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnSave.ShowBase = BaseClass.MBGlassButton.MB_ShowBase.Yes
        Me.btnSave.Size = New System.Drawing.Size(76, 26)
        Me.btnSave.SplitButton = BaseClass.MBGlassButton.MB_SplitButton.No
        Me.btnSave.SplitLocation = BaseClass.MBGlassButton.MB_SplitLocation.Bottom
        Me.btnSave.TabIndex = 1
        Me.btnSave.Text = "       Save"
        Me.btnSave.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(112, 117)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(260, 17)
        Me.Label1.TabIndex = 32
        Me.Label1.Text = "Detail No"
        Me.Label1.Visible = False
        '
        '_lblLabels_9
        '
        Me._lblLabels_9.BackColor = System.Drawing.Color.Transparent
        Me._lblLabels_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_9.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_9, CType(9, Short))
        Me._lblLabels_9.Location = New System.Drawing.Point(247, 121)
        Me._lblLabels_9.Name = "_lblLabels_9"
        Me._lblLabels_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_9.Size = New System.Drawing.Size(57, 18)
        Me._lblLabels_9.TabIndex = 31
        Me._lblLabels_9.Text = "Ref No."
        '
        'lblSource
        '
        Me.lblSource.BackColor = System.Drawing.Color.Transparent
        Me.lblSource.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSource.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSource.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSource.Location = New System.Drawing.Point(112, 90)
        Me.lblSource.Name = "lblSource"
        Me.lblSource.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSource.Size = New System.Drawing.Size(87, 18)
        Me.lblSource.TabIndex = 1
        Me.lblSource.Text = "Source"
        '
        'lblBatch
        '
        Me.lblBatch.BackColor = System.Drawing.Color.Transparent
        Me.lblBatch.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblBatch.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBatch.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblBatch.Location = New System.Drawing.Point(193, 90)
        Me.lblBatch.Name = "lblBatch"
        Me.lblBatch.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblBatch.Size = New System.Drawing.Size(280, 18)
        Me.lblBatch.TabIndex = 29
        Me.lblBatch.Text = "Batch"
        '
        '_lblLabels_8
        '
        Me._lblLabels_8.BackColor = System.Drawing.Color.Transparent
        Me._lblLabels_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_8.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_8, CType(8, Short))
        Me._lblLabels_8.Location = New System.Drawing.Point(12, 10)
        Me._lblLabels_8.Name = "_lblLabels_8"
        Me._lblLabels_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_8.Size = New System.Drawing.Size(72, 18)
        Me._lblLabels_8.TabIndex = 3
        Me._lblLabels_8.Text = "Comment"
        '
        '_Tax_3
        '
        Me._Tax_3.BackColor = System.Drawing.Color.Transparent
        Me._Tax_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Tax_3.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Tax_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Tax.SetIndex(Me._Tax_3, CType(3, Short))
        Me._Tax_3.Location = New System.Drawing.Point(12, 263)
        Me._Tax_3.Name = "_Tax_3"
        Me._Tax_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Tax_3.Size = New System.Drawing.Size(81, 18)
        Me._Tax_3.TabIndex = 27
        Me._Tax_3.Text = "Rate"
        '
        '_Tax_2
        '
        Me._Tax_2.BackColor = System.Drawing.Color.Transparent
        Me._Tax_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Tax_2.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Tax_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Tax.SetIndex(Me._Tax_2, CType(2, Short))
        Me._Tax_2.Location = New System.Drawing.Point(247, 176)
        Me._Tax_2.Name = "_Tax_2"
        Me._Tax_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Tax_2.Size = New System.Drawing.Size(84, 18)
        Me._Tax_2.TabIndex = 26
        Me._Tax_2.Text = "Vat Type"
        '
        '_lblLabels_1
        '
        Me._lblLabels_1.BackColor = System.Drawing.Color.Transparent
        Me._lblLabels_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_1, CType(1, Short))
        Me._lblLabels_1.Location = New System.Drawing.Point(12, 177)
        Me._lblLabels_1.Name = "_lblLabels_1"
        Me._lblLabels_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_1.Size = New System.Drawing.Size(72, 18)
        Me._lblLabels_1.TabIndex = 25
        Me._lblLabels_1.Text = "Location"
        '
        '_Tax_1
        '
        Me._Tax_1.BackColor = System.Drawing.Color.Transparent
        Me._Tax_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Tax_1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Tax_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Tax.SetIndex(Me._Tax_1, CType(1, Short))
        Me._Tax_1.Location = New System.Drawing.Point(247, 236)
        Me._Tax_1.Name = "_Tax_1"
        Me._Tax_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Tax_1.Size = New System.Drawing.Size(72, 18)
        Me._Tax_1.TabIndex = 24
        Me._Tax_1.Text = "Baht"
        '
        '_Tax_0
        '
        Me._Tax_0.BackColor = System.Drawing.Color.Transparent
        Me._Tax_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Tax_0.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Tax_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Tax.SetIndex(Me._Tax_0, CType(0, Short))
        Me._Tax_0.Location = New System.Drawing.Point(247, 204)
        Me._Tax_0.Name = "_Tax_0"
        Me._Tax_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Tax_0.Size = New System.Drawing.Size(72, 18)
        Me._Tax_0.TabIndex = 23
        Me._Tax_0.Text = "Baht"
        '
        '_lblLabels_7
        '
        Me._lblLabels_7.BackColor = System.Drawing.Color.Transparent
        Me._lblLabels_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_7.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_7, CType(7, Short))
        Me._lblLabels_7.Location = New System.Drawing.Point(12, 205)
        Me._lblLabels_7.Name = "_lblLabels_7"
        Me._lblLabels_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_7.Size = New System.Drawing.Size(76, 18)
        Me._lblLabels_7.TabIndex = 21
        Me._lblLabels_7.Text = "Amount"
        '
        '_Tax_6
        '
        Me._Tax_6.BackColor = System.Drawing.Color.Transparent
        Me._Tax_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._Tax_6.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Tax_6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Tax.SetIndex(Me._Tax_6, CType(6, Short))
        Me._Tax_6.Location = New System.Drawing.Point(12, 235)
        Me._Tax_6.Name = "_Tax_6"
        Me._Tax_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Tax_6.Size = New System.Drawing.Size(97, 22)
        Me._Tax_6.TabIndex = 20
        Me._Tax_6.Text = "Tax Amount"
        '
        '_lblLabels_5
        '
        Me._lblLabels_5.BackColor = System.Drawing.Color.Transparent
        Me._lblLabels_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_5.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_5, CType(5, Short))
        Me._lblLabels_5.Location = New System.Drawing.Point(12, 148)
        Me._lblLabels_5.Name = "_lblLabels_5"
        Me._lblLabels_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_5.Size = New System.Drawing.Size(94, 18)
        Me._lblLabels_5.TabIndex = 19
        Me._lblLabels_5.Text = "Name"
        '
        '_lblLabels_4
        '
        Me._lblLabels_4.BackColor = System.Drawing.Color.Transparent
        Me._lblLabels_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_4.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_4, CType(4, Short))
        Me._lblLabels_4.Location = New System.Drawing.Point(12, 121)
        Me._lblLabels_4.Name = "_lblLabels_4"
        Me._lblLabels_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_4.Size = New System.Drawing.Size(131, 22)
        Me._lblLabels_4.TabIndex = 18
        Me._lblLabels_4.Text = "Running No."
        '
        '_lblLabels_3
        '
        Me._lblLabels_3.BackColor = System.Drawing.Color.Transparent
        Me._lblLabels_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_3.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_3, CType(3, Short))
        Me._lblLabels_3.Location = New System.Drawing.Point(12, 95)
        Me._lblLabels_3.Name = "_lblLabels_3"
        Me._lblLabels_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_3.Size = New System.Drawing.Size(72, 18)
        Me._lblLabels_3.TabIndex = 17
        Me._lblLabels_3.Text = "Doc No"
        '
        '_lblLabels_2
        '
        Me._lblLabels_2.BackColor = System.Drawing.Color.Transparent
        Me._lblLabels_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_2.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_2, CType(2, Short))
        Me._lblLabels_2.Location = New System.Drawing.Point(12, 68)
        Me._lblLabels_2.Name = "_lblLabels_2"
        Me._lblLabels_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_2.Size = New System.Drawing.Size(106, 18)
        Me._lblLabels_2.TabIndex = 16
        Me._lblLabels_2.Text = "Tax Inv No."
        '
        '_lblLabels_0
        '
        Me._lblLabels_0.BackColor = System.Drawing.Color.Transparent
        Me._lblLabels_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._lblLabels_0.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._lblLabels_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLabels.SetIndex(Me._lblLabels_0, CType(0, Short))
        Me._lblLabels_0.Location = New System.Drawing.Point(12, 39)
        Me._lblLabels_0.Name = "_lblLabels_0"
        Me._lblLabels_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._lblLabels_0.Size = New System.Drawing.Size(94, 18)
        Me._lblLabels_0.TabIndex = 14
        Me._lblLabels_0.Text = "Post Date"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(12, 11)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(97, 18)
        Me.Label3.TabIndex = 34
        Me.Label3.Text = "Doc Date"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(9, 368)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(49, 15)
        Me.Label4.TabIndex = 36
        Me.Label4.Text = "RunNo."
        '
        'txtRunno
        '
        Me.txtRunno.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRunno.Location = New System.Drawing.Point(112, 368)
        Me.txtRunno.MaxLength = 20
        Me.txtRunno.Name = "txtRunno"
        Me.txtRunno.Size = New System.Drawing.Size(129, 21)
        Me.txtRunno.TabIndex = 16
        '
        'TxtTrans
        '
        Me.TxtTrans.CustomFormat = "dd/MM/yyyy"
        Me.TxtTrans.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtTrans.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.TxtTrans.Location = New System.Drawing.Point(112, 9)
        Me.TxtTrans.Name = "TxtTrans"
        Me.TxtTrans.Size = New System.Drawing.Size(129, 21)
        Me.TxtTrans.TabIndex = 0
        '
        'txtDate
        '
        Me.txtDate.CustomFormat = "dd/MM/yyyy"
        Me.txtDate.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtDate.Location = New System.Drawing.Point(112, 37)
        Me.txtDate.Name = "txtDate"
        Me.txtDate.Size = New System.Drawing.Size(129, 21)
        Me.txtDate.TabIndex = 1
        '
        'Lb_TranNo
        '
        Me.Lb_TranNo.AutoSize = True
        Me.Lb_TranNo.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lb_TranNo.Location = New System.Drawing.Point(245, 368)
        Me.Lb_TranNo.Name = "Lb_TranNo"
        Me.Lb_TranNo.Size = New System.Drawing.Size(85, 15)
        Me.Lb_TranNo.TabIndex = 36
        Me.Lb_TranNo.Text = "Transport No. "
        '
        'Tx_tranNo
        '
        Me.Tx_tranNo.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Tx_tranNo.Location = New System.Drawing.Point(348, 368)
        Me.Tx_tranNo.MaxLength = 20
        Me.Tx_tranNo.Name = "Tx_tranNo"
        Me.Tx_tranNo.Size = New System.Drawing.Size(129, 21)
        Me.Tx_tranNo.TabIndex = 17
        '
        'Lb_Cif
        '
        Me.Lb_Cif.AutoSize = True
        Me.Lb_Cif.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lb_Cif.Location = New System.Drawing.Point(247, 396)
        Me.Lb_Cif.Name = "Lb_Cif"
        Me.Lb_Cif.Size = New System.Drawing.Size(29, 15)
        Me.Lb_Cif.TabIndex = 36
        Me.Lb_Cif.Text = "CIF "
        '
        'Lb_taxCIF
        '
        Me.Lb_taxCIF.AutoSize = True
        Me.Lb_taxCIF.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lb_taxCIF.Location = New System.Drawing.Point(247, 425)
        Me.Lb_taxCIF.Name = "Lb_taxCIF"
        Me.Lb_taxCIF.Size = New System.Drawing.Size(50, 15)
        Me.Lb_taxCIF.TabIndex = 36
        Me.Lb_taxCIF.Text = "Tax CIF "
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel1.Controls.Add(Me.Label15)
        Me.Panel1.Controls.Add(Me.txtExchang)
        Me.Panel1.Controls.Add(Me.txtSourceCurr)
        Me.Panel1.Controls.Add(Me.Label14)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.txtCurr)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Cb_Verify)
        Me.Panel1.Controls.Add(Me.BT_Next)
        Me.Panel1.Controls.Add(Me.BT_Previous)
        Me.Panel1.Controls.Add(Me.txtAmount)
        Me.Panel1.Controls.Add(Me.Tx_TaxCIF)
        Me.Panel1.Controls.Add(Me.Tx_Cif)
        Me.Panel1.Controls.Add(Me.txtTax)
        Me.Panel1.Controls.Add(Me.cboType)
        Me.Panel1.Controls.Add(Me.txtInvNo)
        Me.Panel1.Controls.Add(Me.txtNewDoc)
        Me.Panel1.Controls.Add(Me.txtDate)
        Me.Panel1.Controls.Add(Me.TxtTrans)
        Me.Panel1.Controls.Add(Me.Tx_tranNo)
        Me.Panel1.Controls.Add(Me.txtRunno)
        Me.Panel1.Controls.Add(Me.txtRate)
        Me.Panel1.Controls.Add(Me.txtDocNo)
        Me.Panel1.Controls.Add(Me.cboLocation)
        Me.Panel1.Controls.Add(Me.TxBranch)
        Me.Panel1.Controls.Add(Me.TxTaxID)
        Me.Panel1.Controls.Add(Me.txtName)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me._lblLabels_0)
        Me.Panel1.Controls.Add(Me._lblLabels_2)
        Me.Panel1.Controls.Add(Me._lblLabels_3)
        Me.Panel1.Controls.Add(Me._lblLabels_4)
        Me.Panel1.Controls.Add(Me.Lb_taxCIF)
        Me.Panel1.Controls.Add(Me._lblLabels_5)
        Me.Panel1.Controls.Add(Me._Tax_6)
        Me.Panel1.Controls.Add(Me.Lb_Cif)
        Me.Panel1.Controls.Add(Me._lblLabels_7)
        Me.Panel1.Controls.Add(Me.Lb_TranNo)
        Me.Panel1.Controls.Add(Me._Tax_0)
        Me.Panel1.Controls.Add(Me._Tax_1)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me._lblLabels_1)
        Me.Panel1.Controls.Add(Me._Tax_2)
        Me.Panel1.Controls.Add(Me._Tax_3)
        Me.Panel1.Controls.Add(Me.txtRef)
        Me.Panel1.Controls.Add(Me._lblLabels_9)
        Me.Panel1.Controls.Add(Me.chkNone)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(491, 460)
        Me.Panel1.TabIndex = 0
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.Red
        Me.Label15.Location = New System.Drawing.Point(293, 323)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(182, 15)
        Me.Label15.TabIndex = 45
        Me.Label15.Text = "*กรุณาตรวจสอบ Ex.Rate ก่อน Save"
        '
        'txtExchang
        '
        Me.txtExchang.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtExchang.DecimalPlaces = 7
        Me.txtExchang.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtExchang.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtExchang.Location = New System.Drawing.Point(348, 340)
        Me.txtExchang.Maximum = New Decimal(New Integer() {-1530494977, 232830, 0, 0})
        Me.txtExchang.Minimum = New Decimal(New Integer() {-1530494977, 232830, 0, -2147483648})
        Me.txtExchang.Name = "txtExchang"
        Me.txtExchang.Size = New System.Drawing.Size(129, 21)
        Me.txtExchang.TabIndex = 15
        Me.txtExchang.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtExchang.ThousandsSeparator = True
        '
        'txtSourceCurr
        '
        Me.txtSourceCurr.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSourceCurr.DecimalPlaces = 3
        Me.txtSourceCurr.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSourceCurr.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtSourceCurr.Location = New System.Drawing.Point(112, 340)
        Me.txtSourceCurr.Maximum = New Decimal(New Integer() {-1530494977, 232830, 0, 0})
        Me.txtSourceCurr.Minimum = New Decimal(New Integer() {-1530494977, 232830, 0, -2147483648})
        Me.txtSourceCurr.Name = "txtSourceCurr"
        Me.txtSourceCurr.Size = New System.Drawing.Size(129, 21)
        Me.txtSourceCurr.TabIndex = 14
        Me.txtSourceCurr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtSourceCurr.ThousandsSeparator = True
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(9, 289)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(107, 15)
        Me.Label14.TabIndex = 44
        Me.Label14.Text = "(สกุลเงินต่างประเทศ)"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(247, 344)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(90, 15)
        Me.Label13.TabIndex = 43
        Me.Label13.Text = "Exchange Rate"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.Transparent
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label12.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label12.Location = New System.Drawing.Point(9, 314)
        Me.Label12.Name = "Label12"
        Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label12.Size = New System.Drawing.Size(87, 18)
        Me.Label12.TabIndex = 42
        Me.Label12.Text = "Currency"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.Transparent
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Location = New System.Drawing.Point(9, 345)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(100, 18)
        Me.Label11.TabIndex = 42
        Me.Label11.Text = "Source Currency"
        '
        'txtCurr
        '
        Me.txtCurr.AcceptsReturn = True
        Me.txtCurr.BackColor = System.Drawing.SystemColors.Window
        Me.txtCurr.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCurr.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCurr.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCurr.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtCurr.Location = New System.Drawing.Point(112, 310)
        Me.txtCurr.MaxLength = 13
        Me.txtCurr.Name = "txtCurr"
        Me.txtCurr.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCurr.Size = New System.Drawing.Size(129, 21)
        Me.txtCurr.TabIndex = 13
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.Transparent
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(9, 425)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(90, 22)
        Me.Label10.TabIndex = 40
        Me.Label10.Text = "Branch"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(9, 396)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(65, 22)
        Me.Label9.TabIndex = 40
        Me.Label9.Text = "Tax ID"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Red
        Me.Label8.Location = New System.Drawing.Point(318, 178)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(13, 16)
        Me.Label8.TabIndex = 39
        Me.Label8.Text = "*"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.BackColor = System.Drawing.Color.Transparent
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Red
        Me.Label7.Location = New System.Drawing.Point(94, 149)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(13, 16)
        Me.Label7.TabIndex = 39
        Me.Label7.Text = "*"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Red
        Me.Label6.Location = New System.Drawing.Point(94, 69)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(13, 16)
        Me.Label6.TabIndex = 39
        Me.Label6.Text = "*"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Red
        Me.Label5.Location = New System.Drawing.Point(94, 40)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(13, 16)
        Me.Label5.TabIndex = 39
        Me.Label5.Text = "*"
        '
        'Cb_Verify
        '
        Me.Cb_Verify.AutoSize = True
        Me.Cb_Verify.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cb_Verify.Location = New System.Drawing.Point(277, 9)
        Me.Cb_Verify.Name = "Cb_Verify"
        Me.Cb_Verify.Size = New System.Drawing.Size(77, 19)
        Me.Cb_Verify.TabIndex = 2
        Me.Cb_Verify.Text = "Approved"
        Me.Cb_Verify.UseVisualStyleBackColor = True
        '
        'txtAmount
        '
        Me.txtAmount.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtAmount.DecimalPlaces = 3
        Me.txtAmount.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAmount.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtAmount.Location = New System.Drawing.Point(112, 205)
        Me.txtAmount.Maximum = New Decimal(New Integer() {-1530494977, 232830, 0, 0})
        Me.txtAmount.Minimum = New Decimal(New Integer() {-1530494977, 232830, 0, -2147483648})
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.Size = New System.Drawing.Size(129, 21)
        Me.txtAmount.TabIndex = 10
        Me.txtAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtAmount.ThousandsSeparator = True
        '
        'Tx_TaxCIF
        '
        Me.Tx_TaxCIF.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Tx_TaxCIF.DecimalPlaces = 3
        Me.Tx_TaxCIF.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Tx_TaxCIF.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.Tx_TaxCIF.Location = New System.Drawing.Point(348, 424)
        Me.Tx_TaxCIF.Maximum = New Decimal(New Integer() {-1530494977, 232830, 0, 0})
        Me.Tx_TaxCIF.Minimum = New Decimal(New Integer() {-1530494977, 232830, 0, -2147483648})
        Me.Tx_TaxCIF.Name = "Tx_TaxCIF"
        Me.Tx_TaxCIF.Size = New System.Drawing.Size(129, 21)
        Me.Tx_TaxCIF.TabIndex = 21
        Me.Tx_TaxCIF.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.Tx_TaxCIF.ThousandsSeparator = True
        '
        'Tx_Cif
        '
        Me.Tx_Cif.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Tx_Cif.DecimalPlaces = 3
        Me.Tx_Cif.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Tx_Cif.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.Tx_Cif.Location = New System.Drawing.Point(348, 396)
        Me.Tx_Cif.Maximum = New Decimal(New Integer() {-1530494977, 232830, 0, 0})
        Me.Tx_Cif.Minimum = New Decimal(New Integer() {-1530494977, 232830, 0, -2147483648})
        Me.Tx_Cif.Name = "Tx_Cif"
        Me.Tx_Cif.Size = New System.Drawing.Size(129, 21)
        Me.Tx_Cif.TabIndex = 20
        Me.Tx_Cif.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.Tx_Cif.ThousandsSeparator = True
        '
        'txtTax
        '
        Me.txtTax.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTax.DecimalPlaces = 3
        Me.txtTax.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTax.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.txtTax.Location = New System.Drawing.Point(112, 233)
        Me.txtTax.Maximum = New Decimal(New Integer() {-1530494977, 232830, 0, 0})
        Me.txtTax.Minimum = New Decimal(New Integer() {-1530494977, 232830, 0, -2147483648})
        Me.txtTax.Name = "txtTax"
        Me.txtTax.Size = New System.Drawing.Size(129, 21)
        Me.txtTax.TabIndex = 11
        Me.txtTax.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtTax.ThousandsSeparator = True
        '
        'TxBranch
        '
        Me.TxBranch.AcceptsReturn = True
        Me.TxBranch.BackColor = System.Drawing.SystemColors.Window
        Me.TxBranch.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxBranch.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxBranch.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxBranch.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TxBranch.Location = New System.Drawing.Point(112, 423)
        Me.TxBranch.MaxLength = 200
        Me.TxBranch.Name = "TxBranch"
        Me.TxBranch.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxBranch.Size = New System.Drawing.Size(130, 21)
        Me.TxBranch.TabIndex = 19
        '
        'TxTaxID
        '
        Me.TxTaxID.AcceptsReturn = True
        Me.TxTaxID.BackColor = System.Drawing.SystemColors.Window
        Me.TxTaxID.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TxTaxID.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxTaxID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TxTaxID.ImeMode = System.Windows.Forms.ImeMode.Disable
        Me.TxTaxID.Location = New System.Drawing.Point(112, 394)
        Me.TxTaxID.MaxLength = 13
        Me.TxTaxID.Name = "TxTaxID"
        Me.TxTaxID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TxTaxID.Size = New System.Drawing.Size(129, 21)
        Me.TxTaxID.TabIndex = 18
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Red
        Me.Label2.Location = New System.Drawing.Point(94, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(13, 16)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "*"
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.Panel2.Controls.Add(Me._lblLabels_8)
        Me.Panel2.Controls.Add(Me.lblBatch)
        Me.Panel2.Controls.Add(Me.lblSource)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.btnSave)
        Me.Panel2.Controls.Add(Me.btnCancel)
        Me.Panel2.Controls.Add(Me.txtComment)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 460)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(491, 162)
        Me.Panel2.TabIndex = 0
        '
        'frmAdjust
        '
        Me.AcceptButton = Me.btnSave
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.AutoSize = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(244, Byte), Integer), CType(CType(247, Byte), Integer), CType(CType(247, Byte), Integer))
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(491, 622)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAdjust"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Adj"
        CType(Me.Tax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.txtExchang, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSourceCurr, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtAmount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Tx_TaxCIF, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Tx_Cif, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtTax, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtRunno As System.Windows.Forms.TextBox
    Friend WithEvents TxtTrans As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Lb_TranNo As System.Windows.Forms.Label
    Friend WithEvents Tx_tranNo As System.Windows.Forms.TextBox
    Friend WithEvents Lb_Cif As System.Windows.Forms.Label
    Friend WithEvents Lb_taxCIF As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Public WithEvents txtTax As System.Windows.Forms.NumericUpDown
    Public WithEvents Tx_TaxCIF As System.Windows.Forms.NumericUpDown
    Public WithEvents Tx_Cif As System.Windows.Forms.NumericUpDown
    Public WithEvents txtAmount As System.Windows.Forms.NumericUpDown
    Friend WithEvents BT_Next As BaseClass.MBGlassButton
    Friend WithEvents BT_Previous As BaseClass.MBGlassButton
    Friend WithEvents Cb_Verify As System.Windows.Forms.CheckBox
    Public WithEvents Label8 As System.Windows.Forms.Label
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Label6 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label10 As System.Windows.Forms.Label
    Public WithEvents Label9 As System.Windows.Forms.Label
    Public WithEvents TxBranch As System.Windows.Forms.TextBox
    Public WithEvents TxTaxID As System.Windows.Forms.TextBox
    Public WithEvents txtCurr As System.Windows.Forms.TextBox
    Public WithEvents Label12 As System.Windows.Forms.Label
    Public WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Public WithEvents txtSourceCurr As System.Windows.Forms.NumericUpDown
    Public WithEvents txtExchang As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label15 As System.Windows.Forms.Label
#End Region
End Class