Option Strict Off
Option Explicit On
Friend Class frmCompany
	Inherits System.Windows.Forms.Form
	Public ClickSelect As Boolean
	Public CONAME As String
	Public ADDR01 As String
	Public ADDR02 As String
	Public ADDR03 As String
	Public ADDR04 As String
	Public CITY As String
	Public STATE As String
	Public POSTAL As String
	Public COUNTRY As String
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
		TAXNO = txtTax.Text
		ClickSelect = True
		Me.Close()
	End Sub
	
	Private Sub frmCompany_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        lblAddress(0).Text = Trim(CONAME)
        lblAddress(1).Text = Trim(ADDR01)
        lblAddress(2).Text = Trim(ADDR02)
        lblAddress(3).Text = Trim(ADDR03)
        lblAddress(4).Text = Trim(ADDR04)
        lblAddress(5).Text = Trim(CITY) & "  " & Trim(STATE)
        lblAddress(6).Text = Trim(COUNTRY) & " " & Trim(POSTAL)
        'txtTax.Text = Trim(TAXNO)
        lblTax.Text = "TAX NUMBER " & Trim(TAXNO)
		cmdSave.Enabled = False
		ClickSelect = False
        Me.Text = My.Application.Info.ProductName & " " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build
    End Sub
	
    Private Sub txtTax_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTax.TextChanged
        cmdSave.Enabled = Len(txtTax.Text) > 0
    End Sub
	End Class