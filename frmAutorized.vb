Option Strict Off
Option Explicit On
Friend Class frmAutorized
	Inherits System.Windows.Forms.Form
	Public ClickSelect As Boolean
	Public User As String
	Public Password As String
	Private Sub cmdClose_Click()
		Me.Close()
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		txtUser.Text = ""
		txtPass.Text = ""
	End Sub
	
	Private Sub cmdOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOpen.Click
		User = txtUser.Text
		Password = txtPass.Text
		ClickSelect = True
		Me.Close()
	End Sub
	
	Private Sub frmAutorized_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Text = "Autorized"
	End Sub
	
	Private Sub CheckChange()
		cmdOpen.Enabled = Len(txtUser.Text) > 0
	End Sub

    Private Sub txtPass_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPass.TextChanged
        CheckChange()
    End Sub
	
    Private Sub txtUser_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUser.TextChanged
        CheckChange()
    End Sub
	End Class