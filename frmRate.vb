Option Strict Off
Option Explicit On
Friend Class frmRate
	Inherits System.Windows.Forms.Form
	Public ClickSelect As Boolean
	Dim lbLoad As Boolean
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Hide()
	End Sub
	
	Private Sub cmdUpdate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUpdate.Click
        Call ProcessUpdateRate(CDbl(txtFrom.Text), CDbl(txtTo.Text), CDbl(txtRate.Text))
        DialogResult = Windows.Forms.DialogResult.Yes
		Me.Hide()
	End Sub
	
	Private Sub CheckUpdate()
		If txtFrom.Text <> "0" Or txtTo.Text <> "0" Or txtRate.Text <> "0" Then
			cmdUpdate.Enabled = True
		Else
			cmdUpdate.Enabled = False
		End If
	End Sub
	Private Sub frmRate_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		txtFrom.Text = "0"
		txtTo.Text = "0"
		txtRate.Text = "0"
		Me.Text = "Adjust record rate "
		CheckUpdate()
	End Sub
	
	Private Sub txtFrom_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFrom.Enter
        'HighlightBox(txtFrom)
    End Sub

    Private Sub txtFrom_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFrom.Leave
        If IsNumeric(txtFrom.Text) = False Then
            MsgBox("Wrong format Amount")
            txtFrom.Text = CStr(0)
            txtFrom.Focus()
        End If
        CheckUpdate()
    End Sub

    Private Sub txtRate_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRate.Enter
        'HighlightBox(txtRate)
    End Sub

    Private Sub txtRate_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRate.Leave
        If IsNumeric(txtRate.Text) = False Then
            MsgBox("Wrong format Amount")
            txtRate.Text = CStr(0)
            txtRate.Focus()
        End If
        CheckUpdate()
    End Sub

    Private Sub txtTo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTo.Enter
        'HighlightBox(txtTo)
    End Sub
	
	Private Sub txtTo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTo.Leave
		If IsNumeric(txtTo.Text) = False Then
			MsgBox("Wrong format Amount")
			txtTo.Text = CStr(0)
			txtTo.Focus()
		End If
		CheckUpdate()
	End Sub

End Class