Option Strict Off
Option Explicit On
Friend Class frmAbout
	Inherits System.Windows.Forms.Form
	
    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
	
	Private Sub frmAbout_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Me.Text = "About " & My.Application.Info.ProductName
        lblVersion.Text = My.Application.Info.ProductName & " " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build
		lblTitle.Text = My.Application.Info.ProductName
		lblCompany.Text = "Forward Management Services Co,.Ltd"
        lblDescription.Text = "- Get vat report from ACCPAC Module :" & vbCrLf & "    GL , AP , AR , PO , OE and Cashbook" & vbCrLf & "- EditTable vat list before print to report"
        lblDisclaimer.Text = "293 Tiam Ruam Mitr., Huaykwang, Bangkok 10320" & vbCrLf & "Tel:0-2274-4070-1,0-2274-4670  Fax:0-2274-4881" & vbCrLf & "Email:fms@fmsconsult.com  www.fmsconsult.com" & vbCrLf & Environment.NewLine & "ProgramPath:" & Environment.NewLine & Application.StartupPath
        'lblDisclaimer.Height = (lblDisclaimer.Text.Length * lblDisclaimer.Font.Size) / (lblDisclaimer.Width / lblDisclaimer.Font.Size) + 20
        'Me.Height = lblDisclaimer.Location.Y + lblDisclaimer.Height
        'cmdClose.Location = New Point(cmdClose.Location.X, Me.Height)
        'Me.Height += 10

	End Sub

    Private Sub lblVersion_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblVersion.Click

    End Sub
End Class