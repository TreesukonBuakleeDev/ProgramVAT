Imports System.Drawing.Drawing2D

Public Class FrmLoginFMS
    'Dim OACC As New BassClass.BaseODBCIO(Application.StartupPath, "ACC")
    Dim cn As ADODB.Connection
    Dim rs As BaseClass.BaseODBCIO
    Dim Connstring As String
    Dim strsplit As Object
    Dim STCMD As String

    Public Sub New(ByVal _STCMD As String)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        STCMD = _STCMD
    End Sub
    Private Sub FrmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'DrawBackground()
        AddHandler txtUser.KeyDown, AddressOf btnOK_Click
        AddHandler txtPassword.KeyDown, AddressOf btnOK_Click
      
    End Sub
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click

        Dim frm As FrmConfigDB
        'If STCMD = "New" Then
        '    frm = New FrmConfigDB("New")
        'Else
        frm = New FrmConfigDB("")
        'End If


        Try
            If CType(e, KeyEventArgs).KeyData <> 13 Then
                Exit Sub
            End If
        Catch ex As Exception

        End Try

        Try

            If (txtUser.Text.Trim.ToUpper <> "FMSADMIN" Or txtPassword.Text.Trim.ToUpper <> "FMSADMIN") Then

                ShowToolTip(txtUser, ToolTipIcon.Warning, "UserName or Password Invalid")

            Else

                fmsadmin = txtUser.Text.Trim.ToUpper
                'If STCMD = "New" Then
                '    frmConfigNew.Close()
                'Else
                '    frmConfig.Close()
                'End If
                frmConfig.Close()

                frm.Show(frmMain)


                Me.Close()
                'Exit Sub
            End If

        Catch ex As Exception
            ShowToolTip(txtUser, ToolTipIcon.Warning, "UserName or Password Invalid")
        End Try


        'Try
        '    Dim STR As String = "select * from USERPASS where USERNAME='" & BaseClass.AES_Encrypt(txtUser.Text, "ABCDEFG") & "' and PASS='" & BaseClass.AES_Encrypt(txtPassword.Text, "ABCDEFG") & "'"
        '    rs.Open(STR, cn)
        '    Do While rs.options.Rs.EOF = False
        '        Me.DialogResult = Windows.Forms.DialogResult.Yes
        '        UserLogin = rs.options.Rs.Collect("NAME").ToString
        '        UserName = BaseClass.AES_Decrypt(rs.options.Rs.Collect("USERNAME").ToString, "ABCDEFG")
        '        Me.Close()
        '        Exit Sub
        '    Loop
        '    ShowToolTip(txtUser, ToolTipIcon.Warning, "UserName or Password Invalid")
        'Catch ex As Exception
        '    ShowToolTip(txtUser, ToolTipIcon.Warning, "UserName or Password Invalid")
        'End Try

    End Sub

    Private Sub DrawBackground()
        Dim lbBack As New LinearGradientBrush(Me.ClientRectangle, Color.Gray, Color.White, LinearGradientMode.Vertical)
        Dim bm As New Bitmap(Me.Width, Me.Height)
        Dim g As Graphics = Graphics.FromImage(bm)
        g.FillRectangle(lbBack, 0, 0, Me.ClientRectangle.Width, Me.ClientRectangle.Height)
        Me.BackgroundImage = bm
        lbBack.Dispose()
    End Sub
End Class
