Imports System.Drawing.Drawing2D

Public Class FrmLogin
    'Dim OACC As New BassClass.BaseODBCIO(Application.StartupPath, "ACC")
    Dim cn As ADODB.Connection
    Dim rs As BaseClass.BaseODBCIO
    Dim Connstring As String
    Dim strsplit As Object
    

    Private Sub FrmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DrawBackground()
        AddHandler txtUser.KeyDown, AddressOf btnOK_Click
        AddHandler txtPassword.KeyDown, AddressOf btnOK_Click
        'Dim SQL As String = ""
        'cn = New ADODB.Connection
        'rs = New BaseClass.BaseODBCIO
        ''ConnVAT = GetPrivateProfile("COM" & ComID, "VAT", "")
        'cn.Open(ConnVAT)
        'Try
        '    SQL = "select * from UserPass"
        '    cn.Execute(SQL)
        '    'rs.Open(" select * from UserPass", cn)
        '    Exit Sub
        'Catch ex As Exception

        'End Try

        'SQL = " CREATE TABLE UserPass " + Environment.NewLine
        'SQL &= " ( UserID int, " + Environment.NewLine
        'SQL &= " UserName char(100), " + Environment.NewLine
        'SQL &= " PSW char(100), " + Environment.NewLine
        'SQL &= " Name varchar(200))" + Environment.NewLine
        'cn.Execute(SQL)
        'SQL = " insert into USERPASS  values(0,'ZJSIpmqbQDqJtjdsYIxiPw==','ZJSIpmqbQDqJtjdsYIxiPw==','Administrator') " + Environment.NewLine
        'cn.Execute(SQL)
    End Sub
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If CType(e, KeyEventArgs).KeyData <> 13 Then
                Exit Sub
            End If
        Catch ex As Exception

        End Try
        cn = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        'ConnVAT = GetPrivateProfile("COM" & ComID, "VAT", "")
        cn.Open(ConnVAT)
        Try
            Dim STR As String = "select * from USERPASS where USERNAME='" & BaseClass.AES_Encrypt(txtUser.Text, "ABCDEFG") & "' and PASS='" & BaseClass.AES_Encrypt(txtPassword.Text, "ABCDEFG") & "'"
            rs.Open(STR, cn)
            Do While rs.options.Rs.EOF = False
                Me.DialogResult = Windows.Forms.DialogResult.Yes
                UserLogin = rs.options.Rs.Collect("Name").ToString
                UserName = BaseClass.AES_Decrypt(rs.options.Rs.Collect("username").ToString, "ABCDEFG")
                Me.Close()
                Exit Sub
            Loop
            ShowToolTip(txtUser, ToolTipIcon.Warning, "UserName or Password Invalid")
        Catch ex As Exception
            ShowToolTip(txtUser, ToolTipIcon.Warning, "UserName or Password Invalid")
        End Try

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
