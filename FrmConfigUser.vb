Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class FrmConfigUser
    'Dim h As New BaseFILEIO(Application.StartupPath)
    Dim Index As Integer = 0
    Dim DT As New DataTable
    Dim User, pass, UserID As String
    Dim cn As ADODB.Connection
    Dim rs As New BaseClass.BaseODBCIO

    Private Sub frmConfigUser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim SQL As String
        'SQL = " if (select COUNT(*) from sys.tables where name ='UserPass')=0  " & Environment.NewLine
        'SQL &= " begin " & Environment.NewLine
        'SQL &= " 	CREATE TABLE UserPass " & Environment.NewLine
        'SQL &= " 	(   UserID int,  " & Environment.NewLine
        'SQL &= " 	    UserName char(100),  " & Environment.NewLine
        'SQL &= " 	    PSW char(100), " & Environment.NewLine
        'SQL &= " 	    Name varchar(200)) " & Environment.NewLine
        'SQL = " insert into USERPASS  values(0,'ZJSIpmqbQDqJtjdsYIxiPw==','ZJSIpmqbQDqJtjdsYIxiPw==','Administrator')  " & Environment.NewLine
        'SQL &= "  end " & Environment.NewLine
        'cn = New ADODB.Connection
        'cn.Open(ConnVAT)
        'cn.Execute(SQL)

        SQL = "select * from USERPASS"
        cn = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        cn.Open(ConnVAT)
        rs.Open(SQL, cn)
        DT = rs.options.QueryDT

        If DT.Rows.Count > 0 Then
            btnP.Enabled = False
            btnN.Enabled = True
        Else
            btnP.Enabled = False
            btnN.Enabled = False
        End If

        ToolTip1.SetToolTip(btnNew, "New")
        ToolTip1.SetToolTip(btnP, "Previous")
        ToolTip1.SetToolTip(btnN, "Next")
        EnabledControl(False)

        If UserName = "admin" Then
            txtpass.PasswordChar = ""
            ReadUser(0)
        Else
            SQL = "select * from USERPASS where USERNAME='" & BaseClass.AES_Encrypt(UserName, "ABCDEFG") & "'"
            cn = New ADODB.Connection
            rs = New BaseClass.BaseODBCIO
            cn.Open(ConnVAT)
            rs.Open(SQL, cn)
            DT = rs.options.QueryDT
            btnP.Enabled = False
            btnN.Enabled = False
            btnDelete.Enabled = False
            btnNew.Enabled = False
            If DT.Rows.Count > 0 Then ReadUser(0)
        End If

    End Sub

    Public Sub EnabledControl(ByVal Value As Boolean)
        txtName.Enabled = Value
        txtusername.Enabled = Value
        txtpass.Enabled = Value
        txtusername.Enabled = Value
        txtconfirm.Enabled = Value
    End Sub

    Public Sub ReadUser(ByVal ID As Integer)
        Dim Dr As DataRow = DT.Rows(ID)
        If ID < DT.Rows.Count - 1 Then
            btnN.Enabled = True
        Else
            btnN.Enabled = False
        End If

        If ID >= 1 Then
            btnP.Enabled = True
        Else
            btnP.Enabled = False
        End If
        UserID = Dr("USERID").ToString
        txtName.Text = Dr("NAME").ToString.Trim
        Name = txtName.Text
        txtusername.Text = BaseClass.AES_Decrypt(Dr("USERNAME").ToString(), "ABCDEFG")
        User = txtusername.Text
        txtpass.Text = BaseClass.AES_Decrypt(Dr("PASS").ToString(), "ABCDEFG")
        pass = txtpass.Text
        txtpass.Text = BaseClass.AES_Decrypt(Dr("PASS").ToString(), "ABCDEFG")
        txtconfirm.Text = ""
        EnabledControl(True)
        If UserName <> "admin" Then
            btnP.Enabled = False
            btnN.Enabled = False
            btnDelete.Enabled = False
            btnNew.Enabled = False
        End If
        txtusername.Enabled = False
    End Sub

    Private Sub P_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnP.Click
        Index = Index - 1
        ReadUser(Index)
    End Sub

    Private Sub N_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnN.Click
        Index = Index + 1
        ReadUser(Index)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        btnP.Enabled = False
        btnN.Enabled = False
        txtusername.Text = ""
        txtpass.Text = ""
        txtconfirm.Text = ""
        txtName.Text = ""
        txtconfirm.Text = ""
        btnSave.Text = "&Add"
        txtusername.Enabled = True
        EnabledControl(True)
    End Sub

    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim SQL As String = ""
        Try
            SQL = "select * from USERPASS where NAME='" & txtusername.Text.Trim & "'"
            DT = New DataTable
            cn = New ADODB.Connection
            rs = New BaseClass.BaseODBCIO
            cn.Open(ConnVAT)
            SQL = SQL.Replace("[", "" & Convert.ToChar(34)).Replace("]", "" & Convert.ToChar(34))
            rs.Open(SQL, cn)
            Dim test As New DataTable
            test = rs.options.QueryDT
            'If test.Rows.Count > 0 Then
            '    Dim ob As New BaseVariable
            '    ob.ShowToolTip(txtconfirm, ToolTipIcon.Warning, "UserName alrady!" & Environment.NewLine & "Please try Other UserName")
            '    Exit Sub
            'End If
            If txtpass.Text <> txtconfirm.Text Then
                ShowToolTip(txtconfirm, ToolTipIcon.Warning, "Password Do not match." & Environment.NewLine & "Please try again")
                Exit Sub
            ElseIf test.Rows.Count > 0 Then
                ShowToolTip(txtusername, ToolTipIcon.Warning, "Are already in use." & Environment.NewLine & "Please try again")
                Exit Sub
            End If

            If btnSave.Text = "&Add" Then
                SQL = "select max(userid) from USERPASS "
                DT = New DataTable
                cn = New ADODB.Connection
                rs = New BaseClass.BaseODBCIO
                cn.Open(ConnVAT)
                SQL = SQL.Replace("[", "" & Convert.ToChar(34)).Replace("]", "" & Convert.ToChar(34))
                rs.Open(SQL, cn)
                Dim DTT As New DataTable
                DTT = rs.options.QueryDT

                SQL = " insert into USERPASS values(" & DTT.Rows(0).Item(0).ToString + 1 & ",'" & BaseClass.AES_Encrypt(txtusername.Text, "ABCDEFG") & "','" & BaseClass.AES_Encrypt(txtpass.Text, "ABCDEFG") & "','" & txtName.Text & "') " + Environment.NewLine
                cn = New ADODB.Connection
                cn.Open(ConnVAT)
                cn.Execute(SQL.Replace("[", "" & Convert.ToChar(34)).Replace("]", "" & Convert.ToChar(34)))
                'Dim com As SqlCommand = New SqlCommand(SQL.Replace("[", "" & Convert.ToChar(34)).Replace("]", "" & Convert.ToChar(34)), Conn)
                'com.ExecuteNonQuery()
                Index = Index + 1
                SQL = "select * from USERPASS order by USERID"
                DT = New DataTable
                cn = New ADODB.Connection
                rs = New BaseClass.BaseODBCIO
                cn.Open(ConnVAT)
                SQL = SQL.Replace("[", "" & Convert.ToChar(34)).Replace("]", "" & Convert.ToChar(34))
                rs.Open(SQL, cn)
                DT = rs.options.QueryDT


                'Dim adap As SqlDataAdapter = New SqlDataAdapter(SQL.Replace("[", "" & Convert.ToChar(34)).Replace("]", "" & Convert.ToChar(34)), Conn)
                'adap.Fill(DT)
                Index = DT.Rows.Count - 1
                'Dim Dr2() As DataRow = DT.Select("name='" & name & "'")
            Else
                SQL = " update USERPASS set USERNAME='" & BaseClass.AES_Encrypt(txtusername.Text, "ABCDEFG") & "',PASS='" & BaseClass.AES_Encrypt(txtpass.Text, "ABCDEFG") & "', NAME='" & txtName.Text & "' " + Environment.NewLine
                SQL &= "Where USERID='" & UserID & "'"
                cn = New ADODB.Connection
                cn.Open(ConnVAT)
                cn.Execute(SQL.Replace("[", "" & Convert.ToChar(34)).Replace("]", "" & Convert.ToChar(34)))

                If UserName = "admin" Then
                    SQL = "select * from USERPASS order by USERID"
                Else
                    SQL = "select * from USERPASS where USERNAME='" & BaseClass.AES_Encrypt(UserName, "ABCDEFG") & "'"
                End If

                DT = New DataTable
                cn = New ADODB.Connection
                rs = New BaseClass.BaseODBCIO
                cn.Open(ConnVAT)
                SQL = SQL.Replace("[", "" & Convert.ToChar(34)).Replace("]", "" & Convert.ToChar(34))
                rs.Open(SQL, cn)
                DT = rs.options.QueryDT

            End If
            btnSave.Text = "&Save"
            ReadUser(Index)
            MessageBox.Show("Save Sucess." & Environment.NewLine, "Process...", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Setup configuration Incomplete." & Environment.NewLine & "Please try again", "Warning!!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub BT_Delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If UserName = "admin" And txtusername.Text = "admin" Then
            MessageBox.Show("Cannot Delete Administrator User", "Warning!!", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Exit Sub
        End If

        If MessageBox.Show("Confirm Delete " & txtName.Text, "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> Windows.Forms.DialogResult.Yes Then Exit Sub
        Dim SQL As String = ""
        SQL = "Delete USERPASS where USERID='" & UserID & "'"
        cn = New ADODB.Connection
        cn.Open(ConnVAT)
        cn.Execute(SQL.Replace("[", "" & Convert.ToChar(34)).Replace("]", "" & Convert.ToChar(34)))
        SQL = "select * from USERPASS order by USERID"
        DT = New DataTable
        cn = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        cn.Open(ConnVAT)
        SQL = SQL.Replace("[", "" & Convert.ToChar(34)).Replace("]", "" & Convert.ToChar(34))
        rs.Open(SQL, cn)
        DT = rs.options.QueryDT
        Index = IIf(Index - 1 < 0, 0, Index - 1)
        ReadUser(Index)
    End Sub

    Private Sub txtusername_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtusername.TextChanged

    End Sub

    Private Sub txtpass_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtpass.TextChanged

    End Sub

    Private Sub txtconfirm_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtconfirm.TextChanged

    End Sub
End Class