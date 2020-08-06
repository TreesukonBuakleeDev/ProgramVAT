
Public Class FrmSelect
    Dim strSql As String = ""

    Private Sub btnFindF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim frm As FrmFrindBatch = New FrmFrindBatch(txtBatchF.Text)
        frm.ShowDialog()

        If frm._id <> "" Then
            txtBatchF.Text = frm._id
        End If
    End Sub

    Private Sub btnFindT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim frm As FrmFrindBatch = New FrmFrindBatch(txtBatchT.Text)
        frm.ShowDialog()
        If frm._id <> "" Then

            txtBatchT.Text = frm._id
        End If

    End Sub

    Private Sub btnSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelect.Click

        If cboModule.SelectedIndex = 0 Then
        ElseIf cboModule.SelectedIndex = 1 Then
        ElseIf cboModule.SelectedIndex = 2 Then

            strSql = " SELECT	GLPOST.BATCHNBR , GLPOST.ENTRYNBR, GLPOST.JRNLDATE,GLPJC.COMMENT,"
            strSql &= "         GLPOST.JNLDTLDESC, GLPOST.TRANSQTY, GLPOST.TRANSAMT, GLAMF.ACCTFMTTD, GLPOST.TRANSNBR"
            strSql &= " FROM    ((GLPOST"
            strSql &= "LEFT OUTER JOIN GLPJC ON (GLPOST.TRANSNBR = GLPJC.TRANSNBR) AND (GLPOST.ENTRYNBR = GLPJC.ENTRYNBR) AND (GLPOST.BATCHNBR = GLPJC.BATCHNBR))"
            strSql &= "INNER JOIN GLAMF ON GLPOST.ACCTID = GLAMF.ACCTID)"

            If ComDatabase = "PVSW" Then
                strSql &= "         LEFT JOIN " & ConnVAT.Split(";")(2).Split("=")(1) & "." & "FMSVLACC FMSVLACC ON (FMSVLACC.ACCTVAT = GLAMF.ACCTFMTTD) "
                strSql &= " WHERE   (GLPOST.SRCELEDGER = 'GL') AND (GLPOST.SRCETYPE <> 'NV')"
                strSql &= "         AND GLPOST.BATCHNBR >= '" & txtBatchF.Text & "' AND GLPOST.BATCHNBR <= '" & txtBatchT.Text & "')"
                strSql &= "         AND GLPOST.ENTRYNBR >= '" & txtEntryF.Text & "' AND GLPOST.ENTRYNBR <= '" & txtEntryT.Text & "')"
                strSql &= "         AND (GLAMF.ACCTFMTTD IN (select ACCTVAT from " & ConnVAT.Split(";")(2).Split("=")(1) & "." & "FMSVLACC))"

            Else
                strSql &= "         LEFT JOIN " & ConnVAT.Split(";")(2).Split("=")(1) & ".dbo." & "FMSVLACC FMSVLACC ON (FMSVLACC.ACCTVAT COLLATE Thai_CI_AS = GLAMF.ACCTFMTTD) "
                strSql &= " WHERE   (GLPOST.SRCELEDGER = 'GL') AND (GLPOST.SRCETYPE <> 'NV') "
                strSql &= "         AND GLPOST.BATCHNBR >= '" & txtBatchF.Text & "' AND GLPOST.BATCHNBR <= '" & txtBatchT.Text & "')"
                strSql &= "         AND GLPOST.ENTRYNBR >= '" & txtEntryF.Text & "' AND GLPOST.ENTRYNBR <= '" & txtEntryT.Text & "')"
                strSql &= "         AND (GLAMF.ACCTFMTTD IN (select ACCTVAT  COLLATE Thai_CI_AS from " & ConnVAT.Split(";")(2).Split("=")(1) & ".dbo." & "FMSVLACC))"

            End If

        ElseIf cboModule.SelectedIndex = 3 Then
        End If


    End Sub
End Class