Imports System
Imports System.Data
Imports System.IO
Imports System.Globalization
Imports System.Runtime.InteropServices

Friend Class frmMain
    Inherits Form
    Private lbLoad As Boolean
    Private strOrderBy As String
    Private frmOpasity As Form
    Dim Loadcomplate As Boolean = False
    Dim strReOrder As String
    Dim lbOrder As Boolean = False
    Dim PrintDestination As String
    Dim F As New clsFuntion
    Dim frmData As frmOpen
    Dim m_poperContainerForButton As SuperContextMenu.PoperContainer
    Public WithEvents m_popedContainerForButton As SuperContextMenu.ContextMenuForDeleteVat
    Dim DISCONNECTUSER As Integer = 0
    Dim CountDis As Decimal = 0

    Private Sub btnAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnAbout.Click
        Dim frmData As frmAbout
        frmData = New frmAbout

        frmData.ShowDialog()
        frmData = Nothing
    End Sub

    Public Sub btnOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOpen.Click
        Dim frmData3 As frmProcess
        Dim HistComDatabase As String
        Dim HistComACCPAC, HistComVAT As String

        HistComDatabase = ComDatabase
        HistComACCPAC = ComACCPAC
        HistComVAT = ComVAT
JClose:
        frmData = New frmOpen
        WriteLog("Select Company")
        frmData.ShowDialog()
        If frmData.DialogResult <> Windows.Forms.DialogResult.Yes And eventSender Is Nothing Then
            Application.Exit()
            Exit Sub
        End If
        'Mouse(MON)
        System.Windows.Forms.Cursor.Current = Cursors.AppStarting
        lstVu.Columns(1).Visible = Use_Approved
        HeaderCheckBox.Visible = Use_Approved
        If frmData.ClickSelect = True Then
            If BuildConnection(lAccpac, ComACCPAC, Me) = False Then GoTo JClose
            If BuildConnection(lVat, ComVAT, Me) = False Then GoTo JClose
            Call LoadDBScript()
            'หาก ไม่ได้ใช้ Module PO ให้ตัวเเปร UseINVFROMPO =false คือ ไม่ไปดึงเลขที่ PO Receipt , Po Invoice มาจาก Module PO
            If FindApp("PO") = False Then Use_INVFROMPO = False
            frmData3 = New frmProcess
            frmData3.Process = 0
            frmData3.ShowDialog(Me)

            OpenCompanyProcess()
            RefreshLocation()


            'Refresh view combo
            Call RefreshVatView(cboView)
            '********************************************
            ' Create Table VatInsert
            ' Import VatInsert to FMSVAT
            '*******************************************
            PrepareVatInsertTable()
            GetVatInsert()

            '***************End***********************
            GetVatlist()
            ListShow()


            Dim DESC() As String = GetPrivateProfile("COM" & ComID, "ACCPAC", "").Split(";")
            Dim DESCVAT() As String = GetPrivateProfile("COM" & ComID, "VAT", "").Split(";")

            If GetPrivateProfile("COM" & ComID, "LOGIN", "1") = "0" Then
                btnUserSet.Enabled = GetPrivateProfile("COM" & ComID, "LOGIN", "1")
            End If

            Dim SDESC As String = ""
            If ComDatabase = "MSSQL" Then
                SDESC += "Company :[" & GetPrivateProfile("COMPANY", "COM" & ComID, "").Split(";")(0) & "],"
                SDESC += "    SQL Server:[" & DESC(0).Trim() & "],"
                SDESC += "    DB Accpac:[" & DESC(1).Trim() & "],"
                SDESC += "    DB Vat:[" & DESCVAT(1).Trim() & "],"
                SDESC += "    Login:[" & UserLogin & "],"
            ElseIf ComDatabase = "PVSW" Then
                SDESC += "Company :[" & GetPrivateProfile("COMPANY", "COM" & ComID, "").Split(";")(0) & "],"
                SDESC += "    SQL Server:[" & DESC(0).Trim() & "],"
                SDESC += "    DB Accpac:[" & DESC(0).Trim() & "],"
                SDESC += "    DB Vat:[" & DESCVAT(0).Trim() & "],"
                SDESC += "    Login:[" & UserLogin & "],"
            End If

            ToolStripStatusLabel1.Text = SDESC
            Status.Text = ""
            ToolStripProgressBar1.Visible = False
        End If
        CheckCompany()
        CheckVatEnabled()
        'Mouse(MOFF)
        System.Windows.Forms.Cursor.Current = Cursors.Default

        frmData.Close()
        frmData3 = Nothing
    End Sub

    Private Sub OpenCompanyProcess()
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim nrs As BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim iffs As clsFuntion
        iffs = New clsFuntion
        Dim fi As DateTimeFormatInfo = New DateTimeFormatInfo
        Dim myDateFrom As DateTime '= DateTime.ParseExact(y & mm & "01", "yyyyMMdd", fi)
        Dim myDateTO As DateTime '= DateTime.ParseExact(y & mm & dd, "yyyyMMdd", fi)

        On Error GoTo ErrHandler

        cn = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO

        cn.Open(ConnVAT)
        sqlstr = "Select ORGID,CONAME,TAXNO,DATEFROM,DATETO from FMSVSET"
        rs.Open(sqlstr, cn)
        Do While rs.options.Rs.EOF = False
            ORGID = IIf(iffs.ISNULL(rs.options.Rs, 0), "", Trim(rs.options.Rs.Collect(0)))
            lblCompany.Text = IIf(iffs.ISNULL(rs.options.Rs, 1), "", Trim(rs.options.Rs.Collect(1)))
            TAXNO = IIf(iffs.ISNULL(rs.options.Rs, 2), "", Trim(rs.options.Rs.Collect(2)))

            fi = New DateTimeFormatInfo
            fi.ShortDatePattern = "yyyyMMdd"
            If CStr(IIf(iffs.ISNULL(rs.options.Rs, 3), "", rs.options.Rs.Collect(3))) <> "" Then
                myDateFrom = DateTime.ParseExact(IIf(iffs.ISNULL(rs.options.Rs, 3), "", rs.options.Rs.Collect(3)), "yyyyMMdd", fi)
                txtFrom.Value = myDateFrom
            End If
            If CStr(IIf(iffs.ISNULL(rs.options.Rs, 4), "", rs.options.Rs.Collect(4))) <> "" Then
                myDateTO = DateTime.ParseExact(IIf(iffs.ISNULL(rs.options.Rs, 4), "", rs.options.Rs.Collect(4)), "yyyyMMdd", fi)
                txtTo.Value = myDateTO
            End If
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop

        Dim i, m, y, d As Integer
        Dim mm, dd As String
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")
        m = CDbl(Date.Now.Month)
        y = CDbl(Date.Now.Year)
        If m = 1 Then
            m = 12
            y = y - 1
        Else
            m = m - 1
        End If
        d = F.GetDayInMonth(y, m)
        mm = CStr(m)
        dd = CStr(d)

        If Len(Trim(dd)) = 1 Then
            d = "0" & dd
        End If
        If Len(Trim(mm)) = 1 Then
            mm = "0" & mm
        End If

        fi = New DateTimeFormatInfo
        fi.ShortDatePattern = "yyyyMMdd"
        myDateFrom = DateTime.ParseExact(y & mm & "01", "yyyyMMdd", fi)
        myDateTO = DateTime.ParseExact(y & mm & dd, "yyyyMMdd", fi)

        txtFrom.Value = myDateFrom
        txtTo.Value = myDateTO


        rs = Nothing
        cn = Nothing
        Exit Sub

ErrHandler:
        WriteLog("<GetProfile>" & Err.Number & "  " & Err.Description)
        If Len(Err.Description) > 0 Then MessageBox.Show(Err.Description, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    Private Sub RefreshLocation()
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim i As Integer
        Dim sqlstr As String

        lbLoad = False
        cboLoc.Items.Clear()
        i = 0
        cn = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO

        sqlstr = "SELECT LOCID FROM FMSVLOC ORDER BY LOCID"

        cn.Open(ConnVAT)
        rs.Open(sqlstr, cn)
        cboLoc.Items.Add(New VB6.ListBoxItem("All Location", i))
        Do While rs.options.Rs.EOF = False
            cboLoc.Items.Add(Trim(rs.options.Rs.Collect(0)))
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        cn.Close()
        cboLoc.SelectedIndex = 0
        lbLoad = True
        rs = Nothing
        cn = Nothing
    End Sub
    'Select Data From FmsVat To Grid
    Public Sub GetVatlist()
        Dim loVat As clsVat
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim loClass As clsVat
        Dim ACCPAC_Conn As String
        Dim i As Double
        On Error GoTo ErrHandler

        If Len(ORGID) = 0 Then Exit Sub
        'Mouse(MON)
        System.Windows.Forms.Cursor.Current = Cursors.AppStarting

        cn = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        loVat = New clsVat
        mocolVat = New colVat
        mocolVatIN = New colVat
        mocolVatOut = New colVat
        mocolVatEx = New colVat

        If ComDatabase = "ORCL" Then
            sqlstr = "SELECT VATID,INVDATE,TXDATE,IDINVC,NEWDOCNO,DOCNO," '0-5
            sqlstr = sqlstr & " INVNAME,INVAMT,INVTAX,LOCID,VTYPE," '6-10
            sqlstr = sqlstr & " RATE,TTYPE,ACCTVAT,SOURCE,BATCH,ENTRY," '11-16
            sqlstr = sqlstr & " MARK," & Chr(34) & "VATCOMMENT" & Chr(34) & ",CBREF,TRANSNBR,RUNNO,TypeOfPU,TranNo,CIF,TaxCIF,isnull(Verify,0)as Verify,TAXID,BRANCH,Code FROM FMSVAT " '17-19
            sqlstr &= " WHERE " & IIf(DATEMODE = "DOCU", "INVDATE", "TXDATE") & " BETWEEN '" & txtFrom.Value.ToString("yyyyMMdd") & "' AND '" & txtTo.Value.ToString("yyyyMMdd") & "' "
            sqlstr = sqlstr & " ORDER BY " & strOrderBy
        Else
            sqlstr = "SELECT VATID,INVDATE,TXDATE,IDINVC,NEWDOCNO,DOCNO," '0-5
            sqlstr = sqlstr & " INVNAME,INVAMT,INVTAX,LOCID,VTYPE," '6-10
            sqlstr = sqlstr & " RATE,TTYPE,ACCTVAT,SOURCE,BATCH,ENTRY," '11-16
            sqlstr = sqlstr & " MARK,VATCOMMENT,CBREF,TRANSNBR,RUNNO,TypeOfPU,TranNo,CIF,TaxCIF,isnull(Verify,0)as Verify,TAXID,BRANCH,Code FROM FMSVAT " '17-19
            sqlstr &= " WHERE (" & IIf(DATEMODE = "DOCU", "INVDATE", "TXDATE") & " BETWEEN '" & txtFrom.Value.ToString("yyyyMMdd") & "' AND '" & txtTo.Value.ToString("yyyyMMdd") & "') "

            If S_AP Or S_AR Or S_GL Or S_CB Or S_PO Or s_OE Then
                sqlstr &= " AND SOURCE IN( 'US',"
                If S_AP Then sqlstr &= "'AP',"
                If S_AR Then sqlstr &= "'AR',"
                If S_GL Then sqlstr &= "'GL',"
                If S_CB Then sqlstr &= "'CB',"
                If S_PO Then sqlstr &= "'PO',"
                If s_OE Then sqlstr &= "'OE',"
                sqlstr = sqlstr.Substring(0, sqlstr.Length - 1) & ")"
            Else
                sqlstr &= " AND SOURCE IN('US')"
            End If

            sqlstr = sqlstr & " ORDER BY " & strOrderBy
        End If
        i = 0
        cn.Open(ConnVAT)
        rs.Open(sqlstr, cn)
        Do While rs.options.Rs.EOF = False
            loClass = New clsVat
            With loClass
                i += 1
                .VATID = IIf(ISNULL(rs.options.Rs, "VATID"), 0, (rs.options.Rs.Collect("VATID"))) 'FMSVAT.VATID
                .INVDATE = IIf(ISNULL(rs.options.Rs, "INVDATE"), 0, rs.options.Rs.Collect("INVDATE")) 'FMSVAT.INVDATE
                .TXDATE = IIf(ISNULL(rs.options.Rs, "TXDATE"), "", (rs.options.Rs.Collect("TXDATE"))) 'FMSVAT.TXDATE
                .IDINVC = IIf(ISNULL(rs.options.Rs, "IDINVC"), "", (rs.options.Rs.Collect("IDINVC"))) 'FMSVAT.IDINVC
                .NEWDOCNO = IIf(ISNULL(rs.options.Rs, "NEWDOCNO"), "", (rs.options.Rs.Collect("NEWDOCNO"))) 'FMSVAT.NEWDOCNO
                .DOCNO = IIf(ISNULL(rs.options.Rs, "DOCNO"), "", (rs.options.Rs.Collect("DOCNO"))) 'FMSVAT.DOCNO
                .INVNAME = IIf(ISNULL(rs.options.Rs, "INVNAME"), "", (rs.options.Rs.Collect("INVNAME"))) 'FMSVAT.INVNAME
                .INVAMT = IIf(ISNULL(rs.options.Rs, "INVAMT"), 0, rs.options.Rs.Collect("INVAMT")) 'FMSVAT.INVAMT
                .INVTAX = IIf(ISNULL(rs.options.Rs, "INVTAX") Or rs.options.Rs.Collect("INVTAX") = 0, 0, rs.options.Rs.Collect("INVTAX")) 'FMSVAT.INVTAX
                .LOCID = IIf(ISNULL(rs.options.Rs, "LOCID"), "", (rs.options.Rs.Collect("LOCID"))) 'FMSVAT.LOCID
                .VTYPE = IIf(ISNULL(rs.options.Rs, "VTYPE") Or rs.options.Rs.Collect("VTYPE") = 0, 0, rs.options.Rs.Collect("VTYPE")) 'FMSVAT.VTYPE
                If ISNULL(rs.options.Rs, "RATE") Or rs.options.Rs.Collect("RATE") = 0 Then
                    .RATE_Renamed = 0
                Else
                    .RATE_Renamed = rs.options.Rs.Collect("RATE")
                End If
                .TTYPE = IIf(ISNULL(rs.options.Rs, "TTYPE"), 0, rs.options.Rs.Collect("TTYPE")) 'FMSVAT.TTYPE
                .ACCTVAT = IIf(ISNULL(rs.options.Rs, "ACCTVAT"), "", (rs.options.Rs.Collect("ACCTVAT"))) 'FMSVAT.ACCTVAT
                .Source = IIf(ISNULL(rs.options.Rs, "Source"), "", (rs.options.Rs.Collect("Source"))) 'FMSVAT.SOURCE
                .Batch = IIf(ISNULL(rs.options.Rs, "Batch"), "", (rs.options.Rs.Collect("Batch"))) 'FMSVAT.BATCH
                .Entry = IIf(ISNULL(rs.options.Rs, "Entry"), "", (rs.options.Rs.Collect("Entry"))) 'FMSVAT.ENTRY
                .MARK = IIf(ISNULL(rs.options.Rs, "MARK"), "", (rs.options.Rs.Collect("MARK"))) 'FMSVAT.MARK
                .VATCOMMENT = IIf(ISNULL(rs.options.Rs, "VATCOMMENT"), "", (rs.options.Rs.Collect("VATCOMMENT"))) 'FMSVAT.VATCOMMENT
                .CBRef = IIf(ISNULL(rs.options.Rs, "CBRef"), "", (rs.options.Rs.Collect("CBRef"))) 'FMSVAT.CBREF
                .TRANSNBR = IIf(ISNULL(rs.options.Rs, "TRANSNBR"), "", (rs.options.Rs.Collect("TRANSNBR"))) 'FMSVAT.TRANSNBR
                .Runno = Trim(IIf(IsDBNull(rs.options.Rs.Collect("RUNNO")), "", rs.options.Rs.Collect("RUNNO")))
                .TypeOfPU = Trim(IIf(IsDBNull(rs.options.Rs.Collect("TypeOfPU")), "", rs.options.Rs.Collect("TypeOfPU")))
                .TranNo = Trim(IIf(IsDBNull(rs.options.Rs.Collect("TranNo")), "", rs.options.Rs.Collect("TranNo")))
                .CIF = Trim(IIf(IsDBNull(rs.options.Rs.Collect("CIF")), "0", rs.options.Rs.Collect("CIF")))
                .TaxCIF = Trim(IIf(IsDBNull(rs.options.Rs.Collect("TaxCIF")), "0", rs.options.Rs.Collect("TaxCIF")))
                .Verify = Trim(IIf(IsDBNull(rs.options.Rs.Collect("Verify")), "0", rs.options.Rs.Collect("Verify")))
                .TAXID = Trim(IIf(IsDBNull(rs.options.Rs.Collect("TAXID")), "0", rs.options.Rs.Collect("TAXID")))
                .Branch = Trim(IIf(IsDBNull(rs.options.Rs.Collect("Branch")), "0", rs.options.Rs.Collect("Branch")))
                '.Code1 = Trim(IIf(IsDBNull(rs.options.Rs.Collect("Code")), "0", rs.options.Rs.Collect("Code")))

                If (IIf(ISNULL(rs.options.Rs, "TTYPE"), 0, rs.options.Rs.Collect("TTYPE"))) = "1" Then
                    loVat = mocolVatIN.Add(.VATID, .INVDATE, .TXDATE, .IDINVC, .NEWDOCNO, .DOCNO, .INVNAME, .INVAMT, .INVTAX, .LOCID, CStr(.VTYPE), .RATE_Renamed, .TAXID, .Branch, .TTYPE, .ACCTVAT, .Source, .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, .Runno, .TypeOfPU, .TranNo, .CIF, .TaxCIF, .Verify, "key" & i)
                ElseIf (IIf(ISNULL(rs.options.Rs, "TTYPE"), 0, rs.options.Rs.Collect("TTYPE"))) = "2" Then
                    loVat = mocolVatOut.Add(.VATID, .INVDATE, .TXDATE, .IDINVC, .NEWDOCNO, .DOCNO, .INVNAME, .INVAMT, .INVTAX, .LOCID, CStr(.VTYPE), .RATE_Renamed, .TAXID, .Branch, .TTYPE, .ACCTVAT, .Source, .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, .Runno, .TypeOfPU, .TranNo, .CIF, .TaxCIF, .Verify, "key" & i)
                Else
                    loVat = mocolVatEx.Add(.VATID, .INVDATE, .TXDATE, .IDINVC, .NEWDOCNO, .DOCNO, .INVNAME, .INVAMT, .INVTAX, .LOCID, CStr(.VTYPE), .RATE_Renamed, .TAXID, .Branch, .TTYPE, .ACCTVAT, .Source, .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, .Runno, .TypeOfPU, .TranNo, .CIF, .TaxCIF, .Verify, "key" & i)
                End If

            End With
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        'Mouse(MOFF)
        System.Windows.Forms.Cursor.Current = Cursors.Default
        If cboView.SelectedIndex = 0 Then
            mocolVat = mocolVatIN
        ElseIf cboView.SelectedIndex = 1 Then
            mocolVat = mocolVatOut
        Else
            mocolVat = mocolVatEx
        End If
        rs = Nothing
        cn = Nothing
        Exit Sub
ErrHandler:

        WriteLog("<GetVatlist>" & Err.Number & "  " & Err.Description)
        If Len(Err.Description) > 0 Then MessageBox.Show(Err.Description, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    Private Sub GetVatlistF()
        Dim loVat As clsVat
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim loClass As clsVat
        Dim ACCPAC_Conn As String
        On Error GoTo ErrHandler

        If Len(ORGID) = 0 Then Exit Sub
        'Mouse(MON)
        System.Windows.Forms.Cursor.Current = Cursors.AppStarting

        cn = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        loVat = New clsVat
        mocolVat = New colVat
        If ComDatabase = "ORCL" Then
            sqlstr = "SELECT VATID,INVDATE,TXDATE,IDINVC,NEWDOCNO,DOCNO," '0-5
            sqlstr = sqlstr & " INVNAME,INVAMT,INVTAX,LOCID,VTYPE," '6-10
            sqlstr = sqlstr & " RATE,TTYPE,ACCTVAT,SOURCE,BATCH,ENTRY," '11-16
            sqlstr = sqlstr & " MARK," & Chr(34) & "VATCOMMENT" & Chr(34) & ",CBREF,TRANSNBR,RUNNO,TypeoFpu,tranno,cif,taxcif,TAXID,BRANCH FROM FMSVAT " '17-19
            sqlstr = sqlstr & " ORDER BY " & strOrderBy
        Else
            sqlstr = "SELECT VATID,INVDATE,TXDATE,IDINVC,NEWDOCNO,DOCNO," '0-5
            sqlstr = sqlstr & " INVNAME,INVAMT,INVTAX,LOCID,VTYPE," '6-10
            sqlstr = sqlstr & " RATE,TTYPE,ACCTVAT,SOURCE,BATCH,ENTRY," '11-16
            sqlstr = sqlstr & " MARK,VATCOMMENT,CBREF,TRANSNBR,RUNNO,TypeoFpu,tranno,cif,taxcif,TAXID,BRANCH FROM FMSVAT " '17-19
            sqlstr = sqlstr & " ORDER BY " & strOrderBy
        End If
        cn.Open(ConnVAT)
        rs.Open(sqlstr, cn)
        Do While rs.options.Rs.EOF = False
            loClass = New clsVat
            With loClass
                .VATID = IIf(IsDBNull(rs.options.Rs.Collect(0)), 0, Trim(rs.options.Rs.Collect(0))) 'FMSVAT.VATID
                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect(1)), 0, rs.options.Rs.Collect(1)) 'FMSVAT.INVDATE
                .TXDATE = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", Trim(rs.options.Rs.Collect(2))) 'FMSVAT.TXDATE
                .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect(3)), "", Trim(rs.options.Rs.Collect(3))) 'FMSVAT.IDINVC
                .NEWDOCNO = IIf(IsDBNull(rs.options.Rs.Collect(4)), "", Trim(rs.options.Rs.Collect(4))) 'FMSVAT.NEWDOCNO
                .DOCNO = IIf(IsDBNull(rs.options.Rs.Collect(5)), "", Trim(rs.options.Rs.Collect(5))) 'FMSVAT.DOCNO
                .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect(6)), "", Trim(rs.options.Rs.Collect(6))) 'FMSVAT.INVNAME
                .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7)) 'FMSVAT.INVAMT
                .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(8)) Or rs.options.Rs.Collect(8) = 0, 0, rs.options.Rs.Collect(8)) 'FMSVAT.INVTAX
                .LOCID = IIf(IsDBNull(rs.options.Rs.Collect(9)), "", Trim(rs.options.Rs.Collect(9))) 'FMSVAT.LOCID
                .VTYPE = IIf(IsDBNull(rs.options.Rs.Collect(10)) Or rs.options.Rs.Collect(10) = 0, 0, rs.options.Rs.Collect(10)) 'FMSVAT.VTYPE
                .RATE_Renamed = IIf(IsDBNull(rs.options.Rs.Collect(11) Or rs.options.Rs.Collect(11) = 0), 0, rs.options.Rs.Collect(11)) 'FMSVAT.RATE
                .TTYPE = IIf(IsDBNull(rs.options.Rs.Collect(12)), 0, rs.options.Rs.Collect(12)) 'FMSVAT.TTYPE
                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect(13)), "", Trim(rs.options.Rs.Collect(13))) 'FMSVAT.ACCTVAT
                .Source = IIf(IsDBNull(rs.options.Rs.Collect(14)), "", Trim(rs.options.Rs.Collect(14))) 'FMSVAT.SOURCE
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(15)), "", Trim(rs.options.Rs.Collect(15))) 'FMSVAT.BATCH
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(16)), "", Trim(rs.options.Rs.Collect(16))) 'FMSVAT.ENTRY
                .MARK = IIf(IsDBNull(rs.options.Rs.Collect(17)), "", Trim(rs.options.Rs.Collect(17))) 'FMSVAT.MARK
                .VATCOMMENT = IIf(IsDBNull(rs.options.Rs.Collect(18)), "", Trim(rs.options.Rs.Collect(18))) 'FMSVAT.VATCOMMENT
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(19)), "", Trim(rs.options.Rs.Collect(19))) 'FMSVAT.CBREF
                .TRANSNBR = IIf(IsDBNull(rs.options.Rs.Collect(20)), "", Trim(rs.options.Rs.Collect(20))) 'FMSVAT.TRANSNBR
                .Runno = Trim(IIf(IsDBNull(rs.options.Rs.Collect("RUNNO")), "", (rs.options.Rs.Collect("RUNNO"))))
                .TypeOfPU = Trim(IIf(IsDBNull(rs.options.Rs.Collect("TypeOfPU")), "", rs.options.Rs.Collect("TypeOfPU")))
                .TranNo = Trim(IIf(IsDBNull(rs.options.Rs.Collect("TranNo")), "", rs.options.Rs.Collect("TranNo")))
                .CIF = Trim(IIf(IsDBNull(rs.options.Rs.Collect("CIF")), 0, rs.options.Rs.Collect("CIF")))
                .TaxCIF = Trim(IIf(IsDBNull(rs.options.Rs.Collect("TaxCIF")), 0, rs.options.Rs.Collect("TaxCIF")))
                .TAXID = Trim(IIf(IsDBNull(rs.options.Rs.Collect("TAXID")), 0, rs.options.Rs.Collect("TAXID")))
                .Branch = Trim(IIf(IsDBNull(rs.options.Rs.Collect("BRANCH")), 0, rs.options.Rs.Collect("BRANCH")))
                '.Code = Trim(IIf(IsDBNull(rs.options.Rs.Collect("Code")), 0, rs.options.Rs.Collect("Code")))

                loVat = mocolVat.Add(.VATID, .INVDATE, .TXDATE, .IDINVC, .NEWDOCNO, .DOCNO, .INVNAME, .INVAMT, .INVTAX, .LOCID, CStr(.VTYPE), .RATE_Renamed, .TAXID, .Branch, .TTYPE, .ACCTVAT, .Source, .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, .Runno, .TypeOfPU, .TranNo, .CIF, .TaxCIF, "key" & .VATID)
            End With
            rs.options.Rs.MoveNext()
        Loop
        'Mouse(MOFF)
        System.Windows.Forms.Cursor.Current = Cursors.Default

        rs = Nothing
        cn = Nothing
        Exit Sub
ErrHandler:
        WriteLog("<GetVatlistF>" & Err.Number & "  " & Err.Description)
        If Len(Err.Description) > 0 Then MessageBox.Show(Err.Description, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    'Select Data From fmsvatinsert  and insert to table fmsvat
    Public Sub GetVatInsert()
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim rs1 As New BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim Cnn As ADODB.Connection
        On Error GoTo ErrHandler

        If Len(ORGID) = 0 Then Exit Sub
        'Mouse(MON)
        System.Windows.Forms.Cursor.Current = Cursors.AppStarting

        Cnn = New ADODB.Connection
        cn = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO

        If ComDatabase = "ORCL" Then
            sqlstr = "SELECT VATID,INVDATE,TXDATE,IDINVC,NEWDOCNO,DOCNO," '0-5
            sqlstr = sqlstr & " INVNAME,INVAMT,INVTAX,LOCID,VTYPE," '6-10
            sqlstr = sqlstr & " RATE,TTYPE,ACCTVAT,SOURCE,BATCH,ENTRY," '11-16
            sqlstr = sqlstr & " MARK," & Chr(34) & "VATCOMMENT" & Chr(34) & ",CBREF,STATUS,TRANSNBR,RUNNO,typeofPU,TranNo,CIF,taxCIF,isnull(Verify,0) as Verify,TAXID,BRANCH,CURRENCY,EXCHANGRATE,AMTBASETC FROM FMSVATINSERT " '17-21
            sqlstr = sqlstr & " WHERE (" & IIf(DATEMODE = "DOCU", "INVDATE", "TXDATE") & " BETWEEN '" & txtFrom.Value.ToString("yyyyMMdd") & "' AND '" & txtTo.Value.ToString("yyyyMMdd") & "') "
            sqlstr = sqlstr & " ORDER BY " & strOrderBy
        Else
            sqlstr = "SELECT VATID,INVDATE,TXDATE,IDINVC,NEWDOCNO,DOCNO," '0-5
            sqlstr = sqlstr & " INVNAME,INVAMT,INVTAX,LOCID,VTYPE," '6-10
            sqlstr = sqlstr & " RATE,TTYPE,ACCTVAT,SOURCE,BATCH,ENTRY," '11-16
            sqlstr = sqlstr & " MARK,VATCOMMENT,CBREF,STATUS,TRANSNBR ,RUNNO,typeofPU,TranNo,CIF,taxCIF,isnull(Verify,0) as Verify,TAXID,BRANCH,CURRENCY,EXCHANGRATE,AMTBASETC FROM FMSVATINSERT " '17-21
            sqlstr = sqlstr & " WHERE (" & IIf(DATEMODE = "DOCU", "INVDATE", "TXDATE") & " BETWEEN '" & txtFrom.Value.ToString("yyyyMMdd") & "' AND '" & txtTo.Value.ToString("yyyyMMdd") & "') "
            sqlstr = sqlstr & " ORDER BY " & strOrderBy
        End If
        cn.Open(ConnVAT)
        rs.Open(sqlstr, cn)


        sqlstr = "select max(VATID) FROM  FMSVAT"

        rs1.options.Rs = cn.Execute(sqlstr)
        If IsDBNull(rs1.options.Rs.Fields(0).Value) Then
            gVatId = 0
        Else
            gVatId = rs1.options.Rs.Fields(0).Value + 1
        End If

        Do While rs.options.Rs.EOF = False
            If Trim(rs.options.Rs.Fields("STATUS").Value) = "Edit" Then
                Dim Rrs As New BaseClass.BaseODBCIO
                Cnn = New ADODB.Connection
                Cnn.Open(ConnVAT)
                sqlstr = "select * from FMSVAT"
                sqlstr &= " WHERE LOCID='" & Trim(rs.options.Rs.Collect(9)) & "' AND SOURCE='" & Trim(rs.options.Rs.Collect(14)) & "' "
                sqlstr &= " AND BATCH='" & Trim(rs.options.Rs.Collect(15)) & "' AND ENTRY='" & Trim(rs.options.Rs.Collect(16)) & "' AND TRANSNBR='" & Trim(rs.options.Rs.Collect(21)) & "'"
                Rrs.Open(sqlstr, Cnn)
                If Rrs.options.QueryDT.Rows.Count > 0 Then
                    sqlstr = " Update FMSVAT  SET  INVDATE ='" & Trim(rs.options.Rs.Collect(1)) & "', TXDATE ='" & Trim(rs.options.Rs.Collect(2)) & "', IDINVC ='" & ReQuote(Trim(rs.options.Rs.Collect(3))) & "', "
                    sqlstr &= "DOCNO ='" & ReQuote(Trim(rs.options.Rs.Collect(5))) & "', NEWDOCNO ='" & Trim(rs.options.Rs.Collect(4)) & "', INVNAME ='" & ReQuote(Trim(rs.options.Rs.Collect(6))) & "', INVAMT ='" & rs.options.Rs.Collect(7) & "', "
                    sqlstr &= " INVTAX ='" & rs.options.Rs.Collect(8) & "', LOCID ='" & Trim(rs.options.Rs.Collect(9)) & "', VTYPE ='" & rs.options.Rs.Collect(10) & "', RATE ='" & rs.options.Rs.Collect(11) & "', TTYPE ='" & Trim(rs.options.Rs.Collect(12)) & "', "
                    'sqlstr &= " ACCTVAT ='" & Trim(rs.options.Rs.Collect(13)) & "', "
                    'MARK ='" & Trim(rs.options.Rs.Collect(17)) & "',"
                    sqlstr &= " VATCOMMENT ='" & ReQuote(Trim(rs.options.Rs.Collect(18))) & "', CBREF ='" & ReQuote(Trim(rs.options.Rs.Collect(19))) & "',"
                    sqlstr &= " TypeOfPU ='" & rs.options.Rs.Collect("TypeOfPU").ToString.Trim & "',"
                    sqlstr &= " TranNo='" & rs.options.Rs.Collect("TranNo").ToString.Trim & "',"
                    sqlstr &= " CIF='" & rs.options.Rs.Collect("CIF").ToString.Trim & "',"
                    sqlstr &= " TaxCIF='" & rs.options.Rs.Collect("TaxCIF").ToString.Trim & "',"
                    sqlstr &= " Verify='" & rs.options.Rs.Collect("Verify").ToString.Trim & "',"
                    sqlstr &= " TAXID='" & rs.options.Rs.Collect("TAXID").ToString.Trim & "',"
                    sqlstr &= " BRANCH='" & rs.options.Rs.Collect("BRANCH").ToString.Trim & "',"
                    sqlstr &= " CURRENCY='" & rs.options.Rs.Collect("CURRENCY").ToString.Trim & "',"
                    sqlstr &= " EXCHANGRATE='" & rs.options.Rs.Collect("EXCHANGRATE") & "',"
                    sqlstr &= " AMTBASETC='" & rs.options.Rs.Collect("AMTBASETC") & "'"

                    sqlstr &= " WHERE LOCID='" & Trim(rs.options.Rs.Collect(9)) & "' AND SOURCE='" & Trim(rs.options.Rs.Collect(14)) & "' "
                    sqlstr &= " AND BATCH='" & Trim(rs.options.Rs.Collect(15)) & "' AND ENTRY='" & Trim(rs.options.Rs.Collect(16)) & "' AND TRANSNBR='" & Trim(rs.options.Rs.Collect(21)) & "' AND VATID ='" & rs.options.Rs.Collect("VATID").ToString.Trim & "'"
                Else
                    sqlstr = "INSERT INTO FMSVAT"
                    sqlstr &= "   (INVDATE, TXDATE, IDINVC, DOCNO, NEWDOCNO, INVNAME, INVAMT, INVTAX, LOCID, VTYPE, RATE, TTYPE, ACCTVAT, SOURCE, BATCH,"
                    sqlstr &= " ENTRY, MARK, VATCOMMENT, CBREF, TRANSNBR,typeOfPU,tranNo,CIF,taxCIF,TAXID,BRANCH,Code,CURRENCY,EXCHANGRATE,AMTBASETC)"
                    sqlstr &= " VALUES     ('" & Trim(rs.options.Rs.Collect(1)) & "', '" & Trim(rs.options.Rs.Collect(2)) & "', '" & Trim(rs.options.Rs.Collect(3)) & "', '" & ReQuote(Trim(rs.options.Rs.Collect(5))) & "', '" & Trim(rs.options.Rs.Collect(4)) & "',"
                    sqlstr &= " '" & ReQuote(Trim(rs.options.Rs.Collect(6))) & "', '" & Trim(rs.options.Rs.Collect(7)) & "', '" & rs.options.Rs.Collect(8) & "', '" & Trim(rs.options.Rs.Collect(9)) & "', '" & rs.options.Rs.Collect(10) & "', '" & rs.options.Rs.Collect(11) & "', '" & Trim(rs.options.Rs.Collect(12)) & "',"
                    sqlstr &= " '" & Trim(rs.options.Rs.Collect(13)) & "', '" & Trim(rs.options.Rs.Collect(14)) & "', '" & Trim(rs.options.Rs.Collect(15)) & "',"
                    sqlstr &= " '" & Trim(rs.options.Rs.Collect(16)) & "', '" & Trim(rs.options.Rs.Collect(17)) & "', '" & ReQuote(Trim(rs.options.Rs.Collect(18))) & "', '" & Trim(rs.options.Rs.Collect(19)) & "', '" & Trim(rs.options.Rs.Collect(21)) & "',"
                    sqlstr &= " '" & rs.options.Rs.Collect("TypeOfPU").ToString.Trim & "','" & rs.options.Rs.Collect("TranNo").ToString.Trim & "','" & rs.options.Rs.Collect("CIF").ToString.Trim & "','" & rs.options.Rs.Collect("TaxCIF").ToString.Trim & "',"
                    sqlstr &= " '" & rs.options.Rs.Collect("TAXID").ToString.Trim & "','" & rs.options.Rs.Collect("BRANCH").ToString.Trim & "','','" & rs.options.Rs.Collect("CURRENCY").ToString.Trim & "','" & rs.options.Rs.Collect("EXCHANGRATE") & "','" & rs.options.Rs.Collect("AMTBASETC") & "')"
                End If
                'Dim idAA As String = rs.options.Rs.Collect("VATID").ToString.Trim
                Cnn.Execute(sqlstr)
                Cnn.Close()
            Else
                sqlstr = "select * FROM  FMSVAT"
                sqlstr = sqlstr & " WHERE LOCID='" & Trim(rs.options.Rs.Collect(9)) & "' AND SOURCE='" & Trim(rs.options.Rs.Collect(14)) & "' "
                sqlstr = sqlstr & " AND BATCH='" & Trim(rs.options.Rs.Collect(15)) & "' AND ENTRY='" & Trim(rs.options.Rs.Collect(16)) & "' AND TRANSNBR='" & Trim(rs.options.Rs.Collect(21)) & "'"
                sqlstr = sqlstr & " AND IDINVC='" & Trim(rs.options.Rs.Collect("IDINVC")) & "'"
                rs1.Open(sqlstr, cn)

                If rs1.options.Rs.RecordCount = 0 Then
                    If ComDatabase = "ORCL" Then
                        sqlstr = "INSERT INTO FMSVAT"
                        sqlstr &= "   (VATID, INVDATE, TXDATE, IDINVC, DOCNO, NEWDOCNO, INVNAME, INVAMT, INVTAX, LOCID, VTYPE, RATE, TTYPE, ACCTVAT, SOURCE, BATCH,"
                        sqlstr &= " ENTRY, MARK, VATCOMMENT, CBREF, TRANSNBR,typeOfPU,tranNo,CIF,taxCIF,TAXID,BRANCH,Code,CURRENCY,EXCHANGRATE,AMTBASETC)"
                        sqlstr &= " VALUES     ('" & gVatId & "', '" & Trim(rs.options.Rs.Collect(1)) & "', '" & Trim(rs.options.Rs.Collect(2)) & "', '" & Trim(rs.options.Rs.Collect(3)) & "', '" & ReQuote(Trim(rs.options.Rs.Collect(5))) & "', '" & Trim(rs.options.Rs.Collect(4)) & "',"
                        sqlstr &= " '" & ReQuote((Trim(rs.options.Rs.Collect(6)))) & "', '" & Trim(rs.options.Rs.Collect(7)) & "', '" & rs.options.Rs.Collect(8) & "', '" & Trim(rs.options.Rs.Collect(9)) & "', '" & rs.options.Rs.Collect(10) & "', '" & rs.options.Rs.Collect(11) & "', '" & Trim(rs.options.Rs.Collect(12)) & "',"
                        sqlstr &= " '" & Trim(rs.options.Rs.Collect(13)) & "', '" & Trim(rs.options.Rs.Collect(14)) & "', '" & Trim(rs.options.Rs.Collect(15)) & "',"
                        sqlstr &= " '" & Trim(rs.options.Rs.Collect(16)) & "', '" & Trim(rs.options.Rs.Collect(17)) & "', '" & ReQuote(Trim(rs.options.Rs.Collect(18))) & "', '" & Trim(rs.options.Rs.Collect(19)) & "', '" & Trim(rs.options.Rs.Collect(21)) & "',"
                        sqlstr &= " '" & rs.options.Rs.Collect("TypeOfPU").ToString.Trim & "','" & rs.options.Rs.Collect("TranNo").ToString.Trim & "','" & rs.options.Rs.Collect("CIF").ToString.Trim & "','" & rs.options.Rs.Collect("TaxCIF").ToString.Trim & "',"
                        sqlstr &= " '" & rs.options.Rs.Collect("TAXID").ToString.Trim & "','" & rs.options.Rs.Collect("BRANCH").ToString.Trim & "','','" & rs.options.Rs.Collect("CURRENCY").ToString.Trim & "','" & rs.options.Rs.Collect("EXCHANGRATE") & "','" & rs.options.Rs.Collect("AMTBASETC") & "')"
                    Else
                        sqlstr = "INSERT INTO FMSVAT"
                        sqlstr &= "   (INVDATE, TXDATE, IDINVC, DOCNO, NEWDOCNO, INVNAME, INVAMT, INVTAX, LOCID, VTYPE, RATE, TTYPE, ACCTVAT, SOURCE, BATCH,"
                        sqlstr &= " ENTRY, MARK, VATCOMMENT, CBREF, TRANSNBR,typeOfPU,tranNo,CIF,taxCIF,TAXID,BRANCH,Code,CURRENCY,EXCHANGRATE,AMTBASETC)"
                        sqlstr &= " VALUES     ('" & Trim(rs.options.Rs.Collect(1)) & "', '" & Trim(rs.options.Rs.Collect(2)) & "', '" & Trim(rs.options.Rs.Collect(3)) & "', '" & ReQuote(Trim(rs.options.Rs.Collect(5))) & "', '" & Trim(rs.options.Rs.Collect(4)) & "',"
                        sqlstr &= " '" & ReQuote(Trim(rs.options.Rs.Collect(6))) & "', '" & Trim(rs.options.Rs.Collect(7)) & "', '" & rs.options.Rs.Collect(8) & "', '" & Trim(rs.options.Rs.Collect(9)) & "', '" & rs.options.Rs.Collect(10) & "', '" & rs.options.Rs.Collect(11) & "', '" & Trim(rs.options.Rs.Collect(12)) & "',"
                        sqlstr &= " '" & Trim(rs.options.Rs.Collect(13)) & "', '" & Trim(rs.options.Rs.Collect(14)) & "', '" & Trim(rs.options.Rs.Collect(15)) & "',"
                        sqlstr &= " '" & Trim(rs.options.Rs.Collect(16)) & "', '" & Trim(rs.options.Rs.Collect(17)) & "', '" & ReQuote(Trim(rs.options.Rs.Collect(18))) & "', '" & Trim(rs.options.Rs.Collect(19)) & "', '" & Trim(rs.options.Rs.Collect(21)) & "',"
                        sqlstr &= " '" & rs.options.Rs.Collect("TypeOfPU").ToString.Trim & "','" & rs.options.Rs.Collect("TranNo").ToString.Trim & "','" & rs.options.Rs.Collect("CIF").ToString.Trim & "','" & rs.options.Rs.Collect("TaxCIF").ToString.Trim & "',"
                        sqlstr &= " '" & rs.options.Rs.Collect("TAXID").ToString.Trim & "','" & rs.options.Rs.Collect("BRANCH").ToString.Trim & "','','" & rs.options.Rs.Collect("CURRENCY").ToString.Trim & "','" & rs.options.Rs.Collect("EXCHANGRATE") & "','" & rs.options.Rs.Collect("AMTBASETC") & "')"
                    End If
                    Cnn = New ADODB.Connection
                    Cnn.Open(ConnVAT)
                    Cnn.Execute(sqlstr)
                    Cnn.Close()
                    gVatId = gVatId + 1
                Else
                    sqlstr = " Update FMSVAT  SET  INVDATE ='" & Trim(rs.options.Rs.Collect(1)) & "', TXDATE ='" & Trim(rs.options.Rs.Collect(2)) & "', IDINVC ='" & ReQuote(Trim(rs.options.Rs.Collect(3))) & "', "
                    sqlstr &= "DOCNO ='" & ReQuote(Trim(rs.options.Rs.Collect(5))) & "', NEWDOCNO ='" & Trim(rs.options.Rs.Collect(4)) & "', INVNAME ='" & ReQuote(Trim(rs.options.Rs.Collect(6))) & "', INVAMT ='" & rs.options.Rs.Collect(7) & "', "
                    sqlstr &= " INVTAX ='" & rs.options.Rs.Collect(8) & "', LOCID ='" & Trim(rs.options.Rs.Collect(9)) & "', VTYPE ='" & rs.options.Rs.Collect(10) & "', RATE ='" & rs.options.Rs.Collect(11) & "', TTYPE ='" & Trim(rs.options.Rs.Collect(12)) & "', "
                    'sqlstr &= " ACCTVAT ='" & Trim(rs.options.Rs.Collect(13)) & "', "
                    'MARK ='" & Trim(rs.options.Rs.Collect(17)) & "',"
                    sqlstr &= " VATCOMMENT ='" & ReQuote(Trim(rs.options.Rs.Collect(18))) & "', CBREF ='" & ReQuote(Trim(rs.options.Rs.Collect(19))) & "',"
                    sqlstr &= " TypeOfPU ='" & rs.options.Rs.Collect("TypeOFPU").ToString.Trim & "',"
                    sqlstr &= " TranNo='" & rs.options.Rs.Collect("TranNo").ToString.Trim & "',"
                    sqlstr &= " CIF='" & rs.options.Rs.Collect("CIF").ToString.Trim & "',"
                    sqlstr &= " TaxCIF='" & rs.options.Rs.Collect("TaxCIF").ToString.Trim & "',"
                    sqlstr &= " TAXID='" & rs.options.Rs.Collect("TAXID").ToString.Trim & "',"
                    sqlstr &= " BRANCH='" & rs.options.Rs.Collect("BRANCH").ToString.Trim & "',"
                    sqlstr &= " CURRENCY='" & rs.options.Rs.Collect("CURRENCY").ToString.Trim & "',"
                    sqlstr &= " EXCHANGRATE='" & rs.options.Rs.Collect("EXCHANGRATE") & "',"
                    sqlstr &= " AMTBASETC='" & rs.options.Rs.Collect("AMTBASETC") & "'"

                    sqlstr &= " WHERE LOCID='" & Trim(rs.options.Rs.Collect(9)) & "' AND SOURCE='" & Trim(rs.options.Rs.Collect(14)) & "' "
                    sqlstr &= " AND BATCH='" & Trim(rs.options.Rs.Collect(15)) & "' AND ENTRY='" & Trim(rs.options.Rs.Collect(16)) & "' AND TRANSNBR='" & Trim(rs.options.Rs.Collect(21)) & "'"
                    Cnn = New ADODB.Connection
                    Cnn.Open(ConnVAT)
                    Cnn.Execute(sqlstr)
                    Cnn.Close()
                End If

            End If
            rs.options.Rs.MoveNext()

            Application.DoEvents()
        Loop
        'Mouse(MOFF)
        System.Windows.Forms.Cursor.Current = Cursors.Default

        rs = Nothing
        cn = Nothing
        Cnn = Nothing
        Exit Sub
ErrHandler:
        WriteLog("<GetVatInsert>" & Err.Number & "  " & Err.Description)
        If Len(Err.Description) > 0 Then MessageBox.Show(Err.Description, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub
    '
    Public Sub ListShow()
        Dim loVTYPE As Integer
        Dim loRATE As Double
        Dim loVat As clsVat
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO

        On Error GoTo ErrHandler
        If Len(ORGID) = 0 Then Exit Sub
        'Mouse(MON)
        System.Windows.Forms.Cursor.Current = Cursors.AppStarting
        lstVu.Rows.Clear()
        If mocolVat Is Nothing Then
            'Mouse(MOFF)
            System.Windows.Forms.Cursor.Current = Cursors.Default

            Exit Sub
        End If
        cboPage.Items.Clear()
        Dim pageto As Integer = Math.Ceiling(mocolVat.Count / (Use_DisplayPage * 10)) - 1
        If Use_DisplayPage = 15 Then
            pageto = 0
        End If

        For c As Integer = 0 To pageto
            cboPage.Items.Add((c + 1))
        Next
        If cboPage.Items.Count = 0 Then
            cboPage.Items.Add(1)
        End If
        cboPage.SelectedIndex = 0
        CheckVatEnabled()
        Status.Text = ""
        Status.Image = Nothing
        ToolStripProgressBar1.Visible = False
        System.Windows.Forms.Cursor.Current = Cursors.Default
        Exit Sub

ErrHandler:
        WriteLog("<ListShow>" & Err.Number & "  " & Err.Description)
        If Len(Err.Description) > 0 Then MessageBox.Show(Err.Description, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    Private Sub CheckCompany()
        btnSetTaxNo.Enabled = Len(ORGID) > 0
        btnSetLocation.Enabled = Len(ORGID) > 0
        btnGetVat.Enabled = Len(ORGID) > 0
        btnAddVat.Enabled = Len(ORGID) > 0
        btnEditVat.Enabled = Len(ORGID) > 0
        btnDeleteVat.Enabled = Len(ORGID) > 0
        txtFrom.Enabled = Len(ORGID) > 0
        txtTo.Enabled = Len(ORGID) > 0
        btnGetVat.Enabled = Len(ORGID) > 0
        btnRefreh.Enabled = Len(ORGID) > 0
        btnRunning.Enabled = Len(ORGID) > 0
        btnRate.Enabled = Len(ORGID) > 0
        btnPrint.Enabled = Len(ORGID) > 0
    End Sub

    Private Sub CheckVatEnabled()
        btnGetVat.Enabled = Len(ORGID) > 0 'txtFrom.Enabled
        btnEditVat.Enabled = lstVu.Rows.Count > 0
        btnDeleteVat.Enabled = lstVu.Rows.Count > 0
    End Sub

    Private Sub FillList(ByRef selectvat As clsVat, ByVal i As Integer, ByVal DT As DataTable)
        With selectvat
            Try
                lstVu.EndEdit()
                Dim myDate As Date
                Dim fi As DateTimeFormatInfo = New DateTimeFormatInfo
                fi.ShortDatePattern = "yyyyMMdd"
                myDate = DateTime.ParseExact(CStr(.INVDATE).Trim, "yyyyMMdd", fi)
                lstVu.Rows.Add(New Object() {Format(i, "00000").Trim, .Verify, myDate.ToString("dd/MM/yyyy"), .IDINVC.Trim, .DOCNO.Trim, .INVNAME.Trim, Format(.INVAMT, "#,##0.00").Trim, Format(.INVTAX, "#,##0.00").Trim, .LOCID.Trim, IIf(.VTYPE = 0, "None", .RATE_Renamed), .TAXID, .Branch, .NEWDOCNO.Trim, .Source.Trim, .Batch.Trim, .Entry.Trim, .CBRef.Trim, .MARK.Trim, i, .Runno.Trim, .TypeOfPU, .TranNo, .CIF, .TaxCIF})
                If Math.Abs(.Verify) = 1 Then
                    CType(lstVu.Rows(lstVu.Rows.Count - 1).Cells(1), DataGridViewCheckBoxCell).Value = True
                    lstVu.EndEdit()
                End If
            Catch ex As Exception
                lstVu.Rows.Add(New Object() {Format(i, "00000").Trim, .Verify, "", .IDINVC.Trim, .DOCNO.Trim, .INVNAME.Trim, Format(.INVAMT, "#,##0.00").Trim, Format(.INVTAX, "#,##0.00").Trim, .LOCID.Trim, IIf(.VTYPE = 0, "None", .RATE_Renamed), .TAXID, .Branch, .NEWDOCNO.Trim, .Source.Trim, .Batch.Trim, .Entry.Trim, .CBRef.Trim, .MARK.Trim, i, .Runno.Trim, .TypeOfPU, .TranNo, .CIF, .TaxCIF})
            End Try

            Dim dr() As DataRow = DT.Select("VATID='" & selectvat.VATID & "'")
            If dr.Length > 0 Then
                lstVu.Rows(lstVu.Rows.Count - 1).DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 255, 255, 130)
            End If
        End With
    End Sub

    Private Sub cboLoc_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboLoc.SelectedIndexChanged
        If lbLoad Then ListShow()
    End Sub

    Private Sub cboRate_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboRate.SelectedIndexChanged
        txtRate.Visible = (VB6.GetItemData(cboRate, cboRate.SelectedIndex) = 1)
        lblRate.Visible = (VB6.GetItemData(cboRate, cboRate.SelectedIndex) = 1)
        If txtRate.Visible = True Then
            txtRate.Focus()
        Else
            ListShow()
        End If
    End Sub

    Private Sub cboView_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboView.SelectedIndexChanged
        On Error GoTo ErrHandler
        If cboView.SelectedIndex = 0 Then
            mocolVat = mocolVatIN
        ElseIf cboView.SelectedIndex = 1 Then
            mocolVat = mocolVatOut
        Else
            mocolVat = mocolVatEx
        End If
        If cboView.SelectedIndex + 1 = 2 Then
            lstVu.Columns.Item(4).Width = VB6.TwipsToPixelsX(2650.39 + 1400.31)
            lstVu.Columns.Item(4).Visible = False
        Else
            lstVu.Columns.Item(4).Visible = True
            lstVu.Columns.Item(4).Width = VB6.TwipsToPixelsX(2650.39)
            lstVu.Columns.Item(3).Width = VB6.TwipsToPixelsX(1400.31)
        End If
        ListShow()
        Exit Sub
ErrHandler:
        WriteLog("<cboView_SelectedIndexChanged>" & Err.Number & "  " & Err.Description)
        If Len(Err.Description) > 0 Then MessageBox.Show(Err.Description, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    Public Sub btnEditVat_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnEditVat.Click
        Dim frmData As frmAdjust
        Dim cn As ADODB.Connection
        Dim sqlstr As String = ""
        Dim loVATID As Integer
        Dim rs As New BaseClass.BaseODBCIO
        frmData = New frmAdjust(Me, "EDITVAT")
        Try
            If lstVu.Rows.Count = 0 Then Exit Sub
            If lstVu.CausesValidation = False Then

            End If
            loVATID = CDbl(lstVu.CurrentRow.Cells("PK").Value)
            Application.DoEvents()
            With frmData
                .VATID = mocolVatDisPlay.Item(loVATID).VATID
                .VDate = CStr(mocolVatDisPlay.Item(loVATID).TXDATE)
                .INVDATE = CStr(mocolVatDisPlay.Item(loVATID).INVDATE)
                .IDINVC = mocolVatDisPlay.Item(loVATID).IDINVC
                .DOCNO = mocolVatDisPlay.Item(loVATID).DOCNO
                .NEWDOCNO = mocolVatDisPlay.Item(loVATID).NEWDOCNO
                .INVNAME = mocolVatDisPlay.Item(loVATID).INVNAME
                .INVAMT = mocolVatDisPlay.Item(loVATID).INVAMT
                .TaxAMT = mocolVatDisPlay.Item(loVATID).INVTAX
                .LOCID = mocolVatDisPlay.Item(loVATID).LOCID
                .VTYPE = mocolVatDisPlay.Item(loVATID).VTYPE
                .TTYPE = mocolVatDisPlay.Item(loVATID).TTYPE
                .RATE_Renamed = mocolVatDisPlay.Item(loVATID).RATE_Renamed
                .VATCOMMENT = Trim(mocolVatDisPlay.Item(loVATID).VATCOMMENT)
                .Source = mocolVatDisPlay.Item(loVATID).Source
                .Batch = mocolVatDisPlay.Item(loVATID).Batch
                .Entry = mocolVatDisPlay.Item(loVATID).Entry
                .CBRef = mocolVatDisPlay.Item(loVATID).CBRef
                .DETAILNO = mocolVatDisPlay.Item(loVATID).TRANSNBR
                .RunNo = mocolVatDisPlay.Item(loVATID).Runno
                .TypeofPU = mocolVatDisPlay.Item(loVATID).TypeOfPU
                .TranNo = mocolVatDisPlay.Item(loVATID).TranNo
                .CIF = mocolVatDisPlay.Item(loVATID).CIF
                .TaxCIF = mocolVatDisPlay.Item(loVATID).TaxCIF
                .Verify = mocolVatDisPlay.Item(loVATID).Verify
                .TAXID = mocolVatDisPlay.Item(loVATID).TAXID
                .BRANCH = mocolVatDisPlay.Item(loVATID).Branch
                .ShowDialog()
            End With

            frmData = Nothing
            cn = Nothing
            Application.DoEvents()
            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.Message(), "Error...", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Public Function UpVat(ByRef loVATID As Decimal, ByRef loTXDATE As String, ByRef loIDINVC As String, ByRef loDOCNO As String, ByRef loNEWDOCNO As String, ByRef loINVNAME As String, ByRef loINVAMT As Double, ByRef loINVTAX As Double, ByRef loLOCID As String, ByRef loRATE As Double, ByRef loVTYPE As Short, ByRef loTTYPE As Short, ByRef loVATCOMMENT As String, ByRef loCBRef As String, ByRef loINVDATE As String, ByRef loRunNo As String, ByRef typeofPU As String, ByRef TranNo As String, ByRef CIF As Decimal, ByRef TaxCIF As Decimal, ByRef Verify As Integer, ByVal TaxID As String, ByVal Branch As String) As String

        UpVat = "UPDATE FMSVAT set TXDATE='" & Trim(loTXDATE) & "',"
        UpVat = UpVat & "INVDATE=" & Trim(loINVDATE) & ","
        UpVat = UpVat & "IDINVC='" & Trim(limitStr(ReQuote(loIDINVC), 255)) & "',"
        UpVat = UpVat & "NEWDOCNO='" & Trim(limitStr(ReQuote(loNEWDOCNO), 60)) & "',"
        UpVat = UpVat & "DOCNO='" & Trim(limitStr(ReQuote(loDOCNO), 60)) & "',"
        UpVat = UpVat & "INVNAME='" & Trim(limitStr(ReQuote(loINVNAME), 100)) & "',"
        UpVat = UpVat & "INVAMT=" & loINVAMT & ","
        UpVat = UpVat & "INVTAX=" & loINVTAX & ","
        UpVat = UpVat & "RATE=" & loRATE & ","
        UpVat = UpVat & "LOCID='" & limitStr(ReQuote(loLOCID), 6) & "',"
        UpVat = UpVat & "VATCOMMENT='" & Trim(limitStr(ReQuote(loVATCOMMENT), 250)) & "',"
        UpVat = UpVat & "CBREF='" & ReQuote(Trim(loCBRef)) & "',"
        UpVat = UpVat & "VTYPE=" & loVTYPE & ","
        'UpVat = UpVat & "MARK='',"
        UpVat = UpVat & "Runno='" & ReQuote(Trim(loRunNo)) & "',"
        UpVat = UpVat & "TTYPE=" & loTTYPE & ","
        UpVat = UpVat & "TypeOfPU='" & typeofPU & "',"
        UpVat = UpVat & "TranNo='" & TranNo & "',"
        UpVat = UpVat & "Cif='" & CIF & "',"
        UpVat = UpVat & "TaxCif='" & TaxCIF & "',"
        UpVat = UpVat & "Verify='" & Verify & "',"
        UpVat = UpVat & "TAXID='" & TaxID & "',"
        UpVat = UpVat & "BRANCH='" & Branch & "'"
        UpVat = UpVat & " WHERE VATID=" & Trim(CStr(loVATID))

    End Function

    Public Function StringStrip(ByRef tempStr As String) As String
        Dim pos As Integer
        pos = InStr(1, tempStr, "(", 1)
        If pos Then
            StringStrip = Mid(tempStr, 1, pos - 1)
        Else
            StringStrip = tempStr
        End If
    End Function

    Public Function UpEditVat(ByRef loVATID As Decimal, ByRef loTXDATE As String, ByRef loIDINVC As String, ByRef loDOCNO As String, ByRef loNEWDOCNO As String, ByRef loINVNAME As String, ByRef loINVAMT As Double, ByRef loINVTAX As Double, ByRef loLOCID As String, ByRef loRATE As Double, ByRef loVTYPE As Short, ByRef loTTYPE As Short, ByRef loVATCOMMENT As String, ByRef loCBRef As String, ByRef loINVDATE As String, ByRef loBatch As String, ByRef loEntry As String, ByRef loSource As String, ByRef loTRANSNBR As String, ByRef Runno As String, ByRef typeofPU As String, ByRef TranNo As String, ByRef CIF As Decimal, ByRef TaxCIF As Decimal, ByVal Verify As Integer, ByVal TaxID As String, ByVal Branch As String) As String

        UpEditVat = "UPDATE fmsvatinsert set TXDATE='" & Trim(loTXDATE) & "',"
        UpEditVat &= " INVDATE=" & Trim(loINVDATE) & ","
        UpEditVat &= " IDINVC='" & Trim(limitStr(ReQuote(loIDINVC), 255)) & "',"
        UpEditVat &= " NEWDOCNO='" & Trim(limitStr(ReQuote(loNEWDOCNO), 60)) & "',"
        UpEditVat &= " DOCNO='" & Trim(limitStr(ReQuote(loDOCNO), 60)) & "',"
        UpEditVat &= " INVNAME='" & (Trim(limitStr(ReQuote(loINVNAME), 100))) & "',"
        UpEditVat &= " INVAMT=" & loINVAMT & ","
        UpEditVat &= " INVTAX=" & loINVTAX & ","
        UpEditVat &= " RATE=" & loRATE & ","
        UpEditVat &= " LOCID='" & limitStr(ReQuote(loLOCID), 6) & "',"
        UpEditVat &= " VATCOMMENT='" & (Trim(limitStr(ReQuote(loVATCOMMENT), 250))) & "',"
        UpEditVat &= " CBREF='" & ReQuote(Trim(loCBRef)) & "',"
        UpEditVat &= " VTYPE=" & loVTYPE & ","
        UpEditVat &= " MARK='',"
        UpEditVat &= " Runno='" & ReQuote(Trim(Runno)) & "',"
        UpEditVat &= " TTYPE=" & loTTYPE & ","
        UpEditVat &= " TypeOFPU='" & typeofPU & "',"
        UpEditVat &= " TranNo='" & TranNo & "',"
        UpEditVat &= " CIF='" & CIF & "',"
        UpEditVat &= " TaxCIF='" & TaxCIF & "',"
        UpEditVat &= " Verify='" & Verify & "',"
        UpEditVat &= " TAXID='" & TaxID & "',"
        UpEditVat &= " BRANCH='" & Branch & "'"


        UpEditVat &= " WHERE (SOURCE='" & Trim(loSource) & "')"
        UpEditVat &= " AND (BATCH ='" & Trim(loBatch) & "') AND (ENTRY = '" & Trim(loEntry) & "') "
        UpEditVat &= " AND (LOCID = '" & Trim(limitStr(ReQuote(loLOCID), 6)) & "') AND (TRANSNBR='" & Trim(loTRANSNBR) & "')"
    End Function

    Private Sub btnRate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnRate.Click
        Dim frmData As frmRate
        WriteLog("Update Rate")
        frmData = New frmRate
        If frmData.ShowDialog() = Windows.Forms.DialogResult.Yes Then
            Call btnRefreh_Click(Nothing, Nothing)
        End If
        frmData = Nothing
    End Sub

    Public Sub btnRefreh_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnRefreh.Click
        Application.DoEvents()
        If Loadcomplate Then
            ShowLoading()
        End If
        Application.DoEvents()

        GetVatInsert() 'Update จาก FMSVATINSERT --> FMSVAT
        GetVatlist() 'ดึงจาก FMSVAT ออกมาเก็บในตัวเเปร
        ListShow()

        Application.DoEvents()

        CloseLoading()
    End Sub

    Private Sub btnRunning_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnRunning.Click
        Dim frmData As frmProcess
        Dim strsplit() As String
        Dim losplit As Integer
        On Error GoTo ErrHandler

        frmData = New frmProcess
        frmData.Process = 2

        If UCase(Trim(cboView.Text)) = "PURCHASE" Then
            ModGeneral.CheckTTypeRunning = 1
        ElseIf cboView.Text = "SALE" Then
            ModGeneral.CheckTTypeRunning = 2
        ElseIf cboView.Text = "EXPENSE" Then
            ModGeneral.CheckTTypeRunning = 3
        End If
        ModGeneral.CheckRate = IIf(cboRate.Text.Trim = "Only rate", "Only rate||" & txtRate.Text, cboRate.Text)

        frmData.ShowDialog(Me)
        frmData = Nothing
        If Loadcomplate Then
            ShowLoading()
        End If

        GetVatlist()
        ListShow()
        clsFuntion.FlushMemory()
        CloseLoading()


        Exit Sub
ErrHandler:
        WriteLog("<Running>" & Err.Number & "  " & Err.Description)
        If Len(Err.Description) > 0 Then MessageBox.Show(Err.Description, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    Public WithEvents Loading As New Timer
    Dim Tex As New TextBox
    Dim tic As Integer = 0
    Public Sub ShowLoading()
        tic = 0
        Loading.Interval = 250
        Loading.Start()
        Tex = New TextBox
        frmOpasity = New Form
        frmOpasity.StartPosition = FormStartPosition.Manual
        frmOpasity.ControlBox = False
        frmOpasity.MinimizeBox = False
        frmOpasity.MaximizeBox = False
        frmOpasity.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
        frmOpasity.Opacity = 0.5
        frmOpasity.BackColor = Color.Black
        frmOpasity.Location = New Point(Me.Location.X + 1, Me.Location.Y + 20) 'Me.Location '(New Point(lstVu.Location.X, lstVu.Location.Y))
        frmOpasity.Size = New Size(Me.Width - 2, Me.Height - 21)  ' lstVu.Size
        Tex.Text = "LOADING   "

        Tex.Enabled = False
        Tex.TextAlign = HorizontalAlignment.Center
        Tex.ReadOnly = True
        Tex.Size = New Size(frmOpasity.Width * (20 / 100), frmOpasity.Height * (20 / 100))
        Tex.Font = New System.Drawing.Font("Tahoma", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Tex.Location = New Point(frmOpasity.Width / 2 - Tex.Width / 2, frmOpasity.Height / 2)
        frmOpasity.Controls.Add(Tex)
        frmOpasity.ResumeLayout(False)
        frmOpasity.PerformLayout()
        frmOpasity.Show()
        Panel1.Enabled = False
        Panel2.Enabled = False
        Panel4.Enabled = False
        Application.DoEvents()
    End Sub

    Public Sub CloseLoading()
        If Not frmOpasity Is Nothing Then
            Loading.Stop()
            frmOpasity.Close()
            frmOpasity = Nothing
            Panel1.Enabled = True
            Panel2.Enabled = True
            Panel4.Enabled = True
        End If
    End Sub
    Private Sub btnPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnPrint.Click
        Dim FRMREPORT As New PriviewReport
        Dim rptPath, ReportFilename As String
        Dim ComOWNER, ComSVNAME As String
        Dim ComUSERNAME, ComUSERPSW As String
        Dim i As Integer
        Dim strsplit() As String

        rptPath = My.Application.Info.DirectoryPath & REPORT_PATH & IIf(ComDatabase = "MSSQL", "SQL", IIf(ComDatabase = "ORCL", "ORA", IIf(ComDatabase = "PVSW", "PSW", "")))

        ReportFilename = ""
        ReportChk = ""
        If cboView.SelectedIndex + 1 = 1 Then
            ReportFilename = "FVIN.rpt"
            FRMREPORT.TypeReport = 1
        ElseIf cboView.SelectedIndex + 1 = 2 Then
            ReportFilename = "FVOUT.rpt"
            FRMREPORT.TypeReport = 3
        ElseIf cboView.SelectedIndex + 1 = 3 Then
            ReportFilename = "FVINEXP.rpt"
            FRMREPORT.TypeReport = 2
        End If

        With dlgFile
            .InitialDirectory = rptPath
            .FileName = ReportFilename
            .Filter = "Report files (*.rpt)|*.rpt"
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                ReportFilename = .FileName
            Else
                Exit Sub
            End If
        End With
        ReportChk = ReportFilename
        If ComDatabase = "MSSQL" Then
            ComOWNER = GetPrivateProfile("COM" & ComID, "OWNER", "dbo")
        End If

        ComSVNAME = ""
        ComUSERNAME = ""
        ComUSERNAME = ""
        ComUSERPSW = ""
        strsplit = Split(ComVAT, ";")
        For i = LBound(strsplit) To UBound(strsplit)
            Select Case i
                Case 0
                    FRMREPORT.Server = strsplit(i)
                Case 1
                    FRMREPORT.DataBase = strsplit(i)
                Case 2
                    FRMREPORT.Userid = strsplit(i)
                Case 3
                    FRMREPORT.Password = strsplit(i)
            End Select

        Next


        FRMREPORT.Rate = txtRate.Text
        If cboLoc.Text = "All Location" Then
            FRMREPORT.Loc = ""
        Else
            FRMREPORT.Loc = cboLoc.Text
        End If

        FRMREPORT.txtFrom = txtFrom.Value.ToString("yyyyMMdd")
        FRMREPORT.txtTo = txtTo.Value.ToString("yyyyMMdd")
        FRMREPORT.Company = lblCompany.Text
        FRMREPORT.TAXNO = TAXNO
        FRMREPORT.Vtype = cboRate.SelectedIndex
        FRMREPORT.VTType = cboView.SelectedIndex + 1
        If DATEMODE = "DOCU" Then
            FRMREPORT.sSelection = "{FMSVAT.TXDATE}"
            FRMREPORT.sReplace = "{FMSVAT.INVDATE}"
        Else
            FRMREPORT.sSelection = "{FMSVAT.INVDATE}"
            FRMREPORT.sReplace = "{FMSVAT.TXDATE}"
        End If

        FRMREPORT.ShowDialog()
        clsFuntion.FlushMemory()

    End Sub

    Private Sub frmMain_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        If frmOpasity Is Nothing Then Exit Sub
        frmOpasity.Activate()
    End Sub

    Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Me.Hide()
        Dim i, m, y, d As Integer
        Dim mm, dd As String
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")
        m = CDbl(Date.Now.Month)
        y = CDbl(Date.Now.Year)
        If m = 1 Then
            m = 12
            y = y - 1
        Else
            m = m - 1
        End If
        d = F.GetDayInMonth(y, m)
        mm = CStr(m)
        dd = CStr(d)

        If Len(Trim(dd)) = 1 Then
            d = "0" & dd
        End If
        If Len(Trim(mm)) = 1 Then
            mm = "0" & mm
        End If

        Dim fi As DateTimeFormatInfo = New DateTimeFormatInfo
        fi.ShortDatePattern = "yyyyMMdd"
        Dim myDateFrom As DateTime = DateTime.ParseExact(y & mm & "01", "yyyyMMdd", fi)
        Dim myDateTO As DateTime = DateTime.ParseExact(y & mm & dd, "yyyyMMdd", fi)

        txtFrom.Value = myDateFrom
        txtTo.Value = myDateTO

        StrYear = y

        Dim frmSplash As New BaseClass.frmSplash
        frmSplash.ShowDialog()
        Me.Width = Screen.PrimaryScreen.WorkingArea.Width
        Me.Height = Screen.PrimaryScreen.WorkingArea.Height
        Me.Location = New Point(0, 0)
        Me.MinimumSize = New Size(Me.Width - (Me.Width * 14 / 100), Me.Height - (Me.Height * 14 / 100))

        Label2.Width = Me.Width - 16
        lblCompany.Width = Me.Width - 16
        lstVu.Top = VB6.TwipsToPixelsY(1125)
        lstVu.Left = 3
        lstVu.Width = Me.Width - 23
        lstVu.Height = Me.Height - 190
        GroupBox1.Location = New Point(3, Me.Height - 115)
        GroupBox1.Width = Me.Width - 23

        btnAbout.Location = New Point(Me.Width - 100, 9)
        ListHeader()
        strOrderBy = "TXDATE,DOCNO"
        cboView.Items.Add(New VB6.ListBoxItem("", 0))
        cboView.SelectedIndex = 0

        For i = 0 To 3
            cboRate.Items.Add(New VB6.ListBoxItem(My.Resources.ResourceManager.GetString("str" + CStr(200 + i)), i))
        Next
        cboRate.SelectedIndex = 0
        CreateLog()
        WriteLog("Open application")

        Me.Text = My.Application.Info.ProductName & " " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build

        If ComDatabase = "ORCL" Then
            Me.Text = "FMS Vat Report " & My.Application.Info.ProductName & " " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & " For Oracle 10g"
        ElseIf ComDatabase = "MSSQL" Then
            Me.Text = "FMS VAT Report " & My.Application.Info.ProductName & " " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & " For SQL Server"
        ElseIf ComDatabase = "PVSW" Then
            Me.Text = "FMS VAT Report " & My.Application.Info.ProductName & " " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & " For Pervasive"
        End If

        Call btnOpen_Click(Nothing, Nothing)
        If frmData.DialogResult <> Windows.Forms.DialogResult.Yes Then
            Application.Exit()
            Exit Sub
        End If
        Dim frm As New FrmLogin

        If GetPrivateProfile("COM" & ComID, "LOGIN", "0").Split("")(0) = "1" Then
            frm.ShowDialog()
            If frm.DialogResult <> Windows.Forms.DialogResult.Yes Then
                Application.Exit()
                Me.Close()
                Exit Sub
            End If
        Else
            UserLogin = Environment.MachineName
            UserName = "admin"
        End If

        If ClickSelect1 = True Then
            GetVatInsert()
            GetVatlist()
            ListShow()
        End If
        ToolStripStatusLabel1.AutoSize = True
        Dim DESC() As String = GetPrivateProfile("COM" & ComID, "ACCPAC", "").Split(";")
        Dim DESCVAT() As String = GetPrivateProfile("COM" & ComID, "VAT", "").Split(";")

        If GetPrivateProfile("COM" & ComID, "LOGIN", "1") = "0" Then
            btnUserSet.Enabled = GetPrivateProfile("COM" & ComID, "LOGIN", "1")
        End If
        If UserName <> "admin" Then
            btnDBSet.Enabled = False
        End If

        Dim SDESC As String = ""
        If ComDatabase = "MSSQL" Then
            SDESC += "Company :[" & GetPrivateProfile("COMPANY", "COM" & ComID, "").Split(";")(0) & "],"
            SDESC += "    SQL Server:[" & DESC(0).Trim() & "],"
            SDESC += "    DB Accpac:[" & DESC(1).Trim() & "],"
            SDESC += "    DB Vat:[" & DESCVAT(1).Trim() & "],"
            SDESC += "    Login:[" & UserLogin & "],"
        ElseIf ComDatabase = "PVSW" Then
            SDESC += "Company :[" & GetPrivateProfile("COMPANY", "COM" & ComID, "").Split(";")(0) & "],"
            SDESC += "    SQL Server:[" & DESC(0).Trim() & "],"
            SDESC += "    DB Accpac:[" & DESC(0).Trim() & "],"
            SDESC += "    DB Vat:[" & DESCVAT(0).Trim() & "],"
            SDESC += "    Login:[" & UserLogin & "],"
        End If

        ToolStripStatusLabel1.Text = SDESC
        clsFuntion.FlushMemory()
        Loadcomplate = True

        If Screen.PrimaryScreen.Bounds.Width > 1024 Then
            txtFrom.Width += 20

            LbTo.Left += 20
            txtTo.Left += 20
            txtTo.Width += 20

            LbLocation.Left += 40
            cboLoc.Left += 40
            cboLoc.Width += 20

            LbView.Left += 60
            cboView.Left += 60
            cboView.Width += 20

            LbRate.Left += 80
            cboRate.Left += 80
            cboRate.Width += 20

            txtRate.Left += 100
            lblRate.Left += 100
        End If
        Me.Show()
        Timer1.Enabled = True

    End Sub

    Private Sub ListHeader()
        lstVu.AllowUserToAddRows = False
        lstVu.AllowUserToDeleteRows = False
        lstVu.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        lstVu.ReadOnly = False
        lstVu.MultiSelect = False
        lstVu.Columns.Add("VATID", "No.") ', CInt(VB6.TwipsToPixelsX(700)))
        lstVu.Columns(0).Width = CInt(VB6.TwipsToPixelsX(700))
        lstVu.Columns(0).Visible = False
        lstVu.Columns(0).ReadOnly = True

        Dim Col As New DataGridViewCheckBoxColumn()
        Col.Name = "Verify"
        Col.HeaderText = "Approved"
        Col.ReadOnly = False
        Col.SortMode = DataGridViewColumnSortMode.Programmatic
        Me.lstVu.Columns.Add(Col)
        lstVu.Columns(1).Width = CInt(VB6.TwipsToPixelsX(1200))
        lstVu.Columns(1).Visible = Use_Approved

        lstVu.Columns.Add("TXDATE", "Date") ', CInt(VB6.TwipsToPixelsX(1050)))
        lstVu.Columns(2).Width = CInt(VB6.TwipsToPixelsX(1200))
        lstVu.Columns(2).ReadOnly = True

        lstVu.Columns.Add("IDINVC", "Inv No") ', CInt(VB6.TwipsToPixelsX(1600)))
        lstVu.Columns(3).Width = CInt(VB6.TwipsToPixelsX(1600))
        lstVu.Columns(3).ReadOnly = True

        lstVu.Columns.Add("DOCNO", "Doc No") ', CInt(VB6.TwipsToPixelsX(1400)))
        lstVu.Columns(4).Width = CInt(VB6.TwipsToPixelsX(1400))
        lstVu.Columns(4).ReadOnly = True

        lstVu.Columns.Add("INVNAME", "Name") ', CInt(VB6.TwipsToPixelsX(3000)))
        lstVu.Columns(5).Width = CInt(VB6.TwipsToPixelsX(3000))
        lstVu.Columns(5).ReadOnly = True

        lstVu.Columns.Add("INVAMT", "Amount") ', CInt(VB6.TwipsToPixelsX(1440)), System.Windows.Forms.HorizontalAlignment.Right, "")
        lstVu.Columns(6).Width = CInt(VB6.TwipsToPixelsX(1440))
        lstVu.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
        lstVu.Columns(6).ReadOnly = True

        lstVu.Columns.Add("INVTAX", "Tax Amount") ', CInt(VB6.TwipsToPixelsX(1440)), System.Windows.Forms.HorizontalAlignment.Right, "")
        lstVu.Columns(7).Width = CInt(VB6.TwipsToPixelsX(1440))
        lstVu.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
        lstVu.Columns(7).ReadOnly = True

        lstVu.Columns.Add("LOCID", "Location") ', CInt(VB6.TwipsToPixelsX(900)))
        lstVu.Columns(8).Width = CInt(VB6.TwipsToPixelsX(900))
        lstVu.Columns(8).ReadOnly = True

        lstVu.Columns.Add("RATE", "Rate") ', CInt(VB6.TwipsToPixelsX(600)))
        lstVu.Columns(9).Width = CInt(VB6.TwipsToPixelsX(600))
        lstVu.Columns(9).ReadOnly = True

        'เเสดง TAXID
        lstVu.Columns.Add("TAXID", "TaxID") ', CInt(VB6.TwipsToPixelsX(600)))
        lstVu.Columns(10).Width = CInt(VB6.TwipsToPixelsX(1440))
        lstVu.Columns(10).ReadOnly = True

        'เเสดง BRANCH
        lstVu.Columns.Add("BRANCH", "Branch") ', CInt(VB6.TwipsToPixelsX(600)))
        lstVu.Columns(11).Width = CInt(VB6.TwipsToPixelsX(600))
        lstVu.Columns(11).ReadOnly = True


        lstVu.Columns.Add("NEWDOCNO", "Running") ', CInt(VB6.TwipsToPixelsX(1440)))
        lstVu.Columns(12).Width = CInt(VB6.TwipsToPixelsX(1440))
        lstVu.Columns(12).ReadOnly = True

        lstVu.Columns.Add("SOURCE", "Source") ', CInt(VB6.TwipsToPixelsX(800)))
        lstVu.Columns(13).Width = CInt(VB6.TwipsToPixelsX(800))
        lstVu.Columns(13).ReadOnly = True

        lstVu.Columns.Add("BATCH", "Batch") ', CInt(VB6.TwipsToPixelsX(1000)))
        lstVu.Columns(14).ValueType = Type.GetType("System.Decimal")
        lstVu.Columns(14).Width = CInt(VB6.TwipsToPixelsX(1000))
        lstVu.Columns(14).ReadOnly = True

        lstVu.Columns.Add("ENTRY", "Entry") ', CInt(VB6.TwipsToPixelsX(1000)))
        lstVu.Columns(15).ValueType = Type.GetType("System.Decimal")
        lstVu.Columns(15).Width = CInt(VB6.TwipsToPixelsX(1000))
        lstVu.Columns(15).ReadOnly = True

        lstVu.Columns.Add("CBREF", "Ref") ', CInt(VB6.TwipsToPixelsX(1440)))
        lstVu.Columns(16).Width = CInt(VB6.TwipsToPixelsX(1440))
        lstVu.Columns(16).ReadOnly = True

        lstVu.Columns.Add("MARK", "Mark") ', CInt(VB6.TwipsToPixelsX(1000)))
        lstVu.Columns(17).Width = CInt(VB6.TwipsToPixelsX(1000))
        lstVu.Columns(17).ReadOnly = True

        lstVu.Columns.Add("PK", "PK") ', CInt(VB6.TwipsToPixelsX(0)))
        lstVu.Columns(18).Width = CInt(VB6.TwipsToPixelsX(0))
        lstVu.Columns(18).Visible = False
        lstVu.Columns(18).ReadOnly = True

        lstVu.Columns.Add("RunNo", "RunNo") ', CInt(VB6.TwipsToPixelsX(1440)))
        lstVu.Columns(19).Width = CInt(VB6.TwipsToPixelsX(1440))
        lstVu.Columns(19).ReadOnly = True

        lstVu.Columns.Add("TypeOfPU", "TypeOfPU") ', CInt(VB6.TwipsToPixelsX(1440)))
        lstVu.Columns(20).Width = CInt(VB6.TwipsToPixelsX(1440))
        lstVu.Columns(20).ReadOnly = True

        lstVu.Columns.Add("TranNo", "TranNo") ', CInt(VB6.TwipsToPixelsX(1440)))
        lstVu.Columns(21).Width = CInt(VB6.TwipsToPixelsX(1440))
        lstVu.Columns(21).ReadOnly = True

        lstVu.Columns.Add("CIF", "CIF") ', CInt(VB6.TwipsToPixelsX(1440)))
        lstVu.Columns(22).Width = CInt(VB6.TwipsToPixelsX(1440))
        lstVu.Columns(22).ReadOnly = True

        lstVu.Columns.Add("TaxCIF", "TaxCIF") ', CInt(VB6.TwipsToPixelsX(1440)))
        lstVu.Columns(23).Width = CInt(VB6.TwipsToPixelsX(1440))
        lstVu.Columns(23).ReadOnly = True


        'lstVu.LabelEdit = True
        AddHeaderCheckBox()
        AddHandler HeaderCheckBox.KeyUp, AddressOf HeaderCheckBox_KeyUp
        AddHandler HeaderCheckBox.MouseClick, AddressOf HeaderCheckBox_MouseClick
        AddHandler lstVu.CurrentCellDirtyStateChanged, AddressOf dgvSelectAll_CurrentCellDirtyStateChanged
        AddHandler lstVu.CellPainting, AddressOf dgvSelectAll_CellPainting


    End Sub

    Public HeaderCheckBox As CheckBox
    Private Sub AddHeaderCheckBox()
        HeaderCheckBox = New CheckBox()
        HeaderCheckBox.Size = New Size(15, 15)
        lstVu.Controls.Add(HeaderCheckBox)
        HeaderCheckBox.Visible = Use_Approved
    End Sub
    Private Sub ResetHeaderCheckBoxLocation(ByVal ColumnIndex As Integer, ByVal RowIndex As Integer)
        Dim Rect As Rectangle = lstVu.GetCellDisplayRectangle(ColumnIndex, RowIndex, True)
        Dim Pt As New Point()
        Pt.X = Rect.Location.X + (Rect.Width - HeaderCheckBox.Width) - 2
        Pt.Y = Rect.Location.Y + (Rect.Height - HeaderCheckBox.Height) \ 2
        HeaderCheckBox.Location = Pt
    End Sub
    Private Sub HeaderCheckBox_MouseClick(ByVal sender As Object, ByVal e As MouseEventArgs)
        HeaderCheckBoxClick(DirectCast(sender, CheckBox))
    End Sub
    Private Sub HeaderCheckBoxClick(ByVal HCheckBox As CheckBox)
        lstVu.EndEdit()
        For i As Integer = 0 To lstVu.Rows.Count - 1
            DirectCast(lstVu.Rows(i).Cells("Verify"), DataGridViewCheckBoxCell).Value = HCheckBox.Checked
            lstVu_CellContentClick(lstVu, New DataGridViewCellEventArgs(1, i))
        Next
        lstVu.EndEdit()
    End Sub
    Private Sub HeaderCheckBox_KeyUp(ByVal sender As Object, ByVal e As KeyEventArgs)
        If e.KeyCode = Keys.Space Then
            HeaderCheckBoxClick(DirectCast(sender, CheckBox))
        End If
    End Sub
    Private Sub dgvSelectAll_CellPainting(ByVal sender As Object, ByVal e As DataGridViewCellPaintingEventArgs)
        If Use_Approved = False Then Exit Sub
        If e.RowIndex = -1 And e.ColumnIndex = 1 And lstVu.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, True).Width > HeaderCheckBox.Width Then
            HeaderCheckBox.Visible = True
            ResetHeaderCheckBoxLocation(e.ColumnIndex, e.RowIndex)
        ElseIf lstVu.FirstDisplayedCell IsNot Nothing Then
            If lstVu.FirstDisplayedCell.ColumnIndex <> 1 Then
                HeaderCheckBox.Visible = False
            End If
        End If
    End Sub
    Private Sub dgvSelectAll_CurrentCellDirtyStateChanged(ByVal sender As Object, ByVal e As EventArgs)
        If TypeOf lstVu.CurrentCell Is DataGridViewCheckBoxCell Then
            lstVu.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub


    Private Sub btnAddVat_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnAddVat.Click
        On Error GoTo ErrHandler
        Dim frmData As frmAdjust
        Dim cn As ADODB.Connection
        Dim DETAILNO As String
        Dim rs As BaseClass.BaseODBCIO
        frmData = New frmAdjust(Me, "ADDVAT")
        With frmData
            .VDate = Year(Now) & VB6.Format(Now, "MMDD")
            .INVDATE = Year(Now) & VB6.Format(Now, "MMDD")
            .IDINVC = ""
            .DOCNO = ""
            .NEWDOCNO = ""
            .INVNAME = ""
            .INVAMT = 0
            .TaxAMT = 0
            .LOCID = ""
            .TTYPE = 1
            .VTYPE = 1
            .RATE_Renamed = 0
            .VATCOMMENT = ""
            .CBRef = ""
            .Source = "US"
            DETAILNO = ""
            .ShowDialog()
        End With

        If frmData.DialogResult = Windows.Forms.DialogResult.Yes Then
            Call btnRefreh_Click(Nothing, Nothing)
            Me.lstVu.Select()
            If frmData.LOCID.Trim = cboLoc.Text.Trim Then
                lstVu.CurrentCell = lstVu.Rows(lstVu.RowCount - 1).Cells(1)
                Me.lstVu.CurrentRow.Selected = True
            End If
        End If

        frmData = Nothing
        cn = Nothing
        rs = Nothing
        Exit Sub
ErrHandler:
        WriteLog("<cmdAddVat_Click>" & Err.Number & "  " & Err.Description)
        If Len(Err.Description) > 0 Then MessageBox.Show(Err.Description, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    Private Sub ExitApp()
        WriteLog("Exit application")
        Me.Close()
        clsFuntion.FlushMemory()
    End Sub

    Private Sub btnExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnExit.Click
        ExitApp()
    End Sub

    Public Function FormatDate(ByVal DateNum As String) As String
        Dim mm As String
        Dim dd As String
        Dim yy As String
        Dim dte As String = Trim(DateNum)
        yy = Mid(dte, 1, 4)
        mm = Mid(dte, 5, 2)
        dd = Mid(dte, 7, 2)
        FormatDate = dd & "/" & mm & "/" & yy
    End Function

    Private Sub btnGetVat_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnGetVat.Click
        On Error GoTo ErrHandler
        Dim CN As New ADODB.Connection
        Dim Rs As New BaseClass.BaseODBCIO
        CN.Open(ConnVAT)
        Rs.Open("select INGETVAT,USERGET,RESTR from fmsvset", CN)
        If Rs.options.QueryDT.Rows.Count > 0 Then
            If Rs.options.QueryDT.Rows(0).Item("INGETVAT").ToString.Trim = "1" Then
                If Rs.options.QueryDT.Rows(0).Item("RESTR").ToString.Trim = "1" Then
                    If MessageBox.Show(Rs.options.QueryDT.Rows(0).Item("USERGET").ToString.Trim & " process vat fail,Plase Retry Process Vat.", "Plase Retry Process Vat", MessageBoxButtons.YesNo, MessageBoxIcon.Stop) = Windows.Forms.DialogResult.Yes Then
                        CN.Execute("Update " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "FMSVSET SET INGETVAT=1,USERGET='" & UserLogin & "',DISCONNECTUSER=0,RESTR=0")
                    Else
                        Rs = Nothing
                        CN.Close()
                        CN = Nothing
                        Exit Sub
                    End If
                Else
                    MessageBox.Show("Cannot process vat , " & Rs.options.QueryDT.Rows(0).Item("USERGET").ToString.Trim & " is Processing.", "Just a Moment Please. ...", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                    Rs = Nothing
                    CN.Close()
                    CN = Nothing
                    Exit Sub
                End If
            Else
                CN.Execute("Update " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "FMSVSET SET INGETVAT=1,USERGET='" & UserLogin & "',DISCONNECTUSER=0,RESTR=0")
            End If
        End If
        Dim frmData As frmProcess
        StrYear = FormatDate(Trim(txtFrom.Value.ToString("yyyyMMdd")))
        StrYearF = Trim(txtFrom.Value.ToString("yyyyMMdd"))

        frmData = New frmProcess
        DATEFROM = txtFrom.Value.ToString("yyyyMMdd")
        DATETO = txtTo.Value.ToString("yyyyMMdd")
        INGETVAT = True
        frmData.Process = 1
        frmData.ShowDialog(Me)
        GetVatInsert()
        GetVatlist()
        ListShow()
        INGETVAT = False
        CN.Execute("Update " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "FMSVSET SET INGETVAT=0,USERGET='',STATUS=0 ,DISCONNECTUSER=0,RESTR=0")
        Status.Text = "Time Process : " & frmData.Timmer.Text
        ToolStripProgressBar1.Visible = False
        frmData = Nothing

        WriteLog("<ProcessComplete> GetVat Complete")
        clsFuntion.FlushMemory()
        Rs = Nothing
        CN.Close()
        CN = Nothing


        Exit Sub
ErrHandler:
        INGETVAT = False
        WriteLog("<GetVat>" & Err.Number & "  " & Err.Description)
        If Len(Err.Description) > 0 Then MessageBox.Show(Err.Description, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    Private Sub btnSetLocation_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnSetLocation.Click
        Dim frmData As New FrmLocation
        frmData.ShowDialog(Me)
        RefreshLocation()
        frmData = Nothing
        Exit Sub

ErrHandler:
        WriteLog("<GetLocation>" & Err.Number & "  " & Err.Description)
        If Len(Err.Description) > 0 Then MessageBox.Show(Err.Description, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    Private Sub btnSetTaxNo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnSetTaxNo.Click
        Dim frmData As frmCompany
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim iffs As clsFuntion
        iffs = New clsFuntion

        On Error GoTo ErrHandler
        'Mouse(MON)
        System.Windows.Forms.Cursor.Current = Cursors.AppStarting

        WriteLog("GetCompany")
        frmData = New frmCompany
        cn = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO

        sqlstr = "SELECT ORGID,CONAME,ADDR01,ADDR02,ADDR03,ADDR04,"
        sqlstr = sqlstr & " CITY,STATE,POSTAL,COUNTRY,"
        sqlstr = sqlstr & " PHONE,FAX FROM CSCOM"
        cn.Open(ConnACCPAC)

        rs.Open(sqlstr, cn)
        Do While rs.options.Rs.EOF = False
            With frmData
                .CONAME = IIf(iffs.ISNULL(rs.options.Rs, 1), "", rs.options.Rs.Collect(1)) 'CSCOM.CONAME
                .ADDR01 = IIf(iffs.ISNULL(rs.options.Rs, 2), "", rs.options.Rs.Collect(2)) 'CSCOM.ADDR01
                .ADDR02 = IIf(iffs.ISNULL(rs.options.Rs, 3), "", rs.options.Rs.Collect(3)) 'CSCOM.ADDR02
                .ADDR03 = IIf(iffs.ISNULL(rs.options.Rs, 4), "", rs.options.Rs.Collect(4)) 'CSCOM.ADDR03
                .ADDR04 = IIf(iffs.ISNULL(rs.options.Rs, 5), "", rs.options.Rs.Collect(5)) 'CSCOM.ADDR04
                .CITY = IIf(iffs.ISNULL(rs.options.Rs, 6), "", rs.options.Rs.Collect(6)) 'CSCOM.CITY
                .STATE = IIf(iffs.ISNULL(rs.options.Rs, 7), "", rs.options.Rs.Collect(7)) 'CSCOM.STATE
                .POSTAL = IIf(iffs.ISNULL(rs.options.Rs, 8), "", rs.options.Rs.Collect(8)) 'CSCOM.POSTAL
                .COUNTRY = IIf(iffs.ISNULL(rs.options.Rs, 9), "", rs.options.Rs.Collect(9)) 'CSCOM.COUNTRY
            End With
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        'Mouse(MOFF)
        System.Windows.Forms.Cursor.Current = Cursors.Default

        frmData.ShowDialog()

        'Mouse(MON)
        System.Windows.Forms.Cursor.Current = Cursors.AppStarting

        If frmData.ClickSelect = True Then
            WriteLog("SaveVset")
            SaveVatset()
        End If
        'Mouse(MOFF)
        System.Windows.Forms.Cursor.Current = Cursors.Default

        frmData = Nothing
        rs = Nothing
        cn = Nothing
        Exit Sub

ErrHandler:
        WriteLog("<GetCompany>" & Err.Number & "  " & Err.Description)
        If Len(Err.Description) > 0 Then MessageBox.Show(Err.Description, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    Dim isTrue As Boolean = False
    Private Sub lstVu_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles lstVu.ColumnHeaderMouseClick

        If e.Button = Windows.Forms.MouseButtons.Left Then
            lstVu.Rows.Clear()
            Dim ColumnName As String = lstVu.Columns(e.ColumnIndex).Name
            If e.ColumnIndex = 2 Then
                ColumnName = "INVDATE"
            End If

            If strReOrder = ColumnName Then
                lbOrder = Not lbOrder
            Else
                lbOrder = False
                strReOrder = ColumnName
            End If
            strOrderBy = strReOrder & IIf(lbOrder, " DESC", " ASC")
            btnRefreh_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub lstVu_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lstVu.MouseDown
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        If Button = 2 Then
        End If
    End Sub

    Public Sub mnuViewRefresh_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        lstVu.Refresh()
    End Sub

    Private Sub txtFrom_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFrom.ValueChanged
        CheckVatEnabled()
    End Sub
    Private Sub txtFrom_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtFrom.MouseWheel
        If txtFrom.Focused Then
            If e.Delta > 0 Then
                SendKeys.Send("{UP}")
            Else
                SendKeys.Send("{DOWN}")
            End If
        End If
    End Sub

    Private Sub txtTo_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles txtTo.MouseWheel
        If txtTo.Focused Then
            If e.Delta > 0 Then
                SendKeys.Send("{UP}")
            Else
                SendKeys.Send("{DOWN}")
            End If
        End If
    End Sub

    Private Sub txtFrom_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFrom.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        KeyAscii = OnlyNumericKey(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtRate_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRate.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            ListShow()
        Else
            If KeyAscii <> 46 Then KeyAscii = OnlyNumericKey(KeyAscii)
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTo.ValueChanged
        CheckVatEnabled()
    End Sub

    Private Sub txtTo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        KeyAscii = OnlyNumericKey(KeyAscii)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub frmMain_Move(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Move
        If frmOpasity Is Nothing Then Exit Sub
        frmOpasity.Location = New Point(Me.Location.X + 1, Me.Location.Y + 20) 'Me.Location '(New Point(lstVu.Location.X, lstVu.Location.Y))
        frmOpasity.Size = New Size(Me.Width - 2, Me.Height - 21)  ' lstVu.Size
    End Sub

    Private Sub frmMain_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Me.Refresh()
        If Me.WindowState <> 1 Then      'เช็คว่าฟอร์มไม่อยู่ในสถานะ Minimize
            'กำหนดให้ฟอร์มมีขนาดเล็กได้ไม่เกินเท่าไหร่
            'If Me.Width < 10905 Then Me.Width = 10905
            'If Me.Height < 7650 Then Me.Height = 7650
            'If Me.Width < Screen.PrimaryScreen.WorkingArea.Width Then
            '    Me.Width = Screen.PrimaryScreen.WorkingArea.Width - 20
            lstVu.Width = Me.Width - 23
            lstVu.Height = Me.Height - 190
            GroupBox1.Location = New Point(3, Me.Height - 115)
            GroupBox1.Width = Me.Width - 23
        End If

        If frmOpasity Is Nothing Then Exit Sub
        frmOpasity.Location = New Point(Me.Location.X + 1, Me.Location.Y + 20) 'Me.Location '(New Point(lstVu.Location.X, lstVu.Location.Y))
        frmOpasity.Size = New Size(Me.Width - 2, Me.Height - 21)  ' lstVu.Size

    End Sub

    Private Sub lstVu_CellMouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles lstVu.CellMouseDoubleClick
        If lstVu.Rows.Count > 0 And e.RowIndex > -1 And e.ColumnIndex > -1 Then
            btnEditVat_Click(btnEditVat, New System.EventArgs())
        End If
    End Sub

    Private Sub lstVu_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lstVu.KeyDown
        If e.KeyCode = Keys.Delete Then
            m_popedContainerForButton_DeleteVat()
        End If
    End Sub

    Private Sub btnDBSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDBSet.Click
        'Dim frm As New FrmConfigDB("")
        frmConfig.ShowDialog(Me)
    End Sub

    Private Sub btnUserSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUserSet.Click
        Dim frm As New FrmConfigUser
        frm.ShowDialog()
    End Sub

    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Loadcomplate = False Then Exit Sub
        If MessageBox.Show("Do you really want to exit?", My.Application.Info.ProductName & " " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build, MessageBoxButtons.YesNo) <> Windows.Forms.DialogResult.Yes Then
            e.Cancel = True
        End If
    End Sub

    Private Sub lstVu_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles lstVu.CellContentClick
        If e.ColumnIndex = 1 And e.RowIndex > -1 Then
            lstVu.EndEdit()
            Dim Value As Integer = Math.Abs(CInt(CBool(lstVu.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)))
            Dim cn As ADODB.Connection
            Dim sqlstr As String
            Dim rs As BaseClass.BaseODBCIO
            cn = New ADODB.Connection

            sqlstr = "Select * from FMSVATINSERT"
            sqlstr &= " WHERE LOCID='" & lstVu.Rows(e.RowIndex).Cells("LOCID").Value & "' AND SOURCE='" & lstVu.Rows(e.RowIndex).Cells("SOURCE").Value & "' "
            sqlstr &= " AND BATCH='" & lstVu.Rows(e.RowIndex).Cells("BATCH").Value & "' AND ENTRY='" & lstVu.Rows(e.RowIndex).Cells("ENTRY").Value & "' AND TRANSNBR='" & (mocolVatDisPlay.Item(lstVu.Rows(e.RowIndex).Cells("PK").Value).TRANSNBR) & "'"
            cn = New ADODB.Connection
            rs = New BaseClass.BaseODBCIO
            cn.Open(ConnVAT)
            rs.Open(sqlstr, cn)
            If rs.options.Rs.RecordCount > 0 Then
                'Mouse(MON)
                System.Windows.Forms.Cursor.Current = Cursors.AppStarting

                With lstVu.Rows(e.RowIndex)
                    '*********************************************************************************************
                    '               บันทึกรายการ VAT ที่มีการแก้ไข เพื่อใช้กรณี Print ย้อนหลัง
                    '*********************************************************************************************
                    cn = New ADODB.Connection
                    rs = New BaseClass.BaseODBCIO

                    sqlstr = " Select * from FMSVATINSERT  WHERE (SOURCE = '" & .Cells("SOURCE").Value & "')  "
                    sqlstr = sqlstr & " AND (BATCH ='" & .Cells("Batch").Value & "') AND (ENTRY = '" & .Cells("Entry").Value & "') AND (LOCID = '" & .Cells("LOCID").Value & "') AND (TRANSNBR='" & mocolVatDisPlay.Item(.Cells("PK").Value).TRANSNBR & "')"

                    cn.Open(ConnVAT)
                    rs.Open(sqlstr, cn)
                    If rs.options.Rs.RecordCount > 0 And lstVu.Rows(e.RowIndex).Cells("Verify").Value = False Then
                        sqlstr = "delete from FMSVATINSERT WHERE (SOURCE = '" & .Cells("SOURCE").Value & "')  "
                        sqlstr &= " AND (BATCH ='" & .Cells("Batch").Value & "') AND (ENTRY = '" & .Cells("Entry").Value & "') AND (LOCID = '" & .Cells("LOCID").Value & "') AND (TRANSNBR='" & mocolVatDisPlay.Item(.Cells("PK").Value).TRANSNBR & "');"

                        sqlstr &= " Update FMSVAT SET VERIFY=0 WHERE (SOURCE = '" & .Cells("SOURCE").Value & "')  "
                        sqlstr &= " AND (BATCH ='" & .Cells("Batch").Value & "') AND (ENTRY = '" & .Cells("Entry").Value & "') AND (LOCID = '" & .Cells("LOCID").Value & "') AND (TRANSNBR='" & mocolVatDisPlay.Item(.Cells("PK").Value).TRANSNBR & "');"
                        lstVu.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.White
                    ElseIf rs.options.Rs.RecordCount > 0 Then
                        sqlstr = UpEditVat(CDec(mocolVatDisPlay.Item(.Cells("PK").Value).VATID), mocolVatDisPlay.Item(.Cells("PK").Value).TXDATE.ToString, .Cells("IDINVC").Value.ToString, .Cells("DOCNO").Value.ToString, .Cells("NEWDOCNO").Value.ToString, .Cells("INVNAME").Value.ToString, .Cells("INVAMT").Value.ToString, .Cells("INVTAX").Value.ToString, .Cells("LOCID").Value.ToString, .Cells("RATE").Value.ToString, mocolVatDisPlay.Item(.Cells("PK").Value).VTYPE.ToString, mocolVatDisPlay.Item(.Cells("PK").Value).TTYPE.ToString, mocolVatDisPlay.Item(.Cells("PK").Value).VATCOMMENT.ToString, .Cells("CBREF").Value.ToString, mocolVatDisPlay.Item(.Cells("PK").Value).INVDATE.ToString, .Cells("BATCH").Value.ToString, .Cells("ENTRY").Value.ToString, .Cells("SOURCE").Value.ToString, mocolVatDisPlay.Item(.Cells("PK").Value).TRANSNBR.ToString, .Cells("RUNNO").Value.ToString, .Cells("TypeofPU").Value.ToString, .Cells("TRANNO").Value.ToString, .Cells("CIF").Value.ToString, .Cells("TaxCIF").Value.ToString, Value, .Cells("TAXID").Value.ToString, .Cells("BRANCH").Value.ToString)
                        lstVu.Rows(e.RowIndex).DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 255, 255, 130)
                    Else
                        sqlstr = InsNewVat(CDec(mocolVatDisPlay.Item(.Cells("PK").Value).VATID), mocolVatDisPlay.Item(.Cells("PK").Value).INVDATE.ToString, .Cells("TXDATE").Value.ToString, .Cells("IDINVC").Value.ToString, .Cells("DOCNO").Value.ToString, .Cells("NEWDOCNO").Value.ToString, .Cells("INVNAME").Value.ToString, .Cells("INVAMT").Value, .Cells("INVTAX").Value.ToString, .Cells("LOCID").Value.ToString, mocolVatDisPlay.Item(.Cells("PK").Value).VTYPE.ToString, .Cells("RATE").Value.ToString, mocolVatDisPlay.Item(.Cells("PK").Value).TTYPE.ToString, mocolVatDisPlay.Item(.Cells("PK").Value).ACCTVAT.ToString, .Cells("SOURCE").Value.ToString, .Cells("BATCH").Value.ToString, .Cells("ENTRY").Value.ToString, mocolVatDisPlay.Item(CInt(.Cells("PK").Value)).MARK.ToString, mocolVatDisPlay.Item(.Cells("PK").Value).VATCOMMENT.ToString, .Cells("CBREF").Value.ToString, "Edit", mocolVatDisPlay.Item(.Cells("PK").Value).TRANSNBR.ToString, .Cells("RUNNO").Value.ToString, .Cells("TYPEOFPU").Value.ToString, .Cells("TRANNO").Value.ToString, .Cells("CIF").Value.ToString, .Cells("TAXCIF").Value.ToString, "", "", Value, .Cells("TAXID").Value.ToString, .Cells("BRANCH").Value.ToString, .Cells("CURRENCY").Value.ToString, .Cells("EXCHANGRATE").Value, .Cells("AMTBASETC").Value)
                        lstVu.Rows(e.RowIndex).DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 255, 255, 130)

                    End If

                    cn = New ADODB.Connection
                    cn.Open(ConnVAT)
                    cn.Execute(sqlstr)
                    cn.Close()

                    ' WriteLog
                    WriteLog("Save Data Success")
                    mocolVatDisPlay.Item(lstVu.Rows(e.RowIndex).Cells("PK").Value).Verify = Value
                End With

            ElseIf Value = 1 Then
                With lstVu.Rows(e.RowIndex)
                    Dim a As String = lstVu.Rows(e.RowIndex).Cells("PK").Value
                    cn = New ADODB.Connection
                    rs = New BaseClass.BaseODBCIO

                    '*********************************************************************************************
                    '               บันทึกรายการ VAT ที่มีการ เพิ่ม รายการ
                    '*********************************************************************************************

                    sqlstr = InsNewVat(CDec(mocolVatDisPlay.Item(.Cells("PK").Value).VATID), mocolVatDisPlay.Item(.Cells("PK").Value).INVDATE.ToString, mocolVatDisPlay.Item(.Cells("PK").Value).TXDATE.ToString, .Cells("IDINVC").Value.ToString, .Cells("DOCNO").Value.ToString, .Cells("NEWDOCNO").Value.ToString, .Cells("INVNAME").Value.ToString, .Cells("INVAMT").Value.ToString, .Cells("INVTAX").Value.ToString, .Cells("LOCID").Value.ToString, mocolVatDisPlay.Item(.Cells("PK").Value).VTYPE.ToString, .Cells("RATE").Value.ToString, mocolVatDisPlay.Item(.Cells("PK").Value).TTYPE.ToString, mocolVatDisPlay.Item(.Cells("PK").Value).ACCTVAT.ToString, .Cells("SOURCE").Value.ToString, .Cells("BATCH").Value.ToString, .Cells("ENTRY").Value.ToString, mocolVatDisPlay.Item(.Cells("PK").Value).MARK.ToString, mocolVatDisPlay.Item(.Cells("PK").Value).VATCOMMENT.ToString, .Cells("CBREF").Value.ToString, "Edit", mocolVatDisPlay.Item(.Cells("PK").Value).TRANSNBR.ToString, .Cells("RUNNO").Value.ToString, .Cells("TYPEOFPU").Value.ToString, .Cells("TRANNO").Value.ToString, .Cells("CIF").Value.ToString, .Cells("TAXCIF").Value.ToString, "", "", Value, .Cells("TAXID").Value.ToString, .Cells("BRANCH").Value.ToString, .Cells("CURRENCY").Value.ToString, .Cells("EXCHANGRATE").Value, .Cells("AMTBASETC").Value) & Environment.NewLine
                    sqlstr &= "UPDATE FMSVAT SET VERIFY=1 WHERE VATID='" & CDec(mocolVatDisPlay.Item(.Cells("PK").Value).VATID) & "' and BATCH ='" & .Cells("BATCH").Value & "' and ENTRY='" & .Cells("ENTRY").Value & "' and TRANSNBR ='" & mocolVatDisPlay.Item(.Cells("PK").Value).TRANSNBR & "' and IDINVC='" & .Cells("IDINVC").Value & "'"
                    cn = New ADODB.Connection
                    cn.Open(ConnVAT)
                    cn.Execute(sqlstr)
                    cn.Close()
                    lstVu.Rows(e.RowIndex).DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 255, 255, 130)
                    WriteLog("Save Data Success")
                    mocolVatDisPlay.Item(.Cells("PK").Value).Verify = Value
                End With
            End If
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub BTPTotop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToTop.Click
        cboPage.SelectedIndex = 0
    End Sub

    Private Sub BTPToPre_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToPre.Click
        cboPage.SelectedIndex = cboPage.SelectedIndex - 1
    End Sub

    Private Sub CBPage_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPage.SelectedIndexChanged

        If cboPage.SelectedIndex = 0 Then
            btnToTop.Enabled = False
            btnToPre.Enabled = False
        Else
            btnToTop.Enabled = True
            btnToPre.Enabled = True
        End If

        If cboPage.SelectedIndex = cboPage.Items.Count - 1 Then
            btnToButtom.Enabled = False
            btnToNex.Enabled = False
        Else
            btnToButtom.Enabled = True
            btnToNex.Enabled = True
        End If
        HeaderCheckBox.Checked = False

        If mocolVat.Count = 0 Then
            mocolVatDisPlay = New colVat
            lstVu.Rows.Clear()
            Exit Sub
        End If

        mocolVatDisPlay = New colVat
        For Each TmpCls As clsVat In mocolVat
            Dim Temp_txtRate As String = txtRate.Text.Replace(">", "").Replace(">", "")
            If TmpCls.RATE_Renamed = 0 Then
                Dim aaaa As String = ""
            End If
            If cboLoc.SelectedIndex <> 0 And (Trim(TmpCls.LOCID) <> Trim(cboLoc.Text)) Then GoTo FillNextVat0
            If cboRate.SelectedIndex = 1 And TmpCls.VTYPE = 1 And TmpCls.RATE_Renamed <> IIf(Len(txtRate.Text) = 0, 0, txtRate.Text) Then GoTo FillNextVat0
            If cboRate.SelectedIndex = 2 And TmpCls.VTYPE <> 0 Then GoTo FillNextVat0
            If cboRate.SelectedIndex = 3 And (TmpCls.RATE_Renamed = 7 Or TmpCls.RATE_Renamed = 0) Then GoTo FillNextVat0 ') Then GoTo FillNextVat0
            mocolVatDisPlay.Add(TmpCls)
FillNextVat0:
        Next

        lstVu.Rows.Clear()
        Dim i As Integer = (CInt(cboPage.Text) - 1) * (Use_DisplayPage * 10)
        Dim pageFrom As Integer = (CInt(cboPage.Text) - 1) * (Use_DisplayPage * 10)
        Dim pageto As Integer = (CInt(cboPage.Text)) * (Use_DisplayPage * 10) - 1
        If Use_DisplayPage = 15 Then
            pageFrom = 0
            pageto = mocolVat.Count
        End If

        Dim cn As New ADODB.Connection
        Dim rs As New BaseClass.BaseODBCIO
        Dim str As String = ""
        str = " select FMSVAT.VATID,FMSVAT.LOCID,FMSVAT.SOURCE,FMSVAT.BATCH,FMSVAT.ENTRY,FMSVAT.TRANSNBR,COUNT(*)as CNT "
        str &= " from FMSVAT inner join FMSVATINSERT on  "
        str &= " (FMSVAT.LOCID=FMSVATINSERT .LOCID) and "
        str &= " (FMSVAT.SOURCE=FMSVATINSERT .SOURCE) and "
        str &= " (FMSVAT.BATCH =FMSVATINSERT .BATCH )and  "
        str &= " (FMSVAT .ENTRY =FMSVATINSERT .ENTRY)and "
        str &= " (FMSVAT.TRANSNBR=FMSVATINSERT.TRANSNBR) "
        str &= " group by FMSVAT.VATID,FMSVAT.LOCID,FMSVAT.SOURCE,FMSVAT.BATCH,FMSVAT.ENTRY,FMSVAT.TRANSNBR "
        cn.Open(ConnVAT)
        rs.Open(str, cn)

        For loVat_Count As Integer = pageFrom To pageto
            If loVat_Count > mocolVatDisPlay.Count - 1 Then Exit For
            Dim loVat_ As clsVat = mocolVatDisPlay.Item(loVat_Count + 1)
            Dim Temp_txtRate As String = txtRate.Text.Replace(">", "").Replace(">", "")
            If cboLoc.SelectedIndex <> 0 And (Trim(loVat_.LOCID) <> Trim(cboLoc.Text)) Then
                GoTo FillNextVat
            End If
            If cboRate.SelectedIndex = 1 And loVat_.VTYPE = 1 And loVat_.RATE_Renamed <> IIf(Len(txtRate.Text) = 0, 0, txtRate.Text) Then
                GoTo FillNextVat
            End If
            If cboRate.SelectedIndex = 2 And loVat_.VTYPE <> 0 Then
                GoTo FillNextVat
            End If
            If cboRate.SelectedIndex = 3 And (loVat_.RATE_Renamed = 7 Or loVat_.RATE_Renamed = 0) Then ') Then
                GoTo FillNextVat
            End If
            i = i + 1
            Call FillList(loVat_, i, rs.options.QueryDT)
FillNextVat:
        Next
        System.Windows.Forms.Cursor.Current = Cursors.Default
        lstVu.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)

        If Math.Ceiling(mocolVatDisPlay.Count / (Use_DisplayPage * 10)) < cboPage.Items.Count Then
            For count As Integer = Math.Ceiling(mocolVatDisPlay.Count / (Use_DisplayPage * 10)) To cboPage.Items.Count - 1
                cboPage.Items.RemoveAt(cboPage.Items.Count - 1)
            Next
        End If

        If cboPage.SelectedIndex = 0 Then
            btnToTop.Enabled = False
            btnToPre.Enabled = False
        Else
            btnToTop.Enabled = True
            btnToPre.Enabled = True
        End If

        If cboPage.SelectedIndex = cboPage.Items.Count - 1 Then
            btnToButtom.Enabled = False
            btnToNex.Enabled = False
        Else
            btnToButtom.Enabled = True
            btnToNex.Enabled = True
        End If

        If mocolVatDisPlay.Count = 0 Then
            btnToButtom.Enabled = False
            btnToNex.Enabled = False
            btnToTop.Enabled = False
            btnToPre.Enabled = False
        End If

        Exit Sub
ErrHandler:
        WriteLog("<ListShow>" & Err.Number & "  " & Err.Description)
        If Len(Err.Description) > 0 Then MessageBox.Show(Err.Description, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    Private Sub BTPTONex_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToNex.Click
        cboPage.SelectedIndex = cboPage.SelectedIndex + 1

    End Sub

    Private Sub BTPtoButtom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToButtom.Click
        cboPage.SelectedIndex = Math.Ceiling(mocolVatDisPlay.Count / (Use_DisplayPage * 10)) - 1
    End Sub

    Private Sub BT_Help_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHelp.Click
        Try
            Dim frm As New BaseClass.frmHelp(Application.StartupPath & "\HELP\User Manual - VatReport.pdf", "pdf")
            frm.Text = My.Application.Info.ProductName & " Help"
            frm.Show()
        Catch ex As Exception
            Dim frm As New BaseClass.frmHelp(Application.StartupPath & "\HELP\User Manual - VatReport.mht", "mht")
            frm.Show()
        End Try
    End Sub

    Public FrmFinderText As FrmFinder
    Private Sub lstVu_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles lstVu.KeyPress
        If e.KeyChar = ""c Then
            ShowFinder()
        End If
    End Sub
    Public Sub ShowFinder()
        Try
            If FrmFinderText Is Nothing Then
                FrmFinderText = New FrmFinder
                FrmFinderText.Show(Me)
                FrmFinderText.Location = Me.PointToScreen(New Point((Me.Width / 2) - (FrmFinderText.Width / 2), (Me.Height / 2) - (FrmFinderText.Height / 2)))
            Else
                FrmFinderText.Activate()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub m_popedContainerForButton_DeleteRunning() Handles m_popedContainerForButton.DeleteRunning
        Dim cn As ADODB.Connection
        Dim sql As String
        Dim lbResult As Integer
        Try
            If lstVu.Rows.Count = 0 Then Exit Sub
            lbResult = MsgBox("Clear Running record?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, My.Application.Info.ProductName)
            If lbResult = 6 Then
                sql = "Truncate Table FMSRUN "
                cn = New ADODB.Connection
                cn.Open(ConnVAT)
                cn.Execute(sql)

                cn.Close()
            End If
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("Cannot be deleted" & Environment.NewLine & ex.Message, "Warning!!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub m_popedContainerForButton_DeleteVat() Handles m_popedContainerForButton.DeleteVat
        Dim cn As ADODB.Connection
        Dim sqlstr As String
        Dim sql As String
        Dim lbResult As Integer
        Dim loVATID, i As Integer
        Try
            If lstVu.Rows.Count = 0 Then Exit Sub
            loVATID = CDbl(lstVu.CurrentRow.Cells("PK").Value)
            i = CDbl(lstVu.CurrentRow.Cells(0).Value) - 1
            lbResult = MsgBox("Delete selected record?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, My.Application.Info.ProductName)
            If lbResult = 6 Then

                sql = "DELETE FROM FMSVATINSERT "
                sql = sql & " WHERE (BATCH ='" & Trim(mocolVatDisPlay.Item(loVATID).Batch) & "') AND (ENTRY = '" & Trim(mocolVatDisPlay.Item(loVATID).Entry) & "') AND (LOCID = '" & mocolVatDisPlay.Item(loVATID).LOCID & "') AND (TRANSNBR='" & Trim(mocolVatDisPlay.Item(loVATID).TRANSNBR) & "')"

                cn = New ADODB.Connection
                cn.Open(ConnVAT)
                cn.Execute(sql)
                cn.Close()

                sqlstr = "DELETE FROM FMSVAT WHERE VATID="
                sqlstr = sqlstr & Trim(mocolVatDisPlay.Item(loVATID).VATID)
                cn = New ADODB.Connection
                cn.Open(ConnVAT)
                cn.Execute(sqlstr)
                cn.Close()
                cn = Nothing

                Application.DoEvents()
                If Loadcomplate Then
                    ShowLoading()
                End If
                Application.DoEvents()

                'Reload From
                GetVatInsert()
                GetVatlist()
                Application.DoEvents()
                CloseLoading()
                CBPage_SelectedIndexChanged(Nothing, Nothing)
                '

                Me.lstVu.Select()
                If i = mocolVatDisPlay.Count Then
                    i = i - 1
                End If
                If i > -1 Then
                    Dim ci As Integer = i \ (Use_DisplayPage * 10) + 1
                    i = i Mod (Use_DisplayPage * 10)
                    cboPage.Text = ci
                    Dim pageFrom As Integer = (CInt(cboPage.Text) - 1) * (Use_DisplayPage * 10)
                    Dim pageto As Integer = (CInt(cboPage.Text)) * (Use_DisplayPage * 10) - 1

                    lstVu.CurrentCell = lstVu.Rows(i).Cells(2)
                    Me.lstVu.CurrentRow.Selected = True
                Else
                    CheckVatEnabled()
                End If

            End If
            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.Message & "Please select the item to delete", "Warning!!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub m_popedContainerForButton_DeleteVatInsert() Handles m_popedContainerForButton.DeleteVatInsert
        Dim frm As New DeleteVat
        frm.ShowDialog(Me)
    End Sub

    Private Sub cmdDeleteVat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteVat.Click
        m_popedContainerForButton = New SuperContextMenu.ContextMenuForDeleteVat()
        m_poperContainerForButton = New SuperContextMenu.PoperContainer(m_popedContainerForButton)
        m_popedContainerForButton.EnDelete = (lstVu.Rows.Count > 0)
        m_poperContainerForButton.Show(btnDeleteVat)
        m_popedContainerForButton.btDeleteRun.Text = "Delete Runing"
        m_popedContainerForButton.btDeleteRun.TextAlign = ContentAlignment.MiddleCenter
        m_popedContainerForButton.btDeleteVat.Text = "Delete Vat"
        m_popedContainerForButton.btDeleteVat.TextAlign = ContentAlignment.MiddleCenter
        m_popedContainerForButton.btdeleteVatInsert.Text = "Delete VatInsert"
        m_popedContainerForButton.btdeleteVatInsert.TextAlign = ContentAlignment.MiddleCenter
    End Sub

    Private Sub lblCompany_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblCompany.Click
        If Control.ModifierKeys = Keys.Control Then
            Dim basec As New BaseClass.TaskManager
            basec.ShowDialog()
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            Dim cn As New ADODB.Connection
            Dim rs As New BaseClass.BaseODBCIO
            cn.ConnectionTimeout = 60
            cn.Open(ConnVAT)
            cn.CommandTimeout = 0

            rs.Open("SELECT ORGID, CONAME, TAXNO, DATEFROM, DATETO, ISNULL(INGETVAT, 0) AS INGETVAT, ISNULL(USERGET, '') AS USERGET, ISNULL(STATUS, 0) AS STATUS,ISNULL(DISCONNECTUSER, 0) AS DISCONNECTUSER, ISNULL(RESTR, 0) AS RESTR FROM FMSVSET", cn)
            If rs.options.QueryDT.Rows.Count > 0 Then
                If UserLogin.Trim = rs.options.QueryDT.Rows(0).Item("USERGET").ToString.Trim And rs.options.QueryDT.Rows(0).Item("INGETVAT").ToString.Trim = 1 And rs.options.QueryDT.Rows(0).Item("RESTR").ToString.Trim = 0 And INGETVAT = True Then
                    cn.Execute("Update " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "FMSVSET SET DISCONNECTUSER=DISCONNECTUSER+1 ")
                End If

                If rs.options.QueryDT.Rows(0).Item("INGETVAT").ToString.Trim = 1 And CountDis <> rs.options.QueryDT.Rows(0).Item("DISCONNECTUSER").ToString.Trim Then
                    CountDis = rs.options.QueryDT.Rows(0).Item("DISCONNECTUSER").ToString.Trim
                    DISCONNECTUSER = 0
                ElseIf rs.options.QueryDT.Rows(0).Item("INGETVAT").ToString.Trim = 1 And CountDis = rs.options.QueryDT.Rows(0).Item("DISCONNECTUSER").ToString.Trim Then
                    DISCONNECTUSER += 1
                End If

                If DISCONNECTUSER >= 5 Then
                    DISCONNECTUSER = 0
                    cn.Execute("Update " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "FMSVSET SET RESTR=1")
                    ToolStripProgressBar1.Visible = False
                    Status.Text = rs.options.QueryDT.Rows(0).Item("USERGET") & " process vat fail!"
                    Status.Image = Global.ProgramVAT.My.Resources.Resources.Shocked
                    ToolStripProgressBar1.Value = 0
                    TaskBar.Value = 0
                    TaskBar.ShowInTaskbar = False
                    rs = Nothing
                    cn.Close()
                    Exit Sub
                End If
                If rs.options.QueryDT.Rows(0).Item("INGETVAT").ToString.Trim = 0 And Status.Text <> "" Then
                    Status.Image = Nothing
                    ToolStripProgressBar1.Visible = False
                    If UserLogin.Trim <> Status.Text.Substring(0, Status.Text.IndexOf("Process")).Trim Then
                        If Status.Text.IndexOf("Time Process :") = -1 Then
                            Try
                                Dim fi As DateTimeFormatInfo = New DateTimeFormatInfo
                                fi.ShortDatePattern = "yyyyMMdd"
                                If IsDBNull(rs.options.QueryDT.Rows(0).Item("DATEFROM")) = False Then txtFrom.Value = DateTime.ParseExact(CStr(rs.options.QueryDT.Rows(0).Item("DATEFROM").ToString).Trim, "yyyyMMdd", fi)
                                If IsDBNull(rs.options.QueryDT.Rows(0).Item("DATETO")) = False Then txtTo.Value = DateTime.ParseExact(CStr(rs.options.QueryDT.Rows(0).Item("DATETO").ToString).Trim, "yyyyMMdd", fi)
                            Catch ex As Exception
                            End Try
                            Status.Text = Status.Text.Substring(0, Status.Text.IndexOf("Process")).Trim & " Process vat Sucess!"
                        End If
                    End If
                    ToolStripProgressBar1.Value = 0
                    TaskBar.Value = 0
                    TaskBar.ShowInTaskbar = False
                    rs = Nothing
                    cn.Close()
                    Application.DoEvents()
                    Exit Sub
                End If
                If rs.options.QueryDT.Rows(0).Item("INGETVAT").ToString.Trim = 1 And rs.options.QueryDT.Rows(0).Item("RESTR").ToString.Trim = 0 Then
                    Status.Image = Nothing
                    If IsDBNull(rs.options.QueryDT.Rows(0).Item("STATUS")) = True Then
                        Status.Text = rs.options.QueryDT.Rows(0).Item("USERGET") & " Process Vat " & Math.Round(((0 / 11) * 100), 0) & " %"
                    Else
                        Status.Text = rs.options.QueryDT.Rows(0).Item("USERGET") & " Process Vat " & Math.Round(((rs.options.QueryDT.Rows(0).Item("STATUS") / 11) * 100), 0) & " %"
                    End If
                    ToolStripProgressBar1.Visible = True
                    ToolStripProgressBar1.Maximum = 11
                    ToolStripProgressBar1.Minimum = 0

                    TaskBar.ShowInTaskbar = True
                    TaskBar.Maximum = 11
                    TaskBar.Minimum = 0
                    If IsDBNull(rs.options.QueryDT.Rows(0).Item("STATUS")) = False And rs.options.QueryDT.Rows(0).Item("STATUS") <= ToolStripProgressBar1.Maximum Then
                        ToolStripProgressBar1.Value = rs.options.QueryDT.Rows(0).Item("STATUS")
                    Else
                        ToolStripProgressBar1.Value = 0
                    End If
                    TaskBar.Value = ToolStripProgressBar1.Value()
                End If

            End If
            rs = Nothing
            cn.Close()
            cn = Nothing

            Application.DoEvents()
        Catch ex As Exception
            Status.Text = "Can not Conntct to Server"
            Status.Image = Global.ProgramVAT.My.Resources.Resources.Shocked
            ToolStripProgressBar1.Value = 0
            TaskBar.Value = 0
            TaskBar.ShowInTaskbar = False
        End Try

    End Sub
    Private Sub Loading_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Loading.Tick
        Select Case (tic Mod 4)
            Case 0
                Tex.Text = "LOADING   "
            Case 1
                Tex.Text = "LOADING.  "
            Case 2
                Tex.Text = "LOADING.. "
            Case 3
                Tex.Text = "LOADING..."
        End Select
        tic += 1
    End Sub

   
End Class