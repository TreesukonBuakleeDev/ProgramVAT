Option Strict Off
Option Explicit On
Imports System.Data
Imports System.Data.OleDb
Imports ADODB
Imports System.Globalization
Imports System.Data.SqlClient

Module modVatReport
    Public StrSS As String
    Public ComID As String
    Public ComDatabase As String
    Public ComVatPartial As Boolean
    Public ComVatName As Boolean
    Public VATAutoRun As Boolean
    Public VATFormat As String
    Public DATEMODE As String

    Public ComACCPAC As String
    Public ComVAT As String
    Public Code As String = ""

    Public OPT_AP As Boolean
    Public OPT_CB As Boolean
    Public OPT_GL As Boolean

    Public VNO_AP As Boolean
    Public VNO_CB As Boolean
    Public VNO_AR As Boolean
    Public VNO_GL As Boolean

    Public RERUN As Boolean
    Public UPRUN As Boolean
    Public INGETVAT As Boolean
    Public FormatRunnig As String
    Public Const lAccpac As Short = 1
    Public Const lVat As Short = 0
    Public Const REPORT_PATH As String = "\FMSRPT"
    Private Const LogFile As String = "VatLog.txt"
    Private Const InsMSSQL As String = "INSERT INTO "
    Private Const InsPVSW As String = "INSERT INTO "
    Public InsStr As String

    Public ConnACCPAC As String
    Public ConnVAT As String

    Public ConAcc As String
    Public ConVAT As String

    Public ReportChk As String

    Public gVatId As Integer
    Public ORGID As String
    Public TAXNO As String
    Public StrYear As String
    Public StrYearF As String
    Public DATEFROM As String
    Public DATETO As String
    Public mocolLocation As colLocation
    Public mocolVat As colVat
    Public mocolVatDisPlay As New colVat
    Public mocolVatIN As colVat
    Public mocolVatOut As colVat
    Public mocolVatEx As colVat
    Private VATID As Integer
    Private ErrMsg As String
    'Public RateAVG As String
    Public AMTEXTNDHC As Decimal
    Public AMTGLDIST As Decimal
    Public frmpro As frmProcess
    Dim daNB As SqlDataAdapter
    Dim dsNB As DataSet
    Dim dr As SqlDataReader
    Dim com As SqlCommand
    Dim dtselect As New DataTable
    Dim dtTemp As New DataTable
    Public fmsadmin As String = ""
    Public frmConfig As New FrmConfigDB("")
    'Public frmConfigNew As New FrmConfigDB("New")


    Public ConA As SqlConnection
    Public ConV As SqlConnection
    Public ClickSelect1 As Boolean


    '----------------------------Read INI-----------------------------

    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal sSectionName As String, ByVal sKeyName As String, ByVal sDefault As String, ByVal sReturnedString As String, ByVal lSize As Integer, ByVal sFileName As String) As Integer

    '----------------------------Write INI-----------------------------

    Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal sSectionName As String, ByVal sKeyName As String, ByVal sString As String, ByVal sFileName As String) As Integer

    Public DBType As String = String.Empty

    Function GetPrivateProfileString(ByVal sSectionName As String, ByVal sKeyName As String, ByVal sDefault As String) As String
        Dim lProfile As Integer
        Dim sFile As String
        Dim sString As String = New String(" ", 255)
        sFile = My.Application.Info.DirectoryPath & "\FMSVAT.INI"
        lProfile = GetPrivateProfileString(sSectionName, sKeyName, sDefault, sString, Len(sString), sFile)
        If lProfile = 0 Then
            sString = sDefault
        End If
        sString = Trim(sString).Split("")(0)
        Return sString
    End Function

    Function GetPrivateProfile(ByVal sSectionName As String, ByVal sKeyName As String, ByVal sDefault As String) As String
        Dim lProfile As Integer
        Dim sFile As String
        Dim sString As String
        Dim SSValue As String
        sFile = My.Application.Info.DirectoryPath & "\FMSVAT.INI"
        sString = New String(" ", 255)
        lProfile = GetPrivateProfileString(sSectionName, sKeyName, sDefault, sString, Len(sString), sFile)
        If lProfile = 0 Then
            sString = sDefault
        End If

        sString = Trim(sString).Split("")(0)
        If sKeyName = "DATABASE" Then
            DBType = sString
        End If
        SSValue = sString
        Select Case sKeyName
            Case "ACCPAC"
                If DBType = "PVSW;" Then
                    SSValue = BaseClass.AES_Decrypt(sString, "ABCDEFG")
                ElseIf DBType = "MSSQL;" Then
                    SSValue = BaseClass.AES_Decrypt(sString, "ABCDEFG")
                End If
            Case "VAT"
                If DBType = "PVSW;" Then
                    SSValue = BaseClass.AES_Decrypt(sString, "ABCDEFG")
                ElseIf DBType = "MSSQL;" Then
                    SSValue = BaseClass.AES_Decrypt(sString, "ABCDEFG")
                End If
        End Select
        GetPrivateProfile = Trim(SSValue)


    End Function
    Public Sub Main()

        Dim frmData As frmMain

        frmData = New frmMain

        System.Windows.Forms.Application.Run(frmData)
        frmData = Nothing
    End Sub

    Public Sub LoadDBScript()
        Dim strsql As String = ""

        Try

            '-- เอาไว้ Update Table VAT ให้เข้ากับ โปรเเกรมตัวใหม่ได้
            Dim cn As New ADODB.Connection
            Dim result As New Object
            '--ADD Verify เอาไว้จำว่า มีการกด Verify หรือเปล่า'
            strsql = " if not exists( select * from sys.columns where name='Verify' and object_id =(select object_id  from sys.tables where name='FMSVATINSERT')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVATINSERT') begin"
            strsql &= " ALTER TABLE dbo.FMSVATINSERT ADD "
            strsql &= " 	Verify decimal(1, 0) NULL "
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine

            strsql &= " if not exists( select * from sys.columns where name='Verify' and object_id =(select object_id  from sys.tables where name='FMSVAT')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVAT') begin"
            strsql &= " ALTER TABLE dbo.FMSVAT ADD "
            strsql &= " 	Verify decimal(1, 0) NULL "
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine
            '--END'

            '--ADD TypeOfPU '
            strsql &= " if not exists( select * from sys.columns where name='TypeOfPU' and object_id =(select object_id  from sys.tables where name='FMSVATTEMP')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVATTEMP') begin"
            strsql &= " ALTER TABLE dbo.FMSVATTEMP ADD "
            strsql &= " 	[TypeOfPU] [varchar](30) NULL "
            strsql &= " end ;" & Environment.NewLine
            strsql &= " end ;" & Environment.NewLine


            strsql &= " if not exists( select * from sys.columns where name='TypeOfPU' and object_id =(select object_id  from sys.tables where name='FMSVAT')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVAT') begin"
            strsql &= " ALTER TABLE dbo.FMSVAT ADD "
            strsql &= " 	[TypeOfPU] [varchar](30) NULL "
            strsql &= " end ;" & Environment.NewLine
            strsql &= " end; " & Environment.NewLine

            strsql &= " if not exists( select * from sys.columns where name='TypeOfPU' and object_id =(select object_id  from sys.tables where name='FMSVATINSERT')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVATINSERT') begin"
            strsql &= " ALTER TABLE dbo.FMSVATINSERT ADD "
            strsql &= " 	[TypeOfPU] [varchar](30) NULL "
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine
            '--END'

            '--ADD TranNo'
            strsql &= " if not exists( select * from sys.columns where name='TranNo' and object_id =(select object_id  from sys.tables where name='FMSVATTEMP')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVATTEMP') begin"
            strsql &= " ALTER TABLE dbo.FMSVATTEMP ADD "
            strsql &= " 	[TranNo] [varchar](100) NULL "
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine

            strsql &= " if not exists( select * from sys.columns where name='TranNo' and object_id =(select object_id  from sys.tables where name='FMSVAT')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVAT') begin"
            strsql &= " ALTER TABLE dbo.FMSVAT ADD "
            strsql &= " 	[TranNo] [varchar](100) NULL "
            strsql &= " end ;" & Environment.NewLine
            strsql &= " end ;" & Environment.NewLine

            strsql &= " if not exists( select * from sys.columns where name='TranNo' and object_id =(select object_id  from sys.tables where name='FMSVATINSERT')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVATINSERT') begin"
            strsql &= " ALTER TABLE dbo.FMSVATINSERT ADD "
            strsql &= " 	[TranNo] [varchar](100) NULL "
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine
            '--END'

            '--ADD CIF'
            strsql &= " if not exists( select * from sys.columns where name='CIF' and object_id =(select object_id  from sys.tables where name='FMSVATTEMP')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVATTEMP') begin"
            strsql &= " ALTER TABLE dbo.FMSVATTEMP ADD "
            strsql &= " 	[CIF] [decimal](19, 3) NULL "
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine

            strsql &= " if not exists( select * from sys.columns where name='CIF' and object_id =(select object_id  from sys.tables where name='FMSVAT')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVAT') begin"
            strsql &= " ALTER TABLE dbo.FMSVAT ADD "
            strsql &= " 	[CIF] [decimal](19, 3) NULL "
            strsql &= " end; " & Environment.NewLine
            strsql &= " end ;" & Environment.NewLine

            strsql &= " if not exists( select * from sys.columns where name='CIF' and object_id =(select object_id  from sys.tables where name='FMSVATINSERT')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVATINSERT') begin"
            strsql &= " ALTER TABLE dbo.FMSVATINSERT ADD "
            strsql &= " 	[CIF] [decimal](19, 3) NULL "
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine
            '--END'

            '--ADD TAXCIF'
            strsql &= " if not exists( select * from sys.columns where name='TAXCIF' and object_id =(select object_id  from sys.tables where name='FMSVATTEMP')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVATTEMP') begin"
            strsql &= " ALTER TABLE dbo.FMSVATTEMP ADD "
            strsql &= " 	[TAXCIF] [decimal](19, 3) NULL "
            strsql &= " end ;" & Environment.NewLine
            strsql &= " end ;" & Environment.NewLine

            strsql &= " if not exists( select * from sys.columns where name='TAXCIF' and object_id =(select object_id  from sys.tables where name='FMSVAT')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVAT') begin"
            strsql &= " ALTER TABLE dbo.FMSVAT ADD "
            strsql &= " 	[TAXCIF] [decimal](19, 3) NULL "
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine

            strsql &= " if not exists( select * from sys.columns where name='TAXCIF' and object_id =(select object_id  from sys.tables where name='FMSVATINSERT')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVATINSERT') begin"
            strsql &= " ALTER TABLE dbo.FMSVATINSERT ADD "
            strsql &= " 	[TAXCIF] [decimal](19, 3) NULL "
            strsql &= " end; " & Environment.NewLine
            strsql &= " end ;" & Environment.NewLine
            '--END'

            '-- ADD INGETVAT เอาไว้จำว่ากำลังดำเนินการ GET VAT'
            strsql &= " if not exists( select * from sys.columns where name='INGETVAT' and object_id =(select object_id  from sys.tables where name='FMSVSET')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVAT') begin"
            strsql &= " ALTER TABLE dbo.FMSVSET ADD "
            strsql &= " [INGETVAT] [decimal](1,0)"
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine


            '-- ADD INGETVAT เอาไว้จำว่าเครื่อง ไหนกำลัง GET VAT อยู่'
            strsql &= " if not exists( select * from sys.columns where name='USERGET' and object_id =(select object_id  from sys.tables where name='FMSVSET')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVSET') begin"
            strsql &= " ALTER TABLE dbo.FMSVSET ADD "
            strsql &= " [USERGET] [varchar](100)"
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine


            '-- ADD INGETVAT เอาไว้จำว่าเครื่อง ไหนกำลัง GET VAT อยู่'
            strsql &= " if not exists( select * from sys.columns where name='STATUS' and object_id =(select object_id  from sys.tables where name='FMSVSET')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVSET') begin"
            strsql &= " ALTER TABLE dbo.FMSVSET ADD "
            strsql &= " [STATUS] [decimal](9,0)"
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine

            strsql &= " if not exists( select * from sys.columns where name='DISCONNECTUSER' and object_id =(select object_id  from sys.tables where name='FMSVSET')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVSET') begin"
            strsql &= " ALTER TABLE dbo.FMSVSET ADD "
            strsql &= " [DISCONNECTUSER] [decimal](18,0)"
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine

            strsql &= " if not exists( select * from sys.columns where name='RESTR' and object_id =(select object_id  from sys.tables where name='FMSVSET')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVSET') begin"
            strsql &= " ALTER TABLE dbo.FMSVSET ADD "
            strsql &= " [RESTR] [decimal](1,0)"
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine


            strsql &= " if not exists( select * from sys.columns where name='TAXID' and object_id =(select object_id  from sys.tables where name='FMSVAT')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVAT') begin"
            strsql &= " ALTER TABLE dbo.FMSVAT ADD "
            strsql &= " [TAXID] VARCHAR(13)"
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine

            strsql &= " if not exists( select * from sys.columns where name='TAXID' and object_id =(select object_id  from sys.tables where name='FMSVATINSERT')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVATINSERT') begin"
            strsql &= " ALTER TABLE dbo.FMSVATINSERT ADD "
            strsql &= " [TAXID] VARCHAR(13)"
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine

            strsql &= " if not exists( select * from sys.columns where name='TAXID' and object_id =(select object_id  from sys.tables where name='FMSVATTEMP')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVATTEMP') begin"
            strsql &= " ALTER TABLE dbo.FMSVATTEMP ADD "
            strsql &= " [TAXID] VARCHAR(13)"
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine

            strsql &= " if not exists( select * from sys.columns where name='TAXID' and object_id =(select object_id  from sys.tables where name='FMSCB')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSCB') begin"
            strsql &= " ALTER TABLE dbo.FMSCB ADD "
            strsql &= " [TAXID] VARCHAR(13)"
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine

            strsql &= " if not exists( select * from sys.columns where name='BRANCH' and object_id =(select object_id  from sys.tables where name='FMSVAT')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVAT') begin"
            strsql &= " ALTER TABLE dbo.FMSVAT ADD "
            strsql &= " [BRANCH] VARCHAR(500)"
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine

            strsql &= " if not exists( select * from sys.columns where name='BRANCH' and object_id =(select object_id  from sys.tables where name='FMSVATINSERT')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVATINSERT') begin"
            strsql &= " ALTER TABLE dbo.FMSVATINSERT ADD "
            strsql &= " [BRANCH] VARCHAR(500)"
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine

            strsql &= " if not exists( select * from sys.columns where name='BRANCH' and object_id =(select object_id  from sys.tables where name='FMSVATTEMP')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVATTEMP') begin"
            strsql &= " ALTER TABLE dbo.FMSVATTEMP ADD "
            strsql &= " [BRANCH] VARCHAR(500)"
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine

            strsql &= " if not exists( select * from sys.columns where name='BRANCH' and object_id =(select object_id  from sys.tables where name='FMSCB')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSCB') begin"
            strsql &= " ALTER TABLE dbo.FMSCB ADD "
            strsql &= " [BRANCH] VARCHAR(500)"
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine

            strsql &= " if not exists( select * from sys.columns where name='TOTTAX' and object_id =(select object_id  from sys.tables where name='FMSVATTEMP')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVATTEMP') begin"
            strsql &= " ALTER TABLE dbo.FMSVATTEMP ADD "
            strsql &= " TOTTAX  DECIMAL(19,3)"
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine

            strsql &= " if not exists( select * from sys.columns where name='TOTTAX' and object_id =(select object_id  from sys.tables where name='FMSVAT')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVAT') begin"
            strsql &= " ALTER TABLE dbo.FMSVAT ADD "
            strsql &= " TOTTAX  DECIMAL(19,3)"
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine

            strsql &= " if not exists( select * from sys.columns where name='CODETAX' and object_id =(select object_id  from sys.tables where name='FMSVAT')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSVAT') begin"
            strsql &= " ALTER TABLE dbo.FMSVAT ADD "
            strsql &= " CODETAX  CHAR(12)"
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine


            strsql &= " if not exists( select * from sys.columns where name='TOTTAX' and object_id =(select object_id  from sys.tables where name='FMSCB')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSCB') begin"
            strsql &= " ALTER TABLE dbo.FMSCB ADD "
            strsql &= " TOTTAX  DECIMAL(19,3)"
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine


            strsql &= " if not exists( select * from sys.columns where name='RATERECOV' and object_id =(select object_id  from sys.tables where name='FMSTAX')) "
            strsql &= " begin "
            strsql &= " if exists(select *  from sys.tables where name='FMSTAX') begin"
            strsql &= " ALTER TABLE dbo.FMSTAX ADD "
            strsql &= " RATERECOV  DECIMAL(15,5)"
            strsql &= " end; " & Environment.NewLine
            strsql &= " end; " & Environment.NewLine


            cn = New ADODB.Connection
            cn.Open(ConnVAT)
            cn.Execute(strsql, result)
        Catch ex As Exception
            WriteLog("<LoadDBScript>" & ex.Message.ToString())
        End Try

    End Sub

    '--Build Connection String
    Public Function BuildConnection(ByRef ConnectTo As Integer, ByRef ComSetting As String, Optional ByRef frm As Form = Nothing) As Boolean

Jeditcomplete:
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim Connstring, sqlstr As String


        Dim strsplit As Object
        Dim Server, DBName As String
        Dim User, Password As String
        Dim i As Integer
        On Error GoTo ErrHandler

        cn = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        strsplit = Split(ComSetting, ";")
        Connstring = ""

        For i = LBound(strsplit) To UBound(strsplit)
            Select Case i
                Case Is = 0
                    If ComDatabase = "PVSW" Or (ComDatabase = "ORCL" And ConnectTo = 2) Then
                        Connstring = "Provider=MSDASQL.1;Persist Security Info=False;" & "Data Source=" & strsplit(i)

                    ElseIf ComDatabase = "MSSQL" Then

                        Connstring = "Provider=SQLOLEDB.1;server=" & strsplit(i)

                    ElseIf ComDatabase = "ORCL" Then
                        Connstring = "Provider=MSDAORA.1;Data Source=" & strsplit(i) & ";Persist Security Info=False"
                    End If
                    ConnDB = "Data Source=" & strsplit(i)

                Case Is = 1
                    If ComDatabase = "PVSW" Then
                        Connstring = Connstring
                    ElseIf ComDatabase = "MSSQL" Then
                        Connstring = Connstring & ";database=" & strsplit(i)
                    End If
                    ConnDB &= " ;Initial Catalog=" & strsplit(i)
                Case Is = 2
                    If ComDatabase = "MSSQL" Then
                        Connstring = Connstring & ";uid=" & strsplit(i)
                    ElseIf ComDatabase = "ORCL" Then
                        Connstring = Connstring & ";USER ID=" & strsplit(i)
                    End If
                    'ConnDB &= ";Persist Security Info=True;User ID=" & strsplit(i)
                    ConnDB &= ";User ID=" & strsplit(i)
                Case Is = 3
                    If ComDatabase = "MSSQL" Then
                        Connstring = Connstring & ";pwd=" & strsplit(i)
                    ElseIf ComDatabase = "ORCL" Then
                        Connstring = Connstring & ";PASSWORD=" & strsplit(i)
                    End If
                    ConnDB &= " ;Password=" & strsplit(i)
            End Select
        Next

        cn.Open(Connstring & ";Connect Timeout=3;")
        cn.Close()

        Select Case ConnectTo
            Case 1

                ConnACCPAC = Connstring
                ComACCPAC = ComSetting
                ConAcc = ConnDB
            Case 0
                ConnVAT = Connstring
                ComVAT = ComSetting
                ConVAT = ConnDB
        End Select
        WriteLog("Connection successfull")

        rs = Nothing
        cn = Nothing
        Return True
        Exit Function

ErrHandler:
        WriteLog(Err.Number & "  " & Err.Description)
        MsgBox("Program terminate." & vbCrLf & "Uncorrect Autorized", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly)
        Dim frm_DB As New FrmConfigDB("CanNotFindServer")
        frm_DB.ShowDialog(frm)
        If frm_DB.DialogResult <> DialogResult.Yes Then
            ComID = Nothing
            Return False
        Else
            GoTo Jeditcomplete
        End If

    End Function

    '-- สร้าง Log File
    Public Sub CreateLog()
        Dim fso As Object
        fso = CreateObject("Scripting.FileSystemObject")
        'Create a file
        fso.CreateTextFile(My.Application.Info.DirectoryPath & "\" & LogFile)

        fso = Nothing
    End Sub

    '-- เขียน LOG File
    Public Sub WriteLog(ByRef LogData As String)

        Const ForAppending As Integer = 8
        Const TristateUseDefault As Integer = -2

        Dim ts, fso, F As Object

        fso = CreateObject("Scripting.FileSystemObject")
        F = fso.GetFile(My.Application.Info.DirectoryPath & "\" & LogFile)
        ts = F.OpenAsTextStream(ForAppending, TristateUseDefault)
        ts.writeline(VB6.Format(Now, "dd/mm/yyyy  hh:mm:ss") & "   " & LogData)
        ts.Close()
        ts = Nothing
        F = Nothing
        fso = Nothing
    End Sub

    '-- ดึงข้อมูล เเละตรวจสอบ Table ก่อนเริ่มโปรเเกรม
    Public Sub PrepareVatSet()
        Dim cnACCPAC, cnVAT As ADODB.Connection
        Dim rs, rs2 As BaseClass.BaseODBCIO
        Dim sqlstr, Tmp, TmpTax As String

        On Error GoTo ErrHandler

        cnACCPAC = New ADODB.Connection
        cnVAT = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        rs2 = New BaseClass.BaseODBCIO

        If TableHere("FMSVSET", ConnVAT) = False Then
            sqlstr = "CREATE TABLE FMSVSET"
            sqlstr &= " (ORGID      CHAR(6) ,"
            sqlstr &= " CONAME      CHAR(60) ,"
            sqlstr &= " TAXNO       CHAR(20) ,"
            sqlstr &= " DATEFROM    DECIMAL(9,0),"
            sqlstr &= " DATETO      DECIMAL(9,0),"
            sqlstr &= " INGETVAT    DECIMAL(1,0),"
            sqlstr &= " USERGET     varchar(100),"
            sqlstr &= " STATUS      DECIMAL(2,0),"
            sqlstr &= " DISCONNECTUSER DECIMAL(18,0),"
            sqlstr &= " RESTR       DECIMAL(1,0))"
            cnVAT.Open(ConnVAT)
            cnVAT.Execute(sqlstr)
            cnVAT.Close()

            sqlstr = ""
            sqlstr = " SELECT     PGMVER, SELECTOR  FROM CSAPP"
            sqlstr &= "  WHERE    (SELECTOR = 'GL')"

            cnACCPAC.Open(ConnACCPAC)
            rs.Open(sqlstr, cnACCPAC)
            If rs.options.Rs.EOF = True Then Exit Sub
            Tmp = Mid(Trim(rs.options.Rs.Collect(0)), 1, 2)

            If Tmp > "55" Then
                sqlstr = "SELECT ORGID,CONAME,COUNTRY,TAXNBR FROM CSCOM"
            Else
                sqlstr = "SELECT ORGID,CONAME,COUNTRY FROM CSCOM"
            End If
            cnACCPAC.Close()
            cnACCPAC = New ADODB.Connection
            cnACCPAC.Open(ConnACCPAC)
            rs.Open(sqlstr, cnACCPAC)

            Do While rs.options.Rs.EOF = False
                If Tmp > "55" Then
                    TmpTax = Trim(rs.options.Rs.Collect(3))
                    If Trim(rs.options.Rs.Collect(3)) = "" Then
                        TmpTax = Trim(rs.options.Rs.Collect(3))
                    End If
                Else
                    TmpTax = Trim(rs.options.Rs.Collect(2))
                End If
                sqlstr = InsStr & "FMSVSET(ORGID,CONAME,TAXNO,INGETVAT,USERGET,DISCONNECTUSER,RESTR)"
                sqlstr = sqlstr & " Values('" & Trim(rs.options.Rs.Collect(0)) & "','"
                sqlstr = sqlstr & Trim(rs.options.Rs.Collect(1)) & "','" & TmpTax & "',0,'',0,0)"
                cnVAT.Open(ConnVAT)
                cnVAT.Execute(sqlstr)
                cnVAT.Close()
                rs.options.Rs.MoveNext()
                Application.DoEvents()
            Loop
            rs.options.Rs.Close()
            cnACCPAC.Close()
            cnACCPAC = Nothing
        End If

        rs = Nothing
        cnACCPAC = Nothing
        Exit Sub

ErrHandler:
        cnACCPAC.RollbackTrans()
        cnVAT.RollbackTrans()
        ErrMsg = Err.Description()
        WriteLog("<PrepareVatSet>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    '-- ดึงข้อมูล Company name,Tax number,Company ID
    Public Sub PrepareVatUpdateNew()
        Dim cnACCPAC, cnVAT As ADODB.Connection
        Dim rs, rs2 As BaseClass.BaseODBCIO
        Dim sqlstr As String

        On Error GoTo ErrHandler

        cnACCPAC = New ADODB.Connection
        cnACCPAC.Open(ConnACCPAC)
        cnVAT = New ADODB.Connection
        cnVAT.Open(ConnVAT)
        rs = New BaseClass.BaseODBCIO
        rs2 = New BaseClass.BaseODBCIO

        Dim cn As New ADODB.Connection
        cn.Open(ConnVAT)
        rs = New BaseClass.BaseODBCIO
        rs.Open("select isnull(INGETVAT,0)as INGETVAT,isnull(USERGET,'')as USERGET from fmsvset", cn)
        If rs.options.QueryDT.Rows.Count > 0 Then
            sqlstr = "SELECT ORGID,CONAME,TAXNBR FROM CSCOM"
            rs2.Open(sqlstr, cnACCPAC)
            For i As Integer = 0 To rs2.options.QueryDT.Rows.Count - 1
                sqlstr = "Update FMSVSET SET ORGID='" & Trim(rs2.options.QueryDT.Rows(i).Item(0)) & "',CONAME='" & Trim(rs2.options.QueryDT.Rows(i).Item(1)) & "',TAXNO='" & Trim(rs2.options.QueryDT.Rows(i).Item(2)) & "'"
                cn.Execute(sqlstr)
            Next
        Else
            If rs.options.QueryDT.Rows.Count = 0 Then
                sqlstr = "DELETE FROM FMSVSET"
                cn.Execute(sqlstr)
                sqlstr = "SELECT ORGID,CONAME,TAXNBR FROM CSCOM"
                rs2.Open(sqlstr, cnACCPAC)
                For i As Integer = 0 To rs2.options.QueryDT.Rows.Count - 1
                    sqlstr = InsStr & "FMSVSET(ORGID,CONAME,TAXNO,INGETVAT,USERGET,DISCONNECTUSER,RESTR)"
                    sqlstr = sqlstr & " Values('" & Trim(rs2.options.QueryDT.Rows(i).Item(0)) & "','"
                    sqlstr = sqlstr & Trim(rs2.options.QueryDT.Rows(i).Item(1)) & "','" & Trim(rs2.options.QueryDT.Rows(i).Item(2)) & "',0,'',0,0)"
                    cnVAT.Execute(sqlstr)
                Next
            ElseIf rs.options.QueryDT.Rows(0).Item("INGETVAT").ToString.Trim = 0 Then
                sqlstr = "DELETE FROM FMSVSET"
                cn.Execute(sqlstr)
                sqlstr = "SELECT ORGID,CONAME,TAXNBR FROM CSCOM"
                rs2.Open(sqlstr, cnACCPAC)
                For i As Integer = 0 To rs2.options.QueryDT.Rows.Count - 1
                    sqlstr = InsStr & "FMSVSET(ORGID,CONAME,TAXNO,INGETVAT,USERGET,DISCONNECTUSER,RESTR)"
                    sqlstr = sqlstr & " Values('" & Trim(rs2.options.QueryDT.Rows(i).Item(0)) & "','"
                    sqlstr = sqlstr & Trim(rs2.options.QueryDT.Rows(i).Item(1)) & "','" & Trim(rs2.options.QueryDT.Rows(i).Item(2)) & "',0,'',0,0)"
                    cnVAT.Execute(sqlstr)
                Next
            End If
        End If
        rs = Nothing
        rs2 = Nothing
        cn.Close()
        cnACCPAC.Close()
        cn = Nothing
        cnACCPAC = Nothing
        Exit Sub

ErrHandler:
        ErrMsg = Err.Description
        WriteLog("<PrepareVatSet>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    '-- ดึงข้อมูล Location
    Public Sub PrepareLocation()
        Dim cnACCPAC, cnVAT As ADODB.Connection
        Dim rs, rs2 As BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim loAddress As String
        Dim iffs As clsFuntion
        iffs = New clsFuntion

        On Error GoTo ErrHandler

        cnACCPAC = New ADODB.Connection
        cnVAT = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        rs2 = New BaseClass.BaseODBCIO

        If TableHere("FMSVLOC", ConnVAT) = False Then
            sqlstr = "CREATE TABLE FMSVLOC"
            sqlstr = sqlstr & " (LOCID      CHAR(6) ,"
            sqlstr = sqlstr & " LOCNAME     CHAR(60) ,"
            sqlstr = sqlstr & " LOCADD      CHAR(200) ,"
            sqlstr = sqlstr & " LOCPREFIX   CHAR(4) ,"
            sqlstr = sqlstr & " LOCDESC     CHAR(100) )"
            cnVAT.Open(ConnVAT)
            cnVAT.Execute(sqlstr)
            cnVAT.Close()

            cnACCPAC.Open(ConnACCPAC)
            sqlstr = "SELECT ORGID,CONAME,ADDR01,ADDR02,ADDR03,ADDR04 FROM CSCOM"
            rs.Open(sqlstr, cnACCPAC)
            Do While rs.options.Rs.EOF = False
                loAddress = IIf(iffs.ISNULL(rs.options.Rs, 2), "", Trim(rs.options.Rs.Collect(2)))
                loAddress = loAddress & " " & IIf(iffs.ISNULL(rs.options.Rs, 3), "", Trim(rs.options.Rs.Collect(3)))
                loAddress = loAddress & " " & IIf(iffs.ISNULL(rs.options.Rs, 4), "", Trim(rs.options.Rs.Collect(4)))
                loAddress = loAddress & " " & IIf(iffs.ISNULL(rs.options.Rs, 5), "", Trim(rs.options.Rs.Collect(5)))

                sqlstr = InsStr & "FMSVLOC(LOCID,LOCNAME,LOCADD)"
                sqlstr = sqlstr & " Values('" & Trim(rs.options.Rs.Collect(0)) & "','"
                sqlstr = sqlstr & Trim(rs.options.Rs.Collect(1)) & "','"
                sqlstr = sqlstr & Trim(loAddress) & "')"
                cnVAT.Open(ConnVAT)
                cnVAT.Execute(sqlstr)
                cnVAT.Close()
                rs.options.Rs.MoveNext()
                Application.DoEvents()
            Loop
            cnACCPAC.Close()
            cnACCPAC = Nothing
        End If
        rs = Nothing
        cnACCPAC = Nothing
        Exit Sub

ErrHandler:
        cnVAT.RollbackTrans()
        ErrMsg = Err.Description
        WriteLog("<PrepareLocation>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    '-- สร้าง Table เก็บเลข RUNNING
    Public Sub PrepareRUNNING()
        Dim cn As ADODB.Connection
        Dim sqlstr As String
        On Error GoTo ErrHandler
        cn = New ADODB.Connection
        cn.Open(ConnVAT)

        If TableHere("FMSRUN", ConnVAT) = False Then
            sqlstr = "CREATE TABLE FMSRUN (" & Environment.NewLine
            sqlstr &= "	BATCH     decimal(9, 0) NOT NULL," & Environment.NewLine
            sqlstr &= "	ENTRY     decimal(7, 0) NOT NULL," & Environment.NewLine
            sqlstr &= "	TRANSNBR  varchar(15) NOT NULL," & Environment.NewLine
            sqlstr &= "	RUNNING   varchar(20) NOT NULL," & Environment.NewLine
            sqlstr &= "	SOURCE    char(2) NULL " & Environment.NewLine
            sqlstr &= ") "
            'ON PRIMARY" & Environment.NewLine
            cn.Execute(sqlstr)
        End If
        cn.Close()
        cn = Nothing
        Exit Sub

ErrHandler:
        ErrMsg = Err.Description()
        WriteLog("<PrepareRUNNING>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    '-- สร้าง Table เก็บ Account ของ Location
    Public Sub PrepareAccount()
        Dim cn As ADODB.Connection
        Dim sqlstr As String

        On Error GoTo ErrHandler

        cn = New ADODB.Connection

        cn.Open(ConnVAT)

        If TableHere("FMSVLACC", ConnVAT) = False Then
            sqlstr = "CREATE TABLE FMSVLACC"
            sqlstr = sqlstr & " (ACCTVAT    CHAR(45),"
            sqlstr = sqlstr & " LOCID       CHAR(6) ,"
            sqlstr = sqlstr & " VTYPE       CHAR(1))"
            cn.Execute(sqlstr)
        End If
        cn.Close()
        cn = Nothing
        Exit Sub

ErrHandler:
        ErrMsg = Err.Description
        WriteLog("<PrepareAccount>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    '-- ใช้เก็บ Account Tax
    Public Sub PrepareTax()
        Dim cn As ADODB.Connection
        Dim sqlstr As String

        On Error GoTo ErrHandler

        cn = New ADODB.Connection
        'cn.Open(ConnVAT)

        If TableHere("FMSTAX", ConnVAT) = True Then
            'sqlstr = "DROP TABLE FMSTAX"
            sqlstr = "DELETE FROM FMSTAX"
            cn.Open(ConnVAT)
            cn.Execute(sqlstr)
            cn.Close()
        Else
            cn = New ADODB.Connection
            cn.Open(ConnVAT)
            sqlstr = "CREATE TABLE FMSTAX ("
            sqlstr = sqlstr & " AUTHORITY   CHAR(12)," 'Change from  char(6) to char(12) on 01/10/2004
            sqlstr = sqlstr & " TTYPE       DECIMAL(15,0),"
            sqlstr = sqlstr & " TYPE        CHAR(8),"
            sqlstr = sqlstr & " ITEMRATE1   DECIMAL(15,5),"
            sqlstr = sqlstr & " LIABILITY   CHAR(45),"
            sqlstr = sqlstr & " ACCTRECOV   CHAR(45),"
            sqlstr = sqlstr & " ACCTVAT     CHAR(45),"
            sqlstr = sqlstr & " BUYERCLASS  DECIMAL(15,0),"
            sqlstr = sqlstr & " RATERECOV  DECIMAL(15,5))"

            cn.Execute(sqlstr)
            cn.Close()
        End If

        cn = Nothing
        Exit Sub

ErrHandler:
        ErrMsg = Err.Description
        WriteLog("<PrepareTax>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    '-- เก็บ Account GL
    Public Sub PrepareTaxGL()
        Dim cn As ADODB.Connection
        Dim sqlstr As String

        On Error GoTo ErrHandler

        cn = New ADODB.Connection

        If TableHere("FMSTAXGL", ConnVAT) = True Then
            'sqlstr = "DROP TABLE FMSTAXGL"
            sqlstr = "DELETE FROM  FMSTAXGL"
            cn.Open(ConnVAT)
            cn.Execute(sqlstr)
            cn.Close()
        Else
            cn = New ADODB.Connection
            cn.Open(ConnVAT)
            sqlstr = "CREATE TABLE FMSTAXGL"
            sqlstr = sqlstr & " (TTYPE      SMALLINT,"
            sqlstr = sqlstr & " ACCTVAT     CHAR(45),"
            sqlstr = sqlstr & " AUTHORITY   CHAR(12))"
            cn.Execute(sqlstr)
            cn.Close()
        End If

        cn = Nothing
        Exit Sub

ErrHandler:
        ErrMsg = Err.Description
        WriteLog("<PrepareTaxGL>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    Public Sub SetInitial()
        If ComDatabase = "MSSQL" Then
            InsStr = InsMSSQL
        Else
            InsStr = InsPVSW
        End If
        VATID = 0
    End Sub

    '*************** Create Table Vat
    Public Sub PrepareVatTable()
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim iffs As clsFuntion
        iffs = New clsFuntion

        On Error GoTo ErrHandler
        cn = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        cn.Open(ConnVAT)

        If TableHere("FMSVAT", ConnVAT) = False Then
            sqlstr = "CREATE TABLE FMSVAT"
            If ComDatabase = "MSSQL" Then
                sqlstr = sqlstr & " (VATID      [decimal](19, 0)  IDENTITY(1,1)  ,"
            ElseIf ComDatabase = "ORCL" Then
                sqlstr = sqlstr & " (VATID      Number(19,0)  ,"
            Else
                sqlstr = sqlstr & " (VATID      IDENTITY  ,"
            End If
            sqlstr = sqlstr & " INVDATE     CHAR(8),"
            sqlstr = sqlstr & " TXDATE      CHAR(255),"
            sqlstr = sqlstr & " IDINVC      CHAR(255) ,"
            sqlstr = sqlstr & " DOCNO       CHAR(255) ,"
            sqlstr = sqlstr & " NEWDOCNO    CHAR(60) ,"
            sqlstr = sqlstr & " INVNAME     CHAR(200) ,"
            sqlstr = sqlstr & " INVAMT      DECIMAL(19,3),"
            sqlstr = sqlstr & " INVTAX      DECIMAL(19,3),"
            sqlstr = sqlstr & " LOCID       CHAR(6)NOT NULL ,"
            sqlstr = sqlstr & " VTYPE       DECIMAL(19,0),"
            sqlstr = sqlstr & " RATE        DECIMAL(19,3),"
            sqlstr = sqlstr & " TTYPE       DECIMAL(19,0),"
            sqlstr = sqlstr & " ACCTVAT     CHAR(45) ,"
            sqlstr = sqlstr & " SOURCE      CHAR(2)NOT NULL ,"
            sqlstr = sqlstr & " BATCH       decimal(6, 0),"
            sqlstr = sqlstr & " ENTRY       decimal(5, 0) ,"
            sqlstr = sqlstr & " MARK        CHAR(60) ," ''''
            sqlstr = sqlstr & " VATCOMMENT  CHAR(1000),"
            sqlstr = sqlstr & " CBREF       CHAR(250),"
            sqlstr = sqlstr & " TRANSNBR    VARCHAR(15),"
            sqlstr = sqlstr & " RUNNO       VARCHAR(60)," 'แก้จาก VARCHAR(20) เป็น VARCHAR(60) Edit 07/02/2014 By Pat
            sqlstr = sqlstr & " TypeOfPU    VARCHAR(30),"
            sqlstr = sqlstr & " TranNo      Varchar(100),"
            sqlstr = sqlstr & " CIF         DECIMAL(19,3),"
            sqlstr = sqlstr & " TAXCIF      DECIMAL(19,3),"
            sqlstr = sqlstr & " Verify      DECIMAL(1,0),"
            sqlstr = sqlstr & " TAXID       VARCHAR(100),"
            sqlstr = sqlstr & " BRANCH      VARCHAR(500),"
            sqlstr = sqlstr & " Code        VARCHAR(12),"
            sqlstr = sqlstr & " CURRENCY    VARCHAR(3),"
            sqlstr = sqlstr & " EXCHANGRATE DECIMAL(15,7),"
            sqlstr = sqlstr & " AMTBASETC   DECIMAL(19,3),"
            sqlstr = sqlstr & " TOTTAX      DECIMAL(19,3),"
            sqlstr = sqlstr & " CODETAX     CHAR(12))"


            cn.Execute(sqlstr)
        ElseIf ComDatabase = "ORCL" Then
            sqlstr = "select max(VATID) FROM  FMSVAT"
            rs = cn.Execute(sqlstr)
            gVatId = IIf(iffs.ISNULL(rs.options.Rs, 0), 0, rs.options.Rs.Fields(0).Value + 1)
        End If

        cn.Close()
        rs = Nothing
        cn = Nothing
        Exit Sub

ErrHandler:
        cn.RollbackTrans()
        ErrMsg = Err.Description
        WriteLog("<PrepareVatTable>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub
    '*************** Create Table Vat Insert ***********************************
    Public Sub PrepareVatInsertTable()
        Dim cn As ADODB.Connection
        Dim sqlstr As String
        On Error GoTo ErrHandler
        cn = New ADODB.Connection
        cn.Open(ConnVAT)
        If TableHere("FMSVATINSERT", ConnVAT) = False Then
            sqlstr = "CREATE TABLE FMSVATINSERT"
            If ComDatabase = "MSSQL" Then
                sqlstr = sqlstr & " (VATID      [decimal](19, 0) ,"
            ElseIf ComDatabase = "ORCL" Then
                sqlstr = sqlstr & " (VATID      Number(19,0)  ,"
            Else
                sqlstr = sqlstr & " (VATID      IDENTITY  ,"
            End If
            sqlstr = sqlstr & " INVDATE     CHAR(8),"
            sqlstr = sqlstr & " TXDATE      CHAR(255),"
            sqlstr = sqlstr & " IDINVC      CHAR(255) ,"
            sqlstr = sqlstr & " DOCNO       CHAR(255) ,"
            sqlstr = sqlstr & " NEWDOCNO    CHAR(60) ,"
            sqlstr = sqlstr & " INVNAME     CHAR(250) ,"
            sqlstr = sqlstr & " INVAMT      DECIMAL(19,3),"
            sqlstr = sqlstr & " INVTAX      DECIMAL(19,3),"
            sqlstr = sqlstr & " LOCID       CHAR(6)NOT NULL ,"
            sqlstr = sqlstr & " VTYPE       DECIMAL(19,0),"
            sqlstr = sqlstr & " RATE        DECIMAL(19,3),"
            sqlstr = sqlstr & " TTYPE       DECIMAL(19,0),"
            sqlstr = sqlstr & " ACCTVAT     CHAR(45) ,"
            sqlstr = sqlstr & " SOURCE      CHAR(2)NOT NULL ,"
            sqlstr = sqlstr & " BATCH       decimal(6, 0),"
            sqlstr = sqlstr & " ENTRY       decimal(5, 0) ,"
            sqlstr = sqlstr & " MARK        CHAR(60) ," ''''
            sqlstr = sqlstr & " VATCOMMENT  CHAR(1000) ,"
            sqlstr = sqlstr & " CBREF       CHAR(1000),"
            sqlstr = sqlstr & " STATUS      CHAR(5),"
            sqlstr = sqlstr & " TRANSNBR    CHAR(15),"
            sqlstr = sqlstr & " RUNNO       VARCHAR(60)," 'แก้จาก VARCHAR(20) เป็น VARCHAR(60) Edit 07/02/2014 By Pat
            sqlstr = sqlstr & " TypeOfPU    VARCHAR(30),"
            sqlstr = sqlstr & " TranNo      Varchar(100),"
            sqlstr = sqlstr & " CIF         DECIMAL(19,3),"
            sqlstr = sqlstr & " TAXCIF      DECIMAL(19,3),"
            sqlstr = sqlstr & " Verify      DECIMAL(1,0),"
            sqlstr = sqlstr & " TAXID       VARCHAR(100),"
            sqlstr = sqlstr & " BRANCH      VARCHAR(500),"
            sqlstr = sqlstr & " CURRENCY    VARCHAR(3),"
            sqlstr = sqlstr & " EXCHANGRATE DECIMAL(15,7),"
            sqlstr = sqlstr & " AMTBASETC   DECIMAL(19,3))"

            cn.Execute(sqlstr)
        End If
        cn.Close()

        cn = Nothing
        Exit Sub

ErrHandler:
        cn.RollbackTrans()
        ErrMsg = Err.Description
        WriteLog("<PrepareVatInsertTable>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    '-- Crate Table TEMP ก่อน Get Process AP,Process PO,Process AP-PO
    Public Sub PrepareTempTableAP()

        Dim cn As ADODB.Connection
        Dim sqlstr As String
        Dim sqlIndex As String

        On Error GoTo ErrHandler

        cn = New ADODB.Connection

        If TableHere("APOBLTEMP", ConnVAT) = True Then
            sqlstr = "DROP TABLE APOBLTEMP"
            cn.Open(ConnVAT)
            cn.Execute(sqlstr)
            cn.Close()
        End If
        If TableHere("APPJHTEMP", ConnVAT) = True Then
            sqlstr = "DROP TABLE APPJHTEMP"
            cn.Open(ConnVAT)
            cn.Execute(sqlstr)
            cn.Close()
        End If


        sqlstr = "CREATE TABLE APOBLTEMP"
        sqlstr = sqlstr & " (IDVEND     char(12) NOT NULL,"
        sqlstr = sqlstr & " IDINVC      char(22) NOT NULL,"
        sqlstr = sqlstr & " AUDTDATE    decimal(9, 0) NOT NULL,"
        sqlstr = sqlstr & " IDPONBR     char(22) NOT NULL,"
        sqlstr = sqlstr & " TXTTRXTYPE  smallint NOT NULL,"
        sqlstr = sqlstr & " CNTBTCH     decimal(9, 0) NOT NULL,"
        sqlstr = sqlstr & " CNTITEM     decimal(7, 0) NOT NULL,"
        sqlstr = sqlstr & " DATEINVC    decimal(9, 0) NOT NULL,"
        sqlstr = sqlstr & " AMTINVCHC   decimal(19, 3) NOT NULL,"
        sqlstr = sqlstr & " AMTTAXHC    decimal(19, 3) NOT NULL,"
        sqlstr = sqlstr & " CODETAX1    char(12) NOT NULL,"
        sqlstr = sqlstr & " AMTBASE1HC  decimal(19, 3) NOT NULL,"
        sqlstr = sqlstr & " AMTTAX1HC   decimal(19, 3) NOT NULL,"
        sqlstr = sqlstr & " DATEBUS     decimal(9, 0) NOT NULL,"
        sqlstr = sqlstr & " CODETAXGRP  char(12) NOT NULL,"


        If ComDatabase = "MSSQL" Then
            sqlstr = sqlstr & " [VALUES]      int NOT NULL,"
            sqlstr = sqlstr & " CONSTRAINT  [APOBL_KEY_0] PRIMARY KEY NONCLUSTERED"
            sqlstr = sqlstr & " ([IDVEND] ASC,[IDINVC] ASC))"

            cn.Open(ConnVAT)
            cn.Execute(sqlstr)
            cn.Close()

        Else 'PVSW
            sqlstr = sqlstr & " " & Chr(34) & "VALUES" & Chr(34) & " int NOT NULL);"
            sqlIndex = " CREATE UNIQUE NOT MODIFIABLE INDEX key0 ON APOBLTEMP(IDVEND, IDINVC);"

            cn.Open(ConnVAT)
            cn.Execute(sqlstr)
            cn.Close()

            cn.Open(ConnVAT)
            cn.Execute(sqlIndex)
            cn.Close()
        End If



        sqlstr = "CREATE TABLE APPJHTEMP"
        sqlstr = sqlstr & " (TYPEBTCH   char(2) NOT NULL,"
        sqlstr = sqlstr & " POSTSEQNCE  decimal(9, 0) NOT NULL,"
        sqlstr = sqlstr & " CNTBTCH     decimal(9, 0) NOT NULL,"
        sqlstr = sqlstr & " CNTITEM     decimal(7, 0) NOT NULL,"
        sqlstr = sqlstr & " IDVEND      char(12) NOT NULL,"
        sqlstr = sqlstr & " IDINVC      char(22) NOT NULL,"

        If ComDatabase = "MSSQL" Then
            sqlstr = sqlstr & " [DATEBUS]   [decimal](9, 0) NOT NULL,"
            sqlstr = sqlstr & " CONSTRAINT  [APPJH_KEY_0] PRIMARY KEY CLUSTERED"
            sqlstr = sqlstr & " ([TYPEBTCH] ASC,[POSTSEQNCE] ASC,[CNTBTCH] ASC,[CNTITEM] ASC))"
            cn.Open(ConnVAT)
            cn.Execute(sqlstr)
            cn.Close()
        Else 'PVSW
            sqlstr = sqlstr & " DATEBUS   decimal(9, 0) NOT NULL);"
            sqlIndex = " CREATE UNIQUE NOT MODIFIABLE INDEX key0 ON APPJHTEMP(TYPEBTCH,POSTSEQNCE, CNTBTCH,CNTITEM);"

            cn.Open(ConnVAT)
            cn.Execute(sqlstr)
            cn.Close()

            cn.Open(ConnVAT)
            cn.Execute(sqlIndex)
            cn.Close()

        End If

        cn = Nothing
        Exit Sub


ErrHandler:
        cn.RollbackTrans()
        ErrMsg = Err.Description
        WriteLog("<PrepareTempTableAP>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    '-- Crate Table TEMP ก่อน Get Process AR,Process OE,Process AR-OE
    Public Sub PrepareTempTableAR()

        Dim cn As ADODB.Connection
        Dim sqlstr As String
        Dim sqlIndex As String

        On Error GoTo ErrHandler

        cn = New ADODB.Connection

        If TableHere("AROBLTEMP", ConnVAT) = True Then
            sqlstr = "DROP TABLE AROBLTEMP"
            cn.Open(ConnVAT)
            cn.Execute(sqlstr)
            cn.Close()
        End If
        If TableHere("ARPJHTEMP", ConnVAT) = True Then
            sqlstr = "DROP TABLE ARPJHTEMP"
            cn.Open(ConnVAT)
            cn.Execute(sqlstr)
            cn.Close()
        End If


        sqlstr = "CREATE TABLE AROBLTEMP"
        sqlstr = sqlstr & " (IDCUST     char(12) NOT NULL,"
        sqlstr = sqlstr & " IDINVC      char(22) NOT NULL,"
        sqlstr = sqlstr & " IDRMIT      char(24) NOT NULL,"
        sqlstr = sqlstr & " TRXTYPEID   smallint NOT NULL,"
        sqlstr = sqlstr & " TRXTYPETXT  smallint NOT NULL,"
        sqlstr = sqlstr & " CNTBTCH     decimal(9, 0) NOT NULL,"
        sqlstr = sqlstr & " CNTITEM     decimal(7, 0) NOT NULL,"
        sqlstr = sqlstr & " DATEINVC    decimal(9, 0) NOT NULL,"
        sqlstr = sqlstr & " AMTINVCHC   decimal(19, 3) NOT NULL,"
        sqlstr = sqlstr & " CODECURN    char(3) NOT NULL,"
        sqlstr = sqlstr & " EXCHRATEHC  decimal(15, 7) NOT NULL,"
        sqlstr = sqlstr & " AMTTAXHC    decimal(19, 3) NOT NULL,"
        sqlstr = sqlstr & " CODETAX1    char(12) NOT NULL,"
        sqlstr = sqlstr & " DATEBUS     decimal(9, 0) NOT NULL,"
        sqlstr = sqlstr & " CODETAXGRP  char(12) NOT NULL,"


        If ComDatabase = "MSSQL" Then
            sqlstr = sqlstr & " [VALUES]      int NOT NULL,"
            sqlstr = sqlstr & " CONSTRAINT  [AROBL_KEY_0] PRIMARY KEY NONCLUSTERED"
            sqlstr = sqlstr & " ([IDCUST] ASC,[IDINVC] ASC))"

            cn.Open(ConnVAT)
            cn.Execute(sqlstr)
            cn.Close()

        Else 'PVSW
            sqlstr = sqlstr & " " & Chr(34) & "VALUES" & Chr(34) & " int NOT NULL);"
            sqlIndex = " CREATE UNIQUE NOT MODIFIABLE INDEX key0 ON AROBLTEMP(IDCUST, IDINVC);"

            cn.Open(ConnVAT)
            cn.Execute(sqlstr)
            cn.Close()

            cn.Open(ConnVAT)
            cn.Execute(sqlIndex)
            cn.Close()
        End If



        sqlstr = "CREATE TABLE ARPJHTEMP"
        sqlstr = sqlstr & " (TYPEBTCH   char(2) NOT NULL,"
        sqlstr = sqlstr & " POSTSEQNCE  decimal(9, 0) NOT NULL,"
        sqlstr = sqlstr & " CNTBTCH     decimal(9, 0) NOT NULL,"
        sqlstr = sqlstr & " CNTITEM     decimal(7, 0) NOT NULL,"
        sqlstr = sqlstr & " IDCUST      char(12) NOT NULL,"
        sqlstr = sqlstr & " IDINVC      char(22) NOT NULL,"

        If ComDatabase = "MSSQL" Then
            sqlstr = sqlstr & " [DATEBUS]   [decimal](9, 0) NOT NULL,"
            sqlstr = sqlstr & " CONSTRAINT  [ARPJH_KEY_0] PRIMARY KEY CLUSTERED"
            sqlstr = sqlstr & " ([TYPEBTCH] ASC,[POSTSEQNCE] ASC,[CNTBTCH] ASC,[CNTITEM] ASC))"
            cn.Open(ConnVAT)
            cn.Execute(sqlstr)
            cn.Close()
        Else 'PVSW
            sqlstr = sqlstr & " DATEBUS   decimal(9, 0) NOT NULL);"
            sqlIndex = " CREATE UNIQUE NOT MODIFIABLE INDEX key0 ON ARPJHTEMP(TYPEBTCH,POSTSEQNCE, CNTBTCH,CNTITEM);"

            cn.Open(ConnVAT)
            cn.Execute(sqlstr)
            cn.Close()

            cn.Open(ConnVAT)
            cn.Execute(sqlIndex)
            cn.Close()

        End If

        cn = Nothing
        Exit Sub


ErrHandler:
        cn.RollbackTrans()
        ErrMsg = Err.Description
        WriteLog("<PrepareTempTableAR>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub
    '-- Crate Table TMP ก่อนเอาเข้า FMSVAT
    Public Sub PrepareUserPassTable()
     
        Dim cn As ADODB.Connection
        Dim sqlstr As String

        On Error GoTo ErrHandler
        cn = New ADODB.Connection
        cn.Open(ConnVAT)

        If TableHere("USERPASS", ConnVAT) = False Then

            sqlstr = "CREATE TABLE USERPASS"
            sqlstr = sqlstr & " ( USERID    Int,"
            sqlstr = sqlstr & " USERNAME    CHAR(100),"
            sqlstr = sqlstr & " PASS        CHAR(100) ,"
            sqlstr = sqlstr & " NAME        CHAR(200)) "

            cn.Execute(sqlstr)

            sqlstr = " insert into USERPASS  values(0,'ZJSIpmqbQDqJtjdsYIxiPw==','ZJSIpmqbQDqJtjdsYIxiPw==','Administrator') " + Environment.NewLine

            cn.Execute(sqlstr)
 
        End If

        cn.Close()
        cn = Nothing
        Exit Sub

ErrHandler:
        cn.RollbackTrans()
        ErrMsg = Err.Description
        WriteLog("<PrepareUserPassTable>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub
    '-- Crate Table TMP ก่อนเอาเข้า FMSVAT
    Public Sub PrepareTempTable()
        Dim cn As ADODB.Connection
        Dim sqlstr As String
        On Error GoTo ErrHandler
        cn = New ADODB.Connection

        If TableHere("FMSVATTEMP", ConnVAT) = True Then
            sqlstr = "DROP TABLE FMSVATTEMP"
            cn.Open(ConnVAT)
            cn.Execute(sqlstr)
            cn.Close()
        End If

        sqlstr = "CREATE TABLE FMSVATTEMP" '--------------
        If ComDatabase = "MSSQL" Then
            sqlstr = sqlstr & " (VATID      [decimal](19, 0)  IDENTITY(1,1)  ,"

        ElseIf ComDatabase = "ORCL" Then
            sqlstr = sqlstr & " (VATID      Number(19,0)  ,"
        Else
            sqlstr = sqlstr & " (VATID      IDENTITY  ,"
        End If
        sqlstr = sqlstr & " INVDATE     CHAR(8),"
        sqlstr = sqlstr & " TXDATE      CHAR(255),"
        sqlstr = sqlstr & " IDINVC      CHAR(255) ,"
        sqlstr = sqlstr & " DOCNO       CHAR(255) ,"
        sqlstr = sqlstr & " NEWDOCNO    CHAR(60) ,"
        sqlstr = sqlstr & " INVNAME     CHAR(250) ,"
        sqlstr = sqlstr & " INVAMT      DECIMAL(19,3),"
        sqlstr = sqlstr & " INVTAX      DECIMAL(19,3),"
        sqlstr = sqlstr & " LOCID       CHAR(6) ,"
        sqlstr = sqlstr & " VTYPE       DECIMAL(19,0),"
        sqlstr = sqlstr & " RATE        DECIMAL(19,3),"
        sqlstr = sqlstr & " TTYPE       DECIMAL(19,0),"
        sqlstr = sqlstr & " ACCTVAT     CHAR(45) ,"
        sqlstr = sqlstr & " SOURCE      CHAR(2) ,"
        sqlstr = sqlstr & " BATCH       decimal(6, 0),"
        sqlstr = sqlstr & " ENTRY       decimal(5, 0) ,"
        sqlstr = sqlstr & " MARK        CHAR(60) ," ''''
        sqlstr = sqlstr & " VATCOMMENT  CHAR(1000) ,"
        sqlstr = sqlstr & " CODETAX     CHAR(12) ,"
        sqlstr = sqlstr & " IDDIST      CHAR(12),"
        sqlstr = sqlstr & " CBREF       CHAR(1000),"
        sqlstr = sqlstr & " TRANSNBR    VARCHAR(15),"
        sqlstr = sqlstr & " RUNNO       VARCHAR(60)," 'แก้จาก VARCHAR(20) เป็น VARCHAR(60) Edit 07/02/2014 By Pat
        sqlstr = sqlstr & " TypeOfPU    VARCHAR(30),"
        sqlstr = sqlstr & " TranNo      Varchar(100),"
        sqlstr = sqlstr & " CIF         DECIMAL(19,3),"
        sqlstr = sqlstr & " TAXCIF      DECIMAL(19,3),"
        sqlstr = sqlstr & " TAXID       VARCHAR(100),"
        sqlstr = sqlstr & " BRANCH      VARCHAR(500),"
        sqlstr = sqlstr & " Code        VARCHAR(12),"
        sqlstr = sqlstr & " TOTTAX      DECIMAL(19,3))"
        cn.Open(ConnVAT)
        cn.Execute(sqlstr)
        cn.Close()
        cn = Nothing
        Exit Sub

ErrHandler:
        cn.RollbackTrans()
        ErrMsg = Err.Description
        WriteLog("<PrepareVatTempTable>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    '-- Delete ข้อมูลใน Table FMSVAT ตอนกด GETVAT ใหม่
    Public Sub DropVatTable()
        Dim cn As ADODB.Connection
        Dim sqlstr As String
        On Error GoTo ErrHandler
        cn = New ADODB.Connection
        cn.Open(ConnVAT)
        If TableHere("FMSVAT", ConnVAT) = True Then
            cn.Close()
            If ComDatabase = "PVSW" Then
                sqlstr = "DELETE FROM FMSVAT"
            ElseIf ComDatabase = "MSSQL" Then
                sqlstr = "TRUNCATE TABLE FMSVAT"
            End If
            cn.Open(ConnVAT)
            cn.Execute(sqlstr)
        End If
        cn.Close()

        cn = Nothing
        Exit Sub

ErrHandler:
        cn.RollbackTrans()
        ErrMsg = Err.Description
        WriteLog("<DropVatTable>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    '--ลบเเละดึงข้อมูล Tax ใหม่
    Public Sub ProcessTax()
        Dim cnVAT, cnACCPAC As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String

        On Error GoTo ErrHandler

        cnVAT = New ADODB.Connection
        cnACCPAC = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        cnVAT.ConnectionTimeout = 60
        cnVAT.CommandTimeout = 3600
        cnVAT.Open(ConnVAT)

        cnACCPAC.ConnectionTimeout = 60
        cnACCPAC.Open(ConnACCPAC)
        cnACCPAC.CommandTimeout = 3600


        sqlstr = "DELETE FROM FMSTAX"
        cnVAT.Execute(sqlstr)
        sqlstr = "SELECT TXAUTH.AUTHORITY,TXRATE.TTYPE, TXRATE.ITEMRATE1," '0-2
        sqlstr = sqlstr & " TXAUTH.LIABILITY, TXAUTH.ACCTRECOV, TXRATE.BUYERCLASS,TXAUTH.RATERECOV" '3-6
        If ComDatabase = "ORCL" Then
            sqlstr = sqlstr & " FROM TXRATE INNER JOIN TXAUTH ON TXRATE.AUTHORITY = TXAUTH.AUTHORITY"
        Else
            sqlstr = sqlstr & " FROM TXRATE INNER JOIN TXAUTH ON TXRATE.AUTHORITY = TXAUTH.AUTHORITY"
        End If

        rs.Open(sqlstr, cnACCPAC)
        Do While rs.options.Rs.EOF = False
            sqlstr = InsStr & "FMSTAX(AUTHORITY,TTYPE,ITEMRATE1,LIABILITY,ACCTRECOV,ACCTVAT,TYPE,BUYERCLASS,RATERECOV) "
            sqlstr = sqlstr & " Values('" & Trim(rs.options.Rs.Collect(0)) & "'," & rs.options.Rs.Collect(1) & ",'" & rs.options.Rs.Collect(2) & "','"
            sqlstr = sqlstr & Trim(rs.options.Rs.Collect(3)) & "','" & Trim(rs.options.Rs.Collect(4)) & "','"
            sqlstr = sqlstr & IIf(rs.options.Rs.Collect(1) = 1, Trim(rs.options.Rs.Collect(4)), IIf(rs.options.Rs.Collect(1) = 2, Trim(rs.options.Rs.Collect(3)), "")) & "','"
            sqlstr = sqlstr & IIf(rs.options.Rs.Collect(1) = 1, "PURCHASE", IIf(rs.options.Rs.Collect(1) = 2, "SALE", "NONE")) & "'," & Val(rs.options.Rs.Collect(5)) & "," & rs.options.Rs.Collect(6) & ")"
            cnVAT.Execute(sqlstr)
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()

        sqlstr = "SELECT AUTHORITY,3,1,ACCTEXP FROM TXAUTH "
        sqlstr = sqlstr & " Where EXPSEPARTE = 1"
        rs.Open(sqlstr, cnACCPAC)
        Do While rs.options.Rs.EOF = False
            sqlstr = InsStr & "FMSTAX(AUTHORITY,TTYPE,ITEMRATE1,LIABILITY,ACCTRECOV,ACCTVAT,TYPE)"
            sqlstr = sqlstr & " Values('" & Trim(rs.options.Rs.Collect(0)) & "','" & rs.options.Rs.Collect(1) & "','" & rs.options.Rs.Collect(2) & "','"
            sqlstr = sqlstr & Trim(rs.options.Rs.Collect(3)) & "','" & Trim(rs.options.Rs.Collect(3)) & "','"
            sqlstr = sqlstr & Trim(rs.options.Rs.Collect(3)) & "','EXPENSE')"
            cnVAT.Execute(sqlstr)
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()

        cnVAT.Close()
        cnACCPAC.Close()
        rs = Nothing
        cnVAT = Nothing
        cnACCPAC = Nothing
        Exit Sub

ErrHandler:
        ErrMsg = Err.Description
        WriteLog("<ProcessTax>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessTax> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    '-- ลบเเละดึงข้อมูล GL Tax ใหม่
    Public Sub ProcessTaxGL()
        Dim cnVAT, cnACCPAC As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String

        On Error GoTo ErrHandler

        cnVAT = New ADODB.Connection
        cnACCPAC = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        cnVAT.ConnectionTimeout = 60
        cnVAT.Open(ConnVAT)
        cnVAT.CommandTimeout = 3600

        cnACCPAC.ConnectionTimeout = 60
        cnACCPAC.Open(ConnVAT)
        cnACCPAC.CommandTimeout = 3600
        sqlstr = "DELETE FROM FMSTAXGL"
        cnVAT.Execute(sqlstr)
        sqlstr = " select VTYPE,ACCTVAT,LOCID from FMSVLACC GROUP BY ACCTVAT,VTYPE,LOCID"
        rs.Open(sqlstr, cnACCPAC)
        Do While rs.options.Rs.EOF = False
            sqlstr = InsStr & " FMSTAXGL(TTYPE,ACCTVAT,AUTHORITY)"
            sqlstr = sqlstr & " Values(" & rs.options.Rs.Collect(0) & ",'" & rs.options.Rs.Collect(1) & "','" & rs.options.Rs.Collect(2) & "')"
            cnVAT.Execute(sqlstr)
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop

        rs.options.Rs.Close()
        cnVAT.Close()
        cnACCPAC.Close()
        rs = Nothing
        cnVAT = Nothing
        cnACCPAC = Nothing
        Exit Sub

ErrHandler:
        ErrMsg = Err.Description
        WriteLog("<ProcessTaxGL>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessTaxGL> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    '-- ดึงข้อมูล Module GL
    Public Sub ProcessGL()
        Dim cnACCPAC, cnVAT As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim strsplit(), strsplit2() As String
        Dim losplit, losplit2 As Integer
        Dim splittxt, splittxt2 As String
        Dim loCls As clsVat
        Dim loTAXID As String = ""
        Dim loBRANCH As String = ""
        Dim loMIC As String = ""
        Dim tmpDate As String
        Dim iffs As clsFuntion
        iffs = New clsFuntion
        ErrMsg = ""

        On Error GoTo ErrHandler
        cnACCPAC = New ADODB.Connection
        cnVAT = New ADODB.Connection

        cnACCPAC.ConnectionTimeout = 60
        cnACCPAC.Open(ConnACCPAC)
        cnACCPAC.CommandTimeout = 3600

        cnVAT.ConnectionTimeout = 60
        cnVAT.Open(ConnVAT)
        cnVAT.CommandTimeout = 3600
        'gather data
        rs = New BaseClass.BaseODBCIO

        'Try
        If Use_POST Then
            If ComDatabase = "ORCL" Then
                sqlstr = "SELECT    GLPOST.BATCHNBR , GLPOST.ENTRYNBR, GLPOST.JRNLDATE,GLPJC." & Chr(34) & "COMMENT" & Chr(34) & "," '0-3
                sqlstr &= "         GLPOST.JNLDTLREF,GLPOST.JNLDTLDESC,GLPOST.TRANSQTY," '4-6
                sqlstr &= "         GLPOST.TRANSAMT,GLAMF.ACCTFMTTD,GLPOST.TRANSNBR," '7-9
                sqlstr &= "         GLPOST.POSTINGSEQ,GLPOST.CNTDETAIL,GLPOST.FISCALYR,GLPOST.FISCALPERD,GLJEDO.VALUE" '10-13
                sqlstr &= " FROM    GLPOST , GLPJC,GLAMF,GLJEDO "
                sqlstr &= " WHERE   GLPOST.TRANSNBR = GLPJC.TRANSNBR"
                sqlstr &= "         AND GLPOST.ENTRYNBR = GLPJC.ENTRYNBR"
                sqlstr &= "         AND GLPOST.BATCHNBR = GLPJC.BATCHNBR "
                sqlstr &= "         AND GLPOST.ACCTID = GLAMF.ACCTID"
                sqlstr &= "         AND GLJEDO.BATCHNBR=GLPOST.BATCHNBR"
                sqlstr &= "         AND GLJEDO.JOURNALID=GLPOST.ENTRYNBR"
                sqlstr &= "         AND GLJEDO.OPTFIELD='VOUCHER'"
                sqlstr &= "         AND GLPOST.SRCELEDGER = 'GL' AND GLPOST.SRCETYPE <>'NV' AND"
                sqlstr &= "         GLPOST.JRNLDATE BETWEEN '" & DATEFROM & "' AND '" & DATETO & "'"
            Else
                sqlstr = "SELECT    GLPOST.BATCHNBR , GLPOST.ENTRYNBR, GLPOST.JRNLDATE,GLPJC.COMMENT," '0-3
                sqlstr &= "         GLPOST.JNLDTLREF,GLPOST.JNLDTLDESC,GLPOST.TRANSQTY," '4-6
                sqlstr &= "         GLPOST.TRANSAMT,GLAMF.ACCTFMTTD,GLPOST.TRANSNBR," '7-9
                sqlstr &= "         GLPOST.POSTINGSEQ,GLPOST.CNTDETAIL,GLPOST.FISCALYR,GLPOST.FISCALPERD,GLJEDO.VALUE,FMSVLACC.VTYPE" '10-13
                sqlstr &= " FROM    ((GLPOST "
                sqlstr &= "         LEFT OUTER JOIN GLPJC ON (GLPOST.TRANSNBR = GLPJC.TRANSNBR) AND (GLPOST.ENTRYNBR = GLPJC.ENTRYNBR) AND (GLPOST.BATCHNBR = GLPJC.BATCHNBR)) "
                sqlstr &= "         INNER JOIN GLAMF ON GLPOST.ACCTID = GLAMF.ACCTID)"
                sqlstr &= "         LEFT JOIN  GLJEDO ON (GLJEDO.BATCHNBR=GLPOST.BATCHNBR) AND (GLJEDO.JOURNALID=GLPOST.ENTRYNBR) AND (GLJEDO.OPTFIELD ='VOUCHER')"

                If ComDatabase = "PVSW" Then
                    sqlstr &= "         LEFT JOIN " & ConnVAT.Split(";")(2).Split("=")(1) & "." & "FMSVLACC FMSVLACC ON (FMSVLACC.ACCTVAT = GLAMF.ACCTFMTTD) "
                    sqlstr &= " WHERE   (GLPOST.SRCELEDGER = 'GL') AND (GLPOST.SRCETYPE <> 'NV')"
                    sqlstr &= "         AND (GLPOST.JRNLDATE >= '" & DATEFROM & "' AND GLPOST.JRNLDATE <= '" & DATETO & "')"
                    sqlstr &= "         AND (GLAMF.ACCTFMTTD IN (select ACCTVAT from " & ConnVAT.Split(";")(2).Split("=")(1) & "." & "FMSVLACC))"

                Else
                    sqlstr &= "         LEFT JOIN " & ConnVAT.Split(";")(2).Split("=")(1) & ".dbo." & "FMSVLACC FMSVLACC ON (FMSVLACC.ACCTVAT COLLATE Thai_CI_AS = GLAMF.ACCTFMTTD) "
                    sqlstr &= " WHERE   (GLPOST.SRCELEDGER = 'GL') AND (GLPOST.SRCETYPE <> 'NV') "
                    sqlstr &= "         AND (GLPOST.JRNLDATE >= '" & DATEFROM & "' AND GLPOST.JRNLDATE <= '" & DATETO & "')"
                    sqlstr &= "         AND (GLAMF.ACCTFMTTD IN (select ACCTVAT  COLLATE Thai_CI_AS from " & ConnVAT.Split(";")(2).Split("=")(1) & ".dbo." & "FMSVLACC))"

                End If

                ' หากมีการใช้ GL Reverse
                If GLREVERSE = True Then
                    sqlstr &= "     AND (GLPOST.BATCHNBR+':'+GLPOST.ENTRYNBR+':'+convert(varchar(8),GLPOST .JRNLDATE)) not in  (select (TMP1.BATCHNBR+':'+TMP1.JOURNALID+':'+convert(varchar(8),GLPOST .JRNLDATE)) from ( " & Environment.NewLine
                    sqlstr &= "     select BATCHNBR ,JOURNALID,TRANSNBR " & Environment.NewLine
                    sqlstr &= "     from GLJEC   " & Environment.NewLine
                    sqlstr &= "     group by JECOMMENT ,BATCHNBR ,JOURNALID,TRANSNBR ) as TMP1 inner join  " & Environment.NewLine
                    sqlstr &= " GLPOST on ( TMP1.BATCHNBR=GLPOST.BATCHNBR and GLPOST.ENTRYNBR=TMP1.JOURNALID and TMP1.TRANSNBR =GLPOST.TRANSNBR) inner join " & Environment.NewLine
                    sqlstr &= " GLJEC on ( TMP1.BATCHNBR=GLJEC.BATCHNBR and GLJEC.JOURNALID=TMP1.JOURNALID and TMP1.TRANSNBR =GLJEC.TRANSNBR)  " & Environment.NewLine

                    If ComDatabase = "PVSW" Then
                        sqlstr &= " where GLAMF.ACCTFMTTD in (select ACCTVAT from " & ConnVAT.Split(";")(2).Split("=")(1) & "." & "FMSVLACC) and SRCELEDGER ='GL' " & Environment.NewLine
                    Else
                        sqlstr &= " where GLAMF.ACCTFMTTD in (select ACCTVAT COLLATE Thai_CI_AS  from " & ConnVAT.Split(";")(2).Split("=")(1) & ".dbo." & "FMSVLACC) and SRCELEDGER ='GL' " & Environment.NewLine

                    End If
                    sqlstr &= " group by TMP1.BATCHNBR ,TMP1.JOURNALID,GLPOST.JRNLDATE) " & Environment.NewLine
                End If

                sqlstr &= " GROUP BY GLPOST.BATCHNBR, GLPOST.ENTRYNBR, GLPOST.JRNLDATE, GLPJC.COMMENT, GLPOST.JNLDTLREF, GLPOST.JNLDTLDESC, GLPOST.TRANSQTY, " & _
                          " GLPOST.TRANSAMT, GLAMF.ACCTFMTTD, GLPOST.TRANSNBR, GLPOST.POSTINGSEQ, GLPOST.CNTDETAIL, GLPOST.FISCALYR,  " & _
                          " GLPOST.FISCALPERD, GLJEDO.VALUE,FMSVLACC.VTYPE"
            End If
        Else
            If ComDatabase = "ORCL" Then
                sqlstr = " SELECT   GLJEH.BATCHID as BATCHNBR, GLJEH.BTCHENTRY as ENTRYNBR, GLJED.TRANSDATE as JRNLDATE, GLPJC.COMMENT, " & Environment.NewLine
                sqlstr &= "         GLJED.TRANSREF as JNLDTLREF, GLJED.TRANSDESC as JNLDTLDESC, GLJEH.JRNLQTY as TRANSQTY,   " & Environment.NewLine
                sqlstr &= "         GLJED.TRANSAMT as TRANSAMT, GLAMF.ACCTFMTTD, GLJED.TRANSNBR as TRANSNBR,GLJEH.BATCHID as POSTINGSEQ,GLJEH.BTCHENTRY as CNTDETAIL,  GLJEH.FSCSYR as FISCALYR , GLJEH.FSCSPERD as FISCALPERD,  " & Environment.NewLine
                sqlstr &= "         GLJEDO.VALUE  " & Environment.NewLine
                sqlstr &= " FROM    GLJEH ,GLJED,GLPJC ,GLAMF ,GLJEDO  " & Environment.NewLine
                sqlstr &= " WHERE GLJEH.BATCHID =GLJED.BATCHNBR and GLJEH.BTCHENTRY =GLJED.JOURNALID  " & Environment.NewLine
                sqlstr &= " 		and GLJED.TRANSNBR = GLPJC.TRANSNBR AND GLJED.JOURNALID = GLPJC.ENTRYNBR AND GLJEH.BATCHID = GLPJC.BATCHNBR  " & Environment.NewLine
                sqlstr &= " 		and GLJED.ACCTID = GLAMF.ACCTID " & Environment.NewLine
                sqlstr &= " 		and  GLJEDO.BATCHNBR =GLJEH.BATCHID AND GLJEDO.JOURNALID = GLJED.JOURNALID  AND GLJEDO.OPTFIELD = 'VOUCHER'  " & Environment.NewLine
                sqlstr &= " 		and (GLJEH.SRCELEDGER = 'GL') AND (GLJEH.SRCETYPE <> 'NV') AND (GLJED.TRANSDATE BETWEEN '" & DATEFROM & "' AND '" & DATETO & "')  " & Environment.NewLine
                sqlstr &= " GROUP BY GLJEH.BATCHID, GLJEH.BTCHENTRY, GLJED.TRANSDATE, GLPJC.COMMENT, GLJED.TRANSREF, GLJED.TRANSDESC , GLJEH.JRNLQTY ,   " & Environment.NewLine
                sqlstr &= "         GLJED.TRANSAMT , GLAMF.ACCTFMTTD, GLJED.TRANSNBR,  GLJEH.FSCSYR , GLJEH.FSCSPERD ,   " & Environment.NewLine
                sqlstr &= "         GLJEDO.VALUE  " & Environment.NewLine
            Else
                sqlstr = " SELECT   GLJEH.BATCHID as BATCHNBR, GLJEH.BTCHENTRY as ENTRYNBR, GLJED.TRANSDATE as JRNLDATE, GLPJC.COMMENT, GLJED.TRANSREF as JNLDTLREF, GLJED.TRANSDESC as JNLDTLDESC, GLJEH.JRNLQTY as TRANSQTY,   " & Environment.NewLine
                sqlstr &= "         GLJED.TRANSAMT as TRANSAMT, GLAMF.ACCTFMTTD, GLJED.TRANSNBR as TRANSNBR,'0' as POSTINGSEQ,'0'as CNTDETAIL,  GLJEH.FSCSYR as FISCALYR , GLJEH.FSCSPERD as FISCALPERD,  " & Environment.NewLine
                sqlstr &= "         GLJEDO.VALUE  " & Environment.NewLine
                sqlstr &= " FROM    GLJEH " & Environment.NewLine
                sqlstr &= " 		INNER JOIN GLJED ON GLJEH.BATCHID = GLJED.BATCHNBR AND GLJEH.BTCHENTRY = GLJED.JOURNALID " & Environment.NewLine
                sqlstr &= "         LEFT OUTER JOIN GLPJC ON GLJED.TRANSNBR = GLPJC.TRANSNBR AND GLJED.JOURNALID = GLPJC.ENTRYNBR AND GLJEH.BATCHID = GLPJC.BATCHNBR " & Environment.NewLine
                sqlstr &= "         INNER JOIN GLAMF ON GLJED.ACCTID = GLAMF.ACCTID " & Environment.NewLine
                sqlstr &= "         LEFT OUTER JOIN GLJEDO ON GLJEDO.BATCHNBR = GLJEH.BATCHID AND GLJEDO.JOURNALID = GLJED.JOURNALID  AND GLJEDO.OPTFIELD = 'VOUCHER' " & Environment.NewLine
                sqlstr &= " WHERE   (GLJEH.SRCELEDGER = 'GL') AND (GLJEH.SRCETYPE <> 'NV') AND (GLJED.TRANSDATE >= '" & DATEFROM & "' AND GLJED.TRANSDATE <= '" & DATETO & "') " & Environment.NewLine
                sqlstr &= " GROUP BY GLJEH.BATCHID, GLJEH.BTCHENTRY, GLJED.TRANSDATE, GLPJC.COMMENT, GLJED.TRANSREF, GLJED.TRANSDESC , GLJEH.JRNLQTY ,  " & Environment.NewLine
                sqlstr &= "         GLJED.TRANSAMT , GLAMF.ACCTFMTTD, GLJED.TRANSNBR,  GLJEH.FSCSYR , GLJEH.FSCSPERD ,  " & Environment.NewLine
                sqlstr &= "         GLJEDO.VALUE " & Environment.NewLine
            End If
        End If
        rs.Open(sqlstr, cnACCPAC)

        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessGL> Select GL Complete")
        End If

        'Catch ex As Exception
        '    WriteLog("<LoadDBScript>" & ex.Message.ToString())
        'End Try

        Do While rs.options.Rs.EOF = False

            loCls = New clsVat
            With loCls
                ErrMsg = "Ledger = GL Doc No.=" & Trim(rs.options.Rs.Collect(4))
                ErrMsg = ErrMsg & " Batch No.=" & Trim(rs.options.Rs.Collect(0))
                ErrMsg = ErrMsg & " Entry No.=" & Trim(rs.options.Rs.Collect(1))

                .INVDATE = IIf(ISNULL(rs.options.Rs, 2), 0, rs.options.Rs.Collect(2)) 'GLPOST.JRNLDATE

                If IsDBNull(rs.options.Rs.Collect(3)) = False Then
                    .VATCOMMENT = (Trim(rs.options.Rs.Collect(3)))
                Else
                    .VATCOMMENT = ""
                End If
                loTAXID = ""
                loBRANCH = ""

                'ดึง Vat จาก Comment
                If OPT_GL = False Or (Trim(.VATCOMMENT).Split(",")).Length >= 3 Then
                    strsplit = Split(.VATCOMMENT, ",")
                    If strsplit.Length > 3 Then
                        If strsplit(3).Trim.ToUpper = "LTD." Or strsplit(3).Trim.ToUpper = "LTD" Then
                            strsplit(2) &= ",LTD."
                            strsplit(3) = ""
                            If strsplit.Length > 4 Then
                                strsplit(3) = strsplit(4)
                            End If
                            If strsplit.Length > 5 Then
                                strsplit(4) = strsplit(5)
                            End If
                        End If
                    End If
                    For losplit = LBound(strsplit) To UBound(strsplit)
                        splittxt = Trim(strsplit(losplit))

                        Select Case losplit
                            Case 0
                                .IDINVC = splittxt
                            Case 1
                                splittxt = Replace(splittxt, "-", "/")
                                .INVDATE = TryDate2(ZeroDate(splittxt, "dd/MM/yyyy", "yyyyMMdd"), "dd/MM/yyyy", "yyyyMMdd")
                                If .INVDATE = 0 Then
                                    .INVDATE = TryDate2(splittxt, "dd/MM/yyyy", "yyyyMMdd")
                                End If

                            Case 2
                                .INVNAME = .INVNAME & Trim(splittxt)
                            Case 3
                                If Use_TAX_GL = TAXFROMGL.COMMENT4 Then
                                    loTAXID = Trim(splittxt)
                                ElseIf Use_BRANCH_GL = TAXFROMGL.COMMENT4 Then
                                    loBRANCH = Trim(splittxt)
                                Else
                                    .INVNAME = .INVNAME & "," & Trim(splittxt)
                                End If

                            Case 4
                                If Use_TAX_GL = TAXFROMGL.COMMENT5 Then
                                    loTAXID = Trim(splittxt)
                                ElseIf Use_BRANCH_GL = TAXFROMGL.COMMENT5 Then
                                    loBRANCH = Trim(splittxt)
                                Else
                                    .INVNAME = .INVNAME & "," & Trim(splittxt)
                                End If

                        End Select
                    Next
                Else
                    ' ให้ดึง VAT จาก GL OPF Field
                    Dim RS_WHT As New BaseClass.BaseODBCIO("SELECT isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from GLJEDO where OPTFIELD in ('TAXBASE','TAXINVNO','TAXDATE','TAXNAME','TAXID','BRANCH') and BATCHNBR ='" & rs.options.Rs.Collect("BATCHNBR") & "' and JOURNALID ='" & rs.options.Rs.Collect("ENTRYNBR") & "' and TRANSNBR =" & rs.options.Rs.Collect("TRANSNBR") & "", cnACCPAC)
                    For iwht As Integer = 0 To RS_WHT.options.QueryDT.Rows.Count - 1

                        If RS_WHT.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXBASE" Then
                            If RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE").ToString.Trim <> "" Then
                                .INVAMT = RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE")
                            End If
                        ElseIf RS_WHT.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXINVNO" Then
                            .IDINVC = Trim(RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE"))
                        ElseIf RS_WHT.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXDATE" Then
                            splittxt = Replace(Trim(RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE")), "-", "/")
                            If splittxt <> "" Then
                                If splittxt <> "" Then
                                    tmpDate = ""
                                    strsplit2 = Split(splittxt, "/")
                                    For losplit2 = UBound(strsplit2) To LBound(strsplit2) Step -1
                                        splittxt2 = Trim(strsplit2(losplit2))
                                        Select Case losplit2
                                            Case 2
                                                If Len(splittxt2) < 4 Then
                                                    tmpDate = tmpDate & Mid(CStr(Year(Now)), 1, 4 - Len(splittxt2)) & splittxt2
                                                Else
                                                    tmpDate = tmpDate & splittxt2
                                                End If
                                            Case Else
                                                tmpDate = tmpDate & VB6.Format(splittxt2, "00")
                                        End Select
                                    Next
                                    If Not (IsNumeric(tmpDate)) OrElse Len(tmpDate) < 8 Then
                                        .INVDATE = CDbl(StrYearF)
                                    Else
                                        .INVDATE = CDbl(tmpDate)
                                    End If
                                End If
                            Else
                                .MARK = "F"
                            End If
                        ElseIf RS_WHT.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXNAME" Then

                            Dim ChkOPTF_APTAXNAME As New BaseClass.BaseODBCIO("SELECT IDVEND,VENDNAME,RMITNAME FROM APVEN INNER JOIN APVNR ON APVEN.VENDORID = APVNR.IDVEND WHERE IDVEND = '" & Trim(RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE")) & "'", cnACCPAC)
                            Dim ChkOPTF_ARTAXNAME As New BaseClass.BaseODBCIO("SELECT ARCUS.IDCUST,NAMECUST,NAMELOCN FROM ARCUS INNER JOIN ARCSP ON ARCUS.IDCUST = ARCSP.IDCUST WHERE IDCUST = '" & Trim(RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE")) & "'", cnACCPAC)

                            If ChkOPTF_APTAXNAME.options.QueryDT.Rows.Count > 0 Then
                                If Use_LEGALNAME_inAP = True Then
                                    .INVNAME = ChkOPTF_APTAXNAME.options.QueryDT.Rows(0).Item("RMITNAME").ToString
                                Else
                                    .INVNAME = ChkOPTF_APTAXNAME.options.QueryDT.Rows(0).Item("VENDNAME").ToString
                                End If
                            ElseIf ChkOPTF_ARTAXNAME.options.QueryDT.Rows.Count > 0 Then
                                If Use_LEGALNAME_inAR = True Then
                                    .INVNAME = ChkOPTF_ARTAXNAME.options.QueryDT.Rows(0).Item("NAMELOCN").ToString
                                Else
                                    .INVNAME = ChkOPTF_ARTAXNAME.options.QueryDT.Rows(0).Item("NAMECUST").ToString
                                End If
                            Else
                                .INVNAME = Trim(RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE"))
                            End If


                        ElseIf RS_WHT.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXID" Then
                            If RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE").Trim = "" Then
                                .TAXID = ""
                            Else
                                .TAXID = Trim(RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE"))
                            End If

                        ElseIf RS_WHT.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "BRANCH" Then
                            If RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE").Trim = "" Then
                                .Branch = ""
                            Else
                                .Branch = Trim(RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE"))
                            End If

                        End If


                    Next
                    .VATCOMMENT = .IDINVC & "," & .INVDATE & "," & .INVNAME
                End If

                .TXDATE = IIf(ISNULL(rs.options.Rs, 2), 0, rs.options.Rs.Collect(2)) 'GLPOST.JRNLDATE
                If .TXDATE = 0 Then .TXDATE = .INVDATE

                loMIC = .INVNAME

                If Trim(.TAXID) = "" And Trim(.Branch) = "" Then

                    If Use_TAX_GL = TAXFROMGL.VEN_CUS Then
                        If Trim(rs.options.Rs.Collect("VTYPE")) = 1 Then
                            If FindMiscCodeInAPVEN(Trim(.INVNAME), "AP") = "0" Or FindMiscCodeInAPVEN(Trim(.INVNAME), "AP") = "" Then
                                .TAXID = ""

                            Else
                                .TAXID = FindTaxIDBRANCH(Trim(.INVNAME), "AP", "TAXID")
                                .INVNAME = FindMiscCodeInAPVEN(Trim(.INVNAME), "AP")
                            End If
                        ElseIf Trim(rs.options.Rs.Collect("VTYPE")) = 2 Then
                            If FindMiscCodeInARCUS(Trim(.INVNAME)) = "0" Or FindMiscCodeInARCUS(Trim(.INVNAME)) = "" Then
                                .TAXID = ""
                            Else
                                .TAXID = FindTaxIDBRANCH(Trim(.INVNAME), "AR", "TAXID")
                                .INVNAME = FindMiscCodeInARCUS(Trim(.INVNAME))
                            End If
                        End If
                    ElseIf Use_TAX_GL = TAXFROMGL.OPF Then
                        Dim RS_WHT As New BaseClass.BaseODBCIO("SELECT isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from GLJEDO where OPTFIELD in ('TAXID','BRANCH') and BATCHNBR ='" & rs.options.Rs.Collect("BATCHNBR") & "' and JOURNALID ='" & rs.options.Rs.Collect("ENTRYNBR") & "' and TRANSNBR =" & rs.options.Rs.Collect("TRANSNBR") & "", cnACCPAC)
                        For iwht As Integer = 0 To RS_WHT.options.QueryDT.Rows.Count - 1
                            If RS_WHT.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXID" Then
                                .TAXID = Trim(RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE"))
                            ElseIf RS_WHT.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "BRANCH" Then
                            End If
                        Next
                    Else
                        .TAXID = loTAXID
                    End If

                    If Use_BRANCH_GL = TAXFROMGL.VEN_CUS Then
                        If Trim(rs.options.Rs.Collect("VTYPE")) = 1 Then
                            If FindMiscCodeInAPVEN(Trim(loMIC), "AP") = "0" Or FindMiscCodeInAPVEN(Trim(loMIC), "AP") = "" Then
                                .Branch = ""
                            Else
                                .Branch = FindTaxIDBRANCH(Trim(loMIC), "AP", "BRANCH")
                                .INVNAME = FindMiscCodeInAPVEN(Trim(loMIC), "AP")
                            End If
                        ElseIf Trim(rs.options.Rs.Collect("VTYPE")) = 2 Then
                            If FindMiscCodeInARCUS(Trim(loMIC)) = "0" Or FindMiscCodeInARCUS(Trim(loMIC)) = "" Then
                                .Branch = ""
                            Else
                                .Branch = FindTaxIDBRANCH(Trim(loMIC), "AR", "BRANCH")
                                .INVNAME = FindMiscCodeInARCUS(Trim(loMIC))
                            End If
                        End If
                    ElseIf Use_BRANCH_GL = TAXFROMGL.OPF Then
                        Dim RS_WHT As New BaseClass.BaseODBCIO("SELECT isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from GLJEDO where OPTFIELD in ('TAXID','BRANCH') and BATCHNBR ='" & rs.options.Rs.Collect("BATCHNBR") & "' and JOURNALID ='" & rs.options.Rs.Collect("ENTRYNBR") & "' and TRANSNBR =" & rs.options.Rs.Collect("TRANSNBR") & "", cnACCPAC)
                        For iwht As Integer = 0 To RS_WHT.options.QueryDT.Rows.Count - 1
                            If RS_WHT.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXID" Then
                            ElseIf RS_WHT.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "BRANCH" Then
                                .TAXID = Trim(RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE"))
                            End If
                        Next
                    Else

                        .Branch = loBRANCH
                    End If

                End If


                If Trim(.INVNAME) = "" Then
                    .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect(5)), "", Trim(rs.options.Rs.Collect(5))) 'GLPOST.JNLDTLDESC
                End If
                If Len(CStr(.INVDATE)) > 8 Then .INVDATE = Left(CStr(.INVDATE), 8)
                If Len(CStr(.TXDATE)) > 8 Then .TXDATE = Left(CStr(.TXDATE), 8)

                .DOCNO = IIf(IsDBNull(rs.options.Rs.Collect(4)), "", (rs.options.Rs.Collect(4))) 'GLPOST.JNLDTLREF
                .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7)) 'GLPOST.TRANSAMT

                If .INVAMT = 0 Then
                    .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(6)), 0, rs.options.Rs.Collect(6)) 'GLPOST.TRANSQTY
                End If

                If VNO_GL = True Then

                    Dim VNO_Detail As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from GLJEDO  where OPTFIELD in ('VOUCHER') and BATCHNBR ='" & rs.options.Rs.Collect("BATCHNBR") & "' and JOURNALID ='" & rs.options.Rs.Collect("ENTRYNBR") & "' and TRANSNBR =" & rs.options.Rs.Collect("TRANSNBR") & "", cnACCPAC)
                    For iwht As Integer = 0 To VNO_Detail.options.QueryDT.Rows.Count - 1
                        If VNO_Detail.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "VOUCHER" Then
                            If Trim(VNO_Detail.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                .DOCNO = Trim(VNO_Detail.options.QueryDT.Rows(iwht).Item("VALUE"))
                            End If

                        End If
                    Next

                End If


                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect(8)), "", (rs.options.Rs.Collect(8))) 'GLAMF.ACCTFMTTD
                .Batch = CDbl(IIf(IsDBNull(rs.options.Rs.Collect(0)), 0, (rs.options.Rs.Collect(0)))) 'GLPOST.BATCHNBR
                .Entry = CDbl(IIf(IsDBNull(rs.options.Rs.Collect(1)), 0, (rs.options.Rs.Collect(1)))) 'GLPOST.ENTRYNBR
                .TRANSNBR = IIf(IsDBNull(rs.options.Rs.Collect(9)), "", (rs.options.Rs.Collect(9))) 'GLPOST.TRANSNBR
                .Runno = IIf(IsDBNull(rs.options.Rs.Collect("Value")), "", rs.options.Rs.Collect("Value"))

                ' Insert ข้อมูลลง FMSVATTEMP
                sqlstr = InsVat(0, CStr(.INVDATE), CStr(.TXDATE), Trim(.IDINVC), Trim(.DOCNO), "", Trim(.INVNAME), .INVAMT, .INVTAX, Trim(.LOCID), 1, .RATE_Renamed, Trim(.TTYPE), Trim(.ACCTVAT), "GL", Trim(.Batch), Trim(.Entry), Trim(.MARK), Trim(.VATCOMMENT), Trim(.CBRef), .TRANSNBR, Trim(.Runno), "", "", "", "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, 0)

            End With

            'Debug.Print sqlstr
            cnVAT.Execute(sqlstr)
            ErrMsg = ""
            Code = ""
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing

        'transfer data
        rs = New BaseClass.BaseODBCIO
        sqlstr = "SELECT FMSVATTEMP.INVDATE,FMSVATTEMP.TXDATE,FMSVATTEMP.IDINVC,FMSVATTEMP.DOCNO," '0-3
        sqlstr = sqlstr & " FMSVATTEMP.INVNAME,FMSVATTEMP.INVAMT,FMSVATTEMP.INVTAX," '4-6
        sqlstr = sqlstr & " FMSVLACC.LOCID,FMSVATTEMP.VTYPE,FMSVATTEMP.RATE,FMSTAXGL.TTYPE," '7-10
        sqlstr = sqlstr & " FMSVATTEMP.ACCTVAT,FMSVATTEMP.SOURCE,FMSVATTEMP.BATCH,FMSVATTEMP.ENTRY," '11-14
        sqlstr = sqlstr & " FMSVATTEMP.MARK,FMSVATTEMP.VATCOMMENT,FMSVATTEMP.TRANSNBR,FMSVATTEMP.RUNNO, " '15-17
        sqlstr = sqlstr & " FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH,FMSVATTEMP.Code "
        If ComDatabase = "ORCL" Then
            sqlstr = sqlstr & " FROM FMSVATTEMP,FMSTAXGL,FMSVLACC"
            sqlstr = sqlstr & "  WHERE FMSVATTEMP.ACCTVAT = FMSTAXGL.ACCTVAT"
            sqlstr = sqlstr & " AND FMSVATTEMP.ACCTVAT = FMSVLACC.ACCTVAT"
        Else
            sqlstr = sqlstr & " FROM FMSVATTEMP"
            sqlstr = sqlstr & " INNER JOIN FMSTAXGL ON FMSVATTEMP.ACCTVAT = FMSTAXGL.ACCTVAT"
            sqlstr = sqlstr & " INNER JOIN FMSVLACC ON FMSVATTEMP.ACCTVAT = FMSVLACC.ACCTVAT"
        End If

        rs.Open(sqlstr, cnVAT)

        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessGL> Select FMSVATTEMP Complete")
        End If



        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls
                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect(0)), 0, rs.options.Rs.Collect(0))
                .TXDATE = IIf(IsDBNull(rs.options.Rs.Collect(1)), 0, rs.options.Rs.Collect(1))
                .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", (rs.options.Rs.Collect(2)))
                .DOCNO = IIf(IsDBNull(rs.options.Rs.Collect(3)), "", (rs.options.Rs.Collect(3)))
                .INVNAME = Trim(IIf(IsDBNull(rs.options.Rs.Collect(4)), "", (rs.options.Rs.Collect(4))))
                .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(5)), 0, rs.options.Rs.Collect(5))
                .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(6)), 0, rs.options.Rs.Collect(6))
                .LOCID = IIf(IsDBNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7))
                .VTYPE = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8))
                .TTYPE = IIf(IsDBNull(rs.options.Rs.Collect(10)), 0, rs.options.Rs.Collect(10))
                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect(11)), "", (rs.options.Rs.Collect(11)))
                .Source = IIf(IsDBNull(rs.options.Rs.Collect(12)), "", (rs.options.Rs.Collect(12)))
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(13)), "", (rs.options.Rs.Collect(13)))
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(14)), "", (rs.options.Rs.Collect(14)))
                .MARK = IIf(IsDBNull(rs.options.Rs.Collect(15)), "", (rs.options.Rs.Collect(15)))
                .VATCOMMENT = IIf(IsDBNull(rs.options.Rs.Collect(16)), "", (rs.options.Rs.Collect(16)))
                .TRANSNBR = IIf(IsDBNull(rs.options.Rs.Collect(17)), "", (rs.options.Rs.Collect(17)))
                .Runno = IIf(IsDBNull(rs.options.Rs.Collect("Runno")), "", rs.options.Rs.Collect("Runno"))
                Code = IIf(IsDBNull(rs.options.Rs.Collect("Code")), "", rs.options.Rs.Collect("Code"))


                'Math.Abs >> returns the absolute value of a number
                If .TTYPE = 1 Then 'Purchase
                    If .INVTAX > 0 Then
                        .INVAMT = Math.Abs(.INVAMT)
                        .INVTAX = Math.Abs(.INVTAX)
                    ElseIf .INVTAX < 0 Then
                        .INVAMT = Math.Abs(.INVAMT) * -1
                        .INVTAX = Math.Abs(.INVTAX) * -1
                    End If
                ElseIf .TTYPE = 2 Then  'Sale
                    If .INVTAX > 0 Then
                        .INVAMT = Math.Abs(.INVAMT) * -1
                        .INVTAX = Math.Abs(.INVTAX) * -1
                    ElseIf .INVTAX < 0 Then
                        .INVAMT = Math.Abs(.INVAMT)
                        .INVTAX = Math.Abs(.INVTAX)
                    End If
                End If

                If .INVAMT = 0 Then
                    .RATE_Renamed = 0
                Else
                    .RATE_Renamed = FormatNumber(((.INVTAX / .INVAMT) * 100), 1)
                End If



                'If .INVAMT = 0 Then
                '    .RATE_Renamed = 0
                'Else
                '    .RATE_Renamed = CDbl(FormatNumber((.INVTAX / .INVAMT) * 100, 0))
                'End If

                .TAXID = IIf(IsDBNull(rs.options.Rs.Collect("TAXID")), "", rs.options.Rs.Collect("TAXID"))
                .Branch = IIf(IsDBNull(rs.options.Rs.Collect("BRANCH")), "", rs.options.Rs.Collect("BRANCH"))
                .CODETAX = ""


                'Insert ข้อมูล ลง FMSVAT
                sqlstr = InsVat(1, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, .NEWDOCNO, .INVNAME, .INVAMT, .INVTAX, .LOCID, .VTYPE, .RATE_Renamed, .TTYPE, .ACCTVAT, .Source, .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, Trim(.Runno), .CODETAX, "", "", "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, 0)

            End With

            cnVAT.Execute(sqlstr)
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing
        cnACCPAC.Close()
        cnVAT.Close()
        cnACCPAC = Nothing
        cnVAT = Nothing
        Exit Sub

ErrHandler:
        WriteLog("<ProcessGL>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessGL> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
        WriteLog(ErrMsg)
    End Sub
    ' ดึงข้อมูล ที่คีย์ใน PO
    Public Sub ProcessPO()
        Dim cnACCPAC, cnVAT As ADODB.Connection
        Dim rs, rs2 As BaseClass.BaseODBCIO
        Dim sqlstr As String = ""
        Dim loCls As clsVat
        Dim iffs As clsFuntion
        Dim tmpTAXBASE, tmpTAXDATE As String
        Dim tmpTAXNAME, tmpTAXIDINVC As String
        Dim Ddbo As String
        iffs = New clsFuntion
        ErrMsg = ""

        On Error GoTo ErrHandler
        cnACCPAC = New ADODB.Connection
        cnVAT = New ADODB.Connection

        cnACCPAC.ConnectionTimeout = 60
        cnACCPAC.Open(ConnACCPAC)
        cnACCPAC.CommandTimeout = 3600

        cnVAT.ConnectionTimeout = 60
        cnVAT.Open(ConnVAT)
        cnVAT.CommandTimeout = 3600
        sqlstr = ""

        Ddbo = ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.")

        'gather data ==> Select Data Uses For Module
        'LONG:: ในส่วนดึงVat ที่AP ที่ส่งมาจาก PO ยังไม่ได้ใส่ในส่วน Getvat จาก OPF Field Header(BOCSH)
        rs = New BaseClass.BaseODBCIO
        If Use_POST Then
            If ComDatabase = "MSSQL" Then
                sqlstr = " select * from (     " & Environment.NewLine
                sqlstr &= " SELECT DISTINCT  " & Environment.NewLine
                sqlstr &= Ddbo & "APOBLTEMP.DATEINVC, " & Ddbo & "APOBLTEMP.[VALUES], " & Ddbo & "APOBLTEMP.IDINVC, " & Ddbo & "APOBLTEMP.AUDTDATE,"
                If Use_LEGALNAME_inAP = True Then
                    sqlstr &= " CASE WHEN APVNR.RMITNAME IS NULL THEN APVEN.VENDNAME ELSE APVNR.RMITNAME END AS VENDNAME, APVEN.TEXTSTRE1, '' AS IDDIST, " & Ddbo & "APOBLTEMP.AMTINVCHC,  " & Environment.NewLine
                Else
                    sqlstr &= " APVEN.VENDNAME, APVEN.TEXTSTRE1, '' AS IDDIST, " & Ddbo & "APOBLTEMP.AMTINVCHC,  " & Environment.NewLine
                End If

            Else
                sqlstr &= " SELECT DISTINCT  " & Environment.NewLine
                sqlstr &= Ddbo & "APOBLTEMP.DATEINVC, " & Ddbo & "APOBLTEMP." & Chr(34) & "VALUES" & Chr(34) & ", " & Ddbo & "APOBLTEMP.IDINVC, " & Ddbo & "APOBLTEMP.AUDTDATE, APVEN.VENDNAME, APVEN.TEXTSTRE1, '' AS IDDIST, " & Ddbo & "APOBLTEMP.AMTINVCHC,  " & Environment.NewLine
            End If

            sqlstr &= "             MIN(" & Ddbo & "APOBLTEMP.AMTBASE1HC) AS AMTEXTNDHC, " & Ddbo & "APOBLTEMP.AMTTAXHC, ISNULL(TXAUTH.ACCTRECOV, APPJD.IDACCT) AS IDACCT, " & Ddbo & "APOBLTEMP.CNTBTCH, " & Ddbo & "APOBLTEMP.CNTITEM,  " & Environment.NewLine
            sqlstr &= "             " & Ddbo & "APOBLTEMP.CODETAX1, MIN(PORCPH1.RCPNUMBER) AS Expr1, POCRNH1.RETNUMBER, " & Ddbo & "APOBLTEMP.IDPONBR, MIN(" & Ddbo & "APOBLTEMP.AMTBASE1HC) AS AMTGLDIST,  " & Environment.NewLine
            sqlstr &= "             " & Ddbo & "APOBLTEMP.TXTTRXTYPE, 0 AS TAXBASE, 0 AS TAXDATE, 0 AS TAXNAME, MIN(APIBD.AMTTAXREC1) AS AMTTAXREC1, 1 AS COUNTDUP, APIBHO.VALUE,  " & Environment.NewLine
            sqlstr &= "             APVEN.NAMECITY, '' AS TEXTLINE, MIN(APIBH.AMTRECTAX) AS AMTTOTTAX, '0' AS CNTLINE, POCRNH1.INVDATE, ISNULL(PORCPH2.DATEBUS,  " & Environment.NewLine
            sqlstr &= "             POCRNH1.POSTDATE) AS DATEBUS, POINVH1.VDADDRESS1, 'HEADER' AS VATAS," & Ddbo & "APOBLTEMP.IDVEND  " & Environment.NewLine
            sqlstr &= " FROM        " & Ddbo & "APOBLTEMP" & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APVEN ON " & Ddbo & "APOBLTEMP.IDVEND = APVEN.VENDORID " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN " & Ddbo & "APPJHTEMP ON " & Ddbo & "APOBLTEMP.IDINVC = " & Ddbo & "APPJHTEMP.IDINVC AND " & Ddbo & "APOBLTEMP.IDVEND = " & Ddbo & "APPJHTEMP.IDVEND " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APPJD ON " & Ddbo & "APPJHTEMP.POSTSEQNCE = APPJD.POSTSEQNCE AND " & Ddbo & "APPJHTEMP.CNTBTCH = APPJD.CNTBTCH AND " & Ddbo & "APPJHTEMP.CNTITEM = APPJD.CNTITEM " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APIBH ON " & Ddbo & "APOBLTEMP.IDVEND = APIBH.IDVEND AND " & Ddbo & "APOBLTEMP.IDINVC = APIBH.IDINVC AND " & Ddbo & "APOBLTEMP.CNTBTCH = APIBH.CNTBTCH AND " & Ddbo & "APOBLTEMP.CNTITEM = APIBH.CNTITEM " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN POINVH1 ON APIBH.DRILLDWNLK = POINVH1.INVHSEQ " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN POINVH2 ON POINVH1.INVHSEQ = POINVH2.INVHSEQ " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN PORCPH1 ON POINVH1.RCPHSEQ = PORCPH1.RCPHSEQ " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN PORCPH2 ON PORCPH1.RCPHSEQ = PORCPH2.RCPHSEQ " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN PORCPL ON PORCPH1.RCPHSEQ = PORCPL.RCPHSEQ " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN POCRNH1 ON APIBH.DRILLDWNLK = POCRNH1.CRNHSEQ " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APIBD ON APIBH.CNTBTCH = APIBD.CNTBTCH AND APIBH.CNTITEM = APIBD.CNTITEM " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APIBHO ON APIBH.CNTBTCH = APIBHO.CNTBTCH AND APIBH.CNTITEM = APIBHO.CNTITEM AND APIBHO.OPTFIELD = 'VOUCHER' " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN TXAUTH ON " & Ddbo & "APOBLTEMP.CODETAXGRP = TXAUTH.AUTHORITY " & Environment.NewLine
            'sqlstr &= "             LEFT OUTER JOIN TXAUTH ON PORCPH1.TAXGROUP = TXAUTH.AUTHORITY " & Environment.NewLine
            sqlstr &= "              LEFT OUTER JOIN APVNR ON APVNR.IDVEND = APVEN.VENDORID" & Environment.NewLine
            sqlstr &= " WHERE       (ISNULL(" & Ddbo & "APOBLTEMP.DATEINVC, 0) <> 0) AND  (" & IIf(DATEMODE = "DOCU", "POCRNH1.INVDATE", "ISNULL(PORCPH2.DATEBUS,POCRNH1.POSTDATE)") & " >= " & DATEFROM & " AND " & IIf(DATEMODE = "DOCU", "POCRNH1.INVDATE", "ISNULL(PORCPH2.DATEBUS,POCRNH1.POSTDATE)") & " <= " & DATETO & ") " & Environment.NewLine
            sqlstr &= "             AND (APPJD.IDACCT  IN " & Environment.NewLine
            sqlstr &= "             (SELECT ACCTVAT " & IIf(ComDatabase = "PVSW", "", "COLLATE Thai_CI_AS") & Environment.NewLine
            sqlstr &= "             FROM " & Ddbo & "FMSVLACC)) " & Environment.NewLine
            'sqlstr &= " AND " & Ddbo & "APOBLTEMP.AMTTAXHC<>0 " & Environment.NewLine

            sqlstr &= " GROUP BY    APIBH.DRILLDWNLK, " & Ddbo & "APOBLTEMP.DATEINVC," & Ddbo & "APOBLTEMP." & IIf(ComDatabase = "MSSQL", "[VALUES]", "" & Chr(34) & "VALUES" & Chr(34) & "") & ", " & Ddbo & "APOBLTEMP.IDINVC, " & Ddbo & "APOBLTEMP.AUDTDATE, APVEN.VENDNAME, APVEN.TEXTSTRE1, " & Ddbo & "APOBLTEMP.AMTINVCHC,  " & Environment.NewLine
            sqlstr &= Ddbo & "APOBLTEMP.AMTTAXHC, TXAUTH.ACCTRECOV, " & Ddbo & "APOBLTEMP.CNTBTCH, " & Ddbo & "APOBLTEMP.CNTITEM, " & Ddbo & "APOBLTEMP.CODETAX1, POCRNH1.RETNUMBER, " & Ddbo & "APOBLTEMP.IDPONBR, " & Ddbo & "APOBLTEMP.TXTTRXTYPE,  " & Environment.NewLine
            sqlstr &= "             APIBHO.VALUE, APVEN.NAMECITY, APPJD.IDACCT, " & Ddbo & "APPJHTEMP.POSTSEQNCE, POCRNH1.INVDATE, PORCPH2.DATEBUS, POINVH1.VDADDRESS1, POCRNH1.POSTDATE ," & Ddbo & "APOBLTEMP.IDVEND " & Environment.NewLine

            sqlstr &= " UNION ALL " & Environment.NewLine

            sqlstr &= " SELECT DISTINCT  " & Environment.NewLine

            If ComDatabase = "MSSQL" Then
                sqlstr &= Ddbo & "APOBLTEMP.DATEINVC, " & Ddbo & "APOBLTEMP.[VALUES], " & Ddbo & "APOBLTEMP.IDINVC, " & Ddbo & "APOBLTEMP.AUDTDATE, "
                If Use_LEGALNAME_inAP = True Then
                    sqlstr &= "CASE WHEN APVNR.RMITNAME IS NULL THEN APVEN.VENDNAME ELSE APVNR.RMITNAME END AS VENDNAME, APVEN.TEXTSTRE1, '' AS IDDIST, " & Ddbo & "APOBLTEMP.AMTINVCHC,  " & Environment.NewLine

                Else
                    sqlstr &= "APVEN.VENDNAME, APVEN.TEXTSTRE1, '' AS IDDIST, " & Ddbo & "APOBLTEMP.AMTINVCHC,  " & Environment.NewLine

                End If
            
            Else
                sqlstr &= Ddbo & "APOBLTEMP.DATEINVC, " & Ddbo & "APOBLTEMP." & Chr(34) & "VALUES" & Chr(34) & ", " & Ddbo & "APOBLTEMP.IDINVC, " & Ddbo & "APOBLTEMP.AUDTDATE, APVEN.VENDNAME, APVEN.TEXTSTRE1, '' AS IDDIST, " & Ddbo & "APOBLTEMP.AMTINVCHC,  " & Environment.NewLine

            End If

            sqlstr &= "             (CASE WHEN APIBT.CNTLINE IS NULL THEN MIN(" & Ddbo & "APOBLTEMP.AMTBASE1HC) ELSE SUM(" & Ddbo & "APOBLTEMP.AMTBASE1HC) END) AS AMTEXTNDHC, " & Ddbo & "APOBLTEMP.AMTTAXHC,  " & Environment.NewLine
            sqlstr &= "             ISNULL(TXAUTH.ACCTRECOV, APPJD.IDACCT) AS IDACCT, " & Ddbo & "APOBLTEMP.CNTBTCH, " & Ddbo & "APOBLTEMP.CNTITEM, " & Ddbo & "APOBLTEMP.CODETAX1, MIN(PORCPH1.RCPNUMBER) AS Expr1,  " & Environment.NewLine
            sqlstr &= "             POCRNH1.RETNUMBER, " & Ddbo & "APOBLTEMP.IDPONBR, (CASE WHEN APIBT.CNTLINE IS NULL THEN MIN(" & Ddbo & "APOBLTEMP.AMTBASE1HC) ELSE SUM(" & Ddbo & "APOBLTEMP.AMTBASE1HC) END) AS AMTGLDIST, " & Environment.NewLine
            sqlstr &= "             " & Ddbo & "APOBLTEMP.TXTTRXTYPE, 0 AS TAXBASE, 0 AS TAXDATE, 0 AS TAXNAME, MIN(APIBD.AMTTAXREC1) AS AMTTAXREC1, 1 AS COUNTDUP, APIBHO.VALUE,  " & Environment.NewLine
            sqlstr &= "             APVEN.NAMECITY, APIBT.TEXTLINE, (CASE WHEN APIBT.CNTLINE IS NULL THEN MIN(APIBD.AMTTOTTAX) ELSE " & Environment.NewLine
            sqlstr &= "             (SELECT TOP 1 APPJD.AMTEXTNDHC " & Environment.NewLine
            sqlstr &= "             FROM APPJD " & Environment.NewLine
            sqlstr &= "             WHERE POSTSEQNCE = " & Ddbo & "APPJHTEMP.POSTSEQNCE AND CNTBTCH = " & Ddbo & "APOBLTEMP.CNTBTCH AND CNTITEM = " & Ddbo & "APOBLTEMP.CNTITEM AND CNTLINE = APIBT.CNTLINE) END) AS AMTTOTTAX, " & Environment.NewLine
            sqlstr &= "             APPJD.CNTLINE, POCRNH1.INVDATE, ISNULL(PORCPH2.DATEBUS, POCRNH1.POSTDATE) AS DATEBUS, POINVH1.VDADDRESS1, " & Environment.NewLine
            sqlstr &= "             'DETAIL' AS VATAS ," & Ddbo & "APOBLTEMP.IDVEND " & Environment.NewLine
            sqlstr &= " FROM        " & Ddbo & "APOBLTEMP" & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APVEN ON " & Ddbo & "APOBLTEMP.IDVEND = APVEN.VENDORID " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN " & Ddbo & "APPJHTEMP ON " & Ddbo & "APOBLTEMP.IDINVC = " & Ddbo & "APPJHTEMP.IDINVC AND " & Ddbo & "APOBLTEMP.IDVEND = " & Ddbo & "APPJHTEMP.IDVEND " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APPJD ON " & Ddbo & "APPJHTEMP.POSTSEQNCE = APPJD.POSTSEQNCE AND " & Ddbo & "APPJHTEMP.CNTBTCH = APPJD.CNTBTCH AND " & Ddbo & "APPJHTEMP.CNTITEM = APPJD.CNTITEM " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APIBH ON " & Ddbo & "APOBLTEMP.IDVEND = APIBH.IDVEND AND " & Ddbo & "APOBLTEMP.IDINVC = APIBH.IDINVC AND " & Ddbo & "APOBLTEMP.CNTBTCH = APIBH.CNTBTCH AND " & Ddbo & "APOBLTEMP.CNTITEM = APIBH.CNTITEM " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN POINVH1 ON APIBH.DRILLDWNLK = POINVH1.INVHSEQ " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN POINVH2 ON POINVH1.INVHSEQ = POINVH2.INVHSEQ " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN PORCPH1 ON POINVH1.RCPHSEQ = PORCPH1.RCPHSEQ " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN PORCPH2 ON PORCPH1.RCPHSEQ = PORCPH2.RCPHSEQ " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN PORCPL ON PORCPH1.RCPHSEQ = PORCPL.RCPHSEQ " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN POCRNH1 ON APIBH.DRILLDWNLK = POCRNH1.CRNHSEQ " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APIBD ON APIBH.CNTBTCH = APIBD.CNTBTCH AND APIBH.CNTITEM = APIBD.CNTITEM " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APIBHO ON APIBH.CNTBTCH = APIBHO.CNTBTCH AND APIBH.CNTITEM = APIBHO.CNTITEM AND APIBHO.OPTFIELD = 'VOUCHER' " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN TXAUTH ON " & Ddbo & "APOBLTEMP.CODETAXGRP = TXAUTH.AUTHORITY " & Environment.NewLine
            'sqlstr &= "             LEFT OUTER JOIN TXAUTH ON PORCPH1.TAXGROUP = TXAUTH.AUTHORITY " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APIBT ON APIBH.CNTBTCH = APIBT.CNTBTCH AND APIBH.CNTITEM = APIBT.CNTITEM AND APIBT.CNTLINE = APIBD.CNTLINE " & Environment.NewLine
            sqlstr &= "              LEFT OUTER JOIN APVNR ON APVNR.IDVEND = APVEN.VENDORID" & Environment.NewLine

            sqlstr &= " WHERE       (ISNULL(" & Ddbo & "APOBLTEMP.DATEINVC, 0) <> 0) AND  (" & IIf(DATEMODE = "DOCU", "POCRNH1.INVDATE", "ISNULL(PORCPH2.DATEBUS,POCRNH1.POSTDATE)") & " >= " & DATEFROM & " AND " & IIf(DATEMODE = "DOCU", "POCRNH1.INVDATE", "ISNULL(PORCPH2.DATEBUS,POCRNH1.POSTDATE)") & " <= " & DATETO & ") " & Environment.NewLine
            sqlstr &= "             AND " & Ddbo & "APOBLTEMP.AMTTAXHC = 0 AND (TXAUTH.ACCTRECOV IN " & Environment.NewLine
            sqlstr &= "             (SELECT ACCTVAT " & IIf(ComDatabase = "PVSW", "", "COLLATE Thai_CI_AS") & Environment.NewLine
            sqlstr &= "             FROM " & Ddbo & "FMSVLACC)) " & Environment.NewLine

            sqlstr &= " GROUP BY    APIBH.DRILLDWNLK, " & Ddbo & "APOBLTEMP.DATEINVC, " & Ddbo & "APOBLTEMP." & IIf(ComDatabase = "MSSQL", "[VALUES]", "" & Chr(34) & "VALUES" & Chr(34) & "") & ", " & Ddbo & "APOBLTEMP.IDINVC, " & Ddbo & "APOBLTEMP.AUDTDATE, APVEN.VENDNAME, APVEN.TEXTSTRE1, " & Ddbo & "APOBLTEMP.AMTINVCHC,  " & Environment.NewLine
            sqlstr &= Ddbo & "APOBLTEMP.AMTTAXHC, TXAUTH.ACCTRECOV, " & Ddbo & "APOBLTEMP.CNTBTCH, " & Ddbo & "APOBLTEMP.CNTITEM, " & Ddbo & "APOBLTEMP.CODETAX1, POCRNH1.RETNUMBER, " & Ddbo & "APOBLTEMP.IDPONBR, " & Ddbo & "APOBLTEMP.TXTTRXTYPE,  " & Environment.NewLine
            sqlstr &= "             APIBHO.VALUE, APVEN.NAMECITY, APPJD.IDACCT, APIBT.TEXTLINE, APIBT.CNTLINE, " & Ddbo & "APPJHTEMP.POSTSEQNCE, POCRNH1.INVDATE, PORCPH2.DATEBUS,  " & Environment.NewLine
            sqlstr &= "             POINVH1.VDADDRESS1, POCRNH1.POSTDATE,APPJD.CNTLINE ," & Ddbo & "APOBLTEMP.IDVEND " & Environment.NewLine

            If ComDatabase = "MSSQL" Then
                sqlstr &= " )AS TMP1 Order by DATEINVC,CNTBTCH ,CNTITEM ,CNTLINE " & Environment.NewLine
            End If

        End If

        rs.Open(sqlstr, cnACCPAC)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessPO> Select PO Complete")
        End If


        '***************** Set Data To Temp Variable ==> Not Change
        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls
                ErrMsg = "Ledger = PO IDINVC=" & Trim(rs.options.Rs.Collect(2))
                ErrMsg = ErrMsg & " Batch No.=" & Trim(rs.options.Rs.Collect(11))
                ErrMsg = ErrMsg & " Entry No.=" & Trim(rs.options.Rs.Collect(12))
                Dim INVTAXtemp As Decimal
                Dim TEXTLINE() As String = {}
                TEXTLINE = rs.options.Rs.Collect("TEXTLINE").ToString.Split(",")
                .VATCOMMENT = rs.options.Rs.Collect("TEXTLINE").ToString.Trim

                'If Trim(rs.options.Rs.Collect(11)) = 100 And Trim(rs.options.Rs.Collect(12)) = 5 Then
                '    Dim sss As String = ""
                'End If


                If rs.options.Rs.Collect("VATAS").ToString.Trim = "HEADER" Then

                    Dim RS_WHT_Header As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from APIBHO  where OPTFIELD in ('TAXBASE','TAXINVNO','TAXDATE','TAXNAME') and CNTBTCH ='" & rs.options.Rs.Collect("CNTBTCH") & "' and CNTITEM='" & rs.options.Rs.Collect("CNTITEM") & "'", cnACCPAC)
                    .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect(0)), 0, (rs.options.Rs.Collect(0))) 'APOBL.DATEINVC
                    If Val(tmpTAXDATE) = 0 Then
                        .TXDATE = rs.options.Rs.Collect("DATEBUS")
                    Else
                        .TXDATE = CDbl(Trim(tmpTAXDATE))
                    End If

                    .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", (rs.options.Rs.Collect(2))) 'APOBL.IDINVC

                    If IsDBNull(rs.options.Rs.Collect(14)) Then
                        .DOCNO = ""
                    Else
                        .DOCNO = rs.options.Rs.Collect(14) 'PORCPH1.RCPNUMBER
                    End If

                    If Len(.DOCNO) = 0 Then .DOCNO = IIf(IsDBNull(rs.options.Rs.Collect(15)), "", rs.options.Rs.Collect(15)) 'POCRNH1.RETNUMBER

                    If Trim(tmpTAXNAME) = "0" Or Trim(tmpTAXNAME) = "" Then
                        .INVNAME = ""
                    Else
                        .INVNAME = Trim(tmpTAXNAME)
                    End If

                    If Len(.INVNAME) = 0 Then
                        .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect(4)), "", (rs.options.Rs.Collect(4))) 'APVEN.VENDNAME
                        Code = IIf(IsDBNull(rs.options.Rs.Collect("IDVEND")), "", (rs.options.Rs.Collect("IDVEND"))) 'APVEN.IDVEND

                        If .INVNAME.Trim <> "" Then
                            If ComVatName = True And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" Then 'เพิ่ม And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" กรณีที่ ชื้อ Customer ยาว แล้วต้องคีย์เพิ่มที่ช่อง TEXTSTRE1 (เจอที่ AIT)

                                .INVNAME = CStr(.INVNAME.Trim + IIf(IsDBNull(rs.options.Rs.Collect(5)), "", (Trim(rs.options.Rs.Collect(5))))) 'APVEN.TEXTSTRE1
                            End If
                        End If

                    End If
                    If IsDBNull(rs.options.Rs.Collect(6)) Then Exit Do

                    If Trim(rs.options.Rs.Collect(6)) = "VAT" Then
                        .INVAMT = Val(tmpTAXBASE)
                        .INVTAX = IIf(iffs.ISNULL(rs.options.Rs, 7), 0, rs.options.Rs.Collect(7)) ' APOBL.AMTINVCHC
                        'If .INVTAX <> 0 Then .INVTAX = IIf(IsNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8)) ' APOBL.AMTEXTNDHC
                        If .INVTAX <> 0 Then .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8)) / rs.options.Rs.Fields("CountDup").Value ' APOBL.AMTEXTNDHC
                    Else
                        .INVAMT = Val(tmpTAXBASE)
                        If .INVAMT = 0 Then
                            ' 17 เอาใช้กรณีที่ บรรทัดแรก มีภาษี บันทัดที่สองไม่มีภาษี
                            '.INVAMT = IIf(IsNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7)) ' APOBL.AMTINVCHC       '------------------ Edit by X [09/09/2005]
                            '.INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17)) '/ rs.options.Rs.Fields("CountDup") ' APIBD.AMTGLDIST      '------------------ Edit by X [09/09/2005]  ++++ wut 22/02/2008 add       / rs.options.Rs.Fields("CountDup")
                            '.INVAMT = .INVAMT - IIf(IsNull(rs.options.Rs.Collect(22)), 0, rs.options.Rs.Collect(22)) ' APIBD.AMTTAXREC1
                            '***************by kwan
                            If rs.options.Rs.Collect(23) = 1 Then
                                If rs.options.Rs.Collect(8) < rs.options.Rs.Collect(17) Then
                                    .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8))
                                Else
                                    .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17))
                                End If
                            Else
                                If rs.options.Rs.Collect(23) > 0 Then
                                    If rs.options.Rs.Collect(8) = rs.options.Rs.Collect(17) Then
                                        .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17)) ' / rs.options.Rs.Collect(23)
                                    Else
                                        If rs.options.Rs.Collect(8) < rs.options.Rs.Collect(17) Then
                                            .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8))
                                        Else
                                            .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17))
                                        End If
                                    End If
                                Else
                                    .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17))
                                End If
                            End If

                            '.INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(9)), 0, rs.options.Rs.Collect(9)) '***************by kwan

                            INVTAXtemp = IIf(IsDBNull(rs.options.Rs.Collect(9)), 0, rs.options.Rs.Collect(9)) 'APOBL.AMTTAXHC
                            .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(27)), 0, rs.options.Rs.Collect(27)) 'MIN(APIBH.AMTRECTAX) AS AMTTOTTAX

                            '--------- แก้ไขให้รองรับ VAT AVG เนื่องจากเปลี่ยนฟิลด์หยิบเป็นฟิลด์ (APIBH.AMTRECTAX) และฟิลด์ที่หยิบจะได้ตัวเลขที่เป็น บวกทั้งหมด แต่ตัวเลขต้องแสดงตามจริงของฟิลด์เดิม (APOBL.AMTTAXHC)จึงต้องเช็คเงื่อนไขเพื่อ (*-1) ------
                            If INVTAXtemp < 0 Then
                                .INVTAX = .INVTAX * (-1)
                            End If
                            '-------- end comment -----
                            If .INVTAX = 0 Then
                                .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(9)), 0, rs.options.Rs.Collect(9)) 'APOBL.AMTTAXHC
                            End If


                            If rs.options.Rs.Collect(18) = 3 And .INVTAX > 0 Then '----------- If TXTTRXTYPE = 3 (Credit Note)
                                .INVTAX = .INVTAX * (-1)
                            End If
                        End If

                        If .INVTAX < 0 Then
                            If .INVAMT > 0 Then
                                .INVAMT = .INVAMT * (-1)
                            End If
                        End If

                        If OPT_AP = True Then
                            Dim tINVAMT As Double = 0
                            Dim tTAXINVNO As String = ""
                            Dim tTAXDATE As String = ""
                            Dim tTAXNAME As String = ""
                            For iwht As Integer = 0 To RS_WHT_Header.options.QueryDT.Rows.Count - 1
                                If RS_WHT_Header.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXBASE" Then
                                    Dim Trydec As String = Trim(RS_WHT_Header.options.QueryDT.Rows(iwht).Item("VALUE"))
                                    If Decimal.TryParse(Trydec, 0) = True Then
                                        If CDec(Trydec) <> 0 Then
                                            tINVAMT = CDec(Trydec)
                                        Else
                                            tINVAMT = 0
                                        End If
                                    Else
                                        tINVAMT = 0
                                    End If
                                ElseIf RS_WHT_Header.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXINVNO" Then
                                    If Trim(RS_WHT_Header.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                        tTAXINVNO = Trim(RS_WHT_Header.options.QueryDT.Rows(iwht).Item("VALUE"))
                                    End If
                                ElseIf RS_WHT_Header.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXDATE" Then
                                    Dim myDate As String = Trim(RS_WHT_Header.options.QueryDT.Rows(iwht).Item("VALUE"))
                                    If Decimal.TryParse(myDate, 0) = True Then
                                        If CDec(myDate) <> 0 Then
                                            tTAXDATE = CDec(myDate)
                                        End If
                                    End If
                                ElseIf RS_WHT_Header.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXNAME" Then
                                    If Trim(RS_WHT_Header.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                        tTAXNAME = Trim(RS_WHT_Header.options.QueryDT.Rows(iwht).Item("VALUE"))
                                    End If
                                End If
                            Next
                            If tINVAMT <> 0 Then
                                .INVAMT = tINVAMT
                            End If
                            If tTAXINVNO <> "" Then
                                .IDINVC = tTAXINVNO
                            End If
                            If tTAXDATE <> "" Then
                                .INVDATE = tTAXDATE
                            End If
                            If tTAXNAME <> "" Then
                                .INVNAME = tTAXNAME
                            End If
                        End If
                        .TRANSNBR = ""
                    End If

                ElseIf rs.options.Rs.Collect("VATAS").ToString.Trim = "DETAIL" Then

                    If Trim(rs.options.Rs.Collect(6).ToString) = "VAT" Then
                        .INVAMT = Val(tmpTAXBASE)
                        .INVTAX = IIf(iffs.ISNULL(rs.options.Rs, 7), 0, rs.options.Rs.Collect(7)) ' APOBL.AMTINVCHC
                        If .INVTAX <> 0 Then .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8)) / rs.options.Rs.Fields("CountDup").Value ' APOBL.AMTEXTNDHC
                    Else
                        .INVAMT = Val(tmpTAXBASE)
                        If .INVAMT = 0 Then
                            ' 17 เอาใช้กรณีที่ บรรทัดแรก มีภาษี บันทัดที่สองไม่มีภาษี
                            '.INVAMT = IIf(IsNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7)) ' APOBL.AMTINVCHC       '------------------ Edit by X [09/09/2005]
                            '.INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17)) '/ rs.options.Rs.Fields("CountDup") ' APIBD.AMTGLDIST      '------------------ Edit by X [09/09/2005]  ++++ wut 22/02/2008 add       / rs.options.Rs.Fields("CountDup")
                            '.INVAMT = .INVAMT - IIf(IsNull(rs.options.Rs.Collect(22)), 0, rs.options.Rs.Collect(22)) ' APIBD.AMTTAXREC1
                            '***************by kwan
                            If rs.options.Rs.Collect(23) = 1 Then
                                If rs.options.Rs.Collect(8) < rs.options.Rs.Collect(17) Then
                                    .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8))
                                Else
                                    .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17))
                                End If
                            Else
                                If rs.options.Rs.Collect(23) > 0 Then
                                    If rs.options.Rs.Collect(8) = rs.options.Rs.Collect(17) Then
                                        .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17)) / rs.options.Rs.Collect(23)
                                    Else
                                        If rs.options.Rs.Collect(8) < rs.options.Rs.Collect(17) Then
                                            .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8))
                                        Else
                                            .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17))
                                        End If
                                    End If
                                Else
                                    .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17))
                                End If

                            End If
                        End If

                        .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(9)), 0, rs.options.Rs.Collect(9)) '***************by kwan

                        If rs.options.Rs.Collect(18) = 3 And .INVTAX > 0 Then '----------- If TXTTRXTYPE = 3 (Credit Note)
                            .INVTAX = .INVTAX * (-1)
                        End If
                    End If

                    If .INVTAX < 0 Then
                        If .INVAMT > 0 Then
                            .INVAMT = .INVAMT * (-1)
                        End If
                    End If

                    .IDINVC = rs.options.Rs.Collect("IDINVC")
                    .TXDATE = rs.options.Rs.Collect("DATEBUS")
                    '.INVNAME = rs.options.Rs.Collect("VENDNAME")
                    .INVNAME = ""
                    .INVDATE = rs.options.Rs.Collect("DATEINVC")
                    .DOCNO = rs.options.Rs.Collect("IDPONBR")
                    .INVTAX = rs.options.Rs.Collect("AMTTOTTAX")
                    .TRANSNBR = rs.options.Rs.Collect("CNTLINE")
                    INVTAXtemp = 0
                    'หาก =false เท่ากับให้ดึงค่าจาก Comment
                    If OPT_AP = False Then
                        For i As Integer = 0 To TEXTLINE.Length - 1
                            If IsDBNull(TEXTLINE(i)) = True Then
                                TEXTLINE(i) = ""
                            End If
                        Next
                        If TEXTLINE.Length >= 1 Then
                            If TEXTLINE(0).Trim <> "" And Decimal.TryParse(TEXTLINE(0), 0) Then
                                .INVAMT = CDec(TEXTLINE(0))
                            End If
                        End If

                        If TEXTLINE.Length >= 2 Then
                            If TEXTLINE(1).Trim <> "" Then
                                .IDINVC = TEXTLINE(1)
                            End If
                        End If

                        If TEXTLINE.Length >= 3 Then
                            If TEXTLINE(2).Trim <> "" Then
                                .INVDATE = TryDate2(ZeroDate(TEXTLINE(2), "dd/MM/yyyy", "yyyyMMdd"), "dd/MM/yyyy", "yyyyMMdd")
                                If .INVDATE = 0 Then
                                    .INVDATE = TryDate2(TEXTLINE(2), "dd/MM/yyyy", "yyyyMMdd")
                                End If
                            End If
                        End If

                        If TEXTLINE.Length >= 4 Then
                            If TEXTLINE(3).Trim = "" Then
                                '.INVNAME = rs.options.Rs.Collect("VENDNAME")
                                .INVNAME = rs.options.Rs.Collect("IDVEND")
                            Else
                                .INVNAME = TEXTLINE(3)
                                For CName As Integer = 4 To TEXTLINE.Length - 1
                                    .INVNAME = .INVNAME & "," & TEXTLINE(CName)
                                Next
                            End If
                        End If
                        'หาก =true เท่ากับ ให้ดึงจากจาก Optional field

                    Else 'OPT_AP = TRUE
                        Dim RS_WHT_DETAIL As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from APIBDO  where  OPTFIELD in ('TAXBASE','TAXINVNO','TAXDATE','TAXNAME','TAXID','BRANCH') and CNTBTCH ='" & rs.options.Rs.Collect("CNTBTCH") & "' and CNTITEM='" & rs.options.Rs.Collect("CNTITEM") & "' and CNTLINE ='" & IIf(IsDBNull(rs.options.Rs.Collect("CNTLINE")), "0", rs.options.Rs.Collect("CNTLINE")) & "' ", cnACCPAC)

                        Dim tINVAMT As Double = 0
                        Dim tTAXINVNO As String = ""
                        Dim tTAXDATE As String = ""
                        Dim tTAXNAME As String = ""
                        Dim tTAXID As String = ""
                        Dim tBRANCH As String = ""

                        For iwht As Integer = 0 To RS_WHT_DETAIL.options.QueryDT.Rows.Count - 1
                            'ดึงค่าจาก OPF ชื่อ TAXBASE
                            'ฐานภาษี
                            If RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXBASE" Then
                                Dim Trydec As String = Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE"))
                                If Decimal.TryParse(Trydec, 0) = True Then
                                    If CDec(Trydec) <> 0 Then
                                        tINVAMT = CDec(Trydec)
                                    Else
                                        tINVAMT = 0
                                    End If
                                Else
                                    tINVAMT = 0
                                End If
                                'ดึงค่าจาก OPF ชื่อ TAXINVNO
                                'เลขที่  Invoice
                            ElseIf RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXINVNO" Then
                                If Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    tTAXINVNO = Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If
                                'ดึงค่าจาก OPF ชื่อ TAXDATE
                                'วันที่ Invoice
                            ElseIf RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXDATE" Then
                                Dim myDate As String = Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE"))
                                If Decimal.TryParse(myDate, 0) = True Then
                                    If CDec(myDate) <> 0 Then
                                        tTAXDATE = CDec(myDate)
                                    End If
                                End If
                                'ดึงค่าจาก OPF ชื่อ TAXNAME
                            ElseIf RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXNAME" Then
                                If Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    Dim chkIDVEND_Opt As New BaseClass.BaseODBCIO("select VENDORID,APVNR.RMITNAME AS RMITNAME from APVEN inner join APVNR on APVEN.VENDORID = APVNR.IDVEND where IDVENDRMIT = 'TXTHAI' and VENDORID = '" & Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE")) & "'", cnACCPAC)
                                    If chkIDVEND_Opt.options.QueryDT.Rows.Count = 0 Then
                                        tTAXNAME = Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE"))
                                    Else
                                        tTAXNAME = chkIDVEND_Opt.options.QueryDT.Rows(0).Item("RMITNAME").ToString
                                    End If
                                End If


                            ElseIf RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXID" Then
                                If Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    tTAXID = Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If

                            ElseIf RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "BRANCH" Then
                                If Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    tBRANCH = Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If
                            End If
                        Next
                        If tINVAMT <> 0 Then
                            .INVAMT = tINVAMT
                        Else
                            .INVAMT = 0
                        End If
                        If tTAXINVNO <> "" Then
                            .IDINVC = tTAXINVNO
                        End If
                        If tTAXDATE <> "" Then
                            .INVDATE = tTAXDATE
                        End If
                        If tTAXNAME <> "" Then
                            .INVNAME = tTAXNAME
                        End If
                        If tTAXID <> "" Then
                            .TAXID = tTAXID
                        End If
                        If tBRANCH <> "" Then
                            .Branch = tBRANCH
                        End If

                    End If


                End If 'end if "DETAIL" "HEADER" 

                If .INVNAME.ToString.Trim.ToUpper = "ZZZ" Then .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect("VDADDRESS1")), "", rs.options.Rs.Collect("VDADDRESS1"))

                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect(10)), "", rs.options.Rs.Collect(10)) 'APPJD.IDAPACCT
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(11)), "", (rs.options.Rs.Collect(11))) 'APOBL.CNTBTCH
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(12)), "", (rs.options.Rs.Collect(12))) 'APOBL.CNTITEM
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(16)), "", (rs.options.Rs.Collect(16))) 'APOBL.IDPONBR
                .TypeOfPU = IIf(IsDBNull(rs.options.Rs.Collect("NAMECITY")), "", (rs.options.Rs.Collect("NAMECITY")))
                .Runno = IIf(IsDBNull(rs.options.Rs.Collect("value")), "", rs.options.Rs.Collect("value"))

                If .INVAMT = 0 Then
                    .RATE_Renamed = 0
                Else
                    .RATE_Renamed = FormatNumber(((.INVTAX / .INVAMT) * 100), 1)
                End If

                If FindMiscCodeInAPVEN(Trim(.INVNAME), "AP") = "0" Or FindMiscCodeInAPVEN(Trim(.INVNAME), "AP") = "" Then
                    If Trim(.INVNAME) = "" Then
                        .INVNAME = FindMiscCodeInAPVEN(Trim(rs.options.Rs.Collect("IDVEND")), "AP")
                    Else
                        .INVNAME = Trim(.INVNAME)
                    End If
                    If .TAXID = "" And .Branch = "" Then
                        .TAXID = FindTaxIDBRANCH(Trim(rs.options.Rs.Collect("IDVEND")), "AP", "TAXID")
                        .Branch = FindTaxIDBRANCH(Trim(rs.options.Rs.Collect("IDVEND")), "AP", "BRANCH")
                    End If


                Else
                    If .TAXID = "" And .Branch = "" Then

                        .TAXID = FindTaxIDBRANCH(.INVNAME, "AP", "TAXID")
                        .Branch = FindTaxIDBRANCH(.INVNAME, "AP", "BRANCH")
                        .INVNAME = FindMiscCodeInAPVEN(Trim(.INVNAME), "AP")
                    End If

                    End If
                    .VATCOMMENT = .INVAMT & "," & .IDINVC & "," & .INVDATE & "," & .INVNAME

                    sqlstr = InsVat(0, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, "", (.INVNAME), .INVAMT, .INVTAX, .LOCID, 1, .RATE_Renamed, .TTYPE, .ACCTVAT, "PO", .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, .Runno, IIf(ISNULL(rs.options.Rs, 13), "", Trim(rs.options.Rs.Collect(13))), IIf(ISNULL(rs.options.Rs, 13), "", Trim(rs.options.Rs.Collect(13))), .TypeOfPU, "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, INVTAXtemp)
            End With
            cnVAT.Execute(sqlstr)
            ErrMsg = ""
            Code = ""
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing

        'transfer data
        rs = New BaseClass.BaseODBCIO
        If ComDatabase = "ORCL" Then
            sqlstr = "SELECT FMSVATTEMP.INVDATE,FMSVATTEMP.TXDATE,FMSVATTEMP.IDINVC,FMSVATTEMP.DOCNO," '0-3
            sqlstr = sqlstr & " FMSVATTEMP.INVNAME,FMSVATTEMP.INVAMT,FMSVATTEMP.INVTAX," '4-6
            sqlstr = sqlstr & " FMSVLACC.LOCID,FMSVATTEMP.VTYPE,FMSVATTEMP.RATE,FMSTAX.TTYPE," '7-10
            sqlstr = sqlstr & " FMSTAX.ACCTVAT,FMSVATTEMP.SOURCE,FMSVATTEMP.BATCH,FMSVATTEMP.ENTRY," '11-14
            sqlstr = sqlstr & " FMSVATTEMP.MARK,FMSVATTEMP.VATCOMMENT,FMSTAX.ITEMRATE1,FMSVATTEMP.IDDIST,FMSVATTEMP.CBREF,FMSVATTEMP.RUNNO,FMSVATTEMP.TypeofPu,FMSVATTEMP.tranNo,FMSVATTEMP.Cif,FMSVATTEMP.taxCif,FMSVATTEMP.TRANSNBR, " '15-20
            sqlstr = sqlstr & " FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH,FMSVATTEMP.Code,FMSVATTEMP.TOTTAX ,FMSVATTEMP.CODETAX "
            sqlstr = sqlstr & " FROM FMSVATTEMP,FMSTAX,FMSVLACC WHERE "
            sqlstr = sqlstr & " (FMSVATTEMP.CODETAX = FMSTAX.AUTHORITY AND  FMSTAX.ACCTVAT= FMSVATTEMP.ACCTVAT)"
            sqlstr = sqlstr & " AND FMSTAX.ACCTVAT = FMSVLACC.ACCTVAT"
            sqlstr = sqlstr & " AND (FMSTAX.TTYPE = 1 or FMSTAX.TTYPE = 3) AND (FMSTAX.BUYERCLASS = 1)"
        Else
            sqlstr = " SELECT   FMSVATTEMP.INVDATE, FMSVATTEMP.TXDATE, FMSVATTEMP.IDINVC, FMSVATTEMP.DOCNO," '0-3
            sqlstr = sqlstr & " FMSVATTEMP.INVNAME, sum(FMSVATTEMP.INVAMT) AS INVAMT, sum(FMSVATTEMP.INVTAX) AS INVTAX," '4-6
            sqlstr = sqlstr & " FMSVLACC.LOCID, FMSVATTEMP.VTYPE, FMSVATTEMP.RATE, FMSTAX.TTYPE," '7-10
            sqlstr = sqlstr & " FMSTAX.ACCTVAT,FMSVATTEMP.SOURCE, FMSVATTEMP.BATCH, FMSVATTEMP.ENTRY," '11-14
            sqlstr = sqlstr & " FMSVATTEMP.MARK, FMSVATTEMP.VATCOMMENT, FMSTAX.ITEMRATE1, FMSVATTEMP.IDDIST , FMSVATTEMP.CBREF,FMSVATTEMP.RUNNO,FMSVATTEMP.RUNNO,FMSVATTEMP.TypeofPu,FMSVATTEMP.tranNo,FMSVATTEMP.Cif,FMSVATTEMP.taxCif,FMSVATTEMP.TRANSNBR, " '15-20
            sqlstr = sqlstr & " FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH,FMSVATTEMP.Code,FMSVATTEMP.TOTTAX,FMSVATTEMP.CODETAX "
            sqlstr = sqlstr & " FROM FMSVATTEMP "
            sqlstr = sqlstr & "     INNER JOIN FMSTAX ON (FMSVATTEMP.CODETAX = FMSTAX.AUTHORITY) and  (FMSTAX.ACCTVAT= FMSVATTEMP.ACCTVAT)"
            sqlstr = sqlstr & "     INNER JOIN FMSVLACC ON FMSTAX.ACCTVAT = FMSVLACC.ACCTVAT "
            sqlstr = sqlstr & " Where (FMSTAX.BUYERCLASS = 1) "
            sqlstr = sqlstr & " GROUP BY FMSVATTEMP.INVDATE, FMSVATTEMP.TXDATE, FMSVATTEMP.IDINVC, FMSVATTEMP.DOCNO, FMSVATTEMP.INVNAME, FMSVLACC.LOCID,"
            sqlstr = sqlstr & "     FMSVATTEMP.VTYPE, FMSVATTEMP.RATE, FMSTAX.TTYPE, FMSTAX.ACCTVAT, FMSVATTEMP.SOURCE, FMSVATTEMP.BATCH, FMSVATTEMP.ENTRY,"
            sqlstr = sqlstr & "     FMSVATTEMP.MARK , FMSVATTEMP.VATCOMMENT, FMSTAX.ITEMRATE1, FMSVATTEMP.IDDIST, FMSVATTEMP.CBRef,FMSVATTEMP.RUNNO,FMSVATTEMP.RUNNO,FMSVATTEMP.TypeofPu,FMSVATTEMP.tranNo,FMSVATTEMP.Cif,FMSVATTEMP.taxCif,FMSVATTEMP.TRANSNBR, "
            sqlstr = sqlstr & "     FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH,FMSVATTEMP.Code,FMSVATTEMP.TOTTAX,FMSVATTEMP.CODETAX "
            sqlstr = sqlstr & " HAVING (FMSTAX.TTYPE = 1 OR FMSTAX.TTYPE = 3)"
        End If

        rs.Open(sqlstr, cnVAT)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessPO> Select FMSVATTEMP Complete")
        End If

        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls
                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect(0)), 0, rs.options.Rs.Collect(0))
                .TXDATE = IIf(IsDBNull(rs.options.Rs.Collect(1)), 0, rs.options.Rs.Collect(1))
                .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", (rs.options.Rs.Collect(2)))
                .DOCNO = IIf(IsDBNull(rs.options.Rs.Collect(3)), "", (rs.options.Rs.Collect(3)))
                .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect(4)), "", (rs.options.Rs.Collect(4)))
                .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(5)), 0, rs.options.Rs.Collect(5))
                .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(6)), 0, rs.options.Rs.Collect(6))
                .LOCID = IIf(IsDBNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7))
                .VTYPE = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8))
                .TTYPE = IIf(IsDBNull(rs.options.Rs.Collect(10)), 0, rs.options.Rs.Collect(10))
                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect(11)), "", (rs.options.Rs.Collect(11)))
                .Source = IIf(IsDBNull(rs.options.Rs.Collect(12)), "", (rs.options.Rs.Collect(12)))
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(13)), "", (rs.options.Rs.Collect(13)))
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(14)), "", (rs.options.Rs.Collect(14)))
                .MARK = IIf(IsDBNull(rs.options.Rs.Collect(15)), "", (rs.options.Rs.Collect(15)))
                .VATCOMMENT = IIf(IsDBNull(rs.options.Rs.Collect(16)), "", (rs.options.Rs.Collect(16)))
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(19)), "", (rs.options.Rs.Collect(19)))
                .TRANSNBR = IIf(IsDBNull(rs.options.Rs.Collect("TRANSNBR")), "", (rs.options.Rs.Collect("TRANSNBR")))
                .TypeOfPU = IIf(IsDBNull(rs.options.Rs.Collect("TypeOfPU")), "", rs.options.Rs.Collect("TypeOfPU"))
                .TranNo = IIf(IsDBNull(rs.options.Rs.Collect("TranNo")), "", rs.options.Rs.Collect("TranNo"))
                .CIF = IIf(IsDBNull(rs.options.Rs.Collect("CIF")), "0", rs.options.Rs.Collect("CIF"))
                .TaxCIF = IIf(IsDBNull(rs.options.Rs.Collect("TaxCIF")), "0", rs.options.Rs.Collect("TaxCIF"))
                .TAXID = IIf(IsDBNull(rs.options.Rs.Collect("TAXID")), "0", rs.options.Rs.Collect("TAXID"))
                .Branch = IIf(IsDBNull(rs.options.Rs.Collect("BRANCH")), "0", rs.options.Rs.Collect("BRANCH"))
                .Runno = IIf(IsDBNull(rs.options.Rs.Collect("runno")), "", rs.options.Rs.Collect("runno"))
                Code = IIf(IsDBNull(rs.options.Rs.Collect("Code")), "", (rs.options.Rs.Collect("Code")))
                .TOTTAX = IIf(IsDBNull(rs.options.Rs.Collect("TOTTAX")), 0, (rs.options.Rs.Collect("TOTTAX")))
                .RATE_Renamed = IIf(IsDBNull(rs.options.Rs.Collect("RATE")), 0, (rs.options.Rs.Collect("RATE")))
                .CODETAX = IIf(IsDBNull(rs.options.Rs.Collect("CODETAX")), "", rs.options.Rs.Collect("CODETAX"))
            
                sqlstr = InsVat(1, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, .NEWDOCNO, (.INVNAME), .INVAMT, .INVTAX, .LOCID, .VTYPE, .RATE_Renamed, .TTYPE, .ACCTVAT, .Source, .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, .Runno, .CODETAX, "", .TypeOfPU, .TranNo, .CIF, .TaxCIF, 0, .TAXID, .Branch, Code, "", 0, 0, .TOTTAX)

            End With
            cnVAT.Execute(sqlstr)
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing
        cnACCPAC.Close()
        cnVAT.Close()
        cnACCPAC = Nothing
        cnVAT = Nothing
        Exit Sub
ErrHandler:
        WriteLog("<ProcessPO>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessPO> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
        WriteLog(ErrMsg)

    End Sub
    Public Sub ProcessAPPJHTEMP()
        Dim cn As ADODB.Connection
        Dim sqlstr As String

        On Error GoTo ErrHandler

        cn = New ADODB.Connection

        cn.ConnectionTimeout = 60
        cn.Open(ConnACCPAC)
        cn.CommandTimeout = 3600

        sqlstr = "INSERT INTO " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "APPJHTEMP"
        sqlstr = sqlstr & " (TYPEBTCH,"
        sqlstr = sqlstr & " POSTSEQNCE,"
        sqlstr = sqlstr & " CNTBTCH,"
        sqlstr = sqlstr & " CNTITEM,"
        sqlstr = sqlstr & " IDVEND,"
        sqlstr = sqlstr & " IDINVC,"
        sqlstr = sqlstr & " DATEBUS)"
        sqlstr = sqlstr & "SELECT"
        sqlstr = sqlstr & " TYPEBTCH,"
        sqlstr = sqlstr & " POSTSEQNCE,"
        sqlstr = sqlstr & " CNTBTCH,"
        sqlstr = sqlstr & " CNTITEM,"
        sqlstr = sqlstr & " IDVEND,"
        sqlstr = sqlstr & " IDINVC,"
        sqlstr = sqlstr & " DATEBUS"
        sqlstr = sqlstr & " FROM APPJH"
        sqlstr = sqlstr & " WHERE (APPJH.DATEBUS >= " & DATEFROM & " AND APPJH.DATEBUS <=" & DATETO & ")"
        cn.Execute(sqlstr)
        cn.Close()
        cn = Nothing
        Exit Sub

ErrHandler:
        cn.RollbackTrans()
        ErrMsg = Err.Description
        WriteLog("<ProcessAPPJHTEMP>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    Public Sub ProcessAPOBLTEMP()
        Dim cn As ADODB.Connection
        Dim sqlstr As String

        On Error GoTo ErrHandler

        cn = New ADODB.Connection

        cn.ConnectionTimeout = 60
        cn.Open(ConnACCPAC)
        cn.CommandTimeout = 3600

        sqlstr = "INSERT INTO " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "APOBLTEMP"
        sqlstr = sqlstr & " (IDVEND,"
        sqlstr = sqlstr & " IDINVC,"
        sqlstr = sqlstr & " AUDTDATE,"
        sqlstr = sqlstr & " IDPONBR,"
        sqlstr = sqlstr & " TXTTRXTYPE,"
        sqlstr = sqlstr & " CNTBTCH,"
        sqlstr = sqlstr & " CNTITEM,"
        sqlstr = sqlstr & " DATEINVC,"
        sqlstr = sqlstr & " AMTINVCHC,"
        sqlstr = sqlstr & " AMTTAXHC,"
        sqlstr = sqlstr & " CODETAX1,"
        sqlstr = sqlstr & " AMTBASE1HC,"
        sqlstr = sqlstr & " AMTTAX1HC,"
        sqlstr = sqlstr & " DATEBUS,"
        sqlstr = sqlstr & " CODETAXGRP,"

        If ComDatabase = "MSSQL" Then
            sqlstr = sqlstr & " [VALUES])"
        Else 'PVSW
            sqlstr = sqlstr & " " & Chr(34) & "VALUES" & Chr(34) & ")"
        End If

        sqlstr = sqlstr & " SELECT"
        sqlstr = sqlstr & " IDVEND,"
        sqlstr = sqlstr & " IDINVC,"
        sqlstr = sqlstr & " AUDTDATE,"
        sqlstr = sqlstr & " IDPONBR,"
        sqlstr = sqlstr & " TXTTRXTYPE,"
        sqlstr = sqlstr & " CNTBTCH,"
        sqlstr = sqlstr & " CNTITEM,"
        sqlstr = sqlstr & " DATEINVC,"
        sqlstr = sqlstr & " AMTINVCHC,"
        sqlstr = sqlstr & " AMTTAXHC,"
        sqlstr = sqlstr & " CODETAX1,"
        sqlstr = sqlstr & " AMTBASE1HC,"
        sqlstr = sqlstr & " AMTTAX1HC,"
        sqlstr = sqlstr & " DATEBUS,"
        sqlstr = sqlstr & " CODETAXGRP,"

        If ComDatabase = "MSSQL" Then
            sqlstr = sqlstr & " [VALUES]"
        Else 'PVSW
            sqlstr = sqlstr & " " & Chr(34) & "VALUES" & Chr(34) & ""
        End If

        sqlstr = sqlstr & " FROM APOBL"
        sqlstr = sqlstr & " WHERE APOBL.TXTTRXTYPE IN (1,2,3) AND (APOBL.DATEBUS >= " & DATEFROM & " AND APOBL.DATEBUS <= " & DATETO & ")"
        cn.Execute(sqlstr)
        cn.Close()
        cn = Nothing
        Exit Sub

ErrHandler:
        cn.RollbackTrans()
        ErrMsg = Err.Description
        WriteLog("<ProcessAPOBLTEMP>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub


    Public Sub ProcessAP_PO()
        Dim cnACCPAC, cnVAT As ADODB.Connection
        Dim rs, rs2 As BaseClass.BaseODBCIO
        Dim sqlstr, sqlstr2, sqlstr3 As String
        Dim loCls As clsVat
        Dim iffs As clsFuntion
        Dim tmpTAXDATE As String
        Dim tmpTAXINVC As String
        Dim Ddbo As String
        iffs = New clsFuntion
        ErrMsg = ""

        On Error GoTo ErrHandler
        cnACCPAC = New ADODB.Connection
        cnVAT = New ADODB.Connection

        cnACCPAC.ConnectionTimeout = 60
        cnACCPAC.Open(ConnACCPAC)
        cnACCPAC.CommandTimeout = 3600

        cnVAT.ConnectionTimeout = 60
        cnVAT.Open(ConnVAT)
        cnVAT.CommandTimeout = 3600
        sqlstr = ""
        sqlstr2 = ""
        sqlstr3 = ""

        Ddbo = ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.")

        'gather data ==> Select Data Uses For Module
        rs = New BaseClass.BaseODBCIO

        If Use_POST Then

            If ComDatabase = "MSSQL" Then
                sqlstr = " select * from (  " & Environment.NewLine
                sqlstr &= " SELECT DISTINCT  " & Environment.NewLine
                sqlstr &= " " & Ddbo & "APOBLTEMP.DATEINVC, " & Ddbo & "APOBLTEMP.[VALUES], " & Ddbo & "APOBLTEMP.IDINVC, " & Ddbo & "APOBLTEMP.AUDTDATE, "
                If Use_LEGALNAME_inAP = True Then
                    sqlstr &= "CASE WHEN APVNR.RMITNAME IS NULL THEN APVEN.VENDNAME ELSE APVNR.RMITNAME  END AS VENDNAME, APVEN.TEXTSTRE1, '' AS IDDIST, " & Ddbo & "APOBLTEMP.AMTBASE1HC as AMTINVCHC ,  " & Environment.NewLine
                Else
                    sqlstr &= "APVEN.VENDNAME, APVEN.TEXTSTRE1, '' AS IDDIST, " & Ddbo & "APOBLTEMP.AMTBASE1HC as AMTINVCHC ,  " & Environment.NewLine
                End If

            Else
                sqlstr &= " SELECT DISTINCT  " & Environment.NewLine
                sqlstr &= " " & Ddbo & "APOBLTEMP.DATEINVC," & Ddbo & "APOBLTEMP." & Chr(34) & "VALUES" & Chr(34) & ", " & Ddbo & "APOBLTEMP.IDINVC, " & Ddbo & "APOBLTEMP.AUDTDATE, APVEN.VENDNAME, APVEN.TEXTSTRE1, '' AS IDDIST, " & Ddbo & "APOBLTEMP.AMTBASE1HC as AMTINVCHC ,  " & Environment.NewLine
            End If

            sqlstr &= "            MIN(" & Ddbo & "APOBLTEMP.AMTBASE1HC) AS AMTEXTNDHC, " & Ddbo & "APOBLTEMP.AMTTAXHC,  " & Environment.NewLine
            sqlstr &= "            ISNULL(TXAUTH.ACCTRECOV, APPJD.IDACCT) AS IDACCT, " & Ddbo & "APOBLTEMP.CNTBTCH, " & Ddbo & "APOBLTEMP.CNTITEM, " & Ddbo & "APOBLTEMP.CODETAX1, MIN(PORCPH1.RCPNUMBER) AS Expr1,  " & Environment.NewLine
            sqlstr &= "            POCRNH1.RETNUMBER, " & Ddbo & "APOBLTEMP.IDPONBR,MIN(" & Ddbo & "APOBLTEMP.AMTBASE1HC) AS AMTGLDIST, " & Ddbo & "APOBLTEMP.TXTTRXTYPE, 0 AS TAXBASE, 0 AS TAXDATE,  " & Environment.NewLine
            sqlstr &= "            0 AS TAXNAME, MIN(APIBH.AMTRECTAX) AS AMTTAXREC1, 1 AS COUNTDUP, APIBHO.VALUE, APVEN.NAMECITY,  " & Environment.NewLine
            sqlstr &= "            '' AS TEXTLINE,  MIN(APIBH.AMTRECTAX)AS AMTTOTTAX, '0' AS CNTLINE, " & Ddbo & "APOBLTEMP.DATEBUS,'HEADER' AS VATAS,APVEN.VENDORID AS IDVEND,TXAUTH.RATERECOV AS RATEAVG " & Environment.NewLine
            sqlstr &= " FROM	" & Ddbo & "APOBLTEMP" & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APVEN ON " & Ddbo & "APOBLTEMP.IDVEND = APVEN.VENDORID " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN PORCPH1 ON " & Ddbo & "APOBLTEMP.IDINVC = PORCPH1.INVNUMBER AND " & Ddbo & "APOBLTEMP.IDVEND = PORCPH1.VDCODE " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN POCRNH1 ON " & Ddbo & "APOBLTEMP.IDINVC = POCRNH1.CRNNUMBER " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN " & Ddbo & "APPJHTEMP ON " & Ddbo & "APOBLTEMP.IDINVC = " & Ddbo & "APPJHTEMP.IDINVC AND " & Ddbo & "APOBLTEMP.IDVEND = " & Ddbo & "APPJHTEMP.IDVEND " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APPJD ON " & Ddbo & "APPJHTEMP.POSTSEQNCE = APPJD.POSTSEQNCE AND " & Ddbo & "APPJHTEMP.CNTBTCH = APPJD.CNTBTCH AND " & Ddbo & "APPJHTEMP.CNTITEM = APPJD.CNTITEM " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APIBH ON " & Ddbo & "APOBLTEMP.IDINVC = APIBH.IDINVC AND " & Ddbo & "APOBLTEMP.IDVEND = APIBH.IDVEND " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APIBHO ON APIBH.CNTBTCH = APIBHO.CNTBTCH AND APIBH.CNTITEM = APIBHO.CNTITEM AND APIBHO.OPTFIELD = 'VOUCHER' " & Environment.NewLine
            'sqlstr &= "             LEFT OUTER JOIN TXAUTH ON APIBH.CODETAXGRP = TXAUTH.AUTHORITY  " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN TXAUTH ON " & Ddbo & "APOBLTEMP.CODETAXGRP = TXAUTH.AUTHORITY  " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APVNR ON APVNR.IDVEND = APVEN.VENDORID  "
            'APIBH.ERRENTRY IS NULL  เพิ่มเงื่อนไขกรณี Clear History table APIBH vat ปกติข้อมูลจะยังอยู่แต่ถ้าเป็น vat เฉลี่ยข้อมูลจะหาย
            sqlstr &= " WHERE       (" & Ddbo & "APOBLTEMP.TXTTRXTYPE IN (1,2,3)) AND (" & Ddbo & "APOBLTEMP." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " >= " & DATEFROM & " AND " & Ddbo & "APOBLTEMP." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " <= " & DATETO & " )  AND (APIBH.ERRENTRY IS NULL OR APIBH.ERRENTRY = 0)  " & Environment.NewLine
            sqlstr &= "             AND (TXAUTH.ACCTRECOV IN " & Environment.NewLine
            sqlstr &= "             (SELECT ACCTVAT " & IIf(ComDatabase = "PVSW", "", "COLLATE Thai_CI_AS") & Environment.NewLine
            sqlstr &= "             FROM " & Ddbo & "FMSVLACC))   " & Environment.NewLine
            'sqlstr &= " AND " & Ddbo & "APOBLTEMP.AMTTAXHC<>0 " & Environment.NewLine

            If ComDatabase = "MSSQL" Then
                sqlstr &= " GROUP BY    " & Ddbo & "APOBLTEMP.DATEINVC, " & Ddbo & "APOBLTEMP.[VALUES], " & Ddbo & "APOBLTEMP.IDINVC, " & Ddbo & "APOBLTEMP.AUDTDATE, APVEN.VENDNAME, APVEN.TEXTSTRE1, " & Ddbo & "APOBLTEMP.AMTBASE1HC, " & Ddbo & "APOBLTEMP.AMTTAXHC,  " & Environment.NewLine
            Else 'PVSW
                sqlstr &= " GROUP BY    " & Ddbo & "APOBLTEMP.DATEINVC, " & Ddbo & "APOBLTEMP." & Chr(34) & "VALUES" & Chr(34) & ", " & Ddbo & "APOBLTEMP.IDINVC, " & Ddbo & "APOBLTEMP.AUDTDATE, APVEN.VENDNAME, APVEN.TEXTSTRE1, " & Ddbo & "APOBLTEMP.AMTBASE1HC, " & Ddbo & "APOBLTEMP.AMTTAXHC,  " & Environment.NewLine
            End If

            sqlstr &= "             TXAUTH.ACCTRECOV, " & Ddbo & "APOBLTEMP.CNTBTCH, " & Ddbo & "APOBLTEMP.CNTITEM, " & Ddbo & "APOBLTEMP.CODETAX1, POCRNH1.RETNUMBER, " & Ddbo & "APOBLTEMP.IDPONBR, " & Ddbo & "APOBLTEMP.TXTTRXTYPE, APIBHO.VALUE,  " & Environment.NewLine
            sqlstr &= "             APVEN.NAMECITY, APPJD.IDACCT, " & Ddbo & "APPJHTEMP.POSTSEQNCE, " & Ddbo & "APOBLTEMP.DATEBUS,APVEN.VENDORID,TXAUTH.RATERECOV " & Environment.NewLine

            sqlstr &= " UNION ALL " & Environment.NewLine

            sqlstr &= " SELECT DISTINCT  " & Environment.NewLine

            If ComDatabase = "MSSQL" Then
                sqlstr &= "         " & Ddbo & "APOBLTEMP.DATEINVC, " & Ddbo & "APOBLTEMP.[VALUES], " & Ddbo & "APOBLTEMP.IDINVC, " & Ddbo & "APOBLTEMP.AUDTDATE, "
                If Use_LEGALNAME_inAP = True Then
                    sqlstr &= "CASE WHEN APVNR.RMITNAME IS NULL THEN APVEN.VENDNAME ELSE APVNR.RMITNNAME END AS VENDNAME, APVEN.TEXTSTRE1, '' AS IDDIST, " & Ddbo & "APOBLTEMP.AMTINVCHC, " & Environment.NewLine
                Else
                    sqlstr &= "APVEN.VENDNAME, APVEN.TEXTSTRE1, '' AS IDDIST, " & Ddbo & "APOBLTEMP.AMTINVCHC, " & Environment.NewLine
                End If

            Else
                sqlstr &= "         " & Ddbo & "APOBLTEMP.DATEINVC, " & Ddbo & "APOBLTEMP." & Chr(34) & "VALUES" & Chr(34) & ", " & Ddbo & "APOBLTEMP.IDINVC, " & Ddbo & "APOBLTEMP.AUDTDATE, APVEN.VENDNAME, APVEN.TEXTSTRE1, '' AS IDDIST, " & Ddbo & "APOBLTEMP.AMTINVCHC, " & Environment.NewLine
            End If

            sqlstr &= "             (CASE WHEN APIBT.CNTLINE IS NULL THEN MIN(" & Ddbo & "APOBLTEMP.AMTBASE1HC) ELSE SUM(" & Ddbo & "APOBLTEMP.AMTBASE1HC) END) AS AMTEXTNDHC, SUM(APIBD.AMTTAXREC1) AS AMTTAXHC, " & Environment.NewLine
            sqlstr &= "             ISNULL(TXAUTH.ACCTRECOV, APPJD.IDACCT) AS IDACCT, " & Ddbo & "APOBLTEMP.CNTBTCH, " & Ddbo & "APOBLTEMP.CNTITEM, " & Ddbo & "APOBLTEMP.CODETAX1, MIN(PORCPH1.RCPNUMBER) AS Expr1, " & Environment.NewLine
            sqlstr &= "             POCRNH1.RETNUMBER, " & Ddbo & "APOBLTEMP.IDPONBR, MIN(APIBD.AMTGLDIST) AS AMTGLDIST, " & Ddbo & "APOBLTEMP.TXTTRXTYPE, " & Environment.NewLine
            sqlstr &= "             0 AS TAXBASE, 0 AS TAXDATE, ' ' AS TAXNAME, MIN(APIBD.AMTTAXREC1) AS AMTTAXREC1, " & Environment.NewLine
            sqlstr &= "             1 AS COUNTDUP, APIBHO.VALUE, APVEN.NAMECITY, APIBT.TEXTLINE, (CASE WHEN APIBT.CNTLINE IS NULL THEN SUM(APIBD.AMTTOTTAX) ELSE " & Environment.NewLine
            sqlstr &= "             (SELECT TOP 1 APPJD.AMTEXTNDHC " & Environment.NewLine
            sqlstr &= "             FROM APPJD " & Environment.NewLine
            sqlstr &= "             WHERE POSTSEQNCE = " & Ddbo & "APPJHTEMP.POSTSEQNCE AND CNTBTCH = " & Ddbo & "APOBLTEMP.CNTBTCH AND CNTITEM = " & Ddbo & "APOBLTEMP.CNTITEM AND CNTLINE = APIBT.CNTLINE) END) AS AMTTOTTAX, " & Environment.NewLine
            sqlstr &= "             APPJD.CNTLINE, " & Ddbo & "APOBLTEMP.DATEBUS ,'DETAIL' as VATAS," & Ddbo & "APOBLTEMP.IDVEND,TXAUTH.RATERECOV AS RATEAVG " & Environment.NewLine
            sqlstr &= " FROM	" & Ddbo & "APOBLTEMP " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APVEN ON " & Ddbo & "APOBLTEMP.IDVEND = APVEN.VENDORID " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN PORCPH1 ON " & Ddbo & "APOBLTEMP.IDINVC = PORCPH1.INVNUMBER AND " & Ddbo & "APOBLTEMP.IDVEND = PORCPH1.VDCODE " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN POCRNH1 ON " & Ddbo & "APOBLTEMP.IDINVC = POCRNH1.CRNNUMBER " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN " & Ddbo & "APPJHTEMP ON " & Ddbo & "APOBLTEMP.IDINVC = " & Ddbo & "APPJHTEMP.IDINVC AND " & Ddbo & "APOBLTEMP.IDVEND = " & Ddbo & "APPJHTEMP.IDVEND " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APPJD ON " & Ddbo & "APPJHTEMP.POSTSEQNCE = APPJD.POSTSEQNCE AND " & Ddbo & "APPJHTEMP.CNTBTCH = APPJD.CNTBTCH AND " & Ddbo & "APPJHTEMP.CNTITEM = APPJD.CNTITEM " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APIBH ON " & Ddbo & "APOBLTEMP.IDINVC = APIBH.IDINVC AND " & Ddbo & "APOBLTEMP.IDVEND = APIBH.IDVEND " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APIBD ON APIBH.CNTBTCH = APIBD.CNTBTCH AND APIBH.CNTITEM = APIBD.CNTITEM  and APIBD.CNTLINE =APPJD.CNTLINE " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APIBHO ON APIBH.CNTBTCH = APIBHO.CNTBTCH AND APIBH.CNTITEM = APIBHO.CNTITEM AND APIBHO.OPTFIELD = 'VOUCHER' " & Environment.NewLine
            'sqlstr &= "             LEFT OUTER JOIN TXAUTH ON APIBH.CODETAXGRP = TXAUTH.AUTHORITY " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN TXAUTH ON " & Ddbo & "APOBLTEMP.CODETAXGRP = TXAUTH.AUTHORITY " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APIBT ON APIBH.CNTBTCH = APIBT.CNTBTCH AND APIBH.CNTITEM = APIBT.CNTITEM AND APIBT.CNTLINE = APIBD.CNTLINE " & Environment.NewLine
            sqlstr &= "             LEFT OUTER JOIN APVNR ON APVNR.IDVEND = APVEN.VENDORID " & Environment.NewLine
            'APIBH.ERRENTRY IS NULL และ APIBD.TAXCLASS1 IS NULL  เพิ่มเงื่อนไขกรณี Clear History table APIBH vat ปกติข้อมูลจะยังอยู่แต่ถ้าเป็น vat เฉลี่ยข้อมูลจะหาย
            sqlstr &= " WHERE       (" & Ddbo & "APOBLTEMP.TXTTRXTYPE IN (1,2,3)) AND (" & Ddbo & "APOBLTEMP." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " >= " & DATEFROM & " AND " & Ddbo & "APOBLTEMP." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " <= " & DATETO & " ) AND (APIBD.TAXCLASS1 IS NULL OR APIBD.TAXCLASS1 = 1) AND (APIBH.ERRENTRY IS NULL OR APIBH.ERRENTRY = 0)  " & Environment.NewLine
            sqlstr &= "             AND " & Ddbo & "APOBLTEMP.AMTTAXHC = 0 AND (APPJD.IDACCT IN " & Environment.NewLine
            sqlstr &= "             (SELECT ACCTVAT " & IIf(ComDatabase = "PVSW", "", "COLLATE Thai_CI_AS") & Environment.NewLine
            sqlstr &= "             FROM " & Ddbo & "FMSVLACC))   " & Environment.NewLine

            If ComDatabase = "MSSQL" Then
                sqlstr &= " GROUP BY    " & Ddbo & "APOBLTEMP.DATEINVC, " & Ddbo & "APOBLTEMP.[VALUES], " & Ddbo & "APOBLTEMP.IDINVC, " & Ddbo & "APOBLTEMP.AUDTDATE, APVEN.VENDNAME, APVEN.TEXTSTRE1, " & Ddbo & "APOBLTEMP.AMTINVCHC, " & Ddbo & "APOBLTEMP.AMTTAXHC,  " & Environment.NewLine
                sqlstr &= "             TXAUTH.ACCTRECOV, " & Ddbo & "APOBLTEMP.CNTBTCH, " & Ddbo & "APOBLTEMP.CNTITEM, " & Ddbo & "APOBLTEMP.CODETAX1, POCRNH1.RETNUMBER, " & Ddbo & "APOBLTEMP.IDPONBR, " & Ddbo & "APOBLTEMP.TXTTRXTYPE, APIBHO.VALUE,  " & Environment.NewLine
                sqlstr &= "             APVEN.NAMECITY, APPJD.IDACCT, APIBT.TEXTLINE, APIBT.CNTLINE," & IIf(OPT_AP, "APIBD.CNTLINE,", "") & " " & Ddbo & "APPJHTEMP.POSTSEQNCE, " & Ddbo & "APOBLTEMP.DATEBUS, APPJD.CNTLINE ," & Ddbo & "APOBLTEMP.IDVEND,TXAUTH.RATERECOV " & Environment.NewLine
                sqlstr &= " )AS TMP1 ORDER BY DATEINVC,CNTBTCH ,CNTITEM ,CNTLINE " & Environment.NewLine

            Else 'PVSW
                sqlstr &= " GROUP BY    " & Ddbo & "APOBLTEMP.DATEINVC, " & Ddbo & "APOBLTEMP." & Chr(34) & "VALUES" & Chr(34) & ", " & Ddbo & "APOBLTEMP.IDINVC, " & Ddbo & "APOBLTEMP.AUDTDATE, APVEN.VENDNAME, APVEN.TEXTSTRE1, " & Ddbo & "APOBLTEMP.AMTINVCHC, " & Ddbo & "APOBLTEMP.AMTTAXHC,  " & Environment.NewLine
                sqlstr &= "             TXAUTH.ACCTRECOV, " & Ddbo & "APOBLTEMP.CNTBTCH, " & Ddbo & "APOBLTEMP.CNTITEM, " & Ddbo & "APOBLTEMP.CODETAX1, POCRNH1.RETNUMBER, " & Ddbo & "APOBLTEMP.IDPONBR, " & Ddbo & "APOBLTEMP.TXTTRXTYPE, APIBHO.VALUE,  " & Environment.NewLine
                sqlstr &= "             APVEN.NAMECITY, APPJD.IDACCT, APIBT.TEXTLINE, APIBT.CNTLINE," & IIf(OPT_AP, "APIBD.CNTLINE,", "") & " " & Ddbo & "APPJHTEMP.POSTSEQNCE, " & Ddbo & "APOBLTEMP.DATEBUS, APPJD.CNTLINE ," & Ddbo & "APOBLTEMP.IDVEND,TXAUTH.RATERECOV " & Environment.NewLine
            End If

        End If
        rs.Open(sqlstr, cnACCPAC)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessAP_PO> Select AP-PO Complete")
        End If


        '***************** Set Data To Temp Variable ==> Not Change
        Do While rs.options.Rs.EOF = False
            loCls = New clsVat


            With loCls

                ErrMsg = "Ledger = AP_PO IDINVC=" & Trim(rs.options.Rs.Collect(2))
                ErrMsg = ErrMsg & " Batch No.=" & Trim(rs.options.Rs.Collect(11))
                ErrMsg = ErrMsg & " Entry No.=" & Trim(rs.options.Rs.Collect(12))
                Dim INVTAXtemp As Decimal
                Dim TEXTLINE() As String = {}
                TEXTLINE = rs.options.Rs.Collect("TEXTLINE").ToString.Split(",")
                .VATCOMMENT = rs.options.Rs.Collect("TEXTLINE").ToString.Trim
                AMTEXTNDHC = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8)) 'MIN(APOBL.AMTBASE1HC) AS AMTEXTNDHC
                AMTGLDIST = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17)) 'MIN(APOBL.AMTBASE1HC) AS AMTGLDIST

                If rs.options.Rs.Collect("VATAS").ToString.Trim = "HEADER" Then
                    .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect(0)), 0, (rs.options.Rs.Collect(0))) 'APOBL.DATEINVC
                    .TXDATE = rs.options.Rs.Collect("DATEBUS")
                    .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", (rs.options.Rs.Collect(2))) 'APOBL.IDINVC
                    If IsDBNull(rs.options.Rs.Collect(14)) Then
                        .DOCNO = ""
                    Else
                        .DOCNO = rs.options.Rs.Collect(14) 'PORCPH1.RCPNUMBER
                    End If


                    If Len(.DOCNO) = 0 Then .DOCNO = IIf(IsDBNull(rs.options.Rs.Collect(15)), "", rs.options.Rs.Collect(15)) 'POCRNH1.RETNUMBER


                    If Len(.INVNAME) = 0 Then
                        .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect(4)), "", (rs.options.Rs.Collect(4))) 'APVEN.VENDNAME
                        Code = IIf(IsDBNull(rs.options.Rs.Collect("IDVEND")), "", (rs.options.Rs.Collect("IDVEND")))
                        If .INVNAME.Trim <> "" Then
                            If ComVatName = True And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" Then 'เพิ่ม And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" กรณีที่ ชื้อ Customer ยาว แล้วต้องคีย์เพิ่มที่ช่อง TEXTSTRE1 (เจอที่ AIT)
                                .INVNAME = CStr(.INVNAME.Trim + IIf(IsDBNull(rs.options.Rs.Collect(5)), "", (Trim(rs.options.Rs.Collect(5))))) 'APVEN.TEXTSTRE1
                            End If

                        End If
                    End If


                    If IsDBNull(rs.options.Rs.Collect(6)) Then Exit Do

                    If Trim(rs.options.Rs.Collect(6)) = "VAT" Then
                        .INVTAX = IIf(iffs.ISNULL(rs.options.Rs, 7), 0, rs.options.Rs.Collect(7)) ' APOBL.AMTINVCHC
                        If .INVTAX <> 0 Then .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8)) / rs.options.Rs.Fields("CountDup").Value ' APOBL.AMTEXTNDHC
                    Else
                        If .INVAMT = 0 Then
                            ' 17 เอาใช้กรณีที่ บรรทัดแรก มีภาษี บันทัดที่สองไม่มีภาษี
                            '.INVAMT = IIf(IsNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7)) ' APOBL.AMTINVCHC       '------------------ Edit by X [09/09/2005]
                            '.INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17)) '/ rs.options.Rs.Fields("CountDup") ' APIBD.AMTGLDIST      '------------------ Edit by X [09/09/2005]  ++++ wut 22/02/2008 add       / rs.options.Rs.Fields("CountDup")
                            '.INVAMT = .INVAMT - IIf(IsNull(rs.options.Rs.Collect(22)), 0, rs.options.Rs.Collect(22)) ' APIBD.AMTTAXREC1
                            '***************by kwan
                            If rs.options.Rs.Collect(23) = 1 Then

                                'If rs.options.Rs.Collect(8) < rs.options.Rs.Collect(17) Then 'แก้ไขให้เก็บในตัวแปร
                                If AMTEXTNDHC < AMTGLDIST Then
                                    .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8))
                                Else
                                    .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17))
                                End If
                            Else
                                If rs.options.Rs.Collect(23) > 0 Then
                                    'If rs.options.Rs.Collect(8) = rs.options.Rs.Collect(17) Then 'แก้ไขให้เก็บในตัวแปร
                                    If AMTEXTNDHC = AMTGLDIST Then
                                        .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17)) ' / rs.options.Rs.Collect(23)
                                    Else
                                        'If rs.options.Rs.Collect(8) < rs.options.Rs.Collect(17) Then 'แก้ไขให้เก็บในตัวแปร
                                        If AMTEXTNDHC < AMTGLDIST Then
                                            .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8))
                                        Else
                                            .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17))
                                        End If
                                    End If
                                Else
                                    .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17))
                                End If
                            End If
                        End If


                        INVTAXtemp = IIf(IsDBNull(rs.options.Rs.Collect(9)), 0, rs.options.Rs.Collect(9)) 'APOBL.AMTTAXHC
                        .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(22)), 0, rs.options.Rs.Collect(22)) 'MIN(APIBH.AMTRECTAX) AS AMTTAXREC1

                        '--------- แก้ไขให้รองรับ VAT AVG เนื่องจากเปลี่ยนฟิลด์หยิบ (APIBH.AMTRECTAX) และฟิลด์ที่หยิบจะได้ตัวเลขที่เป็น บวกทั้งหมด แต่ตัวเลขต้องแสดงตามจริงของฟิลด์เดิม (APOBL.AMTTAXHC)จึงต้องเช็คเงื่อนไขเพื่อ (*-1) ------
                        If INVTAXtemp < 0 Then
                            .INVTAX = .INVTAX * (-1)
                        End If
                        '-------- end comment -----
                        If .INVTAX = 0 Then
                            .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(9)), 0, rs.options.Rs.Collect(9)) 'APOBL.AMTTAXHC
                        End If

                        If rs.options.Rs.Collect(18) = 3 And .INVTAX > 0 Then '----------- If TXTTRXTYPE = 3 (Credit Note)
                            .INVTAX = .INVTAX * (-1)
                        End If

                    End If 'End if "VAT" 

                    If .INVTAX < 0 Then
                        If .INVAMT > 0 Then
                            .INVAMT = .INVAMT * (-1)
                        End If
                    End If

                    .TRANSNBR = ""

                    If OPT_AP = True Then 'ดึง VAT จาก Optional field
                        Dim RS_WHT_Header As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from APIBHO  where OPTFIELD in ('TAXBASE','TAXINVNO','TAXDATE','TAXNAME') and CNTBTCH ='" & rs.options.Rs.Collect("CNTBTCH") & "' and CNTITEM='" & rs.options.Rs.Collect("CNTITEM") & "'", cnACCPAC)

                        Dim tINVAMT As Double = 0
                        Dim tTAXINVNO As String = ""
                        Dim tTAXDATE As String = ""
                        Dim tTAXNAME As String = ""
                        For iwht As Integer = 0 To RS_WHT_Header.options.QueryDT.Rows.Count - 1
                            If RS_WHT_Header.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXBASE" Then
                                Dim Trydec As String = Trim(RS_WHT_Header.options.QueryDT.Rows(iwht).Item("VALUE"))
                                If Decimal.TryParse(Trydec, 0) = True Then
                                    If CDec(Trydec) <> 0 Then
                                        tINVAMT = CDec(Trydec)
                                    Else
                                        tINVAMT = 0
                                    End If
                                Else
                                    tINVAMT = 0
                                End If
                            ElseIf RS_WHT_Header.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXINVNO" Then
                                If Trim(RS_WHT_Header.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    tTAXINVNO = Trim(RS_WHT_Header.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If
                            ElseIf RS_WHT_Header.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXDATE" Then
                                Dim myDate As String = Trim(RS_WHT_Header.options.QueryDT.Rows(iwht).Item("VALUE"))
                                If Decimal.TryParse(myDate, 0) = True Then
                                    If CDec(myDate) <> 0 Then
                                        tTAXDATE = CDec(myDate)
                                    End If
                                End If
                            ElseIf RS_WHT_Header.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXNAME" Then

                                If Trim(RS_WHT_Header.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    tTAXNAME = Trim(RS_WHT_Header.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If
                            End If
                        Next
                        If tINVAMT <> 0 Then
                            .INVAMT = tINVAMT
                        End If
                        If tTAXINVNO <> "" Then
                            .IDINVC = tTAXINVNO
                        End If
                        If tTAXDATE <> "" Then
                            .INVDATE = tTAXDATE
                        End If
                        If tTAXNAME <> "" Then
                            .INVNAME = tTAXNAME
                        End If
                    End If


                    If VNO_AP = True Then

                        Dim VNO_Header As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from APIBHO  where OPTFIELD in ('APINVNO') and CNTBTCH ='" & rs.options.Rs.Collect("CNTBTCH") & "' and CNTITEM='" & rs.options.Rs.Collect("CNTITEM") & "'", cnACCPAC)
                        For iwht As Integer = 0 To VNO_Header.options.QueryDT.Rows.Count - 1
                            If VNO_Header.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "APINVNO" Then
                                If Trim(VNO_Header.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    .DOCNO = Trim(VNO_Header.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If

                            End If
                        Next

                    End If


                ElseIf rs.options.Rs.Collect("VATAS").ToString.Trim = "DETAIL" Then


                    If Trim(rs.options.Rs.Collect(6).ToString) = "VAT" Then
                        .INVTAX = IIf(iffs.ISNULL(rs.options.Rs, 7), 0, rs.options.Rs.Collect(7)) ' APOBL.AMTINVCHC
                        If .INVTAX <> 0 Then .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8)) / rs.options.Rs.Fields("CountDup").Value ' APOBL.AMTEXTNDHC
                    Else
                        If .INVAMT = 0 Then
                            ' 17 เอาใช้กรณีที่ บรรทัดแรก มีภาษี บันทัดที่สองไม่มีภาษี
                            '.INVAMT = IIf(IsNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7)) ' APOBL.AMTINVCHC       '------------------ Edit by X [09/09/2005]
                            '.INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17)) '/ rs.options.Rs.Fields("CountDup") ' APIBD.AMTGLDIST      '------------------ Edit by X [09/09/2005]  ++++ wut 22/02/2008 add       / rs.options.Rs.Fields("CountDup")
                            '.INVAMT = .INVAMT - IIf(IsNull(rs.options.Rs.Collect(22)), 0, rs.options.Rs.Collect(22)) ' APIBD.AMTTAXREC1
                            If rs.options.Rs.Collect(23) = 1 Then
                                'If rs.options.Rs.Collect(8) < rs.options.Rs.Collect(17) Then
                                If AMTEXTNDHC < AMTGLDIST Then

                                    .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8))
                                Else
                                    .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17))
                                End If
                            Else
                                If rs.options.Rs.Collect(23) > 0 Then
                                    If AMTEXTNDHC = AMTGLDIST Then
                                        .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17)) / rs.options.Rs.Collect(23)
                                    Else
                                        If AMTEXTNDHC < AMTGLDIST Then
                                            .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8))
                                        Else
                                            .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17))
                                        End If
                                    End If
                                Else
                                    .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17))
                                End If

                            End If
                        End If
                        .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(9)), 0, rs.options.Rs.Collect(9))
                        If .INVTAX = 0 Then
                            .INVTAX = AMTGLDIST 'rs.options.Rs.Collect("AMTGLDIST")
                        End If

                        If rs.options.Rs.Collect(18) = 3 And .INVTAX > 0 Then '----------- If TXTTRXTYPE = 3 (Credit Note)
                            .INVTAX = .INVTAX * (-1)
                        End If
                    End If

                    If .INVTAX < 0 Then
                        If .INVAMT > 0 Then
                            .INVAMT = .INVAMT * (-1)
                        End If
                    End If

                    .IDINVC = rs.options.Rs.Collect("IDINVC")
                    .TXDATE = rs.options.Rs.Collect("DATEBUS")
                    .INVDATE = rs.options.Rs.Collect("DATEINVC")
                    .DOCNO = rs.options.Rs.Collect("IDPONBR")
                    .TRANSNBR = rs.options.Rs.Collect("CNTLINE")
                    INVTAXtemp = 0

                    If OPT_AP = False Then

                        For i As Integer = 0 To TEXTLINE.Length - 1
                            If IsDBNull(TEXTLINE(i)) = True Then
                                TEXTLINE(i) = ""
                            End If
                        Next
                        If TEXTLINE.Length >= 1 Then
                            If TEXTLINE(0).Trim <> "" And Decimal.TryParse(TEXTLINE(0), 0) Then

                                .INVAMT = CDec(TEXTLINE(0))

                            End If
                        End If

                        If TEXTLINE.Length >= 2 Then
                            If TEXTLINE(1).Trim <> "" Then

                                .IDINVC = TEXTLINE(1)

                            End If
                        End If

                        If TEXTLINE.Length >= 3 Then
                            If TEXTLINE(2).Trim <> "" Then

                                .INVDATE = TryDate2(ZeroDate(TEXTLINE(2), "dd/MM/yyyy", "yyyyMMdd"), "dd/MM/yyyy", "yyyyMMdd")
                                If .INVDATE = 0 Then
                                    .INVDATE = TryDate2(TEXTLINE(2), "dd/MM/yyyy", "yyyyMMdd")
                                End If

                            End If
                        End If

                        If TEXTLINE.Length >= 4 Then
                            If TEXTLINE(3).Trim = "" Then
                                .INVNAME = rs.options.Rs.Collect("IDVEND")
                            Else
                                .INVNAME = TEXTLINE(3)
                                For CName As Integer = 4 To TEXTLINE.Length - 1
                                    .INVNAME = .INVNAME & "," & TEXTLINE(CName)
                                Next
                            End If
                        End If

                    Else 'OPT_AP= True 

                        Dim RS_WHT_Detail As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from APIBDO  where OPTFIELD in ('TAXBASE','TAXINVNO','TAXDATE','TAXNAME','TAXID','BRANCH') and CNTBTCH ='" & rs.options.Rs.Collect("CNTBTCH") & "' and CNTITEM='" & rs.options.Rs.Collect("CNTITEM") & "' and CNTLINE ='" & IIf(IsDBNull(rs.options.Rs.Collect("CNTLINE")), "0", rs.options.Rs.Collect("CNTLINE")) & "' ", cnACCPAC)

                        Dim tINVAMT As Double = 0
                        Dim tTAXINVNO As String = ""
                        Dim tTAXDATE As String = ""
                        Dim tTAXNAME As String = ""
                        Dim tTAXID As String = ""
                        Dim tBRANCH As String = ""


                        For iwht As Integer = 0 To RS_WHT_Detail.options.QueryDT.Rows.Count - 1
                            If RS_WHT_Detail.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXBASE" Then
                                Dim Trydec As String = Trim(RS_WHT_Detail.options.QueryDT.Rows(iwht).Item("VALUE"))
                                If Decimal.TryParse(Trydec, 0) = True Then
                                    If CDec(Trydec) <> 0 Then
                                        tINVAMT = CDec(Trydec)
                                    Else
                                        tINVAMT = 0
                                    End If
                                Else
                                    tINVAMT = 0
                                End If
                            ElseIf RS_WHT_Detail.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXINVNO" Then
                                If Trim(RS_WHT_Detail.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    tTAXINVNO = Trim(RS_WHT_Detail.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If
                            ElseIf RS_WHT_Detail.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXDATE" Then
                                Dim myDate As String = Trim(RS_WHT_Detail.options.QueryDT.Rows(iwht).Item("VALUE"))
                                If Decimal.TryParse(myDate, 0) = True Then
                                    If CDec(myDate) <> 0 Then
                                        tTAXDATE = CDec(myDate)
                                    End If
                                End If
                            ElseIf RS_WHT_Detail.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXNAME" Then
                                If Trim(RS_WHT_Detail.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    Dim chkIDVEND_Opt As New BaseClass.BaseODBCIO("select VENDORID,APVNR.RMITNAME AS RMITNAME from APVEN inner join APVNR on APVEN.VENDORID = APVNR.IDVEND where IDVENDRMIT = 'TXTHAI' and VENDORID = '" & Trim(RS_WHT_Detail.options.QueryDT.Rows(iwht).Item("VALUE")) & "'", cnACCPAC)
                                    If chkIDVEND_Opt.options.QueryDT.Rows.Count = 0 Then
                                        tTAXNAME = Trim(RS_WHT_Detail.options.QueryDT.Rows(iwht).Item("VALUE"))
                                    Else
                                        tTAXNAME = chkIDVEND_Opt.options.QueryDT.Rows(0).Item("RMITNAME").ToString
                                    End If
                                End If

                            ElseIf RS_WHT_Detail.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXID" Then
                                If Trim(RS_WHT_Detail.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    tTAXID = Trim(RS_WHT_Detail.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If

                            ElseIf RS_WHT_Detail.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "BRANCH" Then
                                If Trim(RS_WHT_Detail.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    tBRANCH = Trim(RS_WHT_Detail.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If
                            End If
                        Next

                        If tINVAMT <> 0 Then
                            .INVAMT = tINVAMT
                        Else
                            .INVAMT = 0
                        End If
                        If tTAXINVNO <> "" Then
                            .IDINVC = tTAXINVNO
                        End If
                        If tTAXDATE <> "" Then
                            .INVDATE = tTAXDATE
                        End If
                        If tTAXNAME <> "" Then
                            .INVNAME = tTAXNAME
                        End If

                        If tTAXID <> "" Then
                            .TAXID = tTAXID
                        End If
                        If tBRANCH <> "" Then
                            .Branch = tBRANCH
                        End If
                    End If


                    If VNO_AP = True Then

                        Dim VNO_Detail As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from APIBDO  where OPTFIELD in ('APINVNO') and CNTBTCH ='" & rs.options.Rs.Collect("CNTBTCH") & "' and CNTITEM='" & rs.options.Rs.Collect("CNTITEM") & "'", cnACCPAC)
                        For iwht As Integer = 0 To VNO_Detail.options.QueryDT.Rows.Count - 1
                            If VNO_Detail.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "APINVNO" Then
                                If Trim(VNO_Detail.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    .DOCNO = Trim(VNO_Detail.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If

                            End If
                        Next

                    End If


                End If 'End If "DETAIL" "HEADER"

                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect(10)), "", rs.options.Rs.Collect(10)) 'APPJD.IDAPACCT
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(11)), "", (rs.options.Rs.Collect(11))) 'APOBL.CNTBTCH
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(12)), "", (rs.options.Rs.Collect(12))) 'APOBL.CNTITEM
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(16)), "", (rs.options.Rs.Collect(16))) 'APOBL.IDPONBR
                .TypeOfPU = IIf(IsDBNull(rs.options.Rs.Collect("NAMECITY")), "", (rs.options.Rs.Collect("NAMECITY")))
                .Runno = IIf(IsDBNull(rs.options.Rs.Collect("value")), "", rs.options.Rs.Collect("value"))

                If .INVAMT = 0 Then
                    .RATE_Renamed = 0
                Else
                    .RATE_Renamed = FormatNumber(((.INVTAX / .INVAMT) * 100), 1)
                End If


                If CStr(.INVDATE) = 0 Then
                    .INVDATE = Trim(rs.options.Rs.Collect("DATEINVC"))
                End If
                If CStr(.TXDATE) = 0 Then
                    .TXDATE = Trim(rs.options.Rs.Collect("DATEBUS"))
                End If

                If FindMiscCodeInAPVEN(Trim(.INVNAME), "AP") = "0" Or FindMiscCodeInAPVEN(Trim(.INVNAME), "AP") = "" Then
                    If Trim(.INVNAME) = "" Then
                        .INVNAME = FindMiscCodeInAPVEN(Trim(rs.options.Rs.Collect("IDVEND")), "AP")
                    Else
                        .INVNAME = Trim(.INVNAME)
                    End If
                    If .TAXID = "" And .Branch = "" Then

                        .TAXID = FindTaxIDBRANCH(Trim(rs.options.Rs.Collect("IDVEND")), "AP", "TAXID")
                        .Branch = FindTaxIDBRANCH(Trim(rs.options.Rs.Collect("IDVEND")), "AP", "BRANCH")
                    End If
                Else
                    If .TAXID = "" And .Branch = "" Then
                        .TAXID = FindTaxIDBRANCH(.INVNAME, "AP", "TAXID")
                        .Branch = FindTaxIDBRANCH(.INVNAME, "AP", "BRANCH")
                        .INVNAME = FindMiscCodeInAPVEN(Trim(.INVNAME), "AP")

                    End If
                End If


                .VATCOMMENT = .INVAMT & "," & .IDINVC & "," & .INVDATE & "," & .INVNAME
                .MARK = "Invoice"

                sqlstr = InsVat(0, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, "", (.INVNAME), .INVAMT, .INVTAX, .LOCID, 1, .RATE_Renamed, .TTYPE, .ACCTVAT, "AP", .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, .Runno, IIf(ISNULL(rs.options.Rs, 13), "", Trim(rs.options.Rs.Collect(13))), IIf(ISNULL(rs.options.Rs, 13), "", Trim(rs.options.Rs.Collect(13))), .TypeOfPU, "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, INVTAXtemp)

            End With
            cnVAT.Execute(sqlstr)

            ErrMsg = ""
            Code = ""
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing

        'transfer data
        rs = New BaseClass.BaseODBCIO
        If ComDatabase = "ORCL" Then
            sqlstr = "SELECT FMSVATTEMP.INVDATE,FMSVATTEMP.TXDATE,FMSVATTEMP.IDINVC,FMSVATTEMP.DOCNO," '0-3
            sqlstr = sqlstr & " FMSVATTEMP.INVNAME,FMSVATTEMP.INVAMT,FMSVATTEMP.INVTAX," '4-6
            sqlstr = sqlstr & " FMSVLACC.LOCID,FMSVATTEMP.VTYPE,FMSVATTEMP.RATE,FMSTAX.TTYPE," '7-10
            sqlstr = sqlstr & " FMSTAX.ACCTVAT,FMSVATTEMP.SOURCE,FMSVATTEMP.BATCH,FMSVATTEMP.ENTRY," '11-14
            sqlstr = sqlstr & " FMSVATTEMP.MARK,FMSVATTEMP.VATCOMMENT,FMSTAX.ITEMRATE1,FMSVATTEMP.IDDIST,FMSVATTEMP.CBREF,FMSVATTEMP.RUNNO,FMSVATTEMP.TypeofPu,FMSVATTEMP.tranNo,FMSVATTEMP.Cif,FMSVATTEMP.taxCif,FMSVATTEMP.TRANSNBR, " '15-20
            sqlstr = sqlstr & " FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH,FMSVATTEMP.Code,FMSVATTEMP.TOTTAX,FMSVATTEMP.CODETAX "
            sqlstr = sqlstr & " FROM FMSVATTEMP,FMSTAX,FMSVLACC "
            sqlstr = sqlstr & " WHERE (FMSVATTEMP.CODETAX = FMSTAX.AUTHORITY AND  FMSTAX.ACCTVAT = FMSVATTEMP.ACCTVAT)"
            sqlstr = sqlstr & " AND FMSTAX.ACCTVAT = FMSVLACC.ACCTVAT"
            sqlstr = sqlstr & " AND (FMSTAX.TTYPE = 1 or FMSTAX.TTYPE = 3) AND (FMSTAX.BUYERCLASS = 1)"
        Else
            sqlstr = " SELECT FMSVATTEMP.INVDATE, FMSVATTEMP.TXDATE, FMSVATTEMP.IDINVC, FMSVATTEMP.DOCNO," '0-3
            sqlstr = sqlstr & " FMSVATTEMP.INVNAME, sum(FMSVATTEMP.INVAMT) AS INVAMT, sum(FMSVATTEMP.INVTAX) AS INVTAX," '4-6
            sqlstr = sqlstr & " FMSVLACC.LOCID, FMSVATTEMP.VTYPE, FMSVATTEMP.RATE, FMSTAX.TTYPE," '7-10
            sqlstr = sqlstr & " FMSTAX.ACCTVAT,FMSVATTEMP.SOURCE, FMSVATTEMP.BATCH, FMSVATTEMP.ENTRY," '11-14
            sqlstr = sqlstr & " FMSVATTEMP.MARK, FMSVATTEMP.VATCOMMENT, FMSTAX.ITEMRATE1, FMSVATTEMP.IDDIST , FMSVATTEMP.CBREF,FMSVATTEMP.RUNNO,FMSVATTEMP.RUNNO,FMSVATTEMP.TypeofPu,FMSVATTEMP.tranNo,FMSVATTEMP.Cif,FMSVATTEMP.taxCif,FMSVATTEMP.TRANSNBR, " '15-20
            sqlstr = sqlstr & " FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH,FMSVATTEMP.Code,FMSVATTEMP.TOTTAX,FMSVATTEMP.CODETAX "
            sqlstr = sqlstr & " FROM FMSVATTEMP "
            sqlstr = sqlstr & "     INNER JOIN FMSTAX ON (FMSVATTEMP.CODETAX = FMSTAX.AUTHORITY) and  (FMSTAX.ACCTVAT= FMSVATTEMP.ACCTVAT)"
            sqlstr = sqlstr & "     INNER JOIN FMSVLACC ON FMSTAX.ACCTVAT = FMSVLACC.ACCTVAT "
            sqlstr = sqlstr & " Where (FMSTAX.BUYERCLASS = 1) "
            sqlstr = sqlstr & " GROUP BY FMSVATTEMP.INVDATE, FMSVATTEMP.TXDATE, FMSVATTEMP.IDINVC, FMSVATTEMP.DOCNO, FMSVATTEMP.INVNAME, FMSVLACC.LOCID,"
            sqlstr = sqlstr & "     FMSVATTEMP.VTYPE, FMSVATTEMP.RATE, FMSTAX.TTYPE, FMSTAX.ACCTVAT, FMSVATTEMP.SOURCE, FMSVATTEMP.BATCH, FMSVATTEMP.ENTRY,"
            sqlstr = sqlstr & "     FMSVATTEMP.MARK , FMSVATTEMP.VATCOMMENT, FMSTAX.ITEMRATE1, FMSVATTEMP.IDDIST, FMSVATTEMP.CBRef,FMSVATTEMP.RUNNO,FMSVATTEMP.RUNNO,FMSVATTEMP.TypeofPu,FMSVATTEMP.tranNo,FMSVATTEMP.Cif,FMSVATTEMP.taxCif,FMSVATTEMP.TRANSNBR, "
            sqlstr = sqlstr & "     FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH,FMSVATTEMP.Code,FMSVATTEMP.TOTTAX,FMSVATTEMP.CODETAX "
            sqlstr = sqlstr & " HAVING (FMSTAX.TTYPE = 1 OR FMSTAX.TTYPE = 3)"
        End If

        rs.Open(sqlstr, cnVAT)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessAP_PO> Select FMSVATTEMP Complete")
        End If


        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls
                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect(0)), 0, rs.options.Rs.Collect(0))
                .TXDATE = IIf(IsDBNull(rs.options.Rs.Collect(1)), 0, rs.options.Rs.Collect(1))
                .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", (rs.options.Rs.Collect(2)))
                .DOCNO = IIf(IsDBNull(rs.options.Rs.Collect(3)), "", (rs.options.Rs.Collect(3)))
                .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect(4)), "", (rs.options.Rs.Collect(4)))
                .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(5)), 0, rs.options.Rs.Collect(5))
                .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(6)), 0, rs.options.Rs.Collect(6))
                .LOCID = IIf(IsDBNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7))
                .VTYPE = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8))
                .TTYPE = IIf(IsDBNull(rs.options.Rs.Collect(10)), 0, rs.options.Rs.Collect(10))
                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect(11)), "", (rs.options.Rs.Collect(11)))
                .Source = IIf(IsDBNull(rs.options.Rs.Collect(12)), "", (rs.options.Rs.Collect(12)))
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(13)), "", (rs.options.Rs.Collect(13)))
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(14)), "", (rs.options.Rs.Collect(14)))
                .MARK = IIf(IsDBNull(rs.options.Rs.Collect(15)), "", (rs.options.Rs.Collect(15)))
                .VATCOMMENT = IIf(IsDBNull(rs.options.Rs.Collect(16)), "", (rs.options.Rs.Collect(16)))
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(19)), "", (rs.options.Rs.Collect(19)))
                .TRANSNBR = IIf(IsDBNull(rs.options.Rs.Collect("TRANSNBR")), "", (rs.options.Rs.Collect("TRANSNBR")))
                .TypeOfPU = IIf(IsDBNull(rs.options.Rs.Collect("TypeOfPU")), "", rs.options.Rs.Collect("TypeOfPU"))
                .TranNo = IIf(IsDBNull(rs.options.Rs.Collect("TranNo")), "", rs.options.Rs.Collect("TranNo"))
                .CIF = IIf(IsDBNull(rs.options.Rs.Collect("CIF")), "0", rs.options.Rs.Collect("CIF"))
                .TaxCIF = IIf(IsDBNull(rs.options.Rs.Collect("TaxCIF")), "0", rs.options.Rs.Collect("TaxCIF"))
                .TAXID = IIf(IsDBNull(rs.options.Rs.Collect("TAXID")), "", rs.options.Rs.Collect("TAXID"))
                .Branch = IIf(IsDBNull(rs.options.Rs.Collect("BRANCH")), "", rs.options.Rs.Collect("BRANCH"))
                .Runno = IIf(IsDBNull(rs.options.Rs.Collect("runno")), "", rs.options.Rs.Collect("runno"))
                Code = IIf(IsDBNull(rs.options.Rs.Collect("Code")), "", (rs.options.Rs.Collect("Code")))
                .TOTTAX = IIf(IsDBNull(rs.options.Rs.Collect("TOTTAX")), 0, (rs.options.Rs.Collect("TOTTAX")))
                .RATE_Renamed = IIf(IsDBNull(rs.options.Rs.Collect("RATE")), 0, (rs.options.Rs.Collect("RATE")))
                .CODETAX = IIf(IsDBNull(rs.options.Rs.Collect("CODETAX")), "", rs.options.Rs.Collect("CODETAX"))
            
                sqlstr = InsVat(1, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, .NEWDOCNO, (.INVNAME), .INVAMT, .INVTAX, .LOCID, .VTYPE, .RATE_Renamed, .TTYPE, .ACCTVAT, .Source, .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, .Runno, .CODETAX, "", .TypeOfPU, .TranNo, .CIF, .TaxCIF, 0, .TAXID, .Branch, Code, "", 0, 0, .TOTTAX)
            End With
            cnVAT.Execute(sqlstr)
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing
        cnACCPAC.Close()
        cnVAT.Close()
        cnACCPAC = Nothing
        cnVAT = Nothing
        Exit Sub
ErrHandler:
        WriteLog("<ProcessAP_PO>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessAP_PO> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
        WriteLog(ErrMsg)
    End Sub

    Public Sub ProcessAPPay()
        Dim cnACCPAC, cnVAT As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim strsplit() As String
        Dim losplit As Integer
        Dim loCls As clsVat
        Dim tmpTAXBASE, tmpTAXDATE As String
        Dim tmpTAXNAME As String
        Dim iffs As clsFuntion
        iffs = New clsFuntion
        ErrMsg = ""

        tmpTAXDATE = 0
        tmpTAXBASE = 0
        Dim VendorCode As String

        On Error GoTo ErrHandler
        cnACCPAC = New ADODB.Connection
        cnVAT = New ADODB.Connection

        cnACCPAC.ConnectionTimeout = 60
        cnACCPAC.Open(ConnACCPAC)
        cnACCPAC.CommandTimeout = 3600

        cnVAT.ConnectionTimeout = 60
        cnVAT.Open(ConnVAT)
        cnVAT.CommandTimeout = 3600

        'gather data ==> Select Data Uses For Module
        rs = New BaseClass.BaseODBCIO
        sqlstr = "SELECT    APBTA.PAYMTYPE, APBTA.CNTBTCH, APTCR.DATERMIT AS INVDATE, '' AS TXDATE, APTCP.IDINVC, '' AS DOCNO, '' AS NEWDOCNO, '' AS INVNAME, "
        sqlstr &= "         ''  AS INVAMT, APTCU.amtdisthc AS INVTAX, APTCP.AUDTORG AS LOCID, '1' AS VTYPE, '' AS RATE, '1' AS TTTYPE, APTCU.IDACCT,"
        sqlstr &= "         APBTA.SRCEAPPL AS SOURCE, APTCR.CNTBTCH AS BATCH, APTCR.CNTENTR AS ENTRY, '' AS MARK, APTCU.TEXTREF AS VATCOMMENT,APTCU.IDDISTCODE AS CODETAX, "
        sqlstr &= "         APTCR.DOCNBR AS CBREF,APTCR.IDVEND,(case when APTCN.GLDESC IS null then APTCU.TEXTDESC else APTCN.GLDESC end) as GLDESC,APTCR.DATEBUS,APTCU.CNTSEQ "
        sqlstr &= " FROM    APBTA "
        sqlstr &= "         INNER JOIN APTCR ON APBTA.CNTBTCH = APTCR.CNTBTCH AND  APBTA.PAYMTYPE = APTCR.BTCHTYPE  "
        sqlstr &= "         INNER JOIN APTCP ON APTCR.BTCHTYPE = APTCP.BATCHTYPE AND APTCR.CNTBTCH = APTCP.CNTBTCH AND APTCR.CNTENTR = APTCP.CNTRMIT "
        sqlstr &= "         INNER JOIN APTCU ON APTCP.BATCHTYPE = APTCU.BATCHTYPE AND APTCP.CNTBTCH = APTCU.CNTBTCH AND APTCP.CNTRMIT = APTCU.CNTRMIT AND APTCP.CNTLINE = APTCU.CNTLINE "
        sqlstr &= "         LEFT JOIN APTCN ON APTCR.CNTBTCH = APTCN.CNTBTCH and APTCR.CNTENTR = APTCN.CNTRMIT"
        sqlstr &= " WHERE   (APBTA.PAYMTYPE = 'PY' OR APBTA.PAYMTYPE = 'AD') AND (APTCR.RMITTYPE = 1 OR APTCR.RMITTYPE = 5) AND (APTCR." & IIf(DATEMODE = "DOCU", "DATERMIT", "DATEBUS") & " >= " & DATEFROM & " AND APTCR." & IIf(DATEMODE = "DOCU", "DATERMIT", "DATEBUS") & " <= " & DATETO & ") "

        If Use_POST Then
            sqlstr &= " AND (APBTA.BATCHSTAT = 3 AND APTCR.ERRENTRY = 0)" 'แก้ไขจาก APBTA.NBRERRORS = 0  เป็น APTCR.ERRENTRY = 0   เจอ Case AICHI Batch Error จะไม่แสดงทั้ง Batch จึงเปลี่ยนให้ เช็ค Entry แทน เอาเฉพาะ Entry ที่ Error = 0   date 24/04/2014_Pat
        Else
            sqlstr &= " AND (APTCR.ERRENTRY = 0)"
        End If

        sqlstr &= " AND (APTCU.IDACCT IN (select ACCTVAT " & IIf(ComDatabase = "PVSW", "", "COLLATE Thai_CI_AS") & " from " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "FMSVLACC) "
        sqlstr &= " OR  APTCN.IDACCT IN (select ACCTVAT " & IIf(ComDatabase = "PVSW", "", "COLLATE Thai_CI_AS") & " from " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "FMSVLACC))"

        rs.Open(sqlstr, cnACCPAC)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessAPPay> Select APPay Complete")
        End If


        '***************** Set Data To Temp Variable ==> Not Change
        Dim tableau_a_remplir() As String
        Dim nombre_elements, rep, i As Integer
        Dim splittxt As String
        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls
                ErrMsg = "Ledger = APPay IDINVC=" & Trim(rs.options.Rs.Collect(4))
                ErrMsg = ErrMsg & " Batch No.=" & Trim(rs.options.Rs.Collect(16))
                ErrMsg = ErrMsg & " Entry No.=" & Trim(rs.options.Rs.Collect(17))

                tmpTAXBASE = 0
                tmpTAXNAME = ""
                tmpTAXDATE = StrYearF

                If Trim(rs.options.Rs.Collect(19)).Length > 1 Then
                    If InStr(1, Trim(rs.options.Rs.Collect(19)), ",") > 0 Then
                        strsplit = Trim(rs.options.Rs.Collect(19)).Split(",")
                        For losplit = LBound(strsplit) To UBound(strsplit)
                            splittxt = Trim(strsplit(losplit))
                            Select Case losplit
                                Case 0
                                    tmpTAXBASE = IIf(IsNumeric(splittxt) = True, splittxt, 0)
                                Case 1

                                    .IDINVC = Trim(splittxt)
                                Case 2

                                    tmpTAXDATE = TryDate2(ZeroDate(splittxt, "dd/MM/yyyy", "yyyyMMdd"), "dd/MM/yyyy", "yyyyMMdd")
                                    If tmpTAXDATE = 0 Then
                                        tmpTAXDATE = TryDate2(splittxt, "dd/MM/yyyy", "yyyyMMdd")
                                    End If
                                Case 3
                                    tmpTAXNAME = Trim(splittxt)
                                Case Is > 3
                                    tmpTAXNAME = tmpTAXNAME & "," & Trim(splittxt)
                            End Select
                        Next
                        If strsplit.Length > 3 And tmpTAXNAME.Length > 0 Then
                            tmpTAXNAME = tmpTAXNAME.Replace(",", "")
                        End If
                    End If
                End If

                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", (rs.options.Rs.Collect(2)))
                .TXDATE = rs.options.Rs.Collect("DATEBUS")
                If Val(tmpTAXDATE) <> 0 Then
                    .INVDATE = CDbl(Trim(tmpTAXDATE))
                End If

                If Len(.IDINVC) = 0 Then .IDINVC = ""
                If Len(.DOCNO) = 0 Then .DOCNO = ""

                If FindMiscCodeInAPVEN(Trim(tmpTAXNAME), "AP") = "0" Or FindMiscCodeInAPVEN(Trim(tmpTAXNAME), "AP") = "" Then
                    If Trim(tmpTAXNAME) = "" Then
                        .INVNAME = FindMiscCodeInAPVEN(Trim(rs.options.Rs.Collect("IDVEND")), "AP")
                    Else
                        .INVNAME = Trim(tmpTAXNAME)
                    End If
                    .TAXID = FindTaxIDBRANCH(Trim(rs.options.Rs.Collect("IDVEND")), "AP", "TAXID")
                    .Branch = FindTaxIDBranch(Trim(rs.options.Rs.Collect("IDVEND")), "AP", "BRANCH")
                Else
                    .TAXID = FindTaxIDBranch(tmpTAXNAME, "AP", "TAXID")
                    .Branch = FindTaxIDBranch(tmpTAXNAME, "AP", "BRANCH")
                    .INVNAME = FindMiscCodeInAPVEN(Trim(tmpTAXNAME), "AP")
                End If

                If VNO_AP = True Then

                    Dim VNO_Header As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from APTCRO  where OPTFIELD in ('APPYNO') and CNTBTCH ='" & rs.options.Rs.Collect("BATCH") & "' and CNTITEM='" & rs.options.Rs.Collect("ENTRY") & "'", cnACCPAC)
                    For iwht As Integer = 0 To VNO_Header.options.QueryDT.Rows.Count - 1
                        If VNO_Header.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "APPYNO" Then
                            If Trim(VNO_Header.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                .DOCNO = Trim(VNO_Header.options.QueryDT.Rows(iwht).Item("VALUE"))
                            End If

                        End If
                    Next

                End If

                .VATCOMMENT = .INVAMT & "," & .IDINVC & "," & .INVDATE & "," & .INVNAME
                .INVAMT = Val(tmpTAXBASE)
                .INVTAX = CDbl(Trim(rs.options.Rs.Collect(9)))

                If .INVTAX < 0 Then
                    .INVAMT = .INVAMT * (-1)
                End If


                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect(14)), "", rs.options.Rs.Collect(14))
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(16)), "", (rs.options.Rs.Collect(16))) 'APOBL.CNTBTCH
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(17)), "", (rs.options.Rs.Collect(17))) 'APOBL.CNTITEM
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(21)), "", (rs.options.Rs.Collect(21)))
                .TypeOfPU = IIf(IsDBNull(rs.options.Rs.Collect("GLDESC")), "", (rs.options.Rs.Collect("GLDESC")))
                .TRANSNBR = IIf(IsDBNull(rs.options.Rs.Collect("CNTSEQ")), "", (rs.options.Rs.Collect("CNTSEQ"))) 'APTCU.CNTLINE เพิ่มเก็บ Linenumber 12-02-2014 by Pat
                .TTYPE = 1
                .MARK = "Payment"


                If .INVAMT = 0 Then
                    .RATE_Renamed = 0
                Else
                    .RATE_Renamed = FormatNumber(((.INVTAX / .INVAMT) * 100), 1)
                End If

                sqlstr = InsVat(0, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, "", (.INVNAME), .INVAMT, .INVTAX, .LOCID, 1, .RATE_Renamed, .TTYPE, .ACCTVAT, "AP", .Batch, .Entry, .MARK, Trim(.VATCOMMENT), .CBRef, .TRANSNBR, .Runno, IIf(ISNULL(rs.options.Rs, 20), "", Trim(rs.options.Rs.Collect(20))), IIf(ISNULL(rs.options.Rs, 22), "", Trim(rs.options.Rs.Collect(22))), .TypeOfPU, "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, 0)
            End With
            cnVAT.Execute(sqlstr)
            ErrMsg = ""
            Code = ""
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing

        'transfer data
        rs = New BaseClass.BaseODBCIO
        If ComDatabase = "ORCL" Then
            sqlstr = "SELECT FMSVATTEMP.INVDATE,FMSVATTEMP.TXDATE,FMSVATTEMP.IDINVC,FMSVATTEMP.DOCNO," '0-3
            sqlstr = sqlstr & " FMSVATTEMP.INVNAME,FMSVATTEMP.INVAMT,FMSVATTEMP.INVTAX," '4-6
            sqlstr = sqlstr & " FMSVLACC.LOCID,FMSVATTEMP.VTYPE,FMSVATTEMP.RATE,FMSVATTEMP.TTYPE," '7-10
            sqlstr = sqlstr & " FMSVATTEMP.ACCTVAT,FMSVATTEMP.SOURCE,FMSVATTEMP.BATCH,FMSVATTEMP.ENTRY,FMSVATTEMP.TRANSNBR, " '11-14
            sqlstr = sqlstr & " FMSVATTEMP.MARK,FMSVATTEMP.VATCOMMENT,FMSVATTEMP.IDDIST,FMSVATTEMP.CBREF,FMSVATTEMP.TypeofPU,FMSVATTEMP.TranNo,FMSVATTEMP.CIF,FMSVATTEMP.TaxCIF, " '15-20
            sqlstr = sqlstr & " FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH,FMSVATTEMP.Code,FMSVATTEMP.CODETAX "
            sqlstr = sqlstr & " FROM FMSVATTEMP,FMSVLACC WHERE "
            sqlstr = sqlstr & " FMSVATTEMP.ACCTVAT = FMSVLACC.ACCTVAT"
            sqlstr = sqlstr & " AND (FMSVATTEMP.TTYPE = 1 or FMSVATTEMP.TTYPE = 3) "

        Else
            sqlstr = " SELECT   FMSVATTEMP.INVDATE, FMSVATTEMP.TXDATE, FMSVATTEMP.IDINVC, " & Environment.NewLine
            sqlstr &= "         FMSVATTEMP.DOCNO, FMSVATTEMP.INVNAME, SUM(FMSVATTEMP.INVAMT) AS INVAMT, " & Environment.NewLine
            sqlstr &= "         SUM(FMSVATTEMP.INVTAX) AS INVTAX, FMSVLACC.LOCID, FMSVATTEMP.VTYPE, " & Environment.NewLine
            sqlstr &= "         FMSVATTEMP.RATE, FMSVATTEMP.TTYPE, FMSVATTEMP.ACCTVAT,FMSVATTEMP.SOURCE, " & Environment.NewLine
            sqlstr &= "         FMSVATTEMP.BATCH, FMSVATTEMP.ENTRY,FMSVATTEMP.TRANSNBR, FMSVATTEMP.MARK, FMSVATTEMP.VATCOMMENT, " & Environment.NewLine
            sqlstr &= "         FMSVATTEMP.IDDIST , FMSVATTEMP.CBREF ,FMSVATTEMP.TypeofPU,FMSVATTEMP.TranNo, " & Environment.NewLine
            sqlstr &= "         FMSVATTEMP.CIF,FMSVATTEMP.TaxCIF,FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH,FMSVATTEMP.Code,FMSVATTEMP.CODETAX "
            sqlstr &= " FROM    FMSVATTEMP  " & Environment.NewLine
            sqlstr &= "         INNER JOIN FMSVLACC ON FMSVATTEMP.ACCTVAT = FMSVLACC.ACCTVAT  " & Environment.NewLine
            sqlstr &= " GROUP BY FMSVATTEMP.INVDATE, FMSVATTEMP.TXDATE, FMSVATTEMP.IDINVC, " & Environment.NewLine
            sqlstr &= "         FMSVATTEMP.DOCNO, FMSVATTEMP.INVNAME, FMSVLACC.LOCID, FMSVATTEMP.VTYPE, " & Environment.NewLine
            sqlstr &= "         FMSVATTEMP.RATE, FMSVATTEMP.TTYPE, FMSVATTEMP.ACCTVAT, FMSVATTEMP.SOURCE, " & Environment.NewLine
            sqlstr &= "         FMSVATTEMP.BATCH, FMSVATTEMP.ENTRY,FMSVATTEMP.TRANSNBR, FMSVATTEMP.MARK , FMSVATTEMP.VATCOMMENT, " & Environment.NewLine
            sqlstr &= "         FMSVATTEMP.IDDIST, FMSVATTEMP.CBRef ,FMSVATTEMP.TypeofPU,FMSVATTEMP.TranNo,FMSVATTEMP.CIF,FMSVATTEMP.TaxCIF," & Environment.NewLine
            sqlstr &= "         FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH,FMSVATTEMP.Code,FMSVATTEMP.CODETAX "
            sqlstr &= "  HAVING (FMSVATTEMP.TTYPE = 1 OR FMSVATTEMP.TTYPE = 3)" & Environment.NewLine

        End If
        rs.Open(sqlstr, cnVAT)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessAPPay> Select FMSVATTEMP Complete")
        End If


        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls
                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect("INVDATE")), 0, rs.options.Rs.Collect("INVDATE"))
                .TXDATE = IIf(IsDBNull(rs.options.Rs.Collect("TXDATE")), 0, rs.options.Rs.Collect("TXDATE"))
                .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect("IDINVC")), "", (rs.options.Rs.Collect("IDINVC")))
                .DOCNO = IIf(IsDBNull(rs.options.Rs.Collect("DOCNO")), "", (rs.options.Rs.Collect("DOCNO")))
                .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect("INVNAME")), "", (rs.options.Rs.Collect("INVNAME")))
                .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect("INVAMT")), 0, rs.options.Rs.Collect("INVAMT"))
                .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect("INVTAX")), 0, rs.options.Rs.Collect("INVTAX"))
                .LOCID = IIf(IsDBNull(rs.options.Rs.Collect("LOCID")), 0, rs.options.Rs.Collect("LOCID"))
                .VTYPE = IIf(IsDBNull(rs.options.Rs.Collect("VTYPE")), 0, rs.options.Rs.Collect("VTYPE"))
                .TTYPE = IIf(IsDBNull(rs.options.Rs.Collect("TTYPE")), 0, rs.options.Rs.Collect("TTYPE"))
                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect("ACCTVAT")), "", (rs.options.Rs.Collect("ACCTVAT")))
                .Source = IIf(IsDBNull(rs.options.Rs.Collect("SOURCE")), "", (rs.options.Rs.Collect("SOURCE")))
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect("BATCH")), "", (rs.options.Rs.Collect("BATCH")))
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect("ENTRY")), "", (rs.options.Rs.Collect(14)))
                .MARK = IIf(IsDBNull(rs.options.Rs.Collect("MARK")), "", (rs.options.Rs.Collect("MARK")))
                .VATCOMMENT = IIf(IsDBNull(rs.options.Rs.Collect("VATCOMMENT")), "", (rs.options.Rs.Collect("VATCOMMENT")))
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect("CBREF")), "", (rs.options.Rs.Collect("CBREF")))
                .TRANSNBR = IIf(IsDBNull(rs.options.Rs.Collect("TRANSNBR")), "", (rs.options.Rs.Collect("TRANSNBR")))
                .TypeOfPU = IIf(IsDBNull(rs.options.Rs.Collect("TypeOfPU")), "", (rs.options.Rs.Collect("TypeOfPU")))
                .TranNo = IIf(IsDBNull(rs.options.Rs.Collect("TranNo")), "", (rs.options.Rs.Collect("TranNo")))
                .CIF = IIf(IsDBNull(rs.options.Rs.Collect("CIF")), "", (rs.options.Rs.Collect("CIF")))
                .TaxCIF = IIf(IsDBNull(rs.options.Rs.Collect("TaxCIF")), "", (rs.options.Rs.Collect("TaxCIF")))
                .TAXID = IIf(IsDBNull(rs.options.Rs.Collect("TAXID")), "", (rs.options.Rs.Collect("TAXID")))
                .Branch = IIf(IsDBNull(rs.options.Rs.Collect("BRANCH")), "", (rs.options.Rs.Collect("BRANCH")))
                Code = IIf(IsDBNull(rs.options.Rs.Collect("Code")), "", (rs.options.Rs.Collect("Code")))
                .RATE_Renamed = IIf(IsDBNull(rs.options.Rs.Collect("RATE")), 0, (rs.options.Rs.Collect("RATE")))
                .CODETAX = IIf(IsDBNull(rs.options.Rs.Collect("CODETAX")), "", (rs.options.Rs.Collect("CODETAX")))

                'If .Batch = 39 And .Entry = 16 Then
                '    Dim sss As String = ""
                'End If

                sqlstr = InsVat(1, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, .NEWDOCNO, (.INVNAME), .INVAMT, .INVTAX, .LOCID, .VTYPE, .RATE_Renamed, .TTYPE, .ACCTVAT, .Source, .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, .Runno, .CODETAX, "", .TypeOfPU, .TranNo, .CIF, .TaxCIF, 0, .TAXID, .Branch, Code, "", 0, 0, 0)
            End With
            cnVAT.Execute(sqlstr)
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing
        cnACCPAC.Close()
        cnVAT.Close()
        cnACCPAC = Nothing
        cnVAT = Nothing
        Exit Sub

ErrHandler:
        WriteLog("<ProcessAPPay>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessAPPay> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
        WriteLog(ErrMsg)
    End Sub

    Public Sub ProcessARRec()
        Dim cnACCPAC, cnVAT As ADODB.Connection
        Dim rs, rs3 As BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim strsplit() As String
        Dim losplit As Integer
        Dim loCls As clsVat
        Dim tmpTAXBASE, tmpTAXDATE As String
        Dim tmpTAXNAME As String
        Dim iffs As clsFuntion
        iffs = New clsFuntion
        ErrMsg = ""


        Dim VendorCode As String

        On Error GoTo ErrHandler
        cnACCPAC = New ADODB.Connection
        cnVAT = New ADODB.Connection

        cnACCPAC.ConnectionTimeout = 60
        cnACCPAC.Open(ConnACCPAC)
        cnACCPAC.CommandTimeout = 3600

        cnVAT.ConnectionTimeout = 60
        cnVAT.Open(ConnVAT)
        cnVAT.CommandTimeout = 3600

        'gather data ==> Select Data Uses For Module
        rs3 = New BaseClass.BaseODBCIO
        sqlstr = "SELECT HOMECUR FROM CSCOM "
        rs3.Open(sqlstr, cnACCPAC)

        rs = New BaseClass.BaseODBCIO
        sqlstr = "SELECT    ARBTA.CODEPYMTYP, ARBTA.CNTBTCH, ARTCR.DATERMIT AS INVDATE, '' AS TXDATE, ARTCP.IDINVC, '' AS DOCNO, '' AS NEWDOCNO, '' AS INVNAME, "
        sqlstr &= "         ''  AS INVAMT, ARTCU.AMTDIST AS INVTAX, ARTCP.AUDTORG AS LOCID, '1' AS VTYPE, '' AS RATE, '1' AS TTTYPE, ARTCU.IDACCT,"
        sqlstr &= "         ARBTA.SRCEAPPL AS SOURCE, ARTCR.CNTBTCH AS BATCH, ARTCR.CNTITEM AS ENTRY, '' AS MARK, ARTCU.TEXTREF AS VATCOMMENT, ARTCU.IDDISTCODE AS CODETAX, ARTCR.DOCNBR AS CBREF,ARTCR.IDCUST, ARBTA.CODECURN, ARBTA.RATEEXCHHC,ARTCR.DATEBUS,ARTCU.CNTSEQ "
        sqlstr &= " FROM    ARBTA "
        sqlstr &= "         INNER JOIN ARTCR ON ARBTA.CNTBTCH = ARTCR.CNTBTCH AND  ARBTA.CODEPYMTYP = ARTCR.CODEPYMTYP "
        sqlstr &= "         INNER JOIN ARTCP ON ARTCR.CODEPYMTYP = ARTCP.CODEPAYM AND ARTCR.CNTBTCH = ARTCP.CNTBTCH AND ARTCR.CNTITEM = ARTCP.CNTITEM "
        sqlstr &= "         INNER JOIN ARTCU ON ARTCP.CODEPAYM = ARTCU.CODEPAYM AND ARTCP.CNTBTCH = ARTCU.CNTBTCH AND ARTCP.CNTITEM = ARTCU.CNTITEM AND ARTCP.CNTLINE = ARTCU.CNTLINE "
        sqlstr &= " WHERE   (ARBTA.CODEPYMTYP = 'CA' OR ARBTA.CODEPYMTYP = 'AD') AND (ARTCR.RMITTYPE = 1 OR ARTCR.RMITTYPE = 7) AND (ARTCR." & IIf(DATEMODE = "DOCU", "DATERMIT", "DATEBUS") & " >= " & DATEFROM & " AND ARTCR." & IIf(DATEMODE = "DOCU", "DATERMIT", "DATEBUS") & " <= " & DATETO & ") "
        If Use_POST Then
            sqlstr &= " AND (ARBTA.BATCHSTAT = 3 AND ARTCR.ERRENTRY = 0) " 'ARBTA.NBRERRORS = 0
        Else
            sqlstr &= " AND (ARTCR.ERRENTRY = 0) " 'ARBTA.NBRERRORS = 0
        End If
        sqlstr &= " AND IDACCT IN (select ACCTVAT " & IIf(ComDatabase = "PVSW", "", "COLLATE Thai_CI_AS") & " from " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "FMSVLACC)"

        rs.Open(sqlstr, cnACCPAC)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessARRec> Select ARRec Complete")
        End If


        '***************** Set Data To Temp Variable ==> Not Change

        Dim splittxt As String
        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls
                ErrMsg = "Ledger = ARRec IDINVC=" & Trim(rs.options.Rs.Collect(4))
                ErrMsg = ErrMsg & " Batch No.=" & Trim(rs.options.Rs.Collect(16))
                ErrMsg = ErrMsg & " Entry No.=" & Trim(rs.options.Rs.Collect(17))
                tmpTAXBASE = 0
                tmpTAXNAME = ""
                tmpTAXDATE = StrYearF

                If Trim(rs.options.Rs.Collect(16)) = "27" Then
                    sqlstr = ""
                End If

                If InStr(1, Trim(rs.options.Rs.Collect(19)), ",") > 0 Then
                    strsplit = Split(Trim(rs.options.Rs.Collect(19)), ",", -1, CompareMethod.Text)
                    For losplit = LBound(strsplit) To UBound(strsplit)
                        splittxt = Trim(strsplit(losplit))
                        Select Case losplit
                            Case 0
                                tmpTAXBASE = IIf(IsNumeric(splittxt) = True, splittxt, 0)
                            Case 1
                                .DOCNO = Trim(splittxt)
                            Case 2

                                tmpTAXDATE = TryDate2(ZeroDate(splittxt, "dd/MM/yyyy", "yyyyMMdd"), "dd/MM/yyyy", "yyyyMMdd")
                                If tmpTAXDATE = 0 Then
                                    tmpTAXDATE = TryDate2(splittxt, "dd/MM/yyyy", "yyyyMMdd")
                                End If
                            Case 3
                                tmpTAXNAME = Trim(splittxt)
                        End Select
                    Next
                End If


                If tmpTAXDATE = 0 Then
                    .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", (rs.options.Rs.Collect(2)))
                Else
                    .INVDATE = tmpTAXDATE
                End If
                .TXDATE = rs.options.Rs.Collect("DATEBUS")

                .IDINVC = .DOCNO 'IIf(ISNULL(rs.options.Rs, 4), "", Trim(rs.options.Rs.Collect(4)))

                If Len(.DOCNO) = 0 Then
                    .DOCNO = ""
                    .IDINVC = IIf(ISNULL(rs.options.Rs, 4), "", Trim(rs.options.Rs.Collect(4)))
                End If


                .INVAMT = Val(tmpTAXBASE)
                .INVTAX = CDbl(Trim(rs.options.Rs.Collect(9)))


                If .INVTAX < 0 Then
                    .INVTAX = .INVTAX * (-1)
                Else
                    .INVTAX = .INVTAX * (-1)
                    .INVAMT = .INVAMT * (-1)
                End If

                If VNO_AR = True Then

                    Dim VNO_Header As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from ARTCRO  where OPTFIELD in ('ARRVNO') and CNTBTCH ='" & rs.options.Rs.Collect("BATCH") & "' and CNTITEM='" & rs.options.Rs.Collect("ENTRY") & "'", cnACCPAC)
                    For iwht As Integer = 0 To VNO_Header.options.QueryDT.Rows.Count - 1
                        If VNO_Header.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "ARRVNO" Then
                            If Trim(VNO_Header.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                .DOCNO = Trim(VNO_Header.options.QueryDT.Rows(iwht).Item("VALUE"))
                            End If

                        End If
                    Next

                End If


                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect(14)), "", rs.options.Rs.Collect(14))
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(16)), "", (rs.options.Rs.Collect(16))) 'APOBL.CNTBTCH
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(17)), "", (rs.options.Rs.Collect(17))) 'APOBL.CNTITEM
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(21)), "", (rs.options.Rs.Collect(21)))
                .TRANSNBR = IIf(IsDBNull(rs.options.Rs.Collect("CNTSEQ")), "", (rs.options.Rs.Collect("CNTSEQ"))) 'ARTCU.CNTLINE เพิ่มเก็บ Linenumber 12-02-2014 by Pat
                .TTYPE = 2


                If .INVAMT = 0 Then
                    .RATE_Renamed = 0
                Else
                    .RATE_Renamed = FormatNumber(((.INVTAX / .INVAMT) * 100), 1)
                End If


                If FindMiscCodeInARCUS(Trim(tmpTAXNAME)) = "0" Or FindMiscCodeInARCUS(Trim(tmpTAXNAME)) = "" Then
                    If Trim(Trim(tmpTAXNAME)) = "" Then
                        .INVNAME = FindMiscCodeInARCUS(Trim(rs.options.Rs.Collect("IDCUST")))
                    Else
                        '.INVNAME = FindMiscCodeInARCUS(Trim(tmpTAXNAME))
                        .INVNAME = Trim(tmpTAXNAME)
                    End If
                    .TAXID = FindTaxIDBRANCH(Trim(rs.options.Rs.Collect("IDCUST")), "AR", "TAXID")
                    .Branch = FindTaxIDBRANCH(Trim(rs.options.Rs.Collect("IDCUST")), "AR", "BRANCH")
                Else
                    .TAXID = FindTaxIDBRANCH(Trim(tmpTAXNAME), "AR", "TAXID")
                    .Branch = FindTaxIDBRANCH(Trim(tmpTAXNAME), "AR", "BRANCH")
                    .INVNAME = FindMiscCodeInARCUS(Trim(tmpTAXNAME))
                End If
                .VATCOMMENT = .INVAMT & "," & .IDINVC & "," & .INVDATE & "," & .INVNAME
                .MARK = "Receipt"

                sqlstr = InsVat(0, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, "", .INVNAME, .INVAMT, .INVTAX, .LOCID, 1, .RATE_Renamed, .TTYPE, .ACCTVAT, "AR", .Batch, .Entry, .MARK, Trim(.VATCOMMENT), .CBRef, .TRANSNBR, .Runno, IIf(ISNULL(rs.options.Rs, 20), "", Trim(rs.options.Rs.Collect(20))), IIf(ISNULL(rs.options.Rs, 22), "", Trim(rs.options.Rs.Collect(22))), "", "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, 0)
            End With
            cnVAT.Execute(sqlstr)
            ErrMsg = ""
            Code = ""
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing

        rs = New BaseClass.BaseODBCIO
        sqlstr = " SELECT   FMSVATTEMP.INVDATE,FMSVATTEMP.TXDATE,FMSVATTEMP.IDINVC,FMSVATTEMP.DOCNO, " & Environment.NewLine
        sqlstr &= "         FMSVATTEMP.INVNAME,FMSVATTEMP.INVAMT,FMSVATTEMP.INVTAX, FMSVLACC.LOCID," & Environment.NewLine
        sqlstr &= "         FMSVATTEMP.VTYPE,FMSVATTEMP.RATE,FMSVATTEMP.TTYPE, FMSVATTEMP.ACCTVAT,FMSVATTEMP.SOURCE," & Environment.NewLine
        sqlstr &= "         FMSVATTEMP.BATCH,FMSVATTEMP.ENTRY,FMSVATTEMP.TRANSNBR,FMSVATTEMP.MARK,FMSVATTEMP.VATCOMMENT," & Environment.NewLine
        sqlstr &= "         FMSVATTEMP.IDDIST,FMSVATTEMP.CBREF,FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH,FMSVATTEMP.Code,FMSVATTEMP.CODETAX " & Environment.NewLine

        If ComDatabase = "ORCL" Then
            sqlstr &= " FROM FMSVATTEMP "
            sqlstr &= "     INNER JOIN FMSVLACC ON FMSVATTEMP.ACCTVAT = FMSVLACC.ACCTVAT"
            sqlstr &= " WHERE FMSVATTEMP.TTYPE = 2 "
        Else
            sqlstr &= " FROM FMSVATTEMP " & Environment.NewLine
            sqlstr &= "     INNER JOIN FMSVLACC ON FMSVATTEMP.ACCTVAT = FMSVLACC.ACCTVAT " & Environment.NewLine
            sqlstr &= " WHERE FMSVATTEMP.TTYPE = 2 " & Environment.NewLine
        End If

        rs.Open(sqlstr, cnVAT)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessARRec> Select FMSVATTEMP Complete")
        End If


        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls
                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect("INVDATE")), 0, rs.options.Rs.Collect("INVDATE"))
                .TXDATE = IIf(IsDBNull(rs.options.Rs.Collect("TXDATE")), 0, rs.options.Rs.Collect("TXDATE"))
                .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect("IDINVC")), "", (rs.options.Rs.Collect("IDINVC")))
                .DOCNO = IIf(IsDBNull(rs.options.Rs.Collect("DOCNO")), "", (rs.options.Rs.Collect("DOCNO")))
                .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect("INVNAME")), "", (rs.options.Rs.Collect("INVNAME")))
                .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect("INVAMT")), 0, rs.options.Rs.Collect("INVAMT"))
                .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect("INVTAX")), 0, rs.options.Rs.Collect("INVTAX"))
                .LOCID = IIf(IsDBNull(rs.options.Rs.Collect("LOCID")), 0, rs.options.Rs.Collect("LOCID"))
                .VTYPE = IIf(IsDBNull(rs.options.Rs.Collect("VTYPE")), 0, rs.options.Rs.Collect("VTYPE"))
                .TTYPE = IIf(IsDBNull(rs.options.Rs.Collect("TTYPE")), 0, rs.options.Rs.Collect("TTYPE"))
                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect("ACCTVAT")), "", (rs.options.Rs.Collect("ACCTVAT")))
                .Source = IIf(IsDBNull(rs.options.Rs.Collect("Source")), "", (rs.options.Rs.Collect("Source")))
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect("Batch")), "", (rs.options.Rs.Collect("Batch")))
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect("Entry")), "", (rs.options.Rs.Collect("Entry")))
                .MARK = IIf(IsDBNull(rs.options.Rs.Collect("MARK")), "", (rs.options.Rs.Collect("MARK")))
                .VATCOMMENT = IIf(IsDBNull(rs.options.Rs.Collect("VATCOMMENT")), "", (rs.options.Rs.Collect("VATCOMMENT")))
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect("CBRef")), "", (rs.options.Rs.Collect("CBRef")))
                .TRANSNBR = IIf(IsDBNull(rs.options.Rs.Collect("TRANSNBR")), "", (rs.options.Rs.Collect("TRANSNBR")))
                .TAXID = IIf(IsDBNull(rs.options.Rs.Collect("TAXID")), "", (rs.options.Rs.Collect("TAXID")))
                .Branch = IIf(IsDBNull(rs.options.Rs.Collect("BRANCH")), "", (rs.options.Rs.Collect("BRANCH")))
                Code = IIf(IsDBNull(rs.options.Rs.Collect("Code")), "", (rs.options.Rs.Collect("Code")))
                .RATE_Renamed = IIf(IsDBNull(rs.options.Rs.Collect("RATE")), 0, rs.options.Rs.Collect("RATE"))
                .CODETAX = IIf(IsDBNull(rs.options.Rs.Collect("CODETAX")), "", (rs.options.Rs.Collect("CODETAX")))

                sqlstr = InsVat(1, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, .NEWDOCNO, .INVNAME, .INVAMT, .INVTAX, .LOCID, .VTYPE, .RATE_Renamed, .TTYPE, .ACCTVAT, .Source, .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, .Runno, .CODETAX, "", "", "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, 0)

            End With
            cnVAT.Execute(sqlstr)
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing
        cnACCPAC.Close()
        cnVAT.Close()
        cnACCPAC = Nothing
        cnVAT = Nothing
        Exit Sub

ErrHandler:
        WriteLog("<ProcessARRec>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessARRec> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
        WriteLog(ErrMsg)
    End Sub
    Public Sub ProcessAPMisc()
        Dim cnACCPAC, cnVAT As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim strsplit() As String
        Dim losplit As Integer
        Dim loCls As clsVat
        Dim tmpTAXBASE, tmpTAXDATE As String
        Dim tmpTAXNAME As String
        Dim iffs As clsFuntion
        iffs = New clsFuntion
        ErrMsg = ""

        On Error GoTo ErrHandler

        cnACCPAC = New ADODB.Connection
        cnVAT = New ADODB.Connection

        cnACCPAC.ConnectionTimeout = 60
        cnACCPAC.Open(ConnACCPAC)
        cnACCPAC.CommandTimeout = 3600

        cnVAT.ConnectionTimeout = 60
        cnVAT.Open(ConnVAT)
        cnVAT.CommandTimeout = 3600

        'gather data ==> Select Data Uses For Module
        rs = New BaseClass.BaseODBCIO
        sqlstr = " SELECT   APBTA.PAYMTYPE, APBTA.CNTBTCH, APTCR.DATERMIT AS INVDATE, '' AS TXDATE, APTCR.IDINVCMTCH AS IDINVC,'' AS DOCNO, '' AS NEWDOCNO, '' AS INVNAME, '' AS INVAMT, "
        sqlstr &= "         APTCN.AMTNETHC AS INVTAX, APTCR.AUDTORG AS LOCID, '1' AS VTYPE, '' AS RATE, '2' AS TTTYPE, APTCN.IDACCT AS ACCTVAT,"
        sqlstr &= "         APBTA.SRCEAPPL AS SOURCE, APTCR.CNTBTCH AS BATCH, APTCR.CNTENTR AS ENTRY, '' AS MARK, APTCN.GLREF AS VATCOMMENT,"
        sqlstr &= "         APTCN.IDDISTCODE AS CODETAX, APTCR.DOCNBR  AS CBREF,APTCR.IDVEND,APTCN.GLDESC,APTCR.DATEBUS,APTCN.CNTLINE"
        sqlstr &= " FROM    APBTA "
        sqlstr &= "         INNER JOIN APTCR ON APBTA.CNTBTCH = APTCR.CNTBTCH AND APBTA.PAYMTYPE = APTCR.BTCHTYPE "
        sqlstr &= "         INNER JOIN APTCN ON APTCR.BTCHTYPE = APTCN.BATCHTYPE AND APTCR.CNTBTCH = APTCN.CNTBTCH AND APTCR.CNTENTR = APTCN.CNTRMIT"
        sqlstr &= " WHERE   (APBTA.PAYMTYPE = 'PY') AND (APTCR." & IIf(DATEMODE = "DOCU", "DATERMIT", "DATEBUS") & "  >= '" & DATEFROM & "' AND APTCR." & IIf(DATEMODE = "DOCU", "DATERMIT", "DATEBUS") & "  <= '" & DATETO & "') AND (APTCR.RMITTYPE = 4) "
        If Use_POST Then
            sqlstr &= " AND (APBTA.BATCHSTAT = 3 AND APTCR.ERRENTRY = 0)"  'APBTA.NBRERRORS=0
        Else
            sqlstr &= " AND (APTCR.ERRENTRY = 0) "  'APBTA.NBRERRORS=0
        End If
        sqlstr &= " AND IDACCT IN (select ACCTVAT " & IIf(ComDatabase = "PVSW", "", "COLLATE Thai_CI_AS") & " from " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "FMSVLACC)"


        rs.Open(sqlstr, cnACCPAC)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessAPMisc> Select APMisc Complete")
        End If


        '***************** Set Data To Temp Variable ==> Not Change
        Dim splittxt As String
        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls
                ErrMsg = "Ledger = APMisc IDINVC=" & Trim(rs.options.Rs.Collect(4))
                ErrMsg = ErrMsg & " Batch No.=" & Trim(rs.options.Rs.Collect(16))
                ErrMsg = ErrMsg & " Entry No.=" & Trim(rs.options.Rs.Collect(17))

                tmpTAXBASE = 0
                tmpTAXNAME = ""
                tmpTAXDATE = StrYearF

                If InStr(1, Trim(rs.options.Rs.Collect(19)), ",") > 0 Then
                    strsplit = Split(Trim(rs.options.Rs.Collect(19)), ",", -1, CompareMethod.Text)
                    For losplit = LBound(strsplit) To UBound(strsplit)
                        splittxt = Trim(strsplit(losplit))
                        Select Case losplit
                            Case 0
                                tmpTAXBASE = IIf(IsNumeric(splittxt) = True, splittxt, 0)
                            Case 1
                                .IDINVC = Trim(splittxt)
                            Case 2
                                tmpTAXDATE = TryDate2(ZeroDate(splittxt, "dd/MM/yyyy", "yyyyMMdd"), "dd/MM/yyyy", "yyyyMMdd")
                                If tmpTAXDATE = 0 Then
                                    tmpTAXDATE = TryDate2(splittxt, "dd/MM/yyyy", "yyyyMMdd")
                                End If
                            Case 3
                                tmpTAXNAME = Trim(splittxt)
                            Case Is > 3
                                tmpTAXNAME = tmpTAXNAME & "," & Trim(splittxt)
                        End Select
                    Next
                    If strsplit.Length > 3 And tmpTAXNAME.Length > 0 Then
                        tmpTAXNAME = tmpTAXNAME.Replace(",", "")
                    End If
                End If


                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", (rs.options.Rs.Collect(2)))
                .TXDATE = rs.options.Rs.Collect("DATEBUS")

                If Val(tmpTAXDATE) <> 0 Then
                    If IsNumeric(tmpTAXDATE) Then
                        .INVDATE = CDbl(Trim(tmpTAXDATE))
                    End If
                End If

                If Len(.IDINVC) = 0 Then .IDINVC = ""
                If Len(.DOCNO) = 0 Then .DOCNO = ""

                .INVAMT = Val(tmpTAXBASE)
                .INVTAX = CDbl(Trim(rs.options.Rs.Collect(9)))

                If .INVTAX < 0 Then
                    .INVAMT = .INVAMT * (-1)
                End If


                If VNO_AP = True Then

                    Dim VNO_Header As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from APTCRO  where OPTFIELD in ('APPYNO') and CNTBTCH ='" & rs.options.Rs.Collect("BATCH") & "' and CNTITEM='" & rs.options.Rs.Collect("ENTRY") & "'", cnACCPAC)
                    For iwht As Integer = 0 To VNO_Header.options.QueryDT.Rows.Count - 1
                        If VNO_Header.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "APPYNO" Then
                            If Trim(VNO_Header.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                .DOCNO = Trim(VNO_Header.options.QueryDT.Rows(iwht).Item("VALUE"))
                            End If

                        End If
                    Next

                End If

                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect(14)), "", rs.options.Rs.Collect(14))
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(16)), "", (rs.options.Rs.Collect(16))) 'APOBL.CNTBTCH
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(17)), "", (rs.options.Rs.Collect(17))) 'APOBL.CNTITEM
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(21)), "", (rs.options.Rs.Collect(21)))
                .TypeOfPU = IIf(IsDBNull(rs.options.Rs.Collect("GLDESC")), "", (rs.options.Rs.Collect("GLDESC")))
                .TTYPE = 1
                .TRANSNBR = IIf(IsDBNull(rs.options.Rs.Collect("CNTLINE")), "", (rs.options.Rs.Collect("CNTLINE"))) 'APTCN.CNTLINE เพิ่มเก็บ Linenumber 12-02-2014 by Pat


                If .INVAMT = 0 Then
                    .RATE_Renamed = 0
                Else
                    .RATE_Renamed = FormatNumber(((.INVTAX / .INVAMT) * 100), 1)
                End If

                If FindMiscCodeInAPVEN(Trim(tmpTAXNAME), "AP") = "0" Or FindMiscCodeInAPVEN(Trim(tmpTAXNAME), "AP") = "" Then
                    If Trim(tmpTAXNAME) = "" Then
                        .INVNAME = FindMiscCodeInAPVEN(Trim(rs.options.Rs.Collect("IDVEND")), "AP")
                    Else
                        .INVNAME = Trim(tmpTAXNAME)
                    End If
                    .TAXID = FindTaxIDBRANCH(Trim(rs.options.Rs.Collect("IDVEND")), "AP", "TAXID")
                    .Branch = FindTaxIDBRANCH(Trim(rs.options.Rs.Collect("IDVEND")), "AP", "BRANCH")
                Else
                    .TAXID = FindTaxIDBRANCH(tmpTAXNAME, "AP", "TAXID")
                    .Branch = FindTaxIDBRANCH(tmpTAXNAME, "AP", "BRANCH")
                    .INVNAME = FindMiscCodeInAPVEN(Trim(tmpTAXNAME), "AP")
                End If

                .VATCOMMENT = .INVAMT & "," & .IDINVC & "," & .INVDATE & "," & .INVNAME
                .MARK = "Payment"

                sqlstr = InsVat(0, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, "", (.INVNAME), .INVAMT, .INVTAX, .LOCID, 1, .RATE_Renamed, .TTYPE, .ACCTVAT, "AP", .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, .Runno, IIf(ISNULL(rs.options.Rs, 20), "", Trim(rs.options.Rs.Collect(20))), IIf(ISNULL(rs.options.Rs, 22), "", Trim(rs.options.Rs.Collect(22))), .TypeOfPU, "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, 0)

            End With

            cnVAT.Execute(sqlstr)
            ErrMsg = ""
            Code = ""
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop

        rs.options.Rs.Close()
        rs = Nothing

        'transfer data
        rs = New BaseClass.BaseODBCIO
        If ComDatabase = "ORCL" Then
            sqlstr = "SELECT FMSVATTEMP.INVDATE,FMSVATTEMP.TXDATE,FMSVATTEMP.IDINVC,FMSVATTEMP.DOCNO," '0-3
            sqlstr = sqlstr & " FMSVATTEMP.INVNAME,FMSVATTEMP.INVAMT,FMSVATTEMP.INVTAX," '4-6
            sqlstr = sqlstr & " FMSVLACC.LOCID,FMSVATTEMP.VTYPE,FMSVATTEMP.RATE,FMSTAX.TTYPE," '7-10
            sqlstr = sqlstr & " FMSTAX.ACCTVAT,FMSVATTEMP.SOURCE,FMSVATTEMP.BATCH,FMSVATTEMP.ENTRY,FMSVATTEMP.TRANSNBR," '11-14
            sqlstr = sqlstr & " FMSVATTEMP.MARK,FMSVATTEMP.VATCOMMENT,FMSTAX.ITEMRATE1,FMSVATTEMP.IDDIST,FMSVATTEMP.CBREF,FMSVATTEMP.typeofpu,FMSVATTEMP.tranno,FMSVATTEMP.cif,FMSVATTEMP.taxcif, " '15-19
            sqlstr = sqlstr & " FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH,FMSVATTEMP.Code,FMSVATTEMP.CODETAX "
            sqlstr = sqlstr & " FROM FMSVATTEMP,FMSTAX,FMSVLACC WHERE "
            sqlstr = sqlstr & " (FMSVATTEMP.CODETAX = FMSTAX.AUTHORITY AND  FMSTAX.ACCTVAT= FMSVATTEMP.ACCTVAT)"
            sqlstr = sqlstr & " AND FMSTAX.ACCTVAT = FMSVLACC.ACCTVAT"
            sqlstr = sqlstr & " AND (FMSTAX.TTYPE = 1 or FMSTAX.TTYPE = 3) AND (FMSTAX.BUYERCLASS = 1)"
        Else

            sqlstr = " SELECT	FMSVATTEMP.INVDATE, FMSVATTEMP.TXDATE, FMSVATTEMP.IDINVC, FMSVATTEMP.DOCNO, FMSVATTEMP.INVNAME, SUM(FMSVATTEMP.INVAMT) AS INVAMT, " & Environment.NewLine
            sqlstr &= "			SUM(FMSVATTEMP.INVTAX) AS INVTAX,FMSVLACC.LOCID, FMSVATTEMP.VTYPE, FMSVATTEMP.RATE,FMSVATTEMP.TTYPE, " & Environment.NewLine
            sqlstr &= "			FMSVLACC.ACCTVAT ,FMSVATTEMP.SOURCE, FMSVATTEMP.BATCH, FMSVATTEMP.ENTRY,FMSVATTEMP.TRANSNBR, FMSVATTEMP.MARK," & Environment.NewLine
            sqlstr &= "			FMSVATTEMP.VATCOMMENT,FMSVATTEMP.IDDIST , FMSVATTEMP.CBREF,FMSVATTEMP.typeofpu,FMSVATTEMP.tranno, " & Environment.NewLine
            sqlstr &= "			FMSVATTEMP.cif,FMSVATTEMP.taxcif,FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH,FMSVATTEMP.Code,FMSVATTEMP.CODETAX " & Environment.NewLine
            sqlstr &= " FROM	FMSVATTEMP " & Environment.NewLine
            sqlstr &= "			INNER JOIN FMSVLACC ON FMSVATTEMP.ACCTVAT = FMSVLACC.ACCTVAT " & Environment.NewLine
            sqlstr &= " GROUP BY FMSVATTEMP.INVDATE, FMSVATTEMP.TXDATE, FMSVATTEMP.IDINVC, FMSVATTEMP.DOCNO, FMSVATTEMP.INVNAME, FMSVLACC.LOCID, " & Environment.NewLine
            sqlstr &= "			FMSVATTEMP.VTYPE, FMSVATTEMP.RATE,FMSVATTEMP.TTYPE,FMSVLACC.ACCTVAT ,FMSVATTEMP.SOURCE, " & Environment.NewLine
            sqlstr &= "			FMSVATTEMP.BATCH, FMSVATTEMP.ENTRY,FMSVATTEMP.TRANSNBR, FMSVATTEMP.MARK , FMSVATTEMP.VATCOMMENT,  " & Environment.NewLine
            sqlstr &= "			FMSVATTEMP.IDDIST, FMSVATTEMP.CBRef ,FMSVATTEMP.CODETAX ,FMSVATTEMP.typeofpu,FMSVATTEMP.tranno," & Environment.NewLine
            sqlstr &= "			FMSVATTEMP.cif,FMSVATTEMP.taxcif,FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH,FMSVATTEMP.Code,FMSVATTEMP.CODETAX "
            sqlstr &= " HAVING  (FMSVATTEMP.TTYPE = 1 OR FMSVATTEMP.TTYPE = 3) " & Environment.NewLine

        End If
        rs.Open(sqlstr, cnVAT)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessAPMisc> Select FMSVATTEMP Complete")
        End If


        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls
                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect("INVDATE")), 0, rs.options.Rs.Collect("INVDATE"))
                .TXDATE = IIf(IsDBNull(rs.options.Rs.Collect("TXDATE")), 0, rs.options.Rs.Collect("TXDATE"))
                .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect("IDINVC")), "", (rs.options.Rs.Collect("IDINVC")))
                .DOCNO = IIf(IsDBNull(rs.options.Rs.Collect("DOCNO")), "", (rs.options.Rs.Collect("DOCNO")))
                .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect("INVNAME")), "", (rs.options.Rs.Collect("INVNAME")))
                .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect("INVAMT")), 0, rs.options.Rs.Collect("INVAMT"))
                .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect("INVTAX")), 0, rs.options.Rs.Collect("INVTAX"))
                .LOCID = IIf(IsDBNull(rs.options.Rs.Collect("LOCID")), 0, rs.options.Rs.Collect("LOCID"))
                .VTYPE = IIf(IsDBNull(rs.options.Rs.Collect("VTYPE")), 0, rs.options.Rs.Collect("VTYPE"))
                .TTYPE = IIf(IsDBNull(rs.options.Rs.Collect("TTYPE")), 0, rs.options.Rs.Collect("TTYPE"))
                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect("ACCTVAT")), "", (rs.options.Rs.Collect("ACCTVAT")))
                .Source = IIf(IsDBNull(rs.options.Rs.Collect("SOURCE")), "", (rs.options.Rs.Collect("SOURCE")))
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect("BATCH")), "", (rs.options.Rs.Collect("BATCH")))
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect("ENTRY")), "", (rs.options.Rs.Collect("ENTRY")))
                .MARK = IIf(IsDBNull(rs.options.Rs.Collect("MARK")), "", (rs.options.Rs.Collect("MARK")))
                .VATCOMMENT = IIf(IsDBNull(rs.options.Rs.Collect("VATCOMMENT")), "", (rs.options.Rs.Collect("VATCOMMENT")))
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect("CBREF")), "", (rs.options.Rs.Collect("CBREF")))
                .TRANSNBR = IIf(IsDBNull(rs.options.Rs.Collect("TRANSNBR")), "", (rs.options.Rs.Collect("TRANSNBR")))
                .TypeOfPU = IIf(IsDBNull(rs.options.Rs.Collect("TypeofPU")), "", (rs.options.Rs.Collect("TypeofPU")))
                .TranNo = IIf(IsDBNull(rs.options.Rs.Collect("TranNo")), "", (rs.options.Rs.Collect("TranNo")))
                .CIF = IIf(IsDBNull(rs.options.Rs.Collect("CIF")), "0", (rs.options.Rs.Collect("CIF")))
                .TaxCIF = IIf(IsDBNull(rs.options.Rs.Collect("TaxCIF")), "0", (rs.options.Rs.Collect("TaxCIF")))
                .TAXID = IIf(IsDBNull(rs.options.Rs.Collect("TAXID")), "", (rs.options.Rs.Collect("TAXID")))
                .Branch = IIf(IsDBNull(rs.options.Rs.Collect("BRANCH")), "", (rs.options.Rs.Collect("BRANCH")))
                Code = IIf(IsDBNull(rs.options.Rs.Collect("Code")), "", (rs.options.Rs.Collect("Code")))
                .RATE_Renamed = IIf(IsDBNull(rs.options.Rs.Collect("RATE")), 0, (rs.options.Rs.Collect("RATE")))
                .CODETAX = IIf(IsDBNull(rs.options.Rs.Collect("CODETAX")), "", (rs.options.Rs.Collect("CODETAX")))


                sqlstr = InsVat(1, CStr(.INVDATE), CStr(.TXDATE), Trim(.IDINVC), Trim(.DOCNO), Trim(.NEWDOCNO), Trim(.INVNAME), .INVAMT, .INVTAX, Trim(.LOCID), Trim(.VTYPE), .RATE_Renamed, Trim(.TTYPE), Trim(.ACCTVAT), Trim(.Source), Trim(.Batch), Trim(.Entry), Trim(.MARK), Trim(.VATCOMMENT), Trim(.CBRef), .TRANSNBR, .Runno, .CODETAX, "", .TypeOfPU, .TranNo, .CIF, .TaxCIF, 0, .TAXID, .Branch, Code, "", 0, 0, 0)

            End With
            cnVAT.Execute(sqlstr)
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing
        cnACCPAC.Close()
        cnVAT.Close()
        cnACCPAC = Nothing
        cnVAT = Nothing
        Exit Sub

ErrHandler:
        WriteLog("<ProcessAPMisc>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessAPMisc> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
        WriteLog(ErrMsg)
    End Sub
    Public Sub ProcessARMisc()
        Dim cnACCPAC, cnVAT As ADODB.Connection
        Dim rs, rs3 As BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim strsplit() As String
        Dim losplit As Integer
        Dim loCls As clsVat
        Dim tmpTAXBASE, tmpTAXDATE As String
        Dim tmpTAXNAME As String
        Dim iffs As clsFuntion
        iffs = New clsFuntion
        ErrMsg = ""


        Dim VendorCode As String

        On Error GoTo ErrHandler
        cnACCPAC = New ADODB.Connection
        cnVAT = New ADODB.Connection

        cnACCPAC.ConnectionTimeout = 60
        cnACCPAC.Open(ConnACCPAC)
        cnACCPAC.CommandTimeout = 3600

        cnVAT.ConnectionTimeout = 60
        cnVAT.Open(ConnVAT)
        cnVAT.CommandTimeout = 3600

        'gather data ==> Select Data Uses For Module
        rs3 = New BaseClass.BaseODBCIO
        sqlstr = "SELECT HOMECUR FROM CSCOM "
        rs3.Open(sqlstr, cnACCPAC)

        rs = New BaseClass.BaseODBCIO
        sqlstr = "SELECT    ARBTA.CODEPYMTYP, ARBTA.CNTBTCH, ARTCR.DATERMIT AS INVDATE, '' AS TXDATE, ARTCR.IDINVCMTCH AS IDINVC,'' AS DOCNO, '' AS NEWDOCNO, '' AS INVNAME, ARTCN.AMTDISTTC AS  INVAMT, " & vbcrlf
        sqlstr &= "         ARTCN.TXTOTTC AS INVTAX, ARTCR.AUDTORG AS LOCID, '1' AS VTYPE, '' AS RATE, '1' AS TTTYPE" & vbcrlf
        ' sqlstr &= "         ,ARTCN.IDACCT AS ACCTVAT" & vbcrlf
        sqlstr &= "         ,CASE "
        sqlstr &= "             WHEN ARTCN.IDACCT = (select ACCTVAT from [PITVAT].[dbo].[FMSVLACC] WHERE VTYPE = 2) THEN ARTCN.IDACCT"
        sqlstr &= "             ELSE TG.LIABILITY"
        sqlstr &= "         END AS ACCTVAT"
        sqlstr &= "         ,ARBTA.SRCEAPPL AS SOURCE, ARTCR.CNTBTCH AS BATCH, ARTCR.CNTITEM AS ENTRY, '' AS MARK, ARTCN.GLREF AS VATCOMMENT," & vbcrlf
        sqlstr &= "         ARTCN.IDDISTCODE AS CODETAX, ARTCR.DOCNBR  AS CBREF,ARTCR.IDCUST,ARTCR.DATEBUS,ARTCN.CNTLINE " & vbcrlf
        sqlstr &= " FROM    ARBTA " & vbcrlf
        sqlstr &= "         INNER JOIN ARTCR ON ARBTA.CNTBTCH = ARTCR.CNTBTCH AND ARBTA.CODEPYMTYP = ARTCR.CODEPYMTYP " & vbcrlf
        sqlstr &= "         INNER JOIN ARTCN ON ARTCR.CODEPYMTYP = ARTCN.CODEPAYM AND ARTCR.CNTBTCH = ARTCN.CNTBTCH AND ARTCR.CNTITEM = ARTCN.CNTITEM" & vbcrlf
        sqlstr &= "         LEFT JOIN TXAUTH TG ON ARTCR.CODETAXGRP = TG.AUTHORITY " & vbcrlf
        sqlstr &= " WHERE   (ARBTA.CODEPYMTYP = 'CA') AND (ARTCR." & IIf(DATEMODE = "DOCU", "DATERMIT", "DATEBUS") & " >= '" & DATEFROM & "' AND ARTCR." & IIf(DATEMODE = "DOCU", "DATERMIT", "DATEBUS") & " <= '" & DATETO & "') AND (ARTCR.RMITTYPE = 5)" & vbcrlf
        If Use_POST Then
            sqlstr &= " AND (ARBTA.BATCHSTAT = 3 AND ARTCR.ERRENTRY = 0) " & vbcrlf 'ARBTA.NBRERRORS=0
        Else
            sqlstr &= " AND (ARTCR.ERRENTRY = 0) " & vbcrlf 'ARBTA.NBRERRORS=0
        End If

        rs.Open(sqlstr, cnACCPAC)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessARMisc> Select ARMisc Complete")
        End If


        '***************** Set Data To Temp Variable ==> Not Change
        Dim tableau_a_remplir() As String
        Dim nombre_elements, rep, i As Short
        Dim splittxt As String
        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls
                ErrMsg = "Ledger = ARMisc IDINVC=" & Trim(rs.options.Rs.Collect(4))
                ErrMsg = ErrMsg & " Batch No.=" & Trim(rs.options.Rs.Collect(16))
                ErrMsg = ErrMsg & " Entry No.=" & Trim(rs.options.Rs.Collect(17))

                tmpTAXBASE = 0
                tmpTAXNAME = ""
                tmpTAXDATE = StrYearF

                If InStr(1, Trim(rs.options.Rs.Collect(19)), ",") > 0 Then
                    strsplit = Split(Trim(rs.options.Rs.Collect(19)), ",", -1, CompareMethod.Text)
                    For losplit = LBound(strsplit) To UBound(strsplit)
                        splittxt = Trim(strsplit(losplit))
                        Select Case losplit
                            Case 0
                                tmpTAXBASE = IIf(IsNumeric(splittxt) = True, splittxt, 0)
                            Case 1
                                .DOCNO = Trim(splittxt)
                            Case 2

                                tmpTAXDATE = TryDate2(ZeroDate(splittxt, "dd/MM/yyyy", "yyyyMMdd"), "dd/MM/yyyy", "yyyyMMdd")
                                If tmpTAXDATE = 0 Then
                                    tmpTAXDATE = TryDate2(splittxt, "dd/MM/yyyy", "yyyyMMdd")
                                End If
                            Case 3
                                tmpTAXNAME = Trim(splittxt)
                        End Select
                    Next
                End If


                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", Trim(rs.options.Rs.Collect(2)))
                .TXDATE = rs.options.Rs.Collect("DATEBUS")
                If Val(tmpTAXDATE) <> 0 Then
                    If IsNumeric(tmpTAXDATE) Then
                        .INVDATE = CDbl(Trim(tmpTAXDATE))
                    End If
                End If

                .IDINVC = .DOCNO 'IIf(IsDBNull(rs.options.Rs.Collect(4)), "", (rs.options.Rs.Collect(4)))

                If Len(.DOCNO) = 0 Then
                    .DOCNO = ""
                    .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect(4)), "", (rs.options.Rs.Collect(4)))
                End If


                .INVAMT = Val(tmpTAXBASE)
                .INVTAX = CDbl(Trim(rs.options.Rs.Collect(8)))


                If .INVTAX < 0 Then
                    .INVAMT = .INVAMT * (-1)
                End If


                If VNO_AR = True Then

                    Dim VNO_Header As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from ARTCRO  where OPTFIELD in ('ARRVNO') and CNTBTCH ='" & rs.options.Rs.Collect("BATCH") & "' and CNTITEM='" & rs.options.Rs.Collect("ENTRY") & "'", cnACCPAC)
                    For iwht As Integer = 0 To VNO_Header.options.QueryDT.Rows.Count - 1
                        If VNO_Header.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "ARRVNO" Then
                            If Trim(VNO_Header.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                .DOCNO = Trim(VNO_Header.options.QueryDT.Rows(iwht).Item("VALUE"))
                            End If

                        End If
                    Next

                End If

                .VATCOMMENT = IIf(IsDBNull(rs.options.Rs.Collect(19)), "", rs.options.Rs.Collect(19))
                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect(14)), "", rs.options.Rs.Collect(14))
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(16)), "", (rs.options.Rs.Collect(16))) '
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(17)), "", (rs.options.Rs.Collect(17))) '
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(21)), "", (rs.options.Rs.Collect(21)))
                .TRANSNBR = IIf(IsDBNull(rs.options.Rs.Collect("CNTLINE")), "", (rs.options.Rs.Collect("CNTLINE")))
                .TTYPE = 2



                If .INVAMT = 0 Then
                    .RATE_Renamed = 0
                Else
                    .RATE_Renamed = FormatNumber(((.INVTAX / .INVAMT) * 100), 1)
                End If

                If FindMiscCodeInARCUS(Trim(tmpTAXNAME)) = "0" Or FindMiscCodeInARCUS(Trim(tmpTAXNAME)) = "" Then
                    If Trim(tmpTAXNAME) = "" Then
                        .INVNAME = FindMiscCodeInARCUS(Trim(rs.options.Rs.Collect("IDCUST")))
                    Else
                        .INVNAME = Trim(tmpTAXNAME)
                    End If
                    .TAXID = FindTaxIDBranch(Trim(rs.options.Rs.Collect("IDCUST")), "AR", "TAXID")
                    .Branch = FindTaxIDBranch(Trim(rs.options.Rs.Collect("IDCUST")), "AR", "BRANCH")
                Else
                    .TAXID = FindTaxIDBranch(Trim(tmpTAXNAME), "AR", "TAXID")
                    .Branch = FindTaxIDBranch(Trim(tmpTAXNAME), "AR", "BRANCH")
                    .INVNAME = FindMiscCodeInARCUS(Trim(tmpTAXNAME))
                End If

                .VATCOMMENT = .INVAMT & "," & .IDINVC & "," & .INVDATE & "," & .INVNAME
                .MARK = "Receipt"
                sqlstr = InsVat(0, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, "", .INVNAME, .INVAMT, .INVTAX, .LOCID, 1, .RATE_Renamed, .TTYPE, .ACCTVAT, "AR", .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, .Runno, IIf(ISNULL(rs.options.Rs, 20), "", Trim(rs.options.Rs.Collect(20))), IIf(ISNULL(rs.options.Rs, 22), "", Trim(rs.options.Rs.Collect(22))), "", "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, 0)

            End With
            cnVAT.Execute(sqlstr)
            ErrMsg = ""
            Code = ""
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing

        rs = New BaseClass.BaseODBCIO
        sqlstr = " SELECT   FMSVATTEMP.INVDATE,FMSVATTEMP.TXDATE,FMSVATTEMP.IDINVC,FMSVATTEMP.DOCNO,  " & Environment.NewLine
        sqlstr &= "         FMSVATTEMP.INVNAME,FMSVATTEMP.INVAMT,FMSVATTEMP.INVTAX, FMSVLACC.LOCID,FMSVATTEMP.VTYPE, " & Environment.NewLine
        sqlstr &= "         FMSVATTEMP.RATE,FMSVATTEMP.TTYPE, FMSVATTEMP.ACCTVAT, " & Environment.NewLine
        sqlstr &= "         FMSVATTEMP.SOURCE,FMSVATTEMP.BATCH,FMSVATTEMP.ENTRY,FMSVATTEMP.TRANSNBR, " & Environment.NewLine
        sqlstr &= "         FMSVATTEMP.MARK,FMSVATTEMP.VATCOMMENT,FMSVATTEMP.IDDIST,FMSVATTEMP.CBREF, " & Environment.NewLine
        sqlstr &= "         FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH,FMSVATTEMP.Code,FMSVATTEMP.CODETAX "
        If ComDatabase = "ORCL" Then
            sqlstr &= " FROM FMSVATTEMP"
            sqlstr &= "     INNER JOIN FMSVLACC ON FMSVATTEMP.ACCTVAT = FMSVLACC.ACCTVAT"
            sqlstr &= " WHERE FMSVATTEMP.TTYPE = 2 "

        Else
            sqlstr &= " FROM FMSVATTEMP " & Environment.NewLine
            sqlstr &= "     INNER JOIN FMSVLACC ON FMSVATTEMP.ACCTVAT = FMSVLACC.ACCTVAT  " & Environment.NewLine
            sqlstr &= " WHERE FMSVATTEMP.TTYPE = 2 " & Environment.NewLine
        End If

        rs.Open(sqlstr, cnVAT)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessARMisc> Select FMSVATTEMP Complete")
        End If


        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls
                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect("INVDATE")), 0, rs.options.Rs.Collect("INVDATE"))
                .TXDATE = IIf(IsDBNull(rs.options.Rs.Collect("TXDATE")), 0, rs.options.Rs.Collect("TXDATE"))
                .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect("IDINVC")), "", (rs.options.Rs.Collect("IDINVC")))
                .DOCNO = IIf(IsDBNull(rs.options.Rs.Collect("DOCNO")), "", (rs.options.Rs.Collect("DOCNO")))
                .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect("INVNAME")), "", (rs.options.Rs.Collect("INVNAME")))
                .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect("INVAMT")), 0, rs.options.Rs.Collect("INVAMT"))
                .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect("INVTAX")), 0, rs.options.Rs.Collect("INVTAX"))
                .LOCID = IIf(IsDBNull(rs.options.Rs.Collect("LOCID")), 0, rs.options.Rs.Collect("LOCID"))
                .VTYPE = IIf(IsDBNull(rs.options.Rs.Collect("VTYPE")), 0, rs.options.Rs.Collect("VTYPE"))
                .TTYPE = IIf(IsDBNull(rs.options.Rs.Collect("TTYPE")), 0, rs.options.Rs.Collect("TTYPE"))
                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect("ACCTVAT")), "", (rs.options.Rs.Collect("ACCTVAT")))
                .Source = IIf(IsDBNull(rs.options.Rs.Collect("Source")), "", (rs.options.Rs.Collect("Source")))
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect("Batch")), "", (rs.options.Rs.Collect("Batch")))
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect("Entry")), "", (rs.options.Rs.Collect("Entry")))
                .MARK = IIf(IsDBNull(rs.options.Rs.Collect("MARK")), "", (rs.options.Rs.Collect("MARK")))
                .VATCOMMENT = IIf(IsDBNull(rs.options.Rs.Collect("VATCOMMENT")), "", (rs.options.Rs.Collect("VATCOMMENT")))
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect("CBRef")), "", (rs.options.Rs.Collect("CBRef")))
                .TRANSNBR = IIf(IsDBNull(rs.options.Rs.Collect("TRANSNBR")), "", (rs.options.Rs.Collect("TRANSNBR")))
                .TAXID = IIf(IsDBNull(rs.options.Rs.Collect("TAXID")), "", (rs.options.Rs.Collect("TAXID")))
                .Branch = IIf(IsDBNull(rs.options.Rs.Collect("BRANCH")), "", (rs.options.Rs.Collect("BRANCH")))
                Code = IIf(IsDBNull(rs.options.Rs.Collect("Code")), "", (rs.options.Rs.Collect("Code")))
                .RATE_Renamed = IIf(IsDBNull(rs.options.Rs.Collect("RATE")), 0, rs.options.Rs.Collect("RATE"))
                .CODETAX = IIf(IsDBNull(rs.options.Rs.Collect("CODETAX")), "", (rs.options.Rs.Collect("CODETAX")))


                sqlstr = InsVat(1, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, .NEWDOCNO, .INVNAME, .INVAMT, .INVTAX, .LOCID, .VTYPE, .RATE_Renamed, .TTYPE, .ACCTVAT, .Source, .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, .Runno, "", "", "", "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, 0)

            End With
            cnVAT.Execute(sqlstr)
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing
        cnACCPAC.Close()
        cnVAT.Close()
        cnACCPAC = Nothing
        cnVAT = Nothing
        Exit Sub

ErrHandler:
        WriteLog("<ProcessARMisc>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessARMisc> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
        WriteLog(ErrMsg)
    End Sub

    Public Sub ProcessBK()
        Dim cnACCPAC, cnVAT As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim strsplit() As String
        Dim losplit As Integer
        Dim loCls As clsVat
        Dim tmpTAXBASE, tmpTAXDATE, TmpGL As String
        Dim tmpTAXNAME As String
        Dim iffs As clsFuntion
        iffs = New clsFuntion


        Dim VendorCode As String

        On Error GoTo ErrHandler

        cnACCPAC = New ADODB.Connection
        cnVAT = New ADODB.Connection

        cnACCPAC.ConnectionTimeout = 60
        cnACCPAC.Open(ConnACCPAC)
        cnACCPAC.CommandTimeout = 3600

        cnVAT.ConnectionTimeout = 60
        cnVAT.Open(ConnVAT)
        cnVAT.CommandTimeout = 3600

        'gather data ==> Select Data Uses For Module


        rs = New BaseClass.BaseODBCIO
        sqlstr = ""
        sqlstr = " SELECT PGMVER, SELECTOR  FROM CSAPP"
        sqlstr &= " WHERE (SELECTOR = 'GL')"
        rs.Open(sqlstr, cnACCPAC)
        If rs.options.Rs.EOF = True Then Exit Sub
        TmpGL = Mid(Trim(rs.options.Rs.Collect(0)), 1, 2)

        If Trim(TmpGL) < "56" Then Exit Sub

        rs = New BaseClass.BaseODBCIO
        If Use_POST Then
            sqlstr = " SELECT BKJENTH.PSTSEQ, BKJENTH.SEQUENCENO, BKJENTH.TRANSDATE AS INVDATE, '' AS TXDATE, BKJENTH.ENTRYNBR, '' AS DOCNO, '' AS NEWDOCNO, "
            sqlstr &= "       '' AS INVNAME, '0' AS INVAMT, BKJENTD.AMOUNT AS INVTAX, BKJENTH.AUDTORG AS LOCID, '1' AS VTYPE, '' AS RATE, '1' AS TTTYPE, "
            sqlstr &= "       BKJENTD.GLACCOUNT AS ACCTVAT, 'BK' AS SOURCE, BKJENTH.PSTSEQ AS BATCH, BKJENTH.SEQUENCENO AS ENTRY, '' AS MARK, "
            sqlstr &= "       BKJENTD.BIGCOMMENT AS VATCOMMENT, BKJENTD.DISTCODE AS CODETAX, BKJENTD.REFERENCE AS CBREF, BKJENTD.SRCEAMT,BKJENTH.POSTDATE"
            sqlstr &= " FROM  BKJENTH LEFT OUTER JOIN"
            sqlstr &= "       BKJENTD ON BKJENTH.PSTSEQ = BKJENTD.PSTSEQ AND BKJENTH.SEQUENCENO = BKJENTD.SEQUENCENO"
            sqlstr &= " WHERE (BKJENTH.TRANSDATE >= '" & DATEFROM & "' AND BKJENTH.TRANSDATE <= '" & DATETO & "') AND (BKJENTD.BIGCOMMENT <> '')"
        Else
            sqlstr = " SELECT  BKENTH.PSTSEQ, BKENTH.SEQUENCENO, BKENTH.TRANSDATE AS INVDATE, '' AS TXDATE, BKENTH.ENTRYNBR, '' AS DOCNO, '' AS NEWDOCNO, " & Environment.NewLine
            sqlstr &= "        '' AS INVNAME, '0' AS INVAMT, BKENTD.AMOUNT AS INVTAX, BKENTH.AUDTORG AS LOCID, '1' AS VTYPE, '' AS RATE, '1' AS TTTYPE,  " & Environment.NewLine
            sqlstr &= "        BKENTD.GLACCOUNT AS ACCTVAT, 'BK' AS SOURCE, BKENTH.PSTSEQ AS BATCH, BKENTH.SEQUENCENO AS ENTRY, '' AS MARK,  " & Environment.NewLine
            sqlstr &= "        BKENTD.BIGCOMMENT AS VATCOMMENT, BKENTD.DISTCODE AS CODETAX, BKENTD.REFERENCE AS CBREF, BKENTD.SRCEAMT,BKENTH.POSTDATE " & Environment.NewLine
            sqlstr &= " FROM   BKENTH LEFT OUTER JOIN " & Environment.NewLine
            sqlstr &= "        BKENTD ON BKENTH.LINES  = BKENTD.LINE AND BKENTH.SEQUENCENO = BKENTD.SEQUENCENO " & Environment.NewLine
            sqlstr &= " WHERE  (BKENTH.TRANSDATE >= '" & DATEFROM & "' AND BKENTH.TRANSDATE <= '" & DATETO & "') AND (BKENTD.BIGCOMMENT <> '') " & Environment.NewLine

        End If
        rs.Open(sqlstr, cnACCPAC)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessBK> Select BK Complete")
        End If


        '***************** Set Data To Temp Variable ==> Not Change
        Dim tableau_a_remplir() As String
        Dim nombre_elements, rep, i As Short
        Dim splittxt As String
        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls
                ErrMsg = "Ledger = BK IDINVC=" & Trim(rs.options.Rs.Collect(4))
                ErrMsg = ErrMsg & " Batch No.=" & Trim(rs.options.Rs.Collect(16))
                ErrMsg = ErrMsg & " Entry No.=" & Trim(rs.options.Rs.Collect(17))

                tmpTAXBASE = 0
                tmpTAXNAME = ""
                tmpTAXDATE = StrYearF

                If InStr(1, Trim(rs.options.Rs.Collect(19)), ",") > 0 Then
                    strsplit = Split(Trim(rs.options.Rs.Collect(19)), ",", -1, CompareMethod.Text)
                    For losplit = LBound(strsplit) To UBound(strsplit)
                        splittxt = Trim(strsplit(losplit))
                        Select Case losplit
                            Case 0
                                tmpTAXBASE = IIf(IsNumeric(splittxt) = True, splittxt, 0)
                            Case 1
                                .DOCNO = Trim(splittxt)
                            Case 2
                                tmpTAXDATE = TryDate2(ZeroDate(splittxt, "dd/MM/yyyy", "yyyyMMdd"), "dd/MM/yyyy", "yyyyMMdd")
                                If tmpTAXDATE = 0 Then
                                    tmpTAXDATE = TryDate2(splittxt, "dd/MM/yyyy", "yyyyMMdd")
                                End If
                            Case 3
                                tmpTAXNAME = Trim(splittxt)
                        End Select
                    Next
                End If


                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", Trim(rs.options.Rs.Collect(2)))
                .TXDATE = rs.options.Rs.Collect("POSTDATE")
                If Val(tmpTAXDATE) <> 0 Then
                    If IsNumeric(tmpTAXDATE) Then
                        .INVDATE = CDbl(Trim(tmpTAXDATE))
                    End If
                End If
                .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect(4)), "", (rs.options.Rs.Collect(4)))
                If Len(.DOCNO) = 0 Then .DOCNO = ""

                If FindMiscCodeInARCUS(Trim(tmpTAXNAME)) = "0" Or FindMiscCodeInARCUS(Trim(tmpTAXNAME)) = "" Then
                    .INVNAME = Trim(tmpTAXNAME)
                    .TAXID = ""
                    .Branch = ""
                Else
                    .TAXID = FindTaxIDBranch(Trim(tmpTAXNAME), "AR", "TAXID")
                    .Branch = FindTaxIDBranch(Trim(tmpTAXNAME), "AR", "BRANCH")
                    .INVNAME = FindMiscCodeInARCUS(Trim(tmpTAXNAME))
                End If


                .INVAMT = Val(tmpTAXBASE)
                .INVTAX = CDbl(Trim(rs.options.Rs.Collect(9)))


                If .INVTAX < 0 Then
                    .INVAMT = .INVAMT * (-1)
                End If

                .VATCOMMENT = .INVAMT & "," & .IDINVC & "," & .INVDATE & "," & .INVNAME
                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect(14)), "", rs.options.Rs.Collect(14))
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(16)), "", (rs.options.Rs.Collect(16))) '
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(17)), "", (rs.options.Rs.Collect(17))) '
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(21)), "", (rs.options.Rs.Collect(21)))
                .TTYPE = 3

                If .INVAMT = 0 Then
                    .RATE_Renamed = 0
                Else
                    .RATE_Renamed = FormatNumber(((.INVTAX / .INVAMT) * 100), 1)
                End If

                sqlstr = InsVat(0, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, "", .INVNAME, .INVAMT, .INVTAX, .LOCID, 1, .RATE_Renamed, .TTYPE, .ACCTVAT, "BK", .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, "", .Runno, IIf(ISNULL(rs.options.Rs, 20), "", Trim(rs.options.Rs.Collect(20))), IIf(ISNULL(rs.options.Rs, 22), "", Trim(rs.options.Rs.Collect(22))), "", "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, 0)
            End With
            cnVAT.Execute(sqlstr)
            ErrMsg = ""
            Code = ""
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing

        'transfer data
        rs = New BaseClass.BaseODBCIO
        sqlstr = "SELECT FMSVATTEMP.INVDATE,FMSVATTEMP.TXDATE,FMSVATTEMP.IDINVC,FMSVATTEMP.DOCNO," '0-3
        sqlstr = sqlstr & " FMSVATTEMP.INVNAME,FMSVATTEMP.INVAMT,FMSVATTEMP.INVTAX," '4-6
        sqlstr = sqlstr & " FMSVLACC.LOCID,FMSVATTEMP.VTYPE,FMSVATTEMP.RATE,FMSTAX.TTYPE," '7-10
        sqlstr = sqlstr & " FMSTAX.ACCTVAT,FMSVATTEMP.SOURCE,FMSVATTEMP.BATCH,FMSVATTEMP.ENTRY," '11-14
        sqlstr = sqlstr & " FMSVATTEMP.MARK,FMSVATTEMP.VATCOMMENT,FMSTAX.ITEMRATE1,FMSVATTEMP.IDDIST,FMSVATTEMP.CBREF" '15-19
        sqlstr = sqlstr & ",FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH,FMSVATTEMP.Code "
        If ComDatabase = "ORCL" Then
            sqlstr = sqlstr & " FROM FMSVATTEMP"
            sqlstr = sqlstr & " INNER JOIN FMSTAX ON (FMSVATTEMP.CODETAX = FMSTAX.AUTHORITY) and (FMSVATTEMP.ACCTVAT = FMSTAX.ACCTVAT) "
            sqlstr = sqlstr & " INNER JOIN FMSVLACC ON FMSTAX.ACCTVAT = FMSVLACC.ACCTVAT"
            sqlstr = sqlstr & " WHERE  (FMSTAX.BUYERCLASS = 1) and (FMSTAX.TTYPE = 1 OR     FMSTAX.TTYPE = 2 OR     FMSTAX.TTYPE = 3)"
        Else
            sqlstr = sqlstr & " FROM FMSVATTEMP"
            sqlstr = sqlstr & " INNER JOIN FMSTAX ON (FMSVATTEMP.CODETAX = FMSTAX.AUTHORITY) and (FMSVATTEMP.ACCTVAT = FMSTAX.ACCTVAT)"
            sqlstr = sqlstr & " INNER JOIN FMSVLACC ON FMSTAX.ACCTVAT = FMSVLACC.ACCTVAT"
            sqlstr = sqlstr & " WHERE (FMSTAX.BUYERCLASS = 1) and (FMSTAX.TTYPE = 1 OR     FMSTAX.TTYPE = 2  OR     FMSTAX.TTYPE = 3) "
        End If

        rs.Open(sqlstr, cnVAT)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessBK> Select FMSVATTEMP Complete")
        End If


        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls
                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect(0)), 0, rs.options.Rs.Collect(0))
                .TXDATE = IIf(IsDBNull(rs.options.Rs.Collect(1)), 0, rs.options.Rs.Collect(1))
                .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", (rs.options.Rs.Collect(2)))
                .DOCNO = IIf(IsDBNull(rs.options.Rs.Collect(3)), "", (rs.options.Rs.Collect(3)))
                .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect(4)), "", (rs.options.Rs.Collect(4)))
                .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(5)), 0, rs.options.Rs.Collect(5))
                .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(6)), 0, rs.options.Rs.Collect(6))
                .LOCID = IIf(IsDBNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7))
                .VTYPE = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8))
                .TTYPE = IIf(IsDBNull(rs.options.Rs.Collect(10)), 0, rs.options.Rs.Collect(10))
                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect(11)), "", (rs.options.Rs.Collect(11)))
                .Source = IIf(IsDBNull(rs.options.Rs.Collect(12)), "", (rs.options.Rs.Collect(12)))
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(13)), "", (rs.options.Rs.Collect(13)))
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(14)), "", (rs.options.Rs.Collect(14)))
                .MARK = IIf(IsDBNull(rs.options.Rs.Collect(15)), "", (rs.options.Rs.Collect(15)))
                .VATCOMMENT = IIf(IsDBNull(rs.options.Rs.Collect(16)), "", (rs.options.Rs.Collect(16)))
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(19)), "", (rs.options.Rs.Collect(19)))
                .TAXID = IIf(IsDBNull(rs.options.Rs.Collect("TAXID")), "", (rs.options.Rs.Collect("TAXID")))
                .Branch = IIf(IsDBNull(rs.options.Rs.Collect("BRANCH")), "", (rs.options.Rs.Collect("BRANCH")))
                Code = IIf(IsDBNull(rs.options.Rs.Collect("Code")), "", (rs.options.Rs.Collect("Code")))
                .RATE_Renamed = IIf(IsDBNull(rs.options.Rs.Collect("RATE")), 0, rs.options.Rs.Collect("RATE"))

            

                sqlstr = InsVat(1, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, .NEWDOCNO, .INVNAME, .INVAMT, .INVTAX, .LOCID, .VTYPE, .RATE_Renamed, .TTYPE, .ACCTVAT, .Source, .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, "", .Runno, "", "", "", "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, 0)
            End With
            cnVAT.Execute(sqlstr)
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing
        cnACCPAC.Close()
        cnVAT.Close()
        cnACCPAC = Nothing
        cnVAT = Nothing
        Exit Sub

ErrHandler:
        WriteLog("<ProcessBK>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessBK> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
        WriteLog(ErrMsg)
    End Sub
    Public Sub ProcessARPJHTEMP()
        Dim cn As ADODB.Connection
        Dim sqlstr As String

        On Error GoTo ErrHandler

        cn = New ADODB.Connection

        cn.ConnectionTimeout = 60
        cn.Open(ConnACCPAC)
        cn.CommandTimeout = 600

        sqlstr = "INSERT INTO " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "ARPJHTEMP"
        sqlstr = sqlstr & " (TYPEBTCH,"
        sqlstr = sqlstr & " POSTSEQNCE,"
        sqlstr = sqlstr & " CNTBTCH,"
        sqlstr = sqlstr & " CNTITEM,"
        sqlstr = sqlstr & " IDCUST,"
        sqlstr = sqlstr & " IDINVC,"
        sqlstr = sqlstr & " DATEBUS)"

        sqlstr = sqlstr & "SELECT"
        sqlstr = sqlstr & " TYPEBTCH,"
        sqlstr = sqlstr & " POSTSEQNCE,"
        sqlstr = sqlstr & " CNTBTCH,"
        sqlstr = sqlstr & " CNTITEM,"
        sqlstr = sqlstr & " IDCUST,"
        sqlstr = sqlstr & " IDINVC,"
        sqlstr = sqlstr & " DATEBUS"
        sqlstr = sqlstr & " FROM ARPJH"
        sqlstr = sqlstr & " WHERE (ARPJH.DATEBUS >= " & DATEFROM & " AND ARPJH.DATEBUS <=" & DATETO & ")"
        cn.Execute(sqlstr)
        cn.Close()
        cn = Nothing
        Exit Sub

ErrHandler:
        cn.RollbackTrans()
        ErrMsg = Err.Description
        WriteLog("<ProcessARPJHTEMP>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    Public Sub ProcessAROBLTEMP()
        Dim cn As ADODB.Connection
        Dim sqlstr As String

        On Error GoTo ErrHandler

        cn = New ADODB.Connection

        cn.ConnectionTimeout = 60
        cn.Open(ConnACCPAC)
        cn.CommandTimeout = 600

        sqlstr = "INSERT INTO " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "AROBLTEMP"
        sqlstr = sqlstr & " (IDCUST,"
        sqlstr = sqlstr & " IDINVC,"
        sqlstr = sqlstr & " IDRMIT,"
        sqlstr = sqlstr & " TRXTYPEID,"
        sqlstr = sqlstr & " TRXTYPETXT,"
        sqlstr = sqlstr & " CNTBTCH,"
        sqlstr = sqlstr & " CNTITEM,"
        sqlstr = sqlstr & " DATEINVC,"
        sqlstr = sqlstr & " AMTINVCHC,"
        sqlstr = sqlstr & " CODECURN,"
        sqlstr = sqlstr & " EXCHRATEHC,"
        sqlstr = sqlstr & " AMTTAXHC,"
        sqlstr = sqlstr & " CODETAX1,"
        sqlstr = sqlstr & " DATEBUS,"
        sqlstr = sqlstr & " CODETAXGRP,"

        If ComDatabase = "MSSQL" Then
            sqlstr = sqlstr & " [VALUES])"
        Else 'PVSW
            sqlstr = sqlstr & " " & Chr(34) & "VALUES" & Chr(34) & ")"
        End If

        sqlstr = sqlstr & " SELECT"
        sqlstr = sqlstr & " IDCUST,"
        sqlstr = sqlstr & " IDINVC,"
        sqlstr = sqlstr & " IDRMIT,"
        sqlstr = sqlstr & " TRXTYPEID,"
        sqlstr = sqlstr & " TRXTYPETXT,"
        sqlstr = sqlstr & " CNTBTCH,"
        sqlstr = sqlstr & " CNTITEM,"
        sqlstr = sqlstr & " DATEINVC,"
        sqlstr = sqlstr & " AMTINVCHC,"
        sqlstr = sqlstr & " CODECURN,"
        sqlstr = sqlstr & " EXCHRATEHC,"
        sqlstr = sqlstr & " AMTTAXHC,"
        sqlstr = sqlstr & " CODETAX1,"
        sqlstr = sqlstr & " DATEBUS,"
        sqlstr = sqlstr & " CODETAXGRP,"


        If ComDatabase = "MSSQL" Then
            sqlstr = sqlstr & " [VALUES]"
        Else 'PVSW
            sqlstr = sqlstr & " " & Chr(34) & "VALUES" & Chr(34) & ""
        End If

        sqlstr = sqlstr & " FROM AROBL"
        sqlstr = sqlstr & " WHERE AROBL.TRXTYPETXT IN (1,2,3) AND (AROBL.DATEBUS >= " & DATEFROM & " AND AROBL.DATEBUS <= " & DATETO & ")"
        cn.Execute(sqlstr)
        cn.Close()
        cn = Nothing
        Exit Sub

ErrHandler:
        cn.RollbackTrans()
        ErrMsg = Err.Description
        WriteLog("<ProcessAROBLTEMP>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    Public Sub ProcessOE_Backup()
        Dim cnACCPAC, cnVAT As ADODB.Connection
        Dim rs, rs1, rs2, rs3 As BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim loCls As clsVat
        Dim iffs As clsFuntion
        Dim Ddbo As String
        iffs = New clsFuntion


        On Error GoTo ErrHandler
        cnACCPAC = New ADODB.Connection
        cnVAT = New ADODB.Connection
        cnACCPAC.ConnectionTimeout = 60
        cnACCPAC.Open(ConnACCPAC)
        cnACCPAC.CommandTimeout = 3600

        cnVAT.ConnectionTimeout = 60
        cnVAT.Open(ConnVAT)
        cnVAT.CommandTimeout = 3600

        'gather data
        rs3 = New BaseClass.BaseODBCIO
        sqlstr = "SELECT HOMECUR FROM CSCOM "
        rs3.Open(sqlstr, cnACCPAC)


        Ddbo = ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.")

        rs = New BaseClass.BaseODBCIO

        If Use_POST Then
            sqlstr = " SELECT " & Ddbo & "AROBLTEMP.DATEINVC, " & Ddbo & "AROBLTEMP.IDINVC, " & Ddbo & "AROBLTEMP.TRXTYPEID, ARCUS.NAMECUST, ARCUS.TEXTSTRE1, SUM(ARIBD.AMTEXTN) AS AMTEXTN, SUM(ARIBD.TOTTAX) AS TOTTAX, "
            sqlstr &= "         " & Ddbo & "AROBLTEMP.CNTBTCH, " & Ddbo & "AROBLTEMP.CNTITEM, " & Ddbo & "AROBLTEMP.CODETAX1, OEINVH.BILNAME, OEINVH.BILADDR1, OECRDH.BILNAME AS Expr1,  "
            sqlstr &= "         OECRDH.BILADDR1 AS Expr2, " & Ddbo & "AROBLTEMP.IDRMIT, " & Ddbo & "AROBLTEMP.IDCUST, SUM(ARIBD.BASETAX1) AS BASETAX1, SUM(ARIBD.TOTTAX) AS TOTTAX, ARIBH.TEXTTRX,  "
            sqlstr &= "         ARIBH.CODECURN, " & Ddbo & "AROBLTEMP.CODECURN AS CODECURR, " & Ddbo & "AROBLTEMP.EXCHRATEHC, ARIBHO.VALUE, " & Ddbo & "AROBLTEMP.DATEBUS, "
            sqlstr &= "         '' AS TEXTLINE,0 AS CNTLINE,'HEADER' AS VATAS "
            sqlstr &= " FROM " & Ddbo & "AROBLTEMP"
            sqlstr &= "         LEFT OUTER JOIN ARCUS ON " & Ddbo & "AROBLTEMP.IDCUST = ARCUS.IDCUST "
            sqlstr &= "         LEFT OUTER JOIN OEINVH ON " & Ddbo & "AROBLTEMP.IDINVC = OEINVH.INVNUMBER "
            sqlstr &= "         LEFT OUTER JOIN OECRDH ON " & Ddbo & "AROBLTEMP.IDINVC = OECRDH.CRDNUMBER "
            sqlstr &= "         LEFT OUTER JOIN ARIBH ON " & Ddbo & "AROBLTEMP.CNTBTCH = ARIBH.CNTBTCH AND " & Ddbo & "AROBLTEMP.CNTITEM = ARIBH.CNTITEM AND " & Ddbo & "AROBLTEMP.IDCUST = ARIBH.IDCUST AND " & Ddbo & "AROBLTEMP.IDINVC = ARIBH.IDINVC "
            sqlstr &= "         LEFT OUTER JOIN ARIBD ON ARIBH.CNTBTCH = ARIBD.CNTBTCH AND ARIBH.CNTITEM = ARIBD.CNTITEM "
            sqlstr &= "         LEFT OUTER JOIN ARIBHO ON ARIBHO.CNTBTCH = ARIBH.CNTBTCH AND ARIBHO.CNTITEM = ARIBD.CNTITEM AND ARIBHO.OPTFIELD = 'VOUCHER'"
            sqlstr &= "         LEFT OUTER JOIN TXAUTH ON " & Ddbo & "AROBLTEMP.CODETAXGRP = TXAUTH.AUTHORITY "
            sqlstr &= " WHERE   (" & Ddbo & "AROBLTEMP.TRXTYPETXT IN (1,2,3)) AND (" & Ddbo & "AROBLTEMP.DATEBUS >= " & DATEFROM & " AND " & Ddbo & "AROBLTEMP.DATEBUS <= " & DATETO & ") AND (ARIBD.TAXSTTS1 = 1) "
            sqlstr &= "         AND ((ARIBD.BASETAX1 <> 0) OR (ARIBD.TOTTAX <> 0)) "
            'Tax Amount<>0
            ' หาก Tax Amount =0 จะไม่เเสดงรายการนั้น
            'sqlstr &= " and " & Ddbo & "AROBLTEMP.AMTTAXHC<>0  " commentไว้ Vat 0 ก็ออกตามปกติ
            'TXAUTH.ACCTRECOV แก้ไขเนื่องจาก account ขาย อยู่ที่ TXAUTH.LIABILITY  ,ซื้อ อยู่ที่  TXAUTH.ACCTRECOV
            sqlstr &= "         AND  TXAUTH.LIABILITY IN (SELECT ACCTVAT " & IIf(ComDatabase = "PVSW", "", "COLLATE Thai_CI_AS") & " FROM " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "FMSVLACC)   "
            sqlstr &= " GROUP BY " & Ddbo & "AROBLTEMP.DATEINVC, " & Ddbo & "AROBLTEMP.IDINVC, " & Ddbo & "AROBLTEMP.TRXTYPEID, ARCUS.NAMECUST, ARCUS.TEXTSTRE1, " & Ddbo & "AROBLTEMP.AMTINVCHC, " & Ddbo & "AROBLTEMP.AMTTAXHC, " & Ddbo & "AROBLTEMP.CNTBTCH,  "
            sqlstr &= "         " & Ddbo & "AROBLTEMP.CNTITEM, " & Ddbo & "AROBLTEMP.CODETAX1, OEINVH.BILNAME, OEINVH.BILADDR1, OECRDH.BILNAME, OECRDH.BILADDR1, " & Ddbo & "AROBLTEMP.IDRMIT, " & Ddbo & "AROBLTEMP.IDCUST,  "
            sqlstr &= "         ARIBH.TEXTTRX, ARIBH.CODECURN, " & Ddbo & "AROBLTEMP.CODETAX1, " & Ddbo & "AROBLTEMP.CODECURN, " & Ddbo & "AROBLTEMP.EXCHRATEHC, ARIBHO.VALUE, " & Ddbo & "AROBLTEMP.DATEBUS "

            sqlstr &= " UNION ALL "

            sqlstr &= " SELECT " & Ddbo & "AROBLTEMP.DATEINVC, " & Ddbo & "AROBLTEMP.IDINVC, " & Ddbo & "AROBLTEMP.TRXTYPEID, ARCUS.NAMECUST, ARCUS.TEXTSTRE1, SUM(ARIBD.AMTEXTN) AS AMTEXTN, SUM(ARIBD.TOTTAX) AS TOTTAX, "
            sqlstr &= "             " & Ddbo & "AROBLTEMP.CNTBTCH, " & Ddbo & "AROBLTEMP.CNTITEM, " & Ddbo & "AROBLTEMP.CODETAX1, OEINVH.BILNAME, OEINVH.BILADDR1, OECRDH.BILNAME AS Expr1,  "
            sqlstr &= "             OECRDH.BILADDR1 AS Expr2, " & Ddbo & "AROBLTEMP.IDRMIT, " & Ddbo & "AROBLTEMP.IDCUST, SUM(ARIBD.BASETAX1) AS BASETAX1, SUM(ARIBD.TOTTAX) AS TOTTAX, ARIBH.TEXTTRX,  "
            sqlstr &= "             ARIBH.CODECURN, " & Ddbo & "AROBLTEMP.CODECURN AS CODECURR, " & Ddbo & "AROBLTEMP.EXCHRATEHC, ARIBHO.VALUE, " & Ddbo & "AROBLTEMP.DATEBUS, "
            sqlstr &= "             ARIBT.TEXTLINE,ARPJD.CNTLINE,'DETAIL' AS VATAS "
            sqlstr &= " FROM " & Ddbo & "AROBLTEMP "
            sqlstr &= "             LEFT OUTER JOIN ARCUS ON " & Ddbo & "AROBLTEMP.IDCUST = ARCUS.IDCUST "
            sqlstr &= "             LEFT OUTER JOIN OEINVH ON " & Ddbo & "AROBLTEMP.IDINVC = OEINVH.INVNUMBER "
            sqlstr &= "             LEFT OUTER JOIN OECRDH ON " & Ddbo & "AROBLTEMP.IDINVC = OECRDH.CRDNUMBER "
            sqlstr &= "             LEFT OUTER JOIN ARIBH ON " & Ddbo & "AROBLTEMP.CNTBTCH = ARIBH.CNTBTCH AND " & Ddbo & "AROBLTEMP.CNTITEM = ARIBH.CNTITEM AND " & Ddbo & "AROBLTEMP.IDCUST = ARIBH.IDCUST AND " & Ddbo & "AROBLTEMP.IDINVC = ARIBH.IDINVC "
            sqlstr &= "             LEFT OUTER JOIN ARIBD ON ARIBH.CNTBTCH = ARIBD.CNTBTCH AND ARIBH.CNTITEM = ARIBD.CNTITEM "
            sqlstr &= "             LEFT OUTER JOIN " & Ddbo & "ARPJHTEMP ON " & Ddbo & "AROBLTEMP.IDINVC = " & Ddbo & "ARPJHTEMP.IDINVC AND " & Ddbo & "AROBLTEMP.IDCUST  = " & Ddbo & "ARPJHTEMP.IDCUST "
            sqlstr &= "             LEFT OUTER JOIN ARPJD ON " & Ddbo & "ARPJHTEMP.POSTSEQNCE = ARPJD.POSTSEQNCE AND " & Ddbo & "ARPJHTEMP.CNTBTCH = ARPJD.CNTBTCH AND " & Ddbo & "ARPJHTEMP.CNTITEM = ARPJD.CNTITEM "
            sqlstr &= "             LEFT OUTER JOIN ARIBHO ON ARIBHO.CNTBTCH = ARIBH.CNTBTCH AND ARIBHO.CNTITEM = ARIBD.CNTITEM AND ARIBHO.OPTFIELD = 'VOUCHER'"
            sqlstr &= "             LEFT OUTER JOIN ARIBT ON " & Ddbo & "AROBLTEMP.CNTBTCH = ARIBT.CNTBTCH AND " & Ddbo & "AROBLTEMP.CNTITEM = ARIBT.CNTITEM AND ARIBT.CNTLINE = ARIBD.CNTLINE "
            sqlstr &= "             LEFT OUTER JOIN TXAUTH ON " & Ddbo & "AROBLTEMP.CODETAXGRP = TXAUTH.AUTHORITY "
            sqlstr &= " WHERE     (" & Ddbo & "AROBLTEMP.TRXTYPETXT IN (1,2,3)) AND (" & Ddbo & "AROBLTEMP.DATEBUS >= " & DATEFROM & " AND " & Ddbo & "AROBLTEMP.DATEBUS <= " & DATETO & ") AND (ARIBD.TAXSTTS1 = 1) "
            sqlstr &= "           AND ((ARIBD.BASETAX1 <> 0) OR (ARIBD.TOTTAX <> 0))  "
            'Tax Amount=0
            ' หาก Tax Amount =0 จะเเสดงรายการนั้น
            sqlstr &= "             AND " & Ddbo & "AROBLTEMP.AMTTAXHC = 0"
            sqlstr &= "             AND ARPJD.IDACCT IN(SELECT ACCTVAT " & IIf(ComDatabase = "PVSW", "", "COLLATE Thai_CI_AS") & " FROM " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "FMSVLACC)   "
            sqlstr &= " GROUP BY    " & Ddbo & "AROBLTEMP.DATEINVC, " & Ddbo & "AROBLTEMP.IDINVC, " & Ddbo & "AROBLTEMP.TRXTYPEID, ARCUS.NAMECUST, ARCUS.TEXTSTRE1, " & Ddbo & "AROBLTEMP.AMTINVCHC, " & Ddbo & "AROBLTEMP.AMTTAXHC, " & Ddbo & "AROBLTEMP.CNTBTCH,  "
            sqlstr &= "             " & Ddbo & "AROBLTEMP.CNTITEM, " & Ddbo & "AROBLTEMP.CODETAX1, OEINVH.BILNAME, OEINVH.BILADDR1, OECRDH.BILNAME, OECRDH.BILADDR1, " & Ddbo & "AROBLTEMP.IDRMIT, " & Ddbo & "AROBLTEMP.IDCUST,  "
            sqlstr &= "             ARIBH.TEXTTRX, ARIBH.CODECURN, " & Ddbo & "AROBLTEMP.CODETAX1, " & Ddbo & "AROBLTEMP.CODECURN, " & Ddbo & "AROBLTEMP.EXCHRATEHC, ARIBHO.VALUE, " & Ddbo & "AROBLTEMP.DATEBUS,ARIBT.TEXTLINE,ARPJD.CNTLINE "


        End If




        rs.Open(sqlstr, cnACCPAC)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessOE> Select AR-OE Complete")
        End If




        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls
                ErrMsg = "Ledger = OE IDINVC=" & Trim(rs.options.Rs.Collect(1))
                ErrMsg = ErrMsg & " Batch No.=" & Trim(rs.options.Rs.Collect(7))
                ErrMsg = ErrMsg & " Entry No.=" & Trim(rs.options.Rs.Collect(8))

                rs2 = New BaseClass.BaseODBCIO

                '.Batch = IIf(IsDBNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7)) 'AROBL.CNTBTCH
                '.Entry = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8)) 'AROBL.CNTITEM

                'If .Batch = 82 And .Entry = 2 Then
                '    Dim Asss As String = ""
                'End If



                .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(16)), 0, rs.options.Rs.Collect(16)) 'ARIBD.BASETAX1
                sqlstr = " SELECT     CNTBTCH, CNTITEM, OPTFIELD, VALUE"
                sqlstr &= " FROM   ARIBHO "
                sqlstr &= " WHERE     (CNTBTCH = '" & Trim(rs.options.Rs.Collect(7)) & "') AND (CNTITEM = '" & Trim(rs.options.Rs.Collect(8)) & "') AND (OPTFIELD = 'TAXBASE')"
                rs2 = New BaseClass.BaseODBCIO
                rs2.Open(sqlstr, cnACCPAC)

                If rs2.options.Rs.RecordCount > 0 Then
                    .INVAMT = UCase(rs2.options.Rs.Collect(3))
                Else
                    rs2 = New BaseClass.BaseODBCIO
                    If Use_POST Then
                        sqlstr = " SELECT ARIBHO.VALUE AS BASETAX1"
                        sqlstr &= " FROM AROBL "
                        sqlstr &= "         LEFT OUTER JOIN ARCUS ON AROBL.IDCUST = ARCUS.IDCUST "
                        sqlstr &= "         LEFT OUTER JOIN OEINVH ON AROBL.IDINVC = OEINVH.INVNUMBER "
                        sqlstr &= "         LEFT OUTER JOIN OECRDH ON AROBL.IDINVC = OECRDH.CRDNUMBER "
                        sqlstr &= "         LEFT OUTER JOIN ARIBH ON AROBL.IDINVC = ARIBH.IDINVC "
                        sqlstr &= "         LEFT OUTER JOIN ARIBD ON ARIBH.CNTBTCH = ARIBD.CNTBTCH AND ARIBH.CNTITEM = ARIBD.CNTITEM "
                        sqlstr &= "         LEFT OUTER JOIN ARIBHO ON ARIBHO.CNTBTCH = ARIBD.CNTBTCH AND ARIBHO.CNTITEM = ARIBD.CNTITEM "
                        sqlstr &= " WHERE   ARIBHO.OPTFIELD = 'EXPORTAMT' "
                        sqlstr &= "         AND AROBL.CNTBTCH = " & Trim(rs.options.Rs.Collect(7)) & " AND AROBL.CNTITEM = " & Trim(rs.options.Rs.Collect(8)) & " "
                    Else
                        sqlstr = " SELECT ARIBHO.VALUE AS BASETAX1"
                        sqlstr &= " FROM ARIBH "
                        sqlstr &= "         LEFT OUTER JOIN ARCUS ON ARIBH.IDCUST = ARCUS.IDCUST "
                        sqlstr &= "         LEFT OUTER JOIN OEINVH ON ARIBH.IDINVC = OEINVH.INVNUMBER "
                        sqlstr &= "         LEFT OUTER JOIN OECRDH ON ARIBH.IDINVC = OECRDH.CRDNUMBER "
                        sqlstr &= "         LEFT OUTER JOIN ARIBD ON ARIBH.CNTBTCH = ARIBD.CNTBTCH AND ARIBH.CNTITEM = ARIBD.CNTITEM "
                        sqlstr &= "         LEFT OUTER JOIN ARIBHO ON ARIBHO.CNTBTCH = ARIBD.CNTBTCH AND ARIBHO.CNTITEM = ARIBD.CNTITEM"
                        sqlstr &= " WHERE   ARIBHO.OPTFIELD = 'EXPORTAMT' "
                        sqlstr &= "         AND ARIBH.CNTBTCH = " & Trim(rs.options.Rs.Collect("CNTBTCH")) & " AND ARIBH.CNTITEM = " & Trim(rs.options.Rs.Collect("CNTITEM")) & " "
                    End If
                    rs2.Open(sqlstr, cnACCPAC)
                    If rs2.options.Rs.RecordCount > 0 Then
                        .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(0)), 0, rs.options.Rs.Collect(0)) 'ARIBHO.VALUE
                    End If
                End If

                sqlstr = " SELECT CNTBTCH, CNTITEM, OPTFIELD, VALUE"
                sqlstr &= " FROM ARIBHO "
                sqlstr &= " WHERE (CNTBTCH = '" & Trim(rs.options.Rs.Collect(7)) & "') AND (CNTITEM = '" & Trim(rs.options.Rs.Collect(8)) & "') AND (OPTFIELD = 'TAXDATE')"
                rs2 = New BaseClass.BaseODBCIO
                rs2.Open(sqlstr, cnACCPAC)

                If rs2.options.Rs.RecordCount > 0 Then
                    .TXDATE = TryDate2(ZeroDate((rs2.options.Rs.Collect(3)).ToString.Trim, "dd/MM/yyyy", "yyyyMMdd"), "dd/MM/yyyy", "yyyyMMdd")
                    If .TXDATE = 0 Then
                        .TXDATE = TryDate2((rs2.options.Rs.Collect(3)).ToString.Trim, "dd/MM/yyyy", "yyyyMMdd")
                    End If
                Else
                    .TXDATE = rs.options.Rs.Collect("DATEBUS")

                End If

                sqlstr = " SELECT CNTBTCH, CNTITEM, OPTFIELD, VALUE"
                sqlstr &= " FROM ARIBHO "
                sqlstr &= " WHERE (CNTBTCH = '" & Trim(rs.options.Rs.Collect(7)) & "') AND (CNTITEM = '" & Trim(rs.options.Rs.Collect(8)) & "') AND (OPTFIELD = 'TAXNAME')"
                rs2 = New BaseClass.BaseODBCIO
                rs2.Open(sqlstr, cnACCPAC)

                If rs2.options.Rs.RecordCount > 0 Then
                    .INVNAME = (rs2.options.Rs.Collect(3))
                Else
                    If rs.options.Rs.Collect(2) = 14 Then
                        .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect(10)), "", (rs.options.Rs.Collect(10))) 'OEINVH.BILNAME
                        If .INVNAME.Trim <> "" Then
                            If ComVatName = True And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" Then 'เพิ่ม And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" กรณีที่ ชื้อ Customer ยาว แล้วต้องคีย์เพิ่มที่ช่อง TEXTSTRE1 (เจอที่ AIT)
                                .INVNAME = .INVNAME.Trim & " " & IIf(IsDBNull(rs.options.Rs.Collect(11)), "", (Trim(rs.options.Rs.Collect(11)))) 'OEINVH.BILADDR1
                            End If
                        End If
                    ElseIf rs.options.Rs.Collect(2) = 34 Then
                        .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect(12)), "", (rs.options.Rs.Collect(12))) 'OECRDH.BILNAME
                        If .INVNAME.Trim <> "" Then
                            If ComVatName = True And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" Then 'เพิ่ม And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" กรณีที่ ชื้อ Customer ยาว แล้วต้องคีย์เพิ่มที่ช่อง TEXTSTRE1 (เจอที่ AIT)
                                .INVNAME = .INVNAME.Trim & " " & IIf(IsDBNull(rs.options.Rs.Collect(13)), "", (Trim(rs.options.Rs.Collect(13)))) 'OECRDH.BILADDR1
                            End If
                        End If

                    End If

                    If IsDBNull(.INVNAME) Or Len(Trim(.INVNAME)) = 0 Then
                        .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect(3)), "", (rs.options.Rs.Collect(3))) 'ARCUS.NAMECUST
                        If .INVNAME.Trim <> "" Then
                            If ComVatName = True And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" Then 'เพิ่ม And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" กรณีที่ ชื้อ Customer ยาว แล้วต้องคีย์เพิ่มที่ช่อง TEXTSTRE1 (เจอที่ AIT)

                                .INVNAME = .INVNAME.Trim & " " & IIf(IsDBNull(rs.options.Rs.Collect(4)), "", (Trim(rs.options.Rs.Collect(4)))) 'ARCUS.TEXTSTRE1
                            End If
                        End If

                    End If

                End If

                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect(0)), 0, (rs.options.Rs.Collect(0))) 'AROBL.DATEINVC

                .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect(1)), "", (rs.options.Rs.Collect(1))) 'AROBL.IDINVC

                .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(17)), 0, rs.options.Rs.Collect(17)) 'ARIBD.TOTTAX

                rs1 = Nothing

                If rs.options.Rs.Collect("VATAS").ToString.Trim = "DETAIL" Then
                    Dim TEXTLINE() As String = rs.options.Rs.Collect("TEXTLINE").ToString.Split(",")
                    Dim RS_WHT_DETAIL As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from ARIBDO  where OPTFIELD in ('TAXBASE','TAXINVNO','TAXDATE','TAXNAME') and CNTBTCH ='" & rs.options.Rs.Collect("CNTBTCH") & "' and CNTITEM='" & rs.options.Rs.Collect("CNTITEM") & "' and CNTLINE ='" & IIf(IsDBNull(rs.options.Rs.Collect("CNTLINE")), "0", rs.options.Rs.Collect("CNTLINE")) & "' ", cnACCPAC)
                    If OPT_AP = False Then
                        For i As Integer = 0 To TEXTLINE.Length - 1
                            If IsDBNull(TEXTLINE(i)) = True Then
                                TEXTLINE(i) = ""
                            End If
                        Next
                        If TEXTLINE.Length >= 1 Then
                            If TEXTLINE(0).Trim <> "" And Decimal.TryParse(TEXTLINE(0), 0) Then
                                .INVAMT = CDec(TEXTLINE(0))
                            End If
                        End If

                        If TEXTLINE.Length >= 2 Then
                            If TEXTLINE(1).Trim <> "" Then
                                .IDINVC = TEXTLINE(1)
                            End If
                        End If

                        If TEXTLINE.Length >= 3 Then
                            If TEXTLINE(2).Trim <> "" Then
                                .INVDATE = TryDate2(ZeroDate(TEXTLINE(2), "dd/MM/yyyy", "yyyyMMdd"), "dd/MM/yyyy", "yyyyMMdd")
                                If .INVDATE = 0 Then
                                    .INVDATE = TryDate2(TEXTLINE(2), "dd/MM/yyyy", "yyyyMMdd")
                                End If
                            End If
                        End If

                        If TEXTLINE.Length >= 4 Then
                            If TEXTLINE(3).Trim = "" Then
                                '.INVNAME = rs.options.Rs.Collect("VENDNAME")
                                .INVNAME = rs.options.Rs.Collect("IDCUST")
                            Else
                                .INVNAME = TEXTLINE(3)
                                For CName As Integer = 4 To TEXTLINE.Length - 1
                                    .INVNAME = .INVNAME & "," & TEXTLINE(CName)
                                Next
                            End If
                        End If
                    Else
                        Dim tINVAMT As Double = 0
                        Dim tTAXINVNO As String = ""
                        Dim tTAXDATE As String = ""
                        Dim tTAXNAME As String = ""
                        For iwht As Integer = 0 To RS_WHT_DETAIL.options.QueryDT.Rows.Count - 1
                            If RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXBASE" Then
                                Dim Trydec As String = Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE"))
                                If Decimal.TryParse(Trydec, 0) = True Then
                                    If CDec(Trydec) <> 0 Then
                                        tINVAMT = CDec(Trydec)
                                    Else
                                        tINVAMT = 0
                                    End If
                                Else
                                    tINVAMT = 0
                                End If
                            ElseIf RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXINVNO" Then
                                If Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    tTAXINVNO = Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If
                            ElseIf RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXDATE" Then
                                Dim myDate As String = Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE"))
                                If Decimal.TryParse(myDate, 0) = True Then
                                    If CDec(myDate) <> 0 Then
                                        tTAXDATE = CDec(myDate)
                                    End If
                                End If
                            ElseIf RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXNAME" Then
                                If Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    tTAXNAME = Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If
                            End If
                        Next
                        If tINVAMT <> 0 Then
                            .INVAMT = tINVAMT
                        Else
                            .INVAMT = 0
                        End If
                        If tTAXINVNO <> "" Then
                            .IDINVC = tTAXINVNO
                        End If
                        If tTAXDATE <> "" Then
                            .INVDATE = tTAXDATE
                        End If
                        If tTAXNAME <> "" Then
                            .INVNAME = tTAXNAME
                        End If
                    End If
                Else 'Header

                    Code = IIf(IsDBNull(rs.options.Rs.Collect("IDCUST")), "", (rs.options.Rs.Collect("IDCUST"))) 'ARCUS.IDCUST
                    Dim RS_WHT_HEADER As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from ARIBHO  where OPTFIELD in ('TAXBASE','TAXINVNO','TAXDATE','TAXNAME') and CNTBTCH ='" & rs.options.Rs.Collect("CNTBTCH") & "' and CNTITEM='" & rs.options.Rs.Collect("CNTITEM") & "'", cnACCPAC)
                    If OPT_AP = True Then
                        Dim tINVAMT As Double = 0
                        Dim tTAXINVNO As String = ""
                        Dim tTAXDATE As String = ""
                        Dim tTAXNAME As String = ""
                        For iwht As Integer = 0 To RS_WHT_HEADER.options.QueryDT.Rows.Count - 1
                            If RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXBASE" Then
                                Dim Trydec As String = Trim(RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("VALUE"))
                                If Decimal.TryParse(Trydec, 0) = True Then
                                    If CDec(Trydec) <> 0 Then
                                        tINVAMT = CDec(Trydec)
                                    Else
                                        tINVAMT = 0
                                    End If
                                Else
                                    tINVAMT = 0
                                End If
                            ElseIf RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXINVNO" Then
                                If Trim(RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    tTAXINVNO = Trim(RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If
                            ElseIf RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXDATE" Then
                                Dim myDate As String = Trim(RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("VALUE"))
                                If Decimal.TryParse(myDate, 0) = True Then
                                    If CDec(myDate) <> 0 Then
                                        tTAXDATE = CDec(myDate)
                                    End If
                                End If
                            ElseIf RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXNAME" Then
                                If Trim(RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    tTAXNAME = Trim(RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If
                            End If
                        Next
                        If tINVAMT <> 0 Then
                            .INVAMT = tINVAMT
                        End If
                        If tTAXINVNO <> "" Then
                            .IDINVC = tTAXINVNO
                        End If
                        If tTAXDATE <> "" Then
                            .INVDATE = tTAXDATE
                        End If
                        If tTAXNAME <> "" Then
                            .INVNAME = tTAXNAME
                        End If
                    End If
                End If '"DETAIL" "Header"

                If rs.options.Rs.Collect(18) = 3 Then 'APIBH.TEXTTRX
                    If Trim(rs.options.Rs.Collect(19)) <> "THB" Then
                        .INVAMT = 0 - (.INVAMT)
                        .INVTAX = 0 - (.INVTAX)
                    Else
                        .INVAMT = 0 - (.INVAMT)
                        .INVTAX = 0 - (.INVTAX)
                    End If
                End If
                '---------------------------------------------------------------------
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7)) 'AROBL.CNTBTCH
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8)) 'AROBL.CNTITEM
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(14)), 0, rs.options.Rs.Collect(14)) 'AROBL.IDRMIT
                .Runno = Trim(IIf(IsDBNull(rs.options.Rs.Collect("value")), "", rs.options.Rs.Collect("value")))
                .TRANSNBR = Trim(IIf(IsDBNull(rs.options.Rs.Collect("CNTLINE")), "", rs.options.Rs.Collect("CNTLINE")))

                Dim loCODETAX As String = rs.options.Rs.Collect(9)

                If FindMiscCodeInARCUS(Trim(.INVNAME)) = "0" Or FindMiscCodeInARCUS(Trim(.INVNAME)) = "" Then
                    If Trim(.INVNAME) = "" Then
                        .INVNAME = FindMiscCodeInARCUS(Trim(rs.options.Rs.Collect("IDCUST")))
                    Else
                        .INVNAME = Trim(.INVNAME)
                    End If
                    .TAXID = FindTaxIDBRANCH(Trim(rs.options.Rs.Collect("IDCUST")), "AR", "TAXID")
                    .Branch = FindTaxIDBRANCH(Trim(rs.options.Rs.Collect("IDCUST")), "AR", "BRANCH")
                Else
                    .TAXID = FindTaxIDBRANCH(Trim(.INVNAME), "AR", "TAXID")
                    .Branch = FindTaxIDBRANCH(Trim(.INVNAME), "AR", "BRANCH")
                    .INVNAME = FindMiscCodeInARCUS(Trim(.INVNAME))
                End If
                If Not (rs3.options.Rs.EOF) Then
                    If Trim(rs3.options.Rs.Collect(0)) <> Trim(rs.options.Rs.Collect(20)) Then
                        .INVAMT = .INVAMT * rs.options.Rs.Collect(21)
                        .INVTAX = .INVTAX * rs.options.Rs.Collect(21)
                    End If
                End If
                .VATCOMMENT = .INVAMT & "," & .IDINVC & "," & .INVDATE & "," & .INVNAME
                .MARK = "Invoice"
                sqlstr = InsVat(0, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, "", .INVNAME, _
                                .INVAMT, .INVTAX, .LOCID, 1, .RATE_Renamed, .TTYPE, .ACCTVAT, "AR", .Batch, .Entry, _
                                .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, .Runno, loCODETAX, "", "", "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, 0)



            End With
            cnVAT.Execute(sqlstr)

            ErrMsg = ""
            Code = ""
            rs.options.Rs.MoveNext()
            Application.DoEvents()
            'End With
        Loop
        rs.options.Rs.Close()
        rs = Nothing

        'transfer data
        rs = New BaseClass.BaseODBCIO
        sqlstr = "SELECT FMSVATTEMP.INVDATE,FMSVATTEMP.TXDATE,FMSVATTEMP.IDINVC,FMSVATTEMP.DOCNO," '0-3
        sqlstr = sqlstr & " FMSVATTEMP.INVNAME,FMSVATTEMP.INVAMT,FMSVATTEMP.INVTAX," '4-6
        sqlstr = sqlstr & " FMSVLACC.LOCID,FMSVATTEMP.VTYPE,FMSVATTEMP.RATE,FMSTAX.TTYPE," '7-10
        sqlstr = sqlstr & " FMSTAX.ACCTVAT,FMSVATTEMP.SOURCE,FMSVATTEMP.BATCH,FMSVATTEMP.ENTRY," '11-14
        sqlstr = sqlstr & " FMSVATTEMP.MARK,FMSVATTEMP.VATCOMMENT,FMSTAX.ITEMRATE1,FMSVATTEMP.IDDIST,FMSVATTEMP.CBREF,FMSVATTEMP.RUNNO," '15-19
        sqlstr = sqlstr & " FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH,FMSVATTEMP.TRANSNBR,FMSVATTEMP.Code "

        If ComDatabase = "ORCL" Then
            sqlstr = sqlstr & " FROM FMSVATTEMP"
            sqlstr = sqlstr & "     INNER JOIN FMSTAX ON (FMSVATTEMP.CODETAX = FMSTAX.AUTHORITY) "
            sqlstr = sqlstr & "     INNER JOIN FMSVLACC ON FMSTAX.ACCTVAT = FMSVLACC.ACCTVAT"
            sqlstr = sqlstr & " WHERE FMSTAX.TTYPE = 2 AND FMSTAX.BUYERCLASS = 1"

        Else
            sqlstr = sqlstr & " FROM FMSVATTEMP"
            sqlstr = sqlstr & "     INNER JOIN FMSTAX ON (FMSVATTEMP.CODETAX = FMSTAX.AUTHORITY) "
            sqlstr = sqlstr & "     INNER JOIN FMSVLACC ON FMSTAX.ACCTVAT = FMSVLACC.ACCTVAT"
            sqlstr = sqlstr & " WHERE FMSTAX.TTYPE = 2 AND FMSTAX.BUYERCLASS = 1"
        End If

        rs.Open(sqlstr, cnVAT)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessOE> Select FMSVATTEMP Complete")
        End If


        Do While rs.options.Rs.EOF = False
            loCls = New clsVat

            With loCls
                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect(0)), 0, rs.options.Rs.Collect(0))
                .TXDATE = IIf(IsDBNull(rs.options.Rs.Collect(1)), 0, rs.options.Rs.Collect(1))
                .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", (rs.options.Rs.Collect(2)))
                .DOCNO = IIf(IsDBNull(rs.options.Rs.Collect(3)), "", (rs.options.Rs.Collect(3)))
                .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect(4)), "", (rs.options.Rs.Collect(4)))
                .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(5)), 0, rs.options.Rs.Collect(5))
                .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(6)), 0, rs.options.Rs.Collect(6))
                .LOCID = IIf(IsDBNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7))
                .VTYPE = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8))
                .TTYPE = IIf(IsDBNull(rs.options.Rs.Collect(10)), 0, rs.options.Rs.Collect(10))
                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect(11)), "", (rs.options.Rs.Collect(11)))
                .Source = IIf(IsDBNull(rs.options.Rs.Collect(12)), "", (rs.options.Rs.Collect(12)))
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(13)), "", (rs.options.Rs.Collect(13)))
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(14)), "", (rs.options.Rs.Collect(14)))
                .MARK = IIf(IsDBNull(rs.options.Rs.Collect(15)), "", (rs.options.Rs.Collect(15)))
                .VATCOMMENT = IIf(IsDBNull(rs.options.Rs.Collect(16)), "", (rs.options.Rs.Collect(16)))
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(19)), "", (rs.options.Rs.Collect(19)))
                .TRANSNBR = IIf(IsDBNull(rs.options.Rs.Collect("TRANSNBR")), "", rs.options.Rs.Collect("TRANSNBR"))
                .Runno = IIf(IsDBNull(rs.options.Rs.Collect("Runno")), "", rs.options.Rs.Collect("Runno"))
                .TAXID = IIf(IsDBNull(rs.options.Rs.Collect("TAXID")), "", rs.options.Rs.Collect("TAXID"))
                .Branch = IIf(IsDBNull(rs.options.Rs.Collect("BRANCH")), "", rs.options.Rs.Collect("BRANCH"))
                Code = IIf(IsDBNull(rs.options.Rs.Collect("Code")), "", rs.options.Rs.Collect("Code"))


                If .INVAMT = 0 Then
                    .RATE_Renamed = 0
                Else
                    .RATE_Renamed = CDbl(((loCls.INVTAX / loCls.INVAMT) * 100))
                End If
                sqlstr = InsVat(1, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, .NEWDOCNO, .INVNAME, .INVAMT, .INVTAX, .LOCID, .VTYPE, .RATE_Renamed, .TTYPE, .ACCTVAT, .Source, .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, .Runno, "", "", "", "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, 0)

            End With
            cnVAT.Execute(sqlstr)
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing
        cnACCPAC.Close()
        cnVAT.Close()
        cnACCPAC = Nothing
        cnVAT = Nothing
        Exit Sub
ErrHandler:
        WriteLog("<ProcessOE>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessOE> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
        WriteLog(ErrMsg)
    End Sub

    Public Sub ProcessOE()
        Dim cnACCPAC, cnVAT As ADODB.Connection
        Dim rs, rs1, rs2, rs3 As BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim loCls As clsVat
        Dim iffs As clsFuntion
        Dim Ddbo, HOMECUR As String
        iffs = New clsFuntion
        ErrMsg = ""


        On Error GoTo ErrHandler

        cnACCPAC = New ADODB.Connection
        cnACCPAC.ConnectionTimeout = 60
        cnACCPAC.Open(ConnACCPAC)
        cnACCPAC.CommandTimeout = 3600

        Call Connecttion("ACC")


        'gather data
        'rs3 = New BaseClass.BaseODBCIO
        sqlstr = "SELECT HOMECUR FROM CSCOM "
        'rs3.Open(sqlstr, cnACCPAC)
        com = New SqlCommand(sqlstr, ConA)
        dr = com.ExecuteReader
        dr.Read()
        HOMECUR = dr("HOMECUR")
        dr.Close()



        Ddbo = ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.")

        rs = New BaseClass.BaseODBCIO

        If Use_POST Then
            sqlstr = " SELECT " & Ddbo & "AROBLTEMP.DATEINVC, " & Ddbo & "AROBLTEMP.IDINVC, " & Ddbo & "AROBLTEMP.TRXTYPEID, ARCUS.NAMECUST, ARCUS.TEXTSTRE1, SUM(ARIBD.AMTEXTN) AS AMTEXTN, SUM(ARIBD.TOTTAX) AS TOTTAX, "
            sqlstr &= "         " & Ddbo & "AROBLTEMP.CNTBTCH, " & Ddbo & "AROBLTEMP.CNTITEM, " & Ddbo & "AROBLTEMP.CODETAX1, OEINVH.BILNAME, OEINVH.BILADDR1, OECRDH.BILNAME AS Expr1,  "
            sqlstr &= "         OECRDH.BILADDR1 AS Expr2, " & Ddbo & "AROBLTEMP.IDRMIT, " & Ddbo & "AROBLTEMP.IDCUST, SUM(ARIBD.BASETAX1) AS BASETAX1, SUM(ARIBD.TOTTAX) AS TOTTAX, ARIBH.TEXTTRX,  "
            sqlstr &= "         ARIBH.CODECURN, " & Ddbo & "AROBLTEMP.CODECURN AS CODECURR, " & Ddbo & "AROBLTEMP.EXCHRATEHC, ARIBHO.VALUE, " & Ddbo & "AROBLTEMP.DATEBUS, "
            sqlstr &= "         '' AS TEXTLINE,0 AS CNTLINE,'HEADER' AS VATAS "
            sqlstr &= " FROM " & Ddbo & "AROBLTEMP"
            sqlstr &= "         LEFT OUTER JOIN ARCUS ON " & Ddbo & "AROBLTEMP.IDCUST = ARCUS.IDCUST "
            sqlstr &= "         LEFT OUTER JOIN OEINVH ON " & Ddbo & "AROBLTEMP.IDINVC = OEINVH.INVNUMBER "
            sqlstr &= "         LEFT OUTER JOIN OECRDH ON " & Ddbo & "AROBLTEMP.IDINVC = OECRDH.CRDNUMBER "
            sqlstr &= "         LEFT OUTER JOIN ARIBH ON " & Ddbo & "AROBLTEMP.CNTBTCH = ARIBH.CNTBTCH AND " & Ddbo & "AROBLTEMP.CNTITEM = ARIBH.CNTITEM AND " & Ddbo & "AROBLTEMP.IDCUST = ARIBH.IDCUST AND " & Ddbo & "AROBLTEMP.IDINVC = ARIBH.IDINVC "
            sqlstr &= "         LEFT OUTER JOIN ARIBD ON ARIBH.CNTBTCH = ARIBD.CNTBTCH AND ARIBH.CNTITEM = ARIBD.CNTITEM "
            sqlstr &= "         LEFT OUTER JOIN ARIBHO ON ARIBHO.CNTBTCH = ARIBH.CNTBTCH AND ARIBHO.CNTITEM = ARIBD.CNTITEM AND ARIBHO.OPTFIELD = 'VOUCHER'"
            sqlstr &= "         LEFT OUTER JOIN TXAUTH ON " & Ddbo & "AROBLTEMP.CODETAXGRP = TXAUTH.AUTHORITY "
            sqlstr &= " WHERE   (" & Ddbo & "AROBLTEMP.TRXTYPETXT IN (1,2,3)) AND (" & Ddbo & "AROBLTEMP.DATEBUS >= " & DATEFROM & " AND " & Ddbo & "AROBLTEMP.DATEBUS <= " & DATETO & ") AND (ARIBD.TAXSTTS1 = 1) "
            sqlstr &= "         AND ((ARIBD.BASETAX1 <> 0) OR (ARIBD.TOTTAX <> 0)) "
            'Tax Amount<>0
            ' หาก Tax Amount =0 จะไม่เเสดงรายการนั้น
            'sqlstr &= " and " & Ddbo & "AROBLTEMP.AMTTAXHC<>0  " commentไว้ Vat 0 ก็ออกตามปกติ
            'TXAUTH.ACCTRECOV แก้ไขเนื่องจาก account ขาย อยู่ที่ TXAUTH.LIABILITY  ,ซื้อ อยู่ที่  TXAUTH.ACCTRECOV
            sqlstr &= "         AND  TXAUTH.LIABILITY IN (SELECT ACCTVAT " & IIf(ComDatabase = "PVSW", "", "COLLATE Thai_CI_AS") & " FROM " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "FMSVLACC)   "
            sqlstr &= " GROUP BY " & Ddbo & "AROBLTEMP.DATEINVC, " & Ddbo & "AROBLTEMP.IDINVC, " & Ddbo & "AROBLTEMP.TRXTYPEID, ARCUS.NAMECUST, ARCUS.TEXTSTRE1, " & Ddbo & "AROBLTEMP.AMTINVCHC, " & Ddbo & "AROBLTEMP.AMTTAXHC, " & Ddbo & "AROBLTEMP.CNTBTCH,  "
            sqlstr &= "         " & Ddbo & "AROBLTEMP.CNTITEM, " & Ddbo & "AROBLTEMP.CODETAX1, OEINVH.BILNAME, OEINVH.BILADDR1, OECRDH.BILNAME, OECRDH.BILADDR1, " & Ddbo & "AROBLTEMP.IDRMIT, " & Ddbo & "AROBLTEMP.IDCUST,  "
            sqlstr &= "         ARIBH.TEXTTRX, ARIBH.CODECURN, " & Ddbo & "AROBLTEMP.CODETAX1, " & Ddbo & "AROBLTEMP.CODECURN, " & Ddbo & "AROBLTEMP.EXCHRATEHC, ARIBHO.VALUE, " & Ddbo & "AROBLTEMP.DATEBUS "

            sqlstr &= " UNION ALL "

            sqlstr &= " SELECT " & Ddbo & "AROBLTEMP.DATEINVC, " & Ddbo & "AROBLTEMP.IDINVC, " & Ddbo & "AROBLTEMP.TRXTYPEID, ARCUS.NAMECUST, ARCUS.TEXTSTRE1, SUM(ARIBD.AMTEXTN) AS AMTEXTN, SUM(ARIBD.TOTTAX) AS TOTTAX, "
            sqlstr &= "             " & Ddbo & "AROBLTEMP.CNTBTCH, " & Ddbo & "AROBLTEMP.CNTITEM, " & Ddbo & "AROBLTEMP.CODETAX1, OEINVH.BILNAME, OEINVH.BILADDR1, OECRDH.BILNAME AS Expr1,  "
            sqlstr &= "             OECRDH.BILADDR1 AS Expr2, " & Ddbo & "AROBLTEMP.IDRMIT, " & Ddbo & "AROBLTEMP.IDCUST, SUM(ARIBD.BASETAX1) AS BASETAX1, SUM(ARIBD.TOTTAX) AS TOTTAX, ARIBH.TEXTTRX,  "
            sqlstr &= "             ARIBH.CODECURN, " & Ddbo & "AROBLTEMP.CODECURN AS CODECURR, " & Ddbo & "AROBLTEMP.EXCHRATEHC, ARIBHO.VALUE, " & Ddbo & "AROBLTEMP.DATEBUS, "
            sqlstr &= "             ARIBT.TEXTLINE,ARPJD.CNTLINE,'DETAIL' AS VATAS "
            sqlstr &= " FROM " & Ddbo & "AROBLTEMP "
            sqlstr &= "             LEFT OUTER JOIN ARCUS ON " & Ddbo & "AROBLTEMP.IDCUST = ARCUS.IDCUST "
            sqlstr &= "             LEFT OUTER JOIN OEINVH ON " & Ddbo & "AROBLTEMP.IDINVC = OEINVH.INVNUMBER "
            sqlstr &= "             LEFT OUTER JOIN OECRDH ON " & Ddbo & "AROBLTEMP.IDINVC = OECRDH.CRDNUMBER "
            sqlstr &= "             LEFT OUTER JOIN ARIBH ON " & Ddbo & "AROBLTEMP.CNTBTCH = ARIBH.CNTBTCH AND " & Ddbo & "AROBLTEMP.CNTITEM = ARIBH.CNTITEM AND " & Ddbo & "AROBLTEMP.IDCUST = ARIBH.IDCUST AND " & Ddbo & "AROBLTEMP.IDINVC = ARIBH.IDINVC "
            sqlstr &= "             LEFT OUTER JOIN ARIBD ON ARIBH.CNTBTCH = ARIBD.CNTBTCH AND ARIBH.CNTITEM = ARIBD.CNTITEM "
            sqlstr &= "             LEFT OUTER JOIN " & Ddbo & "ARPJHTEMP ON " & Ddbo & "AROBLTEMP.IDINVC = " & Ddbo & "ARPJHTEMP.IDINVC AND " & Ddbo & "AROBLTEMP.IDCUST  = " & Ddbo & "ARPJHTEMP.IDCUST "
            sqlstr &= "             LEFT OUTER JOIN ARPJD ON " & Ddbo & "ARPJHTEMP.POSTSEQNCE = ARPJD.POSTSEQNCE AND " & Ddbo & "ARPJHTEMP.CNTBTCH = ARPJD.CNTBTCH AND " & Ddbo & "ARPJHTEMP.CNTITEM = ARPJD.CNTITEM "
            sqlstr &= "             LEFT OUTER JOIN ARIBHO ON ARIBHO.CNTBTCH = ARIBH.CNTBTCH AND ARIBHO.CNTITEM = ARIBD.CNTITEM AND ARIBHO.OPTFIELD = 'VOUCHER'"
            sqlstr &= "             LEFT OUTER JOIN ARIBT ON " & Ddbo & "AROBLTEMP.CNTBTCH = ARIBT.CNTBTCH AND " & Ddbo & "AROBLTEMP.CNTITEM = ARIBT.CNTITEM AND ARIBT.CNTLINE = ARIBD.CNTLINE "
            sqlstr &= "             LEFT OUTER JOIN TXAUTH ON " & Ddbo & "AROBLTEMP.CODETAXGRP = TXAUTH.AUTHORITY "
            sqlstr &= " WHERE     (" & Ddbo & "AROBLTEMP.TRXTYPETXT IN (1,2,3)) AND (" & Ddbo & "AROBLTEMP.DATEBUS >= " & DATEFROM & " AND " & Ddbo & "AROBLTEMP.DATEBUS <= " & DATETO & ") AND (ARIBD.TAXSTTS1 = 1) "
            sqlstr &= "           AND ((ARIBD.BASETAX1 <> 0) OR (ARIBD.TOTTAX <> 0))  "
            'Tax Amount=0
            ' หาก Tax Amount =0 จะเเสดงรายการนั้น
            sqlstr &= "             AND " & Ddbo & "AROBLTEMP.AMTTAXHC = 0"
            sqlstr &= "             AND ARPJD.IDACCT IN(SELECT ACCTVAT " & IIf(ComDatabase = "PVSW", "", "COLLATE Thai_CI_AS") & " FROM " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "FMSVLACC)   "
            sqlstr &= " GROUP BY    " & Ddbo & "AROBLTEMP.DATEINVC, " & Ddbo & "AROBLTEMP.IDINVC, " & Ddbo & "AROBLTEMP.TRXTYPEID, ARCUS.NAMECUST, ARCUS.TEXTSTRE1, " & Ddbo & "AROBLTEMP.AMTINVCHC, " & Ddbo & "AROBLTEMP.AMTTAXHC, " & Ddbo & "AROBLTEMP.CNTBTCH,  "
            sqlstr &= "             " & Ddbo & "AROBLTEMP.CNTITEM, " & Ddbo & "AROBLTEMP.CODETAX1, OEINVH.BILNAME, OEINVH.BILADDR1, OECRDH.BILNAME, OECRDH.BILADDR1, " & Ddbo & "AROBLTEMP.IDRMIT, " & Ddbo & "AROBLTEMP.IDCUST,  "
            sqlstr &= "             ARIBH.TEXTTRX, ARIBH.CODECURN, " & Ddbo & "AROBLTEMP.CODETAX1, " & Ddbo & "AROBLTEMP.CODECURN, " & Ddbo & "AROBLTEMP.EXCHRATEHC, ARIBHO.VALUE, " & Ddbo & "AROBLTEMP.DATEBUS,ARIBT.TEXTLINE,ARPJD.CNTLINE "

        End If


        Call Connecttion("ACC")

        daNB = New SqlDataAdapter(sqlstr, ConA)
        dtselect = New DataTable
        dtselect.Clear()
        daNB.Fill(dtselect)
        ConA.Close()
        WriteLog("<ProcessOE> Select AR-OE Complete")

        For i As Integer = 0 To dtselect.Rows.Count - 1

            loCls = New clsVat
            With loCls
                ErrMsg = "Ledger = OE IDINVC=" & dtselect.Rows(i)(1).ToString.Trim
                ErrMsg = ErrMsg & " Batch No.=" & dtselect.Rows(i)(7).ToString.Trim
                ErrMsg = ErrMsg & " Entry No.=" & dtselect.Rows(i)(8).ToString.Trim

                If dtselect.Rows(i)(2) = 14 Then
                    .INVNAME = IIf(IsDBNull(dtselect.Rows(i)(10).ToString.Trim), "", dtselect.Rows(i)(10).ToString.Trim) 'OEINVH.BILNAME
                    If .INVNAME.Trim <> "" Then
                        If ComVatName = True And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" Then 'เพิ่ม And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" กรณีที่ ชื้อ Customer ยาว แล้วต้องคีย์เพิ่มที่ช่อง TEXTSTRE1 (เจอที่ AIT)
                            .INVNAME = .INVNAME.Trim & " " & IIf(IsDBNull(dtselect.Rows(i)(11).ToString.Trim), "", dtselect.Rows(i)(11).ToString.Trim) 'OEINVH.BILADDR1
                        End If
                    End If

                ElseIf dtselect.Rows(i)(2) = 34 Then
                    .INVNAME = IIf(IsDBNull(dtselect.Rows(i)(12).ToString.Trim), "", dtselect.Rows(i)(12).ToString.Trim) 'OECRDH.BILNAME
                    If .INVNAME.Trim <> "" Then
                        If ComVatName = True And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" Then 'เพิ่ม And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" กรณีที่ ชื้อ Customer ยาว แล้วต้องคีย์เพิ่มที่ช่อง TEXTSTRE1 (เจอที่ AIT)
                            .INVNAME = .INVNAME.Trim & " " & IIf(IsDBNull(dtselect.Rows(i)(13).ToString.Trim), "", dtselect.Rows(i)(13).ToString.Trim) 'OECRDH.BILADDR1
                        End If
                    End If

                End If

                If IsDBNull(.INVNAME) Or Len(Trim(.INVNAME)) = 0 Then
                    .INVNAME = IIf(IsDBNull(dtselect.Rows(i)(13).ToString.Trim), "", dtselect.Rows(i)(3).ToString.Trim) 'ARCUS.NAMECUST
                    If .INVNAME.Trim <> "" Then
                        If ComVatName = True And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" Then 'เพิ่ม And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" กรณีที่ ชื้อ Customer ยาว แล้วต้องคีย์เพิ่มที่ช่อง TEXTSTRE1 (เจอที่ AIT)

                            .INVNAME = .INVNAME.Trim & " " & IIf(IsDBNull(dtselect.Rows(i)(4).ToString.Trim), "", dtselect.Rows(i)(4).ToString.Trim) 'ARCUS.TEXTSTRE1
                        End If
                    End If

                End If


                .INVAMT = IIf(IsDBNull(dtselect.Rows(i)(16)), 0, dtselect.Rows(i)(16)) 'ARIBD.BASETAX1
                .TXDATE = dtselect.Rows(i)("DATEBUS")
                .INVDATE = IIf(IsDBNull(dtselect.Rows(i)(0)), 0, dtselect.Rows(i)(0)) 'AROBL.DATEINVC
                .IDINVC = IIf(IsDBNull(dtselect.Rows(i)(1).ToString.Trim), "", dtselect.Rows(i)(1).ToString.Trim) 'AROBL.IDINVC
                .INVTAX = IIf(IsDBNull(dtselect.Rows(i)(17)), 0, dtselect.Rows(i)(17)) 'ARIBD.TOTTAX


                sqlstr = " SELECT CNTBTCH, CNTITEM, OPTFIELD, VALUE"
                sqlstr &= " FROM ARIBHO "
                sqlstr &= " WHERE OPTFIELD in ('TAXBASE','TAXINVNO','TAXDATE','TAXNAME') AND (CNTBTCH = '" & dtselect.Rows(i)(7).ToString.Trim & "') AND (CNTITEM = '" & dtselect.Rows(i)(8).ToString.Trim & "') "

                Call Connecttion("ACC")

                daNB = New SqlDataAdapter(sqlstr, ConA)
                dtTemp = New DataTable
                dtTemp.Clear()
                daNB.Fill(dtTemp)
                ConA.Close()

                For k As Integer = 0 To dtTemp.Rows.Count - 1

                    If dtTemp.Rows(k)("OPTFIELD").ToString.Trim = "TAXBASE" Then

                        .INVAMT = dtTemp.Rows(k)("VALUE")

                    ElseIf dtTemp.Rows(k)("OPTFIELD").ToString.Trim = "TAXINVNO" Then

                        .IDINVC = dtTemp.Rows(k)("VALUE")

                    ElseIf dtTemp.Rows(k)("OPTFIELD").ToString.Trim = "TAXDATE" Then

                        .INVDATE = TryDate2(ZeroDate((dtTemp.Rows(k)("VALUE")).ToString.Trim, "dd/MM/yyyy", "yyyyMMdd"), "dd/MM/yyyy", "yyyyMMdd")
                        If .INVDATE = 0 Then
                            .INVDATE = TryDate2((dtTemp.Rows(k)("VALUE")).ToString.Trim, "dd/MM/yyyy", "yyyyMMdd")
                        End If

                    ElseIf dtTemp.Rows(k)("OPTFIELD").ToString.Trim = "TAXNAME" Then

                        .INVNAME = dtTemp.Rows(k)("VALUE")

                    End If

                Next

                If dtselect.Rows(i)("VATAS").ToString.Trim = "HEADER" Then

                    Code = IIf(IsDBNull(dtselect.Rows(i)("IDCUST").ToString.Trim), "", dtselect.Rows(i)("IDCUST").ToString.Trim) 'ARCUS.IDCUST

                    'AR ไม่มีให้เลือกคีย์ VAT ที่ OPT แต่ Code ส่วนนี้ไม่ทราบว่าใครเขียนไว้ เอาส่วนของ AP มาเช็คทำไมก็ไม่รู้  >> OPT_AP เป็นของ AP
                    If OPT_AP = True Then
                        sqlstr = "select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from ARIBHO  where OPTFIELD in ('TAXBASE','TAXINVNO','TAXDATE','TAXNAME') and CNTBTCH ='" & dtselect.Rows(i)("CNTBTCH") & "' and CNTITEM='" & dtselect.Rows(i)("CNTITEM") & "'"


                        Call Connecttion("ACC")

                        daNB = New SqlDataAdapter(sqlstr, ConA)
                        dtTemp = New DataTable
                        dtTemp.Clear()
                        daNB.Fill(dtTemp)
                        ConA.Close()

                        Dim tINVAMT As Double = 0
                        Dim tTAXINVNO As String = ""
                        Dim tTAXDATE As String = ""
                        Dim tTAXNAME As String = ""

                        For iwht As Integer = 0 To dtTemp.Rows.Count - 1

                            If dtTemp.Rows(iwht)("OPTFIELD").ToString.Trim = "TAXBASE" Then
                                Dim Trydec As String = dtTemp.Rows(iwht)("VALUE").ToString.Trim

                                If Decimal.TryParse(Trydec, 0) = True Then
                                    If CDec(Trydec) <> 0 Then
                                        tINVAMT = CDec(Trydec)
                                    Else
                                        tINVAMT = 0
                                    End If
                                Else
                                    tINVAMT = 0
                                End If

                            ElseIf dtTemp.Rows(iwht)("OPTFIELD").ToString.Trim = "TAXINVNO" Then
                                If dtTemp.Rows(iwht).Item("VALUE").ToString.Trim <> "" Then
                                    tTAXINVNO = dtTemp.Rows(iwht)("VALUE").ToString.Trim
                                End If

                            ElseIf dtTemp.Rows(iwht)("OPTFIELD").ToString.Trim = "TAXDATE" Then
                                Dim myDate As String = dtTemp.Rows(iwht)("VALUE").ToString.Trim
                                If Decimal.TryParse(myDate, 0) = True Then
                                    If CDec(myDate) <> 0 Then
                                        tTAXDATE = CDec(myDate)
                                    End If
                                End If

                            ElseIf dtTemp.Rows(iwht)("OPTFIELD").ToString.Trim = "TAXNAME" Then
                                If dtTemp.Rows(iwht).Item("VALUE").ToString.Trim <> "" Then
                                    tTAXNAME = dtTemp.Rows(iwht)("VALUE").ToString.Trim
                                End If
                            End If

                        Next


                        If tINVAMT <> 0 Then
                            .INVAMT = tINVAMT
                        End If
                        If tTAXINVNO <> "" Then
                            .IDINVC = tTAXINVNO
                        End If
                        If tTAXDATE <> "" Then
                            .INVDATE = tTAXDATE
                        End If
                        If tTAXNAME <> "" Then
                            .INVNAME = tTAXNAME
                        End If
                    End If


                    If VNO_AR = True Then

                        Dim VNO_Header As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from ARIBHO  where OPTFIELD in ('ARINVNO') and CNTBTCH ='" & dtselect.Rows(i)("CNTBTCH").ToString.Trim & "' and CNTITEM='" & dtselect.Rows(i)("CNTITEM").ToString.Trim & "'", cnACCPAC)
                        For iwht As Integer = 0 To VNO_Header.options.QueryDT.Rows.Count - 1
                            If VNO_Header.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "ARINVNO" Then
                                If Trim(VNO_Header.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    .DOCNO = Trim(VNO_Header.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If

                            End If
                        Next

                    End If


                ElseIf dtselect.Rows(i)("VATAS").ToString.Trim = "DETAIL" Then

                    Dim TEXTLINE() As String = dtselect.Rows(i)("TEXTLINE").ToString.Split(",")
                    'Dim RS_WHT_DETAIL As New BaseClass.BaseODBCIO


                    'AR ไม่มีให้เลือกคีย์ VAT ที่ OPT แต่ Code ส่วนนี้ไม่ทราบว่าใครเขียนไว้ เอาส่วนของ AP มาเช็คทำไมก็ไม่รู้  >> OPT_AP เป็นของ AP
                    If OPT_AP = False Then

                        For ii As Integer = 0 To TEXTLINE.Length - 1
                            If IsDBNull(TEXTLINE(ii)) = True Then
                                TEXTLINE(ii) = ""
                            End If
                        Next

                        If TEXTLINE.Length >= 1 Then
                            If TEXTLINE(0).Trim <> "" And Decimal.TryParse(TEXTLINE(0), 0) Then
                                .INVAMT = CDec(TEXTLINE(0))
                            End If
                        End If

                        If TEXTLINE.Length >= 2 Then
                            If TEXTLINE(1).Trim <> "" Then
                                .IDINVC = TEXTLINE(1)
                            End If
                        End If

                        If TEXTLINE.Length >= 3 Then
                            If TEXTLINE(2).Trim <> "" Then
                                .INVDATE = TryDate2(ZeroDate(TEXTLINE(2), "dd/MM/yyyy", "yyyyMMdd"), "dd/MM/yyyy", "yyyyMMdd")
                                If .INVDATE = 0 Then
                                    .INVDATE = TryDate2(TEXTLINE(2), "dd/MM/yyyy", "yyyyMMdd")
                                End If
                            End If
                        End If

                        If TEXTLINE.Length >= 4 Then
                            If TEXTLINE(3).Trim = "" Then

                                .INVNAME = dtselect.Rows(i)("IDCUST").ToString.Trim
                            Else
                                .INVNAME = TEXTLINE(3)
                                For CName As Integer = 4 To TEXTLINE.Length - 1
                                    .INVNAME = .INVNAME & "," & TEXTLINE(CName)
                                Next
                            End If
                        End If

                    Else 'OPT_AP = True 

                        sqlstr = "select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from ARIBDO  where OPTFIELD in ('TAXBASE','TAXINVNO','TAXDATE','TAXNAME') and CNTBTCH ='" & dtselect.Rows(i)("CNTBTCH") & "' and CNTITEM='" & dtselect.Rows(i)("CNTITEM") & "' and CNTLINE ='" & IIf(IsDBNull(dtselect.Rows(i)("CNTLINE")), "0", dtselect.Rows(i)("CNTLINE")) & "' "

                        Call Connecttion("ACC")

                        daNB = New SqlDataAdapter(sqlstr, ConA)
                        dtTemp = New DataTable
                        dtTemp.Clear()
                        daNB.Fill(dtTemp)
                        ConA.Close()

                        Dim tINVAMT As Double = 0
                        Dim tTAXINVNO As String = ""
                        Dim tTAXDATE As String = ""
                        Dim tTAXNAME As String = ""

                        For iwht As Integer = 0 To dtTemp.Rows.Count - 1

                            If dtTemp.Rows(iwht)("OPTFIELD").ToString.Trim = "TAXBASE" Then
                                Dim Trydec As String = dtTemp.Rows(iwht)("VALUE").ToString.Trim

                                If Decimal.TryParse(Trydec, 0) = True Then
                                    If CDec(Trydec) <> 0 Then
                                        tINVAMT = CDec(Trydec)
                                    Else
                                        tINVAMT = 0
                                    End If
                                Else
                                    tINVAMT = 0
                                End If

                            ElseIf dtTemp.Rows(iwht)("OPTFIELD").ToString.Trim = "TAXINVNO" Then

                                If dtTemp.Rows(iwht)("VALUE").ToString.Trim <> "" Then
                                    tTAXINVNO = dtTemp.Rows(iwht)("VALUE").ToString.Trim
                                End If

                            ElseIf dtTemp.Rows(iwht)("OPTFIELD").ToString.Trim = "TAXDATE" Then

                                Dim myDate As String = dtTemp.Rows(iwht)("VALUE").ToString.Trim
                                If Decimal.TryParse(myDate, 0) = True Then
                                    If CDec(myDate) <> 0 Then
                                        tTAXDATE = CDec(myDate)
                                    End If
                                End If

                            ElseIf dtTemp.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXNAME" Then
                                If dtTemp.Rows(iwht).Item("VALUE").ToString.Trim <> "" Then
                                    tTAXNAME = dtTemp.Rows(iwht)("VALUE").ToString.Trim
                                End If
                            End If

                        Next


                        If tINVAMT <> 0 Then
                            .INVAMT = tINVAMT
                        Else
                            .INVAMT = 0
                        End If
                        If tTAXINVNO <> "" Then
                            .IDINVC = tTAXINVNO
                        End If
                        If tTAXDATE <> "" Then
                            .INVDATE = tTAXDATE
                        End If
                        If tTAXNAME <> "" Then
                            .INVNAME = tTAXNAME
                        End If
                    End If

                    If VNO_AR = True Then

                        Dim VNO_Detail As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from ARIBDO  where OPTFIELD in ('ARINVNO') and CNTBTCH ='" & dtselect.Rows(i)("CNTBTCH").ToString.Trim & "' and CNTITEM='" & dtselect.Rows(i)("CNTITEM").ToString.Trim & "'", cnACCPAC)
                        For iwht As Integer = 0 To VNO_Detail.options.QueryDT.Rows.Count - 1
                            If VNO_Detail.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "ARINVNO" Then
                                If Trim(VNO_Detail.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    .DOCNO = Trim(VNO_Detail.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If

                            End If
                        Next

                    End If

                End If '"DETAIL" "Header"


                If dtselect.Rows(i)(18) = 3 Then 'APIBH.TEXTTRX
                    If dtselect.Rows(i)(19).ToString.Trim <> "THB" Then
                        .INVAMT = 0 - (.INVAMT)
                        .INVTAX = 0 - (.INVTAX)
                    Else
                        .INVAMT = 0 - (.INVAMT)
                        .INVTAX = 0 - (.INVTAX)
                    End If
                End If

                .Batch = IIf(IsDBNull(dtselect.Rows(i)(7)), 0, dtselect.Rows(i)(7)) 'AROBL.CNTBTCH
                .Entry = IIf(IsDBNull(dtselect.Rows(i)(8)), 0, dtselect.Rows(i)(8)) 'AROBL.CNTITEM
                .CBRef = IIf(IsDBNull(dtselect.Rows(i)(14)), 0, dtselect.Rows(i)(14)) 'AROBL.IDRMIT
                .Runno = IIf(IsDBNull(dtselect.Rows(i)("value").ToString.Trim), "", dtselect.Rows(i)("value").ToString.Trim)
                .TRANSNBR = IIf(IsDBNull(dtselect.Rows(i)("CNTLINE").ToString.Trim), "", dtselect.Rows(i)("CNTLINE").ToString.Trim)

                Dim loCODETAX As String = dtselect.Rows(i)(9).ToString.Trim

                If FindMiscCodeInARCUS(Trim(.INVNAME)) = "0" Or FindMiscCodeInARCUS(Trim(.INVNAME)) = "" Then
                    If Trim(.INVNAME) = "" Then
                        .INVNAME = FindMiscCodeInARCUS(dtselect.Rows(i)("IDCUST").ToString.Trim)
                    Else
                        .INVNAME = Trim(.INVNAME)
                    End If
                    .TAXID = FindTaxIDBRANCH(dtselect.Rows(i)("IDCUST").ToString.Trim, "AR", "TAXID")
                    .Branch = FindTaxIDBRANCH(dtselect.Rows(i)("IDCUST".ToString.Trim), "AR", "BRANCH")
                Else
                    .TAXID = FindTaxIDBRANCH(Trim(.INVNAME), "AR", "TAXID")
                    .Branch = FindTaxIDBRANCH(Trim(.INVNAME), "AR", "BRANCH")
                    .INVNAME = FindMiscCodeInARCUS(Trim(.INVNAME))
                End If

                'ถ้าไม่ใช่ "THB" ให้คูณ ExchangRate
                If HOMECUR.ToString.Trim <> dtselect.Rows(i)(20).ToString.Trim Then
                    .INVAMT = .INVAMT * dtselect.Rows(i)(21)
                    .INVTAX = .INVTAX * dtselect.Rows(i)(21)
                End If


                .VATCOMMENT = .INVAMT & "," & .IDINVC & "," & .INVDATE & "," & .INVNAME
                .MARK = "Invoice"


                If .INVAMT = 0 Then
                    .RATE_Renamed = 0
                Else
                    .RATE_Renamed = FormatNumber(((.INVTAX / .INVAMT) * 100), 1)
                End If


                Application.DoEvents()

                sqlstr = InsVat(0, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, "", .INVNAME, .INVAMT, .INVTAX, .LOCID, 1, .RATE_Renamed, .TTYPE, .ACCTVAT, "AR", .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, .Runno, loCODETAX, "", "", "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, 0)


            End With

            'cnVAT.Execute(sqlstr)

            Call Connecttion("VAT")

            com = New SqlCommand(sqlstr, ConV)
            com.CommandText = sqlstr
            com.CommandTimeout = 0
            com.ExecuteNonQuery()

            ConV.Close()


            ErrMsg = ""
            Code = ""
            'rs.options.Rs.MoveNext()
            Application.DoEvents()
            'End With
            'Loop
        Next


        'rs.options.Rs.Close()
        'rs = Nothing


        'cnVAT = New ADODB.Connection
        ''cnACCPAC.ConnectionTimeout = 60
        ''cnACCPAC.Open(ConnACCPAC)
        ''cnACCPAC.CommandTimeout = 3600

        'cnVAT.ConnectionTimeout = 60
        'cnVAT.Open(ConnVAT)
        'cnVAT.CommandTimeout = 3600

        'transfer data
        ' rs = New BaseClass.BaseODBCIO
        sqlstr = "SELECT FMSVATTEMP.INVDATE,FMSVATTEMP.TXDATE,FMSVATTEMP.IDINVC,FMSVATTEMP.DOCNO," '0-3
        sqlstr = sqlstr & " FMSVATTEMP.INVNAME,FMSVATTEMP.INVAMT,FMSVATTEMP.INVTAX," '4-6
        sqlstr = sqlstr & " FMSVLACC.LOCID,FMSVATTEMP.VTYPE,FMSVATTEMP.RATE,FMSTAX.TTYPE," '7-10
        sqlstr = sqlstr & " FMSTAX.ACCTVAT,FMSVATTEMP.SOURCE,FMSVATTEMP.BATCH,FMSVATTEMP.ENTRY," '11-14
        sqlstr = sqlstr & " FMSVATTEMP.MARK,FMSVATTEMP.VATCOMMENT,FMSTAX.ITEMRATE1,FMSVATTEMP.IDDIST,FMSVATTEMP.CBREF,FMSVATTEMP.RUNNO," '15-19
        sqlstr = sqlstr & " FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH,FMSVATTEMP.TRANSNBR,FMSVATTEMP.Code,FMSVATTEMP.CODETAX "

        If ComDatabase = "ORCL" Then
            sqlstr = sqlstr & " FROM FMSVATTEMP"
            sqlstr = sqlstr & "     INNER JOIN FMSTAX ON (FMSVATTEMP.CODETAX = FMSTAX.AUTHORITY) "
            sqlstr = sqlstr & "     INNER JOIN FMSVLACC ON FMSTAX.ACCTVAT = FMSVLACC.ACCTVAT"
            sqlstr = sqlstr & " WHERE FMSTAX.TTYPE = 2 AND FMSTAX.BUYERCLASS = 1"

        Else
            sqlstr = sqlstr & " FROM FMSVATTEMP"
            sqlstr = sqlstr & "     INNER JOIN FMSTAX ON (FMSVATTEMP.CODETAX = FMSTAX.AUTHORITY) "
            sqlstr = sqlstr & "     INNER JOIN FMSVLACC ON FMSTAX.ACCTVAT = FMSVLACC.ACCTVAT"
            sqlstr = sqlstr & " WHERE FMSTAX.TTYPE = 2 AND FMSTAX.BUYERCLASS = 1"
        End If

        'rs.Open(sqlstr, cnVAT)

        Call Connecttion("VAT")

        daNB = New SqlDataAdapter(sqlstr, ConV)
        dtselect = New DataTable
        dtselect.Clear()
        daNB.Fill(dtselect)
        ConV.Close()



        'If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
        WriteLog("<ProcessOE> Select FMSVATTEMP Complete")
        'End If

        For j As Integer = 0 To dtselect.Rows.Count - 1

            loCls = New clsVat
            With loCls
                .INVDATE = IIf(IsDBNull(dtselect.Rows(j)(0)), 0, dtselect.Rows(j)(0))
                .TXDATE = IIf(IsDBNull(dtselect.Rows(j)(1)), 0, dtselect.Rows(j)(1))
                .IDINVC = IIf(IsDBNull(dtselect.Rows(j)(2).ToString), "", dtselect.Rows(j)(2))
                .DOCNO = IIf(IsDBNull(dtselect.Rows(j)(3).ToString), "", dtselect.Rows(j)(3))
                .INVNAME = IIf(IsDBNull(dtselect.Rows(j)(4).ToString), "", dtselect.Rows(j)(4))
                .INVAMT = IIf(IsDBNull(dtselect.Rows(j)(5)), 0, dtselect.Rows(j)(5))
                .INVTAX = IIf(IsDBNull(dtselect.Rows(j)(6)), 0, dtselect.Rows(j)(6))
                .LOCID = IIf(IsDBNull(dtselect.Rows(j)(7)), 0, dtselect.Rows(j)(7))
                .VTYPE = IIf(IsDBNull(dtselect.Rows(j)(8)), 0, dtselect.Rows(j)(8))
                .TTYPE = IIf(IsDBNull(dtselect.Rows(j)(10)), 0, dtselect.Rows(j)(10))
                .ACCTVAT = IIf(IsDBNull(dtselect.Rows(j)(11).ToString), "", dtselect.Rows(j)(11))
                .Source = IIf(IsDBNull(dtselect.Rows(j)(12).ToString), "", dtselect.Rows(j)(12))
                .Batch = IIf(IsDBNull(dtselect.Rows(j)(13).ToString), "", dtselect.Rows(j)(13))
                .Entry = IIf(IsDBNull(dtselect.Rows(j)(14).ToString), "", dtselect.Rows(j)(14))
                .MARK = IIf(IsDBNull(dtselect.Rows(j)(15).ToString), "", dtselect.Rows(j)(15))
                .VATCOMMENT = IIf(IsDBNull(dtselect.Rows(j)(16).ToString), "", dtselect.Rows(j)(16))
                .CBRef = IIf(IsDBNull(dtselect.Rows(j)(19).ToString), "", dtselect.Rows(j)(19))
                .TRANSNBR = IIf(IsDBNull(dtselect.Rows(j)("TRANSNBR").ToString), "", dtselect.Rows(j)("TRANSNBR"))
                .Runno = IIf(IsDBNull(dtselect.Rows(j)("Runno").ToString), "", dtselect.Rows(j)("Runno"))
                .TAXID = IIf(IsDBNull(dtselect.Rows(j)("TAXID").ToString), "", dtselect.Rows(j)("TAXID"))
                .Branch = IIf(IsDBNull(dtselect.Rows(j)("BRANCH").ToString), "", dtselect.Rows(j)("BRANCH"))
                Code = IIf(IsDBNull(dtselect.Rows(j)("Code").ToString), "", dtselect.Rows(j)("Code"))
                .RATE_Renamed = IIf(IsDBNull(dtselect.Rows(j)("RATE")), 0, dtselect.Rows(j)("RATE"))
                .CODETAX = IIf(IsDBNull(dtselect.Rows(j)("CODETAX").ToString), "", dtselect.Rows(j)("CODETAX"))
            

                sqlstr = InsVat(1, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, .NEWDOCNO, .INVNAME, .INVAMT, .INVTAX, .LOCID, .VTYPE, .RATE_Renamed, .TTYPE, .ACCTVAT, .Source, .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, .Runno, .CODETAX, "", "", "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, 0)

            End With
            'cnVAT.Execute(sqlstr)

            Call Connecttion("VAT")

            com = New SqlCommand(sqlstr, ConV)
            com.CommandText = sqlstr
            com.ExecuteNonQuery()

            ConV.Close()



            'rs.options.Rs.MoveNext()
            Application.DoEvents()
        Next

        'rs.options.Rs.Close()
        'rs = Nothing
        ''cnACCPAC.Close()
        'cnVAT.Close()
        ''cnACCPAC = Nothing
        'cnVAT = Nothing
        Exit Sub
ErrHandler:
        WriteLog("<ProcessOE>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessOE> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
        WriteLog(ErrMsg)
    End Sub


    Public Sub ProcessOENOVAT()
        Dim cnACCPAC, cnVAT As ADODB.Connection
        Dim rs, rs3 As BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim strsplit() As String
        Dim losplit, i As Integer
        Dim loCls As clsVat
        Dim iffs As clsFuntion
        iffs = New clsFuntion

        On Error GoTo ErrHandler

        cnACCPAC = New ADODB.Connection
        cnVAT = New ADODB.Connection
        cnACCPAC.ConnectionTimeout = 60
        cnACCPAC.Open(ConnACCPAC)
        cnACCPAC.CommandTimeout = 3600

        cnVAT.ConnectionTimeout = 60
        cnVAT.Open(ConnVAT)
        cnVAT.CommandTimeout = 3600

        'gather data
        rs3 = New BaseClass.BaseODBCIO
        sqlstr = "SELECT HOMECUR FROM CSCOM "
        rs3.Open(sqlstr, cnACCPAC)

        rs = New BaseClass.BaseODBCIO
        If Use_POST Then
            sqlstr = "SELECT AROBL.DATEINVC,AROBL.IDINVC,AROBL.TRXTYPEID,AROBL.TRXTYPETXT," '0-3
            sqlstr &= " ARCUS.NAMECUST,ARCUS.TEXTSTRE1,AROBL.AMTINVCHC,AROBL.AMTTAXHC," '4-7
            sqlstr &= " AROBL.CNTBTCH,AROBL.CNTITEM," '8-9
            sqlstr &= " OEINVH.BILNAME , OEINVH.BILADDR1, OECRDH.BILNAME, OECRDH.BILADDR1,AROBL.IDRMIT,AROBL.DATEBUS,ARCUS.IDCUST " '10-14
            If ComDatabase = "ORCL" Then
                sqlstr = sqlstr & " FROM ((AROBL LEFT JOIN ARCUS ON AROBL.IDCUST = ARCUS.IDCUST)"
                sqlstr = sqlstr & "     LEFT JOIN OEINVH ON AROBL.IDINVC = OEINVH.INVNUMBER)"
                sqlstr = sqlstr & "     LEFT JOIN OECRDH ON AROBL.IDINVC = OECRDH.CRDNUMBER"
                sqlstr = sqlstr & " WHERE (AROBL.CODECURN<>'THB')"
                sqlstr = sqlstr & " AND (AROBL." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " >= " & DATEFROM & " AND AROBL." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " <= " & DATETO & ")"
            Else
                sqlstr = sqlstr & " FROM ((AROBL LEFT JOIN ARCUS ON AROBL.IDCUST = ARCUS.IDCUST)"
                sqlstr = sqlstr & "     LEFT JOIN OEINVH ON AROBL.IDINVC = OEINVH.INVNUMBER)"
                sqlstr = sqlstr & "     LEFT JOIN OECRDH ON AROBL.IDINVC = OECRDH.CRDNUMBER"
                sqlstr = sqlstr & " WHERE (AROBL.CODECURN<>'THB')"
                sqlstr = sqlstr & " AND (AROBL." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " >= " & DATEFROM & " AND AROBL." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " <= " & DATETO & ")"
            End If
        Else
            sqlstr = "SELECT  ARIBH.DATEINVC,ARIBH.IDINVC,ARIBH.IDTRX  AS TRXTYPEID,ARIBH.TEXTTRX AS TRXTYPETXT, ARCUS.NAMECUST,ARCUS.TEXTSTRE1 "
            sqlstr &= ",ARIBH.AMTDUEHC AS AMTINVCHC,ARIBH.TXAMT1HC AS AMTTAXHC, ARIBH.CNTBTCH,ARIBH.CNTITEM, OEINVH.BILNAME , OEINVH.BILADDR1, "
            sqlstr &= "OECRDH.BILNAME, OECRDH.BILADDR1, '' AS IDRMIT ,ARIBH.DATEBUS "
            If ComDatabase = "ORCL" Then
                sqlstr &= "FROM (( ARIBC LEFT JOIN ARIBH  ON ARIBC.CNTBTCH =ARIBH.CNTBTCH  LEFT JOIN ARCUS ON ARIBH.IDCUST = ARCUS.IDCUST) "
                sqlstr &= "LEFT JOIN OEINVH ON ARIBH.IDINVC = OEINVH.INVNUMBER) "
                sqlstr &= "LEFT JOIN OECRDH ON ARIBH.IDINVC = OECRDH.CRDNUMBER "
                sqlstr &= "WHERE (ARIBH.CODECURN <> 'THB') "
                sqlstr &= "AND (ARIBH." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " >= " & DATEFROM & " AND ARIBH." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " <= " & DATETO & ") "
            Else
                sqlstr &= "FROM (( ARIBC LEFT JOIN ARIBH  ON ARIBC.CNTBTCH = ARIBH.CNTBTCH  LEFT JOIN ARCUS ON ARIBH.IDCUST = ARCUS.IDCUST) "
                sqlstr &= "LEFT JOIN OEINVH ON ARIBH.IDINVC = OEINVH.INVNUMBER) "
                sqlstr &= "LEFT JOIN OECRDH ON ARIBH.IDINVC = OECRDH.CRDNUMBER "
                sqlstr &= "WHERE (ARIBH.CODECURN <> 'THB') "
                sqlstr &= "AND (ARIBH." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " >= " & DATEFROM & " AND ARIBH." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " <= " & DATETO & ") "
            End If
        End If

        rs.Open(sqlstr, cnACCPAC)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessOENOVAT> Select OENOVAT Complete")
        End If


        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls
                ErrMsg = "Ledger = OENOVAT IDINVC=" & Trim(rs.options.Rs.Collect(1)) & ":" & i
                ErrMsg = ErrMsg & " Batch No.=" & Trim(rs.options.Rs.Collect(8))
                ErrMsg = ErrMsg & " Entry No.=" & Trim(rs.options.Rs.Collect(9))

                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect(0)), 0, (rs.options.Rs.Collect(0))) 'AROBL.DATEINVC
                .TXDATE = IIf(IsDBNull(rs.options.Rs.Collect("DATEBUS")), 0, (rs.options.Rs.Collect("DATEBUS"))) 'AROBL.DATEINVC
                .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect(0)), 0, (rs.options.Rs.Collect(1))) 'AROBL.IDINVC
                If rs.options.Rs.Collect(2) = 14 Then 'OEINVH.BILNAME  
                    .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect(10)), "", (rs.options.Rs.Collect(10))) 'OEINVH.BILNAME
                    If .INVNAME.Trim <> "" Then
                        If ComVatName = True And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" Then 'เพิ่ม And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" กรณีที่ ชื้อ Customer ยาว แล้วต้องคีย์เพิ่มที่ช่อง TEXTSTRE1 (เจอที่ AIT)
                            .INVNAME = .INVNAME.Trim & " " & IIf(IsDBNull(rs.options.Rs.Collect(11)), "", (Trim(rs.options.Rs.Collect(11)))) 'OEINVH.BILADDR1
                        End If
                    End If

                ElseIf rs.options.Rs.Collect(2) = 34 Then 'OECRDH.BILNAME
                    .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect(12)), "", (rs.options.Rs.Collect(12))) 'OECRDH.BILNAME
                    If .INVNAME.Trim <> "" Then
                        If ComVatName = True And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" Then 'เพิ่ม And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" กรณีที่ ชื้อ Customer ยาว แล้วต้องคีย์เพิ่มที่ช่อง TEXTSTRE1 (เจอที่ AIT)
                            .INVNAME = .INVNAME.Trim & " " & IIf(IsDBNull(rs.options.Rs.Collect(13)), "", (Trim(rs.options.Rs.Collect(13)))) 'OECRDH.BILADDR1
                        End If
                    End If
                End If

                If IsDBNull(.INVNAME) Or Len(Trim(.INVNAME)) = 0 Then 'ARcus.CusName
                    .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect(4)), "", RTrim(rs.options.Rs.Collect(4))) 'ARCUS.NAMECUST
                    If .INVNAME.Trim <> "" Then
                        If ComVatName = True And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" Then 'เพิ่ม And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" กรณีที่ ชื้อ Customer ยาว แล้วต้องคีย์เพิ่มที่ช่อง TEXTSTRE1 (เจอที่ AIT)
                            .INVNAME = .INVNAME.Trim & " " & IIf(IsDBNull(rs.options.Rs.Collect(5)), "", (Trim(rs.options.Rs.Collect(5)))) 'ARCUS.TEXTSTRE1
                        End If
                    End If
                End If
                .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7)) 'AROBL.AMTTAXHC
                .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(6)), 0, rs.options.Rs.Collect(6)) - .INVTAX 'AROBL.AMTINVCHC
                If rs.options.Rs.Collect(3) = 1 Then
                    .INVAMT = .INVAMT * (-1)
                    .INVTAX = .INVTAX * (-1)
                End If
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8)) 'AROBL.CNTBTCH
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(9)), 0, rs.options.Rs.Collect(9)) 'AROBL.CNTITEM
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(14)), 0, rs.options.Rs.Collect(14)) 'AROBL.IDRMIT
                Dim loCODETAX As String = rs.options.Rs.Collect(9)
                If FindMiscCodeInARCUS(Trim(rs.options.Rs.Collect("IDCUST"))) = "0" Or FindMiscCodeInARCUS(Trim(rs.options.Rs.Collect("IDCUST"))) = "" Then
                    .TAXID = ""
                    .Branch = ""
                Else
                    .TAXID = FindTaxIDBRANCH(Trim(rs.options.Rs.Collect("IDCUST")), "AR", "TAXID")
                    .Branch = FindTaxIDBRANCH(Trim(rs.options.Rs.Collect("IDCUST")), "AR", "BRANCH")
                    .INVNAME = FindMiscCodeInARCUS(Trim(rs.options.Rs.Collect("IDCUST")))
                End If
                .VATCOMMENT = .INVAMT & "," & .IDINVC & "," & .INVDATE & "," & .INVNAME

                sqlstr = InsVat(0, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, "", .INVNAME, .INVAMT, .INVTAX, .LOCID, 1, .RATE_Renamed, .TTYPE, .ACCTVAT, "AR", .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, "", .Runno, loCODETAX, "", "", "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, 0)
            End With
            i += 1
            cnVAT.Execute(sqlstr)
            ErrMsg = ""
            Code = ""
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing

        'transfer data
        rs = New BaseClass.BaseODBCIO
        sqlstr = "SELECT FMSVATTEMP.INVDATE,FMSVATTEMP.TXDATE,FMSVATTEMP.IDINVC,FMSVATTEMP.DOCNO," '0-3
        sqlstr = sqlstr & " FMSVATTEMP.INVNAME,FMSVATTEMP.INVAMT,FMSVATTEMP.INVTAX," '4-6
        sqlstr = sqlstr & " FMSVLACC.LOCID,FMSVATTEMP.VTYPE,FMSVATTEMP.RATE,FMSTAX.TTYPE," '7-10
        sqlstr = sqlstr & " FMSTAX.ACCTVAT,FMSVATTEMP.SOURCE,FMSVATTEMP.BATCH,FMSVATTEMP.ENTRY," '11-14
        sqlstr = sqlstr & " FMSVATTEMP.MARK,FMSVATTEMP.VATCOMMENT,FMSTAX.ITEMRATE1,FMSVATTEMP.IDDIST,FMSVATTEMP.CBREF" '15-19
        sqlstr = sqlstr & " ,FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH,FMSVATTEMP.Code "
        If ComDatabase = "ORCL" Then
            sqlstr = sqlstr & " FROM FMSVATTEMP"
            sqlstr = sqlstr & "     INNER JOIN FMSVLACC ON FMSVATTEMP.ACCTVAT = FMSVLACC.ACCTVAT"
            sqlstr = sqlstr & "     INNER JOIN FMSTAX ON (FMSVATTEMP.CODETAX = FMSTAX.AUTHORITY) "
            sqlstr = sqlstr & " Where FMSTAX.TTYPE = 2"
        Else

            sqlstr = sqlstr & " FROM FMSVATTEMP"
            sqlstr = sqlstr & "     INNER JOIN FMSVLACC ON FMSVATTEMP.ACCTVAT = FMSVLACC.ACCTVAT"
            sqlstr = sqlstr & "     INNER JOIN FMSTAX ON (FMSVATTEMP.CODETAX = FMSTAX.AUTHORITY) "
            sqlstr = sqlstr & " Where FMSTAX.TTYPE = 2"
        End If
        rs.Open(sqlstr, cnVAT)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessOENOVAT> Select FMSVATTEMP Complete")
        End If


        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls
                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect(0)), 0, rs.options.Rs.Collect(0))
                .TXDATE = IIf(IsDBNull(rs.options.Rs.Collect(1)), 0, rs.options.Rs.Collect(1))
                .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", (rs.options.Rs.Collect(2)))
                .DOCNO = IIf(IsDBNull(rs.options.Rs.Collect(3)), "", (rs.options.Rs.Collect(3)))
                .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect(4)), "", (rs.options.Rs.Collect(4)))
                .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(5)), 0, rs.options.Rs.Collect(5))
                .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(6)), 0, rs.options.Rs.Collect(6))
                .LOCID = IIf(IsDBNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7))
                .VTYPE = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8))
                .TTYPE = IIf(IsDBNull(rs.options.Rs.Collect(10)), 0, rs.options.Rs.Collect(10))
                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect(11)), "", (rs.options.Rs.Collect(11)))
                .Source = IIf(IsDBNull(rs.options.Rs.Collect(12)), "", (rs.options.Rs.Collect(12)))
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(13)), "", (rs.options.Rs.Collect(13)))
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(14)), "", (rs.options.Rs.Collect(14)))
                .MARK = IIf(IsDBNull(rs.options.Rs.Collect(15)), "", (rs.options.Rs.Collect(15)))
                .VATCOMMENT = IIf(IsDBNull(rs.options.Rs.Collect(16)), "", (rs.options.Rs.Collect(16)))
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(19)), "", (rs.options.Rs.Collect(19)))
                .TAXID = IIf(IsDBNull(rs.options.Rs.Collect("TAXID")), "", (rs.options.Rs.Collect("TAXID")))
                .Branch = IIf(IsDBNull(rs.options.Rs.Collect("BRANCH")), "", (rs.options.Rs.Collect("BRANCH")))
                Code = IIf(IsDBNull(rs.options.Rs.Collect("Code")), "", (rs.options.Rs.Collect("Code")))

                If .INVAMT = 0 Then
                    .RATE_Renamed = 0
                Else
                    .RATE_Renamed = CDbl(FormatNumber((.INVTAX / .INVAMT) * 100, 0))
                End If

                sqlstr = InsVat(1, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, .NEWDOCNO, .INVNAME, .INVAMT, .INVTAX, .LOCID, .VTYPE, .RATE_Renamed, .TTYPE, .ACCTVAT, .Source, .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, "", .Runno, "", "", "", "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, 0)
                sqlstr = StrSS
            End With
            cnVAT.Execute(sqlstr)
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing
        cnACCPAC.Close()
        cnVAT.Close()
        cnACCPAC = Nothing
        cnVAT = Nothing
        Exit Sub

ErrHandler:
        WriteLog("<ProcessOENOVAT>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessOENOVAT> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
        WriteLog(ErrMsg)
    End Sub

    Public Sub ProcessARNOVAT()
        Dim cnACCPAC, cnVAT As ADODB.Connection
        Dim rs, rs3 As BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim strsplit() As String
        Dim losplit As Integer
        Dim loCls As clsVat
        Dim iffs As clsFuntion
        iffs = New clsFuntion

        On Error GoTo ErrHandler
        cnACCPAC = New ADODB.Connection
        cnVAT = New ADODB.Connection
        cnACCPAC.ConnectionTimeout = 60
        cnACCPAC.Open(ConnACCPAC)
        cnACCPAC.CommandTimeout = 3600

        cnVAT.ConnectionTimeout = 60
        cnVAT.Open(ConnVAT)
        cnVAT.CommandTimeout = 3600

        rs3 = New BaseClass.BaseODBCIO
        sqlstr = "SELECT HOMECUR FROM CSCOM "
        rs3.Open(sqlstr, cnACCPAC)

        'gather data
        rs = New BaseClass.BaseODBCIO
        If Use_POST Then
            sqlstr = "SELECT AROBL.DATEINVC,AROBL.IDINVC,AROBL.TRXTYPEID,AROBL.TRXTYPETXT," '0-3
            sqlstr = sqlstr & " ARCUS.NAMECUST,ARCUS.TEXTSTRE1,AROBL.AMTINVCHC,AROBL.AMTTAXHC," '4-7
            sqlstr = sqlstr & " AROBL.CNTBTCH,AROBL.CNTITEM,AROBL.IDRMIT,AROBL.DATEBUS ,AROBL.IDCUST" '8-9
            If ComDatabase = "ORCL" Then
                sqlstr = sqlstr & " FROM (AROBL LEFT JOIN ARCUS ON AROBL.IDCUST = ARCUS.IDCUST)"
                sqlstr = sqlstr & " WHERE (AROBL.CODECURN <>'THB')"
                sqlstr = sqlstr & " AND (AROBL." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " >= " & DATEFROM & " AND AROBL." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " <= " & DATETO & ")"
            Else
                sqlstr = sqlstr & " FROM (AROBL LEFT JOIN ARCUS ON AROBL.IDCUST = ARCUS.IDCUST)"
                sqlstr = sqlstr & " WHERE (AROBL.CODECURN <>'THB')"
                sqlstr = sqlstr & " AND (AROBL." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " >= " & DATEFROM & " AND AROBL." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " <= " & DATETO & ")"
            End If
        Else
            sqlstr = " SELECT ARIBH.DATEINVC, ARIBH.IDINVC, ARIBH.IDTRX  AS TRXTYPEID, ARIBH.TEXTTRX  AS TRXTYPETXT, ARCUS.NAMECUST, ARCUS.TEXTSTRE1, aribh.AMTDUEHC as AMTINVCHC, aribh.TXAMT1HC as AMTTAXHC, " & Environment.NewLine
            sqlstr &= "       ARIBH.CNTBTCH, ARIBH.CNTITEM, '' AS IDRMIT,ARIBH.DATEBUS ,ARCUS.IDCUST " & Environment.NewLine
            If ComDatabase = "ORCL" Then
                sqlstr &= " FROM ARIBH LEFT OUTER JOIN " & Environment.NewLine
                sqlstr &= " ARCUS ON ARIBH.IDCUST = ARCUS.IDCUST " & Environment.NewLine
                sqlstr &= " WHERE (ARIBH.CODECURN <> 'THB') AND (ARIBH." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " >= " & DATEFROM & " AND ARIBH." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " <= " & DATETO & ")) " & Environment.NewLine
            Else
                sqlstr &= " FROM ARIBH LEFT OUTER JOIN " & Environment.NewLine
                sqlstr &= " ARCUS ON ARIBH.IDCUST = ARCUS.IDCUST " & Environment.NewLine
                sqlstr &= " WHERE (ARIBH.CODECURN <> 'THB') AND (ARIBH." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " >= " & DATEFROM & " AND ARIBH." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " <= " & DATETO & ") " & Environment.NewLine
            End If
        End If
        rs.Open(sqlstr, cnACCPAC)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessARNOVAT> Select ARNOVAT Complete")
        End If



        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls
                ErrMsg = "Ledger = ARNOVAT IDINVC=" & Trim(rs.options.Rs.Collect(1))
                ErrMsg = ErrMsg & " Batch No.=" & Trim(rs.options.Rs.Collect(8))
                ErrMsg = ErrMsg & " Entry No.=" & Trim(rs.options.Rs.Collect(9))

                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect(0)), 0, (rs.options.Rs.Collect(0))) 'AROBL.DATEINVC
                .TXDATE = IIf(IsDBNull(rs.options.Rs.Collect("DATEBUS")), 0, (rs.options.Rs.Collect("DATEBUS"))) 'AROBL.DATEBUS
                .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect(1)), "", (rs.options.Rs.Collect(1))) 'AROBL.IDINVC

                .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect(4)), "", (rs.options.Rs.Collect(4))) 'ARCUS.NAMECUST
                If .INVNAME.Trim <> "" Then
                    If ComVatName = True And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" Then 'เพิ่ม And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" กรณีที่ ชื้อ Customer ยาว แล้วต้องคีย์เพิ่มที่ช่อง TEXTSTRE1 (เจอที่ AIT)

                        .INVNAME = .INVNAME.Trim & " " & IIf(IsDBNull(rs.options.Rs.Collect(5)), "", (Trim(rs.options.Rs.Collect(5)))) 'ARCUS.TEXTSTRE1
                    End If
                End If

                .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7)) 'AROBL.AMTTAXHC
                .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(6)), 0, rs.options.Rs.Collect(6)) - .INVTAX 'AROBL.AMTINVCHC
                If rs.options.Rs.Collect(3) = 1 Then
                    .INVAMT = .INVAMT * (-1)
                    .INVTAX = .INVTAX * (-1)
                End If
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8)) 'AROBL.CNTBTCH
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(9)), 0, rs.options.Rs.Collect(9)) 'AROBL.CNTITEM
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(10)), 0, rs.options.Rs.Collect(10)) 'AROBL.IDRMIT

                If FindMiscCodeInARCUS(Trim(rs.options.Rs.Collect("IDCUST"))) = "0" Or FindMiscCodeInARCUS(Trim(rs.options.Rs.Collect("IDCUST"))) = "" Then
                    .TAXID = ""
                    .Branch = ""
                Else
                    .TAXID = FindTaxIDBRANCH(Trim(rs.options.Rs.Collect("IDCUST")), "AR", "TAXID")
                    .Branch = FindTaxIDBRANCH(Trim(rs.options.Rs.Collect("IDCUST")), "AR", "BRANCH")
                    .INVNAME = FindMiscCodeInARCUS(Trim(rs.options.Rs.Collect("IDCUST")))
                End If
                .VATCOMMENT = .INVAMT & "," & .IDINVC & "," & .INVDATE & "," & .INVNAME


                sqlstr = InsVat(0, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, "", .INVNAME, .INVAMT, .INVTAX, .LOCID, 1, .RATE_Renamed, .TTYPE, .ACCTVAT, "AR", .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, "", .Runno, rs.options.Rs.Collect(10).ToString, "", "", "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, 0)
                'sqlstr = StrSS

            End With
            cnVAT.Execute(sqlstr)
            ErrMsg = ""
            Code = ""
            rs.options.Rs.MoveNext()
            Application.DoEvents()

        Loop
        rs.options.Rs.Close()
        rs = Nothing

        'transfer data
        rs = New BaseClass.BaseODBCIO
        sqlstr = "SELECT FMSVATTEMP.INVDATE,FMSVATTEMP.TXDATE,FMSVATTEMP.IDINVC,FMSVATTEMP.DOCNO," '0-3
        sqlstr = sqlstr & " FMSVATTEMP.INVNAME,FMSVATTEMP.INVAMT,FMSVATTEMP.INVTAX," '4-6
        sqlstr = sqlstr & " FMSVLACC.LOCID,FMSVATTEMP.VTYPE,FMSVATTEMP.RATE,FMSTAX.TTYPE," '7-10
        sqlstr = sqlstr & " FMSTAX.ACCTVAT,FMSVATTEMP.SOURCE,FMSVATTEMP.BATCH,FMSVATTEMP.ENTRY," '11-14
        sqlstr = sqlstr & " FMSVATTEMP.MARK,FMSVATTEMP.VATCOMMENT,FMSTAX.ITEMRATE1,FMSVATTEMP.IDDIST,FMSVATTEMP.CBREF, " '15-19
        sqlstr = sqlstr & " FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH,FMSVATTEMP.Code "
        If ComDatabase = "ORCL" Then
            sqlstr = sqlstr & " FROM FMSVATTEMP"
            sqlstr = sqlstr & "     INNER JOIN FMSVLACC ON FMSVATTEMP.ACCTVAT = FMSVLACC.ACCTVAT"
            sqlstr = sqlstr & "     INNER JOIN FMSTAX ON (FMSVATTEMP.CODETAX = FMSTAX.AUTHORITY) "
            sqlstr = sqlstr & " Where FMSTAX.TTYPE = 2"
        Else
            sqlstr = sqlstr & " FROM FMSVATTEMP"
            sqlstr = sqlstr & "     INNER JOIN FMSVLACC ON FMSVATTEMP.ACCTVAT = FMSVLACC.ACCTVAT"
            sqlstr = sqlstr & "     INNER JOIN FMSTAX ON (FMSVATTEMP.CODETAX = FMSTAX.AUTHORITY) "
            sqlstr = sqlstr & " Where FMSTAX.TTYPE = 2"
        End If
        rs.Open(sqlstr, cnVAT)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessARNOVAT> Select FMSVATTEMP Complete")
        End If


        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls
                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect(0)), 0, rs.options.Rs.Collect(0))
                .TXDATE = IIf(IsDBNull(rs.options.Rs.Collect(1)), 0, rs.options.Rs.Collect(1))
                .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", (rs.options.Rs.Collect(2)))
                .DOCNO = IIf(IsDBNull(rs.options.Rs.Collect(3)), "", (rs.options.Rs.Collect(3)))
                .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect(4)), "", (rs.options.Rs.Collect(4)))
                .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(5)), 0, rs.options.Rs.Collect(5))
                .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(6)), 0, rs.options.Rs.Collect(6))
                .LOCID = IIf(IsDBNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7))
                .VTYPE = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8))
                .TTYPE = IIf(IsDBNull(rs.options.Rs.Collect(10)), 0, rs.options.Rs.Collect(10))
                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect(11)), "", (rs.options.Rs.Collect(11)))
                .Source = IIf(IsDBNull(rs.options.Rs.Collect(12)), "", (rs.options.Rs.Collect(12)))
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(13)), "", (rs.options.Rs.Collect(13)))
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(14)), "", (rs.options.Rs.Collect(14)))
                .MARK = IIf(IsDBNull(rs.options.Rs.Collect(15)), "", (rs.options.Rs.Collect(15)))
                .VATCOMMENT = IIf(IsDBNull(rs.options.Rs.Collect(16)), "", (rs.options.Rs.Collect(16)))
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(19)), "", (rs.options.Rs.Collect(19)))
                .TAXID = IIf(IsDBNull(rs.options.Rs.Collect("TAXID")), "", (rs.options.Rs.Collect("TAXID")))
                .Branch = IIf(IsDBNull(rs.options.Rs.Collect("BRANCH")), "", (rs.options.Rs.Collect("BRANCH")))
                Code = IIf(IsDBNull(rs.options.Rs.Collect("Code")), "", (rs.options.Rs.Collect("Code")))

                If .INVAMT = 0 Then
                    .RATE_Renamed = 0
                Else
                    .RATE_Renamed = CDbl(FormatNumber((.INVTAX / .INVAMT) * 100, 0))
                End If

                sqlstr = InsVat(1, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, .NEWDOCNO, .INVNAME, .INVAMT, .INVTAX, .LOCID, .VTYPE, .RATE_Renamed, .TTYPE, .ACCTVAT, .Source, .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, "", .Runno, rs.options.Rs.Collect(10), "", "", "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, 0)
            End With
            cnVAT.Execute(sqlstr)
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing
        cnACCPAC.Close()
        cnVAT.Close()
        cnACCPAC = Nothing
        cnVAT = Nothing
        Exit Sub

ErrHandler:
        WriteLog("<ProcessARNOVAT>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessARNOVAT> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
        WriteLog(ErrMsg)
    End Sub

    Public Sub ProcessAP()
        Dim cnACCPAC, cnVAT As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr, sqlstr2, sqlstr3, TmpStr As String
        Dim loCls As clsVat
        'Dim tmpTAXBASE, tmpTAXDATE As String
        Dim iffs As clsFuntion
        Dim Ddbo As String
        iffs = New clsFuntion
        ErrMsg = ""

        On Error GoTo ErrHandler

        cnACCPAC = New ADODB.Connection
        cnVAT = New ADODB.Connection
        cnACCPAC.ConnectionTimeout = 60
        cnACCPAC.Open(ConnACCPAC)
        cnACCPAC.CommandTimeout = 3600
        cnVAT.ConnectionTimeout = 60
        cnVAT.Open(ConnVAT)
        cnVAT.CommandTimeout = 3600

        'gather data
        sqlstr = ""
        sqlstr2 = ""
        sqlstr3 = ""
        TmpStr = ""

        Ddbo = ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.")

        rs = New BaseClass.BaseODBCIO
        If Use_POST Then

            If ComDatabase = "MSSQL" Then
                sqlstr = " select * from ( " & Environment.NewLine
                sqlstr &= " SELECT DISTINCT  " & Environment.NewLine
                sqlstr &= Ddbo & "APOBLTEMP.DATEINVC, " & Ddbo & "APOBLTEMP.[VALUES], " & Ddbo & "APOBLTEMP.IDINVC, "
                If Use_LEGALNAME_inAP = True Then

                    sqlstr &= "CASE WHEN APVNR.RMITNAME IS NULL THEN APVEN.VENDNAME ELSE APVNR.RMITNAME END AS THNAME, APVEN.TEXTSTRE1, " & Ddbo & "APOBLTEMP.AUDTDATE, '' AS IDDIST, " & Ddbo & "APOBLTEMP.AMTINVCHC,  " & Environment.NewLine
                Else
                    sqlstr &= "APVEN.VENDNAME, APVEN.TEXTSTRE1, " & Ddbo & "APOBLTEMP.AUDTDATE, '' AS IDDIST, " & Ddbo & "APOBLTEMP.AMTINVCHC,  " & Environment.NewLine
                End If

            Else
                sqlstr &= " SELECT DISTINCT  " & Environment.NewLine
                sqlstr &= Ddbo & "APOBLTEMP.DATEINVC, " & Ddbo & "APOBLTEMP." & Chr(34) & "VALUES" & Chr(34) & ", " & Ddbo & "APOBLTEMP.IDINVC, APVEN.VENDNAME, APVEN.TEXTSTRE1, " & Ddbo & "APOBLTEMP.AUDTDATE, '' AS IDDIST, " & Ddbo & "APOBLTEMP.AMTINVCHC,  " & Environment.NewLine
            End If

            sqlstr &= "         SUM(" & Ddbo & "APOBLTEMP.AMTBASE1HC) AS AMTEXTNDHC, SUM(" & Ddbo & "APOBLTEMP.AMTBASE1HC) AS AMTTAXHC, ISNULL(TXAUTH.ACCTRECOV, APIBH.ACCTREC1 ) AS IDACCT,  " & Environment.NewLine
            sqlstr &= "         " & Ddbo & "APOBLTEMP.CNTBTCH, " & Ddbo & "APOBLTEMP.CNTITEM, " & Ddbo & "APOBLTEMP.CODETAX1, " & Ddbo & "APOBLTEMP.IDPONBR, MIN(" & Ddbo & "APOBLTEMP.AMTTAX1HC) AS AMTTAXREC1, 0 AS TAXBASE,  " & Environment.NewLine
            sqlstr &= "         0 AS TAXDATE, '' AS TAXNAME,MIN(" & Ddbo & "APOBLTEMP.AMTBASE1HC)  AS BASETAX1, " & Ddbo & "APOBLTEMP.TXTTRXTYPE, APVEN.NAMECITY,  " & Environment.NewLine
            sqlstr &= "         '' AS TEXTLINE, " & Ddbo & "APOBLTEMP.AMTBASE1HC AS AMTTOTTAX, '0' AS CNTLINE, " & Ddbo & "APOBLTEMP.IDVEND, " & Ddbo & "APOBLTEMP.DATEBUS, APIBH.INVCDESC ,'HEADER' AS VATAS,MIN(APIBH.AMTRECTAX) AS AMTRECTAX " & Environment.NewLine
            sqlstr &= " FROM    " & Ddbo & "APOBLTEMP " & Environment.NewLine
            'sqlstr &= " FROM    APIBC " & Environment.NewLine
            'sqlstr &= "         LEFT OUTER JOIN " & Ddbo & "APOBLTEMP ON APIBC.CNTBTCH = " & Ddbo & "APOBLTEMP.CNTBTCH " & Environment.NewLine
            sqlstr &= "         LEFT OUTER JOIN APVEN ON " & Ddbo & "APOBLTEMP.IDVEND = APVEN.VENDORID " & Environment.NewLine
            sqlstr &= "         LEFT OUTER JOIN " & Ddbo & "APPJHTEMP ON " & Ddbo & "APOBLTEMP.IDINVC = " & Ddbo & "APPJHTEMP.IDINVC AND " & Ddbo & "APOBLTEMP.IDVEND = " & Ddbo & "APPJHTEMP.IDVEND " & Environment.NewLine
            sqlstr &= "         LEFT OUTER JOIN APPJD ON " & Ddbo & "APPJHTEMP.POSTSEQNCE = APPJD.POSTSEQNCE AND " & Ddbo & "APPJHTEMP.CNTBTCH = APPJD.CNTBTCH AND " & Ddbo & "APPJHTEMP.CNTITEM = APPJD.CNTITEM " & Environment.NewLine
            sqlstr &= "         LEFT OUTER JOIN APIBH ON " & Ddbo & "APOBLTEMP.IDINVC = APIBH.IDINVC AND " & Ddbo & "APOBLTEMP.IDVEND = APIBH.IDVEND " & Environment.NewLine
            sqlstr &= "         LEFT OUTER JOIN APIBHO ON APIBH.CNTBTCH = APIBHO.CNTBTCH AND APIBH.CNTITEM = APIBHO.CNTITEM AND APIBHO.OPTFIELD = 'VOUCHER' " & Environment.NewLine
            sqlstr &= "         LEFT OUTER JOIN TXAUTH ON " & Ddbo & "APOBLTEMP.CODETAXGRP = TXAUTH.AUTHORITY " & Environment.NewLine
            sqlstr &= "         LEFT OUTER JOIN APVNR ON APVNR.IDVEND = " & Ddbo & "APOBLTEMP.IDVEND AND APVNR.IDVENDRMIT = 'TXTHAI'" & Environment.NewLine
            sqlstr &= " WHERE   (" & Ddbo & "APOBLTEMP.TXTTRXTYPE IN(1,2,3)) " & Environment.NewLine
            'หาก Use_INVFROMPO=TRUE ให้ดึง ข้อมูลยกเว้นที่ถูกส่งมาจาก PO
            If Use_INVFROMPO Then
                sqlstr &= "     AND APIBH.SRCEAPPL <> 'PO' AND APIBH.DRILLAPP <> 'PO' " & Environment.NewLine
            End If
            sqlstr &= "         AND (APIBH.ERRENTRY IS NULL OR APIBH.ERRENTRY = 0) AND (" & Ddbo & "APOBLTEMP." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " >= " & DATEFROM & " AND " & Ddbo & "APOBLTEMP." & IIf(DATEMODE = "DOCU", "DATEINVC", "DATEBUS") & " <= " & DATETO & ") " & Environment.NewLine
            sqlstr &= "         AND  (TXAUTH.ACCTRECOV IN    " & Environment.NewLine
            sqlstr &= "         (SELECT ACCTVAT " & IIf(ComDatabase = "PVSW", "", "COLLATE Thai_CI_AS") & Environment.NewLine

            sqlstr &= "         FROM " & Ddbo & "FMSVLACC))   " & Environment.NewLine
            'CASE VAT SHIPPING NOT SHOW HEADER-NONVAT 
            sqlstr &= "         AND " & Ddbo & "APOBLTEMP.CODETAX1 NOT LIKE '%NONVAT%'" & Environment.NewLine
            'sqlstr &= " AND " & Ddbo & "APOBLTEMP.AMTTAXHC<>0 " & Environment.NewLine
            If Use_LEGALNAME_inAP = True Then
                sqlstr &= " GROUP BY " & Ddbo & "APOBLTEMP.DATEINVC, " & Ddbo & "APOBLTEMP." & IIf(ComDatabase = "MSSQL", "[VALUES]", "" & Chr(34) & "VALUES" & Chr(34) & "") & ", " & Ddbo & "APOBLTEMP.IDINVC,APIBH.ACCTREC1, " & Ddbo & "APOBLTEMP.AMTBASE1HC,CASE WHEN APVNR.RMITNAME IS NULL THEN APVEN.VENDNAME ELSE APVNR.RMITNAME END, APVEN.TEXTSTRE1, " & Ddbo & "APOBLTEMP.AUDTDATE, " & Ddbo & "APOBLTEMP.AMTINVCHC,  " & Environment.NewLine
                sqlstr &= "         " & Ddbo & "APOBLTEMP.CNTBTCH, " & Ddbo & "APOBLTEMP.CNTITEM, " & Ddbo & "APOBLTEMP.CODETAX1, " & Ddbo & "APOBLTEMP.IDPONBR,  " & Environment.NewLine
                sqlstr &= "         " & Ddbo & "APOBLTEMP.TXTTRXTYPE, APVEN.NAMECITY, " & Ddbo & "APPJHTEMP.POSTSEQNCE, TXAUTH.ACCTRECOV,  " & Environment.NewLine
                sqlstr &= "         " & Ddbo & "APPJHTEMP.POSTSEQNCE, " & Ddbo & "APOBLTEMP.IDVEND, " & Ddbo & "APOBLTEMP.DATEBUS, APIBH.INVCDESC" & Environment.NewLine
            Else
                sqlstr &= " GROUP BY " & Ddbo & "APOBLTEMP.DATEINVC, " & Ddbo & "APOBLTEMP." & IIf(ComDatabase = "MSSQL", "[VALUES]", "" & Chr(34) & "VALUES" & Chr(34) & "") & ", " & Ddbo & "APOBLTEMP.IDINVC,APIBH.ACCTREC1, " & Ddbo & "APOBLTEMP.AMTBASE1HC,APVEN.VENDNAME, APVEN.TEXTSTRE1, " & Ddbo & "APOBLTEMP.AUDTDATE, " & Ddbo & "APOBLTEMP.AMTINVCHC,  " & Environment.NewLine
                sqlstr &= "         " & Ddbo & "APOBLTEMP.CNTBTCH, " & Ddbo & "APOBLTEMP.CNTITEM, " & Ddbo & "APOBLTEMP.CODETAX1, " & Ddbo & "APOBLTEMP.IDPONBR,  " & Environment.NewLine
                sqlstr &= "         " & Ddbo & "APOBLTEMP.TXTTRXTYPE, APVEN.NAMECITY, " & Ddbo & "APPJHTEMP.POSTSEQNCE, TXAUTH.ACCTRECOV,  " & Environment.NewLine
                sqlstr &= "         " & Ddbo & "APPJHTEMP.POSTSEQNCE, " & Ddbo & "APOBLTEMP.IDVEND, " & Ddbo & "APOBLTEMP.DATEBUS, APIBH.INVCDESC" & Environment.NewLine
            End If
            sqlstr &= " UNION ALL " & Environment.NewLine

            sqlstr &= " SELECT DISTINCT  " & Environment.NewLine

            If ComDatabase = "MSSQL" Then
                If Use_LEGALNAME_inAP = True Then
                    sqlstr &= Ddbo & "APOBLTEMP.DATEINVC, " & Ddbo & "APOBLTEMP.[VALUES], " & Ddbo & "APOBLTEMP.IDINVC, CASE WHEN APVNR.RMITNAME IS NULL THEN APVEN.VENDNAME ELSE APVNR.RMITNAME END AS THNAME, APVEN.TEXTSTRE1, " & Ddbo & "APOBLTEMP.AUDTDATE, '' AS IDDIST, " & Ddbo & "APOBLTEMP.AMTINVCHC,  " & Environment.NewLine
                Else
                    sqlstr &= Ddbo & "APOBLTEMP.DATEINVC, " & Ddbo & "APOBLTEMP.[VALUES], " & Ddbo & "APOBLTEMP.IDINVC, APVEN.VENDNAME, APVEN.TEXTSTRE1, " & Ddbo & "APOBLTEMP.AUDTDATE, '' AS IDDIST, " & Ddbo & "APOBLTEMP.AMTINVCHC,  " & Environment.NewLine
                End If
            Else
                sqlstr &= Ddbo & "APOBLTEMP.DATEINVC, " & Ddbo & "APOBLTEMP." & Chr(34) & "VALUES" & Chr(34) & ", " & Ddbo & "APOBLTEMP.IDINVC, APVEN.VENDNAME, APVEN.TEXTSTRE1, " & Ddbo & "APOBLTEMP.AUDTDATE, '' AS IDDIST, " & Ddbo & "APOBLTEMP.AMTINVCHC,  " & Environment.NewLine
            End If

            sqlstr &= "         SUM(" & Ddbo & "APOBLTEMP.AMTBASE1HC) AS AMTEXTNDHC, SUM(APIBD.AMTTAXREC1) AS AMTTAXHC, ISNULL(TXAUTH.ACCTRECOV, APIBD.IDGLACCT) AS IDACCT,  " & Environment.NewLine
            sqlstr &= "         " & Ddbo & "APOBLTEMP.CNTBTCH, " & Ddbo & "APOBLTEMP.CNTITEM, " & Ddbo & "APOBLTEMP.CODETAX1, " & Ddbo & "APOBLTEMP.IDPONBR, SUM(APIBD.AMTTAXREC1) AS AMTTAXREC1, 0 AS TAXBASE,  " & Environment.NewLine
            sqlstr &= "         0 AS TAXDATE, '' AS TAXNAME, SUM(APPJD.AMTEXTNDTC) AS BASETAX1, " & Ddbo & "APOBLTEMP.TXTTRXTYPE, APVEN.NAMECITY,  " & Environment.NewLine
            sqlstr &= "         APIBT.TEXTLINE, (CASE WHEN APIBT.CNTLINE IS NULL THEN SUM(APIBD.AMTTOTTAX) ELSE " & Environment.NewLine
            sqlstr &= "         (SELECT TOP 1 APPJD.AMTEXTNDHC " & Environment.NewLine
            sqlstr &= "         FROM APPJD " & Environment.NewLine
            sqlstr &= "         WHERE POSTSEQNCE = " & Ddbo & "APPJHTEMP.POSTSEQNCE AND CNTBTCH = " & Ddbo & "APOBLTEMP.CNTBTCH AND CNTITEM = " & Ddbo & "APOBLTEMP.CNTITEM AND CNTLINE = APIBT.CNTLINE) END) AS AMTTOTTAX, " & Environment.NewLine
            sqlstr &= "         APPJD.CNTLINE, " & Ddbo & "APOBLTEMP.IDVEND, " & Ddbo & "APOBLTEMP.DATEBUS, APIBH.INVCDESC ,'DETAIL' AS VATAS,MIN(APIBH.AMTRECTAX) AS AMTRECTAX " & Environment.NewLine
            sqlstr &= " FROM    " & Ddbo & "APOBLTEMP" & Environment.NewLine
            'sqlstr &= " FROM    APIBC " & Environment.NewLine
            'sqlstr &= "         LEFT OUTER JOIN " & Ddbo & "APOBLTEMP ON APIBC.CNTBTCH = " & Ddbo & "APOBLTEMP.CNTBTCH " & Environment.NewLine
            sqlstr &= "         LEFT OUTER JOIN APVEN ON " & Ddbo & "APOBLTEMP.IDVEND = APVEN.VENDORID " & Environment.NewLine
            sqlstr &= "         LEFT OUTER JOIN " & Ddbo & "APPJHTEMP ON " & Ddbo & "APOBLTEMP.IDINVC = " & Ddbo & "APPJHTEMP.IDINVC AND " & Ddbo & "APOBLTEMP.IDVEND = " & Ddbo & "APPJHTEMP.IDVEND " & Environment.NewLine
            sqlstr &= "         LEFT OUTER JOIN APPJD ON " & Ddbo & "APPJHTEMP.POSTSEQNCE = APPJD.POSTSEQNCE AND " & Ddbo & "APPJHTEMP.CNTBTCH = APPJD.CNTBTCH AND " & Ddbo & "APPJHTEMP.CNTITEM = APPJD.CNTITEM " & Environment.NewLine
            sqlstr &= "         LEFT OUTER JOIN APIBH ON " & Ddbo & "APOBLTEMP.IDINVC = APIBH.IDINVC AND " & Ddbo & "APOBLTEMP.IDVEND = APIBH.IDVEND " & Environment.NewLine
            sqlstr &= "         LEFT OUTER JOIN APIBD ON APIBH.CNTBTCH = APIBD.CNTBTCH AND APIBH.CNTITEM = APIBD.CNTITEM  and APIBD.CNTLINE =APPJD.CNTLINE " & Environment.NewLine
            sqlstr &= "         LEFT OUTER JOIN APIBHO ON APIBH.CNTBTCH = APIBHO.CNTBTCH AND APIBH.CNTITEM = APIBHO.CNTITEM AND APIBHO.OPTFIELD = 'VOUCHER' " & Environment.NewLine
            sqlstr &= "         LEFT OUTER JOIN TXAUTH ON " & Ddbo & "APOBLTEMP.CODETAXGRP = TXAUTH.AUTHORITY " & Environment.NewLine
            sqlstr &= "         LEFT OUTER JOIN APIBT ON APPJD.CNTBTCH = APIBT.CNTBTCH AND APPJD.CNTITEM = APIBT.CNTITEM AND APPJD.CNTLINE = APIBT.CNTLINE " & Environment.NewLine
            sqlstr &= "         LEFT OUTER JOIN APVNR ON APVNR.IDVEND = " & Ddbo & "APOBLTEMP.IDVEND AND APVNR.IDVENDRMIT = 'TXTHAI' " & Environment.NewLine
            sqlstr &= " WHERE   (" & Ddbo & "APOBLTEMP.TXTTRXTYPE IN(1,2,3)) AND (" & Ddbo & "APOBLTEMP.AMTTAXHC = 0) " & Environment.NewLine
            'หาก Use_INVFROMPO=TRUE ให้ดึง ข้อมูลยกเว้นที่ถูกส่งมาจาก PO
            If Use_INVFROMPO Then
                sqlstr &= "     AND APIBH.SRCEAPPL <>'PO' AND APIBH.DRILLAPP <>'PO' " & Environment.NewLine
            End If

            sqlstr &= "         AND (APIBD.TAXCLASS1 IS NULL OR APIBD.TAXCLASS1 = 1) AND (APIBH.ERRENTRY IS NULL OR APIBH.ERRENTRY = 0) AND (" & Ddbo & "APOBLTEMP.DATEBUS >= " & DATEFROM & " AND " & Ddbo & "APOBLTEMP.DATEBUS <= " & DATETO & ") " & Environment.NewLine
            sqlstr &= "         AND (APPJD.IDACCT IN " & Environment.NewLine
            sqlstr &= "         (SELECT ACCTVAT " & IIf(ComDatabase = "PVSW", "", "COLLATE Thai_CI_AS") & Environment.NewLine
            sqlstr &= "         FROM " & Ddbo & "FMSVLACC))   " & Environment.NewLine
            If Use_LEGALNAME_inAP = True Then
                sqlstr &= " GROUP BY " & Ddbo & "APOBLTEMP.DATEINVC, " & Ddbo & "APOBLTEMP." & IIf(ComDatabase = "MSSQL", "[VALUES]", "" & Chr(34) & "VALUES" & Chr(34) & "") & ", " & Ddbo & "APOBLTEMP.IDINVC, APIBD.IDGLACCT, CASE WHEN APVNR.RMITNAME IS NULL THEN APVEN.VENDNAME ELSE APVNR.RMITNAME END, APVEN.TEXTSTRE1, " & Ddbo & "APOBLTEMP.AUDTDATE, " & Ddbo & "APOBLTEMP.AMTINVCHC,  " & Environment.NewLine
                sqlstr &= "         " & Ddbo & "APOBLTEMP.CNTBTCH, " & Ddbo & "APOBLTEMP.CNTITEM, " & Ddbo & "APOBLTEMP.CODETAX1, " & Ddbo & "APOBLTEMP.IDPONBR,  " & Environment.NewLine
                sqlstr &= "         " & Ddbo & "APOBLTEMP.TXTTRXTYPE, APVEN.NAMECITY, " & Ddbo & "APPJHTEMP.POSTSEQNCE, TXAUTH.ACCTRECOV, APIBT.TEXTLINE, APIBT.CNTLINE," & IIf(OPT_AP, "APIBD.CNTLINE,", "") & " " & Environment.NewLine
                sqlstr &= "         " & Ddbo & "APPJHTEMP.POSTSEQNCE, " & Ddbo & "APOBLTEMP.IDVEND, " & Ddbo & "APOBLTEMP.DATEBUS, APIBH.INVCDESC , APPJD.CNTLINE" & Environment.NewLine

            Else
                sqlstr &= " GROUP BY " & Ddbo & "APOBLTEMP.DATEINVC, " & Ddbo & "APOBLTEMP." & IIf(ComDatabase = "MSSQL", "[VALUES]", "" & Chr(34) & "VALUES" & Chr(34) & "") & ", " & Ddbo & "APOBLTEMP.IDINVC, APIBD.IDGLACCT, APVEN.VENDNAME, APVEN.TEXTSTRE1, " & Ddbo & "APOBLTEMP.AUDTDATE, " & Ddbo & "APOBLTEMP.AMTINVCHC,  " & Environment.NewLine
                sqlstr &= "         " & Ddbo & "APOBLTEMP.CNTBTCH, " & Ddbo & "APOBLTEMP.CNTITEM, " & Ddbo & "APOBLTEMP.CODETAX1, " & Ddbo & "APOBLTEMP.IDPONBR,  " & Environment.NewLine
                sqlstr &= "         " & Ddbo & "APOBLTEMP.TXTTRXTYPE, APVEN.NAMECITY, " & Ddbo & "APPJHTEMP.POSTSEQNCE, TXAUTH.ACCTRECOV, APIBT.TEXTLINE, APIBT.CNTLINE," & IIf(OPT_AP, "APIBD.CNTLINE,", "") & " " & Environment.NewLine
                sqlstr &= "         " & Ddbo & "APPJHTEMP.POSTSEQNCE, " & Ddbo & "APOBLTEMP.IDVEND, " & Ddbo & "APOBLTEMP.DATEBUS, APIBH.INVCDESC , APPJD.CNTLINE" & Environment.NewLine
            End If

            If ComDatabase = "MSSQL" Then
                sqlstr &= " )AS TMP1  Order by DATEINVC,CNTBTCH ,CNTITEM ,CNTLINE  " & Environment.NewLine
            End If

        End If
        rs.Open(sqlstr, cnACCPAC)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessAP> Select AP Complete")
        End If



        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls
                ErrMsg = "Ledger = AP IDINVC=" & Trim(rs.options.Rs.Collect(2))
                ErrMsg = ErrMsg & " Batch No.=" & Trim(rs.options.Rs.Collect(11))
                ErrMsg = ErrMsg & " Entry No.=" & Trim(rs.options.Rs.Collect(12))
                Dim INVTAXtemp As Decimal
                Dim TEXTLINE() As String = {}
                TEXTLINE = rs.options.Rs.Collect("TEXTLINE").ToString.Split(",")
                .VATCOMMENT = rs.options.Rs.Collect("TEXTLINE").ToString.Trim
                Code = ""
                '.Batch = IIf(IsDBNull(rs.options.Rs.Collect(11)), "", (rs.options.Rs.Collect(11))) 'APOBL.CNTBTCH
                '.Entry = IIf(IsDBNull(rs.options.Rs.Collect(12)), "", (rs.options.Rs.Collect(12))) 'APOBL.CNTITEM

                'If .Batch = 45 And .Entry = 1 Then
                '    Dim Asss As String = ""
                'End If

                If rs.options.Rs.Collect("VATAS").ToString.Trim = "HEADER" Then

                    Dim RS_WHT_HEADER As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from APIBHO  where OPTFIELD in ('TAXBASE','TAXINVNO','TAXDATE','TAXNAME') and CNTBTCH ='" & rs.options.Rs.Collect("CNTBTCH") & "' and CNTITEM='" & rs.options.Rs.Collect("CNTITEM") & "'", cnACCPAC)
                    .TXDATE = rs.options.Rs.Collect("DATEBUS")
                    .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", Trim(rs.options.Rs.Collect(2))) 'APOBL.IDINVC
                    .DOCNO = rs.options.Rs.Collect("IDPONBR")

                    If Len(Trim(.INVNAME)) = 0 Then ' Or Trim(tmpTAXNAME) = "0" Then
                        .INVNAME = IIf(iffs.ISNULL(rs.options.Rs, 3), "", Replace(RTrim(rs.options.Rs.Collect(3)), "'", "''")) 'APVEN.VENDNAME
                        Code = IIf(IsDBNull(rs.options.Rs.Collect("IDVEND")), "", (rs.options.Rs.Collect("IDVEND")))
                    End If

                    If IsDBNull(rs.options.Rs.Collect(6)) Then Exit Do

                    If Trim(rs.options.Rs.Collect(6)) = "VAT" Then

                        .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7)) ' APOBL.AMTINVCHC
                        .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8)) ' APOBL.AMTEXTNDHC
                    Else

                        If .INVAMT = 0 Then
                            .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(19)), 0, rs.options.Rs.Collect(19)) ' APOBL.AMTINVCHC
                        End If

                        INVTAXtemp = IIf(IsDBNull(rs.options.Rs.Collect(15)), 0, rs.options.Rs.Collect(15)) ' APOBL.AMTTAX1HC
                        .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect("AMTRECTAX")), 0, rs.options.Rs.Collect("AMTRECTAX")) 'MIN(APIBH.AMTRECTAX) AS AMTRECTAX

                        '--------- แก้ไขให้รองรับ VAT AVG เนื่องจากเปลี่ยนฟิลด์หยิบเป็นฟิลด์ (APIBH.AMTRECTAX) และฟิลด์ที่หยิบจะได้ตัวเลขที่เป็น บวกทั้งหมด แต่ตัวเลขต้องแสดงตามฟิลด์เดิม (APOBL.AMTTAX1HC)จึงต้องเช็คเงื่อนไขเพื่อ (*-1) ------
                        If INVTAXtemp < 0 Then
                            .INVTAX = .INVTAX * (-1)
                        End If
                        '-------- end comment -----

                        If .INVTAX = 0 Then
                            .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(15)), 0, rs.options.Rs.Collect(15)) ' APOBL.AMTTAX1HC
                        End If


                        If rs.options.Rs.Collect(20) = 3 And .INVTAX > 0 Then '----------- If TXTTRXTYPE = 3 (Credit Note)
                            .INVTAX = .INVTAX * (-1)
                        End If

                    End If

                    If .INVTAX < 0 And .INVAMT > 0 Then
                        .INVAMT = .INVAMT * (-1)
                    End If



                    If OPT_AP = True Then ' ดึง VAT จาก OPT Field
                        Dim tINVAMT As Double = 0
                        Dim tTAXINVNO As String = ""
                        Dim tTAXDATE As String = ""
                        Dim tTAXNAME As String = ""
                        For iwht As Integer = 0 To RS_WHT_HEADER.options.QueryDT.Rows.Count - 1
                            If RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXBASE" Then
                                Dim Trydec As String = Trim(RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("VALUE"))
                                If Decimal.TryParse(Trydec, 0) = True Then
                                    If CDec(Trydec) <> 0 Then
                                        tINVAMT = CDec(Trydec)
                                    Else
                                        tINVAMT = 0
                                    End If
                                Else
                                    tINVAMT = 0
                                End If
                            ElseIf RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXINVNO" Then
                                If Trim(RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    tTAXINVNO = Trim(RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If
                            ElseIf RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXDATE" Then
                                Dim myDate As String = Trim(RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("VALUE"))

                                If IsNumeric(myDate) = True Then

                                    If Decimal.TryParse(myDate, 0) = True Then
                                        If CDec(myDate) <> 0 Then
                                            tTAXDATE = CDec(myDate)
                                        End If
                                    End If
                                Else

                                    tTAXDATE = TryDate2(myDate, "dd/MM/yyyy", "yyyyMMdd")

                                End If

                            ElseIf RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXNAME" Then
                                If Trim(RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    tTAXNAME = Trim(RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If
                            End If
                        Next

                        If tINVAMT <> 0 Then
                            .INVAMT = tINVAMT
                        End If
                        If tTAXINVNO <> "" Then
                            .IDINVC = tTAXINVNO
                        End If
                        If tTAXDATE <> "" Then
                            .INVDATE = tTAXDATE
                        End If
                        If tTAXNAME <> "" Then
                            .INVNAME = tTAXNAME
                        End If
                    End If 'end if OPT_AP = True



                    If VNO_AP = True Then

                        Dim VNO_Header As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from APIBHO  where OPTFIELD in ('APINVNO') and CNTBTCH ='" & rs.options.Rs.Collect("CNTBTCH") & "' and CNTITEM='" & rs.options.Rs.Collect("CNTITEM") & "'", cnACCPAC)
                        For iwht As Integer = 0 To VNO_Header.options.QueryDT.Rows.Count - 1
                            If VNO_Header.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "APINVNO" Then
                                If Trim(VNO_Header.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    .DOCNO = Trim(VNO_Header.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If

                            End If
                        Next

                    End If


                ElseIf rs.options.Rs.Collect("VATAS").ToString.Trim = "DETAIL" Then


                    If Trim(rs.options.Rs.Collect(6)) = "VAT" Then

                        .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7)) ' APOBL.AMTINVCHC
                        .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8)) ' APOBL.AMTEXTNDHC
                    Else

                        If .INVAMT = 0 Then
                            .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(19)), 0, rs.options.Rs.Collect(19)) ' APOBL.AMTINVCHC
                        End If

                        .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect("BASETAX1")), 0, rs.options.Rs.Collect("BASETAX1")) ' SUM(APPJD.AMTEXTNDTC) AS BASETAX1 

                        If rs.options.Rs.Collect(20) = 3 And .INVTAX > 0 Then '----------- If TXTTRXTYPE = 3 (Credit Note)
                            .INVTAX = .INVTAX * (-1)
                        End If

                    End If

                    If .INVTAX < 0 And .INVAMT > 0 Then
                        .INVAMT = .INVAMT * (-1)
                    End If

                    .IDINVC = rs.options.Rs.Collect("IDINVC")
                    .TXDATE = rs.options.Rs.Collect("DATEBUS")
                    .INVDATE = rs.options.Rs.Collect("DATEINVC")
                    .DOCNO = rs.options.Rs.Collect("IDPONBR")
                    .TRANSNBR = rs.options.Rs.Collect("CNTLINE")
                    .INVAMT = rs.options.Rs.Collect("TAXBASE")
                    INVTAXtemp = 0

                    If OPT_AP = False Then 'เอาจาก Comment ใน Detail
                        For i As Integer = 0 To TEXTLINE.Length - 1
                            If IsDBNull(TEXTLINE(i)) = True Then
                                TEXTLINE(i) = ""
                            End If
                        Next
                        If TEXTLINE.Length >= 1 Then 'ฐาน
                            If TEXTLINE(0).Trim <> "" And Decimal.TryParse(TEXTLINE(0), 0) Then
                                .INVAMT = CDec(TEXTLINE(0))
                            End If
                        End If

                        If TEXTLINE.Length >= 2 Then 'เลขที่
                            If TEXTLINE(1).Trim <> "" Then
                                .IDINVC = TEXTLINE(1)
                            End If
                        End If

                        If TEXTLINE.Length >= 3 Then 'วันที่
                            If TEXTLINE(2).Trim <> "" Then
                                .INVDATE = TryDate2(ZeroDate(TEXTLINE(2), "dd/MM/yyyy", "yyyyMMdd"), "dd/MM/yyyy", "yyyyMMdd")
                                If .INVDATE = 0 Then
                                    .INVDATE = TryDate2(TEXTLINE(2), "dd/MM/yyyy", "yyyyMMdd")
                                End If
                            End If
                        End If

                        If TEXTLINE.Length >= 4 Then 'เลข Vendor Code หรือ ชื่อ Vendor
                            If TEXTLINE(3).Trim = "" Then
                                '.INVNAME = rs.options.Rs.Collect("VENDNAME")
                                .INVNAME = rs.options.Rs.Collect("IDVEND")
                            Else
                                .INVNAME = TEXTLINE(3)
                                For CName As Integer = 4 To TEXTLINE.Length - 1
                                    .INVNAME = .INVNAME & "," & TEXTLINE(CName)
                                Next
                            End If
                        End If

                    Else ' ดึง VAT จาก OPT Field

                        Dim RS_WHT_DETAIL As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from APIBDO  where OPTFIELD in ('TAXBASE','TAXINVNO','TAXDATE','TAXNAME','TAXID','BRANCH') and CNTBTCH ='" & rs.options.Rs.Collect("CNTBTCH") & "' and CNTITEM='" & rs.options.Rs.Collect("CNTITEM") & "' and CNTLINE ='" & IIf(IsDBNull(rs.options.Rs.Collect("CNTLINE")), "0", rs.options.Rs.Collect("CNTLINE")) & "' ", cnACCPAC)

                        Dim tINVAMT As Double = 0
                        Dim tTAXINVNO As String = ""
                        Dim tTAXDATE As String = ""
                        Dim tTAXNAME As String = ""
                        Dim tTAXID As String = ""
                        Dim tBRANCH As String = ""

                        For iwht As Integer = 0 To RS_WHT_DETAIL.options.QueryDT.Rows.Count - 1
                            If RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXBASE" Then
                                Dim Trydec As String = Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE"))
                                If Decimal.TryParse(Trydec, 0) = True Then
                                    If CDec(Trydec) <> 0 Then
                                        tINVAMT = CDec(Trydec)
                                    Else
                                        tINVAMT = 0
                                    End If
                                Else
                                    tINVAMT = 0
                                End If
                            ElseIf RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXINVNO" Then
                                If Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    tTAXINVNO = Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If
                            ElseIf RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXDATE" Then
                                Dim myDate As String = Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE"))

                                If IsNumeric(myDate) = True Then

                                    If Decimal.TryParse(myDate, 0) = True Then
                                        If CDec(myDate) <> 0 Then
                                            tTAXDATE = CDec(myDate)
                                        End If
                                    End If
                                Else

                                    tTAXDATE = TryDate2(myDate, "dd/MM/yyyy", "yyyyMMdd")

                                End If

                            ElseIf RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXNAME" Then
                                If Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    Dim chkIDVEND_Opt As New BaseClass.BaseODBCIO("select VENDORID,APVNR.RMITNAME AS RMITNAME from APVEN inner join APVNR on APVEN.VENDORID = APVNR.IDVEND where IDVENDRMIT = 'TXTHAI' and VENDORID = '" & Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE")) & "'", cnACCPAC)
                                    If chkIDVEND_Opt.options.QueryDT.Rows.Count = 0 Then
                                        tTAXNAME = Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE"))
                                    Else
                                        tTAXNAME = chkIDVEND_Opt.options.QueryDT.Rows(0).Item("RMITNAME").ToString
                                    End If
                                End If

                                ElseIf RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXID" Then
                                    If Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                        tTAXID = Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE"))
                                    End If

                                ElseIf RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "BRANCH" Then
                                    If Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                        tBRANCH = Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE"))
                                    End If

                                End If
                        Next

                        If tINVAMT <> 0 Then
                            .INVAMT = tINVAMT
                        Else
                            .INVAMT = 0
                        End If
                        If tTAXINVNO <> "" Then
                            .IDINVC = tTAXINVNO
                        End If
                        If tTAXDATE <> "" Then
                            .INVDATE = tTAXDATE
                        End If
                        If tTAXNAME <> "" Then
                            .INVNAME = tTAXNAME
                        End If

                        If tTAXID <> "" Then
                            .TAXID = tTAXID
                        End If
                        If tBRANCH <> "" Then
                            .Branch = tBRANCH
                        End If

                    End If



                    If VNO_AP = True Then

                        Dim VNO_Detail As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from APIBDO  where OPTFIELD in ('APINVNO') and CNTBTCH ='" & rs.options.Rs.Collect("CNTBTCH") & "' and CNTITEM='" & rs.options.Rs.Collect("CNTITEM") & "'", cnACCPAC)
                        For iwht As Integer = 0 To VNO_Detail.options.QueryDT.Rows.Count - 1
                            If VNO_Detail.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "APINVNO" Then
                                If Trim(VNO_Detail.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    .DOCNO = Trim(VNO_Detail.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If

                            End If
                        Next

                    End If


                End If 'end if "DETAIL" "HEADER"

                Dim vDOCNO As String = rs.options.Rs.Collect("INVCDESC").ToString.Trim()
                If vDOCNO.IndexOf(",") > -1 Then
                    .DOCNO = vDOCNO.Substring(vDOCNO.IndexOf(",") + 1, vDOCNO.Length - vDOCNO.IndexOf(",") - 1)
                End If

                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect(10)), "", rs.options.Rs.Collect(10)) 'APATR.IDAPACCT
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(11)), "", Trim(rs.options.Rs.Collect(11))) 'APOBL.CNTBTCH
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(12)), "", Trim(rs.options.Rs.Collect(12))) 'APOBL.CNTITEM
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(14)), "", Trim(rs.options.Rs.Collect(14))) 'APOBL.IDPONBR
                .TypeOfPU = IIf(IsDBNull(rs.options.Rs.Collect("NAMECITY")), "", Trim(rs.options.Rs.Collect("NAMECITY")))

                If CDec(rs.options.Rs.Collect("AMTINVCHC")) < 0 Then
                    .INVAMT = Math.Abs(.INVAMT) * -1
                End If

                If .INVAMT = 0 Then
                    .RATE_Renamed = 0
                Else
                    .RATE_Renamed = FormatNumber(((.INVTAX / .INVAMT) * 100), 1)
                End If


                Dim f13 As String = IIf(ISNULL(rs.options.Rs, 13), "", Trim(rs.options.Rs.Collect(13)))
                Dim f6 As String = IIf(ISNULL(rs.options.Rs, 6), "", Trim(rs.options.Rs.Collect(6)))

                If CStr(.INVDATE) = 0 Then
                    .INVDATE = Trim(rs.options.Rs.Collect("DATEINVC"))
                End If
                If CStr(.TXDATE) = 0 Then
                    .TXDATE = Trim(rs.options.Rs.Collect("DATEBUS"))
                End If

                If FindMiscCodeInAPVEN(Trim(.INVNAME), "AP") = "0" Or FindMiscCodeInAPVEN(Trim(.INVNAME), "AP") = "" Then
                    If Trim(.INVNAME) = "" Then
                        .INVNAME = FindMiscCodeInAPVEN(Trim(rs.options.Rs.Collect("IDVEND")), "AP")
                    Else
                        .INVNAME = Trim(.INVNAME)
                    End If
                    If .TAXID = "" And .Branch = "" Then
                        .TAXID = FindTaxIDBRANCH(Trim(rs.options.Rs.Collect("IDVEND")), "AP", "TAXID") 'Vendor Code ของ Entry
                        .Branch = FindTaxIDBRANCH(Trim(rs.options.Rs.Collect("IDVEND")), "AP", "BRANCH")
                    End If
                Else
                    If .TAXID = "" And .Branch = "" Then
                        .TAXID = FindTaxIDBRANCH(.INVNAME, "AP", "TAXID")
                        .Branch = FindTaxIDBRANCH(.INVNAME, "AP", "BRANCH")
                        .INVNAME = FindMiscCodeInAPVEN(Trim(.INVNAME), "AP")
                    End If

                End If

                .VATCOMMENT = .INVAMT & "," & .IDINVC & "," & .INVDATE & "," & .INVNAME
                .MARK = "Invoice"

                sqlstr = InsVat(0, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, "", .INVNAME, .INVAMT, .INVTAX, .LOCID, 1, .RATE_Renamed, .TTYPE, .ACCTVAT, "AP", .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, .Runno, f13, f6, .TypeOfPU, "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, INVTAXtemp)

            End With



            cnVAT.Execute(sqlstr)
            ErrMsg = ""
            Code = ""
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing

        'transfer data
        rs = New BaseClass.BaseODBCIO
        sqlstr = "SELECT FMSVATTEMP.INVDATE,FMSVATTEMP.TXDATE,FMSVATTEMP.IDINVC,FMSVATTEMP.DOCNO," '0-3
        sqlstr = sqlstr & " FMSVATTEMP.INVNAME,FMSVATTEMP.INVAMT,FMSVATTEMP.INVTAX," '4-6
        sqlstr = sqlstr & " FMSVLACC.LOCID,FMSVATTEMP.VTYPE,FMSVATTEMP.RATE,FMSTAX.TTYPE," '7-10
        sqlstr = sqlstr & " FMSTAX.ACCTVAT,FMSVATTEMP.SOURCE,FMSVATTEMP.BATCH,FMSVATTEMP.ENTRY," '11-14
        sqlstr = sqlstr & " FMSVATTEMP.MARK,FMSVATTEMP.VATCOMMENT,FMSTAX.ITEMRATE1,FMSVATTEMP.IDDIST,FMSVATTEMP.CBREF ,FMSVATTEMP.typeofpu,FMSVATTEMP.tranno,FMSVATTEMP.cif,FMSVATTEMP.taxcif,FMSVATTEMP.TRANSNBR, " '15-20
        sqlstr = sqlstr & " FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH,FMSVATTEMP.Code ,FMSVATTEMP.TOTTAX,FMSVATTEMP.CODETAX "

        If ComDatabase = "ORCL" Then
            sqlstr = sqlstr & " FROM FMSVATTEMP"
            sqlstr = sqlstr & " INNER JOIN FMSTAX ON (FMSVATTEMP.CODETAX = FMSTAX.AUTHORITY) and  (FMSTAX.ACCTVAT= FMSVATTEMP.ACCTVAT) "
            sqlstr = sqlstr & " INNER JOIN FMSVLACC ON FMSTAX.ACCTVAT = FMSVLACC.ACCTVAT"
            sqlstr = sqlstr & " Where FMSTAX.TTYPE IN(1,3) AND FMSTAX.BUYERCLASS = 1 "
        Else
            sqlstr = sqlstr & " FROM FMSVATTEMP"
            sqlstr = sqlstr & " INNER JOIN FMSTAX ON (FMSVATTEMP.CODETAX = FMSTAX.AUTHORITY) and  (FMSTAX.ACCTVAT= FMSVATTEMP.ACCTVAT) "
            sqlstr = sqlstr & " INNER JOIN FMSVLACC ON FMSTAX.ACCTVAT = FMSVLACC.ACCTVAT"
            sqlstr = sqlstr & " Where FMSTAX.TTYPE IN(1,3) AND FMSTAX.BUYERCLASS = 1 "
        End If

        rs.Open(sqlstr, cnVAT)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessAP> Select FMSVATTEMP Complete")
        End If


        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls

                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect(0)), 0, rs.options.Rs.Collect(0))
                .TXDATE = IIf(IsDBNull(rs.options.Rs.Collect(1)), 0, rs.options.Rs.Collect(1))
                .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", (rs.options.Rs.Collect(2)))
                .DOCNO = IIf(IsDBNull(rs.options.Rs.Collect(3)), "", (rs.options.Rs.Collect(3)))
                .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect(4)), "", Replace((rs.options.Rs.Collect(4)), "'", "''"))
                .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(5)), 0, rs.options.Rs.Collect(5))
                .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(6)), 0, rs.options.Rs.Collect(6))
                .LOCID = IIf(IsDBNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7))
                .VTYPE = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8))
                .TTYPE = IIf(IsDBNull(rs.options.Rs.Collect(10)), 0, rs.options.Rs.Collect(10))
                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect(11)), "", (rs.options.Rs.Collect(11)))
                .Source = IIf(IsDBNull(rs.options.Rs.Collect(12)), "", (rs.options.Rs.Collect(12)))
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(13)), "", (rs.options.Rs.Collect(13)))
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(14)), "", (rs.options.Rs.Collect(14)))
                .MARK = IIf(IsDBNull(rs.options.Rs.Collect(15)), "", (rs.options.Rs.Collect(15)))
                .VATCOMMENT = IIf(IsDBNull(rs.options.Rs.Collect(16)), "", (rs.options.Rs.Collect(16)))
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(19)), "", (rs.options.Rs.Collect(19)))
                .TRANSNBR = IIf(IsDBNull(rs.options.Rs.Collect("TRANSNBR")), "", (rs.options.Rs.Collect("TRANSNBR")))
                .TypeOfPU = IIf(IsDBNull(rs.options.Rs.Collect("TypeOfPU")), "", (rs.options.Rs.Collect("TypeOfPU")))
                .TranNo = IIf(IsDBNull(rs.options.Rs.Collect("TranNo")), "", (rs.options.Rs.Collect("TranNo")))
                .CIF = IIf(IsDBNull(rs.options.Rs.Collect("CIF")), "0", (rs.options.Rs.Collect("CIF")))
                .TaxCIF = IIf(IsDBNull(rs.options.Rs.Collect("TaxCIF")), "0", (rs.options.Rs.Collect("TaxCIF")))
                .TAXID = IIf(IsDBNull(rs.options.Rs.Collect("TAXID")), "", (rs.options.Rs.Collect("TAXID")))
                .Branch = IIf(IsDBNull(rs.options.Rs.Collect("BRANCH")), "", (rs.options.Rs.Collect("BRANCH")))
                Code = IIf(IsDBNull(rs.options.Rs.Collect("Code")), "", (rs.options.Rs.Collect("Code")))
                .TOTTAX = IIf(IsDBNull(rs.options.Rs.Collect("TOTTAX")), 0, (rs.options.Rs.Collect("TOTTAX")))
                .RATE_Renamed = IIf(IsDBNull(rs.options.Rs.Collect("RATE")), 0, (rs.options.Rs.Collect("RATE")))
                .CODETAX = IIf(IsDBNull(rs.options.Rs.Collect("CODETAX")), "", (rs.options.Rs.Collect("CODETAX")))


                sqlstr = InsVat(1, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, .NEWDOCNO, .INVNAME, .INVAMT, .INVTAX, .LOCID, .VTYPE, .RATE_Renamed, .TTYPE, .ACCTVAT, .Source, .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, .Runno, .CODETAX, "", .TypeOfPU, .TranNo, .CIF, .TaxCIF, 0, .TAXID, .Branch, Code, "", 0, 0, .TOTTAX)

            End With
            cnVAT.Execute(sqlstr)
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing
        cnACCPAC.Close()
        cnVAT.Close()
        cnACCPAC = Nothing
        cnVAT = Nothing
        Exit Sub
ErrHandler:
        WriteLog("<ProcessAP>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessAP> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
        WriteLog(ErrMsg)
    End Sub

    Public Sub ProcessAR()
        Dim cnACCPAC, cnVAT As ADODB.Connection
        Dim rs, rs1, rs3 As BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim strsplit() As String
        Dim losplit As Integer
        Dim loCls As clsVat
        Dim iffs As clsFuntion
        Dim Ddbo As String
        iffs = New clsFuntion
        ErrMsg = ""

        On Error GoTo ErrHandler
        cnACCPAC = New ADODB.Connection
        cnVAT = New ADODB.Connection
        cnACCPAC.ConnectionTimeout = 60
        cnACCPAC.Open(ConnACCPAC)
        cnACCPAC.CommandTimeout = 3600

        cnVAT.ConnectionTimeout = 60
        cnVAT.Open(ConnVAT)
        cnVAT.CommandTimeout = 3600

        'gather data
        rs3 = New BaseClass.BaseODBCIO
        sqlstr = "SELECT HOMECUR FROM CSCOM "
        rs3.Open(sqlstr, cnACCPAC)


        Ddbo = ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.")

        rs = New BaseClass.BaseODBCIO
        If Use_POST Then
            sqlstr = "SELECT " & Ddbo & "AROBLTEMP.DATEINVC, " & Ddbo & "AROBLTEMP.IDINVC, " & Ddbo & "AROBLTEMP.TRXTYPEID, "
            If Use_LEGALNAME_inAR = True Then
                sqlstr &= "CASE WHEN ARCSP.NAMELOCN IS NULL THEN ARCUS.NAMECUST ELSE ARCSP.NAMELOCN END AS THNAME, ARCUS.TEXTSTRE1, " & Ddbo & "AROBLTEMP.AMTINVCHC, " & Ddbo & "AROBLTEMP.AMTTAXHC, " & Ddbo & "AROBLTEMP.CNTBTCH, " & vbCrLf
            Else
                sqlstr &= "ARCUS.NAMECUST, ARCUS.TEXTSTRE1, " & Ddbo & "AROBLTEMP.AMTINVCHC, " & Ddbo & "AROBLTEMP.AMTTAXHC, " & Ddbo & "AROBLTEMP.CNTBTCH, " & vbCrLf
            End If

            '----------------------------- bond Edit --------------------------
            ' sqlstr &= "         " & Ddbo & "AROBLTEMP.CNTITEM, " & Ddbo & "AROBLTEMP.CODETAX1, " & Ddbo & "AROBLTEMP.IDRMIT, SUM(ARIBD.BASETAX1) AS Expr1, SUM(ARIBD.TOTTAX) AS Expr2, " & Ddbo & "AROBLTEMP.DATEBUS, " & Ddbo & "AROBLTEMP.IDCUST, " & vbCrLf 
            sqlstr &= "         " & Ddbo & "AROBLTEMP.CNTITEM, " & Ddbo & "AROBLTEMP.CODETAX1, " & Ddbo & "AROBLTEMP.IDRMIT, SUM(ARIBD.AMTEXTNHC) AS Expr1, SUM(ARIBD.TOTTAX) AS Expr2, " & Ddbo & "AROBLTEMP.DATEBUS, " & Ddbo & "AROBLTEMP.IDCUST, " & vbCrLf
            sqlstr &= "         '' AS TEXTLINE,0 AS CNTLINE,'HEADER' AS VATAS " & vbCrLf
            sqlstr &= "FROM     " & Ddbo & "AROBLTEMP " & vbCrLf
            sqlstr &= "         INNER JOIN ARIBD ON " & Ddbo & "AROBLTEMP.CNTBTCH = ARIBD.CNTBTCH AND " & Ddbo & "AROBLTEMP.CNTITEM = ARIBD.CNTITEM " & vbCrLf
            sqlstr &= "         LEFT OUTER JOIN ARCUS ON " & Ddbo & "AROBLTEMP.IDCUST = ARCUS.IDCUST " & vbCrLf
            sqlstr &= "         LEFT OUTER JOIN TXAUTH ON " & Ddbo & "AROBLTEMP.CODETAXGRP = TXAUTH.AUTHORITY " & vbCrLf
            sqlstr &= "         LEFT OUTER JOIN ARCSP ON ARCSP.IDCUST = " & Ddbo & "AROBLTEMP.IDCUST AND ARCSP.IDCUSTSHPT = 'TXTHAI' " & vbCrLf
            sqlstr &= "WHERE    (" & Ddbo & "AROBLTEMP.TRXTYPETXT IN (1,2,3)) AND (" & Ddbo & "AROBLTEMP.DATEBUS >= " & DATEFROM & " AND " & Ddbo & "AROBLTEMP.DATEBUS <= " & DATETO & ") AND (ARIBD.TAXSTTS1 = 1) " & vbCrLf
            'sqlstr &= "and AMTTAXHC<>0  " comment เพื่อให้ Vat 0 ออกทุกตัว
            'TXAUTH.ACCTRECOV แก้ไขเนื่องจาก account ขาย อยู่ที่ TXAUTH.LIABILITY  ,ซื้อ อยู่ที่  TXAUTH.ACCTRECOV
            sqlstr &= "         AND  TXAUTH.LIABILITY IN (SELECT ACCTVAT " & IIf(ComDatabase = "PVSW", "", "COLLATE Thai_CI_AS") & " FROM " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "FMSVLACC) " & vbCrLf
            If Use_LEGALNAME_inAR = True Then
                sqlstr &= "GROUP BY " & Ddbo & "AROBLTEMP.DATEINVC, " & Ddbo & "AROBLTEMP.IDINVC, " & Ddbo & "AROBLTEMP.TRXTYPEID, CASE WHEN ARCSP.NAMELOCN IS NULL THEN ARCUS.NAMECUST ELSE ARCSP.NAMELOCN END , ARCUS.TEXTSTRE1, " & Ddbo & "AROBLTEMP.AMTINVCHC, " & Ddbo & "AROBLTEMP.AMTTAXHC, " & Ddbo & "AROBLTEMP.CNTBTCH,  " & vbCrLf
                sqlstr &= "         " & Ddbo & "AROBLTEMP.CNTITEM, " & Ddbo & "AROBLTEMP.CODETAX1, " & Ddbo & "AROBLTEMP.IDRMIT, " & Ddbo & "AROBLTEMP.DATEBUS, " & Ddbo & "AROBLTEMP.IDCUST " & vbCrLf
            Else
                sqlstr &= "GROUP BY " & Ddbo & "AROBLTEMP.DATEINVC, " & Ddbo & "AROBLTEMP.IDINVC, " & Ddbo & "AROBLTEMP.TRXTYPEID, ARCUS.NAMECUST, ARCUS.TEXTSTRE1, " & Ddbo & "AROBLTEMP.AMTINVCHC, " & Ddbo & "AROBLTEMP.AMTTAXHC, " & Ddbo & "AROBLTEMP.CNTBTCH,  " & vbCrLf
                sqlstr &= "         " & Ddbo & "AROBLTEMP.CNTITEM, " & Ddbo & "AROBLTEMP.CODETAX1, " & Ddbo & "AROBLTEMP.IDRMIT, " & Ddbo & "AROBLTEMP.DATEBUS, " & Ddbo & "AROBLTEMP.IDCUST " & vbCrLf
            End If

            sqlstr &= "UNION ALL " & vbCrLf

            sqlstr &= "SELECT " & Ddbo & "AROBLTEMP.DATEINVC, " & Ddbo & "AROBLTEMP.IDINVC, " & Ddbo & "AROBLTEMP.TRXTYPEID, "
            If Use_LEGALNAME_inAR = True Then
                sqlstr &= "CASE WHEN ARCSP.NAMELOCN IS NULL THEN ARCUS.NAMECUST ELSE ARCSP.NAMELOCN END , ARCUS.TEXTSTRE1, " & Ddbo & "AROBLTEMP.AMTINVCHC, " & Ddbo & "AROBLTEMP.AMTTAXHC, " & Ddbo & "AROBLTEMP.CNTBTCH,  " & vbCrLf
            Else
                sqlstr &= "ARCUS.NAMECUST, ARCUS.TEXTSTRE1, " & Ddbo & "AROBLTEMP.AMTINVCHC, " & Ddbo & "AROBLTEMP.AMTTAXHC, " & Ddbo & "AROBLTEMP.CNTBTCH,  " & vbCrLf
            End If

            '----------------------------- bond Edit --------------------------
            ' sqlstr &= "         " & Ddbo & "AROBLTEMP.CNTITEM, " & Ddbo & "AROBLTEMP.CODETAX1, " & Ddbo & "AROBLTEMP.IDRMIT, SUM(ARIBD.BASETAX1) AS Expr1, SUM(ARIBD.TOTTAX) AS Expr2, " & Ddbo & "AROBLTEMP.DATEBUS, " & Ddbo & "AROBLTEMP.IDCUST, " & vbCrLf 
            sqlstr &= "         " & Ddbo & "AROBLTEMP.CNTITEM, " & Ddbo & "AROBLTEMP.CODETAX1, " & Ddbo & "AROBLTEMP.IDRMIT, SUM(ARIBD.AMTEXTNHC) AS Expr1, SUM(ARIBD.TOTTAX) AS Expr2, " & Ddbo & "AROBLTEMP.DATEBUS, " & Ddbo & "AROBLTEMP.IDCUST, " & vbCrLf
            sqlstr &= "         ARIBT.TEXTLINE,ARPJD.CNTLINE,'DETAIL' AS VATAS " & vbCrLf
            sqlstr &= "FROM " & Ddbo & "AROBLTEMP " & vbCrLf
            sqlstr &= "         INNER JOIN ARIBD ON " & Ddbo & "AROBLTEMP.CNTBTCH = ARIBD.CNTBTCH AND " & Ddbo & "AROBLTEMP.CNTITEM = ARIBD.CNTITEM " & vbCrLf
            sqlstr &= "         LEFT OUTER JOIN " & Ddbo & "ARPJHTEMP ON " & Ddbo & "AROBLTEMP.IDINVC = " & Ddbo & "ARPJHTEMP.IDINVC AND " & Ddbo & "AROBLTEMP.IDCUST  = " & Ddbo & "ARPJHTEMP.IDCUST " & vbCrLf
            sqlstr &= "         LEFT OUTER JOIN ARPJD ON " & Ddbo & "ARPJHTEMP.POSTSEQNCE = ARPJD.POSTSEQNCE AND " & Ddbo & "ARPJHTEMP.CNTBTCH = ARPJD.CNTBTCH AND " & Ddbo & "ARPJHTEMP.CNTITEM = ARPJD.CNTITEM " & vbCrLf
            sqlstr &= "         LEFT OUTER JOIN ARCUS ON " & Ddbo & "AROBLTEMP.IDCUST = ARCUS.IDCUST " & vbCrLf
            sqlstr &= "         LEFT OUTER JOIN ARIBT ON " & Ddbo & "AROBLTEMP.CNTBTCH = ARIBT.CNTBTCH AND " & Ddbo & "AROBLTEMP.CNTITEM = ARIBT.CNTITEM AND ARIBT.CNTLINE = ARIBD.CNTLINE " & vbCrLf
            sqlstr &= "         LEFT OUTER JOIN ARCSP ON ARCSP.IDCUST = " & Ddbo & "AROBLTEMP.IDCUST AND ARCSP.IDCUSTSHPT = 'TXTHAI'" & vbCrLf
            sqlstr &= "WHERE     (" & Ddbo & "AROBLTEMP.TRXTYPETXT IN (1,2,3)) AND (" & Ddbo & "AROBLTEMP.DATEBUS >= " & DATEFROM & " AND " & Ddbo & "AROBLTEMP.DATEBUS <= " & DATETO & ") AND (ARIBD.TAXSTTS1 = 1) " & vbCrLf
            sqlstr &= "         AND " & Ddbo & "AROBLTEMP.AMTTAXHC = 0 " & vbCrLf
            sqlstr &= "         AND ARPJD.IDACCT IN (SELECT ACCTVAT " & IIf(ComDatabase = "PVSW", "", "COLLATE Thai_CI_AS") & " FROM " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "FMSVLACC)   " & vbCrLf
            If Use_LEGALNAME_inAR = True Then
                sqlstr &= "GROUP BY " & Ddbo & "AROBLTEMP.DATEINVC, " & Ddbo & "AROBLTEMP.IDINVC, " & Ddbo & "AROBLTEMP.TRXTYPEID, CASE WHEN ARCSP.NAMELOCN IS NULL THEN ARCUS.NAMECUST ELSE ARCSP.NAMELOCN END, ARCUS.TEXTSTRE1, " & Ddbo & "AROBLTEMP.AMTINVCHC, " & Ddbo & "AROBLTEMP.AMTTAXHC, " & Ddbo & "AROBLTEMP.CNTBTCH,  " & vbCrLf
                sqlstr &= "         " & Ddbo & "AROBLTEMP.CNTITEM, " & Ddbo & "AROBLTEMP.CODETAX1, " & Ddbo & "AROBLTEMP.IDRMIT, " & Ddbo & "AROBLTEMP.DATEBUS, " & Ddbo & "AROBLTEMP.IDCUST,ARPJD.CNTLINE ,ARIBT.TEXTLINE " & vbCrLf
            Else
                sqlstr &= "GROUP BY " & Ddbo & "AROBLTEMP.DATEINVC, " & Ddbo & "AROBLTEMP.IDINVC, " & Ddbo & "AROBLTEMP.TRXTYPEID, ARCUS.NAMECUST, ARCUS.TEXTSTRE1, " & Ddbo & "AROBLTEMP.AMTINVCHC, " & Ddbo & "AROBLTEMP.AMTTAXHC, " & Ddbo & "AROBLTEMP.CNTBTCH,  " & vbCrLf
                sqlstr &= "         " & Ddbo & "AROBLTEMP.CNTITEM, " & Ddbo & "AROBLTEMP.CODETAX1, " & Ddbo & "AROBLTEMP.IDRMIT, " & Ddbo & "AROBLTEMP.DATEBUS, " & Ddbo & "AROBLTEMP.IDCUST,ARPJD.CNTLINE ,ARIBT.TEXTLINE " & vbCrLf
            End If
        End If

        rs.Open(sqlstr, cnACCPAC)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessAR> Select AR Complete")
        End If

        Do While rs.options.Rs.EOF = False

            loCls = New clsVat
            With loCls
                ErrMsg = "Ledger = AR IDINVC=" & Trim(rs.options.Rs.Collect(1))
                ErrMsg = ErrMsg & " Batch No.=" & rs.options.Rs.Collect(7)
                ErrMsg = ErrMsg & " Entry No.=" & rs.options.Rs.Collect(8)

                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect("DATEINVC")), "", (rs.options.Rs.Collect("DATEINVC"))) 'AROBL.DATEINVC
                .TXDATE = IIf(IsDBNull(rs.options.Rs.Collect("DATEBUS")), "", (rs.options.Rs.Collect("DATEBUS"))) 'AROBL.DATEBUS
                .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect(1)), "", (rs.options.Rs.Collect(1))) 'AROBL.IDINVC
                .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect(3)), "", (rs.options.Rs.Collect(3))) 'ARCUS.NAMECUST


                rs1 = New BaseClass.BaseODBCIO

                If ComDatabase = "PVSW" Then
                    sqlstr = " SELECT ARIBHO.VALUE FROM ARIBHO "
                Else
                    sqlstr = " SELECT ARIBHO.[VALUE] FROM ARIBHO "
                End If
                sqlstr = sqlstr & " WHERE ARIBHO.OPTFIELD = 'EXPORTAMT'"
                sqlstr = sqlstr & " AND ARIBHO.CNTBTCH = '" & rs.options.Rs.Collect(7) & "' AND ARIBHO.CNTITEM = '" & rs.options.Rs.Collect(8) & "'"
                rs1.Open(sqlstr, cnACCPAC)

                If rs1.options.Rs.EOF Then
                    .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(11)), 0, rs.options.Rs.Collect(11)) 'ARIBD.BASETAX1
                Else
                    .INVAMT = IIf(IsDBNull(rs1.options.Rs.Collect(0)), 0, rs1.options.Rs.Collect(0)) 'ARIBD.BASETAX1
                End If

                .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(12)), 0, rs.options.Rs.Collect(12)) 'ARIBD.TOTTAX

                If rs.options.Rs.Collect(2) = 31 Or rs.options.Rs.Collect(2) = 32 Or rs.options.Rs.Collect(2) = 34 Or rs.options.Rs.Collect(2) = 35 Then
                    .INVTAX = .INVTAX * (-1)
                    .INVAMT = .INVAMT * (-1)
                End If


                If rs.options.Rs.Collect("VATAS").ToString.Trim = "HEADER" Then

                    Code = IIf(IsDBNull(rs.options.Rs.Collect("IDCUST")), "", (rs.options.Rs.Collect("IDCUST"))) 'ARCUS.IDCUST

                    Dim RS_WHT_HEADER As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from ARIBHO  where OPTFIELD in ('TAXBASE','TAXINVNO','TAXDATE','TAXNAME') and CNTBTCH ='" & rs.options.Rs.Collect("CNTBTCH") & "' and CNTITEM='" & rs.options.Rs.Collect("CNTITEM") & "'", cnACCPAC)

                    If OPT_AP = True Then
                        Dim tINVAMT As Double = 0
                        Dim tTAXINVNO As String = ""
                        Dim tTAXDATE As String = ""
                        Dim tTAXNAME As String = ""
                        For iwht As Integer = 0 To RS_WHT_HEADER.options.QueryDT.Rows.Count - 1
                            If RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXBASE" Then
                                Dim Trydec As String = Trim(RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("VALUE"))
                                If Decimal.TryParse(Trydec, 0) = True Then
                                    If CDec(Trydec) <> 0 Then
                                        tINVAMT = CDec(Trydec)
                                    Else
                                        tINVAMT = 0
                                    End If
                                Else
                                    tINVAMT = 0
                                End If
                            ElseIf RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXINVNO" Then
                                If Trim(RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    tTAXINVNO = Trim(RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If
                            ElseIf RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXDATE" Then
                                Dim myDate As String = Trim(RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("VALUE"))
                                If Decimal.TryParse(myDate, 0) = True Then
                                    If CDec(myDate) <> 0 Then
                                        tTAXDATE = CDec(myDate)
                                    End If
                                End If
                            ElseIf RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXNAME" Then
                                If Trim(RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    tTAXNAME = Trim(RS_WHT_HEADER.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If
                            End If
                        Next
                        If tINVAMT <> 0 Then
                            .INVAMT = tINVAMT
                        End If
                        If tTAXINVNO <> "" Then
                            .IDINVC = tTAXINVNO
                        End If
                        If tTAXDATE <> "" Then
                            .INVDATE = tTAXDATE
                        End If
                        If tTAXNAME <> "" Then
                            .INVNAME = tTAXNAME
                        End If
                    End If

                    If VNO_AR = True Then

                        Dim VNO_Header As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from ARIBHO  where OPTFIELD in ('ARINVNO') and CNTBTCH ='" & rs.options.Rs.Collect("CNTBTCH") & "' and CNTITEM='" & rs.options.Rs.Collect("CNTITEM") & "'", cnACCPAC)
                        For iwht As Integer = 0 To VNO_Header.options.QueryDT.Rows.Count - 1
                            If VNO_Header.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "ARINVNO" Then
                                If Trim(VNO_Header.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    .DOCNO = Trim(VNO_Header.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If

                            End If
                        Next

                    End If


                ElseIf rs.options.Rs.Collect("VATAS").ToString.Trim = "DETAIL" Then
                    Dim TEXTLINE() As String = rs.options.Rs.Collect("TEXTLINE").ToString.Split(",")
                    Dim RS_WHT_DETAIL As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from ARIBDO  where OPTFIELD in ('TAXBASE','TAXINVNO','TAXDATE','TAXNAME') and CNTBTCH ='" & rs.options.Rs.Collect("CNTBTCH") & "' and CNTITEM='" & rs.options.Rs.Collect("CNTITEM") & "' and CNTLINE ='" & IIf(IsDBNull(rs.options.Rs.Collect("CNTLINE")), "0", rs.options.Rs.Collect("CNTLINE")) & "' ", cnACCPAC)

                    If OPT_AP = False Then 'OPT_AP คือการเช็คว่าให้ดึง VAT ที่ OPT หรือไม่ ส่วนนี้ดึง VAT จาก Comment
                        For i As Integer = 0 To TEXTLINE.Length - 1
                            If IsDBNull(TEXTLINE(i)) = True Then
                                TEXTLINE(i) = ""
                            End If
                        Next
                        If TEXTLINE.Length >= 1 Then
                            If TEXTLINE(0).Trim <> "" And Decimal.TryParse(TEXTLINE(0), 0) Then
                                .INVAMT = CDec(TEXTLINE(0))
                            End If
                        End If

                        If TEXTLINE.Length >= 2 Then
                            If TEXTLINE(1).Trim <> "" Then
                                .IDINVC = TEXTLINE(1)
                            End If
                        End If

                        If TEXTLINE.Length >= 3 Then
                            If TEXTLINE(2).Trim <> "" Then
                                .INVDATE = TryDate2(ZeroDate(TEXTLINE(2), "dd/MM/yyyy", "yyyyMMdd"), "dd/MM/yyyy", "yyyyMMdd")
                                If .INVDATE = 0 Then
                                    .INVDATE = TryDate2(TEXTLINE(2), "dd/MM/yyyy", "yyyyMMdd")
                                End If
                            End If
                        End If

                        If TEXTLINE.Length >= 4 Then
                            If TEXTLINE(3).Trim = "" Then
                                '.INVNAME = rs.options.Rs.Collect("VENDNAME")
                                .INVNAME = rs.options.Rs.Collect("IDCUST")
                            Else
                                .INVNAME = TEXTLINE(3)
                                For CName As Integer = 4 To TEXTLINE.Length - 1
                                    .INVNAME = .INVNAME & "," & TEXTLINE(CName)
                                Next
                            End If
                        End If
                    Else 'ดึง VAT จาก OPT
                        Dim tINVAMT As Double = 0
                        Dim tTAXINVNO As String = ""
                        Dim tTAXDATE As String = ""
                        Dim tTAXNAME As String = ""
                        For iwht As Integer = 0 To RS_WHT_DETAIL.options.QueryDT.Rows.Count - 1
                            If RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXBASE" Then
                                Dim Trydec As String = Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE"))
                                If Decimal.TryParse(Trydec, 0) = True Then
                                    If CDec(Trydec) <> 0 Then
                                        tINVAMT = CDec(Trydec)
                                    Else
                                        tINVAMT = 0
                                    End If
                                Else
                                    tINVAMT = 0
                                End If
                            ElseIf RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXINVNO" Then
                                If Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    tTAXINVNO = Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If
                            ElseIf RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXDATE" Then
                                Dim myDate As String = Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE"))
                                If Decimal.TryParse(myDate, 0) = True Then
                                    If CDec(myDate) <> 0 Then
                                        tTAXDATE = CDec(myDate)
                                    End If
                                End If
                            ElseIf RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXNAME" Then
                                If Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    Dim chkIDCUST_Opt As New BaseClass.BaseODBCIO("select ARCUS.IDCUST, ARCSP.NAMELOCN AS ARSHIPTO from ARCUS INNER JOIN  ARCSP ON ARCUS.IDCUST = ARCSP.IDCUST WHERE IDCUSTSHPT = 'TXTHAI' and ARCUS.IDCUST = '" & Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE")) & "'", cnACCPAC)
                                    If chkIDCUST_Opt.options.QueryDT.Rows.Count = 0 Then
                                        tTAXNAME = Trim(RS_WHT_DETAIL.options.QueryDT.Rows(iwht).Item("VALUE"))
                                    Else
                                        tTAXNAME = chkIDCUST_Opt.options.QueryDT.Rows(0).Item("ARSHIPTO").ToString
                                    End If
                                End If
                            End If
                        Next
                        If tINVAMT <> 0 Then
                            .INVAMT = tINVAMT
                        Else
                            .INVAMT = 0
                        End If
                        If tTAXINVNO <> "" Then
                            .IDINVC = tTAXINVNO
                        End If
                        If tTAXDATE <> "" Then
                            .INVDATE = tTAXDATE
                        End If
                        If tTAXNAME <> "" Then
                            .INVNAME = tTAXNAME
                        End If
                    End If

                    If VNO_AR = True Then

                        Dim VNO_Detail As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from ARIBDO  where OPTFIELD in ('ARINVNO') and CNTBTCH ='" & rs.options.Rs.Collect("CNTBTCH") & "' and CNTITEM='" & rs.options.Rs.Collect("CNTITEM") & "'", cnACCPAC)
                        For iwht As Integer = 0 To VNO_Detail.options.QueryDT.Rows.Count - 1
                            If VNO_Detail.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "ARINVNO" Then
                                If Trim(VNO_Detail.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                    .DOCNO = Trim(VNO_Detail.options.QueryDT.Rows(iwht).Item("VALUE"))
                                End If

                            End If
                        Next

                    End If

                End If '"DETAIL" "HEADER"

                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7)) 'AROBL.CNTBTCH
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8)) 'AROBL.CNTITEM
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(10)), 0, rs.options.Rs.Collect(10)) 'AROBL.IDRMIT
                .TRANSNBR = IIf(IsDBNull(rs.options.Rs.Collect("CNTLINE")), "", rs.options.Rs.Collect("CNTLINE")) 'CNTLINE

                Dim C9 As String = rs.options.Rs.Collect(9)

                If FindMiscCodeInARCUS(Trim(.INVNAME)) = "0" Or FindMiscCodeInARCUS(Trim(.INVNAME)) = "" Then
                    If Trim(.INVNAME) = "" Then
                        .INVNAME = FindMiscCodeInARCUS(Trim(rs.options.Rs.Collect("IDCUST")))
                    Else
                        .INVNAME = Trim(.INVNAME)
                    End If
                    .TAXID = FindTaxIDBRANCH(Trim(rs.options.Rs.Collect("IDCUST")), "AR", "TAXID")
                    .Branch = FindTaxIDBRANCH(Trim(rs.options.Rs.Collect("IDCUST")), "AR", "BRANCH")
                Else
                    .TAXID = FindTaxIDBRANCH(Trim(.INVNAME), "AR", "TAXID")
                    .Branch = FindTaxIDBRANCH(Trim(.INVNAME), "AR", "BRANCH")
                    .INVNAME = FindMiscCodeInARCUS(Trim(.INVNAME))
                End If

                .VATCOMMENT = .INVAMT & "," & .IDINVC & "," & .INVDATE & "," & .INVNAME
                .MARK = "Invoice"

                If .INVAMT = 0 Then
                    .RATE_Renamed = 0
                Else
                    .RATE_Renamed = FormatNumber(((.INVTAX / .INVAMT) * 100), 1)
                End If


                sqlstr = InsVat(0, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, "", .INVNAME, .INVAMT, .INVTAX, .LOCID, 1, .RATE_Renamed, .TTYPE, .ACCTVAT, "AR", .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, .Runno, C9, "", "", "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, 0)

            End With

            cnVAT.Execute(sqlstr)
            ErrMsg = ""
            Code = ""
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing

        'transfer data
        rs = New BaseClass.BaseODBCIO
        sqlstr = "SELECT FMSVATTEMP.INVDATE,FMSVATTEMP.TXDATE,FMSVATTEMP.IDINVC,FMSVATTEMP.DOCNO," '0-3
        sqlstr = sqlstr & " FMSVATTEMP.INVNAME,FMSVATTEMP.INVAMT,FMSVATTEMP.INVTAX," '4-6
        sqlstr = sqlstr & " FMSVLACC.LOCID,FMSVATTEMP.VTYPE,FMSVATTEMP.RATE,FMSTAX.TTYPE," '7-10
        sqlstr = sqlstr & " FMSTAX.ACCTVAT,FMSVATTEMP.SOURCE,FMSVATTEMP.BATCH,FMSVATTEMP.ENTRY," '11-14
        sqlstr = sqlstr & " FMSVATTEMP.MARK,FMSVATTEMP.VATCOMMENT,FMSTAX.ITEMRATE1,FMSVATTEMP.IDDIST,FMSVATTEMP.CBREF, " '15-19
        sqlstr = sqlstr & " FMSVATTEMP.TAXID,FMSVATTEMP.BRANCH ,FMSVATTEMP.TRANSNBR,FMSVATTEMP.Code,FMSVATTEMP.CODETAX "
        If ComDatabase = "ORCL" Then
            sqlstr = sqlstr & " FROM FMSVATTEMP"
            sqlstr = sqlstr & "     INNER JOIN FMSTAX ON (FMSVATTEMP.CODETAX = FMSTAX.AUTHORITY) "
            sqlstr = sqlstr & "     INNER JOIN FMSVLACC ON FMSTAX.ACCTVAT = FMSVLACC.ACCTVAT"
            sqlstr = sqlstr & " WHERE FMSTAX.TTYPE = 2 AND FMSTAX.BUYERCLASS = 1"
        Else
            sqlstr = sqlstr & " FROM FMSVATTEMP"
            sqlstr = sqlstr & "     INNER JOIN FMSTAX ON (FMSVATTEMP.CODETAX = FMSTAX.AUTHORITY) "
            sqlstr = sqlstr & "     INNER JOIN FMSVLACC ON FMSTAX.ACCTVAT = FMSVLACC.ACCTVAT"
            sqlstr = sqlstr & " WHERE FMSTAX.TTYPE = 2 AND FMSTAX.BUYERCLASS = 1"
        End If

        rs.Open(sqlstr, cnVAT)
        If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
            WriteLog("<ProcessAR> Select FMSVATTEMP Complete")
        End If


        Do While rs.options.Rs.EOF = False
            loCls = New clsVat
            With loCls
                .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect(0)), 0, rs.options.Rs.Collect(0))
                .TXDATE = IIf(IsDBNull(rs.options.Rs.Collect(1)), 0, rs.options.Rs.Collect(1))
                .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", (rs.options.Rs.Collect(2)))
                .DOCNO = IIf(IsDBNull(rs.options.Rs.Collect(3)), "", (rs.options.Rs.Collect(3)))
                .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect(4)), "", (rs.options.Rs.Collect(4)))
                .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(5)), 0, rs.options.Rs.Collect(5))
                .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(6)), 0, rs.options.Rs.Collect(6))
                .LOCID = IIf(IsDBNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7))
                .VTYPE = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8))
                .TTYPE = IIf(IsDBNull(rs.options.Rs.Collect(10)), 0, rs.options.Rs.Collect(10))
                .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect(11)), "", (rs.options.Rs.Collect(11)))
                .Source = IIf(IsDBNull(rs.options.Rs.Collect(12)), "", (rs.options.Rs.Collect(12)))
                .Batch = IIf(IsDBNull(rs.options.Rs.Collect(13)), "", (rs.options.Rs.Collect(13)))
                .Entry = IIf(IsDBNull(rs.options.Rs.Collect(14)), "", (rs.options.Rs.Collect(14)))
                .MARK = IIf(IsDBNull(rs.options.Rs.Collect(15)), "", (rs.options.Rs.Collect(15)))
                .VATCOMMENT = IIf(IsDBNull(rs.options.Rs.Collect(16)), "", (rs.options.Rs.Collect(16)))
                .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(19)), "", (rs.options.Rs.Collect(19)))
                .TRANSNBR = IIf(IsDBNull(rs.options.Rs.Collect("TRANSNBR")), "", rs.options.Rs.Collect("TRANSNBR"))
                .TAXID = IIf(IsDBNull(rs.options.Rs.Collect("TAXID")), "", (rs.options.Rs.Collect("TAXID")))
                .Branch = IIf(IsDBNull(rs.options.Rs.Collect("BRANCH")), "", (rs.options.Rs.Collect("BRANCH")))
                Code = IIf(IsDBNull(rs.options.Rs.Collect("Code")), "", (rs.options.Rs.Collect("Code")))
                .RATE_Renamed = IIf(IsDBNull(rs.options.Rs.Collect("RATE")), 0, rs.options.Rs.Collect("RATE"))
                .CODETAX = IIf(IsDBNull(rs.options.Rs.Collect("CODETAX")), "", (rs.options.Rs.Collect("CODETAX")))

                sqlstr = InsVat(1, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, .NEWDOCNO, .INVNAME, .INVAMT, .INVTAX, .LOCID, .VTYPE, .RATE_Renamed, .TTYPE, .ACCTVAT, .Source, .Batch, .Entry, .MARK, .VATCOMMENT, .CBRef, .TRANSNBR, .Runno, .CODETAX, "", "", "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, 0)


            End With
            cnVAT.Execute(sqlstr)
            rs.options.Rs.MoveNext()

            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        rs = Nothing
        cnACCPAC.Close()
        cnVAT.Close()
        cnACCPAC = Nothing
        cnVAT = Nothing
        Exit Sub

ErrHandler:
        WriteLog("<ProcessAR>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessAR> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
        WriteLog(ErrMsg)
    End Sub


    Public Sub PrepareCB()
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String

        On Error GoTo ErrHandler

        cn = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO

        If TableHere("FMSCB", ConnVAT) = True Then
            sqlstr = "DELETE FROM FMSCB"
            cn.Open(ConnVAT)
            cn.Execute(sqlstr)
            cn.Close()
        Else
            cn = New ADODB.Connection
            cn.Open(ConnVAT)
            sqlstr = "CREATE TABLE FMSCB"
            sqlstr = sqlstr & " (CBDATE     DECIMAL(9,0),   "
            sqlstr = sqlstr & " TXDATE      DECIMAL(9,0),   "
            sqlstr = sqlstr & " BATCH       CHAR(6) ,"
            sqlstr = sqlstr & " ENTRY       CHAR(6) ,"
            sqlstr = sqlstr & " DOCNO       CHAR(255) ,"
            sqlstr = sqlstr & " REFERENCE   CHAR(12) ,"
            sqlstr = sqlstr & " NAME        CHAR(100) ,"
            sqlstr = sqlstr & " INVAMT      DECIMAL(19,3),"
            sqlstr = sqlstr & " INVTAX      DECIMAL(19,3),"
            sqlstr = sqlstr & " ACCTVAT     CHAR(45) ,"
            sqlstr = sqlstr & " VATCOMMENT  CHAR(1000),"
            sqlstr = sqlstr & " TRANSNBR    VARCHAR(15),"
            sqlstr = sqlstr & " RUNNO       VARCHAR(60)," 'แก้จาก VARCHAR(20) เป็น VARCHAR(60) Edit 07/02/2014 By Pat
            sqlstr = sqlstr & " TAXID       VARCHAR(13),"
            sqlstr = sqlstr & " BRANCH      VARCHAR(500),"
            sqlstr = sqlstr & " Code        VARCHAR(12),"
            sqlstr = sqlstr & " TOTTAX      DECIMAL(19,3))"


            cn.Execute(sqlstr)
            cn.Close()
        End If

        WriteLog("<PrepareVatCB> table ready")
        rs = Nothing
        cn = Nothing
        Exit Sub

ErrHandler:
        cn.RollbackTrans()
        ErrMsg = Err.Description
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
        WriteLog("<PrepareVatCB>" & Err.Number & "  " & Err.Description)
    End Sub

    Public Sub ProcessCB()
        Dim cnACCPAC, cnVAT As ADODB.Connection
        Dim rs, rs1 As BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim strsplit(), strsplit2() As String
        Dim losplit, losplit2 As Integer
        Dim splittxt, splittxt2 As String
        Dim loCls As clsVat
        Dim loCBRef, locbCBDATE, locbTXDATE, locbDOCNO, locbNAME, TmpComment, locbTrans As String
        Dim MISC As String = ""
        Dim locbTAXID As String = ""
        Dim locbBRANCH As String = ""
        Dim locbINVAMT, locbINVTAX, locbTOTTAX As Double
        Dim locbVATCOMMENT, locbACCTVAT, loRunno As String
        Dim locbDescription, locbMiscCode As String
        Dim locbBatch, locbEntry As String
        Dim locbType As Short
        Dim tmpName, TmpCB As String
        Dim tmpDate As Double
        Dim Temp_DateCB As String
        Dim Temp_YearCb As Double
        Dim iffs As clsFuntion
        iffs = New clsFuntion
        ErrMsg = ""

        ' On Error GoTo ErrHandler
        Try

            cnACCPAC = New ADODB.Connection

            cnVAT = New ADODB.Connection
            cnACCPAC.ConnectionTimeout = 60
            cnACCPAC.Open(ConnACCPAC)
            cnACCPAC.CommandTimeout = 3600

            cnVAT.ConnectionTimeout = 60
            cnVAT.Open(ConnVAT)
            cnVAT.CommandTimeout = 3600



            Try
                '***************************************
                'Chang year to ค.ศ
                Temp_YearCb = CShort(Right(CStr(Today.ToString("dd/MM/yyyy")), 4))
                'DATETO = 20050331
                'DATEFROM = 20050301
            Catch ex As Exception
                WriteLog("<Chang year> " & ex.Message)
           
            End Try


            If Temp_YearCb > 2500 Then
                Temp_DateCB = Mid(DATETO, 5, 4)
                DATETO = Trim(Str(Val(Left(DATETO, 4)) + 543) & Temp_DateCB)
                Temp_DateCB = Mid(DATEFROM, 5, 4)
                DATEFROM = Trim(Str(Val(Left(DATEFROM, 4)) + 543) & Temp_DateCB)
            End If


            rs = New BaseClass.BaseODBCIO
            sqlstr = ""
            sqlstr = " SELECT   PGMVER, SELECTOR  FROM CSAPP"
            sqlstr &= "  WHERE  (SELECTOR = 'CB')"
            rs.Open(sqlstr, cnACCPAC)
            If rs.options.Rs.EOF = True Then Exit Sub
            TmpCB = Mid(Trim(rs.options.Rs.Collect(0)), 1, 2)

            rs = New BaseClass.BaseODBCIO
            sqlstr = ""

            'ใช้ Get Format date จากเครื่องที่รันโปรแกรม เช่น  dd/MM/yyyy , M/dd/yyyy
            Dim Dateformat_PC As String = Globalization.DateTimeFormatInfo.CurrentInfo.ShortDatePattern


            'gather data
            rs = New BaseClass.BaseODBCIO
            If Use_POST Then
                If ComDatabase = "PVSW" Or ComDatabase = "ORCL" Then
                    sqlstr = "SELECT CBTRHD.""DATE"" , CBTRHD.INTERTYPE," '0-1
                ElseIf ComDatabase = "MSSQL" Then
                    sqlstr = "SELECT CBTRHD.[DATE] , CBTRHD.INTERTYPE," '0-1
                End If

                sqlstr &= "     CBTRHD.REFERENCE,CBTRHD.TEXTDESC," '2-3
                sqlstr &= "     CBTRDT.AMOUNT*CBTRHD.BK2GLRATE AS AMOUNT ,CBTRDT.QUANTITY," '4-5
                sqlstr &= "     CBTRDT.COMMENTS ,CBTRHD.ACCTREC1 AS ACCTID," '6-7
                sqlstr &= "     CBTRHD.BATCHID ,(case when ((CBTRDT.RCPTENTRY = null) or (CBTRDT.RCPTENTRY = '')) then CBTRHD.ENTRYNO else CBTRDT.RCPTENTRY end) AS ENTRYNO," '8-9,CBTRHD.ENTRYNO isnull(CBTRDT.RCPTENTRY ,CBTRHD.ENTRYNO) AS ENTRYNO
                sqlstr &= "     CBTRHD.MISCCODE," '10
                sqlstr &= "     CBTRDT.BANKCODE, CBTRDT.REFERENCE, CBTRDT.UNIQUIFIER, CBTRDT.DETAILNO,CBTRHDO.VALUE,"
                If Use_LEGALNAME_inCB = True Then
                    sqlstr &= "     isnull(CBMISC.COMMENTS ,VENDNAME )AS Vname,"
                Else
                    sqlstr &= "     isnull(CBMISC.MISCNAME ,VENDNAME )AS Vname," ' 11-14
                End If

                sqlstr &= "     CBTRHD.TAXGROUP,CBTRDT.AMTTAXREC1*CBTRHD.BK2GLRATE AS AMTTAXREC1,CBTRDT.AMTTAX1*CBTRHD.BK2GLRATE AS AMTTAX1,'VATAVG' AS VATTYPE" & Environment.NewLine
                sqlstr &= " FROM (CBTRHD"
                sqlstr &= "     LEFT JOIN CBTRDT ON (CBTRHD.BANKCODE = CBTRDT.BANKCODE) AND (CBTRHD.REFERENCE = CBTRDT.REFERENCE) AND (CBTRHD.UNIQUIFIER = CBTRDT.UNIQUIFIER)) "
                sqlstr &= "     LEFT JOIN CBTRHDO ON CBTRHDO.BANKCODE = CBTRDT.BANKCODE AND CBTRHDO.REFERENCE = CBTRDT.REFERENCE AND CBTRHDO.UNIQUIFIER = CBTRHD.UNIQUIFIER AND CBTRHDO.OPTFIELD = 'VOUCHER'"
                sqlstr &= "     LEFT JOIN CBMISC ON (CBTRHD.MISCCODE = CBMISC.MISCCODE ) "
                sqlstr &= "     LEFT JOIN APVEN ON (CBTRHD.MISCCODE = APVEN.VENDORID )" & Environment.NewLine

                If ComDatabase = "PVSW" Or ComDatabase = "ORCL" Then
                    sqlstr = sqlstr & "  WHERE (CBTRHD.""DATE"" >= '" & DATEFROM & "' AND CBTRHD.""DATE"" <= '" & DATETO & "')"
                ElseIf ComDatabase = "MSSQL" Then
                    sqlstr = sqlstr & "  WHERE (CBTRHD.[DATE] >= '" & DATEFROM & "' AND CBTRHD.[DATE] <= '" & DATETO & "')"
                End If
                sqlstr &= " AND (CBTRDT.AMTTAX1 <> 0 "
                sqlstr &= " AND CBTRHD.ACCTREC1 IN (select ACCTVAT " & IIf(ComDatabase = "PVSW", "", "COLLATE Thai_CI_AS") & " from " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "FMSVLACC)) "
                sqlstr &= " AND CBTRDT.ACCTID NOT IN (select ACCTVAT " & IIf(ComDatabase = "PVSW", "", "COLLATE Thai_CI_AS") & " from " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "FMSVLACC) " & Environment.NewLine
                If CBREVERSE = True Then sqlstr &= " AND (CONVERT(varchar(8), CBTRHD.DATE)+':'+ltrim(rtrim(CBTRHD.REFERENCE))) not in (SELECT CONVERT(varchar(8), DATE)+':'+ltrim(rtrim(REFERENCE)) AS Expr1 FROM CBTRHD where reconciled = 3 GROUP BY  DATE,REFERENCE having count(*)>1 )" & Environment.NewLine

                sqlstr &= " UNION ALL" & Environment.NewLine

                If ComDatabase = "PVSW" Or ComDatabase = "ORCL" Then
                    sqlstr &= "SELECT CBTRHD.""DATE"" , CBTRHD.INTERTYPE," '0-1
                ElseIf ComDatabase = "MSSQL" Then
                    sqlstr &= "SELECT CBTRHD.[DATE] , CBTRHD.INTERTYPE," '0-1
                End If

                sqlstr &= "     CBTRHD.REFERENCE,CBTRHD.TEXTDESC," '2-3
                sqlstr &= "     CBTRDT.AMOUNT*CBTRHD.BK2GLRATE AS AMOUNT,CBTRDT.QUANTITY," '4-5
                sqlstr &= "     CBTRDT.COMMENTS ,CBTRDT.ACCTID AS ACCTID," '6-7
                sqlstr &= "     CBTRHD.BATCHID ,(case when ((CBTRDT.RCPTENTRY = null) or (CBTRDT.RCPTENTRY = '')) then CBTRHD.ENTRYNO else CBTRDT.RCPTENTRY end) AS ENTRYNO," '8-9,CBTRHD.ENTRYNO isnull(CBTRDT.RCPTENTRY ,CBTRHD.ENTRYNO) AS ENTRYNO
                sqlstr &= "     CBTRHD.MISCCODE," '10
                sqlstr &= "     CBTRDT.BANKCODE, CBTRDT.REFERENCE, CBTRDT.UNIQUIFIER, CBTRDT.DETAILNO,CBTRHDO.VALUE,"
                If Use_LEGALNAME_inCB = True Then
                    sqlstr &= "isnull(CBMISC.COMMENTS ,VENDNAME )AS Vname," ' 11-14
                Else
                    sqlstr &= "isnull(CBMISC.MISCNAME ,VENDNAME )AS Vname," ' 11-14
                End If

                sqlstr &= "     CBTRHD.TAXGROUP,0 AS AMTTAXREC1,0 AS AMTTAX1,'VAT' AS VATTYPE" & Environment.NewLine
                sqlstr &= " FROM (CBTRHD"
                sqlstr &= "     LEFT JOIN CBTRDT ON (CBTRHD.BANKCODE = CBTRDT.BANKCODE) AND (CBTRHD.REFERENCE = CBTRDT.REFERENCE) AND (CBTRHD.UNIQUIFIER = CBTRDT.UNIQUIFIER)) "
                sqlstr &= "     LEFT JOIN CBTRHDO ON CBTRHDO.BANKCODE = CBTRDT.BANKCODE AND CBTRHDO.REFERENCE = CBTRDT.REFERENCE AND CBTRHDO.OPTFIELD = 'VOUCHER' AND CBTRHDO.UNIQUIFIER = CBTRHD.UNIQUIFIER "
                sqlstr &= "     LEFT JOIN CBMISC ON (CBTRHD.MISCCODE = CBMISC.MISCCODE ) "
                sqlstr &= "     LEFT JOIN APVEN ON (CBTRHD.MISCCODE = APVEN.VENDORID )" & Environment.NewLine

                If ComDatabase = "PVSW" Or ComDatabase = "ORCL" Then
                    sqlstr = sqlstr & "  WHERE (CBTRHD.""DATE"" >= '" & DATEFROM & "' AND CBTRHD.""DATE"" <= '" & DATETO & "')"
                ElseIf ComDatabase = "MSSQL" Then
                    sqlstr = sqlstr & "  WHERE (CBTRHD.[DATE] >= '" & DATEFROM & "' AND CBTRHD.[DATE] <= '" & DATETO & "')"
                End If
                sqlstr &= " AND CBTRDT.ACCTID IN (select ACCTVAT " & IIf(ComDatabase = "PVSW", "", "COLLATE Thai_CI_AS") & " from " & ConnVAT.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "FMSVLACC) "
                If CBREVERSE = True Then sqlstr &= " AND (CONVERT(varchar(8), CBTRHD.DATE)+':'+ltrim(rtrim(CBTRHD.REFERENCE))) not in (SELECT CONVERT(varchar(8), DATE)+':'+ltrim(rtrim(REFERENCE)) AS Expr1 FROM CBTRHD where reconciled = 3 GROUP BY  DATE,REFERENCE having count(*)>1 )"

            Else 'ข้อมูลที่ยังไม่ได้ Post ไม่ได้ใช่แล้วเนื่องจากบังคับให้ Post ก่อน Get vat
                If ComDatabase = "PVSW" Or ComDatabase = "ORCL" Then
                    sqlstr = "SELECT CBBTHD.""DATE"" ," '0-1
                ElseIf ComDatabase = "MSSQL" Then
                    sqlstr = "SELECT CBBTHD.[DATE] ," '0-1
                End If
                sqlstr &= "         CBBTHD.ENTRYTYPE AS INTERTYPE, CBBTHD.REFERENCE, CBBTHD.TEXTDESC, cbbtdt.DTLAMOUNT  * CBBTHD.BK2GLRATE AS Expr1, cbbtdt.QUANTITY, " & Environment.NewLine
                sqlstr &= "         CBBTDT.COMMENTS, CBBTDT.ACCTID, CBBTHD.BATCHID, CBBTHD.ENTRYNO, CBBTHD.MISCCODE, CBBTHD.BANKCODE, CBBTHD.REFERENCE AS Expr2, " & Environment.NewLine
                sqlstr &= "         '' AS UNIQUIFIER, CBBTDT.DETAILNO, CBBTHDO.VALUE,ISNULL(CBMISC.MISCNAME ,VENDNAME )AS Vname" & Environment.NewLine
                sqlstr &= " FROM    CBBTHD " & Environment.NewLine
                sqlstr &= " 		LEFT OUTER JOIN CBBTDT ON CBBTHD.BATCHID =CBBTDT.BATCHID and CBBTHD.ENTRYNO =CBBTDT.ENTRYNO  " & Environment.NewLine
                sqlstr &= " 		LEFT OUTER JOIN CBBTHDO ON CBBTHD.BATCHID =CBBTHDO.BATCHID and CBBTHD.ENTRYNO =CBBTHDO.ENTRYNO and CBBTHDO.OPTFIELD = 'VOUCHER' AND CBTRHDO.UNIQUIFIER = CBTRHD.UNIQUIFIER " & Environment.NewLine
                sqlstr &= "         LEFT JOIN CBMISC ON (CBTRHD.MISCCODE = CBMISC.MISCCODE ) " & Environment.NewLine
                sqlstr &= "         LEFT JOIN APVEN ON (cbtrhd.MISCCODE = APVEN .VENDORID ) " & Environment.NewLine

                If ComDatabase = "PVSW" Or ComDatabase = "ORCL" Then
                    sqlstr = sqlstr & "  WHERE (CBBTHD.""DATE"" >= '" & DATEFROM & "' AND CBBTHD.""DATE"" <= '" & DATETO & "')"
                ElseIf ComDatabase = "MSSQL" Then
                    sqlstr = sqlstr & "  WHERE (CBBTHD.[DATE] >= '" & DATEFROM & "' AND CBBTHD.[DATE] <= '" & DATETO & "')"
                End If

            End If

            rs.Open(sqlstr, cnACCPAC)
            If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
                WriteLog("<ProcessCB> Select CB Complete")
            End If


            Do While rs.options.Rs.EOF = False
                ErrMsg = "Ledger = CB IDINVC=" & rs.options.Rs.Collect(2)
                ErrMsg = ErrMsg & " Batch No.=" & rs.options.Rs.Collect(8)
                ErrMsg = ErrMsg & " Entry No.=" & rs.options.Rs.Collect(9)


                locbTAXID = ""
                locbBRANCH = ""
                locbDOCNO = ""
                MISC = ""
                Code = ""
                locbNAME = ""
                locbINVAMT = 0

                locbCBDATE = IIf(IsDBNull(rs.options.Rs.Collect(0)), "", rs.options.Rs.Collect(0)) ' CBTRHD.Date
                locbDOCNO = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", rs.options.Rs.Collect(2)) 'CBTRHD.REFERENCE
                loCBRef = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", rs.options.Rs.Collect(2)) 'CBTRHD.REFERENCE
                locbDescription = IIf(IsDBNull(rs.options.Rs.Collect(3)), "", rs.options.Rs.Collect(3)) 'CBTRHD.Description
                locbMiscCode = IIf(IsDBNull(rs.options.Rs.Collect(10)), "", rs.options.Rs.Collect(10)) 'CBTRHD.MISCCODE
                locbType = IIf(IsDBNull(rs.options.Rs.Collect(1)), 0, rs.options.Rs.Collect(1)) 'CBTRHD.INTERTYPE
                locbTXDATE = IIf(IsDBNull(rs.options.Rs.Collect(0)), 0, rs.options.Rs.Collect(0))
                locbVATCOMMENT = IIf(IsDBNull(rs.options.Rs.Collect(6)), "", (rs.options.Rs.Collect(6))) 'CBTRDT.VATCOMMENTs

                If rs.options.Rs.Collect("VATTYPE").ToString.Trim = "VAT" Then
                    locbINVTAX = IIf(IsDBNull(rs.options.Rs.Collect(4)), 0, rs.options.Rs.Collect(4)) 'CBTRDT.AMOUNT            
                ElseIf rs.options.Rs.Collect("VATTYPE").ToString.Trim = "VATAVG" Then
                    locbINVTAX = IIf(IsDBNull(rs.options.Rs.Collect("AMTTAXREC1")), 0, rs.options.Rs.Collect("AMTTAXREC1")) 'CBTRDT.AMTTAXREC1*CBTRHD.BK2GLRATE AS AMTTAXREC1          
                End If

                'locbBatch = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8)) 'CBTRHD.BATCHID
                'locbEntry = IIf(IsDBNull(rs.options.Rs.Collect(9)), 0, rs.options.Rs.Collect(9)) 'CBTRHD.ENTRYNO

                'If locbBatch = 88 And locbEntry = 1 Then
                '    Dim asss As String = ""
                'End If



                'ดึง Vat ที่ Comment
                If OPT_CB = False Or (Trim(locbVATCOMMENT).Split(",")).Length >= 3 Then
                    strsplit = Split(IIf(IsDBNull(rs.options.Rs.Collect(6)), "", rs.options.Rs.Collect(6)), ",") 'CBTRDT.VATCOMMENTs

                    For losplit = LBound(strsplit) To UBound(strsplit)
                        splittxt = Trim(strsplit(losplit))
                        Select Case losplit
                            Case 0
                                locbDOCNO = IIf(splittxt = "", locbDOCNO, splittxt)
                            Case 1
                                locbCBDATE = TryDate2(ZeroDate(splittxt, "dd/MM/yyyy", "yyyyMMdd"), "dd/MM/yyyy", "yyyyMMdd")
                                If locbCBDATE = 0 Then
                                    locbCBDATE = TryDate2(splittxt, "dd/MM/yyyy", "yyyyMMdd")
                                End If
                            Case 2
                                locbNAME = locbNAME & splittxt
                                Dim rsMISS As New BaseClass.BaseODBCIO
                                'rsMISS.Open("select TOP 1 * from(select MISCCODE as CODE , MISCNAME as NAME from CBMISC union all select VENDORID as CODE , VENDNAME as NAME  from APVEN union all select IDCUST as CODE , NAMECUST as NAME  from ARCUS ) As TMP where CODE ='" & IIf(locbNAME.Trim = "", rs.options.Rs.Collect("MISCCODE"), locbNAME.Replace("'", "''")) & "'", cnACCPAC)

                                If locbType = 0 Then 'Type Cashbook
                                    rsMISS.Open(" select MISCCODE as CODE , MISCNAME as NAME from CBMISC where MISCCODE ='" & IIf(locbNAME.Trim = "", rs.options.Rs.Collect("MISCCODE"), locbNAME.Replace("'", "''")) & "'", cnACCPAC)
                                ElseIf locbType = 1 Then 'Type Account Payable
                                   
                                    rsMISS.Open(" select VENDORID as CODE , VENDNAME as NAME  from APVEN where VENDORID ='" & IIf(locbNAME.Trim = "", rs.options.Rs.Collect("MISCCODE"), locbNAME.Replace("'", "''")) & "'", cnACCPAC)
                                   
                                ElseIf locbType = 2 Then 'Type Account Receivable
                                    rsMISS.Open(" select IDCUST as CODE , NAMECUST as NAME  from ARCUS where IDCUST ='" & IIf(locbNAME.Trim = "", rs.options.Rs.Collect("MISCCODE"), locbNAME.Replace("'", "''")) & "'", cnACCPAC)
                                End If

                                If rsMISS.options.QueryDT.Rows.Count > 0 Then
                                    locbNAME = rsMISS.options.QueryDT.Rows(0).Item("NAME").ToString
                                    'บรรทัดนี้เก็บข้อมูลผิด ทำให้ Taxid ไม่แสดง 
                                    'MISC = IIf(locbNAME.Trim = "", rsMISS.options.Rs.Collect("CODE"), locbNAME.Replace("'", "''"))
                                    'บรรทัดนี้ที่เก็บถูกต้อง
                                    MISC = rsMISS.options.Rs.Collect("CODE")

                                End If
                            Case Is > 2
                                locbNAME = locbNAME & "," & splittxt
                                Dim rsMISS As New BaseClass.BaseODBCIO
                                'rsMISS.Open("select TOP 1 * from(select MISCCODE as CODE , MISCNAME as NAME from CBMISC union all select VENDORID as CODE , VENDNAME as NAME  from APVEN union all select IDCUST as CODE , NAMECUST as NAME  from ARCUS ) As TMP where CODE ='" & IIf(locbNAME.Trim = "", rs.options.Rs.Collect("MISCCODE"), locbNAME.Replace("'", "''")) & "'", cnACCPAC)

                                If locbType = 0 Then 'Type Cashbook
                                    rsMISS.Open(" select MISCCODE as CODE , MISCNAME as NAME from CBMISC where MISCCODE ='" & IIf(locbNAME.Trim = "", rs.options.Rs.Collect("MISCCODE"), locbNAME.Replace("'", "''")) & "'", cnACCPAC)
                                ElseIf locbType = 1 Then 'Type Account Payable
                                    If Use_LEGALNAME_inAP = True Then
                                        rsMISS.Open("select IDVEND AS CODE, RMITNAME AS NAME FROM APVNR WHERE IDVEND ='" & IIf(locbNAME.Trim = "", rs.options.Rs.Collect("MISCCODE"), locbNAME.Replace("'", "''")) & "'", cnACCPAC)
                                    Else
                                        rsMISS.Open(" select VENDORID as CODE , VENDNAME as NAME  from APVEN where VENDORID ='" & IIf(locbNAME.Trim = "", rs.options.Rs.Collect("MISCCODE"), locbNAME.Replace("'", "''")) & "'", cnACCPAC)
                                    End If

                                ElseIf locbType = 2 Then 'Type Account Receivable
                                    rsMISS.Open(" select IDCUST as CODE , NAMECUST as NAME  from ARCUS where IDCUST ='" & IIf(locbNAME.Trim = "", rs.options.Rs.Collect("MISCCODE"), locbNAME.Replace("'", "''")) & "'", cnACCPAC)
                                End If

                                If rsMISS.options.QueryDT.Rows.Count > 0 Then
                                    locbNAME = rsMISS.options.QueryDT.Rows(0).Item("NAME").ToString
                                    'บรรทัดนี้เก็บข้อมูลผิด ทำให้ Taxid ไม่แสดง 
                                    'MISC = IIf(locbNAME.Trim = "", rsMISS.options.Rs.Collect("CODE"), locbNAME.Replace("'", "''"))
                                    'บรรทัดนี้ที่เก็บถูกต้อง
                                    MISC = rsMISS.options.Rs.Collect("CODE")
                                End If
                        End Select
                    Next

                    If Len(MISC.Trim) = 0 And locbNAME.Trim = "" Then
                        MISC = IIf(IsDBNull(rs.options.Rs.Collect("MISCCODE")), "", rs.options.Rs.Collect("MISCCODE"))
                        locbNAME = IIf(IsDBNull(rs.options.Rs.Collect("Vname")), "", rs.options.Rs.Collect("Vname"))
                    End If

                Else 'ดึง Vat ที่ Optional Fields

                    Dim RS_WHT As New BaseClass.BaseODBCIO("select * from CBTRDTO where OPTFIELD in ('TAXBASE','TAXINVNO','TAXDATE','TAXNAME','TAXID','BRANCH') and BANKCODE = '" & rs.options.Rs.Collect("BANKCODE") & "' and UNIQUIFIER = " & rs.options.Rs.Collect("UNIQUIFIER") & " and REFERENCE ='" & rs.options.Rs.Collect("REFERENCE") & "' and DETAILNO ='" & rs.options.Rs.Collect("DETAILNO") & "'", cnACCPAC)

                    For iwht As Integer = 0 To RS_WHT.options.QueryDT.Rows.Count - 1

                        If RS_WHT.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXBASE" Then
                            If IsNumeric(RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE")) = True Then
                                locbINVAMT = IIf(Trim(RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE")) = "", 0, Trim(RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE")))
                            End If

                        ElseIf RS_WHT.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXINVNO" Then
                            locbDOCNO = Trim(RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE"))

                        ElseIf RS_WHT.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXDATE" Then

                            If IsNumeric(RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE")) = True Then 'And RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE") <> 0
                                If RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE") <> 0 Then
                                    locbCBDATE = RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE").ToString.Trim
                                End If
                            Else
                                Dim CBDATE As String = ""
                                Dim ConDate As Date
                                Dim OPT_CBDATE As String = RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE").ToString().Trim

                                Try

                                    If OPT_CBDATE <> "" Then
                                        CBDATE = TryDate2(OPT_CBDATE, "dd/MM/yyyy", "yyyyMMdd")
                                    End If

                                Catch ex As Exception

                                    ConDate = FormatDateTime(OPT_CBDATE, DateFormat.ShortDate)
                                    Dim ArDate As Object = ConDate.GetDateTimeFormats
                                    ConDate = ArDate(0)
                                    CBDATE = TryDate2(ConDate, "d/M/yyyy", "yyyyMMdd")

                                End Try

                                If OPT_CBDATE <> "" And CBDATE = "0" Then

                                    ConDate = FormatDateTime(OPT_CBDATE, DateFormat.ShortDate)
                                    Dim ArDate As Object = ConDate.GetDateTimeFormats
                                    ConDate = ArDate(0)
                                    CBDATE = TryDate2(ConDate, "d/M/yyyy", "yyyyMMdd")
                                End If



                                locbCBDATE = IIf(CBDATE = "", locbCBDATE, CBDATE)

                            End If

                            locbCBDATE = IIf(locbCBDATE = 0, rs.options.Rs.Collect(0), locbCBDATE)

                        ElseIf RS_WHT.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXNAME" Then
                            If RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE").Trim = "" Then
                                locbNAME = ""
                            Else
                                locbNAME = Trim(RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE"))
                            End If
                            Dim rsMISS As New BaseClass.BaseODBCIO

                            If locbType = 0 Then 'Type Cashbook
                                If Use_LEGALNAME_inCB = True Then
                                    rsMISS.Open(" select MISCCODE as CODE , CASE WHEN COMMENTS IS NULL THEN MISCNAME ELSE COMMENTS END as NAME from CBMISC where MISCCODE ='" & IIf(locbNAME.Trim = "", rs.options.Rs.Collect("MISCCODE"), locbNAME.Replace("'", "''")) & "'", cnACCPAC)
                                    If rsMISS.options.QueryDT.Rows.Count = 0 Then
                                        rsMISS.Open(" select MISCCODE as CODE , MISCNAME as NAME from CBMISC where MISCCODE ='" & IIf(locbNAME.Trim = "", rs.options.Rs.Collect("MISCCODE"), locbNAME.Replace("'", "''")) & "'", cnACCPAC)
                                    Else
                                    End If
                                Else
                                    rsMISS.Open(" select MISCCODE as CODE , MISCNAME as NAME from CBMISC where MISCCODE ='" & IIf(locbNAME.Trim = "", rs.options.Rs.Collect("MISCCODE"), locbNAME.Replace("'", "''")) & "'", cnACCPAC)
                                End If

                            ElseIf locbType = 1 Then 'Type Account Payable
                                If Use_LEGALNAME_inAP = True Then
                                    rsMISS.Open("select IDVEND AS CODE, CASE WHEN RMITNAME IS NULL THEN APVEN.VENDNAME ELSE RMITNAME END AS NAME FROM APVNR INNER JOIN APVEN ON APVNR.IDVEND = APVEN.VENDORID WHERE IDVENDRMIT = 'TXTHAI' AND IDVEND ='" & IIf(locbNAME.Trim = "", rs.options.Rs.Collect("MISCCODE"), locbNAME.Replace("'", "''")) & "'", cnACCPAC)
                                    If rsMISS.options.QueryDT.Rows.Count = 0 Then
                                        rsMISS.Open(" select VENDORID as CODE , VENDNAME as NAME  from APVEN where VENDORID ='" & IIf(locbNAME.Trim = "", rs.options.Rs.Collect("MISCCODE"), locbNAME.Replace("'", "''")) & "'", cnACCPAC)
                                    Else
                                    End If
                                Else
                                    rsMISS.Open(" select VENDORID as CODE , VENDNAME as NAME  from APVEN where VENDORID ='" & IIf(locbNAME.Trim = "", rs.options.Rs.Collect("MISCCODE"), locbNAME.Replace("'", "''")) & "'", cnACCPAC)
                                End If

                            ElseIf locbType = 2 Then 'Type Account Receivable
                                If Use_LEGALNAME_inAR = True Then
                                    rsMISS.Open("SELECT NAMELOCN FROM ARCSP WHERE IDCUSTSHPT = 'TXTHAI' AND  ARCSP.IDCUST ='" & IIf(locbNAME.Trim = "", rs.options.Rs.Collect("MISCCODE"), locbNAME.Replace("'", "''")) & "'", cnACCPAC)
                                    If rsMISS.options.QueryDT.Rows.Count = 0 Then
                                        rsMISS.Open(" select IDCUST as CODE , NAMECUST as NAME  from ARCUS where IDCUST ='" & IIf(locbNAME.Trim = "", rs.options.Rs.Collect("MISCCODE"), locbNAME.Replace("'", "''")) & "'", cnACCPAC)
                                    Else
                                        rsMISS.Open("SELECT ARCSP.IDCUST AS CODE ,CASE WHEN NAMELOCN IS NULL THEN ARCUS.NAMECUST ELSE NAMELOCN END AS NAME FROM ARCSP LEFT OUTER JOIN ARCUS ON ARCSP.IDCUST = ARCUS.IDCUST WHERE IDCUSTSHPT = 'TXTHAI' AND  ARCSP.IDCUST ='" & IIf(locbNAME.Trim = "", rs.options.Rs.Collect("MISCCODE"), locbNAME.Replace("'", "''")) & "'", cnACCPAC)
                                    End If
                                Else
                                    rsMISS.Open(" select IDCUST as CODE , NAMECUST as NAME  from ARCUS where IDCUST ='" & IIf(locbNAME.Trim = "", rs.options.Rs.Collect("MISCCODE"), locbNAME.Replace("'", "''")) & "'", cnACCPAC)
                                End If
                            End If

                            ' rsMISS.Open("select TOP 1 * from(select MISCCODE as CODE , MISCNAME as NAME from CBMISC union all select VENDORID as CODE , VENDNAME as NAME  from APVEN union all select IDCUST as CODE , NAMECUST as NAME  from ARCUS ) As TMP where CODE ='" & IIf(locbNAME.Trim = "", rs.options.Rs.Collect("MISCCODE"), locbNAME.Replace("'", "''")) & "'", cnACCPAC)


                            If rsMISS.options.QueryDT.Rows.Count > 0 Then
                                locbNAME = rsMISS.options.QueryDT.Rows(0).Item("NAME").ToString
                                MISC = rsMISS.options.QueryDT.Rows(0).Item("CODE").ToString
                            End If

                        ElseIf RS_WHT.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "TAXID" Then
                            If RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE").Trim = "" Then
                                locbTAXID = ""
                            Else
                                locbTAXID = Trim(RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE"))
                            End If

                        ElseIf RS_WHT.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "BRANCH" Then
                            If RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE").Trim = "" Then
                                locbBRANCH = ""
                            Else
                                locbBRANCH = Trim(RS_WHT.options.QueryDT.Rows(iwht).Item("VALUE"))
                            End If
                        End If

                    Next

                    If Len(MISC.Trim) = 0 And locbNAME.Trim = "" Then
                        MISC = IIf(IsDBNull(rs.options.Rs.Collect("MISCCODE")), "", rs.options.Rs.Collect("MISCCODE"))
                        'locbNAME = IIf(IsDBNull(rs.options.Rs.Collect("Vname")), "", rs.options.Rs.Collect("Vname"))
                    End If

                End If 'OPT_CB = False เช็ค Vat ดึงจากที่ไหน

                Code = MISC

           

                If locbTAXID = "" And locbBRANCH = "" Then

                    If locbType = 0 Then 'Type Cashbook
                        locbTAXID = FindTaxIDBRANCH(MISC, "CB", "TAXID")
                        locbBRANCH = FindTaxIDBRANCH(MISC, "CB", "BRANCH")
                        If locbNAME.Trim = "" Then locbNAME = FindMiscCodeInMisccode(MISC)
                    ElseIf locbType = 1 Then 'Type Account Payable
                        locbTAXID = FindTaxIDBRANCH(MISC, "AP", "TAXID")
                        locbBRANCH = FindTaxIDBRANCH(MISC, "AP", "BRANCH")
                        If locbNAME.Trim = "" Then locbNAME = FindMiscCodeInAPVEN(MISC, "AP")
                    ElseIf locbType = 2 Then 'Type Account Receivable
                        locbTAXID = FindTaxIDBRANCH(MISC, "AR", "TAXID")
                        locbBRANCH = FindTaxIDBRANCH(MISC, "AR", "BRANCH")
                        If locbNAME.Trim = "" Then locbNAME = FindMiscCodeInARCUS(MISC)
                    End If

                End If


                locbVATCOMMENT = locbDOCNO & "," & locbCBDATE & "," & locbNAME
                If locbINVAMT = 0 Then
                    locbINVAMT = IIf(IsDBNull(rs.options.Rs.Collect(5)), 0, rs.options.Rs.Collect(5)) 'CBTRDT.QUANTITY
                End If

                'If rs.options.Rs.Collect("VATTYPE").ToString.Trim = "VAT" Then
                '    locbINVTAX = IIf(IsDBNull(rs.options.Rs.Collect(4)), 0, rs.options.Rs.Collect(4)) 'CBTRDT.AMOUNT            
                'ElseIf rs.options.Rs.Collect("VATTYPE").ToString.Trim = "VATAVG" Then
                '    locbINVTAX = IIf(IsDBNull(rs.options.Rs.Collect("AMTTAXREC1")), 0, rs.options.Rs.Collect("AMTTAXREC1")) 'CBTRDT.AMTTAXREC1*CBTRHD.BK2GLRATE AS AMTTAXREC1          
                'End If


                If VNO_CB = True Then

                    Dim VNO_Detail As New BaseClass.BaseODBCIO("select isnull(OPTFIELD,'') as 'OPTFIELD',isnull(VALUE,'')as 'VALUE' from CBTRDTO  where OPTFIELD in ('VOUCHER') and BANKCODE = '" & rs.options.Rs.Collect("BANKCODE") & "' and UNIQUIFIER = " & rs.options.Rs.Collect("UNIQUIFIER") & " and REFERENCE ='" & rs.options.Rs.Collect("REFERENCE") & "' and DETAILNO ='" & rs.options.Rs.Collect("DETAILNO") & "'", cnACCPAC)
                    For iwht As Integer = 0 To VNO_Detail.options.QueryDT.Rows.Count - 1
                        If VNO_Detail.options.QueryDT.Rows(iwht).Item("OPTFIELD").ToString.Trim = "VOUCHER" Then
                            If Trim(VNO_Detail.options.QueryDT.Rows(iwht).Item("VALUE")).Trim <> "" Then
                                loCBRef = Trim(VNO_Detail.options.QueryDT.Rows(iwht).Item("VALUE"))
                            End If

                        End If
                    Next

                End If


                locbACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect(7)), "", (rs.options.Rs.Collect(7))) 'AS ACCTID
                locbBatch = IIf(IsDBNull(rs.options.Rs.Collect(8)), 0, rs.options.Rs.Collect(8)) 'CBTRHD.BATCHID
                locbEntry = IIf(IsDBNull(rs.options.Rs.Collect(9)), 0, rs.options.Rs.Collect(9)) 'CBTRHD.ENTRYNO
                locbTrans = IIf(IsDBNull(rs.options.Rs.Collect(14)), 0, rs.options.Rs.Collect(14)) 'CBTRDT.DETAILNO
                loRunno = Trim(IIf(IsDBNull(rs.options.Rs.Collect("Value")), "", rs.options.Rs.Collect("Value"))) 'CBTRHDO.VALUE
                locbTOTTAX = Trim(IIf(IsDBNull(rs.options.Rs.Collect("AMTTAX1")), 0, rs.options.Rs.Collect("AMTTAX1"))) 'AS AMTTAX1


                sqlstr = InsCB(locbBatch, locbEntry, locbCBDATE, locbDOCNO, loCBRef, locbNAME, locbINVAMT, locbINVTAX, Trim(locbVATCOMMENT), locbACCTVAT, locbTrans, locbTXDATE, loRunno, locbTAXID, locbBRANCH, Code, locbTOTTAX)

                cnVAT.Execute(sqlstr)
                rs.options.Rs.MoveNext()
                Application.DoEvents()
              
            Loop
            rs.options.Rs.Close()
            rs = Nothing

            'transfer data
            rs = New BaseClass.BaseODBCIO
            sqlstr = "SELECT FMSCB.CBDATE,FMSCB.DOCNO,FMSCB.NAME," '0-2
            sqlstr = sqlstr & " FMSCB.INVAMT , FMSCB.INVTAX, FMSVLACC.LOCID, " '3-5
            sqlstr = sqlstr & " FMSCB.ACCTVAT, FMSVLACC.VTYPE, FMSCB.VATCOMMENT," '6-8
            sqlstr = sqlstr & " FMSCB.BATCH , FMSCB.ENTRY, FMSCB.REFERENCE,FMSCB.TRANSNBR,FMSCB.TXDATE,FMSCB.RUNNO," '9-13
            sqlstr = sqlstr & " FMSCB.TAXID,FMSCB.BRANCH,FMSCB.Code,FMSCB.TOTTAX "
            If ComDatabase = "ORCL" Then
                sqlstr = sqlstr & " FROM FMSCB INNER JOIN FMSVLACC ON FMSCB.ACCTVAT = FMSVLACC.ACCTVAT"
                sqlstr = sqlstr & " WHERE (FMSCB.TXDATE BETWEEN '" & DATEFROM & "' AND '" & DATETO & "')"
            Else
                sqlstr = sqlstr & " FROM FMSCB INNER JOIN FMSVLACC ON FMSCB.ACCTVAT = FMSVLACC.ACCTVAT"
                sqlstr = sqlstr & " WHERE (FMSCB.TXDATE BETWEEN '" & DATEFROM & "' AND '" & DATETO & "')"
            End If

            rs.Open(sqlstr, cnVAT)
            If rs.options.Rs.EOF = False Or rs.options.Rs.EOF = True Then
                WriteLog("<ProcessCB> Select FMSCB Complete")
            End If


            Do While rs.options.Rs.EOF = False
                loCls = New clsVat
                With loCls
                    .INVDATE = IIf(IsDBNull(rs.options.Rs.Collect(0)), 0, rs.options.Rs.Collect(0))
                    .TXDATE = IIf(IsDBNull(rs.options.Rs.Collect(13)), 0, rs.options.Rs.Collect(13))
                    .INVNAME = IIf(IsDBNull(rs.options.Rs.Collect(2)), 0, (rs.options.Rs.Collect(2)))
                    .IDINVC = IIf(IsDBNull(rs.options.Rs.Collect(1)), "", (rs.options.Rs.Collect(1)))
                    .DOCNO = IIf(IsDBNull(rs.options.Rs.Collect(11)), "", (rs.options.Rs.Collect(11)))
                    .TTYPE = IIf(IsDBNull(rs.options.Rs.Collect(7)), 0, rs.options.Rs.Collect(7))
                    .INVAMT = IIf(IsDBNull(rs.options.Rs.Collect(3)), 0, rs.options.Rs.Collect(3)) 'Up to user key
                    .INVTAX = IIf(IsDBNull(rs.options.Rs.Collect(4)), 0, rs.options.Rs.Collect(4))
                    .TOTTAX = Trim(IIf(IsDBNull(rs.options.Rs.Collect("TOTTAX")), 0, rs.options.Rs.Collect("TOTTAX")))

                    '** Purchase => -tax  Sale => Tax

                    If .TTYPE = 1 Or .TTYPE = 3 Then
                        .INVTAX = .INVTAX * (-1)
                        .TOTTAX = .TOTTAX * (-1)
                    End If

                    If .INVTAX < 0 Then
                        .INVAMT = .INVAMT * (-1)
                    End If


                    If .INVAMT = 0 Then
                        .RATE_Renamed = 0
                    Else
                        .RATE_Renamed = FormatNumber(((.INVTAX / .INVAMT) * 100), 1)
                    End If


                    .LOCID = IIf(IsDBNull(rs.options.Rs.Collect(5)), "", rs.options.Rs.Collect(5))
                    .ACCTVAT = IIf(IsDBNull(rs.options.Rs.Collect(6)), "", rs.options.Rs.Collect(6))
                    .VATCOMMENT = IIf(IsDBNull(rs.options.Rs.Collect(8)), "", (rs.options.Rs.Collect(8)))
                    .Batch = IIf(IsDBNull(rs.options.Rs.Collect(9)), 0, rs.options.Rs.Collect(9))
                    .Entry = IIf(IsDBNull(rs.options.Rs.Collect(10)), 0, rs.options.Rs.Collect(10))
                    .CBRef = IIf(IsDBNull(rs.options.Rs.Collect(11)), "", (rs.options.Rs.Collect(11)))
                    .TRANSNBR = IIf(IsDBNull(rs.options.Rs.Collect(12)), 0, rs.options.Rs.Collect(12))
                    .Runno = Trim(IIf(IsDBNull(rs.options.Rs.Collect("Runno")), "", rs.options.Rs.Collect("Runno")))
                    .TAXID = Trim(IIf(IsDBNull(rs.options.Rs.Collect("TAXID")), "", rs.options.Rs.Collect("TAXID")))
                    .Branch = Trim(IIf(IsDBNull(rs.options.Rs.Collect("BRANCH")), "", rs.options.Rs.Collect("BRANCH")))
                    Code = Trim(IIf(IsDBNull(rs.options.Rs.Collect("Code")), "", rs.options.Rs.Collect("Code")))
                    .CODETAX = ""


                    sqlstr = InsVat(1, CStr(.INVDATE), CStr(.TXDATE), .IDINVC, .DOCNO, .NEWDOCNO, .INVNAME, .INVAMT, .INVTAX, .LOCID, 1, .RATE_Renamed, .TTYPE, .ACCTVAT, "CB", .Batch, .Entry, .MARK, Trim(.VATCOMMENT), .CBRef, .TRANSNBR, .Runno, .CODETAX, "", "", "", 0, 0, 0, .TAXID, .Branch, Code, "", 0, 0, .TOTTAX)

                End With
                cnVAT.Execute(sqlstr)
                rs.options.Rs.MoveNext()
                Application.DoEvents()
            Loop
            rs.options.Rs.Close()
            rs = Nothing
            cnACCPAC.Close()
            cnVAT.Close()
            cnACCPAC = Nothing
            cnVAT = Nothing
            Exit Sub

            'ErrHandler:
        Catch ex As Exception
            WriteLog("<ProcessCB>" & Err.Number & "  " & Err.Description)
            WriteLog("<ProcessCB> sqlstr : " & sqlstr)
            If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
            WriteLog(ErrMsg)
        End Try

    End Sub

    Public Sub ProcessUpdateRate(ByRef Ratefrom As Double, ByRef RateTo As Double, ByRef RateNew As Double)
        Dim cn As ADODB.Connection
        Dim sqlstr As String
        cn = New ADODB.Connection
        cn.Open(ConnVAT)
        sqlstr = "UPDATE FMSVAT set RATE = " & RateNew
        sqlstr = sqlstr & " WHERE RATE between " & Ratefrom & " and " & RateTo
        cn.Execute(sqlstr)
        cn.Close()
        WriteLog("<ProcessUpdateRate> ready")
        cn = Nothing
        Exit Sub
ErrHandler:
        cn.RollbackTrans()
        ErrMsg = Err.Description
        WriteLog("<ProcessUpdateRate>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessUpdateRate> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    Public Function CountCharacter(ByVal value As String, ByVal ch As Char) As String
        Dim cnt As Integer = 0
        For Each c As Char In value
            If c = ch Then cnt += 1
        Next
        Dim result As String = ""
        For i As Integer = 0 To cnt - 1
            result += ch
        Next
        Return result
    End Function

    Function SQLNewDocNO(ByVal loPrefix As String, ByVal SQ As String, ByVal Type As Integer) As String
        Try
            FormatRunnig = FormatRunnig
            SQ = SQ.Trim
            Dim Tx As String = FormatRunnig.ToUpper.Replace("/", "'/'")
            Dim T As String = FormatRunnig.Replace("/", "'/'")
            Dim Y As String = CountCharacter(Tx, "Y")
            Dim M As String = CountCharacter(Tx, "M")
            If Len(Y) <> 0 Then Tx = Tx.Replace(Y, " RIGHT(LEFT(LTRIM(RTRIM(" + SQ + ")),4)," + Y.Length.ToString + ")")
            If Len(M) <> 0 Then Tx = Tx.Replace(M, " RIGHT(LEFT(LTRIM(RTRIM(" + SQ + ")),6),2)")
            Select Case Type
                Case 0
                    Return loPrefix & "+" & Tx
                Case 1
                    Return Tx
                Case 2
                    Dim ST As String = TryDate2(SQ.Trim, "yyyyMMdd", T.Replace("+", "").Replace("'", ""))
                    Return ST
                Case 3
                    Dim ST As String = TryDate2(SQ.Trim, "yyyyMMdd", T.Replace("+", "").Replace("'", ""))
                    Return loPrefix & ST
            End Select
        Catch ex As Exception

        End Try

        Return SQ
    End Function

    Public Sub ProcessNewDocNoCustom()
        On Error GoTo ErrHandler
        Dim cn, cnACCPAC As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr, abc As String
        Dim loPrefix As String
        Dim loRunning As Integer = 0
        Dim PreLocation As String
        cn = New ADODB.Connection
        cnACCPAC = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        cn.ConnectionTimeout = 60
        cn.Open(ConnVAT)
        cn.CommandTimeout = 3600

        cnACCPAC.ConnectionTimeout = 60
        cnACCPAC.Open(ConnACCPAC)
        cnACCPAC.CommandTimeout = 3600
        PreLocation = ""
        Dim IsSource As String = ""
        Dim Location As New Collection

        If RERUN Then
            sqlstr = "UPDATE FMSVAT " & Environment.NewLine
            sqlstr &= "SET " & Environment.NewLine
            sqlstr &= "FMSVAT.NEWDOCNO = " & SQLNewDocNO("FMSVLOC.LOCPREFIX", "FMSVAT.TXDATE", 0) & "+FMSRUN.RUNNING " & Environment.NewLine
            sqlstr &= "FROM FMSVAT " & Environment.NewLine
            sqlstr &= "LEFT JOIN FMSVLOC ON FMSVAT.LOCID = FMSVLOC.LOCID " & Environment.NewLine
            sqlstr &= "LEFT JOIN FMSRUN ON (FMSVAT.BATCH =FMSRUN.BATCH) AND (FMSVAT.ENTRY = FMSRUN.ENTRY  ) AND ((CASE WHEN isnull(FMSVAT.TRANSNBR,0)='' then 0 else isnull(FMSVAT.TRANSNBR,0) end) = (CASE WHEN isnull(FMSRUN.TRANSNBR,0)='' then 0 else isnull(FMSRUN.TRANSNBR,0) end))and (FMSVAT.SOURCE = FMSRUN.SOURCE  )" & Environment.NewLine
            sqlstr &= "WHERE ISNULL (FMSRUN.RUNNING,'')<>'';" & Environment.NewLine

            sqlstr &= "UPDATE FMSVATINSERT" & Environment.NewLine
            sqlstr &= "SET " & Environment.NewLine
            sqlstr &= "FMSVATINSERT.NEWDOCNO  =" & SQLNewDocNO("FMSVLOC.LOCPREFIX", "FMSVATINSERT.TXDATE", 0) & "+FMSRUN.RUNNING " & Environment.NewLine
            sqlstr &= "FROM FMSVATINSERT " & Environment.NewLine
            sqlstr &= "LEFT JOIN FMSVLOC ON FMSVATINSERT.LOCID = FMSVLOC.LOCID " & Environment.NewLine
            sqlstr &= "LEFT JOIN FMSRUN ON (FMSVATINSERT.BATCH = FMSRUN.BATCH) AND (FMSVATINSERT.ENTRY =FMSRUN.ENTRY) AND ((CASE WHEN isnull(FMSVATINSERT.TRANSNBR,0)='' then 0 else isnull(FMSVATINSERT.TRANSNBR,0) end) = (CASE WHEN isnull(FMSRUN.TRANSNBR,0)='' then 0 else isnull(FMSRUN.TRANSNBR,0) end))and (FMSVATINSERT.SOURCE  = FMSRUN.SOURCE)" & Environment.NewLine
            sqlstr &= "WHERE ISNULL (FMSRUN.RUNNING,'')<>'';" & Environment.NewLine
            cn.Execute(sqlstr)

            Dim RS_Location As New BaseClass.BaseODBCIO("select  isnull(max(RUNNING),0)as RUNNING ,LOCID ," & SQLNewDocNO("", "TXDATE", 1) & " as TXDATE from FMSVAT  LEFT JOIN FMSRUN on (FMSVAT.BATCH =FMSRUN.BATCH) and (FMSVAT.ENTRY =FMSRUN.ENTRY  ) and ((CASE WHEN isnull(FMSVAT.TRANSNBR,0)='' then 0 else isnull(FMSVAT.TRANSNBR,0) end) =(CASE WHEN isnull(FMSRUN.TRANSNBR,0)='' then 0 else isnull(FMSRUN.TRANSNBR,0) end)) and (FMSVAT.SOURCE  =FMSRUN.SOURCE  ) group by LOCID," & SQLNewDocNO("", "TXDATE", 1), cn)
            For iLoca As Integer = 0 To RS_Location.options.QueryDT.Rows.Count - 1
                Location.Add(New String() {RS_Location.options.QueryDT.Rows(iLoca).Item("LOCID"), RS_Location.options.QueryDT.Rows(iLoca).Item("TXDATE"), RS_Location.options.QueryDT.Rows(iLoca).Item("RUNNING")}, RS_Location.options.QueryDT.Rows(iLoca).Item("LOCID") + RS_Location.options.QueryDT.Rows(iLoca).Item("TXDATE"))
            Next
        End If

        sqlstr = "SELECT FMSVAT.TXDATE,FMSVAT.DOCNO,FMSVAT.VATID,FMSVAT.INVDATE,FMSVLOC.LOCPREFIX,FMSVLOC.LOCID,FMSVAT.BATCH ,FMSVAT.ENTRY ,(CASE WHEN isnull(FMSVAT.TRANSNBR,0)='' then 0 else isnull(FMSVAT.TRANSNBR,0) end)as LINE ,FMSVAT.SOURCE" & IIf(RERUN, " ,isnull(FMSRUN.RUNNING,'') as RUNNING", "") & ",FMSVAT.NEWDOCNO " & Environment.NewLine
        sqlstr &= ",(CASE	WHEN ltrim(rtrim(FMSVAT.SOURCE))in('AP','PO') then (select top 1 APOBL.AUDTDATE from " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & " APOBL left join " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "APIBD on (APOBL.CNTBTCH =APIBD.CNTBTCH and APOBL.CNTITEM =APIBD.CNTITEM ) where APOBL.CNTBTCH =FMSVAT.BATCH  and  APOBL.CNTITEM =FMSVAT.ENTRY) " & Environment.NewLine
        sqlstr &= "		    WHEN ltrim(rtrim(FMSVAT.SOURCE))='GL' then (select top 1 GLPOST.AUDTDATE  from " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "GLPOST left join " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "GLJED  on (GLPOST.BATCHNBR =GLJED.BATCHNBR and GLPOST.ENTRYNBR =GLJED.JOURNALID ) where GLPOST.BATCHNBR=convert(char(6),FMSVAT.BATCH)  and  GLPOST.ENTRYNBR =convert(char(6),FMSVAT.ENTRY)) " & Environment.NewLine
        If FindApp("CB") = True Then
            sqlstr &= "		    WHEN ltrim(rtrim(FMSVAT.SOURCE))='CB' then (select top 1 CBTRHD.AUDTDATE  from " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "CBTRHD left join " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "CBBTDT   on (CBTRHD.BATCHID  =CBBTDT.BATCHID and CBTRHD.ENTRYNO =CBBTDT.ENTRYNO  ) where CBTRHD.BATCHID=FMSVAT.BATCH  and  CBTRHD.ENTRYNO =FMSVAT.ENTRY) " & Environment.NewLine
        End If
        sqlstr &= "		    WHEN ltrim(rtrim(FMSVAT.SOURCE))in('AR','OE') then (select top 1 AROBL.AUDTDATE  from " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "AROBL  left join " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "ARIBD    on (AROBL.CNTBTCH  =ARIBD.CNTBTCH and AROBL.CNTITEM =ARIBD.CNTITEM  ) where AROBL.CNTBTCH=FMSVAT.BATCH  and  ARIBD.CNTITEM =FMSVAT.ENTRY) " & Environment.NewLine
        sqlstr &= "end) as AUDTDATE" & Environment.NewLine
        If ComDatabase = "ORCL" Then
            sqlstr = sqlstr & " FROM FMSVAT LEFT JOIN FMSVLOC ON FMSVAT.LOCID = FMSVLOC.LOCID" & Environment.NewLine
        Else
            sqlstr = sqlstr & " FROM FMSVAT LEFT JOIN FMSVLOC ON FMSVAT.LOCID = FMSVLOC.LOCID" & Environment.NewLine
        End If
        If RERUN Then sqlstr &= " LEFT JOIN FMSRUN on (FMSVAT.BATCH =FMSRUN.BATCH) and (FMSVAT.ENTRY =FMSRUN.ENTRY  ) and ((CASE WHEN isnull(FMSVAT.TRANSNBR,0)='' then 0 else isnull(FMSVAT.TRANSNBR,0) end) =(CASE WHEN isnull(FMSRUN.TRANSNBR,0)='' then 0 else isnull(FMSRUN.TRANSNBR,0)end))" & Environment.NewLine
        If CheckTTypeRunning = 1 Then
            sqlstr = sqlstr & " WHERE TTYPE = 1" & Environment.NewLine
        ElseIf CheckTTypeRunning = 2 Then
            sqlstr = sqlstr & " WHERE TTYPE = 2" & Environment.NewLine
        ElseIf CheckTTypeRunning = 3 Then
            sqlstr = sqlstr & " WHERE TTYPE = 3" & Environment.NewLine
        End If
        If CheckRate.Trim.Split("||").Length = 3 Then
            sqlstr &= " and RATE = " & CheckRate.Trim.Split("||")(2) & Environment.NewLine
        ElseIf CheckRate.Trim = "None" Then
            sqlstr &= " and (VTYPE = 0)" & Environment.NewLine
        ElseIf CheckRate.Trim = "unlike 7" Then
            sqlstr &= " and (RATE <> 7 and RATE <> 0)" & Environment.NewLine
        End If
        sqlstr &= " ORDER BY FMSVAT.INVDATE,IDINVC,BATCH,ENTRY,LINE"
        rs.Open(sqlstr, cn)
        Dim collvat As New Collection
        Dim TMPTXDATE As String = ""
        Do While rs.options.Rs.EOF = False
            Dim RunNO As String = ""
            If RERUN = True Then
                If rs.options.Rs.Fields("RUNNING").Value.ToString.Trim <> "" Then
                    GoTo Gonext
                Else
                    If PreLocation <> rs.options.Rs.Fields("LOCID").Value Then
                        If PreLocation <> "" Then
                            collvat.Remove(PreLocation)
                            collvat.Add(loRunning, PreLocation)
                        End If
                        If Location.Count > 0 Then
                            loRunning = CType(Location(rs.options.Rs.Fields("LOCID").Value & SQLNewDocNO("", rs.options.Rs.Fields("TXDATE").Value.ToString, 2)), String())(2)
                        Else
                            collvat.Add(0, rs.options.Rs.Fields("LOCID").Value)
                            loRunning = 0
                        End If
                        PreLocation = rs.options.Rs.Fields("LOCID").Value
                    ElseIf TMPTXDATE <> SQLNewDocNO("", rs.options.Rs.Fields("TXDATE").Value.ToString, 2) Then
                        If Location.Count > 0 Then
                            CType(Location(rs.options.Rs.Fields("LOCID").Value & SQLNewDocNO("", rs.options.Rs.Fields("TXDATE").Value.ToString, 2)), String())(2) = loRunning
                            loRunning = 0
                        End If
                    End If
                End If
            Else
                If PreLocation <> rs.options.Rs.Fields("LOCID").Value Then
                    If PreLocation <> "" Then
                        collvat.Remove(PreLocation)
                        collvat.Add(loRunning, PreLocation)
                    End If

                    If collvat.Contains(rs.options.Rs.Fields("LOCID").Value) Then
                        loRunning = collvat(rs.options.Rs.Fields("LOCID").Value).ToString
                    Else
                        collvat.Add(0, rs.options.Rs.Fields("LOCID").Value)
                        loRunning = 0
                    End If
                    PreLocation = rs.options.Rs.Fields("LOCID").Value
                End If
            End If
            loPrefix = IIf(IsDBNull(rs.options.Rs.Fields("LOCPREFIX").Value), "", rs.options.Rs.Fields("LOCPREFIX").Value)
            loPrefix = loPrefix.Trim
            loRunning = loRunning + 1
            TMPTXDATE = SQLNewDocNO("", rs.options.Rs.Fields("TXDATE").Value.ToString, 2)
            If RERUN = True Then
                RunNO = SQLNewDocNO(loPrefix, rs.options.Rs.Fields("TXDATE").Value.ToString, 3) & Format(loRunning, VATFormat)
                abc = ("UPDATE FMSVAT set NEWDOCNO ='" & RunNO & "' WHERE VATID=" & rs.options.Rs.Fields("VATID").Value)
                cn.Execute(abc)
                abc = ("UPDATE FMSVATINSERT set NEWDOCNO ='" & RunNO & "'  WHERE BATCH =" & rs.options.Rs.Fields("BATCH").Value & " and ENTRY =" & rs.options.Rs.Fields("ENTRY").Value & " and TRANSNBR =" & rs.options.Rs.Fields("LINE").Value)
                cn.Execute(abc)
                If RERUN Then
                    abc = ("if not exists (select * from FMSRUN where BATCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and ENTRY ='" & rs.options.Rs.Fields("ENTRY").Value & "' and TRANSNBR ='" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "' and SOURCE ='" & rs.options.Rs.Fields("SOURCE").Value & "') begin INSERT INTO FMSRUN VALUES('" & rs.options.Rs.Fields("BATCH").Value & "','" & rs.options.Rs.Fields("ENTRY").Value & "','" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "','" & Format(loRunning, VATFormat) & "','" & rs.options.Rs.Fields("SOURCE").Value & "') end")
                    cn.Execute(abc)
                End If
            ElseIf DATEMODE = "DOCU" Then
                RunNO = SQLNewDocNO(loPrefix, rs.options.Rs.Fields("TXDATE").Value.ToString, 3) & Format(loRunning, VATFormat)
                abc = ("UPDATE FMSVAT set NEWDOCNO ='" & RunNO & "' WHERE VATID=" & rs.options.Rs.Fields("VATID").Value)
                cn.Execute(abc)
                abc = ("UPDATE FMSVATINSERT set NEWDOCNO ='" & RunNO & "'  WHERE BATCH =" & rs.options.Rs.Fields("BATCH").Value & " and ENTRY =" & rs.options.Rs.Fields("ENTRY").Value & " and TRANSNBR =" & rs.options.Rs.Fields("LINE").Value)
                cn.Execute(abc)
                If RERUN Then
                    abc = ("if not exists (select * from FMSRUN where BATCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and ENTRY ='" & rs.options.Rs.Fields("ENTRY").Value & "' and TRANSNBR ='" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "' and SOURCE ='" & rs.options.Rs.Fields("SOURCE").Value & "') begin INSERT INTO FMSRUN VALUES('" & rs.options.Rs.Fields("BATCH").Value & "','" & rs.options.Rs.Fields("ENTRY").Value & "','" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "','" & Format(loRunning, VATFormat) & "','" & rs.options.Rs.Fields("SOURCE").Value & "') end")
                    cn.Execute(abc)
                End If
            Else
                RunNO = SQLNewDocNO(loPrefix, rs.options.Rs.Fields("TXDATE").Value.ToString, 3) & Format(loRunning, VATFormat)
                abc = ("UPDATE FMSVAT set NEWDOCNO='" & RunNO & "' WHERE VATID=" & rs.options.Rs.Fields("VATID").Value)
                cn.Execute(abc)
                abc = ("UPDATE FMSVATINSERT set NEWDOCNO='" & RunNO & "'  WHERE BATCH =" & rs.options.Rs.Fields("BATCH").Value & " and ENTRY =" & rs.options.Rs.Fields("ENTRY").Value & " and TRANSNBR =" & rs.options.Rs.Fields("LINE").Value)
                cn.Execute(abc)
                If RERUN Then
                    abc = ("if not exists (select * from FMSRUN where BATCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and ENTRY ='" & rs.options.Rs.Fields("ENTRY").Value & "' and TRANSNBR ='" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "' and SOURCE ='" & rs.options.Rs.Fields("SOURCE").Value & "') begin INSERT INTO FMSRUN VALUES('" & rs.options.Rs.Fields("BATCH").Value & "','" & rs.options.Rs.Fields("ENTRY").Value & "','" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "','" & Format(loRunning, VATFormat) & "','" & rs.options.Rs.Fields("SOURCE").Value & "') end")
                    cn.Execute(abc)
                End If
            End If

Gonext:

            If UPRUN = True Then
                If RunNO = "" Then
                    RunNO = rs.options.Rs.Fields("NEWDOCNO").Value.ToString.Trim
                End If
                If rs.options.Rs.Fields("SOURCE").Value.ToString.Trim = "AP" Then
                    If IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) = "0" Then
                        sqlstr = "Update APIBHO SET VALUE='" & RunNO & "' WHERE OPTFIELD='VATVOUCHER' and CNTBTCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and CNTITEM ='" & rs.options.Rs.Fields("ENTRY").Value & "'"
                    Else
                        sqlstr = "Update APIBDO SET VALUE='" & RunNO & "' WHERE OPTFIELD='VATVOUCHER' and CNTBTCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and CNTITEM ='" & rs.options.Rs.Fields("ENTRY").Value & "' and CNTLINE ='" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value.ToString.Trim) & "'"
                    End If
                    cnACCPAC.Execute(sqlstr)
                End If
                If rs.options.Rs.Fields("SOURCE").Value.ToString.Trim = "PO" Then
                    sqlstr = "Update APIBDO SET VALUE='" & RunNO & "' WHERE OPTFIELD='VATVOUCHER' and CNTBTCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and CNTITEM ='" & rs.options.Rs.Fields("ENTRY").Value & "' and CNTLINE ='" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value.ToString.Trim) & "'"
                    cnACCPAC.Execute(sqlstr)
                End If
                If rs.options.Rs.Fields("SOURCE").Value.ToString.Trim = "CB" Then
                    sqlstr = "Update CBBTDTO  set VALUE ='" & RunNO & "' from CBBTDTO where OPTFIELD='VATVOUCHER' and BATCHID =" & rs.options.Rs.Fields("BATCH").Value & " and ENTRYNO =" & rs.options.Rs.Fields("ENTRY").Value & " "
                    cnACCPAC.Execute(sqlstr)
                End If
                If rs.options.Rs.Fields("SOURCE").Value.ToString.Trim = "GL" Then
                    sqlstr = "Update GLJEDO SET VALUE='" & RunNO & "' WHERE OPTFIELD='VATVOUCHER' and BATCHNBR =" & rs.options.Rs.Fields("BATCH").Value & " and JOURNALID =" & rs.options.Rs.Fields("ENTRY").Value & " and TRANSNBR =" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value.ToString.Trim) & ""
                    cnACCPAC.Execute(sqlstr)
                End If

            End If
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        cn.Close()
        cnACCPAC.Close()
        rs = Nothing
        cn = Nothing
        Exit Sub

ErrHandler:
        ErrMsg = Err.Description
        WriteLog("<ProcessNewDocNo>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessNewDocNo> sqlstr : " & abc & Environment.NewLine & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub


    'สำหรับ STD
    Public Sub ProcessNewDocNo()
        On Error GoTo ErrHandler
        Dim cn, cnACCPAC As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr, abc As String
        Dim loPrefix As String
        Dim loRunning As Integer = 0
        Dim PreLocation As String
        cn = New ADODB.Connection
        cnACCPAC = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        cn.ConnectionTimeout = 60
        cn.Open(ConnVAT)
        cn.CommandTimeout = 3600

        cnACCPAC.ConnectionTimeout = 60
        cnACCPAC.Open(ConnACCPAC)
        cnACCPAC.CommandTimeout = 3600
        PreLocation = ""
        Dim IsSource As String = ""
        Dim Location As New Collection

        If RERUN Then
            sqlstr = "UPDATE " & Environment.NewLine
            sqlstr &= "FMSVAT" & Environment.NewLine
            sqlstr &= "SET" & Environment.NewLine
            sqlstr &= "FMSVAT.NEWDOCNO  =FMSVLOC.LOCPREFIX+right(left(FMSVAT.TXDATE,6),2)+'/'+FMSRUN.RUNNING " & Environment.NewLine
            sqlstr &= "FROM" & Environment.NewLine
            sqlstr &= "FMSVAT LEFT JOIN FMSVLOC ON FMSVAT.LOCID = FMSVLOC.LOCID" & Environment.NewLine
            sqlstr &= "LEFT JOIN FMSRUN on (FMSVAT.BATCH =FMSRUN.BATCH) and (FMSVAT.ENTRY =FMSRUN.ENTRY  ) and ((CASE WHEN isnull(FMSVAT.TRANSNBR,0)='' then 0 else isnull(FMSVAT.TRANSNBR,0) end) =(CASE WHEN isnull(FMSRUN.TRANSNBR,0)='' then 0 else isnull(FMSRUN.TRANSNBR,0) end))and (FMSVAT.SOURCE  =FMSRUN.SOURCE  )" & Environment.NewLine
            sqlstr &= "where ISNULL (fmsrun.running,'')<>'';" & Environment.NewLine

            sqlstr &= "UPDATE " & Environment.NewLine
            sqlstr &= "FMSVATINSERT" & Environment.NewLine
            sqlstr &= "SET" & Environment.NewLine
            sqlstr &= "FMSVATINSERT.NEWDOCNO  = FMSVLOC.LOCPREFIX+right(left(FMSVATINSERT.TXDATE,6),2)+'/'+FMSRUN.RUNNING" & Environment.NewLine
            sqlstr &= "FROM" & Environment.NewLine
            sqlstr &= "FMSVATINSERT LEFT JOIN FMSVLOC ON FMSVATINSERT.LOCID = FMSVLOC.LOCID" & Environment.NewLine
            sqlstr &= "LEFT JOIN FMSRUN on (FMSVATINSERT.BATCH =FMSRUN.BATCH) and (FMSVATINSERT.ENTRY =FMSRUN.ENTRY  ) and ((CASE WHEN isnull(FMSVATINSERT.TRANSNBR,0)='' then 0 else isnull(FMSVATINSERT.TRANSNBR,0) end) =(CASE WHEN isnull(FMSRUN.TRANSNBR,0)='' then 0 else isnull(FMSRUN.TRANSNBR,0) end))and (FMSVATINSERT.SOURCE  =FMSRUN.SOURCE  )" & Environment.NewLine
            sqlstr &= "where ISNULL (fmsrun.running,'')<>'';" & Environment.NewLine
            cn.Execute(sqlstr)

            Dim RS_Location As New BaseClass.BaseODBCIO("select  isnull(max(RUNNING),0)as RUNNING ,LOCID ,right(left(TXDATE,6),2) as TXDATE from FMSVAT  LEFT JOIN FMSRUN on (FMSVAT.BATCH =FMSRUN.BATCH) and (FMSVAT.ENTRY =FMSRUN.ENTRY  ) and ((CASE WHEN isnull(FMSVAT.TRANSNBR,0)='' then 0 else isnull(FMSVAT.TRANSNBR,0) end) =(CASE WHEN isnull(FMSRUN.TRANSNBR,0)='' then 0 else isnull(FMSRUN.TRANSNBR,0) end)) and (FMSVAT.SOURCE  =FMSRUN.SOURCE  ) group by LOCID,right(left(TXDATE,6),2)", cn)
            For iLoca As Integer = 0 To RS_Location.options.QueryDT.Rows.Count - 1
                Location.Add(New String() {RS_Location.options.QueryDT.Rows(iLoca).Item("LOCID"), RS_Location.options.QueryDT.Rows(iLoca).Item("TXDATE"), RS_Location.options.QueryDT.Rows(iLoca).Item("RUNNING")}, RS_Location.options.QueryDT.Rows(iLoca).Item("LOCID") + RS_Location.options.QueryDT.Rows(iLoca).Item("TXDATE"))
            Next
        End If

        sqlstr = "SELECT FMSVAT.TXDATE,FMSVAT.DOCNO,FMSVAT.VATID,FMSVAT.TXDATE,FMSVLOC.LOCPREFIX,FMSVLOC.LOCID,FMSVAT.BATCH ,FMSVAT.ENTRY ,(CASE WHEN isnull(FMSVAT.TRANSNBR,0)='' then 0 else isnull(FMSVAT.TRANSNBR,0) end)as LINE ,FMSVAT.SOURCE" & IIf(RERUN, " ,isnull(FMSRUN.RUNNING,'') as RUNNING", "") & ",FMSVAT.NEWDOCNO " & Environment.NewLine
        sqlstr &= ",(CASE	WHEN ltrim(rtrim(FMSVAT.SOURCE))in('AP','PO') then (select top 1 APOBL.AUDTDATE  from " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & " APOBL left join " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "APIBD on (APOBL.CNTBTCH =APIBD.CNTBTCH and APOBL.CNTITEM =APIBD.CNTITEM ) where APOBL.CNTBTCH =FMSVAT.BATCH  and  APOBL.CNTITEM =FMSVAT.ENTRY) " & Environment.NewLine
        sqlstr &= "		    WHEN ltrim(rtrim(FMSVAT.SOURCE))='GL' then (select top 1 GLPOST.AUDTDATE  from " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "GLPOST left join " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "GLJED  on (GLPOST.BATCHNBR =GLJED.BATCHNBR and GLPOST.ENTRYNBR =GLJED.JOURNALID ) where GLPOST.BATCHNBR=convert(char(6),FMSVAT.BATCH)  and  GLPOST.ENTRYNBR =convert(char(6),FMSVAT.ENTRY)) " & Environment.NewLine
        If FindApp("CB") = True Then
            sqlstr &= "		    WHEN ltrim(rtrim(FMSVAT.SOURCE))='CB' then (select top 1 CBTRHD.AUDTDATE  from " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "CBTRHD left join " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "CBBTDT   on (CBTRHD.BATCHID  =CBBTDT.BATCHID and CBTRHD.ENTRYNO =CBBTDT.ENTRYNO  ) where CBTRHD.BATCHID=FMSVAT.BATCH  and  CBTRHD.ENTRYNO =FMSVAT.ENTRY) " & Environment.NewLine
        End If
        sqlstr &= "		    WHEN ltrim(rtrim(FMSVAT.SOURCE))in('AR','OE') then (select top 1 AROBL.AUDTDATE  from " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "AROBL  left join " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "ARIBD    on (AROBL.CNTBTCH  =ARIBD.CNTBTCH and AROBL.CNTITEM =ARIBD.CNTITEM  ) where AROBL.CNTBTCH=FMSVAT.BATCH  and  ARIBD.CNTITEM =FMSVAT.ENTRY) " & Environment.NewLine
        sqlstr &= "end) as AUDTDATE" & Environment.NewLine
        If ComDatabase = "ORCL" Then
            sqlstr = sqlstr & " FROM FMSVAT LEFT JOIN FMSVLOC ON FMSVAT.LOCID = FMSVLOC.LOCID" & Environment.NewLine
        Else
            sqlstr = sqlstr & " FROM FMSVAT LEFT JOIN FMSVLOC ON FMSVAT.LOCID = FMSVLOC.LOCID" & Environment.NewLine
        End If
        If RERUN Then sqlstr &= " LEFT JOIN FMSRUN on (FMSVAT.BATCH =FMSRUN.BATCH) and (FMSVAT.ENTRY =FMSRUN.ENTRY  ) and ((CASE WHEN isnull(FMSVAT.TRANSNBR,0)='' then 0 else isnull(FMSVAT.TRANSNBR,0) end) =(CASE WHEN isnull(FMSRUN.TRANSNBR,0)='' then 0 else isnull(FMSRUN.TRANSNBR,0)end))" & Environment.NewLine
        If CheckTTypeRunning = 1 Then
            sqlstr = sqlstr & " WHERE TTYPE = 1" & Environment.NewLine
        ElseIf CheckTTypeRunning = 2 Then
            sqlstr = sqlstr & " WHERE TTYPE = 2" & Environment.NewLine
        ElseIf CheckTTypeRunning = 3 Then
            sqlstr = sqlstr & " WHERE TTYPE = 3" & Environment.NewLine
        End If
        If CheckRate.Trim.Split("||").Length = 3 Then
            sqlstr &= " and RATE = " & CheckRate.Trim.Split("||")(2) & Environment.NewLine
        ElseIf CheckRate.Trim = "None" Then
            sqlstr &= " and (VTYPE = 0)" & Environment.NewLine
        ElseIf CheckRate.Trim = "unlike 7" Then
            sqlstr &= " and (RATE <>7 and RATE <>0)" & Environment.NewLine
        End If
        sqlstr &= " ORDER BY AUDTDATE,BATCH,ENTRY ,LINE"
        rs.Open(sqlstr, cn)
        Dim collvat As New Collection
        Dim TMPTXDATE As String = ""
        Do While rs.options.Rs.EOF = False
            Dim RunNO As String = ""
            If RERUN = True Then
                If rs.options.Rs.Fields("RUNNING").Value.ToString.Trim <> "" Then
                    GoTo Gonext
                Else
                    If PreLocation <> rs.options.Rs.Fields("LOCID").Value Then
                        If PreLocation <> "" Then
                            collvat.Remove(PreLocation)
                            collvat.Add(loRunning, PreLocation)
                        End If
                        If Location.Count > 0 Then
                            loRunning = CType(Location(rs.options.Rs.Fields("LOCID").Value & rs.options.Rs.Fields("TXDATE").Value.ToString.Substring(4, 2)), String())(2)
                        Else
                            collvat.Add(0, rs.options.Rs.Fields("LOCID").Value)
                            loRunning = 0
                        End If
                        PreLocation = rs.options.Rs.Fields("LOCID").Value
                    ElseIf TMPTXDATE <> Mid(rs.options.Rs.Fields("TXDATE").Value, 5, 2) Then
                        If Location.Count > 0 Then
                            CType(Location(rs.options.Rs.Fields("LOCID").Value & rs.options.Rs.Fields("TXDATE").Value.ToString.Substring(4, 2)), String())(2) = loRunning
                            loRunning = 0
                        End If
                    End If
                End If
            Else
                If PreLocation <> rs.options.Rs.Fields("LOCID").Value Then
                    If PreLocation <> "" Then
                        collvat.Remove(PreLocation)
                        collvat.Add(loRunning, PreLocation)
                    End If

                    If collvat.Contains(rs.options.Rs.Fields("LOCID").Value) Then
                        loRunning = collvat(rs.options.Rs.Fields("LOCID").Value).ToString
                    Else
                        collvat.Add(0, rs.options.Rs.Fields("LOCID").Value)
                        loRunning = 0
                    End If
                    PreLocation = rs.options.Rs.Fields("LOCID").Value
                End If
            End If
            loPrefix = IIf(IsDBNull(rs.options.Rs.Fields("LOCPREFIX").Value), "", Trim(rs.options.Rs.Fields("LOCPREFIX").Value))
            loRunning = loRunning + 1
            TMPTXDATE = Mid(rs.options.Rs.Fields("TXDATE").Value, 5, 2)

            If RERUN = True Then
                RunNO = loPrefix & Mid(rs.options.Rs.Fields("TXDATE").Value, 5, 2) & "/" & Format(loRunning, VATFormat)
                abc = ("UPDATE FMSVAT set NEWDOCNO='" & RunNO & "' WHERE VATID=" & rs.options.Rs.Fields("VATID").Value)
                cn.Execute(abc)
                abc = ("UPDATE FMSVATINSERT set NEWDOCNO='" & RunNO & "'  WHERE VATID=" & rs.options.Rs.Fields("VATID").Value)
                cn.Execute(abc)
                If RERUN Then
                    abc = ("if not exists (select * from FMSRUN where BATCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and ENTRY ='" & rs.options.Rs.Fields("ENTRY").Value & "' and TRANSNBR ='" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "' and SOURCE ='" & rs.options.Rs.Fields("SOURCE").Value & "') begin INSERT INTO FMSRUN VALUES('" & rs.options.Rs.Fields("BATCH").Value & "','" & rs.options.Rs.Fields("ENTRY").Value & "','" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "','" & Format(loRunning, VATFormat) & "','" & rs.options.Rs.Fields("SOURCE").Value & "') end")
                    cn.Execute(abc)
                End If
            ElseIf DATEMODE = "DOCU" Then
                RunNO = loPrefix & Mid(rs.options.Rs.Fields("TXDATE").Value, 5, 2) & "/" & Format(loRunning, VATFormat)
                abc = ("UPDATE FMSVAT set NEWDOCNO='" & RunNO & "' WHERE VATID=" & rs.options.Rs.Fields("VATID").Value)
                cn.Execute(abc)
                abc = ("UPDATE FMSVATINSERT set NEWDOCNO='" & RunNO & "'  WHERE VATID=" & rs.options.Rs.Fields("VATID").Value)
                cn.Execute(abc)
                If RERUN Then
                    abc = ("if not exists (select * from FMSRUN where BATCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and ENTRY ='" & rs.options.Rs.Fields("ENTRY").Value & "' and TRANSNBR ='" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "' and SOURCE ='" & rs.options.Rs.Fields("SOURCE").Value & "') begin INSERT INTO FMSRUN VALUES('" & rs.options.Rs.Fields("BATCH").Value & "','" & rs.options.Rs.Fields("ENTRY").Value & "','" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "','" & Format(loRunning, VATFormat) & "','" & rs.options.Rs.Fields("SOURCE").Value & "') end")
                    cn.Execute(abc)
                End If
            Else
                RunNO = loPrefix & Mid(rs.options.Rs.Fields("TXDATE").Value, 5, 2) & "/" & Format(loRunning, VATFormat)
                abc = ("UPDATE FMSVAT set NEWDOCNO='" & RunNO & "' WHERE VATID=" & rs.options.Rs.Fields("VATID").Value)
                cn.Execute(abc)
                abc = ("UPDATE FMSVATINSERT set NEWDOCNO='" & RunNO & "'  WHERE VATID=" & rs.options.Rs.Fields("VATID").Value)
                cn.Execute(abc)
                If RERUN Then
                    abc = ("if not exists (select * from FMSRUN where BATCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and ENTRY ='" & rs.options.Rs.Fields("ENTRY").Value & "' and TRANSNBR ='" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "' and SOURCE ='" & rs.options.Rs.Fields("SOURCE").Value & "') begin INSERT INTO FMSRUN VALUES('" & rs.options.Rs.Fields("BATCH").Value & "','" & rs.options.Rs.Fields("ENTRY").Value & "','" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "','" & Format(loRunning, VATFormat) & "','" & rs.options.Rs.Fields("SOURCE").Value & "') end")
                    cn.Execute(abc)
                End If
            End If

Gonext:

            If UPRUN = True Then
                If RunNO = "" Then
                    RunNO = rs.options.Rs.Fields("NEWDOCNO").Value.ToString.Trim
                End If
                If rs.options.Rs.Fields("SOURCE").Value.ToString.Trim = "AP" Then
                    If IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) = "0" Then
                        sqlstr = "Update APIBHO SET VALUE='" & RunNO & "' WHERE OPTFIELD='VATVOUCHER' and CNTBTCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and CNTITEM ='" & rs.options.Rs.Fields("ENTRY").Value & "'"
                    Else
                        sqlstr = "Update APIBDO SET VALUE='" & RunNO & "' WHERE OPTFIELD='VATVOUCHER' and CNTBTCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and CNTITEM ='" & rs.options.Rs.Fields("ENTRY").Value & "' and CNTLINE ='" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value.ToString.Trim) & "'"
                    End If
                    cnACCPAC.Execute(sqlstr)
                End If
                If rs.options.Rs.Fields("SOURCE").Value.ToString.Trim = "PO" Then
                    sqlstr = "Update APIBDO SET VALUE='" & RunNO & "' WHERE OPTFIELD='VATVOUCHER' and CNTBTCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and CNTITEM ='" & rs.options.Rs.Fields("ENTRY").Value & "' and CNTLINE ='" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value.ToString.Trim) & "'"
                    cnACCPAC.Execute(sqlstr)
                End If
                If rs.options.Rs.Fields("SOURCE").Value.ToString.Trim = "CB" Then
                    sqlstr = "update CBBTDTO  set VALUE ='" & RunNO & "' from CBBTDTO where OPTFIELD='VATVOUCHER' and BATCHID =" & rs.options.Rs.Fields("BATCH").Value & " and ENTRYNO =" & rs.options.Rs.Fields("ENTRY").Value & " "
                    cnACCPAC.Execute(sqlstr)
                End If
                If rs.options.Rs.Fields("SOURCE").Value.ToString.Trim = "GL" Then
                    sqlstr = "Update GLJEDO SET VALUE='" & RunNO & "' WHERE OPTFIELD='VATVOUCHER' and BATCHNBR =" & rs.options.Rs.Fields("BATCH").Value & " and JOURNALID =" & rs.options.Rs.Fields("ENTRY").Value & " and TRANSNBR =" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value.ToString.Trim) & ""
                    cnACCPAC.Execute(sqlstr)
                End If

            End If
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        cn.Close()
        cnACCPAC.Close()
        rs = Nothing
        cn = Nothing
        Exit Sub

ErrHandler:
        ErrMsg = Err.Description
        WriteLog("<ProcessNewDocNo>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessNewDocNo> sqlstr : " & abc & Environment.NewLine & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    'SNC
    Public Sub ProcessNewDocNoSNC()
        On Error GoTo ErrHandler
        Dim cn, cnACCPAC As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr, abc As String
        Dim loPrefix As String
        Dim loRunning As Integer = 0
        Dim PreLocation As String
        cn = New ADODB.Connection
        cnACCPAC = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        cn.ConnectionTimeout = 60
        cn.Open(ConnVAT)
        cn.CommandTimeout = 3600

        cnACCPAC.ConnectionTimeout = 60
        cnACCPAC.Open(ConnACCPAC)
        cnACCPAC.CommandTimeout = 3600
        PreLocation = ""
        Dim IsSource As String = ""
        Dim Location As New Collection

        If RERUN Then
            sqlstr = "UPDATE " & Environment.NewLine
            sqlstr &= "FMSVAT" & Environment.NewLine
            sqlstr &= "SET" & Environment.NewLine
            sqlstr &= "FMSVAT.NEWDOCNO  =FMSVLOC.LOCPREFIX+RIGHT(left(FMSVAT.TXDATE,4),2)+right(left(FMSVAT.TXDATE,6),2)+FMSRUN.RUNNING " & Environment.NewLine
            sqlstr &= "FROM" & Environment.NewLine
            sqlstr &= "FMSVAT LEFT JOIN FMSVLOC ON FMSVAT.LOCID = FMSVLOC.LOCID" & Environment.NewLine
            sqlstr &= "LEFT JOIN FMSRUN on (FMSVAT.BATCH =FMSRUN.BATCH) and (FMSVAT.ENTRY =FMSRUN.ENTRY  ) and ((CASE WHEN isnull(FMSVAT.TRANSNBR,0)='' then 0 else isnull(FMSVAT.TRANSNBR,0) end) =(CASE WHEN isnull(FMSRUN.TRANSNBR,0)='' then 0 else isnull(FMSRUN.TRANSNBR,0) end))and (FMSVAT.SOURCE  =FMSRUN.SOURCE  )" & Environment.NewLine
            sqlstr &= "where ISNULL (fmsrun.running,'')<>'';" & Environment.NewLine

            sqlstr &= "UPDATE " & Environment.NewLine
            sqlstr &= "FMSVATINSERT" & Environment.NewLine
            sqlstr &= "SET" & Environment.NewLine
            sqlstr &= "FMSVATINSERT.NEWDOCNO  = FMSVLOC.LOCPREFIX+RIGHT(left(FMSVAT.TXDATE,4),2)+right(left(FMSVATINSERT.TXDATE,6),2)+FMSRUN.RUNNING" & Environment.NewLine
            sqlstr &= "FROM" & Environment.NewLine
            sqlstr &= "FMSVATINSERT LEFT JOIN FMSVLOC ON FMSVATINSERT.LOCID = FMSVLOC.LOCID" & Environment.NewLine
            sqlstr &= "LEFT JOIN FMSRUN on (FMSVATINSERT.BATCH =FMSRUN.BATCH) and (FMSVATINSERT.ENTRY =FMSRUN.ENTRY  ) and ((CASE WHEN isnull(FMSVATINSERT.TRANSNBR,0)='' then 0 else isnull(FMSVATINSERT.TRANSNBR,0) end) =(CASE WHEN isnull(FMSRUN.TRANSNBR,0)='' then 0 else isnull(FMSRUN.TRANSNBR,0) end))and (FMSVATINSERT.SOURCE  =FMSRUN.SOURCE  )" & Environment.NewLine
            sqlstr &= "where ISNULL (fmsrun.running,'')<>'';" & Environment.NewLine
            cn.Execute(sqlstr)

            Dim RS_Location As New BaseClass.BaseODBCIO("select  isnull(max(RUNNING),0)as RUNNING ,LOCID ,RIGHT(left(FMSVAT.TXDATE,4),2)+right(left(TXDATE,6),2) as TXDATE from FMSVAT  LEFT JOIN FMSRUN on (FMSVAT.BATCH =FMSRUN.BATCH) and (FMSVAT.ENTRY =FMSRUN.ENTRY  ) and ((CASE WHEN isnull(FMSVAT.TRANSNBR,0)='' then 0 else isnull(FMSVAT.TRANSNBR,0) end) =(CASE WHEN isnull(FMSRUN.TRANSNBR,0)='' then 0 else isnull(FMSRUN.TRANSNBR,0) end)) and (FMSVAT.SOURCE  =FMSRUN.SOURCE  ) group by LOCID,RIGHT(left(FMSVAT.TXDATE,4),2)+right(left(TXDATE,6),2)", cn)
            For iLoca As Integer = 0 To RS_Location.options.QueryDT.Rows.Count - 1
                Location.Add(New String() {RS_Location.options.QueryDT.Rows(iLoca).Item("LOCID"), RS_Location.options.QueryDT.Rows(iLoca).Item("TXDATE"), RS_Location.options.QueryDT.Rows(iLoca).Item("RUNNING")}, RS_Location.options.QueryDT.Rows(iLoca).Item("LOCID") + RS_Location.options.QueryDT.Rows(iLoca).Item("TXDATE"))
            Next
        End If

        sqlstr = "SELECT FMSVAT.TXDATE,FMSVAT.DOCNO,FMSVAT.VATID,FMSVAT.TXDATE,FMSVLOC.LOCPREFIX,FMSVLOC.LOCID,FMSVAT.BATCH ,FMSVAT.ENTRY ,(CASE WHEN isnull(FMSVAT.TRANSNBR,0)='' then 0 else isnull(FMSVAT.TRANSNBR,0) end)as LINE ,FMSVAT.SOURCE" & IIf(RERUN, " ,isnull(FMSRUN.RUNNING,'') as RUNNING", "") & ",FMSVAT.NEWDOCNO " & Environment.NewLine
        sqlstr &= ",(CASE	WHEN ltrim(rtrim(FMSVAT.SOURCE))in('AP','PO') then (select top 1 APOBL.AUDTDATE  from " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & " APOBL left join " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "APIBD on (APOBL.CNTBTCH =APIBD.CNTBTCH and APOBL.CNTITEM =APIBD.CNTITEM ) where APOBL.CNTBTCH =FMSVAT.BATCH  and  APOBL.CNTITEM =FMSVAT.ENTRY) " & Environment.NewLine
        sqlstr &= "		    WHEN ltrim(rtrim(FMSVAT.SOURCE))='GL' then (select top 1 GLPOST.AUDTDATE  from " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "GLPOST left join " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "GLJED  on (GLPOST.BATCHNBR =GLJED.BATCHNBR and GLPOST.ENTRYNBR =GLJED.JOURNALID ) where GLPOST.BATCHNBR=convert(char(6),FMSVAT.BATCH)  and  GLPOST.ENTRYNBR =convert(char(6),FMSVAT.ENTRY)) " & Environment.NewLine
        If FindApp("CB") = True Then
            sqlstr &= "		    WHEN ltrim(rtrim(FMSVAT.SOURCE))='CB' then (select top 1 CBTRHD.AUDTDATE  from " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "CBTRHD left join " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "CBBTDT   on (CBTRHD.BATCHID  =CBBTDT.BATCHID and CBTRHD.ENTRYNO =CBBTDT.ENTRYNO  ) where CBTRHD.BATCHID=FMSVAT.BATCH  and  CBTRHD.ENTRYNO =FMSVAT.ENTRY) " & Environment.NewLine
        End If
        sqlstr &= "		    WHEN ltrim(rtrim(FMSVAT.SOURCE))in('AR','OE') then (select top 1 AROBL.AUDTDATE  from " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "AROBL  left join " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "ARIBD    on (AROBL.CNTBTCH  =ARIBD.CNTBTCH and AROBL.CNTITEM =ARIBD.CNTITEM  ) where AROBL.CNTBTCH=FMSVAT.BATCH  and  ARIBD.CNTITEM =FMSVAT.ENTRY) " & Environment.NewLine
        sqlstr &= "end) as AUDTDATE" & Environment.NewLine
        If ComDatabase = "ORCL" Then
            sqlstr = sqlstr & " FROM FMSVAT LEFT JOIN FMSVLOC ON FMSVAT.LOCID = FMSVLOC.LOCID" & Environment.NewLine
        Else
            sqlstr = sqlstr & " FROM FMSVAT LEFT JOIN FMSVLOC ON FMSVAT.LOCID = FMSVLOC.LOCID" & Environment.NewLine
        End If
        If RERUN Then sqlstr &= " LEFT JOIN FMSRUN on (FMSVAT.BATCH =FMSRUN.BATCH) and (FMSVAT.ENTRY =FMSRUN.ENTRY  ) and ((CASE WHEN isnull(FMSVAT.TRANSNBR,0)='' then 0 else isnull(FMSVAT.TRANSNBR,0) end) =(CASE WHEN isnull(FMSRUN.TRANSNBR,0)='' then 0 else isnull(FMSRUN.TRANSNBR,0)end))" & Environment.NewLine
        If CheckTTypeRunning = 1 Then
            sqlstr = sqlstr & " WHERE TTYPE = 1" & Environment.NewLine
        ElseIf CheckTTypeRunning = 2 Then
            sqlstr = sqlstr & " WHERE TTYPE = 2" & Environment.NewLine
        ElseIf CheckTTypeRunning = 3 Then
            sqlstr = sqlstr & " WHERE TTYPE = 3" & Environment.NewLine
        End If
        If CheckRate.Trim.Split("||").Length = 3 Then
            sqlstr &= " and RATE = " & CheckRate.Trim.Split("||")(2) & Environment.NewLine
        ElseIf CheckRate.Trim = "None" Then
            sqlstr &= " and (VTYPE = 0)" & Environment.NewLine
        ElseIf CheckRate.Trim = "unlike 7" Then
            sqlstr &= " and (RATE <>7 and RATE <>0)" & Environment.NewLine
        End If
        sqlstr &= " ORDER BY AUDTDATE,BATCH,ENTRY ,LINE"
        rs.Open(sqlstr, cn)
        Dim collvat As New Collection
        Dim TMPTXDATE As String = ""
        Do While rs.options.Rs.EOF = False
            Dim RunNO As String = ""
            If RERUN = True Then
                If rs.options.Rs.Fields("RUNNING").Value.ToString.Trim <> "" Then
                    GoTo Gonext
                Else
                    If PreLocation <> rs.options.Rs.Fields("LOCID").Value Then
                        If PreLocation <> "" Then
                            collvat.Remove(PreLocation)
                            collvat.Add(loRunning, PreLocation)
                        End If
                        If Location.Count > 0 Then
                            loRunning = CType(Location(rs.options.Rs.Fields("LOCID").Value & rs.options.Rs.Fields("TXDATE").Value.ToString.Substring(2, 2) & rs.options.Rs.Fields("TXDATE").Value.ToString.Substring(4, 2)), String())(2)
                        Else
                            collvat.Add(0, rs.options.Rs.Fields("LOCID").Value)
                            loRunning = 0
                        End If
                        PreLocation = rs.options.Rs.Fields("LOCID").Value
                    ElseIf TMPTXDATE <> Mid(rs.options.Rs.Fields("TXDATE").Value, 3, 2) & Mid(rs.options.Rs.Fields("TXDATE").Value, 5, 2) Then
                        If Location.Count > 0 Then
                            CType(Location(rs.options.Rs.Fields("LOCID").Value & rs.options.Rs.Fields("TXDATE").Value.ToString.Substring(2, 2) & rs.options.Rs.Fields("TXDATE").Value.ToString.Substring(4, 2)), String())(2) = loRunning
                            loRunning = 0
                        End If
                    End If
                End If
            Else
                If PreLocation <> rs.options.Rs.Fields("LOCID").Value Then
                    If PreLocation <> "" Then
                        collvat.Remove(PreLocation)
                        collvat.Add(loRunning, PreLocation)
                    End If

                    If collvat.Contains(rs.options.Rs.Fields("LOCID").Value) Then
                        loRunning = collvat(rs.options.Rs.Fields("LOCID").Value).ToString
                    Else
                        collvat.Add(0, rs.options.Rs.Fields("LOCID").Value)
                        loRunning = 0
                    End If
                    PreLocation = rs.options.Rs.Fields("LOCID").Value
                End If
            End If
            loPrefix = IIf(IsDBNull(rs.options.Rs.Fields("LOCPREFIX").Value), "", Trim(rs.options.Rs.Fields("LOCPREFIX").Value))
            loRunning = loRunning + 1
            TMPTXDATE = Mid(rs.options.Rs.Fields("TXDATE").Value, 3, 2) & Mid(rs.options.Rs.Fields("TXDATE").Value, 5, 2)

            If RERUN = True Then
                RunNO = loPrefix & Mid(rs.options.Rs.Fields("TXDATE").Value, 3, 2) & Mid(rs.options.Rs.Fields("TXDATE").Value, 5, 2) & Format(loRunning, VATFormat)
                abc = ("UPDATE FMSVAT set NEWDOCNO='" & RunNO & "' WHERE VATID=" & rs.options.Rs.Fields("VATID").Value)
                cn.Execute(abc)
                abc = ("UPDATE FMSVATINSERT set NEWDOCNO='" & RunNO & "'  WHERE VATID=" & rs.options.Rs.Fields("VATID").Value)
                cn.Execute(abc)
                If RERUN Then
                    abc = ("if not exists (select * from FMSRUN where BATCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and ENTRY ='" & rs.options.Rs.Fields("ENTRY").Value & "' and TRANSNBR ='" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "' and SOURCE ='" & rs.options.Rs.Fields("SOURCE").Value & "') begin INSERT INTO FMSRUN VALUES('" & rs.options.Rs.Fields("BATCH").Value & "','" & rs.options.Rs.Fields("ENTRY").Value & "','" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "','" & Format(loRunning, VATFormat) & "','" & rs.options.Rs.Fields("SOURCE").Value & "') end")
                    cn.Execute(abc)
                End If
            ElseIf DATEMODE = "DOCU" Then
                RunNO = loPrefix & Mid(rs.options.Rs.Fields("TXDATE").Value, 3, 2) & Mid(rs.options.Rs.Fields("TXDATE").Value, 5, 2) & Format(loRunning, VATFormat)
                abc = ("UPDATE FMSVAT set NEWDOCNO='" & RunNO & "' WHERE VATID=" & rs.options.Rs.Fields("VATID").Value)
                cn.Execute(abc)
                abc = ("UPDATE FMSVATINSERT set NEWDOCNO='" & RunNO & "'  WHERE VATID=" & rs.options.Rs.Fields("VATID").Value)
                cn.Execute(abc)
                If RERUN Then
                    abc = ("if not exists (select * from FMSRUN where BATCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and ENTRY ='" & rs.options.Rs.Fields("ENTRY").Value & "' and TRANSNBR ='" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "' and SOURCE ='" & rs.options.Rs.Fields("SOURCE").Value & "') begin INSERT INTO FMSRUN VALUES('" & rs.options.Rs.Fields("BATCH").Value & "','" & rs.options.Rs.Fields("ENTRY").Value & "','" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "','" & Format(loRunning, VATFormat) & "','" & rs.options.Rs.Fields("SOURCE").Value & "') end")
                    cn.Execute(abc)
                End If
            Else
                RunNO = loPrefix & Mid(rs.options.Rs.Fields("TXDATE").Value, 3, 2) & Mid(rs.options.Rs.Fields("TXDATE").Value, 5, 2) & Format(loRunning, VATFormat)
                abc = ("UPDATE FMSVAT set NEWDOCNO='" & RunNO & "' WHERE VATID=" & rs.options.Rs.Fields("VATID").Value)
                cn.Execute(abc)
                abc = ("UPDATE FMSVATINSERT set NEWDOCNO='" & RunNO & "'  WHERE VATID=" & rs.options.Rs.Fields("VATID").Value)
                cn.Execute(abc)
                If RERUN Then
                    abc = ("if not exists (select * from FMSRUN where BATCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and ENTRY ='" & rs.options.Rs.Fields("ENTRY").Value & "' and TRANSNBR ='" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "' and SOURCE ='" & rs.options.Rs.Fields("SOURCE").Value & "') begin INSERT INTO FMSRUN VALUES('" & rs.options.Rs.Fields("BATCH").Value & "','" & rs.options.Rs.Fields("ENTRY").Value & "','" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "','" & Format(loRunning, VATFormat) & "','" & rs.options.Rs.Fields("SOURCE").Value & "') end")
                    cn.Execute(abc)
                End If
            End If

Gonext:

            If UPRUN = True Then
                If RunNO = "" Then
                    RunNO = rs.options.Rs.Fields("NEWDOCNO").Value.ToString.Trim
                End If
                If rs.options.Rs.Fields("SOURCE").Value.ToString.Trim = "AP" Then
                    If IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) = "0" Then
                        sqlstr = "Update APIBHO SET VALUE='" & RunNO & "' WHERE OPTFIELD='VATVOUCHER' and CNTBTCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and CNTITEM ='" & rs.options.Rs.Fields("ENTRY").Value & "'"
                    Else
                        sqlstr = "Update APIBDO SET VALUE='" & RunNO & "' WHERE OPTFIELD='VATVOUCHER' and CNTBTCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and CNTITEM ='" & rs.options.Rs.Fields("ENTRY").Value & "' and CNTLINE ='" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value.ToString.Trim) & "'"
                    End If
                    cnACCPAC.Execute(sqlstr)
                End If
                If rs.options.Rs.Fields("SOURCE").Value.ToString.Trim = "PO" Then
                    sqlstr = "Update APIBDO SET VALUE='" & RunNO & "' WHERE OPTFIELD='VATVOUCHER' and CNTBTCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and CNTITEM ='" & rs.options.Rs.Fields("ENTRY").Value & "' and CNTLINE ='" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value.ToString.Trim) & "'"
                    cnACCPAC.Execute(sqlstr)
                End If
                If rs.options.Rs.Fields("SOURCE").Value.ToString.Trim = "CB" Then
                    sqlstr = "update CBBTDTO  set VALUE ='" & RunNO & "' from CBBTDTO where OPTFIELD='VATVOUCHER' and BATCHID =" & rs.options.Rs.Fields("BATCH").Value & " and ENTRYNO =" & rs.options.Rs.Fields("ENTRY").Value & " "
                    cnACCPAC.Execute(sqlstr)
                End If
                If rs.options.Rs.Fields("SOURCE").Value.ToString.Trim = "GL" Then
                    sqlstr = "Update GLJEDO SET VALUE='" & RunNO & "' WHERE OPTFIELD='VATVOUCHER' and BATCHNBR =" & rs.options.Rs.Fields("BATCH").Value & " and JOURNALID =" & rs.options.Rs.Fields("ENTRY").Value & " and TRANSNBR =" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value.ToString.Trim) & ""
                    cnACCPAC.Execute(sqlstr)
                End If

            End If
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        cn.Close()
        cnACCPAC.Close()
        rs = Nothing
        cn = Nothing
        Exit Sub

ErrHandler:
        ErrMsg = Err.Description
        WriteLog("<ProcessNewDocNo>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessNewDocNo> sqlstr : " & abc & Environment.NewLine & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    'สำหรับ BOSHCH
    Public Sub ProcessNewDocNoBOSHCH()
        On Error GoTo ErrHandler
        Dim cn, cnACCPAC As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim loPrefix As String
        Dim loRunning As Integer = 0
        Dim PreLocation As String

        cn = New ADODB.Connection
        cnACCPAC = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        cn.ConnectionTimeout = 60
        cn.Open(ConnVAT)
        cn.CommandTimeout = 3600

        cnACCPAC.ConnectionTimeout = 60
        cnACCPAC.Open(ConnACCPAC)
        cnACCPAC.CommandTimeout = 3600
        PreLocation = ""
        Dim IsSource As String = ""
        Dim Location As New Collection

        If RERUN Then
            sqlstr = "UPDATE " & Environment.NewLine
            sqlstr &= "FMSVAT" & Environment.NewLine
            sqlstr &= "SET" & Environment.NewLine
            sqlstr &= "FMSVAT.NEWDOCNO  =FMSVLOC.LOCPREFIX+right(left(FMSVAT.TXDATE,6),4)+FMSRUN.RUNNING " & Environment.NewLine
            sqlstr &= "FROM" & Environment.NewLine
            sqlstr &= "FMSVAT LEFT JOIN FMSVLOC ON FMSVAT.LOCID = FMSVLOC.LOCID" & Environment.NewLine
            sqlstr &= "LEFT JOIN FMSRUN on (FMSVAT.BATCH =FMSRUN.BATCH) and (FMSVAT.ENTRY =FMSRUN.ENTRY  ) and ((CASE WHEN isnull(FMSVAT.TRANSNBR,0)='' then 0 else isnull(FMSVAT.TRANSNBR,0) end) =(CASE WHEN isnull(FMSRUN.TRANSNBR,0)='' then 0 else isnull(FMSRUN.TRANSNBR,0) end))and (FMSVAT.SOURCE  =FMSRUN.SOURCE  )" & Environment.NewLine
            sqlstr &= "where ISNULL (fmsrun.running,'')<>'';" & Environment.NewLine

            sqlstr &= "UPDATE " & Environment.NewLine
            sqlstr &= "FMSVATINSERT" & Environment.NewLine
            sqlstr &= "SET" & Environment.NewLine
            sqlstr &= "FMSVATINSERT.NEWDOCNO  = FMSVLOC.LOCPREFIX+right(left(FMSVATINSERT.TXDATE,6),4)+FMSRUN.RUNNING" & Environment.NewLine
            sqlstr &= "FROM" & Environment.NewLine
            sqlstr &= "FMSVATINSERT LEFT JOIN FMSVLOC ON FMSVATINSERT.LOCID = FMSVLOC.LOCID" & Environment.NewLine
            sqlstr &= "LEFT JOIN FMSRUN on (FMSVATINSERT.BATCH =FMSRUN.BATCH) and (FMSVATINSERT.ENTRY =FMSRUN.ENTRY  ) and ((CASE WHEN isnull(FMSVATINSERT.TRANSNBR,0)='' then 0 else isnull(FMSVATINSERT.TRANSNBR,0) end) =(CASE WHEN isnull(FMSRUN.TRANSNBR,0)='' then 0 else isnull(FMSRUN.TRANSNBR,0) end))and (FMSVATINSERT.SOURCE  =FMSRUN.SOURCE  )" & Environment.NewLine
            sqlstr &= "where ISNULL (fmsrun.running,'')<>'';" & Environment.NewLine
            cn.Execute(sqlstr)

            Dim RS_Location As New BaseClass.BaseODBCIO("select  isnull(max(RUNNING),0)as RUNNING ,LOCID ,right(left(TXDATE,6),4) as TXDATE from FMSVAT  LEFT JOIN FMSRUN on (FMSVAT.BATCH =FMSRUN.BATCH) and (FMSVAT.ENTRY =FMSRUN.ENTRY  ) and ((CASE WHEN isnull(FMSVAT.TRANSNBR,0)='' then 0 else isnull(FMSVAT.TRANSNBR,0) end) =(CASE WHEN isnull(FMSRUN.TRANSNBR,0)='' then 0 else isnull(FMSRUN.TRANSNBR,0) end)) and (FMSVAT.SOURCE  =FMSRUN.SOURCE  ) group by LOCID,right(left(TXDATE,6),4)", cn)
            For iLoca As Integer = 0 To RS_Location.options.QueryDT.Rows.Count - 1
                Location.Add(New String() {RS_Location.options.QueryDT.Rows(iLoca).Item("LOCID"), RS_Location.options.QueryDT.Rows(iLoca).Item("TXDATE"), RS_Location.options.QueryDT.Rows(iLoca).Item("RUNNING")}, RS_Location.options.QueryDT.Rows(iLoca).Item("LOCID") + RS_Location.options.QueryDT.Rows(iLoca).Item("TXDATE"))
            Next
        End If

        sqlstr = "SELECT FMSVAT.TXDATE,FMSVAT.DOCNO,FMSVAT.VATID,FMSVAT.TXDATE,FMSVLOC.LOCPREFIX,FMSVLOC.LOCID,FMSVAT.BATCH ,FMSVAT.ENTRY ,(CASE WHEN isnull(FMSVAT.TRANSNBR,0)='' then 0 else isnull(FMSVAT.TRANSNBR,0) end)as LINE ,FMSVAT.SOURCE" & IIf(RERUN, " ,isnull(FMSRUN.RUNNING,'') as RUNNING", "") & ",FMSVAT.NEWDOCNO " & Environment.NewLine
        sqlstr &= ",(CASE	WHEN ltrim(rtrim(FMSVAT.SOURCE))in('AP','PO') then (select top 1 APOBL.AUDTDATE  from " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & " APOBL left join " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "APIBD on (APOBL.CNTBTCH =APIBD.CNTBTCH and APOBL.CNTITEM =APIBD.CNTITEM ) where APOBL.CNTBTCH =FMSVAT.BATCH  and  APOBL.CNTITEM =FMSVAT.ENTRY) " & Environment.NewLine
        sqlstr &= "		    WHEN ltrim(rtrim(FMSVAT.SOURCE))='GL' then (select top 1 GLPOST.AUDTDATE  from " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "GLPOST left join " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "GLJED  on (GLPOST.BATCHNBR =GLJED.BATCHNBR and GLPOST.ENTRYNBR =GLJED.JOURNALID ) where GLPOST.BATCHNBR=convert(char(6),FMSVAT.BATCH)  and  GLPOST.ENTRYNBR =convert(char(6),FMSVAT.ENTRY)) " & Environment.NewLine
        If FindApp("CB") = True Then
            sqlstr &= "		    WHEN ltrim(rtrim(FMSVAT.SOURCE))='CB' then (select top 1 CBTRHD.AUDTDATE  from " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "CBTRHD left join " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "CBBTDT   on (CBTRHD.BATCHID  =CBBTDT.BATCHID and CBTRHD.ENTRYNO =CBBTDT.ENTRYNO  ) where CBTRHD.BATCHID=FMSVAT.BATCH  and  CBTRHD.ENTRYNO =FMSVAT.ENTRY) " & Environment.NewLine
        End If
        sqlstr &= "		    WHEN ltrim(rtrim(FMSVAT.SOURCE))in('AR','OE') then (select top 1 AROBL.AUDTDATE  from " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "AROBL  left join " & ConnACCPAC.Split(";")(2).Split("=")(1) & IIf(ComDatabase = "PVSW", ".", ".dbo.") & "ARIBD    on (AROBL.CNTBTCH  =ARIBD.CNTBTCH and AROBL.CNTITEM =ARIBD.CNTITEM  ) where AROBL.CNTBTCH=FMSVAT.BATCH  and  ARIBD.CNTITEM =FMSVAT.ENTRY) " & Environment.NewLine
        sqlstr &= "end) as AUDTDATE" & Environment.NewLine
        If ComDatabase = "ORCL" Then
            sqlstr = sqlstr & " FROM FMSVAT LEFT JOIN FMSVLOC ON FMSVAT.LOCID = FMSVLOC.LOCID" & Environment.NewLine
        Else
            sqlstr = sqlstr & " FROM FMSVAT LEFT JOIN FMSVLOC ON FMSVAT.LOCID = FMSVLOC.LOCID" & Environment.NewLine
        End If
        If RERUN Then sqlstr &= " LEFT JOIN FMSRUN on (FMSVAT.BATCH =FMSRUN.BATCH) and (FMSVAT.ENTRY =FMSRUN.ENTRY  ) and ((CASE WHEN isnull(FMSVAT.TRANSNBR,0)='' then 0 else isnull(FMSVAT.TRANSNBR,0) end) =(CASE WHEN isnull(FMSRUN.TRANSNBR,0)='' then 0 else isnull(FMSRUN.TRANSNBR,0)end))" & Environment.NewLine
        If CheckTTypeRunning = 1 Then
            sqlstr = sqlstr & " WHERE TTYPE = 1" & Environment.NewLine
        ElseIf CheckTTypeRunning = 2 Then
            sqlstr = sqlstr & " WHERE TTYPE = 2" & Environment.NewLine
        ElseIf CheckTTypeRunning = 3 Then
            sqlstr = sqlstr & " WHERE TTYPE = 3" & Environment.NewLine
        End If
        If CheckRate.Trim.Split("||").Length = 3 Then
            sqlstr &= " and RATE = " & CheckRate.Trim.Split("||")(2) & Environment.NewLine
        ElseIf CheckRate.Trim = "None" Then
            sqlstr &= " and (VTYPE = 0)" & Environment.NewLine
        ElseIf CheckRate.Trim = "unlike 7" Then
            sqlstr &= " and (RATE <>7 and RATE <>0)" & Environment.NewLine
        End If
        sqlstr &= " ORDER BY AUDTDATE,BATCH,ENTRY ,LINE"
        rs.Open(sqlstr, cn)
        Dim collvat As New Collection
        Dim TMPTXDATE As String = ""
        Do While rs.options.Rs.EOF = False
            Dim RunNO As String = ""
            If RERUN = True Then
                If rs.options.Rs.Fields("RUNNING").Value.ToString.Trim <> "" Then
                    GoTo Gonext
                Else
                    If PreLocation <> rs.options.Rs.Fields("LOCID").Value Then
                        If PreLocation <> "" Then
                            collvat.Remove(PreLocation)
                            collvat.Add(loRunning, PreLocation)
                        End If
                        If Location.Count > 0 Then
                            loRunning = CType(Location(rs.options.Rs.Fields("LOCID").Value & rs.options.Rs.Fields("TXDATE").Value.ToString.Substring(2, 4)), String())(2)
                        Else
                            collvat.Add(0, rs.options.Rs.Fields("LOCID").Value)
                            loRunning = 0
                        End If
                        PreLocation = rs.options.Rs.Fields("LOCID").Value
                    ElseIf TMPTXDATE <> Mid(rs.options.Rs.Fields("TXDATE").Value, 3, 4) Then
                        If Location.Count > 0 Then
                            CType(Location(rs.options.Rs.Fields("LOCID").Value & rs.options.Rs.Fields("TXDATE").Value.ToString.Substring(2, 4)), String())(2) = loRunning
                            loRunning = 0
                        End If
                    End If
                End If
            Else
                If PreLocation <> rs.options.Rs.Fields("LOCID").Value Then
                    If PreLocation <> "" Then
                        collvat.Remove(PreLocation)
                        collvat.Add(loRunning, PreLocation)
                    End If

                    If collvat.Contains(rs.options.Rs.Fields("LOCID").Value) Then
                        loRunning = collvat(rs.options.Rs.Fields("LOCID").Value).ToString
                    Else
                        collvat.Add(0, rs.options.Rs.Fields("LOCID").Value)
                        loRunning = 0
                    End If
                    PreLocation = rs.options.Rs.Fields("LOCID").Value
                End If
            End If
            loPrefix = IIf(IsDBNull(rs.options.Rs.Fields("LOCPREFIX").Value), "", Trim(rs.options.Rs.Fields("LOCPREFIX").Value))
            loRunning = loRunning + 1
            TMPTXDATE = Mid(rs.options.Rs.Fields("TXDATE").Value, 3, 4)
            Dim abc As String = ""
            If RERUN = True Then
                RunNO = loPrefix & Mid(rs.options.Rs.Fields("TXDATE").Value, 3, 4) & Format(loRunning, VATFormat)
                abc = ("UPDATE FMSVAT set NEWDOCNO='" & RunNO & "' WHERE VATID=" & rs.options.Rs.Fields("VATID").Value)
                cn.Execute(abc)
                abc = ("UPDATE FMSVATINSERT set NEWDOCNO='" & RunNO & "'" & " WHERE VATID=" & rs.options.Rs.Fields("VATID").Value)
                cn.Execute(abc)
                If RERUN Then
                    abc = ("if not exists (select * from FMSRUN where BATCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and ENTRY ='" & rs.options.Rs.Fields("ENTRY").Value & "' and TRANSNBR ='" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "' and SOURCE ='" & rs.options.Rs.Fields("SOURCE").Value & "') begin INSERT INTO FMSRUN VALUES('" & rs.options.Rs.Fields("BATCH").Value & "','" & rs.options.Rs.Fields("ENTRY").Value & "','" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "','" & Format(loRunning, VATFormat) & "','" & rs.options.Rs.Fields("SOURCE").Value & "') end")
                    cn.Execute(abc)
                End If
            ElseIf DATEMODE = "DOCU" Then
                RunNO = loPrefix & Mid(rs.options.Rs.Fields("TXDATE").Value, 3, 4) & Format(loRunning, VATFormat)
                abc = ("UPDATE FMSVAT set NEWDOCNO='" & RunNO & "' WHERE VATID=" & rs.options.Rs.Fields("VATID").Value)
                cn.Execute(abc)
                abc = ("UPDAT FMSVATINSERT set NEWDOCNO='" & RunNO & "'" & " WHERE VATID=" & rs.options.Rs.Fields("VATID").Value)
                cn.Execute(abc)
                If RERUN Then
                    abc = ("if not exists (select * from FMSRUN where BATCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and ENTRY ='" & rs.options.Rs.Fields("ENTRY").Value & "' and TRANSNBR ='" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "' and SOURCE ='" & rs.options.Rs.Fields("SOURCE").Value & "') begin INSERT INTO FMSRUN VALUES('" & rs.options.Rs.Fields("BATCH").Value & "','" & rs.options.Rs.Fields("ENTRY").Value & "','" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "','" & Format(loRunning, VATFormat) & "','" & rs.options.Rs.Fields("SOURCE").Value & "') end")
                    cn.Execute(abc)
                End If
            Else
                RunNO = loPrefix & Mid(rs.options.Rs.Fields("TXDATE").Value, 3, 4) & Format(loRunning, VATFormat)
                abc = ("UPDATE FMSVAT set NEWDOCNO='" & RunNO & "' WHERE VATID=" & rs.options.Rs.Fields("VATID").Value)
                cn.Execute(abc)
                abc = ("UPDATE FMSVATINSERT set NEWDOCNO='" & RunNO & "'" & " WHERE VATID=" & rs.options.Rs.Fields("VATID").Value)
                cn.Execute(abc)
                If RERUN Then
                    abc = ("if not exists (select * from FMSRUN where BATCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and ENTRY ='" & rs.options.Rs.Fields("ENTRY").Value & "' and TRANSNBR ='" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "' and SOURCE ='" & rs.options.Rs.Fields("SOURCE").Value & "') begin INSERT INTO FMSRUN VALUES('" & rs.options.Rs.Fields("BATCH").Value & "','" & rs.options.Rs.Fields("ENTRY").Value & "','" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) & "','" & Format(loRunning, VATFormat) & "','" & rs.options.Rs.Fields("SOURCE").Value & "') end")
                    cn.Execute(abc)
                End If
            End If

Gonext:

            If UPRUN = True Then
                If RunNO = "" Then
                    RunNO = rs.options.Rs.Fields("NEWDOCNO").Value.ToString.Trim
                End If
                If rs.options.Rs.Fields("SOURCE").Value.ToString.Trim = "AP" Then
                    If IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value) = "0" Then
                        sqlstr = "Update APIBHO SET VALUE='" & RunNO & "' WHERE OPTFIELD='VATVOUCHER' and CNTBTCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and CNTITEM ='" & rs.options.Rs.Fields("ENTRY").Value & "'"
                    Else
                        sqlstr = "Update APIBDO SET VALUE='" & RunNO & "' WHERE OPTFIELD='VATVOUCHER' and CNTBTCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and CNTITEM ='" & rs.options.Rs.Fields("ENTRY").Value & "' and CNTLINE ='" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value.ToString.Trim) & "'"
                    End If
                    cnACCPAC.Execute(sqlstr)
                End If
                If rs.options.Rs.Fields("SOURCE").Value.ToString.Trim = "PO" Then
                    sqlstr = "Update APIBDO SET VALUE='" & RunNO & "' WHERE OPTFIELD='VATVOUCHER' and CNTBTCH ='" & rs.options.Rs.Fields("BATCH").Value & "' and CNTITEM ='" & rs.options.Rs.Fields("ENTRY").Value & "' and CNTLINE ='" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value.ToString.Trim) & "'"
                    cnACCPAC.Execute(sqlstr)
                End If
                If rs.options.Rs.Fields("SOURCE").Value.ToString.Trim = "CB" Then
                    sqlstr = "update CBBTDTO  set VALUE ='" & RunNO & "' from CBBTDTO where OPTFIELD='VATVOUCHER' and BATCHID =" & rs.options.Rs.Fields("BATCH").Value & " and ENTRYNO =" & rs.options.Rs.Fields("ENTRY").Value & " "
                    cnACCPAC.Execute(sqlstr)
                End If
                If rs.options.Rs.Fields("SOURCE").Value.ToString.Trim = "GL" Then
                    sqlstr = "Update GLJEDO SET VALUE='" & RunNO & "' WHERE OPTFIELD='VATVOUCHER' and BATCHNBR =" & rs.options.Rs.Fields("BATCH").Value & " and JOURNALID =" & rs.options.Rs.Fields("ENTRY").Value & " and TRANSNBR =" & IIf(rs.options.Rs.Fields("LINE").Value.ToString.Trim = "", "0", rs.options.Rs.Fields("LINE").Value.ToString.Trim) & ""
                    cnACCPAC.Execute(sqlstr)
                End If

            End If
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop
        rs.options.Rs.Close()
        cn.Close()
        cnACCPAC.Close()
        rs = Nothing
        cn = Nothing
        Exit Sub

ErrHandler:
        ErrMsg = Err.Description
        WriteLog("<ProcessNewDocNo>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessNewDocNo> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    Public Sub ProcessManulaNewDocNo()

        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim rs1 As BaseClass.BaseODBCIO
        Dim rs2 As BaseClass.BaseODBCIO

        Dim sqlstr As String
        Dim loPrefix, loMonth As String
        Dim loRunning As String
        Dim PreLocation As String
        Dim cnACCPAC As ADODB.Connection
        Dim iffs As clsFuntion
        iffs = New clsFuntion

        Dim VATCOMMENT As String
        Dim VatRunning() As String
        Dim LastNonEmpty As Integer
        Dim i As Integer
        On Error GoTo ErrHandler

        PreLocation = ""
        cnACCPAC = New ADODB.Connection

        cnACCPAC.ConnectionTimeout = 60
        cnACCPAC.Open(ConnACCPAC)
        cnACCPAC.CommandTimeout = 3600


        cn = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        cn.ConnectionTimeout = 60
        cn.Open(ConnVAT)
        cn.CommandTimeout = 3600
        loRunning = ""

        sqlstr = "SELECT     FMSVAT.SOURCE, FMSVAT.BATCH, FMSVAT.ENTRY, FMSVAT.IDINVC, FMSVAT.TXDATE, FMSVAT.DOCNO, FMSVAT.VATID, FMSVAT.INVDATE, " & " FMSVLOC.LOCPREFIX, FMSVLOC.LOCID, FMSVAT.NEWDOCNO,FMSVAT.VATCOMMENT,FMSVAT.TRANSNBR"

        If ComDatabase = "ORCL" Then
            sqlstr = sqlstr & " FROM FMSVAT LEFT JOIN FMSVLOC ON FMSVAT.LOCID = FMSVLOC.LOCID"
        Else
            sqlstr = sqlstr & " FROM FMSVAT LEFT JOIN FMSVLOC ON FMSVAT.LOCID = FMSVLOC.LOCID"
        End If

        If CheckTTypeRunning = 1 Then
            sqlstr = sqlstr & " WHERE TTYPE = 1"
            loRunning = ""
        ElseIf CheckTTypeRunning = 2 Then
            sqlstr = sqlstr & " WHERE TTYPE = 2"
            loRunning = ""
        ElseIf CheckTTypeRunning = 3 Then
            sqlstr = sqlstr & " WHERE TTYPE = 3"
            loRunning = ""
        End If

        If CheckRate.Trim.Split("||").Length = 3 Then
            sqlstr &= " and RATE = " & CheckRate.Trim.Split("||")(2)
        ElseIf CheckRate.Trim = "None" Then
            sqlstr &= " and (VTYPE = 0)"
        ElseIf CheckRate.Trim = "unlike 7" Then
            sqlstr &= " and (RATE <> 7 or RATE <> 0)"
        End If

        sqlstr = sqlstr & " ORDER BY FMSVAT.INVDATE,FMSVAT.BATCH, FMSVLOC.LOCID, FMSVAT.DOCNO, FMSVAT.VATID"
        rs.Open(sqlstr, cn)

        Do While rs.options.Rs.EOF = False

            If PreLocation <> rs.options.Rs.Fields("LOCID").Value Then
                PreLocation = rs.options.Rs.Fields("LOCID").Value
                loRunning = ""
            End If

            ' หา เลข Running จาก Optional field จาก ap ,gl และ CB ให้ไปหยิบที่ misccode
            If Trim(rs.options.Rs.Fields("SOURCE").Value) = "AP" Then

                sqlstr = " SELECT CNTBTCH, CNTITEM, OPTFIELD, VALUE FROM APIBHO " & Environment.NewLine
                sqlstr &= " WHERE (OPTFIELD = 'VATRUNNO') AND (CNTBTCH = " & rs.options.Rs.Fields("BATCH").Value & " AND CNTITEM = " & rs.options.Rs.Fields("ENTRY").Value & ")"
                rs1 = New BaseClass.BaseODBCIO
                rs1.Open(sqlstr, cnACCPAC)

                If rs1.options.Rs.RecordCount > 0 Then
                    loRunning = rs1.options.Rs.Fields("VALUE").Value.ToString.Trim
                Else
                    loRunning = ""
                End If



                If loRunning.Trim = "" And Trim(rs.options.Rs.Fields("TRANSNBR").Value) <> "" Then

                    sqlstr = " SELECT CNTBTCH, CNTITEM, OPTFIELD, VALUE FROM APIBDO " & Environment.NewLine
                    sqlstr &= " WHERE (OPTFIELD = 'VATRUNNO') AND (CNTBTCH = " & rs.options.Rs.Fields("BATCH").Value & " AND CNTITEM = " & rs.options.Rs.Fields("ENTRY").Value & ") AND (CNTLINE = " & Trim(rs.options.Rs.Fields("TRANSNBR").Value) & ")"
                    rs2 = New BaseClass.BaseODBCIO
                    rs2.Open(sqlstr, cnACCPAC)

                    If rs2.options.Rs.RecordCount > 0 Then
                        loRunning = rs2.options.Rs.Fields("VALUE").Value.ToString.Trim
                    Else
                        loRunning = ""
                    End If

                End If


            ElseIf Trim(rs.options.Rs.Fields("SOURCE").Value) = "GL" Then

                sqlstr = " SELECT GLJEH.BATCHID, GLJEH.BTCHENTRY, GLJEDO.OPTFIELD, GLJEDO.VALUE,GLJED.TRANSNBR " & Environment.NewLine
                sqlstr &= " FROM GLJEH "
                sqlstr &= " LEFT JOIN GLJED ON GLJEH.BATCHID = GLJED.BATCHNBR AND GLJEH.BTCHENTRY = GLJED.JOURNALID " & Environment.NewLine
                sqlstr &= " LEFT JOIN GLJEDO ON GLJED.BATCHNBR = GLJEDO.BATCHNBR AND GLJED.JOURNALID = GLJEDO.JOURNALID " & Environment.NewLine
                sqlstr &= " AND GLJED.TRANSNBR = GLJEDO.TRANSNBR " & Environment.NewLine
                sqlstr &= " WHERE (GLJEDO.OPTFIELD = 'VATRUNNO') " & Environment.NewLine
                sqlstr &= " AND (GLJEH.BATCHID = " & Trim(rs.options.Rs.Fields("BATCH").Value) & " AND GLJEH.BTCHENTRY = " & Trim(rs.options.Rs.Fields("ENTRY").Value) & ") AND (GLJED.TRANSNBR = " & Trim(rs.options.Rs.Fields("TRANSNBR").Value) & ")"
                rs1 = New BaseClass.BaseODBCIO
                rs1.Open(sqlstr, cnACCPAC)

                If rs1.options.Rs.RecordCount > 0 Then
                    loRunning = rs1.options.Rs.Fields("VALUE").Value.ToString.Trim
                Else
                    loRunning = ""
                End If

            ElseIf Trim(rs.options.Rs.Fields("SOURCE").Value) = "CB" Then

                VATCOMMENT = Trim(rs.options.Rs.Fields("VATCOMMENT").Value)
                VatRunning = Split(VATCOMMENT, ",", -1, CompareMethod.Text)
                i = UBound(VatRunning)

                If i <> 3 Then
                    loRunning = ""
                Else
                    loRunning = VatRunning(i)
                End If

                If loRunning.Trim = "" Then

                    sqlstr = " SELECT BATCHID, ENTRYNO, OPTFIELD, VALUE FROM CBTRHD " & Environment.NewLine
                    sqlstr &= " LEFT JOIN CBTRDT ON CBTRHD.BANKCODE = CBTRDT.BANKCODE AND CBTRHD.REFERENCE = CBTRDT.REFERENCE AND CBTRHD.UNIQUIFIER = CBTRDT.UNIQUIFIER" & Environment.NewLine
                    sqlstr &= " LEFT JOIN CBTRDTO ON CBTRDT.BANKCODE = CBTRDTO.BANKCODE AND CBTRDT.REFERENCE = CBTRDTO.REFERENCE AND CBTRDT.UNIQUIFIER = CBTRDTO.UNIQUIFIER" & Environment.NewLine
                    sqlstr &= " AND CBTRDT.DETAILNO = CBTRDTO.DETAILNO" & Environment.NewLine
                    sqlstr &= " WHERE OPTFIELD = 'VATRUNNO' AND (BATCHID = " & rs.options.Rs.Fields("BATCH").Value & " AND ENTRYNO = " & rs.options.Rs.Fields("ENTRY").Value & ") AND (CBTRDTO.DETAILNO = " & Trim(rs.options.Rs.Fields("TRANSNBR").Value) & ")"
                    rs1 = New BaseClass.BaseODBCIO
                    rs1.Open(sqlstr, cnACCPAC)

                    If rs1.options.Rs.RecordCount > 0 Then
                        loRunning = rs1.options.Rs.Fields("VALUE").Value.ToString.Trim
                    Else
                        loRunning = ""
                    End If

                End If


            End If


            loMonth = rs.options.Rs.Collect(7)
            loMonth = Mid(rs.options.Rs.Fields("INVDATE").Value, 5, 2)
            loPrefix = IIf(IsDBNull(rs.options.Rs.Fields("LOCPREFIX").Value), "", (rs.options.Rs.Fields("LOCPREFIX").Value))
            loPrefix = loPrefix.Trim

            sqlstr = "UPDATE FMSVAT SET NEWDOCNO = '" & loPrefix & loMonth & "/"
            sqlstr = sqlstr & loRunning & "'"
            sqlstr = sqlstr & " WHERE VATID = " & rs.options.Rs.Fields("VATID").Value
            rs1 = Nothing
            loRunning = ""
            cn.Execute(sqlstr)
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop

        rs.options.Rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
        Exit Sub

ErrHandler:
        ErrMsg = Err.Description
        WriteLog("<ProcessNewDocNo>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessNewDocNo> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    Public Sub ProcessSetDate()
        Dim cn As ADODB.Connection
        Dim rs, rs2 As BaseClass.BaseODBCIO
        Dim sqlstr As String
        On Error GoTo ErrHandler
        cn = New ADODB.Connection
        cn.Open(ConnVAT)
        sqlstr = "Update FMSVSET set DATEFROM = '" & DATEFROM & "',DATETO='" & DATETO & "'"
        cn.Execute(sqlstr)
        WriteLog("<ProcessComplete> table ready")
        cn.Close()
        cn = Nothing
        Exit Sub

ErrHandler:
        ErrMsg = Err.Description
        WriteLog("<ProcessComplete>" & Err.Number & "  " & Err.Description)
        WriteLog("<ProcessComplete> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    Public Sub SaveVatset()
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String

        On Error GoTo ErrHandler

        cn = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        cn.ConnectionTimeout = 60
        cn.Open(ConnVAT)

        WriteLog("<SaveVatset> Check table")
        If TableHere("FMSVSET", ConnVAT) = False Then
            PrepareVatSet()
        End If

        sqlstr = "Update FMSVSET set TAXNO = '" & TAXNO & "'"
        cn.Execute(sqlstr)
        cn.Close()
        WriteLog("<SaveVatset> table ready")
        rs = Nothing
        cn = Nothing
        Exit Sub

ErrHandler:
        cn.RollbackTrans()
        ErrMsg = Err.Description
        WriteLog("<SaveVatset>" & Err.Number & "  " & Err.Description)
        WriteLog("<SaveVatset> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub
    Public Sub PrepareType()
        Dim cnVAT As ADODB.Connection
        Dim sqlstr As String

        On Error GoTo ErrHandler

        cnVAT = New ADODB.Connection

        If TableHere("FMSTTYPE", ConnVAT) = False Then
            sqlstr = "CREATE TABLE FMSTTYPE"
            sqlstr = sqlstr & " (TTYPE  DECIMAL(9,0),"
            sqlstr = sqlstr & " TTYPENAME   CHAR(50))"
            cnVAT.Open(ConnVAT)
            cnVAT.Execute(sqlstr)
            sqlstr = InsStr & "FMSTTYPE(TTYPE,TTYPENAME) Values(1,'PURCHASE')"
            cnVAT.Execute(sqlstr)
            sqlstr = InsStr & "FMSTTYPE(TTYPE,TTYPENAME) Values(2,'SALE')"
            cnVAT.Execute(sqlstr)
            sqlstr = InsStr & "FMSTTYPE(TTYPE,TTYPENAME) Values(3,'EXPENSE')"
            cnVAT.Execute(sqlstr)
            cnVAT.Close()
        End If

        Exit Sub

ErrHandler:
        cnVAT.RollbackTrans()
        ErrMsg = Err.Description
        WriteLog("<PrepareType>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    Public Function ViewHere(ByRef viewname As String) As Boolean

        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String

        On Error GoTo ErrHandler

        cn = New ADODB.Connection
        sqlstr = ""
        rs = New BaseClass.BaseODBCIO

        If ComDatabase = "MSSQL" Then
            sqlstr = "select * from sysobjects"
            sqlstr = sqlstr & " where name='" & viewname & "'"
        ElseIf ComDatabase = "PVSW" Then
            sqlstr = "SELECT * FROM X$View"
            sqlstr = sqlstr & " where Xv$Name = '" & viewname & "'"
        ElseIf ComDatabase = "ORCL" Then
            sqlstr = "SELECT * FROM USER_TABLES"
            sqlstr = sqlstr & " where TABLE_NAME = '" & viewname & "'"
        End If
        cn.ConnectionTimeout = 60
        cn.Open(ConnACCPAC)
        rs.Open(sqlstr, cn)
        If rs.options.Rs.EOF = False Then
            ViewHere = True
        Else
            ViewHere = False
        End If

        rs = Nothing
        cn = Nothing
        Exit Function

ErrHandler:
        If Err.Number = -2147217865 Then ' no table
            ViewHere = False
            WriteLog("No Table " & viewname)
        Else
            ErrMsg = Err.Description
            WriteLog(Err.Number & "  " & Err.Description)
            If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End If
    End Function

    Public Function TableHere(ByRef tablename As String, ByVal STRDB As String) As Boolean

        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String

        On Error GoTo ErrHandler

        cn = New ADODB.Connection
        sqlstr = ""
        rs = New BaseClass.BaseODBCIO

        If ComDatabase = "MSSQL" Then
            sqlstr = "select * from sysobjects"
            sqlstr = sqlstr & " where name='" & tablename & "'"
        ElseIf ComDatabase = "PVSW" Then
            sqlstr = " SELECT " & """X$File"".""Xf$Name""" & " FROM " & """X$File""" & " WHERE (" & """X$File"".""Xf$Name""" & "='" & tablename & "')"
        ElseIf ComDatabase = "ORCL" Then
            sqlstr = "SELECT * FROM USER_TABLES"
            sqlstr = sqlstr & " where TABLE_NAME = '" & tablename & "'"
        End If
        cn.ConnectionTimeout = 60
        cn.Open(STRDB)
        rs.Open(sqlstr, cn)

        If rs.options.QueryDT.Rows.Count > 0 Then
            cn.Close()
            TableHere = True
        Else
            cn.Close()
            TableHere = False
        End If

        rs = Nothing
        cn = Nothing
        Exit Function

ErrHandler:
        If Err.Number = -2147217865 Then ' no table
            TableHere = False
            WriteLog("No Table " & tablename)
        Else
            ErrMsg = Err.Description()
            WriteLog(Err.Number & "  " & Err.Description)
            If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End If
    End Function

    Public Sub DeleteLocation(ByRef aLOCID As String)
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String

        On Error GoTo ErrHandler

        cn = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO

        cn.Open(ConnVAT)

        WriteLog("<DeleteLocation> Check table")
        If TableHere("FMSVLOC", ConnVAT) = False Then
            PrepareLocation()
        End If

        sqlstr = "DELETE FROM FMSVLACC WHERE LOCID = '" & aLOCID & "'"
        cn.Execute(sqlstr)
        sqlstr = "DELETE FROM FMSVLOC WHERE LOCID = '" & aLOCID & "'"
        cn.Execute(sqlstr)
        cn.Close()
        WriteLog("<DeleteLocation> ready")
        rs = Nothing
        cn = Nothing
        Exit Sub

ErrHandler:
        cn.RollbackTrans()
        ErrMsg = Err.Description
        WriteLog("<DeleteLocation>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub


    Public Sub SaveLocation(ByRef aLOCID As String, ByRef aLOCNAME As String, ByRef aLOCADD As String, ByRef aLOCPrefix As String, ByRef aLOCDESC As String)
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String

        On Error GoTo ErrHandler

        cn = New ADODB.Connection
        rs = New BaseClass.BaseODBCIO
        cn.ConnectionTimeout = 60
        cn.Open(ConnVAT)

        WriteLog("<SaveLocation> Check table")
        If TableHere("FMSVLOC", ConnVAT) = False Then
            PrepareLocation()
        End If

        sqlstr = "INSERT INTO FMSVLOC(LOCID,LOCNAME,LOCADD,LOCPREFIX,LOCDESC )"
        sqlstr = sqlstr & " Values('" & limitStr(ReQuote(aLOCID), 6) & "','"
        sqlstr = sqlstr & limitStr(ReQuote(aLOCNAME), 50) & "','"
        sqlstr = sqlstr & limitStr(ReQuote(aLOCADD), 200) & "','"
        sqlstr = sqlstr & limitStr(ReQuote(aLOCPrefix), 4) & "','"
        sqlstr = sqlstr & limitStr(ReQuote(aLOCDESC), 100) & "')"
        cn.Execute(sqlstr)
        cn.Close()
        WriteLog("<SaveLocation> ready")
        rs = Nothing
        cn = Nothing
        Exit Sub

ErrHandler:
        cn.RollbackTrans()

        ErrMsg = Err.Description
        WriteLog("<SaveLocation>" & Err.Number & "  " & Err.Description)
        WriteLog("<SaveLocation> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    Public Function FindApp(ByRef loAccpacModule As String, Optional ByVal StrCon As String = "") As Boolean
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim lbFind As Boolean
        Dim sqlstr As String

        On Error GoTo ErrHandler

        cn = New ADODB.Connection
        cn.ConnectionTimeout = 60
        cn.Open(IIf(StrCon <> "", StrCon, ConnACCPAC))
        cn.CommandTimeout = 3600

        rs = New BaseClass.BaseODBCIO
        'loAccpacModule = "AR"
        sqlstr = "select * from CSAPP where SELECTOR = '" & loAccpacModule & "'"
        rs.Open(sqlstr, cn)

        If rs.options.Rs.EOF = False Then lbFind = True
        FindApp = lbFind

        cn.Close()
        cn = Nothing

        rs = Nothing

        Exit Function

ErrHandler:

        ErrMsg = Err.Description()
        WriteLog("<FindApp>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Function

    Public Function InsCB(ByRef loBatch As String, ByRef loEntry As String, ByRef loCBDATE As String, ByRef loDOCNO As String, ByRef loCHECK As String, ByRef loNAME As String, ByRef loINVAMT As Double, ByRef loINVTAX As Double, ByRef loVATCOMMENT As String, ByRef loACCTVAT As String, ByVal locbTrans As String, ByVal TXDATE As String, ByVal loRunno As String, ByVal TAXID As String, ByVal BRANCH As String, ByVal loCode As String, ByVal loTOTTAX As Double) As String
        If loNAME <> "" Then

            If loNAME.ToString.Trim.Length > 100 Then
                loNAME = Replace(loNAME, "'", "''").Trim.Substring(0, 100)
            Else
                loNAME = Replace(loNAME, "'", "''").Trim
            End If
        End If

        If loVATCOMMENT <> "" Then

            If loVATCOMMENT.ToString.Trim.Length > 1000 Then 'And loVATCOMMENT <> ""
                loVATCOMMENT = Replace(loVATCOMMENT, "'", "''").Trim.Substring(0, 1000)
            Else
                loVATCOMMENT = Replace(loVATCOMMENT, "'", "''").Trim
            End If
        End If

        If TAXID.Trim <> "" Then
            If TAXID.Trim.Length > 13 Then
                TAXID = TAXID.Substring(0, 13)
            End If
        End If
        'loVATCOMMENT = Replace(loVATCOMMENT & Space(1000 - Len(loVATCOMMENT)), "'", "''")

        loDOCNO = Replace(loDOCNO, "'", "''")
        loCHECK = Replace(loCHECK, "'", "''")
        BRANCH = ReQuote(Trim(BRANCH))
        InsCB = InsStr & "FMSCB(BATCH,ENTRY,CBDATE,DOCNO,REFERENCE,NAME,INVAMT,INVTAX,VATCOMMENT,ACCTVAT,TRANSNBR,TXDATE,RUNNO,TAXID,BRANCH,Code,TOTTAX)" & " Values('" & loBatch & "','" & loEntry & "','" & loCBDATE & "','" & limitStr(loDOCNO, 255) & "','" & limitStr(loCHECK, 12) & "','" & Trim(loNAME) & "'," & loINVAMT & "," & loINVTAX & ",'" & Trim(loVATCOMMENT) & "','" & loACCTVAT & "','" & locbTrans & "','" & TXDATE.Trim & "','" & loRunno & "','" & TAXID & "','" & BRANCH & "','" & loCode & "'," & loTOTTAX & ")"

    End Function

    Public Function InsVat(ByRef loTable As Integer, ByRef loINVDATE As String, ByRef loTXDATE As String, ByRef loIDINVC As String, ByRef loDOCNO As String, ByRef loNEWDOCNO As String, ByRef loINVNAME As String, ByRef loINVAMT As Double, ByRef loINVTAX As Double, ByRef loLOCID As String, ByRef loVTYPE As Short, ByRef loRATE As Double, ByRef loTTYPE As Integer, ByRef loACCTVAT As String, ByRef loSource As String, ByRef loBatch As String, ByRef loEntry As String, ByRef loMark As String, ByRef loVATCOMMENT As String, ByRef loCBRef As String, ByRef loTRANSNBR As String, ByRef loRunno As String, ByRef loCODETAX As String, ByRef loIDDIST As String, ByRef typeofPU As String, ByRef TranNo As String, ByRef CIF As Decimal, ByRef TaxCIF As Decimal, ByRef Verify As Integer, ByVal TAXID As String, ByVal BRANCH As String, ByRef loCode As String, ByVal loCurrency As String, ByVal loExchangrate As Decimal, ByVal loAmoutbase As Decimal, ByVal loTotalTax As Decimal) As String

        loIDINVC = Replace(Trim(loIDINVC), "'", "''") '& Space(22 - Len(Trim(loIDINVC)))
        loDOCNO = Replace(Trim(loDOCNO), "'", "''") '& Space(60 - Len(Trim(loDOCNO)))
        loNEWDOCNO = Replace(Trim(loNEWDOCNO), "'", "''") ' & Space(60 - Len(Trim(loNEWDOCNO)))
        loINVNAME = Replace(Trim(loINVNAME), "'", "''") '& Space(100 - Len(Trim(loINVNAME)))
        loACCTVAT = Trim(loACCTVAT) '& Space(45 - Len(Trim(loACCTVAT)))
        loLOCID = loLOCID '& Space(6 - Len(loLOCID))
        loSource = loSource '& Space(2 - Len(loSource))
        loBatch = Trim(loBatch) '& Space(20 - Len(Trim(loBatch)))
        loEntry = Trim(loEntry) '& Space(20 - Len(Trim(loEntry)))
        loMark = Trim(loMark) '& Space(1 - Len(loMark))
        loVATCOMMENT = Replace(Trim(loVATCOMMENT), "'", "''") '& Space(250 - Len(Trim(loVATCOMMENT)))
        loCODETAX = loCODETAX '& Space(12 - Len(loCODETAX))
        loIDDIST = loIDDIST '& Space(6 - Len(loIDDIST))
        loCBRef = Replace(Trim(loCBRef), "'", "''") '& Space(30 - Len(Trim(loCBRef)))
        loTRANSNBR = CStr(loTRANSNBR)
        loRunno = Trim(loRunno)
        typeofPU = Trim(typeofPU)
        TranNo = Trim(TranNo)
        BRANCH = ReQuote(Trim(BRANCH))
        loCode = loCode
        loCurrency = ReQuote(Trim(loCurrency))



        If loTable = 0 Then 'To Temp Table
            If ComDatabase = "PVSW" Then
                Return InsStr & "FMSVATTEMP(INVDATE,TXDATE,IDINVC,DOCNO,NEWDOCNO," & "INVNAME,INVAMT,INVTAX,LOCID,VTYPE,RATE," & "TTYPE,ACCTVAT,SOURCE,BATCH,ENTRY,MARK,VATCOMMENT,CODETAX,IDDIST,CBREF,TRANSNBR,RUNNO,TypeofPu,TranNo,CIF,TaxCIF,TAXID,BRANCH,Code,TOTTAX)" & " VALUES('" & loINVDATE & "','" & loTXDATE.Trim & "','" & loIDINVC & "','" & loDOCNO & "','" & loNEWDOCNO & "','" & loINVNAME & "','" & loINVAMT & "','" & loINVTAX & "','" & loLOCID & "','" & loVTYPE & "','" & loRATE & "','" & loTTYPE & "','" & loACCTVAT & "','" & loSource & "','" & loBatch & "','" & loEntry & "','" & loMark & "','" & loVATCOMMENT & "','" & loCODETAX & "','" & loIDDIST & "','" & loCBRef & "','" & loTRANSNBR & "','" & loRunno & "',CONVERT('" & typeofPU.Replace("'", "''") & "',SQL_VARCHAR,30),'" & TranNo & "','" & CIF & "','" & TaxCIF & "','" & TAXID & "','" & BRANCH & "','" & loCode & "'," & loTotalTax & ")"

            Else
                Return InsStr & "FMSVATTEMP(INVDATE,TXDATE,IDINVC,DOCNO,NEWDOCNO," & "INVNAME,INVAMT,INVTAX,LOCID,VTYPE,RATE," & "TTYPE,ACCTVAT,SOURCE,BATCH,ENTRY,MARK,VATCOMMENT,CODETAX,IDDIST,CBREF,TRANSNBR,RUNNO,TypeofPu,TranNo,CIF,TaxCIF,TAXID,BRANCH,Code,TOTTAX)" & " VALUES('" & loINVDATE & "','" & loTXDATE.Trim & "','" & loIDINVC & "','" & loDOCNO & "','" & loNEWDOCNO & "','" & loINVNAME & "','" & loINVAMT & "','" & loINVTAX & "','" & loLOCID & "','" & loVTYPE & "','" & loRATE & "','" & loTTYPE & "','" & loACCTVAT & "','" & loSource & "','" & loBatch & "','" & loEntry & "','" & loMark & "','" & loVATCOMMENT & "','" & loCODETAX & "','" & loIDDIST & "','" & loCBRef & "','" & loTRANSNBR & "','" & loRunno & "',convert(varchar(30),'" & typeofPU.Replace("'", "''") & "'),'" & TranNo & "','" & CIF & "','" & TaxCIF & "','" & TAXID & "','" & BRANCH & "','" & loCode & "'," & loTotalTax & ")"

            End If
        Else 'To Vat Table
            If ComDatabase = "ORCL" Then
                gVatId = gVatId + 1
                Return InsStr & "FMSVAT(VATID,INVDATE,TXDATE,IDINVC,DOCNO,NEWDOCNO," & "INVNAME,INVAMT,INVTAX,LOCID,VTYPE,RATE," & "TTYPE,ACCTVAT,SOURCE,BATCH,ENTRY,MARK,VATCOMMENT,CBREF,TRANSNBR,RUNNO,TypeofPu,TranNo,CIF,TaxCIF,Verify,TAXID,BRANCH,Code,CURRENCY,EXCHANGRATE,AMTBASETC,TOTTAX,CODETAX)" & " VALUES(" & gVatId & ",'" & loINVDATE & "','" & loTXDATE.Trim & "','" & loIDINVC & "','" & loDOCNO & "','" & loNEWDOCNO & "','" & loINVNAME & "'," & loINVAMT & "," & loINVTAX & ",'" & loLOCID & "'," & loVTYPE & "," & loRATE & "," & loTTYPE & ",'" & loACCTVAT & "','" & loSource & "','" & loBatch & "','" & loEntry & "','" & loMark & "','" & loVATCOMMENT & "','" & loCBRef & "','" & loTRANSNBR & "','" & loRunno & "',convert(varchar(30),'" & typeofPU.Replace("'", "''") & "'),'" & TranNo & "','" & CIF & "','" & TaxCIF & "','" & Verify & "','" & TAXID & "','" & BRANCH & "','" & loCode & "','" & loCurrency & "'," & loExchangrate & "," & loAmoutbase & "," & loTotalTax & ",'" & loCODETAX & "')"
            ElseIf ComDatabase = "PVSW" Then
                Return InsStr & "FMSVAT(INVDATE,TXDATE,IDINVC,DOCNO,NEWDOCNO," & "INVNAME,INVAMT,INVTAX,LOCID,VTYPE,RATE," & "TTYPE,ACCTVAT,SOURCE,BATCH,ENTRY,MARK,VATCOMMENT,CBREF,TRANSNBR,RUNNO,TypeofPu,TranNo,CIF,TaxCIF,Verify,TAXID,BRANCH,Code,CURRENCY,EXCHANGRATE,AMTBASETC,TOTTAX,CODETAX)" & " VALUES('" & loINVDATE & "','" & loTXDATE.Trim & "','" & loIDINVC & "','" & loDOCNO & "','" & loNEWDOCNO & "','" & loINVNAME & "'," & loINVAMT & "," & loINVTAX & ",'" & loLOCID & "'," & loVTYPE & "," & loRATE & "," & loTTYPE & ",'" & loACCTVAT & "','" & loSource & "','" & loBatch & "','" & loEntry & "','" & loMark & "','" & loVATCOMMENT & "','" & loCBRef & "','" & loTRANSNBR & "','" & loRunno & "',CONVERT('" & typeofPU.Replace("'", "''") & "',SQL_VARCHAR,30),'" & TranNo & "','" & CIF & "','" & TaxCIF & "','" & Verify & "','" & TAXID & "','" & BRANCH & "','" & loCode & "','" & loCurrency & "'," & loExchangrate & "," & loAmoutbase & "," & loTotalTax & ",'" & loCODETAX & "')"
            Else
                Return InsStr & "FMSVAT(INVDATE,TXDATE,IDINVC,DOCNO,NEWDOCNO," & "INVNAME,INVAMT,INVTAX,LOCID,VTYPE,RATE," & "TTYPE,ACCTVAT,SOURCE,BATCH,ENTRY,MARK,VATCOMMENT,CBREF,TRANSNBR,RUNNO,TypeofPu,TranNo,CIF,TaxCIF,Verify,TAXID,BRANCH,Code,CURRENCY,EXCHANGRATE,AMTBASETC,TOTTAX,CODETAX)" & " VALUES('" & loINVDATE & "','" & loTXDATE.Trim & "','" & loIDINVC & "','" & loDOCNO & "','" & loNEWDOCNO & "','" & loINVNAME & "'," & loINVAMT & "," & loINVTAX & ",'" & loLOCID & "'," & loVTYPE & "," & loRATE & "," & loTTYPE & ",'" & loACCTVAT & "','" & loSource & "','" & loBatch & "','" & loEntry & "','" & loMark & "','" & loVATCOMMENT & "','" & loCBRef & "','" & loTRANSNBR & "','" & loRunno & "',convert(varchar(30),'" & typeofPU.Replace("'", "''") & "'),'" & TranNo & "','" & CIF & "','" & TaxCIF & "','" & Verify & "','" & TAXID & "','" & BRANCH & "','" & loCode & "','" & loCurrency & "'," & loExchangrate & "," & loAmoutbase & "," & loTotalTax & ",'" & loCODETAX & "')"
            End If
        End If

    End Function

    'บันทึกข้อมูลลง Table เมื่อมีการ Add New VAT หรือ มีการ Edit Vat
    Public Function InsNewVat(ByRef loVATID As Decimal, ByRef loINVDATE As String, ByRef loTXDATE As String, ByRef loIDINVC As String, ByRef loDOCNO As String, ByRef loNEWDOCNO As String, ByRef loINVNAME As String, ByRef loINVAMT As Double, ByRef loINVTAX As Double, ByRef loLOCID As String, ByRef loVTYPE As Short, ByRef loRATE As Double, ByRef loTTYPE As Short, ByRef loACCTVAT As String, ByRef loSource As String, ByRef loBatch As String, ByRef loEntry As String, ByRef loMark As String, ByRef loVATCOMMENT As String, ByRef loCBRef As String, ByRef loStatus As String, ByRef loDETAILNO As String, ByRef loRunno As String, ByRef typeofPU As String, ByRef TranNo As String, ByRef CIF As Decimal, ByRef TaxCIF As Decimal, ByRef loCODETAX As String, ByRef loIDDIST As String, ByRef loVerify As Integer, ByVal TAXID As String, ByVal BRANCH As String, ByVal loCurrency As String, ByVal loExchangrate As Decimal, ByVal loAmountbase As Decimal) As String

        loIDINVC = Trim(ReQuote(limitStr(loIDINVC, 255)))
        loDOCNO = Trim(ReQuote(limitStr(loDOCNO, 60)))
        loNEWDOCNO = Trim(ReQuote(limitStr(loNEWDOCNO, 60)))
        loINVNAME = Trim(ReQuote(limitStr(loINVNAME, 100)))
        loACCTVAT = Trim(ReQuote(limitStr(loACCTVAT, 45)))
        loLOCID = Trim(ReQuote(limitStr(loLOCID, 6)))
        loSource = ReQuote(limitStr(loSource, 2))
        loBatch = Trim(ReQuote(limitStr(loBatch, 20)))
        loEntry = Trim(ReQuote(limitStr(loEntry, 20)))
        loMark = ReQuote(limitStr(loMark, 1))
        loVATCOMMENT = Trim(ReQuote(limitStr(loVATCOMMENT, 250)))
        loCODETAX = ReQuote(limitStr(loCODETAX, 12))
        loIDDIST = Trim(ReQuote(limitStr(loIDDIST, 6)))
        loCBRef = ReQuote(limitStr(loCBRef, 12))
        loDETAILNO = IIf(IsDBNull(loDETAILNO), " ", Trim(loDETAILNO))
        loRunno = ReQuote(Trim(IIf(IsDBNull(loRunno), "", loRunno)))
        BRANCH = ReQuote(Trim(BRANCH))
        loCurrency = ReQuote(Trim(loCurrency))
        'Exchangrate = ReQuote(Trim(Exchangrate))
        'Amountbase = ReQuote(Trim(Amountbase))


        If loDETAILNO = "" Then
            loDETAILNO = " "
        End If

        If ComDatabase = "ORCL" Then
            InsNewVat = InsStr & "FMSVATINSERT(VATID,INVDATE,TXDATE,IDINVC,DOCNO,NEWDOCNO," & "INVNAME,INVAMT,INVTAX,LOCID,VTYPE,RATE," & "TTYPE,ACCTVAT,SOURCE,BATCH,ENTRY,MARK,VATCOMMENT,CBREF,STATUS,TRANSNBR,RUNNO,typeofPU,TranNo,CIF,TAXCIF,Verify,TAXID,BRANCH,CURRENCY,EXCHANGRATE,AMTBASETC)" & " VALUES('" & loVATID & "','" & loINVDATE & "','" & loTXDATE.Trim & "','" & loIDINVC & "','" & loDOCNO & "','" & loNEWDOCNO & "','" & loINVNAME & "'," & loINVAMT & "," & loINVTAX & ",'" & loLOCID & "'," & loVTYPE & "," & loRATE & "," & loTTYPE & ",'" & loACCTVAT & "','" & loSource & "','" & loBatch & "','" & loEntry & "','" & loMark & "','" & loVATCOMMENT & "','" & Trim(loCBRef) & "','" & loStatus & "','" & loDETAILNO & "','" & loRunno & "','" & typeofPU & "','" & TranNo & "','" & CIF & "','" & TaxCIF & "','" & loVerify & "','" & TAXID & "','" & BRANCH & "','" & loCurrency & "'," & loExchangrate & "," & loAmountbase & ")"
        Else
            InsNewVat = InsStr & "FMSVATINSERT(VATID,INVDATE,TXDATE,IDINVC,DOCNO,NEWDOCNO," & "INVNAME,INVAMT,INVTAX,LOCID,VTYPE,RATE," & "TTYPE,ACCTVAT,SOURCE,BATCH,ENTRY,MARK,VATCOMMENT,CBREF,STATUS,TRANSNBR,RUNNO,typeofPU,TranNo,CIF,TAXCIF,Verify,TAXID,BRANCH,CURRENCY,EXCHANGRATE,AMTBASETC)" & " VALUES('" & loVATID & "','" & loINVDATE & "','" & loTXDATE.Trim & "','" & loIDINVC & "','" & loDOCNO & "','" & loNEWDOCNO & "','" & loINVNAME & "'," & loINVAMT & "," & loINVTAX & ",'" & loLOCID & "'," & loVTYPE & "," & loRATE & "," & loTTYPE & ",'" & loACCTVAT & "','" & loSource & "','" & loBatch & "','" & loEntry & "','" & loMark & "','" & loVATCOMMENT & "','" & Trim(loCBRef) & "','" & loStatus & "','" & loDETAILNO & "','" & loRunno & "','" & typeofPU & "','" & TranNo & "','" & CIF & "','" & TaxCIF & "','" & loVerify & "','" & TAXID & "','" & BRANCH & "','" & loCurrency & "'," & loExchangrate & "," & loAmountbase & ")"
        End If

    End Function

    Public Function FindTaxIDBRANCH(ByVal Code As String, ByVal Source As String, ByVal Field As String) As String
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String = ""
        Dim Result As String = ""
        On Error GoTo ErrHandler

        If FindApp(Source) = False Then Return Result
        Code = ReQuote(Code)

        cn = New ADODB.Connection
        cn.ConnectionTimeout = 60
        cn.Open(ConnACCPAC)
        cn.CommandTimeout = 3600

        rs = New BaseClass.BaseODBCIO
        Code = Trim(Code)

        If Source = "CB" Then

            If Field = "TAXID" Then

                If Use_TAX_CB = TAXFROM.EMAIL Then
                    sqlstr = "select EMAILADDR  from CBMISC  where  MISCCODE ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Dim tmResult As String = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                        If tmResult.Length > 13 Then
                            Result = tmResult.Substring(0, 13)
                        Else
                            Result = tmResult
                        End If
                    End If
                ElseIf Use_TAX_CB = TAXFROM.URL Then
                    sqlstr = "select URLADDR  from CBMISC  where MISCCODE ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Dim tmResult As String = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                        If tmResult.Length > 13 Then
                            Result = tmResult.Substring(0, 13)
                        Else
                            Result = tmResult
                        End If
                    End If

                ElseIf Use_TAX_CB = TAXFROM.State Then

                    sqlstr = "select [STATE]  from CBMISC  where MISCCODE ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Dim tmResult As String = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                        If tmResult.Length > 13 Then
                            Result = tmResult.Substring(0, 13)
                        Else
                            Result = tmResult
                        End If
                    End If

                ElseIf Use_TAX_CB = TAXFROM.Country Then

                    sqlstr = "select COUNTRY  from CBMISC  where MISCCODE ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Dim tmResult As String = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                        If tmResult.Length > 13 Then
                            Result = tmResult.Substring(0, 13)
                        Else
                            Result = tmResult
                        End If
                    End If

                ElseIf Use_TAX_CB = TAXFROM.BRN Then
                    sqlstr = "select BRN  from CBMISC  where MISCCODE ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Dim tmResult As String = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                        If tmResult.Length > 13 Then
                            Result = tmResult.Substring(0, 13)
                        Else
                            Result = tmResult
                        End If
                    End If
                End If


            ElseIf Field = "BRANCH" Then

                If Use_BRANCH_CB = TAXFROM.EMAIL Then
                    sqlstr = "select EMAILADDR  from CBMISC  where  MISCCODE ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Dim tmResult As String = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                        Result = tmResult
                    End If
                ElseIf Use_BRANCH_CB = TAXFROM.URL Then
                    sqlstr = "select URLADDR  from CBMISC  where MISCCODE ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Dim tmResult As String = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                        Result = tmResult
                    End If

                ElseIf Use_BRANCH_CB = TAXFROM.State Then
                    sqlstr = "select [STATE]  from CBMISC  where MISCCODE ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Dim tmResult As String = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                        Result = tmResult
                    End If

                ElseIf Use_BRANCH_CB = TAXFROM.Country Then
                    sqlstr = "select COUNTRY  from CBMISC  where MISCCODE ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Dim tmResult As String = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                        Result = tmResult
                    End If

                End If

            End If 'end if CB


        ElseIf Source = "AP" Then

            If Field = "TAXID" Then
                If Use_TAX_AP = TAXFROM.OPF Then
                    sqlstr = "SELECT VALUE  FROM APVENO WHERE OPTFIELD ='TAXID' AND VENDORID ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Result = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                        If Result.Length > 13 Then
                            Result = Result.Substring(0, 13)
                        Else
                            Result = Result
                        End If
                    End If

                ElseIf Use_TAX_AP = TAXFROM.RegisNumber Then
                    sqlstr = "select IDTAXREGI1  from APVEN where VENDORID ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Result = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                        If Result.Length > 13 Then
                            Result = Result.Substring(0, 13)
                        Else
                            Result = Result
                        End If
                    End If

                ElseIf Use_TAX_AP = TAXFROM.EMAIL Then
                    sqlstr = "select EMAIL2  from APVEN where VENDORID ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Dim tmResult As String = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                        If tmResult.Length > 13 Then
                            Result = tmResult.Substring(0, 13)
                        Else
                            Result = tmResult
                        End If
                    End If
                ElseIf Use_TAX_AP = TAXFROM.URL Then
                    sqlstr = "select WEBSITE  from APVEN where VENDORID ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Dim tmResult As String = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                        If tmResult.Length > 13 Then
                            Result = tmResult.Substring(0, 13)
                        Else
                            Result = tmResult
                        End If
                    End If

                ElseIf Use_TAX_AP = TAXFROM.BRN Then
                    sqlstr = "select BRN  from APVEN where VENDORID ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Dim tmResult As String = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                        If tmResult.Length > 13 Then
                            Result = tmResult.Substring(0, 13)
                        Else
                            Result = tmResult
                        End If
                    End If

                End If

            ElseIf Field = "BRANCH" Then

                If Use_BRANCH_AP = TAXFROM.OPF Then
                    sqlstr = "SELECT VALUE  FROM APVENO WHERE OPTFIELD ='BRANCH' AND VENDORID ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Result = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                    End If

                ElseIf Use_BRANCH_AP = TAXFROM.RegisNumber Then
                    sqlstr = "select IDTAXREGI1  from APVEN where VENDORID ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Result = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                    End If

                ElseIf Use_BRANCH_AP = TAXFROM.EMAIL Then
                    sqlstr = "select EMAIL2  from APVEN where VENDORID ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Result = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                    End If

                ElseIf Use_BRANCH_AP = TAXFROM.URL Then
                    sqlstr = "select WEBSITE  from APVEN where VENDORID ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Result = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                    End If

                ElseIf Use_BRANCH_AP = TAXFROM.BRN Then
                    sqlstr = "select BRN  from APVEN where VENDORID ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Result = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                    End If

                End If
            End If 'end if AP

        ElseIf Source = "AR" Then

            If Field = "TAXID" Then
                If Use_TAX_AR = TAXFROM.OPF Then
                    sqlstr = "SELECT VALUE FROM ARCUSO WHERE OPTFIELD ='TAXID' AND IDCUST ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Result = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                        If Result.Length > 13 Then
                            Result = Result.Substring(0, 13)
                        Else
                            Result = Result
                        End If
                    End If

                ElseIf Use_TAX_AR = TAXFROM.RegisNumber Then
                    sqlstr = "select IDTAXREGI1  from ARCUS where IDCUST ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Result = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                        If Result.Length > 13 Then
                            Result = Result.Substring(0, 13)
                        Else
                            Result = Result
                        End If
                    End If

                ElseIf Use_TAX_AR = TAXFROM.EMAIL Then
                    sqlstr = "select EMAIL2  from ARCUS  where IDCUST  ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Dim tmResult As String = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                        If tmResult.Length > 13 Then
                            Result = tmResult.Substring(0, 13)
                        Else
                            Result = tmResult
                        End If
                    End If

                ElseIf Use_TAX_AR = TAXFROM.URL Then
                    sqlstr = "select WEBSITE  from ARCUS  where IDCUST  ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Dim tmResult As String = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                        If tmResult.Length > 13 Then
                            Result = tmResult.Substring(0, 13)
                        Else
                            Result = tmResult
                        End If
                    End If

                ElseIf Use_TAX_AR = TAXFROM.BRN Then
                    sqlstr = "select BRN  from ARCUS  where IDCUST  ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Dim tmResult As String = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                        If tmResult.Length > 13 Then
                            Result = tmResult.Substring(0, 13)
                        Else
                            Result = tmResult
                        End If
                    End If

                End If


            ElseIf Field = "BRANCH" Then

                If Use_BRANCH_AR = TAXFROM.OPF Then
                    sqlstr = "SELECT VALUE FROM ARCUSO WHERE OPTFIELD ='BRANCH' AND IDCUST ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Result = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                    End If
                ElseIf Use_BRANCH_AR = TAXFROM.RegisNumber Then
                    sqlstr = "select IDTAXREGI1  from ARCUS where IDCUST ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Result = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                    End If

                ElseIf Use_BRANCH_AR = TAXFROM.EMAIL Then
                    sqlstr = "select EMAIL2  from ARCUS  where IDCUST  ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Result = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                    End If

                ElseIf Use_BRANCH_AR = TAXFROM.URL Then
                    sqlstr = "select WEBSITE  from ARCUS  where IDCUST  ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Result = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                    End If

                ElseIf Use_BRANCH_AR = TAXFROM.BRN Then
                    sqlstr = "select BRN  from ARCUS  where IDCUST  ='" & Code & "'"
                    rs.Open(sqlstr, cn)
                    If rs.options.QueryDT.Rows.Count > 0 Then
                        Result = rs.options.QueryDT.Rows(0).Item(0).ToString.Trim
                    End If

                End If



            End If 'end if AR

        End If 'end if Source

        cn.Close()
        cn = Nothing
        rs = Nothing

        Return Result

ErrHandler:

        ErrMsg = Err.Description
        WriteLog("<FindTaxIDBRANCH>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Function

    Public Function FindMiscCodeInAPVEN(ByVal aCode As String, ByVal source As String) As String
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String
        'Code = ""

        On Error GoTo ErrHandler
        If FindApp("AP") = False Then Exit Function
        aCode = ReQuote(aCode)
        aCode = Trim(aCode)
        cn = New ADODB.Connection

        cn.ConnectionTimeout = 60
        cn.Open(ConnACCPAC)
        cn.CommandTimeout = 3600

        rs = New BaseClass.BaseODBCIO
        If source = "CB" Then
            If Use_LEGALNAME_inCB Then
                sqlstr = "select APVEN.LEGALNAME ,APVEN.TEXTSTRE1 "
            Else
                sqlstr = "SELECT APVEN.VENDNAME, APVEN.TEXTSTRE1,APVEN.VENDORID"
            End If
        ElseIf source = "AP" Then
            If Use_LEGALNAME_inAP Then
                sqlstr = "select APVEN.LEGALNAME ,APVEN.TEXTSTRE1 "
            Else
                sqlstr = "SELECT APVEN.VENDNAME, APVEN.TEXTSTRE1,APVEN.VENDORID"
            End If
        End If
        sqlstr = sqlstr & " From APVEN"
        sqlstr = sqlstr & " WHERE APVEN.VENDORID ='" & aCode & "'"
        rs.Open(sqlstr, cn)

        Do While rs.options.Rs.EOF = False
            FindMiscCodeInAPVEN = Trim(IIf(IsDBNull(rs.options.Rs.Collect(0)), "", rs.options.Rs.Collect(0)))
            If FindMiscCodeInAPVEN.Trim <> "" Then
                If ComVatName = True And FindMiscCodeInAPVEN.Trim.Substring(Len(FindMiscCodeInAPVEN.Trim) - 1) = "&" Then 'เพิ่ม And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" กรณีที่ ชื้อ Customer ยาว แล้วต้องคีย์เพิ่มที่ช่อง TEXTSTRE1 (เจอที่ AIT)

                    FindMiscCodeInAPVEN = FindMiscCodeInAPVEN.Trim & RTrim(IIf(IsDBNull(rs.options.Rs.Collect(1)), "", Trim(rs.options.Rs.Collect(1))))

                End If
            End If

            If Use_LEGALNAME_inAP = False Then
                Code = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", Trim(rs.options.Rs.Collect(2)))
            End If

            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop

        cn.Close()
        cn = Nothing
        rs = Nothing

        Exit Function

ErrHandler:

        ErrMsg = Err.Description
        WriteLog("<FindMiscCodeInAPVEN>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Function
    Public Function FindMiscCodeInARCUS(ByVal aCode As String) As String
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String
        'Code = ""
        On Error GoTo ErrHandler

        If FindApp("AR") = False Then Exit Function
        aCode = ReQuote(aCode)

        cn = New ADODB.Connection
        cn.ConnectionTimeout = 60
        cn.Open(ConnACCPAC)
        cn.CommandTimeout = 3600

        rs = New BaseClass.BaseODBCIO
        sqlstr = "SELECT ARCUS.NAMECUST, ARCUS.TEXTSTRE1, ARCUS.IDCUST"
        sqlstr = sqlstr & " From ARCUS"
        sqlstr = sqlstr & " WHERE ARCUS.IDCUST = '" & aCode & "'"
        rs.Open(sqlstr, cn)

        Do While rs.options.Rs.EOF = False
            FindMiscCodeInARCUS = Trim(IIf(IsDBNull(rs.options.Rs.Collect(0)), "", rs.options.Rs.Collect(0)))

            If FindMiscCodeInARCUS.Trim <> "" Then

                If ComVatName = True And FindMiscCodeInARCUS.Trim.Substring(Len(FindMiscCodeInARCUS.Trim) - 1) = "&" Then 'เพิ่ม And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" กรณีที่ ชื้อ Customer ยาว แล้วต้องคีย์เพิ่มที่ช่อง TEXTSTRE1 (เจอที่ AIT)
                    FindMiscCodeInARCUS = FindMiscCodeInARCUS.Trim & RTrim(IIf(IsDBNull(rs.options.Rs.Collect(1)), "", Trim(rs.options.Rs.Collect(1))))
                End If
            End If


            Code = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", Trim(rs.options.Rs.Collect(2)))
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop

        cn.Close()
        cn = Nothing
        rs = Nothing

        Exit Function

ErrHandler:
        ErrMsg = Err.Description()
        WriteLog("<FindMiscCodeInARCUS>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Function

    Public Function FindMiscCodeInMisccode(ByVal aCode As String) As String
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String
        'Code = ""

        On Error GoTo ErrHandler

        cn = New ADODB.Connection
        cn.ConnectionTimeout = 60
        cn.Open(ConnACCPAC)
        cn.CommandTimeout = 3600

        rs = New BaseClass.BaseODBCIO
        aCode = ReQuote(aCode)
        sqlstr = "SELECT CBMISC.MISCNAME, CBMISC.ADDRESS1,CBMISC.MISCCODE"
        sqlstr = sqlstr & " From CBMISC"
        sqlstr = sqlstr & " WHERE CBMISC.MISCCODE ='" & aCode & "'"
        rs.Open(sqlstr, cn)

        Do While rs.options.Rs.EOF = False
            FindMiscCodeInMisccode = Trim(IIf(IsDBNull(rs.options.Rs.Collect(0)), "", rs.options.Rs.Collect(0)))
            If FindMiscCodeInMisccode.Trim <> "" Then

                If ComVatName = True And FindMiscCodeInMisccode.Trim.Substring(Len(FindMiscCodeInMisccode.Trim) - 1) = "&" Then 'เพิ่ม And .INVNAME.Trim.Substring(Len(.INVNAME.Trim) - 1) = "&" กรณีที่ ชื้อ Customer ยาว แล้วต้องคีย์เพิ่มที่ช่อง TEXTSTRE1 (เจอที่ AIT)
                    FindMiscCodeInMisccode = FindMiscCodeInMisccode.Trim & RTrim(IIf(IsDBNull(rs.options.Rs.Collect(1)), "", Trim(rs.options.Rs.Collect(1))))
                End If
            End If

            Code = IIf(IsDBNull(rs.options.Rs.Collect(2)), "", Trim(rs.options.Rs.Collect(2)))
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop

        cn.Close()
        cn = Nothing
        rs = Nothing

        Exit Function

ErrHandler:

        ErrMsg = Err.Description
        WriteLog("<FindMiscCodeInMisccode>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Function

    Public Sub RefreshVatView(ByRef ao_Cbo As ComboBox)
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String
        Dim i As Integer

        On Error GoTo ErrHandler

        ao_Cbo.Items.Clear()

        cn = New ADODB.Connection
        cn.Open(ConnVAT)
        rs = New BaseClass.BaseODBCIO
        sqlstr = "SELECT TTYPE,TTYPENAME From FMSTTYPE"
        rs.Open(sqlstr, cn)
        'test = rs
        Do While rs.options.Rs.EOF = False
            ao_Cbo.Items.Add(Trim(rs.options.Rs.Collect(1)))
            rs.options.Rs.MoveNext()
        Loop

        cn.Close()
        cn = Nothing
        If ao_Cbo.Items.Count > 0 Then ao_Cbo.SelectedIndex = 0
        Exit Sub

ErrHandler:
        ErrMsg = Err.Description()
        WriteLog("<RefreshVatView>" & Err.Number & "  " & Err.Description)
        WriteLog("<RefreshVatView> sqlstr : " & sqlstr)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub

    Public Function GetVatType(ByRef aCode As String) As String
        Dim cn As ADODB.Connection
        Dim rs As BaseClass.BaseODBCIO
        Dim sqlstr As String

        On Error GoTo ErrHandler

        cn = New ADODB.Connection
        cn.Open(ConnVAT)
        rs = New BaseClass.BaseODBCIO
        sqlstr = "SELECT TTYPE,TTYPENAME From FMSTTYPE Where TTYPE='" & aCode & "'"
        rs.Open(sqlstr, cn)

        Do While rs.options.Rs.EOF = False
            GetVatType = Trim(rs.options.Rs.Collect(1))
            rs.options.Rs.MoveNext()
            Application.DoEvents()
        Loop

        cn.Close()
        cn = Nothing
        Exit Function

ErrHandler:

        ErrMsg = Err.Description()
        WriteLog("<GetVatType>" & Err.Number & "  " & Err.Description)
        If Len(ErrMsg) > 0 Then MessageBox.Show(ErrMsg, "Error on this data detail", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Function

    Public Sub Connecttion(ByVal Connect As String)

        ConA = New SqlConnection
        ConV = New SqlConnection

        If Connect = "ACC" Then

            If ConA.State = ConnectionState.Open Then
                ConA.Close()
                ConA.ConnectionString = ConAcc
                ConA.Open()

            Else

                Try
                    ConA.ConnectionString = ConAcc
                    ConA.Open()
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try


            End If
        Else

            If ConV.State = ConnectionState.Open Then
                ConV.Close()
                ConV.ConnectionString = ConVAT
                ConV.Open()

            Else

                Try
                    ConV.ConnectionString = ConVAT
                    ConV.Open()
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try


            End If

        End If



    End Sub

    
End Module