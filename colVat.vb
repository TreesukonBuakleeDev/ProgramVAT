Option Strict Off
Option Explicit On
Friend Class colVat
	Implements System.Collections.IEnumerable
	
	'local variable to hold collection
	Private mCol As Collection
	
	'UPGRADE_NOTE: RATE was upgraded to RATE_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public Function Add(ByRef VATID As Decimal, ByRef INVDATE As Double, ByRef TXDATE As Double, ByRef IDINVC As String, ByRef NEWDOCNO As String, ByRef DOCNO As String, ByRef INVNAME As String, ByRef INVAMT As Double, ByRef INVTAX As Double, ByRef LOCID As String, ByRef VTYPE As String, ByRef RATE_Renamed As Double, ByRef TAXID As String, ByRef BRANCH As String, ByRef TTYPE As Short, ByRef ACCTVAT As String, ByRef Source As String, ByRef Batch As String, ByRef Entry As String, ByRef MARK As String, ByRef VATCOMMENT As String, ByRef CBRef As String, ByRef TRANSNBR As String, ByRef Runno As String, ByRef TypeOfPu As String, ByRef TranNO As String, ByRef CIF As Decimal, ByRef TaxCIF As Decimal, ByRef Verify As Boolean, Optional ByRef sKey As String = "") As clsVat

        'create a new object
        Dim objNewMember As clsVat
        objNewMember = New clsVat

        'set the properties passed into the method
        objNewMember.VATID = VATID
        objNewMember.INVDATE = INVDATE
        objNewMember.TXDATE = TXDATE
        objNewMember.IDINVC = IDINVC
        objNewMember.NEWDOCNO = NEWDOCNO
        objNewMember.DOCNO = DOCNO
        objNewMember.INVNAME = INVNAME
        objNewMember.INVAMT = INVAMT
        objNewMember.INVTAX = INVTAX
        objNewMember.LOCID = LOCID
        objNewMember.VTYPE = CShort(VTYPE)
        objNewMember.RATE_Renamed = RATE_Renamed
        objNewMember.TTYPE = TTYPE
        objNewMember.ACCTVAT = ACCTVAT
        objNewMember.Source = Source
        objNewMember.Batch = Batch
        objNewMember.Entry = Entry
        objNewMember.MARK = MARK
        objNewMember.CBRef = CBRef
        objNewMember.Runno = Runno
        objNewMember.VATCOMMENT = VATCOMMENT
        objNewMember.TRANSNBR = TRANSNBR
        objNewMember.TypeOfPU = TypeOfPu
        objNewMember.TranNo = TranNO
        objNewMember.CIF = CIF
        objNewMember.TaxCIF = TaxCIF
        objNewMember.Verify = Verify
        objNewMember.TAXID = TAXID
        objNewMember.Branch = BRANCH
        ' objNewMember.Code = Code

        If Len(sKey) = 0 Then
            mCol.Add(objNewMember)
        Else
            mCol.Add(objNewMember, sKey)
        End If

        'return the object created
        Add = objNewMember
        'UPGRADE_NOTE: Object objNewMember may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objNewMember = Nothing
    End Function

    Public Sub Add(ByVal TmpCls As clsVat, Optional ByRef sKey As String = "")
        If Len(sKey) = 0 Then
            mCol.Add(TmpCls)
        Else
            mCol.Add(TmpCls, sKey)
        End If

        TmpCls = Nothing
    End Sub

    Default Public ReadOnly Property Item(ByVal vntIndexKey As Integer) As clsVat
        Get
            'used when referencing an element in the collection
            'vntIndexKey contains either the Index or Key to the collection,
            'this is why it is declared as a Variant
            'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
            Item = mCol.Item(vntIndexKey)
        End Get
    End Property
	
	Public ReadOnly Property Count() As Integer
		Get
			'used when retrieving the number of elements in the
			'collection. Syntax: Debug.Print x.Count
			Count = mCol.Count()
		End Get
	End Property
	
	
	'UPGRADE_NOTE: NewEnum property was commented out. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3FC1610-34F3-43F5-86B7-16C984F0E88E"'
	'Public ReadOnly Property NewEnum() As stdole.IUnknown
		'Get
			'this property allows you to enumerate
			'this collection with the For...Each syntax
			'NewEnum = mCol._NewEnum
		'End Get
	'End Property
	
	Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
		'UPGRADE_TODO: Uncomment and change the following line to return the collection enumerator. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="95F9AAD0-1319-4921-95F0-B9D3C4FF7F1C"'
        GetEnumerator = mCol.GetEnumerator
	End Function
	
	
	Public Sub Remove(ByRef vntIndexKey As Object)
		'used when removing an element from the collection
		'vntIndexKey contains either the Index or Key, which is why
		'it is declared as a Variant
		'Syntax: x.Remove(xyz)
		mCol.Remove(vntIndexKey)
	End Sub
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'creates the collection when this class is created
		mCol = New Collection
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'destroys collection when this class is terminated
		'UPGRADE_NOTE: Object mCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mCol = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class