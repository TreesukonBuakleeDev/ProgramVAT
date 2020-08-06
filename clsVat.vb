Option Strict Off
Option Explicit On
Friend Class clsVat
	
	'local variable(s) to hold property value(s)
    Private mvarVATID As Decimal  'local copy
	Private mvarINVDATE As Double 'local copy
	Private mvarTXDATE As Double 'local copy
	Private mvarIDINVC As String 'local copy
	Private mvarNEWDOCNO As String 'local copy
	Private mvarDOCNO As String 'local copy
	Private mvarINVNAME As String 'local copy
	Private mvarINVAMT As Double 'local copy
	Private mvarINVTAX As Double 'local copy
	Private mvarLOCID As String 'local copy
	Private mvarVTYPE As String 'local copy
	Private mvarRATE As Double 'local copy
    Private mvarTTYPE As Integer 'local copy
	Private mvarACCTVAT As String 'local copy
	Private mvarSOURCE As String 'local copy
    Private mvarBATCH As String 'local copy
    Private mvarRunno As String 'local copy
	Private mvarENTRY As String 'local copy
	Private mvarMARK As String 'local copy
	Private mvarVATCOMMENT As String 'local copy
	Private mvarCBREF As String 'local copy
	Private mvarTRANSNBR As String
    Private mvarTypeOfPU As String
    Private mvarTranNo As String
    Private mvarCIF As Decimal
    Private mvarTaxCIF As Decimal
    Private mvarVerify As Integer
    Private mvarTAXID As String
    Private mvarBranch As String
    Private mvarCode As String
    Private mvarTOTTAX As Double
    Private mvarCODETAX As String

    Public Property Verify() As Integer
        Get
            Verify = mvarVerify
        End Get
        Set(ByVal value As Integer)
            mvarVerify = value
        End Set
    End Property
    Public Property TypeOfPU() As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.VATCOMMENT
            TypeOfPU = mvarTypeOfPU
        End Get
        Set(ByVal Value As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.VATCOMMENT = 5
            mvarTypeOfPU = Value
        End Set
    End Property

    Public Property TranNo() As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.VATCOMMENT
            TranNo = mvarTranNo
        End Get
        Set(ByVal Value As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.VATCOMMENT = 5
            mvarTranNo = Value
        End Set
    End Property

    Public Property CIF() As Decimal
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.VATCOMMENT
            CIF = mvarCIF
        End Get
        Set(ByVal Value As Decimal)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.VATCOMMENT = 5
            mvarCIF = Value
        End Set
    End Property

    Public Property TaxCIF() As Decimal
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.VATCOMMENT
            TaxCIF = mvarTaxCIF
        End Get
        Set(ByVal Value As Decimal)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.VATCOMMENT = 5
            mvarTaxCIF = Value
        End Set
    End Property

	Public Property CBRef() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.VATCOMMENT
			CBRef = mvarCBREF
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.VATCOMMENT = 5
			mvarCBREF = Value
		End Set
	End Property
	
	Public Property Entry() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.VATCOMMENT
			Entry = mvarENTRY
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.VATCOMMENT = 5
            mvarENTRY = Value
		End Set
	End Property
	
	Public Property Batch() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.VATCOMMENT
			Batch = mvarBATCH
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
               'Syntax: X.VATCOMMENT = 5
            mvarBATCH = Value
		End Set
	End Property
	
	Public Property VATCOMMENT() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.VATCOMMENT
			VATCOMMENT = mvarVATCOMMENT
        End Get

		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.VATCOMMENT = 5
			mvarVATCOMMENT = Value
		End Set
    End Property

 Public Property Runno() As String
     Get
        Return mvarRunno
     End Get
     Set(ByVal value As String)
       mvarRunno = value
     End Set
 End Property

	Public Property MARK() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.MustCHECK
			MARK = mvarMARK
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.MustCHECK = 5
			mvarMARK = Value
		End Set
	End Property

	Public Property Source() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.SOURCE
			Source = mvarSOURCE
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.SOURCE = 5
			mvarSOURCE = Value
		End Set
	End Property
		
	Public Property ACCTVAT() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.ACCTVAT
			ACCTVAT = mvarACCTVAT
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.ACCTVAT = 5
			mvarACCTVAT = Value
		End Set
	End Property
    Public Property TTYPE() As Integer
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.TTYPE
            TTYPE = mvarTTYPE
        End Get
        Set(ByVal Value As Integer)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.TTYPE = 5
            mvarTTYPE = Value
        End Set
    End Property
	
	'UPGRADE_NOTE: RATE was upgraded to RATE_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Property RATE_Renamed() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.RATE
			RATE_Renamed = mvarRATE
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.RATE = 5
			mvarRATE = Value
		End Set
	End Property
		
    Public Property VTYPE() As Integer
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.VTYPE
            VTYPE = mvarVTYPE
        End Get
        Set(ByVal Value As Integer)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.VTYPE = 5
            mvarVTYPE = Value
        End Set
    End Property

    Public Property LOCID() As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.LOCID
            LOCID = mvarLOCID
        End Get
        Set(ByVal Value As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.LOCID = 5
            mvarLOCID = Value
        End Set
    End Property
	
	Public Property INVTAX() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.INVTAX
			INVTAX = mvarINVTAX
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.INVTAX = 5
			mvarINVTAX = Value
		End Set
	End Property

	Public Property INVAMT() As Double
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.INVAMT
			INVAMT = mvarINVAMT
		End Get
		Set(ByVal Value As Double)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.INVAMT = 5
			mvarINVAMT = Value
		End Set
	End Property
	
    Public Property INVNAME() As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.INVNAME
            INVNAME = mvarINVNAME
        End Get
        Set(ByVal Value As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.INVNAME = 5
            mvarINVNAME = Value
        End Set
    End Property

    Public Property DOCNO() As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.DOCNO
            DOCNO = mvarDOCNO
        End Get
        Set(ByVal Value As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.DOCNO = 5
            mvarDOCNO = Value
        End Set
    End Property
    Public Property NEWDOCNO() As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.NEWDOCNO
            NEWDOCNO = mvarNEWDOCNO
        End Get
        Set(ByVal Value As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.NEWDOCNO = 5
            mvarNEWDOCNO = Value
        End Set
    End Property
    Public Property IDINVC() As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.IDINVC
            IDINVC = mvarIDINVC
        End Get
        Set(ByVal Value As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.IDINVC = 5
            mvarIDINVC = Value
        End Set
    End Property
    Public Property TXDATE() As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.TXDATE
            TXDATE = mvarTXDATE
        End Get
        Set(ByVal Value As Double)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.TXDATE = 5
            mvarTXDATE = Value
        End Set
    End Property
    Public Property INVDATE() As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.INVDATE
            INVDATE = mvarINVDATE
        End Get
        Set(ByVal Value As Double)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.INVDATE = 5
            mvarINVDATE = Value
        End Set
    End Property
    Public Property VATID() As Decimal
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.VATID
            VATID = mvarVATID
        End Get
        Set(ByVal Value As Decimal)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.VATID = 5
            mvarVATID = Value
        End Set
    End Property
    Public Property TRANSNBR() As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.VATID
            TRANSNBR = mvarTRANSNBR
        End Get
        Set(ByVal Value As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.VATID = 5
            mvarTRANSNBR = Value
        End Set
    End Property

    Public Property TAXID() As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.VATID
            TAXID = mvarTAXID
        End Get
        Set(ByVal Value As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.VATID = 5
            mvarTAXID = Value
        End Set
    End Property

    Public Property Branch() As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.VATID
            Branch = mvarBranch
        End Get
        Set(ByVal Value As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.VATID = 5
            mvarBranch = Value
        End Set
    End Property

    Public Property Code1() As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.VATID
            Code1 = mvarCode
        End Get
        Set(ByVal Value As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.VATID = 5
            mvarCode = Value
        End Set
    End Property

    Public Property TOTTAX() As Double
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.VATID
            TOTTAX = mvarTOTTAX
        End Get
        Set(ByVal Value As Double)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.VATID = 5
            mvarTOTTAX = Value
        End Set
    End Property

    Public Property CODETAX() As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.VATID
            CODETAX = mvarCODETAX
        End Get
        Set(ByVal Value As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.VATID = 5
            mvarCODETAX = Value
        End Set
    End Property

End Class