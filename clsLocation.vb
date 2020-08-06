Option Strict Off
Option Explicit On
Friend Class clsLocation
	
	'local variable(s) to hold property value(s)
	Private mvarLOCID As String 'local copy
	Private mvarLOCNAME As String 'local copy
	Private mvarLOCADDR As String 'local copy
	Private mvarLOCDESC As String 'local copy
	Private mvarLOCPREFIX As String 'local copy
	Private mvarcolVLACC As colVlacc

	Public Property colVLACC() As colVlacc
		Get
            If mvarcolVLACC Is Nothing Then
                mvarcolVLACC = New colVlacc
            End If

			colVLACC = mvarcolVLACC
		End Get
		Set(ByVal Value As colVlacc)
			mvarcolVLACC = Value
		End Set
	End Property
	
	
	Public Property LOCPREFIX() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.LOCDESC
			LOCPREFIX = mvarLOCPREFIX
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.LOCDESC = 5
			mvarLOCPREFIX = Value
		End Set
	End Property
	
    Public Property LOCDESC() As String
        Get
            'used when retrieving value of a property, on the right side of an assignment.
            'Syntax: Debug.Print X.LOCDESC
            LOCDESC = mvarLOCDESC
        End Get
        Set(ByVal Value As String)
            'used when assigning a value to the property, on the left side of an assignment.
            'Syntax: X.LOCDESC = 5
            mvarLOCDESC = Value
        End Set
    End Property

	Public Property LOCADDR() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.LOCADDR
			LOCADDR = mvarLOCADDR
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.LOCADDR = 5
			mvarLOCADDR = Value
		End Set
	End Property

	Public Property LOCNAME() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.LOCNAME
			LOCNAME = mvarLOCNAME
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.LOCNAME = 5
			mvarLOCNAME = Value
		End Set
	End Property

	Public Property LOCID() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.LocId
			LOCID = mvarLOCID
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.LocId = 5
			mvarLOCID = Value
		End Set
	End Property
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object mvarcolVLACC may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mvarcolVLACC = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class