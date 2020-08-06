Option Strict Off
Option Explicit On
Friend Class clsVlacc
	
	Private mvarACCID As String 'local copy
	Private mvarVTYPE As String 'local copy
	
	
	Public Property VTYPE() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.VTYPE
			VTYPE = mvarVTYPE
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.VTYPE = 5
			mvarVTYPE = Value
		End Set
	End Property
	
	
	
	
	
	Public Property ACCID() As String
		Get
			'used when retrieving value of a property, on the right side of an assignment.
			'Syntax: Debug.Print X.ACCID
			ACCID = mvarACCID
		End Get
		Set(ByVal Value As String)
			'used when assigning a value to the property, on the left side of an assignment.
			'Syntax: X.ACCID = 5
			mvarACCID = Value
		End Set
	End Property
End Class