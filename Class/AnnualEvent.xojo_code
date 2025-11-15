#tag Interface
Protected Interface AnnualEvent
	#tag Method, Flags = &h0
		Function Calculate(Year as integer) As DateAndCaption
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Caption() As string
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Caption(assigns s as string)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CycleFirstYear() As Integer
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CycleFirstYear(assigns value as Integer)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CycleYearDuration() As Integer
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CycleYearDuration(assigns value as Integer)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function DateValue(Year as integer) As DateTime
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function DayOff() As Boolean
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub DayOff(assigns value as Boolean)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function DbRow(RegionIdentifier as Variant = Nil, encoding as TextEncoding = nil) As DatabaseRow
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function DefinitionObject() As Variant
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function EndOfValidity() As DateTime
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub EndOfValidity(assigns d as DateTime)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FingerPrint() As string
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function StartOfValidity() As DateTime
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub StartOfValidity(assigns d as DateTime)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Tag() As Variant
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Tag(assigns value as Variant)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TestDate(d as DateTime) As Boolean
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ToString() As String
		  
		End Function
	#tag EndMethod


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Interface
#tag EndInterface
