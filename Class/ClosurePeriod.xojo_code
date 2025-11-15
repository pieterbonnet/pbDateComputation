#tag Class
Protected Class ClosurePeriod
	#tag Method, Flags = &h0
		Sub Constructor(copy as ClosurePeriod)
		  Me.mFirstDay = copy.FirstDay
		  Me.mLastDay = copy.LastDay
		  Me.Caption = copy.Caption
		  me.Tag = copy.Tag
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(lFirstDay as DateTime = Nil, lLastDay as DateTime = Nil, lCaption as string = "")
		  Me.FirstDay = lFirstDay
		  Me.LastDay= lLastDay
		  me.Caption = lCaption
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(FromString as string)
		  Var s() As String = FromString.Split("|")
		  If s.LastIndex <> 2 Then 
		    Raise New InvalidArgumentException
		    Exit Sub
		  end
		  
		  If s(0).Trim <> "" Then Me.Caption = DecodeBase64(s(0))
		  Me.FirstDay = DateTime.FromString(s(1))
		  Me.LastDay = DateTime.FromString(s(2))
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Duration() As Integer
		  Var interval As DateInterval = Me.LastDay - Me.FirstDay
		  Return interval.Days + 1
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function FromString(StringDefinition as String) As ClosurePeriod
		  Var s() As String = StringDefinition.Split("|")
		  If s.LastIndex <> 2 Then 
		    Raise new InvalidArgumentException
		    Return Nil
		  end
		  Var cp As New ClosurePeriod
		  
		  If s(0).Trim <> "" Then cp.Caption = DecodeBase64(s(0))
		  cp.FirstDay = DateTime.FromString(s(1))
		  cp.LastDay = DateTime.FromString(s(2))
		  
		  Return cp
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsIncluded(lDate as DateTime) As Boolean
		  Return (lDate >= me.FirstDay and lDate <= me.LastDay)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ToString() As String
		  Var s as string
		  
		  s = s + EncodeBase64(Me.Caption,0) 
		  s = s + Me.FirstDay.ToString("yyyy-MM-dd") + "|"
		  s = s + Me.LastDay.ToString("yyyy-MM-dd") 
		  
		  Return s
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		Caption As String
	#tag EndProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return mFirstDay
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  Var d1 as new DateTime(value.Year, value.Month,  value.day, 0, 0, 0)
			  mFirstDay = d1
			End Set
		#tag EndSetter
		FirstDay As DateTime
	#tag EndComputedProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  Return mLastDay
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  var d1 as new DateTime(value.Year, value.Month,  value.day, 23, 59, 59)
			  mLastDay = value
			End Set
		#tag EndSetter
		LastDay As DateTime
	#tag EndComputedProperty

	#tag Property, Flags = &h21
		Private mFirstDay As DateTime
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mLastDay As DateTime
	#tag EndProperty

	#tag Property, Flags = &h0
		Tag As Variant
	#tag EndProperty


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
		#tag ViewProperty
			Name="Caption"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
