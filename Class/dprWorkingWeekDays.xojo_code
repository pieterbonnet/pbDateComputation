#tag Class
Protected Class dprWorkingWeekDays
	#tag Method, Flags = &h0
		Function WorkingDay(DayOfWeek as Integer) As Boolean
		  Select Case DayOfWeek 
		    
		  Case 1
		    
		    Return Me.WorkingSunday
		    
		  Case 2
		    
		    Return Me.WorkingMonday
		    
		  Case 3
		    
		    Return Me.WorkingTuesday
		    
		  Case 4
		    
		    Return Me.WorkingWednesday
		    
		  Case 5
		    
		    Return Me.WorkingThursday
		    
		  Case 6
		    
		    Return Me.WorkingFriday
		    
		  Case 7
		    
		    Return Me.WorkingSaturday
		    
		  Case Else
		    
		    Raise new InvalidArgumentException
		    
		  end
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub WorkingDay(DayOfWeek as Integer, assigns value as Boolean)
		  Select Case DayOfWeek 
		    
		  Case 1
		    
		    Me.WorkingSunday = value
		    
		  Case 2
		    
		    Me.WorkingMonday = value
		    
		  Case 3
		    
		    Me.WorkingTuesday = value
		    
		  Case 4
		    
		    Me.WorkingWednesday = value
		    
		  Case 5
		    
		    Me.WorkingThursday = value
		    
		  Case 6
		    
		    Me.WorkingFriday = value
		    
		  Case 7
		    
		    Me.WorkingSaturday = value
		    
		  Case Else
		    
		    Raise new InvalidArgumentException
		    
		  end
		End Sub
	#tag EndMethod


	#tag Note, Name = This class is required for the RegionDatesWorked class
		
		
	#tag EndNote


	#tag Property, Flags = &h0
		WorkingFriday As Boolean = True
	#tag EndProperty

	#tag Property, Flags = &h0
		WorkingMonday As Boolean = True
	#tag EndProperty

	#tag Property, Flags = &h0
		WorkingSaturday As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		WorkingSunday As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		WorkingThursday As Boolean = True
	#tag EndProperty

	#tag Property, Flags = &h0
		WorkingTuesday As Boolean = True
	#tag EndProperty

	#tag Property, Flags = &h0
		WorkingWednesday As Boolean = True
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
			Name="WorkingMonday"
			Visible=false
			Group="Behavior"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="WorkingTuesday"
			Visible=false
			Group="Behavior"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="WorkingWednesday"
			Visible=false
			Group="Behavior"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="WorkingSaturday"
			Visible=false
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="WorkingSunday"
			Visible=false
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="WorkingThursday"
			Visible=false
			Group="Behavior"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="WorkingFriday"
			Visible=false
			Group="Behavior"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
