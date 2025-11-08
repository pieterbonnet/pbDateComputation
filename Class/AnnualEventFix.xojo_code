#tag Class
Protected Class AnnualEventFix
Implements AnnualEvent
	#tag Method, Flags = &h0
		Function Calculate(year as Integer) As DateAndCaption
		  If year < 1 Then 
		    Break
		    Return Nil
		  End
		  
		  If Me.CycleYearDuration > 1 Then
		    if me.CycleFirstYear > year then Return Nil
		    if (year - me.CycleFirstYear) mod Me.CycleYearDuration > 0 then Return Nil
		  end
		  
		  If year < Me.StartOfValidity.Year Or year > Me.EndOfValidity.Year Then Return Nil
		  
		  If Day < 1 Or Month < 1 Or Month > 12 or day > 31 Then Return Nil
		  
		  Var d As DateTime
		  
		  #Pragma BreakOnExceptions False
		  Try
		    d = new DateTime(year, Me.Month, Me.Day, 0, 0, 0, 0)
		  Catch
		    Return Nil
		  end
		  #Pragma BreakOnExceptions True
		  
		  If d.DayOfWeek = 7 And Me.FridayIfSaturday Then 
		    d = d.SubtractInterval(0,0,1)
		    If d < Me.StartOfValidity Or d > Me.EndOfValidity Then Return Nil
		    Return new DateAndCaption(d, Me.mCaption)
		  end
		  
		  If d.DayOfWeek = 7 And Me.MondayIfSaturday Then 
		    d = d.AddInterval(0,0,2)
		    If d < Me.StartOfValidity Or d > Me.EndOfValidity Then Return Nil
		    Return New DateAndCaption(d, Me.mCaption)
		  end
		  
		  If d.DayOfWeek = 1 And Me.MondayIfSunday Then 
		    d = d.AddInterval(0,0,1)
		    If d < Me.StartOfValidity Or d > Me.EndOfValidity Then Return Nil
		    Return New DateAndCaption(d, Me.mCaption)
		  end
		  
		  if AlwaysPreviousWeekDay > 0 then
		    
		    Do 
		      d = d.SubtractInterval(0,0,1)
		    Loop Until d.DayOfWeek = AlwaysPreviousWeekDay
		    
		  elseif AlwaysNextWeekDay > 0 then
		    
		    Do 
		      d = d.AddInterval(0,0,1)
		    Loop Until d.DayOfWeek = AlwaysNextWeekDay
		    
		  elseif PreviousWeekDay > 0 then
		    
		    Do Until d.DayOfWeek = PreviousWeekDay
		      d = d.SubtractInterval(0,0,1)
		    Loop
		    
		  elseIf NextWeekDay > 0 Then
		    
		    Do Until d.DayOfWeek = NextWeekDay 
		      d = d.AddInterval(0,0,1)
		    Loop
		    
		  ElseIf AddDays <> 0 Then
		    d = d.AddInterval(0,0, AddDays)
		  End
		  
		  If d < Me.StartOfValidity Or d > Me.EndOfValidity Then Return Nil
		  
		  Return New DateAndCaption(d, Me.mCaption)
		  
		  Return Nil
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Caption() As string
		  Return mCaption
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Caption(assigns s as string)
		  me.mCaption = s
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(a as AnnualEvent)
		  If a = Nil Then 
		    Raise new NilObjectException
		    Exit Sub
		  end
		  
		  If not (a IsA AnnualEventFix) Then
		    Raise New InvalidArgumentException
		    Exit Sub
		  end
		  
		  Var vo As AnnualEventFix = a.DefinitionObject
		  
		  Me.mCaption = vo.Caption
		  Me.Day = vo.Day
		  Me.Month = vo.Month
		  
		  Me.DayOff = vo.DayOff
		  
		  Me.CycleFirstYear = vo.CycleFirstYear
		  Me.CycleYearDuration = vo.CycleYearDuration
		  
		  Me.StartOfValidity = vo.StartOfValidity
		  Me.EndOfValidity = vo.EndOfValidity
		  
		  Me.AddDays = vo.AddDays
		  Me.FridayIfSaturday = vo.FridayIfSaturday
		  Me.MondayIfSunday = vo.MondayIfSunday
		  Me.MondayIfSaturday = vo.MondayIfSaturday
		  Me.NextWeekDay = vo.NextWeekDay
		  Me.PreviousWeekDay = vo.PreviousWeekDay
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(Copy as AnnualEventFix)
		  Me.mCaption = Copy.Caption
		  Me.Day = Copy.Day
		  Me.Month = copy.Month
		  
		  me.DayOff = Copy.DayOff
		  
		  me.CycleFirstYear = Copy.CycleFirstYear
		  me.CycleYearDuration = Copy.CycleYearDuration
		  
		  Me.StartOfValidity = copy.StartOfValidity
		  Me.EndOfValidity = copy.EndOfValidity
		  
		  Me.AddDays = copy.AddDays
		  Me.FridayIfSaturday = copy.FridayIfSaturday
		  Me.MondayIfSunday = copy.MondayIfSunday
		  me.MondayIfSaturday = Copy.MondayIfSaturday
		  Me.NextWeekDay = copy.NextWeekDay
		  me.PreviousWeekDay = copy.PreviousWeekDay
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(lCaption as string = "", lMonth as Integer = 1, lDay as Integer = 1, lFridayIfSaturday as Boolean = False, lMondayIfSunday as Boolean = False)
		  Me.mCaption = lCaption
		  Me.Day = lDay
		  Me.Month = lMonth
		  
		  
		  Me.StartOfValidity = New DateTime(1,1,1,0,0,0)
		  Me.EndOfValidity = New DateTime(3999,12,31,23,59,59)
		  Me.FridayIfSaturday = lFridayIfSaturday
		  me.MondayIfSunday = lMondayIfSunday
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CycleFirstYear() As Integer
		  Return mCycleFirstYear
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CycleFirstYear(assigns value as Integer)
		  mCycleFirstYear = value
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CycleYearDuration() As Integer
		  Return mCycleYearDuration
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CycleYearDuration(assigns value as Integer)
		  mCycleYearDuration = value
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function DateValue(year as Integer) As DateTime
		  Var dc As DateAndCaption = Me.Calculate(year)
		  If dc = Nil Then Return Nil
		  Return dc.DateValue
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function DayOff() As Boolean
		  Return mDayOff
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub DayOff(assigns value as Boolean)
		  mDayOff = value
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function DbRow(RegionIdentifier as Variant = Nil, encoding as TextEncoding = nil) As DatabaseRow
		  Var row as new DatabaseRow
		  Var dft as string = "F"
		  
		  
		  If Encoding = Nil Or Encoding = Encodings.UTF8 Then
		    
		    row.Column("caption").StringValue = Me.mCaption.DefineEncoding(Encodings.UTF8)
		    if RegionIdentifier <> nil then row.Column("region").StringValue = RegionIdentifier.StringValue.DefineEncoding(Encodings.UTF8) else row.Column("region").StringValue = ""
		    row.Column("fingerprint").StringValue = me.FingerPrint.DefineEncoding(Encodings.UTF8)
		    row.Column("definitiontype").StringValue = dft.DefineEncoding(Encodings.UTF8)
		    
		  Else
		    
		    row.Column("caption").StringValue = Me.mCaption.DefineEncoding(Encodings.UTF8).ConvertEncoding(Encoding).DefineEncoding(encoding)
		    if RegionIdentifier <> nil then row.Column("region").StringValue = RegionIdentifier.StringValue.DefineEncoding(Encodings.UTF8).ConvertEncoding(Encoding).DefineEncoding(encoding) else row.Column("region").StringValue = ""
		    row.Column("fingerprint").StringValue = me.FingerPrint.DefineEncoding(Encodings.UTF8).ConvertEncoding(Encoding).DefineEncoding(encoding)
		    row.Column("definitiontype").StringValue = dft.DefineEncoding(Encodings.UTF8).ConvertEncoding(Encoding).DefineEncoding(encoding)
		    
		  End
		  
		  
		  row.Column("cyclefirstyear").IntegerValue = me.mCycleFirstYear
		  row.Column("cycleyearduration").IntegerValue = me.mCycleYearDuration 
		  row.Column("start").DateTimeValue = me.mStartOfValidity
		  row.Column("end").DateTimeValue = me.mEndOfValidity
		  row.Column("dayoff").BooleanValue = me.mDayOff
		  
		  row.Column("month").IntegerValue = Me.Month
		  row.Column("day").IntegerValue = Me.Day
		  row.Column("sundaytomonday").BooleanValue = me.MondayIfSunday
		  row.Column("saturdaytofriday").BooleanValue = Me.FridayIfSaturday
		  row.Column("saturdaytomonday").BooleanValue = me.MondayIfSaturday
		  
		  row.Column("adddays").IntegerValue = me.AddDays
		  if me.AlwaysNextWeekDay > 0 or me.AlwaysPreviousWeekDay > 0 then
		    row.Column("nextweekday").IntegerValue = Me.AlwaysNextWeekDay
		    row.Column("previousweekday").IntegerValue = Me.AlwaysPreviousWeekDay
		    row.Column("alwaysshift").BooleanValue = True
		  else
		    row.Column("nextweekday").IntegerValue = Me.NextWeekDay
		    row.Column("previousweekday").IntegerValue = Me.PreviousWeekDay
		    row.Column("alwaysshift").BooleanValue = False
		  end
		  
		  Return row
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function DefinitionObject() As Variant
		  Return me
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function EndOfValidity() As DateTime
		  Return mEndOfValidity
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub EndOfValidity(assigns d as DateTime)
		  Var d1 as new DateTime(d.Year, d.Month, d.day, 23, 59, 59)
		  mEndOfValidity = d
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FingerPrint() As string
		  Return "F" _
		  + Format(Me.Month, "0#") _
		  + ";" + Format(Me.Day, "0#") _
		  + ";" + me.AddDays.ToString _
		  + "|" + Me.mStartOfValidity.SQLDate _
		  + "|" + Me.mEndOfValidity.SQLDate 
		  
		  'Return "F" _
		  '+ Format(Me.Month, "0#") _
		  '+ Format(Me.Day, "0#") _
		  '+ If(Me.FridayIfSaturday, "1", "0") _
		  '+ If(Me.MondayIfSaturday, "1", "0") _
		  '+ If(Me.MondayIfSunday, "1", "0") _
		  '+ Format(Me.AddDays, "0#") _
		  '+ Me.NextWeekDay.ToString _
		  '+ Me.mStartOfValidity.SQLDate _
		  '+ Me.mEndOfValidity.SQLDate 
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function StartOfValidity() As DateTime
		  Return mStartOfValidity
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub StartOfValidity(assigns d as DateTime)
		  Var d1 as new DateTime(d.Year, d.Month, d.day, 0, 0, 0)
		  mStartOfValidity = d1
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Tag() As Variant
		  Return mTag
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Tag(assigns value as Variant)
		  me.mTag = value
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TestDate(d as DateTime) As Boolean
		  // Simple case
		  
		  If Me.CycleYearDuration > 1 Then
		    if me.CycleFirstYear > d.year then Return False
		    If (d.year - Me.CycleFirstYear) Mod Me.CycleYearDuration > 0 Then Return false
		  end
		  
		  If d < Me.StartOfValidity Or d > Me.EndOfValidity Then Return False
		  
		  If Me.FridayIfSaturday And d.DayOfWeek = 6 Then
		    Var dt As DateTime = d.AddInterval(0,0,1)
		    If dt.Day = Me.Day And dt.Month = Me.Month Then Return True
		  end
		  
		  If Me.MondayIfSunday And d.DayOfWeek = 2 Then
		    Var dt As DateTime = d.SubtractInterval(0,0,1)
		    If dt.Day = Me.Day And dt.Month = Me.Month Then Return True
		  End
		  
		  If Me.MondayIfSaturday And d.DayOfWeek = 2 Then
		    Var dt As DateTime = d.SubtractInterval(0,0,2)
		    If dt.Day = Me.Day And dt.Month = Me.Month Then Return True
		  End
		  
		  if not FridayIfSaturday and not MondayIfSunday and not MondayIfSaturday then
		    
		    if me.AlwaysPreviousWeekDay > 0 then
		      
		      If d.DayOfWeek <> Me.AlwaysPreviousWeekDay Then Return False // No need to continue
		      
		      Var dt As New DateTime(d) // We move the days forward
		      
		      For x As Integer = 1 To 7
		        dt = d.AddInterval(0,0,x)
		        If Me.Day = dt.Day And Me.Month = dt.Month Then Return True 
		      Next
		      
		      Return False
		      
		    ElseIf Me.AlwaysNextWeekDay > 0 Then 
		      
		      If d.DayOfWeek <> Me.AlwaysNextWeekDay Then Return False // No need to continue
		      
		      Var dt As New DateTime(d) // We go back in time
		      
		      For x As Integer = 1 To 7
		        dt = d.SubtractInterval(0,0,x)
		        If Me.Day = dt.Day And Me.Month = dt.Month Then Return True 
		      Next
		      
		      Return False
		      
		    elseif me.PreviousWeekDay > 0 then
		      
		      If d.DayOfWeek <> Me.PreviousWeekDay Then Return False // No need to continue
		      
		      Var dt As New DateTime(d) // We move the days forward
		      
		      For x As Integer = 0 To 6
		        dt = d.AddInterval(0,0,x)
		        If Me.Day = dt.Day And Me.Month = dt.Month Then Return True 
		      Next
		      
		      Return False
		      
		    elseIf Me.NextWeekDay > 0 Then 
		      
		      If d.DayOfWeek <> Me.NextWeekDay Then Return false // No need to continue
		      
		      Var dt As new DateTime(d) // We go back in time
		      
		      For x As Integer = 0 To 6
		        dt = d.SubtractInterval(0,0,x)
		        If Me.Day = dt.Day And Me.Month = dt.Month Then Return True 
		      Next
		      
		      Return False
		      
		    ElseIf Me.AddDays <> 0 Then
		      
		      Var dt As DateTime = d.SubtractInterval(0,0, Me.AddDays)
		      If Me.Day = dt.Day And Me.Month = dt.Month Then Return True else return False
		      
		    End
		    
		  end
		  
		  If Me.Day = d.Day And Me.Month = d.Month Then Return True 
		  
		  Return false
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		AddDays As Integer = 0
	#tag EndProperty

	#tag Property, Flags = &h0
		AlwaysNextWeekDay As Integer = 0
	#tag EndProperty

	#tag Property, Flags = &h0
		AlwaysPreviousWeekDay As Integer = 0
	#tag EndProperty

	#tag Property, Flags = &h0
		Day As Integer = 1
	#tag EndProperty

	#tag Property, Flags = &h0
		FridayIfSaturday As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mCaption As string
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mCycleFirstYear As Integer = 1900
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mCycleYearDuration As Integer = 1
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDayOff As Boolean = True
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mEndOfValidity As DateTime
	#tag EndProperty

	#tag Property, Flags = &h0
		MondayIfSaturday As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		MondayIfSunday As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		Month As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mStartOfValidity As DateTime
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mTag As Variant
	#tag EndProperty

	#tag Property, Flags = &h0
		NextWeekDay As Integer = 0
	#tag EndProperty

	#tag Property, Flags = &h0
		PreviousWeekDay As Integer = 0
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
			Name="Day"
			Visible=false
			Group="Behavior"
			InitialValue="1"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Month"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="MondayIfSunday"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="MondayIfSaturday"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="NextWeekDay"
			Visible=false
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="AddDays"
			Visible=false
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="FridayIfSaturday"
			Visible=false
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="PreviousWeekDay"
			Visible=false
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="AlwaysNextWeekDay"
			Visible=false
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="AlwaysPreviousWeekDay"
			Visible=false
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
