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
		  Me.AlwaysNextWeekDay = vo.AlwaysNextWeekDay
		  me.AlwaysPreviousWeekDay = vo.AlwaysPreviousWeekDay
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
		  me.AlwaysNextWeekDay = copy.AlwaysNextWeekDay
		  me.AlwaysPreviousWeekDay = copy.AlwaysPreviousWeekDay
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
		  
		  
		  If Encoding = Nil Or Encoding = Encodings.UTF8 Then
		    
		    if RegionIdentifier <> nil then row.Column("region").StringValue = RegionIdentifier.StringValue.DefineEncoding(Encodings.UTF8) else row.Column("region").StringValue = ""
		    row.Column("value").StringValue = me.ToString.DefineEncoding(Encodings.UTF8)
		    
		  Else
		    
		    if RegionIdentifier <> nil then row.Column("region").StringValue = RegionIdentifier.StringValue.DefineEncoding(Encodings.UTF8).ConvertEncoding(Encoding).DefineEncoding(encoding) else row.Column("region").StringValue = ""
		    row.Column("value").StringValue = me.ToString.DefineEncoding(Encodings.UTF8).ConvertEncoding(Encoding).DefineEncoding(encoding)
		    
		    
		  End
		  
		  
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
		  Var s As String
		  
		  s = "F|" // 0
		  s = s + Me.StartOfValidity.ToString("yyyy-MM-dd") + "|" // 2
		  s = s + Me.EndOfValidity.ToString("yyyy-MM-dd") + "|" // 3
		  s = s + Me.CycleFirstYear.ToString(Locale.Raw) + "|" // 4
		  s = s + Me.CycleYearDuration.ToString(Locale.Raw) + "|" // 5
		  If Me.DayOff Then s = s + "1|" Else s = s + "0|" // 6
		  s = s + Me.Month.ToString(Locale.Raw) + "|" // 7
		  s = s + Me.Day.ToString(Locale.Raw) + "|" // 8
		  
		  
		  If Me.MondayIfSunday And Me.FridayIfSaturday Then
		    s = s + "A"
		  ElseIf Me.FridayIfSaturday Then
		    s = s + "B"
		  ElseIf Me.MondayIfSunday Then
		    s = s + "C"
		  ElseIf Me.AlwaysNextWeekDay > 0 Then
		    s = s + "D" + Me.AlwaysNextWeekDay.ToString(Locale.Raw)
		  ElseIf  Me.AlwaysPreviousWeekDay > 0 Then
		    s = s + "D-" + Me.AlwaysPreviousWeekDay.ToString(Locale.Raw)
		  ElseIf Me.NextWeekDay > 0 Then
		    s = s + "E" + Me.NextWeekDay.ToString(Locale.Raw)
		  ElseIf Me.PreviousWeekDay > 0 Then
		    s = s + "E-" + Me.PreviousWeekDay.ToString(Locale.Raw) 
		  ElseIf Me.AddDays <> 0 Then
		    s = s + "F" + Me.AddDays.ToString(Locale.Raw)
		  Else
		    s = s + "N"
		  End
		  
		  
		  Return s
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function FromString(value as String) As AnnualEvent
		  Var s() As String = value.Split("|")
		  Var x As Integer
		  
		  If s.LastIndex < 7 Then 
		    Raise new InvalidArgumentException
		    Return Nil
		  end
		  
		  select case s(0)
		    
		  case "F"
		    
		    if s.LastIndex <> 9 then
		      Raise New InvalidArgumentException
		      Return Nil
		    end
		    
		    Var adf As New AnnualEventFix
		    If s(1).Trim <> "" Then adf.Caption = DecodeBase64(s(1))
		    
		    adf.StartOfValidity = DateTime.FromString(s(2))
		    adf.EndOfValidity = DateTime.FromString(s(3))
		    adf.mCycleFirstYear = s(4).ToInteger
		    adf.mCycleYearDuration = s(5).ToInteger
		    adf.DayOff = (s(6)="1")
		    adf.Month = s(7).ToInteger
		    adf.day = s(8).ToInteger
		    
		    select case s(9).Left(1)
		      
		    case "N"
		      
		      // Nothing
		      
		    Case "A"
		      
		      adf.FridayIfSaturday = True
		      adf.MondayIfSunday = True
		      
		    case "B"
		      
		      adf.FridayIfSaturday = True
		      
		    case "C"
		      
		      adf.MondayIfSunday = True
		      
		    case "D"
		      
		      x = s(9).right(s(9).Length-1).ToInteger
		      
		      If x > 0 Then
		        adf.AlwaysNextWeekDay = x
		      ElseIf x < 0 Then
		        adf.AlwaysPreviousWeekDay = x * -1
		      End
		      
		    case "E"
		      
		      x = s(9).right(s(9).Length-1).ToInteger
		      
		      If x > 0 Then
		        adf.NextWeekDay = x
		      ElseIf x < 0 Then
		        adf.PreviousWeekDay = x * -1
		      End
		      
		    case "F"
		      
		      x = s(9).right(s(9).Length-1).ToInteger
		      
		      adf.AddDays = x
		      
		    end Select
		    
		    Return adf
		    
		  case "E"
		    
		    Var ae as new AnnualEventEaster
		    
		    If s.LastIndex <> 7 Then
		      Raise New InvalidArgumentException
		      Return Nil
		    end
		    
		    If s(1).Trim <> "" Then ae.Caption = DecodeBase64(s(1))
		    
		    ae.StartOfValidity = DateTime.FromString(s(2))
		    ae.EndOfValidity = DateTime.FromString(s(3))
		    ae.CycleFirstYear = s(4).ToInteger
		    ae.CycleYearDuration = s(5).ToInteger
		    ae.DayOff = (s(6)="1")
		    ae.DeltaEaster = s(7).ToInteger
		    
		    Return ae
		    
		  case "O"
		    
		    Var ae As New AnnualEventOrthodoxEaster
		    
		    If s.LastIndex <> 7 Then
		      Raise New InvalidArgumentException
		      Return nil
		    End
		    
		    If s(1).Trim <> "" Then ae.Caption = DecodeBase64(s(1))
		    
		    ae.StartOfValidity = DateTime.FromString(s(2))
		    ae.EndOfValidity = DateTime.FromString(s(3))
		    ae.CycleFirstYear = s(4).ToInteger
		    ae.CycleYearDuration = s(5).ToInteger
		    ae.DayOff = (s(6)="1")
		    ae.DeltaEaster = s(7).ToInteger
		    
		    Return ae
		    
		  case "W"
		    
		    If s.LastIndex <> 10 Then 
		      Raise New InvalidArgumentException
		      Return Nil
		    End
		    
		    Var aw as new AnnualEventWeekDay
		    
		    If s(1).Trim <> "" Then aw.Caption = DecodeBase64(s(1))
		    
		    aw.StartOfValidity = DateTime.FromString(s(2))
		    aw.EndOfValidity = DateTime.FromString(s(3))
		    aw.CycleFirstYear = s(4).ToInteger
		    aw.CycleYearDuration = s(5).ToInteger
		    aw.DayOff = (s(6)="1")
		    
		    aw.Month = s(7).ToInteger
		    aw.WeekDay = s(8).ToInteger
		    aw.Rank = s(9).ToInteger
		    
		    
		    
		    Select Case s(10).Left(1)
		      
		    Case "N"
		      
		      // Nothing
		      
		    Case "A", "B", "C"
		      
		      // Nothing
		      
		    Case "D"
		      
		      x = s(10).Right(s(10).Length-1).ToInteger
		      
		      If x > 0 Then
		        aw.NextWeekDay = x
		      ElseIf x < 0 Then
		        aw.PreviousWeekDay = x * -1
		      End
		      
		    Case "E"
		      
		      x = s(10).Right(s(10).Length-1).ToInteger
		      
		      aw.AddDays = x
		      
		    end Select
		    
		  Else
		    
		    Raise New InvalidArgumentException
		    Return Nil
		    
		  end
		  
		  
		  
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
		  
		  If d < Me.StartOfValidity Or d > Me.EndOfValidity Then Return False
		  
		  If Me.CycleYearDuration > 1 Then
		    if me.CycleFirstYear > d.year then Return False
		    If (d.year - Me.CycleFirstYear) Mod Me.CycleYearDuration > 0 Then Return false
		  end
		  
		  
		  
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

	#tag Method, Flags = &h0
		Function ToString() As String
		  Var s as String
		  
		  s = "F|" // 0
		  s = s + EncodeBase64(Me.Caption,0) + "|" // 1
		  s = s + Me.StartOfValidity.ToString("yyyy-MM-dd") + "|" // 2
		  s = s + Me.EndOfValidity.ToString("yyyy-MM-dd") + "|" // 3
		  s = s + Me.CycleFirstYear.ToString(Locale.Raw) + "|" // 4
		  s = s + Me.CycleYearDuration.ToString(Locale.Raw) + "|" // 5
		  If me.DayOff Then s = s + "1|" Else s = s + "0|" // 6
		  s = s + Me.Month.ToString(Locale.Raw) + "|" // 7
		  s = s + Me.Day.ToString(Locale.Raw) + "|" // 8
		  
		  
		  if me.MondayIfSunday and me.FridayIfSaturday then // 9
		    s = s + "A"
		  ElseIf Me.FridayIfSaturday Then
		    s = s + "B"
		  ElseIf Me.MondayIfSunday Then
		    s = s + "C"
		  ElseIf Me.AlwaysNextWeekDay > 0 Then
		    s = s + "D" + Me.AlwaysNextWeekDay.ToString(Locale.Raw)
		  ElseIf  Me.AlwaysPreviousWeekDay > 0 Then
		    s = s + "D-" + Me.AlwaysPreviousWeekDay.ToString(Locale.Raw)
		  Elseif Me.NextWeekDay > 0 Then
		    s = s + "E" + Me.NextWeekDay.ToString(Locale.Raw)
		  ElseIf Me.PreviousWeekDay > 0 Then
		    s = s + "E-" + Me.PreviousWeekDay.ToString(Locale.Raw) 
		  Elseif me.AddDays <> 0 Then
		    s = s + "F" + Me.AddDays.ToString(Locale.Raw)
		  else
		    s = s + "N"
		  End
		  
		  
		  Return s
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
