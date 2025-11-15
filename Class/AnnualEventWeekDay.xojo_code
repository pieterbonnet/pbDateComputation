#tag Class
Protected Class AnnualEventWeekDay
Implements AnnualEvent
	#tag Method, Flags = &h0
		Function Calculate(year as Integer) As DateAndCaption
		  If year < 1607 Then 
		    Break
		    Return Nil
		  End
		  
		  If Me.CycleYearDuration > 1 Then
		    if me.CycleFirstYear > year then Return Nil
		    If (year - Me.CycleFirstYear) Mod Me.CycleYearDuration > 0 Then Return Nil
		  End
		  
		  
		  If year < Me.StartOfValidity.Year Or year > Me.EndOfValidity.Year Then Return Nil
		  
		  If WeekDay < 1 or WeekDay > 7 or  Rank < 1 or Rank > 6 or month < 1 or month > 12 Then Return nil
		  
		  Var d As New DateTime(year, Month, 1, 0, 0, 0) // First day of the month
		  
		  // Trivial, the first day is the good day
		  
		  If d.DayOfWeek = Me.WeekDay And Me.Rank = 1 Then 
		    
		    // Nothing else
		    
		  elseIf Rank = 6 Then // Last
		    d = d.AddInterval(0,1)         // First day of the next month
		    d = d.SubtractInterval(0,0,1)  // Last day of the "current" month
		    Do Until d.DayOfWeek = WeekDay
		      d = d.SubtractInterval(0,0,1)
		    Loop
		    
		  else 
		    
		    Var rang As Integer = 0
		    
		    Do 
		      If rang = 0 Then d = d.AddInterval (0,0,1) Else d = d.AddInterval(0,0,7)
		      If d.DayOfWeek = Me.WeekDay Then rang = rang + 1
		    Loop Until  (d.DayOfWeek = Me.WeekDay And rang = Me.Rank) Or (d.Month <> Me.Month)
		    
		    If d.Month <> Me.Month Then Return Nil // Sometimes, there is no 5th Monday of the month.
		    
		  End
		  
		  If PreviousWeekDay > 0 Then
		    Do Until d.DayOfWeek = PreviousWeekDay 
		      d = d.SubtractInterval(0,0,1)
		    Loop
		  elseIf NextWeekDay > 0 Then
		    Do Until d.DayOfWeek = NextWeekDay 
		      d = d.AddInterval(0,0,1)
		    Loop
		  ElseIf AddDays <> 0 Then
		    d = d.AddInterval(0,0,AddDays)
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
		  mCaption = s
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(d as AnnualEvent)
		  If d = Nil Then 
		    Raise New NilObjectException
		    Exit Sub
		  End
		  
		  If Not (d IsA AnnualEventWeekDay) Then
		    Raise New InvalidArgumentException
		    Exit Sub
		  End
		  
		  Var vo As AnnualEventWeekDay = d.DefinitionObject
		  
		  Me.mCaption = vo.Caption
		  Me.Month = vo.Month
		  Me.WeekDay = vo.WeekDay
		  Me.Rank = vo.Rank
		  Me.NextWeekDay = vo.NextWeekDay
		  Me.PreviousWeekDay = vo.PreviousWeekDay
		  Me.DayOff = vo.DayOff
		  Me.AddDays = vo.AddDays
		  
		  Me.CycleFirstYear = vo.CycleFirstYear
		  Me.CycleYearDuration = vo.CycleYearDuration
		  
		  Me.StartOfValidity = vo.StartOfValidity
		  Me.EndOfValidity = vo.EndOfValidity
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(copy as AnnualEventWeekDay)
		  Me.mCaption = Copy.Caption
		  Me.Month = copy.Month
		  Me.WeekDay = copy.WeekDay
		  Me.Rank = copy.Rank
		  Me.NextWeekDay = copy.NextWeekDay
		  me.PreviousWeekDay = copy.PreviousWeekDay
		  me.DayOff = copy.DayOff
		  me.AddDays = copy.AddDays
		  
		  Me.CycleFirstYear = Copy.CycleFirstYear
		  me.CycleYearDuration = Copy.CycleYearDuration
		  
		  Me.StartOfValidity = copy.StartOfValidity
		  Me.EndOfValidity = copy.EndOfValidity
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(lCaption as string = "", lMonth as Integer = 1, lWeekDay as Integer = 1, lRank as Integer = 1)
		  Me.mCaption = lCaption
		  
		  Me.Month = lMonth
		  me.WeekDay = lWeekDay
		  me.Rank = lRank
		  
		  
		  Me.StartOfValidity = New DateTime(1,1,1,0,0,0)
		  Me.EndOfValidity = New DateTime(3999,12,31,23,59,59)
		  
		  
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
		  Var jf As DateAndCaption = Me.Calculate(year)
		  
		  If jf = Nil Then 
		    Var d As DateTime
		    Return d
		  end
		  
		  Return jf.DateValue
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
		  
		  s = "W|"
		  s = s + Me.StartOfValidity.ToString("yyyy-MM-dd") + "|"
		  s = s + Me.EndOfValidity.ToString("yyyy-MM-dd") + "|"
		  s = s + Me.CycleFirstYear.ToString(Locale.Raw) + "|"
		  s = s + Me.CycleYearDuration.ToString(Locale.Raw) + "|"
		  If Me.DayOff Then s = s + "1|" Else s = s + "0|"
		  s = s + Me.Month.ToString(locale.Raw) + "|"
		  s = s + Me.WeekDay.ToString(locale.Raw) + "|"
		  s = s + Me.Rank.ToString(locale.Raw) + "|"
		  
		  
		  If Me.NextWeekDay > 0 Then
		    s = s + "D" + Me.NextWeekDay.ToString(Locale.Raw)
		  ElseIf Me.PreviousWeekDay > 0 Then
		    s = s + "D-" + Me.PreviousWeekDay.ToString(Locale.Raw) 
		  ElseIf Me.AddDays <> 0 Then
		    s = s + "E" + Me.AddDays.ToString(Locale.Raw)
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
		    adf.CycleFirstYear = s(4).ToInteger
		    adf.CycleYearDuration = s(5).ToInteger
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
		  If Me.CycleYearDuration > 1 Then
		    if me.CycleFirstYear > d.year then Return False
		    If (d.year - Me.CycleFirstYear) Mod Me.CycleYearDuration > 0 Then Return false
		  end
		  
		  If d < Me.StartOfValidity Or d > Me.EndOfValidity Then Return False
		  
		  Var DaL as DateAndCaption = Me.Calculate(d.Year)
		  if DaL = Nil then Return False
		  
		  var dt as DateTime = DaL.DateValue
		  
		  If dt.day = d.Day And dt.month = d.Month And d.Year = dt.Year Then Return True
		  
		  Return false
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ToString() As String
		  Var s as String
		  
		  s = "W|"
		  s = s + EncodeBase64(Me.Caption,0) + "|"
		  s = s + Me.StartOfValidity.ToString("yyyy-MM-dd") + "|"
		  s = s + Me.EndOfValidity.ToString("yyyy-MM-dd") + "|"
		  s = s + Me.CycleFirstYear.ToString(Locale.Raw) + "|"
		  s = s + Me.CycleYearDuration.ToString(Locale.Raw) + "|"
		  If Me.DayOff Then s = s + "1|" Else s = s + "0|"
		  s = s + Me.Month.ToString(locale.Raw) + "|"
		  s = s + Me.WeekDay.ToString(locale.Raw) + "|"
		  s = s + Me.Rank.ToString(locale.Raw) + "|"
		  
		  
		  If Me.NextWeekDay > 0 Then
		    s = s + "D" + Me.NextWeekDay.ToString(Locale.Raw)
		  ElseIf Me.PreviousWeekDay > 0 Then
		    s = s + "D-" + Me.PreviousWeekDay.ToString(Locale.Raw) 
		  ElseIf Me.AddDays <> 0 then
		    s = s + "E" + Me.AddDays.ToString(Locale.Raw)
		  Else
		    s = s + "N"
		  End
		  
		  
		  Return s
		End Function
	#tag EndMethod


	#tag Note, Name = Untitled
		
		
	#tag EndNote


	#tag Property, Flags = &h0
		AddDays As Integer = 0
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

	#tag Property, Flags = &h0
		Rank As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		WeekDay As Integer = 0
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
			Name="Month"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="WeekDay"
			Visible=false
			Group="Behavior"
			InitialValue="=0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Rank"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="Integer"
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
			Name="PreviousWeekDay"
			Visible=false
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
