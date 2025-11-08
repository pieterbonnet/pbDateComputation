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
		  Var row As New DatabaseRow
		  Var dft As String = "WD"
		  
		  
		  If Encoding = Nil Or Encoding = Encodings.UTF8 Then
		    
		    row.Column("caption").StringValue = Me.mCaption.DefineEncoding(Encodings.UTF8)
		    If RegionIdentifier <> Nil Then row.Column("region").StringValue = RegionIdentifier.StringValue.DefineEncoding(Encodings.UTF8) Else row.Column("region").StringValue = ""
		    row.Column("fingerprint").StringValue = Me.FingerPrint.DefineEncoding(Encodings.UTF8)
		    row.Column("definitiontype").StringValue = dft.DefineEncoding(Encodings.UTF8)
		    
		  Else
		    
		    row.Column("caption").StringValue = Me.mCaption.DefineEncoding(Encodings.UTF8).ConvertEncoding(Encoding).DefineEncoding(Encoding)
		    If RegionIdentifier <> Nil Then row.Column("region").StringValue = RegionIdentifier.StringValue.DefineEncoding(Encodings.UTF8).ConvertEncoding(Encoding).DefineEncoding(Encoding) Else row.Column("region").StringValue = ""
		    row.Column("fingerprint").StringValue = Me.FingerPrint.DefineEncoding(Encodings.UTF8).ConvertEncoding(Encoding).DefineEncoding(Encoding)
		    row.Column("definitiontype").StringValue = dft.DefineEncoding(Encodings.UTF8).ConvertEncoding(Encoding).DefineEncoding(Encoding)
		    
		  End
		  
		  If Encoding = Nil Or Encoding = Encodings.UTF8 Then
		    
		    row.Column("caption").StringValue = Me.mCaption.DefineEncoding(Encodings.UTF8)
		    If RegionIdentifier <> Nil Then row.Column("region").StringValue = RegionIdentifier.StringValue.DefineEncoding(Encodings.UTF8) Else row.Column("region").StringValue = ""
		    row.Column("fingerprint").StringValue = Me.FingerPrint.DefineEncoding(Encodings.UTF8)
		  Else
		    
		    row.Column("caption").StringValue = Me.mCaption.DefineEncoding(Encodings.UTF8).ConvertEncoding(Encoding).DefineEncoding(Encoding)
		    If RegionIdentifier <> Nil Then row.Column("region").StringValue = RegionIdentifier.StringValue.DefineEncoding(Encodings.UTF8).ConvertEncoding(Encoding).DefineEncoding(Encoding) Else row.Column("region").StringValue = ""
		    row.Column("fingerprint").StringValue = Me.FingerPrint.DefineEncoding(Encodings.UTF8).ConvertEncoding(Encoding).DefineEncoding(Encoding)
		    
		  End
		  
		  row.Column("cyclefirstyear").IntegerValue = Me.mCycleFirstYear
		  row.Column("cycleyearduration").IntegerValue = Me.mCycleYearDuration 
		  row.Column("start").DateTimeValue = Me.mStartOfValidity
		  row.Column("end").DateTimeValue = Me.mEndOfValidity
		  row.Column("dayoff").BooleanValue = Me.mDayOff
		  
		  row.Column("month").IntegerValue = me.Month
		  row.Column("weekday").IntegerValue = Me.WeekDay
		  row.Column("rank").IntegerValue =  me.Rank
		  
		  row.Column("adddays").IntegerValue = Me.AddDays
		  row.Column("nextweekday").IntegerValue = Me.NextWeekDay
		  row.Column("previousweekday").IntegerValue = Me.PreviousWeekDay
		  
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
		  Return "W" _
		  + Format(Me.Month, "0#") _
		  + ";" + me.WeekDay.ToString _
		  + ";" + me.rank.ToString _
		  + ";" + me.AddDays.ToString _
		  + "|" + Me.mStartOfValidity.SQLDate _
		  + "|" + Me.mEndOfValidity.SQLDate 
		  
		  
		  'Return "W" _
		  '+ Format(Me.Month, "0#") _
		  '+ Me.WeekDay.ToString _
		  '+ Me.rank.ToString _
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
