#tag Class
Protected Class AnnualEventOrthodoxEaster
Implements AnnualEvent
	#tag Method, Flags = &h0
		Function Calculate(year as Integer) As DateAndCaption
		  If year < 1 Then 
		    Break
		    Return Nil
		  End
		  
		  If Me.CycleYearDuration > 1 Then
		    if me.CycleFirstYear > year then Return Nil
		    If Not (year - Me.CycleFirstYear) Mod Me.CycleYearDuration > 0 Then Return Nil
		  end
		  
		  
		  If year < Me.StartOfValidity.Year Or year > Me.EndOfValidity.Year Then Return Nil
		  
		  
		  Var d As DateTime = Me.OrthodoxEaster(year)
		  d = d.AddInterval(0,0,me.DeltaEaster)
		  
		  'If NextWeekDay > 0 Then
		  'Do Until d.DayOfWeek = NextWeekDay 
		  'd = d.AddInterval(0,0,1)
		  'Loop
		  'end
		  
		  If d < Me.StartOfValidity Or d > Me.EndOfValidity Then Return Nil
		  
		  Return New DateAndCaption(d, Me.mCaption)
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
		Sub Constructor(d as AnnualEvent)
		  If d = Nil Then 
		    Raise New NilObjectException
		    Exit Sub
		  End
		  
		  If Not (d IsA AnnualEventOrthodoxEaster) Then
		    Raise New InvalidArgumentException
		    Exit Sub
		  End
		  
		  
		  Var vo As AnnualEventEaster = d.DefinitionObject
		  
		  Me.mCaption = vo.Caption
		  Me.DeltaEaster = vo.DeltaEaster
		  
		  me.DayOff = vo.DayOff
		  
		  Me.CycleFirstYear = vo.CycleFirstYear
		  Me.CycleYearDuration = vo.CycleYearDuration
		  
		  Me.StartOfValidity = vo.StartOfValidity
		  Me.EndOfValidity = vo.EndOfValidity
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(copy as AnnualEventOrthodoxEaster)
		  Me.mCaption = Copy.Caption
		  Me.DeltaEaster = copy.DeltaEaster
		  me.DayOff = copy.DayOff
		  
		  Me.CycleFirstYear = Copy.CycleFirstYear
		  me.CycleYearDuration = Copy.CycleYearDuration
		  
		  Me.StartOfValidity = copy.StartOfValidity
		  Me.EndOfValidity = copy.EndOfValidity
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(lCaption as string = "", lDeltaEaster as Integer = 0)
		  Me.mCaption = lCaption
		  Me.DeltaEaster = lDeltaEaster
		  
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
		  If jf = Nil Then Return Nil
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
		  Var dft As String = "EO"
		  
		  
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
		  
		  
		  row.Column("cyclefirstyear").IntegerValue = Me.mCycleFirstYear
		  row.Column("cycleyearduration").IntegerValue = Me.mCycleYearDuration 
		  row.Column("start").DateTimeValue = Me.mStartOfValidity
		  row.Column("end").DateTimeValue = Me.mEndOfValidity
		  row.Column("dayoff").BooleanValue = Me.mDayOff
		  
		  row.Column("day").IntegerValue = Me.DeltaEaster
		  
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
		  Return "OE" _
		  + Format(Me.DeltaEaster, "00#") _
		  + "|" + Me.mStartOfValidity.SQLDate _
		  + "|" + Me.mEndOfValidity.SQLDate 
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function OrthodoxEaster(year as integer) As DateTime
		  Dim a,b,c As Integer
		  Dim d,e As Integer
		  Dim month As Integer, day As Integer
		  
		  a = year Mod 19
		  b = Floor(year / 4)
		  d = ((a * 19) + 15) Mod 30
		  C = (year + B + d) Mod 7
		  e = d - C
		  
		  e = e + Floor(0.75 * (Floor(year / 100) - 20 / 12))
		  
		  Month = 3 + Floor((e + 40) / 44) + Floor(e / 34)
		  
		  day = e + 28 - 31 * Floor(Month / 4) - 30 * Floor(Month / 5)
		  
		  Return New DateTime(year, month, day, 0, 0 , 0, 0)
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
		  
		  Var dtarget As DateTime = OrthodoxEaster(d.Year).AddInterval(0, 0, Me.DeltaEaster)
		  If dtarget.day = d.Day And dtarget.month = d.Month Then Return True 
		  
		  
		  Return false
		End Function
	#tag EndMethod


	#tag Note, Name = Gregorian vs Julian calendar
		The date of Easter is the Orthodox date, but it's actually the date in the Gregorian calendar that's given
		 (because Xojo works with this calendar!)
		If necessary, you'll have to convert it to the Julian calendar yourself.
	#tag EndNote


	#tag Property, Flags = &h0
		DeltaEaster As Integer = 0
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
		Private mDayOff As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mEndOfValidity As DateTime
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mStartOfValidity As DateTime
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mTag As Variant
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
			Name="DeltaEaster"
			Visible=false
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
