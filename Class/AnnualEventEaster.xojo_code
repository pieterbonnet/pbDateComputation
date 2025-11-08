#tag Class
Protected Class AnnualEventEaster
Implements AnnualEvent
	#tag Method, Flags = &h0
		Function Calculate(year as Integer) As DateAndCaption
		  If year < 1 Then 
		    Break
		    Return Nil
		  End
		  
		  If Me.CycleYearDuration > 1 Then
		    if me.CycleFirstYear > year then Return Nil
		    If (year - Me.CycleFirstYear) Mod Me.CycleYearDuration > 0 Then Return Nil
		  end
		  
		  
		  If year < Me.StartOfValidity.Year Or year > Me.EndOfValidity.Year Then Return Nil
		  
		  
		  Var d As DateTime = Me.Easter(year)
		  d = d.AddInterval(0,0,me.DeltaEaster)
		  
		  
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
		  
		  If Not (d IsA AnnualEventEaster) Then
		    Raise New InvalidArgumentException
		    Exit Sub
		  End
		  
		  
		  Var vo As AnnualEventEaster = d.DefinitionObject
		  
		  Me.mCaption = vo.Caption
		  Me.DeltaEaster = vo.DeltaEaster
		  
		  me.DayOff = vo.DayOff
		  
		  me.mCycleFirstYear = vo.CycleFirstYear
		  me.mCycleYearDuration = vo.CycleYearDuration
		  
		  Me.StartOfValidity = vo.StartOfValidity
		  Me.EndOfValidity = vo.EndOfValidity
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(copy as AnnualEventEaster)
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
		  Var dft As String = "E"
		  
		  
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

	#tag Method, Flags = &h21
		Private Function Easter(year as integer) As DateTime
		  Var h As Integer
		  Var i As Integer
		  Var j As Integer
		  Var l As Integer
		  Var m As Integer
		  Var d As Integer
		  
		  
		  H = (24 + 19*(year Mod 19)) Mod 30
		  I = H - (H\28)
		  J = (year + year\4 + I - 13) Mod 7
		  L = I - J
		  m = 3 + (L + 40)\44
		  d= L + 28 - 31*(m\4)
		  
		  Return New DateTime(year, m, d, 0, 0 , 0, 0)
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
		  Return "E" _
		  + Format(Me.DeltaEaster, "00#") _
		  + "|" + Me.mStartOfValidity.SQLDate _
		  + "|" + Me.mEndOfValidity.SQLDate 
		  
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
		    Var y As Integer = d.year - Me.CycleFirstYear
		    If (d.year - Me.CycleFirstYear) Mod Me.CycleYearDuration > 0 Then Return false
		    
		  end
		  
		  If d < Me.StartOfValidity Or d > Me.EndOfValidity Then Return False
		  
		  Var dtarget As DateTime = Easter(d.Year).AddInterval(0, 0, Me.DeltaEaster)
		  If dtarget.day = d.Day And dtarget.month = d.Month Then Return True 
		  
		  Return false
		End Function
	#tag EndMethod


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
		Private mDayOff As Boolean = True
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


	#tag Constant, Name = AshWednesday, Type = Double, Dynamic = False, Default = \"-46", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"-46"
	#tag EndConstant

	#tag Constant, Name = CareSunday, Type = Double, Dynamic = False, Default = \"-14", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"-14"
	#tag EndConstant

	#tag Constant, Name = CleanMonday, Type = Double, Dynamic = False, Default = \"-48", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"-48"
	#tag EndConstant

	#tag Constant, Name = CorpusChristi, Type = Double, Dynamic = False, Default = \"60", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"60"
	#tag EndConstant

	#tag Constant, Name = CorpusDomini, Type = Double, Dynamic = False, Default = \"60", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"60"
	#tag EndConstant

	#tag Constant, Name = Easter, Type = Double, Dynamic = False, Default = \"0", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"0"
	#tag EndConstant

	#tag Constant, Name = FridayOfSorrows, Type = Double, Dynamic = False, Default = \"-9", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"-9"
	#tag EndConstant

	#tag Constant, Name = GoodFriday, Type = Double, Dynamic = False, Default = \"-2", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"-2"
	#tag EndConstant

	#tag Constant, Name = HolyFriday, Type = Double, Dynamic = False, Default = \"-2", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"-2"
	#tag EndConstant

	#tag Constant, Name = HolySaturday, Type = Double, Dynamic = False, Default = \"-1", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"-1"
	#tag EndConstant

	#tag Constant, Name = HolyThursday, Type = Double, Dynamic = False, Default = \"-3", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"-3"
	#tag EndConstant

	#tag Constant, Name = MardiGras, Type = Double, Dynamic = False, Default = \"-47", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"-47"
	#tag EndConstant

	#tag Constant, Name = MaundyThursday, Type = Double, Dynamic = False, Default = \"-3", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"-3"
	#tag EndConstant

	#tag Constant, Name = MotheringSunday, Type = Double, Dynamic = False, Default = \"-21", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"-21"
	#tag EndConstant

	#tag Constant, Name = PalmSunday, Type = Double, Dynamic = False, Default = \"-7", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"-7"
	#tag EndConstant

	#tag Constant, Name = PancakeDay, Type = Double, Dynamic = False, Default = \"-47", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"-47"
	#tag EndConstant

	#tag Constant, Name = PassionSunday, Type = Double, Dynamic = False, Default = \"-14", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"-14"
	#tag EndConstant

	#tag Constant, Name = Pentecost, Type = Double, Dynamic = False, Default = \"49", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"49"
	#tag EndConstant

	#tag Constant, Name = PentecostMonday, Type = Double, Dynamic = False, Default = \"50", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"50"
	#tag EndConstant

	#tag Constant, Name = Quasimodo, Type = Double, Dynamic = False, Default = \"7", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"7"
	#tag EndConstant

	#tag Constant, Name = ResurrectionSunday, Type = Double, Dynamic = False, Default = \"0", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"0"
	#tag EndConstant

	#tag Constant, Name = SecondSundayofEaster, Type = Double, Dynamic = False, Default = \"7", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"7"
	#tag EndConstant

	#tag Constant, Name = ShroveMonday, Type = Double, Dynamic = False, Default = \"-48", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"-48"
	#tag EndConstant

	#tag Constant, Name = ShroveTuesday, Type = Double, Dynamic = False, Default = \"-47", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"-47"
	#tag EndConstant

	#tag Constant, Name = SundayAfterCorpusChristi, Type = Double, Dynamic = False, Default = \"63", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"63"
	#tag EndConstant

	#tag Constant, Name = TrinitySunday, Type = Double, Dynamic = False, Default = \"56", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"56"
	#tag EndConstant

	#tag Constant, Name = WhitMonday, Type = Double, Dynamic = False, Default = \"50", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"50"
	#tag EndConstant


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
