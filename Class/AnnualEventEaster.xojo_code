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
		  
		  
		  Var d As DateTime = AnnualEventEaster.Easter(year)
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
		Shared Function Easter(Year as Integer) As DateTime
		  Var h As Integer
		  Var i As Integer
		  Var j As Integer
		  Var l As Integer
		  Var m As Integer
		  Var d As Integer
		  
		  
		  H = (24 + 19*(Year Mod 19)) Mod 30
		  I = H - (H\28)
		  J = (Year + Year\4 + I - 13) Mod 7
		  L = I - J
		  m = 3 + (L + 40)\44
		  d= L + 28 - 31*(m\4)
		  
		  Return New DateTime(Year, m, d, 0, 0 , 0, 0)
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
		  
		  s = "E|"
		  s = s + Me.StartOfValidity.ToString("yyyy-MM-dd") + "|"
		  s = s + Me.EndOfValidity.ToString("yyyy-MM-dd") + "|"
		  s = s + Me.CycleFirstYear.ToString(Locale.Raw) + "|"
		  s = s + Me.CycleYearDuration.ToString(Locale.Raw) + "|"
		  If Me.DayOff Then s = s + "1|" Else s = s + "0|"
		  
		  s = s + Me.DeltaEaster.ToString(Locale.Raw) 
		  
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
		    Var y As Integer = d.year - Me.CycleFirstYear
		    If (d.year - Me.CycleFirstYear) Mod Me.CycleYearDuration > 0 Then Return false
		    
		  end
		  
		  If d < Me.StartOfValidity Or d > Me.EndOfValidity Then Return False
		  
		  Var dtarget As DateTime = AnnualEventEaster.Easter(d.Year).AddInterval(0, 0, Me.DeltaEaster)
		  If dtarget.day = d.Day And dtarget.month = d.Month Then Return True 
		  
		  Return false
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ToString() As String
		  Var s As String
		  
		  s = "E|"
		  s = s + EncodeBase64(Me.Caption,0) + "|"
		  s = s + Me.StartOfValidity.ToString("yyyy-MM-dd") + "|"
		  s = s + Me.EndOfValidity.ToString("yyyy-MM-dd") + "|"
		  s = s + Me.CycleFirstYear.ToString(Locale.Raw) + "|"
		  s = s + Me.CycleYearDuration.ToString(Locale.Raw) + "|"
		  If Me.DayOff Then s = s + "1|" Else s = s + "0|"
		  
		  s = s + Me.DeltaEaster.ToString(Locale.Raw) 
		  
		  Return s
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
