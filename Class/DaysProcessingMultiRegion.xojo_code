#tag Class
Protected Class DaysProcessingMultiRegion
	#tag Method, Flags = &h21
		Private Sub AddClosurePeriods(FirstDay as DateTime, LastDay as DateTime, cp() as ClosurePeriod, R as DaysProcessingRegion)
		  For c As Integer = 0 To R.ClosurePeriods.LastIndex
		    If FirstDay <= R.ClosurePeriods(c).LastDay And LastDay >= R.ClosurePeriods(c).FirstDay Then cp.Add R.ClosurePeriods(c)
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub AddInArray(d() as DateAndCaption, valeur as DateAndCaption)
		  If d.LastIndex = -1 Then
		    d.Add valeur
		    Exit Sub
		  End
		  
		  If valeur.DateValue >= d(d.LastIndex).DateValue Then
		    d.Add valeur
		    Exit Sub
		  End
		  
		  For i As Integer = 0 To d.LastIndex
		    
		    If valeur.DateValue <= d(i).DateValue Then
		      d.AddAt(i, valeur)
		      Exit Sub
		    End
		    
		  Next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function AnnualEventMatch(d as DateTime, OnlyIfDayOff as Boolean = False) As Boolean
		  If d = Nil Then 
		    Raise New NilObjectException
		    Return False
		  end
		  
		  if not OnRegionEnabled then Return False
		  
		  // The date must be a public holiday for ALL regions.
		  
		  For i As Integer = 0 To Me.Regions.LastIndex
		    
		    if not me.Regions(i).Enabled then Continue
		    
		    // Therefore, it is enough for just one of them to not match.
		    if not me.Regions(i).AnnualEventMatch(d, OnlyIfDayOff) then Return false
		    
		  Next
		  
		  Return True
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function BusinessDays(starting as DateTime, ending as DateTime) As DateTime()
		  If starting = Nil Then 
		    Raise New NilObjectException
		    Return Nil
		  End
		  
		  If ending = Nil Then 
		    Raise New NilObjectException
		    Return Nil
		  End
		  
		  Var list() As DateTime
		  if not me.OnRegionEnabled then Return list
		  
		  Var nStart As New DateTime(starting.Year, starting.Month, starting.day)
		  Var nEnd As New DateTime(ending.Year, ending.Month, ending.day)
		  
		  If nStart > nEnd Then Return list
		  
		  // Copy the relevant items ClosurePeriod
		  Var cp() As ClosurePeriod 
		  For r As Integer = 0 To Me.Regions.LastIndex
		    If Me.Regions(r).Enabled Then
		      AddClosurePeriods(starting, ending, cp, me.Regions(r))
		    End If
		  next
		  
		  // Determines which days of the week are open (or closed) for all regions
		  Var WorkingWeekDays(6) as Boolean
		  For r As Integer = 0 To 6
		    WorkingWeekDays(r) = True
		  next
		  
		  For r As Integer = 0 To Me.Regions.LastIndex
		    For i As Integer = 1 To 7
		      if not me.Regions(r).WorkingWeekDays.WorkingDay(i) then WorkingWeekDays(i-1) = False // It is enough that in just one region that day is not a working day...
		    Next i
		  next r
		  
		  Var d As DateTime = nStart.SubtractInterval(0,0,1)
		  
		  Do Until d = nEnd
		    
		    d = d.AddInterval(0,0,1)
		    
		    if WorkingWeekDays(d.DayOfWeek - 1) = False then Continue do 
		    
		    for c as integer = 0 to cp.LastIndex
		      if d >= cp(c).FirstDay and d <= cp(c).LastDay then continue do // It's a closed period....
		    next 
		    
		    for i as Integer = 0 to me.Regions.LastIndex
		      
		      if not me.Regions(i).Enabled then Continue // ignore
		      
		      If Me.Regions(i).AnnualEventMatch(d, True) Then  // if is a holiday
		        Continue Do // It is enough that in just one region that day is not a working day...
		      end
		      
		    next 
		    
		    list.Add New DateTime(d) // Working day for all the regions
		    
		  Loop
		  
		  Return list
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function BusinessDays(year as Integer) As DateTime()
		  Var d As New DateTime(year, 1, 1, 0, 0, 0)
		  Var f As New DateTime(year, 12, 31, 0, 0, 0)
		  
		  Return BusinessDays(d,f)
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor()
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(copy as DaysProcessingMultiRegion)
		  If copy.Regions.Count > 0 Then
		    
		    For r As Integer = 0 To copy.Regions.LastIndex
		      
		      me.Regions.Add new DaysProcessingRegion(copy.Regions(r))
		      
		    next
		    
		    
		  end
		  
		  
		  'Me.WorkingWeekDays = New dprWorkingWeekDays
		  '
		  'For i As Integer = 1 To 7
		  'Me.WorkingWeekDays.WorkingDay(i) = copy.WorkingWeekDays.WorkingDay(i)
		  'Next
		  '
		  'Var df As AnnualEventFix
		  'Var de As AnnualEventEaster
		  'Var doe As AnnualEventOrthodoxEaster
		  'var dwd as AnnualEventWeekDay
		  '
		  'If copy.AnnualEvents.Count > 0 Then
		  '
		  'For i As Integer = 0 To copy.AnnualEvents.LastIndex
		  '
		  'select case copy.AnnualEvents(i) 
		  '
		  'Case IsA AnnualEventFix
		  '
		  'df = copy.AnnualEvents(i).DefinitionObject
		  'me.AnnualEvents.Add new AnnualEventFix(df)
		  '
		  'Case IsA AnnualEventEaster
		  '
		  'de =  copy.AnnualEvents(i).DefinitionObject
		  'me.AnnualEvents.Add new AnnualEventEaster(de)
		  '
		  'Case IsA AnnualEventOrthodoxEaster
		  '
		  'doe =  copy.AnnualEvents(i).DefinitionObject
		  'Me.AnnualEvents.Add New AnnualEventOrthodoxEaster(doe)
		  '
		  'Case IsA AnnualEventWeekDay
		  '
		  'dwd =  copy.AnnualEvents(i).DefinitionObject
		  'Me.AnnualEvents.Add New AnnualEventWeekDay(dwd)
		  '
		  'end
		  '
		  'Next
		  '
		  'End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CountBusinessDays(Date1 as datetime, Date2 as datetime) As Integer
		  // Trivial case
		  
		  If Date1 = Nil Then 
		    Raise New NilObjectException
		    Return -1
		  End
		  
		  If date2 = Nil Then 
		    Raise New NilObjectException
		    Return -1
		  End
		  
		  if not OnRegionEnabled then Return -1
		  
		  If date1.day= date2.day _
		  And date1.Month = date2.Month _
		  and date1.Year = date2.Year Then Return 0
		  
		  Var dtCurrent As DateTime, dtMax As DateTime, counter As Integer
		  
		  If date1 > Date2 Then
		    dtMax = New DateTime(date1)
		    dtCurrent = New DateTime(Date2)
		  else
		    dtMax = New DateTime(date2)
		    dtCurrent = New DateTime(Date1)
		  End
		  
		  
		  // Copy the relevant items ClosurePeriod
		  Var cp() As ClosurePeriod 
		  For r As Integer = 0 To Me.Regions.LastIndex
		    If Me.Regions(r).Enabled Then
		      AddClosurePeriods(date1, date2, cp, Me.Regions(r))
		    End If
		  next
		  
		  // Determines which days of the week are open (or closed) for all regions
		  Var WorkingWeekDays(6) As Boolean
		  For r As Integer = 0 To 6
		    WorkingWeekDays(r) = True
		  Next
		  
		  For r As Integer = 0 To Me.Regions.LastIndex
		    For i As Integer = 1 To 7
		      If Not Me.Regions(r).WorkingWeekDays.WorkingDay(i) Then WorkingWeekDays(i-1) = False // It is enough that in just one region that day is not a working day...
		    Next i
		  next r
		  
		  dtCurrent = dtCurrent.SubtractInterval(0,0,1)
		  
		  Do 
		    
		    dtCurrent = dtCurrent.AddInterval (0,0,1)
		    
		    if WorkingWeekDays(dtCurrent.DayOfWeek - 1) = False then Continue do 
		    
		    For c As Integer = 0 To cp.LastIndex
		      If dtCurrent >= cp(c).FirstDay And dtCurrent <= cp(c).LastDay Then Continue Do // It's a closed period....
		    next 
		    
		    
		    for i as integer = 0 to me.Regions.LastIndex
		      
		      if me.Regions(i).Enabled then continue for
		      
		      // it's a working (week)day but is this a public holiday - day off ?
		      If Me.Regions(i).AnnualEventMatch(dtCurrent, True) Then Continue Do // It is enough that in just one region that day is not a working day...
		      
		      
		    next i
		    
		    
		    
		    counter = counter + 1
		    
		  loop until dtCurrent >= dtMax
		  
		  Return counter
		  
		  
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CountNonWorkingDays(Date1 as datetime, Date2 as datetime) As Integer
		  // Trivial cases
		  
		  If Date1 = Nil Then 
		    Raise New NilObjectException
		    Return -1
		  End
		  
		  If date2 = Nil Then 
		    Raise New NilObjectException
		    Return -1
		  End
		  
		  If Not OnRegionEnabled Then Return -1
		  
		  If date1.day= date2.day _
		  And date1.Month = date2.Month _
		  And date1.Year = date2.Year Then Return 0
		  
		  // End of trivial cases
		  
		  Var dtCurrent As DateTime, dtMax As DateTime, counter As Integer
		  Var blOK As Boolean
		  
		  If date1 > Date2 Then
		    dtMax = New DateTime(date1)
		    dtCurrent = New DateTime(Date2)
		  Else
		    dtMax = New DateTime(date2)
		    dtCurrent = New DateTime(Date1)
		  End
		  
		  
		  // Copy the relevant items ClosurePeriod
		  Var cp() As ClosurePeriod 
		  For r As Integer = 0 To Me.Regions.LastIndex
		    If Me.Regions(r).Enabled Then
		      AddClosurePeriods(date1, date2, cp, Me.Regions(r))
		    End If
		  next
		  
		  // Determines which days of the week are open (or closed) for all regions
		  Var WorkingWeekDays(6) As Boolean
		  For r As Integer = 0 To 6
		    WorkingWeekDays(r) = True
		  Next
		  
		  For r As Integer = 0 To Me.Regions.LastIndex
		    For i As Integer = 1 To 7
		      If Not Me.Regions(r).WorkingWeekDays.WorkingDay(i) Then WorkingWeekDays(i-1) = False // It is enough that in just one region that day is not a working day...
		    Next i
		  next r
		  
		  
		  dtCurrent = dtCurrent.SubtractInterval(0,0,1)
		  
		  Do 
		    
		    dtCurrent = dtCurrent.AddInterval (0,0,1)
		    
		    blOK = True // The day is initially considered a non-working day
		    
		    If WorkingWeekDays(dtCurrent.DayOfWeek - 1) = True Then 
		      counter = counter + 1
		      Continue Do 
		    end
		    
		    For c As Integer = 0 To cp.LastIndex
		      If dtCurrent >= cp(c).FirstDay And dtCurrent <= cp(c).LastDay Then 
		        counter = counter + 1
		        Continue Do // It's a closed period....
		      end
		    next 
		    
		    
		    For i As Integer = 0 To Me.Regions.LastIndex
		      
		      If Me.Regions(i).Enabled Then Continue
		      
		      If not Me.Regions(i).AnnualEventMatch(dtCurrent, True) Then 
		        counter = counter + 1
		        Continue Do // It's a closed period....
		      end
		      
		    Next i
		    
		  Loop Until dtCurrent >= dtMax
		  
		  Return counter
		  
		  
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsEvent(d as DateTime, OnlyDayOff as Boolean = False) As Boolean
		  for r as Integer = 0 to me.Regions.LastIndex
		    
		    For i As Integer = 0 To me.Regions(r).AnnualEvents.LastIndex
		      
		      if not me.Regions(r).AnnualEvents(i).DayOff and OnlyDayOff then Continue
		      if me.Regions(r).AnnualEvents(i).TestDate(d) then Return True
		      
		    Next
		    
		  next r
		  
		  Return False
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsWorkingDay(d as DateTime) As Boolean
		  for r as Integer = 0 to me.Regions.LastIndex
		    
		    If Not me.Regions(r).WorkingWeekDays.WorkingDay(d.DayOfWeek) Then Return False // It's never a working day
		    
		    For p As Integer = 0 To Me.Regions(r).ClosurePeriods.LastIndex
		      If d >=  Me.Regions(r).ClosurePeriods(p).FirstDay And d <=  Me.Regions(r).ClosurePeriods(p).LastDay Then Return False // It's as closed period
		    Next p
		    
		    For i As Integer = 0 To me.Regions(r).AnnualEvents.LastIndex
		      
		      if not me.Regions(r).AnnualEvents(i).DayOff then Continue // No need to calculate the date, it's not a holiday
		      if me.Regions(r).AnnualEvents(i).TestDate(d) then Return False
		      
		    Next i
		    
		  next r
		  
		  Return True
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsWorkingWeekDay(day as integer) As Boolean
		  if not OnRegionEnabled then Return False
		  
		  // The date must be a working week day for ALL regions.
		  
		  For i As Integer = 0 To Me.Regions.LastIndex
		    
		    if not me.Regions(i).Enabled then Continue
		    
		    // Therefore, it is enough for just one of them to not match.
		    if not me.Regions(i).WorkingWeekDays.WorkingDay(day) then Return false
		    
		  Next
		  
		  Return True
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ListFromDefinitions(Starting as DateTime, Ending as DateTime, IncludeSaturday as Boolean = True, IncludeSunday as Boolean = True, IncludeMonday as Boolean = True, IncludeTuesday as Boolean = True, IncludeWednesday as Boolean = True, IncludeThursday as Boolean = True, IncludeFriday as Boolean = True) As DateAndCaption()
		  If Starting = Nil Then 
		    Raise New NilObjectException
		    Return Nil
		  End
		  
		  If Ending = Nil Then 
		    Raise New NilObjectException
		    Return Nil
		  End
		  
		  Var list() As DateAndCaption
		  if not me.OnRegionEnabled then Return list // return empty array
		  
		  Var nStart As New date(starting.Year, starting.Month, starting.Day, 0, 0, 0)
		  Var nEnd As New date(ending.Year, ending.Month, ending.Day, 0, 0, 0)
		  
		  If nStart > nEnd Then 
		    Break // At design time, warns the developer
		    Return list // return empty array
		  end
		  
		  Var Valeur As DateAndCaption
		  
		  
		  For year As Integer = nStart.Year To nEnd.Year
		    
		    for r as integer = 0 to me.Regions.LastIndex
		      
		      if me.Regions(r).AnnualEvents.Count = 0 then Continue 
		      
		      For i As Integer = 0 To me.Regions(r).AnnualEvents.LastIndex
		        
		        Valeur =  me.Regions(r).AnnualEvents(i).Calculate(year)
		        
		        If valeur = Nil Then Continue
		        If valeur.DateValue < nStart Or valeur.DateValue > nEnd Then Continue
		        
		        If valeur.DateValue.DayOfWeek = 7 And Not IncludeSaturday Then Continue
		        If valeur.DateValue.DayOfWeek = 1 And Not IncludeSunday Then Continue
		        If valeur.DateValue.DayOfWeek = 2 And Not IncludeMonday Then Continue
		        If valeur.DateValue.DayOfWeek = 3 And Not IncludeTuesday Then Continue
		        If valeur.DateValue.DayOfWeek = 4 And Not IncludeWednesday  Then Continue
		        If valeur.DateValue.DayOfWeek = 6 And Not IncludeThursday Then Continue
		        If valeur.DateValue.DayOfWeek = 6 And Not IncludeFriday Then Continue
		        
		        AddInArray(list, New DateAndCaption(valeur))
		        
		      Next i
		      
		    next r
		    
		  Next year
		  
		  Return list
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ListFromDefinitions(Year as integer, IncludeSaturday as Boolean = True, IncludeSunday as Boolean = True, IncludeMonday as Boolean = True, IncludeTuesday as Boolean = True, IncludeWednesday as Boolean = True, IncludeThursday as Boolean = True, IncludeFriday as Boolean = True) As DateAndCaption()
		  Var list() As DateAndCaption
		  if not me.OnRegionEnabled then Return list // return empty array
		  
		  Var Valeur As DateAndCaption
		  
		  for r as Integer = 0 to me.Regions.LastIndex
		    
		    if me.Regions(r).AnnualEvents.Count = 0 then  continue
		    
		    For i As Integer = 0 To me.Regions(r).AnnualEvents.LastIndex
		      
		      Valeur = me.Regions(r).AnnualEvents(i).Calculate(year)
		      
		      If valeur = Nil Then Continue
		      
		      If valeur.DateValue.DayOfWeek = 7 And Not IncludeSaturday Then Continue
		      If valeur.DateValue.DayOfWeek = 1 And Not IncludeSunday Then Continue
		      If valeur.DateValue.DayOfWeek = 2 And Not IncludeMonday Then Continue
		      If valeur.DateValue.DayOfWeek = 3 And Not IncludeTuesday Then Continue
		      If valeur.DateValue.DayOfWeek = 4 And Not IncludeWednesday Then Continue
		      If valeur.DateValue.DayOfWeek = 6 And Not IncludeThursday Then Continue
		      If valeur.DateValue.DayOfWeek = 6 And Not IncludeFriday Then Continue
		      
		      AddInArray(list, New DateAndCaption(valeur))
		      
		    Next i
		    
		  next r
		  
		  Return list
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ListOfRegions() As Variant()
		  Dim r() As Variant
		  
		  If Me.Regions.Count = 0 Then return r() // Retun empty array
		  
		  For i As Integer = 0 To Me.Regions.LastIndex
		    r.Add me.Regions(i).Identifier
		  Next
		  
		  Return r
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Sub LoadClosurePeriods(regions() as DaysProcessingRegion, rs as RowSet, Encoding as TextEncoding = Nil)
		  If rs = Nil Then Exit Sub
		  If Regions.Count = 0 Then Exit Sub
		  
		  Var r1 as string, caption as string
		  
		  do until rs.AfterLastRow
		    
		    If Encoding = Nil Or Encoding = Encodings.UTF8 Then
		      r1 = rs.Column("region").StringValue.DefineEncoding(Encodings.UTF8).Lowercase.Trim
		      caption = rs.Column("caption").StringValue.DefineEncoding(Encodings.UTF8)
		    else
		      r1 = rs.Column("region").StringValue.DefineEncoding(Encoding).ConvertEncoding(Encodings.UTF8).DefineEncoding(Encodings.UTF8).Lowercase.Trim
		      caption = rs.Column("caption").StringValue.DefineEncoding(Encoding).ConvertEncoding(Encodings.UTF8).DefineEncoding(Encodings.UTF8)
		    end
		    
		    for r as Integer = 0 to Regions.LastIndex
		      
		      If r1 <> Regions(r).Identifier.StringValue.Lowercase.Trim Then continue
		      Regions(r).ClosurePeriods.Add new ClosurePeriod(rs.Column("firstday").DateTimeValue, rs.Column("lastday").DateTimeValue, caption)
		      
		    Next r
		    
		    rs.MoveToNextRow
		    
		  Loop
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Sub LoadEventsFromRowSet(Regions() as DaysProcessingRegion, rs as RowSet, Encoding as TextEncoding = Nil)
		  // if Encoding = nil, using UTF8
		  if Regions.Count = 0 then exit Sub
		  
		  Var DefinitionsArray() as AnnualEvent
		  Var label as string, identifier as String
		  
		  Var df As AnnualEventFix
		  Var de As AnnualEventEaster
		  Var deo As AnnualEventOrthodoxEaster
		  var dw as AnnualEventWeekDay
		  
		  Var CurrentRegion as DaysProcessingRegion
		  If rs = Nil Then exit sub
		  // using the field "day", for fix and easters definitions
		  
		  Do Until rs.AfterLastRow
		    
		    
		    If Encoding = Nil Or Encoding = Encodings.UTF8 Then // Converting all text data
		      label = rs.Column("caption").StringValue.DefineEncoding(Encoding).ConvertEncoding(Encodings.UTF8).DefineEncoding(Encodings.UTF8)
		      identifier = rs.Column("region").StringValue.DefineEncoding(Encoding).ConvertEncoding(Encodings.UTF8).DefineEncoding(Encodings.UTF8)
		    else
		      label = rs.Column("caption").StringValue.DefineEncoding(Encodings.UTF8)
		      identifier = rs.Column("region").StringValue.DefineEncoding(Encodings.UTF8)
		    end
		    
		    CurrentRegion = nil
		    
		    for i as Integer = 0 to Regions.LastIndex
		      if Regions(i).Identifier.StringValue.trim.Lowercase = identifier.trim.Lowercase then CurrentRegion = Regions(i)
		    next
		    
		    if CurrentRegion = nil then continue // No region to adding
		    
		    
		    Select Case rs.Column("definitiontype").StringValue.Uppercase
		      
		    Case "F" // Fix
		      
		      df = New AnnualEventFix(label, rs.Column("month").IntegerValue, rs.Column("day").IntegerValue, rs.Column("saturdaytofriday").BooleanValue, rs.Column("sundaytomonday").BooleanValue)
		      df.MondayIfSaturday = rs.Column("saturdaytomonday")
		      
		      df.AddDays = rs.Column("adddays").IntegerValue
		      df.NextWeekDay = rs.Column("nextweekday").IntegerValue
		      df.PreviousWeekDay = rs.Column("previousweekday").IntegerValue
		      
		      CurrentRegion.AnnualEvents.Add df
		      
		    Case "WD" // Weekday
		      
		      dw = new AnnualEventWeekDay(label, rs.Column("month").IntegerValue, rs.Column("weekday").IntegerValue, rs.Column("rank").IntegerValue)
		      
		      dw.AddDays = rs.Column("adddays").IntegerValue
		      dw.NextWeekDay = rs.Column("nextweekday").IntegerValue
		      dw.PreviousWeekDay = rs.Column("previousweekday").IntegerValue
		      
		      CurrentRegion.AnnualEvents.Add dw  
		      
		    Case  "E" // Easter
		      
		      de = New AnnualEventEaster(label, rs.Column("day").IntegerValue)
		      
		      CurrentRegion.AnnualEvents.Add de
		      
		      
		    case "EO" // Orthodox Easter 
		      
		      deo = New AnnualEventOrthodoxEaster(label, rs.Column("day").IntegerValue)
		      
		      CurrentRegion.AnnualEvents.Add deo
		      
		      
		    end
		    
		    CurrentRegion.AnnualEvents(CurrentRegion.AnnualEvents.LastIndex).CycleFirstYear = rs.Column("cyclefirstyear").IntegerValue
		    CurrentRegion.AnnualEvents(CurrentRegion.AnnualEvents.LastIndex).CycleYearDuration = rs.Column("cycleyearduration").IntegerValue
		    CurrentRegion.AnnualEvents(CurrentRegion.AnnualEvents.LastIndex).StartOfValidity = rs.Column("start").DateTimeValue
		    CurrentRegion.AnnualEvents(CurrentRegion.AnnualEvents.LastIndex).EndOfValidity = rs.Column("end").DateTimeValue
		    CurrentRegion.AnnualEvents(CurrentRegion.AnnualEvents.LastIndex).DayOff = rs.Column("dayoff").BooleanValue
		    
		    rs.MoveToNextRow
		  Loop
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Sub LoadRegions(regions() as DaysProcessingRegion, rs as RowSet, encoding as TextEncoding = Nil)
		  If rs = Nil Then Exit Sub
		  
		  Var identifier As String
		  
		  Do Until rs.AfterLastRow
		    
		    If Encoding = Nil Or Encoding = Encodings.UTF8 Then
		      identifier = rs.Column("identifier").StringValue.DefineEncoding(Encodings.UTF8).Lowercase.Trim
		    Else
		      identifier = rs.Column("identifier").StringValue.DefineEncoding(Encoding).ConvertEncoding(Encodings.UTF8).DefineEncoding(Encodings.UTF8).Lowercase.Trim
		    End
		    
		    regions.Add new DaysProcessingRegion(identifier)
		    
		    rs.MoveToNextRow
		    
		  Loop
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Sub LoadWeekDaysFromRowSet(regions() as DaysProcessingRegion, rs as RowSet, Encoding as TextEncoding = Nil)
		  If rs = Nil Then Exit Sub
		  If Regions.Count = 0 Then Exit Sub
		  
		  Var r1 as string
		  
		  do until rs.AfterLastRow
		    
		    If Encoding = Nil Or Encoding = Encodings.UTF8 Then
		      r1 = rs.Column("region").StringValue.DefineEncoding(Encodings.UTF8).Lowercase.Trim
		    else
		      r1 = rs.Column("region").StringValue.DefineEncoding(Encoding).ConvertEncoding(Encodings.UTF8).DefineEncoding(Encodings.UTF8)
		    end
		    
		    for r as Integer = 0 to Regions.LastIndex
		      
		      If r1 <> Regions(r).Identifier.StringValue.Lowercase.Trim Then Continue
		      
		      Regions(r).WorkingWeekDays.WorkingSunday = rs.Column("sunday").BooleanValue
		      Regions(r).WorkingWeekDays.WorkingMonday = rs.Column("monday").BooleanValue
		      Regions(r).WorkingWeekDays.WorkingTuesday = rs.Column("tuesday").BooleanValue
		      Regions(r).WorkingWeekDays.WorkingWednesday = rs.Column("wednesday").BooleanValue
		      Regions(r).WorkingWeekDays.WorkingThursday = rs.Column("thursday").BooleanValue
		      Regions(r).WorkingWeekDays.WorkingFriday = rs.Column("friday").BooleanValue
		      Regions(r).WorkingWeekDays.WorkingSaturday = rs.Column("saturday").BooleanValue
		      
		    Next r
		    
		    rs.MoveToNextRow
		    
		  Loop
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function NextBusinessDay(d as datetime, number as integer = 1) As DateTime
		  // if number is 1, this means the immediately following business day. 
		  // If the number is 2, it means the second following business day, and so on.
		  
		  If d = Nil Then 
		    Raise New NilObjectException
		    Return Nil
		  End
		  
		  If number = 0 Then Return d // Trivial case
		  
		  if not OnRegionEnabled then Return d
		  
		  // Duplicating the date and using a counter
		  Var dt As New DateTime(d), counter As Integer
		  Var blOK as Boolean
		  
		  Var cp() As ClosurePeriod 
		  
		  If number > 0 Then
		    
		    Var d2 as new DateTime(3999, 12, 31)
		    
		    // Copy the relevant items ClosurePeriod
		    
		    For r As Integer = 0 To Me.Regions.LastIndex
		      If Me.Regions(r).Enabled Then
		        AddClosurePeriods(d, d2 , cp, Me.Regions(r))
		      End If
		    next
		    
		  Else
		    
		    Var d2 as new DateTime(1, 1, 1)
		    
		    For r As Integer = 0 To Me.Regions.LastIndex
		      If Me.Regions(r).Enabled Then
		        AddClosurePeriods(d2, d, cp, Me.Regions(r))
		      End If
		    next
		    
		  end
		  
		  // Determines which days of the week are open (or closed) for all regions
		  Var WorkingWeekDays(6) As Boolean
		  For r As Integer = 0 To 6
		    WorkingWeekDays(r) = True
		  Next
		  
		  For r As Integer = 0 To Me.Regions.LastIndex
		    For i As Integer = 1 To 7
		      If Not Me.Regions(r).WorkingWeekDays.WorkingDay(i) Then WorkingWeekDays(i-1) = False // It is enough that in just one region that day is not a working day...
		    Next i
		  next r
		  
		  Do 
		    
		    blOK = True
		    
		    If number > 0 Then
		      dt = dt.AddInterval (0,0,1)
		    Else
		      dt = dt.SubtractInterval (0,0,1)
		    End
		    
		    if WorkingWeekDays(dt.DayOfWeek - 1) = False then Continue do 
		    
		    For c As Integer = 0 To cp.LastIndex
		      If dt >= cp(c).FirstDay And dt <= cp(c).LastDay Then Continue Do // It's a closed period....
		    next 
		    
		    for r as Integer = 0 to me.Regions.LastIndex
		      
		      if not me.Regions(r).Enabled then Continue for
		      
		      If Me.Regions(r).AnnualEventMatch(dt, True) Then continue do
		      
		    Next r
		    
		    counter = counter + 1
		    
		  Loop Until counter = abs(number)
		  
		  Return dt
		  
		  
		  
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function NonWorkingDays(starting as DateTime, ending as DateTime) As DateTime()
		  If Starting = Nil Then 
		    Raise New NilObjectException
		    Return Nil
		  End
		  
		  If Ending = Nil Then 
		    Raise New NilObjectException
		    Return Nil
		  End
		  
		  Var list() As DateTime
		  if not OnRegionEnabled then Return list
		  
		  Var nStart As New DateTime(starting.Year, starting.Month, starting.day)
		  Var nEnd As New DateTime(ending.Year, ending.Month, ending.day)
		  
		  If nStart > nEnd Then Return list
		  
		  // Copy the relevant items ClosurePeriod
		  Var cp() As ClosurePeriod 
		  For r As Integer = 0 To Me.Regions.LastIndex
		    If Me.Regions(r).Enabled Then
		      AddClosurePeriods(starting, ending, cp, Me.Regions(r))
		    End If
		  next
		  
		  // Determines which days of the week are open (or closed) for all regions
		  Var WorkingWeekDays(6) As Boolean
		  For r As Integer = 0 To 6
		    WorkingWeekDays(r) = True
		  Next
		  
		  For r As Integer = 0 To Me.Regions.LastIndex
		    For i As Integer = 1 To 7
		      If Not Me.Regions(r).WorkingWeekDays.WorkingDay(i) Then WorkingWeekDays(i-1) = False // It is enough that in just one region that day is not a working day...
		    Next i
		  next r
		  
		  Var d As DateTime = nStart.SubtractInterval(0,0,1)
		  Var blOK As Boolean
		  
		  Do Until d = nEnd
		    
		    d = d.AddInterval(0,0,1)
		    blOK = False
		    
		    For i As Integer = 0 To Me.Regions.LastIndex
		      
		      If not Me.Regions(i).WorkingWeekDays.WorkingDay(d.DayOfWeek) Then 
		        blOK = True
		        Exit For
		      End
		      
		      If Me.Regions(i).AnnualEventMatch(d, True) Then 
		        blOK = True
		        Exit For
		      End
		      
		    Next 
		    
		    if blOK then list.Add New DateTime(d)
		    
		  Loop
		  
		  Return list
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function NonWorkingDays(year as Integer) As DateTime()
		  Var d As New DateTime(year, 1, 1, 0, 0, 0)
		  Var f As New DateTime(year, 12, 31, 0, 0, 0)
		  
		  Return NonWorkingDays(d,f)
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function OnRegionEnabled() As Boolean
		  
		  if me.Regions.Count = 0 then Return False
		  
		  For i As Integer = 0 To Me.Regions.LastIndex
		    If Me.Regions(i).Enabled Then Return True
		  Next i
		  
		  Return False
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function PreviousBusinessDay(d as datetime, number as integer = 1) As DateTime
		  // if number is 1, this means the immediately following business day. 
		  // If the number is 2, it means the second following business day, and so on.
		  
		  If d = Nil Then 
		    Raise New NilObjectException
		    Return Nil
		  End
		  
		  if not OnRegionEnabled then Return d
		  
		  If number = 0 Then Return d // Trivial case
		  
		  number = number * -1
		  
		  Return NextBusinessDay(d, number)
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		Regions() As DaysProcessingRegion
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
	#tag EndViewBehavior
End Class
#tag EndClass
