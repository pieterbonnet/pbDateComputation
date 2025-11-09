#tag Class
Protected Class RegionDatesWorked
	#tag Method, Flags = &h0
		Sub AddAnnualEvents(d() as AnnualEvent, NoDuplicates as Boolean = False)
		  If d.Count = 0 Then Exit Sub
		  
		  // Create a dictionnary to speed up duplicate searches
		  
		  Var dic As New Dictionary, dicCaption as new Dictionary
		  
		  If NoDuplicates Then 
		    
		    // Filling the dictionary with existing values
		    
		    If Me.AnnualEvents.Count > 0 Then
		      
		      For i As Integer = 0 To Me.AnnualEvents.LastIndex
		        dic.value(me.AnnualEvents(i).FingerPrint) = i
		        dicCaption.Value(me.AnnualEvents(i).Caption.Lowercase) = i
		      Next
		      
		    End If
		    
		  end
		  
		  For i As Integer = 0 To d.LastIndex
		    
		    if NoDuplicates then
		      If Not dic.Lookup(d(i).FingerPrint, "clear") = "clear" Then Continue // Already exist ! Skip this definition
		      If Not dicCaption.Lookup(d(i).Caption.Lowercase, "clear") = "clear" Then Continue // Already exist ! Skip this definition
		    end
		    
		    me.AnnualEvents.Add d(i)
		    
		    // Adding the new value in the dictionary
		    If NoDuplicates Then 
		      dic.value(Me.AnnualEvents(Me.AnnualEvents.LastIndex).FingerPrint) = Me.AnnualEvents.LastIndex
		      dicCaption.value(Me.AnnualEvents(Me.AnnualEvents.LastIndex).Caption.Lowercase) = Me.AnnualEvents.LastIndex
		    end
		  next
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AddAnnualEvents(NewEvent as AnnualEvent, NoDuplicates as Boolean = False)
		  // Create a dictionnary to speed up duplicate searches
		  
		  if me.AnnualEvents.Count > 0 and NoDuplicates then
		    
		    Var dic As New Dictionary
		    
		    // Filling the dictionary with existing values
		    
		    If Me.AnnualEvents.Count > 0 Then
		      
		      For i As Integer = 0 To Me.AnnualEvents.LastIndex
		        dic.value(me.AnnualEvents(i).FingerPrint) = i
		      Next
		      
		    End If
		    
		    If Not dic.Lookup(NewEvent, "clear") = "clear" Then exit sub // Already exist ! Skip this definition
		  end
		  
		  me.AnnualEvents.Add NewEvent
		  
		  
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

	#tag Method, Flags = &h21
		Private Function AffectedClosurePeriods(FirstDay as DateTime, LastDay as DateTime) As ClosurePeriod()
		  Var cp() As ClosurePeriod // Copy the relevant items ClosurePeriod
		  
		  If Me.ClosurePeriods.Count > 0 Then
		    
		    For c As Integer = 0 To Me.ClosurePeriods.LastIndex
		      
		      If FirstDay <= ClosurePeriods(c).LastDay And LastDay >= ClosurePeriods(c).FirstDay Then cp.Add ClosurePeriods(c)
		      
		    Next
		    
		  End
		  
		  Return cp
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function AnnualEventMatch(d as DateTime, OnlyIfDayOff as Boolean = False) As Boolean
		  If d = Nil Then 
		    Raise New NilObjectException
		    Return False
		  End
		  
		  
		  Var r As DateAndCaption
		  r = AnnualEventMatchDateAndCaption(d, OnlyIfDayOff)
		  Return r <> Nil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function AnnualEventMatchCaption(d as DateTime, OnlyIfDayOff as Boolean = False) As String
		  If d = Nil Then 
		    Raise New NilObjectException
		    Return ""
		  End
		  
		  Var r As DateAndCaption
		  r = AnnualEventMatchDateAndCaption(d, OnlyIfDayOff)
		  if r = Nil then Return "" else return r.Caption
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function AnnualEventMatchDateAndCaption(d as DateTime, OnlyIfDayOff as Boolean = False) As DateAndCaption
		  If d = Nil Then 
		    Raise New NilObjectException
		    Return Nil
		  End
		  
		  If Me.AnnualEvents.LastIndex = -1 Then Return nil
		  
		  For i As Integer = 0 To Me.AnnualEvents.LastIndex
		    
		    if OnlyIfDayOff and not me.AnnualEvents(i).DayOff then Continue
		    If d < Me.AnnualEvents(i).StartOfValidity Or d > Me.AnnualEvents(i).EndOfValidity Then Continue
		    If Me.AnnualEvents(i).TestDate(d) Then Return new DateAndCaption(d, me.AnnualEvents(i).Caption, me.Identifier)
		    
		  Next i
		  
		  Return Nil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function AnnualEventMatchObject(d as DateTime, OnlyIfDayOff as Boolean = False) As AnnualEvent
		  If d = Nil Then 
		    Raise New NilObjectException
		    Return Nil
		  End
		  
		  If Me.AnnualEvents.LastIndex = -1 Then Return Nil
		  
		  For i As Integer = 0 To Me.AnnualEvents.LastIndex
		    
		    If OnlyIfDayOff And Not Me.AnnualEvents(i).DayOff Then Continue
		    If d < Me.AnnualEvents(i).StartOfValidity Or d > Me.AnnualEvents(i).EndOfValidity Then Continue
		    If Me.AnnualEvents(i).TestDate(d) Then return Me.AnnualEvents(i)
		    
		  Next i
		  
		  Return Nil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function AnnualEventMatchRegion(d as DateTime, OnlyIfDayOff as Boolean = False) As Variant
		  If d = Nil Then 
		    Raise New NilObjectException
		    Return Nil
		  End
		  
		  Var r As DateAndCaption
		  r = AnnualEventMatchDateAndCaption(d, OnlyIfDayOff)
		  if r = Nil then Return Nil else return r.RegionIdentifier
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function AnnualEventsBelgium(Region as Belgium = Belgium.NationauxEnFrancais, PublicService as Boolean = True) As AnnualEvent()
		  Var df() As AnnualEvent
		  
		  // I use Belgium so as not to offend any Belgian language community.
		  
		  Select Case Region
		    
		  Case Belgium.NationauxEnFrancais
		    
		    df.Add New AnnualEventFix("Premier janvier", 1, 1)
		    df.Add New AnnualEventFix("Fête du travail", 5, 1)
		    df.Add New AnnualEventFix("Fête Nationale", 7, 21)
		    df.Add New AnnualEventFix("Assomption", 8, 15)
		    
		    df.Add New AnnualEventFix("Toussaint", 11, 1)
		    df.Add New AnnualEventFix("Armistice 1918", 11, 11)
		    
		    df.Add New AnnualEventFix("Noël", 12, 25)
		    df.Add New AnnualEventEaster("Lundi de Pâques", 1)
		    df.Add New AnnualEventEaster("Ascension", 39)
		    df.Add New AnnualEventEaster("Lundi de Pentecôte", 50)
		    
		  case Belgium.NationaleDagenInDeVlaamseTaal
		    
		    df.Add New AnnualEventFix("Nieuwjaarsdag", 1, 1) 
		    df.Add New AnnualEventFix("Feest van de Arbeid", 5, 1)
		    df.Add New AnnualEventFix("Feest van de Vlaamse Gemeenschap", 7, 11)
		    df.Add New AnnualEventFix("Nationale Feestdag", 7, 21)
		    df.Add New AnnualEventFix("Onze-Lieve-Vrouw Hemelvaart", 8, 15)
		    
		    df.Add New AnnualEventFix("Allerheiligen", 11, 1)
		    df.Add New AnnualEventFix("Wapenstilstand", 11, 11)
		    
		    df.Add New AnnualEventFix("Kerstmis", 12, 25)
		    
		    df.Add New AnnualEventEaster("Paasmaandag", 1)
		    df.Add New AnnualEventEaster("Onze-Lieve-Heer Hemelvaart", 39)
		    df.Add New AnnualEventEaster("Pinkstermaandag", 50)
		    
		  case Belgium.NationalfeiertageAufDeutsch
		    
		    df.Add New AnnualEventFix("Neujahr", 1, 1)
		    df.Add New AnnualEventFix("Tag der Arbeit", 5, 1)
		    df.Add New AnnualEventFix("Nationalfeiertag", 7, 21)
		    df.Add New AnnualEventFix("Mariä Himmelfahrt", 8, 15)
		    
		    df.Add New AnnualEventFix("Allerheiligen", 11, 1)
		    df.Add New AnnualEventFix("Waffenstillstand", 11, 11)
		    
		    df.Add New AnnualEventFix("Weihnachten", 12, 25)
		    df.Add New AnnualEventEaster("Ostermontag", 1)
		    df.Add New AnnualEventEaster("Christi Himmelfahrt", 39)
		    df.Add New AnnualEventEaster("Pfingstmontag", 50)
		    
		  Case Belgium.BruxellesCapitale
		    
		    // https://www.talent.brussels/fr/fonctionnaires/jours-feries-la-region-bruxelloise
		    
		    df.Add New AnnualEventFix("Premier janvier", 1, 1)
		    df.Add New AnnualEventFix("Fête du travail", 5, 1)
		    df.Add New AnnualEventFix("Fête Nationale", 7, 21)
		    df.Add New AnnualEventFix("Assomption", 8, 15)
		    
		    df.Add New AnnualEventFix("Toussaint", 11, 1)
		    df.Add New AnnualEventFix("Armistice 1918", 11, 11)
		    df.Add New AnnualEventFix("Noël", 12, 25)
		    
		    df.Add New AnnualEventEaster("Lundi de Pâques", 1)
		    df.Add New AnnualEventEaster("Ascension", 39)
		    df.Add New AnnualEventEaster("Lundi de Pentecôte", 50)
		    
		    
		    if PublicService then
		      df.Add New AnnualEventFix("Fête de l'Iris", 5, 8)
		      df.Add New AnnualEventFix("Fête de la Fédération Wallonie-Bruxelles", 9,27)
		      df.Add New AnnualEventFix("Fête des morts", 11, 2)
		      df.Add New AnnualEventFix("Jour de la Dynastie", 11, 15)
		      df.Add New AnnualEventFix("Deuxième jour de Noël", 12, 26)
		    end
		    
		  Case Belgium.vlaanderen
		    
		    // https://www.vlaanderen.be/intern/personeel/verloven-en-afwezigheden/feestdagen
		    
		    df.Add New AnnualEventFix("Nieuwjaarsdag", 1, 1) 
		    df.Add New AnnualEventFix("Feest van de Arbeid", 5, 1)
		    df.Add New AnnualEventFix("Feest van de Vlaamse Gemeenschap", 7, 11)
		    df.Add New AnnualEventFix("Nationale Feestdag", 7, 21)
		    df.Add New AnnualEventFix("Onze-Lieve-Vrouw Hemelvaart", 8, 15)
		    
		    df.Add New AnnualEventFix("Allerheiligen", 11, 1)
		    df.Add New AnnualEventFix("Wapenstilstand", 11, 11)
		    
		    df.Add New AnnualEventFix("Kerstmis", 12, 25)
		    
		    df.Add New AnnualEventEaster("Paasmaandag", 1)
		    df.Add New AnnualEventEaster("Onze-Lieve-Heer Hemelvaart", 39)
		    df.Add New AnnualEventEaster("Pinkstermaandag", 50)
		    
		    If PublicService Then
		      df.Add New AnnualEventFix("Allerzielen", 11, 2)
		      df.Add New AnnualEventFix("Koningsdag", 11, 15)
		      df.Add New AnnualEventFix("Tweede kerstdag", 12, 26)
		    End
		    
		    
		  Case Belgium.Wallonie
		    
		    df.Add New AnnualEventFix("Premier janvier", 1, 1)
		    df.Add New AnnualEventFix("Fête du travail", 5, 1)
		    df.Add New AnnualEventFix("Fête Nationale", 7, 21)
		    df.Add New AnnualEventFix("Assomption", 8, 15)
		    
		    df.Add New AnnualEventFix("Toussaint", 11, 1)
		    df.Add New AnnualEventFix("Armistice 1918", 11, 11)
		    df.Add New AnnualEventFix("Noël", 12, 25)
		    
		    df.Add New AnnualEventEaster("Lundi de Pâques", 1)
		    df.Add New AnnualEventEaster("Ascension", 39)
		    df.Add New AnnualEventEaster("Lundi de Pentecôte", 50)
		    
		    
		    If PublicService Then
		      df.Add New AnnualEventFix("Fête de la Fédération Wallonie-Bruxelles", 9,27)
		      
		      df.Add New AnnualEventFix("Fête des morts", 11, 2)
		      df.Add New AnnualEventFix("Jour de la Dynastie", 11, 15)
		      df.Add New AnnualEventFix("Deuxième jour de Noël", 12, 26)
		    end
		    
		    
		    df.Add new AnnualEventWeekDay("Fête de la Région wallone", 9, 1, 3) // 3ème dimanche de septembre
		    df(df.LastIndex).DayOff = False
		    
		  case Belgium.DeutschsprachigeGemeinschaft
		    
		    // https://de.wikipedia.org/wiki/Deutschsprachige_Gemeinschaft
		    
		    df.Add New AnnualEventFix("Neujahr", 1, 1)
		    df.Add New AnnualEventFix("Tag der Arbeit", 5, 1)
		    df.Add New AnnualEventFix("Nationalfeiertag", 7, 21)
		    df.Add New AnnualEventFix("Mariä Himmelfahrt", 8, 15)
		    
		    df.Add New AnnualEventFix("Allerheiligen", 11, 1)
		    df.Add New AnnualEventFix("Waffenstillstand", 11, 11)
		    
		    df.Add New AnnualEventFix("Weihnachten", 12, 25)
		    df.Add New AnnualEventEaster("Ostermontag", 1)
		    df.Add New AnnualEventEaster("Christi Himmelfahrt", 39)
		    df.Add New AnnualEventEaster("Pfingstmontag", 50)
		    
		    df.Add New AnnualEventFix("Tag der Deutschsprachigen Gemeinschaft", 11, 15)
		    
		  End
		  
		  Return df
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function AnnualEventsFrance(Region as France = France.General) As AnnualEvent()
		  Var df() As AnnualEvent
		  
		  
		  Select Case region
		    
		  Case France.General
		    
		    df.Add New AnnualEventFix("Jour de l'an", 1, 1)
		    df.Add New AnnualEventFix("Fête du travail", 5, 1)
		    df.Add New AnnualEventFix("Victoire 1945", 5, 8)
		    df.Add New AnnualEventFix("Fête Nationale", 7, 14)
		    df.Add New AnnualEventFix("Assomption", 8, 15)
		    df.Add New AnnualEventFix("Toussaint", 11, 1)
		    df.Add New AnnualEventFix("Armistice 1918", 11, 11)
		    df.Add New AnnualEventFix("Noël", 12, 25)
		    
		    df.Add New AnnualEventEaster("Lundi de Pâques", 1)
		    df.Add New AnnualEventEaster("Ascension", 39)
		    df.Add New AnnualEventEaster("Lundi de Pentecôte", 50)
		    
		  Case France.Guadeloupe
		    
		    df.Add New AnnualEventFix("Jour de l'an", 1, 1)
		    df.Add New AnnualEventFix("Fête du travail", 5, 1)
		    df.Add New AnnualEventFix("Victoire 1945", 5, 8)
		    df.Add New AnnualEventFix("Fête Nationale", 7, 14)
		    df.Add New AnnualEventFix("Assomption", 8, 15)
		    df.Add New AnnualEventFix("Toussaint", 11, 1)
		    df.Add New AnnualEventFix("Armistice 1918", 11, 11)
		    df.Add New AnnualEventFix("Noël", 12, 25)
		    
		    df.Add New AnnualEventEaster("Lundi de Pâques", 1)
		    df.Add New AnnualEventEaster("Ascension", 39)
		    df.Add New AnnualEventEaster("Lundi de Pentecôte", 50)
		    
		    df.Add New AnnualEventFix("Abolition de l'esclavage", 05, 27)
		    
		  case France.SaintMartin
		    
		    df.Add New AnnualEventFix("Jour de l'an", 1, 1)
		    df.Add New AnnualEventFix("Fête du travail", 5, 1)
		    df.Add New AnnualEventFix("Victoire 1945", 5, 8)
		    df.Add New AnnualEventFix("Fête Nationale", 7, 14)
		    df.Add New AnnualEventFix("Assomption", 8, 15)
		    df.Add New AnnualEventFix("Toussaint", 11, 1)
		    df.Add New AnnualEventFix("Armistice 1918", 11, 11)
		    df.Add New AnnualEventFix("Noël", 12, 25)
		    
		    df.Add New AnnualEventEaster("Lundi de Pâques", 1)
		    df.Add New AnnualEventEaster("Ascension", 39)
		    df.Add New AnnualEventEaster("Lundi de Pentecôte", 50)
		    
		    df.Add New AnnualEventFix("Abolition de l'esclavage", 05, 27)
		    
		  Case France.AlsaceMoselle
		    
		    df.Add New AnnualEventFix("Jour de l'an", 1, 1)
		    df.Add New AnnualEventFix("Fête du travail", 5, 1)
		    df.Add New AnnualEventFix("Victoire 1945", 5, 8)
		    df.Add New AnnualEventFix("Fête Nationale", 7, 14)
		    df.Add New AnnualEventFix("Assomption", 8, 15)
		    df.Add New AnnualEventFix("Toussaint", 11, 1)
		    df.Add New AnnualEventFix("Armistice 1918", 11, 11)
		    df.Add New AnnualEventFix("Noël", 12, 25)
		    
		    df.Add New AnnualEventEaster("Lundi de Pâques", 1)
		    df.Add New AnnualEventEaster("Ascension", 39)
		    df.Add New AnnualEventEaster("Lundi de Pentecôte", 50)
		    
		    
		    df.Add New AnnualEventEaster("Vendredi Saint", -2)
		    df.Add New AnnualEventFix("Lendemain Noël - St Étienne", 12, 26)
		    
		  Case France.Martinique
		    
		    df.Add New AnnualEventFix("Jour de l'an", 1, 1)
		    df.Add New AnnualEventFix("Fête du travail", 5, 1)
		    df.Add New AnnualEventFix("Victoire 1945", 5, 8)
		    df.Add New AnnualEventFix("Fête Nationale", 7, 14)
		    df.Add New AnnualEventFix("Assomption", 8, 15)
		    df.Add New AnnualEventFix("Toussaint", 11, 1)
		    df.Add New AnnualEventFix("Armistice 1918", 11, 11)
		    df.Add New AnnualEventFix("Noël", 12, 25)
		    
		    df.Add New AnnualEventEaster("Lundi de Pâques", 1)
		    df.Add New AnnualEventEaster("Ascension", 39)
		    df.Add New AnnualEventEaster("Lundi de Pentecôte", 50)
		    
		    df.Add New AnnualEventFix("Abolition de l'esclavage", 05, 22)
		    
		  Case France.Guyane
		    
		    df.Add New AnnualEventFix("Jour de l'an", 1, 1)
		    df.Add New AnnualEventFix("Fête du travail", 5, 1)
		    df.Add New AnnualEventFix("Victoire 1945", 5, 8)
		    df.Add New AnnualEventFix("Fête Nationale", 7, 14)
		    df.Add New AnnualEventFix("Assomption", 8, 15)
		    df.Add New AnnualEventFix("Toussaint", 11, 1)
		    df.Add New AnnualEventFix("Armistice 1918", 11, 11)
		    df.Add New AnnualEventFix("Noël", 12, 25)
		    
		    df.Add New AnnualEventEaster("Lundi de Pâques", 1)
		    df.Add New AnnualEventEaster("Ascension", 39)
		    df.Add New AnnualEventEaster("Lundi de Pentecôte", 50)
		    
		    df.Add New AnnualEventFix("Abolition de l'esclavage", 06, 10)
		    
		  Case France.Reunion
		    
		    df.Add New AnnualEventFix("Jour de l'an", 1, 1)
		    df.Add New AnnualEventFix("Fête du travail", 5, 1)
		    df.Add New AnnualEventFix("Victoire 1945", 5, 8)
		    df.Add New AnnualEventFix("Fête Nationale", 7, 14)
		    df.Add New AnnualEventFix("Assomption", 8, 15)
		    df.Add New AnnualEventFix("Toussaint", 11, 1)
		    df.Add New AnnualEventFix("Armistice 1918", 11, 11)
		    df.Add New AnnualEventFix("Noël", 12, 25)
		    
		    df.Add New AnnualEventEaster("Lundi de Pâques", 1)
		    df.Add New AnnualEventEaster("Ascension", 39)
		    df.Add New AnnualEventEaster("Lundi de Pentecôte", 50)
		    
		    df.Add New AnnualEventFix("Abolition de l'esclavage", 12, 20)
		    
		  Case France.Mayotte
		    
		    df.Add New AnnualEventFix("Jour de l'an", 1, 1)
		    df.Add New AnnualEventFix("Fête du travail", 5, 1)
		    df.Add New AnnualEventFix("Victoire 1945", 5, 8)
		    df.Add New AnnualEventFix("Fête Nationale", 7, 14)
		    df.Add New AnnualEventFix("Assomption", 8, 15)
		    df.Add New AnnualEventFix("Toussaint", 11, 1)
		    df.Add New AnnualEventFix("Armistice 1918", 11, 11)
		    df.Add New AnnualEventFix("Noël", 12, 25)
		    
		    df.Add New AnnualEventEaster("Lundi de Pâques", 1)
		    df.Add New AnnualEventEaster("Ascension", 39)
		    df.Add New AnnualEventEaster("Lundi de Pentecôte", 50)
		    
		    
		    df.Add New AnnualEventFix("Abolition de l'esclavage", 04, 27)
		    
		  Case France.SaintBarthelemy
		    
		    df.Add New AnnualEventFix("Jour de l'an", 1, 1)
		    df.Add New AnnualEventFix("Fête du travail", 5, 1)
		    df.Add New AnnualEventFix("Victoire 1945", 5, 8)
		    df.Add New AnnualEventFix("Fête Nationale", 7, 14)
		    df.Add New AnnualEventFix("Assomption", 8, 15)
		    df.Add New AnnualEventFix("Toussaint", 11, 1)
		    df.Add New AnnualEventFix("Armistice 1918", 11, 11)
		    df.Add New AnnualEventFix("Noël", 12, 25)
		    
		    df.Add New AnnualEventEaster("Lundi de Pâques", 1)
		    df.Add New AnnualEventEaster("Ascension", 39)
		    df.Add New AnnualEventEaster("Lundi de Pentecôte", 50)
		    
		    df.Add New AnnualEventFix("Abolition de l'esclavage", 10, 9)
		    
		  End
		  
		  Return df
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function AnnualEventsFromRowSet(rs as RowSet, Identifier as Variant = Nil, Encoding as TextEncoding = Nil) As AnnualEvent()
		  // if Encoding = nil, using UTF8
		  
		  Var EventsArray() as AnnualEvent
		  Var caption as string, region as String
		  
		  Var df As AnnualEventFix
		  Var de As AnnualEventEaster
		  Var deo As AnnualEventOrthodoxEaster
		  var dw as AnnualEventWeekDay
		  
		  Var id As String
		  if Identifier = nil Then id = "" else id = Identifier.StringValue.Lowercase.Trim
		  
		  If rs = Nil Then Return EventsArray
		  
		  // using the field "day", for fix and easters definitions
		  
		  Do Until rs.AfterLastRow
		    
		    
		    If Encoding <> Nil Then // Converting all text data
		      caption = rs.Column("caption").StringValue.DefineEncoding(Encoding).ConvertEncoding(Encodings.UTF8).DefineEncoding(Encodings.UTF8)
		      region = rs.Column("region").StringValue.DefineEncoding(Encoding).ConvertEncoding(Encodings.UTF8).DefineEncoding(Encodings.UTF8)
		    else
		      caption = rs.Column("caption").StringValue.DefineEncoding(Encodings.UTF8)
		      region = rs.Column("region").StringValue.DefineEncoding(Encodings.UTF8)
		    end
		    
		    If region.trim.Lowercase <> id Then 
		      rs.MoveToNextRow
		      Continue Do
		    End
		    
		    
		    Select Case rs.Column("definitiontype").StringValue
		      
		    Case "f" // Fix
		      
		      df = New AnnualEventFix(caption, rs.Column("month").IntegerValue, rs.Column("day").IntegerValue, rs.Column("saturdaytofriday").BooleanValue, rs.Column("sundaytomonday").BooleanValue)
		      df.MondayIfSaturday = rs.Column("saturdaytomonday")
		      
		      df.AddDays = rs.Column("adddays").IntegerValue
		      
		      Var nd As Integer =  rs.Column("nextweekday").IntegerValue
		      Var pd as Integer = rs.Column("previousweekday").IntegerValue
		      
		      if rs.Column("alwaysdayshift").BooleanValue then
		        df.AlwaysNextWeekDay = rs.Column("nextweekday").IntegerValue
		        df.AlwaysPreviousWeekDay = rs.Column("previousweekday").IntegerValue
		      else
		        df.NextWeekDay = rs.Column("nextweekday").IntegerValue
		        df.PreviousWeekDay = rs.Column("previousweekday").IntegerValue
		      end
		      EventsArray.Add df
		      
		    Case "WD" // Weekday
		      
		      dw = new AnnualEventWeekDay(caption, rs.Column("month").IntegerValue, rs.Column("weekday").IntegerValue, rs.Column("rank").IntegerValue)
		      
		      dw.AddDays = rs.Column("adddays").IntegerValue
		      dw.NextWeekDay = rs.Column("nextweekday").IntegerValue
		      dw.PreviousWeekDay = rs.Column("previousweekday").IntegerValue
		      
		      EventsArray.Add dw  
		      
		    Case  "E" // Easter
		      
		      de = New AnnualEventEaster(caption, rs.Column("day").IntegerValue)
		      
		      EventsArray.add de
		      
		      
		    case "EO" // Orthodox Easter 
		      
		      deo = New AnnualEventOrthodoxEaster(caption, rs.Column("day").IntegerValue)
		      
		      EventsArray.add deo
		      
		      
		    end
		    
		    EventsArray(EventsArray.LastIndex).CycleFirstYear = rs.Column("cyclefirstyear").IntegerValue
		    EventsArray(EventsArray.LastIndex).CycleYearDuration = rs.Column("cycleyearduration").IntegerValue
		    EventsArray(EventsArray.LastIndex).StartOfValidity = rs.Column("start").DateTimeValue
		    EventsArray(EventsArray.LastIndex).EndOfValidity = rs.Column("end").DateTimeValue
		    EventsArray(EventsArray.LastIndex).DayOff = rs.Column("dayoff").BooleanValue
		    
		    rs.MoveToNextRow
		  Loop
		  
		  Return EventsArray
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function AnnualEventsUSA(State as USA = USA.Federal) As AnnualEvent()
		  Var List() As AnnualEvent
		  
		  Select Case state
		    
		  Case usa.Federal
		    
		    // https://www.opm.gov/policy-data-oversight/pay-leave/federal-holidays/#url=Overview
		    
		    // The change from Sunday to Monday, or from Saturday to Friday, 
		    // was announced on the official website.
		    
		    List.Add New AnnualEventFix("New Year's Day", 1, 1,  True, True) // January, 1st
		    List.Add New AnnualEventFix("Juneteenth National Independence Day", 6,19,  True, True) // June, 19
		    List.Add New AnnualEventFix("Independence Day", 7,4,  True, True) // July, 4
		    List.Add New AnnualEventFix("Christmas",12,25,  True, True) // December, 25
		    List.Add New AnnualEventFix("Veterans Day", 11,11,  True, True) // November, 11
		    
		    
		    list.Add New AnnualEventWeekDay("Birthday of Martin Luther King, Jr.", 1,2,3) // January, Monday, The Third
		    list.Add New AnnualEventWeekDay("Washington's Birthday", 2,2,3) // Frebruary, Monday, The Third
		    List.Add New AnnualEventWeekDay("Memorial Day", 5, 2, 6) // May, the last monday
		    List.Add New AnnualEventWeekDay("Labor Day",9, 2, 1) // September, the first monday
		    List.Add New AnnualEventWeekDay("Columbus Day",10, 2, 2) // October, the second monday
		    List.Add New AnnualEventWeekDay("Thanksgiving",11, 5, 4) // November, the 4th thursday
		    
		  Case usa.Alabama // Alabama
		    
		    // https://personnel.alabama.gov/Downloads/StateHolidays.pdf
		    // https://admincode.legislature.state.al.us/administrative-code/670-X-12-.01
		    // No mention of a schedule change from Sunday to Monday, or from Saturday to Friday, 
		    // was found on this website (or I couldn't find it). 
		    
		    List.Add New AnnualEventFix("New Year's Day", 1, 1) // January, 1st
		    List.Add New AnnualEventWeekDay("Martin Luther King Day & Robert E. Lee’s Birthday",1,2,3)' Januray, third monday
		    List.Add new AnnualEventEaster("Mardi gras", -47)
		    List.Add New AnnualEventWeekDay("George Washington & Thomas Jefferson’s Birthday", 2, 2, 3) // February, the third monday
		    List.Add New AnnualEventWeekDay("Confederate Memorial Day",4,2,4) // April, the fourth monday
		    List.Add New AnnualEventWeekDay("National Memorial Day",5,2,6) // May, the last monday
		    List.Add New AnnualEventWeekDay("Jefferson Davis birthday",6,2,1) // June, the first monday
		    List.Add New AnnualEventFix("Independence Day", 7,4) // July, 4
		    List.Add New AnnualEventWeekDay("Labor Day",9, 2, 1) // September, the first monday
		    list.Add New AnnualEventWeekDay("Columbus Day & Fraternal Day & American Indian Heritage Day",10,2,2) // October, the second monday
		    List.Add New AnnualEventFix("Veterans Day", 11,11,  True, True) // November, 11
		    List.Add New AnnualEventWeekDay("Thanksgiving",11, 5, 4) // November, the 4th thursday
		    List.Add New AnnualEventFix("Christmas",12,25) // December, 25
		    
		  Case USA.Alaska // Alaska
		    
		    // https://library.alaska.gov/akholidays.html
		    // No mention of a schedule change from Sunday to Monday, or from Saturday to Friday, 
		    // was found on this website (or I couldn't find it). 
		    
		    list.Add New AnnualEventWeekDay("President's day", 2,2,3) // Frebruary, Monday, The Third
		    list.Add New AnnualEventWeekDay("Seward Day",3,2,6) // March, the last monday
		    List.Add New AnnualEventWeekDay("Memorial Day", 5, 2, 6) // May, the last monday
		    List.Add New AnnualEventWeekDay("Labor Day",9, 2, 1) // September, the first monday
		    List.Add New AnnualEventWeekDay("Thanksgiving",11, 5, 4) // November, the 4th thursday
		    
		    List.Add New AnnualEventFix("New Year's Day", 1, 1) // January, 1st
		    List.Add New AnnualEventFix("Independence Day", 7,4) // July, 4
		    List.Add New AnnualEventFix("Alaska day",10,18) // October, 18
		    List.Add New AnnualEventFix("Veterans Day", 11,11) // November, 
		    List.Add New AnnualEventFix("Christmas",12,25) // December, 25
		    
		  Case usa.Arkansas // Arkansas
		    
		    // https://law.justia.com/codes/arkansas/title-1/chapter-5/section-1-5-101/
		    // https://arkleg.state.ar.us/Home/FTPDocument?path=%2FACTS%2F2001%2FPublic%2FACT304.pdf
		    // https://arkleg.state.ar.us/Home/FTPDocument?path=%2FACTS%2F2017R%2FPublic%2FACT561.pdf
		    
		    // The change from Sunday to Monday, or from Saturday to Friday, 
		    // was announced in the law.
		    
		    list.Add New AnnualEventWeekDay("Dr. Martin Luther King Jr.'s Birthday", 1,2,3) // January, Monday, The Third
		    list.Add New AnnualEventWeekDay("Washington's Birthday and Daisy Gatson Bates Day", 2,2,3) // Frebruary, Monday, The Third
		    List.Add New AnnualEventWeekDay("Memorial Day", 5, 2, 6) // May, the last monday
		    List.Add New AnnualEventWeekDay("Labor Day",9, 2, 1) // September, the first monday
		    List.Add New AnnualEventWeekDay("Thanksgiving",11, 5, 4) // November, the 4th thursday
		    
		    
		    List.Add New AnnualEventFix("New Year's Day", 1, 1, True, True) // January, 1st
		    List.Add New AnnualEventFix("Independence Day", 7,4, True, True) // July, 4
		    List.Add New AnnualEventFix("Veterans Day", 11,11, True, True) // November, 11
		    
		    List.Add New AnnualEventFix("Christmas Eve",12,24, True, True) // December, 24
		    List.Add New AnnualEventFix("Christmas",12,25, True, True) // December, 25
		    
		    
		  Case USA.Arizona // Arizona
		    // https://hr.az.gov/2024-holiday-calendar
		    // The change from Sunday to Monday, or from Saturday to Friday, 
		    // was announced on the this website.
		    
		    
		    List.Add New AnnualEventFix("New Year's Day", 1, 1, True, True) // January, 1st
		    List.Add New AnnualEventFix("Independence Day", 7,4, True, True) // July, 4
		    List.Add New AnnualEventFix("Veterans Day", 11,11, True, True) // November, 11
		    List.Add New AnnualEventFix("Christmas",12,25, True, True) // December, 25
		    
		    
		    
		    list.Add New AnnualEventWeekDay("Martin Luther King, Jr. / Civil Rights", 1,2,3) // January, Monday, The Third
		    list.Add New AnnualEventWeekDay("President's day", 2,2,3) // Frebruary, Monday, The Third
		    List.Add New AnnualEventWeekDay("Memorial Day", 5, 2, 6) // May, the last monday
		    List.Add New AnnualEventWeekDay("Labor Day",9, 2, 1) // September, the first monday
		    List.add New AnnualEventWeekDay("Columbus Day",10, 2, 2) // October, the second monday
		    List.Add New AnnualEventWeekDay("Thanksgiving",11, 5, 4) // November, the 4th thursday
		    
		    
		    
		  Case USA.California // California
		    
		    // https://www.sos.ca.gov/state-holidays
		    // No mention of a schedule change from Sunday to Monday, or from Saturday to Friday, 
		    // was found on this website (or I couldn't find it). 
		    
		    List.Add New AnnualEventFix("New Year's Day", 1, 1, True, True) // January, 1st
		    List.Add New AnnualEventFix("Independence Day", 7,4, True, True) // July, 4
		    List.Add New AnnualEventFix("Veterans Day", 11,11, True, True) // November, 11
		    List.Add New AnnualEventFix("Christmas",12,25, True, True) // December, 25
		    List.Add New AnnualEventFix("Cesar Chavez Day",03,31) // October, 18
		    
		    
		    list.Add New AnnualEventWeekDay("Martin Luther King Jr. Day", 1,2,3) // January, Monday, The Third
		    list.Add New AnnualEventWeekDay("President's day", 2,2,3) // Frebruary, Monday, The Third
		    List.Add New AnnualEventWeekDay("Memorial Day", 5, 2, 6) // May, the last monday
		    List.Add New AnnualEventWeekDay("Labor Day",9, 2, 1) // September, the first monday
		    List.Add New AnnualEventWeekDay("Thanksgiving",11, 5, 4) // November, the 4th thursday
		    
		    var dw as New AnnualEventWeekDay("Day after thanksgiving",11,5,4) // // November, the 4th thursday
		    dw.AddDays = 1 // The day after...
		    List.Add dw
		    
		  Case USA.Colorado // Colorado
		    
		    // Colorado C.R.S. 24-11-101
		    // https://tax.colorado.gov/holidays
		    
		    List.Add New AnnualEventFix("New Year's Day", 1, 1, False, True) // January, 1st
		    List.Add New AnnualEventFix("Juneteenth National Independence Day", 6,19, False, True)
		    List.Add New AnnualEventFix("Independence Day", 7,4, False, True) // July, 4
		    List.Add New AnnualEventFix("Veterans Day", 11,11, False, True) // November, 11
		    List.Add New AnnualEventFix("Christmas",12,25, False, True) // December, 25
		    
		    list.Add New AnnualEventWeekDay("Martin Luther King Jr. Day", 1,2,3) // January, Monday, The Third
		    list.Add New AnnualEventWeekDay("Washington-Lincoln day", 2,2,3) // Frebruary, Monday, The Third
		    List.Add New AnnualEventWeekDay("Memorial Day", 5, 2, 6) // May, the last monday
		    List.Add New AnnualEventWeekDay("Labor Day",9, 2, 1) // September, the first monday
		    list.Add New AnnualEventWeekDay("Frances Xavier Cabrini day", 10,2,1) // October, the first monday
		    List.Add New AnnualEventWeekDay("Thanksgiving",11, 5, 4) // November, the 4th thursday
		    
		  case usa.Connecticut // Connecticut
		    
		    // https://portal.ct.gov/about/legal-state-holidays
		    
		    List.Add New AnnualEventFix("New Year's Day", 1, 1, True, True) // January, 1st
		    
		    Var df as New AnnualEventFix("Martin Luther King Jr. Day", 1,15) // The first monday after the 15
		    df.NextWeekDay = 2
		    List.Add df
		    
		    List.Add New AnnualEventFix("Lincoln's Birthday", 2, 12, True, True) // January, 1st
		    List.Add New AnnualEventFix("Juneteenth", 6, 19, true, true)
		    List.Add New AnnualEventFix("Independence Day", 7,4, True, True) // July, 4
		    List.Add New AnnualEventFix("Veterans Day", 11,11, True, True) // November, 11
		    List.Add New AnnualEventFix("Christmas",12,25, True, True) // December, 25
		    
		    list.Add New AnnualEventWeekDay("Washington's Birthday", 2,2,3) // Frebruary, Monday, The Third
		    List.Add New AnnualEventWeekDay("Memorial Day", 5, 2, 6) // May, the last monday
		    List.Add New AnnualEventWeekDay("Labor Day",9, 2, 1) // September, the first monday
		    List.Add New AnnualEventWeekDay("Columbus Day",10, 2, 2) // October, the second monday
		    List.Add New AnnualEventWeekDay("Thanksgiving",11, 5, 4) // November, the 4th thursday
		    
		  Case USA.Delaware // Delaware
		    
		    // https://delcode.delaware.gov/title1/c005/index.html
		    
		    List.Add New AnnualEventFix("New Year's Day", 1, 1, True, True) // January, 1st
		    List.Add New AnnualEventFix("Juneteenth", 6, 19, true, true)
		    List.Add New AnnualEventFix("Independence Day", 7,4, True, True) // July, 4
		    List.Add New AnnualEventFix("Veterans Day", 11,11, True, True) // November, 11
		    List.Add New AnnualEventFix("Christmas",12,25, True, True) // December, 25
		    
		    list.Add new AnnualEventEaster("Good friday", AnnualEventEaster.GoodFriday)
		    
		    list.Add New AnnualEventWeekDay("Martin Luther King, Jr.", 1,2,3) // January, Monday, The Third
		    List.Add New AnnualEventWeekDay("Labor Day",9, 2, 1) // September, the first monday
		    List.Add New AnnualEventWeekDay("Thanksgiving",11, 5, 4) // November, the 4th thursday
		    Var dw As New AnnualEventWeekDay("Day after thanksgiving",11,5,4) // // November, the 4th thursday
		    dw.AddDays = 1 // The day after...
		    List.Add dw
		    
		  case usa.DistrictOfColumbia
		    
		    // https://edpm.dc.gov/chapter/12/#section-1220
		    
		    Var df as New AnnualEventFix("Inauguration Day", 1, 20, False, True)
		    df.CycleFirstYear = 1849
		    df.CycleYearDuration = 4
		    
		    List.Add New AnnualEventFix("New Year's Day", 1, 1,  True, True) // January, 1st
		    List.Add New AnnualEventFix("Independence Day", 7,4,  True, True) // July, 4
		    List.Add New AnnualEventFix("Veterans Day", 11,11,  True, True) // November, 11
		    List.Add New AnnualEventFix("Christmas",12,25,  True, True) // December, 25
		    
		    list.Add New AnnualEventWeekDay("Birthday of Martin Luther King, Jr.", 1,2,3) // January, Monday, The Third
		    list.Add New AnnualEventWeekDay("Washington's Birthday", 2,2,3) // Frebruary, Monday, The Third
		    List.Add New AnnualEventWeekDay("Memorial Day", 5, 2, 6) // May, the last monday
		    List.Add New AnnualEventWeekDay("Labor Day",9, 2, 1) // September, the first monday
		    List.Add New AnnualEventWeekDay("Columbus Day",10, 2, 2) // October, the second monday
		    List.Add New AnnualEventWeekDay("Thanksgiving",11, 5, 4) // November, the 4th thursday
		    
		    
		  case usa.Florida 
		    
		  Case USA.Georgia 
		    
		  Case USA.Hawai
		    
		  case usa.Idaho
		    
		    
		    
		  end Select
		  
		  
		  
		  Return List
		  
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
		  Var nStart As New DateTime(starting.Year, starting.Month, starting.day,0,0,0)
		  Var nEnd As New DateTime(ending.Year, ending.Month, ending.day,0,0,0)
		  If nStart > nEnd Then Return list
		  
		  Var cp() as ClosurePeriod = me.AffectedClosurePeriods(nStart, nEnd) // Copy the relevant items ClosurePeriod
		  
		  
		  Var d As DateTime = nStart.SubtractInterval(0,0,1)
		  
		  Do Until d = nEnd
		    
		    d = d.AddInterval(0,0,1)
		    
		    //  Is this never a workday ?
		    if not me.WorkingWeekDays.WorkingDay(d.DayOfWeek) then Continue Do // Sufficient condition, no need to check anything else
		    
		    // Is this a Closure period ?
		    For i As Integer = 0 To cp.LastIndex
		      If d >= cp(i).FirstDay And d <= cp(i).LastDay Then Continue Do // Sufficient condition, no need to check anything else
		    Next
		    
		    // Is this a non-worked event ?
		    // This is the most time-consuming part to calculate, so it is calculated last only if necessary.
		    If AnnualEventMatch(d, true) Then Continue  Do
		    
		    list.Add New DateTime(d) // okay, we work
		    
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
		Function ClosurePeriodMatch(d as DateTime) As Boolean
		  Var d1 As New DateTime(d.Year, d.Month, d.day,0,0,0) // 00:00
		  
		  For p As Integer = 0 To Me.ClosurePeriods.LastIndex
		    
		    if me.ClosurePeriods(p).FirstDay <= d1 and me.ClosurePeriods(p).LastDay >= d1 then Return true
		    
		  Next
		  
		  Return False
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ClosurePeriodMatchCaption(d as DateTime) As string
		  Var d1 As New DateTime(d.Year, d.Month, d.day,0,0,0) // 00:00
		  
		  For p As Integer = 0 To Me.ClosurePeriods.LastIndex
		    
		    if me.ClosurePeriods(p).FirstDay <= d1 and me.ClosurePeriods(p).LastDay >= d1 then Return me.ClosurePeriods(p).Caption
		    
		  Next
		  
		  Return ""
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ClosurePeriodMatchObject(d as DateTime) As ClosurePeriod
		  Var d1 As New DateTime(d.Year, d.Month, d.day,0,0,0) // 00:00
		  
		  For p As Integer = 0 To Me.ClosurePeriods.LastIndex
		    
		    if me.ClosurePeriods(p).FirstDay <= d1 and me.ClosurePeriods(p).LastDay >= d1 then Return me.ClosurePeriods(p)
		    
		  Next
		  
		  Return nil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function ClosurePeriodsFromRowSet(rs as RowSet, Identifier as variant = nil, Encoding as TextEncoding = Nil) As ClosurePeriod()
		  Var id As String
		  if Identifier = nil Then id = "" else id = Identifier.StringValue.Lowercase.Trim
		  
		  Var cp() as ClosurePeriod
		  
		  If rs = Nil Then Return cp // Empty array
		  If rs.AfterLastRow Then Return cp // Empty array
		  Var r1 As String, caption As String
		  
		  Do Until rs.AfterLastRow
		    
		    If Encoding = Nil Or Encoding = Encodings.UTF8 Then
		      r1 = rs.Column("region").StringValue.DefineEncoding(Encodings.UTF8).Lowercase.Trim
		      caption = rs.Column("caption").StringValue.DefineEncoding(Encodings.UTF8)
		    Else
		      r1 = rs.Column("region").StringValue.DefineEncoding(Encoding).ConvertEncoding(Encodings.UTF8).DefineEncoding(Encodings.UTF8).Lowercase.Trim
		      caption = rs.Column("caption").StringValue.DefineEncoding(Encoding).ConvertEncoding(Encodings.UTF8).DefineEncoding(Encodings.UTF8)
		    End
		    
		    
		    If r1.Trim.Lowercase <> id Then 
		      rs.MoveToNextRow
		      Continue do
		    End
		    
		    
		    cp.Add New ClosurePeriod(rs.Column("firstday").DateTimeValue, rs.Column("lastday").DateTimeValue, caption)
		    
		    rs.MoveToNextRow
		  Loop
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor()
		  me.WorkingWeekDays = new dprWorkingWeekDays
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(copy as RegionDatesWorked, NewIdentifier as Variant = Nil)
		  Me.WorkingWeekDays = New dprWorkingWeekDays
		  
		  If NewIdentifier = Nil Then
		    Me.Identifier  = copy.Identifier
		  Else
		    me.Identifier = Identifier
		  End If
		  
		  For i As Integer = 1 To 7
		    Me.WorkingWeekDays.WorkingDay(i) = copy.WorkingWeekDays.WorkingDay(i)
		  Next
		  
		  Var df As AnnualEventFix
		  Var de As AnnualEventEaster
		  Var doe As AnnualEventOrthodoxEaster
		  var dwd as AnnualEventWeekDay
		  
		  If copy.AnnualEvents.Count > 0 Then
		    
		    
		    For i As Integer = 0 To copy.AnnualEvents.LastIndex
		      
		      select case copy.AnnualEvents(i) 
		        
		      Case IsA AnnualEventFix
		        
		        df = copy.AnnualEvents(i).DefinitionObject
		        me.AnnualEvents.Add new AnnualEventFix(df)
		        
		      Case IsA AnnualEventEaster
		        
		        de =  copy.AnnualEvents(i).DefinitionObject
		        me.AnnualEvents.Add new AnnualEventEaster(de)
		        
		      Case IsA AnnualEventOrthodoxEaster
		        
		        doe =  copy.AnnualEvents(i).DefinitionObject
		        Me.AnnualEvents.Add New AnnualEventOrthodoxEaster(doe)
		        
		      Case IsA AnnualEventWeekDay
		        
		        dwd =  copy.AnnualEvents(i).DefinitionObject
		        Me.AnnualEvents.Add New AnnualEventWeekDay(dwd)
		        
		      end
		      
		    Next i
		    
		  End If
		  
		  If copy.ClosurePeriods.LastIndex > 0 Then
		    
		    For p As Integer = 0 To copy.ClosurePeriods.LastIndex
		      me.ClosurePeriods.Add new ClosurePeriod(copy.ClosurePeriods(p))
		    Next p
		    
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(lIdentifier as Variant)
		  Me.WorkingWeekDays = New dprWorkingWeekDays
		  me.Identifier  = lIdentifier
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(lIdentifier as Variant, lDefinitions() as AnnualEvent)
		  Me.WorkingWeekDays = New dprWorkingWeekDays
		  Me.Identifier  = lIdentifier
		  me.AnnualEvents = lDefinitions
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CountBusinessDays(Date1 as datetime, Date2 as datetime) As Integer
		  // cas triviaux
		  
		  If Date1 = Nil Then 
		    Raise New NilObjectException
		    Return 0
		  End
		  
		  If date2 = Nil Then 
		    Raise New NilObjectException
		    Return 0
		  End
		  
		  If date1.day= date2.day _
		    And date1.Month = date2.Month _
		    And date1.Year = date2.Year Then 
		    
		    if IsWorkingDay(date1)  then Return 1 else return 0
		    
		  end
		  
		  Var dtCurrent As DateTime, dtMax As DateTime, counter As Integer
		  
		  If date1 > Date2 Then
		    dtMax = New DateTime(date1)
		    dtCurrent = New DateTime(Date2)
		  else
		    dtMax = New DateTime(date2)
		    dtCurrent = New DateTime(Date1)
		  End
		  
		  Var cp() As ClosurePeriod  = AffectedClosurePeriods(date1, date2) // Copy the relevant items ClosurePeriod
		  
		  
		  dtCurrent = dtCurrent.SubtractInterval(0,0,1)
		  
		  Do 
		    
		    dtCurrent = dtCurrent.AddInterval (0,0,1)
		    
		    //  Is this never a workday ?
		    If not me.WorkingWeekDays.WorkingDay(dtCurrent.DayOfWeek ) then Continue do
		    
		    // Is this a Closure period ?
		    For i As Integer = 0 To cp.LastIndex
		      If dtCurrent >= cp(i).FirstDay And dtCurrent <= cp(i).LastDay Then Continue Do
		    Next
		    
		    // Is this a non-worked event ?
		    // This is the most time-consuming part to calculate, so it's calculated last, only if necessary.
		    If Me.AnnualEventMatch(dtCurrent, True) Then Continue
		    
		    counter = counter + 1
		    
		  loop until dtCurrent >= dtMax
		  
		  Return counter
		  
		  
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CountNonWorkingDays(Date1 as datetime, Date2 as datetime) As Integer
		  // cas triviaux
		  
		  If Date1 = Nil Then 
		    Raise New NilObjectException
		    Return -1
		  End
		  
		  If date2 = Nil Then 
		    Raise New NilObjectException
		    Return -1
		  End
		  
		  If date1.day= date2.day _
		    And date1.Month = date2.Month _
		    And date1.Year = date2.Year Then 
		    
		    if IsWorkingDay(date1) then Return 0 else return 1
		    
		  end
		  
		  Var dtCurrent As DateTime, dtMax As DateTime, counter As Integer
		  
		  If date1 > Date2 Then
		    dtMax = New DateTime(date1)
		    dtCurrent = New DateTime(Date2)
		  else
		    dtMax = New DateTime(date2)
		    dtCurrent = New DateTime(Date1)
		  End
		  
		  Var cp() As ClosurePeriod  = AffectedClosurePeriods(date1, date2) // Copy the relevant items ClosurePeriod
		  
		  dtCurrent = dtCurrent.SubtractInterval(0,0,1)
		  
		  Do 
		    
		    dtCurrent = dtCurrent.AddInterval (0,0,1)
		    
		    //  Is this never a workday ?
		    If Not Me.WorkingWeekDays.WorkingDay(dtCurrent.DayOfWeek) Then
		      counter = counter + 1
		      Continue
		    end
		    
		    // Is this a Closure period ?
		    For i As Integer = 0 To cp.LastIndex
		      If dtCurrent >= cp(i).FirstDay And dtCurrent <= cp(i).LastDay Then 
		        counter = counter + 1
		        Continue do
		      end
		    Next
		    
		    // Is this a non-worked event ?
		    // This is the most time-consuming part to calculate, so it is calculated last only if necessary.
		    If AnnualEventMatch(dtCurrent, true) Then 
		      counter = counter + 1
		      Continue
		    End
		    
		  loop until dtCurrent >= dtMax
		  
		  Return counter
		  
		  
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsWorkingDay(d as DateTime) As Boolean
		  If Not Me.WorkingWeekDays.WorkingDay(d.DayOfWeek) Then Return False
		  
		  For p As Integer = 0 To Me.ClosurePeriods.LastIndex
		    If Me.ClosurePeriods(p).IsIncluded(d) Then Return False
		  Next p
		  
		  For i As Integer = 0 To Me.AnnualEvents.LastIndex
		    
		    if not (d >= me.AnnualEvents(i).StartOfValidity and d <= me.AnnualEvents(i).EndOfValidity) then Continue 
		    if not me.AnnualEvents(i).DayOff then Continue // No need to test, it's not a holiday
		    
		    // Is this a non-worked event ?
		    // This is the most time-consuming part to calculate, so it is calculated last only if necessary.
		    if me.AnnualEvents(i).TestDate(d) then Return False
		    
		  Next i
		  
		  Return True
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ListOfAnnualEvents(Starting as DateTime, Ending as DateTime, OnlyDayOff as boolean = False, IncludeSaturday as Boolean = True, IncludeSunday as Boolean = True, IncludeMonday as Boolean = True, IncludeTuesday as Boolean = True, IncludeWednesday as Boolean = True, IncludeThursday as Boolean = True, IncludeFriday as Boolean = True) As DateAndCaption()
		  Var list() As DateAndCaption
		  If Me.AnnualEvents.LastIndex = -1 Then Return list
		  
		  If Starting = Nil Then 
		    Raise New NilObjectException
		    Return Nil
		  End
		  
		  If Ending = Nil Then 
		    Raise New NilObjectException
		    Return Nil
		  End
		  
		  
		  Var nStart As New date(starting.Year, starting.Month, starting.Day, 0, 0, 0)
		  Var nEnd As New date(ending.Year, ending.Month, ending.Day, 0, 0, 0)
		  
		  If nStart > nEnd Then Return list // Empty
		  If nStart = nEnd Then
		    Var temp As DateAndCaption = Me.AnnualEventMatchDateAndCaption(nStart)
		    If temp <> Nil Then list.Add temp
		    Return list
		  End
		  
		  Var Valeur As DateAndCaption
		  
		  
		  For year As Integer = nStart.Year To nEnd.Year
		    
		    For i As Integer = 0 To Me.AnnualEvents.LastIndex
		      
		      if OnlyDayOff and not me.AnnualEvents(i).DayOff then Continue
		      
		      Valeur = Me.AnnualEvents(i).Calculate(year)
		      
		      If valeur = Nil Then Continue
		      If valeur.DateValue < nStart Or valeur.DateValue > nEnd Then Continue
		      
		      
		      If valeur.DateValue.DayOfWeek = 7 And Not IncludeSaturday Then Continue
		      If valeur.DateValue.DayOfWeek = 1 And Not IncludeSunday Then Continue
		      If valeur.DateValue.DayOfWeek = 2 And Not IncludeMonday Then Continue
		      If valeur.DateValue.DayOfWeek = 3 And Not IncludeTuesday Then Continue
		      If valeur.DateValue.DayOfWeek = 4 And Not IncludeWednesday Then Continue
		      If valeur.DateValue.DayOfWeek = 6 And Not IncludeThursday Then Continue
		      If valeur.DateValue.DayOfWeek = 6 And Not IncludeFriday Then Continue
		      
		      AddInArray(list, New DateAndCaption(valeur))
		      
		    Next
		    
		  Next
		  
		  Return list
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ListOfAnnualEvents(Year as integer, OnlyDayOff as boolean = False, IncludeSaturday as Boolean = True, IncludeSunday as Boolean = True, IncludeMonday as Boolean = True, IncludeTuesday as Boolean = True, IncludeWednesday as Boolean = True, IncludeThursday as Boolean = True, IncludeFriday as Boolean = True) As DateAndCaption()
		  Var list() As DateAndCaption
		  If Me.AnnualEvents.LastIndex = -1 Then Return list
		  
		  Var Valeur As DateAndCaption
		  Var nStart As New DateTime(Year,1,1,0,0,0)
		  Var nEnd as new DateTime(Year,12,31,23,59,59)
		  
		  
		  For i As Integer = 0 To Me.AnnualEvents.LastIndex
		    
		    if OnlyDayOff and not me.AnnualEvents(i).DayOff then Continue
		    
		    Valeur = Me.AnnualEvents(i).Calculate(year)
		    
		    If valeur = Nil Then Continue
		    If valeur.DateValue < nStart Or valeur.DateValue > nEnd Then Continue
		    
		    If valeur.DateValue.DayOfWeek = 7 And Not IncludeSaturday Then Continue
		    If valeur.DateValue.DayOfWeek = 1 And Not IncludeSunday Then Continue
		    If valeur.DateValue.DayOfWeek = 2 And Not IncludeMonday Then Continue
		    If valeur.DateValue.DayOfWeek = 3 And Not IncludeTuesday Then Continue
		    If valeur.DateValue.DayOfWeek = 4 And Not IncludeWednesday Then Continue
		    If valeur.DateValue.DayOfWeek = 6 And Not IncludeThursday Then Continue
		    If valeur.DateValue.DayOfWeek = 6 And Not IncludeFriday Then Continue
		    
		    AddInArray(list, New DateAndCaption(valeur))
		    
		  Next
		  
		  Return list
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadClosurePeriods(rs as RowSet, Encoding as TextEncoding = Nil)
		  If rs = Nil Then Exit Sub
		  if rs.AfterLastRow then exit sub
		  
		  Var cp() As ClosurePeriod = RegionDatesWorked.ClosurePeriodsFromRowSet(rs, Me.Identifier, Encoding)
		  
		  For p As Integer = 0 To cp.LastIndex
		    Me.ClosurePeriods.add cp(p)
		  next p
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadEventsFromRowSet(rs as RowSet, Encoding as TextEncoding = Nil)
		  if rs = nil then exit sub
		  if rs.AfterLastRow then exit sub
		  
		  Var ae() as AnnualEvent = RegionDatesWorked.AnnualEventsFromRowSet(rs, me.Identifier, encoding)
		  
		  For i As Integer = 0 To ae.LastIndex
		    me.AnnualEvents.Add ae(i)
		  next 
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadWeekDaysFromRowSet(rs as RowSet, encoding as TextEncoding = Nil)
		  If rs = Nil Then Exit Sub
		  if rs.AfterLastRow then exit sub
		  
		  Var r1 As String, id As String
		  
		  if me.Identifier = nil Then id = "" else id = Me.Identifier.StringValue.Lowercase.Trim
		  
		  do until rs.AfterLastRow
		    
		    if Encoding = Nil Or Encoding = Encodings.UTF8 then
		      r1 = rs.Column("region").StringValue.DefineEncoding(encodings.UTF8).Lowercase.Trim 
		    else
		      r1 = rs.Column("region").StringValue.DefineEncoding(Encoding).ConvertEncoding(Encodings.utf8).DefineEncoding(encodings.UTF8).Lowercase.Trim 
		    end
		    
		    If r1 = id Then 
		      
		      me.WorkingWeekDays.WorkingSunday = rs.Column("sunday").BooleanValue
		      Me.WorkingWeekDays.WorkingMonday = rs.Column("monday").BooleanValue
		      Me.WorkingWeekDays.WorkingTuesday = rs.Column("tuesday").BooleanValue
		      Me.WorkingWeekDays.WorkingWednesday = rs.Column("wednesday").BooleanValue
		      Me.WorkingWeekDays.WorkingThursday = rs.Column("thursday").BooleanValue
		      Me.WorkingWeekDays.WorkingFriday = rs.Column("friday").BooleanValue
		      Me.WorkingWeekDays.WorkingSaturday = rs.Column("saturday").BooleanValue
		      
		      exit sub
		      
		    end
		    
		  loop
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function NextBusinessDay(d as datetime, number as integer = 1) As DateTime
		  // if number is 1, this means the immediately following business day. 
		  // If the number is 2, it means the second following business day, and so on.
		  // If you need to add 6 business days, then number = 6
		  
		  If d = Nil Then 
		    Raise New NilObjectException
		    Return Nil
		  End
		  
		  If number = 0 Then Return d // Trivial case
		  
		  // Duplicating the date and using a counter
		  Var dt As New DateTime(d), counter As Integer
		  
		  Do 
		    
		    If number > 0 Then
		      dt = dt.AddInterval (0,0,1)
		    Else
		      dt = dt.SubtractInterval (0,0,1)
		    End
		    
		    //  Is this never a workday ?
		    if not me.WorkingWeekDays.WorkingDay(dt.DayOfWeek) then Continue // Skip
		    
		    // Is this a Closure period ?
		    For i As Integer = 0 To Me.ClosurePeriods.LastIndex
		      if dt >= Me.ClosurePeriods(i).FirstDay  and dt <= Me.ClosurePeriods(i).LastDay then continue // skip
		    Next
		    
		    // Is this a non-worked event ?
		    // This is the most time-consuming part to calculate, so it is calculated last only if necessary.
		    If Me.AnnualEventMatch(dt, True) Then Continue // Skip
		    
		    counter = counter + 1
		    
		  Loop Until counter = abs(number)
		  
		  Return dt
		  
		  
		  
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function NextNonWorkingDay(d as datetime, number as integer = 1) As DateTime
		  // if number is 1, this means the immediately following non working day. 
		  // If the number is 2, it means the second following non working day, and so on.
		  // If you need to add 6 non working days, then number = 6
		  
		  If d = Nil Then 
		    Raise New NilObjectException
		    Return Nil
		  End
		  
		  If number = 0 Then Return d // Trivial case
		  
		  // Duplicating the date and using a counter
		  Var dt As New DateTime(d), counter As Integer
		  
		  Do 
		    
		    If number > 0 Then
		      dt = dt.AddInterval (0,0,1)
		    Else
		      dt = dt.SubtractInterval (0,0,1)
		    End
		    
		    //  Is this never a workday ?
		    If Not Me.WorkingWeekDays.WorkingDay(dt.DayOfWeek) Then 
		      counter = counter + 1
		      Continue // Skip
		    end
		    
		    // Is this a Closure period ?
		    For i As Integer = 0 To Me.ClosurePeriods.LastIndex
		      If dt >= Me.ClosurePeriods(i).FirstDay  And dt <= Me.ClosurePeriods(i).LastDay Then 
		        counter = counter + 1
		        continue do // skip
		      end
		    Next
		    
		    // Is this a non-worked event ?
		    // This is the most time-consuming part to calculate, so it is calculated last only if necessary.
		    If Me.AnnualEventMatch(dt, True) Then counter = counter + 1
		    
		  Loop Until counter = abs(number)
		  
		  Return dt
		  
		  
		  
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function NonWorkingDayReasonCaption(d as DateTime) As String
		  If d = Nil Then 
		    Raise New NilObjectException
		    Return ""
		  End
		  
		  if not me.WorkingWeekDays.WorkingDay(d.DayOfWeek) then Return d.ToString("EEEE")
		  
		  For i As Integer = 0 To Me.ClosurePeriods.LastIndex
		    If Me.ClosurePeriods(i).IsIncluded(d) Then Return Me.ClosurePeriods(i).Caption
		  next
		  
		  
		  
		  For i As Integer = 0 To Me.AnnualEvents.LastIndex
		    
		    If Not Me.AnnualEvents(i).DayOff Then Continue
		    If d < Me.AnnualEvents(i).StartOfValidity Or d > Me.AnnualEvents(i).EndOfValidity Then Continue
		    If Me.AnnualEvents(i).TestDate(d) Then Return me.AnnualEvents(i).Caption
		    
		  Next i
		  
		  Return ""
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function NonWorkingDayReasonType(d as DateTime) As Integer
		  If d = Nil Then 
		    Raise New NilObjectException
		    Return 0
		  End
		  
		  if not me.WorkingWeekDays.WorkingDay(d.DayOfWeek) then Return 1
		  
		  For i As Integer = 0 To Me.ClosurePeriods.LastIndex
		    If Me.ClosurePeriods(i).IsIncluded(d) Then Return 2
		  next
		  
		  
		  
		  For i As Integer = 0 To Me.AnnualEvents.LastIndex
		    
		    If Not Me.AnnualEvents(i).DayOff Then Continue
		    If d < Me.AnnualEvents(i).StartOfValidity Or d > Me.AnnualEvents(i).EndOfValidity Then Continue
		    If Me.AnnualEvents(i).TestDate(d) Then Return 3
		    
		  Next i
		  
		  Return 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function NonWorkingDays(starting as DateTime, ending as DateTime) As DateTime()
		  Var list() As DateTime
		  Var nStart As New DateTime(starting.Year, starting.Month, starting.day)
		  Var nEnd As New DateTime(ending.Year, ending.Month, ending.day)
		  
		  If nStart > nEnd Then Return list
		  
		  
		  Var cp() As ClosurePeriod  = AffectedClosurePeriods(starting, ending) // Copy the relevant items ClosurePeriod
		  
		  Var d As DateTime =  nStart.SubtractInterval(0,0,1)
		  
		  Do Until d >= nEnd
		    
		    d = d.AddInterval(0,0,1)
		    
		    //  Is this never a workday ?
		    If Not Me.WorkingWeekDays.WorkingDay(d.DayOfWeek) Then  
		      list.Add new DateTime(d)
		      continue // Sufficient condition, no need to check anything else
		    end
		    
		    // Is this a Closure period ?
		    For i As Integer = 0 To cp.LastIndex
		      If d >= cp(i).FirstDay And d <= cp(i).LastDay Then 
		        list.Add New DateTime(d)
		        Continue Do // Sufficient condition, no need to check anything else
		      end
		    Next
		    
		    // Is this a non-worked event ?
		    // This is the most time-consuming part to calculate, so it is calculated last only if necessary.
		    
		    If AnnualEventMatch(d, True) Then 
		      list.Add New DateTime(d)
		    end
		    
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

	#tag Method, Flags = &h0
		Function NonWorkingDaysDateAndCaption(starting as DateTime, ending as DateTime) As DateAndCaption()
		  Var list() As DateAndCaption
		  Var nStart As New DateTime(starting.Year, starting.Month, starting.day)
		  Var nEnd As New DateTime(ending.Year, ending.Month, ending.day)
		  
		  If nStart > nEnd Then Return list
		  
		  
		  Var cp() As ClosurePeriod  = AffectedClosurePeriods(starting, ending) // Copy the relevant items ClosurePeriod
		  
		  Var d As DateTime =  nStart.SubtractInterval(0,0,1)
		  Var s as string
		  
		  Do Until d >= nEnd
		    
		    d = d.AddInterval(0,0,1)
		    
		    //  Is this never a workday ?
		    If Not Me.WorkingWeekDays.WorkingDay(d.DayOfWeek) Then  
		      list.Add new DateAndCaption(DateTime(d), d.ToString("EEEE"))
		      continue // Sufficient condition, no need to check anything else
		    end
		    
		    // Is this a Closure period ?
		    For i As Integer = 0 To cp.LastIndex
		      If d >= cp(i).FirstDay And d <= cp(i).LastDay Then 
		        list.Add new DateAndCaption(DateTime(d), cp(i).Caption)
		        Continue Do // Sufficient condition, no need to check anything else
		      end
		    Next
		    
		    // Is this a non-worked event ?
		    // This is the most time-consuming part to calculate, so it is calculated last only if necessary.
		    
		    s = AnnualEventMatchCaption(d, True)
		    if s <> "" then list.Add new DateAndCaption(DateTime(d), s)
		    
		    
		  Loop
		  
		  Return list
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function NonWorkingDaysDateAndCaption(year as Integer) As DateAndCaption()
		  Var d As New DateTime(year, 1, 1, 0, 0, 0)
		  Var f As New DateTime(year, 12, 31, 0, 0, 0)
		  
		  Return NonWorkingDaysDateAndCaption(d,f)
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function PreviousBusinessDay(d as datetime, number as integer = 1) As DateTime
		  // if number is 1, this means the immediately previous business day. 
		  // If the number is 2, it means the second following business day, and so on.
		  
		  If d = Nil Then 
		    Raise New NilObjectException
		    Return Nil
		  End
		  
		  If number = 0 Then Return d // Trivial case
		  number = number * -1
		  
		  Return NextBusinessDay(d, number)
		  
		  
		  
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function PreviousNonWorkingDay(d as datetime, number as integer = 1) As DateTime
		  // if number is 1, this means the immediately previous non working day. 
		  // If the number is 2, it means the second following non working day, and so on.
		  
		  If d = Nil Then 
		    Raise New NilObjectException
		    Return Nil
		  End
		  
		  If number = 0 Then Return d // Trivial case
		  number = number * -1
		  
		  Return NextNonWorkingDay(d, number)
		  
		  
		  
		  
		  
		  
		  
		  
		End Function
	#tag EndMethod


	#tag Note, Name = DB Table example
		CREATE TABLE "tdays" (
		"id_date"    INTEGER NOT NULL,
		"start"    DATE NOT NULL DEFAULT '1-1-1',
		"end"    DATE NOT NULL DEFAULT '2999-12-31',
		"label"    TEXT,
		"region"    TEXT,
		"definitiontype"    TEXT NOT NULL,
		"dayoff"    BOOLEAN NOT NULL DEFAULT 1,
		"cyclefirstyear"    INTEGER NOT NULL DEFAULT 1,
		"cycleyearduration"    INTEGER NOT NULL DEFAULT 1,
		"month"    INTEGER NOT NULL DEFAULT 0,
		"day"    INTEGER NOT NULL DEFAULT 0,
		"weekday"    INTEGER NOT NULL DEFAULT 0,
		"rank"    INTEGER NOT NULL DEFAULT 0,
		"adddays"    INTEGER NOT NULL DEFAULT 0,
		"nextweekday"    INTEGER NOT NULL DEFAULT 0,
		"previousweekday"    INTEGER NOT NULL DEFAULT 0,
		"saturdaytofriday"    BOOLEAN NOT NULL DEFAULT 0,
		"sundaytomonday"    BOOLEAN NOT NULL DEFAULT 0,
		"saturdaytomonday"    BOOLEAN NOT NULL DEFAULT 0,
		"fingerprint"    TEXT NOT NULL UNIQUE,
		PRIMARY KEY("id_date" AUTOINCREMENT)
		);
		
	#tag EndNote


	#tag Property, Flags = &h0
		AnnualEvents() As AnnualEvent
	#tag EndProperty

	#tag Property, Flags = &h0
		ClosurePeriods() As ClosurePeriod
	#tag EndProperty

	#tag Property, Flags = &h0
		Enabled As Boolean = True
	#tag EndProperty

	#tag Property, Flags = &h0
		Identifier As Variant
	#tag EndProperty

	#tag Property, Flags = &h0
		Tag As Variant
	#tag EndProperty

	#tag Property, Flags = &h0
		WorkingWeekDays As dprWorkingWeekDays
	#tag EndProperty


	#tag Enum, Name = Belgium, Type = Integer, Flags = &h0
		NationauxEnFrancais
		  NationaleDagenInDeVlaamseTaal
		  NationalfeiertageAufDeutsch
		  Vlaanderen
		  DeutschsprachigeGemeinschaft
		  BruxellesCapitale
		Wallonie
	#tag EndEnum

	#tag Enum, Name = France, Flags = &h0
		General
		  AlsaceMoselle
		  Guyane
		  Guadeloupe
		  Martinique
		  Reunion
		  SaintBarthelemy
		  SaintMartin
		Mayotte
	#tag EndEnum

	#tag Enum, Name = USA, Type = Integer, Flags = &h0
		Federal
		  Alabama
		  Alaska
		  Arkansas
		  Arizona
		  California
		  Colorado
		  Connecticut
		  Delaware
		  DistrictOfColumbia
		  Florida
		  Georgia
		  Hawai
		  Idaho
		  Illinois
		  Indiana
		  Iowa
		  Kansas
		  Kentucky
		  Louisiana
		  Maine
		  Maryland
		  Massachusetts
		  Michigan
		  Minnesota
		  Mississipi
		  Missouri
		  Montana
		  Nebraska
		  Nevada
		  NewHampshire
		  NewJersey
		  NewMexico
		  NewYork
		  NorthCarolina
		  NorthDakota
		  Ohio
		  Oklahoma
		  Orgon
		  Pennsylvania
		  RhodeIsland
		  SouthCarolina
		  SouthDakota
		  Tennessee
		  Texas
		  Utah
		  Vermont
		  Virginia
		  Washington
		  WestVirginia
		  Wisconsin
		Wyoming
	#tag EndEnum


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
			Name="Enabled"
			Visible=false
			Group="Behavior"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
