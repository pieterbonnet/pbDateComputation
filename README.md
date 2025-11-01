# pbDateComputation
Calculate holidays, closures, and working days with Xojo.

The repository is provided with an example project, allowing you to show how to calculate event dates and how to calculate working days, the next working day after a date, etc.
# Features
The features consist of 8 classes.
### 4 classes for calculating annual events, such as holidays and commemorations:
- AnnualEventFix
- AnnualEventEaster and AnnualEventOrthodoxEaster,
- AnnualEventWeekDay

These 4 classes implement the **AnnualEvent** interface, allowing them to be used in the same array, for example.

#### A class to indicate periods of closure or holidays, for the class DaysProcessingRegion

- ClosingPeriod

#### A class to calculate, for the same region, the working days, closed days, etc. according to holidays, closure periods and working days of the week.

- DaysProcessingRegion

#### A class to work with multiple regions, which do not have the same public holidays or vacation periods, for example.

- DaysProcessingMultiRegion

### A class that is used to return the result of certain methods

- DateAndCaption

For example, you can request a list of all public holidays for a given year. An array of this type will be returned, specifying the date and name of the event for each row.

*A final class (dprWorkingWeekDays), which is only there to be a property of the DaysProcessingRegion class, is also provided, but is not designed to be used on its own.*

# The 4 classes that implement the AnnualEvent interface
 These classes are used to manipulate "annual events", such as holidays, celebrations, or events where calculating the date is a bit complicated (like "The Tuesday after the first Monday of November, for the U.S. Election Day").

| Class|For                       
|----------------|------------------------------- |
|AnnualEventFix  |for events whose date is calculated using a fixed month and day (Independence Day, Christmas, etc.)          |
|AnnualEventEaster |for events whose date is calculated based on Easter (Mardi Gras, Ash Wednesday, etc.)                      |
|AnnualEventOrthodoxEaster         |for events whose date is calculated based on Orthodox Easter 
|AnnualEventWeekDay|for events whose date is calculated based on a fixed month and the week number of a weekday (Labor Day, Thanksgiving, etc.) 


 ## Common methods and properties (implemented by the interface)
Note: An interface defines only methods and not properties. "Properties" are a pair of overloaded methods that simulate real properties.

    Public Sub Caption(assigns s as string) 
    Public Function Caption() As string

### Common properties
#### Caption 
Type : String
Defines the caption of the event, for example Christmas, Easter, presidential election, etc.
#### DayOff 
Type : Boolean
Defines if it is a non-working day or not (for non-working day the value is true)
#### StartOfValidity, EndOfValidity
Type : DateTime
Defines the dates on which the event occurs. For example, we can imagine a new public holiday being created starting next year (or an existing public holiday being discontinued).

The default dates are from January 1st of year 1 to December 31st of year 3999.
#### CycleFirstYear, CycleYearDuration
Type : Integer
An event may not occur every year. For example, an election.

If this is the case, you must define the first year and the cycle duration.

**Note :** that if the CycleYearDuration property is less than 2, these parameters are ignored.

The default values ​​are 1900 for CycleFirstYear and 1 for CycleYearDuration.

For example, the cycle for the US presidential election can be defined as follows: 

    Var PresidentialElection As New AnnualEventWeekDay("U.S. Presidential Election",11,2,1) // The first monday of november
    PresidentialElection.NextWeekDay = 3 // Shift to the next tuesday (so the next tuesday after the first monday)
    PresidentialElection.CycleFirstYear = 1792  // The event is not every year
    PresidentialElection.CycleYearDuration = 4

### Common methods

#### DateValue(Year as integer) as DateTime
Returns the date of the event for the specified year.

If the event did not occur that year, Nil is returned.
The reasons for a Nil value could be that the year falls outside the event's validity period or that the year does not correspond to the cycle.
Hours, minutes, and seconds are all set to 0.

    Var GoodFriday As New AnnualEventEaster("Good Friday",-2)
     // Determine the date for the year 2025
    Var Value2025 as Datetime = GoodFriday.DateValue(2025)


#### Calculate(Year as integer) as DateAndCaption
Returns a DateAndCaption object for the event for the specified year.
The object's date property contains the event date, and the Caption property replicates the event's Caption property.

If the event did not occur that year, Nil is returned.
The reasons for a Nil value could be that the year falls outside the event's validity period or that the year does not correspond to the cycle.
Hours, minutes, and seconds of the Date propertie of the object are all set to 0.

#### TestDate(d as DateTime) as Boolean
Determines if the event occurs on the tested date.
Hours, minutes, and seconds are ignored.

     Var GoodFriday As New AnnualEventEaster("Good Friday",-2)
     Var t as new DateTime(2025,4,18)
     if GoodFriday.TestDate(t) then 
        MessageBox "It's the date of the Good Friday !"
     else
       MessageBox "It's NOT the date of the Good Friday !"
     end
#### DbRow(RegionIdentifier as Variant = Nil, encoding as TextEncoding = nil) as DatabaseRow
Returns a DatabaseRow object.
This allows you to quickly save an event definition to a database.
Example :

    Var LaborDayEvent as New AnnualEventWeekDay("Labor Day",9, 2, 1) // September, the first monday
    Call MyDatabase.AddRow "tdays", LaborDayEvent.dbrow

You can specify the RegionIdentifer for the event if you are working with multiple regions. For example, the date of the commemoration of the abolition of slavery is not the same in the different French regions concerned. This allows you to filter them in the database.

You can also specify a specific encoding, particularly for databases that do not support UTF-8. By default, it is UTF-8.

Below is the SQL (SQLite) code to create a table to store event definitions.
The DaysProcessingRegion class provides a method to load events from a RowSet.

    CREATE TABLE "tdays" (
	"id_date"	INTEGER NOT NULL,
	"start"	DATE NOT NULL DEFAULT '1-1-1',
	"end"	DATE NOT NULL DEFAULT '3999-12-31',
	"caption"	TEXT,
	"region"	TEXT,
	"definitiontype"	TEXT NOT NULL,
	"dayoff"	BOOLEAN NOT NULL DEFAULT 1,
	"cyclefirstyear"	INTEGER NOT NULL DEFAULT 1,
	"cycleyearduration"	INTEGER NOT NULL DEFAULT 1,
	"month"	INTEGER NOT NULL DEFAULT 0,
	"day"	INTEGER NOT NULL DEFAULT 0,
	"weekday"	INTEGER NOT NULL DEFAULT 0,
	"rank"	INTEGER NOT NULL DEFAULT 0,
	"adddays"	INTEGER NOT NULL DEFAULT 0,
	"nextweekday"	INTEGER NOT NULL DEFAULT 0,
	"previousweekday"	INTEGER NOT NULL DEFAULT 0,
	"saturdaytofriday"	BOOLEAN NOT NULL DEFAULT 0,
	"sundaytomonday"	BOOLEAN NOT NULL DEFAULT 0,
	"saturdaytomonday"	BOOLEAN NOT NULL DEFAULT 0,
	"fingerprint"	TEXT NOT NULL,
	"alwaysdayshift"	BOOLEAN NOT NULL DEFAULT 0,
	PRIMARY KEY("id_date" AUTOINCREMENT)
	);
### FingerPrint() as string
Returns a value calculated based on the nature of the event, its parameters, and its validity dates. 
The caption and the shift parameters properties is not part of the fingerprint.

    Var VictoriaDay As New AnnualEventFix("Victoria Day ou Fête des patriotes",5,25) // May, 25th
    VictoriaDay.PreviousWeekDay = 2 // Previous Monday, if this not a monday
    MessageBox dFixedDay.FingerPrint

This code return "F05;25;0|0001-01-01|3999-12-31"
### DefinitionObject() as variant
Returns the object as a variant

Example :
AnnualEvents() is a array of AnnualEvent

		Var dFix as AnnualEventFix 
		Var dEaster as AnnualEventEaster
		Var dEasterOrthodox as AnnualEventOrthodoxEaster
		Var dWeekDay as AnnualEventWeekDay
    For i as integer = 0 to AnnualEvents.LastIndex
    
    	Select Case AnnualEvents(i)
    	
    	Case IsA AnnualEventFix 
      
    			dFix = AnnualEvents(i).DefinitionObject
                // Change something in dFix 
    	Case IsA AnnualEventEaster
    	  
    	   dEaster = AnnualEvents(i).DefinitionObject
           // Change something in dEaster
    	Case IsA AnnualEventOrthodoxEaster
          
		   dEasterOrthodox = AnnualEvents(i).DefinitionObject
           // Change something in dEasterOrthodox
       Case IsA AnnualEventWeekDay
     
	      dWeekDay =  AnnualEvents(i).DefinitionObject
          // Change something in dWeekDay
    
	End  Select
    
    Next i
### AnnualEventFix Class
In addition to the methods and properties common to the interface, here are the properties specific to this class.
#### Methods
#### Constructor(lCaption as string = "", lMonth as Integer = 1, lDay as Integer = 1, lFridayIfSaturday as Boolean = False, lMondayIfSunday as Boolean = False)
Creates a new object AnnualEventFix  whose properties are defined in the method arguments.

    Var dFixedDay As New AnnualEventFix("Día de la Independencia",9,16) // Mexico

#### Constructor(copy as AnnualEventFix)
Creates a new object AnnualEventFix whose properties are the same as the **copy** object passed (duplication).

    Var dFixedDay As New AnnualEventFix("Día de la Independencia",9,16) // Mexico
    Var dDupli as New AnnualEventFix(dFixedDay)

#### Constructor(a as AnnualEvent)
Creates a new object AnnualEventFix whose properties are the same as the **copy** object passed (duplication).
If the AnnualEvent is not a AnnualEventFix then a InvalidArgumentException is raised.

#### Properties 
#### Month, Day
Type : Integer

These two values ​​define the month and day of the event each year.

    Var dFix as AnnualEventFix
    // Bastille Day
    dFix.month = 7
    dFix.day = 14

#### MondayIfSunday
Type : Boolean

Determines whether the event should be postponed to Monday if it takes place on a Sunday.

    Var dFix as AnnualEventFix
    // Independence day
    dFix.month = 7
    dFix.day = 4
    dFix.MondayIfSunday = True

***Note :*** The property can be associated with FridayIfSaturday


#### FridayIfSaturday
Type : Boolean

Determines whether the event should be moved to the previous Friday if it takes place on a Saturday.

    Var dFix as AnnualEventFix
    // Independence day
    dFix.month = 7
    dFix.day = 4
    dFix.FridayIfSaturday = True

***Note :***  The property can be associated with MondayIfSunday

    Var dFix as AnnualEventFix
    // Independence day
    dFix.month = 7
    dFix.day = 4
    // Never a weekend :-)
    dFix.FridayIfSaturday = True
    dFix.MondayIfSunday = True
#### NextWeekDay, PreviousWeekDay, AlwaysNextWeekDay, AlwaysPreviousWeekDay
Type : Integer

Determines whether the day will be shifted to the corresponding day of the following week (or the previous one)

Use NextWeekDay/PreviousWeekDat if shifting is only necessary if the calculation does not fall on the requested day of the week.

Use AlwaysNextWeekDay/AlwaysPreviousWeekDay if shifting is ALWAYS necessary, even if the calculation falls on the correct day of the week (in this case, it's a one-week shift).
For example, in Canada, Victoria Day is ALWAYS shifted to the Monday before May 25th, even if May 25th is a Monday (This is the case for 2026).

The value is the same as the DayOfWeek property of the DateTime class. 
The value less that 1 indicates that there will be no offset
 
    // Victoria Day ou Fête des patriotes
    Var dFixedDay As New AnnualEventFix
    dFixedDay.Month = 5
    dFixedDay.Day = 25
    dFixedDay.AlwaysPreviousWeekDay = 2
#### AddDays
Type : Integer

Shifts the event by the specified number of days. Accepts negative values.
The value less that 0 indicates that there will be no offset

For example, imagine a holiday that falls on the day after February 28th.

    Var dFixedDay As New AnnualEventFix
    dFixedDay.Caption = "The birthday of people born on February 29th"
    dFixedDay.Month = 2
    dFixedDay.Day = 28
    dFixedDay.AddDays = 1
The calculation will give March 1st except for leap years where it will give February 29th.
### :warning: Important note on the combination and order of the shifts
The calculation function will first consider the MondayIfSunday and FridayIfSaturday properties, which are cumulative. If these properties are used, it will stop there and ignore any other offsets.
Otherwise, it will apply only ONE of the following conditions in the order below:

 1. AlwaysPreviousWeekDay 
 2. AlwaysNextWeekDay 
 3. PreviousWeekDay
 4.  NextWeekDay
 5. AddDays
