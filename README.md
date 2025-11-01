# pbDateComputaion
Calculate holidays, closures, and working days with Xojo.

The repository is provided with an example project, allowing you to show how to calculate event dates and how to calculate working days, the next working day after a date, etc.

## Features
The features consist of 8 classes.
### 4 classes for calculating annual events, such as holidays and commemorations:
- AnnualEventFix, for events whose date is calculated using a fixed month and day (Independence Day, Christmas, etc.) 
- AnnualEventEaster and AnnualEventOrthodoxEaster, for events whose date is calculated based on Easter (Mardi Gras, Ash Wednesday, etc.) 
- AnnualEventWeekDay, for events whose date is calculated based on a fixed month and the week number of a weekday (Labor Day, Thanksgiving, etc.) 

These 4 classes implement the **AnnualEvent** interface, allowing them to be used in the same array, for example.

### A class to indicate periods of closure or holidays, for the class DaysProcessingRegion

- ClosingPeriod

### A class to calculate, for the same region, the working days, closed days, etc. according to holidays, closure periods and working days of the week.

- DaysProcessingRegion

### A class to work with multiple regions, which do not have the same public holidays or vacation periods, for example.

- DaysProcessingMultiRegion

### A class that is used to return the result of certain methods

- DateAndCaption

For example, you can request a list of all public holidays for a given year. An array of this type will be returned, specifying the date and name of the event for each row.

*A final class (dprWorkingWeekDays), which is only there to be a property of the DaysProcessingRegion class, is also provided, but is not designed to be used on its own.*

# The 4 classes that implement the AnnualEvent interface
 ## Common methods and properties (implemented by the interface)
Note: An interface defines only methods and not properties. "Properties" are a pair of overloaded methods that simulate real properties.

    Public Sub Caption(assigns s as string) 
    Public Function Caption() As string

#### Caption (propertie)
The name of the event, for example Christmas, Easter, presidential election, etc.
#### DayOff (propertie)
