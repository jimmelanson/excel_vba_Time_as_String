# excel_vba_Time_as_String
A class module to let you treat time as a string.

There are some reasons why you may want to treat time as as string. For example, I created a large scheduling appliction in excel through VBA. There was about 30 users scheduling 400 people over seven different shifts. I had to not only enusre data was input correctly, but that various labour laws and collective agreement rules were applied to the people's break times.

The biggest challenge was having users enter things correctly. No matte what I did with the formatting of cells, they quite often did something that changed the formatting. That made every other method/procedure using an input time be put at risk. Eventually I just rewrote everything to use time as a string.

Benefits of evaluating time as string:

- Make sure you can process an input time value exactly as it was entered.
- Reformat time strings exactly as you require.
- Easily examine hours and minutes separately.
- Not have to worry about cells time-format being reformatted.
- Sometimes it is just about preference and what you are more comfortable with.

Here are the things you can do with this class library:
- Choose one of four formats for your output: 24-Hour, 12-Hour, 12-Hour with AM/PM, French format (23h00).
- Convert times to decimals.
- Add one time to another time (eg, 8:30 A + 3:30 = 12:00 P)
- Subtract one time from another (eg, 10:00 P - 02:15 = 07:45 P)
- Add hours or minutes to a time.
- Subtract hours or minutes from a time.
- Evaluate if a time value falls between two other time values.
- Add multiple time ranges together.

The example workbook has two worksheets that demostrate the different things you can. If you look at the VBA Editor there is a module with some more testing subroutines.

You can download the clsTimeString file to include in your project. Here is a list of the methods and properties.

<code>TimeFormat = [String]</code> Tells the class how to output the formatted time value.

<code>AMPM = Boolean</code> Tells the class to put "A" or "P" after the forrmatted time, but only if the 12-Hour format is selected.

<code>RoundTo = [integer]</code> Will force the formatted time value to be rounded to the nearest value provided. Tested and written to work with 5, 10, 15, 20, and 30. Not sure how well it would work with other values.

<code>AddTimeValue([StartTime], [TimeValue])</code> Adds a string time of hhmm to another time value. Eg, adding 0400 to 0930 would return 1330

<code>SubtractTimeValue([StartTime], [TimeValue])</code> Subtracts a time of hhmm from another time value. Eg, subtracting 02:35 from 1425 would return 1150

<code>AddHours([TimeValue], [Integer Hours])</code> Adds an integer number of hours to a time value.

<code>AddMinutes([TimeValue], [Integer Minutes])</code> Adds an integer number of minutes to a time value.

<code>SubtractHours([TimeValue], [Integer Hours])</code> Subtracts an integer number of hours from a time value.

<code>SubtractMinutes([TimeValue], [Integer Minutes])</code> Subtracts an integer number of minutes from a time value.

<code>Function ToDecimal([timevalue])</code> Converts a time value to a single (decimal) value. Eg, 03:05 P would return 15.08

<code>FromDecimal([timevalue])</code> Converts a decimal to time value. Eg, 7.42 would return 07:25

<code>objVar.DifferenceAsTime([FirstTime], [SecondTime])</code> Returns an hours and minutes value that shows the difference between two times. Eg, the difference between 10:00 A and 11:30 A would be 01:30

<code>objVar.DifferenceAsSingle([FirstTime], [SecondTime])</code> Returns a single (decimal) value that shows the difference, in hours, between two times. Eg, the difference between 10:00 A and 11:30 A would be 1.5

<code>objVar.InRange([TimeValue], [StartTime], [EndTime])</code>  Returns true if the time value falls within the time range created by start time and end time. If the end time is earlier than the start time, the procedure assumes the time range spans midnight.


