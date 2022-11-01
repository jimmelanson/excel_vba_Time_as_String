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

You can download the clsTimeString file to include in your project.

