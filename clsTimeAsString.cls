VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTimeAsString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'   Created 2022.10.29
'   Jim Melanson
'   jmelanson1965@gmail.com

Private p_TimeValue As String
Private p_TimeFormat As String
Private p_AMPM As Boolean
Private p_RoundTo As Integer

Private Sub Class_Initialize()
    p_TimeValue = ""
    p_TimeFormat = "24"
    p_AMPM = False
    p_RoundTo = -1
End Sub

Private Sub Class_Terminate()
    p_TimeValue = ""
    p_TimeFormat = ""
    p_AMPM = False
    p_RoundTo = -1
End Sub

Property Get AMPM() As Boolean
    AMPM = p_AMPM
End Property
Property Let AMPM(ByVal boolValue As Boolean)
        p_AMPM = boolValue
End Property

Property Get TimeFormat() As String
    TimeFormat = p_TimeFormat
End Property
Property Let TimeFormat(ByVal strValue As String)
    If Left(strValue, 2) = "12" Then
        p_TimeFormat = "12"
    ElseIf UCase(Left(strValue, 1)) = "F" Then
        p_TimeFormat = "FRENCH"
    Else
        p_TimeFormat = "24"
    End If
End Property

Property Get RoundTo() As Integer
    RoundTo = p_RoundTo
End Property
Property Let RoundTo(ByVal intValue As Integer)
    p_RoundTo = intValue
End Property

'=========================================PROCEDURES & METHODS==========================================

'=============================================================
'       FORMAT TIME (Public function)
'
'       Debug.Print objVar.FormatTime([input time value])
'
'=============================================================

Public Function FormatTimeValue(ByVal strTimeValue As String, Optional ByVal boolForce24 As Boolean) As String
    'Optional boolForce24
    'Internal processes using this procedure will want a 24-hour format. This procedure
    'automatically changes the input values to 24-hour format for formatting, before
    'changing back to the user's desired format. This boolean flag simply prevents that
    'switch back from happening. It can be called publicly.

    'Fix typos:
    '   • letter O in place of number 0
    '   • semi-colon instead of colon
    '   • letter "l" instead of number 1
    strTimeValue = Trim(strTimeValue)
    strTimeValue = Replace(strTimeValue, "O", "0")
    strTimeValue = Replace(strTimeValue, "o", "0")
    strTimeValue = Replace(strTimeValue, "l", "1")
    strTimeValue = Replace(strTimeValue, " ", "")

    'Default some troublesome entries
    If strTimeValue = "00:00" Or strTimeValue = "00:00 A" Or strTimeValue = "24:00" Or strTimeValue = "2400" Or strTimeValue = "000" Or strTimeValue = "00" Or strTimeValue = "0" Then
        FormatTimeValue = IIf(boolForce24 = True, "0000", c_FinalConversion("0000"))
        Exit Function
    End If
    
    'Default return value in case of error
    FormatTimeValue = "Error"

    'Declare regex
    Dim regexThis As RegExp
    Set regexThis = New RegExp
    regexThis.IgnoreCase = True
    
    '1 character
    '  "^\d$"
    '2 characters
    '  "^\d\d$"
    '3 characters
    '  "^\d\d\d$"    "^\:\d\d$"    "^\d\d\:$"
    '4 characters
    '  "^\d\d\d\d$"    "^\d\:\d\d$"    "^\d\d\:\d$"
    '5 Characters
    '  "^\d\d\:\d\d$"

    If Len(strTimeValue) = 1 Then
        regexThis.Pattern = "^\d$"
        If regexThis.test(strTimeValue) = True Then
            FormatTimeValue = IIf(boolForce24 = True, "0" & strTimeValue & "00", c_FinalConversion("0" & strTimeValue & "00"))
            Set regexThis = Nothing
            Exit Function
        End If
    ElseIf Len(strTimeValue) = 2 Then
        regexThis.Pattern = "^\d\d$"
        If regexThis.test(strTimeValue) = True Then
            If CInt(strTimeValue) < 24 Then
                FormatTimeValue = IIf(boolForce24 = True, strTimeValue & "00", c_FinalConversion(strTimeValue & "00"))
                Set regexThis = Nothing
                Exit Function
            ElseIf CInt(strTimeValue) < 60 Then
                FormatTimeValue = IIf(boolForce24 = True, c_RoundTime("00" & strTimeValue), c_FinalConversion("00" & strTimeValue))
                'FormatTimeValue = c_FinalConversion("00" & strTimeValue)
                Set regexThis = Nothing
                Exit Function
            End If
        End If
        
        regexThis.Pattern = "^\dA$"
        If regexThis.test(strTimeValue) = True Then
            FormatTimeValue = IIf(boolForce24 = True, "0" & Left(strTimeValue, 1) & "00", c_FinalConversion("0" & Left(strTimeValue, 1) & "00"))
            Set regexThis = Nothing
            Exit Function
        End If

        regexThis.Pattern = "^\dP$"
        If regexThis.test(strTimeValue) = True Then
            FormatTimeValue = IIf(boolForce24 = True, (Left(strTimeValue, 1) + 12) & "00", c_FinalConversion((Left(strTimeValue, 1) + 12) & "00"))
            Set regexThis = Nothing
            Exit Function
        End If

    ElseIf Len(strTimeValue) = 3 Then
        regexThis.Pattern = "^\d\d\d$"
        If regexThis.test(strTimeValue) = True Then
            FormatTimeValue = IIf(boolForce24 = True, c_RoundTime("0" & strTimeValue), c_FinalConversion("0" & strTimeValue))
            Set regexThis = Nothing
            Exit Function
        End If

        regexThis.Pattern = "^\D\d\d$"
        If regexThis.test(strTimeValue) = True Then
            FormatTimeValue = IIf(boolForce24 = True, c_RoundTime("00" & Right(strTimeValue, 2)), c_FinalConversion("00" & Right(strTimeValue, 2)))
            Set regexThis = Nothing
            Exit Function
        End If

        regexThis.Pattern = "^\d\dA$"
        If regexThis.test(strTimeValue) = True Then
            FormatTimeValue = IIf(boolForce24 = True, Left(strTimeValue, 2) & "00", c_FinalConversion(Left(strTimeValue, 2) & "00"))
            Set regexThis = Nothing
            Exit Function
        End If

        regexThis.Pattern = "^\d\dP$"
        If regexThis.test(strTimeValue) = True Then
            FormatTimeValue = IIf(boolForce24 = True, (CInt(Left(strTimeValue, 2) + 12)) & "00", c_FinalConversion((CInt(Left(strTimeValue, 2)) + 12) & "00"))
            Set regexThis = Nothing
            Exit Function
        End If

        regexThis.Pattern = "^\d\d\D$"
        If regexThis.test(strTimeValue) = True Then
            FormatTimeValue = IIf(boolForce24 = True, Left(strTimeValue, 2) & "00", c_FinalConversion(Left(strTimeValue, 2) & "00"))
            Set regexThis = Nothing
            Exit Function
        End If
    ElseIf Len(strTimeValue) = 4 Then
        regexThis.Pattern = "^\d\d\d\d$"
        If regexThis.test(strTimeValue) = True Then
            FormatTimeValue = IIf(boolForce24 = True, c_RoundTime(strTimeValue), c_FinalConversion(strTimeValue))
            Set regexThis = Nothing
            Exit Function
        End If

        regexThis.Pattern = "^\d\D\d\d$"
        If regexThis.test(strTimeValue) = True Then
            FormatTimeValue = IIf(boolForce24 = True, c_RoundTime("0" & Left(strTimeValue, 1) & Right(strTimeValue, 2)), c_FinalConversion("0" & Left(strTimeValue, 1) & Right(strTimeValue, 2)))
            Set regexThis = Nothing
            Exit Function
        End If

        regexThis.Pattern = "^\d\d\D\d$"
        If regexThis.test(strTimeValue) = True Then
            FormatTimeValue = IIf(boolForce24 = True, c_RoundTime(Left(strTimeValue, 2) & Right(strTimeValue, 1) & "0"), c_FinalConversion(Left(strTimeValue, 2) & Right(strTimeValue, 1) & "0"))
            Set regexThis = Nothing
            Exit Function
        End If
    ElseIf Len(strTimeValue) >= 5 Then
        regexThis.Pattern = "^\d\d\D\d\d$"
        If regexThis.test(strTimeValue) = True Then
            FormatTimeValue = IIf(boolForce24 = True, c_RoundTime(Left(strTimeValue, 2) & Right(strTimeValue, 2)), c_FinalConversion(Left(strTimeValue, 2) & Right(strTimeValue, 2)))
            Set regexThis = Nothing
            Exit Function
        End If
        Dim intThisHour As Integer
        Dim strThisHour As String
        Dim strThisMinute As String
        'Remember: This procedure process 24-hour format internally, then converts to selected format afterwards
        regexThis.Pattern = "^\d\d\D\d\d(P|PM)?$"
        If regexThis.test(strTimeValue) = True Then
            'If true, then this is a PM value.
            intThisHour = Left(strTimeValue, 2)
            If intThisHour < 12 Then
                'Hour value is less than 12 so it needs to be converted to the 24-hour version of PM by adding 12.
                strThisHour = CStr(intThisHour + 12)
            ElseIf intThisHour = 12 Then
                'The hour value is exactly 12 so it doesn't change
                strThisHour = 12
            ElseIf intThisHour < 10 Then
                'The hour is less than 10, a single digit, so we need to put a zero in front of it.
                strThisHour = "0" & CStr(intThisHour)
            Else
                'The hour value is two digits and does not need to be changed
                strThisHour = CStr(intThisHour)
            End If
            strThisMinute = Mid(strTimeValue, 4, 2)
            FormatTimeValue = IIf(boolForce24 = True, c_RoundTime(strThisHour & strThisMinute), c_FinalConversion(strThisHour & strThisMinute))
            Set regexThis = Nothing
            Exit Function
        End If
        
        regexThis.Pattern = "^\d\d\D\d\d(A|AM)?$"
        If regexThis.test(strTimeValue) = True Then
            'If it is not true, then it is an AM value.
            FormatTimeValue = IIf(boolForce24 = True, (Left(strTimeValue, 2) & Mid(strTimeValue, 4, 2)), c_FinalConversion(Left(strTimeValue, 2) & Mid(strTimeValue, 4, 2)))
            Set regexThis = Nothing
            Exit Function
        End If

        regexThis.Pattern = "^\d\D\d\d(P|PM)?$"
        If regexThis.test(strTimeValue) = True Then
            'If true, then this is a PM value.
            intThisHour = Left(strTimeValue, 1)
            If intThisHour < 12 Then
                'Hour value is less than 12 so it needs to be converted to the 24-hour version of PM by adding 12.
                strThisHour = CStr(intThisHour + 12)
            ElseIf intThisHour = 12 Then
                'The hour value is exactly 12 so it doesn't change
                strThisHour = 12
            ElseIf intThisHour < 10 Then
                'The hour is less than 10, a single digit, so we need to put a zero in front of it.
                strThisHour = "0" & CStr(intThisHour)
            Else
                'The hour value is two digits and does not need to be changed
                strThisHour = CStr(intThisHour)
            End If
            strThisMinute = Mid(strTimeValue, 3, 2)
            FormatTimeValue = IIf(boolForce24 = True, c_RoundTime(strThisHour & strThisMinute), c_FinalConversion(strThisHour & strThisMinute))
            'FormatTimeValue = c_FinalConversion(Left(strTimeValue, 2) & Right(strTimeValue, 2))
            Set regexThis = Nothing
            Exit Function
        End If
        
        regexThis.Pattern = "^\d\D\d\d(A|AM)?$"
        If regexThis.test(strTimeValue) = True Then
            'If it is not true, then it is an AM value.
            FormatTimeValue = IIf(boolForce24 = True, "0" & c_RoundTime(Left(strTimeValue, 1) & Mid(strTimeValue, 3, 2)), c_FinalConversion("0" & Left(strTimeValue, 1) & Mid(strTimeValue, 3, 2)))
            'FormatTimeValue = c_FinalConversion(Left(strTimeValue, 2) & Right(strTimeValue, 2))
            Set regexThis = Nothing
            Exit Function
        End If

    End If
    Set regexThis = Nothing
End Function

'==========================================================
'       ADD / SUBTRACT TIME VALUES (hh:mm + hh:mm)
'==========================================================

Public Function AddTimeValue(ByVal strBaseTime As String, ByVal strAddTime As String) As String
    'AddTimes([hhmm], [hhmm])
    strBaseTime = FormatTimeValue(strBaseTime, True)
    strAddTime = FormatTimeValue(strAddTime, True)
    AddTimeValue = strBaseTime
    If Len(strBaseTime) = 4 And Len(strAddTime) = 4 Then
        Dim intStartHour As Integer
        intStartHour = CInt(Mid(strBaseTime, 1, 2))
        Dim intStartMinute As Integer
        intStartMinute = CInt(Mid(strBaseTime, 3, 2))
        
        Dim intAddHour As Integer
        intAddHour = CInt(Mid(strAddTime, 1, 2))
        Dim intAddMinute As Integer
        intAddMinute = CInt(Mid(strAddTime, 3, 2))
    
        Dim intFinalHour As Integer
        intFinalHour = intStartHour + intAddHour
        Dim intFinalMinute As Integer
        intFinalMinute = intStartMinute + intAddMinute
        If intFinalMinute >= 60 Then
            Do While intFinalMinute >= 60
                intFinalHour = intFinalHour + 1
                intFinalMinute = intFinalMinute - 60
            Loop
        End If
        If intFinalHour >= 24 Then
            intFinalHour = intFinalHour - 24
        End If
        If intFinalHour < 10 Then
            AddTimeValue = "0" & CStr(intFinalHour)
        Else
            AddTimeValue = CStr(intFinalHour)
        End If
        If intFinalMinute < 10 Then
            AddTimeValue = AddTimeValue & "0" & CStr(intFinalMinute)
        Else
            AddTimeValue = AddTimeValue & CStr(intFinalMinute)
        End If
        AddTimeValue = c_FinalConversion(AddTimeValue)
    End If
End Function

Public Function SubtractTimeValue(ByVal strBaseTime As String, ByVal strSubtractTime As String) As String
    'AddTimes([hhmm], [hhmm])
    'First we need to convert the input times to 24-hour format for ease of processing
    strBaseTime = FormatTimeValue(strBaseTime, True)
    strSubtractTime = FormatTimeValue(strSubtractTime, True)
    'Default the return value to the beginning value so they don't get a null result
    SubtractTimeValue = strBaseTime
    If Len(strBaseTime) = 4 And Len(strSubtractTime) = 4 Then
        'Convert the base hour/minute to integers.
        Dim intStartHour As Integer
        intStartHour = CInt(Mid(strBaseTime, 1, 2))
        Dim intStartMinute As Integer
        intStartMinute = CInt(Mid(strBaseTime, 3, 2))
        'Convert the subtracting hour/minute to integers.
        Dim intSubtractHour As Integer
        intSubtractHour = CInt(Mid(strSubtractTime, 1, 2))
        Dim intSubtractMinute As Integer
        intSubtractMinute = CInt(Mid(strSubtractTime, 3, 2))
        'First we math the hour, understanding it may get modified after mathing the minutes.
        Dim intFinalHour As Integer
        If intStartHour > intSubtractHour Then
            intFinalHour = intStartHour - intSubtractHour
        ElseIf intStartHour = intSubtractHour Then
            intFinalHour = 0
        Else
            'Subtract hour is larger than the start hour. So we reduce the subtract hour
            'by the value of start hour and then deduct THAT value from 24.
            intFinalHour = 24 - (intSubtractHour - intStartHour)
        End If
        'Now we math the minutes
        Dim intFinalMinute As Integer
        If intStartMinute > intSubtractMinute Then
            intFinalMinute = intStartMinute - intSubtractMinute
        ElseIf intStartMinute = intSubtractMinute Then
            intFinalMinute = 0
        Else
            Do While intSubtractMinute >= 60
                intFinalHour = intFinalHour - 1
                intSubtractMinute = intSubtractMinute - 60
            Loop
            If intStartMinute > intSubtractMinute Then
                intFinalMinute = intStartMinute - intSubtractMinute
            ElseIf intSubtractMinute = intStartMinute Then
                intFinalMinute = 0
            Else
                'Subtract minute is larger than the start minute. So we reduce the subtract minute
                'by the value of start minute and then deduct THAT value from 60.
                intFinalMinute = 60 - (intSubtractMinute - intStartMinute)
                intFinalHour = intFinalHour - 1
            End If
        End If
        If intFinalHour < 0 Then
            intFinalHour = 24 - Abs(intFinalHour)
        End If
        'Put the numbers back together again.
        If intFinalHour < 10 Then
            SubtractTimeValue = "0" & CStr(intFinalHour)
        Else
            SubtractTimeValue = CStr(intFinalHour)
        End If
        If intFinalMinute < 10 Then
            SubtractTimeValue = SubtractTimeValue & "0" & CStr(intFinalMinute)
        Else
            SubtractTimeValue = SubtractTimeValue & CStr(intFinalMinute)
        End If
        SubtractTimeValue = c_FinalConversion(SubtractTimeValue)
    End If
End Function

'==========================================================
'       ADD / SUBTRACT HOURS OR MINUTES (hhmm + h)
'==========================================================

Public Function AddHours(ByVal strBaseTime As String, ByVal intHours As Integer) As String
    strBaseTime = FormatTimeValue(strBaseTime, True)
    If intHours > 0 Then
        If intHours < 10 Then
            AddHours = AddTimeValue(strBaseTime, "0" & CStr(intHours) & "00")
        Else
            AddHours = AddTimeValue(strBaseTime, CStr(intHours) & "00")
        End If
    End If
    'function AddTimeValue already applies c_FinalConversion so doing it here would cause an error.
End Function

Public Function AddMinutes(ByVal strBaseTime As String, ByVal intMinutes As Integer) As String
    strBaseTime = FormatTimeValue(strBaseTime, True)
    If intMinutes > 0 Then
        If intMinutes < 10 Then
            AddMinutes = AddTimeValue(strBaseTime, "000" & CStr(intMinutes))
        Else
            AddMinutes = AddTimeValue(strBaseTime, "00" & CStr(intMinutes))
        End If
    End If
    'function AddTimeValue already applies c_FinalConversion so doing it here would cause an error.
End Function

Public Function SubtractHours(ByVal strBaseTime As String, ByVal intHours As Integer) As String
    strBaseTime = FormatTimeValue(strBaseTime, True)
    If intHours > 0 Then
        If intHours < 10 Then
            SubtractHours = SubtractTimeValue(strBaseTime, "0" & CStr(intHours) & "00")
        Else
            SubtractHours = SubtractTimeValue(strBaseTime, CStr(intHours) & "00")
        End If
    End If
    'function SubtractTimeValue already applies c_FinalConversion so doing it here would cause an error.
End Function

Public Function SubtractMinutes(ByVal strBaseTime As String, ByVal intMinutes As Integer) As String
    strBaseTime = FormatTimeValue(strBaseTime, True)
    'Debug.Print "DEBUG Subtract Minutes strBaseTime: " & strBaseTime
    If intMinutes > 0 Then
        If intMinutes < 10 Then
            SubtractMinutes = SubtractTimeValue(strBaseTime, "000" & CStr(intMinutes))
        Else
            SubtractMinutes = SubtractTimeValue(strBaseTime, "00" & CStr(intMinutes))
        End If
    End If
    'Debug.Print "DEBUG FinalValue: " & SubtractMinutes
End Function


'==========================================================
'           CONVERT TIME VALUE TO/FROM DECIMAL
'==========================================================

Public Function ToDecimal(ByVal strBaseTime As String) As Single
    If strBaseTime <> "" Then
        strBaseTime = FormatTimeValue(strBaseTime, True)
        ToDecimal = CInt(Right(strBaseTime, 2)) / 60
        ToDecimal = ToDecimal + CInt(Left(strBaseTime, 2))
    End If
End Function

Public Function FromDecimal(ByVal sngDecimal As Single) As String
    If sngDecimal > 0 Then
        Dim intFinalHour As Integer
        intFinalHour = Int(sngDecimal)

        Dim intFinalMinute As Integer
        intFinalMinute = 60 * (sngDecimal - Int(sngDecimal))

        If intFinalHour < 10 Then
            FromDecimal = "0" & CStr(intFinalHour)
        Else
            FromDecimal = CStr(intFinalHour)
        End If
        If intFinalMinute < 10 Then
            FromDecimal = FromDecimal & "0" & CStr(intFinalMinute)
        Else
            FromDecimal = FromDecimal & CStr(intFinalMinute)
        End If
        FromDecimal = c_FinalConversion(FromDecimal)
    ElseIf sngDecimal = 0 Then
        FromDecimal = c_FinalConversion("0000")
    End If
End Function

'============================================================
'       DIFFERENCE BETWEEN TWO TIMES : TIME OR SINGLE
'============================================================

Public Function DifferenceAsTime(ByVal strBaseTime As String, ByVal strSecondTime As String) As String
    DifferenceAsTime = ""
    'DifferenceAsTime = FromDecimal(DifferenceAsSingle(strBaseTime, strSecondTime))

    Dim sngTotal As Single
    sngTotal = DifferenceAsSingle(strBaseTime, strSecondTime)

    Dim intFinalHour As Integer
    intFinalHour = Int(sngTotal)
    Dim intFinalMinute As Integer
    intFinalMinute = 0
    If sngTotal - Int(sngTotal) <> 0 Then
        intFinalMinute = (sngTotal - Int(sngTotal)) * 60
    End If
    
    If intFinalHour < 10 Then
        DifferenceAsTime = "0" & intFinalHour
    Else
        DifferenceAsTime = intFinalHour
    End If
    
    If intFinalMinute < 10 Then
        DifferenceAsTime = DifferenceAsTime & ":0" & intFinalMinute
    Else
        DifferenceAsTime = DifferenceAsTime & ":" & intFinalMinute
    End If
End Function

Public Function DifferenceAsSingle(ByVal strBaseTime As String, ByVal strSecondTime As String) As Single
    'This procedure relies on both times being supplied in 24-hour format.
    '
    'This procedure also relies on the two times being entered in the correct order. For example, if you
    'enter the times DifferenceAsSingle("0700", "0300") then it reads from 0700 past midnight to 0300 the
    'next morning. If you enter it as DifferenceAsSingle("0300", "0700") then it reads from 3am till 7am
    'the same morning. Two others to note are:
    'DifferenceAsSingle("0000", "2200") reads from midnight to 10pm (22 hours)
    'DifferenceAsSingle("2200", "0000") reads from 10 pm to midnight (2 hours)
    If strBaseTime <> "" And strSecondTime <> "" Then
        strBaseTime = FormatTimeValue(strBaseTime, True)
        strSecondTime = FormatTimeValue(strSecondTime, True)
        If strBaseTime = strSecondTime Then
            DifferenceAsSingle = 0
            Exit Function
        End If
        Dim sngTime1 As Single
        sngTime1 = ToDecimal(strBaseTime)
        Dim sngTime2 As Single
        sngTime2 = ToDecimal(strSecondTime)

        If sngTime1 = 0 Then
            'we are reading from midnight so it's just a decimal value of the second time
            'Eg, DifferenceAsSingle("0000", "2230") = 22.5 - 0 = 22.5
            DifferenceAsSingle = Round(sngTime2, 2)
        ElseIf sngTime2 = 0 Then
            'We are reading from the first time up to midnight, so it is 24 minus the first time's decimal value
            'Eg, DifferenceAsSingle("2230", "0000") = 24 - 22.5 = 1.5
            DifferenceAsSingle = Round(24 - sngTime1, 2)
        Else
            'If sngTime2 is greater than sngTime1 then it is just a subtraction of decimal values.
            'Eg, DifferenceAsDecimal("0700", "0900") = 9 - 7 = 2
            'However, if sngTime2 is smaller than sngTime1, then it is first time until midnight and then
            'midnight until second time.
            '0700-0800, 0900-1700, 1400-0200
            If sngTime2 > sngTime1 Then
                DifferenceAsSingle = Round(sngTime2 - sngTime1, 2)
            Else
                'First time until midnight
                DifferenceAsSingle = Round(24 - sngTime1, 2)
                'Midnight till second time
                DifferenceAsSingle = Round(DifferenceAsSingle + sngTime2, 2)
            End If
        End If
    End If
End Function


'============================================================
'       TIME VALUE BETWEEN TWO OTHER TIME VALUES
'============================================================

Public Function InRange(ByVal strTimeValue As String, ByVal strTimeStart As String, ByVal strTimeEnd As String) As Boolean
    InRange = False
    'The purpose of this procedure is to look at a time, and to see if
    'that time falls between a start and end time. The time CAN match the
    'start time but CANNOT match the end time … Think about it.
    '
    'First we need to make sure that all necessary values are present. There is
    'an argument as to whether this is needed since the arguments are all required.
    'However, I'm an old fart who does things old skool. Deal.
    If strTimeValue <> "" And strTimeStart <> "" And strTimeEnd <> "" Then
        'A time value matching the start time is in range. However, a time value matching
        'the end time is NOT in range. So, if the start time and end time are the same,
        'then no time range can be in between them.
        If strTimeValue = strTimeEnd Then
            Exit Function
        End If
        '----------TIME VALUE TO COMPARE----------
        'First step is to convert the time value we are looking at to the 24-hour format.
        'You may be wondering why we are spending the time to assign all of these values
        'to variables instead of nesting the procedures - because in the end it is faster.
        'This way we only call the conversions once, rather than calling them a full
        'dozen times near the end of the procedure.
        strTimeValue = FormatTimeValue(strTimeValue, True)
        'Change the time to a decimal value
        Dim decimalValueTime As Single
        decimalValueTime = ToDecimal(strTimeValue)
        '----------------START TIME---------------
        'First step is to convert the start time we are looking at to the 24-hour format.
        strTimeStart = FormatTimeValue(strTimeStart, True)
        'Change the time to a decimal value
        Dim decimalValueStart As Single
        decimalValueStart = ToDecimal(strTimeStart)
        '-----------------END TIME----------------
        'First step is to convert the end time we are looking at to the 24-hour format.
        strTimeEnd = FormatTimeValue(strTimeEnd, True)
        'Change the time to a decimal value
        Dim decimalValueEnd As Single
        decimalValueEnd = ToDecimal(strTimeEnd)
        '-----------------------------------------
        If decimalValueEnd < decimalValueStart Then
            'This means that the range starts before midnight and ends after midnight
            If (decimalValueTime >= decimalValueStart And decimalValueTime <= 24) Or (decimalValueTime >= 0 And decimalValueTime <= decimalValueEnd) Then
                InRange = True
            End If
        Else
            'This means the start and end times in the same day
            If decimalValueTime >= decimalValueStart And decimalValueTime < decimalValueEnd Then
                InRange = True
            End If
        End If
    End If
End Function

'=======================================
'       CHECK MINUTE ROUNDING
'       This should be limited to:
'       5, 10, 15, 30, & 30
'=======================================

Public Function c_RoundTime(ByVal strTimeValue As String) As String
    'Make sure that there is both a time value and a round-to value.
    If p_RoundTo >= 0 And strTimeValue <> "" Then
        'We need to work with the minutes as an integer
        Dim intMinutes As Integer
        intMinutes = CInt(Right(strTimeValue, 2))
        If p_RoundTo = 0 Or intMinutes = 0 Then
            'If the round-to value OR the minutes value is a zero, then we will not have to
            'make any changes, so just return the starting value.
            c_RoundTime = strTimeValue
            Exit Function
        ElseIf intMinutes Mod p_RoundTo = 0 Then
            'If the modulus of the time value and the round-to value is zero, then we will not have
            'to make any changes, so just return the starting value.
            c_RoundTime = strTimeValue
            Exit Function
        ElseIf p_RoundTo Mod 5 = 0 And p_RoundTo <= 30 Then
            'If the round-to value is a multiple of 5 AND it is less than or equal to 30, then we can
            'proceed with the procedure.
            Dim i_Before As Integer
            i_Before = 0
            Dim i_After As Integer
            i_After = 0
            Dim i As Integer
            'First we will look from the minutes given to zero to find the first modulus of
            'a mintue and the round-to value that is zero. We need to know how many minutes
            'difference it is.
            For i = intMinutes To 0 Step -1
                If i_Before = 0 Then
                    If i Mod p_RoundTo = 0 Then
                        i_Before = intMinutes - i
                    End If
                End If
            Next i
            'Next we will look from the minutes given to 60 to find the first modulus of
            'a mintue and the round-to value that is zero. We need to know how many minutes
            'difference it is.
            For i = intMinutes To 60
                If i_After = 0 Then
                    If i Mod p_RoundTo = 0 Then
                        i_After = i - intMinutes
                    End If
                End If
            Next i
            'If the number of minutes towards zero is larger than the number of minutes towards 60, then we
            'are going to round UP towards 60.
            'If the number of minutes towards 60 is larger than the number of minutes towards zero, then we
            'are going to round DOWN towards 0.
            If i_After < i_Before Then
                'Rounding UP
                '
                'We need to know if the match was on 60. If it was, then we need to increase the hour by
                'one as well.
                If intMinutes + i_After = 60 Then
                    Dim intHour As Integer
                    intHour = CInt(Left(strTimeValue, 2))
                    intHour = intHour + 1
                    If intHour < 10 Then
                        c_RoundTime = "0" & CStr(intHour) & "00"
                    Else
                        c_RoundTime = CStr(intHour) & "00"
                    End If
                Else
                    If intMinutes = 0 Then
                        c_RoundTime = Left(strTimeValue, 2) & "00"
                    ElseIf intMinutes + i_After < 10 Then
                        c_RoundTime = Left(strTimeValue, 2) & "0" & CStr(intMinutes + i_After)
                        'c_RoundTime = Left(strTimeValue, 2) & "0" & CStr(intMinutes)
                    Else
                        c_RoundTime = Left(strTimeValue, 2) & CStr(intMinutes + i_After)
                        'c_RoundTime = Left(strTimeValue, 2) & CStr(intMinutes)
                    End If
                End If
            Else
                'Rounding DOWN
                If intMinutes - i_Before < 10 Then
                    c_RoundTime = Left(strTimeValue, 2) & "0" & intMinutes - i_Before
                Else
                    c_RoundTime = Left(strTimeValue, 2) & intMinutes - i_Before
                End If
            End If
        Else
            'something was not true so we are exiting without doing anything.
            c_RoundTime = strTimeValue
        End If
    Else
        c_RoundTime = strTimeValue
    End If
End Function

'=============================================================
'       FINAL FORMAT (Private function for class only)
'
'       Call c_FinalConversion   - No arguments
'
'       All processing is done with the
'       24-hour clock. Final format will
'       convert it to 12-hour clock or
'       French clock if needed.
'=============================================================

Private Function c_FinalConversion(ByVal strTimeValue As String) As String
    c_FinalConversion = ""
    'Debug.Print "DEBUG Converting: " & strTimeValue
    If Len(strTimeValue) = 4 Then
        strTimeValue = c_RoundTime(strTimeValue)
        'Make sure hour value is valid
        If CInt(Left(strTimeValue, 2)) > 24 Then
            c_FinalConversion = "Error"
            Exit Function
        ElseIf CInt(Left(strTimeValue, 2)) = 24 Then
            strTimeValue = "00" & Right(strTimeValue, 2)
        End If
        
        'Make sure minute value is valid
        If CInt(Right(strTimeValue, 2)) > 60 Then
            c_FinalConversion = "Error"
            Exit Function
        ElseIf CInt(Right(strTimeValue, 2)) = 60 Then
            strTimeValue = Left(strTimeValue, 2) & "00"
        End If
        'Default
        c_FinalConversion = strTimeValue
        'Do the conversion if necessary
        If p_TimeFormat = "12" Then
            If CInt(Left(strTimeValue, 2)) > 12 Then
                If (CInt(Left(strTimeValue, 2)) - 12) < 10 Then
                    c_FinalConversion = "0" & CStr(CInt(Left(strTimeValue, 2)) - 12) & ":" & Right(strTimeValue, 2)
                Else
                    c_FinalConversion = CStr(CInt(Left(strTimeValue, 2)) - 12) & ":" & Right(strTimeValue, 2)
                End If
                If p_AMPM = True Then
                    c_FinalConversion = c_FinalConversion & " P"
                End If
            Else
                c_FinalConversion = Left(strTimeValue, 2) & ":" & Right(strTimeValue, 2)
                If p_AMPM = True Then
                    If Left(strTimeValue, 2) = 12 Then
                        c_FinalConversion = c_FinalConversion & " P"
                    Else
                        c_FinalConversion = c_FinalConversion & " A"
                    End If
                End If
            End If
        ElseIf UCase(Left(p_TimeFormat, 1)) = "F" Then
            c_FinalConversion = Left(strTimeValue, 2) & "h" & Right(strTimeValue, 2)
        End If
    Else
        c_FinalConversion = "Error"
    End If
End Function



