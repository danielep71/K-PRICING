Attribute VB_Name = "CORE_Dates"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: CORE_Dates
'------------------------------------------------------------------------------
' PURPOSE
'   The date layer's calendar core: the supported window, the Gregorian month, year and
'   weekday rules, period boundaries, shifts by days and months, weekday locators, and
'   the parse, shift and format of pillar tokens.
'
' WHY THIS EXISTS
'   Calendar arithmetic is shared by many worksheet functions. Keeping it in one
'   module that knows no worksheet means that month length, month shifting, the window
'   and pillar rounding are computed one way only, and a facade holds no date
'   algorithm.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides, by calculation contract:
'     - CALC-02: DayOfWeek
'     - CALC-03: IsLeapYear, DaysInMonth, DaysInYear
'     - CALC-04: BeginOfMonth, EndOfMonth, BeginOfQuarter, EndOfQuarter, BeginOfYear,
'       EndOfYear
'     - CALC-05: IsMonthEnd, IsQuarterEnd, IsYearEnd
'     - CALC-06: TryAddDays, TryAddMonths
'     - CALC-07: TryNthWeekday, TryLastWeekday
'     - CALC-08: TryParsePillar, TryAddPillar
'     - CALC-09: TryFormatPillar
'     - the window predicate IsDateInWindow, and the argument domains CheckYear,
'       CheckMonth, CheckWeekday and CheckOccurrence
'
' ERROR POLICY
'   No member raises to its caller. The calendar functions are total for the values they
'   document. The Try members return False, with a condition where they report one, and
'   the two that trap runtime errors clear them. The module never builds a worksheet
'   error value: the facade maps each condition with CORE_Err.ErrForCondition.
'
' DEPENDENCIES
'   CORE_Err, for the condition vocabulary, and CORE_Parse, which owns the supported
'   window and the ASCII text helpers. No Excel object model: the VBA date functions
'   are used directly. Calculation contracts: docs/methodology/CALCULATION_CONTRACTS.md,
'   CALC-02 to CALC-09 and the common rules
'
' NOTES
'   - Date arithmetic follows VBA DateSerial, DateDiff and Weekday, which agree with
'     Excel's serials from 1900-03-01, the first day of the window
'   - Day distances between two known whole days are taken by serial subtraction, not
'     DateDiff: the results are identical and the subtraction is cheaper on the per-cell
'     path
'   - AddYears is AddMonths with twelve times the count, through TryAddMonths, so the
'     29 February, short-month and month-end edges cannot drift between them
'   - DaysInMonth is the one source of month length, and IsLeapYear the one leap rule.
'     The day-0 idiom DateSerial(y, m + 1, 0) is not used: it raises an error at both ends
'     of the VBA Date range, so month length is derived arithmetically
'   - The Begin* boundaries are not gated here: BeginOfYear of a 1900 date is 1900-01-01,
'     outside the window, and the facade reports it with IsDateInWindow
'   - Option Private Module keeps every member out of the Excel function list. Nothing
'     here is supported API
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'The last month index, 9999 * 12 + 11, which CORE_Parse.WINDOW_MAX_YEAR gives; the
    'index counts months from year 0, so that a month shift is one addition
    Private Const MAX_MONTH_INDEX   As Long = 119999    'December 9999

'------------------------------------------------------------------------------
' PUBLIC ENUMS
'------------------------------------------------------------------------------
    'The rounding mode of a pillar token, calculation contract CALC-09
    Public Enum KPRPillarRounding
        prNearest = 0   'Candidate closest to the end date
        prFloor = 1     'Farthest candidate that does not pass the end date
        prCeiling = 2   'Closest candidate that reaches or passes the end date
    End Enum


'
'------------------------------------------------------------------------------
'
'                              WINDOW AND DOMAINS
'
'------------------------------------------------------------------------------
'

Public Function IsDateInWindow(ByVal DateIn As Date) As Boolean

'
'------------------------------------------------------------------------------
'                        TEST A DATE AGAINST THE WINDOW
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a date lies inside the supported window, 1900-03-01 to 9999-12-31,
'   both included.
'
' WHY THIS EXISTS
'   Every date a function returns is gated like a result (calculation contracts, common
'   rules, Dates), and the gates must all use the one window.
'
' INPUTS
'   DateIn
'     A VBA Date; a time of day is compared as part of its serial, so callers pass
'     date-only values.
'
' RETURNS
'   Boolean
'     True when WINDOW_MIN_SERIAL <= DateIn <= WINDOW_MAX_SERIAL.
'
' BEHAVIOR
'   Compares the serial of DateIn with the two bounds that CORE_Parse owns.
'
' ERROR POLICY
'   None. Total for every Date.
'
' DEPENDENCIES
'   CORE_Parse.WINDOW_MIN_SERIAL, CORE_Parse.WINDOW_MAX_SERIAL. CALC-02 to CALC-10, common rules
'
' NOTES
'   The window starts at serial 61 so that it excludes the region where Excel's 1900
'   date system and VBA disagree, because of the 29 February 1900 that never existed
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Both bounds are included
        IsDateInWindow = (VBA.CDbl(DateIn) >= CORE_Parse.WINDOW_MIN_SERIAL) And _
            (VBA.CDbl(DateIn) <= CORE_Parse.WINDOW_MAX_SERIAL)

End Function

Public Function CheckYear(ByVal ValueIn As Long) As KPRCondition

'
'------------------------------------------------------------------------------
'                                 CHECK A YEAR
'------------------------------------------------------------------------------
' PURPOSE
'   Checks that a calendar year is inside its domain, CORE_Parse.WINDOW_MIN_YEAR to CORE_Parse.WINDOW_MAX_YEAR.
'
' WHY THIS EXISTS
'   A parsed whole number can still be outside the domain of its argument. The facade
'   checks each argument as it reads it, so that the first failing argument of a
'   position decides its result, and the domain stays defined once.
'
' INPUTS
'   ValueIn
'     The parsed whole number: a calendar year.
'
' RETURNS
'   KPRCondition
'     cdNone when CORE_Parse.WINDOW_MIN_YEAR <= ValueIn <= CORE_Parse.WINDOW_MAX_YEAR, otherwise cdDomainYear.
'
' BEHAVIOR
'   Compares the value with the two bounds.
'
' ERROR POLICY
'   None. Total for every Long.
'
' DEPENDENCIES
'   CORE_Err (KPRCondition). Data contract: Function arguments, Integer values; CALC-03, CALC-07
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Report the condition, or none
        If (ValueIn < CORE_Parse.WINDOW_MIN_YEAR) Or (ValueIn > CORE_Parse.WINDOW_MAX_YEAR) Then
            CheckYear = cdDomainYear
        Else
            CheckYear = cdNone
        End If

End Function

Public Function CheckMonth(ByVal ValueIn As Long) As KPRCondition

'
'------------------------------------------------------------------------------
'                                CHECK A MONTH
'------------------------------------------------------------------------------
' PURPOSE
'   Checks that a calendar month is inside its domain, 1 to 12.
'
' WHY THIS EXISTS
'   A parsed whole number can still be outside the domain of its argument. The facade
'   checks each argument as it reads it, so that the first failing argument of a
'   position decides its result, and the domain stays defined once.
'
' INPUTS
'   ValueIn
'     The parsed whole number: a calendar month.
'
' RETURNS
'   KPRCondition
'     cdNone when 1 <= ValueIn <= 12, otherwise cdDomainMonth.
'
' BEHAVIOR
'   Compares the value with the two bounds.
'
' ERROR POLICY
'   None. Total for every Long.
'
' DEPENDENCIES
'   CORE_Err (KPRCondition). Data contract: Function arguments, Integer values; CALC-07
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Report the condition, or none
        If (ValueIn < 1) Or (ValueIn > 12) Then
            CheckMonth = cdDomainMonth
        Else
            CheckMonth = cdNone
        End If

End Function

Public Function CheckWeekday(ByVal ValueIn As Long) As KPRCondition

'
'------------------------------------------------------------------------------
'                            CHECK A WEEKDAY NUMBER
'------------------------------------------------------------------------------
' PURPOSE
'   Checks that a weekday number in the call's weekday base is inside its domain, 1 to 7.
'
' WHY THIS EXISTS
'   A parsed whole number can still be outside the domain of its argument. The facade
'   checks each argument as it reads it, so that the first failing argument of a
'   position decides its result, and the domain stays defined once.
'
' INPUTS
'   ValueIn
'     The parsed whole number: a weekday number in the call's weekday base.
'
' RETURNS
'   KPRCondition
'     cdNone when 1 <= ValueIn <= 7, otherwise cdDomainWeekday.
'
' BEHAVIOR
'   Compares the value with the two bounds.
'
' ERROR POLICY
'   None. Total for every Long.
'
' DEPENDENCIES
'   CORE_Err (KPRCondition). Data contract: Function arguments, Integer values; CALC-07
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Report the condition, or none
        If (ValueIn < 1) Or (ValueIn > 7) Then
            CheckWeekday = cdDomainWeekday
        Else
            CheckWeekday = cdNone
        End If

End Function

Public Function CheckOccurrence(ByVal ValueIn As Long) As KPRCondition

'
'------------------------------------------------------------------------------
'                             CHECK AN OCCURRENCE
'------------------------------------------------------------------------------
' PURPOSE
'   Checks that the occurrence of a weekday in a month is inside its domain, 1 to 5.
'
' WHY THIS EXISTS
'   A parsed whole number can still be outside the domain of its argument. The facade
'   checks each argument as it reads it, so that the first failing argument of a
'   position decides its result, and the domain stays defined once.
'
' INPUTS
'   ValueIn
'     The parsed whole number: the occurrence of a weekday in a month.
'
' RETURNS
'   KPRCondition
'     cdNone when 1 <= ValueIn <= 5, otherwise cdDomainOccurrence.
'
' BEHAVIOR
'   Compares the value with the two bounds.
'
' ERROR POLICY
'   None. Total for every Long.
'
' DEPENDENCIES
'   CORE_Err (KPRCondition). Data contract: Function arguments, Integer values; CALC-07
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Report the condition, or none
        If (ValueIn < 1) Or (ValueIn > 5) Then
            CheckOccurrence = cdDomainOccurrence
        Else
            CheckOccurrence = cdNone
        End If

End Function

'
'------------------------------------------------------------------------------
'
'                                   CALENDAR
'
'------------------------------------------------------------------------------
'

Public Function IsLeapYear(ByVal YearIn As Long) As Boolean

'
'------------------------------------------------------------------------------
'                      APPLY THE GREGORIAN LEAP-YEAR RULE
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the Gregorian leap-year rule to a calendar year.
'
' WHY THIS EXISTS
'   The rule is short enough to inline and subtle enough to get wrong. Stating
'   the century exceptions once means DaysInMonth is the only place February
'   is resolved, and every other month length derives from that.
'
' INPUTS
'   YearIn
'     Calendar year, not a date serial. Callers holding a Date must pass
'     Year(DateIn).
'
' RETURNS
'   Boolean
'     TRUE => YearIn is a leap year under the proleptic Gregorian rule.
'
' BEHAVIOR
'   Evaluates the divisibility rule below on YearIn.
'
' ERROR POLICY
'   None. Total for every Long; the caller owns range validation.
'
' DEPENDENCIES
'   None
'
' NOTES
'   - Rule: divisible by 4 => leap, unless divisible by 100, unless also
'     divisible by 400. So 1900 is not a leap year and 2000 is.
'   - The rule is applied proleptically. VBA Date arithmetic is Gregorian
'     over its whole range, so no Julian calendar or 1582 changeover is
'     modelled here and 1500 reports TRUE. That is consistent with the rest
'     of the layer, not a historical claim.
'   - The supported window is not enforced. Any Long is evaluated, including
'     zero and negatives; VBA Mod takes the sign of the dividend, so -4
'     reports TRUE. Values outside the window are the caller's to reject.
'
' UPDATED
'   2026-09-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Gregorian rule
        IsLeapYear = ((YearIn Mod 4 = 0) And ((YearIn Mod 100 <> 0) Or (YearIn Mod 400 = 0)))

End Function

Public Function DaysInMonth(ByVal YearIn As Long, ByVal MonthIn As Long) As Long

'
'------------------------------------------------------------------------------
'                         RETURN THE LENGTH OF A MONTH
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the length in days of a given calendar month.
'
' WHY THIS EXISTS
'   Month length is needed by EndOfMonth, by the EOM anchor test, and by the
'   day clip in TryAddMonths. Deriving it arithmetically rather than through
'   DateSerial avoids its upper-bound normalization trap:
'   DateSerial(9999, 13, 0) raises error 5. The lower supported boundary is
'   not symmetric; DateSerial(1900, 1, 0) returns 1899-12-31 successfully.
'
' INPUTS
'   YearIn
'     Calendar year, not a date serial. Read only for February.
'
'   MonthIn
'     Calendar month, 1 to 12.
'
' RETURNS
'   Long
'     28, 29, 30 or 31 for a valid month; 0 for any MonthIn outside 1 to 12.
'
' BEHAVIOR
'   Returns the fixed length of a 30- or 31-day month, applies IsLeapYear for
'   February, and returns 0 for an invalid month.
'
' ERROR POLICY
'   None. Total for every pair of Longs, including an invalid month.
'
' DEPENDENCIES
'   IsLeapYear
'
' NOTES
'   - A zero return signals an invalid month rather than raising, because this
'     is an internal member reached only with an already-resolved month. A
'     caller that can produce an out-of-range month must test the result; one
'     that cannot may use it directly.
'   - No date is constructed, so the result is valid for any year the caller
'     can represent, including years outside the supported window.
'
' UPDATED
'   2026-09-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Fixed lengths first; February is the only month needing the leap rule
        Select Case MonthIn
            Case 1, 3, 5, 7, 8, 10, 12
                DaysInMonth = 31
            Case 4, 6, 9, 11
                DaysInMonth = 30
            Case 2
                If IsLeapYear(YearIn) Then
                    DaysInMonth = 29
                Else
                    DaysInMonth = 28
                End If
            Case Else
                'Invalid month; reported rather than raised
                DaysInMonth = 0
        End Select

End Function

Public Function DaysInYear(ByVal YearIn As Long) As Long

'
'------------------------------------------------------------------------------
'                         RETURN THE LENGTH OF A YEAR
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of days of a calendar year: 366 for a leap year, 365 otherwise.
'
' WHY THIS EXISTS
'   KPR_Dates_DaysInYear takes a year, and the leap rule is stated once, in IsLeapYear.
'
' INPUTS
'   YearIn
'     Calendar year, not a date serial.
'
' RETURNS
'   Long
'     365 or 366.
'
' BEHAVIOR
'   Applies IsLeapYear.
'
' ERROR POLICY
'   None. Total for every Long.
'
' DEPENDENCIES
'   IsLeapYear. CALC-03 step 3
'
' NOTES
'   No date is constructed, so no window applies: the result is correct for 1900 and
'   for any other year
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'A leap year has one more day
        If IsLeapYear(YearIn) Then
            DaysInYear = 366
        Else
            DaysInYear = 365
        End If

End Function

Public Function DayOfWeek( _
    ByVal DateIn As Date, _
    ByVal MondayBase As Boolean) As Long
'
'------------------------------------------------------------------------------
'                        RETURN THE NUMBER OF A WEEKDAY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of a date's weekday, 1 to 7, Monday-based or Sunday-based.
'
' WHY THIS EXISTS
'   KPR_Dates_DayOfWeek and the weekday locators number weekdays the same way, so the base
'   is applied once.
'
' INPUTS
'   DateIn
'     A date inside the supported window.
'
'   MondayBase
'     True: Monday is 1 and Sunday is 7. False: Sunday is 1 and Saturday is 7.
'
' RETURNS
'   Long
'     1 to 7.
'
' BEHAVIOR
'   Reads the weekday of the date under the base that MondayBase selects.
'
' ERROR POLICY
'   None. Total for every Date.
'
' DEPENDENCIES
'   VBA.Weekday. CALC-02
'
' NOTES
'   Sunday-based numbering is WEEKDAY's return type 1 of Excel. The default of the
'   worksheet function is Monday-based, the reverse of WEEKDAY's, and the facade passes it
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Dates.DayOfWeek"

    Dim FirstDay        As VbDayOfWeek  'Weekday that numbers as 1

'------------------------------------------------------------------------------
' SELECT BASE
'------------------------------------------------------------------------------
    'Monday-based numbering starts the week on Monday
        If MondayBase Then
            FirstDay = vbMonday
        Else
            FirstDay = vbSunday
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Read the weekday under that base
        DayOfWeek = VBA.Weekday(DateIn, FirstDay)

End Function

'
'------------------------------------------------------------------------------
'
'                              PERIOD BOUNDARIES
'
'------------------------------------------------------------------------------
'

Public Function BeginOfMonth(ByVal DateIn As Date) As Date

'
'------------------------------------------------------------------------------
'                       RETURN THE FIRST DAY OF A MONTH
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the first day of the month that holds a date.
'
' WHY THIS EXISTS
'   KPR_Dates_BeginOfMonth, and the quarter and year boundaries, build on one month
'   boundary.
'
' INPUTS
'   DateIn
'     A date inside the supported window.
'
' RETURNS
'   Date
'     The first day of its month. For a date inside the window it is inside the window too.
'
' BEHAVIOR
'   Builds the date from the year and month of DateIn and day 1.
'
' ERROR POLICY
'   None. Total for the dates of the window.
'
' DEPENDENCIES
'   VBA.DateSerial. CALC-04 step 2
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Day 1 of the same month
        BeginOfMonth = VBA.DateSerial(VBA.Year(DateIn), VBA.Month(DateIn), 1)

End Function

Public Function EndOfMonth(ByVal DateIn As Date) As Date

'
'------------------------------------------------------------------------------
'                       RETURN THE LAST DAY OF THE MONTH
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the last calendar day of the month containing DateIn.
'
' WHY THIS EXISTS
'   Month-end is the anchor for the whole month-arithmetic family. Isolating
'   it here means no caller re-derives it.
'
' INPUTS
'   DateIn
'     Already-parsed VBA Date. A time component is tolerated and discarded.
'
' RETURNS
'   Date
'     Midnight on the last calendar day of the containing month.
'
' BEHAVIOR
'   Splits DateIn into year and month, then builds the date whose day is the
'   month length from DaysInMonth.
'
' ERROR POLICY
'   None. Total over the whole VBA Date range, both boundary months included.
'
' DEPENDENCIES
'   DaysInMonth
'
' NOTES
'   - The month length comes from DaysInMonth, so no date outside the month
'     is ever constructed. The day-0 idiom would be shorter but raises error 5
'     for DateSerial(9999, 13, 0). DateSerial(1900, 1, 0) succeeds; see MONTH
'     LENGTH in the module header for the distinct VBA-range lower boundary.
'   - Year and Month ignore any time fraction, and DateSerial always returns
'     midnight, so the result is date-only regardless of what arrives. No
'     caller needs to strip the time first.
'   - Totality is asserted over the whole VBA Date range, which is wider than
'     the supported window. Pre-1900 dates are negative serials and behave
'     identically.
'
' UPDATED
'   2026-09-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Dates.EndOfMonth"

    Dim YearPart        As Long     'Calendar year of DateIn
    Dim MonthPart       As Long     'Calendar month of DateIn, 1 to 12

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Decompose once; both parts feed the month-length lookup
        YearPart = VBA.Year(DateIn)
        MonthPart = VBA.Month(DateIn)

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Last day of the containing month, built without leaving it
        EndOfMonth = VBA.DateSerial(YearPart, MonthPart, DaysInMonth(YearPart, MonthPart))

End Function

Public Function BeginOfQuarter(ByVal DateIn As Date) As Date

'
'------------------------------------------------------------------------------
'                     RETURN THE FIRST DAY OF THE QUARTER
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the first day of the calendar quarter containing DateIn, using
'   quarters Jan-Mar, Apr-Jun, Jul-Sep and Oct-Dec.
'
' WHY THIS EXISTS
'   The facade element Elem_BeginOfQuarter takes the calendar result from
'   here and applies the window gate itself.
'
' INPUTS
'   DateIn
'     Already-parsed VBA Date.
'
' RETURNS
'   Date
'     The first day of the containing quarter: 1 January, 1 April, 1 July or
'     1 October of the year of DateIn.
'
' BEHAVIOR
'   Maps the month of DateIn to the quarter start month 1, 4, 7 or 10 and
'   builds day 1 of that month.
'
' ERROR POLICY
'   None. The year and month come from a valid Date, so the date is always
'   constructible.
'
' DEPENDENCIES
'   None
'
' NOTES
'   - Pure calendar function: no window gate. A Q1-1900 input yields
'     1900-01-01, which is a valid VBA Date but outside the supported window;
'     the facade element gates the result.
'
' UPDATED
'   2026-09-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Quarter start month is 1, 4, 7 or 10
        BeginOfQuarter = VBA.DateSerial(VBA.Year(DateIn), ((VBA.Month(DateIn) - 1) \ 3) * 3 + 1, 1)

End Function

Public Function EndOfQuarter(ByVal DateIn As Date) As Date

'
'------------------------------------------------------------------------------
'                      RETURN THE LAST DAY OF THE QUARTER
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the last day of the calendar quarter containing DateIn.
'
' WHY THIS EXISTS
'   The facade element Elem_EndOfQuarter takes the quarter end from here, so
'   the month length is read from DaysInMonth like the rest of the module.
'
' INPUTS
'   DateIn
'     Already-parsed VBA Date.
'
' RETURNS
'   Date
'     The last day of the containing quarter: 31 March, 30 June,
'     30 September or 31 December of the year of DateIn.
'
' BEHAVIOR
'   Maps the month of DateIn to the quarter end month 3, 6, 9 or 12 and
'   builds the last day of that month from DaysInMonth.
'
' ERROR POLICY
'   None. The year and month come from a valid Date, so the date is always
'   constructible.
'
' DEPENDENCIES
'   DaysInMonth
'
' NOTES
'   - Month length comes from DaysInMonth, never from the day-zero idiom.
'
' UPDATED
'   2026-09-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Dates.EndOfQuarter"

    Dim LastMonth       As Long     'Last month of the quarter: 3, 6, 9 or 12

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Quarter end month, then its length
        LastMonth = ((VBA.Month(DateIn) - 1) \ 3) * 3 + 3
        EndOfQuarter = VBA.DateSerial(VBA.Year(DateIn), LastMonth, DaysInMonth(VBA.Year(DateIn), LastMonth))

End Function

Public Function BeginOfYear(ByVal DateIn As Date) As Date

'
'------------------------------------------------------------------------------
'                       RETURN THE FIRST DAY OF THE YEAR
'------------------------------------------------------------------------------
' PURPOSE
'   Returns 1 January of the year containing DateIn.
'
' WHY THIS EXISTS
'   The facade element Elem_BeginOfYear takes the calendar result from here
'   and applies the window gate itself.
'
' INPUTS
'   DateIn
'     Already-parsed VBA Date.
'
' RETURNS
'   Date
'     1 January of the year of DateIn.
'
' BEHAVIOR
'   Builds day 1 of month 1 in the year of DateIn.
'
' ERROR POLICY
'   None. The year comes from a valid Date, so the date is always
'   constructible.
'
' DEPENDENCIES
'   None
'
' NOTES
'   - Pure calendar function: no window gate. Any 1900 input yields
'     1900-01-01, which the facade element gates as RESULT_WINDOW.
'
' UPDATED
'   2026-09-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'First day of the containing year
        BeginOfYear = VBA.DateSerial(VBA.Year(DateIn), 1, 1)

End Function

Public Function EndOfYear(ByVal DateIn As Date) As Date

'
'------------------------------------------------------------------------------
'                       RETURN THE LAST DAY OF THE YEAR
'------------------------------------------------------------------------------
' PURPOSE
'   Returns 31 December of the year containing DateIn.
'
' WHY THIS EXISTS
'   The facade element Elem_EndOfYear takes the year end from here.
'
' INPUTS
'   DateIn
'     Already-parsed VBA Date.
'
' RETURNS
'   Date
'     31 December of the year of DateIn.
'
' BEHAVIOR
'   Builds day 31 of month 12 in the year of DateIn.
'
' ERROR POLICY
'   None. The year comes from a valid Date, so the date is always
'   constructible.
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-09-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Last day of the containing year
        EndOfYear = VBA.DateSerial(VBA.Year(DateIn), 12, 31)

End Function

Public Function IsMonthEnd(ByVal DateIn As Date) As Boolean

'
'------------------------------------------------------------------------------
'                             TEST FOR A MONTH END
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a date is the last day of its month.
'
' WHY THIS EXISTS
'   KPR_Dates_IsMonthEnd tests the boundary that the period boundaries define.
'
' INPUTS
'   DateIn
'     A date inside the supported window.
'
' RETURNS
'   Boolean
'     True when DateIn is the last day of its month.
'
' BEHAVIOR
'   Compares the day, and the month where the period needs it, with the period's end.
'
' ERROR POLICY
'   None. Total for the dates of the window.
'
' DEPENDENCIES
'   DaysInMonth. CALC-05 step 1
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'The last day of the month
        IsMonthEnd = (VBA.Day(DateIn) = DaysInMonth(VBA.Year(DateIn), VBA.Month(DateIn)))

End Function

Public Function IsQuarterEnd(ByVal DateIn As Date) As Boolean

'
'------------------------------------------------------------------------------
'                            TEST FOR A QUARTER END
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a date is the last day of its quarter.
'
' WHY THIS EXISTS
'   KPR_Dates_IsQuarterEnd tests the boundary that the period boundaries define.
'
' INPUTS
'   DateIn
'     A date inside the supported window.
'
' RETURNS
'   Boolean
'     True when DateIn is the last day of its quarter.
'
' BEHAVIOR
'   Compares the day, and the month where the period needs it, with the period's end.
'
' ERROR POLICY
'   None. Total for the dates of the window.
'
' DEPENDENCIES
'   DaysInMonth. IsMonthEnd. CALC-05 step 1
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'The last day of the quarter
        IsQuarterEnd = (VBA.Month(DateIn) Mod 3 = 0) And IsMonthEnd(DateIn)

End Function

Public Function IsYearEnd(ByVal DateIn As Date) As Boolean

'
'------------------------------------------------------------------------------
'                             TEST FOR A YEAR END
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a date is the last day of its year.
'
' WHY THIS EXISTS
'   KPR_Dates_IsYearEnd tests the boundary that the period boundaries define.
'
' INPUTS
'   DateIn
'     A date inside the supported window.
'
' RETURNS
'   Boolean
'     True when DateIn is the last day of its year.
'
' BEHAVIOR
'   Compares the day, and the month where the period needs it, with the period's end.
'
' ERROR POLICY
'   None. Total for the dates of the window.
'
' DEPENDENCIES
'   DaysInMonth. CALC-05 step 1
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'The last day of the year
        IsYearEnd = (VBA.Month(DateIn) = 12) And (VBA.Day(DateIn) = 31)

End Function

'
'------------------------------------------------------------------------------
'
'                                 DATE SHIFTS
'
'------------------------------------------------------------------------------
'

Public Function TryAddDays( _
    ByVal DateIn As Date, _
    ByVal Days As Double, _
    ByRef DateOut As Date) As Boolean
'
'------------------------------------------------------------------------------
'                             SHIFT A DATE BY DAYS
'------------------------------------------------------------------------------
' PURPOSE
'   Adds a signed number of calendar days to a date, if the result is inside the window.
'
' WHY THIS EXISTS
'   AddDays and AddWeeks differ only in the factor, so one shifter serves both and gates
'   the result once.
'
' INPUTS
'   DateIn
'     A date inside the supported window; a time of day is discarded.
'
'   Days
'     Signed number of days. A Double, so that seven times the largest Long
'     (AddWeeks) cannot overflow before the window rejects it.
'
'   DateOut (ByRef, output)
'     Assigned only on success.
'
' RETURNS
'   Boolean
'     True when the result is inside the window; False otherwise, with DateOut untouched.
'
' BEHAVIOR
'   Adds Days to the serial of DateIn and tests the sum against the window before it
'   builds a Date.
'
' ERROR POLICY
'   Does not raise. A result outside the window is False, which the facade reports as
'   RESULT_WINDOW.
'
' DEPENDENCIES
'   CORE_Parse.WINDOW_MIN_SERIAL, CORE_Parse.WINDOW_MAX_SERIAL. CALC-06 steps 1, 2 and 6
'
' NOTES
'   AddWeeks passes seven times its count
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Dates.TryAddDays"

    Dim ResultSerial    As Double   'Shifted serial, gated before conversion

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default outcome is failure
        TryAddDays = False

'------------------------------------------------------------------------------
' SHIFT
'------------------------------------------------------------------------------
    'Shift the whole day of the date
        ResultSerial = VBA.Int(VBA.CDbl(DateIn)) + Days
    'Gate the result before any Date is built
        If (ResultSerial < CORE_Parse.WINDOW_MIN_SERIAL) Or _
            (ResultSerial > CORE_Parse.WINDOW_MAX_SERIAL) Then Exit Function

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once the result is validated
        DateOut = VBA.CDate(ResultSerial)
        TryAddDays = True

End Function

Public Function TryAddMonths( _
    ByVal DateIn As Date, _
    ByVal nMonths As Double, _
    ByVal KeepEOM As Boolean, _
    ByRef DateOut As Date) As Boolean
'
'------------------------------------------------------------------------------
'                         SHIFT A DATE BY WHOLE MONTHS
'------------------------------------------------------------------------------
' PURPOSE
'   Single month shifter for the module. Adds nMonths to an already-parsed
'   date with explicit end-of-month semantics and full range safety.
'
' WHY THIS EXISTS
'   AddMonths, AddYears and the pillar resolver all need the same shift with
'   the same edge-case behavior. Duplicating it three times is how the three
'   drift apart.
'
' INPUTS
'   DateIn
'     Already-parsed VBA Date. A time component is tolerated and discarded.
'
'   nMonths
'     Signed month delta. A Double, so that twelve times the largest Long
'     (AddYears) cannot overflow before the gate rejects it.
'
'   KeepEOM
'     TRUE  => if DateIn is month-end, return the target month-end
'     FALSE => clip day-of-month to the target month length
'
'   DateOut (ByRef, output)
'     Assigned ONLY on success.
'
' RETURNS
'   Boolean
'     TRUE  => shift succeeded and the result is inside the supported window
'     FALSE => the target month or the result falls outside it
'
' BEHAVIOR
'   - Converts year / month into an absolute month index, shifts it, and gates
'     the index before it is ever handed to DateSerial. This is why an absurd
'     nMonths returns FALSE instead of raising an overflow.
'   - Resolves the day-of-month by the KeepEOM rule, then re-gates the result.
'
' ERROR POLICY
'   - Does not raise. All failure paths return FALSE.
'   - Clears Err before returning FALSE, so no stale error state survives.
'
' DEPENDENCIES
'   DaysInMonth
'
' NOTES
'   - DateOut is safe to pass as the same variable as an input date at the call
'     site, because it is written only once, at the end.
'   - Both month lengths come from DaysInMonth, so neither the EOM anchor test
'     nor the day clip constructs a date outside the month it is measuring.
'   - The index gate is deliberately coarse, admitting whole years at both
'     ends. The terminal gate is what enforces the exact window, so if either
'     bound ever moves off a year boundary the pair still holds.
'   - The index is computed in Double so that an extreme nMonths cannot
'     overflow before the gate can reject it.
'
' UPDATED
'   2026-09-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Dates.TryAddMonths"

    Dim Y               As Long     'Year component of DateIn
    Dim M               As Long     'Month component of DateIn
    Dim DayIn           As Long     'Day-of-month of DateIn
    Dim IsEOM           As Boolean  'TRUE when DateIn is its own month-end

    Dim MonthIndex      As Double   'Absolute 0-based month index after the shift

    Dim TargetY         As Long     'Resolved target year
    Dim TargetM         As Long     'Resolved target month
    Dim TargetDim       As Long     'Days in the target month
    Dim TargetD         As Long     'Resolved day-of-month in the target month

    Dim Candidate       As Date     'Shifted date before the terminal range gate

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Trap runtime errors and convert them to a FALSE return
        On Error GoTo Fail
    'Default return is failure unless we explicitly succeed
        TryAddMonths = False

'------------------------------------------------------------------------------
' CACHE INPUT COMPONENTS
'------------------------------------------------------------------------------
    'Split the input date once
        Y = VBA.Year(DateIn)
        M = VBA.Month(DateIn)
        DayIn = VBA.Day(DateIn)
    'Detect the EOM anchor before shifting
        IsEOM = (DayIn = DaysInMonth(Y, M))

'------------------------------------------------------------------------------
' RESOLVE TARGET MONTH
'------------------------------------------------------------------------------
    'Absolute 0-based month index, shifted, computed in Double
        MonthIndex = (VBA.CDbl(Y) * 12# + VBA.CDbl(M) - 1#) + nMonths
    'Gate the index before any coercion so DateSerial never sees a wild value
        If (MonthIndex < (VBA.CDbl(CORE_Parse.WINDOW_MIN_YEAR) * 12#)) Or _
            (MonthIndex > VBA.CDbl(MAX_MONTH_INDEX)) Then GoTo Fail
    'Split the gated index back into year and month
        TargetY = VBA.CLng(VBA.Int(MonthIndex / 12#))
        TargetM = VBA.CLng(MonthIndex - (VBA.CDbl(TargetY) * 12#)) + 1

'------------------------------------------------------------------------------
' RESOLVE DAY OF MONTH
'------------------------------------------------------------------------------
    'Target month length, read without constructing a date
        TargetDim = DaysInMonth(TargetY, TargetM)
    'Preserve EOM when asked and the input was EOM; otherwise clip
        If KeepEOM And IsEOM Then
            TargetD = TargetDim
        Else
            TargetD = DayIn
            If TargetD > TargetDim Then TargetD = TargetDim
        End If

'------------------------------------------------------------------------------
' TERMINAL RANGE GATE
'------------------------------------------------------------------------------
    'Build the candidate and gate it against the supported window
        Candidate = VBA.DateSerial(TargetY, TargetM, TargetD)
        If (VBA.CDbl(Candidate) < CORE_Parse.WINDOW_MIN_SERIAL) Or _
            (VBA.CDbl(Candidate) > CORE_Parse.WINDOW_MAX_SERIAL) Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once the value is fully validated
        DateOut = Candidate
    'Contract: TRUE only when DateOut was assigned
        TryAddMonths = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the failure path
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Fail:
    'Return FALSE per contract, leaving DateOut untouched
        TryAddMonths = False
    'Do not leave stale error state for the caller to observe
        Err.Clear

End Function

Public Function TryNthWeekday( _
    ByVal YearIn As Long, _
    ByVal MonthIn As Long, _
    ByVal WeekdayIn As Long, _
    ByVal Occurrence As Long, _
    ByVal MondayBase As Boolean, _
    ByRef DateOut As Date, _
    ByRef Condition As KPRCondition) As Boolean
'
'------------------------------------------------------------------------------
'                       FIND THE NTH WEEKDAY OF A MONTH
'------------------------------------------------------------------------------
' PURPOSE
'   Finds the date of the nth given weekday of a month.
'
' WHY THIS EXISTS
'   KPR_Dates_NthWeekdayOfMonth needs the date, the absent fifth weekday and the window
'   result decided in one place.
'
' INPUTS
'   YearIn, MonthIn, WeekdayIn, Occurrence
'     Parsed whole numbers. Each is checked against its domain, in this order, and the
'     first that fails is reported.
'
'   MondayBase
'     The weekday base of WeekdayIn.
'
'   DateOut (ByRef, output)
'     Assigned only on success.
'
'   Condition (ByRef, output)
'     Always assigned: cdNone, or the failing condition.
'
' RETURNS
'   Boolean
'     True when DateOut was assigned. False with cdDomainYear, cdDomainMonth,
'     cdDomainWeekday, cdDomainOccurrence, cdOccurrenceAbsent (the month has fewer
'     than Occurrence of that weekday) or cdResultWindow (the date is before the window).
'
' BEHAVIOR
'   Counts from the month's first day to the first matching weekday, then adds seven
'   days for each further occurrence, and tests the occurrence against the month's
'   length before the window.
'
' ERROR POLICY
'   Does not raise. Every failure is False with a condition.
'
' DEPENDENCIES
'   CheckYear, CheckMonth, CheckWeekday, CheckOccurrence, DaysInMonth, VBA.Weekday,
'   CORE_Parse window constants. CALC-07 steps 1, 2, 3 and 5
'
' NOTES
'   The window can reject only a result in January or February 1900
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Dates.TryNthWeekday"

    Dim FirstDay        As Date         'First day of the month
    Dim FirstSerial     As Double       'Its serial
    Dim Offset          As Long         'Days from the first day to the first matching weekday
    Dim ResultSerial    As Double       'Candidate serial, gated before conversion
    Dim BaseDay         As VbDayOfWeek  'Weekday that numbers as 1

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default outcome is failure with an explicit condition
        TryNthWeekday = False
        Condition = cdNone

'------------------------------------------------------------------------------
' CHECK DOMAINS
'------------------------------------------------------------------------------
    'Each argument in signature order; the first failure decides
        Condition = CheckYear(YearIn)
        If Condition <> cdNone Then Exit Function
        Condition = CheckMonth(MonthIn)
        If Condition <> cdNone Then Exit Function
        Condition = CheckWeekday(WeekdayIn)
        If Condition <> cdNone Then Exit Function
        Condition = CheckOccurrence(Occurrence)
        If Condition <> cdNone Then Exit Function

'------------------------------------------------------------------------------
' LOCATE
'------------------------------------------------------------------------------
    'Select the weekday base
        If MondayBase Then
            BaseDay = vbMonday
        Else
            BaseDay = vbSunday
        End If
    'Days from the first day of the month to the first matching weekday
        FirstDay = VBA.DateSerial(YearIn, MonthIn, 1)
        FirstSerial = VBA.CDbl(FirstDay)
        Offset = (WeekdayIn - VBA.Weekday(FirstDay, BaseDay) + 7) Mod 7
    'Add a week for each further occurrence
        ResultSerial = FirstSerial + VBA.CDbl(Offset) + (7# * VBA.CDbl(Occurrence - 1))
    'A fifth weekday that the month does not have is absent
        If ResultSerial > (FirstSerial + VBA.CDbl(DaysInMonth(YearIn, MonthIn)) - 1#) Then
            Condition = cdOccurrenceAbsent
            Exit Function
        End If
    'Gate the result before any Date is built
        If (ResultSerial < CORE_Parse.WINDOW_MIN_SERIAL) Or _
            (ResultSerial > CORE_Parse.WINDOW_MAX_SERIAL) Then
            Condition = cdResultWindow
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once the result is validated
        DateOut = VBA.CDate(ResultSerial)
        TryNthWeekday = True

End Function

Public Function TryLastWeekday( _
    ByVal YearIn As Long, _
    ByVal MonthIn As Long, _
    ByVal WeekdayIn As Long, _
    ByVal MondayBase As Boolean, _
    ByRef DateOut As Date, _
    ByRef Condition As KPRCondition) As Boolean
'
'------------------------------------------------------------------------------
'                       FIND THE LAST WEEKDAY OF A MONTH
'------------------------------------------------------------------------------
' PURPOSE
'   Finds the date of the last given weekday of a month.
'
' WHY THIS EXISTS
'   KPR_Dates_LastWeekdayOfMonth needs the date and the window result decided in one place.
'
' INPUTS
'   YearIn, MonthIn, WeekdayIn
'     Parsed whole numbers. Each is checked against its domain, in this order, and the
'     first that fails is reported.
'
'   MondayBase
'     The weekday base of WeekdayIn.
'
'   DateOut (ByRef, output)
'     Assigned only on success.
'
'   Condition (ByRef, output)
'     Always assigned: cdNone, or the failing condition.
'
' RETURNS
'   Boolean
'     True when DateOut was assigned. False with cdDomainYear, cdDomainMonth,
'     cdDomainWeekday or cdResultWindow (the date is before the window).
'
' BEHAVIOR
'   Counts back from the month's last day to the given weekday.
'
' ERROR POLICY
'   Does not raise. Every failure is False with a condition.
'
' DEPENDENCIES
'   CheckYear, CheckMonth, CheckWeekday, DaysInMonth, VBA.Weekday, CORE_Parse window
'   constants. CALC-07 steps 1, 4 and 5
'
' NOTES
'   The window can reject only a result in January or February 1900
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Dates.TryLastWeekday"

    Dim LastDay         As Date         'Last day of the month
    Dim Back            As Long         'Days from the last day back to the weekday
    Dim ResultSerial    As Double       'Candidate serial, gated before conversion
    Dim BaseDay         As VbDayOfWeek  'Weekday that numbers as 1

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default outcome is failure with an explicit condition
        TryLastWeekday = False
        Condition = cdNone

'------------------------------------------------------------------------------
' CHECK DOMAINS
'------------------------------------------------------------------------------
    'Each argument in signature order; the first failure decides
        Condition = CheckYear(YearIn)
        If Condition <> cdNone Then Exit Function
        Condition = CheckMonth(MonthIn)
        If Condition <> cdNone Then Exit Function
        Condition = CheckWeekday(WeekdayIn)
        If Condition <> cdNone Then Exit Function

'------------------------------------------------------------------------------
' LOCATE
'------------------------------------------------------------------------------
    'Select the weekday base
        If MondayBase Then
            BaseDay = vbMonday
        Else
            BaseDay = vbSunday
        End If
    'Days back from the last day of the month to the weekday
        LastDay = VBA.DateSerial(YearIn, MonthIn, DaysInMonth(YearIn, MonthIn))
        Back = (VBA.Weekday(LastDay, BaseDay) - WeekdayIn + 7) Mod 7
        ResultSerial = VBA.CDbl(LastDay) - VBA.CDbl(Back)
    'Gate the result before any Date is built
        If (ResultSerial < CORE_Parse.WINDOW_MIN_SERIAL) Or _
            (ResultSerial > CORE_Parse.WINDOW_MAX_SERIAL) Then
            Condition = cdResultWindow
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once the result is validated
        DateOut = VBA.CDate(ResultSerial)
        TryLastWeekday = True

End Function

'
'------------------------------------------------------------------------------
'
'                                   PILLARS
'
'------------------------------------------------------------------------------
'

Public Function TryParsePillar( _
    ByVal PillarIn As Variant, _
    ByRef TotalMonths As Double, _
    ByRef TotalDays As Double, _
    ByRef Condition As KPRCondition) As Boolean
'
'------------------------------------------------------------------------------
'                             PARSE A PILLAR TOKEN
'------------------------------------------------------------------------------
' PURPOSE
'   Parses a pillar token into a signed month delta and a signed day delta.
'
' WHY THIS EXISTS
'   The facade reads pillar tokens through this one parser, so every grammar
'   rejection is classified under its own condition identifier in one place
'   (data contract, Pillar tokens).
'
' INPUTS
'   PillarIn
'     Scalar Variant expected to hold pillar text.
'
'   TotalMonths (ByRef, output)
'     Signed aggregate of the Y and M components: sign * (12Y + M).
'     Assigned ONLY on success.
'
'   TotalDays (ByRef, output)
'     Signed aggregate of the W and D components: sign * (7W + D).
'     Assigned ONLY on success.
'
'   Condition (ByRef, output)
'     Always assigned. cdNone on success; otherwise one of
'       PILLAR_TYPE_REJECTED     non-text payload
'       PILLAR_ALIAS_SIGNED      whole-token alias carrying a leading sign
'       PILLAR_DUPLICATE_UNIT    a unit appearing more than once
'       PILLAR_AGGREGATE_RANGE   valid numeric component/aggregate out of range
'       PILLAR_TOKEN_MALFORMED   every other grammar violation
'
' RETURNS
'   Boolean
'     TRUE  => the token matched the accepted grammar
'     FALSE => blank, malformed, unknown unit, repeated unit, signed alias,
'              non-text payload, or valid numeric component/aggregate out of range
'
' BEHAVIOR
'   - Trims and upper-cases, consumes an optional leading sign, then checks the
'     whole-token aliases ON / O/N / TN / T/N.
'   - Otherwise validates every repeated [digits][unit] pair to the end of the
'     string before converting any numeric quantity.
'   - Any character that is not a digit in a numeric position or a known unit
'     in a unit position fails the parse.
'   - Each unit may appear at most once, so "1M1M" fails rather than summing.
'
' ERROR POLICY
'   - Does not raise. All failure paths return FALSE.
'   - Clears Err before returning FALSE, so no stale error state survives.
'
' DEPENDENCIES
'   CORE_Err condition vocabulary; Pillar_AsciiTrim, Pillar_AsciiUpper
'
' NOTES
'   - No characters are stripped from the body, so "1 M" and "1/M" fail rather
'     than being silently reinterpreted as "1M". The aliases are the only
'     tokens containing "/", and they are matched whole.
'   - Fractional tenors are not part of the grammar: "1.5M" fails. Half months
'     are not a market convention, and admitting them would force a rounding
'     policy into a parser.
'   - A leading sign on an alias is rejected. The aliases name fixed points at
'     the short end, not quantities, so "-ON" has no meaning to negate.
'   - A repeated unit is a typo, not a sum. Accepting "1M1M" as two months
'     would return a plausible number for input the caller did not intend;
'     the data contract therefore rejects duplicate units.
'   - Finite quantities are not bounded here. A token such as "999999999999M"
'     parses; the facade rejects a month aggregate outside the Long range as
'     PILLAR_AGGREGATE_RANGE, and TryAddMonths' month-index gate rejects any
'     remaining out-of-window shift. A digits-only component that cannot be
'     represented as Double, or an aggregate that overflows Double, is
'     PILLAR_AGGREGATE_RANGE.
'   - CDbl is applied to a digits-only substring, so no locale decimal
'     separator is involved. Replacing it with Val or CLng would change the
'     overflow behaviour.
'   - The digit scan compares character codes rather than one-character
'     strings. The unit test stays on strings because it runs once per
'     component rather than once per character.
'   - A signed result can be negative zero, because the sign multiplier is
'     applied to a zero aggregate. It compares equal to zero and behaves
'     identically in arithmetic; only a formatter would notice.
'
' UPDATED
'   2026-10-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Dates.TryParsePillar"

    Dim S               As String   'Upper-cased, trimmed pillar text
    Dim SBody           As String   'Pillar body after the optional leading sign
    Dim ChCode          As Long     'Character code at the current scan position
    Dim UnitChar        As String   'Validated unit for the current numeric token
    Dim HasSign         As Boolean  'TRUE when a leading sign was consumed

    Dim ScanPos         As Long     'Current scan position in SBody
    Dim TokenStart      As Long     'Start position of the current numeric token
    Dim TokenCount      As Long     'Number of components parsed so far
    Dim BodyLen         As Long     'Length of SBody, read once

    Dim SignMul         As Double   'Global sign multiplier (+1 or -1)
    Dim QtyD            As Double   'Current component quantity

    Dim SeenY           As Boolean  'TRUE once a Y component has been consumed
    Dim SeenM           As Boolean  'TRUE once an M component has been consumed
    Dim SeenW           As Boolean  'TRUE once a W component has been consumed
    Dim SeenD           As Boolean  'TRUE once a D component has been consumed

    Dim YearsD          As Double   'Y component
    Dim MonthsD         As Double   'M component
    Dim WeeksD          As Double   'W component
    Dim DaysD           As Double   'D component, or the resolved alias
    Dim ParsedMonths    As Double   'Local month aggregate; assigned out only on success
    Dim ParsedDays      As Double   'Local day aggregate; assigned out only on success

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Trap runtime errors and convert them to a FALSE return
        On Error GoTo Fail
    'Default classification for any structural rejection below
        Condition = cdPillarTokenMalformed
    'Default return is failure unless we explicitly succeed
        TryParsePillar = False
    'Default sign is positive
        SignMul = 1#

'------------------------------------------------------------------------------
' NORMALIZE PILLAR TEXT
'------------------------------------------------------------------------------
    'Require a text payload
        If VBA.VarType(PillarIn) <> vbString Then
            Condition = cdPillarTypeRejected
            GoTo Fail
        End If
    'Trim and upper-case once
        S = CORE_Parse.AsciiUpper(CORE_Parse.AsciiTrim(VBA.CStr(PillarIn)))
    'Reject a blank token
        If VBA.Len(S) = 0 Then GoTo Fail

'------------------------------------------------------------------------------
' CONSUME OPTIONAL LEADING SIGN
'------------------------------------------------------------------------------
    'A sign applies to the whole token
        Select Case VBA.Left$(S, 1)
            Case "+"
                HasSign = True
                SBody = VBA.Mid$(S, 2)
            Case "-"
                HasSign = True
                SignMul = -1#
                SBody = VBA.Mid$(S, 2)
            Case Else
                SBody = S
        End Select
    'Reject a sign with no body
        BodyLen = VBA.Len(SBody)
        If BodyLen = 0 Then GoTo Fail

'------------------------------------------------------------------------------
' WHOLE-TOKEN ALIASES
'------------------------------------------------------------------------------
    'Overnight and tom-next are matched whole, never as generic tokens
        Select Case SBody
            Case "ON", "O/N", "TN", "T/N"
                'An alias names a fixed point and cannot carry a sign
                If HasSign Then
                    Condition = cdPillarAliasSigned
                    GoTo Fail
                End If
                'Resolve the alias to its day count
                If (SBody = "ON") Or (SBody = "O/N") Then
                    DaysD = 1#
                Else
                    DaysD = 2#
                End If
                TokenCount = 1
            Case Else
        'Not an alias: the generic scanner below reads the token
        End Select

'------------------------------------------------------------------------------
' GENERIC INTEGER + UNIT SCANNER
'------------------------------------------------------------------------------
    'Only run the scanner when no alias matched
        If TokenCount = 0 Then

'------------------------------------------------------------------------------
' PASS 1 - VALIDATE THE COMPLETE TOKEN GRAMMAR
'------------------------------------------------------------------------------
            'Grammar has absolute precedence over numerical range. Scan the
            'entire token first so a later unknown/repeated/missing unit wins
            'even when an earlier digits-only quantity would overflow Double.
            ScanPos = 1
            Do While ScanPos <= BodyLen
                TokenStart = ScanPos
                Do While ScanPos <= BodyLen
                    ChCode = VBA.AscW(VBA.Mid$(SBody, ScanPos, 1))
                    If (ChCode < 48) Or (ChCode > 57) Then Exit Do
                    ScanPos = ScanPos + 1
                Loop
                If TokenStart = ScanPos Then GoTo Fail
                If ScanPos > BodyLen Then GoTo Fail

                UnitChar = VBA.Mid$(SBody, ScanPos, 1)
                Select Case UnitChar
                    Case "Y"
                        If SeenY Then
                            Condition = cdPillarDuplicateUnit
                            GoTo Fail
                        End If
                        SeenY = True
                    Case "M"
                        If SeenM Then
                            Condition = cdPillarDuplicateUnit
                            GoTo Fail
                        End If
                        SeenM = True
                    Case "W"
                        If SeenW Then
                            Condition = cdPillarDuplicateUnit
                            GoTo Fail
                        End If
                        SeenW = True
                    Case "D"
                        If SeenD Then
                            Condition = cdPillarDuplicateUnit
                            GoTo Fail
                        End If
                        SeenD = True
                    Case Else
                        GoTo Fail
                End Select

                TokenCount = TokenCount + 1
                ScanPos = ScanPos + 1
            Loop

'------------------------------------------------------------------------------
' PASS 2 - CONVERT VALIDATED QUANTITIES
'------------------------------------------------------------------------------
            'Only a token whose complete grammar is valid reaches conversion.
            'A conversion failure therefore unambiguously means numerical range.
            ScanPos = 1
            Do While ScanPos <= BodyLen
                TokenStart = ScanPos
                Do While ScanPos <= BodyLen
                    ChCode = VBA.AscW(VBA.Mid$(SBody, ScanPos, 1))
                    If (ChCode < 48) Or (ChCode > 57) Then Exit Do
                    ScanPos = ScanPos + 1
                Loop

                UnitChar = VBA.Mid$(SBody, ScanPos, 1)
                On Error GoTo RangeFail
                QtyD = VBA.CDbl(VBA.Mid$(SBody, TokenStart, ScanPos - TokenStart))
                On Error GoTo Fail

                Select Case UnitChar
                    Case "Y"
                        YearsD = QtyD
                    Case "M"
                        MonthsD = QtyD
                    Case "W"
                        WeeksD = QtyD
                    Case "D"
                        DaysD = QtyD
                    Case Else
                'Unreachable: pass 1 admitted only Y, M, W and D
                End Select
                ScanPos = ScanPos + 1
            Loop

        End If

    'Unreachable in practice: the scanner fails before exiting with no
    'components, and a blank body was rejected earlier. Kept as a guard so
    'the TRUE return can never be reached without a parsed component.
        If TokenCount = 0 Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULTS
'------------------------------------------------------------------------------
    'Aggregate into locals first so a range failure never partially modifies
    'the caller's ByRef outputs.
        On Error GoTo RangeFail
        ParsedMonths = SignMul * ((12# * YearsD) + MonthsD)
        ParsedDays = SignMul * ((7# * WeeksD) + DaysD)
        On Error GoTo Fail
    'Contract: both outputs are assigned only after both aggregates succeed.
        TotalMonths = ParsedMonths
        TotalDays = ParsedDays
        Condition = cdNone
        TryParsePillar = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the failure paths
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
RangeFail:
    'The token grammar is valid; the numeric component or aggregate is not
    'representable in the parser's Double domain.
        Condition = cdPillarAggregateRange
        Resume Fail

Fail:
    'Return FALSE per contract, leaving the outputs untouched
        TryParsePillar = False
    'Do not leave stale error state for the caller to observe
        Err.Clear

End Function

Public Function TryAddPillar( _
    ByVal StartDate As Date, _
    ByVal TotalMonths As Double, _
    ByVal TotalDays As Double, _
    ByRef DateOut As Date, _
    ByRef Condition As KPRCondition) As Boolean
'
'------------------------------------------------------------------------------
'                        ADD A PARSED PILLAR TO A DATE
'------------------------------------------------------------------------------
' PURPOSE
'   Moves a date by the month and day deltas of a parsed pillar token.
'
' WHY THIS EXISTS
'   KPR_Dates_DateFromPillar applies the month delta first, with the month clip, and
'   then the day delta, and gates the result as a pillar total.
'
' INPUTS
'   StartDate
'     A date inside the supported window.
'
'   TotalMonths, TotalDays
'     The signed deltas that TryParsePillar returned.
'
'   DateOut (ByRef, output)
'     Assigned only on success.
'
'   Condition (ByRef, output)
'     Always assigned: cdNone, or cdPillarAggregateRange.
'
' RETURNS
'   Boolean
'     True when DateOut was assigned. False with cdPillarAggregateRange when the month
'     shift or the result leaves the window.
'
' BEHAVIOR
'   Applies TotalMonths with TryAddMonths and no month-end keeping, then adds
'   TotalDays and gates the sum.
'
' ERROR POLICY
'   Does not raise. Every failure is False with a condition.
'
' DEPENDENCIES
'   TryAddMonths, CORE_Parse window constants. CALC-08 steps 2, 3 and 4
'
' NOTES
'   A month delta beyond the Long range is beyond the window too, so TryAddMonths
'   rejects it with the same condition
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Dates.TryAddPillar"

    Dim WorkDate        As Date     'Date after the month shift
    Dim ResultSerial    As Double   'Final serial, gated before conversion

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default outcome is failure with an explicit condition
        TryAddPillar = False
        Condition = cdPillarAggregateRange
        WorkDate = StartDate

'------------------------------------------------------------------------------
' SHIFT
'------------------------------------------------------------------------------
    'Months first, with the clip of a short month
        If TotalMonths <> 0# Then
            If Not TryAddMonths(WorkDate, TotalMonths, False, WorkDate) Then Exit Function
        End If
    'Then the exact days, gated before any Date is built
        ResultSerial = VBA.CDbl(WorkDate) + TotalDays
        If (ResultSerial < CORE_Parse.WINDOW_MIN_SERIAL) Or _
            (ResultSerial > CORE_Parse.WINDOW_MAX_SERIAL) Then Exit Function

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once the result is validated
        DateOut = VBA.CDate(ResultSerial)
        Condition = cdNone
        TryAddPillar = True

End Function

Public Function TryFormatPillar( _
    ByVal DtStart As Date, _
    ByVal DtEnd As Date, _
    ByVal Mode As KPRPillarRounding, _
    ByRef TokenOut As String, _
    ByRef Condition As KPRCondition) As Boolean
'
'------------------------------------------------------------------------------
'                     FORMAT AN INTERVAL AS A PILLAR TOKEN
'------------------------------------------------------------------------------
' PURPOSE
'   Formats the interval from DtStart to DtEnd as one canonical pillar token
'   under the requested rounding mode, or reports why no token exists.
'
' WHY THIS EXISTS
'   The three rounding modes share one candidate set and one set of tie rules
'   (calculation contract CALC-09). Implementing them in one function keeps the modes
'   from drifting apart.
'
' INPUTS
'   DtStart, DtEnd
'     Date-only values inside the supported window. Order is free; a negative
'     interval yields a "-" prefix.
'
'   Mode
'     prNearest, prFloor or prCeiling.
'
'   TokenOut (ByRef, output)
'     Assigned ONLY on success.
'
'   Condition (ByRef, output)
'     Always assigned; cdNone on success.
'
' RETURNS
'   Boolean
'     TRUE  => TokenOut assigned
'     FALSE => Condition is RESULT_WINDOW: the anchor the mode requires would
'              leave the supported window, so no candidate satisfies it
'
' BEHAVIOR
'   - Returns the exact day token for an interval shorter than seven days.
'   - Otherwise builds the week and month candidates, selects one under the
'     mode, and formats it. CANDIDATE SET and SELECTION below give the rules.
'
' ERROR POLICY
'   - Does not raise. The single failure path returns FALSE with a condition.
'   - A local handler converts any unexpected runtime error into that same
'     FALSE return, so a raise can never escape to facade containment.
'
' DEPENDENCIES
'   TryAddMonths, CORE_Parse window constants; AnchorPasses, TieToMonth,
'   SignPrefix; CORE_Err condition vocabulary
'
' NOTES
'   - Month counts of 12 or more format as years plus residual months
'   - Candidate set, one under every mode: anchors are generated from the original
'     DtStart in the direction of DtEnd, with signed counts, so that an emitted
'     negative token reads back with the clipping that TryAddPillar applies. Weeks:
'     1W, 2W and 3W, never 4W or beyond. Months: the floor month, the largest
'     positive count whose anchor does not pass DtEnd in the interval direction, and
'     the ceiling month, the floor count plus one. When no positive count qualifies,
'     the floor count is 0, has no anchor and is not a candidate, and the ceiling
'     month is 1M. An anchor outside the supported window is excluded, never
'     approximated, because there is no nominal distance for an anchor that cannot be
'     built. Month anchors come from TryAddMonths, so their clipping cannot diverge
'     from the parser
'   - Selection: an interval shorter than seven days returns its exact day token
'     under every mode. FLOOR takes the candidate that does not pass DtEnd and lies
'     farthest from DtStart. CEILING takes the candidate that reaches or passes DtEnd
'     and lies closest to it, and none is RESULT_WINDOW. NEAREST takes the smallest
'     calendar-day distance to DtEnd. Ties are broken the same way in every mode:
'     a month beats a week, and between months the larger count wins. An exact
'     anchor wins each selection outright
'   - Under NEAREST the boundary of 3W and 1M is derived, not fixed: 24 days is
'     always 3W and 27 days is always 1M while the 1M anchor is in the window, and
'     25 and 26 days go to the closer anchor, with a tie to the month
'   - FLOOR and NEAREST always have a candidate: for an interval of seven days or
'     more, the 1W anchor lies between the two dates, which are both in the window.
'     Only CEILING can find every in-window candidate short of DtEnd
'
' UPDATED
'   2026-09-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Dates.TryFormatPillar"

    Dim Direction               As Long     '+1 when DtEnd follows DtStart, else -1
    Dim SerialStart             As Double   'Serial of DtStart
    Dim SerialEnd               As Double   'Serial of DtEnd
    Dim TotalDays               As Long     'Absolute calendar days between the dates

    Dim CandSerial(1 To 5)      As Double   'Anchor serials, in candidate order
    Dim CandIsMonth(1 To 5)     As Boolean  'TRUE for a month anchor
    Dim CandCount(1 To 5)       As Long     'Week or month count (unsigned)
    Dim CandN                   As Long     'Number of admitted candidates

    Dim K                       As Long     'Loop and count cursor
    Dim WeekSerial              As Double   'Candidate week anchor, tested before any Date conversion
    Dim FloorMonths             As Long     'Largest positive month count not passing DtEnd
    Dim Anchor                  As Date     'Working anchor from TryAddMonths
    Dim Probe                   As Long     'Working month count while locating the floor

    Dim Best                    As Long     'Index of the selected candidate, 0 = none
    Dim Reach                   As Double   'Signed progress of a candidate from DtStart
    Dim BestReach               As Double   'Progress of the current best
    Dim Dist                    As Long     'Absolute day distance of a candidate to DtEnd
    Dim BestDist                As Long     'Distance of the current best
    Dim Passes                  As Boolean  'TRUE when a candidate reaches or passes DtEnd
    Dim Better                  As Boolean  'TRUE when candidate K displaces Best

    Dim TokenYears              As Long     'Year component of a month count
    Dim TokenMonths             As Long     'Residual month component

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default outcome is failure with an explicit condition
        TryFormatPillar = False
        Condition = cdResultWindow
    'Trap runtime errors and convert them to that FALSE return
        On Error GoTo Fail

'------------------------------------------------------------------------------
' RESOLVE DIRECTION
'------------------------------------------------------------------------------
    'Direction and magnitude from the ORIGINAL start date
        SerialStart = VBA.Int(VBA.CDbl(DtStart))
        SerialEnd = VBA.Int(VBA.CDbl(DtEnd))
        If SerialEnd >= SerialStart Then Direction = 1 Else Direction = -1
        TotalDays = VBA.CLng(VBA.Abs(SerialEnd - SerialStart))

'------------------------------------------------------------------------------
' EXACT DAY PILLARS
'------------------------------------------------------------------------------
    'Equal dates and short intervals are exact under every mode
        If TotalDays = 0 Then
            TokenOut = "0D"
            Condition = cdNone
            TryFormatPillar = True
            Exit Function
        End If
        If TotalDays < 7 Then
            TokenOut = SignPrefix(Direction) & VBA.CStr(TotalDays) & "D"
            Condition = cdNone
            TryFormatPillar = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' WEEK CANDIDATES
'------------------------------------------------------------------------------
    'Admit 1W, 2W and 3W when their anchors stay inside the window. The test
    'is on the serial as a Double: converting an out-of-range serial to a Date
    'raises, and an anchor that cannot be constructed must be excluded, not
    'allowed to abort the call.
        CandN = 0
        For K = 1 To 3
            WeekSerial = SerialStart + VBA.CDbl(Direction) * 7# * VBA.CDbl(K)
            If (WeekSerial >= CORE_Parse.WINDOW_MIN_SERIAL) And (WeekSerial <= CORE_Parse.WINDOW_MAX_SERIAL) Then
                CandN = CandN + 1
                CandSerial(CandN) = WeekSerial
                CandIsMonth(CandN) = False
                CandCount(CandN) = K
            End If
        Next K

'------------------------------------------------------------------------------
' MONTH CANDIDATES
'------------------------------------------------------------------------------
    'Locate the largest positive count whose anchor does not pass DtEnd,
    'starting from the calendar-month difference and correcting for clipping
        FloorMonths = VBA.Abs(VBA.DateDiff("m", DtStart, DtEnd))
        If FloorMonths < 1 Then FloorMonths = 1
    'Step down while the anchor passes DtEnd or cannot be constructed
        Do While FloorMonths >= 1
            If TryAddMonths(DtStart, Direction * FloorMonths, False, Anchor) Then
                If Not AnchorPasses(VBA.CDbl(Anchor), SerialEnd, Direction) Then Exit Do
            End If
            FloorMonths = FloorMonths - 1
        Loop
    'Step up while the next anchor still does not pass DtEnd
        Do
            Probe = FloorMonths + 1
            If Not TryAddMonths(DtStart, Direction * Probe, False, Anchor) Then Exit Do
            If AnchorPasses(VBA.CDbl(Anchor), SerialEnd, Direction) Then Exit Do
            FloorMonths = Probe
        Loop
    'Admit the floor month anchor, if any
        If FloorMonths >= 1 Then
            If TryAddMonths(DtStart, Direction * FloorMonths, False, Anchor) Then
                CandN = CandN + 1
                CandSerial(CandN) = VBA.CDbl(Anchor)
                CandIsMonth(CandN) = True
                CandCount(CandN) = FloorMonths
            End If
        End If
    'Admit the ceiling month anchor only if it can be constructed in-window
        If TryAddMonths(DtStart, Direction * (FloorMonths + 1), False, Anchor) Then
            CandN = CandN + 1
            CandSerial(CandN) = VBA.CDbl(Anchor)
            CandIsMonth(CandN) = True
            CandCount(CandN) = FloorMonths + 1
        End If

'------------------------------------------------------------------------------
' SELECT
'------------------------------------------------------------------------------
    'One pass; the mode decides the ordering and the tie rules are shared
        Best = 0
        For K = 1 To CandN

            'Signed progress from DtStart and absolute distance to DtEnd
            Reach = VBA.CDbl(Direction) * (CandSerial(K) - SerialStart)
            Dist = VBA.CLng(VBA.Abs(CandSerial(K) - SerialEnd))
            Passes = AnchorPasses(CandSerial(K), SerialEnd, Direction) Or (Dist = 0)

            'Eligibility and ordering by mode
            Better = False
            Select Case Mode

                Case prFloor
                    'Must not pass; prefer the farthest reach
                    If Not AnchorPasses(CandSerial(K), SerialEnd, Direction) Then
                        If Best = 0 Then
                            Better = True
                        ElseIf Reach > BestReach Then
                            Better = True
                        ElseIf Reach = BestReach Then
                            Better = TieToMonth(K, Best, CandIsMonth, CandCount)
                        End If
                    End If

                Case prCeiling
                    'Must reach or pass; prefer the smallest reach
                    If Passes Then
                        If Best = 0 Then
                            Better = True
                        ElseIf Reach < BestReach Then
                            Better = True
                        ElseIf Reach = BestReach Then
                            Better = TieToMonth(K, Best, CandIsMonth, CandCount)
                        End If
                    End If

                Case Else
                    'NEAREST: smallest distance, ties by the shared rule
                    If Best = 0 Then
                        Better = True
                    ElseIf Dist < BestDist Then
                        Better = True
                    ElseIf Dist = BestDist Then
                        Better = TieToMonth(K, Best, CandIsMonth, CandCount)
                    End If

            End Select

            'Record the new best
            If Better Then
                Best = K
                BestReach = Reach
                BestDist = Dist
            End If

        Next K
    'No candidate satisfied the mode: only CEILING can arrive here
        If Best = 0 Then Exit Function

'------------------------------------------------------------------------------
' FORMAT
'------------------------------------------------------------------------------
    'Weeks are a single token; months split into years and residual months
        If Not CandIsMonth(Best) Then
            TokenOut = SignPrefix(Direction) & VBA.CStr(CandCount(Best)) & "W"
        Else
            TokenYears = CandCount(Best) \ 12
            TokenMonths = CandCount(Best) Mod 12
            If TokenYears = 0 Then
                TokenOut = SignPrefix(Direction) & VBA.CStr(TokenMonths) & "M"
            ElseIf TokenMonths = 0 Then
                TokenOut = SignPrefix(Direction) & VBA.CStr(TokenYears) & "Y"
            Else
                TokenOut = SignPrefix(Direction) & VBA.CStr(TokenYears) & "Y" & VBA.CStr(TokenMonths) & "M"
            End If
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Contract: TRUE only when TokenOut was assigned
        Condition = cdNone
        TryFormatPillar = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the failure path
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Fail:
    'Return FALSE per contract, leaving TokenOut untouched
        TryFormatPillar = False
        Condition = cdResultWindow
    'Do not leave stale error state for the caller to observe
        Err.Clear

End Function

Private Function AnchorPasses( _
    ByVal AnchorSerial As Double, _
    ByVal SerialEnd As Double, _
    ByVal Direction As Long) As Boolean
'
'------------------------------------------------------------------------------
'                  TEST WHETHER AN ANCHOR PASSES THE END DATE
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether an anchor lies strictly beyond the end date in the interval direction.
'
' WHY THIS EXISTS
'   Floor, ceiling and the month search all ask the same question of an anchor.
'
' INPUTS
'   AnchorSerial
'     Serial of the anchor.
'
'   SerialEnd
'     Serial of the end date.
'
'   Direction
'     +1 when the end date follows the start date, -1 when it precedes it.
'
' RETURNS
'   Boolean
'     True when the anchor is strictly beyond the end date in that direction.
'
' BEHAVIOR
'   Multiplies the signed distance by the direction and tests it for being positive.
'
' ERROR POLICY
'   None.
'
' DEPENDENCIES
'   None. CALC-09 steps 3 and 5
'
' NOTES
'   An anchor equal to the end date does not pass it
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Strictly beyond, in the signed direction
        AnchorPasses = (VBA.CDbl(Direction) * (AnchorSerial - SerialEnd) > 0#)

End Function

Private Function TieToMonth( _
    ByVal Candidate As Long, _
    ByVal Best As Long, _
    ByRef IsMonth() As Boolean, _
    ByRef Count() As Long) As Boolean
'
'------------------------------------------------------------------------------
'                          APPLY THE SHARED TIE RULE
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the shared tie rule: a month beats a week, and between months the larger count wins.
'
' WHY THIS EXISTS
'   The tie rule is identical under every rounding mode, so the three modes share it.
'
' INPUTS
'   Candidate
'     Index of the candidate that ties with the current best.
'
'   Best
'     Index of the current best candidate.
'
'   IsMonth()
'     Whether each candidate is a month anchor.
'
'   Count()
'     The week or month count of each candidate.
'
' RETURNS
'   Boolean
'     True when Candidate should displace Best.
'
' BEHAVIOR
'   A month displaces a week. A month displaces a month only with a larger count.
'   Otherwise the current best is kept.
'
' ERROR POLICY
'   None.
'
' DEPENDENCIES
'   None. CALC-09 step 6
'
' NOTES
'   The larger month count is the ceiling month anchor
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Month over week
        If IsMonth(Candidate) And Not IsMonth(Best) Then
            TieToMonth = True

    'Between months, the larger count (the ceiling anchor)
        ElseIf IsMonth(Candidate) And IsMonth(Best) Then
            TieToMonth = (Count(Candidate) > Count(Best))

    'Otherwise keep the current best
        Else
            TieToMonth = False
        End If

End Function

Private Function SignPrefix(ByVal Direction As Long) As String

'
'------------------------------------------------------------------------------
'                    RETURN THE SIGN PREFIX OF AN INTERVAL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the sign prefix of an interval: a minus for a negative interval, otherwise nothing.
'
' WHY THIS EXISTS
'   Every emitted token carries the same prefix, so it is decided once.
'
' INPUTS
'   Direction
'     +1 when the end date follows the start date, -1 when it precedes it.
'
' RETURNS
'   String
'     "-" or an empty string.
'
' BEHAVIOR
'   Tests the direction.
'
' ERROR POLICY
'   None.
'
' DEPENDENCIES
'   None. CALC-09 step 8
'
' NOTES
'   A plus sign is never emitted
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Only a negative interval carries a prefix; "+" is never emitted
        If Direction < 0 Then SignPrefix = "-" Else SignPrefix = vbNullString

End Function
