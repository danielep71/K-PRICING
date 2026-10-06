Attribute VB_Name = "KPR_Demo_Dates"
'==============================================================================
' MODULE: KPR_Demo_Dates
'------------------------------------------------------------------------------
' PURPOSE
'   Content of the Date Primitives demo sheet, built by KPR_Demo__Builder: the
'   22 KPR_Dates_* functions beside native-Excel reference formulas, driven
'   by editable inputs, plus fixed native-error cases and dynamic-array
'   spills.
'
' WHY THIS EXISTS
'   The sheet is rebuilt from this tracked source rather than stored in a
'   workbook, so it can be reviewed, versioned and checked by
'   KPR_Tests_RunDemo. Layout and style live in KPR_Demo__Builder; this module
'   only says what the sheet shows.
'
' SCOPE
'   - KPR_Demo_BuildDates       build the Date Primitives demo sheet
'
' REFERENCE FORMULAS
'   Each reference is plain Excel that reproduces the documented behaviour of
'   docs/DATE_LAYER_CONTRACT.md for any valid input, including Opt_KeepEOM
'   and the Monday/Sunday weekday base, so every row shows OK for any valid
'   input. Leap years use the Gregorian rule rather than DATE arithmetic,
'   because Excel's DATE counts a 29 February 1900 that never existed. The
'   two month-end policy rows compare against EDATE and EOMONTH on purpose:
'   they show DIFFERS when Keep EOM changes the result.
'
' DEFAULT INPUTS
'   2024-04-30 is a month end, so the default sheet shows one DIFFERS row
'   (AddMonths against EDATE with Keep EOM) and no FAIL. The fixed cases were
'   computed by the independent reference model in tools/gen_fixtures.py.
'
' VISIBILITY
'   Option Private Module. Callable from this project, the Immediate window
'   and Application.Run.
'
' ALLOWED DEPENDENCIES
'   KPR_Demo__Builder. KPR_Dates_* names appear only inside formula text.
'
' UPDATED
'   2026-10-06
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit         'Force explicit variable declarations
    Option Private Module   'Demo infrastructure: invisible outside this project

'------------------------------------------------------------------------------
' CONSTANTS
'------------------------------------------------------------------------------
    Private Const DEMO_SHEET As String = "KPR Dates Demo"

    'Gregorian leap-year test of the year of DateIn
        Private Const LEAP_DATEIN As String = _
            "OR(MOD(YEAR(DateIn),400)=0,AND(MOD(YEAR(DateIn),4)=0,MOD(YEAR(DateIn),100)<>0))"

    'First day of the month of DateIn, and the weekday-number mode of WEEKDAY
        Private Const FIRST_OF_MONTH As String = "DATE(YEAR(DateIn),MONTH(DateIn),1)"
        Private Const WEEKDAY_MODE   As String = "IF(WeekBaseMonday,2,1)"

    'Keep EOM applies: the flag is on and DateIn is a month end
        Private Const KEEP_EOM_APPLIES As String = "AND(KeepEOM,DateIn=EOMONTH(DateIn,0))"

'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Function KPR_Demo_BuildDates(Optional ByVal OutputPath As String = "") As Boolean
'
'==============================================================================
'                             KPR_Demo_BuildDates
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the Date Primitives demo sheet.
'
' INPUTS
'   OutputPath
'     Optional. "" adds the sheet to the active workbook, or to a new
'     workbook when the active one cannot take it. A path builds into a new
'     .xlsx file there, which must not exist yet.
'
' RETURNS
'   TRUE when the sheet was built. FALSE otherwise; KPR_Demo_LastReport says
'   why, and nothing the build created is left behind.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
    Dim NthRef  As String   'Reference date of the n-th weekday, before the month check
    Dim LastRef As String   'Reference date of the last weekday

    KPR_Demo_BuildDates = False
    If Not Demo_Begin(DEMO_SHEET, "DATE PRIMITIVES", _
                      "Calendar-day date functions beside native Excel. Edit the orange cells.", _
                      OutputPath) Then Exit Function
    On Error GoTo Fail

'------------------------------------------------------------------------------
' INPUTS
'------------------------------------------------------------------------------
    Demo_Inputs "Inputs (edit these)"
        Demo_Input "DateIn", "DateIn", DateSerial(2024, 4, 30), "Any date from 1900-03-01 to 9999-12-31"
        Demo_Input "nDays", "nDays", 2, "Whole number of days"
        Demo_Input "nWeeks", "nWeeks", 3, "Whole number of weeks"
        Demo_Input "nMonths", "nMonths", 1, "Whole number of months"
        Demo_Input "nYears", "nYears", 2, "Whole number of years"
        Demo_Input "Keep EOM", "KeepEOM", True, "TRUE keeps a month end at month end"
        Demo_Input "Weekday index", "WdIndex", 4, "1 to 7 in the weekday base below"
        Demo_Input "Occurrence n", "nOccur", 2, "1 to 5"
        Demo_Input "WeekBase Monday", "WeekBaseMonday", True, "TRUE: Monday = 1; FALSE: Sunday = 1"

'------------------------------------------------------------------------------
' DAY PRIMITIVES
'------------------------------------------------------------------------------
    Demo_Section "Day primitives"
        Demo_Compare "=KPR_Dates_DayOfWeek(DateIn,WeekBaseMonday)", _
                     "=WEEKDAY(DateIn," & WEEKDAY_MODE & ")", _
                     "Monday = 1 by default; WeekBaseMonday = FALSE matches WEEKDAY's Sunday = 1."
        Demo_Compare "=KPR_Dates_DaysInMonth(DateIn)", _
                     "=DAY(EOMONTH(DateIn,0))", _
                     "Gregorian month length."
        Demo_Compare "=KPR_Dates_DaysInYear(YEAR(DateIn))", _
                     "=IF(" & LEAP_DATEIN & ",366,365)", _
                     "Takes a year, not a date."

'------------------------------------------------------------------------------
' PERIOD BOUNDARIES
'------------------------------------------------------------------------------
    Demo_Section "Period boundaries"
        Demo_Compare "=KPR_Dates_BeginOfMonth(DateIn)", "=EOMONTH(DateIn,-1)+1", _
                     "First day of the month.", DEMO_DATE
        Demo_Compare "=KPR_Dates_EndOfMonth(DateIn)", "=EOMONTH(DateIn,0)", _
                     "Last day of the month.", DEMO_DATE
        Demo_Compare "=KPR_Dates_BeginOfQuarter(DateIn)", _
                     "=DATE(YEAR(DateIn),3*INT((MONTH(DateIn)-1)/3)+1,1)", _
                     "Calendar quarters start in January, April, July and October.", DEMO_DATE
        Demo_Compare "=KPR_Dates_EndOfQuarter(DateIn)", _
                     "=EOMONTH(DATE(YEAR(DateIn),3*INT((MONTH(DateIn)-1)/3)+3,1),0)", _
                     "Last day of the calendar quarter.", DEMO_DATE
        Demo_Compare "=KPR_Dates_BeginOfYear(DateIn)", "=DATE(YEAR(DateIn),1,1)", _
                     "1 January.", DEMO_DATE
        Demo_Compare "=KPR_Dates_EndOfYear(DateIn)", "=DATE(YEAR(DateIn),12,31)", _
                     "31 December.", DEMO_DATE

'------------------------------------------------------------------------------
' PREDICATES
'------------------------------------------------------------------------------
    Demo_Section "Predicates"
        Demo_Compare "=KPR_Dates_IsMonthEnd(DateIn)", "=DateIn=EOMONTH(DateIn,0)", _
                     "TRUE on the last day of a month."
        Demo_Compare "=KPR_Dates_IsQuarterEnd(DateIn)", _
                     "=AND(DateIn=EOMONTH(DateIn,0),MOD(MONTH(DateIn),3)=0)", _
                     "TRUE on 31 March, 30 June, 30 September and 31 December."
        Demo_Compare "=KPR_Dates_IsYearEnd(DateIn)", "=DateIn=DATE(YEAR(DateIn),12,31)", _
                     "TRUE on 31 December."
        Demo_Compare "=KPR_Dates_IsLeapYear(YEAR(DateIn))", "=" & LEAP_DATEIN, _
                     "Takes a year. 1900 is not a leap year, whatever Excel's DATE says."

'------------------------------------------------------------------------------
' DATE ARITHMETIC
'------------------------------------------------------------------------------
    Demo_Section "Date arithmetic"
        Demo_Compare "=KPR_Dates_AddDays(DateIn,nDays)", "=DateIn+nDays", _
                     "Exact calendar days.", DEMO_DATE
        Demo_Compare "=KPR_Dates_AddWeeks(DateIn,nWeeks)", "=DateIn+7*nWeeks", _
                     "Weeks of seven days.", DEMO_DATE
        Demo_Compare "=KPR_Dates_AddMonths(DateIn,nMonths,KeepEOM)", _
                     "=IF(" & KEEP_EOM_APPLIES & ",EOMONTH(DateIn,nMonths),EDATE(DateIn,nMonths))", _
                     "Clips to a shorter month end; Keep EOM moves a month end to the target month end.", _
                     DEMO_DATE
        Demo_Compare "=KPR_Dates_AddYears(DateIn,nYears,KeepEOM)", _
                     "=IF(" & KEEP_EOM_APPLIES & ",EOMONTH(DateIn,12*nYears),EDATE(DateIn,12*nYears))", _
                     "Twelve months per year: 29 February clips to 28 February.", DEMO_DATE

'------------------------------------------------------------------------------
' MONTH-END POLICY
'------------------------------------------------------------------------------
    Demo_Section "Month-end policy against EDATE and EOMONTH"
        Demo_Compare "=KPR_Dates_AddMonths(DateIn,nMonths,KeepEOM)", "=EDATE(DateIn,nMonths)", _
                     "EDATE always clips. DIFFERS when Keep EOM moves a month end further.", _
                     DEMO_DATE, KEEP_EOM_APPLIES
        Demo_Compare "=KPR_Dates_AddMonths(DateIn,nMonths,KeepEOM)", "=EOMONTH(DateIn,nMonths)", _
                     "EOMONTH always lands on a month end. Equal only when Keep EOM applies.", _
                     DEMO_DATE, "NOT(" & KEEP_EOM_APPLIES & ")"

'------------------------------------------------------------------------------
' WEEKDAY LOCATORS
'------------------------------------------------------------------------------
    NthRef = FIRST_OF_MONTH & "+MOD(WdIndex-WEEKDAY(" & FIRST_OF_MONTH & "," & WEEKDAY_MODE & "),7)+7*(nOccur-1)"
    LastRef = "EOMONTH(DateIn,0)-MOD(WEEKDAY(EOMONTH(DateIn,0)," & WEEKDAY_MODE & ")-WdIndex,7)"
    Demo_Section "Weekday locators"
        Demo_Compare "=KPR_Dates_NthWeekdayOfMonth(YEAR(DateIn),MONTH(DateIn),WdIndex,nOccur,WeekBaseMonday)", _
                     "=IF(MONTH(" & NthRef & ")=MONTH(DateIn)," & NthRef & ",SQRT(-1))", _
                     "The n-th weekday of the month of DateIn; #NUM! when the month has no such occurrence.", _
                     DEMO_DATE
        Demo_Compare "=KPR_Dates_LastWeekdayOfMonth(YEAR(DateIn),MONTH(DateIn),WdIndex,WeekBaseMonday)", _
                     "=" & LastRef, _
                     "The last such weekday of the month of DateIn.", DEMO_DATE

'------------------------------------------------------------------------------
' PILLARS
'------------------------------------------------------------------------------
    Demo_Section "Pillars"
        Demo_Compare "=KPR_Dates_DateFromPillar(DateIn,""ON"")", "=DateIn+1", _
                     "Overnight: one calendar day.", DEMO_DATE
        Demo_Compare "=KPR_Dates_DateFromPillar(DateIn,""TN"")", "=DateIn+2", _
                     "Tom-next: two calendar days. Business days arrive with calendars.", DEMO_DATE
        Demo_Compare "=KPR_Dates_DateFromPillar(DateIn,""1W"")", "=DateIn+7", _
                     "One week.", DEMO_DATE
        Demo_Compare "=KPR_Dates_DateFromPillar(DateIn,""3M"")", "=EDATE(DateIn,3)", _
                     "Months clip like EDATE.", DEMO_DATE
        Demo_Compare "=KPR_Dates_DateFromPillar(DateIn,""1Y6M"")", "=EDATE(DateIn,18)", _
                     "Years and months combine.", DEMO_DATE
        Demo_Compare "=KPR_Dates_DateFromPillar(DateIn,""-1D"")", "=DateIn-1", _
                     "A sign applies to the whole token.", DEMO_DATE
        Demo_Compare "=KPR_Dates_PillarFromDates(DateIn,DateIn+7)", "=""1W""", _
                     "Seven days format as one week."
        Demo_Compare "=KPR_Dates_PillarFromDates(DateIn,EDATE(DateIn,3))", "=""3M""", _
                     "An exact month count formats as months."
        Demo_Compare "=KPR_Dates_PillarFromDates(DateIn,EDATE(DateIn,18))", "=""1Y6M""", _
                     "Twelve months or more format as years and months."
        Demo_Compare "=KPR_Dates_PillarFromDates(DateIn,DateIn-7)", "=""-1W""", _
                     "An end before the start gives a minus sign."
        Demo_Compare "=KPR_Dates_PillarFromDates(""2024-01-15"",""2024-02-10"",""NEAREST"")", "=""1M""", _
                     "Rounding NEAREST: 26 days are nearer one month than three weeks."
        Demo_Compare "=KPR_Dates_PillarFromDates(""2024-01-15"",""2024-02-10"",""FLOOR"")", "=""3W""", _
                     "Rounding FLOOR: the largest pillar not after the end."
        Demo_Compare "=KPR_Dates_PillarFromDates(""2024-01-15"",""2024-02-10"",""CEILING"")", "=""1M""", _
                     "Rounding CEILING: the smallest pillar not before the end."

    Demo_Note "Dates are passed as cells or as YYYY-MM-DD text, so nothing here depends on the regional date format."
    Demo_Note "In a 1904-date-system workbook every value-taking function returns #N/A instead of a shifted date."

'------------------------------------------------------------------------------
' NATIVE ERRORS
'------------------------------------------------------------------------------
    Demo_Cases "Invalid input gives a native Excel error"
        Demo_Case "aa", "=KPR_Dates_DaysInMonth({IN})", CVErr(xlErrValue), _
                  "DaysInMonth: text that is not YYYY-MM-DD gives #VALUE!."
        Demo_Case "32/04/2020", "=KPR_Dates_BeginOfMonth({IN})", CVErr(xlErrValue), _
                  "BeginOfMonth: regional date text is never guessed."
        Demo_Case """2024-04-30""", "=KPR_Dates_IsMonthEnd({IN})", CVErr(xlErrValue), _
                  "IsMonthEnd: quotes inside the text make it invalid."
        Demo_Case True, "=KPR_Dates_DaysInYear({IN})", CVErr(xlErrValue), _
                  "DaysInYear: TRUE is not a year."
        Demo_Case 1899, "=KPR_Dates_IsLeapYear({IN})", CVErr(xlErrValue), _
                  "IsLeapYear: years run from 1900 to 9999."
        Demo_Case -4, "=KPR_Dates_DayOfWeek({IN})", CVErr(xlErrNum), _
                  "DayOfWeek: a serial before 1900-03-01 gives #NUM!."
        Demo_Case 0, "=KPR_Dates_EndOfMonth({IN})", CVErr(xlErrNum), _
                  "EndOfMonth: serial 0 is outside the supported window."
        Demo_Case 32, "=KPR_Dates_IsQuarterEnd({IN})", CVErr(xlErrNum), _
                  "IsQuarterEnd: 1900-02-01 is before the window."
        Demo_Case "9999-12-31", "=KPR_Dates_AddDays({IN},1)", CVErr(xlErrNum), _
                  "AddDays: a result after 9999-12-31 gives #NUM!."
        Demo_Case 1.5, "=KPR_Dates_AddDays(""2024-01-31"",{IN})", CVErr(xlErrValue), _
                  "AddDays: a count is never truncated or rounded."
        Demo_Case "yes", "=KPR_Dates_AddMonths(""2024-01-31"",1,{IN})", CVErr(xlErrValue), _
                  "AddMonths: Keep EOM takes TRUE or FALSE only."
        Demo_Case 5, "=KPR_Dates_NthWeekdayOfMonth(2026,2,1,{IN})", CVErr(xlErrNum), _
                  "NthWeekdayOfMonth: February 2026 has four Mondays."
        Demo_Case "-ON", "=KPR_Dates_DateFromPillar(""2024-01-31"",{IN})", CVErr(xlErrValue), _
                  "DateFromPillar: an alias never carries a sign."
        Demo_Case "1M2M", "=KPR_Dates_DateFromPillar(""2024-01-31"",{IN})", CVErr(xlErrValue), _
                  "DateFromPillar: a unit may appear once."
        Demo_Case "ROUND", "=KPR_Dates_PillarFromDates(""2024-01-15"",""2024-04-15"",{IN})", CVErr(xlErrValue), _
                  "PillarFromDates: rounding is NEAREST, FLOOR or CEILING."
        Demo_Case CVErr(xlErrDiv0), "=KPR_Dates_EndOfMonth({IN})", CVErr(xlErrDiv0), _
                  "EndOfMonth: an error in the input passes through unchanged."
        Demo_Case Empty, "=KPR_Dates_HostDateSystem()", 1900, _
                  "HostDateSystem: this workbook uses the 1900 date system."

'------------------------------------------------------------------------------
' DYNAMIC ARRAYS
'------------------------------------------------------------------------------
    'EOMONTH refuses a spill or range reference as an array; +0 turns it into values
    Demo_Spill "One formula, many dates", "Date", "=SEQUENCE(30,1,DateIn,25)", 30, DEMO_DATE
        Demo_SpillColumn "Days in month", "=KPR_Dates_DaysInMonth({SRC})", "DAY(EOMONTH({SRC}+0,0))"
        Demo_SpillColumn "Begin of month", "=KPR_Dates_BeginOfMonth({SRC})", "EOMONTH({SRC}+0,-1)+1", DEMO_DATE
        Demo_SpillColumn "Is leap year", "=KPR_Dates_IsLeapYear(YEAR({SRC}))", _
                         "((MOD(YEAR({SRC}),400)=0)+(MOD(YEAR({SRC}),4)=0)*(MOD(YEAR({SRC}),100)<>0)>0)"
    If Demo_HasDynamicArrays() Then
        Demo_Cases "Array limits"
            Demo_Case "100,000 dates", "=ROWS(KPR_Dates_AddDays(SEQUENCE(100000,1,45000),0))", 100000, _
                      "Up to 100,000 elements per call are accepted."
            Demo_Case "100,001 dates", "=KPR_Dates_AddDays(SEQUENCE(100001,1,45000),0)", CVErr(xlErrNum), _
                      "One element more gives a single #NUM!."
            Demo_Case "2x1 and 1x3", "=KPR_Dates_AddDays({""2024-01-01"";""2024-01-02""},{1,2,3})", _
                      CVErr(xlErrValue), "Array arguments of different shapes give #VALUE!."
    End If

'------------------------------------------------------------------------------
' FINISH
'------------------------------------------------------------------------------
    On Error GoTo 0
    KPR_Demo_BuildDates = Demo_Finish()
    Exit Function

Fail:
    KPR_Demo_BuildDates = Demo_Abort(Err.Number, Err.Description)

End Function
