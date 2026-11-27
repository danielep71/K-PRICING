Attribute VB_Name = "TEST_Dates"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: TEST_Dates
'------------------------------------------------------------------------------
' PURPOSE
'   Tests of the KPR_Dates facade: one call of each of the 22 functions, every condition it can
'   meet from VBA, the array rules and the optional controls
'
' WHY THIS EXISTS
'   The facade joins the host rule, the shape engine, the parsers and the date core. These cases
'   check the joined behaviour through the supported entry points, as a worksheet calls them
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   RunDateTests runs the suite. Option Private Module keeps it out of the workbook's external
'   surface. The cases call the worksheet functions from VBA, so no worksheet host is identified
'   and the 1900 serial contract applies; they change no Excel state. CertifyDates certifies
'   its normal run, as .github/excel-evidence-policy.json lists it
'
' ERROR POLICY
'   As TEST_Harness.EndRun: a run that does not pass raises TEST_Harness.ERR_TEST_FAILED,
'   or the unexpected error that escaped the runner. A refused start raises
'   TEST_Harness.ERR_TEST_DIRTY_START
'
' DEPENDENCIES
'   KPR_Dates, TEST_Harness
'
' NOTES
'   In a development workbook holding the modules of src/ and tests/modules/, type
'   TEST_Dates.RunDateTests in the Immediate window. Expected serials come from
'   tests/derivations/date_core_reference.py; error values from the data contract's Conditions
'   table. A worksheet caller, a 1904 workbook and volatility need a worksheet, and are left to
'   the Excel run
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'A complete run starts every case and evaluates every assertion once
    Private Const SUITE_NAME            As String = "dates"             'Suite name in the run log
    Private Const EXPECTED_CASES        As Long = 4                     'Cases a complete run starts
    Private Const EXPECTED_ASSERTIONS   As Long = 85                    'Assertions a complete run evaluates

    'The native error numbers that the Conditions table lists, and one an error passed in may hold
    Private Const ERROR_VALUE           As Long = 2015                  '#VALUE!
    Private Const ERROR_NUM             As Long = 2036                  '#NUM!
    Private Const ERROR_NA              As Long = 2042                  '#N/A
    Private Const ERROR_DIV0            As Long = 2007                  '#DIV/0!


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub RunDateTests()

'
'------------------------------------------------------------------------------
'                               RUN DATES TESTS
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case in a fixed order between BeginRun and EndRun. A passing run ends with
'   RESULT=PASS; suite=dates; operation=PASS; completeness=COMPLETE; cases=4;
'   assertions=85; failures=0; cleanup=PASS; recovery=NONE
'
' WHY THIS EXISTS
'   One entry point runs the suite and reports it through the harness, so that a run is
'   complete or visibly incomplete
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Starts the run, calls each case, then passes any error that escaped them to EndRun
'
' ERROR POLICY
'   An error escaping a case handler is passed to EndRun, which reports it and re-raises it
'   unchanged. BeginRun's refusal propagates directly
'
' DEPENDENCIES
'   TEST_Harness.BeginRun, TEST_Harness.EndRun
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Dates.RunDateTests"

    Dim ErrNumber       As Long     'Runner error number; 0 for none
    Dim ErrSource       As String   'Runner error source
    Dim ErrDescription  As String   'Runner error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the run; a refused start propagates directly
        TEST_Harness.BeginRun SUITE_NAME, "NORMAL", EXPECTED_CASES, EXPECTED_ASSERTIONS
    'Send an error that escapes a case to the runner handler
        On Error GoTo RunFailed

'------------------------------------------------------------------------------
' RUN CASES
'------------------------------------------------------------------------------
    'Run every case in a fixed order
        CaseScalarResults
        CaseConditions
        CaseArrays
        CaseControls

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
Finish:
    'Restore normal error handling
        On Error GoTo 0
    'Report the run, raising its failure or the escaped error
        TEST_Harness.EndRun ErrNumber, ErrSource, ErrDescription
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
RunFailed:
    'Capture the escaped error
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
    'Finish the run with the captured error
        Resume Finish

End Sub

Public Sub CertifyDates( _
    ByVal Folder As String, _
    ByVal Fixtures As String)
'
'------------------------------------------------------------------------------
'                              CERTIFY THE SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Certifies the suite's normal run and saves its evidence in the folder.
'
' WHY THIS EXISTS
'   The certification entry point that .github/excel-evidence-policy.json names.
'
' INPUTS
'   Folder
'     The evidence folder.
'
'   Fixtures
'     The tests/fixtures/ folder, which this suite does not read.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs RunDateTests, saves its log, certifies the run, then saves the certification log.
'
' ERROR POLICY
'   TEST_Harness.EndCertification raises when the certification did not pass.
'
' DEPENDENCIES
'   RunDateTests, TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Dates.CertifyDates"

    Dim RunErrorNumber          As Long     'Error that ended the run; 0 for none
    Dim RunErrorDescription     As String   'Its description

'------------------------------------------------------------------------------
' CERTIFY
'------------------------------------------------------------------------------
    'Refuse state left by an earlier run and start the certification
        TEST_Harness.BeginCertification SUITE_NAME
    'Run the suite and capture the error that ended it
        On Error Resume Next
        RunDateTests
        RunErrorNumber = Err.Number
        RunErrorDescription = Err.Description
    'Restore normal error handling
        On Error GoTo 0
    'Save the log and certify the run
        TEST_Harness.SaveLog Folder & "\dates.log"
        TEST_Harness.CertifyRun "NORMAL", True, RunErrorNumber, RunErrorDescription
    'Save the certification log, raising when the certification did not pass
        TEST_Harness.EndCertification Folder & "\dates-certification.log"

End Sub

'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseScalarResults()

'
'------------------------------------------------------------------------------
'                       CHECK ONE CALL OF EACH FUNCTION
'------------------------------------------------------------------------------
' PURPOSE
'   Each of the 22 functions returns its documented result for a nominal call. 28 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the array rules hold as the calculation contracts state them
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls every function once with scalar arguments, both weekday bases and every rounding mode
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   KPR_Dates, TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Dates.CaseScalarResults"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "scalar-results"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'One nominal call of each function, with a date, a serial or ISO text
        TEST_Harness.AssertEqualDouble "day-of-week", 3, KPR_Dates_DayOfWeek(46351#)
        TEST_Harness.AssertEqualDouble "day-of-week-sunday", 4, KPR_Dates_DayOfWeek(46351#, False)
        TEST_Harness.AssertEqualDouble "days-in-month", 28, KPR_Dates_DaysInMonth("2026-02-10")
        TEST_Harness.AssertEqualDouble "days-in-year", 366, KPR_Dates_DaysInYear(2024)
        TEST_Harness.AssertEqualDouble "begin-of-month", 46357, KPR_Dates_BeginOfMonth("2026-12-31")
        TEST_Harness.AssertEqualDouble "end-of-month", 46053, KPR_Dates_EndOfMonth(CDate(46023))
        TEST_Harness.AssertEqualDouble "begin-of-quarter", 46296, KPR_Dates_BeginOfQuarter(46351#)
        TEST_Harness.AssertEqualDouble "end-of-quarter", 46387, KPR_Dates_EndOfQuarter(46351#)
        TEST_Harness.AssertEqualDouble "begin-of-year", 46023, KPR_Dates_BeginOfYear(46351#)
        TEST_Harness.AssertEqualDouble "end-of-year", 46387, KPR_Dates_EndOfYear(46023#)
        TEST_Harness.AssertTrue "is-month-end", KPR_Dates_IsMonthEnd(46053#) = True, "the result is not True"
        TEST_Harness.AssertTrue "is-quarter-end", KPR_Dates_IsQuarterEnd("2026-03-31") = True, "the result is not True"
        TEST_Harness.AssertTrue "is-year-end", KPR_Dates_IsYearEnd(46386#) = False, "the result is not False"
        TEST_Harness.AssertTrue "is-leap-year", KPR_Dates_IsLeapYear(2100) = False, "the result is not False"
        TEST_Harness.AssertEqualDouble "add-days", 46388, KPR_Dates_AddDays(46387#, 1)
        TEST_Harness.AssertEqualDouble "add-weeks", 46380, KPR_Dates_AddWeeks(46387#, -1)
        TEST_Harness.AssertEqualDouble "add-months-clip", 46081, KPR_Dates_AddMonths(46053#, 1)
        TEST_Harness.AssertEqualDouble "add-months-keep-end", 46112, KPR_Dates_AddMonths(46081#, 1, True)
        TEST_Harness.AssertEqualDouble "add-years-leap-day", 45716, KPR_Dates_AddYears("2024-02-29", 1)
        TEST_Harness.AssertEqualDouble "nth-weekday", 46352, KPR_Dates_NthWeekdayOfMonth(2026, 11, 4, 4)
        TEST_Harness.AssertEqualDouble "last-weekday", 46167, KPR_Dates_LastWeekdayOfMonth(2026, 5, 1)
        TEST_Harness.AssertEqualText "pillar-from-dates", "3M", KPR_Dates_PillarFromDates(46023#, 46113#)
        TEST_Harness.AssertEqualText "pillar-from-dates-nearest", "3M", KPR_Dates_PillarFromDates(46023#, 46123#)
        TEST_Harness.AssertEqualText "pillar-from-dates-floor", "3M", KPR_Dates_PillarFromDates(46023#, 46123#, "floor")
        TEST_Harness.AssertEqualText "pillar-from-dates-ceiling", "4M", KPR_Dates_PillarFromDates(46023#, 46123#, " CEILING ")
        TEST_Harness.AssertEqualDouble "date-from-pillar", 46081, KPR_Dates_DateFromPillar(46053#, "1M")
        TEST_Harness.AssertEqualDouble "date-from-pillar-alias", 46055, KPR_Dates_DateFromPillar(46053#, "T/N")
    'Without a worksheet caller the 1900 serial contract applies
        TEST_Harness.AssertEqualDouble "host-date-system", 1900, KPR_Dates_HostDateSystem()

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the unexpected error against the case
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseConditions()

'
'------------------------------------------------------------------------------
'                            CHECK EVERY CONDITION
'------------------------------------------------------------------------------
' PURPOSE
'   Each condition the facade can meet returns the Excel error of the data contract. 32 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the array rules hold as the calculation contracts state them
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Breaks each argument rule once and checks the error value, an error passed in included
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   KPR_Dates, TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Dates.CaseConditions"

    Dim Item              As Collection   'An object that is not a Range
    Dim Large()           As Variant      'A row of 100,001 elements

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "conditions"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Date values
        TEST_Harness.AssertErrorValue "date-text-format", ERROR_VALUE, KPR_Dates_EndOfMonth("2026-1-5")
        TEST_Harness.AssertErrorValue "date-text-locale", ERROR_VALUE, KPR_Dates_EndOfMonth("31/12/2026")
        TEST_Harness.AssertErrorValue "date-text-numeric", ERROR_VALUE, KPR_Dates_EndOfMonth("45292")
        TEST_Harness.AssertErrorValue "date-text-impossible", ERROR_VALUE, KPR_Dates_EndOfMonth("1900-02-29")
        TEST_Harness.AssertErrorValue "date-type-rejected", ERROR_VALUE, KPR_Dates_EndOfMonth(True)
        TEST_Harness.AssertErrorValue "date-window", ERROR_NUM, KPR_Dates_EndOfMonth(60)
        TEST_Harness.AssertErrorValue "input-blank-required", ERROR_VALUE, KPR_Dates_EndOfMonth(Empty)
    'An error passed in is returned unchanged
        TEST_Harness.AssertErrorValue "input-error-propagated-na", ERROR_NA, KPR_Dates_EndOfMonth(CVErr(ERROR_NA))
        TEST_Harness.AssertErrorValue "input-error-propagated-div0", ERROR_DIV0, KPR_Dates_AddDays(46387#, CVErr(ERROR_DIV0))
    'Whole numbers and their domains
        TEST_Harness.AssertErrorValue "integer-fraction", ERROR_VALUE, KPR_Dates_AddDays(46387#, 2.5)
        TEST_Harness.AssertErrorValue "integer-range", ERROR_NUM, KPR_Dates_AddDays(46387#, 10000000000#)
        TEST_Harness.AssertErrorValue "integer-type-rejected", ERROR_VALUE, KPR_Dates_AddDays(46387#, "7")
        TEST_Harness.AssertErrorValue "domain-year", ERROR_VALUE, KPR_Dates_IsLeapYear(1899)
        TEST_Harness.AssertErrorValue "domain-month", ERROR_VALUE, KPR_Dates_LastWeekdayOfMonth(2026, 13, 1)
        TEST_Harness.AssertErrorValue "domain-weekday", ERROR_VALUE, KPR_Dates_LastWeekdayOfMonth(2026, 1, 8)
        TEST_Harness.AssertErrorValue "domain-occurrence", ERROR_VALUE, KPR_Dates_NthWeekdayOfMonth(2026, 1, 1, 6)
        TEST_Harness.AssertErrorValue "occurrence-absent", ERROR_NUM, KPR_Dates_NthWeekdayOfMonth(2026, 2, 1, 5)
        TEST_Harness.AssertErrorValue "result-window", ERROR_NUM, KPR_Dates_AddDays(2958465#, 1)
        TEST_Harness.AssertErrorValue "result-window-begin-of-year", ERROR_NUM, KPR_Dates_BeginOfYear(61)
    'Pillar tokens
        TEST_Harness.AssertErrorValue "pillar-type-rejected", ERROR_VALUE, KPR_Dates_DateFromPillar(46387#, 3)
        TEST_Harness.AssertErrorValue "pillar-token-malformed", ERROR_VALUE, KPR_Dates_DateFromPillar(46387#, "1 M")
        TEST_Harness.AssertErrorValue "pillar-duplicate-unit", ERROR_VALUE, KPR_Dates_DateFromPillar(46387#, "1M2M")
        TEST_Harness.AssertErrorValue "pillar-alias-signed", ERROR_VALUE, KPR_Dates_DateFromPillar(46387#, "-ON")
        TEST_Harness.AssertErrorValue "pillar-aggregate-range", ERROR_NUM, KPR_Dates_DateFromPillar(46387#, "999999999999M")
    'Controls, for the whole call
        TEST_Harness.AssertErrorValue "control-type-rejected", ERROR_VALUE, KPR_Dates_DayOfWeek(46351#, 0)
        TEST_Harness.AssertErrorValue "control-token-unknown", ERROR_VALUE, KPR_Dates_PillarFromDates(46023#, 46113#, "ROUND")
        TEST_Harness.AssertErrorValue "control-not-scalar", _
            ERROR_VALUE, _
            KPR_Dates_DayOfWeek(46351#, Array(True, False))
        TEST_Harness.AssertErrorValue "control-error-propagated", ERROR_NA, KPR_Dates_DayOfWeek(46351#, CVErr(ERROR_NA))
    'Shapes and capacity
        Set Item = New Collection
        TEST_Harness.AssertErrorValue "shape-unsupported", ERROR_VALUE, KPR_Dates_EndOfMonth(Item)
        Set Item = Nothing
        TEST_Harness.AssertErrorValue "shape-mismatch", _
            ERROR_VALUE, _
            KPR_Dates_AddDays(Array(46387, 46388), Array(1, 2, 3))
        ReDim Large(1 To 100001)
        TEST_Harness.AssertErrorValue "capacity-exceeded", ERROR_NUM, KPR_Dates_EndOfMonth(Large)
    'CEILING with no candidate in the window
        TEST_Harness.AssertErrorValue "ceiling-without-candidate", ERROR_NUM, KPR_Dates_PillarFromDates("9999-12-20", "9999-12-31", "CEILING")

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the unexpected error against the case
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseArrays()

'
'------------------------------------------------------------------------------
'                            CHECK THE ARRAY RULES
'------------------------------------------------------------------------------
' PURPOSE
'   Shapes are kept, scalars repeat, failures stay at their position, and the first failing argument
'   decides. 15 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the array rules hold as the calculation contracts state them
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Evaluates rows, a column, pairs of arrays and mismatched shapes, and checks errors element by element
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   KPR_Dates, TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Dates.CaseArrays"

    Dim Result            As Variant      'Result of a call
    Dim Column()          As Variant      'A 2 by 1 array

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "arrays"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Each position is evaluated on its own, by rows, and keeps its error
        Result = KPR_Dates_AddDays(Array(46387, "bad", 2958465), 1)
        TEST_Harness.AssertTrue "row-shape", _
            (LBound(Result, 1) = 1) And (UBound(Result, 1) = 1) And (UBound(Result, 2) = 3), _
            "the result is not 1 by 3"
        TEST_Harness.AssertEqualDouble "row-first", 46388, Result(1, 1)
        TEST_Harness.AssertErrorValue "row-second", ERROR_VALUE, Result(1, 2)
        TEST_Harness.AssertErrorValue "row-third", ERROR_NUM, Result(1, 3)
    'A scalar repeats over a column
        ReDim Column(1 To 2, 1 To 1)
        Column(1, 1) = 46053
        Column(2, 1) = 46081
        Result = KPR_Dates_AddMonths(Column, 1, True)
        TEST_Harness.AssertTrue "column-shape", (UBound(Result, 1) = 2) And (UBound(Result, 2) = 1), "the result is not 2 by 1"
        TEST_Harness.AssertEqualDouble "column-first", 46081, Result(1, 1)
        TEST_Harness.AssertEqualDouble "column-second", 46112, Result(2, 1)
    'Two arrays of one shape pair element by element
        Result = KPR_Dates_AddDays(Array(46387, 46388), Array(1, -1))
        TEST_Harness.AssertEqualDouble "pair-first", 46388, Result(1, 1)
        TEST_Harness.AssertEqualDouble "pair-second", 46387, Result(1, 2)
    'An all-scalar call, and a one-element array, return a scalar
        TEST_Harness.AssertTrue "scalar-call", Not IsArray(KPR_Dates_AddDays(46387#, 1)), "a scalar call returned an array"
        TEST_Harness.AssertTrue "one-element-call", Not IsArray(KPR_Dates_AddDays(Array(46387), 1)), "a one-element array returned an array"
    'A row and a column do not form an outer product
        TEST_Harness.AssertErrorValue "row-and-column", _
            ERROR_VALUE, _
            KPR_Dates_AddDays(Array(46387, 46388), Column)
    'At one position, the first failing argument in signature order decides
        TEST_Harness.AssertErrorValue "first-failing-date", ERROR_VALUE, KPR_Dates_AddDays("bad", "x")
        TEST_Harness.AssertErrorValue "first-failing-error", ERROR_NA, KPR_Dates_AddDays(CVErr(ERROR_NA), "x")
        TEST_Harness.AssertErrorValue "first-failing-year", ERROR_VALUE, KPR_Dates_NthWeekdayOfMonth(1899, 13, 1, 1)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the unexpected error against the case
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseControls()

'
'------------------------------------------------------------------------------
'                         CHECK THE OPTIONAL CONTROLS
'------------------------------------------------------------------------------
' PURPOSE
'   Omitted, Empty and Missing controls select the default; one element is a scalar. 10 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the array rules hold as the calculation contracts state them
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls with omitted, Empty, forwarded Missing and one-element controls
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   KPR_Dates, TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Dates.CaseControls"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "controls"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Omitted, Empty and a forwarded Missing select the default
        TEST_Harness.AssertEqualDouble "omitted", 3, KPR_Dates_DayOfWeek(46351#)
        TEST_Harness.AssertEqualDouble "empty", 3, KPR_Dates_DayOfWeek(46351#, Empty)
        TEST_Harness.AssertEqualDouble "forwarded-missing", 3, ForwardDayOfWeek(46351#)
        TEST_Harness.AssertEqualDouble "forwarded-missing-keep-eom", 45380, ForwardAddMonths(45351#, 1)
        TEST_Harness.AssertEqualText "forwarded-missing-rounding-below-half", "1W", ForwardPillarFromDates(46023#, 46033#)
        TEST_Harness.AssertEqualText "forwarded-missing-rounding-above-half", "2W", ForwardPillarFromDates(46023#, 46034#)
    'A one-element control is a scalar
        TEST_Harness.AssertEqualDouble "one-element", 4, KPR_Dates_DayOfWeek(46351#, Array(False))
    'Rounding text is trimmed of spaces and tabs and folded in ASCII
        TEST_Harness.AssertEqualText "rounding-empty", "3M", KPR_Dates_PillarFromDates(46023#, 46123#, Empty)
        TEST_Harness.AssertEqualText "rounding-padded-tab", "3W", KPR_Dates_PillarFromDates(46096#, 46121#, "  floor" & vbTab)
    'A control never vectorizes: its array is CONTROL_NOT_SCALAR, not a shape
        TEST_Harness.AssertErrorValue "control-array", _
            ERROR_VALUE, _
            KPR_Dates_AddMonths(46053#, 1, Array(True, True))

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the unexpected error against the case
        TEST_Harness.RecordCaseError

End Sub

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function ForwardDayOfWeek( _
    ByVal DateIn As Variant, _
    Optional ByVal WeekBase As Variant) As Variant
'
'------------------------------------------------------------------------------
'                          FORWARD AN OMITTED CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Calls KPR_Dates_DayOfWeek with its own omitted Optional argument, which arrives as Missing.
'
' WHY THIS EXISTS
'   A VBA caller that forwards an omitted Optional argument passes Missing, which must select the
'   default like an omitted control (data contract, Control values)
'
' INPUTS
'   DateIn
'     The date
'
'   WeekBase
'     Left out, so that it is Missing
'
' RETURNS
'   Variant
'     The result of KPR_Dates_DayOfWeek
'
' BEHAVIOR
'   Forwards both arguments
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   KPR_Dates_DayOfWeek
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' FORWARD
'------------------------------------------------------------------------------
    'WeekBase is Missing when the caller left it out
        ForwardDayOfWeek = KPR_Dates_DayOfWeek(DateIn, WeekBase)

End Function

Private Function ForwardAddMonths( _
    ByVal DateIn As Variant, _
    ByVal nMonths As Variant, _
    Optional ByVal Opt_KeepEOM As Variant) As Variant
'
'------------------------------------------------------------------------------
'                          FORWARD AN OMITTED CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Calls KPR_Dates_AddMonths with its own omitted Optional argument, which arrives as Missing.
'
' WHY THIS EXISTS
'   A VBA caller that forwards an omitted Optional argument passes Missing, which must select the
'   default like an omitted control (data contract, Control values)
'
' INPUTS
'   DateIn
'     The date
'
'   nMonths
'     The months to add
'
'   Opt_KeepEOM
'     Left out, so that it is Missing
'
' RETURNS
'   Variant
'     The result of KPR_Dates_AddMonths
'
' BEHAVIOR
'   Forwards every argument
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   KPR_Dates_AddMonths
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' FORWARD
'------------------------------------------------------------------------------
    'Opt_KeepEOM is Missing when the caller left it out
        ForwardAddMonths = KPR_Dates_AddMonths(DateIn, nMonths, Opt_KeepEOM)

End Function

Private Function ForwardPillarFromDates( _
    ByVal StartDate As Variant, _
    ByVal EndDate As Variant, _
    Optional ByVal Opt_Rounding As Variant) As Variant
'
'------------------------------------------------------------------------------
'                          FORWARD AN OMITTED CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Calls KPR_Dates_PillarFromDates with its own omitted Optional argument, which arrives as Missing.
'
' WHY THIS EXISTS
'   A VBA caller that forwards an omitted Optional argument passes Missing, which must select the
'   default like an omitted control (data contract, Control values)
'
' INPUTS
'   StartDate
'     The anchor date
'
'   EndDate
'     The end date
'
'   Opt_Rounding
'     Left out, so that it is Missing
'
' RETURNS
'   Variant
'     The result of KPR_Dates_PillarFromDates
'
' BEHAVIOR
'   Forwards every argument
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   KPR_Dates_PillarFromDates
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' FORWARD
'------------------------------------------------------------------------------
    'Opt_Rounding is Missing when the caller left it out
        ForwardPillarFromDates = KPR_Dates_PillarFromDates(StartDate, EndDate, Opt_Rounding)

End Function
