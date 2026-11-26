Attribute VB_Name = "TEST_CoreDates"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: TEST_CoreDates
'------------------------------------------------------------------------------
' PURPOSE
'   Tests of CORE_Dates: the window and the domains, the calendar rules, the period boundaries,
'   the shifts by days and months, the weekday locators and the pillar parser, shift and
'   format
'
' WHY THIS EXISTS
'   The calculation contracts CALC-02 to CALC-09 define each rule. These cases check the
'   code against reference values that the code does not produce, on and beside every
'   edge of the window
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   RunCoreDatesTests runs the suite. Option Private Module keeps it out of the workbook's
'   external surface. The cases build their inputs in memory and change no Excel state.
'   CertifyCoreDates certifies its normal run, as .github/excel-evidence-policy.json lists it
'
' ERROR POLICY
'   As TEST_Harness.EndRun: a run that does not pass raises TEST_Harness.ERR_TEST_FAILED,
'   or the unexpected error that escaped the runner. A refused start raises
'   TEST_Harness.ERR_TEST_DIRTY_START
'
' DEPENDENCIES
'   CORE_Dates, CORE_Err, TEST_Harness
'
' NOTES
'   In a development workbook holding the modules of src/ and tests/modules/, type
'   TEST_CoreDates.RunCoreDatesTests in the Immediate window. Every expected serial and
'   token is a literal computed by tests/derivations/date_core_reference.py: the calendar
'   rules from Python's calendar, and the pillar format from a second implementation of
'   CALC-09 written from the contract. Its author also wrote the VBA, so the pillar values
'   are a cross-check, not an independent reference
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'A complete run starts every case and evaluates every assertion once
    Private Const SUITE_NAME            As String = "core-dates"        'Suite name in the run log
    Private Const EXPECTED_CASES        As Long = 13                 'Cases a complete run starts
    Private Const EXPECTED_ASSERTIONS   As Long = 2111               'Assertions a complete run evaluates

    'The values an output holds before a call, which a rejection must leave in place
    Private Const SENTINEL_SERIAL       As Double = 36526#              '2000-01-01
    Private Const SENTINEL_DELTA        As Double = -424242#            'A delta no token gives
    Private Const SENTINEL_TOKEN        As String = "UNTOUCHED"         'A token no format returns


'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    'The arguments of each assertion of the last probe, by assertion and argument
    Private mFact(1 To 4, 1 To 3) As Variant


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub RunCoreDatesTests()

'
'------------------------------------------------------------------------------
'                             RUN CORE DATE TESTS
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case in a fixed order between BeginRun and EndRun. A passing run ends with
'   RESULT=PASS; suite=core-dates; operation=PASS; completeness=COMPLETE; cases=13;
'   assertions=2111; failures=0; cleanup=PASS; recovery=NONE
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
    Dim ErrNumber       As Long     'Runner error number; 0 for none
    Dim ErrSource       As String   'Runner error source
    Dim ErrDescription  As String   'Runner error description

'------------------------------------------------------------------------------
' RUN CASES
'------------------------------------------------------------------------------
    'Start the run; a refused start propagates directly
        TEST_Harness.BeginRun SUITE_NAME, "NORMAL", EXPECTED_CASES, EXPECTED_ASSERTIONS
    'Send an error that escapes a case to the runner handler
        On Error GoTo RunFailed
    'Run every case in a fixed order
        CaseWindowAndDomains
        CaseLeapYearsAndLengths
        CaseDayOfWeek
        CasePeriodBoundaries
        CasePeriodEndTests
        CaseAddDays
        CaseAddMonths
        CaseNthWeekday
        CaseLastWeekday
        CaseParsePillar
        CaseAddPillar
        CaseFormatPillar
        CasePillarLegacy

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

Public Sub CertifyCoreDates( _
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
'   Runs RunCoreDatesTests, saves its log, certifies the run, then saves the certification log.
'
' ERROR POLICY
'   TEST_Harness.EndCertification raises when the certification did not pass.
'
' DEPENDENCIES
'   RunCoreDatesTests, TEST_Harness
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
    Dim RunErrorNumber          As Long     'Error that ended the run; 0 for none
    Dim RunErrorDescription     As String   'Its description

'------------------------------------------------------------------------------
' CERTIFY
'------------------------------------------------------------------------------
    'Refuse state left by an earlier run and start the certification
        TEST_Harness.BeginCertification SUITE_NAME
    'Run the suite and capture the error that ended it
        On Error Resume Next
        RunCoreDatesTests
        RunErrorNumber = Err.Number
        RunErrorDescription = Err.Description
    'Restore normal error handling
        On Error GoTo 0
    'Save the log and certify the run
        TEST_Harness.SaveLog Folder & "\core-dates.log"
        TEST_Harness.CertifyRun "NORMAL", True, RunErrorNumber, RunErrorDescription
    'Save the certification log, raising when the certification did not pass
        TEST_Harness.EndCertification Folder & "\core-dates-certification.log"

End Sub

'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseWindowAndDomains()

'
'------------------------------------------------------------------------------
'                       CHECK THE WINDOW AND THE DOMAINS
'------------------------------------------------------------------------------
' PURPOSE
'   The window includes both of its edges, and each domain its bounds. 25 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date core follows the calculation contracts
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Checks the window predicate at its edges and each domain check on and beside its bounds
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Dates, CORE_Err, TEST_Harness
'
' NOTES
'   Expected values come from tests/ reference computations in Python, not from the code
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "window-and-domains"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'The window starts at serial 61 and ends at serial 2958465, both included
        ProbeFlag False, CORE_Dates.IsDateInWindow(CDate(2#))
        TEST_Harness.AssertTrue "window-2", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsDateInWindow(CDate(59#))
        TEST_Harness.AssertTrue "window-59", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsDateInWindow(CDate(60#))
        TEST_Harness.AssertTrue "window-60", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsDateInWindow(CDate(61#))
        TEST_Harness.AssertTrue "window-61", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsDateInWindow(CDate(2958465#))
        TEST_Harness.AssertTrue "window-2958465", mFact(1, 1), mFact(1, 2)
    'The year domain is 1900 to 9999
        ProbeCondition cdDomainYear, CORE_Dates.CheckYear(-1)
        TEST_Harness.AssertEqualText "year--1", mFact(1, 1), mFact(1, 2)
        ProbeCondition cdDomainYear, CORE_Dates.CheckYear(0)
        TEST_Harness.AssertEqualText "year-0", mFact(1, 1), mFact(1, 2)
        ProbeCondition cdDomainYear, CORE_Dates.CheckYear(1899)
        TEST_Harness.AssertEqualText "year-1899", mFact(1, 1), mFact(1, 2)
        ProbeCondition cdNone, CORE_Dates.CheckYear(1900)
        TEST_Harness.AssertEqualText "year-1900", mFact(1, 1), mFact(1, 2)
        ProbeCondition cdNone, CORE_Dates.CheckYear(5000)
        TEST_Harness.AssertEqualText "year-5000", mFact(1, 1), mFact(1, 2)
        ProbeCondition cdNone, CORE_Dates.CheckYear(9999)
        TEST_Harness.AssertEqualText "year-9999", mFact(1, 1), mFact(1, 2)
        ProbeCondition cdDomainYear, CORE_Dates.CheckYear(10000)
        TEST_Harness.AssertEqualText "year-10000", mFact(1, 1), mFact(1, 2)
    'The month domain is 1 to 12
        ProbeCondition cdDomainMonth, CORE_Dates.CheckMonth(-1)
        TEST_Harness.AssertEqualText "month--1", mFact(1, 1), mFact(1, 2)
        ProbeCondition cdDomainMonth, CORE_Dates.CheckMonth(0)
        TEST_Harness.AssertEqualText "month-0", mFact(1, 1), mFact(1, 2)
        ProbeCondition cdNone, CORE_Dates.CheckMonth(1)
        TEST_Harness.AssertEqualText "month-1", mFact(1, 1), mFact(1, 2)
        ProbeCondition cdNone, CORE_Dates.CheckMonth(12)
        TEST_Harness.AssertEqualText "month-12", mFact(1, 1), mFact(1, 2)
        ProbeCondition cdDomainMonth, CORE_Dates.CheckMonth(13)
        TEST_Harness.AssertEqualText "month-13", mFact(1, 1), mFact(1, 2)
    'The weekday domain is 1 to 7
        ProbeCondition cdDomainWeekday, CORE_Dates.CheckWeekday(0)
        TEST_Harness.AssertEqualText "weekday-0", mFact(1, 1), mFact(1, 2)
        ProbeCondition cdNone, CORE_Dates.CheckWeekday(1)
        TEST_Harness.AssertEqualText "weekday-1", mFact(1, 1), mFact(1, 2)
        ProbeCondition cdNone, CORE_Dates.CheckWeekday(7)
        TEST_Harness.AssertEqualText "weekday-7", mFact(1, 1), mFact(1, 2)
        ProbeCondition cdDomainWeekday, CORE_Dates.CheckWeekday(8)
        TEST_Harness.AssertEqualText "weekday-8", mFact(1, 1), mFact(1, 2)
    'The occurrence domain is 1 to 5
        ProbeCondition cdDomainOccurrence, CORE_Dates.CheckOccurrence(0)
        TEST_Harness.AssertEqualText "occurrence-0", mFact(1, 1), mFact(1, 2)
        ProbeCondition cdNone, CORE_Dates.CheckOccurrence(1)
        TEST_Harness.AssertEqualText "occurrence-1", mFact(1, 1), mFact(1, 2)
        ProbeCondition cdNone, CORE_Dates.CheckOccurrence(5)
        TEST_Harness.AssertEqualText "occurrence-5", mFact(1, 1), mFact(1, 2)
        ProbeCondition cdDomainOccurrence, CORE_Dates.CheckOccurrence(6)
        TEST_Harness.AssertEqualText "occurrence-6", mFact(1, 1), mFact(1, 2)

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

Private Sub CaseLeapYearsAndLengths()

'
'------------------------------------------------------------------------------
'                         CHECK LEAP YEARS AND LENGTHS
'------------------------------------------------------------------------------
' PURPOSE
'   Leap years, month lengths and year lengths follow the Gregorian calendar. 46 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date core follows the calculation contracts
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Checks the leap rule on centuries and on the edges of the window, and every month length in a common and a leap year
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Dates, CORE_Err, TEST_Harness
'
' NOTES
'   Expected values come from tests/ reference computations in Python, not from the code
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "leap-years-and-lengths"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'The Gregorian leap rule: divisible by 4, except centuries not divisible by 400
        ProbeFlag False, CORE_Dates.IsLeapYear(1900)
        TEST_Harness.AssertTrue "leap-1900", mFact(1, 1), mFact(1, 2)
        ProbeLong 365, CORE_Dates.DaysInYear(1900)
        TEST_Harness.AssertEqualDouble "year-days-1900", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsLeapYear(1904)
        TEST_Harness.AssertTrue "leap-1904", mFact(1, 1), mFact(1, 2)
        ProbeLong 366, CORE_Dates.DaysInYear(1904)
        TEST_Harness.AssertEqualDouble "year-days-1904", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsLeapYear(1999)
        TEST_Harness.AssertTrue "leap-1999", mFact(1, 1), mFact(1, 2)
        ProbeLong 365, CORE_Dates.DaysInYear(1999)
        TEST_Harness.AssertEqualDouble "year-days-1999", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsLeapYear(2000)
        TEST_Harness.AssertTrue "leap-2000", mFact(1, 1), mFact(1, 2)
        ProbeLong 366, CORE_Dates.DaysInYear(2000)
        TEST_Harness.AssertEqualDouble "year-days-2000", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsLeapYear(2023)
        TEST_Harness.AssertTrue "leap-2023", mFact(1, 1), mFact(1, 2)
        ProbeLong 365, CORE_Dates.DaysInYear(2023)
        TEST_Harness.AssertEqualDouble "year-days-2023", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsLeapYear(2024)
        TEST_Harness.AssertTrue "leap-2024", mFact(1, 1), mFact(1, 2)
        ProbeLong 366, CORE_Dates.DaysInYear(2024)
        TEST_Harness.AssertEqualDouble "year-days-2024", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsLeapYear(2100)
        TEST_Harness.AssertTrue "leap-2100", mFact(1, 1), mFact(1, 2)
        ProbeLong 365, CORE_Dates.DaysInYear(2100)
        TEST_Harness.AssertEqualDouble "year-days-2100", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsLeapYear(2400)
        TEST_Harness.AssertTrue "leap-2400", mFact(1, 1), mFact(1, 2)
        ProbeLong 366, CORE_Dates.DaysInYear(2400)
        TEST_Harness.AssertEqualDouble "year-days-2400", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsLeapYear(9996)
        TEST_Harness.AssertTrue "leap-9996", mFact(1, 1), mFact(1, 2)
        ProbeLong 366, CORE_Dates.DaysInYear(9996)
        TEST_Harness.AssertEqualDouble "year-days-9996", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsLeapYear(9999)
        TEST_Harness.AssertTrue "leap-9999", mFact(1, 1), mFact(1, 2)
        ProbeLong 365, CORE_Dates.DaysInYear(9999)
        TEST_Harness.AssertEqualDouble "year-days-9999", mFact(1, 1), mFact(1, 2)
    'The length of each month in a common and in a leap year
        ProbeLong 31, CORE_Dates.DaysInMonth(2023, 1)
        TEST_Harness.AssertEqualDouble "month-2023-1", mFact(1, 1), mFact(1, 2)
        ProbeLong 28, CORE_Dates.DaysInMonth(2023, 2)
        TEST_Harness.AssertEqualDouble "month-2023-2", mFact(1, 1), mFact(1, 2)
        ProbeLong 31, CORE_Dates.DaysInMonth(2023, 3)
        TEST_Harness.AssertEqualDouble "month-2023-3", mFact(1, 1), mFact(1, 2)
        ProbeLong 30, CORE_Dates.DaysInMonth(2023, 4)
        TEST_Harness.AssertEqualDouble "month-2023-4", mFact(1, 1), mFact(1, 2)
        ProbeLong 31, CORE_Dates.DaysInMonth(2023, 5)
        TEST_Harness.AssertEqualDouble "month-2023-5", mFact(1, 1), mFact(1, 2)
        ProbeLong 30, CORE_Dates.DaysInMonth(2023, 6)
        TEST_Harness.AssertEqualDouble "month-2023-6", mFact(1, 1), mFact(1, 2)
        ProbeLong 31, CORE_Dates.DaysInMonth(2023, 7)
        TEST_Harness.AssertEqualDouble "month-2023-7", mFact(1, 1), mFact(1, 2)
        ProbeLong 31, CORE_Dates.DaysInMonth(2023, 8)
        TEST_Harness.AssertEqualDouble "month-2023-8", mFact(1, 1), mFact(1, 2)
        ProbeLong 30, CORE_Dates.DaysInMonth(2023, 9)
        TEST_Harness.AssertEqualDouble "month-2023-9", mFact(1, 1), mFact(1, 2)
        ProbeLong 31, CORE_Dates.DaysInMonth(2023, 10)
        TEST_Harness.AssertEqualDouble "month-2023-10", mFact(1, 1), mFact(1, 2)
        ProbeLong 30, CORE_Dates.DaysInMonth(2023, 11)
        TEST_Harness.AssertEqualDouble "month-2023-11", mFact(1, 1), mFact(1, 2)
        ProbeLong 31, CORE_Dates.DaysInMonth(2023, 12)
        TEST_Harness.AssertEqualDouble "month-2023-12", mFact(1, 1), mFact(1, 2)
        ProbeLong 31, CORE_Dates.DaysInMonth(2024, 1)
        TEST_Harness.AssertEqualDouble "month-2024-1", mFact(1, 1), mFact(1, 2)
        ProbeLong 29, CORE_Dates.DaysInMonth(2024, 2)
        TEST_Harness.AssertEqualDouble "month-2024-2", mFact(1, 1), mFact(1, 2)
        ProbeLong 31, CORE_Dates.DaysInMonth(2024, 3)
        TEST_Harness.AssertEqualDouble "month-2024-3", mFact(1, 1), mFact(1, 2)
        ProbeLong 30, CORE_Dates.DaysInMonth(2024, 4)
        TEST_Harness.AssertEqualDouble "month-2024-4", mFact(1, 1), mFact(1, 2)
        ProbeLong 31, CORE_Dates.DaysInMonth(2024, 5)
        TEST_Harness.AssertEqualDouble "month-2024-5", mFact(1, 1), mFact(1, 2)
        ProbeLong 30, CORE_Dates.DaysInMonth(2024, 6)
        TEST_Harness.AssertEqualDouble "month-2024-6", mFact(1, 1), mFact(1, 2)
        ProbeLong 31, CORE_Dates.DaysInMonth(2024, 7)
        TEST_Harness.AssertEqualDouble "month-2024-7", mFact(1, 1), mFact(1, 2)
        ProbeLong 31, CORE_Dates.DaysInMonth(2024, 8)
        TEST_Harness.AssertEqualDouble "month-2024-8", mFact(1, 1), mFact(1, 2)
        ProbeLong 30, CORE_Dates.DaysInMonth(2024, 9)
        TEST_Harness.AssertEqualDouble "month-2024-9", mFact(1, 1), mFact(1, 2)
        ProbeLong 31, CORE_Dates.DaysInMonth(2024, 10)
        TEST_Harness.AssertEqualDouble "month-2024-10", mFact(1, 1), mFact(1, 2)
        ProbeLong 30, CORE_Dates.DaysInMonth(2024, 11)
        TEST_Harness.AssertEqualDouble "month-2024-11", mFact(1, 1), mFact(1, 2)
        ProbeLong 31, CORE_Dates.DaysInMonth(2024, 12)
        TEST_Harness.AssertEqualDouble "month-2024-12", mFact(1, 1), mFact(1, 2)
    'An invalid month has no length
        ProbeLong 0, CORE_Dates.DaysInMonth(2026, 0)
        TEST_Harness.AssertEqualDouble "month-invalid-0", mFact(1, 1), mFact(1, 2)
        ProbeLong 0, CORE_Dates.DaysInMonth(2026, 13)
        TEST_Harness.AssertEqualDouble "month-invalid-13", mFact(1, 1), mFact(1, 2)

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

Private Sub CaseDayOfWeek()

'
'------------------------------------------------------------------------------
'                          CHECK THE DAY OF THE WEEK
'------------------------------------------------------------------------------
' PURPOSE
'   The weekday number follows the chosen base: Monday is 1, or Sunday is 1. 44 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date core follows the calculation contracts
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Numbers the weekday of each sample date under both bases
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Dates, CORE_Err, TEST_Harness
'
' NOTES
'   Expected values come from tests/ reference computations in Python, not from the code
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "day-of-week"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Monday-based and Sunday-based numbering of the sample dates
        ProbeLong 4, CORE_Dates.DayOfWeek(CDate(61#), True)
        TEST_Harness.AssertEqualDouble "monday-1900-03-01", mFact(1, 1), mFact(1, 2)
        ProbeLong 5, CORE_Dates.DayOfWeek(CDate(61#), False)
        TEST_Harness.AssertEqualDouble "sunday-1900-03-01", mFact(1, 1), mFact(1, 2)
        ProbeLong 6, CORE_Dates.DayOfWeek(CDate(91#), True)
        TEST_Harness.AssertEqualDouble "monday-1900-03-31", mFact(1, 1), mFact(1, 2)
        ProbeLong 7, CORE_Dates.DayOfWeek(CDate(91#), False)
        TEST_Harness.AssertEqualDouble "sunday-1900-03-31", mFact(1, 1), mFact(1, 2)
        ProbeLong 6, CORE_Dates.DayOfWeek(CDate(182#), True)
        TEST_Harness.AssertEqualDouble "monday-1900-06-30", mFact(1, 1), mFact(1, 2)
        ProbeLong 7, CORE_Dates.DayOfWeek(CDate(182#), False)
        TEST_Harness.AssertEqualDouble "sunday-1900-06-30", mFact(1, 1), mFact(1, 2)
        ProbeLong 1, CORE_Dates.DayOfWeek(CDate(366#), True)
        TEST_Harness.AssertEqualDouble "monday-1900-12-31", mFact(1, 1), mFact(1, 2)
        ProbeLong 2, CORE_Dates.DayOfWeek(CDate(366#), False)
        TEST_Harness.AssertEqualDouble "sunday-1900-12-31", mFact(1, 1), mFact(1, 2)
        ProbeLong 5, CORE_Dates.DayOfWeek(CDate(36525#), True)
        TEST_Harness.AssertEqualDouble "monday-1999-12-31", mFact(1, 1), mFact(1, 2)
        ProbeLong 6, CORE_Dates.DayOfWeek(CDate(36525#), False)
        TEST_Harness.AssertEqualDouble "sunday-1999-12-31", mFact(1, 1), mFact(1, 2)
        ProbeLong 1, CORE_Dates.DayOfWeek(CDate(36584#), True)
        TEST_Harness.AssertEqualDouble "monday-2000-02-28", mFact(1, 1), mFact(1, 2)
        ProbeLong 2, CORE_Dates.DayOfWeek(CDate(36584#), False)
        TEST_Harness.AssertEqualDouble "sunday-2000-02-28", mFact(1, 1), mFact(1, 2)
        ProbeLong 2, CORE_Dates.DayOfWeek(CDate(36585#), True)
        TEST_Harness.AssertEqualDouble "monday-2000-02-29", mFact(1, 1), mFact(1, 2)
        ProbeLong 3, CORE_Dates.DayOfWeek(CDate(36585#), False)
        TEST_Harness.AssertEqualDouble "sunday-2000-02-29", mFact(1, 1), mFact(1, 2)
        ProbeLong 3, CORE_Dates.DayOfWeek(CDate(36586#), True)
        TEST_Harness.AssertEqualDouble "monday-2000-03-01", mFact(1, 1), mFact(1, 2)
        ProbeLong 4, CORE_Dates.DayOfWeek(CDate(36586#), False)
        TEST_Harness.AssertEqualDouble "sunday-2000-03-01", mFact(1, 1), mFact(1, 2)
        ProbeLong 4, CORE_Dates.DayOfWeek(CDate(45351#), True)
        TEST_Harness.AssertEqualDouble "monday-2024-02-29", mFact(1, 1), mFact(1, 2)
        ProbeLong 5, CORE_Dates.DayOfWeek(CDate(45351#), False)
        TEST_Harness.AssertEqualDouble "sunday-2024-02-29", mFact(1, 1), mFact(1, 2)
        ProbeLong 4, CORE_Dates.DayOfWeek(CDate(46023#), True)
        TEST_Harness.AssertEqualDouble "monday-2026-01-01", mFact(1, 1), mFact(1, 2)
        ProbeLong 5, CORE_Dates.DayOfWeek(CDate(46023#), False)
        TEST_Harness.AssertEqualDouble "sunday-2026-01-01", mFact(1, 1), mFact(1, 2)
        ProbeLong 6, CORE_Dates.DayOfWeek(CDate(46053#), True)
        TEST_Harness.AssertEqualDouble "monday-2026-01-31", mFact(1, 1), mFact(1, 2)
        ProbeLong 7, CORE_Dates.DayOfWeek(CDate(46053#), False)
        TEST_Harness.AssertEqualDouble "sunday-2026-01-31", mFact(1, 1), mFact(1, 2)
        ProbeLong 6, CORE_Dates.DayOfWeek(CDate(46081#), True)
        TEST_Harness.AssertEqualDouble "monday-2026-02-28", mFact(1, 1), mFact(1, 2)
        ProbeLong 7, CORE_Dates.DayOfWeek(CDate(46081#), False)
        TEST_Harness.AssertEqualDouble "sunday-2026-02-28", mFact(1, 1), mFact(1, 2)
        ProbeLong 2, CORE_Dates.DayOfWeek(CDate(46112#), True)
        TEST_Harness.AssertEqualDouble "monday-2026-03-31", mFact(1, 1), mFact(1, 2)
        ProbeLong 3, CORE_Dates.DayOfWeek(CDate(46112#), False)
        TEST_Harness.AssertEqualDouble "sunday-2026-03-31", mFact(1, 1), mFact(1, 2)
        ProbeLong 4, CORE_Dates.DayOfWeek(CDate(46142#), True)
        TEST_Harness.AssertEqualDouble "monday-2026-04-30", mFact(1, 1), mFact(1, 2)
        ProbeLong 5, CORE_Dates.DayOfWeek(CDate(46142#), False)
        TEST_Harness.AssertEqualDouble "sunday-2026-04-30", mFact(1, 1), mFact(1, 2)
        ProbeLong 2, CORE_Dates.DayOfWeek(CDate(46203#), True)
        TEST_Harness.AssertEqualDouble "monday-2026-06-30", mFact(1, 1), mFact(1, 2)
        ProbeLong 3, CORE_Dates.DayOfWeek(CDate(46203#), False)
        TEST_Harness.AssertEqualDouble "sunday-2026-06-30", mFact(1, 1), mFact(1, 2)
        ProbeLong 3, CORE_Dates.DayOfWeek(CDate(46295#), True)
        TEST_Harness.AssertEqualDouble "monday-2026-09-30", mFact(1, 1), mFact(1, 2)
        ProbeLong 4, CORE_Dates.DayOfWeek(CDate(46295#), False)
        TEST_Harness.AssertEqualDouble "sunday-2026-09-30", mFact(1, 1), mFact(1, 2)
        ProbeLong 3, CORE_Dates.DayOfWeek(CDate(46351#), True)
        TEST_Harness.AssertEqualDouble "monday-2026-11-25", mFact(1, 1), mFact(1, 2)
        ProbeLong 4, CORE_Dates.DayOfWeek(CDate(46351#), False)
        TEST_Harness.AssertEqualDouble "sunday-2026-11-25", mFact(1, 1), mFact(1, 2)
        ProbeLong 4, CORE_Dates.DayOfWeek(CDate(46387#), True)
        TEST_Harness.AssertEqualDouble "monday-2026-12-31", mFact(1, 1), mFact(1, 2)
        ProbeLong 5, CORE_Dates.DayOfWeek(CDate(46387#), False)
        TEST_Harness.AssertEqualDouble "sunday-2026-12-31", mFact(1, 1), mFact(1, 2)
        ProbeLong 7, CORE_Dates.DayOfWeek(CDate(73109#), True)
        TEST_Harness.AssertEqualDouble "monday-2100-02-28", mFact(1, 1), mFact(1, 2)
        ProbeLong 1, CORE_Dates.DayOfWeek(CDate(73109#), False)
        TEST_Harness.AssertEqualDouble "sunday-2100-02-28", mFact(1, 1), mFact(1, 2)
        ProbeLong 1, CORE_Dates.DayOfWeek(CDate(73110#), True)
        TEST_Harness.AssertEqualDouble "monday-2100-03-01", mFact(1, 1), mFact(1, 2)
        ProbeLong 2, CORE_Dates.DayOfWeek(CDate(73110#), False)
        TEST_Harness.AssertEqualDouble "sunday-2100-03-01", mFact(1, 1), mFact(1, 2)
        ProbeLong 5, CORE_Dates.DayOfWeek(CDate(2958101#), True)
        TEST_Harness.AssertEqualDouble "monday-9999-01-01", mFact(1, 1), mFact(1, 2)
        ProbeLong 6, CORE_Dates.DayOfWeek(CDate(2958101#), False)
        TEST_Harness.AssertEqualDouble "sunday-9999-01-01", mFact(1, 1), mFact(1, 2)
        ProbeLong 5, CORE_Dates.DayOfWeek(CDate(2958465#), True)
        TEST_Harness.AssertEqualDouble "monday-9999-12-31", mFact(1, 1), mFact(1, 2)
        ProbeLong 6, CORE_Dates.DayOfWeek(CDate(2958465#), False)
        TEST_Harness.AssertEqualDouble "sunday-9999-12-31", mFact(1, 1), mFact(1, 2)

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

Private Sub CasePeriodBoundaries()

'
'------------------------------------------------------------------------------
'                         CHECK THE PERIOD BOUNDARIES
'------------------------------------------------------------------------------
' PURPOSE
'   The first and last day of a month, quarter and year follow the calendar. 132 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date core follows the calculation contracts
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Builds each boundary of the sample dates, the early 1900 ones that leave the window included
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Dates, CORE_Err, TEST_Harness
'
' NOTES
'   Expected values come from tests/ reference computations in Python, not from the code
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "period-boundaries"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'The first and last day of the month, quarter and year of the sample dates
        ProbeDate 61#, CORE_Dates.BeginOfMonth(CDate(61#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-1900-03-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 91#, CORE_Dates.EndOfMonth(CDate(61#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-1900-03-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 2#, CORE_Dates.BeginOfQuarter(CDate(61#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-1900-03-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 91#, CORE_Dates.EndOfQuarter(CDate(61#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-1900-03-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 2#, CORE_Dates.BeginOfYear(CDate(61#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-1900-03-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 366#, CORE_Dates.EndOfYear(CDate(61#))
        TEST_Harness.AssertEqualDouble "EndOfYear-1900-03-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 61#, CORE_Dates.BeginOfMonth(CDate(91#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-1900-03-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 91#, CORE_Dates.EndOfMonth(CDate(91#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-1900-03-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 2#, CORE_Dates.BeginOfQuarter(CDate(91#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-1900-03-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 91#, CORE_Dates.EndOfQuarter(CDate(91#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-1900-03-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 2#, CORE_Dates.BeginOfYear(CDate(91#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-1900-03-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 366#, CORE_Dates.EndOfYear(CDate(91#))
        TEST_Harness.AssertEqualDouble "EndOfYear-1900-03-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 153#, CORE_Dates.BeginOfMonth(CDate(182#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-1900-06-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 182#, CORE_Dates.EndOfMonth(CDate(182#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-1900-06-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 92#, CORE_Dates.BeginOfQuarter(CDate(182#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-1900-06-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 182#, CORE_Dates.EndOfQuarter(CDate(182#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-1900-06-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 2#, CORE_Dates.BeginOfYear(CDate(182#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-1900-06-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 366#, CORE_Dates.EndOfYear(CDate(182#))
        TEST_Harness.AssertEqualDouble "EndOfYear-1900-06-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 336#, CORE_Dates.BeginOfMonth(CDate(366#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-1900-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 366#, CORE_Dates.EndOfMonth(CDate(366#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-1900-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 275#, CORE_Dates.BeginOfQuarter(CDate(366#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-1900-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 366#, CORE_Dates.EndOfQuarter(CDate(366#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-1900-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 2#, CORE_Dates.BeginOfYear(CDate(366#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-1900-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 366#, CORE_Dates.EndOfYear(CDate(366#))
        TEST_Harness.AssertEqualDouble "EndOfYear-1900-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 36495#, CORE_Dates.BeginOfMonth(CDate(36525#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-1999-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 36525#, CORE_Dates.EndOfMonth(CDate(36525#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-1999-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 36434#, CORE_Dates.BeginOfQuarter(CDate(36525#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-1999-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 36525#, CORE_Dates.EndOfQuarter(CDate(36525#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-1999-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 36161#, CORE_Dates.BeginOfYear(CDate(36525#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-1999-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 36525#, CORE_Dates.EndOfYear(CDate(36525#))
        TEST_Harness.AssertEqualDouble "EndOfYear-1999-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 36557#, CORE_Dates.BeginOfMonth(CDate(36584#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-2000-02-28", mFact(1, 1), mFact(1, 2)
        ProbeDate 36585#, CORE_Dates.EndOfMonth(CDate(36584#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-2000-02-28", mFact(1, 1), mFact(1, 2)
        ProbeDate 36526#, CORE_Dates.BeginOfQuarter(CDate(36584#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-2000-02-28", mFact(1, 1), mFact(1, 2)
        ProbeDate 36616#, CORE_Dates.EndOfQuarter(CDate(36584#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-2000-02-28", mFact(1, 1), mFact(1, 2)
        ProbeDate 36526#, CORE_Dates.BeginOfYear(CDate(36584#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-2000-02-28", mFact(1, 1), mFact(1, 2)
        ProbeDate 36891#, CORE_Dates.EndOfYear(CDate(36584#))
        TEST_Harness.AssertEqualDouble "EndOfYear-2000-02-28", mFact(1, 1), mFact(1, 2)
        ProbeDate 36557#, CORE_Dates.BeginOfMonth(CDate(36585#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-2000-02-29", mFact(1, 1), mFact(1, 2)
        ProbeDate 36585#, CORE_Dates.EndOfMonth(CDate(36585#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-2000-02-29", mFact(1, 1), mFact(1, 2)
        ProbeDate 36526#, CORE_Dates.BeginOfQuarter(CDate(36585#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-2000-02-29", mFact(1, 1), mFact(1, 2)
        ProbeDate 36616#, CORE_Dates.EndOfQuarter(CDate(36585#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-2000-02-29", mFact(1, 1), mFact(1, 2)
        ProbeDate 36526#, CORE_Dates.BeginOfYear(CDate(36585#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-2000-02-29", mFact(1, 1), mFact(1, 2)
        ProbeDate 36891#, CORE_Dates.EndOfYear(CDate(36585#))
        TEST_Harness.AssertEqualDouble "EndOfYear-2000-02-29", mFact(1, 1), mFact(1, 2)
        ProbeDate 36586#, CORE_Dates.BeginOfMonth(CDate(36586#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-2000-03-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 36616#, CORE_Dates.EndOfMonth(CDate(36586#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-2000-03-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 36526#, CORE_Dates.BeginOfQuarter(CDate(36586#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-2000-03-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 36616#, CORE_Dates.EndOfQuarter(CDate(36586#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-2000-03-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 36526#, CORE_Dates.BeginOfYear(CDate(36586#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-2000-03-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 36891#, CORE_Dates.EndOfYear(CDate(36586#))
        TEST_Harness.AssertEqualDouble "EndOfYear-2000-03-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 45323#, CORE_Dates.BeginOfMonth(CDate(45351#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-2024-02-29", mFact(1, 1), mFact(1, 2)
        ProbeDate 45351#, CORE_Dates.EndOfMonth(CDate(45351#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-2024-02-29", mFact(1, 1), mFact(1, 2)
        ProbeDate 45292#, CORE_Dates.BeginOfQuarter(CDate(45351#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-2024-02-29", mFact(1, 1), mFact(1, 2)
        ProbeDate 45382#, CORE_Dates.EndOfQuarter(CDate(45351#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-2024-02-29", mFact(1, 1), mFact(1, 2)
        ProbeDate 45292#, CORE_Dates.BeginOfYear(CDate(45351#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-2024-02-29", mFact(1, 1), mFact(1, 2)
        ProbeDate 45657#, CORE_Dates.EndOfYear(CDate(45351#))
        TEST_Harness.AssertEqualDouble "EndOfYear-2024-02-29", mFact(1, 1), mFact(1, 2)
        ProbeDate 46023#, CORE_Dates.BeginOfMonth(CDate(46023#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-2026-01-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 46053#, CORE_Dates.EndOfMonth(CDate(46023#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-2026-01-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 46023#, CORE_Dates.BeginOfQuarter(CDate(46023#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-2026-01-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 46112#, CORE_Dates.EndOfQuarter(CDate(46023#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-2026-01-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 46023#, CORE_Dates.BeginOfYear(CDate(46023#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-2026-01-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 46387#, CORE_Dates.EndOfYear(CDate(46023#))
        TEST_Harness.AssertEqualDouble "EndOfYear-2026-01-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 46023#, CORE_Dates.BeginOfMonth(CDate(46053#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-2026-01-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 46053#, CORE_Dates.EndOfMonth(CDate(46053#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-2026-01-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 46023#, CORE_Dates.BeginOfQuarter(CDate(46053#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-2026-01-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 46112#, CORE_Dates.EndOfQuarter(CDate(46053#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-2026-01-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 46023#, CORE_Dates.BeginOfYear(CDate(46053#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-2026-01-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 46387#, CORE_Dates.EndOfYear(CDate(46053#))
        TEST_Harness.AssertEqualDouble "EndOfYear-2026-01-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 46054#, CORE_Dates.BeginOfMonth(CDate(46081#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-2026-02-28", mFact(1, 1), mFact(1, 2)
        ProbeDate 46081#, CORE_Dates.EndOfMonth(CDate(46081#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-2026-02-28", mFact(1, 1), mFact(1, 2)
        ProbeDate 46023#, CORE_Dates.BeginOfQuarter(CDate(46081#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-2026-02-28", mFact(1, 1), mFact(1, 2)
        ProbeDate 46112#, CORE_Dates.EndOfQuarter(CDate(46081#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-2026-02-28", mFact(1, 1), mFact(1, 2)
        ProbeDate 46023#, CORE_Dates.BeginOfYear(CDate(46081#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-2026-02-28", mFact(1, 1), mFact(1, 2)
        ProbeDate 46387#, CORE_Dates.EndOfYear(CDate(46081#))
        TEST_Harness.AssertEqualDouble "EndOfYear-2026-02-28", mFact(1, 1), mFact(1, 2)
        ProbeDate 46082#, CORE_Dates.BeginOfMonth(CDate(46112#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-2026-03-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 46112#, CORE_Dates.EndOfMonth(CDate(46112#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-2026-03-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 46023#, CORE_Dates.BeginOfQuarter(CDate(46112#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-2026-03-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 46112#, CORE_Dates.EndOfQuarter(CDate(46112#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-2026-03-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 46023#, CORE_Dates.BeginOfYear(CDate(46112#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-2026-03-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 46387#, CORE_Dates.EndOfYear(CDate(46112#))
        TEST_Harness.AssertEqualDouble "EndOfYear-2026-03-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 46113#, CORE_Dates.BeginOfMonth(CDate(46142#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-2026-04-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 46142#, CORE_Dates.EndOfMonth(CDate(46142#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-2026-04-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 46113#, CORE_Dates.BeginOfQuarter(CDate(46142#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-2026-04-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 46203#, CORE_Dates.EndOfQuarter(CDate(46142#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-2026-04-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 46023#, CORE_Dates.BeginOfYear(CDate(46142#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-2026-04-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 46387#, CORE_Dates.EndOfYear(CDate(46142#))
        TEST_Harness.AssertEqualDouble "EndOfYear-2026-04-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 46174#, CORE_Dates.BeginOfMonth(CDate(46203#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-2026-06-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 46203#, CORE_Dates.EndOfMonth(CDate(46203#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-2026-06-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 46113#, CORE_Dates.BeginOfQuarter(CDate(46203#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-2026-06-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 46203#, CORE_Dates.EndOfQuarter(CDate(46203#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-2026-06-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 46023#, CORE_Dates.BeginOfYear(CDate(46203#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-2026-06-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 46387#, CORE_Dates.EndOfYear(CDate(46203#))
        TEST_Harness.AssertEqualDouble "EndOfYear-2026-06-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 46266#, CORE_Dates.BeginOfMonth(CDate(46295#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-2026-09-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 46295#, CORE_Dates.EndOfMonth(CDate(46295#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-2026-09-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 46204#, CORE_Dates.BeginOfQuarter(CDate(46295#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-2026-09-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 46295#, CORE_Dates.EndOfQuarter(CDate(46295#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-2026-09-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 46023#, CORE_Dates.BeginOfYear(CDate(46295#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-2026-09-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 46387#, CORE_Dates.EndOfYear(CDate(46295#))
        TEST_Harness.AssertEqualDouble "EndOfYear-2026-09-30", mFact(1, 1), mFact(1, 2)
        ProbeDate 46327#, CORE_Dates.BeginOfMonth(CDate(46351#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-2026-11-25", mFact(1, 1), mFact(1, 2)
        ProbeDate 46356#, CORE_Dates.EndOfMonth(CDate(46351#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-2026-11-25", mFact(1, 1), mFact(1, 2)
        ProbeDate 46296#, CORE_Dates.BeginOfQuarter(CDate(46351#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-2026-11-25", mFact(1, 1), mFact(1, 2)
        ProbeDate 46387#, CORE_Dates.EndOfQuarter(CDate(46351#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-2026-11-25", mFact(1, 1), mFact(1, 2)
        ProbeDate 46023#, CORE_Dates.BeginOfYear(CDate(46351#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-2026-11-25", mFact(1, 1), mFact(1, 2)
        ProbeDate 46387#, CORE_Dates.EndOfYear(CDate(46351#))
        TEST_Harness.AssertEqualDouble "EndOfYear-2026-11-25", mFact(1, 1), mFact(1, 2)
        ProbeDate 46357#, CORE_Dates.BeginOfMonth(CDate(46387#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-2026-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 46387#, CORE_Dates.EndOfMonth(CDate(46387#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-2026-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 46296#, CORE_Dates.BeginOfQuarter(CDate(46387#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-2026-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 46387#, CORE_Dates.EndOfQuarter(CDate(46387#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-2026-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 46023#, CORE_Dates.BeginOfYear(CDate(46387#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-2026-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 46387#, CORE_Dates.EndOfYear(CDate(46387#))
        TEST_Harness.AssertEqualDouble "EndOfYear-2026-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 73082#, CORE_Dates.BeginOfMonth(CDate(73109#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-2100-02-28", mFact(1, 1), mFact(1, 2)
        ProbeDate 73109#, CORE_Dates.EndOfMonth(CDate(73109#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-2100-02-28", mFact(1, 1), mFact(1, 2)
        ProbeDate 73051#, CORE_Dates.BeginOfQuarter(CDate(73109#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-2100-02-28", mFact(1, 1), mFact(1, 2)
        ProbeDate 73140#, CORE_Dates.EndOfQuarter(CDate(73109#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-2100-02-28", mFact(1, 1), mFact(1, 2)
        ProbeDate 73051#, CORE_Dates.BeginOfYear(CDate(73109#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-2100-02-28", mFact(1, 1), mFact(1, 2)
        ProbeDate 73415#, CORE_Dates.EndOfYear(CDate(73109#))
        TEST_Harness.AssertEqualDouble "EndOfYear-2100-02-28", mFact(1, 1), mFact(1, 2)
        ProbeDate 73110#, CORE_Dates.BeginOfMonth(CDate(73110#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-2100-03-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 73140#, CORE_Dates.EndOfMonth(CDate(73110#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-2100-03-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 73051#, CORE_Dates.BeginOfQuarter(CDate(73110#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-2100-03-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 73140#, CORE_Dates.EndOfQuarter(CDate(73110#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-2100-03-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 73051#, CORE_Dates.BeginOfYear(CDate(73110#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-2100-03-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 73415#, CORE_Dates.EndOfYear(CDate(73110#))
        TEST_Harness.AssertEqualDouble "EndOfYear-2100-03-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 2958101#, CORE_Dates.BeginOfMonth(CDate(2958101#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-9999-01-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 2958131#, CORE_Dates.EndOfMonth(CDate(2958101#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-9999-01-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 2958101#, CORE_Dates.BeginOfQuarter(CDate(2958101#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-9999-01-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 2958190#, CORE_Dates.EndOfQuarter(CDate(2958101#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-9999-01-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 2958101#, CORE_Dates.BeginOfYear(CDate(2958101#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-9999-01-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 2958465#, CORE_Dates.EndOfYear(CDate(2958101#))
        TEST_Harness.AssertEqualDouble "EndOfYear-9999-01-01", mFact(1, 1), mFact(1, 2)
        ProbeDate 2958435#, CORE_Dates.BeginOfMonth(CDate(2958465#))
        TEST_Harness.AssertEqualDouble "BeginOfMonth-9999-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 2958465#, CORE_Dates.EndOfMonth(CDate(2958465#))
        TEST_Harness.AssertEqualDouble "EndOfMonth-9999-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 2958374#, CORE_Dates.BeginOfQuarter(CDate(2958465#))
        TEST_Harness.AssertEqualDouble "BeginOfQuarter-9999-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 2958465#, CORE_Dates.EndOfQuarter(CDate(2958465#))
        TEST_Harness.AssertEqualDouble "EndOfQuarter-9999-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 2958101#, CORE_Dates.BeginOfYear(CDate(2958465#))
        TEST_Harness.AssertEqualDouble "BeginOfYear-9999-12-31", mFact(1, 1), mFact(1, 2)
        ProbeDate 2958465#, CORE_Dates.EndOfYear(CDate(2958465#))
        TEST_Harness.AssertEqualDouble "EndOfYear-9999-12-31", mFact(1, 1), mFact(1, 2)

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

Private Sub CasePeriodEndTests()

'
'------------------------------------------------------------------------------
'                          CHECK THE PERIOD-END TESTS
'------------------------------------------------------------------------------
' PURPOSE
'   A date is a month, quarter or year end exactly when it is the last day of the period. 66 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date core follows the calculation contracts
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Tests the sample dates against each period end
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Dates, CORE_Err, TEST_Harness
'
' NOTES
'   Expected values come from tests/ reference computations in Python, not from the code
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "period-end-tests"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'A date is a period end when it is the last day of the period
        ProbeFlag False, CORE_Dates.IsMonthEnd(CDate(61#))
        TEST_Harness.AssertTrue "IsMonthEnd-1900-03-01", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsQuarterEnd(CDate(61#))
        TEST_Harness.AssertTrue "IsQuarterEnd-1900-03-01", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsYearEnd(CDate(61#))
        TEST_Harness.AssertTrue "IsYearEnd-1900-03-01", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsMonthEnd(CDate(91#))
        TEST_Harness.AssertTrue "IsMonthEnd-1900-03-31", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsQuarterEnd(CDate(91#))
        TEST_Harness.AssertTrue "IsQuarterEnd-1900-03-31", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsYearEnd(CDate(91#))
        TEST_Harness.AssertTrue "IsYearEnd-1900-03-31", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsMonthEnd(CDate(182#))
        TEST_Harness.AssertTrue "IsMonthEnd-1900-06-30", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsQuarterEnd(CDate(182#))
        TEST_Harness.AssertTrue "IsQuarterEnd-1900-06-30", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsYearEnd(CDate(182#))
        TEST_Harness.AssertTrue "IsYearEnd-1900-06-30", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsMonthEnd(CDate(366#))
        TEST_Harness.AssertTrue "IsMonthEnd-1900-12-31", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsQuarterEnd(CDate(366#))
        TEST_Harness.AssertTrue "IsQuarterEnd-1900-12-31", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsYearEnd(CDate(366#))
        TEST_Harness.AssertTrue "IsYearEnd-1900-12-31", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsMonthEnd(CDate(36525#))
        TEST_Harness.AssertTrue "IsMonthEnd-1999-12-31", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsQuarterEnd(CDate(36525#))
        TEST_Harness.AssertTrue "IsQuarterEnd-1999-12-31", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsYearEnd(CDate(36525#))
        TEST_Harness.AssertTrue "IsYearEnd-1999-12-31", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsMonthEnd(CDate(36584#))
        TEST_Harness.AssertTrue "IsMonthEnd-2000-02-28", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsQuarterEnd(CDate(36584#))
        TEST_Harness.AssertTrue "IsQuarterEnd-2000-02-28", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsYearEnd(CDate(36584#))
        TEST_Harness.AssertTrue "IsYearEnd-2000-02-28", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsMonthEnd(CDate(36585#))
        TEST_Harness.AssertTrue "IsMonthEnd-2000-02-29", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsQuarterEnd(CDate(36585#))
        TEST_Harness.AssertTrue "IsQuarterEnd-2000-02-29", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsYearEnd(CDate(36585#))
        TEST_Harness.AssertTrue "IsYearEnd-2000-02-29", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsMonthEnd(CDate(36586#))
        TEST_Harness.AssertTrue "IsMonthEnd-2000-03-01", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsQuarterEnd(CDate(36586#))
        TEST_Harness.AssertTrue "IsQuarterEnd-2000-03-01", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsYearEnd(CDate(36586#))
        TEST_Harness.AssertTrue "IsYearEnd-2000-03-01", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsMonthEnd(CDate(45351#))
        TEST_Harness.AssertTrue "IsMonthEnd-2024-02-29", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsQuarterEnd(CDate(45351#))
        TEST_Harness.AssertTrue "IsQuarterEnd-2024-02-29", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsYearEnd(CDate(45351#))
        TEST_Harness.AssertTrue "IsYearEnd-2024-02-29", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsMonthEnd(CDate(46023#))
        TEST_Harness.AssertTrue "IsMonthEnd-2026-01-01", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsQuarterEnd(CDate(46023#))
        TEST_Harness.AssertTrue "IsQuarterEnd-2026-01-01", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsYearEnd(CDate(46023#))
        TEST_Harness.AssertTrue "IsYearEnd-2026-01-01", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsMonthEnd(CDate(46053#))
        TEST_Harness.AssertTrue "IsMonthEnd-2026-01-31", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsQuarterEnd(CDate(46053#))
        TEST_Harness.AssertTrue "IsQuarterEnd-2026-01-31", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsYearEnd(CDate(46053#))
        TEST_Harness.AssertTrue "IsYearEnd-2026-01-31", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsMonthEnd(CDate(46081#))
        TEST_Harness.AssertTrue "IsMonthEnd-2026-02-28", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsQuarterEnd(CDate(46081#))
        TEST_Harness.AssertTrue "IsQuarterEnd-2026-02-28", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsYearEnd(CDate(46081#))
        TEST_Harness.AssertTrue "IsYearEnd-2026-02-28", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsMonthEnd(CDate(46112#))
        TEST_Harness.AssertTrue "IsMonthEnd-2026-03-31", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsQuarterEnd(CDate(46112#))
        TEST_Harness.AssertTrue "IsQuarterEnd-2026-03-31", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsYearEnd(CDate(46112#))
        TEST_Harness.AssertTrue "IsYearEnd-2026-03-31", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsMonthEnd(CDate(46142#))
        TEST_Harness.AssertTrue "IsMonthEnd-2026-04-30", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsQuarterEnd(CDate(46142#))
        TEST_Harness.AssertTrue "IsQuarterEnd-2026-04-30", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsYearEnd(CDate(46142#))
        TEST_Harness.AssertTrue "IsYearEnd-2026-04-30", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsMonthEnd(CDate(46203#))
        TEST_Harness.AssertTrue "IsMonthEnd-2026-06-30", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsQuarterEnd(CDate(46203#))
        TEST_Harness.AssertTrue "IsQuarterEnd-2026-06-30", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsYearEnd(CDate(46203#))
        TEST_Harness.AssertTrue "IsYearEnd-2026-06-30", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsMonthEnd(CDate(46295#))
        TEST_Harness.AssertTrue "IsMonthEnd-2026-09-30", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsQuarterEnd(CDate(46295#))
        TEST_Harness.AssertTrue "IsQuarterEnd-2026-09-30", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsYearEnd(CDate(46295#))
        TEST_Harness.AssertTrue "IsYearEnd-2026-09-30", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsMonthEnd(CDate(46351#))
        TEST_Harness.AssertTrue "IsMonthEnd-2026-11-25", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsQuarterEnd(CDate(46351#))
        TEST_Harness.AssertTrue "IsQuarterEnd-2026-11-25", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsYearEnd(CDate(46351#))
        TEST_Harness.AssertTrue "IsYearEnd-2026-11-25", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsMonthEnd(CDate(46387#))
        TEST_Harness.AssertTrue "IsMonthEnd-2026-12-31", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsQuarterEnd(CDate(46387#))
        TEST_Harness.AssertTrue "IsQuarterEnd-2026-12-31", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsYearEnd(CDate(46387#))
        TEST_Harness.AssertTrue "IsYearEnd-2026-12-31", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsMonthEnd(CDate(73109#))
        TEST_Harness.AssertTrue "IsMonthEnd-2100-02-28", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsQuarterEnd(CDate(73109#))
        TEST_Harness.AssertTrue "IsQuarterEnd-2100-02-28", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsYearEnd(CDate(73109#))
        TEST_Harness.AssertTrue "IsYearEnd-2100-02-28", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsMonthEnd(CDate(73110#))
        TEST_Harness.AssertTrue "IsMonthEnd-2100-03-01", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsQuarterEnd(CDate(73110#))
        TEST_Harness.AssertTrue "IsQuarterEnd-2100-03-01", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsYearEnd(CDate(73110#))
        TEST_Harness.AssertTrue "IsYearEnd-2100-03-01", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsMonthEnd(CDate(2958101#))
        TEST_Harness.AssertTrue "IsMonthEnd-9999-01-01", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsQuarterEnd(CDate(2958101#))
        TEST_Harness.AssertTrue "IsQuarterEnd-9999-01-01", mFact(1, 1), mFact(1, 2)
        ProbeFlag False, CORE_Dates.IsYearEnd(CDate(2958101#))
        TEST_Harness.AssertTrue "IsYearEnd-9999-01-01", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsMonthEnd(CDate(2958465#))
        TEST_Harness.AssertTrue "IsMonthEnd-9999-12-31", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsQuarterEnd(CDate(2958465#))
        TEST_Harness.AssertTrue "IsQuarterEnd-9999-12-31", mFact(1, 1), mFact(1, 2)
        ProbeFlag True, CORE_Dates.IsYearEnd(CDate(2958465#))
        TEST_Harness.AssertTrue "IsYearEnd-9999-12-31", mFact(1, 1), mFact(1, 2)

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

Private Sub CaseAddDays()

'
'------------------------------------------------------------------------------
'                           CHECK THE SHIFT BY DAYS
'------------------------------------------------------------------------------
' PURPOSE
'   A date moves by whole days, and a result outside the window is rejected. 36 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date core follows the calculation contracts
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Shifts dates by days on and beside both edges of the window, by a week count and by extreme counts
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Dates, CORE_Err, TEST_Harness
'
' NOTES
'   Expected values come from tests/ reference computations in Python, not from the code
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "add-days"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Whole calendar days, with the window checked on the result
        ProbeAddDays 46053#, 1#, True, 46054#
        TEST_Harness.AssertTrue "2026-01-31+(1)-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-01-31+(1)-value", mFact(2, 1), mFact(2, 2)
        ProbeAddDays 46023#, -1#, True, 46022#
        TEST_Harness.AssertTrue "2026-01-01+(-1)-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-01-01+(-1)-value", mFact(2, 1), mFact(2, 2)
        ProbeAddDays 46082#, -365#, True, 45717#
        TEST_Harness.AssertTrue "2026-03-01+(-365)-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-03-01+(-365)-value", mFact(2, 1), mFact(2, 2)
        ProbeAddDays 61#, 0#, True, 61#
        TEST_Harness.AssertTrue "1900-03-01+(0)-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-01+(0)-value", mFact(2, 1), mFact(2, 2)
        ProbeAddDays 61#, -1#, False, 0#
        TEST_Harness.AssertTrue "1900-03-01+(-1)-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-01+(-1)-value", mFact(2, 1), mFact(2, 2)
        ProbeAddDays 62#, -1#, True, 61#
        TEST_Harness.AssertTrue "1900-03-02+(-1)-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-02+(-1)-value", mFact(2, 1), mFact(2, 2)
        ProbeAddDays 2958464#, 1#, True, 2958465#
        TEST_Harness.AssertTrue "9999-12-30+(1)-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-12-30+(1)-value", mFact(2, 1), mFact(2, 2)
        ProbeAddDays 2958465#, 0#, True, 2958465#
        TEST_Harness.AssertTrue "9999-12-31+(0)-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-12-31+(0)-value", mFact(2, 1), mFact(2, 2)
        ProbeAddDays 2958465#, 1#, False, 0#
        TEST_Harness.AssertTrue "9999-12-31+(1)-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-12-31+(1)-value", mFact(2, 1), mFact(2, 2)
        ProbeAddDays 2958465#, -2958404#, True, 61#
        TEST_Harness.AssertTrue "9999-12-31+(-2958404)-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-12-31+(-2958404)-value", mFact(2, 1), mFact(2, 2)
        ProbeAddDays 2958465#, -2958405#, False, 0#
        TEST_Harness.AssertTrue "9999-12-31+(-2958405)-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-12-31+(-2958405)-value", mFact(2, 1), mFact(2, 2)
        ProbeAddDays 46023#, 364#, True, 46387#
        TEST_Harness.AssertTrue "2026-01-01+(364)-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-01-01+(364)-value", mFact(2, 1), mFact(2, 2)
        ProbeAddDays 46023#, -7#, True, 46016#
        TEST_Harness.AssertTrue "2026-01-01+(-7)-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-01-01+(-7)-value", mFact(2, 1), mFact(2, 2)
        ProbeAddDays 46387#, 364#, True, 46751#
        TEST_Harness.AssertTrue "2026-12-31+(364)-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-12-31+(364)-value", mFact(2, 1), mFact(2, 2)
        ProbeAddDays 46188#, -7000000#, False, 0#
        TEST_Harness.AssertTrue "2026-06-15+(-7000000)-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-06-15+(-7000000)-value", mFact(2, 1), mFact(2, 2)
        ProbeAddDays 46188#, (10# ^ 300), False, 0#
        TEST_Harness.AssertTrue "2026-06-15-huge+-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-06-15-huge+-value", mFact(2, 1), mFact(2, 2)
        ProbeAddDays 46188#, -(10# ^ 300), False, 0#
        TEST_Harness.AssertTrue "2026-06-15-huge--accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-06-15-huge--value", mFact(2, 1), mFact(2, 2)
    'The time of day is discarded before the shift
        ProbeAddDays 46387.75, 1#, True, 46388#
        TEST_Harness.AssertTrue "time-of-day-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "time-of-day-value", mFact(2, 1), mFact(2, 2)

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

Private Sub CaseAddMonths()

'
'------------------------------------------------------------------------------
'                          CHECK THE SHIFT BY MONTHS
'------------------------------------------------------------------------------
' PURPOSE
'   A date moves by months with the clip of a short month, or keeps its month end. 286 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date core follows the calculation contracts
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Shifts month ends and leap days by negative, zero and positive counts, to and past both edges of the window
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Dates, CORE_Err, TEST_Harness
'
' NOTES
'   Expected values come from tests/ reference computations in Python, not from the code
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "add-months"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Month shifts with the clip of a short month, with and without keeping a month end
        ProbeAddMonths 46053#, -13#, False, True, 45657#
        TEST_Harness.AssertTrue "2026-01-31-13-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-01-31-13-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46053#, -13#, True, True, 45657#
        TEST_Harness.AssertTrue "2026-01-31-13k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-01-31-13k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46053#, -1#, False, True, 46022#
        TEST_Harness.AssertTrue "2026-01-31-1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-01-31-1-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46053#, -1#, True, True, 46022#
        TEST_Harness.AssertTrue "2026-01-31-1k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-01-31-1k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46053#, 0#, False, True, 46053#
        TEST_Harness.AssertTrue "2026-01-31+0-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-01-31+0-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46053#, 0#, True, True, 46053#
        TEST_Harness.AssertTrue "2026-01-31+0k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-01-31+0k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46053#, 1#, False, True, 46081#
        TEST_Harness.AssertTrue "2026-01-31+1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-01-31+1-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46053#, 1#, True, True, 46081#
        TEST_Harness.AssertTrue "2026-01-31+1k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-01-31+1k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46053#, 2#, False, True, 46112#
        TEST_Harness.AssertTrue "2026-01-31+2-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-01-31+2-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46053#, 2#, True, True, 46112#
        TEST_Harness.AssertTrue "2026-01-31+2k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-01-31+2k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46053#, 12#, False, True, 46418#
        TEST_Harness.AssertTrue "2026-01-31+12-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-01-31+12-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46053#, 12#, True, True, 46418#
        TEST_Harness.AssertTrue "2026-01-31+12k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-01-31+12k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46053#, 48#, False, True, 47514#
        TEST_Harness.AssertTrue "2026-01-31+48-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-01-31+48-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46053#, 48#, True, True, 47514#
        TEST_Harness.AssertTrue "2026-01-31+48k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-01-31+48k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 45351#, -13#, False, True, 44955#
        TEST_Harness.AssertTrue "2024-02-29-13-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2024-02-29-13-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 45351#, -13#, True, True, 44957#
        TEST_Harness.AssertTrue "2024-02-29-13k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2024-02-29-13k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 45351#, -1#, False, True, 45320#
        TEST_Harness.AssertTrue "2024-02-29-1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2024-02-29-1-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 45351#, -1#, True, True, 45322#
        TEST_Harness.AssertTrue "2024-02-29-1k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2024-02-29-1k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 45351#, 0#, False, True, 45351#
        TEST_Harness.AssertTrue "2024-02-29+0-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2024-02-29+0-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 45351#, 0#, True, True, 45351#
        TEST_Harness.AssertTrue "2024-02-29+0k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2024-02-29+0k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 45351#, 1#, False, True, 45380#
        TEST_Harness.AssertTrue "2024-02-29+1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2024-02-29+1-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 45351#, 1#, True, True, 45382#
        TEST_Harness.AssertTrue "2024-02-29+1k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2024-02-29+1k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 45351#, 2#, False, True, 45411#
        TEST_Harness.AssertTrue "2024-02-29+2-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2024-02-29+2-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 45351#, 2#, True, True, 45412#
        TEST_Harness.AssertTrue "2024-02-29+2k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2024-02-29+2k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 45351#, 12#, False, True, 45716#
        TEST_Harness.AssertTrue "2024-02-29+12-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2024-02-29+12-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 45351#, 12#, True, True, 45716#
        TEST_Harness.AssertTrue "2024-02-29+12k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2024-02-29+12k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 45351#, 48#, False, True, 46812#
        TEST_Harness.AssertTrue "2024-02-29+48-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2024-02-29+48-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 45351#, 48#, True, True, 46812#
        TEST_Harness.AssertTrue "2024-02-29+48k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2024-02-29+48k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46081#, -13#, False, True, 45685#
        TEST_Harness.AssertTrue "2026-02-28-13-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-02-28-13-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46081#, -13#, True, True, 45688#
        TEST_Harness.AssertTrue "2026-02-28-13k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-02-28-13k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46081#, -1#, False, True, 46050#
        TEST_Harness.AssertTrue "2026-02-28-1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-02-28-1-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46081#, -1#, True, True, 46053#
        TEST_Harness.AssertTrue "2026-02-28-1k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-02-28-1k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46081#, 0#, False, True, 46081#
        TEST_Harness.AssertTrue "2026-02-28+0-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-02-28+0-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46081#, 0#, True, True, 46081#
        TEST_Harness.AssertTrue "2026-02-28+0k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-02-28+0k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46081#, 1#, False, True, 46109#
        TEST_Harness.AssertTrue "2026-02-28+1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-02-28+1-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46081#, 1#, True, True, 46112#
        TEST_Harness.AssertTrue "2026-02-28+1k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-02-28+1k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46081#, 2#, False, True, 46140#
        TEST_Harness.AssertTrue "2026-02-28+2-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-02-28+2-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46081#, 2#, True, True, 46142#
        TEST_Harness.AssertTrue "2026-02-28+2k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-02-28+2k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46081#, 12#, False, True, 46446#
        TEST_Harness.AssertTrue "2026-02-28+12-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-02-28+12-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46081#, 12#, True, True, 46446#
        TEST_Harness.AssertTrue "2026-02-28+12k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-02-28+12k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46081#, 48#, False, True, 47542#
        TEST_Harness.AssertTrue "2026-02-28+48-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-02-28+48-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46081#, 48#, True, True, 47542#
        TEST_Harness.AssertTrue "2026-02-28+48k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-02-28+48k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46112#, -13#, False, True, 45716#
        TEST_Harness.AssertTrue "2026-03-31-13-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-03-31-13-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46112#, -13#, True, True, 45716#
        TEST_Harness.AssertTrue "2026-03-31-13k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-03-31-13k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46112#, -1#, False, True, 46081#
        TEST_Harness.AssertTrue "2026-03-31-1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-03-31-1-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46112#, -1#, True, True, 46081#
        TEST_Harness.AssertTrue "2026-03-31-1k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-03-31-1k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46112#, 0#, False, True, 46112#
        TEST_Harness.AssertTrue "2026-03-31+0-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-03-31+0-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46112#, 0#, True, True, 46112#
        TEST_Harness.AssertTrue "2026-03-31+0k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-03-31+0k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46112#, 1#, False, True, 46142#
        TEST_Harness.AssertTrue "2026-03-31+1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-03-31+1-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46112#, 1#, True, True, 46142#
        TEST_Harness.AssertTrue "2026-03-31+1k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-03-31+1k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46112#, 2#, False, True, 46173#
        TEST_Harness.AssertTrue "2026-03-31+2-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-03-31+2-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46112#, 2#, True, True, 46173#
        TEST_Harness.AssertTrue "2026-03-31+2k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-03-31+2k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46112#, 12#, False, True, 46477#
        TEST_Harness.AssertTrue "2026-03-31+12-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-03-31+12-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46112#, 12#, True, True, 46477#
        TEST_Harness.AssertTrue "2026-03-31+12k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-03-31+12k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46112#, 48#, False, True, 47573#
        TEST_Harness.AssertTrue "2026-03-31+48-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-03-31+48-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46112#, 48#, True, True, 47573#
        TEST_Harness.AssertTrue "2026-03-31+48k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-03-31+48k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46387#, -13#, False, True, 45991#
        TEST_Harness.AssertTrue "2026-12-31-13-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-12-31-13-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46387#, -13#, True, True, 45991#
        TEST_Harness.AssertTrue "2026-12-31-13k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-12-31-13k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46387#, -1#, False, True, 46356#
        TEST_Harness.AssertTrue "2026-12-31-1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-12-31-1-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46387#, -1#, True, True, 46356#
        TEST_Harness.AssertTrue "2026-12-31-1k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-12-31-1k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46387#, 0#, False, True, 46387#
        TEST_Harness.AssertTrue "2026-12-31+0-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-12-31+0-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46387#, 0#, True, True, 46387#
        TEST_Harness.AssertTrue "2026-12-31+0k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-12-31+0k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46387#, 1#, False, True, 46418#
        TEST_Harness.AssertTrue "2026-12-31+1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-12-31+1-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46387#, 1#, True, True, 46418#
        TEST_Harness.AssertTrue "2026-12-31+1k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-12-31+1k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46387#, 2#, False, True, 46446#
        TEST_Harness.AssertTrue "2026-12-31+2-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-12-31+2-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46387#, 2#, True, True, 46446#
        TEST_Harness.AssertTrue "2026-12-31+2k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-12-31+2k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46387#, 12#, False, True, 46752#
        TEST_Harness.AssertTrue "2026-12-31+12-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-12-31+12-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46387#, 12#, True, True, 46752#
        TEST_Harness.AssertTrue "2026-12-31+12k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-12-31+12k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46387#, 48#, False, True, 47848#
        TEST_Harness.AssertTrue "2026-12-31+48-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-12-31+48-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46387#, 48#, True, True, 47848#
        TEST_Harness.AssertTrue "2026-12-31+48k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2026-12-31+48k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 36585#, -13#, False, True, 36189#
        TEST_Harness.AssertTrue "2000-02-29-13-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2000-02-29-13-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 36585#, -13#, True, True, 36191#
        TEST_Harness.AssertTrue "2000-02-29-13k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2000-02-29-13k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 36585#, -1#, False, True, 36554#
        TEST_Harness.AssertTrue "2000-02-29-1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2000-02-29-1-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 36585#, -1#, True, True, 36556#
        TEST_Harness.AssertTrue "2000-02-29-1k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2000-02-29-1k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 36585#, 0#, False, True, 36585#
        TEST_Harness.AssertTrue "2000-02-29+0-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2000-02-29+0-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 36585#, 0#, True, True, 36585#
        TEST_Harness.AssertTrue "2000-02-29+0k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2000-02-29+0k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 36585#, 1#, False, True, 36614#
        TEST_Harness.AssertTrue "2000-02-29+1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2000-02-29+1-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 36585#, 1#, True, True, 36616#
        TEST_Harness.AssertTrue "2000-02-29+1k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2000-02-29+1k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 36585#, 2#, False, True, 36645#
        TEST_Harness.AssertTrue "2000-02-29+2-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2000-02-29+2-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 36585#, 2#, True, True, 36646#
        TEST_Harness.AssertTrue "2000-02-29+2k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2000-02-29+2k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 36585#, 12#, False, True, 36950#
        TEST_Harness.AssertTrue "2000-02-29+12-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2000-02-29+12-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 36585#, 12#, True, True, 36950#
        TEST_Harness.AssertTrue "2000-02-29+12k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2000-02-29+12k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 36585#, 48#, False, True, 38046#
        TEST_Harness.AssertTrue "2000-02-29+48-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2000-02-29+48-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 36585#, 48#, True, True, 38046#
        TEST_Harness.AssertTrue "2000-02-29+48k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "2000-02-29+48k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 91#, -13#, False, False, 0#
        TEST_Harness.AssertTrue "1900-03-31-13-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-31-13-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 91#, -13#, True, False, 0#
        TEST_Harness.AssertTrue "1900-03-31-13k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-31-13k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 91#, -1#, False, False, 0#
        TEST_Harness.AssertTrue "1900-03-31-1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-31-1-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 91#, -1#, True, False, 0#
        TEST_Harness.AssertTrue "1900-03-31-1k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-31-1k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 91#, 0#, False, True, 91#
        TEST_Harness.AssertTrue "1900-03-31+0-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-31+0-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 91#, 0#, True, True, 91#
        TEST_Harness.AssertTrue "1900-03-31+0k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-31+0k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 91#, 1#, False, True, 121#
        TEST_Harness.AssertTrue "1900-03-31+1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-31+1-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 91#, 1#, True, True, 121#
        TEST_Harness.AssertTrue "1900-03-31+1k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-31+1k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 91#, 2#, False, True, 152#
        TEST_Harness.AssertTrue "1900-03-31+2-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-31+2-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 91#, 2#, True, True, 152#
        TEST_Harness.AssertTrue "1900-03-31+2k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-31+2k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 91#, 12#, False, True, 456#
        TEST_Harness.AssertTrue "1900-03-31+12-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-31+12-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 91#, 12#, True, True, 456#
        TEST_Harness.AssertTrue "1900-03-31+12k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-31+12k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 91#, 48#, False, True, 1552#
        TEST_Harness.AssertTrue "1900-03-31+48-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-31+48-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 91#, 48#, True, True, 1552#
        TEST_Harness.AssertTrue "1900-03-31+48k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-31+48k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 61#, -13#, False, False, 0#
        TEST_Harness.AssertTrue "1900-03-01-13-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-01-13-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 61#, -13#, True, False, 0#
        TEST_Harness.AssertTrue "1900-03-01-13k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-01-13k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 61#, -1#, False, False, 0#
        TEST_Harness.AssertTrue "1900-03-01-1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-01-1-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 61#, -1#, True, False, 0#
        TEST_Harness.AssertTrue "1900-03-01-1k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-01-1k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 61#, 0#, False, True, 61#
        TEST_Harness.AssertTrue "1900-03-01+0-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-01+0-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 61#, 0#, True, True, 61#
        TEST_Harness.AssertTrue "1900-03-01+0k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-01+0k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 61#, 1#, False, True, 92#
        TEST_Harness.AssertTrue "1900-03-01+1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-01+1-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 61#, 1#, True, True, 92#
        TEST_Harness.AssertTrue "1900-03-01+1k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-01+1k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 61#, 2#, False, True, 122#
        TEST_Harness.AssertTrue "1900-03-01+2-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-01+2-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 61#, 2#, True, True, 122#
        TEST_Harness.AssertTrue "1900-03-01+2k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-01+2k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 61#, 12#, False, True, 426#
        TEST_Harness.AssertTrue "1900-03-01+12-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-01+12-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 61#, 12#, True, True, 426#
        TEST_Harness.AssertTrue "1900-03-01+12k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-01+12k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 61#, 48#, False, True, 1522#
        TEST_Harness.AssertTrue "1900-03-01+48-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-01+48-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 61#, 48#, True, True, 1522#
        TEST_Harness.AssertTrue "1900-03-01+48k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "1900-03-01+48k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958434#, -13#, False, True, 2958038#
        TEST_Harness.AssertTrue "9999-11-30-13-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-11-30-13-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958434#, -13#, True, True, 2958039#
        TEST_Harness.AssertTrue "9999-11-30-13k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-11-30-13k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958434#, -1#, False, True, 2958403#
        TEST_Harness.AssertTrue "9999-11-30-1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-11-30-1-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958434#, -1#, True, True, 2958404#
        TEST_Harness.AssertTrue "9999-11-30-1k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-11-30-1k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958434#, 0#, False, True, 2958434#
        TEST_Harness.AssertTrue "9999-11-30+0-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-11-30+0-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958434#, 0#, True, True, 2958434#
        TEST_Harness.AssertTrue "9999-11-30+0k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-11-30+0k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958434#, 1#, False, True, 2958464#
        TEST_Harness.AssertTrue "9999-11-30+1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-11-30+1-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958434#, 1#, True, True, 2958465#
        TEST_Harness.AssertTrue "9999-11-30+1k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-11-30+1k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958434#, 2#, False, False, 0#
        TEST_Harness.AssertTrue "9999-11-30+2-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-11-30+2-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958434#, 2#, True, False, 0#
        TEST_Harness.AssertTrue "9999-11-30+2k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-11-30+2k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958434#, 12#, False, False, 0#
        TEST_Harness.AssertTrue "9999-11-30+12-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-11-30+12-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958434#, 12#, True, False, 0#
        TEST_Harness.AssertTrue "9999-11-30+12k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-11-30+12k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958434#, 48#, False, False, 0#
        TEST_Harness.AssertTrue "9999-11-30+48-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-11-30+48-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958434#, 48#, True, False, 0#
        TEST_Harness.AssertTrue "9999-11-30+48k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-11-30+48k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958465#, -13#, False, True, 2958069#
        TEST_Harness.AssertTrue "9999-12-31-13-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-12-31-13-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958465#, -13#, True, True, 2958069#
        TEST_Harness.AssertTrue "9999-12-31-13k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-12-31-13k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958465#, -1#, False, True, 2958434#
        TEST_Harness.AssertTrue "9999-12-31-1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-12-31-1-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958465#, -1#, True, True, 2958434#
        TEST_Harness.AssertTrue "9999-12-31-1k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-12-31-1k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958465#, 0#, False, True, 2958465#
        TEST_Harness.AssertTrue "9999-12-31+0-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-12-31+0-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958465#, 0#, True, True, 2958465#
        TEST_Harness.AssertTrue "9999-12-31+0k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-12-31+0k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958465#, 1#, False, False, 0#
        TEST_Harness.AssertTrue "9999-12-31+1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-12-31+1-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958465#, 1#, True, False, 0#
        TEST_Harness.AssertTrue "9999-12-31+1k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-12-31+1k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958465#, 2#, False, False, 0#
        TEST_Harness.AssertTrue "9999-12-31+2-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-12-31+2-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958465#, 2#, True, False, 0#
        TEST_Harness.AssertTrue "9999-12-31+2k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-12-31+2k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958465#, 12#, False, False, 0#
        TEST_Harness.AssertTrue "9999-12-31+12-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-12-31+12-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958465#, 12#, True, False, 0#
        TEST_Harness.AssertTrue "9999-12-31+12k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-12-31+12k-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958465#, 48#, False, False, 0#
        TEST_Harness.AssertTrue "9999-12-31+48-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-12-31+48-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 2958465#, 48#, True, False, 0#
        TEST_Harness.AssertTrue "9999-12-31+48k-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "9999-12-31+48k-value", mFact(2, 1), mFact(2, 2)
    'An extreme count is outside the window, not an overflow
        ProbeAddMonths 46188#, (10# ^ 15), False, False, 0#
        TEST_Harness.AssertTrue "extreme-1e+15-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "extreme-1e+15-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46188#, -(10# ^ 15), False, False, 0#
        TEST_Harness.AssertTrue "extreme--1e+15-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "extreme--1e+15-value", mFact(2, 1), mFact(2, 2)
        ProbeAddMonths 46188#, (10# ^ 300), False, False, 0#
        TEST_Harness.AssertTrue "extreme-1e+300-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualDouble "extreme-1e+300-value", mFact(2, 1), mFact(2, 2)

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

Private Sub CaseNthWeekday()

'
'------------------------------------------------------------------------------
'                            CHECK THE NTH WEEKDAY
'------------------------------------------------------------------------------
' PURPOSE
'   The nth weekday of a month is found under both bases, or reported absent or out of the window. 240 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date core follows the calculation contracts
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Locates the first and fifth Monday-based and Sunday-based weekday of months in 1900, 2026 and 9999, and fails each domain
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Dates, CORE_Err, TEST_Harness
'
' NOTES
'   Expected values come from tests/ reference computations in Python, not from the code
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "nth-weekday"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'The nth weekday of a month, with a fifth that may be absent and a result that may be before the window
        ProbeNthWeekday 2026, 1, 1, 1, True, cdNone, 46027#
        TEST_Harness.AssertTrue "2026-1-w1-n1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-1-w1-n1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-1-w1-n1M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 1, 1, 1, False, cdNone, 46026#
        TEST_Harness.AssertTrue "2026-1-w1-n1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-1-w1-n1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-1-w1-n1S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 1, 1, 5, True, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "2026-1-w1-n5M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-1-w1-n5M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-1-w1-n5M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 1, 1, 5, False, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "2026-1-w1-n5S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-1-w1-n5S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-1-w1-n5S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 1, 7, 1, True, cdNone, 46026#
        TEST_Harness.AssertTrue "2026-1-w7-n1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-1-w7-n1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-1-w7-n1M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 1, 7, 1, False, cdNone, 46025#
        TEST_Harness.AssertTrue "2026-1-w7-n1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-1-w7-n1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-1-w7-n1S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 1, 7, 5, True, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "2026-1-w7-n5M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-1-w7-n5M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-1-w7-n5M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 1, 7, 5, False, cdNone, 46053#
        TEST_Harness.AssertTrue "2026-1-w7-n5S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-1-w7-n5S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-1-w7-n5S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 2, 1, 1, True, cdNone, 46055#
        TEST_Harness.AssertTrue "2026-2-w1-n1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-2-w1-n1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-2-w1-n1M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 2, 1, 1, False, cdNone, 46054#
        TEST_Harness.AssertTrue "2026-2-w1-n1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-2-w1-n1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-2-w1-n1S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 2, 1, 5, True, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "2026-2-w1-n5M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-2-w1-n5M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-2-w1-n5M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 2, 1, 5, False, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "2026-2-w1-n5S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-2-w1-n5S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-2-w1-n5S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 2, 7, 1, True, cdNone, 46054#
        TEST_Harness.AssertTrue "2026-2-w7-n1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-2-w7-n1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-2-w7-n1M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 2, 7, 1, False, cdNone, 46060#
        TEST_Harness.AssertTrue "2026-2-w7-n1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-2-w7-n1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-2-w7-n1S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 2, 7, 5, True, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "2026-2-w7-n5M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-2-w7-n5M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-2-w7-n5M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 2, 7, 5, False, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "2026-2-w7-n5S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-2-w7-n5S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-2-w7-n5S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 12, 1, 1, True, cdNone, 46363#
        TEST_Harness.AssertTrue "2026-12-w1-n1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-12-w1-n1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-12-w1-n1M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 12, 1, 1, False, cdNone, 46362#
        TEST_Harness.AssertTrue "2026-12-w1-n1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-12-w1-n1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-12-w1-n1S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 12, 1, 5, True, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "2026-12-w1-n5M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-12-w1-n5M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-12-w1-n5M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 12, 1, 5, False, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "2026-12-w1-n5S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-12-w1-n5S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-12-w1-n5S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 12, 7, 1, True, cdNone, 46362#
        TEST_Harness.AssertTrue "2026-12-w7-n1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-12-w7-n1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-12-w7-n1M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 12, 7, 1, False, cdNone, 46361#
        TEST_Harness.AssertTrue "2026-12-w7-n1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-12-w7-n1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-12-w7-n1S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 12, 7, 5, True, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "2026-12-w7-n5M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-12-w7-n5M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-12-w7-n5M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 12, 7, 5, False, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "2026-12-w7-n5S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-12-w7-n5S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-12-w7-n5S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 1, 1, 1, True, cdResultWindow, 0#
        TEST_Harness.AssertTrue "1900-1-w1-n1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-1-w1-n1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-1-w1-n1M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 1, 1, 1, False, cdResultWindow, 0#
        TEST_Harness.AssertTrue "1900-1-w1-n1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-1-w1-n1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-1-w1-n1S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 1, 1, 5, True, cdResultWindow, 0#
        TEST_Harness.AssertTrue "1900-1-w1-n5M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-1-w1-n5M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-1-w1-n5M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 1, 1, 5, False, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "1900-1-w1-n5S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-1-w1-n5S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-1-w1-n5S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 1, 7, 1, True, cdResultWindow, 0#
        TEST_Harness.AssertTrue "1900-1-w7-n1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-1-w7-n1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-1-w7-n1M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 1, 7, 1, False, cdResultWindow, 0#
        TEST_Harness.AssertTrue "1900-1-w7-n1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-1-w7-n1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-1-w7-n1S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 1, 7, 5, True, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "1900-1-w7-n5M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-1-w7-n5M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-1-w7-n5M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 1, 7, 5, False, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "1900-1-w7-n5S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-1-w7-n5S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-1-w7-n5S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 2, 1, 1, True, cdResultWindow, 0#
        TEST_Harness.AssertTrue "1900-2-w1-n1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-2-w1-n1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-2-w1-n1M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 2, 1, 1, False, cdResultWindow, 0#
        TEST_Harness.AssertTrue "1900-2-w1-n1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-2-w1-n1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-2-w1-n1S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 2, 1, 5, True, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "1900-2-w1-n5M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-2-w1-n5M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-2-w1-n5M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 2, 1, 5, False, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "1900-2-w1-n5S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-2-w1-n5S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-2-w1-n5S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 2, 7, 1, True, cdResultWindow, 0#
        TEST_Harness.AssertTrue "1900-2-w7-n1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-2-w7-n1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-2-w7-n1M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 2, 7, 1, False, cdResultWindow, 0#
        TEST_Harness.AssertTrue "1900-2-w7-n1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-2-w7-n1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-2-w7-n1S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 2, 7, 5, True, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "1900-2-w7-n5M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-2-w7-n5M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-2-w7-n5M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 2, 7, 5, False, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "1900-2-w7-n5S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-2-w7-n5S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-2-w7-n5S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 12, 1, 1, True, cdNone, 338#
        TEST_Harness.AssertTrue "1900-12-w1-n1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-12-w1-n1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-12-w1-n1M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 12, 1, 1, False, cdNone, 337#
        TEST_Harness.AssertTrue "1900-12-w1-n1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-12-w1-n1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-12-w1-n1S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 12, 1, 5, True, cdNone, 366#
        TEST_Harness.AssertTrue "1900-12-w1-n5M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-12-w1-n5M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-12-w1-n5M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 12, 1, 5, False, cdNone, 365#
        TEST_Harness.AssertTrue "1900-12-w1-n5S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-12-w1-n5S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-12-w1-n5S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 12, 7, 1, True, cdNone, 337#
        TEST_Harness.AssertTrue "1900-12-w7-n1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-12-w7-n1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-12-w7-n1M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 12, 7, 1, False, cdNone, 336#
        TEST_Harness.AssertTrue "1900-12-w7-n1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-12-w7-n1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-12-w7-n1S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 12, 7, 5, True, cdNone, 365#
        TEST_Harness.AssertTrue "1900-12-w7-n5M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-12-w7-n5M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-12-w7-n5M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1900, 12, 7, 5, False, cdNone, 364#
        TEST_Harness.AssertTrue "1900-12-w7-n5S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-12-w7-n5S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-12-w7-n5S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 1, 1, 1, True, cdNone, 2958104#
        TEST_Harness.AssertTrue "9999-1-w1-n1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-1-w1-n1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-1-w1-n1M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 1, 1, 1, False, cdNone, 2958103#
        TEST_Harness.AssertTrue "9999-1-w1-n1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-1-w1-n1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-1-w1-n1S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 1, 1, 5, True, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "9999-1-w1-n5M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-1-w1-n5M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-1-w1-n5M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 1, 1, 5, False, cdNone, 2958131#
        TEST_Harness.AssertTrue "9999-1-w1-n5S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-1-w1-n5S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-1-w1-n5S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 1, 7, 1, True, cdNone, 2958103#
        TEST_Harness.AssertTrue "9999-1-w7-n1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-1-w7-n1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-1-w7-n1M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 1, 7, 1, False, cdNone, 2958102#
        TEST_Harness.AssertTrue "9999-1-w7-n1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-1-w7-n1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-1-w7-n1S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 1, 7, 5, True, cdNone, 2958131#
        TEST_Harness.AssertTrue "9999-1-w7-n5M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-1-w7-n5M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-1-w7-n5M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 1, 7, 5, False, cdNone, 2958130#
        TEST_Harness.AssertTrue "9999-1-w7-n5S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-1-w7-n5S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-1-w7-n5S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 2, 1, 1, True, cdNone, 2958132#
        TEST_Harness.AssertTrue "9999-2-w1-n1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-2-w1-n1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-2-w1-n1M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 2, 1, 1, False, cdNone, 2958138#
        TEST_Harness.AssertTrue "9999-2-w1-n1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-2-w1-n1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-2-w1-n1S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 2, 1, 5, True, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "9999-2-w1-n5M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-2-w1-n5M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-2-w1-n5M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 2, 1, 5, False, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "9999-2-w1-n5S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-2-w1-n5S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-2-w1-n5S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 2, 7, 1, True, cdNone, 2958138#
        TEST_Harness.AssertTrue "9999-2-w7-n1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-2-w7-n1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-2-w7-n1M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 2, 7, 1, False, cdNone, 2958137#
        TEST_Harness.AssertTrue "9999-2-w7-n1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-2-w7-n1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-2-w7-n1S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 2, 7, 5, True, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "9999-2-w7-n5M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-2-w7-n5M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-2-w7-n5M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 2, 7, 5, False, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "9999-2-w7-n5S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-2-w7-n5S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-2-w7-n5S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 12, 1, 1, True, cdNone, 2958440#
        TEST_Harness.AssertTrue "9999-12-w1-n1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-w1-n1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-12-w1-n1M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 12, 1, 1, False, cdNone, 2958439#
        TEST_Harness.AssertTrue "9999-12-w1-n1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-w1-n1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-12-w1-n1S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 12, 1, 5, True, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "9999-12-w1-n5M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-w1-n5M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-12-w1-n5M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 12, 1, 5, False, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "9999-12-w1-n5S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-w1-n5S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-12-w1-n5S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 12, 7, 1, True, cdNone, 2958439#
        TEST_Harness.AssertTrue "9999-12-w7-n1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-w7-n1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-12-w7-n1M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 12, 7, 1, False, cdNone, 2958438#
        TEST_Harness.AssertTrue "9999-12-w7-n1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-w7-n1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-12-w7-n1S-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 12, 7, 5, True, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "9999-12-w7-n5M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-w7-n5M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-12-w7-n5M-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 9999, 12, 7, 5, False, cdOccurrenceAbsent, 0#
        TEST_Harness.AssertTrue "9999-12-w7-n5S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-w7-n5S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-12-w7-n5S-value", mFact(3, 1), mFact(3, 2)
    'Domain failures, in the order of the arguments
        ProbeNthWeekday 1899, 1, 1, 1, True, cdDomainYear, 0#
        TEST_Harness.AssertTrue "domain-1899-1-1-1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "domain-1899-1-1-1-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "domain-1899-1-1-1-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 0, 1, 1, True, cdDomainMonth, 0#
        TEST_Harness.AssertTrue "domain-2026-0-1-1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "domain-2026-0-1-1-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "domain-2026-0-1-1-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 1, 0, 1, True, cdDomainWeekday, 0#
        TEST_Harness.AssertTrue "domain-2026-1-0-1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "domain-2026-1-0-1-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "domain-2026-1-0-1-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 1, 1, 0, True, cdDomainOccurrence, 0#
        TEST_Harness.AssertTrue "domain-2026-1-1-0-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "domain-2026-1-1-0-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "domain-2026-1-1-0-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 1899, 13, 8, 6, True, cdDomainYear, 0#
        TEST_Harness.AssertTrue "domain-1899-13-8-6-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "domain-1899-13-8-6-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "domain-1899-13-8-6-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 13, 8, 6, True, cdDomainMonth, 0#
        TEST_Harness.AssertTrue "domain-2026-13-8-6-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "domain-2026-13-8-6-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "domain-2026-13-8-6-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 2026, 1, 8, 6, True, cdDomainWeekday, 0#
        TEST_Harness.AssertTrue "domain-2026-1-8-6-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "domain-2026-1-8-6-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "domain-2026-1-8-6-value", mFact(3, 1), mFact(3, 2)
        ProbeNthWeekday 10000, 1, 1, 1, True, cdDomainYear, 0#
        TEST_Harness.AssertTrue "domain-10000-1-1-1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "domain-10000-1-1-1-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "domain-10000-1-1-1-value", mFact(3, 1), mFact(3, 2)

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

Private Sub CaseLastWeekday()

'
'------------------------------------------------------------------------------
'                            CHECK THE LAST WEEKDAY
'------------------------------------------------------------------------------
' PURPOSE
'   The last weekday of a month is found under both bases, or reported out of the window. 123 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date core follows the calculation contracts
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Locates the last Monday-based and Sunday-based weekday of months in 1900, 2026 and 9999, and fails each domain
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Dates, CORE_Err, TEST_Harness
'
' NOTES
'   Expected values come from tests/ reference computations in Python, not from the code
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "last-weekday"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'The last weekday of a month, with a result that may be before the window
        ProbeLastWeekday 2026, 1, 1, True, cdNone, 46048#
        TEST_Harness.AssertTrue "2026-1-w1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-1-w1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-1-w1M-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 2026, 1, 1, False, cdNone, 46047#
        TEST_Harness.AssertTrue "2026-1-w1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-1-w1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-1-w1S-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 2026, 1, 7, True, cdNone, 46047#
        TEST_Harness.AssertTrue "2026-1-w7M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-1-w7M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-1-w7M-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 2026, 1, 7, False, cdNone, 46053#
        TEST_Harness.AssertTrue "2026-1-w7S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-1-w7S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-1-w7S-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 2026, 2, 1, True, cdNone, 46076#
        TEST_Harness.AssertTrue "2026-2-w1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-2-w1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-2-w1M-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 2026, 2, 1, False, cdNone, 46075#
        TEST_Harness.AssertTrue "2026-2-w1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-2-w1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-2-w1S-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 2026, 2, 7, True, cdNone, 46075#
        TEST_Harness.AssertTrue "2026-2-w7M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-2-w7M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-2-w7M-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 2026, 2, 7, False, cdNone, 46081#
        TEST_Harness.AssertTrue "2026-2-w7S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-2-w7S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-2-w7S-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 2026, 12, 1, True, cdNone, 46384#
        TEST_Harness.AssertTrue "2026-12-w1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-12-w1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-12-w1M-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 2026, 12, 1, False, cdNone, 46383#
        TEST_Harness.AssertTrue "2026-12-w1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-12-w1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-12-w1S-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 2026, 12, 7, True, cdNone, 46383#
        TEST_Harness.AssertTrue "2026-12-w7M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-12-w7M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-12-w7M-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 2026, 12, 7, False, cdNone, 46382#
        TEST_Harness.AssertTrue "2026-12-w7S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-12-w7S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-12-w7S-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 1900, 1, 1, True, cdResultWindow, 0#
        TEST_Harness.AssertTrue "1900-1-w1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-1-w1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-1-w1M-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 1900, 1, 1, False, cdResultWindow, 0#
        TEST_Harness.AssertTrue "1900-1-w1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-1-w1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-1-w1S-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 1900, 1, 7, True, cdResultWindow, 0#
        TEST_Harness.AssertTrue "1900-1-w7M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-1-w7M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-1-w7M-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 1900, 1, 7, False, cdResultWindow, 0#
        TEST_Harness.AssertTrue "1900-1-w7S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-1-w7S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-1-w7S-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 1900, 2, 1, True, cdResultWindow, 0#
        TEST_Harness.AssertTrue "1900-2-w1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-2-w1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-2-w1M-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 1900, 2, 1, False, cdResultWindow, 0#
        TEST_Harness.AssertTrue "1900-2-w1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-2-w1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-2-w1S-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 1900, 2, 7, True, cdResultWindow, 0#
        TEST_Harness.AssertTrue "1900-2-w7M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-2-w7M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-2-w7M-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 1900, 2, 7, False, cdResultWindow, 0#
        TEST_Harness.AssertTrue "1900-2-w7S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-2-w7S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-2-w7S-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 1900, 12, 1, True, cdNone, 366#
        TEST_Harness.AssertTrue "1900-12-w1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-12-w1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-12-w1M-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 1900, 12, 1, False, cdNone, 365#
        TEST_Harness.AssertTrue "1900-12-w1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-12-w1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-12-w1S-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 1900, 12, 7, True, cdNone, 365#
        TEST_Harness.AssertTrue "1900-12-w7M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-12-w7M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-12-w7M-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 1900, 12, 7, False, cdNone, 364#
        TEST_Harness.AssertTrue "1900-12-w7S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-12-w7S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-12-w7S-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 9999, 1, 1, True, cdNone, 2958125#
        TEST_Harness.AssertTrue "9999-1-w1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-1-w1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-1-w1M-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 9999, 1, 1, False, cdNone, 2958131#
        TEST_Harness.AssertTrue "9999-1-w1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-1-w1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-1-w1S-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 9999, 1, 7, True, cdNone, 2958131#
        TEST_Harness.AssertTrue "9999-1-w7M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-1-w7M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-1-w7M-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 9999, 1, 7, False, cdNone, 2958130#
        TEST_Harness.AssertTrue "9999-1-w7S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-1-w7S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-1-w7S-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 9999, 2, 1, True, cdNone, 2958153#
        TEST_Harness.AssertTrue "9999-2-w1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-2-w1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-2-w1M-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 9999, 2, 1, False, cdNone, 2958159#
        TEST_Harness.AssertTrue "9999-2-w1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-2-w1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-2-w1S-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 9999, 2, 7, True, cdNone, 2958159#
        TEST_Harness.AssertTrue "9999-2-w7M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-2-w7M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-2-w7M-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 9999, 2, 7, False, cdNone, 2958158#
        TEST_Harness.AssertTrue "9999-2-w7S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-2-w7S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-2-w7S-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 9999, 12, 1, True, cdNone, 2958461#
        TEST_Harness.AssertTrue "9999-12-w1M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-w1M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-12-w1M-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 9999, 12, 1, False, cdNone, 2958460#
        TEST_Harness.AssertTrue "9999-12-w1S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-w1S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-12-w1S-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 9999, 12, 7, True, cdNone, 2958460#
        TEST_Harness.AssertTrue "9999-12-w7M-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-w7M-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-12-w7M-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 9999, 12, 7, False, cdNone, 2958459#
        TEST_Harness.AssertTrue "9999-12-w7S-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-w7S-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-12-w7S-value", mFact(3, 1), mFact(3, 2)
    'Domain failures, in the order of the arguments
        ProbeLastWeekday 1899, 1, 1, True, cdDomainYear, 0#
        TEST_Harness.AssertTrue "domain-1899-1-1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "domain-1899-1-1-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "domain-1899-1-1-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 2026, 0, 1, True, cdDomainMonth, 0#
        TEST_Harness.AssertTrue "domain-2026-0-1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "domain-2026-0-1-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "domain-2026-0-1-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 2026, 1, 0, True, cdDomainWeekday, 0#
        TEST_Harness.AssertTrue "domain-2026-1-0-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "domain-2026-1-0-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "domain-2026-1-0-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 1899, 13, 8, True, cdDomainYear, 0#
        TEST_Harness.AssertTrue "domain-1899-13-8-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "domain-1899-13-8-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "domain-1899-13-8-value", mFact(3, 1), mFact(3, 2)
        ProbeLastWeekday 2026, 13, 8, True, cdDomainMonth, 0#
        TEST_Harness.AssertTrue "domain-2026-13-8-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "domain-2026-13-8-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "domain-2026-13-8-value", mFact(3, 1), mFact(3, 2)

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

Private Sub CaseParsePillar()

'
'------------------------------------------------------------------------------
'                           CHECK THE PILLAR PARSER
'------------------------------------------------------------------------------
' PURPOSE
'   A pillar token gives its month and day deltas, or the condition that rejects it. 176 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date core follows the calculation contracts
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Parses accepted tokens of every unit and alias and each rejected form
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Dates, CORE_Err, TEST_Harness
'
' NOTES
'   Expected values come from tests/ reference computations in Python, not from the code
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "parse-pillar"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Pillar tokens: every unit, the aliases, case, blanks, and each rejection
        ProbeParsePillar "3M", 3#, 0#, cdNone
        TEST_Harness.AssertTrue "3M-cdNone-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "3M-cdNone-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "3M-cdNone-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "3M-cdNone-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "1Y6M", 18#, 0#, cdNone
        TEST_Harness.AssertTrue "1Y6M-cdNone-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1Y6M-cdNone-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1Y6M-cdNone-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "1Y6M-cdNone-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "-1Y", -12#, 0#, cdNone
        TEST_Harness.AssertTrue "minus1Y-cdNone-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "minus1Y-cdNone-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "minus1Y-cdNone-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "minus1Y-cdNone-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "+1M", 1#, 0#, cdNone
        TEST_Harness.AssertTrue "plus1M-cdNone-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "plus1M-cdNone-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "plus1M-cdNone-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "plus1M-cdNone-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "2W3D", 0#, 17#, cdNone
        TEST_Harness.AssertTrue "2W3D-cdNone-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2W3D-cdNone-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2W3D-cdNone-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "2W3D-cdNone-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "3D2W", 0#, 17#, cdNone
        TEST_Harness.AssertTrue "3D2W-cdNone-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "3D2W-cdNone-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "3D2W-cdNone-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "3D2W-cdNone-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "-2W3D", 0#, -17#, cdNone
        TEST_Harness.AssertTrue "minus2W3D-cdNone-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "minus2W3D-cdNone-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "minus2W3D-cdNone-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "minus2W3D-cdNone-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "1Y2M3W4D", 14#, 25#, cdNone
        TEST_Harness.AssertTrue "1Y2M3W4D-cdNone-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1Y2M3W4D-cdNone-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1Y2M3W4D-cdNone-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "1Y2M3W4D-cdNone-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "10D", 0#, 10#, cdNone
        TEST_Harness.AssertTrue "10D-cdNone-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "10D-cdNone-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "10D-cdNone-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "10D-cdNone-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "0D", 0#, 0#, cdNone
        TEST_Harness.AssertTrue "0D-cdNone-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "0D-cdNone-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "0D-cdNone-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "0D-cdNone-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "0M", 0#, 0#, cdNone
        TEST_Harness.AssertTrue "0M-cdNone-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "0M-cdNone-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "0M-cdNone-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "0M-cdNone-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "ON", 0#, 1#, cdNone
        TEST_Harness.AssertTrue "ON-cdNone-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "ON-cdNone-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "ON-cdNone-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "ON-cdNone-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "O/N", 0#, 1#, cdNone
        TEST_Harness.AssertTrue "OminusN-cdNone-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "OminusN-cdNone-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "OminusN-cdNone-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "OminusN-cdNone-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "TN", 0#, 2#, cdNone
        TEST_Harness.AssertTrue "TN-cdNone-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "TN-cdNone-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "TN-cdNone-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "TN-cdNone-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "T/N", 0#, 2#, cdNone
        TEST_Harness.AssertTrue "TminusN-cdNone-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "TminusN-cdNone-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "TminusN-cdNone-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "TminusN-cdNone-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "on", 0#, 1#, cdNone
        TEST_Harness.AssertTrue "on-cdNone-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "on-cdNone-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "on-cdNone-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "on-cdNone-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "t/n", 0#, 2#, cdNone
        TEST_Harness.AssertTrue "tminusn-cdNone-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "tminusn-cdNone-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "tminusn-cdNone-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "tminusn-cdNone-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "1m", 1#, 0#, cdNone
        TEST_Harness.AssertTrue "1m-cdNone-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1m-cdNone-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1m-cdNone-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "1m-cdNone-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "  1M  ", 1#, 0#, cdNone
        TEST_Harness.AssertTrue "__1M__-cdNone-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "__1M__-cdNone-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "__1M__-cdNone-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "__1M__-cdNone-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "" & vbTab & "2Y" & vbTab & "", 24#, 0#, cdNone
        TEST_Harness.AssertTrue "tab2Ytab-cdNone-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "tab2Ytab-cdNone-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "tab2Ytab-cdNone-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "tab2Ytab-cdNone-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "999999999999M", 999999999999#, 0#, cdNone
        TEST_Harness.AssertTrue "999999999999M-cdNone-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "999999999999M-cdNone-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "999999999999M-cdNone-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "999999999999M-cdNone-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "1 M", 0#, 0#, cdPillarTokenMalformed
        TEST_Harness.AssertTrue "1_M-cdPillarTokenMalformed-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1_M-cdPillarTokenMalformed-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1_M-cdPillarTokenMalformed-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "1_M-cdPillarTokenMalformed-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "1M 2W", 0#, 0#, cdPillarTokenMalformed
        TEST_Harness.AssertTrue "1M_2W-cdPillarTokenMalformed-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1M_2W-cdPillarTokenMalformed-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1M_2W-cdPillarTokenMalformed-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "1M_2W-cdPillarTokenMalformed-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "M", 0#, 0#, cdPillarTokenMalformed
        TEST_Harness.AssertTrue "M-cdPillarTokenMalformed-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "M-cdPillarTokenMalformed-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "M-cdPillarTokenMalformed-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "M-cdPillarTokenMalformed-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "1.5Y", 0#, 0#, cdPillarTokenMalformed
        TEST_Harness.AssertTrue "1.5Y-cdPillarTokenMalformed-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1.5Y-cdPillarTokenMalformed-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1.5Y-cdPillarTokenMalformed-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "1.5Y-cdPillarTokenMalformed-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "1X", 0#, 0#, cdPillarTokenMalformed
        TEST_Harness.AssertTrue "1X-cdPillarTokenMalformed-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1X-cdPillarTokenMalformed-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1X-cdPillarTokenMalformed-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "1X-cdPillarTokenMalformed-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "12", 0#, 0#, cdPillarTokenMalformed
        TEST_Harness.AssertTrue "12-cdPillarTokenMalformed-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "12-cdPillarTokenMalformed-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "12-cdPillarTokenMalformed-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "12-cdPillarTokenMalformed-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "", 0#, 0#, cdPillarTokenMalformed
        TEST_Harness.AssertTrue "empty-cdPillarTokenMalformed-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "empty-cdPillarTokenMalformed-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "empty-cdPillarTokenMalformed-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "empty-cdPillarTokenMalformed-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "   ", 0#, 0#, cdPillarTokenMalformed
        TEST_Harness.AssertTrue "___-cdPillarTokenMalformed-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "___-cdPillarTokenMalformed-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "___-cdPillarTokenMalformed-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "___-cdPillarTokenMalformed-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "+", 0#, 0#, cdPillarTokenMalformed
        TEST_Harness.AssertTrue "plus-cdPillarTokenMalformed-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "plus-cdPillarTokenMalformed-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "plus-cdPillarTokenMalformed-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "plus-cdPillarTokenMalformed-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "-", 0#, 0#, cdPillarTokenMalformed
        TEST_Harness.AssertTrue "minus-cdPillarTokenMalformed-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "minus-cdPillarTokenMalformed-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "minus-cdPillarTokenMalformed-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "minus-cdPillarTokenMalformed-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "1M2", 0#, 0#, cdPillarTokenMalformed
        TEST_Harness.AssertTrue "1M2-cdPillarTokenMalformed-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1M2-cdPillarTokenMalformed-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1M2-cdPillarTokenMalformed-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "1M2-cdPillarTokenMalformed-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "ONE", 0#, 0#, cdPillarTokenMalformed
        TEST_Harness.AssertTrue "ONE-cdPillarTokenMalformed-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "ONE-cdPillarTokenMalformed-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "ONE-cdPillarTokenMalformed-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "ONE-cdPillarTokenMalformed-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "O/N1D", 0#, 0#, cdPillarTokenMalformed
        TEST_Harness.AssertTrue "OminusN1D-cdPillarTokenMalformed-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "OminusN1D-cdPillarTokenMalformed-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "OminusN1D-cdPillarTokenMalformed-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "OminusN1D-cdPillarTokenMalformed-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "1M2M", 0#, 0#, cdPillarDuplicateUnit
        TEST_Harness.AssertTrue "1M2M-cdPillarDuplicateUnit-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1M2M-cdPillarDuplicateUnit-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1M2M-cdPillarDuplicateUnit-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "1M2M-cdPillarDuplicateUnit-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "3D4D", 0#, 0#, cdPillarDuplicateUnit
        TEST_Harness.AssertTrue "3D4D-cdPillarDuplicateUnit-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "3D4D-cdPillarDuplicateUnit-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "3D4D-cdPillarDuplicateUnit-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "3D4D-cdPillarDuplicateUnit-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "1Y1Y", 0#, 0#, cdPillarDuplicateUnit
        TEST_Harness.AssertTrue "1Y1Y-cdPillarDuplicateUnit-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1Y1Y-cdPillarDuplicateUnit-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1Y1Y-cdPillarDuplicateUnit-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "1Y1Y-cdPillarDuplicateUnit-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "-ON", 0#, 0#, cdPillarAliasSigned
        TEST_Harness.AssertTrue "minusON-cdPillarAliasSigned-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "minusON-cdPillarAliasSigned-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "minusON-cdPillarAliasSigned-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "minusON-cdPillarAliasSigned-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "+T/N", 0#, 0#, cdPillarAliasSigned
        TEST_Harness.AssertTrue "plusTminusN-cdPillarAliasSigned-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "plusTminusN-cdPillarAliasSigned-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "plusTminusN-cdPillarAliasSigned-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "plusTminusN-cdPillarAliasSigned-days", mFact(4, 1), mFact(4, 2)
    'Not text, and a quantity that overflows a Double
        ProbeParsePillar 3, 0#, 0#, cdPillarTypeRejected
        TEST_Harness.AssertTrue "type-number-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "type-number-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "type-number-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "type-number-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar Empty, 0#, 0#, cdPillarTypeRejected
        TEST_Harness.AssertTrue "type-empty-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "type-empty-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "type-empty-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "type-empty-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar Null, 0#, 0#, cdPillarTypeRejected
        TEST_Harness.AssertTrue "type-null-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "type-null-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "type-null-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "type-null-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar True, 0#, 0#, cdPillarTypeRejected
        TEST_Harness.AssertTrue "type-boolean-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "type-boolean-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "type-boolean-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "type-boolean-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar String(310, "9") & "D", 0#, 0#, cdPillarAggregateRange
        TEST_Harness.AssertTrue "type-overflow-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "type-overflow-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "type-overflow-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "type-overflow-days", mFact(4, 1), mFact(4, 2)

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

Private Sub CaseAddPillar()

'
'------------------------------------------------------------------------------
'                            CHECK THE PILLAR SHIFT
'------------------------------------------------------------------------------
' PURPOSE
'   A parsed pillar moves a date by months first, then days, inside the window. 45 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date core follows the calculation contracts
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Shifts month ends and leap days by pillar deltas, to and past both edges of the window
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Dates, CORE_Err, TEST_Harness
'
' NOTES
'   Expected values come from tests/ reference computations in Python, not from the code
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "add-pillar"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Months first with the clip, then days; a shift out of the window is an aggregate range
        ProbeAddPillar 46053#, 1#, 0#, cdNone, 46081#
        TEST_Harness.AssertTrue "2026-01-31_1_0-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_1_0-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-01-31_1_0-value", mFact(3, 1), mFact(3, 2)
        ProbeAddPillar 46053#, 0#, 17#, cdNone, 46070#
        TEST_Harness.AssertTrue "2026-01-31_0_17-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_0_17-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-01-31_0_17-value", mFact(3, 1), mFact(3, 2)
        ProbeAddPillar 46112#, -1#, 0#, cdNone, 46081#
        TEST_Harness.AssertTrue "2026-03-31_-1_0-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-03-31_-1_0-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-03-31_-1_0-value", mFact(3, 1), mFact(3, 2)
        ProbeAddPillar 45351#, 12#, 0#, cdNone, 45716#
        TEST_Harness.AssertTrue "2024-02-29_12_0-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_12_0-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2024-02-29_12_0-value", mFact(3, 1), mFact(3, 2)
        ProbeAddPillar 45351#, 48#, 0#, cdNone, 46812#
        TEST_Harness.AssertTrue "2024-02-29_48_0-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_48_0-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2024-02-29_48_0-value", mFact(3, 1), mFact(3, 2)
        ProbeAddPillar 46037#, -12#, -3#, cdNone, 45669#
        TEST_Harness.AssertTrue "2026-01-15_-12_-3-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-15_-12_-3-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-01-15_-12_-3-value", mFact(3, 1), mFact(3, 2)
        ProbeAddPillar 46188#, 18#, 25#, cdNone, 46761#
        TEST_Harness.AssertTrue "2026-06-15_18_25-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-06-15_18_25-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-06-15_18_25-value", mFact(3, 1), mFact(3, 2)
        ProbeAddPillar 61#, -1#, 0#, cdPillarAggregateRange, 0#
        TEST_Harness.AssertTrue "1900-03-01_-1_0-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-03-01_-1_0-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-03-01_-1_0-value", mFact(3, 1), mFact(3, 2)
        ProbeAddPillar 61#, 0#, -1#, cdPillarAggregateRange, 0#
        TEST_Harness.AssertTrue "1900-03-01_0_-1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-03-01_0_-1-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "1900-03-01_0_-1-value", mFact(3, 1), mFact(3, 2)
        ProbeAddPillar 2958465#, 1#, 0#, cdPillarAggregateRange, 0#
        TEST_Harness.AssertTrue "9999-12-31_1_0-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-31_1_0-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-12-31_1_0-value", mFact(3, 1), mFact(3, 2)
        ProbeAddPillar 2958465#, 0#, 1#, cdPillarAggregateRange, 0#
        TEST_Harness.AssertTrue "9999-12-31_0_1-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-31_0_1-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-12-31_0_1-value", mFact(3, 1), mFact(3, 2)
        ProbeAddPillar 2958434#, 1#, 0#, cdNone, 2958464#
        TEST_Harness.AssertTrue "9999-11-30_1_0-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-11-30_1_0-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "9999-11-30_1_0-value", mFact(3, 1), mFact(3, 2)
        ProbeAddPillar 46188#, (10# ^ 15), 0#, cdPillarAggregateRange, 0#
        TEST_Harness.AssertTrue "2026-06-15_1e+15_0-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-06-15_1e+15_0-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-06-15_1e+15_0-value", mFact(3, 1), mFact(3, 2)
        ProbeAddPillar 46188#, 0#, (10# ^ 300), cdPillarAggregateRange, 0#
        TEST_Harness.AssertTrue "2026-06-15_0_1e+300-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-06-15_0_1e+300-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-06-15_0_1e+300-value", mFact(3, 1), mFact(3, 2)
        ProbeAddPillar 46188#, 0#, -(10# ^ 300), cdPillarAggregateRange, 0#
        TEST_Harness.AssertTrue "2026-06-15_0_-1e+300-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-06-15_0_-1e+300-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "2026-06-15_0_-1e+300-value", mFact(3, 1), mFact(3, 2)

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

Private Sub CaseFormatPillar()

'
'------------------------------------------------------------------------------
'                           CHECK THE PILLAR FORMAT
'------------------------------------------------------------------------------
' PURPOSE
'   An interval formats as one canonical pillar under each rounding mode, or reports the window. 864 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date core follows the calculation contracts
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Formats intervals of 0 to 40 days from one start, negative intervals, clipped month ends, leap days and the edges of the window
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Dates, CORE_Err, TEST_Harness
'
' NOTES
'   Expected values come from tests/ reference computations in Python, not from the code
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "format-pillar"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Every rounding mode on short intervals, month edges, leap days, and both edges of the window
        ProbeFormatPillar 46023#, 46023#, prNearest, "0D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-01_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-01_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-01_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46023#, prFloor, "0D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-01_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-01_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-01_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46023#, prCeiling, "0D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-01_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-01_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-01_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46024#, prNearest, "1D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-02_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-02_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-02_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46024#, prFloor, "1D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-02_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-02_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-02_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46024#, prCeiling, "1D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-02_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-02_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-02_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46025#, prNearest, "2D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-03_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-03_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-03_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46025#, prFloor, "2D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-03_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-03_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-03_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46025#, prCeiling, "2D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-03_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-03_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-03_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46026#, prNearest, "3D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-04_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-04_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-04_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46026#, prFloor, "3D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-04_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-04_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-04_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46026#, prCeiling, "3D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-04_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-04_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-04_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46027#, prNearest, "4D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-05_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-05_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-05_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46027#, prFloor, "4D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-05_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-05_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-05_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46027#, prCeiling, "4D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-05_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-05_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-05_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46028#, prNearest, "5D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-06_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-06_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-06_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46028#, prFloor, "5D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-06_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-06_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-06_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46028#, prCeiling, "5D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-06_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-06_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-06_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46029#, prNearest, "6D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-07_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-07_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-07_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46029#, prFloor, "6D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-07_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-07_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-07_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46029#, prCeiling, "6D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-07_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-07_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-07_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46030#, prNearest, "1W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-08_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-08_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-08_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46030#, prFloor, "1W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-08_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-08_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-08_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46030#, prCeiling, "1W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-08_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-08_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-08_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46031#, prNearest, "1W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-09_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-09_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-09_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46031#, prFloor, "1W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-09_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-09_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-09_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46031#, prCeiling, "2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-09_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-09_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-09_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46032#, prNearest, "1W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-10_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-10_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-10_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46032#, prFloor, "1W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-10_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-10_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-10_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46032#, prCeiling, "2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-10_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-10_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-10_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46033#, prNearest, "1W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-11_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-11_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-11_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46033#, prFloor, "1W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-11_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-11_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-11_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46033#, prCeiling, "2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-11_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-11_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-11_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46034#, prNearest, "2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-12_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-12_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-12_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46034#, prFloor, "1W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-12_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-12_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-12_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46034#, prCeiling, "2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-12_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-12_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-12_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46035#, prNearest, "2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-13_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-13_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-13_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46035#, prFloor, "1W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-13_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-13_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-13_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46035#, prCeiling, "2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-13_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-13_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-13_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46036#, prNearest, "2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-14_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-14_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-14_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46036#, prFloor, "1W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-14_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-14_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-14_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46036#, prCeiling, "2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-14_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-14_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-14_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46037#, prNearest, "2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-15_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-15_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-15_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46037#, prFloor, "2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-15_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-15_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-15_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46037#, prCeiling, "2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-15_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-15_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-15_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46038#, prNearest, "2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-16_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-16_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-16_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46038#, prFloor, "2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-16_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-16_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-16_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46038#, prCeiling, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-16_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-16_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-16_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46039#, prNearest, "2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-17_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-17_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-17_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46039#, prFloor, "2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-17_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-17_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-17_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46039#, prCeiling, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-17_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-17_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-17_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46040#, prNearest, "2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-18_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-18_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-18_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46040#, prFloor, "2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-18_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-18_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-18_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46040#, prCeiling, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-18_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-18_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-18_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46041#, prNearest, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-19_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-19_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-19_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46041#, prFloor, "2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-19_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-19_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-19_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46041#, prCeiling, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-19_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-19_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-19_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46042#, prNearest, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-20_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-20_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-20_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46042#, prFloor, "2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-20_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-20_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-20_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46042#, prCeiling, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-20_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-20_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-20_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46043#, prNearest, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-21_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-21_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-21_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46043#, prFloor, "2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-21_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-21_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-21_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46043#, prCeiling, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-21_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-21_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-21_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46044#, prNearest, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-22_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-22_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-22_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46044#, prFloor, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-22_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-22_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-22_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46044#, prCeiling, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-22_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-22_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-22_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46045#, prNearest, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-23_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-23_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-23_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46045#, prFloor, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-23_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-23_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-23_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46045#, prCeiling, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-23_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-23_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-23_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46046#, prNearest, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-24_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-24_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-24_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46046#, prFloor, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-24_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-24_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-24_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46046#, prCeiling, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-24_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-24_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-24_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46047#, prNearest, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-25_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-25_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-25_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46047#, prFloor, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-25_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-25_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-25_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46047#, prCeiling, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-25_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-25_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-25_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46048#, prNearest, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-26_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-26_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-26_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46048#, prFloor, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-26_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-26_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-26_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46048#, prCeiling, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-26_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-26_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-26_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46049#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-27_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-27_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-27_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46049#, prFloor, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-27_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-27_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-27_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46049#, prCeiling, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-27_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-27_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-27_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46050#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-28_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-28_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-28_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46050#, prFloor, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-28_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-28_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-28_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46050#, prCeiling, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-28_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-28_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-28_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46051#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-29_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-29_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-29_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46051#, prFloor, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-29_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-29_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-29_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46051#, prCeiling, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-29_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-29_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-29_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46052#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-30_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-30_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-30_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46052#, prFloor, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-30_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-30_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-30_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46052#, prCeiling, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-30_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-30_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-30_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46053#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-31_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-31_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-31_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46053#, prFloor, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-31_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-31_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-31_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46053#, prCeiling, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-01-31_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-31_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-01-31_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46054#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-01_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-01_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-01_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46054#, prFloor, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-01_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-01_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-01_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46054#, prCeiling, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-01_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-01_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-01_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46055#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-02_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-02_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-02_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46055#, prFloor, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-02_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-02_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-02_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46055#, prCeiling, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-02_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-02_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-02_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46056#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-03_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-03_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-03_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46056#, prFloor, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-03_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-03_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-03_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46056#, prCeiling, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-03_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-03_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-03_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46057#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-04_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-04_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-04_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46057#, prFloor, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-04_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-04_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-04_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46057#, prCeiling, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-04_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-04_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-04_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46058#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-05_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-05_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-05_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46058#, prFloor, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-05_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-05_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-05_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46058#, prCeiling, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-05_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-05_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-05_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46059#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-06_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-06_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-06_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46059#, prFloor, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-06_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-06_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-06_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46059#, prCeiling, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-06_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-06_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-06_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46060#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-07_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-07_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-07_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46060#, prFloor, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-07_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-07_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-07_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46060#, prCeiling, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-07_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-07_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-07_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46061#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-08_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-08_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-08_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46061#, prFloor, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-08_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-08_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-08_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46061#, prCeiling, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-08_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-08_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-08_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46062#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-09_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-09_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-09_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46062#, prFloor, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-09_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-09_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-09_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46062#, prCeiling, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-09_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-09_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-09_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46063#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-10_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-10_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-10_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46063#, prFloor, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-10_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-10_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-10_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46063#, prCeiling, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-10_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-10_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-10_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46022#, prNearest, "-1D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-31_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-31_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-31_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46022#, prFloor, "-1D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-31_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-31_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-31_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46022#, prCeiling, "-1D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-31_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-31_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-31_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46017#, prNearest, "-6D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-26_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-26_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-26_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46017#, prFloor, "-6D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-26_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-26_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-26_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46017#, prCeiling, "-6D", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-26_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-26_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-26_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46016#, prNearest, "-1W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-25_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-25_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-25_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46016#, prFloor, "-1W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-25_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-25_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-25_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46016#, prCeiling, "-1W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-25_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-25_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-25_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46013#, prNearest, "-1W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-22_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-22_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-22_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46013#, prFloor, "-1W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-22_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-22_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-22_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46013#, prCeiling, "-2W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-22_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-22_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-22_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46002#, prNearest, "-3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-11_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-11_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-11_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46002#, prFloor, "-3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-11_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-11_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-11_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46002#, prCeiling, "-3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-11_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-11_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-11_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 45999#, prNearest, "-3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-08_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-08_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-08_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 45999#, prFloor, "-3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-08_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-08_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-08_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 45999#, prCeiling, "-1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-08_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-08_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-08_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 45998#, prNearest, "-3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-07_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-07_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-07_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 45998#, prFloor, "-3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-07_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-07_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-07_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 45998#, prCeiling, "-1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-07_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-07_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-07_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 45997#, prNearest, "-1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-06_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-06_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-06_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 45997#, prFloor, "-3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-06_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-06_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-06_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 45997#, prCeiling, "-1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-06_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-06_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-06_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 45996#, prNearest, "-1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-05_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-05_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-05_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 45996#, prFloor, "-3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-05_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-05_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-05_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 45996#, prCeiling, "-1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-05_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-05_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-05_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 45993#, prNearest, "-1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-02_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-02_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-02_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 45993#, prFloor, "-3W", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-02_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-02_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-02_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 45993#, prCeiling, "-1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-02_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-02_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-02_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 45992#, prNearest, "-1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-01_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-01_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-01_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 45992#, prFloor, "-1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-01_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-01_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-01_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 45992#, prCeiling, "-1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-12-01_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-01_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-12-01_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 45658#, prNearest, "-1Y", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-01-01_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-01-01_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-01-01_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 45658#, prFloor, "-1Y", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-01-01_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-01-01_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-01-01_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 45658#, prCeiling, "-1Y", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2025-01-01_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-01-01_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2025-01-01_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46068#, prNearest, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-15_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-15_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-15_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46068#, prFloor, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-15_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-15_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-15_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46068#, prCeiling, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-02-15_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-15_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-02-15_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46082#, prNearest, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-03-01_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-03-01_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-03-01_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46082#, prFloor, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-03-01_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-03-01_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-03-01_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46082#, prCeiling, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-03-01_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-03-01_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-03-01_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46083#, prNearest, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-03-02_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-03-02_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-03-02_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46083#, prFloor, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-03-02_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-03-02_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-03-02_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46083#, prCeiling, "3M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-03-02_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-03-02_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-03-02_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46113#, prNearest, "3M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-04-01_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-04-01_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-04-01_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46113#, prFloor, "3M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-04-01_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-04-01_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-04-01_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46113#, prCeiling, "3M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-04-01_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-04-01_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-04-01_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46203#, prNearest, "6M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-06-30_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-06-30_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-06-30_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46203#, prFloor, "5M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-06-30_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-06-30_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-06-30_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46203#, prCeiling, "6M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-06-30_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-06-30_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-06-30_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46387#, prNearest, "1Y", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-12-31_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-12-31_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-12-31_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46387#, prFloor, "11M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-12-31_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-12-31_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-12-31_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46387#, prCeiling, "1Y", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2026-12-31_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-12-31_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2026-12-31_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46388#, prNearest, "1Y", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2027-01-01_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2027-01-01_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2027-01-01_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46388#, prFloor, "1Y", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2027-01-01_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2027-01-01_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2027-01-01_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46388#, prCeiling, "1Y", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2027-01-01_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2027-01-01_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2027-01-01_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46389#, prNearest, "1Y", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2027-01-02_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2027-01-02_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2027-01-02_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46389#, prFloor, "1Y", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2027-01-02_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2027-01-02_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2027-01-02_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46389#, prCeiling, "1Y1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2027-01-02_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2027-01-02_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2027-01-02_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46423#, prNearest, "1Y1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2027-02-05_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2027-02-05_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2027-02-05_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46423#, prFloor, "1Y1M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2027-02-05_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2027-02-05_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2027-02-05_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46423#, prCeiling, "1Y2M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2027-02-05_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2027-02-05_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2027-02-05_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46753#, prNearest, "2Y", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2028-01-01_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2028-01-01_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2028-01-01_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46753#, prFloor, "2Y", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2028-01-01_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2028-01-01_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2028-01-01_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 46753#, prCeiling, "2Y", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2028-01-01_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2028-01-01_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2028-01-01_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 49673#, prNearest, "10Y", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2035-12-30_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2035-12-30_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2035-12-30_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 49673#, prFloor, "9Y11M", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2035-12-30_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2035-12-30_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2035-12-30_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46023#, 49673#, prCeiling, "10Y", cdNone
        TEST_Harness.AssertTrue "2026-01-01_2035-12-30_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2035-12-30_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-01_2035-12-30_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46060#, prNearest, "1W", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-02-07_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-07_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-07_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46060#, prFloor, "1W", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-02-07_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-07_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-07_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46060#, prCeiling, "1W", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-02-07_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-07_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-07_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46078#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-02-25_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-25_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-25_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46078#, prFloor, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-02-25_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-25_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-25_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46078#, prCeiling, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-02-25_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-25_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-25_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46079#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-02-26_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-26_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-26_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46079#, prFloor, "3W", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-02-26_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-26_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-26_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46079#, prCeiling, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-02-26_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-26_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-26_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46081#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-02-28_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-28_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-28_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46081#, prFloor, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-02-28_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-28_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-28_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46081#, prCeiling, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-02-28_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-28_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-02-28_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46082#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-03-01_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-01_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-01_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46082#, prFloor, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-03-01_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-01_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-01_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46082#, prCeiling, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-03-01_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-01_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-01_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46083#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-03-02_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-02_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-02_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46083#, prFloor, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-03-02_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-02_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-02_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46083#, prCeiling, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-03-02_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-02_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-02_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46084#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-03-03_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-03_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-03_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46084#, prFloor, "1M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-03-03_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-03_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-03_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46084#, prCeiling, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-03-03_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-03_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-03_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46112#, prNearest, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-03-31_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-31_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-31_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46112#, prFloor, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-03-31_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-31_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-31_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46112#, prCeiling, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-03-31_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-31_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-03-31_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46113#, prNearest, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-04-01_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-04-01_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-04-01_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46113#, prFloor, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-04-01_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-04-01_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-04-01_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46113#, prCeiling, "3M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-04-01_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-04-01_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-04-01_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46114#, prNearest, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-04-02_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-04-02_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-04-02_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46114#, prFloor, "2M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-04-02_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-04-02_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-04-02_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46114#, prCeiling, "3M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-04-02_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-04-02_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-04-02_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46418#, prNearest, "1Y", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2027-01-31_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2027-01-31_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2027-01-31_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46418#, prFloor, "1Y", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2027-01-31_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2027-01-31_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2027-01-31_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46418#, prCeiling, "1Y", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2027-01-31_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2027-01-31_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2027-01-31_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46028#, prNearest, "-3W", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-01-06_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-01-06_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-01-06_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46028#, prFloor, "-3W", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-01-06_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-01-06_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-01-06_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46028#, prCeiling, "-1M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-01-06_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-01-06_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-01-06_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46025#, prNearest, "-1M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-01-03_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-01-03_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-01-03_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46025#, prFloor, "-3W", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-01-03_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-01-03_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-01-03_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 46025#, prCeiling, "-1M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2026-01-03_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-01-03_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2026-01-03_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 45994#, prNearest, "-2M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2025-12-03_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2025-12-03_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2025-12-03_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 45994#, prFloor, "-1M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2025-12-03_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2025-12-03_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2025-12-03_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 46053#, 45994#, prCeiling, "-2M", cdNone
        TEST_Harness.AssertTrue "2026-01-31_2025-12-03_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2025-12-03_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2026-01-31_2025-12-03_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 45381#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2024-03-30_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-03-30_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-03-30_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 45381#, prFloor, "1M", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2024-03-30_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-03-30_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-03-30_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 45381#, prCeiling, "2M", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2024-03-30_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-03-30_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-03-30_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 45382#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2024-03-31_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-03-31_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-03-31_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 45382#, prFloor, "1M", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2024-03-31_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-03-31_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-03-31_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 45382#, prCeiling, "2M", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2024-03-31_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-03-31_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-03-31_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 45716#, prNearest, "1Y", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2025-02-28_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2025-02-28_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2025-02-28_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 45716#, prFloor, "1Y", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2025-02-28_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2025-02-28_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2025-02-28_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 45716#, prCeiling, "1Y", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2025-02-28_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2025-02-28_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2025-02-28_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 45717#, prNearest, "1Y", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2025-03-01_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2025-03-01_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2025-03-01_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 45717#, prFloor, "1Y", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2025-03-01_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2025-03-01_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2025-03-01_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 45717#, prCeiling, "1Y1M", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2025-03-01_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2025-03-01_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2025-03-01_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 46812#, prNearest, "4Y", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2028-02-29_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2028-02-29_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2028-02-29_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 46812#, prFloor, "4Y", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2028-02-29_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2028-02-29_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2028-02-29_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 46812#, prCeiling, "4Y", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2028-02-29_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2028-02-29_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2028-02-29_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 45321#, prNearest, "-1M", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2024-01-30_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-01-30_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-01-30_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 45321#, prFloor, "-3W", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2024-01-30_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-01-30_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-01-30_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 45321#, prCeiling, "-1M", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2024-01-30_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-01-30_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-01-30_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 45320#, prNearest, "-1M", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2024-01-29_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-01-29_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-01-29_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 45320#, prFloor, "-1M", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2024-01-29_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-01-29_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-01-29_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 45320#, prCeiling, "-1M", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2024-01-29_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-01-29_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2024-01-29_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 44986#, prNearest, "-1Y", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2023-03-01_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2023-03-01_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2023-03-01_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 44986#, prFloor, "-11M", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2023-03-01_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2023-03-01_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2023-03-01_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 45351#, 44986#, prCeiling, "-1Y", cdNone
        TEST_Harness.AssertTrue "2024-02-29_2023-03-01_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2023-03-01_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "2024-02-29_2023-03-01_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 70#, 61#, prNearest, "-1W", cdNone
        TEST_Harness.AssertTrue "1900-03-10_1900-03-01_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-03-10_1900-03-01_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "1900-03-10_1900-03-01_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 70#, 61#, prFloor, "-1W", cdNone
        TEST_Harness.AssertTrue "1900-03-10_1900-03-01_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-03-10_1900-03-01_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "1900-03-10_1900-03-01_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 70#, 61#, prCeiling, "", cdResultWindow
        TEST_Harness.AssertTrue "1900-03-10_1900-03-01_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-03-10_1900-03-01_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "1900-03-10_1900-03-01_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 65#, 61#, prNearest, "-4D", cdNone
        TEST_Harness.AssertTrue "1900-03-05_1900-03-01_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-03-05_1900-03-01_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "1900-03-05_1900-03-01_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 65#, 61#, prFloor, "-4D", cdNone
        TEST_Harness.AssertTrue "1900-03-05_1900-03-01_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-03-05_1900-03-01_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "1900-03-05_1900-03-01_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 65#, 61#, prCeiling, "-4D", cdNone
        TEST_Harness.AssertTrue "1900-03-05_1900-03-01_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-03-05_1900-03-01_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "1900-03-05_1900-03-01_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 61#, 92#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "1900-03-01_1900-04-01_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-03-01_1900-04-01_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "1900-03-01_1900-04-01_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 61#, 92#, prFloor, "1M", cdNone
        TEST_Harness.AssertTrue "1900-03-01_1900-04-01_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-03-01_1900-04-01_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "1900-03-01_1900-04-01_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 61#, 92#, prCeiling, "1M", cdNone
        TEST_Harness.AssertTrue "1900-03-01_1900-04-01_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-03-01_1900-04-01_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "1900-03-01_1900-04-01_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 61#, 426#, prNearest, "1Y", cdNone
        TEST_Harness.AssertTrue "1900-03-01_1901-03-01_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-03-01_1901-03-01_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "1900-03-01_1901-03-01_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 61#, 426#, prFloor, "1Y", cdNone
        TEST_Harness.AssertTrue "1900-03-01_1901-03-01_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-03-01_1901-03-01_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "1900-03-01_1901-03-01_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 61#, 426#, prCeiling, "1Y", cdNone
        TEST_Harness.AssertTrue "1900-03-01_1901-03-01_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-03-01_1901-03-01_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "1900-03-01_1901-03-01_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 2958454#, 2958465#, prNearest, "1W", cdNone
        TEST_Harness.AssertTrue "9999-12-20_9999-12-31_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-20_9999-12-31_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "9999-12-20_9999-12-31_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 2958454#, 2958465#, prFloor, "1W", cdNone
        TEST_Harness.AssertTrue "9999-12-20_9999-12-31_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-20_9999-12-31_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "9999-12-20_9999-12-31_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 2958454#, 2958465#, prCeiling, "", cdResultWindow
        TEST_Harness.AssertTrue "9999-12-20_9999-12-31_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-20_9999-12-31_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "9999-12-20_9999-12-31_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 2958465#, 2958454#, prNearest, "-2W", cdNone
        TEST_Harness.AssertTrue "9999-12-31_9999-12-20_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-31_9999-12-20_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "9999-12-31_9999-12-20_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 2958465#, 2958454#, prFloor, "-1W", cdNone
        TEST_Harness.AssertTrue "9999-12-31_9999-12-20_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-31_9999-12-20_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "9999-12-31_9999-12-20_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 2958465#, 2958454#, prCeiling, "-2W", cdNone
        TEST_Harness.AssertTrue "9999-12-31_9999-12-20_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-31_9999-12-20_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "9999-12-31_9999-12-20_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 2958435#, 2958465#, prNearest, "3W", cdNone
        TEST_Harness.AssertTrue "9999-12-01_9999-12-31_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-01_9999-12-31_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "9999-12-01_9999-12-31_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 2958435#, 2958465#, prFloor, "3W", cdNone
        TEST_Harness.AssertTrue "9999-12-01_9999-12-31_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-01_9999-12-31_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "9999-12-01_9999-12-31_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 2958435#, 2958465#, prCeiling, "", cdResultWindow
        TEST_Harness.AssertTrue "9999-12-01_9999-12-31_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-01_9999-12-31_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "9999-12-01_9999-12-31_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 2958434#, 2958465#, prNearest, "1M", cdNone
        TEST_Harness.AssertTrue "9999-11-30_9999-12-31_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-11-30_9999-12-31_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "9999-11-30_9999-12-31_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 2958434#, 2958465#, prFloor, "1M", cdNone
        TEST_Harness.AssertTrue "9999-11-30_9999-12-31_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-11-30_9999-12-31_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "9999-11-30_9999-12-31_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 2958434#, 2958465#, prCeiling, "", cdResultWindow
        TEST_Harness.AssertTrue "9999-11-30_9999-12-31_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-11-30_9999-12-31_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "9999-11-30_9999-12-31_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 61#, 2958465#, prNearest, "8099Y9M", cdNone
        TEST_Harness.AssertTrue "1900-03-01_9999-12-31_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-03-01_9999-12-31_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "1900-03-01_9999-12-31_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 61#, 2958465#, prFloor, "8099Y9M", cdNone
        TEST_Harness.AssertTrue "1900-03-01_9999-12-31_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-03-01_9999-12-31_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "1900-03-01_9999-12-31_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 61#, 2958465#, prCeiling, "", cdResultWindow
        TEST_Harness.AssertTrue "1900-03-01_9999-12-31_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "1900-03-01_9999-12-31_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "1900-03-01_9999-12-31_C-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 2958465#, 61#, prNearest, "-8099Y9M", cdNone
        TEST_Harness.AssertTrue "9999-12-31_1900-03-01_N-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-31_1900-03-01_N-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "9999-12-31_1900-03-01_N-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 2958465#, 61#, prFloor, "-8099Y9M", cdNone
        TEST_Harness.AssertTrue "9999-12-31_1900-03-01_F-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-31_1900-03-01_F-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "9999-12-31_1900-03-01_F-token", mFact(3, 1), mFact(3, 2)
        ProbeFormatPillar 2958465#, 61#, prCeiling, "", cdResultWindow
        TEST_Harness.AssertTrue "9999-12-31_1900-03-01_C-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "9999-12-31_1900-03-01_C-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "9999-12-31_1900-03-01_C-token", mFact(3, 1), mFact(3, 2)

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

Private Sub CasePillarLegacy()

'
'------------------------------------------------------------------------------
'                        CHECK THE LEGACY PILLAR TOKENS
'------------------------------------------------------------------------------
' PURPOSE
'   Oversized pillar tokens of the earlier project: grammar before numerical range. 28 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the pillar parser reads tokens as the data contract says
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Parses each token and asserts its result, condition, months and days
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Dates, CORE_Err, TEST_Harness
'
' NOTES
'   Ported from the earlier project's pillar regression suite: the tokens that
'   CaseParsePillar does not already hold
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HugeDigits      As String   'Four hundred digits, beyond the Double range

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "pillar-legacy"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Digits that no Double holds
        HugeDigits = String(400, "9")
    'Digits beyond the Double range are a range failure once the grammar is valid
        ProbeParsePillar HugeDigits & "M", 0#, 0#, cdPillarAggregateRange
        TEST_Harness.AssertTrue "oversized-months-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "oversized-months-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "oversized-months-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "oversized-months-days", mFact(4, 1), mFact(4, 2)
    'A finite year quantity whose month aggregate leaves the Double range
        ProbeParsePillar "2" & String(307, "0") & "Y", 0#, 0#, cdPillarAggregateRange
        TEST_Harness.AssertTrue "year-aggregate-overflow-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "year-aggregate-overflow-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "year-aggregate-overflow-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "year-aggregate-overflow-days", mFact(4, 1), mFact(4, 2)
    'The grammar is checked before any number is converted
        ProbeParsePillar HugeDigits, 0#, 0#, cdPillarTokenMalformed
        TEST_Harness.AssertTrue "oversized-missing-unit-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "oversized-missing-unit-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "oversized-missing-unit-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "oversized-missing-unit-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar HugeDigits & "X", 0#, 0#, cdPillarTokenMalformed
        TEST_Harness.AssertTrue "oversized-unknown-unit-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "oversized-unknown-unit-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "oversized-unknown-unit-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "oversized-unknown-unit-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar "1M" & HugeDigits & "M", 0#, 0#, cdPillarDuplicateUnit
        TEST_Harness.AssertTrue "oversized-repeated-unit-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "oversized-repeated-unit-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "oversized-repeated-unit-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "oversized-repeated-unit-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar HugeDigits & "M1M", 0#, 0#, cdPillarDuplicateUnit
        TEST_Harness.AssertTrue "overflow-first-then-duplicate-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "overflow-first-then-duplicate-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "overflow-first-then-duplicate-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "overflow-first-then-duplicate-days", mFact(4, 1), mFact(4, 2)
        ProbeParsePillar HugeDigits & "M1X", 0#, 0#, cdPillarTokenMalformed
        TEST_Harness.AssertTrue "overflow-first-then-unknown-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "overflow-first-then-unknown-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "overflow-first-then-unknown-months", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertEqualDouble "overflow-first-then-unknown-days", mFact(4, 1), mFact(4, 2)

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

Private Sub ProbeLong( _
    ByVal Expected As Long, _
    ByVal Actual As Long)

'
'------------------------------------------------------------------------------
'                             PROBE A WHOLE NUMBER
'------------------------------------------------------------------------------
' PURPOSE
'   Asserts that a whole number equals its expectation. 1 assertion
'
' WHY THIS EXISTS
'   The same checks repeat for many rows, so they are stated once
'
' INPUTS
'   Label
'     Name of the assertion
'
'   Expected
'     The value from the reference
'
'   Actual
'     The value from the code under test
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Compares the two values exactly
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
    'Whole numbers compare exactly
        mFact(1, 1) = Expected
        mFact(1, 2) = Actual

End Sub

Private Sub ProbeFlag( _
    ByVal Expected As Boolean, _
    ByVal Actual As Boolean)

'
'------------------------------------------------------------------------------
'                               PROBE A BOOLEAN
'------------------------------------------------------------------------------
' PURPOSE
'   Asserts that a Boolean equals its expectation. 1 assertion
'
' WHY THIS EXISTS
'   The same checks repeat for many rows, so they are stated once
'
' INPUTS
'   Label
'     Name of the assertion
'
'   Expected
'     The value from the reference
'
'   Actual
'     The value from the code under test
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Compares the two values
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
    'Booleans compare exactly
        mFact(1, 1) = (Actual = Expected)
        mFact(1, 2) = "the value differs from the expectation"

End Sub

Private Sub ProbeDate( _
    ByVal Expected As Double, _
    ByVal Actual As Date)

'
'------------------------------------------------------------------------------
'                                 PROBE A DATE
'------------------------------------------------------------------------------
' PURPOSE
'   Asserts that a date has the expected serial. 1 assertion
'
' WHY THIS EXISTS
'   The same checks repeat for many rows, so they are stated once
'
' INPUTS
'   Label
'     Name of the assertion
'
'   Expected
'     The serial from the reference
'
'   Actual
'     The date from the code under test
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Compares the serial of the date with the expectation
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
    'Whole days compare exactly
        mFact(1, 1) = Expected
        mFact(1, 2) = CDbl(Actual)

End Sub

Private Sub ProbeCondition( _
    ByVal Expected As KPRCondition, _
    ByVal Actual As KPRCondition)

'
'------------------------------------------------------------------------------
'                              PROBE A CONDITION
'------------------------------------------------------------------------------
' PURPOSE
'   Asserts that a condition is the expected one. 1 assertion
'
' WHY THIS EXISTS
'   The same checks repeat for many rows, so they are stated once
'
' INPUTS
'   Label
'     Name of the assertion
'
'   Expected
'     The condition from the contract
'
'   Actual
'     The condition from the code under test
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Compares the identifiers of the two conditions
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   CORE_Err.ConditionName, TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
    'The identifier is the semantic name that evidence cites
        mFact(1, 1) = CORE_Err.ConditionName(Expected)
        mFact(1, 2) = CORE_Err.ConditionName(Actual)

End Sub

Private Sub ProbeAddDays( _
    ByVal StartSerial As Double, _
    ByVal Days As Double, _
    ByVal ExpectedAccepted As Boolean, _
    ByVal ExpectedSerial As Double)

'
'------------------------------------------------------------------------------
'                            PROBE A SHIFT BY DAYS
'------------------------------------------------------------------------------
' PURPOSE
'   Shifts a date by days and checks the result and the output. 2 assertions
'
' WHY THIS EXISTS
'   The same checks repeat for many rows, so they are stated once
'
' INPUTS
'   Label
'     Prefix of the assertion names
'
'   StartSerial, Days
'     The inputs
'
'   ExpectedAccepted
'     Whether the shift stays inside the window
'
'   ExpectedSerial
'     The serial of the result; ignored when the shift is rejected
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Starts from a sentinel output so that an assignment on failure is seen
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   CORE_Dates.TryAddDays, TEST_Harness
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
    Dim DateOut         As Date     'Output of the shift, a sentinel on entry
    Dim Accepted        As Boolean  'Whether the shift was accepted
    Dim Expected        As Double   'Serial the output must hold

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Start from a sentinel, then shift
        DateOut = CDate(SENTINEL_SERIAL)
        Accepted = CORE_Dates.TryAddDays(CDate(StartSerial), Days, DateOut)
    'A rejection leaves the sentinel untouched
        If ExpectedAccepted Then
            Expected = ExpectedSerial
        Else
            Expected = SENTINEL_SERIAL
        End If

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        mFact(1, 1) = (CBool(Accepted) = CBool(ExpectedAccepted))
        mFact(1, 2) = "the value differs from the expectation"
        mFact(2, 1) = CDbl(Expected)
        mFact(2, 2) = CDbl(CDate(DateOut))

End Sub

Private Sub ProbeAddMonths( _
    ByVal StartSerial As Double, _
    ByVal Months As Double, _
    ByVal KeepEOM As Boolean, _
    ByVal ExpectedAccepted As Boolean, _
    ByVal ExpectedSerial As Double)

'
'------------------------------------------------------------------------------
'                           PROBE A SHIFT BY MONTHS
'------------------------------------------------------------------------------
' PURPOSE
'   Shifts a date by months and checks the result and the output. 2 assertions
'
' WHY THIS EXISTS
'   The same checks repeat for many rows, so they are stated once
'
' INPUTS
'   Label
'     Prefix of the assertion names
'
'   StartSerial, Months, KeepEOM
'     The inputs
'
'   ExpectedAccepted
'     Whether the shift stays inside the window
'
'   ExpectedSerial
'     The serial of the result; ignored when the shift is rejected
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Starts from a sentinel output so that an assignment on failure is seen
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   CORE_Dates.TryAddMonths, TEST_Harness
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
    Dim DateOut         As Date     'Output of the shift, a sentinel on entry
    Dim Accepted        As Boolean  'Whether the shift was accepted
    Dim Expected        As Double   'Serial the output must hold

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Start from a sentinel, then shift
        DateOut = CDate(SENTINEL_SERIAL)
        Accepted = CORE_Dates.TryAddMonths(CDate(StartSerial), Months, KeepEOM, DateOut)
    'A rejection leaves the sentinel untouched
        If ExpectedAccepted Then
            Expected = ExpectedSerial
        Else
            Expected = SENTINEL_SERIAL
        End If

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        mFact(1, 1) = (CBool(Accepted) = CBool(ExpectedAccepted))
        mFact(1, 2) = "the value differs from the expectation"
        mFact(2, 1) = CDbl(Expected)
        mFact(2, 2) = CDbl(CDate(DateOut))

End Sub

Private Sub ProbeNthWeekday( _
    ByVal YearIn As Long, _
    ByVal MonthIn As Long, _
    ByVal WeekdayIn As Long, _
    ByVal Occurrence As Long, _
    ByVal MondayBase As Boolean, _
    ByVal ExpectedCondition As KPRCondition, _
    ByVal ExpectedSerial As Double)

'
'------------------------------------------------------------------------------
'                            PROBE THE NTH WEEKDAY
'------------------------------------------------------------------------------
' PURPOSE
'   Locates a weekday and checks the result, the condition and the output. 3 assertions
'
' WHY THIS EXISTS
'   The same checks repeat for many rows, so they are stated once
'
' INPUTS
'   Label
'     Prefix of the assertion names
'
'   YearIn, MonthIn, WeekdayIn, Occurrence, MondayBase
'     The inputs
'
'   ExpectedCondition
'     cdNone for a date, otherwise the failing condition
'
'   ExpectedSerial
'     The serial of the result; ignored when a condition is expected
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Starts from a sentinel output so that an assignment on failure is seen
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   CORE_Dates, CORE_Err.ConditionName, TEST_Harness
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
    Dim DateOut         As Date         'Output of the locator, a sentinel on entry
    Dim Condition       As KPRCondition 'Condition that the locator reported
    Dim Accepted        As Boolean      'Whether a date was found
    Dim Expected        As Double       'Serial the output must hold

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Start from a sentinel, then locate
        DateOut = CDate(SENTINEL_SERIAL)
        Accepted = CORE_Dates.TryNthWeekday(YearIn, MonthIn, WeekdayIn, Occurrence, MondayBase, DateOut, Condition)
    'A failure leaves the sentinel untouched
        If ExpectedCondition = cdNone Then
            Expected = ExpectedSerial
        Else
            Expected = SENTINEL_SERIAL
        End If

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        mFact(1, 1) = (CBool(Accepted) = CBool((ExpectedCondition = cdNone)))
        mFact(1, 2) = "the value differs from the expectation"
        mFact(2, 1) = CORE_Err.ConditionName((ExpectedCondition))
        mFact(2, 2) = CORE_Err.ConditionName((Condition))
        mFact(3, 1) = CDbl(Expected)
        mFact(3, 2) = CDbl(CDate(DateOut))

End Sub

Private Sub ProbeLastWeekday( _
    ByVal YearIn As Long, _
    ByVal MonthIn As Long, _
    ByVal WeekdayIn As Long, _
    ByVal MondayBase As Boolean, _
    ByVal ExpectedCondition As KPRCondition, _
    ByVal ExpectedSerial As Double)

'
'------------------------------------------------------------------------------
'                            PROBE THE LAST WEEKDAY
'------------------------------------------------------------------------------
' PURPOSE
'   Locates a weekday and checks the result, the condition and the output. 3 assertions
'
' WHY THIS EXISTS
'   The same checks repeat for many rows, so they are stated once
'
' INPUTS
'   Label
'     Prefix of the assertion names
'
'   YearIn, MonthIn, WeekdayIn, MondayBase
'     The inputs
'
'   ExpectedCondition
'     cdNone for a date, otherwise the failing condition
'
'   ExpectedSerial
'     The serial of the result; ignored when a condition is expected
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Starts from a sentinel output so that an assignment on failure is seen
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   CORE_Dates, CORE_Err.ConditionName, TEST_Harness
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
    Dim DateOut         As Date         'Output of the locator, a sentinel on entry
    Dim Condition       As KPRCondition 'Condition that the locator reported
    Dim Accepted        As Boolean      'Whether a date was found
    Dim Expected        As Double       'Serial the output must hold

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Start from a sentinel, then locate
        DateOut = CDate(SENTINEL_SERIAL)
        Accepted = CORE_Dates.TryLastWeekday(YearIn, MonthIn, WeekdayIn, MondayBase, DateOut, Condition)
    'A failure leaves the sentinel untouched
        If ExpectedCondition = cdNone Then
            Expected = ExpectedSerial
        Else
            Expected = SENTINEL_SERIAL
        End If

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        mFact(1, 1) = (CBool(Accepted) = CBool((ExpectedCondition = cdNone)))
        mFact(1, 2) = "the value differs from the expectation"
        mFact(2, 1) = CORE_Err.ConditionName((ExpectedCondition))
        mFact(2, 2) = CORE_Err.ConditionName((Condition))
        mFact(3, 1) = CDbl(Expected)
        mFact(3, 2) = CDbl(CDate(DateOut))

End Sub

Private Sub ProbeParsePillar( _
    ByVal PillarIn As Variant, _
    ByVal ExpectedMonths As Double, _
    ByVal ExpectedDays As Double, _
    ByVal ExpectedCondition As KPRCondition)

'
'------------------------------------------------------------------------------
'                             PROBE A PILLAR PARSE
'------------------------------------------------------------------------------
' PURPOSE
'   Parses a pillar token and checks the result, the condition and both deltas. 4 assertions
'
' WHY THIS EXISTS
'   The same checks repeat for many rows, so they are stated once
'
' INPUTS
'   Label
'     Prefix of the assertion names
'
'   PillarIn
'     The token
'
'   ExpectedMonths, ExpectedDays
'     The deltas; ignored when a condition is expected
'
'   ExpectedCondition
'     cdNone for a token, otherwise the failing condition
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Starts from sentinel outputs so that an assignment on failure is seen
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   CORE_Dates.TryParsePillar, TEST_Harness
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
    Dim Months          As Double       'Month delta, a sentinel on entry
    Dim Days            As Double       'Day delta, a sentinel on entry
    Dim Condition       As KPRCondition 'Condition that the parser reported
    Dim Accepted        As Boolean      'Whether the token was accepted
    Dim WantMonths      As Double       'Value the month output must hold
    Dim WantDays        As Double       'Value the day output must hold

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Start from sentinels, then parse
        Months = SENTINEL_DELTA
        Days = SENTINEL_DELTA
        Accepted = CORE_Dates.TryParsePillar(PillarIn, Months, Days, Condition)
    'A rejection leaves the sentinels untouched
        If ExpectedCondition = cdNone Then
            WantMonths = ExpectedMonths
            WantDays = ExpectedDays
        Else
            WantMonths = SENTINEL_DELTA
            WantDays = SENTINEL_DELTA
        End If

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        mFact(1, 1) = (CBool(Accepted) = CBool((ExpectedCondition = cdNone)))
        mFact(1, 2) = "the value differs from the expectation"
        mFact(2, 1) = CORE_Err.ConditionName((ExpectedCondition))
        mFact(2, 2) = CORE_Err.ConditionName((Condition))
        mFact(3, 1) = WantMonths
        mFact(3, 2) = Months
        mFact(4, 1) = WantDays
        mFact(4, 2) = Days

End Sub

Private Sub ProbeAddPillar( _
    ByVal StartSerial As Double, _
    ByVal Months As Double, _
    ByVal Days As Double, _
    ByVal ExpectedCondition As KPRCondition, _
    ByVal ExpectedSerial As Double)

'
'------------------------------------------------------------------------------
'                             PROBE A PILLAR SHIFT
'------------------------------------------------------------------------------
' PURPOSE
'   Adds a parsed pillar to a date and checks the result, the condition and the output. 3 assertions
'
' WHY THIS EXISTS
'   The same checks repeat for many rows, so they are stated once
'
' INPUTS
'   Label
'     Prefix of the assertion names
'
'   StartSerial, Months, Days
'     The inputs
'
'   ExpectedCondition
'     cdNone for a date, otherwise cdPillarAggregateRange
'
'   ExpectedSerial
'     The serial of the result; ignored when a condition is expected
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Starts from a sentinel output so that an assignment on failure is seen
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   CORE_Dates.TryAddPillar, TEST_Harness
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
    Dim DateOut         As Date         'Output of the shift, a sentinel on entry
    Dim Condition       As KPRCondition 'Condition that the shift reported
    Dim Accepted        As Boolean      'Whether the shift was accepted
    Dim Expected        As Double       'Serial the output must hold

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Start from a sentinel, then shift
        DateOut = CDate(SENTINEL_SERIAL)
        Accepted = CORE_Dates.TryAddPillar(CDate(StartSerial), Months, Days, DateOut, Condition)
    'A rejection leaves the sentinel untouched
        If ExpectedCondition = cdNone Then
            Expected = ExpectedSerial
        Else
            Expected = SENTINEL_SERIAL
        End If

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        mFact(1, 1) = (CBool(Accepted) = CBool((ExpectedCondition = cdNone)))
        mFact(1, 2) = "the value differs from the expectation"
        mFact(2, 1) = CORE_Err.ConditionName((ExpectedCondition))
        mFact(2, 2) = CORE_Err.ConditionName((Condition))
        mFact(3, 1) = CDbl(Expected)
        mFact(3, 2) = CDbl(CDate(DateOut))

End Sub

Private Sub ProbeFormatPillar( _
    ByVal StartSerial As Double, _
    ByVal EndSerial As Double, _
    ByVal Mode As KPRPillarRounding, _
    ByVal ExpectedToken As String, _
    ByVal ExpectedCondition As KPRCondition)

'
'------------------------------------------------------------------------------
'                            PROBE A PILLAR FORMAT
'------------------------------------------------------------------------------
' PURPOSE
'   Formats an interval as a pillar and checks the result, the condition and the token. 3 assertions
'
' WHY THIS EXISTS
'   The same checks repeat for many rows, so they are stated once
'
' INPUTS
'   Label
'     Prefix of the assertion names
'
'   StartSerial, EndSerial, Mode
'     The inputs
'
'   ExpectedToken
'     The token; ignored when a condition is expected
'
'   ExpectedCondition
'     cdNone for a token, otherwise the failing condition
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Starts from a sentinel output so that an assignment on failure is seen
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   CORE_Dates.TryFormatPillar, TEST_Harness
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
    Dim TokenOut        As String       'Output of the format, a sentinel on entry
    Dim Condition       As KPRCondition 'Condition that the format reported
    Dim Accepted        As Boolean      'Whether a token was formatted
    Dim Expected        As String       'Text the output must hold

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Start from a sentinel, then format
        TokenOut = SENTINEL_TOKEN
        Accepted = CORE_Dates.TryFormatPillar(CDate(StartSerial), CDate(EndSerial), Mode, TokenOut, Condition)
    'A failure leaves the sentinel untouched
        If ExpectedCondition = cdNone Then
            Expected = ExpectedToken
        Else
            Expected = SENTINEL_TOKEN
        End If

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        mFact(1, 1) = (CBool(Accepted) = CBool((ExpectedCondition = cdNone)))
        mFact(1, 2) = "the value differs from the expectation"
        mFact(2, 1) = CORE_Err.ConditionName((ExpectedCondition))
        mFact(2, 2) = CORE_Err.ConditionName((Condition))
        mFact(3, 1) = Expected
        mFact(3, 2) = TokenOut

End Sub
