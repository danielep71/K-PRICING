Attribute VB_Name = "TEST_CoreErrParse"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: TEST_CoreErrParse
' PURPOSE
'   Tests of CORE_Err and CORE_Parse: the native error values, the registry of
'   conditions and the four argument parsers of the date layer
'
' WHY THIS EXISTS
'   The data contract defines how each argument is read and which error each
'   rejection returns. These cases check the code against the contract, and show that
'   each condition has its own identifier, so that no two rejections are confused
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   RunCoreErrParseTests runs the suite. Option Private Module keeps it out of the
'   workbook's external surface. The cases build their inputs in memory and change no
'   Excel state. The suite is not certified: it is not listed in
'   .github/excel-evidence-policy.json, and holds no case of the date functions
'
' ERROR POLICY
'   As TEST_Harness.EndRun: a run that does not pass raises
'   TEST_Harness.ERR_TEST_FAILED, or the unexpected error that escaped the runner. A
'   refused start raises TEST_Harness.ERR_TEST_DIRTY_START
'
' DEPENDENCIES
'   CORE_Err, CORE_Parse, TEST_Harness
'
' NOTES
'   In a development workbook holding the modules of src/ and tests/modules/, type
'   TEST_CoreErrParse.RunCoreErrParseTests in the Immediate window. Expected serials
'   are literals of the Excel 1900 date system and the condition expectations come from
'   the Conditions table of docs/methodology/DATA_CONTRACT.md: neither is read back
'   from the code under test. The 64-bit LongLong branch of the parsers is not tested,
'   since a LongLong cannot be built on a 32-bit host
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'A complete run starts every case and evaluates every assertion once
    Private Const SUITE_NAME            As String = "core-err-parse"    'Suite name in the run log
    Private Const EXPECTED_CASES        As Long = 11                  'Cases a complete run starts
    Private Const EXPECTED_ASSERTIONS   As Long = 455                 'Assertions a complete run evaluates

    'The native error numbers that the Conditions table lists
    Private Const ERROR_VALUE_CODE      As Long = 2015                  '#VALUE!
    Private Const ERROR_NUM_CODE        As Long = 2036                  '#NUM!
    Private Const ERROR_NA_CODE         As Long = 2042                  '#N/A

    'The values an output holds before a parse, which a rejection must leave in place
    Private Const SENTINEL_SERIAL       As Double = 36526#              '2000-01-01
    Private Const SENTINEL_LONG         As Long = 424242                'A Long no case parses
    Private Const SENTINEL_TOKEN        As String = "UNTOUCHED"         'A token no parser returns


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub RunCoreErrParseTests()

'
'------------------------------------------------------------------------------
'                       RUN CORE ERROR AND PARSER TESTS
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case in a fixed order between BeginRun and EndRun. A passing run ends
'   with RESULT=PASS; suite=core-err-parse; operation=PASS; completeness=COMPLETE;
'   cases=11; assertions=455; failures=0; cleanup=PASS; recovery=NONE
'
' WHY THIS EXISTS
'   One entry point runs the suite and reports it through the harness, so that a run
'   is complete or visibly incomplete
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
'   An error escaping a case handler is passed to EndRun, which reports it and
'   re-raises it unchanged. BeginRun's refusal propagates directly
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
        CaseErrorConstructors
        CaseConditionRegistry
        CaseUnmappedConditions
        CaseDateAccepted
        CaseDateTextRejected
        CaseDateTypeRejected
        CaseDateWindow
        CaseLongAccepted
        CaseLongRejected
        CaseBoolControl
        CaseRoundingControl

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

'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseErrorConstructors()

'
'------------------------------------------------------------------------------
'                         CHECK THE ERROR CONSTRUCTORS
'------------------------------------------------------------------------------
' PURPOSE
'   Each constructor returns its native Excel error. 3 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date layer reads its arguments as the data contract says
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Asserts the error number of ErrValue, ErrNum and ErrNA
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Err, TEST_Harness
'
' NOTES
'   None
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
        TEST_Harness.BeginCase "err-constructors"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Each constructor gives its own error
        TEST_Harness.AssertErrorValue "value", ERROR_VALUE_CODE, CORE_Err.ErrValue()
        TEST_Harness.AssertErrorValue "num", ERROR_NUM_CODE, CORE_Err.ErrNum()
        TEST_Harness.AssertErrorValue "na", ERROR_NA_CODE, CORE_Err.ErrNA()

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

Private Sub CaseConditionRegistry()

'
'------------------------------------------------------------------------------
'                         CHECK THE CONDITION REGISTRY
'------------------------------------------------------------------------------
' PURPOSE
'   Every condition has its identifier and its error, as the data contract lists them. 62 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date layer reads its arguments as the data contract says
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Checks the 29 conditions that map to an error, then the identifiers of the success
'   sentinel, the two propagation conditions and an unknown code
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Err, TEST_Harness
'
' NOTES
'   None
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
        TEST_Harness.BeginCase "condition-registry"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'The 29 conditions of the Conditions table that map to an error
        CheckCondition "DATE_TEXT_FORMAT", cdDateTextFormat, ERROR_VALUE_CODE
        CheckCondition "DATE_TEXT_LOCALE", cdDateTextLocale, ERROR_VALUE_CODE
        CheckCondition "DATE_TEXT_NUMERIC", cdDateTextNumeric, ERROR_VALUE_CODE
        CheckCondition "DATE_TEXT_IMPOSSIBLE", cdDateTextImpossible, ERROR_VALUE_CODE
        CheckCondition "DATE_TYPE_REJECTED", cdDateTypeRejected, ERROR_VALUE_CODE
        CheckCondition "DATE_WINDOW", cdDateWindow, ERROR_NUM_CODE
        CheckCondition "INPUT_BLANK_REQUIRED", cdInputBlankRequired, ERROR_VALUE_CODE
        CheckCondition "INTEGER_FRACTION", cdIntegerFraction, ERROR_VALUE_CODE
        CheckCondition "INTEGER_TYPE_REJECTED", cdIntegerTypeRejected, ERROR_VALUE_CODE
        CheckCondition "INTEGER_RANGE", cdIntegerRange, ERROR_NUM_CODE
        CheckCondition "DOMAIN_YEAR", cdDomainYear, ERROR_VALUE_CODE
        CheckCondition "DOMAIN_MONTH", cdDomainMonth, ERROR_VALUE_CODE
        CheckCondition "DOMAIN_WEEKDAY", cdDomainWeekday, ERROR_VALUE_CODE
        CheckCondition "DOMAIN_OCCURRENCE", cdDomainOccurrence, ERROR_VALUE_CODE
        CheckCondition "OCCURRENCE_ABSENT", cdOccurrenceAbsent, ERROR_NUM_CODE
        CheckCondition "RESULT_WINDOW", cdResultWindow, ERROR_NUM_CODE
        CheckCondition "PILLAR_TYPE_REJECTED", cdPillarTypeRejected, ERROR_VALUE_CODE
        CheckCondition "PILLAR_TOKEN_MALFORMED", cdPillarTokenMalformed, ERROR_VALUE_CODE
        CheckCondition "PILLAR_DUPLICATE_UNIT", cdPillarDuplicateUnit, ERROR_VALUE_CODE
        CheckCondition "PILLAR_ALIAS_SIGNED", cdPillarAliasSigned, ERROR_VALUE_CODE
        CheckCondition "PILLAR_AGGREGATE_RANGE", cdPillarAggregateRange, ERROR_NUM_CODE
        CheckCondition "CONTROL_TYPE_REJECTED", cdControlTypeRejected, ERROR_VALUE_CODE
        CheckCondition "CONTROL_TOKEN_UNKNOWN", cdControlTokenUnknown, ERROR_VALUE_CODE
        CheckCondition "CONTROL_NOT_SCALAR", cdControlNotScalar, ERROR_VALUE_CODE
        CheckCondition "SHAPE_UNSUPPORTED", cdShapeUnsupported, ERROR_VALUE_CODE
        CheckCondition "SHAPE_MISMATCH", cdShapeMismatch, ERROR_VALUE_CODE
        CheckCondition "CAPACITY_EXCEEDED", cdCapacityExceeded, ERROR_NUM_CODE
        CheckCondition "HOST_DATE1904", cdHostDate1904, ERROR_NA_CODE
        CheckCondition "HOST_UNRESOLVED", cdHostUnresolved, ERROR_NA_CODE
    'The conditions that map to none: no identifier for success or an unknown code,
    'and an identifier for each propagation condition
        CheckName "success", cdNone, vbNullString
        CheckName "unknown", 99, vbNullString
        CheckName "INPUT_ERROR_PROPAGATED", cdInputErrorPropagated, "INPUT_ERROR_PROPAGATED"
        CheckName "CONTROL_ERROR_PROPAGATED", cdControlErrorPropagated, "CONTROL_ERROR_PROPAGATED"

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

Private Sub CaseUnmappedConditions()

'
'------------------------------------------------------------------------------
'                        CHECK THE UNMAPPED CONDITIONS
'------------------------------------------------------------------------------
' PURPOSE
'   A condition with no worksheet error raises the one documented error. 12 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date layer reads its arguments as the data contract says
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Maps the success sentinel, the two propagation conditions and an unknown code, and
'   asserts the raised number, source and description
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Err, TEST_Harness
'
' NOTES
'   None
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
        TEST_Harness.BeginCase "unmapped-conditions"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'None of these has a worksheet error
        CheckUnmapped "success", cdNone
        CheckUnmapped "input-error-propagated", cdInputErrorPropagated
        CheckUnmapped "control-error-propagated", cdControlErrorPropagated
        CheckUnmapped "unknown", 99

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

Private Sub CaseDateAccepted()

'
'------------------------------------------------------------------------------
'                           CHECK THE ACCEPTED DATES
'------------------------------------------------------------------------------
' PURPOSE
'   Dates in every accepted form give their date-only serial. 45 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date layer reads its arguments as the data contract says
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Parses native dates, numbers of every numeric type and exact ISO text
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Err, CORE_Parse, TEST_Harness
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
    Dim WithTime      As Variant  'A date with a time of day

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "date-accepted"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'A native Date, with and without a time of day
        CheckDate "date-value", DateSerial(2026, 12, 31), 46387#, cdNone
        WithTime = DateSerial(2026, 12, 31) + TimeSerial(13, 45, 0)
        CheckDate "date-with-time", WithTime, 46387#, cdNone
    'Numbers of every numeric type, which keep their whole part
        CheckDate "double", 46387#, 46387#, cdNone
        CheckDate "double-with-time", 46387.75, 46387#, cdNone
        CheckDate "long", CLng(46387), 46387#, cdNone
        CheckDate "integer", CInt(100), 100#, cdNone
        CheckDate "single", CSng(46387.5), 46387#, cdNone
        CheckDate "currency", CCur(46387), 46387#, cdNone
        CheckDate "byte", CByte(100), 100#, cdNone
        CheckDate "decimal", CDec(46387.25), 46387#, cdNone
    'Exact ISO text, with the edges of the window and leap days
        CheckDate "iso-text", "2026-12-31", 46387#, cdNone
        CheckDate "iso-leap-day", "2024-02-29", 45351#, cdNone
        CheckDate "iso-leap-century", "2000-02-29", 36585#, cdNone
        CheckDate "iso-first-day", "1900-03-01", 61#, cdNone
        CheckDate "iso-last-day", "9999-12-31", 2958465#, cdNone

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

Private Sub CaseDateTextRejected()

'
'------------------------------------------------------------------------------
'                         CHECK THE REJECTED DATE TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Text that is not an exact, existing ISO date is rejected with its own condition. 90 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date layer reads its arguments as the data contract says
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Parses text of every rejected shape and asserts the condition, which keeps the
'   provenance of the error
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Err, CORE_Parse, TEST_Harness
'
' NOTES
'   None
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
        TEST_Harness.BeginCase "date-text-rejected"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Text that is not exact ISO and neither locale- nor number-shaped
        CheckDate "format-short-month-day", "2026-1-5", 0#, cdDateTextFormat
        CheckDate "format-slash-year-first", "2026/01/05", 0#, cdDateTextFormat
        CheckDate "format-trailing-blank", "2024-02-10 ", 0#, cdDateTextFormat
        CheckDate "format-leading-blank", " 2024-02-10", 0#, cdDateTextFormat
        CheckDate "format-time-suffix", "2024-02-10 12.30", 0#, cdDateTextFormat
        CheckDate "format-empty", "", 0#, cdDateTextFormat
        CheckDate "format-sign-only", "+", 0#, cdDateTextFormat
        CheckDate "format-point-only", ".", 0#, cdDateTextFormat
        CheckDate "format-two-points", "1.2.3", 0#, cdDateTextFormat
        CheckDate "format-letter-digit", "2026-01-0A", 0#, cdDateTextFormat
        CheckDate "format-timestamp", "2026-01-05T10:00", 0#, cdDateTextFormat
        CheckDate "format-localized-digit", "2026-01-0" & ChrW(&H665), 0#, cdDateTextFormat
    'Text shaped like a locale date
        CheckDate "locale-day-first", "31/12/2026", 0#, cdDateTextLocale
        CheckDate "locale-month-first", "12/31/2026", 0#, cdDateTextLocale
        CheckDate "locale-dots", "10.02.2024", 0#, cdDateTextLocale
        CheckDate "locale-short-year", "5/1/26", 0#, cdDateTextLocale
    'Text that looks like a number, never read as a serial
        CheckDate "numeric-serial", "61", 0#, cdDateTextNumeric
        CheckDate "numeric-serial-45292", "45292", 0#, cdDateTextNumeric
        CheckDate "numeric-serial-with-time", "45292.5", 0#, cdDateTextNumeric
        CheckDate "numeric-trailing-point", "1.", 0#, cdDateTextNumeric
        CheckDate "numeric-leading-point", ".5", 0#, cdDateTextNumeric
        CheckDate "numeric-negative", "-61", 0#, cdDateTextNumeric
        CheckDate "numeric-positive", "+7", 0#, cdDateTextNumeric
    'Exact ISO text naming a date that does not exist, 1900-02-29 included
        CheckDate "impossible-feb-29-common", "2025-02-29", 0#, cdDateTextImpossible
        CheckDate "impossible-feb-29-1900", "1900-02-29", 0#, cdDateTextImpossible
        CheckDate "impossible-month-13", "2026-13-01", 0#, cdDateTextImpossible
        CheckDate "impossible-month-0", "2026-00-10", 0#, cdDateTextImpossible
        CheckDate "impossible-day-32", "2026-01-32", 0#, cdDateTextImpossible
        CheckDate "impossible-day-0", "2026-01-00", 0#, cdDateTextImpossible
        CheckDate "impossible-april-31", "2026-04-31", 0#, cdDateTextImpossible

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

Private Sub CaseDateTypeRejected()

'
'------------------------------------------------------------------------------
'                        CHECK THE REJECTED DATE TYPES
'------------------------------------------------------------------------------
' PURPOSE
'   A blank, Null, Boolean, error value or object is not a date. 15 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date layer reads its arguments as the data contract says
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Parses each type and asserts its condition
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Err, CORE_Parse, TEST_Harness
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
    Dim Item            As Collection   'An object, which is never a date

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "date-type-rejected"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'An object is never a date
        Set Item = New Collection
    'A blank is contract-invalid, not a zero serial
        CheckDate "empty", Empty, 0#, cdInputBlankRequired
    'Every other type is a rejected type
        CheckDate "null", Null, 0#, cdDateTypeRejected
        CheckDate "boolean", True, 0#, cdDateTypeRejected
        CheckDate "error-value", CORE_Err.ErrValue(), 0#, cdDateTypeRejected
        CheckDate "object", Item, 0#, cdDateTypeRejected
    'Release the object
        Set Item = Nothing

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

Private Sub CaseDateWindow()

'
'------------------------------------------------------------------------------
'                            CHECK THE DATE WINDOW
'------------------------------------------------------------------------------
' PURPOSE
'   Dates outside 1900-03-01 to 9999-12-31 are rejected, and the edges are accepted. 42 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date layer reads its arguments as the data contract says
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Parses numbers, dates and text on and beside each edge of the window
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Err, CORE_Parse, TEST_Harness
'
' NOTES
'   None
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
        TEST_Harness.BeginCase "date-window"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'The window is serial 61 to 2958465, applied to the date-only value
        CheckDate "serial-60", 60#, 0#, cdDateWindow
        CheckDate "serial-61", 61#, 61#, cdNone
        CheckDate "serial-zero", 0#, 0#, cdDateWindow
        CheckDate "serial-negative", -1#, 0#, cdDateWindow
        CheckDate "serial-max", 2958465#, 2958465#, cdNone
        CheckDate "serial-over-max", 2958466#, 0#, cdDateWindow
        CheckDate "serial-last-day-with-time", 2958465.99, 2958465#, cdNone
        CheckDate "serial-below-min-with-time", 60.99, 0#, cdDateWindow
        CheckDate "date-below-min", DateSerial(1899, 12, 31), 0#, cdDateWindow
        CheckDate "text-1899", "1899-12-31", 0#, cdDateWindow
        CheckDate "text-year-zero", "0000-01-01", 0#, cdDateWindow
        CheckDate "text-year-50", "0050-01-01", 0#, cdDateWindow
        CheckDate "text-1900-02-28", "1900-02-28", 0#, cdDateWindow
        CheckDate "text-1900-01-01", "1900-01-01", 0#, cdDateWindow

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

Private Sub CaseLongAccepted()

'
'------------------------------------------------------------------------------
'                         CHECK THE ACCEPTED INTEGERS
'------------------------------------------------------------------------------
' PURPOSE
'   Whole numbers, and values within 1E-9 of one, are read as that whole number. 45 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date layer reads its arguments as the data contract says
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Parses numbers of every numeric type, the edges of the Long range and values on
'   each side of the tolerance
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Err, CORE_Parse, TEST_Harness
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
    Dim Residue          As Double   'A Double a hair below or above 3
    Dim DecimalInside    As Variant  'A Decimal 1E-9 less one unit in 3 away from 3

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "long-accepted"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'A Double that carries binary residue, and a Decimal inside the tolerance that a
    'Double conversion would push outside it
        Residue = (0.1 + 0.2) * 10#
        DecimalInside = CDec(3) + CDec(999999999) / CDec(1000000000000000000#)
    'Whole numbers of every numeric type, and values within 1E-9 of a whole number
        CheckLong "integer", CInt(7), 7, cdNone
        CheckLong "long", CLng(-5), -5, cdNone
        CheckLong "byte", CByte(200), 200, cdNone
        CheckLong "single", CSng(12), 12, cdNone
        CheckLong "double-whole", 3#, 3, cdNone
        CheckLong "zero", 0#, 0, cdNone
        CheckLong "currency", CCur(40), 40, cdNone
        CheckLong "decimal", CDec(12), 12, cdNone
        CheckLong "long-min", -2147483648#, -2147483647 - 1, cdNone
        CheckLong "long-max", 2147483647#, 2147483647, cdNone
        CheckLong "residue", Residue, 3, cdNone
        CheckLong "inside-high", 3.0000000001, 3, cdNone
        CheckLong "inside-low", 2.9999999999, 3, cdNone
        CheckLong "negative-inside", -2.9999999999, -3, cdNone
        CheckLong "decimal-inside", DecimalInside, 3, cdNone

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

Private Sub CaseLongRejected()

'
'------------------------------------------------------------------------------
'                         CHECK THE REJECTED INTEGERS
'------------------------------------------------------------------------------
' PURPOSE
'   A fraction, a value outside Long, a rejected type and a blank each have their own condition. 54 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date layer reads its arguments as the data contract says
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Parses each rejected form and asserts the condition
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Err, CORE_Parse, TEST_Harness
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
    Dim Item             As Collection   'An object, which is never an integer
    Dim DecimalOutside   As Variant      'A Decimal 1E-8 above 3

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "long-rejected"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'An object and a Decimal just outside the tolerance
        Set Item = New Collection
        DecimalOutside = CDec(3) + CDec(1) / CDec(100000000)
    'A real fraction is never truncated or rounded, and the range is tested first
        CheckLong "fraction-half", 2.5, 0, cdIntegerFraction
        CheckLong "fraction-small", 3.00000001, 0, cdIntegerFraction
        CheckLong "fraction-negative-half", -0.5, 0, cdIntegerFraction
        CheckLong "fraction-positive-half", 0.5, 0, cdIntegerFraction
        CheckLong "fraction-decimal", DecimalOutside, 0, cdIntegerFraction
        CheckLong "range-over-max", 2147483648#, 0, cdIntegerRange
        CheckLong "range-below-min", -2147483649#, 0, cdIntegerRange
        CheckLong "range-over-max-with-fraction", 2147483647.5, 0, cdIntegerRange
        CheckLong "range-below-min-with-fraction", -2147483648.5, 0, cdIntegerRange
        CheckLong "range-huge", (10000000000# * 10000000000#), 0, cdIntegerRange
        CheckLong "type-boolean", True, 0, cdIntegerTypeRejected
        CheckLong "type-date", DateSerial(2026, 1, 5), 0, cdIntegerTypeRejected
        CheckLong "type-text-number", "7", 0, cdIntegerTypeRejected
        CheckLong "type-text-empty", "", 0, cdIntegerTypeRejected
        CheckLong "type-null", Null, 0, cdIntegerTypeRejected
        CheckLong "type-object", Item, 0, cdIntegerTypeRejected
        CheckLong "type-error-value", CORE_Err.ErrValue(), 0, cdIntegerTypeRejected
        CheckLong "blank", Empty, 0, cdInputBlankRequired
    'Release the object
        Set Item = Nothing

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

Private Sub CaseBoolControl()

'
'------------------------------------------------------------------------------
'                          CHECK THE BOOLEAN CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   A Boolean control accepts Empty, True and False only. 30 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date layer reads its arguments as the data contract says
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Parses each accepted form and each rejected type and asserts the condition and the output
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Err, CORE_Parse, TEST_Harness
'
' NOTES
'   None
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
        TEST_Harness.BeginCase "bool-control"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Empty selects the default; only a native Boolean is accepted
        CheckBool "empty-default-true", Empty, True, True, cdNone
        CheckBool "empty-default-false", Empty, False, False, cdNone
        CheckBool "true", True, False, True, cdNone
        CheckBool "false", False, True, False, cdNone
        CheckBool "number-zero", 0, True, False, cdControlTypeRejected
        CheckBool "number-one", 1, False, False, cdControlTypeRejected
        CheckBool "text-true", "TRUE", False, False, cdControlTypeRejected
        CheckBool "null", Null, False, False, cdControlTypeRejected
        CheckBool "date", DateSerial(2026, 1, 5), False, False, cdControlTypeRejected
        CheckBool "error-value", CORE_Err.ErrValue(), False, False, cdControlTypeRejected

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

Private Sub CaseRoundingControl()

'
'------------------------------------------------------------------------------
'                          CHECK THE ROUNDING CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Opt_Rounding resolves to NEAREST, FLOOR or CEILING in every locale, or is rejected. 57 assertions.
'
' WHY THIS EXISTS
'   Part of the evidence that the date layer reads its arguments as the data contract says
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Parses each accepted form, each unknown token and each rejected type
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Err, CORE_Parse, TEST_Harness
'
' NOTES
'   None
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
        TEST_Harness.BeginCase "rounding-control"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Empty selects NEAREST; text is trimmed of blanks and tabs and folded in ASCII only
        CheckRounding "empty", Empty, "NEAREST", cdNone
        CheckRounding "nearest", "NEAREST", "NEAREST", cdNone
        CheckRounding "floor", "FLOOR", "FLOOR", cdNone
        CheckRounding "ceiling", "CEILING", "CEILING", cdNone
        CheckRounding "lower-case", "floor", "FLOOR", cdNone
        CheckRounding "mixed-case", "Ceiling", "CEILING", cdNone
        CheckRounding "padded-with-blanks", "  nearest  ", "NEAREST", cdNone
        CheckRounding "padded-with-tabs", vbTab & "FLOOR" & vbTab, "FLOOR", cdNone
        CheckRounding "unknown-token", "ROUND", "", cdControlTokenUnknown
        CheckRounding "empty-text", "", "", cdControlTokenUnknown
        CheckRounding "blanks-only", "   ", "", cdControlTokenUnknown
        CheckRounding "inner-blank", "CE ILING", "", cdControlTokenUnknown
        CheckRounding "dotless-i", "ce" & ChrW(&H131) & "ling", "", cdControlTokenUnknown
        CheckRounding "line-feed", vbLf & "FLOOR", "", cdControlTokenUnknown
        CheckRounding "number", 1, "", cdControlTypeRejected
        CheckRounding "boolean", True, "", cdControlTypeRejected
        CheckRounding "null", Null, "", cdControlTypeRejected
        CheckRounding "date", DateSerial(2026, 1, 5), "", cdControlTypeRejected
        CheckRounding "error-value", CORE_Err.ErrValue(), "", cdControlTypeRejected

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

Private Sub CheckCondition( _
    ByVal Identifier As String, _
    ByVal Condition As KPRCondition, _
    ByVal ExpectedCode As Long)

'
'------------------------------------------------------------------------------
'                             CHECK ONE CONDITION
'------------------------------------------------------------------------------
' PURPOSE
'   Checks the identifier and the native error of one condition. 2 assertions
'
' WHY THIS EXISTS
'   Each condition of the data contract's table is checked the same way
'
' INPUTS
'   Identifier
'     The identifier of the data contract's Conditions table
'
'   Condition
'     The enum member that mirrors it
'
'   ExpectedCode
'     The native error number the table gives it
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Asserts the identifier text and the error that the mapper returns
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   CORE_Err.ConditionName, CORE_Err.ErrForCondition, TEST_Harness
'
' NOTES
'   The expectations are generated from the Conditions table of
'   docs/methodology/DATA_CONTRACT.md, not read from the code
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN CHECKS
'------------------------------------------------------------------------------
    'The identifier is the semantic name that test cases and evidence cite
        TEST_Harness.AssertEqualText Identifier & "-identifier", Identifier, _
            CORE_Err.ConditionName(Condition)
    'The error is the one the data contract's Conditions table gives
        TEST_Harness.AssertErrorValue Identifier & "-error", ExpectedCode, _
            CORE_Err.ErrForCondition(Condition)

End Sub

Private Sub CheckName( _
    ByVal Label As String, _
    ByVal Condition As KPRCondition, _
    ByVal Expected As String)

'
'------------------------------------------------------------------------------
'                           CHECK ONE CONDITION NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Checks the identifier text of one condition that the mapper does not map.
'   1 assertion
'
' WHY THIS EXISTS
'   The propagation conditions, the success sentinel and an unknown code have no error
'   value to check, only an identifier or none
'
' INPUTS
'   Label
'     Stable prefix of the assertion name
'
'   Condition
'     The member or code
'
'   Expected
'     The identifier, or empty text for none
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Asserts the text that ConditionName returns
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
' RUN CHECK
'------------------------------------------------------------------------------
    'Compare the identifier text
        TEST_Harness.AssertEqualText Label & "-identifier", Expected, CORE_Err.ConditionName(Condition)

End Sub

Private Sub CheckUnmapped( _
    ByVal Label As String, _
    ByVal Condition As KPRCondition)

'
'------------------------------------------------------------------------------
'                         CHECK AN UNMAPPED CONDITION
'------------------------------------------------------------------------------
' PURPOSE
'   Checks that the mapper raises the one documented error for a condition that has no
'   worksheet error. 3 assertions
'
' WHY THIS EXISTS
'   A success sentinel or a propagation condition that reaches the mapper is a
'   defect that must not become a plausible error value
'
' INPUTS
'   Label
'     Stable prefix of the assertion names
'
'   Condition
'     A member or code that must not be mapped
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls the mapper, captures the error number, source and description, and asserts
'   each
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   CORE_Err.ErrForCondition, CORE_Err.ERR_UNMAPPED_CONDITION, TEST_Harness
'
' NOTES
'   The number, source and description are those of the data contract, Internal
'   failures
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Mapped          As Variant      'Result of the mapper, never assigned when it raises
    Dim RaisedNumber    As Long         'Err.Number captured after the call
    Dim RaisedSource    As String       'Err.Source captured after the call
    Dim RaisedText      As String       'Err.Description captured after the call

'------------------------------------------------------------------------------
' CALL MAPPER
'------------------------------------------------------------------------------
    'Capture the error of the call on the lines after it
        On Error Resume Next
        Mapped = CORE_Err.ErrForCondition(Condition)
        RaisedNumber = Err.Number
        RaisedSource = Err.Source
        RaisedText = Err.Description
        On Error GoTo 0

'------------------------------------------------------------------------------
' CHECK ERROR
'------------------------------------------------------------------------------
    'The error is the one internal defect, with its documented source and description
        TEST_Harness.AssertRaised Label & "-number", CORE_Err.ERR_UNMAPPED_CONDITION, RaisedNumber
        TEST_Harness.AssertEqualText Label & "-source", "CORE_Err.ErrForCondition", RaisedSource
        TEST_Harness.AssertEqualText Label & "-description", _
            "Condition has no worksheet error mapping; this is an internal defect.", RaisedText

End Sub

Private Sub CheckDate( _
    ByVal Label As String, _
    ByVal ScalarIn As Variant, _
    ByVal ExpectedSerial As Double, _
    ByVal ExpectedCondition As KPRCondition)

'
'------------------------------------------------------------------------------
'                            CHECK ONE DATE SCALAR
'------------------------------------------------------------------------------
' PURPOSE
'   Parses one scalar as a date and checks the result, the condition and the output.
'   3 assertions
'
' WHY THIS EXISTS
'   Every date case checks the same three facts, so they are stated once
'
' INPUTS
'   Label
'     Stable prefix of the assertion names
'
'   ScalarIn
'     The scalar to parse
'
'   ExpectedSerial
'     The serial the parser must assign; ignored when a condition is expected
'
'   ExpectedCondition
'     cdNone for an accepted scalar, otherwise the rejecting condition
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Parses with a sentinel in the output, then asserts the result, the condition name
'   and the value of the output. A rejection must leave the sentinel untouched
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   CORE_Parse.TryParseDateScalar, CORE_Err.ConditionName, TEST_Harness
'
' NOTES
'   Expected serials are written as literals from the Excel 1900 date system, never
'   read back from VBA
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date             'Output of the parser, a sentinel on entry
    Dim Condition       As KPRCondition     'Condition that the parser reported
    Dim Accepted        As Boolean          'Whether the parser accepted the scalar
    Dim ExpectedValue   As Double           'Serial the output must hold afterwards

'------------------------------------------------------------------------------
' RUN PARSER
'------------------------------------------------------------------------------
    'Start from a sentinel so that an assignment on failure is seen
        ParsedDate = CDate(SENTINEL_SERIAL)
    'Parse the scalar
        Accepted = CORE_Parse.TryParseDateScalar(ScalarIn, ParsedDate, Condition)
    'A failure leaves the output untouched; a success assigns the expected serial
        If ExpectedCondition = cdNone Then
            ExpectedValue = ExpectedSerial
        Else
            ExpectedValue = SENTINEL_SERIAL
        End If

'------------------------------------------------------------------------------
' CHECK RESULT
'------------------------------------------------------------------------------
    'The result, the condition and the output are three separate facts
        TEST_Harness.AssertTrue Label & "-result", Accepted = (ExpectedCondition = cdNone), _
            "the acceptance differs from the expectation"
        TEST_Harness.AssertEqualText Label & "-condition", _
            CORE_Err.ConditionName(ExpectedCondition), CORE_Err.ConditionName(Condition)
        TEST_Harness.AssertEqualDouble Label & "-value", ExpectedValue, VBA.CDbl(ParsedDate)

End Sub

Private Sub CheckLong( _
    ByVal Label As String, _
    ByVal ScalarIn As Variant, _
    ByVal ExpectedValue As Long, _
    ByVal ExpectedCondition As KPRCondition)

'
'------------------------------------------------------------------------------
'                           CHECK ONE INTEGER SCALAR
'------------------------------------------------------------------------------
' PURPOSE
'   Parses one scalar as a Long and checks the result, the condition and the output.
'   3 assertions
'
' WHY THIS EXISTS
'   Every integer case checks the same three facts, so they are stated once
'
' INPUTS
'   Label
'     Stable prefix of the assertion names
'
'   ScalarIn
'     The scalar to parse
'
'   ExpectedValue
'     The Long the parser must assign; ignored when a condition is expected
'
'   ExpectedCondition
'     cdNone for an accepted scalar, otherwise the rejecting condition
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Parses with a sentinel in the output, then asserts the result, the condition name
'   and the value of the output. A rejection must leave the sentinel untouched
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   CORE_Parse.TryParseLongScalar, CORE_Err.ConditionName, TEST_Harness
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
    Dim ParsedLong      As Long             'Output of the parser, a sentinel on entry
    Dim Condition       As KPRCondition     'Condition that the parser reported
    Dim Accepted        As Boolean          'Whether the parser accepted the scalar
    Dim ExpectedOutput  As Long             'Value the output must hold afterwards

'------------------------------------------------------------------------------
' RUN PARSER
'------------------------------------------------------------------------------
    'Start from a sentinel so that an assignment on failure is seen
        ParsedLong = SENTINEL_LONG
    'Parse the scalar
        Accepted = CORE_Parse.TryParseLongScalar(ScalarIn, ParsedLong, Condition)
    'A failure leaves the output untouched; a success assigns the expected value
        If ExpectedCondition = cdNone Then
            ExpectedOutput = ExpectedValue
        Else
            ExpectedOutput = SENTINEL_LONG
        End If

'------------------------------------------------------------------------------
' CHECK RESULT
'------------------------------------------------------------------------------
    'The result, the condition and the output are three separate facts
        TEST_Harness.AssertTrue Label & "-result", Accepted = (ExpectedCondition = cdNone), _
            "the acceptance differs from the expectation"
        TEST_Harness.AssertEqualText Label & "-condition", _
            CORE_Err.ConditionName(ExpectedCondition), CORE_Err.ConditionName(Condition)
        TEST_Harness.AssertEqualDouble Label & "-value", ExpectedOutput, ParsedLong

End Sub

Private Sub CheckBool( _
    ByVal Label As String, _
    ByVal ControlIn As Variant, _
    ByVal DefaultValue As Boolean, _
    ByVal ExpectedValue As Boolean, _
    ByVal ExpectedCondition As KPRCondition)

'
'------------------------------------------------------------------------------
'                          CHECK ONE BOOLEAN CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Parses one control as a Boolean and checks the result, the condition and the output.
'   3 assertions
'
' WHY THIS EXISTS
'   Every Boolean control case checks the same three facts, so they are stated once
'
' INPUTS
'   Label
'     Stable prefix of the assertion names
'
'   ControlIn
'     The control to parse
'
'   DefaultValue
'     The default the control selects when it is Empty
'
'   ExpectedValue
'     The Boolean the parser must assign; ignored when a condition is expected
'
'   ExpectedCondition
'     cdNone for an accepted control, otherwise the rejecting condition
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Parses with True in the output, then asserts the result, the condition name and
'   the value of the output. A rejection must leave True untouched
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   CORE_Parse.TryParseBoolControl, CORE_Err.ConditionName, TEST_Harness
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
    Dim ParsedBool      As Boolean          'Output of the parser, True on entry
    Dim Condition       As KPRCondition     'Condition that the parser reported
    Dim Accepted        As Boolean          'Whether the parser accepted the control
    Dim ExpectedOutput  As Boolean          'Value the output must hold afterwards

'------------------------------------------------------------------------------
' RUN PARSER
'------------------------------------------------------------------------------
    'Start from True, so that an assignment of False is seen and a rejection is not
    'mistaken for one
        ParsedBool = True
    'Parse the control
        Accepted = CORE_Parse.TryParseBoolControl(ControlIn, DefaultValue, ParsedBool, Condition)
    'A failure leaves the output untouched; a success assigns the expected value
        If ExpectedCondition = cdNone Then
            ExpectedOutput = ExpectedValue
        Else
            ExpectedOutput = True
        End If

'------------------------------------------------------------------------------
' CHECK RESULT
'------------------------------------------------------------------------------
    'The result, the condition and the output are three separate facts
        TEST_Harness.AssertTrue Label & "-result", Accepted = (ExpectedCondition = cdNone), _
            "the acceptance differs from the expectation"
        TEST_Harness.AssertEqualText Label & "-condition", _
            CORE_Err.ConditionName(ExpectedCondition), CORE_Err.ConditionName(Condition)
        TEST_Harness.AssertTrue Label & "-value", ParsedBool = ExpectedOutput, _
            "the output differs from the expectation"

End Sub

Private Sub CheckRounding( _
    ByVal Label As String, _
    ByVal ControlIn As Variant, _
    ByVal ExpectedToken As String, _
    ByVal ExpectedCondition As KPRCondition)

'
'------------------------------------------------------------------------------
'                          CHECK ONE ROUNDING CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Parses one control as a rounding token and checks the result, the condition and the
'   output. 3 assertions
'
' WHY THIS EXISTS
'   Every rounding case checks the same three facts, so they are stated once
'
' INPUTS
'   Label
'     Stable prefix of the assertion names
'
'   ControlIn
'     The control to parse
'
'   ExpectedToken
'     The token the parser must assign; ignored when a condition is expected
'
'   ExpectedCondition
'     cdNone for an accepted control, otherwise the rejecting condition
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Parses with a sentinel in the output, then asserts the result, the condition name
'   and the output. A rejection must leave the sentinel untouched
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   CORE_Parse.TryParseRoundingControl, CORE_Err.ConditionName, TEST_Harness
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
    Dim TokenOut        As String           'Output of the parser, a sentinel on entry
    Dim Condition       As KPRCondition     'Condition that the parser reported
    Dim Accepted        As Boolean          'Whether the parser accepted the control
    Dim ExpectedOutput  As String           'Text the output must hold afterwards

'------------------------------------------------------------------------------
' RUN PARSER
'------------------------------------------------------------------------------
    'Start from a sentinel so that an assignment on failure is seen
        TokenOut = SENTINEL_TOKEN
    'Parse the control
        Accepted = CORE_Parse.TryParseRoundingControl(ControlIn, TokenOut, Condition)
    'A failure leaves the output untouched; a success assigns the expected token
        If ExpectedCondition = cdNone Then
            ExpectedOutput = ExpectedToken
        Else
            ExpectedOutput = SENTINEL_TOKEN
        End If

'------------------------------------------------------------------------------
' CHECK RESULT
'------------------------------------------------------------------------------
    'The result, the condition and the output are three separate facts
        TEST_Harness.AssertTrue Label & "-result", Accepted = (ExpectedCondition = cdNone), _
            "the acceptance differs from the expectation"
        TEST_Harness.AssertEqualText Label & "-condition", _
            CORE_Err.ConditionName(ExpectedCondition), CORE_Err.ConditionName(Condition)
        TEST_Harness.AssertEqualText Label & "-token", ExpectedOutput, TokenOut

End Sub
