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
'   Excel state. It holds no case of the date functions. CertifyCoreErrParse certifies its normal run,
'   as .github/excel-evidence-policy.json lists it
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
    Private Const EXPECTED_CASES        As Long = 13                    'Cases a complete run starts
    Private Const EXPECTED_ASSERTIONS   As Long = 632                   'Assertions a complete run evaluates

    'The native error numbers that the Conditions table lists
    Private Const ERROR_VALUE_CODE      As Long = 2015                  '#VALUE!
    Private Const ERROR_NUM_CODE        As Long = 2036                  '#NUM!
    Private Const ERROR_NA_CODE         As Long = 2042                  '#N/A

    'The values an output holds before a parse, which a rejection must leave in place
    Private Const SENTINEL_SERIAL       As Double = 36526#              '2000-01-01
    Private Const SENTINEL_LONG         As Long = 424242                'A Long no case parses
    Private Const SENTINEL_TOKEN        As String = "UNTOUCHED"         'A token no parser returns

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

Public Sub RunCoreErrParseTests()

'
'------------------------------------------------------------------------------
'                       RUN CORE ERROR AND PARSER TESTS
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case in a fixed order between BeginRun and EndRun. A passing run ends
'   with RESULT=PASS; suite=core-err-parse; operation=PASS; completeness=COMPLETE;
'   cases=13; assertions=632; failures=0; cleanup=PASS; recovery=NONE
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
    Const PROC_NAME     As String = "TEST_CoreErrParse.RunCoreErrParseTests"

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
        CaseDateRegression
        CaseLongRegression

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

Public Sub CertifyCoreErrParse( _
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
'   Runs RunCoreErrParseTests, saves its log, certifies the run, then saves the certification log.
'
' ERROR POLICY
'   TEST_Harness.EndCertification raises when the certification did not pass.
'
' DEPENDENCIES
'   RunCoreErrParseTests, TEST_Harness
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
    Const PROC_NAME     As String = "TEST_CoreErrParse.CertifyCoreErrParse"

    Dim RunErrorNumber          As Long     'Error that ended the run; 0 for none
    Dim RunErrorDescription     As String   'Its description

'------------------------------------------------------------------------------
' CERTIFY
'------------------------------------------------------------------------------
    'Refuse state left by an earlier run and start the certification
        TEST_Harness.BeginCertification SUITE_NAME
    'Run the suite and capture the error that ended it
        On Error Resume Next
        RunCoreErrParseTests
        RunErrorNumber = Err.Number
        RunErrorDescription = Err.Description
    'Restore normal error handling
        On Error GoTo 0
    'Save the log and certify the run
        TEST_Harness.SaveLog Folder & "\core-err-parse.log"
        TEST_Harness.CertifyRun "NORMAL", True, RunErrorNumber, RunErrorDescription
    'Save the certification log, raising when the certification did not pass
        TEST_Harness.EndCertification Folder & "\core-err-parse-certification.log"

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
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_CoreErrParse.CaseErrorConstructors"

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
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_CoreErrParse.CaseConditionRegistry"

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
        ProbeCondition "DATE_TEXT_FORMAT", cdDateTextFormat, ERROR_VALUE_CODE
        TEST_Harness.AssertEqualText "DATE_TEXT_FORMAT-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "DATE_TEXT_FORMAT-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "DATE_TEXT_LOCALE", cdDateTextLocale, ERROR_VALUE_CODE
        TEST_Harness.AssertEqualText "DATE_TEXT_LOCALE-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "DATE_TEXT_LOCALE-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "DATE_TEXT_NUMERIC", cdDateTextNumeric, ERROR_VALUE_CODE
        TEST_Harness.AssertEqualText "DATE_TEXT_NUMERIC-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "DATE_TEXT_NUMERIC-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "DATE_TEXT_IMPOSSIBLE", cdDateTextImpossible, ERROR_VALUE_CODE
        TEST_Harness.AssertEqualText "DATE_TEXT_IMPOSSIBLE-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "DATE_TEXT_IMPOSSIBLE-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "DATE_TYPE_REJECTED", cdDateTypeRejected, ERROR_VALUE_CODE
        TEST_Harness.AssertEqualText "DATE_TYPE_REJECTED-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "DATE_TYPE_REJECTED-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "DATE_WINDOW", cdDateWindow, ERROR_NUM_CODE
        TEST_Harness.AssertEqualText "DATE_WINDOW-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "DATE_WINDOW-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "INPUT_BLANK_REQUIRED", cdInputBlankRequired, ERROR_VALUE_CODE
        TEST_Harness.AssertEqualText "INPUT_BLANK_REQUIRED-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "INPUT_BLANK_REQUIRED-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "INTEGER_FRACTION", cdIntegerFraction, ERROR_VALUE_CODE
        TEST_Harness.AssertEqualText "INTEGER_FRACTION-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "INTEGER_FRACTION-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "INTEGER_TYPE_REJECTED", cdIntegerTypeRejected, ERROR_VALUE_CODE
        TEST_Harness.AssertEqualText "INTEGER_TYPE_REJECTED-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "INTEGER_TYPE_REJECTED-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "INTEGER_RANGE", cdIntegerRange, ERROR_NUM_CODE
        TEST_Harness.AssertEqualText "INTEGER_RANGE-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "INTEGER_RANGE-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "DOMAIN_YEAR", cdDomainYear, ERROR_VALUE_CODE
        TEST_Harness.AssertEqualText "DOMAIN_YEAR-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "DOMAIN_YEAR-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "DOMAIN_MONTH", cdDomainMonth, ERROR_VALUE_CODE
        TEST_Harness.AssertEqualText "DOMAIN_MONTH-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "DOMAIN_MONTH-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "DOMAIN_WEEKDAY", cdDomainWeekday, ERROR_VALUE_CODE
        TEST_Harness.AssertEqualText "DOMAIN_WEEKDAY-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "DOMAIN_WEEKDAY-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "DOMAIN_OCCURRENCE", cdDomainOccurrence, ERROR_VALUE_CODE
        TEST_Harness.AssertEqualText "DOMAIN_OCCURRENCE-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "DOMAIN_OCCURRENCE-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "OCCURRENCE_ABSENT", cdOccurrenceAbsent, ERROR_NUM_CODE
        TEST_Harness.AssertEqualText "OCCURRENCE_ABSENT-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "OCCURRENCE_ABSENT-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "RESULT_WINDOW", cdResultWindow, ERROR_NUM_CODE
        TEST_Harness.AssertEqualText "RESULT_WINDOW-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "RESULT_WINDOW-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "PILLAR_TYPE_REJECTED", cdPillarTypeRejected, ERROR_VALUE_CODE
        TEST_Harness.AssertEqualText "PILLAR_TYPE_REJECTED-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "PILLAR_TYPE_REJECTED-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "PILLAR_TOKEN_MALFORMED", cdPillarTokenMalformed, ERROR_VALUE_CODE
        TEST_Harness.AssertEqualText "PILLAR_TOKEN_MALFORMED-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "PILLAR_TOKEN_MALFORMED-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "PILLAR_DUPLICATE_UNIT", cdPillarDuplicateUnit, ERROR_VALUE_CODE
        TEST_Harness.AssertEqualText "PILLAR_DUPLICATE_UNIT-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "PILLAR_DUPLICATE_UNIT-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "PILLAR_ALIAS_SIGNED", cdPillarAliasSigned, ERROR_VALUE_CODE
        TEST_Harness.AssertEqualText "PILLAR_ALIAS_SIGNED-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "PILLAR_ALIAS_SIGNED-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "PILLAR_AGGREGATE_RANGE", cdPillarAggregateRange, ERROR_NUM_CODE
        TEST_Harness.AssertEqualText "PILLAR_AGGREGATE_RANGE-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "PILLAR_AGGREGATE_RANGE-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "CONTROL_TYPE_REJECTED", cdControlTypeRejected, ERROR_VALUE_CODE
        TEST_Harness.AssertEqualText "CONTROL_TYPE_REJECTED-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "CONTROL_TYPE_REJECTED-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "CONTROL_TOKEN_UNKNOWN", cdControlTokenUnknown, ERROR_VALUE_CODE
        TEST_Harness.AssertEqualText "CONTROL_TOKEN_UNKNOWN-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "CONTROL_TOKEN_UNKNOWN-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "CONTROL_NOT_SCALAR", cdControlNotScalar, ERROR_VALUE_CODE
        TEST_Harness.AssertEqualText "CONTROL_NOT_SCALAR-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "CONTROL_NOT_SCALAR-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "SHAPE_UNSUPPORTED", cdShapeUnsupported, ERROR_VALUE_CODE
        TEST_Harness.AssertEqualText "SHAPE_UNSUPPORTED-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "SHAPE_UNSUPPORTED-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "SHAPE_MISMATCH", cdShapeMismatch, ERROR_VALUE_CODE
        TEST_Harness.AssertEqualText "SHAPE_MISMATCH-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "SHAPE_MISMATCH-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "CAPACITY_EXCEEDED", cdCapacityExceeded, ERROR_NUM_CODE
        TEST_Harness.AssertEqualText "CAPACITY_EXCEEDED-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "CAPACITY_EXCEEDED-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "HOST_DATE1904", cdHostDate1904, ERROR_NA_CODE
        TEST_Harness.AssertEqualText "HOST_DATE1904-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "HOST_DATE1904-error", mFact(2, 1), mFact(2, 2)
        ProbeCondition "HOST_UNRESOLVED", cdHostUnresolved, ERROR_NA_CODE
        TEST_Harness.AssertEqualText "HOST_UNRESOLVED-identifier", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertErrorValue "HOST_UNRESOLVED-error", mFact(2, 1), mFact(2, 2)
    'The conditions that map to none: no identifier for success or an unknown code,
    'and an identifier for each propagation condition
        ProbeName cdNone, vbNullString
        TEST_Harness.AssertEqualText "success-identifier", mFact(1, 1), mFact(1, 2)
        ProbeName 99, vbNullString
        TEST_Harness.AssertEqualText "unknown-identifier", mFact(1, 1), mFact(1, 2)
        ProbeName cdInputErrorPropagated, "INPUT_ERROR_PROPAGATED"
        TEST_Harness.AssertEqualText "INPUT_ERROR_PROPAGATED-identifier", mFact(1, 1), mFact(1, 2)
        ProbeName cdControlErrorPropagated, "CONTROL_ERROR_PROPAGATED"
        TEST_Harness.AssertEqualText "CONTROL_ERROR_PROPAGATED-identifier", mFact(1, 1), mFact(1, 2)

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
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_CoreErrParse.CaseUnmappedConditions"

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
        ProbeUnmapped cdNone
        TEST_Harness.AssertRaised "success-number", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "success-source", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "success-description", mFact(3, 1), mFact(3, 2)
        ProbeUnmapped cdInputErrorPropagated
        TEST_Harness.AssertRaised "input-error-propagated-number", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "input-error-propagated-source", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "input-error-propagated-description", mFact(3, 1), mFact(3, 2)
        ProbeUnmapped cdControlErrorPropagated
        TEST_Harness.AssertRaised "control-error-propagated-number", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "control-error-propagated-source", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "control-error-propagated-description", mFact(3, 1), mFact(3, 2)
        ProbeUnmapped 99
        TEST_Harness.AssertRaised "unknown-number", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "unknown-source", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "unknown-description", mFact(3, 1), mFact(3, 2)

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
'   Dates in every accepted form give their date-only serial. 48 assertions.
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
    Const PROC_NAME     As String = "TEST_CoreErrParse.CaseDateAccepted"

    Dim WithTime      As Variant  'A date with a time of day
    Dim BelowNextDay  As Variant  'A Decimal serial 1E-14 below 46388

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
        ProbeDate DateSerial(2026, 12, 31), 46387#, cdNone
        TEST_Harness.AssertTrue "date-value-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "date-value-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "date-value-value", mFact(3, 1), mFact(3, 2)
        WithTime = DateSerial(2026, 12, 31) + TimeSerial(13, 45, 0)
        ProbeDate WithTime, 46387#, cdNone
        TEST_Harness.AssertTrue "date-with-time-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "date-with-time-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "date-with-time-value", mFact(3, 1), mFact(3, 2)
    'Numbers of every numeric type, which keep their whole part
        ProbeDate 46387#, 46387#, cdNone
        TEST_Harness.AssertTrue "double-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "double-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "double-value", mFact(3, 1), mFact(3, 2)
        ProbeDate 46387.75, 46387#, cdNone
        TEST_Harness.AssertTrue "double-with-time-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "double-with-time-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "double-with-time-value", mFact(3, 1), mFact(3, 2)
        ProbeDate CLng(46387), 46387#, cdNone
        TEST_Harness.AssertTrue "long-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "long-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "long-value", mFact(3, 1), mFact(3, 2)
        ProbeDate CInt(100), 100#, cdNone
        TEST_Harness.AssertTrue "integer-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "integer-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "integer-value", mFact(3, 1), mFact(3, 2)
        ProbeDate CSng(46387.5), 46387#, cdNone
        TEST_Harness.AssertTrue "single-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "single-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "single-value", mFact(3, 1), mFact(3, 2)
        ProbeDate CCur(46387), 46387#, cdNone
        TEST_Harness.AssertTrue "currency-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "currency-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "currency-value", mFact(3, 1), mFact(3, 2)
        ProbeDate CByte(100), 100#, cdNone
        TEST_Harness.AssertTrue "byte-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "byte-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "byte-value", mFact(3, 1), mFact(3, 2)
        ProbeDate CDec(46387.25), 46387#, cdNone
        TEST_Harness.AssertTrue "decimal-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "decimal-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "decimal-value", mFact(3, 1), mFact(3, 2)
    'A Decimal a hair below the next day keeps the day before; a Double conversion would not
        BelowNextDay = CDec(46388) - CDec(1) / CDec(100000000000000#)
        ProbeDate BelowNextDay, 46387#, cdNone
        TEST_Harness.AssertTrue "decimal-below-next-day-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "decimal-below-next-day-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "decimal-below-next-day-value", mFact(3, 1), mFact(3, 2)
    'Exact ISO text, with the edges of the window and leap days
        ProbeDate "2026-12-31", 46387#, cdNone
        TEST_Harness.AssertTrue "iso-text-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "iso-text-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "iso-text-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "2024-02-29", 45351#, cdNone
        TEST_Harness.AssertTrue "iso-leap-day-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "iso-leap-day-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "iso-leap-day-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "2000-02-29", 36585#, cdNone
        TEST_Harness.AssertTrue "iso-leap-century-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "iso-leap-century-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "iso-leap-century-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "1900-03-01", 61#, cdNone
        TEST_Harness.AssertTrue "iso-first-day-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "iso-first-day-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "iso-first-day-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "9999-12-31", 2958465#, cdNone
        TEST_Harness.AssertTrue "iso-last-day-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "iso-last-day-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "iso-last-day-value", mFact(3, 1), mFact(3, 2)

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
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_CoreErrParse.CaseDateTextRejected"

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
        ProbeDate "2026-1-5", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "format-short-month-day-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "format-short-month-day-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "format-short-month-day-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "2026/01/05", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "format-slash-year-first-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "format-slash-year-first-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "format-slash-year-first-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "2024-02-10 ", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "format-trailing-blank-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "format-trailing-blank-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "format-trailing-blank-value", mFact(3, 1), mFact(3, 2)
        ProbeDate " 2024-02-10", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "format-leading-blank-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "format-leading-blank-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "format-leading-blank-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "2024-02-10 12.30", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "format-time-suffix-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "format-time-suffix-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "format-time-suffix-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "format-empty-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "format-empty-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "format-empty-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "+", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "format-sign-only-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "format-sign-only-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "format-sign-only-value", mFact(3, 1), mFact(3, 2)
        ProbeDate ".", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "format-point-only-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "format-point-only-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "format-point-only-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "1.2.3", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "format-two-points-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "format-two-points-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "format-two-points-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "2026-01-0A", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "format-letter-digit-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "format-letter-digit-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "format-letter-digit-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "2026-01-05T10:00", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "format-timestamp-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "format-timestamp-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "format-timestamp-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "2026-01-0" & ChrW(&H665), 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "format-localized-digit-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "format-localized-digit-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "format-localized-digit-value", mFact(3, 1), mFact(3, 2)
    'Text shaped like a locale date
        ProbeDate "31/12/2026", 0#, cdDateTextLocale
        TEST_Harness.AssertTrue "locale-day-first-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "locale-day-first-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "locale-day-first-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "12/31/2026", 0#, cdDateTextLocale
        TEST_Harness.AssertTrue "locale-month-first-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "locale-month-first-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "locale-month-first-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "10.02.2024", 0#, cdDateTextLocale
        TEST_Harness.AssertTrue "locale-dots-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "locale-dots-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "locale-dots-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "5/1/26", 0#, cdDateTextLocale
        TEST_Harness.AssertTrue "locale-short-year-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "locale-short-year-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "locale-short-year-value", mFact(3, 1), mFact(3, 2)
    'Text that looks like a number, never read as a serial
        ProbeDate "61", 0#, cdDateTextNumeric
        TEST_Harness.AssertTrue "numeric-serial-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "numeric-serial-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "numeric-serial-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "45292", 0#, cdDateTextNumeric
        TEST_Harness.AssertTrue "numeric-serial-45292-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "numeric-serial-45292-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "numeric-serial-45292-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "45292.5", 0#, cdDateTextNumeric
        TEST_Harness.AssertTrue "numeric-serial-with-time-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "numeric-serial-with-time-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "numeric-serial-with-time-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "1.", 0#, cdDateTextNumeric
        TEST_Harness.AssertTrue "numeric-trailing-point-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "numeric-trailing-point-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "numeric-trailing-point-value", mFact(3, 1), mFact(3, 2)
        ProbeDate ".5", 0#, cdDateTextNumeric
        TEST_Harness.AssertTrue "numeric-leading-point-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "numeric-leading-point-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "numeric-leading-point-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "-61", 0#, cdDateTextNumeric
        TEST_Harness.AssertTrue "numeric-negative-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "numeric-negative-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "numeric-negative-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "+7", 0#, cdDateTextNumeric
        TEST_Harness.AssertTrue "numeric-positive-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "numeric-positive-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "numeric-positive-value", mFact(3, 1), mFact(3, 2)
    'Exact ISO text naming a date that does not exist, 1900-02-29 included
        ProbeDate "2025-02-29", 0#, cdDateTextImpossible
        TEST_Harness.AssertTrue "impossible-feb-29-common-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "impossible-feb-29-common-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "impossible-feb-29-common-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "1900-02-29", 0#, cdDateTextImpossible
        TEST_Harness.AssertTrue "impossible-feb-29-1900-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "impossible-feb-29-1900-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "impossible-feb-29-1900-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "2026-13-01", 0#, cdDateTextImpossible
        TEST_Harness.AssertTrue "impossible-month-13-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "impossible-month-13-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "impossible-month-13-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "2026-00-10", 0#, cdDateTextImpossible
        TEST_Harness.AssertTrue "impossible-month-0-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "impossible-month-0-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "impossible-month-0-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "2026-01-32", 0#, cdDateTextImpossible
        TEST_Harness.AssertTrue "impossible-day-32-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "impossible-day-32-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "impossible-day-32-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "2026-01-00", 0#, cdDateTextImpossible
        TEST_Harness.AssertTrue "impossible-day-0-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "impossible-day-0-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "impossible-day-0-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "2026-04-31", 0#, cdDateTextImpossible
        TEST_Harness.AssertTrue "impossible-april-31-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "impossible-april-31-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "impossible-april-31-value", mFact(3, 1), mFact(3, 2)

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
    Const PROC_NAME     As String = "TEST_CoreErrParse.CaseDateTypeRejected"

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
        ProbeDate Empty, 0#, cdInputBlankRequired
        TEST_Harness.AssertTrue "empty-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "empty-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "empty-value", mFact(3, 1), mFact(3, 2)
    'Every other type is a rejected type
        ProbeDate Null, 0#, cdDateTypeRejected
        TEST_Harness.AssertTrue "null-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "null-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "null-value", mFact(3, 1), mFact(3, 2)
        ProbeDate True, 0#, cdDateTypeRejected
        TEST_Harness.AssertTrue "boolean-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "boolean-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "boolean-value", mFact(3, 1), mFact(3, 2)
        ProbeDate CORE_Err.ErrValue(), 0#, cdDateTypeRejected
        TEST_Harness.AssertTrue "error-value-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "error-value-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "error-value-value", mFact(3, 1), mFact(3, 2)
        ProbeDate Item, 0#, cdDateTypeRejected
        TEST_Harness.AssertTrue "object-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "object-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "object-value", mFact(3, 1), mFact(3, 2)
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
'   Dates outside 1900-03-01 to 9999-12-31 are rejected, and the edges are accepted. 45 assertions.
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
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_CoreErrParse.CaseDateWindow"

    Dim BelowMin      As Variant  'A Decimal serial 1E-14 below 61

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
        ProbeDate 60#, 0#, cdDateWindow
        TEST_Harness.AssertTrue "serial-60-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "serial-60-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "serial-60-value", mFact(3, 1), mFact(3, 2)
        ProbeDate 61#, 61#, cdNone
        TEST_Harness.AssertTrue "serial-61-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "serial-61-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "serial-61-value", mFact(3, 1), mFact(3, 2)
        ProbeDate 0#, 0#, cdDateWindow
        TEST_Harness.AssertTrue "serial-zero-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "serial-zero-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "serial-zero-value", mFact(3, 1), mFact(3, 2)
        ProbeDate -1#, 0#, cdDateWindow
        TEST_Harness.AssertTrue "serial-negative-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "serial-negative-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "serial-negative-value", mFact(3, 1), mFact(3, 2)
        ProbeDate 2958465#, 2958465#, cdNone
        TEST_Harness.AssertTrue "serial-max-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "serial-max-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "serial-max-value", mFact(3, 1), mFact(3, 2)
        ProbeDate 2958466#, 0#, cdDateWindow
        TEST_Harness.AssertTrue "serial-over-max-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "serial-over-max-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "serial-over-max-value", mFact(3, 1), mFact(3, 2)
        ProbeDate 2958465.99, 2958465#, cdNone
        TEST_Harness.AssertTrue "serial-last-day-with-time-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "serial-last-day-with-time-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "serial-last-day-with-time-value", mFact(3, 1), mFact(3, 2)
        ProbeDate 60.99, 0#, cdDateWindow
        TEST_Harness.AssertTrue "serial-below-min-with-time-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "serial-below-min-with-time-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "serial-below-min-with-time-value", mFact(3, 1), mFact(3, 2)
        ProbeDate DateSerial(1899, 12, 31), 0#, cdDateWindow
        TEST_Harness.AssertTrue "date-below-min-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "date-below-min-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "date-below-min-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "1899-12-31", 0#, cdDateWindow
        TEST_Harness.AssertTrue "text-1899-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "text-1899-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "text-1899-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "0000-01-01", 0#, cdDateWindow
        TEST_Harness.AssertTrue "text-year-zero-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "text-year-zero-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "text-year-zero-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "0050-01-01", 0#, cdDateWindow
        TEST_Harness.AssertTrue "text-year-50-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "text-year-50-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "text-year-50-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "1900-02-28", 0#, cdDateWindow
        TEST_Harness.AssertTrue "text-1900-02-28-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "text-1900-02-28-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "text-1900-02-28-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "1900-01-01", 0#, cdDateWindow
        TEST_Harness.AssertTrue "text-1900-01-01-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "text-1900-01-01-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "text-1900-01-01-value", mFact(3, 1), mFact(3, 2)
    'A Decimal a hair below serial 61 floors to 60, outside the window
        BelowMin = CDec(61) - CDec(1) / CDec(100000000000000#)
        ProbeDate BelowMin, 0#, cdDateWindow
        TEST_Harness.AssertTrue "decimal-below-min-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "decimal-below-min-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "decimal-below-min-value", mFact(3, 1), mFact(3, 2)

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
    Const PROC_NAME     As String = "TEST_CoreErrParse.CaseLongAccepted"

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
        ProbeLong CInt(7), 7, cdNone
        TEST_Harness.AssertTrue "integer-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "integer-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "integer-value", mFact(3, 1), mFact(3, 2)
        ProbeLong CLng(-5), -5, cdNone
        TEST_Harness.AssertTrue "long-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "long-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "long-value", mFact(3, 1), mFact(3, 2)
        ProbeLong CByte(200), 200, cdNone
        TEST_Harness.AssertTrue "byte-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "byte-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "byte-value", mFact(3, 1), mFact(3, 2)
        ProbeLong CSng(12), 12, cdNone
        TEST_Harness.AssertTrue "single-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "single-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "single-value", mFact(3, 1), mFact(3, 2)
        ProbeLong 3#, 3, cdNone
        TEST_Harness.AssertTrue "double-whole-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "double-whole-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "double-whole-value", mFact(3, 1), mFact(3, 2)
        ProbeLong 0#, 0, cdNone
        TEST_Harness.AssertTrue "zero-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "zero-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "zero-value", mFact(3, 1), mFact(3, 2)
        ProbeLong CCur(40), 40, cdNone
        TEST_Harness.AssertTrue "currency-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "currency-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "currency-value", mFact(3, 1), mFact(3, 2)
        ProbeLong CDec(12), 12, cdNone
        TEST_Harness.AssertTrue "decimal-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "decimal-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "decimal-value", mFact(3, 1), mFact(3, 2)
        ProbeLong -2147483648#, -2147483647 - 1, cdNone
        TEST_Harness.AssertTrue "long-min-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "long-min-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "long-min-value", mFact(3, 1), mFact(3, 2)
        ProbeLong 2147483647#, 2147483647, cdNone
        TEST_Harness.AssertTrue "long-max-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "long-max-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "long-max-value", mFact(3, 1), mFact(3, 2)
        ProbeLong Residue, 3, cdNone
        TEST_Harness.AssertTrue "residue-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "residue-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "residue-value", mFact(3, 1), mFact(3, 2)
        ProbeLong 3.0000000001, 3, cdNone
        TEST_Harness.AssertTrue "inside-high-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "inside-high-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "inside-high-value", mFact(3, 1), mFact(3, 2)
        ProbeLong 2.9999999999, 3, cdNone
        TEST_Harness.AssertTrue "inside-low-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "inside-low-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "inside-low-value", mFact(3, 1), mFact(3, 2)
        ProbeLong -2.9999999999, -3, cdNone
        TEST_Harness.AssertTrue "negative-inside-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "negative-inside-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "negative-inside-value", mFact(3, 1), mFact(3, 2)
        ProbeLong DecimalInside, 3, cdNone
        TEST_Harness.AssertTrue "decimal-inside-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "decimal-inside-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "decimal-inside-value", mFact(3, 1), mFact(3, 2)

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
'   A fraction, a value outside Long, a rejected type and a blank each have their own condition. 60 assertions.
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
    Const PROC_NAME     As String = "TEST_CoreErrParse.CaseLongRejected"

    Dim Item             As Collection   'An object, which is never an integer
    Dim DecimalOutside   As Variant      'A Decimal 1E-8 above 3
    Dim DecimalAboveMax  As Variant      'A Decimal 1E-10 above the largest Long
    Dim DecimalBelowMin  As Variant      'A Decimal 1E-10 below the smallest Long

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
        DecimalAboveMax = CDec(2147483647) + CDec(1) / CDec(10000000000#)
        DecimalBelowMin = CDec(-2147483648#) - CDec(1) / CDec(10000000000#)
    'A real fraction is never truncated or rounded, and the range is tested first
        ProbeLong 2.5, 0, cdIntegerFraction
        TEST_Harness.AssertTrue "fraction-half-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "fraction-half-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "fraction-half-value", mFact(3, 1), mFact(3, 2)
        ProbeLong 3.00000001, 0, cdIntegerFraction
        TEST_Harness.AssertTrue "fraction-small-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "fraction-small-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "fraction-small-value", mFact(3, 1), mFact(3, 2)
        ProbeLong -0.5, 0, cdIntegerFraction
        TEST_Harness.AssertTrue "fraction-negative-half-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "fraction-negative-half-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "fraction-negative-half-value", mFact(3, 1), mFact(3, 2)
        ProbeLong 0.5, 0, cdIntegerFraction
        TEST_Harness.AssertTrue "fraction-positive-half-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "fraction-positive-half-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "fraction-positive-half-value", mFact(3, 1), mFact(3, 2)
        ProbeLong DecimalOutside, 0, cdIntegerFraction
        TEST_Harness.AssertTrue "fraction-decimal-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "fraction-decimal-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "fraction-decimal-value", mFact(3, 1), mFact(3, 2)
        ProbeLong 2147483648#, 0, cdIntegerRange
        TEST_Harness.AssertTrue "range-over-max-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "range-over-max-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "range-over-max-value", mFact(3, 1), mFact(3, 2)
        ProbeLong -2147483649#, 0, cdIntegerRange
        TEST_Harness.AssertTrue "range-below-min-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "range-below-min-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "range-below-min-value", mFact(3, 1), mFact(3, 2)
        ProbeLong 2147483647.5, 0, cdIntegerRange
        TEST_Harness.AssertTrue "range-over-max-with-fraction-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "range-over-max-with-fraction-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "range-over-max-with-fraction-value", mFact(3, 1), mFact(3, 2)
        ProbeLong -2147483648.5, 0, cdIntegerRange
        TEST_Harness.AssertTrue "range-below-min-with-fraction-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "range-below-min-with-fraction-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "range-below-min-with-fraction-value", mFact(3, 1), mFact(3, 2)
        ProbeLong DecimalAboveMax, 0, cdIntegerRange
        TEST_Harness.AssertTrue "range-decimal-above-max-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "range-decimal-above-max-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "range-decimal-above-max-value", mFact(3, 1), mFact(3, 2)
        ProbeLong DecimalBelowMin, 0, cdIntegerRange
        TEST_Harness.AssertTrue "range-decimal-below-min-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "range-decimal-below-min-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "range-decimal-below-min-value", mFact(3, 1), mFact(3, 2)
        ProbeLong (10000000000# * 10000000000#), 0, cdIntegerRange
        TEST_Harness.AssertTrue "range-huge-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "range-huge-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "range-huge-value", mFact(3, 1), mFact(3, 2)
        ProbeLong True, 0, cdIntegerTypeRejected
        TEST_Harness.AssertTrue "type-boolean-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "type-boolean-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "type-boolean-value", mFact(3, 1), mFact(3, 2)
        ProbeLong DateSerial(2026, 1, 5), 0, cdIntegerTypeRejected
        TEST_Harness.AssertTrue "type-date-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "type-date-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "type-date-value", mFact(3, 1), mFact(3, 2)
        ProbeLong "7", 0, cdIntegerTypeRejected
        TEST_Harness.AssertTrue "type-text-number-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "type-text-number-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "type-text-number-value", mFact(3, 1), mFact(3, 2)
        ProbeLong "", 0, cdIntegerTypeRejected
        TEST_Harness.AssertTrue "type-text-empty-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "type-text-empty-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "type-text-empty-value", mFact(3, 1), mFact(3, 2)
        ProbeLong Null, 0, cdIntegerTypeRejected
        TEST_Harness.AssertTrue "type-null-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "type-null-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "type-null-value", mFact(3, 1), mFact(3, 2)
        ProbeLong Item, 0, cdIntegerTypeRejected
        TEST_Harness.AssertTrue "type-object-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "type-object-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "type-object-value", mFact(3, 1), mFact(3, 2)
        ProbeLong CORE_Err.ErrValue(), 0, cdIntegerTypeRejected
        TEST_Harness.AssertTrue "type-error-value-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "type-error-value-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "type-error-value-value", mFact(3, 1), mFact(3, 2)
        ProbeLong Empty, 0, cdInputBlankRequired
        TEST_Harness.AssertTrue "blank-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "blank-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "blank-value", mFact(3, 1), mFact(3, 2)
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
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_CoreErrParse.CaseBoolControl"

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
        ProbeBool Empty, True, True, cdNone
        TEST_Harness.AssertTrue "empty-default-true-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "empty-default-true-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertTrue "empty-default-true-value", mFact(3, 1), mFact(3, 2)
        ProbeBool Empty, False, False, cdNone
        TEST_Harness.AssertTrue "empty-default-false-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "empty-default-false-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertTrue "empty-default-false-value", mFact(3, 1), mFact(3, 2)
        ProbeBool True, False, True, cdNone
        TEST_Harness.AssertTrue "true-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "true-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertTrue "true-value", mFact(3, 1), mFact(3, 2)
        ProbeBool False, True, False, cdNone
        TEST_Harness.AssertTrue "false-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "false-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertTrue "false-value", mFact(3, 1), mFact(3, 2)
        ProbeBool 0, True, False, cdControlTypeRejected
        TEST_Harness.AssertTrue "number-zero-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "number-zero-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertTrue "number-zero-value", mFact(3, 1), mFact(3, 2)
        ProbeBool 1, False, False, cdControlTypeRejected
        TEST_Harness.AssertTrue "number-one-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "number-one-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertTrue "number-one-value", mFact(3, 1), mFact(3, 2)
        ProbeBool "TRUE", False, False, cdControlTypeRejected
        TEST_Harness.AssertTrue "text-true-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "text-true-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertTrue "text-true-value", mFact(3, 1), mFact(3, 2)
        ProbeBool Null, False, False, cdControlTypeRejected
        TEST_Harness.AssertTrue "null-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "null-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertTrue "null-value", mFact(3, 1), mFact(3, 2)
        ProbeBool DateSerial(2026, 1, 5), False, False, cdControlTypeRejected
        TEST_Harness.AssertTrue "date-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "date-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertTrue "date-value", mFact(3, 1), mFact(3, 2)
        ProbeBool CORE_Err.ErrValue(), False, False, cdControlTypeRejected
        TEST_Harness.AssertTrue "error-value-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "error-value-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertTrue "error-value-value", mFact(3, 1), mFact(3, 2)

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
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_CoreErrParse.CaseRoundingControl"

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
        ProbeRounding Empty, "NEAREST", cdNone
        TEST_Harness.AssertTrue "empty-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "empty-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "empty-token", mFact(3, 1), mFact(3, 2)
        ProbeRounding "NEAREST", "NEAREST", cdNone
        TEST_Harness.AssertTrue "nearest-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "nearest-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "nearest-token", mFact(3, 1), mFact(3, 2)
        ProbeRounding "FLOOR", "FLOOR", cdNone
        TEST_Harness.AssertTrue "floor-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "floor-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "floor-token", mFact(3, 1), mFact(3, 2)
        ProbeRounding "CEILING", "CEILING", cdNone
        TEST_Harness.AssertTrue "ceiling-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "ceiling-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "ceiling-token", mFact(3, 1), mFact(3, 2)
        ProbeRounding "floor", "FLOOR", cdNone
        TEST_Harness.AssertTrue "lower-case-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "lower-case-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "lower-case-token", mFact(3, 1), mFact(3, 2)
        ProbeRounding "Ceiling", "CEILING", cdNone
        TEST_Harness.AssertTrue "mixed-case-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "mixed-case-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "mixed-case-token", mFact(3, 1), mFact(3, 2)
        ProbeRounding "  nearest  ", "NEAREST", cdNone
        TEST_Harness.AssertTrue "padded-with-blanks-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "padded-with-blanks-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "padded-with-blanks-token", mFact(3, 1), mFact(3, 2)
        ProbeRounding vbTab & "FLOOR" & vbTab, "FLOOR", cdNone
        TEST_Harness.AssertTrue "padded-with-tabs-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "padded-with-tabs-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "padded-with-tabs-token", mFact(3, 1), mFact(3, 2)
        ProbeRounding "ROUND", "", cdControlTokenUnknown
        TEST_Harness.AssertTrue "unknown-token-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "unknown-token-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "unknown-token-token", mFact(3, 1), mFact(3, 2)
        ProbeRounding "", "", cdControlTokenUnknown
        TEST_Harness.AssertTrue "empty-text-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "empty-text-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "empty-text-token", mFact(3, 1), mFact(3, 2)
        ProbeRounding "   ", "", cdControlTokenUnknown
        TEST_Harness.AssertTrue "blanks-only-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "blanks-only-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "blanks-only-token", mFact(3, 1), mFact(3, 2)
        ProbeRounding "CE ILING", "", cdControlTokenUnknown
        TEST_Harness.AssertTrue "inner-blank-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "inner-blank-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "inner-blank-token", mFact(3, 1), mFact(3, 2)
        ProbeRounding "ce" & ChrW(&H131) & "ling", "", cdControlTokenUnknown
        TEST_Harness.AssertTrue "dotless-i-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "dotless-i-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "dotless-i-token", mFact(3, 1), mFact(3, 2)
        ProbeRounding vbLf & "FLOOR", "", cdControlTokenUnknown
        TEST_Harness.AssertTrue "line-feed-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "line-feed-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "line-feed-token", mFact(3, 1), mFact(3, 2)
        ProbeRounding 1, "", cdControlTypeRejected
        TEST_Harness.AssertTrue "number-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "number-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "number-token", mFact(3, 1), mFact(3, 2)
        ProbeRounding True, "", cdControlTypeRejected
        TEST_Harness.AssertTrue "boolean-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "boolean-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "boolean-token", mFact(3, 1), mFact(3, 2)
        ProbeRounding Null, "", cdControlTypeRejected
        TEST_Harness.AssertTrue "null-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "null-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "null-token", mFact(3, 1), mFact(3, 2)
        ProbeRounding DateSerial(2026, 1, 5), "", cdControlTypeRejected
        TEST_Harness.AssertTrue "date-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "date-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "date-token", mFact(3, 1), mFact(3, 2)
        ProbeRounding CORE_Err.ErrValue(), "", cdControlTypeRejected
        TEST_Harness.AssertTrue "error-value-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "error-value-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "error-value-token", mFact(3, 1), mFact(3, 2)

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

Private Sub CaseDateRegression()

'
'------------------------------------------------------------------------------
'                       CHECK THE REGRESSION DATE INPUTS
'------------------------------------------------------------------------------
' PURPOSE
'   Regression date inputs that the cases above do not hold. 114 assertions.
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
'   Parses each input and asserts its result, condition and serial
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Err, CORE_Parse, TEST_Harness
'
' NOTES
'   Each expectation is the data contract's; none is read back from the code under test
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_CoreErrParse.CaseDateRegression"

    'None

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "date-regression"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'A native Date and a Date with a time of day keep the day
        ProbeDate DateSerial(2026, 3, 15), 46096#, cdNone
        TEST_Harness.AssertTrue "native-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "native-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "native-value", mFact(3, 1), mFact(3, 2)
        ProbeDate DateSerial(2026, 3, 15) + 0.75, 46096#, cdNone
        TEST_Harness.AssertTrue "native-with-time-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "native-with-time-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "native-with-time-value", mFact(3, 1), mFact(3, 2)
    'Each numeric type is a serial, and its fraction is dropped
        ProbeDate CByte(255), 255#, cdNone
        TEST_Harness.AssertTrue "byte-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "byte-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "byte-value", mFact(3, 1), mFact(3, 2)
        ProbeDate CInt(1000), 1000#, cdNone
        TEST_Harness.AssertTrue "integer-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "integer-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "integer-value", mFact(3, 1), mFact(3, 2)
        ProbeDate CLng(46096), 46096#, cdNone
        TEST_Harness.AssertTrue "long-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "long-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "long-value", mFact(3, 1), mFact(3, 2)
        ProbeDate CSng(46096.75), 46096#, cdNone
        TEST_Harness.AssertTrue "single-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "single-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "single-value", mFact(3, 1), mFact(3, 2)
        ProbeDate 46096#, 46096#, cdNone
        TEST_Harness.AssertTrue "integral-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "integral-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "integral-value", mFact(3, 1), mFact(3, 2)
        ProbeDate 46096.75, 46096#, cdNone
        TEST_Harness.AssertTrue "fractional-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "fractional-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "fractional-value", mFact(3, 1), mFact(3, 2)
        ProbeDate CCur(46096.75), 46096#, cdNone
        TEST_Harness.AssertTrue "currency-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "currency-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "currency-value", mFact(3, 1), mFact(3, 2)
        ProbeDate CDec(46096.75), 46096#, cdNone
        TEST_Harness.AssertTrue "decimal-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "decimal-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "decimal-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "2026-03-15", 46096#, cdNone
        TEST_Harness.AssertTrue "iso-exact-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "iso-exact-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "iso-exact-value", mFact(3, 1), mFact(3, 2)
    'False and the #N/A error value are rejected types
        ProbeDate False, 0#, cdDateTypeRejected
        TEST_Harness.AssertTrue "boolean-false-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "boolean-false-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "boolean-false-value", mFact(3, 1), mFact(3, 2)
        ProbeDate CVErr(2042), 0#, cdDateTypeRejected
        TEST_Harness.AssertTrue "error-na-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "error-na-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "error-na-value", mFact(3, 1), mFact(3, 2)
    'Day-first, month-first and short-year text is locale-dependent
        ProbeDate "31.12.2026", 0#, cdDateTextLocale
        TEST_Harness.AssertTrue "dotted-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "dotted-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "dotted-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "1/2/26", 0#, cdDateTextLocale
        TEST_Harness.AssertTrue "short-dmy-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "short-dmy-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "short-dmy-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "10.02/2024", 0#, cdDateTextLocale
        TEST_Harness.AssertTrue "mixed-separators-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "mixed-separators-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "mixed-separators-value", mFact(3, 1), mFact(3, 2)
    'Other text that is not yyyy-mm-dd is malformed
        ProbeDate "2024/02/10", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "year-first-slash-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "year-first-slash-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "year-first-slash-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "2024.02.10", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "year-first-dot-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "year-first-dot-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "year-first-dot-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "310/12/2026", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "three-digit-day-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "three-digit-day-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "three-digit-day-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "31/12/20260", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "five-digit-year-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "five-digit-year-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "five-digit-year-value", mFact(3, 1), mFact(3, 2)
        ProbeDate " 31/12/2026", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "padded-dmy-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "padded-dmy-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "padded-dmy-value", mFact(3, 1), mFact(3, 2)
        ProbeDate " 2026-03-15", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "leading-space-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "leading-space-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "leading-space-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "2026-03-15 ", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "trailing-space-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "trailing-space-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "trailing-space-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "2026-3-15", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "omitted-zeros-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "omitted-zeros-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "omitted-zeros-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "2026-03-15T00:00", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "time-suffix-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "time-suffix-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "time-suffix-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "2026_03_15", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "underscores-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "underscores-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "underscores-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "20X6-03-15", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "non-digit-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "non-digit-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "non-digit-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "-", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "sign-only-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "sign-only-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "sign-only-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "46,096", 0#, cdDateTextFormat
        TEST_Harness.AssertTrue "thousands-separator-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "thousands-separator-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "thousands-separator-value", mFact(3, 1), mFact(3, 2)
    'Text that reads as a number is a serial written as text
        ProbeDate "46096", 0#, cdDateTextNumeric
        TEST_Harness.AssertTrue "serial-digits-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "serial-digits-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "serial-digits-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "46096.75", 0#, cdDateTextNumeric
        TEST_Harness.AssertTrue "serial-decimal-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "serial-decimal-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "serial-decimal-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "20260315", 0#, cdDateTextNumeric
        TEST_Harness.AssertTrue "no-separators-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "no-separators-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "no-separators-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "31.12", 0#, cdDateTextNumeric
        TEST_Harness.AssertTrue "dotted-numeric-shape-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "dotted-numeric-shape-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "dotted-numeric-shape-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "-1.", 0#, cdDateTextNumeric
        TEST_Harness.AssertTrue "signed-trailing-point-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "signed-trailing-point-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "signed-trailing-point-value", mFact(3, 1), mFact(3, 2)
    'A well-formed date that the calendar lacks is impossible
        ProbeDate "2024-02-30", 0#, cdDateTextImpossible
        TEST_Harness.AssertTrue "feb-30-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "feb-30-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "feb-30-value", mFact(3, 1), mFact(3, 2)
        ProbeDate "2026-03-00", 0#, cdDateTextImpossible
        TEST_Harness.AssertTrue "day-00-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "day-00-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "day-00-value", mFact(3, 1), mFact(3, 2)
    'The last serial with a time of day is accepted; serial 59 is outside the window
        ProbeDate 2958465.999, 2958465#, cdNone
        TEST_Harness.AssertTrue "max-serial-with-time-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "max-serial-with-time-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "max-serial-with-time-value", mFact(3, 1), mFact(3, 2)
        ProbeDate 59#, 0#, cdDateWindow
        TEST_Harness.AssertTrue "serial-59-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "serial-59-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "serial-59-value", mFact(3, 1), mFact(3, 2)

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

Private Sub CaseLongRegression()

'
'------------------------------------------------------------------------------
'                     CHECK THE REGRESSION INTEGER INPUTS
'------------------------------------------------------------------------------
' PURPOSE
'   Regression integer inputs that the cases above do not hold. 51 assertions.
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
'   Parses each input and asserts its result, condition and value
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Err, CORE_Parse, TEST_Harness
'
' NOTES
'   Each expectation is the data contract's; none is read back from the code under test
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_CoreErrParse.CaseLongRegression"

    Dim DecimalOutside  As Variant      'A Decimal 1.000000001E-9 above a whole number

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "long-regression"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Decimal keeps the offset exactly, which Double would round
        DecimalOutside = CDec(3) + CDec(1000000001) / CDec(1000000000000000000#)
    'Whole numbers of every numeric type
        ProbeLong CByte(12), 12, cdNone
        TEST_Harness.AssertTrue "byte-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "byte-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "byte-value", mFact(3, 1), mFact(3, 2)
        ProbeLong CInt(12), 12, cdNone
        TEST_Harness.AssertTrue "integer-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "integer-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "integer-value", mFact(3, 1), mFact(3, 2)
        ProbeLong CLng(12), 12, cdNone
        TEST_Harness.AssertTrue "long-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "long-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "long-value", mFact(3, 1), mFact(3, 2)
        ProbeLong CDbl(12), 12, cdNone
        TEST_Harness.AssertTrue "double-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "double-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "double-value", mFact(3, 1), mFact(3, 2)
        ProbeLong CCur(12), 12, cdNone
        TEST_Harness.AssertTrue "currency-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "currency-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "currency-value", mFact(3, 1), mFact(3, 2)
        ProbeLong 0, 0, cdNone
        TEST_Harness.AssertTrue "zero-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "zero-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "zero-value", mFact(3, 1), mFact(3, 2)
        ProbeLong -12, -12, cdNone
        TEST_Harness.AssertTrue "negative-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "negative-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "negative-value", mFact(3, 1), mFact(3, 2)
    'A binary residue below a whole number is within the tolerance
        ProbeLong (0.7 + 0.1) * 10#, 8, cdNone
        TEST_Harness.AssertTrue "residue-below-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "residue-below-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "residue-below-value", mFact(3, 1), mFact(3, 2)
    'A Decimal just outside the 1E-9 tolerance, and plain fractions, are fractions
        ProbeLong DecimalOutside, 0, cdIntegerFraction
        TEST_Harness.AssertTrue "decimal-outside-tolerance-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "decimal-outside-tolerance-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "decimal-outside-tolerance-value", mFact(3, 1), mFact(3, 2)
        ProbeLong 2.6, 0, cdIntegerFraction
        TEST_Harness.AssertTrue "fraction-up-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "fraction-up-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "fraction-up-value", mFact(3, 1), mFact(3, 2)
        ProbeLong 2.4, 0, cdIntegerFraction
        TEST_Harness.AssertTrue "fraction-down-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "fraction-down-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "fraction-down-value", mFact(3, 1), mFact(3, 2)
        ProbeLong 3.0000001, 0, cdIntegerFraction
        TEST_Harness.AssertTrue "tiny-fraction-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "tiny-fraction-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "tiny-fraction-value", mFact(3, 1), mFact(3, 2)
    'A fractional value above the Long range is out of range
        ProbeLong 2147483648.5, 0, cdIntegerRange
        TEST_Harness.AssertTrue "over-range-fractional-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "over-range-fractional-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "over-range-fractional-value", mFact(3, 1), mFact(3, 2)
    'False, numeric text, other text and a Date are rejected types
        ProbeLong False, 0, cdIntegerTypeRejected
        TEST_Harness.AssertTrue "boolean-false-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "boolean-false-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "boolean-false-value", mFact(3, 1), mFact(3, 2)
        ProbeLong "12", 0, cdIntegerTypeRejected
        TEST_Harness.AssertTrue "numeric-text-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "numeric-text-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "numeric-text-value", mFact(3, 1), mFact(3, 2)
        ProbeLong "twelve", 0, cdIntegerTypeRejected
        TEST_Harness.AssertTrue "text-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "text-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "text-value", mFact(3, 1), mFact(3, 2)
        ProbeLong DateSerial(2026, 3, 15), 0, cdIntegerTypeRejected
        TEST_Harness.AssertTrue "date-result", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "date-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "date-value", mFact(3, 1), mFact(3, 2)

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

Private Sub ProbeCondition( _
    ByVal Identifier As String, _
    ByVal Condition As KPRCondition, _
    ByVal ExpectedCode As Long)
'
'------------------------------------------------------------------------------
'                             PROBE ONE CONDITION
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
        mFact(1, 1) = Identifier
        mFact(1, 2) = CORE_Err.ConditionName(Condition)
    'The error is the one the data contract's Conditions table gives
        mFact(2, 1) = ExpectedCode
        mFact(2, 2) = CORE_Err.ErrForCondition(Condition)

End Sub

Private Sub ProbeName( _
    ByVal Condition As KPRCondition, _
    ByVal Expected As String)
'
'------------------------------------------------------------------------------
'                           PROBE ONE CONDITION NAME
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
        mFact(1, 1) = Expected
        mFact(1, 2) = CORE_Err.ConditionName(Condition)

End Sub

Private Sub ProbeUnmapped( _
    ByVal Condition As KPRCondition)
'
'------------------------------------------------------------------------------
'                         PROBE AN UNMAPPED CONDITION
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
    Const PROC_NAME     As String = "TEST_CoreErrParse.ProbeUnmapped"

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
' PROBE ERROR
'------------------------------------------------------------------------------
    'The error is the one internal defect, with its documented source and description
        mFact(1, 1) = CORE_Err.ERR_UNMAPPED_CONDITION
        mFact(1, 2) = RaisedNumber
        mFact(2, 1) = "CORE_Err.ErrForCondition"
        mFact(2, 2) = RaisedSource
        mFact(3, 1) = "Condition has no worksheet error mapping; this is an internal defect."
        mFact(3, 2) = RaisedText

End Sub

Private Sub ProbeDate( _
    ByVal ScalarIn As Variant, _
    ByVal ExpectedSerial As Double, _
    ByVal ExpectedCondition As KPRCondition)
'
'------------------------------------------------------------------------------
'                            PROBE ONE DATE SCALAR
'------------------------------------------------------------------------------
' PURPOSE
'   Parses one scalar as a date and checks the result, the condition and the output.
'   3 assertions
'
' WHY THIS EXISTS
'   Every date case checks the same three facts, so they are stated once
'
' INPUTS
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
    Const PROC_NAME     As String = "TEST_CoreErrParse.ProbeDate"

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
' PROBE RESULT
'------------------------------------------------------------------------------
    'The result, the condition and the output are three separate facts
        mFact(1, 1) = (Accepted = (ExpectedCondition = cdNone))
        mFact(1, 2) = "the acceptance differs from the expectation"
        mFact(2, 1) = CORE_Err.ConditionName(ExpectedCondition)
        mFact(2, 2) = CORE_Err.ConditionName(Condition)
        mFact(3, 1) = ExpectedValue
        mFact(3, 2) = VBA.CDbl(ParsedDate)

End Sub

Private Sub ProbeLong( _
    ByVal ScalarIn As Variant, _
    ByVal ExpectedValue As Long, _
    ByVal ExpectedCondition As KPRCondition)
'
'------------------------------------------------------------------------------
'                           PROBE ONE INTEGER SCALAR
'------------------------------------------------------------------------------
' PURPOSE
'   Parses one scalar as a Long and checks the result, the condition and the output.
'   3 assertions
'
' WHY THIS EXISTS
'   Every integer case checks the same three facts, so they are stated once
'
' INPUTS
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
    Const PROC_NAME     As String = "TEST_CoreErrParse.ProbeLong"

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
' PROBE RESULT
'------------------------------------------------------------------------------
    'The result, the condition and the output are three separate facts
        mFact(1, 1) = (Accepted = (ExpectedCondition = cdNone))
        mFact(1, 2) = "the acceptance differs from the expectation"
        mFact(2, 1) = CORE_Err.ConditionName(ExpectedCondition)
        mFact(2, 2) = CORE_Err.ConditionName(Condition)
        mFact(3, 1) = ExpectedOutput
        mFact(3, 2) = ParsedLong

End Sub

Private Sub ProbeBool( _
    ByVal ControlIn As Variant, _
    ByVal DefaultValue As Boolean, _
    ByVal ExpectedValue As Boolean, _
    ByVal ExpectedCondition As KPRCondition)
'
'------------------------------------------------------------------------------
'                          PROBE ONE BOOLEAN CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Parses one control as a Boolean and checks the result, the condition and the output.
'   3 assertions
'
' WHY THIS EXISTS
'   Every Boolean control case checks the same three facts, so they are stated once
'
' INPUTS
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
    Const PROC_NAME     As String = "TEST_CoreErrParse.ProbeBool"

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
' PROBE RESULT
'------------------------------------------------------------------------------
    'The result, the condition and the output are three separate facts
        mFact(1, 1) = (Accepted = (ExpectedCondition = cdNone))
        mFact(1, 2) = "the acceptance differs from the expectation"
        mFact(2, 1) = CORE_Err.ConditionName(ExpectedCondition)
        mFact(2, 2) = CORE_Err.ConditionName(Condition)
        mFact(3, 1) = (ParsedBool = ExpectedOutput)
        mFact(3, 2) = "the output differs from the expectation"

End Sub

Private Sub ProbeRounding( _
    ByVal ControlIn As Variant, _
    ByVal ExpectedToken As String, _
    ByVal ExpectedCondition As KPRCondition)
'
'------------------------------------------------------------------------------
'                          PROBE ONE ROUNDING CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Parses one control as a rounding token and checks the result, the condition and the
'   output. 3 assertions
'
' WHY THIS EXISTS
'   Every rounding case checks the same three facts, so they are stated once
'
' INPUTS
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
    Const PROC_NAME     As String = "TEST_CoreErrParse.ProbeRounding"

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
' PROBE RESULT
'------------------------------------------------------------------------------
    'The result, the condition and the output are three separate facts
        mFact(1, 1) = (Accepted = (ExpectedCondition = cdNone))
        mFact(1, 2) = "the acceptance differs from the expectation"
        mFact(2, 1) = CORE_Err.ConditionName(ExpectedCondition)
        mFact(2, 2) = CORE_Err.ConditionName(Condition)
        mFact(3, 1) = ExpectedOutput
        mFact(3, 2) = TokenOut

End Sub
