Attribute VB_Name = "TEST_Examples"

Option Explicit
Option Private Module

'
'==============================================================================
' MODULE: TEST_Examples
'==============================================================================
' PURPOSE
'   Neutral example suite of TEST_Harness: one case per kind of assertion,
'   one case that records the results of the illustrative test cases, and
'   the reference implementation of the cleanup pattern of
'   docs/VBA_HOUSE_STYLE.md (Operation and cleanup outcomes) with its
'   failure-path cases. Nothing here is a calculation of the workbook or a
'   regulatory value
'
' WHY THIS EXISTS
'   An Excel run of this suite in its three modes is the evidence that the
'   harness itself works: that a passing run passes, and that a failed or
'   incomplete run is never reported as a pass
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   RunExampleTests is the documented test entry point.
'   RunExampleTestsWithInjectedFailure and RunExampleTestsWithSkippedCase
'   show what a failed and an incomplete run report. CertifyExamples
'   certifies all three and saves their logs, the certification log and the
'   results of the normal run as the evidence of an Excel run. Each takes
'   the folder of the test cases, tests/fixtures/ of the checkout.
'   Option Private Module keeps them out of the workbook's external surface.
'
'   The module owns the mode and the fixtures folder of the running suite,
'   and reads the illustrative test cases of tests/fixtures/.
'   RunExampleOperation changes Application.Calculation and
'   Application.ScreenUpdating and restores them before it returns, on the
'   success and the error path
'
' ERROR POLICY
'   As TEST_Harness.EndRun: a run that does not pass raises ERR_TEST_FAILED,
'   or the unexpected error that escaped the runner. A refused start raises
'   TEST_Harness.ERR_TEST_DIRTY_START. CertifyExamples records those errors
'   instead of raising them, saves every log, and raises only when the
'   certification did not pass
'
' DEPENDENCIES
'   TEST_Harness; the VBA and Excel object models
'
' NOTES
'   In a development workbook holding the modules of tests/modules/, type
'   TEST_Examples.RunExampleTests "<checkout>\tests\fixtures" in the
'   Immediate window. To record an Excel run, type
'   TEST_Examples.CertifyExamples "<folder>", "<checkout>\tests\fixtures"
'   instead (INSTALLATION.md, Recording an Excel run)
'
' UPDATED
'   2026-11-04
'==============================================================================

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'A complete run starts every case and evaluates every assertion once
    Private Const SUITE_NAME            As String = "examples"  'Suite name in the run log
    Private Const EXPECTED_CASES        As Long = 14            'Cases a complete run starts
    Private Const EXPECTED_ASSERTIONS   As Long = 34            'Assertions a complete run evaluates

'------------------------------------------------------------------------------
' ERROR NUMBERS
'------------------------------------------------------------------------------
    'Errors of the example procedures, test-only, as literals equal to
    'vbObjectError + n, n first in each comment (see TEST_Harness, Module
    'constants)
    Private Const ERR_EXAMPLE_INVALID       As Long = &H80040FF0    '4080: Input outside the domain
    Private Const ERR_EXAMPLE_OPERATION     As Long = &H80040FF1    '4081: Primary error of the operation
    Private Const ERR_EXAMPLE_RESTORE       As Long = &H80040FF2    '4082: Simulated restore failure
    Private Const ERR_CLEANUP_FAILED        As Long = &H80040FF3    '4083: Only the cleanup failed
    Private Const ERR_EXAMPLE_HOST          As Long = &H80040FF4    '4084: Simulated capture or change failure
    Private Const ERR_EXAMPLE_FIXTURE       As Long = &H80040FF5    '4085: Test case file not as expected

'------------------------------------------------------------------------------
' RESULT QUANTITIES
'------------------------------------------------------------------------------
    'The quantity of the illustrative test cases (tests/expected/)
    Private Const ILLUSTRATIVE_QUANTITY     As String = "illustrative_total_notional_amount"    'Recorded total

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private mMode       As String   'NORMAL, INJECTED_FAILURE or SKIPPED_CASE
    Private mFixtures   As String   'Folder of the test cases, tests/fixtures/


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub RunExampleTests( _
    ByVal Fixtures As String)
'
'------------------------------------------------------------------------------
'                              RUN EXAMPLE TESTS
'------------------------------------------------------------------------------
' PURPOSE
'   Run the example suite. A passing run ends with
'   RESULT=PASS; suite=examples; operation=PASS; completeness=COMPLETE;
'   cases=14; assertions=34; failures=0; cleanup=PASS; recovery=NONE
'
' WHY THIS EXISTS
'   It is the documented test entry point of the suite
'
' INPUTS
'   Fixtures
'     The full path of tests/fixtures/ in the checkout, without a trailing
'     separator
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs RunSuite in the NORMAL mode
'
' ERROR POLICY
'   As RunSuite: a run that does not pass raises
'   TEST_Harness.ERR_TEST_FAILED after it has been reported
'
' DEPENDENCIES
'   RunSuite
'
' NOTES
'   None
'
' UPDATED
'   2026-11-04
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN SUITE
'------------------------------------------------------------------------------
    'Run the suite in its normal mode
        RunSuite "NORMAL", Fixtures

End Sub

Public Sub RunExampleTestsWithInjectedFailure( _
    ByVal Fixtures As String)
'
'------------------------------------------------------------------------------
'                  RUN EXAMPLE TESTS WITH AN INJECTED FAILURE
'------------------------------------------------------------------------------
' PURPOSE
'   Show the failure path: exact-double expects a wrong value. Expected
'   end: RESULT=FAIL; ...; operation=FAIL; completeness=COMPLETE; cases=14;
'   assertions=34; failures=1, then the error ERR_TEST_FAILED
'
' WHY THIS EXISTS
'   It shows what a failed run reports
'
' INPUTS
'   Fixtures
'     As RunExampleTests
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs RunSuite in the INJECTED_FAILURE mode, in which CaseExactDouble
'   expects a wrong value
'
' ERROR POLICY
'   Raises TEST_Harness.ERR_TEST_FAILED after the failed run has been
'   reported
'
' DEPENDENCIES
'   RunSuite
'
' NOTES
'   None
'
' UPDATED
'   2026-11-04
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN SUITE
'------------------------------------------------------------------------------
    'Run the suite with one assertion made to fail
        RunSuite "INJECTED_FAILURE", Fixtures

End Sub

Public Sub RunExampleTestsWithSkippedCase( _
    ByVal Fixtures As String)
'
'------------------------------------------------------------------------------
'                    RUN EXAMPLE TESTS WITH A SKIPPED CASE
'------------------------------------------------------------------------------
' PURPOSE
'   Show that an incomplete run is never a pass: the last case is skipped.
'   Expected end: RESULT=FAIL; ...; operation=PASS;
'   completeness=INCOMPLETE; cases=13; assertions=32; failures=1, the
'   failure being harness.completeness, then ERR_TEST_FAILED
'
' WHY THIS EXISTS
'   It shows what an incomplete run reports
'
' INPUTS
'   Fixtures
'     As RunExampleTests
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs RunSuite in the SKIPPED_CASE mode, in which RunSuite does not call
'   CaseCleanupOnlyFailure
'
' ERROR POLICY
'   Raises TEST_Harness.ERR_TEST_FAILED after the incomplete run has been
'   reported
'
' DEPENDENCIES
'   RunSuite
'
' NOTES
'   None
'
' UPDATED
'   2026-11-04
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN SUITE
'------------------------------------------------------------------------------
    'Run the suite with its last case skipped
        RunSuite "SKIPPED_CASE", Fixtures

End Sub

Public Sub CertifyExamples( _
    ByVal Folder As String, _
    ByVal Fixtures As String)
'
'------------------------------------------------------------------------------
'                          CERTIFY THE EXAMPLE SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   The certification entry point: refuse state left by an earlier run,
'   run the suite in its three modes, save each complete log in Folder
'   (examples.log, injected.log and skipped.log) and the results of the
'   normal run (results.json), then save certification.log
'
' WHY THIS EXISTS
'   The saved logs are the evidence of an Excel run, which
'   tools/check_excel_run.py checks against the commit
'
' INPUTS
'   Folder
'     An existing folder outside the checkout or in the ignored
'     test-results/ directory, without a trailing backslash
'
'   Fixtures
'     The full path of tests/fixtures/ in the same checkout, without a
'     trailing separator
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   The normal run must pass and be complete. The injected-failure and
'   skipped-case runs must fail and reach this caller as
'   TEST_Harness.ERR_TEST_FAILED with a description: they are the
'   self-test that a failed run is never silent
'
' ERROR POLICY
'   The error that ends each run is recorded, not raised, so that every log
'   is written. EndCertification raises ERR_TEST_FAILED after writing
'   certification.log when the certification did not pass. A refused start
'   raises, before any run, TEST_Harness.ERR_TEST_DIRTY_START when a run is
'   active or recovery is required, or TEST_Harness.ERR_TEST_USAGE while an
'   earlier certification is still active. An error of SaveLog or
'   SaveResults propagates; run TEST_Harness.ResetTests before trying again
'
' DEPENDENCIES
'   CertifyMode; TEST_Harness.BeginCertification, TEST_Harness.SaveResults
'   and TEST_Harness.EndCertification
'
' NOTES
'   tools/check_excel_run.py checks the folder against the commit
'
' UPDATED
'   2026-11-04
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CERTIFY
'------------------------------------------------------------------------------
    'Refuse state left by an earlier run and start the certification
        TEST_Harness.BeginCertification SUITE_NAME
    'Run the normal mode, which must pass, then save its results before the next
    'mode resets them, and run the two failure modes, which must fail
        CertifyMode "NORMAL", True, Folder & "\examples.log", Fixtures
        TEST_Harness.SaveResults Folder & "\results.json"
        CertifyMode "INJECTED_FAILURE", False, Folder & "\injected.log", Fixtures
        CertifyMode "SKIPPED_CASE", False, Folder & "\skipped.log", Fixtures
    'Save the certification log, raising when the certification did not pass
        TEST_Harness.EndCertification Folder & "\certification.log"

End Sub

'
'------------------------------------------------------------------------------
'
'                                    RUNNER
'
'------------------------------------------------------------------------------
'

Private Sub RunSuite( _
    ByVal Mode As String, _
    ByVal Fixtures As String)
'
'------------------------------------------------------------------------------
'                          RUN THE SUITE IN ONE MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case in a fixed order between BeginRun and EndRun
'
' WHY THIS EXISTS
'   The three entry points and the certification share one runner, so that
'   every mode runs the same cases
'
' INPUTS
'   Mode
'     NORMAL, INJECTED_FAILURE or SKIPPED_CASE
'
'   Fixtures
'     The folder of the test cases, as RunExampleTests
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Stores the mode and the fixtures folder for the cases, starts the run,
'   calls every case (CaseCleanupOnlyFailure only outside SKIPPED_CASE),
'   clears the stored mode and folder, then ends the run
'
' ERROR POLICY
'   An error escaping a case handler is passed to EndRun, which reports it
'   and re-raises it unchanged. BeginRun's refusal propagates directly
'
' DEPENDENCIES
'   TEST_Harness.BeginRun, TEST_Harness.EndRun; the Case procedures
'
' NOTES
'   None
'
' UPDATED
'   2026-11-04
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrNumber       As Long     'Runner error number; 0 for none
    Dim ErrSource       As String   'Runner error source
    Dim ErrDescription  As String   'Runner error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Store the mode and the fixtures folder for the cases
        mMode = Mode
        mFixtures = Fixtures
    'Start the run; a refused start propagates directly
        TEST_Harness.BeginRun SUITE_NAME, Mode, EXPECTED_CASES, EXPECTED_ASSERTIONS
    'Send an error that escapes a case to the runner handler
        On Error GoTo RunFailed

'------------------------------------------------------------------------------
' RUN CASES
'------------------------------------------------------------------------------
    'Run every case in a fixed order
        CaseExactDouble
        CaseAbsoluteTolerance
        CaseRelativeTolerance
        CaseExactText
        CaseErrorValue
        CaseRaisedError
        CaseIllustrativeResults
        CaseCleanupRestoresState
        CaseCleanupKeepsPrimaryError
        CaseCleanupPartialCapture
        CaseCleanupChangeFailure
        CaseCleanupPrimaryOnly
        CaseCleanupEachRestorationFailure
    'Skip the last case in the SKIPPED_CASE mode
        If mMode <> "SKIPPED_CASE" Then
            CaseCleanupOnlyFailure
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
Finish:
    'Restore normal error handling
        On Error GoTo 0
    'Clear the stored mode and fixtures folder
        mMode = vbNullString
        mFixtures = vbNullString
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

Private Sub CertifyMode( _
    ByVal Mode As String, _
    ByVal MustPass As Boolean, _
    ByVal Path As String, _
    ByVal Fixtures As String)
'
'------------------------------------------------------------------------------
'                           RUN AND CERTIFY ONE MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Run the suite in Mode, save its log to Path, then certify how the run
'   reached this caller
'
' WHY THIS EXISTS
'   Each mode is certified the same way, after its log is saved
'
' INPUTS
'   Mode
'     NORMAL, INJECTED_FAILURE or SKIPPED_CASE
'
'   MustPass
'     True for the normal run, False for a deliberate failure mode
'
'   Path
'     The file the run's log is saved to
'
'   Fixtures
'     The folder of the test cases, as RunExampleTests
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs the suite, captures the error that ended it, saves the log, prints
'   a SAVED= line, then passes the outcome to TEST_Harness.CertifyRun
'
' ERROR POLICY
'   The run's own error is captured, printed and certified after the log is
'   saved; an error of SaveLog propagates
'
' DEPENDENCIES
'   RunSuite; TEST_Harness.SaveLog, TEST_Harness.Echo and
'   TEST_Harness.CertifyRun
'
' NOTES
'   None
'
' UPDATED
'   2026-11-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RunErrorNumber          As Long     'Error that ended the run; 0 for none
    Dim RunErrorDescription     As String   'Its description

'------------------------------------------------------------------------------
' RUN SUITE
'------------------------------------------------------------------------------
    'Run the suite and capture the error that ended it
        On Error Resume Next
        RunSuite Mode, Fixtures
        RunErrorNumber = Err.Number
        RunErrorDescription = Err.Description
    'Restore normal error handling
        On Error GoTo 0

'------------------------------------------------------------------------------
' SAVE AND CERTIFY
'------------------------------------------------------------------------------
    'Save the log and print where it went
        TEST_Harness.SaveLog Path
        TEST_Harness.Echo "SAVED=" & Path & "; mode=" & Mode & "; error=" & CStr(RunErrorNumber) & _
            IIf(RunErrorNumber = 0, vbNullString, " " & RunErrorDescription)
    'Certify how the run reached this caller
        TEST_Harness.CertifyRun Mode, MustPass, RunErrorNumber, RunErrorDescription

End Sub

'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseExactDouble()

'
'------------------------------------------------------------------------------
'                         CHECK EXACT DOUBLE EQUALITY
'------------------------------------------------------------------------------
' PURPOSE
'   Exact equality on a result that is exact in binary. 1 assertion.
'
' WHY THIS EXISTS
'   It shows the one kind of result that may be compared without a
'   tolerance, and it is the case the INJECTED_FAILURE mode makes fail
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Compares 0.5 + 0.25 with 0.75, or with 0.76 in the INJECTED_FAILURE
'   mode
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Expected    As Double   'Expected sum; wrong on purpose when injected

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "exact-double"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Expect the exact sum, or a wrong value when the failure is injected
        Expected = 0.75
        If mMode = "INJECTED_FAILURE" Then
            Expected = 0.76
        End If
    'Compare exactly
        TEST_Harness.AssertEqualDouble "half-plus-quarter", Expected, 0.5 + 0.25

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

Private Sub CaseAbsoluteTolerance()

'
'------------------------------------------------------------------------------
'                         CHECK THE ABSOLUTE TOLERANCE
'------------------------------------------------------------------------------
' PURPOSE
'   The absolute bound: 0.1 + 0.2 is not exactly 0.3 in binary, and near
'   zero only the absolute bound counts. 2 assertions.
'
' WHY THIS EXISTS
'   A sum that is not exact in binary needs a tolerance, and a relative
'   bound is meaningless near zero
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Compares 0.1 + 0.2 with 0.3, then a value near zero with zero, each
'   within its bounds
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "tolerance-absolute"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Compare within the absolute bound, away from and near zero
        TEST_Harness.AssertNear "tenth-plus-fifth", 0.3, 0.1 + 0.2, 0.000000000000001, 0#
        TEST_Harness.AssertNear "near-zero", 0#, 1E-20, 0.000000000001, 0.000000001

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

Private Sub CaseRelativeTolerance()

'
'------------------------------------------------------------------------------
'                         CHECK THE RELATIVE TOLERANCE
'------------------------------------------------------------------------------
' PURPOSE
'   The relative bound governs large values, the bound is inclusive, and a
'   value outside both bounds is rejected. 3 assertions.
'
' WHY THIS EXISTS
'   Large values need a bound proportional to their size
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Compares a large value within its relative bound, then checks that a
'   difference equal to the bound is accepted and that one outside both
'   bounds is rejected
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "tolerance-relative"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Compare a large value within the relative bound
        TEST_Harness.AssertNear "large-value", 1000000000#, 1000000000.1, 0.000001, 0.000000001
    'Check that the bound is inclusive
        TEST_Harness.AssertTrue "bound-is-inclusive", _
            TEST_Harness.WithinTolerance(2#, 2.5, 0.5, 0#), "|2.5 - 2| = 0.5 must be within 0.5"
    'Check that a value outside both bounds is rejected
        TEST_Harness.AssertTrue "outside-both-bounds", _
            Not TEST_Harness.WithinTolerance(1#, 1.1, 0.01, 0.01), _
            "|1.1 - 1| = 0.1 must exceed max(0.01, 0.01 * 1)"

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

Private Sub CaseExactText()

'
'------------------------------------------------------------------------------
'                          CHECK EXACT TEXT EQUALITY
'------------------------------------------------------------------------------
' PURPOSE
'   Text is compared exactly, case included. 2 assertions.
'
' WHY THIS EXISTS
'   Identifiers and codes are compared as text, where case matters
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Compares a concatenation with its expected text, then checks that a
'   binary comparison tells "a" from "A"
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "text-exact"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Compare a built text exactly
        TEST_Harness.AssertEqualText "concatenation", "CASE-1", "CASE" & "-" & CStr(1)
    'Check that case matters
        TEST_Harness.AssertTrue "case-matters", _
            StrComp("a", "A", vbBinaryCompare) <> 0, "a and A must differ"

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

Private Sub CaseErrorValue()

'
'------------------------------------------------------------------------------
'                         CHECK A RETURNED ERROR VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   An expected worksheet error value, returned rather than raised.
'   1 assertion.
'
' WHY THIS EXISTS
'   A worksheet function returns its errors as values, which need their own
'   assertion
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Checks that ExampleRatio returns #DIV/0! for a zero denominator
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   ExampleRatio, TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "error-value"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Check the error value of a zero denominator
        TEST_Harness.AssertErrorValue "zero-denominator", xlErrDiv0, ExampleRatio(1#, 0#)

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

Private Sub CaseRaisedError()

'
'------------------------------------------------------------------------------
'                             CHECK A RAISED ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   An expected raised error, captured on the line after the call, and the
'   same call with a valid input raising nothing. 2 assertions.
'
' WHY THIS EXISTS
'   A core procedure raises its errors, and a test must capture Err.Number
'   before any other call can reset it
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls ExampleSquareRoot with a negative and with a valid input, each
'   under On Error Resume Next, and checks the error number of each call
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   ExampleSquareRoot, TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result          As Double   'Value of the checked square root
    Dim RaisedError     As Long     'Err.Number right after the call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "raised-error"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Capture the error of a negative input on the line after the call
        On Error Resume Next
        Result = ExampleSquareRoot(-1#)
        RaisedError = Err.Number
        On Error GoTo CaseFailed
    'Check the expected error
        TEST_Harness.AssertRaised "negative-input", ERR_EXAMPLE_INVALID, RaisedError
    'Capture the error of a valid input on the line after the call
        On Error Resume Next
        Result = ExampleSquareRoot(4#)
        RaisedError = Err.Number
        On Error GoTo CaseFailed
    'Check that nothing was raised
        TEST_Harness.AssertRaised "valid-input", 0, RaisedError

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

Private Sub CaseIllustrativeResults()

'
'------------------------------------------------------------------------------
'                       RECORD THE ILLUSTRATIVE RESULTS
'------------------------------------------------------------------------------
' PURPOSE
'   The results of an Excel run: read the illustrative test cases from the
'   fixtures folder, add up the notional amounts of each, and record each
'   total with TEST_Harness.RecordResult, for the accuracy contract ACC-01
'   (tests/accuracy/contracts.json). tools/check_accuracy.py compares the
'   totals with their expected values; this case checks only that each file
'   held amounts. 3 assertions.
'
' WHY THIS EXISTS
'   The accuracy check needs results computed in Excel, recorded in the
'   results file of the normal run
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   For the JSON format example and for each line-ending variant of the
'   file import example: totals the notional amounts, checks that the file
'   held amounts, and records the total
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   FixturePath, SumJsonNumbers, SumCsvColumn, TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Total       As Double   'Sum of the notional amounts of one file
    Dim Records     As Long     'Amounts read from that file

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "illustrative-results"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Total, check and record the JSON format example
        Total = SumJsonNumbers(FixturePath("illustrative-format-example.json"), _
            "notional_amount", Records)
        TEST_Harness.AssertTrue "format-example-amounts", Records > 0, _
            "illustrative-format-example.json holds no notional_amount"
        TEST_Harness.RecordResult "illustrative-format-example", vbNullString, _
            ILLUSTRATIVE_QUANTITY, vbNullString, vbNullString, Total
    'Total, check and record the CR LF variant of the file import example
        Total = SumCsvColumn(FixturePath( _
            "illustrative-file-import-example/crlf-line-endings/positions.csv"), "notional", Records)
        TEST_Harness.AssertTrue "crlf-variant-amounts", Records > 0, _
            "the crlf-line-endings positions.csv holds no notional amount"
        TEST_Harness.RecordResult "illustrative-file-import-example", "crlf-line-endings", _
            ILLUSTRATIVE_QUANTITY, vbNullString, vbNullString, Total
    'Total, check and record the LF variant of the file import example
        Total = SumCsvColumn(FixturePath( _
            "illustrative-file-import-example/lf-line-endings/positions.csv"), "notional", Records)
        TEST_Harness.AssertTrue "lf-variant-amounts", Records > 0, _
            "the lf-line-endings positions.csv holds no notional amount"
        TEST_Harness.RecordResult "illustrative-file-import-example", "lf-line-endings", _
            ILLUSTRATIVE_QUANTITY, vbNullString, vbNullString, Total

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

Private Sub CaseCleanupRestoresState()

'
'------------------------------------------------------------------------------
'                      CHECK THAT CLEANUP RESTORES STATE
'------------------------------------------------------------------------------
' PURPOSE
'   A successful operation restores what it changed and reports no cleanup
'   failure. 2 assertions.
'
' WHY THIS EXISTS
'   It is the success path of the cleanup pattern
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs RunExampleOperation with no failure, then checks that it reported
'   no cleanup failure and that ScreenUpdating is as before
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   RunExampleOperation, TEST_Harness; the Excel object model
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details         As String   'Cleanup failures reported by the operation
    Dim SavedScreen     As Boolean  'ScreenUpdating before the operation

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "cleanup-restores-state"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Run the operation with no failure, keeping ScreenUpdating to compare
        SavedScreen = Application.ScreenUpdating
        RunExampleOperation vbNullString, False, False, Details
    'Check that nothing failed and the setting was restored
        TEST_Harness.AssertEqualText "no-cleanup-failure", vbNullString, Details
        TEST_Harness.AssertTrue "screen-updating-restored", _
            Application.ScreenUpdating = SavedScreen, "ScreenUpdating differs after the operation"

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

Private Sub CaseCleanupKeepsPrimaryError()

'
'------------------------------------------------------------------------------
'                  CHECK THAT CLEANUP KEEPS THE PRIMARY ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   When the operation and its cleanup both fail, the caller receives the
'   primary error, and the cleanup failure is still reported. 3 assertions.
'
' WHY THIS EXISTS
'   Cleanup must never replace the primary error, and its failures must
'   still be reported
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs RunExampleOperation failing in its work and in both restorations,
'   then checks the raised error and that both restoration failures were
'   reported
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   RunExampleOperation, TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details         As String   'Cleanup failures reported by the operation
    Dim RaisedError     As Long     'Err.Number right after the call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "cleanup-keeps-primary-error"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Run the operation failing everywhere and capture its error
        On Error Resume Next
        RunExampleOperation "work", True, True, Details
        RaisedError = Err.Number
        On Error GoTo CaseFailed
    'Check the primary error and both reported restoration failures
        TEST_Harness.AssertRaised "primary-error", ERR_EXAMPLE_OPERATION, RaisedError
        TEST_Harness.AssertTrue "calculation-failure-reported", _
            InStr(Details, "Calculation not restored") > 0, "details='" & Details & "'"
        TEST_Harness.AssertTrue "screen-failure-reported", _
            InStr(Details, "ScreenUpdating not restored") > 0, "details='" & Details & "'"

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

Private Sub CaseCleanupPartialCapture()

'
'------------------------------------------------------------------------------
'                    CHECK CLEANUP AFTER A PARTIAL CAPTURE
'------------------------------------------------------------------------------
' PURPOSE
'   When capture fails part-way, nothing is restored: no setting was
'   changed, and an uncaptured value is never restored. Both restore seams
'   are set, so any restoration attempt would be reported. 2 assertions.
'
' WHY THIS EXISTS
'   An uncaptured setting must never be restored
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs RunExampleOperation failing while it captures ScreenUpdating, then
'   checks the raised error and that no restoration was reported
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   RunExampleOperation, TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details         As String   'Cleanup failures reported by the operation
    Dim RaisedError     As Long     'Err.Number right after the call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "cleanup-partial-capture"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Run the operation failing during capture and capture its error
        On Error Resume Next
        RunExampleOperation "capture-screen", True, True, Details
        RaisedError = Err.Number
        On Error GoTo CaseFailed
    'Check the capture error and that nothing was restored
        TEST_Harness.AssertRaised "capture-error", ERR_EXAMPLE_HOST, RaisedError
        TEST_Harness.AssertEqualText "nothing-restored", vbNullString, Details

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

Private Sub CaseCleanupChangeFailure()

'
'------------------------------------------------------------------------------
'                     CHECK CLEANUP AFTER A FAILED CHANGE
'------------------------------------------------------------------------------
' PURPOSE
'   When a change fails after its restoration was registered, that setting
'   is still restored, as is the one changed before it, and the change's
'   error stays the primary error. 3 assertions.
'
' WHY THIS EXISTS
'   A setter that fails half way must still be restored
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs RunExampleOperation failing while it changes Calculation, then
'   checks the raised error, that the Calculation restoration was
'   attempted, and that ScreenUpdating is as before
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   RunExampleOperation, TEST_Harness; the Excel object model
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details         As String   'Cleanup failures reported by the operation
    Dim RaisedError     As Long     'Err.Number right after the call
    Dim SavedScreen     As Boolean  'ScreenUpdating before the operation

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "cleanup-change-failure"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Keep ScreenUpdating to compare after the operation
        SavedScreen = Application.ScreenUpdating
    'Run the operation failing during the change and capture its error
        On Error Resume Next
        RunExampleOperation "change-calculation", True, False, Details
        RaisedError = Err.Number
        On Error GoTo CaseFailed
    'Check the change error, the restoration attempt and the restored setting
        TEST_Harness.AssertRaised "change-error", ERR_EXAMPLE_HOST, RaisedError
        TEST_Harness.AssertTrue "calculation-restore-attempted", _
            InStr(Details, "Calculation not restored") > 0, "details='" & Details & "'"
        TEST_Harness.AssertTrue "screen-updating-restored", _
            Application.ScreenUpdating = SavedScreen, "ScreenUpdating differs after the operation"

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

Private Sub CaseCleanupPrimaryOnly()

'
'------------------------------------------------------------------------------
'               CHECK A PRIMARY FAILURE WITH SUCCESSFUL CLEANUP
'------------------------------------------------------------------------------
' PURPOSE
'   When only the operation fails and every restoration succeeds, the
'   caller receives the primary error, no cleanup failure is reported and
'   both settings are restored: a failed operation stays failed even if its
'   cleanup succeeded. 4 assertions.
'
' WHY THIS EXISTS
'   A successful cleanup must not turn a failed operation into a success
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs RunExampleOperation failing in its work only, then checks the
'   raised error, that no cleanup failure was reported, and that both
'   settings are as before
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   RunExampleOperation, TEST_Harness; the Excel object model
'
' NOTES
'   None
'
' UPDATED
'   2026-11-04
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details             As String           'Cleanup failures reported by the operation
    Dim RaisedError         As Long             'Err.Number right after the call
    Dim SavedCalculation    As XlCalculation    'Calculation before the operation
    Dim SavedScreen         As Boolean          'ScreenUpdating before the operation

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "cleanup-primary-only"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Keep both settings to compare after the operation
        SavedCalculation = Application.Calculation
        SavedScreen = Application.ScreenUpdating
    'Run the operation failing in its work only and capture its error
        On Error Resume Next
        RunExampleOperation "work", False, False, Details
        RaisedError = Err.Number
        On Error GoTo CaseFailed
    'Check the primary error, the empty details and both restored settings
        TEST_Harness.AssertRaised "primary-error", ERR_EXAMPLE_OPERATION, RaisedError
        TEST_Harness.AssertEqualText "no-cleanup-failure", vbNullString, Details
        TEST_Harness.AssertTrue "calculation-restored", _
            Application.Calculation = SavedCalculation, "Calculation differs after the operation"
        TEST_Harness.AssertTrue "screen-updating-restored", _
            Application.ScreenUpdating = SavedScreen, "ScreenUpdating differs after the operation"

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

Private Sub CaseCleanupEachRestorationFailure()

'
'------------------------------------------------------------------------------
'                        CHECK EACH RESTORATION FAILURE
'------------------------------------------------------------------------------
' PURPOSE
'   Each restoration failure on its own is reported, names only its
'   setting and does not stop the other restoration. 4 assertions.
'
' WHY THIS EXISTS
'   Every restoration is attempted even when another one fails
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs RunExampleOperation with only the Calculation restoration failing,
'   then with only the ScreenUpdating restoration failing, and checks each
'   time the raised error and that only the failed setting was reported
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   RunExampleOperation, TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details         As String   'Cleanup failures reported by the operation
    Dim RaisedError     As Long     'Err.Number right after the call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "cleanup-each-restoration-failure"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Fail only the Calculation restoration and capture the error
        On Error Resume Next
        RunExampleOperation vbNullString, True, False, Details
        RaisedError = Err.Number
        On Error GoTo CaseFailed
    'Check the cleanup error and that only Calculation was reported
        TEST_Harness.AssertRaised "calculation-cleanup-error", ERR_CLEANUP_FAILED, RaisedError
        TEST_Harness.AssertTrue "only-calculation-reported", _
            InStr(Details, "Calculation not restored") > 0 _
            And InStr(Details, "ScreenUpdating") = 0, "details='" & Details & "'"
    'Fail only the ScreenUpdating restoration and capture the error
        On Error Resume Next
        RunExampleOperation vbNullString, False, True, Details
        RaisedError = Err.Number
        On Error GoTo CaseFailed
    'Check the cleanup error and that only ScreenUpdating was reported
        TEST_Harness.AssertRaised "screen-cleanup-error", ERR_CLEANUP_FAILED, RaisedError
        TEST_Harness.AssertTrue "only-screen-reported", _
            InStr(Details, "ScreenUpdating not restored") > 0 _
            And InStr(Details, "Calculation") = 0, "details='" & Details & "'"

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

Private Sub CaseCleanupOnlyFailure()

'
'------------------------------------------------------------------------------
'                         CHECK A CLEANUP-ONLY FAILURE
'------------------------------------------------------------------------------
' PURPOSE
'   When only the cleanup fails, the operation raises ERR_CLEANUP_FAILED,
'   never success. 2 assertions.
'
' WHY THIS EXISTS
'   An operation whose cleanup failed is not a success
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs RunExampleOperation with both restorations failing and no other
'   failure, then checks the raised error and that details were returned
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   RunExampleOperation, TEST_Harness
'
' NOTES
'   RunSuite skips this case in the SKIPPED_CASE mode
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details         As String   'Cleanup failures reported by the operation
    Dim RaisedError     As Long     'Err.Number right after the call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "cleanup-only-failure"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Run the operation failing only in cleanup and capture its error
        On Error Resume Next
        RunExampleOperation vbNullString, True, True, Details
        RaisedError = Err.Number
        On Error GoTo CaseFailed
    'Check the cleanup error and that details were returned
        TEST_Harness.AssertRaised "cleanup-error", ERR_CLEANUP_FAILED, RaisedError
        TEST_Harness.AssertTrue "failure-reported", Len(Details) > 0, "no cleanup detail returned"

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
'                        EXAMPLE PROCEDURES UNDER TEST
'
'------------------------------------------------------------------------------
'

Private Function ExampleRatio( _
    ByVal Numerator As Double, _
    ByVal Denominator As Double) _
    As Variant
'
'------------------------------------------------------------------------------
'                           DIVIDE OR RETURN #DIV/0!
'------------------------------------------------------------------------------
' PURPOSE
'   Numerator / Denominator, or the worksheet error #DIV/0! for a zero
'   denominator, as a worksheet function returns errors (house style)
'
' WHY THIS EXISTS
'   CaseErrorValue needs a procedure that returns an error value
'
' INPUTS
'   Numerator
'     The dividend
'
'   Denominator
'     The divisor
'
' RETURNS
'   The ratio, or CVErr(xlErrDiv0) when Denominator is zero
'
' BEHAVIOR
'   Returns the error value for a zero denominator, the ratio otherwise
'
' ERROR POLICY
'   None: a zero denominator is returned as an error value, not raised
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return #DIV/0! for a zero denominator, the ratio otherwise
        If Denominator = 0# Then
            ExampleRatio = CVErr(xlErrDiv0)
        Else
            ExampleRatio = Numerator / Denominator
        End If

End Function

Private Function ExampleSquareRoot( _
    ByVal Value As Double) _
    As Double
'
'------------------------------------------------------------------------------
'                          TAKE A CHECKED SQUARE ROOT
'------------------------------------------------------------------------------
' PURPOSE
'   Square root of Value >= 0; a negative Value raises ERR_EXAMPLE_INVALID,
'   as a core procedure raises (house style)
'
' WHY THIS EXISTS
'   CaseRaisedError needs a procedure that raises an error
'
' INPUTS
'   Value
'     The number whose square root is taken; must be >= 0
'
' RETURNS
'   The square root of Value
'
' BEHAVIOR
'   Rejects a negative Value, then returns its square root
'
' ERROR POLICY
'   Raises ERR_EXAMPLE_INVALID when Value is negative
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a negative value
        If Value < 0# Then
            Err.Raise ERR_EXAMPLE_INVALID, "TEST_Examples.ExampleSquareRoot", _
                "Value must be >= 0; got " & CStr(Value) & "."
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the square root
        ExampleSquareRoot = Sqr(Value)

End Function

Private Function FixturePath( _
    ByVal Relative As String) _
    As String
'
'------------------------------------------------------------------------------
'                             BUILD A FIXTURE PATH
'------------------------------------------------------------------------------
' PURPOSE
'   The full path of a file of tests/fixtures/, given relative to it with
'   "/" separators
'
' WHY THIS EXISTS
'   The suite runs in Excel on Windows and in LibreOffice on Linux, whose
'   paths use different separators
'
' INPUTS
'   Relative
'     The path of the file relative to tests/fixtures/, with "/" separators
'
' RETURNS
'   The full path of the file
'
' BEHAVIOR
'   The separator is the one the fixtures folder uses: "/" when it has
'   only forward slashes, as in LibreOffice on Linux, otherwise "\"
'
' ERROR POLICY
'   Raises ERR_EXAMPLE_FIXTURE when no fixtures folder is set
'
' DEPENDENCIES
'   The fixtures folder stored by RunSuite
'
' NOTES
'   None
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Separator   As String   'Path separator of the fixtures folder

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Require a fixtures folder
        If Len(mFixtures) = 0 Then
            Err.Raise ERR_EXAMPLE_FIXTURE, "TEST_Examples.FixturePath", _
                "No fixtures folder: pass the full path of tests/fixtures."
        End If

'------------------------------------------------------------------------------
' COMPOSE
'------------------------------------------------------------------------------
    'Use the separator of the fixtures folder
        Separator = "\"
        If InStr(mFixtures, "/") > 0 And InStr(mFixtures, "\") = 0 Then
            Separator = "/"
        End If
    'Join the folder and the relative path
        FixturePath = mFixtures & Separator & Replace(Relative, "/", Separator)

End Function

Private Function ReadFixture( _
    ByVal Path As String) _
    As String
'
'------------------------------------------------------------------------------
'                            READ AN ASCII FIXTURE
'------------------------------------------------------------------------------
' PURPOSE
'   The bytes of an ASCII test case file, as text with its line endings
'   kept
'
' WHY THIS EXISTS
'   The test case files are ASCII, and the line-ending variants must reach
'   the readers unchanged
'
' INPUTS
'   Path
'     The full path of the file
'
' RETURNS
'   The content of the file
'
' BEHAVIOR
'   Reads the file in binary, then rejects any character outside ASCII
'
' ERROR POLICY
'   A missing or unreadable file raises the file error; a byte above 127
'   raises ERR_EXAMPLE_FIXTURE. The file is closed first
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FileNumber          As Integer  'Handle of the open file
    Dim Text                As String   'Content of the file
    Dim Index               As Long     'Position in Text
    Dim ErrorNumber         As Long     'File error number
    Dim ErrorSource         As String   'File error source
    Dim ErrorDescription    As String   'File error description

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Open the file; Access Read raises on a missing file instead of
    'creating it
        FileNumber = FreeFile
        Open Path For Binary Access Read As #FileNumber
    'Close the file on a read error
        On Error GoTo ReadFailed
    'Read the whole file
        Text = Space$(LOF(FileNumber))
        Get #FileNumber, , Text
        Close #FileNumber
    'Restore normal error handling
        On Error GoTo 0

'------------------------------------------------------------------------------
' REQUIRE ASCII
'------------------------------------------------------------------------------
    'Reject any character outside ASCII
        For Index = 1 To Len(Text)
            If AscW(Mid$(Text, Index, 1)) > 127 Or AscW(Mid$(Text, Index, 1)) < 0 Then
                Err.Raise ERR_EXAMPLE_FIXTURE, "TEST_Examples.ReadFixture", _
                    Path & " is not ASCII at byte " & CStr(Index) & "."
            End If
        Next Index

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the content
        ReadFixture = Text

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ReadFailed:
    'Capture the file error
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
    'Restore normal error handling and close the file
        On Error GoTo 0
        Close #FileNumber
    'Raise the file error again
        Err.Raise ErrorNumber, ErrorSource, ErrorDescription

End Function

Private Function SumJsonNumbers( _
    ByVal Path As String, _
    ByVal Key As String, _
    ByRef Records As Long) _
    As Double
'
'------------------------------------------------------------------------------
'                        SUM THE NUMBERS OF A JSON KEY
'------------------------------------------------------------------------------
' PURPOSE
'   The sum of every number that follows "Key": in a JSON fixture, and how
'   many there are
'
' WHY THIS EXISTS
'   CaseIllustrativeResults totals the amounts of the JSON format example
'
' INPUTS
'   Path
'     The full path of the JSON fixture
'
'   Key
'     The key whose numbers are added up
'
'   Records
'     Receives how many numbers were added up
'
' RETURNS
'   The sum of the numbers
'
' BEHAVIOR
'   Finds each "Key": marker in turn and adds up the number after it
'
' ERROR POLICY
'   As ReadFixture
'
' DEPENDENCIES
'   ReadFixture
'
' NOTES
'   Enough for the illustrative fixture, whose amounts are plain numbers;
'   not a JSON parser
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Text        As String   'Content of the fixture
    Dim Marker      As String   'The key and its colon
    Dim Position    As Long     'Position of the next marker
    Dim Total       As Double   'Sum so far

'------------------------------------------------------------------------------
' ADD UP
'------------------------------------------------------------------------------
    'Read the fixture and build the marker
        Text = ReadFixture(Path)
        Marker = """" & Key & """:"
        Records = 0
    'Add up the number after each marker. Val reads it with a period as the
    'decimal separator, whatever the regional settings, and skips the blanks
        Position = InStr(1, Text, Marker, vbBinaryCompare)
        Do While Position > 0
            Total = Total + Val(Mid$(Text, Position + Len(Marker)))
            Records = Records + 1
            Position = InStr(Position + Len(Marker), Text, Marker, vbBinaryCompare)
        Loop
    'Return the sum
        SumJsonNumbers = Total

End Function

Private Function SumCsvColumn( _
    ByVal Path As String, _
    ByVal Column As String, _
    ByRef Records As Long) _
    As Double
'
'------------------------------------------------------------------------------
'                              SUM ONE CSV COLUMN
'------------------------------------------------------------------------------
' PURPOSE
'   The sum of one column of a comma-separated file with a header row, and
'   how many rows there are
'
' WHY THIS EXISTS
'   CaseIllustrativeResults totals the amounts of the positions files of
'   the file import example
'
' INPUTS
'   Path
'     The full path of the file
'
'   Column
'     The header of the column to add up
'
'   Records
'     Receives how many rows were added up
'
' RETURNS
'   The sum of the column
'
' BEHAVIOR
'   Lines end in CR LF or LF; empty lines are ignored. Finds the column in
'   the header row, then adds up its field in every other line
'
' ERROR POLICY
'   Raises ERR_EXAMPLE_FIXTURE when the header lacks Column or a row lacks
'   the field
'
' DEPENDENCIES
'   ReadFixture
'
' NOTES
'   Enough for the illustrative positions file, whose fields hold no comma
'   or quote; not a general CSV reader
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Lines   As Variant  'The file's lines
    Dim Header  As Variant  'The header's fields
    Dim Fields  As Variant  'One row's fields
    Dim Target  As Long     'Index of Column, -1 until found
    Dim Index   As Long     'Line or field index
    Dim Total   As Double   'Sum so far

'------------------------------------------------------------------------------
' FIND THE COLUMN
'------------------------------------------------------------------------------
    'Split the file into lines and the header into fields
        Lines = Split(Replace(ReadFixture(Path), vbCr, vbNullString), vbLf)
        Header = Split(Lines(0), ",")
    'Find the column in the header
        Target = -1
        For Index = 0 To UBound(Header)
            If Header(Index) = Column Then
                Target = Index
            End If
        Next Index
    'Reject a header without the column
        If Target < 0 Then
            Err.Raise ERR_EXAMPLE_FIXTURE, "TEST_Examples.SumCsvColumn", _
                Path & " has no column " & Column & "."
        End If

'------------------------------------------------------------------------------
' ADD UP
'------------------------------------------------------------------------------
    'Add up the field of every non-empty row, rejecting a row without it
        Records = 0
        For Index = 1 To UBound(Lines)
            If Len(Lines(Index)) > 0 Then
                Fields = Split(Lines(Index), ",")
                If UBound(Fields) < Target Then
                    Err.Raise ERR_EXAMPLE_FIXTURE, "TEST_Examples.SumCsvColumn", _
                        Path & " line " & CStr(Index + 1) & " has no " & Column & "."
                End If
                Total = Total + Val(Fields(Target))
                Records = Records + 1
            End If
        Next Index
    'Return the sum
        SumCsvColumn = Total

End Function

Private Sub RunExampleOperation( _
    ByVal FailAt As String, _
    ByVal FailRestoreCalculation As Boolean, _
    ByVal FailRestoreScreen As Boolean, _
    ByRef CleanupDetails As String)
'
'------------------------------------------------------------------------------
'                          RUN THE EXAMPLE OPERATION
'------------------------------------------------------------------------------
' PURPOSE
'   Reference implementation of the cleanup pattern of
'   docs/VBA_HOUSE_STYLE.md (Operation and cleanup outcomes): capture each
'   setting, register its restoration before changing it, restore only what
'   was registered on both paths, preserve the primary error and report
'   every cleanup failure
'
' WHY THIS EXISTS
'   The cleanup cases test the pattern on an operation that changes real
'   Excel settings
'
' INPUTS
'   FailAt
'     Test seam; empty for none, or where to raise ERR_EXAMPLE_HOST or
'     ERR_EXAMPLE_OPERATION: "capture-screen" (after Calculation was
'     captured), "change-calculation" (after its restoration was
'     registered) or "work"
'
'   FailRestoreCalculation, FailRestoreScreen
'     Test seams; the restore helper restores its setting and then reports
'     a simulated failure, so that the failure paths can be tested without
'     leaving Excel changed
'
'   CleanupDetails
'     Receives the cleanup failures joined by "; ", empty on success; set
'     also when an error is raised
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Changes Application.ScreenUpdating to False and Application.Calculation
'   to manual, and restores each one it registered to the value it captured
'
' ERROR POLICY
'   Re-raises the primary error after cleanup; when only the cleanup failed,
'   raises ERR_CLEANUP_FAILED with the details
'
' DEPENDENCIES
'   RestoreCalculation, RestoreScreenUpdating, AppendDetail; the Excel
'   object model
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SavedCalculation        As XlCalculation    'Calculation mode before the operation
    Dim SavedScreen             As Boolean          'ScreenUpdating before the operation
    Dim RestoreCalculationDue   As Boolean          'Calculation captured and about to change
    Dim RestoreScreenDue        As Boolean          'ScreenUpdating captured and about to change
    Dim ErrNumber               As Long             'Primary error number; 0 for none
    Dim ErrSource               As String           'Primary error source
    Dim ErrDescription          As String           'Primary error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start with no cleanup failure and send errors to the handler
        CleanupDetails = vbNullString
        On Error GoTo Failed

'------------------------------------------------------------------------------
' CAPTURE, REGISTER AND CHANGE
'------------------------------------------------------------------------------
    'A restoration is registered only for a captured setting, just before
    'its change, so a change that fails part-way is still restored and an
    'unchanged setting is never touched
    'Capture Calculation, then ScreenUpdating
        SavedCalculation = Application.Calculation
        If FailAt = "capture-screen" Then
            Err.Raise ERR_EXAMPLE_HOST, "TEST_Examples.RunExampleOperation", _
                "Simulated failure capturing ScreenUpdating."
        End If
        SavedScreen = Application.ScreenUpdating
    'Register the ScreenUpdating restoration, then change it
        RestoreScreenDue = True
        Application.ScreenUpdating = False
    'Register the Calculation restoration, then change it
        RestoreCalculationDue = True
        If FailAt = "change-calculation" Then
            Err.Raise ERR_EXAMPLE_HOST, "TEST_Examples.RunExampleOperation", _
                "Simulated failure changing Calculation."
        End If
        Application.Calculation = xlCalculationManual

'------------------------------------------------------------------------------
' WORK
'------------------------------------------------------------------------------
    'Do the work, here only a simulated failure
        If FailAt = "work" Then
            Err.Raise ERR_EXAMPLE_OPERATION, "TEST_Examples.RunExampleOperation", _
                "Simulated failure of the operation."
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
CleanUp:
    'Restore normal error handling
        On Error GoTo 0
    'Restore each registered setting, collecting every failure
        If RestoreCalculationDue Then
            CleanupDetails = AppendDetail(CleanupDetails, _
                RestoreCalculation(SavedCalculation, FailRestoreCalculation))
        End If
        If RestoreScreenDue Then
            CleanupDetails = AppendDetail(CleanupDetails, _
                RestoreScreenUpdating(SavedScreen, FailRestoreScreen))
        End If
    'Raise the primary error again, or a cleanup failure
        If ErrNumber <> 0 Then
            Err.Raise ErrNumber, ErrSource, ErrDescription
        End If
        If Len(CleanupDetails) > 0 Then
            Err.Raise ERR_CLEANUP_FAILED, "TEST_Examples.RunExampleOperation", CleanupDetails
        End If
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Capture the primary error
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
    'Restore what was registered
        Resume CleanUp

End Sub

Private Function RestoreCalculation( _
    ByVal SavedCalculation As XlCalculation, _
    ByVal SimulateFailure As Boolean) _
    As String
'
'------------------------------------------------------------------------------
'                         RESTORE THE CALCULATION MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Restore Application.Calculation; contain any error
'
' WHY THIS EXISTS
'   Each setting is restored through its own helper and handler, so that
'   one failed restoration does not stop the others
'
' INPUTS
'   SavedCalculation
'     The captured calculation mode
'
'   SimulateFailure
'     Test seam: report a simulated failure after restoring
'
' RETURNS
'   Empty on success, otherwise what failed
'
' BEHAVIOR
'   Restores the calculation mode, then raises the simulated failure when
'   asked
'
' ERROR POLICY
'   Contains every error and returns it as text
'
' DEPENDENCIES
'   The Excel object model
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Contain any error
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Restore the captured mode, then simulate a failure when asked
        Application.Calculation = SavedCalculation
        If SimulateFailure Then
            Err.Raise ERR_EXAMPLE_RESTORE, "TEST_Examples.RestoreCalculation", "simulated"
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Return what failed
        RestoreCalculation = "Calculation not restored: " & Err.Description

End Function

Private Function RestoreScreenUpdating( _
    ByVal SavedScreen As Boolean, _
    ByVal SimulateFailure As Boolean) _
    As String
'
'------------------------------------------------------------------------------
'                           RESTORE SCREEN UPDATING
'------------------------------------------------------------------------------
' PURPOSE
'   Restore Application.ScreenUpdating; contain any error
'
' WHY THIS EXISTS
'   Each setting is restored through its own helper and handler, so that
'   one failed restoration does not stop the others
'
' INPUTS
'   SavedScreen
'     The captured ScreenUpdating value
'
'   SimulateFailure
'     Test seam: report a simulated failure after restoring
'
' RETURNS
'   Empty on success, otherwise what failed
'
' BEHAVIOR
'   Restores ScreenUpdating, then raises the simulated failure when asked
'
' ERROR POLICY
'   Contains every error and returns it as text
'
' DEPENDENCIES
'   The Excel object model
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Contain any error
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Restore the captured value, then simulate a failure when asked
        Application.ScreenUpdating = SavedScreen
        If SimulateFailure Then
            Err.Raise ERR_EXAMPLE_RESTORE, "TEST_Examples.RestoreScreenUpdating", "simulated"
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Return what failed
        RestoreScreenUpdating = "ScreenUpdating not restored: " & Err.Description

End Function

Private Function AppendDetail( _
    ByVal First As String, _
    ByVal Second As String) _
    As String
'
'------------------------------------------------------------------------------
'                           APPEND A FAILURE DETAIL
'------------------------------------------------------------------------------
' PURPOSE
'   Join two failure texts with "; ", skipping empty ones
'
' WHY THIS EXISTS
'   RunExampleOperation collects every cleanup failure in one text
'
' INPUTS
'   First
'     The failures collected so far
'
'   Second
'     The failure to add, empty for none
'
' RETURNS
'   The joined text
'
' BEHAVIOR
'   Returns the non-empty text alone, or both joined with "; "
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Join the texts, skipping an empty one
        If Len(First) = 0 Then
            AppendDetail = Second
        ElseIf Len(Second) = 0 Then
            AppendDetail = First
        Else
            AppendDetail = First & "; " & Second
        End If

End Function
