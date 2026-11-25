Attribute VB_Name = "TEST_Harness"

Option Explicit
Option Private Module

'
'==============================================================================
' MODULE: TEST_Harness
'==============================================================================
' PURPOSE
'   Deterministic regression harness for the workbook's VBA test suites:
'   one run per suite, stable case identifiers, exact and tolerance
'   assertions, expected error values and raised errors, expected case and
'   assertion counts, a complete key=value log, and the certification of a
'   suite's runs as the evidence of an Excel run. It holds no calculation
'   of the workbook
'
' WHY THIS EXISTS
'   Every suite needs the same run lifecycle, assertions, log and
'   certification, so that an Excel run leaves evidence that
'   tools/check_excel_run.py can read; the harness provides them without a
'   reference, a workbook fixture or an external framework
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Public surface:
'     - suite lifecycle: BeginRun, BeginCase, RecordCaseError, EndRun
'     - assertions: AssertEqualDouble, AssertNear, AssertEqualText,
'       AssertTrue, AssertErrorValue, AssertRaised; WithinTolerance is the
'       comparison rule
'     - log: Say, Echo, SaveLog, LastResult
'     - results: RecordResult, SaveResults
'     - recovery: ResetTests
'     - certification: BeginCertification, CertifyRun, EndCertification
'   Option Private Module keeps all of them out of the workbook's external
'   surface.
'   State: owns the run counters, the case identifiers, the failure
'   records, the retained log, the recorded results, the certification
'   records and the active-run, recovery-required and active-certification
'   flags. Reads four Application properties, the number of open workbooks
'   and the caption of the active window to prove a run left them
'   unchanged; never changes them. Reads the build stamp's commit from
'   ThisWorkbook for SaveResults; never writes a workbook or worksheet
'
' ERROR POLICY
'   Assertion failures are recorded and the run continues. EndRun re-raises
'   an unexpected runner error unchanged after cleanup and reporting, or
'   raises ERR_TEST_FAILED for any other failed outcome. A run that starts
'   while another is active, or while recovery is required, is refused with
'   ERR_TEST_DIRTY_START. Misuse of the lifecycle raises ERR_TEST_USAGE.
'   EndCertification writes the certification log first, then raises
'   ERR_TEST_FAILED when the certification did not pass
'
' DEPENDENCIES
'   The VBA and Excel object models, and the build_commit cell of
'   ThisWorkbook, read by SaveResults; no reference, workbook fixture or
'   external framework
'
' NOTES
'   Log format: one record per line, KEY=value; fields separated by "; ".
'   The keys and their order are documented in
'   docs/methodology/TEST_HARNESS.md (Log format, Certification log).
'   SaveResults writes the JSON results file of
'   docs/methodology/TEST_CASES.md (Accuracy contracts).
'
'   Compatibility: Excel VBA on Windows. The log records the host, its
'   version, the Office bitness and the VBA generation.
'
'   Usage: called by the suites in tests/modules/, for example
'   TEST_Examples.RunExampleTests and TEST_Examples.CertifyExamples; SaveLog
'   and ResetTests are also typed in the Immediate window. It has no suite
'   of its own
'
' UPDATED
'   2026-11-04
'==============================================================================

'------------------------------------------------------------------------------
' PUBLIC CONSTANTS
'------------------------------------------------------------------------------
    'Test-only error numbers, kept apart from the production range in use.
    'Written as literals equal to vbObjectError + n: LibreOffice, the optional
    'smoke host, cannot evaluate vbObjectError in a constant expression
    Public Const ERR_TEST_DIRTY_START   As Long = &H80040FE0    'vbObjectError + 4064: start refused
    Public Const ERR_TEST_FAILED        As Long = &H80040FE1    'vbObjectError + 4065: run finished without PASS
    Public Const ERR_TEST_USAGE         As Long = &H80040FE2    'vbObjectError + 4066: lifecycle used out of order

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'The defined name that carries the build stamp's commit
    '(tools/devbuild/workbook.json, stamp.commit)
    Private Const STAMP_COMMIT_NAME     As String = "build_commit"  'Defined name of the stamp's commit

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    'One run at a time. Counters and the log stay readable after EndRun,
    'until the next BeginRun or ResetTests
    Private mRunActive              As Boolean          'True from BeginRun to EndRun
    Private mRecoveryRequired       As Boolean          'True until ResetTests after an unsafe end
    Private mRecoveryReason         As String           'Why recovery is required
    Private mSuite                  As String           'Name of the running or last suite
    Private mExpectedCases          As Long             'Cases a complete run starts
    Private mExpectedAssertions     As Long             'Assertions a complete run evaluates
    Private mCaseCount              As Long             'Cases started in this run
    Private mAssertionCount         As Long             'Assertions evaluated in this run
    Private mFailureCount           As Long             'Failures of any kind in this run
    Private mErrorCount             As Long             'Unexpected errors in this run
    Private mCurrentCase            As String           'Identifier of the running case
    Private mCaseIds                As Collection       'Case identifiers seen in this run
    Private mPrimaryFailure         As String           'First failure recorded in this run
    Private mLastResult             As String           'RESULT record of the last finished run
    Private mLog                    As Collection       'Log since the last BeginRun or ResetTests
    Private mResults                As Collection       'Results recorded since the last BeginRun or ResetTests
    Private mSnapshotValid          As Boolean          'True once all four properties were read
    Private mInitialCalculation     As XlCalculation    'Calculation at BeginRun
    Private mInitialAlerts          As Boolean          'DisplayAlerts at BeginRun
    Private mInitialEvents          As Boolean          'EnableEvents at BeginRun
    Private mInitialScreen          As Boolean          'ScreenUpdating at BeginRun
    Private mInitialWorkbooks       As Long             'Open workbooks at BeginRun
    Private mInitialWindow          As String           'Active window caption at BeginRun
    Private mCertActive             As Boolean          'True from BeginCertification to EndCertification
    Private mCertSuite              As String           'Suite being certified
    Private mCertLog                As Collection       'Certification records
    Private mCertModes              As Long             'Runs certified
    Private mCertFailures           As Long             'Runs that did not end as required
    Private mCertMandatory          As Long             'Runs required to pass
    Private mCertComplete           As Long             'Required runs that were complete


'
'------------------------------------------------------------------------------
'
'                               SUITE LIFECYCLE
'
'------------------------------------------------------------------------------
'

Public Sub BeginRun( _
    ByVal SuiteName As String, _
    ByVal Mode As String, _
    ByVal ExpectedCases As Long, _
    ByVal ExpectedAssertions As Long)
'
'------------------------------------------------------------------------------
'                              START A SUITE RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Start one suite run: refuse an unsafe start, reset the run state, record
'   the host state to compare at the end and write the log header
'
' WHY THIS EXISTS
'   A run must not overwrite the state of an interrupted run, and EndRun
'   needs the host state as it was at the start to prove cleanup
'
' INPUTS
'   SuiteName
'     Stable suite identifier, written in the RUN and RESULT records
'
'   Mode
'     NORMAL, or the name of a deliberate failure mode of the suite
'
'   ExpectedCases
'     Cases of a complete run, >= 1
'
'   ExpectedAssertions
'     Assertions of a complete run, >= 1
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Refuses an unsafe start, resets the run state, claims the run, reads the
'   host state, then writes the RUN and ENVIRONMENT records
'
' ERROR POLICY
'   Raises ERR_TEST_DIRTY_START, after logging why, when a run is active or
'   recovery is required; raises ERR_TEST_USAGE for counts below 1. A host
'   property that cannot be read leaves the snapshot invalid, which fails
'   cleanup at EndRun; it does not stop the run
'
' DEPENDENCIES
'   ResetRunState, ActiveWindowCaption, EnvironmentSummary, the Excel
'   Application object
'
' NOTES
'   A project reset (End or Reset in the VBE) clears the run state, so a run
'   stopped that way is not seen as interrupted
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' REFUSE AN UNSAFE START
'------------------------------------------------------------------------------
    'An interrupted run or an unverified cleanup must be acknowledged with
    'ResetTests before a new run may overwrite its state
        If mRunActive Then
            mRecoveryRequired = True
            mRecoveryReason = "run of " & mSuite & " still active (interrupted?)"
        End If
    'Refuse the start, logging the recovery reason and a failed result
        If mRecoveryRequired Then
            mLastResult = "RESULT=FAIL; suite=" & SuiteName & "; operation=NOT_RUN" & _
                "; completeness=INCOMPLETE; cases=0; assertions=0; failures=0" & _
                "; cleanup=NOT_RUN; recovery=REQUIRED"
            'RESULT is always the last record of a run
            Say "RECOVERY=REQUIRED; reason=" & mRecoveryReason & _
                "; action=check Excel state, then run TEST_Harness.ResetTests"
            Say mLastResult
            Err.Raise ERR_TEST_DIRTY_START, "TEST_Harness.BeginRun", _
                "Recovery required: " & mRecoveryReason & ". Run TEST_Harness.ResetTests."
        End If
    'Reject expected counts below 1
        If ExpectedCases < 1 Or ExpectedAssertions < 1 Then
            Err.Raise ERR_TEST_USAGE, "TEST_Harness.BeginRun", _
                "Expected counts must be at least 1; got cases=" & CStr(ExpectedCases) & _
                ", assertions=" & CStr(ExpectedAssertions) & "."
        End If

'------------------------------------------------------------------------------
' RESET AND CLAIM THE RUN
'------------------------------------------------------------------------------
    'Clear the previous run, its log and its results
        ResetRunState
        Set mLog = New Collection
        Set mResults = New Collection
    'Store the suite and its expected counts, then mark the run active
        mSuite = SuiteName
        mExpectedCases = ExpectedCases
        mExpectedAssertions = ExpectedAssertions
        mRunActive = True

'------------------------------------------------------------------------------
' SNAPSHOT HOST STATE
'------------------------------------------------------------------------------
    'Mark the snapshot valid only after every read, so that cleanup never
    'compares a value that was not read
        On Error GoTo SnapshotFailed
        mInitialCalculation = Application.Calculation
        mInitialAlerts = Application.DisplayAlerts
        mInitialEvents = Application.EnableEvents
        mInitialScreen = Application.ScreenUpdating
        mInitialWorkbooks = Application.Workbooks.Count
        mInitialWindow = ActiveWindowCaption()
        mSnapshotValid = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
WriteHeader:
    'Restore normal error handling
        On Error GoTo 0
    'Write the run header and the host description
        Say "RUN=" & SuiteName & "; mode=" & Mode & _
            "; expected_cases=" & CStr(ExpectedCases) & _
            "; expected_assertions=" & CStr(ExpectedAssertions) & _
            "; started=" & Format$(Now, "yyyy-mm-dd hh:mm:ss")
        Say "ENVIRONMENT=" & EnvironmentSummary()
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
SnapshotFailed:
    'Leave the snapshot invalid, so that EndRun fails cleanup, and go on
        mSnapshotValid = False
        Resume WriteHeader

End Sub

Public Sub BeginCase( _
    ByVal CaseId As String)
'
'------------------------------------------------------------------------------
'                                 START A CASE
'------------------------------------------------------------------------------
' PURPOSE
'   Count and log the start of one case
'
' WHY THIS EXISTS
'   Stable case identifiers let logs be compared case by case, and the case
'   count proves a run complete
'
' INPUTS
'   CaseId
'     Lower-case kebab-case identifier, unique within the run and stable
'     across runs
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Counts the case, makes it the running case, writes its CASE record and
'   registers its identifier
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a run. A malformed or repeated identifier
'   is recorded as a failure; the case still counts and runs
'
' DEPENDENCIES
'   RequireActiveRun, Say, IsCaseId, IsKnownCase, RecordFailure
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' REGISTER THE CASE
'------------------------------------------------------------------------------
    'Refuse a case outside a run
        RequireActiveRun "BeginCase"
    'Count the case and log its start
        mCaseCount = mCaseCount + 1
        mCurrentCase = CaseId
        Say "CASE=" & CaseId
    'Record a malformed or repeated identifier, otherwise register it
        If Not IsCaseId(CaseId) Then
            RecordFailure "harness.case-id", "malformed case identifier '" & CaseId & "'"
        ElseIf IsKnownCase(CaseId) Then
            RecordFailure "harness.case-id", "repeated case identifier '" & CaseId & "'"
        Else
            mCaseIds.Add CaseId, CaseId
        End If

End Sub

Public Sub RecordCaseError()

'
'------------------------------------------------------------------------------
'                             RECORD A CASE ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   Record the error that ended the running case, so that the remaining
'   cases still run
'
' WHY THIS EXISTS
'   One unexpected error must fail its case without stopping the suite
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Reads the Err object before anything else can reset it, then counts an
'   unexpected error and a failure, and logs the ERROR record. It
'   evaluates no assertion
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a run
'
' DEPENDENCIES
'   RequireActiveRun, Say, RecordFailure
'
' NOTES
'   Call it first thing in the case's error handler
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrorNumber         As Long     'Number of the case error
    Dim ErrorSource         As String   'Source of the case error
    Dim ErrorDescription    As String   'Description of the case error

'------------------------------------------------------------------------------
' RECORD
'------------------------------------------------------------------------------
    'Capture the case error before any other call
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
    'Refuse a call outside a run
        RequireActiveRun "RecordCaseError"
    'Count the error, log it and record it as a failure of the case
        mErrorCount = mErrorCount + 1
        Say "ERROR=" & mCurrentCase & "; number=" & CStr(ErrorNumber) & _
            "; source=" & ErrorSource & "; description=" & ErrorDescription
        RecordFailure mCurrentCase & ".unexpected-error", _
            "number=" & CStr(ErrorNumber) & "; source=" & ErrorSource

End Sub

Public Sub EndRun( _
    ByVal RunErrorNumber As Long, _
    ByVal RunErrorSource As String, _
    ByVal RunErrorDescription As String)
'
'------------------------------------------------------------------------------
'                              FINISH A SUITE RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Finish the run: check completeness, release the run, verify the host
'   state, write the summary and raise the outcome
'
' WHY THIS EXISTS
'   A run passes only when what ran passed, the suite ran as written and the
'   host state is unchanged; each outcome is reported on its own
'
' INPUTS
'   RunErrorNumber
'     Number of the error that escaped the suite's runner, or 0 when none
'     did
'
'   RunErrorSource
'     Its source, or an empty string
'
'   RunErrorDescription
'     Its description, or an empty string
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Outcomes written to the log:
'     - operation: what ran: PASS, FAIL (an assertion or a case identifier
'       failed) or ERROR (an unexpected error occurred)
'     - completeness: COMPLETE only when both expected counts were met
'       exactly
'     - cleanup: PASS when the host state is unchanged
'     - recovery: REQUIRED when the host state changed or could not be
'       checked, or when a start was refused while this run was active
'   RESULT=PASS needs operation PASS, COMPLETE, cleanup PASS and no
'   recovery
'
' ERROR POLICY
'   Re-raises the runner error unchanged, after reporting, so that cleanup
'   never replaces the primary failure; otherwise raises ERR_TEST_FAILED
'   unless RESULT=PASS. Raises ERR_TEST_USAGE outside a run
'
' DEPENDENCIES
'   RequireActiveRun, Say, RecordFailure, VerifyHostState
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
    Dim CleanupDetail   As String   'Reason for the cleanup outcome
    Dim CleanupPassed   As Boolean  'True when the host state is unchanged
    Dim Completed       As Boolean  'True when both expected counts were met
    Dim Operation       As String   'PASS, FAIL or ERROR
    Dim Verdict         As String   'PASS or FAIL

'------------------------------------------------------------------------------
' COLLECT THE OUTCOMES
'------------------------------------------------------------------------------
    'Refuse a call outside a run
        RequireActiveRun "EndRun"
        mCurrentCase = "run"
    'Count, log and record an error that escaped the runner
        If RunErrorNumber <> 0 Then
            mErrorCount = mErrorCount + 1
            Say "ERROR=run; number=" & CStr(RunErrorNumber) & "; source=" & RunErrorSource & _
                "; description=" & RunErrorDescription
            RecordFailure "runner.unexpected-error", _
                "number=" & CStr(RunErrorNumber) & "; source=" & RunErrorSource
        End If
    'The operation outcome covers what ran; completeness and cleanup are
    'judged separately below, so that each is reported on its own
        If mErrorCount > 0 Then
            Operation = "ERROR"
        ElseIf mFailureCount > 0 Then
            Operation = "FAIL"
        Else
            Operation = "PASS"
        End If
    'Exact counts: a skipped case or assertion, or an extra one, means the
    'run did not execute the suite as written
        Completed = (mCaseCount = mExpectedCases) And _
                    (mAssertionCount = mExpectedAssertions)
        If Not Completed Then
            RecordFailure "harness.completeness", _
                "expected cases=" & CStr(mExpectedCases) & _
                ", assertions=" & CStr(mExpectedAssertions) & _
                "; observed cases=" & CStr(mCaseCount) & _
                ", assertions=" & CStr(mAssertionCount)
        End If
    'Verify the host state, release the run and require recovery on change
        CleanupPassed = VerifyHostState(CleanupDetail)
        mRunActive = False
        If Not CleanupPassed Then
            mRecoveryRequired = True
            mRecoveryReason = CleanupDetail
        End If
    'Pass only when every outcome passed
        If Operation = "PASS" And Completed And CleanupPassed And Not mRecoveryRequired Then
            Verdict = "PASS"
        Else
            Verdict = "FAIL"
        End If

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Write the counts and the primary failure
        Say "CASES=" & CStr(mCaseCount) & "; expected=" & CStr(mExpectedCases)
        Say "ASSERTIONS=" & CStr(mAssertionCount) & "; expected=" & CStr(mExpectedAssertions)
        Say "FAILURES=" & CStr(mFailureCount) & "; errors=" & CStr(mErrorCount)
        If Len(mPrimaryFailure) > 0 Then
            Say "PRIMARY_FAILURE=" & mPrimaryFailure
        End If
    'Write the cleanup and recovery outcomes
        Say "CLEANUP=" & IIf(CleanupPassed, "PASS", "FAIL") & "; detail=" & CleanupDetail
        If mRecoveryRequired Then
            Say "RECOVERY=REQUIRED; reason=" & mRecoveryReason & _
                "; action=check Excel state, then run TEST_Harness.ResetTests"
        Else
            Say "RECOVERY=NONE"
        End If
    'Keep and write the RESULT record, always the last of a run
        mLastResult = "RESULT=" & Verdict & "; suite=" & mSuite & _
            "; operation=" & Operation & _
            "; completeness=" & IIf(Completed, "COMPLETE", "INCOMPLETE") & _
            "; cases=" & CStr(mCaseCount) & _
            "; assertions=" & CStr(mAssertionCount) & _
            "; failures=" & CStr(mFailureCount) & _
            "; cleanup=" & IIf(CleanupPassed, "PASS", "FAIL") & _
            "; recovery=" & IIf(mRecoveryRequired, "REQUIRED", "NONE")
        Say mLastResult

'------------------------------------------------------------------------------
' RAISE THE OUTCOME
'------------------------------------------------------------------------------
    'Re-raise the runner error unchanged
        If RunErrorNumber <> 0 Then
            Err.Raise RunErrorNumber, RunErrorSource, RunErrorDescription
        End If
    'Raise a failed run
        If Verdict <> "PASS" Then
            Err.Raise ERR_TEST_FAILED, "TEST_Harness.EndRun", _
                "Suite " & mSuite & " did not pass; see the log."
        End If

End Sub

Public Sub ResetTests()

'
'------------------------------------------------------------------------------
'                          ACKNOWLEDGE AN UNSAFE END
'------------------------------------------------------------------------------
' PURPOSE
'   Acknowledge an interrupted run or an unverified cleanup, so that the
'   next run may start
'
' WHY THIS EXISTS
'   BeginRun refuses to start while recovery is required; only a check by
'   hand can clear that state
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears the run state, both flags and an unfinished certification, starts
'   a new log and new results, and logs the reset
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   ResetRunState, Say
'
' NOTES
'   Run only after the interrupted run has stopped and the Excel state has
'   been checked by hand; it restores nothing itself
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESET
'------------------------------------------------------------------------------
    'Clear the run state and the recovery flag
        ResetRunState
        mRecoveryRequired = False
        mRecoveryReason = vbNullString
    'Drop an unfinished certification
        mCertActive = False
        Set mCertLog = Nothing
    'Start a new log and new results, and log the reset
        Set mLog = New Collection
        Set mResults = New Collection
        Say "RESET=DONE; run_active=no; recovery=NONE"

End Sub

'
'------------------------------------------------------------------------------
'
'                                  ASSERTIONS
'
'------------------------------------------------------------------------------
'

Public Sub AssertEqualDouble( _
    ByVal AssertionName As String, _
    ByVal Expected As Double, _
    ByVal Actual As Double)
'
'------------------------------------------------------------------------------
'                         ASSERT EXACT DOUBLE EQUALITY
'------------------------------------------------------------------------------
' PURPOSE
'   Require exact Double equality, for results that are exact in binary
'
' WHY THIS EXISTS
'   A result that is exact in binary needs no tolerance
'
' INPUTS
'   AssertionName
'     Stable name within the case
'
'   Expected
'     Expected value
'
'   Actual
'     Value computed
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Counts and logs one assertion that passes when Actual = Expected
'
' ERROR POLICY
'   As LogAssertion
'
' DEPENDENCIES
'   LogAssertion, NumberText
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Compare exactly and log both values
        LogAssertion AssertionName, (Actual = Expected), _
            "expected=" & NumberText(Expected) & "; actual=" & NumberText(Actual)

End Sub

Public Sub AssertNear( _
    ByVal AssertionName As String, _
    ByVal Expected As Double, _
    ByVal Actual As Double, _
    ByVal AbsoluteTolerance As Double, _
    ByVal RelativeTolerance As Double)
'
'------------------------------------------------------------------------------
'                       ASSERT A VALUE WITHIN TOLERANCE
'------------------------------------------------------------------------------
' PURPOSE
'   Require |Actual - Expected| <= max(AbsoluteTolerance, RelativeTolerance
'   * |Expected|), the rule of docs/methodology/TEST_CASES.md (Tolerance)
'
' WHY THIS EXISTS
'   Most computed amounts are not exact in binary and need a stated
'   tolerance
'
' INPUTS
'   AssertionName
'     Stable name within the case
'
'   Expected
'     Expected value
'
'   Actual
'     Value computed
'
'   AbsoluteTolerance
'     Absolute bound, >= 0, always stated
'
'   RelativeTolerance
'     Relative bound, >= 0, always stated
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Counts and logs one assertion. Near zero the absolute bound governs;
'   both 0 demands exact equality
'
' ERROR POLICY
'   A negative tolerance fails the assertion; nothing is raised
'
' DEPENDENCIES
'   LogAssertion, WithinTolerance, NumberText
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Fail a negative tolerance, otherwise compare within the tolerance
        If AbsoluteTolerance < 0# Or RelativeTolerance < 0# Then
            LogAssertion AssertionName, False, "tolerances must not be negative"
        Else
            LogAssertion AssertionName, _
                WithinTolerance(Expected, Actual, AbsoluteTolerance, RelativeTolerance), _
                "expected=" & NumberText(Expected) & "; actual=" & NumberText(Actual) & _
                "; absolute=" & NumberText(AbsoluteTolerance) & _
                "; relative=" & NumberText(RelativeTolerance)
        End If

End Sub

Public Sub AssertEqualText( _
    ByVal AssertionName As String, _
    ByVal Expected As String, _
    ByVal Actual As String)
'
'------------------------------------------------------------------------------
'                              ASSERT EQUAL TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Require the same text, compared binary: case and spaces count
'
' WHY THIS EXISTS
'   Identifiers and messages must match exactly
'
' INPUTS
'   AssertionName
'     Stable name within the case
'
'   Expected
'     Expected text
'
'   Actual
'     Text produced
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Counts and logs one assertion that passes when StrComp in binary mode
'   returns 0
'
' ERROR POLICY
'   As LogAssertion
'
' DEPENDENCIES
'   LogAssertion
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Compare in binary mode and log both texts
        LogAssertion AssertionName, (StrComp(Actual, Expected, vbBinaryCompare) = 0), _
            "expected='" & Expected & "'; actual='" & Actual & "'"

End Sub

Public Sub AssertTrue( _
    ByVal AssertionName As String, _
    ByVal Condition As Boolean, _
    ByVal Detail As String)
'
'------------------------------------------------------------------------------
'                              ASSERT A CONDITION
'------------------------------------------------------------------------------
' PURPOSE
'   Require a condition the suite evaluated; Detail is logged on failure
'
' WHY THIS EXISTS
'   Some checks are a condition the suite computes itself
'
' INPUTS
'   AssertionName
'     Stable name within the case
'
'   Condition
'     The condition that must hold
'
'   Detail
'     Text logged when it does not
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Counts and logs one assertion that passes when Condition is True
'
' ERROR POLICY
'   As LogAssertion
'
' DEPENDENCIES
'   LogAssertion
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Log the condition
        LogAssertion AssertionName, Condition, Detail

End Sub

Public Sub AssertErrorValue( _
    ByVal AssertionName As String, _
    ByVal ExpectedCode As Long, _
    ByVal Actual As Variant)
'
'------------------------------------------------------------------------------
'                        ASSERT A WORKSHEET ERROR VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Require a specific worksheet error value, such as CVErr(xlErrValue)
'
' WHY THIS EXISTS
'   A worksheet function may return a documented error value, which cannot
'   be compared with the other assertions
'
' INPUTS
'   AssertionName
'     Stable name within the case
'
'   ExpectedCode
'     An xlErr constant
'
'   Actual
'     The value returned
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Fails a value that is not an error; otherwise compares the error values
'   as text
'
' ERROR POLICY
'   As LogAssertion
'
' DEPENDENCIES
'   LogAssertion
'
' NOTES
'   Comparing a value with an error value by "=" raises error 13, Type
'   mismatch
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Check the type first, then compare the error values as text
    '("Error 2007"): comparing a value with an error value by "=" raises
    'error 13, Type mismatch
        If Not IsError(Actual) Then
            LogAssertion AssertionName, False, _
                "expected error " & CStr(ExpectedCode) & "; actual type=" & TypeName(Actual)
        Else
            LogAssertion AssertionName, (CStr(Actual) = CStr(CVErr(ExpectedCode))), _
                "expected error " & CStr(ExpectedCode) & "; actual=" & CStr(Actual)
        End If

End Sub

Public Sub AssertRaised( _
    ByVal AssertionName As String, _
    ByVal ExpectedNumber As Long, _
    ByVal ActualNumber As Long)
'
'------------------------------------------------------------------------------
'                            ASSERT A RAISED ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   Require that a call raised a specific error number
'
' WHY THIS EXISTS
'   Error paths are part of a procedure's contract and need their own
'   assertion
'
' INPUTS
'   AssertionName
'     Stable name within the case
'
'   ExpectedNumber
'     The error number the call must raise
'
'   ActualNumber
'     Err.Number captured on the line after the call; 0 means nothing was
'     raised
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Counts and logs one assertion that passes when the numbers are equal
'
' ERROR POLICY
'   As LogAssertion
'
' DEPENDENCIES
'   LogAssertion
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Compare the error numbers and log both
        LogAssertion AssertionName, (ActualNumber = ExpectedNumber), _
            "expected error=" & CStr(ExpectedNumber) & "; actual error=" & CStr(ActualNumber)

End Sub

Public Function WithinTolerance( _
    ByVal Expected As Double, _
    ByVal Actual As Double, _
    ByVal AbsoluteTolerance As Double, _
    ByVal RelativeTolerance As Double) _
    As Boolean
'
'------------------------------------------------------------------------------
'                           COMPARE WITHIN TOLERANCE
'------------------------------------------------------------------------------
' PURPOSE
'   The comparison rule of AssertNear, exposed so that suites can test it
'
' WHY THIS EXISTS
'   A relation test checks that the rule rejects a corrupted output, which
'   needs the rule without an assertion
'
' INPUTS
'   Expected
'     Expected value
'
'   Actual
'     Value computed
'
'   AbsoluteTolerance
'     Absolute bound
'
'   RelativeTolerance
'     Relative bound
'
' RETURNS
'   True when |Actual - Expected| <= max(AbsoluteTolerance,
'   RelativeTolerance * |Expected|); the bound is inclusive
'
' BEHAVIOR
'   Takes the larger of the two bounds and compares the difference with it
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
' DECLARE
'------------------------------------------------------------------------------
    Dim Bound   As Double   'Larger of the two bounds

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Take the larger of the two bounds
        Bound = RelativeTolerance * Abs(Expected)
        If AbsoluteTolerance > Bound Then
            Bound = AbsoluteTolerance
        End If
    'Compare the difference with it, inclusively
        WithinTolerance = (Abs(Actual - Expected) <= Bound)

End Function

'
'------------------------------------------------------------------------------
'
'                                CERTIFICATION
'
'------------------------------------------------------------------------------
'

Public Sub BeginCertification( _
    ByVal SuiteName As String)
'
'------------------------------------------------------------------------------
'                            START A CERTIFICATION
'------------------------------------------------------------------------------
' PURPOSE
'   Start the certification of one suite: refuse state left by an earlier
'   run, then keep a certification log apart from the run logs
'
' WHY THIS EXISTS
'   The certification is the evidence of an Excel run and must not mix with
'   state left by an earlier run
'
' INPUTS
'   SuiteName
'     The suite identifier, as in its RUN records
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Refuses leftover state, resets the certification counters and the last
'   result, starts the certification log and writes its CERTIFY record
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE, changing nothing, while another certification
'   is active. A run still active is taken as interrupted: it marks
'   recovery required, with the reason, so that ResetTests is needed before
'   any run, then raises ERR_TEST_DIRTY_START, as it does when recovery is
'   already required. It changes no other state before raising
'
' DEPENDENCIES
'   CertSay
'
' NOTES
'   None
'
' UPDATED
'   2026-11-04
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' REFUSE LEFTOVER STATE
'------------------------------------------------------------------------------
    'Refuse a second active certification
        If mCertActive Then
            Err.Raise ERR_TEST_USAGE, "TEST_Harness.BeginCertification", _
                "The certification of " & mCertSuite & " is still active. " & _
                "Run TEST_Harness.ResetTests."
        End If
    'Take a run still active as interrupted
        If mRunActive Then
            mRecoveryRequired = True
            mRecoveryReason = "run of " & mSuite & " still active (interrupted?)"
        End If
    'Refuse to start while recovery is required
        If mRecoveryRequired Then
            Err.Raise ERR_TEST_DIRTY_START, "TEST_Harness.BeginCertification", _
                "Recovery required: " & mRecoveryReason & ". Run TEST_Harness.ResetTests."
        End If

'------------------------------------------------------------------------------
' CLAIM THE CERTIFICATION
'------------------------------------------------------------------------------
    'Mark the certification active and reset its counters
        mCertActive = True
        mCertSuite = SuiteName
        mCertModes = 0
        mCertFailures = 0
        mCertMandatory = 0
        mCertComplete = 0
        mLastResult = vbNullString
    'Start the certification log
        Set mCertLog = New Collection
        CertSay "CERTIFY=" & SuiteName & "; started=" & Format$(Now, "yyyy-mm-dd hh:mm:ss")

End Sub

Public Sub CertifyRun( _
    ByVal Mode As String, _
    ByVal MustPass As Boolean, _
    ByVal RaisedNumber As Long, _
    ByVal RaisedDescription As String)
'
'------------------------------------------------------------------------------
'                               CERTIFY ONE RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Record how one run of the suite reached its caller, and whether that is
'   how the mode must end
'
' WHY THIS EXISTS
'   A deliberate failure mode proves that a failure is never silent; it
'   must end as a failure, as a normal run must end as a pass
'
' INPUTS
'   Mode
'     The run's mode
'
'   MustPass
'     True for a run that must pass, False for a deliberate failure mode
'
'   RaisedNumber
'     Number of the error that reached the caller of the suite's entry
'     point, or 0 when none did
'
'   RaisedDescription
'     Its description, or an empty string
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   A run that must pass ends with no error and a RESULT=PASS that is
'   COMPLETE. A deliberate failure ends with RESULT=FAIL and reaches its
'   caller as ERR_TEST_FAILED with a description: the self-test that a
'   failure is never silent. A run with no RESULT of its own fails. The
'   outcome is written as a MODE record
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a certification. A run that is still
'   active did not end: it requires recovery and is recorded as failed, so
'   that the certification still writes its log
'
' DEPENDENCIES
'   RequireCertification, CertSay
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
    Dim Completeness    As String   'COMPLETE, INCOMPLETE or NONE
    Dim Passed          As Boolean  'True when the run ended as required
    Dim Verdict         As String   'PASS, FAIL or NONE

'------------------------------------------------------------------------------
' READ THE RUN'S RESULT
'------------------------------------------------------------------------------
    'Refuse a call outside a certification
        RequireCertification "CertifyRun"
    'Treat a run still active as one that did not end
        If mRunActive Then
            mRunActive = False
            mRecoveryRequired = True
            mRecoveryReason = "run of " & mSuite & " did not end"
            mLastResult = vbNullString
        End If
    'Read the verdict and completeness of the run's RESULT record, if any
        If Len(mLastResult) = 0 Then
            Verdict = "NONE"
            Completeness = "NONE"
        Else
            Verdict = IIf(Left$(mLastResult, 12) = "RESULT=PASS;", "PASS", "FAIL")
            Completeness = IIf(InStr(mLastResult, "; completeness=COMPLETE;") > 0, _
                               "COMPLETE", "INCOMPLETE")
        End If
    'Consume the RESULT record
        mLastResult = vbNullString

'------------------------------------------------------------------------------
' JUDGE AND RECORD
'------------------------------------------------------------------------------
    'Judge the run against how its mode must end
        If MustPass Then
            Passed = (RaisedNumber = 0) And (Verdict = "PASS") And (Completeness = "COMPLETE")
            mCertMandatory = mCertMandatory + 1
            If Completeness = "COMPLETE" Then
                mCertComplete = mCertComplete + 1
            End If
        Else
            Passed = (RaisedNumber = ERR_TEST_FAILED) And (Len(RaisedDescription) > 0) _
                     And (Verdict = "FAIL")
        End If
    'Count the run and its failure
        mCertModes = mCertModes + 1
        If Not Passed Then
            mCertFailures = mCertFailures + 1
        End If
    'Write the MODE record
        CertSay "MODE=" & Mode & "; expected=" & IIf(MustPass, "PASS", "FAIL") & _
            "; error=" & CStr(RaisedNumber) & "; result=" & Verdict & _
            "; completeness=" & Completeness & "; status=" & IIf(Passed, "PASS", "FAIL")

End Sub

Public Sub EndCertification( _
    ByVal Path As String)
'
'------------------------------------------------------------------------------
'                            FINISH A CERTIFICATION
'------------------------------------------------------------------------------
' PURPOSE
'   Finish the certification: write its verdict, save the certification log
'   and only then raise the outcome
'
' WHY THIS EXISTS
'   The certification log is the evidence of the run, so it is written even
'   when the certification fails
'
' INPUTS
'   Path
'     Full path of the certification log, created or overwritten
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   completeness: COMPLETE when every run that must pass was complete.
'   CERTIFICATION=PASS needs at least one such run, COMPLETE, and every run
'   ending as its mode requires. Writes the CERTIFICATION record, releases
'   the certification and saves its log
'
' ERROR POLICY
'   A file error is re-raised after the file is closed. Otherwise raises
'   ERR_TEST_FAILED, after the log is written, unless the certification
'   passed. Raises ERR_TEST_USAGE outside a certification
'
' DEPENDENCIES
'   RequireCertification, CertSay, the private record writer of this module
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
    Dim Completed   As Boolean  'True when every required run was complete
    Dim Verdict     As String   'PASS or FAIL

'------------------------------------------------------------------------------
' JUDGE
'------------------------------------------------------------------------------
    'Refuse a call outside a certification
        RequireCertification "EndCertification"
    'Pass only a complete certification without failures
        Completed = (mCertMandatory > 0) And (mCertComplete = mCertMandatory)
        If Completed And mCertFailures = 0 Then
            Verdict = "PASS"
        Else
            Verdict = "FAIL"
        End If
    'Write the verdict and release the certification
        CertSay "CERTIFICATION=" & Verdict & "; suite=" & mCertSuite & _
            "; completeness=" & IIf(Completed, "COMPLETE", "INCOMPLETE") & _
            "; modes=" & CStr(mCertModes) & "; failures=" & CStr(mCertFailures)
        mCertActive = False

'------------------------------------------------------------------------------
' WRITE THE EVIDENCE, THEN RAISE
'------------------------------------------------------------------------------
    'Save the certification log
        WriteRecords Path, mCertLog
    'Raise a failed certification only after the log is written
        If Verdict <> "PASS" Then
            Err.Raise ERR_TEST_FAILED, "TEST_Harness.EndCertification", _
                "Certification of " & mCertSuite & " did not pass; see " & Path & "."
        End If

End Sub

'
'------------------------------------------------------------------------------
'
'                                     LOG
'
'------------------------------------------------------------------------------
'

Public Sub Say( _
    ByVal Text As String)
'
'------------------------------------------------------------------------------
'                              WRITE A LOG RECORD
'------------------------------------------------------------------------------
' PURPOSE
'   Write one log record to the Immediate window and keep it for SaveLog
'
' WHY THIS EXISTS
'   The Immediate window keeps only its last lines; the kept log is the
'   record of a run
'
' INPUTS
'   Text
'     The record
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Creates the log when there is none, adds the record and echoes it
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   Echo
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' KEEP AND ECHO
'------------------------------------------------------------------------------
    'Create the log when there is none
        If mLog Is Nothing Then
            Set mLog = New Collection
        End If
    'Keep the record and print it
        mLog.Add Text
        Echo Text

End Sub

Public Sub Echo( _
    ByVal Text As String)
'
'------------------------------------------------------------------------------
'                                 PRINT A LINE
'------------------------------------------------------------------------------
' PURPOSE
'   Print one line to the Immediate window, without keeping it in the log
'
' WHY THIS EXISTS
'   Records must show while a run progresses, on every host
'
' INPUTS
'   Text
'     The line
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Prints the line with Debug.Print
'
' ERROR POLICY
'   A host without an Immediate window, such as LibreOffice run headless,
'   raises on Debug.Print; that failure is ignored, because the kept log,
'   not the window, is the record of a run
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRINT
'------------------------------------------------------------------------------
    'Ignore a host without an Immediate window
        On Error Resume Next
        Debug.Print Text
        On Error GoTo 0

End Sub

Public Sub SaveLog( _
    ByVal Path As String)
'
'------------------------------------------------------------------------------
'                                 SAVE THE LOG
'------------------------------------------------------------------------------
' PURPOSE
'   Write every record kept since the last BeginRun or ResetTests to a text
'   file
'
' WHY THIS EXISTS
'   The Immediate window keeps only its last 200 lines; the file is the
'   complete log
'
' INPUTS
'   Path
'     Full path of the file, created or overwritten; outside the checkout or
'     in the ignored test-results/ directory
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes the kept records, one per line
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE when no record is kept; a file error is re-raised
'   after the file is closed
'
' DEPENDENCIES
'   WriteRecords
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Refuse when no record is kept
        If mLog Is Nothing Then
            Err.Raise ERR_TEST_USAGE, "TEST_Harness.SaveLog", "No log record is kept."
        End If
    'Write the records
        WriteRecords Path, mLog

End Sub

Public Function LastResult() _
    As String
'
'------------------------------------------------------------------------------
'                             READ THE LAST RESULT
'------------------------------------------------------------------------------
' PURPOSE
'   The RESULT record of the last finished or refused run
'
' WHY THIS EXISTS
'   A caller reads the outcome of a run without parsing the log
'
' INPUTS
'   None
'
' RETURNS
'   The RESULT record; empty before the first run, and after CertifyRun has
'   read it
'
' BEHAVIOR
'   Returns the kept record
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
    'Return the kept record
        LastResult = mLastResult

End Function

'
'------------------------------------------------------------------------------
'
'                                   RESULTS
'
'------------------------------------------------------------------------------
'

Public Sub RecordResult( _
    ByVal CaseId As String, _
    ByVal VariantName As String, _
    ByVal Quantity As String, _
    ByVal Dataset As String, _
    ByVal Record As String, _
    ByVal Value As Double)
'
'------------------------------------------------------------------------------
'                           RECORD A COMPUTED RESULT
'------------------------------------------------------------------------------
' PURPOSE
'   Keep one computed value of the running suite for SaveResults: the
'   result of one expected output of a test case
'
' WHY THIS EXISTS
'   tools/check_accuracy.py --results compares each result with its expected
'   value
'
' INPUTS
'   CaseId
'     The test case's id (tests/fixtures/)
'
'   VariantName
'     The variant of a file-based case, or an empty string for a JSON
'     fixture
'
'   Quantity
'     The output's quantity
'
'   Dataset
'     The dataset of a record-scoped output, or an empty string for a
'     result of the whole calculation
'
'   Record
'     The record id of a record-scoped output, or an empty string for a
'     result of the whole calculation
'
'   Value
'     The result, written with the 15 significant digits of Str$
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Keeps the result as one JSON object
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a run. A repeated result is kept; the
'   accuracy check rejects it
'
' DEPENDENCIES
'   RequireActiveRun, JsonText, JsonTextOrNull, ScopeJson, NumberText
'
' NOTES
'   None
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RECORD
'------------------------------------------------------------------------------
    'Refuse a call outside a run
        RequireActiveRun "RecordResult"
    'Keep the result as one JSON object
        mResults.Add "{""case"": " & JsonText(CaseId) & _
            ", ""variant"": " & JsonTextOrNull(VariantName) & _
            ", ""quantity"": " & JsonText(Quantity) & _
            ", ""scope"": " & ScopeJson(Dataset, Record) & _
            ", ""value"": " & NumberText(Value) & "}"

End Sub

Public Sub SaveResults( _
    ByVal Path As String)
'
'------------------------------------------------------------------------------
'                            SAVE THE RESULTS FILE
'------------------------------------------------------------------------------
' PURPOSE
'   Write the results recorded since the last BeginRun or ResetTests as the
'   results file of docs/methodology/TEST_CASES.md (Accuracy contracts), with
'   the commit of the workbook's build stamp
'
' WHY THIS EXISTS
'   The accuracy check reads the results of a run from a file that names the
'   exact commit they were computed from
'
' INPUTS
'   Path
'     Full path of the file, created or overwritten; outside the checkout or
'     in the ignored test-results/ directory
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Composes the JSON file and writes it. The file is ASCII, a subset of
'   UTF-8, with LF line endings. A workbook without a readable build stamp
'   writes an empty commit, and a build with local changes writes its
'   marked commit; the accuracy check rejects both, because the results
'   then name no exact commit
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE before the first run; a file error is re-raised
'   after the file is closed
'
' DEPENDENCIES
'   JsonText, StampedCommit, WriteText
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
    Dim Text    As String   'The whole file
    Dim Entry   As Variant  'One recorded result
    Dim Count   As Long     'Results written so far

'------------------------------------------------------------------------------
' COMPOSE
'------------------------------------------------------------------------------
    'Refuse before the first run
        If mResults Is Nothing Then
            Err.Raise ERR_TEST_USAGE, "TEST_Harness.SaveResults", "No run has recorded results."
        End If
    'Open the file object with its schema, kind and commit
        Text = "{" & vbLf & "  ""schema_version"": 1," & vbLf & _
            "  ""kind"": ""test-results""," & vbLf & _
            "  ""commit"": " & JsonText(StampedCommit()) & "," & vbLf & _
            "  ""results"": ["
    'Add every recorded result, separated by commas
        For Each Entry In mResults
            Count = Count + 1
            Text = Text & IIf(Count = 1, vbNullString, ",") & vbLf & "    " & CStr(Entry)
        Next Entry
    'Close the results list and the file object
        Text = Text & IIf(Count = 0, vbNullString, vbLf & "  ") & "]" & vbLf & "}" & vbLf

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Write the file exactly as composed
        WriteText Path, Text

End Sub

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub LogAssertion( _
    ByVal AssertionName As String, _
    ByVal Passed As Boolean, _
    ByVal Detail As String)
'
'------------------------------------------------------------------------------
'                               LOG AN ASSERTION
'------------------------------------------------------------------------------
' PURPOSE
'   Count one assertion, log its outcome and record it when it failed
'
' WHY THIS EXISTS
'   Every assertion is counted and logged the same way
'
' INPUTS
'   AssertionName
'     Stable name within the case
'
'   Passed
'     Whether the assertion held
'
'   Detail
'     Text logged when it did not
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Counts the assertion and writes its ASSERT record; a failure is also
'   recorded
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a run. An assertion outside a case is
'   counted and recorded as a failure
'
' DEPENDENCIES
'   RequireActiveRun, RecordFailure, Say
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' LOG
'------------------------------------------------------------------------------
    'Refuse an assertion outside a run, and count it
        RequireActiveRun "an assertion"
        mAssertionCount = mAssertionCount + 1
    'Record an assertion outside a case
        If Len(mCurrentCase) = 0 Then
            RecordFailure "harness.no-case", "assertion '" & AssertionName & "' outside a case"
        End If
    'Log the outcome; record a failure
        If Passed Then
            Say "ASSERT=" & mCurrentCase & "/" & AssertionName & "; status=PASS"
        Else
            Say "ASSERT=" & mCurrentCase & "/" & AssertionName & "; status=FAIL; " & Detail
            RecordFailure mCurrentCase & "/" & AssertionName, Detail
        End If

End Sub

Private Sub RecordFailure( _
    ByVal FailureId As String, _
    ByVal Detail As String)
'
'------------------------------------------------------------------------------
'                               RECORD A FAILURE
'------------------------------------------------------------------------------
' PURPOSE
'   Count and log one failure; the first one of a run is its primary failure
'
' WHY THIS EXISTS
'   The summary names the first failure, which usually explains the others
'
' INPUTS
'   FailureId
'     Identifier of the failure
'
'   Detail
'     What failed
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Counts the failure, keeps the first one and writes its FAILURE record
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   Say
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RECORD
'------------------------------------------------------------------------------
    'Count the failure and keep the first one
        mFailureCount = mFailureCount + 1
        If Len(mPrimaryFailure) = 0 Then
            mPrimaryFailure = FailureId & "; " & Detail
        End If
    'Log it
        Say "FAILURE=" & FailureId & "; " & Detail

End Sub

Private Sub RequireActiveRun( _
    ByVal Caller As String)
'
'------------------------------------------------------------------------------
'                            REQUIRE AN ACTIVE RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Refuse a lifecycle call or an assertion made outside BeginRun..EndRun
'
' WHY THIS EXISTS
'   A call out of order would count against no run
'
' INPUTS
'   Caller
'     Name of the refused call, used in the message
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Returns when a run is active
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE when no run is active
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
' VALIDATE STATE
'------------------------------------------------------------------------------
    'Refuse a call outside a run
        If Not mRunActive Then
            Err.Raise ERR_TEST_USAGE, "TEST_Harness", _
                "No test run is active for " & Caller & "; call BeginRun first."
        End If

End Sub

Private Sub RequireCertification( _
    ByVal Caller As String)
'
'------------------------------------------------------------------------------
'                       REQUIRE AN ACTIVE CERTIFICATION
'------------------------------------------------------------------------------
' PURPOSE
'   Refuse a certification call made outside
'   BeginCertification..EndCertification
'
' WHY THIS EXISTS
'   A call out of order would count against no certification
'
' INPUTS
'   Caller
'     Name of the refused call, used in the message
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Returns when a certification is active
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE when no certification is active
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE STATE
'------------------------------------------------------------------------------
    'Refuse a call outside a certification
        If Not mCertActive Then
            Err.Raise ERR_TEST_USAGE, "TEST_Harness", _
                "No certification is active for " & Caller & _
                "; call BeginCertification first."
        End If

End Sub

Private Sub CertSay( _
    ByVal Text As String)
'
'------------------------------------------------------------------------------
'                         WRITE A CERTIFICATION RECORD
'------------------------------------------------------------------------------
' PURPOSE
'   Write one certification record to the Immediate window and keep it for
'   EndCertification
'
' WHY THIS EXISTS
'   The certification log is kept apart from the run logs
'
' INPUTS
'   Text
'     The record
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Adds the record to the certification log and echoes it
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   Echo
'
' NOTES
'   None
'
' UPDATED
'   2026-11-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' KEEP AND ECHO
'------------------------------------------------------------------------------
    'Keep the record and print it
        mCertLog.Add Text
        Echo Text

End Sub

Private Sub WriteRecords( _
    ByVal Path As String, _
    ByVal Records As Collection)
'
'------------------------------------------------------------------------------
'                           WRITE RECORDS TO A FILE
'------------------------------------------------------------------------------
' PURPOSE
'   Write records to a text file, one per line
'
' WHY THIS EXISTS
'   The run log and the certification log are saved the same way
'
' INPUTS
'   Path
'     Full path of the file, created or overwritten
'
'   Records
'     The records, in order
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Opens the file for output, prints each record and closes the file
'
' ERROR POLICY
'   A file error is re-raised after the file is closed
'
' DEPENDENCIES
'   None
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
    Dim FileNumber          As Integer  'Handle of the open file
    Dim Entry               As Variant  'One record
    Dim ErrorNumber         As Long     'File error number
    Dim ErrorSource         As String   'File error source
    Dim ErrorDescription    As String   'File error description

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Open the file, then close it on a write error
        FileNumber = FreeFile
        Open Path For Output As #FileNumber
        On Error GoTo WriteFailed
    'Print each record on its own line
        For Each Entry In Records
            Print #FileNumber, CStr(Entry)
        Next Entry
        Close #FileNumber

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
WriteFailed:
    'Capture the file error
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
    'Restore normal error handling and close the file
        On Error GoTo 0
        Close #FileNumber
    'Re-raise the file error
        Err.Raise ErrorNumber, ErrorSource, ErrorDescription

End Sub

Private Sub WriteText( _
    ByVal Path As String, _
    ByVal Text As String)
'
'------------------------------------------------------------------------------
'                             WRITE TEXT TO A FILE
'------------------------------------------------------------------------------
' PURPOSE
'   Write text to a file exactly as given, without a line ending added
'
' WHY THIS EXISTS
'   The results file has LF line endings; a plain Print would add CR LF
'
' INPUTS
'   Path
'     Full path of the file, created or overwritten
'
'   Text
'     ASCII text
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Opens the file for output, prints the text and closes the file
'
' ERROR POLICY
'   A write error is re-raised after the file is closed; an Open error is
'   raised as it is, with no file open
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
    Dim ErrorNumber         As Long     'File error number
    Dim ErrorSource         As String   'File error source
    Dim ErrorDescription    As String   'File error description

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Open the file, then close it on a write error
        FileNumber = FreeFile
        Open Path For Output As #FileNumber
        On Error GoTo WriteFailed
    'The trailing semicolon keeps Print from adding CR LF
        Print #FileNumber, Text;
        Close #FileNumber

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
WriteFailed:
    'Capture the file error
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
    'Restore normal error handling and close the file
        On Error GoTo 0
        Close #FileNumber
    'Re-raise the file error
        Err.Raise ErrorNumber, ErrorSource, ErrorDescription

End Sub

Private Function StampedCommit() _
    As String
'
'------------------------------------------------------------------------------
'                           READ THE STAMPED COMMIT
'------------------------------------------------------------------------------
' PURPOSE
'   The commit of the workbook's build stamp: the value of the defined name
'   STAMP_COMMIT_NAME in ThisWorkbook, the workbook that holds this module
'
' WHY THIS EXISTS
'   The results file names the commit its results were computed from
'
' INPUTS
'   None
'
' RETURNS
'   The stamp as written by tools/devbuild/build_workbook.py, or an empty
'   string when the workbook has no such name or it cannot be read
'
' BEHAVIOR
'   Reads the value of the defined name's range as text
'
' ERROR POLICY
'   A failed read returns an empty string; nothing is raised
'
' DEPENDENCIES
'   ThisWorkbook.Names
'
' NOTES
'   None
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Read the stamp, returning an empty string when it cannot be read
        On Error GoTo ReadFailed
        StampedCommit = CStr(ThisWorkbook.Names(STAMP_COMMIT_NAME).RefersToRange.Value)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ReadFailed:
    'Return an empty stamp
        StampedCommit = vbNullString

End Function

Private Function JsonText( _
    ByVal Value As String) _
    As String
'
'------------------------------------------------------------------------------
'                              QUOTE TEXT AS JSON
'------------------------------------------------------------------------------
' PURPOSE
'   A JSON string literal of Value in ASCII
'
' WHY THIS EXISTS
'   The results file is ASCII JSON, whatever the text it holds
'
' INPUTS
'   Value
'     The text to quote
'
' RETURNS
'   Value in double quotes, with the quote, the backslash, the control
'   characters and every character above code 126 escaped
'
' BEHAVIOR
'   Escapes one character at a time; other characters are kept
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
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Index       As Long     'Position in Value
    Dim Character   As String   'Character at Index
    Dim Code        As Long     'Its UTF-16 code unit, 0 to 65535
    Dim Text        As String   'Escaped text so far

'------------------------------------------------------------------------------
' ESCAPE
'------------------------------------------------------------------------------
    'Escape the quote, the backslash, control and non-ASCII characters
        For Index = 1 To Len(Value)
            Character = Mid$(Value, Index, 1)
            Code = AscW(Character) And &HFFFF&
            If Character = """" Or Character = "\" Then
                Text = Text & "\" & Character
            ElseIf Code < 32 Or Code > 126 Then
                Text = Text & "\u" & Right$("000" & Hex$(Code), 4)
            Else
                Text = Text & Character
            End If
        Next Index
    'Return the text in quotes
        JsonText = """" & Text & """"

End Function

Private Function JsonTextOrNull( _
    ByVal Value As String) _
    As String
'
'------------------------------------------------------------------------------
'                          QUOTE TEXT AS JSON OR NULL
'------------------------------------------------------------------------------
' PURPOSE
'   JSON null for an empty string, otherwise the string literal of Value
'
' WHY THIS EXISTS
'   An optional field of the results file is null when it has no value
'
' INPUTS
'   Value
'     The text to quote
'
' RETURNS
'   null, or the JSON string literal of Value
'
' BEHAVIOR
'   Tests for an empty string, otherwise quotes the text
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   JsonText
'
' NOTES
'   None
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return null for an empty string, otherwise the quoted text
        If Len(Value) = 0 Then
            JsonTextOrNull = "null"
        Else
            JsonTextOrNull = JsonText(Value)
        End If

End Function

Private Function ScopeJson( _
    ByVal Dataset As String, _
    ByVal Record As String) _
    As String
'
'------------------------------------------------------------------------------
'                         WRITE THE SCOPE OF A RESULT
'------------------------------------------------------------------------------
' PURPOSE
'   The scope of a result as the expected-results format writes it
'
' WHY THIS EXISTS
'   tools/check_accuracy.py matches a record-scoped output only with the
'   shape {"dataset": ..., "record": ...}
'
' INPUTS
'   Dataset
'     The dataset of a record-scoped output, or an empty string
'
'   Record
'     The record id of a record-scoped output, or an empty string
'
' RETURNS
'   null for the whole calculation, otherwise {"dataset": ..., "record":
'   ...}
'
' BEHAVIOR
'   Tests for two empty strings, otherwise quotes both values
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   JsonText
'
' NOTES
'   None
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return null for the whole calculation, otherwise the dataset and record
        If Len(Dataset) = 0 And Len(Record) = 0 Then
            ScopeJson = "null"
        Else
            ScopeJson = "{""dataset"": " & JsonText(Dataset) & ", ""record"": " & _
                JsonText(Record) & "}"
        End If

End Function

Private Function ActiveWindowCaption() _
    As String
'
'------------------------------------------------------------------------------
'                        READ THE ACTIVE WINDOW CAPTION
'------------------------------------------------------------------------------
' PURPOSE
'   The caption of Excel's active window, or an empty string when there is
'   none
'
' WHY THIS EXISTS
'   A run which leaves another window active must be seen
'
' INPUTS
'   None
'
' RETURNS
'   The caption, or an empty string
'
' BEHAVIOR
'   Reads Application.ActiveWindow
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   The Excel Application object
'
' NOTES
'   None
'
' UPDATED
'   2026-11-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the caption, or an empty string without an active window
        If Application.ActiveWindow Is Nothing Then
            ActiveWindowCaption = vbNullString
        Else
            ActiveWindowCaption = Application.ActiveWindow.Caption
        End If

End Function

Private Function VerifyHostState( _
    ByRef Detail As String) _
    As Boolean
'
'------------------------------------------------------------------------------
'                            VERIFY THE HOST STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Compare the four Application properties, the number of open workbooks
'   and the caption of the active window with their BeginRun values
'
' WHY THIS EXISTS
'   A run passes only when it left the Excel state unchanged
'
' INPUTS
'   Detail
'     Set to the names of the changed values, or to why they could not be
'     compared
'
' RETURNS
'   True when all are unchanged
'
' BEHAVIOR
'   Compares each value with its snapshot and names the changed ones
'
' ERROR POLICY
'   A failed read returns False, with the error in Detail; nothing is raised
'
' DEPENDENCIES
'   ActiveWindowCaption, the Excel Application object
'
' NOTES
'   The harness never restores them: it never changed them, and the
'   procedure that did owns their restoration
'
' UPDATED
'   2026-11-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Changed     As String   'Names of the changed properties

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Refuse to compare a snapshot that was not read
        If Not mSnapshotValid Then
            Detail = "Excel state not compared: it could not be read at BeginRun"
            VerifyHostState = False
            Exit Function
        End If
    'Name each value that changed
        On Error GoTo ReadFailed
        If Application.Calculation <> mInitialCalculation Then Changed = Changed & " Calculation"
        If Application.DisplayAlerts <> mInitialAlerts Then Changed = Changed & " DisplayAlerts"
        If Application.EnableEvents <> mInitialEvents Then Changed = Changed & " EnableEvents"
        If Application.ScreenUpdating <> mInitialScreen Then Changed = Changed & " ScreenUpdating"
        If Application.Workbooks.Count <> mInitialWorkbooks Then Changed = Changed & " Workbooks"
        If ActiveWindowCaption() <> mInitialWindow Then Changed = Changed & " ActiveWindow"
        On Error GoTo 0
    'Return the outcome and its detail
        If Len(Changed) = 0 Then
            Detail = "Excel state unchanged (Calculation, DisplayAlerts, EnableEvents, " & _
                "ScreenUpdating, Workbooks, ActiveWindow)"
            VerifyHostState = True
        Else
            Detail = "Excel state changed during the run:" & Changed
            VerifyHostState = False
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ReadFailed:
    'Report the failed read
        Detail = "Excel state not compared: read failed with error " & CStr(Err.Number) & _
            " (" & Err.Description & ")"
        VerifyHostState = False

End Function

Private Function IsCaseId( _
    ByVal CaseId As String) _
    As Boolean
'
'------------------------------------------------------------------------------
'                           CHECK A CASE IDENTIFIER
'------------------------------------------------------------------------------
' PURPOSE
'   Whether CaseId is lower-case kebab-case
'
' WHY THIS EXISTS
'   Case identifiers use the form of the test-case files (TEST_CASES.md,
'   Files)
'
' INPUTS
'   CaseId
'     The identifier to check
'
' RETURNS
'   True when CaseId is a-z and 0-9 in groups joined by single hyphens
'
' BEHAVIOR
'   Rejects an empty identifier, a leading, trailing or double hyphen, and
'   any other character
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
' DECLARE
'------------------------------------------------------------------------------
    Dim Character   As String   'Character under test
    Dim i           As Long     'Position in CaseId

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
    'Reject an empty identifier or a hyphen at either end
        If Len(CaseId) = 0 Or Left$(CaseId, 1) = "-" Or Right$(CaseId, 1) = "-" Then
            IsCaseId = False
            Exit Function
        End If
    'Reject a double hyphen
        If InStr(CaseId, "--") > 0 Then
            IsCaseId = False
            Exit Function
        End If
    'Reject any character outside a-z, 0-9 and the hyphen
        For i = 1 To Len(CaseId)
            Character = Mid$(CaseId, i, 1)
            If Not (Character Like "[a-z0-9-]") Then
                IsCaseId = False
                Exit Function
            End If
        Next i
    'Accept the identifier
        IsCaseId = True

End Function

Private Function IsKnownCase( _
    ByVal CaseId As String) _
    As Boolean
'
'------------------------------------------------------------------------------
'                            CHECK FOR A KNOWN CASE
'------------------------------------------------------------------------------
' PURPOSE
'   Whether CaseId already started in this run
'
' WHY THIS EXISTS
'   A repeated case identifier is a failure
'
' INPUTS
'   CaseId
'     The identifier to look up
'
' RETURNS
'   True when CaseId is registered in this run
'
' BEHAVIOR
'   Looks the identifier up in the registered case identifiers
'
' ERROR POLICY
'   The error of a missing key is the answer; nothing is raised
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
' DECLARE
'------------------------------------------------------------------------------
    Dim Found   As Variant  'Item stored under CaseId, when present

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
    'A Collection raises error 5 for an unknown key; that error is the answer
        On Error Resume Next
        Found = mCaseIds.Item(CaseId)
        IsKnownCase = (Err.Number = 0)
        On Error GoTo 0

End Function

Private Function NumberText( _
    ByVal Value As Double) _
    As String
'
'------------------------------------------------------------------------------
'                               FORMAT A NUMBER
'------------------------------------------------------------------------------
' PURPOSE
'   A Double as text that does not depend on the regional settings
'
' WHY THIS EXISTS
'   Logs and results files must read the same on every machine
'
' INPUTS
'   Value
'     The number
'
' RETURNS
'   The number with a period as the decimal separator and a zero before
'   the point ("0.75", "-0.75")
'
' BEHAVIOR
'   Str$ always uses a period as the decimal separator. Str$ drops the zero
'   before the point (" .75"), which is put back
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
' DECLARE
'------------------------------------------------------------------------------
    Dim Text    As String   'Str$ result without its sign space

'------------------------------------------------------------------------------
' FORMAT
'------------------------------------------------------------------------------
    'Format the number without its sign space
        Text = Trim$(Str$(Value))
    'Put back the zero before the point
        If Left$(Text, 1) = "." Then
            Text = "0" & Text
        ElseIf Left$(Text, 2) = "-." Then
            Text = "-0" & Mid$(Text, 2)
        End If
    'Return the text
        NumberText = Text

End Function

Private Function HostProperty( _
    ByVal PropertyName As String) _
    As String
'
'------------------------------------------------------------------------------
'                             READ A HOST PROPERTY
'------------------------------------------------------------------------------
' PURPOSE
'   One Application property as text, or an empty string when the host
'   cannot read it
'
' WHY THIS EXISTS
'   A host that cannot read a property must still start the run
'
' INPUTS
'   PropertyName
'     Name, Version, Build or OperatingSystem
'
' RETURNS
'   The property as text, or an empty string
'
' BEHAVIOR
'   Reads the named property of the Excel Application object
'
' ERROR POLICY
'   A failed read returns an empty string; nothing is raised
'
' DEPENDENCIES
'   The Excel Application object
'
' NOTES
'   None
'
' UPDATED
'   2026-11-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Read the named property, returning an empty string when it cannot be read
        On Error GoTo ReadFailed
        Select Case PropertyName
            Case "Name"
                HostProperty = Application.Name
            Case "Version"
                HostProperty = Application.Version
            Case "Build"
                HostProperty = CStr(Application.Build)
            Case "OperatingSystem"
                HostProperty = Application.OperatingSystem
        End Select

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ReadFailed:
    'Return an empty value
        HostProperty = vbNullString

End Function

Private Function EnvironmentSummary() _
    As String
'
'------------------------------------------------------------------------------
'                           DESCRIBE THE ENVIRONMENT
'------------------------------------------------------------------------------
' PURPOSE
'   Describe the Excel host and the compiled VBA environment
'
' WHY THIS EXISTS
'   The evidence of an Excel run names the host it ran on
'
' INPUTS
'   None
'
' RETURNS
'   host, version, build, os, office (bitness) and runtime (VBA
'   generation), as key=value fields
'
' BEHAVIOR
'   The last two come from conditional compilation, not from Windows
'
' ERROR POLICY
'   A host property that cannot be read is left empty, so that the run
'   still starts and tools/check_excel_run.py rejects the evidence
'
' DEPENDENCIES
'   HostProperty
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
    Dim Bitness     As String   'Office bitness
    Dim Runtime     As String   'VBA generation

'------------------------------------------------------------------------------
' DESCRIBE
'------------------------------------------------------------------------------
    'Office bitness and VBA generation, from conditional compilation
#If Win64 Then
        Bitness = "64-bit"
#Else
        Bitness = "32-bit"
#End If
#If VBA7 Then
        Runtime = "VBA7"
#Else
        Runtime = "VBA6"
#End If
    'Join the host properties and the compiled environment
        EnvironmentSummary = "host=" & HostProperty("Name") & "; version=" & HostProperty("Version") & _
            "; build=" & HostProperty("Build") & "; os=" & HostProperty("OperatingSystem") & _
            "; office=" & Bitness & "; runtime=" & Runtime

End Function

Private Sub ResetRunState()

'
'------------------------------------------------------------------------------
'                             RESET THE RUN STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Return the counters, case identifiers, failures and flags of a run to
'   their idle values
'
' WHY THIS EXISTS
'   BeginRun and ResetTests start from the same idle state
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resets every run field; the recovery flag and the log are left to
'   callers
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
' RESET
'------------------------------------------------------------------------------
    'Mark the run inactive and clear its suite and counts
        mRunActive = False
        mSuite = vbNullString
        mExpectedCases = 0
        mExpectedAssertions = 0
        mCaseCount = 0
        mAssertionCount = 0
        mFailureCount = 0
        mErrorCount = 0
    'Clear the case, the identifiers, the failure and the snapshot
        mCurrentCase = vbNullString
        Set mCaseIds = New Collection
        mPrimaryFailure = vbNullString
        mSnapshotValid = False

End Sub
