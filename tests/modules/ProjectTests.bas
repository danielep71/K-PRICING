Attribute VB_Name = "ProjectTests"
'==============================================================================
' MODULE: ProjectTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run a deterministic, dependency-free regression suite for the neutral
'   facade/core starter and report complete evidence to the Immediate window.
'
' PUBLIC SURFACE
'   RunProjectTests is the documented test entry point. ResetProjectTests is a
'   project-private recovery command for an interrupted run; Option Private
'   Module keeps both procedures out of the external workbook automation API.
'
' DEPENDENCIES
'   ProjectFacade and the built-in VBA/Excel object models only. No external
'   references, workbook fixture, worksheet, donor project, or test framework.
'
' STATE OWNERSHIP
'   Owns private counters, failure text, suite-completeness state, and a re-entry
'   flag. Cleanup clears the re-entry flag; report state is retained until the
'   next run/reset so the final summary remains inspectable.
'
' ERROR POLICY
'   Expected facade errors are captured and asserted. Unexpected runner errors
'   are recorded, cleanup runs, and the original number/source/description is
'   re-raised. Assertion failures raise one test-suite error after reporting.
'
' WORKSHEET SAFETY
'   Reads selected Application properties to report the environment and prove
'   calculation, display-alert, event, and screen-updating state did not change.
'   It never creates, selects, activates, edits, calculates, saves, or closes
'   workbook or worksheet state.
'
' TEST SEAM
'   Tests the supported ProjectFacade surface. The fixed core boundary can be
'   exercised by future focused tests without adding production API.
'
' COMPATIBILITY
'   Excel VBA on Windows; host evidence identifies the actual Office bitness
'   and VBA generation. No external framework or workbook fixture is required.
'
' USAGE
'   Import the required production modules first, then run
'   ProjectTests.RunProjectTests from the VBE Immediate window.
'
' UPDATED
'   2026-09-09
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations; preserve the configured component visibility.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Keep harness error codes separate from the production failure contract;
    'the expected counts define when the suite is complete.
        Private Const TEST_ERROR_DIRTY_START   As Long = vbObjectError + 2060    'Refused re-entry or interrupted run
        Private Const TEST_ERROR_FAILURES      As Long = vbObjectError + 2061    'Failed or incomplete suite outcome
        Private Const EXPECTED_ASSERTIONS      As Long = 6                       'Required assertions for a complete run
        Private Const EXPECTED_CASES           As Long = 4                       'Required cases for a complete run

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'Own counters and diagnostics for one run; cleanup releases the active
    'flag while report values remain available until the next reset.
        Private mCaseCount        As Long       'Cases started in the current run
        Private mAssertionCount   As Long       'Assertions evaluated in the current run
        Private mFailureCount     As Long       'Failures recorded, including runner errors
        Private mFailureDetails   As String     'Diagnostics retained until the next reset
        Private mRunActive        As Boolean    'Re-entry guard; cleared by cleanup or reset
        Private mSuiteCompleted   As Boolean    'True only when both expected counts match


'
'------------------------------------------------------------------------------
'
'                              TEST ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub RunProjectTests()
'
'==============================================================================
'                               RunProjectTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run all four cases and report six assertions with cleanup evidence.
'
' USAGE
'   Run ProjectTests.RunProjectTests from the VBE Immediate window.
'
' STATE OWNERSHIP
'   Reject an active run before resetting counters. Snapshot four Excel
'   properties for comparison; the harness never changes those properties.
'
' ERROR POLICY
'   Record unexpected runner failures, attempt cleanup, print the summary,
'   then re-raise the saved error. Other failed outcomes raise the suite error.
'   A dirty start raises immediately and does not run cleanup.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the host snapshot separate from the saved runner error so cleanup
    'can be reported without losing the original failure.
    Dim cleanupDetail           As String           'Cleanup explanation included in the report
    Dim cleanupPassed           As Boolean          'True only when cleanup verification succeeds
    Dim initialCalculation      As XlCalculation    'Calculation mode captured before the suite
    Dim initialDisplayAlerts    As Boolean          'Display-alert setting captured before the suite
    Dim initialEnableEvents     As Boolean          'Event setting captured before the suite
    Dim initialScreenUpdating   As Boolean          'Screen-update setting captured before the suite
    Dim savedDescription        As String           'Original diagnostic preserved through cleanup
    Dim savedNumber             As Long             'Original error number for later re-raise
    Dim savedSource             As String           'Original runner error source for re-raise
    Dim stateSnapshotValid      As Boolean          'True only after all four properties were read

'------------------------------------------------------------------------------
' GUARD ENTRY
'------------------------------------------------------------------------------
    'Refuse an active run before clearing its counters or failure details.
    'An interrupted execution must be reset explicitly by the caller.
        If mRunActive Then
            Debug.Print "RESULT=FAIL_DIRTY_START; cleanup=NOT_RUN"
            Err.Raise _
                TEST_ERROR_DIRTY_START, _
                "ProjectTests.RunProjectTests", _
                "A ProjectTests run is already active. Run ResetProjectTests after an interrupted execution."
        End If

'------------------------------------------------------------------------------
' INITIALIZE RUN
'------------------------------------------------------------------------------
    'Start with empty report state, then claim the run before executing
    'anything that can fail through the shared runner handler.
        ResetRun
        mRunActive = True
        On Error GoTo RunFailed

'------------------------------------------------------------------------------
' SNAPSHOT HOST STATE
'------------------------------------------------------------------------------
    'Read all four properties before marking the snapshot valid. If a read
    'fails, cleanup must report an unavailable snapshot rather than compare
    'uninitialized values or claim that host state was unchanged.
        initialCalculation = Application.Calculation
        initialDisplayAlerts = Application.DisplayAlerts
        initialEnableEvents = Application.EnableEvents
        initialScreenUpdating = Application.ScreenUpdating
        stateSnapshotValid = True

'------------------------------------------------------------------------------
' RUN SUITE
'------------------------------------------------------------------------------
    'Run cases in the evidence contract order; each case records unexpected
    'errors so later cases can still contribute to the report.
        PrintEnvironment
        TestExactEquality
        TestTolerance
        TestExpectedError
        TestRepeatability

    'Require both counts so an early return cannot produce a complete PASS.
        mSuiteCompleted = _
            (mCaseCount = EXPECTED_CASES) And _
            (mAssertionCount = EXPECTED_ASSERTIONS)
        If Not mSuiteCompleted Then
            RecordFailure _
                "suite.completeness", _
                "expected cases=" & CStr(EXPECTED_CASES) & _
                    ", assertions=" & CStr(EXPECTED_ASSERTIONS) & _
                    "; observed cases=" & CStr(mCaseCount) & _
                    ", assertions=" & CStr(mAssertionCount)
        End If

'------------------------------------------------------------------------------
' CLEANUP AND REPORT
'------------------------------------------------------------------------------
CleanExit:
    'Disable the runner handler before cleanup and reporting. A later raise
    'must escape to the caller instead of re-entering the runner handler.
        On Error GoTo 0
        cleanupPassed = CleanupRun( _
            cleanupDetail, _
            stateSnapshotValid, _
            initialCalculation, _
            initialDisplayAlerts, _
            initialEnableEvents, _
            initialScreenUpdating)
        PrintSummary cleanupPassed, cleanupDetail

    'Propagate the original runner error only after cleanup and reporting.
        If savedNumber <> 0 Then
            Err.Raise savedNumber, savedSource, savedDescription
        End If

    'Make assertion, cleanup and completeness failures visible to callers
    'even when no unexpected runner error was saved.
        If mFailureCount <> 0 Or Not cleanupPassed Or Not mSuiteCompleted Then
            Err.Raise _
                TEST_ERROR_FAILURES, _
                "ProjectTests.RunProjectTests", _
                "Regression failed; review the Immediate window report."
        End If
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE RUNNER ERROR
'------------------------------------------------------------------------------
RunFailed:
    'Save the escaping error before recording its diagnostic; Resume then
    'leaves the active handler through the common cleanup path.
        savedNumber = Err.Number
        savedSource = Err.Source
        savedDescription = Err.Description
        RecordFailure _
            "runner.unexpected", _
            "error=" & CStr(savedNumber) & _
                "; source=" & savedSource & _
                "; description=" & savedDescription
        Resume CleanExit

End Sub


Public Sub ResetProjectTests()
'
'==============================================================================
'                              ResetProjectTests
'------------------------------------------------------------------------------
' PURPOSE
'   Recover harness-owned state after an interrupted execution.
'
' USAGE
'   Run only after the previous execution has stopped; then rerun the suite.
'
' SIDE EFFECTS
'   Clear counters, failure text, completion, and the active-run flag.
'   Print confirmation; do not alter Excel application properties.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' RESET
'------------------------------------------------------------------------------
    'Recover only harness-owned state after execution has stopped; the
    'confirmation distinguishes an explicit reset from a completed test run.
        ResetRun
        Debug.Print "PROJECT TESTS RESET; run_active=no; counters=0"

End Sub


'
'------------------------------------------------------------------------------
'
'                               REGRESSION CASES
'
'------------------------------------------------------------------------------
'

Private Sub TestExactEquality()
'
'==============================================================================
'                              TestExactEquality
'------------------------------------------------------------------------------
' PURPOSE
'   Check one exactly representable quotient through the public facade.
'
' ERROR POLICY
'   Record unexpected case errors and let the remaining suite continue.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Use a binary-exact quotient so this case checks equality without a
    'tolerance that could hide an incorrect result.
        On Error GoTo CaseFailed

        BeginCase "ratio.exact"
        AssertEqualDouble _
            "ProjectRatio(10, 4)", _
            2.5, _
            ProjectFacade.ProjectRatio(10#, 4#)
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Record this case as failed without aborting the remaining cases.
        RecordUnexpectedCaseError "ratio.exact"

End Sub


Private Sub TestTolerance()
'
'==============================================================================
'                                TestTolerance
'------------------------------------------------------------------------------
' PURPOSE
'   Check a recurring quotient against an explicit absolute tolerance.
'
' ERROR POLICY
'   Record unexpected case errors and let the remaining suite continue.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Use a recurring quotient to exercise the absolute-tolerance assertion
    'with an explicit reference value and bound.
        On Error GoTo CaseFailed

        BeginCase "ratio.tolerance"
        AssertNear _
            "ProjectRatio(1, 3)", _
            0.333333333333333, _
            ProjectFacade.ProjectRatio(1#, 3#), _
            0.000000000001
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Keep the unexpected failure associated with the tolerance case.
        RecordUnexpectedCaseError "ratio.tolerance"

End Sub


Private Sub TestExpectedError()
'
'==============================================================================
'                              TestExpectedError
'------------------------------------------------------------------------------
' PURPOSE
'   Verify the zero-denominator error number, source, and description.
'
' ERROR POLICY
'   Capture the expected error before further calls can alter Err. A normal
'   return is a failure; errors during verification are unexpected failures.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Capture the facade error as values before calling assertion helpers.
    Dim actualDescription   As String    'Error description captured from the facade
    Dim actualNumber        As Long      'Error number captured from the facade
    Dim actualSource        As String    'Error source captured from the facade
    Dim ignored             As Double    'Unexpected return value if the call fails to raise

'------------------------------------------------------------------------------
' CALL EXPECTED FAILURE
'------------------------------------------------------------------------------
    'A zero denominator must raise. Reaching the next statement is itself
    'a failure, regardless of the returned numeric value.
        BeginCase "ratio.zero-denominator"
        On Error GoTo ExpectedError

        ignored = ProjectFacade.ProjectRatio(1#, 0#)
        RecordFailure _
            "ratio.zero-denominator.raises", _
            "Expected an error, but the call returned " & CStr(ignored) & "."
        Exit Sub

'------------------------------------------------------------------------------
' VERIFY EXPECTED ERROR
'------------------------------------------------------------------------------
ExpectedError:
    'Snapshot the expected error before resetting error-handling mode and
    'installing a separate handler for assertion-time failures.
        actualNumber = Err.Number
        actualSource = Err.Source
        actualDescription = Err.Description
        On Error GoTo 0
        On Error GoTo CaseFailed

    'Check captured values rather than the mutable live Err object.
        AssertExpectedError _
            "ratio.zero-denominator", _
            ProjectFacade.PROJECT_ERROR_ZERO_DENOMINATOR, _
            "ProjectFacade.ProjectRatio", _
            "Denominator must not be zero.", _
            actualNumber, _
            actualSource, _
            actualDescription
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Treat a failure during verification as unexpected, not as evidence
    'that the original facade call raised the correct error.
        RecordUnexpectedCaseError "ratio.zero-denominator"

End Sub


Private Sub TestRepeatability()
'
'==============================================================================
'                              TestRepeatability
'------------------------------------------------------------------------------
' PURPOSE
'   Verify identical facade results for two calls with the same inputs.
'
' ERROR POLICY
'   Record unexpected case errors and let the remaining suite continue.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Retain both results so the comparison observes two separate calls.
    Dim firstResult    As Double    'Baseline result for the repeatability check
    Dim secondResult   As Double    'Result of the identical second call

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Call the facade twice with identical signed inputs to detect a result
    'that depends on residual state from an earlier invocation.
        On Error GoTo CaseFailed

        BeginCase "ratio.repeatability"
        firstResult = ProjectFacade.ProjectRatio(-9#, 4#)
        secondResult = ProjectFacade.ProjectRatio(-9#, 4#)
        AssertEqualDouble "Repeated calls", firstResult, secondResult
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE CASE ERROR
'------------------------------------------------------------------------------
CaseFailed:
    'Report the failed repeatability case and allow suite finalization.
        RecordUnexpectedCaseError "ratio.repeatability"

End Sub


'
'------------------------------------------------------------------------------
'
'                          CASE AND ASSERTION HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub BeginCase( _
    ByVal caseName As String)
'
'==============================================================================
'                                  BeginCase
'------------------------------------------------------------------------------
' PURPOSE
'   Register a started case in both the counter and machine-readable log.
'
' INPUTS
'   caseName: stable identifier used by the evidence contract.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' REGISTER CASE
'------------------------------------------------------------------------------
    'Count a case when it starts, even if its first assertion later fails;
    'the matching log record identifies which case was reached.
        mCaseCount = mCaseCount + 1
        Debug.Print "CASE=" & caseName

End Sub


Private Sub AssertEqualDouble( _
    ByVal assertionName As String, _
    ByVal expected As Double, _
    ByVal actual As Double)
'
'==============================================================================
'                              AssertEqualDouble
'------------------------------------------------------------------------------
' PURPOSE
'   Count one assertion and require exact Double equality.
'
' INPUTS
'   assertionName identifies the check; expected and actual are compared.
'
' ERROR POLICY
'   Append a mismatch to the run report without raising an assertion error.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Count the assertion before evaluating it, then record any exact-value
    'mismatch without terminating the suite.
        mAssertionCount = mAssertionCount + 1
        If actual <> expected Then
            RecordFailure _
                assertionName, _
                "expected=" & CStr(expected) & "; actual=" & CStr(actual)
        End If

End Sub


Private Sub AssertNear( _
    ByVal assertionName As String, _
    ByVal expected As Double, _
    ByVal actual As Double, _
    ByVal tolerance As Double)
'
'==============================================================================
'                                  AssertNear
'------------------------------------------------------------------------------
' PURPOSE
'   Count one assertion and apply an absolute Double tolerance.
'
' INPUTS
'   assertionName identifies the check; expected and actual are compared.
'   tolerance must be nonnegative and is inclusive at the boundary.
'
' ERROR POLICY
'   Record a negative tolerance or an excessive difference as one failure.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Use an absolute error bound. A negative tolerance is a failed test
    'contract; equality at the permitted bound remains a pass.
        mAssertionCount = mAssertionCount + 1
        If tolerance < 0# Then
            RecordFailure assertionName, "Tolerance must not be negative."
        ElseIf Abs(actual - expected) > tolerance Then
            RecordFailure _
                assertionName, _
                "expected=" & CStr(expected) & _
                    "; actual=" & CStr(actual) & _
                    "; tolerance=" & CStr(tolerance)
        End If

End Sub


Private Sub AssertExpectedError( _
    ByVal assertionName As String, _
    ByVal expectedNumber As Long, _
    ByVal expectedSource As String, _
    ByVal expectedDescription As String, _
    ByVal actualNumber As Long, _
    ByVal actualSource As String, _
    ByVal actualDescription As String)
'
'==============================================================================
'                             AssertExpectedError
'------------------------------------------------------------------------------
' PURPOSE
'   Verify the three stable fields of a captured facade error.
'
' INPUTS
'   assertionName prefixes three checks. The expected and actual number,
'   source, and description are supplied as captured scalar values.
'
' SIDE EFFECTS
'   Delegate to three assertions; do not read the live Err object.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSERT ERROR CONTRACT
'------------------------------------------------------------------------------
    'Keep number, source and description as three independent assertions
    'so a partially correct error cannot satisfy the complete contract.
        AssertEqualLong _
            assertionName & ".number", _
            expectedNumber, _
            actualNumber
        AssertEqualString _
            assertionName & ".source", _
            expectedSource, _
            actualSource
        AssertEqualString _
            assertionName & ".description", _
            expectedDescription, _
            actualDescription

End Sub


Private Sub AssertEqualLong( _
    ByVal assertionName As String, _
    ByVal expected As Long, _
    ByVal actual As Long)
'
'==============================================================================
'                               AssertEqualLong
'------------------------------------------------------------------------------
' PURPOSE
'   Count one assertion and require exact Long equality.
'
' INPUTS
'   assertionName identifies the check; expected and actual are compared.
'
' ERROR POLICY
'   Append a mismatch to the run report without raising an assertion error.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Compare the integer contract exactly; a mismatched error number must
    'not be accepted because the source or description happens to match.
        mAssertionCount = mAssertionCount + 1
        If actual <> expected Then
            RecordFailure _
                assertionName, _
                "expected=" & CStr(expected) & "; actual=" & CStr(actual)
        End If

End Sub


Private Sub AssertEqualString( _
    ByVal assertionName As String, _
    ByVal expected As String, _
    ByVal actual As String)
'
'==============================================================================
'                              AssertEqualString
'------------------------------------------------------------------------------
' PURPOSE
'   Count one assertion and require binary, case-sensitive string equality.
'
' INPUTS
'   assertionName identifies the check; expected and actual are compared.
'
' ERROR POLICY
'   Append a mismatch to the run report without raising an assertion error.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
    'Use binary comparison so case changes in the error source or message
    'remain observable contract differences, independent of text settings.
        mAssertionCount = mAssertionCount + 1
        If StrComp(actual, expected, vbBinaryCompare) <> 0 Then
            RecordFailure _
                assertionName, _
                "expected=""" & expected & """; actual=""" & actual & """"
        End If

End Sub


'
'------------------------------------------------------------------------------
'
'                              FAILURE RECORDING
'
'------------------------------------------------------------------------------
'

Private Sub RecordUnexpectedCaseError( _
    ByVal caseName As String)
'
'==============================================================================
'                          RecordUnexpectedCaseError
'------------------------------------------------------------------------------
' PURPOSE
'   Preserve the active case error before recording its diagnostic fields.
'
' INPUTS
'   caseName: stable case identifier; Err supplies the original failure.
'
' SIDE EFFECTS
'   Add a failure record without incrementing the assertion count.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Retain Err fields before any error-state reset or reporting call.
    Dim errorDescription   As String    'Case diagnostic captured before recording
    Dim errorNumber        As Long      'Case error number captured before recording
    Dim errorSource        As String    'Case error source captured before recording

'------------------------------------------------------------------------------
' CAPTURE ERROR
'------------------------------------------------------------------------------
    'Read the original diagnostic before On Error GoTo 0 can clear Err.
        errorNumber = Err.Number
        errorSource = Err.Source
        errorDescription = Err.Description
        On Error GoTo 0

'------------------------------------------------------------------------------
' RECORD DIAGNOSTIC
'------------------------------------------------------------------------------
    'Add the case-qualified failure without inventing a completed assertion
    'for an operation that raised before its check could run.
        RecordFailure _
            caseName & ".unexpected", _
            "error=" & CStr(errorNumber) & _
                "; source=" & errorSource & _
                "; description=" & errorDescription

End Sub


Private Sub RecordFailure( _
    ByVal assertionName As String, _
    ByVal detail As String)
'
'==============================================================================
'                                RecordFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Accumulate failure details for the final report.
'
' INPUTS
'   assertionName identifies the failure; detail supplies its diagnostic.
'
' STATE OWNERSHIP
'   Increment the failure count and append text to the owned report buffer.
'   Do not change case or assertion counts.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' RECORD FAILURE
'------------------------------------------------------------------------------
    'Retain every diagnostic in arrival order. A separator belongs only
    'between records so the final report has no leading empty entry.
        mFailureCount = mFailureCount + 1
        If Len(mFailureDetails) > 0 Then
            mFailureDetails = mFailureDetails & vbNewLine
        End If
        mFailureDetails = mFailureDetails & assertionName & ": " & detail

End Sub


'
'------------------------------------------------------------------------------
'
'                          CLEANUP AND HOST EVIDENCE
'
'------------------------------------------------------------------------------
'

Private Function CleanupRun( _
    ByRef cleanupDetail As String, _
    ByVal stateSnapshotValid As Boolean, _
    ByVal initialCalculation As XlCalculation, _
    ByVal initialDisplayAlerts As Boolean, _
    ByVal initialEnableEvents As Boolean, _
    ByVal initialScreenUpdating As Boolean) _
    As Boolean
'
'==============================================================================
'                                  CleanupRun
'------------------------------------------------------------------------------
' PURPOSE
'   Release the run flag and verify that observed Excel state is unchanged.
'
' INPUTS
'   stateSnapshotValid indicates whether all four initial properties were
'   read. The initial values belong to the current run only.
'
' RETURNS
'   Boolean success and a ByRef cleanupDetail for the final report.
'
' ERROR POLICY
'   Return False for an unavailable snapshot, changed state, or a cleanup
'   error. Report the reason; do not restore properties the harness never owned.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Track the combined comparison without taking ownership of host state.
    Dim excelStateUnchanged   As Boolean    'All four observed properties match the snapshot

'------------------------------------------------------------------------------
' VERIFY CLEANUP
'------------------------------------------------------------------------------
    'Release the re-entry flag first, even when no complete host snapshot
    'is available to verify the remaining cleanup contract.
        On Error GoTo CleanupFailed

        mRunActive = False
        If Not stateSnapshotValid Then
            cleanupDetail = "run flag cleared; Excel state snapshot unavailable"
            CleanupRun = False
            Exit Function
        End If

    'Compare only properties that were captured successfully. The harness
    'does not restore them because it never changed or owned them.
        excelStateUnchanged = _
            (Application.Calculation = initialCalculation) And _
            (Application.DisplayAlerts = initialDisplayAlerts) And _
            (Application.EnableEvents = initialEnableEvents) And _
            (Application.ScreenUpdating = initialScreenUpdating)

    'Return the combined result and retain a reason for either outcome.
        CleanupRun = excelStateUnchanged
        If excelStateUnchanged Then
            cleanupDetail = _
                "run flag cleared; verified Excel state unchanged " & _
                "(calculation/display alerts/events/screen updating)"
        Else
            cleanupDetail = "run flag cleared; Excel application state changed during test run"
        End If
        Exit Function

'------------------------------------------------------------------------------
' HANDLE CLEANUP ERROR
'------------------------------------------------------------------------------
CleanupFailed:
    'Describe a failed verification and return False. The runner retains
    'its own saved diagnostic for any later re-raise.
        cleanupDetail = _
            "cleanup error=" & CStr(Err.Number) & _
            "; description=" & Err.Description
        Err.Clear
        CleanupRun = False

End Function


Private Sub PrintEnvironment()
'
'==============================================================================
'                               PrintEnvironment
'------------------------------------------------------------------------------
' PURPOSE
'   Write the report preamble and current Excel environment.
'
' ERROR POLICY
'   Host-property read failures propagate to the runner error path.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' REPORT ENVIRONMENT
'------------------------------------------------------------------------------
    'Write the preamble before the host fields so retained output can be
    'recognized and attributed to the environment that produced it.
        Debug.Print "PROJECT TESTS"
        Debug.Print "ENVIRONMENT=" & EnvironmentSummary()

End Sub


Private Function EnvironmentSummary() _
    As String
'
'==============================================================================
'                              EnvironmentSummary
'------------------------------------------------------------------------------
' PURPOSE
'   Describe the current Excel host and compiled VBA environment.
'
' RETURNS
'   Machine-readable host, version, operating system, Office bitness, and
'   VBA generation fields. Compile-time branches select the last two values.
'
' ERROR POLICY
'   Host-property read failures propagate to the runner error path.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep compile-time properties distinct from values read from Excel.
    Dim bitness         As String    'Office bitness selected by conditional compilation
    Dim vbaGeneration   As String    'Compiled VBA generation reported as text

'------------------------------------------------------------------------------
' RESOLVE COMPILED ENVIRONMENT
'------------------------------------------------------------------------------
    'Describe the running Office/VBA build from compiler constants; do not
    'infer Office bitness from the Windows operating-system description.
#If Win64 Then
        bitness = "64-bit"
#Else
        bitness = "32-bit"
#End If

#If VBA7 Then
        vbaGeneration = "VBA7+"
#Else
        vbaGeneration = "legacy VBA"
#End If

'------------------------------------------------------------------------------
' BUILD ENVIRONMENT RECORD
'------------------------------------------------------------------------------
    'Combine live host properties with the compiled environment fields,
    'keeping the evidence parser keys and order stable.
        EnvironmentSummary = _
            "host=" & Application.Name & _
            "; version=" & Application.Version & _
            "; os=" & Application.OperatingSystem & _
            "; office=" & bitness & _
            "; runtime=" & vbaGeneration

End Function


Private Sub PrintSummary( _
    ByVal cleanupPassed As Boolean, _
    ByVal cleanupDetail As String)
'
'==============================================================================
'                                 PrintSummary
'------------------------------------------------------------------------------
' PURPOSE
'   Report counts, cleanup, failure details, and the complete run verdict.
'
' INPUTS
'   cleanupPassed and cleanupDetail: the outcome of CleanupRun.
'
' STATE OWNERSHIP
'   Read the retained counters and completion flag without resetting them.
'   PASS requires zero failures, successful cleanup, and a complete suite.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Resolve one verdict before writing the machine-readable report.
    Dim verdict   As String    'Final status derived from failures and cleanup

'------------------------------------------------------------------------------
' RESOLVE VERDICT
'------------------------------------------------------------------------------
    'A zero failure count is insufficient: cleanup must succeed and the
    'suite must have reached both required counts before reporting PASS.
        If mFailureCount = 0 And cleanupPassed And mSuiteCompleted Then
            verdict = "PASS"
        Else
            verdict = "FAIL"
        End If

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Write the retained counts and cleanup first, then any diagnostics and
    'the final result record consumed by the evidence validator.
        Debug.Print "CASES=" & CStr(mCaseCount)
        Debug.Print "ASSERTIONS=" & CStr(mAssertionCount)
        Debug.Print "FAILURES=" & CStr(mFailureCount)
        Debug.Print "CLEANUP=" & IIf(cleanupPassed, "PASS", "FAIL") & _
            "; detail=" & cleanupDetail

    'Omit the diagnostic record on a clean run; retain all details on failure.
        If Len(mFailureDetails) > 0 Then
            Debug.Print "FAILURE_DETAILS=" & mFailureDetails
        End If

        Debug.Print _
            "RESULT=" & verdict & _
            "; completeness=" & IIf(mSuiteCompleted, "COMPLETE", "INCOMPLETE") & _
            "; cases=" & CStr(mCaseCount) & _
            "; assertions=" & CStr(mAssertionCount) & _
            "; failures=" & CStr(mFailureCount) & _
            "; cleanup=" & IIf(cleanupPassed, "PASS", "FAIL")

End Sub


'
'------------------------------------------------------------------------------
'
'                              OWNED STATE RESET
'
'------------------------------------------------------------------------------
'

Private Sub ResetRun()
'
'==============================================================================
'                                   ResetRun
'------------------------------------------------------------------------------
' PURPOSE
'   Return all harness-owned state to its idle baseline.
'
' STATE OWNERSHIP
'   Clear counters, failure text, active-run state, and completion.
'   This helper never changes Excel application properties.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' RESET
'------------------------------------------------------------------------------
    'Clear all harness-owned state together so no count, diagnostic, or
    'completion flag can leak from a previous execution into the next one.
        mCaseCount = 0
        mAssertionCount = 0
        mFailureCount = 0
        mFailureDetails = vbNullString
        mRunActive = False
        mSuiteCompleted = False

End Sub
