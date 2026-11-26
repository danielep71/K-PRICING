Attribute VB_Name = "TEST_Ranges"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: TEST_Ranges
'------------------------------------------------------------------------------
' PURPOSE
'   Tests of M_Ranges: the description and the read of Range arguments, and the hand-off of
'   every other argument to CORE_Array
'
' WHY THIS EXISTS
'   M_Ranges is the one module that reads a Range, so its paths are tested on real Ranges of
'   the workbook, which only a host with Excel's object model provides
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   RunRangeTests runs the suite. Option Private Module keeps it out of the workbook's external
'   surface. The cases read Ranges of the first worksheet, whose cells they leave unchanged, and
'   write nothing. CertifyRanges certifies its normal run, as .github/excel-evidence-policy.json
'   lists it
'
' ERROR POLICY
'   As TEST_Harness.EndRun: a run that does not pass raises TEST_Harness.ERR_TEST_FAILED,
'   or the unexpected error that escaped the runner. A refused start raises
'   TEST_Harness.ERR_TEST_DIRTY_START
'
' DEPENDENCIES
'   M_Ranges, CORE_Err, TEST_Harness; the Excel object model (Worksheets, Range)
'
' NOTES
'   In a development workbook holding the modules of src/ and tests/modules/, type
'   TEST_Ranges.RunRangeTests in the Immediate window. The cases need the cells Z1000 and
'   X1000:Z1001 of the first worksheet to be blank
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'A complete run starts every case and evaluates every assertion once
    Private Const SUITE_NAME            As String = "ranges"            'Suite name in the run log
    Private Const EXPECTED_CASES        As Long = 5                 'Cases a complete run starts
    Private Const EXPECTED_ASSERTIONS   As Long = 62                'Assertions a complete run evaluates


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

Public Sub RunRangeTests()

'
'------------------------------------------------------------------------------
'                               RUN RANGES TESTS
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case in a fixed order between BeginRun and EndRun. A passing run ends with
'   RESULT=PASS; suite=ranges; operation=PASS; completeness=COMPLETE; cases=5;
'   assertions=62; failures=0; cleanup=PASS; recovery=NONE
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
        CaseDescribeRanges
        CaseReadRanges
        CaseReadScalars
        CaseReadControls
        CaseRangeElements

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

Public Sub CertifyRanges( _
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
'   Runs RunRangeTests, saves its log, certifies the run, then saves the certification log.
'
' ERROR POLICY
'   TEST_Harness.EndCertification raises when the certification did not pass.
'
' DEPENDENCIES
'   RunRangeTests, TEST_Harness
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
        RunRangeTests
        RunErrorNumber = Err.Number
        RunErrorDescription = Err.Description
    'Restore normal error handling
        On Error GoTo 0
    'Save the log and certify the run
        TEST_Harness.SaveLog Folder & "\ranges.log"
        TEST_Harness.CertifyRun "NORMAL", True, RunErrorNumber, RunErrorDescription
    'Save the certification log, raising when the certification did not pass
        TEST_Harness.EndCertification Folder & "\ranges-certification.log"

End Sub

'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseDescribeRanges()

'
'------------------------------------------------------------------------------
'                         CHECK THE RANGE DESCRIPTION
'------------------------------------------------------------------------------
' PURPOSE
'   A single-area Range is described by its counts; several areas or another object are not. 28 assertions.
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
'   Describes a cell, a block, a whole column and a two-area Range, then a plain array and another object
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   M_Ranges, CORE_Err, TEST_Harness; the Excel object model (Worksheets, Range)
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
    Dim Sheet             As Worksheet    'The first worksheet
    Dim Item              As Collection   'An object that is not a Range
    Dim Table()           As Variant      'A 2 by 3 array

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "describe-ranges"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Read Ranges of the first worksheet
        Set Sheet = ThisWorkbook.Worksheets(1)
    'A cell is a scalar, a block keeps its rows and columns, a column is read at its full height
        ProbeShape Sheet.Range("B2"), cdNone, skScalar, 1, 1
        TEST_Harness.AssertTrue "cell-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "cell-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "cell-kind", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertTrue "cell-size", mFact(4, 1), mFact(4, 2)
        ProbeShape Sheet.Range("A1:C2"), cdNone, skArray, 2, 3
        TEST_Harness.AssertTrue "block-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "block-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "block-kind", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertTrue "block-size", mFact(4, 1), mFact(4, 2)
        ProbeShape Sheet.Range("A1:D1"), cdNone, skArray, 1, 4
        TEST_Harness.AssertTrue "row-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "row-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "row-kind", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertTrue "row-size", mFact(4, 1), mFact(4, 2)
        ProbeShape Sheet.Range("A:A"), cdNone, skArray, Sheet.Rows.Count, 1
        TEST_Harness.AssertTrue "whole-column-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "whole-column-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "whole-column-kind", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertTrue "whole-column-size", mFact(4, 1), mFact(4, 2)
    'Two areas are not one rectangle
        ProbeShape Sheet.Range("A1:A2,C1:C2"), cdShapeUnsupported, 0, 0, 0
        TEST_Harness.AssertTrue "two-areas-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "two-areas-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "two-areas-kind", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertTrue "two-areas-size", mFact(4, 1), mFact(4, 2)
    'A plain array is the core's, and another object is unsupported
        ReDim Table(1 To 2, 1 To 3)
        ProbeShape Table, cdNone, skArray, 2, 3
        TEST_Harness.AssertTrue "plain-array-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "plain-array-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "plain-array-kind", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertTrue "plain-array-size", mFact(4, 1), mFact(4, 2)
        Set Item = New Collection
        ProbeShape Item, cdShapeUnsupported, 0, 0, 0
        TEST_Harness.AssertTrue "other-object-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "other-object-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "other-object-kind", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertTrue "other-object-size", mFact(4, 1), mFact(4, 2)
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

Private Sub CaseReadRanges()

'
'------------------------------------------------------------------------------
'                             CHECK THE RANGE READ
'------------------------------------------------------------------------------
' PURPOSE
'   A Range is read once into a plain payload, and only after its size passes the cap. 9 assertions.
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
'   Reads a cell and a block of empty cells, then a Range over the cap and a two-area Range
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   M_Ranges, CORE_Err, TEST_Harness; the Excel object model (Worksheets, Range)
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
    Dim Sheet             As Worksheet    'The first worksheet
    Dim Payload           As Variant      'Payload read
    Dim Kind              As KPRArgShape  'Kind reported
    Dim RowCount          As Long         'Rows reported
    Dim ColCount          As Long         'Columns reported
    Dim Condition         As KPRCondition 'Condition reported

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "read-ranges"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Read Ranges of the first worksheet
        Set Sheet = ThisWorkbook.Worksheets(1)
    'A blank cell reads as Empty, and a block as a 1-based two-dimensional array
        TEST_Harness.AssertTrue "cell", _
            M_Ranges.TryReadArgument(Sheet.Range("Z1000"), Payload, Kind, RowCount, ColCount, Condition) And _
            (Kind = skScalar), _
            "a cell was not read as a scalar"
        TEST_Harness.AssertTrue "cell-empty", IsEmpty(Payload), "a blank cell was not read as Empty"
        TEST_Harness.AssertTrue "block", _
            M_Ranges.TryReadArgument(Sheet.Range("X1000:Z1001"), Payload, Kind, RowCount, ColCount, Condition), _
            "a block was not read"
        TEST_Harness.AssertTrue "block-bounds", _
            (LBound(Payload, 1) = 1) And _
            (UBound(Payload, 1) = 2) And _
            (LBound(Payload, 2) = 1) And _
            (UBound(Payload, 2) = 3), _
            "the block is not 2 by 3 from 1"
    'Over the cap, the Range is refused from its counts
        TEST_Harness.AssertTrue "over-cap", _
            Not M_Ranges.TryReadArgument(Sheet.Range("A1:A100001"), Payload, Kind, RowCount, ColCount, Condition), _
            "100,001 cells were read"
        TEST_Harness.AssertEqualText "over-cap-condition", "CAPACITY_EXCEEDED", CORE_Err.ConditionName(Condition)
        TEST_Harness.AssertTrue "at-cap", _
            M_Ranges.TryReadArgument(Sheet.Range("A1:A100000"), Payload, Kind, RowCount, ColCount, Condition), _
            "100,000 cells were refused"
    'Two areas are not read
        TEST_Harness.AssertTrue "two-areas", _
            Not M_Ranges.TryReadArgument(Sheet.Range("A1:A2,C1:C2"), Payload, Kind, RowCount, ColCount, Condition), _
            "two areas were read"
        TEST_Harness.AssertEqualText "two-areas-condition", "SHAPE_UNSUPPORTED", CORE_Err.ConditionName(Condition)

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

Private Sub CaseReadScalars()

'
'------------------------------------------------------------------------------
'                            CHECK THE SCALAR READ
'------------------------------------------------------------------------------
' PURPOSE
'   A one-cell Range reads as its value; a larger Range is not a scalar. 3 assertions.
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
'   Reads a cell, a two-cell Range and a plain one-element array
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   M_Ranges, CORE_Err, TEST_Harness; the Excel object model (Worksheets, Range)
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
    Dim Sheet             As Worksheet    'The first worksheet
    Dim ScalarOut         As Variant      'Scalar read

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "read-scalars"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Read Ranges of the first worksheet
        Set Sheet = ThisWorkbook.Worksheets(1)
    'One cell is a scalar; two cells are not
        TEST_Harness.AssertTrue "cell", _
            M_Ranges.TryReadScalar(Sheet.Range("Z1000"), ScalarOut) And IsEmpty(ScalarOut), _
            "a blank cell was not read as Empty"
        TEST_Harness.AssertTrue "two-cells", Not M_Ranges.TryReadScalar(Sheet.Range("Y1000:Z1000"), ScalarOut), "two cells were read as a scalar"
    'A plain value is the core's
        TEST_Harness.AssertTrue "plain", _
            M_Ranges.TryReadScalar(Array(8), ScalarOut) And (ScalarOut = 8), _
            "a one-element array was not unwrapped"

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

Private Sub CaseReadControls()

'
'------------------------------------------------------------------------------
'                            CHECK THE CONTROL READ
'------------------------------------------------------------------------------
' PURPOSE
'   A one-cell Range is a control; a larger one is not a scalar, whatever its size. 15 assertions.
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
'   Reads controls from a cell, a block, a Range over the cap, two areas and a plain array
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   M_Ranges, CORE_Err, TEST_Harness; the Excel object model (Worksheets, Range)
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
    Dim Sheet             As Worksheet    'The first worksheet

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "read-controls"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Read Ranges of the first worksheet
        Set Sheet = ThisWorkbook.Worksheets(1)
    'One cell is a control, and a blank one reads as Empty
        ProbeControl Sheet.Range("Z1000"), cdNone, ""
        TEST_Harness.AssertTrue "cell-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "cell-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "cell-value", mFact(3, 1), mFact(3, 2)
    'More cells are not a scalar, even over the cap
        ProbeControl Sheet.Range("A1:B2"), cdControlNotScalar, "UNTOUCHED"
        TEST_Harness.AssertTrue "block-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "block-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "block-value", mFact(3, 1), mFact(3, 2)
        ProbeControl Sheet.Range("A1:A100001"), cdControlNotScalar, "UNTOUCHED"
        TEST_Harness.AssertTrue "over-cap-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "over-cap-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "over-cap-value", mFact(3, 1), mFact(3, 2)
    'Two areas are not a shape, and a plain array is the core's
        ProbeControl Sheet.Range("A1,C1"), cdShapeUnsupported, "UNTOUCHED"
        TEST_Harness.AssertTrue "two-areas-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "two-areas-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "two-areas-value", mFact(3, 1), mFact(3, 2)
        ProbeControl Array("CEILING"), cdNone, "CEILING"
        TEST_Harness.AssertTrue "plain-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "plain-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "plain-value", mFact(3, 1), mFact(3, 2)

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

Private Sub CaseRangeElements()

'
'------------------------------------------------------------------------------
'                        CHECK A RANGE INSIDE AN ARRAY
'------------------------------------------------------------------------------
' PURPOSE
'   A Range held inside an array is an object element, refused on every path. 7 assertions.
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
'   Puts a one-cell Range inside one-element and two-element arrays and reads them as an argument, a scalar and a control
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   M_Ranges, CORE_Err, TEST_Harness; the Excel object model (Worksheets, Range)
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
    Dim Sheet             As Worksheet    'The first worksheet
    Dim One(1 To 1)       As Variant      'A one-element array holding a Range
    Dim Two(1 To 2)       As Variant      'A row whose second element is a Range
    Dim Payload           As Variant      'Payload read
    Dim Kind              As KPRArgShape  'Kind reported
    Dim RowCount          As Long         'Rows reported
    Dim ColCount          As Long         'Columns reported
    Dim Condition         As KPRCondition 'Condition reported
    Dim ScalarOut         As Variant      'Scalar read

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "range-elements"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Read Ranges of the first worksheet
        Set Sheet = ThisWorkbook.Worksheets(1)
    'A one-cell Range has a default member, which no path may evaluate
        Set One(1) = Sheet.Range("Z1000")
        Two(1) = 1
        Set Two(2) = Sheet.Range("Z1000")
        TEST_Harness.AssertTrue "row-read", _
            Not M_Ranges.TryReadArgument(Two, Payload, Kind, RowCount, ColCount, Condition), _
            "a Range element was read"
        TEST_Harness.AssertEqualText "row-read-condition", "SHAPE_UNSUPPORTED", CORE_Err.ConditionName(Condition)
        TEST_Harness.AssertTrue "one-read", _
            Not M_Ranges.TryReadArgument(One, Payload, Kind, RowCount, ColCount, Condition), _
            "a Range element was read as a scalar"
        TEST_Harness.AssertTrue "one-scalar", Not M_Ranges.TryReadScalar(One, ScalarOut), "a Range element was unwrapped"
        ProbeControl One, cdShapeUnsupported, "UNTOUCHED"
        TEST_Harness.AssertTrue "one-control-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "one-control-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "one-control-value", mFact(3, 1), mFact(3, 2)
        Set One(1) = Nothing
        Set Two(2) = Nothing

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

Private Sub ProbeShape( _
    ByVal ValueIn As Variant, _
    ByVal ExpectedCondition As KPRCondition, _
    ByVal ExpectedKind As KPRArgShape, _
    ByVal ExpectedRows As Long, _
    ByVal ExpectedCols As Long)

'
'------------------------------------------------------------------------------
'                        PROBE THE SHAPE OF AN ARGUMENT
'------------------------------------------------------------------------------
' PURPOSE
'   Describes an argument and checks the condition and, on success, its kind and size. 4
'   assertions
'
' WHY THIS EXISTS
'   The same four facts are checked for every argument shape
'
' INPUTS
'   Label
'     Prefix of the assertion names
'
'   ValueIn
'     The argument
'
'   ExpectedCondition
'     cdNone for a shape, otherwise the rejecting condition
'
'   ExpectedKind, ExpectedRows, ExpectedCols
'     The shape; checked only when no condition is expected
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Describes the argument, then asserts the result, the condition and the shape
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   M_Ranges.TryDescribeArgument, CORE_Err.ConditionName, TEST_Harness
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
    Dim Kind            As KPRArgShape  'Kind reported
    Dim RowCount        As Long         'Rows reported
    Dim ColCount        As Long         'Columns reported
    Dim Condition       As KPRCondition 'Condition reported
    Dim Accepted        As Boolean      'Whether the shape was accepted

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Start from values that no accepted shape has
        Kind = -1
        RowCount = -1
        ColCount = -1
        Accepted = M_Ranges.TryDescribeArgument(ValueIn, Kind, RowCount, ColCount, Condition)

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        mFact(1, 1) = (Accepted = (ExpectedCondition = cdNone))
        mFact(1, 2) = "the acceptance differs"
        mFact(2, 1) = CORE_Err.ConditionName(ExpectedCondition)
        mFact(2, 2) = CORE_Err.ConditionName(Condition)
    'A rejection leaves the shape unassigned
        If ExpectedCondition <> cdNone Then
            ExpectedKind = -1
            ExpectedRows = -1
            ExpectedCols = -1
        End If
        mFact(3, 1) = ExpectedKind
        mFact(3, 2) = Kind
        mFact(4, 1) = (RowCount = ExpectedRows) And (ColCount = ExpectedCols)
        mFact(4, 2) = "size " & RowCount & " by " & ColCount & ", expected " & ExpectedRows & " by " & ExpectedCols

End Sub

Private Sub ProbeControl( _
    ByVal ValueIn As Variant, _
    ByVal ExpectedCondition As KPRCondition, _
    ByVal ExpectedText As String)

'
'------------------------------------------------------------------------------
'                          PROBE AN OPTIONAL CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Unwraps a control and checks the condition and, on success, its value as text. 3
'   assertions
'
' WHY THIS EXISTS
'   The same three facts are checked for every control shape
'
' INPUTS
'   Label
'     Prefix of the assertion names
'
'   ValueIn
'     The control
'
'   ExpectedCondition
'     cdNone for a control that unwraps, otherwise the rejecting condition
'
'   ExpectedText
'     The unwrapped value as text, or UNTOUCHED when a condition is expected
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Unwraps the control from a sentinel, then asserts the result, the condition and the value
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   M_Ranges.TryReadControl, CORE_Err.ConditionName, TEST_Harness
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
    Dim ScalarOut       As Variant      'Unwrapped control, a sentinel on entry
    Dim Condition       As KPRCondition 'Condition reported
    Dim Accepted        As Boolean      'Whether the control unwrapped

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        ScalarOut = "UNTOUCHED"
        Accepted = M_Ranges.TryReadControl(ValueIn, ScalarOut, Condition)

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        mFact(1, 1) = (Accepted = (ExpectedCondition = cdNone))
        mFact(1, 2) = "the acceptance differs"
        mFact(2, 1) = CORE_Err.ConditionName(ExpectedCondition)
        mFact(2, 2) = CORE_Err.ConditionName(Condition)
        mFact(3, 1) = ExpectedText
        mFact(3, 2) = CStr(ScalarOut)

End Sub
