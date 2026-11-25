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
'   write nothing. The suite is not certified: it is not listed in
'   .github/excel-evidence-policy.json
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
        CheckShape "cell", Sheet.Range("B2"), cdNone, skScalar, 1, 1
        CheckShape "block", Sheet.Range("A1:C2"), cdNone, skArray, 2, 3
        CheckShape "row", Sheet.Range("A1:D1"), cdNone, skArray, 1, 4
        CheckShape "whole-column", Sheet.Range("A:A"), cdNone, skArray, Sheet.Rows.Count, 1
    'Two areas are not one rectangle
        CheckShape "two-areas", Sheet.Range("A1:A2,C1:C2"), cdShapeUnsupported, 0, 0, 0
    'A plain array is the core's, and another object is unsupported
        ReDim Table(1 To 2, 1 To 3)
        CheckShape "plain-array", Table, cdNone, skArray, 2, 3
        Set Item = New Collection
        CheckShape "other-object", Item, cdShapeUnsupported, 0, 0, 0
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
        CheckControl "cell", Sheet.Range("Z1000"), cdNone, ""
    'More cells are not a scalar, even over the cap
        CheckControl "block", Sheet.Range("A1:B2"), cdControlNotScalar, "UNTOUCHED"
        CheckControl "over-cap", Sheet.Range("A1:A100001"), cdControlNotScalar, "UNTOUCHED"
    'Two areas are not a shape, and a plain array is the core's
        CheckControl "two-areas", Sheet.Range("A1,C1"), cdShapeUnsupported, "UNTOUCHED"
        CheckControl "plain", Array("CEILING"), cdNone, "CEILING"

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
        CheckControl "one-control", One, cdShapeUnsupported, "UNTOUCHED"
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

Private Sub CheckShape( _
    ByVal Label As String, _
    ByVal ValueIn As Variant, _
    ByVal ExpectedCondition As KPRCondition, _
    ByVal ExpectedKind As KPRArgShape, _
    ByVal ExpectedRows As Long, _
    ByVal ExpectedCols As Long)

'
'------------------------------------------------------------------------------
'                        CHECK THE SHAPE OF AN ARGUMENT
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
        TEST_Harness.AssertTrue Label & "-accepted", Accepted = (ExpectedCondition = cdNone), "the acceptance differs"
        TEST_Harness.AssertEqualText Label & "-condition", _
            CORE_Err.ConditionName(ExpectedCondition), _
            CORE_Err.ConditionName(Condition)
    'A rejection leaves the shape unassigned
        If ExpectedCondition <> cdNone Then
            ExpectedKind = -1
            ExpectedRows = -1
            ExpectedCols = -1
        End If
        TEST_Harness.AssertEqualDouble Label & "-kind", ExpectedKind, Kind
        TEST_Harness.AssertTrue Label & "-size", _
            (RowCount = ExpectedRows) And (ColCount = ExpectedCols), _
            _
            "size " & RowCount & " by " & ColCount & ", expected " & ExpectedRows & " by " & ExpectedCols

End Sub

Private Sub CheckControl( _
    ByVal Label As String, _
    ByVal ValueIn As Variant, _
    ByVal ExpectedCondition As KPRCondition, _
    ByVal ExpectedText As String)

'
'------------------------------------------------------------------------------
'                          CHECK AN OPTIONAL CONTROL
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
        TEST_Harness.AssertTrue Label & "-accepted", Accepted = (ExpectedCondition = cdNone), "the acceptance differs"
        TEST_Harness.AssertEqualText Label & "-condition", _
            CORE_Err.ConditionName(ExpectedCondition), _
            CORE_Err.ConditionName(Condition)
        TEST_Harness.AssertEqualText Label & "-value", ExpectedText, CStr(ScalarOut)

End Sub
