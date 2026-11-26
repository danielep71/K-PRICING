Attribute VB_Name = "TEST_CoreArray"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: TEST_CoreArray
'------------------------------------------------------------------------------
' PURPOSE
'   Tests of CORE_Array on plain values: rank, shape classification, the capacity gate, shape
'   resolution, materialization, element access, output allocation and the unwrapping of
'   scalars and controls
'
' WHY THIS EXISTS
'   The array rules of the calculation contracts hold for every function. These cases check
'   the core on every plain shape, without Excel, against the rules as the contract states
'   them
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   RunCoreArrayTests runs the suite. Option Private Module keeps it out of the workbook's
'   external surface. The cases build their inputs in memory and change no Excel state.
'   CertifyCoreArray certifies its normal run, as .github/excel-evidence-policy.json lists it
'
' ERROR POLICY
'   As TEST_Harness.EndRun: a run that does not pass raises TEST_Harness.ERR_TEST_FAILED,
'   or the unexpected error that escaped the runner. A refused start raises
'   TEST_Harness.ERR_TEST_DIRTY_START
'
' DEPENDENCIES
'   CORE_Array, CORE_Err, TEST_Harness
'
' NOTES
'   In a development workbook holding the modules of src/ and tests/modules/, type
'   TEST_CoreArray.RunCoreArrayTests in the Immediate window. Range arguments are tested in
'   TEST_Ranges
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'A complete run starts every case and evaluates every assertion once
    Private Const SUITE_NAME            As String = "core-array"        'Suite name in the run log
    Private Const EXPECTED_CASES        As Long = 8                 'Cases a complete run starts
    Private Const EXPECTED_ASSERTIONS   As Long = 122               'Assertions a complete run evaluates


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

Public Sub RunCoreArrayTests()

'
'------------------------------------------------------------------------------
'                             RUN CORE ARRAY TESTS
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case in a fixed order between BeginRun and EndRun. A passing run ends with
'   RESULT=PASS; suite=core-array; operation=PASS; completeness=COMPLETE; cases=8;
'   assertions=122; failures=0; cleanup=PASS; recovery=NONE
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
        CaseArrayRank
        CaseClassifyShape
        CaseCapacity
        CaseAccumulateShape
        CaseMaterialize
        CaseTraversal
        CaseUnwrapScalar
        CaseUnwrapControl

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

Public Sub CertifyCoreArray( _
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
'   Runs RunCoreArrayTests, saves its log, certifies the run, then saves the certification log.
'
' ERROR POLICY
'   TEST_Harness.EndCertification raises when the certification did not pass.
'
' DEPENDENCIES
'   RunCoreArrayTests, TEST_Harness
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
        RunCoreArrayTests
        RunErrorNumber = Err.Number
        RunErrorDescription = Err.Description
    'Restore normal error handling
        On Error GoTo 0
    'Save the log and certify the run
        TEST_Harness.SaveLog Folder & "\core-array.log"
        TEST_Harness.CertifyRun "NORMAL", True, RunErrorNumber, RunErrorDescription
    'Save the certification log, raising when the certification did not pass
        TEST_Harness.EndCertification Folder & "\core-array-certification.log"

End Sub

'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseArrayRank()

'
'------------------------------------------------------------------------------
'                             CHECK THE ARRAY RANK
'------------------------------------------------------------------------------
' PURPOSE
'   The rank is the number of dimensions an array exposes. 6 assertions.
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
'   Probes a scalar, an unallocated array, an empty array and arrays of rank 1, 2 and 3
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Array, CORE_Err, TEST_Harness
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
    Dim Unallocated()     As Variant  'A dynamic array never sized
    Dim Cube()            As Variant  'A rank-3 array
    Dim Table()           As Variant  'A rank-2 array

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "array-rank"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Not an array, or never sized, is rank 0
        TEST_Harness.AssertEqualDouble "scalar", 0, CORE_Array.ArrayRank(5)
        TEST_Harness.AssertEqualDouble "unallocated", 0, CORE_Array.ArrayRank(Unallocated)
    'An empty Array() is rank 1
        TEST_Harness.AssertEqualDouble "empty", 1, CORE_Array.ArrayRank(Array())
        TEST_Harness.AssertEqualDouble "one-dimensional", 1, CORE_Array.ArrayRank(Array(1, 2, 3))
        ReDim Table(1 To 2, 1 To 3)
        TEST_Harness.AssertEqualDouble "two-dimensional", 2, CORE_Array.ArrayRank(Table)
        ReDim Cube(1 To 2, 1 To 2, 1 To 2)
        TEST_Harness.AssertEqualDouble "three-dimensional", 3, CORE_Array.ArrayRank(Cube)

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

Private Sub CaseClassifyShape()

'
'------------------------------------------------------------------------------
'                        CHECK THE SHAPE CLASSIFICATION
'------------------------------------------------------------------------------
' PURPOSE
'   Each plain argument is a scalar or an array of known size, or an unsupported shape. 40 assertions.
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
'   Classifies scalars, one-element arrays, rows, blocks with any lower bounds, and each unsupported shape
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Array, CORE_Err, TEST_Harness
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
    Dim Table()           As Variant  'A 2 by 3 array
    Dim Offset()          As Variant  'A 3 by 2 array with lower bounds 5 and -1
    Dim One()             As Variant  'A 1 by 1 array
    Dim Cube()            As Variant  'A rank-3 array
    Dim Item              As Collection   'An object

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "classify-shape"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Scalars and one-element arrays are scalars
        ProbeShape 7, cdNone, skScalar, 1, 1
        TEST_Harness.AssertTrue "scalar-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "scalar-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "scalar-kind", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertTrue "scalar-size", mFact(4, 1), mFact(4, 2)
        ProbeShape Empty, cdNone, skScalar, 1, 1
        TEST_Harness.AssertTrue "empty-value-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "empty-value-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "empty-value-kind", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertTrue "empty-value-size", mFact(4, 1), mFact(4, 2)
        ProbeShape Array(7), cdNone, skScalar, 1, 1
        TEST_Harness.AssertTrue "one-element-row-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "one-element-row-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "one-element-row-kind", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertTrue "one-element-row-size", mFact(4, 1), mFact(4, 2)
        ReDim One(1 To 1, 1 To 1)
        ProbeShape One, cdNone, skScalar, 1, 1
        TEST_Harness.AssertTrue "one-by-one-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "one-by-one-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "one-by-one-kind", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertTrue "one-by-one-size", mFact(4, 1), mFact(4, 2)
    'A one-dimensional array is a row; a two-dimensional one keeps its size
        ProbeShape Array(1, 2, 3), cdNone, skArray, 1, 3
        TEST_Harness.AssertTrue "row-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "row-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "row-kind", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertTrue "row-size", mFact(4, 1), mFact(4, 2)
        ReDim Table(1 To 2, 1 To 3)
        ProbeShape Table, cdNone, skArray, 2, 3
        TEST_Harness.AssertTrue "block-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "block-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "block-kind", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertTrue "block-size", mFact(4, 1), mFact(4, 2)
        ReDim Offset(5 To 7, -1 To 0)
        ProbeShape Offset, cdNone, skArray, 3, 2
        TEST_Harness.AssertTrue "block-with-offsets-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "block-with-offsets-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "block-with-offsets-kind", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertTrue "block-with-offsets-size", mFact(4, 1), mFact(4, 2)
    'An empty array, a third dimension and an object are not shapes
        ProbeShape Array(), cdShapeUnsupported, 0, 0, 0
        TEST_Harness.AssertTrue "empty-array-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "empty-array-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "empty-array-kind", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertTrue "empty-array-size", mFact(4, 1), mFact(4, 2)
        ReDim Cube(1 To 2, 1 To 2, 1 To 2)
        ProbeShape Cube, cdShapeUnsupported, 0, 0, 0
        TEST_Harness.AssertTrue "three-dimensional-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "three-dimensional-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "three-dimensional-kind", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertTrue "three-dimensional-size", mFact(4, 1), mFact(4, 2)
        Set Item = New Collection
        ProbeShape Item, cdShapeUnsupported, 0, 0, 0
        TEST_Harness.AssertTrue "object-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "object-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualDouble "object-kind", mFact(3, 1), mFact(3, 2)
        TEST_Harness.AssertTrue "object-size", mFact(4, 1), mFact(4, 2)
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

Private Sub CaseCapacity()

'
'------------------------------------------------------------------------------
'                           CHECK THE CAPACITY GATE
'------------------------------------------------------------------------------
' PURPOSE
'   A result of 100,000 elements passes and one of 100,001 does not, at any size. 7 assertions.
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
'   Gates rows and columns on both sides of the cap, and a whole sheet, without overflow
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Array, CORE_Err, TEST_Harness
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
    Dim Condition         As KPRCondition 'Condition reported

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "capacity"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'The cap is 100,000 elements, both orientations
        TEST_Harness.AssertTrue "column-at-cap", CORE_Array.CheckCapacity(100000, 1, Condition), "100,000 by 1 was refused"
        TEST_Harness.AssertTrue "row-at-cap", CORE_Array.CheckCapacity(1, 100000, Condition), "1 by 100,000 was refused"
        TEST_Harness.AssertTrue "column-over-cap", Not CORE_Array.CheckCapacity(100001, 1, Condition), "100,001 by 1 passed"
        TEST_Harness.AssertEqualText "column-over-cap-condition", "CAPACITY_EXCEEDED", CORE_Err.ConditionName(Condition)
        TEST_Harness.AssertTrue "block-over-cap", Not CORE_Array.CheckCapacity(316, 317, Condition), "316 by 317 passed"
    'A whole sheet is counted in Double, so it is refused, not an overflow
        TEST_Harness.AssertTrue "whole-sheet", Not CORE_Array.CheckCapacity(1048576, 16384, Condition), "a whole sheet passed"
        TEST_Harness.AssertEqualText "whole-sheet-condition", "CAPACITY_EXCEEDED", CORE_Err.ConditionName(Condition)

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

Private Sub CaseAccumulateShape()

'
'------------------------------------------------------------------------------
'                          CHECK THE SHAPE RESOLUTION
'------------------------------------------------------------------------------
' PURPOSE
'   A scalar repeats over an array, equal arrays agree, and any other pair mismatches. 7 assertions.
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
'   Folds scalars and arrays of equal and of different sizes, a row against a column included
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Array, CORE_Err, TEST_Harness
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
    Dim OutKind           As KPRArgShape  'Resolved kind
    Dim OutRows           As Long         'Resolved rows
    Dim OutCols           As Long         'Resolved columns
    Dim Condition         As KPRCondition 'Condition reported

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "accumulate-shape"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'A scalar leaves the result a scalar, and an array sets it
        OutKind = skScalar
        OutRows = 1
        OutCols = 1
        TEST_Harness.AssertTrue "scalar-scalar", _
            CORE_Array.AccumulateShape(OutKind, OutRows, OutCols, skScalar, 1, 1, Condition) And _
            (OutKind = skScalar), _
            "two scalars did not stay a scalar"
        TEST_Harness.AssertTrue "scalar-array", _
            CORE_Array.AccumulateShape(OutKind, OutRows, OutCols, skArray, 2, 3, Condition) And _
            (OutRows = 2) And _
            (OutCols = 3), _
            "an array did not set the result"
        TEST_Harness.AssertTrue "array-scalar", _
            CORE_Array.AccumulateShape(OutKind, OutRows, OutCols, skScalar, 1, 1, Condition) And _
            (OutRows = 2) And _
            (OutCols = 3), _
            "a scalar changed the result"
        TEST_Harness.AssertTrue "equal-arrays", _
            CORE_Array.AccumulateShape(OutKind, OutRows, OutCols, skArray, 2, 3, Condition), _
            "equal arrays did not agree"
    'Different sizes mismatch: no outer product of a row and a column
        TEST_Harness.AssertTrue "transposed", _
            Not CORE_Array.AccumulateShape(OutKind, OutRows, OutCols, skArray, 3, 2, Condition), _
            "3 by 2 agreed with 2 by 3"
        TEST_Harness.AssertEqualText "transposed-condition", "SHAPE_MISMATCH", CORE_Err.ConditionName(Condition)
        OutKind = skArray
        OutRows = 1
        OutCols = 3
        TEST_Harness.AssertTrue "row-and-column", _
            Not CORE_Array.AccumulateShape(OutKind, OutRows, OutCols, skArray, 3, 1, Condition), _
            "a column agreed with a row"

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

Private Sub CaseMaterialize()

'
'------------------------------------------------------------------------------
'                          CHECK THE MATERIALIZATION
'------------------------------------------------------------------------------
' PURPOSE
'   Arrays become 1-based two-dimensional copies, and jagged or oversized ones are rejected. 17 assertions.
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
'   Materializes a row, a block with offsets and a one-element array, then a jagged array, an oversized array and the unsupported shapes
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Array, CORE_Err, TEST_Harness
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
    Dim Payload           As Variant      'Materialized payload
    Dim Kind              As KPRArgShape  'Kind reported
    Dim RowCount          As Long         'Rows reported
    Dim ColCount          As Long         'Columns reported
    Dim Condition         As KPRCondition 'Condition reported
    Dim Offset()          As Variant      'A 2 by 2 array with lower bounds 0 and 3
    Dim Large()           As Variant      'A row of 100,001 elements
    Dim Item              As Collection   'An object
    Dim Holder(1 To 2)    As Variant      'A row whose second element is an object

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "materialize"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'A row becomes 1 by 3, 1-based, with its values in order
        TEST_Harness.AssertTrue "row", _
            CORE_Array.TryMaterialize(Array(10, 20, 30), Payload, Kind, RowCount, ColCount, Condition), _
            "a row was refused"
        TEST_Harness.AssertTrue "row-bounds", _
            (LBound(Payload, 1) = 1) And _
            (UBound(Payload, 1) = 1) And _
            (LBound(Payload, 2) = 1) And _
            (UBound(Payload, 2) = 3), _
            "the row is not 1 by 3 from 1"
        TEST_Harness.AssertEqualDouble "row-third", 30, Payload(1, 3)
    'A block with any lower bounds becomes 1-based, element for element
        ReDim Offset(0 To 1, 3 To 4)
        Offset(0, 3) = 1
        Offset(0, 4) = 2
        Offset(1, 3) = 3
        Offset(1, 4) = 4
        TEST_Harness.AssertTrue "block", _
            CORE_Array.TryMaterialize(Offset, Payload, Kind, RowCount, ColCount, Condition), _
            "a block was refused"
        TEST_Harness.AssertTrue "block-shape", (Kind = skArray) And (RowCount = 2) And (ColCount = 2), "the block is not 2 by 2"
        TEST_Harness.AssertEqualDouble "block-last", 4, Payload(2, 2)
        TEST_Harness.AssertEqualDouble "block-row-major", 2, Payload(1, 2)
    'A one-element array materializes as its scalar
        TEST_Harness.AssertTrue "one-element", _
            CORE_Array.TryMaterialize(Array(42), Payload, Kind, RowCount, ColCount, Condition) And _
            (Kind = skScalar), _
            "a one-element array is not a scalar"
        TEST_Harness.AssertEqualDouble "one-element-value", 42, Payload
    'An element that is itself an array is a jagged array
        TEST_Harness.AssertTrue "jagged", _
            Not CORE_Array.TryMaterialize(Array(Array(1), 2), Payload, Kind, RowCount, ColCount, Condition), _
            "a jagged array passed"
        TEST_Harness.AssertEqualText "jagged-condition", "SHAPE_UNSUPPORTED", CORE_Err.ConditionName(Condition)
    'More than 100,000 elements are refused before any element is copied
        ReDim Large(1 To 100001)
        TEST_Harness.AssertTrue "over-cap", _
            Not CORE_Array.TryMaterialize(Large, Payload, Kind, RowCount, ColCount, Condition), _
            "100,001 elements passed"
        TEST_Harness.AssertEqualText "over-cap-condition", "CAPACITY_EXCEEDED", CORE_Err.ConditionName(Condition)
    'An object is not a plain value
        Set Item = New Collection
        TEST_Harness.AssertTrue "object", _
            Not CORE_Array.TryMaterialize(Item, Payload, Kind, RowCount, ColCount, Condition), _
            "an object passed"
        TEST_Harness.AssertEqualText "object-condition", "SHAPE_UNSUPPORTED", CORE_Err.ConditionName(Condition)
    'An element that is an object is refused, never read through a default member
        Holder(1) = 1
        Set Holder(2) = Item
        TEST_Harness.AssertTrue "object-element", _
            Not CORE_Array.TryMaterialize(Holder, Payload, Kind, RowCount, ColCount, Condition), _
            "an object element passed"
        TEST_Harness.AssertEqualText "object-element-condition", "SHAPE_UNSUPPORTED", CORE_Err.ConditionName(Condition)
        Set Holder(2) = Nothing
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

Private Sub CaseTraversal()

'
'------------------------------------------------------------------------------
'                       CHECK ELEMENT ACCESS AND OUTPUT
'------------------------------------------------------------------------------
' PURPOSE
'   A scalar repeats at every position, and an output has the resolved size, 1-based. 7 assertions.
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
'   Reads elements of a scalar and of an array payload, and allocates outputs of a valid, an empty and an oversized size
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Array, CORE_Err, TEST_Harness
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
    Dim Block()           As Variant      'A 2 by 2 payload
    Dim Scalar            As Variant      'A scalar payload
    Dim Allocated         As Variant      'An allocated output
    Dim Condition         As KPRCondition 'Condition reported

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "traversal"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'A scalar payload is the same at every position
        Scalar = 9
        TEST_Harness.AssertEqualDouble "scalar-anywhere", 9, CORE_Array.ElementAt(Scalar, skScalar, 2, 3)
    'An array payload is read at its row and column
        ReDim Block(1 To 2, 1 To 2)
        Block(2, 1) = 21
        TEST_Harness.AssertEqualDouble "array-element", 21, CORE_Array.ElementAt(Block, skArray, 2, 1)
    'An output has the resolved size, from 1
        TEST_Harness.AssertTrue "allocate", CORE_Array.TryAllocateOutput(2, 3, Allocated, Condition), "a 2 by 3 output was refused"
        TEST_Harness.AssertTrue "allocate-bounds", _
            (LBound(Allocated, 1) = 1) And _
            (UBound(Allocated, 1) = 2) And _
            (LBound(Allocated, 2) = 1) And _
            (UBound(Allocated, 2) = 3), _
            "the output is not 2 by 3 from 1"
    'No output has no element, or more than 100,000
        TEST_Harness.AssertTrue "allocate-empty", Not CORE_Array.TryAllocateOutput(0, 1, Allocated, Condition), "an empty output was allocated"
        TEST_Harness.AssertTrue "allocate-over-cap", _
            Not CORE_Array.TryAllocateOutput(100001, 1, Allocated, Condition), _
            "100,001 elements were allocated"
        TEST_Harness.AssertEqualText "allocate-over-cap-condition", "CAPACITY_EXCEEDED", CORE_Err.ConditionName(Condition)

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

Private Sub CaseUnwrapScalar()

'
'------------------------------------------------------------------------------
'                         CHECK THE SCALAR UNWRAPPING
'------------------------------------------------------------------------------
' PURPOSE
'   One element unwraps to its value; more elements, an object or a wrapped object do not. 8 assertions.
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
'   Unwraps scalars and one-element arrays of rank 1 and 2, then each rejected form
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Array, CORE_Err, TEST_Harness
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
    Dim ScalarOut         As Variant      'Unwrapped value
    Dim One()             As Variant      'A 1 by 1 array
    Dim Item              As Collection   'An object
    Dim Wrapped(1 To 1)   As Variant      'An array holding an object

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "unwrap-scalar"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'A scalar and one-element arrays unwrap
        TEST_Harness.AssertTrue "scalar", CORE_Array.TryUnwrapScalar(5, ScalarOut) And (ScalarOut = 5), "a scalar did not pass through"
        TEST_Harness.AssertTrue "row-of-one", _
            CORE_Array.TryUnwrapScalar(Array(6), ScalarOut) And (ScalarOut = 6), _
            "a one-element row did not unwrap"
        ReDim One(2 To 2, 3 To 3)
        One(2, 3) = 7
        TEST_Harness.AssertTrue "block-of-one", _
            CORE_Array.TryUnwrapScalar(One, ScalarOut) And (ScalarOut = 7), _
            "a 1 by 1 block did not unwrap"
    'More elements, an object, a wrapped object and a third dimension do not
        TEST_Harness.AssertTrue "two-elements", Not CORE_Array.TryUnwrapScalar(Array(1, 2), ScalarOut), "two elements unwrapped"
        Set Item = New Collection
        TEST_Harness.AssertTrue "object", Not CORE_Array.TryUnwrapScalar(Item, ScalarOut), "an object unwrapped"
        Set Wrapped(1) = Item
        TEST_Harness.AssertTrue "wrapped-object", Not CORE_Array.TryUnwrapScalar(Wrapped, ScalarOut), "a wrapped object unwrapped"
        Set Wrapped(1) = Nothing
        Set Item = Nothing
    'An element that is itself an array is not a scalar
        TEST_Harness.AssertTrue "nested-row", _
            Not CORE_Array.TryUnwrapScalar(Array(Array(1, 2)), ScalarOut), _
            "a nested row unwrapped"
        One(2, 3) = Array(1, 2)
        TEST_Harness.AssertTrue "nested-block", Not CORE_Array.TryUnwrapScalar(One, ScalarOut), "a 1 by 1 block holding an array unwrapped"

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

Private Sub CaseUnwrapControl()

'
'------------------------------------------------------------------------------
'                         CHECK THE CONTROL UNWRAPPING
'------------------------------------------------------------------------------
' PURPOSE
'   A control is one element; a larger one is not a scalar, whatever its size. 30 assertions.
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
'   Unwraps Empty, a scalar and one-element arrays, then multi-element, oversized, empty, jagged and object controls
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   CORE_Array, CORE_Err, TEST_Harness
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
    Dim Large()           As Variant      'A row of 100,001 elements
    Dim Block()           As Variant      'A 2 by 2 array
    Dim Item              As Collection   'An object
    Dim Holder(1 To 1)    As Variant      'A one-element array holding an object

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "unwrap-control"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Empty, a scalar and one-element arrays unwrap
        ProbeControl Empty, cdNone, ""
        TEST_Harness.AssertTrue "empty-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "empty-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "empty-value", mFact(3, 1), mFact(3, 2)
        ProbeControl True, cdNone, CStr(True)
        TEST_Harness.AssertTrue "boolean-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "boolean-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "boolean-value", mFact(3, 1), mFact(3, 2)
        ProbeControl Array("FLOOR"), cdNone, "FLOOR"
        TEST_Harness.AssertTrue "row-of-one-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "row-of-one-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "row-of-one-value", mFact(3, 1), mFact(3, 2)
    'More than one element is not a scalar, and 100,001 is not a capacity failure
        ProbeControl Array(1, 2), cdControlNotScalar, "UNTOUCHED"
        TEST_Harness.AssertTrue "row-of-two-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "row-of-two-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "row-of-two-value", mFact(3, 1), mFact(3, 2)
        ReDim Block(1 To 2, 1 To 2)
        ProbeControl Block, cdControlNotScalar, "UNTOUCHED"
        TEST_Harness.AssertTrue "block-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "block-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "block-value", mFact(3, 1), mFact(3, 2)
        ReDim Large(1 To 100001)
        ProbeControl Large, cdControlNotScalar, "UNTOUCHED"
        TEST_Harness.AssertTrue "over-cap-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "over-cap-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "over-cap-value", mFact(3, 1), mFact(3, 2)
    'An empty array, a jagged element and an object are not shapes
        ProbeControl Array(), cdShapeUnsupported, "UNTOUCHED"
        TEST_Harness.AssertTrue "empty-array-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "empty-array-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "empty-array-value", mFact(3, 1), mFact(3, 2)
        ProbeControl Array(Array(1)), cdShapeUnsupported, "UNTOUCHED"
        TEST_Harness.AssertTrue "jagged-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "jagged-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "jagged-value", mFact(3, 1), mFact(3, 2)
        Set Item = New Collection
        ProbeControl Item, cdShapeUnsupported, "UNTOUCHED"
        TEST_Harness.AssertTrue "object-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "object-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "object-value", mFact(3, 1), mFact(3, 2)
    'A one-element array holding an object is not a control
        Set Holder(1) = Item
        ProbeControl Holder, cdShapeUnsupported, "UNTOUCHED"
        TEST_Harness.AssertTrue "object-element-accepted", mFact(1, 1), mFact(1, 2)
        TEST_Harness.AssertEqualText "object-element-condition", mFact(2, 1), mFact(2, 2)
        TEST_Harness.AssertEqualText "object-element-value", mFact(3, 1), mFact(3, 2)
        Set Holder(1) = Nothing
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
'   CORE_Array.TryClassifyShape, CORE_Err.ConditionName, TEST_Harness
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
        Accepted = CORE_Array.TryClassifyShape(ValueIn, Kind, RowCount, ColCount, Condition)

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
'   CORE_Array.TryUnwrapControl, CORE_Err.ConditionName, TEST_Harness
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
        Accepted = CORE_Array.TryUnwrapControl(ValueIn, ScalarOut, Condition)

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
