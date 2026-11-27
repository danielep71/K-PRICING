Attribute VB_Name = "CORE_Array"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: CORE_Array
'------------------------------------------------------------------------------
' PURPOSE
'   The shape engine of the date layer, on plain values only: classification, scalar and
'   control unwrapping, the capacity gate, materialization, exact-shape broadcasting and
'   output allocation.
'
' WHY THIS EXISTS
'   Every worksheet function applies the same array rules. Keeping them in one module that
'   knows no Excel object means that they can run, and be tested, without Excel, and
'   that a Range reaches them only as the plain value M_Ranges reads from it.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - ArrayRank          the true rank of an array, 0 to 60
'     - TryUnwrapScalar    a one-element array to its scalar
'     - TryUnwrapControl   the same for an Opt_ control, with CONTROL_NOT_SCALAR
'     - TryClassifyShape   the shape of a value from its type and bounds alone
'     - CheckCapacity      rows times columns against MAX_ELEMENTS, in Double
'     - TryMaterialize     the 1-based payload of a value, after the capacity gate
'     - AccumulateShape    one argument's shape folded into the result's
'     - ElementAt          the element of a payload at a position
'     - TryAllocateOutput  a 1-based result of a validated shape
'
' ERROR POLICY
'   No member raises; every failure is a return value with a condition.
'
' DEPENDENCIES
'   CORE_Err, for the condition vocabulary. No Excel object model: an object of any
'   kind is an unsupported shape here, and M_Ranges handles a Range before the core sees
'   it. Calculation contracts: docs/methodology/CALCULATION_CONTRACTS.md, Arrays and
'   evaluation
'
' NOTES
'   Order of operations that a worksheet function keeps (calculation contracts, Arrays
'   and evaluation):
'     1. Read the caller and apply the host rule. A Range in a 1904 workbook is refused
'        before it is read: once its serials are read, nothing can tell they meant
'        other days.
'     2. Describe every value argument with M_Ranges.TryDescribeArgument, from type and
'        dimensions alone, so an unsupported shape anywhere is reported first.
'     3. Read the controls with M_Ranges.TryReadControl, which never reads a
'        multi-element control, and fold the shapes with AccumulateShape
'        (SHAPE_MISMATCH).
'     4. CheckCapacity once, on the resolved shape. Only now is any content read.
'     5. Read each value argument with M_Ranges.TryReadArgument, then
'        TryAllocateOutput.
'     6. Traverse by rows: every column of row 1, then row 2, with ElementAt.
'   A jagged array is the one shape rejected after the capacity gate, because only its
'   elements reveal it. The per-element loop lives in the facade: VBA has no clean way
'   to pass an element function to this module. Option Private Module keeps every member
'   out of the Excel function list; nothing here is supported API
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PUBLIC CONSTANTS
'------------------------------------------------------------------------------
    'The largest result a call may have, in elements (calculation contracts, Arrays and
    'evaluation, stage 4)
    Public Const MAX_ELEMENTS   As Long = 100000    'Capacity of a resolved result

'------------------------------------------------------------------------------
' PUBLIC ENUMS
'------------------------------------------------------------------------------
    'The shape of an argument as the engine sees it
    Public Enum KPRArgShape
        skScalar = 0    'A scalar, a single cell or a 1 by 1 array
        skArray = 1     'RowCount by columns, with more than one element
    End Enum


'
'------------------------------------------------------------------------------
'
'                             SHAPE CLASSIFICATION
'
'------------------------------------------------------------------------------
'

Public Function ArrayRank(ByVal ArrayIn As Variant) As Long

'
'------------------------------------------------------------------------------
'                       REPORT THE TRUE RANK OF AN ARRAY
'------------------------------------------------------------------------------
' PURPOSE
'   Reports how many dimensions an array payload actually exposes.
'
' WHY THIS EXISTS
'   Callers must reject any shape they do not support, and VBA offers no
'   non-raising way to ask. Keeping the probe here means no caller has to
'   hold its own On Error Resume Next block.
'
'   Reporting true rank rather than a capped classification lets a caller
'   reject rank 3 and above at its own boundary. Folding 3D into 2D would
'   let such an array through to be indexed with two subscripts, raising
'   error 9 deep inside the caller instead.
'
' INPUTS
'   ArrayIn
'     Variant that may or may not hold an array payload. Never written to.
'
' RETURNS
'   Long
'     0  => non-array, or dynamic array that has never been ReDim'd
'     1  => 1D array, including an empty one returned by Array()
'     2  => 2D array
'     3+ => rank as declared, up to the VBA ceiling of 60
'
' BEHAVIOR
'   Returns 0 for a non-array. Otherwise probes LBound on dimensions 1 to 60
'   and returns the last dimension that probed clean.
'
' ERROR POLICY
'   - Does NOT raise.
'   - Clears Err as a side effect; callers must not hold pending error
'     state across this call.
'   - Restores normal error handling for the remainder of THIS procedure
'     only. VBA scopes error state per procedure, so the caller's handler
'     is unaffected either way.
'
' DEPENDENCIES
'   VBA library only
'
' NOTES
'   - Any result >= 1 guarantees LBound / UBound are safe for dimensions
'     1 to that value. No caller-side guard is required.
'   - Rank 0 covers both "not an array" and "array with no allocated
'     dimensions". Callers needing to separate these should use IsArray,
'     which does not raise.
'   - A jagged array is rank 1. Its elements being arrays is not visible
'     here.
'   - The probe is bounded at the VBA ceiling of 60 dimensions so that a
'     payload whose probe never fails cannot spin.
'   - ProbeBound is assigned and never read. It exists because LBound must
'     have a destination; only the success or failure of the call matters.
'     Removing it breaks the probe.
'
' UPDATED
'   2026-09-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Array.ArrayRank"

    Dim ProbeDim        As Long                                     'Dimension currently under probe
    Dim ProbeBound      As Long                                     'Probe target; only success / failure matters

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default is 0 unless at least one dimension probes clean
        ArrayRank = 0

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Non-array input has no dimensions to count
        If Not IsArray(ArrayIn) Then Exit Function

'------------------------------------------------------------------------------
' DIMENSION PROBE
'------------------------------------------------------------------------------
    'Probe upward until a dimension is unavailable; VBA caps rank at 60
        On Error Resume Next
        For ProbeDim = 1 To 60
            Err.Clear
            ProbeBound = LBound(ArrayIn, ProbeDim)
            If Err.Number <> 0 Then Exit For
            ArrayRank = ProbeDim
        Next ProbeDim
    'Restore normal error handling
        Err.Clear
        On Error GoTo 0

End Function

Public Function TryUnwrapScalar(ByVal ValueIn As Variant, ByRef ScalarOut As Variant) As Boolean

'
'------------------------------------------------------------------------------
'                               UNWRAP A SCALAR
'------------------------------------------------------------------------------
' PURPOSE
'   Reduces a plain value argument to its single scalar: a one-element array to its
'   element, and anything else that is not an array to itself.
'
' WHY THIS EXISTS
'   A worksheet function may receive its scalar wrapped in a one-element array, and every
'   scalar path needs the same unwrapping.
'
' INPUTS
'   ValueIn
'     A plain value: a scalar or a VBA array, never a Range, which M_Ranges reads first.
'
'   ScalarOut (ByRef, output)
'     Assigned only on success.
'
' RETURNS
'   Boolean
'     True when ScalarOut holds the single scalar. False for an object, a rank other than
'     1 or 2, more than one element, or an element that is itself an object or an array.
'
' BEHAVIOR
'   Tests the argument itself before any copy, then unwraps a one-element array of rank
'   1 or 2, or passes a scalar through.
'
' ERROR POLICY
'   Does not raise. A runtime error while reading is a False result, and Err is cleared.
'
' DEPENDENCIES
'   ArrayRank. Calculation contracts: Arrays and evaluation
'
' NOTES
'   The argument is never copied with Let before its type is tested: a Let copy of a
'   Variant that holds an object calls the object's default member
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Array.TryUnwrapScalar"

    Dim Payload         As Variant                                      'Isolated scalar, assigned out only on success
    Dim ArrRank         As Long                                         'Array rank as reported by ArrayRank

    Dim LB1             As Long                                         'Lower bound, dimension 1
    Dim UB1             As Long                                         'Upper bound, dimension 1
    Dim LB2             As Long                                         'Lower bound, dimension 2
    Dim UB2             As Long                                         'Upper bound, dimension 2

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Trap runtime errors from property access and convert them to a FALSE return
        On Error GoTo Fail
    'Default return is failure unless we explicitly succeed
        TryUnwrapScalar = False

'------------------------------------------------------------------------------
' ISOLATE PAYLOAD
'------------------------------------------------------------------------------
    'Test the argument itself; a Let copy would dereference the default member
        If IsObject(ValueIn) Then
            'An object is never a plain value: M_Ranges reads a Range first
            GoTo Fail

        'If an array arrived, accept only a 1x1 wrapper (scalar contract)
        ElseIf IsArray(ValueIn) Then
            'Classify the array shape once
            ArrRank = ArrayRank(ValueIn)
            'Reject unsupported shapes before any bound is read
            If ArrRank < 1 Or ArrRank > 2 Then GoTo Fail
            'Two-dimensional wrapper
            If ArrRank = 2 Then
                'Capture bounds
                LB1 = LBound(ValueIn, 1)
                UB1 = UBound(ValueIn, 1)
                LB2 = LBound(ValueIn, 2)
                UB2 = UBound(ValueIn, 2)
                'Reject anything larger than 1x1
                If (UB1 <> LB1) Or (UB2 <> LB2) Then GoTo Fail
                'A wrapped object or array is not a scalar payload
                If IsObject(ValueIn(LB1, LB2)) Or IsArray(ValueIn(LB1, LB2)) Then GoTo Fail
                'Unwrap the scalar payload
                Payload = ValueIn(LB1, LB2)
            'One-dimensional wrapper
            Else
                'Capture bounds
                LB1 = LBound(ValueIn)
                UB1 = UBound(ValueIn)
                'Reject vectors with more than one element
                If UB1 <> LB1 Then GoTo Fail
                'A wrapped object or array is not a scalar payload
                If IsObject(ValueIn(LB1)) Or IsArray(ValueIn(LB1)) Then GoTo Fail
                'Unwrap the scalar payload
                Payload = ValueIn(LB1)
            End If

        'Neither object nor array: the argument is already a scalar
        Else
            Payload = ValueIn
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once a single scalar has been isolated
        ScalarOut = Payload
    'Contract: TRUE only when ScalarOut was assigned
        TryUnwrapScalar = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the failure handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Fail:
    'Return FALSE per contract, leaving ScalarOut untouched
        TryUnwrapScalar = False
    'Do not leave stale error state for the caller to observe
        Err.Clear

End Function

Public Function CheckCapacity( _
    ByVal RowCount As Long, _
    ByVal ColCount As Long, _
    ByRef Condition As KPRCondition) As Boolean
'
'------------------------------------------------------------------------------
'                    GATE A SHAPE AGAINST THE CAPACITY CAP
'------------------------------------------------------------------------------
' PURPOSE
'   Gates a shape against the 100,000-element cap.
'
' WHY THIS EXISTS
'   The the calculation contracts cap (MAX_ELEMENTS) is applied in one place,
'   from the dimensions alone, before any content is read.
'
' INPUTS
'   RowCount
'     Row count of the shape.
'
'   ColCount
'     Column count of the shape.
'
'   Condition (ByRef)
'     Always assigned: cdNone or cdCapacityExceeded.
'
' RETURNS
'   Boolean
'     TRUE  => RowCount x ColCount is at most MAX_ELEMENTS; Condition is NONE
'     FALSE => CAPACITY_EXCEEDED
'
' BEHAVIOR
'   Compares RowCount x ColCount, taken in Double, with MAX_ELEMENTS and assigns
'   Condition and the result accordingly.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   MAX_ELEMENTS; CORE_Err for the condition vocabulary
'
' NOTES
'   - The product is taken in Double. A worksheet Range can have over a
'     million rows and sixteen thousand columns, and the Long product of a
'     large Range overflows before it can be compared.
'   - Called from the dimensions alone, before any content is read.
'
' UPDATED
'   2026-09-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Compare the element count in Double so a large Range cannot overflow
        If (VBA.CDbl(RowCount) * VBA.CDbl(ColCount)) > VBA.CDbl(MAX_ELEMENTS) Then
            Condition = cdCapacityExceeded
            CheckCapacity = False
        Else
            Condition = cdNone
            CheckCapacity = True
        End If

End Function

Public Function TryClassifyShape( _
    ByVal ValueIn As Variant, _
    ByRef Kind As KPRArgShape, _
    ByRef RowCount As Long, _
    ByRef ColCount As Long, _
    ByRef Condition As KPRCondition) As Boolean
'
'------------------------------------------------------------------------------
'                        CLASSIFY THE SHAPE OF A VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Classifies a plain value argument as a scalar or an array, with its rows and columns,
'   from its type and bounds alone.
'
' WHY THIS EXISTS
'   Stage 2 of a call classifies every argument before any content is read, so an
'   unsupported shape anywhere is reported first.
'
' INPUTS
'   ValueIn
'     A plain value: a scalar or a VBA array. A Range is described by M_Ranges instead.
'
'   Kind, RowCount, ColCount (ByRef, output)
'     The shape; assigned on success.
'
'   Condition (ByRef, output)
'     Always assigned: cdNone or cdShapeUnsupported.
'
' RETURNS
'   Boolean
'     True with the shape. False with cdShapeUnsupported for an object, an empty array or
'     a rank other than 1 or 2.
'
' BEHAVIOR
'   A one-dimensional array is 1 by N; a two-dimensional array keeps its rows and
'   columns; a 1 by 1 array and a non-array are scalars. No element is read.
'
' ERROR POLICY
'   Does not raise. A runtime error is a shape rejection, and Err is cleared.
'
' DEPENDENCIES
'   ArrayRank, CORE_Err (KPRCondition). Calculation contracts: Arrays and evaluation,
'   stage 2
'
' NOTES
'   A jagged array is rank 1 here: its elements are arrays, which only reading them
'   reveals, so TryMaterialize reports it after the capacity gate
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Array.TryClassifyShape"

    Dim Rank            As Long                                         'Array rank
    Dim LB1             As Long                                         'Lower bound, dimension 1
    Dim UB1             As Long                                         'Upper bound, dimension 1
    Dim LB2             As Long                                         'Lower bound, dimension 2
    Dim UB2             As Long                                         'Upper bound, dimension 2

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Convert any runtime error during classification into a shape rejection
        On Error GoTo Unsupported
    'Default to a rejection until a shape is classified
        TryClassifyShape = False
        Condition = cdShapeUnsupported

'------------------------------------------------------------------------------
' OBJECT
'------------------------------------------------------------------------------
    'An object is never a plain value: M_Ranges reads a Range before the core sees it
        If IsObject(ValueIn) Then GoTo Unsupported

'------------------------------------------------------------------------------
' VBA ARRAY
'------------------------------------------------------------------------------
    'Classify a rank-1 or rank-2 array from its bounds, without reading elements
        If IsArray(ValueIn) Then
            Rank = ArrayRank(ValueIn)
            Select Case Rank
                Case 1
                    LB1 = LBound(ValueIn)
                    UB1 = UBound(ValueIn)
                    If UB1 < LB1 Then GoTo Unsupported
                    RowCount = 1
                    ColCount = UB1 - LB1 + 1
                Case 2
                    LB1 = LBound(ValueIn, 1)
                    UB1 = UBound(ValueIn, 1)
                    LB2 = LBound(ValueIn, 2)
                    UB2 = UBound(ValueIn, 2)
                    If (UB1 < LB1) Or (UB2 < LB2) Then GoTo Unsupported
                    RowCount = UB1 - LB1 + 1
                    ColCount = UB2 - LB2 + 1
                Case Else
                    GoTo Unsupported
            End Select
            If (RowCount = 1) And (ColCount = 1) Then
                Kind = skScalar
            Else
                Kind = skArray
            End If
            Condition = cdNone
            TryClassifyShape = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' SCALAR
'------------------------------------------------------------------------------
    'Anything else is a scalar
        Kind = skScalar
        RowCount = 1
        ColCount = 1
        Condition = cdNone
        TryClassifyShape = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the rejection handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Unsupported:
    'Any failure while classifying is a shape rejection, never a raise
        Err.Clear
        Condition = cdShapeUnsupported
        TryClassifyShape = False

End Function

'
'------------------------------------------------------------------------------
'
'                        MATERIALIZATION AND TRAVERSAL
'
'------------------------------------------------------------------------------
'

Public Function TryMaterialize( _
    ByVal ValueIn As Variant, _
    ByRef Payload As Variant, _
    ByRef Kind As KPRArgShape, _
    ByRef RowCount As Long, _
    ByRef ColCount As Long, _
    ByRef Condition As KPRCondition) As Boolean
'
'------------------------------------------------------------------------------
'                        COPY A VALUE INTO ITS PAYLOAD
'------------------------------------------------------------------------------
' PURPOSE
'   Copies a plain value argument into the payload the traversal reads: its scalar, or a
'   1-based two-dimensional copy of its array.
'
' WHY THIS EXISTS
'   Stage 5 reads the arguments only once the shapes and the capacity have passed, and
'   the traversal indexes every array the same way.
'
' INPUTS
'   ValueIn
'     A plain value: a scalar or a VBA array. A Range is read by M_Ranges instead.
'
'   Payload, Kind, RowCount, ColCount (ByRef, output)
'     The copy and its shape; assigned on success.
'
'   Condition (ByRef, output)
'     Always assigned: cdNone, cdCapacityExceeded or cdShapeUnsupported.
'
' RETURNS
'   Boolean
'     True with the payload. False with cdCapacityExceeded for more than MAX_ELEMENTS
'     elements, or cdShapeUnsupported for an object, an empty array, a rank other than 1
'     or 2, or an element that is itself an array or an object.
'
' BEHAVIOR
'   Gates the capacity from the bounds, then copies the elements row by row into a
'   1-based array, rejecting an element that is an array or an object as it meets it, so
'   that no default member is ever evaluated here.
'
' ERROR POLICY
'   Does not raise. A runtime error is a shape rejection, and Err is cleared.
'
' DEPENDENCIES
'   ArrayRank, CheckCapacity, CORE_Err (KPRCondition). Calculation contracts: Arrays and
'   evaluation, stages 4 and 5
'
' NOTES
'   The capacity gate here is a defence: the facade gates the resolved shape first, and
'   every array matched it
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Array.TryMaterialize"

    Dim Rank            As Long                                         'Array rank
    Dim LB1             As Long                                         'Lower bound, dimension 1
    Dim UB1             As Long                                         'Upper bound, dimension 1
    Dim LB2             As Long                                         'Lower bound, dimension 2
    Dim UB2             As Long                                         'Upper bound, dimension 2
    Dim R               As Long                                         'Row cursor
    Dim C               As Long                                         'Column cursor
    Dim Snapshot        As Variant                                      'Normalized 1-based copy

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Trap runtime errors from property access and convert them to a rejection
        On Error GoTo Unsupported
        TryMaterialize = False
        Condition = cdShapeUnsupported

'------------------------------------------------------------------------------
' OBJECT
'------------------------------------------------------------------------------
    'An object is never a plain value: M_Ranges reads a Range before the core sees it
        If IsObject(ValueIn) Then GoTo Unsupported

'------------------------------------------------------------------------------
' VBA ARRAY
'------------------------------------------------------------------------------
    'Normalize a rank-1 or rank-2 array to a 1-based copy after the gate
        If IsArray(ValueIn) Then

            Rank = ArrayRank(ValueIn)

            Select Case Rank

                Case 1
                    'One-dimensional is 1 x N
                    LB1 = LBound(ValueIn)
                    UB1 = UBound(ValueIn)
                    If UB1 < LB1 Then GoTo Unsupported
                    RowCount = 1
                    ColCount = UB1 - LB1 + 1
                    If Not CheckCapacity(RowCount, ColCount, Condition) Then Exit Function

                    'A single element is a scalar; an element that is an array is jagged
                    If ColCount = 1 Then
                        If IsArray(ValueIn(LB1)) Or IsObject(ValueIn(LB1)) Then GoTo Unsupported
                        Kind = skScalar
                        Payload = ValueIn(LB1)
                    Else
                        ReDim Snapshot(1 To 1, 1 To ColCount)
                        For C = 1 To ColCount
                            If IsArray(ValueIn(LB1 + C - 1)) Or IsObject(ValueIn(LB1 + C - 1)) Then GoTo Unsupported
                            Snapshot(1, C) = ValueIn(LB1 + C - 1)
                        Next C
                        Kind = skArray
                        Payload = Snapshot
                    End If

                Case 2
                    'Two-dimensional: rows then columns, bounds normalized to 1
                    LB1 = LBound(ValueIn, 1)
                    UB1 = UBound(ValueIn, 1)
                    LB2 = LBound(ValueIn, 2)
                    UB2 = UBound(ValueIn, 2)
                    If (UB1 < LB1) Or (UB2 < LB2) Then GoTo Unsupported
                    RowCount = UB1 - LB1 + 1
                    ColCount = UB2 - LB2 + 1
                    If Not CheckCapacity(RowCount, ColCount, Condition) Then Exit Function

                    'A single element is a scalar; any element that is an array is jagged
                    If (RowCount = 1) And (ColCount = 1) Then
                        If IsArray(ValueIn(LB1, LB2)) Or IsObject(ValueIn(LB1, LB2)) Then GoTo Unsupported
                        Kind = skScalar
                        Payload = ValueIn(LB1, LB2)
                    Else
                        ReDim Snapshot(1 To RowCount, 1 To ColCount)
                        For R = 1 To RowCount
                            For C = 1 To ColCount
                                If IsArray(ValueIn(LB1 + R - 1, LB2 + C - 1)) Or _
                                    IsObject(ValueIn(LB1 + R - 1, LB2 + C - 1)) Then GoTo Unsupported
                                Snapshot(R, C) = ValueIn(LB1 + R - 1, LB2 + C - 1)
                            Next C
                        Next R
                        Kind = skArray
                        Payload = Snapshot
                    End If

                Case Else
                    'Rank 0 cannot occur here; rank 3 and above is unsupported
                    GoTo Unsupported

            End Select

            Condition = cdNone
            TryMaterialize = True
            Exit Function

        End If

'------------------------------------------------------------------------------
' SCALAR
'------------------------------------------------------------------------------
    'Everything else is a scalar payload, including Empty, Null and errors
        Kind = skScalar
        RowCount = 1
        ColCount = 1
        Payload = ValueIn
        Condition = cdNone
        TryMaterialize = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the rejection handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Unsupported:
    'Any failure while classifying is a shape rejection, never a raise
        Err.Clear
        Condition = cdShapeUnsupported
        TryMaterialize = False

End Function

Public Function AccumulateShape( _
    ByRef OutKind As KPRArgShape, _
    ByRef OutRows As Long, _
    ByRef OutCols As Long, _
    ByVal ArgKind As KPRArgShape, _
    ByVal ArgRows As Long, _
    ByVal ArgCols As Long, _
    ByRef Condition As KPRCondition) As Boolean
'
'------------------------------------------------------------------------------
'                   FOLD ONE ARGUMENT INTO THE OUTPUT SHAPE
'------------------------------------------------------------------------------
' PURPOSE
'   Folds one argument's shape into the running output shape.
'
' WHY THIS EXISTS
'   The facade resolves one output shape across all classified arguments
'   (module header, step 3); this keeps the exact-match broadcast rule in one
'   place.
'
' INPUTS
'   OutKind / OutRows / OutCols (ByRef)
'     The running output shape, updated in place.
'
'   ArgKind / ArgRows / ArgCols
'     The shape of the argument being folded in.
'
'   Condition (ByRef)
'     Always assigned: cdNone or cdShapeMismatch.
'
' RETURNS
'   Boolean
'     TRUE  => the argument fits the output shape
'     FALSE => SHAPE_MISMATCH
'
' BEHAVIOR
'   - Start with OutKind = skScalar and OutRows = OutCols = 1.
'   - A scalar argument leaves the output shape unchanged: it expands.
'   - The first array argument sets the output shape.
'   - Every later array argument must match it exactly in rows and columns;
'     otherwise SHAPE_MISMATCH. There is no outer product and no
'     cross-broadcast: a 1 x N against an N x 1 is a mismatch.
'   - An all-scalar call leaves the output scalar, never a 1 x 1 array.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   CORE_Err for the condition vocabulary
'
' NOTES
'   None
'
' UPDATED
'   2026-09-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Assume the argument fits until a mismatch is found
        Condition = cdNone
        AccumulateShape = True

'------------------------------------------------------------------------------
' FOLD
'------------------------------------------------------------------------------
    'A scalar never constrains the shape
        If ArgKind = skScalar Then Exit Function

    'First array sets the shape
        If OutKind = skScalar Then
            OutKind = skArray
            OutRows = ArgRows
            OutCols = ArgCols
            Exit Function
        End If

    'Later arrays must match exactly
        If (ArgRows <> OutRows) Or (ArgCols <> OutCols) Then
            Condition = cdShapeMismatch
            AccumulateShape = False
        End If

End Function

Public Function ElementAt( _
    ByRef Payload As Variant, _
    ByVal Kind As KPRArgShape, _
    ByVal R As Long, _
    ByVal C As Long) As Variant
'
'------------------------------------------------------------------------------
'                        READ ONE ELEMENT OF A PAYLOAD
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the element at 1-based position (R, C) of a materialized payload.
'
' WHY THIS EXISTS
'   The facade reads scalar and array payloads through one call; returning a
'   scalar payload at every position is what scalar expansion means.
'
' INPUTS
'   Payload
'     A payload from a successful TryMaterialize. Not modified.
'
'   Kind
'     The Kind TryMaterialize reported for that payload.
'
'   R / C
'     1-based position inside the output shape AccumulateShape resolved.
'
' RETURNS
'   Variant
'     The element at (R, C), or the scalar payload itself.
'
' BEHAVIOR
'   Returns Payload for a scalar and Payload(R, C) for an array.
'
' ERROR POLICY
'   No condition channel and no handler; see NOTES for the preconditions.
'
' DEPENDENCIES
'   None
'
' NOTES
'   Total over a valid descriptor: Payload and Kind must come from a successful
'   TryMaterialize, and (R, C) must lie inside the output shape that
'   AccumulateShape resolved. Under those preconditions no failure exists, so
'   there is no condition channel. A scalar payload returns itself at every
'   position, which IS scalar expansion.
'
' UPDATED
'   2026-09-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'A scalar expands to every position; an array is indexed
        If Kind = skScalar Then
            ElementAt = Payload
        Else
            ElementAt = Payload(R, C)
        End If

End Function

Public Function TryAllocateOutput( _
    ByVal RowCount As Long, _
    ByVal ColCount As Long, _
    ByRef OutArr As Variant, _
    ByRef Condition As KPRCondition) As Boolean
'
'------------------------------------------------------------------------------
'                          ALLOCATE THE OUTPUT ARRAY
'------------------------------------------------------------------------------
' PURPOSE
'   Allocates a 1-based RowCount x ColCount Variant array for the output, after the
'   capacity gate.
'
' WHY THIS EXISTS
'   The facade allocates its output only once the shape has passed the
'   capacity gate (module header, step 5); this re-checks the shape it is
'   given before allocating.
'
' INPUTS
'   RowCount / ColCount
'     Dimensions of the resolved output shape.
'
'   OutArr (ByRef)
'     Assigned ONLY on success.
'
'   Condition (ByRef)
'     Always assigned.
'
' RETURNS
'   Boolean
'     TRUE  => OutArr allocated
'     FALSE => CAPACITY_EXCEEDED, or SHAPE_UNSUPPORTED for a non-positive
'              dimension
'
' BEHAVIOR
'   Rejects a non-positive dimension, gates capacity, then allocates and
'   assigns the array.
'
' ERROR POLICY
'   No handler; every expected failure returns FALSE with a condition.
'
' DEPENDENCIES
'   CheckCapacity
'
' NOTES
'   None
'
' UPDATED
'   2026-09-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Array.TryAllocateOutput"

    Dim Fresh           As Variant                                      'Newly allocated output

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default return is failure
        TryAllocateOutput = False

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a non-positive dimension
        If (RowCount < 1) Or (ColCount < 1) Then
            Condition = cdShapeUnsupported
            Exit Function
        End If
    'Gate the capacity
        If Not CheckCapacity(RowCount, ColCount, Condition) Then Exit Function

'------------------------------------------------------------------------------
' ALLOCATE
'------------------------------------------------------------------------------
    'Allocate the 1-based output and report success
        ReDim Fresh(1 To RowCount, 1 To ColCount)
        OutArr = Fresh
        Condition = cdNone
        TryAllocateOutput = True

End Function

'
'------------------------------------------------------------------------------
'
'                              OPTIONAL CONTROLS
'
'------------------------------------------------------------------------------
'

Public Function TryUnwrapControl( _
    ByVal ValueIn As Variant, _
    ByRef ScalarOut As Variant, _
    ByRef Condition As KPRCondition) As Boolean
'
'------------------------------------------------------------------------------
'                          UNWRAP AN OPTIONAL CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Reduces a plain control argument to its single scalar, reporting a multi-element
'   control and an unsupported shape separately.
'
' WHY THIS EXISTS
'   A control with more than one element is CONTROL_NOT_SCALAR, not a capacity or shape
'   failure, so controls need their own unwrapping.
'
' INPUTS
'   ValueIn
'     A plain control: a scalar or a VBA array. A Range is read by M_Ranges instead.
'
'   ScalarOut (ByRef, output)
'     Assigned only on success.
'
'   Condition (ByRef, output)
'     Always assigned: cdNone, cdControlNotScalar or cdShapeUnsupported.
'
' RETURNS
'   Boolean
'     True with the scalar. False with cdControlNotScalar for more than one element, or
'     cdShapeUnsupported for an object, an empty array, a rank other than 1 or 2, or an
'     element that is itself an array or an object.
'
' BEHAVIOR
'   Counts the elements from the bounds alone and never reads a multi-element control.
'   It is not capacity-gated: a control of 100,001 elements is CONTROL_NOT_SCALAR.
'
' ERROR POLICY
'   Does not raise. A runtime error is a shape rejection, and Err is cleared.
'
' DEPENDENCIES
'   ArrayRank, CORE_Err (KPRCondition). Data contract: Function arguments, Control values
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
    Const PROC_NAME     As String = "CORE_Array.TryUnwrapControl"

    Dim Rank            As Long                                         'Array rank
    Dim Count           As Double                                       'Element count, in Double for large Ranges

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Convert any runtime error during classification into a shape rejection
        On Error GoTo Unsupported
    'Default to a rejection until the control is unwrapped
        TryUnwrapControl = False
        Condition = cdShapeUnsupported

'------------------------------------------------------------------------------
' OBJECT
'------------------------------------------------------------------------------
    'An object is never a plain value: M_Ranges reads a Range before the core sees it
        If IsObject(ValueIn) Then GoTo Unsupported

'------------------------------------------------------------------------------
' VBA ARRAY
'------------------------------------------------------------------------------
    'Unwrap a one-element array; reject a larger one from its bounds alone
        If IsArray(ValueIn) Then
            Rank = ArrayRank(ValueIn)
            Select Case Rank
                Case 1
                    If UBound(ValueIn) < LBound(ValueIn) Then GoTo Unsupported
                    Count = VBA.CDbl(UBound(ValueIn) - LBound(ValueIn) + 1)
                    If Count <> 1# Then
                        Condition = cdControlNotScalar
                        Exit Function
                    End If
                    If IsArray(ValueIn(LBound(ValueIn))) Or IsObject(ValueIn(LBound(ValueIn))) Then GoTo Unsupported
                    ScalarOut = ValueIn(LBound(ValueIn))
                Case 2
                    If (UBound(ValueIn, 1) < LBound(ValueIn, 1)) Or _
                        (UBound(ValueIn, 2) < LBound(ValueIn, 2)) Then GoTo Unsupported
                    Count = VBA.CDbl(UBound(ValueIn, 1) - LBound(ValueIn, 1) + 1) * _
                        VBA.CDbl(UBound(ValueIn, 2) - LBound(ValueIn, 2) + 1)
                    If Count <> 1# Then
                        Condition = cdControlNotScalar
                        Exit Function
                    End If
                    If IsArray(ValueIn(LBound(ValueIn, 1), LBound(ValueIn, 2))) Or _
                        IsObject(ValueIn(LBound(ValueIn, 1), LBound(ValueIn, 2))) Then GoTo Unsupported
                    ScalarOut = ValueIn(LBound(ValueIn, 1), LBound(ValueIn, 2))
                Case Else
                    GoTo Unsupported
            End Select
            Condition = cdNone
            TryUnwrapControl = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' SCALAR
'------------------------------------------------------------------------------
    'Anything else is already a scalar control value
        ScalarOut = ValueIn
        Condition = cdNone
        TryUnwrapControl = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the rejection handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Unsupported:
    'Any failure while classifying is a shape rejection, never a raise
        Err.Clear
        Condition = cdShapeUnsupported
        TryUnwrapControl = False

End Function
