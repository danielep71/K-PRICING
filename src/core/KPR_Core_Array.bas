Attribute VB_Name = "KPR_Core_Array"
'==============================================================================
' MODULE: KPR_Core_Array
'------------------------------------------------------------------------------
' PURPOSE
'   Internal boundary for input shape: classification, scalar/control
'   unwrapping, capacity gating, materialization, exact-shape broadcasting and
'   output allocation.
'
' WHY THIS EXISTS
'   The pre-split baseline mixed three concerns inside Parse_Date: unwrapping
'   shape, converting a scalar to a Date, and gating the supported window.
'   Shape lives here, value conversion lives in KPR_Core_Parse, and the window
'   belongs to the date domain in KPR_Core_Dates. Their separation lets the
'   facade traverse shapes without duplicating value parsing.
'
' SCOPE
'   Shape services for the worksheet-facing date facade:
'
'   - Array_Rank          true array rank, 0 through 60
'   - TryUnwrapScalar     single-cell Range and 1x1 array unwrapping (scalar path)
'   - TryUnwrapControl    the same for Opt_ controls, reporting CONTROL_NOT_SCALAR
'                         for a valid multi-element shape without reading it
'   - TryClassifyShape    classify one argument from type and dimensions alone
'   - CheckCapacity       Rows x Cols <= 100,000, computed in Double
'   - TryMaterialize      classify, gate capacity, then snapshot one argument
'                         into a scalar or a 1-based 2-D payload
'   - AccumulateShape     fold one argument's shape into the output shape
'   - ElementAt           the element at a position; total over a valid descriptor
'   - TryAllocateOutput   a 1-based 2-D output of a validated shape
'
'   The engine owns shape classification, broadcast resolution, allocation and
'   unwrapping ONLY. It contains no date algorithm, never classifies the host,
'   and performs no function-pointer dispatch: VBA has no clean way to call an
'   element implementation from here, so the per-element loop lives in the
'   facade and calls these services.
'
' ORDER OF OPERATIONS THE FACADE MUST KEEP
'   1. PassHostGuard first. A Range in a 1904 workbook must be refused BEFORE
'      it is materialized: once Value2 has been read, its serials are in hand
'      and nothing downstream can tell they meant something else. The engine
'      never sees the guard by design, so this ordering is the facade's duty.
'   2. TryClassifyShape EVERY value argument, so an unsupported shape anywhere
'      in the argument list is reported before anything else.
'   3. Resolve Opt_ controls with TryUnwrapControl, which never reads the
'      contents of a multi-element control, and AccumulateShape across the
'      classified arguments (SHAPE_MISMATCH).
'   4. CheckCapacity once, on the resolved output shape. Only now may content
'      be read.
'   5. TryMaterialize each value argument (its internal gate is a defence that
'      cannot fire differently, since every array matched the gated shape),
'      then TryAllocateOutput.
'   6. Traverse row-major: R = 1 To Rows, C = 1 To Cols. Traversal is
'      deterministic and touches no Excel state.
'
'   This is contract section 5.2's order. A jagged array is the one shape
'   rejection reported after the capacity gate, because detecting it requires
'   inspecting elements.
'
' SHAPE VOCABULARY
'   KPR_SHAPE_SCALAR   a scalar, a single-cell Range or a 1x1 array; expands to
'                      every output position
'   KPR_SHAPE_ARRAY    a Rows x Cols payload with Rows or Cols above 1
'
'   A one-dimensional VBA array is 1xN. Worksheet Range orientation is
'   preserved. All non-scalar arguments must match exactly; there is no
'   row-to-column outer product and no implicit cross-broadcast. A multi-area,
'   zero-element, jagged or rank-3+ input and any non-Range object are
'   SHAPE_UNSUPPORTED.
'
' RANK
'   Rank is reported rather than capped, so a caller can reject a shape it does
'   not support at its own boundary. Folding rank 3 and above into rank 2 would
'   let such an array reach code that indexes it with two subscripts, raising
'   error 9 inside the caller instead of failing cleanly here.
'
' TYPE HANDOFF
'   A single-cell Range is read through Value2, so a date-formatted cell leaves
'   this module as vbDouble rather than the vbDate that a Value read produces.
'
'   Nothing downstream depends on which of the two arrives: KPR_Core_Parse
'   resolves both through the same arithmetic path and returns the same
'   date-only result, verified on the host. Value2 is used because it applies
'   no Date or Currency coercion of its own, not because the parse layer
'   requires it.
'
'   What does matter is that the argument is never Let-copied before its type
'   is tested. A Let assignment from a Variant holding an object invokes that
'   object's default member, which for a Range is Value. That silently turns a
'   wrapper into a value before any shape test runs, which makes the non-Range
'   rejection unreachable except through error 438 and materializes a
'   multi-cell Range in full before CountLarge can reject it.
'
' VISIBILITY
'   Option Private Module. Members are declared Public so other modules in the
'   project can call them; the module-level Private setting is what keeps them
'   out of the Excel function list and off the project's external surface.
'   Nothing here is supported API.
'
' ALLOWED DEPENDENCIES
'   KPR_Core_Err, for the condition vocabulary. No member raises; every
'   failure is a return value with a classified condition.
'
' PURITY
'   The static gate forbids in this module: UsedRange, Select, Activate,
'   Calculate, EnableEvents, ScreenUpdating, Selection, Application.Caller,
'   AddressOf, CallByName, Application.Run, and the date intrinsics
'   DateSerial, DateValue, CDate, Weekday, Year, Month and Day. A shape engine
'   that reads calendar values or Excel state has stopped being a shape engine.
'
' UPDATED
'   2026-09-02
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit         'Force explicit variable declarations
    Option Private Module   'Internal module: invisible outside this VBA project

'------------------------------------------------------------------------------
' SHAPE VOCABULARY
'------------------------------------------------------------------------------
    'Argument shape as seen by the engine
        Public Enum KPR_ArgShape
            KPR_SHAPE_SCALAR = 0
            KPR_SHAPE_ARRAY = 1
        End Enum

    'Capacity cap on the resolved output, per contract section 5
        Public Const KPR_MAX_ELEMENTS As Long = 100000

'
'------------------------------------------------------------------------------
'
'                              SHAPE CLASSIFICATION
'
'------------------------------------------------------------------------------
'

Public Function Array_Rank( _
    ByVal ArrayIn As Variant) _
    As Long
'
'==============================================================================
'                                  Array_Rank
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
' SIGNATURE
'   Array_Rank(ArrayIn) -> Long
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
' ERROR POLICY
'   - Does NOT raise.
'   - Clears Err as a side effect; callers must not hold pending error
'     state across this call.
'   - Restores normal error handling for the remainder of THIS procedure
'     only. VBA scopes error state per procedure, so the caller's handler
'     is unaffected either way.
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
'   2026-08-31
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ProbeDim    As Long      'Dimension currently under probe
    Dim ProbeBound  As Long      'Probe target; only success / failure matters
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default is 0 unless at least one dimension probes clean
        Array_Rank = 0
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
            Array_Rank = ProbeDim
        Next ProbeDim
    'Restore normal error handling
        Err.Clear
        On Error GoTo 0
End Function


Public Function TryUnwrapScalar( _
    ByVal ValueIn As Variant, _
    ByRef ScalarOut As Variant) _
    As Boolean
'
'==============================================================================
'                                TryUnwrapScalar
'------------------------------------------------------------------------------
' PURPOSE
'   Reduces an incoming argument to the single scalar it carries, or fails.
'
' WHY THIS EXISTS
'   Worksheet arguments arrive as scalars, as single-cell Range objects, or as
'   1x1 array wrappers depending on how the formula was written. All three mean
'   one value. This routine is the single place that decides which wrappers are
'   equivalent to a scalar, so no caller re-implements shape logic.
'
' SIGNATURE
'   TryUnwrapScalar(ValueIn, ScalarOut) -> Boolean
'
' INPUTS
'   ValueIn
'     Any Variant argument, including Range objects and arrays.
'
' OUTPUTS
'   ScalarOut (ByRef)
'     Assigned ONLY on success, to the unwrapped scalar payload.
'
' RETURNS
'   Boolean
'     TRUE  => a single scalar was extracted; ScalarOut assigned
'     FALSE => the input is not single-valued; ScalarOut untouched
'
' BEHAVIOR
'   - An object is accepted only when it is a Range of exactly one cell, read
'     through Value2.
'   - An array is accepted only when it is rank 1 or rank 2 and holds exactly
'     one element. Rank 0 and rank 3 or higher are rejected outright.
'   - A single element that is itself an object is rejected; it is a wrapper,
'     not a scalar payload.
'   - Anything else, including a multi-cell Range and a non-Range object, fails.
'   - The scalar payload is NOT inspected here. Type acceptance is owned by
'     KPR_Core_Parse.
'
' ERROR POLICY
'   - Does not raise. All failure paths return FALSE.
'   - Clears Err before returning FALSE, so no stale error state survives.
'
' DEPENDENCIES
'   - Array_Rank
'
' NOTES
'   - ValueIn is tested with IsObject directly and never Let-copied first. See
'     TYPE HANDOFF in the module header for what a Let copy would do.
'   - Rank is tested before any bound is read, so an unsupported shape is
'     rejected at the boundary rather than through the error handler.
'   - The wrapped-object rejection is a guard rather than a migrated rule. The
'     baseline had no equivalent, because its own Let copy meant an object
'     element was dereferenced before anything could test it.
'   - This helper is deliberately scalar-only. Multi-element value arguments
'     are handled by TryMaterialize plus facade traversal; element resolvers
'     use this routine as the defensive 1x1 boundary.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Rng         As Range        'Bound Range when ValueIn is a Range object
    Dim Payload     As Variant      'Isolated scalar, assigned out only on success
    Dim ArrRank     As Long         'Array rank as reported by Array_Rank

    Dim LB1         As Long         'Lower bound, dimension 1
    Dim UB1         As Long         'Upper bound, dimension 1
    Dim LB2         As Long         'Lower bound, dimension 2
    Dim UB2         As Long         'Upper bound, dimension 2

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
            'Reject non-Range objects, Nothing included
                If Not TypeOf ValueIn Is Range Then GoTo Fail
            'Bind strongly typed
                Set Rng = ValueIn
            'Require exactly one cell before reading any content
                If Rng.CountLarge <> 1 Then GoTo Fail
            'Read through Value2 so no Date or Currency coercion is applied
                Payload = Rng.Value2

    'If an array arrived, accept only a 1x1 wrapper (scalar contract)
        ElseIf IsArray(ValueIn) Then
            'Classify the array shape once
                ArrRank = Array_Rank(ValueIn)
            'Reject unsupported shapes before any bound is read
                If ArrRank < 1 Or ArrRank > 2 Then GoTo Fail
            'Two-dimensional wrapper
                If ArrRank = 2 Then
                    'Capture bounds
                        LB1 = LBound(ValueIn, 1): UB1 = UBound(ValueIn, 1)
                        LB2 = LBound(ValueIn, 2): UB2 = UBound(ValueIn, 2)
                    'Reject anything larger than 1x1
                        If (UB1 <> LB1) Or (UB2 <> LB2) Then GoTo Fail
                    'A wrapped object is not a scalar payload
                        If IsObject(ValueIn(LB1, LB2)) Then GoTo Fail
                    'Unwrap the scalar payload
                        Payload = ValueIn(LB1, LB2)
            'One-dimensional wrapper
                Else
                    'Capture bounds
                        LB1 = LBound(ValueIn): UB1 = UBound(ValueIn)
                    'Reject vectors with more than one element
                        If UB1 <> LB1 Then GoTo Fail
                    'A wrapped object is not a scalar payload
                        If IsObject(ValueIn(LB1)) Then GoTo Fail
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
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return FALSE per contract, leaving ScalarOut untouched
        TryUnwrapScalar = False
    'Do not leave stale error state for the caller to observe
        Err.Clear

End Function

'
'------------------------------------------------------------------------------
'
'                              ARRAY ENGINE SERVICES
'
'------------------------------------------------------------------------------
'

Public Function CheckCapacity( _
    ByVal Rows As Long, _
    ByVal Cols As Long, _
    ByRef Condition As KPR_Condition) _
    As Boolean
'
'==============================================================================
'                                 CheckCapacity
'------------------------------------------------------------------------------
' PURPOSE
'   Gates a shape against the 100,000-element cap.
'
' RETURNS
'   Boolean
'     TRUE  => Rows x Cols is at most KPR_MAX_ELEMENTS; Condition is NONE
'     FALSE => CAPACITY_EXCEEDED
'
' NOTES
'   - The product is taken in Double. A worksheet Range can have over a
'     million rows and sixteen thousand columns, and the Long product of a
'     large Range overflows before it can be compared.
'   - Called from the dimensions alone, before any content is read.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    If (CDbl(Rows) * CDbl(Cols)) > CDbl(KPR_MAX_ELEMENTS) Then
        Condition = KPR_COND_CAPACITY_EXCEEDED
        CheckCapacity = False
    Else
        Condition = KPR_COND_NONE
        CheckCapacity = True
    End If

End Function

Public Function TryClassifyShape( _
    ByVal ValueIn As Variant, _
    ByRef Kind As KPR_ArgShape, _
    ByRef Rows As Long, _
    ByRef Cols As Long, _
    ByRef Condition As KPR_Condition) _
    As Boolean
'
'==============================================================================
'                               TryClassifyShape
'------------------------------------------------------------------------------
' PURPOSE
'   Classifies one value argument from its type and dimensions ALONE: no
'   capacity gate, no Value2 read, no allocation, no element inspection.
'
' WHY THIS EXISTS
'   Contract section 5.2 orders the call-level stages: shape classification of
'   every argument, then control validation and broadcast resolution, then the
'   capacity gate, then traversal. A service that classified and gated in one
'   step could report CAPACITY_EXCEEDED for the first argument before a later
'   argument's SHAPE_UNSUPPORTED or a SHAPE_MISMATCH had been seen. This
'   preflight lets the facade classify everything first and gate once, on the
'   resolved output shape.
'
' OUTPUTS
'   Kind      (ByRef)  KPR_SHAPE_SCALAR or KPR_SHAPE_ARRAY
'   Rows/Cols (ByRef)  the argument's dimensions; 1 x 1 for a scalar
'   Condition (ByRef)  always assigned
'
' BEHAVIOR
'   - A Variant holding a Range is never Let-assigned before its type is tested.
'   - A single-area Range reports its Rows and Columns counts. A one-cell Range
'     is a scalar. Value2 is NOT read.
'   - A rank-1 array is 1 x N; a rank-2 array reports its bounds; a one-element
'     array is a scalar. No element is read.
'   - A multi-area Range, a zero-element array, a rank above 2 and any
'     non-Range object are SHAPE_UNSUPPORTED.
'   - A jagged array cannot be detected here, because detecting it requires
'     inspecting elements and that is forbidden before the capacity gate. It is
'     detected by TryMaterialize and reported as SHAPE_UNSUPPORTED then. The
'     contract records this as the one shape rejection that follows capacity.
'
' ERROR POLICY
'   - Does not raise. Every failure path returns FALSE with a condition.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Rng         As Range        'Bound Range when ValueIn is a Range object
    Dim Rank        As Long         'Array rank
    Dim LB1         As Long         'Lower bound, dimension 1
    Dim UB1         As Long         'Upper bound, dimension 1
    Dim LB2         As Long         'Lower bound, dimension 2
    Dim UB2         As Long         'Upper bound, dimension 2

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo Unsupported
    TryClassifyShape = False
    Condition = KPR_COND_SHAPE_UNSUPPORTED

'------------------------------------------------------------------------------
' RANGE
'------------------------------------------------------------------------------
    If IsObject(ValueIn) Then
        If Not TypeOf ValueIn Is Range Then GoTo Unsupported
        Set Rng = ValueIn
        If Rng.Areas.Count <> 1 Then GoTo Unsupported
        Rows = Rng.Rows.Count
        Cols = Rng.Columns.Count
        If (Rows = 1) And (Cols = 1) Then Kind = KPR_SHAPE_SCALAR Else Kind = KPR_SHAPE_ARRAY
        Condition = KPR_COND_NONE
        TryClassifyShape = True
        Exit Function
    End If

'------------------------------------------------------------------------------
' VBA ARRAY
'------------------------------------------------------------------------------
    If IsArray(ValueIn) Then
        Rank = Array_Rank(ValueIn)
        Select Case Rank
            Case 1
                LB1 = LBound(ValueIn): UB1 = UBound(ValueIn)
                If UB1 < LB1 Then GoTo Unsupported
                Rows = 1
                Cols = UB1 - LB1 + 1
            Case 2
                LB1 = LBound(ValueIn, 1): UB1 = UBound(ValueIn, 1)
                LB2 = LBound(ValueIn, 2): UB2 = UBound(ValueIn, 2)
                If (UB1 < LB1) Or (UB2 < LB2) Then GoTo Unsupported
                Rows = UB1 - LB1 + 1
                Cols = UB2 - LB2 + 1
            Case Else
                GoTo Unsupported
        End Select
        If (Rows = 1) And (Cols = 1) Then Kind = KPR_SHAPE_SCALAR Else Kind = KPR_SHAPE_ARRAY
        Condition = KPR_COND_NONE
        TryClassifyShape = True
        Exit Function
    End If

'------------------------------------------------------------------------------
' SCALAR
'------------------------------------------------------------------------------
    Kind = KPR_SHAPE_SCALAR
    Rows = 1
    Cols = 1
    Condition = KPR_COND_NONE
    TryClassifyShape = True
    Exit Function

'------------------------------------------------------------------------------
' UNSUPPORTED
'------------------------------------------------------------------------------
Unsupported:
    Err.Clear
    Condition = KPR_COND_SHAPE_UNSUPPORTED
    TryClassifyShape = False

End Function

Public Function TryMaterialize( _
    ByVal ValueIn As Variant, _
    ByRef Payload As Variant, _
    ByRef Kind As KPR_ArgShape, _
    ByRef Rows As Long, _
    ByRef Cols As Long, _
    ByRef Condition As KPR_Condition) _
    As Boolean
'
'==============================================================================
'                                TryMaterialize
'------------------------------------------------------------------------------
' PURPOSE
'   Classifies one value argument, gates its capacity, and snapshots it into a
'   payload the engine can index: the scalar itself, or a 1-based 2-D Variant.
'
' SIGNATURE
'   TryMaterialize(ValueIn, Payload, Kind, Rows, Cols, Condition) -> Boolean
'
' OUTPUTS
'   Payload   (ByRef)  assigned ONLY on success
'   Kind      (ByRef)  KPR_SHAPE_SCALAR or KPR_SHAPE_ARRAY
'   Rows/Cols (ByRef)  the argument's dimensions; 1 x 1 for a scalar
'   Condition (ByRef)  always assigned
'
' BEHAVIOR
'   - Type is tested before anything is copied. A Variant holding a Range is
'     never Let-assigned; that would invoke Range.Value and materialize the
'     whole Range before capacity or area count could be checked.
'   - A single-area Range reports Rows and Cols from the Range object, is
'     capacity-gated, and only THEN read once through Value2. A one-cell Range
'     is a scalar. Value2 is used as supplied: no UsedRange intersection, no
'     trimming, no per-cell reads.
'   - A 2-D VBA array reports its bounds, is capacity-gated, then normalized to
'     a 1-based copy. A 1-D array becomes 1 x N the same way. A one-element
'     array is a scalar.
'   - Elements are inspected only during the normalizing copy, which happens
'     after the gate. An element that is itself an array makes the input
'     jagged: SHAPE_UNSUPPORTED.
'   - Empty, Null and error elements are preserved at their positions. The
'     engine does not interpret them; the facade parses each element.
'   - A multi-area Range, a zero-element array, a rank above 2, and any
'     non-Range object are SHAPE_UNSUPPORTED.
'
' ERROR POLICY
'   - Does not raise. Every failure path returns FALSE with a condition.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Rng         As Range        'Bound Range when ValueIn is a Range object
    Dim Rank        As Long         'Array rank
    Dim LB1         As Long         'Lower bound, dimension 1
    Dim UB1         As Long         'Upper bound, dimension 1
    Dim LB2         As Long         'Lower bound, dimension 2
    Dim UB2         As Long         'Upper bound, dimension 2
    Dim R           As Long         'Row cursor
    Dim C           As Long         'Column cursor
    Dim Snapshot    As Variant      'Normalized 1-based copy

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Trap runtime errors from property access and convert them to a rejection
        On Error GoTo Unsupported
        TryMaterialize = False
        Condition = KPR_COND_SHAPE_UNSUPPORTED

'------------------------------------------------------------------------------
' RANGE
'------------------------------------------------------------------------------
    If IsObject(ValueIn) Then

        'Only a Range is a shape
            If Not TypeOf ValueIn Is Range Then GoTo Unsupported
            Set Rng = ValueIn

        'One area only; a union has no single rectangle
            If Rng.Areas.Count <> 1 Then GoTo Unsupported

        'Dimensions from the object, before any value is read
            Rows = Rng.Rows.Count
            Cols = Rng.Columns.Count
            If Not CheckCapacity(Rows, Cols, Condition) Then Exit Function

        'One read. A single cell comes back as a scalar, a block as 1-based 2-D
            If (Rows = 1) And (Cols = 1) Then
                Kind = KPR_SHAPE_SCALAR
                Payload = Rng.Value2
            Else
                Kind = KPR_SHAPE_ARRAY
                Payload = Rng.Value2
            End If
            Condition = KPR_COND_NONE
            TryMaterialize = True
            Exit Function

    End If

'------------------------------------------------------------------------------
' VBA ARRAY
'------------------------------------------------------------------------------
    If IsArray(ValueIn) Then

        Rank = Array_Rank(ValueIn)

        Select Case Rank

            Case 1
                'One-dimensional is 1 x N
                    LB1 = LBound(ValueIn): UB1 = UBound(ValueIn)
                    If UB1 < LB1 Then GoTo Unsupported
                    Rows = 1
                    Cols = UB1 - LB1 + 1
                    If Not CheckCapacity(Rows, Cols, Condition) Then Exit Function

                'A single element is a scalar; an element that is an array is jagged
                    If Cols = 1 Then
                        If IsArray(ValueIn(LB1)) Then GoTo Unsupported
                        Kind = KPR_SHAPE_SCALAR
                        Payload = ValueIn(LB1)
                    Else
                        ReDim Snapshot(1 To 1, 1 To Cols)
                        For C = 1 To Cols
                            If IsArray(ValueIn(LB1 + C - 1)) Then GoTo Unsupported
                            Snapshot(1, C) = ValueIn(LB1 + C - 1)
                        Next C
                        Kind = KPR_SHAPE_ARRAY
                        Payload = Snapshot
                    End If

            Case 2
                'Two-dimensional: rows then columns, bounds normalized to 1
                    LB1 = LBound(ValueIn, 1): UB1 = UBound(ValueIn, 1)
                    LB2 = LBound(ValueIn, 2): UB2 = UBound(ValueIn, 2)
                    If (UB1 < LB1) Or (UB2 < LB2) Then GoTo Unsupported
                    Rows = UB1 - LB1 + 1
                    Cols = UB2 - LB2 + 1
                    If Not CheckCapacity(Rows, Cols, Condition) Then Exit Function

                    If (Rows = 1) And (Cols = 1) Then
                        If IsArray(ValueIn(LB1, LB2)) Then GoTo Unsupported
                        Kind = KPR_SHAPE_SCALAR
                        Payload = ValueIn(LB1, LB2)
                    Else
                        ReDim Snapshot(1 To Rows, 1 To Cols)
                        For R = 1 To Rows
                            For C = 1 To Cols
                                If IsArray(ValueIn(LB1 + R - 1, LB2 + C - 1)) Then GoTo Unsupported
                                Snapshot(R, C) = ValueIn(LB1 + R - 1, LB2 + C - 1)
                            Next C
                        Next R
                        Kind = KPR_SHAPE_ARRAY
                        Payload = Snapshot
                    End If

            Case Else
                'Rank 0 cannot occur here; rank 3 and above is unsupported
                    GoTo Unsupported

        End Select

        Condition = KPR_COND_NONE
        TryMaterialize = True
        Exit Function

    End If

'------------------------------------------------------------------------------
' SCALAR
'------------------------------------------------------------------------------
    'Everything else is a scalar payload, including Empty, Null and errors
        Kind = KPR_SHAPE_SCALAR
        Rows = 1
        Cols = 1
        Payload = ValueIn
        Condition = KPR_COND_NONE
        TryMaterialize = True
        Exit Function

'------------------------------------------------------------------------------
' UNSUPPORTED
'------------------------------------------------------------------------------
Unsupported:
    'Any failure while classifying is a shape rejection, never a raise
        Err.Clear
        Condition = KPR_COND_SHAPE_UNSUPPORTED
        TryMaterialize = False

End Function

Public Function AccumulateShape( _
    ByRef OutKind As KPR_ArgShape, _
    ByRef OutRows As Long, _
    ByRef OutCols As Long, _
    ByVal ArgKind As KPR_ArgShape, _
    ByVal ArgRows As Long, _
    ByVal ArgCols As Long, _
    ByRef Condition As KPR_Condition) _
    As Boolean
'
'==============================================================================
'                               AccumulateShape
'------------------------------------------------------------------------------
' PURPOSE
'   Folds one argument's shape into the running output shape.
'
' BEHAVIOR
'   - Start with OutKind = KPR_SHAPE_SCALAR and OutRows = OutCols = 1.
'   - A scalar argument leaves the output shape unchanged: it expands.
'   - The first array argument sets the output shape.
'   - Every later array argument must match it exactly in rows and columns;
'     otherwise SHAPE_MISMATCH. There is no outer product and no
'     cross-broadcast: a 1 x N against an N x 1 is a mismatch.
'   - An all-scalar call leaves the output scalar, never a 1 x 1 array.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' FOLD
'------------------------------------------------------------------------------
    Condition = KPR_COND_NONE
    AccumulateShape = True

    'A scalar never constrains the shape
        If ArgKind = KPR_SHAPE_SCALAR Then Exit Function

    'First array sets the shape
        If OutKind = KPR_SHAPE_SCALAR Then
            OutKind = KPR_SHAPE_ARRAY
            OutRows = ArgRows
            OutCols = ArgCols
            Exit Function
        End If

    'Later arrays must match exactly
        If (ArgRows <> OutRows) Or (ArgCols <> OutCols) Then
            Condition = KPR_COND_SHAPE_MISMATCH
            AccumulateShape = False
        End If

End Function

Public Function ElementAt( _
    ByRef Payload As Variant, _
    ByVal Kind As KPR_ArgShape, _
    ByVal R As Long, _
    ByVal C As Long) _
    As Variant
'
'==============================================================================
'                                   ElementAt
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the element at 1-based position (R, C) of a materialized payload.
'
' TOTALITY
'   Total over a valid descriptor: Payload and Kind must come from a successful
'   TryMaterialize, and (R, C) must lie inside the output shape that
'   AccumulateShape resolved. Under those preconditions no failure exists, so
'   there is no condition channel. A scalar payload returns itself at every
'   position, which IS scalar expansion.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    If Kind = KPR_SHAPE_SCALAR Then
        ElementAt = Payload
    Else
        ElementAt = Payload(R, C)
    End If

End Function

Public Function TryAllocateOutput( _
    ByVal Rows As Long, _
    ByVal Cols As Long, _
    ByRef OutArr As Variant, _
    ByRef Condition As KPR_Condition) _
    As Boolean
'
'==============================================================================
'                               TryAllocateOutput
'------------------------------------------------------------------------------
' PURPOSE
'   Allocates a 1-based Rows x Cols Variant array for the output, after the
'   capacity gate.
'
' RETURNS
'   Boolean
'     TRUE  => OutArr allocated
'     FALSE => CAPACITY_EXCEEDED, or SHAPE_UNSUPPORTED for a non-positive
'              dimension
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Fresh       As Variant      'Newly allocated output

'------------------------------------------------------------------------------
' ALLOCATE
'------------------------------------------------------------------------------
    TryAllocateOutput = False
    If (Rows < 1) Or (Cols < 1) Then
        Condition = KPR_COND_SHAPE_UNSUPPORTED
        Exit Function
    End If
    If Not CheckCapacity(Rows, Cols, Condition) Then Exit Function

    ReDim Fresh(1 To Rows, 1 To Cols)
    OutArr = Fresh
    Condition = KPR_COND_NONE
    TryAllocateOutput = True

End Function

Public Function TryUnwrapControl( _
    ByVal ValueIn As Variant, _
    ByRef ScalarOut As Variant, _
    ByRef Condition As KPR_Condition) _
    As Boolean
'
'==============================================================================
'                               TryUnwrapControl
'------------------------------------------------------------------------------
' PURPOSE
'   Reduces an Opt_ control to its single scalar, reporting the contract's
'   distinct conditions for a multi-element control and an unsupported shape.
'
' BEHAVIOR
'   - Classifies independently of TryMaterialize and is NOT capacity-gated. A
'     control of 100,001 cells is CONTROL_NOT_SCALAR, not CAPACITY_EXCEEDED,
'     because the failure is that it has more than one element, however many.
'   - A multi-element control is rejected from its dimensions alone; its
'     contents are never read.
'   - A single-cell Range or one-element array unwraps to its payload.
'   - A multi-area Range, an empty array, a rank above 2 and any non-Range
'     object are SHAPE_UNSUPPORTED.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Rng         As Range        'Bound Range when ValueIn is a Range object
    Dim Rank        As Long         'Array rank
    Dim Count       As Double       'Element count, in Double for large Ranges

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo Unsupported
    TryUnwrapControl = False
    Condition = KPR_COND_SHAPE_UNSUPPORTED

'------------------------------------------------------------------------------
' RANGE
'------------------------------------------------------------------------------
    If IsObject(ValueIn) Then
        If Not TypeOf ValueIn Is Range Then GoTo Unsupported
        Set Rng = ValueIn
        If Rng.Areas.Count <> 1 Then GoTo Unsupported
        Count = CDbl(Rng.Rows.Count) * CDbl(Rng.Columns.Count)
        If Count <> 1# Then
            Condition = KPR_COND_CONTROL_NOT_SCALAR
            Exit Function
        End If
        ScalarOut = Rng.Value2
        Condition = KPR_COND_NONE
        TryUnwrapControl = True
        Exit Function
    End If

'------------------------------------------------------------------------------
' VBA ARRAY
'------------------------------------------------------------------------------
    If IsArray(ValueIn) Then
        Rank = Array_Rank(ValueIn)
        Select Case Rank
            Case 1
                If UBound(ValueIn) < LBound(ValueIn) Then GoTo Unsupported
                Count = CDbl(UBound(ValueIn) - LBound(ValueIn) + 1)
                If Count <> 1# Then Condition = KPR_COND_CONTROL_NOT_SCALAR: Exit Function
                If IsArray(ValueIn(LBound(ValueIn))) Then GoTo Unsupported
                ScalarOut = ValueIn(LBound(ValueIn))
            Case 2
                If (UBound(ValueIn, 1) < LBound(ValueIn, 1)) Or (UBound(ValueIn, 2) < LBound(ValueIn, 2)) Then GoTo Unsupported
                Count = CDbl(UBound(ValueIn, 1) - LBound(ValueIn, 1) + 1) * CDbl(UBound(ValueIn, 2) - LBound(ValueIn, 2) + 1)
                If Count <> 1# Then Condition = KPR_COND_CONTROL_NOT_SCALAR: Exit Function
                If IsArray(ValueIn(LBound(ValueIn, 1), LBound(ValueIn, 2))) Then GoTo Unsupported
                ScalarOut = ValueIn(LBound(ValueIn, 1), LBound(ValueIn, 2))
            Case Else
                GoTo Unsupported
        End Select
        Condition = KPR_COND_NONE
        TryUnwrapControl = True
        Exit Function
    End If

'------------------------------------------------------------------------------
' SCALAR
'------------------------------------------------------------------------------
    ScalarOut = ValueIn
    Condition = KPR_COND_NONE
    TryUnwrapControl = True
    Exit Function

'------------------------------------------------------------------------------
' UNSUPPORTED
'------------------------------------------------------------------------------
Unsupported:
    Err.Clear
    Condition = KPR_COND_SHAPE_UNSUPPORTED
    TryUnwrapControl = False

End Function
