Attribute VB_Name = "M_Ranges"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: M_Ranges
'------------------------------------------------------------------------------
' PURPOSE
'   The Excel adapter of the shape engine: it describes and reads Range arguments and
'   hands every other argument to CORE_Array, so that the core sees only plain values.
'
' WHY THIS EXISTS
'   The core may not touch the Excel object model, but worksheet functions receive Range
'   objects. This module is the one place that reads a Range, and it reads one only after
'   its dimensions have passed the shape and capacity rules.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - TryDescribeArgument  the shape of an argument, from dimensions alone
'     - TryReadArgument      the plain payload of an argument, after the capacity gate
'     - TryReadScalar        the scalar of a one-cell Range or a one-element array
'     - TryReadControl       the scalar of an Opt_ control, or CONTROL_NOT_SCALAR
'
' ERROR POLICY
'   No member raises; every failure is a return value, with a condition where the core
'   reports one.
'
' DEPENDENCIES
'   The Excel object model (Range: Areas, RowCount, Columns, Value2),
'   CORE_Array and CORE_Err. Calculation contracts: Arrays and evaluation
'
' NOTES
'   - The order of operations that a worksheet function keeps is in the header of
'     CORE_Array
'   - A Range is never trimmed to the used range: it is read at its full size
'   - Option Private Module keeps every member out of the Excel function list. Nothing
'     here is supported API
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------


'
'------------------------------------------------------------------------------
'
'                               RANGE ARGUMENTS
'
'------------------------------------------------------------------------------
'

Public Function TryDescribeArgument( _
    ByVal ValueIn As Variant, _
    ByRef Kind As KPRArgShape, _
    ByRef RowCount As Long, _
    ByRef ColCount As Long, _
    ByRef Condition As KPRCondition) As Boolean

'
'------------------------------------------------------------------------------
'                      DESCRIBE THE SHAPE OF AN ARGUMENT
'------------------------------------------------------------------------------
' PURPOSE
'   Describes the shape of a value argument from its type and dimensions alone: a Range
'   from its counts, anything else through CORE_Array.TryClassifyShape.
'
' WHY THIS EXISTS
'   Stage 2 must classify every argument before any content is read. A Range's
'   dimensions come from Excel, which the core may not touch.
'
' INPUTS
'   ValueIn
'     A value argument as the worksheet passed it.
'
'   Kind, RowCount, ColCount (ByRef, output)
'     The shape; assigned on success.
'
'   Condition (ByRef, output)
'     Always assigned: cdNone or cdShapeUnsupported.
'
' RETURNS
'   Boolean
'     True with the shape. False with cdShapeUnsupported for a Range of several areas, an
'     object that is not a Range, or a plain value the core rejects.
'
' BEHAVIOR
'   A single-area Range is described by its row and column counts; its values are not
'   read. Every other argument goes to CORE_Array.TryClassifyShape.
'
' ERROR POLICY
'   Does not raise. A runtime error is a shape rejection, and Err is cleared.
'
' DEPENDENCIES
'   Excel Range (Areas, RowCount, Columns), CORE_Array.TryClassifyShape, CORE_Err
'
' NOTES
'   The argument is tested before any copy: a Let copy of a Variant holding a Range would
'   read its Value
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Rng             As Range    'The argument, bound as a Range

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Convert any runtime error while reading the Range into a shape rejection
        On Error GoTo Unsupported
        TryDescribeArgument = False
        Condition = cdShapeUnsupported

'------------------------------------------------------------------------------
' RANGE
'------------------------------------------------------------------------------
    'Describe a single-area Range from its counts, without reading Value2
        If IsObject(ValueIn) Then
            If TypeOf ValueIn Is Range Then
                Set Rng = ValueIn
                If Rng.Areas.Count <> 1 Then GoTo Unsupported
                RowCount = Rng.Rows.Count
                ColCount = Rng.Columns.Count
                If (RowCount = 1) And (ColCount = 1) Then
                    Kind = skScalar
                Else
                    Kind = skArray
                End If
                Condition = cdNone
                TryDescribeArgument = True
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' PLAIN VALUE
'------------------------------------------------------------------------------
    'Every other argument, another object included, is the core's to classify
        TryDescribeArgument = CORE_Array.TryClassifyShape(ValueIn, Kind, RowCount, ColCount, Condition)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the rejection handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Unsupported:
    'Any failure while reading the Range is a shape rejection, never a raise
        Err.Clear
        Condition = cdShapeUnsupported
        TryDescribeArgument = False

End Function

Public Function TryReadArgument( _
    ByVal ValueIn As Variant, _
    ByRef Payload As Variant, _
    ByRef Kind As KPRArgShape, _
    ByRef RowCount As Long, _
    ByRef ColCount As Long, _
    ByRef Condition As KPRCondition) As Boolean

'
'------------------------------------------------------------------------------
'                      READ AN ARGUMENT INTO ITS PAYLOAD
'------------------------------------------------------------------------------
' PURPOSE
'   Reads a value argument into the plain payload the core traverses: a Range through
'   one read of Value2, anything else through CORE_Array.TryMaterialize.
'
' WHY THIS EXISTS
'   This is where a Range becomes a plain Variant array, after its dimensions have passed
'   the capacity gate, so the core never sees a Range.
'
' INPUTS
'   ValueIn
'     A value argument as the worksheet passed it.
'
'   Payload, Kind, RowCount, ColCount (ByRef, output)
'     The payload and its shape; assigned on success.
'
'   Condition (ByRef, output)
'     Always assigned: cdNone, cdCapacityExceeded or cdShapeUnsupported.
'
' RETURNS
'   Boolean
'     True with the payload: the cell's value for one cell, a 1-based two-dimensional array
'     for a block. False with cdCapacityExceeded or cdShapeUnsupported.
'
' BEHAVIOR
'   A single-area Range is gated by CORE_Array.CheckCapacity on its counts, then read once
'   through Value2, which applies no Date or Currency coercion. Every other argument goes
'   to CORE_Array.TryMaterialize.
'
' ERROR POLICY
'   Does not raise. A runtime error is a shape rejection, and Err is cleared.
'
' DEPENDENCIES
'   Excel Range (Areas, RowCount, Columns, Value2), CORE_Array.CheckCapacity,
'   CORE_Array.TryMaterialize, CORE_Err
'
' NOTES
'   A date-formatted cell is read through Value2 as a Double, not a Date; CORE_Parse reads
'   both the same way
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Rng             As Range    'The argument, bound as a Range

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Convert any runtime error while reading the Range into a shape rejection
        On Error GoTo Unsupported
        TryReadArgument = False
        Condition = cdShapeUnsupported

'------------------------------------------------------------------------------
' RANGE
'------------------------------------------------------------------------------
    'Read a single-area Range once, only after its dimensions pass the gate
        If IsObject(ValueIn) Then
            If TypeOf ValueIn Is Range Then
                Set Rng = ValueIn
                If Rng.Areas.Count <> 1 Then GoTo Unsupported
                RowCount = Rng.Rows.Count
                ColCount = Rng.Columns.Count
                If Not CORE_Array.CheckCapacity(RowCount, ColCount, Condition) Then Exit Function
                If (RowCount = 1) And (ColCount = 1) Then
                    Kind = skScalar
                Else
                    Kind = skArray
                End If
                Payload = Rng.Value2
                Condition = cdNone
                TryReadArgument = True
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' PLAIN VALUE
'------------------------------------------------------------------------------
    'Every other argument, another object included, is the core's to read
        TryReadArgument = CORE_Array.TryMaterialize(ValueIn, Payload, Kind, RowCount, ColCount, Condition)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the rejection handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Unsupported:
    'Any failure while reading the Range is a shape rejection, never a raise
        Err.Clear
        Condition = cdShapeUnsupported
        TryReadArgument = False

End Function

Public Function TryReadScalar( _
    ByVal ValueIn As Variant, _
    ByRef ScalarOut As Variant) As Boolean

'
'------------------------------------------------------------------------------
'                            READ A SCALAR ARGUMENT
'------------------------------------------------------------------------------
' PURPOSE
'   Reduces an argument to its single scalar: a one-cell Range to its value, anything else
'   through CORE_Array.TryUnwrapScalar.
'
' WHY THIS EXISTS
'   A scalar path may receive a one-cell Range, which only Excel can read.
'
' INPUTS
'   ValueIn
'     An argument as the worksheet passed it.
'
'   ScalarOut (ByRef, output)
'     Assigned only on success.
'
' RETURNS
'   Boolean
'     True with the scalar. False for a Range of more than one cell, or a value the core
'     does not unwrap.
'
' BEHAVIOR
'   Requires one area of one row and one column before reading the Range, then reads
'   its cell through Value2. Every other argument goes to CORE_Array.TryUnwrapScalar.
'
' ERROR POLICY
'   Does not raise. A runtime error is a False result, and Err is cleared.
'
' DEPENDENCIES
'   Excel Range (Areas, RowCount, Columns, Value2), CORE_Array.TryUnwrapScalar
'
' NOTES
'   One area of one row and one column is exactly one cell, as CountLarge would count it;
'   the counts are the ones the other members read
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Rng             As Range    'The argument, bound as a Range

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Convert any runtime error while reading the Range into a rejection
        On Error GoTo Unsupported
        TryReadScalar = False

'------------------------------------------------------------------------------
' RANGE
'------------------------------------------------------------------------------
    'Read a one-cell Range; reject a larger one from its count alone
        If IsObject(ValueIn) Then
            If TypeOf ValueIn Is Range Then
                Set Rng = ValueIn
                If Rng.Areas.Count <> 1 Then GoTo Unsupported
                If (Rng.Rows.Count <> 1) Or (Rng.Columns.Count <> 1) Then GoTo Unsupported
                ScalarOut = Rng.Value2
                TryReadScalar = True
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' PLAIN VALUE
'------------------------------------------------------------------------------
    'Every other argument, another object included, is the core's to unwrap
        TryReadScalar = CORE_Array.TryUnwrapScalar(ValueIn, ScalarOut)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the rejection handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Unsupported:
    'Any failure while reading the Range is a shape rejection, never a raise
        Err.Clear
        TryReadScalar = False

End Function

Public Function TryReadControl( _
    ByVal ValueIn As Variant, _
    ByRef ScalarOut As Variant, _
    ByRef Condition As KPRCondition) As Boolean

'
'------------------------------------------------------------------------------
'                           READ AN OPTIONAL CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Reduces an Opt_ control to its single scalar: a one-cell Range to its value, anything
'   else through CORE_Array.TryUnwrapControl.
'
' WHY THIS EXISTS
'   A control may be a Range, and a multi-cell one is CONTROL_NOT_SCALAR whatever its size,
'   so its cells are counted, never read.
'
' INPUTS
'   ValueIn
'     A control as the worksheet passed it.
'
'   ScalarOut (ByRef, output)
'     Assigned only on success.
'
'   Condition (ByRef, output)
'     Always assigned: cdNone, cdControlNotScalar or cdShapeUnsupported.
'
' RETURNS
'   Boolean
'     True with the scalar. False with cdControlNotScalar for more than one cell, or
'     cdShapeUnsupported for a Range of several areas or a value the core rejects.
'
' BEHAVIOR
'   Counts a single-area Range in Double, reads it only when it is one cell, and sends
'   every other control to CORE_Array.TryUnwrapControl. It is not capacity-gated.
'
' ERROR POLICY
'   Does not raise. A runtime error is a shape rejection, and Err is cleared.
'
' DEPENDENCIES
'   Excel Range (Areas, RowCount, Columns, Value2), CORE_Array.TryUnwrapControl, CORE_Err.
'   Data contract: Function arguments, Control values
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
    Dim Rng             As Range    'The argument, bound as a Range
    Dim Cells           As Double   'Cell count, in Double for a large Range

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Convert any runtime error while reading the Range into a shape rejection
        On Error GoTo Unsupported
        TryReadControl = False
        Condition = cdShapeUnsupported

'------------------------------------------------------------------------------
' RANGE
'------------------------------------------------------------------------------
    'Read a one-cell Range; reject a larger one from its counts alone
        If IsObject(ValueIn) Then
            If TypeOf ValueIn Is Range Then
                Set Rng = ValueIn
                If Rng.Areas.Count <> 1 Then GoTo Unsupported
                Cells = VBA.CDbl(Rng.Rows.Count) * VBA.CDbl(Rng.Columns.Count)
                If Cells <> 1# Then
                    Condition = cdControlNotScalar
                    Exit Function
                End If
                ScalarOut = Rng.Value2
                Condition = cdNone
                TryReadControl = True
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' PLAIN VALUE
'------------------------------------------------------------------------------
    'Every other control, another object included, is the core's to unwrap
        TryReadControl = CORE_Array.TryUnwrapControl(ValueIn, ScalarOut, Condition)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the rejection handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Unsupported:
    'Any failure while reading the Range is a shape rejection, never a raise
        Err.Clear
        Condition = cdShapeUnsupported
        TryReadControl = False

End Function
