Attribute VB_Name = "KPR_DATES_DAYS"
'==============================================================================
' MODULE: KPR_DATES_DAYS
'------------------------------------------------------------------------------
' PURPOSE
'   The worksheet-facing date facade for the KPR toolkit: date primitives, date
'   arithmetic, weekday locators and pillar labeling.
'
'   Public functions are designed for direct use in Excel formulas (Function
'   Wizard, tooltips, scalar usage) with deterministic input handling and native
'   Excel error values on failure.
'
' WHY THE NAME
'   Facade components are named in upper case to mark them as worksheet-facing
'   rather than internal, and this one is suffixed _DAYS because
'   KPR_DATES_CALENDARS and KPR_DATES_BUSINESSDAYS follow in later milestones.
'   Internal modules keep mixed-case KPR_Core_* names. The convention is
'   documented rather than enforced by a static casing rule.
'
' SCOPE / PUBLIC SURFACE
'   The complete 22-name surface frozen in contract section 2. Twenty-one
'   value-taking functions, each with a private element implementation, and
'   one scalar diagnostic.
'
'   - Day primitives
'       * KPR_Dates_DayOfWeek           (Optional Opt_WeekBaseMonday)
'       * KPR_Dates_DaysInMonth
'       * KPR_Dates_DaysInYear          (takes YearIn, not a date)
'
'   - Month, quarter and year boundaries
'       * KPR_Dates_BeginOfMonth        * KPR_Dates_EndOfMonth
'       * KPR_Dates_BeginOfQuarter      * KPR_Dates_EndOfQuarter
'       * KPR_Dates_BeginOfYear         * KPR_Dates_EndOfYear
'
'   - Boundary predicates
'       * KPR_Dates_IsMonthEnd          * KPR_Dates_IsQuarterEnd
'       * KPR_Dates_IsYearEnd           * KPR_Dates_IsLeapYear (takes YearIn)
'
'   - Date arithmetic (EOM-aware)
'       * KPR_Dates_AddDays             * KPR_Dates_AddWeeks
'       * KPR_Dates_AddMonths           (Optional Opt_KeepEOM)
'       * KPR_Dates_AddYears            (Optional Opt_KeepEOM)
'
'   - Weekday locators
'       * KPR_Dates_NthWeekdayOfMonth   (Optional Opt_WeekBaseMonday)
'       * KPR_Dates_LastWeekdayOfMonth  (Optional Opt_WeekBaseMonday)
'
'   - Pillar / tenor formatting
'       * KPR_Dates_PillarFromDates     (Optional Opt_Rounding)
'       * KPR_Dates_DateFromPillar
'
'   - Host diagnostic (scalar only, deliberately volatile)
'       * KPR_Dates_HostDateSystem
'
'   Semantic result types, all declared As Variant so a native error can be
'   returned: Date for the thirteen boundary, arithmetic, locator and
'   DateFromPillar functions; Long for DayOfWeek, DaysInMonth, DaysInYear and
'   HostDateSystem; Boolean for the four Is* predicates; String for
'   PillarFromDates.
'
'   The static gate pins this inventory exactly: a missing name, an extra
'   KPR_Dates_* name, the legacy plural DatesFromPillar or any _Spill twin
'   fails the build. That control is temporary and is replaced by the #26
'   manifest.
'
' ARCHITECTURE
'   This module is the only place a supported KPR_Dates_* function is declared.
'   It owns dispatch and worksheet-facing error mapping, and delegates every
'   computation:
'
'       KPR_Core_Array   shape: classification, materialization and broadcasting
'       KPR_Core_Parse   value: scalar -> date-only VBA Date
'       KPR_Core_Dates   calendar: window, arithmetic, boundaries, pillars
'       KPR_Core_Err     construction of native Excel error values
'
'   The four core modules declare Option Private Module and never depend on this
'   facade, on registration, on UI, on tests or on demo code. This module
'   declares no Option Private Module, because its members must reach the
'   worksheet.
'
' CALLER AND DATE-SYSTEM CONTRACT
'   Every value-taking function runs PassHostGuard once, before any argument
'   is resolved. The guard reads Application.Caller through one guarded object
'   assignment and classifies it:
'
'       worksheet Range   read that Range's own workbook Date1904
'                           1900 => proceed
'                           1904 => one call-level #N/A (HOST_DATE1904)
'                           unreadable => one call-level #N/A (HOST_UNRESOLVED)
'       anything else     "no worksheet host could be identified":
'                         proceed under the documented 1900 serial contract
'
'   A 1904 host is refused rather than compensated, because every serial the
'   caller supplied would otherwise have two possible meanings. ActiveWorkbook,
'   ThisWorkbook and ActiveSheet are never consulted; the static gate rejects
'   them on the host-resolution path.
'
'   "No worksheet host could be identified" is not a claim that the caller is
'   direct VBA. Direct VBA, the Immediate window, Application.Run and the
'   regression harness are the certified uses of that path. Other non-Range
'   callers exist (shape macros, data validation, chart series, conditional
'   formatting, defined names) and v0.0.2 makes no claim for them; they are
'   #29 probe targets. See TryResolveHostDateSystem for the full list.
'
'   Library-produced #N/A and a propagated incoming #N/A are the same Excel
'   value. KPR_Dates_HostDateSystem is the discriminator: 1904 identifies host
'   refusal; 1900 leaves an incoming error or another input path as the
'   source. It is the only volatile function in the library.
'
' ELEMENT IMPLEMENTATIONS
'   Every value-taking public function is a thin wrapper over one private
'   Elem_* function. The wrapper is call-level: it runs the host guard once,
'   resolves the optional controls once, and calls the element. The element is
'   element-level: it resolves the raw value arguments in signature order,
'   computes, and returns the value or the element's own native error.
'
'       host guard -> optional controls -> element arguments in signature order
'
'   The scalar call is the 1x1 case of the element. A multi-element call loops
'   the same Elem_* function over the resolved shape through KPR_Core_Array;
'   nothing about an element differs between the two, and no element
'   ever calls the guard, which the static gate enforces.
'
' SHAPE CONTRACT
'   - Value arguments vectorize. A scalar, single-cell Range or 1x1 wrapper
'     expands to the resolved output shape; every non-scalar value argument
'     must match exactly in rows and columns. No outer product, no implicit
'     cross-broadcast. Worksheet orientation is preserved and a 1-D VBA array
'     is 1xN.
'   - An all-scalar call returns a scalar. Any other call returns a 1-based
'     2-D Variant of the resolved shape, evaluated row-major, each element
'     resolved independently so a blank or an incoming error at one position
'     never suppresses a valid neighbour.
'   - Opt_ controls never vectorize: a multi-element control is call-level
'     #VALUE! under CONTROL_NOT_SCALAR.
'   - The output is capped at 100,000 elements (#NUM!, CAPACITY_EXCEEDED),
'     decided from dimensions before any Range is read. Multi-area, empty,
'     rank-3+ and non-Range inputs are call-level #VALUE!; a jagged array is
'     detected during materialization, after the cap, because detecting it
'     requires inspecting elements.
'   - Every wrapper applies contract section 5.2's stages in order: host guard,
'     classify every value argument, controls and broadcast resolution, cap,
'     materialize, traverse. No element implementation ever runs the guard or
'     reads Application.Caller.
'
' COMPATIBILITY
'   Scalar calls are supported on every Excel version certified for scalar
'   use. Multi-cell calls are tested, supported and claimed on dynamic-array
'   Excel only. No Ctrl+Shift+Enter or other legacy multi-cell claim is made,
'   there is no version detection, and Excel's own spill-placement errors are
'   Excel's to raise.
'
' DESIGN / INPUT NORMALIZATION
'   - Every DateIn-style argument is a Variant funnelled through TryResolveDate,
'     which composes the three boundaries in a fixed order:
'         unwrap shape -> propagate incoming error -> strict parse/window gate
'   - KPR_Core_Parse applies the window once after date-only normalization. Its
'     serial/year constants are statically pinned to the Date constants owned by
'     KPR_Core_Dates because the dependency matrix forbids a direct call.
'   - The two weekday locators take a year and a month rather than a date, so
'     they never reach TryResolveDate. They gate their own result against the
'     window instead. Without that gate a year check alone is too coarse: the
'     window floor is 1 March 1900, so January and February 1900 pass a
'     year-granularity test and would return an out-of-window date.
'
' ERROR POLICY (USER FACING)
'   Public UDFs return either a valid scalar result or a native Excel error.
'   No public function returns a message string.
'
'       #VALUE!  the input cannot be interpreted
'                (unparseable date, unsupported or mismatched shape, malformed
'                 pillar, month outside 1..12, weekday outside 1..7,
'                 occurrence outside 1..5, a non-scalar control, or a control
'                 of the wrong type or with an unknown token)
'
'       #NUM!    the input is well formed but the answer does not exist or
'                falls outside the supported date window
'                (no such weekday occurrence in the month, a located weekday
'                 outside the window, arithmetic or pillar shift beyond
'                 KPR_MIN_DATE / KPR_MAX_DATE)
'
'       #N/A     the result is unavailable in this host configuration
'                (identified 1904 worksheet caller, or an identified worksheet
'                 caller whose date system cannot be read)
'
'       incoming native errors are returned unchanged
'
'   Rationale: message strings are invisible to IFERROR / ISERROR, break any
'   downstream arithmetic that references the cell, and cannot be aggregated or
'   filtered. Native error values are the only return that composes.
'
' CALENDAR ARITHMETIC
'   No function in this module writes the day-0 idiom DateSerial(y, m + 1, 0).
'   It raises error 5 at both ends of the VBA Date range, and behind a
'   defensive handler that surfaces as a worksheet error indistinguishable
'   from a genuine rejection. Month length comes from KPR_Core_Dates.DaysInMonth
'   and month-end from KPR_Core_Dates.EndOfMonth, both of which are total.
'
' INTERNAL ERROR POLICY
'   - Core routines are Try-style (Boolean return, result ByRef) or plain
'     computation on already-validated inputs.
'   - This facade owns all worksheet-facing error behavior.
'   - Defensive catch-all handlers, in wrappers and in elements alike, are
'     containment only. They must be unreachable in contract-conforming
'     execution; an activation is a defect, never an expected outcome, and
'     is not a documented #VALUE! condition. A raise that a handler converts into a
'     plausible worksheet error is invisible from the sheet, so a condition
'     that can be tested is tested rather than trapped.
'
' DEPENDENCIES / INTEGRATION
'   - KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates, KPR_Core_Err
'   - Private helpers in this module: the Elem_* element implementations,
'     TryResolveDate, TryResolveLong, TryResolveBool, TryResolveRounding,
'     TryResolvePillar, TryResolveHostDateSystem, PassHostGuard
'   - No dependency on registration, UI, tests or demo code.
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
    Option Explicit  'Force explicit variable declarations

'
'------------------------------------------------------------------------------
'
'                         PUBLIC API - DAY PRIMITIVES                          
'
'------------------------------------------------------------------------------
'

Public Function KPR_Dates_DayOfWeek( _
    ByVal DateIn As Variant, _
    Optional ByVal Opt_WeekBaseMonday As Variant = True) _
    As Variant
'
'==============================================================================
'                          KPR_Dates_DayOfWeek
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the weekday index of a date under the selected week base.
'
' SIGNATURE
'   KPR_Dates_DayOfWeek(DateIn, [Opt_WeekBaseMonday]) -> Variant (Long)
'
' INPUTS
'   DateIn
'     Date, numeric serial, or ISO YYYY-MM-DD text; single-cell Range or 1x1 wrapper accepted.
'   Opt_WeekBaseMonday
'     Optional. Omitted or Empty selects True. Native Boolean only.
'
' RETURNS
'   Variant
'     Long on success, or a native Excel error value.
'
' ERROR POLICY (USER FACING)
'   #VALUE!  a multi-cell optional control; a shape that is not a scalar,
'            Range or array, or a multi-area, empty, jagged or rank-3+ input;
'            two non-scalar arguments of different shapes; a date argument that is not scalar, not a date, or malformed text; an optional control of the wrong type or an unknown token
'   #NUM!    an output of 100,001 or more elements; a date outside KPR_MIN_DATE .. KPR_MAX_DATE
'   #N/A     an identified 1904 worksheet caller, or an identified caller
'            whose date system cannot be read
'   Incoming native errors are returned unchanged.
'
' SHAPE
'   Value arguments may be a scalar, a single-cell Range or 1x1 wrapper, or a
'   multi-element Range or array. A scalar expands; every non-scalar argument
'   must match exactly in rows and columns. An all-scalar call returns a
'   scalar; any other call returns a 1-based 2-D Variant of the resolved
'   shape, evaluated row-major, with each element resolved independently.
'   Multi-cell results are supported on dynamic-array Excel only.
'
' ORDER OF EVALUATION (call-level, contract section 5.2)
'   host guard -> classify every value argument -> optional controls and
'   broadcast resolution -> element cap -> materialize -> traverse.
'   Within an element, value arguments resolve in signature order.
'
' DEPENDENCIES
'   - PassHostGuard
'   - Elem_DayOfWeek
'   - KPR_Core_Array: TryClassifyShape, AccumulateShape, CheckCapacity,
'     TryMaterialize, TryAllocateOutput, ElementAt
'   - KPR_Core_Err.ErrForCondition
'   - TryResolveBool (KPR_Core_Array, KPR_Core_Parse)
'
' NOTES
'   - TRUE selects Monday-based numbering (ISO habit); FALSE selects Sunday-based.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P1              As Variant   'Materialized payload of DateIn
    Dim K1              As KPR_ArgShape 'Shape kind of DateIn
    Dim R1              As Long      'Rows of DateIn
    Dim C1              As Long      'Cols of DateIn
    Dim WkMonday        As Boolean    'Resolved Opt_WeekBaseMonday
    Dim OutKind         As KPR_ArgShape 'Resolved output kind
    Dim OutRows         As Long      'Resolved output rows
    Dim OutCols         As Long      'Resolved output cols
    Dim OutArr          As Variant   'Output array for a multi-element call
    Dim R               As Long      'Row cursor
    Dim C               As Long      'Column cursor
    Dim Condition       As KPR_Condition 'Call-level shape or capacity condition
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' CALL-LEVEL STAGES (contract section 5.2)
'------------------------------------------------------------------------------
    'Stage 1: refuse a 1904 worksheet host before any argument is touched
        If Not PassHostGuard(FailErr) Then GoTo Fail

    'Stage 2: classify every value argument from type and dimensions alone
        If Not TryClassifyShape(DateIn, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 3a: optional controls, scalar or 1x1 only (CONTROL_NOT_SCALAR)
        If Not TryResolveBool(Opt_WeekBaseMonday, True, WkMonday, FailErr) Then GoTo Fail

    'Stage 3b: exact-shape broadcast resolution (SHAPE_MISMATCH)
        OutKind = KPR_SHAPE_SCALAR: OutRows = 1: OutCols = 1
        If Not AccumulateShape(OutKind, OutRows, OutCols, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 4: the element cap, once, on the resolved output shape
        If Not CheckCapacity(OutRows, OutCols, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 5: only now may content be read
        If Not TryMaterialize(DateIn, P1, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

'------------------------------------------------------------------------------
' TRAVERSE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar: the 1x1 case of the element
        If OutKind = KPR_SHAPE_SCALAR Then
            KPR_Dates_DayOfWeek = Elem_DayOfWeek(P1, WkMonday)
            Exit Function
        End If

    'A multi-element call returns a 1-based 2-D array of the resolved shape,
    'evaluated row-major; each element resolves its own inputs
        If Not TryAllocateOutput(OutRows, OutCols, OutArr, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        For R = 1 To OutRows
            For C = 1 To OutCols
                OutArr(R, C) = Elem_DayOfWeek(ElementAt(P1, K1, R, C), WkMonday)
            Next C
        Next R
        KPR_Dates_DayOfWeek = OutArr
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_DayOfWeek = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed with #VALUE!; reaching here is a defect
        FailErr = ErrValue()
        Resume Fail

End Function
Public Function KPR_Dates_DaysInMonth( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
'                          KPR_Dates_DaysInMonth
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Gregorian length of the containing month.
'
' SIGNATURE
'   KPR_Dates_DaysInMonth(DateIn) -> Variant (Long)
'
' INPUTS
'   DateIn
'     Date, numeric serial, or ISO YYYY-MM-DD text; single-cell Range or 1x1 wrapper accepted.
'
' RETURNS
'   Variant
'     Long on success, or a native Excel error value.
'
' ERROR POLICY (USER FACING)
'   #VALUE!  a multi-cell optional control; a shape that is not a scalar,
'            Range or array, or a multi-area, empty, jagged or rank-3+ input;
'            two non-scalar arguments of different shapes; a date argument that is not scalar, not a date, or malformed text
'   #NUM!    an output of 100,001 or more elements; a date outside KPR_MIN_DATE .. KPR_MAX_DATE
'   #N/A     an identified 1904 worksheet caller, or an identified caller
'            whose date system cannot be read
'   Incoming native errors are returned unchanged.
'
' SHAPE
'   Value arguments may be a scalar, a single-cell Range or 1x1 wrapper, or a
'   multi-element Range or array. A scalar expands; every non-scalar argument
'   must match exactly in rows and columns. An all-scalar call returns a
'   scalar; any other call returns a 1-based 2-D Variant of the resolved
'   shape, evaluated row-major, with each element resolved independently.
'   Multi-cell results are supported on dynamic-array Excel only.
'
' ORDER OF EVALUATION (call-level, contract section 5.2)
'   host guard -> classify every value argument -> optional controls and
'   broadcast resolution -> element cap -> materialize -> traverse.
'   Within an element, value arguments resolve in signature order.
'
' DEPENDENCIES
'   - PassHostGuard
'   - Elem_DaysInMonth
'   - KPR_Core_Array: TryClassifyShape, AccumulateShape, CheckCapacity,
'     TryMaterialize, TryAllocateOutput, ElementAt
'   - KPR_Core_Err.ErrForCondition
'
' NOTES
'   - None.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P1              As Variant   'Materialized payload of DateIn
    Dim K1              As KPR_ArgShape 'Shape kind of DateIn
    Dim R1              As Long      'Rows of DateIn
    Dim C1              As Long      'Cols of DateIn
    Dim OutKind         As KPR_ArgShape 'Resolved output kind
    Dim OutRows         As Long      'Resolved output rows
    Dim OutCols         As Long      'Resolved output cols
    Dim OutArr          As Variant   'Output array for a multi-element call
    Dim R               As Long      'Row cursor
    Dim C               As Long      'Column cursor
    Dim Condition       As KPR_Condition 'Call-level shape or capacity condition
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' CALL-LEVEL STAGES (contract section 5.2)
'------------------------------------------------------------------------------
    'Stage 1: refuse a 1904 worksheet host before any argument is touched
        If Not PassHostGuard(FailErr) Then GoTo Fail

    'Stage 2: classify every value argument from type and dimensions alone
        If Not TryClassifyShape(DateIn, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 3b: exact-shape broadcast resolution (SHAPE_MISMATCH)
        OutKind = KPR_SHAPE_SCALAR: OutRows = 1: OutCols = 1
        If Not AccumulateShape(OutKind, OutRows, OutCols, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 4: the element cap, once, on the resolved output shape
        If Not CheckCapacity(OutRows, OutCols, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 5: only now may content be read
        If Not TryMaterialize(DateIn, P1, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

'------------------------------------------------------------------------------
' TRAVERSE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar: the 1x1 case of the element
        If OutKind = KPR_SHAPE_SCALAR Then
            KPR_Dates_DaysInMonth = Elem_DaysInMonth(P1)
            Exit Function
        End If

    'A multi-element call returns a 1-based 2-D array of the resolved shape,
    'evaluated row-major; each element resolves its own inputs
        If Not TryAllocateOutput(OutRows, OutCols, OutArr, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        For R = 1 To OutRows
            For C = 1 To OutCols
                OutArr(R, C) = Elem_DaysInMonth(ElementAt(P1, K1, R, C))
            Next C
        Next R
        KPR_Dates_DaysInMonth = OutArr
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_DaysInMonth = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed with #VALUE!; reaching here is a defect
        FailErr = ErrValue()
        Resume Fail

End Function
Public Function KPR_Dates_DaysInYear( _
    ByVal YearIn As Variant) _
    As Variant
'
'==============================================================================
'                          KPR_Dates_DaysInYear
'------------------------------------------------------------------------------
' PURPOSE
'   Returns 366 for a Gregorian leap year and 365 otherwise.
'
' SIGNATURE
'   KPR_Dates_DaysInYear(YearIn) -> Variant (Long)
'
' INPUTS
'   YearIn
'     Native integral numeric, domain 1900 through 9999.
'
' RETURNS
'   Variant
'     Long on success, or a native Excel error value.
'
' ERROR POLICY (USER FACING)
'   #VALUE!  a multi-cell optional control; a shape that is not a scalar,
'            Range or array, or a multi-area, empty, jagged or rank-3+ input;
'            two non-scalar arguments of different shapes; an integer argument that is fractional, Boolean, text, or outside its domain
'   #NUM!    an output of 100,001 or more elements; an integer outside the Long range
'   #N/A     an identified 1904 worksheet caller, or an identified caller
'            whose date system cannot be read
'   Incoming native errors are returned unchanged.
'
' SHAPE
'   Value arguments may be a scalar, a single-cell Range or 1x1 wrapper, or a
'   multi-element Range or array. A scalar expands; every non-scalar argument
'   must match exactly in rows and columns. An all-scalar call returns a
'   scalar; any other call returns a 1-based 2-D Variant of the resolved
'   shape, evaluated row-major, with each element resolved independently.
'   Multi-cell results are supported on dynamic-array Excel only.
'
' ORDER OF EVALUATION (call-level, contract section 5.2)
'   host guard -> classify every value argument -> optional controls and
'   broadcast resolution -> element cap -> materialize -> traverse.
'   Within an element, value arguments resolve in signature order.
'
' DEPENDENCIES
'   - PassHostGuard
'   - Elem_DaysInYear
'   - KPR_Core_Array: TryClassifyShape, AccumulateShape, CheckCapacity,
'     TryMaterialize, TryAllocateOutput, ElementAt
'   - KPR_Core_Err.ErrForCondition
'
' NOTES
'   - Takes a calendar year, not a date. No date is constructed, so no window gate applies and year 1900 is fully in the domain.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P1              As Variant   'Materialized payload of YearIn
    Dim K1              As KPR_ArgShape 'Shape kind of YearIn
    Dim R1              As Long      'Rows of YearIn
    Dim C1              As Long      'Cols of YearIn
    Dim OutKind         As KPR_ArgShape 'Resolved output kind
    Dim OutRows         As Long      'Resolved output rows
    Dim OutCols         As Long      'Resolved output cols
    Dim OutArr          As Variant   'Output array for a multi-element call
    Dim R               As Long      'Row cursor
    Dim C               As Long      'Column cursor
    Dim Condition       As KPR_Condition 'Call-level shape or capacity condition
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' CALL-LEVEL STAGES (contract section 5.2)
'------------------------------------------------------------------------------
    'Stage 1: refuse a 1904 worksheet host before any argument is touched
        If Not PassHostGuard(FailErr) Then GoTo Fail

    'Stage 2: classify every value argument from type and dimensions alone
        If Not TryClassifyShape(YearIn, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 3b: exact-shape broadcast resolution (SHAPE_MISMATCH)
        OutKind = KPR_SHAPE_SCALAR: OutRows = 1: OutCols = 1
        If Not AccumulateShape(OutKind, OutRows, OutCols, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 4: the element cap, once, on the resolved output shape
        If Not CheckCapacity(OutRows, OutCols, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 5: only now may content be read
        If Not TryMaterialize(YearIn, P1, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

'------------------------------------------------------------------------------
' TRAVERSE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar: the 1x1 case of the element
        If OutKind = KPR_SHAPE_SCALAR Then
            KPR_Dates_DaysInYear = Elem_DaysInYear(P1)
            Exit Function
        End If

    'A multi-element call returns a 1-based 2-D array of the resolved shape,
    'evaluated row-major; each element resolves its own inputs
        If Not TryAllocateOutput(OutRows, OutCols, OutArr, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        For R = 1 To OutRows
            For C = 1 To OutCols
                OutArr(R, C) = Elem_DaysInYear(ElementAt(P1, K1, R, C))
            Next C
        Next R
        KPR_Dates_DaysInYear = OutArr
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_DaysInYear = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed with #VALUE!; reaching here is a defect
        FailErr = ErrValue()
        Resume Fail

End Function
'
'------------------------------------------------------------------------------
'
'               PUBLIC API - MONTH, QUARTER AND YEAR BOUNDARIES                
'
'------------------------------------------------------------------------------
'

Public Function KPR_Dates_BeginOfMonth( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
'                          KPR_Dates_BeginOfMonth
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the first date of the containing month.
'
' SIGNATURE
'   KPR_Dates_BeginOfMonth(DateIn) -> Variant (Date)
'
' INPUTS
'   DateIn
'     Date, numeric serial, or ISO YYYY-MM-DD text; single-cell Range or 1x1 wrapper accepted.
'
' RETURNS
'   Variant
'     Date on success, or a native Excel error value.
'
' ERROR POLICY (USER FACING)
'   #VALUE!  a multi-cell optional control; a shape that is not a scalar,
'            Range or array, or a multi-area, empty, jagged or rank-3+ input;
'            two non-scalar arguments of different shapes; a date argument that is not scalar, not a date, or malformed text
'   #NUM!    an output of 100,001 or more elements; a date outside KPR_MIN_DATE .. KPR_MAX_DATE; a result outside the supported window
'   #N/A     an identified 1904 worksheet caller, or an identified caller
'            whose date system cannot be read
'   Incoming native errors are returned unchanged.
'
' SHAPE
'   Value arguments may be a scalar, a single-cell Range or 1x1 wrapper, or a
'   multi-element Range or array. A scalar expands; every non-scalar argument
'   must match exactly in rows and columns. An all-scalar call returns a
'   scalar; any other call returns a 1-based 2-D Variant of the resolved
'   shape, evaluated row-major, with each element resolved independently.
'   Multi-cell results are supported on dynamic-array Excel only.
'
' ORDER OF EVALUATION (call-level, contract section 5.2)
'   host guard -> classify every value argument -> optional controls and
'   broadcast resolution -> element cap -> materialize -> traverse.
'   Within an element, value arguments resolve in signature order.
'
' DEPENDENCIES
'   - PassHostGuard
'   - Elem_BeginOfMonth
'   - KPR_Core_Array: TryClassifyShape, AccumulateShape, CheckCapacity,
'     TryMaterialize, TryAllocateOutput, ElementAt
'   - KPR_Core_Err.ErrForCondition
'
' NOTES
'   - None.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P1              As Variant   'Materialized payload of DateIn
    Dim K1              As KPR_ArgShape 'Shape kind of DateIn
    Dim R1              As Long      'Rows of DateIn
    Dim C1              As Long      'Cols of DateIn
    Dim OutKind         As KPR_ArgShape 'Resolved output kind
    Dim OutRows         As Long      'Resolved output rows
    Dim OutCols         As Long      'Resolved output cols
    Dim OutArr          As Variant   'Output array for a multi-element call
    Dim R               As Long      'Row cursor
    Dim C               As Long      'Column cursor
    Dim Condition       As KPR_Condition 'Call-level shape or capacity condition
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' CALL-LEVEL STAGES (contract section 5.2)
'------------------------------------------------------------------------------
    'Stage 1: refuse a 1904 worksheet host before any argument is touched
        If Not PassHostGuard(FailErr) Then GoTo Fail

    'Stage 2: classify every value argument from type and dimensions alone
        If Not TryClassifyShape(DateIn, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 3b: exact-shape broadcast resolution (SHAPE_MISMATCH)
        OutKind = KPR_SHAPE_SCALAR: OutRows = 1: OutCols = 1
        If Not AccumulateShape(OutKind, OutRows, OutCols, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 4: the element cap, once, on the resolved output shape
        If Not CheckCapacity(OutRows, OutCols, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 5: only now may content be read
        If Not TryMaterialize(DateIn, P1, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

'------------------------------------------------------------------------------
' TRAVERSE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar: the 1x1 case of the element
        If OutKind = KPR_SHAPE_SCALAR Then
            KPR_Dates_BeginOfMonth = Elem_BeginOfMonth(P1)
            Exit Function
        End If

    'A multi-element call returns a 1-based 2-D array of the resolved shape,
    'evaluated row-major; each element resolves its own inputs
        If Not TryAllocateOutput(OutRows, OutCols, OutArr, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        For R = 1 To OutRows
            For C = 1 To OutCols
                OutArr(R, C) = Elem_BeginOfMonth(ElementAt(P1, K1, R, C))
            Next C
        Next R
        KPR_Dates_BeginOfMonth = OutArr
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_BeginOfMonth = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed with #VALUE!; reaching here is a defect
        FailErr = ErrValue()
        Resume Fail

End Function
Public Function KPR_Dates_EndOfMonth( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
'                          KPR_Dates_EndOfMonth
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the last date of the containing month.
'
' SIGNATURE
'   KPR_Dates_EndOfMonth(DateIn) -> Variant (Date)
'
' INPUTS
'   DateIn
'     Date, numeric serial, or ISO YYYY-MM-DD text; single-cell Range or 1x1 wrapper accepted.
'
' RETURNS
'   Variant
'     Date on success, or a native Excel error value.
'
' ERROR POLICY (USER FACING)
'   #VALUE!  a multi-cell optional control; a shape that is not a scalar,
'            Range or array, or a multi-area, empty, jagged or rank-3+ input;
'            two non-scalar arguments of different shapes; a date argument that is not scalar, not a date, or malformed text
'   #NUM!    an output of 100,001 or more elements; a date outside KPR_MIN_DATE .. KPR_MAX_DATE; a result outside the supported window
'   #N/A     an identified 1904 worksheet caller, or an identified caller
'            whose date system cannot be read
'   Incoming native errors are returned unchanged.
'
' SHAPE
'   Value arguments may be a scalar, a single-cell Range or 1x1 wrapper, or a
'   multi-element Range or array. A scalar expands; every non-scalar argument
'   must match exactly in rows and columns. An all-scalar call returns a
'   scalar; any other call returns a 1-based 2-D Variant of the resolved
'   shape, evaluated row-major, with each element resolved independently.
'   Multi-cell results are supported on dynamic-array Excel only.
'
' ORDER OF EVALUATION (call-level, contract section 5.2)
'   host guard -> classify every value argument -> optional controls and
'   broadcast resolution -> element cap -> materialize -> traverse.
'   Within an element, value arguments resolve in signature order.
'
' DEPENDENCIES
'   - PassHostGuard
'   - Elem_EndOfMonth
'   - KPR_Core_Array: TryClassifyShape, AccumulateShape, CheckCapacity,
'     TryMaterialize, TryAllocateOutput, ElementAt
'   - KPR_Core_Err.ErrForCondition
'
' NOTES
'   - None.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P1              As Variant   'Materialized payload of DateIn
    Dim K1              As KPR_ArgShape 'Shape kind of DateIn
    Dim R1              As Long      'Rows of DateIn
    Dim C1              As Long      'Cols of DateIn
    Dim OutKind         As KPR_ArgShape 'Resolved output kind
    Dim OutRows         As Long      'Resolved output rows
    Dim OutCols         As Long      'Resolved output cols
    Dim OutArr          As Variant   'Output array for a multi-element call
    Dim R               As Long      'Row cursor
    Dim C               As Long      'Column cursor
    Dim Condition       As KPR_Condition 'Call-level shape or capacity condition
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' CALL-LEVEL STAGES (contract section 5.2)
'------------------------------------------------------------------------------
    'Stage 1: refuse a 1904 worksheet host before any argument is touched
        If Not PassHostGuard(FailErr) Then GoTo Fail

    'Stage 2: classify every value argument from type and dimensions alone
        If Not TryClassifyShape(DateIn, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 3b: exact-shape broadcast resolution (SHAPE_MISMATCH)
        OutKind = KPR_SHAPE_SCALAR: OutRows = 1: OutCols = 1
        If Not AccumulateShape(OutKind, OutRows, OutCols, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 4: the element cap, once, on the resolved output shape
        If Not CheckCapacity(OutRows, OutCols, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 5: only now may content be read
        If Not TryMaterialize(DateIn, P1, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

'------------------------------------------------------------------------------
' TRAVERSE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar: the 1x1 case of the element
        If OutKind = KPR_SHAPE_SCALAR Then
            KPR_Dates_EndOfMonth = Elem_EndOfMonth(P1)
            Exit Function
        End If

    'A multi-element call returns a 1-based 2-D array of the resolved shape,
    'evaluated row-major; each element resolves its own inputs
        If Not TryAllocateOutput(OutRows, OutCols, OutArr, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        For R = 1 To OutRows
            For C = 1 To OutCols
                OutArr(R, C) = Elem_EndOfMonth(ElementAt(P1, K1, R, C))
            Next C
        Next R
        KPR_Dates_EndOfMonth = OutArr
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_EndOfMonth = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed with #VALUE!; reaching here is a defect
        FailErr = ErrValue()
        Resume Fail

End Function
Public Function KPR_Dates_BeginOfQuarter( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
'                          KPR_Dates_BeginOfQuarter
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the first date of the containing calendar quarter.
'
' SIGNATURE
'   KPR_Dates_BeginOfQuarter(DateIn) -> Variant (Date)
'
' INPUTS
'   DateIn
'     Date, numeric serial, or ISO YYYY-MM-DD text; single-cell Range or 1x1 wrapper accepted.
'
' RETURNS
'   Variant
'     Date on success, or a native Excel error value.
'
' ERROR POLICY (USER FACING)
'   #VALUE!  a multi-cell optional control; a shape that is not a scalar,
'            Range or array, or a multi-area, empty, jagged or rank-3+ input;
'            two non-scalar arguments of different shapes; a date argument that is not scalar, not a date, or malformed text
'   #NUM!    an output of 100,001 or more elements; a date outside KPR_MIN_DATE .. KPR_MAX_DATE; a result outside the supported window
'   #N/A     an identified 1904 worksheet caller, or an identified caller
'            whose date system cannot be read
'   Incoming native errors are returned unchanged.
'
' SHAPE
'   Value arguments may be a scalar, a single-cell Range or 1x1 wrapper, or a
'   multi-element Range or array. A scalar expands; every non-scalar argument
'   must match exactly in rows and columns. An all-scalar call returns a
'   scalar; any other call returns a 1-based 2-D Variant of the resolved
'   shape, evaluated row-major, with each element resolved independently.
'   Multi-cell results are supported on dynamic-array Excel only.
'
' ORDER OF EVALUATION (call-level, contract section 5.2)
'   host guard -> classify every value argument -> optional controls and
'   broadcast resolution -> element cap -> materialize -> traverse.
'   Within an element, value arguments resolve in signature order.
'
' DEPENDENCIES
'   - PassHostGuard
'   - Elem_BeginOfQuarter
'   - KPR_Core_Array: TryClassifyShape, AccumulateShape, CheckCapacity,
'     TryMaterialize, TryAllocateOutput, ElementAt
'   - KPR_Core_Err.ErrForCondition
'
' NOTES
'   - A Q1-1900 input names 1900-01-01, which is outside the supported window: RESULT_WINDOW, not a date.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P1              As Variant   'Materialized payload of DateIn
    Dim K1              As KPR_ArgShape 'Shape kind of DateIn
    Dim R1              As Long      'Rows of DateIn
    Dim C1              As Long      'Cols of DateIn
    Dim OutKind         As KPR_ArgShape 'Resolved output kind
    Dim OutRows         As Long      'Resolved output rows
    Dim OutCols         As Long      'Resolved output cols
    Dim OutArr          As Variant   'Output array for a multi-element call
    Dim R               As Long      'Row cursor
    Dim C               As Long      'Column cursor
    Dim Condition       As KPR_Condition 'Call-level shape or capacity condition
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' CALL-LEVEL STAGES (contract section 5.2)
'------------------------------------------------------------------------------
    'Stage 1: refuse a 1904 worksheet host before any argument is touched
        If Not PassHostGuard(FailErr) Then GoTo Fail

    'Stage 2: classify every value argument from type and dimensions alone
        If Not TryClassifyShape(DateIn, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 3b: exact-shape broadcast resolution (SHAPE_MISMATCH)
        OutKind = KPR_SHAPE_SCALAR: OutRows = 1: OutCols = 1
        If Not AccumulateShape(OutKind, OutRows, OutCols, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 4: the element cap, once, on the resolved output shape
        If Not CheckCapacity(OutRows, OutCols, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 5: only now may content be read
        If Not TryMaterialize(DateIn, P1, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

'------------------------------------------------------------------------------
' TRAVERSE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar: the 1x1 case of the element
        If OutKind = KPR_SHAPE_SCALAR Then
            KPR_Dates_BeginOfQuarter = Elem_BeginOfQuarter(P1)
            Exit Function
        End If

    'A multi-element call returns a 1-based 2-D array of the resolved shape,
    'evaluated row-major; each element resolves its own inputs
        If Not TryAllocateOutput(OutRows, OutCols, OutArr, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        For R = 1 To OutRows
            For C = 1 To OutCols
                OutArr(R, C) = Elem_BeginOfQuarter(ElementAt(P1, K1, R, C))
            Next C
        Next R
        KPR_Dates_BeginOfQuarter = OutArr
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_BeginOfQuarter = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed with #VALUE!; reaching here is a defect
        FailErr = ErrValue()
        Resume Fail

End Function
Public Function KPR_Dates_EndOfQuarter( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
'                          KPR_Dates_EndOfQuarter
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the last date of the containing calendar quarter.
'
' SIGNATURE
'   KPR_Dates_EndOfQuarter(DateIn) -> Variant (Date)
'
' INPUTS
'   DateIn
'     Date, numeric serial, or ISO YYYY-MM-DD text; single-cell Range or 1x1 wrapper accepted.
'
' RETURNS
'   Variant
'     Date on success, or a native Excel error value.
'
' ERROR POLICY (USER FACING)
'   #VALUE!  a multi-cell optional control; a shape that is not a scalar,
'            Range or array, or a multi-area, empty, jagged or rank-3+ input;
'            two non-scalar arguments of different shapes; a date argument that is not scalar, not a date, or malformed text
'   #NUM!    an output of 100,001 or more elements; a date outside KPR_MIN_DATE .. KPR_MAX_DATE; a result outside the supported window
'   #N/A     an identified 1904 worksheet caller, or an identified caller
'            whose date system cannot be read
'   Incoming native errors are returned unchanged.
'
' SHAPE
'   Value arguments may be a scalar, a single-cell Range or 1x1 wrapper, or a
'   multi-element Range or array. A scalar expands; every non-scalar argument
'   must match exactly in rows and columns. An all-scalar call returns a
'   scalar; any other call returns a 1-based 2-D Variant of the resolved
'   shape, evaluated row-major, with each element resolved independently.
'   Multi-cell results are supported on dynamic-array Excel only.
'
' ORDER OF EVALUATION (call-level, contract section 5.2)
'   host guard -> classify every value argument -> optional controls and
'   broadcast resolution -> element cap -> materialize -> traverse.
'   Within an element, value arguments resolve in signature order.
'
' DEPENDENCIES
'   - PassHostGuard
'   - Elem_EndOfQuarter
'   - KPR_Core_Array: TryClassifyShape, AccumulateShape, CheckCapacity,
'     TryMaterialize, TryAllocateOutput, ElementAt
'   - KPR_Core_Err.ErrForCondition
'
' NOTES
'   - Always inside the window for an in-window input: the quarter end is never earlier than the input.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P1              As Variant   'Materialized payload of DateIn
    Dim K1              As KPR_ArgShape 'Shape kind of DateIn
    Dim R1              As Long      'Rows of DateIn
    Dim C1              As Long      'Cols of DateIn
    Dim OutKind         As KPR_ArgShape 'Resolved output kind
    Dim OutRows         As Long      'Resolved output rows
    Dim OutCols         As Long      'Resolved output cols
    Dim OutArr          As Variant   'Output array for a multi-element call
    Dim R               As Long      'Row cursor
    Dim C               As Long      'Column cursor
    Dim Condition       As KPR_Condition 'Call-level shape or capacity condition
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' CALL-LEVEL STAGES (contract section 5.2)
'------------------------------------------------------------------------------
    'Stage 1: refuse a 1904 worksheet host before any argument is touched
        If Not PassHostGuard(FailErr) Then GoTo Fail

    'Stage 2: classify every value argument from type and dimensions alone
        If Not TryClassifyShape(DateIn, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 3b: exact-shape broadcast resolution (SHAPE_MISMATCH)
        OutKind = KPR_SHAPE_SCALAR: OutRows = 1: OutCols = 1
        If Not AccumulateShape(OutKind, OutRows, OutCols, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 4: the element cap, once, on the resolved output shape
        If Not CheckCapacity(OutRows, OutCols, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 5: only now may content be read
        If Not TryMaterialize(DateIn, P1, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

'------------------------------------------------------------------------------
' TRAVERSE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar: the 1x1 case of the element
        If OutKind = KPR_SHAPE_SCALAR Then
            KPR_Dates_EndOfQuarter = Elem_EndOfQuarter(P1)
            Exit Function
        End If

    'A multi-element call returns a 1-based 2-D array of the resolved shape,
    'evaluated row-major; each element resolves its own inputs
        If Not TryAllocateOutput(OutRows, OutCols, OutArr, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        For R = 1 To OutRows
            For C = 1 To OutCols
                OutArr(R, C) = Elem_EndOfQuarter(ElementAt(P1, K1, R, C))
            Next C
        Next R
        KPR_Dates_EndOfQuarter = OutArr
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_EndOfQuarter = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed with #VALUE!; reaching here is a defect
        FailErr = ErrValue()
        Resume Fail

End Function
Public Function KPR_Dates_BeginOfYear( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
'                          KPR_Dates_BeginOfYear
'------------------------------------------------------------------------------
' PURPOSE
'   Returns 1 January of the containing year.
'
' SIGNATURE
'   KPR_Dates_BeginOfYear(DateIn) -> Variant (Date)
'
' INPUTS
'   DateIn
'     Date, numeric serial, or ISO YYYY-MM-DD text; single-cell Range or 1x1 wrapper accepted.
'
' RETURNS
'   Variant
'     Date on success, or a native Excel error value.
'
' ERROR POLICY (USER FACING)
'   #VALUE!  a multi-cell optional control; a shape that is not a scalar,
'            Range or array, or a multi-area, empty, jagged or rank-3+ input;
'            two non-scalar arguments of different shapes; a date argument that is not scalar, not a date, or malformed text
'   #NUM!    an output of 100,001 or more elements; a date outside KPR_MIN_DATE .. KPR_MAX_DATE; a result outside the supported window
'   #N/A     an identified 1904 worksheet caller, or an identified caller
'            whose date system cannot be read
'   Incoming native errors are returned unchanged.
'
' SHAPE
'   Value arguments may be a scalar, a single-cell Range or 1x1 wrapper, or a
'   multi-element Range or array. A scalar expands; every non-scalar argument
'   must match exactly in rows and columns. An all-scalar call returns a
'   scalar; any other call returns a 1-based 2-D Variant of the resolved
'   shape, evaluated row-major, with each element resolved independently.
'   Multi-cell results are supported on dynamic-array Excel only.
'
' ORDER OF EVALUATION (call-level, contract section 5.2)
'   host guard -> classify every value argument -> optional controls and
'   broadcast resolution -> element cap -> materialize -> traverse.
'   Within an element, value arguments resolve in signature order.
'
' DEPENDENCIES
'   - PassHostGuard
'   - Elem_BeginOfYear
'   - KPR_Core_Array: TryClassifyShape, AccumulateShape, CheckCapacity,
'     TryMaterialize, TryAllocateOutput, ElementAt
'   - KPR_Core_Err.ErrForCondition
'
' NOTES
'   - Any 1900 input names 1900-01-01, which is outside the supported window: RESULT_WINDOW, not a date.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P1              As Variant   'Materialized payload of DateIn
    Dim K1              As KPR_ArgShape 'Shape kind of DateIn
    Dim R1              As Long      'Rows of DateIn
    Dim C1              As Long      'Cols of DateIn
    Dim OutKind         As KPR_ArgShape 'Resolved output kind
    Dim OutRows         As Long      'Resolved output rows
    Dim OutCols         As Long      'Resolved output cols
    Dim OutArr          As Variant   'Output array for a multi-element call
    Dim R               As Long      'Row cursor
    Dim C               As Long      'Column cursor
    Dim Condition       As KPR_Condition 'Call-level shape or capacity condition
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' CALL-LEVEL STAGES (contract section 5.2)
'------------------------------------------------------------------------------
    'Stage 1: refuse a 1904 worksheet host before any argument is touched
        If Not PassHostGuard(FailErr) Then GoTo Fail

    'Stage 2: classify every value argument from type and dimensions alone
        If Not TryClassifyShape(DateIn, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 3b: exact-shape broadcast resolution (SHAPE_MISMATCH)
        OutKind = KPR_SHAPE_SCALAR: OutRows = 1: OutCols = 1
        If Not AccumulateShape(OutKind, OutRows, OutCols, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 4: the element cap, once, on the resolved output shape
        If Not CheckCapacity(OutRows, OutCols, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 5: only now may content be read
        If Not TryMaterialize(DateIn, P1, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

'------------------------------------------------------------------------------
' TRAVERSE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar: the 1x1 case of the element
        If OutKind = KPR_SHAPE_SCALAR Then
            KPR_Dates_BeginOfYear = Elem_BeginOfYear(P1)
            Exit Function
        End If

    'A multi-element call returns a 1-based 2-D array of the resolved shape,
    'evaluated row-major; each element resolves its own inputs
        If Not TryAllocateOutput(OutRows, OutCols, OutArr, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        For R = 1 To OutRows
            For C = 1 To OutCols
                OutArr(R, C) = Elem_BeginOfYear(ElementAt(P1, K1, R, C))
            Next C
        Next R
        KPR_Dates_BeginOfYear = OutArr
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_BeginOfYear = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed with #VALUE!; reaching here is a defect
        FailErr = ErrValue()
        Resume Fail

End Function
Public Function KPR_Dates_EndOfYear( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
'                          KPR_Dates_EndOfYear
'------------------------------------------------------------------------------
' PURPOSE
'   Returns 31 December of the containing year.
'
' SIGNATURE
'   KPR_Dates_EndOfYear(DateIn) -> Variant (Date)
'
' INPUTS
'   DateIn
'     Date, numeric serial, or ISO YYYY-MM-DD text; single-cell Range or 1x1 wrapper accepted.
'
' RETURNS
'   Variant
'     Date on success, or a native Excel error value.
'
' ERROR POLICY (USER FACING)
'   #VALUE!  a multi-cell optional control; a shape that is not a scalar,
'            Range or array, or a multi-area, empty, jagged or rank-3+ input;
'            two non-scalar arguments of different shapes; a date argument that is not scalar, not a date, or malformed text
'   #NUM!    an output of 100,001 or more elements; a date outside KPR_MIN_DATE .. KPR_MAX_DATE; a result outside the supported window
'   #N/A     an identified 1904 worksheet caller, or an identified caller
'            whose date system cannot be read
'   Incoming native errors are returned unchanged.
'
' SHAPE
'   Value arguments may be a scalar, a single-cell Range or 1x1 wrapper, or a
'   multi-element Range or array. A scalar expands; every non-scalar argument
'   must match exactly in rows and columns. An all-scalar call returns a
'   scalar; any other call returns a 1-based 2-D Variant of the resolved
'   shape, evaluated row-major, with each element resolved independently.
'   Multi-cell results are supported on dynamic-array Excel only.
'
' ORDER OF EVALUATION (call-level, contract section 5.2)
'   host guard -> classify every value argument -> optional controls and
'   broadcast resolution -> element cap -> materialize -> traverse.
'   Within an element, value arguments resolve in signature order.
'
' DEPENDENCIES
'   - PassHostGuard
'   - Elem_EndOfYear
'   - KPR_Core_Array: TryClassifyShape, AccumulateShape, CheckCapacity,
'     TryMaterialize, TryAllocateOutput, ElementAt
'   - KPR_Core_Err.ErrForCondition
'
' NOTES
'   - None.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P1              As Variant   'Materialized payload of DateIn
    Dim K1              As KPR_ArgShape 'Shape kind of DateIn
    Dim R1              As Long      'Rows of DateIn
    Dim C1              As Long      'Cols of DateIn
    Dim OutKind         As KPR_ArgShape 'Resolved output kind
    Dim OutRows         As Long      'Resolved output rows
    Dim OutCols         As Long      'Resolved output cols
    Dim OutArr          As Variant   'Output array for a multi-element call
    Dim R               As Long      'Row cursor
    Dim C               As Long      'Column cursor
    Dim Condition       As KPR_Condition 'Call-level shape or capacity condition
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' CALL-LEVEL STAGES (contract section 5.2)
'------------------------------------------------------------------------------
    'Stage 1: refuse a 1904 worksheet host before any argument is touched
        If Not PassHostGuard(FailErr) Then GoTo Fail

    'Stage 2: classify every value argument from type and dimensions alone
        If Not TryClassifyShape(DateIn, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 3b: exact-shape broadcast resolution (SHAPE_MISMATCH)
        OutKind = KPR_SHAPE_SCALAR: OutRows = 1: OutCols = 1
        If Not AccumulateShape(OutKind, OutRows, OutCols, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 4: the element cap, once, on the resolved output shape
        If Not CheckCapacity(OutRows, OutCols, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 5: only now may content be read
        If Not TryMaterialize(DateIn, P1, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

'------------------------------------------------------------------------------
' TRAVERSE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar: the 1x1 case of the element
        If OutKind = KPR_SHAPE_SCALAR Then
            KPR_Dates_EndOfYear = Elem_EndOfYear(P1)
            Exit Function
        End If

    'A multi-element call returns a 1-based 2-D array of the resolved shape,
    'evaluated row-major; each element resolves its own inputs
        If Not TryAllocateOutput(OutRows, OutCols, OutArr, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        For R = 1 To OutRows
            For C = 1 To OutCols
                OutArr(R, C) = Elem_EndOfYear(ElementAt(P1, K1, R, C))
            Next C
        Next R
        KPR_Dates_EndOfYear = OutArr
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_EndOfYear = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed with #VALUE!; reaching here is a defect
        FailErr = ErrValue()
        Resume Fail

End Function
'
'------------------------------------------------------------------------------
'
'                       PUBLIC API - BOUNDARY PREDICATES                       
'
'------------------------------------------------------------------------------
'

Public Function KPR_Dates_IsMonthEnd( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
'                          KPR_Dates_IsMonthEnd
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a date is the last day of its month.
'
' SIGNATURE
'   KPR_Dates_IsMonthEnd(DateIn) -> Variant (Boolean)
'
' INPUTS
'   DateIn
'     Date, numeric serial, or ISO YYYY-MM-DD text; single-cell Range or 1x1 wrapper accepted.
'
' RETURNS
'   Variant
'     Boolean on success, or a native Excel error value.
'
' ERROR POLICY (USER FACING)
'   #VALUE!  a multi-cell optional control; a shape that is not a scalar,
'            Range or array, or a multi-area, empty, jagged or rank-3+ input;
'            two non-scalar arguments of different shapes; a date argument that is not scalar, not a date, or malformed text
'   #NUM!    an output of 100,001 or more elements; a date outside KPR_MIN_DATE .. KPR_MAX_DATE
'   #N/A     an identified 1904 worksheet caller, or an identified caller
'            whose date system cannot be read
'   Incoming native errors are returned unchanged.
'
' SHAPE
'   Value arguments may be a scalar, a single-cell Range or 1x1 wrapper, or a
'   multi-element Range or array. A scalar expands; every non-scalar argument
'   must match exactly in rows and columns. An all-scalar call returns a
'   scalar; any other call returns a 1-based 2-D Variant of the resolved
'   shape, evaluated row-major, with each element resolved independently.
'   Multi-cell results are supported on dynamic-array Excel only.
'
' ORDER OF EVALUATION (call-level, contract section 5.2)
'   host guard -> classify every value argument -> optional controls and
'   broadcast resolution -> element cap -> materialize -> traverse.
'   Within an element, value arguments resolve in signature order.
'
' DEPENDENCIES
'   - PassHostGuard
'   - Elem_IsMonthEnd
'   - KPR_Core_Array: TryClassifyShape, AccumulateShape, CheckCapacity,
'     TryMaterialize, TryAllocateOutput, ElementAt
'   - KPR_Core_Err.ErrForCondition
'
' NOTES
'   - None.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P1              As Variant   'Materialized payload of DateIn
    Dim K1              As KPR_ArgShape 'Shape kind of DateIn
    Dim R1              As Long      'Rows of DateIn
    Dim C1              As Long      'Cols of DateIn
    Dim OutKind         As KPR_ArgShape 'Resolved output kind
    Dim OutRows         As Long      'Resolved output rows
    Dim OutCols         As Long      'Resolved output cols
    Dim OutArr          As Variant   'Output array for a multi-element call
    Dim R               As Long      'Row cursor
    Dim C               As Long      'Column cursor
    Dim Condition       As KPR_Condition 'Call-level shape or capacity condition
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' CALL-LEVEL STAGES (contract section 5.2)
'------------------------------------------------------------------------------
    'Stage 1: refuse a 1904 worksheet host before any argument is touched
        If Not PassHostGuard(FailErr) Then GoTo Fail

    'Stage 2: classify every value argument from type and dimensions alone
        If Not TryClassifyShape(DateIn, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 3b: exact-shape broadcast resolution (SHAPE_MISMATCH)
        OutKind = KPR_SHAPE_SCALAR: OutRows = 1: OutCols = 1
        If Not AccumulateShape(OutKind, OutRows, OutCols, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 4: the element cap, once, on the resolved output shape
        If Not CheckCapacity(OutRows, OutCols, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 5: only now may content be read
        If Not TryMaterialize(DateIn, P1, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

'------------------------------------------------------------------------------
' TRAVERSE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar: the 1x1 case of the element
        If OutKind = KPR_SHAPE_SCALAR Then
            KPR_Dates_IsMonthEnd = Elem_IsMonthEnd(P1)
            Exit Function
        End If

    'A multi-element call returns a 1-based 2-D array of the resolved shape,
    'evaluated row-major; each element resolves its own inputs
        If Not TryAllocateOutput(OutRows, OutCols, OutArr, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        For R = 1 To OutRows
            For C = 1 To OutCols
                OutArr(R, C) = Elem_IsMonthEnd(ElementAt(P1, K1, R, C))
            Next C
        Next R
        KPR_Dates_IsMonthEnd = OutArr
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_IsMonthEnd = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed with #VALUE!; reaching here is a defect
        FailErr = ErrValue()
        Resume Fail

End Function
Public Function KPR_Dates_IsQuarterEnd( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
'                          KPR_Dates_IsQuarterEnd
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a date is the last day of a calendar quarter.
'
' SIGNATURE
'   KPR_Dates_IsQuarterEnd(DateIn) -> Variant (Boolean)
'
' INPUTS
'   DateIn
'     Date, numeric serial, or ISO YYYY-MM-DD text; single-cell Range or 1x1 wrapper accepted.
'
' RETURNS
'   Variant
'     Boolean on success, or a native Excel error value.
'
' ERROR POLICY (USER FACING)
'   #VALUE!  a multi-cell optional control; a shape that is not a scalar,
'            Range or array, or a multi-area, empty, jagged or rank-3+ input;
'            two non-scalar arguments of different shapes; a date argument that is not scalar, not a date, or malformed text
'   #NUM!    an output of 100,001 or more elements; a date outside KPR_MIN_DATE .. KPR_MAX_DATE
'   #N/A     an identified 1904 worksheet caller, or an identified caller
'            whose date system cannot be read
'   Incoming native errors are returned unchanged.
'
' SHAPE
'   Value arguments may be a scalar, a single-cell Range or 1x1 wrapper, or a
'   multi-element Range or array. A scalar expands; every non-scalar argument
'   must match exactly in rows and columns. An all-scalar call returns a
'   scalar; any other call returns a 1-based 2-D Variant of the resolved
'   shape, evaluated row-major, with each element resolved independently.
'   Multi-cell results are supported on dynamic-array Excel only.
'
' ORDER OF EVALUATION (call-level, contract section 5.2)
'   host guard -> classify every value argument -> optional controls and
'   broadcast resolution -> element cap -> materialize -> traverse.
'   Within an element, value arguments resolve in signature order.
'
' DEPENDENCIES
'   - PassHostGuard
'   - Elem_IsQuarterEnd
'   - KPR_Core_Array: TryClassifyShape, AccumulateShape, CheckCapacity,
'     TryMaterialize, TryAllocateOutput, ElementAt
'   - KPR_Core_Err.ErrForCondition
'
' NOTES
'   - None.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P1              As Variant   'Materialized payload of DateIn
    Dim K1              As KPR_ArgShape 'Shape kind of DateIn
    Dim R1              As Long      'Rows of DateIn
    Dim C1              As Long      'Cols of DateIn
    Dim OutKind         As KPR_ArgShape 'Resolved output kind
    Dim OutRows         As Long      'Resolved output rows
    Dim OutCols         As Long      'Resolved output cols
    Dim OutArr          As Variant   'Output array for a multi-element call
    Dim R               As Long      'Row cursor
    Dim C               As Long      'Column cursor
    Dim Condition       As KPR_Condition 'Call-level shape or capacity condition
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' CALL-LEVEL STAGES (contract section 5.2)
'------------------------------------------------------------------------------
    'Stage 1: refuse a 1904 worksheet host before any argument is touched
        If Not PassHostGuard(FailErr) Then GoTo Fail

    'Stage 2: classify every value argument from type and dimensions alone
        If Not TryClassifyShape(DateIn, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 3b: exact-shape broadcast resolution (SHAPE_MISMATCH)
        OutKind = KPR_SHAPE_SCALAR: OutRows = 1: OutCols = 1
        If Not AccumulateShape(OutKind, OutRows, OutCols, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 4: the element cap, once, on the resolved output shape
        If Not CheckCapacity(OutRows, OutCols, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 5: only now may content be read
        If Not TryMaterialize(DateIn, P1, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

'------------------------------------------------------------------------------
' TRAVERSE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar: the 1x1 case of the element
        If OutKind = KPR_SHAPE_SCALAR Then
            KPR_Dates_IsQuarterEnd = Elem_IsQuarterEnd(P1)
            Exit Function
        End If

    'A multi-element call returns a 1-based 2-D array of the resolved shape,
    'evaluated row-major; each element resolves its own inputs
        If Not TryAllocateOutput(OutRows, OutCols, OutArr, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        For R = 1 To OutRows
            For C = 1 To OutCols
                OutArr(R, C) = Elem_IsQuarterEnd(ElementAt(P1, K1, R, C))
            Next C
        Next R
        KPR_Dates_IsQuarterEnd = OutArr
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_IsQuarterEnd = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed with #VALUE!; reaching here is a defect
        FailErr = ErrValue()
        Resume Fail

End Function
Public Function KPR_Dates_IsYearEnd( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
'                          KPR_Dates_IsYearEnd
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a date is 31 December.
'
' SIGNATURE
'   KPR_Dates_IsYearEnd(DateIn) -> Variant (Boolean)
'
' INPUTS
'   DateIn
'     Date, numeric serial, or ISO YYYY-MM-DD text; single-cell Range or 1x1 wrapper accepted.
'
' RETURNS
'   Variant
'     Boolean on success, or a native Excel error value.
'
' ERROR POLICY (USER FACING)
'   #VALUE!  a multi-cell optional control; a shape that is not a scalar,
'            Range or array, or a multi-area, empty, jagged or rank-3+ input;
'            two non-scalar arguments of different shapes; a date argument that is not scalar, not a date, or malformed text
'   #NUM!    an output of 100,001 or more elements; a date outside KPR_MIN_DATE .. KPR_MAX_DATE
'   #N/A     an identified 1904 worksheet caller, or an identified caller
'            whose date system cannot be read
'   Incoming native errors are returned unchanged.
'
' SHAPE
'   Value arguments may be a scalar, a single-cell Range or 1x1 wrapper, or a
'   multi-element Range or array. A scalar expands; every non-scalar argument
'   must match exactly in rows and columns. An all-scalar call returns a
'   scalar; any other call returns a 1-based 2-D Variant of the resolved
'   shape, evaluated row-major, with each element resolved independently.
'   Multi-cell results are supported on dynamic-array Excel only.
'
' ORDER OF EVALUATION (call-level, contract section 5.2)
'   host guard -> classify every value argument -> optional controls and
'   broadcast resolution -> element cap -> materialize -> traverse.
'   Within an element, value arguments resolve in signature order.
'
' DEPENDENCIES
'   - PassHostGuard
'   - Elem_IsYearEnd
'   - KPR_Core_Array: TryClassifyShape, AccumulateShape, CheckCapacity,
'     TryMaterialize, TryAllocateOutput, ElementAt
'   - KPR_Core_Err.ErrForCondition
'
' NOTES
'   - None.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P1              As Variant   'Materialized payload of DateIn
    Dim K1              As KPR_ArgShape 'Shape kind of DateIn
    Dim R1              As Long      'Rows of DateIn
    Dim C1              As Long      'Cols of DateIn
    Dim OutKind         As KPR_ArgShape 'Resolved output kind
    Dim OutRows         As Long      'Resolved output rows
    Dim OutCols         As Long      'Resolved output cols
    Dim OutArr          As Variant   'Output array for a multi-element call
    Dim R               As Long      'Row cursor
    Dim C               As Long      'Column cursor
    Dim Condition       As KPR_Condition 'Call-level shape or capacity condition
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' CALL-LEVEL STAGES (contract section 5.2)
'------------------------------------------------------------------------------
    'Stage 1: refuse a 1904 worksheet host before any argument is touched
        If Not PassHostGuard(FailErr) Then GoTo Fail

    'Stage 2: classify every value argument from type and dimensions alone
        If Not TryClassifyShape(DateIn, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 3b: exact-shape broadcast resolution (SHAPE_MISMATCH)
        OutKind = KPR_SHAPE_SCALAR: OutRows = 1: OutCols = 1
        If Not AccumulateShape(OutKind, OutRows, OutCols, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 4: the element cap, once, on the resolved output shape
        If Not CheckCapacity(OutRows, OutCols, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 5: only now may content be read
        If Not TryMaterialize(DateIn, P1, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

'------------------------------------------------------------------------------
' TRAVERSE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar: the 1x1 case of the element
        If OutKind = KPR_SHAPE_SCALAR Then
            KPR_Dates_IsYearEnd = Elem_IsYearEnd(P1)
            Exit Function
        End If

    'A multi-element call returns a 1-based 2-D array of the resolved shape,
    'evaluated row-major; each element resolves its own inputs
        If Not TryAllocateOutput(OutRows, OutCols, OutArr, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        For R = 1 To OutRows
            For C = 1 To OutCols
                OutArr(R, C) = Elem_IsYearEnd(ElementAt(P1, K1, R, C))
            Next C
        Next R
        KPR_Dates_IsYearEnd = OutArr
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_IsYearEnd = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed with #VALUE!; reaching here is a defect
        FailErr = ErrValue()
        Resume Fail

End Function
Public Function KPR_Dates_IsLeapYear( _
    ByVal YearIn As Variant) _
    As Variant
'
'==============================================================================
'                          KPR_Dates_IsLeapYear
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a calendar year is a Gregorian leap year.
'
' SIGNATURE
'   KPR_Dates_IsLeapYear(YearIn) -> Variant (Boolean)
'
' INPUTS
'   YearIn
'     Native integral numeric, domain 1900 through 9999.
'
' RETURNS
'   Variant
'     Boolean on success, or a native Excel error value.
'
' ERROR POLICY (USER FACING)
'   #VALUE!  a multi-cell optional control; a shape that is not a scalar,
'            Range or array, or a multi-area, empty, jagged or rank-3+ input;
'            two non-scalar arguments of different shapes; an integer argument that is fractional, Boolean, text, or outside its domain
'   #NUM!    an output of 100,001 or more elements; an integer outside the Long range
'   #N/A     an identified 1904 worksheet caller, or an identified caller
'            whose date system cannot be read
'   Incoming native errors are returned unchanged.
'
' SHAPE
'   Value arguments may be a scalar, a single-cell Range or 1x1 wrapper, or a
'   multi-element Range or array. A scalar expands; every non-scalar argument
'   must match exactly in rows and columns. An all-scalar call returns a
'   scalar; any other call returns a 1-based 2-D Variant of the resolved
'   shape, evaluated row-major, with each element resolved independently.
'   Multi-cell results are supported on dynamic-array Excel only.
'
' ORDER OF EVALUATION (call-level, contract section 5.2)
'   host guard -> classify every value argument -> optional controls and
'   broadcast resolution -> element cap -> materialize -> traverse.
'   Within an element, value arguments resolve in signature order.
'
' DEPENDENCIES
'   - PassHostGuard
'   - Elem_IsLeapYear
'   - KPR_Core_Array: TryClassifyShape, AccumulateShape, CheckCapacity,
'     TryMaterialize, TryAllocateOutput, ElementAt
'   - KPR_Core_Err.ErrForCondition
'
' NOTES
'   - Takes a calendar year, not a date. IsLeapYear(1900) is False; no window gate applies.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P1              As Variant   'Materialized payload of YearIn
    Dim K1              As KPR_ArgShape 'Shape kind of YearIn
    Dim R1              As Long      'Rows of YearIn
    Dim C1              As Long      'Cols of YearIn
    Dim OutKind         As KPR_ArgShape 'Resolved output kind
    Dim OutRows         As Long      'Resolved output rows
    Dim OutCols         As Long      'Resolved output cols
    Dim OutArr          As Variant   'Output array for a multi-element call
    Dim R               As Long      'Row cursor
    Dim C               As Long      'Column cursor
    Dim Condition       As KPR_Condition 'Call-level shape or capacity condition
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' CALL-LEVEL STAGES (contract section 5.2)
'------------------------------------------------------------------------------
    'Stage 1: refuse a 1904 worksheet host before any argument is touched
        If Not PassHostGuard(FailErr) Then GoTo Fail

    'Stage 2: classify every value argument from type and dimensions alone
        If Not TryClassifyShape(YearIn, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 3b: exact-shape broadcast resolution (SHAPE_MISMATCH)
        OutKind = KPR_SHAPE_SCALAR: OutRows = 1: OutCols = 1
        If Not AccumulateShape(OutKind, OutRows, OutCols, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 4: the element cap, once, on the resolved output shape
        If Not CheckCapacity(OutRows, OutCols, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 5: only now may content be read
        If Not TryMaterialize(YearIn, P1, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

'------------------------------------------------------------------------------
' TRAVERSE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar: the 1x1 case of the element
        If OutKind = KPR_SHAPE_SCALAR Then
            KPR_Dates_IsLeapYear = Elem_IsLeapYear(P1)
            Exit Function
        End If

    'A multi-element call returns a 1-based 2-D array of the resolved shape,
    'evaluated row-major; each element resolves its own inputs
        If Not TryAllocateOutput(OutRows, OutCols, OutArr, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        For R = 1 To OutRows
            For C = 1 To OutCols
                OutArr(R, C) = Elem_IsLeapYear(ElementAt(P1, K1, R, C))
            Next C
        Next R
        KPR_Dates_IsLeapYear = OutArr
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_IsLeapYear = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed with #VALUE!; reaching here is a defect
        FailErr = ErrValue()
        Resume Fail

End Function
'
'------------------------------------------------------------------------------
'
'                         PUBLIC API - DATE ARITHMETIC                         
'
'------------------------------------------------------------------------------
'

Public Function KPR_Dates_AddDays( _
    ByVal DateIn As Variant, _
    ByVal nDays As Variant) _
    As Variant
'
'==============================================================================
'                          KPR_Dates_AddDays
'------------------------------------------------------------------------------
' PURPOSE
'   Shifts a date by an exact number of calendar days.
'
' SIGNATURE
'   KPR_Dates_AddDays(DateIn, nDays) -> Variant (Date)
'
' INPUTS
'   DateIn
'     Date, numeric serial, or ISO YYYY-MM-DD text; single-cell Range or 1x1 wrapper accepted.
'   nDays
'     Native integral numeric, domain the Long range.
'
' RETURNS
'   Variant
'     Date on success, or a native Excel error value.
'
' ERROR POLICY (USER FACING)
'   #VALUE!  a multi-cell optional control; a shape that is not a scalar,
'            Range or array, or a multi-area, empty, jagged or rank-3+ input;
'            two non-scalar arguments of different shapes; a date argument that is not scalar, not a date, or malformed text; an integer argument that is fractional, Boolean, text, or outside its domain
'   #NUM!    an output of 100,001 or more elements; a date outside KPR_MIN_DATE .. KPR_MAX_DATE; an integer outside the Long range; a result outside the supported window
'   #N/A     an identified 1904 worksheet caller, or an identified caller
'            whose date system cannot be read
'   Incoming native errors are returned unchanged.
'
' SHAPE
'   Value arguments may be a scalar, a single-cell Range or 1x1 wrapper, or a
'   multi-element Range or array. A scalar expands; every non-scalar argument
'   must match exactly in rows and columns. An all-scalar call returns a
'   scalar; any other call returns a 1-based 2-D Variant of the resolved
'   shape, evaluated row-major, with each element resolved independently.
'   Multi-cell results are supported on dynamic-array Excel only.
'
' ORDER OF EVALUATION (call-level, contract section 5.2)
'   host guard -> classify every value argument -> optional controls and
'   broadcast resolution -> element cap -> materialize -> traverse.
'   Within an element, value arguments resolve in signature order.
'
' DEPENDENCIES
'   - PassHostGuard
'   - Elem_AddDays
'   - KPR_Core_Array: TryClassifyShape, AccumulateShape, CheckCapacity,
'     TryMaterialize, TryAllocateOutput, ElementAt
'   - KPR_Core_Err.ErrForCondition
'
' NOTES
'   - The shift is computed in Double and gated before coercion back to Date, so an overflow can never reach CDate.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P1              As Variant   'Materialized payload of DateIn
    Dim K1              As KPR_ArgShape 'Shape kind of DateIn
    Dim R1              As Long      'Rows of DateIn
    Dim C1              As Long      'Cols of DateIn
    Dim P2              As Variant   'Materialized payload of nDays
    Dim K2              As KPR_ArgShape 'Shape kind of nDays
    Dim R2              As Long      'Rows of nDays
    Dim C2              As Long      'Cols of nDays
    Dim OutKind         As KPR_ArgShape 'Resolved output kind
    Dim OutRows         As Long      'Resolved output rows
    Dim OutCols         As Long      'Resolved output cols
    Dim OutArr          As Variant   'Output array for a multi-element call
    Dim R               As Long      'Row cursor
    Dim C               As Long      'Column cursor
    Dim Condition       As KPR_Condition 'Call-level shape or capacity condition
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' CALL-LEVEL STAGES (contract section 5.2)
'------------------------------------------------------------------------------
    'Stage 1: refuse a 1904 worksheet host before any argument is touched
        If Not PassHostGuard(FailErr) Then GoTo Fail

    'Stage 2: classify every value argument from type and dimensions alone
        If Not TryClassifyShape(DateIn, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryClassifyShape(nDays, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 3b: exact-shape broadcast resolution (SHAPE_MISMATCH)
        OutKind = KPR_SHAPE_SCALAR: OutRows = 1: OutCols = 1
        If Not AccumulateShape(OutKind, OutRows, OutCols, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not AccumulateShape(OutKind, OutRows, OutCols, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 4: the element cap, once, on the resolved output shape
        If Not CheckCapacity(OutRows, OutCols, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 5: only now may content be read
        If Not TryMaterialize(DateIn, P1, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryMaterialize(nDays, P2, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

'------------------------------------------------------------------------------
' TRAVERSE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar: the 1x1 case of the element
        If OutKind = KPR_SHAPE_SCALAR Then
            KPR_Dates_AddDays = Elem_AddDays(P1, P2)
            Exit Function
        End If

    'A multi-element call returns a 1-based 2-D array of the resolved shape,
    'evaluated row-major; each element resolves its own inputs
        If Not TryAllocateOutput(OutRows, OutCols, OutArr, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        For R = 1 To OutRows
            For C = 1 To OutCols
                OutArr(R, C) = Elem_AddDays(ElementAt(P1, K1, R, C), ElementAt(P2, K2, R, C))
            Next C
        Next R
        KPR_Dates_AddDays = OutArr
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_AddDays = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed with #VALUE!; reaching here is a defect
        FailErr = ErrValue()
        Resume Fail

End Function
Public Function KPR_Dates_AddWeeks( _
    ByVal DateIn As Variant, _
    ByVal nWeeks As Variant) _
    As Variant
'
'==============================================================================
'                          KPR_Dates_AddWeeks
'------------------------------------------------------------------------------
' PURPOSE
'   Shifts a date by whole weeks.
'
' SIGNATURE
'   KPR_Dates_AddWeeks(DateIn, nWeeks) -> Variant (Date)
'
' INPUTS
'   DateIn
'     Date, numeric serial, or ISO YYYY-MM-DD text; single-cell Range or 1x1 wrapper accepted.
'   nWeeks
'     Native integral numeric, domain the Long range.
'
' RETURNS
'   Variant
'     Date on success, or a native Excel error value.
'
' ERROR POLICY (USER FACING)
'   #VALUE!  a multi-cell optional control; a shape that is not a scalar,
'            Range or array, or a multi-area, empty, jagged or rank-3+ input;
'            two non-scalar arguments of different shapes; a date argument that is not scalar, not a date, or malformed text; an integer argument that is fractional, Boolean, text, or outside its domain
'   #NUM!    an output of 100,001 or more elements; a date outside KPR_MIN_DATE .. KPR_MAX_DATE; an integer outside the Long range; a result outside the supported window
'   #N/A     an identified 1904 worksheet caller, or an identified caller
'            whose date system cannot be read
'   Incoming native errors are returned unchanged.
'
' SHAPE
'   Value arguments may be a scalar, a single-cell Range or 1x1 wrapper, or a
'   multi-element Range or array. A scalar expands; every non-scalar argument
'   must match exactly in rows and columns. An all-scalar call returns a
'   scalar; any other call returns a 1-based 2-D Variant of the resolved
'   shape, evaluated row-major, with each element resolved independently.
'   Multi-cell results are supported on dynamic-array Excel only.
'
' ORDER OF EVALUATION (call-level, contract section 5.2)
'   host guard -> classify every value argument -> optional controls and
'   broadcast resolution -> element cap -> materialize -> traverse.
'   Within an element, value arguments resolve in signature order.
'
' DEPENDENCIES
'   - PassHostGuard
'   - Elem_AddWeeks
'   - KPR_Core_Array: TryClassifyShape, AccumulateShape, CheckCapacity,
'     TryMaterialize, TryAllocateOutput, ElementAt
'   - KPR_Core_Err.ErrForCondition
'
' NOTES
'   - None.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P1              As Variant   'Materialized payload of DateIn
    Dim K1              As KPR_ArgShape 'Shape kind of DateIn
    Dim R1              As Long      'Rows of DateIn
    Dim C1              As Long      'Cols of DateIn
    Dim P2              As Variant   'Materialized payload of nWeeks
    Dim K2              As KPR_ArgShape 'Shape kind of nWeeks
    Dim R2              As Long      'Rows of nWeeks
    Dim C2              As Long      'Cols of nWeeks
    Dim OutKind         As KPR_ArgShape 'Resolved output kind
    Dim OutRows         As Long      'Resolved output rows
    Dim OutCols         As Long      'Resolved output cols
    Dim OutArr          As Variant   'Output array for a multi-element call
    Dim R               As Long      'Row cursor
    Dim C               As Long      'Column cursor
    Dim Condition       As KPR_Condition 'Call-level shape or capacity condition
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' CALL-LEVEL STAGES (contract section 5.2)
'------------------------------------------------------------------------------
    'Stage 1: refuse a 1904 worksheet host before any argument is touched
        If Not PassHostGuard(FailErr) Then GoTo Fail

    'Stage 2: classify every value argument from type and dimensions alone
        If Not TryClassifyShape(DateIn, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryClassifyShape(nWeeks, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 3b: exact-shape broadcast resolution (SHAPE_MISMATCH)
        OutKind = KPR_SHAPE_SCALAR: OutRows = 1: OutCols = 1
        If Not AccumulateShape(OutKind, OutRows, OutCols, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not AccumulateShape(OutKind, OutRows, OutCols, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 4: the element cap, once, on the resolved output shape
        If Not CheckCapacity(OutRows, OutCols, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 5: only now may content be read
        If Not TryMaterialize(DateIn, P1, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryMaterialize(nWeeks, P2, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

'------------------------------------------------------------------------------
' TRAVERSE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar: the 1x1 case of the element
        If OutKind = KPR_SHAPE_SCALAR Then
            KPR_Dates_AddWeeks = Elem_AddWeeks(P1, P2)
            Exit Function
        End If

    'A multi-element call returns a 1-based 2-D array of the resolved shape,
    'evaluated row-major; each element resolves its own inputs
        If Not TryAllocateOutput(OutRows, OutCols, OutArr, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        For R = 1 To OutRows
            For C = 1 To OutCols
                OutArr(R, C) = Elem_AddWeeks(ElementAt(P1, K1, R, C), ElementAt(P2, K2, R, C))
            Next C
        Next R
        KPR_Dates_AddWeeks = OutArr
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_AddWeeks = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed with #VALUE!; reaching here is a defect
        FailErr = ErrValue()
        Resume Fail

End Function
Public Function KPR_Dates_AddMonths( _
    ByVal DateIn As Variant, _
    ByVal nMonths As Variant, _
    Optional ByVal Opt_KeepEOM As Variant = False) _
    As Variant
'
'==============================================================================
'                          KPR_Dates_AddMonths
'------------------------------------------------------------------------------
' PURPOSE
'   Shifts a date by calendar months with clip or EOM-preserving semantics.
'
' SIGNATURE
'   KPR_Dates_AddMonths(DateIn, nMonths, [Opt_KeepEOM]) -> Variant (Date)
'
' INPUTS
'   DateIn
'     Date, numeric serial, or ISO YYYY-MM-DD text; single-cell Range or 1x1 wrapper accepted.
'   nMonths
'     Native integral numeric, domain the Long range.
'   Opt_KeepEOM
'     Optional. Omitted or Empty selects False. Native Boolean only.
'
' RETURNS
'   Variant
'     Date on success, or a native Excel error value.
'
' ERROR POLICY (USER FACING)
'   #VALUE!  a multi-cell optional control; a shape that is not a scalar,
'            Range or array, or a multi-area, empty, jagged or rank-3+ input;
'            two non-scalar arguments of different shapes; a date argument that is not scalar, not a date, or malformed text; an integer argument that is fractional, Boolean, text, or outside its domain; an optional control of the wrong type or an unknown token
'   #NUM!    an output of 100,001 or more elements; a date outside KPR_MIN_DATE .. KPR_MAX_DATE; an integer outside the Long range; a result outside the supported window
'   #N/A     an identified 1904 worksheet caller, or an identified caller
'            whose date system cannot be read
'   Incoming native errors are returned unchanged.
'
' SHAPE
'   Value arguments may be a scalar, a single-cell Range or 1x1 wrapper, or a
'   multi-element Range or array. A scalar expands; every non-scalar argument
'   must match exactly in rows and columns. An all-scalar call returns a
'   scalar; any other call returns a 1-based 2-D Variant of the resolved
'   shape, evaluated row-major, with each element resolved independently.
'   Multi-cell results are supported on dynamic-array Excel only.
'
' ORDER OF EVALUATION (call-level, contract section 5.2)
'   host guard -> classify every value argument -> optional controls and
'   broadcast resolution -> element cap -> materialize -> traverse.
'   Within an element, value arguments resolve in signature order.
'
' DEPENDENCIES
'   - PassHostGuard
'   - Elem_AddMonths
'   - KPR_Core_Array: TryClassifyShape, AccumulateShape, CheckCapacity,
'     TryMaterialize, TryAllocateOutput, ElementAt
'   - KPR_Core_Err.ErrForCondition
'   - TryResolveBool (KPR_Core_Array, KPR_Core_Parse)
'
' NOTES
'   - Clip mode keeps the day when the target month has it and otherwise uses that month's last day. With Opt_KeepEOM=True an EOM input maps to the target EOM; a non-EOM input still clips.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P1              As Variant   'Materialized payload of DateIn
    Dim K1              As KPR_ArgShape 'Shape kind of DateIn
    Dim R1              As Long      'Rows of DateIn
    Dim C1              As Long      'Cols of DateIn
    Dim P2              As Variant   'Materialized payload of nMonths
    Dim K2              As KPR_ArgShape 'Shape kind of nMonths
    Dim R2              As Long      'Rows of nMonths
    Dim C2              As Long      'Cols of nMonths
    Dim KeepEOM         As Boolean    'Resolved Opt_KeepEOM
    Dim OutKind         As KPR_ArgShape 'Resolved output kind
    Dim OutRows         As Long      'Resolved output rows
    Dim OutCols         As Long      'Resolved output cols
    Dim OutArr          As Variant   'Output array for a multi-element call
    Dim R               As Long      'Row cursor
    Dim C               As Long      'Column cursor
    Dim Condition       As KPR_Condition 'Call-level shape or capacity condition
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' CALL-LEVEL STAGES (contract section 5.2)
'------------------------------------------------------------------------------
    'Stage 1: refuse a 1904 worksheet host before any argument is touched
        If Not PassHostGuard(FailErr) Then GoTo Fail

    'Stage 2: classify every value argument from type and dimensions alone
        If Not TryClassifyShape(DateIn, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryClassifyShape(nMonths, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 3a: optional controls, scalar or 1x1 only (CONTROL_NOT_SCALAR)
        If Not TryResolveBool(Opt_KeepEOM, False, KeepEOM, FailErr) Then GoTo Fail

    'Stage 3b: exact-shape broadcast resolution (SHAPE_MISMATCH)
        OutKind = KPR_SHAPE_SCALAR: OutRows = 1: OutCols = 1
        If Not AccumulateShape(OutKind, OutRows, OutCols, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not AccumulateShape(OutKind, OutRows, OutCols, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 4: the element cap, once, on the resolved output shape
        If Not CheckCapacity(OutRows, OutCols, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 5: only now may content be read
        If Not TryMaterialize(DateIn, P1, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryMaterialize(nMonths, P2, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

'------------------------------------------------------------------------------
' TRAVERSE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar: the 1x1 case of the element
        If OutKind = KPR_SHAPE_SCALAR Then
            KPR_Dates_AddMonths = Elem_AddMonths(P1, P2, KeepEOM)
            Exit Function
        End If

    'A multi-element call returns a 1-based 2-D array of the resolved shape,
    'evaluated row-major; each element resolves its own inputs
        If Not TryAllocateOutput(OutRows, OutCols, OutArr, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        For R = 1 To OutRows
            For C = 1 To OutCols
                OutArr(R, C) = Elem_AddMonths(ElementAt(P1, K1, R, C), ElementAt(P2, K2, R, C), KeepEOM)
            Next C
        Next R
        KPR_Dates_AddMonths = OutArr
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_AddMonths = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed with #VALUE!; reaching here is a defect
        FailErr = ErrValue()
        Resume Fail

End Function
Public Function KPR_Dates_AddYears( _
    ByVal DateIn As Variant, _
    ByVal nYears As Variant, _
    Optional ByVal Opt_KeepEOM As Variant = False) _
    As Variant
'
'==============================================================================
'                          KPR_Dates_AddYears
'------------------------------------------------------------------------------
' PURPOSE
'   Shifts a date by calendar years through the single month shifter.
'
' SIGNATURE
'   KPR_Dates_AddYears(DateIn, nYears, [Opt_KeepEOM]) -> Variant (Date)
'
' INPUTS
'   DateIn
'     Date, numeric serial, or ISO YYYY-MM-DD text; single-cell Range or 1x1 wrapper accepted.
'   nYears
'     Native integral numeric, domain the Long range.
'   Opt_KeepEOM
'     Optional. Omitted or Empty selects False. Native Boolean only.
'
' RETURNS
'   Variant
'     Date on success, or a native Excel error value.
'
' ERROR POLICY (USER FACING)
'   #VALUE!  a multi-cell optional control; a shape that is not a scalar,
'            Range or array, or a multi-area, empty, jagged or rank-3+ input;
'            two non-scalar arguments of different shapes; a date argument that is not scalar, not a date, or malformed text; an integer argument that is fractional, Boolean, text, or outside its domain; an optional control of the wrong type or an unknown token
'   #NUM!    an output of 100,001 or more elements; a date outside KPR_MIN_DATE .. KPR_MAX_DATE; an integer outside the Long range; a result outside the supported window
'   #N/A     an identified 1904 worksheet caller, or an identified caller
'            whose date system cannot be read
'   Incoming native errors are returned unchanged.
'
' SHAPE
'   Value arguments may be a scalar, a single-cell Range or 1x1 wrapper, or a
'   multi-element Range or array. A scalar expands; every non-scalar argument
'   must match exactly in rows and columns. An all-scalar call returns a
'   scalar; any other call returns a 1-based 2-D Variant of the resolved
'   shape, evaluated row-major, with each element resolved independently.
'   Multi-cell results are supported on dynamic-array Excel only.
'
' ORDER OF EVALUATION (call-level, contract section 5.2)
'   host guard -> classify every value argument -> optional controls and
'   broadcast resolution -> element cap -> materialize -> traverse.
'   Within an element, value arguments resolve in signature order.
'
' DEPENDENCIES
'   - PassHostGuard
'   - Elem_AddYears
'   - KPR_Core_Array: TryClassifyShape, AccumulateShape, CheckCapacity,
'     TryMaterialize, TryAllocateOutput, ElementAt
'   - KPR_Core_Err.ErrForCondition
'   - TryResolveBool (KPR_Core_Array, KPR_Core_Parse)
'
' NOTES
'   - Delegates to TryAddMonths so 29-Feb, short-month and EOM behaviour cannot diverge from AddMonths.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P1              As Variant   'Materialized payload of DateIn
    Dim K1              As KPR_ArgShape 'Shape kind of DateIn
    Dim R1              As Long      'Rows of DateIn
    Dim C1              As Long      'Cols of DateIn
    Dim P2              As Variant   'Materialized payload of nYears
    Dim K2              As KPR_ArgShape 'Shape kind of nYears
    Dim R2              As Long      'Rows of nYears
    Dim C2              As Long      'Cols of nYears
    Dim KeepEOM         As Boolean    'Resolved Opt_KeepEOM
    Dim OutKind         As KPR_ArgShape 'Resolved output kind
    Dim OutRows         As Long      'Resolved output rows
    Dim OutCols         As Long      'Resolved output cols
    Dim OutArr          As Variant   'Output array for a multi-element call
    Dim R               As Long      'Row cursor
    Dim C               As Long      'Column cursor
    Dim Condition       As KPR_Condition 'Call-level shape or capacity condition
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' CALL-LEVEL STAGES (contract section 5.2)
'------------------------------------------------------------------------------
    'Stage 1: refuse a 1904 worksheet host before any argument is touched
        If Not PassHostGuard(FailErr) Then GoTo Fail

    'Stage 2: classify every value argument from type and dimensions alone
        If Not TryClassifyShape(DateIn, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryClassifyShape(nYears, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 3a: optional controls, scalar or 1x1 only (CONTROL_NOT_SCALAR)
        If Not TryResolveBool(Opt_KeepEOM, False, KeepEOM, FailErr) Then GoTo Fail

    'Stage 3b: exact-shape broadcast resolution (SHAPE_MISMATCH)
        OutKind = KPR_SHAPE_SCALAR: OutRows = 1: OutCols = 1
        If Not AccumulateShape(OutKind, OutRows, OutCols, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not AccumulateShape(OutKind, OutRows, OutCols, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 4: the element cap, once, on the resolved output shape
        If Not CheckCapacity(OutRows, OutCols, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 5: only now may content be read
        If Not TryMaterialize(DateIn, P1, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryMaterialize(nYears, P2, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

'------------------------------------------------------------------------------
' TRAVERSE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar: the 1x1 case of the element
        If OutKind = KPR_SHAPE_SCALAR Then
            KPR_Dates_AddYears = Elem_AddYears(P1, P2, KeepEOM)
            Exit Function
        End If

    'A multi-element call returns a 1-based 2-D array of the resolved shape,
    'evaluated row-major; each element resolves its own inputs
        If Not TryAllocateOutput(OutRows, OutCols, OutArr, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        For R = 1 To OutRows
            For C = 1 To OutCols
                OutArr(R, C) = Elem_AddYears(ElementAt(P1, K1, R, C), ElementAt(P2, K2, R, C), KeepEOM)
            Next C
        Next R
        KPR_Dates_AddYears = OutArr
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_AddYears = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed with #VALUE!; reaching here is a defect
        FailErr = ErrValue()
        Resume Fail

End Function
'
'------------------------------------------------------------------------------
'
'                        PUBLIC API - WEEKDAY LOCATORS                         
'
'------------------------------------------------------------------------------
'

Public Function KPR_Dates_NthWeekdayOfMonth( _
    ByVal YearIn As Variant, _
    ByVal MonthIn As Variant, _
    ByVal WdIndex As Variant, _
    ByVal n As Variant, _
    Optional ByVal Opt_WeekBaseMonday As Variant = True) _
    As Variant
'
'==============================================================================
'                          KPR_Dates_NthWeekdayOfMonth
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the nth occurrence of a weekday in a month.
'
' SIGNATURE
'   KPR_Dates_NthWeekdayOfMonth(YearIn, MonthIn, WdIndex, n, [Opt_WeekBaseMonday]) -> Variant (Date)
'
' INPUTS
'   YearIn
'     Native integral numeric, domain 1900 through 9999.
'   MonthIn
'     Native integral numeric, domain 1 through 12.
'   WdIndex
'     Native integral numeric, domain 1 through 7.
'   n
'     Native integral numeric, domain 1 through 5.
'   Opt_WeekBaseMonday
'     Optional. Omitted or Empty selects True. Native Boolean only.
'
' RETURNS
'   Variant
'     Date on success, or a native Excel error value.
'
' ERROR POLICY (USER FACING)
'   #VALUE!  a multi-cell optional control; a shape that is not a scalar,
'            Range or array, or a multi-area, empty, jagged or rank-3+ input;
'            two non-scalar arguments of different shapes; an integer argument that is fractional, Boolean, text, or outside its domain; an optional control of the wrong type or an unknown token
'   #NUM!    an output of 100,001 or more elements; an integer outside the Long range; a result outside the supported window
'   #N/A     an identified 1904 worksheet caller, or an identified caller
'            whose date system cannot be read
'   Incoming native errors are returned unchanged.
'
' SHAPE
'   Value arguments may be a scalar, a single-cell Range or 1x1 wrapper, or a
'   multi-element Range or array. A scalar expands; every non-scalar argument
'   must match exactly in rows and columns. An all-scalar call returns a
'   scalar; any other call returns a 1-based 2-D Variant of the resolved
'   shape, evaluated row-major, with each element resolved independently.
'   Multi-cell results are supported on dynamic-array Excel only.
'
' ORDER OF EVALUATION (call-level, contract section 5.2)
'   host guard -> classify every value argument -> optional controls and
'   broadcast resolution -> element cap -> materialize -> traverse.
'   Within an element, value arguments resolve in signature order.
'
' DEPENDENCIES
'   - PassHostGuard
'   - Elem_NthWeekdayOfMonth
'   - KPR_Core_Array: TryClassifyShape, AccumulateShape, CheckCapacity,
'     TryMaterialize, TryAllocateOutput, ElementAt
'   - KPR_Core_Err.ErrForCondition
'   - TryResolveBool (KPR_Core_Array, KPR_Core_Parse)
'
' NOTES
'   - A fifth occurrence that the month does not contain is OCCURRENCE_ABSENT; a valid locator before 1900-03-01 is RESULT_WINDOW.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P1              As Variant   'Materialized payload of YearIn
    Dim K1              As KPR_ArgShape 'Shape kind of YearIn
    Dim R1              As Long      'Rows of YearIn
    Dim C1              As Long      'Cols of YearIn
    Dim P2              As Variant   'Materialized payload of MonthIn
    Dim K2              As KPR_ArgShape 'Shape kind of MonthIn
    Dim R2              As Long      'Rows of MonthIn
    Dim C2              As Long      'Cols of MonthIn
    Dim P3              As Variant   'Materialized payload of WdIndex
    Dim K3              As KPR_ArgShape 'Shape kind of WdIndex
    Dim R3              As Long      'Rows of WdIndex
    Dim C3              As Long      'Cols of WdIndex
    Dim P4              As Variant   'Materialized payload of n
    Dim K4              As KPR_ArgShape 'Shape kind of n
    Dim R4              As Long      'Rows of n
    Dim C4              As Long      'Cols of n
    Dim WkMonday        As Boolean    'Resolved Opt_WeekBaseMonday
    Dim OutKind         As KPR_ArgShape 'Resolved output kind
    Dim OutRows         As Long      'Resolved output rows
    Dim OutCols         As Long      'Resolved output cols
    Dim OutArr          As Variant   'Output array for a multi-element call
    Dim R               As Long      'Row cursor
    Dim C               As Long      'Column cursor
    Dim Condition       As KPR_Condition 'Call-level shape or capacity condition
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' CALL-LEVEL STAGES (contract section 5.2)
'------------------------------------------------------------------------------
    'Stage 1: refuse a 1904 worksheet host before any argument is touched
        If Not PassHostGuard(FailErr) Then GoTo Fail

    'Stage 2: classify every value argument from type and dimensions alone
        If Not TryClassifyShape(YearIn, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryClassifyShape(MonthIn, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryClassifyShape(WdIndex, K3, R3, C3, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryClassifyShape(n, K4, R4, C4, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 3a: optional controls, scalar or 1x1 only (CONTROL_NOT_SCALAR)
        If Not TryResolveBool(Opt_WeekBaseMonday, True, WkMonday, FailErr) Then GoTo Fail

    'Stage 3b: exact-shape broadcast resolution (SHAPE_MISMATCH)
        OutKind = KPR_SHAPE_SCALAR: OutRows = 1: OutCols = 1
        If Not AccumulateShape(OutKind, OutRows, OutCols, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not AccumulateShape(OutKind, OutRows, OutCols, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not AccumulateShape(OutKind, OutRows, OutCols, K3, R3, C3, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not AccumulateShape(OutKind, OutRows, OutCols, K4, R4, C4, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 4: the element cap, once, on the resolved output shape
        If Not CheckCapacity(OutRows, OutCols, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 5: only now may content be read
        If Not TryMaterialize(YearIn, P1, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryMaterialize(MonthIn, P2, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryMaterialize(WdIndex, P3, K3, R3, C3, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryMaterialize(n, P4, K4, R4, C4, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

'------------------------------------------------------------------------------
' TRAVERSE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar: the 1x1 case of the element
        If OutKind = KPR_SHAPE_SCALAR Then
            KPR_Dates_NthWeekdayOfMonth = Elem_NthWeekdayOfMonth(P1, P2, P3, P4, WkMonday)
            Exit Function
        End If

    'A multi-element call returns a 1-based 2-D array of the resolved shape,
    'evaluated row-major; each element resolves its own inputs
        If Not TryAllocateOutput(OutRows, OutCols, OutArr, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        For R = 1 To OutRows
            For C = 1 To OutCols
                OutArr(R, C) = Elem_NthWeekdayOfMonth(ElementAt(P1, K1, R, C), ElementAt(P2, K2, R, C), ElementAt(P3, K3, R, C), ElementAt(P4, K4, R, C), WkMonday)
            Next C
        Next R
        KPR_Dates_NthWeekdayOfMonth = OutArr
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_NthWeekdayOfMonth = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed with #VALUE!; reaching here is a defect
        FailErr = ErrValue()
        Resume Fail

End Function
Public Function KPR_Dates_LastWeekdayOfMonth( _
    ByVal YearIn As Variant, _
    ByVal MonthIn As Variant, _
    ByVal WdIndex As Variant, _
    Optional ByVal Opt_WeekBaseMonday As Variant = True) _
    As Variant
'
'==============================================================================
'                          KPR_Dates_LastWeekdayOfMonth
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the final occurrence of a weekday in a month.
'
' SIGNATURE
'   KPR_Dates_LastWeekdayOfMonth(YearIn, MonthIn, WdIndex, [Opt_WeekBaseMonday]) -> Variant (Date)
'
' INPUTS
'   YearIn
'     Native integral numeric, domain 1900 through 9999.
'   MonthIn
'     Native integral numeric, domain 1 through 12.
'   WdIndex
'     Native integral numeric, domain 1 through 7.
'   Opt_WeekBaseMonday
'     Optional. Omitted or Empty selects True. Native Boolean only.
'
' RETURNS
'   Variant
'     Date on success, or a native Excel error value.
'
' ERROR POLICY (USER FACING)
'   #VALUE!  a multi-cell optional control; a shape that is not a scalar,
'            Range or array, or a multi-area, empty, jagged or rank-3+ input;
'            two non-scalar arguments of different shapes; an integer argument that is fractional, Boolean, text, or outside its domain; an optional control of the wrong type or an unknown token
'   #NUM!    an output of 100,001 or more elements; an integer outside the Long range; a result outside the supported window
'   #N/A     an identified 1904 worksheet caller, or an identified caller
'            whose date system cannot be read
'   Incoming native errors are returned unchanged.
'
' SHAPE
'   Value arguments may be a scalar, a single-cell Range or 1x1 wrapper, or a
'   multi-element Range or array. A scalar expands; every non-scalar argument
'   must match exactly in rows and columns. An all-scalar call returns a
'   scalar; any other call returns a 1-based 2-D Variant of the resolved
'   shape, evaluated row-major, with each element resolved independently.
'   Multi-cell results are supported on dynamic-array Excel only.
'
' ORDER OF EVALUATION (call-level, contract section 5.2)
'   host guard -> classify every value argument -> optional controls and
'   broadcast resolution -> element cap -> materialize -> traverse.
'   Within an element, value arguments resolve in signature order.
'
' DEPENDENCIES
'   - PassHostGuard
'   - Elem_LastWeekdayOfMonth
'   - KPR_Core_Array: TryClassifyShape, AccumulateShape, CheckCapacity,
'     TryMaterialize, TryAllocateOutput, ElementAt
'   - KPR_Core_Err.ErrForCondition
'   - TryResolveBool (KPR_Core_Array, KPR_Core_Parse)
'
' NOTES
'   - None.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P1              As Variant   'Materialized payload of YearIn
    Dim K1              As KPR_ArgShape 'Shape kind of YearIn
    Dim R1              As Long      'Rows of YearIn
    Dim C1              As Long      'Cols of YearIn
    Dim P2              As Variant   'Materialized payload of MonthIn
    Dim K2              As KPR_ArgShape 'Shape kind of MonthIn
    Dim R2              As Long      'Rows of MonthIn
    Dim C2              As Long      'Cols of MonthIn
    Dim P3              As Variant   'Materialized payload of WdIndex
    Dim K3              As KPR_ArgShape 'Shape kind of WdIndex
    Dim R3              As Long      'Rows of WdIndex
    Dim C3              As Long      'Cols of WdIndex
    Dim WkMonday        As Boolean    'Resolved Opt_WeekBaseMonday
    Dim OutKind         As KPR_ArgShape 'Resolved output kind
    Dim OutRows         As Long      'Resolved output rows
    Dim OutCols         As Long      'Resolved output cols
    Dim OutArr          As Variant   'Output array for a multi-element call
    Dim R               As Long      'Row cursor
    Dim C               As Long      'Column cursor
    Dim Condition       As KPR_Condition 'Call-level shape or capacity condition
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' CALL-LEVEL STAGES (contract section 5.2)
'------------------------------------------------------------------------------
    'Stage 1: refuse a 1904 worksheet host before any argument is touched
        If Not PassHostGuard(FailErr) Then GoTo Fail

    'Stage 2: classify every value argument from type and dimensions alone
        If Not TryClassifyShape(YearIn, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryClassifyShape(MonthIn, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryClassifyShape(WdIndex, K3, R3, C3, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 3a: optional controls, scalar or 1x1 only (CONTROL_NOT_SCALAR)
        If Not TryResolveBool(Opt_WeekBaseMonday, True, WkMonday, FailErr) Then GoTo Fail

    'Stage 3b: exact-shape broadcast resolution (SHAPE_MISMATCH)
        OutKind = KPR_SHAPE_SCALAR: OutRows = 1: OutCols = 1
        If Not AccumulateShape(OutKind, OutRows, OutCols, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not AccumulateShape(OutKind, OutRows, OutCols, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not AccumulateShape(OutKind, OutRows, OutCols, K3, R3, C3, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 4: the element cap, once, on the resolved output shape
        If Not CheckCapacity(OutRows, OutCols, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 5: only now may content be read
        If Not TryMaterialize(YearIn, P1, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryMaterialize(MonthIn, P2, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryMaterialize(WdIndex, P3, K3, R3, C3, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

'------------------------------------------------------------------------------
' TRAVERSE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar: the 1x1 case of the element
        If OutKind = KPR_SHAPE_SCALAR Then
            KPR_Dates_LastWeekdayOfMonth = Elem_LastWeekdayOfMonth(P1, P2, P3, WkMonday)
            Exit Function
        End If

    'A multi-element call returns a 1-based 2-D array of the resolved shape,
    'evaluated row-major; each element resolves its own inputs
        If Not TryAllocateOutput(OutRows, OutCols, OutArr, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        For R = 1 To OutRows
            For C = 1 To OutCols
                OutArr(R, C) = Elem_LastWeekdayOfMonth(ElementAt(P1, K1, R, C), ElementAt(P2, K2, R, C), ElementAt(P3, K3, R, C), WkMonday)
            Next C
        Next R
        KPR_Dates_LastWeekdayOfMonth = OutArr
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_LastWeekdayOfMonth = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed with #VALUE!; reaching here is a defect
        FailErr = ErrValue()
        Resume Fail

End Function
'
'------------------------------------------------------------------------------
'
'                        PUBLIC API - PILLAR FORMATTING                        
'
'------------------------------------------------------------------------------
'

Public Function KPR_Dates_PillarFromDates( _
    ByVal StartDate As Variant, _
    ByVal EndDate As Variant, _
    Optional ByVal Opt_Rounding As Variant = "NEAREST") _
    As Variant
'
'==============================================================================
'                          KPR_Dates_PillarFromDates
'------------------------------------------------------------------------------
' PURPOSE
'   Labels the interval between two dates as a canonical pillar token.
'
' SIGNATURE
'   KPR_Dates_PillarFromDates(StartDate, EndDate, [Opt_Rounding]) -> Variant (String)
'
' INPUTS
'   StartDate
'     Date, numeric serial, or ISO YYYY-MM-DD text; single-cell Range or 1x1 wrapper accepted.
'   EndDate
'     Date, numeric serial, or ISO YYYY-MM-DD text; single-cell Range or 1x1 wrapper accepted.
'   Opt_Rounding
'     Optional. Omitted or Empty selects "NEAREST". Native text NEAREST, FLOOR or CEILING only.
'
' RETURNS
'   Variant
'     String on success, or a native Excel error value.
'
' ERROR POLICY (USER FACING)
'   #VALUE!  a multi-cell optional control; a shape that is not a scalar,
'            Range or array, or a multi-area, empty, jagged or rank-3+ input;
'            two non-scalar arguments of different shapes; a date argument that is not scalar, not a date, or malformed text; an optional control of the wrong type or an unknown token
'   #NUM!    an output of 100,001 or more elements; a date outside KPR_MIN_DATE .. KPR_MAX_DATE
'   #N/A     an identified 1904 worksheet caller, or an identified caller
'            whose date system cannot be read
'   Incoming native errors are returned unchanged.
'
' SHAPE
'   Value arguments may be a scalar, a single-cell Range or 1x1 wrapper, or a
'   multi-element Range or array. A scalar expands; every non-scalar argument
'   must match exactly in rows and columns. An all-scalar call returns a
'   scalar; any other call returns a 1-based 2-D Variant of the resolved
'   shape, evaluated row-major, with each element resolved independently.
'   Multi-cell results are supported on dynamic-array Excel only.
'
' ORDER OF EVALUATION (call-level, contract section 5.2)
'   host guard -> classify every value argument -> optional controls and
'   broadcast resolution -> element cap -> materialize -> traverse.
'   Within an element, value arguments resolve in signature order.
'
' DEPENDENCIES
'   - PassHostGuard
'   - Elem_PillarFromDates
'   - KPR_Core_Array: TryClassifyShape, AccumulateShape, CheckCapacity,
'     TryMaterialize, TryAllocateOutput, ElementAt
'   - KPR_Core_Err.ErrForCondition
'   - TryResolveRounding (KPR_Core_Array, KPR_Core_Parse)
'
' NOTES
'   - Rounding modes, the candidate set and tie rules are specified in contract section 8.4.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P1              As Variant   'Materialized payload of StartDate
    Dim K1              As KPR_ArgShape 'Shape kind of StartDate
    Dim R1              As Long      'Rows of StartDate
    Dim C1              As Long      'Cols of StartDate
    Dim P2              As Variant   'Materialized payload of EndDate
    Dim K2              As KPR_ArgShape 'Shape kind of EndDate
    Dim R2              As Long      'Rows of EndDate
    Dim C2              As Long      'Cols of EndDate
    Dim Mode            As KPR_PillarRounding 'Resolved Opt_Rounding
    Dim OutKind         As KPR_ArgShape 'Resolved output kind
    Dim OutRows         As Long      'Resolved output rows
    Dim OutCols         As Long      'Resolved output cols
    Dim OutArr          As Variant   'Output array for a multi-element call
    Dim R               As Long      'Row cursor
    Dim C               As Long      'Column cursor
    Dim Condition       As KPR_Condition 'Call-level shape or capacity condition
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' CALL-LEVEL STAGES (contract section 5.2)
'------------------------------------------------------------------------------
    'Stage 1: refuse a 1904 worksheet host before any argument is touched
        If Not PassHostGuard(FailErr) Then GoTo Fail

    'Stage 2: classify every value argument from type and dimensions alone
        If Not TryClassifyShape(StartDate, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryClassifyShape(EndDate, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 3a: optional controls, scalar or 1x1 only (CONTROL_NOT_SCALAR)
        If Not TryResolveRounding(Opt_Rounding, Mode, FailErr) Then GoTo Fail

    'Stage 3b: exact-shape broadcast resolution (SHAPE_MISMATCH)
        OutKind = KPR_SHAPE_SCALAR: OutRows = 1: OutCols = 1
        If Not AccumulateShape(OutKind, OutRows, OutCols, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not AccumulateShape(OutKind, OutRows, OutCols, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 4: the element cap, once, on the resolved output shape
        If Not CheckCapacity(OutRows, OutCols, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 5: only now may content be read
        If Not TryMaterialize(StartDate, P1, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryMaterialize(EndDate, P2, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

'------------------------------------------------------------------------------
' TRAVERSE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar: the 1x1 case of the element
        If OutKind = KPR_SHAPE_SCALAR Then
            KPR_Dates_PillarFromDates = Elem_PillarFromDates(P1, P2, Mode)
            Exit Function
        End If

    'A multi-element call returns a 1-based 2-D array of the resolved shape,
    'evaluated row-major; each element resolves its own inputs
        If Not TryAllocateOutput(OutRows, OutCols, OutArr, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        For R = 1 To OutRows
            For C = 1 To OutCols
                OutArr(R, C) = Elem_PillarFromDates(ElementAt(P1, K1, R, C), ElementAt(P2, K2, R, C), Mode)
            Next C
        Next R
        KPR_Dates_PillarFromDates = OutArr
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_PillarFromDates = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed with #VALUE!; reaching here is a defect
        FailErr = ErrValue()
        Resume Fail

End Function
Public Function KPR_Dates_DateFromPillar( _
    ByVal StartDate As Variant, _
    ByVal Pillar As Variant) _
    As Variant
'
'==============================================================================
'                          KPR_Dates_DateFromPillar
'------------------------------------------------------------------------------
' PURPOSE
'   Shifts a start date by a parsed pillar token.
'
' SIGNATURE
'   KPR_Dates_DateFromPillar(StartDate, Pillar) -> Variant (Date)
'
' INPUTS
'   StartDate
'     Date, numeric serial, or ISO YYYY-MM-DD text; single-cell Range or 1x1 wrapper accepted.
'   Pillar
'     Pillar token text under the contract grammar (section 3.4).
'
' RETURNS
'   Variant
'     Date on success, or a native Excel error value.
'
' ERROR POLICY (USER FACING)
'   #VALUE!  a multi-cell optional control; a shape that is not a scalar,
'            Range or array, or a multi-area, empty, jagged or rank-3+ input;
'            two non-scalar arguments of different shapes; a date argument that is not scalar, not a date, or malformed text; a pillar token outside the accepted grammar
'   #NUM!    an output of 100,001 or more elements; a date outside KPR_MIN_DATE .. KPR_MAX_DATE; a result outside the supported window
'   #N/A     an identified 1904 worksheet caller, or an identified caller
'            whose date system cannot be read
'   Incoming native errors are returned unchanged.
'
' SHAPE
'   Value arguments may be a scalar, a single-cell Range or 1x1 wrapper, or a
'   multi-element Range or array. A scalar expands; every non-scalar argument
'   must match exactly in rows and columns. An all-scalar call returns a
'   scalar; any other call returns a 1-based 2-D Variant of the resolved
'   shape, evaluated row-major, with each element resolved independently.
'   Multi-cell results are supported on dynamic-array Excel only.
'
' ORDER OF EVALUATION (call-level, contract section 5.2)
'   host guard -> classify every value argument -> optional controls and
'   broadcast resolution -> element cap -> materialize -> traverse.
'   Within an element, value arguments resolve in signature order.
'
' DEPENDENCIES
'   - PassHostGuard
'   - Elem_DateFromPillar
'   - KPR_Core_Array: TryClassifyShape, AccumulateShape, CheckCapacity,
'     TryMaterialize, TryAllocateOutput, ElementAt
'   - KPR_Core_Err.ErrForCondition
'
' NOTES
'   - Months first with clip semantics, then exact days, matching the parser's grammar. The name is singular because the function returns exactly one date per element.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim P1              As Variant   'Materialized payload of StartDate
    Dim K1              As KPR_ArgShape 'Shape kind of StartDate
    Dim R1              As Long      'Rows of StartDate
    Dim C1              As Long      'Cols of StartDate
    Dim P2              As Variant   'Materialized payload of Pillar
    Dim K2              As KPR_ArgShape 'Shape kind of Pillar
    Dim R2              As Long      'Rows of Pillar
    Dim C2              As Long      'Cols of Pillar
    Dim OutKind         As KPR_ArgShape 'Resolved output kind
    Dim OutRows         As Long      'Resolved output rows
    Dim OutCols         As Long      'Resolved output cols
    Dim OutArr          As Variant   'Output array for a multi-element call
    Dim R               As Long      'Row cursor
    Dim C               As Long      'Column cursor
    Dim Condition       As KPR_Condition 'Call-level shape or capacity condition
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' CALL-LEVEL STAGES (contract section 5.2)
'------------------------------------------------------------------------------
    'Stage 1: refuse a 1904 worksheet host before any argument is touched
        If Not PassHostGuard(FailErr) Then GoTo Fail

    'Stage 2: classify every value argument from type and dimensions alone
        If Not TryClassifyShape(StartDate, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryClassifyShape(Pillar, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 3b: exact-shape broadcast resolution (SHAPE_MISMATCH)
        OutKind = KPR_SHAPE_SCALAR: OutRows = 1: OutCols = 1
        If Not AccumulateShape(OutKind, OutRows, OutCols, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not AccumulateShape(OutKind, OutRows, OutCols, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 4: the element cap, once, on the resolved output shape
        If Not CheckCapacity(OutRows, OutCols, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

    'Stage 5: only now may content be read
        If Not TryMaterialize(StartDate, P1, K1, R1, C1, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        If Not TryMaterialize(Pillar, P2, K2, R2, C2, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail

'------------------------------------------------------------------------------
' TRAVERSE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar: the 1x1 case of the element
        If OutKind = KPR_SHAPE_SCALAR Then
            KPR_Dates_DateFromPillar = Elem_DateFromPillar(P1, P2)
            Exit Function
        End If

    'A multi-element call returns a 1-based 2-D array of the resolved shape,
    'evaluated row-major; each element resolves its own inputs
        If Not TryAllocateOutput(OutRows, OutCols, OutArr, Condition) Then FailErr = ErrForCondition(Condition): GoTo Fail
        For R = 1 To OutRows
            For C = 1 To OutCols
                OutArr(R, C) = Elem_DateFromPillar(ElementAt(P1, K1, R, C), ElementAt(P2, K2, R, C))
            Next C
        Next R
        KPR_Dates_DateFromPillar = OutArr
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_DateFromPillar = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed with #VALUE!; reaching here is a defect
        FailErr = ErrValue()
        Resume Fail

End Function
'
'------------------------------------------------------------------------------
'
'                       PRIVATE ELEMENT IMPLEMENTATIONS                        
'
'------------------------------------------------------------------------------
'


Private Function Elem_DayOfWeek( _
    ByVal DateIn As Variant, _
    ByVal WkMonday As Boolean) _
    As Variant
'
'==============================================================================
'                              Elem_DayOfWeek
'------------------------------------------------------------------------------
' PURPOSE
'   One element of KPR_Dates_DayOfWeek: resolves the raw value arguments,
'   computes, and returns the Long or the element's native error value.
'
' ELEMENT CONTRACT
'   - Value arguments arrive raw (scalar, single-cell Range or 1x1 wrapper)
'     and are resolved here in signature order; the first failing argument
'     determines the result.
'   - Controls arrive already resolved, because a control is call-level and
'     is never re-resolved per element.
'   - Never calls the host guard: the guard runs once per public call, before
'     materialization, so scalar and array calls share one host decision.
'   - Never raises. Every failure returns a value the caller can place at the
'     element's output position.
'
' ERROR POLICY
'   Registered conditions only, mapped through KPR_Core_Err.ErrForCondition.
'   The containment handler is a defect if reached.
'
' NOTES
'   - TRUE selects Monday-based numbering (ISO habit); FALSE selects Sunday-based.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim D               As Date     'Resolved DateIn
    Dim WkBase          As VbDayOfWeek 'Weekday base selected by the control
    Dim ErrOut          As Variant   'Element-level error value from a resolver

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: an element must never expose a runtime exception
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Resolve DateIn: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveDate(DateIn, D, ErrOut) Then Elem_DayOfWeek = ErrOut: Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        If WkMonday Then WkBase = vbMonday Else WkBase = vbSunday
        Elem_DayOfWeek = CLng(Weekday(D, WkBase))
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed at this element; reaching here is a defect
        Elem_DayOfWeek = ErrValue()

End Function

Private Function Elem_DaysInMonth( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
'                              Elem_DaysInMonth
'------------------------------------------------------------------------------
' PURPOSE
'   One element of KPR_Dates_DaysInMonth: resolves the raw value arguments,
'   computes, and returns the Long or the element's native error value.
'
' ELEMENT CONTRACT
'   - Value arguments arrive raw (scalar, single-cell Range or 1x1 wrapper)
'     and are resolved here in signature order; the first failing argument
'     determines the result.
'   - Controls arrive already resolved, because a control is call-level and
'     is never re-resolved per element.
'   - Never calls the host guard: the guard runs once per public call, before
'     materialization, so scalar and array calls share one host decision.
'   - Never raises. Every failure returns a value the caller can place at the
'     element's output position.
'
' ERROR POLICY
'   Registered conditions only, mapped through KPR_Core_Err.ErrForCondition.
'   The containment handler is a defect if reached.
'
' NOTES
'   - None.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim D               As Date     'Resolved DateIn
    Dim ErrOut          As Variant   'Element-level error value from a resolver

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: an element must never expose a runtime exception
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Resolve DateIn: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveDate(DateIn, D, ErrOut) Then Elem_DaysInMonth = ErrOut: Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        Elem_DaysInMonth = DaysInMonth(Year(D), Month(D))
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed at this element; reaching here is a defect
        Elem_DaysInMonth = ErrValue()

End Function

Private Function Elem_DaysInYear( _
    ByVal YearIn As Variant) _
    As Variant
'
'==============================================================================
'                              Elem_DaysInYear
'------------------------------------------------------------------------------
' PURPOSE
'   One element of KPR_Dates_DaysInYear: resolves the raw value arguments,
'   computes, and returns the Long or the element's native error value.
'
' ELEMENT CONTRACT
'   - Value arguments arrive raw (scalar, single-cell Range or 1x1 wrapper)
'     and are resolved here in signature order; the first failing argument
'     determines the result.
'   - Controls arrive already resolved, because a control is call-level and
'     is never re-resolved per element.
'   - Never calls the host guard: the guard runs once per public call, before
'     materialization, so scalar and array calls share one host decision.
'   - Never raises. Every failure returns a value the caller can place at the
'     element's output position.
'
' ERROR POLICY
'   Registered conditions only, mapped through KPR_Core_Err.ErrForCondition.
'   The containment handler is a defect if reached.
'
' NOTES
'   - Takes a calendar year, not a date. No date is constructed, so no window gate applies and year 1900 is fully in the domain.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Y               As Long     'Resolved YearIn
    Dim ErrOut          As Variant   'Element-level error value from a resolver

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: an element must never expose a runtime exception
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Resolve YearIn: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveLong(YearIn, Y, ErrOut) Then Elem_DaysInYear = ErrOut: Exit Function
    'Domain of YearIn is 1900 through 9999
        If (Y < Year(KPR_MIN_DATE)) Or (Y > Year(KPR_MAX_DATE)) Then Elem_DaysInYear = ErrForCondition(KPR_COND_DOMAIN_YEAR): Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        If IsLeapYear(Y) Then Elem_DaysInYear = 366& Else Elem_DaysInYear = 365&
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed at this element; reaching here is a defect
        Elem_DaysInYear = ErrValue()

End Function

Private Function Elem_BeginOfMonth( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
'                              Elem_BeginOfMonth
'------------------------------------------------------------------------------
' PURPOSE
'   One element of KPR_Dates_BeginOfMonth: resolves the raw value arguments,
'   computes, and returns the Date or the element's native error value.
'
' ELEMENT CONTRACT
'   - Value arguments arrive raw (scalar, single-cell Range or 1x1 wrapper)
'     and are resolved here in signature order; the first failing argument
'     determines the result.
'   - Controls arrive already resolved, because a control is call-level and
'     is never re-resolved per element.
'   - Never calls the host guard: the guard runs once per public call, before
'     materialization, so scalar and array calls share one host decision.
'   - Never raises. Every failure returns a value the caller can place at the
'     element's output position.
'
' ERROR POLICY
'   Registered conditions only, mapped through KPR_Core_Err.ErrForCondition.
'   The containment handler is a defect if reached.
'
' NOTES
'   - None.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim D               As Date     'Resolved DateIn
    Dim ErrOut          As Variant   'Element-level error value from a resolver

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: an element must never expose a runtime exception
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Resolve DateIn: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveDate(DateIn, D, ErrOut) Then Elem_BeginOfMonth = ErrOut: Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        Elem_BeginOfMonth = DateSerial(Year(D), Month(D), 1)
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed at this element; reaching here is a defect
        Elem_BeginOfMonth = ErrValue()

End Function

Private Function Elem_EndOfMonth( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
'                              Elem_EndOfMonth
'------------------------------------------------------------------------------
' PURPOSE
'   One element of KPR_Dates_EndOfMonth: resolves the raw value arguments,
'   computes, and returns the Date or the element's native error value.
'
' ELEMENT CONTRACT
'   - Value arguments arrive raw (scalar, single-cell Range or 1x1 wrapper)
'     and are resolved here in signature order; the first failing argument
'     determines the result.
'   - Controls arrive already resolved, because a control is call-level and
'     is never re-resolved per element.
'   - Never calls the host guard: the guard runs once per public call, before
'     materialization, so scalar and array calls share one host decision.
'   - Never raises. Every failure returns a value the caller can place at the
'     element's output position.
'
' ERROR POLICY
'   Registered conditions only, mapped through KPR_Core_Err.ErrForCondition.
'   The containment handler is a defect if reached.
'
' NOTES
'   - None.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim D               As Date     'Resolved DateIn
    Dim ErrOut          As Variant   'Element-level error value from a resolver

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: an element must never expose a runtime exception
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Resolve DateIn: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveDate(DateIn, D, ErrOut) Then Elem_EndOfMonth = ErrOut: Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        Elem_EndOfMonth = EndOfMonth(D)
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed at this element; reaching here is a defect
        Elem_EndOfMonth = ErrValue()

End Function

Private Function Elem_BeginOfQuarter( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
'                              Elem_BeginOfQuarter
'------------------------------------------------------------------------------
' PURPOSE
'   One element of KPR_Dates_BeginOfQuarter: resolves the raw value arguments,
'   computes, and returns the Date or the element's native error value.
'
' ELEMENT CONTRACT
'   - Value arguments arrive raw (scalar, single-cell Range or 1x1 wrapper)
'     and are resolved here in signature order; the first failing argument
'     determines the result.
'   - Controls arrive already resolved, because a control is call-level and
'     is never re-resolved per element.
'   - Never calls the host guard: the guard runs once per public call, before
'     materialization, so scalar and array calls share one host decision.
'   - Never raises. Every failure returns a value the caller can place at the
'     element's output position.
'
' ERROR POLICY
'   Registered conditions only, mapped through KPR_Core_Err.ErrForCondition.
'   The containment handler is a defect if reached.
'
' NOTES
'   - A Q1-1900 input names 1900-01-01, which is outside the supported window: RESULT_WINDOW, not a date.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim D               As Date     'Resolved DateIn
    Dim R               As Date      'Boundary from the calendar core
    Dim ErrOut          As Variant   'Element-level error value from a resolver

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: an element must never expose a runtime exception
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Resolve DateIn: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveDate(DateIn, D, ErrOut) Then Elem_BeginOfQuarter = ErrOut: Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        R = BeginOfQuarter(D)
        If Not IsDateInWindow(R) Then Elem_BeginOfQuarter = ErrForCondition(KPR_COND_RESULT_WINDOW): Exit Function
        Elem_BeginOfQuarter = R
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed at this element; reaching here is a defect
        Elem_BeginOfQuarter = ErrValue()

End Function

Private Function Elem_EndOfQuarter( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
'                              Elem_EndOfQuarter
'------------------------------------------------------------------------------
' PURPOSE
'   One element of KPR_Dates_EndOfQuarter: resolves the raw value arguments,
'   computes, and returns the Date or the element's native error value.
'
' ELEMENT CONTRACT
'   - Value arguments arrive raw (scalar, single-cell Range or 1x1 wrapper)
'     and are resolved here in signature order; the first failing argument
'     determines the result.
'   - Controls arrive already resolved, because a control is call-level and
'     is never re-resolved per element.
'   - Never calls the host guard: the guard runs once per public call, before
'     materialization, so scalar and array calls share one host decision.
'   - Never raises. Every failure returns a value the caller can place at the
'     element's output position.
'
' ERROR POLICY
'   Registered conditions only, mapped through KPR_Core_Err.ErrForCondition.
'   The containment handler is a defect if reached.
'
' NOTES
'   - Always inside the window for an in-window input: the quarter end is never earlier than the input.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim D               As Date     'Resolved DateIn
    Dim ErrOut          As Variant   'Element-level error value from a resolver

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: an element must never expose a runtime exception
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Resolve DateIn: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveDate(DateIn, D, ErrOut) Then Elem_EndOfQuarter = ErrOut: Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        Elem_EndOfQuarter = EndOfQuarter(D)
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed at this element; reaching here is a defect
        Elem_EndOfQuarter = ErrValue()

End Function

Private Function Elem_BeginOfYear( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
'                              Elem_BeginOfYear
'------------------------------------------------------------------------------
' PURPOSE
'   One element of KPR_Dates_BeginOfYear: resolves the raw value arguments,
'   computes, and returns the Date or the element's native error value.
'
' ELEMENT CONTRACT
'   - Value arguments arrive raw (scalar, single-cell Range or 1x1 wrapper)
'     and are resolved here in signature order; the first failing argument
'     determines the result.
'   - Controls arrive already resolved, because a control is call-level and
'     is never re-resolved per element.
'   - Never calls the host guard: the guard runs once per public call, before
'     materialization, so scalar and array calls share one host decision.
'   - Never raises. Every failure returns a value the caller can place at the
'     element's output position.
'
' ERROR POLICY
'   Registered conditions only, mapped through KPR_Core_Err.ErrForCondition.
'   The containment handler is a defect if reached.
'
' NOTES
'   - Any 1900 input names 1900-01-01, which is outside the supported window: RESULT_WINDOW, not a date.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim D               As Date     'Resolved DateIn
    Dim R               As Date      'Boundary from the calendar core
    Dim ErrOut          As Variant   'Element-level error value from a resolver

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: an element must never expose a runtime exception
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Resolve DateIn: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveDate(DateIn, D, ErrOut) Then Elem_BeginOfYear = ErrOut: Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        R = BeginOfYear(D)
        If Not IsDateInWindow(R) Then Elem_BeginOfYear = ErrForCondition(KPR_COND_RESULT_WINDOW): Exit Function
        Elem_BeginOfYear = R
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed at this element; reaching here is a defect
        Elem_BeginOfYear = ErrValue()

End Function

Private Function Elem_EndOfYear( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
'                              Elem_EndOfYear
'------------------------------------------------------------------------------
' PURPOSE
'   One element of KPR_Dates_EndOfYear: resolves the raw value arguments,
'   computes, and returns the Date or the element's native error value.
'
' ELEMENT CONTRACT
'   - Value arguments arrive raw (scalar, single-cell Range or 1x1 wrapper)
'     and are resolved here in signature order; the first failing argument
'     determines the result.
'   - Controls arrive already resolved, because a control is call-level and
'     is never re-resolved per element.
'   - Never calls the host guard: the guard runs once per public call, before
'     materialization, so scalar and array calls share one host decision.
'   - Never raises. Every failure returns a value the caller can place at the
'     element's output position.
'
' ERROR POLICY
'   Registered conditions only, mapped through KPR_Core_Err.ErrForCondition.
'   The containment handler is a defect if reached.
'
' NOTES
'   - None.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim D               As Date     'Resolved DateIn
    Dim ErrOut          As Variant   'Element-level error value from a resolver

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: an element must never expose a runtime exception
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Resolve DateIn: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveDate(DateIn, D, ErrOut) Then Elem_EndOfYear = ErrOut: Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        Elem_EndOfYear = EndOfYear(D)
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed at this element; reaching here is a defect
        Elem_EndOfYear = ErrValue()

End Function

Private Function Elem_IsMonthEnd( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
'                              Elem_IsMonthEnd
'------------------------------------------------------------------------------
' PURPOSE
'   One element of KPR_Dates_IsMonthEnd: resolves the raw value arguments,
'   computes, and returns the Boolean or the element's native error value.
'
' ELEMENT CONTRACT
'   - Value arguments arrive raw (scalar, single-cell Range or 1x1 wrapper)
'     and are resolved here in signature order; the first failing argument
'     determines the result.
'   - Controls arrive already resolved, because a control is call-level and
'     is never re-resolved per element.
'   - Never calls the host guard: the guard runs once per public call, before
'     materialization, so scalar and array calls share one host decision.
'   - Never raises. Every failure returns a value the caller can place at the
'     element's output position.
'
' ERROR POLICY
'   Registered conditions only, mapped through KPR_Core_Err.ErrForCondition.
'   The containment handler is a defect if reached.
'
' NOTES
'   - None.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim D               As Date     'Resolved DateIn
    Dim ErrOut          As Variant   'Element-level error value from a resolver

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: an element must never expose a runtime exception
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Resolve DateIn: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveDate(DateIn, D, ErrOut) Then Elem_IsMonthEnd = ErrOut: Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        Elem_IsMonthEnd = (Day(D) = DaysInMonth(Year(D), Month(D)))
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed at this element; reaching here is a defect
        Elem_IsMonthEnd = ErrValue()

End Function

Private Function Elem_IsQuarterEnd( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
'                              Elem_IsQuarterEnd
'------------------------------------------------------------------------------
' PURPOSE
'   One element of KPR_Dates_IsQuarterEnd: resolves the raw value arguments,
'   computes, and returns the Boolean or the element's native error value.
'
' ELEMENT CONTRACT
'   - Value arguments arrive raw (scalar, single-cell Range or 1x1 wrapper)
'     and are resolved here in signature order; the first failing argument
'     determines the result.
'   - Controls arrive already resolved, because a control is call-level and
'     is never re-resolved per element.
'   - Never calls the host guard: the guard runs once per public call, before
'     materialization, so scalar and array calls share one host decision.
'   - Never raises. Every failure returns a value the caller can place at the
'     element's output position.
'
' ERROR POLICY
'   Registered conditions only, mapped through KPR_Core_Err.ErrForCondition.
'   The containment handler is a defect if reached.
'
' NOTES
'   - None.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim D               As Date     'Resolved DateIn
    Dim ErrOut          As Variant   'Element-level error value from a resolver

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: an element must never expose a runtime exception
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Resolve DateIn: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveDate(DateIn, D, ErrOut) Then Elem_IsQuarterEnd = ErrOut: Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        If (Month(D) Mod 3) <> 0 Then
            Elem_IsQuarterEnd = False
        Else
            Elem_IsQuarterEnd = (Day(D) = DaysInMonth(Year(D), Month(D)))
        End If
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed at this element; reaching here is a defect
        Elem_IsQuarterEnd = ErrValue()

End Function

Private Function Elem_IsYearEnd( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
'                              Elem_IsYearEnd
'------------------------------------------------------------------------------
' PURPOSE
'   One element of KPR_Dates_IsYearEnd: resolves the raw value arguments,
'   computes, and returns the Boolean or the element's native error value.
'
' ELEMENT CONTRACT
'   - Value arguments arrive raw (scalar, single-cell Range or 1x1 wrapper)
'     and are resolved here in signature order; the first failing argument
'     determines the result.
'   - Controls arrive already resolved, because a control is call-level and
'     is never re-resolved per element.
'   - Never calls the host guard: the guard runs once per public call, before
'     materialization, so scalar and array calls share one host decision.
'   - Never raises. Every failure returns a value the caller can place at the
'     element's output position.
'
' ERROR POLICY
'   Registered conditions only, mapped through KPR_Core_Err.ErrForCondition.
'   The containment handler is a defect if reached.
'
' NOTES
'   - None.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim D               As Date     'Resolved DateIn
    Dim ErrOut          As Variant   'Element-level error value from a resolver

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: an element must never expose a runtime exception
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Resolve DateIn: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveDate(DateIn, D, ErrOut) Then Elem_IsYearEnd = ErrOut: Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        Elem_IsYearEnd = ((Month(D) = 12) And (Day(D) = 31))
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed at this element; reaching here is a defect
        Elem_IsYearEnd = ErrValue()

End Function

Private Function Elem_IsLeapYear( _
    ByVal YearIn As Variant) _
    As Variant
'
'==============================================================================
'                              Elem_IsLeapYear
'------------------------------------------------------------------------------
' PURPOSE
'   One element of KPR_Dates_IsLeapYear: resolves the raw value arguments,
'   computes, and returns the Boolean or the element's native error value.
'
' ELEMENT CONTRACT
'   - Value arguments arrive raw (scalar, single-cell Range or 1x1 wrapper)
'     and are resolved here in signature order; the first failing argument
'     determines the result.
'   - Controls arrive already resolved, because a control is call-level and
'     is never re-resolved per element.
'   - Never calls the host guard: the guard runs once per public call, before
'     materialization, so scalar and array calls share one host decision.
'   - Never raises. Every failure returns a value the caller can place at the
'     element's output position.
'
' ERROR POLICY
'   Registered conditions only, mapped through KPR_Core_Err.ErrForCondition.
'   The containment handler is a defect if reached.
'
' NOTES
'   - Takes a calendar year, not a date. IsLeapYear(1900) is False; no window gate applies.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Y               As Long     'Resolved YearIn
    Dim ErrOut          As Variant   'Element-level error value from a resolver

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: an element must never expose a runtime exception
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Resolve YearIn: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveLong(YearIn, Y, ErrOut) Then Elem_IsLeapYear = ErrOut: Exit Function
    'Domain of YearIn is 1900 through 9999
        If (Y < Year(KPR_MIN_DATE)) Or (Y > Year(KPR_MAX_DATE)) Then Elem_IsLeapYear = ErrForCondition(KPR_COND_DOMAIN_YEAR): Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        Elem_IsLeapYear = IsLeapYear(Y)
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed at this element; reaching here is a defect
        Elem_IsLeapYear = ErrValue()

End Function

Private Function Elem_AddDays( _
    ByVal DateIn As Variant, _
    ByVal nDays As Variant) _
    As Variant
'
'==============================================================================
'                              Elem_AddDays
'------------------------------------------------------------------------------
' PURPOSE
'   One element of KPR_Dates_AddDays: resolves the raw value arguments,
'   computes, and returns the Date or the element's native error value.
'
' ELEMENT CONTRACT
'   - Value arguments arrive raw (scalar, single-cell Range or 1x1 wrapper)
'     and are resolved here in signature order; the first failing argument
'     determines the result.
'   - Controls arrive already resolved, because a control is call-level and
'     is never re-resolved per element.
'   - Never calls the host guard: the guard runs once per public call, before
'     materialization, so scalar and array calls share one host decision.
'   - Never raises. Every failure returns a value the caller can place at the
'     element's output position.
'
' ERROR POLICY
'   Registered conditions only, mapped through KPR_Core_Err.ErrForCondition.
'   The containment handler is a defect if reached.
'
' NOTES
'   - The shift is computed in Double and gated before coercion back to Date, so an overflow can never reach CDate.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim D               As Date     'Resolved DateIn
    Dim N               As Long     'Resolved nDays
    Dim ResultD         As Double    'Shifted serial, gated before conversion
    Dim ErrOut          As Variant   'Element-level error value from a resolver

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: an element must never expose a runtime exception
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Resolve DateIn: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveDate(DateIn, D, ErrOut) Then Elem_AddDays = ErrOut: Exit Function
    'Resolve nDays: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveLong(nDays, N, ErrOut) Then Elem_AddDays = ErrOut: Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        ResultD = CDbl(D) + CDbl(N)
        If (ResultD < CDbl(KPR_MIN_DATE)) Or (ResultD > CDbl(KPR_MAX_DATE)) Then
            Elem_AddDays = ErrForCondition(KPR_COND_RESULT_WINDOW): Exit Function
        End If
        Elem_AddDays = CDate(ResultD)
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed at this element; reaching here is a defect
        Elem_AddDays = ErrValue()

End Function

Private Function Elem_AddWeeks( _
    ByVal DateIn As Variant, _
    ByVal nWeeks As Variant) _
    As Variant
'
'==============================================================================
'                              Elem_AddWeeks
'------------------------------------------------------------------------------
' PURPOSE
'   One element of KPR_Dates_AddWeeks: resolves the raw value arguments,
'   computes, and returns the Date or the element's native error value.
'
' ELEMENT CONTRACT
'   - Value arguments arrive raw (scalar, single-cell Range or 1x1 wrapper)
'     and are resolved here in signature order; the first failing argument
'     determines the result.
'   - Controls arrive already resolved, because a control is call-level and
'     is never re-resolved per element.
'   - Never calls the host guard: the guard runs once per public call, before
'     materialization, so scalar and array calls share one host decision.
'   - Never raises. Every failure returns a value the caller can place at the
'     element's output position.
'
' ERROR POLICY
'   Registered conditions only, mapped through KPR_Core_Err.ErrForCondition.
'   The containment handler is a defect if reached.
'
' NOTES
'   - None.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim D               As Date     'Resolved DateIn
    Dim N               As Long     'Resolved nWeeks
    Dim ResultD         As Double    'Shifted serial, gated before conversion
    Dim ErrOut          As Variant   'Element-level error value from a resolver

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: an element must never expose a runtime exception
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Resolve DateIn: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveDate(DateIn, D, ErrOut) Then Elem_AddWeeks = ErrOut: Exit Function
    'Resolve nWeeks: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveLong(nWeeks, N, ErrOut) Then Elem_AddWeeks = ErrOut: Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        ResultD = CDbl(D) + (7# * CDbl(N))
        If (ResultD < CDbl(KPR_MIN_DATE)) Or (ResultD > CDbl(KPR_MAX_DATE)) Then
            Elem_AddWeeks = ErrForCondition(KPR_COND_RESULT_WINDOW): Exit Function
        End If
        Elem_AddWeeks = CDate(ResultD)
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed at this element; reaching here is a defect
        Elem_AddWeeks = ErrValue()

End Function

Private Function Elem_AddMonths( _
    ByVal DateIn As Variant, _
    ByVal nMonths As Variant, _
    ByVal KeepEOM As Boolean) _
    As Variant
'
'==============================================================================
'                              Elem_AddMonths
'------------------------------------------------------------------------------
' PURPOSE
'   One element of KPR_Dates_AddMonths: resolves the raw value arguments,
'   computes, and returns the Date or the element's native error value.
'
' ELEMENT CONTRACT
'   - Value arguments arrive raw (scalar, single-cell Range or 1x1 wrapper)
'     and are resolved here in signature order; the first failing argument
'     determines the result.
'   - Controls arrive already resolved, because a control is call-level and
'     is never re-resolved per element.
'   - Never calls the host guard: the guard runs once per public call, before
'     materialization, so scalar and array calls share one host decision.
'   - Never raises. Every failure returns a value the caller can place at the
'     element's output position.
'
' ERROR POLICY
'   Registered conditions only, mapped through KPR_Core_Err.ErrForCondition.
'   The containment handler is a defect if reached.
'
' NOTES
'   - Clip mode keeps the day when the target month has it and otherwise uses that month's last day. With Opt_KeepEOM=True an EOM input maps to the target EOM; a non-EOM input still clips.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim D               As Date     'Resolved DateIn
    Dim N               As Long     'Resolved nMonths
    Dim R               As Date      'Shifted date from the calendar core
    Dim ErrOut          As Variant   'Element-level error value from a resolver

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: an element must never expose a runtime exception
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Resolve DateIn: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveDate(DateIn, D, ErrOut) Then Elem_AddMonths = ErrOut: Exit Function
    'Resolve nMonths: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveLong(nMonths, N, ErrOut) Then Elem_AddMonths = ErrOut: Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        If Not TryAddMonths(D, N, KeepEOM, R) Then
            Elem_AddMonths = ErrForCondition(KPR_COND_RESULT_WINDOW): Exit Function
        End If
        Elem_AddMonths = R
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed at this element; reaching here is a defect
        Elem_AddMonths = ErrValue()

End Function

Private Function Elem_AddYears( _
    ByVal DateIn As Variant, _
    ByVal nYears As Variant, _
    ByVal KeepEOM As Boolean) _
    As Variant
'
'==============================================================================
'                              Elem_AddYears
'------------------------------------------------------------------------------
' PURPOSE
'   One element of KPR_Dates_AddYears: resolves the raw value arguments,
'   computes, and returns the Date or the element's native error value.
'
' ELEMENT CONTRACT
'   - Value arguments arrive raw (scalar, single-cell Range or 1x1 wrapper)
'     and are resolved here in signature order; the first failing argument
'     determines the result.
'   - Controls arrive already resolved, because a control is call-level and
'     is never re-resolved per element.
'   - Never calls the host guard: the guard runs once per public call, before
'     materialization, so scalar and array calls share one host decision.
'   - Never raises. Every failure returns a value the caller can place at the
'     element's output position.
'
' ERROR POLICY
'   Registered conditions only, mapped through KPR_Core_Err.ErrForCondition.
'   The containment handler is a defect if reached.
'
' NOTES
'   - Delegates to TryAddMonths so 29-Feb, short-month and EOM behaviour cannot diverge from AddMonths.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim D               As Date     'Resolved DateIn
    Dim N               As Long     'Resolved nYears
    Dim MonthsD         As Double    'Year count as months, range-gated before CLng
    Dim R               As Date      'Shifted date from the calendar core
    Dim ErrOut          As Variant   'Element-level error value from a resolver

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: an element must never expose a runtime exception
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Resolve DateIn: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveDate(DateIn, D, ErrOut) Then Elem_AddYears = ErrOut: Exit Function
    'Resolve nYears: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveLong(nYears, N, ErrOut) Then Elem_AddYears = ErrOut: Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        MonthsD = 12# * CDbl(N)
        If (MonthsD < -2147483648#) Or (MonthsD > 2147483647#) Then
            Elem_AddYears = ErrForCondition(KPR_COND_RESULT_WINDOW): Exit Function
        End If
        If Not TryAddMonths(D, CLng(MonthsD), KeepEOM, R) Then
            Elem_AddYears = ErrForCondition(KPR_COND_RESULT_WINDOW): Exit Function
        End If
        Elem_AddYears = R
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed at this element; reaching here is a defect
        Elem_AddYears = ErrValue()

End Function

Private Function Elem_NthWeekdayOfMonth( _
    ByVal YearIn As Variant, _
    ByVal MonthIn As Variant, _
    ByVal WdIndex As Variant, _
    ByVal n As Variant, _
    ByVal WkMonday As Boolean) _
    As Variant
'
'==============================================================================
'                              Elem_NthWeekdayOfMonth
'------------------------------------------------------------------------------
' PURPOSE
'   One element of KPR_Dates_NthWeekdayOfMonth: resolves the raw value arguments,
'   computes, and returns the Date or the element's native error value.
'
' ELEMENT CONTRACT
'   - Value arguments arrive raw (scalar, single-cell Range or 1x1 wrapper)
'     and are resolved here in signature order; the first failing argument
'     determines the result.
'   - Controls arrive already resolved, because a control is call-level and
'     is never re-resolved per element.
'   - Never calls the host guard: the guard runs once per public call, before
'     materialization, so scalar and array calls share one host decision.
'   - Never raises. Every failure returns a value the caller can place at the
'     element's output position.
'
' ERROR POLICY
'   Registered conditions only, mapped through KPR_Core_Err.ErrForCondition.
'   The containment handler is a defect if reached.
'
' NOTES
'   - A fifth occurrence that the month does not contain is OCCURRENCE_ABSENT; a valid locator before 1900-03-01 is RESULT_WINDOW.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Y               As Long     'Resolved YearIn
    Dim M               As Long     'Resolved MonthIn
    Dim W               As Long     'Resolved WdIndex
    Dim K               As Long     'Resolved n
    Dim WkBase          As VbDayOfWeek 'Weekday base selected by the control
    Dim D0              As Date      'First day of the requested month
    Dim D0Serial        As Double    'Its serial
    Dim MonthLen        As Long      'Length of the requested month
    Dim Off             As Long      'Days from D0 to the first matching weekday
    Dim OutSerial       As Double    'Candidate serial, gated before conversion
    Dim ErrOut          As Variant   'Element-level error value from a resolver

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: an element must never expose a runtime exception
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Resolve YearIn: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveLong(YearIn, Y, ErrOut) Then Elem_NthWeekdayOfMonth = ErrOut: Exit Function
    'Domain of YearIn is 1900 through 9999
        If (Y < Year(KPR_MIN_DATE)) Or (Y > Year(KPR_MAX_DATE)) Then Elem_NthWeekdayOfMonth = ErrForCondition(KPR_COND_DOMAIN_YEAR): Exit Function
    'Resolve MonthIn: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveLong(MonthIn, M, ErrOut) Then Elem_NthWeekdayOfMonth = ErrOut: Exit Function
    'Domain of MonthIn is 1 through 12
        If (M < 1) Or (M > 12) Then Elem_NthWeekdayOfMonth = ErrForCondition(KPR_COND_DOMAIN_MONTH): Exit Function
    'Resolve WdIndex: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveLong(WdIndex, W, ErrOut) Then Elem_NthWeekdayOfMonth = ErrOut: Exit Function
    'Domain of WdIndex is 1 through 7
        If (W < 1) Or (W > 7) Then Elem_NthWeekdayOfMonth = ErrForCondition(KPR_COND_DOMAIN_WEEKDAY): Exit Function
    'Resolve n: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveLong(n, K, ErrOut) Then Elem_NthWeekdayOfMonth = ErrOut: Exit Function
    'Domain of n is 1 through 5
        If (K < 1) Or (K > 5) Then Elem_NthWeekdayOfMonth = ErrForCondition(KPR_COND_DOMAIN_OCCURRENCE): Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        If WkMonday Then WkBase = vbMonday Else WkBase = vbSunday
        D0 = DateSerial(Y, M, 1)
        D0Serial = CDbl(D0)
        MonthLen = DaysInMonth(Y, M)
        Off = (W - Weekday(D0, WkBase) + 7) Mod 7
        OutSerial = D0Serial + CDbl(Off) + (7# * CDbl(K - 1))
        If OutSerial > (D0Serial + CDbl(MonthLen) - 1#) Then
            Elem_NthWeekdayOfMonth = ErrForCondition(KPR_COND_OCCURRENCE_ABSENT): Exit Function
        End If
        If (OutSerial < CDbl(KPR_MIN_DATE)) Or (OutSerial > CDbl(KPR_MAX_DATE)) Then
            Elem_NthWeekdayOfMonth = ErrForCondition(KPR_COND_RESULT_WINDOW): Exit Function
        End If
        Elem_NthWeekdayOfMonth = CDate(OutSerial)
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed at this element; reaching here is a defect
        Elem_NthWeekdayOfMonth = ErrValue()

End Function

Private Function Elem_LastWeekdayOfMonth( _
    ByVal YearIn As Variant, _
    ByVal MonthIn As Variant, _
    ByVal WdIndex As Variant, _
    ByVal WkMonday As Boolean) _
    As Variant
'
'==============================================================================
'                              Elem_LastWeekdayOfMonth
'------------------------------------------------------------------------------
' PURPOSE
'   One element of KPR_Dates_LastWeekdayOfMonth: resolves the raw value arguments,
'   computes, and returns the Date or the element's native error value.
'
' ELEMENT CONTRACT
'   - Value arguments arrive raw (scalar, single-cell Range or 1x1 wrapper)
'     and are resolved here in signature order; the first failing argument
'     determines the result.
'   - Controls arrive already resolved, because a control is call-level and
'     is never re-resolved per element.
'   - Never calls the host guard: the guard runs once per public call, before
'     materialization, so scalar and array calls share one host decision.
'   - Never raises. Every failure returns a value the caller can place at the
'     element's output position.
'
' ERROR POLICY
'   Registered conditions only, mapped through KPR_Core_Err.ErrForCondition.
'   The containment handler is a defect if reached.
'
' NOTES
'   - None.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Y               As Long     'Resolved YearIn
    Dim M               As Long     'Resolved MonthIn
    Dim W               As Long     'Resolved WdIndex
    Dim WkBase          As VbDayOfWeek 'Weekday base selected by the control
    Dim EOM             As Date      'Last day of the requested month
    Dim Diff            As Long      'Days back from EOM to the weekday
    Dim OutSerial       As Double    'Candidate serial, gated before conversion
    Dim ErrOut          As Variant   'Element-level error value from a resolver

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: an element must never expose a runtime exception
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Resolve YearIn: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveLong(YearIn, Y, ErrOut) Then Elem_LastWeekdayOfMonth = ErrOut: Exit Function
    'Domain of YearIn is 1900 through 9999
        If (Y < Year(KPR_MIN_DATE)) Or (Y > Year(KPR_MAX_DATE)) Then Elem_LastWeekdayOfMonth = ErrForCondition(KPR_COND_DOMAIN_YEAR): Exit Function
    'Resolve MonthIn: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveLong(MonthIn, M, ErrOut) Then Elem_LastWeekdayOfMonth = ErrOut: Exit Function
    'Domain of MonthIn is 1 through 12
        If (M < 1) Or (M > 12) Then Elem_LastWeekdayOfMonth = ErrForCondition(KPR_COND_DOMAIN_MONTH): Exit Function
    'Resolve WdIndex: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveLong(WdIndex, W, ErrOut) Then Elem_LastWeekdayOfMonth = ErrOut: Exit Function
    'Domain of WdIndex is 1 through 7
        If (W < 1) Or (W > 7) Then Elem_LastWeekdayOfMonth = ErrForCondition(KPR_COND_DOMAIN_WEEKDAY): Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        If WkMonday Then WkBase = vbMonday Else WkBase = vbSunday
        EOM = DateSerial(Y, M, DaysInMonth(Y, M))
        Diff = (Weekday(EOM, WkBase) - W + 7) Mod 7
        OutSerial = CDbl(EOM) - CDbl(Diff)
        If (OutSerial < CDbl(KPR_MIN_DATE)) Or (OutSerial > CDbl(KPR_MAX_DATE)) Then
            Elem_LastWeekdayOfMonth = ErrForCondition(KPR_COND_RESULT_WINDOW): Exit Function
        End If
        Elem_LastWeekdayOfMonth = CDate(OutSerial)
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed at this element; reaching here is a defect
        Elem_LastWeekdayOfMonth = ErrValue()

End Function

Private Function Elem_PillarFromDates( _
    ByVal StartDate As Variant, _
    ByVal EndDate As Variant, _
    ByVal Mode As KPR_PillarRounding) _
    As Variant
'
'==============================================================================
'                              Elem_PillarFromDates
'------------------------------------------------------------------------------
' PURPOSE
'   One element of KPR_Dates_PillarFromDates: resolves the raw value arguments,
'   computes, and returns the String or the element's native error value.
'
' ELEMENT CONTRACT
'   - Value arguments arrive raw (scalar, single-cell Range or 1x1 wrapper)
'     and are resolved here in signature order; the first failing argument
'     determines the result.
'   - Controls arrive already resolved, because a control is call-level and
'     is never re-resolved per element.
'   - Never calls the host guard: the guard runs once per public call, before
'     materialization, so scalar and array calls share one host decision.
'   - Never raises. Every failure returns a value the caller can place at the
'     element's output position.
'
' ERROR POLICY
'   Registered conditions only, mapped through KPR_Core_Err.ErrForCondition.
'   The containment handler is a defect if reached.
'
' NOTES
'   - Rounding modes, the candidate set and tie rules are specified in contract section 8.4.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim S               As Date     'Resolved StartDate
    Dim E               As Date     'Resolved EndDate
    Dim Token           As String    'Formatted pillar token
    Dim Condition       As KPR_Condition 'Formatter failure condition
    Dim ErrOut          As Variant   'Element-level error value from a resolver

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: an element must never expose a runtime exception
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Resolve StartDate: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveDate(StartDate, S, ErrOut) Then Elem_PillarFromDates = ErrOut: Exit Function
    'Resolve EndDate: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveDate(EndDate, E, ErrOut) Then Elem_PillarFromDates = ErrOut: Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        If Not TryPillar_Format(S, E, Mode, Token, Condition) Then
            Elem_PillarFromDates = ErrForCondition(Condition): Exit Function
        End If
        Elem_PillarFromDates = Token
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed at this element; reaching here is a defect
        Elem_PillarFromDates = ErrValue()

End Function

Private Function Elem_DateFromPillar( _
    ByVal StartDate As Variant, _
    ByVal Pillar As Variant) _
    As Variant
'
'==============================================================================
'                              Elem_DateFromPillar
'------------------------------------------------------------------------------
' PURPOSE
'   One element of KPR_Dates_DateFromPillar: resolves the raw value arguments,
'   computes, and returns the Date or the element's native error value.
'
' ELEMENT CONTRACT
'   - Value arguments arrive raw (scalar, single-cell Range or 1x1 wrapper)
'     and are resolved here in signature order; the first failing argument
'     determines the result.
'   - Controls arrive already resolved, because a control is call-level and
'     is never re-resolved per element.
'   - Never calls the host guard: the guard runs once per public call, before
'     materialization, so scalar and array calls share one host decision.
'   - Never raises. Every failure returns a value the caller can place at the
'     element's output position.
'
' ERROR POLICY
'   Registered conditions only, mapped through KPR_Core_Err.ErrForCondition.
'   The containment handler is a defect if reached.
'
' NOTES
'   - Months first with clip semantics, then exact days, matching the parser's grammar. The name is singular because the function returns exactly one date per element.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim S               As Date     'Resolved StartDate
    Dim PMonths         As Double    'Parsed month delta of Pillar
    Dim PDays           As Double    'Parsed day delta of Pillar
    Dim WorkDate        As Date      'Date after the month delta
    Dim ResultD         As Double    'Final serial, gated before conversion
    Dim ErrOut          As Variant   'Element-level error value from a resolver

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Containment only: an element must never expose a runtime exception
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Resolve StartDate: unwrap, propagate an incoming error, parse strictly
        If Not TryResolveDate(StartDate, S, ErrOut) Then Elem_DateFromPillar = ErrOut: Exit Function
    'Resolve Pillar: unwrap, propagate an incoming error, parse the grammar
        If Not TryResolvePillar(Pillar, PMonths, PDays, ErrOut) Then Elem_DateFromPillar = ErrOut: Exit Function

'------------------------------------------------------------------------------
' COMPUTE
'------------------------------------------------------------------------------
        WorkDate = S
        If PMonths <> 0# Then
            If (PMonths < -2147483648#) Or (PMonths > 2147483647#) Then
                Elem_DateFromPillar = ErrForCondition(KPR_COND_PILLAR_AGGREGATE_RANGE): Exit Function
            End If
            If Not TryAddMonths(WorkDate, CLng(PMonths), False, WorkDate) Then
                Elem_DateFromPillar = ErrForCondition(KPR_COND_RESULT_WINDOW): Exit Function
            End If
        End If
        ResultD = CDbl(WorkDate) + PDays
        If (ResultD < CDbl(KPR_MIN_DATE)) Or (ResultD > CDbl(KPR_MAX_DATE)) Then
            Elem_DateFromPillar = ErrForCondition(KPR_COND_RESULT_WINDOW): Exit Function
        End If
        Elem_DateFromPillar = CDate(ResultD)
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Containment: fail closed at this element; reaching here is a defect
        Elem_DateFromPillar = ErrValue()

End Function

'
'------------------------------------------------------------------------------
'
'                          PUBLIC API - HOST DIAGNOSTIC
'
'------------------------------------------------------------------------------
'

Public Function KPR_Dates_HostDateSystem() As Variant
'
'==============================================================================
'                           KPR_Dates_HostDateSystem
'------------------------------------------------------------------------------
' PURPOSE
'   Reports the date system the library is answering under for this caller:
'   1900, 1904, or #N/A when an identified worksheet host cannot be read.
'
' WHY THIS EXISTS
'   A library-produced host-configuration #N/A and a propagated incoming #N/A
'   are the same Excel value. This diagnostic supplies the caller context that
'   the returned value cannot: 1904 identifies host refusal, 1900 leaves an
'   incoming error or another documented input path as the source.
'
' SIGNATURE
'   KPR_Dates_HostDateSystem() -> Variant
'
' RETURNS
'   Variant
'     1900  identified worksheet caller in a 1900 workbook, OR no worksheet
'           host could be identified (direct VBA, Immediate window,
'           Application.Run, the regression harness)
'     1904  identified worksheet caller in a 1904 workbook
'     #N/A  an identified worksheet Range whose workbook date system could
'           not be read reliably (HOST_UNRESOLVED)
'
' VOLATILITY
'   Application.Volatile True is the first executable statement. A
'   zero-argument function has no precedents, so without it Excel would
'   evaluate this cell once on entry and never again, and a toggled date
'   system would leave it reporting a stale answer at the one moment it is
'   being relied on. Only this diagnostic is volatile; every date calculation
'   stays non-volatile, and the static gate enforces both halves of that.
'
'   Volatility is contagious: any cell that references this one becomes
'   volatile too. Use it as a debugging aid in a cell or two, not as a
'   building block fanned across a model.
'
' ERROR POLICY (USER FACING)
'   #N/A     identified worksheet host whose date system cannot be read
'   The containment handler is not a documented condition; reaching it is a
'   defect.
'
' DEPENDENCIES
'   - TryResolveHostDateSystem
'   - KPR_Core_Err.ErrForCondition
'
' NOTES
'   - This function calls the classifier directly rather than PassHostGuard,
'     because an identified 1904 workbook must be REPORTED as 1904 here, not
'     refused with #N/A as the value functions do.
'   - Scalar only: it has no value argument to vectorize over.
'
' UPDATED
'   2026-09-01
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim DateSystem      As Long             'Classified host date system
    Dim Condition       As KPR_Condition    'Failure condition, if any

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Re-evaluate on every ordinary recalculation; this must be the first statement
        Application.Volatile True

    'Containment only: a raise reaching the handler is a defect, never an outcome
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' CLASSIFY HOST
'------------------------------------------------------------------------------
    'Report rather than refuse: 1904 is an answer here
        If Not TryResolveHostDateSystem(DateSystem, Condition) Then
            KPR_Dates_HostDateSystem = ErrForCondition(Condition)
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the year that names the date system
        KPR_Dates_HostDateSystem = DateSystem
        Exit Function

'------------------------------------------------------------------------------
' ERR_HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Unexpected runtime error => #VALUE! (containment only)
        KPR_Dates_HostDateSystem = ErrValue()

End Function

'
'------------------------------------------------------------------------------
'
'                              PRIVATE COMPOSITION
'
'------------------------------------------------------------------------------
'

Private Function TryResolveDate( _
    ByVal DateIn As Variant, _
    ByRef ParsedDate As Date, _
    ByRef ErrOut As Variant) _
    As Boolean
'
'==============================================================================
'                                TryResolveDate
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves one worksheet date argument into a normalized, in-window VBA Date,
'   or assigns the exact worksheet error the contract requires.
'
' SIGNATURE
'   TryResolveDate(DateIn, ParsedDate, ErrOut) -> Boolean
'
' OUTPUTS
'   ParsedDate (ByRef)   assigned ONLY on success
'   ErrOut     (ByRef)   assigned ONLY on failure, to the value to return
'
' BEHAVIOR
'   1. KPR_Core_Array.TryUnwrapScalar reduces wrappers to one scalar. Failure is
'      SHAPE_UNSUPPORTED.
'   2. An incoming native Excel error is returned VERBATIM. This is the
'      public/element propagation boundary: the error never reaches the parser
'      and is never collapsed into an ordinary parse failure.
'   3. KPR_Core_Parse.TryParseDateScalar classifies and window-gates the value.
'   4. The condition is mapped through KPR_Core_Err.ErrForCondition, so a
'      malformed value yields #VALUE! and an out-of-window value yields #NUM!.
'
' NOTES
'   - Propagation is deliberately not routed through ErrForCondition. The value
'     to return is the incoming error itself, not an error chosen for a
'     condition, and ErrForCondition raises rather than mapping propagation.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ScalarValue     As Variant          'Payload after shape unwrapping
    Dim Candidate       As Date             'Parsed candidate
    Dim Condition       As KPR_Condition    'Classified failure condition

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default return is failure unless every stage succeeds
        TryResolveDate = False

'------------------------------------------------------------------------------
' RESOLVE SHAPE
'------------------------------------------------------------------------------
    'Reduce a single-cell Range or 1x1 wrapper to its scalar payload
        If Not TryUnwrapScalar(DateIn, ScalarValue) Then
            ErrOut = ErrForCondition(KPR_COND_SHAPE_UNSUPPORTED)
            Exit Function
        End If

'------------------------------------------------------------------------------
' PROPAGATE INCOMING ERRORS
'------------------------------------------------------------------------------
    'An incoming Excel error is the caller's answer, returned unchanged
        If VarType(ScalarValue) = vbError Then
            ErrOut = ScalarValue
            Exit Function
        End If

'------------------------------------------------------------------------------
' RESOLVE VALUE
'------------------------------------------------------------------------------
    'Classify strictly and apply the supported window
        If Not TryParseDateScalar(ScalarValue, Candidate, Condition) Then
            ErrOut = ErrForCondition(Condition)
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once the value is fully validated
        ParsedDate = Candidate

    'Contract: TRUE only when ParsedDate was assigned
        TryResolveDate = True

End Function

Private Function TryResolveLong( _
    ByVal ValueIn As Variant, _
    ByRef ParsedLong As Long, _
    ByRef ErrOut As Variant) _
    As Boolean
'
'==============================================================================
'                                TryResolveLong
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves one worksheet integer argument into a Long, or assigns the exact
'   worksheet error the contract requires.
'
' BEHAVIOR
'   Same four stages as TryResolveDate, with KPR_Core_Parse.TryParseLongScalar
'   doing the classification. A fraction yields #VALUE! under INTEGER_FRACTION
'   and a value outside the Long range yields #NUM! under INTEGER_RANGE.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ScalarValue     As Variant          'Payload after shape unwrapping
    Dim Candidate       As Long             'Parsed candidate
    Dim Condition       As KPR_Condition    'Classified failure condition

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default return is failure unless every stage succeeds
        TryResolveLong = False

'------------------------------------------------------------------------------
' RESOLVE SHAPE
'------------------------------------------------------------------------------
    'Reduce a single-cell Range or 1x1 wrapper to its scalar payload
        If Not TryUnwrapScalar(ValueIn, ScalarValue) Then
            ErrOut = ErrForCondition(KPR_COND_SHAPE_UNSUPPORTED)
            Exit Function
        End If

'------------------------------------------------------------------------------
' PROPAGATE INCOMING ERRORS
'------------------------------------------------------------------------------
    'An incoming Excel error is the caller's answer, returned unchanged
        If VarType(ScalarValue) = vbError Then
            ErrOut = ScalarValue
            Exit Function
        End If

'------------------------------------------------------------------------------
' RESOLVE VALUE
'------------------------------------------------------------------------------
    'Classify strictly: no truncation, no rounding, no Boolean coercion
        If Not TryParseLongScalar(ScalarValue, Candidate, Condition) Then
            ErrOut = ErrForCondition(Condition)
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once the value is fully validated
        ParsedLong = Candidate

    'Contract: TRUE only when ParsedLong was assigned
        TryResolveLong = True

End Function

Private Function TryResolveBool( _
    ByVal ControlIn As Variant, _
    ByVal DefaultValue As Boolean, _
    ByRef ParsedBool As Boolean, _
    ByRef ErrOut As Variant) _
    As Boolean
'
'==============================================================================
'                                TryResolveBool
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves one optional Boolean control, or assigns the worksheet error the
'   contract requires.
'
' BEHAVIOR
'   - Each public Optional Variant declaration supplies its documented default,
'     so an omitted argument reaches this helper as a native Boolean.
'   - A blank cell arrives as Empty and selects the same default.
'   - A control larger than 1x1 is CONTROL_NOT_SCALAR, reported by
'     TryUnwrapControl from the control's dimensions alone.
'   - An incoming error in a control propagates verbatim as a call-level result.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ScalarValue     As Variant          'Payload after shape unwrapping
    Dim Candidate       As Boolean          'Parsed candidate
    Dim Condition       As KPR_Condition    'Classified failure condition

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default return is failure unless every stage succeeds
        TryResolveBool = False

'------------------------------------------------------------------------------
' RESOLVE SHAPE
'------------------------------------------------------------------------------
    'A control must be scalar or a 1x1 wrapper; it never vectorizes. A valid
    'multi-element control is CONTROL_NOT_SCALAR, an unsupported shape is
    'SHAPE_UNSUPPORTED, and neither reads the control's contents.
        If Not TryUnwrapControl(ControlIn, ScalarValue, Condition) Then
            ErrOut = ErrForCondition(Condition)
            Exit Function
        End If

'------------------------------------------------------------------------------
' PROPAGATE INCOMING ERRORS
'------------------------------------------------------------------------------
    'An incoming Excel error is the caller's answer, returned unchanged
        If VarType(ScalarValue) = vbError Then
            ErrOut = ScalarValue
            Exit Function
        End If

'------------------------------------------------------------------------------
' RESOLVE VALUE
'------------------------------------------------------------------------------
    'Accept a native Boolean only; Empty selects the documented default
        If Not TryParseBoolControl(ScalarValue, DefaultValue, Candidate, Condition) Then
            ErrOut = ErrForCondition(Condition)
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once the control is validated
        ParsedBool = Candidate

    'Contract: TRUE only when ParsedBool was assigned
        TryResolveBool = True

End Function

Private Function TryResolveRounding( _
    ByVal ControlIn As Variant, _
    ByRef Mode As KPR_PillarRounding, _
    ByRef ErrOut As Variant) _
    As Boolean
'
'==============================================================================
'                              TryResolveRounding
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves Opt_Rounding to a KPR_PillarRounding mode, or assigns the
'   call-level error the contract requires.
'
' BEHAVIOR
'   - Omitted selects NEAREST here; Empty selects it in the parser.
'   - Shape, propagation and parsing follow the same four stages as the other
'     resolvers. The parser returns normalized text because it may not depend
'     on KPR_Core_Dates; the mapping to the enum happens here, where both
'     modules are reachable.
'
' UPDATED
'   2026-09-01
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ScalarValue     As Variant          'Payload after shape unwrapping
    Dim Token           As String           'Normalized token from the parser
    Dim Condition       As KPR_Condition    'Classified failure condition

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default return is failure unless every stage succeeds
        TryResolveRounding = False

'------------------------------------------------------------------------------
' RESOLVE OMISSION
'------------------------------------------------------------------------------
    'An omitted control selects the documented default
        If IsMissing(ControlIn) Then
            Mode = KPR_ROUND_NEAREST
            TryResolveRounding = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' RESOLVE SHAPE
'------------------------------------------------------------------------------
    'A control must be scalar or a 1x1 wrapper; it never vectorizes. A valid
    'multi-element control is CONTROL_NOT_SCALAR, an unsupported shape is
    'SHAPE_UNSUPPORTED, and neither reads the control's contents.
        If Not TryUnwrapControl(ControlIn, ScalarValue, Condition) Then
            ErrOut = ErrForCondition(Condition)
            Exit Function
        End If

'------------------------------------------------------------------------------
' PROPAGATE INCOMING ERRORS
'------------------------------------------------------------------------------
    'An incoming Excel error is the caller's answer, returned unchanged
        If VarType(ScalarValue) = vbError Then
            ErrOut = ScalarValue
            Exit Function
        End If

'------------------------------------------------------------------------------
' RESOLVE VALUE
'------------------------------------------------------------------------------
    'Locale-independent normalization and matching in the parser
        If Not TryParseRoundingControl(ScalarValue, Token, Condition) Then
            ErrOut = ErrForCondition(Condition)
            Exit Function
        End If

'------------------------------------------------------------------------------
' MAP TO MODE
'------------------------------------------------------------------------------
    'The parser guarantees one of exactly three tokens
        Select Case Token
            Case "FLOOR":   Mode = KPR_ROUND_FLOOR
            Case "CEILING": Mode = KPR_ROUND_CEILING
            Case Else:      Mode = KPR_ROUND_NEAREST
        End Select

    'Contract: TRUE only when Mode was assigned
        TryResolveRounding = True

End Function

Private Function TryResolvePillar( _
    ByVal PillarIn As Variant, _
    ByRef TotalMonths As Double, _
    ByRef TotalDays As Double, _
    ByRef ErrOut As Variant) _
    As Boolean
'
'==============================================================================
'                               TryResolvePillar
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves one worksheet pillar argument into signed month and day deltas,
'   or assigns the worksheet error the contract requires.
'
' BEHAVIOR
'   Same four stages as TryResolveDate. Before this resolver existed, an
'   incoming Excel error at the Pillar position reached the parser and was
'   rejected as a non-text payload; the contract requires it to propagate
'   verbatim, and now it does.
'
' UPDATED
'   2026-09-01
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ScalarValue     As Variant          'Payload after shape unwrapping
    Dim Months          As Double           'Parsed month delta
    Dim Days            As Double           'Parsed day delta
    Dim Condition       As KPR_Condition    'Classified failure condition

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default return is failure unless every stage succeeds
        TryResolvePillar = False

'------------------------------------------------------------------------------
' RESOLVE SHAPE
'------------------------------------------------------------------------------
    'Reduce a single-cell Range or 1x1 wrapper to its scalar payload
        If Not TryUnwrapScalar(PillarIn, ScalarValue) Then
            ErrOut = ErrForCondition(KPR_COND_SHAPE_UNSUPPORTED)
            Exit Function
        End If

'------------------------------------------------------------------------------
' PROPAGATE INCOMING ERRORS
'------------------------------------------------------------------------------
    'An incoming Excel error is the caller's answer, returned unchanged
        If VarType(ScalarValue) = vbError Then
            ErrOut = ScalarValue
            Exit Function
        End If

'------------------------------------------------------------------------------
' RESOLVE VALUE
'------------------------------------------------------------------------------
    'The grammar and its classified rejections live in the calendar core
        If Not TryPillar_Parse(ScalarValue, Months, Days, Condition) Then
            ErrOut = ErrForCondition(Condition)
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the outputs only once the token is fully validated
        TotalMonths = Months
        TotalDays = Days

    'Contract: TRUE only when the outputs were assigned
        TryResolvePillar = True

End Function

'
'------------------------------------------------------------------------------
'
'                            PRIVATE HOST RESOLUTION
'
'------------------------------------------------------------------------------
'

Private Function TryResolveHostDateSystem( _
    ByRef DateSystem As Long, _
    ByRef Condition As KPR_Condition) _
    As Boolean
'
'==============================================================================
'                          TryResolveHostDateSystem
'------------------------------------------------------------------------------
' PURPOSE
'   Classifies the caller of the current public call and reports the date
'   system the library answers under.
'
' SIGNATURE
'   TryResolveHostDateSystem(DateSystem, Condition) -> Boolean
'
' OUTPUTS
'   DateSystem (ByRef)   1900 or 1904, assigned ONLY on success
'   Condition  (ByRef)   KPR_COND_NONE on success, else KPR_COND_HOST_UNRESOLVED
'
' RETURNS
'   Boolean
'     TRUE  => DateSystem assigned
'     FALSE => a worksheet Range was identified but its workbook date system
'              could not be read reliably
'
' CALLER CONTRACT
'   1. Application.Caller is read ONCE through a single guarded object
'      assignment. A worksheet Range is the only caller form that identifies a
'      worksheet host.
'   2. For a Range, the workbook is reached through the caller's own
'      Worksheet.Parent. ActiveWorkbook, ThisWorkbook and ActiveSheet are never
'      consulted: the caller's workbook is the only authority on its date
'      system, and any other workbook would be an unrelated one.
'   3. If the Range was identified but reading its date system raises, the
'      result is HOST_UNRESOLVED. This is the ONLY path that fails.
'   4. Every other outcome, including a non-object caller, a non-Range object,
'      and a failure while obtaining or classifying Application.Caller itself,
'      means "no worksheet host could be identified". The documented 1900
'      serial contract applies and the caller owns the interpretation of the
'      values it constructed.
'
' WHAT "NO WORKSHEET HOST" DOES NOT MEAN
'   It does not mean the caller is proven direct VBA. Direct VBA, the
'   Immediate window, Application.Run and the regression harness are the
'   CERTIFIED uses of this path. Excel also exposes non-Range callers in other
'   contexts, and v0.0.2 makes no compatibility claim for any of them:
'       - a String naming a macro-attached shape
'       - an Error value from the Immediate window and from Application.Run
'       - Error or other non-Range forms in data-validation, chart-series,
'         conditional-formatting and defined-name evaluation
'   Their caller forms and observed behaviour are #29 probe targets, as is the
'   case where Application.Caller raises rather than returning a value.
'
' ERROR POLICY
'   - Does not raise. Every failure while reading the caller is absorbed into
'     the no-host path; only a date-system read failure on an identified Range
'     reports a condition.
'
' DEPENDENCIES
'   - Excel.Application.Caller, Excel.Workbook.Date1904
'
' NOTES
'   - The two reads are deliberately not cached across calls. A cache that
'     outlived a date-system toggle would reintroduce the 1,462-day shift this
'     routine exists to prevent. The cost is one Caller read and one property
'     read per public call, regardless of the number of elements in that call.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CallerObject    As Object       'Application.Caller when it is an object
    Dim CallerRange     As Range        'Strongly typed caller when it is a Range
    Dim HostWorkbook    As Workbook     'The caller's own workbook
    Dim Is1904          As Boolean      'Workbook.Date1904 for the caller

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default outcome: no worksheet host identified => documented 1900 contract
        TryResolveHostDateSystem = True
        Condition = KPR_COND_NONE
        DateSystem = 1900

'------------------------------------------------------------------------------
' IDENTIFY THE CALLER
'------------------------------------------------------------------------------
    'One guarded object assignment. A non-object caller (Error from the
    'Immediate window, String from a shape) fails the Set and stays Nothing; a
    'raise during the read is absorbed into the same no-host outcome.
        On Error Resume Next
        Set CallerObject = Application.Caller
        On Error GoTo 0

    'No object => no worksheet host could be identified
        If CallerObject Is Nothing Then Exit Function

    'An object that is not a Range is also not a worksheet host
        If Not TypeOf CallerObject Is Range Then Exit Function

    'Bind strongly typed; the caller is a worksheet Range from here on
        Set CallerRange = CallerObject

'------------------------------------------------------------------------------
' READ THE CALLER'S DATE SYSTEM
'------------------------------------------------------------------------------
    'Reach the workbook through the caller itself, never through an active
    'object. Any failure from here on is HOST_UNRESOLVED: the host was
    'identified, so answering under 1900 would be a guess.
        On Error GoTo Unresolved
        Set HostWorkbook = CallerRange.Worksheet.Parent
        Is1904 = HostWorkbook.Date1904
        On Error GoTo 0

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Report the identified system
        If Is1904 Then DateSystem = 1904
        Exit Function

'------------------------------------------------------------------------------
' UNRESOLVED
'------------------------------------------------------------------------------
Unresolved:
    'Identified worksheet host, unreadable date system
        On Error GoTo 0
        Condition = KPR_COND_HOST_UNRESOLVED
        TryResolveHostDateSystem = False

End Function

Private Function PassHostGuard( _
    ByRef ErrOut As Variant) _
    As Boolean
'
'==============================================================================
'                                PassHostGuard
'------------------------------------------------------------------------------
' PURPOSE
'   The once-per-call date-system guard for every value-taking public
'   function.
'
' SIGNATURE
'   PassHostGuard(ErrOut) -> Boolean
'
' OUTPUTS
'   ErrOut (ByRef)   assigned ONLY on refusal, to the call-level #N/A
'
' RETURNS
'   Boolean
'     TRUE  => proceed under the 1900 serial contract
'     FALSE => refuse: 1904 worksheet host (HOST_DATE1904) or an identified
'              host whose date system cannot be read (HOST_UNRESOLVED)
'
' BEHAVIOR
'   - Runs before any argument resolver or calculation, so it is call-level
'     and precedes every element-level outcome.
'   - A 1904 host is refused rather than compensated. Compensation would mean
'     silently choosing between two interpretations of every serial the caller
'     supplied; refusal makes the mismatch visible.
'
' DEPENDENCIES
'   - TryResolveHostDateSystem
'   - KPR_Core_Err.ErrForCondition
'
' UPDATED
'   2026-09-01
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim DateSystem      As Long             'Classified host date system
    Dim Condition       As KPR_Condition    'Failure condition, if any

'------------------------------------------------------------------------------
' CLASSIFY
'------------------------------------------------------------------------------
    'Default is refusal until the classifier says otherwise
        PassHostGuard = False

    'An identified host with an unreadable date system is refused
        If Not TryResolveHostDateSystem(DateSystem, Condition) Then
            ErrOut = ErrForCondition(Condition)
            Exit Function
        End If

    'An identified 1904 host is refused
        If DateSystem = 1904 Then
            ErrOut = ErrForCondition(KPR_COND_HOST_DATE1904)
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Proceed under the documented 1900 contract
        PassHostGuard = True

End Function
