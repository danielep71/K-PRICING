Attribute VB_Name = "KPR_Dates"

Option Explicit

'
'------------------------------------------------------------------------------
' MODULE: KPR_Dates
'------------------------------------------------------------------------------
' PURPOSE
'   The worksheet functions of the date layer: the 22 KPR_Dates_* functions, the supported API.
'
' WHY THIS EXISTS
'   A worksheet calls these functions by name. Each one checks the calling workbook, reads its
'   arguments in the order and with the rules of the calculation contracts, and returns its
'   result or an Excel error, cell for cell. No date algorithm lives here: CORE_Dates computes,
'   CORE_Parse reads values, CORE_Array and M_Ranges handle shapes, and M_Host reads the host.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Each value-taking function stores its value arguments in signature order and calls
'   EvaluateCall, which runs the stages of the calculation contracts, Arrays and evaluation:
'     1. the host rule (M_Host.TryCheckHost);
'     2. the shape of every value argument, from type and dimensions alone;
'     3. the optional control;
'     4. the resolved shape (SHAPE_MISMATCH);
'     5. the capacity, from the dimensions alone;
'     6. the read of every value argument;
'     7. the elements, by rows, through EvaluateElement.
'   KPR_Dates_HostDateSystem reads the host only.
'
' ERROR POLICY
'   No function raises to the worksheet. Every rejection is an Excel error built by
'   CORE_Err.ErrForCondition, or an error passed in, returned unchanged. A runtime error that a
'   contract-conforming call cannot cause is contained as #VALUE!: reaching it is a defect.
'
' DEPENDENCIES
'   CORE_Err, CORE_Parse, CORE_Dates, CORE_Array, M_Ranges, M_Host. Calculation contracts:
'   docs/methodology/CALCULATION_CONTRACTS.md, CALC-02 to CALC-10 and the common rules; data
'   contract: docs/methodology/DATA_CONTRACT.md, Function arguments
'
' NOTES
'   - The public functions are exactly the 22 names of docs/PUBLIC_API.txt, [supported]. Every
'     other member is Private
'   - Only KPR_Dates_HostDateSystem calls Application.Volatile
'   - The facade never reads a Range: it passes every argument to M_Ranges unread
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE ENUMS
'------------------------------------------------------------------------------
    'The function an evaluation serves
    Private Enum KPRDateFunction
        dfDayOfWeek = 1                 'KPR_Dates_DayOfWeek
        dfDaysInMonth = 2               'KPR_Dates_DaysInMonth
        dfDaysInYear = 3                'KPR_Dates_DaysInYear
        dfBeginOfMonth = 4              'KPR_Dates_BeginOfMonth
        dfEndOfMonth = 5                'KPR_Dates_EndOfMonth
        dfBeginOfQuarter = 6            'KPR_Dates_BeginOfQuarter
        dfEndOfQuarter = 7              'KPR_Dates_EndOfQuarter
        dfBeginOfYear = 8               'KPR_Dates_BeginOfYear
        dfEndOfYear = 9                 'KPR_Dates_EndOfYear
        dfIsMonthEnd = 10               'KPR_Dates_IsMonthEnd
        dfIsQuarterEnd = 11             'KPR_Dates_IsQuarterEnd
        dfIsYearEnd = 12                'KPR_Dates_IsYearEnd
        dfIsLeapYear = 13               'KPR_Dates_IsLeapYear
        dfAddDays = 14                  'KPR_Dates_AddDays
        dfAddWeeks = 15                 'KPR_Dates_AddWeeks
        dfAddMonths = 16                'KPR_Dates_AddMonths
        dfAddYears = 17                 'KPR_Dates_AddYears
        dfNthWeekdayOfMonth = 18        'KPR_Dates_NthWeekdayOfMonth
        dfLastWeekdayOfMonth = 19       'KPR_Dates_LastWeekdayOfMonth
        dfPillarFromDates = 20          'KPR_Dates_PillarFromDates
        dfDateFromPillar = 21           'KPR_Dates_DateFromPillar
    End Enum

    'The domain of a whole-number argument
    Private Enum KPRDomain
        dmYear = 1                      'YearIn, 1900 to 9999
        dmMonth = 2                     'MonthIn, 1 to 12
        dmWeekday = 3                   'WdIndex, 1 to 7
        dmOccurrence = 4                'n, 1 to 5
    End Enum

    'The kind of optional control a function takes
    Private Enum KPRControlKind
        ckNone = 0                      'No control
        ckBoolean = 1                   'Opt_WeekBaseMonday or Opt_KeepEOM
        ckRounding = 2                  'Opt_Rounding
    End Enum


'
'------------------------------------------------------------------------------
'
'                             WORKSHEET FUNCTIONS
'
'------------------------------------------------------------------------------
'

Public Function KPR_Dates_DayOfWeek( _
    ByVal DateIn As Variant, _
    Optional ByVal Opt_WeekBaseMonday As Variant = True) As Variant
'
'------------------------------------------------------------------------------
'                         WORKSHEET FUNCTION DAYOFWEEK
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of the weekday of a date, 1 to 7, Monday-based by default.
'
' WHY THIS EXISTS
'   The supported worksheet function of calculation contract CALC-02.
'
' INPUTS
'   DateIn
'     A date: a VBA Date, a 1900-system serial or exact YYYY-MM-DD text; a scalar, an array or a Range.
'
'   Opt_WeekBaseMonday
'     True: Monday is 1. False: Sunday is 1, as WEEKDAY. A Boolean, or Empty for the default; one
'     cell or element at most, default True.
'
' RETURNS
'   Variant
'     For scalar arguments, a scalar: Long, 1 to 7. For an array or a multi-cell Range, a
'     two-dimensional array of the arguments' shape, with an Excel error at each position
'     that fails. A call-level failure is one Excel error for the whole call.
'
' BEHAVIOR
'   Runs EvaluateCall, which applies the stages of the calculation contracts and evaluates
'   each position in signature order.
'
' ERROR POLICY
'   Does not raise. Each condition returns the error of the data contract's Conditions
'   table, and an error passed in is returned unchanged.
'
' DEPENDENCIES
'   StoreArgument, EvaluateCall. CALC-02
'
' NOTES
'   Not volatile.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_DayOfWeek"

    Dim Values(1 To 1)  As Variant  'Value arguments, in signature order

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Store each argument unread, then evaluate the call
        StoreArgument Values(1), DateIn
        KPR_Dates_DayOfWeek = EvaluateCall(dfDayOfWeek, Values, Opt_WeekBaseMonday)

End Function

Public Function KPR_Dates_DaysInMonth( _
    ByVal DateIn As Variant) As Variant
'
'------------------------------------------------------------------------------
'                        WORKSHEET FUNCTION DAYSINMONTH
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of days of the month that holds a date.
'
' WHY THIS EXISTS
'   The supported worksheet function of calculation contract CALC-03.
'
' INPUTS
'   DateIn
'     A date: a VBA Date, a 1900-system serial or exact YYYY-MM-DD text; a scalar, an array or a Range.
'
' RETURNS
'   Variant
'     For scalar arguments, a scalar: Long, 28 to 31. For an array or a multi-cell Range, a
'     two-dimensional array of the arguments' shape, with an Excel error at each position
'     that fails. A call-level failure is one Excel error for the whole call.
'
' BEHAVIOR
'   Runs EvaluateCall, which applies the stages of the calculation contracts and evaluates
'   each position in signature order.
'
' ERROR POLICY
'   Does not raise. Each condition returns the error of the data contract's Conditions
'   table, and an error passed in is returned unchanged.
'
' DEPENDENCIES
'   StoreArgument, EvaluateCall. CALC-03
'
' NOTES
'   Not volatile.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_DaysInMonth"

    Dim Values(1 To 1)  As Variant  'Value arguments, in signature order

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Store each argument unread, then evaluate the call
        StoreArgument Values(1), DateIn
        KPR_Dates_DaysInMonth = EvaluateCall(dfDaysInMonth, Values)

End Function

Public Function KPR_Dates_DaysInYear( _
    ByVal YearIn As Variant) As Variant
'
'------------------------------------------------------------------------------
'                        WORKSHEET FUNCTION DAYSINYEAR
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of days of a calendar year.
'
' WHY THIS EXISTS
'   The supported worksheet function of calculation contract CALC-03.
'
' INPUTS
'   YearIn
'     A calendar year, 1900 to 9999, never a date; a scalar, an array or a Range.
'
' RETURNS
'   Variant
'     For scalar arguments, a scalar: Long, 365 or 366. For an array or a multi-cell Range, a
'     two-dimensional array of the arguments' shape, with an Excel error at each position
'     that fails. A call-level failure is one Excel error for the whole call.
'
' BEHAVIOR
'   Runs EvaluateCall, which applies the stages of the calculation contracts and evaluates
'   each position in signature order.
'
' ERROR POLICY
'   Does not raise. Each condition returns the error of the data contract's Conditions
'   table, and an error passed in is returned unchanged.
'
' DEPENDENCIES
'   StoreArgument, EvaluateCall. CALC-03
'
' NOTES
'   Not volatile.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_DaysInYear"

    Dim Values(1 To 1)  As Variant  'Value arguments, in signature order

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Store each argument unread, then evaluate the call
        StoreArgument Values(1), YearIn
        KPR_Dates_DaysInYear = EvaluateCall(dfDaysInYear, Values)

End Function

Public Function KPR_Dates_BeginOfMonth( _
    ByVal DateIn As Variant) As Variant
'
'------------------------------------------------------------------------------
'                       WORKSHEET FUNCTION BEGINOFMONTH
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the first day of the month that holds a date.
'
' WHY THIS EXISTS
'   The supported worksheet function of calculation contract CALC-04.
'
' INPUTS
'   DateIn
'     A date: a VBA Date, a 1900-system serial or exact YYYY-MM-DD text; a scalar, an array or a Range.
'
' RETURNS
'   Variant
'     For scalar arguments, a scalar: Date. For an array or a multi-cell Range, a
'     two-dimensional array of the arguments' shape, with an Excel error at each position
'     that fails. A call-level failure is one Excel error for the whole call.
'
' BEHAVIOR
'   Runs EvaluateCall, which applies the stages of the calculation contracts and evaluates
'   each position in signature order.
'
' ERROR POLICY
'   Does not raise. Each condition returns the error of the data contract's Conditions
'   table, and an error passed in is returned unchanged.
'
' DEPENDENCIES
'   StoreArgument, EvaluateCall. CALC-04
'
' NOTES
'   Not volatile.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_BeginOfMonth"

    Dim Values(1 To 1)  As Variant  'Value arguments, in signature order

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Store each argument unread, then evaluate the call
        StoreArgument Values(1), DateIn
        KPR_Dates_BeginOfMonth = EvaluateCall(dfBeginOfMonth, Values)

End Function

Public Function KPR_Dates_EndOfMonth( _
    ByVal DateIn As Variant) As Variant
'
'------------------------------------------------------------------------------
'                        WORKSHEET FUNCTION ENDOFMONTH
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the last day of the month that holds a date.
'
' WHY THIS EXISTS
'   The supported worksheet function of calculation contract CALC-04.
'
' INPUTS
'   DateIn
'     A date: a VBA Date, a 1900-system serial or exact YYYY-MM-DD text; a scalar, an array or a Range.
'
' RETURNS
'   Variant
'     For scalar arguments, a scalar: Date. For an array or a multi-cell Range, a
'     two-dimensional array of the arguments' shape, with an Excel error at each position
'     that fails. A call-level failure is one Excel error for the whole call.
'
' BEHAVIOR
'   Runs EvaluateCall, which applies the stages of the calculation contracts and evaluates
'   each position in signature order.
'
' ERROR POLICY
'   Does not raise. Each condition returns the error of the data contract's Conditions
'   table, and an error passed in is returned unchanged.
'
' DEPENDENCIES
'   StoreArgument, EvaluateCall. CALC-04
'
' NOTES
'   Not volatile.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_EndOfMonth"

    Dim Values(1 To 1)  As Variant  'Value arguments, in signature order

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Store each argument unread, then evaluate the call
        StoreArgument Values(1), DateIn
        KPR_Dates_EndOfMonth = EvaluateCall(dfEndOfMonth, Values)

End Function

Public Function KPR_Dates_BeginOfQuarter( _
    ByVal DateIn As Variant) As Variant
'
'------------------------------------------------------------------------------
'                      WORKSHEET FUNCTION BEGINOFQUARTER
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the first day of the calendar quarter that holds a date.
'
' WHY THIS EXISTS
'   The supported worksheet function of calculation contract CALC-04.
'
' INPUTS
'   DateIn
'     A date: a VBA Date, a 1900-system serial or exact YYYY-MM-DD text; a scalar, an array or a Range.
'
' RETURNS
'   Variant
'     For scalar arguments, a scalar: Date; #NUM! for a first quarter of 1900, whose first day is
'     outside the window. For an array or a multi-cell Range, a
'     two-dimensional array of the arguments' shape, with an Excel error at each position
'     that fails. A call-level failure is one Excel error for the whole call.
'
' BEHAVIOR
'   Runs EvaluateCall, which applies the stages of the calculation contracts and evaluates
'   each position in signature order.
'
' ERROR POLICY
'   Does not raise. Each condition returns the error of the data contract's Conditions
'   table, and an error passed in is returned unchanged.
'
' DEPENDENCIES
'   StoreArgument, EvaluateCall. CALC-04
'
' NOTES
'   Not volatile.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_BeginOfQuarter"

    Dim Values(1 To 1)  As Variant  'Value arguments, in signature order

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Store each argument unread, then evaluate the call
        StoreArgument Values(1), DateIn
        KPR_Dates_BeginOfQuarter = EvaluateCall(dfBeginOfQuarter, Values)

End Function

Public Function KPR_Dates_EndOfQuarter( _
    ByVal DateIn As Variant) As Variant
'
'------------------------------------------------------------------------------
'                       WORKSHEET FUNCTION ENDOFQUARTER
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the last day of the calendar quarter that holds a date.
'
' WHY THIS EXISTS
'   The supported worksheet function of calculation contract CALC-04.
'
' INPUTS
'   DateIn
'     A date: a VBA Date, a 1900-system serial or exact YYYY-MM-DD text; a scalar, an array or a Range.
'
' RETURNS
'   Variant
'     For scalar arguments, a scalar: Date. For an array or a multi-cell Range, a
'     two-dimensional array of the arguments' shape, with an Excel error at each position
'     that fails. A call-level failure is one Excel error for the whole call.
'
' BEHAVIOR
'   Runs EvaluateCall, which applies the stages of the calculation contracts and evaluates
'   each position in signature order.
'
' ERROR POLICY
'   Does not raise. Each condition returns the error of the data contract's Conditions
'   table, and an error passed in is returned unchanged.
'
' DEPENDENCIES
'   StoreArgument, EvaluateCall. CALC-04
'
' NOTES
'   Not volatile.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_EndOfQuarter"

    Dim Values(1 To 1)  As Variant  'Value arguments, in signature order

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Store each argument unread, then evaluate the call
        StoreArgument Values(1), DateIn
        KPR_Dates_EndOfQuarter = EvaluateCall(dfEndOfQuarter, Values)

End Function

Public Function KPR_Dates_BeginOfYear( _
    ByVal DateIn As Variant) As Variant
'
'------------------------------------------------------------------------------
'                        WORKSHEET FUNCTION BEGINOFYEAR
'------------------------------------------------------------------------------
' PURPOSE
'   Returns 1 January of the year that holds a date.
'
' WHY THIS EXISTS
'   The supported worksheet function of calculation contract CALC-04.
'
' INPUTS
'   DateIn
'     A date: a VBA Date, a 1900-system serial or exact YYYY-MM-DD text; a scalar, an array or a Range.
'
' RETURNS
'   Variant
'     For scalar arguments, a scalar: Date; #NUM! for a 1900 date, whose year begins outside the
'     window. For an array or a multi-cell Range, a
'     two-dimensional array of the arguments' shape, with an Excel error at each position
'     that fails. A call-level failure is one Excel error for the whole call.
'
' BEHAVIOR
'   Runs EvaluateCall, which applies the stages of the calculation contracts and evaluates
'   each position in signature order.
'
' ERROR POLICY
'   Does not raise. Each condition returns the error of the data contract's Conditions
'   table, and an error passed in is returned unchanged.
'
' DEPENDENCIES
'   StoreArgument, EvaluateCall. CALC-04
'
' NOTES
'   Not volatile.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_BeginOfYear"

    Dim Values(1 To 1)  As Variant  'Value arguments, in signature order

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Store each argument unread, then evaluate the call
        StoreArgument Values(1), DateIn
        KPR_Dates_BeginOfYear = EvaluateCall(dfBeginOfYear, Values)

End Function

Public Function KPR_Dates_EndOfYear( _
    ByVal DateIn As Variant) As Variant
'
'------------------------------------------------------------------------------
'                         WORKSHEET FUNCTION ENDOFYEAR
'------------------------------------------------------------------------------
' PURPOSE
'   Returns 31 December of the year that holds a date.
'
' WHY THIS EXISTS
'   The supported worksheet function of calculation contract CALC-04.
'
' INPUTS
'   DateIn
'     A date: a VBA Date, a 1900-system serial or exact YYYY-MM-DD text; a scalar, an array or a Range.
'
' RETURNS
'   Variant
'     For scalar arguments, a scalar: Date. For an array or a multi-cell Range, a
'     two-dimensional array of the arguments' shape, with an Excel error at each position
'     that fails. A call-level failure is one Excel error for the whole call.
'
' BEHAVIOR
'   Runs EvaluateCall, which applies the stages of the calculation contracts and evaluates
'   each position in signature order.
'
' ERROR POLICY
'   Does not raise. Each condition returns the error of the data contract's Conditions
'   table, and an error passed in is returned unchanged.
'
' DEPENDENCIES
'   StoreArgument, EvaluateCall. CALC-04
'
' NOTES
'   Not volatile.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_EndOfYear"

    Dim Values(1 To 1)  As Variant  'Value arguments, in signature order

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Store each argument unread, then evaluate the call
        StoreArgument Values(1), DateIn
        KPR_Dates_EndOfYear = EvaluateCall(dfEndOfYear, Values)

End Function

Public Function KPR_Dates_IsMonthEnd( _
    ByVal DateIn As Variant) As Variant
'
'------------------------------------------------------------------------------
'                        WORKSHEET FUNCTION ISMONTHEND
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a date is the last day of its month.
'
' WHY THIS EXISTS
'   The supported worksheet function of calculation contract CALC-05.
'
' INPUTS
'   DateIn
'     A date: a VBA Date, a 1900-system serial or exact YYYY-MM-DD text; a scalar, an array or a Range.
'
' RETURNS
'   Variant
'     For scalar arguments, a scalar: Boolean. For an array or a multi-cell Range, a
'     two-dimensional array of the arguments' shape, with an Excel error at each position
'     that fails. A call-level failure is one Excel error for the whole call.
'
' BEHAVIOR
'   Runs EvaluateCall, which applies the stages of the calculation contracts and evaluates
'   each position in signature order.
'
' ERROR POLICY
'   Does not raise. Each condition returns the error of the data contract's Conditions
'   table, and an error passed in is returned unchanged.
'
' DEPENDENCIES
'   StoreArgument, EvaluateCall. CALC-05
'
' NOTES
'   Not volatile.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_IsMonthEnd"

    Dim Values(1 To 1)  As Variant  'Value arguments, in signature order

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Store each argument unread, then evaluate the call
        StoreArgument Values(1), DateIn
        KPR_Dates_IsMonthEnd = EvaluateCall(dfIsMonthEnd, Values)

End Function

Public Function KPR_Dates_IsQuarterEnd( _
    ByVal DateIn As Variant) As Variant
'
'------------------------------------------------------------------------------
'                       WORKSHEET FUNCTION ISQUARTEREND
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a date is the last day of its calendar quarter.
'
' WHY THIS EXISTS
'   The supported worksheet function of calculation contract CALC-05.
'
' INPUTS
'   DateIn
'     A date: a VBA Date, a 1900-system serial or exact YYYY-MM-DD text; a scalar, an array or a Range.
'
' RETURNS
'   Variant
'     For scalar arguments, a scalar: Boolean. For an array or a multi-cell Range, a
'     two-dimensional array of the arguments' shape, with an Excel error at each position
'     that fails. A call-level failure is one Excel error for the whole call.
'
' BEHAVIOR
'   Runs EvaluateCall, which applies the stages of the calculation contracts and evaluates
'   each position in signature order.
'
' ERROR POLICY
'   Does not raise. Each condition returns the error of the data contract's Conditions
'   table, and an error passed in is returned unchanged.
'
' DEPENDENCIES
'   StoreArgument, EvaluateCall. CALC-05
'
' NOTES
'   Not volatile.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_IsQuarterEnd"

    Dim Values(1 To 1)  As Variant  'Value arguments, in signature order

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Store each argument unread, then evaluate the call
        StoreArgument Values(1), DateIn
        KPR_Dates_IsQuarterEnd = EvaluateCall(dfIsQuarterEnd, Values)

End Function

Public Function KPR_Dates_IsYearEnd( _
    ByVal DateIn As Variant) As Variant
'
'------------------------------------------------------------------------------
'                         WORKSHEET FUNCTION ISYEAREND
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a date is 31 December.
'
' WHY THIS EXISTS
'   The supported worksheet function of calculation contract CALC-05.
'
' INPUTS
'   DateIn
'     A date: a VBA Date, a 1900-system serial or exact YYYY-MM-DD text; a scalar, an array or a Range.
'
' RETURNS
'   Variant
'     For scalar arguments, a scalar: Boolean. For an array or a multi-cell Range, a
'     two-dimensional array of the arguments' shape, with an Excel error at each position
'     that fails. A call-level failure is one Excel error for the whole call.
'
' BEHAVIOR
'   Runs EvaluateCall, which applies the stages of the calculation contracts and evaluates
'   each position in signature order.
'
' ERROR POLICY
'   Does not raise. Each condition returns the error of the data contract's Conditions
'   table, and an error passed in is returned unchanged.
'
' DEPENDENCIES
'   StoreArgument, EvaluateCall. CALC-05
'
' NOTES
'   Not volatile.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_IsYearEnd"

    Dim Values(1 To 1)  As Variant  'Value arguments, in signature order

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Store each argument unread, then evaluate the call
        StoreArgument Values(1), DateIn
        KPR_Dates_IsYearEnd = EvaluateCall(dfIsYearEnd, Values)

End Function

Public Function KPR_Dates_IsLeapYear( _
    ByVal YearIn As Variant) As Variant
'
'------------------------------------------------------------------------------
'                        WORKSHEET FUNCTION ISLEAPYEAR
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a calendar year is a Gregorian leap year.
'
' WHY THIS EXISTS
'   The supported worksheet function of calculation contract CALC-03.
'
' INPUTS
'   YearIn
'     A calendar year, 1900 to 9999, never a date; a scalar, an array or a Range.
'
' RETURNS
'   Variant
'     For scalar arguments, a scalar: Boolean. For an array or a multi-cell Range, a
'     two-dimensional array of the arguments' shape, with an Excel error at each position
'     that fails. A call-level failure is one Excel error for the whole call.
'
' BEHAVIOR
'   Runs EvaluateCall, which applies the stages of the calculation contracts and evaluates
'   each position in signature order.
'
' ERROR POLICY
'   Does not raise. Each condition returns the error of the data contract's Conditions
'   table, and an error passed in is returned unchanged.
'
' DEPENDENCIES
'   StoreArgument, EvaluateCall. CALC-03
'
' NOTES
'   Not volatile.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_IsLeapYear"

    Dim Values(1 To 1)  As Variant  'Value arguments, in signature order

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Store each argument unread, then evaluate the call
        StoreArgument Values(1), YearIn
        KPR_Dates_IsLeapYear = EvaluateCall(dfIsLeapYear, Values)

End Function

Public Function KPR_Dates_AddDays( _
    ByVal DateIn As Variant, _
    ByVal nDays As Variant) As Variant
'
'------------------------------------------------------------------------------
'                          WORKSHEET FUNCTION ADDDAYS
'------------------------------------------------------------------------------
' PURPOSE
'   Moves a date by a signed number of calendar days.
'
' WHY THIS EXISTS
'   The supported worksheet function of calculation contract CALC-06.
'
' INPUTS
'   DateIn
'     A date: a VBA Date, a 1900-system serial or exact YYYY-MM-DD text; a scalar, an array or a Range.
'
'   nDays
'     A signed whole number of days; a scalar, an array or a Range.
'
' RETURNS
'   Variant
'     For scalar arguments, a scalar: Date; #NUM! when the result leaves the window. For an array or
'     a multi-cell Range, a
'     two-dimensional array of the arguments' shape, with an Excel error at each position
'     that fails. A call-level failure is one Excel error for the whole call.
'
' BEHAVIOR
'   Runs EvaluateCall, which applies the stages of the calculation contracts and evaluates
'   each position in signature order.
'
' ERROR POLICY
'   Does not raise. Each condition returns the error of the data contract's Conditions
'   table, and an error passed in is returned unchanged.
'
' DEPENDENCIES
'   StoreArgument, EvaluateCall. CALC-06
'
' NOTES
'   Not volatile.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_AddDays"

    Dim Values(1 To 2)  As Variant  'Value arguments, in signature order

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Store each argument unread, then evaluate the call
        StoreArgument Values(1), DateIn
        StoreArgument Values(2), nDays
        KPR_Dates_AddDays = EvaluateCall(dfAddDays, Values)

End Function

Public Function KPR_Dates_AddWeeks( _
    ByVal DateIn As Variant, _
    ByVal nWeeks As Variant) As Variant
'
'------------------------------------------------------------------------------
'                         WORKSHEET FUNCTION ADDWEEKS
'------------------------------------------------------------------------------
' PURPOSE
'   Moves a date by a signed number of weeks of seven calendar days.
'
' WHY THIS EXISTS
'   The supported worksheet function of calculation contract CALC-06.
'
' INPUTS
'   DateIn
'     A date: a VBA Date, a 1900-system serial or exact YYYY-MM-DD text; a scalar, an array or a Range.
'
'   nWeeks
'     A signed whole number of weeks; a scalar, an array or a Range.
'
' RETURNS
'   Variant
'     For scalar arguments, a scalar: Date; #NUM! when the result leaves the window. For an array or
'     a multi-cell Range, a
'     two-dimensional array of the arguments' shape, with an Excel error at each position
'     that fails. A call-level failure is one Excel error for the whole call.
'
' BEHAVIOR
'   Runs EvaluateCall, which applies the stages of the calculation contracts and evaluates
'   each position in signature order.
'
' ERROR POLICY
'   Does not raise. Each condition returns the error of the data contract's Conditions
'   table, and an error passed in is returned unchanged.
'
' DEPENDENCIES
'   StoreArgument, EvaluateCall. CALC-06
'
' NOTES
'   Not volatile.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_AddWeeks"

    Dim Values(1 To 2)  As Variant  'Value arguments, in signature order

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Store each argument unread, then evaluate the call
        StoreArgument Values(1), DateIn
        StoreArgument Values(2), nWeeks
        KPR_Dates_AddWeeks = EvaluateCall(dfAddWeeks, Values)

End Function

Public Function KPR_Dates_AddMonths( _
    ByVal DateIn As Variant, _
    ByVal nMonths As Variant, _
    Optional ByVal Opt_KeepEOM As Variant = False) As Variant
'
'------------------------------------------------------------------------------
'                         WORKSHEET FUNCTION ADDMONTHS
'------------------------------------------------------------------------------
' PURPOSE
'   Moves a date by a signed number of months, clipping a day that the target month lacks.
'
' WHY THIS EXISTS
'   The supported worksheet function of calculation contract CALC-06.
'
' INPUTS
'   DateIn
'     A date: a VBA Date, a 1900-system serial or exact YYYY-MM-DD text; a scalar, an array or a Range.
'
'   nMonths
'     A signed whole number of months; a scalar, an array or a Range.
'
'   Opt_KeepEOM
'     True: a month end moves to the target month's end. A Boolean, or Empty for the default; one
'     cell or element at most, default False.
'
' RETURNS
'   Variant
'     For scalar arguments, a scalar: Date; #NUM! when the result leaves the window. For an array or
'     a multi-cell Range, a
'     two-dimensional array of the arguments' shape, with an Excel error at each position
'     that fails. A call-level failure is one Excel error for the whole call.
'
' BEHAVIOR
'   Runs EvaluateCall, which applies the stages of the calculation contracts and evaluates
'   each position in signature order.
'
' ERROR POLICY
'   Does not raise. Each condition returns the error of the data contract's Conditions
'   table, and an error passed in is returned unchanged.
'
' DEPENDENCIES
'   StoreArgument, EvaluateCall. CALC-06
'
' NOTES
'   Not volatile.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_AddMonths"

    Dim Values(1 To 2)  As Variant  'Value arguments, in signature order

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Store each argument unread, then evaluate the call
        StoreArgument Values(1), DateIn
        StoreArgument Values(2), nMonths
        KPR_Dates_AddMonths = EvaluateCall(dfAddMonths, Values, Opt_KeepEOM)

End Function

Public Function KPR_Dates_AddYears( _
    ByVal DateIn As Variant, _
    ByVal nYears As Variant, _
    Optional ByVal Opt_KeepEOM As Variant = False) As Variant
'
'------------------------------------------------------------------------------
'                         WORKSHEET FUNCTION ADDYEARS
'------------------------------------------------------------------------------
' PURPOSE
'   Moves a date by a signed number of years, as twelve months each.
'
' WHY THIS EXISTS
'   The supported worksheet function of calculation contract CALC-06.
'
' INPUTS
'   DateIn
'     A date: a VBA Date, a 1900-system serial or exact YYYY-MM-DD text; a scalar, an array or a Range.
'
'   nYears
'     A signed whole number of years; a scalar, an array or a Range.
'
'   Opt_KeepEOM
'     True: a month end moves to the target month's end. A Boolean, or Empty for the default; one
'     cell or element at most, default False.
'
' RETURNS
'   Variant
'     For scalar arguments, a scalar: Date; #NUM! when the result leaves the window. For an array or
'     a multi-cell Range, a
'     two-dimensional array of the arguments' shape, with an Excel error at each position
'     that fails. A call-level failure is one Excel error for the whole call.
'
' BEHAVIOR
'   Runs EvaluateCall, which applies the stages of the calculation contracts and evaluates
'   each position in signature order.
'
' ERROR POLICY
'   Does not raise. Each condition returns the error of the data contract's Conditions
'   table, and an error passed in is returned unchanged.
'
' DEPENDENCIES
'   StoreArgument, EvaluateCall. CALC-06
'
' NOTES
'   Not volatile.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_AddYears"

    Dim Values(1 To 2)  As Variant  'Value arguments, in signature order

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Store each argument unread, then evaluate the call
        StoreArgument Values(1), DateIn
        StoreArgument Values(2), nYears
        KPR_Dates_AddYears = EvaluateCall(dfAddYears, Values, Opt_KeepEOM)

End Function

Public Function KPR_Dates_NthWeekdayOfMonth( _
    ByVal YearIn As Variant, _
    ByVal MonthIn As Variant, _
    ByVal WdIndex As Variant, _
    ByVal n As Variant, _
    Optional ByVal Opt_WeekBaseMonday As Variant = True) As Variant
'
'------------------------------------------------------------------------------
'                     WORKSHEET FUNCTION NTHWEEKDAYOFMONTH
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the date of the nth given weekday of a month.
'
' WHY THIS EXISTS
'   The supported worksheet function of calculation contract CALC-07.
'
' INPUTS
'   YearIn
'     A calendar year, 1900 to 9999; a scalar, an array or a Range.
'
'   MonthIn
'     A month, 1 to 12; a scalar, an array or a Range.
'
'   WdIndex
'     A weekday number, 1 to 7, in the weekday base; a scalar, an array or a Range.
'
'   n
'     The occurrence, 1 to 5; a scalar, an array or a Range.
'
'   Opt_WeekBaseMonday
'     True: Monday is 1. False: Sunday is 1, as WEEKDAY. A Boolean, or Empty for the default; one
'     cell or element at most, default True.
'
' RETURNS
'   Variant
'     For scalar arguments, a scalar: Date; #NUM! when the month has no such occurrence or the date
'     is before the window. For an array or a multi-cell Range, a
'     two-dimensional array of the arguments' shape, with an Excel error at each position
'     that fails. A call-level failure is one Excel error for the whole call.
'
' BEHAVIOR
'   Runs EvaluateCall, which applies the stages of the calculation contracts and evaluates
'   each position in signature order.
'
' ERROR POLICY
'   Does not raise. Each condition returns the error of the data contract's Conditions
'   table, and an error passed in is returned unchanged.
'
' DEPENDENCIES
'   StoreArgument, EvaluateCall. CALC-07
'
' NOTES
'   Not volatile.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_NthWeekdayOfMonth"

    Dim Values(1 To 4)  As Variant  'Value arguments, in signature order

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Store each argument unread, then evaluate the call
        StoreArgument Values(1), YearIn
        StoreArgument Values(2), MonthIn
        StoreArgument Values(3), WdIndex
        StoreArgument Values(4), n
        KPR_Dates_NthWeekdayOfMonth = EvaluateCall(dfNthWeekdayOfMonth, Values, Opt_WeekBaseMonday)

End Function

Public Function KPR_Dates_LastWeekdayOfMonth( _
    ByVal YearIn As Variant, _
    ByVal MonthIn As Variant, _
    ByVal WdIndex As Variant, _
    Optional ByVal Opt_WeekBaseMonday As Variant = True) As Variant
'
'------------------------------------------------------------------------------
'                    WORKSHEET FUNCTION LASTWEEKDAYOFMONTH
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the date of the last given weekday of a month.
'
' WHY THIS EXISTS
'   The supported worksheet function of calculation contract CALC-07.
'
' INPUTS
'   YearIn
'     A calendar year, 1900 to 9999; a scalar, an array or a Range.
'
'   MonthIn
'     A month, 1 to 12; a scalar, an array or a Range.
'
'   WdIndex
'     A weekday number, 1 to 7, in the weekday base; a scalar, an array or a Range.
'
'   Opt_WeekBaseMonday
'     True: Monday is 1. False: Sunday is 1, as WEEKDAY. A Boolean, or Empty for the default; one
'     cell or element at most, default True.
'
' RETURNS
'   Variant
'     For scalar arguments, a scalar: Date; #NUM! when the date is before the window. For an array
'     or a multi-cell Range, a
'     two-dimensional array of the arguments' shape, with an Excel error at each position
'     that fails. A call-level failure is one Excel error for the whole call.
'
' BEHAVIOR
'   Runs EvaluateCall, which applies the stages of the calculation contracts and evaluates
'   each position in signature order.
'
' ERROR POLICY
'   Does not raise. Each condition returns the error of the data contract's Conditions
'   table, and an error passed in is returned unchanged.
'
' DEPENDENCIES
'   StoreArgument, EvaluateCall. CALC-07
'
' NOTES
'   Not volatile.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_LastWeekdayOfMonth"

    Dim Values(1 To 3)  As Variant  'Value arguments, in signature order

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Store each argument unread, then evaluate the call
        StoreArgument Values(1), YearIn
        StoreArgument Values(2), MonthIn
        StoreArgument Values(3), WdIndex
        KPR_Dates_LastWeekdayOfMonth = EvaluateCall(dfLastWeekdayOfMonth, Values, Opt_WeekBaseMonday)

End Function

Public Function KPR_Dates_PillarFromDates( _
    ByVal StartDate As Variant, _
    ByVal EndDate As Variant, _
    Optional ByVal Opt_Rounding As Variant = "NEAREST") As Variant
'
'------------------------------------------------------------------------------
'                      WORKSHEET FUNCTION PILLARFROMDATES
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the pillar token that labels the interval between two dates, rounded to a curve bucket.
'
' WHY THIS EXISTS
'   The supported worksheet function of calculation contract CALC-09.
'
' INPUTS
'   StartDate
'     A date: a VBA Date, a 1900-system serial or exact YYYY-MM-DD text; a scalar, an array or a Range.
'
'   EndDate
'     A date: a VBA Date, a 1900-system serial or exact YYYY-MM-DD text; a scalar, an array or a Range.
'
'   Opt_Rounding
'     NEAREST, FLOOR or CEILING, in any case. Text, or Empty for the default; one cell or element at
'     most, default NEAREST.
'
' RETURNS
'   Variant
'     For scalar arguments, a scalar: String, such as 3M, 1Y6M or -2W; #NUM! when CEILING has no
'     candidate in the window. For an array or a multi-cell Range, a
'     two-dimensional array of the arguments' shape, with an Excel error at each position
'     that fails. A call-level failure is one Excel error for the whole call.
'
' BEHAVIOR
'   Runs EvaluateCall, which applies the stages of the calculation contracts and evaluates
'   each position in signature order.
'
' ERROR POLICY
'   Does not raise. Each condition returns the error of the data contract's Conditions
'   table, and an error passed in is returned unchanged.
'
' DEPENDENCIES
'   StoreArgument, EvaluateCall. CALC-09
'
' NOTES
'   Not volatile.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_PillarFromDates"

    Dim Values(1 To 2)  As Variant  'Value arguments, in signature order

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Store each argument unread, then evaluate the call
        StoreArgument Values(1), StartDate
        StoreArgument Values(2), EndDate
        KPR_Dates_PillarFromDates = EvaluateCall(dfPillarFromDates, Values, Opt_Rounding)

End Function

Public Function KPR_Dates_DateFromPillar( _
    ByVal StartDate As Variant, _
    ByVal Pillar As Variant) As Variant
'
'------------------------------------------------------------------------------
'                      WORKSHEET FUNCTION DATEFROMPILLAR
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the date that a pillar token names from a start date.
'
' WHY THIS EXISTS
'   The supported worksheet function of calculation contract CALC-08.
'
' INPUTS
'   StartDate
'     A date: a VBA Date, a 1900-system serial or exact YYYY-MM-DD text; a scalar, an array or a Range.
'
'   Pillar
'     A pillar token, such as 3M, 1Y6M, 2W3D, ON or TN; a scalar, an array or a Range.
'
' RETURNS
'   Variant
'     For scalar arguments, a scalar: Date; #NUM! when the pillar or its result leaves the domain.
'     For an array or a multi-cell Range, a
'     two-dimensional array of the arguments' shape, with an Excel error at each position
'     that fails. A call-level failure is one Excel error for the whole call.
'
' BEHAVIOR
'   Runs EvaluateCall, which applies the stages of the calculation contracts and evaluates
'   each position in signature order.
'
' ERROR POLICY
'   Does not raise. Each condition returns the error of the data contract's Conditions
'   table, and an error passed in is returned unchanged.
'
' DEPENDENCIES
'   StoreArgument, EvaluateCall. CALC-08
'
' NOTES
'   Not volatile.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_DateFromPillar"

    Dim Values(1 To 2)  As Variant  'Value arguments, in signature order

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Store each argument unread, then evaluate the call
        StoreArgument Values(1), StartDate
        StoreArgument Values(2), Pillar
        KPR_Dates_DateFromPillar = EvaluateCall(dfDateFromPillar, Values)

End Function

Public Function KPR_Dates_HostDateSystem() As Variant

'
'------------------------------------------------------------------------------
'                      WORKSHEET FUNCTION HOSTDATESYSTEM
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the date system of the calling workbook: 1900 or 1904.
'
' WHY THIS EXISTS
'   A library #N/A and an #N/A passed in are the same Excel value; this function tells a
'   user which one a 1904 workbook produced. Calculation contract CALC-10.
'
' INPUTS
'   None
'
' RETURNS
'   Variant
'     1900 or 1904, a Long; #N/A when a worksheet caller's date system cannot be read.
'
' BEHAVIOR
'   Marks itself volatile, so that an ordinary recalculation follows a change of the
'   workbook's date system, then reads the host. With no worksheet caller it returns 1900.
'
' ERROR POLICY
'   Does not raise. An unexpected runtime error is contained as #VALUE!.
'
' DEPENDENCIES
'   Application.Volatile, M_Host.TryReadHostDateSystem, CORE_Err. CALC-10
'
' NOTES
'   The only volatile function of the date layer.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.KPR_Dates_HostDateSystem"

    Dim DateSystem      As Long         '1900 or 1904
    Dim Condition       As KPRCondition 'Host condition, when it cannot be read

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Re-evaluate on every ordinary recalculation (CALC-10 step 1)
        Application.Volatile True
    'Contain an unexpected error
        On Error GoTo Contained

'------------------------------------------------------------------------------
' READ HOST
'------------------------------------------------------------------------------
    'Return the date system, or the condition that hides it
        If Not M_Host.TryReadHostDateSystem(DateSystem, Condition) Then
            KPR_Dates_HostDateSystem = CORE_Err.ErrForCondition(Condition)
            Exit Function
        End If
        KPR_Dates_HostDateSystem = DateSystem

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the containment handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Contained:
    'A defect, contained as #VALUE!
        KPR_Dates_HostDateSystem = CORE_Err.ErrValue()

End Function

'
'------------------------------------------------------------------------------
'
'                                 CALL ENGINE
'
'------------------------------------------------------------------------------
'

Private Sub StoreArgument( _
    ByRef Slot As Variant, _
    ByVal ValueIn As Variant)
'
'------------------------------------------------------------------------------
'                           STORE AN ARGUMENT UNREAD
'------------------------------------------------------------------------------
' PURPOSE
'   Stores an argument in a slot without reading it.
'
' WHY THIS EXISTS
'   A Let copy of a Variant that holds a Range reads the Range's Value. An object is stored
'   with Set, so M_Ranges receives the Range itself.
'
' INPUTS
'   Slot (ByRef, output)
'     The slot of the value arguments.
'
'   ValueIn
'     The argument as the worksheet passed it.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Sets an object; assigns anything else.
'
' ERROR POLICY
'   None.
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' STORE
'------------------------------------------------------------------------------
    'An object is stored as a reference, never through its default member
        If IsObject(ValueIn) Then
            Set Slot = ValueIn
        Else
            Slot = ValueIn
        End If

End Sub

Private Function EvaluateCall( _
    ByVal Fn As KPRDateFunction, _
    ByRef Values() As Variant, _
    Optional ByVal Control As Variant) As Variant
'
'------------------------------------------------------------------------------
'                               EVALUATE A CALL
'------------------------------------------------------------------------------
' PURPOSE
'   Evaluates one call of a value-taking function: its call-level stages, then each
'   position of the result.
'
' WHY THIS EXISTS
'   Every function applies the same stages in the same order, so they are written once.
'
' INPUTS
'   Fn
'     The function.
'
'   Values()
'     The value arguments, in signature order, stored unread.
'
'   Control
'     The optional control as the worksheet passed it; omitted for a function that has none.
'
' RETURNS
'   Variant
'     A scalar result, a two-dimensional array of results, or one call-level Excel error.
'
' BEHAVIOR
'   Stage 1, the host; stage 2, the shape of every value argument; stage 3, the control,
'   then the resolved shape; stage 4, the capacity; stage 5, the read; then each position by
'   rows, every column of the first row first.
'
' ERROR POLICY
'   Does not raise. Each failing stage returns the Excel error of its condition. An
'   unexpected runtime error is contained as #VALUE!.
'
' DEPENDENCIES
'   M_Host, M_Ranges, CORE_Array, CORE_Err, TryResolveControl, EvaluateElement. Calculation
'   contracts: Arrays and evaluation, Host and date system
'
' NOTES
'   Stage 3 resolves the control before the shapes are folded together, so a bad control
'   is reported before SHAPE_MISMATCH.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.EvaluateCall"

    Dim Count           As Long              'Number of value arguments
    Dim Kinds()         As KPRArgShape       'Shape kind of each value argument
    Dim RowCounts()     As Long              'Rows of each value argument
    Dim ColCounts()     As Long              'Columns of each value argument
    Dim Payloads()      As Variant           'Read payload of each value argument
    Dim Elements()      As Variant           'The value of each argument at one position
    Dim OutKind         As KPRArgShape       'Resolved shape kind
    Dim OutRows         As Long              'Resolved rows
    Dim OutCols         As Long              'Resolved columns
    Dim Result          As Variant           'Two-dimensional result of a non-scalar call
    Dim Flag            As Boolean           'Resolved Boolean control
    Dim Mode            As KPRPillarRounding 'Resolved rounding control
    Dim Failure         As Variant           'The error a failing stage returns
    Dim Condition       As KPRCondition      'Condition of a failing stage
    Dim I               As Long              'Argument cursor
    Dim R               As Long              'Row cursor
    Dim C               As Long              'Column cursor

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Contain an unexpected error as #VALUE!
        On Error GoTo Contained
    'Size the per-argument arrays
        Count = UBound(Values) - LBound(Values) + 1
        ReDim Kinds(1 To Count)
        ReDim RowCounts(1 To Count)
        ReDim ColCounts(1 To Count)
        ReDim Payloads(1 To Count)
        ReDim Elements(1 To Count)

'------------------------------------------------------------------------------
' CHECK HOST
'------------------------------------------------------------------------------
    'Stage 1: a 1904 or unreadable worksheet host refuses the call before any argument is read
        If Not M_Host.TryCheckHost(Condition) Then
            EvaluateCall = CORE_Err.ErrForCondition(Condition)
            Exit Function
        End If

'------------------------------------------------------------------------------
' CLASSIFY SHAPES
'------------------------------------------------------------------------------
    'Stage 2: every value argument, from its type and dimensions alone
        For I = 1 To Count
            If Not M_Ranges.TryDescribeArgument(Values(LBound(Values) + I - 1), Kinds(I), RowCounts(I), _
                ColCounts(I), Condition) Then
                EvaluateCall = CORE_Err.ErrForCondition(Condition)
                Exit Function
            End If
        Next I

'------------------------------------------------------------------------------
' RESOLVE CONTROL
'------------------------------------------------------------------------------
    'Stage 3: the optional control, one element at most
        If Not TryResolveControl(Fn, Control, Flag, Mode, Failure) Then
            EvaluateCall = Failure
            Exit Function
        End If

'------------------------------------------------------------------------------
' RESOLVE SHAPE
'------------------------------------------------------------------------------
    'Stage 3: a scalar repeats over the arrays, which must all have one shape
        OutKind = skScalar
        OutRows = 1
        OutCols = 1
        For I = 1 To Count
            If Not CORE_Array.AccumulateShape(OutKind, OutRows, OutCols, Kinds(I), RowCounts(I), _
                ColCounts(I), Condition) Then
                EvaluateCall = CORE_Err.ErrForCondition(Condition)
                Exit Function
            End If
        Next I
    'Stage 4: the capacity, from the dimensions alone
        If Not CORE_Array.CheckCapacity(OutRows, OutCols, Condition) Then
            EvaluateCall = CORE_Err.ErrForCondition(Condition)
            Exit Function
        End If

'------------------------------------------------------------------------------
' READ ARGUMENTS
'------------------------------------------------------------------------------
    'Stage 5: only now is any content read
        For I = 1 To Count
            If Not M_Ranges.TryReadArgument(Values(LBound(Values) + I - 1), Payloads(I), Kinds(I), _
                RowCounts(I), ColCounts(I), Condition) Then
                EvaluateCall = CORE_Err.ErrForCondition(Condition)
                Exit Function
            End If
        Next I

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'An all-scalar call returns a scalar
        If OutKind = skScalar Then
            For I = 1 To Count
                Elements(I) = Payloads(I)
            Next I
            EvaluateCall = EvaluateElement(Fn, Elements, Flag, Mode)
            Exit Function
        End If
    'Any other call returns a two-dimensional array, evaluated by rows
        If Not CORE_Array.TryAllocateOutput(OutRows, OutCols, Result, Condition) Then
            EvaluateCall = CORE_Err.ErrForCondition(Condition)
            Exit Function
        End If
        For R = 1 To OutRows
            For C = 1 To OutCols
                For I = 1 To Count
                    Elements(I) = CORE_Array.ElementAt(Payloads(I), Kinds(I), R, C)
                Next I
                Result(R, C) = EvaluateElement(Fn, Elements, Flag, Mode)
            Next C
        Next R
        EvaluateCall = Result

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the containment handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Contained:
    'A defect, contained as #VALUE! (data contract, Internal failures)
        EvaluateCall = CORE_Err.ErrValue()

End Function

Private Function TryResolveControl( _
    ByVal Fn As KPRDateFunction, _
    ByVal Control As Variant, _
    ByRef Flag As Boolean, _
    ByRef Mode As KPRPillarRounding, _
    ByRef Failure As Variant) As Boolean
'
'------------------------------------------------------------------------------
'                         RESOLVE THE OPTIONAL CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Reads and parses the optional control of a function, or selects its default.
'
' WHY THIS EXISTS
'   Every control follows the same rules (data contract, Control values): Empty, Missing and
'   an omitted control select the default; an error passed in is returned unchanged; a control
'   of more than one element is not a scalar.
'
' INPUTS
'   Fn
'     The function, which decides the kind and default of its control.
'
'   Control
'     The control as the worksheet passed it, or Missing.
'
'   Flag, Mode (ByRef, output)
'     The resolved Boolean control and rounding mode.
'
'   Failure (ByRef, output)
'     The call-level error when the control is rejected.
'
' RETURNS
'   Boolean
'     True with the control resolved; False with Failure.
'
' BEHAVIOR
'   Selects the default for no control or Missing, reads the control through M_Ranges,
'   returns an error passed in, then parses it with CORE_Parse.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   ControlKind, DefaultBoolean, M_Ranges.TryReadControl, CORE_Parse, CORE_Err. Data contract:
'   Function arguments, Control values
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
    Const PROC_NAME     As String = "KPR_Dates.TryResolveControl"

    Dim ScalarOut       As Variant      'The control's single element
    Dim Condition       As KPRCondition 'Condition of a rejected control
    Dim Token           As String       'Normalized rounding token
    Dim DefaultFlag     As Boolean      'The documented default of a Boolean control

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Defaults until the control is read
        TryResolveControl = False
        Mode = prNearest
        DefaultFlag = DefaultBoolean(Fn)
        Flag = DefaultFlag

'------------------------------------------------------------------------------
' READ CONTROL
'------------------------------------------------------------------------------
    'A function without a control, or a control forwarded as Missing, takes the default
        If ControlKind(Fn) = ckNone Then
            TryResolveControl = True
            Exit Function
        End If
        If IsMissing(Control) Then
            TryResolveControl = True
            Exit Function
        End If
    'One cell or element at most, never read when it is more
        If Not M_Ranges.TryReadControl(Control, ScalarOut, Condition) Then
            Failure = CORE_Err.ErrForCondition(Condition)
            Exit Function
        End If
    'An error passed as the control is returned unchanged for the whole call
        If VarType(ScalarOut) = vbError Then
            Failure = ScalarOut
            Exit Function
        End If

'------------------------------------------------------------------------------
' PARSE CONTROL
'------------------------------------------------------------------------------
    'Parse the control by its kind
        If ControlKind(Fn) = ckBoolean Then
            If Not CORE_Parse.TryParseBoolControl(ScalarOut, DefaultFlag, Flag, Condition) Then
                Failure = CORE_Err.ErrForCondition(Condition)
                Exit Function
            End If
        Else
            If Not CORE_Parse.TryParseRoundingControl(ScalarOut, Token, Condition) Then
                Failure = CORE_Err.ErrForCondition(Condition)
                Exit Function
            End If
            Select Case Token
                Case "FLOOR"
                    Mode = prFloor
                Case "CEILING"
                    Mode = prCeiling
                Case Else
                    Mode = prNearest
            End Select
        End If
        TryResolveControl = True

End Function

Private Function ControlKind(ByVal Fn As KPRDateFunction) As KPRControlKind

'
'------------------------------------------------------------------------------
'                          NAME THE KIND OF A CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the kind of optional control a function takes.
'
' WHY THIS EXISTS
'   The control of each function is stated once.
'
' INPUTS
'   Fn
'     The function.
'
' RETURNS
'   KPRControlKind
'
' BEHAVIOR
'   Looks the function up.
'
' ERROR POLICY
'   None.
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'By function
        Select Case Fn
            Case dfDayOfWeek, dfAddMonths, dfAddYears, dfNthWeekdayOfMonth, dfLastWeekdayOfMonth
                ControlKind = ckBoolean
            Case dfPillarFromDates
                ControlKind = ckRounding
            Case Else
                ControlKind = ckNone
        End Select

End Function

Private Function DefaultBoolean(ByVal Fn As KPRDateFunction) As Boolean

'
'------------------------------------------------------------------------------
'                    NAME THE DEFAULT OF A BOOLEAN CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the documented default of a function's Boolean control.
'
' WHY THIS EXISTS
'   The control of each function is stated once.
'
' INPUTS
'   Fn
'     The function.
'
' RETURNS
'   Boolean
'
' BEHAVIOR
'   Looks the function up.
'
' ERROR POLICY
'   None.
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'By function
        Select Case Fn
            Case dfDayOfWeek, dfNthWeekdayOfMonth, dfLastWeekdayOfMonth
                DefaultBoolean = True
            Case Else
                DefaultBoolean = False
        End Select

End Function

'
'------------------------------------------------------------------------------
'
'                                   ELEMENTS
'
'------------------------------------------------------------------------------
'

Private Function EvaluateElement( _
    ByVal Fn As KPRDateFunction, _
    ByRef Elements() As Variant, _
    ByVal Flag As Boolean, _
    ByVal Mode As KPRPillarRounding) As Variant
'
'------------------------------------------------------------------------------
'                            EVALUATE ONE POSITION
'------------------------------------------------------------------------------
' PURPOSE
'   Evaluates a function at one position of its result, from the values of its arguments there.
'
' WHY THIS EXISTS
'   Each function reads its arguments in signature order and applies one routine of
'   CORE_Dates, so the per-element work of the 21 value-taking functions is one dispatch.
'
' INPUTS
'   Fn
'     The function.
'
'   Elements()
'     The value of each argument at the position, in signature order.
'
'   Flag, Mode
'     The resolved control.
'
' RETURNS
'   Variant
'     The result at the position, or the Excel error of the first argument that fails, of the
'     rule that rejects the element, or of an error passed in.
'
' BEHAVIOR
'   Reads each argument with its reader: a date, a whole number, or a whole number with its
'   domain, checked as it is read; then calls CORE_Dates. A boundary is gated by the window.
'
' ERROR POLICY
'   Does not raise. An unexpected runtime error is contained as #VALUE!.
'
' DEPENDENCIES
'   TryElementDate, TryElementLong, TryElementYear, TryElementDomain, TryElementPillar,
'   WindowedDate, CORE_Dates, CORE_Err. CALC-02 to CALC-09
'
' NOTES
'   A failed element never stops its neighbours: each position is evaluated on its own.
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "KPR_Dates.EvaluateElement"

    Dim D1              As Date         'First date argument
    Dim D2              As Date         'Second date argument
    Dim L1              As Long         'First whole-number argument
    Dim L2              As Long         'Second whole-number argument
    Dim L3              As Long         'Third whole-number argument
    Dim L4              As Long         'Fourth whole-number argument
    Dim Months          As Double       'Month delta of a pillar
    Dim Days            As Double       'Day delta of a pillar
    Dim Shifted         As Date         'A shifted or located date
    Dim Token           As String       'A formatted pillar
    Dim Failure         As Variant      'The error of the first failing argument
    Dim Condition       As KPRCondition 'Condition of a failing rule

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Contain an unexpected error as #VALUE!
        On Error GoTo Contained

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    'Read the arguments in signature order; the first that fails decides the element
        Select Case Fn
            Case dfDayOfWeek
                If Not TryElementDate(Elements(1), D1, Failure) Then GoTo Failed
                EvaluateElement = CORE_Dates.DayOfWeek(D1, Flag)
            Case dfDaysInMonth
                If Not TryElementDate(Elements(1), D1, Failure) Then GoTo Failed
                EvaluateElement = CORE_Dates.DaysInMonth(VBA.Year(D1), VBA.Month(D1))
            Case dfDaysInYear
                If Not TryElementYear(Elements(1), L1, Failure) Then GoTo Failed
                EvaluateElement = CORE_Dates.DaysInYear(L1)
            Case dfBeginOfMonth
                If Not TryElementDate(Elements(1), D1, Failure) Then GoTo Failed
                EvaluateElement = WindowedDate(CORE_Dates.BeginOfMonth(D1))
            Case dfEndOfMonth
                If Not TryElementDate(Elements(1), D1, Failure) Then GoTo Failed
                EvaluateElement = WindowedDate(CORE_Dates.EndOfMonth(D1))
            Case dfBeginOfQuarter
                If Not TryElementDate(Elements(1), D1, Failure) Then GoTo Failed
                EvaluateElement = WindowedDate(CORE_Dates.BeginOfQuarter(D1))
            Case dfEndOfQuarter
                If Not TryElementDate(Elements(1), D1, Failure) Then GoTo Failed
                EvaluateElement = WindowedDate(CORE_Dates.EndOfQuarter(D1))
            Case dfBeginOfYear
                If Not TryElementDate(Elements(1), D1, Failure) Then GoTo Failed
                EvaluateElement = WindowedDate(CORE_Dates.BeginOfYear(D1))
            Case dfEndOfYear
                If Not TryElementDate(Elements(1), D1, Failure) Then GoTo Failed
                EvaluateElement = WindowedDate(CORE_Dates.EndOfYear(D1))
            Case dfIsMonthEnd
                If Not TryElementDate(Elements(1), D1, Failure) Then GoTo Failed
                EvaluateElement = CORE_Dates.IsMonthEnd(D1)
            Case dfIsQuarterEnd
                If Not TryElementDate(Elements(1), D1, Failure) Then GoTo Failed
                EvaluateElement = CORE_Dates.IsQuarterEnd(D1)
            Case dfIsYearEnd
                If Not TryElementDate(Elements(1), D1, Failure) Then GoTo Failed
                EvaluateElement = CORE_Dates.IsYearEnd(D1)
            Case dfIsLeapYear
                If Not TryElementYear(Elements(1), L1, Failure) Then GoTo Failed
                EvaluateElement = CORE_Dates.IsLeapYear(L1)
            Case dfAddDays
                If Not TryElementDate(Elements(1), D1, Failure) Then GoTo Failed
                If Not TryElementLong(Elements(2), L1, Failure) Then GoTo Failed
                If Not CORE_Dates.TryAddDays(D1, VBA.CDbl(L1), Shifted) Then GoTo OutOfWindow
                EvaluateElement = Shifted
            Case dfAddWeeks
                If Not TryElementDate(Elements(1), D1, Failure) Then GoTo Failed
                If Not TryElementLong(Elements(2), L1, Failure) Then GoTo Failed
                If Not CORE_Dates.TryAddDays(D1, 7# * VBA.CDbl(L1), Shifted) Then GoTo OutOfWindow
                EvaluateElement = Shifted
            Case dfAddMonths
                If Not TryElementDate(Elements(1), D1, Failure) Then GoTo Failed
                If Not TryElementLong(Elements(2), L1, Failure) Then GoTo Failed
                If Not CORE_Dates.TryAddMonths(D1, VBA.CDbl(L1), Flag, Shifted) Then GoTo OutOfWindow
                EvaluateElement = Shifted
            Case dfAddYears
                If Not TryElementDate(Elements(1), D1, Failure) Then GoTo Failed
                If Not TryElementLong(Elements(2), L1, Failure) Then GoTo Failed
                If Not CORE_Dates.TryAddMonths(D1, 12# * VBA.CDbl(L1), Flag, Shifted) Then GoTo OutOfWindow
                EvaluateElement = Shifted
            Case dfNthWeekdayOfMonth
                If Not TryElementDomain(Elements(1), dmYear, L1, Failure) Then GoTo Failed
                If Not TryElementDomain(Elements(2), dmMonth, L2, Failure) Then GoTo Failed
                If Not TryElementDomain(Elements(3), dmWeekday, L3, Failure) Then GoTo Failed
                If Not TryElementDomain(Elements(4), dmOccurrence, L4, Failure) Then GoTo Failed
                If Not CORE_Dates.TryNthWeekday(L1, L2, L3, L4, Flag, Shifted, Condition) Then GoTo Rejected
                EvaluateElement = Shifted
            Case dfLastWeekdayOfMonth
                If Not TryElementDomain(Elements(1), dmYear, L1, Failure) Then GoTo Failed
                If Not TryElementDomain(Elements(2), dmMonth, L2, Failure) Then GoTo Failed
                If Not TryElementDomain(Elements(3), dmWeekday, L3, Failure) Then GoTo Failed
                If Not CORE_Dates.TryLastWeekday(L1, L2, L3, Flag, Shifted, Condition) Then GoTo Rejected
                EvaluateElement = Shifted
            Case dfPillarFromDates
                If Not TryElementDate(Elements(1), D1, Failure) Then GoTo Failed
                If Not TryElementDate(Elements(2), D2, Failure) Then GoTo Failed
                If Not CORE_Dates.TryFormatPillar(D1, D2, Mode, Token, Condition) Then GoTo Rejected
                EvaluateElement = Token
            Case dfDateFromPillar
                If Not TryElementDate(Elements(1), D1, Failure) Then GoTo Failed
                If Not TryElementPillar(Elements(2), Months, Days, Failure) Then GoTo Failed
                If Not CORE_Dates.TryAddPillar(D1, Months, Days, Shifted, Condition) Then GoTo Rejected
                EvaluateElement = Shifted
            Case Else
                'Every function the facade evaluates has its own case above
        End Select

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Done
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'An argument was rejected, or passed an error in
        EvaluateElement = Failure
        Exit Function
OutOfWindow:
    'A valid shift whose result leaves the window
        EvaluateElement = CORE_Err.ErrForCondition(cdResultWindow)
        Exit Function
Rejected:
    'A rule of the calculation rejected the element
        EvaluateElement = CORE_Err.ErrForCondition(Condition)
        Exit Function
Contained:
    'A defect, contained as #VALUE! (data contract, Internal failures)
        EvaluateElement = CORE_Err.ErrValue()

End Function

Private Function TryElementDate( _
    ByVal ValueIn As Variant, _
    ByRef DateOut As Date, _
    ByRef Failure As Variant) As Boolean
'
'------------------------------------------------------------------------------
'                             READ A DATE ARGUMENT
'------------------------------------------------------------------------------
' PURPOSE
'   Reads a date argument at one position.
'
' WHY THIS EXISTS
'   Every argument of its kind is read the same way, in the order of the data contract.
'
' INPUTS
'   ValueIn
'     The value of the argument at one position.
'
'   DateOut (ByRef, output)
'     The date-only, in-window date.
'
'   Failure (ByRef, output)
'     The Excel error of a rejection, or the error passed in.
'
' RETURNS
'   Boolean
'     True with the date; False with Failure.
'
' BEHAVIOR
'   Unwraps the scalar, returns an error passed in unchanged, then parses the value.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   CORE_Array.TryUnwrapScalar, CORE_Parse.TryParseDateScalar, CORE_Err. Data contract: Date values
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
    Const PROC_NAME     As String = "KPR_Dates.TryElementDate"

    Dim ScalarOut       As Variant      'The element's single value
    Dim Condition       As KPRCondition 'Condition of a rejected value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default outcome is failure
        TryElementDate = False

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Isolate the scalar; a payload element is always one
        If Not CORE_Array.TryUnwrapScalar(ValueIn, ScalarOut) Then
            Failure = CORE_Err.ErrForCondition(cdShapeUnsupported)
            Exit Function
        End If
    'An error passed in is returned unchanged (INPUT_ERROR_PROPAGATED)
        If VarType(ScalarOut) = vbError Then
            Failure = ScalarOut
            Exit Function
        End If
    'Parse the date under the data contract's Date values
        If Not CORE_Parse.TryParseDateScalar(ScalarOut, DateOut, Condition) Then
            Failure = CORE_Err.ErrForCondition(Condition)
            Exit Function
        End If
        TryElementDate = True

End Function

Private Function TryElementLong( _
    ByVal ValueIn As Variant, _
    ByRef LongOut As Long, _
    ByRef Failure As Variant) As Boolean
'
'------------------------------------------------------------------------------
'                         READ A WHOLE-NUMBER ARGUMENT
'------------------------------------------------------------------------------
' PURPOSE
'   Reads a whole-number argument at one position.
'
' WHY THIS EXISTS
'   Every argument of its kind is read the same way, in the order of the data contract.
'
' INPUTS
'   ValueIn
'     The value of the argument at one position.
'
'   LongOut (ByRef, output)
'     The whole number.
'
'   Failure (ByRef, output)
'     The Excel error of a rejection, or the error passed in.
'
' RETURNS
'   Boolean
'     True with the whole number; False with Failure.
'
' BEHAVIOR
'   Unwraps the scalar, returns an error passed in unchanged, then parses the value.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   CORE_Array.TryUnwrapScalar, CORE_Parse.TryParseLongScalar, CORE_Err. Data contract: Integer values
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
    Const PROC_NAME     As String = "KPR_Dates.TryElementLong"

    Dim ScalarOut       As Variant      'The element's single value
    Dim Condition       As KPRCondition 'Condition of a rejected value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default outcome is failure
        TryElementLong = False

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Isolate the scalar; a payload element is always one
        If Not CORE_Array.TryUnwrapScalar(ValueIn, ScalarOut) Then
            Failure = CORE_Err.ErrForCondition(cdShapeUnsupported)
            Exit Function
        End If
    'An error passed in is returned unchanged (INPUT_ERROR_PROPAGATED)
        If VarType(ScalarOut) = vbError Then
            Failure = ScalarOut
            Exit Function
        End If
    'Parse the whole number under the data contract's Integer values
        If Not CORE_Parse.TryParseLongScalar(ScalarOut, LongOut, Condition) Then
            Failure = CORE_Err.ErrForCondition(Condition)
            Exit Function
        End If
        TryElementLong = True

End Function

Private Function TryElementDomain( _
    ByVal ValueIn As Variant, _
    ByVal Domain As KPRDomain, _
    ByRef LongOut As Long, _
    ByRef Failure As Variant) As Boolean
'
'------------------------------------------------------------------------------
'                      READ A WHOLE NUMBER IN ITS DOMAIN
'------------------------------------------------------------------------------
' PURPOSE
'   Reads a whole-number argument and checks it against the domain of its argument.
'
' WHY THIS EXISTS
'   The contract checks each argument's domain as it is read, so that the first failing
'   argument of a position decides its error.
'
' INPUTS
'   ValueIn
'     The value of the argument at one position.
'
'   Domain
'     dmYear, dmMonth, dmWeekday or dmOccurrence.
'
'   LongOut (ByRef, output)
'     The whole number.
'
'   Failure (ByRef, output)
'     The Excel error of a rejection.
'
' RETURNS
'   Boolean
'     True with the whole number inside its domain; False with Failure.
'
' BEHAVIOR
'   Reads the whole number with TryElementLong, then checks its domain with CORE_Dates.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   TryElementLong, CORE_Dates.Check*, CORE_Err. Data contract: Integer values
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
    Const PROC_NAME     As String = "KPR_Dates.TryElementDomain"

    Dim Condition       As KPRCondition 'Domain condition

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Parse, then check the domain before the next argument is read
        TryElementDomain = False
        If Not TryElementLong(ValueIn, LongOut, Failure) Then Exit Function
        Select Case Domain
            Case dmYear
                Condition = CORE_Dates.CheckYear(LongOut)
            Case dmMonth
                Condition = CORE_Dates.CheckMonth(LongOut)
            Case dmWeekday
                Condition = CORE_Dates.CheckWeekday(LongOut)
            Case Else
                Condition = CORE_Dates.CheckOccurrence(LongOut)
        End Select
        If Condition <> cdNone Then
            Failure = CORE_Err.ErrForCondition(Condition)
            Exit Function
        End If
        TryElementDomain = True

End Function

Private Function TryElementYear( _
    ByVal ValueIn As Variant, _
    ByRef YearOut As Long, _
    ByRef Failure As Variant) As Boolean
'
'------------------------------------------------------------------------------
'                             READ A YEAR ARGUMENT
'------------------------------------------------------------------------------
' PURPOSE
'   Reads a calendar-year argument, 1900 to 9999.
'
' WHY THIS EXISTS
'   DaysInYear and IsLeapYear take a year, never a date.
'
' INPUTS
'   ValueIn
'     The value of the argument at one position.
'
'   YearOut (ByRef, output)
'     The year.
'
'   Failure (ByRef, output)
'     The Excel error of a rejection.
'
' RETURNS
'   Boolean
'     True with the year; False with Failure.
'
' BEHAVIOR
'   Reads the year through TryElementDomain with dmYear.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   TryElementDomain. CALC-03 step 5
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'A year is a whole number in its domain
        TryElementYear = TryElementDomain(ValueIn, dmYear, YearOut, Failure)

End Function

Private Function TryElementPillar( _
    ByVal ValueIn As Variant, _
    ByRef Months As Double, _
    ByRef Days As Double, _
    ByRef Failure As Variant) As Boolean
'
'------------------------------------------------------------------------------
'                            READ A PILLAR ARGUMENT
'------------------------------------------------------------------------------
' PURPOSE
'   Reads a pillar-token argument at one position.
'
' WHY THIS EXISTS
'   Every argument of its kind is read the same way, in the order of the data contract.
'
' INPUTS
'   ValueIn
'     The value of the argument at one position.
'
'   Months, Days (ByRef, output)
'     The month and day deltas of the token.
'
'   Failure (ByRef, output)
'     The Excel error of a rejection, or the error passed in.
'
' RETURNS
'   Boolean
'     True with both deltas; False with Failure.
'
' BEHAVIOR
'   Unwraps the scalar, returns an error passed in unchanged, then parses the value.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   CORE_Array.TryUnwrapScalar, CORE_Dates.TryParsePillar, CORE_Err. Data contract: Pillar tokens
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
    Const PROC_NAME     As String = "KPR_Dates.TryElementPillar"

    Dim ScalarOut       As Variant      'The element's single value
    Dim Condition       As KPRCondition 'Condition of a rejected value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default outcome is failure
        TryElementPillar = False

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Isolate the scalar; a payload element is always one
        If Not CORE_Array.TryUnwrapScalar(ValueIn, ScalarOut) Then
            Failure = CORE_Err.ErrForCondition(cdShapeUnsupported)
            Exit Function
        End If
    'An error passed in is returned unchanged (INPUT_ERROR_PROPAGATED)
        If VarType(ScalarOut) = vbError Then
            Failure = ScalarOut
            Exit Function
        End If
    'Parse the token under the data contract's Pillar tokens
        If Not CORE_Dates.TryParsePillar(ScalarOut, Months, Days, Condition) Then
            Failure = CORE_Err.ErrForCondition(Condition)
            Exit Function
        End If
        TryElementPillar = True

End Function

Private Function WindowedDate(ByVal DateIn As Date) As Variant

'
'------------------------------------------------------------------------------
'                        GATE A BOUNDARY BY THE WINDOW
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a boundary date, or #NUM! when it is outside the window.
'
' WHY THIS EXISTS
'   A boundary is a result and is gated like one (CALC-04 step 4): BeginOfYear of a 1900 date
'   is 1900-01-01, outside the window.
'
' INPUTS
'   DateIn
'     A boundary date that CORE_Dates built.
'
' RETURNS
'   Variant
'     The date, or #NUM! (RESULT_WINDOW).
'
' BEHAVIOR
'   Tests the date with CORE_Dates.IsDateInWindow.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   CORE_Dates.IsDateInWindow, CORE_Err. CALC-04 step 4
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Inside the window, the date; outside, RESULT_WINDOW
        If CORE_Dates.IsDateInWindow(DateIn) Then
            WindowedDate = DateIn
        Else
            WindowedDate = CORE_Err.ErrForCondition(cdResultWindow)
        End If

End Function
