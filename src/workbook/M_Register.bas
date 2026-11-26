Attribute VB_Name = "M_Register"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: M_Register
'------------------------------------------------------------------------------
' PURPOSE
'   The Excel adapter that registers the 22 KPR_Dates_* worksheet functions with the host: a
'   category, a description and one description per argument in the Insert Function dialog.
'
' WHY THIS EXISTS
'   Excel shows a worksheet function in the Insert Function dialog only with what
'   Application.MacroOptions gives it. Reaching the host's registry touches the Excel object
'   model, which the facade and the core may not.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - RegisterFunctions    registers the 22 functions under the category CATEGORY_NAME
'     - UnregisterFunctions  returns the 22 functions to the User Defined category with no
'                            description
'   Both are repeatable: every call sets the same options again, so no call adds an entry.
'   The descriptions are the purposes and the argument texts of the facade's documentation,
'   which restate calculation contracts CALC-02 to CALC-10.
'
' ERROR POLICY
'   No member raises; a failure of one function is counted, never raised, and the others are
'   still processed.
'
' DEPENDENCIES
'   The Excel object model (Application.MacroOptions, ThisWorkbook.Name). Calculation
'   contracts: CALC-02 to CALC-10; public API: docs/PUBLIC_API.txt, [supported]
'
' NOTES
'   - The names are those of the facade: the functions are addressed through the workbook
'     that holds this module, never through the active workbook
'   - Excel keeps a category name until it closes: UnregisterFunctions empties the category
'     but cannot delete its name from the running session
'   - Excel limits a description to 255 characters
'   - Option Private Module keeps every member out of the Excel function list
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PUBLIC CONSTANTS
'------------------------------------------------------------------------------
    'The number of functions the module registers
    Public Const FUNCTION_COUNT As Long = 22    'The 22 KPR_Dates_* functions

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'The category of the Insert Function dialog
    Private Const CATEGORY_NAME As String = "KPR Dates"
    'The User Defined category, where a function with no registration stands
    Private Const CATEGORY_USER_DEFINED As Long = 15

    'Argument descriptions that several functions share
    Private Const ARG_DATE       As String = "A date: a VBA Date, a 1900-system serial or exact YYYY-MM-DD text; a scalar, an array or a Range."
    Private Const ARG_YEAR_ONLY  As String = "A calendar year, 1900 to 9999, never a date; a scalar, an array or a Range."
    Private Const ARG_YEAR       As String = "A calendar year, 1900 to 9999; a scalar, an array or a Range."
    Private Const ARG_WEEK_BASE  As String = "True: Monday is 1. False: Sunday is 1, as WEEKDAY. A Boolean, or Empty for the default; one cell or element at most, default True."
    Private Const ARG_KEEP_EOM   As String = "True: a month end moves to the target month's end. A Boolean, or Empty for the default; one cell or element at most, default False."
    Private Const ARG_MONTH      As String = "A month, 1 to 12; a scalar, an array or a Range."
    Private Const ARG_WEEKDAY    As String = "A weekday number, 1 to 7, in the weekday base; a scalar, an array or a Range."

'
'------------------------------------------------------------------------------
'
'                       WORKSHEET FUNCTION REGISTRATION
'
'------------------------------------------------------------------------------
'

Public Function RegisterFunctions() As Long

'
'------------------------------------------------------------------------------
'                            REGISTER THE FUNCTIONS
'------------------------------------------------------------------------------
' PURPOSE
'   Registers the 22 worksheet functions: category, description and argument descriptions.
'
' WHY THIS EXISTS
'   The Insert Function dialog and the argument help read what MacroOptions stores.
'
' INPUTS
'   None
'
' RETURNS
'   Long
'     The number of functions registered, FUNCTION_COUNT when every one was.
'
' BEHAVIOR
'   Describes each function in turn and gives MacroOptions its category, description and
'   argument descriptions. A function that fails does not stop the others.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   Describe, ApplyOptions
'
' NOTES
'   Repeatable: a second call sets the same options and changes nothing else
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Index           As Long         'Position of the function, 1 to FUNCTION_COUNT
    Dim FunctionName    As String       'Name of the function
    Dim Summary         As String       'Description of the function
    Dim ArgTexts        As Variant      'Argument descriptions, Empty for none

'------------------------------------------------------------------------------
' REGISTER
'------------------------------------------------------------------------------
    'Register every function, counting those the host accepted
        For Index = 1 To FUNCTION_COUNT
            Describe Index, FunctionName, Summary, ArgTexts
            If ApplyOptions(FunctionName, Summary, CATEGORY_NAME, ArgTexts) Then
                RegisterFunctions = RegisterFunctions + 1
            End If
        Next Index

End Function

Public Function UnregisterFunctions() As Long

'
'------------------------------------------------------------------------------
'                           UNREGISTER THE FUNCTIONS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the 22 worksheet functions to the User Defined category with no description.
'
' WHY THIS EXISTS
'   A workbook that is closed, or a template that is removed, must not leave descriptions
'   behind in the dialog.
'
' INPUTS
'   None
'
' RETURNS
'   Long
'     The number of functions cleared, FUNCTION_COUNT when every one was.
'
' BEHAVIOR
'   Describes each function in turn, replaces its category with the User Defined category,
'   its description with an empty text and each argument description with an empty text. A
'   function that fails does not stop the others.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   Describe, ApplyOptions
'
' NOTES
'   Repeatable. The category name stays in the list until Excel closes
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Index           As Long         'Position of the function, 1 to FUNCTION_COUNT
    Dim Position        As Long         'Position of an argument
    Dim FunctionName    As String       'Name of the function
    Dim Summary         As String       'Description of the function
    Dim ArgTexts        As Variant      'Argument descriptions, Empty for none

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
    'Clear every function, counting those the host accepted
        For Index = 1 To FUNCTION_COUNT
            Describe Index, FunctionName, Summary, ArgTexts
            If Not IsEmpty(ArgTexts) Then
                For Position = LBound(ArgTexts) To UBound(ArgTexts)
                    ArgTexts(Position) = vbNullString
                Next Position
            End If
            If ApplyOptions(FunctionName, vbNullString, CATEGORY_USER_DEFINED, ArgTexts) Then
                UnregisterFunctions = UnregisterFunctions + 1
            End If
        Next Index

End Function

Private Sub Describe( _
    ByVal Index As Long, _
    ByRef FunctionName As String, _
    ByRef Summary As String, _
    ByRef ArgTexts As Variant)

'
'------------------------------------------------------------------------------
'                             DESCRIBE A FUNCTION
'------------------------------------------------------------------------------
' PURPOSE
'   Gives the name, the description and the argument descriptions of one function.
'
' WHY THIS EXISTS
'   The texts are written once, in signature order, and both procedures read them.
'
' INPUTS
'   Index
'     The position of the function, 1 to FUNCTION_COUNT, in the order of docs/PUBLIC_API.txt.
'
'   FunctionName, Summary, ArgTexts (ByRef, output)
'     The function name, its description and one description per argument, in signature
'     order; ArgTexts is Empty for a function with no argument.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Selects the function by its position. The descriptions are the purposes and the argument
'   texts of the facade's documentation.
'
' ERROR POLICY
'   Does not raise. An Index outside 1 to FUNCTION_COUNT returns an empty name.
'
' DEPENDENCIES
'   None
'
' NOTES
'   - A description holds at most 255 characters, the limit of the host
'   - tests/test_register_contract.py compares every text with the facade and its signature
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'An unknown position describes nothing
        FunctionName = vbNullString
        Summary = vbNullString
        ArgTexts = Empty

'------------------------------------------------------------------------------
' DESCRIBE
'------------------------------------------------------------------------------
    'Select the function
        Select Case Index
        Case 1
            FunctionName = "KPR_Dates_DayOfWeek"
            Summary = "Returns the number of the weekday of a date, 1 to 7, Monday-based by default."
            ArgTexts = VBA.Array( _
                ARG_DATE, _
                ARG_WEEK_BASE)
        Case 2
            FunctionName = "KPR_Dates_DaysInMonth"
            Summary = "Returns the number of days of the month that holds a date."
            ArgTexts = VBA.Array( _
                ARG_DATE)
        Case 3
            FunctionName = "KPR_Dates_DaysInYear"
            Summary = "Returns the number of days of a calendar year."
            ArgTexts = VBA.Array( _
                ARG_YEAR_ONLY)
        Case 4
            FunctionName = "KPR_Dates_BeginOfMonth"
            Summary = "Returns the first day of the month that holds a date."
            ArgTexts = VBA.Array( _
                ARG_DATE)
        Case 5
            FunctionName = "KPR_Dates_EndOfMonth"
            Summary = "Returns the last day of the month that holds a date."
            ArgTexts = VBA.Array( _
                ARG_DATE)
        Case 6
            FunctionName = "KPR_Dates_BeginOfQuarter"
            Summary = "Returns the first day of the calendar quarter that holds a date."
            ArgTexts = VBA.Array( _
                ARG_DATE)
        Case 7
            FunctionName = "KPR_Dates_EndOfQuarter"
            Summary = "Returns the last day of the calendar quarter that holds a date."
            ArgTexts = VBA.Array( _
                ARG_DATE)
        Case 8
            FunctionName = "KPR_Dates_BeginOfYear"
            Summary = "Returns 1 January of the year that holds a date."
            ArgTexts = VBA.Array( _
                ARG_DATE)
        Case 9
            FunctionName = "KPR_Dates_EndOfYear"
            Summary = "Returns 31 December of the year that holds a date."
            ArgTexts = VBA.Array( _
                ARG_DATE)
        Case 10
            FunctionName = "KPR_Dates_IsMonthEnd"
            Summary = "Reports whether a date is the last day of its month."
            ArgTexts = VBA.Array( _
                ARG_DATE)
        Case 11
            FunctionName = "KPR_Dates_IsQuarterEnd"
            Summary = "Reports whether a date is the last day of its calendar quarter."
            ArgTexts = VBA.Array( _
                ARG_DATE)
        Case 12
            FunctionName = "KPR_Dates_IsYearEnd"
            Summary = "Reports whether a date is 31 December."
            ArgTexts = VBA.Array( _
                ARG_DATE)
        Case 13
            FunctionName = "KPR_Dates_IsLeapYear"
            Summary = "Reports whether a calendar year is a Gregorian leap year."
            ArgTexts = VBA.Array( _
                ARG_YEAR_ONLY)
        Case 14
            FunctionName = "KPR_Dates_AddDays"
            Summary = "Moves a date by a signed number of calendar days."
            ArgTexts = VBA.Array( _
                ARG_DATE, _
                "A signed whole number of days; a scalar, an array or a Range.")
        Case 15
            FunctionName = "KPR_Dates_AddWeeks"
            Summary = "Moves a date by a signed number of weeks of seven calendar days."
            ArgTexts = VBA.Array( _
                ARG_DATE, _
                "A signed whole number of weeks; a scalar, an array or a Range.")
        Case 16
            FunctionName = "KPR_Dates_AddMonths"
            Summary = "Moves a date by a signed number of months, clipping a day that the target month lacks."
            ArgTexts = VBA.Array( _
                ARG_DATE, _
                "A signed whole number of months; a scalar, an array or a Range.", _
                ARG_KEEP_EOM)
        Case 17
            FunctionName = "KPR_Dates_AddYears"
            Summary = "Moves a date by a signed number of years, as twelve months each."
            ArgTexts = VBA.Array( _
                ARG_DATE, _
                "A signed whole number of years; a scalar, an array or a Range.", _
                ARG_KEEP_EOM)
        Case 18
            FunctionName = "KPR_Dates_NthWeekdayOfMonth"
            Summary = "Returns the date of the nth given weekday of a month."
            ArgTexts = VBA.Array( _
                ARG_YEAR, _
                ARG_MONTH, _
                ARG_WEEKDAY, _
                "The occurrence, 1 to 5; a scalar, an array or a Range.", _
                ARG_WEEK_BASE)
        Case 19
            FunctionName = "KPR_Dates_LastWeekdayOfMonth"
            Summary = "Returns the date of the last given weekday of a month."
            ArgTexts = VBA.Array( _
                ARG_YEAR, _
                ARG_MONTH, _
                ARG_WEEKDAY, _
                ARG_WEEK_BASE)
        Case 20
            FunctionName = "KPR_Dates_PillarFromDates"
            Summary = "Returns the pillar token that labels the interval between two dates, rounded to a curve bucket."
            ArgTexts = VBA.Array( _
                ARG_DATE, _
                ARG_DATE, _
                "NEAREST, FLOOR or CEILING, in any case. Text, or Empty for the default; one cell or element at most, default NEAREST.")
        Case 21
            FunctionName = "KPR_Dates_DateFromPillar"
            Summary = "Returns the date that a pillar token names from a start date."
            ArgTexts = VBA.Array( _
                ARG_DATE, _
                "A pillar token, such as 3M, 1Y6M, 2W3D, ON or TN; a scalar, an array or a Range.")
        Case 22
            FunctionName = "KPR_Dates_HostDateSystem"
            Summary = "Returns the date system of the calling workbook: 1900 or 1904."
            ArgTexts = Empty
        End Select

End Sub

Private Function ApplyOptions( _
    ByVal FunctionName As String, _
    ByVal Summary As String, _
    ByVal Category As Variant, _
    ByRef ArgTexts As Variant) As Boolean

'
'------------------------------------------------------------------------------
'                        SET THE OPTIONS OF A FUNCTION
'------------------------------------------------------------------------------
' PURPOSE
'   Gives one function its options in the host.
'
' WHY THIS EXISTS
'   Both registration and unregistration reach MacroOptions the same way, and a host refusal
'   must become a result, not an error.
'
' INPUTS
'   FunctionName
'     The name of a function of the facade.
'
'   Summary
'     The description; an empty text clears it.
'
'   Category
'     A category name or number.
'
'   ArgTexts
'     One description per argument, or Empty for a function with no argument.
'
' RETURNS
'   Boolean
'     True when the host accepted the options; False when it refused them or the name is empty.
'
' BEHAVIOR
'   Addresses the function through the workbook that holds this module, with the name quoted
'   as a reference, and calls Application.MacroOptions. Arguments are passed only when the
'   function has some.
'
' ERROR POLICY
'   Does not raise. A host error is False, and Err is cleared.
'
' DEPENDENCIES
'   Application.MacroOptions, ThisWorkbook.Name
'
' NOTES
'   Excel refuses more argument descriptions than the function has arguments
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Macro           As String       'Workbook-qualified name of the function

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Refuse an empty name, and contain any host error
        ApplyOptions = False
        If Len(FunctionName) = 0 Then Exit Function
        On Error GoTo Refused

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
    'Address the function through this workbook, never the active one
        Macro = "'" & VBA.Replace(ThisWorkbook.Name, "'", "''") & "'!" & FunctionName

    'Set the options; a function with no argument gets none
        If IsEmpty(ArgTexts) Then
            Application.MacroOptions Macro:=Macro, Description:=Summary, Category:=Category
        Else
            Application.MacroOptions Macro:=Macro, Description:=Summary, Category:=Category, _
                ArgumentDescriptions:=ArgTexts
        End If
        ApplyOptions = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Refused:
    'The host refused the options
        Err.Clear
        ApplyOptions = False

End Function
