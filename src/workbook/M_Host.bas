Attribute VB_Name = "M_Host"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: M_Host
'------------------------------------------------------------------------------
' PURPOSE
'   The Excel adapter of the host rule: which date system the calling worksheet uses.
'
' WHY THIS EXISTS
'   Only a worksheet caller decides the date system (calculation contracts, Host and date
'   system). Reading the caller and its workbook touches the Excel object model, which the
'   facade and the core may not.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - TryReadHostDateSystem  1900 or 1904, or HOST_UNRESOLVED
'     - TryCheckHost           whether a value-taking call may proceed
'
' ERROR POLICY
'   No member raises; a failure is a return value with a condition.
'
' DEPENDENCIES
'   The Excel object model (Application.Caller, Range.Worksheet, Workbook.Date1904), CORE_Err.
'   Calculation contracts: Host and date system
'
' NOTES
'   - No other workbook, such as ActiveWorkbook or ThisWorkbook, is ever consulted
'   - A caller that is not a Range, or reading Application.Caller failing, means no worksheet
'     host: the 1900 serial contract applies
'   - Option Private Module keeps every member out of the Excel function list
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------


'
'------------------------------------------------------------------------------
'
'                                  HOST RULE
'
'------------------------------------------------------------------------------
'

Public Function TryReadHostDateSystem( _
    ByRef DateSystem As Long, _
    ByRef Condition As KPRCondition) As Boolean

'
'------------------------------------------------------------------------------
'                        READ THE CALLER'S DATE SYSTEM
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the date system of the worksheet that calls the function, or 1900 without one.
'
' WHY THIS EXISTS
'   A workbook on the 1904 date system numbers every day 1,462 lower; a value-taking
'   function must refuse it rather than return a shifted date.
'
' INPUTS
'   DateSystem (ByRef, output)
'     1900 or 1904.
'
'   Condition (ByRef, output)
'     cdNone, or cdHostUnresolved.
'
' RETURNS
'   Boolean
'     True with the date system; False with cdHostUnresolved when a worksheet caller was
'     identified and its workbook's date system could not be read.
'
' BEHAVIOR
'   Reads Application.Caller once. A Range identifies the worksheet host, whose workbook's
'   Date1904 decides; anything else, or a failure to read the caller, is no worksheet host.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   Application.Caller, Range.Worksheet, Workbook.Date1904. Calculation contracts: Host and
'   date system, steps 1 to 5
'
' NOTES
'   Only the sequence of an identified Range and a failed read gives HOST_UNRESOLVED
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CallerObject    As Object       'Application.Caller, when it is an object
    Dim CallerRange     As Range        'The calling Range
    Dim Is1904          As Boolean      'Date system of the caller's workbook

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Without a worksheet host the 1900 serial contract applies
        TryReadHostDateSystem = True
        Condition = cdNone
        DateSystem = 1900

'------------------------------------------------------------------------------
' CLASSIFY CALLER
'------------------------------------------------------------------------------
    'A caller that is not an object, or cannot be read, identifies no worksheet host
        On Error Resume Next
        Set CallerObject = Application.Caller
        On Error GoTo 0
        If CallerObject Is Nothing Then Exit Function
        If Not TypeOf CallerObject Is Range Then Exit Function

'------------------------------------------------------------------------------
' READ DATE SYSTEM
'------------------------------------------------------------------------------
    'A worksheet host was identified: its own workbook decides
        Set CallerRange = CallerObject
        On Error GoTo Unresolved
        Is1904 = CallerRange.Worksheet.Parent.Date1904
        On Error GoTo 0
        If Is1904 Then DateSystem = 1904

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Unresolved:
    'The identified host's date system could not be read
        Condition = cdHostUnresolved
        TryReadHostDateSystem = False

End Function

Public Function TryCheckHost(ByRef Condition As KPRCondition) As Boolean

'
'------------------------------------------------------------------------------
'                      CHECK THAT THE HOST ALLOWS A CALL
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a value-taking function may evaluate in the calling host.
'
' WHY THIS EXISTS
'   Stage 1 of every value-taking call refuses a 1904 or unreadable worksheet host before any
'   argument is read: once its serials are read, nothing can tell they meant other days.
'
' INPUTS
'   Condition (ByRef, output)
'     cdNone, cdHostDate1904 or cdHostUnresolved.
'
' RETURNS
'   Boolean
'     True for a 1900 worksheet host or no worksheet host.
'
' BEHAVIOR
'   Reads the date system and refuses 1904.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   TryReadHostDateSystem. Calculation contracts: Host and date system, step 3
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
    Dim DateSystem      As Long         '1900 or 1904

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
    'An unreadable host, or a 1904 one, refuses the call
        TryCheckHost = False
        If Not TryReadHostDateSystem(DateSystem, Condition) Then Exit Function
        If DateSystem = 1904 Then
            Condition = cdHostDate1904
            Exit Function
        End If
        TryCheckHost = True

End Function
