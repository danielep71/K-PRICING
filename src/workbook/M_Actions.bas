Attribute VB_Name = "M_Actions"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: M_Actions
'------------------------------------------------------------------------------
' PURPOSE
'   The Excel adapter behind every menu of the workbook: insert one of the 22 KPR_Dates_*
'   functions, register or clear their help in the Insert Function dialog, and, in a
'   development build, run the project's test suites. The Ribbon (M_Ribbon) and the toolbar
'   (M_CommandBars) both call these procedures, so the two surfaces offer the same actions
'   with the same behavior.
'
' WHY THIS EXISTS
'   A menu callback stays thin: it hands the action text of the clicked control to RunAction.
'   Each action here checks its preconditions, delegates to the module that owns the work and
'   tells the user the outcome.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Actions, each telling the user its outcome in a message box:
'     - InsertFunction   put one function in the active cell and open Function Arguments
'     - RegisterHelp     register the help of the 22 functions (M_Register)
'     - ClearHelp        clear that help; the functions stay callable
'     - RunTests         development build: run the project suites and show their results
'     - RunAction        run the action an action text names
'   Menu content, with no message and no change to Excel:
'     - FunctionNames    the 22 function names, in registration order
'     - ActionKnown      whether an action text names an action this build can run
'     - TestsAvailable   whether the test modules are part of this build
'
' ERROR POLICY
'   No action raises to the menu that called it: a failure is reported in a message box.
'   InsertFunction gives the cell back its content when it fails or when the dialog is
'   cancelled, and reports both outcomes when that restoration fails too.
'
' DEPENDENCIES
'   M_Register; the Excel object model (Selection, ActiveCell, Range.FunctionWizard,
'   Application.Run, MsgBox). The test modules only through Application.Run, by name
'
' NOTES
'   - An action text is ACTION_FUNCTION and a function name, or one of ACTION_REGISTER,
'     ACTION_CLEAR and ACTION_TESTS. The Ribbon stores it in a control's tag, the toolbar in a
'     button's Parameter; tools/check_ribbon.py keeps the Ribbon's tags equal to these
'     constants
'   - The test modules are part of the development build only. This module never names their
'     procedures in code: it reaches TEST_Suites through Application.Run, so a production
'     build compiles without them
'   - Option Private Module keeps every member out of the macro list and the function list
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE TYPES
'------------------------------------------------------------------------------
    Private Type CellContent                'What a cell held before an action wrote into it
        HadFormula      As Boolean          'The cell held a formula
        Formula         As String           'That formula's text
        Dynamic         As Boolean          'The formula was read through Formula2
        Value           As Variant          'A constant's Value2, error values included
        Prefix          As String           'PrefixCharacter: an apostrophe marks text
    End Type

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'The procedures of the test modules, reached by name
    Private Const TESTS_PROBE           As String = "TEST_Suites.SuitesLoaded"      'Answers True
    Private Const TESTS_RUNNER          As String = "TEST_Suites.RunProjectSuites"  'Runs them

    'The first line of a passing report of TESTS_RUNNER
    Private Const TESTS_PASSED          As String = "RESULT=PASS"                   'Pass line

'------------------------------------------------------------------------------
' PUBLIC CONSTANTS
'------------------------------------------------------------------------------
    'The action texts the menus store on their controls
    Public Const ACTION_FUNCTION        As String = "fn:"               'Followed by a name
    Public Const ACTION_REGISTER        As String = "setup:register"    'RegisterHelp
    Public Const ACTION_CLEAR           As String = "setup:clear"       'ClearHelp
    Public Const ACTION_TESTS           As String = "dev:tests"         'RunTests

    'The title of every message box of the menus
    Public Const UI_TITLE               As String = "K-PRICING"         'Message box title


'
'------------------------------------------------------------------------------
'
'                                   ACTIONS
'
'------------------------------------------------------------------------------
'

Public Sub InsertFunction(ByVal Chosen As String)

'
'------------------------------------------------------------------------------
'                         INSERT A FUNCTION IN A CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Puts one of the 22 functions in the active cell and opens Excel's Function Arguments
'   dialog for it.
'
' WHY THIS EXISTS
'   A wrong target must not lose content, and a cancelled dialog must leave the cell as it was.
'
' INPUTS
'   Chosen
'     A function name, such as KPR_Dates_EndOfMonth; the case does not matter.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Refuses a name outside the registration list, a selection that is not a cell and a locked
'   cell of a protected sheet before it writes anything. Records the cell's content, writes
'   "=<name>()" and, for a function with arguments, opens the dialog; a cancelled dialog
'   leaves the formula as written, and the cell gets its content back. The formula is
'   written through Formula2 where Excel has it, so that a result spills. A cell of another
'   workbook gets the name qualified with this workbook's name, unless this workbook is an
'   add-in.
'
' ERROR POLICY
'   Does not raise. An error gives the cell back its content, then reports the error and
'   whether that restoration succeeded.
'
' DEPENDENCIES
'   FunctionIndex, M_Register, SaveCell, WriteFormula, TryRestoreCell, FormulaQualifier,
'   Notify; Range.FunctionWizard
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_Actions.InsertFunction"

    Dim Index           As Long         'Position of the function in the registration list
    Dim Target          As Range        'The cell that receives the formula
    Dim Saved           As CellContent  'What the cell held on entry
    Dim Written         As Boolean      'The cell may no longer hold its content
    Dim Inserted        As String       'The formula as Excel spells it after the write
    Dim ErrNumber       As Long         'Number of the error that stopped the action
    Dim ErrDescription  As String       'Its description
    Dim Restored        As Boolean      'The cell got its content back

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Refuse before touching a cell: a wrong target must not lose content
        Index = FunctionIndex(Chosen)
        If Index = 0 Then
            Notify "Not a K-PRICING function: " & Chosen, vbExclamation
            Exit Sub
        End If
        If TypeName(Selection) <> "Range" Then
            Notify "Select a worksheet cell, then choose the function again.", vbExclamation
            Exit Sub
        End If
        Set Target = ActiveCell
        If Target.Worksheet.ProtectContents And Target.Locked Then
            Notify "The active cell is locked on a protected sheet.", vbExclamation
            Exit Sub
        End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Report a failure instead of raising it to the menu
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' INSERT
'------------------------------------------------------------------------------
    'Record the content first, so that a write that fails half way is restored
        SaveCell Target, Saved
        Written = True
        WriteFormula Target, "=" & FormulaQualifier(Target.Worksheet.Parent) & _
            M_Register.RegisteredName(Index) & "()"
    'A function with no argument needs no dialog
        If M_Register.RegisteredArgumentCount(Index) = 0 Then Exit Sub

    'A cancel leaves Excel's spelling of the formula unchanged; OK adds the arguments
        Inserted = VBA.CStr(Target.Formula)
        Target.FunctionWizard
        If VBA.StrComp(VBA.CStr(Target.Formula), Inserted, vbTextCompare) = 0 Then
            If Not TryRestoreCell(Target, Saved) Then
                Notify "The dialog was cancelled, but the cell could not get its content back.", _
                    vbExclamation
            End If
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Save the error before any other call
        ErrNumber = Err.Number
        ErrDescription = Err.Description
    'Give the cell back its content, then report both outcomes
        Restored = True
        If Written Then Restored = TryRestoreCell(Target, Saved)
        If Restored Then
            Notify "Could not insert " & Chosen & ": error " & VBA.CStr(ErrNumber) & " (" & _
                ErrDescription & "). The cell has its content back.", vbExclamation
        Else
            Notify "Could not insert " & Chosen & ": error " & VBA.CStr(ErrNumber) & " (" & _
                ErrDescription & "). The cell could not get its content back.", vbCritical
        End If

End Sub

Public Sub RegisterHelp()

'
'------------------------------------------------------------------------------
'                        REGISTER THE HELP OF THE FUNCTIONS
'------------------------------------------------------------------------------
' PURPOSE
'   Registers the 22 functions in the Insert Function dialog and says how many Excel accepted.
'
' WHY THIS EXISTS
'   M_Register returns a count and shows nothing; the menu tells the user the outcome.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls M_Register.RegisterFunctions and reports its count.
'
' ERROR POLICY
'   Does not raise: M_Register counts a refused function instead of raising.
'
' DEPENDENCIES
'   M_Register, ReportHelp
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' REGISTER
'------------------------------------------------------------------------------
    'Register every function and report the count
        ReportHelp M_Register.RegisterFunctions(), "registered in the category ""KPR Dates"""

End Sub

Public Sub ClearHelp()

'
'------------------------------------------------------------------------------
'                         CLEAR THE HELP OF THE FUNCTIONS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the 22 functions to the User Defined category with no description, and says how
'   many Excel accepted. The functions stay callable.
'
' WHY THIS EXISTS
'   M_Register returns a count and shows nothing; the menu tells the user the outcome.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls M_Register.UnregisterFunctions and reports its count.
'
' ERROR POLICY
'   Does not raise: M_Register counts a refused function instead of raising.
'
' DEPENDENCIES
'   M_Register, ReportHelp
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
    'Clear every function and report the count
        ReportHelp M_Register.UnregisterFunctions(), "moved back to ""User Defined"""

End Sub

Public Sub RunTests()

'
'------------------------------------------------------------------------------
'                           RUN THE PROJECT SUITES
'------------------------------------------------------------------------------
' PURPOSE
'   Development build only: runs the project's test suites and shows their results.
'
' WHY THIS EXISTS
'   A developer checks a change in Excel without typing a runner in the Immediate window.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs TEST_Suites.RunProjectSuites through Application.Run and shows its report: one line
'   per suite, and where the logs are. Each suite restores what it changes and checks it.
'
' ERROR POLICY
'   Does not raise. A build without the test modules, and an error that escapes the runner,
'   are reported in a message box.
'
' DEPENDENCIES
'   TestsAvailable, Qualified, Notify; TEST_Suites through Application.Run
'
' NOTES
'   This is not an evidence run: the evidence of a commit comes from
'   tools/devbuild/excel_run.py, as INSTALLATION.md describes
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_Actions.RunTests"

    Dim Report          As String       'What the runner returned
    Dim ErrNumber       As Long         'Number of the error that escaped the runner
    Dim ErrDescription  As String       'Its description

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Stop when the test modules are not part of this build
        If Not TestsAvailable() Then
            Notify "The test modules are not part of this build.", vbExclamation
            Exit Sub
        End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Report a failure instead of raising it to the menu
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Run the suites and show the report
        Report = VBA.CStr(Application.Run(Qualified(TESTS_RUNNER)))
        If VBA.StrComp(VBA.Left$(Report, VBA.Len(TESTS_PASSED)), TESTS_PASSED, _
            vbBinaryCompare) = 0 Then
            Notify Report, vbInformation
        Else
            Notify Report, vbExclamation
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Save the error before any other call, then report it
        ErrNumber = Err.Number
        ErrDescription = Err.Description
        Notify "The test run stopped: error " & VBA.CStr(ErrNumber) & " (" & ErrDescription & _
            ").", vbExclamation

End Sub

Public Sub RunAction(ByVal Action As String)

'
'------------------------------------------------------------------------------
'                              RUN A MENU ACTION
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the action an action text names. The Ribbon and the toolbar call it with the text
'   stored on the clicked control.
'
' WHY THIS EXISTS
'   Both surfaces reach the actions through one dispatch, so that they cannot differ.
'
' INPUTS
'   Action
'     An action text: ACTION_FUNCTION and a function name, or ACTION_REGISTER, ACTION_CLEAR
'     or ACTION_TESTS.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Selects the action by its text. Each action checks its own argument again.
'
' ERROR POLICY
'   Does not raise. An unknown text is reported in a message box, and nothing runs.
'
' DEPENDENCIES
'   InsertFunction, RegisterHelp, ClearHelp, RunTests, Notify
'
' NOTES
'   The texts are compared exactly, with their case
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DISPATCH
'------------------------------------------------------------------------------
    'Run the action the text names
        If IsFunctionAction(Action) Then
            InsertFunction VBA.Mid$(Action, VBA.Len(ACTION_FUNCTION) + 1)
        ElseIf VBA.StrComp(Action, ACTION_REGISTER, vbBinaryCompare) = 0 Then
            RegisterHelp
        ElseIf VBA.StrComp(Action, ACTION_CLEAR, vbBinaryCompare) = 0 Then
            ClearHelp
        ElseIf VBA.StrComp(Action, ACTION_TESTS, vbBinaryCompare) = 0 Then
            RunTests
        Else
            Notify "Unknown K-PRICING action: " & Action, vbExclamation
        End If

End Sub

'
'------------------------------------------------------------------------------
'
'                                 MENU CONTENT
'
'------------------------------------------------------------------------------
'

Public Function FunctionNames() As Variant

'
'------------------------------------------------------------------------------
'                            LIST THE FUNCTION NAMES
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the names of the 22 functions, in registration order, for a Functions menu.
'
' WHY THIS EXISTS
'   M_Register holds the one list of the functions; both menus read it through here.
'
' INPUTS
'   None
'
' RETURNS
'   Variant
'     A zero-based array of M_Register.FUNCTION_COUNT names.
'
' BEHAVIOR
'   Reads each name from M_Register.RegisteredName.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   M_Register
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Names()         As String       'The names
    Dim Index           As Long         'Position in the registration list

'------------------------------------------------------------------------------
' LIST
'------------------------------------------------------------------------------
    'Read every name in registration order
        ReDim Names(0 To M_Register.FUNCTION_COUNT - 1)
        For Index = 1 To M_Register.FUNCTION_COUNT
            Names(Index - 1) = M_Register.RegisteredName(Index)
        Next Index
        FunctionNames = Names

End Function

Public Function ActionKnown(ByVal Action As String) As Boolean

'
'------------------------------------------------------------------------------
'                           RECOGNIZE AN ACTION TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   True when an action text names an action this build can run, with an argument it
'   accepts. Runs nothing, so that a test can check every menu item without a dialog.
'
' WHY THIS EXISTS
'   A menu item whose text RunAction does not know would only show an error when clicked.
'
' INPUTS
'   Action
'     An action text.
'
' RETURNS
'   Boolean
'     True for ACTION_FUNCTION and one of the 22 names, ACTION_REGISTER, ACTION_CLEAR, and
'     ACTION_TESTS when the test modules are part of this build; False otherwise.
'
' BEHAVIOR
'   Compares the text as RunAction does.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   IsFunctionAction, TestsAvailable
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLASSIFY
'------------------------------------------------------------------------------
    'Accept only what RunAction runs
        If IsFunctionAction(Action) Then
            ActionKnown = True
        ElseIf VBA.StrComp(Action, ACTION_REGISTER, vbBinaryCompare) = 0 Then
            ActionKnown = True
        ElseIf VBA.StrComp(Action, ACTION_CLEAR, vbBinaryCompare) = 0 Then
            ActionKnown = True
        ElseIf VBA.StrComp(Action, ACTION_TESTS, vbBinaryCompare) = 0 Then
            ActionKnown = TestsAvailable()
        Else
            ActionKnown = False
        End If

End Function

Public Function TestsAvailable() As Boolean

'
'------------------------------------------------------------------------------
'                        FIND THE TEST MODULES IN THE BUILD
'------------------------------------------------------------------------------
' PURPOSE
'   True when the test modules are part of this build.
'
' WHY THIS EXISTS
'   The Development menus show only in a development build.
'
' INPUTS
'   None
'
' RETURNS
'   Boolean
'     True when TEST_Suites.SuitesLoaded answers True; False when it is absent or fails.
'
' BEHAVIOR
'   Calls TEST_Suites.SuitesLoaded through Application.Run.
'
' ERROR POLICY
'   Does not raise. The error Application.Run raises for an absent procedure is False.
'
' DEPENDENCIES
'   Qualified; TEST_Suites through Application.Run
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'An absent probe raises in Application.Run; that is False
        On Error GoTo Absent

'------------------------------------------------------------------------------
' ASK
'------------------------------------------------------------------------------
    'Ask the test modules whether they are loaded
        TestsAvailable = VBA.CBool(Application.Run(Qualified(TESTS_PROBE)))

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Absent:
    'The test modules are not part of this build
        Err.Clear
        TestsAvailable = False

End Function

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function IsFunctionAction(ByVal Action As String) As Boolean

'
'------------------------------------------------------------------------------
'                        RECOGNIZE A FUNCTION ACTION
'------------------------------------------------------------------------------
' PURPOSE
'   True when an action text is ACTION_FUNCTION followed by one of the 22 names.
'
' WHY THIS EXISTS
'   RunAction and ActionKnown read the text the same way.
'
' INPUTS
'   Action
'     An action text.
'
' RETURNS
'   Boolean
'     True for a known function; False otherwise.
'
' BEHAVIOR
'   Compares the prefix exactly, then looks the rest up in the registration list.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   FunctionIndex
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLASSIFY
'------------------------------------------------------------------------------
    'Require the prefix, then a known name
        If VBA.StrComp(VBA.Left$(Action, VBA.Len(ACTION_FUNCTION)), ACTION_FUNCTION, _
            vbBinaryCompare) <> 0 Then Exit Function
        IsFunctionAction = (FunctionIndex(VBA.Mid$(Action, VBA.Len(ACTION_FUNCTION) + 1)) > 0)

End Function

Private Function FunctionIndex(ByVal Chosen As String) As Long

'
'------------------------------------------------------------------------------
'                           FIND A FUNCTION BY NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the position of a function in the registration list.
'
' WHY THIS EXISTS
'   Only the 22 registered functions may be inserted.
'
' INPUTS
'   Chosen
'     A function name; the case does not matter, as in Excel.
'
' RETURNS
'   Long
'     The position, 1 to M_Register.FUNCTION_COUNT; 0 for any other name, the empty one
'     included.
'
' BEHAVIOR
'   Compares the name with each registered name.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   M_Register
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Index           As Long         'Position in the registration list

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
    'Return the first position whose name matches
        If VBA.Len(Chosen) = 0 Then Exit Function
        For Index = 1 To M_Register.FUNCTION_COUNT
            If VBA.StrComp(M_Register.RegisteredName(Index), Chosen, vbTextCompare) = 0 Then
                FunctionIndex = Index
                Exit Function
            End If
        Next Index

End Function

Private Function FormulaQualifier(ByVal Book As Workbook) As String

'
'------------------------------------------------------------------------------
'                         QUALIFY A NAME IN A FORMULA
'------------------------------------------------------------------------------
' PURPOSE
'   Returns what a formula puts before a function name of this workbook.
'
' WHY THIS EXISTS
'   A formula in another workbook reaches a function of this one only through its name.
'
' INPUTS
'   Book
'     The workbook of the cell that receives the formula.
'
' RETURNS
'   String
'     An empty text for a cell of this workbook, or when this workbook is an add-in;
'     otherwise this workbook's name, quoted, and an exclamation mark.
'
' BEHAVIOR
'   None
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   ThisWorkbook
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' QUALIFY
'------------------------------------------------------------------------------
    'Qualify only a cell of another workbook
        If ThisWorkbook.IsAddin Or Book Is ThisWorkbook Then
            FormulaQualifier = vbNullString
        Else
            FormulaQualifier = "'" & VBA.Replace(ThisWorkbook.Name, "'", "''") & "'!"
        End If

End Function

Private Function Qualified(ByVal ProcedureName As String) As String

'
'------------------------------------------------------------------------------
'                           QUALIFY A PROCEDURE NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the name Application.Run needs to reach a procedure of this workbook.
'
' WHY THIS EXISTS
'   An unqualified name could reach a procedure of another open workbook; the workbook's name
'   also reaches a module with Option Private Module.
'
' INPUTS
'   ProcedureName
'     A procedure name, such as TEST_Suites.SuitesLoaded.
'
' RETURNS
'   String
'     This workbook's name, quoted, an exclamation mark and the procedure name.
'
' BEHAVIOR
'   None
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   ThisWorkbook
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' QUALIFY
'------------------------------------------------------------------------------
    'Quote the workbook name, doubling any apostrophe in it
        Qualified = "'" & VBA.Replace(ThisWorkbook.Name, "'", "''") & "'!" & ProcedureName

End Function

Private Sub SaveCell(ByVal Cell As Range, ByRef Saved As CellContent)

'
'------------------------------------------------------------------------------
'                           RECORD A CELL'S CONTENT
'------------------------------------------------------------------------------
' PURPOSE
'   Records what a cell holds, so that TryRestoreCell can give it back.
'
' WHY THIS EXISTS
'   A cancelled dialog or a failed write must not lose the content of the cell.
'
' INPUTS
'   Cell
'     The cell.
'
'   Saved (ByRef, output)
'     Receives the formula or the constant, and the prefix character.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Reads a formula through Formula2 where Excel has it, else through Formula.
'
' ERROR POLICY
'   Raises what Excel raises when the cell cannot be read; the caller's handler reports it.
'
' DEPENDENCIES
'   Range.HasFormula, Range.Value2, Range.PrefixCharacter, Range.Formula and Formula2
'
' NOTES
'   Formula2 is read late bound, so that the module compiles in an Excel without it
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Probe           As Object       'The cell, late bound, for Formula2
    Dim ErrNumber       As Long         'Error of the Formula2 read; 0 for none

'------------------------------------------------------------------------------
' RECORD
'------------------------------------------------------------------------------
    'Record a constant, or the formula through Formula2 where Excel has it
        Saved.HadFormula = Cell.HasFormula
        Saved.Value = Cell.Value2
        Saved.Prefix = VBA.CStr(Cell.PrefixCharacter)
        Saved.Formula = vbNullString
        Saved.Dynamic = False
        If Not Saved.HadFormula Then Exit Sub
        Set Probe = Cell
        On Error Resume Next
        Saved.Formula = VBA.CStr(Probe.Formula2)
        ErrNumber = Err.Number
        On Error GoTo 0
        Saved.Dynamic = (ErrNumber = 0)
        If Not Saved.Dynamic Then Saved.Formula = VBA.CStr(Cell.Formula)
        Set Probe = Nothing

End Sub

Private Sub WriteFormula(ByVal Cell As Range, ByVal Formula As String)

'
'------------------------------------------------------------------------------
'                             WRITE A FORMULA
'------------------------------------------------------------------------------
' PURPOSE
'   Writes a formula into a cell, through Formula2 where Excel has it.
'
' WHY THIS EXISTS
'   Formula applies implicit intersection, so a result of several cells would not spill.
'
' INPUTS
'   Cell
'     The cell.
'
'   Formula
'     The formula, starting with "=".
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Tries Formula2 late bound; when Excel has no Formula2, writes Formula.
'
' ERROR POLICY
'   Raises what Excel raises when Formula refuses the text; the caller's handler reports it.
'
' DEPENDENCIES
'   Range.Formula and Formula2
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Probe           As Object       'The cell, late bound, for Formula2
    Dim ErrNumber       As Long         'Error of the Formula2 write; 0 for none

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Try Formula2, then fall back to Formula
        Set Probe = Cell
        On Error Resume Next
        Probe.Formula2 = Formula
        ErrNumber = Err.Number
        On Error GoTo 0
        Set Probe = Nothing
        If ErrNumber <> 0 Then Cell.Formula = Formula

End Sub

Private Function TryRestoreCell(ByVal Cell As Range, ByRef Saved As CellContent) As Boolean

'
'------------------------------------------------------------------------------
'                          GIVE A CELL ITS CONTENT BACK
'------------------------------------------------------------------------------
' PURPOSE
'   Writes back what SaveCell recorded, and says whether it succeeded.
'
' WHY THIS EXISTS
'   The cleanup of InsertFunction must report its own failure without hiding the error that
'   made it necessary.
'
' INPUTS
'   Cell
'     The cell.
'
'   Saved
'     What SaveCell recorded; not changed.
'
' RETURNS
'   Boolean
'     True when the cell holds its recorded content again; False when Excel refused a write.
'
' BEHAVIOR
'   Writes back the formula, through Formula2 when it was read that way; clears an empty
'   cell; writes back a constant, with its apostrophe for text entered as text.
'
' ERROR POLICY
'   Does not raise. A refused write is False, and Err is cleared.
'
' DEPENDENCIES
'   Range.Formula and Formula2, Range.Value2, Range.ClearContents
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Probe           As Object       'The cell, late bound, for Formula2

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'A refused write is a failed restoration
        On Error GoTo Refused

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Write back the formula or the constant the cell held
        If Saved.HadFormula Then
            If Saved.Dynamic Then
                Set Probe = Cell
                Probe.Formula2 = Saved.Formula
                Set Probe = Nothing
            Else
                Cell.Formula = Saved.Formula
            End If
        ElseIf IsEmpty(Saved.Value) Then
            Cell.ClearContents
        ElseIf VarType(Saved.Value) = vbString And Saved.Prefix = "'" Then
            Cell.Value2 = "'" & Saved.Value
        Else
            Cell.Value2 = Saved.Value
        End If
        TryRestoreCell = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Refused:
    'Excel refused the write
        Err.Clear
        Set Probe = Nothing
        TryRestoreCell = False

End Function

Private Sub ReportHelp(ByVal Accepted As Long, ByVal Outcome As String)

'
'------------------------------------------------------------------------------
'                        REPORT A REGISTRATION COUNT
'------------------------------------------------------------------------------
' PURPOSE
'   Tells the user how many of the 22 functions Excel accepted.
'
' WHY THIS EXISTS
'   RegisterHelp and ClearHelp report their counts the same way.
'
' INPUTS
'   Accepted
'     The count M_Register returned.
'
'   Outcome
'     What happened to the accepted functions, such as "registered in the category ...".
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   An information message when Excel accepted every function; a warning otherwise.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   M_Register.FUNCTION_COUNT, Notify
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Say how many functions Excel accepted
        If Accepted = M_Register.FUNCTION_COUNT Then
            Notify VBA.CStr(Accepted) & " K-PRICING functions " & Outcome & ".", vbInformation
        Else
            Notify VBA.CStr(Accepted) & " of " & VBA.CStr(M_Register.FUNCTION_COUNT) & _
                " K-PRICING functions " & Outcome & "; Excel refused the others.", vbExclamation
        End If

End Sub

Private Sub Notify(ByVal Message As String, ByVal Style As VbMsgBoxStyle)

'
'------------------------------------------------------------------------------
'                              TELL THE USER
'------------------------------------------------------------------------------
' PURPOSE
'   Shows one message box with the menus' title.
'
' WHY THIS EXISTS
'   Every action tells the user its outcome the same way.
'
' INPUTS
'   Message
'     The text.
'
'   Style
'     The icon: vbInformation, vbExclamation or vbCritical.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   None
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   MsgBox
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' SHOW
'------------------------------------------------------------------------------
    'Show the message with the menus' title
        MsgBox Message, Style, UI_TITLE

End Sub
