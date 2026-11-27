Attribute VB_Name = "TEST_Ui"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: TEST_Ui
'------------------------------------------------------------------------------
' PURPOSE
'   Tests of the menus of the workbook: the list of functions both menus show, the action
'   texts M_Actions runs, the cells InsertFunction refuses, the Functions menu of the Ribbon,
'   the toolbar's lifecycle and the checks of the Dates Demo sheet
'
' WHY THIS EXISTS
'   The menus run only in Excel. A menu item whose action is unknown, a Ribbon menu that does
'   not list the 22 functions and a toolbar that a second install doubles would show only when
'   a user clicks them
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   RunUiTests runs the suite. Option Private Module keeps it out of the workbook's external
'   surface. The cases open no dialog and show no message: they call the procedures of the
'   menus that do neither. The toolbar case removes and installs the toolbar, and gives it
'   back when it was installed before the run. CertifyUi certifies its normal run, as
'   .github/excel-evidence-policy.json lists it
'
' ERROR POLICY
'   As TEST_Harness.EndRun: a run that does not pass raises TEST_Harness.ERR_TEST_FAILED,
'   or the unexpected error that escaped the runner. A refused start raises
'   TEST_Harness.ERR_TEST_DIRTY_START
'
' DEPENDENCIES
'   M_Actions, M_Ribbon, M_CommandBars, M_Register, TEST_WorksheetSupport, TEST_Harness; the
'   Excel object model and the Office CommandBars objects
'
' NOTES
'   In a development workbook holding the modules of src/ and tests/modules/, type
'   TEST_Ui.RunUiTests in the Immediate window. The Ribbon's callbacks and a click on a menu
'   item need Excel's user interface: the certification run checks them by hand (#20)
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'A complete run starts every case and evaluates every assertion once
    Private Const SUITE_NAME            As String = "ui"            'Suite name in the run log
    Private Const EXPECTED_CASES        As Long = 6                 'Cases a complete run starts
    Private Const EXPECTED_ASSERTIONS   As Long = 36                'Assertions a complete run evaluates

    'The Office constant of a popup, written out because the objects are late bound
    Private Const MSO_CONTROL_POPUP     As Long = 10                'msoControlPopup

    'The functions both menus list, and the buttons of a development toolbar
    Private Const FUNCTIONS             As Long = 22                'Functions of the menus
    Private Const TOOLBAR_BUTTONS       As Long = 25                '22 functions, 2 setup, 1 test

    'The checks of the Dates Demo sheet: its statuses, and the one that differs by design with
    'the sheet's own inputs (AddMonths with Keep EOM against EDATE)
    Private Const DEMO_ANCHOR           As String = "DatesDemo_Checks"  'First header cell
    Private Const DEMO_CHECKS           As Long = 57                    'Statuses on the sheet
    Private Const DEMO_DIFFERS          As Long = 1                     'Differences by design

    'The demo's input names, separated by ";", in the order of their cells
    Private Const DEMO_INPUTS           As String = "DateIn;nDays;nWeeks;nMonths;nYears;KeepEOM;WdIndex;nOccur;WeekBaseMonday"


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub RunUiTests()

'
'------------------------------------------------------------------------------
'                                 RUN UI TESTS
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case in a fixed order between BeginRun and EndRun. A passing run ends with
'   RESULT=PASS; suite=ui; operation=PASS; completeness=COMPLETE; cases=6; assertions=36;
'   failures=0; cleanup=PASS; recovery=NONE
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
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Ui.RunUiTests"

    Dim ErrNumber       As Long     'Runner error number; 0 for none
    Dim ErrSource       As String   'Runner error source
    Dim ErrDescription  As String   'Runner error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the run; a refused start propagates directly
        TEST_Harness.BeginRun SUITE_NAME, "NORMAL", EXPECTED_CASES, EXPECTED_ASSERTIONS
    'Send an error that escapes a case to the runner handler
        On Error GoTo RunFailed

'------------------------------------------------------------------------------
' RUN CASES
'------------------------------------------------------------------------------
    'Run every case in a fixed order
        CaseFunctionList
        CaseActions
        CaseInsertRefusal
        CaseRibbonMenu
        CaseToolbar
        CaseDemoSheet

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

Public Sub CertifyUi( _
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
'   Runs RunUiTests, saves its log, certifies the run, then saves the certification log.
'
' ERROR POLICY
'   TEST_Harness.EndCertification raises when the certification did not pass.
'
' DEPENDENCIES
'   RunUiTests, TEST_Harness
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
    Const PROC_NAME             As String = "TEST_Ui.CertifyUi"

    Dim RunErrorNumber          As Long     'Error that ended the run; 0 for none
    Dim RunErrorDescription     As String   'Its description

'------------------------------------------------------------------------------
' CERTIFY
'------------------------------------------------------------------------------
    'Refuse state left by an earlier run and start the certification
        TEST_Harness.BeginCertification SUITE_NAME
    'Run the suite and capture the error that ended it
        On Error Resume Next
        RunUiTests
        RunErrorNumber = Err.Number
        RunErrorDescription = Err.Description
    'Restore normal error handling
        On Error GoTo 0
    'Save the log and certify the run
        TEST_Harness.SaveLog Folder & "\ui.log"
        TEST_Harness.CertifyRun "NORMAL", True, RunErrorNumber, RunErrorDescription
    'Save the certification log, raising when the certification did not pass
        TEST_Harness.EndCertification Folder & "\ui-certification.log"

End Sub

'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseFunctionList()

'
'------------------------------------------------------------------------------
'                         CHECK THE LIST OF FUNCTIONS
'------------------------------------------------------------------------------
' PURPOSE
'   The menus list the 22 functions M_Register registers, in its order. 2 assertions.
'
' WHY THIS EXISTS
'   Both menus read M_Actions.FunctionNames; a function missing there is missing from both
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Compares the list with M_Register.RegisteredName, position by position
'
' ERROR POLICY
'   An unexpected error is recorded against the case
'
' DEPENDENCIES
'   M_Actions.FunctionNames, M_Register.RegisteredName, TEST_Harness
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
    Const PROC_NAME     As String = "TEST_Ui.CaseFunctionList"

    Dim Names           As Variant      'M_Actions.FunctionNames
    Dim Index           As Long         'Position in the list
    Dim Differences     As String       'Positions whose name differs

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "function-list"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'The list holds the 22 functions, in registration order
        Names = M_Actions.FunctionNames()
        TEST_Harness.AssertEqualDouble "count", FUNCTIONS, UBound(Names) - LBound(Names) + 1
        For Index = LBound(Names) To UBound(Names)
            If Names(Index) <> M_Register.RegisteredName(Index - LBound(Names) + 1) Then
                Differences = Differences & " " & VBA.CStr(Index)
            End If
        Next Index
        TEST_Harness.AssertEqualText "registration-order", vbNullString, Differences

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

Private Sub CaseActions()

'
'------------------------------------------------------------------------------
'                            CHECK THE ACTION TEXTS
'------------------------------------------------------------------------------
' PURPOSE
'   M_Actions knows the action of every function and of every fixed menu item, and refuses
'   malformed texts. 5 assertions.
'
' WHY THIS EXISTS
'   RunAction runs what ActionKnown accepts; an unknown text would show an error on a click
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Checks every function action, the case of a name, the setup and test actions, then a list
'   of texts RunAction must refuse
'
' ERROR POLICY
'   An unexpected error is recorded against the case
'
' DEPENDENCIES
'   M_Actions, TEST_Harness
'
' NOTES
'   This suite runs in a development build, where the test action is known
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Ui.CaseActions"

    Dim Item            As Variant      'Cursor over names and texts
    Dim Unknown         As String       'Function actions refused
    Dim Accepted        As String       'Malformed texts accepted

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "actions"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Every function has a known action; a name is matched in any case, as Excel does
        For Each Item In M_Actions.FunctionNames()
            If Not M_Actions.ActionKnown(M_Actions.ACTION_FUNCTION & VBA.CStr(Item)) Then
                Unknown = Unknown & " " & VBA.CStr(Item)
            End If
        Next Item
        TEST_Harness.AssertEqualText "function-actions", vbNullString, Unknown
        TEST_Harness.AssertTrue "name-case", M_Actions.ActionKnown("fn:kpr_dates_endofmonth"), _
            "a function name in lower case was refused"
    'The fixed items are known; the test action needs the test modules, which are loaded
        TEST_Harness.AssertTrue "setup-actions", M_Actions.ActionKnown(M_Actions.ACTION_REGISTER) _
            And M_Actions.ActionKnown(M_Actions.ACTION_CLEAR), "a setup action was refused"
        TEST_Harness.AssertTrue "test-action", M_Actions.ActionKnown(M_Actions.ACTION_TESTS), _
            "the test action was refused although the test modules are loaded"
    'Malformed and unknown texts are refused, the prefix and the fixed texts in their case
        For Each Item In VBA.Array("", "fn", "fn:", ":x", "fn:KPR_Dates_NoSuch", _
            "FN:KPR_Dates_EndOfMonth", "fn: KPR_Dates_EndOfMonth", "setup:delete", _
            "Setup:register", "setup:register ", "dev:all", "dev:evidence", "other:x")
            If M_Actions.ActionKnown(VBA.CStr(Item)) Then
                Accepted = Accepted & " '" & VBA.CStr(Item) & "'"
            End If
        Next Item
        TEST_Harness.AssertEqualText "refused", vbNullString, Accepted

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

Private Sub CaseInsertRefusal()

'
'------------------------------------------------------------------------------
'                   CHECK THE CELLS A FUNCTION CANNOT GO IN
'------------------------------------------------------------------------------
' PURPOSE
'   InsertFunction refuses an unknown name, no cell and a cell of a table, accepts a plain
'   cell, and writes nothing while it decides. 6 assertions.
'
' WHY THIS EXISTS
'   InsertFunction restores one cell only: in an empty column of a table, Excel would copy the
'   temporary formula to every row, and a cancelled dialog would leave them filled
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Builds a table of two rows with an empty column on a scratch workbook, asks
'   M_Actions.InsertRefusal about its cells and a cell outside it, then closes the workbook
'
' ERROR POLICY
'   An unexpected error is recorded against the case, once the scratch workbook is closed
'
' DEPENDENCIES
'   M_Actions.InsertRefusal, TEST_WorksheetSupport, TEST_Harness
'
' NOTES
'   InsertRefusal shows no message, so the case needs no dialog
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Ui.CaseInsertRefusal"

    Dim Sheet           As Worksheet    'The scratch sheet
    Dim Nowhere         As Range        'No cell
    Dim Detail          As String       'What closing the workbook could not restore
    Dim ErrNumber       As Long         'Error of the case; 0 for none
    Dim ErrSource       As String       'Its source
    Dim ErrDescription  As String       'Its description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "insert-refusal"
    'Close the scratch workbook on every path
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'A table of two rows whose second column is empty
        Set Sheet = TEST_WorksheetSupport.OpenScratch(False)
        Sheet.Range("A1").Value2 = "Date"
        Sheet.Range("B1").Value2 = "Result"
        Sheet.Range("A2").Value2 = 46022
        Sheet.Range("A3").Value2 = 46053
        Sheet.ListObjects.Add SourceType:=xlSrcRange, Source:=Sheet.Range("A1:B3"), _
            XlListObjectHasHeaders:=xlYes
    'A cell of the table is refused, a cell outside it is not
        TEST_Harness.AssertTrue "table-cell", _
            VBA.Len(M_Actions.InsertRefusal("KPR_Dates_EndOfMonth", Sheet.Range("B2"))) > 0, _
            "a cell of an empty table column was accepted"
        TEST_Harness.AssertEqualText "plain-cell", vbNullString, _
            M_Actions.InsertRefusal("KPR_Dates_EndOfMonth", Sheet.Range("D5"))
    'An unknown name and no cell are refused
        TEST_Harness.AssertTrue "unknown-name", _
            VBA.Len(M_Actions.InsertRefusal("KPR_Dates_NoSuch", Sheet.Range("D5"))) > 0, _
            "an unknown function was accepted"
        TEST_Harness.AssertTrue "no-cell", _
            VBA.Len(M_Actions.InsertRefusal("KPR_Dates_EndOfMonth", Nowhere)) > 0, _
            "no cell was accepted"
    'Deciding wrote nothing
        TEST_Harness.AssertEqualDouble "nothing-written", 4, _
            Application.WorksheetFunction.CountA(Sheet.UsedRange)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
CleanUp:
    'Restore normal error handling
        On Error GoTo 0
    'Close the scratch workbook and restore the calculation mode
        Detail = TEST_WorksheetSupport.CloseScratch()
        TEST_Harness.AssertTrue "scratch-closed", VBA.Len(Detail) = 0, Detail
    'Record the case's error, once Excel is restored
        If ErrNumber <> 0 Then
            Err.Number = ErrNumber
            Err.Source = ErrSource
            Err.Description = ErrDescription
            TEST_Harness.RecordCaseError
        End If
    'Exit before the error handlers
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Keep the case's error and close the workbook
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
        Resume CleanUp

CaseFailed:
    'Record the unexpected error against the case
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseRibbonMenu()

'
'------------------------------------------------------------------------------
'                         CHECK THE RIBBON'S MENU XML
'------------------------------------------------------------------------------
' PURPOSE
'   The Functions menu of the Ribbon is a customUI menu with one button per function, each
'   carrying the function's action and the module's onAction callback. 4 assertions.
'
' WHY THIS EXISTS
'   Excel rejects a menu it cannot read and shows it empty, with no error
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Builds the menu with M_Ribbon.BuildFunctionMenu and checks its root, its number of
'   buttons and callbacks, and the tag of every function
'
' ERROR POLICY
'   An unexpected error is recorded against the case
'
' DEPENDENCIES
'   M_Ribbon.BuildFunctionMenu, M_Actions, TEST_Harness
'
' NOTES
'   tools/check_ribbon.py checks the static XML; this case checks what VBA builds
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Ui.CaseRibbonMenu"

    Dim Xml             As String       'The menu XML
    Dim Item            As Variant      'Function names cursor
    Dim Missing         As String       'Functions without their button

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "ribbon-menu"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'The menu is a customUI menu element
        M_Ribbon.BuildFunctionMenu Xml
        TEST_Harness.AssertTrue "root", VBA.InStr(1, Xml, _
            "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">", _
            vbBinaryCompare) = 1 And VBA.Right$(Xml, 7) = "</menu>", "the menu root is " & _
            VBA.Left$(Xml, 80)
    'One button per function, each with the module's callback
        TEST_Harness.AssertEqualDouble "buttons", FUNCTIONS, Occurrences(Xml, "<button ")
        TEST_Harness.AssertEqualDouble "callbacks", FUNCTIONS, _
            Occurrences(Xml, " onAction=""RibbonOnAction""/>")
    'Each function's button carries its action text
        For Each Item In M_Actions.FunctionNames()
            If Occurrences(Xml, " tag=""" & M_Actions.ACTION_FUNCTION & VBA.CStr(Item) & _
                """ ") <> 1 Then Missing = Missing & " " & VBA.CStr(Item)
        Next Item
        TEST_Harness.AssertEqualText "tags", vbNullString, Missing

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

Private Sub CaseToolbar()

'
'------------------------------------------------------------------------------
'                        CHECK THE TOOLBAR'S LIFECYCLE
'------------------------------------------------------------------------------
' PURPOSE
'   The toolbar installs with one known action per button, a second install adds nothing,
'   and removal, repeated, leaves nothing; none of them moves the user. 14 assertions.
'
' WHY THIS EXISTS
'   A toolbar that a second install doubles, or that leaves a control behind, stays in the
'   user's Excel until it closes
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Removes the toolbar, installs it twice and removes it twice, counting what it owns after
'   each step, then gives it back when it was installed before the case
'
' ERROR POLICY
'   An unexpected error is recorded against the case; the toolbar is then given back as it
'   was, when that is possible
'
' DEPENDENCIES
'   M_CommandBars, M_Actions, TEST_Harness; Application.CommandBars
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
    Const PROC_NAME     As String = "TEST_Ui.CaseToolbar"

    Dim HadToolbar      As Boolean      'The toolbar was installed before the case
    Dim Done            As Boolean      'Outcome of an install or a removal
    Dim Report          As String       'Why an install failed
    Dim First           As Long         'What the first install owns
    Dim Actions         As Collection   'Action texts of the toolbar's buttons
    Dim Item            As Variant      'Actions cursor
    Dim Unknown         As String       'Actions M_Actions does not know
    Dim BookName        As String       'Active workbook before the case
    Dim SheetName       As String       'Active sheet before the case
    Dim Selected        As String       'Selection before the case

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "toolbar"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Keep what the case must give back
        HadToolbar = (OwnedCount() > 0)
        BookName = ActiveWorkbook.Name
        SheetName = ActiveSheet.Name
        Selected = SelectionText()
    'Start with no toolbar
        M_CommandBars.TryRemoveToolbar Done
        TEST_Harness.AssertTrue "clean-start", Done And OwnedCount() = 0, _
            VBA.CStr(OwnedCount()) & " owned bar(s) or control(s) left"
    'The first install builds a visible toolbar whose every button has a known action
        M_CommandBars.TryInstallToolbar Done, Report
        TEST_Harness.AssertTrue "install", Done, Report
        TEST_Harness.AssertTrue "visible", ToolbarVisible(), "the toolbar is hidden or absent"
        First = OwnedCount()
        Set Actions = ToolbarActions()
        TEST_Harness.AssertEqualDouble "buttons", TOOLBAR_BUTTONS, Actions.Count
        For Each Item In Actions
            If Not M_Actions.ActionKnown(VBA.CStr(Item)) Then
                Unknown = Unknown & " '" & VBA.CStr(Item) & "'"
            End If
        Next Item
        TEST_Harness.AssertEqualText "actions-known", vbNullString, Unknown
    'A second install replaces the first and adds nothing
        M_CommandBars.TryInstallToolbar Done, Report
        TEST_Harness.AssertTrue "repeat-install", Done, Report
        TEST_Harness.AssertEqualDouble "repeat-adds-nothing", First, OwnedCount()
    'Removal leaves nothing, and a second removal is not an error
        M_CommandBars.TryRemoveToolbar Done
        TEST_Harness.AssertTrue "remove", Done, "TryRemoveToolbar reported a control left"
        TEST_Harness.AssertEqualDouble "nothing-left", 0, OwnedCount()
        M_CommandBars.TryRemoveToolbar Done
        TEST_Harness.AssertTrue "repeat-remove", Done And OwnedCount() = 0, _
            "the second removal failed or left something"
    'No step moved the user
        TEST_Harness.AssertEqualText "active-workbook", BookName, ActiveWorkbook.Name
        TEST_Harness.AssertEqualText "active-sheet", SheetName, ActiveSheet.Name
        TEST_Harness.AssertEqualText "selection", Selected, SelectionText()
    'Give back a toolbar the user had installed
        If HadToolbar Then M_CommandBars.TryInstallToolbar Done, Report
        TEST_Harness.AssertTrue "given-back", (OwnedCount() > 0) = HadToolbar, _
            "the toolbar was not left as the case found it"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the unexpected error against the case, then give the toolbar back
        TEST_Harness.RecordCaseError
        On Error Resume Next
        If HadToolbar Then
            M_CommandBars.TryInstallToolbar Done, Report
        Else
            M_CommandBars.TryRemoveToolbar Done
        End If
        Err.Clear
        On Error GoTo 0

End Sub

Private Sub CaseDemoSheet()

'
'------------------------------------------------------------------------------
'                     CHECK THE DATES DEMO SHEET IN EXCEL
'------------------------------------------------------------------------------
' PURPOSE
'   The Dates Demo sheet, calculated with the inputs it is generated with, counts every status
'   and none fails: each function gives what Excel's own formula gives, but the one difference
'   by design. The inputs a user typed are given back. 5 assertions.
'
' WHY THIS EXISTS
'   The template holds the demo's formulas and no result: only Excel calculates them, so only
'   this case shows that the demo works as a user sees it
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Keeps the content of the inputs DEMO_INPUTS, formulas included, writes the generated ones, calculates the sheet
'   of the anchor DEMO_ANCHOR and reads the four counts below it: OK, Differs by design, Fail
'   and Checks. Then writes the kept values back, on every path
'
' ERROR POLICY
'   An unexpected error is recorded against the case, once the inputs are written back
'
' DEPENDENCIES
'   The defined names DEMO_ANCHOR and DEMO_INPUTS of this workbook, ReadContent,
'   RestoreInputs, TEST_Harness
'
' NOTES
'   tests/test_gen_workbook_template.py keeps DEMO_CHECKS, DEMO_INPUTS and the generated values
'   equal to tools/gen_workbook_template.py and tools/demo_dates.py. The spill and the array
'   limits need an Excel with dynamic arrays
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Ui.CaseDemoSheet"

    Dim Anchor          As Range        'The first header cell of the demo's checks
    Dim Names           As Variant      'The input names
    Dim Generated       As Variant      'The values the sheet is generated with
    Dim Saved()         As Variant      'The content on entry: a formula or a value
    Dim Formulas()      As Boolean      'Which of them are formulas
    Dim Kept            As Long         'How many of them were kept
    Dim Index           As Long         'Inputs cursor
    Dim Restored        As Boolean      'Every kept value was written back
    Dim ErrNumber       As Long         'Error of the case; 0 for none
    Dim ErrSource       As String       'Its source
    Dim ErrDescription  As String       'Its description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "demo-sheet"
    'Write the inputs back on every path
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Keep the inputs, then write the values the counts belong to; DateIn is 2024-04-30
        Names = VBA.Split(DEMO_INPUTS, ";")
        Generated = VBA.Array(45412, 2, 3, 1, 2, True, 4, 2, True)
        ReDim Saved(LBound(Names) To UBound(Names))
        ReDim Formulas(LBound(Names) To UBound(Names))
        For Index = LBound(Names) To UBound(Names)
            Formulas(Index) = ReadContent(ThisWorkbook.Names(Names(Index)).RefersToRange, _
                Saved(Index))
            Kept = Kept + 1
        Next Index
        For Index = LBound(Names) To UBound(Names)
            ThisWorkbook.Names(Names(Index)).RefersToRange.Value2 = Generated(Index)
        Next Index
    'Calculate the demo
        Set Anchor = ThisWorkbook.Names(DEMO_ANCHOR).RefersToRange
        Anchor.Worksheet.Calculate
    'Every status is counted, none fails, and only the documented one differs
        TEST_Harness.AssertEqualDouble "checks", DEMO_CHECKS, Anchor.Offset(4, 1).Value2
        TEST_Harness.AssertEqualDouble "fails", 0, Anchor.Offset(3, 1).Value2
        TEST_Harness.AssertEqualDouble "differs", DEMO_DIFFERS, Anchor.Offset(2, 1).Value2
        TEST_Harness.AssertEqualDouble "ok", DEMO_CHECKS - DEMO_DIFFERS, Anchor.Offset(1, 1).Value2
        Set Anchor = Nothing

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
CleanUp:
    'Restore normal error handling
        On Error GoTo 0
    'Write the kept inputs back
        Restored = RestoreInputs(Names, Saved, Formulas, Kept)
        TEST_Harness.AssertTrue "inputs-restored", Restored, "an input of the demo was not written back"
    'Record the case's error, once the inputs are back
        If ErrNumber <> 0 Then
            Err.Number = ErrNumber
            Err.Source = ErrSource
            Err.Description = ErrDescription
            TEST_Harness.RecordCaseError
        End If
    'Exit before the error handlers
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Keep the case's error and write the inputs back
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
        Set Anchor = Nothing
        Resume CleanUp

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

Private Function RestoreInputs( _
    ByVal Names As Variant, _
    ByRef Saved() As Variant, _
    ByRef Formulas() As Boolean, _
    ByVal Kept As Long) As Boolean
'
'------------------------------------------------------------------------------
'                         WRITE THE DEMO'S INPUTS BACK
'------------------------------------------------------------------------------
' PURPOSE
'   Writes back the first Kept inputs that CaseDemoSheet kept, formulas as formulas, and says
'   whether every one now holds its kept content.
'
' WHY THIS EXISTS
'   The case changes the cells a user may have edited; it must give them back on every path,
'   a formula such as =TODAY() included, and report a failure to do so without hiding the
'   case's own error
'
' INPUTS
'   Names
'     The input names, or Empty when none was read.
'
'   Saved
'     The kept content, in the order of Names: a formula's text or a value; not changed.
'
'   Formulas
'     Which kept contents are formulas; not changed.
'
'   Kept
'     How many inputs were kept, from the first.
'
' RETURNS
'   Boolean
'     True when every kept content is back; True when none was kept.
'
' BEHAVIOR
'   Writes each kept content with WriteContent, then reads it again with ReadContent and
'   compares: the formula's text, or the value and its type
'
' ERROR POLICY
'   Does not raise. A write that fails is False, and Err is cleared.
'
' DEPENDENCIES
'   ReadContent, WriteContent; the defined names of this workbook
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
    Const PROC_NAME     As String = "TEST_Ui.RestoreInputs"

    Dim Index           As Long         'Inputs cursor
    Dim Cell            As Range        'An input cell
    Dim Content         As Variant      'What the cell holds after the write
    Dim IsFormula       As Boolean      'The cell holds a formula after the write

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'A write that fails is a failed restoration
        RestoreInputs = True
        If Kept = 0 Then Exit Function
        On Error GoTo Refused

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Write each kept content, then read it again and compare
        For Index = LBound(Names) To LBound(Names) + Kept - 1
            Set Cell = ThisWorkbook.Names(Names(Index)).RefersToRange
            WriteContent Cell, Saved(Index), Formulas(Index)
            IsFormula = ReadContent(Cell, Content)
            If IsFormula <> Formulas(Index) Then
                RestoreInputs = False
            ElseIf VBA.VarType(Content) <> VBA.VarType(Saved(Index)) Then
                RestoreInputs = False
            ElseIf Not IsEmpty(Saved(Index)) And Not IsError(Saved(Index)) Then
                If Content <> Saved(Index) Then RestoreInputs = False
            End If
        Next Index
        Set Cell = Nothing

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Refused:
    'Excel refused a write
        Err.Clear
        Set Cell = Nothing
        RestoreInputs = False

End Function

Private Function ReadContent(ByVal Cell As Range, ByRef Content As Variant) As Boolean

'
'------------------------------------------------------------------------------
'                            READ WHAT A CELL HOLDS
'------------------------------------------------------------------------------
' PURPOSE
'   Reads a cell's formula, or its value when it holds no formula, and says which it read.
'
' WHY THIS EXISTS
'   An input may hold a formula, which its value alone would lose
'
' INPUTS
'   Cell
'     The cell.
'
'   Content (ByRef, output)
'     Receives the formula's text, through Formula2 where Excel has it, else the Value2.
'
' RETURNS
'   Boolean
'     True when the cell holds a formula.
'
' BEHAVIOR
'   Reads Formula2 late bound, so that the module compiles in an Excel without it
'
' ERROR POLICY
'   Raises what Excel raises when the cell cannot be read; the caller handles it.
'
' DEPENDENCIES
'   Range.HasFormula, Range.Formula and Formula2, Range.Value2
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
    Const PROC_NAME     As String = "TEST_Ui.ReadContent"

    Dim Probe           As Object       'The cell, late bound, for Formula2
    Dim ErrNumber       As Long         'Error of the Formula2 read; 0 for none

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'A constant is read as its value
        If Not Cell.HasFormula Then
            Content = Cell.Value2
            Exit Function
        End If
    'A formula is read through Formula2 where Excel has it, else through Formula
        ReadContent = True
        Set Probe = Cell
        On Error Resume Next
        Content = VBA.CStr(Probe.Formula2)
        ErrNumber = Err.Number
        On Error GoTo 0
        Set Probe = Nothing
        If ErrNumber <> 0 Then Content = VBA.CStr(Cell.Formula)

End Function

Private Sub WriteContent(ByVal Cell As Range, ByVal Content As Variant, ByVal IsFormula As Boolean)

'
'------------------------------------------------------------------------------
'                            WRITE WHAT A CELL HELD
'------------------------------------------------------------------------------
' PURPOSE
'   Writes back what ReadContent read: a formula as a formula, a value as a value.
'
' WHY THIS EXISTS
'   RestoreInputs gives an input back as the user left it
'
' INPUTS
'   Cell
'     The cell.
'
'   Content
'     A formula's text or a value.
'
'   IsFormula
'     True when Content is a formula.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes a formula through Formula2 late bound where Excel has it, else through Formula
'
' ERROR POLICY
'   Raises what Excel raises when it refuses the write; RestoreInputs handles it.
'
' DEPENDENCIES
'   Range.Formula and Formula2, Range.Value2
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
    Const PROC_NAME     As String = "TEST_Ui.WriteContent"

    Dim Probe           As Object       'The cell, late bound, for Formula2
    Dim ErrNumber       As Long         'Error of the Formula2 write; 0 for none

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'A value is written as a value
        If Not IsFormula Then
            Cell.Value2 = Content
            Exit Sub
        End If
    'A formula is written through Formula2 where Excel has it, else through Formula
        Set Probe = Cell
        On Error Resume Next
        Probe.Formula2 = Content
        ErrNumber = Err.Number
        On Error GoTo 0
        Set Probe = Nothing
        If ErrNumber <> 0 Then Cell.Formula = Content

End Sub

Private Function Occurrences(ByVal Text As String, ByVal Part As String) As Long

'
'------------------------------------------------------------------------------
'                            COUNT A PART OF A TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns how many times a part occurs in a text, without overlap.
'
' WHY THIS EXISTS
'   The menu case counts the elements and attributes of the menu XML
'
' INPUTS
'   Text
'     The text searched.
'
'   Part
'     The part counted; not empty.
'
' RETURNS
'   Long
'     The number of occurrences, compared case by case.
'
' BEHAVIOR
'   None
'
' ERROR POLICY
'   Does not raise
'
' DEPENDENCIES
'   None
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
    Const PROC_NAME     As String = "TEST_Ui.Occurrences"

    Dim Position        As Long         'Where the next search starts

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
    'Count each occurrence and search after it
        Position = VBA.InStr(1, Text, Part, vbBinaryCompare)
        Do While Position > 0
            Occurrences = Occurrences + 1
            Position = VBA.InStr(Position + VBA.Len(Part), Text, Part, vbBinaryCompare)
        Loop

End Function

Private Function OwnedCount() As Long

'
'------------------------------------------------------------------------------
'                         COUNT WHAT THE TOOLBAR OWNS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of bars named, and of controls tagged, as the toolbar's.
'
' WHY THIS EXISTS
'   The toolbar case reads what an install or a removal left without trusting the procedure
'   it checks
'
' INPUTS
'   None
'
' RETURNS
'   Long
'     The count.
'
' BEHAVIOR
'   Counts the bars by M_CommandBars.TOOLBAR_NAME, then the controls by
'   M_CommandBars.TOOLBAR_TAG
'
' ERROR POLICY
'   Raises what Office raises; the case records it
'
' DEPENDENCIES
'   Application.CommandBars, M_CommandBars
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
    Const PROC_NAME     As String = "TEST_Ui.OwnedCount"

    Dim Bar             As Object       'CommandBars cursor
    Dim Found           As Object       'FindControls result, Nothing when none

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
    'Count by name, then by tag
        For Each Bar In Application.CommandBars
            If VBA.StrComp(Bar.Name, M_CommandBars.TOOLBAR_NAME, vbTextCompare) = 0 Then
                OwnedCount = OwnedCount + 1
            End If
        Next Bar
        Set Found = Application.CommandBars.FindControls(Tag:=M_CommandBars.TOOLBAR_TAG)
        If Not Found Is Nothing Then OwnedCount = OwnedCount + Found.Count
        Set Found = Nothing

End Function

Private Function ToolbarVisible() As Boolean

'
'------------------------------------------------------------------------------
'                        READ WHETHER THE TOOLBAR SHOWS
'------------------------------------------------------------------------------
' PURPOSE
'   True when the toolbar exists and is visible.
'
' WHY THIS EXISTS
'   CommandBars raises for a name it does not hold
'
' INPUTS
'   None
'
' RETURNS
'   Boolean
'     True when the bar exists and is visible; False otherwise.
'
' BEHAVIOR
'   None
'
' ERROR POLICY
'   Does not raise
'
' DEPENDENCIES
'   Application.CommandBars, M_CommandBars.TOOLBAR_NAME
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'An absent bar raises; that is False
        On Error Resume Next
        ToolbarVisible = Application.CommandBars(M_CommandBars.TOOLBAR_NAME).Visible
        Err.Clear
        On Error GoTo 0

End Function

Private Function ToolbarActions() As Collection

'
'------------------------------------------------------------------------------
'                      COLLECT THE ACTIONS OF THE TOOLBAR
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Parameter of every button of the toolbar, menus searched in order.
'
' WHY THIS EXISTS
'   The toolbar case checks that every button runs an action M_Actions knows
'
' INPUTS
'   None
'
' RETURNS
'   Collection
'     The action texts; empty when the toolbar is absent.
'
' BEHAVIOR
'   Reads the bar's menus and their buttons
'
' ERROR POLICY
'   Raises what Office raises; the case records it
'
' DEPENDENCIES
'   Application.CommandBars, M_CommandBars.TOOLBAR_NAME
'
' NOTES
'   The toolbar's menus hold buttons only, one level deep
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Ui.ToolbarActions"

    Dim Found           As Collection   'The action texts
    Dim Bar             As Object       'The toolbar
    Dim Menu            As Object       'Menu cursor
    Dim Button          As Object       'Button cursor

'------------------------------------------------------------------------------
' COLLECT
'------------------------------------------------------------------------------
    'Find the bar; an absent one has no action
        Set Found = New Collection
        Set ToolbarActions = Found
        On Error Resume Next
        Set Bar = Application.CommandBars(M_CommandBars.TOOLBAR_NAME)
        Err.Clear
        On Error GoTo 0
        If Bar Is Nothing Then Exit Function
    'Read each button of each menu
        For Each Menu In Bar.Controls
            If Menu.Type = MSO_CONTROL_POPUP Then
                For Each Button In Menu.Controls
                    Found.Add VBA.CStr(Button.Parameter)
                Next Button
            End If
        Next Menu
        Set Bar = Nothing

End Function

Private Function SelectionText() As String

'
'------------------------------------------------------------------------------
'                            DESCRIBE THE SELECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the selection's external address, or its type when it is not a Range.
'
' WHY THIS EXISTS
'   The toolbar case checks that no step moved the selection
'
' INPUTS
'   None
'
' RETURNS
'   String
'     The address, or the type name; an empty text when there is no selection.
'
' BEHAVIOR
'   None
'
' ERROR POLICY
'   Does not raise
'
' DEPENDENCIES
'   Selection
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Read the selection, which may be missing
        On Error Resume Next
        If TypeName(Selection) = "Range" Then
            SelectionText = Selection.Address(External:=True)
        Else
            SelectionText = TypeName(Selection)
        End If
        Err.Clear
        On Error GoTo 0

End Function
