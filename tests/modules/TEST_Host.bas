Attribute VB_Name = "TEST_Host"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: TEST_Host
'------------------------------------------------------------------------------
' PURPOSE
'   Tests of the date layer in its Excel host: a worksheet caller whose workbook changes date
'   system, a multi-cell result spilled by a dynamic-array formula, the read of a block of a
'   workbook's cells, and the registration of the 22 worksheet functions
'
' WHY THIS EXISTS
'   The date system, the spill and MacroOptions exist only in Excel: no call from VBA and no
'   static check reaches them. These cases check them where a user meets them
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   RunHostTests runs the suite. Option Private Module keeps it out of the workbook's external
'   surface. The worksheet cases open scratch workbooks through TEST_WorksheetSupport, close
'   them on every path and assert that they did; the registration case leaves the functions
'   registered. CertifyHost certifies its normal run, as .github/excel-evidence-policy.json
'   lists it
'
' ERROR POLICY
'   As TEST_Harness.EndRun: a run that does not pass raises TEST_Harness.ERR_TEST_FAILED,
'   or the unexpected error that escaped the runner. A refused start raises
'   TEST_Harness.ERR_TEST_DIRTY_START
'
' DEPENDENCIES
'   KPR_Dates, M_Ranges, M_Register, CORE_Array, CORE_Err, TEST_WorksheetSupport,
'   TEST_DateCaseSupport, TEST_Harness; the Excel object model
'
' NOTES
'   In a development workbook holding the modules of src/ and tests/modules/, type
'   TEST_Host.RunHostTests in the Immediate window. The spill case needs dynamic-array Excel:
'   on another host its assertions fail, and say so. Expected dates are those of the 1900
'   date system, read through TEST_WorksheetSupport.EncodeCell
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'A complete run starts every case and evaluates every assertion once
    Private Const SUITE_NAME            As String = "excel-host"        'Suite name in the run log
    Private Const EXPECTED_CASES        As Long = 4                 'Cases a complete run starts
    Private Const EXPECTED_ASSERTIONS   As Long = 41                'Assertions a complete run evaluates


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub RunHostTests()

'
'------------------------------------------------------------------------------
'                               RUN HOST TESTS
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case in a fixed order between BeginRun and EndRun. A passing run ends with
'   RESULT=PASS; suite=excel-host; operation=PASS; completeness=COMPLETE; cases=4;
'   assertions=41; failures=0; cleanup=PASS; recovery=NONE
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
'   2026-11-26
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
        CaseDateSystemToggle
        CaseSpill
        CaseRangeBlock
        CaseRegistration

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

Public Sub CertifyHost( _
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
'   Runs RunHostTests, saves its log, certifies the run, then saves the certification log.
'
' ERROR POLICY
'   TEST_Harness.EndCertification raises when the certification did not pass.
'
' DEPENDENCIES
'   RunHostTests, TEST_Harness
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
        RunHostTests
        RunErrorNumber = Err.Number
        RunErrorDescription = Err.Description
    'Restore normal error handling
        On Error GoTo 0
    'Save the log and certify the run
        TEST_Harness.SaveLog Folder & "\excel-host.log"
        TEST_Harness.CertifyRun "NORMAL", True, RunErrorNumber, RunErrorDescription
    'Save the certification log, raising when the certification did not pass
        TEST_Harness.EndCertification Folder & "\excel-host-certification.log"

End Sub

'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseDateSystemToggle()

'
'------------------------------------------------------------------------------
'                        CHECK A CHANGE OF DATE SYSTEM
'------------------------------------------------------------------------------
' PURPOSE
'   One workbook goes from the 1900 to the 1904 date system and back: the diagnostic follows
'   each change on an ordinary recalculation, a value function proceeds under 1900 and is
'   refused under 1904, and a propagated #N/A stays #N/A. 10 assertions.
'
' WHY THIS EXISTS
'   KPR_Dates_HostDateSystem is volatile so that a workbook whose date system changed reports
'   it without a full rebuild (CALC-10); the refusal under 1904 is the host rule (CALC-01)
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Enters the diagnostic once, and the two value formulas afresh under each system, then
'   calculates with Application.Calculate and reads the three cells
'
' ERROR POLICY
'   An unexpected error is kept, the workbook closed, and the error recorded against the case
'   by TEST_Harness.RecordCaseError; it does not stop the run
'
' DEPENDENCIES
'   TEST_WorksheetSupport, TEST_Harness
'
' NOTES
'   A full rebuild would hide a diagnostic that does not recalculate, so none is asked for
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet           As Worksheet    'The scratch sheet
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
        TEST_Harness.BeginCase "date-system-toggle"
    'Close the scratch workbook on every path
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'A 1900 workbook, with the diagnostic entered once
        Set Sheet = TEST_WorksheetSupport.OpenScratch(False)
        Sheet.Range("A1").Formula = "=" & TEST_WorksheetSupport.Qualified("KPR_Dates_HostDateSystem") & "()"
        EnterValueFormulas Sheet
        TEST_Harness.AssertEqualText "diagnostic-1900", "lng:1900", _
            TEST_WorksheetSupport.EncodeCell(Sheet.Range("A1").Value2, "lng")
        TEST_Harness.AssertEqualText "value-1900", "date:2026-03-01", _
            TEST_WorksheetSupport.EncodeCell(Sheet.Range("A2").Value2, "date")
        TEST_Harness.AssertEqualText "propagated-na-1900", "err:#N/A", _
            TEST_WorksheetSupport.EncodeCell(Sheet.Range("A3").Value2, "err")
    'The same workbook in the 1904 date system
        Sheet.Parent.Date1904 = True
        EnterValueFormulas Sheet
        TEST_Harness.AssertEqualText "diagnostic-1904", "lng:1904", _
            TEST_WorksheetSupport.EncodeCell(Sheet.Range("A1").Value2, "lng")
        TEST_Harness.AssertEqualText "value-1904", "err:#N/A", _
            TEST_WorksheetSupport.EncodeCell(Sheet.Range("A2").Value2, "err")
        TEST_Harness.AssertEqualText "propagated-na-1904", "err:#N/A", _
            TEST_WorksheetSupport.EncodeCell(Sheet.Range("A3").Value2, "err")
    'Back to the 1900 date system
        Sheet.Parent.Date1904 = False
        EnterValueFormulas Sheet
        TEST_Harness.AssertEqualText "diagnostic-1900-again", "lng:1900", _
            TEST_WorksheetSupport.EncodeCell(Sheet.Range("A1").Value2, "lng")
        TEST_Harness.AssertEqualText "value-1900-again", "date:2026-03-01", _
            TEST_WorksheetSupport.EncodeCell(Sheet.Range("A2").Value2, "date")
        TEST_Harness.AssertEqualText "propagated-na-1900-again", "err:#N/A", _
            TEST_WorksheetSupport.EncodeCell(Sheet.Range("A3").Value2, "err")

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
CleanUp:
    'Restore normal error handling
        On Error GoTo 0
    'Close the scratch workbook and restore the calculation mode
        Detail = TEST_WorksheetSupport.CloseScratch()
        TEST_Harness.AssertTrue "scratch-closed", Len(Detail) = 0, Detail
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

Private Sub CaseSpill()

'
'------------------------------------------------------------------------------
'                         CHECK A SPILLED RESULT
'------------------------------------------------------------------------------
' PURPOSE
'   A dynamic-array formula over a 3x1 Range spills three cells under 1900, with a blank row's
'   own #VALUE! between two dates, and under 1904 returns one #N/A and spills nothing.
'   9 assertions.
'
' WHY THIS EXISTS
'   A multi-cell result reaches a user only as a spill; an element's error must not hide its
'   neighbours, and a refused call must not write beyond its cell
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes two dates as text around a blank cell, enters the formula with Formula2 and reads
'   the spilled cells, then changes the date system and enters it again
'
' ERROR POLICY
'   An unexpected error is kept, the workbook closed, and the error recorded against the case
'   by TEST_Harness.RecordCaseError; it does not stop the run
'
' DEPENDENCIES
'   EnterSpillFormula, SpillRows, TEST_WorksheetSupport, TEST_Harness
'
' NOTES
'   Formula2 and SpillingToRange are reached through an Object reference, so the module
'   compiles on a host without them; there the first assertion fails and names the cause
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet           As Worksheet    'The scratch sheet
    Dim Entered         As Boolean      'Whether Formula2 accepted the formula
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
        TEST_Harness.BeginCase "spill"
    'Close the scratch workbook on every path
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Two dates as text around a blank cell, in a 1900 workbook
        Set Sheet = TEST_WorksheetSupport.OpenScratch(False)
        Sheet.Range("A1:A3").NumberFormat = "@"
        Sheet.Range("A1").Value = "2026-03-15"
        Sheet.Range("A3").Value = "2026-04-09"
    'The formula spills one result per row
        Entered = EnterSpillFormula(Sheet.Range("C1"))
        TEST_Harness.AssertTrue "formula2-available", Entered, "this host has no dynamic-array formulas"
        TEST_Harness.AssertEqualDouble "spill-rows-1900", 3, SpillRows(Sheet.Range("C1"))
        TEST_Harness.AssertEqualText "row-1-1900", "date:2026-03-31", _
            TEST_WorksheetSupport.EncodeCell(Sheet.Range("C1").Value2, "date")
        TEST_Harness.AssertEqualText "row-2-1900", "err:#VALUE!", _
            TEST_WorksheetSupport.EncodeCell(Sheet.Range("C2").Value2, "err")
        TEST_Harness.AssertEqualText "row-3-1900", "date:2026-04-30", _
            TEST_WorksheetSupport.EncodeCell(Sheet.Range("C3").Value2, "date")
    'Under 1904 the call is refused as a whole and writes one cell
        Sheet.Parent.Date1904 = True
        EnterSpillFormula Sheet.Range("C1")
        TEST_Harness.AssertEqualText "call-1904", "err:#N/A", _
            TEST_WorksheetSupport.EncodeCell(Sheet.Range("C1").Value2, "err")
        TEST_Harness.AssertTrue "no-spill-1904", SpillRows(Sheet.Range("C1")) <= 1, "the refused call spilled"
        TEST_Harness.AssertEqualText "neighbour-1904", "empty", _
            TEST_WorksheetSupport.EncodeCell(Sheet.Range("C2").Value2, "empty")

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
CleanUp:
    'Restore normal error handling
        On Error GoTo 0
    'Close the scratch workbook and restore the calculation mode
        Detail = TEST_WorksheetSupport.CloseScratch()
        TEST_Harness.AssertTrue "scratch-closed", Len(Detail) = 0, Detail
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

Private Sub CaseRangeBlock()

'
'------------------------------------------------------------------------------
'                        CHECK A BLOCK OF CELLS READ
'------------------------------------------------------------------------------
' PURPOSE
'   A 3x3 block of a workbook's cells is read with its blank, its error and its values in
'   place; a column and a Range beyond the used range keep their shapes; reading moves no
'   selection. 9 assertions.
'
' WHY THIS EXISTS
'   M_Ranges reads a worksheet caller's Ranges; TEST_Ranges reads the development workbook's
'   first sheet, which holds no error value
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Fills the block, calculates its #N/A, reads it and three of its elements, then the column
'   and the 50-row Range, and compares the selection with the one before
'
' ERROR POLICY
'   An unexpected error is kept, the workbook closed, and the error recorded against the case
'   by TEST_Harness.RecordCaseError; it does not stop the run
'
' DEPENDENCIES
'   ReadSize, SelectionText, M_Ranges.TryReadArgument, CORE_Array.ElementAt,
'   TEST_WorksheetSupport, TEST_DateCaseSupport, TEST_Harness
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
    Dim Sheet           As Worksheet    'The scratch sheet
    Dim Selected        As String       'The selection before the reads
    Dim Payload         As Variant      'The block read
    Dim Kind            As KPRArgShape  'Its kind
    Dim RowCount        As Long         'Its rows
    Dim ColCount        As Long         'Its columns
    Dim Condition       As KPRCondition 'The condition reported
    Dim Accepted        As Boolean      'Whether the block was read
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
        TEST_Harness.BeginCase "range-block"
    'Close the scratch workbook on every path
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'A block with a blank at B2 and an error at C3
        Set Sheet = TEST_WorksheetSupport.OpenScratch(False)
        Sheet.Range("A1:C1").Value = Array(1, 2, 3)
        Sheet.Range("A2").Value = 4
        Sheet.Range("C2").Value = 6
        Sheet.Range("A3:B3").Value = Array(7, 8)
        Sheet.Range("C3").Formula = "=NA()"
        Application.Calculate
        Selected = SelectionText()
    'The block keeps its shape and each element its value
        Accepted = M_Ranges.TryReadArgument(Sheet.Range("A1:C3"), Payload, Kind, RowCount, ColCount, Condition)
        TEST_Harness.AssertTrue "block-accepted", Accepted, CORE_Err.ConditionName(Condition)
        TEST_Harness.AssertEqualText "block-size", "3x3", ReadSize(Sheet.Range("A1:C3"))
        If Not Accepted Then Err.Raise vbObjectError + 1, "TEST_Host.CaseRangeBlock", "the block was refused"
        TEST_Harness.AssertEqualText "blank-element", "empty", _
            TEST_DateCaseSupport.Encode(CORE_Array.ElementAt(Payload, Kind, 2, 2))
        TEST_Harness.AssertEqualText "error-element", "err:#N/A", _
            TEST_DateCaseSupport.Encode(CORE_Array.ElementAt(Payload, Kind, 3, 3))
        TEST_Harness.AssertEqualText "value-element", "num:8", _
            TEST_DateCaseSupport.Encode(CORE_Array.ElementAt(Payload, Kind, 3, 2))
    'A column and a Range beyond the used range keep their shapes
        TEST_Harness.AssertEqualText "column-size", "3x1", ReadSize(Sheet.Range("A1:A3"))
        TEST_Harness.AssertEqualText "beyond-used-range-size", "50x1", ReadSize(Sheet.Range("A1:A50"))
    'Reading moved no selection
        TEST_Harness.AssertEqualText "selection-unchanged", Selected, SelectionText()

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
CleanUp:
    'Restore normal error handling
        On Error GoTo 0
    'Close the scratch workbook and restore the calculation mode
        Detail = TEST_WorksheetSupport.CloseScratch()
        TEST_Harness.AssertTrue "scratch-closed", Len(Detail) = 0, Detail
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

Private Sub CaseRegistration()

'
'------------------------------------------------------------------------------
'                        CHECK THE FUNCTION REGISTRATION
'------------------------------------------------------------------------------
' PURPOSE
'   Registering twice, unregistering twice and registering again each reach the 22 functions,
'   and leave the workbook's saved state, the application settings, the active workbook and
'   sheet and the selection as they were. 13 assertions.
'
' WHY THIS EXISTS
'   Registration runs when the add-in opens and may run again; each call must be complete and
'   repeatable and must not disturb the user's session
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Captures the state, makes the five calls, asserts each count, then compares the state
'
' ERROR POLICY
'   An unexpected error is recorded against the case by TEST_Harness.RecordCaseError and does
'   not stop the run
'
' DEPENDENCIES
'   M_Register, SelectionText, TEST_Harness
'
' NOTES
'   The last call registers, so the functions stay registered after the run
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WasSaved        As Boolean          'ThisWorkbook.Saved before the calls
    Dim SavedCalculation As XlCalculation   'Calculation mode before the calls
    Dim Events          As Boolean          'EnableEvents before the calls
    Dim Screen          As Boolean          'ScreenUpdating before the calls
    Dim Alerts          As Boolean          'DisplayAlerts before the calls
    Dim BookName        As String           'Active workbook before the calls
    Dim SheetName       As String           'Active sheet before the calls
    Dim Selected        As String           'Selection before the calls

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "registration"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Keep the state the calls must leave unchanged
        WasSaved = ThisWorkbook.Saved
        SavedCalculation = Application.Calculation
        Events = Application.EnableEvents
        Screen = Application.ScreenUpdating
        Alerts = Application.DisplayAlerts
        BookName = ActiveWorkbook.Name
        SheetName = ActiveSheet.Name
        Selected = SelectionText()
    'Each call reaches every function, whatever came before
        TEST_Harness.AssertEqualDouble "register-1", 22, M_Register.RegisterFunctions()
        TEST_Harness.AssertEqualDouble "register-2", 22, M_Register.RegisterFunctions()
        TEST_Harness.AssertEqualDouble "unregister-1", 22, M_Register.UnregisterFunctions()
        TEST_Harness.AssertEqualDouble "unregister-2", 22, M_Register.UnregisterFunctions()
        TEST_Harness.AssertEqualDouble "register-3", 22, M_Register.RegisterFunctions()
    'The session is as it was
        TEST_Harness.AssertTrue "saved-state", ThisWorkbook.Saved = WasSaved, "ThisWorkbook.Saved changed"
        TEST_Harness.AssertTrue "calculation", Application.Calculation = SavedCalculation, "the calculation mode changed"
        TEST_Harness.AssertTrue "events", Application.EnableEvents = Events, "EnableEvents changed"
        TEST_Harness.AssertTrue "screen-updating", Application.ScreenUpdating = Screen, "ScreenUpdating changed"
        TEST_Harness.AssertTrue "alerts", Application.DisplayAlerts = Alerts, "DisplayAlerts changed"
        TEST_Harness.AssertEqualText "active-workbook", BookName, ActiveWorkbook.Name
        TEST_Harness.AssertEqualText "active-sheet", SheetName, ActiveSheet.Name
        TEST_Harness.AssertEqualText "selection", Selected, SelectionText()

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

Private Sub EnterValueFormulas( _
    ByVal Sheet As Worksheet)

'
'------------------------------------------------------------------------------
'                      ENTER THE VALUE FORMULAS AND CALCULATE
'------------------------------------------------------------------------------
' PURPOSE
'   Enters a value function in A2 and a propagated #N/A in A3, then recalculates.
'
' WHY THIS EXISTS
'   The date-system case enters them afresh under each system.
'
' INPUTS
'   Sheet
'     The scratch sheet.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes the two formulas and calls Application.Calculate, an ordinary recalculation.
'
' ERROR POLICY
'   Raises what the Excel object model raises.
'
' DEPENDENCIES
'   TEST_WorksheetSupport.Qualified
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ENTER
'------------------------------------------------------------------------------
    'Enter the two formulas and recalculate
        Sheet.Range("A2").Formula = "=" & TEST_WorksheetSupport.Qualified("KPR_Dates_BeginOfMonth") & _
            "(""2026-03-15"")"
        Sheet.Range("A3").Formula = "=" & TEST_WorksheetSupport.Qualified("KPR_Dates_BeginOfMonth") & "(NA())"
        Application.Calculate

End Sub

Private Function EnterSpillFormula( _
    ByVal Anchor As Range) As Boolean

'
'------------------------------------------------------------------------------
'                         ENTER A DYNAMIC-ARRAY FORMULA
'------------------------------------------------------------------------------
' PURPOSE
'   Enters EndOfMonth over A1:A3 as a dynamic-array formula and recalculates.
'
' WHY THIS EXISTS
'   Only a formula entered with Formula2 spills.
'
' INPUTS
'   Anchor
'     The formula's cell.
'
' RETURNS
'   Boolean
'     True when the host accepted Formula2.
'
' BEHAVIOR
'   Sets Formula2 through an Object reference, then calls Application.Calculate.
'
' ERROR POLICY
'   Does not raise for a host without Formula2: returns False.
'
' DEPENDENCIES
'   TEST_WorksheetSupport.Qualified
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
    Dim Cell            As Object       'The formula's cell, late-bound

'------------------------------------------------------------------------------
' ENTER
'------------------------------------------------------------------------------
    'Set Formula2 through a late-bound reference, since an older host lacks it
        Set Cell = Anchor
        On Error Resume Next
        Cell.Formula2 = "=" & TEST_WorksheetSupport.Qualified("KPR_Dates_EndOfMonth") & "(A1:A3)"
        EnterSpillFormula = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
    'Recalculate
        Application.Calculate

End Function

Private Function SpillRows( _
    ByVal Anchor As Range) As Long

'
'------------------------------------------------------------------------------
'                         COUNT THE ROWS OF A SPILL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of rows the formula in the cell spills to.
'
' WHY THIS EXISTS
'   The spill case compares the spilled range with the result's shape.
'
' INPUTS
'   Anchor
'     The formula's cell.
'
' RETURNS
'   Long
'     The rows of SpillingToRange, or 0 when there is none or the host lacks it.
'
' BEHAVIOR
'   Reads SpillingToRange through an Object reference.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   None
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
    Dim Cell            As Object       'The formula's cell, late-bound
    Dim Spill           As Range        'The spilled range, or Nothing

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
    'Read the spilled range through a late-bound reference, since an older host lacks it
        Set Cell = Anchor
        On Error Resume Next
        Set Spill = Cell.SpillingToRange
        Err.Clear
        On Error GoTo 0
    'Count its rows
        If Not Spill Is Nothing Then SpillRows = Spill.Rows.Count

End Function

Private Function ReadSize( _
    ByVal Block As Range) As String

'
'------------------------------------------------------------------------------
'                           READ A RANGE'S SHAPE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the rows and columns M_Ranges reads from a Range, or the condition it refused it
'   with.
'
' WHY THIS EXISTS
'   The block case compares the shape read with the Range's own.
'
' INPUTS
'   Block
'     The Range.
'
' RETURNS
'   String
'     RxC, or the condition's name.
'
' BEHAVIOR
'   Calls M_Ranges.TryReadArgument.
'
' ERROR POLICY
'   Raises what M_Ranges raises.
'
' DEPENDENCIES
'   M_Ranges.TryReadArgument, CORE_Err.ConditionName
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
    Dim Payload         As Variant      'The values read
    Dim Kind            As KPRArgShape  'Their kind
    Dim RowCount        As Long         'Their rows
    Dim ColCount        As Long         'Their columns
    Dim Condition       As KPRCondition 'The condition reported

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Read the Range and describe what came back
        If M_Ranges.TryReadArgument(Block, Payload, Kind, RowCount, ColCount, Condition) Then
            ReadSize = CStr(RowCount) & "x" & CStr(ColCount)
        Else
            ReadSize = CORE_Err.ConditionName(Condition)
        End If

End Function

Private Function SelectionText() As String

'
'------------------------------------------------------------------------------
'                          DESCRIBE THE SELECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the selection's external address, or its type name when it is not a Range.
'
' WHY THIS EXISTS
'   A case compares the selection before and after the calls it makes.
'
' INPUTS
'   None
'
' RETURNS
'   String
'     The address, the type name, or empty when nothing can be read.
'
' BEHAVIOR
'   Reads Selection.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   The Excel object model
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DESCRIBE
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
