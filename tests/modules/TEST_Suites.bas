Attribute VB_Name = "TEST_Suites"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: TEST_Suites
'------------------------------------------------------------------------------
' PURPOSE
'   Runs every project suite of .github/excel-evidence-policy.json in a row, for the Run tests
'   button of the menus, and tells the menus that the test modules are loaded
'
' WHY THIS EXISTS
'   M_Actions is production code: it reaches the test modules only by name, through
'   Application.Run, so a production build compiles without them. This module is the one name
'   it needs
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - SuitesLoaded      True; Application.Run raises when this module is absent
'     - RunProjectSuites  run each suite's runner and return one line per suite
'   Option Private Module keeps it out of the workbook's external surface
'
' ERROR POLICY
'   RunProjectSuites does not raise: a suite that fails, or that refuses to start, is a FAIL
'   line of its report, and the next suite still runs
'
' DEPENDENCIES
'   The runner of each project suite, through Application.Run; TEST_Harness.LastResult
'
' NOTES
'   - SUITES lists the policy's project suites in its order, each with its runner;
'     tests/test_check_excel_run.py keeps them equal. The example suite is left out: its
'     runners need the fixtures folder
'   - This is not an evidence run: tools/devbuild/excel_run.py certifies each suite and saves
'     its evidence, as INSTALLATION.md describes
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'Each project suite and its runner, "<suite>=<module>.<runner>", separated by ";"
    Private Const SUITES As String = _
        "core-err-parse=TEST_CoreErrParse.RunCoreErrParseTests;" & _
        "core-dates=TEST_CoreDates.RunCoreDatesTests;" & _
        "core-array=TEST_CoreArray.RunCoreArrayTests;" & _
        "ranges=TEST_Ranges.RunRangeTests;" & _
        "dates=TEST_Dates.RunDateTests;" & _
        "date-cases=TEST_DateCases.RunDateCaseTests;" & _
        "worksheet-cases=TEST_WorksheetCases.RunWorksheetCaseTests;" & _
        "excel-host=TEST_Host.RunHostTests;" & _
        "excel-oracle=TEST_ExcelOracle.RunExcelOracleTests;" & _
        "ui=TEST_Ui.RunUiTests"

    'The start of a passing RESULT record of the harness
    Private Const PASSED                As String = "RESULT=PASS"   'Passing record


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Function SuitesLoaded() As Boolean

'
'------------------------------------------------------------------------------
'                        SAY THAT THE TESTS ARE LOADED
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True, so that M_Actions can tell a development build from a production one
'
' WHY THIS EXISTS
'   Application.Run raises when the procedure it names is absent: an answer means the test
'   modules are loaded
'
' INPUTS
'   None
'
' RETURNS
'   Boolean
'     True
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
' RETURN VALUE
'------------------------------------------------------------------------------
    'Answer that the test modules are loaded
        SuitesLoaded = True

End Function

Public Function RunProjectSuites() As String

'
'------------------------------------------------------------------------------
'                            RUN THE PROJECT SUITES
'------------------------------------------------------------------------------
' PURPOSE
'   Runs every project suite in the policy's order and returns a report: RESULT=PASS or
'   RESULT=FAIL, a summary, then one line per suite
'
' WHY THIS EXISTS
'   The Run tests button shows one result per suite, and a failing suite must not stop the
'   others
'
' INPUTS
'   None
'
' RETURNS
'   String
'     The report. Its first line is RESULT=PASS only when every suite passed; a suite's line
'     is its name and PASS, or FAIL and the RESULT record or error that ended it
'
' BEHAVIOR
'   Calls each runner through Application.Run with this workbook's name, then reads the
'   suite's RESULT record from TEST_Harness.LastResult. Each suite's log goes to the
'   Immediate window, as the harness writes it
'
' ERROR POLICY
'   Does not raise. An error that ends a suite is part of its FAIL line
'
' DEPENDENCIES
'   TEST_Harness.LastResult; each runner through Application.Run
'
' NOTES
'   A suite that leaves recovery required makes the next ones refuse to start; their lines
'   say so, and TEST_Harness.ResetTests clears it after a check by hand
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Suites.RunProjectSuites"

    Dim Entries         As Variant      'The SUITES entries
    Dim Index           As Long         'Entries cursor
    Dim Suite           As String       'Name of the suite
    Dim Runner          As String       'Its runner, "<module>.<runner>"
    Dim Outcome         As String       'Its RESULT record, or the error that ended it
    Dim ErrNumber       As Long         'Error that ended the runner; 0 for none
    Dim ErrDescription  As String       'Its description
    Dim Lines           As String       'One line per suite
    Dim Failed          As Long         'Suites that did not pass

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Run each suite, keeping its outcome; a failure does not stop the next one
        Entries = VBA.Split(SUITES, ";")
        For Index = LBound(Entries) To UBound(Entries)
            Suite = VBA.Split(Entries(Index), "=")(0)
            Runner = VBA.Split(Entries(Index), "=")(1)
            On Error Resume Next
            Application.Run "'" & VBA.Replace(ThisWorkbook.Name, "'", "''") & "'!" & Runner
            ErrNumber = Err.Number
            ErrDescription = Err.Description
            On Error GoTo 0
            Outcome = TEST_Harness.LastResult()
            If ErrNumber = 0 And VBA.InStr(1, Outcome, PASSED & ";", vbBinaryCompare) = 1 Then
                Lines = Lines & vbCrLf & Suite & ": PASS"
            Else
                Failed = Failed + 1
                If ErrNumber <> 0 Then Outcome = Outcome & " (error " & VBA.CStr(ErrNumber) & _
                    ": " & ErrDescription & ")"
                Lines = Lines & vbCrLf & Suite & ": FAIL " & Outcome
            End If
        Next Index

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'The result first, then the summary and the suites
        If Failed = 0 Then
            RunProjectSuites = PASSED & vbCrLf & "Every one of the " & _
                VBA.CStr(UBound(Entries) - LBound(Entries) + 1) & " project suites passed."
        Else
            RunProjectSuites = "RESULT=FAIL" & vbCrLf & VBA.CStr(Failed) & " of the " & _
                VBA.CStr(UBound(Entries) - LBound(Entries) + 1) & " project suites did not pass."
        End If
        RunProjectSuites = RunProjectSuites & vbCrLf & Lines & vbCrLf & vbCrLf & _
            "The log of each suite is in the Immediate window of the VBA editor."

End Function
