Attribute VB_Name = "KPR_Demo_Dates"
'==============================================================================
' MODULE: KPR_Demo_Dates
'------------------------------------------------------------------------------
' PURPOSE
'   Deterministic builder of the date-layer demonstration workbook: one new
'   workbook, five fixed sheets of live examples, each beside the result the
'   contract documents and a cell that shows whether the two match.
'
' WHY THIS EXISTS
'   A generated workbook is an Office binary and is never committed. This
'   tracked builder is the authoritative source of the demo: rebuilding it from
'   the same source in the same host gives the same sheets, labels, formulas,
'   examples, widths, styles and names. A generated copy may be attached to a
'   release only after exact-source certification (#52).
'
' SCOPE
'   - KPR_Demo_BuildDates       build the demo and save it to an explicit path
'   - KPR_Demo_LastReport       why the last build failed, or "" on success
'
' SHEETS
'   About     purpose, evidence boundary, future scope, the workbook date
'             system and the total number of mismatches
'   Scalar    one example per supported function, plus the Monday/Sunday
'             weekday base and the near-integer count rule
'   Arrays    dynamic-array examples: column, row and 2-D spills, a scalar
'             broadcast, an element-level error, and the 100,000-element
'             capacity policy shown through single-cell formulas
'   Errors    every native-error category the contract documents
'   Pillars   pillar parsing, aliases, signs, rounding modes and formatting
'
' EXPECTED VALUES
'   Every expected value was computed by the independent reference model in
'   tools/gen_fixtures.py from docs/DATE_LAYER_CONTRACT.md, never by running
'   this VBA. Dates are passed as YYYY-MM-DD text, so no example depends on
'   the regional date format.
'
' ONE SUPPORTED SURFACE
'   The 21 value-taking functions keep the same names for scalar and
'   multi-cell calls; KPR_Dates_HostDateSystem is shown as scalar-only. No
'   _Spill function exists or is shown. Multi-cell results are claimed only on
'   dynamic-array Excel, and no Ctrl+Shift+Enter behaviour is claimed.
'
' FORMULAS AND THE HOST
'   The new workbook calls the functions of the VBA project that built it.
'   When that project is an add-in the formulas use the bare function names;
'   otherwise each call is qualified with the host workbook name, as Excel
'   requires for a function in another open workbook. The live results
'   therefore need the host project to be open. The Formula column shows the
'   unqualified formula. Formulas are written with Range.Formula2 where Excel
'   provides it, so multi-cell results spill.
'
' STATE
'   Calculation, events, screen updating and alerts are restored to the
'   values found on entry, and the caller's workbook and sheet are made active
'   again. The demo workbook is closed after it is saved. An existing file is
'   never overwritten: the build stops before creating anything when the path
'   exists, and the workbook is saved under a unique temporary name in the
'   same folder and then renamed with the Name statement, which refuses an
'   existing destination. A file that appears at OutputPath while the demo is
'   being built is therefore kept, and the temporary file is deleted.
'
' ERROR POLICY
'   No entry point raises. A failure closes the unsaved demo workbook,
'   restores the caller's state, returns FALSE and is described by
'   KPR_Demo_LastReport.
'
' VISIBILITY
'   Option Private Module. The entry points are unsupported demo
'   infrastructure, callable from the Immediate window or Application.Run.
'
' ALLOWED DEPENDENCIES
'   No KPR module: function names appear only inside formula text.
'
' UPDATED
'   2026-10-05
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit         'Force explicit variable declarations
    Option Private Module   'Demo infrastructure: invisible outside this project

'------------------------------------------------------------------------------
' CONSTANTS
'------------------------------------------------------------------------------
    'Sheet names, in workbook order
        Private Const SHEET_ABOUT   As String = "About"
        Private Const SHEET_SCALAR  As String = "Scalar"
        Private Const SHEET_ARRAYS  As String = "Arrays"
        Private Const SHEET_ERRORS  As String = "Errors"
        Private Const SHEET_PILLARS As String = "Pillars"

    'First data row of the Scalar, Errors and Pillars tables
        Private Const TABLE_FIRST_ROW As Long = 5

    'Excel's Open XML workbook format, the only format the builder writes
        Private Const XL_OPEN_XML_WORKBOOK As Long = 51

    'Display format of every date the demo shows
        Private Const DATE_FORMAT As String = "yyyy-mm-dd"

'------------------------------------------------------------------------------
' BUILD STATE
'------------------------------------------------------------------------------
    Private mLastReport     As String       'Failure description of the last build
    Private mPrefix         As String       'Host qualifier inserted before KPR_Dates_

'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Function KPR_Demo_BuildDates(ByVal OutputPath As String) As Boolean
'
'==============================================================================
'                             KPR_Demo_BuildDates
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the demonstration workbook and saves it as an .xlsx file at
'   OutputPath.
'
' INPUTS
'   OutputPath
'     Full path of a file that does not exist yet, ending in .xlsx, in a
'     folder that exists.
'
' RETURNS
'   TRUE when the workbook was built, calculated, saved and closed. FALSE
'   otherwise; KPR_Demo_LastReport says why.
'
' BEHAVIOR
'   Refuses an existing path, a missing folder or another extension before
'   creating anything. Restores calculation, events, screen updating, alerts
'   and the caller's active workbook and sheet on every path.
'
' UPDATED
'   2026-10-05
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CallerBook      As Workbook     'Active workbook on entry
    Dim CallerSheet     As Object       'Active sheet on entry
    Dim OldCalculation  As Long         'Calculation mode on entry
    Dim OldEvents       As Boolean      'EnableEvents on entry
    Dim OldScreen       As Boolean      'ScreenUpdating on entry
    Dim OldAlerts       As Boolean      'DisplayAlerts on entry
    Dim Wb              As Workbook     'The demo workbook under construction
    Dim Detail          As String       'Path validation message
    Dim Folder          As String       'Folder of OutputPath
    Dim TempPath        As String       'Unique temporary file the workbook is saved to
    Dim Stage           As String       'Build stage, for the failure report

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Default outcome is failure with an explicit report
        KPR_Demo_BuildDates = False
        mLastReport = ""
    'Nothing is created unless the destination is acceptable
        If Not TryValidatePath(OutputPath, Folder, Detail) Then
            mLastReport = Detail
            Exit Function
        End If

'------------------------------------------------------------------------------
' CAPTURE STATE
'------------------------------------------------------------------------------
        On Error GoTo Fail
        Stage = "capture Excel state"
        Set CallerBook = ActiveWorkbook
        Set CallerSheet = ActiveSheet
        OldCalculation = Application.Calculation
        OldEvents = Application.EnableEvents
        OldScreen = Application.ScreenUpdating
        OldAlerts = Application.DisplayAlerts

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
        mPrefix = HostQualifier()

        Stage = "create the workbook"
        Set Wb = Workbooks.Add(xlWBATWorksheet)
        Wb.Date1904 = False
        Wb.Styles("Normal").Font.Name = "Calibri"
        Wb.Styles("Normal").Font.Size = 11
        PrepareSheets Wb

        Stage = "write the About sheet"
        BuildAbout Wb.Worksheets(SHEET_ABOUT)
        Stage = "write the Scalar sheet"
        BuildScalar Wb, Wb.Worksheets(SHEET_SCALAR)
        Stage = "write the Arrays sheet"
        BuildArrays Wb, Wb.Worksheets(SHEET_ARRAYS)
        Stage = "write the Errors sheet"
        BuildErrors Wb, Wb.Worksheets(SHEET_ERRORS)
        Stage = "write the Pillars sheet"
        BuildPillars Wb, Wb.Worksheets(SHEET_PILLARS)
        Stage = "write the summary"
        BuildSummary Wb, Wb.Worksheets(SHEET_ABOUT)

        Stage = "calculate the workbook"
        CalculateDemo Wb

        Stage = "save the workbook"
        Wb.Worksheets(SHEET_ABOUT).Activate
        TempPath = UniqueTempPath(Folder)
        Wb.SaveAs Filename:=TempPath, FileFormat:=XL_OPEN_XML_WORKBOOK
        Stage = "close the workbook"
        Wb.Close SaveChanges:=False
        Set Wb = Nothing

    'Name refuses an existing destination, so a file that appeared at
    'OutputPath during the build is never replaced
        Stage = "publish the workbook"
        If Len(Dir$(OutputPath)) > 0 Then
            mLastReport = "The output file appeared during the build and is never overwritten: " & OutputPath
            GoTo Discard
        End If
        Name TempPath As OutputPath
        TempPath = ""

        KPR_Demo_BuildDates = True
        GoTo Restore

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        mLastReport = "Could not " & Stage & ": error " & CStr(Err.Number) & _
                      " (" & Err.Description & ")"
        Resume Discard

Discard:
    'Leave the failed workbook unsaved and remove its temporary file; the file
    'at OutputPath, if any, is never touched
        On Error Resume Next
        If Not Wb Is Nothing Then Wb.Close SaveChanges:=False
        Set Wb = Nothing
        If Len(TempPath) > 0 Then Kill TempPath

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
Restore:
        On Error Resume Next
        If Not CallerBook Is Nothing Then CallerBook.Activate
        If Not CallerSheet Is Nothing Then CallerSheet.Activate
        Application.Calculation = OldCalculation
        Application.DisplayAlerts = OldAlerts
        Application.ScreenUpdating = OldScreen
        Application.EnableEvents = OldEvents
        Err.Clear

End Function

Public Function KPR_Demo_LastReport() As String
'
'==============================================================================
'                             KPR_Demo_LastReport
'------------------------------------------------------------------------------
' PURPOSE
'   Returns why the last KPR_Demo_BuildDates call failed, or "" when it
'   succeeded or has not run.
'
' UPDATED
'   2026-10-05
'==============================================================================
'
    KPR_Demo_LastReport = mLastReport

End Function

'
'------------------------------------------------------------------------------
'
'                                    SHEETS
'
'------------------------------------------------------------------------------
'

Private Sub PrepareSheets(ByVal Wb As Workbook)
'
' Names the single sheet of a new workbook and adds the rest in order.
'
    Dim SheetNames As Variant   'Sheet names in workbook order
    Dim K          As Long      'Sheet cursor

    SheetNames = Array(SHEET_ABOUT, SHEET_SCALAR, SHEET_ARRAYS, SHEET_ERRORS, SHEET_PILLARS)
    Wb.Worksheets(1).Name = SheetNames(0)
    For K = 1 To UBound(SheetNames)
        Wb.Worksheets.Add(After:=Wb.Worksheets(Wb.Worksheets.Count)).Name = SheetNames(K)
    Next K

End Sub

Private Sub BuildAbout(ByVal Sh As Worksheet)
'
' Purpose, evidence boundary and future scope. The summary cells are added by
' BuildSummary once the check ranges exist.
'
    Sh.Range("B1").Value = "KPR date functions: demonstration workbook"
    Sh.Range("B1").Font.Bold = True
    Sh.Range("B1").Font.Size = 14
    Sh.Range("B3").Value = "Built by KPR_Demo_BuildDates from the tracked source " & _
                           "examples/modules/KPR_Demo_Dates.bas. Rebuild it rather than edit it."
    Sh.Range("B4").Value = "Each example shows a formula, its live result, the result " & _
                           "docs/DATE_LAYER_CONTRACT.md documents, and whether the two match."
    Sh.Range("B5").Value = "The live results need the K-PRICING VBA project that built " & _
                           "this workbook to be open."
    Sh.Range("B6").Value = "Scalar examples work in any supported Excel. The Arrays sheet " & _
                           "needs dynamic-array Excel; no Ctrl+Shift+Enter behaviour is claimed."
    Sh.Range("B7").Value = "This workbook demonstrates documented behaviour. It does not " & _
                           "certify accuracy, production readiness or any untested environment."
    Sh.Range("B8").Value = "Calendars, holidays and business-day arithmetic are future scope " & _
                           "(v0.0.5 and v0.0.6). Every day count and pillar here is in calendar days."
    Sh.Columns("A").ColumnWidth = 2
    Sh.Columns("B").ColumnWidth = 30
    Sh.Columns("C").ColumnWidth = 14
    Sh.Columns("D").ColumnWidth = 70

End Sub

Private Sub BuildSummary(ByVal Wb As Workbook, ByVal Sh As Worksheet)
'
' The workbook date system, the number of checked examples and the number of
' mismatches, each with a workbook name.
'
    Sh.Range("B10").Value = "Workbook date system"
    SetFormula Sh.Range("C10"), "=KPR_Dates_HostDateSystem()"
    Sh.Range("D10").Value = "Scalar only and volatile. In a 1904 workbook every " & _
                            "value-taking function returns #N/A."
    Sh.Range("B11").Value = "Examples checked"
    SetFormula Sh.Range("C11"), "=COUNTIF(Demo_Scalar_Checks,TRUE)+COUNTIF(Demo_Scalar_Checks,FALSE)" & _
                                "+COUNTIF(Demo_Array_Checks,TRUE)+COUNTIF(Demo_Array_Checks,FALSE)" & _
                                "+COUNTIF(Demo_Error_Checks,TRUE)+COUNTIF(Demo_Error_Checks,FALSE)" & _
                                "+COUNTIF(Demo_Pillar_Checks,TRUE)+COUNTIF(Demo_Pillar_Checks,FALSE)"
    Sh.Range("D11").Value = "One check per example on the Scalar, Arrays, Errors and Pillars sheets."
    Sh.Range("B12").Value = "Mismatches"
    SetFormula Sh.Range("C12"), "=COUNTIF(Demo_Scalar_Checks,FALSE)+COUNTIF(Demo_Array_Checks,FALSE)" & _
                                "+COUNTIF(Demo_Error_Checks,FALSE)+COUNTIF(Demo_Pillar_Checks,FALSE)"
    Sh.Range("D12").Value = "0 when every live result matches its documented expectation."
    Sh.Range("B10:B12").Font.Bold = True
    Sh.Range("C10:C12").HorizontalAlignment = xlLeft

    Wb.Names.Add Name:="Demo_HostDateSystem", RefersTo:="='" & SHEET_ABOUT & "'!$C$10"
    Wb.Names.Add Name:="Demo_ExampleCount", RefersTo:="='" & SHEET_ABOUT & "'!$C$11"
    Wb.Names.Add Name:="Demo_Mismatches", RefersTo:="='" & SHEET_ABOUT & "'!$C$12"

End Sub

Private Sub BuildScalar(ByVal Wb As Workbook, ByVal Sh As Worksheet)
'
' One scalar example per supported function. Expected values come from the
' reference model in tools/gen_fixtures.py.
'
    Dim R As Long   'Next table row

    WriteTableHeader Sh, "Scalar examples", _
        "One value in, one value out. These work in any supported Excel."
    R = TABLE_FIRST_ROW
        WriteRow Sh, R, "Day of week, Monday = 1", "=KPR_Dates_DayOfWeek(""2026-09-28"")", 1, "Default Monday = 1, unlike WEEKDAY": R = R + 1
        WriteRow Sh, R, "Day of week, Sunday = 1", "=KPR_Dates_DayOfWeek(""2026-09-28"",FALSE)", 2, "Opt_WeekBaseMonday = FALSE matches WEEKDAY": R = R + 1
        WriteRow Sh, R, "Days in month", "=KPR_Dates_DaysInMonth(""2024-02-10"")", 29, "Gregorian month length": R = R + 1
        WriteRow Sh, R, "Days in year", "=KPR_Dates_DaysInYear(2024)", 366, "Takes a year, not a date": R = R + 1
        WriteRow Sh, R, "Begin of month", "=KPR_Dates_BeginOfMonth(""2024-02-10"")", DateSerial(2024, 2, 1), "First day of the month": R = R + 1
        WriteRow Sh, R, "End of month", "=KPR_Dates_EndOfMonth(""2024-02-10"")", DateSerial(2024, 2, 29), "Leap-year month end": R = R + 1
        WriteRow Sh, R, "Begin of quarter", "=KPR_Dates_BeginOfQuarter(""2024-05-20"")", DateSerial(2024, 4, 1), "Calendar quarters": R = R + 1
        WriteRow Sh, R, "End of quarter", "=KPR_Dates_EndOfQuarter(""2024-05-20"")", DateSerial(2024, 6, 30), "Calendar quarters": R = R + 1
        WriteRow Sh, R, "Begin of year", "=KPR_Dates_BeginOfYear(""2024-05-20"")", DateSerial(2024, 1, 1), "1 January": R = R + 1
        WriteRow Sh, R, "End of year", "=KPR_Dates_EndOfYear(""2024-05-20"")", DateSerial(2024, 12, 31), "31 December": R = R + 1
        WriteRow Sh, R, "Is month end", "=KPR_Dates_IsMonthEnd(""2024-02-29"")", True, "Boundary predicate": R = R + 1
        WriteRow Sh, R, "Is quarter end", "=KPR_Dates_IsQuarterEnd(""2024-06-30"")", True, "Boundary predicate": R = R + 1
        WriteRow Sh, R, "Is year end", "=KPR_Dates_IsYearEnd(""2024-12-31"")", True, "Boundary predicate": R = R + 1
        WriteRow Sh, R, "Is leap year", "=KPR_Dates_IsLeapYear(1900)", False, "Century rule: 1900 is not leap": R = R + 1
        WriteRow Sh, R, "Add days", "=KPR_Dates_AddDays(""2024-02-28"",2)", DateSerial(2024, 3, 1), "Exact calendar days": R = R + 1
        WriteRow Sh, R, "Add weeks", "=KPR_Dates_AddWeeks(""2024-01-31"",4)", DateSerial(2024, 2, 28), "Weeks of 7 days": R = R + 1
        WriteRow Sh, R, "Add months (clip)", "=KPR_Dates_AddMonths(""2024-01-31"",1)", DateSerial(2024, 2, 29), "Clips to the shorter month end": R = R + 1
        WriteRow Sh, R, "Add months (keep EOM)", "=KPR_Dates_AddMonths(""2024-02-29"",1,TRUE)", DateSerial(2024, 3, 31), "Opt_KeepEOM keeps month ends": R = R + 1
        WriteRow Sh, R, "Add years", "=KPR_Dates_AddYears(""2024-02-29"",1)", DateSerial(2025, 2, 28), "29 Feb clips to 28 Feb": R = R + 1
        WriteRow Sh, R, "Nth weekday of month", "=KPR_Dates_NthWeekdayOfMonth(2026,9,1,2)", DateSerial(2026, 9, 14), "Second Monday of September 2026": R = R + 1
        WriteRow Sh, R, "Last weekday of month", "=KPR_Dates_LastWeekdayOfMonth(2026,9,5)", DateSerial(2026, 9, 25), "Last Friday of September 2026": R = R + 1
        WriteRow Sh, R, "Pillar from dates", "=KPR_Dates_PillarFromDates(""2024-01-15"",""2024-04-15"")", "3M", "Canonical emitted pillar": R = R + 1
        WriteRow Sh, R, "Date from pillar", "=KPR_Dates_DateFromPillar(""2024-01-31"",""1M"")", DateSerial(2024, 2, 29), "Month delta with clip semantics": R = R + 1
        WriteRow Sh, R, "Near-integer count", "=KPR_Dates_AddDays(""2024-01-31"",(0.1+0.2)*10)", DateSerial(2024, 2, 3), "Residue within 1E-9 reads as 3": R = R + 1
    NameChecks Wb, Sh, "Demo_Scalar_Checks", TABLE_FIRST_ROW, R - 1

End Sub

Private Sub BuildErrors(ByVal Wb As Workbook, ByVal Sh As Worksheet)
'
' Every native-error category of the contract. A 1904 workbook cannot be
' shown inside this 1900 workbook, so that case is described in text.
'
    Dim R As Long   'Next table row

    WriteTableHeader Sh, "Native errors", _
        "Invalid input gives #VALUE!, a date outside 1900-03-01 to 9999-12-31 gives #NUM!."
    R = TABLE_FIRST_ROW
        WriteRow Sh, R, "Locale text date", "=KPR_Dates_DaysInMonth(""31/12/2026"")", CVErr(xlErrValue), "#VALUE!: only YYYY-MM-DD text is accepted": R = R + 1
        WriteRow Sh, R, "Impossible date", "=KPR_Dates_EndOfMonth(""2024-02-30"")", CVErr(xlErrValue), "#VALUE!: no such calendar date": R = R + 1
        WriteRow Sh, R, "Fractional count", "=KPR_Dates_AddDays(""2024-01-31"",1.5)", CVErr(xlErrValue), "#VALUE!: no truncation or rounding": R = R + 1
        WriteRow Sh, R, "Boolean count", "=KPR_Dates_AddDays(""2024-01-31"",TRUE)", CVErr(xlErrValue), "#VALUE!: TRUE is not a count": R = R + 1
        WriteRow Sh, R, "Year outside 1900-9999", "=KPR_Dates_IsLeapYear(1899)", CVErr(xlErrValue), "#VALUE!: year domain": R = R + 1
        WriteRow Sh, R, "Invalid control", "=KPR_Dates_AddMonths(""2024-01-31"",1,""yes"")", CVErr(xlErrValue), "#VALUE!: controls take TRUE or FALSE only": R = R + 1
        WriteRow Sh, R, "Unknown rounding token", "=KPR_Dates_PillarFromDates(""2024-01-15"",""2024-04-15"",""ROUND"")", CVErr(xlErrValue), "#VALUE!: NEAREST, FLOOR or CEILING": R = R + 1
        WriteRow Sh, R, "Date before the window", "=KPR_Dates_EndOfMonth(""1900-02-28"")", CVErr(xlErrNum), "#NUM!: supported window starts 1900-03-01": R = R + 1
        WriteRow Sh, R, "Result after the window", "=KPR_Dates_AddDays(""9999-12-31"",1)", CVErr(xlErrNum), "#NUM!: result leaves the window": R = R + 1
        WriteRow Sh, R, "Occurrence absent", "=KPR_Dates_NthWeekdayOfMonth(2026,2,1,5)", CVErr(xlErrNum), "#NUM!: February 2026 has four Mondays": R = R + 1
        WriteRow Sh, R, "Propagated input error", "=KPR_Dates_EndOfMonth(1/0)", CVErr(xlErrDiv0), "An input error passes through unchanged": R = R + 1
        WriteRow Sh, R, "Shape mismatch", "=KPR_Dates_AddDays({""2024-01-01"";""2024-01-02""},{1,2,3})", CVErr(xlErrValue), "#VALUE!: array arguments of different shapes": R = R + 1
    NameChecks Wb, Sh, "Demo_Error_Checks", TABLE_FIRST_ROW, R - 1
    R = R + 1
    Sh.Cells(R, 2).Value = "1904 workbooks"
    Sh.Cells(R, 2).Font.Bold = True
    Sh.Cells(R, 3).Value = "In a workbook that uses the 1904 date system every value-taking " & _
                           "function returns #N/A instead of a date shifted by 1,462 days."

End Sub

Private Sub BuildPillars(ByVal Wb As Workbook, ByVal Sh As Worksheet)
'
' Pillar tokens in both directions. ON and TN are calendar days in every
' KPR_Dates_* function.
'
    Dim R As Long   'Next table row

    WriteTableHeader Sh, "Pillars", _
        "Tokens such as ON, 1W, 3M or 1Y6M; ON and TN count calendar days."
    R = TABLE_FIRST_ROW
        WriteRow Sh, R, "Date from pillar ON", "=KPR_Dates_DateFromPillar(""2024-01-31"",""ON"")", DateSerial(2024, 2, 1), "Overnight: 1 calendar day": R = R + 1
        WriteRow Sh, R, "Date from pillar TN", "=KPR_Dates_DateFromPillar(""2024-01-31"",""TN"")", DateSerial(2024, 2, 2), "Tom-next: 2 calendar days": R = R + 1
        WriteRow Sh, R, "Date from pillar 1W", "=KPR_Dates_DateFromPillar(""2024-01-31"",""1W"")", DateSerial(2024, 2, 7), "One week": R = R + 1
        WriteRow Sh, R, "Date from pillar 3M", "=KPR_Dates_DateFromPillar(""2024-01-31"",""3M"")", DateSerial(2024, 4, 30), "Months clip to month end": R = R + 1
        WriteRow Sh, R, "Date from pillar 1Y6M", "=KPR_Dates_DateFromPillar(""2024-01-31"",""1Y6M"")", DateSerial(2025, 7, 31), "Years and months combine": R = R + 1
        WriteRow Sh, R, "Date from pillar 2W3D", "=KPR_Dates_DateFromPillar(""2024-01-31"",""2W3D"")", DateSerial(2024, 2, 17), "Weeks and days combine": R = R + 1
        WriteRow Sh, R, "Date from pillar -1D", "=KPR_Dates_DateFromPillar(""2024-01-31"",""-1D"")", DateSerial(2024, 1, 30), "A sign applies to the whole token": R = R + 1
        WriteRow Sh, R, "Date from pillar -ON", "=KPR_Dates_DateFromPillar(""2024-01-31"",""-ON"")", CVErr(xlErrValue), "#VALUE!: an alias never carries a sign": R = R + 1
        WriteRow Sh, R, "Date from pillar 1M2M", "=KPR_Dates_DateFromPillar(""2024-01-31"",""1M2M"")", CVErr(xlErrValue), "#VALUE!: a unit may appear once": R = R + 1
        WriteRow Sh, R, "Exact day pillar", "=KPR_Dates_PillarFromDates(""2024-01-15"",""2024-01-20"")", "5D", "Under 7 days: exact days": R = R + 1
        WriteRow Sh, R, "Rounding NEAREST", "=KPR_Dates_PillarFromDates(""2024-01-15"",""2024-02-10"",""NEAREST"")", "1M", "Nearest candidate": R = R + 1
        WriteRow Sh, R, "Rounding FLOOR", "=KPR_Dates_PillarFromDates(""2024-01-15"",""2024-02-10"",""FLOOR"")", "3W", "Floor candidate": R = R + 1
        WriteRow Sh, R, "Rounding CEILING", "=KPR_Dates_PillarFromDates(""2024-01-15"",""2024-02-10"",""CEILING"")", "1M", "Ceiling candidate": R = R + 1
        WriteRow Sh, R, "Years and months", "=KPR_Dates_PillarFromDates(""2024-01-15"",""2025-07-15"")", "1Y6M", "12 months or more format as years": R = R + 1
        WriteRow Sh, R, "Negative interval", "=KPR_Dates_PillarFromDates(""2024-04-15"",""2024-01-15"")", "-3M", "End before start gives a minus sign": R = R + 1
    NameChecks Wb, Sh, "Demo_Pillar_Checks", TABLE_FIRST_ROW, R - 1

End Sub

Private Sub BuildArrays(ByVal Wb As Workbook, ByVal Sh As Worksheet)
'
' Dynamic-array examples. Each band holds its inputs (B:D), the spilled
' result (F:H), the expected result (J:L), one match per element (N:P) and
' the example's overall check (R). The capacity policy is shown with
' single-cell formulas, so no 100,000-cell spill is materialized.
'
    Dim Top    As Long      'Title row of the current band
    Dim Dates  As Variant   'Shared column input

    Sh.Range("B1").Value = "Dynamic-array examples"
    Sh.Range("B1").Font.Bold = True
    Sh.Range("B1").Font.Size = 14
    Sh.Range("B2").Value = "These need dynamic-array Excel. The same function names serve " & _
                           "scalar and multi-cell calls; no _Spill variant exists."
    Sh.Range("B3").Value = "Inputs"
    Sh.Range("F3").Value = "Result (spills)"
    Sh.Range("J3").Value = "Expected"
    Sh.Range("N3").Value = "Match"
    Sh.Range("R3").Value = "Check"
    Sh.Range("S3").Value = "Rule"
    Sh.Range("B3:S3").Font.Bold = True

    Dates = Grid(5, 1, DateSerial(2024, 1, 15), DateSerial(2024, 2, 29), _
                 DateSerial(2024, 11, 30), DateSerial(2023, 12, 31), DateSerial(2025, 6, 10))
    Top = 5
    Top = WriteArrayExample(Sh, Top, "Column spill", "=KPR_Dates_EndOfMonth({IN})", Dates, _
        Grid(5, 1, DateSerial(2024, 1, 31), DateSerial(2024, 2, 29), DateSerial(2024, 11, 30), _
             DateSerial(2023, 12, 31), DateSerial(2025, 6, 30)), _
        "A 5x1 range gives a 5x1 result")
    Top = WriteArrayExample(Sh, Top, "Scalar broadcast", "=KPR_Dates_AddMonths({IN},1)", Dates, _
        Grid(5, 1, DateSerial(2024, 2, 15), DateSerial(2024, 3, 29), DateSerial(2024, 12, 30), _
             DateSerial(2024, 1, 31), DateSerial(2025, 7, 10)), _
        "The scalar 1 is applied to every element")
    Top = WriteArrayExample(Sh, Top, "Row spill", "=KPR_Dates_DayOfWeek({IN})", _
        Grid(1, 3, DateSerial(2026, 9, 28), DateSerial(2026, 10, 3), DateSerial(2026, 10, 4)), _
        Grid(1, 3, 1, 6, 7), _
        "A 1x3 row gives a 1x3 result; Monday = 1")
    Top = WriteArrayExample(Sh, Top, "Two-dimensional spill", "=KPR_Dates_IsMonthEnd({IN})", _
        Grid(2, 2, DateSerial(2024, 2, 29), DateSerial(2024, 3, 15), _
             DateSerial(2024, 12, 31), DateSerial(2025, 1, 1)), _
        Grid(2, 2, True, False, True, False), _
        "Shape is preserved element by element")
    Top = WriteArrayExample(Sh, Top, "Element-level error", "=KPR_Dates_EndOfMonth({IN})", _
        Grid(3, 1, DateSerial(2024, 1, 15), "31/12/2026", DateSerial(2024, 3, 10)), _
        Grid(3, 1, DateSerial(2024, 1, 31), CVErr(xlErrValue), DateSerial(2024, 3, 31)), _
        "One invalid element gives its own #VALUE!; its neighbours are unaffected")
    Top = WriteArrayExample(Sh, Top, "Capacity: 100,000 elements", _
        "=ROWS(KPR_Dates_AddDays(SEQUENCE(100000,1,45000),0))", Empty, _
        Grid(1, 1, 100000), _
        "A 100,000-element result is allowed; ROWS keeps it out of the sheet")
    Top = WriteArrayExample(Sh, Top, "Capacity: 100,001 elements", _
        "=KPR_Dates_AddDays(SEQUENCE(100001),0)", Empty, _
        Grid(1, 1, CVErr(xlErrNum)), _
        "More than 100,000 elements gives one #NUM! for the whole call")
    NameChecks Wb, Sh, "Demo_Array_Checks", 5, Top - 1, "R"

    Sh.Columns("A").ColumnWidth = 2
    Sh.Columns("B:D").ColumnWidth = 12
    Sh.Columns("E").ColumnWidth = 2
    Sh.Columns("F:H").ColumnWidth = 12
    Sh.Columns("I").ColumnWidth = 2
    Sh.Columns("J:L").ColumnWidth = 12
    Sh.Columns("M").ColumnWidth = 2
    Sh.Columns("N:P").ColumnWidth = 8
    Sh.Columns("Q").ColumnWidth = 2
    Sh.Columns("R").ColumnWidth = 8
    Sh.Columns("S").ColumnWidth = 60

End Sub

'
'------------------------------------------------------------------------------
'
'                                   WRITERS
'
'------------------------------------------------------------------------------
'

Private Sub WriteTableHeader(ByVal Sh As Worksheet, ByVal Title As String, ByVal Subtitle As String)
'
' Title, subtitle, column headings and fixed widths of a five-column table.
'
    Sh.Range("B1").Value = Title
    Sh.Range("B1").Font.Bold = True
    Sh.Range("B1").Font.Size = 14
    Sh.Range("B2").Value = Subtitle
    Sh.Range("B4").Value = "Example"
    Sh.Range("C4").Value = "Formula"
    Sh.Range("D4").Value = "Result"
    Sh.Range("E4").Value = "Expected"
    Sh.Range("F4").Value = "Match"
    Sh.Range("G4").Value = "Rule"
    Sh.Range("B4:G4").Font.Bold = True
    Sh.Range("B4:G4").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Sh.Columns("A").ColumnWidth = 2
    Sh.Columns("B").ColumnWidth = 26
    Sh.Columns("C").ColumnWidth = 64
    Sh.Columns("D").ColumnWidth = 13
    Sh.Columns("E").ColumnWidth = 13
    Sh.Columns("F").ColumnWidth = 8
    Sh.Columns("G").ColumnWidth = 48

End Sub

Private Sub WriteRow( _
    ByVal Sh As Worksheet, _
    ByVal R As Long, _
    ByVal Label As String, _
    ByVal Formula As String, _
    ByVal Expected As Variant, _
    ByVal Rule As String)
'
' One example: label, the formula as text, the live formula, the expected
' value, the match test and the rule it shows.
'
    Sh.Cells(R, 2).Value = Label
    WriteFormulaText Sh.Cells(R, 3), Formula
    ApplyValueFormat Sh.Cells(R, 4), Expected
    SetFormula Sh.Cells(R, 4), Formula
    ApplyValueFormat Sh.Cells(R, 5), Expected
    Sh.Cells(R, 5).Value = Expected
    Sh.Cells(R, 6).Formula = MatchFormula(Sh.Cells(R, 4).Address(False, False), _
                                          Sh.Cells(R, 5).Address(False, False))
    Sh.Cells(R, 7).Value = Rule

End Sub

Private Function WriteArrayExample( _
    ByVal Sh As Worksheet, _
    ByVal Top As Long, _
    ByVal Title As String, _
    ByVal Template As String, _
    ByVal Inputs As Variant, _
    ByVal Expected As Variant, _
    ByVal Rule As String) _
    As Long
'
' One band of the Arrays sheet. {IN} in Template is replaced by the address of
' the input block. Returns the title row of the next band.
'
    Dim RowCount As Long     'Result rows
    Dim ColCount As Long     'Result columns
    Dim I        As Long     'Row cursor
    Dim J        As Long     'Column cursor
    Dim Formula  As String   'Unqualified formula
    Dim InBlock  As Range    'Input block, when the example has inputs

    RowCount = UBound(Expected, 1)
    ColCount = UBound(Expected, 2)
    Formula = Template
    If Not IsEmpty(Inputs) Then
        Set InBlock = Sh.Cells(Top + 1, 2).Resize(UBound(Inputs, 1), UBound(Inputs, 2))
        For I = 1 To UBound(Inputs, 1)
            For J = 1 To UBound(Inputs, 2)
                ApplyValueFormat InBlock.Cells(I, J), Inputs(I, J)
                InBlock.Cells(I, J).Value = Inputs(I, J)
            Next J
        Next I
        Formula = Replace(Formula, "{IN}", InBlock.Address(False, False))
    End If

    Sh.Cells(Top, 2).Value = Title
    Sh.Cells(Top, 2).Font.Bold = True
    WriteFormulaText Sh.Cells(Top, 6), Formula
    Sh.Cells(Top, 19).Value = Rule

    For I = 1 To RowCount
        For J = 1 To ColCount
            ApplyValueFormat Sh.Cells(Top + I, 5 + J), Expected(I, J)
            ApplyValueFormat Sh.Cells(Top + I, 9 + J), Expected(I, J)
            Sh.Cells(Top + I, 9 + J).Value = Expected(I, J)
            Sh.Cells(Top + I, 13 + J).Formula = _
                MatchFormula(Sh.Cells(Top + I, 5 + J).Address(False, False), _
                             Sh.Cells(Top + I, 9 + J).Address(False, False))
        Next J
    Next I
    SetFormula Sh.Cells(Top + 1, 6), Formula
    Sh.Cells(Top, 18).Formula = "=AND(" & _
        Sh.Cells(Top + 1, 14).Resize(RowCount, ColCount).Address(False, False) & ")"

    WriteArrayExample = Top + RowCount + 2

End Function

Private Sub NameChecks( _
    ByVal Wb As Workbook, _
    ByVal Sh As Worksheet, _
    ByVal CheckName As String, _
    ByVal FirstRow As Long, _
    ByVal LastRow As Long, _
    Optional ByVal CheckColumn As String = "F")
'
' Names the column of match results of one sheet.
'
    Wb.Names.Add Name:=CheckName, RefersTo:="='" & Sh.Name & "'!$" & CheckColumn & "$" & _
                 CStr(FirstRow) & ":$" & CheckColumn & "$" & CStr(LastRow)

End Sub

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub WriteFormulaText(ByVal Target As Range, ByVal Formula As String)
'
' Shows a formula as text: a text number format and a leading apostrophe, so
' Excel never evaluates it.
'
    Target.NumberFormat = "@"
    Target.Value = "'" & Formula

End Sub

Private Sub SetFormula(ByVal Target As Object, ByVal Formula As String)
'
' Writes Formula with the host qualifier. Formula2 keeps dynamic-array
' semantics where Excel provides it; earlier Excel falls back to Formula.
'
    Dim Qualified As String     'Formula as Excel must receive it

    Qualified = Replace(Formula, "KPR_Dates_", mPrefix & "KPR_Dates_")
    On Error Resume Next
    Target.Formula2 = Qualified
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Target.Formula = Qualified
    End If
    On Error GoTo 0

End Sub

Private Function HostQualifier() As String
'
' "" when this project is an add-in; otherwise 'Workbook name'! so a formula
' in another workbook reaches this project's functions.
'
    If ThisWorkbook.IsAddin Then
        HostQualifier = ""
    Else
        HostQualifier = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!"
    End If

End Function

Private Function MatchFormula(ByVal Actual As String, ByVal Expected As String) As String
'
' TRUE when the live result equals the expected value, including the same
' error type when an error is expected. Never itself an error.
'
    MatchFormula = "=IF(ISERROR(" & Expected & "),IFERROR(ERROR.TYPE(" & Actual & _
                   ")=ERROR.TYPE(" & Expected & "),FALSE),IFERROR(" & Actual & "=" & _
                   Expected & ",FALSE))"

End Function

Private Sub ApplyValueFormat(ByVal Target As Range, ByVal Value As Variant)
'
' Dates show as YYYY-MM-DD, text stays text, everything else is General.
'
    Select Case VarType(Value)
        Case vbDate
            Target.NumberFormat = DATE_FORMAT
        Case vbString
            Target.NumberFormat = "@"
        Case Else
            Target.NumberFormat = "General"
    End Select

End Sub

Private Function Grid(ByVal RowCount As Long, ByVal ColCount As Long, ParamArray Items() As Variant) As Variant
'
' A 1-based RowCount x ColCount array filled row by row from Items.
'
    Dim Result() As Variant     '1-based grid
    Dim I        As Long        'Row cursor
    Dim J        As Long        'Column cursor

    ReDim Result(1 To RowCount, 1 To ColCount)
    For I = 1 To RowCount
        For J = 1 To ColCount
            Result(I, J) = Items((I - 1) * ColCount + (J - 1))
        Next J
    Next I
    Grid = Result

End Function

Private Sub CalculateDemo(ByVal Wb As Workbook)
'
' Calculates only the demo sheets, twice, so cross-sheet summaries see the
' final check results. No other open workbook is recalculated.
'
    Dim Pass As Long        'Calculation pass
    Dim Sh   As Worksheet   'Sheet cursor

    For Pass = 1 To 2
        For Each Sh In Wb.Worksheets
            Sh.Calculate
        Next Sh
    Next Pass

End Sub

Private Function TryValidatePath( _
    ByVal OutputPath As String, _
    ByRef Folder As String, _
    ByRef Detail As String) _
    As Boolean
'
' Accepts only a new .xlsx file in an existing folder; Folder receives the
' parent folder.
'
    Dim Cut    As Long      'Position of the last path separator

    TryValidatePath = False
    If Len(OutputPath) = 0 Then
        Detail = "An output path is required."
        Exit Function
    End If
    If LCase$(Right$(OutputPath, 5)) <> ".xlsx" Then
        Detail = "The output path must end in .xlsx: " & OutputPath
        Exit Function
    End If
    Cut = InStrRev(OutputPath, Application.PathSeparator)
    If Cut < 2 Then
        Detail = "The output path must include its folder: " & OutputPath
        Exit Function
    End If
    Folder = Left$(OutputPath, Cut - 1)

    On Error Resume Next
    If Len(Dir$(Folder, vbDirectory)) = 0 Then
        Detail = "The output folder does not exist: " & Folder
        Exit Function
    End If
    If Len(Dir$(OutputPath)) > 0 Then
        Detail = "The output file already exists and is never overwritten: " & OutputPath
        Exit Function
    End If
    If Err.Number <> 0 Then
        Detail = "The output path cannot be checked: " & OutputPath
        Exit Function
    End If
    On Error GoTo 0

    TryValidatePath = True

End Function

Private Function UniqueTempPath(ByVal Folder As String) As String
'
' A file name in Folder that does not exist yet, for the save before the
' final rename.
'
    Dim Candidate As String     'Proposed temporary path
    Dim Attempt   As Long       'Attempt counter

    Randomize
    Do
        Attempt = Attempt + 1
        Candidate = Folder & Application.PathSeparator & "~kpr-demo-" & _
                    Format$(Now, "yyyymmddhhnnss") & "-" & CStr(Int(Rnd * 1000000)) & ".xlsx"
    Loop While Len(Dir$(Candidate)) > 0 And Attempt < 100
    UniqueTempPath = Candidate

End Function
