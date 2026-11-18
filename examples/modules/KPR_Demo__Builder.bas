Attribute VB_Name = "KPR_Demo__Builder"
'==============================================================================
' MODULE: KPR_Demo__Builder
'------------------------------------------------------------------------------
' PURPOSE
'   Shared engine for K-PRICING demonstration sheets. A demo content module
'   (for example KPR_Demo_Dates) describes its sheet as a short list of calls;
'   this module owns the layout, the house style, the status rules, the
'   summary and the Excel state around the build.
'
' WHY THIS EXISTS
'   Every demo must look the same, be rebuilt from tracked source rather than
'   stored as a workbook binary, and be checkable by the regression suite.
'   Keeping layout and style in one module means a new demo is one small
'   content module plus one catalog line, and a change of house style is one
'   change here.
'
' SHEET LAYOUT
'   Row 1-2     title band and subtitle band across B:M, 24 and 26.1 points
'               high as on the Home sheet
'   Column A    with the K-PRICING Home sheet in the workbook: the black
'               navigation column of Home (21.7 wide, B a 4-wide gutter),
'               rows 1 to 3 and column A frozen, and the navigation bar of
'               Home as two shape groups (NavBarTop with this sheet
'               highlighted, NavBar with the link to Home); without it, a
'               narrow margin and no bar
'   Main column C:G, top to bottom:
'     inputs    labels in C, editable values in D (orange), notes in E;
'               each value gets a sheet-level name used by the formulas
'     sections  Formula | KPR UDF | Expected (Excel) | Status | Notes; the
'               Formula column shows the KPR formula as text, D holds it live
'               and E holds a native-Excel reference formula
'     notes     free text lines
'   Side panel I:M, top to bottom:
'     checks    counts of OK, DIFFERS and FAIL, with sheet-level names
'     cases     Input | Result | Expected | Status | Rule, for fixed cases
'               such as native errors
'     spills    one source spill and up to three dependent spills, each with
'               a status cell
'
' STATUS RULES
'   OK       the KPR result equals the reference, or both are the same error
'   DIFFERS  they differ while the row's DiffersWhen condition is TRUE: a
'            documented difference by design, never a failure
'   FAIL     any other difference
'   NO VBA   the KPR cell shows #NAME?: the K-PRICING project is not loaded
'   Fails in the summary are every status that is neither OK nor DIFFERS.
'
' WHERE A DEMO IS BUILT
'   Without an output path the demo is a new sheet in the active workbook,
'   named after the demo with " (2)", " (3)" ... added when the name is
'   taken. No existing sheet is ever changed, replaced or deleted. When no
'   ordinary workbook is active, or it is a 1904 workbook, read-only or
'   structure-protected, the demo goes into a new unsaved workbook instead.
'   With an output path the demo goes into a new workbook saved as that .xlsx
'   file, which must not exist yet; it is saved under a temporary name and
'   then renamed with the Name statement, which refuses an existing file.
'
' FORMULAS AND THE HOST
'   When this project is an add-in, or the demo sheet is in this project's own
'   workbook, formulas use the bare KPR_ names. Otherwise every "KPR_" in a
'   formula is qualified with this workbook's name, as Excel requires for a
'   function in another open workbook. The Formula column always shows the
'   unqualified text. Formulas are written with Range.Formula2 where Excel
'   provides it.
'
' STATE
'   Calculation, events, screen updating and alerts are restored on every
'   path. After a build into an existing or new unsaved workbook the demo
'   sheet is left active so the user sees it; after a build to a file the
'   caller's workbook and sheet are made active again. Ctrl+Z cannot remove a
'   sheet added by a macro, so the summary tells the user to delete the sheet.
'
' ERROR POLICY
'   Demo_Begin and Demo_Finish never raise: they return FALSE and
'   KPR_Demo_LastReport says why. The writers raise to the content module,
'   whose handler calls Demo_Abort, which removes everything the build
'   created and restores the state.
'
' VISIBILITY
'   Option Private Module. Callable from this project, the Immediate window
'   and Application.Run.
'
' ALLOWED DEPENDENCIES
'   No KPR module: function names appear only inside formula text, and
'   catalog entries are run by name.
'
' UPDATED
'   2026-11-18
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
' PUBLIC CONSTANTS
'------------------------------------------------------------------------------
    'Display format of every date a demo shows
        Public Const DEMO_DATE As String = "yyyy-mm-dd"

'------------------------------------------------------------------------------
' CONSTANTS
'------------------------------------------------------------------------------
    'Main column
        Private Const COL_TITLE    As Long = 2      'B: title bands start here
        Private Const COL_FORMULA  As Long = 3      'C: input labels, formula text
        Private Const COL_KPR      As Long = 4      'D: input values, live KPR result
        Private Const COL_EXCEL    As Long = 5      'E: input notes, Excel reference
        Private Const COL_STATUS   As Long = 6      'F: status
        Private Const COL_NOTES    As Long = 7      'G: notes
    'Side panel
        Private Const COL_S_INPUT  As Long = 9      'I: case input, spill source
        Private Const COL_S_RESULT As Long = 10     'J: case result
        Private Const COL_S_EXPECT As Long = 11     'K: case expected value
        Private Const COL_S_STATUS As Long = 12     'L: case status
        Private Const COL_S_RULE   As Long = 13     'M: case rule
        Private Const COL_LAST     As Long = 13     'Last column of the title bands

    'First row below the title bands
        Private Const FIRST_ROW As Long = 4

    'Status texts
        Private Const STATUS_OK      As String = "OK"
        Private Const STATUS_DIFFERS As String = "DIFFERS"

    'House style
        Private Const FONT_NAME    As String = "Aptos"
        Private Const SIZE_BODY    As Long = 10
        Private Const SIZE_HEADER  As Long = 11
        Private Const SIZE_BAND    As Long = 12
        Private Const SIZE_TITLE   As Long = 14

    'Excel's Open XML workbook format and the longest sheet name
        Private Const XL_OPEN_XML_WORKBOOK As Long = 51
        Private Const SHEET_NAME_MAX       As Long = 31

    'Layout shared with the Home sheet of the template (tools/workbook_template.py)
        Private Const HOME_SHEET       As String = "Home"       'Sheet that holds the navigation bar
        Private Const HOME_WORDMARK    As String = "K-PRICING"  'Home!B1, to tell the K-PRICING Home
        Private Const ROW_1_HEIGHT     As Double = 24           'Black title band, points
        Private Const ROW_2_HEIGHT     As Double = 26.1         'Navy subtitle band, points
        Private Const FROZEN_ROWS      As Long = 3              'Rows frozen with column A
        Private Const NAV_COL_A        As Double = 21.7         'Navigation column, characters
        Private Const NAV_COL_B        As Double = 4            'Gutter, characters
    'Navigation bar, in pixels at 96 dpi (a pixel is 0.75 point)
        Private Const NAV_BAR_X        As Long = 8              'Left margin inside column A
        Private Const NAV_BAR_W        As Long = 141            'Width of the bar
        Private Const NAV_PAD_TOP      As Long = 6              'Above the section title
        Private Const NAV_HEAD_H       As Long = 17             'Section title height
        Private Const NAV_BTN_H        As Long = 21             'Button height
        Private Const NAV_STRIP_W      As Long = 3              'Coloured strip of a button
        Private Const NAV_GROUP_TOP    As String = "NavBarTop"  'Group in the frozen rows
        Private Const NAV_GROUP_REST   As String = "NavBar"     'Group below the frozen rows
        Private Const NAV_TEXT_SIZE    As Double = 9
        Private Const NAV_HEAD_SIZE    As Double = 7
    'Office constants used by the navigation bar
        Private Const MSO_RECTANGLE    As Long = 1
        Private Const MSO_ANCHOR_MID   As Long = 3
        Private Const MSO_ANCHOR_BASE  As Long = 4
        Private Const MSO_ALIGN_LEFT   As Long = 1

'------------------------------------------------------------------------------
' BUILD STATE
'------------------------------------------------------------------------------
    Private mLastReport     As String       'Failure description of the last build
    Private mBuilding       As Boolean      'A build is between Begin and Finish
    Private mWb             As Workbook     'Workbook holding the demo sheet
    Private mSh             As Worksheet    'The demo sheet
    Private mOwnBook        As Boolean      'mWb was created by this build
    Private mOutputPath     As String       'File to publish, or ""
    Private mFolder         As String       'Folder of mOutputPath
    Private mTempPath       As String       'Temporary file before the rename
    Private mPrefix         As String       'Qualifier inserted before KPR_
    Private mDynamic        As Boolean      'Excel provides Range.Formula2
    Private mNav            As Boolean      'The workbook has the K-PRICING Home: draw the navigation bar
    Private mRow            As Long         'Next free row of the main column
    Private mSideRow        As Long         'Next free row of the side panel
    Private mSummaryRow     As Long         'First row of the checks summary
    Private mStatus         As Collection   'Address of every status cell
    Private mSpillAnchor    As Range        'Source cell of the open spill block
    Private mSpillCheckRow  As Long         'Status row of the open spill block
    Private mSpillCol       As Long         'Next column of the open spill block
    Private mSpillOpen      As Boolean      'A spill block accepts columns
    'Excel state on entry
    Private mCallerBook     As Workbook
    Private mCallerSheet    As Object
    Private mOldCalculation As Long
    Private mOldEvents      As Boolean
    Private mOldScreen      As Boolean
    Private mOldAlerts      As Boolean

'
'------------------------------------------------------------------------------
'
'                                   CATALOG
'
'------------------------------------------------------------------------------
'

Public Function KPR_Demo_Catalog() As Variant
'
'==============================================================================
'                              KPR_Demo_Catalog
'------------------------------------------------------------------------------
' PURPOSE
'   The demos this project can build, one "EntryPoint|Caption|Group" text per
'   demo. A new demo adds one line here; KPR_Demo_Build, the regression suite
'   and the Ribbon and CommandBars menus (KPR_UI_Actions) read this list.
'
' FIELDS
'   EntryPoint  the content module's build procedure
'   Caption     the menu item text
'   Group       the functional module the demo belongs to; the menus show
'               one Demo sheets submenu per group, so it matches a caption
'               of KPR_UI_FunctionGroups
'
' RETURNS
'   A zero-based array of "EntryPoint|Caption|Group" texts.
'
' UPDATED
'   2026-11-03
'==============================================================================
'
    KPR_Demo_Catalog = Array( _
        "KPR_Demo_BuildDates|Date primitives|Date primitives")

End Function

Public Function KPR_Demo_Build( _
    ByVal EntryPoint As String, _
    Optional ByVal OutputPath As String = "") _
    As Boolean
'
'==============================================================================
'                               KPR_Demo_Build
'------------------------------------------------------------------------------
' PURPOSE
'   Builds one catalog demo by its entry-point name.
'
' INPUTS
'   EntryPoint
'     An entry point listed by KPR_Demo_Catalog.
'   OutputPath
'     Optional. "" builds into the active workbook; a path builds into a new
'     .xlsx file there.
'
' RETURNS
'   The entry point's result; FALSE for a name outside the catalog.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
    Dim Entry As Variant    'Catalog cursor

    KPR_Demo_Build = False
    For Each Entry In KPR_Demo_Catalog()
        If StrComp(Split(CStr(Entry), "|")(0), EntryPoint, vbTextCompare) = 0 Then
            On Error GoTo Fail
            KPR_Demo_Build = CBool(Application.Run("'" & Replace(ThisWorkbook.Name, "'", "''") & _
                                                   "'!" & EntryPoint, OutputPath))
            Exit Function
        End If
    Next Entry
    mLastReport = "Not a catalog demo: " & EntryPoint
    Exit Function

Fail:
    mLastReport = "Could not run " & EntryPoint & ": error " & CStr(Err.Number) & _
                  " (" & Err.Description & ")"

End Function

Public Function KPR_Demo_LastReport() As String
'
'==============================================================================
'                             KPR_Demo_LastReport
'------------------------------------------------------------------------------
' PURPOSE
'   Returns why the last demo build failed, or "" when it succeeded or none
'   has run.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
    KPR_Demo_LastReport = mLastReport

End Function

'
'------------------------------------------------------------------------------
'
'                                  LIFECYCLE
'
'------------------------------------------------------------------------------
'

Public Function Demo_Begin( _
    ByVal SheetName As String, _
    ByVal Title As String, _
    ByVal Subtitle As String, _
    ByVal OutputPath As String) _
    As Boolean
'
'==============================================================================
'                                 Demo_Begin
'------------------------------------------------------------------------------
' PURPOSE
'   Starts a demo: checks the destination, saves the Excel state, creates the
'   demo sheet and writes the title bands.
'
' INPUTS
'   SheetName   preferred sheet name; a free variant is used when taken
'   Title       text of the black title band
'   Subtitle    text of the navy subtitle band
'   OutputPath  "" for the active workbook, or a new .xlsx file path
'
' RETURNS
'   TRUE when the sheet is ready for the writers. FALSE otherwise, with the
'   state already restored and KPR_Demo_LastReport set.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
    Dim Detail As String    'Path validation message
    Dim Stage  As String    'Step, for the failure report

    Demo_Begin = False
    mLastReport = ""
    If mBuilding Then
        mLastReport = "Another demo build is still in progress."
        Exit Function
    End If
    If Len(OutputPath) > 0 Then
        If Not TryValidatePath(OutputPath, mFolder, Detail) Then
            mLastReport = Detail
            Exit Function
        End If
    End If

    On Error GoTo Fail
    Stage = "capture the Excel state"
    mBuilding = True
    mOutputPath = OutputPath
    mTempPath = ""
    mOwnBook = False
    Set mWb = Nothing
    Set mSh = Nothing
    Set mCallerBook = ActiveWorkbook
    Set mCallerSheet = ActiveSheet
    mOldCalculation = Application.Calculation
    mOldEvents = Application.EnableEvents
    mOldScreen = Application.ScreenUpdating
    mOldAlerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Stage = "create the demo sheet"
    If Len(OutputPath) = 0 And CanHostDemo(mCallerBook) Then
        Set mWb = mCallerBook
        Set mSh = mWb.Worksheets.Add(After:=mWb.Sheets(mWb.Sheets.Count))
    Else
        Set mWb = Workbooks.Add(xlWBATWorksheet)
        mOwnBook = True
        mWb.Date1904 = False
        Set mSh = mWb.Worksheets(1)
    End If
    mSh.Name = FreeSheetName(mWb, SheetName)
    mPrefix = HostQualifier(mWb)
    mNav = HasKprHome(mWb)
    mDynamic = HasFormula2(mSh.Range("A1"))
    Set mStatus = New Collection
    mSpillOpen = False

    Stage = "write the title"
    PrepareSheet
    WriteTitle Title, Subtitle
    mRow = FIRST_ROW
    mSummaryRow = FIRST_ROW
    mSideRow = mSummaryRow + 7

    Demo_Begin = True
    Exit Function

Fail:
    mLastReport = "Could not " & Stage & ": error " & CStr(Err.Number) & " (" & Err.Description & ")"
    Resume Discard

Discard:
    DiscardBuild

End Function

Public Function Demo_Finish() As Boolean
'
'==============================================================================
'                                 Demo_Finish
'------------------------------------------------------------------------------
' PURPOSE
'   Completes a demo: writes the checks summary, calculates the sheet and,
'   with an output path, saves and publishes the file.
'
' RETURNS
'   TRUE when the demo is complete. FALSE otherwise, with everything the
'   build created removed and KPR_Demo_LastReport set.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
    Dim Stage As String     'Step, for the failure report

    Demo_Finish = False
    If Not mBuilding Then
        mLastReport = "Demo_Finish was called without Demo_Begin."
        Exit Function
    End If

    On Error GoTo Fail
    Stage = "write the summary"
    CloseSpill
    WriteSummary
    Stage = "lay out the sheet"
    mSh.UsedRange.Rows.AutoFit
    AlignTitleRows
    DrawNavigation
    Stage = "calculate the sheet"
    mSh.Calculate
    mSh.Calculate
    ShowSheet

    If Len(mOutputPath) > 0 Then
        Stage = "save the workbook"
        mTempPath = UniqueTempPath(mFolder)
        mWb.SaveAs Filename:=mTempPath, FileFormat:=XL_OPEN_XML_WORKBOOK
        Stage = "close the workbook"
        mWb.Close SaveChanges:=False
        Set mWb = Nothing
        Set mSh = Nothing
        Stage = "publish the workbook"
        If Len(Dir$(mOutputPath)) > 0 Then
            mLastReport = "The output file appeared during the build and is never overwritten: " & mOutputPath
            GoTo Discard
        End If
        Name mTempPath As mOutputPath
        mTempPath = ""
        RestoreState True
    Else
        RestoreState False
    End If

    Demo_Finish = True
    Exit Function

Fail:
    mLastReport = "Could not " & Stage & ": error " & CStr(Err.Number) & " (" & Err.Description & ")"
    Resume Discard

Discard:
    DiscardBuild

End Function

Public Function Demo_Abort(ByVal Number As Long, ByVal Description As String) As Boolean
'
'==============================================================================
'                                 Demo_Abort
'------------------------------------------------------------------------------
' PURPOSE
'   Ends a failed demo from the content module's error handler: removes the
'   demo sheet or the new workbook, restores the state and records the error.
'
' INPUTS
'   Number, Description   the error the content module caught
'
' RETURNS
'   FALSE, so a content module can write Build = Demo_Abort(...).
'
' UPDATED
'   2026-10-06
'==============================================================================
'
    If Len(mLastReport) = 0 Then
        mLastReport = "Could not build the demo: error " & CStr(Number) & " (" & Description & ")"
    End If
    If mBuilding Then DiscardBuild
    Demo_Abort = False

End Function

Public Function Demo_HasDynamicArrays() As Boolean
'
'==============================================================================
'                           Demo_HasDynamicArrays
'------------------------------------------------------------------------------
' PURPOSE
'   TRUE when the Excel running the build has dynamic arrays (Range.Formula2),
'   so a content module can add array-only examples.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
    Demo_HasDynamicArrays = mDynamic

End Function

'
'------------------------------------------------------------------------------
'
'                                MAIN COLUMN
'
'------------------------------------------------------------------------------
'

Public Sub Demo_Inputs(ByVal Caption As String)
'
'==============================================================================
'                                 Demo_Inputs
'------------------------------------------------------------------------------
' PURPOSE
'   Starts the block of editable inputs with a band across C:E.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
    RequireBuild
    CloseSpill
    WriteBand mSh.Range(mSh.Cells(mRow, COL_FORMULA), mSh.Cells(mRow, COL_EXCEL)), Caption
    mRow = mRow + 1

End Sub

Public Sub Demo_Input( _
    ByVal Label As String, _
    ByVal InputName As String, _
    ByVal Value As Variant, _
    Optional ByVal Note As String = "")
'
'==============================================================================
'                                 Demo_Input
'------------------------------------------------------------------------------
' PURPOSE
'   One editable input: label in C, value in D, note in E. InputName becomes
'   a sheet-level name of the value cell, so formulas read DateIn rather than
'   $D$5 and keep working in a renamed or second copy of the demo.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
    Dim Row As Range    'C:E of the input row

    RequireBuild
    Set Row = mSh.Range(mSh.Cells(mRow, COL_FORMULA), mSh.Cells(mRow, COL_EXCEL))
    StyleRow Row
    With mSh.Cells(mRow, COL_FORMULA)
        .Value = Label
        .Font.Bold = True
        .Interior.Color = ColourHeader()
    End With
    With mSh.Cells(mRow, COL_KPR)
        ApplyValueFormat .Cells(1, 1), Value
        .Value = Value
        .Interior.Color = ColourInput()
        .Font.Color = ColourInputText()
        .HorizontalAlignment = xlCenter
    End With
    With mSh.Cells(mRow, COL_EXCEL)
        .Value = Note
        .Font.Italic = True
        .WrapText = False
    End With
    mSh.Names.Add Name:=InputName, RefersTo:="=" & SheetRef() & "!$" & ColumnLetter(COL_KPR) & "$" & CStr(mRow)
    mRow = mRow + 1

End Sub

Public Sub Demo_Section(ByVal Caption As String)
'
'==============================================================================
'                                Demo_Section
'------------------------------------------------------------------------------
' PURPOSE
'   Starts a comparison table: a band and the header row Formula | KPR UDF |
'   Expected (Excel) | Status | Notes across C:G.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
    RequireBuild
    mRow = mRow + 1
    WriteBand mSh.Range(mSh.Cells(mRow, COL_FORMULA), mSh.Cells(mRow, COL_NOTES)), Caption
    mRow = mRow + 1
    WriteHeader mSh.Cells(mRow, COL_FORMULA), Array("Formula", "KPR UDF", "Expected (Excel)", "Status", "Notes")
    mRow = mRow + 1

End Sub

Public Sub Demo_Compare( _
    ByVal Formula As String, _
    ByVal Reference As String, _
    ByVal Note As String, _
    Optional ByVal NumberFormat As String = "General", _
    Optional ByVal DiffersWhen As String = "")
'
'==============================================================================
'                                Demo_Compare
'------------------------------------------------------------------------------
' PURPOSE
'   One comparison row: the KPR formula as text and live, the native-Excel
'   reference, the status and a note.
'
' INPUTS
'   Formula       the KPR formula, starting with "="
'   Reference     the native-Excel formula or value it must equal
'   Note          what the row shows
'   NumberFormat  display format of both results
'   DiffersWhen   optional Excel condition under which a difference is by
'                 design and shows DIFFERS instead of FAIL
'
' UPDATED
'   2026-10-06
'==============================================================================
'
    Dim Kpr As Range    'Live KPR cell
    Dim Ref As Range    'Reference cell

    RequireBuild
    StyleRow mSh.Range(mSh.Cells(mRow, COL_FORMULA), mSh.Cells(mRow, COL_NOTES))
    WriteFormulaText mSh.Cells(mRow, COL_FORMULA), Formula
    Set Kpr = mSh.Cells(mRow, COL_KPR)
    Set Ref = mSh.Cells(mRow, COL_EXCEL)
    Kpr.NumberFormat = NumberFormat
    Ref.NumberFormat = NumberFormat
    SetFormula Kpr, Formula
    SetFormula Ref, Reference
    Kpr.HorizontalAlignment = xlCenter
    Ref.HorizontalAlignment = xlCenter
    WriteStatus mSh.Cells(mRow, COL_STATUS), Kpr, Ref, DiffersWhen
    With mSh.Cells(mRow, COL_NOTES)
        .Value = Note
        .WrapText = True
    End With
    mRow = mRow + 1

End Sub

Public Sub Demo_Note(ByVal Text As String)
'
'==============================================================================
'                                  Demo_Note
'------------------------------------------------------------------------------
' PURPOSE
'   A line of free text across the main column.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
    RequireBuild
    mRow = mRow + 1
    With mSh.Cells(mRow, COL_FORMULA)
        .Value = Text
        .Font.Italic = True
    End With
    mRow = mRow + 1

End Sub

'
'------------------------------------------------------------------------------
'
'                                 SIDE PANEL
'
'------------------------------------------------------------------------------
'

Public Sub Demo_Cases(ByVal Caption As String)
'
'==============================================================================
'                                 Demo_Cases
'------------------------------------------------------------------------------
' PURPOSE
'   Starts a side-panel table of fixed cases: a band and the header row
'   Input | Result | Expected | Status | Rule across I:M.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
    RequireBuild
    CloseSpill
    WriteBand mSh.Range(mSh.Cells(mSideRow, COL_S_INPUT), mSh.Cells(mSideRow, COL_S_RULE)), Caption
    mSideRow = mSideRow + 1
    WriteHeader mSh.Cells(mSideRow, COL_S_INPUT), Array("Input", "Result", "Expected", "Status", "Rule")
    mSideRow = mSideRow + 1

End Sub

Public Sub Demo_Case( _
    ByVal InputValue As Variant, _
    ByVal Formula As String, _
    ByVal Expected As Variant, _
    ByVal Rule As String, _
    Optional ByVal NumberFormat As String = "General")
'
'==============================================================================
'                                  Demo_Case
'------------------------------------------------------------------------------
' PURPOSE
'   One fixed case: an input value, a formula that may refer to it as {IN},
'   the expected value or native error, the status and the rule it shows.
'
' INPUTS
'   InputValue    value of the input cell; Empty leaves it blank and a text
'                 is stored as text
'   Formula       the KPR formula, starting with "="
'   Expected      the expected value, or CVErr(...) for a native error
'   Rule          what the case shows
'   NumberFormat  display format of the input, result and expected cells
'
' UPDATED
'   2026-10-06
'==============================================================================
'
    Dim InCell  As Range    'Input cell
    Dim Result  As Range    'Live result cell
    Dim Expect  As Range    'Expected cell

    RequireBuild
    StyleRow mSh.Range(mSh.Cells(mSideRow, COL_S_INPUT), mSh.Cells(mSideRow, COL_S_RULE))
    Set InCell = mSh.Cells(mSideRow, COL_S_INPUT)
    Set Result = mSh.Cells(mSideRow, COL_S_RESULT)
    Set Expect = mSh.Cells(mSideRow, COL_S_EXPECT)
    If Not IsEmpty(InputValue) Then
        If VarType(InputValue) = vbString Then
            InCell.NumberFormat = "@"
        ElseIf VarType(InputValue) = vbDate Then
            InCell.NumberFormat = DEMO_DATE
        Else
            InCell.NumberFormat = NumberFormat
        End If
        InCell.Value = InputValue
    End If
    Result.NumberFormat = NumberFormat
    Expect.NumberFormat = NumberFormat
    SetFormula Result, Replace(Formula, "{IN}", InCell.Address(False, False))
    Expect.Value = Expected
    InCell.HorizontalAlignment = xlCenter
    Result.HorizontalAlignment = xlCenter
    Expect.HorizontalAlignment = xlCenter
    WriteStatus mSh.Cells(mSideRow, COL_S_STATUS), Result, Expect, ""
    With mSh.Cells(mSideRow, COL_S_RULE)
        .Value = Rule
        .WrapText = True
    End With
    mSideRow = mSideRow + 1

End Sub

Public Sub Demo_Spill( _
    ByVal Caption As String, _
    ByVal SourceHeader As String, _
    ByVal SourceFormula As String, _
    ByVal SourceRows As Long, _
    Optional ByVal NumberFormat As String = "General")
'
'==============================================================================
'                                 Demo_Spill
'------------------------------------------------------------------------------
' PURPOSE
'   Starts a side-panel spill block: one source formula in column I that
'   spills SourceRows rows. Demo_SpillColumn then adds up to three dependent
'   spills in J:L. Without dynamic arrays the block shows a note instead.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
    RequireBuild
    CloseSpill
    mSideRow = mSideRow + 1
    WriteBand mSh.Range(mSh.Cells(mSideRow, COL_S_INPUT), mSh.Cells(mSideRow, COL_S_STATUS)), Caption
    mSideRow = mSideRow + 1
    If Not mDynamic Then
        mSh.Cells(mSideRow, COL_S_INPUT).Value = "This example needs dynamic-array Excel."
        mSh.Cells(mSideRow, COL_S_INPUT).Font.Italic = True
        mSideRow = mSideRow + 2
        Exit Sub
    End If

    WriteHeader mSh.Cells(mSideRow, COL_S_INPUT), Array(SourceHeader)
    mSpillCheckRow = mSideRow + 1
    With mSh.Cells(mSpillCheckRow, COL_S_INPUT)
        .Value = "Check"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    StyleRow mSh.Cells(mSpillCheckRow, COL_S_INPUT)
    Set mSpillAnchor = mSh.Cells(mSpillCheckRow + 1, COL_S_INPUT)
    mSpillAnchor.Resize(SourceRows, 1).NumberFormat = NumberFormat
    mSpillAnchor.Resize(SourceRows, 1).HorizontalAlignment = xlCenter
    SetFormula mSpillAnchor, SourceFormula
    mSpillCol = COL_S_RESULT
    mSpillOpen = True
    mSideRow = mSpillAnchor.Row + SourceRows + 1

End Sub

Public Sub Demo_SpillColumn( _
    ByVal Header As String, _
    ByVal Template As String, _
    ByVal Reference As String, _
    Optional ByVal NumberFormat As String = "General")
'
'==============================================================================
'                              Demo_SpillColumn
'------------------------------------------------------------------------------
' PURPOSE
'   One dependent spill of the open spill block. {SRC} in Template and
'   Reference stands for the whole source spill. The status cell above the
'   column is OK when the spill has the source's rows and equals Reference
'   element by element.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
    Dim Source  As String   'Spill reference of the source
    Dim Top     As Range    'First cell of this spill
    Dim Spill   As String   'Spill reference of this column

    RequireBuild
    If Not mSpillOpen Then Exit Sub
    If mSpillCol > COL_S_STATUS Then Err.Raise 5, "Demo_SpillColumn", "A spill block holds at most three columns."

    Source = mSpillAnchor.Address(False, False) & "#"
    Set Top = mSh.Cells(mSpillAnchor.Row, mSpillCol)
    Spill = Top.Address(False, False) & "#"
    WriteHeader mSh.Cells(mSpillAnchor.Row - 2, mSpillCol), Array(Header)
    Top.Resize(mSideRow - mSpillAnchor.Row - 1, 1).NumberFormat = NumberFormat
    Top.Resize(mSideRow - mSpillAnchor.Row - 1, 1).HorizontalAlignment = xlCenter
    SetFormula Top, Replace(Template, "{SRC}", Source)

    StyleRow mSh.Cells(mSpillCheckRow, mSpillCol)
    SetFormula mSh.Cells(mSpillCheckRow, mSpillCol), _
        "=IF(IFERROR(ERROR.TYPE(" & Top.Address(False, False) & ")=5,FALSE),""NO VBA""," & _
        "IFERROR(IF(AND(ROWS(" & Spill & ")=ROWS(" & Source & ")," & Spill & "=" & _
        Replace(Reference, "{SRC}", Source) & "),""OK"",""FAIL""),""FAIL""))"
    mSh.Cells(mSpillCheckRow, mSpillCol).HorizontalAlignment = xlCenter
    mStatus.Add mSh.Cells(mSpillCheckRow, mSpillCol).Address(False, False)
    mSpillCol = mSpillCol + 1

End Sub

Public Function Demo_Sheet() As Worksheet
'
'==============================================================================
'                                 Demo_Sheet
'------------------------------------------------------------------------------
' PURPOSE
'   The demo sheet, for a one-off layout no writer covers. Cells written this
'   way must stay below the last writer row so they never overlap a block.
'
' UPDATED
'   2026-10-06
'==============================================================================
'
    RequireBuild
    Set Demo_Sheet = mSh

End Function

'
'------------------------------------------------------------------------------
'
'                                   WRITERS
'
'------------------------------------------------------------------------------
'

Private Sub RequireBuild()
'
' Stops a writer that runs outside Demo_Begin ... Demo_Finish.
'
    If Not mBuilding Then Err.Raise 5, "KPR_Demo__Builder", "No demo build is in progress."

End Sub

Private Sub PrepareSheet()
'
' House font and column widths.
'
    With mSh.Cells
        .Font.Name = FONT_NAME
        .Font.Size = SIZE_BODY
        .VerticalAlignment = xlCenter
    End With
    If mNav Then
        mSh.Columns(1).ColumnWidth = NAV_COL_A
        mSh.Columns(1).Interior.Color = RGB(0, 0, 0)
        mSh.Columns(COL_TITLE).ColumnWidth = NAV_COL_B
    Else
        mSh.Columns(1).ColumnWidth = 2.6
        mSh.Columns(COL_TITLE).ColumnWidth = 2.6
    End If
    mSh.Columns(COL_FORMULA).ColumnWidth = 44
    mSh.Columns(COL_KPR).ColumnWidth = 15.6
    mSh.Columns(COL_EXCEL).ColumnWidth = 15.6
    mSh.Columns(COL_STATUS).ColumnWidth = 10
    mSh.Columns(COL_NOTES).ColumnWidth = 56
    mSh.Columns(COL_NOTES + 1).ColumnWidth = 2.6
    mSh.Columns(COL_S_INPUT).ColumnWidth = 15.6
    mSh.Columns(COL_S_RESULT).ColumnWidth = 15.6
    mSh.Columns(COL_S_EXPECT).ColumnWidth = 15.6
    mSh.Columns(COL_S_STATUS).ColumnWidth = 10
    mSh.Columns(COL_S_RULE).ColumnWidth = 48

End Sub

Private Sub WriteTitle(ByVal Title As String, ByVal Subtitle As String)
'
' Black title band with yellow text, navy subtitle band with white text.
'
    With mSh.Range(mSh.Cells(1, COL_TITLE), mSh.Cells(1, COL_LAST))
        .Interior.Color = RGB(0, 0, 0)
    End With
    With mSh.Cells(1, COL_TITLE)
        .Value = Title
        .Font.Bold = True
        .Font.Size = SIZE_TITLE
        .Font.Color = RGB(255, 192, 0)
    End With
    With mSh.Range(mSh.Cells(2, COL_TITLE), mSh.Cells(2, COL_LAST))
        .Interior.Color = ColourBand()
    End With
    With mSh.Cells(2, COL_TITLE)
        .Value = Subtitle
        .Font.Bold = True
        .Font.Size = SIZE_BAND
        .Font.Color = RGB(255, 255, 255)
    End With

End Sub

Private Sub WriteBand(ByVal Target As Range, ByVal Caption As String)
'
' Navy band with a white caption centred across Target.
'
    With Target
        .Interior.Color = ColourBand()
        .HorizontalAlignment = xlCenterAcrossSelection
        .Font.Bold = True
        .Font.Size = SIZE_BAND
        .Font.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = ColourBand()
    End With
    Target.Cells(1, 1).Value = Caption

End Sub

Private Sub WriteHeader(ByVal First As Range, ByVal Captions As Variant)
'
' Light-blue header cells to the right of First, one per caption.
'
    Dim K As Long   'Caption cursor

    For K = 0 To UBound(Captions)
        With First.Offset(0, K)
            .Value = Captions(K)
            .Font.Bold = True
            .Font.Size = SIZE_HEADER
            .Interior.Color = ColourHeader()
            .HorizontalAlignment = xlCenter
        End With
        StyleRow First.Offset(0, K)
    Next K

End Sub

Private Sub StyleRow(ByVal Target As Range)
'
' Thin grey borders on every edge of Target.
'
    With Target.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(166, 166, 166)
    End With

End Sub

Private Sub WriteFormulaText(ByVal Target As Range, ByVal Formula As String)
'
' Shows a formula as text: a text number format and a leading apostrophe, so
' Excel never evaluates it.
'
    Target.NumberFormat = "@"
    Target.Value = "'" & Formula
    Target.WrapText = True
    Target.Font.Name = "Consolas"

End Sub

Private Sub WriteStatus( _
    ByVal Target As Range, _
    ByVal Actual As Range, _
    ByVal Expected As Range, _
    ByVal DiffersWhen As String)
'
' The status formula of one row, registered for the summary.
'
    Dim A     As String     'Address of the live result
    Dim E     As String     'Address of the expected value
    Dim Same  As String     'Equality, error-aware
    Dim Other As String     'Outcome when the two differ

    A = Actual.Address(False, False)
    E = Expected.Address(False, False)
    Same = "IF(ISERROR(" & E & "),IFERROR(ERROR.TYPE(" & A & ")=ERROR.TYPE(" & E & "),FALSE)," & _
           "IFERROR(" & A & "=" & E & ",FALSE))"
    If Len(DiffersWhen) > 0 Then
        Other = "IF(IFERROR(" & DiffersWhen & ",FALSE),""" & STATUS_DIFFERS & """,""FAIL"")"
    Else
        Other = """FAIL"""
    End If
    SetFormula Target, "=IF(IFERROR(ERROR.TYPE(" & A & ")=5,FALSE),""NO VBA"",IF(" & Same & _
                       ",""" & STATUS_OK & """," & Other & "))"
    Target.HorizontalAlignment = xlCenter
    Target.Font.Bold = True
    mStatus.Add Target.Address(False, False)

End Sub

Private Sub CloseSpill()
'
' Ends the open spill block, if any.
'
    mSpillOpen = False
    Set mSpillAnchor = Nothing

End Sub

Private Sub WriteSummary()
'
' Counts of OK, DIFFERS, failing and all checks at the top of the side
' panel, each with a sheet-level name, and how to remove the demo.
'
    Dim R       As Long     'Summary row cursor
    Dim OkSum   As String   'Sum of OK statuses
    Dim DiffSum As String   'Sum of DIFFERS statuses
    Dim Item    As Variant  'Status address

    For Each Item In mStatus
        OkSum = OkSum & "+(" & CStr(Item) & "=""" & STATUS_OK & """)"
        DiffSum = DiffSum & "+(" & CStr(Item) & "=""" & STATUS_DIFFERS & """)"
    Next Item
    If Len(OkSum) = 0 Then
        OkSum = "+0"
        DiffSum = "+0"
    End If

    R = mSummaryRow
    WriteBand mSh.Range(mSh.Cells(R, COL_S_INPUT), mSh.Cells(R, COL_S_RULE)), "Checks"
    SummaryLine R + 1, "OK", "=" & Mid$(OkSum, 2), "Demo_OK", _
                "The KPR result equals the native-Excel reference."
    SummaryLine R + 2, "Differs by design", "=" & Mid$(DiffSum, 2), "Demo_Differs", _
                "A documented difference; the row's note explains it."
    SummaryLine R + 3, "Fail", "=" & CStr(mStatus.Count) & "-" & _
                ColumnLetter(COL_S_RESULT) & CStr(R + 1) & "-" & ColumnLetter(COL_S_RESULT) & CStr(R + 2), _
                "Demo_Fails", "Must be 0. NO VBA means the K-PRICING project is not loaded."
    SummaryLine R + 4, "Checks", "=" & CStr(mStatus.Count), "Demo_Checks", _
                "Every status cell on this sheet."
    With mSh.Cells(R + 5, COL_S_INPUT)
        .Value = "Built by K-PRICING. Ctrl+Z cannot remove it: delete this sheet instead, and run the demo again to rebuild it."
        .Font.Italic = True
    End With

End Sub

Private Sub SummaryLine( _
    ByVal R As Long, _
    ByVal Label As String, _
    ByVal Formula As String, _
    ByVal CellName As String, _
    ByVal Note As String)
'
' One summary line: label, count, note; the count cell gets CellName.
'
    StyleRow mSh.Range(mSh.Cells(R, COL_S_INPUT), mSh.Cells(R, COL_S_RULE))
    With mSh.Cells(R, COL_S_INPUT)
        .Value = Label
        .Font.Bold = True
        .Interior.Color = ColourHeader()
    End With
    With mSh.Cells(R, COL_S_RESULT)
        .Formula = Formula
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
    End With
    mSh.Cells(R, COL_S_EXPECT).Value = Note
    mSh.Names.Add Name:=CellName, RefersTo:="=" & SheetRef() & "!$" & ColumnLetter(COL_S_RESULT) & "$" & CStr(R)

End Sub

Private Sub AlignTitleRows()
'
' The two title bands have the same heights on every sheet of the workbook,
' as on the Home sheet, instead of the heights AutoFit gives their fonts.
'
    mSh.Rows(1).RowHeight = ROW_1_HEIGHT
    mSh.Rows(2).RowHeight = ROW_2_HEIGHT

End Sub

Private Sub DrawNavigation()
'
' Draws the navigation bar of the Home sheet: NavBarTop in the frozen rows,
' with the button of this sheet highlighted and not linked, and NavBar below
' them, with a button to Home. Purely cosmetic: a failure here removes what
' was drawn and never fails the build.
'
    If Not mNav Then Exit Sub
    On Error GoTo Failed
    mSh.Tab.Color = RGB(84, 130, 53)
    NavGroup NAV_GROUP_TOP, 0, "DEMO", mSh.Name, RGB(255, 192, 0), _
             "Each function next to Excel's own formula, with a status per row."
    NavGroup NAV_GROUP_REST, mSh.Rows(FROZEN_ROWS + 1).Top, "START", HOME_SHEET, RGB(143, 170, 220), _
             "What K-PRICING provides, how to use it and the commit it was built from."
    Exit Sub

Failed:
    Err.Clear
    RemoveNavigation

End Sub

Private Sub NavGroup( _
    ByVal GroupName As String, _
    ByVal BlockTop As Double, _
    ByVal Section As String, _
    ByVal Target As String, _
    ByVal Strip As Long, _
    ByVal Tip As String)
'
' One block of the bar: a section title and one button, grouped and named as
' in the template. The button of the sheet itself is blue and not linked.
'
    Dim Here   As Boolean   'The button is for this sheet
    Dim BarLeft  As Double  'Left of the bar, points
    Dim BarWidth As Double  'Width of the bar, points
    Dim BtnTop   As Double  'Top of the button, points
    Dim Head   As Shape     'Section title
    Dim Button As Shape     'Button
    Dim Mark   As Shape     'Strip at the left of the button
    Dim Whole  As Shape     'The group

    Here = (StrComp(Target, mSh.Name, vbBinaryCompare) = 0)
    BarLeft = NAV_BAR_X * 0.75
    BarWidth = NAV_BAR_W * 0.75
    BtnTop = BlockTop + (NAV_PAD_TOP + NAV_HEAD_H) * 0.75

    Set Head = NavShape("Nav_Section_" & StrConv(Section, vbProperCase), BarLeft, BlockTop + NAV_PAD_TOP * 0.75, BarWidth, NAV_HEAD_H * 0.75, _
                        -1, Section, RGB(255, 192, 0), True, NAV_HEAD_SIZE, MSO_ANCHOR_BASE, 4)
    If Here Then
        Set Button = NavShape("Nav_" & Target, BarLeft, BtnTop, BarWidth, NAV_BTN_H * 0.75, _
                              RGB(38, 90, 153), Target, RGB(255, 255, 255), True, NAV_TEXT_SIZE, MSO_ANCHOR_MID, 12)
    Else
        Set Button = NavShape("Nav_" & Target, BarLeft, BtnTop, BarWidth, NAV_BTN_H * 0.75, _
                              RGB(38, 38, 38), Target, RGB(217, 217, 217), False, NAV_TEXT_SIZE, MSO_ANCHOR_MID, 12)
        mSh.Hyperlinks.Add Anchor:=Button, Address:="", _
            SubAddress:="'" & Replace(Target, "'", "''") & "'!A1", ScreenTip:=Tip
    End If
    Set Mark = NavShape("Nav_" & Target & "_Mark", BarLeft, BtnTop, NAV_STRIP_W * 0.75, NAV_BTN_H * 0.75, _
                        IIf(Here, RGB(255, 192, 0), Strip), "", 0, False, 0, MSO_ANCHOR_MID, 0)

    Set Whole = mSh.Shapes.Range(Array(Head.Name, Button.Name, Mark.Name)).Group
    Whole.Name = GroupName
    Whole.Placement = xlMove

End Sub

Private Function NavShape( _
    ByVal ShapeName As String, _
    ByVal ShapeLeft As Double, _
    ByVal ShapeTop As Double, _
    ByVal ShapeWidth As Double, _
    ByVal ShapeHeight As Double, _
    ByVal Fill As Long, _
    ByVal Caption As String, _
    ByVal TextColour As Long, _
    ByVal Bold As Boolean, _
    ByVal Size As Double, _
    ByVal Anchor As Long, _
    ByVal Inset As Long) _
    As Shape
'
' A borderless rectangle; Fill is -1 for none. The text is left aligned and
' never wraps; Inset is its left margin in pixels.
'
    Dim Item As Shape   'The new shape

    Set Item = mSh.Shapes.AddShape(MSO_RECTANGLE, ShapeLeft, ShapeTop, ShapeWidth, ShapeHeight)
    Item.Name = ShapeName
    Item.Line.Visible = False
    If Fill < 0 Then
        Item.Fill.Visible = False
    Else
        Item.Fill.Visible = True
        Item.Fill.Solid
        Item.Fill.ForeColor.RGB = Fill
    End If
    If Len(Caption) > 0 Then
        With Item.TextFrame2
            .WordWrap = False
            .MarginLeft = Inset * 0.75
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = IIf(Anchor = MSO_ANCHOR_BASE, 1.5, 0)
            .VerticalAnchor = Anchor
            .TextRange.Text = Caption
            .TextRange.ParagraphFormat.Alignment = MSO_ALIGN_LEFT
            .TextRange.Font.Name = FONT_NAME
            .TextRange.Font.Size = Size
            .TextRange.Font.Bold = Bold
            .TextRange.Font.Fill.ForeColor.RGB = TextColour
        End With
    End If
    Set NavShape = Item

End Function

Private Sub RemoveNavigation()
'
' Deletes every shape of the navigation bar, whole or half drawn.
'
    Dim Index As Long   'Shape cursor, from the last

    On Error Resume Next
    For Index = mSh.Shapes.Count To 1 Step -1
        If Left$(mSh.Shapes(Index).Name, 3) = "Nav" Then mSh.Shapes(Index).Delete
    Next Index
    Err.Clear

End Sub

Private Function HasKprHome(ByVal Wb As Workbook) As Boolean
'
' TRUE when the workbook holds the K-PRICING Home sheet, the one that carries
' the navigation bar.
'
    On Error Resume Next
    HasKprHome = (StrComp(CStr(Wb.Worksheets(HOME_SHEET).Range("B1").Value), HOME_WORDMARK, vbBinaryCompare) = 0)
    If Err.Number <> 0 Then HasKprHome = False
    Err.Clear

End Function

Private Sub ShowSheet()
'
' Activates the demo sheet at its top-left corner without gridlines. Purely
' cosmetic: a failure here never fails the build.
'
    On Error Resume Next
    mWb.Activate
    mSh.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    If mNav Then
        ActiveWindow.FreezePanes = False
        ActiveWindow.SplitColumn = 1
        ActiveWindow.SplitRow = FROZEN_ROWS
        ActiveWindow.FreezePanes = True
    End If
    Err.Clear

End Sub

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub SetFormula(ByVal Target As Object, ByVal Formula As String)
'
' Writes Formula with the host qualifier. Formula2 keeps dynamic-array
' semantics where Excel provides it; earlier Excel falls back to Formula.
'
    Dim Qualified As String     'Formula as Excel must receive it

    If Len(mPrefix) > 0 Then
        Qualified = Replace(Formula, "KPR_", mPrefix & "KPR_")
    Else
        Qualified = Formula
    End If
    If mDynamic Then
        Target.Formula2 = Qualified
    Else
        Target.Formula = Qualified
    End If

End Sub

Private Function HasFormula2(ByVal Probe As Object) As Boolean
'
' TRUE when the running Excel provides Range.Formula2.
'
    Dim Text As String  'Probe result

    On Error Resume Next
    Text = Probe.Formula2
    HasFormula2 = (Err.Number = 0)
    Err.Clear

End Function

Private Function HostQualifier(ByVal Wb As Workbook) As String
'
' "" when this project is an add-in or holds the demo sheet itself;
' otherwise 'Workbook name'! so a formula reaches this project's functions.
'
    If ThisWorkbook.IsAddin Or Wb Is ThisWorkbook Then
        HostQualifier = ""
    Else
        HostQualifier = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!"
    End If

End Function

Private Function CanHostDemo(ByVal Wb As Workbook) As Boolean
'
' TRUE for an ordinary 1900-system workbook that accepts a new sheet.
'
    CanHostDemo = False
    If Wb Is Nothing Then Exit Function
    On Error GoTo Refuse
    If Wb.IsAddin Then Exit Function
    If Wb.ReadOnly Then Exit Function
    If Wb.ProtectStructure Then Exit Function
    If Wb.Date1904 Then Exit Function
    CanHostDemo = True
    Exit Function

Refuse:
    CanHostDemo = False

End Function

Private Function FreeSheetName(ByVal Wb As Workbook, ByVal Wanted As String) As String
'
' Wanted, or Wanted (2), (3) ... the first name no sheet of Wb uses.
'
    Dim Candidate As String     'Proposed name
    Dim Suffix    As String     'Numbered suffix
    Dim K         As Long       'Suffix counter

    Candidate = Left$(Wanted, SHEET_NAME_MAX)
    K = 1
    Do While SheetExists(Wb, Candidate)
        K = K + 1
        Suffix = " (" & CStr(K) & ")"
        Candidate = Left$(Wanted, SHEET_NAME_MAX - Len(Suffix)) & Suffix
    Loop
    FreeSheetName = Candidate

End Function

Private Function SheetExists(ByVal Wb As Workbook, ByVal SheetName As String) As Boolean
'
' TRUE when any sheet of Wb, chart sheets included, has SheetName.
'
    Dim Sh As Object    'Sheet cursor

    For Each Sh In Wb.Sheets
        If StrComp(Sh.Name, SheetName, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next Sh

End Function

Private Function SheetRef() As String
'
' The demo sheet's name, quoted for a reference.
'
    SheetRef = "'" & Replace(mSh.Name, "'", "''") & "'"

End Function

Private Function ColumnLetter(ByVal Column As Long) As String
'
' Column letter of a column number up to Z.
'
    ColumnLetter = Chr$(64 + Column)

End Function

Private Sub ApplyValueFormat(ByVal Target As Range, ByVal Value As Variant)
'
' Dates show as YYYY-MM-DD, text stays text, everything else is General.
'
    Select Case VarType(Value)
        Case vbDate
            Target.NumberFormat = DEMO_DATE
        Case vbString
            Target.NumberFormat = "@"
        Case Else
            Target.NumberFormat = "General"
    End Select

End Sub

Private Function ColourBand() As Long
    ColourBand = RGB(31, 56, 100)
End Function

Private Function ColourHeader() As Long
    ColourHeader = RGB(218, 233, 248)
End Function

Private Function ColourInput() As Long
    ColourInput = RGB(255, 204, 153)
End Function

Private Function ColourInputText() As Long
    ColourInputText = RGB(63, 63, 118)
End Function

Private Sub RestoreState(ByVal ReactivateCaller As Boolean)
'
' Puts the Excel state back and ends the build.
'
    On Error Resume Next
    If ReactivateCaller Then
        If Not mCallerBook Is Nothing Then mCallerBook.Activate
        If Not mCallerSheet Is Nothing Then mCallerSheet.Activate
    End If
    Application.Calculation = mOldCalculation
    Application.DisplayAlerts = mOldAlerts
    Application.ScreenUpdating = mOldScreen
    Application.EnableEvents = mOldEvents
    Set mCallerBook = Nothing
    Set mCallerSheet = Nothing
    Set mWb = Nothing
    Set mSh = Nothing
    Set mSpillAnchor = Nothing
    mSpillOpen = False
    mBuilding = False
    Err.Clear

End Sub

Private Sub DiscardBuild()
'
' Removes what the failed build created: its new workbook, or its sheet in
' the caller's workbook, and its temporary file. Nothing else is touched.
'
    On Error Resume Next
    Application.DisplayAlerts = False
    If mOwnBook Then
        If Not mWb Is Nothing Then mWb.Close SaveChanges:=False
    ElseIf Not mSh Is Nothing Then
        mSh.Delete
    End If
    If Len(mTempPath) > 0 Then Kill mTempPath
    mTempPath = ""
    Err.Clear
    RestoreState True

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
    Dim Cut As Long     'Position of the last path separator

    TryValidatePath = False
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
