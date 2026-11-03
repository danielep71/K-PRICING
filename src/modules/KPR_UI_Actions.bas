Attribute VB_Name = "KPR_UI_Actions"
'==============================================================================
' MODULE: KPR_UI_Actions
'------------------------------------------------------------------------------
' PURPOSE
'   The commands behind every K-PRICING menu: insert a supported function,
'   build a demo sheet, register or clear the Function Wizard help, and the
'   development-only regression runs. The Ribbon (KPR_UI_Ribbon) and the
'   classic toolbar (KPR_UI_CommandBars) both call these procedures, so the
'   two surfaces always offer the same commands with the same behavior.
'
' WHY THIS EXISTS
'   A menu callback must stay thin: no date algorithm, no layout and no
'   registration logic. Each command here only checks its preconditions,
'   delegates to the module that owns the work and tells the user the
'   outcome.
'
' SCOPE
'   Commands (each shows its outcome in a message box)
'   - KPR_UI_InsertFunction       insert one supported function
'   - KPR_UI_BuildDemo            build one catalog demo sheet
'   - KPR_UI_RegisterFunctions    register the Function Wizard help
'   - KPR_UI_ClearFunctions       clear the Function Wizard help
'   - KPR_UI_RunTests             development: pure regression suites
'   - KPR_UI_RunEvidence          development: durable evidence run
'   Menu content (no message, no side effect)
'   - KPR_UI_FunctionGroups       functional modules and their name prefixes
'   - KPR_UI_FunctionNames        supported functions, by name prefix
'   - KPR_UI_DemoGroups           demo catalog groups
'   - KPR_UI_DemoEntries          demo catalog entries of one group
'   - KPR_UI_DemosAvailable       the demo modules are loaded
'   - KPR_UI_TestsAvailable       the regression modules are loaded
'   - KPR_UI_ActionKnown          a menu action text names a command
'   - KPR_UI_RunAction            run a menu action text
'
' MENU ACTIONS
'   A menu item names its command with one text, "<kind>:<argument>", that
'   both surfaces store on the control (CommandBar Parameter, Ribbon tag):
'     fn:<function name>        KPR_UI_InsertFunction
'     demo:<entry point>        KPR_UI_BuildDemo
'     setup:register            KPR_UI_RegisterFunctions
'     setup:clear               KPR_UI_ClearFunctions
'     dev:tests                 KPR_UI_RunTests
'     dev:evidence              KPR_UI_RunEvidence
'
' OPTIONAL MODULES
'   The demo modules (examples/) and the regression modules (tests/) are not
'   part of every build. This module never names their procedures in code: it
'   reaches them with Application.Run and the workbook-qualified name, so the
'   project compiles without them. A menu shows demo commands only when the
'   demo catalog answers, and development commands only when
'   KPR_Tests_Available answers.
'
' VISIBILITY
'   Option Private Module. Callable from this project, including the menu
'   modules, and through Application.Run. Not supported API: menus are
'   unsupported infrastructure (docs/PUBLIC_API.txt lists the API).
'
' ERROR POLICY
'   No command raises to the menu that called it. A failure is reported in a
'   message box, and Excel's alerts, events, status bar, active workbook,
'   active sheet and selection are restored to what the command found.
'
' ALLOWED DEPENDENCIES
'   KPR_REGISTER_PUBLIC_UDFS for the function manifest and registration.
'   Demo and regression procedures by name through Application.Run only.
'
' UPDATED
'   2026-11-03
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit         'Force explicit variable declarations
    Option Private Module   'UI infrastructure: invisible outside this VBA project

'------------------------------------------------------------------------------
' CONSTANTS
'------------------------------------------------------------------------------
    'Title of every message box
        Private Const UI_TITLE As String = "K-PRICING"
    'Separator of a menu action text and of a demo catalog entry
        Private Const ACTION_SEP  As String = ":"
        Private Const CATALOG_SEP As String = "|"

'------------------------------------------------------------------------------
' TYPES
'------------------------------------------------------------------------------
    'Excel state a command restores when it fails
    Private Type UIState
        Alerts          As Boolean      'Application.DisplayAlerts
        Events          As Boolean      'Application.EnableEvents
        StatusBar       As Variant      'Application.StatusBar: FALSE or a text
        ShowStatus      As Boolean      'Application.DisplayStatusBar
        Book            As Workbook     'Active workbook, or Nothing
        Sheet           As Object       'Active sheet, or Nothing
        Selected        As Object       'Selection, or Nothing
    End Type

'
'------------------------------------------------------------------------------
'
'                                   COMMANDS
'
'------------------------------------------------------------------------------
'

Public Sub KPR_UI_InsertFunction( _
    ByVal FunctionName As String)
'
'==============================================================================
'                            KPR_UI_InsertFunction
'------------------------------------------------------------------------------
' PURPOSE
'   Puts one supported function into the active cell and opens Excel's
'   Function Arguments dialog for it.
'
' INPUTS
'   FunctionName
'     A name from the registration manifest, for example KPR_Dates_EndOfMonth.
'
' BEHAVIOR
'   A function without arguments is entered directly. Otherwise the cell gets
'   "=<name>()" and the dialog opens; when the user cancels, the cell gets
'   back the content it had before. When this project is not an add-in and
'   the cell is in another workbook, the name is qualified with this
'   workbook's name, as Excel requires.
'
' ERROR POLICY
'   Refuses, with a message, a name outside the manifest, a selection that is
'   not a cell and a locked cell on a protected sheet.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim State           As UIState      'Excel state on entry
    Dim Index           As Long         'Manifest record of FunctionName
    Dim Target          As Range        'Cell that receives the formula
    Dim Inserted        As String       'Formula this command wrote
    Dim OldContent      As String       'Formula or value of Target on entry
    Dim OldDynamic      As Boolean      'OldContent was read through Formula2

'------------------------------------------------------------------------------
' GUARD ENTRY
'------------------------------------------------------------------------------
    'Refuse before touching a cell: a wrong target must not lose content
        CaptureState State
        On Error GoTo Fail
        Index = ManifestIndex(FunctionName)
        If Index = 0 Then
            Notify "Not a K-PRICING function: " & FunctionName, vbExclamation
            Exit Sub
        End If
        If TypeName(Selection) <> "Range" Then
            Notify "Select a worksheet cell first, then choose the function again.", vbExclamation
            Exit Sub
        End If
        Set Target = ActiveCell
        If Target.Worksheet.ProtectContents And Target.Locked Then
            Notify "The active cell is locked on a protected sheet.", vbExclamation
            Exit Sub
        End If

'------------------------------------------------------------------------------
' INSERT
'------------------------------------------------------------------------------
    'Keep the old content so a cancelled dialog leaves the cell as it was
        OldDynamic = ReadFormula(Target, OldContent)
        Inserted = "=" & FormulaQualifier(Target.Worksheet.Parent) & FunctionName & "()"
        Target.Formula = Inserted
        If KPR_Register_ManifestArgCount(Index) = 0 Then Exit Sub

    'Compare with Excel's own spelling of the inserted formula: a cancel leaves
    'it unchanged, while OK in Function Arguments adds the arguments
        Inserted = CStr(Target.Formula)
        Target.FunctionWizard
        If StrComp(CStr(Target.Formula), Inserted, vbTextCompare) = 0 Then
            WriteFormula Target, OldContent, OldDynamic
        End If
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE FAILURE
'------------------------------------------------------------------------------
Fail:
    'Put the cell back before reporting, then restore Excel's state
        Notify "Could not insert " & FunctionName & ": error " & CStr(Err.Number) & _
               " (" & Err.Description & ")", vbExclamation
        On Error Resume Next
        If Not Target Is Nothing And Len(Inserted) > 0 Then WriteFormula Target, OldContent, OldDynamic
        RestoreState State

End Sub


Public Sub KPR_UI_BuildDemo( _
    ByVal EntryPoint As String)
'
'==============================================================================
'                               KPR_UI_BuildDemo
'------------------------------------------------------------------------------
' PURPOSE
'   Builds one demo sheet of the demo catalog into the active workbook.
'
' INPUTS
'   EntryPoint
'     An entry point listed by the demo catalog, for example
'     KPR_Demo_BuildDates.
'
' BEHAVIOR
'   Delegates to KPR_Demo_Build, which owns where the sheet goes, its layout
'   and the Excel state around the build. On success the new sheet is left
'   active. On failure the message gives the builder's report.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim State           As UIState      'Excel state on entry
    Dim Built           As Boolean      'Result of KPR_Demo_Build

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
    'The builder reports its own failures; Run raises only when it is absent
        CaptureState State
        If Not KPR_UI_DemosAvailable() Then
            Notify "The demo modules are not loaded in this build.", vbExclamation
            Exit Sub
        End If
        On Error GoTo Fail
        Built = CBool(Application.Run(Qualified("KPR_Demo_Build"), EntryPoint))
        If Not Built Then
            Notify "The demo was not built. " & CStr(Application.Run(Qualified("KPR_Demo_LastReport"))), _
                   vbExclamation
            RestoreState State
        End If
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE FAILURE
'------------------------------------------------------------------------------
Fail:
    'An error escaping the builder still restores the caller's state
        Notify "Could not build the demo " & EntryPoint & ": error " & CStr(Err.Number) & _
               " (" & Err.Description & ")", vbExclamation
        On Error Resume Next
        RestoreState State

End Sub


Public Sub KPR_UI_RegisterFunctions()
'
'==============================================================================
'                          KPR_UI_RegisterFunctions
'------------------------------------------------------------------------------
' PURPOSE
'   Registers the Function Wizard help of every supported function and says
'   how many Excel accepted.
'
' UPDATED
'   2026-11-03
'==============================================================================
'
    ReportRegistration KPR_Register_PublicUDFs(), "registered in the category ""KPR Dates"""

End Sub


Public Sub KPR_UI_ClearFunctions()
'
'==============================================================================
'                            KPR_UI_ClearFunctions
'------------------------------------------------------------------------------
' PURPOSE
'   Clears the Function Wizard help of every supported function and says how
'   many Excel accepted. The functions stay callable.
'
' UPDATED
'   2026-11-03
'==============================================================================
'
    ReportRegistration KPR_Register_ClearPublicUDFs(), "moved back to ""User Defined"""

End Sub


Public Sub KPR_UI_RunTests()
'
'==============================================================================
'                               KPR_UI_RunTests
'------------------------------------------------------------------------------
' PURPOSE
'   Development only: runs the pure regression suites (KPR_Tests_Run). The
'   report goes to the Immediate window of the VBA editor.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim State           As UIState      'Excel state on entry

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'The runner restores its own state; this guards only its absence
        CaptureState State
        If Not KPR_UI_TestsAvailable() Then
            Notify "The regression modules are not loaded in this build.", vbExclamation
            Exit Sub
        End If
        On Error GoTo Fail
        Application.Run Qualified("KPR_Tests_Run")
        Notify "The pure regression suites have run. The report is in the Immediate window " & _
               "of the VBA editor (Alt+F11, then Ctrl+G).", vbInformation
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE FAILURE
'------------------------------------------------------------------------------
Fail:
    'Report and restore; a runner failure never reaches the menu
        Notify "The regression run stopped: error " & CStr(Err.Number) & _
               " (" & Err.Description & ")", vbExclamation
        On Error Resume Next
        RestoreState State

End Sub


Public Sub KPR_UI_RunEvidence()
'
'==============================================================================
'                              KPR_UI_RunEvidence
'------------------------------------------------------------------------------
' PURPOSE
'   Development only: asks for the exact source SHA and an output folder,
'   then runs every durable suite (KPR_Test_RunAll) and says where the
'   evidence file is.
'
' INPUTS
'   Two input boxes: the 40-character commit SHA of the loaded source and an
'   existing folder. Cancelling either stops without running anything.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim State           As UIState      'Excel state on entry
    Dim Sha             As String       'Exact source SHA of the loaded modules
    Dim Folder          As String       'Folder that receives the evidence
    Dim Passed          As Boolean      'Result of KPR_Test_RunAll

'------------------------------------------------------------------------------
' ASK
'------------------------------------------------------------------------------
    'Validate both answers before the run: the evidence is bound to the SHA
        CaptureState State
        If Not KPR_UI_TestsAvailable() Then
            Notify "The regression modules are not loaded in this build.", vbExclamation
            Exit Sub
        End If
        Sha = LCase$(Trim$(InputBox("Exact 40-character commit SHA of the loaded source:", UI_TITLE)))
        If Len(Sha) = 0 Then Exit Sub
        If Not IsSha(Sha) Then
            Notify "Not a 40-character commit SHA: " & Sha, vbExclamation
            Exit Sub
        End If
        Folder = Trim$(InputBox("Existing folder for the evidence file:", UI_TITLE, CurDir$))
        If Len(Folder) = 0 Then Exit Sub
        If Len(Dir$(Folder, vbDirectory)) = 0 Then
            Notify "The folder does not exist: " & Folder, vbExclamation
            Exit Sub
        End If

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'The runner writes the evidence and restores its own state
        On Error GoTo Fail
        Passed = CBool(Application.Run(Qualified("KPR_Test_RunAll"), Sha, Folder))
        If Passed Then
            Notify "Every durable suite passed. The evidence file is in " & Folder & ".", vbInformation
        Else
            Notify "The durable run found failures or could not complete. " & _
                   "See the evidence file in " & Folder & " and the Immediate window.", vbExclamation
        End If
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE FAILURE
'------------------------------------------------------------------------------
Fail:
    'Report and restore; a runner failure never reaches the menu
        Notify "The durable run stopped: error " & CStr(Err.Number) & _
               " (" & Err.Description & ")", vbExclamation
        On Error Resume Next
        RestoreState State

End Sub

'
'------------------------------------------------------------------------------
'
'                                 MENU CONTENT
'
'------------------------------------------------------------------------------
'

Public Function KPR_UI_FunctionGroups() As Variant
'
'==============================================================================
'                            KPR_UI_FunctionGroups
'------------------------------------------------------------------------------
' PURPOSE
'   The functional modules a Functions menu shows, one "Caption|Prefix" text
'   each: the submenu caption and the name prefix of its functions. A new
'   functional module adds one line here and, for the Ribbon, one group in
'   src/ribbon/customUI14.xml.
'
' RETURNS
'   A zero-based array of "Caption|Prefix" texts.
'
' UPDATED
'   2026-11-03
'==============================================================================
'
    KPR_UI_FunctionGroups = Array( _
        "Date primitives|KPR_Dates_")

End Function


Public Function KPR_UI_FunctionNames( _
    ByVal Prefix As String) _
    As Variant
'
'==============================================================================
'                             KPR_UI_FunctionNames
'------------------------------------------------------------------------------
' PURPOSE
'   The supported functions whose names start with Prefix, in manifest order,
'   for a Functions menu.
'
' INPUTS
'   Prefix
'     A name prefix such as "KPR_Dates_"; "" lists every function.
'
' RETURNS
'   A zero-based array of names, or Empty when none matches.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Names()         As String       'Matching names
    Dim Found           As Long         'Names collected
    Dim Count           As Long         'Manifest records
    Dim I               As Long         'Manifest cursor
    Dim FnName          As String       'Name of record I

'------------------------------------------------------------------------------
' COLLECT
'------------------------------------------------------------------------------
    'The manifest is the single list of supported functions
        Count = KPR_Register_ManifestCount()
        If Count = 0 Then Exit Function
        ReDim Names(0 To Count - 1)
        For I = 1 To Count
            FnName = KPR_Register_ManifestName(I)
            If StrComp(Left$(FnName, Len(Prefix)), Prefix, vbTextCompare) = 0 Then
                Names(Found) = FnName
                Found = Found + 1
            End If
        Next I
        If Found = 0 Then Exit Function
        ReDim Preserve Names(0 To Found - 1)
        KPR_UI_FunctionNames = Names

End Function


Public Function KPR_UI_DemoGroups() As Variant
'
'==============================================================================
'                              KPR_UI_DemoGroups
'------------------------------------------------------------------------------
' PURPOSE
'   The distinct groups of the demo catalog, in catalog order, for a menu
'   with one submenu per functional module.
'
' RETURNS
'   A zero-based array of group names, or Empty when the demo modules are not
'   loaded or the catalog is empty.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Catalog         As Variant      'KPR_Demo_Catalog result
    Dim Groups()        As String       'Distinct groups found
    Dim Found           As Long         'Groups collected
    Dim Entry           As Variant      'Catalog cursor
    Dim Group           As String       'Group of Entry
    Dim J               As Long         'Groups cursor
    Dim Seen            As Boolean      'Group was already collected

'------------------------------------------------------------------------------
' COLLECT
'------------------------------------------------------------------------------
    'Keep the first occurrence of each group so menus follow catalog order
        Catalog = DemoCatalog()
        If IsEmpty(Catalog) Then Exit Function
        ReDim Groups(0 To UBound(Catalog) - LBound(Catalog))
        For Each Entry In Catalog
            Group = CatalogField(CStr(Entry), 2)
            Seen = False
            For J = 0 To Found - 1
                If StrComp(Groups(J), Group, vbTextCompare) = 0 Then Seen = True
            Next J
            If Not Seen And Len(Group) > 0 Then
                Groups(Found) = Group
                Found = Found + 1
            End If
        Next Entry
        If Found = 0 Then Exit Function
        ReDim Preserve Groups(0 To Found - 1)
        KPR_UI_DemoGroups = Groups

End Function


Public Function KPR_UI_DemoEntries( _
    ByVal Group As String) _
    As Variant
'
'==============================================================================
'                              KPR_UI_DemoEntries
'------------------------------------------------------------------------------
' PURPOSE
'   The demo catalog entries of one group, for a Demo sheets menu.
'
' INPUTS
'   Group
'     A catalog group such as "Date primitives"; "" lists every entry.
'
' RETURNS
'   A zero-based array of "EntryPoint|Caption" texts, or Empty when the demo
'   modules are not loaded or the group has no entry.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Catalog         As Variant      'KPR_Demo_Catalog result
    Dim Entries()       As String       'Matching entries
    Dim Found           As Long         'Entries collected
    Dim Entry           As Variant      'Catalog cursor

'------------------------------------------------------------------------------
' COLLECT
'------------------------------------------------------------------------------
    'An entry without a group belongs to no submenu but still to ""
        Catalog = DemoCatalog()
        If IsEmpty(Catalog) Then Exit Function
        ReDim Entries(0 To UBound(Catalog) - LBound(Catalog))
        For Each Entry In Catalog
            If Len(Group) = 0 Or StrComp(CatalogField(CStr(Entry), 2), Group, vbTextCompare) = 0 Then
                Entries(Found) = CatalogField(CStr(Entry), 0) & CATALOG_SEP & CatalogField(CStr(Entry), 1)
                Found = Found + 1
            End If
        Next Entry
        If Found = 0 Then Exit Function
        ReDim Preserve Entries(0 To Found - 1)
        KPR_UI_DemoEntries = Entries

End Function


Public Function KPR_UI_DemosAvailable() As Boolean
'
' TRUE when the demo catalog answers, that is when the demo modules are loaded.
'
    KPR_UI_DemosAvailable = Not IsEmpty(DemoCatalog())

End Function


Public Function KPR_UI_TestsAvailable() As Boolean
'
' TRUE when the regression modules are loaded: KPR_Tests_Available answers.
'
    On Error GoTo Absent
    KPR_UI_TestsAvailable = CBool(Application.Run(Qualified("KPR_Tests_Available")))
    Exit Function

Absent:
    KPR_UI_TestsAvailable = False

End Function


Public Function KPR_UI_ActionKnown( _
    ByVal Action As String) _
    As Boolean
'
'==============================================================================
'                              KPR_UI_ActionKnown
'------------------------------------------------------------------------------
' PURPOSE
'   TRUE when a menu action text names a command KPR_UI_RunAction can run,
'   with an argument the command accepts. Runs nothing, so the regression
'   suite can check every menu item without opening a dialog.
'
' INPUTS
'   Action
'     A menu action text, "<kind>:<argument>" (see MENU ACTIONS).
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Kind            As String       'Text before the first ":"
    Dim Argument        As String       'Text after the first ":"
    Dim Entry           As Variant      'Catalog cursor

'------------------------------------------------------------------------------
' CLASSIFY
'------------------------------------------------------------------------------
    'Every argument is checked against the list its command reads
        KPR_UI_ActionKnown = False
        If Not SplitAction(Action, Kind, Argument) Then Exit Function
        Select Case Kind
            Case "fn"
                KPR_UI_ActionKnown = (ManifestIndex(Argument) > 0)
            Case "demo"
                If Not KPR_UI_DemosAvailable() Then Exit Function
                For Each Entry In KPR_UI_DemoEntries("")
                    If StrComp(CatalogField(CStr(Entry), 0), Argument, vbTextCompare) = 0 Then
                        KPR_UI_ActionKnown = True
                    End If
                Next Entry
            Case "setup"
                KPR_UI_ActionKnown = (Argument = "register" Or Argument = "clear")
            Case "dev"
                KPR_UI_ActionKnown = (Argument = "tests" Or Argument = "evidence")
        End Select

End Function


Public Sub KPR_UI_RunAction( _
    ByVal Action As String)
'
'==============================================================================
'                               KPR_UI_RunAction
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the command a menu action text names. Both menu surfaces call this
'   with the text stored on the clicked control.
'
' INPUTS
'   Action
'     A menu action text, "<kind>:<argument>" (see MENU ACTIONS).
'
' ERROR POLICY
'   An unknown text is reported in a message box; nothing runs.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Kind            As String       'Text before the first ":"
    Dim Argument        As String       'Text after the first ":"

'------------------------------------------------------------------------------
' DISPATCH
'------------------------------------------------------------------------------
    'Each command validates its own argument again before acting
        If Not SplitAction(Action, Kind, Argument) Then Kind = ""
        Select Case Kind
            Case "fn":      KPR_UI_InsertFunction Argument
            Case "demo":    KPR_UI_BuildDemo Argument
            Case "setup"
                Select Case Argument
                    Case "register":    KPR_UI_RegisterFunctions
                    Case "clear":       KPR_UI_ClearFunctions
                    Case Else:          Notify "Unknown K-PRICING command: " & Action, vbExclamation
                End Select
            Case "dev"
                Select Case Argument
                    Case "tests":       KPR_UI_RunTests
                    Case "evidence":    KPR_UI_RunEvidence
                    Case Else:          Notify "Unknown K-PRICING command: " & Action, vbExclamation
                End Select
            Case Else
                Notify "Unknown K-PRICING command: " & Action, vbExclamation
        End Select

End Sub

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function Qualified( _
    ByVal ProcedureName As String) _
    As String
'
' The workbook-qualified name Application.Run needs to reach a procedure of
' this project, including one in an Option Private Module module.
'
    Qualified = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & ProcedureName

End Function


Private Function DemoCatalog() As Variant
'
' KPR_Demo_Catalog, or Empty when the demo modules are not loaded or the
' catalog is not a non-empty array.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Catalog         As Variant      'KPR_Demo_Catalog result

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Run raises when the module is absent; that is the "not loaded" answer
        On Error GoTo Absent
        Catalog = Application.Run(Qualified("KPR_Demo_Catalog"))
        If Not IsArray(Catalog) Then Exit Function
        If UBound(Catalog) < LBound(Catalog) Then Exit Function
        DemoCatalog = Catalog
        Exit Function

Absent:
    'An absent or failing catalog shows as no demos
        DemoCatalog = Empty

End Function


Private Function CatalogField( _
    ByVal Entry As String, _
    ByVal Index As Long) _
    As String
'
' Field Index (0 entry point, 1 caption, 2 group) of an
' "EntryPoint|Caption|Group" catalog text; "" when the entry has no such field.
'
    Dim Fields          As Variant      'Split fields of Entry

    Fields = Split(Entry, CATALOG_SEP)
    If Index <= UBound(Fields) Then CatalogField = Trim$(CStr(Fields(Index)))

End Function


Private Function SplitAction( _
    ByVal Action As String, _
    ByRef Kind As String, _
    ByRef Argument As String) _
    As Boolean
'
' Splits "<kind>:<argument>" at the first ":". FALSE when either part is empty.
'
    Dim At              As Long         'Position of the separator

    At = InStr(1, Action, ACTION_SEP)
    If At < 2 Or At = Len(Action) Then Exit Function
    Kind = LCase$(Left$(Action, At - 1))
    Argument = Mid$(Action, At + 1)
    SplitAction = True

End Function


Private Function ManifestIndex( _
    ByVal FunctionName As String) _
    As Long
'
' 1-based manifest record of FunctionName; 0 when it is not a supported function.
'
    Dim I               As Long         'Manifest cursor

    For I = 1 To KPR_Register_ManifestCount()
        If StrComp(KPR_Register_ManifestName(I), FunctionName, vbTextCompare) = 0 Then
            ManifestIndex = I
            Exit Function
        End If
    Next I

End Function


Private Function FormulaQualifier( _
    ByVal Wb As Workbook) _
    As String
'
' "" when this project is an add-in or holds the cell itself; otherwise
' 'Workbook name'! so the formula reaches this project's functions.
'
    If ThisWorkbook.IsAddin Or Wb Is ThisWorkbook Then
        FormulaQualifier = ""
    Else
        FormulaQualifier = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!"
    End If

End Function


Private Function ReadFormula( _
    ByVal Cell As Object, _
    ByRef Content As String) _
    As Boolean
'
' Reads the cell's formula or constant into Content, through Formula2 when
' the running Excel provides it (so a dynamic-array formula keeps its form).
' Returns TRUE when Formula2 was used.
'
    On Error Resume Next
    Content = CStr(Cell.Formula2)
    If Err.Number = 0 Then
        ReadFormula = True
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0
    Content = CStr(Cell.Formula)

End Function


Private Sub WriteFormula( _
    ByVal Cell As Object, _
    ByVal Content As String, _
    ByVal Dynamic As Boolean)
'
' Writes back what ReadFormula read, through the same property.
'
    If Dynamic Then
        Cell.Formula2 = Content
    Else
        Cell.Formula = Content
    End If

End Sub


Private Sub ReportRegistration( _
    ByVal Accepted As Long, _
    ByVal Outcome As String)
'
' Tells the user how many manifest functions Excel accepted, with the
' registration module's per-function report when some were refused.
'
    Dim Total           As Long         'Manifest records

    Total = KPR_Register_ManifestCount()
    If Accepted = Total Then
        Notify CStr(Accepted) & " K-PRICING functions " & Outcome & ".", vbInformation
    Else
        Notify CStr(Accepted) & " of " & CStr(Total) & " K-PRICING functions " & Outcome & "." & _
               vbCrLf & vbCrLf & KPR_Register_LastReport(), vbExclamation
    End If

End Sub


Private Function IsSha( _
    ByVal Text As String) _
    As Boolean
'
' TRUE for exactly 40 lowercase hexadecimal characters.
'
    Dim I               As Long         'Character cursor

    If Len(Text) <> 40 Then Exit Function
    For I = 1 To 40
        If InStr(1, "0123456789abcdef", Mid$(Text, I, 1), vbBinaryCompare) = 0 Then Exit Function
    Next I
    IsSha = True

End Function


Private Sub Notify( _
    ByVal Message As String, _
    ByVal Style As VbMsgBoxStyle)
'
' The one message box of every command.
'
    MsgBox Message, Style, UI_TITLE

End Sub


Private Sub CaptureState( _
    ByRef State As UIState)
'
' Records the Excel state a failed command restores. Never raises.
'
    On Error Resume Next
    State.Alerts = Application.DisplayAlerts
    State.Events = Application.EnableEvents
    State.StatusBar = Application.StatusBar
    State.ShowStatus = Application.DisplayStatusBar
    Set State.Book = ActiveWorkbook
    Set State.Sheet = ActiveSheet
    Set State.Selected = Selection
    Err.Clear

End Sub


Private Sub RestoreState( _
    ByRef State As UIState)
'
' Restores what CaptureState recorded, each part independently, so one part
' that can no longer be restored (a closed workbook) does not stop the others.
' Never raises.
'
    On Error Resume Next
    Application.DisplayAlerts = State.Alerts
    Application.EnableEvents = State.Events
    Application.DisplayStatusBar = State.ShowStatus
    If VarType(State.StatusBar) = vbBoolean Then
        Application.StatusBar = False
    Else
        Application.StatusBar = State.StatusBar
    End If
    If Not State.Book Is Nothing Then State.Book.Activate
    If Not State.Sheet Is Nothing Then State.Sheet.Activate
    If Not State.Selected Is Nothing Then State.Selected.Select
    Err.Clear

End Sub
