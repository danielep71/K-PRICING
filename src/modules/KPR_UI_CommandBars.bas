Attribute VB_Name = "KPR_UI_CommandBars"
'==============================================================================
' MODULE: KPR_UI_CommandBars
'------------------------------------------------------------------------------
' PURPOSE
'   The classic K-PRICING toolbar, "KPR Dates", for users and environments
'   that rely on CommandBars rather than the Ribbon. In Excel 2007 and later
'   it appears on the Ribbon's Add-ins tab.
'
' SCOPE
'   - KPR_UI_InstallCommandBar        build the toolbar (repeatable)
'   - KPR_UI_RemoveCommandBar         remove every KPR-owned bar and control
'   - KPR_UI_CommandBarClick          OnAction target of every button
'   - KPR_UI_CommandBarLastReport     why the last install failed
'   - KPR_UI_CommandBarControlCount   KPR-owned bars and controls, for tests
'   - KPR_UI_CommandBarActions        action text of every button, for tests
'
' TOOLBAR
'   Functions       one submenu per functional module (KPR_UI_FunctionGroups),
'                   one button per supported function
'   Demo sheets     one submenu per demo catalog group; a disabled note when
'                   the demo modules are not loaded
'   Setup           register / clear the Function Wizard help
'   Development     shown only when the regression modules are loaded
'   Every button stores its menu action text ("fn:KPR_Dates_EndOfMonth",
'   "demo:KPR_Demo_BuildDates", ...) in its Parameter and runs
'   KPR_UI_CommandBarClick, which hands that text to KPR_UI_RunAction.
'
' LIFECYCLE
'   Install first removes every KPR-owned bar and control, then builds the
'   toolbar again, so a second install never duplicates anything. Remove
'   deletes the bar named "KPR Dates" and every control tagged "KPR_UI"
'   anywhere, and tolerates their absence; both are safe to call repeatedly
'   and in any order. The bar and its controls are Temporary: Excel discards
'   them when it closes, so they never persist into the next session.
'
' INTEGRATION
'   Call KPR_UI_InstallCommandBar from Workbook_Open (or the add-in's
'   startup) and KPR_UI_RemoveCommandBar from Workbook_BeforeClose. Both are
'   also in the macro list (Alt+F8).
'
' VISIBILITY
'   No Option Private Module: the install and remove entry points are meant
'   for the macro list, and every button reaches KPR_UI_CommandBarClick by
'   its workbook-qualified name. Unsupported infrastructure, not supported
'   API (docs/PUBLIC_API.txt lists the API).
'
' ERROR POLICY
'   No entry point raises. A failed install removes what it built and
'   returns FALSE; KPR_UI_CommandBarLastReport says why.
'
' ALLOWED DEPENDENCIES
'   KPR_UI_Actions only. Office CommandBar objects are late bound.
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

'------------------------------------------------------------------------------
' CONSTANTS
'------------------------------------------------------------------------------
    'Name of the toolbar and tag of every KPR-owned control
        Private Const BAR_NAME As String = "KPR Dates"
        Private Const UI_TAG   As String = "KPR_UI"
    'Office constants, spelled out because the objects are late bound
        Private Const MSO_BAR_TOP        As Long = 1    'msoBarTop
        Private Const MSO_CONTROL_BUTTON As Long = 1    'msoControlButton
        Private Const MSO_CONTROL_POPUP  As Long = 10   'msoControlPopup
        Private Const MSO_BUTTON_CAPTION As Long = 2    'msoButtonCaption
    'Upper bound of the removal loops: a control that refuses deletion must
    'not hold teardown in an endless loop
        Private Const MAX_REMOVALS As Long = 500

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    Private mLastReport     As String       'Why the last install failed, or ""

'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Function KPR_UI_InstallCommandBar() As Boolean
'
'==============================================================================
'                           KPR_UI_InstallCommandBar
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the "KPR Dates" toolbar and shows it.
'
' RETURNS
'   TRUE when the toolbar was built; FALSE when it failed, with nothing of it
'   left behind and the reason in KPR_UI_CommandBarLastReport.
'
' BEHAVIOR
'   Repeatable: every KPR-owned bar and control is removed first, so a second
'   install leaves exactly one toolbar. Menu content is read at install time;
'   install again after loading or removing the demo or regression modules.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Bar             As Object       'The new CommandBar
    Dim Stage           As String       'Step in progress, for the report

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
    'Clear stale KPR controls first: that is what makes install idempotent
        mLastReport = vbNullString
        On Error GoTo Fail
        Stage = "remove the previous toolbar"
        KPR_UI_RemoveCommandBar
        Stage = "create the toolbar"
        Set Bar = Application.CommandBars.Add(Name:=BAR_NAME, Position:=MSO_BAR_TOP, Temporary:=True)
        Stage = "add the Functions menu"
        AddFunctionsMenu Bar
        Stage = "add the Demo sheets menu"
        AddDemoMenu Bar
        Stage = "add the Setup menu"
        AddSetupMenu Bar
        If KPR_UI_TestsAvailable() Then
            Stage = "add the Development menu"
            AddDevelopmentMenu Bar
        End If
        Stage = "show the toolbar"
        Bar.Visible = True
        KPR_UI_InstallCommandBar = True
        Exit Function

'------------------------------------------------------------------------------
' HANDLE FAILURE
'------------------------------------------------------------------------------
Fail:
    'Leave no half-built toolbar: remove everything KPR-owned again
        mLastReport = "Could not " & Stage & ": error " & CStr(Err.Number) & " (" & Err.Description & ")"
        On Error Resume Next
        KPR_UI_RemoveCommandBar
        KPR_UI_InstallCommandBar = False

End Function


Public Function KPR_UI_RemoveCommandBar() As Boolean
'
'==============================================================================
'                           KPR_UI_RemoveCommandBar
'------------------------------------------------------------------------------
' PURPOSE
'   Removes the "KPR Dates" toolbar and every control tagged "KPR_UI", and
'   nothing else.
'
' RETURNS
'   TRUE when no KPR-owned bar or control remains.
'
' BEHAVIOR
'   Repeatable and tolerant: absent UI is not an error. Other toolbars and
'   controls are never touched.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Bar             As Object       'A KPR-owned CommandBar
    Dim Ctl             As Object       'A KPR-tagged control outside the bar
    Dim Guard           As Long         'Loop bound

'------------------------------------------------------------------------------
' REMOVE
'------------------------------------------------------------------------------
    'The bar first (its controls go with it), then strays elsewhere
        On Error Resume Next
        For Guard = 1 To MAX_REMOVALS
            Set Bar = Nothing
            Set Bar = Application.CommandBars(BAR_NAME)
            If Bar Is Nothing Then Exit For
            Bar.Delete
        Next Guard
        Err.Clear
        For Guard = 1 To MAX_REMOVALS
            Set Ctl = Nothing
            Set Ctl = Application.CommandBars.FindControl(Tag:=UI_TAG)
            If Ctl Is Nothing Then Exit For
            Ctl.Delete
        Next Guard
        Err.Clear
        KPR_UI_RemoveCommandBar = (KPR_UI_CommandBarControlCount() = 0)

End Function


Public Sub KPR_UI_CommandBarClick()
'
'==============================================================================
'                            KPR_UI_CommandBarClick
'------------------------------------------------------------------------------
' PURPOSE
'   OnAction target of every toolbar button: runs the menu action text the
'   clicked button stores in its Parameter.
'
' ERROR POLICY
'   Called from anywhere but a toolbar button (the macro list, the Immediate
'   window), it does nothing.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Action          As String       'Parameter of the clicked button

'------------------------------------------------------------------------------
' DISPATCH
'------------------------------------------------------------------------------
    'ActionControl is Nothing outside a click; that is not a failure
        On Error Resume Next
        Action = CStr(Application.CommandBars.ActionControl.Parameter)
        Err.Clear
        On Error GoTo 0
        If Len(Action) > 0 Then KPR_UI_RunAction Action

End Sub


Public Function KPR_UI_CommandBarLastReport() As String
'
' Why the last KPR_UI_InstallCommandBar failed, or "" when it succeeded.
'
    KPR_UI_CommandBarLastReport = mLastReport

End Function


Public Function KPR_UI_CommandBarControlCount() As Long
'
'==============================================================================
'                        KPR_UI_CommandBarControlCount
'------------------------------------------------------------------------------
' PURPOSE
'   Counts what KPR owns in the CommandBars collection: each bar named
'   "KPR Dates" plus every control tagged "KPR_UI" that Excel's FindControls
'   reports. Every KPR control, menus included, carries the tag, so the count
'   is 0 after a complete removal and the same after every complete install.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Bar             As Object       'CommandBar cursor
    Dim Found           As Object       'FindControls result, Nothing when none
    Dim Count           As Long         'Running total

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
    'Name and tag are the two marks of KPR ownership
        On Error Resume Next
        For Each Bar In Application.CommandBars
            If Bar.Name = BAR_NAME Then Count = Count + 1
        Next Bar
        Set Found = Application.CommandBars.FindControls(Tag:=UI_TAG)
        If Not Found Is Nothing Then Count = Count + Found.Count
        Err.Clear
        KPR_UI_CommandBarControlCount = Count

End Function


Public Function KPR_UI_CommandBarActions() As Variant
'
'==============================================================================
'                          KPR_UI_CommandBarActions
'------------------------------------------------------------------------------
' PURPOSE
'   The menu action text of every enabled button on the toolbar, at any
'   depth, so the regression suite can check each one with
'   KPR_UI_ActionKnown without clicking.
'
' RETURNS
'   A zero-based array of texts, or Empty when the toolbar is absent.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Bar             As Object       'The toolbar
    Dim Actions         As Collection   'Texts collected
    Dim Texts()         As String       'Result
    Dim I               As Long         'Result cursor

'------------------------------------------------------------------------------
' COLLECT
'------------------------------------------------------------------------------
    'An absent toolbar has no actions
        On Error Resume Next
        Set Bar = Application.CommandBars(BAR_NAME)
        Err.Clear
        On Error GoTo 0
        If Bar Is Nothing Then Exit Function
        Set Actions = New Collection
        CollectActions Bar.Controls, Actions
        If Actions.Count = 0 Then Exit Function
        ReDim Texts(0 To Actions.Count - 1)
        For I = 1 To Actions.Count
            Texts(I - 1) = Actions(I)
        Next I
        KPR_UI_CommandBarActions = Texts

End Function

'
'------------------------------------------------------------------------------
'
'                                    MENUS
'
'------------------------------------------------------------------------------
'

Private Sub AddFunctionsMenu( _
    ByVal Bar As Object)
'
' Functions > one submenu per functional module > one button per function.
'
    Dim Menu            As Object       'The Functions popup
    Dim Submenu         As Object       'One functional module
    Dim Group           As Variant      'KPR_UI_FunctionGroups cursor
    Dim Names           As Variant      'Functions of the module
    Dim FnName          As Variant      'Function cursor

    Set Menu = AddPopup(Bar.Controls, "&Functions")
    For Each Group In KPR_UI_FunctionGroups()
        Names = KPR_UI_FunctionNames(Split(CStr(Group), "|")(1))
        If IsArray(Names) Then
            Set Submenu = AddPopup(Menu.Controls, Split(CStr(Group), "|")(0))
            For Each FnName In Names
                AddButton Submenu.Controls, CStr(FnName), "fn:" & CStr(FnName), True
            Next FnName
        End If
    Next Group

End Sub


Private Sub AddDemoMenu( _
    ByVal Bar As Object)
'
' Demo sheets > one submenu per catalog group > one button per demo; a single
' disabled note when the demo modules are not loaded.
'
    Dim Menu            As Object       'The Demo sheets popup
    Dim Submenu         As Object       'One catalog group
    Dim Groups          As Variant      'KPR_UI_DemoGroups result
    Dim Group           As Variant      'Group cursor
    Dim Entries         As Variant      'Demos of the group
    Dim Entry           As Variant      'Demo cursor

    Set Menu = AddPopup(Bar.Controls, "&Demo sheets")
    Groups = KPR_UI_DemoGroups()
    If Not IsArray(Groups) Then
        AddButton Menu.Controls, "Demo modules not loaded", "", False
        Exit Sub
    End If
    For Each Group In Groups
        Entries = KPR_UI_DemoEntries(CStr(Group))
        If IsArray(Entries) Then
            Set Submenu = AddPopup(Menu.Controls, CStr(Group))
            For Each Entry In Entries
                AddButton Submenu.Controls, Split(CStr(Entry), "|")(1), _
                          "demo:" & Split(CStr(Entry), "|")(0), True
            Next Entry
        End If
    Next Group

End Sub


Private Sub AddSetupMenu( _
    ByVal Bar As Object)
'
' Setup > register / clear the Function Wizard help.
'
    Dim Menu            As Object       'The Setup popup

    Set Menu = AddPopup(Bar.Controls, "&Setup")
    AddButton Menu.Controls, "Register function help", "setup:register", True
    AddButton Menu.Controls, "Clear function help", "setup:clear", True

End Sub


Private Sub AddDevelopmentMenu( _
    ByVal Bar As Object)
'
' Development > the two regression runs; added only when the regression
' modules are loaded.
'
    Dim Menu            As Object       'The Development popup

    Set Menu = AddPopup(Bar.Controls, "De&velopment")
    AddButton Menu.Controls, "Run the pure regression suites", "dev:tests", True
    AddButton Menu.Controls, "Durable evidence run...", "dev:evidence", True

End Sub

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function AddPopup( _
    ByVal Controls As Object, _
    ByVal Caption As String) _
    As Object
'
' A temporary, KPR-tagged submenu at the end of Controls.
'
    Dim Popup           As Object       'The new popup

    Set Popup = Controls.Add(Type:=MSO_CONTROL_POPUP, Temporary:=True)
    Popup.Caption = Caption
    Popup.Tag = UI_TAG
    Set AddPopup = Popup

End Function


Private Sub AddButton( _
    ByVal Controls As Object, _
    ByVal Caption As String, _
    ByVal Action As String, _
    ByVal Enabled As Boolean)
'
' A temporary, KPR-tagged caption button that runs KPR_UI_CommandBarClick
' with Action as its Parameter; a disabled button carries no action.
'
    Dim Button          As Object       'The new button

    Set Button = Controls.Add(Type:=MSO_CONTROL_BUTTON, Temporary:=True)
    Button.Caption = Caption
    Button.Style = MSO_BUTTON_CAPTION
    Button.Tag = UI_TAG
    Button.Enabled = Enabled
    If Enabled Then
        Button.Parameter = Action
        Button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!KPR_UI_CommandBarClick"
    End If

End Sub


Private Sub CollectActions( _
    ByVal Controls As Object, _
    ByVal Actions As Collection)
'
' Adds the Parameter of every enabled button in Controls, at any depth.
'
    Dim Ctl             As Object       'Control cursor

    For Each Ctl In Controls
        If Ctl.Type = MSO_CONTROL_POPUP Then
            CollectActions Ctl.Controls, Actions
        ElseIf Ctl.Enabled Then
            Actions.Add CStr(Ctl.Parameter)
        End If
    Next Ctl

End Sub
