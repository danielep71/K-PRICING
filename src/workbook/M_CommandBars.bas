Attribute VB_Name = "M_CommandBars"

Option Explicit

'
'------------------------------------------------------------------------------
' MODULE: M_CommandBars
'------------------------------------------------------------------------------
' PURPOSE
'   The workbook's toolbar, "KPR Dates", for an Excel where the Ribbon tab is not available.
'   In Excel 2007 and later it shows on the Ribbon's Add-ins tab.
'
' WHY THIS EXISTS
'   An Excel that does not load the workbook's Ribbon, such as one whose policy turns custom
'   tabs off, still shows a toolbar built from VBA. It offers the actions of the Ribbon tab,
'   through M_Actions.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - InstallToolbar     macro list: install the toolbar, and say when that failed
'     - RemoveToolbar      macro list: remove it, and say when a control remained
'     - ToolbarClick       the OnAction of every button: runs the button's action
'     - TryInstallToolbar  install the toolbar and return the outcome, with no message
'     - TryRemoveToolbar   remove it and return the outcome, with no message
'   The toolbar has three menus: Functions, one button per function; Setup, which registers
'   or clears the help of the functions; and Development, which runs the project suites and
'   shows only in a development build. Each button stores its action text in its Parameter.
'
' ERROR POLICY
'   No member raises. A failed install removes what it built and reports why.
'
' DEPENDENCIES
'   M_Actions; the Office CommandBars objects, late bound
'
' NOTES
'   - Install first removes the bar named TOOLBAR_NAME and every control tagged TOOLBAR_TAG,
'     then builds the toolbar again, so a second install adds nothing. Both are safe to call
'     in any order and more than once. The bar and its controls are temporary: Excel discards
'     them when it closes
'   - No Option Private Module: the macro list shows InstallToolbar, RemoveToolbar and
'     ToolbarClick, which does nothing outside a click. The two Try procedures are Subs with
'     arguments, so they show in neither the macro list nor the function list
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'Office constants, written out because the objects are late bound
    Private Const MSO_BAR_TOP           As Long = 1             'msoBarTop
    Private Const MSO_CONTROL_BUTTON    As Long = 1             'msoControlButton
    Private Const MSO_CONTROL_POPUP     As Long = 10            'msoControlPopup
    Private Const MSO_BUTTON_CAPTION    As Long = 2             'msoButtonCaption

    'A bound on the removal loops: a control that refuses deletion must not loop forever
    Private Const MAX_REMOVALS          As Long = 500           'Removal loop bound

'------------------------------------------------------------------------------
' PUBLIC CONSTANTS
'------------------------------------------------------------------------------
    'The marks of what the toolbar owns
    Public Const TOOLBAR_NAME           As String = "KPR Dates" 'Name of the toolbar
    Public Const TOOLBAR_TAG            As String = "KPR_UI"    'Tag of every control it owns


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub InstallToolbar()

'
'------------------------------------------------------------------------------
'                            INSTALL THE TOOLBAR
'------------------------------------------------------------------------------
' PURPOSE
'   Installs the toolbar from the macro list, and says why when that failed.
'
' WHY THIS EXISTS
'   The macro list shows only Subs without arguments.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls TryInstallToolbar; shows a message only when it failed.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   TryInstallToolbar, M_Actions.UI_TITLE
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
    Dim Installed       As Boolean      'The toolbar is installed
    Dim Report          As String       'Why it was not

'------------------------------------------------------------------------------
' INSTALL
'------------------------------------------------------------------------------
    'Install, and say why when that failed
        TryInstallToolbar Installed, Report
        If Not Installed Then
            MsgBox "The " & TOOLBAR_NAME & " toolbar was not installed. " & Report, _
                vbExclamation, M_Actions.UI_TITLE
        End If

End Sub

Public Sub RemoveToolbar()

'
'------------------------------------------------------------------------------
'                             REMOVE THE TOOLBAR
'------------------------------------------------------------------------------
' PURPOSE
'   Removes the toolbar from the macro list, and says so when a control remained.
'
' WHY THIS EXISTS
'   The macro list shows only Subs without arguments.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls TryRemoveToolbar; shows a message only when a control remained.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   TryRemoveToolbar, M_Actions.UI_TITLE
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
    Dim Removed         As Boolean      'Nothing of the toolbar remains

'------------------------------------------------------------------------------
' REMOVE
'------------------------------------------------------------------------------
    'Remove, and say so when a control remained
        TryRemoveToolbar Removed
        If Not Removed Then
            MsgBox "Some controls of the " & TOOLBAR_NAME & " toolbar could not be removed.", _
                vbExclamation, M_Actions.UI_TITLE
        End If

End Sub

Public Sub ToolbarClick()

'
'------------------------------------------------------------------------------
'                          RUN A TOOLBAR BUTTON'S ACTION
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the action stored in the Parameter of the clicked button.
'
' WHY THIS EXISTS
'   Every button of the toolbar has this procedure as its OnAction.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Reads the Parameter of CommandBars.ActionControl and hands it to M_Actions.RunAction.
'   Outside a click there is no ActionControl, and nothing runs.
'
' ERROR POLICY
'   Does not raise: M_Actions.RunAction reports every failure itself.
'
' DEPENDENCIES
'   M_Actions.RunAction; Application.CommandBars.ActionControl
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
    Dim Clicked         As Object       'The clicked control, or Nothing outside a click
    Dim Action          As String       'Its Parameter

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Outside a click there is no control to read
        Set Clicked = Application.CommandBars.ActionControl
        If Clicked Is Nothing Then Exit Sub
        Action = VBA.CStr(Clicked.Parameter)
        Set Clicked = Nothing

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Run the stored action text
        If VBA.Len(Action) > 0 Then M_Actions.RunAction Action

End Sub

Public Sub TryInstallToolbar(ByRef Installed As Boolean, ByRef Report As String)

'
'------------------------------------------------------------------------------
'                         BUILD THE TOOLBAR, QUIETLY
'------------------------------------------------------------------------------
' PURPOSE
'   Installs the toolbar and returns the outcome, without a message.
'
' WHY THIS EXISTS
'   InstallToolbar shows the outcome to the user; a test reads it.
'
' INPUTS
'   Installed (ByRef, output)
'     Receives True when the toolbar is built and visible.
'
'   Report (ByRef, output)
'     Receives why the install failed, or an empty text.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Removes every control the toolbar owns, adds the bar, its Functions and Setup menus and,
'   when the test modules are part of this build, its Development menu, then shows it.
'
' ERROR POLICY
'   Does not raise. A failure removes what was built, and Report names the step that failed
'   and its error.
'
' DEPENDENCIES
'   TryRemoveToolbar, AddPopup, AddButton, M_Actions; Application.CommandBars
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
    Dim Bar             As Object       'The new toolbar
    Dim Menu            As Object       'The menu being filled
    Dim Item            As Variant      'Function names cursor
    Dim Stage           As String       'The step in progress, for the report
    Dim Removed         As Boolean      'Outcome of a removal
    Dim ErrNumber       As Long         'Number of the error that stopped the install
    Dim ErrDescription  As String       'Its description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start from no outcome, and report a failure instead of raising it
        Installed = False
        Report = vbNullString
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
    'Remove what an earlier install left: that is what makes install repeatable
        TryRemoveToolbar Removed
        If Not Removed Then
            Report = "Some controls of an earlier toolbar could not be removed."
            Exit Sub
        End If
    'Add the bar
        Stage = "create the toolbar"
        Set Bar = Application.CommandBars.Add(Name:=TOOLBAR_NAME, Position:=MSO_BAR_TOP, _
            Temporary:=True)
    'Add the Functions menu, one button per function
        Stage = "add the Functions menu"
        Set Menu = AddPopup(Bar.Controls, "&Functions")
        For Each Item In M_Actions.FunctionNames()
            AddButton Menu.Controls, VBA.CStr(Item), M_Actions.ACTION_FUNCTION & VBA.CStr(Item)
        Next Item
    'Add the Setup menu
        Stage = "add the Setup menu"
        Set Menu = AddPopup(Bar.Controls, "&Setup")
        AddButton Menu.Controls, "Register function help", M_Actions.ACTION_REGISTER
        AddButton Menu.Controls, "Clear function help", M_Actions.ACTION_CLEAR
    'Add the Development menu only when the test modules are part of this build
        If M_Actions.TestsAvailable() Then
            Stage = "add the Development menu"
            Set Menu = AddPopup(Bar.Controls, "De&velopment")
            AddButton Menu.Controls, "Run the project suites", M_Actions.ACTION_TESTS
        End If
    'Show the toolbar
        Stage = "show the toolbar"
        Bar.Visible = True
        Installed = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
Finish:
    'Release the objects
        Set Menu = Nothing
        Set Bar = Nothing
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Save the error before any other call
        ErrNumber = Err.Number
        ErrDescription = Err.Description
    'Leave no half-built toolbar, and say whether that cleanup succeeded
        Set Menu = Nothing
        Set Bar = Nothing
        TryRemoveToolbar Removed
        Report = "Could not " & Stage & ": error " & VBA.CStr(ErrNumber) & " (" & _
            ErrDescription & ")."
        If Not Removed Then Report = Report & " Some of its controls could not be removed."
        Installed = False
        Resume Finish

End Sub

Public Sub TryRemoveToolbar(ByRef Removed As Boolean)

'
'------------------------------------------------------------------------------
'                         REMOVE THE TOOLBAR, QUIETLY
'------------------------------------------------------------------------------
' PURPOSE
'   Removes the toolbar and every control it owns, and returns whether none remains.
'
' WHY THIS EXISTS
'   RemoveToolbar shows the outcome to the user; an install and a test read it.
'
' INPUTS
'   Removed (ByRef, output)
'     Receives True when no bar named TOOLBAR_NAME and no control tagged TOOLBAR_TAG remains.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Deletes the bar named TOOLBAR_NAME, then every control tagged TOOLBAR_TAG elsewhere, each
'   at most MAX_REMOVALS times; absent ones are not an error.
'
' ERROR POLICY
'   Does not raise. A control that refuses deletion leaves Removed False.
'
' DEPENDENCIES
'   OwnedCount; Application.CommandBars
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
    Dim Owned           As Object       'A bar or a control of the toolbar
    Dim Guard           As Long         'Loop bound

'------------------------------------------------------------------------------
' REMOVE
'------------------------------------------------------------------------------
    'Delete the bar, whose controls go with it
        For Guard = 1 To MAX_REMOVALS
            Set Owned = FindBar()
            If Owned Is Nothing Then Exit For
            If Not TryDelete(Owned) Then Exit For
        Next Guard
    'Delete every tagged control left elsewhere
        For Guard = 1 To MAX_REMOVALS
            Set Owned = FindTagged()
            If Owned Is Nothing Then Exit For
            If Not TryDelete(Owned) Then Exit For
        Next Guard
        Set Owned = Nothing
    'Report whether anything the toolbar owns remains
        Removed = (OwnedCount() = 0)

End Sub

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function AddPopup(ByVal Controls As Object, ByVal Caption As String) As Object

'
'------------------------------------------------------------------------------
'                              ADD A MENU
'------------------------------------------------------------------------------
' PURPOSE
'   Adds a temporary popup, tagged as the toolbar's, and returns it.
'
' WHY THIS EXISTS
'   Every menu carries the tag the removal looks for.
'
' INPUTS
'   Controls
'     The controls of the bar.
'
'   Caption
'     The menu's caption, with an ampersand before its access key.
'
' RETURNS
'   Object
'     The new popup.
'
' BEHAVIOR
'   None
'
' ERROR POLICY
'   Raises what Office raises; TryInstallToolbar reports it.
'
' DEPENDENCIES
'   CommandBarControls.Add
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
    Dim Popup           As Object       'The new popup

'------------------------------------------------------------------------------
' ADD
'------------------------------------------------------------------------------
    'Add the popup and tag it
        Set Popup = Controls.Add(Type:=MSO_CONTROL_POPUP, Temporary:=True)
        Popup.Caption = Caption
        Popup.Tag = TOOLBAR_TAG
        Set AddPopup = Popup
        Set Popup = Nothing

End Function

Private Sub AddButton(ByVal Controls As Object, ByVal Caption As String, ByVal Action As String)

'
'------------------------------------------------------------------------------
'                             ADD A MENU BUTTON
'------------------------------------------------------------------------------
' PURPOSE
'   Adds a temporary caption button that runs one action when clicked.
'
' WHY THIS EXISTS
'   Every button carries the tag the removal looks for, its action text and the same click
'   handler.
'
' INPUTS
'   Controls
'     The controls of a menu.
'
'   Caption
'     The button's caption.
'
'   Action
'     The action text, stored in the Parameter.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   The OnAction names ToolbarClick through this workbook's name, so that a click reaches
'   this workbook even when another one is active.
'
' ERROR POLICY
'   Raises what Office raises; TryInstallToolbar reports it.
'
' DEPENDENCIES
'   CommandBarControls.Add, ThisWorkbook
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
    Dim Button          As Object       'The new button

'------------------------------------------------------------------------------
' ADD
'------------------------------------------------------------------------------
    'Add the button, tag it and give it its action
        Set Button = Controls.Add(Type:=MSO_CONTROL_BUTTON, Temporary:=True)
        Button.Caption = Caption
        Button.Style = MSO_BUTTON_CAPTION
        Button.Tag = TOOLBAR_TAG
        Button.Parameter = Action
        Button.OnAction = "'" & VBA.Replace(ThisWorkbook.Name, "'", "''") & "'!ToolbarClick"
        Set Button = Nothing

End Sub

Private Function FindBar() As Object

'
'------------------------------------------------------------------------------
'                              FIND THE TOOLBAR
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the bar named TOOLBAR_NAME, or Nothing.
'
' WHY THIS EXISTS
'   CommandBars raises for a name it does not hold.
'
' INPUTS
'   None
'
' RETURNS
'   Object
'     The bar, or Nothing when there is none.
'
' BEHAVIOR
'   None
'
' ERROR POLICY
'   Does not raise. The error of an absent bar is Nothing, and Err is cleared.
'
' DEPENDENCIES
'   Application.CommandBars
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
    'An absent bar raises; that is Nothing
        On Error Resume Next
        Set FindBar = Application.CommandBars(TOOLBAR_NAME)
        Err.Clear
        On Error GoTo 0

End Function

Private Function FindTagged() As Object

'
'------------------------------------------------------------------------------
'                           FIND A TAGGED CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns one control tagged TOOLBAR_TAG, or Nothing.
'
' WHY THIS EXISTS
'   A control can outlive its bar, for example when it was moved to another bar.
'
' INPUTS
'   None
'
' RETURNS
'   Object
'     A tagged control, or Nothing when there is none.
'
' BEHAVIOR
'   None
'
' ERROR POLICY
'   Does not raise. An error is Nothing, and Err is cleared.
'
' DEPENDENCIES
'   Application.CommandBars.FindControl
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
    'Search every bar for the tag
        On Error Resume Next
        Set FindTagged = Application.CommandBars.FindControl(Tag:=TOOLBAR_TAG)
        Err.Clear
        On Error GoTo 0

End Function

Private Function TryDelete(ByVal Owned As Object) As Boolean

'
'------------------------------------------------------------------------------
'                           DELETE A BAR OR A CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Deletes one bar or control and says whether Office accepted.
'
' WHY THIS EXISTS
'   A control that refuses deletion must end the removal loop, not repeat it.
'
' INPUTS
'   Owned
'     A bar or a control of the toolbar.
'
' RETURNS
'   Boolean
'     True when Office deleted it; False when it refused.
'
' BEHAVIOR
'   None
'
' ERROR POLICY
'   Does not raise. A refusal is False, and Err is cleared.
'
' DEPENDENCIES
'   CommandBar.Delete, CommandBarControl.Delete
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
    Dim ErrNumber       As Long         'Error of the deletion; 0 for none

'------------------------------------------------------------------------------
' DELETE
'------------------------------------------------------------------------------
    'Delete, and read whether Office refused
        On Error Resume Next
        Owned.Delete
        ErrNumber = Err.Number
        On Error GoTo 0
        TryDelete = (ErrNumber = 0)

End Function

Private Function OwnedCount() As Long

'
'------------------------------------------------------------------------------
'                     COUNT WHAT THE TOOLBAR STILL OWNS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of bars named TOOLBAR_NAME and of controls tagged TOOLBAR_TAG.
'
' WHY THIS EXISTS
'   A removal succeeded only when this is 0.
'
' INPUTS
'   None
'
' RETURNS
'   Long
'     The count; a CommandBars collection that cannot be read counts as 1, so that it is
'     never taken for a clean removal.
'
' BEHAVIOR
'   Counts the bars by name, then the controls by tag.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   Application.CommandBars, CommandBars.FindControls
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
    Dim Bar             As Object       'CommandBars cursor
    Dim Found           As Object       'FindControls result, Nothing when none
    Dim Count           As Long         'Running total

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'A collection that cannot be read is not a clean removal
        On Error GoTo Unreadable

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
    'The name and the tag are the two marks of what the toolbar owns
        For Each Bar In Application.CommandBars
            If VBA.StrComp(Bar.Name, TOOLBAR_NAME, vbTextCompare) = 0 Then Count = Count + 1
        Next Bar
        Set Found = Application.CommandBars.FindControls(Tag:=TOOLBAR_TAG)
        If Not Found Is Nothing Then Count = Count + Found.Count
        OwnedCount = Count

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
Finish:
    'Release the objects
        Set Found = Nothing
        Set Bar = Nothing
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Unreadable:
    'Count an unreadable collection as something left
        Err.Clear
        OwnedCount = Count + 1
        Resume Finish

End Function
