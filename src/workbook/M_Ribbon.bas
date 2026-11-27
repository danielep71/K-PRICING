Attribute VB_Name = "M_Ribbon"

Option Explicit

'
'------------------------------------------------------------------------------
' MODULE: M_Ribbon
'------------------------------------------------------------------------------
' PURPOSE
'   The callbacks of the workbook's Ribbon tab, src/workbook/customUI14.xml, and the menu of
'   its Functions button.
'
' WHY THIS EXISTS
'   Excel calls a Ribbon callback by its name. Each callback hands the clicked control's tag to
'   M_Actions or builds a menu from its lists; none holds an action of its own.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Callbacks the XML names:
'     - RibbonOnAction               onAction of every button: runs the tag's action
'     - RibbonGetFunctions           getContent of the Functions menu
'     - RibbonGetDevelopmentVisible  getVisible of the Development group
'   Other entry point:
'     - BuildFunctionMenu            the menu XML of the Functions button, for the callback
'                                    and the tests
'
' ERROR POLICY
'   No callback raises: Excel would show an error for the whole tab. A menu that cannot be
'   built shows one disabled item, and a group whose visibility cannot be read stays hidden.
'
' DEPENDENCIES
'   M_Actions; the Office library types IRibbonControl
'
' NOTES
'   - The Ribbon is Office 2010's, schema customUI/2009/07. Control ids start with KPR_ and
'     stay stable; icons are built-in imageMso names. tools/check_ribbon.py checks the XML
'     and that every callback it names is a Public Sub of this module with the signature
'     Office calls
'   - The Functions menu is a dynamicMenu with invalidateContentOnDrop, so Excel asks for its
'     content each time it opens: no Ribbon object is kept and nothing invalidates the tab
'   - No Option Private Module: Excel finds the callbacks by name. Every public member is a
'     Sub with arguments, so none shows in the macro list or the function list
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'The namespace of the menu XML a getContent callback returns
    Private Const CUSTOMUI_NS           As String = "http://schemas.microsoft.com/office/2009/07/customui"

    'The onAction callback of every button of the Functions menu
    Private Const ON_ACTION             As String = "RibbonOnAction"    'This module's callback

    'The id prefix of a function's button in the Functions menu
    Private Const FUNCTION_ID           As String = "KPR_Fn_"           'Followed by the name


'
'------------------------------------------------------------------------------
'
'                                  CALLBACKS
'
'------------------------------------------------------------------------------
'

Public Sub RibbonOnAction(ByVal Control As IRibbonControl)

'
'------------------------------------------------------------------------------
'                           RUN A BUTTON'S ACTION
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the action the clicked button's tag names.
'
' WHY THIS EXISTS
'   Every button of the tab, those of the Functions menu included, stores its action text in
'   its tag, as the toolbar stores it in a button's Parameter.
'
' INPUTS
'   Control
'     The clicked button.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Hands the tag to M_Actions.RunAction.
'
' ERROR POLICY
'   Does not raise: M_Actions.RunAction reports every failure itself.
'
' DEPENDENCIES
'   M_Actions.RunAction
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Run the action text held in the tag
        M_Actions.RunAction Control.Tag

End Sub

Public Sub RibbonGetFunctions(ByVal Control As IRibbonControl, ByRef Content As Variant)

'
'------------------------------------------------------------------------------
'                         BUILD THE FUNCTIONS MENU
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the content of the Functions menu: one button per function.
'
' WHY THIS EXISTS
'   The menu is built from M_Actions.FunctionNames, the list the toolbar reads, so the two
'   surfaces list the same functions.
'
' INPUTS
'   Control
'     The Functions menu; not read.
'
'   Content (ByRef, output)
'     Receives the menu XML.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Returns what BuildFunctionMenu builds.
'
' ERROR POLICY
'   Does not raise: BuildFunctionMenu returns one disabled item when it cannot build the menu.
'
' DEPENDENCIES
'   BuildFunctionMenu
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
    Dim Xml             As String       'The menu XML

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
    'Return the menu XML
        BuildFunctionMenu Xml
        Content = Xml

End Sub

Public Sub RibbonGetDevelopmentVisible(ByVal Control As IRibbonControl, ByRef Visible As Variant)

'
'------------------------------------------------------------------------------
'                      SHOW THE DEVELOPMENT GROUP OR NOT
'------------------------------------------------------------------------------
' PURPOSE
'   Shows the Development group only in a development build.
'
' WHY THIS EXISTS
'   Its button runs the test modules, which a production build does not hold.
'
' INPUTS
'   Control
'     The Development group; not read.
'
'   Visible (ByRef, output)
'     Receives True when the test modules are part of this build.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Returns M_Actions.TestsAvailable.
'
' ERROR POLICY
'   Does not raise: M_Actions.TestsAvailable is False when it cannot tell.
'
' DEPENDENCIES
'   M_Actions.TestsAvailable
'
' NOTES
'   Excel asks once, when it loads the tab
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ANSWER
'------------------------------------------------------------------------------
    'Show the group when the test modules are loaded
        Visible = M_Actions.TestsAvailable()

End Sub

'
'------------------------------------------------------------------------------
'
'                                  MENU XML
'
'------------------------------------------------------------------------------
'

Public Sub BuildFunctionMenu(ByRef Xml As String)

'
'------------------------------------------------------------------------------
'                        WRITE THE FUNCTIONS MENU XML
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the XML of the Functions menu: one button per function, in registration order.
'
' WHY THIS EXISTS
'   RibbonGetFunctions returns it to Excel, and the tests read it without a Ribbon.
'
' INPUTS
'   Xml (ByRef, output)
'     Receives a menu element of the customUI namespace.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Each button has the id FUNCTION_ID and the function name, the name as its label, the
'   action text M_Actions.ACTION_FUNCTION and the name as its tag, and RibbonOnAction as its
'   onAction. When the list cannot be read, the menu holds one disabled button.
'
' ERROR POLICY
'   Does not raise. A failure gives the menu with the disabled button, and Err is cleared.
'
' DEPENDENCIES
'   M_Actions.FunctionNames, M_Actions.ACTION_FUNCTION, MenuButton
'
' NOTES
'   A Sub with an output argument, not a Function, so that it does not show in Excel's
'   function list
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Names           As Variant      'The function names
    Dim Item            As Variant      'Names cursor
    Dim Body            As String       'The buttons

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'A failure gives the menu with the disabled button
        On Error GoTo Unavailable

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
    'One button per function, from the list the toolbar reads
        Names = M_Actions.FunctionNames()
        For Each Item In Names
            Body = Body & MenuButton(FUNCTION_ID & VBA.CStr(Item), VBA.CStr(Item), _
                M_Actions.ACTION_FUNCTION & VBA.CStr(Item))
        Next Item
        If VBA.Len(Body) = 0 Then GoTo Unavailable
        Xml = "<menu xmlns=""" & CUSTOMUI_NS & """>" & Body & "</menu>"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Unavailable:
    'One disabled button tells the user why the menu is empty
        Err.Clear
        Xml = "<menu xmlns=""" & CUSTOMUI_NS & """><button id=""KPR_Fn_None"" " & _
            "label=""Not available"" enabled=""false""/></menu>"

End Sub

Private Function MenuButton( _
    ByVal Id As String, _
    ByVal Label As String, _
    ByVal Action As String) As String

'
'------------------------------------------------------------------------------
'                            WRITE A MENU BUTTON
'------------------------------------------------------------------------------
' PURPOSE
'   Returns one button element of a dynamic menu.
'
' WHY THIS EXISTS
'   Every text that goes into an attribute is escaped the same way.
'
' INPUTS
'   Id
'     The control id: letters, digits and underscores.
'
'   Label
'     The text the menu shows.
'
'   Action
'     The action text, stored in the tag.
'
' RETURNS
'   String
'     A button element with RibbonOnAction as its onAction.
'
' BEHAVIOR
'   None
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   XmlText
'
' NOTES
'   None
'
' UPDATED
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Build the element with its escaped texts
        MenuButton = "<button id=""" & XmlText(Id) & """ label=""" & XmlText(Label) & _
            """ tag=""" & XmlText(Action) & """ onAction=""" & ON_ACTION & """/>"

End Function

Private Function XmlText(ByVal Text As String) As String

'
'------------------------------------------------------------------------------
'                          ESCAPE AN ATTRIBUTE TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a text that can stand between double quotes in an XML attribute.
'
' WHY THIS EXISTS
'   A label or a tag with an ampersand, an angle bracket or a quote would break the menu.
'
' INPUTS
'   Text
'     The text.
'
' RETURNS
'   String
'     The text with &, <, > and " replaced by their entities.
'
' BEHAVIOR
'   Escapes the ampersand first, so that the entities written after it stay intact.
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
'   2026-11-27
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ESCAPE
'------------------------------------------------------------------------------
    'Escape the ampersand first, then the brackets and the quote
        XmlText = VBA.Replace(Text, "&", "&amp;")
        XmlText = VBA.Replace(XmlText, "<", "&lt;")
        XmlText = VBA.Replace(XmlText, ">", "&gt;")
        XmlText = VBA.Replace(XmlText, """", "&quot;")

End Function
