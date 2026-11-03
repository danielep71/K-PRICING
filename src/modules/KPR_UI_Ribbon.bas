Attribute VB_Name = "KPR_UI_Ribbon"
'==============================================================================
' MODULE: KPR_UI_Ribbon
'------------------------------------------------------------------------------
' PURPOSE
'   The callbacks of the K-PRICING Ribbon tab, src/ribbon/customUI14.xml.
'   Every callback is a few lines that hand the clicked control's tag to
'   KPR_UI_Actions or build a menu from its lists; none holds a date
'   algorithm, registration logic or demo layout.
'
' SCOPE
'   Callbacks named in customUI14.xml
'   - KPR_UI_Ribbon_OnLoad                  onLoad: keeps the Ribbon object
'   - KPR_UI_Ribbon_OnAction                onAction of every button
'   - KPR_UI_Ribbon_GetFunctions            getContent of a Functions menu
'   - KPR_UI_Ribbon_GetDemos                getContent of a Demo sheets menu
'   - KPR_UI_Ribbon_GetDemosEnabled         getEnabled of a Demo sheets menu
'   - KPR_UI_Ribbon_GetDevelopmentVisible   getVisible of the Development group
'   Other entry points
'   - KPR_UI_RibbonRefresh                  rebuild the tab after modules change
'   - KPR_UI_RibbonLoaded                   Excel has loaded the tab, for the
'                                           one-click Excel check
'   - KPR_UI_RibbonMenuXml                  the menu XML, for the callbacks and
'                                           the regression suite
'
' RIBBON CONTRACT
'   Office 2010 and later, schema customUI/2009/07. Control ids are stable
'   and start with KPR_. Icons are built-in imageMso names. Each Functions and
'   Demo sheets menu is a dynamicMenu whose tag is a functional module's
'   caption ("Date primitives"); invalidateContentOnDrop rebuilds it every
'   time it opens, from the registration manifest and the demo catalog. A
'   button's tag is its menu action text ("setup:register", "fn:..."), which
'   KPR_UI_RunAction runs, exactly as the toolbar does.
'
' RIBBON OBJECT
'   onLoad keeps the IRibbonUI object for KPR_UI_RibbonRefresh. Resetting the
'   VBA project (an unhandled error, the editor's Reset) loses it; refresh
'   then does nothing, and the tab still works, because the menus rebuild
'   themselves on every open. Close and reopen the workbook or add-in to get
'   it back. Nothing else invalidates the Ribbon.
'
' VISIBILITY
'   No Option Private Module: Excel calls these callbacks by name. They take
'   Ribbon arguments, so they stay out of the macro list. Unsupported
'   infrastructure, not supported API (docs/PUBLIC_API.txt lists the API).
'
' ERROR POLICY
'   No callback raises: Office would show an error for the whole add-in. A
'   menu that cannot be built shows one disabled "not available" item.
'
' ALLOWED DEPENDENCIES
'   KPR_UI_Actions only. The Office library types IRibbonUI and
'   IRibbonControl.
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
    'Namespace of the menu XML a getContent callback returns
        Private Const CUSTOMUI_NS As String = "http://schemas.microsoft.com/office/2009/07/customui"
    'onAction callback of every button, static or built here
        Private Const ON_ACTION As String = "KPR_UI_Ribbon_OnAction"
    'Menu kinds KPR_UI_RibbonMenuXml builds
        Private Const KIND_FUNCTIONS As String = "functions"
        Private Const KIND_DEMOS     As String = "demos"

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    Private mRibbon         As IRibbonUI    'The loaded tab, or Nothing after a reset

'
'------------------------------------------------------------------------------
'
'                                  CALLBACKS
'
'------------------------------------------------------------------------------
'

Public Sub KPR_UI_Ribbon_OnLoad( _
    ByVal Ribbon As IRibbonUI)
'
' onLoad: keeps the Ribbon object for KPR_UI_RibbonRefresh.
'
    Set mRibbon = Ribbon

End Sub


Public Sub KPR_UI_Ribbon_OnAction( _
    ByVal Control As IRibbonControl)
'
' onAction of every button: runs the menu action text in the button's tag.
'
    On Error Resume Next
    KPR_UI_RunAction Control.Tag
    Err.Clear

End Sub


Public Sub KPR_UI_Ribbon_GetFunctions( _
    ByVal Control As IRibbonControl, _
    ByRef Content As Variant)
'
' getContent of a Functions menu: one button per supported function of the
' functional module named by the menu's tag.
'
    Content = KPR_UI_RibbonMenuXml(KIND_FUNCTIONS, Control.Tag)

End Sub


Public Sub KPR_UI_Ribbon_GetDemos( _
    ByVal Control As IRibbonControl, _
    ByRef Content As Variant)
'
' getContent of a Demo sheets menu: one button per catalog demo of the group
' named by the menu's tag.
'
    Content = KPR_UI_RibbonMenuXml(KIND_DEMOS, Control.Tag)

End Sub


Public Sub KPR_UI_Ribbon_GetDemosEnabled( _
    ByVal Control As IRibbonControl, _
    ByRef Enabled As Variant)
'
' getEnabled of a Demo sheets menu: TRUE when the demo modules are loaded and
' the menu's group has at least one demo.
'
    On Error Resume Next
    Enabled = IsArray(KPR_UI_DemoEntries(Control.Tag))
    If Err.Number <> 0 Then Enabled = False
    Err.Clear

End Sub


Public Sub KPR_UI_Ribbon_GetDevelopmentVisible( _
    ByVal Control As IRibbonControl, _
    ByRef Visible As Variant)
'
' getVisible of the Development group: TRUE only when the regression modules
' are loaded.
'
    On Error Resume Next
    Visible = KPR_UI_TestsAvailable()
    If Err.Number <> 0 Then Visible = False
    Err.Clear

End Sub

'
'------------------------------------------------------------------------------
'
'                                ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub KPR_UI_RibbonRefresh()
'
'==============================================================================
'                             KPR_UI_RibbonRefresh
'------------------------------------------------------------------------------
' PURPOSE
'   Asks Excel to rebuild the K-PRICING tab, for example after the demo or
'   regression modules were loaded or removed, so the Development group and
'   the Demo sheets menus follow.
'
' BEHAVIOR
'   The only place that invalidates the Ribbon. Does nothing when the Ribbon
'   object was lost to a VBA reset or the tab is not loaded. Never raises.
'
' UPDATED
'   2026-11-03
'==============================================================================
'
    If mRibbon Is Nothing Then Exit Sub
    On Error Resume Next
    mRibbon.Invalidate
    Err.Clear

End Sub


Public Function KPR_UI_RibbonLoaded() As Boolean
'
' TRUE once Excel has loaded this project's K-PRICING tab (onLoad ran) and
' no VBA reset has lost the Ribbon object since. KPR_Dev_CheckBuild reads it
' after opening a copy that carries the tab.
'
    KPR_UI_RibbonLoaded = Not mRibbon Is Nothing

End Function


Public Function KPR_UI_RibbonMenuXml( _
    ByVal Kind As String, _
    ByVal Group As String) _
    As String
'
'==============================================================================
'                             KPR_UI_RibbonMenuXml
'------------------------------------------------------------------------------
' PURPOSE
'   The <menu> XML of a dynamicMenu: one button per function or demo of one
'   functional module.
'
' INPUTS
'   Kind
'     "functions" or "demos".
'   Group
'     A functional module's caption, as in KPR_UI_FunctionGroups and the demo
'     catalog, for example "Date primitives".
'
' RETURNS
'   A <menu> element in the customUI/2009/07 namespace. Every button has a
'   stable id (KPR_Fn_<name> or KPR_DemoItem_<entry point>), the escaped
'   label, the menu action text as its tag and KPR_UI_Ribbon_OnAction. When
'   nothing is available, or the menu cannot be built, a single disabled
'   button says so.
'
' UPDATED
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Items           As Variant      'Function names or "EntryPoint|Caption" texts
    Dim Item            As Variant      'Items cursor
    Dim Body            As String       'The buttons

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
    'Read the same lists the toolbar reads, so both surfaces agree
        On Error GoTo Unavailable
        Select Case LCase$(Kind)
            Case KIND_FUNCTIONS
                Items = KPR_UI_FunctionNames(GroupPrefix(Group))
                If IsArray(Items) Then
                    For Each Item In Items
                        Body = Body & MenuButton("KPR_Fn_" & IdPart(CStr(Item)), CStr(Item), "fn:" & CStr(Item))
                    Next Item
                End If
            Case KIND_DEMOS
                Items = KPR_UI_DemoEntries(Group)
                If IsArray(Items) Then
                    For Each Item In Items
                        Body = Body & MenuButton("KPR_DemoItem_" & IdPart(Split(CStr(Item), "|")(0)), _
                                                 Split(CStr(Item), "|")(1), _
                                                 "demo:" & Split(CStr(Item), "|")(0))
                    Next Item
                End If
        End Select
        If Len(Body) = 0 Then GoTo Unavailable
        KPR_UI_RibbonMenuXml = MenuXml(Body)
        Exit Function

'------------------------------------------------------------------------------
' HANDLE EMPTY OR FAILED MENU
'------------------------------------------------------------------------------
Unavailable:
    'One disabled item tells the user why the menu is empty
        KPR_UI_RibbonMenuXml = MenuXml("<button id=""KPR_" & IdPart(Kind) & "_None"" " & _
                                       "label=""Not available in this build"" enabled=""false""/>")

End Function

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function GroupPrefix( _
    ByVal Group As String) _
    As String
'
' The function-name prefix of a functional module's caption, from
' KPR_UI_FunctionGroups; a prefix no function has when the caption is unknown.
'
    Dim Entry           As Variant      'KPR_UI_FunctionGroups cursor

    GroupPrefix = "KPR_NoSuchGroup_"
    For Each Entry In KPR_UI_FunctionGroups()
        If StrComp(Split(CStr(Entry), "|")(0), Group, vbTextCompare) = 0 Then
            GroupPrefix = Split(CStr(Entry), "|")(1)
        End If
    Next Entry

End Function


Private Function MenuXml( _
    ByVal Body As String) _
    As String
'
' Wraps buttons in the <menu> element a getContent callback returns.
'
    MenuXml = "<menu xmlns=""" & CUSTOMUI_NS & """>" & Body & "</menu>"

End Function


Private Function MenuButton( _
    ByVal Id As String, _
    ByVal Label As String, _
    ByVal Action As String) _
    As String
'
' One menu button that runs Action through KPR_UI_Ribbon_OnAction.
'
    MenuButton = "<button id=""" & Id & """ label=""" & XmlText(Label) & """ tag=""" & XmlText(Action) & _
                 """ onAction=""" & ON_ACTION & """/>"

End Function


Private Function IdPart( _
    ByVal Text As String) _
    As String
'
' Text with every character outside A-Z, a-z, 0-9 and "_" replaced by "_",
' so it can be part of a Ribbon control id.
'
    Dim I               As Long         'Character cursor
    Dim Ch              As String       'Character I

    For I = 1 To Len(Text)
        Ch = Mid$(Text, I, 1)
        If Ch Like "[A-Za-z0-9_]" Then
            IdPart = IdPart & Ch
        Else
            IdPart = IdPart & "_"
        End If
    Next I

End Function


Private Function XmlText( _
    ByVal Text As String) _
    As String
'
' Text escaped for a double-quoted XML attribute value.
'
    XmlText = Replace(Text, "&", "&amp;")
    XmlText = Replace(XmlText, "<", "&lt;")
    XmlText = Replace(XmlText, ">", "&gt;")
    XmlText = Replace(XmlText, """", "&quot;")

End Function
