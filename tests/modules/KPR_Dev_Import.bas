Attribute VB_Name = "KPR_Dev_Import"
'==============================================================================
' MODULE: KPR_Dev_Import
'------------------------------------------------------------------------------
' PURPOSE
'   Development loader: replaces the K-PRICING modules of an open workbook or
'   add-in with the versions of one commit, in a single call, instead of
'   removing and importing each module by hand.
'
' WHY THIS EXISTS
'   Every Excel check of a pull request means importing the changed modules
'   into the developer's build. This loader takes the module list from the
'   commit's own .github/repository-profile.json, so new, renamed and removed
'   modules follow the repository without editing the loader.
'
' SCOPE
'   - KPR_Dev_CheckLatest       the one macro to run (Alt+F8): checks the
'                               latest main in KPR.xlam
'   - KPR_Dev_CheckBuild        import, compile, save, two regression runs
'                               and a Ribbon load check of one commit
'   - KPR_Dev_LastCheckFolder   the evidence folder of the last check, for
'                               tools/Check-KPR.ps1
'   - KPR_Dev_ImportModules     import one commit's modules into a target
'
' WHERE IT LIVES
'   In a workbook other than the target: the Personal Macro Workbook
'   (PERSONAL.XLSB) or a small development workbook. It refuses to change
'   the project it runs from, and it never imports itself.
'
' SOURCES
'   - A commit SHA or branch name: files are downloaded from
'     raw.githubusercontent.com with MSXML2.XMLHTTP, which uses the Windows
'     proxy settings. A full SHA is exact; a branch may be served from a
'     cache for a few minutes after a push.
'   - A local folder: the root of a clone of the repository.
'
' WHAT IT CHANGES
'   - Replaces every standard module the profile lists in the selected roles.
'   - Removes repository modules the selected commit no longer has: a module
'     the loader imported before (it records each import as a hidden name
'     KPRDEV_<module> in the target) or one the repository has tracked at any
'     time (KNOWN_MODULES), when the commit's profile does not list it in any
'     role. Loading an older commit therefore restores that commit's set.
'   - Leaves every other module alone, such as private ribbon or menu
'     modules, and modules of roles not selected.
'   All files are fetched and checked before the target is touched, so a
'   failed download changes nothing. The target is not saved: compile it,
'   then save it.
'
' REQUIREMENTS
'   Excel option "Trust access to the VBA project object model" (File >
'   Options > Trust Center > Trust Center Settings > Macro Settings), and an
'   unlocked target project. Without them the loader stops and says why.
'
' ERROR POLICY
'   Never raises. Returns the number of modules imported, or -1 with the
'   reason printed to the Immediate window. A failure after the first change
'   leaves the target partly updated and unsaved; close it without saving to
'   discard.
'
' VISIBILITY
'   No Option Private Module, so KPR_Dev_CheckLatest appears in the macro
'   list (Alt+F8). Development infrastructure, never imported into a build.
'
' ALLOWED DEPENDENCIES
'   None.
'
' UPDATED
'   2026-11-09
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
    'Raw-file address of the repository; the ref and path are appended
        Private Const RAW_BASE As String = "https://raw.githubusercontent.com/danielep71/k-pricing/"

    'Profile that lists every module with its role
        Private Const PROFILE_PATH As String = ".github/repository-profile.json"

    'Every standard module the repository has tracked, under any name. One the
    'selected commit's profile does not list is removed from the target
        Private Const KNOWN_MODULES As String = _
            "KPR_Core_Array,KPR_Core_Dates,KPR_Core_Err,KPR_Core_Parse,KPR_DATES_DAYS," & _
            "KPR_REGISTER_PUBLIC_UDFS,KPR_REGRESSION_TESTS,KPR_Test_Fixtures_Generated," & _
            "KPR_Test_Oracle,KPR_DateExample,KPR_Demo_Builder,KPR__Demo_Builder," & _
            "KPR_Demo__Builder,KPR_Demo_Dates,KPR_Demo_DirectVBA,KPR_UI_Actions," & _
            "KPR_UI_CommandBars,KPR_UI_Ribbon"

    'Prefix of the hidden workbook names that record what the loader imported
        Private Const RECORD_PREFIX As String = "KPRDEV_"

    'Default roles to import
        Private Const DEFAULT_ROLES As String = "public,internal,ui,test,example"

    'VBComponent type of a standard module
        Private Const STANDARD_MODULE As Long = 1

    'Build KPR_Dev_CheckLatest checks, and where a branch is resolved to a SHA
        Private Const DEFAULT_TARGET As String = "KPR.xlam"
        Private Const API_COMMITS As String = "https://api.github.com/repos/danielep71/K-PRICING/commits/"

    'Control id of Debug > Compile in the VBA editor
        Private Const COMPILE_COMMAND As Long = 578

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    Private mLastFolder     As String       'Evidence folder of the last check

'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Function KPR_Dev_ImportModules( _
    ByVal Source As String, _
    ByVal TargetWorkbook As String, _
    Optional ByVal Roles As String = DEFAULT_ROLES) _
    As Long
'
'==============================================================================
'                            KPR_Dev_ImportModules
'------------------------------------------------------------------------------
' PURPOSE
'   Replaces the K-PRICING modules of TargetWorkbook with those of Source.
'
' INPUTS
'   Source
'     A commit SHA or branch name to download, or the path of a local clone.
'   TargetWorkbook
'     Name of the open workbook or add-in to update, e.g. "KPR.xlam".
'   Roles
'     Optional. Comma-separated profile roles to import; default all five
'     (public, internal, ui, test, example).
'
' RETURNS
'   The number of modules imported, or -1 when the loader stopped; the
'   Immediate window says why.
'
' EXAMPLE
'   ? KPR_Dev_ImportModules("ce88bd1fead665c88760d01f132e75e5da1000d0", "KPR.xlam")
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Target      As Workbook     'Workbook or add-in to update
    Dim Project     As Object       'Its VBProject
    Dim Folder      As String       'Staging folder for the fetched files
    Dim Profile     As String       'Text of the repository profile
    Dim Paths       As Collection   'Repository paths of the modules to import
    Dim Listed      As Collection   'Every module the profile lists, in any role
    Dim Names       As Collection   'Module names, in the same order
    Dim Files       As Collection   'Staged files, in the same order
    Dim Detail      As String       'Reason a step failed
    Dim Item        As Variant      'Cursor
    Dim K           As Long         'Module cursor
    Dim Removed     As Long         'Retired modules removed
    Dim Stage       As String       'Step, for the failure report
    Dim Imported    As Object       'Module created by one import

    KPR_Dev_ImportModules = -1
    On Error GoTo Fail

'------------------------------------------------------------------------------
' TARGET
'------------------------------------------------------------------------------
    Stage = "find the target"
    If Not TryTarget(TargetWorkbook, Target, Project, Detail) Then GoTo Refuse

'------------------------------------------------------------------------------
' FETCH AND CHECK EVERYTHING BEFORE ANY CHANGE
'------------------------------------------------------------------------------
    Stage = "prepare the staging folder"
    Folder = StagingFolder()
    Stage = "read the repository profile"
    If Not TryFetch(Source, PROFILE_PATH, Folder & Application.PathSeparator & "profile.json", Detail) Then GoTo Refuse
    Profile = ReadText(Folder & Application.PathSeparator & "profile.json")
    If Not TrySelect(Profile, Roles, Paths, Listed, Detail) Then GoTo Refuse

    Stage = "fetch the modules"
    Set Names = New Collection
    Set Files = New Collection
    For Each Item In Paths
        K = K + 1
        Files.Add Folder & Application.PathSeparator & CStr(K) & "_" & FileName(CStr(Item))
        If Not TryFetch(Source, CStr(Item), Files(K), Detail) Then GoTo Refuse
        If Not TryModuleName(Files(K), FileStem(CStr(Item)), Detail) Then GoTo Refuse
        Names.Add FileStem(CStr(Item))
    Next Item

'------------------------------------------------------------------------------
' REPLACE
'------------------------------------------------------------------------------
    Stage = "remove modules the commit does not have"
    For Each Item In Candidates(Target)
        If Not Contains(Listed, CStr(Item)) Then
            If RemoveModule(Project, CStr(Item)) Then
                Removed = Removed + 1
                Debug.Print "  removed  " & CStr(Item) & " (not in this commit)"
            End If
            Forget Target, CStr(Item)
        End If
    Next Item

    For K = 1 To Names.Count
        Stage = "replace " & Names(K)
        RemoveModule Project, Names(K)
        Set Imported = Project.VBComponents.Import(Files(K))
        If StrComp(Imported.Name, Names(K), vbBinaryCompare) <> 0 Then
            Detail = Names(K) & " was imported as " & Imported.Name & "; rename or remove it by hand."
            GoTo Refuse
        End If
        Remember Target, Names(K)
        Debug.Print "  imported " & Names(K) & "  (" & Paths(K) & ")"
    Next K

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    Debug.Print "KPR_Dev_ImportModules: " & CStr(Names.Count) & " module(s) imported and " & _
                CStr(Removed) & " module(s) removed in " & Target.Name & " from " & Source & "."
    Debug.Print "Next: Debug > Compile VBAProject, then save " & Target.Name & "."
    KPR_Dev_ImportModules = Names.Count
    CleanUp Folder
    Exit Function

Refuse:
    Debug.Print "KPR_Dev_ImportModules stopped: " & Detail
    CleanUp Folder
    Exit Function

Fail:
    Debug.Print "KPR_Dev_ImportModules stopped: could not " & Stage & ": error " & _
                CStr(Err.Number) & " (" & Err.Description & ")"
    CleanUp Folder

End Function

'
'------------------------------------------------------------------------------
'
'                              ONE-CLICK EXCEL CHECK
'
'------------------------------------------------------------------------------
'

Public Sub KPR_Dev_CheckLatest()
'
'==============================================================================
'                             KPR_Dev_CheckLatest
'------------------------------------------------------------------------------
' PURPOSE
'   The one macro to run (Alt+F8) to check the latest main in Excel: runs
'   KPR_Dev_CheckBuild on the current main for KPR.xlam and shows the
'   outcome. Send the files of the folder it opens.
'
' UPDATED
'   2026-11-03
'==============================================================================
'
    KPR_Dev_CheckBuild "main", DEFAULT_TARGET

End Sub


Public Function KPR_Dev_CheckBuild( _
    ByVal Source As String, _
    ByVal TargetWorkbook As String, _
    Optional ByVal Unattended As Boolean = False) _
    As Boolean
'
'==============================================================================
'                              KPR_Dev_CheckBuild
'------------------------------------------------------------------------------
' PURPOSE
'   Does every step of an Excel check of one commit, unattended:
'     1. resolves Source to its exact 40-character commit SHA;
'     2. imports that commit's modules into TargetWorkbook;
'     3. compiles its VBA project and saves it;
'     4. runs KPR_Test_RunAll twice, into run-1 and run-2;
'     5. loads the commit's Ribbon into a copy of the target and checks that
'        Excel loaded it (Windows only);
'   then writes summary.txt, opens the folder and shows the outcome.
'   tools/Check-KPR.ps1 calls it Unattended, from a fresh Excel it starts.
'
' INPUTS
'   Source
'     A branch name such as "main", or a full commit SHA.
'   TargetWorkbook
'     The open workbook or add-in to update, e.g. "KPR.xlam".
'   Unattended
'     TRUE: no message box and no Explorer window, for a script that drives
'     Excel; KPR_Dev_LastCheckFolder then gives the folder.
'
' RETURNS
'   TRUE when every step passed.
'
' OUTPUT
'   A new folder Documents\KPR-evidence\<date>-<SHA prefix> holding
'   run-1\kpr-test-evidence.json, run-2\kpr-test-evidence.json and
'   summary.txt. The Ribbon copy is deleted after the check.
'
' ERROR POLICY
'   Never raises. A failed step is reported in summary.txt and the message;
'   the steps after a failed import or compile are skipped, because they
'   would test the wrong code.
'
' UPDATED
'   2026-11-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sha         As String       'Exact commit checked
    Dim Folder      As String       'Output folder
    Dim Report      As String       'summary.txt text
    Dim Detail      As String       'Reason a step failed
    Dim Target      As Workbook     'The updated workbook or add-in
    Dim Imported    As Long         'Result of KPR_Dev_ImportModules
    Dim Run1        As Boolean      'First regression run passed
    Dim Run2        As Boolean      'Second regression run passed
    Dim Ribbon      As String       'PASS, FAIL or NOT RUN
    Dim Passed      As Boolean      'Every step passed

'------------------------------------------------------------------------------
' RESOLVE, IMPORT, COMPILE AND SAVE
'------------------------------------------------------------------------------
    'Stop at the first failure: later steps would test the wrong code
        On Error GoTo Fail
        Application.StatusBar = "K-PRICING check: resolving " & Source & "..."
        If Not TryResolveSha(Source, Sha, Detail) Then GoTo Finish
        Report = "Source:  " & Source & vbCrLf & "Commit:  " & Sha & vbCrLf & _
                 "Target:  " & TargetWorkbook & vbCrLf & "Started: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
        Application.StatusBar = "K-PRICING check: importing " & Left$(Sha, 8) & "..."
        Imported = KPR_Dev_ImportModules(Sha, TargetWorkbook)
        If Imported < 0 Then
            Detail = "Import failed; the Immediate window (Ctrl+G in the VBA editor) says why."
            GoTo Finish
        End If
        Report = Report & "Import:  PASS (" & CStr(Imported) & " modules)" & vbCrLf
        Set Target = Workbooks(TargetWorkbook)
        Application.StatusBar = "K-PRICING check: compiling..."
        If Not TryCompile(Target.VBProject, Detail) Then GoTo Finish
        Report = Report & "Compile: PASS" & vbCrLf
        Target.Save
        Report = Report & "Save:    PASS" & vbCrLf

'------------------------------------------------------------------------------
' REGRESSION RUNS
'------------------------------------------------------------------------------
    'Two runs, so the evidence can be compared for determinism
        Folder = EvidenceFolder(Sha)
        Application.StatusBar = "K-PRICING check: regression run 1 of 2..."
        Run1 = CBool(Application.Run(QualifiedIn(Target, "KPR_Test_RunAll"), Sha, _
                                     Folder & Application.PathSeparator & "run-1"))
        Report = Report & "Run 1:   " & IIf(Run1, "PASS", "FAIL") & vbCrLf
        Application.StatusBar = "K-PRICING check: regression run 2 of 2..."
        Run2 = CBool(Application.Run(QualifiedIn(Target, "KPR_Test_RunAll"), Sha, _
                                     Folder & Application.PathSeparator & "run-2"))
        Report = Report & "Run 2:   " & IIf(Run2, "PASS", "FAIL") & vbCrLf

'------------------------------------------------------------------------------
' RIBBON
'------------------------------------------------------------------------------
    'A copy carries the tab, so the user's own add-in is never rewritten
        Application.StatusBar = "K-PRICING check: loading the Ribbon in a copy..."
        Ribbon = RibbonCheck(Target, Sha, Folder, Detail)
        Report = Report & "Ribbon:  " & Ribbon & IIf(Len(Detail) > 0, " (" & Detail & ")", "") & vbCrLf
        Detail = vbNullString
        Passed = Run1 And Run2 And Ribbon = "PASS"

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
Finish:
    'Every path ends here: write what is known, show it, open the folder
        On Error Resume Next
        Application.StatusBar = False
        If Len(Detail) > 0 Then Report = Report & "STOPPED: " & Detail & vbCrLf
        Report = Report & "Result:  " & IIf(Passed, "PASS", "FAIL") & vbCrLf
        If Len(Folder) = 0 And Len(Sha) > 0 Then Folder = EvidenceFolder(Sha)
        mLastFolder = Folder
        If Len(Folder) > 0 Then
            WriteText Folder & Application.PathSeparator & "summary.txt", Report
            If Not Unattended Then Shell "explorer.exe """ & Folder & """", vbNormalFocus
            Report = Report & vbCrLf & "Send the files of this folder:" & vbCrLf & Folder
        End If
        Debug.Print Report
        If Not Unattended Then MsgBox Report, IIf(Passed, vbInformation, vbExclamation), "K-PRICING check"
        KPR_Dev_CheckBuild = Passed
        Exit Function

Fail:
    'An unexpected error still reaches the report
        Detail = "error " & CStr(Err.Number) & " (" & Err.Description & ")"
        Resume Finish

End Function


Public Function KPR_Dev_LastCheckFolder() As String
'
'==============================================================================
'                           KPR_Dev_LastCheckFolder
'------------------------------------------------------------------------------
' PURPOSE
'   The evidence folder KPR_Dev_CheckBuild last wrote, or "" when it wrote
'   none (it stopped before resolving the commit). tools/Check-KPR.ps1 reads
'   it after an unattended check.
'
' UPDATED
'   2026-11-09
'==============================================================================
'
    KPR_Dev_LastCheckFolder = mLastFolder

End Function


Private Function TryResolveSha( _
    ByVal Source As String, _
    ByRef Sha As String, _
    ByRef Detail As String) _
    As Boolean
'
' A full SHA as given, or the commit a branch points to now, from the GitHub
' API, so the evidence is bound to an exact commit.
'
    Dim Http    As Object       'MSXML2.XMLHTTP request

    If IsSha(Source) Then
        Sha = LCase$(Source)
        TryResolveSha = True
        Exit Function
    End If
    On Error GoTo Failed
    Set Http = CreateObject("MSXML2.XMLHTTP.6.0")
    Http.Open "GET", API_COMMITS & Source, False
    Http.setRequestHeader "Accept", "application/vnd.github.sha"
    Http.send
    Sha = LCase$(Trim$(Replace(Replace(CStr(Http.responseText), vbCr, ""), vbLf, "")))
    If Http.Status = 200 And IsSha(Sha) Then
        TryResolveSha = True
    Else
        Detail = "GitHub returned " & CStr(Http.Status) & " for the branch " & Source & "."
    End If
    Exit Function

Failed:
    Detail = "could not ask GitHub for " & Source & ": error " & CStr(Err.Number) & " (" & Err.Description & ")"

End Function


Private Function TryCompile( _
    ByVal Project As Object, _
    ByRef Detail As String) _
    As Boolean
'
' Compiles Project with the VBA editor's own Debug > Compile command. The
' command is disabled once the project is compiled, so it still being
' enabled afterwards means a compile error, which the editor shows.
'
    Dim CompileCmd  As Object       'Debug > Compile in the VBA editor

    On Error GoTo Failed
    Set Application.VBE.ActiveVBProject = Project
    Set CompileCmd = Application.VBE.CommandBars.FindControl(Type:=1, Id:=COMPILE_COMMAND)
    If CompileCmd Is Nothing Then
        Detail = "the VBA editor has no Debug > Compile command."
        Exit Function
    End If
    If CompileCmd.Enabled Then CompileCmd.Execute
    DoEvents
    TryCompile = Not CompileCmd.Enabled
    If Not TryCompile Then Detail = "compile error; the VBA editor shows the line."
    Exit Function

Failed:
    Detail = "could not compile: error " & CStr(Err.Number) & " (" & Err.Description & ")"

End Function


Private Function RibbonCheck( _
    ByVal Target As Workbook, _
    ByVal Sha As String, _
    ByVal Folder As String, _
    ByRef Detail As String) _
    As String
'
' Writes the commit's src/ribbon/customUI14.xml into a copy of Target with
' PowerShell, opens and closes the copy with events off (the caller's setting
' is restored on every exit) and returns PASS when
' Excel loaded the tab (KPR_UI_RibbonLoaded) and its Functions menu builds,
' FAIL otherwise, or NOT RUN where PowerShell is unavailable. The copy is
' closed and deleted.
'
    Dim Work        As String       'Ribbon working folder
    Dim CopyPath    As String       'The copy that receives the tab
    Dim Script      As String       'PowerShell injector
    Dim ExitCode    As Long         'Its exit code
    Dim Copy        As Workbook     'The opened copy
    Dim Events      As Boolean      'EnableEvents on entry
    Dim Loaded      As Boolean      'onLoad ran in the copy
    Dim Started     As Single       'Wait start
    Dim Menu        As String       'Functions menu XML from the copy

    RibbonCheck = "FAIL"
    Events = Application.EnableEvents
    If Application.OperatingSystem Like "*Mac*" Then
        RibbonCheck = "NOT RUN"
        Detail = "needs Windows PowerShell"
        Exit Function
    End If
    On Error GoTo Failed
    Work = Folder & Application.PathSeparator & "ribbon"
    MkDir Work
    If Not TryFetch(Sha, "src/ribbon/customUI14.xml", Work & "\customUI14.xml", Detail) Then GoTo CleanUp
    CopyPath = Work & "\KPR-ribbon-check." & LCase$(Mid$(Target.Name, InStrRev(Target.Name, ".") + 1))
    Target.SaveCopyAs CopyPath
    Script = Work & "\inject.ps1"
    WriteText Script, InjectorScript()
    ExitCode = CreateObject("WScript.Shell").Run("powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & _
                                              Script & """ -Package """ & CopyPath & """ -CustomUI """ & _
                                              Work & "\customUI14.xml""", 0, True)
    If ExitCode <> 0 Then
        Detail = "PowerShell could not write the tab (exit " & CStr(ExitCode) & ")"
        GoTo CleanUp
    End If
    Application.EnableEvents = False
    Set Copy = Workbooks.Open(CopyPath)
    Started = Timer
    Do
        DoEvents
        Loaded = CBool(Application.Run(QualifiedIn(Copy, "KPR_UI_RibbonLoaded")))
    Loop Until Loaded Or Timer - Started > 10 Or Timer < Started
    Menu = CStr(Application.Run(QualifiedIn(Copy, "KPR_UI_RibbonMenuXml"), "functions", "Date primitives"))
    If Not Loaded Then
        Detail = "Excel did not load the tab (onLoad never ran)"
    ElseIf InStr(1, Menu, "KPR_Fn_", vbBinaryCompare) = 0 Then
        Detail = "the tab loaded but its Functions menu is empty"
    Else
        RibbonCheck = "PASS"
    End If

CleanUp:
    'Close the copy with its events still off, so none of its handlers run,
    'then give back the caller's setting; never leave the copy on disk
        On Error Resume Next
        Application.EnableEvents = False
        If Not Copy Is Nothing Then Copy.Close SaveChanges:=False
        Application.EnableEvents = Events
        Kill CopyPath
        Kill Script
        Kill Work & "\customUI14.xml"
        RmDir Work
        Err.Clear
        Exit Function

Failed:
    Detail = "error " & CStr(Err.Number) & " (" & Err.Description & ")"
    Resume CleanUp

End Function


Private Function InjectorScript() As String
'
' The PowerShell script RibbonCheck runs on the copy: it replaces
' customUI/customUI14.xml, keeps exactly one Office 2010 ribbon relationship
' and makes sure .xml parts have a content type. It is the minimal Windows
' equivalent of tools/inject_ribbon.py, for a throwaway copy only.
'
    Dim S As String     'Script text

    S = "param([string]$Package, [string]$CustomUI)" & vbCrLf
    S = S & "$ErrorActionPreference = 'Stop'" & vbCrLf
    S = S & "Add-Type -AssemblyName System.IO.Compression" & vbCrLf
    S = S & "Add-Type -AssemblyName System.IO.Compression.FileSystem" & vbCrLf
    S = S & "$type = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'" & vbCrLf
    S = S & "$utf8 = New-Object System.Text.UTF8Encoding($false)" & vbCrLf
    S = S & "function ReadEntry($zip, $name) { $r = New-Object System.IO.StreamReader($zip.GetEntry($name).Open()); " & _
            "$t = $r.ReadToEnd(); $r.Close(); return $t }" & vbCrLf
    S = S & "function WriteEntry($zip, $name, $text) { $old = $zip.GetEntry($name); if ($old) { $old.Delete() }; " & _
            "$w = New-Object System.IO.StreamWriter($zip.CreateEntry($name).Open(), $utf8); $w.Write($text); $w.Close() }" & vbCrLf
    S = S & "$zip = [System.IO.Compression.ZipFile]::Open($Package, 'Update')" & vbCrLf
    S = S & "try {" & vbCrLf
    S = S & "  WriteEntry $zip 'customUI/customUI14.xml' ([System.IO.File]::ReadAllText($CustomUI, $utf8))" & vbCrLf
    S = S & "  $rels = ReadEntry $zip '_rels/.rels'" & vbCrLf
    S = S & "  $rels = [regex]::Replace($rels, '<Relationship[^>]*' + [regex]::Escape($type) + '[^>]*/>', '')" & vbCrLf
    S = S & "  $rels = $rels.Replace('</Relationships>', '<Relationship Id=""rIdKprRibbon"" Type=""' + $type + " & _
            "'"" Target=""customUI/customUI14.xml""/></Relationships>')" & vbCrLf
    S = S & "  WriteEntry $zip '_rels/.rels' $rels" & vbCrLf
    S = S & "  $types = ReadEntry $zip '[Content_Types].xml'" & vbCrLf
    S = S & "  if ($types -notmatch 'Extension=""xml""') { $types = $types.Replace('</Types>', " & _
            "'<Default Extension=""xml"" ContentType=""application/xml""/></Types>'); WriteEntry $zip '[Content_Types].xml' $types }" & vbCrLf
    S = S & "} finally { $zip.Dispose() }" & vbCrLf
    InjectorScript = S

End Function


Private Function EvidenceFolder( _
    ByVal Sha As String) _
    As String
'
' Creates and returns Documents\KPR-evidence\<date-time>-<SHA prefix>.
'
    Dim Base        As String       'Documents\KPR-evidence

    On Error Resume Next
    Base = CStr(CreateObject("WScript.Shell").SpecialFolders("MyDocuments"))
    If Len(Base) = 0 Then Base = Environ$("USERPROFILE")
    Base = Base & Application.PathSeparator & "KPR-evidence"
    MkDir Base
    EvidenceFolder = Base & Application.PathSeparator & Format$(Now, "yyyymmdd-hhnnss") & "-" & Left$(Sha, 8)
    MkDir EvidenceFolder
    Err.Clear

End Function


Private Function QualifiedIn( _
    ByVal Wb As Workbook, _
    ByVal ProcedureName As String) _
    As String
'
' The workbook-qualified name Application.Run needs for a procedure of Wb.
'
    QualifiedIn = "'" & Replace(Wb.Name, "'", "''") & "'!" & ProcedureName

End Function


Private Sub WriteText( _
    ByVal File As String, _
    ByVal Text As String)
'
' Writes Text to File, replacing it.
'
    Dim FileNo      As Integer      'Output file

    FileNo = FreeFile
    Open File For Output As #FileNo
    Print #FileNo, Text;
    Close #FileNo

End Sub


Private Function IsSha( _
    ByVal Text As String) _
    As Boolean
'
' TRUE for exactly 40 hexadecimal characters.
'
    Dim I           As Long         'Character cursor

    If Len(Text) <> 40 Then Exit Function
    For I = 1 To 40
        If InStr(1, "0123456789abcdef", LCase$(Mid$(Text, I, 1)), vbBinaryCompare) = 0 Then Exit Function
    Next I
    IsSha = True

End Function


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function TryTarget( _
    ByVal TargetWorkbook As String, _
    ByRef Target As Workbook, _
    ByRef Project As Object, _
    ByRef Detail As String) _
    As Boolean
'
' The open target and its unlocked VBProject; refuses this loader's own
' workbook and explains a missing trust setting.
'
    Dim Count As Long   'Probe of the object model

    TryTarget = False
    On Error Resume Next
    Set Target = Workbooks(TargetWorkbook)
    On Error GoTo 0
    If Target Is Nothing Then
        Detail = TargetWorkbook & " is not open. Open or load it first; use its file name, e.g. KPR.xlam."
        Exit Function
    End If
    If Target Is ThisWorkbook Then
        Detail = "the loader cannot update its own workbook. Keep KPR_Dev_Import in PERSONAL.XLSB " & _
                 "or another development workbook."
        Exit Function
    End If
    On Error Resume Next
    Set Project = Target.VBProject
    Count = Project.VBComponents.Count
    If Err.Number <> 0 Then
        Detail = "Excel refuses access to the VBA project. Turn on File > Options > Trust Center > " & _
                 "Trust Center Settings > Macro Settings > Trust access to the VBA project object model."
        Exit Function
    End If
    On Error GoTo 0
    If Project.Protection <> 0 Then
        Detail = "the VBA project of " & TargetWorkbook & " is locked. Unlock it in the VBA editor first."
        Exit Function
    End If
    TryTarget = True

End Function

Private Function TrySelect( _
    ByVal Profile As String, _
    ByVal Roles As String, _
    ByRef Paths As Collection, _
    ByRef Listed As Collection, _
    ByRef Detail As String) _
    As Boolean
'
' Paths receives the .bas paths of the profile's "components" block whose
' role is in Roles, skipping this loader itself; Listed receives the module
' name of every .bas the block lists, in any role.
'
    Dim Start  As Long      'Start of the components block
    Dim Finish As Long      'End of the components block
    Dim Lines  As Variant   'Lines of the block
    Dim Row    As Variant   'Line cursor
    Dim Parts  As Variant   'Quoted pieces of one line
    Dim Wanted As String    'Roles, comma-delimited for lookup

    TrySelect = False
    Set Paths = New Collection
    Set Listed = New Collection
    Start = InStr(1, Profile, """components"": {", vbBinaryCompare)
    If Start = 0 Then
        Detail = "the repository profile has no components list."
        Exit Function
    End If
    Finish = InStr(Start, Profile, "}", vbBinaryCompare)
    Wanted = "," & LCase$(Replace(Roles, " ", "")) & ","
    Lines = Split(Replace(Mid$(Profile, Start, Finish - Start), vbCr, ""), vbLf)
    For Each Row In Lines
        Parts = Split(CStr(Row), """")
        If UBound(Parts) >= 3 Then
            If LCase$(Right$(Parts(1), 4)) = ".bas" Then
                Listed.Add FileStem(Parts(1))
                If InStr(1, Wanted, "," & LCase$(Parts(3)) & ",") > 0 And _
                   StrComp(FileStem(Parts(1)), "KPR_Dev_Import", vbTextCompare) <> 0 Then Paths.Add Parts(1)
            End If
        End If
    Next Row
    If Paths.Count = 0 Then
        Detail = "no module of the roles " & Roles & " is listed in the profile."
        Exit Function
    End If
    TrySelect = True

End Function

Private Function TryFetch( _
    ByVal Source As String, _
    ByVal RepoPath As String, _
    ByVal Destination As String, _
    ByRef Detail As String) _
    As Boolean
'
' Copies RepoPath of Source (a local clone, or a ref downloaded from GitHub)
' to Destination, byte for byte.
'
    Dim Copied  As String       'Path of the file in a local clone
    Dim Http    As Object       'MSXML2.XMLHTTP request
    Dim Bytes() As Byte         'Downloaded content
    Dim FileNo  As Integer      'Output file

    TryFetch = False
    If IsLocalFolder(Source) Then
        Copied = Source & Application.PathSeparator & Replace(RepoPath, "/", Application.PathSeparator)
        If Len(Dir$(Copied)) = 0 Then
            Detail = "the local clone has no " & RepoPath & "."
            Exit Function
        End If
        FileCopy Copied, Destination
        TryFetch = True
        Exit Function
    End If

    On Error GoTo Failed
    Set Http = CreateObject("MSXML2.XMLHTTP.6.0")
    Http.Open "GET", RAW_BASE & Source & "/" & RepoPath, False
    Http.send
    If Http.Status <> 200 Then
        Detail = "GitHub returned " & CStr(Http.Status) & " for " & RepoPath & " at " & Source & _
                 ". Check the SHA or branch name."
        Exit Function
    End If
    Bytes = Http.responseBody
    FileNo = FreeFile
    Open Destination For Binary Access Write As #FileNo
    Put #FileNo, , Bytes
    Close #FileNo
    TryFetch = True
    Exit Function

Failed:
    Detail = "could not download " & RepoPath & ": error " & CStr(Err.Number) & " (" & Err.Description & ")"

End Function

Private Function TryModuleName( _
    ByVal File As String, _
    ByVal Expected As String, _
    ByRef Detail As String) _
    As Boolean
'
' TRUE when File is a VBA export whose VB_Name is Expected.
'
    Dim Text As String      'File content

    Text = ReadText(File)
    TryModuleName = InStr(1, Text, "Attribute VB_Name = """ & Expected & """", vbBinaryCompare) = 1
    If Not TryModuleName Then
        Detail = FileName(File) & " is not the VBA module " & Expected & "."
    End If

End Function

Private Function Candidates(ByVal Target As Workbook) As Collection
'
' Repository modules that may need removing: KNOWN_MODULES plus every module
' this loader recorded in Target, without duplicates.
'
    Dim Result As Collection    'Names found
    Dim Item   As Variant       'Known-name cursor
    Dim Entry  As Object        'Workbook-name cursor
    Dim Recorded As String      'Recorded module name

    Set Result = New Collection
    For Each Item In Split(KNOWN_MODULES, ",")
        If Not Contains(Result, CStr(Item)) Then Result.Add CStr(Item)
    Next Item
    For Each Entry In Target.Names
        If StrComp(Left$(Entry.Name, Len(RECORD_PREFIX)), RECORD_PREFIX, vbTextCompare) = 0 Then
            Recorded = Mid$(Entry.Name, Len(RECORD_PREFIX) + 1)
            If Not Contains(Result, Recorded) Then Result.Add Recorded
        End If
    Next Entry
    Set Candidates = Result

End Function

Private Sub Remember(ByVal Target As Workbook, ByVal ModuleName As String)
'
' Records in Target, as a hidden workbook name, that this loader imported
' ModuleName. Never raises.
'
    On Error Resume Next
    Target.Names.Add Name:=RECORD_PREFIX & ModuleName, RefersTo:="=TRUE", Visible:=False
    Err.Clear

End Sub

Private Sub Forget(ByVal Target As Workbook, ByVal ModuleName As String)
'
' Deletes the record of ModuleName from Target, if present. Never raises.
'
    On Error Resume Next
    Target.Names(RECORD_PREFIX & ModuleName).Delete
    Err.Clear

End Sub

Private Function RemoveModule(ByVal Project As Object, ByVal ModuleName As String) As Boolean
'
' Removes a standard module of that name, if present. The module is renamed
' first: Excel defers the removal until the macro ends, and an import under
' the old name would otherwise receive a numbered name.
'
    Dim Component As Object     'Existing module

    RemoveModule = False
    On Error Resume Next
    Set Component = Project.VBComponents(ModuleName)
    On Error GoTo 0
    If Component Is Nothing Then Exit Function
    If Component.Type <> STANDARD_MODULE Then Exit Function
    Component.Name = Left$(ModuleName, 24) & "_old" & Format$(Timer * 100 Mod 1000, "000")
    Project.VBComponents.Remove Component
    RemoveModule = True

End Function

Private Function StagingFolder() As String
'
' A new empty folder under %TEMP% for the fetched files.
'
    Dim Base As String      'Temporary folder

    Base = Environ$("TEMP")
    If Len(Base) = 0 Then Base = CurDir$
    StagingFolder = Base & Application.PathSeparator & "kpr-import-" & Format$(Now, "yyyymmddhhnnss")
    MkDir StagingFolder

End Function

Private Sub CleanUp(ByVal Folder As String)
'
' Deletes the staging folder and its files. Never raises.
'
    On Error Resume Next
    If Len(Folder) = 0 Then Exit Sub
    Kill Folder & Application.PathSeparator & "*.*"
    RmDir Folder
    Err.Clear

End Sub

Private Function ReadText(ByVal File As String) As String
'
' The whole file as text, byte for byte.
'
    Dim FileNo As Integer   'Input file

    FileNo = FreeFile
    Open File For Binary Access Read As #FileNo
    ReadText = Space$(LOF(FileNo))
    Get #FileNo, , ReadText
    Close #FileNo

End Function

Private Function IsLocalFolder(ByVal Source As String) As Boolean
'
' TRUE when Source is a folder holding the repository profile.
'
    On Error Resume Next
    IsLocalFolder = Len(Dir$(Source & Application.PathSeparator & ".github" & _
                             Application.PathSeparator & "repository-profile.json")) > 0
    Err.Clear

End Function

Private Function Contains(ByVal Items As Collection, ByVal Wanted As String) As Boolean
'
' TRUE when Items holds Wanted, ignoring case.
'
    Dim Item As Variant     'Cursor

    For Each Item In Items
        If StrComp(CStr(Item), Wanted, vbTextCompare) = 0 Then
            Contains = True
            Exit Function
        End If
    Next Item

End Function

Private Function FileName(ByVal Path As String) As String
'
' The last part of a path, with either separator.
'
    Dim Cut As Long     'Last separator

    Cut = InStrRev(Replace(Path, "\", "/"), "/")
    FileName = Mid$(Path, Cut + 1)

End Function

Private Function FileStem(ByVal Path As String) As String
'
' The file name without its extension.
'
    Dim Leaf As String  'File name

    Leaf = FileName(Path)
    FileStem = Left$(Leaf, InStrRev(Leaf, ".") - 1)

End Function
