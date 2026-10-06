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
'   Only standard modules named in the profile, of the selected roles, plus
'   the retired names listed below. Every other module of the target (for
'   example private ribbon or menu modules) is left alone. All files are
'   fetched and checked before the target is touched, so a failed download
'   changes nothing. The target is not saved: compile it, then save it.
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
'   Option Private Module. Development infrastructure, called from the
'   Immediate window.
'
' ALLOWED DEPENDENCIES
'   None.
'
' UPDATED
'   2026-10-06
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit         'Force explicit variable declarations
    Option Private Module   'Development infrastructure: invisible outside this project

'------------------------------------------------------------------------------
' CONSTANTS
'------------------------------------------------------------------------------
    'Raw-file address of the repository; the ref and path are appended
        Private Const RAW_BASE As String = "https://raw.githubusercontent.com/danielep71/k-pricing/"

    'Profile that lists every module with its role
        Private Const PROFILE_PATH As String = ".github/repository-profile.json"

    'Module names the repository no longer uses; removed from the target when
    'the selected commit does not import a module of the same name
        Private Const RETIRED_MODULES As String = "KPR_Demo_Builder,KPR__Demo_Builder,KPR_DateExample"

    'Default roles to import
        Private Const DEFAULT_ROLES As String = "public,internal,test,example"

    'VBComponent type of a standard module
        Private Const STANDARD_MODULE As Long = 1

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
'     Optional. Comma-separated profile roles to import; default all four.
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
    If Not TrySelect(Profile, Roles, Paths, Detail) Then GoTo Refuse

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
    Stage = "remove retired modules"
    For Each Item In Split(RETIRED_MODULES, ",")
        If Not Contains(Names, CStr(Item)) Then
            If RemoveModule(Project, CStr(Item)) Then
                Removed = Removed + 1
                Debug.Print "  removed  " & CStr(Item) & " (retired)"
            End If
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
        Debug.Print "  imported " & Names(K) & "  (" & Paths(K) & ")"
    Next K

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    Debug.Print "KPR_Dev_ImportModules: " & CStr(Names.Count) & " module(s) imported and " & _
                CStr(Removed) & " retired module(s) removed in " & Target.Name & " from " & Source & "."
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
    ByRef Detail As String) _
    As Boolean
'
' The .bas paths of the profile's "components" block whose role is in Roles,
' skipping this loader itself.
'
    Dim Start  As Long      'Start of the components block
    Dim Finish As Long      'End of the components block
    Dim Lines  As Variant   'Lines of the block
    Dim Row    As Variant   'Line cursor
    Dim Parts  As Variant   'Quoted pieces of one line
    Dim Wanted As String    'Roles, comma-delimited for lookup

    TrySelect = False
    Set Paths = New Collection
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
            If LCase$(Right$(Parts(1), 4)) = ".bas" And InStr(1, Wanted, "," & LCase$(Parts(3)) & ",") > 0 Then
                If StrComp(FileStem(Parts(1)), "KPR_Dev_Import", vbTextCompare) <> 0 Then Paths.Add Parts(1)
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
