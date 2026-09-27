Attribute VB_Name = "KPR_REGISTER_PUBLIC_UDFS"
'==============================================================================
' MODULE: KPR_REGISTER_PUBLIC_UDFS
'------------------------------------------------------------------------------
' PURPOSE
'   Function Wizard registration for the 22 supported KPR_Dates_* worksheet
'   functions: one manifest, one Excel category and an explicit, repeatable
'   register / clean-up lifecycle.
'
' WHY THIS EXISTS
'   Application.MacroOptions gives each worksheet function a description, a
'   Function Wizard category and argument help. Keeping that metadata here
'   keeps descriptions and lifecycle code out of the calculation modules, and
'   gives the supported surface exactly one manifest that
'   tools/check_kpr_contract.py compares with docs/PUBLIC_API.txt.
'
' SCOPE
'   - KPR_Register_PublicUDFs           register every manifest record
'   - KPR_Register_ClearPublicUDFs      best-effort metadata clean-up
'   - KPR_Register_LastReport           per-function failures of the last run
'   - KPR_Register_ManifestCount        manifest accessors for the regression
'   - KPR_Register_ManifestName           harness and later tooling
'   - KPR_Register_ManifestArgCount
'   - KPR_Register_ManifestArgDescriptions
'
' MANIFEST
'   LoadManifest holds one AddRecord statement per supported function: the
'   function name, its argument names in signature order separated by "|",
'   the Function Wizard description, then one argument description per
'   argument. Each argument description begins with its argument name, and an
'   optional argument's with "<name> (optional, default ...)". The static gate
'   reads these records as data; it never interprets the prose.
'
'   Every function is registered in the single category KPR_UDF_CATEGORY,
'   "KPR Dates". Argument descriptions are passed as a complete 1-based array
'   matching the signature. KPR_Dates_HostDateSystem takes no argument, so its
'   registration omits ArgumentDescriptions entirely.
'
' LIFECYCLE
'   Registration overwrites the metadata of each function by name, so running
'   it again is idempotent and creates no duplicate category or entry.
'   Excel has no operation that unregisters a VBA function: clean-up can only
'   blank the description and argument help and move each function back to
'   Excel's built-in "User Defined" category. The functions themselves remain
'   callable for as long as this VBA project is loaded.
'
'   Excel offers no way to read MacroOptions metadata back, so each entry
'   point reports only whether Excel accepted every call.
'
' VISIBILITY
'   Option Private Module. The entry points are Public so a workbook or add-in
'   event, Application.Run or a UI callback can call them; the module-level
'   Private setting keeps them out of the worksheet function list and the
'   macro list. They are unsupported infrastructure, not supported API.
'
' ERROR POLICY
'   Registration is a usability feature, never a calculation dependency. Each
'   function is registered independently; a failure is recorded for
'   KPR_Register_LastReport and the remaining functions are still attempted.
'   No entry point raises to its caller.
'
' STATE
'   MacroOptions marks ThisWorkbook as changed. Each entry point restores the
'   Saved flag it found, and none selects, activates, recalculates or changes
'   alerts, events or screen updating.
'
' INTEGRATION
'   Call KPR_Register_PublicUDFs once from Workbook_Open (or an add-in's
'   startup), from the Immediate window, or through
'   Application.Run "'<workbook>'!KPR_Register_PublicUDFs". It returns the
'   number of functions Excel registered: 22 on success.
'
' ALLOWED DEPENDENCIES
'   No KPR module: function names appear only as manifest text. Excel's
'   Application.MacroOptions and ThisWorkbook.
'
' UPDATED
'   2026-09-27
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit         'Force explicit variable declarations
    Option Private Module   'Infrastructure: invisible outside this VBA project

'------------------------------------------------------------------------------
' CONSTANTS
'------------------------------------------------------------------------------
    'The single Function Wizard category of the supported surface
        Private Const KPR_UDF_CATEGORY As String = "KPR Dates"
    'Excel's built-in "User Defined" category, the clean-up destination
        Private Const USER_DEFINED_CATEGORY As Long = 14

'------------------------------------------------------------------------------
' MANIFEST STATE
'------------------------------------------------------------------------------
    Private mCount          As Long         'Manifest records loaded
    Private mNames()        As String       'Function names, 1-based
    Private mArgNames()     As String       'Argument names joined by "|", 1-based
    Private mDescriptions() As String       'Function Wizard descriptions, 1-based
    Private mArgDescs()     As Variant      'Per record: 1-based argument descriptions or Empty
    Private mArgCounts()    As Long         'Per record: number of arguments
    Private mLastReport     As String       'Failures of the last register or clean-up run

'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Function KPR_Register_PublicUDFs() As Long
'
'==============================================================================
'                           KPR_Register_PublicUDFs
'------------------------------------------------------------------------------
' PURPOSE
'   Registers the description, category and argument help of every supported
'   function with Application.MacroOptions.
'
' RETURNS
'   The number of functions Excel accepted: 22 when all succeed. Failures are
'   listed by KPR_Register_LastReport.
'
' BEHAVIOR
'   Idempotent: each run overwrites the same metadata by function name.
'   ThisWorkbook.Saved is restored to the value found on entry.
'
' UPDATED
'   2026-09-27
'==============================================================================
'
    KPR_Register_PublicUDFs = ApplyAll(False)

End Function

Public Function KPR_Register_ClearPublicUDFs() As Long
'
'==============================================================================
'                         KPR_Register_ClearPublicUDFs
'------------------------------------------------------------------------------
' PURPOSE
'   Best-effort removal of the Function Wizard metadata this module registers.
'
' RETURNS
'   The number of functions whose metadata Excel accepted a reset for: 22
'   when all succeed. Failures are listed by KPR_Register_LastReport.
'
' BEHAVIOR
'   Blanks each description and argument description and moves the function
'   to Excel's built-in "User Defined" category. Excel exposes no symmetric
'   unregister operation: the VBA functions are not removed and stay callable
'   while this project is loaded. Repeating the clean-up is harmless.
'
' UPDATED
'   2026-09-27
'==============================================================================
'
    KPR_Register_ClearPublicUDFs = ApplyAll(True)

End Function

Public Function KPR_Register_LastReport() As String
'
' One line per function that Excel rejected during the last register or
' clean-up run; empty when every call succeeded.
'
    KPR_Register_LastReport = mLastReport

End Function

Public Function KPR_Register_ManifestCount() As Long
'
' Number of manifest records.
'
    LoadManifest
    KPR_Register_ManifestCount = mCount

End Function

Public Function KPR_Register_ManifestName( _
    ByVal Index As Long) _
    As String
'
' Function name of manifest record Index (1-based); empty outside the manifest.
'
    LoadManifest
    If Index >= 1 And Index <= mCount Then KPR_Register_ManifestName = mNames(Index)

End Function

Public Function KPR_Register_ManifestArgCount( _
    ByVal Index As Long) _
    As Long
'
' Argument count of manifest record Index (1-based); -1 outside the manifest.
'
    LoadManifest
    If Index >= 1 And Index <= mCount Then
        KPR_Register_ManifestArgCount = mArgCounts(Index)
    Else
        KPR_Register_ManifestArgCount = -1
    End If

End Function

Public Function KPR_Register_ManifestArgDescriptions( _
    ByVal Index As Long) _
    As Variant
'
' Copy of the 1-based argument-description array of manifest record Index;
' Empty for a function without arguments or outside the manifest.
'
    LoadManifest
    If Index >= 1 And Index <= mCount Then KPR_Register_ManifestArgDescriptions = mArgDescs(Index)

End Function

'
'------------------------------------------------------------------------------
'
'                                   LIFECYCLE
'
'------------------------------------------------------------------------------
'

Private Function ApplyAll( _
    ByVal Clear As Boolean) _
    As Long
'
' Registers (Clear = False) or resets (Clear = True) every manifest record and
' returns how many Excel accepted. ThisWorkbook.Saved is restored on exit.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WasSaved        As Boolean      'ThisWorkbook.Saved found on entry
    Dim HaveSaved       As Boolean      'TRUE once WasSaved was read
    Dim Accepted        As Long         'Records Excel accepted
    Dim I               As Long         'Record cursor
    Dim Why             As String       'Failure description of one record

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo Apply_Error
    mLastReport = vbNullString
    LoadManifest
    WasSaved = ThisWorkbook.Saved
    HaveSaved = True

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
    'Each record independently, so one rejection never blocks the others
        For I = 1 To mCount
            If TryApply(I, Clear, Why) Then
                Accepted = Accepted + 1
            Else
                Note mNames(I) & ": " & Why
            End If
        Next I

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
Apply_Exit:
    'MacroOptions marks the workbook as changed; put back what the caller had
        If HaveSaved Then
            On Error Resume Next
            If WasSaved And Not ThisWorkbook.Saved Then ThisWorkbook.Saved = True
            If Err.Number <> 0 Then Note "ThisWorkbook.Saved: error " & CStr(Err.Number) & ": " & Err.Description
            Err.Clear
            On Error GoTo 0
        End If
    ApplyAll = Accepted
    Exit Function

Apply_Error:
    Note "lifecycle: error " & CStr(Err.Number) & ": " & Err.Description
    Resume Apply_Exit

End Function

Private Function TryApply( _
    ByVal Index As Long, _
    ByVal Clear As Boolean, _
    ByRef Why As String) _
    As Boolean
'
' Applies one record. The workbook-qualified macro name is tried first so the
' call cannot resolve to a same-named macro in another open workbook; the bare
' name is the fallback Excel accepts when this project is the active one.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrQualified    As String       'Outcome of the workbook-qualified call
    Dim ErrBare         As String       'Outcome of the bare-name call

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
    Why = vbNullString
    ErrQualified = CallMacroOptions(QualifiedName(mNames(Index)), Index, Clear)
    If Len(ErrQualified) = 0 Then
        TryApply = True
        Exit Function
    End If
    ErrBare = CallMacroOptions(mNames(Index), Index, Clear)
    If Len(ErrBare) = 0 Then
        TryApply = True
    Else
        Why = ErrQualified & " (bare name: " & ErrBare & ")"
    End If

End Function

Private Function CallMacroOptions( _
    ByVal Target As String, _
    ByVal Index As Long, _
    ByVal Clear As Boolean) _
    As String
'
' One Application.MacroOptions call; returns "" on success, otherwise the
' error Excel raised. ArgumentDescriptions is omitted for a function without
' arguments, as an empty array is rejected by some Excel builds.
'

'------------------------------------------------------------------------------
' CALL
'------------------------------------------------------------------------------
    On Error GoTo Call_Error
    If Clear Then
        If mArgCounts(Index) > 0 Then
            Application.MacroOptions Macro:=Target, Description:=vbNullString, _
                                     Category:=USER_DEFINED_CATEGORY, _
                                     ArgumentDescriptions:=BlankDescriptions(mArgCounts(Index))
        Else
            Application.MacroOptions Macro:=Target, Description:=vbNullString, _
                                     Category:=USER_DEFINED_CATEGORY
        End If
    Else
        If mArgCounts(Index) > 0 Then
            Application.MacroOptions Macro:=Target, Description:=mDescriptions(Index), _
                                     Category:=KPR_UDF_CATEGORY, _
                                     ArgumentDescriptions:=mArgDescs(Index)
        Else
            Application.MacroOptions Macro:=Target, Description:=mDescriptions(Index), _
                                     Category:=KPR_UDF_CATEGORY
        End If
    End If
    Exit Function

Call_Error:
    CallMacroOptions = "error " & CStr(Err.Number) & ": " & Err.Description
    Err.Clear

End Function

Private Function QualifiedName( _
    ByVal FunctionName As String) _
    As String
'
' 'Workbook name'!FunctionName, with any apostrophe in the name doubled.
'
    QualifiedName = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & FunctionName

End Function

Private Function BlankDescriptions( _
    ByVal Total As Long) _
    As Variant
'
' A 1-based array of Total empty argument descriptions for the clean-up.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Blank()         As Variant      'Result
    Dim I               As Long         'Element cursor

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
    ReDim Blank(1 To Total)
    For I = 1 To Total
        Blank(I) = vbNullString
    Next I
    BlankDescriptions = Blank

End Function

Private Sub Note( _
    ByVal Entry As String)
'
' Appends one line to the report of the current run.
'
    If Len(mLastReport) > 0 Then mLastReport = mLastReport & vbCrLf
    mLastReport = mLastReport & Entry

End Sub

'
'------------------------------------------------------------------------------
'
'                                   MANIFEST
'
'------------------------------------------------------------------------------
'

Private Sub AddRecord( _
    ByVal FunctionName As String, _
    ByVal ArgNames As String, _
    ByVal Description As String, _
    ParamArray ArgDescriptions() As Variant)
'
' Appends one manifest record. The argument descriptions arrive as a 0-based
' ParamArray and are stored as the 1-based array MacroOptions receives.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Descs()         As Variant      'Argument descriptions, 1-based
    Dim Total           As Long         'Number of argument descriptions
    Dim I               As Long         'Element cursor

'------------------------------------------------------------------------------
' STORE
'------------------------------------------------------------------------------
    mCount = mCount + 1
    ReDim Preserve mNames(1 To mCount)
    ReDim Preserve mArgNames(1 To mCount)
    ReDim Preserve mDescriptions(1 To mCount)
    ReDim Preserve mArgDescs(1 To mCount)
    ReDim Preserve mArgCounts(1 To mCount)
    mNames(mCount) = FunctionName
    mArgNames(mCount) = ArgNames
    mDescriptions(mCount) = Description
    Total = UBound(ArgDescriptions) - LBound(ArgDescriptions) + 1
    mArgCounts(mCount) = Total
    If Total > 0 Then
        ReDim Descs(1 To Total)
        For I = 1 To Total
            Descs(I) = CStr(ArgDescriptions(LBound(ArgDescriptions) + I - 1))
        Next I
        mArgDescs(mCount) = Descs
    Else
        mArgDescs(mCount) = Empty
    End If

End Sub

Private Sub LoadManifest()
'
' The supported worksheet surface: exactly one record per KPR_Dates_* function
' in docs/PUBLIC_API.txt, in contract order. tools/check_kpr_contract.py
' (kpr-registration-manifest) checks the names, argument names and order,
' description lengths and category against the public signatures.
'

'------------------------------------------------------------------------------
' RESET
'------------------------------------------------------------------------------
    mCount = 0
    Erase mNames
    Erase mArgNames
    Erase mDescriptions
    Erase mArgDescs
    Erase mArgCounts

'------------------------------------------------------------------------------
' RECORDS
'------------------------------------------------------------------------------
    'Days, months, quarters and years
        AddRecord "KPR_Dates_DayOfWeek", "DateIn|Opt_WeekBaseMonday", _
            "Returns the weekday number of DateIn, 1 to 7; by default Monday = 1, unlike WEEKDAY. Scalar input gives one value; a range or array spills one result per cell (dynamic-array Excel). Invalid input: #VALUE!; date out of range: #NUM!.", _
            "DateIn: a date, date serial or ISO text yyyy-mm-dd from 1900-03-01 to 9999-12-31. A range or array gives one result per cell (dynamic-array Excel). Other text gives #VALUE!; a date outside the window gives #NUM!.", _
            "Opt_WeekBaseMonday (optional, default TRUE): TRUE numbers Monday=1 to Sunday=7; FALSE numbers Sunday=1 to Saturday=7. Must be a single TRUE or FALSE, otherwise #VALUE!."
        AddRecord "KPR_Dates_DaysInMonth", "DateIn", _
            "Returns the number of days, 28 to 31, in the month containing DateIn. Scalar input gives one value; a range or array spills one result per cell (dynamic-array Excel). Invalid input: #VALUE!; date out of range: #NUM!.", _
            "DateIn: a date, date serial or ISO text yyyy-mm-dd from 1900-03-01 to 9999-12-31. A range or array gives one result per cell (dynamic-array Excel). Other text gives #VALUE!; a date outside the window gives #NUM!."
        AddRecord "KPR_Dates_DaysInYear", "YearIn", _
            "Returns 366 for a Gregorian leap year and 365 otherwise. Takes a year, not a date. Scalar input gives one value; a range or array spills one result per cell (dynamic-array Excel). Invalid input: #VALUE!; date out of range: #NUM!.", _
            "YearIn: a calendar year as a whole number, 1900 to 9999; not a date (use YEAR(date)). A range or array gives one result per cell (dynamic-array Excel). Other values give #VALUE!."
        AddRecord "KPR_Dates_BeginOfMonth", "DateIn", _
            "Returns the first day of the month containing DateIn. Scalar input gives one value; a range or array spills one result per cell (dynamic-array Excel). Invalid input: #VALUE!; date out of range: #NUM!.", _
            "DateIn: a date, date serial or ISO text yyyy-mm-dd from 1900-03-01 to 9999-12-31. A range or array gives one result per cell (dynamic-array Excel). Other text gives #VALUE!; a date outside the window gives #NUM!."
        AddRecord "KPR_Dates_EndOfMonth", "DateIn", _
            "Returns the last day of the month containing DateIn. Scalar input gives one value; a range or array spills one result per cell (dynamic-array Excel). Invalid input: #VALUE!; date out of range: #NUM!.", _
            "DateIn: a date, date serial or ISO text yyyy-mm-dd from 1900-03-01 to 9999-12-31. A range or array gives one result per cell (dynamic-array Excel). Other text gives #VALUE!; a date outside the window gives #NUM!."
        AddRecord "KPR_Dates_BeginOfQuarter", "DateIn", _
            "Returns the first day of the calendar quarter containing DateIn. Scalar input gives one value; a range or array spills one result per cell (dynamic-array Excel). Invalid input: #VALUE!; date out of range: #NUM!.", _
            "DateIn: a date, date serial or ISO text yyyy-mm-dd from 1900-03-01 to 9999-12-31. A range or array gives one result per cell (dynamic-array Excel). Other text gives #VALUE!; a date outside the window gives #NUM!."
        AddRecord "KPR_Dates_EndOfQuarter", "DateIn", _
            "Returns the last day of the calendar quarter containing DateIn. Scalar input gives one value; a range or array spills one result per cell (dynamic-array Excel). Invalid input: #VALUE!; date out of range: #NUM!.", _
            "DateIn: a date, date serial or ISO text yyyy-mm-dd from 1900-03-01 to 9999-12-31. A range or array gives one result per cell (dynamic-array Excel). Other text gives #VALUE!; a date outside the window gives #NUM!."
        AddRecord "KPR_Dates_BeginOfYear", "DateIn", _
            "Returns 1 January of the year containing DateIn. Scalar input gives one value; a range or array spills one result per cell (dynamic-array Excel). Invalid input: #VALUE!; date out of range: #NUM!.", _
            "DateIn: a date, date serial or ISO text yyyy-mm-dd from 1900-03-01 to 9999-12-31. A range or array gives one result per cell (dynamic-array Excel). Other text gives #VALUE!; a date outside the window gives #NUM!."
        AddRecord "KPR_Dates_EndOfYear", "DateIn", _
            "Returns 31 December of the year containing DateIn. Scalar input gives one value; a range or array spills one result per cell (dynamic-array Excel). Invalid input: #VALUE!; date out of range: #NUM!.", _
            "DateIn: a date, date serial or ISO text yyyy-mm-dd from 1900-03-01 to 9999-12-31. A range or array gives one result per cell (dynamic-array Excel). Other text gives #VALUE!; a date outside the window gives #NUM!."
        AddRecord "KPR_Dates_IsMonthEnd", "DateIn", _
            "Returns TRUE if DateIn is the last day of its month, otherwise FALSE. Scalar input gives one value; a range or array spills one result per cell (dynamic-array Excel). Invalid input: #VALUE!; date out of range: #NUM!.", _
            "DateIn: a date, date serial or ISO text yyyy-mm-dd from 1900-03-01 to 9999-12-31. A range or array gives one result per cell (dynamic-array Excel). Other text gives #VALUE!; a date outside the window gives #NUM!."
        AddRecord "KPR_Dates_IsQuarterEnd", "DateIn", _
            "Returns TRUE if DateIn is 31 Mar, 30 Jun, 30 Sep or 31 Dec. Scalar input gives one value; a range or array spills one result per cell (dynamic-array Excel). Invalid input: #VALUE!; date out of range: #NUM!.", _
            "DateIn: a date, date serial or ISO text yyyy-mm-dd from 1900-03-01 to 9999-12-31. A range or array gives one result per cell (dynamic-array Excel). Other text gives #VALUE!; a date outside the window gives #NUM!."
        AddRecord "KPR_Dates_IsYearEnd", "DateIn", _
            "Returns TRUE if DateIn is 31 December, otherwise FALSE. Scalar input gives one value; a range or array spills one result per cell (dynamic-array Excel). Invalid input: #VALUE!; date out of range: #NUM!.", _
            "DateIn: a date, date serial or ISO text yyyy-mm-dd from 1900-03-01 to 9999-12-31. A range or array gives one result per cell (dynamic-array Excel). Other text gives #VALUE!; a date outside the window gives #NUM!."
        AddRecord "KPR_Dates_IsLeapYear", "YearIn", _
            "Returns TRUE if YearIn is a Gregorian leap year. Takes a year, not a date. Scalar input gives one value; a range or array spills one result per cell (dynamic-array Excel). Invalid input: #VALUE!; date out of range: #NUM!.", _
            "YearIn: a calendar year as a whole number, 1900 to 9999; not a date (use YEAR(date)). A range or array gives one result per cell (dynamic-array Excel). Other values give #VALUE!."

    'Date arithmetic
        AddRecord "KPR_Dates_AddDays", "DateIn|nDays", _
            "Adds nDays calendar days to DateIn. Scalar input gives one value; a range or array spills one result per cell (dynamic-array Excel). Invalid input: #VALUE!; date out of range: #NUM!.", _
            "DateIn: a date, date serial or ISO text yyyy-mm-dd from 1900-03-01 to 9999-12-31. A range or array gives one result per cell (dynamic-array Excel). Other text gives #VALUE!; a date outside the window gives #NUM!.", _
            "nDays: whole number of calendar days to add; negative values move back. A range or array gives one result per cell (dynamic-array Excel). Fractions or text give #VALUE!."
        AddRecord "KPR_Dates_AddWeeks", "DateIn|nWeeks", _
            "Adds nWeeks weeks of 7 days to DateIn. Scalar input gives one value; a range or array spills one result per cell (dynamic-array Excel). Invalid input: #VALUE!; date out of range: #NUM!.", _
            "DateIn: a date, date serial or ISO text yyyy-mm-dd from 1900-03-01 to 9999-12-31. A range or array gives one result per cell (dynamic-array Excel). Other text gives #VALUE!; a date outside the window gives #NUM!.", _
            "nWeeks: whole number of weeks to add; negative values move back. A range or array gives one result per cell (dynamic-array Excel). Fractions or text give #VALUE!."
        AddRecord "KPR_Dates_AddMonths", "DateIn|nMonths|Opt_KeepEOM", _
            "Adds nMonths months to DateIn, clipping to a shorter month's last day; Opt_KeepEOM keeps month-ends. Scalar input gives one value; a range or array spills one result per cell (dynamic-array Excel). Invalid input: #VALUE!; date out of range: #NUM!.", _
            "DateIn: a date, date serial or ISO text yyyy-mm-dd from 1900-03-01 to 9999-12-31. A range or array gives one result per cell (dynamic-array Excel). Other text gives #VALUE!; a date outside the window gives #NUM!.", _
            "nMonths: whole number of calendar months to add; negative values move back. A range or array gives one result per cell (dynamic-array Excel). Fractions or text give #VALUE!.", _
            "Opt_KeepEOM (optional, default FALSE): FALSE keeps the day, clipped to the target month's length; TRUE maps a month-end input to the target month-end. Must be a single TRUE or FALSE."
        AddRecord "KPR_Dates_AddYears", "DateIn|nYears|Opt_KeepEOM", _
            "Adds nYears years to DateIn with the month-end rules of AddMonths (29 Feb clips to 28 Feb). Scalar input gives one value; a range or array spills one result per cell (dynamic-array Excel). Invalid input: #VALUE!; date out of range: #NUM!.", _
            "DateIn: a date, date serial or ISO text yyyy-mm-dd from 1900-03-01 to 9999-12-31. A range or array gives one result per cell (dynamic-array Excel). Other text gives #VALUE!; a date outside the window gives #NUM!.", _
            "nYears: whole number of years to add; negative values move back. A range or array gives one result per cell (dynamic-array Excel). Fractions or text give #VALUE!.", _
            "Opt_KeepEOM (optional, default FALSE): FALSE keeps the day, clipped to the target month's length; TRUE maps a month-end input to the target month-end. Must be a single TRUE or FALSE."

    'Weekday locators
        AddRecord "KPR_Dates_NthWeekdayOfMonth", "YearIn|MonthIn|WdIndex|n|Opt_WeekBaseMonday", _
            "Returns occurrence n of weekday WdIndex in MonthIn of YearIn; #NUM! if it does not exist. Scalar input gives one value; a range or array spills one result per cell (dynamic-array Excel). Invalid input: #VALUE!; date out of range: #NUM!.", _
            "YearIn: a calendar year as a whole number, 1900 to 9999; not a date (use YEAR(date)). A range or array gives one result per cell (dynamic-array Excel). Other values give #VALUE!.", _
            "MonthIn: month number, 1 to 12. A range or array gives one result per cell (dynamic-array Excel). Other values give #VALUE!.", _
            "WdIndex: weekday number, 1 to 7, counted as Opt_WeekBaseMonday selects (default Monday=1). A range or array gives one result per cell (dynamic-array Excel).", _
            "n: occurrence, 1 to 5 (1 = first). A range or array gives one result per cell (dynamic-array Excel). Values outside 1 to 5 give #VALUE!.", _
            "Opt_WeekBaseMonday (optional, default TRUE): TRUE numbers Monday=1 to Sunday=7; FALSE numbers Sunday=1 to Saturday=7. Must be a single TRUE or FALSE, otherwise #VALUE!."
        AddRecord "KPR_Dates_LastWeekdayOfMonth", "YearIn|MonthIn|WdIndex|Opt_WeekBaseMonday", _
            "Returns the last occurrence of weekday WdIndex in MonthIn of YearIn. Scalar input gives one value; a range or array spills one result per cell (dynamic-array Excel). Invalid input: #VALUE!; date out of range: #NUM!.", _
            "YearIn: a calendar year as a whole number, 1900 to 9999; not a date (use YEAR(date)). A range or array gives one result per cell (dynamic-array Excel). Other values give #VALUE!.", _
            "MonthIn: month number, 1 to 12. A range or array gives one result per cell (dynamic-array Excel). Other values give #VALUE!.", _
            "WdIndex: weekday number, 1 to 7, counted as Opt_WeekBaseMonday selects (default Monday=1). A range or array gives one result per cell (dynamic-array Excel).", _
            "Opt_WeekBaseMonday (optional, default TRUE): TRUE numbers Monday=1 to Sunday=7; FALSE numbers Sunday=1 to Saturday=7. Must be a single TRUE or FALSE, otherwise #VALUE!."

    'Pillars
        AddRecord "KPR_Dates_PillarFromDates", "StartDate|EndDate|Opt_Rounding", _
            "Returns the pillar from StartDate to EndDate, such as 3D, 2W, 6M or 1Y3M, rounded per Opt_Rounding. Scalar input gives one value; a range or array spills one result per cell (dynamic-array Excel). Invalid input: #VALUE!; date out of range: #NUM!.", _
            "StartDate: anchor date; a date, date serial or ISO text yyyy-mm-dd. A range or array gives one result per cell (dynamic-array Excel).", _
            "EndDate: target date; a date, date serial or ISO text yyyy-mm-dd. The result is a pillar such as 0D, 3D, 2W, 6M, 1Y3M or -1Y. A range or array gives one result per cell (dynamic-array Excel).", _
            "Opt_Rounding (optional, default ""NEAREST""): ""NEAREST"", ""FLOOR"" or ""CEILING"", case-insensitive; selects how an interval between anchors is rounded. Other values give #VALUE!."
        AddRecord "KPR_Dates_DateFromPillar", "StartDate|Pillar", _
            "Returns StartDate moved by a pillar token such as ON, 1W, 3M, 2Y6M or -1D. Scalar input gives one value; a range or array spills one result per cell (dynamic-array Excel). Invalid input: #VALUE!; date out of range: #NUM!.", _
            "StartDate: anchor date; a date, date serial or ISO text yyyy-mm-dd. A range or array gives one result per cell (dynamic-array Excel).", _
            "Pillar: token such as ON, TN, 1W, 3M, 2Y6M or -1D; units Y, M, W and D, each at most once, no spaces. A range or array gives one result per cell (dynamic-array Excel). A malformed token gives #VALUE!."

    'Host diagnostic
        AddRecord "KPR_Dates_HostDateSystem", "", _
            "Returns 1900 or 1904, the date system of the calling workbook (1900 when called from VBA). The other KPR functions return #N/A in a 1904 workbook. Scalar only; volatile."

End Sub
