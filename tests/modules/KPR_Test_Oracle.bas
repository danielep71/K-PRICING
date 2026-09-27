Attribute VB_Name = "KPR_Test_Oracle"
'==============================================================================
' MODULE: KPR_Test_Oracle
'------------------------------------------------------------------------------
' PURPOSE
'   Cross-checks the KPR date primitives against native Excel worksheet
'   functions where the two contracts genuinely overlap (issue #41). Excel is
'   an independent oracle only on that overlap; it is never authoritative
'   where its behaviour differs from docs/DATE_LAYER_CONTRACT.md.
'
' ORACLE FUNCTIONS
'   EOMONTH, EDATE, WEEKDAY, DAY, YEAR and MONTH, evaluated with
'   Worksheet.Evaluate on a scratch workbook whose date system is 1900.
'   WORKDAY.INTL and NETWORKDAYS.INTL are not used; check_kpr_contract.py
'   rejects any other worksheet function name in this module.
'
' OVERLAP ASSUMPTIONS
'   Each assertion's label states the identity it tests, for example
'   "EndOfMonth(d) = EOMONTH(d,0)". Inputs are whole-day serials inside the
'   supported window 1900-03-01 .. 9999-12-31, passed to KPR as native VBA
'   Dates under the direct-VBA 1900 contract.
'
' DOCUMENTED EXCLUSIONS (asserted, never skipped)
'   - A result Excel places outside the supported window, or refuses: the
'     contract requires #NUM!, so the assertion expects #NUM! instead of the
'     Excel value.
'   - Year 1900 leap-year identities: Excel's February 1900 has a fictitious
'     29th day, so IsLeapYear(1900) and DaysInYear(1900) are asserted against
'     the Gregorian answers FALSE and 365 rather than against Excel.
'   - March 1900 month start: Excel's EOMONTH(d,-1)+1 does not give
'     1-Mar-1900 across the fictitious 29-Feb-1900, so BeginOfMonth is
'     asserted against d-DAY(d)+1 there, and the weekday locators always
'     take the month start as d-DAY(d)+1.
'   - Not compared at all: text or coerced inputs (Excel parses permissively,
'     KPR strictly), fractional serials (KPR normalizes to the date), serials
'     before 1900-03-01, and the 1904 date system.
'
' SAMPLES
'   Fixed boundary dates, each with two parameter sets, then ORACLE_SAMPLES
'   deterministic samples from a Park-Miller generator seeded with
'   ORACLE_SEED. Case order and values are identical on every run.
'
' PUBLIC SURFACE (test infrastructure, not supported API)
'   KPR_Oracle_RunCases     runs every oracle case into a caller's counters
'
' DEPENDENCIES
'   KPR_DATES_DAYS only. No core module, no fixture module.
'
' STATE OWNERSHIP
'   Creates one scratch workbook, sets manual calculation for the run, and
'   restores calculation and closes the workbook with SaveChanges:=False on
'   every exit path. Macro-only: it adds and closes a workbook.
'
' UPDATED
'   2026-09-26
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit         'Force explicit variable declarations
    Option Private Module   'Visible to this VBA project only

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Deterministic sampling
        Private Const ORACLE_SEED       As Double = 20260926#   'Park-Miller seed
        Private Const ORACLE_SAMPLES    As Long = 300           'Generated samples
        Private Const PM_MODULUS        As Double = 2147483647# 'Park-Miller modulus
        Private Const PM_MULTIPLIER     As Double = 16807#      'Park-Miller multiplier

    'Supported window as 1900 serials, and the native error it implies
        Private Const MIN_SERIAL        As Long = 61            '1900-03-01
        Private Const MAX_SERIAL        As Long = 2958465       '9999-12-31
        Private Const ERR_NUM           As Long = 2036          '#NUM!

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'Scratch sheet the oracle formulas are evaluated on
        Private mSheet          As Worksheet

    'Caller's failure collection and this run's assertion count
        Private mFailures       As Collection
        Private mChecks         As Long

    'Current Park-Miller state
        Private mSeed           As Double


'
'------------------------------------------------------------------------------
'
'                                  ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Sub KPR_Oracle_RunCases( _
    ByRef Checks As Long, _
    ByVal Failures As Collection)
'
'==============================================================================
'                              KPR_Oracle_RunCases
'------------------------------------------------------------------------------
' PURPOSE
'   Runs every oracle case, adds its assertion count to Checks and appends
'   each failure to Failures as Array(label, detail).
'
' UPDATED
'   2026-09-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Scratch         As Workbook     'The scratch workbook, held by reference
    Dim PriorCalc       As XlCalculation 'Caller's calculation mode
    Dim CalcChanged     As Boolean      'TRUE once PriorCalc was captured
    Dim Boundaries      As Variant      'Fixed boundary dates
    Dim I               As Long         'Sample cursor

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    Set mFailures = Failures
    mChecks = 0
    On Error GoTo Cleanup
    PriorCalc = Application.Calculation
    CalcChanged = True
    Application.Calculation = xlCalculationManual
    Set Scratch = Application.Workbooks.Add
    Scratch.Date1904 = False
    Set mSheet = Scratch.Worksheets(1)

'------------------------------------------------------------------------------
' BOUNDARY SAMPLES
'------------------------------------------------------------------------------
    'Window edges, month ends, leap days and century years, each with a
    'forward and a backward parameter set
        Boundaries = Array( _
            "1900-03-01", "1900-03-31", "1900-12-31", "1901-01-01", "1904-02-29", _
            "1999-12-31", "2000-02-28", "2000-02-29", "2000-03-01", "2023-02-28", _
            "2024-01-31", "2024-02-29", "2024-03-31", "2024-06-30", "2024-09-30", _
            "2024-12-31", "2100-02-28", "2100-03-01", "2400-02-29", "9999-11-30", _
            "9999-12-01", "9999-12-31")
        For I = LBound(Boundaries) To UBound(Boundaries)
            CheckSample SerialOf(CStr(Boundaries(I))), 1, 1, 1, 1, 5, True
            CheckSample SerialOf(CStr(Boundaries(I))), -1, -1, -1, 7, 1, False
        Next I

'------------------------------------------------------------------------------
' DETERMINISTIC SAMPLES
'------------------------------------------------------------------------------
    mSeed = ORACLE_SEED
    For I = 1 To ORACLE_SAMPLES
        CheckSample NextInt(MIN_SERIAL, MAX_SERIAL), NextInt(-1200, 1200), NextInt(-100000, 100000), _
                    NextInt(-100, 100), NextInt(1, 7), NextInt(1, 5), (NextInt(0, 1) = 1)
    Next I

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
Cleanup:
    If Err.Number <> 0 Then
        Fail "oracle/runner", "unexpected runtime error " & CStr(Err.Number) & ": " & Err.Description
        Err.Clear
    End If
    On Error Resume Next
    Set mSheet = Nothing
    If Not Scratch Is Nothing Then
        Scratch.Close SaveChanges:=False
        If Err.Number <> 0 Then
            Fail "oracle/cleanup", "scratch workbook did not close (error " & CStr(Err.Number) & ")"
            Err.Clear
        End If
    End If
    If CalcChanged Then
        Application.Calculation = PriorCalc
        If Err.Number <> 0 Then
            Fail "oracle/cleanup", "calculation mode was not restored (error " & CStr(Err.Number) & ")"
            Err.Clear
        ElseIf Application.Calculation <> PriorCalc Then
            Fail "oracle/cleanup", "calculation mode is " & CStr(Application.Calculation) & ", expected " & CStr(PriorCalc)
        End If
    End If
    Err.Clear
    On Error GoTo 0
    Checks = Checks + mChecks
    Set mFailures = Nothing

End Sub

'
'------------------------------------------------------------------------------
'
'                                  ORACLE CASES
'
'------------------------------------------------------------------------------
'

Private Sub CheckSample( _
    ByVal Serial As Long, _
    ByVal Months As Long, _
    ByVal Days As Long, _
    ByVal Years As Long, _
    ByVal WeekdayIndex As Long, _
    ByVal Occurrence As Long, _
    ByVal MondayBase As Boolean)
'
' Runs every overlap identity for one date serial and one parameter set.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim D               As Date         'The sample as a native VBA Date
    Dim Tag             As String       'Sample label
    Dim Yr              As Long         'YEAR(d)
    Dim Mo              As Long         'MONTH(d)
    Dim Dy              As Long         'DAY(d)
    Dim Q               As Long         'Months since the quarter began
    Dim Eom             As Double       'EOMONTH(d,0)
    Dim QuarterEnd      As Double       'EOMONTH(d,2-Q)
    Dim FebDays         As Long         'DAY(EOMONTH(d,2-MONTH(d)))
    Dim First           As Double       'd-DAY(d)+1
    Dim WeekType        As Long         'WEEKDAY return type: 2 Monday base, 1 Sunday base
    Dim FirstWeekday    As Long         'WEEKDAY(first,type)
    Dim LastWeekday     As Long         'WEEKDAY(EOMONTH(d,0),type)
    Dim Nth             As Double       'Oracle occurrence date

'------------------------------------------------------------------------------
' SAMPLE
'------------------------------------------------------------------------------
    D = CDate(Serial)
    Tag = IsoText(Serial)
    Yr = CLng(Xl("YEAR(" & CStr(Serial) & ")"))
    Mo = CLng(Xl("MONTH(" & CStr(Serial) & ")"))
    Dy = CLng(Xl("DAY(" & CStr(Serial) & ")"))
    Eom = CDbl(Xl("EOMONTH(" & CStr(Serial) & ",0)"))
    Q = (Mo - 1) Mod 3

'------------------------------------------------------------------------------
' MONTH, QUARTER AND YEAR BOUNDARIES
'------------------------------------------------------------------------------
    ExpectDate "EndOfMonth(d) = EOMONTH(d,0)", Tag, KPR_Dates_EndOfMonth(D), Eom
    If Yr = 1900 And Mo = 3 Then
        'Documented exclusion: Excel's EOMONTH(d,-1)+1 is not 1-Mar-1900 across its fictitious 29-Feb-1900
            ExpectDate "BeginOfMonth(d) = d-DAY(d)+1 (March 1900: EOMONTH(d,-1)+1 crosses Excel's fictitious 29-Feb-1900)", _
                       Tag, KPR_Dates_BeginOfMonth(D), Xl(CStr(Serial) & "-DAY(" & CStr(Serial) & ")+1")
    Else
        ExpectDate "BeginOfMonth(d) = EOMONTH(d,-1)+1", Tag, KPR_Dates_BeginOfMonth(D), Xl("EOMONTH(" & CStr(Serial) & ",-1)+1")
    End If
    ExpectNumber "DaysInMonth(d) = DAY(EOMONTH(d,0))", Tag, KPR_Dates_DaysInMonth(D), Xl("DAY(EOMONTH(" & CStr(Serial) & ",0))")
    ExpectBoolean "IsMonthEnd(d) = (d = EOMONTH(d,0))", Tag, KPR_Dates_IsMonthEnd(D), (CDbl(Serial) = Eom)
    ExpectDate "BeginOfQuarter(d) = EOMONTH(d,-MOD(MONTH(d)-1,3)-1)+1", Tag, KPR_Dates_BeginOfQuarter(D), _
               Xl("EOMONTH(" & CStr(Serial) & "," & CStr(-Q - 1) & ")+1")
    QuarterEnd = CDbl(Xl("EOMONTH(" & CStr(Serial) & "," & CStr(2 - Q) & ")"))
    ExpectDate "EndOfQuarter(d) = EOMONTH(d,2-MOD(MONTH(d)-1,3))", Tag, KPR_Dates_EndOfQuarter(D), QuarterEnd
    ExpectBoolean "IsQuarterEnd(d) = (d = EOMONTH(d,2-MOD(MONTH(d)-1,3)))", Tag, KPR_Dates_IsQuarterEnd(D), _
                  (CDbl(Serial) = QuarterEnd)
    ExpectDate "BeginOfYear(d) = EOMONTH(d,-MONTH(d))+1", Tag, KPR_Dates_BeginOfYear(D), _
               Xl("EOMONTH(" & CStr(Serial) & "," & CStr(-Mo) & ")+1")
    ExpectDate "EndOfYear(d) = EOMONTH(d,12-MONTH(d))", Tag, KPR_Dates_EndOfYear(D), _
               Xl("EOMONTH(" & CStr(Serial) & "," & CStr(12 - Mo) & ")")
    ExpectBoolean "IsYearEnd(d) = (MONTH(d)=12 AND DAY(d)=31)", Tag, KPR_Dates_IsYearEnd(D), (Mo = 12 And Dy = 31)

'------------------------------------------------------------------------------
' LEAP YEARS
'------------------------------------------------------------------------------
    If Yr = 1900 Then
        'Documented exclusion: Excel's February 1900 has a fictitious 29th day
            ExpectBoolean "IsLeapYear(1900) = FALSE (excluded from the oracle: Excel's fictitious 29-Feb-1900)", _
                          Tag, KPR_Dates_IsLeapYear(Yr), False
            ExpectNumber "DaysInYear(1900) = 365 (excluded from the oracle: Excel's fictitious 29-Feb-1900)", _
                         Tag, KPR_Dates_DaysInYear(Yr), 365
    Else
        FebDays = CLng(Xl("DAY(EOMONTH(" & CStr(Serial) & "," & CStr(2 - Mo) & "))"))
        ExpectBoolean "IsLeapYear(YEAR(d)) = (DAY(EOMONTH(d,2-MONTH(d)))=29)", Tag, KPR_Dates_IsLeapYear(Yr), _
                      (FebDays = 29)
        ExpectNumber "DaysInYear(YEAR(d)) = 337+DAY(EOMONTH(d,2-MONTH(d)))", Tag, KPR_Dates_DaysInYear(Yr), _
                     337 + FebDays
    End If

'------------------------------------------------------------------------------
' WEEKDAYS
'------------------------------------------------------------------------------
    ExpectNumber "DayOfWeek(d) = WEEKDAY(d,2)", Tag, KPR_Dates_DayOfWeek(D), Xl("WEEKDAY(" & CStr(Serial) & ",2)")
    ExpectNumber "DayOfWeek(d,FALSE) = WEEKDAY(d,1)", Tag, KPR_Dates_DayOfWeek(D, False), Xl("WEEKDAY(" & CStr(Serial) & ",1)")

'------------------------------------------------------------------------------
' ARITHMETIC
'------------------------------------------------------------------------------
    ExpectDate "AddMonths(d,k) = EDATE(d,k)", Tag & " k=" & CStr(Months), KPR_Dates_AddMonths(D, Months), _
               Xl("EDATE(" & CStr(Serial) & "," & CStr(Months) & ")")
    ExpectDate "AddYears(d,j) = EDATE(d,12*j)", Tag & " j=" & CStr(Years), KPR_Dates_AddYears(D, Years), _
               Xl("EDATE(" & CStr(Serial) & "," & CStr(12 * Years) & ")")
    ExpectDate "EndOfMonth(AddMonths(d,k)) = EOMONTH(d,k)", Tag & " k=" & CStr(Months), _
               KPR_Dates_EndOfMonth(KPR_Dates_AddMonths(D, Months)), Xl("EOMONTH(" & CStr(Serial) & "," & CStr(Months) & ")")
    ExpectDate "AddDays(d,n) = d+n", Tag & " n=" & CStr(Days), KPR_Dates_AddDays(D, Days), _
               Xl(CStr(Serial) & "+(" & CStr(Days) & ")")
    ExpectDate "AddWeeks(d,w) = d+7*w", Tag & " w=" & CStr(Months), KPR_Dates_AddWeeks(D, Months), _
               Xl(CStr(Serial) & "+7*(" & CStr(Months) & ")")

'------------------------------------------------------------------------------
' WEEKDAY LOCATORS
'------------------------------------------------------------------------------
    If MondayBase Then WeekType = 2 Else WeekType = 1
    'd-DAY(d)+1, not EOMONTH(d,-1)+1, which crosses Excel's fictitious 29-Feb-1900 in March 1900
    First = CDbl(Xl(CStr(Serial) & "-DAY(" & CStr(Serial) & ")+1"))
    FirstWeekday = CLng(Xl("WEEKDAY(" & CStr(First) & "," & CStr(WeekType) & ")"))
    Nth = First + ((WeekdayIndex - FirstWeekday + 7) Mod 7) + 7 * (Occurrence - 1)
    If Nth > Eom Then
        'The occurrence is absent from the month: the contract requires #NUM!
            Nth = MAX_SERIAL + 1
    End If
    ExpectDate "NthWeekdayOfMonth(y,m,i,n) = first+MOD(i-WEEKDAY(first,t),7)+7*(n-1), first = d-DAY(d)+1, absent after EOMONTH(d,0)", _
               Tag & " i=" & CStr(WeekdayIndex) & " n=" & CStr(Occurrence) & " t=" & CStr(WeekType), _
               KPR_Dates_NthWeekdayOfMonth(Yr, Mo, WeekdayIndex, Occurrence, MondayBase), Nth
    LastWeekday = CLng(Xl("WEEKDAY(" & CStr(Eom) & "," & CStr(WeekType) & ")"))
    ExpectDate "LastWeekdayOfMonth(y,m,i) = EOMONTH(d,0)-MOD(WEEKDAY(EOMONTH(d,0),t)-i,7)", _
               Tag & " i=" & CStr(WeekdayIndex) & " t=" & CStr(WeekType), _
               KPR_Dates_LastWeekdayOfMonth(Yr, Mo, WeekdayIndex, MondayBase), _
               Eom - ((LastWeekday - WeekdayIndex + 7) Mod 7)

End Sub

'
'------------------------------------------------------------------------------
'
'                                   ASSERTIONS
'
'------------------------------------------------------------------------------
'

Private Sub ExpectDate( _
    ByVal Assumption As String, _
    ByVal Tag As String, _
    ByVal Actual As Variant, _
    ByVal Oracle As Variant)
'
' A date identity. When Excel's answer lies outside the supported window, or
' Excel refuses, the contract requires #NUM! and that is what is asserted.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim InWindow        As Boolean      'Excel's answer is a supported date

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    mChecks = mChecks + 1
    If VarType(Oracle) <> vbError Then
        InWindow = (CDbl(Oracle) >= MIN_SERIAL And CDbl(Oracle) <= MAX_SERIAL)
    End If
    If Not InWindow Then
        If VarType(Actual) <> vbError Then
            Fail Label(Tag, Assumption), "outside the supported window: expected #NUM!, got " & Describe(Actual)
        ElseIf CLng(Actual) <> ERR_NUM Then
            Fail Label(Tag, Assumption), "outside the supported window: expected #NUM!, got " & Describe(Actual)
        End If
    ElseIf VarType(Actual) <> vbDate Then
        Fail Label(Tag, Assumption), "expected Date serial " & CStr(CDbl(Oracle)) & ", got " & Describe(Actual)
    ElseIf CDbl(Actual) <> CDbl(Oracle) Then
        Fail Label(Tag, Assumption), "expected Date serial " & CStr(CDbl(Oracle)) & ", got " & Describe(Actual)
    End If

End Sub

Private Sub ExpectNumber( _
    ByVal Assumption As String, _
    ByVal Tag As String, _
    ByVal Actual As Variant, _
    ByVal Oracle As Variant)
'
' A Long identity.
'

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    mChecks = mChecks + 1
    If VarType(Oracle) = vbError Then
        Fail Label(Tag, Assumption), "Excel returned " & Describe(Oracle)
    ElseIf VarType(Actual) <> vbLong Then
        Fail Label(Tag, Assumption), "expected Long " & CStr(CDbl(Oracle)) & ", got " & Describe(Actual)
    ElseIf CDbl(Actual) <> CDbl(Oracle) Then
        Fail Label(Tag, Assumption), "expected Long " & CStr(CDbl(Oracle)) & ", got " & Describe(Actual)
    End If

End Sub

Private Sub ExpectBoolean( _
    ByVal Assumption As String, _
    ByVal Tag As String, _
    ByVal Actual As Variant, _
    ByVal Expected As Boolean)
'
' A Boolean identity.
'

'------------------------------------------------------------------------------
' EVALUATE
'------------------------------------------------------------------------------
    mChecks = mChecks + 1
    If VarType(Actual) <> vbBoolean Then
        Fail Label(Tag, Assumption), "expected Boolean " & CStr(Expected) & ", got " & Describe(Actual)
    ElseIf CBool(Actual) <> Expected Then
        Fail Label(Tag, Assumption), "expected Boolean " & CStr(Expected) & ", got " & Describe(Actual)
    End If

End Sub

Private Sub Fail( _
    ByVal CaseLabel As String, _
    ByVal Detail As String)
'
' Records one failing case in the caller's collection.
'
    mFailures.Add Array(CaseLabel, Detail)

End Sub

'
'------------------------------------------------------------------------------
'
'                                    HELPERS
'
'------------------------------------------------------------------------------
'

Private Function Xl( _
    ByVal Formula As String) _
    As Variant
'
' Evaluates one oracle expression on the scratch 1900 sheet.
'
    Xl = mSheet.Evaluate(Formula)

End Function

Private Function NextInt( _
    ByVal Low As Long, _
    ByVal High As Long) _
    As Long
'
' Next Park-Miller value mapped to Low .. High. Every product stays below
' 2^53, so the sequence is exact in Double arithmetic on every host.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Span            As Double       'Number of values in the range

'------------------------------------------------------------------------------
' STEP
'------------------------------------------------------------------------------
    mSeed = mSeed * PM_MULTIPLIER
    mSeed = mSeed - Int(mSeed / PM_MODULUS) * PM_MODULUS
    Span = CDbl(High) - CDbl(Low) + 1#
    NextInt = Low + CLng(mSeed - Int(mSeed / Span) * Span)

End Function

Private Function SerialOf( _
    ByVal Iso As String) _
    As Long
'
' 1900 serial of an ISO date inside the supported window, where VBA and
' Excel serials coincide.
'
    SerialOf = CLng(DateSerial(CInt(Left$(Iso, 4)), CInt(Mid$(Iso, 6, 2)), CInt(Right$(Iso, 2))))

End Function

Private Function IsoText( _
    ByVal Serial As Long) _
    As String
'
' yyyy-mm-dd for a label, independent of the host's date separator.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim D               As Date         'The serial as a VBA Date

'------------------------------------------------------------------------------
' FORMAT
'------------------------------------------------------------------------------
    D = CDate(Serial)
    IsoText = Right$("000" & CStr(Year(D)), 4) & "-" & Right$("0" & CStr(Month(D)), 2) & "-" & _
              Right$("0" & CStr(Day(D)), 2)

End Function

Private Function Label( _
    ByVal Tag As String, _
    ByVal Assumption As String) _
    As String
'
' The case label: the sample, then the overlap identity it tests.
'
    Label = "oracle " & Tag & ": " & Assumption

End Function

Private Function Describe( _
    ByVal V As Variant) _
    As String
'
' A value with its subtype for a failure message.
'
    Select Case VarType(V)
        Case vbError:       Describe = "error " & CStr(CLng(V))
        Case vbDate:        Describe = "Date serial " & CStr(CDbl(V))
        Case vbBoolean:     Describe = "Boolean " & CStr(V)
        Case vbString:      Describe = "String """ & CStr(V) & """"
        Case vbEmpty:       Describe = "Empty"
        Case Else
            If IsArray(V) Then
                Describe = "an array"
            Else
                Describe = TypeName(V) & " " & Trim$(Str$(V))
            End If
    End Select

End Function
