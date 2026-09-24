Attribute VB_Name = "KPR_Core_Dates"
'==============================================================================
' MODULE: KPR_Core_Dates
'------------------------------------------------------------------------------
' PURPOSE
'   Internal boundary for calendar computation: Gregorian arithmetic, month
'   boundary logic, EOM-aware month shifting, and the pillar parse/format core.
'
'   This module owns the supported date window, because the window is a
'   calendar fact rather than a parsing rule.
'
' SCOPE
'   - Supported window constants and the window predicate
'   - IsLeapYear
'   - DaysInMonth
'   - EndOfMonth, BeginOfQuarter, EndOfQuarter, BeginOfYear, EndOfYear
'   - TryAddMonths
'   - TryPillar_Parse
'   - TryPillar_Format
'
' VISIBILITY
'   Option Private Module. Members are declared Public so other modules in the
'   project can call them; the module-level Private setting is what keeps them
'   out of the Excel function list and off the project's external surface.
'   Nothing here is supported API.
'
' ALLOWED DEPENDENCIES
'   KPR_Core_Err, for the condition vocabulary only. This module never
'   constructs a worksheet error value: it reports a classified condition and
'   the facade maps it. It never parses text other than pillar tokens and
'   never inspects shape. VBA date intrinsics are used directly.
'
' SUPPORTED WINDOW
'       KPR_MIN_DATE = 1900-03-01
'       KPR_MAX_DATE = 9999-12-31
'
'   The lower bound is serial 61 by design: it excludes the Excel 1900-system
'   region where worksheet serials and VBA dates disagree because of the
'   fictitious 29-Feb-1900. At and above serial 61 the two systems agree.
'
'   The baseline applied this gate inside its combined parser. The gate is now
'   applied by the facade immediately after parsing, using the predicate below.
'   It remains a single gate on every accepted path; only its location moved,
'   so that KPR_Core_Parse need not depend on this module.
'
'   The derived index constants below restate the window in absolute months.
'   They exist to keep per-call arithmetic off the hot paths and MUST be
'   updated together with the two date constants.
'
' MONTH LENGTH
'   DaysInMonth is the single source of month length for this module, and it
'   is the only member that consults the leap rule. EndOfMonth is built from
'   it, and TryAddMonths reads it directly.
'
'   The day-0 idiom DateSerial(y, m + 1, 0) is deliberately NOT used anywhere
'   in this module. It raises error 5 at both ends of the VBA Date range:
'   DateSerial(9999, 13, 0) and DateSerial(100, 1, 0) both fail, verified in
'   isolation on the host. Deriving month length arithmetically removes the
'   boundary crossing rather than guarding it.
'
' PILLAR POLICY
'   TryPillar_Format implements the three rounding modes over one uniform
'   candidate set, and TryPillar_Parse reports every grammar rejection under
'   its own condition identifier. Both are specified in contract section 8.4
'   and 3.4; the module header does not restate the rules.
'
' NOTES
'   - Date math follows Excel / VBA DateSerial, DateAdd, DateDiff and Weekday
'     semantics. Gregorian leap-year logic is used throughout.
'   - Day distances between two known date-only values are taken by serial
'     subtraction rather than DateDiff. The results are identical because both
'     operands are integral, and the subtraction is materially cheaper on the
'     per-cell paths.
'   - The facade AddYears delegates to TryAddMonths so year-roll edge cases
'     (29-Feb, short months, EOM) cannot drift between the two surfaces.
'
' UPDATED
'   2026-09-02
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit         'Force explicit variable declarations
    Option Private Module   'Internal module: invisible outside this VBA project

'------------------------------------------------------------------------------
' ROUNDING VOCABULARY
'------------------------------------------------------------------------------
    'Pillar rounding modes. A calendar concept, so it lives with the calendar
    'core. The facade maps the normalized Opt_Rounding token onto these; the
    'parser cannot, because KPR_Core_Parse may not depend on this module.
        Public Enum KPR_PillarRounding
            KPR_ROUND_NEAREST = 0
            KPR_ROUND_FLOOR = 1
            KPR_ROUND_CEILING = 2
        End Enum

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Supported date window (single source of truth for every parse / result gate)
        Public Const KPR_MIN_DATE   As Date = #3/1/1900#     'Serial 61; excludes the Excel 1900 leap-bug region
        Public Const KPR_MAX_DATE   As Date = #12/31/9999#   'Serial 2958465; upper bound of the Excel date system

    'The same window as absolute 0-based month indices, so the hot paths do not
    'recompute Year / Month of a constant on every call. Keep in step with the
    'two date constants above.
        Private Const KPR_MIN_YEAR      As Long = 1900       'Year(KPR_MIN_DATE)
        Private Const KPR_MAX_YEAR      As Long = 9999       'Year(KPR_MAX_DATE)
        Private Const KPR_MAX_MONTHIDX  As Long = 119999     '9999 * 12 + 12 - 1

'
'------------------------------------------------------------------------------
'
'                              SUPPORTED WINDOW
'
'------------------------------------------------------------------------------
'

Public Function IsDateInWindow( _
    ByVal DateIn As Date) _
    As Boolean
'
'==============================================================================
'                                IsDateInWindow
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a parsed date lies inside the supported window.
'
' SIGNATURE
'   IsDateInWindow(DateIn) -> Boolean
'
' INPUTS
'   DateIn
'     Any VBA Date value.
'
' RETURNS
'   Boolean
'     TRUE  => KPR_MIN_DATE <= DateIn <= KPR_MAX_DATE
'     FALSE => outside the supported window
'
' ERROR POLICY
'   None. The caller decides which worksheet error an out-of-window value maps
'   to.
'
' NOTES
'   - Input parsing reports DATE_WINDOW and result construction reports
'     RESULT_WINDOW; both map to #NUM! at the facade. This predicate remains
'     policy-free so callers can select the correct condition for their path.
'   - A value carrying a time component is compared as a serial, so 9999-12-31
'     with a non-zero time reports FALSE. Callers holding a raw parse result
'     should floor it before gating if that is not intended.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'One window test for every accepted path
        IsDateInWindow = ((DateIn >= KPR_MIN_DATE) And (DateIn <= KPR_MAX_DATE))

End Function

'
'------------------------------------------------------------------------------
'
'                            CALENDAR COMPUTATION CORE
'
'------------------------------------------------------------------------------
'

Public Function IsLeapYear( _
    ByVal YearIn As Long) _
    As Boolean
'
'==============================================================================
'                                  IsLeapYear
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the Gregorian leap-year rule to a calendar year.
'
' WHY THIS EXISTS
'   The rule is short enough to inline and subtle enough to get wrong. Stating
'   the century exceptions once means DaysInMonth is the only place February
'   is resolved, and every other month length derives from that.
'
' SIGNATURE
'   IsLeapYear(YearIn) -> Boolean
'
' INPUTS
'   YearIn
'     Calendar year, not a date serial. Callers holding a Date must pass
'     Year(DateIn).
'
' RETURNS
'   Boolean
'     TRUE => YearIn is a leap year under the proleptic Gregorian rule.
'
' ERROR POLICY
'   None. Total for every Long; the caller owns range validation.
'
' NOTES
'   - Rule: divisible by 4 => leap, unless divisible by 100, unless also
'     divisible by 400. So 1900 is not a leap year and 2000 is.
'   - The rule is applied proleptically. VBA Date arithmetic is Gregorian
'     over its whole range, so no Julian calendar or 1582 changeover is
'     modelled here and 1500 reports TRUE. That is consistent with the rest
'     of the layer, not a historical claim.
'   - The supported window is not enforced. Any Long is evaluated, including
'     zero and negatives; VBA Mod takes the sign of the dividend, so -4
'     reports TRUE. Values outside the window are the caller's to reject.
'
' UPDATED
'   2026-08-31
'==============================================================================
'
'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Gregorian rule
        IsLeapYear = ((YearIn Mod 4 = 0) And ((YearIn Mod 100 <> 0) Or (YearIn Mod 400 = 0)))
End Function


Public Function DaysInMonth( _
    ByVal YearIn As Long, _
    ByVal MonthIn As Long) _
    As Long
'
'==============================================================================
'                                  DaysInMonth
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the length in days of a given calendar month.
'
' WHY THIS EXISTS
'   Month length is needed by EndOfMonth, by the EOM anchor test, and by the
'   day clip in TryAddMonths. Deriving it arithmetically rather than through
'   DateSerial avoids its upper-bound normalization trap:
'   DateSerial(9999, 13, 0) raises error 5. The lower supported boundary is
'   not symmetric; DateSerial(1900, 1, 0) returns 1899-12-31 successfully.
'
' SIGNATURE
'   DaysInMonth(YearIn, MonthIn) -> Long
'
' INPUTS
'   YearIn
'     Calendar year, not a date serial. Read only for February.
'
'   MonthIn
'     Calendar month, 1 to 12.
'
' RETURNS
'   Long
'     28, 29, 30 or 31 for a valid month; 0 for any MonthIn outside 1 to 12.
'
' ERROR POLICY
'   None. Total for every pair of Longs, including an invalid month.
'
' DEPENDENCIES
'   - IsLeapYear
'
' NOTES
'   - A zero return signals an invalid month rather than raising, because this
'     is an internal member reached only with an already-resolved month. A
'     caller that can produce an out-of-range month must test the result; one
'     that cannot may use it directly.
'   - No date is constructed, so the result is valid for any year the caller
'     can represent, including years outside the supported window.
'
' UPDATED
'   2026-08-31
'==============================================================================
'
'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Fixed lengths first; February is the only month needing the leap rule
        Select Case MonthIn
            Case 1, 3, 5, 7, 8, 10, 12
                DaysInMonth = 31
            Case 4, 6, 9, 11
                DaysInMonth = 30
            Case 2
                If IsLeapYear(YearIn) Then
                    DaysInMonth = 29
                Else
                    DaysInMonth = 28
                End If
            Case Else
                'Invalid month; reported rather than raised
                    DaysInMonth = 0
        End Select
End Function


Public Function EndOfMonth( _
    ByVal DateIn As Date) _
    As Date
'
'==============================================================================
'                                  EndOfMonth
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the last calendar day of the month containing DateIn.
'
' WHY THIS EXISTS
'   Month-end is the anchor for the whole month-arithmetic family. Isolating
'   it here means no caller re-derives it.
'
' SIGNATURE
'   EndOfMonth(DateIn) -> Date
'
' INPUTS
'   DateIn
'     Already-parsed VBA Date. A time component is tolerated and discarded.
'
' RETURNS
'   Date
'     Midnight on the last calendar day of the containing month.
'
' ERROR POLICY
'   None. Total over the whole VBA Date range, both boundary months included.
'
' DEPENDENCIES
'   - DaysInMonth
'
' NOTES
'   - The month length comes from DaysInMonth, so no date outside the month
'     is ever constructed. The day-0 idiom would be shorter but raises error 5
'     for DateSerial(9999, 13, 0). DateSerial(1900, 1, 0) succeeds; see MONTH
'     LENGTH in the module header for the distinct VBA-range lower boundary.
'   - Year and Month ignore any time fraction, and DateSerial always returns
'     midnight, so the result is date-only regardless of what arrives. No
'     caller needs to strip the time first.
'   - Totality is asserted over the whole VBA Date range, which is wider than
'     the supported window. Pre-1900 dates are negative serials and behave
'     identically.
'
' UPDATED
'   2026-08-31
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim YearPart    As Long     'Calendar year of DateIn
    Dim MonthPart   As Long     'Calendar month of DateIn, 1 to 12
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Decompose once; both parts feed the month-length lookup
        YearPart = Year(DateIn)
        MonthPart = Month(DateIn)
'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Last day of the containing month, built without leaving it
        EndOfMonth = DateSerial(YearPart, MonthPart, DaysInMonth(YearPart, MonthPart))
End Function


Public Function BeginOfQuarter( _
    ByVal DateIn As Date) _
    As Date
'
'==============================================================================
'                                BeginOfQuarter
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the first day of the calendar quarter containing DateIn, using
'   quarters Jan-Mar, Apr-Jun, Jul-Sep and Oct-Dec.
'
' NOTES
'   - Pure calendar function: no window gate. A Q1-1900 input yields
'     1900-01-01, which is a valid VBA Date but outside the supported window;
'     the facade element gates the result.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Quarter start month is 1, 4, 7 or 10
        BeginOfQuarter = DateSerial(Year(DateIn), ((Month(DateIn) - 1) \ 3) * 3 + 1, 1)

End Function

Public Function EndOfQuarter( _
    ByVal DateIn As Date) _
    As Date
'
'==============================================================================
'                                 EndOfQuarter
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the last day of the calendar quarter containing DateIn.
'
' NOTES
'   - Month length comes from DaysInMonth, never from the day-zero idiom.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LastMonth       As Long     'Last month of the quarter: 3, 6, 9 or 12

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Quarter end month, then its length
        LastMonth = ((Month(DateIn) - 1) \ 3) * 3 + 3
        EndOfQuarter = DateSerial(Year(DateIn), LastMonth, DaysInMonth(Year(DateIn), LastMonth))

End Function

Public Function BeginOfYear( _
    ByVal DateIn As Date) _
    As Date
'
'==============================================================================
'                                 BeginOfYear
'------------------------------------------------------------------------------
' PURPOSE
'   Returns 1 January of the year containing DateIn.
'
' NOTES
'   - Pure calendar function: no window gate. Any 1900 input yields
'     1900-01-01, which the facade element gates as RESULT_WINDOW.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    BeginOfYear = DateSerial(Year(DateIn), 1, 1)

End Function

Public Function EndOfYear( _
    ByVal DateIn As Date) _
    As Date
'
'==============================================================================
'                                  EndOfYear
'------------------------------------------------------------------------------
' PURPOSE
'   Returns 31 December of the year containing DateIn.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    EndOfYear = DateSerial(Year(DateIn), 12, 31)

End Function

Public Function TryAddMonths( _
    ByVal DateIn As Date, _
    ByVal nMonths As Long, _
    ByVal KeepEOM As Boolean, _
    ByRef DateOut As Date) _
    As Boolean
'
'==============================================================================
'                                 TryAddMonths
'------------------------------------------------------------------------------
' PURPOSE
'   Single month shifter for the module. Adds nMonths to an already-parsed
'   date with explicit end-of-month semantics and full range safety.
'
' WHY THIS EXISTS
'   AddMonths, AddYears and the pillar resolver all need the same shift with
'   the same edge-case behavior. Duplicating it three times is how the three
'   drift apart.
'
' SIGNATURE
'   TryAddMonths(DateIn, nMonths, KeepEOM, DateOut) -> Boolean
'
' INPUTS
'   DateIn
'     Already-parsed VBA Date. A time component is tolerated and discarded.
'
'   nMonths
'     Signed month delta.
'
'   KeepEOM
'     TRUE  => if DateIn is month-end, return the target month-end
'     FALSE => clip day-of-month to the target month length
'
' OUTPUTS
'   DateOut (ByRef)
'     Assigned ONLY on success.
'
' RETURNS
'   Boolean
'     TRUE  => shift succeeded and the result is inside the supported window
'     FALSE => the target month or the result falls outside it
'
' BEHAVIOR
'   - Converts year / month into an absolute month index, shifts it, and gates
'     the index before it is ever handed to DateSerial. This is why an absurd
'     nMonths returns FALSE instead of raising an overflow.
'   - Resolves the day-of-month by the KeepEOM rule, then re-gates the result.
'
' ERROR POLICY
'   - Does not raise. All failure paths return FALSE.
'   - Clears Err before returning FALSE, so no stale error state survives.
'
' DEPENDENCIES
'   - DaysInMonth
'
' NOTES
'   - DateOut is safe to pass as the same variable as an input date at the call
'     site, because it is written only once, at the end.
'   - Both month lengths come from DaysInMonth, so neither the EOM anchor test
'     nor the day clip constructs a date outside the month it is measuring.
'   - The index gate is deliberately coarse, admitting whole years at both
'     ends. The terminal gate is what enforces the exact window, so if either
'     bound ever moves off a year boundary the pair still holds.
'   - The index is computed in Double so that an extreme nMonths cannot
'     overflow before the gate can reject it.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Y               As Long      'Year component of DateIn
    Dim M               As Long      'Month component of DateIn
    Dim DayIn           As Long      'Day-of-month of DateIn
    Dim IsEOM           As Boolean   'TRUE when DateIn is its own month-end

    Dim MonthIndex      As Double    'Absolute 0-based month index after the shift

    Dim TargetY         As Long      'Resolved target year
    Dim TargetM         As Long      'Resolved target month
    Dim TargetDim       As Long      'Days in the target month
    Dim TargetD         As Long      'Resolved day-of-month in the target month

    Dim Candidate       As Date      'Shifted date before the terminal range gate

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Trap runtime errors and convert them to a FALSE return
        On Error GoTo Fail
    'Default return is failure unless we explicitly succeed
        TryAddMonths = False

'------------------------------------------------------------------------------
' CACHE INPUT COMPONENTS
'------------------------------------------------------------------------------
    'Split the input date once
        Y = Year(DateIn)
        M = Month(DateIn)
        DayIn = Day(DateIn)
    'Detect the EOM anchor before shifting
        IsEOM = (DayIn = DaysInMonth(Y, M))

'------------------------------------------------------------------------------
' RESOLVE TARGET MONTH
'------------------------------------------------------------------------------
    'Absolute 0-based month index, shifted, computed in Double
        MonthIndex = (CDbl(Y) * 12# + CDbl(M) - 1#) + CDbl(nMonths)
    'Gate the index before any coercion so DateSerial never sees a wild value
        If (MonthIndex < (CDbl(KPR_MIN_YEAR) * 12#)) Or _
           (MonthIndex > (CDbl(KPR_MAX_YEAR) * 12# + 11#)) Then GoTo Fail
    'Split the gated index back into year and month
        TargetY = CLng(Int(MonthIndex / 12#))
        TargetM = CLng(MonthIndex - (CDbl(TargetY) * 12#)) + 1

'------------------------------------------------------------------------------
' RESOLVE DAY OF MONTH
'------------------------------------------------------------------------------
    'Target month length, read without constructing a date
        TargetDim = DaysInMonth(TargetY, TargetM)

    'Preserve EOM when asked and the input was EOM; otherwise clip
        If KeepEOM And IsEOM Then
            TargetD = TargetDim
        Else
            TargetD = DayIn
            If TargetD > TargetDim Then TargetD = TargetDim
        End If

'------------------------------------------------------------------------------
' TERMINAL RANGE GATE
'------------------------------------------------------------------------------
    'Build the candidate and gate it against the supported window
        Candidate = DateSerial(TargetY, TargetM, TargetD)
        If (Candidate < KPR_MIN_DATE) Or (Candidate > KPR_MAX_DATE) Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once the value is fully validated
        DateOut = Candidate
    'Contract: TRUE only when DateOut was assigned
        TryAddMonths = True
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return FALSE per contract, leaving DateOut untouched
        TryAddMonths = False
    'Do not leave stale error state for the caller to observe
        Err.Clear

End Function

'
'------------------------------------------------------------------------------
'
'                                 PILLAR CORE
'
'------------------------------------------------------------------------------
'

Public Function TryPillar_Parse( _
    ByVal PillarIn As Variant, _
    ByRef TotalMonths As Double, _
    ByRef TotalDays As Double, _
    ByRef Condition As KPR_Condition) _
    As Boolean
'
'==============================================================================
'                               TryPillar_Parse
'------------------------------------------------------------------------------
' PURPOSE
'   Parses a pillar token into a signed month delta and a signed day delta.
'
' SIGNATURE
'   TryPillar_Parse(PillarIn, TotalMonths, TotalDays, Condition) -> Boolean
'
' INPUTS
'   PillarIn
'     Scalar Variant expected to hold pillar text.
'
' OUTPUTS
'   TotalMonths (ByRef)
'     Signed aggregate of the Y and M components: sign * (12Y + M).
'
'   Condition (ByRef)
'     Always assigned. KPR_COND_NONE on success; otherwise one of
'       PILLAR_TYPE_REJECTED     non-text payload
'       PILLAR_ALIAS_SIGNED      whole-token alias carrying a leading sign
'       PILLAR_DUPLICATE_UNIT    a unit appearing more than once
'       PILLAR_AGGREGATE_RANGE   valid numeric component/aggregate out of range
'       PILLAR_TOKEN_MALFORMED   every other grammar violation
'
'   TotalDays (ByRef)
'     Signed aggregate of the W and D components: sign * (7W + D).
'
'   Both are assigned ONLY on success.
'
' RETURNS
'   Boolean
'     TRUE  => the token matched the accepted grammar
'     FALSE => blank, malformed, unknown unit, repeated unit, signed alias,
'              non-text payload, or valid numeric component/aggregate out of range
'
' BEHAVIOR
'   - Trims and upper-cases, consumes an optional leading sign, then checks the
'     whole-token aliases ON / O/N / TN / T/N.
'   - Otherwise scans repeated [digits][unit] pairs to the end of the string.
'   - Any character that is not a digit in a numeric position or a known unit
'     in a unit position fails the parse.
'   - Each unit may appear at most once, so "1M1M" fails rather than summing.
'
' ERROR POLICY
'   - Does not raise. All failure paths return FALSE.
'   - Clears Err before returning FALSE, so no stale error state survives.
'
' DEPENDENCIES
'   - KPR_Core_Err condition vocabulary
'
' NOTES
'   - No characters are stripped from the body, so "1 M" and "1/M" fail rather
'     than being silently reinterpreted as "1M". The aliases are the only
'     tokens containing "/", and they are matched whole.
'   - Fractional tenors are not part of the grammar: "1.5M" fails. Half months
'     are not a market convention, and admitting them would force a rounding
'     policy into a parser.
'   - A leading sign on an alias is rejected. The aliases name fixed points at
'     the short end, not quantities, so "-ON" has no meaning to negate.
'   - A repeated unit is a typo, not a sum. Accepting "1M1M" as two months
'     would return a plausible number for input the caller did not intend;
'     contract section 3.4 therefore rejects duplicate units.
'   - Finite quantities are not bounded here. A token such as "999999999999M"
'     parses and the magnitude is left for the shift layer to reject at its
'     month-index gate. A digits-only component that cannot be represented as
'     Double, or an aggregate that overflows Double, is PILLAR_AGGREGATE_RANGE.
'   - CDbl is applied to a digits-only substring, so no locale decimal
'     separator is involved. Replacing it with Val or CLng would change the
'     overflow behaviour.
'   - The digit scan compares character codes rather than one-character
'     strings. The unit test stays on strings because it runs once per
'     component rather than once per character.
'   - A signed result can be negative zero, because the sign multiplier is
'     applied to a zero aggregate. It compares equal to zero and behaves
'     identically in arithmetic; only a formatter would notice.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim S               As String    'Upper-cased, trimmed pillar text
    Dim SBody           As String    'Pillar body after the optional leading sign
    Dim ChCode          As Long      'Character code at the current scan position
    Dim HasSign         As Boolean   'TRUE when a leading sign was consumed

    Dim ScanPos         As Long      'Current scan position in SBody
    Dim TokenStart      As Long      'Start position of the current numeric token
    Dim TokenCount      As Long      'Number of components parsed so far
    Dim BodyLen         As Long      'Length of SBody, read once

    Dim SignMul         As Double    'Global sign multiplier (+1 or -1)
    Dim QtyD            As Double    'Current component quantity

    Dim SeenY           As Boolean   'TRUE once a Y component has been consumed
    Dim SeenM           As Boolean   'TRUE once an M component has been consumed
    Dim SeenW           As Boolean   'TRUE once a W component has been consumed
    Dim SeenD           As Boolean   'TRUE once a D component has been consumed

    Dim YearsD          As Double    'Y component
    Dim MonthsD         As Double    'M component
    Dim WeeksD          As Double    'W component
    Dim DaysD           As Double    'D component, or the resolved alias
    Dim ParsedMonths    As Double    'Local month aggregate; assigned out only on success
    Dim ParsedDays      As Double    'Local day aggregate; assigned out only on success

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Trap runtime errors and convert them to a FALSE return
        On Error GoTo Fail

    'Default classification for any structural rejection below
        Condition = KPR_COND_PILLAR_TOKEN_MALFORMED
    'Default return is failure unless we explicitly succeed
        TryPillar_Parse = False
    'Default sign is positive
        SignMul = 1#

'------------------------------------------------------------------------------
' NORMALIZE PILLAR TEXT
'------------------------------------------------------------------------------
    'Require a text payload
        If VarType(PillarIn) <> vbString Then
            Condition = KPR_COND_PILLAR_TYPE_REJECTED
            GoTo Fail
        End If
    'Trim and upper-case once
        S = Pillar_AsciiUpper(Pillar_AsciiTrim(CStr(PillarIn)))
    'Reject a blank token
        If Len(S) = 0 Then GoTo Fail

'------------------------------------------------------------------------------
' CONSUME OPTIONAL LEADING SIGN
'------------------------------------------------------------------------------
    'A sign applies to the whole token
        Select Case Left$(S, 1)
            Case "+"
                HasSign = True
                SBody = Mid$(S, 2)
            Case "-"
                HasSign = True
                SignMul = -1#
                SBody = Mid$(S, 2)
            Case Else
                SBody = S
        End Select

    'Reject a sign with no body
        BodyLen = Len(SBody)
        If BodyLen = 0 Then GoTo Fail

'------------------------------------------------------------------------------
' WHOLE-TOKEN ALIASES
'------------------------------------------------------------------------------
    'Overnight and tom-next are matched whole, never as generic tokens
        Select Case SBody
            Case "ON", "O/N", "TN", "T/N"
                'An alias names a fixed point and cannot carry a sign
                    If HasSign Then
                        Condition = KPR_COND_PILLAR_ALIAS_SIGNED
                        GoTo Fail
                    End If
                'Resolve the alias to its day count
                    If (SBody = "ON") Or (SBody = "O/N") Then
                        DaysD = 1#
                    Else
                        DaysD = 2#
                    End If
                    TokenCount = 1
        End Select

'------------------------------------------------------------------------------
' GENERIC INTEGER + UNIT SCANNER
'------------------------------------------------------------------------------
    'Only run the scanner when no alias matched
        If TokenCount = 0 Then

            'Start at the first character of the body
                ScanPos = 1
            'Consume [digits][unit] pairs until the body is exhausted
                Do While ScanPos <= BodyLen
                    'Mark the start of the numeric token
                        TokenStart = ScanPos
                    'Advance through contiguous digits, comparing codes
                        Do While ScanPos <= BodyLen
                            ChCode = AscW(Mid$(SBody, ScanPos, 1))
                            If (ChCode < 48) Or (ChCode > 57) Then Exit Do
                            ScanPos = ScanPos + 1
                        Loop
                    'Reject a component with no digits
                        If TokenStart = ScanPos Then GoTo Fail
                    'Coerce the numeric token once. Grammar is already known to
                    'be digits-only, so conversion failure is numerical range,
                    'not malformed syntax.
                        On Error GoTo RangeFail
                        QtyD = CDbl(Mid$(SBody, TokenStart, ScanPos - TokenStart))
                        On Error GoTo Fail
                    'Reject a quantity with no trailing unit
                        If ScanPos > BodyLen Then GoTo Fail
                    'Record the component by unit, rejecting any repeat
                        Select Case Mid$(SBody, ScanPos, 1)
                            Case "Y"
                                If SeenY Then Condition = KPR_COND_PILLAR_DUPLICATE_UNIT: GoTo Fail
                                SeenY = True
                                YearsD = QtyD
                            Case "M"
                                If SeenM Then Condition = KPR_COND_PILLAR_DUPLICATE_UNIT: GoTo Fail
                                SeenM = True
                                MonthsD = QtyD
                            Case "W"
                                If SeenW Then Condition = KPR_COND_PILLAR_DUPLICATE_UNIT: GoTo Fail
                                SeenW = True
                                WeeksD = QtyD
                            Case "D"
                                If SeenD Then Condition = KPR_COND_PILLAR_DUPLICATE_UNIT: GoTo Fail
                                SeenD = True
                                DaysD = QtyD
                            Case Else
                                GoTo Fail
                        End Select
                    'Count the component and step past the unit
                        TokenCount = TokenCount + 1
                        ScanPos = ScanPos + 1
                Loop

        End If

    'Unreachable in practice: the scanner fails before exiting with no
    'components, and a blank body was rejected earlier. Kept as a guard so
    'the TRUE return can never be reached without a parsed component.
        If TokenCount = 0 Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULTS
'------------------------------------------------------------------------------
    'Aggregate into locals first so a range failure never partially modifies
    'the caller's ByRef outputs.
        On Error GoTo RangeFail
        ParsedMonths = SignMul * ((12# * YearsD) + MonthsD)
        ParsedDays = SignMul * ((7# * WeeksD) + DaysD)
        On Error GoTo Fail

    'Contract: both outputs are assigned only after both aggregates succeed.
        TotalMonths = ParsedMonths
        TotalDays = ParsedDays
        Condition = KPR_COND_NONE
        TryPillar_Parse = True
        Exit Function

'------------------------------------------------------------------------------
' RANGE FAIL
'------------------------------------------------------------------------------
RangeFail:
    'The token grammar is valid; the numeric component or aggregate is not
    'representable in the parser's Double domain.
        Condition = KPR_COND_PILLAR_AGGREGATE_RANGE
        Resume Fail

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return FALSE per contract, leaving the outputs untouched
        TryPillar_Parse = False
    'Do not leave stale error state for the caller to observe
        Err.Clear

End Function


Public Function TryPillar_Format( _
    ByVal DtStart As Date, _
    ByVal DtEnd As Date, _
    ByVal Mode As KPR_PillarRounding, _
    ByRef TokenOut As String, _
    ByRef Condition As KPR_Condition) _
    As Boolean
'
'==============================================================================
'                                TryPillar_Format
'------------------------------------------------------------------------------
' PURPOSE
'   Formats the interval from DtStart to DtEnd as one canonical pillar token
'   under the requested rounding mode, or reports why no token exists.
'
' SIGNATURE
'   TryPillar_Format(DtStart, DtEnd, Mode, TokenOut, Condition) -> Boolean
'
' INPUTS
'   DtStart, DtEnd
'     Date-only values inside the supported window. Order is free; a negative
'     interval yields a "-" prefix.
'   Mode
'     KPR_ROUND_NEAREST, KPR_ROUND_FLOOR or KPR_ROUND_CEILING.
'
' OUTPUTS
'   TokenOut  (ByRef)   assigned ONLY on success
'   Condition (ByRef)   always assigned; KPR_COND_NONE on success
'
' RETURNS
'   Boolean
'     TRUE  => TokenOut assigned
'     FALSE => Condition is RESULT_WINDOW: the anchor the mode requires would
'              leave the supported window, so no candidate satisfies it
'
' CANDIDATE SET (contract section 8.4)
'   One set under every mode. Anchors are generated from the ORIGINAL DtStart
'   in the direction of DtEnd, using signed counts, so that an emitted negative
'   token re-parses with the same clipping DateFromPillar applies.
'
'       weeks    1W, 2W, 3W          (never 4W or beyond)
'       months   floor month  = largest positive count whose anchor does not
'                               pass DtEnd in the interval direction
'                ceiling month = floor + 1
'
'   An anchor outside the supported window is excluded, never approximated:
'   there is no nominal distance for an anchor that cannot be constructed.
'   Month anchors come from TryAddMonths so their clipping cannot diverge from
'   the parser.
'
' SELECTION
'   Exact day pillars: an absolute interval shorter than seven days returns
'   the exact "nD" token under every mode; rounding is not involved.
'
'       FLOOR    the candidate that does not pass DtEnd and lies farthest from
'                DtStart in the interval direction
'       CEILING  the candidate that reaches or passes DtEnd and lies closest
'                to DtEnd; none => RESULT_WINDOW
'       NEAREST  the candidate with the smallest calendar-day distance to DtEnd
'
'   Tie rules, identical under every mode: an equal outcome between a week
'   anchor and a month anchor takes the month; an equal outcome between two
'   month anchors takes the larger count. An exact anchor is exact under every
'   mode, because it wins each selection outright.
'
'   Under NEAREST the 3W / 1M boundary is therefore derived, not pinned: 24
'   days is always 3W and 27 days is always 1M whenever the 1M anchor is in
'   the window; 25 and 26 resolve by distance to that anchor, ties to the
'   month.
'
' WHY FLOOR AND NEAREST ALWAYS SUCCEED
'   For any interval of seven days or more, the 1W anchor lies strictly
'   between the two dates, both of which are in the window, so at least one
'   candidate exists and it does not pass DtEnd. Only CEILING can find every
'   in-window candidate short of DtEnd.
'
' ERROR POLICY
'   - Does not raise. The single failure path returns FALSE with a condition.
'
' DEPENDENCIES
'   - TryAddMonths, KPR_MIN_DATE, KPR_MAX_DATE
'   - KPR_Core_Err condition vocabulary
'
' NOTES
'   - Month counts of 12 or more format as years plus residual months.
'   - The old half-up "nearest week then cap" mechanism is gone. It made the
'     week family mean different things under different modes and produced a
'     3W / 1M boundary that was an artifact of integer division rather than
'     of distance.
'
' UPDATED
'   2026-09-01
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Dir             As Long     '+1 when DtEnd follows DtStart, else -1
    Dim SerialStart     As Double   'Serial of DtStart
    Dim SerialEnd       As Double   'Serial of DtEnd
    Dim TotalDays       As Long     'Absolute calendar days between the dates

    Dim CandSerial(1 To 5) As Double    'Anchor serials, in candidate order
    Dim CandIsMonth(1 To 5) As Boolean  'TRUE for a month anchor
    Dim CandCount(1 To 5) As Long       'Week or month count (unsigned)
    Dim CandN           As Long         'Number of admitted candidates

    Dim K               As Long     'Loop and count cursor
    Dim WeekSerial      As Double   'Candidate week anchor, tested before any Date conversion
    Dim FloorMonths     As Long     'Largest positive month count not passing DtEnd
    Dim Anchor          As Date     'Working anchor from TryAddMonths
    Dim Probe           As Long     'Working month count while locating the floor

    Dim Best            As Long     'Index of the selected candidate, 0 = none
    Dim Reach           As Double   'Signed progress of a candidate from DtStart
    Dim BestReach       As Double   'Progress of the current best
    Dim Dist            As Long     'Absolute day distance of a candidate to DtEnd
    Dim BestDist        As Long     'Distance of the current best
    Dim Passes          As Boolean  'TRUE when a candidate reaches or passes DtEnd
    Dim Better          As Boolean  'TRUE when candidate K displaces Best

    Dim nYears          As Long     'Year component of a month count
    Dim nMonths         As Long     'Residual month component

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default outcome is failure with an explicit condition
        TryPillar_Format = False
        Condition = KPR_COND_RESULT_WINDOW

    'Direction and magnitude from the ORIGINAL start date
        SerialStart = Int(CDbl(DtStart))
        SerialEnd = Int(CDbl(DtEnd))
        If SerialEnd >= SerialStart Then Dir = 1 Else Dir = -1
        TotalDays = CLng(Abs(SerialEnd - SerialStart))

'------------------------------------------------------------------------------
' EXACT DAY PILLARS
'------------------------------------------------------------------------------
    'Equal dates and short intervals are exact under every mode
        If TotalDays = 0 Then
            TokenOut = "0D"
            Condition = KPR_COND_NONE
            TryPillar_Format = True
            Exit Function
        End If
        If TotalDays < 7 Then
            TokenOut = SignPrefix(Dir) & CStr(TotalDays) & "D"
            Condition = KPR_COND_NONE
            TryPillar_Format = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' WEEK CANDIDATES
'------------------------------------------------------------------------------
    'Admit 1W, 2W and 3W when their anchors stay inside the window. The test
    'is on the serial as a Double: converting an out-of-range serial to a Date
    'raises, and an anchor that cannot be constructed must be excluded, not
    'allowed to abort the call.
        CandN = 0
        For K = 1 To 3
            WeekSerial = SerialStart + CDbl(Dir) * 7# * CDbl(K)
            If (WeekSerial >= CDbl(KPR_MIN_DATE)) And (WeekSerial <= CDbl(KPR_MAX_DATE)) Then
                CandN = CandN + 1
                CandSerial(CandN) = WeekSerial
                CandIsMonth(CandN) = False
                CandCount(CandN) = K
            End If
        Next K

'------------------------------------------------------------------------------
' MONTH CANDIDATES
'------------------------------------------------------------------------------
    'Locate the largest positive count whose anchor does not pass DtEnd,
    'starting from the calendar-month difference and correcting for clipping
        FloorMonths = Abs(DateDiff("m", DtStart, DtEnd))
        If FloorMonths < 1 Then FloorMonths = 1

    'Step down while the anchor passes DtEnd or cannot be constructed
        Do While FloorMonths >= 1
            If TryAddMonths(DtStart, Dir * FloorMonths, False, Anchor) Then
                If Not AnchorPasses(CDbl(Anchor), SerialEnd, Dir) Then Exit Do
            End If
            FloorMonths = FloorMonths - 1
        Loop

    'Step up while the next anchor still does not pass DtEnd
        Do
            Probe = FloorMonths + 1
            If Not TryAddMonths(DtStart, Dir * Probe, False, Anchor) Then Exit Do
            If AnchorPasses(CDbl(Anchor), SerialEnd, Dir) Then Exit Do
            FloorMonths = Probe
        Loop

    'Admit the floor month anchor, if any
        If FloorMonths >= 1 Then
            If TryAddMonths(DtStart, Dir * FloorMonths, False, Anchor) Then
                CandN = CandN + 1
                CandSerial(CandN) = CDbl(Anchor)
                CandIsMonth(CandN) = True
                CandCount(CandN) = FloorMonths
            End If
        End If

    'Admit the ceiling month anchor only if it can be constructed in-window
        If TryAddMonths(DtStart, Dir * (FloorMonths + 1), False, Anchor) Then
            CandN = CandN + 1
            CandSerial(CandN) = CDbl(Anchor)
            CandIsMonth(CandN) = True
            CandCount(CandN) = FloorMonths + 1
        End If

'------------------------------------------------------------------------------
' SELECT
'------------------------------------------------------------------------------
    'One pass; the mode decides the ordering and the tie rules are shared
        Best = 0
        For K = 1 To CandN

            'Signed progress from DtStart and absolute distance to DtEnd
                Reach = CDbl(Dir) * (CandSerial(K) - SerialStart)
                Dist = CLng(Abs(CandSerial(K) - SerialEnd))
                Passes = AnchorPasses(CandSerial(K), SerialEnd, Dir) Or (Dist = 0)

            'Eligibility and ordering by mode
                Better = False
                Select Case Mode

                    Case KPR_ROUND_FLOOR
                        'Must not pass; prefer the farthest reach
                            If Not AnchorPasses(CandSerial(K), SerialEnd, Dir) Then
                                If Best = 0 Then
                                    Better = True
                                ElseIf Reach > BestReach Then
                                    Better = True
                                ElseIf Reach = BestReach Then
                                    Better = TieToMonth(K, Best, CandIsMonth, CandCount)
                                End If
                            End If

                    Case KPR_ROUND_CEILING
                        'Must reach or pass; prefer the smallest reach
                            If Passes Then
                                If Best = 0 Then
                                    Better = True
                                ElseIf Reach < BestReach Then
                                    Better = True
                                ElseIf Reach = BestReach Then
                                    Better = TieToMonth(K, Best, CandIsMonth, CandCount)
                                End If
                            End If

                    Case Else
                        'NEAREST: smallest distance, ties by the shared rule
                            If Best = 0 Then
                                Better = True
                            ElseIf Dist < BestDist Then
                                Better = True
                            ElseIf Dist = BestDist Then
                                Better = TieToMonth(K, Best, CandIsMonth, CandCount)
                            End If

                End Select

            'Record the new best
                If Better Then
                    Best = K
                    BestReach = Reach
                    BestDist = Dist
                End If

        Next K

    'No candidate satisfied the mode: only CEILING can arrive here
        If Best = 0 Then Exit Function

'------------------------------------------------------------------------------
' FORMAT
'------------------------------------------------------------------------------
    'Weeks are a single token; months split into years and residual months
        If Not CandIsMonth(Best) Then
            TokenOut = SignPrefix(Dir) & CStr(CandCount(Best)) & "W"
        Else
            nYears = CandCount(Best) \ 12
            nMonths = CandCount(Best) Mod 12
            If nYears = 0 Then
                TokenOut = SignPrefix(Dir) & CStr(nMonths) & "M"
            ElseIf nMonths = 0 Then
                TokenOut = SignPrefix(Dir) & CStr(nYears) & "Y"
            Else
                TokenOut = SignPrefix(Dir) & CStr(nYears) & "Y" & CStr(nMonths) & "M"
            End If
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Contract: TRUE only when TokenOut was assigned
        Condition = KPR_COND_NONE
        TryPillar_Format = True

End Function

Private Function AnchorPasses( _
    ByVal AnchorSerial As Double, _
    ByVal SerialEnd As Double, _
    ByVal Dir As Long) _
    As Boolean
'
' TRUE when the anchor lies strictly beyond DtEnd in the interval direction.
'

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Strictly beyond, in the signed direction
        AnchorPasses = (CDbl(Dir) * (AnchorSerial - SerialEnd) > 0#)

End Function

Private Function TieToMonth( _
    ByVal Candidate As Long, _
    ByVal Best As Long, _
    ByRef IsMonth() As Boolean, _
    ByRef Count() As Long) _
    As Boolean
'
' The shared tie rule: month beats week; between months, the larger count.
' Returns TRUE when Candidate should displace Best.
'

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Month over week
        If IsMonth(Candidate) And Not IsMonth(Best) Then
            TieToMonth = True

    'Between months, the larger count (the ceiling anchor)
        ElseIf IsMonth(Candidate) And IsMonth(Best) Then
            TieToMonth = (Count(Candidate) > Count(Best))

    'Otherwise keep the current best
        Else
            TieToMonth = False
        End If

End Function

Private Function SignPrefix( _
    ByVal Dir As Long) _
    As String
'
' "-" for a negative interval, empty otherwise.
'

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Only a negative interval carries a prefix; "+" is never emitted
        If Dir < 0 Then SignPrefix = "-" Else SignPrefix = vbNullString

End Function

Private Function Pillar_AsciiTrim( _
    ByVal TextIn As String) _
    As String
'
' Removes leading and trailing ASCII spaces and tabs only, so pillar text is
' trimmed identically in every locale.
'
' A private twin of the helper in KPR_Core_Parse: the dependency matrix does
' not let this module call that one, and a ten-line helper is cheaper than a
' shared utility module. Keep the two in step.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim First       As Long     'First kept position
    Dim Last        As Long     'Last kept position
    Dim Ch          As String   'Character under test

'------------------------------------------------------------------------------
' SCAN
'------------------------------------------------------------------------------
    First = 1
    Last = Len(TextIn)
    Do While First <= Last
        Ch = Mid$(TextIn, First, 1)
        If (Ch <> " ") And (Ch <> vbTab) Then Exit Do
        First = First + 1
    Loop
    Do While Last >= First
        Ch = Mid$(TextIn, Last, 1)
        If (Ch <> " ") And (Ch <> vbTab) Then Exit Do
        Last = Last - 1
    Loop

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    If Last < First Then Pillar_AsciiTrim = vbNullString Else Pillar_AsciiTrim = Mid$(TextIn, First, Last - First + 1)

End Function

Private Function Pillar_AsciiUpper( _
    ByVal TextIn As String) _
    As String
'
' Folds a-z to A-Z by code point. UCase$ consults the host locale and is not
' used anywhere in pillar parsing. Private twin of the KPR_Core_Parse helper;
' keep the two in step.
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim I           As Long     'Character cursor
    Dim Code        As Long     'Code point under test
    Dim Result      As String   'Accumulated output

'------------------------------------------------------------------------------
' FOLD
'------------------------------------------------------------------------------
    Result = TextIn
    For I = 1 To Len(Result)
        Code = AscW(Mid$(Result, I, 1))
        If (Code >= 97) And (Code <= 122) Then Mid$(Result, I, 1) = ChrW$(Code - 32)
    Next I

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    Pillar_AsciiUpper = Result

End Function
