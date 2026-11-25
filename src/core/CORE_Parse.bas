Attribute VB_Name = "CORE_Parse"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: CORE_Parse
'------------------------------------------------------------------------------
' PURPOSE
'   Reads one argument of a date-layer function: converts an already-unwrapped
'   scalar into a date, a whole number, a Boolean or a rounding token, or
'   reports the condition that rejects it.
'
' WHY THIS EXISTS
'   Worksheet arguments arrive as many runtime types for the same intent. This
'   module holds the one acceptance policy for each argument kind, so callers
'   never re-implement type logic and no runtime error reaches the worksheet.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - TryParseDateScalar       scalar to a date-only, in-window VBA Date
'     - TryParseLongScalar       scalar to a VBA Long
'     - TryParseBoolControl      optional control to a Boolean, or its default
'     - TryParseRoundingControl  Opt_Rounding to NEAREST, FLOOR or CEILING
'     - WINDOW_MIN_SERIAL, WINDOW_MAX_SERIAL, WINDOW_MIN_YEAR, WINDOW_MAX_YEAR
'       the supported date window, which this module owns
'
' ERROR POLICY
'   The parsers do not raise. Every rejection is reported through the Boolean
'   result and the Condition output.
'
' DEPENDENCIES
'   CORE_Err, for the KPRCondition vocabulary. No Excel object model. Data
'   contract: docs/methodology/DATA_CONTRACT.md, Function arguments (Date
'   values, Integer values, Control values); calculation contracts: common
'   rules, Dates (the 1900-03-01 to 9999-12-31 window)
'
' NOTES
'   - Shape is not handled here. The caller passes a scalar that the array
'     engine has already unwrapped
'   - Incoming Excel errors are not rejections. The element boundary returns
'     them unchanged before it calls a parser
'   - A single-cell Range reaches the parsers read through Value2, so a
'     date-formatted cell presents as vbDouble, not vbDate. Both types take the
'     same arithmetic path and give the same date-only result
'   - The supported window is owned here and restated nowhere: the date core
'     reads the same constants. The serial bounds are the 1900 date system's
'     serials of 1900-03-01 and 9999-12-31, which VBA's Date names the same way
'   - Option Private Module keeps every member out of the Excel function list.
'     Nothing here is supported API
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'The Long range, tested before any narrowing conversion
    Private Const LONG_MIN_VALUE    As Double = -2147483648#    'Smallest VBA Long
    Private Const LONG_MAX_VALUE    As Double = 2147483647#     'Largest VBA Long

    'Largest distance from a whole number still read as that whole number;
    'absorbs binary floating-point residue such as 2.9999999999999996
    Private Const INTEGER_TOLERANCE As Double = 1# / 1000000000#    '1E-9

'------------------------------------------------------------------------------
' PUBLIC CONSTANTS
'------------------------------------------------------------------------------
    'The supported date window, 1900-03-01 to 9999-12-31, both included. The
    'lower bound excludes Excel's serial 60, the 1900-02-29 that never existed
    Public Const WINDOW_MIN_SERIAL  As Double = 61#             '1900-03-01
    Public Const WINDOW_MAX_SERIAL  As Double = 2958465#        '9999-12-31
    Public Const WINDOW_MIN_YEAR    As Long = 1900              'First supported year
    Public Const WINDOW_MAX_YEAR    As Long = 9999              'Last supported year


'
'------------------------------------------------------------------------------
'
'                              DATE VALUE PARSING
'
'------------------------------------------------------------------------------
'

Public Function TryParseDateScalar( _
    ByVal ScalarIn As Variant, _
    ByRef ParsedDate As Date, _
    ByRef Condition As KPRCondition) As Boolean
'
'------------------------------------------------------------------------------
'                             PARSE A DATE SCALAR
'------------------------------------------------------------------------------
' PURPOSE
'   Converts one already-unwrapped scalar into a normalized, in-window VBA Date
'   under the data contract's Date values rules, reporting the exact condition
'   on failure.
'
' WHY THIS EXISTS
'   Every date position of the library shares one strict, locale-independent
'   acceptance policy, so no caller re-implements type logic.
'
' INPUTS
'   ScalarIn
'     A scalar Variant already isolated by the array engine.
'
'     Accepted:
'       - native VBA Date
'       - native numeric 1900-system serial
'       - text in exactly the form YYYY-MM-DD
'
'     Rejected:
'       - locale-formatted text            DATE_TEXT_LOCALE
'       - numeric-looking text             DATE_TEXT_NUMERIC
'       - any other malformed text         DATE_TEXT_FORMAT
'       - exact ISO naming a date that does not exist
'                                          DATE_TEXT_IMPOSSIBLE
'       - Empty / blank required cell      INPUT_BLANK_REQUIRED
'       - Null, Boolean, object            DATE_TYPE_REJECTED
'       - accepted value outside the window
'                                          DATE_WINDOW
'
'     An incoming Excel error never reaches this routine. It is detected and
'     returned verbatim at the public/element boundary.
'
'   ParsedDate
'     Output, see RETURNS.
'
'   Condition
'     Output, see RETURNS.
'
' RETURNS
'   Boolean
'     TRUE  => ParsedDate assigned and Condition is cdNone
'     FALSE => ParsedDate untouched and Condition names the failure
'
'   ParsedDate (ByRef)
'     Assigned ONLY on success, date-only and inside the supported window.
'
'   Condition (ByRef)
'     Always assigned. cdNone on success, otherwise the failing
'     condition.
'
' BEHAVIOR
'   - Text is matched against the exact ten-character ISO form with ASCII digits
'     and hyphens at fixed positions. No trimming, no alternate separator, no
'     omitted zero, no time suffix, no localized digit.
'   - Numeric-looking text is rejected before any numeric interpretation, so a
'     serial written as text is never silently reinterpreted.
'   - A native Date or numeric serial has its time component removed by flooring
'     the serial, then the supported window is applied to the resulting
'     date-only value.
'   - No locale-sensitive text conversion is used. IsDate, DateValue, CVDate and
'     IsNumeric do not appear; CDate is reached only after a numeric serial has
'     been normalized and window-validated.
'
' ERROR POLICY
'   - Does not raise. Every failure path returns FALSE with a condition.
'   - The caller maps the condition through CORE_Err.ErrForCondition.
'
' DEPENDENCIES
'   CORE_Err (KPRCondition), TryClassifyDateText, FloorSerial
'
' NOTES
'   - Component validity is checked by round trip rather than by a private
'     calendar. DateSerial is called only once the year is known to be inside
'     the window, and the result must return the same year, month and day.
'     2025-02-29 rolls to 2025-03-01 and is rejected, so rollover can never pass
'     silently, and no leap-year logic is duplicated here.
'   - The year gate precedes construction for a second reason: DateSerial
'     reinterprets a year below 100 as a 20th-century year, so DateSerial(50,1,1)
'     is 1950. A syntactically exact 0050-01-01 is reported as DATE_WINDOW
'     without ever being constructed.
'   - The window bounds are this module's own public constants
'
' UPDATED
'   2026-09-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Parse.TryParseDateScalar"

    Dim VT              As VbVarType    'Cached VarType of the incoming scalar
    Dim Candidate       As Date         'Parsed candidate before assignment
    Dim Serial          As Double       'Numeric working value for serial inputs
    Dim S               As String       'Text payload, used without trimming

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default outcome is failure with an explicit condition
        TryParseDateScalar = False
        Condition = cdDateTypeRejected

'------------------------------------------------------------------------------
' CLASSIFY SCALAR
'------------------------------------------------------------------------------
    'Classify the already-unwrapped payload
        VT = VBA.VarType(ScalarIn)
    'Route each runtime type to its acceptance rule
        Select Case VT

            'A blank required cell is contract-invalid, not a zero serial
            Case vbEmpty
                Condition = cdInputBlankRequired
                Exit Function

            'Null, Boolean and objects are never dates. An incoming error should
            'have been intercepted at the boundary; classify rather than guess.
            Case vbNull, vbBoolean, vbObject, vbDataObject, vbError
                Condition = cdDateTypeRejected
                Exit Function

            'Native Date => remove the time component from the serial
            Case vbDate
                Serial = FloorSerial(VBA.CDbl(ScalarIn))

            'Text is accepted in exactly one form, checked without locale help;
            'classify the text before attempting any conversion
            Case vbString
                S = VBA.CStr(ScalarIn)
                If Not TryClassifyDateText(S, Serial, Condition) Then Exit Function

            'Native numeric serial => remove the time component
            Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
                Serial = FloorSerial(VBA.CDbl(ScalarIn))

            'LongLong is a native numeric subtype on VBA7. The test lives inside
            'the Case body rather than as its own Case, so the conditional
            'directive never sits between two Case clauses, and it is gated on
            'VBA7 rather than Win64 because the vbLongLong constant exists in
            'the VBA7 enum regardless of Office bitness.
            Case Else
#If VBA7 Then
                If VT = vbLongLong Then
                    Serial = FloorSerial(VBA.CDbl(ScalarIn))
                Else
                    Condition = cdDateTypeRejected
                    Exit Function
                End If
#Else
                Condition = cdDateTypeRejected
                Exit Function
#End If

        End Select

'------------------------------------------------------------------------------
' APPLY SUPPORTED WINDOW
'------------------------------------------------------------------------------
    'One window test for every accepted path, on the date-only value
        If (Serial < WINDOW_MIN_SERIAL) Or (Serial > WINDOW_MAX_SERIAL) Then
            Condition = cdDateWindow
            Exit Function
        End If
    'Safe: the serial is inside the window, so conversion cannot overflow
        Candidate = VBA.CDate(Serial)

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once the value is fully validated
        ParsedDate = Candidate
        Condition = cdNone
    'Contract: TRUE only when ParsedDate was assigned
        TryParseDateScalar = True

End Function

Private Function TryClassifyDateText( _
    ByVal TextIn As String, _
    ByRef SerialOut As Double, _
    ByRef Condition As KPRCondition) As Boolean
'
'------------------------------------------------------------------------------
'                            CLASSIFY ISO DATE TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Accepts text only in the exact form YYYY-MM-DD and returns its date serial.
'
' WHY THIS EXISTS
'   Keeps the text branch of TryParseDateScalar, including the provenance of
'   each text rejection, in one routine.
'
' INPUTS
'   TextIn
'     The raw text payload, untrimmed. Surrounding whitespace is a rejection,
'     not something to clean up.
'
'   SerialOut
'     Output, see RETURNS.
'
'   Condition
'     Output, see RETURNS.
'
' RETURNS
'   Boolean
'     TRUE  => SerialOut assigned
'     FALSE => Condition names the rejection
'
'   SerialOut (ByRef)
'     Assigned ONLY on success, to the date serial of the parsed value.
'
'   Condition (ByRef)
'     Always assigned on failure.
'
' BEHAVIOR
'   - Length must be exactly ten characters.
'   - Positions 5 and 8 must be hyphens; the other eight must be ASCII digits.
'   - Text that is numeric as a whole is reported as DATE_TEXT_NUMERIC, and text
'     shaped like a day-month-year or month-day-year date with "/" or "."
'     separators is reported as DATE_TEXT_LOCALE, so provenance survives even
'     though both return #VALUE!. Every other malformed text, including a
'     year-first "2024/02/10" or a time suffix, is DATE_TEXT_FORMAT.
'   - The year is gated against the supported window before construction; the
'     month, day and leap day are validated by DateSerial round trip.
'
' ERROR POLICY
'   - Does not raise. Digit validation precedes every numeric conversion, so no
'     conversion in this routine can fail.
'
' DEPENDENCIES
'   LooksNumeric, LooksLocaleDate
'
' NOTES
'   None
'
' UPDATED
'   2026-11-11
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Parse.TryClassifyDateText"

    Dim I               As Long         'Character cursor
    Dim Ch              As String       'Single character under test
    Dim YearPart        As Long         'Parsed year component
    Dim MonthPart       As Long         'Parsed month component
    Dim DayPart         As Long         'Parsed day component
    Dim RoundTrip       As Date         'Constructed date, verified componentwise

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default outcome is failure with an explicit condition
        TryClassifyDateText = False
        Condition = cdDateTextFormat

'------------------------------------------------------------------------------
' REJECT NON-ISO SHAPES
'------------------------------------------------------------------------------
    'Empty text carries no date
        If VBA.Len(TextIn) = 0 Then Exit Function

    'A numeric-looking string is never reinterpreted as a serial.
    'The test is our own, not the host's: VBA IsNumeric consults the locale and
    'reads "31.12.2026" as thirty-one million under an Italian decimal
    'convention, which would misclassify a locale-formatted date as numeric.
        If LooksNumeric(TextIn) Then
            Condition = cdDateTextNumeric
            Exit Function
        End If

    'Distinguish a locale-shaped date from generic malformed text for provenance.
    'Only the whole d/m/y or m/d/y shape is a locale marker (contract 3.1): a
    'separator alone is not, and neither is whitespace, so padded, time-suffixed
    'and year-first text stays DATE_TEXT_FORMAT with every other malformed shape.
        If LooksLocaleDate(TextIn) Then
            Condition = cdDateTextLocale
            Exit Function
        End If

    'The accepted form is exactly ten characters
        If VBA.Len(TextIn) <> 10 Then Exit Function

    'Separators sit at fixed positions
        If VBA.Mid$(TextIn, 5, 1) <> "-" Then Exit Function
        If VBA.Mid$(TextIn, 8, 1) <> "-" Then Exit Function

    'Every other position must be an ASCII digit
        For I = 1 To 10

            'Skip the two separator positions already verified
                If (I <> 5) And (I <> 8) Then

                    'Compare by code point so no locale digit is accepted
                        Ch = VBA.Mid$(TextIn, I, 1)
                        If (Ch < "0") Or (Ch > "9") Then Exit Function

                End If

        Next I

'------------------------------------------------------------------------------
' VALIDATE COMPONENTS
'------------------------------------------------------------------------------
    'Safe: every character position has been verified as a digit
        YearPart = VBA.CLng(VBA.Left$(TextIn, 4))
        MonthPart = VBA.CLng(VBA.Mid$(TextIn, 6, 2))
        DayPart = VBA.CLng(VBA.Right$(TextIn, 2))

    'Gate the year before construction: DateSerial reinterprets years below 100
        If (YearPart < WINDOW_MIN_YEAR) Or (YearPart > WINDOW_MAX_YEAR) Then
            Condition = cdDateWindow
            Exit Function
        End If

    'Reject an impossible month before construction
        If (MonthPart < 1) Or (MonthPart > 12) Then
            Condition = cdDateTextImpossible
            Exit Function
        End If

    'Reject an impossible day the same way
        If (DayPart < 1) Or (DayPart > 31) Then
            Condition = cdDateTextImpossible
            Exit Function
        End If

    'Construct, then verify componentwise so rollover cannot pass silently
        RoundTrip = VBA.DateSerial(YearPart, MonthPart, DayPart)

    'A rolled-over date differs in at least one component
        If (VBA.Year(RoundTrip) <> YearPart) Or _
            (VBA.Month(RoundTrip) <> MonthPart) Or _
            (VBA.Day(RoundTrip) <> DayPart) Then
            Condition = cdDateTextImpossible
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Report the serial; the caller applies the supported window
        SerialOut = VBA.CDbl(RoundTrip)
        Condition = cdNone
    'Contract: TRUE only when SerialOut was assigned
        TryClassifyDateText = True

End Function

'
'------------------------------------------------------------------------------
'
'                         INTEGER AND CONTROL PARSING
'
'------------------------------------------------------------------------------
'

Public Function TryParseLongScalar( _
    ByVal ScalarIn As Variant, _
    ByRef ParsedLong As Long, _
    ByRef Condition As KPRCondition) As Boolean
'
'------------------------------------------------------------------------------
'                             PARSE A LONG SCALAR
'------------------------------------------------------------------------------
' PURPOSE
'   Converts one already-unwrapped scalar into a Long under the data contract's
'   Integer values rules, reporting the exact condition on failure.
'
' WHY THIS EXISTS
'   Every integer position of the library shares one strict acceptance policy,
'   so no caller re-implements type or tolerance logic.
'
' INPUTS
'   ScalarIn
'     Accepted: native numeric inside the VBA Long range that is integral or
'     within INTEGER_TOLERANCE (1E-9) of an integer; it is read as that
'     nearest integer.
'
'     Rejected:
'       - numeric farther than 1E-9 from an integer
'                                         INTEGER_FRACTION
'       - numeric outside the Long range  INTEGER_RANGE
'       - Boolean, Date, ANY text including numeric-looking text, Null, object
'                                         INTEGER_TYPE_REJECTED
'       - Empty / blank required cell     INPUT_BLANK_REQUIRED
'
'   ParsedLong
'     Output, see RETURNS.
'
'   Condition
'     Output, see RETURNS.
'
' RETURNS
'   Boolean
'     TRUE  => ParsedLong assigned and Condition is cdNone
'     FALSE => ParsedLong untouched and Condition names the failure
'
'   ParsedLong (ByRef)   assigned ONLY on success
'   Condition  (ByRef)   always assigned
'
' BEHAVIOR
'   - Range is tested BEFORE integrality, per the data contract: a value
'     outside the Long range returns INTEGER_RANGE and #NUM! even when it also
'     has a fractional part.
'   - The nearest integer is Int(X + 0.5); a value within the tolerance is
'     never near a half, so no tie rule is involved. CLng is reached only on
'     that whole number, so its round-half-to-even behaviour is never
'     observed.
'
' ERROR POLICY
'   - Does not raise. Every failure path returns FALSE with a condition.
'
' DEPENDENCIES
'   CORE_Err (KPRCondition)
'
' NOTES
'   - Boolean is rejected rather than coerced. TRUE would otherwise become -1,
'     which is a silently plausible shift count.
'   - Text is rejected even when it looks like a number, mirroring the date
'     rule that numeric-looking text is never reinterpreted.
'
' UPDATED
'   2026-09-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Parse.TryParseLongScalar"

    Dim VT              As VbVarType    'Cached VarType of the incoming scalar
    Dim X               As Double       'Numeric working value
    Dim Nearest         As Double       'Nearest whole number to X
    Dim DecIn           As Variant      'Decimal payload, compared without a lossy Double step
    Dim DecNearest      As Variant      'Nearest whole number to DecIn

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default outcome is failure with an explicit condition
        TryParseLongScalar = False
        Condition = cdIntegerTypeRejected

'------------------------------------------------------------------------------
' CLASSIFY SCALAR
'------------------------------------------------------------------------------
    'Classify the already-unwrapped payload
        VT = VBA.VarType(ScalarIn)
    'Route each runtime type to its acceptance rule
        Select Case VT

            'A blank required cell is contract-invalid, not a zero
            Case vbEmpty
                Condition = cdInputBlankRequired
                Exit Function

            'Accepted numeric types; coerce once to a common working type
            Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
                X = VBA.CDbl(ScalarIn)

            'LongLong is accepted on VBA7 and the Long-range gate below still
            'owns whether its value can be taken. Everything else (Boolean,
            'Date, text, Null, objects, errors) is rejected. See
            'TryParseDateScalar for why the directive sits inside the Case body
            'and is gated on VBA7.
            Case Else
#If VBA7 Then
                If VT = vbLongLong Then
                    X = VBA.CDbl(ScalarIn)
                Else
                    Condition = cdIntegerTypeRejected
                    Exit Function
                End If
#Else
                Condition = cdIntegerTypeRejected
                Exit Function
#End If

        End Select

'------------------------------------------------------------------------------
' VALIDATE RANGE THEN INTEGRALITY
'------------------------------------------------------------------------------
    'Range precedes integrality when both apply (data contract, Integer values)
        If (X < LONG_MIN_VALUE) Or (X > LONG_MAX_VALUE) Then
            Condition = cdIntegerRange
            Exit Function
        End If

    'No truncation and no rounding of a real fraction: only floating-point
    'residue within the tolerance is read as the nearest whole number. A
    'Decimal carries more digits than a Double, so its distance is measured
    'in Decimal; converting first could push a value just inside the
    'tolerance outside it.
        If VT = vbDecimal Then
            DecIn = VBA.CDec(ScalarIn)
            DecNearest = VBA.Int(DecIn + VBA.CDec(0.5))
            If VBA.Abs(DecIn - DecNearest) > VBA.CDec(1) / VBA.CDec(1000000000) Then
                Condition = cdIntegerFraction
                Exit Function
            End If
            Nearest = VBA.CDbl(DecNearest)
        Else
            Nearest = VBA.Int(X + 0.5)
            If VBA.Abs(X - Nearest) > INTEGER_TOLERANCE Then
                Condition = cdIntegerFraction
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Safe: Nearest is whole and inside the Long range, because X is
        ParsedLong = VBA.CLng(Nearest)
        Condition = cdNone
    'Contract: TRUE only when ParsedLong was assigned
        TryParseLongScalar = True

End Function

Public Function TryParseBoolControl( _
    ByVal ControlIn As Variant, _
    ByVal DefaultValue As Boolean, _
    ByRef ParsedBool As Boolean, _
    ByRef Condition As KPRCondition) As Boolean
'
'------------------------------------------------------------------------------
'                           PARSE A BOOLEAN CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves a present Boolean optional control under the data contract's Control values rules.
'
' WHY THIS EXISTS
'   A Boolean control must accept a native Boolean only. Excel coercion of an
'   As Boolean declaration would accept 0, 1 and "TRUE" before the function is
'   entered (see NOTES).
'
' INPUTS
'   ControlIn
'     Accepted: Empty, meaning the documented default, or a native Boolean.
'
'     Rejected with CONTROL_TYPE_REJECTED: numeric 0 and 1, Boolean-looking text
'     such as "TRUE", Null, dates, objects. At the public boundary an object
'     never arrives here: the array engine reports it as SHAPE_UNSUPPORTED.
'
'   DefaultValue
'     The documented default for this control, selected when ControlIn is Empty.
'
'   ParsedBool
'     Output, see RETURNS.
'
'   Condition
'     Output, see RETURNS.
'
' RETURNS
'   Boolean
'     TRUE  => ParsedBool assigned and Condition is cdNone
'     FALSE => ParsedBool untouched and Condition names the failure
'
'   ParsedBool (ByRef)   assigned ONLY on success
'   Condition  (ByRef)   always assigned
'
' BEHAVIOR
'   - An omitted argument never reaches this routine: the public declaration
'     supplies its default, and a forwarded Missing is resolved to the default
'     by the caller. Omitted, Missing and
'     Empty select the same documented default.
'   - An incoming Excel error is NOT classified here. A control carrying an
'     error propagates verbatim as a call-level result, which is a decision for
'     the boundary rather than the parser.
'
' ERROR POLICY
'   - Does not raise. Every failure path returns FALSE with a condition.
'
' DEPENDENCIES
'   CORE_Err (KPRCondition)
'
' NOTES
'   - Accepting a native Boolean only is what makes the control strict. The
'     previous As Boolean declaration accepted 0, 1 and "TRUE" through Excel
'     coercion before the function was ever entered.
'
' UPDATED
'   2026-11-11
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Parse.TryParseBoolControl"

    Dim VT              As VbVarType    'Cached VarType of the incoming control

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default outcome is failure with an explicit condition
        TryParseBoolControl = False
        Condition = cdControlTypeRejected

'------------------------------------------------------------------------------
' CLASSIFY CONTROL
'------------------------------------------------------------------------------
    'Classify the payload
        VT = VBA.VarType(ControlIn)
    'Route each runtime type to its acceptance rule
        Select Case VT

            'A blank control selects the documented default
            Case vbEmpty
                ParsedBool = DefaultValue

            'The only accepted explicit form
            Case vbBoolean
                ParsedBool = VBA.CBool(ControlIn)

            'Numeric 0 / 1, Boolean-looking text, Null, dates and objects
            Case Else
                Condition = cdControlTypeRejected
                Exit Function

        End Select

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Contract: TRUE only when ParsedBool was assigned
        Condition = cdNone
        TryParseBoolControl = True

End Function

Public Function TryParseRoundingControl( _
    ByVal ControlIn As Variant, _
    ByRef TokenOut As String, _
    ByRef Condition As KPRCondition) As Boolean
'
'------------------------------------------------------------------------------
'                          PARSE THE ROUNDING CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves a present Opt_Rounding control to one of the three normalized
'   tokens NEAREST, FLOOR or CEILING, under the data contract's Control values rules.
'
' WHY THIS EXISTS
'   The Opt_Rounding text must resolve to the same token in every locale, so
'   it is trimmed and folded here without the host's locale-aware functions.
'
' INPUTS
'   ControlIn
'     Accepted: Empty, meaning the documented default NEAREST, or text whose
'     outer whitespace trims and whose ASCII-folded form is one of the three
'     tokens.
'
'     Rejected:
'       - any non-text, non-Empty type      CONTROL_TYPE_REJECTED
'       - text that is not one of the three CONTROL_TOKEN_UNKNOWN
'
'   TokenOut
'     Output, see RETURNS.
'
'   Condition
'     Output, see RETURNS.
'
' RETURNS
'   Boolean
'     TRUE  => TokenOut assigned and Condition is cdNone
'     FALSE => TokenOut untouched and Condition names the failure
'
'   TokenOut  (ByRef)   assigned ONLY on success, always upper-case ASCII
'   Condition (ByRef)   always assigned
'
' BEHAVIOR
'   - Case folding is ASCII-only: a-z map to A-Z and nothing else changes.
'     UCase$ is deliberately not used, because it consults the host locale and
'     under Turkish casing the i of "ceiling" would fold to the dotted capital
'     I (U+0130) rather than to I.
'   - Trimming removes leading and trailing ASCII spaces and tabs only.
'   - An incoming Excel error is NOT classified here; the boundary propagates
'     it before this routine is called.
'   - The result is text rather than an enum because this module may not
'     depend on CORE_Dates, which owns the rounding vocabulary. The facade
'     performs the final mapping.
'
' ERROR POLICY
'   - Does not raise. Every failure path returns FALSE with a condition.
'
' DEPENDENCIES
'   CORE_Err (KPRCondition), AsciiTrim, AsciiUpper
'
' NOTES
'   None
'
' UPDATED
'   2026-10-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Parse.TryParseRoundingControl"

    Dim Folded          As String       'Trimmed, ASCII-folded candidate token

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default outcome is failure with an explicit condition
        TryParseRoundingControl = False
        Condition = cdControlTypeRejected

'------------------------------------------------------------------------------
' CLASSIFY CONTROL
'------------------------------------------------------------------------------
    'Route the control by its runtime type
        Select Case VBA.VarType(ControlIn)

            'A blank control selects the documented default
            Case vbEmpty
                Folded = "NEAREST"

            'Normalize without consulting the locale
            Case vbString
                Folded = AsciiUpper(AsciiTrim(VBA.CStr(ControlIn)))

            'Numbers, Booleans, dates, Null and objects are not tokens
            Case Else
                Condition = cdControlTypeRejected
                Exit Function

        End Select

'------------------------------------------------------------------------------
' MATCH TOKEN
'------------------------------------------------------------------------------
    'Exactly three tokens are recognized
        Select Case Folded

            'A recognized token is returned as folded
            Case "NEAREST", "FLOOR", "CEILING"
                TokenOut = Folded

            'Any other text is an unknown token
            Case Else
                Condition = cdControlTokenUnknown
                Exit Function

        End Select

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Contract: TRUE only when TokenOut was assigned
        Condition = cdNone
        TryParseRoundingControl = True

End Function

Private Function AsciiTrim(ByVal TextIn As String) As String

'
'------------------------------------------------------------------------------
'                          TRIM ASCII SPACES AND TABS
'------------------------------------------------------------------------------
' PURPOSE
'   Removes leading and trailing ASCII spaces and tabs.
'
' WHY THIS EXISTS
'   Nothing else counts as whitespace here, so the result is identical in
'   every locale.
'
' INPUTS
'   TextIn
'     Text to trim.
'
' RETURNS
'   TextIn without its leading and trailing spaces and tabs; vbNullString when
'   nothing else remains.
'
' BEHAVIOR
'   Scans inward from both ends past spaces and tabs, then returns the span
'   between them.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-09-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Parse.AsciiTrim"

    Dim StartAt         As Long         'First kept position
    Dim EndAt           As Long         'Last kept position
    Dim Ch              As String       'Character under test

'------------------------------------------------------------------------------
' SCAN
'------------------------------------------------------------------------------
    'Start from the whole text
        StartAt = 1
        EndAt = VBA.Len(TextIn)
    'Skip leading spaces and tabs
        Do While StartAt <= EndAt
            Ch = VBA.Mid$(TextIn, StartAt, 1)
            If (Ch <> " ") And (Ch <> vbTab) Then Exit Do
            StartAt = StartAt + 1
        Loop
    'Skip trailing spaces and tabs
        Do While EndAt >= StartAt
            Ch = VBA.Mid$(TextIn, EndAt, 1)
            If (Ch <> " ") And (Ch <> vbTab) Then Exit Do
            EndAt = EndAt - 1
        Loop

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the kept span, or an empty string when nothing is kept
        If EndAt < StartAt Then
            AsciiTrim = vbNullString
        Else
            AsciiTrim = VBA.Mid$(TextIn, StartAt, EndAt - StartAt + 1)
        End If

End Function

Private Function AsciiUpper(ByVal TextIn As String) As String

'
'------------------------------------------------------------------------------
'                       FOLD ASCII LETTERS TO UPPER CASE
'------------------------------------------------------------------------------
' PURPOSE
'   Folds a-z to A-Z by code point and leaves every other character untouched.
'
' WHY THIS EXISTS
'   Locale-independent by construction; see TryParseRoundingControl.
'
' INPUTS
'   TextIn
'     Text to fold.
'
' RETURNS
'   TextIn with each ASCII letter a-z replaced by its upper-case form.
'
' BEHAVIOR
'   Replaces each lower-case ASCII letter in a copy of the text, in place.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-09-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Parse.AsciiUpper"

    Dim I               As Long         'Character cursor
    Dim Code            As Long         'Code point under test
    Dim Result          As String       'Accumulated output

'------------------------------------------------------------------------------
' FOLD
'------------------------------------------------------------------------------
    'Replace each lower-case ASCII letter in place
        Result = TextIn
        For I = 1 To VBA.Len(Result)
            Code = VBA.AscW(VBA.Mid$(Result, I, 1))
            If (Code >= 97) And (Code <= 122) Then Mid$(Result, I, 1) = VBA.ChrW$(Code - 32)
        Next I

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the folded text
        AsciiUpper = Result

End Function

Private Function FloorSerial(ByVal SerialIn As Double) As Double

'
'------------------------------------------------------------------------------
'                             FLOOR A DATE SERIAL
'------------------------------------------------------------------------------
' PURPOSE
'   Removes the time component from a date serial.
'
' WHY THIS EXISTS
'   Every accepted date is a whole day, so the time of a native Date or a
'   numeric serial is removed before the window is applied.
'
' INPUTS
'   SerialIn
'     A date serial, with or without a time fraction.
'
' RETURNS
'   Double
'     The largest whole serial not greater than SerialIn.
'
' BEHAVIOR
'   Floors the serial.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   - Int floors; Fix truncates toward zero. They agree for every serial inside
'     the supported window, which is entirely positive, but Int is used so that
'     a negative serial cannot move toward the window before being gated
'
' UPDATED
'   2026-09-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Floor rather than truncate toward zero
        FloorSerial = VBA.Int(SerialIn)

End Function

Private Function LooksNumeric(ByVal TextIn As String) As Boolean

'
'------------------------------------------------------------------------------
'                            TEST FOR NUMERIC TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether text is numeric under a fixed, locale-independent rule.
'
' WHY THIS EXISTS
'   VBA IsNumeric asks the host how to read the text. Under a locale that uses
'   "." as a group separator it accepts "31.12.2026" as a number, which would
'   classify a locale-formatted date as DATE_TEXT_NUMERIC and lose the
'   provenance the contract requires. It also accepts currency symbols,
'   thousands separators and surrounding whitespace, none of which this library
'   treats as numeric. The rule here is fixed and identical in every region.
'
' INPUTS
'   TextIn
'     Text to test.
'
' RETURNS
'   Boolean
'     TRUE  => the text is numeric under the rule above
'     FALSE => otherwise
'
' BEHAVIOR
'   Accepted shape: an optional leading + or -, ASCII digits, and at most one
'   "." acting as a decimal point. At least one digit must be present.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-09-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Parse.LooksNumeric"

    Dim I               As Long         'Character cursor
    Dim Ch              As String       'Single character under test
    Dim Digits          As Long         'Count of digits seen
    Dim Points          As Long         'Count of decimal points seen
    Dim StartAt         As Long         'First position after any sign

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default is not numeric
        LooksNumeric = False

    'Empty text is not numeric
        If VBA.Len(TextIn) = 0 Then Exit Function

    'Skip a single leading sign
        StartAt = 1
        Ch = VBA.Left$(TextIn, 1)
        If (Ch = "+") Or (Ch = "-") Then StartAt = 2

    'A sign on its own is not numeric
        If StartAt > VBA.Len(TextIn) Then Exit Function

'------------------------------------------------------------------------------
' SCAN
'------------------------------------------------------------------------------
    'Every remaining character must be a digit or the single decimal point
        For I = StartAt To VBA.Len(TextIn)

            'Read one character
                Ch = VBA.Mid$(TextIn, I, 1)

            'Digits accumulate
                If (Ch >= "0") And (Ch <= "9") Then
                    Digits = Digits + 1

            'At most one decimal point is allowed
                ElseIf Ch = "." Then
                    Points = Points + 1
                    If Points > 1 Then Exit Function

            'Anything else disqualifies the token
                Else
                    Exit Function
                End If

        Next I

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'At least one digit is required
        LooksNumeric = (Digits > 0)

End Function

Private Function LooksLocaleDate(ByVal TextIn As String) As Boolean

'
'------------------------------------------------------------------------------
'                          TEST FOR LOCALE DATE TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether text is shaped like a locale-formatted date.
'
' WHY THIS EXISTS
'   The contract reserves DATE_TEXT_LOCALE for text shaped like a
'   locale-formatted date. A year-first "2024/02/10", a time suffix such as
'   "2024-02-10 12.30" or "1.2.3" contains a separator but is not that shape,
'   so it is DATE_TEXT_FORMAT.
'
' INPUTS
'   TextIn
'     Text to test.
'
' RETURNS
'   Boolean
'     TRUE  => the text has the shape above
'     FALSE => otherwise
'
' BEHAVIOR
'   Accepted shape: three runs of ASCII digits separated by "/" or ".": one or
'   two digits, one or two digits, then two to four digits, as in
'   "31/12/2026", "12/31/2026" or "10.02.2024". Nothing may precede, follow or
'   separate them otherwise. The test is about shape only; the digits are
'   never read as a date.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-10
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Parse.LooksLocaleDate"

    Dim I               As Long         'Character cursor
    Dim Ch              As String       'Single character under test
    Dim Part            As Long         'Digit run under construction, 1 to 3
    Dim RunLen(1 To 3)  As Long         'Digits in each run

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default is not locale-shaped
        LooksLocaleDate = False
        Part = 1

'------------------------------------------------------------------------------
' SCAN
'------------------------------------------------------------------------------
    'Every character must be a digit or one of the two separators
        For I = 1 To VBA.Len(TextIn)

            'Read one character
                Ch = VBA.Mid$(TextIn, I, 1)

            'Digits lengthen the current run
                If (Ch >= "0") And (Ch <= "9") Then
                    RunLen(Part) = RunLen(Part) + 1

            'A separator closes the current run; a third one is too many
                ElseIf (Ch = "/") Or (Ch = ".") Then
                    If Part = 3 Then Exit Function
                    Part = Part + 1

            'Anything else disqualifies the text
                Else
                    Exit Function
                End If

        Next I

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Exactly three runs of the required lengths
        If Part <> 3 Then Exit Function
        If (RunLen(1) < 1) Or (RunLen(1) > 2) Then Exit Function
        If (RunLen(2) < 1) Or (RunLen(2) > 2) Then Exit Function
        If (RunLen(3) < 2) Or (RunLen(3) > 4) Then Exit Function
        LooksLocaleDate = True

End Function
