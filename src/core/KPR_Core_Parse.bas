Attribute VB_Name = "KPR_Core_Parse"
'==============================================================================
' MODULE: KPR_Core_Parse
'------------------------------------------------------------------------------
' PURPOSE
'   Internal boundary for value classification: converting one already-unwrapped
'   scalar into a normalized VBA Date, or failing.
'
' WHY THIS EXISTS
'   Worksheet arguments carry many runtime types for the same intent. This
'   module is the single place where the acceptance policy for a date value
'   lives, so callers never re-implement fragile type logic and no unhandled
'   runtime error can reach the worksheet.
'
' SCOPE
'   - TryParseDateScalar   scalar -> normalized, in-window VBA Date
'   - TryParseLongScalar   scalar -> exact VBA Long
'   - TryParseBoolControl  scalar optional control -> Boolean/default
'   - TryParseRoundingControl  Opt_Rounding text -> normalized token
'
' SCOPE BOUNDARY
'   - Shape is NOT handled here. Callers pass a scalar already unwrapped by
'     KPR_Core_Array.TryUnwrapScalar.
'   - The supported date window is applied here after date-only normalization.
'     The serial/year bounds are restated because this module may not depend on
'     KPR_Core_Dates; the static gate pins both representations against each
'     other so they cannot drift.
'   - Incoming Excel errors are not parser failures. The public/element boundary
'     returns them verbatim before calling these routines.
'
' TYPE HANDOFF
'   A single-cell Range arrives from KPR_Core_Array read through Value2, so a
'   date-formatted cell presents as vbDouble rather than vbDate. Both types
'   resolve through the same arithmetic path below and produce the same
'   date-only result, so the surface does not depend on which one arrives.
'
' VISIBILITY
'   Option Private Module. Members are declared Public so other modules in the
'   project can call them; the module-level Private setting is what keeps them
'   out of the Excel function list and off the project's external surface.
'   Nothing here is supported API.
'
' ALLOWED DEPENDENCIES
'   KPR_Core_Err, for the shared KPR_Condition vocabulary. Every parser failure
'   is still reported through its Boolean return and condition ByRef output.
'
' UPDATED
'   2026-09-01
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
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Supported window expressed as serials and years.
    '
    'KPR_Core_Dates owns the window as Date constants. This module may not
    'depend on it, so the same bounds are restated here in the form this parser
    'needs. The static gate pins the two representations against each other, so
    'they cannot drift apart silently.
        Private Const KPR_MIN_SERIAL    As Double = 61#         '1900-03-01
        Private Const KPR_MAX_SERIAL    As Double = 2958465#    '9999-12-31
        Private Const KPR_MIN_YEAR      As Long = 1900
        Private Const KPR_MAX_YEAR      As Long = 9999

    'Long range, tested before any narrowing conversion
        Private Const KPR_LONG_MIN      As Double = -2147483648#
        Private Const KPR_LONG_MAX      As Double = 2147483647#

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
    ByRef Condition As KPR_Condition) _
    As Boolean
'
'==============================================================================
'                              TryParseDateScalar
'------------------------------------------------------------------------------
' PURPOSE
'   Converts one already-unwrapped scalar into a normalized, in-window VBA Date
'   under the strict rules of contract sections 3.1 and 4, reporting the exact
'   condition on failure.
'
' SIGNATURE
'   TryParseDateScalar(ScalarIn, ParsedDate, Condition) -> Boolean
'
' INPUTS
'   ScalarIn
'     A scalar Variant already isolated by KPR_Core_Array.
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
' OUTPUTS
'   ParsedDate (ByRef)
'     Assigned ONLY on success, date-only and inside the supported window.
'
'   Condition (ByRef)
'     Always assigned. KPR_COND_NONE on success, otherwise the failing
'     condition.
'
' RETURNS
'   Boolean
'     TRUE  => ParsedDate assigned and Condition is KPR_COND_NONE
'     FALSE => ParsedDate untouched and Condition names the failure
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
'   - The caller maps the condition through KPR_Core_Err.ErrForCondition.
'
' NOTES
'   - Component validity is checked by round trip rather than by a private
'     calendar. DateSerial is called only once the year is known to be inside
'     the window, and the result must return the same year, month and day.
'     2025-02-29 rolls to 2025-03-01 and is rejected, so rollover can never pass
'     silently, and no leap-year logic is duplicated out of KPR_Core_Dates.
'   - The year gate precedes construction for a second reason: DateSerial
'     reinterprets a year below 100 as a 20th-century year, so DateSerial(50,1,1)
'     is 1950. A syntactically exact 0050-01-01 is reported as DATE_WINDOW
'     without ever being constructed.
'   - The window bounds are inlined as serials rather than read from
'     KPR_Core_Dates, which this module may not call. They are pinned by the
'     static gate against the constants that own them.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim VT              As VbVarType    'Cached VarType of the incoming scalar
    Dim Candidate       As Date         'Parsed candidate before assignment
    Dim Serial          As Double       'Numeric working value for serial inputs
    Dim S               As String       'Text payload, used without trimming

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default outcome is failure with an explicit condition
        TryParseDateScalar = False
        Condition = KPR_COND_DATE_TYPE_REJECTED

'------------------------------------------------------------------------------
' CLASSIFY SCALAR
'------------------------------------------------------------------------------
    'Classify the already-unwrapped payload
        VT = VarType(ScalarIn)

    Select Case VT

        Case vbEmpty
            'A blank required cell is contract-invalid, not a zero serial
                Condition = KPR_COND_INPUT_BLANK_REQUIRED
                Exit Function

        Case vbNull, vbBoolean, vbObject, vbDataObject, vbError
            'Null, Boolean and objects are never dates. An incoming error should
            'have been intercepted at the boundary; classify rather than guess.
                Condition = KPR_COND_DATE_TYPE_REJECTED
                Exit Function

        Case vbDate
            'Native Date => remove the time component from the serial
                Serial = KPR_FloorSerial(CDbl(ScalarIn))

        Case vbString
            'Text is accepted in exactly one form, checked without locale help
                S = CStr(ScalarIn)

            'Classify the text before attempting any conversion
                If Not TryClassifyDateText(S, Serial, Condition) Then Exit Function

        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            'Native numeric serial => remove the time component
                Serial = KPR_FloorSerial(CDbl(ScalarIn))

        Case Else
            'LongLong is a native numeric subtype on VBA7. The test lives inside
            'the Case body rather than as its own Case, so the conditional
            'directive never sits between two Case clauses, and it is gated on
            'VBA7 rather than Win64 because the vbLongLong constant exists in
            'the VBA7 enum regardless of Office bitness.
#If VBA7 Then
                If VT = vbLongLong Then
                    Serial = KPR_FloorSerial(CDbl(ScalarIn))
                Else
                    Condition = KPR_COND_DATE_TYPE_REJECTED
                    Exit Function
                End If
#Else
                Condition = KPR_COND_DATE_TYPE_REJECTED
                Exit Function
#End If

    End Select

'------------------------------------------------------------------------------
' APPLY SUPPORTED WINDOW
'------------------------------------------------------------------------------
    'One window test for every accepted path, on the date-only value
        If (Serial < KPR_MIN_SERIAL) Or (Serial > KPR_MAX_SERIAL) Then
            Condition = KPR_COND_DATE_WINDOW
            Exit Function
        End If

    'Safe: the serial is inside the window, so conversion cannot overflow
        Candidate = CDate(Serial)

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once the value is fully validated
        ParsedDate = Candidate
        Condition = KPR_COND_NONE

    'Contract: TRUE only when ParsedDate was assigned
        TryParseDateScalar = True

End Function

Private Function TryClassifyDateText( _
    ByVal TextIn As String, _
    ByRef SerialOut As Double, _
    ByRef Condition As KPR_Condition) _
    As Boolean
'
'==============================================================================
'                             TryClassifyDateText
'------------------------------------------------------------------------------
' PURPOSE
'   Accepts text only in the exact form YYYY-MM-DD and returns its date serial.
'
' INPUTS
'   TextIn
'     The raw text payload, untrimmed. Surrounding whitespace is a rejection,
'     not something to clean up.
'
' OUTPUTS
'   SerialOut (ByRef)
'     Assigned ONLY on success, to the date serial of the parsed value.
'
'   Condition (ByRef)
'     Always assigned on failure.
'
' RETURNS
'   Boolean
'     TRUE  => SerialOut assigned
'     FALSE => Condition names the rejection
'
' BEHAVIOR
'   - Length must be exactly ten characters.
'   - Positions 5 and 8 must be hyphens; the other eight must be ASCII digits.
'   - Text that is numeric as a whole is reported as DATE_TEXT_NUMERIC, and text
'     that merely contains a separator the host would accept is reported as
'     DATE_TEXT_LOCALE, so provenance survives even though both return #VALUE!.
'   - The year is gated against the supported window before construction; the
'     month, day and leap day are validated by DateSerial round trip.
'
' ERROR POLICY
'   - Does not raise. Digit validation precedes every numeric conversion, so no
'     conversion in this routine can fail.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
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
        Condition = KPR_COND_DATE_TEXT_FORMAT

'------------------------------------------------------------------------------
' REJECT NON-ISO SHAPES
'------------------------------------------------------------------------------
    'Empty text carries no date
        If Len(TextIn) = 0 Then Exit Function

    'A numeric-looking string is never reinterpreted as a serial.
    'The test is our own, not the host's: VBA IsNumeric consults the locale and
    'reads "31.12.2026" as thirty-one million under an Italian decimal
    'convention, which would misclassify a locale-formatted date as numeric.
        If LooksNumeric(TextIn) Then
            Condition = KPR_COND_DATE_TEXT_NUMERIC
            Exit Function
        End If

    'Distinguish a locale-shaped date from generic malformed text for provenance.
    'Whitespace is NOT a locale marker: the contract classifies padded text as
    'DATE_TEXT_FORMAT, alongside every other malformed ISO shape.
        If (InStr(1, TextIn, "/", vbBinaryCompare) > 0) Or _
           (InStr(1, TextIn, ".", vbBinaryCompare) > 0) Then
            Condition = KPR_COND_DATE_TEXT_LOCALE
            Exit Function
        End If

    'The accepted form is exactly ten characters
        If Len(TextIn) <> 10 Then Exit Function

    'Separators sit at fixed positions
        If Mid$(TextIn, 5, 1) <> "-" Then Exit Function
        If Mid$(TextIn, 8, 1) <> "-" Then Exit Function

    'Every other position must be an ASCII digit
        For I = 1 To 10

            'Skip the two separator positions already verified
                If (I <> 5) And (I <> 8) Then

                    'Compare by code point so no locale digit is accepted
                        Ch = Mid$(TextIn, I, 1)
                        If (Ch < "0") Or (Ch > "9") Then Exit Function

                End If

        Next I

'------------------------------------------------------------------------------
' VALIDATE COMPONENTS
'------------------------------------------------------------------------------
    'Safe: every character position has been verified as a digit
        YearPart = CLng(Left$(TextIn, 4))
        MonthPart = CLng(Mid$(TextIn, 6, 2))
        DayPart = CLng(Right$(TextIn, 2))

    'Gate the year before construction: DateSerial reinterprets years below 100
        If (YearPart < KPR_MIN_YEAR) Or (YearPart > KPR_MAX_YEAR) Then
            Condition = KPR_COND_DATE_WINDOW
            Exit Function
        End If

    'Reject an impossible month before construction
        If (MonthPart < 1) Or (MonthPart > 12) Then
            Condition = KPR_COND_DATE_TEXT_IMPOSSIBLE
            Exit Function
        End If

    'Reject an impossible day the same way
        If (DayPart < 1) Or (DayPart > 31) Then
            Condition = KPR_COND_DATE_TEXT_IMPOSSIBLE
            Exit Function
        End If

    'Construct, then verify componentwise so rollover cannot pass silently
        RoundTrip = DateSerial(YearPart, MonthPart, DayPart)

    'A rolled-over date differs in at least one component
        If (Year(RoundTrip) <> YearPart) Or _
           (Month(RoundTrip) <> MonthPart) Or _
           (Day(RoundTrip) <> DayPart) Then
            Condition = KPR_COND_DATE_TEXT_IMPOSSIBLE
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Report the serial; the caller applies the supported window
        SerialOut = CDbl(RoundTrip)
        Condition = KPR_COND_NONE

    'Contract: TRUE only when SerialOut was assigned
        TryClassifyDateText = True

End Function

'
'------------------------------------------------------------------------------
'
'                          INTEGER AND CONTROL PARSING
'
'------------------------------------------------------------------------------
'

Public Function TryParseLongScalar( _
    ByVal ScalarIn As Variant, _
    ByRef ParsedLong As Long, _
    ByRef Condition As KPR_Condition) _
    As Boolean
'
'==============================================================================
'                              TryParseLongScalar
'------------------------------------------------------------------------------
' PURPOSE
'   Converts one already-unwrapped scalar into a Long under the strict rules of
'   contract section 3.2, reporting the exact condition on failure.
'
' INPUTS
'   ScalarIn
'     Accepted: native numeric that is mathematically integral and inside the
'     VBA Long range.
'
'     Rejected:
'       - fractional numeric              INTEGER_FRACTION
'       - numeric outside the Long range  INTEGER_RANGE
'       - Boolean, Date, ANY text including numeric-looking text, Null, object
'                                         INTEGER_TYPE_REJECTED
'       - Empty / blank required cell     INPUT_BLANK_REQUIRED
'
' OUTPUTS
'   ParsedLong (ByRef)   assigned ONLY on success
'   Condition  (ByRef)   always assigned
'
' RETURNS
'   Boolean
'     TRUE  => ParsedLong assigned and Condition is KPR_COND_NONE
'     FALSE => ParsedLong untouched and Condition names the failure
'
' BEHAVIOR
'   - Range is tested BEFORE integrality, per contract section 3.2: a value
'     outside the Long range returns INTEGER_RANGE and #NUM! even when it also
'     has a fractional part.
'   - CLng is reached only after both tests pass, so its round-half-to-even
'     behaviour can never be observed.
'
' ERROR POLICY
'   - Does not raise. Every failure path returns FALSE with a condition.
'
' NOTES
'   - Boolean is rejected rather than coerced. TRUE would otherwise become -1,
'     which is a silently plausible shift count.
'   - Text is rejected even when it looks like a number, mirroring the date
'     rule that numeric-looking text is never reinterpreted.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim VT              As VbVarType    'Cached VarType of the incoming scalar
    Dim X               As Double       'Numeric working value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default outcome is failure with an explicit condition
        TryParseLongScalar = False
        Condition = KPR_COND_INTEGER_TYPE_REJECTED

'------------------------------------------------------------------------------
' CLASSIFY SCALAR
'------------------------------------------------------------------------------
    'Classify the already-unwrapped payload
        VT = VarType(ScalarIn)

    Select Case VT

        Case vbEmpty
            'A blank required cell is contract-invalid, not a zero
                Condition = KPR_COND_INPUT_BLANK_REQUIRED
                Exit Function

        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            'Accepted numeric types; coerce once to a common working type
                X = CDbl(ScalarIn)

        Case Else
            'LongLong is accepted on VBA7 and the Long-range gate below still
            'owns whether its value can be taken. Everything else (Boolean,
            'Date, text, Null, objects, errors) is rejected. See
            'TryParseDateScalar for why the directive sits inside the Case body
            'and is gated on VBA7.
#If VBA7 Then
                If VT = vbLongLong Then
                    X = CDbl(ScalarIn)
                Else
                    Condition = KPR_COND_INTEGER_TYPE_REJECTED
                    Exit Function
                End If
#Else
                Condition = KPR_COND_INTEGER_TYPE_REJECTED
                Exit Function
#End If

    End Select

'------------------------------------------------------------------------------
' VALIDATE RANGE THEN INTEGRALITY
'------------------------------------------------------------------------------
    'Range precedes integrality when both apply (contract section 3.2)
        If (X < KPR_LONG_MIN) Or (X > KPR_LONG_MAX) Then
            Condition = KPR_COND_INTEGER_RANGE
            Exit Function
        End If

    'No silent truncation and no rounding: a fraction is a rejection
        If X <> Int(X) Then
            Condition = KPR_COND_INTEGER_FRACTION
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Safe: the value is integral and inside the Long range
        ParsedLong = CLng(X)
        Condition = KPR_COND_NONE

    'Contract: TRUE only when ParsedLong was assigned
        TryParseLongScalar = True

End Function

Public Function TryParseBoolControl( _
    ByVal ControlIn As Variant, _
    ByVal DefaultValue As Boolean, _
    ByRef ParsedBool As Boolean, _
    ByRef Condition As KPR_Condition) _
    As Boolean
'
'==============================================================================
'                             TryParseBoolControl
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves a present Boolean optional control under contract section 3.3.
'
' INPUTS
'   ControlIn
'     Accepted: Empty, meaning the documented default, or a native Boolean.
'
'     Rejected with CONTROL_TYPE_REJECTED: numeric 0 and 1, Boolean-looking text
'     such as "TRUE", Null, dates, objects.
'
'   DefaultValue
'     The documented default for this control, selected when ControlIn is Empty.
'
' OUTPUTS
'   ParsedBool (ByRef)   assigned ONLY on success
'   Condition  (ByRef)   always assigned
'
' RETURNS
'   Boolean
'     TRUE  => ParsedBool assigned and Condition is KPR_COND_NONE
'     FALSE => ParsedBool untouched and Condition names the failure
'
' BEHAVIOR
'   - An omitted argument is handled by the caller, which cannot pass a missing
'     Variant onward without losing that state. Omitted and Empty select the
'     same documented default.
'   - An incoming Excel error is NOT classified here. A control carrying an
'     error propagates verbatim as a call-level result, which is a decision for
'     the boundary rather than the parser.
'
' ERROR POLICY
'   - Does not raise. Every failure path returns FALSE with a condition.
'
' NOTES
'   - Accepting a native Boolean only is what makes the control strict. The
'     previous As Boolean declaration accepted 0, 1 and "TRUE" through Excel
'     coercion before the function was ever entered.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim VT              As VbVarType    'Cached VarType of the incoming control

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default outcome is failure with an explicit condition
        TryParseBoolControl = False
        Condition = KPR_COND_CONTROL_TYPE_REJECTED

'------------------------------------------------------------------------------
' CLASSIFY CONTROL
'------------------------------------------------------------------------------
    'Classify the payload
        VT = VarType(ControlIn)

    Select Case VT

        Case vbEmpty
            'A blank control selects the documented default
                ParsedBool = DefaultValue

        Case vbBoolean
            'The only accepted explicit form
                ParsedBool = CBool(ControlIn)

        Case Else
            'Numeric 0 / 1, Boolean-looking text, Null, dates and objects
                Condition = KPR_COND_CONTROL_TYPE_REJECTED
                Exit Function

    End Select

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Contract: TRUE only when ParsedBool was assigned
        Condition = KPR_COND_NONE
        TryParseBoolControl = True

End Function

Public Function TryParseRoundingControl( _
    ByVal ControlIn As Variant, _
    ByRef TokenOut As String, _
    ByRef Condition As KPR_Condition) _
    As Boolean
'
'==============================================================================
'                           TryParseRoundingControl
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves a present Opt_Rounding control to one of the three normalized
'   tokens NEAREST, FLOOR or CEILING, under contract section 3.3.
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
' OUTPUTS
'   TokenOut  (ByRef)   assigned ONLY on success, always upper-case ASCII
'   Condition (ByRef)   always assigned
'
' RETURNS
'   Boolean
'     TRUE  => TokenOut assigned and Condition is KPR_COND_NONE
'     FALSE => TokenOut untouched and Condition names the failure
'
' BEHAVIOR
'   - Case folding is ASCII-only: a-z map to A-Z and nothing else changes.
'     UCase$ is deliberately not used, because it consults the host locale and
'     under Turkish casing the dotless i in CEILING would not fold to I.
'   - Trimming removes leading and trailing ASCII spaces and tabs only.
'   - An incoming Excel error is NOT classified here; the boundary propagates
'     it before this routine is called.
'   - The result is text rather than an enum because this module may not
'     depend on KPR_Core_Dates, which owns the rounding vocabulary. The facade
'     performs the final mapping.
'
' ERROR POLICY
'   - Does not raise. Every failure path returns FALSE with a condition.
'
' UPDATED
'   2026-09-01
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Folded          As String   'Trimmed, ASCII-folded candidate token

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default outcome is failure with an explicit condition
        TryParseRoundingControl = False
        Condition = KPR_COND_CONTROL_TYPE_REJECTED

'------------------------------------------------------------------------------
' CLASSIFY CONTROL
'------------------------------------------------------------------------------
    Select Case VarType(ControlIn)

        Case vbEmpty
            'A blank control selects the documented default
                Folded = "NEAREST"

        Case vbString
            'Normalize without consulting the locale
                Folded = AsciiUpper(AsciiTrim(CStr(ControlIn)))

        Case Else
            'Numbers, Booleans, dates, Null and objects are not tokens
                Condition = KPR_COND_CONTROL_TYPE_REJECTED
                Exit Function

    End Select

'------------------------------------------------------------------------------
' MATCH TOKEN
'------------------------------------------------------------------------------
    'Exactly three tokens are recognized
        Select Case Folded
            Case "NEAREST", "FLOOR", "CEILING"
                TokenOut = Folded
            Case Else
                Condition = KPR_COND_CONTROL_TOKEN_UNKNOWN
                Exit Function
        End Select

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Contract: TRUE only when TokenOut was assigned
        Condition = KPR_COND_NONE
        TryParseRoundingControl = True

End Function

Private Function AsciiTrim( _
    ByVal TextIn As String) _
    As String
'
' Removes leading and trailing ASCII spaces and tabs. Nothing else counts as
' whitespace here, so the result is identical in every locale.
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
    If Last < First Then AsciiTrim = vbNullString Else AsciiTrim = Mid$(TextIn, First, Last - First + 1)

End Function

Private Function AsciiUpper( _
    ByVal TextIn As String) _
    As String
'
' Folds a-z to A-Z by code point and leaves every other character untouched.
' Locale-independent by construction; see TryParseRoundingControl.
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
    AsciiUpper = Result

End Function

Private Function KPR_FloorSerial( _
    ByVal SerialIn As Double) _
    As Double
'
'==============================================================================
'                                KPR_FloorSerial
'------------------------------------------------------------------------------
' PURPOSE
'   Removes the time component from a date serial.
'
' RETURNS
'   Double
'     The largest whole serial not greater than SerialIn.
'
' NOTES
'   - Int floors; Fix truncates toward zero. They agree for every serial inside
'     the supported window, which is entirely positive, but Int is used so that
'     a negative serial cannot move toward the window before being gated.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Floor rather than truncate toward zero
        KPR_FloorSerial = Int(SerialIn)

End Function

Private Function LooksNumeric( _
    ByVal TextIn As String) _
    As Boolean
'
'==============================================================================
'                                 LooksNumeric
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether text is numeric under a fixed, locale-independent rule.
'
' ACCEPTED SHAPE
'   An optional leading + or -, ASCII digits, and at most one "." acting as a
'   decimal point. At least one digit must be present.
'
' WHY NOT IsNumeric
'   VBA IsNumeric asks the host how to read the text. Under a locale that uses
'   "." as a group separator it accepts "31.12.2026" as a number, which would
'   classify a locale-formatted date as DATE_TEXT_NUMERIC and lose the
'   provenance the contract requires. It also accepts currency symbols,
'   thousands separators and surrounding whitespace, none of which this library
'   treats as numeric. The rule here is fixed and identical in every region.
'
' RETURNS
'   Boolean
'     TRUE  => the text is numeric under the rule above
'     FALSE => otherwise
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim I           As Long     'Character cursor
    Dim Ch          As String   'Single character under test
    Dim Digits      As Long     'Count of digits seen
    Dim Points      As Long     'Count of decimal points seen
    Dim StartAt     As Long     'First position after any sign

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default is not numeric
        LooksNumeric = False

    'Empty text is not numeric
        If Len(TextIn) = 0 Then Exit Function

    'Skip a single leading sign
        StartAt = 1
        Ch = Left$(TextIn, 1)
        If (Ch = "+") Or (Ch = "-") Then StartAt = 2

    'A sign on its own is not numeric
        If StartAt > Len(TextIn) Then Exit Function

'------------------------------------------------------------------------------
' SCAN
'------------------------------------------------------------------------------
    'Every remaining character must be a digit or the single decimal point
        For I = StartAt To Len(TextIn)

            'Read one character
                Ch = Mid$(TextIn, I, 1)

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
