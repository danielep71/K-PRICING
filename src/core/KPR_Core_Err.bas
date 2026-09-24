Attribute VB_Name = "KPR_Core_Err"
'==============================================================================
' MODULE: KPR_Core_Err
'------------------------------------------------------------------------------
' PURPOSE
'   Internal boundary for native Excel error construction.
'
'   Every worksheet-facing failure value the KPR date layer CONSTRUCTS is built
'   here, so that the set of errors the library can originate is enumerable
'   from one place rather than recovered by grepping for CVErr.
'
' WHY THIS EXISTS
'   The pre-split baseline constructed CVErr(xlErrValue) and CVErr(xlErrNum)
'   inline at forty call sites. That made the error surface impossible to audit
'   and impossible to change coherently. Centralized construction plus one
'   condition-to-error mapper makes the complete error policy auditable here.
'
' SCOPE
'   - KPR_Condition   stable condition vocabulary from contract section 7
'   - ErrForCondition condition-to-native-error mapping
'   - ErrValue   #VALUE!
'   - ErrNum     #NUM!
'   - ErrNA      #N/A
'
' VISIBILITY
'   Option Private Module. Members are declared Public so other modules in the
'   project can call them; the module-level Private setting is what keeps them
'   out of the Excel function list and off the project's external surface.
'   Nothing here is supported API, and nothing here is visible to the
'   worksheet, the Function Wizard, the macro list or another VBA project.
'
' ALLOWED DEPENDENCIES
'   No KPR module. This is the bottom of the project dependency graph.
'
'   It does depend on the host: CVErr is a VBA conversion function and the
'   xlErr* names are members of XlCVError in the Excel object library. Without
'   that reference the module does not compile. That is the only external
'   surface it touches.
'
' NOTES
'   - ErrNA is used for HOST_DATE1904 and HOST_UNRESOLVED. It is constructed
'     here just like #VALUE! and #NUM!, so the library-originated error surface
'     remains centralized.
'   - Incoming errors are detected at the public/element boundary and returned
'     verbatim without passing through this module. ErrForCondition deliberately
'     refuses propagation identifiers so a discarded incoming value cannot be
'     hidden behind a newly constructed error.
'   - Each member is a function rather than a constant because VBA cannot hold
'     an error value in a Const. The cost is one call per constructed error.
'   - A caller that assigns a failure value before knowing whether it will fail
'     pays that call on the success path too. On a per-cell surface evaluated
'     across a large range this is not free; construct at the failure point
'     unless a handler needs to substitute a different value first.
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
' CONDITION VOCABULARY
'------------------------------------------------------------------------------
    'Stable condition identifiers from section 7 of the date-layer contract.
    '
    'The enum NAMES mirror the registry identifiers so that a condition is
    'compile-checked rather than spelled out in prose. The enum NUMBERS are an
    'internal implementation detail: they must never appear in fixtures or
    'evidence, which cite the semantic string form. Members are never renumbered
    'or reused, matching the registry rule.
    '
    'KPR_COND_NONE is a success sentinel, not a condition. The contract assigns
    'no identifier to successful evaluation.
        Public Enum KPR_Condition
            KPR_COND_NONE = 0

            'Date value conditions
            KPR_COND_DATE_TEXT_FORMAT = 1
            KPR_COND_DATE_TEXT_LOCALE = 2
            KPR_COND_DATE_TEXT_NUMERIC = 3
            KPR_COND_DATE_TEXT_IMPOSSIBLE = 4
            KPR_COND_DATE_TYPE_REJECTED = 5
            KPR_COND_DATE_WINDOW = 6

            'Shared input conditions
            KPR_COND_INPUT_BLANK_REQUIRED = 10
            KPR_COND_INPUT_ERROR_PROPAGATED = 11

            'Integer value conditions
            KPR_COND_INTEGER_FRACTION = 20
            KPR_COND_INTEGER_TYPE_REJECTED = 21
            KPR_COND_INTEGER_RANGE = 22

            'Argument domain conditions
            KPR_COND_DOMAIN_YEAR = 30
            KPR_COND_DOMAIN_MONTH = 31
            KPR_COND_DOMAIN_WEEKDAY = 32
            KPR_COND_DOMAIN_OCCURRENCE = 33
            KPR_COND_OCCURRENCE_ABSENT = 34
            KPR_COND_RESULT_WINDOW = 35

            'Optional control conditions
            KPR_COND_CONTROL_TYPE_REJECTED = 40
            KPR_COND_CONTROL_ERROR_PROPAGATED = 41
            KPR_COND_CONTROL_TOKEN_UNKNOWN = 42
            KPR_COND_CONTROL_NOT_SCALAR = 43

            'Shape conditions
            KPR_COND_SHAPE_UNSUPPORTED = 50
            KPR_COND_SHAPE_MISMATCH = 51
            KPR_COND_CAPACITY_EXCEEDED = 52

            'Host date-system conditions
            KPR_COND_HOST_DATE1904 = 60
            KPR_COND_HOST_UNRESOLVED = 61

            'Pillar token conditions
            KPR_COND_PILLAR_TYPE_REJECTED = 70
            KPR_COND_PILLAR_TOKEN_MALFORMED = 71
            KPR_COND_PILLAR_DUPLICATE_UNIT = 72
            KPR_COND_PILLAR_ALIAS_SIGNED = 73
            KPR_COND_PILLAR_AGGREGATE_RANGE = 74
        End Enum

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Raised when ErrForCondition is asked to map a code that has no worksheet
    'error. Reaching it is an internal defect, never an expected outcome.
        Private Const KPR_ERR_UNMAPPED_CONDITION As Long = vbObjectError + 512

'
'------------------------------------------------------------------------------
'
'                           NATIVE ERROR CONSTRUCTORS
'
'------------------------------------------------------------------------------
'

Public Function ErrValue() As Variant
'
'==============================================================================
'                                   ErrValue
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the native Excel #VALUE! error value.
'
' SIGNATURE
'   ErrValue() -> Variant
'
' RETURNS
'   Variant
'     CVErr(xlErrValue), a Variant of subtype vbError.
'
' ERROR POLICY
'   None. Construction only.
'
' NOTES
'   - The result must be Let-assigned, never Set. Testing it requires IsError;
'     comparing it to another value with = raises.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'The input cannot be interpreted under the module contract
        ErrValue = CVErr(xlErrValue)

End Function


Public Function ErrNum() As Variant
'
'==============================================================================
'                                    ErrNum
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the native Excel #NUM! error value.
'
' SIGNATURE
'   ErrNum() -> Variant
'
' RETURNS
'   Variant
'     CVErr(xlErrNum), a Variant of subtype vbError.
'
' ERROR POLICY
'   None. Construction only.
'
' NOTES
'   - The strict parser classifies an out-of-window date as DATE_WINDOW and the
'     facade maps it here to #NUM!.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'The input is well formed but the answer does not exist or is out of range
        ErrNum = CVErr(xlErrNum)

End Function


Public Function ErrNA() As Variant
'
'==============================================================================
'                                     ErrNA
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the native Excel #N/A error value.
'
' SIGNATURE
'   ErrNA() -> Variant
'
' RETURNS
'   Variant
'     CVErr(xlErrNA), a Variant of subtype vbError.
'
' ERROR POLICY
'   None. Construction only.
'
' NOTES
'   - Reached only through ErrForCondition, for HOST_DATE1904 and
'     HOST_UNRESOLVED.
'   - #N/A here means the library declines to answer, not that a lookup failed.
'     The distinction matters because a sheet cannot tell the two apart, so
'     the provenance of this value belongs in the contract rather than in a
'     cell comment.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'The library refuses to answer for this caller or host configuration
        ErrNA = CVErr(xlErrNA)

End Function

Public Function ErrForCondition( _
    ByVal Condition As KPR_Condition) _
    As Variant
'
'==============================================================================
'                               ErrForCondition
'------------------------------------------------------------------------------
' PURPOSE
'   Maps a classified condition to the native Excel error the contract requires
'   for it.
'
' INPUTS
'   Condition
'     A KPR_Condition member produced by a parser or a domain check.
'
' RETURNS
'   Variant
'     CVErr(xlErrValue), CVErr(xlErrNum) or CVErr(xlErrNA), per section 7 of
'     the contract.
'
' ERROR POLICY
'   Raises KPR_ERR_UNMAPPED_CONDITION for KPR_COND_NONE, for the propagation
'   conditions, and for any unknown code.
'
'   These are deliberately NOT mapped to #VALUE!. A success sentinel reaching an
'   error mapper means a caller ignored a TRUE return. A propagation condition
'   reaching it means a caller discarded the incoming error value it was
'   supposed to return verbatim. Silently answering #VALUE! would hide both.
'   Reaching this path is an internal defect and a regression failure; the
'   facade's defensive handler contains it, but containment is not an expected
'   outcome.
'
' NOTES
'   - The two host conditions are the only #N/A the library produces. A
'     propagated incoming #N/A is the same Excel value, but it never passes
'     through this mapper: it is returned verbatim at the boundary.
'
' UPDATED
'   2026-09-02
'==============================================================================
'

'------------------------------------------------------------------------------
' MAP CONDITION
'------------------------------------------------------------------------------
    Select Case Condition

        'Contract-invalid input: the value cannot be interpreted
            Case KPR_COND_DATE_TEXT_FORMAT, _
                 KPR_COND_DATE_TEXT_LOCALE, _
                 KPR_COND_DATE_TEXT_NUMERIC, _
                 KPR_COND_DATE_TEXT_IMPOSSIBLE, _
                 KPR_COND_DATE_TYPE_REJECTED, _
                 KPR_COND_INPUT_BLANK_REQUIRED, _
                 KPR_COND_INTEGER_FRACTION, _
                 KPR_COND_INTEGER_TYPE_REJECTED, _
                 KPR_COND_DOMAIN_YEAR, _
                 KPR_COND_DOMAIN_MONTH, _
                 KPR_COND_DOMAIN_WEEKDAY, _
                 KPR_COND_DOMAIN_OCCURRENCE, _
                 KPR_COND_CONTROL_TYPE_REJECTED, _
                 KPR_COND_CONTROL_TOKEN_UNKNOWN, _
                 KPR_COND_CONTROL_NOT_SCALAR, _
                 KPR_COND_SHAPE_UNSUPPORTED, _
                 KPR_COND_SHAPE_MISMATCH, _
                 KPR_COND_PILLAR_TYPE_REJECTED, _
                 KPR_COND_PILLAR_TOKEN_MALFORMED, _
                 KPR_COND_PILLAR_DUPLICATE_UNIT, _
                 KPR_COND_PILLAR_ALIAS_SIGNED

                ErrForCondition = ErrValue()

        'Well formed, but the answer does not exist or leaves the domain
            Case KPR_COND_DATE_WINDOW, _
                 KPR_COND_INTEGER_RANGE, _
                 KPR_COND_OCCURRENCE_ABSENT, _
                 KPR_COND_RESULT_WINDOW, _
                 KPR_COND_PILLAR_AGGREGATE_RANGE, _
                 KPR_COND_CAPACITY_EXCEEDED

                ErrForCondition = ErrNum()

        'Result unavailable in this host configuration
            Case KPR_COND_HOST_DATE1904, _
                 KPR_COND_HOST_UNRESOLVED

                ErrForCondition = ErrNA()

        'Success sentinel, propagation conditions and unknown codes
            Case Else

                Err.Raise KPR_ERR_UNMAPPED_CONDITION, "KPR_Core_Err.ErrForCondition", _
                          "Condition has no worksheet error mapping; this is an internal defect."

    End Select

End Function
