Attribute VB_Name = "CORE_Err"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: CORE_Err
'------------------------------------------------------------------------------
' PURPOSE
'   The one place that builds the native Excel errors the date layer returns,
'   names the conditions that cause them, and maps each condition to its error.
'
' WHY THIS EXISTS
'   Every error value the library originates is built here, so the set of
'   errors it can return is enumerable from one module instead of recovered by
'   searching for CVErr. A condition is an expected outcome of one element, not
'   a failure, so the parsers return it as a code and the facade maps it, once.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - KPRCondition   the condition vocabulary of the data contract
'     - ErrValue       the #VALUE! error
'     - ErrNum         the #NUM! error
'     - ErrNA          the #N/A error
'     - ErrForCondition  the native error a condition maps to
'     - ConditionName  the identifier of a condition, as text
'
' ERROR POLICY
'   The constructors and ConditionName never raise. ErrForCondition raises
'   ERR_UNMAPPED_CONDITION for cdNone, for the two propagation conditions and
'   for any unknown code: reaching it is an internal defect, never an expected
'   outcome.
'
' DEPENDENCIES
'   None: this is the bottom of the dependency graph. It uses the VBA function
'   CVErr and no Excel object model; the native error numbers are written as
'   constants. Data contract: docs/methodology/DATA_CONTRACT.md, Function
'   arguments (Error values, Conditions, Internal failures)
'
' NOTES
'   - The enum names mirror the identifiers of the data contract's registry,
'     and ConditionName returns the identifier itself. The enum numbers are
'     internal: they never appear in test cases or evidence, are never
'     renumbered and are never reused
'   - An incoming Excel error is detected at the boundary and returned
'     unchanged without passing through this module. ErrForCondition refuses
'     the propagation conditions so that a discarded incoming error cannot be
'     hidden behind a new one
'   - ErrNA serves the two host conditions only. It is the same Excel value as
'     an #N/A passed in; the condition identifier tells them apart
'   - Option Private Module keeps every member out of the Excel function list.
'     Nothing here is supported API
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'The native Excel error numbers that CVErr takes (the values of the Excel
    'library's XlCVError members, written here so that this module needs no
    'Excel reference)
    Private Const ERROR_VALUE_NUMBER    As Long = 2015  '#VALUE!
    Private Const ERROR_NUM_NUMBER      As Long = 2036  '#NUM!
    Private Const ERROR_NA_NUMBER       As Long = 2042  '#N/A

    'The description of the one error this module raises
    Private Const UNMAPPED_DESCRIPTION  As String = _
        "Condition has no worksheet error mapping; this is an internal defect."

'------------------------------------------------------------------------------
' PUBLIC CONSTANTS
'------------------------------------------------------------------------------
    'The one error this module raises, in the project's range (vbObjectError +
    '2048 through +4063); written as a literal because LibreOffice does not
    'compile vbObjectError in a constant expression
    Public Const ERR_UNMAPPED_CONDITION As Long = &H80040800    'vbObjectError + 2048

'------------------------------------------------------------------------------
' PUBLIC ENUMS
'------------------------------------------------------------------------------
    'The conditions of the data contract, Conditions table. cdNone is the
    'success sentinel, not a condition: success has no identifier
    Public Enum KPRCondition
        cdNone = 0                      'Success sentinel

        'Date value conditions
        cdDateTextFormat = 1            'DATE_TEXT_FORMAT
        cdDateTextLocale = 2            'DATE_TEXT_LOCALE
        cdDateTextNumeric = 3           'DATE_TEXT_NUMERIC
        cdDateTextImpossible = 4        'DATE_TEXT_IMPOSSIBLE
        cdDateTypeRejected = 5          'DATE_TYPE_REJECTED
        cdDateWindow = 6                'DATE_WINDOW

        'Shared input conditions
        cdInputBlankRequired = 10       'INPUT_BLANK_REQUIRED
        cdInputErrorPropagated = 11     'INPUT_ERROR_PROPAGATED

        'Integer value conditions
        cdIntegerFraction = 20          'INTEGER_FRACTION
        cdIntegerTypeRejected = 21      'INTEGER_TYPE_REJECTED
        cdIntegerRange = 22             'INTEGER_RANGE

        'Argument domain and result conditions
        cdDomainYear = 30               'DOMAIN_YEAR
        cdDomainMonth = 31              'DOMAIN_MONTH
        cdDomainWeekday = 32            'DOMAIN_WEEKDAY
        cdDomainOccurrence = 33         'DOMAIN_OCCURRENCE
        cdOccurrenceAbsent = 34         'OCCURRENCE_ABSENT
        cdResultWindow = 35             'RESULT_WINDOW

        'Optional control conditions
        cdControlTypeRejected = 40      'CONTROL_TYPE_REJECTED
        cdControlErrorPropagated = 41   'CONTROL_ERROR_PROPAGATED
        cdControlTokenUnknown = 42      'CONTROL_TOKEN_UNKNOWN
        cdControlNotScalar = 43         'CONTROL_NOT_SCALAR

        'Shape conditions
        cdShapeUnsupported = 50         'SHAPE_UNSUPPORTED
        cdShapeMismatch = 51            'SHAPE_MISMATCH
        cdCapacityExceeded = 52         'CAPACITY_EXCEEDED

        'Host date-system conditions
        cdHostDate1904 = 60             'HOST_DATE1904
        cdHostUnresolved = 61           'HOST_UNRESOLVED

        'Pillar token conditions
        cdPillarTypeRejected = 70       'PILLAR_TYPE_REJECTED
        cdPillarTokenMalformed = 71     'PILLAR_TOKEN_MALFORMED
        cdPillarDuplicateUnit = 72      'PILLAR_DUPLICATE_UNIT
        cdPillarAliasSigned = 73        'PILLAR_ALIAS_SIGNED
        cdPillarAggregateRange = 74     'PILLAR_AGGREGATE_RANGE
    End Enum


'
'------------------------------------------------------------------------------
'
'                          NATIVE ERROR CONSTRUCTORS
'
'------------------------------------------------------------------------------
'

Public Function ErrValue() As Variant

'
'------------------------------------------------------------------------------
'                            BUILD THE #VALUE! ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the native Excel #VALUE! error value.
'
' WHY THIS EXISTS
'   Keeps #VALUE! construction in this module, so every error the library
'   originates is built in one auditable place.
'
' INPUTS
'   None
'
' RETURNS
'   Variant
'     CVErr(2015), a Variant of subtype vbError.
'
' BEHAVIOR
'   Assigns CVErr(ERROR_VALUE_NUMBER) to the result.
'
' ERROR POLICY
'   None. Construction only.
'
' DEPENDENCIES
'   CVErr
'
' NOTES
'   - The result is assigned with Let, never Set. Test it with IsError: comparing
'     it to another value with = raises
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'The input cannot be interpreted under the contract
        ErrValue = VBA.CVErr(ERROR_VALUE_NUMBER)

End Function

Public Function ErrNum() As Variant

'
'------------------------------------------------------------------------------
'                             BUILD THE #NUM! ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the native Excel #NUM! error value.
'
' WHY THIS EXISTS
'   Keeps #NUM! construction in this module, so every error the library
'   originates is built in one auditable place.
'
' INPUTS
'   None
'
' RETURNS
'   Variant
'     CVErr(2036), a Variant of subtype vbError.
'
' BEHAVIOR
'   Assigns CVErr(ERROR_NUM_NUMBER) to the result.
'
' ERROR POLICY
'   None. Construction only.
'
' DEPENDENCIES
'   CVErr
'
' NOTES
'   - The result is assigned with Let, never Set. Test it with IsError
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'The input is well formed but its value or result is outside the domain
        ErrNum = VBA.CVErr(ERROR_NUM_NUMBER)

End Function

Public Function ErrNA() As Variant

'
'------------------------------------------------------------------------------
'                              BUILD THE #N/A ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the native Excel #N/A error value.
'
' WHY THIS EXISTS
'   Keeps #N/A construction in this module, so every error the library
'   originates is built in one auditable place. The library produces #N/A for
'   the two host date-system conditions only.
'
' INPUTS
'   None
'
' RETURNS
'   Variant
'     CVErr(2042), a Variant of subtype vbError.
'
' BEHAVIOR
'   Assigns CVErr(ERROR_NA_NUMBER) to the result.
'
' ERROR POLICY
'   None. Construction only.
'
' DEPENDENCIES
'   CVErr
'
' NOTES
'   - The result is assigned with Let, never Set. Test it with IsError
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'The result is unavailable in the calling worksheet's configuration
        ErrNA = VBA.CVErr(ERROR_NA_NUMBER)

End Function

'
'------------------------------------------------------------------------------
'
'                             CONDITION MAPPING
'
'------------------------------------------------------------------------------
'

Public Function ErrForCondition(ByVal Condition As KPRCondition) As Variant

'
'------------------------------------------------------------------------------
'                    MAP A CONDITION TO ITS NATIVE ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   Maps a condition to the native Excel error the data contract gives it.
'
' WHY THIS EXISTS
'   One mapper makes the condition-to-error policy auditable in one place
'   instead of at every failure site.
'
' INPUTS
'   Condition
'     A KPRCondition member produced by a parser or a domain check.
'
' RETURNS
'   Variant
'     #VALUE!, #NUM! or #N/A, as the Conditions table of the data contract
'     lists for it.
'
' BEHAVIOR
'   - Contract-invalid input conditions map to #VALUE!
'   - Well-formed input whose answer does not exist or leaves the domain maps
'     to #NUM!
'   - The two host date-system conditions map to #N/A
'   - Anything else raises
'
' ERROR POLICY
'   Raises ERR_UNMAPPED_CONDITION, with source CORE_Err.ErrForCondition and the
'   description "Condition has no worksheet error mapping; this is an internal
'   defect.", for cdNone, for cdInputErrorPropagated and
'   cdControlErrorPropagated, and for any unknown code.
'
'   These are deliberately not mapped to #VALUE!. A success sentinel that
'   reaches an error mapper means a caller ignored a True return; a
'   propagation condition means a caller discarded the incoming error it had
'   to return unchanged. Answering #VALUE! would hide both defects.
'
' DEPENDENCIES
'   ErrValue, ErrNum, ErrNA. Data contract: Function arguments, Conditions
'
' NOTES
'   - The two host conditions are the only #N/A the library produces. A
'     propagated incoming #N/A never passes through this mapper
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Err.ErrForCondition"

'------------------------------------------------------------------------------
' MAP CONDITION
'------------------------------------------------------------------------------
    'Choose the native error from the condition's class
        Select Case Condition

            'Contract-invalid input: the value cannot be interpreted
            Case cdDateTextFormat, _
                cdDateTextLocale, _
                cdDateTextNumeric, _
                cdDateTextImpossible, _
                cdDateTypeRejected, _
                cdInputBlankRequired, _
                cdIntegerFraction, _
                cdIntegerTypeRejected, _
                cdDomainYear, _
                cdDomainMonth, _
                cdDomainWeekday, _
                cdDomainOccurrence, _
                cdControlTypeRejected, _
                cdControlTokenUnknown, _
                cdControlNotScalar, _
                cdShapeUnsupported, _
                cdShapeMismatch, _
                cdPillarTypeRejected, _
                cdPillarTokenMalformed, _
                cdPillarDuplicateUnit, _
                cdPillarAliasSigned
                ErrForCondition = ErrValue()

            'Well formed, but the answer does not exist or leaves the domain
            Case cdDateWindow, _
                cdIntegerRange, _
                cdOccurrenceAbsent, _
                cdResultWindow, _
                cdPillarAggregateRange, _
                cdCapacityExceeded
                ErrForCondition = ErrNum()

            'Result unavailable in this host configuration
            Case cdHostDate1904, _
                cdHostUnresolved
                ErrForCondition = ErrNA()

            'Success sentinel, propagation conditions and unknown codes
            Case Else
                Err.Raise ERR_UNMAPPED_CONDITION, PROC_NAME, UNMAPPED_DESCRIPTION

        End Select

End Function

Public Function ConditionName(ByVal Condition As KPRCondition) As String

'
'------------------------------------------------------------------------------
'                      NAME A CONDITION BY ITS IDENTIFIER
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the identifier of a condition as text, such as DATE_WINDOW.
'
' WHY THIS EXISTS
'   Test cases and evidence cite the semantic identifier, never the enum
'   number, because the Excel error alone cannot say which condition produced
'   it.
'
' INPUTS
'   Condition
'     A KPRCondition member.
'
' RETURNS
'   String
'     The identifier of the data contract's Conditions table, in upper case,
'     or vbNullString for cdNone, which has none, and for an unknown code.
'
' BEHAVIOR
'   Looks the member up in a fixed list.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   None. Data contract: Function arguments, Conditions
'
' NOTES
'   None
'
' UPDATED
'   2026-11-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' LOOK UP IDENTIFIER
'------------------------------------------------------------------------------
    'Return the identifier that the member's name mirrors
        Select Case Condition
            Case cdDateTextFormat
                ConditionName = "DATE_TEXT_FORMAT"
            Case cdDateTextLocale
                ConditionName = "DATE_TEXT_LOCALE"
            Case cdDateTextNumeric
                ConditionName = "DATE_TEXT_NUMERIC"
            Case cdDateTextImpossible
                ConditionName = "DATE_TEXT_IMPOSSIBLE"
            Case cdDateTypeRejected
                ConditionName = "DATE_TYPE_REJECTED"
            Case cdDateWindow
                ConditionName = "DATE_WINDOW"
            Case cdInputBlankRequired
                ConditionName = "INPUT_BLANK_REQUIRED"
            Case cdInputErrorPropagated
                ConditionName = "INPUT_ERROR_PROPAGATED"
            Case cdIntegerFraction
                ConditionName = "INTEGER_FRACTION"
            Case cdIntegerTypeRejected
                ConditionName = "INTEGER_TYPE_REJECTED"
            Case cdIntegerRange
                ConditionName = "INTEGER_RANGE"
            Case cdDomainYear
                ConditionName = "DOMAIN_YEAR"
            Case cdDomainMonth
                ConditionName = "DOMAIN_MONTH"
            Case cdDomainWeekday
                ConditionName = "DOMAIN_WEEKDAY"
            Case cdDomainOccurrence
                ConditionName = "DOMAIN_OCCURRENCE"
            Case cdOccurrenceAbsent
                ConditionName = "OCCURRENCE_ABSENT"
            Case cdResultWindow
                ConditionName = "RESULT_WINDOW"
            Case cdControlTypeRejected
                ConditionName = "CONTROL_TYPE_REJECTED"
            Case cdControlErrorPropagated
                ConditionName = "CONTROL_ERROR_PROPAGATED"
            Case cdControlTokenUnknown
                ConditionName = "CONTROL_TOKEN_UNKNOWN"
            Case cdControlNotScalar
                ConditionName = "CONTROL_NOT_SCALAR"
            Case cdShapeUnsupported
                ConditionName = "SHAPE_UNSUPPORTED"
            Case cdShapeMismatch
                ConditionName = "SHAPE_MISMATCH"
            Case cdCapacityExceeded
                ConditionName = "CAPACITY_EXCEEDED"
            Case cdHostDate1904
                ConditionName = "HOST_DATE1904"
            Case cdHostUnresolved
                ConditionName = "HOST_UNRESOLVED"
            Case cdPillarTypeRejected
                ConditionName = "PILLAR_TYPE_REJECTED"
            Case cdPillarTokenMalformed
                ConditionName = "PILLAR_TOKEN_MALFORMED"
            Case cdPillarDuplicateUnit
                ConditionName = "PILLAR_DUPLICATE_UNIT"
            Case cdPillarAliasSigned
                ConditionName = "PILLAR_ALIAS_SIGNED"
            Case cdPillarAggregateRange
                ConditionName = "PILLAR_AGGREGATE_RANGE"
            Case Else
                ConditionName = vbNullString
        End Select

End Function
