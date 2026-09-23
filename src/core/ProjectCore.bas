Attribute VB_Name = "ProjectCore"
'==============================================================================
' MODULE: ProjectCore
'------------------------------------------------------------------------------
' PURPOSE
'   Implement the neutral arithmetic example behind the supported facade.
'
' PUBLIC SURFACE
'   None outside this VBA project. DivideChecked and ERR_ZERO_DENOMINATOR are
'   Public only for in-project use; Option Private Module keeps them off the
'   supported external surface.
'
' DEPENDENCIES
'   VBA runtime only. This core never depends on the facade, tests, examples,
'   workbook objects, or optional references.
'
' STATE OWNERSHIP
'   Stateless. Results depend only on explicit scalar arguments.
'
' ERROR POLICY
'   Own the single internal zero-denominator error number and raise it with a
'   stable description. The facade exposes the same value and normalizes the
'   public error source.
'
' WORKSHEET SAFETY
'   Performs no Excel object-model access and changes no caller-owned state.
'
' TEST SEAM
'   ProjectTests verifies behavior through ProjectFacade. Direct core access is
'   available inside the project for focused future tests without widening API.
'
' COMPATIBILITY
'   Excel VBA; scalar VBA arithmetic requires no optional references.
'
' UPDATED
'   2026-09-09
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations; preserve the configured component visibility.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Own the internal error code used by the facade boundary.
        Public Const ERR_ZERO_DENOMINATOR   As Long = vbObjectError + 2048    'Core-owned zero-denominator code


'
'------------------------------------------------------------------------------
'
'                              CHECKED ARITHMETIC
'
'------------------------------------------------------------------------------
'

Public Function DivideChecked( _
    ByVal numerator As Double, _
    ByVal denominator As Double) _
    As Double
'
'==============================================================================
'                                DivideChecked
'------------------------------------------------------------------------------
' PURPOSE
'   Provide stateless checked division for in-project callers.
'
' INPUTS
'   numerator, denominator: explicit scalar Double values.
'
' RETURNS
'   Double quotient for a nonzero denominator.
'
' ERROR POLICY
'   Raise ERR_ZERO_DENOMINATOR for zero; arithmetic errors propagate.
'   The facade owns the supported external error source.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' VALIDATE DENOMINATOR
'------------------------------------------------------------------------------
    'Reject zero explicitly so callers receive the stable project error
    'instead of depending on the runtime division-by-zero diagnostic.
        If denominator = 0# Then
            Err.Raise _
                ERR_ZERO_DENOMINATOR, _
                "ProjectCore.DivideChecked", _
                "Denominator must not be zero."
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the quotient without reading or changing ambient host state.
        DivideChecked = numerator / denominator

End Function
