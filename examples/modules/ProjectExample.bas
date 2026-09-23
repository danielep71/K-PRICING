Attribute VB_Name = "ProjectExample"
'==============================================================================
' MODULE: ProjectExample
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrate one supported facade call from explicit scalar inputs.
'
' PUBLIC SURFACE
'   RunProjectExample is an in-project example macro, not production API.
'
' DEPENDENCIES
'   ProjectFacade only.
'
' STATE OWNERSHIP
'   No mutable state. Output is written only to the VBE Immediate window.
'
' ERROR POLICY
'   Does not suppress facade errors; callers see the supported error contract.
'
' WORKSHEET SAFETY
'   Does not read or modify Application, workbook, worksheet, Range, selection,
'   calculation, events, display settings, or other host state.
'
' TEST SEAM
'   ProjectTests covers the same facade behavior with deterministic assertions.
'
' COMPATIBILITY
'   Excel VBA; scalar VBA arithmetic requires no optional references.
'
' USAGE
'   Import the required production modules first, then run
'   ProjectExample.RunProjectExample from the VBE Immediate window.
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


'
'------------------------------------------------------------------------------
'
'                             EXAMPLE ENTRY POINT
'
'------------------------------------------------------------------------------
'

Public Sub RunProjectExample()
'
'==============================================================================
'                              RunProjectExample
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrate a supported facade call without a workbook fixture.
'
' USAGE
'   Run ProjectExample.RunProjectExample from the VBE Immediate window.
'
' SIDE EFFECTS
'   Print one result to the Immediate window; no host state is changed.
'
' ERROR POLICY
'   Facade errors propagate to the caller.
'
' UPDATED
'   2026-09-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'Keep the demonstration reproducible with explicit scalar inputs.
    Const NUMERATOR     As Double = 12#    'Explicit sample dividend
    Const DENOMINATOR   As Double = 4#     'Nonzero sample divisor

'------------------------------------------------------------------------------
' REPORT EXAMPLE
'------------------------------------------------------------------------------
    'Call the supported facade and print its result without creating a
    'worksheet fixture or depending on the active selection.
        Debug.Print "ProjectRatio(12, 4) = " & _
            CStr(ProjectFacade.ProjectRatio(NUMERATOR, DENOMINATOR))

End Sub
