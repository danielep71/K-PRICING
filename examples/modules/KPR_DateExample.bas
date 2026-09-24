Attribute VB_Name = "KPR_DateExample"
'==============================================================================
' MODULE: KPR_DateExample
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrate one supported KPR date-layer call from direct VBA.
'
' PUBLIC SURFACE
'   RunDateExample is an in-project example macro, not supported production API.
'
' DEPENDENCIES
'   KPR_DATES_DAYS only.
'
' STATE OWNERSHIP
'   No mutable Excel state. Output is written only to the VBE Immediate window.
'
' HOST CONTRACT
'   Direct VBA is a documented non-Range caller and therefore uses the documented
'   1900 serial contract. Worksheet callers remain subject to the caller
'   workbook's date-system rules in docs/DATE_LAYER_CONTRACT.md.
'==============================================================================

Option Explicit
Option Private Module

Public Sub RunDateExample()
    Dim Result As Variant

    Result = KPR_Dates_DaysInMonth("2026-02-15")
    Debug.Print "KPR_Dates_DaysInMonth(2026-02-15) = " & CStr(Result)
End Sub
