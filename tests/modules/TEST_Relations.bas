Attribute VB_Name = "TEST_Relations"

Option Explicit
Option Private Module

'
'==============================================================================
' MODULE: TEST_Relations
'==============================================================================
' PURPOSE
'   Illustrative relation tests: checks that need no external expected
'   value, run on a stand-in for the calculation under test, the total of a
'   list of amounts. Nothing here is a calculation of the workbook or a
'   regulatory value
'
' WHY THIS EXISTS
'   A relation (zero input, additivity, symmetry, monotonicity) can be
'   tested before any expected value exists. Each case asserts that its
'   relation holds, then that the same comparison rejects a deliberately
'   corrupted output, so that a relation which would pass any output is seen
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   RunRelationTests runs the suite. Option Private Module keeps it out of
'   the workbook's external surface. The cases build their inputs in memory
'   and change no Excel state. The suite is not certified: it is not listed
'   in .github/excel-evidence-policy.json
'
' ERROR POLICY
'   As TEST_Harness.EndRun: a run that does not pass raises
'   TEST_Harness.ERR_TEST_FAILED, or the unexpected error that escaped the
'   runner. A refused start raises TEST_Harness.ERR_TEST_DIRTY_START
'
' DEPENDENCIES
'   TEST_Harness
'
' NOTES
'   In a development workbook holding the modules of tests/modules/, type
'   TEST_Relations.RunRelationTests in the Immediate window. Replace
'   IllustrativeTotal with a call to the core procedure the relations test
'   (docs/methodology/TEST_HARNESS.md, Relation tests)
'
' UPDATED
'   2026-11-03
'==============================================================================

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'A complete run starts every case and evaluates every assertion once
    Private Const SUITE_NAME            As String = "relations"     'Suite name in the run log
    Private Const EXPECTED_CASES        As Long = 4                 'Cases a complete run starts
    Private Const EXPECTED_ASSERTIONS   As Long = 8                 'Assertions a complete run evaluates

    'The tolerance of every comparison, in currency units and as a proportion
    '(docs/methodology/TEST_CASES.md, Tolerance)
    Private Const ABSOLUTE_TOLERANCE    As Double = 0.000001        'Absolute bound, currency units
    Private Const RELATIVE_TOLERANCE    As Double = 0.000000001     'Relative bound, proportion

    'The shift that corrupts an output: one currency unit, far outside both
    'bounds for the totals of these cases, which stay below 2,000,000
    Private Const CORRUPTION            As Double = 1#              'Shift added to a corrupted output

    'The amount by which the monotonicity case raises one input
    Private Const INCREMENT             As Double = 1000#           'Rise of one input


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub RunRelationTests()

'
'------------------------------------------------------------------------------
'                              RUN RELATION TESTS
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case in a fixed order between BeginRun and EndRun. A passing
'   run ends with RESULT=PASS; suite=relations; operation=PASS;
'   completeness=COMPLETE; cases=4; assertions=8; failures=0; cleanup=PASS;
'   recovery=NONE
'
' WHY THIS EXISTS
'   One entry point runs the suite and reports it through the harness, so
'   that a run is complete or visibly incomplete
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Starts the run, calls CaseZeroInput, CaseAdditivity, CaseSymmetry and
'   CaseMonotonicity, then passes any error that escaped them to EndRun
'
' ERROR POLICY
'   An error escaping a case handler is passed to EndRun, which reports it
'   and re-raises it unchanged. BeginRun's refusal propagates directly
'
' DEPENDENCIES
'   TEST_Harness.BeginRun, TEST_Harness.EndRun
'
' NOTES
'   None
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrNumber       As Long     'Runner error number; 0 for none
    Dim ErrSource       As String   'Runner error source
    Dim ErrDescription  As String   'Runner error description

'------------------------------------------------------------------------------
' RUN CASES
'------------------------------------------------------------------------------
    'Start the run; a refused start propagates directly
        TEST_Harness.BeginRun SUITE_NAME, "NORMAL", EXPECTED_CASES, EXPECTED_ASSERTIONS
    'Send an error that escapes a case to the runner handler
        On Error GoTo RunFailed
    'Run every case in a fixed order
        CaseZeroInput
        CaseAdditivity
        CaseSymmetry
        CaseMonotonicity

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
Finish:
    'Restore normal error handling
        On Error GoTo 0
    'Report the run, raising its failure or the escaped error
        TEST_Harness.EndRun ErrNumber, ErrSource, ErrDescription
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
RunFailed:
    'Capture the escaped error
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
    'Finish the run with the captured error
        Resume Finish

End Sub

'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseZeroInput()

'
'------------------------------------------------------------------------------
'                          CHECK THE ZERO-INPUT TOTAL
'------------------------------------------------------------------------------
' PURPOSE
'   Zero inputs give a zero total, matched within the absolute bound only;
'   a total shifted by CORRUPTION is rejected. 2 assertions.
'
' WHY THIS EXISTS
'   A relative bound is meaningless at zero, so this case exercises the
'   absolute bound alone
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Totals three zero amounts, asserts the total is zero, then asserts that
'   the total shifted by CORRUPTION is outside the tolerance
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   IllustrativeTotal, TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Amounts(1 To 3)     As Double   'Three zero amounts, as VBA initializes them
    Dim Total               As Double   'Total of Amounts

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "zero-input"

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Total the zero amounts
        Total = IllustrativeTotal(Amounts)
    'Check the relation, then that a corrupted total is rejected
        TEST_Harness.AssertNear "holds", 0#, Total, ABSOLUTE_TOLERANCE, RELATIVE_TOLERANCE
        TEST_Harness.AssertTrue "detects-corruption", _
            Not TEST_Harness.WithinTolerance(0#, Total + CORRUPTION, ABSOLUTE_TOLERANCE, _
            RELATIVE_TOLERANCE), "a corrupted total of zero amounts was accepted"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the unexpected error against the case
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseAdditivity()

'
'------------------------------------------------------------------------------
'                            CHECK TOTAL ADDITIVITY
'------------------------------------------------------------------------------
' PURPOSE
'   The total of two lists together equals the sum of their totals; a
'   combined total shifted by CORRUPTION is rejected. 2 assertions.
'
' WHY THIS EXISTS
'   Additivity is the relation a total must satisfy whatever its inputs
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Builds two lists and their concatenation, compares the total of the
'   concatenation with the sum of the two totals, then checks that the
'   combined total shifted by CORRUPTION is rejected
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   IllustrativeTotal, TEST_Harness
'
' NOTES
'   The amounts are not exact in binary, so the sums differ in their last
'   bits and the comparison needs its tolerance
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FirstList(1 To 2)   As Double   'The first list
    Dim SecondList(1 To 3)  As Double   'The second list
    Dim BothLists(1 To 5)   As Double   'Both lists together
    Dim Index               As Long     'Position in a list
    Dim Whole               As Double   'Total of BothLists
    Dim Parts               As Double   'Sum of the totals of FirstList and SecondList

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "additivity"

'------------------------------------------------------------------------------
' BUILD INPUTS
'------------------------------------------------------------------------------
    'Fill the two lists with amounts that are not exact in binary
        FirstList(1) = 600000.1
        FirstList(2) = 400000.2
        SecondList(1) = 250000.3
        SecondList(2) = 125000.4
        SecondList(3) = 0.5
    'Copy both lists, in order, into one
        For Index = 1 To 2
            BothLists(Index) = FirstList(Index)
        Next Index
        For Index = 1 To 3
            BothLists(2 + Index) = SecondList(Index)
        Next Index

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Total the combined list and the two lists separately
        Whole = IllustrativeTotal(BothLists)
        Parts = IllustrativeTotal(FirstList) + IllustrativeTotal(SecondList)
    'Check the relation, then that a corrupted total is rejected
        TEST_Harness.AssertNear "holds", Parts, Whole, ABSOLUTE_TOLERANCE, RELATIVE_TOLERANCE
        TEST_Harness.AssertTrue "detects-corruption", _
            Not TEST_Harness.WithinTolerance(Parts, Whole + CORRUPTION, ABSOLUTE_TOLERANCE, _
            RELATIVE_TOLERANCE), "a corrupted combined total was accepted"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the unexpected error against the case
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseSymmetry()

'
'------------------------------------------------------------------------------
'                         CHECK ORDER-INDEPENDENT TOTAL
'------------------------------------------------------------------------------
' PURPOSE
'   The order of the inputs does not change the total; a total of the
'   reversed list shifted by CORRUPTION is rejected. 2 assertions.
'
' WHY THIS EXISTS
'   A total must not depend on the order in which its inputs arrive
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Builds a list and its reverse, compares their totals, then checks that
'   the reversed total shifted by CORRUPTION is rejected
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   IllustrativeTotal, TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Forward(1 To 4)     As Double   'The list
    Dim Reversed(1 To 4)    As Double   'The same list in reverse order
    Dim Index               As Long     'Position in the list
    Dim Expected            As Double   'Total of Forward
    Dim Actual              As Double   'Total of Reversed

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "symmetry"

'------------------------------------------------------------------------------
' BUILD INPUTS
'------------------------------------------------------------------------------
    'Fill the list with amounts of very different sizes
        Forward(1) = 0.1
        Forward(2) = 1000000.7
        Forward(3) = 0.2
        Forward(4) = 333333.3
    'Copy the list in reverse order
        For Index = 1 To 4
            Reversed(Index) = Forward(5 - Index)
        Next Index

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Total the list in both orders
        Expected = IllustrativeTotal(Forward)
        Actual = IllustrativeTotal(Reversed)
    'Check the relation, then that a corrupted total is rejected
        TEST_Harness.AssertNear "holds", Expected, Actual, ABSOLUTE_TOLERANCE, RELATIVE_TOLERANCE
        TEST_Harness.AssertTrue "detects-corruption", _
            Not TEST_Harness.WithinTolerance(Expected, Actual + CORRUPTION, ABSOLUTE_TOLERANCE, _
            RELATIVE_TOLERANCE), "a corrupted total of the reversed list was accepted"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the unexpected error against the case
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseMonotonicity()

'
'------------------------------------------------------------------------------
'                          CHECK MONOTONIC TOTAL RISE
'------------------------------------------------------------------------------
' PURPOSE
'   Raising one input by INCREMENT never lowers the total; a raised total
'   lowered by twice INCREMENT, below the original total, is rejected.
'   2 assertions.
'
' WHY THIS EXISTS
'   A larger input must never give a smaller total
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Builds a list and a copy with its second amount raised by INCREMENT,
'   asserts the raised total is not below the original, then checks that a
'   raised total lowered by twice INCREMENT is rejected
'
' ERROR POLICY
'   An unexpected error is recorded against the case by
'   TEST_Harness.RecordCaseError and does not stop the run
'
' DEPENDENCIES
'   IllustrativeTotal, TEST_Harness
'
' NOTES
'   None
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Original(1 To 3)    As Double   'The list
    Dim Raised(1 To 3)      As Double   'The same list with its second amount raised
    Dim Index               As Long     'Position in the list
    Dim OriginalTotal       As Double   'Total of Original
    Dim RaisedTotal         As Double   'Total of Raised

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error against the case
        On Error GoTo CaseFailed
    'Start the case
        TEST_Harness.BeginCase "monotonicity"

'------------------------------------------------------------------------------
' BUILD INPUTS
'------------------------------------------------------------------------------
    'Fill the list
        Original(1) = 500000#
        Original(2) = 250000.25
        Original(3) = 125000.125
    'Copy the list and raise its second amount
        For Index = 1 To 3
            Raised(Index) = Original(Index)
        Next Index
        Raised(2) = Raised(2) + INCREMENT

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Total both lists
        OriginalTotal = IllustrativeTotal(Original)
        RaisedTotal = IllustrativeTotal(Raised)
    'Check the relation, then that a corrupted total is rejected
        TEST_Harness.AssertTrue "holds", RaisedTotal >= OriginalTotal, _
            "raised total " & CStr(RaisedTotal) & " is below " & CStr(OriginalTotal)
        TEST_Harness.AssertTrue "detects-corruption", _
            Not (RaisedTotal - 2# * INCREMENT >= OriginalTotal), _
            "a corrupted raised total was accepted"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the unexpected error against the case
        TEST_Harness.RecordCaseError

End Sub

'
'------------------------------------------------------------------------------
'
'                         EXAMPLE PROCEDURE UNDER TEST
'
'------------------------------------------------------------------------------
'

Private Function IllustrativeTotal( _
    ByRef Amounts() As Double) _
    As Double
'
'------------------------------------------------------------------------------
'                              TOTAL THE AMOUNTS
'------------------------------------------------------------------------------
' PURPOSE
'   The sum of Amounts: a stand-in for the calculation the relations test
'
' WHY THIS EXISTS
'   The relation cases need a calculation to run on before the workbook has
'   one of its own
'
' INPUTS
'   Amounts
'     Amounts in one currency, read only; VBA passes an array by reference
'
' RETURNS
'   The sum of every element of Amounts
'
' BEHAVIOR
'   Adds the elements from LBound to UBound
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   Replace it with a call to that calculation's core procedure
'
' UPDATED
'   2026-11-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Index   As Long     'Position in Amounts
    Dim Total   As Double   'Sum so far

'------------------------------------------------------------------------------
' ADD UP
'------------------------------------------------------------------------------
    'Add every amount
        For Index = LBound(Amounts) To UBound(Amounts)
            Total = Total + Amounts(Index)
        Next Index
    'Return the sum
        IllustrativeTotal = Total

End Function
