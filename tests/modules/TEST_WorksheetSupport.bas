Attribute VB_Name = "TEST_WorksheetSupport"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: TEST_WorksheetSupport
'------------------------------------------------------------------------------
' PURPOSE
'   The helpers of the suites that call the worksheet functions from a worksheet: a scratch
'   workbook in either date system, the arguments written to its cells, a formula calculated
'   there and its result in the value encoding of docs/methodology/TEST_CASES.md
'
' WHY THIS EXISTS
'   A worksheet caller reads its date system from the workbook that holds the formula and its
'   arguments from Ranges. Only a formula in a workbook of that date system tests that path,
'   and the steps that open, fill and close one are written once here
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - OpenScratch     a new workbook in the 1900 or 1904 date system, with manual
'                       calculation
'     - CloseScratch    closes it without saving and restores the calculation mode
'     - CalculateCall   writes the arguments to cells, calculates one formula that refers to
'                       them and returns its encoded result
'     - Qualified       a worksheet function's name as a formula in another workbook calls it
'     - EncodeCell      the encoded text of a cell's value, read as the expected kind
'
' ERROR POLICY
'   OpenScratch and CalculateCall raise what the Excel object model raises; CloseScratch does not
'   raise and returns what failed
'
' DEPENDENCIES
'   CORE_Array, TEST_DateCaseSupport; the Excel object model
'
' NOTES
'   - Option Private Module keeps every member out of the workbook's external surface
'   - The formula calls the functions of the workbook that holds this module, through its
'     name, so the scratch workbook never holds code
'   - A cell holds a number, not a date: EncodeCell reads a number as a date or a whole number
'     when the expected text is one, so a cell is compared with its case as a date or as a
'     number. Every date a case expects is in the 1900 date system
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'Where the arguments and the formula stand on the scratch sheet
    Private Const ARGUMENT_ROW      As Long = 3     'First row of every argument
    Private Const ARGUMENT_SPACING  As Long = 26    'Columns between two arguments
    Private Const FORMULA_CELL      As String = "A1"  'The cell of the formula

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    'The scratch workbook and the calculation mode it replaced
    Private mScratch            As Workbook         'The open scratch workbook, or Nothing
    Private mSavedCalculation   As XlCalculation    'Calculation mode before OpenScratch
    Private mCalculationDue     As Boolean          'True while the mode must be restored


'
'------------------------------------------------------------------------------
'
'                              SCRATCH WORKBOOK
'
'------------------------------------------------------------------------------
'

Public Function OpenScratch( _
    ByVal Date1904 As Boolean) As Worksheet

'
'------------------------------------------------------------------------------
'                          OPEN A SCRATCH WORKBOOK
'------------------------------------------------------------------------------
' PURPOSE
'   Opens a new workbook in the date system asked for, with manual calculation, and returns
'   its first sheet.
'
' WHY THIS EXISTS
'   A worksheet caller's date system is its workbook's; a new workbook leaves the development
'   workbook unchanged.
'
' INPUTS
'   Date1904
'     True for the 1904 date system, False for the 1900 one.
'
' RETURNS
'   Worksheet
'     The first sheet of the new workbook.
'
' BEHAVIOR
'   Keeps the calculation mode and registers its restoration before setting it to manual, then
'   adds the workbook and sets its date system. Nothing calculates unless a case asks.
'
' ERROR POLICY
'   Raises what the Excel object model raises; CloseScratch restores what was registered.
'
' DEPENDENCIES
'   The Excel object model
'
' NOTES
'   A workbook still open from an earlier case is closed first
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Detail          As String       'What an earlier close could not do

'------------------------------------------------------------------------------
' OPEN
'------------------------------------------------------------------------------
    'Close a workbook an earlier case left open
        If Not mScratch Is Nothing Then
            Detail = CloseScratch()
            If Len(Detail) > 0 Then
                Err.Raise vbObjectError + 1, "TEST_WorksheetSupport.OpenScratch", Detail
            End If
        End If
    'Keep the calculation mode and register its restoration before changing it
        mSavedCalculation = Application.Calculation
        mCalculationDue = True
        Application.Calculation = xlCalculationManual
    'Add the workbook and set its date system
        Set mScratch = Application.Workbooks.Add
        mScratch.Date1904 = Date1904
        Set OpenScratch = mScratch.Worksheets(1)

End Function

Public Function CloseScratch() As String

'
'------------------------------------------------------------------------------
'                         CLOSE THE SCRATCH WORKBOOK
'------------------------------------------------------------------------------
' PURPOSE
'   Closes the scratch workbook without saving and restores the calculation mode.
'
' WHY THIS EXISTS
'   A case leaves Excel as it found it on its success and its error path alike.
'
' INPUTS
'   None
'
' RETURNS
'   String
'     Empty when everything was restored, otherwise what failed, joined by "; ".
'
' BEHAVIOR
'   Closes the workbook if one is open, then restores the calculation mode if it was
'   registered. Each step runs even when the other failed.
'
' ERROR POLICY
'   Does not raise: each failure is returned.
'
' DEPENDENCIES
'   The Excel object model
'
' NOTES
'   Safe to call when nothing is open
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Detail          As String       'Failures, joined by "; "

'------------------------------------------------------------------------------
' CLOSE
'------------------------------------------------------------------------------
    'Contain each failure so that every step runs
        On Error Resume Next
    'Close the workbook without saving it
        If Not mScratch Is Nothing Then
            mScratch.Close SaveChanges:=False
            If Err.Number <> 0 Then
                Detail = "the scratch workbook did not close: " & Err.Description
                Err.Clear
            End If
            Set mScratch = Nothing
        End If
    'Restore the calculation mode
        If mCalculationDue Then
            Application.Calculation = mSavedCalculation
            If Err.Number <> 0 Then
                If Len(Detail) > 0 Then Detail = Detail & "; "
                Detail = Detail & "the calculation mode was not restored: " & Err.Description
                Err.Clear
            End If
            mCalculationDue = False
        End If
    'Restore normal error handling and return what failed
        On Error GoTo 0
        CloseScratch = Detail

End Function

'
'------------------------------------------------------------------------------
'
'                                  FORMULAS
'
'------------------------------------------------------------------------------
'

Public Function CalculateCall( _
    ByVal Sheet As Worksheet, _
    ByVal FunctionName As String, _
    ByVal Kind As String, _
    ParamArray Arguments() As Variant) As String

'
'------------------------------------------------------------------------------
'                     CALCULATE ONE FORMULA ON THE SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the arguments to cells, calculates a formula that calls the function on them and
'   returns its encoded result.
'
' WHY THIS EXISTS
'   A worksheet-function case of a worksheet caller passes its arguments as Ranges.
'
' INPUTS
'   Sheet
'     A sheet of the scratch workbook.
'
'   FunctionName
'     The worksheet function, such as KPR_Dates_EndOfMonth.
'
'   Kind
'     The kind of the expected text, such as date or lng, which EncodeCell reads a number as.
'
'   Arguments
'     The case's arguments: a scalar, or a one- or two-dimensional array.
'
' RETURNS
'   String
'     The encoded value of the formula's cell.
'
' BEHAVIOR
'   Clears the sheet, writes each argument from row ARGUMENT_ROW, ARGUMENT_SPACING columns
'   apart, enters the formula in FORMULA_CELL, calculates and encodes the cell's value.
'
' ERROR POLICY
'   Raises what the Excel object model raises.
'
' DEPENDENCIES
'   PutArgument, Qualified, EncodeCell
'
' NOTES
'   Application.Calculate is an ordinary recalculation, as a user's
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim References      As String       'The arguments' addresses, comma-separated
    Dim Index           As Long         'Position of an argument

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Start from an empty sheet
        Sheet.Cells.Clear
    'Write each argument and collect its address
        For Index = LBound(Arguments) To UBound(Arguments)
            If Index > LBound(Arguments) Then References = References & ","
            References = References & PutArgument(Sheet, Index - LBound(Arguments), Arguments(Index))
        Next Index

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Enter the formula, calculate and encode its value
        Sheet.Range(FORMULA_CELL).Formula = "=" & Qualified(FunctionName) & "(" & References & ")"
        Application.Calculate
        CalculateCall = EncodeCell(Sheet.Range(FORMULA_CELL).Value2, Kind)

End Function

Public Function Qualified( _
    ByVal FunctionName As String) As String

'
'------------------------------------------------------------------------------
'                    QUALIFY A FUNCTION WITH ITS WORKBOOK
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the function's name as a formula in another workbook calls it.
'
' WHY THIS EXISTS
'   The scratch workbook holds no code: its formulas call the functions of this workbook.
'
' INPUTS
'   FunctionName
'     The worksheet function.
'
' RETURNS
'   String
'     'Workbook name'!FunctionName, with an apostrophe in the name doubled.
'
' BEHAVIOR
'   Quotes the name of the workbook that holds this module.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   ThisWorkbook
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' QUALIFY
'------------------------------------------------------------------------------
    'Quote the workbook name and append the function
        Qualified = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & FunctionName

End Function

Public Function EncodeCell( _
    ByVal Value As Variant, _
    ByVal Kind As String) As String

'
'------------------------------------------------------------------------------
'                           ENCODE A CELL'S VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the encoded text of a cell's value, reading a number as the expected kind.
'
' WHY THIS EXISTS
'   A cell holds a date or a whole number as a Double; the case's expected text says which.
'
' INPUTS
'   Value
'     The cell's Value2.
'
'   Kind
'     The prefix of the expected text: date, lng or another.
'
' RETURNS
'   String
'     The encoded text, as TEST_DateCaseSupport.Encode writes it.
'
' BEHAVIOR
'   A Double read as date becomes the Date of that 1900-system serial; a whole Double in the
'   Long range read as lng becomes a Long. Every other value is encoded as it is, so a value
'   of another type fails the comparison.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   TEST_DateCaseSupport.Encode
'
' NOTES
'   A serial below 61 is not read as a date: VBA and the 1900 date system disagree there
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ENCODE
'------------------------------------------------------------------------------
    'Read a number as the expected kind, then encode it
        If VarType(Value) = vbDouble Then
            If Kind = "date" And Value >= 61 And Value < 2958466 Then
                EncodeCell = TEST_DateCaseSupport.Encode(CDate(Value))
                Exit Function
            End If
            If Kind = "lng" And Value = Int(Value) And Abs(Value) <= 2147483647# Then
                EncodeCell = TEST_DateCaseSupport.Encode(CLng(Value))
                Exit Function
            End If
        End If
        EncodeCell = TEST_DateCaseSupport.Encode(Value)

End Function

'
'------------------------------------------------------------------------------
'
'                                  ARGUMENTS
'
'------------------------------------------------------------------------------
'

Private Function PutArgument( _
    ByVal Sheet As Worksheet, _
    ByVal Position As Long, _
    ByVal Value As Variant) As String

'
'------------------------------------------------------------------------------
'                         WRITE ONE ARGUMENT TO CELLS
'------------------------------------------------------------------------------
' PURPOSE
'   Writes one argument to the cells of its position and returns their address.
'
' WHY THIS EXISTS
'   A worksheet caller passes each argument as a Range.
'
' INPUTS
'   Sheet
'     The scratch sheet.
'
'   Position
'     The argument's position, from 0.
'
'   Value
'     A scalar, or a one- or two-dimensional array.
'
' RETURNS
'   String
'     The address of the cell or the block, such as A3 or AA3:AC3.
'
' BEHAVIOR
'   A scalar fills one cell, a one-dimensional array a row and a two-dimensional array a block,
'   each element through PutCell.
'
' ERROR POLICY
'   Raises what the Excel object model raises.
'
' DEPENDENCIES
'   PutCell, CORE_Array.ArrayRank
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Anchor          As Range        'The argument's first cell
    Dim RowIndex        As Long         'Row of an element
    Dim ColIndex        As Long         'Column of an element
    Dim RowCount        As Long         'Rows of the block
    Dim ColCount        As Long         'Columns of the block

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Find the argument's first cell
        Set Anchor = Sheet.Cells(ARGUMENT_ROW, 1 + Position * ARGUMENT_SPACING)
    'A scalar fills one cell
        If Not IsArray(Value) Then
            PutCell Anchor, Value
            PutArgument = Anchor.Address(False, False)
            Exit Function
        End If
    'An array fills a row or a block
        If CORE_Array.ArrayRank(Value) = 1 Then
            ColCount = UBound(Value) - LBound(Value) + 1
            For ColIndex = 1 To ColCount
                PutCell Anchor.Offset(0, ColIndex - 1), Value(LBound(Value) + ColIndex - 1)
            Next ColIndex
            RowCount = 1
        Else
            RowCount = UBound(Value, 1) - LBound(Value, 1) + 1
            ColCount = UBound(Value, 2) - LBound(Value, 2) + 1
            For RowIndex = 1 To RowCount
                For ColIndex = 1 To ColCount
                    PutCell Anchor.Offset(RowIndex - 1, ColIndex - 1), _
                        Value(LBound(Value, 1) + RowIndex - 1, LBound(Value, 2) + ColIndex - 1)
                Next ColIndex
            Next RowIndex
        End If
        PutArgument = Anchor.Resize(RowCount, ColCount).Address(False, False)

End Function

Private Sub PutCell( _
    ByVal Cell As Range, _
    ByVal Value As Variant)

'
'------------------------------------------------------------------------------
'                          WRITE ONE VALUE TO A CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Writes one scalar to a cell so that the cell holds the same value and type.
'
' WHY THIS EXISTS
'   Excel converts typed text such as 2024-02-10 to a date; a case's text must stay text.
'
' INPUTS
'   Cell
'     The cell.
'
'   Value
'     A scalar: a Date, a number, text, a Boolean, an error value or Empty.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Text is written to a cell formatted as text; Empty leaves the cell blank; every other value
'   is written as it is, a Date becoming the serial of the workbook's date system.
'
' ERROR POLICY
'   Raises what the Excel object model raises.
'
' DEPENDENCIES
'   The Excel object model
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Keep text as text, leave Empty blank, and write every other value
        Select Case VarType(Value)
            Case vbString
                Cell.NumberFormat = "@"
                Cell.Value = Value
            Case vbEmpty
                Cell.ClearContents
            Case Else
                Cell.Value = Value
        End Select

End Sub
