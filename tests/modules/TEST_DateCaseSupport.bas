Attribute VB_Name = "TEST_DateCaseSupport"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: TEST_DateCaseSupport
'------------------------------------------------------------------------------
' PURPOSE
'   The helpers of the worksheet-function case suite: building the arguments of a case and
'   writing a result in the value encoding of docs/methodology/TEST_CASES.md
'
' WHY THIS EXISTS
'   TEST_DateCases compares each result with the encoded text of its expected file and records
'   that text for the exact accuracy contracts. The encoding and the arrays are written once here
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - Encode          the encoded text of a value, an array as arr2
'     - EncodeFilled    the same, but an array whose items all encode alike as fill2
'     - A1, A2, F2      a one-dimensional, a two-dimensional and a filled array
'
' ERROR POLICY
'   Encode returns the text "unsupported:" and the type name for a value the encoding has no form
'   for, so the case's assertion fails and names it; the array builders raise nothing
'
' DEPENDENCIES
'   None
'
' NOTES
'   - Option Private Module keeps every member out of the workbook's external surface
'   - The VBA functions are called unqualified: LibreOffice, which runs the early warning of
'     tools/devbuild/lo_smoke.py, resolves Join and Str$ only so
'   - A date is written from its year, month and day, and a whole number with CStr of a Long,
'     neither of which depends on the regional settings
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------


'
'------------------------------------------------------------------------------
'
'                                   ENCODING
'
'------------------------------------------------------------------------------
'

Public Function Encode( _
    ByVal Value As Variant) As String
'
'------------------------------------------------------------------------------
'                              ENCODE ONE RESULT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the encoded text of a value, as the expected files of the cases write it.
'
' WHY THIS EXISTS
'   The cases compare texts, so a result's type is compared with its value: a Date is not a
'   Double, and a Long is not a Boolean.
'
' INPUTS
'   Value
'     A scalar result or a two-dimensional array of results.
'
' RETURNS
'   String
'     date:, lng:, num:, str:, bool:, err:, empty or null for a scalar; arr2:RxC[...] for an
'     array, its items by rows.
'
' BEHAVIOR
'   Encodes an array item by item, and a scalar by its VarType.
'
' ERROR POLICY
'   Does not raise. A value of another type gives "unsupported:" and its type name.
'
' DEPENDENCIES
'   EncodeArray, EncodeScalar
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ENCODE
'------------------------------------------------------------------------------
    'An array lists every item; a scalar encodes alone
        If IsArray(Value) Then
            Encode = EncodeArray(Value, False)
        Else
            Encode = EncodeScalar(Value)
        End If

End Function

Public Function EncodeFilled( _
    ByVal Value As Variant) As String
'
'------------------------------------------------------------------------------
'                      ENCODE A RESULT THAT MAY BE FILLED
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the encoded text of a value, writing an array whose items all encode alike as fill2.
'
' WHY THIS EXISTS
'   A case whose arguments fill an array of up to 100,000 elements expects a fill2 text, which
'   keeps its expected file small.
'
' INPUTS
'   Value
'     A scalar result or a two-dimensional array of results.
'
' RETURNS
'   String
'     fill2:RxC[item] for an array whose items all encode alike; otherwise as Encode.
'
' BEHAVIOR
'   Encodes the array's items, then writes one of them when all are equal.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   EncodeArray, EncodeScalar
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ENCODE
'------------------------------------------------------------------------------
    'An array may be filled; a scalar encodes alone
        If IsArray(Value) Then
            EncodeFilled = EncodeArray(Value, True)
        Else
            EncodeFilled = EncodeScalar(Value)
        End If

End Function

Private Function EncodeArray( _
    ByVal Value As Variant, _
    ByVal Filled As Boolean) As String
'
'------------------------------------------------------------------------------
'                               ENCODE AN ARRAY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the arr2 text of a two-dimensional array, or its fill2 text when asked and possible.
'
' WHY THIS EXISTS
'   The facade returns an array result as a two-dimensional array of its arguments' shape.
'
' INPUTS
'   Value
'     A two-dimensional array.
'
'   Filled
'     True to write an array whose items all encode alike as fill2.
'
' RETURNS
'   String
'     arr2:RxC[...] or fill2:RxC[item]; "unsupported:array" for an array of another rank.
'
' BEHAVIOR
'   Encodes the items by rows into a list, then joins it.
'
' ERROR POLICY
'   Does not raise: a failure to read the bounds gives "unsupported:array".
'
' DEPENDENCIES
'   EncodeScalar
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
    Const PROC_NAME     As String = "TEST_DateCaseSupport.EncodeArray"

    Dim RowCount        As Long         'Rows of the array
    Dim ColCount        As Long         'Columns of the array
    Dim Items()         As String       'The encoded items, by rows
    Dim R               As Long         'Row offset
    Dim C               As Long         'Column offset
    Dim Same            As Boolean      'True while every item equals the first

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'An array of another rank, or unreadable bounds, has no encoding
        On Error GoTo Unsupported
        RowCount = UBound(Value, 1) - LBound(Value, 1) + 1
        ColCount = UBound(Value, 2) - LBound(Value, 2) + 1
        On Error GoTo 0

'------------------------------------------------------------------------------
' ENCODE ITEMS
'------------------------------------------------------------------------------
    'Encode the items by rows, noting whether they all match the first
        ReDim Items(0 To RowCount * ColCount - 1)
        Same = True
        For R = 0 To RowCount - 1
            For C = 0 To ColCount - 1
                Items(R * ColCount + C) = EncodeScalar(Value(LBound(Value, 1) + R, LBound(Value, 2) + C))
                If Items(R * ColCount + C) <> Items(0) Then Same = False
            Next C
        Next R

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Write one item of a filled array, or every item
        If Filled And Same Then
            EncodeArray = "fill2:" & CStr(RowCount) & "x" & CStr(ColCount) & "[" & Items(0) & "]"
        Else
            EncodeArray = "arr2:" & CStr(RowCount) & "x" & CStr(ColCount) & "[" & _
                Join(Items, ";") & "]"
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Unsupported:
    'An array this encoding does not hold
        Err.Clear
        EncodeArray = "unsupported:array"

End Function

Private Function EncodeScalar( _
    ByVal Value As Variant) As String
'
'------------------------------------------------------------------------------
'                               ENCODE A SCALAR
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the encoded text of one scalar value.
'
' WHY THIS EXISTS
'   Each item of a result, or a scalar result, is compared as its encoded text.
'
' INPUTS
'   Value
'     A scalar.
'
' RETURNS
'   String
'     date:YYYY-MM-DD, lng:N, num:N, str:"...", bool:TRUE or FALSE, err:#..., empty or null.
'
' BEHAVIOR
'   Selects the form by VarType; a Date with a time keeps it as datetime. A date is written
'   from its year, month and day, never with Format$, whose output can follow the settings.
'
' ERROR POLICY
'   Does not raise. Another type gives "unsupported:" and its type name.
'
' DEPENDENCIES
'   ErrorName, QuotedText, TwoDigits
'
' NOTES
'   A Double is written with Str$, whose decimal point does not depend on the regional settings
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ENCODE
'------------------------------------------------------------------------------
    'Select the form of the value's type
        Select Case VarType(Value)
            Case vbDate
                EncodeScalar = "date:" & CStr(Year(Value)) & "-" & TwoDigits(Month(Value)) & _
                    "-" & TwoDigits(Day(Value))
                If CDbl(Value) <> Int(CDbl(Value)) Then
                    EncodeScalar = "datetime:" & Mid$(EncodeScalar, 6) & "T" & TwoDigits(Hour(Value)) & _
                        ":" & TwoDigits(Minute(Value)) & ":" & TwoDigits(Second(Value))
                End If
            Case vbLong, vbInteger
                EncodeScalar = "lng:" & CStr(Value)
            Case vbDouble
                EncodeScalar = "num:" & Trim$(Str$(Value))
            Case vbString
                EncodeScalar = "str:" & QuotedText(CStr(Value))
            Case vbBoolean
                EncodeScalar = "bool:" & IIf(Value, "TRUE", "FALSE")
            Case vbError
                EncodeScalar = "err:" & ErrorName(Value)
            Case vbEmpty
                EncodeScalar = "empty"
            Case vbNull
                EncodeScalar = "null"
            Case Else
                EncodeScalar = "unsupported:" & TypeName(Value)
        End Select

End Function

Private Function TwoDigits( _
    ByVal Number As Long) As String
'
'------------------------------------------------------------------------------
'                               WRITE TWO DIGITS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a number from 0 to 99 as two digits, with a leading zero below 10.
'
' WHY THIS EXISTS
'   A month, a day and the parts of a time are written with two digits.
'
' INPUTS
'   Number
'     A whole number from 0 to 99.
'
' RETURNS
'   String
'     Two digits.
'
' BEHAVIOR
'   Pads with one zero below 10.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   None
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
    'Pad one digit with a zero
        TwoDigits = Right$("0" & CStr(Number), 2)

End Function

Private Function ErrorName( _
    ByVal Value As Variant) As String
'
'------------------------------------------------------------------------------
'                          NAME AN EXCEL ERROR VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the worksheet name of an Excel error value, such as #N/A.
'
' WHY THIS EXISTS
'   The encoding names an error as a worksheet shows it, not by its number.
'
' INPUTS
'   Value
'     A Variant of the Error subtype.
'
' RETURNS
'   String
'     #NULL!, #DIV/0!, #VALUE!, #REF!, #NAME?, #NUM! or #N/A; another number as #ERROR and it.
'
' BEHAVIOR
'   Maps the error number with CLng.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' NAME
'------------------------------------------------------------------------------
    'Map the native error numbers
        Select Case CLng(Value)
            Case 2000
                ErrorName = "#NULL!"
            Case 2007
                ErrorName = "#DIV/0!"
            Case 2015
                ErrorName = "#VALUE!"
            Case 2023
                ErrorName = "#REF!"
            Case 2029
                ErrorName = "#NAME?"
            Case 2036
                ErrorName = "#NUM!"
            Case 2042
                ErrorName = "#N/A"
            Case Else
                ErrorName = "#ERROR" & CStr(CLng(Value))
        End Select

End Function

Private Function QuotedText( _
    ByVal Text As String) As String
'
'------------------------------------------------------------------------------
'                             QUOTE A TEXT AS JSON
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a text in double quotes, with its quotes and backslashes escaped, as JSON writes it.
'
' WHY THIS EXISTS
'   The encoding writes a text as a JSON string, so that outer spaces show.
'
' INPUTS
'   Text
'     A text of printable ASCII.
'
' RETURNS
'   String
'     The quoted text.
'
' BEHAVIOR
'   Escapes the backslash first, then the quote.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   None
'
' NOTES
'   A result never holds another character: the pillar tokens are ASCII
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' QUOTE
'------------------------------------------------------------------------------
    'Escape, then quote
        QuotedText = """" & Replace(Replace(Text, "\", "\\"), """", "\""") & """"

End Function

'
'------------------------------------------------------------------------------
'
'                                    ARRAYS
'
'------------------------------------------------------------------------------
'

Public Function A1( _
    ParamArray Items() As Variant) As Variant
'
'------------------------------------------------------------------------------
'                        BUILD A ONE-DIMENSIONAL ARRAY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a one-dimensional Variant array of the items, from index 0.
'
' WHY THIS EXISTS
'   A case's argument encoded as arr1 is built from its items.
'
' INPUTS
'   Items
'     The items, possibly none, possibly arrays themselves.
'
' RETURNS
'   Variant
'     An array of the items; with none, an empty array whose upper bound is below its lower.
'
' BEHAVIOR
'   Copies the items, so that the result is a plain array rather than the ParamArray.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   None
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
    Const PROC_NAME     As String = "TEST_DateCaseSupport.A1"

    Dim Result()        As Variant      'The array built
    Dim Index           As Long         'Position of an item

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
    'An empty list gives an empty array
        If UBound(Items) < LBound(Items) Then
            A1 = Array()
            Exit Function
        End If
    'Copy the items from index 0
        ReDim Result(0 To UBound(Items) - LBound(Items))
        For Index = 0 To UBound(Result)
            If IsObject(Items(LBound(Items) + Index)) Then
                Set Result(Index) = Items(LBound(Items) + Index)
            Else
                Result(Index) = Items(LBound(Items) + Index)
            End If
        Next Index
        A1 = Result

End Function

Public Function A2( _
    ByVal RowCount As Long, _
    ByVal ColCount As Long, _
    ParamArray Items() As Variant) As Variant
'
'------------------------------------------------------------------------------
'                        BUILD A TWO-DIMENSIONAL ARRAY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a two-dimensional Variant array of the items, by rows, from index 1.
'
' WHY THIS EXISTS
'   A case's argument encoded as arr2 is built from its dimensions and items.
'
' INPUTS
'   RowCount, ColCount
'     The dimensions, each at least 1.
'
'   Items
'     RowCount times ColCount items, by rows.
'
' RETURNS
'   Variant
'     The array, 1 to RowCount by 1 to ColCount, as a worksheet passes a Range's values.
'
' BEHAVIOR
'   Places item k at row k \ ColCount + 1 and column k Mod ColCount + 1.
'
' ERROR POLICY
'   Does not raise for a list of the right length, which the generator writes.
'
' DEPENDENCIES
'   None
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
    Const PROC_NAME     As String = "TEST_DateCaseSupport.A2"

    Dim Result()        As Variant      'The array built
    Dim Index           As Long         'Position of an item in the list

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
    'Place the items by rows
        ReDim Result(1 To RowCount, 1 To ColCount)
        For Index = 0 To RowCount * ColCount - 1
            Result(Index \ ColCount + 1, Index Mod ColCount + 1) = Items(LBound(Items) + Index)
        Next Index
        A2 = Result

End Function

Public Function F2( _
    ByVal RowCount As Long, _
    ByVal ColCount As Long, _
    ByVal Item As Variant) As Variant
'
'------------------------------------------------------------------------------
'                             BUILD A FILLED ARRAY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a two-dimensional Variant array with the same item in every position.
'
' WHY THIS EXISTS
'   The capacity cases pass arrays of 100,000 and 100,001 elements, which their fixtures write
'   as one item and the dimensions.
'
' INPUTS
'   RowCount, ColCount
'     The dimensions, each at least 1.
'
'   Item
'     The item of every position, possibly an array.
'
' RETURNS
'   Variant
'     The array, 1 to RowCount by 1 to ColCount.
'
' BEHAVIOR
'   Fills every position with the item.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   None
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
    Const PROC_NAME     As String = "TEST_DateCaseSupport.F2"

    Dim Result()        As Variant      'The array built
    Dim R               As Long         'Row
    Dim C               As Long         'Column

'------------------------------------------------------------------------------
' BUILD
'------------------------------------------------------------------------------
    'Fill every position
        ReDim Result(1 To RowCount, 1 To ColCount)
        For R = 1 To RowCount
            For C = 1 To ColCount
                Result(R, C) = Item
            Next C
        Next R
        F2 = Result

End Function
