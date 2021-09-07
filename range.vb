' function library for ranges
' dependencies:
'   string

Function visibleCellsCountNotZero(rng As Range) As Boolean
' returns if the count of visible cells is NOT zero
' useful with range.autofilter for deleting rows that match a criteria
' eg. if visibleCellCountNotZero(selection) then selection.entirerow.delete
    Dim tmp As Long

    On Error GoTo error

    If rng.SpecialCells(xlCellTypeVisible).count Then
        visibleCellsCountNotZero = True
        Exit Function
    End If

error:
    visibleCellsCountNotZero = False
End Function

Function sumRange(r As String) As Double
' returns the sum of a range of numbers as double
    Dim rng As Range
    Dim cel As Range
    Dim sum As Double

    Set rng = Range(r)
    sum = 0
    For Each cel In rng
        If VarType(cel.value) = vbDouble Or VarType(cel.value) = vbInteger Then
            sum = sum + cel.value
        End If
    Next cel
    sumRange = sum
End Function

Function firstCell(r As Range) As String
' returns the address of the first cell
    Dim cel As Range
    For Each cel In r
        firstCell = cel.Address(0, 0)
        Exit Function
    Next cel
End Function

Function lastCell(r As Range) As String
' returns the address of the last cell
    Dim cel As Range
    Dim last As Range
    For Each cel In r
        Set last = cel
    Next cel
    lastCell = last.Address(0, 0)
End Function

Function nRows(r As Range) As Long
' returns the number of rows in the provided range
    Dim a As Long
    Dim b As Long

    If word_count(r.Address(0, 0), ":") > 0 Then
        a = numbers_only(Split(r.Address, ":")(1))
        b = numbers_only(Split(r.Address, ":")(0))
    Else
        nRows = 1
        Exit Function
    End If
    nRows = a - b + 1
End Function

Function nCols(r As Range) As Long
' returns the number of columns in the provided range
    Dim a As Long
    Dim b As Long

    If word_count(r.Address(0, 0), ":") > 0 Then
        a = Range(Split(r.Address, ":")(1)).Column
        b = Range(Split(r.Address, ":")(0)).Column
    Else
        nCols = 1
        Exit Function
    End If
    nCols = a - b + 1
End Function

Function contiguousRange() As String
' returns the contiguous range of cells with origin = selection
    Dim start_r As Long
    Dim end_r As Long
    Dim start_c As Long
    Dim end_c As Long
    start_r = Selection.End(xlUp).Row
    end_r = Selection.End(xlDown).Row
    start_c = Selection.End(xlToRight).Column
    end_c = Selection.End(xlToLeft).Column
    contiguousRange = Range(Cells(start_r, start_c), Cells(end_r, end_c)).Address
End Function


Function nextCol(r As Range, offset As Long) As String
' return the range (as string) of the next sub-column in provided range
    Dim c As Long
    Dim r_start As Long
    Dim r_end As Long
    Dim addr As String

    If word_count(r.Address, ":") > 0 Then
        addr = Split(r.Address, ":")(0)
        c = Range(addr).Column + offset - 1
        r_start = Range(addr).Row
        r_end = numbers_only(Split(r.Address, ":")(1))
    End If
    nextCol = colToLetter(c) & r_start & ":" & colToLetter(c) & r_end
End Function
