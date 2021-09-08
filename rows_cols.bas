' functions pertaining to rows and columns

Public Function has_row(text As String) As Boolean
' finds the first column header that matches given text in the given sheet
    Dim thissheet As Long
    Dim tmp As Long
    On Error GoTo error

    tmp = Cells.find(What:=text, after:=Cells(1, 1), LookIn:=xlValues, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row

    has_row = True
    Exit Function

error:
    has_row = False
End Function

Public Function lastcol(r As Long) As Long
' returns last col of argument row # '
    lastcol = ActiveSheet.Cells(r, Columns.count).End(xlToLeft).Column
End Function

Public Function lastrow(c As String) As Long
' returns last row of column c (provide character instead of number) '
    lastrow = ActiveSheet.Cells(Rows.count, c).End(xlUp).Row
End Function

Function colToLetter(lngCol As Long) As String
' split returns an array, so use the (1) to return only
' the first member of the array, which is a string
    colToLetter = Split(Cells(1, lngCol).Address, "$")(1)
End Function

Function letterToCol(col As String) As Long
' converts a column letter to long
    colLettertoNum = Range(col & 1).Column
End Function

Function getMaxRow() As Long
' gets the max row in the sheet
' this is an important constant for detecting errors in a couple functions
    getMaxRow = ActiveSheet.Rows.count
End Function


Function firstRow(c As String, offset As Long) As Long
' returns first row that has a value '
' offset can be used to ignore headers, etc. '

    Dim r As Long
    r = 1 + offset
    Range(c & r).Select
    Do While ActiveCell.value = xlblank
        Range(c & r).Select
        r = r + 1
    Loop
    firstRow = r
End Function


Public Function findColAfter(sheet As String, text As String, aftr As Range) As Long
' finds the first column header that matches given text in the given sheet
    Dim thissheet As Long
    Dim tmp As Long
    On Error GoTo error

    thissheet = ActiveSheet.Index
    If thissheet <> Sheets(sheet).Index Then
        Sheets(sheet).Select
    End If

    tmp = Cells.find(What:=text, after:=aftr, LookIn:=xlValues, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Column
    Sheets(thissheet).Select
    findColAfter = tmp
    Exit Function

error:
    MsgBox "Column not found: " & text
End Function

Function getRowStart(criteria As String, rng As Range) As Long
' gets a criteria and a range and returns the row that corresponds to the start of the criteria in the range
' assuming the criteria is continguious in the range AND that the criteria is to be found in the range
    Dim cel As Range
    For Each cel In rng
        If cel.value = criteria Then
            getRowStart = cel.Row
            Exit Function
        End If
    Next cel
End Function

Function getRowEnd(criteria As String, rng As Range) As Long
' gets a criteria and a range and returns the row that corresponds to the end of the criteria in the range
' assuming the criteria is continguious in the range AND that the criteria is to be found in the range
    Dim cel As Range
    Dim last As Range
    Dim in_range As Boolean
    For Each cel In rng
        If in_range = True Then
            If cel.value = criteria Then
                Set last = cel
            ElseIf cel.value <> criteria Then
                getRowEnd = last.Row
                Exit Function
            End If
        ElseIf in_range = False And cel.value = criteria Then
            in_range = True
            Set last = cel
        End If
    Next cel
End Function

Public Function findFirstCol(sheet As String, text As String) As Long
' finds the first column header that matches given text in the given sheet
    Dim thissheet As Long
    Dim tmp As Long
    On Error GoTo error

    thissheet = ActiveSheet.Index
    If thissheet <> Sheets(sheet).Index Then
        Sheets(sheet).Select
    End If

    tmp = Cells.find(What:=text, after:=Cells(1, 1), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Column
    Sheets(thissheet).Select
    findFirstCol = tmp
    Exit Function

error:
    MsgBox "Column not found: " & text
End Function

Public Function findFirstRow(sheet As String, text As String) As Long
' finds the first column header that matches given text in the given sheet
    Dim thissheet As Long
    Dim tmp As Long
    On Error GoTo error

    thissheet = ActiveSheet.Index
    If thissheet <> Sheets(sheet).Index Then
        Sheets(sheet).Select
    End If

    tmp = Cells.find(What:=text, after:=Cells(1, 1), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row
    Sheets(thissheet).Select
    findFirstRow = tmp
    Exit Function

error:
    MsgBox "Row not found: " & text
End Function

Function isFound(s As String) As Boolean
' returns whether or not the search text exists on the active sheet
    Dim tmp As Long
    On Error GoTo error

    tmp = Cells.find(What:=s, after:=Cells(1, 1), LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Row
    isFound = True
    Exit Function

error:
    isFound = False
End Function

Function findCumulativeTotalRow(strt As Range, amt As Double, write_total As Boolean) As Long
' from starting cel, sum the amounts in that column until the sum = amt
' if write_total is true, it will write cumulative totals to the right of each cel

    Dim cel As Range
    Dim rng As Range
    Dim sum As Double

    Set rng = Range(strt.Address(0, 0) & ":" & colToLetter(strt.Column) & lastrow(colToLetter(strt.Column)))

    For Each cel In rng
        If cel.value <> error Then
            If cel.value <> xlblank Then
                If IsNumeric(cel.value) Then
                    sum = sum + cel.value
                    If write_total = True Then cel.offset(0, 1).value = sum
                    If sum = amt Then
                        findCumulativeTotalRow = sum
                        Exit Function
                    End If
                End If
            End If
        End If
    Next cel

    findCumulativeTotalRow = 0 ' fail state

End Function
