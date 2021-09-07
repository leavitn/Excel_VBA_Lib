' function library for cell references for excel vba
' dependency libraries:
'   string

Function isReference(formula As String) As Boolean
' returns true if cell formula contains a reference
    If find("=", formula, 0) > 0 Then
    ' Formula established, but does it contain a reference?
        If find("!", formula, 0) > 0 Then
            isReference = True
            Exit Function
        Else
            For i = 1 To Len(formula)
                If isalpha(Right(Left(formula, i), 1)) Then
                    If i = Len(formula) Then
                    ' do nothing
                    Else
                        If IsNumeric(Right(Left(formula, i + 1), 1)) Then
                        'let's assume this is a reference
                            isReference = True
                            Exit Function
                        End If
                    End If
                End If
            Next i
        End If
    End If

    isReference = False
End Function

Function isFormula(f As String) As Boolean
' checks if formula
    If Left(f, 1) = "=" Then
        isFormula = True
    Else
        isFormula = False
    End If
End Function

Function hasParentRef(list As String, delimiter As String) As Boolean
' receives a list of cell references and determines if the list contains parent tab
' references or not
    Dim first As String ' first member of the list '
    first = Split(list, delimiter)(0)
    If word_count(first, "!") > 0 Then
    ' test to see if the list contains parent tab references '
        hasParentRef = True
    Else
        hasParentRef = False
    End If
End Function

Function issequential(ByVal curr As String, ByVal last As String) As Boolean
' returns if two cell addresses are sequential or not'
    Dim parent1 As String
    Dim parent2 As String
    Dim col1 As String
    Dim col2 As String

    If word_count(curr, "!") > 0 Then
        parent1 = Split(curr, "!")(0)
        parent2 = Split(last, "!")(0)
        col1 = not_numbers(Split(curr, "!")(1))
        col2 = not_numbers(Split(last, "!")(1))
        curr = Split(curr, "!")(1)
        last = Split(last, "!")(1)
    Else
        parent1 = "1"
        parent2 = parent1
        col1 = not_numbers(curr)
        col2 = not_numbers(last)
    End If
    issequential = parent1 = parent2 And col1 = col2 And numbers_only(curr) - 1 = numbers_only(last)
End Function

Function getRef(f As String) As String
' gets a formula and returns a reference
    Dim inword As Boolean
    Dim curr As String
    Dim ref As String

    inword = False
    For i = 1 To Len(f)
        curr = Right(Left(f, i), 1)
        If inword Then
            If curr = "]" Then
                Exit For
            Else
                ref = ref & curr
            End If
        Else
            If curr = "[" Then inword = True
        End If
    Next i

    getRef = ref

End Function

Sub removeReferences()
' Removes references to other workbooks in range
    Dim cel As Range
    Dim output As String
    Dim offset As Integer

    Application.DisplayAlerts = False
    offset = InputBox("How many rows to offset?")
    For Each cel In Selection
        output = cel.formula
        While find("[", output, 0) > 0
            output = remove_all_between(output, "'", "]")
        Wend
        output = remove_char(output, "'")
        MsgBox output
        cel.formula = offset_row_ref(offset, output)
    Next cel

End Sub


Function convert_range_to_list(ByVal s As String) As String
' converts a range demarked by a ':' to a list of references. List broken up by commas '
' Example Input: RSM!E121:F123
' Example Output: RSM!E121,RSM!E122,RSM!E123,RSM!F121,RSM!F122,RSM!F123
    Dim ref_sheet As String
    Dim start_cell As String
    Dim tmp As String
    Dim n As Long   ' number of cells per each column '
    Dim m As Long   ' number of columns '
    Dim i As Long
    Dim j As Long
    Dim col As String

    If find("!", s, 0) > 0 Then ' reference found
        ref_sheet = Split(s, "!")(0)                    ' Split reference sheet from cell info '
        s = printAfter(s, find("!", s, 0))
    End If

    start_cell = Split(s, ":")(0)    ' get the start cell '
    n = numbers_only(Split(s, ":")(1)) - numbers_only(start_cell)    ' get the number of cells per col '
    m = char_diff(Left(Split(s, ":")(1), 1), Left(start_cell, 1))    ' get the number of columns '

    For j = 0 To m
        ' for each column '
        For i = 0 To n
            ' add to the list the referenced cells in RSM!E121 format and separated by a comma '
            If Len(ref_sheet) > 0 Then
            ' If there is a reference, add the reference to the cell
                tmp = tmp & ref_sheet & "!"
            End If
            tmp = tmp & int_to_char(char_to_int(Left(start_cell, 1)) + j) & numbers_only(start_cell) + i & ","
        Next i
    Next j
    convert_range_to_list = Left(tmp, Len(tmp) - 1)


End Function

Private Sub highlightDereferencedCells()
' Deference all cell references until a cell with a value is found. Then highlight it.
' LIMITATION: As of now, won't work for cells located in other worksheets

    Dim s As String
    Dim cel As Range
    Dim curr As String
    Dim this_sheet As String

    this_sheet = ActiveSheet.name

    For Each cel In Selection
        If isReference(cel.formula) And find("sum", cel.formula, 1) = 0 Then
            s = s & reflistSuccincttoVerbose(cel.formula) & "," ' comma is added to the end to separate / stitch together each list
        End If
    Next cel

    If Len(s) > 0 Then
    ' if list is populated
        s = Left(s, Len(s) - 1) ' remove the comma at the end of the list
        For i = 0 To word_count(s, ",")
        ' for each cell in the list, go to it and recursively call the function again until the list is length 0
            curr = Split(s, ",")(i)
            If find("!", curr, 0) = 0 Then
            ' If no sheet reference is added cell address, add
                curr = this_sheet & "!" & curr
            End If

            Call select_cell(curr)
            highlightDereferencedCells ' recursively self-call function until cell with a value is found
        Next i
    Else
        ' Cell with only a value has been found, highlight
        Selection.Interior.Color = 65535
    End If

End Sub


Function reflistSuccincttoVerbose(ByVal s As String) As String
' The opposite of function reflistVerbosetoSuccinct
' Gets a cell formula and expands all ranges (A1:A3,A5:A6) of referenced cells into one monster list(A1,A2,A3,A5,A6...etc.)
    Dim tmp As String
    Dim n As Integer    ' number of groups '
    Dim m As Integer    ' number of sums
    Dim output As String
    Dim c As Integer    ' character number for find()
    s = remove_char(s, "=+-()$") ' strip unnecessary characters that have nothing to do with the references
    s = remove_word(s, "SUM") ' strip SUM() formulas to expose the reference ranges
    s = addCommaSeparators(s) ' will only add comma separators if the formula contains more than one range of cells
    If word_count(s, ",") > 0 Then
    ' more than one group - must separate '
        n = word_count(s, ",")
    Else
        n = 0
    End If
    tmp = s
    For i = 0 To n
    ' for each group in sum '
        If n > 0 Then
            tmp = Split(s, ",")(i)
        End If
        If find(":", tmp, 0) <> False Then
            output = output + convert_range_to_list(tmp) + ","
        Else
            output = output + tmp + ","
        End If
    Next i
    output = Left(output, Len(output) - 1)
    reflistSuccincttoVerbose = output ' output is verbose whereas input as succinct
End Function

Function refListVerboseToSuccinct(list As String, delimiter As String) As String
' formats the input (a list of references) as output: =sum(sheet!a1:a3,sheet!a5)
' list = string of cell references separated by delimiter '
    Dim n As Long               ' number of members in the list '
    Dim i As Long               ' current member index'
    Dim m As Long               ' number of members in current contiguous series '
    Dim output As String        ' formatted output '
    Dim curr As String          ' current member '
    Dim last As String          ' last member '
    Dim tmp As String           ' temp '
    Dim hasparent As Boolean    ' does the list have parent tab references? '

    ' check if the list has parent references '
    hasparent = hasParentRef(list, delimiter)

    If Right(list, 1) = delimiter Then
    ' formatting - make sure it's correct '
        list = Left(list, Len(list) - 1)    ' if next char in list is delimiter, truncate '
    End If

    ' Initialization '
    n = word_count(list, delimiter) + 1     ' get number of members in list '
    curr = Split(list, delimiter)(0)
    output = curr
    last = curr
    m = 0

    ' Main Loop '
    For i = 2 To n
        curr = Split(list, delimiter)(i - 1)
        If issequential(curr, last) = True Then
            m = m + 1                       ' increment the number of members in the sequence '
            If i = n Then
            ' if the list ends while in a sequence '
                If hasparent = True Then
                    output = output & ":" & Split(curr, "!")(1)
                Else
                    output = output & ":" & curr
                End If
            End If
        ElseIf m > 0 Then
        ' if issequential is false and m > 0, then write tail of sequence '
            If hasparent = True Then
                output = output & ":" & Split(last, "!")(1) & "," & curr
            Else
                output = output & ":" & last & "," & curr
            End If

            m = 0
        Else
        ' if issequential = false and m = 0, then simply write next cell
            output = output & "," & curr
        End If
        last = curr
    Next i
    refListVerboseToSuccinct = output

End Function

Function refcnt(ByVal f As String) As Integer
' returns a count of the number of references in the formula
    Dim curr As String ' current char
    Dim nxt As String ' next char
    Dim n As Integer ' the count

    While word_count(f, "'") > 0
        f = remove_all_between(f, "'", "'")
    Wend

    For i = 1 To Len(f) - 1
        curr = Right(Left(f, i), 1)
        nxt = Right(Left(f, i + 1), 1)
        If isalpha(curr) And (nxt = "$" Or IsNumeric(nxt)) Then
        ' reference found
            n = n + 1
        End If
    Next i

    refcnt = n

End Function

Function getNthRef(ByVal n As Integer, ByVal f As String) As String
' returns the nth reference in the formula 'f'
    Dim curr As String ' current char
    Dim nxt As String ' next char
    Dim output As String ' output
    Dim inref As Boolean ' in the reference?

    ' remove unnecessary data
    While word_count(f, "'") > 0
        f = remove_all_between(f, "'", "'")
    Wend

    For i = 1 To Len(f)
        curr = Right(Left(f, i), 1)

        If i <> Len(f) Then
            nxt = Right(Left(f, i + 1), 1)
        Else
            nxt = "!" ' to prevent errors
        End If

        If inref And (IsNumeric(curr) Or curr = "$" Or isalpha(curr)) Then
            output = output & curr
        ElseIf inref And Not (IsNumeric(curr) Or curr = "$") Then
            Exit For
        ElseIf Not inref And ((curr = "$" And isalpha(nxt)) Or (isalpha(curr) And IsNumeric(nxt))) Then
            If n > 1 Then
            ' find the nth reference
                n = n - 1
            Else
                ' start writing
                output = curr
                inref = True
            End If
        End If
    Next i

    getNthRef = output

End Function


Function convert_sum_to_ref_list(ByVal s As String) As String
' receives a formula with the following format sum(TAB!x1:xn,TAB!zm) format '
' converts to a list of all the referenced cells  '
    Dim tmp As String
    Dim n As Integer    ' number of groups '
    Dim output As String

    ' formatting '
    s = remove_char(s, "=+-()")
    s = remove_word(s, "SUM")

    If word_count(s, ",") > 0 Then
    ' more than one group - must separate '
        n = word_count(s, ",")
    Else
        n = 0
    End If
    tmp = s
    For i = 0 To n
    ' for each group in sum '
        If n > 0 Then
            tmp = Split(s, ",")(i)
        End If
        If word_count(tmp, ":") > 0 = True Then
            output = output + convert_range_to_list(tmp) + ","
        Else
            output = output + tmp + ","
        End If
    Next i
    output = Left(output, Len(output) - 1)
    convert_sum_to_ref_list = output
End Function


Function create_cell_list(rng As Range) As String
' creates a list (contained in a string) from the range of cells provided '
    Dim cel As Range
    Dim output As String

    For Each cel In rng
        output = output + Selection.Parent.name & "!" & cel.Address(0, 0) & " "
    Next cel
    create_cell_list = output
End Function

Function create_cell_list_if(criteria As String, rng As Range, val_col As String) As String
' returns a list of formatted cell addresses that match the criteria function provided '
' list includes parent references '
' ################################################################################################## '
' Parameters: '
' criteria = name of a boolean FUNCTION that receives the current cel and returns a true/false evaluation on that cell '
' rng = range '
' val_col = column that address will be grabbed from '
' ################################################################################################## '

    Dim cel As Range    ' current cel address '
    Dim tmp As String   ' formatted output '

    For Each cel In rng
        If Application.Run(criteria, cel) = True Then  ' function pointer to function returning conditional expression '
        ' match found - add to list '
           tmp = tmp & rng.Parent.name & "!" & Range(val_col & cel.Row).Address(0, 0) & " "
        End If
    Next
    create_cell_list_if = tmp

End Function

Sub paste_remove_references()
' Pasting from another workbook contains references to that workbook if sheets are referred to
' This removes those references when pasting
    Dim cel As Range

    Selection.PasteSpecial
    For Each cel In Selection
        cel.formula = remRefs(cel.formula)
    End
End Sub

Function rmvRefs(ByVal f As String) As String
' Removes references in the formula received
    For i = 1 To word_count(f, "[")
         f = remove_all_between(f, "'", "]")
         f = rmvFirstChar(f, "'")
    Next i

    rmvRefs = f
End Function

Sub offset_cell_row_reference()
' Offset the cell reference row (for example, if off by one row then +1 or -1 is input
    Dim adjust As Integer
    Dim curr As Long
    adjust = InputBox("How many rows do you want to adjust?")
    Dim cel As Range
    For Each cel In Selection
       cel.formula = offset_row_ref(adjust, cel.formula)
    Next cel

End Sub

Function addCommaSeparators(s As String) As String
' For sum formulas that have been stripped of extraneous characters, this formula adds comma separators between ranges
' see convert_sum_to_references function

    Dim curr As String
    Dim last As String
    Dim output As String

    For i = 1 To Len(s)
        If i = 1 Then
            curr = Right(Left(s, i), 1)
        Else
            last = curr
            curr = Right(Left(s, i), 1)
            If isalpha(curr) And IsNumeric(last) Then
                output = output & ","
            End If
        End If
        output = output & curr
    Next i

    addCommaSeparators = output

End Function


Function offset_row_ref(offset As Integer, output As String) As String
' Offset the cell reference row # (for example, if off by one row then +1 or -1 is input
    Dim adjust As Long
    Dim curr As Long
    Dim cel As Range

    adjust = offset

    For i = 1 To n_integers(output)
        curr = integer_breakout(output, CInt(i))
        output = replace_word(output, CStr(curr), CStr(curr + adjust))
    Next i

    offset_row_ref = output
End Function


Private Sub select_cell(ByVal addr As String)
' receives an address to select, goes to that cell (even if in another tab) '
' and selects it '
' address must be in "Sheet!RC" format '

    Dim sht As String   ' sheet
    Dim cel As String    ' cell address

    sht = Split(addr, "!")(0)
    cel = Split(addr, "!")(1)
    Sheets(sht).Select
    Range(cel).Select

End Sub

Function vlookup_refs_only(value As String, rng As Range, offset As Long) As String
' works like vlookup in excel, returns cell references of matches
    Dim cel As Range
    For Each cel In rng
        If cel.value = value Then
            vlookup_refs_only = cel.Parent.name & "!" & cel.offset(0, offset).Address(0, 0)
            Exit Function
        End If
    Next cel
    ' if cannot be found '
    vlookup_refs_only = "%N/A"
End Function
