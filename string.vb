' String Library for Excel VBA

Function isalpha(c As String) As Boolean
' receives a character and returns whether it is a letter or not '
    Const alpha = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim i As Integer

    i = 0
    While i < Len(alpha) And Right(Left(alpha, i), 1) <> c
        i = i + 1
    Wend
    If i < Len(alpha) Then
        isalpha = True
    Else
        isalpha = False
    End If

End Function

Function is_integer(s As String) As Boolean
' Tests whether the value is an integer
    Const integers As String = "0123456789"
    Dim check As Boolean
    check = False
    For j = 1 To Len(s)
        For i = 1 To Len(integers)
            If Right(Left(integers, i), 1) = Right(Left(s, j), 1) Then
                check = True
            End If
        Next i
    Next j
    is_integer = check
End Function

Function doesNotContain(s As String, chars As String) As Boolean
' returns if s string contains the supplied chars
    For i = 1 To Len(s)
        For j = 1 To Len(chars)
            If Right(Left(s, i), 1) = Right(Left(chars, j), 1) Then
                ' contains, exit
                doesNotContain = False
                Exit Function
            End If
        Next j
    Next i

    doesNotContain = True
End Function

Function has_alpha(val As String) As Boolean
' returns whether an alpha character is detected in value
  Const alpha = "-ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
  Dim i As Long
  Dim char As String
  Dim found As Boolean

  found = False
  For i = 1 To Len(val)
    char = Right(Left(val, i), 1)
    For j = 1 To Len(alpha)
      If char = Right(Left(alpha, j), 1) Then
        found = True
        Exit For
      End If
    Next j
    If found = True Then
      Exit For
    End If
 Next i
 has_alpha = found
End Function

Function find(find_text As String, within_text As String, start_num As Integer) As Integer
' returns first position of find_text within the provided text
' start_num needs to be > 0
    Dim inword As Boolean
    find_text = upper(find_text)
    within_text = upper(within_text)
    For i = start_num To Len(within_text)
        If Right(Left(within_text, i), 1) = Left(find_text, 1) Then
        ' First letter found - check to see if the rest of the word matches
            For j = 1 To Len(find_text)
                If Right(Left(within_text, i + j - 1), 1) <> Right(Left(find_text, j), 1) Then
                    Exit For
                ElseIf Right(Left(within_text, i + j - 1), 1) = Right(Left(find_text, j), 1) And j = Len(find_text) Then
                    find = i
                    Exit Function
                End If
            Next j
        End If
    Next i
    ' not found '
    find = 0 ' error '
End Function

Function word_count(s As String, word As String) As Integer
' returns the # of instances of word that occurs in s '
    Dim cnt As Integer
    Dim inword As Boolean
    Dim char As String
    Dim i As Long ' index in s
    Dim j As Long ' index in comparison word
    i = 1
    cnt = 0
    While i <= Len(s)
        char = Right(Left(s, i), 1)
        If inword = True Then
            If char <> Right(Left(word, j), 1) Then
                ' not a match '
                inword = False
            ElseIf char = Right(Left(word, j), 1) And j = Len(word) Then
                ' if it reaches this point, match was found. Count incremented. '
                cnt = cnt + 1
                inword = False
            Else
                j = j + 1
            End If
        ' if out of word and current character = start of word, then inword = true '
        ElseIf char = Left(word, 1) Then
            If Len(word) = 1 Then
            ' exception for if word is only 1 character in length '
                cnt = cnt + 1
            Else
              inword = True
              j = 2
            End If
        End If
        i = i + 1
    Wend
    word_count = cnt
End Function

Function remove_word(s As String, word As String) As String
' Outputs a string that will be missing the specified instances of word supplied to function '

    Dim i As Long
    Dim j As Long
    Dim k As Long

    Dim tmp As String

    For i = 1 To Len(s)
        j = 1
        k = i   ' temporary variable for 'looking ahead' of i '
        If Right(Left(s, i), 1) = Left(word, 1) Then
            While Right(Left(s, k), 1) = Right(Left(word, j), 1) And j <= Len(word)
                j = j + 1
                k = k + 1
            Wend
            If j = Len(word) + 1 Then
                ' match found - advance past word'
                i = k
            End If
        End If
        If i <= Len(s) Then
        ' fixes bug: if k advances past the length of string, then do nothing '
            tmp = tmp + Right(Left(s, i), 1)
        End If
    Next i

    remove_word = tmp
End Function

Function replace_word(s As String, word As String, replacement As String) As String
' Outputs a string that will replace all instances of word with replacement'

    Dim i As Long
    Dim j As Long
    Dim k As Long

    Dim tmp As String

    For i = 1 To Len(s)
        j = 1
        k = i   ' temporary variable for 'looking ahead' of i '
        If Right(Left(s, i), 1) = Left(word, 1) Then
            While Right(Left(s, k), 1) = Right(Left(word, j), 1) And j <= Len(word)
                j = j + 1
                k = k + 1
            Wend
            If j = Len(word) + 1 Then
                ' match found - advance past word'
                i = k
                tmp = tmp & replacement
            End If
        End If
        If i <= Len(s) Then
        ' fixes bug: if k advances past the length of string, then do nothing '
            tmp = tmp + Right(Left(s, i), 1)
        End If
    Next i

    replace_word = tmp
End Function


Function word_count_range(rng As Range, word As String) As Long
' Does a word count for a range '
    Dim cel As Range
    Dim cnt As Long
    Dim tmp As Long

    For Each cel In rng
        tmp = word_count(cel.value, word)
        If tmp > 0 Then
            cnt = cnt + tmp
        End If
    Next cel
    word_count_range = cnt

End Function


Function not_numbers(ByVal s As String) As String
' receives a string and returns everhthing BUT the numbers in that string '
    Dim i As Long
    Dim tmp As String
    For i = 1 To Len(s)
        If IsNumeric(Right(Left(s, i), 1)) Then
            ' do nothing
        Else
            tmp = tmp & Right(Left(s, i), 1)
        End If
    Next i

    not_numbers = tmp
End Function


Function numbers_only(ByVal s As String) As Long
' receives a string and returns ONLY the numbers in that string '
    Dim i As Long
    Dim tmp As String
    For i = 1 To Len(s)
        If IsNumeric(Right(Left(s, i), 1)) Then
            tmp = tmp & Right(Left(s, i), 1)
        End If
    Next i

    numbers_only = tmp
End Function

Function remove_char(s As String, c As String) As String
' Outputs a string that will be missing the specified character or characters '
    Dim i As Long
    Dim tmp As String
    Dim copy As Boolean

    For i = 1 To Len(s)
        copy = True
        For j = 1 To Len(c)
            If Right(Left(s, i), 1) = Right(Left(c, j), 1) Then
                copy = False
            End If
        Next j
        If copy = True Then
           tmp = tmp + Right(Left(s, i), 1)
        End If
    Next i

    remove_char = tmp
End Function

Function nextAlpha(c As String)
' get a character and return the next character in the alphabet
    nextAlpha = int_to_char(1 + char_to_int(c))
End Function

Function char_diff(ByVal c1 As String, ByVal c2 As String)
' given two characters, outputs the difference via character arithmetic '
    Const alpha = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim i As Integer
    Dim j As Integer

    i = 1
    While Right(Left(alpha, i), 1) <> c1
        i = i + 1
    Wend
    j = 1
    While Right(Left(alpha, j), 1) <> c2
        j = j + 1
    Wend
    char_diff = i - j

End Function

Function alphaOnly(s As String) As String
' returns ONLY alpha characters
    Dim output As String
    Dim curr As String

    For i = 1 To Len(s)
        curr = Right(Left(s, i), 1)
        If isalpha(curr) Then
            output = output & curr
        End If
    Next i

    alphaOnly = output

End Function

Function int_to_char(i As Integer) As String
' receives an integer and converts to corresponding character '
    Const alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    int_to_char = Right(Left(alpha, i), 1)
End Function

Function char_to_int(c As String) As Integer
' receives a character and returns a corresponding integer '
    Const alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim i As Integer
    i = 1
    While Right(Left(alpha, i), 1) <> c
        i = i + 1
    Wend

    char_to_int = i
End Function

Function upper(s As String) As String
'Converts lower case to upper case '
    Const alpha = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim output As String
    Dim i As Long
    Dim j As Long

    For i = 1 To Len(s)
        j = 1
        While j <= Len(alpha) And Right(Left(alpha, j), 1) <> Right(Left(s, i), 1)
            ' find the
            j = j + 1
        Wend
        If j <= 26 Then
            output = output & Right(Left(alpha, j + 26), 1)
        Else
            output = output & Right(Left(s, i), 1)
        End If
    Next i
    upper = output
End Function

Function remove_all_between(ByVal s As String, first As String, last As String) As String
' given a string and a right and left symbol, such as braces [], all content will be removed between symbols

    Dim i As Long
    Dim output As String
    Dim length As Long
    Dim inside As Boolean

    length = Len(s)
    inside = False
    For i = 1 To length
        If Right(Left(s, i), 1) = first And inside = False Then
            inside = True
            While Right(Left(s, i + 1), 1) <> last And i + 1 <= length
                i = i + 1
            Wend
            i = i + 2
        End If
        output = output & Right(Left(s, i), 1)
    Next i
    remove_all_between = output
End Function

Function onlyBeforeDelimiter(s As String, delim As String) As String
' Returns a string that has only the characters before the delimiter is reached
    Dim output As String
    Dim c As String

    For i = 1 To Len(s)
        c = Right(Left(s, i), 1)
        If c <> delim Then
            output = output & c
        Else
            i = Len(s)
        End If
    Next i
    onlyBeforeDelimiter = output
End Function

Function rmvFirstChar(s As String, c As String) As String
' remove the first char
    Dim output As String
    Dim curr As String

    For i = 1 To Len(s)
        curr = Right(Left(s, i), 1)
        If found = False And curr = c Then
            found = True
        Else
            output = output & curr
        End If
    Next i

    rmvFirstChar = output

End Function

Function integer_breakout(s As String, n As Integer) As Integer
' Returns the nth integer found in the string received

    Dim in_number As Boolean
    Dim curr As String
    Dim output As Long

    in_number = False
    For i = 1 To Len(s)
        curr = CStr(Right(Left(s, i), 1))

        If is_integer(curr) = True And in_number = False And n > 1 Then
        ' Integer found, but this is an integer found before the one we want
            in_number = True
        ElseIf is_integer(curr) = True And in_number = False And n = 1 Then
        ' Integer found and it's the one we want
            in_number = True
            output = output & curr
        ElseIf is_integer(curr) = True And in_number = True And n = 1 Then
        ' In a number, and n = 1
            output = output & curr
        ElseIf is_integer(curr) = False And in_number = True And n = 1 Then
        ' Integer found and it's the one we want and we've found the whole thing
            integer_breakout = output
            Exit Function
        ElseIf is_integer(curr) = False And in_number = True And n > 1 Then
        ' No longer in integer
            n = n - 1
            in_number = False
        End If
    Next i
    integer_breakout = output
End Function

Function n_integers(s As String) As Long
' Returns the number of integers found in received string
    Dim curr As String
    Dim in_integer As Boolean
    Dim n As Long
    Dim i As Long

    n = 0

    For i = 1 To Len(s)
        curr = Right(Left(s, i), 1)
        If is_integer(curr) = True And in_integer = False Then
            in_integer = True
            n = n + 1
        ElseIf is_integer(curr) = False And in_integer = True Then
            in_integer = False
        End If
    Next i

    n_integers = n
End Function

Function printAfter(s As String, pos As Integer) As String
' print at given position for the rest of the string
    For i = pos To Len(s)
        output = output & Right(Left(s, i), 1)
    Next i
    printAfter = output
End Function

Function printAfterNth(find_text As String, ByVal within_text As String, n As Integer) As String
' prints after nth find_text in within_text
    Dim tmp As String
    Dim i As Integer
    Dim s As String

    If n > 0 Then
        i = find(find_text, within_text, 1) ' find index of find_text in within_text
        s = printAfter(within_text, i + Len(find_text)) ' new string, prints all chars after index
        tmp = printAfterNth(find_text, s, n - 1) ' call function again with new string to either (1) find next instance or (2) return result
    Else
        tmp = within_text
    End If
    printAfterNth = tmp

End Function

Function findNth(find_text As String, within_text As String, n As Integer) As String
' returns start position of nth find_text within the provided text
    If word_count(within_text, find_text) < n Then
        findNth = 0
    Else

    findNth = find(printAfterNth(find_text, within_text, n), within_text, 1)
End Function

Function slice(s As String, start_pos As Integer, end_pos As Integer) As String
' slices a string at the start_pos and end_pos
    end_pos = end_pos - 1
    If end_pos - 1 > Len(s) Then
        end_pos = Len(s)
    End If

    For i = start_pos To end_pos
        output = output & Right(Left(s, i), 1)
    Next i
    slice = output
End Function

Function nWords(s As String) As Integer
' returns the number of words in the string
' a word is defined as a string containing more than 3 letters
    Dim spaces As Integer
    Dim cnt As Integer

    If Len(s) > 3 Then
        spaces = nWhiteSpace(s)
        If spaces > 1 Then
            cnt = spaces + 1
        ElseIf spaces = 1 Then
            cnt = spaces
        Else
            cnt = 0
        End If
    Else
        cnt = 0
    End If

    nWords = cnt

End Function

Function nWhiteSpace(s As String) As Integer
' returns the number of white space characters in a string
    Dim cnt As Integer

    cnt = 0

    For i = 1 To Len(s)
        If Right(Left(s, i), 1) = " " Then
            cnt = cnt + 1
        End If
    Next i
    nWhiteSpace = cnt
End Function

Function getNthWord(s As String, n As Integer) As String
' returns the first word of a string
    Dim tmp As String

    If find(" ", s, 1) > 0 And nWords(s) >= n Then
        tmp = Split(s, " ")(n - 1)
    ElseIf nWords(s) < n Then
        MsgBox "Error in getNthWord: n exceeds word count of arg s"
    Else
        tmp = s
    End If

    getNthWord = upper(tmp)
End Function


Function remove_duplicates(s As String, delimiter As String) As String
' Removes duplicates from a string of list members separated by delimiter
' e.g. "first,second,third" or "1,2,3"
' and returns the list without the duplicates
    Dim i As Long
    Dim output As String
    Dim count As String
    Dim sel As String

    count = word_count(s, delimiter)
    For i = 0 To count - 1
        sel = Split(s, delimiter)(i)
        If word_count(output, sel) > 0 Then
            ' do nothing, already in list
        Else
            output = output & sel & ","
        End If
    Next i
    remove_duplicates = output

End Function
