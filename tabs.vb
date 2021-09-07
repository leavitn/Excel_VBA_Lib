' function library for tabs

Function doesTabExist(name As String) As Boolean
' returns whether a tab exists or not
    For i = 1 To Sheets.count
        If Sheets(i).name = name Then
            doesTabExist = True
            Exit Function
        End If
    Next i

    ' fail state
    doesTabExist = False
End Function

Function find_tab_index(keyword As String) As Long
' finds a sheet based on the keyword
   Dim tmp As Long

    For i = 1 To Sheets.count
        If find(keyword, Sheets(i).name, 1) > 0 Then
            find_tab_index = Sheets(i).Index
            Exit Function
        End If
    Next i
    ' error - could not find
    MsgBox "Error! Sheet could not be found."
    find_tab_index = 0
End Function
