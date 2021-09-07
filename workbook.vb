' workbook library

Function activatewb(wbName As String) As Boolean
' activates an open workbook given a name

    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If wb.name Like wbName & "*" Then
            wb.Activate
            activatewb = True
            Exit Function
        End If
    Next wb
    ' wb not found
    activatewb = False
End Function

Function openwb(path As String, file As String) As Boolean
' opens specified workbook
    Dim filename As String
    filename = Dir(path & file)
    If Len(filename) = 0 Then
        'failure'
        MsgBox "Error: Filepath is invalid"
        openwb = False
        Exit Function
    Else
        'success'
        Workbooks.Open filename:=path & filename
        openwb = True
    End If
End Function


Function openwb_path_only(ByVal path As String) As Boolean
' opens specified workbook
    Dim filename As String
    filename = Dir(path)
    If Len(filename) = 0 Then
        'failure'
        MsgBox "Error: Filepath is invalid"
        openwb_path_only = False
        Exit Function
    Else
        'success'
        Workbooks.Open filename:=path
        openwb_path_only = True
    End If
