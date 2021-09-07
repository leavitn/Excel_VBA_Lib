' library for storing keys and values in a string

' create
Function keyval_put(ByVal keyval As String, key As String, val As String) As String
' adds key and val to keyval, returns transformed keyval
  Dim tmp As String
  If Len(keyval) > 0 Then
    tmp = keyval & "," & key & ":" & val
  Else
    tmp = keyval & key & ":" & val
  End If
  keyval_put = tmp
End Function

' read
Function keyval_get(key As String, data As String) As Double
' receives a list of key value pairs deliminated by ","
' key1:value1,key2:value2
' returns value associated with key argument
' value assumed to be a double
    For i = 0 To word_count(data, ",")
        keyval = Split(data, ",")(i)
        If key = Split(keyval, ":")(0) Then
            getval = CDbl(Split(keyval, ":")(1))
            Exit Function
        End If
    Next i

    MsgBox "Error: Key could not be found!" & vbNewLine & _
        "key: " & key
    getval = "error"
End Function

' update
Function keyval_update(ByVal data As String, key As String, new_val As String) As String
' updates an existing key in a keyval
    Dim tmp As String
    Dim keyval As String
    Dim count As Long
    Dim found As Boolean
    found = False
    tmp = data

    If keyval_key_exists(data, key) Then
        tmp = keyval_delete(tmp, key)
    End If
    tmp = keyval_put(tmp, key, new_val)

    keyval_update = tmp
End Function

' delete
Function keyval_delete(ByVal data As String, key As String) As String
' deletes a keyval given a keyval string and a key
    Dim tmp As String

    For i = 0 To word_count(data, ",")
        keyval = Split(data, ",")(i)
        If key = Split(keyval, ":")(0) Then
            ' do nothing
        Else
            If Len(tmp) > 0 Then
                tmp = tmp & "," & keyval
            Else
                tmp = keyval
            End If
        End If
    Next i
    keyval_delete = tmp
End Function

Function keyval_key_exists(ByVal data As String, key As String) As Boolean
' returns whether the key already exists in the keyval
    Dim i As Long
    Dim count As Long

    If Len(data) > 0 Then
        For i = 0 To word_count(data, ",")
            keyval = Split(data, ",")(i)
            If key = Split(keyval, ":")(0) Then
                keyval_key_exists = True
                Exit Function
            End If
        Next i
    End If

    keyval_key_exists = False
End Function
