' Date Library

Function monthShort(mo As Integer) As String
' receives a month as an integer and returns the month as a string in short form '
    Dim m(12) As String
    m(0) = "Jan"
    m(1) = "Feb"
    m(2) = "Mar"
    m(3) = "Apr"
    m(4) = "May"
    m(5) = "June"
    m(6) = "July"
    m(7) = "Aug"
    m(8) = "Sep"
    m(9) = "Oct"
    m(10) = "Nov"
    m(11) = "Dec"
    monthShort = m(mo - 1)
End Function

Function monthLong(mo As Integer) As String
' receives a month as an integer and returns the month as a string in short form '
    Dim m(12) As String
    m(0) = "January"
    m(1) = "February"
    m(2) = "March"
    m(3) = "April"
    m(4) = "May"
    m(5) = "June"
    m(6) = "July"
    m(7) = "August"
    m(8) = "September"
    m(9) = "October"
    m(10) = "November"
    m(11) = "December"
    monthLong = m(mo - 1)
End Function

Function monthStringtoInt(s As String) As Integer
' receives a month in long form as a string and converts to int
    Dim tmp As Integer
    Select Case s:
        Case "January": tmp = 1
        Case "February": tmp = 2
        Case "March": tmp = 3
        Case "April": tmp = 4
        Case "May": tmp = 5
        Case "June": tmp = 6
        Case "July": tmp = 7
        Case "August": tmp = 8
        Case "September": tmp = 9
        Case "October": tmp = 10
        Case "November": tmp = 11
        Case "December": tmp = 12
    End Select
    monthStringtoInt = tmp
End Function

Function lastDayofMonth(mo As Integer) As String
' Returns last day of the month '
    Dim m(12) As Integer
    m(0) = 31
    m(1) = 28
    m(2) = 31
    m(3) = 30
    m(4) = 31
    m(5) = 30
    m(6) = 31
    m(7) = 31
    m(8) = 30
    m(9) = 31
    m(10) = 30
    m(11) = 31

    If Year(Date) Mod 4 = 0 Then
    ' if leap year '
        m(1) = m(1) + 1
    End If

    lastDayofMonth = m(mo - 1)
End Function
