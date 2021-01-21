Sub RunTotals()

'Variables
Dim dates() As Date
Dim totals() As Double
Dim time() As Double
Dim isNew As Boolean
isNew = False
Dim i As Integer
i = 2
Dim curRange As String
curRange = "E" + CStr(i)
Dim curDate As String
Dim curRadRange As String
curRadRange = "K" + CStr(i)
Dim curRad As String
Dim curTimeRange As String
curTimeRange = "L" + CStr(i)
Dim curTime As Double


'Sort the sheet by Date (column E)
Dim numDates As Long
numDates = Cells(Rows.Count, "E").End(xlUp).Row - 1
Dim sortRange As String
sortRange = "A2:N" + CStr(numDates)
'DEBUGGING - MsgBox ("sortRange is " + sortRange)
Range(sortRange).Sort key1:=Range("E:E"), order1:=xlAscending, Header:=xlNo

' While theres a value in the E column, walk the dates() array check for a match
Do While Not Range(curRange).Value = ""
    isNew = True

    'Check for first run
    If i = 2 Then
        curDate = Range(curRange).Value
        curRad = Range(curRadRange).Value
        curTime = Range(curTimeRange).Value
        ReDim dates(1)
        dates(0) = curDate
        ReDim totals(1)
        totals(0) = curRad
        ReDim time(1)
        time(0) = curTime
        i = i + 1
        curRange = "E" + CStr(i)
        curRadRange = "K" + CStr(i)
        curTimeRange = "L" + CStr(i)
    Else

        'Not first run. Grab the values of columns of interest
        curDate = Range(curRange).Value
        curRad = Range(curRadRange).Value
        curTime = Range(curTimeRange).Value

        'Check the dates() array for a match
        Dim j As Integer
        j = 0
        Do While j < UBound(dates)
            'If we find a match, add the value in the K column to the total
            If (curDate = dates(j)) Then
                totals(j) = totals(j) + curRad
                time(j) = time(j) + curTime
                isNew = False
            End If
            j = j + 1
        Loop

        'If the date is new, increase the size of each array and add the values
        If (isNew) Then
            ReDim Preserve dates(UBound(dates) + 1)
            dates(UBound(dates) - 1) = curDate
            ReDim Preserve totals(UBound(totals) + 1)
            totals(UBound(totals) - 1) = curRad
            ReDim Preserve time(UBound(time) + 1)
            time(UBound(time) - 1) = curTime
        End If
        i = i + 1
        curRange = "E" + CStr(i)
        curRadRange = "K" + CStr(i)
        curTimeRange = "L" + CStr(i)
    End If
Loop
Dim msg As String
Dim k As Integer

'Make a header for the arrays
Range("E1").Copy Range("S1")
Range("K1").Copy Range("T1")
Range("L1").Copy Range("U1")

'Put the values in the columns
Dim l As Integer
l = 0
Dim m As Integer
m = 2
While l < UBound(dates)
    Cells(m, 19).Value = CStr(dates(l))
    Cells(m, 20).Value = CStr(totals(l) / 1000)
    Cells(m, 21).Value = CStr(time(l))
    l = l + 1
    m = m + 1
Wend
End Sub
