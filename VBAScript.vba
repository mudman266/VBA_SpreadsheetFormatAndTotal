Sub RunTotals()

'Unmerge all cells
ActiveSheet.Cells.UnMerge

'Check row 9 for any blank headers and remove the relative column
Dim colCheck As Integer
colCheck = 1
Dim littleLoopCounter As Integer 'Keeps the removal process from infinite looping when it reaches the last column early.
littleLoopCounter = 0

While colCheck < 27
    If IsEmpty(Cells(9, colCheck).Value) Then
' DEBUGGING        MsgBox ("Removing Column " + CStr(colCheck) + " it contained: " + CStr(Cells(9, colCheck).Value))
        Columns(colCheck).EntireColumn.Delete
        If littleLoopCounter > 3 Then
            littleLoopCounter = 0
            colCheck = colCheck + 1
        Else
            littleLoopCounter = littleLoopCounter + 1
        End If
    Else
        colCheck = colCheck + 1
    End If
Wend

'Variables
Dim dates() As String
Dim totals() As Double
Dim time() As Double
Dim isNew As Boolean
isNew = False
Dim i As Integer
i = 10
Dim curRange As String
curRange = "D" + CStr(i)
Dim curDate As String
Dim curRadRange As String
curRadRange = "J" + CStr(i)
Dim curRad As String
Dim curTimeRange As String
curTimeRange = "K" + CStr(i)
Dim curTime As Double


'Sort the sheet by Date (column D)
Dim numDates As Long
numDates = Cells(Rows.Count, "D").End(xlUp).Row - 1
Dim sortRange As String
sortRange = "A10:N" + CStr(numDates)
'DEBUGGING - MsgBox ("sortRange is " + sortRange)
Range(sortRange).Sort key1:=Range("D:D"), order1:=xlAscending, Header:=xlNo

' While theres a value in the E column, walk the dates() array check for a match
Do While Not Range(curRange).Value = ""
    isNew = True

    'Check for first run
    If i = 10 Then
        curDate = Range(curRange).Value
        curRad = Range(curRadRange).Value
        curTime = Range(curTimeRange).Value
        ReDim dates(0)
        dates(0) = curDate
        ReDim totals(0)
        totals(0) = curRad
        ReDim time(0)
        time(0) = curTime
        i = i + 1
        curRange = "D" + CStr(i)
        curRadRange = "J" + CStr(i)
        curTimeRange = "K" + CStr(i)
    Else

        'Not first run. Grab the values of columns of interest
        curDate = Range(curRange).Value
        curRad = Range(curRadRange).Value
        curTime = Range(curTimeRange).Value

        'Check the dates() array for a match
        Dim j As Integer
        j = 0
        Do While j <= UBound(dates)
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
            'First add blank data for dates that are not found
            Dim diff As Integer
            diff = DateDiff("d", DateValue(dates(UBound(dates))), DateValue(curDate))
            If (diff > 1) Then
                'MsgBox ("Newest date in list is " + CStr(dates(UBound(dates))) + " | Going to add: " + CStr(DateValue(curDate)))
                'MsgBox ("Adding " + CStr(diff) + " blank dates first.")
                For counter = 0 To (diff - 2)
                    ' Build the blank date
                    Dim blankDate As Date
                    blankDate = DateValue(dates(UBound(dates)))
                    Dim addDate As Date
                    addDate = DateAdd("d", 1, blankDate)
                    ' Add the date and zeros to the array
                    ReDim Preserve dates(UBound(dates) + 1)
                    dates(UBound(dates)) = CStr(addDate)
                    ReDim Preserve totals(UBound(totals) + 1)
                    totals(UBound(totals)) = 0
                    ReDim Preserve time(UBound(time) + 1)
                    time(UBound(time)) = 0
                Next
            End If
            ' If the difference is less than two days, add as normal.
            ReDim Preserve dates(UBound(dates) + 1)
            dates(UBound(dates)) = curDate
            ReDim Preserve totals(UBound(totals) + 1)
            totals(UBound(totals)) = curRad
            ReDim Preserve time(UBound(time) + 1)
            time(UBound(time)) = curTime
        End If
        i = i + 1
        curRange = "D" + CStr(i)
        curRadRange = "J" + CStr(i)
        curTimeRange = "K" + CStr(i)
    End If
Loop
Dim msg As String
Dim k As Integer

'Clear N:R to prevent duplicating data during multiple macro runs
Columns("N:R").EntireColumn.Delete

'Make a header for the arrays
Range("D9").Copy Range("P9")
Range("J9").Copy Range("R9")
Range("K9").Copy Range("Q9")

'Put the values in the columns
Dim l As Integer
l = 0
Dim m As Integer
m = 10
While l < UBound(dates)
    Cells(m, 16).Value = CStr(dates(l))
    Cells(m, 18).Value = CStr(totals(l))
    Cells(m, 17).Value = CStr(time(l))
    l = l + 1
    m = m + 1
Wend
Columns("A:M").AutoFit
Columns("P:R").AutoFit
End Sub
