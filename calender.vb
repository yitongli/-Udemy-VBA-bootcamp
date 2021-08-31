Sub loop_Practise()

my_date = CDate("30/08/2021")

For i = 1 To 10
    Sheets("Sheet1").Cells(i, 1) = my_date
    Sheets("Sheet1").Cells(i, 2) = Format(my_date, "dddd")


Select Case Weekday(my_date, vbMonday)
Case 1, 2, 3, 4, 5
    Sheets("Sheet1").Cells(i, 3) = "It's a weekday"
Case Else
    Sheets("Sheet1").Cells(i, 3) = "It's a weekend!"
End Select

my_date = my_date + 1

Next

End Sub
