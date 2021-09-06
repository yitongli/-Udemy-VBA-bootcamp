Sub Main_Process()

Dim folder_path As String, file_extension As String, input_file As String
Dim wbData As Workbook, shtdata As Worksheet
Dim shop_number As Integer


Call InitializeVariables

folder_path = "C:\Users\lyt\Documents\Excel_VBA\Input Files\"
file_extension = "*.xlsx"
input_file = Dir(folder_path & file_extension)

Do While input_file <> ""

    Set wbData = Workbooks.Open(Filename:=folder_path & input_file)
    Set shtdata = wbData.Sheets("Data")

    shop_number = shtdata.Cells(2, 3)

    Select Case shop_number
        Case 1
          For i = 2 To shtdata.Range("A1048576").End(xlUp).Row
            shtSH1.Cells(row_sh1, 3).Value = shtdata.Cells(i, 1)
            shtSH1.Cells(row_sh1, 4).Value = shtdata.Cells(i, 2)
            shtSH1.Cells(row_sh1, 5).Value = shtdata.Cells(i, 4)
            shtSH1.Cells(row_sh1, 6).Value = shtdata.Cells(i, 5)
            shtSH1.Cells(row_sh1, 7).Value = shtdata.Cells(i, 6)
            shtSH1.Cells(row_sh1, 8).Value = shtdata.Cells(i, 8)
            shtSH1.Cells(row_sh1, 9).Value = shtdata.Cells(i, 9)
            row_sh1 = row_sh1 + 1
          Next

        Case 2
          For i = 2 To shtdata.Range("A1048576").End(xlUp).Row
            shtSH2.Cells(row_sh2, 3).Value = shtdata.Cells(i, 1)
            shtSH2.Cells(row_sh2, 4).Value = shtdata.Cells(i, 2)
            shtSH2.Cells(row_sh2, 5).Value = shtdata.Cells(i, 4)
            shtSH2.Cells(row_sh2, 6).Value = shtdata.Cells(i, 5)
            shtSH2.Cells(row_sh2, 7).Value = shtdata.Cells(i, 6)
            shtSH2.Cells(row_sh2, 8).Value = shtdata.Cells(i, 8)
            shtSH2.Cells(row_sh2, 9).Value = shtdata.Cells(i, 9)
            row_sh2 = row_sh2 + 1
          Next
        Case 3
          For i = 2 To shtdata.Range("A1048576").End(xlUp).Row
            shtSH3.Cells(row_sh3, 3).Value = shtdata.Cells(i, 1)
            shtSH3.Cells(row_sh3, 4).Value = shtdata.Cells(i, 2)
            shtSH3.Cells(row_sh3, 5).Value = shtdata.Cells(i, 4)
            shtSH3.Cells(row_sh3, 6).Value = shtdata.Cells(i, 5)
            shtSH3.Cells(row_sh3, 7).Value = shtdata.Cells(i, 6)
            shtSH3.Cells(row_sh3, 8).Value = shtdata.Cells(i, 8)
            shtSH3.Cells(row_sh3, 9).Value = shtdata.Cells(i, 9)
            row_sh3 = row_sh3 + 1
          Next
        Case 4
          For i = 2 To shtdata.Range("A1048576").End(xlUp).Row
            shtSH4.Cells(row_sh4, 3).Value = shtdata.Cells(i, 1)
            shtSH4.Cells(row_sh4, 4).Value = shtdata.Cells(i, 2)
            shtSH4.Cells(row_sh4, 5).Value = shtdata.Cells(i, 4)
            shtSH4.Cells(row_sh4, 6).Value = shtdata.Cells(i, 5)
            shtSH4.Cells(row_sh4, 7).Value = shtdata.Cells(i, 6)
            shtSH4.Cells(row_sh4, 8).Value = shtdata.Cells(i, 8)
            shtSH4.Cells(row_sh4, 9).Value = shtdata.Cells(i, 9)
            row_sh4 = row_sh4 + 1
          Next

        Case 5
          For i = 2 To shtdata.Range("A1048576").End(xlUp).Row
            shtSH5.Cells(row_sh5, 3).Value = shtdata.Cells(i, 1)
            shtSH5.Cells(row_sh5, 4).Value = shtdata.Cells(i, 2)
            shtSH5.Cells(row_sh5, 5).Value = shtdata.Cells(i, 4)
            shtSH5.Cells(row_sh5, 6).Value = shtdata.Cells(i, 5)
            shtSH5.Cells(row_sh5, 7).Value = shtdata.Cells(i, 6)
            shtSH5.Cells(row_sh5, 8).Value = shtdata.Cells(i, 8)
            shtSH5.Cells(row_sh5, 9).Value = shtdata.Cells(i, 9)
            row_sh5 = row_sh5 + 1
          Next

    End Select
    
    wbData.Close
    input_file = Dir 'this will give the loop the other file

Loop

End Sub

