Sub Main_Process()

Dim folder_path As String, file_extension As String, input_file As String
Dim wbData As Workbook, shtdata As Worksheet
Dim shop_number As Integer


Call InitializeVariables()

folder_path = "C:\Users\lyt\Documents\Excel_VBA\Input Files"
file_extension = "*.xlsx"
input_file = Dir(folder_path & file_extension)

Do While input_file <> ""

    Set wbData = Workbooks.Open(Filename:=folder_path & file_extension)
    Set shtdata = wbData.Sheets("Data")

    shop_number = shtdata.Cells(2,3)

    Select Case shop_number
        Case 1
          Call Loop_transfer(shtSH1,shtdata,row_sh1)
        Case 2
          Call Loop_transfer(shtSH2,shtdata,row_sh2)
        Case 3
          Call Loop_transfer(shtSH3,shtdata,row_sh3)
        Case 4
          Call Loop_transfer(shtSH4,shtdata,row_sh4)
        Case 5
          Call Loop_transfer(shtSH5,shtdata,row_sh5)

    End Select
    
    wbData.Close
    input_file = Dir

Loop

End Sub

Public Sub Loop_transfer(input_sheet as Worksheet, output_sheet as Worksheet, output_row as Integer)
    for i = 2 to output_sheet.Range("A1048576").End(xlup).Row
            input_sheet.Cells(output_row,3).Value = output_sheet.Cells(i,1)
            input_sheet.Cells(output_row,4).Value = output_sheet.Cells(i,2)
            input_sheet.Cells(output_row,5).Value = output_sheet.Cells(i,4)
            input_sheet.Cells(output_row,6).Value = output_sheet.Cells(i,5)
            input_sheet.Cells(output_row,7).Value = output_sheet.Cells(i,6)
            input_sheet.Cells(output_row,8).Value = output_sheet.Cells(i,8)
            input_sheet.Cells(output_row,9).Value = output_sheet.Cells(i,9)
            output_row = output_row + 1
    next
End Sub
