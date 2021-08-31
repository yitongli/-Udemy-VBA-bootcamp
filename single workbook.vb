Sub Main_Process()

Dim wbMain as workbook, wbData as workbook
Dim shtmain as worksheet, shtdata as worksheet
dim folder_path as string, file_name as string, full_path as string

Set wbMain = ActiveWorkbook
with wbMain
    set shtmain = .sheets("Main")
end with

folder_path = "C:\Users\lyt\Documents\Excel_VBA"
file_name = "Sales_Input.xlsx"
full_path = folder_path & file_name

Set wbData = Workbooks.open(filename:=full_path)
with wbData
    set shtdata = .sheets("data")
end with

shtmain.Range("C12:K166").Value = shtdata.range("A2:I156").Value

wbData.Close

End Sub
