Public wbMain As Workbook
Public shtSH1 As Worksheet, shtSH2 As Worksheet, shtSH3 As Worksheet, shtSH4 As Worksheet, shtSH5 As Worksheet
Public row_sh1 As Integer, row_sh2 As Integer, row_sh3 As Integer, row_sh4 As Integer, row_sh5 As Integer

Public Sub InitializeVariables()

Set wbMain = ActiveWorkbook

With wbMain
    Set shtSH1 = .Sheets("Shop 1")
    Set shtSH2 = .Sheets("Shop 2")
    Set shtSH3 = .Sheets("Shop 3")
    Set shtSH4 = .Sheets("Shop 4")
    Set shtSH5 = .Sheets("Shop 5")
End With

row_sh1 = 12
row_sh2 = 12
row_sh3 = 12
row_sh4 = 12
row_sh5 = 12

End Sub

