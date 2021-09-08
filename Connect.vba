'Below codes contain three modules, connection module contains the access database and excel, download module could download data from access database to excel main worksheet, and update module update data of access database with data in excel main worksheet.

'connection module:

Option Explicit

Public blnisConnected As Boolean

Public cnnDb As ADODB.Connection


Public Sub connection_database()

blnisConnected = False

On Error GoTo Errhandling

Set cnnDb = New ADODB.Connection

With cnnDb
    .Provider = "Microsoft.ACE.OLEDB.12.0"
    .ConnectionString = "C:\Users\lyt\Documents\Excel_VBA\week 4\Finance_DB1.accdb"
    .Properties("Jet OLEDB:Database Password") = ""
    .Open
End With

blnisConnected = True


Exit Sub


Errhandling:
MsgBox "Connection failed", vbCritical, "Error"


End Sub


Public Sub Disconnect_Database()


cnnDb.Close

Set cnnDb = Nothing
blnisConnected = False

End Sub

'download module
Option Explicit

Sub Download_transaction()

Application.ScreenUpdating = False

Application.Calculation = xlCalculationManual

Dim shtDwn As Worksheet
Dim sql_query As String
Dim rstDb As ADODB.Recordset
Dim download_row As Integer



Set shtDwn = Sheets("DOWNLOAD")
sql_query = "select sales_date,product_id,sales_status,sales_price from tbSales WHERE shop_id = 1 "


Call connection_database

Set rstDb = New ADODB.Recordset
download_row = 12

With rstDb

    .Open Source:=sql_query, ActiveConnection:=cnnDb
    Do While Not .EOF
        shtDwn.Cells(download_row, 3).Value = .Fields("sales_date")
        shtDwn.Cells(download_row, 4).Value = .Fields("product_id")
        shtDwn.Cells(download_row, 5).Value = .Fields("sales_status")
        shtDwn.Cells(download_row, 6).Value = .Fields("sales_price")
        download_row = download_row + 1
        .MoveNext
        
    Loop
    
    

End With

Call Disconnect_Database


End Sub


'update modules

Option Explicit

Sub Upload_Data()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim shtUp As Worksheet
Dim row_upload As Integer
Dim sql_query As String
Dim rstDatabase As ADODB.Recordset

Call connection_database

Set shtUp = Sheets("Upload")
Set rstDatabase = New ADODB.Recordset

row_upload = 12

With rstDatabase

    .Open Source:="tbPerformance", ActiveConnection:=cnnDb, LockType:=adLockOptimistic

    Do While Not IsEmpty(shtUp.Cells(row_upload, 3))
        
        If shtUp.Cells(row_upload, 7) = "Target correction" Then
            'UPDATE code
            sql_query = "UPDATE tbPerformance SET target_value = " & shtUp.Cells(row_upload, 6) & " WHERE target_id = " & shtUp.Cells(row_upload, 3)
            Debug.Print sql_query
            cnnDb.Execute sql_query
            
        ElseIf shtUp.Cells(row_upload, 7) = "New" Then
            'INSERT code
            .AddNew
            .Fields("target_id") = shtUp.Cells(row_upload, 3)
            .Fields("target_week") = shtUp.Cells(row_upload, 4)
            .Fields("shop_id") = shtUp.Cells(row_upload, 5)
            .Fields("target_value") = shtUp.Cells(row_upload, 6)
            .Update
        Else
            MsgBox "Status not recognized!", vbCritical, "Error!"
        End If
        
        row_upload = row_upload + 1
    Loop

End With

Call Disconnect_Database

End Sub



