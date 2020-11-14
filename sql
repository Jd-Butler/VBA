Option Explicit

Sub f()
Dim cn As New ADODB.Connection
Dim query As String
Dim rs As New ADODB.Recordset
Dim i As Long
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName _
& ";Extended Properties=""Excel 8.0;HDR=YES"";"

query = "Select * from [Sheet1$] Where Scorecard = 'CSC'"
Debug.Print query

rs.Open query, cn

Sheet2.Cells.ClearContents

'write query

For i = 0 To rs.Fields.Count - 1
    Sheet2.Cells(1, i + 1).Value2 = rs.Fields(i).Name
Next i

Sheet2.Range("A2").CopyFromRecordset rs

Sheet2.Cells.Columns.AutoFit

set rs = nothing
cn.Close
End Sub
