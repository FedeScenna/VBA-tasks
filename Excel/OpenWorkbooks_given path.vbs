Dim wbk As Workbook
Dim w_path As Object
For i = 1 To Range("A3", Range("a3").End(xlDown)).Rows.Count Step 1
If Right(Cells(j + 2, 1), 4) <> ".zip" Then
Set w_path = Cells(j + 2, 1)
Set wbk = Workbooks.Open(w_path)
wbk.Worksheets("data").Activate
wbk.Save
wbk.Close
End If
Next i