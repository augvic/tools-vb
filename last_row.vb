Function lastRow(coluna As Integer)

lastRow = ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Rows.Count, coluna).End(xlUp).Row

End Function