Function AdicionarDiasUteis(ByVal inicioTransporte As Date, ByVal prazoInt As Integer) As Date

feriado1 = CDate(ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(2, "K").Value)
feriado2 = CDate(ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(2, "L").Value)
feriado3 = CDate(ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(2, "M").Value)
feriado4 = CDate(ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(2, "N").Value)
feriado5 = CDate(ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(2, "O").Value)
feriado6 = CDate(ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(3, "K").Value)
feriado7 = CDate(ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(3, "L").Value)
feriado8 = CDate(ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(3, "M").Value)
feriado9 = CDate(ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(3, "N").Value)
feriado10 = CDate(ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(3, "O").Value)
feriado11 = CDate(ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(2, "P").Value)
feriado12 = CDate(ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(3, "P").Value)
feriados = Array(feriado1, feriado2, feriado3, feriado4, feriado5, feriado6, feriado7, feriado8, feriado9, feriado10, feriado11, feriado12)
contador = 0

Do While contador < prazoInt
    inicioTransporte = inicioTransporte + 1
    If Weekday(inicioTransporte) <> vbSaturday And Weekday(inicioTransporte) <> vbSunday And Not IsFeriado(inicioTransporte, feriados) Then
        contador = contador + 1
    End If
Loop

AdicionarDiasUteis = inicioTransporte

End Function