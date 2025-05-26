Sub cotação_de_frete()

Set objeto_outlook = CreateObject("Outlook.Application")
Set Email = objeto_outlook.CreateItem(0)

Email.Display
Email.To = Worksheets("COTAÇÃO").Cells(3, 23)
Email.CC = Worksheets("COTAÇÃO").Cells(4, 23)
Email.Subject = "COTAÇÃO DE FRETE | PEDIDO / COTAÇÃO: " & Worksheets("COTAÇÃO").Cells(5, 10) & " | ORDEM: " & Worksheets("COTAÇÃO").Cells(5, 14)
Email.HTMLBody = RangetoHTML(Range("B2:T22"))

' ~~ Coletando dados. ~~ '
Dim valortotal As Currency
If ThisWorkbook.Worksheets("COTAÇÃO").Cells(5, "D").Value = "81.243.735/0034-06 - POSILHE - 45658-464 - Ilhéus - BA" Then
    planta = "IOS"
End If
If ThisWorkbook.Worksheets("COTAÇÃO").Cells(5, "D").Value = "81.243.735/0019-77 - POSMAN - 69075-110 - Manaus - AM" Then
    planta = "MAO"
End If
If ThisWorkbook.Worksheets("COTAÇÃO").Cells(5, "D").Value = "81.243.735/0001-48 - POSCTB - 81350-000 - Curitiba - PR" Then
    planta = "CWB"
End If
pedidocotacao = ThisWorkbook.Worksheets("COTAÇÃO").Cells(5, "J").Value
ordem = ThisWorkbook.Worksheets("COTAÇÃO").Cells(5, "N").Value
valortotal = CCur(ThisWorkbook.Worksheets("COTAÇÃO").Cells(18, "J").Value)

' ~~ Inserindo na tabela. ~~ '
ultimaLinha = lastRow("A") + 1
ThisWorkbook.Worksheets("TABELA MENSAL").Cells(ultimaLinha, "A") = planta
ThisWorkbook.Worksheets("TABELA MENSAL").Cells(ultimaLinha, "B") = pedidocotacao
ThisWorkbook.Worksheets("TABELA MENSAL").Cells(ultimaLinha, "C") = ordem
ThisWorkbook.Worksheets("TABELA MENSAL").Cells(ultimaLinha, "D") = valortotal

End Sub