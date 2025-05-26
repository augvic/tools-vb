Public diferenca As Integer
Public prazoInt As Integer
Public centro As Integer
Public nfToLeft As String
Public posicaoHifen As Variant
Public notaFiscal As String
Public docT As String
Public tipoT As String
Public entregaRemessa As String
Public dataIT As Variant
Public dataDF As Variant
Public dataENT As Variant
Public ordem As String
Public remessa As String
Public docDT As String
Public localEXP As String
Public diferencaCell As String
Public feriado1 As Date
Public feriado2 As Date
Public feriado3 As Date
Public feriado4 As Date
Public feriado5 As Date
Public feriado6 As Date
Public feriado7 As Date
Public feriado8 As Date
Public feriado9 As Date
Public feriado10 As Date
Public feriados As Variant
Public contadora As Integer
Public prazoStr As String
Public mensagem As String

Sub mon_entregas()

' ~~ Primeiro checa se a primeira linha está em branco. ~~ '
If ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(5, "A").Value = "" Then
    MsgBox ("Nenhuma linha para monitoramento.")
    Exit Sub
End If

' # Tirando "-" das NF's. # '
ultimaLinha = lastRow(1)
For NF = 5 To ultimaLinha
    nfToLeft = ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(NF, "A").Value
    posicaoHifen = InStr(nfToLeft, "-")
    If posicaoHifen > 0 Then
        nfToLeft = Left(nfToLeft, posicaoHifen - 1)
        ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(NF, "A") = nfToLeft
    End If
Next NF

' # Iniciando Loop. # '
On Error Resume Next
ultimaLinha = lastRow(1)
For linha = 5 To ultimaLinha
  
    ' ~~ Definindo valores. ~~ '
    notaFiscal = ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "A").text
    ordem = "-"
    cliente = "-"
    remessa = "-"
    docDT = "-"
    localEXP = "-"
    estado = "-"
    cidade = "-"
    transportadora = "-"
    dataDF = "-"
    dataIT = "-"
    dataENT = "-"
    
    ' ~~ Acessando ZSD022. ~~ '
    Set SapGuiAuto = GetObject("SAPGUI")
    Set objGui = SapGuiAuto.GetScriptingEngine
    Set objConn = objGui.Children(0)
    Set session = objConn.Children(0)
    session.FindById("wnd[0]/tbar[0]/okcd").text = "/NZSD022"
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[0]/tbar[1]/btn[17]").press
    session.FindById("wnd[1]/usr/txtENAME-LOW").text = "72776"
    session.FindById("wnd[1]/tbar[0]/btn[8]").press
    session.FindById("wnd[0]/usr/ctxtS_NFNUM-LOW").text = notaFiscal
    periodoInicial = ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(3, "B").text
    periodoFinal = "31.12.9999"
    session.FindById("wnd[0]/usr/ctxtS_DOCDAT-LOW").text = periodoInicial
    session.FindById("wnd[0]/usr/ctxtS_DOCDAT-HIGH").text = periodoFinal
    session.FindById("wnd[0]/tbar[1]/btn[8]").press
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
    
    ' ~~ Coletando dados da ZSD022. ~~ '
    dataIT = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "DATBG")
    dataIT = Split(dataIT, ".")
    dataIT = dataIT(1) & "/" & dataIT(0) & "/" & dataIT(2)
    dataIT = Format(CDate(dataIT), "DD/MM/YYYY")
    dataENT = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "DATRECEB")
    dataENT = Split(dataENT, ".")
    dataENT = dataENT(1) & "/" & dataENT(0) & "/" & dataENT(2)
    dataENT = Format(CDate(dataENT), "DD/MM/YYYY")
    cidade = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "ORT01")
    estado = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "REGIO")
    dataDF = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "DOCDAT")
    dataDF = Split(dataDF, ".")
    dataDF = dataDF(1) & "/" & dataDF(0) & "/" & dataDF(2)
    dataDF = Format(CDate(dataDF), "DD/MM/YYYY")
    ordem = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "ORDEM")
    cliente = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "NOME_PARID")
    remessa = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "VBELN")
    docDT = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "TKNUM")
    localEXP = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "VSTEL")
    transportadora = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "NOME_TRANSP")
    
    ' ~~ Inserindo dados na planilha. ~~
    ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "B") = ordem
    ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "C") = cliente
    ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "D") = remessa
    ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "E") = docDT
    ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "F") = localEXP
    ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "G") = estado
    ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "H") = cidade
    ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "I") = transportadora
    ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "J") = dataDF
    ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "K") = dataIT
    If dataIT = "" Then
        ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "K") = "-"
    End If
    If dataENT = "" Then
        ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "M") = "-"
    Else
        ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "M") = dataENT
    End If
    cidadeConversao = UCase(ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "H").text)
    cidadeConversao = Acento(cidadeConversao)
    ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "H") = cidadeConversao
    
    ' ~~ Caso não encontre dados, flega ocorrência. ~~ '
    If ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "B").Value = "-" Then
        session.FindById("wnd[0]/usr/chkP_S_OCOR").Selected = True
        session.FindById("wnd[0]/tbar[1]/btn[8]").press
        session.FindById("wnd[1]/tbar[0]/btn[0]").press
        ' ~~ Coletando dados da ZSD022. ~~ '
        dataIT = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "DATBG")
        dataIT = Split(dataIT, ".")
        dataIT = dataIT(1) & "/" & dataIT(0) & "/" & dataIT(2)
        dataIT = Format(CDate(dataIT), "DD/MM/YYYY")
        dataENT = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "DATRECEB")
        dataENT = Split(dataENT, ".")
        dataENT = dataENT(1) & "/" & dataENT(0) & "/" & dataENT(2)
        dataENT = Format(CDate(dataENT), "DD/MM/YYYY")
        cidade = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "ORT01")
        estado = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "REGIO")
        dataDF = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "DOCDAT")
        dataDF = Split(dataDF, ".")
        dataDF = dataDF(1) & "/" & dataDF(0) & "/" & dataDF(2)
        dataDF = Format(CDate(dataDF), "DD/MM/YYYY")
        ordem = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "ORDEM")
        cliente = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "NOME_PARID")
        remessa = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "VBELN")
        docDT = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "TKNUM")
        localEXP = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "VSTEL")
        transportadora = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(0, "NOME_TRANSP")
        ' ~~ Inserindo dados na planilha. ~~
        ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "B") = ordem
        ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "C") = cliente
        ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "D") = remessa
        ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "E") = docDT
        ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "F") = localEXP
        ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "G") = estado
        ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "H") = cidade
        ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "I") = transportadora
        ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "J") = dataDF
        ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "K") = dataIT
        If dataENT = "" Then
            ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "M") = "-"
        Else
            ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "M") = dataENT
        End If
        cidadeConversao = UCase(ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "H").text)
        cidadeConversao = Acento(cidadeConversao)
        ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "H") = cidadeConversao
    End If
    
    ' ~~ Pegando previsão. ~~ '
    If Not ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "K") = "-" Then ' <= Se não tiver início de transporte, ignora cálculo. '
        docT = ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "E").Value
        If ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "F").Value = "1910" Then
            ThisWorkbook.Worksheets("PRAZOS").Cells(2, "E") = "MANAUS"
        End If
        If ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "F").Value = "3540" Then
            ThisWorkbook.Worksheets("PRAZOS").Cells(2, "E") = "MANAUS"
        End If
        If ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "F").Value = "3010" Then
            ThisWorkbook.Worksheets("PRAZOS").Cells(2, "E") = "ILHEUS"
        End If
        If ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "F").Value = "3550" Then
            ThisWorkbook.Worksheets("PRAZOS").Cells(2, "E") = "ILHEUS"
        End If
        If ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "F").Value = "1010" Then
            ThisWorkbook.Worksheets("PRAZOS").Cells(2, "E") = "CURITIBA"
        End If
        If ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "F").Value = "1099" Then
            ThisWorkbook.Worksheets("PRAZOS").Cells(2, "E") = "CURITIBA"
        End If
        If ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "F").Value = "3520" Then
            ThisWorkbook.Worksheets("PRAZOS").Cells(2, "E") = "CURITIBA"
        End If
        ThisWorkbook.Worksheets("PRAZOS").Cells(2, "F") = ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "G")
        ThisWorkbook.Worksheets("PRAZOS").Cells(2, "G") = ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "H")
        Set SapGuiAuto = GetObject("SAPGUI")
        Set objGui = SapGuiAuto.GetScriptingEngine
        Set objConn = objGui.Children(0)
        Set session = objConn.Children(0)
        session.FindById("wnd[0]/tbar[0]/okcd").text = "/NVT03N"
        session.FindById("wnd[0]").SendVKey 0
        session.FindById("wnd[0]/usr/ctxtVTTK-TKNUM").text = docT
        session.FindById("wnd[0]").SendVKey 0
        tipoT = CStr(session.FindById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_PR/ssubG_HEADER_SUBSCREEN1:SAPMV56A:1021/ctxtVTTK-VSART").text)
        Set prazos = ThisWorkbook.Worksheets("PRAZOS")
        Set colunaFiltrada = prazos.Range("A:A")
        prazos.AutoFilterMode = False
        colunaFiltrada.AutoFilter
        prazoConcat = ThisWorkbook.Worksheets("PRAZOS").Cells(2, "H").Value
        colunaFiltrada.AutoFilter Field:=1, Criteria1:=prazoConcat
        If tipoT = "02" Or "14" Then
            Set visibleRange = ThisWorkbook.Worksheets("PRAZOS").Columns("B").SpecialCells(xlCellTypeVisible)
            Set secondRow = visibleRange.Areas(2).Cells(1, 1)
            prazoInt = secondRow.Value
        End If
        If tipoT = "01" Then
            Set visibleRange = ThisWorkbook.Worksheets("PRAZOS").Columns("C").SpecialCells(xlCellTypeVisible)
            Set secondRow = visibleRange.Areas(2).Cells(1, 1)
            prazoInt = secondRow.Value
        End If
        If tipoT = "17" Or tipoT = "16" Then
            prazoInt = 4
        End If
        inicioTransporte = CDate(ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "K").Value)
        previsaoData = AdicionarDiasUteis(inicioTransporte, prazoInt)
        ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "L") = previsaoData
        mensagem = " DIAS ÚTEIS"
        prazoStr = CStr(prazoInt)
        msgfinal = prazoStr + mensagem
        ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "P") = msgfinal
    Else ' <= Preenche célula de previsão com nulo, caso não haja início de transporte. '
        ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "L") = "-"
        docT = ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "E").Value
        If ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "F").Value = "1910" Then
            ThisWorkbook.Worksheets("PRAZOS").Cells(2, "E") = "MANAUS"
        End If
        If ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "F").Value = "3540" Then
            ThisWorkbook.Worksheets("PRAZOS").Cells(2, "E") = "MANAUS"
        End If
        If ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "F").Value = "3010" Then
            ThisWorkbook.Worksheets("PRAZOS").Cells(2, "E") = "ILHEUS"
        End If
        If ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "F").Value = "3550" Then
            ThisWorkbook.Worksheets("PRAZOS").Cells(2, "E") = "ILHEUS"
        End If
        If ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "F").Value = "1010" Then
            ThisWorkbook.Worksheets("PRAZOS").Cells(2, "E") = "CURITIBA"
        End If
        If ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "F").Value = "1099" Then
            ThisWorkbook.Worksheets("PRAZOS").Cells(2, "E") = "CURITIBA"
        End If
        If ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "F").Value = "3520" Then
            ThisWorkbook.Worksheets("PRAZOS").Cells(2, "E") = "CURITIBA"
        End If
        ThisWorkbook.Worksheets("PRAZOS").Cells(2, "F") = ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "G")
        ThisWorkbook.Worksheets("PRAZOS").Cells(2, "G") = ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "H")
        Set SapGuiAuto = GetObject("SAPGUI")
        Set objGui = SapGuiAuto.GetScriptingEngine
        Set objConn = objGui.Children(0)
        Set session = objConn.Children(0)
        session.FindById("wnd[0]/tbar[0]/okcd").text = "/NVT03N"
        session.FindById("wnd[0]").SendVKey 0
        session.FindById("wnd[0]/usr/ctxtVTTK-TKNUM").text = docT
        session.FindById("wnd[0]").SendVKey 0
        tipoT = CStr(session.FindById("wnd[0]/usr/tabsHEADER_TABSTRIP1/tabpTABS_OV_PR/ssubG_HEADER_SUBSCREEN1:SAPMV56A:1021/ctxtVTTK-VSART").text)
        Set prazos = ThisWorkbook.Worksheets("PRAZOS")
        Set colunaFiltrada = prazos.Range("A:A")
        prazos.AutoFilterMode = False
        colunaFiltrada.AutoFilter
        prazoConcat = ThisWorkbook.Worksheets("PRAZOS").Cells(2, "H").Value
        colunaFiltrada.AutoFilter Field:=1, Criteria1:=prazoConcat
        If tipoT = "02" Or "14" Then
            Set visibleRange = ThisWorkbook.Worksheets("PRAZOS").Columns("B").SpecialCells(xlCellTypeVisible)
            Set secondRow = visibleRange.Areas(2).Cells(1, 1)
            prazoInt = secondRow.Value
        End If
        If tipoT = "01" Then
            Set visibleRange = ThisWorkbook.Worksheets("PRAZOS").Columns("C").SpecialCells(xlCellTypeVisible)
            Set secondRow = visibleRange.Areas(2).Cells(1, 1)
            prazoInt = secondRow.Value
        End If
        If tipoT = "17" Or tipoT = "16" Then
            prazoInt = 4
        End If
        mensagem = " DIAS ÚTEIS"
        prazoStr = CStr(prazoInt)
        msgfinal = prazoStr + mensagem
        ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "P") = msgfinal
    End If
    
    ' ~~ Validando diferença. ~~ '
    inicioTransporte = ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "K").Value
    dataEntrega = ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "M").Value
    previsao = ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "L").Value
    dataAtual = ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(1, "P").Value
    If Not inicioTransporte = "-" Then ' <= Se transporte tiver iniciado. '
        If Not dataEntrega = "-" Then ' <= Se tiver sido entregue. '
            diferenca = calcularDiferenca(dataEntrega, previsao)
            ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "N") = diferenca
        Else ' <= Se não tiver sido entregue. '
            ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "N") = "-"
        End If
    Else ' <= Se transporte não tiver iniciado. '
        ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "N") = "-"
    End If
    
    ' ~~ Definindo status. ~~ '
    diferencaCell = ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "N").Value
    If Not ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "K") = "-" Then ' <= Se tiver início de transporte. '
        If diferencaCell = "-" Then ' <= Se ainda não entregue. '
            If dataAtual <= previsao Then
                ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "O") = "AGUARDANDO ENTREGA - NO PRAZO"
            Else
                ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "O") = "AGUARDANDO ENTREGA - EM ATRASO"
            End If
        Else
            If diferencaCell <= 0 Then ' <= Se entregue no prazo. '
                ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "O") = "ENTREGUE NO PRAZO"
            Else ' <= Se entregue com atraso. '
                ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "O") = "ENTREGUE COM ATRASO"
            End If
        End If
    Else ' <= Se não tiver início de transporte. '
        If Not ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "M") = "-" Then ' <= Se constar entrega. '
            ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "K") = "-"
            ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "L") = "-"
            ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "N") = "!-"
            ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "O") = "ENTREGUE - SEM IT"
        Else ' <= Se não constar entrega. '
            ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "K") = "-"
            ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "L") = "-"
            ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "M") = "-"
            ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "N") = "--"
            ThisWorkbook.Worksheets("MONITORAMENTO ENTREGAS").Cells(linha, "O") = "NÃO EXPEDIDO"
        End If
    End If
        
Next linha

Set session = Nothing
MsgBox ("Notas monitoradas com sucesso.")

End Sub