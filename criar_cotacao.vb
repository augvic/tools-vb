Public SapGuiAuto
Public objGui As GuiApplication
Public objConn As GuiConnection
Public session As GuiSession

Sub ORDEM_COTAÇÃO()

' # Variáveis # '
Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)
Dim DOC, TRANSAÇÃOCRIAR, TRANSAÇÃOMODIFICAR As String
Dim ORGANIZAÇÃO, CANAL, ESCRITÓRIO, EQUIPE As String
Dim EMISSOR, RECEBEDOR, PEDIDO, dataAtual As String
Dim CONDPAG, INCOTERM, MOTIVO As String
Dim TABELA, EXPEDIÇÃO, TIPOPAG, AGENTE As String
Dim texto As String
Dim GARANTIA10, GARANTIA20, GARANTIA30, GARANTIA40, GARANTIA50, GARANTIA60, GARANTIA70, GARANTIA80, GARANTIA90 As String
Dim Z2, Z3, Z5, Z6, Z7, Z2VALOR, Z3VALOR, Z5VALOR, Z6VALOR, Z7VALOR As String
Dim ITEM10, QUANTIDADE10, CENTRO10, DEPÓSITO10, VALOR10NOTEONSITE, VALOR10DESKONSITE, VALOR10UNIONONSITE, VALOR10DESKBALCÃO, VALOR10NOTEBALCÃO, VALOR10SEMGARANTIA As String
Dim ITEM20, QUANTIDADE20, CENTRO20, DEPÓSITO20, VALOR20NOTEONSITE, VALOR20DESKONSITE, VALOR20UNIONONSITE, VALOR20DESKBALCÃO, VALOR20NOTEBALCÃO, VALOR20SEMGARANTIA As String
Dim ITEM30, QUANTIDADE30, CENTRO30, DEPÓSITO30, VALOR30NOTEONSITE, VALOR30DESKONSITE, VALOR30UNIONONSITE, VALOR30DESKBALCÃO, VALOR30NOTEBALCÃO, VALOR30SEMGARANTIA As String
Dim ITEM40, QUANTIDADE40, CENTRO40, DEPÓSITO40, VALOR40NOTEONSITE, VALOR40DESKONSITE, VALOR40UNIONONSITE, VALOR40DESKBALCÃO, VALOR40NOTEBALCÃO, VALOR40SEMGARANTIA As String
Dim ITEM50, QUANTIDADE50, CENTRO50, DEPÓSITO50, VALOR50NOTEONSITE, VALOR50DESKONSITE, VALOR50UNIONONSITE, VALOR50DESKBALCÃO, VALOR50NOTEBALCÃO, VALOR50SEMGARANTIA As String
Dim ITEM60, QUANTIDADE60, CENTRO60, DEPÓSITO60, VALOR60NOTEONSITE, VALOR60DESKONSITE, VALOR60UNIONONSITE, VALOR60DESKBALCÃO, VALOR60NOTEBALCÃO, VALOR60SEMGARANTIA As String
Dim ITEM70, QUANTIDADE70, CENTRO70, DEPÓSITO70, VALOR70NOTEONSITE, VALOR70DESKONSITE, VALOR70UNIONONSITE, VALOR70DESKBALCÃO, VALOR70NOTEBALCÃO, VALOR70SEMGARANTIA As String
Dim ITEM80, QUANTIDADE80, CENTRO80, DEPÓSITO80, VALOR80NOTEONSITE, VALOR80DESKONSITE, VALOR80UNIONONSITE, VALOR80DESKBALCÃO, VALOR80NOTEBALCÃO, VALOR80SEMGARANTIA As String
Dim ITEM90, QUANTIDADE90, CENTRO90, DEPÓSITO90, VALOR90NOTEONSITE, VALOR90DESKONSITE, VALOR90UNIONONSITE, VALOR90DESKBALCÃO, VALOR90NOTEBALCÃO, VALOR90SEMGARANTIA As String

' # Atribuição de Valores # '
DOC = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(3, 3).Value
TRANSAÇÃOCRIAR = ThisWorkbook.Worksheets("DADOS").Cells(5, 3).Value
TRANSAÇÃOMODIFICAR = ThisWorkbook.Worksheets("DADOS").Cells(18, 3).Value

ORGANIZAÇÃO = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(5, 4).Value
CANAL = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(6, 4).Value
ESCRITÓRIO = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(7, 4).Value
EQUIPE = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(8, 4).Value

EMISSOR = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(10, 4).Value
RECEBEDOR = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(11, 4).Value
PEDIDO = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(12, 4).Value
dataAtual = ThisWorkbook.Worksheets("DADOS").Cells(2, 1).text

CONDPAG = ThisWorkbook.Worksheets("DADOS").Cells(2, 29).Value
INCOTERM = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(11, 7).Value
MOTIVO = ThisWorkbook.Worksheets("DADOS").Cells(2, "BB").Value

TABELA = ThisWorkbook.Worksheets("DADOS").Cells(2, "BF").Value
EXPEDIÇÃO = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(6, 7).Value
TIPOPAG = ThisWorkbook.Worksheets("DADOS").Cells(2, 50).Value
AGENTE = ThisWorkbook.Worksheets("DADOS").Cells(2, 43).Value
texto = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(3, 6).Value
REVENDA = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(3, "T").Value

GARANTIA10 = ThisWorkbook.Worksheets("DADOS").Cells(3, 24).Value
GARANTIA20 = ThisWorkbook.Worksheets("DADOS").Cells(4, 24).Value
GARANTIA30 = ThisWorkbook.Worksheets("DADOS").Cells(5, 24).Value
GARANTIA40 = ThisWorkbook.Worksheets("DADOS").Cells(6, 24).Value
GARANTIA50 = ThisWorkbook.Worksheets("DADOS").Cells(7, 24).Value
GARANTIA60 = ThisWorkbook.Worksheets("DADOS").Cells(8, 24).Value
GARANTIA70 = ThisWorkbook.Worksheets("DADOS").Cells(9, 24).Value
GARANTIA80 = ThisWorkbook.Worksheets("DADOS").Cells(10, 24).Value
GARANTIA90 = ThisWorkbook.Worksheets("DADOS").Cells(11, 24).Value

PRODUTO10 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(3, "Q").Value
PRODUTO20 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(4, "Q").Value
PRODUTO30 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(5, "Q").Value
PRODUTO40 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(6, "Q").Value
PRODUTO50 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(7, "Q").Value
PRODUTO60 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(8, "Q").Value
PRODUTO70 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(9, "Q").Value
PRODUTO80 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(10, "Q").Value
PRODUTO90 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(11, "Q").Value

Z1 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(5, "T").Value
Z2 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(6, "T").Value
Z3 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(7, "T").Value
Z4 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(8, "T").Value
Z5 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(9, "T").Value
Z6 = ThisWorkbook.Worksheets("DADOS").Cells(2, 39).Value
Z7 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(11, "T").Value
Z1VALOR = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(5, "U").Value
Z2VALOR = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(6, "U").Value
Z3VALOR = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(7, "U").Value
Z4VALOR = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(8, "U").Value
Z5VALOR = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(9, "U").Value
Z6VALOR = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(10, "U").Value
Z7VALOR = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(11, "U").Value

ITEM10 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(3, 10).Value
QUANTIDADE10 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(3, 11).Value
CENTRO10 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(3, 12).Value
DEPÓSITO10 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(3, 13).Value
VALOR10NOTEONSITE = ThisWorkbook.Worksheets("DADOS").Cells(3, 6).Value
VALOR10DESKONSITE = ThisWorkbook.Worksheets("DADOS").Cells(3, 7).Value
VALOR10UNIONONSITE = ThisWorkbook.Worksheets("DADOS").Cells(3, 8).Value
VALOR10DESKBALCÃO = ThisWorkbook.Worksheets("DADOS").Cells(3, 9).Value
VALOR10NOTEBALCÃO = ThisWorkbook.Worksheets("DADOS").Cells(3, 10).Value
VALOR10SEMGARANTIA = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(3, 14).Value
OVER10 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(3, "O").Value

ITEM20 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(4, 10).Value
QUANTIDADE20 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(4, 11).Value
CENTRO20 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(4, 12).Value
DEPÓSITO20 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(4, 13).Value
VALOR20NOTEONSITE = ThisWorkbook.Worksheets("DADOS").Cells(4, 6).Value
VALOR20DESKONSITE = ThisWorkbook.Worksheets("DADOS").Cells(4, 7).Value
VALOR20UNIONONSITE = ThisWorkbook.Worksheets("DADOS").Cells(4, 8).Value
VALOR20DESKBALCÃO = ThisWorkbook.Worksheets("DADOS").Cells(4, 9).Value
VALOR20NOTEBALCÃO = ThisWorkbook.Worksheets("DADOS").Cells(4, 10).Value
VALOR20SEMGARANTIA = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(4, 14).Value
OVER20 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(4, "O").Value

ITEM30 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(5, 10).Value
QUANTIDADE30 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(5, 11).Value
CENTRO30 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(5, 12).Value
DEPÓSITO30 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(5, 13).Value
VALOR30NOTEONSITE = ThisWorkbook.Worksheets("DADOS").Cells(5, 6).Value
VALOR30DESKONSITE = ThisWorkbook.Worksheets("DADOS").Cells(5, 7).Value
VALOR30UNIONONSITE = ThisWorkbook.Worksheets("DADOS").Cells(5, 8).Value
VALOR30DESKBALCÃO = ThisWorkbook.Worksheets("DADOS").Cells(5, 9).Value
VALOR30NOTEBALCÃO = ThisWorkbook.Worksheets("DADOS").Cells(5, 10).Value
VALOR30SEMGARANTIA = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(5, 14).Value
OVER30 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(5, "O").Value

ITEM40 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(6, 10).Value
QUANTIDADE40 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(6, 11).Value
CENTRO40 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(6, 12).Value
DEPÓSITO40 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(6, 13).Value
VALOR40NOTEONSITE = ThisWorkbook.Worksheets("DADOS").Cells(6, 6).Value
VALOR40DESKONSITE = ThisWorkbook.Worksheets("DADOS").Cells(6, 7).Value
VALOR40UNIONONSITE = ThisWorkbook.Worksheets("DADOS").Cells(6, 8).Value
VALOR40DESKBALCÃO = ThisWorkbook.Worksheets("DADOS").Cells(6, 9).Value
VALOR40NOTEBALCÃO = ThisWorkbook.Worksheets("DADOS").Cells(6, 10).Value
VALOR40SEMGARANTIA = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(6, 14).Value
OVER40 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(6, "O").Value

ITEM50 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(7, 10).Value
QUANTIDADE50 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(7, 11).Value
CENTRO50 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(7, 12).Value
DEPÓSITO50 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(7, 13).Value
VALOR50NOTEONSITE = ThisWorkbook.Worksheets("DADOS").Cells(7, 6).Value
VALOR50DESKONSITE = ThisWorkbook.Worksheets("DADOS").Cells(7, 7).Value
VALOR50UNIONONSITE = ThisWorkbook.Worksheets("DADOS").Cells(7, 8).Value
VALOR50DESKBALCÃO = ThisWorkbook.Worksheets("DADOS").Cells(7, 9).Value
VALOR50NOTEBALCÃO = ThisWorkbook.Worksheets("DADOS").Cells(7, 10).Value
VALOR50SEMGARANTIA = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(7, 14).Value
OVER50 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(7, "O").Value

ITEM60 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(8, 10).Value
QUANTIDADE60 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(8, 11).Value
CENTRO60 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(8, 12).Value
DEPÓSITO60 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(8, 13).Value
VALOR60NOTEONSITE = ThisWorkbook.Worksheets("DADOS").Cells(8, 6).Value
VALOR60DESKONSITE = ThisWorkbook.Worksheets("DADOS").Cells(8, 7).Value
VALOR60UNIONONSITE = ThisWorkbook.Worksheets("DADOS").Cells(8, 8).Value
VALOR60DESKBALCÃO = ThisWorkbook.Worksheets("DADOS").Cells(8, 9).Value
VALOR60NOTEBALCÃO = ThisWorkbook.Worksheets("DADOS").Cells(8, 10).Value
VALOR60SEMGARANTIA = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(8, 14).Value
OVER60 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(8, "O").Value

ITEM70 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(9, 10).Value
QUANTIDADE70 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(9, 11).Value
CENTRO70 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(9, 12).Value
DEPÓSITO70 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(9, 13).Value
VALOR70NOTEONSITE = ThisWorkbook.Worksheets("DADOS").Cells(9, 6).Value
VALOR70DESKONSITE = ThisWorkbook.Worksheets("DADOS").Cells(9, 7).Value
VALOR70UNIONONSITE = ThisWorkbook.Worksheets("DADOS").Cells(9, 8).Value
VALOR70DESKBALCÃO = ThisWorkbook.Worksheets("DADOS").Cells(9, 9).Value
VALOR70NOTEBALCÃO = ThisWorkbook.Worksheets("DADOS").Cells(9, 10).Value
VALOR70SEMGARANTIA = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(9, 14).Value
OVER70 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(9, "O").Value

ITEM80 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(10, 10).Value
QUANTIDADE80 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(10, 11).Value
CENTRO80 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(10, 12).Value
DEPÓSITO80 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(10, 13).Value
VALOR80NOTEONSITE = ThisWorkbook.Worksheets("DADOS").Cells(10, 6).Value
VALOR80DESKONSITE = ThisWorkbook.Worksheets("DADOS").Cells(10, 7).Value
VALOR80UNIONONSITE = ThisWorkbook.Worksheets("DADOS").Cells(10, 8).Value
VALOR80DESKBALCÃO = ThisWorkbook.Worksheets("DADOS").Cells(10, 9).Value
VALOR80NOTEBALCÃO = ThisWorkbook.Worksheets("DADOS").Cells(10, 10).Value
VALOR80SEMGARANTIA = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(10, 14).Value
OVER80 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(10, "O").Value

ITEM90 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(11, 10).Value
QUANTIDADE90 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(11, 11).Value
CENTRO90 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(11, 12).Value
DEPÓSITO90 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(11, 13).Value
VALOR90NOTEONSITE = ThisWorkbook.Worksheets("DADOS").Cells(11, 6).Value
VALOR90DESKONSITE = ThisWorkbook.Worksheets("DADOS").Cells(11, 7).Value
VALOR90UNIONONSITE = ThisWorkbook.Worksheets("DADOS").Cells(11, 8).Value
VALOR90DESKBALCÃO = ThisWorkbook.Worksheets("DADOS").Cells(11, 9).Value
VALOR90NOTEBALCÃO = ThisWorkbook.Worksheets("DADOS").Cells(11, 10).Value
VALOR90SEMGARANTIA = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(11, 14).Value
OVER90 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(11, "O").Value

' ---------- # Script SAP # ---------- '
' # Inserindo dados para criar ordem/cotação # '
session.FindById("wnd[0]").maximize
session.FindById("wnd[0]/tbar[0]/okcd").text = TRANSAÇÃOCRIAR
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/ctxtVBAK-AUART").text = DOC
session.FindById("wnd[0]/usr/ctxtVBAK-VKORG").text = ORGANIZAÇÃO
session.FindById("wnd[0]/usr/ctxtVBAK-VTWEG").text = CANAL
session.FindById("wnd[0]/usr/ctxtVBAK-SPART").text = "00"
session.FindById("wnd[0]/usr/ctxtVBAK-VKBUR").text = ESCRITÓRIO
session.FindById("wnd[0]/usr/ctxtVBAK-VKGRP").text = "0" & EQUIPE
session.FindById("wnd[0]").SendVKey 0

' # Dados do pedido e emissor e recebedor # '
session.FindById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text = PEDIDO
session.FindById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBKD-BSTDK").text = dataAtual
session.FindById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").text = EMISSOR
session.FindById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR").text = RECEBEDOR
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtVBKD-ZTERM").text = CONDPAG
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtVBKD-INCO1").text = INCOTERM
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-AUGRU").Key = MOTIVO

' # Sistema para ignorar validade # '
On Error Resume Next
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/ctxtVBAK-BNDDT").text = ThisWorkbook.Worksheets("DADOS").Cells(2, 2).text
On Error GoTo 0
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]").SendVKey 0

' # Marcando fornecimento completo # '
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/chkVBAK-AUTLF").Selected = True

' # Entrando da síntese e inserindo itens # '
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").Select
If Not ITEM10 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]").text = ITEM10
        End If
If Not ITEM20 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,1]").text = ITEM20
        End If
If Not ITEM30 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,2]").text = ITEM30
        End If
If Not ITEM40 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,3]").text = ITEM40
        End If
If Not ITEM50 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,4]").text = ITEM50
        End If
If Not ITEM60 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,5]").text = ITEM60
        End If
If Not ITEM70 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,6]").text = ITEM70
        End If
If Not ITEM80 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,7]").text = ITEM80
        End If
If Not ITEM90 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,8]").text = ITEM90
        End If
If Not QUANTIDADE10 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,0]").text = QUANTIDADE10
        End If
If Not QUANTIDADE20 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,1]").text = QUANTIDADE20
        End If
If Not QUANTIDADE30 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,2]").text = QUANTIDADE30
        End If
If Not QUANTIDADE40 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,3]").text = QUANTIDADE40
        End If
If Not QUANTIDADE50 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,4]").text = QUANTIDADE50
        End If
If Not QUANTIDADE60 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,5]").text = QUANTIDADE60
        End If
If Not QUANTIDADE70 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,6]").text = QUANTIDADE70
        End If
If Not QUANTIDADE80 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,7]").text = QUANTIDADE80
        End If
If Not QUANTIDADE90 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,8]").text = QUANTIDADE90
        End If
If Not CENTRO10 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-WERKS[13,0]").text = CENTRO10
        End If
If Not CENTRO20 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-WERKS[13,1]").text = CENTRO20
        End If
If Not CENTRO30 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-WERKS[13,2]").text = CENTRO30
        End If
If Not CENTRO40 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-WERKS[13,3]").text = CENTRO40
        End If
If Not CENTRO50 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-WERKS[13,4]").text = CENTRO50
        End If
If Not CENTRO60 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-WERKS[13,5]").text = CENTRO60
        End If
If Not CENTRO70 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-WERKS[13,6]").text = CENTRO70
        End If
If Not CENTRO80 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-WERKS[13,7]").text = CENTRO80
        End If
If Not CENTRO90 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-WERKS[13,8]").text = CENTRO90
        End If
If Not DEPÓSITO10 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-LGORT[3,0]").text = "0" & DEPÓSITO10
        End If
If Not DEPÓSITO20 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-LGORT[3,1]").text = "0" & DEPÓSITO20
        End If
If Not DEPÓSITO30 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-LGORT[3,2]").text = "0" & DEPÓSITO30
        End If
If Not DEPÓSITO40 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-LGORT[3,3]").text = "0" & DEPÓSITO40
        End If
If Not DEPÓSITO50 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-LGORT[3,4]").text = "0" & DEPÓSITO50
        End If
If Not DEPÓSITO60 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-LGORT[3,5]").text = "0" & DEPÓSITO60
        End If
If Not DEPÓSITO70 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-LGORT[3,6]").text = "0" & DEPÓSITO70
        End If
If Not DEPÓSITO80 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-LGORT[3,7]").text = "0" & DEPÓSITO80
        End If
If Not DEPÓSITO90 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-LGORT[3,8]").text = "0" & DEPÓSITO90
        End If

' # Sistema para ignorar lista técnica # '
On Error Resume Next
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[1]/usr/tblSAPMC29ACNTL/txtRC29K-STKTX[1,0]").SetFocus
session.FindById("wnd[1]").SendVKey 2
session.FindById("wnd[1]/usr/tblSAPMC29ACNTL/txtRC29K-STKTX[1,0]").SetFocus
session.FindById("wnd[1]").SendVKey 2
session.FindById("wnd[1]/usr/tblSAPMC29ACNTL/txtRC29K-STKTX[1,0]").SetFocus
session.FindById("wnd[1]").SendVKey 2
On Error GoTo 0

' # Inserindo dados no cabeçalho # '
session.FindById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/cmbVBKD-PLTYP").Key = TABELA
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\02").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4302/ctxtVBKD-VSART").text = "0" & EXPEDIÇÃO
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\05").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\05/ssubSUBSCREEN_BODY:SAPMV45A:4311/ctxtVBKD-ZLSCH").text = TIPOPAG
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,9]").Key = "ZW"
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,10]").Key = "ZA"
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,9]").text = "1001380210"
If REVENDA = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,10]").text = AGENTE
Else
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,10]").text = REVENDA
End If
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").SetSelectionIndexes 0, 0
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").SelectItem "9002", "Column1"
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").EnsureVisibleHorizontalItem "9002", "Column1"
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").DoubleClickItem "9002", "Column1"
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = texto + vbCr + ""
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").SetSelectionIndexes 0, 0
session.FindById("wnd[0]/tbar[0]/btn[3]").press

' # Inserindo valores e garantias # '
On Error Resume Next
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]").SetFocus
session.FindById("wnd[0]").SendVKey 2
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/ctxtVBAP-ZZCDGARANTIAEXT").text = GARANTIA10
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/subSUBSCREEN_PUSHBUTTONS:SAPLV69A:1000/btnBT_KONY").press
session.FindById("wnd[1]/usr/lbl[1,4]").SetFocus
session.FindById("wnd[1]").SendVKey 2
If Not OVER10 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,11]").text = OVER10
End If
If PRODUTO10 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR10SEMGARANTIA
        End If
If PRODUTO10 = "NOTEBOOK ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR10NOTEONSITE
        End If
If PRODUTO10 = "DESKTOP ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR10DESKONSITE
        End If
If PRODUTO10 = "UNION ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR10UNIONONSITE
        End If
If PRODUTO10 = "DESKTOP BALCÃO" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR10DESKBALCÃO
        End If
If PRODUTO10 = "NOTEBOOK BALCÃO" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR10NOTEBALCÃO
        End If
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/tbar[0]/btn[3]").press
If Not session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,1]").text = "20" Then
    GoTo IgnorarRestante:
        End If
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,1]").SetFocus
session.FindById("wnd[0]").SendVKey 2
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/ctxtVBAP-ZZCDGARANTIAEXT").text = GARANTIA20
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/subSUBSCREEN_PUSHBUTTONS:SAPLV69A:1000/btnBT_KONY").press
session.FindById("wnd[1]/usr/lbl[1,4]").SetFocus
session.FindById("wnd[1]").SendVKey 2
If Not OVER20 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,11]").text = OVER20
End If
If PRODUTO20 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR20SEMGARANTIA
        End If
If PRODUTO20 = "NOTEBOOK ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR20NOTEONSITE
        End If
If PRODUTO20 = "DESKTOP ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR20DESKONSITE
        End If
If PRODUTO20 = "UNION ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR20UNIONONSITE
        End If
If PRODUTO20 = "DESKTOP BALCÃO" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR20DESKBALCÃO
        End If
If PRODUTO20 = "NOTEBOOK BALCÃO" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR20NOTEBALCÃO
        End If
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/tbar[0]/btn[3]").press
If Not session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,2]").text = "30" Then
    GoTo IgnorarRestante:
        End If
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,2]").SetFocus
session.FindById("wnd[0]").SendVKey 2
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/ctxtVBAP-ZZCDGARANTIAEXT").text = GARANTIA30
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/subSUBSCREEN_PUSHBUTTONS:SAPLV69A:1000/btnBT_KONY").press
session.FindById("wnd[1]/usr/lbl[1,4]").SetFocus
session.FindById("wnd[1]").SendVKey 2
If Not OVER30 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,11]").text = OVER30
End If
If PRODUTO30 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR30SEMGARANTIA
        End If
If PRODUTO30 = "NOTEBOOK ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR30NOTEONSITE
        End If
If PRODUTO30 = "DESKTOP ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR30DESKONSITE
        End If
If PRODUTO30 = "UNION ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR30UNIONONSITE
        End If
If PRODUTO30 = "DESKTOP BALCÃO" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR30DESKBALCÃO
        End If
If PRODUTO30 = "NOTEBOOK BALCÃO" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR30NOTEBALCÃO
        End If
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/tbar[0]/btn[3]").press
If Not session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,3]").text = "40" Then
    GoTo IgnorarRestante:
        End If
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,3]").SetFocus
session.FindById("wnd[0]").SendVKey 2
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/ctxtVBAP-ZZCDGARANTIAEXT").text = GARANTIA40
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/subSUBSCREEN_PUSHBUTTONS:SAPLV69A:1000/btnBT_KONY").press
session.FindById("wnd[1]/usr/lbl[1,4]").SetFocus
session.FindById("wnd[1]").SendVKey 2
If Not OVER40 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,11]").text = OVER40
End If
If PRODUTO40 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR40SEMGARANTIA
        End If
If PRODUTO40 = "NOTEBOOK ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR40NOTEONSITE
        End If
If PRODUTO40 = "DESKTOP ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR40DESKONSITE
        End If
If PRODUTO40 = "UNION ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR40UNIONONSITE
        End If
If PRODUTO40 = "DESKTOP BALCÃO" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR40DESKBALCÃO
        End If
If PRODUTO40 = "NOTEBOOK BALCÃO" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR40NOTEBALCÃO
        End If
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/tbar[0]/btn[3]").press
If Not session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,4]").text = "50" Then
    GoTo IgnorarRestante:
        End If
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,4]").SetFocus
session.FindById("wnd[0]").SendVKey 2
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/ctxtVBAP-ZZCDGARANTIAEXT").text = GARANTIA50
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/subSUBSCREEN_PUSHBUTTONS:SAPLV69A:1000/btnBT_KONY").press
session.FindById("wnd[1]/usr/lbl[1,4]").SetFocus
session.FindById("wnd[1]").SendVKey 2
If Not OVER50 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,11]").text = OVER50
End If
If PRODUTO50 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR50SEMGARANTIA
        End If
If PRODUTO50 = "NOTEBOOK ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR50NOTEONSITE
        End If
If PRODUTO50 = "DESKTOP ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR50DESKONSITE
        End If
If PRODUTO50 = "UNION ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR50UNIONONSITE
        End If
If PRODUTO50 = "DESKTOP BALCÃO" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR50DESKBALCÃO
        End If
If PRODUTO50 = "NOTEBOOK BALCÃO" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR50NOTEBALCÃO
        End If
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/tbar[0]/btn[3]").press
If Not session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,5]").text = "60" Then
    GoTo IgnorarRestante:
        End If
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,5]").SetFocus
session.FindById("wnd[0]").SendVKey 2
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/ctxtVBAP-ZZCDGARANTIAEXT").text = GARANTIA60
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/subSUBSCREEN_PUSHBUTTONS:SAPLV69A:1000/btnBT_KONY").press
session.FindById("wnd[1]/usr/lbl[1,4]").SetFocus
session.FindById("wnd[1]").SendVKey 2
If Not OVER60 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,11]").text = OVER60
End If
If PRODUTO60 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR60SEMGARANTIA
        End If
If PRODUTO60 = "NOTEBOOK ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR60NOTEONSITE
        End If
If PRODUTO60 = "DESKTOP ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR60DESKONSITE
        End If
If PRODUTO60 = "UNION ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR60UNIONONSITE
        End If
If PRODUTO60 = "DESKTOP BALCÃO" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR60DESKBALCÃO
        End If
If PRODUTO60 = "NOTEBOOK BALCÃO" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR60NOTEBALCÃO
        End If
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/tbar[0]/btn[3]").press
If Not session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,6]").text = "70" Then
    GoTo IgnorarRestante:
        End If
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,6]").SetFocus
session.FindById("wnd[0]").SendVKey 2
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/ctxtVBAP-ZZCDGARANTIAEXT").text = GARANTIA70
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/subSUBSCREEN_PUSHBUTTONS:SAPLV69A:1000/btnBT_KONY").press
session.FindById("wnd[1]/usr/lbl[1,4]").SetFocus
session.FindById("wnd[1]").SendVKey 2
If Not OVER70 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,11]").text = OVER70
End If
If PRODUTO70 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR70SEMGARANTIA
        End If
If PRODUTO70 = "NOTEBOOK ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR70NOTEONSITE
        End If
If PRODUTO70 = "DESKTOP ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR70DESKONSITE
        End If
If PRODUTO70 = "UNION ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR70UNIONONSITE
        End If
If PRODUTO70 = "DESKTOP BALCÃO" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR70DESKBALCÃO
        End If
If PRODUTO70 = "NOTEBOOK BALCÃO" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR70NOTEBALCÃO
        End If
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/tbar[0]/btn[3]").press
If Not session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,7]").text = "80" Then
    GoTo IgnorarRestante:
        End If
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,7]").SetFocus
session.FindById("wnd[0]").SendVKey 2
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/ctxtVBAP-ZZCDGARANTIAEXT").text = GARANTIA80
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/subSUBSCREEN_PUSHBUTTONS:SAPLV69A:1000/btnBT_KONY").press
session.FindById("wnd[1]/usr/lbl[1,4]").SetFocus
session.FindById("wnd[1]").SendVKey 2
If Not OVER80 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,11]").text = OVER80
End If
If PRODUTO80 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR80SEMGARANTIA
        End If
If PRODUTO80 = "NOTEBOOK ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR80NOTEONSITE
        End If
If PRODUTO80 = "DESKTOP ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR80DESKONSITE
        End If
If PRODUTO80 = "UNION ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR80UNIONONSITE
        End If
If PRODUTO80 = "DESKTOP BALCÃO" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR80DESKBALCÃO
        End If
If PRODUTO80 = "NOTEBOOK BALCÃO" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR80NOTEBALCÃO
        End If
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/tbar[0]/btn[3]").press
If Not session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,8]").text = "90" Then
    GoTo IgnorarRestante:
        End If
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,8]").SetFocus
session.FindById("wnd[0]").SendVKey 2
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/ctxtVBAP-ZZCDGARANTIAEXT").text = GARANTIA90
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/subSUBSCREEN_PUSHBUTTONS:SAPLV69A:1000/btnBT_KONY").press
session.FindById("wnd[1]/usr/lbl[1,4]").SetFocus
session.FindById("wnd[1]").SendVKey 2
If Not OVER90 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,11]").text = OVER90
End If
If PRODUTO90 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR90SEMGARANTIA
        End If
If PRODUTO90 = "NOTEBOOK ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR90NOTEONSITE
        End If
If PRODUTO90 = "DESKTOP ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR90DESKONSITE
        End If
If PRODUTO90 = "UNION ON SITE" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR90UNIONONSITE
        End If
If PRODUTO90 = "DESKTOP BALCÃO" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR90DESKBALCÃO
        End If
If PRODUTO90 = "NOTEBOOK BALCÃO" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = VALOR90NOTEBALCÃO
        End If
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/tbar[0]/btn[3]").press
IgnorarRestante:
On Error GoTo 0

' # Colocando comissão # '
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]").SetFocus
session.FindById("wnd[0]").SendVKey 2
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/btnPB_ADD").press
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/btnPB_ADD").press
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/btnPB_ADD").press
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/btnPB_ADD").press
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/btnPB_ADD").press
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/btnPB_ADD").press
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/btnPB_ADD").press
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/cmbTG_TABCOM-PARVW[0,0]").Key = "Z2"
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/cmbTG_TABCOM-PARVW[0,1]").Key = "Z5"
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/cmbTG_TABCOM-PARVW[0,2]").Key = "Z6"
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/cmbTG_TABCOM-PARVW[0,3]").Key = "Z7"
If Not Z3 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/cmbTG_TABCOM-PARVW[0,4]").Key = "Z3"
End If
If Not Z1 = "" Then
    If Z3 = "" Then
        session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/cmbTG_TABCOM-PARVW[0,4]").Key = "Z1"
    Else
        session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/cmbTG_TABCOM-PARVW[0,5]").Key = "Z1"
    End If
End If
If Not Z4 = "" Then
    If Z3 = "" Then
        session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/cmbTG_TABCOM-PARVW[0,4]").Key = "Z4"
    Else
        If Z1 = "" Then
            session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/cmbTG_TABCOM-PARVW[0,5]").Key = "Z4"
        Else
            session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/cmbTG_TABCOM-PARVW[0,6]").Key = "Z4"
        End If
    End If
End If
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/ctxtTG_TABCOM-LIFNR[1,0]").text = Z2
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/ctxtTG_TABCOM-LIFNR[1,1]").text = Z5
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/ctxtTG_TABCOM-LIFNR[1,2]").text = Z6
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/ctxtTG_TABCOM-LIFNR[1,3]").text = Z7
If Not Z3 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/ctxtTG_TABCOM-LIFNR[1,4]").text = Z3
End If
If Not Z1 = "" Then
    If Z3 = "" Then
        session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/ctxtTG_TABCOM-LIFNR[1,4]").text = Z1
    Else
        session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/ctxtTG_TABCOM-LIFNR[1,5]").text = Z1
    End If
End If
If Not Z4 = "" Then
    If Z3 = "" Then
        session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/ctxtTG_TABCOM-LIFNR[1,4]").text = Z4
    Else
        If Z1 = "" Then
            session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/ctxtTG_TABCOM-LIFNR[1,5]").text = Z4
        Else
            session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/ctxtTG_TABCOM-LIFNR[1,6]").text = Z4
        End If
    End If
End If
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/txtTG_TABCOM-KBETR[3,0]").text = Z2VALOR
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/txtTG_TABCOM-KBETR[3,1]").text = Z5VALOR
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/txtTG_TABCOM-KBETR[3,2]").text = Z6VALOR
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/txtTG_TABCOM-KBETR[3,3]").text = Z7VALOR
If Not Z3 = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/txtTG_TABCOM-KBETR[3,4]").text = Z3VALOR
End If
If Not Z1 = "" Then
    If Z3 = "" Then
        session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/txtTG_TABCOM-KBETR[3,4]").text = Z1VALOR
    Else
        session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/txtTG_TABCOM-KBETR[3,5]").text = Z1VALOR
    End If
End If
If Not Z4 = "" Then
    If Z3 = "" Then
        session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/txtTG_TABCOM-KBETR[3,4]").text = Z4VALOR
    Else
        If Z1 = "" Then
            session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/txtTG_TABCOM-KBETR[3,5]").text = Z4VALOR
        Else
            session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/tblSAPMV45ATC_TABCOMISS/txtTG_TABCOM-KBETR[3,6]").text = Z4VALOR
        End If
    End If
End If
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\15/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/btnBT_REPL_COMISS").press
session.FindById("wnd[1]/usr/btnBUTTON_1").press
session.FindById("wnd[1]/tbar[0]/btn[0]").press
session.FindById("wnd[0]/tbar[0]/btn[3]").press
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select

' # Verificando se valor final da ZCOT bate e corrigindo se necessário # '
If DOC = "ZCOT" Then
    Dim zcotMount As String
    Dim zcotMountWS As String
    ThisWorkbook.Worksheets("DADOS").Cells(15, "F") = session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/txtRV45A-ENDBK").text
    zcotMountWS = ThisWorkbook.Worksheets("DADOS").Cells(15, "F").text
    zcotMount = removeWS(zcotMountWS)
    finalMount = ThisWorkbook.Worksheets("DADOS").Cells(15, "G").text
    If zcotMount <> finalMount Then
        Do
            If ThisWorkbook.Worksheets("DADOS").Cells(16, 6).Value < 0.1 And ThisWorkbook.Worksheets("DADOS").Cells(16, 6).Value > 0 Then
                session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").Select
                session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]").SetFocus
                session.FindById("wnd[0]").SendVKey 2
                session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06").Select
                session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN").VerticalScrollbar.Position = 74
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,1]").text = "Resid. Arred. Manual" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,1]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    theValue = ThisWorkbook.Worksheets("DADOS").Cells(15, "H").text
                    correctValue = Replace(theValue, "-", "")
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = correctValue
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(18, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,2]").text = "Resid. Arred. Manual" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,2]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    theValue = ThisWorkbook.Worksheets("DADOS").Cells(15, "H").text
                    correctValue = Replace(theValue, "-", "")
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = correctValue
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(18, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,3]").text = "Resid. Arred. Manual" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,3]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    theValue = ThisWorkbook.Worksheets("DADOS").Cells(15, "H").text
                    correctValue = Replace(theValue, "-", "")
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = correctValue
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(18, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,4]").text = "Resid. Arred. Manual" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,4]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    theValue = ThisWorkbook.Worksheets("DADOS").Cells(15, "H").text
                    correctValue = Replace(theValue, "-", "")
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = correctValue
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(18, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,5]").text = "Resid. Arred. Manual" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,5]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    theValue = ThisWorkbook.Worksheets("DADOS").Cells(15, "H").text
                    correctValue = Replace(theValue, "-", "")
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = correctValue
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(18, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,6]").text = "Resid. Arred. Manual" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,6]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    theValue = ThisWorkbook.Worksheets("DADOS").Cells(15, "H").text
                    correctValue = Replace(theValue, "-", "")
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = correctValue
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(18, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,7]").text = "Resid. Arred. Manual" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,7]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    theValue = ThisWorkbook.Worksheets("DADOS").Cells(15, "H").text
                    correctValue = Replace(theValue, "-", "")
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = correctValue
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(18, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,8]").text = "Resid. Arred. Manual" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,8]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    theValue = ThisWorkbook.Worksheets("DADOS").Cells(15, "H").text
                    correctValue = Replace(theValue, "-", "")
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = correctValue
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(18, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,9]").text = "Resid. Arred. Manual" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,9]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    theValue = ThisWorkbook.Worksheets("DADOS").Cells(15, "H").text
                    correctValue = Replace(theValue, "-", "")
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = correctValue
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(18, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
            End If
            If ThisWorkbook.Worksheets("DADOS").Cells(16, 6).Value > -0.1 And ThisWorkbook.Worksheets("DADOS").Cells(16, 6).Value < 0 Then
                session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").Select
                session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]").SetFocus
                session.FindById("wnd[0]").SendVKey 2
                session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06").Select
                session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN").VerticalScrollbar.Position = 74
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,1]").text = "Resid. Arred. Manual" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,1]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(17, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(17, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,2]").text = "Resid. Arred. Manual" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,2]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(17, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(17, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,3]").text = "Resid. Arred. Manual" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,3]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(17, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(17, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,4]").text = "Resid. Arred. Manual" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,4]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(17, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(17, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,5]").text = "Resid. Arred. Manual" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,5]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(17, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(17, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,6]").text = "Resid. Arred. Manual" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,6]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(17, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(17, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,7]").text = "Resid. Arred. Manual" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,7]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(17, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(17, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,8]").text = "Resid. Arred. Manual" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,8]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(17, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(17, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,9]").text = "Resid. Arred. Manual" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,9]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(17, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(17, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
            End If
            If ThisWorkbook.Worksheets("DADOS").Cells(16, 6).Value > 0.1 Then
                session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").Select
                session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]").SetFocus
                session.FindById("wnd[0]").SendVKey 2
                session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06").Select
                session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN").VerticalScrollbar.Position = 5
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,1]").text = "Desconto Manual(Vlr)" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,1]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    theValue = ThisWorkbook.Worksheets("DADOS").Cells(15, "H").text
                    correctValue = Replace(theValue, "-", "")
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = correctValue
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(18, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,2]").text = "Desconto Manual(Vlr)" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,2]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    theValue = ThisWorkbook.Worksheets("DADOS").Cells(15, "H").text
                    correctValue = Replace(theValue, "-", "")
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = correctValue
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(18, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,3]").text = "Desconto Manual(Vlr)" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,3]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    theValue = ThisWorkbook.Worksheets("DADOS").Cells(15, "H").text
                    correctValue = Replace(theValue, "-", "")
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = correctValue
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(18, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,4]").text = "Desconto Manual(Vlr)" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,4]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    theValue = ThisWorkbook.Worksheets("DADOS").Cells(15, "H").text
                    correctValue = Replace(theValue, "-", "")
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = correctValue
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(18, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,5]").text = "Desconto Manual(Vlr)" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,5]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    theValue = ThisWorkbook.Worksheets("DADOS").Cells(15, "H").text
                    correctValue = Replace(theValue, "-", "")
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = correctValue
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(18, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,6]").text = "Desconto Manual(Vlr)" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,6]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    theValue = ThisWorkbook.Worksheets("DADOS").Cells(15, "H").text
                    correctValue = Replace(theValue, "-", "")
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = correctValue
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(18, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,7]").text = "Desconto Manual(Vlr)" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,7]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    theValue = ThisWorkbook.Worksheets("DADOS").Cells(15, "H").text
                    correctValue = Replace(theValue, "-", "")
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = correctValue
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(18, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,8]").text = "Desconto Manual(Vlr)" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,8]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    theValue = ThisWorkbook.Worksheets("DADOS").Cells(15, "H").text
                    correctValue = Replace(theValue, "-", "")
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = correctValue
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(18, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,9]").text = "Desconto Manual(Vlr)" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,9]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    theValue = ThisWorkbook.Worksheets("DADOS").Cells(15, "H").text
                    correctValue = Replace(theValue, "-", "")
                    ThisWorkbook.Worksheets("DADOS").Cells(15, "H") = correctValue
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(18, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
            End If
            If ThisWorkbook.Worksheets("DADOS").Cells(16, 6).Value < -0.1 Then
                session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").Select
                session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]").SetFocus
                session.FindById("wnd[0]").SendVKey 2
                session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06").Select
                session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN").VerticalScrollbar.Position = 5
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,1]").text = "AcréscimoManual(Vlr)" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,1]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(17, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(17, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,2]").text = "AcréscimoManual(Vlr))" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,2]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(17, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(17, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,3]").text = "AcréscimoManual(Vlr)" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,3]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(17, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(17, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,4]").text = "AcréscimoManual(Vlr)" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,4]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(17, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(17, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,5]").text = "AcréscimoManual(Vlr)" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,5]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(17, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(17, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,6]").text = "AcréscimoManual(Vlr)" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,6]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(17, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(17, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,7]").text = "AcréscimoManual(Vlr)" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,7]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(17, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(17, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,8]").text = "AcréscimoManual(Vlr)" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,8]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(17, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(17, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,9]").text = "AcréscimoManual(Vlr)" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,9]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(17, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(17, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,10]").text = "AcréscimoManual(Vlr)" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,10]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(17, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(17, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,11]").text = "AcréscimoManual(Vlr)" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,11]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(17, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(17, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
                If session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,12]").text = "AcréscimoManual(Vlr)" Then
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtT685T-VTEXT[2,12]").SetFocus
                    session.FindById("wnd[0]").SendVKey 2
                    ThisWorkbook.Worksheets("DADOS").Cells(17, "H") = session.FindById("wnd[0]/usr/txtKOMV-KBETR").text
                    session.FindById("wnd[0]/usr/txtKOMV-KBETR").text = ThisWorkbook.Worksheets("DADOS").Cells(17, "I").text
                    session.FindById("wnd[0]").SendVKey 0
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01").Select
                    GoTo SkipThis
                End If
            End If
SkipThis:
        ThisWorkbook.Worksheets("DADOS").Cells(15, "F") = session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/txtRV45A-ENDBK").text
        zcotMountWS = ThisWorkbook.Worksheets("DADOS").Cells(15, "F").text
        zcotMount = removeWS(zcotMountWS)
        If zcotMount = finalMount Then
            Exit Do
        End If
        Loop
    End If
End If

' # Salvando ordem/cotação # '
On Error Resume Next
session.FindById("wnd[0]/tbar[0]/btn[11]").press
session.FindById("wnd[1]/tbar[0]/btn[0]").press
session.FindById("wnd[1]/usr/btnSPOP-VAROPTION1").press
On Error GoTo 0

' # Inserindo nº do doc na planilha # '
session.FindById("wnd[0]/tbar[0]/okcd").text = TRANSAÇÃOMODIFICAR
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBAK-VBELN").SetFocus
ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(3, "W") = session.FindById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBAK-VBELN").text
nDOC = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(3, "W").Value

Set session = Nothing
MsgBox ("Documento criado com sucesso: " & nDOC)

End Sub