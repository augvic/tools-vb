Public SapGuiAuto
Public objGui As GuiApplication
Public objConn As GuiConnection
Public session As GuiSession

Sub CADASTRO_SAP()

' # Variáveis # '
Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)
Dim grupo As String
Dim EMPRESA As String
Dim organizaçãocadastro As String
Dim canalcadastro As String
Dim escritóriocadastro As String
Dim equipecadastro As String
Dim hierarquia As String
Dim centro As String
Dim tratamento As String
Dim razão1 As String
Dim razão2 As String
Dim fantasia1 As String
Dim fantasia2 As String
Dim pesquisa1 As String
Dim pesquisa2 As String
Dim cnpj As String
Dim ie As String
Dim cpf As String
Dim suframa As String
Dim rua As String
Dim número As String
Dim complemento1 As String
Dim complemento2 As String
Dim bairro As String
Dim cep As String
Dim cidade As String
Dim estado As String
Dim setorindustrial As String
Dim domiciliointeiro As String
Dim domicilioresumido As String
Dim telefone As String
Dim emailcontato As String
Dim emailfatura As String

' # Atribuição de valores # '
grupo = ThisWorkbook.Worksheets("CADASTRO").Cells(3, 4).text
EMPRESA = ThisWorkbook.Worksheets("CADASTRO").Cells(4, 4).text
organizaçãocadastro = ThisWorkbook.Worksheets("CADASTRO").Cells(5, 4).text
canalcadastro = ThisWorkbook.Worksheets("CADASTRO").Cells(6, 4).text
escritóriocadastro = ThisWorkbook.Worksheets("CADASTRO").Cells(7, 4).text
equipecadastro = ThisWorkbook.Worksheets("CADASTRO").Cells(8, 4).text
hierarquia = ThisWorkbook.Worksheets("DADOS CADASTRO").Cells(2, 8).text
centro = ThisWorkbook.Worksheets("DADOS CADASTRO").Cells(2, 12).text

tratamento = ThisWorkbook.Worksheets("DADOS CADASTRO").Cells(1, "Q").text
razão1 = ThisWorkbook.Worksheets("CADASTRO").Cells(4, 7).text
razão2 = ThisWorkbook.Worksheets("CADASTRO").Cells(5, 7).text
fantasia1 = ThisWorkbook.Worksheets("CADASTRO").Cells(6, 7).text
fantasia2 = ThisWorkbook.Worksheets("CADASTRO").Cells(7, 7).text
pesquisa1 = ThisWorkbook.Worksheets("CADASTRO").Cells(8, 7).text
pesquisa2 = ThisWorkbook.Worksheets("CADASTRO").Cells(9, 10).text
cnpj = ThisWorkbook.Worksheets("CADASTRO").Cells(3, 10).text
cpf = ThisWorkbook.Worksheets("CADASTRO").Cells(3, 12).text
ie = ThisWorkbook.Worksheets("CADASTRO").Cells(12, 14).text
suframa = ThisWorkbook.Worksheets("CADASTRO").Cells(15, "O").text
tipo = ThisWorkbook.Worksheets("CADASTRO").Cells(9, 4).text

rua = ThisWorkbook.Worksheets("CADASTRO").Cells(10, "G").text
rua2 = ThisWorkbook.Worksheets("CADASTRO").Cells(11, "G").text
número = ThisWorkbook.Worksheets("CADASTRO").Cells(10, "J").text
complemento1 = ThisWorkbook.Worksheets("CADASTRO").Cells(10, "L").text
complemento2 = ThisWorkbook.Worksheets("CADASTRO").Cells(12, "G").text
bairro = ThisWorkbook.Worksheets("CADASTRO").Cells(13, "G").text
cep = ThisWorkbook.Worksheets("CADASTRO").Cells(14, "G").text
cidade = ThisWorkbook.Worksheets("CADASTRO").Cells(13, "K").text
estado = ThisWorkbook.Worksheets("CADASTRO").Cells(13, "I").text
setorindustrial = ThisWorkbook.Worksheets("DADOS CADASTRO").Cells(2, 4).text

telefone = ThisWorkbook.Worksheets("CADASTRO").Cells(16, 7).text
emailcontato = ThisWorkbook.Worksheets("CADASTRO").Cells(17, 7).text
emailfatura = ThisWorkbook.Worksheets("CADASTRO").Cells(18, 7).text

' # Checando se todos os campos estão preenchidos. # '
If grupo = "" Or EMPRESA = "" Or organizaçãocadastro = "" Or canalcadastro = "" Or escritóriocadastro = "" _
Or equipecadastro = "" Or hierarquia = "" Or tratamento = "" Or razão1 = "" _
Or pesquisa1 = "" Or rua = "" Or número = "" Or bairro = "" Or cep = "" Or estado = "" _
Or cidade = "" Or telefone = "" Or ie = "" Then
    MsgBox ("Preencha todos os campos para cadastrar.")
    Set session = Nothing
    Exit Sub
End If

If grupo = "ZPPJ" Or grupo = "ZAGE" Or grupo = "ZPOP" Then
    If cnpj = "" Then
        MsgBox ("Preencha todos os campos para cadastrar.")
        Set session = Nothing
        Exit Sub
    End If
End If
If grupo = "ZPPF" Then
    If cpf = "" Then
        MsgBox ("Preencha todos os campos para cadastrar.")
        Set session = Nothing
        Exit Sub
    End If
End If

' # Dados # '
session.FindById("wnd[0]").maximize
session.FindById("wnd[0]/tbar[0]/okcd").text = "/NXD01"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[1]/usr/ctxtRF02D-REF_KUNNR").text = ""
session.FindById("wnd[1]/usr/ctxtRF02D-REF_BUKRS").text = ""
session.FindById("wnd[1]/usr/ctxtRF02D-REF_VKORG").text = ""
session.FindById("wnd[1]/usr/ctxtRF02D-REF_VTWEG").text = ""
session.FindById("wnd[1]/usr/ctxtRF02D-REF_SPART").text = ""
session.FindById("wnd[1]/usr/cmbRF02D-KTOKD").Key = grupo
session.FindById("wnd[1]/usr/ctxtRF02D-BUKRS").text = EMPRESA
session.FindById("wnd[1]/usr/ctxtRF02D-VKORG").text = organizaçãocadastro
session.FindById("wnd[1]/usr/ctxtRF02D-VTWEG").text = canalcadastro
session.FindById("wnd[1]/usr/ctxtRF02D-SPART").text = "00"
session.FindById("wnd[1]").SendVKey 0

' # Razão e Endereço # '
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/cmbSZA1_D0100-TITLE_MEDI").Key = tratamento
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME1").text = razão1
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME2").text = razão2
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-SORT1").text = pesquisa1
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-SORT2").text = pesquisa2
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-STREET").text = rua
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-HOUSE_NUM1").text = número
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-CITY2").text = bairro
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-POST_CODE1").text = cep
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-CITY1").text = cidade
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-COUNTRY").text = "BR"
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-REGION").text = estado
On Error GoTo erroSetorIndustrial
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-TRANSPZONE").text = setorindustrial
On Error GoTo 0

' # Domicílio # '
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/ctxtADDR1_DATA-TAXJURCODE").SetFocus
session.FindById("wnd[0]").SendVKey 4
session.FindById("wnd[1]/tbar[0]/btn[0]").press
If Not ThisWorkbook.Worksheets("CADASTRO").Cells(14, 15) = "" Then
    ThisWorkbook.Worksheets("DADOS CADASTRO").Cells(2, 23) = session.FindById("wnd[1]/usr/lbl[88,4]").text
    If ThisWorkbook.Worksheets("DADOS CADASTRO").Cells(3, 23) = "ZF" Then
        session.FindById("wnd[1]/usr/lbl[88,4]").SetFocus
        session.FindById("wnd[1]").SendVKey 2
        GoTo SkipSuframa
    Else
        ThisWorkbook.Worksheets("DADOS CADASTRO").Cells(2, 23) = session.FindById("wnd[1]/usr/lbl[88,5]").text
        If ThisWorkbook.Worksheets("DADOS CADASTRO").Cells(3, 23) = "ZF" Then
            session.FindById("wnd[1]/usr/lbl[88,5]").SetFocus
            session.FindById("wnd[1]").SendVKey 2
            GoTo SkipSuframa
        Else
            ThisWorkbook.Worksheets("DADOS CADASTRO").Cells(2, 23) = session.FindById("wnd[1]/usr/lbl[88,6]").text
            If ThisWorkbook.Worksheets("DADOS CADASTRO").Cells(3, 23) = "ZF" Then
                session.FindById("wnd[1]/usr/lbl[88,6]").SetFocus
                session.FindById("wnd[1]").SendVKey 2
                GoTo SkipSuframa
            End If
        End If
    End If
End If

SkipSuframa:
If Not ie = "ISENTO" Then
    session.FindById("wnd[1]/usr/lbl[88,3]").SetFocus
    session.FindById("wnd[1]").SendVKey 2
Else
    ThisWorkbook.Worksheets("DADOS CADASTRO").Cells(2, 21) = session.FindById("wnd[1]/usr/lbl[88,4]").text
    If ThisWorkbook.Worksheets("DADOS CADASTRO").Cells(3, 21) = "Z" Then
        session.FindById("wnd[1]/usr/lbl[88,4]").SetFocus
        session.FindById("wnd[1]").SendVKey 2
    Else
        ThisWorkbook.Worksheets("DADOS CADASTRO").Cells(2, 21) = session.FindById("wnd[1]/usr/lbl[88,5]").text
        If ThisWorkbook.Worksheets("DADOS CADASTRO").Cells(3, 21) = "Z" Then
            session.FindById("wnd[1]/usr/lbl[88,5]").SetFocus
            session.FindById("wnd[1]").SendVKey 2
        Else
            ThisWorkbook.Worksheets("DADOS CADASTRO").Cells(2, 21) = session.FindById("wnd[1]/usr/lbl[88,6]").text
            If ThisWorkbook.Worksheets("DADOS CADASTRO").Cells(3, 21) = "Z" Then
                session.FindById("wnd[1]/usr/lbl[88,6]").SetFocus
                session.FindById("wnd[1]").SendVKey 2
            End If
        End If
    End If
End If

' # Contato # '
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtSZA1_D0100-TEL_NUMBER").text = telefone
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtSZA1_D0100-FAX_NUMBER").text = telefone
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtSZA1_D0100-SMTP_ADDR").text = emailcontato

' # Complementos # '
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/btnG_D0100_DUMMY_TRANSPZONE").press
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-STR_SUPPL3").text = rua2
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-HOUSE_NUM2").text = complemento1
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-LOCATION").text = complemento2
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/btnG_D0100_DUMMY_NAME2").press
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME3").text = fantasia1
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/txtADDR1_DATA-NAME4").text = fantasia2

' # Dados de controle # '
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").Select
If grupo = "ZPPJ" Then
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA2:SAPMF02D:7123/ctxtKNA1-BRSCH").text = "0006"
        End If
If grupo = "ZAGE" Then
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA2:SAPMF02D:7123/ctxtKNA1-BRSCH").text = "0006"
        End If
If grupo = "ZPPF" Then
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA2:SAPMF02D:7123/ctxtKNA1-BRSCH").text = "0010"
        End If
If tipo = "ÓRGÃO FEDERAL" Then
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA2:SAPMF02D:7123/ctxtKNA1-BRSCH").text = "0004"
        End If
If tipo = "ÓRGÃO ESTADUAL" Then
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA2:SAPMF02D:7123/ctxtKNA1-BRSCH").text = "0011"
        End If
If tipo = "ÓRGÃO MUNICIPAL" Then
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA2:SAPMF02D:7123/ctxtKNA1-BRSCH").text = "0013"
        End If
If grupo = "ZPPF" Then
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPMF02D:7122/txtKNA1-STCD2").text = cpf
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPMF02D:7122/chkKNA1-STKZN").Selected = True
Else
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPMF02D:7122/txtKNA1-STCD1").text = cnpj
End If
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPMF02D:7122/txtKNA1-STCD3").text = ie
If Not ie = "ISENTO" Then
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPMF02D:7122/ctxtKNA1-XSUBT").text = "Z"
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPMF02D:7122/ctxtKNA1-CFOPC").text = "0"
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPMF02D:7122/ctxtKNA1-ICMSTAXPAY").text = "ZA"
        End If
If ie = "ISENTO" Then
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPMF02D:7122/ctxtKNA1-XSUBT").text = "X"
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPMF02D:7122/ctxtKNA1-CFOPC").text = "6"
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPMF02D:7122/ctxtKNA1-ICMSTAXPAY").text = "ZC"
        End If
If grupo = "ZPPF" Then
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPMF02D:7122/ctxtKNA1-ICMSTAXPAY").text = "ZE"
        End If
If Not suframa = "" Then
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPMF02D:7122/txtKNA1-SUFRAMA").text = suframa
End If
session.FindById("wnd[0]/tbar[1]/btn[26]").press

' # Administração de conta # '
If Not grupo = "ZPOP" Then
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0202/subAREA1:SAPMF02D:7211/ctxtKNB1-AKONT").text = "1102001001"
        End If
If grupo = "ZPOP" Then
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0202/subAREA1:SAPMF02D:7211/ctxtKNB1-AKONT").text = "1102001003"
        End If
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0202/subAREA1:SAPMF02D:7211/ctxtKNB1-ZUAWA").text = "001"
If Not grupo = "ZPOP" Then
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0202/subAREA1:SAPMF02D:7211/ctxtKNB1-FDGRV").text = "E1"
        End If
If grupo = "ZPOP" Then
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0202/subAREA1:SAPMF02D:7211/ctxtKNB1-FDGRV").text = "E2"
        End If
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0202/subAREA2:SAPMF02D:7212/ctxtKNB1-VZSKZ").text = "01"
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").Select

' # Pagamentos # '
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0202/subAREA1:SAPMF02D:7215/chkKNB1-XZVER").Selected = True
session.FindById("wnd[0]/tbar[1]/btn[27]").press

' # Área de vendas - Vendas # '
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPMF02D:7310/ctxtKNVV-VKBUR").text = escritóriocadastro
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPMF02D:7310/ctxtKNVV-VKGRP").text = equipecadastro
If grupo = "ZPPJ" Then
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPMF02D:7310/ctxtKNVV-KDGRP").text = "03"
        End If
If grupo = "ZAGE" Then
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPMF02D:7310/ctxtKNVV-KDGRP").text = "03"
        End If
If grupo = "ZPOP" Then
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPMF02D:7310/ctxtKNVV-KDGRP").text = "08"
        End If
If grupo = "ZPPF" Then
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPMF02D:7310/ctxtKNVV-KDGRP").text = "01"
        End If
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA2:SAPMF02D:7311/ctxtKNVV-KALKS").text = "1"
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA2:SAPMF02D:7311/ctxtKNVV-VERSG").text = "1"
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPMF02D:7313/ctxtKNVH-HITYP").text = "A"
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPMF02D:7313/ctxtLINK_KNVH-HKUNNR").text = hierarquia
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").Select

' # Área de vendas - Expedição # '
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPMF02D:7315/ctxtKNVV-LPRIO").text = "02"
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPMF02D:7315/ctxtKNVV-VSBED").text = "Z1"
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPMF02D:7315/ctxtKNVV-VWERK").text = centro
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").Select

' # Área de vendas - Doc.faturamento # '
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPMF02D:7320/chkKNVV-BOKRE").Selected = True
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPMF02D:7320/chkKNVV-PRFRE").Selected = True
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0200/subAREA2:SAPMF02D:7321/ctxtKNVV-INCO1").text = "CIF"
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0200/subAREA2:SAPMF02D:7321/txtKNVV-INCO2").text = cidade
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPMF02D:7322/ctxtKNVV-KTGRD").text = "01"
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:SAPLATAB:0200/subAREA4:SAPMF02D:7323/subSUB_STEUER:SAPMF02D:7350/tblSAPMF02DTCTRL_STEUERN/ctxtKNVI-TAXKD[4,0]").text = "1"

' # Imposto Ret. # '
If grupo = "ZPOP" Then
    session.FindById("wnd[0]/tbar[1]/btn[26]").press
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05").Select
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7610/subWT_WITS:SAPMFWTC:0610/tblSAPMFWTCTCTRL_QUELLENSTEUER/chkKNBW-WT_AGENT[2,0]").Selected = True
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7610/subWT_WITS:SAPMFWTC:0610/tblSAPMFWTCTCTRL_QUELLENSTEUER/chkKNBW-WT_AGENT[2,1]").Selected = True
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7610/subWT_WITS:SAPMFWTC:0610/tblSAPMFWTCTCTRL_QUELLENSTEUER/chkKNBW-WT_AGENT[2,2]").Selected = True
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7610/subWT_WITS:SAPMFWTC:0610/tblSAPMFWTCTCTRL_QUELLENSTEUER/chkKNBW-WT_AGENT[2,3]").Selected = True
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7610/subWT_WITS:SAPMFWTC:0610/tblSAPMFWTCTCTRL_QUELLENSTEUER/ctxtKNBW-WITHT[0,0]").text = "CB"
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7610/subWT_WITS:SAPMFWTC:0610/tblSAPMFWTCTCTRL_QUELLENSTEUER/ctxtKNBW-WITHT[0,1]").text = "PB"
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7610/subWT_WITS:SAPMFWTC:0610/tblSAPMFWTCTCTRL_QUELLENSTEUER/ctxtKNBW-WITHT[0,2]").text = "RB"
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7610/subWT_WITS:SAPMFWTC:0610/tblSAPMFWTCTCTRL_QUELLENSTEUER/ctxtKNBW-WITHT[0,3]").text = "SB"
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7610/subWT_WITS:SAPMFWTC:0610/tblSAPMFWTCTCTRL_QUELLENSTEUER/ctxtKNBW-WT_WITHCD[1,0]").text = "CB"
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7610/subWT_WITS:SAPMFWTC:0610/tblSAPMFWTCTCTRL_QUELLENSTEUER/ctxtKNBW-WT_WITHCD[1,1]").text = "PB"
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7610/subWT_WITS:SAPMFWTC:0610/tblSAPMFWTCTCTRL_QUELLENSTEUER/ctxtKNBW-WT_WITHCD[1,2]").text = "RB"
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7610/subWT_WITS:SAPMFWTC:0610/tblSAPMFWTCTCTRL_QUELLENSTEUER/ctxtKNBW-WT_WITHCD[1,3]").text = "SB"
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7610/subWT_WITS:SAPMFWTC:0610/tblSAPMFWTCTCTRL_QUELLENSTEUER/ctxtKNBW-WT_AGTDF[3,0]").text = "01.01.2009"
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7610/subWT_WITS:SAPMFWTC:0610/tblSAPMFWTCTCTRL_QUELLENSTEUER/ctxtKNBW-WT_AGTDF[3,1]").text = "01.01.2009"
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7610/subWT_WITS:SAPMFWTC:0610/tblSAPMFWTCTCTRL_QUELLENSTEUER/ctxtKNBW-WT_AGTDF[3,2]").text = "01.01.2009"
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7610/subWT_WITS:SAPMFWTC:0610/tblSAPMFWTCTCTRL_QUELLENSTEUER/ctxtKNBW-WT_AGTDF[3,3]").text = "01.01.2009"
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7610/subWT_WITS:SAPMFWTC:0610/tblSAPMFWTCTCTRL_QUELLENSTEUER/ctxtKNBW-WT_AGTDT[4,0]").text = "30.12.9999"
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7610/subWT_WITS:SAPMFWTC:0610/tblSAPMFWTCTCTRL_QUELLENSTEUER/ctxtKNBW-WT_AGTDT[4,1]").text = "30.12.9999"
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7610/subWT_WITS:SAPMFWTC:0610/tblSAPMFWTCTCTRL_QUELLENSTEUER/ctxtKNBW-WT_AGTDT[4,2]").text = "30.12.9999"
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7610/subWT_WITS:SAPMFWTC:0610/tblSAPMFWTCTCTRL_QUELLENSTEUER/ctxtKNBW-WT_AGTDT[4,3]").text = "30.12.9999"
End If

' # E-mail fatura # '
session.FindById("wnd[0]/tbar[1]/btn[28]").press
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPLZSD_VEN_011:9001/txtKNA1-ZZEMAILNFE").text = emailfatura
If Not suframa = "" Then
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPLZSD_VEN_011:9001/txtKNA1-ZZSUFRAMA").text = suframa
End If
session.FindById("wnd[0]/tbar[1]/btn[25]").press

' ~~ Se for contribuinte, flega atributo 4. ~~ '
If Not ie = "ISENTO" Then
    session.FindById("wnd[0]/mbar/menu[3]/menu[8]").Select
    session.FindById("wnd[1]/usr/ctxtKNA1-KATR4").text = "X"
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
End If

' ~~ Salvando cliente novo e inserindo nº na planilha. ~~ '
session.FindById("wnd[0]/tbar[0]/btn[11]").press
session.FindById("wnd[0]/tbar[0]/okcd").text = "/NXD02"
session.FindById("wnd[0]").SendVKey 0
ThisWorkbook.Worksheets("CADASTRO").Cells(11, 3) = session.FindById("wnd[1]/usr/ctxtRF02D-KUNNR").text
nDOC = ThisWorkbook.Worksheets("CADASTRO").Cells(11, 3).Value

Set session = Nothing
MsgBox ("Cliente criado com sucesso: " & nDOC)

erroSetorIndustrial:
MsgBox ("Erro ao inserir Setor Industrial. Verificar estado e cidade.")
Exit Sub

End Sub