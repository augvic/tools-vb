Sub ampliação_cliente_cadastro()

' # Variáveis # '
Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)
Dim EMISSOR As String
Dim ORGANIZAÇÃO As String
Dim CANAL As String
Dim ESCRITÓRIO As String
Dim EQUIPE As String
Dim centro As String
Dim EMPRESA As String

' # Atribuindo # '
EMPRESA = ThisWorkbook.Worksheets("CADASTRO").Cells(4, 4).text
EMISSOR = ThisWorkbook.Worksheets("CADASTRO").Cells(11, 3).text
ORGANIZAÇÃO = ThisWorkbook.Worksheets("CADASTRO").Cells(5, 4).text
CANAL = ThisWorkbook.Worksheets("CADASTRO").Cells(6, 4).text
ESCRITÓRIO = ThisWorkbook.Worksheets("CADASTRO").Cells(7, 4).text
EQUIPE = 0 & ThisWorkbook.Worksheets("CADASTRO").Cells(8, 4).text
centro = ThisWorkbook.Worksheets("DADOS CADASTRO").Cells(2, 12).text
CTGHIE = "A"
CLIHIE = ThisWorkbook.Worksheets("DADOS CADASTRO").Cells(2, 8).text

' # Checando se campos estão preenchidos antes de ampliar. # '
If EMPRESA = "" Or ORGANIZAÇÃO = "" Or CANAL = "" Or CLIHIE = "" Or EMISSOR = "" Then
    MsgBox ("Preencha todos os campos obrigatórios para ampliação.")
    Set session = Nothing
    Exit Sub
End If

' # Checando se cliente está ampliado # '
session.FindById("wnd[0]").maximize
session.FindById("wnd[0]/tbar[0]/okcd").text = "/NXD02"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[1]/usr/ctxtRF02D-KUNNR").text = EMISSOR
session.FindById("wnd[1]/usr/btnBUTTON2").press
session.FindById("wnd[2]/usr/ctxt*TVKO-VKORG").text = ORGANIZAÇÃO
session.FindById("wnd[2]/tbar[0]/btn[0]").press
If session.FindById("wnd[2]/usr/tblSAPMF02DTCTRL_KUNDENVERTRIEB/ctxtRF02D-VKOKU[0,0]").text = ORGANIZAÇÃO And session.FindById("wnd[2]/usr/tblSAPMF02DTCTRL_KUNDENVERTRIEB/ctxtRF02D-VTWKU[2,0]").text = CANAL Then
    session.FindById("wnd[2]/usr/tblSAPMF02DTCTRL_KUNDENVERTRIEB/ctxtRF02D-VKOKU[0,0]").SetFocus
    session.FindById("wnd[2]/tbar[0]/btn[0]").press
    MsgBox ("Cliente já ampliado!")
    Set session = Nothing
    Exit Sub
End If
If session.FindById("wnd[2]/usr/tblSAPMF02DTCTRL_KUNDENVERTRIEB/ctxtRF02D-VKOKU[0,1]").text = ORGANIZAÇÃO And session.FindById("wnd[2]/usr/tblSAPMF02DTCTRL_KUNDENVERTRIEB/ctxtRF02D-VTWKU[2,1]").text = CANAL Then
    session.FindById("wnd[2]/usr/tblSAPMF02DTCTRL_KUNDENVERTRIEB/ctxtRF02D-VKOKU[0,0]").SetFocus
    session.FindById("wnd[2]/tbar[0]/btn[0]").press
    MsgBox ("Cliente já ampliado!")
    Set session = Nothing
    Exit Sub
End If
If session.FindById("wnd[2]/usr/tblSAPMF02DTCTRL_KUNDENVERTRIEB/ctxtRF02D-VKOKU[0,2]").text = ORGANIZAÇÃO And session.FindById("wnd[2]/usr/tblSAPMF02DTCTRL_KUNDENVERTRIEB/ctxtRF02D-VTWKU[2,2]").text = CANAL Then
    session.FindById("wnd[2]/usr/tblSAPMF02DTCTRL_KUNDENVERTRIEB/ctxtRF02D-VKOKU[0,0]").SetFocus
    session.FindById("wnd[2]/tbar[0]/btn[0]").press
    MsgBox ("Cliente já ampliado!")
    Set session = Nothing
    Exit Sub
Else
    session.FindById("wnd[2]/usr/tblSAPMF02DTCTRL_KUNDENVERTRIEB/ctxtRF02D-VKOKU[0,0]").SetFocus
    session.FindById("wnd[2]/tbar[0]/btn[0]").press
    GoTo Ampliar
End If

' # Script SAP # '
Ampliar:
session.FindById("wnd[0]").maximize
session.FindById("wnd[0]/tbar[0]/okcd").text = "/NXD01"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[1]/usr/cmbRF02D-KTOKD").Key = ""
session.FindById("wnd[1]/usr/ctxtRF02D-REF_KUNNR").text = EMISSOR
session.FindById("wnd[1]/usr/ctxtRF02D-REF_KUNNR").SetFocus
session.FindById("wnd[1]/usr/btnBUTTON2").press
session.FindById("wnd[2]/usr/tblSAPMF02DTCTRL_KUNDENVERTRIEB/ctxtRF02D-VKOKU[0,0]").SetFocus
session.FindById("wnd[2]").SendVKey 2
session.FindById("wnd[1]/usr/ctxtRF02D-KUNNR").text = EMISSOR
session.FindById("wnd[1]/usr/ctxtRF02D-BUKRS").text = EMPRESA
session.FindById("wnd[1]/usr/ctxtRF02D-VKORG").text = ORGANIZAÇÃO
session.FindById("wnd[1]/usr/ctxtRF02D-VTWEG").text = CANAL
session.FindById("wnd[1]/usr/ctxtRF02D-SPART").text = "00"
session.FindById("wnd[1]").SendVKey 0
On Error Resume Next
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0202/subAREA1:SAPMF02D:7211/ctxtKNB1-AKONT").text = "1102001001"
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0202/subAREA1:SAPMF02D:7211/ctxtKNB1-ZUAWA").text = "001"
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0202/subAREA1:SAPMF02D:7211/ctxtKNB1-FDGRV").text = "E1"
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0202/subAREA2:SAPMF02D:7212/ctxtKNB1-VZSKZ").text = "01"
session.FindById("wnd[0]/tbar[0]/btn[11]").press
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0202/subAREA1:SAPMF02D:7215/chkKNB1-XZVER").Selected = True
session.FindById("wnd[0]/tbar[0]/btn[11]").press
If ORGANIZAÇÃO = "1900" Or ORGANIZAÇÃO = "3400" Then
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPMF02D:7310/ctxtKNVV-VKBUR").text = "2000"
    session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPMF02D:7310/ctxtKNVV-VKGRP").text = "043"
    session.FindById("wnd[0]/tbar[0]/btn[11]").press
End If
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPMF02D:7313/ctxtKNVH-HITYP").text = CTGHIE
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA4:SAPMF02D:7313/ctxtKNVH-HITYP").text = CTGHIE
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA3:SAPMF02D:7313/ctxtLINK_KNVH-HKUNNR").text = CLIHIE
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA4:SAPMF02D:7313/ctxtLINK_KNVH-HKUNNR").text = CLIHIE
session.FindById("wnd[0]/tbar[0]/btn[11]").press
session.FindById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPMF02D:7315/ctxtKNVV-VWERK").text = centro
session.FindById("wnd[0]/tbar[0]/btn[11]").press

Set session = Nothing
MsgBox ("Ampliado com sucesso.")

End Sub