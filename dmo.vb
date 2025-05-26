Sub ORDEM_COTAÇÃO_DMO()

' # Variáveis # '
Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)
Dim NOMEDMO, EMAIL1, EMAIL2, FORMULÁRIO, COMPROVANTE, TRANSAÇÃOMODIFICAR As String

' # Atribuição de Valores # '
NOMEDMO = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(6, "X").Value
ANEXO1 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(7, "X").Value
ANEXO2 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(8, "X").Value
ANEXO3 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(9, "X").Value
ANEXO4 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(10, "X").Value
TRANSAÇÃOMODIFICAR = ThisWorkbook.Worksheets("DADOS").Cells(18, 3).Value

' # Script SAP # '
session.FindById("wnd[0]").maximize
session.FindById("wnd[0]/tbar[0]/okcd").text = "/NCV01N"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/ctxtDRAW-DOKNR").text = ""
session.FindById("wnd[0]/usr/ctxtDRAW-DOKAR").text = "DMO"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/txtDRAT-DKTXT").text = NOMEDMO
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/ctxtDRAW-LABOR").text = "D01"

' # Sistema para verificar se arquivo é pdf ou xlsx # '
On Error Resume Next
If Not ANEXO1 = "" Then
    session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/btnPB_FILE_BROWSER").press
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = Environ("USERPROFILE") & "\Desktop\Arquivos\" & ANEXO1 & ".pdf"
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = Environ("USERPROFILE") & "\Desktop\Arquivos\" & ANEXO1 & ".jpg"
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = Environ("USERPROFILE") & "\Desktop\Arquivos\" & ANEXO1 & ".jpeg"
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = Environ("USERPROFILE") & "\Desktop\Arquivos\" & ANEXO1 & ".xlsx"
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
    session.FindById("wnd[0]").SendVKey 0
        End If
If Not ANEXO2 = "" Then
    session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/btnPB_FILE_BROWSER").press
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = Environ("USERPROFILE") & "\Desktop\Arquivos\" & ANEXO2 & ".pdf"
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = Environ("USERPROFILE") & "\Desktop\Arquivos\" & ANEXO2 & ".jpg"
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
        session.FindById("wnd[1]/usr/ctxtDY_PATH").text = Environ("USERPROFILE") & "\Desktop\Arquivos\" & ANEXO2 & ".jpeg"
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = Environ("USERPROFILE") & "\Desktop\Arquivos\" & ANEXO2 & ".xlsx"
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
    session.FindById("wnd[0]").SendVKey 0
        End If
If Not ANEXO3 = "" Then
    session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/btnPB_FILE_BROWSER").press
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = Environ("USERPROFILE") & "\Desktop\Arquivos\" & ANEXO3 & ".pdf"
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = Environ("USERPROFILE") & "\Desktop\Arquivos\" & ANEXO3 & ".jpg"
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
        session.FindById("wnd[1]/usr/ctxtDY_PATH").text = Environ("USERPROFILE") & "\Desktop\Arquivos\" & ANEXO3 & ".jpeg"
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = Environ("USERPROFILE") & "\Desktop\Arquivos\" & ANEXO3 & ".xlsx"
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
    session.FindById("wnd[0]").SendVKey 0
        End If
If Not ANEXO4 = "" Then
    session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/btnPB_FILE_BROWSER").press
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = Environ("USERPROFILE") & "\Desktop\Arquivos\" & ANEXO4 & ".pdf"
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = Environ("USERPROFILE") & "\Desktop\Arquivos\" & ANEXO4 & ".jpg"
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = Environ("USERPROFILE") & "\Desktop\Arquivos\" & ANEXO4 & ".jpg"
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = Environ("USERPROFILE") & "\Desktop\Arquivos\" & ANEXO4 & ".xlsx"
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
    session.FindById("wnd[0]").SendVKey 0
        End If
On Error GoTo 0

' # Selecionando arquivos para taro doc e salvando # '
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").SelectNode "          1"
If Not ANEXO2 = "" Then
    session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").SelectNode "          2"
        End If
If Not ANEXO3 = "" Then
    session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").SelectNode "          3"
        End If
If Not ANEXO4 = "" Then
    session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").SelectNode "          4"
        End If
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/btnPB_CHECKIN").press
session.FindById("wnd[1]/usr/tblSAPLCV117TBL_SCAT").GetAbsoluteRow(1).Selected = True
session.FindById("wnd[1]/tbar[0]/btn[0]").press
session.FindById("wnd[0]/tbar[0]/btn[11]").press

' # Salvando na planilha o doc gerado # '
ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(11, "X").Value = session.FindById("wnd[0]/usr/ctxtDRAW-DOKNR").text

' # Anexando na ordem/cotação # '
session.FindById("wnd[0]/tbar[0]/okcd").text = TRANSAÇÃOMODIFICAR
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/ctxtVBAK-VBELN").text = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(3, "W").Value
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").Select
session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]").SetFocus
session.FindById("wnd[0]/mbar/menu[3]/menu[10]").Select
If Not session.FindById("wnd[1]/usr/tblSAPLCVOBTCTRL_DOKUMENTE/ctxtDRAW-DOKAR[0,0]").text = "DMO" Then
    session.FindById("wnd[1]/usr/tblSAPLCVOBTCTRL_DOKUMENTE/ctxtDRAW-DOKAR[0,0]").text = "DMO"
    session.FindById("wnd[1]/usr/tblSAPLCVOBTCTRL_DOKUMENTE/ctxtDRAW-DOKNR[1,0]").text = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(11, "X").Value
    GoTo IgnorarRestante:
        End If
If Not session.FindById("wnd[1]/usr/tblSAPLCVOBTCTRL_DOKUMENTE/ctxtDRAW-DOKAR[0,1]").text = "DMO" Then
    session.FindById("wnd[1]/usr/tblSAPLCVOBTCTRL_DOKUMENTE/ctxtDRAW-DOKAR[0,1]").text = "DMO"
    session.FindById("wnd[1]/usr/tblSAPLCVOBTCTRL_DOKUMENTE/ctxtDRAW-DOKNR[1,1]").text = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(11, "X").Value
    GoTo IgnorarRestante:
        End If
If Not session.FindById("wnd[1]/usr/tblSAPLCVOBTCTRL_DOKUMENTE/ctxtDRAW-DOKAR[0,2]").text = "DMO" Then
    session.FindById("wnd[1]/usr/tblSAPLCVOBTCTRL_DOKUMENTE/ctxtDRAW-DOKAR[0,2]").text = "DMO"
    session.FindById("wnd[1]/usr/tblSAPLCVOBTCTRL_DOKUMENTE/ctxtDRAW-DOKNR[1,2]").text = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(11, "X").Value
    GoTo IgnorarRestante:
        End If
IgnorarRestante:
session.FindById("wnd[1]").SendVKey 0

' # Sistema para salvar ordem/cotação # '
On Error Resume Next
session.FindById("wnd[0]/tbar[0]/btn[11]").press
session.FindById("wnd[1]/tbar[0]/btn[0]").press
session.FindById("wnd[1]/usr/btnSPOP-VAROPTION1").press
On Error GoTo 0

Set session = Nothing
MsgBox ("DMO anexado com sucesso.")

End Sub