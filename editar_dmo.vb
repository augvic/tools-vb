Sub editar_DMO()

Dim ordemCotacao As String
Dim dmo As String

ordemCotacao = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(3, "W").Value
ANEXO1 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(7, "X").Value
ANEXO2 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(8, "X").Value
ANEXO3 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(9, "X").Value
ANEXO4 = ThisWorkbook.Worksheets("ORDEM | COTAÇÃO").Cells(10, "X").Value

Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)
Dim EMISSOR, ORGANIZAÇÃO, CANAL, ESCRITÓRIO, EQUIPE, centro As String

session.FindById("wnd[0]/tbar[0]/okcd").text = "/nva23"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/ctxtVBAK-VBELN").text = ordemCotacao
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/mbar/menu[3]/menu[10]").Select
dmo = session.FindById("wnd[1]/usr/tblSAPLCVOBTCTRL_DOKUMENTE/ctxtDRAW-DOKNR[1,0]").text
session.FindById("wnd[1]").Close
session.FindById("wnd[0]/tbar[0]/okcd").text = "/ncv02n"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/ctxtDRAW-DOKNR").text = dmo
session.FindById("wnd[0]/usr/ctxtDRAW-DOKAR").text = "DMO"
session.FindById("wnd[0]/usr/ctxtDRAW-DOKTL").text = "000"
session.FindById("wnd[0]/usr/ctxtDRAW-DOKVR").text = "00"
session.FindById("wnd[0]").SendVKey 0

On Error Resume Next

If Not ANEXO1 = "" Then
    session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/btnPB_FILE_BROWSER").press
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = Environ("USERPROFILE") & "\Desktop\Arquivos\" & ANEXO1 & ".pdf"
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = Environ("USERPROFILE") & "\Desktop\Arquivos\" & ANEXO1 & ".jpg"
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
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = Environ("USERPROFILE") & "\Desktop\Arquivos\" & "ANEXO2" & ".xlsx"
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
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = Environ("USERPROFILE") & "\Desktop\Arquivos\" & ANEXO4 & ".xlsx"
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
    session.FindById("wnd[0]").SendVKey 0
        End If

' # Selecionando arquivos para taro doc e salvando # '
On Error Resume Next
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").SelectNode "          1"
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/btnPB_CHECKIN").press
session.FindById("wnd[1]/usr/tblSAPLCV117TBL_SCAT").GetAbsoluteRow(1).Selected = True
session.FindById("wnd[1]/tbar[0]/btn[0]").press
session.FindById("wnd[1]/tbar[0]/btn[0]").press
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").UnselectNode "          1"
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").SelectNode "          2"
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/btnPB_CHECKIN").press
session.FindById("wnd[1]/usr/tblSAPLCV117TBL_SCAT").GetAbsoluteRow(1).Selected = True
session.FindById("wnd[1]/tbar[0]/btn[0]").press
session.FindById("wnd[1]/tbar[0]/btn[0]").press
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").UnselectNode "          2"
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").SelectNode "          3"
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/btnPB_CHECKIN").press
session.FindById("wnd[1]/usr/tblSAPLCV117TBL_SCAT").GetAbsoluteRow(1).Selected = True
session.FindById("wnd[1]/tbar[0]/btn[0]").press
session.FindById("wnd[1]/tbar[0]/btn[0]").press
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").UnselectNode "          3"
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").SelectNode "          4"
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/btnPB_CHECKIN").press
session.FindById("wnd[1]/usr/tblSAPLCV117TBL_SCAT").GetAbsoluteRow(1).Selected = True
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").UnselectNode "          4"
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").SelectNode "          5"
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/btnPB_CHECKIN").press
session.FindById("wnd[1]/usr/tblSAPLCV117TBL_SCAT").GetAbsoluteRow(1).Selected = True
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").UnselectNode "          5"
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").SelectNode "          6"
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/btnPB_CHECKIN").press
session.FindById("wnd[1]/usr/tblSAPLCV117TBL_SCAT").GetAbsoluteRow(1).Selected = True
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").UnselectNode "          6"
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").SelectNode "          7"
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/btnPB_CHECKIN").press
session.FindById("wnd[1]/usr/tblSAPLCV117TBL_SCAT").GetAbsoluteRow(1).Selected = True
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").UnselectNode "          7"
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").SelectNode "          8"
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/btnPB_CHECKIN").press
session.FindById("wnd[1]/usr/tblSAPLCV117TBL_SCAT").GetAbsoluteRow(1).Selected = True
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").UnselectNode "          8"
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").SelectNode "          9"
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/btnPB_CHECKIN").press
session.FindById("wnd[1]/usr/tblSAPLCV117TBL_SCAT").GetAbsoluteRow(1).Selected = True
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").UnselectNode "          9"
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").SelectNode "          10"
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/btnPB_CHECKIN").press
session.FindById("wnd[1]/usr/tblSAPLCV117TBL_SCAT").GetAbsoluteRow(1).Selected = True
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").UnselectNode "          10"
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/cntlCTL_FILES1/shellcont/shell/shellcont[1]/shell").SelectNode "          11"
session.FindById("wnd[0]/usr/tabsTAB_MAIN/tabpTSMAIN/ssubSCR_MAIN:SAPLCV110:0102/btnPB_CHECKIN").press
session.FindById("wnd[1]/usr/tblSAPLCV117TBL_SCAT").GetAbsoluteRow(1).Selected = True
session.FindById("wnd[1]/tbar[0]/btn[0]").press
session.FindById("wnd[1]/tbar[0]/btn[0]").press
session.FindById("wnd[0]/tbar[0]/btn[11]").press

End Sub
