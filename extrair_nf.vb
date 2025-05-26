Sub extrair_nf()

' # Variáveis # '
Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)
Dim hwndParent As LongPtr
Dim hwndChild As LongPtr
Const BM_CLICK As Long = &HF5
Dim hwndSaveDialog As LongPtr
Dim hwndPathField As LongPtr
Dim hwndCuteChild As LongPtr
Const WM_SETTEXT As Long = &HC
Dim fileName As String
Dim diretório As String

' # Script SAP # '
ultimaLinha = lastRow(3)
For linha = 8 To ultimaLinha
    periodoInicial = ThisWorkbook.Worksheets("EXTRAIR NF").Cells(3, "D").text
    NF = ThisWorkbook.Worksheets("EXTRAIR NF").Cells(linha, 3)
    série = ThisWorkbook.Worksheets("EXTRAIR NF").Cells(linha, 4)
    session.FindById("wnd[0]").maximize
    session.FindById("wnd[0]/tbar[0]/okcd").text = "/NJ1B3N"
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[0]").SendVKey 4
    session.FindById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").text = série
    session.FindById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[1,24]").text = NF
    session.FindById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[4,24]").text = "WE"
    session.FindById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/btnG_SELFLD_TAB-MORE[5,56]").press
    session.FindById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "1000"
    session.FindById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "3500"
    session.FindById("wnd[2]/tbar[0]/btn[8]").press
    session.FindById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[8,24]").SetFocus
    session.FindById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/btnG_SELFLD_TAB-MORE[8,56]").press
    session.FindById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/btnRSCSEL_255-SOP_I[0,0]").SetFocus
    session.FindById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/btnRSCSEL_255-SOP_I[0,0]").press
    session.FindById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").SetCurrentCell 1, "TEXT"
    session.FindById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").SelectedRows = "1"
    session.FindById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").DoubleClickCurrentCell
    session.FindById("wnd[2]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = periodoInicial
    session.FindById("wnd[2]/tbar[0]/btn[8]").press
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
    session.FindById("wnd[1]").SendVKey 2
    ThisWorkbook.Worksheets("EXTRAIR NF").Cells(linha, 5) = session.FindById("wnd[0]/usr/ctxtJ_1BDYDOC-DOCNUM").text
    session.FindById("wnd[0]/mbar/menu[0]/menu[8]").Select
    session.FindById("wnd[1]/usr/btnSPOP-OPTION1").press
    session.FindById("wnd[1]/tbar[0]/btn[0]").press

    ' # Colocando pausa até tela de print, abrir # '
    Application.Wait (Now + TimeValue("0:00:07"))

    ' # Localizando janela de print e dando OK # '
    arrombado = FindWindow(vbNullString, "Imprimir")
    If arrombado <> 0 Then
        arrombadinho = FindWindowEx(arrombado, 0, vbNullString, "OK")
        If arrombadinho <> 0 Then
            SendMessage arrombadinho, BM_CLICK, 0, 0
        Else
            MsgBox ("Botão não encontrado.")
        End If
    Else
        MsgBox ("Janela não encontrada.")
    End If

    ' # Colocando pausa até janela de salvar, abrir # '
    Application.Wait (Now + TimeValue("0:00:05"))

    ' # Localizando janela do CutePDF, definindo path e salvando # '
    arrombador = FindWindow(vbNullString, "Salvar Saída de Impressão como")
    If arrombador <> 0 Then
        arrombadator = FindWindowEx(arrombador, 0, vbNullString, Edit)
        If arrombadator <> 0 Then
            SetForegroundWindow arrombador
            Matricula = Environ("USERPROFILE")
            If Matricula = "60064" Then
                diretório = "C:\Users\60064\OneDrive - Positivo\Área de Trabalho\Arquivos\"
            Else
                diretório = Environ("USERPROFILE") & "\Desktop\Arquivos\"
            End If
            fileName = "NF " & ThisWorkbook.Worksheets("EXTRAIR NF").Cells(linha, 3)
            SendKeys diretório & fileName & ".pdf", True
            Application.Wait (Now + TimeValue("0:00:02"))
        Else
            MsgBox ("Campo de diretório não encontrado.")
        End If
    Else
        MsgBox ("Janela de salvar não encontrada.")
    End If
    arrombadaço = FindWindowEx(arrombador, 0, vbNullString, "Sa&lvar")
    If arrombadaço <> 0 Then
        SendMessage arrombadaço, BM_CLICK, 0, 0
    Else
        MsgBox ("Botão de salvar não encontrado.")
    End If

Next

Set session = Nothing
MsgBox ("NF's salvas com sucesso.")

End Sub