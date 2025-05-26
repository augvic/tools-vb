Sub gerar_tudo()

' ---------- DELETANDO ARQUIVOS ---------- '

Dim fullPath As String
Dim files As Variant
Dim file As Variant

' # Atribuindo valores # '
fullPath = "N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\"

' # Deletando arquivos na pasta # '
fileName = Dir(fullPath & "\*.*")
Do While fileName <> ""
    Kill fullPath & "\" & fileName
    fileName = Dir
Loop

' ------------------------------ '

' ---------- DECLARAÇÃO ---------- '

Set WordApplication = CreateObject("Word.Application")
tipo = ThisWorkbook.Worksheets("GERAR DOC").Cells(9, "C").Value

If tipo = "SIMPLES" Then
    WordApplication.Visible = True
    WordApplication.Documents.Open ("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\BASES\" & "SIMPLES.docm")
    WordApplication.Activate
End If

If tipo = "CUMULATIVO" Then
    WordApplication.Visible = True
    WordApplication.Documents.Open ("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\BASES\" & "CUMULATIVO.docm")
    WordApplication.Activate
End If

If tipo = "NÃO CUMULATIVO" Then
    WordApplication.Visible = True
    WordApplication.Documents.Open ("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\BASES\" & "NÃO CUMULATIVO.docm")
    WordApplication.Activate
End If

razão = ThisWorkbook.Worksheets("GERAR DOC").Cells(2, "C").Value
cnpj = ThisWorkbook.Worksheets("GERAR DOC").Cells(3, "C").Value
representante = ThisWorkbook.Worksheets("GERAR DOC").Cells(4, "C").Value
cpf = ThisWorkbook.Worksheets("GERAR DOC").Cells(5, "C").Value
ruaN = ThisWorkbook.Worksheets("GERAR DOC").Cells(6, "C").Value
cidade = ThisWorkbook.Worksheets("GERAR DOC").Cells(7, "C").Value
estado = ThisWorkbook.Worksheets("GERAR DOC").Cells(8, "C").Value
dataH = ThisWorkbook.Worksheets("GERAR DOC").Cells(10, "C").text

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#RAZÃO-SOCIAL#"
            .replacement.text = razão
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#RUA-N#"
            .replacement.text = ruaN
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#CIDADE#"
            .replacement.text = cidade
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#ESTADO#"
            .replacement.text = estado
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#CNPJ#"
            .replacement.text = cnpj
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#REPRESENTANTE#"
            .replacement.text = representante
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#CPF#"
            .replacement.text = cpf
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#DATA#"
            .replacement.text = dataH
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

If tipo = "SIMPLES" Then
    WordApplication.ActiveDocument.SaveAs fileName:=("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\DEC-SIMPLES.docm")
    WordApplication.ActiveDocument.ExportAsFixedFormat OutputFilename:="N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\DEC-SIMPLES.pdf", ExportFormat:=wdExportFormatPDF
    WordApplication.ActiveDocument.Close
    WordApplication.Visible = False
End If

If tipo = "CUMULATIVO" Then
    WordApplication.ActiveDocument.SaveAs fileName:=("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\DEC-CUMULATIVO.docm")
    WordApplication.ActiveDocument.ExportAsFixedFormat OutputFilename:="N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\DEC-CUMULATIVO.pdf", ExportFormat:=wdExportFormatPDF
    WordApplication.ActiveDocument.Close
    WordApplication.Visible = False
End If

If tipo = "NÃO CUMULATIVO" Then
    WordApplication.ActiveDocument.SaveAs fileName:=("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\DEC-NÃO CUMULATIVO.docm")
    WordApplication.ActiveDocument.ExportAsFixedFormat OutputFilename:="N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\DEC-NÃO CUMULATIVO.pdf", ExportFormat:=wdExportFormatPDF
    WordApplication.ActiveDocument.Close
    WordApplication.Visible = False
End If

' ---------- CONTRATO ---------- '

WordApplication.Visible = True
WordApplication.Documents.Open ("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\BASES\" & "CONTRATO.docm")
WordApplication.Activate

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#RAZÃO-SOCIAL#"
            .replacement.text = razão
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#RUA-N#"
            .replacement.text = ruaN
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#CIDADE#"
            .replacement.text = cidade
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#ESTADO#"
            .replacement.text = estado
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#CNPJ#"
            .replacement.text = cnpj
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#REPRESENTANTE#"
            .replacement.text = representante
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#CPF#"
            .replacement.text = cpf
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#DATA#"
            .replacement.text = dataH
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

WordApplication.ActiveDocument.SaveAs fileName:=("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\CONTRATO.docm")
WordApplication.ActiveDocument.ExportAsFixedFormat OutputFilename:="N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\CONTRATO.pdf", ExportFormat:=wdExportFormatPDF
WordApplication.ActiveDocument.Close
WordApplication.Visible = False

' ---------- ASSINATURA ---------- '

If tipo = "SIMPLES" Then
    WordApplication.Visible = True
    WordApplication.Documents.Open ("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\BASES\" & "ASSINATURA - SIMPLES.docm")
    WordApplication.Activate
End If

If tipo = "CUMULATIVO" Then
    WordApplication.Visible = True
    WordApplication.Documents.Open ("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\BASES\" & "ASSINATURA - CUMULATIVO.docm")
    WordApplication.Activate
End If

If tipo = "NÃO CUMULATIVO" Then
    WordApplication.Visible = True
    WordApplication.Documents.Open ("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\BASES\" & "ASSINATURA - NÃO CUMULATIVO.docm")
    WordApplication.Activate
End If

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#RAZÃO-SOCIAL#"
            .replacement.text = razão
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#RUA-N#"
            .replacement.text = ruaN
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#CIDADE#"
            .replacement.text = cidade
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#ESTADO#"
            .replacement.text = estado
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#CNPJ#"
            .replacement.text = cnpj
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#REPRESENTANTE#"
            .replacement.text = representante
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#CPF#"
            .replacement.text = cpf
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#DATA#"
            .replacement.text = dataH
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

WordApplication.ActiveDocument.SaveAs fileName:=("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\" & ThisWorkbook.Worksheets("GERAR DOC").Cells(7, "D") & ".docm")
WordApplication.ActiveDocument.ExportAsFixedFormat OutputFilename:="N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\" & ThisWorkbook.Worksheets("GERAR DOC").Cells(7, "D") & ".pdf", ExportFormat:=wdExportFormatPDF
WordApplication.ActiveDocument.Close
WordApplication.Visible = False

End Sub
Sub gerar_dec()

' ---------- DELETANDO ARQUIVOS ---------- '

Dim fullPath As String
Dim files As Variant
Dim file As Variant

' # Atribuindo valores # '
fullPath = "N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\"

' # Deletando arquivos na pasta # '
fileName = Dir(fullPath & "\*.*")
Do While fileName <> ""
    Kill fullPath & "\" & fileName
    fileName = Dir
Loop

' ------------------------------ '

Set WordApplication = CreateObject("Word.Application")
tipo = ThisWorkbook.Worksheets("GERAR DOC").Cells(9, "C").Value

If tipo = "SIMPLES" Then
    WordApplication.Visible = True
    WordApplication.Documents.Open ("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\BASES\" & "SIMPLES.docm")
    WordApplication.Activate
End If

If tipo = "CUMULATIVO" Then
    WordApplication.Visible = True
    WordApplication.Documents.Open ("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\BASES\" & "CUMULATIVO.docm")
    WordApplication.Activate
End If

If tipo = "NÃO CUMULATIVO" Then
    WordApplication.Visible = True
    WordApplication.Documents.Open ("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\BASES\" & "NÃO CUMULATIVO.docm")
    WordApplication.Activate
End If

razão = ThisWorkbook.Worksheets("GERAR DOC").Cells(2, "C").Value
cnpj = ThisWorkbook.Worksheets("GERAR DOC").Cells(3, "C").Value
representante = ThisWorkbook.Worksheets("GERAR DOC").Cells(4, "C").Value
cpf = ThisWorkbook.Worksheets("GERAR DOC").Cells(5, "C").Value
ruaN = ThisWorkbook.Worksheets("GERAR DOC").Cells(6, "C").Value
cidade = ThisWorkbook.Worksheets("GERAR DOC").Cells(7, "C").Value
estado = ThisWorkbook.Worksheets("GERAR DOC").Cells(8, "C").Value
dataH = ThisWorkbook.Worksheets("GERAR DOC").Cells(10, "C").text

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#RAZÃO-SOCIAL#"
            .replacement.text = razão
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#RUA-N#"
            .replacement.text = ruaN
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#CIDADE#"
            .replacement.text = cidade
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#ESTADO#"
            .replacement.text = estado
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#CNPJ#"
            .replacement.text = cnpj
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#REPRESENTANTE#"
            .replacement.text = representante
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#CPF#"
            .replacement.text = cpf
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#DATA#"
            .replacement.text = dataH
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

If tipo = "SIMPLES" Then
    WordApplication.ActiveDocument.SaveAs fileName:=("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\DEC-SIMPLES.docm")
    WordApplication.ActiveDocument.ExportAsFixedFormat OutputFilename:="N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\DEC-SIMPLES.pdf", ExportFormat:=wdExportFormatPDF
    WordApplication.ActiveDocument.Close
    WordApplication.Visible = False
End If

If tipo = "CUMULATIVO" Then
    WordApplication.ActiveDocument.SaveAs fileName:=("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\DEC-CUMULATIVO.docm")
    WordApplication.ActiveDocument.ExportAsFixedFormat OutputFilename:="N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\DEC-CUMULATIVO.pdf", ExportFormat:=wdExportFormatPDF
    WordApplication.ActiveDocument.Close
    WordApplication.Visible = False
End If

If tipo = "NÃO CUMULATIVO" Then
    WordApplication.ActiveDocument.SaveAs fileName:=("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\DEC-NÃO CUMULATIVO.docm")
    WordApplication.ActiveDocument.ExportAsFixedFormat OutputFilename:="N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\DEC-NÃO CUMULATIVO.pdf", ExportFormat:=wdExportFormatPDF
    WordApplication.ActiveDocument.Close
    WordApplication.Visible = False
End If

End Sub
Sub gerar_contrato()

' ---------- DELETANDO ARQUIVOS ---------- '

Dim fullPath As String
Dim files As Variant
Dim file As Variant

' # Atribuindo valores # '
fullPath = "N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\"

' # Deletando arquivos na pasta # '
fileName = Dir(fullPath & "\*.*")
Do While fileName <> ""
    Kill fullPath & "\" & fileName
    fileName = Dir
Loop

' ------------------------------ '

Set WordApplication = CreateObject("Word.Application")
tipo = ThisWorkbook.Worksheets("GERAR DOC").Cells(9, "C").Value

WordApplication.Visible = True
WordApplication.Documents.Open ("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\BASES\" & "CONTRATO.docm")
WordApplication.Activate

razão = ThisWorkbook.Worksheets("GERAR DOC").Cells(2, "C").Value
cnpj = ThisWorkbook.Worksheets("GERAR DOC").Cells(3, "C").Value
representante = ThisWorkbook.Worksheets("GERAR DOC").Cells(4, "C").Value
cpf = ThisWorkbook.Worksheets("GERAR DOC").Cells(5, "C").Value
ruaN = ThisWorkbook.Worksheets("GERAR DOC").Cells(6, "C").Value
cidade = ThisWorkbook.Worksheets("GERAR DOC").Cells(7, "C").Value
estado = ThisWorkbook.Worksheets("GERAR DOC").Cells(8, "C").Value
dataH = ThisWorkbook.Worksheets("GERAR DOC").Cells(10, "C").text

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#RAZÃO-SOCIAL#"
            .replacement.text = razão
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#RUA-N#"
            .replacement.text = ruaN
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#CIDADE#"
            .replacement.text = cidade
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#ESTADO#"
            .replacement.text = estado
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#CNPJ#"
            .replacement.text = cnpj
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#REPRESENTANTE#"
            .replacement.text = representante
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#CPF#"
            .replacement.text = cpf
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#DATA#"
            .replacement.text = dataH
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

WordApplication.ActiveDocument.SaveAs fileName:=("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\CONTRATO.docm")
WordApplication.ActiveDocument.ExportAsFixedFormat OutputFilename:="N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\CONTRATO.pdf", ExportFormat:=wdExportFormatPDF
WordApplication.ActiveDocument.Close
WordApplication.Visible = False

End Sub
Sub gerar_ass()

' ---------- DELETANDO ARQUIVOS ---------- '

Dim fullPath As String
Dim files As Variant
Dim file As Variant

' # Atribuindo valores # '
fullPath = "N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\"

' # Deletando arquivos na pasta # '
fileName = Dir(fullPath & "\*.*")
Do While fileName <> ""
    Kill fullPath & "\" & fileName
    fileName = Dir
Loop

' ------------------------------ '

Set WordApplication = CreateObject("Word.Application")
tipo = ThisWorkbook.Worksheets("GERAR DOC").Cells(9, "C").Value

If tipo = "SIMPLES" Then
    WordApplication.Visible = True
    WordApplication.Documents.Open ("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\BASES\" & "ASSINATURA - SIMPLES.docm")
    WordApplication.Activate
End If

If tipo = "CUMULATIVO" Then
    WordApplication.Visible = True
    WordApplication.Documents.Open ("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\BASES\" & "ASSINATURA - CUMULATIVO.docm")
    WordApplication.Activate
End If

If tipo = "NÃO CUMULATIVO" Then
    WordApplication.Visible = True
    WordApplication.Documents.Open ("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\BASES\" & "ASSINATURA - NÃO CUMULATIVO.docm")
    WordApplication.Activate
End If

razão = ThisWorkbook.Worksheets("GERAR DOC").Cells(2, "C").Value
cnpj = ThisWorkbook.Worksheets("GERAR DOC").Cells(3, "C").Value
representante = ThisWorkbook.Worksheets("GERAR DOC").Cells(4, "C").Value
cpf = ThisWorkbook.Worksheets("GERAR DOC").Cells(5, "C").Value
ruaN = ThisWorkbook.Worksheets("GERAR DOC").Cells(6, "C").Value
cidade = ThisWorkbook.Worksheets("GERAR DOC").Cells(7, "C").Value
estado = ThisWorkbook.Worksheets("GERAR DOC").Cells(8, "C").Value
dataH = ThisWorkbook.Worksheets("GERAR DOC").Cells(10, "C").text

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#RAZÃO-SOCIAL#"
            .replacement.text = razão
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#RUA-N#"
            .replacement.text = ruaN
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#CIDADE#"
            .replacement.text = cidade
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#ESTADO#"
            .replacement.text = estado
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#CNPJ#"
            .replacement.text = cnpj
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#REPRESENTANTE#"
            .replacement.text = representante
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#CPF#"
            .replacement.text = cpf
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

With WordApplication.ActiveDocument
    With .Range
        With .Find
            .text = "#DATA#"
            .replacement.text = dataH
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End With
End With

WordApplication.ActiveDocument.SaveAs fileName:=("N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\" & ThisWorkbook.Worksheets("GERAR DOC").Cells(7, "D") & ".docm")
WordApplication.ActiveDocument.ExportAsFixedFormat OutputFilename:="N:\NNPI\NN\Projeto Vendas 1 a 1\Ferramentas\FerramentasBack\GERAR DOCUMENTOS PLATAFORMA\FINAL\" & ThisWorkbook.Worksheets("GERAR DOC").Cells(7, "D") & ".pdf", ExportFormat:=wdExportFormatPDF
WordApplication.ActiveDocument.Close
WordApplication.Visible = False

End Sub