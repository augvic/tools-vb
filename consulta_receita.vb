Private Sub consultarBTN_Click()

Unload Me

' # Variáveis # '
Dim requisicao As New MSXML2.XMLHTTP60
Dim resposta As Object
Dim url, cnpjconsultado, urlsuframa, payloadconverted As String
Dim payload As New Dictionary

' ---------- # Consultando Receita # ---------- '

' # Definindo API e trasferindo CNPJ para célula # '
cnpjconsultado = cnpjconsultadobox.text
ThisWorkbook.Worksheets("CADASTRO").Cells(3, "J") = cnpjconsultado
url = "https://publica.cnpj.ws/cnpj/" & cnpjconsultado

' # Realizando requisição Receita # '
requisicao.Open "Get", url, False
requisicao.Send

' # Se ocorrer erro na requisição # '
If requisicao.Status <> 200 Then
    MsgBox "Erro ao consultar: " & requisicao.responseText
    Exit Sub
End If

' # Convertendo para JSON a resposta da requisição # '
Set resposta = JsonConverter.ParseJson(requisicao.responseText)

' # Integrando dados da Receita na planilha # '
ThisWorkbook.Worksheets("CADASTRO").Cells(4, "G") = resposta("razao_social")
ThisWorkbook.Worksheets("CADASTRO").Cells(6, "G") = resposta("estabelecimento")("nome_fantasia")
ThisWorkbook.Worksheets("CADASTRO").Cells(10, "G") = resposta("estabelecimento")("tipo_logradouro") & " " & resposta("estabelecimento")("logradouro")
ThisWorkbook.Worksheets("CADASTRO").Cells(10, "J") = resposta("estabelecimento")("numero")
ThisWorkbook.Worksheets("CADASTRO").Cells(10, "L") = resposta("estabelecimento")("complemento")
ThisWorkbook.Worksheets("CADASTRO").Cells(13, "G") = resposta("estabelecimento")("bairro")
ThisWorkbook.Worksheets("CADASTRO").Cells(14, "G") = resposta("estabelecimento")("cep")
ThisWorkbook.Worksheets("CADASTRO").Cells(13, "K") = resposta("estabelecimento")("cidade")("nome")
ThisWorkbook.Worksheets("CADASTRO").Cells(13, "I") = resposta("estabelecimento")("estado")("sigla")
ThisWorkbook.Worksheets("CADASTRO").Cells(20, "F") = resposta("natureza_juridica")("id") & " - " & resposta("natureza_juridica")("descricao")
On Error Resume Next
ThisWorkbook.Worksheets("CADASTRO").Cells(3, 15) = resposta("estabelecimento")("inscricoes_estaduais")(1)("inscricao_estadual")
ThisWorkbook.Worksheets("DADOS CADASTRO").Cells(1, 19) = resposta("estabelecimento")("inscricoes_estaduais")(1)("ativo")
ThisWorkbook.Worksheets("CADASTRO").Cells(6, 15) = resposta("estabelecimento")("inscricoes_estaduais")(2)("inscricao_estadual")
ThisWorkbook.Worksheets("DADOS CADASTRO").Cells(2, 19) = resposta("estabelecimento")("inscricoes_estaduais")(2)("ativo")
ThisWorkbook.Worksheets("CADASTRO").Cells(9, 15) = resposta("estabelecimento")("inscricoes_estaduais")(3)("inscricao_estadual")
ThisWorkbook.Worksheets("DADOS CADASTRO").Cells(3, 19) = resposta("estabelecimento")("inscricoes_estaduais")(3)("ativo")
On Error GoTo 0

' # Tirando acento. # '
ThisWorkbook.Worksheets("CADASTRO").Cells(13, "K") = ThisWorkbook.Worksheets("CADASTRO").Cells(1, "C")

' # UCase # '
Dim Rng As Range
For Each Rng In Range("B2:L19")
    Rng.Value = UCase(Rng.Value)
Next

' # Cortando caracteres da string Razão Social # '
razãoInteira = ThisWorkbook.Worksheets("CADASTRO").Cells(4, "G").text
razãoCurtinha = Left(razãoInteira, 15)
ThisWorkbook.Worksheets("CADASTRO").Cells(8, "G") = razãoCurtinha

' # Verificando IE # '
If ThisWorkbook.Worksheets("CADASTRO").Cells(4, "O") = "IE HABILITADA" Then
    ThisWorkbook.Worksheets("CADASTRO").Cells(12, "N") = ThisWorkbook.Worksheets("CADASTRO").Cells(3, "O")
Else
    If ThisWorkbook.Worksheets("CADASTRO").Cells(7, "O") = "IE HABILITADA" Then
        ThisWorkbook.Worksheets("CADASTRO").Cells(12, "N") = ThisWorkbook.Worksheets("CADASTRO").Cells(6, "O")
    Else
        If ThisWorkbook.Worksheets("CADASTRO").Cells(10, "O") = "IE HABILITADA" Then
            ThisWorkbook.Worksheets("CADASTRO").Cells(12, "N") = ThisWorkbook.Worksheets("CADASTRO").Cells(9, "O")
        Else
            ThisWorkbook.Worksheets("CADASTRO").Cells(12, "N") = "ISENTO"
        End If
    End If
End If

' # Verificando Suframa # '
Dim estados As Variant
Dim estado As String
estados = Array("AC", "AP", "AM", "RO", "RR")
estado = ThisWorkbook.Worksheets("CADASTRO").Cells(13, "I")

If IsInArray(estado, estados) = True Then
    ThisWorkbook.Worksheets("CADASTRO").Cells(16, "O") = "VERIFICAR SUFRAMA"
End If

End Sub