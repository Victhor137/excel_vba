# README — Explaining project files

Este README foi gerado automaticamente pelo assistente. Ele descreve os arquivos encontrados no ZIP enviado e inclui trechos dos códigos legíveis. Se seu projeto contiver arquivos VBA embutidos em `.xlsm`/`.xlsb`, esses aparecem como binários e não foram extraídos para texto aqui — veja nota abaixo.

## Resumo da estrutura do projeto

- `excel_vba-main/README.md`
- `excel_vba-main/alterar_datas_definir_uma_padrao.vba`
- `excel_vba-main/colorir_historicos_iguais.vba`
- `excel_vba-main/exportar_dados_txt_personalizado_mesma_pasta_do_arquivo.vba`
- `excel_vba-main/exportar_txt_padrao_dominiocontabil.vba`
- `excel_vba-main/extrair_extratobancario.vba`
- `excel_vba-main/moverdados.vba`
- `excel_vba-main/remover_linhas_vazias.vba`
- `excel_vba-main/removersinal_cred_deb.vba`

---

## Arquivos analisados (trechos)

### `excel_vba-main/README.md`

- Tipo: **Texto legível**.

- Trecho inicial (até 5000 caracteres):

```vb
# excel_vba
Códigos em VBA para automatização de planilhas.

```

- Tamanho: 63 bytes.

### `excel_vba-main/alterar_datas_definir_uma_padrao.vba`

- Tipo: **Texto** (possivelmente script/descrição).

- Trecho inicial (até 5000 caracteres):

```vb
'Defini todas as datas de uma coluna para um valor definido. Economizando tempo repetitivo.

Sub Planilha_Modelo_PreencherDatas()

    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim dia As String, mes As String, ano As String
    Dim dataFormatada As String

    ' Define a aba específica
    Set ws = ThisWorkbook.Sheets("PLANILHA_MODELO")

    ' Lê os valores de dia, mês e ano
    dia = Format(ws.Range("B4").Value, "00")
    mes = Format(ws.Range("C4").Value, "00")
    ano = ws.Range("D4").Value

    ' Encontra a última linha com valor na coluna C
    ultimaLinha = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    ' Preenche a coluna B a partir da linha 9
    For i = 9 To ultimaLinha
        If ws.Cells(i, "C").Value <> "" Then
            dataFormatada = dia & "/" & mes & "/" & ano
            ws.Cells(i, "B").Value = dataFormatada
        Else
            ws.Cells(i, "B").ClearContents
        End If
    Next i

End Sub

```

- Tamanho: 965 bytes.

### `excel_vba-main/colorir_historicos_iguais.vba`

- Tipo: **Texto** (possivelmente script/descrição).

- Trecho inicial (até 5000 caracteres):

```vb
'este modulo vai analisar tudo antes do segundo hifen (-) e vai colorir em tons pasteis, os historicos que forem iguais
  Sub Planilha_Modelo_Colorir()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim dictEnderecos As Object
    Dim textoHistorico As String
    Dim endereco As String
    Dim posUltimo As Long
    Dim corIndex As Long
    Dim cores()

    ' Cores
    cores = Array( _
        RGB(255, 204, 153), _
        RGB(204, 255, 204), _
        RGB(204, 229, 255), _
        RGB(255, 204, 229), _
        RGB(255, 255, 153), _
        RGB(204, 204, 255), _
        RGB(255, 230, 153), _
        RGB(153, 255, 204), _
        RGB(255, 179, 179), _
        RGB(204, 255, 255) _
    )

    Set ws = ActiveSheet
    ultimaLinha = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    Set dictEnderecos = CreateObject("Scripting.Dictionary")
    corIndex = 0

    For i = 9 To ultimaLinha
        textoHistorico = ws.Cells(i, "F").Value
        endereco = ""

        ' Analisa tudo antes do último hífen
        posUltimo = InStrRev(textoHistorico, "-")
        If posUltimo > 0 Then
            endereco = Trim(Left(textoHistorico, posUltimo - 1))
        Else
            endereco = Trim(textoHistorico)
        End If

        If endereco <> "" Then
            If Not dictEnderecos.Exists(endereco) Then
                dictEnderecos.Add endereco, cores(corIndex Mod (UBound(cores) + 1))
                corIndex = corIndex + 1
            End If
            ws.Range("B" & i & ":F" & i).Interior.Color = dictEnderecos(endereco)
        End If
    Next i

    MsgBox "Colorização com tons pastéis aplicada com sucesso!", vbInformation
End Sub

```

- Tamanho: 1679 bytes.

### `excel_vba-main/exportar_dados_txt_personalizado_mesma_pasta_do_arquivo.vba`

- Tipo: **Texto** (possivelmente script/descrição).

- Trecho inicial (até 5000 caracteres):

```vb
'Este codigo necessita de um menu criado na parte de formularios'
Sub Planilha_Modelo_Exportar_Personalizado()
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long, colDestino As Long
    Dim linhaDestino As Long
    ' Exibe o formulário de seleção
    frmSelecionarColunas.Tag = ""
    frmSelecionarColunas.Show

    ' Verifica se o usuário clicou em Cancelar
    If frmSelecionarColunas.Tag = "CANCELADO" Then
        Unload frmSelecionarColunas
        MsgBox "Operação cancelada pelo usuário.", vbExclamation
        Exit Sub
    End If

    ' Se o formulário ainda está visível, foi cancelado
    If frmSelecionarColunas.Visible Then
        Unload frmSelecionarColunas
        Exit Sub
    End If

    ' Define as planilhas
    Set wsOrigem = ThisWorkbook.Sheets("PLANILHA_MODELO")
    Set wsDestino = ThisWorkbook.Sheets("PERSONALIZADOS")

    Sheets("PERSONALIZADOS").Select
    Range("A1").Select
    Call Planilha_Personalizados_LimparDados
    Sheets("PLANILHA_MODELO").Select
    Range("A1").Select
    ' Encontra a última linha com dados na origem (coluna B)
    ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "B").End(xlUp).Row
    If ultimaLinha < 9 Then
        MsgBox "ERRO! Revise os dados antes de continuar.", vbExclamation
        Exit Sub
    End If

    ' Limpa dados anteriores na aba DOMINIO
    wsDestino.Range("A2:F" & wsDestino.Rows.Count).ClearContents

    ' Começa a colar na linha 2
    linhaDestino = 2

    ' Loop pelas linhas da origem
    Dim colCabecalho As Long
    colCabecalho = 1
    
    With wsDestino
        If frmSelecionarColunas.chkColA.Value Then
            .Cells(1, colCabecalho).Value = "EXTRATO"
            colCabecalho = colCabecalho + 1
        End If
    
        If frmSelecionarColunas.chkColB.Value Then
            .Cells(1, colCabecalho).Value = "DATA"
            colCabecalho = colCabecalho + 1
        End If
    
        If frmSelecionarColunas.chkColC.Value Then
            .Cells(1, colCabecalho).Value = "VALOR"
            colCabecalho = colCabecalho + 1
        End If
    
        If frmSelecionarColunas.chkColD.Value Then
            .Cells(1, colCabecalho).Value = "CONTA DÉBITO"
            colCabecalho = colCabecalho + 1
        End If
    
        If frmSelecionarColunas.chkColE.Value Then
            .Cells(1, colCabecalho).Value = "CONTA CRÉDITO"
            colCabecalho = colCabecalho + 1
        End If
    
        If frmSelecionarColunas.chkColF.Value Then
            .Cells(1, colCabecalho).Value = "HISTÓRICO"
            colCabecalho = colCabecalho + 1
        End If
    End With


    For i = 9 To ultimaLinha
        colDestino = 1

        ' Coluna A: EXTRATO BANCÁRIO
        If frmSelecionarColunas.chkColA.Value Then
            wsDestino.Cells(linhaDestino, colDestino).Value = wsOrigem.Cells(i, 1).Value
            wsDestino.Cells(linhaDestino, colDestino).NumberFormat = "@"
            colDestino = colDestino + 1
        End If

        ' Coluna B: DATA
        If frmSelecionarColunas.chkColB.Value Then
            wsDestino.Cells(linhaDestino, colDestino).Value = wsOrigem.Cells(i, 2).Value
            wsDestino.Cells(linhaDestino, colDestino).NumberFormat = "dd/mm/yyyy"
            colDestino = colDestino + 1
        End If

        ' Coluna C: VALOR
        If frmSelecionarColunas.chkColC.Value Then
            wsDestino.Cells(linhaDestino, colDestino).Value = wsOrigem.Cells(i, 3).Value
            wsDestino.Cells(linhaDestino, colDestino).NumberFormat = "0.00"
            colDestino = colDestino + 1
        End If

        ' Coluna D: CONTA DÉBITO
        If frmSelecionarColunas.chkColD.Value Then
            wsDestino.Cells(linhaDestino, colDestino).Value = wsOrigem.Cells(i, 4).Value
            wsDestino.Cells(linhaDestino, colDestino).NumberFormat = "@"
            colDestino = colDestino + 1
        End If

        ' Coluna E: CONTA CRÉDITO
        If frmSelecionarColunas.chkColE.Value Then
            wsDestino.Cells(linhaDestino, colDestino).Value = wsOrigem.Cells(i, 5).Value
            wsDestino.Cells(linhaDestino, colDestino).NumberFormat = "@"
            colDestino = colDestino + 1
        End If

        ' Coluna F: HISTÓRICO
        If frmSelecionarColunas.chkColF.Value Then
            wsDestino.Cells(linhaDestino, colDestino).Value = wsOrigem.Cells(i, 6).Value
            wsDestino.Cells(linhaDestino, colDestino).NumberFormat = "@"
            colDestino = colDestino + 1
        End If

        linhaDestino = linhaDestino + 1
    Next i

    Unload frmSelecionarColunas
    Sheets("PERSONALIZADOS").Select
    Range("A1").Select
    MsgBox "Planilha exportada para PERSONALIZADOS!", vbInformation
End Sub

```

- Tamanho: 4736 bytes.

### `excel_vba-main/exportar_txt_padrao_dominiocontabil.vba`

- Tipo: **Texto** (possivelmente script/descrição).

- Trecho inicial (até 5000 caracteres):

```vb
' Exportar uma planilha no formato dominio contabil para TXT na mesma pasta do arquivo excel.
Sub Planilha_Dominio_Exportar_TXT()
    Dim ws As Worksheet
    Dim caminhoArquivo As String
    Dim linhaTexto As String
    Dim arquivoNum As Integer
    Dim pasta As String
    Dim nomeArquivo As String
    Dim ultimaLinha As Long
    Dim i As Long
    Dim linhaAtual As Long
    Dim celulaTexto As String

    ' Tentar definir a aba "DOMINIO"
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("DOMINIO")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "A planilha 'DOMINIO' não foi encontrada.", vbExclamation
        Exit Sub
    End If

    ' Verificar a última linha com dados na coluna A
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Obter a pasta onde o arquivo Excel está salvo
    If ThisWorkbook.Path = "" Then
        MsgBox "Salve o arquivo Excel antes de exportar.", vbExclamation
        Exit Sub
    End If

    pasta = ThisWorkbook.Path
    nomeArquivo = "Dominio_Exportado_YHWH_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    caminhoArquivo = pasta & "\" & nomeArquivo

    ' Abrir o arquivo para escrita
    arquivoNum = FreeFile
    Open caminhoArquivo For Output As #arquivoNum

    ' Percorrer da linha 2 até a última linha (ignorando o cabeçalho)
    For linhaAtual = 2 To ultimaLinha
        linhaTexto = ""
        For i = 1 To 6 ' Colunas A (1) a F (6)
            celulaTexto = ws.Cells(linhaAtual, i).Text
            ' Remover quebras de linha da célula
            celulaTexto = Replace(celulaTexto, vbCrLf, " ")
            celulaTexto = Replace(celulaTexto, vbCr, " ")
            celulaTexto = Replace(celulaTexto, vbLf, " ")
            
            linhaTexto = linhaTexto & celulaTexto
            If i < 6 Then
                linhaTexto = linhaTexto & ";"
            End If
        Next i
        Print #arquivoNum, linhaTexto
    Next linhaAtual

    ' Fechar o arquivo
    Close #arquivoNum

   MsgBox "Dados exportados para TXT na pasta:" & vbCrLf & vbCrLf & caminhoArquivo, vbInformation
End Sub

```

- Tamanho: 2085 bytes.

### `excel_vba-main/extrair_extratobancario.vba`

- Tipo: **Texto** (possivelmente script/descrição).

- Trecho inicial (até 5000 caracteres):

```vb
'Este modulo vai buscar o nome da coluna e com base nisso modificar e extrair um extrato bancario para a formatação padrão da dominio contabilidade "DATA", "HIST", "VALOR"
Sub Planilha_Processo_TratarDados()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long, lastCol As Long
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    lastRow = 1000

    Dim i As Long
    Dim colData As Long, colHist1 As Long, colHist2 As Long, colHist3 As Long
    Dim colValor1 As Long, colValor2 As Long
    Dim colINSS As Long, colIRRF As Long, colDesconto As Long

    ' Inicializar colunas
    colData = 0: colHist1 = 0: colHist2 = 0: colHist3 = 0
    colValor1 = 0: colValor2 = 0
    colINSS = 0: colIRRF = 0: colDesconto = 0

    ' Identificar colunas pela linha 2
    For i = 1 To lastCol
        Select Case Trim(UCase(ws.Cells(2, i).Value))
            Case "DATA": colData = i
            Case "HST_1": colHist1 = i
            Case "HST_2": colHist2 = i
            Case "HST_3": colHist3 = i
            Case "VALOR_DEBITO": colValor1 = i
            Case "VALOR_CRÉDITO": colValor2 = i
            Case "OUTRO_VALOR_1": colINSS = i
            Case "OUTRO_VALOR_2": colIRRF = i
            Case "OUTRO_VALOR_3": colDesconto = i
        End Select
    Next i

    ' Checar colunas obrigatórias
    If colData = 0 Or colHist1 = 0 Then
        MsgBox "Colunas obrigatórias ausentes!", vbCritical
        Exit Sub
    End If

    Dim novaLinha As Long
    novaLinha = 3

    For i = 3 To lastRow
        Dim valorData As Variant
        valorData = ws.Cells(i, colData).Value

        If IsDate(valorData) Then
            ' Montar histórico base
            Dim historico As String
            historico = ws.Cells(i, colHist1).Text

            If colHist2 > 0 Then
                If Trim(ws.Cells(i, colHist2).Text) <> "" Then
                    historico = historico & " - " & ws.Cells(i, colHist2).Text
                End If
            End If

            If colHist3 > 0 Then
                If Trim(ws.Cells(i, colHist3).Text) <> "" Then
                    historico = historico & " - " & ws.Cells(i, colHist3).Text
                End If
            End If

            ' VALOR 1 (débito)
            If colValor1 > 0 Then
                If Trim(ws.Cells(i, colValor1).Text) <> "" And IsNumeric(ws.Cells(i, colValor1).Value) Then
                    ws.Cells(novaLinha, 1).Value = valorData
                    ws.Cells(novaLinha, 2).Value = historico
                    ws.Cells(novaLinha, 3).Value = ws.Cells(i, colValor1).Value
                    novaLinha = novaLinha + 1
                End If
            End If

            ' VALOR 2 (crédito negativo)
            If colValor2 > 0 Then
                If Trim(ws.Cells(i, colValor2).Text) <> "" And IsNumeric(ws.Cells(i, colValor2).Value) Then
                    ws.Cells(novaLinha, 1).Value = valorData
                    ws.Cells(novaLinha, 2).Value = historico
                    ws.Cells(novaLinha, 3).Value = ws.Cells(i, colValor2).Value * -1
                    novaLinha = novaLinha + 1
                End If
            End If

            ' INSS
            If colINSS > 0 Then
                If IsNumeric(ws.Cells(i, colINSS).Value) Then
                    If ws.Cells(i, colINSS).Value <> 0 Then
                        ws.Cells(novaLinha, 1).Value = valorData
                        ws.Cells(novaLinha, 2).Value = historico
                        ws.Cells(novaLinha, 3).Value = ws.Cells(i, colINSS).Value
                        novaLinha = novaLinha + 1
                    End If
                End If
            End If

            ' IRRF
            If colIRRF > 0 Then
                If IsNumeric(ws.Cells(i, colIRRF).Value) Then
                    If ws.Cells(i, colIRRF).Value <> 0 Then
                        ws.Cells(novaLinha, 1).Value = valorData
                        ws.Cells(novaLinha, 2).Value = historico
                        ws.Cells(novaLinha, 3).Value = ws.Cells(i, colIRRF).Value
                        novaLinha = novaLinha + 1
                    End If
                End If
            End If

            ' DESCONTO
            If colDesconto > 0 Then
                If IsNumeric(ws.Cells(i, colDesconto).Value) Then
                    If ws.Cells(i, colDesconto).Value <> 0 Then
                        ws.Cells(novaLinha, 1).Value = valorData
                        ws.Cells(novaLinha, 2).Value = historico
                        ws.Cells(novaLinha, 3).Value = ws.Cells(i, colDesconto).Value
                        novaLinha = novaLinha + 1
                    End If
                End If
            End If
        End If
    Next i

    MsgBox "Dados organizados com sucesso!", vbInformation

End Sub

```

- Tamanho: 4782 bytes.

### `excel_vba-main/moverdados.vba`

- Tipo: **Texto** (possivelmente script/descrição).

- Trecho inicial (até 5000 caracteres):

```vb
'este modulo move dados de uma matriz na aba 'A' para uma aba 'B'
Sub Planilha_Extrato_CopiarValoresParaPlanilhaModelo()
    Sheets("EXTRATO").Select
    Range("A1").Select
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaLinha As Long
    Dim intervaloOrigem As Range
    Dim intervaloDestino As Range
    
    ' Definindo as planilhas
    Set wsOrigem = ThisWorkbook.ActiveSheet
    Set wsDestino = ThisWorkbook.Sheets("PLANILHA_MODELO")
    ' Encontrar a última linha não vazia na coluna L a P (última linha da matriz)
    ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "L").End(xlUp).Row
    If wsOrigem.Cells(wsOrigem.Rows.Count, "M").End(xlUp).Row > ultimaLinha Then ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "M").End(xlUp).Row
    If wsOrigem.Cells(wsOrigem.Rows.Count, "N").End(xlUp).Row > ultimaLinha Then ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "N").End(xlUp).Row
    If wsOrigem.Cells(wsOrigem.Rows.Count, "O").End(xlUp).Row > ultimaLinha Then ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "O").End(xlUp).Row
    If wsOrigem.Cells(wsOrigem.Rows.Count, "P").End(xlUp).Row > ultimaLinha Then ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "P").End(xlUp).Row
    
    ' Definindo intervalo de origem (valores de L3 até Púltima)
    Set intervaloOrigem = wsOrigem.Range("E3:I" & ultimaLinha)
    
    ' Definindo destino a partir da próxima linha vazia em PLANILHA_MODELO, coluna B
    Dim linhaDestino As Long
    linhaDestino = wsDestino.Cells(wsDestino.Rows.Count, "B").End(xlUp).Row + 1
    If linhaDestino < 2 Then linhaDestino = 2 ' Garante que comece em pelo menos a linha 2
    
    Set intervaloDestino = wsDestino.Range("B" & linhaDestino & ":F" & (linhaDestino + intervaloOrigem.Rows.Count - 1))
    
    ' Copiar apenas os VALORES (não fórmulas)
    intervaloDestino.Value = intervaloOrigem.Value
    Sheets("PLANILHA_MODELO").Select
    Range("A1").Select
    MsgBox "Os novos dados foram enviados para a Planilha Modelo!", vbInformation
End Sub

```

- Tamanho: 2028 bytes.

### `excel_vba-main/remover_linhas_vazias.vba`

- Tipo: **Texto** (possivelmente script/descrição).

- Trecho inicial (até 5000 caracteres):

```vb
'Esta planilha remove celulas vazias dos dados movendo eles pra cima'
'OBS: Esse codigo pode misturar conteudos de uma linha com outra ao subir os dados'
Sub Remover_Vazias()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim intervalo As Range

    Set ws = ActiveSheet
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Aqui você pode definir a coluna que começa e termina'
    For i = ultimaLinha To 3 Step -1
        Set intervalo = ws.Range("A" & i & ":J" & i)
        
        If Application.WorksheetFunction.CountA(intervalo) = 0 Then
            ws.Rows(i).Delete
        End If
    Next i

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Linhas com colunas A até J totalmente vazias foram removidas com segurança.", vbInformation
End Sub

```

- Tamanho: 942 bytes.

### `excel_vba-main/removersinal_cred_deb.vba`

- Tipo: **Texto** (possivelmente script/descrição).

- Trecho inicial (até 5000 caracteres):

```vb
'Este modulo remove os " D" e " C" no final dos valores em um extrato bancario'
Sub Planilha_Extrato_SepararCreditoDebito()
    Dim ws As Worksheet
    Dim i As Long
    Dim ultimaLinha As Long
    Dim credito As Variant
    Dim debito As Variant

    Set ws = ActiveSheet
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Percorre de baixo para cima a partir da linha 3
    For i = ultimaLinha To 3 Step -1
        credito = ws.Cells(i, 3).Value ' Coluna C
        debito = ws.Cells(i, 4).Value  ' Coluna D

        If Not IsEmpty(credito) And Not IsEmpty(debito) Then
            ' Insere nova linha abaixo
            ws.Rows(i + 1).Insert Shift:=xlDown

            ' Copia conteúdo da linha original
            ws.Rows(i).Copy Destination:=ws.Rows(i + 1)

            ' Na linha original, zera débito
            ws.Cells(i, 4).Value = ""

            ' Na linha nova, zera crédito
            ws.Cells(i + 1, 3).Value = ""
        End If
    Next i

    MsgBox "Processo concluído com sucesso!", vbInformation
End Sub

```

- Tamanho: 1047 bytes.

---

## Instruções gerais e recomendações

- Se o seu projeto contém macros dentro de arquivos `.xlsm` ou `.xlsb`, para versioná-las no GitHub recomendamos **exportar os módulos VBA** como `.bas`, `.cls`, `.frm` e adicioná-los ao repositório.
  - No Excel: `ALT+F11` -> no Project Explorer, clique com o botão direito no módulo/classe/form -> `Export File...`.
- Para editar/vistar o código VBA em texto fora do Excel, use ferramentas como `VBA Editor`, `Office Developer Tools`, ou utilitários de terceiros que possam extrair `vbaProject.bin`.
- Mantenha um `README.md` (este arquivo) no repositório com instruções de instalação/uso, dependências, e como executar macros.
- Se desejar, eu posso continuar e:
  - Separar módulos VBA exportados (se você exportar) e documentar cada função/sub.
  - Gerar um README formatado com exemplos de uso e screenshots.
