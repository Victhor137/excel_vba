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
