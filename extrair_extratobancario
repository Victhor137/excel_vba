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
