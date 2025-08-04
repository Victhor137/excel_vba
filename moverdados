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
