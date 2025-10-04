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
