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
