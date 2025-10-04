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
