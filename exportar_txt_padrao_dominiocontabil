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
