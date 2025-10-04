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
