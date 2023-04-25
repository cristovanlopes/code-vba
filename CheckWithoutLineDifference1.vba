Sub CompararPlanilhas()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim r1 As Range, r2 As Range
    Dim c1 As Range, c2 As Range
    Dim diffCount As Long
    Dim maxRows As Long, maxCols As Long, row As Long, col As Long

    ' Defina aqui os nomes dos arquivos e das planilhas que você deseja comparar
    Set ws1 = Workbooks("ARQUIVO1").Worksheets("NEGATIVOS")
    Set ws2 = Workbooks("ARQUIVO2").Worksheets("NEGATIVOS")

    ' Defina os intervalos a serem comparados (neste exemplo, comparamos a planilha inteira)
    Set r1 = ws1.UsedRange
    Set r2 = ws2.UsedRange

    ' Determine o tamanho das áreas comuns das duas planilhas
    maxRows = Application.WorksheetFunction.Min(r1.Rows.Count, r2.Rows.Count)
    maxCols = Application.WorksheetFunction.Min(r1.Columns.Count, r2.Columns.Count)

    ' Inicialize o contador de diferenças
    diffCount = 0

    ' Compare as células nas áreas comuns das planilhas
    For row = 1 To maxRows
        For col = 1 To maxCols
            Set c1 = r1.Cells(row, col)
            Set c2 = r2.Cells(row, col)

            ' Verifique se os tipos de dados das células são compatíveis
            If TypeName(c1.Value) = TypeName(c2.Value) Then
                If c1.Value <> c2.Value Then
                    c1.Interior.Color = RGB(255, 0, 0) ' Destaque a célula com diferença na Planilha1 com cor vermelha
                    c2.Interior.Color = RGB(255, 0, 0) ' Destaque a célula com diferença na Planilha2 com cor vermelha
                    diffCount = diffCount + 1
                Else
                    c1.Interior.Color = xlNone ' Remova a cor de fundo se as células forem iguais
                    c2.Interior.Color = xlNone ' Remova a cor de fundo se as células forem iguais
                End If
            Else
                c1.Interior.Color = RGB(255, 255, 0) ' Destaque a célula com tipo de dado diferente na Planilha1 com cor amarela
                c2.Interior.Color = RGB(255, 255, 0) ' Destaque a célula com tipo de dado diferente na Planilha2 com cor amarela
                diffCount = diffCount + 1
            End If
        Next col
    Next row

    ' Exibir uma mensagem com o número de diferenças encontradas
    If diffCount = 0 Then
        MsgBox "Nenhuma diferença encontrada!"
    Else
        MsgBox diffCount & " diferenças encontradas!"
    End If
End Sub
