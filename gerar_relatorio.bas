Attribute VB_Name = "Módulo3"
Sub automatizar_relatório()


Dim linha, linha2 As Integer
linha = 2
linha2 = 8

While Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha, 21) <> "":
    Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 13) = Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha, 21)
    Call AjustarCorretorasDestaques
    Call AjustarCorretorasDestaques
    While Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 21) <> ""
        If Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 21) = "VERDADEIRO" Then
            linha2 = linha2 + 1
        Else:
            MsgBox ("tem algo errado")
    linha2 = linha2 + 1
    Wend
Call exportar3
linha = linha + 1
Wend

    
    
End Sub
