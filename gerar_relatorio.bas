Attribute VB_Name = "M�dulo3"
Sub automatizar_relat�rio()


Dim linha, linha2 As Integer
linha = 2
linha2 = 8

While Worksheets("RELAT�RIO 5 CORRETORAS").Cells(linha, 21) <> "":
    Worksheets("RELAT�RIO 5 CORRETORAS").Cells(1, 13) = Worksheets("RELAT�RIO 5 CORRETORAS").Cells(linha, 21)
    Call AjustarCorretorasDestaques
    Call AjustarCorretorasDestaques
    While Worksheets("RELAT�RIO 5 CORRETORAS").Cells(linha2, 21) <> ""
        If Worksheets("RELAT�RIO 5 CORRETORAS").Cells(linha2, 21) = "VERDADEIRO" Then
            linha2 = linha2 + 1
        Else:
            MsgBox ("tem algo errado")
    linha2 = linha2 + 1
    Wend
Call exportar3
linha = linha + 1
Wend

    
    
End Sub
