Attribute VB_Name = "M�dulo2"
'essa � a certa
Sub exportar3()
Dim fundo, codigo, mes As String
Dim ano As Integer
Dim LocalPDF As String
Dim intervalo As Range
Dim fundo1 As String

codigo = Worksheets("RELAT�RIO 5 CORRETORAS").Range("N1")
fundo = Worksheets("RELAT�RIO 5 CORRETORAS").Range("V1")
ano = Worksheets("RELAT�RIO 5 CORRETORAS").Range("Q1")
mes = Worksheets("RELAT�RIO 5 CORRETORAS").Range("S1")
fundo1 = Cells(3, 14).Value


LocalPDF = "G:\depto\RENDA\Formador de Mercado\FUNDOS\" & fundo & "\" & codigo & "\RELAT�RIOS\" & ano & "\" & mes & "\" & codigo & " " & Format(Range("N2"), "dd.mm.yyyy") & ".pdf"

If codigo = "SPXS" Then
Set intervalo = Sheets("RELAT�RIO 5 CORRETORAS").Range("A1:K342")
intervalo.ExportAsFixedFormat Type:=xlTypePDF, _
Filename:=LocalPDF
Else
Set intervalo = Sheets("RELAT�RIO 5 CORRETORAS").Range("A1:K285")
intervalo.ExportAsFixedFormat Type:=xlTypePDF, _
Filename:=LocalPDF
End If
End Sub

