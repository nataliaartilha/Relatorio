Attribute VB_Name = "Módulo6"
Public Sub BODB()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14) = "BODB"
Worksheets("INTRADAY").Activate
    Application.Run "RefreshEntireWorksheet"
    Application.Run "RefreshAllWorkbooks"
    Application.Run "RefreshAllStaticData"
    Application.OnTime (Now + TimeValue("00:00:20")), "ajustar_BODB"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub
Sub ajustar_BODB()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Call AjustarCorretorasDestaques
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Application.OnTime (Now + TimeValue("00:00:20")), "imprimir_BODB"
Application.CutCopyMode = True
Application.ScreenUpdating = True

End Sub
Sub imprimir_BODB()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Call AjustarCorretorasDestaques
'MsgBox ("Deu td certo!")
While Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14) <> ""
       If Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "VERDADEIRO" Or Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "não teve operação" Then
            linha2 = linha2 + 1
            contador = contador + 1
        Else:
            MsgBox ("tem algo errado na linha" & linha2)
            linha2 = linha2 + 1
        End If
    Wend
    
    If contador = 11 Then
        Call exportar3
        nome_fundo = Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14).Value
        'MsgBox ("foi impresso fundo: " & nome_fundo)
        linha = linha + 1
    Else:
        linha = linha + 1
      
        MsgBox ("Não imprimiu o fundo: " & nome_fundo & " pulou para o próximo")
    End If

linha2 = 8

Application.OnTime (Now + TimeValue("00:00:01")), "BIDB"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub

Public Sub BIDB()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14) = "BIDB"
Worksheets("INTRADAY").Activate
    Application.Run "RefreshEntireWorksheet"
    Application.Run "RefreshAllWorkbooks"
    Application.Run "RefreshAllStaticData"
    Application.OnTime (Now + TimeValue("00:00:20")), "ajustar_BIDB"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub
Sub ajustar_BIDB()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Call AjustarCorretorasDestaques
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Application.OnTime (Now + TimeValue("00:00:20")), "imprimir_BIDB"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub
Sub imprimir_BIDB()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Call AjustarCorretorasDestaques
'MsgBox ("Deu td certo!")
While Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14) <> ""
       If Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "VERDADEIRO" Or Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "não teve operação" Then
            linha2 = linha2 + 1
            contador = contador + 1
        Else:
            MsgBox ("tem algo errado na linha" & linha2)
            linha2 = linha2 + 1
        End If
    Wend
    
    If contador = 11 Then
        Call exportar3
        nome_fundo = Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14).Value
        'MsgBox ("foi impresso fundo: " & nome_fundo)
        linha = linha + 1
    Else:
        linha = linha + 1
      
        MsgBox ("Não imprimiu o fundo: " & nome_fundo & " pulou para o próximo")
    End If

linha2 = 8

Application.OnTime (Now + TimeValue("00:00:01")), "ITIP"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub

Public Sub ITIP()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14) = "ITIP"
Worksheets("INTRADAY").Activate
    Application.Run "RefreshEntireWorksheet"
    Application.Run "RefreshAllWorkbooks"
    Application.Run "RefreshAllStaticData"
    Application.OnTime (Now + TimeValue("00:00:20")), "ajustar_ITIP"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub
Sub ajustar_ITIP()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Call AjustarCorretorasDestaques
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Application.OnTime (Now + TimeValue("00:00:20")), "imprimir_ITIP"
Application.CutCopyMode = True
Application.ScreenUpdating = True

End Sub
Sub imprimir_ITIP()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Call AjustarCorretorasDestaques
'MsgBox ("Deu td certo!")
While Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14) <> ""
       If Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "VERDADEIRO" Or Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "não teve operação" Then
            linha2 = linha2 + 1
            contador = contador + 1
        Else:
            MsgBox ("tem algo errado na linha" & linha2)
            linha2 = linha2 + 1
        End If
    Wend
    
    If contador = 11 Then
        Call exportar3
        nome_fundo = Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14).Value
        'MsgBox ("foi impresso fundo: " & nome_fundo)
        linha = linha + 1
    Else:
        linha = linha + 1
      
        MsgBox ("Não imprimiu o fundo: " & nome_fundo & " pulou para o próximo")
    End If

linha2 = 8

Application.OnTime (Now + TimeValue("00:00:01")), "ITIT"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub
Public Sub ITIT()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14) = "ITIT"
Worksheets("INTRADAY").Activate
    Application.Run "RefreshEntireWorksheet"
    Application.Run "RefreshAllWorkbooks"
    Application.Run "RefreshAllStaticData"
    Application.OnTime (Now + TimeValue("00:00:20")), "ajustar_ITIT"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub
Sub ajustar_ITIT()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Call AjustarCorretorasDestaques
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Application.OnTime (Now + TimeValue("00:00:20")), "imprimir_ITIT"
Application.CutCopyMode = True
Application.ScreenUpdating = True

End Sub
Sub imprimir_ITIT()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Call AjustarCorretorasDestaques
'MsgBox ("Deu td certo!")
While Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14) <> ""
       If Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "VERDADEIRO" Or Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "não teve operação" Then
            linha2 = linha2 + 1
            contador = contador + 1
        Else:
            MsgBox ("tem algo errado na linha" & linha2)
            linha2 = linha2 + 1
        End If
    Wend
    
    If contador = 11 Then
        Call exportar3
        nome_fundo = Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14).Value
        'MsgBox ("foi impresso fundo: " & nome_fundo)
        linha = linha + 1
    Else:
        linha = linha + 1
      
        MsgBox ("Não imprimiu o fundo: " & nome_fundo & " pulou para o próximo")
    End If

linha2 = 8

Application.OnTime (Now + TimeValue("00:00:01")), "SADI"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub
Public Sub SADI()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14) = "SADI"
Worksheets("INTRADAY").Activate
    Application.Run "RefreshEntireWorksheet"
    Application.Run "RefreshAllWorkbooks"
    Application.Run "RefreshAllStaticData"
    Application.OnTime (Now + TimeValue("00:00:20")), "ajustar_SADI"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub
Sub ajustar_SADI()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Call AjustarCorretorasDestaques
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Application.OnTime (Now + TimeValue("00:00:20")), "imprimir_SADI"
Application.CutCopyMode = True
Application.ScreenUpdating = True

End Sub
Sub imprimir_SADI()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Call AjustarCorretorasDestaques
'MsgBox ("Deu td certo!")
While Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14) <> ""
       If Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "VERDADEIRO" Or Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "não teve operação" Then
            linha2 = linha2 + 1
            contador = contador + 1
        Else:
            MsgBox ("tem algo errado na linha" & linha2)
            linha2 = linha2 + 1
        End If
    Wend
    
    If contador = 11 Then
        Call exportar3
        nome_fundo = Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14).Value
        'MsgBox ("foi impresso fundo: " & nome_fundo)
        linha = linha + 1
    Else:
        linha = linha + 1
      
        MsgBox ("Não imprimiu o fundo: " & nome_fundo & " pulou para o próximo")
    End If

linha2 = 8

Application.OnTime (Now + TimeValue("00:00:01")), "SARE"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub
Public Sub SARE()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14) = "SARE"
Worksheets("INTRADAY").Activate
    Application.Run "RefreshEntireWorksheet"
    Application.Run "RefreshAllWorkbooks"
    Application.Run "RefreshAllStaticData"
    Application.OnTime (Now + TimeValue("00:00:20")), "ajustar_SARE"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub
Sub ajustar_SARE()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Call AjustarCorretorasDestaques
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Application.OnTime (Now + TimeValue("00:00:20")), "imprimir_SARE"
Application.CutCopyMode = True
Application.ScreenUpdating = True

End Sub
Sub imprimir_SARE()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Call AjustarCorretorasDestaques
'MsgBox ("Deu td certo!")
While Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14) <> ""
       If Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "VERDADEIRO" Or Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "não teve operação" Then
            linha2 = linha2 + 1
            contador = contador + 1
        Else:
            MsgBox ("tem algo errado na linha" & linha2)
            linha2 = linha2 + 1
        End If
    Wend
    
    If contador = 11 Then
        Call exportar3
        nome_fundo = Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14).Value
        'MsgBox ("foi impresso fundo: " & nome_fundo)
        linha = linha + 1
    Else:
        linha = linha + 1
      
        MsgBox ("Não imprimiu o fundo: " & nome_fundo & " pulou para o próximo")
    End If

linha2 = 8

Application.OnTime (Now + TimeValue("00:00:01")), "SPXS"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub

Public Sub SPXS()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14) = "SPXS"
Worksheets("INTRADAY").Activate
    Application.Run "RefreshEntireWorksheet"
    Application.Run "RefreshAllWorkbooks"
    Application.Run "RefreshAllStaticData"
    Application.OnTime (Now + TimeValue("00:00:20")), "ajustar_SPXS"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub
Sub ajustar_SPXS()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Call AjustarCorretorasDestaques
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Application.OnTime (Now + TimeValue("00:00:20")), "imprimir_SPXS"
Application.CutCopyMode = True
Application.ScreenUpdating = True

End Sub
Sub imprimir_SPXS()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Call AjustarCorretorasDestaques
'MsgBox ("Deu td certo!")
While Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14) <> ""
       If Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "VERDADEIRO" Or Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "não teve operação" Then
            linha2 = linha2 + 1
            contador = contador + 1
        Else:
            MsgBox ("tem algo errado na linha" & linha2)
            linha2 = linha2 + 1
        End If
    Wend
    
    If contador = 11 Then
        Call exportar3
        nome_fundo = Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14).Value
        'MsgBox ("foi impresso fundo: " & nome_fundo)
        linha = linha + 1
    Else:
        linha = linha + 1
      
        MsgBox ("Não imprimiu o fundo: " & nome_fundo & " pulou para o próximo")
    End If

linha2 = 8

Application.OnTime (Now + TimeValue("00:00:01")), "VIUR"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub

Public Sub VIUR()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14) = "VIUR"
Worksheets("INTRADAY").Activate
    Application.Run "RefreshEntireWorksheet"
    Application.Run "RefreshAllWorkbooks"
    Application.Run "RefreshAllStaticData"
    Application.OnTime (Now + TimeValue("00:00:20")), "ajustar_VIUR"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub
Sub ajustar_VIUR()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Call AjustarCorretorasDestaques
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Application.OnTime (Now + TimeValue("00:00:20")), "imprimir_VIUR"
Application.CutCopyMode = True
Application.ScreenUpdating = True

End Sub
Sub imprimir_VIUR()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Call AjustarCorretorasDestaques
'MsgBox ("Deu td certo!")
While Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14) <> ""
       If Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "VERDADEIRO" Or Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "não teve operação" Then
            linha2 = linha2 + 1
            contador = contador + 1
        Else:
            MsgBox ("tem algo errado na linha" & linha2)
            linha2 = linha2 + 1
        End If
    Wend
    
    If contador = 11 Then
        Call exportar3
        nome_fundo = Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14).Value
        'MsgBox ("foi impresso fundo: " & nome_fundo)
        linha = linha + 1
    Else:
        linha = linha + 1
      
        MsgBox ("Não imprimiu o fundo: " & nome_fundo & " pulou para o próximo")
    End If

linha2 = 8

Application.OnTime (Now + TimeValue("00:00:01")), "WHGR"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub
Public Sub WHGR()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14) = "WHGR"
Worksheets("INTRADAY").Activate
    Application.Run "RefreshEntireWorksheet"
    Application.Run "RefreshAllWorkbooks"
    Application.Run "RefreshAllStaticData"
    Application.OnTime (Now + TimeValue("00:00:20")), "ajustar_WHGR"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub
Sub ajustar_WHGR()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Call AjustarCorretorasDestaques
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Application.OnTime (Now + TimeValue("00:00:20")), "imprimir_WHGR"
Application.CutCopyMode = True
Application.ScreenUpdating = True

End Sub
Sub imprimir_WHGR()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Call AjustarCorretorasDestaques
'MsgBox ("Deu td certo!")
While Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14) <> ""
       If Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "VERDADEIRO" Or Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "não teve operação" Then
            linha2 = linha2 + 1
            contador = contador + 1
        Else:
            MsgBox ("tem algo errado na linha" & linha2)
            linha2 = linha2 + 1
        End If
    Wend
    
    If contador = 11 Then
        Call exportar3
        nome_fundo = Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14).Value
        'MsgBox ("foi impresso fundo: " & nome_fundo)
        linha = linha + 1
    Else:
        linha = linha + 1
      
        MsgBox ("Não imprimiu o fundo: " & nome_fundo & " pulou para o próximo")
    End If

linha2 = 8

Application.OnTime (Now + TimeValue("00:00:01")), "XPID"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub

Public Sub XPID()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14) = "XPID"
Worksheets("INTRADAY").Activate
    Application.Run "RefreshEntireWorksheet"
    Application.Run "RefreshAllWorkbooks"
    Application.Run "RefreshAllStaticData"
    Application.OnTime (Now + TimeValue("00:00:20")), "ajustar_XPID"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub
Sub ajustar_XPID()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Call AjustarCorretorasDestaques
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Application.OnTime (Now + TimeValue("00:00:20")), "imprimir_XPID"
Application.CutCopyMode = True
Application.ScreenUpdating = True

End Sub
Sub imprimir_XPID()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Call AjustarCorretorasDestaques
'MsgBox ("Deu td certo!")
While Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14) <> ""
       If Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "VERDADEIRO" Or Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "não teve operação" Then
            linha2 = linha2 + 1
            contador = contador + 1
        Else:
            MsgBox ("tem algo errado na linha" & linha2)
            linha2 = linha2 + 1
        End If
    Wend
    
    If contador = 11 Then
        Call exportar3
        nome_fundo = Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14).Value
        'MsgBox ("foi impresso fundo: " & nome_fundo)
        linha = linha + 1
    Else:
        linha = linha + 1
      
        MsgBox ("Não imprimiu o fundo: " & nome_fundo & " pulou para o próximo")
    End If

linha2 = 8

Application.OnTime (Now + TimeValue("00:00:01")), "XPIE"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub

Public Sub XPIE()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14) = "XPIE"
Worksheets("INTRADAY").Activate
    Application.Run "RefreshEntireWorksheet"
    Application.Run "RefreshAllWorkbooks"
    Application.Run "RefreshAllStaticData"
    Application.OnTime (Now + TimeValue("00:00:20")), "ajustar_XPIE"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub
Sub ajustar_XPIE()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Call AjustarCorretorasDestaques
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Application.OnTime (Now + TimeValue("00:00:20")), "imprimir_XPIE"
Application.CutCopyMode = True
Application.ScreenUpdating = True

End Sub
Sub imprimir_XPIE()
Application.CutCopyMode = False
Application.ScreenUpdating = False
Dim nome_fundo, starttime, endtime As String
Dim linha, linha2, contador As Integer
linha = 2
linha2 = 8
contador = 0
Application.Run "RefreshAllWorkbooks"
Application.Run "RefreshAllStaticData"
Call AjustarCorretorasDestaques
'MsgBox ("Deu td certo!")
While Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14) <> ""
       If Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "VERDADEIRO" Or Worksheets("RELATÓRIO 5 CORRETORAS").Cells(linha2, 14).Value = "não teve operação" Then
            linha2 = linha2 + 1
            contador = contador + 1
        Else:
            MsgBox ("tem algo errado na linha" & linha2)
            linha2 = linha2 + 1
        End If
    Wend
    
    If contador = 11 Then
        Call exportar3
        nome_fundo = Worksheets("RELATÓRIO 5 CORRETORAS").Cells(1, 14).Value
        'MsgBox ("foi impresso fundo: " & nome_fundo)
        linha = linha + 1
    Else:
        linha = linha + 1
      
        MsgBox ("Não imprimiu o fundo: " & nome_fundo & " pulou para o próximo")
    End If

linha2 = 8
MsgBox ("Atualizou todos relatórios")
'Application.OnTime (Now + TimeValue("00:00:20")), "XPIE"
Application.CutCopyMode = True
Application.ScreenUpdating = True
End Sub
