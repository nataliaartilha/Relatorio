Attribute VB_Name = "Módulo2_Emails"
Sub Email_Geral()

Dim EmailTo As Range, cl As Range
Dim sTo As String
Dim mes, ano, gestora, fundo As String
Dim data As Date

mes = Worksheets("RELATÓRIO 5 CORRETORAS").Range("S1")
ano = Worksheets("INTRADAY").Range("B6")
gestora = Worksheets("RELATÓRIO 5 CORRETORAS").Range("V1")
fundo = Worksheets("INTRADAY").Range("B1")
data = Worksheets("INTRADAY").Range("B2")

Set objeto_outlook = CreateObject("Outlook.Application")
Set Email = objeto_outlook.createitem(0)

    Set EmailTo = Worksheets("EMAILS").Range("A2", Range("A100").End(xlUp))

    For Each cl In EmailTo
        sTo = sTo & ";" & cl.Value
    Next
    
Email.display

Email.to = sTo
Email.cc = "BancoFatorTesouraria@fator.com.br"
'Email.bcc

Email.Subject = "[Fator] - Acompanhamento de Mercado Secundário - " & fundo & "11 - " & Format(data, "DD.MM.YYYY")

Email.Body = "Prezado(a)," & Chr(10) & Chr(10) _
& "Segue relatório de acompanhamento diário de mercado secundário do " & fundo & "11 referente ao dia " & Format(data, "DD/MMMM/YYYY") & "." & Chr(10) & Chr(10) _
& "Atenciosamente,"
Email.Attachments.Add "G:\depto\RENDA\Formador de Mercado\FUNDOS\" & gestora & "\" & fundo & "\RELATÓRIOS\" & ano & "\" & mes & "\" & fundo & " " & Format(data, "DD.MM.YYYY") & ".pdf"
'Email.Attachments.Add "G:\depto\RENDA\Formador de Mercado\FUNDOS\     INTER     \    BIDB     \RELATÓRIOS\   2022    \  05. Maio \    BIDB          11.05.2022               .pdf"

End Sub

Sub chamarEmails()

Call Email_BODB
Call Email_ITIT_ITIP
'Call Email_JSAF
Call Email_SADI_SARE
Call Email_SPSX
Call Email_VIUR
Call Email_WHGR
Call Email_XPID_XPIE
Call Email_BIDB

End Sub

Sub Email_BIDB()

Dim EmailTo As Range, cl As Range
Dim sTo As String
Dim mes, ano, gestora, fundo, mensagem1, assinatura As String

mes = Worksheets("RELATÓRIO 5 CORRETORAS").Range("S1")
ano = Worksheets("INTRADAY").Range("B6")
gestora = Worksheets("RELATÓRIO 5 CORRETORAS").Range("V1")
fundo = Worksheets("INTRADAY").Range("B1")
data = Worksheets("INTRADAY").Range("B2")

Set objeto_outlook = CreateObject("Outlook.Application")
Set Email = objeto_outlook.createitem(0)

Worksheets("EMAILS").Activate

Set EmailTo = Worksheets("EMAILS").Range("A2", Range("A100").End(xlUp))
    For Each cl In EmailTo
        sTo = sTo & ";" & cl.Value
    Next

Worksheets("RELATÓRIO 5 CORRETORAS").Activate

Email.display
Email.to = sTo
Email.cc = "BancoFatorTesouraria@fator.com.br"
Email.Subject = "[Fator] - Acompanhamento de Mercado Secundário - BIDB11 - " & Format(data, "DD.MM.YYYY")

assinatura = Email.HTMLBody

Email.Body = "Prezado(a)," & Chr(10) & Chr(10) _
& "Segue relatório de acompanhamento diário de mercado secundário do BIDB11 referente ao dia " & Format(data, "DD/MMMM/YYYY") & "." & Chr(10) & Chr(10) _
& "Atenciosamente,"

mensagem1 = Email.HTMLBody

Email.HTMLBody = mensagem1 & assinatura

Email.Attachments.Add "G:\depto\RENDA\Formador de Mercado\FUNDOS\INTER\BIDB\RELATÓRIOS\" & ano & "\" & mes & "\BIDB " & Format(data, "DD.MM.YYYY") & ".pdf"
'Email.Attachments.Add "G:\depto\RENDA\Formador de Mercado\FUNDOS\INTER\BIDB\RELATÓRIOS\   2022   \ 05. Maio  \BIDB           11.05.2022               .pdf"

Email.Attachments.Add "G:\depto\RENDA\Formador de Mercado\FUNDOS\INTER\BIDB\RELATÓRIOS\" & ano & "\" & mes & "\INFORMATIVO_BIDB_" & Format(data, "DD.MM.YYYY") & ".XLS"

End Sub

Sub Email_ITIT_ITIP()

Dim EmailTo As Range, cl As Range
Dim sTo As String
Dim mes, ano, gestora, fundo As String

mes = Worksheets("RELATÓRIO 5 CORRETORAS").Range("S1")
ano = Worksheets("INTRADAY").Range("B6")
gestora = Worksheets("RELATÓRIO 5 CORRETORAS").Range("V1")
fundo = Worksheets("INTRADAY").Range("B1")
data = Worksheets("INTRADAY").Range("B2")

Set objeto_outlook = CreateObject("Outlook.Application")
Set Email = objeto_outlook.createitem(0)

Worksheets("EMAILS").Activate

Set EmailTo = Worksheets("EMAILS").Range("B2", Range("B100").End(xlUp))
    For Each cl In EmailTo
        sTo = sTo & ";" & cl.Value
    Next

Worksheets("RELATÓRIO 5 CORRETORAS").Activate

Email.display
Email.to = sTo
Email.cc = "BancoFatorTesouraria@fator.com.br"
Email.Subject = "[Fator] - Acompanhamento de Mercado Secundário - ITIT11 e ITIP11 - " & Format(data, "DD.MM.YYYY")

assinatura = Email.HTMLBody

Email.Body = "Prezado(a)," & Chr(10) & Chr(10) _
& "Segue relatório de acompanhamento diário de mercado secundário do ITIT11 e ITIP11 referente ao dia " & Format(data, "DD/MMMM/YYYY") & "." & Chr(10) & Chr(10) _
& "Atenciosamente,"

mensagem1 = Email.HTMLBody

Email.HTMLBody = mensagem1 & assinatura

Email.Attachments.Add "G:\depto\RENDA\Formador de Mercado\FUNDOS\INTER\ITIT\RELATÓRIOS\" & ano & "\" & mes & "\ITIT " & Format(data, "DD.MM.YYYY") & ".pdf"
Email.Attachments.Add "G:\depto\RENDA\Formador de Mercado\FUNDOS\INTER\ITIP\RELATÓRIOS\" & ano & "\" & mes & "\ITIP " & Format(data, "DD.MM.YYYY") & ".pdf"

End Sub

Sub Email_JSAF()

Dim EmailTo As Range, cl As Range
Dim sTo As String
Dim mes, ano, gestora, fundo As String

mes = Worksheets("RELATÓRIO 5 CORRETORAS").Range("S1")
ano = Worksheets("INTRADAY").Range("B6")
gestora = Worksheets("RELATÓRIO 5 CORRETORAS").Range("V1")
fundo = Worksheets("INTRADAY").Range("B1")
data = Worksheets("INTRADAY").Range("B2")

Set objeto_outlook = CreateObject("Outlook.Application")
Set Email = objeto_outlook.createitem(0)

Worksheets("EMAILS").Activate

Set EmailTo = Worksheets("EMAILS").Range("C2", Range("C100").End(xlUp))
    For Each cl In EmailTo
        sTo = sTo & ";" & cl.Value
    Next

Worksheets("RELATÓRIO 5 CORRETORAS").Activate

Email.display
Email.to = sTo
Email.cc = "BancoFatorTesouraria@fator.com.br"
Email.Subject = "[Fator] - Acompanhamento de Mercado Secundário - JSAF11 - " & Format(data, "DD.MM.YYYY")

assinatura = Email.HTMLBody

Email.Body = "Prezado(a)," & Chr(10) & Chr(10) _
& "Segue relatório de acompanhamento diário de mercado secundário do JSAF11 referente ao dia " & Format(data, "DD/MMMM/YYYY") & "." & Chr(10) & Chr(10) _
& "Atenciosamente,"

mensagem1 = Email.HTMLBody

Email.HTMLBody = mensagem1 & assinatura

Email.Attachments.Add "G:\depto\RENDA\Formador de Mercado\FUNDOS\SAFRA\JSAF\RELATÓRIOS\" & ano & "\" & mes & "\JSAF " & Format(data, "DD.MM.YYYY") & ".pdf"

End Sub

Sub Email_SADI_SARE()

Dim EmailTo As Range, cl As Range
Dim sTo As String
Dim mes, ano, gestora, fundo As String

mes = Worksheets("RELATÓRIO 5 CORRETORAS").Range("S1")
ano = Worksheets("INTRADAY").Range("B6")
gestora = Worksheets("RELATÓRIO 5 CORRETORAS").Range("V1")
fundo = Worksheets("INTRADAY").Range("B1")
data = Worksheets("INTRADAY").Range("B2")

Set objeto_outlook = CreateObject("Outlook.Application")
Set Email = objeto_outlook.createitem(0)

Worksheets("EMAILS").Activate

Set EmailTo = Worksheets("EMAILS").Range("E2", Range("E100").End(xlUp))
    For Each cl In EmailTo
        sTo = sTo & ";" & cl.Value
    Next

Worksheets("RELATÓRIO 5 CORRETORAS").Activate

Email.display
Email.to = sTo
Email.cc = "BancoFatorTesouraria@fator.com.br"
Email.Subject = "[Fator] - Acompanhamento de Mercado Secundário - SADI11 e SARE11 - " & Format(data, "DD.MM.YYYY")

assinatura = Email.HTMLBody

Email.Body = "Prezado(a)," & Chr(10) & Chr(10) _
& "Segue relatório de acompanhamento diário de mercado secundário do SADI11 e SARE11 referente ao dia " & Format(data, "DD/MMMM/YYYY") & "." & Chr(10) & Chr(10) _
& "Atenciosamente,"

mensagem1 = Email.HTMLBody

Email.HTMLBody = mensagem1 & assinatura

Email.Attachments.Add "G:\depto\RENDA\Formador de Mercado\FUNDOS\SANTANDER\SADI\RELATÓRIOS\" & ano & "\" & mes & "\SADI " & Format(data, "DD.MM.YYYY") & ".pdf"
Email.Attachments.Add "G:\depto\RENDA\Formador de Mercado\FUNDOS\SANTANDER\SARE\RELATÓRIOS\" & ano & "\" & mes & "\SARE " & Format(data, "DD.MM.YYYY") & ".pdf"

End Sub

Sub Email_XPID_XPIE()

Dim EmailTo As Range, cl As Range
Dim sTo As String
Dim mes, ano, gestora, fundo As String

mes = Worksheets("RELATÓRIO 5 CORRETORAS").Range("S1")
ano = Worksheets("INTRADAY").Range("B6")
gestora = Worksheets("RELATÓRIO 5 CORRETORAS").Range("V1")
fundo = Worksheets("INTRADAY").Range("B1")
data = Worksheets("INTRADAY").Range("B2")

Set objeto_outlook = CreateObject("Outlook.Application")
Set Email = objeto_outlook.createitem(0)

Worksheets("EMAILS").Activate

Set EmailTo = Worksheets("EMAILS").Range("F2", Range("F100").End(xlUp))
    For Each cl In EmailTo
        sTo = sTo & ";" & cl.Value
    Next

Worksheets("RELATÓRIO 5 CORRETORAS").Activate

Email.display
Email.to = sTo
Email.cc = "BancoFatorTesouraria@fator.com.br"
Email.Subject = "[Fator] - Acompanhamento de Mercado Secundário - XPID11 e XPIE11 - " & Format(data, "DD.MM.YYYY")

assinatura = Email.HTMLBody

Email.Body = "Prezado(a)," & Chr(10) & Chr(10) _
& "Segue relatório de acompanhamento diário de mercado secundário do XPID11 e XPIE11 referente ao dia " & Format(data, "DD/MMMM/YYYY") & "." & Chr(10) & Chr(10) _
& "Atenciosamente,"

mensagem1 = Email.HTMLBody

Email.HTMLBody = mensagem1 & assinatura

Email.Attachments.Add "G:\depto\RENDA\Formador de Mercado\FUNDOS\XP\XPID\RELATÓRIOS\" & ano & "\" & mes & "\XPID " & Format(data, "DD.MM.YYYY") & ".pdf"
Email.Attachments.Add "G:\depto\RENDA\Formador de Mercado\FUNDOS\XP\XPIE\RELATÓRIOS\" & ano & "\" & mes & "\XPIE " & Format(data, "DD.MM.YYYY") & ".pdf"

End Sub

Sub Email_BODB()

Dim EmailTo As Range, cl As Range
Dim sTo As String
Dim mes, ano, gestora, fundo As String

mes = Worksheets("RELATÓRIO 5 CORRETORAS").Range("S1")
ano = Worksheets("INTRADAY").Range("B6")
gestora = Worksheets("RELATÓRIO 5 CORRETORAS").Range("V1")
fundo = Worksheets("INTRADAY").Range("B1")
data = Worksheets("INTRADAY").Range("B2")

Set objeto_outlook = CreateObject("Outlook.Application")
Set Email = objeto_outlook.createitem(0)

Worksheets("EMAILS").Activate

Set EmailTo = Worksheets("EMAILS").Range("G2", Range("G100").End(xlUp))
    For Each cl In EmailTo
        sTo = sTo & ";" & cl.Value
    Next

Worksheets("RELATÓRIO 5 CORRETORAS").Activate

Email.display
Email.to = sTo
Email.cc = "BancoFatorTesouraria@fator.com.br"
Email.Subject = "[Fator] - Acompanhamento de Mercado Secundário - BODB11 - " & Format(data, "DD.MM.YYYY")

assinatura = Email.HTMLBody

Email.Body = "Prezado(a)," & Chr(10) & Chr(10) _
& "Segue relatório de acompanhamento diário de mercado secundário do BODB11 referente ao dia " & Format(data, "DD/MMMM/YYYY") & "." & Chr(10) & Chr(10) _
& "Atenciosamente,"

mensagem1 = Email.HTMLBody

Email.HTMLBody = mensagem1 & assinatura

Email.Attachments.Add "G:\depto\RENDA\Formador de Mercado\FUNDOS\BOCAINA\BODB\RELATÓRIOS\" & ano & "\" & mes & "\BODB " & Format(data, "DD.MM.YYYY") & ".pdf"

End Sub

Sub Email_VIUR()

Dim EmailTo As Range, cl As Range
Dim sTo As String
Dim mes, ano, gestora, fundo As String

mes = Worksheets("RELATÓRIO 5 CORRETORAS").Range("S1")
ano = Worksheets("INTRADAY").Range("B6")
gestora = Worksheets("RELATÓRIO 5 CORRETORAS").Range("V1")
fundo = Worksheets("INTRADAY").Range("B1")
data = Worksheets("INTRADAY").Range("B2")

Set objeto_outlook = CreateObject("Outlook.Application")
Set Email = objeto_outlook.createitem(0)

Worksheets("EMAILS").Activate

Set EmailTo = Worksheets("EMAILS").Range("H2", Range("H100").End(xlUp))
    For Each cl In EmailTo
        sTo = sTo & ";" & cl.Value
    Next

Worksheets("RELATÓRIO 5 CORRETORAS").Activate

Email.display
Email.to = sTo
Email.cc = "BancoFatorTesouraria@fator.com.br"
Email.Subject = "[Fator] - Acompanhamento de Mercado Secundário - VIUR11 - " & Format(data, "DD.MM.YYYY")

assinatura = Email.HTMLBody

Email.Body = "Prezado(a)," & Chr(10) & Chr(10) _
& "Segue relatório de acompanhamento diário de mercado secundário do VIUR11 referente ao dia " & Format(data, "DD/MMMM/YYYY") & "." & Chr(10) & Chr(10) _
& "Atenciosamente,"

mensagem1 = Email.HTMLBody

Email.HTMLBody = mensagem1 & assinatura

Email.Attachments.Add "G:\depto\RENDA\Formador de Mercado\FUNDOS\VINCI\VIUR\RELATÓRIOS\" & ano & "\" & mes & "\VIUR " & Format(data, "DD.MM.YYYY") & ".pdf"

End Sub

Sub Email_SPSX()

Dim EmailTo As Range, cl As Range
Dim sTo As String
Dim mes, ano, gestora, fundo As String

mes = Worksheets("RELATÓRIO 5 CORRETORAS").Range("S1")
ano = Worksheets("INTRADAY").Range("B6")
gestora = Worksheets("RELATÓRIO 5 CORRETORAS").Range("V1")
fundo = Worksheets("INTRADAY").Range("B1")
data = Worksheets("INTRADAY").Range("B2")

Set objeto_outlook = CreateObject("Outlook.Application")
Set Email = objeto_outlook.createitem(0)

Worksheets("EMAILS").Activate

Set EmailTo = Worksheets("EMAILS").Range("I2", Range("I100").End(xlUp))
    For Each cl In EmailTo
        sTo = sTo & ";" & cl.Value
    Next

Worksheets("RELATÓRIO 5 CORRETORAS").Activate

Email.display
Email.to = sTo
Email.cc = "BancoFatorTesouraria@fator.com.br"
Email.Subject = "[Fator] - Acompanhamento de Mercado Secundário - SPXS11 - " & Format(data, "DD.MM.YYYY")

assinatura = Email.HTMLBody

Email.Body = "Prezado(a)," & Chr(10) & Chr(10) _
& "Segue relatório de acompanhamento diário de mercado secundário do SPXS11 referente ao dia " & Format(data, "DD/MMMM/YYYY") & "." & Chr(10) & Chr(10) _
& "Atenciosamente,"

mensagem1 = Email.HTMLBody

Email.HTMLBody = mensagem1 & assinatura

Email.Attachments.Add "G:\depto\RENDA\Formador de Mercado\FUNDOS\SPX\SPXS\RELATÓRIOS\" & ano & "\" & mes & "\SPXS " & Format(data, "DD.MM.YYYY") & ".pdf"

End Sub
Sub Email_WHGR()

Dim EmailTo As Range, cl As Range
Dim sTo As String
Dim mes, ano, gestora, fundo As String

mes = Worksheets("RELATÓRIO 5 CORRETORAS").Range("S1")
ano = Worksheets("INTRADAY").Range("B6")
gestora = Worksheets("RELATÓRIO 5 CORRETORAS").Range("V1")
fundo = Worksheets("INTRADAY").Range("B1")
data = Worksheets("INTRADAY").Range("B2")

Set objeto_outlook = CreateObject("Outlook.Application")
Set Email = objeto_outlook.createitem(0)

Worksheets("EMAILS").Activate

Set EmailTo = Worksheets("EMAILS").Range("J2", Range("J100").End(xlUp))
    For Each cl In EmailTo
        sTo = sTo & ";" & cl.Value
    Next

Worksheets("RELATÓRIO 5 CORRETORAS").Activate

Email.display
Email.to = sTo
Email.cc = "BancoFatorTesouraria@fator.com.br"
Email.Subject = "[Fator] - Acompanhamento de Mercado Secundário - WHGR11 - " & Format(data, "DD.MM.YYYY")

assinatura = Email.HTMLBody

Email.Body = "Prezado(a)," & Chr(10) & Chr(10) _
& "Segue relatório de acompanhamento diário de mercado secundário do WHGR11 referente ao dia " & Format(data, "DD/MMMM/YYYY") & "." & Chr(10) & Chr(10) _
& "Atenciosamente,"

mensagem1 = Email.HTMLBody

Email.HTMLBody = mensagem1 & assinatura

Email.Attachments.Add "G:\depto\RENDA\Formador de Mercado\FUNDOS\WHG\WHGR\RELATÓRIOS\" & ano & "\" & mes & "\WHGR " & Format(data, "DD.MM.YYYY") & ".pdf"

End Sub


