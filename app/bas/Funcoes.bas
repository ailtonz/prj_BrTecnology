Attribute VB_Name = "Funcoes"
Option Compare Database
Option Explicit
Global VarEmail(30, 3)

Public Function Compara(Optional MES1 As Integer, Optional ANO1 As Integer, Optional MES2 As Integer, Optional ANO2 As Integer)

Dim data1, data2 As String

data1 = "01/" & MES1 & "/" & ANO1
data2 = "01/" & MES2 & "/" & ANO2


If MES1 = 0 Or ANO1 = 0 Then
   Compara = True
Else
   Compara = CDate(data1) < CDate(data2)
End If


End Function

Sub teste()

       If EnviarEmail("Boleto.htm", "cleber@brloja.com.br", "Cleber Torezan", "Cobrança de mensalidade") Then
          MsgBox "OK"
       Else
          MsgBox "Erro de envio!"
       End If
          
End Sub

    
Public Function Monta_CodBarras(valor1, vencimento1, Moeda, Banco, Agencia, Conta, Conta_dac, Carteira, nossonumero, dv_nossonumero)
    
    Dim database, fator, Valor, dvcb, codigo_sequencia As String
    
    database = CDate("7/10/1997")
    fator = DateDiff("d", database, vencimento1)
    Valor = Int(valor1 * 100)
    
    Do While Len(Valor) < 10
        Valor = "0" & Valor
    Loop
    
    codigo_sequencia = Banco & Moeda & fator & Valor & Carteira & nossonumero & dv_nossonumero & Agencia & Conta & Conta_dac & "000"
    
    dvcb = calcula_DV_CodBarras(codigo_sequencia)
    
    Monta_CodBarras = Left(codigo_sequencia, 4) & dvcb & Right(codigo_sequencia, 39)
    
End Function
    
Function Linha_Digitavel(sequencia_codigo_barra)
    
    Dim seq1, seq2, seq3, seq4, dvcb, dv1, dv2, dv3
    
    '         10        20        30        40
    '12345678901234567890123456789012345678901234
    '3419 7233 5 00000059 00175 02280204 4 2923 05456 9 000
    '3419 1233 2 00000059 00175 01250204 2 2923 05456 9 000
    
    seq1 = Left(sequencia_codigo_barra, 4) & Mid(sequencia_codigo_barra, 20, 5)
    seq2 = Mid(sequencia_codigo_barra, 25, 10)
    seq3 = Mid(sequencia_codigo_barra, 35, 10)
    seq4 = Mid(sequencia_codigo_barra, 6, 14)
    
    'seq1 = banco & moeda & carteira & left(nossonumero,2)
    'seq2 = right( nossonumero, 6 ) & dv_nossonumero & left( agencia , 3 )
    'seq3 = right( agencia, 1 ) & conta & dv_conta & "000"
    'seq4 = mid(sequencia_codigo_barra,6,14)
    
    dvcb = Mid(sequencia_codigo_barra, 5, 1)
    
    dv1 = Calculo_DV10(seq1)
    dv2 = Calculo_DV10(seq2)
    dv3 = Calculo_DV10(seq3)
    
    seq1 = Left(seq1 & dv1, 5) & "." & Mid(seq1 & dv1, 6, 5)
    seq2 = Left(seq2 & dv2, 5) & "." & Mid(seq2 & dv2, 6, 6)
    seq3 = Left(seq3 & dv3, 5) & "." & Mid(seq3 & dv3, 6, 6)
    
    Linha_Digitavel = seq1 & " " & seq2 & " " & seq3 & " " & dvcb & " " & seq4
    
End Function
    
Function Calculo_DV10(strNumero)
    
    Dim fator, total, numero, resto, i As Integer
    
    fator = 2
    total = 0
    For i = Len(strNumero) To 1 Step -1
        numero = Mid(strNumero, i, 1) * fator
        If numero > 9 Then
            numero = CInt(Left(numero, 1)) + CInt(Right(numero, 1))
        End If
        total = total + numero
        If fator = 2 Then
            fator = 1
        Else
            fator = 2
        End If
    Next
    resto = total Mod 10
    resto = 10 - resto
    If resto = 10 Then
        Calculo_DV10 = 0
    Else
        Calculo_DV10 = resto
    End If
    
End Function
    
Function calcula_DV_CodBarras(sequencia)
    
    Dim fator, total, numero, resto, i, resultado As Integer
    
    fator = 2
    total = 0
    For i = 43 To 1 Step -1
        numero = Val(Mid(sequencia, i, 1))
        If fator > 9 Then
            fator = 2
        End If
        numero = numero * fator
        total = total + numero
        fator = fator + 1
    Next
    resto = total Mod 11
    resultado = 11 - resto
    If resultado = 10 Or resultado = 0 Or resultado = 11 Then
        calcula_DV_CodBarras = 1
    Else
        calcula_DV_CodBarras = resultado
    End If
    
End Function

Public Function WBarCode(Valor)
    Dim f, f1, f2, i
    Dim Texto
    Const fino = 1
    Const largo = 3
    Const altura = 50
    Dim BarCodes(99)
    Dim retorno As String
    
    
    If BarCodes(0) = "" Then
      BarCodes(0) = "00110"
      BarCodes(1) = "10001"
      BarCodes(2) = "01001"
      BarCodes(3) = "11000"
      BarCodes(4) = "00101"
      BarCodes(5) = "10100"
      BarCodes(6) = "01100"
      BarCodes(7) = "00011"
      BarCodes(8) = "10010"
      BarCodes(9) = "01010"
      For f1 = 9 To 0 Step -1
        For f2 = 9 To 0 Step -1
          f = f1 * 10 + f2
          Texto = ""
          For i = 1 To 5
            Texto = Texto & Mid(BarCodes(f1), i, 1) + Mid(BarCodes(f2), i, 1)
          Next
          BarCodes(f) = Texto
        Next
      Next
    End If
    
    'Desenho da barra
    
    retorno = retorno + "<img src=http://www.brloja.com.br/boletos/img/p.gif width=" & fino & " height=" & altura & "  border=0><img"
    retorno = retorno + " src=http://www.brloja.com.br/boletos/img/b.gif width=" & fino & "  height=" & altura & "  border=0><img"
    retorno = retorno + " src=http://www.brloja.com.br/boletos/img/p.gif width=" & fino & "  height=" & altura & "  border=0><img"
    retorno = retorno + " src=http://www.brloja.com.br/boletos/img/b.gif width=" & fino & "  height=" & altura & "  border=0><img"
    
    
    Texto = Valor
    If Len(Texto) Mod 2 <> 0 Then
      Texto = "0" & Texto
    End If
    
    ' Draw dos dados
    Do While Len(Texto) > 0
      i = CInt(Left(Texto, 2))
      Texto = Right(Texto, Len(Texto) - 2)
      f = BarCodes(i)
      For i = 1 To 10 Step 2
        If Mid(f, i, 1) = "0" Then
          f1 = fino
        Else
          f1 = largo
        End If
    
        retorno = retorno + " src=http://www.brloja.com.br/boletos/img/p.gif width=" & f1 & " height=" & altura & " border=0><img"
     
        If Mid(f, i + 1, 1) = "0" Then
          f2 = fino
        Else
          f2 = largo
        End If
    
        retorno = retorno + " src=http://www.brloja.com.br/boletos/img/b.gif width=" & f2 & " height=" & altura & " border=0><img"
    
      Next
    Loop
    
    ' Draw guarda final
    retorno = retorno + " src=http://www.brloja.com.br/boletos/img/p.gif width=" & largo & " height=" & altura & " border=0><img"
    retorno = retorno + " src=http://www.brloja.com.br/boletos/img/b.gif width=" & fino & " height=" & altura & " border=0><img"
    retorno = retorno + " src=http://www.brloja.com.br/boletos/img/p.gif width=" & f1 & " height=" & altura & " border=0>"
    
    WBarCode = retorno
    
End Function

Public Function EnviarEmail(Arquivo As String, Para As String, ParaNome As String, Assunto As String)

   Dim retorno As Boolean
   Dim Texto As String
   Dim Saida As String
   Dim a As Integer
   Dim TextoRun As String
   Dim Para1 As String
   Dim Para2 As String
   
   
   Saida = Application.CurrentProject.Path + "\Saida.htm"
   If Dir(Saida) <> "" Then Kill Saida

   Open Application.CurrentProject.Path + "\" + Arquivo For Input As #1
   Open Saida For Output As #2
   Do While Not EOF(1)
      Line Input #1, Texto
      For a = 1 To 30
         If VarEmail(a, 1) = "(FIM)" Then Exit For
         If Not IsNull(VarEmail(a, 1)) Then
            If IsNull(VarEmail(a, 2)) Then VarEmail(a, 2) = ""
            Texto = Replace2(Texto, VarEmail(a, 1), VarEmail(a, 2))
         End If
      Next a
      Print #2, Texto
      DoEvents
   Loop
   Close #1
   Close #2

   retorno = False
   If Dir(Application.CurrentProject.Path + "\ok.txt") <> "" Then Kill Application.CurrentProject.Path + "\ok.txt"
   If Dir(Application.CurrentProject.Path + "\error.txt") <> "" Then Kill Application.CurrentProject.Path + "\error.txt"
     
   ChDir Application.CurrentProject.Path
   
   'Para = "cleber@brloja.com.br"
   Para1 = Left(Para, Int(Len(Para) / 2))
   Para2 = Mid(Para, Int(Len(Para) / 2))

   TextoRun = "Roda.bat " & Para & " " & Chr(34) & Assunto & Chr(34) & " " & Chr(34) & ParaNome & Chr(34) & Chr(34)

   Shell TextoRun, vbMinimizedNoFocus
   
   Do While True
      If Dir(Application.CurrentProject.Path + "\ok.txt") <> "" Then
         retorno = True
         Exit Do
      ElseIf Dir(Application.CurrentProject.Path + "\error.txt") <> "" Then
         retorno = False
         Exit Do
      End If
      DoEvents
   Loop
   If retorno = False Then MsgBox "ATENÇÃO: Erro no envio da mensagem!"
   EnviarEmail = retorno
   
End Function

Public Function Replace2(Txt, Busca, Troca)

Dim a As Integer
Dim aux As String

For a = 1 To Len(Txt)
   If Mid(Txt, a, Len(Busca)) = Busca Then
      aux = aux & Troca
      a = a + Len(Busca) - 1
   Else
      aux = aux & Mid(Txt, a, 1)
   End If
Next a

Replace2 = aux

End Function



Public Function MandarEmail()

Dim Faturamentos As DAO.Recordset
Dim Dados As DAO.Recordset
Dim Clientes As DAO.Recordset
Dim codBarras As String


Set Faturamentos = CurrentDb.OpenRecordset("Select * from Faturamentos Where status = 2 and mandaremail = -1")
Set Dados = CurrentDb.OpenRecordset("Select * from Dados")
Set Clientes = CurrentDb.OpenRecordset("Select * from Cadastros")

While Not Faturamentos.EOF

    If Faturamentos.Fields("Status") = 2 And IsNull(Faturamentos.Fields("EmailCobranca")) Then
       
       Clientes.FindFirst "codCadastro = " & Faturamentos.Fields("codCadastro")
       
       If Not Clientes.NoMatch Then
       
       VarEmail(1, 1) = "[cedente]"
       VarEmail(1, 2) = Dados.Fields("favorecido")
       
       VarEmail(2, 1) = "[agencia]"
       VarEmail(2, 2) = Dados.Fields("agencia")
       
       VarEmail(3, 1) = "[conta]"
       VarEmail(3, 2) = Dados.Fields("conta") & "-" & Dados.Fields("conta_dac")
       
       VarEmail(4, 1) = "[data_doc]"
       VarEmail(4, 2) = Date
       
       VarEmail(5, 1) = "[vencimento]"
       VarEmail(5, 2) = Faturamentos.Fields("DatadeVencimento")
       
       VarEmail(6, 1) = "[sacado]"
       VarEmail(6, 2) = Clientes.Fields("Razao")
       
       VarEmail(7, 1) = "[numero_doc]"
       VarEmail(7, 2) = FormatNumber(Faturamentos.Fields("codCadastro"), 0) & Format(Faturamentos.Fields("DatadeVencimento"), "ddmmyy")
       
       VarEmail(8, 1) = "[nosso_num]"
       VarEmail(8, 2) = Dados.Fields("carteira") & "/" & VarEmail(7, 2) & "-" & Calculo_DV10(Dados.Fields("agencia") & Dados.Fields("conta") & Dados.Fields("carteira") & VarEmail(7, 2))
    
       VarEmail(9, 1) = "[valor_doc]"
       VarEmail(9, 2) = FormatNumber(Faturamentos.Fields("ValorDoFaturamento"), 2)
              
       codBarras = Monta_CodBarras(Faturamentos.Fields("ValorDoFaturamento"), Faturamentos.Fields("DatadeVencimento"), Dados.Fields("cdmoeda"), Dados.Fields("banco"), Dados.Fields("agencia"), Dados.Fields("conta"), Dados.Fields("conta_dac"), Dados.Fields("carteira"), VarEmail(7, 2), Calculo_DV10(Dados.Fields("agencia") & Dados.Fields("conta") & Dados.Fields("carteira") & VarEmail(7, 2)))
       VarEmail(10, 1) = "[codigo_boleto]"
       VarEmail(10, 2) = Linha_Digitavel(codBarras)
       
       VarEmail(11, 1) = "[local1]"
       VarEmail(11, 2) = Dados.Fields("local1")
       
       VarEmail(12, 1) = "[local2]"
       VarEmail(12, 2) = Dados.Fields("local2")
       
       VarEmail(13, 1) = "[especie]"
       VarEmail(13, 2) = Dados.Fields("especie")
       
       VarEmail(14, 1) = "[aceite]"
       VarEmail(14, 2) = Dados.Fields("aceite")
       
       VarEmail(15, 1) = "[moeda]"
       VarEmail(15, 2) = Dados.Fields("moeda")
       
       VarEmail(16, 1) = "[instr1]"
       VarEmail(16, 2) = Dados.Fields("instrucao1")
       
       VarEmail(17, 1) = "[instr2]"
       VarEmail(17, 2) = Dados.Fields("instrucao2")
       
       VarEmail(18, 1) = "[instr3]"
       VarEmail(18, 2) = Dados.Fields("instrucao3")
       
       VarEmail(19, 1) = "[instr4]"
       VarEmail(19, 2) = Dados.Fields("instrucao4")
       
       VarEmail(20, 1) = "[sacado_endereco]"
       VarEmail(20, 2) = Clientes.Fields("endereco")
       
       VarEmail(21, 1) = "[sacado_documento]"
       VarEmail(21, 2) = Clientes.Fields("CNPJ_CPF")
       
       VarEmail(22, 1) = "[codigo_barra]"
       VarEmail(22, 2) = WBarCode(codBarras)
       
       VarEmail(23, 1) = "[sacado_cep]"
       VarEmail(23, 2) = Clientes.Fields("cep")
       
       VarEmail(24, 1) = "[sacado_cidade]"
       VarEmail(24, 2) = Clientes.Fields("municipio")
       
       VarEmail(25, 1) = "[sacado_uf]"
       VarEmail(25, 2) = Clientes.Fields("estado")
       
       VarEmail(26, 1) = "[carteira]"
       VarEmail(26, 2) = Dados.Fields("carteira")
       
       VarEmail(27, 1) = "(FIM)"
       
       If EnviarEmail("Boleto.htm", Clientes.Fields("eMailCob"), Clientes.Fields("Nome"), "Cobrança de mensalidade") Then
          JogaValor "EmailCobranca", Date, Faturamentos
       End If
       
       End If
       
    ElseIf Faturamentos.Fields("Status") = 2 And IsNull(Faturamentos.Fields("EmailCancelar")) And Date >= (Faturamentos.Fields("DatadeVencimento") + Dados.Fields("DiasParaAtrazoDeRC") * 4) Then
       
       Clientes.FindFirst "codCadastro = " & Faturamentos.Fields("codCadastro")
       
       If Not Clientes.NoMatch Then
       
       VarEmail(1, 1) = "[vencimento]"
       VarEmail(1, 2) = Faturamentos.Fields("DatadeVencimento")
       
       VarEmail(2, 1) = "[sacado]"
       VarEmail(2, 2) = Clientes.Fields("Razao")
           
       VarEmail(3, 1) = "[valor_doc]"
       VarEmail(3, 2) = FormatNumber(Faturamentos.Fields("ValorDoFaturamento"), 2)
                     
       VarEmail(4, 1) = "[sacado_documento]"
       VarEmail(4, 2) = Clientes.Fields("CNPJ_CPF")
       
       VarEmail(5, 1) = "[servico]"
       VarEmail(5, 2) = Faturamentos.Fields("descricaodofaturamento")
       
       VarEmail(6, 1) = "[aviso]"
       VarEmail(6, 2) = "Aviso de cancelamento. (4o. Aviso)"
       
       VarEmail(7, 1) = "(FIM)"
       
       If EnviarEmail("Cancelar.htm", Clientes.Fields("eMailCob"), Clientes.Fields("Nome"), "Aviso de cancelamento. (4o Aviso)") Then
          JogaValor "EmailCancelar", Date, Faturamentos
       End If
       
       End If
       
    ElseIf Faturamentos.Fields("Status") = 2 And IsNull(Faturamentos.Fields("EmailAtraso3")) And Date >= (Faturamentos.Fields("DatadeVencimento") + Dados.Fields("DiasParaAtrazoDeRC") * 3) Then
       
       Clientes.FindFirst "codCadastro = " & Faturamentos.Fields("codCadastro")
       
       If Not Clientes.NoMatch Then
       
       VarEmail(1, 1) = "[vencimento]"
       VarEmail(1, 2) = Faturamentos.Fields("DatadeVencimento")
       
       VarEmail(2, 1) = "[sacado]"
       VarEmail(2, 2) = Clientes.Fields("Razao")
           
       VarEmail(3, 1) = "[valor_doc]"
       VarEmail(3, 2) = FormatNumber(Faturamentos.Fields("ValorDoFaturamento"), 2)
                     
       VarEmail(4, 1) = "[sacado_documento]"
       VarEmail(4, 2) = Clientes.Fields("CNPJ_CPF")
       
       VarEmail(5, 1) = "[servico]"
       VarEmail(5, 2) = Faturamentos.Fields("descricaodofaturamento")
       
       VarEmail(6, 1) = "[aviso]"
       VarEmail(6, 2) = "Pagamento em atraso (3o. Aviso)"
       
       VarEmail(7, 1) = "(FIM)"
       If EnviarEmail("Atraso.htm", Clientes.Fields("eMailCob"), Clientes.Fields("Nome"), "Pagamento em atraso (3o Aviso)") Then
          JogaValor "EmailAtraso3", Date, Faturamentos
       End If
       
       End If
       
    ElseIf Faturamentos.Fields("Status") = 2 And IsNull(Faturamentos.Fields("EmailAtraso2")) And Date >= (Faturamentos.Fields("DatadeVencimento") + Dados.Fields("DiasParaAtrazoDeRC") * 2) Then
       
       Clientes.FindFirst "codCadastro = " & Faturamentos.Fields("codCadastro")
       
       If Not Clientes.NoMatch Then
       
       VarEmail(1, 1) = "[vencimento]"
       VarEmail(1, 2) = Faturamentos.Fields("DatadeVencimento")
       
       VarEmail(2, 1) = "[sacado]"
       VarEmail(2, 2) = Clientes.Fields("Razao")
           
       VarEmail(3, 1) = "[valor_doc]"
       VarEmail(3, 2) = FormatNumber(Faturamentos.Fields("ValorDoFaturamento"), 2)
                     
       VarEmail(4, 1) = "[sacado_documento]"
       VarEmail(4, 2) = Clientes.Fields("CNPJ_CPF")
       
       VarEmail(5, 1) = "[servico]"
       VarEmail(5, 2) = Faturamentos.Fields("descricaodofaturamento")
       
       VarEmail(6, 1) = "[aviso]"
       VarEmail(6, 2) = "Pagamento em atraso (2o. Aviso)"
       
       VarEmail(7, 1) = "(FIM)"
       
       If EnviarEmail("Atraso.htm", Clientes.Fields("eMailCob"), Clientes.Fields("Nome"), "Pagamento em atraso (2o Aviso)") Then
          JogaValor "EmailAtraso2", Date, Faturamentos
       End If
       
       End If
       
    ElseIf Faturamentos.Fields("Status") = 2 And IsNull(Faturamentos.Fields("EmailAtraso1")) And Date >= (Faturamentos.Fields("DatadeVencimento") + Dados.Fields("DiasParaAtrazoDeRC")) Then
       
       Clientes.FindFirst "codCadastro = " & Faturamentos.Fields("codCadastro")
       
       If Not Clientes.NoMatch Then
       
       VarEmail(1, 1) = "[vencimento]"
       VarEmail(1, 2) = Faturamentos.Fields("DatadeVencimento")
       
       VarEmail(2, 1) = "[sacado]"
       VarEmail(2, 2) = Clientes.Fields("Razao")
           
       VarEmail(3, 1) = "[valor_doc]"
       VarEmail(3, 2) = FormatNumber(Faturamentos.Fields("ValorDoFaturamento"), 2)
                     
       VarEmail(4, 1) = "[sacado_documento]"
       VarEmail(4, 2) = Clientes.Fields("CNPJ_CPF")
       
       VarEmail(5, 1) = "[servico]"
       VarEmail(5, 2) = Faturamentos.Fields("descricaodofaturamento")
       
       VarEmail(6, 1) = "[aviso]"
       VarEmail(6, 2) = "Pagamento em atraso (1o. Aviso)"
       
       VarEmail(7, 1) = "(FIM)"
       
       If EnviarEmail("Atraso.htm", Clientes.Fields("eMailCob"), Clientes.Fields("Nome"), "Pagamento em atraso (1o Aviso)") Then
          JogaValor "EmailAtraso1", Date, Faturamentos
       End If
       
       End If
       
    End If
    Faturamentos.MoveNext

Wend

Faturamentos.Close
Dados.Close
Clientes.Close

End Function


Public Function ImportarContas()

If MsgBox("A data do computador esta correta? " & Date, vbYesNo, "Importação") = vbYes Then
   ImportarMovimentosFixos Month(Date), Year(Date)
   ImportarFaturamentosFixos Month(Date), Year(Date)
   ImportarMovimentosPessoaisFixos Month(Date), Year(Date)
   ImportarFaturamentosPessoaisFixos Month(Date), Year(Date)
   'MandarEmail
End If

End Function



Public Function ImportarMovimentosFixos(Optional MES As Integer, Optional ANO As Integer)

Dim Movimentos As DAO.Recordset
Dim MovimentosFixos As DAO.Recordset
Dim Dados As DAO.Recordset

Set Movimentos = CurrentDb.OpenRecordset("Select * from Movimentos")
Set MovimentosFixos = CurrentDb.OpenRecordset("Select * from MovimentosFixos where MovimentoAtivo = true")
Set Dados = CurrentDb.OpenRecordset("Select * from Dados")

BeginTrans

While Not MovimentosFixos.EOF

If IIf(MovimentosFixos.Fields("mesinicio") > 0 And MovimentosFixos.Fields("mesfinal") > 0, MovimentosFixos.Fields("mesinicio") <= Month(Date) And MovimentosFixos.Fields("mesfinal") >= Month(Date), False) Or _
   IIf(MovimentosFixos.Fields("mesinicio") > 0 And MovimentosFixos.Fields("mesfinal") = 0, MovimentosFixos.Fields("mesinicio") <= Month(Date), False) Or _
   IIf(MovimentosFixos.Fields("mesinicio") = 0 And MovimentosFixos.Fields("mesfinal") > 0, MovimentosFixos.Fields("mesfinal") >= Month(Date), False) Or _
   IIf(MovimentosFixos.Fields("mesinicio") = 0 And MovimentosFixos.Fields("mesfinal") = 0, True, False) Then

    If Compara(MovimentosFixos.Fields("MesGerado"), MovimentosFixos.Fields("AnoGerado"), MES, ANO) And Val(MovimentosFixos.Fields("DiaDeVencimento")) <= Day(Date) + Dados.Fields("DiasParaGerarPG") Then
    
       Movimentos.AddNew
       Movimentos.Fields("codMovimento") = NovoCodigo("Movimentos", "codMovimento")
       Movimentos.Fields("DescricaoDoMovimento") = MovimentosFixos.Fields("Descricao")
       Movimentos.Fields("codCadastro") = MovimentosFixos.Fields("codCadastro")
       Movimentos.Fields("ValorDoMovimento") = MovimentosFixos.Fields("ValorDoMovimento")
       Movimentos.Fields("DataDeVencimento") = CalcularVencimento(MovimentosFixos.Fields("DiaDeVencimento"), MES, ANO)
       Movimentos.Fields("Status") = 2
       Movimentos.Fields("Obs") = MovimentosFixos.Fields("Obs")
       Movimentos.Fields("codCategoriadePagamento") = MovimentosFixos.Fields("codCategoriadePagamento")
       
       Movimentos.Update
       Movimentos.MoveLast
       
       MovimentosFixos.Edit
       MovimentosFixos.Fields("MesGerado") = MES
       MovimentosFixos.Fields("AnoGerado") = ANO
       MovimentosFixos.Update
       
    End If
    
End If
MovimentosFixos.MoveNext
DoEvents

Wend

CommitTrans

Movimentos.Close
MovimentosFixos.Close
Dados.Close

End Function


Public Function ImportarFaturamentosFixos(Optional MES As Integer, Optional ANO As Integer)

Dim Faturamentos As DAO.Recordset
Dim FaturamentosFixos As DAO.Recordset
Dim Dados As DAO.Recordset

Set Faturamentos = CurrentDb.OpenRecordset("Select * from Faturamentos")
Set FaturamentosFixos = CurrentDb.OpenRecordset("Select * from FaturamentosFixos Where FaturamentoAtivo = True")
Set Dados = CurrentDb.OpenRecordset("Select * from Dados")

BeginTrans

While Not FaturamentosFixos.EOF

If IIf(FaturamentosFixos.Fields("mesinicio") > 0 And FaturamentosFixos.Fields("mesfinal") > 0, FaturamentosFixos.Fields("mesinicio") <= Month(Date) And FaturamentosFixos.Fields("mesfinal") >= Month(Date), False) Or _
   IIf(FaturamentosFixos.Fields("mesinicio") > 0 And FaturamentosFixos.Fields("mesfinal") = 0, FaturamentosFixos.Fields("mesinicio") <= Month(Date), False) Or _
   IIf(FaturamentosFixos.Fields("mesinicio") = 0 And FaturamentosFixos.Fields("mesfinal") > 0, FaturamentosFixos.Fields("mesfinal") >= Month(Date), False) Or _
   IIf(FaturamentosFixos.Fields("mesinicio") = 0 And FaturamentosFixos.Fields("mesfinal") = 0, True, False) Then

    If Compara(FaturamentosFixos.Fields("MesGerado"), FaturamentosFixos.Fields("AnoGerado"), MES, ANO) And Val(FaturamentosFixos.Fields("DiaDeVencimento")) <= (Day(Date) + Dados.Fields("DiasParaGerarRC")) Then

       Faturamentos.AddNew
       Faturamentos.Fields("codFaturamento") = NovoCodigo("Faturamentos", "codFaturamento")
       Faturamentos.Fields("codCadastro") = FaturamentosFixos.Fields("codCadastro")
       Faturamentos.Fields("MandarEmail") = FaturamentosFixos.Fields("MandarEmail")
       Faturamentos.Fields("DescricaoDoFaturamento") = FaturamentosFixos.Fields("Descricao")
       Faturamentos.Fields("ValorDoFaturamento") = FaturamentosFixos.Fields("ValorDoFaturamento")
       Faturamentos.Fields("DataDeVencimento") = CalcularVencimento(FaturamentosFixos.Fields("DiaDeVencimento"), MES, ANO)
       Faturamentos.Fields("Status") = 2
       Faturamentos.Fields("Obs") = FaturamentosFixos.Fields("Obs")
       Faturamentos.Fields("codCategoriadePagamento") = FaturamentosFixos.Fields("codCategoriadePagamento")

       Faturamentos.Update
       Faturamentos.MoveLast
    
       FaturamentosFixos.Edit
       FaturamentosFixos.Fields("MesGerado") = MES
       FaturamentosFixos.Fields("AnoGerado") = ANO
       FaturamentosFixos.Update
    End If
    
End If
DoEvents
FaturamentosFixos.MoveNext

Wend

CommitTrans

Faturamentos.Close
FaturamentosFixos.Close
Dados.Close

End Function


Public Function ImportarMovimentosPessoaisFixos(Optional MES As Integer, Optional ANO As Integer)

Dim Movimentos As DAO.Recordset
Dim MovimentosFixos As DAO.Recordset
Dim Dados As DAO.Recordset

Set Movimentos = CurrentDb.OpenRecordset("Select * from MovimentosPessoais")
Set MovimentosFixos = CurrentDb.OpenRecordset("Select * from MovimentosPessoaisFixos where MovimentoAtivo = true")
Set Dados = CurrentDb.OpenRecordset("Select * from Dados")

BeginTrans

While Not MovimentosFixos.EOF

If IIf(MovimentosFixos.Fields("mesinicio") > 0 And MovimentosFixos.Fields("mesfinal") > 0, MovimentosFixos.Fields("mesinicio") <= Month(Date) And MovimentosFixos.Fields("mesfinal") >= Month(Date), False) Or _
   IIf(MovimentosFixos.Fields("mesinicio") > 0 And MovimentosFixos.Fields("mesfinal") = 0, MovimentosFixos.Fields("mesinicio") <= Month(Date), False) Or _
   IIf(MovimentosFixos.Fields("mesinicio") = 0 And MovimentosFixos.Fields("mesfinal") > 0, MovimentosFixos.Fields("mesfinal") >= Month(Date), False) Or _
   IIf(MovimentosFixos.Fields("mesinicio") = 0 And MovimentosFixos.Fields("mesfinal") = 0, True, False) Then

    If Compara(MovimentosFixos.Fields("MesGerado"), MovimentosFixos.Fields("AnoGerado"), MES, ANO) And Val(MovimentosFixos.Fields("DiaDeVencimento")) <= Day(Date) + Dados.Fields("DiasParaGerarPG") Then
    
       Movimentos.AddNew
       Movimentos.Fields("codMovimento") = NovoCodigo("MovimentosPessoais", "codMovimento")
       Movimentos.Fields("DescricaoDoMovimento") = MovimentosFixos.Fields("Descricao")
       Movimentos.Fields("codCategoriadeMovimento") = MovimentosFixos.Fields("codCategoriadeMovimento")
       Movimentos.Fields("ValorDoMovimento") = MovimentosFixos.Fields("ValorDoMovimento")
       Movimentos.Fields("DataDeVencimento") = CalcularVencimento(MovimentosFixos.Fields("DiaDeVencimento"), MES, ANO)
       Movimentos.Fields("Status") = 2
       Movimentos.Fields("Obs") = MovimentosFixos.Fields("Obs")
       Movimentos.Fields("codCategoriadePagamento") = MovimentosFixos.Fields("codCategoriadePagamento")
       
       Movimentos.Update
       Movimentos.MoveLast
       
       MovimentosFixos.Edit
       MovimentosFixos.Fields("MesGerado") = MES
       MovimentosFixos.Fields("AnoGerado") = ANO
       MovimentosFixos.Update
       
    End If
    
End If
MovimentosFixos.MoveNext
DoEvents

Wend

CommitTrans

Movimentos.Close
MovimentosFixos.Close
Dados.Close

End Function


Public Function ImportarFaturamentosPessoaisFixos(Optional MES As Integer, Optional ANO As Integer)

Dim Faturamentos As DAO.Recordset
Dim FaturamentosFixos As DAO.Recordset
Dim Dados As DAO.Recordset

Set Faturamentos = CurrentDb.OpenRecordset("Select * from FaturamentosPessoais")
Set FaturamentosFixos = CurrentDb.OpenRecordset("Select * from FaturamentosPessoaisFixos Where FaturamentoAtivo = True")
Set Dados = CurrentDb.OpenRecordset("Select * from Dados")

BeginTrans

While Not FaturamentosFixos.EOF

If IIf(FaturamentosFixos.Fields("mesinicio") > 0 And FaturamentosFixos.Fields("mesfinal") > 0, FaturamentosFixos.Fields("mesinicio") <= Month(Date) And FaturamentosFixos.Fields("mesfinal") >= Month(Date), False) Or _
   IIf(FaturamentosFixos.Fields("mesinicio") > 0 And FaturamentosFixos.Fields("mesfinal") = 0, FaturamentosFixos.Fields("mesinicio") <= Month(Date), False) Or _
   IIf(FaturamentosFixos.Fields("mesinicio") = 0 And FaturamentosFixos.Fields("mesfinal") > 0, FaturamentosFixos.Fields("mesfinal") >= Month(Date), False) Or _
   IIf(FaturamentosFixos.Fields("mesinicio") = 0 And FaturamentosFixos.Fields("mesfinal") = 0, True, False) Then

    If Compara(FaturamentosFixos.Fields("MesGerado"), FaturamentosFixos.Fields("AnoGerado"), MES, ANO) And Val(FaturamentosFixos.Fields("DiaDeVencimento")) <= (Day(Date) + Dados.Fields("DiasParaGerarRC")) Then

       Faturamentos.AddNew
       Faturamentos.Fields("codFaturamento") = NovoCodigo("Faturamentos", "codFaturamento")
       Faturamentos.Fields("DescricaoDoFaturamento") = FaturamentosFixos.Fields("Descricao")
       Faturamentos.Fields("ValorDoFaturamento") = FaturamentosFixos.Fields("ValorDoFaturamento")
       Faturamentos.Fields("DataDeVencimento") = CalcularVencimento(FaturamentosFixos.Fields("DiaDeVencimento"), MES, ANO)
       Faturamentos.Fields("Status") = 2
       Faturamentos.Fields("Obs") = FaturamentosFixos.Fields("Obs")
       Faturamentos.Fields("codCategoriadePagamento") = FaturamentosFixos.Fields("codCategoriadePagamento")

       Faturamentos.Update
       Faturamentos.MoveLast
    
       FaturamentosFixos.Edit
       FaturamentosFixos.Fields("MesGerado") = MES
       FaturamentosFixos.Fields("AnoGerado") = ANO
       FaturamentosFixos.Update
    End If
    
End If

FaturamentosFixos.MoveNext
DoEvents

Wend

CommitTrans

Faturamentos.Close
FaturamentosFixos.Close
Dados.Close

End Function

Public Function CalcularVencimento(DIA As Integer, Optional MES As Integer, Optional ANO As Integer) As Date

If MES > 0 And ANO > 0 Then
    CalcularVencimento = Format((DateSerial(ANO, MES, DIA)), "dd/mm/yyyy")
ElseIf MES = 0 And ANO > 0 Then
    CalcularVencimento = Format((DateSerial(ANO, Month(Now), DIA)), "dd/mm/yyyy")
ElseIf MES = 0 And ANO = 0 Then
    CalcularVencimento = Format((DateSerial(Year(Now), Month(Now), DIA)), "dd/mm/yyyy")
End If

End Function

Function JogaValor(Campo As String, Valor As String, Tabela As DAO.Recordset)

Tabela.Edit
Tabela.Fields(Campo) = Valor
Tabela.Update

End Function


Public Function Chancelamento(Inicio As Integer, Final As Integer) As String

Dim ch_X As Boolean
Dim Texto As String
Dim a As Integer

ch_X = True

For a = Inicio To Final
    Texto = Texto + IIf(ch_X, "x", "-")
    ch_X = Not ch_X
Next

Chancelamento = Texto


End Function


