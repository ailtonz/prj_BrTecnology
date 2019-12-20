VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_NotasFiscais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detalhe_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo Err_Detalhe_Format

'Banco de dados corrente
Dim dbBase As database

'Dados da Nota Fiscal
Dim rNotaFiscal As DAO.Recordset
Dim sNotaFiscal As String
Dim rNotaFiscalItens As DAO.Recordset
Dim sNotaFiscalItens As String

'Dados do Cliente
Dim rCliente As DAO.Recordset
Dim sCliente As String

'Enderecos
Dim rEndereco As DAO.Recordset
Dim sEndereco As String

'Dados do Transportador
Dim rTransportador As DAO.Recordset
Dim sTransportador As String

'Somatórias
Dim SomaPesoBruto As Currency
Dim SomaPesoLiquido As Currency

sNotaFiscal = "Select * from NotasFiscais where codNotaFiscal = " & Me.codNotaFiscal
sNotaFiscalItens = "Select * from NotasFiscaisItens where codNotaFiscal = " & Me.codNotaFiscal
sCliente = "Select * from Cadastros Where codCadastro = " & Me.codCliente


Set dbBase = CurrentDb
Set rNotaFiscal = dbBase.OpenRecordset(sNotaFiscal)
Set rNotaFiscalItens = dbBase.OpenRecordset(sNotaFiscalItens)
Set rCliente = dbBase.OpenRecordset(sCliente)

'#########################
'### FORMATAÇÃO DA NF ####
'#########################

With Me
    .ScaleMode = 6
    .FontName = "Arial"
    .FontBold = False
    .FontSize = 9
End With

'Set Printer = Application.Printers("NotasFiscais")

With Printer

    .TopMargin = 360
    .BottomMargin = 0
    .LeftMargin = 110
    .RightMargin = 0
        
End With


'################
'### OPERAÇÃO ###
'################

'Linha 03
If rNotaFiscal.Fields("codOperacao") = "Entrada" Then
   Prt 5, 157, "X"  ' Entrada
ElseIf rNotaFiscal.Fields("codOperacao") = "Saída" Then
   Prt 5, 139, "X" ' Saida
End If

'############################
'### NATUREZA DA OPERAÇÃO ###
'############################

'Linha 19
Prt 19, 13, Mid(rNotaFiscal.Fields("NaturezaDeOperacao"), 1, 20)
Prt 19, 74, rNotaFiscal.Fields("codCFOP")

'##################################
'### RAZÃO / CNPJ / DT. EMISSÃO ###
'##################################

'Linha 27
Prt 29, 13, UCase(rCliente.Fields("Razao"))
Prt 29, 133, rCliente.Fields("CNPJ_CPF")
Prt 29, 184, rNotaFiscal.Fields("DataDeEmissao")

'##########################
'### END / BAIRRO / CEP ###
'##########################

'Linha 34
Prt 36, 13, rCliente.Fields("Endereco")
Prt 36, 116, rCliente.Fields("Bairro")
Prt 36, 154, rCliente.Fields("CEP")
Prt 36, 184, rNotaFiscal.Fields("DataDeSaida")

'############################
'### MUN / FONE / UF / IE ###
'############################

'Linha 40
Prt 45, 13, rCliente.Fields("Municipio")
Prt 45, 78, rCliente.Fields("Telefone1")
Prt 45, 120, rCliente.Fields("Estado")
Prt 45, 135, rCliente.Fields("IE_RG")
Prt 45, 184, rNotaFiscal.Fields("HoraDeSaida")

'###############################
'### FATURA(S) / ENDEREÇO(S) ###
'###############################

'Linha 49 - Endereço de cobrança
Prt 57, 13, rNotaFiscal.Fields("endCobranca")

'Linha 52 - Fatura "A"
Prt 62, 140, rNotaFiscal.Fields("Fatura_A_Numero")
Prt 62, 160, rNotaFiscal.Fields("Fatura_A_Vencimento")
If Not IsNull(rNotaFiscal.Fields("Fatura_A_Valor")) Then
     Prt 62, 184, Space(9 - Len(Format(rNotaFiscal.Fields("Fatura_A_Valor"), "##,##0.00"))) & Format(rNotaFiscal.Fields("Fatura_A_Valor"), "##,##0.00")
End If


'Linha 55 - Praça de pagamento
Prt 65, 13, rNotaFiscal.Fields("prcPagamento")

'Linha 55 - Fatura "B"
Prt 65, 140, rNotaFiscal.Fields("Fatura_B_Numero")
Prt 65, 160, rNotaFiscal.Fields("Fatura_B_Vencimento")
If Not IsNull(rNotaFiscal.Fields("Fatura_B_Valor")) Then
     Prt 65, 184, Space(9 - Len(Format(rNotaFiscal.Fields("Fatura_B_Valor"), "##,##0.00"))) & Format(rNotaFiscal.Fields("Fatura_B_Valor"), "##,##0.00")
End If

'Linha 58 - Fatura "C"
Prt 68, 140, rNotaFiscal.Fields("Fatura_C_Numero")
Prt 68, 160, rNotaFiscal.Fields("Fatura_C_Vencimento")
If Not IsNull(rNotaFiscal.Fields("Fatura_C_Valor")) Then
     Prt 68, 184, Space(9 - Len(Format(rNotaFiscal.Fields("Fatura_C_Valor"), "##,##0.00"))) & Format(rNotaFiscal.Fields("Fatura_C_Valor"), "##,##0.00")
End If

'Linha 61 - Fatura "D"
Prt 71, 140, rNotaFiscal.Fields("Fatura_D_Numero")
Prt 71, 160, rNotaFiscal.Fields("Fatura_D_Vencimento")
If Not IsNull(rNotaFiscal.Fields("Fatura_D_Valor")) Then
     Prt 71, 184, Space(9 - Len(Format(rNotaFiscal.Fields("Fatura_D_Valor"), "##,##0.00"))) & Format(rNotaFiscal.Fields("Fatura_D_Valor"), "##,##0.00")
End If


'#######################
'   VALOR POR EXTENSO
'#######################

'Linhas 60 / 63

Dim extenso01 As String
Dim extenso02 As String
Dim Tamanho As Integer
Dim LimiteDoCampo As Integer
Dim ContinuacaoDoCampo As Integer

LimiteDoCampo = 62
ContinuacaoDoCampo = 63

If Not IsNull(rNotaFiscal.Fields("ValorTotalDaNota")) Then
    Tamanho = Len(UCase(Extenso(rNotaFiscal.Fields("ValorTotalDaNota"))))

    If Tamanho > LimiteDoCampo Then
    
       extenso01 = Trim(Mid(UCase(Extenso(rNotaFiscal.Fields("ValorTotalDaNota"))), 1, LimiteDoCampo))
       extenso02 = Trim(Mid(UCase(Extenso(rNotaFiscal.Fields("ValorTotalDaNota"))), ContinuacaoDoCampo, Tamanho))
    
       Prt 70, 31, UCase(extenso01) & " " & Chancelamento(Len(extenso01), LimiteDoCampo)
       Prt 74, 31, extenso02 & " " & Chancelamento(Len(extenso02), LimiteDoCampo)
    
    Else
    
       Prt 70, 31, UCase(Extenso(rNotaFiscal.Fields("ValorTotalDaNota"))) & " " & Chancelamento(Len(Extenso(rNotaFiscal.Fields("ValorTotalDaNota"))), LimiteDoCampo)
       Prt 74, 31, Chancelamento(1, LimiteDoCampo)
    
    End If

End If

'##########################
'### DADOS DOS PRODUTOS ###
'##########################

If Not rNotaFiscalItens.EOF Then

   Dim x As Integer
   Dim Linha As Integer
   Dim LimiteDeRegistros As Byte
   Dim TotalDeRegistros As Byte

   LimiteDeRegistros = 22
   Linha = 89

   rNotaFiscalItens.MoveLast
   rNotaFiscalItens.MoveFirst

   Me.FontBold = False
   Me.FontSize = 8

   TotalDeRegistros = rNotaFiscalItens.RecordCount

   If TotalDeRegistros > LimiteDeRegistros Then TotalDeRegistros = LimiteDeRegistros

   For x = 1 To TotalDeRegistros

       Prt Linha, 13, rNotaFiscalItens.Fields("codProduto")
       Prt Linha, 31, rNotaFiscalItens.Fields("DescricaoDoProduto")

       If Not IsNull(rNotaFiscalItens.Fields("Quantidade")) Then
            Prt Linha, 130, Space(5 - Len(rNotaFiscalItens.Fields("Quantidade"))) & rNotaFiscalItens.Fields("Quantidade")
       End If

       If Not IsNull(rNotaFiscalItens.Fields("ValorUnitario")) Then
            Prt Linha, 142, Space(12 - Len(Format(rNotaFiscalItens.Fields("ValorUnitario"), "##,##0.00"))) & Format(rNotaFiscalItens.Fields("ValorUnitario"), "##,##0.00")
       End If

       If Not IsNull(rNotaFiscalItens.Fields("ValorTotal")) Then
            Prt Linha, 165, Space(12 - Len(Format(rNotaFiscalItens.Fields("ValorTotal"), "##,##0.00"))) & Format(rNotaFiscalItens.Fields("ValorTotal"), "##,##0.00")
       End If

       Prt Linha, 186, IIf(Not IsNull(rNotaFiscalItens.Fields("ICMS")), rNotaFiscalItens.Fields("ICMS"), "")

       Linha = Linha + 3

       rNotaFiscalItens.MoveNext

   Next x

End If


'##############################
'### PRESTAÇÃO DE SERVIÇOS ####
'##############################

Prt 171, 13, rNotaFiscal.Fields("servDescricao01")
Prt 175, 13, rNotaFiscal.Fields("servDescricao02")
Prt 179, 13, rNotaFiscal.Fields("servDescricao03")

Prt 174, 183, IIf(Not IsNull(rNotaFiscal.Fields("servValorDoISS")), Space(12 - Len(Format(rNotaFiscal.Fields("servValorDoISS"), "##,##0.00"))) & Format(rNotaFiscal.Fields("servValorDoISS"), "##,##0.00"), "")
Prt 180, 183, IIf(Not IsNull(rNotaFiscal.Fields("servTotal")), Space(12 - Len(Format(rNotaFiscal.Fields("servTotal"), "##,##0.00"))) & Format(rNotaFiscal.Fields("servTotal"), "##,##0.00"), "")

'#########################
'### FORMATAÇÃO DA NF ####
'#########################

With Me
    .ScaleMode = 6
    .FontName = "Arial"
    .FontBold = False
    .FontSize = 9
End With

'###########################
'### CALCULOS DO IMPOSTO ###
'###########################

'Linha 197
If Not IsNull(rNotaFiscal.Fields("BaseDeCalculoDoICMS")) Then
    Prt 190, 16, Format(rNotaFiscal.Fields("BaseDeCalculoDoICMS"), "##,##0.00")
End If

If Not IsNull(rNotaFiscal.Fields("ValorDoICMS")) Then
    Prt 190, 55, Format(rNotaFiscal.Fields("ValorDoICMS"), "##,##0.00")
End If

If Not IsNull(rNotaFiscal.Fields("BaseDeCalculoICMSSubstituicao")) Then
    Prt 190, 95, Format(rNotaFiscal.Fields("BaseDeCalculoICMSSubstituicao"), "##,##0.00")
End If

If Not IsNull(rNotaFiscal.Fields("ValorDoICMSSubstituicao")) Then
    Prt 190, 135, Format(rNotaFiscal.Fields("ValorDoICMSSubstituicao"), "##,##0.00")
End If

If Not IsNull(rNotaFiscal.Fields("ValorTotalDosProdutos")) Then
    Prt 190, 175, Format(rNotaFiscal.Fields("ValorTotalDosProdutos"), "##,##0.00")
End If

'Linha 205
If Not IsNull(rNotaFiscal.Fields("ValorDoFrete")) Then
    Prt 199, 16, Format(rNotaFiscal.Fields("ValorDoFrete"), "##,##0.00")
End If

If Not IsNull(rNotaFiscal.Fields("ValorDoSeguro")) Then
    Prt 199, 55, Format(rNotaFiscal.Fields("ValorDoSeguro"), "##,##0.00")
End If

If Not IsNull(rNotaFiscal.Fields("OutrasDespesasAcessorias")) Then
    Prt 199, 95, Format(rNotaFiscal.Fields("OutrasDespesasAcessorias"), "##,##0.00")
End If

If Not IsNull(rNotaFiscal.Fields("ValorTotalDoIPI")) Then
    Prt 199, 135, Format(rNotaFiscal.Fields("ValorTotalDoIPI"), "##,##0.00")
End If

If Not IsNull(rNotaFiscal.Fields("ValorTotalDaNota")) Then
    Prt 199, 175, Format(rNotaFiscal.Fields("ValorTotalDaNota"), "##,##0.00")
End If

'#####################
'### TRANSPORTADOR ###
'#####################

'Linha 216 / 215
Prt 175, 13, UCase(rNotaFiscal.Fields("transpNome"))
Prt 175, 120, rNotaFiscal.Fields("transpFretePorConta")
Prt 175, 135, rNotaFiscal.Fields("transpPlaca")
Prt 175, 158, rNotaFiscal.Fields("transpPlaca_UF")
Prt 175, 168, rNotaFiscal.Fields("transCNPJ_CPF")

Prt 181, 13, rNotaFiscal.Fields("transpEndereco")
Prt 181, 158, rNotaFiscal.Fields("transpEndereco_UF")
Prt 181, 168, rNotaFiscal.Fields("transIE_RG")

'########################
'### DADOS ADICIONAIS ###
'########################

Prt 238, 15, rNotaFiscal.Fields("DadoAdicional_1")

Prt 242, 15, rNotaFiscal.Fields("DadoAdicional_2")

Prt 246, 15, rNotaFiscal.Fields("DadoAdicional_3")

Prt 250, 15, rNotaFiscal.Fields("DadoAdicional_4")

Exit_Detalhe_Format:
    Exit Sub

Err_Detalhe_Format:
    MsgBox Err.Description
    Resume Exit_Detalhe_Format
    
End Sub

Sub Prt(Linha, Coluna, Texto)

Dim mSup As Long
Dim mEsq As Long

'If Me.Page > 1 Then
'   If Me.Page = 2 Then
'      mSup = 18#
'   Else
'      mSup = 17# - ((Me.Page - 1) * 0.5)
'   End If
'Else
'   mSup = 14#
'End If

mEsq = 0
'' Okidata 520
' mSup = 12 - 1

'' LX300
'mSup = 10 - 1

'' EPSON LX Series 1 (80)
 mSup = 17 - 1

With Me
  .CurrentY = Linha + mSup
  .CurrentX = Coluna + mEsq
  If Not IsNull(Texto) Then
     .Print Texto
  End If
End With

End Sub

Private Sub Report_Open(Cancel As Integer)
    DoCmd.Maximize
End Sub
