VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Duplicatas"
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

'Dados do titulo
Dim rDuplicata As DAO.Recordset

'Dados do Cliente
Dim rCliente As DAO.Recordset

Set dbBase = CurrentDb
Set rDuplicata = dbBase.OpenRecordset("Select * from Duplicatas where codDuplicata = " & Me.codDuplicata)
Set rCliente = dbBase.OpenRecordset("Select * from Cadastros where codCadastro = " & Me.codCadastro)

'##########################
'### FORMATAÇÃO DA DPL ####
'##########################

With Me
    .ScaleMode = 6
    .FontName = "Darf"
    .FontBold = False
    .FontSize = 9
End With

Printer.TopMargin = 0
Printer.BottomMargin = 0
Printer.LeftMargin = 0
Printer.RightMargin = 0

'#######################
'   DADOS DA DUPLICATA
'#######################

Prt 55, 135, rDuplicata.Fields("Emissao")

Prt 70, 47, rDuplicata.Fields("Fatura")
Prt 70, 70, FormatCurrency(Format(rDuplicata.Fields("Valor"), "##,##0.00"))
Prt 70, 105, rDuplicata.Fields("Ordem")
Prt 70, 132, IIf(Not IsNull(rDuplicata.Fields("Vencimento")), Format(rDuplicata.Fields("Vencimento"), "dd/mm/yyyy"), "")

Prt 76, 71, rDuplicata.Fields("DescontoPorcento")
Prt 76, 95, rDuplicata.Fields("DescontoValor")
Prt 76, 130, rDuplicata.Fields("DescontoAte")

Prt 85, 71, rDuplicata.Fields("CondEspeciais")

'Dados cadastrais
Prt 94, 67, rCliente.Fields("Nome")
Prt 99, 67, rCliente.Fields("Endereco") & " - " & rCliente.Fields("Bairro")
Prt 104, 67, rCliente.Fields("Municipio")
Prt 104, 120, rCliente.Fields("Estado")
Prt 104, 160, rCliente.Fields("CEP")

'Praca de pagamento
Prt 108, 67, rDuplicata.Fields("Praca")

'Documentação
Prt 112, 62, rCliente.Fields("CNPJ_CPF") 'C.N.P.J
Prt 112, 135, rCliente.Fields("IE_RG")  'I.E

'#######################
'   VALOR POR EXTENSO
'#######################

Dim extrenso01 As String
Dim extrenso02 As String
Dim Tamanho As Integer
Dim LimiteDoCampo As Integer
Dim ContinuacaoDoCampo As Integer

LimiteDoCampo = 80
ContinuacaoDoCampo = 81

Tamanho = Len(UCase(Extenso(rDuplicata.Fields("Valor"))))

If Tamanho > LimiteDoCampo Then

   extrenso01 = Mid(UCase(Extenso(rDuplicata.Fields("Valor"))), 1, LimiteDoCampo)
   extrenso02 = Mid(UCase(Extenso(rDuplicata.Fields("Valor"))), ContinuacaoDoCampo, Tamanho)

   Prt 117, 61, UCase(Extenso(rDuplicata.Fields("Valor"))) & " " & Chancelamento(Len(extrenso02), LimiteDoCampo)
   Prt 124, 61, extrenso02 & " " & Chancelamento(Len(extrenso02), LimiteDoCampo)

Else

   Prt 117, 61, UCase(Extenso(rDuplicata.Fields("Valor"))) & " " & Chancelamento(Len(UCase(Extenso(rDuplicata.Fields("Valor")))), LimiteDoCampo)
   Prt 124, 61, Chancelamento(1, LimiteDoCampo)

End If




Exit_Detalhe_Format:
    Exit Sub

Err_Detalhe_Format:
    MsgBox Err.Description
    Resume Exit_Detalhe_Format

End Sub

Sub Prt(Linha, Coluna, Texto)

Dim mSup As Long
Dim mEsq As Long

mEsq = 15
mSup = -23

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
