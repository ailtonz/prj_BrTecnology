VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Duplicatas_Historico"
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

'Dados do Historico
Dim rHistorico As DAO.Recordset

'Posição da linha
Dim Linha As Integer


Set dbBase = CurrentDb
Set rHistorico = dbBase.OpenRecordset("Select * from Duplicatas_Historico where codDuplicata = " & Form_Duplicatas.codDuplicata)


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
'   HISTORICO
'#######################

Linha = 70

rHistorico.MoveFirst

While Not rHistorico.EOF
    Prt Linha, 4, Format(rHistorico.Fields("Emissao"), "dd/mm/yy")
    Prt Linha, 20, rHistorico.Fields("Fatura")
    Prt Linha, 38, FormatCurrency(Format(rHistorico.Fields("Valor"), "##,##0.00"))
    Linha = Linha + 3
    rHistorico.MoveNext
Wend



Exit_Detalhe_Format:
    Exit Sub

Err_Detalhe_Format:
    MsgBox Err.Description
    Resume Exit_Detalhe_Format

End Sub

Sub Prt(Linha, Coluna, Texto)

Dim mSup As Long
Dim mEsq As Long

mEsq = 0
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
