VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Inicializar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public ValorAnterior As Currency
Public Sinal As String




Private Sub cmdAnterior_Click()
varQualMes = varQualMes - 1
If varQualMes = 0 Then
   varQualMes = 12
   varQualAno = varQualAno - 1
End If
If varQualMes < Month(Date) Or varQualAno < Year(Date) Then
   varQualData = IIf(varQualMes = 2, "28", "30") & "/" & varQualMes & "/" & varQualAno
Else
   varQualData = Date
End If
AtualizaData
End Sub

Private Sub cmdProximo_Click()
varQualMes = varQualMes + 1
If varQualMes = 13 Then
   varQualMes = 1
   varQualAno = varQualAno + 1
End If
If varQualMes < Month(Date) Or varQualAno < Year(Date) Then
   varQualData = IIf(varQualMes = 2, "28", "30") & "/" & varQualMes & "/" & varQualAno
Else
   varQualData = Date
End If
AtualizaData
End Sub

Private Sub Form_Open(Cancel As Integer)
'DoCmd.Maximize

ValorAnterior = 0
Sinal = ""
Me.Move 0, 0
ImportarContas

varQualMes = Month(Date)
varQualAno = Year(Date)
varQualData = Date

AtualizaData


End Sub


Sub AtualizaData()
Dim varSoma As Currency

lblData.Caption = IIf(varQualMes = Month(Date) And varQualAno = Year(Date), "** (", "") & MonthName(varQualMes) & " de " & varQualAno & IIf(varQualMes = Month(Date) And varQualAno = Year(Date), ") **", "")
varSoma = SomarCampo("qryMovimentosPessoaisHoje", "Valor", "Total") + SomarCampo("qryMovimentosHoje", "Valor", "Total") + SomarCampo("qryFaturamentosPessoaisHoje", "Valor", "Total") + SomarCampo("qryFaturamentosHoje", "Valor", "Total")
If varSoma = 0 Then
   Guia.Pages(1).Visible = False
   Guia.Pages(1).Caption = "Sem pagamentos para hoje"
Else
   Guia.Pages(1).Visible = True
   Guia.Pages(1).Caption = "Existe(m) pagamento(s) para hoje: " & FormatCurrency(varSoma, 2)
End If

Me.Refresh

End Sub




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Calculadora
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Valores_KeyDown(KeyCode As Integer, Shift As Integer)

If ((KeyCode >= 48 And KeyCode <= 57) Or KeyCode = 188) And Shift = 0 Then

ElseIf KeyCode = 187 And Shift = 1 Then '+
If ValorAnterior = 0 Then
   ValorAnterior = Valores.Text
Else
   ValorAnterior = ValorAnterior + Valores.Text
End If
KeyCode = 0
Valores.Text = ValorAnterior
Sinal = "+"

Valores.SetFocus
Valores.SelStart = 0
Valores.SelLength = Len(ValorAnterior)


ElseIf KeyCode = 189 And Shift = 0 Then '-
If ValorAnterior = 0 Then
   ValorAnterior = Valores.Text
Else
   ValorAnterior = ValorAnterior - Valores.Text
End If
KeyCode = 0
Valores.Text = ValorAnterior
Sinal = "-"

Valores.SetFocus
Valores.SelStart = 0
Valores.SelLength = Len(ValorAnterior)

ElseIf KeyCode = 191 And Shift = 0 Then '/
If ValorAnterior = 0 Then
   ValorAnterior = Valores.Text
Else
   ValorAnterior = Valores.Text / ValorAnterior
End If
KeyCode = 0
Valores.Text = ValorAnterior
Sinal = "/"

Valores.SetFocus
Valores.SelStart = 0
Valores.SelLength = Len(ValorAnterior)

ElseIf KeyCode = 88 Or (KeyCode = 56 And Shift = 1) Then 'x
If ValorAnterior = 0 Then
   ValorAnterior = Valores.Text
Else
   ValorAnterior = ValorAnterior * Valores.Text
End If
KeyCode = 0
Valores.Text = ""
Sinal = "*"

Valores.SetFocus
Valores.SelStart = 0
Valores.SelLength = Len(ValorAnterior)

ElseIf KeyCode = 67 Or KeyCode = 27 Then 'c
KeyCode = 0
ValorAnterior = 0
Valores.Text = ValorAnterior
Sinal = ""

Valores.SetFocus
Valores.SelStart = 0
Valores.SelLength = Len(ValorAnterior)

ElseIf (KeyCode = 187 And Shift = 0) Or KeyCode = 13 Then '=
If Sinal <> "" Then
   If Sinal = "+" Then
      ValorAnterior = ValorAnterior + Valores.Text
   ElseIf Sinal = "-" Then
      ValorAnterior = ValorAnterior - Valores.Text
   ElseIf Sinal = "/" Then
      ValorAnterior = ValorAnterior / Valores.Text
   ElseIf Sinal = "*" Then
      ValorAnterior = ValorAnterior * Valores.Text
   End If
End If
Sinal = ""
KeyCode = 0
Valores.Text = ValorAnterior
ValorAnterior = 0

Valores.SetFocus
Valores.SelStart = 0
Valores.SelLength = Len(ValorAnterior)

ElseIf (KeyCode = 8 Or KeyCode = 37 Or KeyCode = 38 Or KeyCode = 39 Or KeyCode = 40 Or KeyCode = 46) And Shift = 0 Then '<-

Else
   KeyCode = 0
End If

End Sub
