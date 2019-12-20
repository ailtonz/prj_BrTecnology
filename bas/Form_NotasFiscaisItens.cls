VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_NotasFiscaisItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub ICMS_AfterUpdate()
    Me.ICMS_Valor = (Me.ValorTotal / 100) * Me.ICMS
End Sub

Private Sub ICMS_Exit(Cancel As Integer)
    Call ICMS_AfterUpdate
End Sub

Private Sub Quantidade_Exit(Cancel As Integer)
    If Not IsNull(Me.DescricaoDoProduto) Then Me.ValorTotal = Me.Quantidade * Me.ValorUnitario
End Sub

Private Sub ValorUnitario_Exit(Cancel As Integer)
    If Not IsNull(Me.DescricaoDoProduto) Then Me.ValorTotal = Me.Quantidade * Me.ValorUnitario
End Sub
