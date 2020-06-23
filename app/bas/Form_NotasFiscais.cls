VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_NotasFiscais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Function AtualizarCampo(sql As String)
Dim rDados As DAO.Recordset

Set rDados = CurrentDb.OpenRecordset(sql)

If Not rDados.EOF Then

    AtualizarCampo = rDados.Fields(0)

End If

rDados.Close

End Function


Private Sub codCliente_Click()
    Me.Razao = Me.codCliente.Column(1)
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    
    If Me.NewRecord Then
       Me.Codigo = NovoCodigo(Me.RecordSource, Me.Codigo.ControlSource)
       Me.DataDeEmissao = Format(Now(), "dd/mm/yy")
       Me.DadoAdicional_1 = AtualizarCampo("SELECT DadosAdicionais.Adicional_01 FROM DadosAdicionais")
       Me.DadoAdicional_2 = AtualizarCampo("SELECT DadosAdicionais.Adicional_02 FROM DadosAdicionais")
       Me.DadoAdicional_3 = AtualizarCampo("SELECT DadosAdicionais.Adicional_03 FROM DadosAdicionais")
    End If
    
End Sub

Private Sub cmdSalvar_Click()
On Error GoTo Err_cmdSalvar_Click

    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    Form_Pesquisar.lstCadastro.Requery
    DoCmd.Close

Exit_cmdSalvar_Click:
    Exit Sub

Err_cmdSalvar_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_cmdSalvar_Click
End Sub

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click

    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
    DoCmd.CancelEvent
    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_cmdFechar_Click

End Sub

Private Sub NotasFiscaisItens_Exit(Cancel As Integer)
    Me.ValorTotalDosProdutos = Me.txtSomaProdutos
    Me.ValorTotalDaNota = Me.ValorTotalDosProdutos + Me.servTotal
    Me.BaseDeCalculoDoICMS = Me.txtSomaProdutos
    Me.ValorDoICMS = txtSomaICMS
End Sub
Private Sub cmdVisualizar_Click()
On Error GoTo Err_cmdVisualizar_Click

    Dim stDocName As String

    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    stDocName = "NotasFiscais"
    DoCmd.OpenReport stDocName, acPreview, , "codNotaFiscal = " & Me.Codigo

Exit_cmdVisualizar_Click:
    Exit Sub

Err_cmdVisualizar_Click:
    MsgBox Err.Description
    Resume Exit_cmdVisualizar_Click
    
End Sub

Private Sub servPorcento_AfterUpdate()
    
    Me.servValorDoISS = (Me.servTotal / 100) * Me.servPorcento
    
    If Not IsNull(Me.ValorTotalDaNota) Then
        Me.ValorTotalDaNota = Me.ValorTotalDaNota + Me.servTotal
    Else
        Me.ValorTotalDaNota = Me.servTotal
    End If
    
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    
End Sub
