VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Duplicatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdHistorico_Click()
On Error GoTo Err_cmdHistorico_Click

    Dim stDocName As String
    
    'Salvar Registro
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    
    stDocName = "Duplicatas_Historico"
    DoCmd.OpenReport stDocName, acPreview, , "codDuplicata = " & Me.Codigo

Exit_cmdHistorico_Click:
    Exit Sub

Err_cmdHistorico_Click:
    MsgBox Err.Description
    Resume Exit_cmdHistorico_Click
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    If Me.NewRecord Then
       Me.Codigo = NovoCodigo(Me.RecordSource, Me.Codigo.ControlSource)
       Me.Emissao = Format(Now(), "dd/mm/yy")
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

Private Sub cmdVisualizar_Click()
On Error GoTo Err_cmdVisualizar_Click

    Dim stDocName As String
    
    'Salvar Registro
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    
    stDocName = "Duplicatas"
    DoCmd.OpenReport stDocName, acPreview, , "codDuplicata = " & Me.Codigo

Exit_cmdVisualizar_Click:
    Exit Sub

Err_cmdVisualizar_Click:
    MsgBox Err.Description
    Resume Exit_cmdVisualizar_Click
    
End Sub
