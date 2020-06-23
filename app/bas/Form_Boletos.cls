VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Boletos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim WithEvents mRelatorio As Report
Attribute mRelatorio.VB_VarHelpID = -1

Private Sub Form_Open(Cancel As Integer)

codboletos.DefaultValue = NovoCodigo(Me.RecordSource, Me.codboletos.ControlSource)

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
On Error GoTo Err_cmdDesfazer_Click

    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
    DoCmd.CancelEvent
    DoCmd.Close
    
Exit_cmdDesfazer_Click:
    Exit Sub

Err_cmdDesfazer_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_cmdDesfazer_Click

End Sub


Private Sub cmdVisualizar_Click()
On Error GoTo Err_cmdVisualizar_Click


 Set mRelatorio = New Report_rptBoletos
 
  With mRelatorio
   .Caption = "Visualizando: " & codboletos.Value
   .Filter = "codboletos = " & codboletos.Value
   .FilterOn = True
   .Visible = True
  End With

Exit_cmdVisualizar_Click:
    Exit Sub

Err_cmdVisualizar_Click:
    MsgBox Err.Description
    Resume Exit_cmdVisualizar_Click
    
End Sub
