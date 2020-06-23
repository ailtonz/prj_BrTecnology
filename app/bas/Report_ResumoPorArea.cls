VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_ResumoPorArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Open(Cancel As Integer)
    DoCmd.Maximize
End Sub

Private Sub Detalhe_Print(Cancel As Integer, PrintCount As Integer)

If Me.Calculo >= 0 Then
    Me.Calculo.ForeColor = RGB(0, 0, 0)
Else
    Me.Calculo.ForeColor = RGB(255, 0, 0)
End If

End Sub

Public Function Zebrar(rpt As Report)
Static fCinza As Boolean
Const conCinza = 15198183
Const conBranco = 16777215

On Error Resume Next

rpt.Section(0).BackColor = IIf(fCinza, conCinza, conBranco)
fCinza = Not fCinza

End Function


