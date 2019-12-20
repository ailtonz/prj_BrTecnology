Attribute VB_Name = "ADM"
Option Compare Database
Option Explicit

Public strTabela As String
Public varQualMes As Integer
Public varQualAno As Integer
Public varQualData As Date

Public Function QualMes()

QualMes = varQualMes

End Function

Public Function QualMesAnt()
Dim varQualMes2, varQualAno2
varQualMes2 = varQualMes - 1
If varQualMes2 = 0 Then
   varQualMes2 = 12
   varQualAno2 = varQualAno - 1
Else
   varQualAno2 = varQualAno
End If
QualMesAnt = varQualMes2
End Function

Public Function QualAnoAnt()
Dim varQualMes2, varQualAno2
varQualMes2 = varQualMes - 1
If varQualMes2 = 0 Then
   varQualMes2 = 12
   varQualAno2 = varQualAno - 1
Else
   varQualAno2 = varQualAno
End If
QualAnoAnt = varQualAno2
End Function


Public Function QualAno()

QualAno = varQualAno

End Function

Public Function QualData()

QualData = varQualData

End Function

Public Function NovoCodigo(Tabela, Campo)

Dim rstTabela As DAO.Recordset
Set rstTabela = CurrentDb.OpenRecordset("SELECT Max([" & Campo & "])+1 AS CodigoNovo FROM " & Tabela & ";")

If Not rstTabela.EOF Then
   NovoCodigo = rstTabela.Fields("CodigoNovo")
   If IsNull(NovoCodigo) Then
      NovoCodigo = 1
   End If
Else
   NovoCodigo = 1
End If
rstTabela.Close

End Function


Public Function SomarCampo(Tabela, Campo, Como)

Dim rstTabela As DAO.Recordset
Set rstTabela = CurrentDb.OpenRecordset("SELECT Sum([" & Campo & "]) AS " & Como & " FROM " & Tabela & ";")

If Not rstTabela.EOF Then
   SomarCampo = rstTabela.Fields(Como)
   If IsNull(SomarCampo) Then SomarCampo = 0
Else
   SomarCampo = 0
End If
   

rstTabela.Close

End Function

Public Function Pesquisar(Tabela As String)
                                   
On Error GoTo Err_Pesquisar
  
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Pesquisar"
    strTabela = Tabela
       
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    
Exit_Pesquisar:
    Exit Function

Err_Pesquisar:
    MsgBox Err.Description
    Resume Exit_Pesquisar
    
End Function


Public Sub LimparSistema()

DoCmd.SetWarnings False

DoCmd.RunSQL "Delete * From Faturamentos"
DoCmd.RunSQL "Delete * From Movimentos"

DoCmd.SetWarnings True

MsgBox "ATENÇÃO: Os cadastros de contas a Receber e a Pagar foram limpos!", vbInformation + vbOKOnly, "Limpar Sistema."

End Sub
