Attribute VB_Name = "ModBásico"
Public batch As Boolean

Sub obter_caminho_atualizado()
 caminho = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "")
 Range("C4").Value = caminho
End Sub

Sub mudar_delimitadores()

'Muda Separadores de Sistema

If Application.UseSystemSeparators = True Then
Sheets("PRINCIPAL").Shapes("Button12").TextFrame.Characters.Text = "Mudar para Português"
With Application
     .DecimalSeparator = "."
     .ThousandsSeparator = ","
     .UseSystemSeparators = False
End With

Else
Sheets("PRINCIPAL").Shapes("Button12").TextFrame.Characters.Text = "Mudar para Inglês"
With Application
     .DecimalSeparator = "."
     .ThousandsSeparator = ","
     .UseSystemSeparators = True
End With
End If
End Sub

Sub DeletaPasta(Pasta As String, Ext As String)
'You can use this to delete all the files in the folder Test
    On Error Resume Next
    Kill Pasta & "*" & Ext
    On Error GoTo 0
End Sub

Sub rodar_tudo()

batch = True
t_start = Time

'Muda Separadores de Sistema, se necessário
If Application.UseSystemSeparators = True Then
    Call mudar_delimitadores
    retorna_delimit = True
End If

Call obter_caminho_atualizado
Call ArquivosMatrizes
Call ArquivosGrade
Call criar_macros
Call rodar_emme
Call ImportarArquivosResultados

ActiveWorkbook.Sheets("PRINCIPAL").Activate

'Retorna aos Separadores de Sistema originais
If retorna_delimit = True Then
    Call mudar_delimitadores
End If

batch = False
t_end = Time

MsgBox ("Alocação completa finalizada em " & Round(((t_end - t_start) * 24 * 60 * 60), 0) & " segundos.")

End Sub
