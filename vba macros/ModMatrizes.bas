Attribute VB_Name = "ModMatrizes"
Sub ArquivosMatrizes()
  Call Importar_Arquivo_Matrizes
  Call preenche_ODzona
  Call preenche_FatorExp
  Call preenche_ExpansaoOD
  Call preenche_ResumoHoraExpansaoOD
  Call preenche_ExpansaoSimulacao
  Call preenche_ResumoHoraExpansaoSimulacao
  Call salvar_matrizes
End Sub

Sub Listar_Arquivos_Matrizes()

Set wb = ActiveWorkbook

'Arquivos de Matrizes
i = 1
Pasta = wb.Sheets("PRINCIPAL").Range("C4") & "demanda-matriz\"
tpArq = "*.xlsx"   ' Tipo de arquivos
Arq = Dir(Pasta & tpArq, vbDirectory)

wb.Sheets("arquivos").Range("B:B").Clear

While Arq <> ""
wb.Sheets("arquivos").Range("B" & i) = Arq
i = i + 1
Arq = Dir()
Wend

'Arquivos de Hora-Hora
i = 1
Pasta = wb.Sheets("PRINCIPAL").Range("C4") & "demanda-hora\"
tpArq = "*.xlsx"   ' Tipo de arquivos
Arq = Dir(Pasta & tpArq, vbDirectory)

wb.Sheets("arquivos").Range("C:C").Clear

While Arq <> ""
wb.Sheets("arquivos").Range("C" & i) = Arq
i = i + 1
Arq = Dir()
Wend

End Sub

Sub Importar_Arquivo_Matrizes()
'Desativa Alertas
Application.DisplayAlerts = False

'Memorizar workbook atual
Set wb = ActiveWorkbook

'Limpar planilhas a receber dados
wb.Sheets("HORA-HORA").Range("A2", wb.Sheets("HORA-HORA").Range("Z2").End(xlDown)).ClearContents
wb.Sheets("OD-TOTAL").Range("A2", wb.Sheets("OD-TOTAL").Range("H2").End(xlDown)).ClearContents

'Abrir arquivo com Matrizes OD
fName = wb.Sheets("PRINCIPAL").Range("C4").Value & "demanda-matriz\" & wb.Sheets("PRINCIPAL").Range("C11").Value
Set wbMatriz = Workbooks.Open(fName)

'Copiar prefixos
With wbMatriz.ActiveSheet
    .Activate
    .Range("A1", .Range("E1").End(xlDown)).Copy
End With

'Colar prefixos
With wb.Sheets("OD-TOTAL")
    .Activate
    .Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End With

'Fechar arquivo com grades
wbMatriz.Close

'Abrir arquivo com Hora-Hora
fName = wb.Sheets("PRINCIPAL").Range("C4").Value & "demanda-hora\" & wb.Sheets("PRINCIPAL").Range("H11").Value
Set wbHora = Workbooks.Open(fName)

'Copiar prefixos
With wbHora.ActiveSheet
    .Activate
    .Range("Z1", .Range("A1").End(xlDown)).Copy
End With

'Colar prefixos
With wb.Sheets("HORA-HORA")
    .Activate
    .Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End With

'Fechar arquivo com grades
wbHora.Close

'Reativa Alertas
Application.DisplayAlerts = True

'Preenche data simulada
wb.Sheets("PRINCIPAL").Range("H13").Value = Format(wb.Sheets("HORA-HORA").Range("A2").Value, "yyyymmdd")
  
End Sub

Sub preenche_ODzona()
Set wb = ActiveWorkbook

  myrow = 2
  With wb.Sheets("OD-TOTAL")
  .Range("F1").Value = "CONCAT OD-ZONAS"
    While .Cells(myrow, 1) <> ""
     zona_origem = Application.VLookup(.Cells(myrow, 3), wb.Sheets("ZNS-EMME").Range("B:C"), 2, False)
     zona_destino = Application.VLookup(.Cells(myrow, 4), wb.Sheets("ZNS-EMME").Range("B:C"), 2, False)
     .Cells(myrow, 6) = .Cells(myrow, 2) & "-" & zona_origem & "-" & zona_destino
     myrow = myrow + 1
    Wend
  End With
    
End Sub

Sub preenche_FatorExp()
Set wb = ActiveWorkbook

'Limpar planilhas a receber dados
wb.Sheets("EXP").Range("D:F").ClearContents

  myrow = 2
  With wb.Sheets("EXP")
  .Range("D1").Value = "PAX OD"
  .Range("E1").Value = "PAX REAL"
  .Range("F1").Value = "FATOR EXPANSAO"
    While .Cells(myrow, 2) <> ""
      .Cells(myrow, 4) = Application.SumIfs(wb.Sheets("OD-TOTAL").Range("E:E"), wb.Sheets("OD-TOTAL").Range("B:B"), .Cells(myrow, 3), wb.Sheets("OD-TOTAL").Range("C:C"), .Cells(myrow, 2))
      .Cells(myrow, 5) = Application.VLookup(.Cells(myrow, 2), wb.Sheets("HORA-HORA").Range("B:Z"), .Cells(myrow, 3) + 2, False)
      
      If Not (IsNumeric(.Cells(myrow, 5))) Then .Cells(myrow, 5) = 0
      
      If .Cells(myrow, 4) > 0 Then
           .Cells(myrow, 6) = .Cells(myrow, 5) / .Cells(myrow, 4)
      Else
        .Cells(myrow, 6) = 0
      End If
     myrow = myrow + 1
    Wend
  End With

End Sub

Sub preenche_ExpansaoOD()
Set wb = ActiveWorkbook

  myrow = 2
  With wb.Sheets("OD-TOTAL")
  .Range("G1").Value = "TOTAL EXPANDIDO"
    While .Cells(myrow, 1) <> ""
     .Cells(myrow, 7) = .Cells(myrow, 5) * Application.VLookup(.Cells(myrow, 3) & "-" & .Cells(myrow, 2), wb.Sheets("EXP").Range("A:F"), 6, False)
     myrow = myrow + 1
    Wend
  End With
End Sub

Sub preenche_ResumoHoraExpansaoOD()
Set wb = ActiveWorkbook

  With wb.Sheets("MATRIZES")
    myrow = 8
    While .Cells(myrow, 1) <> "fim"
     .Cells(myrow, 11) = Application.SumIf(wb.Sheets("OD-TOTAL").Range("B:B"), .Cells(myrow, 1), wb.Sheets("OD-TOTAL").Range("G:G"))
     myrow = myrow + 1
    Wend
  .Cells(32, 11) = Application.Sum(.Range("K8:K31"))
  
  'Preenche demanda de entrada
  wb.Sheets("PRINCIPAL").Cells(13, 3) = .Cells(32, 11)
  
  'Atualiza Fator de Expansão
  wb.Sheets("PRINCIPAL").Cells(15, 4) = wb.Sheets("PRINCIPAL").Cells(15, 3) / wb.Sheets("PRINCIPAL").Cells(13, 3) - 1
  
  'Busca Fatores de Expansão para planilha "Matrizes"
    myrow = 8
    While .Cells(myrow, 1) <> "fim"
     fExp = wb.Sheets("PRINCIPAL").Cells(15, 4)
     .Cells(myrow, 10) = (1 + fExp) * (1 + .Cells(myrow, 9))
     myrow = myrow + 1
    Wend
  
  End With
End Sub


Sub preenche_ExpansaoSimulacao()
Set wb = ActiveWorkbook

  myrow = 2
  With wb.Sheets("OD-TOTAL")
  .Range("H1").Value = "TOTAL SIMULADO"
    While .Cells(myrow, 1) <> ""
     .Cells(myrow, 8) = .Cells(myrow, 7) * Application.VLookup(.Cells(myrow, 2), wb.Sheets("MATRIZES").Range("A:J"), 10, False)
     myrow = myrow + 1
    Wend
  End With
End Sub

Sub preenche_ResumoHoraExpansaoSimulacao()
Set wb = ActiveWorkbook

  With wb.Sheets("MATRIZES")
    myrow = 8
    While .Cells(myrow, 1) <> "fim"
     .Cells(myrow, 12) = Application.SumIf(wb.Sheets("OD-TOTAL").Range("B:B"), .Cells(myrow, 1), wb.Sheets("OD-TOTAL").Range("H:H"))
     myrow = myrow + 1
    Wend
  .Cells(32, 12) = Application.Sum(.Range("L8:L31"))
  .Cells(33, 12) = .Cells(32, 12) / .Cells(32, 11) - 1
  End With
End Sub

Sub salvar_matrizes()

ActiveWorkbook.Sheets("MATRIZES").Activate

Range("M8:M31").ClearContents

'Desabilita o cálculo antes de salvar
Application.CalculateBeforeSave = False

'Desabilita os avisos de sobrescrever
Application.DisplayAlerts = False

'Guarda o caminho do arquivo para salvar no fim
ArqPrincipal = ThisWorkbook.FullName

'Guarda a data para a nomenclatura dos arquivos
Data = ActiveWorkbook.Sheets("PRINCIPAL").Range("H13").Value

'Inicializa o nome dos arquivos de matrizes
Dim caminho As String
caminho = ActiveWorkbook.Sheets("PRINCIPAL").Range("C4").Value & "matrizes\"

'Limpa a Pasta
Call DeletaPasta(caminho, ".311")

'Inicializa a seleção de horas
Range("A8").Select

Do While ActiveCell.Value <> "fim"
    
    'Atualiza para a hora do loop e realiza o calculo
    hora = ActiveCell.Value
    Range("H2").Value = hora
    ActiveWorkbook.Sheets("MTZ-CALCULO").Calculate

    'Atualiza o nome do arquivo para a hora do loop e escreve na planilha
    arquivo = caminho & Data & "-OD" & Format(hora, "00") & ".311"
    ActiveCell.Offset(0, 1).Value = arquivo
    
    'Escreve a soma da matriz na planilha para conferência
    ActiveCell.Offset(0, 12).Value = Application.Sum(Sheets("MTZ-CALCULO").Range("C:C"))
    
    'Ativa a planilha de cálculo e salva-a como CSV
    ActiveWorkbook.Sheets("MTZ-CALCULO").Activate
    ActiveWorkbook.SaveAs Filename:=arquivo, FileFormat:=xlTextPrinter, CreateBackup:=False
    
    'Reestabelece o nome da planilha de cálculo
    ActiveSheet.Name = "MTZ-CALCULO"
    
    ActiveWorkbook.Sheets("MATRIZES").Activate
    ActiveCell.Offset(1, 0).Select
Loop

'Atualiza a soma
  ActiveWorkbook.Sheets("MATRIZES").Cells(32, 12) = Application.Sum(ActiveWorkbook.Sheets("MATRIZES").Range("M8:M31"))

'Retorna para a Planilha Principal
ActiveWorkbook.Sheets("PRINCIPAL").Activate

'Salva novamente o arquivo original
ActiveWorkbook.SaveAs Filename:=ArqPrincipal, FileFormat:=52, CreateBackup:=False

Application.DisplayAlerts = True
Application.CalculateBeforeSave = True
If Not (batch) Then
MsgBox ("Scripts .311 exportados")
End If
End Sub
