Attribute VB_Name = "ModArquivosGrade"
Sub ArquivosGrade()
  Call Importar_Arquivo_Grade
  Call Preenche_Hora
  Call Procv_Linhas_Marchas
  Call Reordenar_Prefixos
  Call Preencher_Headways
  Call Salvar_Headways
End Sub

Sub Listar_Arquivos_Grade()

Set wb = ActiveWorkbook
i = 1
Pasta = wb.Sheets("PRINCIPAL").Range("C19")
tpArq = "*.xlsx"   ' Tipo de arquivos
Arq = Dir(Pasta & tpArq, vbDirectory)

wb.Sheets("arquivos").Range("A:A").Clear

While Arq <> ""
wb.Sheets("arquivos").Range("A" & i) = Arq
i = i + 1
Arq = Dir()
Wend
End Sub

Sub Importar_Arquivo_Grade()
'Desativa Alertas
Application.DisplayAlerts = False

'Memorizar workbook atual
Set wb = ActiveWorkbook

'Limpar planilhas a receber dados
wb.Sheets("prefixos").Range("A:Z").ClearContents

'Abrir arquivo com grades
fName = wb.Sheets("PRINCIPAL").Range("C19").Value & wb.Sheets("PRINCIPAL").Range("C21").Value
Set wbGrade = Workbooks.Open(fName)

'Copiar prefixos
With wbGrade.Sheets("Prefixos")
    .Activate
    .Range("A1", Range("H1").End(xlDown)).Copy
End With

'Colar prefixos
With wb.Sheets("prefixos")
    .Activate
    .Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End With

'Fechar arquivo com grades
wbGrade.Close

'Reativa Alertas
Application.DisplayAlerts = True
End Sub

Sub Preenche_Hora()

Set wb = ActiveWorkbook

  myrow = 2
  With wb.Sheets("prefixos")
  .Range("I1").Value = "H_PARTIDA"
    While .Cells(myrow, 8) <> ""
     .Cells(myrow, 9) = Hour(.Cells(myrow, 5))
     myrow = myrow + 1
    Wend
  End With
    
End Sub

Sub Procv_Linhas_Marchas()

Set wb = ActiveWorkbook

  myrow = 2
  With wb.Sheets("prefixos")
  .Range("J1").Value = "LINHA"
    While .Cells(myrow, 8) <> ""
     .Cells(myrow, 10) = Application.VLookup(.Cells(myrow, 8), wb.Sheets("linhas-marchas").Range("C:E"), 3, False)
     myrow = myrow + 1
    Wend
  End With
End Sub


Sub Reordenar_Prefixos()

Set wb = ActiveWorkbook
wb.Sheets("prefixos").Activate

'Reordena planilha prefixos
With wb.Sheets("prefixos").Sort
    .SortFields.Clear
    
    .SortFields.Add Key:=Range( _
    "J:J"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
    
    .SortFields.Add Key:=Range( _
    "E:E"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
    .SetRange Range("A:Q")
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

End Sub

Sub Preencher_Headways()

Set wb = ActiveWorkbook
wb.Sheets("prefixos").Activate

  myrow = 2
  With wb.Sheets("prefixos")
  .Range("K1").Value = "HEADWAY"
    While .Cells(myrow, 10) <> ""
     'Caso não seja a primeira viagem da linha
     If .Cells(myrow, 10) = .Cells(myrow - 1, 10) Then
     
       'Caso a diferença para a viagem anterior seja inferior a 1 hora
       If Hour(.Cells(myrow, 5) - .Cells(myrow - 1, 5)) < 1 Then
          .Cells(myrow, 11) = Minute(.Cells(myrow, 5) - .Cells(myrow - 1, 5))
          
       'Caso a diferença para a viagem anterior seja igual ou superior a 1 hora
       Else
          .Cells(myrow, 11) = 60
       End If
       
     'Caso seja a primeira viagem da linha
     Else
     
       'Caso seja a unica viagem da hora
       If Hour(.Cells(myrow, 5)) <> Hour(.Cells(myrow + 1, 5)) Then
         .Cells(myrow, 11) = 60
       Else
         'Caso seja a unica viagem da linha em questão
         If .Cells(myrow, 10) <> .Cells(myrow + 1, 10) Then
          .Cells(myrow, 11) = 60
         End If
       End If
     End If
     myrow = myrow + 1
    Wend
  End With

End Sub

Sub Salvar_Headways()
Set wb = ActiveWorkbook

'Desabilita o cálculo antes de salvar
Application.CalculateBeforeSave = False

'Desabilita os avisos de sobrescrever
Application.DisplayAlerts = False

'Guarda o caminho do arquivo para salvar no fim
ArqPrincipal = ThisWorkbook.FullName

'Inicializa o nome dos arquivos de matrizes
Dim caminho As String
caminho = wb.Sheets("PRINCIPAL").Range("C4").Value & "headways\"

'Limpa a Pasta de Headways
Call DeletaPasta(caminho, ".211")

'Atualiza Planilha dos Headways Automaticos
wb.Sheets("HDW-FORMULA").Calculate

'Inicializa a seleção de horas
wb.Sheets("HDW").Activate
Range("A2").Select


Do While ActiveCell.Value <> "fim"
    
    'Atualiza para a hora do loop
    hora = ActiveCell.Value
    Range("I2").Value = hora
    
    'Atualizar Planilhas para início dos calculos
    wb.Sheets("HDW").Calculate
    wb.Sheets("HDW-SCRIPT").Calculate
    
    'Atualiza o nome do arquivo para a hora do loop e escreve na planilha
    arquivo = caminho & "hdw-h" & Format(hora, "00") & ".211"
    ActiveCell.Offset(0, 2).Value = arquivo
   
    ActiveCell.Offset(0, 1).Value = WorksheetFunction.Average(Range("D29:D54"))
   
    'Ativa a planilha de cálculo e salva-a como CSV
    wb.Sheets("HDW-SCRIPT").Activate
    wb.SaveAs Filename:=arquivo, FileFormat:=xlCSV, CreateBackup:=False
    
    'Reestabelece o nome da planilha de cálculo
    ActiveSheet.Name = "HDW-SCRIPT"
    
    wb.Sheets("HDW").Activate
    ActiveCell.Offset(1, 0).Select
Loop


'Salva novamente o arquivo original
wb.SaveAs Filename:=ArqPrincipal, FileFormat:=52, CreateBackup:=False

Application.DisplayAlerts = True
Application.CalculateBeforeSave = True
wb.Sheets("PRINCIPAL").Activate

If Not (batch) Then
MsgBox ("Scripts .211 exportados")
End If

End Sub
