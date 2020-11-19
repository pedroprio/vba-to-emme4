Attribute VB_Name = "ModImportResultados"
Public linha_importa As Integer

'Puxar arquivos gerados na pasta "resultados"
Sub ListarArquivosResultados()

Set wb = ActiveWorkbook
i = 1
Pasta = wb.Sheets("PRINCIPAL").Range("C4") & "resultados\"
tpArq = "*.txt"   ' Tipo de arquivos
Arq = Dir(Pasta & tpArq, vbDirectory)

wb.Sheets("arquivos").Range("D:D").Clear

While Arq <> ""
wb.Sheets("arquivos").Range("D" & i) = Arq
i = i + 1
Arq = Dir()
Wend
End Sub


'Importar arquivos de resultados
Sub ImportarArquivosResultados()
Call ListarArquivosResultados

Set wb = ActiveWorkbook

'Limpar Planilhas a Receber os resultados
With wb.Sheets("RESULT-ITINERARIES")
  .Activate
  .Range("A2:Z9999").ClearContents
  .Range("B2").Select
End With

With wb.Sheets("RESULT-LINHAS")
  .Activate
  .Range("A2:Z9999").ClearContents
  .Range("A2").Select
End With

With wb.Sheets("RESULT-NODES")
  .Activate
  .Range("A2:Z9999").ClearContents
  .Range("A2").Select
End With

With wb.Sheets("RESULT-IMPED")
  .Activate
  .Range("A2:Z9999").ClearContents
  .Range("A2").Select
End With

With wb.Sheets("RESULT-OD")
  .Activate
  .Range("A2:Z9999").ClearContents
  .Range("A2").Select
End With

With wb.Sheets("IMPED-OD")
  .Activate
  .Range("A2:Z9999").ClearContents
  .Range("A2").Select
End With

'Chamar a importação de acordo com os arquivos gerados
With wb.Sheets("arquivos")
  myrow = 1
  While .Cells(myrow, 4) <> ""
    Select Case True
      Case .Cells(myrow, 4) Like "transit_line_summary_hora*"
        Call Importa("RESULT-LINHAS", .Cells(myrow, 4), 10)
      Case .Cells(myrow, 4) Like "nodes_hora*"
        Call Importa("RESULT-NODES", .Cells(myrow, 4), 13)
      Case .Cells(myrow, 4) Like "itineraries_hora*"
        Call ImportaItinerarios("RESULT-ITINERARIES", .Cells(myrow, 4), 3)
      Case .Cells(myrow, 4) Like "matriz_tempos_hora*"
        Call ImportaMatrizes("RESULT-IMPED", .Cells(myrow, 4), 5)
      Case .Cells(myrow, 4) Like "matriz_od_hora*"
        Call ImportaMatrizes("RESULT-OD", .Cells(myrow, 4), 5)
      Case .Cells(myrow, 4) Like "matriz_imped_hora*"
        Call ImportaMatrizes("IMPED-OD", .Cells(myrow, 4), 5)
    End Select
    myrow = myrow + 1
  Wend
End With

If Not (batch) Then
MsgBox ("Resultados Importados")
End If

End Sub


'Sub de Importacao
Sub Importa(sheet_saida As String, arquivo As String, inicio_arquivo As Integer)

Set wb = ActiveWorkbook
fName = wb.Sheets("PRINCIPAL").Range("C4").Value & "resultados\" & arquivo

wb.Sheets(sheet_saida).Activate

    With wb.Sheets(sheet_saida).QueryTables.Add(Connection:="TEXT;" & fName, _
        Destination:=wb.Sheets(sheet_saida).Range(ActiveCell, ActiveCell))
            .Name = "sample"
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = False
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = 437
            .TextFileStartRow = inicio_arquivo
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierNone
            .TextFileConsecutiveDelimiter = True
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = False
            .TextFileSpaceDelimiter = True
            .TextFileOtherDelimiter = "" & Chr(10) & ""
            .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, _
               1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
               1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
               1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
    End With
    
wb.Connections.Item(ActiveWorkbook.Connections.Count).Delete

'Preenchimento das Horas
hora = Mid(arquivo, InStr(arquivo, "hora") + 4, 2)

While ActiveCell.Offset(0, 1).Value <> ""
  ActiveCell.Value = hora
  ActiveCell.Offset(1, 0).Select
Wend

'Limpar linhas não desejadas da importação
wb.Sheets(sheet_saida).Range(ActiveCell, ActiveCell.Offset(9999, 99)).ClearContents
End Sub


'Sub especial para itineraries
Sub ImportaItinerarios(sheet_saida As String, arquivo As String, inicio_arquivo As Integer)

Set wb = ActiveWorkbook
fName = wb.Sheets("PRINCIPAL").Range("C4").Value & "resultados\" & arquivo

wb.Sheets(sheet_saida).Activate

    With wb.Sheets(sheet_saida).QueryTables.Add(Connection:="TEXT;" & fName, _
        Destination:=wb.Sheets(sheet_saida).Range(ActiveCell, ActiveCell))
            .Name = "sample"
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = False
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = 437
            .TextFileStartRow = inicio_arquivo
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierNone
            .TextFileConsecutiveDelimiter = True
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = False
            .TextFileSpaceDelimiter = True
            .TextFileOtherDelimiter = "" & Chr(10) & ""
            .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, _
               1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
               1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
               1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
    End With
    
wb.Connections.Item(ActiveWorkbook.Connections.Count).Delete

'Preenchimento das Horas
hora = Mid(arquivo, InStr(arquivo, "hora") + 4, 2)

While cont_skip < 30
    If (ActiveCell.Offset(0, 1) = "Transit") Then
        transit_line = ActiveCell.Offset(0, 3)
    End If
    If IsNumeric(ActiveCell.Offset(0, 1)) And (ActiveCell.Offset(0, 1) > 0) And (ActiveCell.Offset(0, 3) <> "-") Then
        cont_skip = 0
        ActiveCell.Offset(0, -1).Value = hora
        ActiveCell.Value = transit_line
        ActiveCell.Offset(1, 0).Select
    Else
        ActiveCell.EntireRow.Delete
        cont_skip = cont_skip + 1
    End If

Wend
End Sub

'Sub especial para matrizes
Sub ImportaMatrizes(sheet_saida As String, arquivo As String, inicio_arquivo As Integer)

Set wb = ActiveWorkbook
fName = wb.Sheets("PRINCIPAL").Range("C4").Value & "resultados\" & arquivo

wb.Sheets(sheet_saida).Activate

    With wb.Sheets(sheet_saida).QueryTables.Add(Connection:="TEXT;" & fName, _
        Destination:=wb.Sheets(sheet_saida).Range(ActiveCell, ActiveCell))
            .Name = "sample"
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = False
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = 437
            .TextFileStartRow = inicio_arquivo
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierNone
            .TextFileConsecutiveDelimiter = True
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = False
            .TextFileSpaceDelimiter = True
            .TextFileOtherDelimiter = ":"
            .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, _
               1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
               1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
               1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
    End With
    
wb.Connections.Item(ActiveWorkbook.Connections.Count).Delete

'Preenchimento das Horas
hora = Mid(arquivo, InStr(arquivo, "hora") + 4, 2)

While ActiveCell.Offset(0, 1).Value <> ""
  ActiveCell.Value = hora
  ActiveCell.Offset(1, 0).Select
Wend

'Limpar linhas não desejadas da importação
wb.Sheets(sheet_saida).Range(ActiveCell, ActiveCell.Offset(9999, 99)).ClearContents
End Sub



