Attribute VB_Name = "ModSalvarMacro"
Sub criar_macros()

'Atualizar Planilhas para início dos calculos
ActiveWorkbook.Sheets("MACRO").Calculate

EMME_macro_path = ActiveWorkbook.Sheets("PRINCIPAL").Range("C4").Value & "Macros"

myrow = 2
While Sheets("MACRO").Cells(myrow, 1) <> ""
    If Sheets("MACRO").Cells(myrow, 1) <> Sheets("MACRO").Cells(myrow - 1, 1) Then
        If myrow <> 2 Then
            Close #1
        End If
        Open EMME_macro_path + "\" + Sheets("MACRO").Cells(myrow, 1) For Output As #1
    End If
    Print #1, Trim(Sheets("MACRO").Cells(myrow, 2))
    myrow = myrow + 1
Wend
Close #1

If Not (batch) Then
MsgBox ("Macros atualizadas!")
End If

End Sub


