Attribute VB_Name = "ModRodarEmme"
Sub rodar_emme()

Dim myEmmeProjectFolder
Dim myMatrixRef
Dim myArray
Dim myOArray
Dim Denom
Dim myscenario As Variant

t_start = Time


'Read EMME Database Location
EmmeProjectFolder = Range("C8")
If Right(EmmeProjectFolder, 1) <> "\" Then
    EmmeProjectFolder = EmmeProjectFolder & "\"
End If
EmmeDatabaseFolder = EmmeProjectFolder & "Database\"

'Ler Caminho arquivo
caminho = Range("C4")
If Right(caminho, 1) <> "\" Then
    caminho = caminho & "\"
End If

'Limpa a Pasta Resultados e Logs
Call DeletaPasta(caminho & "resultados\", ".txt")
Call DeletaPasta(caminho & "logs\", ".txt")

'Read Emme Macros Location
EmmeMacrosFolder = caminho & "macros\"
EmmeMacrosFolder = Replace(EmmeMacrosFolder, "\", "/")

'Read Emme Program Executable Location
Emmeprogramsfolder = Range("C6")
If Right(Emmeprogramsfolder, 1) <> "\" Then
    Emmeprogramsfolder = Emmeprogramsfolder & "\"
End If

'Prepara bat que roda macro
macro_name = "batch.mac"
    
Open caminho & "Run.Bat" For Output As #1
Print #1, "title Modelo Ocupacao SPV EMME"
Print #1, "color 0A"
Print #1, "path=" & Emmeprogramsfolder
Print #1, "set EMACPATH=%EMACPATH%;""" & EmmeMacrosFolder & """"
Print #1, "cd " & EmmeDatabaseFolder
Print #1, "call emme -ng 000 -m " & macro_name
Print #1, "del ~processing.now"
Close #1
Application.StatusBar = "Rodando modelo EMME ..."

'Rodar Emme
Open EmmeDatabaseFolder & "~processing.now" For Output As #1
        Print #1, ""
    Close #1
    
Application.Wait (Now + TimeValue("0:00:02"))
    
ChDrive (Left(caminho, 2))
ChDir caminho
Shell ("""" & caminho & "Run.bat"""), vbNormalFocus

myTime = Now
Fim = False
While Not Fim
    If (Not fileexists(EmmeDatabaseFolder + "~processing.now")) Then
        Fim = True
    End If
    Application.Wait (myTime + TimeValue("00:00:01"))
    myTime = Now
Wend

Kill (caminho & "Run.bat")

'Fim
t_end = Time

If Not (batch) Then
MsgBox ("Finalizado em " & Round(((t_end - t_start) * 24 * 60 * 60), 0) & " segundos.")
End If

Application.StatusBar = False

End Sub

Function fileexists(fullfname As String, Optional ispath As Boolean = False) As Boolean

    Dim fs As Object
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    If (ispath) Then
        fileexists = fs.FolderExists(fullfname)
    Else
        fileexists = fs.fileexists(fullfname)
    End If

End Function
