Option Explicit

Dim shell, fso, baseDir, pythonExe, scriptPath, reportPath, command, exitCode, reportText

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
scriptPath = baseDir & "\modelo\importar_pasta_propostas.py"
reportPath = baseDir & "\resultados\nao-mexer\resultado_importacao.txt"
pythonExe = FindPython()

If Not fso.FileExists(pythonExe) Then
    MsgBox "Python nao foi encontrado nesta maquina." & vbCrLf & vbCrLf & _
           "Instale o Python 3 ou coloque uma versao portatil em:" & vbCrLf & _
           baseDir & "\modelo\python\pythonw.exe", vbCritical, "Importar propostas"
    WScript.Quit 1
End If

If Not fso.FileExists(scriptPath) Then
    MsgBox "Script nao encontrado em:" & vbCrLf & scriptPath, vbCritical, "Importar propostas"
    WScript.Quit 1
End If

command = """" & pythonExe & """ """ & scriptPath & """"
exitCode = shell.Run(command, 0, True)

If fso.FileExists(reportPath) Then
    reportText = ReadUtf8File(reportPath)
Else
    reportText = "Importacao finalizada, mas o relatorio nao foi encontrado."
End If

If exitCode = 0 Then
    MsgBox reportText, vbInformation, "Importar propostas"
Else
    MsgBox reportText, vbExclamation, "Importar propostas"
End If

Function FindPython()
    Dim candidates, version, path, i, found

    ReDim candidates(0)
    candidates(0) = baseDir & "\modelo\python\pythonw.exe"

    For version = 314 To 38 Step -1
        AddCandidate candidates, shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Programs\Python\Python" & version & "\pythonw.exe"
        AddCandidate candidates, shell.ExpandEnvironmentStrings("%ProgramFiles%") & "\Python" & version & "\pythonw.exe"
        AddCandidate candidates, shell.ExpandEnvironmentStrings("%ProgramFiles(x86)%") & "\Python" & version & "\pythonw.exe"
    Next

    AddCandidate candidates, "C:\Python314\pythonw.exe"
    AddCandidate candidates, "C:\Python313\pythonw.exe"
    AddCandidate candidates, "C:\Python312\pythonw.exe"
    AddCandidate candidates, "C:\Python311\pythonw.exe"
    AddCandidate candidates, "C:\Python310\pythonw.exe"
    AddCandidate candidates, "C:\Python39\pythonw.exe"
    AddCandidate candidates, "C:\Python38\pythonw.exe"

    found = FirstFromWhere("pythonw.exe")
    If found <> "" Then AddCandidate candidates, found

    found = FirstFromWhere("pyw.exe")
    If found <> "" Then AddCandidate candidates, found

    found = FirstFromWhere("python.exe")
    If found <> "" Then AddCandidate candidates, found

    found = FirstFromWhere("py.exe")
    If found <> "" Then AddCandidate candidates, found

    For i = 0 To UBound(candidates)
        path = candidates(i)
        If path <> "" Then
            If fso.FileExists(path) Then
                FindPython = path
                Exit Function
            End If
        End If
    Next

    FindPython = ""
End Function

Sub AddCandidate(ByRef candidates, ByVal path)
    Dim nextIndex
    nextIndex = UBound(candidates) + 1
    ReDim Preserve candidates(nextIndex)
    candidates(nextIndex) = path
End Sub

Function FirstFromWhere(ByVal executableName)
    Dim exec, line
    On Error Resume Next
    Set exec = shell.Exec("%ComSpec% /c where " & executableName)
    If Err.Number <> 0 Then
        Err.Clear
        FirstFromWhere = ""
        Exit Function
    End If
    On Error GoTo 0

    Do While Not exec.StdOut.AtEndOfStream
        line = Trim(exec.StdOut.ReadLine)
        If InStr(1, line, "\WindowsApps\", vbTextCompare) = 0 Then
            FirstFromWhere = line
            Exit Function
        End If
    Loop

    FirstFromWhere = ""
End Function

Function ReadUtf8File(ByVal path)
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    ReadUtf8File = stream.ReadText
    stream.Close
End Function
