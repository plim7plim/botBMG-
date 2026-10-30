Set shell = CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")
pasta = fs.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = pasta
On Error Resume Next
shell.Run "pythonw.exe """ & pasta & "\iniciar.pyw""", 1, False
If Err.Number <> 0 Then
    Err.Clear
    shell.Run "pyw.exe -3 """ & pasta & "\iniciar.pyw""", 1, False
    If Err.Number <> 0 Then
        MsgBox "Python nao encontrado. Instale Python 3.12 ou superior e abra Iniciar botBMG.cmd.", 16, "botBMG"
    End If
End If
