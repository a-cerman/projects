Option Explicit
Dim fso, shell, folder, python, command
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
folder = fso.GetParentFolderName(WScript.ScriptFullName)
python = fso.BuildPath(folder, ".venv\Scripts\pythonw.exe")
If Not fso.FileExists(python) Then
    MsgBox "Python environment not found. Follow the setup instructions in README.md.", 48, "Narrated Presentations"
    WScript.Quit 1
End If
shell.CurrentDirectory = folder
command = Chr(34) & python & Chr(34) & " " & Chr(34) & fso.BuildPath(folder, "app.py") & Chr(34)
shell.Run command, 0, False
