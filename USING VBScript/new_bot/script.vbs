Set objShell = CreateObject("WScript.Shell")

' Use NirCmd to send a refresh command to the active Explorer window
objShell.Run "nircmd.exe cmdwait 1000 sendkeypress {F5}", 0, True
