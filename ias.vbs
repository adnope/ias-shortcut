Set objShell = CreateObject("WScript.Shell")
strScriptPath = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\") - 1)
objShell.Run "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & strScriptPath & "\ias.ps1""", 0, True
