Do
set WshShell = CreateObject("WScript.Shell")
WScript.Sleep 1
WshShell.SendKeys "{CAPSLOCK}"
WScript.Sleep 1
WshShell.SendKeys "{ScrollLock}"
WScript.Sleep 1
WshShell.SendKeys "{NumLock}"
Loop