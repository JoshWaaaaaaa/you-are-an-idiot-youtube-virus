do
Dim URL,WshShell,i
URL = "https://www.youtube.com/watch?v=JQF9M_aT6l4"
Set WshShell = CreateObject("WScript.shell")
For i = 0 to 50
    WshShell.SendKeys(chr(175))
Next
loop