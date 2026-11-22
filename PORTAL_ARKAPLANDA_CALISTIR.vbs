Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "C:\isg"
WshShell.Run "cmd.exe /c ""C:\isg\baslat.bat""", 0, False
