Set objShell = CreateObject("WScript.Shell")

' Navigate to Chrome application folder
chromePath = objShell.ExpandEnvironmentStrings("%ProgramFiles%\Google\Chrome\Application\")

' Copy everything in Chrome application folder to current directory
objShell.Run "xcopy /s /y """ & chromePath & "*.*"" .\", 0, True

' Rename chrome.exe to abc.exe in current directory
Set objFSO = CreateObject("Scripting.FileSystemObject")
chromeExePath = ".\chrome.exe"
abcExePath = ".\abc.exe"
objFSO.MoveFile chromeExePath, abcExePath
MsgBox "chrome.exe renamed to abc.exe in current directory"

' Create a shortcut for abc.exe with the flag --disable-extensions
Set objShortcut = objShell.CreateShortcut(objShell.SpecialFolders("Desktop") & "\abc.lnk")

' Get the user's profile directory
userProfile = objShell.ExpandEnvironmentStrings("%USERPROFILE%")

' Construct the path to the Downloads folder
downloadsPath = userProfile & "\Downloads"

' Set the TargetPath of the shortcut to the Downloads folder
objShortcut.TargetPath = downloadsPath & "\abc.exe"

objShortcut.Arguments = " --disable-extensions"
objShortcut.Save
MsgBox "Shortcut created on Desktop for abc.exe with --disable-extensions flag"

' Close all Chrome tabs
objShell.Run "taskkill /f /im chrome.exe", 0, True
MsgBox "All Chrome tabs closed"

' Open the shortcut
Set objShortcut = objShell.CreateShortcut(objShell.SpecialFolders("Desktop") & "\abc.lnk")
objShell.Run objShortcut.TargetPath, 1, False
MsgBox "Shortcut opened"