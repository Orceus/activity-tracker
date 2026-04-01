Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

' Get the directory where this VBS script is located
scriptDir = FSO.GetParentFolderName(WScript.ScriptFullName)

' Path to the Python tracker (avoid launcher entirely)
trackerPath = scriptDir & "\activity_tracker.py"

' Check if Python tracker exists
If FSO.FileExists(trackerPath) Then
    ' Use pythonw.exe to run completely hidden (no console window)
    ' Window style 0 = completely hidden
    ' Wait parameter False = don't wait for completion
    WshShell.Run "pythonw.exe """ & trackerPath & """", 0, False
    WScript.Quit 0
Else
    ' Fall back to compiled version if Python version not found
    exePath = scriptDir & "\keytrk_tracker.exe"
    If FSO.FileExists(exePath) Then
        ' Run compiled version completely hidden
        WshShell.Run """" & exePath & """", 0, False
        WScript.Quit 0
    Else
        ' Silent fail - no error dialogs
        WScript.Quit 1
    End If
End If