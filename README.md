# Screenshot Index Updater VBScript
A simple VBScript to update the Screenshot Index value in the Windows Registry.

![MainWindow](https://user-images.githubusercontent.com/54836559/215480076-062d6728-79c7-4c68-8b6b-fc802e59ecf7.png)

![Dialog](https://user-images.githubusercontent.com/54836559/215480088-6a620471-c8b0-4ec4-9f7d-740c8ab53579.png)


## What is Screenshot Index
The Screenshot Index is a value stored in the Windows Registry that is used to keep track of the next number to use when naming screenshots taken using the `Print Screen`. This value is stored in the registry key `HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\ScreenshotIndex`.

## Usage
1. [Download](https://github.com/kame404/Screenshot-Index-Updater/archive/refs/heads/main.zip) the `UpdateScreenshotIndex.vbs` file.
1. Double-click the file to run it.
1. Enter the new value for the Screenshot Index.
1. The script will update the Screenshot Index value in the Windows Registry.

## Note
This script is designed to run on Windows operating systems and requires administrative privileges to modify the Windows Registry.

## Script
```vb
Option Explicit

Dim regKey, value, newValue
Const HKEY_CURRENT_USER = &H80000001

Set regKey = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
regKey.GetDWORDValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer", "ScreenshotIndex", value

newValue = InputBox("Current ScreenshotIndex value is " & value & ". Please enter the new value:", "Update ScreenshotIndex", value)

If Not IsNull(newValue) And newValue <> "" Then
  If IsNumeric(newValue) Then
    If CLng(newValue) >= 0 And CLng(newValue) <= 9999999 Then
      regKey.SetDWORDValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer", "ScreenshotIndex", CLng(newValue)
      MsgBox "ScreenshotIndex value has been updated from " & value & " to " & newValue & "."
    Else
      MsgBox "Invalid input. ScreenshotIndex value must be a number."
    End If
  Else
    MsgBox "Invalid input. ScreenshotIndex value must be a number."
  End If
Else
  MsgBox "No update made."
End If
```

## License
This project is licensed under the 0BSD License.
