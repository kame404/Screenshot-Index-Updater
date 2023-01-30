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