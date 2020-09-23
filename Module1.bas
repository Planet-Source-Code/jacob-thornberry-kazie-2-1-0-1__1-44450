Attribute VB_Name = "Module1"
Declare Function SendPauseCommands Lib "ptmd5.dll" () As Integer
Declare Function SendResumeCommands Lib "ptmd5.dll" () As Integer
Declare Function SendMoreCommands Lib "ptmd5.dll" () As Integer

'This module allows the program, without a title bar, to still be moved
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hWnd, &HA1, 2, 0&)
End Sub

