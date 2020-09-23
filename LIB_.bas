Attribute VB_Name = "LIB_"
Option Explicit

Public blnAlarm As Boolean

Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hWnd, &HA1, 2, 0&)
End Sub

