Attribute VB_Name = "modAtomicClockAndSettings"
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Const ERR_UNERR = "Unexpected error"
Public Const ERR_INACC = "Error - server time inaccurate"

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Public ACDate           As Date
Public Offset           As Long
Public DateExists       As Boolean
Public DontSync         As Boolean
Public LastTicks        As Long
Public LastTicks2       As Long
Public Message          As String
Public bUsed            As Boolean
Public msADV            As Integer
Public ErrCode          As Long

Public UseChime         As Boolean
Public ShowSecondHand   As Boolean
Public RunWin           As Boolean
Public OnTop            As Boolean

Public txtLabel4        As String
Public txtLabel5        As String
Public intCheck1        As Integer
Public intCheck2        As Integer
Public intCheck3        As Integer
Public intCheck4        As Integer

Public Function AppPathContainsSpaces() As Boolean
AppPathContainsSpaces = InStr(1, App.Path & "\" & App.EXEName & ".exe", " ")
End Function

Public Function AppLocation() As String
Dim Q As String
Q = IIf(AppPathContainsSpaces, """", "")
AppLocation = Q & App.Path & "\" & App.EXEName & ".exe" & Q & " 1"
End Function

Public Sub ProcessMsg(MyMessage As String)
On Error Resume Next
If MyMessage = "" Then MyMessage = "Clock not synchronized yet"
Message = MyMessage
If frmAtomicClock.Visible Then frmAtomicClock.Label1 = Message
End Sub

Public Sub ProcessDt(MyMessage As String)
On Error Resume Next
If IsDate(MyMessage) Then
    If ACDate <> CDate(MyMessage) Then ACDate = CDate(MyMessage)
    If txtLabel5 <> CStr(ACDate) Then txtLabel5 = CStr(ACDate)
    If frmAtomicClock.Visible Then frmAtomicClock.Label5 = txtLabel5
    If Not DateExists Then DateExists = True
Else
    If txtLabel5 <> "N/A" Then txtLabel5 = "N/A"
    If frmAtomicClock.Visible Then frmAtomicClock.Label5 = txtLabel5
    If DateExists Then DateExists = False
End If
End Sub

Public Sub SyncTime()
If Not bUsed Then
    bUsed = True
    ProcessMsg "Connecting to time server; please be patient"
    LastTicks = GetTickCount
    DoEvents
    frmMain.WebSock1.ConnectTo "time.nist.gov", 13
    LastTicks2 = GetTickCount
End If
End Sub

Public Function GetTimeOffSet() As Single
GetTimeOffSet = ((GetTickCount - LastTicks) / 1000)
End Function

Public Function SetWinPos(ByVal hwnd As Long, Optional OnTop As Boolean = True) As Long
SetWinPos = SetWindowPos(hwnd, IIf(OnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, FLAGS)
End Function
