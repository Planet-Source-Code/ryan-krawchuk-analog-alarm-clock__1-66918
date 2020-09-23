Attribute VB_Name = "modPublicFunctions"
'Option Explicit
'
'Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal wndrpcPrev As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'
'Public Const GWL_WNDPROC = (-4)
'
'Public Declare Function GetTickCount Lib "kernel32" () As Long
'
'Public OldWndProc                   As Long
'
'Private UC                          As New Collection
'Private MaxUCCount                  As Integer
'
'Public CryptionObject               As Object
'Public ICanUseCryptionObject        As Boolean
'Public IShouldUseCryptionObject()   As Boolean
'Public CryptionKey()                As String
'
'Public WinsockStates(9)             As String
'Public CurrentState()               As Integer
'
'Public m_lngSocks()                 As Long
'Public m_intSocketAsync()           As Integer
'Public m_intMaxSockCount            As Integer
'Public m_intConnectionsAlert        As Integer
'
'Public Function GetIndexFromsID(SocketID As Long) As Integer
'Dim X As Integer
'
'For X = 1 To m_intMaxSockCount
'    If m_lngSocks(X) = SocketID Then
'        GetIndexFromsID = X
'        Exit Function
'    End If
'Next
'
'GetIndexFromsID = -1
'End Function
'
'Public Function WaitJustOneSecond(Optional WaitTime As Single = 1) As Boolean
'Dim sTimer As Variant
'
'sTimer = Timer
'Do Until Timer > sTimer + WaitTime
'    DoEvents
'Loop
'
'WaitJustOneSecond = True
'End Function
'
'Public Function SetControlHost(ByVal ControlInstance As WebSock) As String
'Dim objWebSock  As WebSock
'Dim NewKey      As String
'
'NewKey = "a" & UC.Count + 1
'
'Set objWebSock = ControlInstance
'UC.Add objWebSock, NewKey
'
'If UC.Count > MaxUCCount Then MaxUCCount = UC.Count
'
'Set objWebSock = Nothing
'Set ControlInstance = Nothing
'
'SetControlHost = NewKey
'End Function
'
'Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'On Error GoTo ErrHandle
'
'If lParam = 0 Then Exit Function
'
'If uMsg > 4025 And uMsg < 4026 + MaxUCCount Then
'    Dim WSAEvent As Long
'    Dim WSAError As Long
''    Dim TempUC As WebSock
'
''    Set TempUC = UC.Item("a" & uMsg - 4025)
'
'    WSAEvent = WSAGetSelectEvent(lParam)
'    WSAError = WSAGetAsyncError(lParam)
'
'    Select Case WSAEvent
'        Case FD_ACCEPT: TempUC.RaiseConnectionRequest wParam
'        Case FD_READ: ReceiveDataNew wParam, "a" & uMsg - 4025
'        Case FD_CONNECT: TempUC.RaiseConnected wParam
'        Case FD_CLOSE: TempUC.RaisePeerClosing wParam
'        Case FD_WRITE
'        Case FD_OOB
'    End Select
'Else
'    WindowProc = CallWindowProc(OldWndProc, hWnd, uMsg, wParam, ByVal lParam)
'End If
'
'Set TempUC = Nothing
'
'ErrHandle:
'End Function
'
'Private Function ReceiveDataNew(SocketID As Long, UCKey As String)
'Dim RecvBuffer  As String
'Dim fixstr      As String * 1024
'Dim RetByteErr  As Integer
'
'fixstr = ""
'RecvBuffer = ""
'
'RetByteErr = recv(SocketID, fixstr, 1024, 0)
'
'If RetByteErr < 0 Then
'    Exit Function
'ElseIf RetByteErr = 0 Then
'    Exit Function
'Else
'    RecvBuffer = left$(fixstr, RetByteErr)
'End If
'
'If RecvBuffer <> "" Then
'    Dim TempUC As WebSock
'    Set TempUC = UC.Item(UCKey)
'
'    If ICanUseCryptionObject And IShouldUseCryptionObject(GetIndexFromsID(SocketID)) Then RecvBuffer = CryptionObject.Decrypt(RecvBuffer, CryptionKey(GetIndexFromsID(SocketID)))
'    TempUC.RaiseDataArrival SocketID, RecvBuffer
'    Set TempUC = Nothing
'End If
'End Function
'
'Public Sub CleanUp(UCKey As String)
'On Error Resume Next
'UC.Remove UCKey
'End Sub
'
'Public Sub CleanUpAll()
'On Error Resume Next
'Dim X As Integer
'
'For X = UC.Count To 0 Step -1
'    UC.Remove X
'Next
'End Sub
'
'Public Function ResolveIPtoNBO(IP As String) As Long
'Dim NBO As Long
'NBO = inet_addr(IP)
'
'If NBO = -1 Then NBO = GetHostByNameAlias(IP)
'ResolveIPtoNBO = NBO
'End Function
