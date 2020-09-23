VERSION 5.00
Begin VB.UserControl WebSock 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1140
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   810
   ScaleWidth      =   1140
   ToolboxBitmap   =   "WebSock.ctx":0000
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "WebSock.ctx":0312
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "WebSock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private MyUCKey         As String
Private UserEnvironment As Boolean

Public Event ConnectionRequest(ByVal FromListeningSocket As Long)
Public Event Connected(ByVal SocketID As Long)
Public Event DataArrival(ByVal SocketID As Long, sData As String)
Public Event Error(ByVal SocketID As Long, ByVal number As Integer, Description As String)
Public Event ConnectionsAlert(ByVal SocketID As Long)
Public Event PeerClosing(ByVal SocketID As Long)
Public Event SendComplete(ByVal SocketID As Long)
Public Event StateChanged(ByVal SocketID As Long)

Public Property Get State(SocketID As Long) As Integer
Dim X As Integer
X = GetIndexFromsID(SocketID)

If X = -1 Then
    State = -1
Else
    State = CurrentState(X)
End If
End Property

Public Property Let MaxConnectionsAlert(v_intMaxConnectionsAlert As Integer)
m_intConnectionsAlert = v_intMaxConnectionsAlert
End Property

Public Property Get MaxConnectionsAlert() As Integer
MaxConnectionsAlert = m_intConnectionsAlert
End Property

Public Function AcceptConnection(FromListeningSocket As Long) As Long
Dim RC              As Long
Dim i               As Integer
Dim ReadSockBuffer  As sockaddr
      
i = GetNextSocksIndex
m_lngSocks(i) = accept(FromListeningSocket, ReadSockBuffer, Len(ReadSockBuffer))
  
If m_lngSocks(i) = -1 Then Exit Function

RC = WSAAsyncSelect(m_lngSocks(i), hwnd, ByVal (4025 + Int(Right$(MyUCKey, Len(MyUCKey) - 1))), ByVal FD_READ Or FD_CLOSE Or FD_WRITE Or FD_OOB Or FD_CONNECT)
If RC = -1 Then Exit Function
m_intSocketAsync(i) = RC

RaiseEvent Connected(m_lngSocks(i))
AcceptConnection = m_lngSocks(i)
End Function

Public Function ConnectTo(IP As String, OnPort As Integer) As Long
Dim RC              As Long
Dim SocketBuffer    As sockaddr
Dim lngSocket       As Long
Dim i               As Integer
Dim lngResolvedIP   As Long

lngResolvedIP = ResolveIPtoNBO(IP)
If lngResolvedIP = -1 Then Exit Function

OldWndProc = SetWindowLong(UserControl.hwnd, GWL_WNDPROC, AddressOf WindowProc)

lngSocket = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
If lngSocket = -1 Then Exit Function

SocketBuffer.sin_family = AF_INET
SocketBuffer.sin_port = htons(OnPort)
SocketBuffer.sin_addr = lngResolvedIP
SocketBuffer.sin_zero = String$(8, 0)

RC = Connect(lngSocket, SocketBuffer, Len(SocketBuffer))
If RC <> 0 Then
    closesocket CInt(lngSocket)
    Exit Function
End If

RC = WSAAsyncSelect(lngSocket, hwnd, ByVal (4025 + Int(Right$(MyUCKey, Len(MyUCKey) - 1))), ByVal FD_READ Or FD_CLOSE Or FD_WRITE Or FD_OOB Or FD_CONNECT)
If RC <> 0 Then
    closesocket CInt(lngSocket)
    lngSocket = -1
    Exit Function
End If

i = GetNextSocksIndex
m_lngSocks(i) = lngSocket
m_intSocketAsync(i) = RC

ConnectTo = m_lngSocks(i)
End Function

Public Function Disconnect(SocketID As Long) As Boolean
On Error Resume Next
Dim X   As Integer
Dim RC  As Integer

closesocket CInt(SocketID)
X = GetIndexFromsID(SocketID)

m_lngSocks(X) = -1
Disconnect = True
End Function

Public Function LocalIP(SocketID As Long) As String
Dim X   As String
Dim Y   As Integer
    
If Not IDExists(SocketID) Then
    LocalIP = "0.0.0.0"
    Exit Function
End If

X = GetSockAddress(Int(SocketID))

Y = InStr(1, X, ":")
X = Left$(X, Y - 1)

LocalIP = X
End Function

Public Function LocalPort(SocketID As Long) As Variant
Dim X   As String
Dim Y   As Integer

If Not IDExists(SocketID) Then
    LocalPort = 0
    Exit Function
End If

X = GetSockAddress(Int(SocketID))

Y = InStr(1, X, ":")
X = Right$(X, Len(X) - Y)

LocalPort = CLng(X)
End Function

Public Function ListenNow(OnPort As Long) As Long
Dim RC              As Long
Dim SocketBuffer    As sockaddr
Dim lngSocket       As Long
Dim i               As Integer
  
OldWndProc = SetWindowLong(UserControl.hwnd, GWL_WNDPROC, AddressOf WindowProc)

lngSocket = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
If lngSocket = -1 Then Exit Function

SocketBuffer.sin_family = AF_INET
SocketBuffer.sin_port = htons(OnPort)
SocketBuffer.sin_addr = 0
SocketBuffer.sin_zero = String$(8, 0)
    
RC = bind(lngSocket, SocketBuffer, 16)
  
If RC <> 0 Then
    closesocket CInt(lngSocket)
    lngSocket = -1
    Exit Function
End If
    
RC = listen(ByVal lngSocket, ByVal 5)
If RC <> 0 Then
    closesocket CInt(lngSocket)
    lngSocket = -1
    Exit Function
End If
  
RC = WSAAsyncSelect(lngSocket, hwnd, ByVal (4025 + Int(Right$(MyUCKey, Len(MyUCKey) - 1))), ByVal FD_CONNECT Or FD_ACCEPT)
If RC <> 0 Then
    closesocket CInt(lngSocket)
    lngSocket = -1
    Exit Function
End If

i = GetNextSocksIndex
m_lngSocks(i) = lngSocket
m_intSocketAsync(i) = RC
  
ListenNow = lngSocket
End Function

Public Function ReleaseInstance()
CleanUp MyUCKey
End Function

Public Function RemoteIP(SocketID As Long) As String
Dim X   As String
Dim Y   As Integer

If Not IDExists(SocketID) Then
    RemoteIP = "0.0.0.0"
    Exit Function
End If

X = GetPeerAddress(Int(SocketID))

Y = InStr(1, X, ":")
X = Left$(X, Y - 1)

RemoteIP = X
End Function

Public Function RemotePort(SocketID As Long) As Long
Dim X   As String
Dim Y   As Integer
  
If Not IDExists(SocketID) Then
    RemotePort = 0
    Exit Function
End If
    
X = GetPeerAddress(Int(SocketID))

Y = InStr(1, X, ":")
X = Right$(X, Len(X) - Y)

RemotePort = CLng(X)
End Function

Public Function SendDataTo(DataToSend As String, Optional SocketID As Long = 0)
Dim DummyDataToSend As String

If SocketID = 0 Then
    Dim X As Integer
    
    For X = 1 To m_intMaxSockCount
        If m_lngSocks(X) <> -1 Then
            DummyDataToSend = DataToSend
            If ICanUseCryptionObject And IShouldUseCryptionObject(X) Then DummyDataToSend = CryptionObject.Encrypt(DataToSend, CryptionKey(X))
            SendToSock DummyDataToSend, m_lngSocks(X)
        End If
    Next
Else
    If Not IDExists(SocketID) Then Exit Function
    
    DummyDataToSend = DataToSend
    If ICanUseCryptionObject And IShouldUseCryptionObject(GetIndexFromsID(SocketID)) Then DummyDataToSend = CryptionObject.Encrypt(DataToSend, CryptionKey(GetIndexFromsID(SocketID)))
    SendToSock DummyDataToSend, SocketID
End If
End Function

Public Function SetCryptionKey(sKey As String, SocketID As Long) As Boolean
If Not IDExists(SocketID) Then
    SetCryptionKey = False
    Exit Function
End If

SetCryptionKey = True
CryptionKey(GetIndexFromsID(SocketID)) = sKey
End Function

Public Function SetCryptionObject(CryptionObj As Object) As Integer
On Error GoTo Error_Handle

Dim TestString              As String
Dim EncryptReturnString     As String
Dim DecryptReturnString1    As String
Dim DecryptReturnString2    As String
Dim X                       As Integer

TestString = vbCrLf

For X = 32 To 126
  TestString = TestString & Chr$(X)
Next

EncryptReturnString = CryptionObj.Encrypt(TestString, "TestEncryptionString")
DecryptReturnString1 = CryptionObj.Decrypt(EncryptReturnString, "TestEncryptionString")

EncryptReturnString = CryptionObj.Encrypt(TestString, "a")
DecryptReturnString2 = CryptionObj.Decrypt(EncryptReturnString, "a")
  
EncryptReturnString = CryptionObj.Encrypt(TestString, "")
DecryptReturnString2 = CryptionObj.Decrypt(EncryptReturnString, "")

If TestString = DecryptReturnString1 And TestString = DecryptReturnString2 Then
    SetCryptionObject = 0
    ICanUseCryptionObject = True
    Set CryptionObject = CryptionObj
Else
    SetCryptionObject = 1
End If
Exit Function

Error_Handle:
    SetCryptionObject = 2
    Exit Function
    Resume Next
End Function

Public Function StateDescription(Optional SocketID As Long = 0, Optional v_intSelState As Integer = -1) As String
Dim X As Integer
  
If v_intSelState < 0 Or v_intSelState > 9 Then
    X = GetIndexFromsID(SocketID)
    If X = -1 Then
        StateDescription = "Socket Does Not Exist"
    Else
        StateDescription = WinsockStates(CurrentState(X))
    End If
Else
    StateDescription = WinsockStates(v_intSelState)
End If
End Function

Public Function UseCryption(SocketID As Long, Optional OnOff As Integer = -1) As Integer
If Not IDExists(SocketID) Then
    UseCryption = 0
    Exit Function
End If
  
If OnOff = 0 Then
    IShouldUseCryptionObject(GetIndexFromsID(SocketID)) = False
    UseCryption = 1
ElseIf OnOff = 1 Then
    If ICanUseCryptionObject Then
        IShouldUseCryptionObject(GetIndexFromsID(SocketID)) = True
        UseCryption = 2
    Else
        IShouldUseCryptionObject(GetIndexFromsID(SocketID)) = False
        UseCryption = 3
    End If
Else
    If IShouldUseCryptionObject(GetIndexFromsID(SocketID)) Then
        UseCryption = 4
    Else
        UseCryption = 5
    End If
End If
End Function

Public Sub RaiseConnected(ByVal SocketID As Long)
RaiseEvent Connected(SocketID)
End Sub

Public Sub RaiseConnectionRequest(ByVal SocketID As Long)
RaiseEvent ConnectionRequest(SocketID)
End Sub

Public Sub RaiseDataArrival(ByVal SocketID As Long, sData As String)
RaiseEvent DataArrival(SocketID, sData)
End Sub

Public Sub RaiseError(ByVal SocketID As Long, ByVal number As Integer, Description As String)
RaiseEvent Error(SocketID, number, Description)
End Sub

Public Sub RaisePeerClosing(ByVal SocketID As Long)
Disconnect SocketID
RaiseEvent PeerClosing(SocketID)
End Sub

Public Sub RaiseSendComplete(ByVal SocketID As Long)
RaiseEvent SendComplete(SocketID)
End Sub

Public Sub RaiseStateChanged(ByVal SocketID As Long)
RaiseEvent StateChanged(SocketID)
End Sub

Private Function GetNextSocksIndex() As Integer
Dim X As Integer
  
For X = 1 To m_intMaxSockCount
    If m_lngSocks(X) = -1 Then
        GetNextSocksIndex = X
        Exit Function
    End If
Next
m_intMaxSockCount = m_intMaxSockCount + 1
  
ReDim Preserve m_lngSocks(m_intMaxSockCount)
ReDim Preserve m_intSocketAsync(m_intMaxSockCount)
ReDim Preserve CurrentState(m_intMaxSockCount)
ReDim Preserve IShouldUseCryptionObject(m_intMaxSockCount)
ReDim Preserve CryptionKey(m_intMaxSockCount)

m_lngSocks(m_intMaxSockCount) = -1
m_intSocketAsync(m_intMaxSockCount) = -1
CurrentState(m_intMaxSockCount) = -1
IShouldUseCryptionObject(m_intMaxSockCount) = False
CryptionKey(m_intMaxSockCount) = ""

GetNextSocksIndex = m_intMaxSockCount
End Function

Private Function IDExists(SocketID As Long) As Boolean
Dim X As Integer
  
For X = 1 To m_intMaxSockCount
    If m_lngSocks(X) = SocketID Then
        IDExists = True
        Exit Function
    End If
Next
  
IDExists = False
End Function

Private Sub SendToSock(msg As String, SocketID As Long)
On Error Resume Next
Dim RC As Long

RC = SendData(SocketID, msg)
If RC <> -1 Then RaiseEvent SendComplete(SocketID)
End Sub

Private Sub Initialize()
If Not UserEnvironment Then Exit Sub
On Error Resume Next

If Not WinsockStartedUp Then
    Dim RC          As Long
    Dim StartupData As WSADataType
    
    RC = WSAStartup(&H101, StartupData)
    
    If RC = -1 Then Exit Sub
    
    WinsockStartedUp = True
    
    WinsockStates(0) = "Closed"
    WinsockStates(1) = "Open"
    WinsockStates(2) = "Listening"
    WinsockStates(3) = "Connection Pending"
    WinsockStates(4) = "Resolving Host"
    WinsockStates(5) = "Host Resolved"
    WinsockStates(6) = "Connecting"
    WinsockStates(7) = "Connected"
    WinsockStates(8) = "Peer Is Closing The Connection"
    WinsockStates(9) = "Error On Socket"
      
    m_intMaxSockCount = 1
    
    ReDim Preserve m_lngSocks(1)
    ReDim Preserve m_intSocketAsync(1)
    ReDim Preserve CurrentState(1)
    ReDim Preserve IShouldUseCryptionObject(1)
    ReDim Preserve CryptionKey(1)
    
    m_lngSocks(1) = -1
    m_intSocketAsync(1) = -1
    CurrentState(1) = -1
    IShouldUseCryptionObject(1) = False
    
    ICanUseCryptionObject = False
    
    CryptionKey(1) = ""
        
End If

MyUCKey = SetControlHost(Me)
  
m_intConnectionsAlert = 50
End Sub

Private Sub UserControl_Terminate()
If Not UserEnvironment Then Exit Sub
On Error Resume Next

Dim RC  As Long
Dim X   As Integer

For X = 1 To m_intMaxSockCount
  closesocket CInt(m_lngSocks(X))
  RC = WSACancelAsyncRequest(m_intSocketAsync(X))
Next

RC = WSCleanUp()
End Sub

Private Sub UserControl_Resize()
UserControl.Width = 480
UserControl.Height = 480
End Sub

Private Sub UserControl_InitProperties()
UserEnvironment = Ambient.UserMode
Initialize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
UserEnvironment = Ambient.UserMode
Initialize
End Sub
