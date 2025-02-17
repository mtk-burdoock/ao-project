VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Private Declare Function api_socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Private Declare Function api_GlobalLock Lib "kernel32" Alias "GlobalLock" (ByVal hMem As Long) As Long
Private Declare Function api_GlobalUnlock Lib "kernel32" Alias "GlobalUnlock" (ByVal hMem As Long) As Long
Private Declare Function api_htons Lib "ws2_32.dll" Alias "htons" (ByVal hostshort As Integer) As Integer
Private Declare Function api_ntohs Lib "ws2_32.dll" Alias "ntohs" (ByVal netshort As Integer) As Integer
Private Declare Function api_connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef Name As sockaddr_in, ByVal namelen As Long) As Long
Private Declare Function api_gethostname Lib "ws2_32.dll" Alias "gethostname" (ByVal host_name As String, ByVal namelen As Long) As Long
Private Declare Function api_gethostbyname Lib "ws2_32.dll" Alias "gethostbyname" (ByVal host_name As String) As Long
Private Declare Function api_bind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, ByRef Name As sockaddr_in, ByVal namelen As Long) As Long
Private Declare Function api_getsockname Lib "ws2_32.dll" Alias "getsockname" (ByVal s As Long, ByRef Name As sockaddr_in, ByRef namelen As Long) As Long
Private Declare Function api_getpeername Lib "ws2_32.dll" Alias "getpeername" (ByVal s As Long, ByRef Name As sockaddr_in, ByRef namelen As Long) As Long
Private Declare Function api_inet_addr Lib "ws2_32.dll" Alias "inet_addr" (ByVal cp As String) As Long
Private Declare Function api_send Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function api_sendto Lib "ws2_32.dll" Alias "sendto" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long, ByRef toaddr As sockaddr_in, ByVal tolen As Long) As Long
Private Declare Function api_getsockopt Lib "ws2_32.dll" Alias "getsockopt" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Private Declare Function api_setsockopt Lib "ws2_32.dll" Alias "setsockopt" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function api_recv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function api_recvfrom Lib "ws2_32.dll" Alias "recvfrom" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long, ByRef From As sockaddr_in, ByRef fromlen As Long) As Long
Private Declare Function api_WSACancelAsyncRequest Lib "ws2_32.dll" Alias "WSACancelAsyncRequest" (ByVal hAsyncTaskHandle As Long) As Long
Private Declare Function api_listen Lib "ws2_32.dll" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function api_accept Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, ByRef Addr As sockaddr_in, ByRef addrlen As Long) As Long
Private Declare Function api_inet_ntoa Lib "ws2_32.dll" Alias "inet_ntoa" (ByVal inn As Long) As Long
Private Declare Function api_ioctlsocket Lib "ws2_32.dll" Alias "ioctlsocket" (ByVal s As Long, ByVal Cmd As Long, ByRef argp As Long) As Long
Private Declare Function api_closesocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long

Public Enum SockState
    sckClosed = 0
    sckOpen
    sckListening
    sckConnectionPending
    sckResolvingHost
    sckHostResolved
    sckConnecting
    sckConnected
    sckClosing
    sckError
End Enum
 
Private Const SOMAXCONN As Long = 5
 
Public Enum ProtocolConstants
    sckTCPProtocol = 0
    sckUDPProtocol = 1
End Enum
 
Private Const MSG_PEEK  As Long = &H2
 
Public Event CloseSck()
Public Event Connect()
Public Event ConnectionRequest(ByVal requestID As Long)
Public Event DataArrival(ByVal bytesTotal As Long)
Public Event Error(ByVal number As Integer, Description As String, ByVal sCode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Public Event SendComplete()
Public Event SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
 
Private m_lngSocketHandle       As Long
Private m_enmState              As SockState
Private m_strTag                As String
Private m_strRemoteHost         As String
Private m_lngRemotePort         As Long
Private m_strRemoteHostIP       As String
Private m_lngLocalPort          As Long
Private m_lngLocalPortBind      As Long
Private m_strLocalIP            As String
Private m_enmProtocol           As ProtocolConstants
Private m_lngMemoryPointer  As Long
Private m_lngMemoryHandle   As Long
Private m_lngSendBufferLen  As Long
Private m_lngRecvBufferLen  As Long
Private m_strSendBuffer As String
Private m_strRecvBuffer As String
Private m_blnAcceptClass As Boolean
Private m_colWaitingResolutions As Collection
 
Public Sub WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    Select Case uMsg
        Case RESOLVE_MESSAGE
            PostResolution wParam, HiWord(lParam)
    
        Case SOCKET_MESSAGE
            PostSocket LoWord(lParam), HiWord(lParam)
    End Select
End Sub
 
Private Sub Class_Initialize()
    m_lngSocketHandle = INVALID_SOCKET
    Set m_colWaitingResolutions = New Collection
    modSocket.InitiateProcesses
End Sub
 
Private Sub Class_Terminate()
    CleanResolutionSystem
    If Not m_blnAcceptClass Then DestroySocket
    modSocket.FinalizeProcesses
    Set m_colWaitingResolutions = Nothing
End Sub
  
Public Property Get RemotePort() As Long
    RemotePort = m_lngRemotePort
End Property
 
Public Property Let RemotePort(ByVal lngPort As Long)
    If m_enmProtocol = sckTCPProtocol And m_enmState <> sckClosed Then
        Exit Property
    End If
    If lngPort < 0 Or lngPort > 65535 Then
        Exit Property
    Else
        m_lngRemotePort = lngPort
    End If
End Property
 
Public Property Get RemoteHost() As String
    RemoteHost = m_strRemoteHost
End Property
 
Public Property Let RemoteHost(ByVal strHost As String)
    m_strRemoteHost = strHost
End Property
 
Public Property Get RemoteHostIP() As String
    RemoteHostIP = m_strRemoteHostIP
End Property
 
Public Property Get LocalPort() As Long
    If m_lngLocalPortBind = 0 Then
        LocalPort = m_lngLocalPort
    Else
        LocalPort = m_lngLocalPortBind
    End If
End Property
 
Public Property Let LocalPort(ByVal lngPort As Long)
    If m_enmState <> sckClosed Then
        Exit Property
    End If
    If lngPort < 0 Or lngPort > 65535 Then
        Exit Property
    Else
        m_lngLocalPort = lngPort
    End If
End Property
 
Public Property Get State() As SockState
    State = m_enmState
End Property
 
Public Property Get LocalHostName() As String
    LocalHostName = GetLocalHostName
End Property
 
Public Property Get LocalIP() As String
    If m_enmState = sckConnected Then
        LocalIP = m_strLocalIP
    Else
        LocalIP = GetLocalIP
    End If
End Property
 
Public Property Get BytesReceived() As Long
    If m_enmProtocol = sckTCPProtocol Then
        BytesReceived = Len(m_strRecvBuffer)
    Else
        BytesReceived = GetBufferLenUDP
    End If
End Property
 
Public Property Get SocketHandle() As Long
    SocketHandle = m_lngSocketHandle
End Property
 
Public Property Get Tag() As String
    Tag = m_strTag
End Property
 
Public Property Let Tag(ByVal strTag As String)
    m_strTag = strTag
End Property
 
Public Property Get Protocol() As ProtocolConstants
    Protocol = m_enmProtocol
End Property
 
Public Property Let Protocol(ByVal enmProtocol As ProtocolConstants)
    If m_enmState <> sckClosed Then
        Exit Property
    Else
        m_enmProtocol = enmProtocol
    End If
End Property
 
Private Sub DestroySocket()
    If Not m_lngSocketHandle = INVALID_SOCKET Then
        Dim lngResult As Long
        lngResult = api_closesocket(m_lngSocketHandle)
        If lngResult = Socket_Error Then
            m_enmState = sckError
            Dim lngErrorCode As Long
            lngErrorCode = Err.LastDllError
            Exit Sub
        Else
            modSocket.UnregisterSocket m_lngSocketHandle
            m_lngSocketHandle = INVALID_SOCKET
        End If
    End If
End Sub
 
Public Sub CloseSck()
    If m_lngSocketHandle = INVALID_SOCKET Then Exit Sub
    m_enmState = sckClosing
    CleanResolutionSystem
    DestroySocket
    m_lngLocalPortBind = 0
    m_strRemoteHostIP = ""
    m_strRecvBuffer = ""
    m_strSendBuffer = ""
    m_lngSendBufferLen = 0
    m_lngRecvBufferLen = 0
    m_enmState = sckClosed
End Sub
 
Private Function SocketExists() As Boolean
    SocketExists = True
    Dim lngResult    As Long
    Dim lngErrorCode As Long
    If m_lngSocketHandle = INVALID_SOCKET Then
        If m_enmProtocol = sckTCPProtocol Then
            lngResult = api_socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
        Else
            lngResult = api_socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP)
        End If
        If lngResult = INVALID_SOCKET Then
            m_enmState = sckError
            SocketExists = False
            lngErrorCode = Err.LastDllError
            Dim blnCancelDisplay As Boolean
            blnCancelDisplay = True
            RaiseEvent Error(lngErrorCode, GetErrorDescription(lngErrorCode), 0, "Socket.SocketExists", "", 0, blnCancelDisplay)
            If blnCancelDisplay = False Then MsgBox GetErrorDescription(lngErrorCode), vbOKOnly, "Socket.SocketExists"
        Else
            m_lngSocketHandle = lngResult
            ProcessOptions
            SocketExists = modSocket.RegisterSocket(m_lngSocketHandle, ObjPtr(Me), True)
        End If
    End If
End Function
 
Public Sub Connect(Optional RemoteHost As String, Optional RemotePort As Integer)
    If m_enmState <> sckClosed Then
        Exit Sub
    End If
    If Not IsMissing(RemoteHost) Then
        m_strRemoteHost = CStr(RemoteHost)
    End If
    If m_strRemoteHost = vbNullString Then
        m_strRemoteHost = ""
    End If
    If Not IsMissing(RemotePort) Then
        If IsNumeric(RemotePort) Then
            If CLng(RemotePort) > 65535 Or CLng(RemotePort) < 1 Then
                Exit Sub
            Else
                m_lngRemotePort = CLng(RemotePort)
            End If
        Else
            Exit Sub
        End If
    End If
    If Not SocketExists Then Exit Sub
    If Not BindInternal Then Exit Sub
    If m_enmProtocol = sckUDPProtocol Then
        m_enmState = sckOpen
        Exit Sub
    End If
    Dim lngAddress As Long
    lngAddress = ResolveIfHostname(m_strRemoteHost)
    If lngAddress <> vbNull Then
        ConnectToIP lngAddress, 0
    End If
End Sub
 
Private Sub PostResolution(ByVal lngAsynHandle As Long, ByVal lngErrorCode As Long)
    m_colWaitingResolutions.Remove "R" & lngAsynHandle
    UnregisterResolution lngAsynHandle
    If m_enmState <> sckResolvingHost Then Exit Sub
    If lngErrorCode = 0 Then
        m_enmState = sckHostResolved
        Dim udtHostent           As HOSTENT
        Dim lngPtrToIP           As Long
        Dim arrIpAddress(1 To 4) As Byte
        Dim lngRemoteHostAddress As Long
        Dim Count                As Integer
        Dim strIpAddress         As String
        api_CopyMemory udtHostent, ByVal m_lngMemoryPointer, LenB(udtHostent)
        api_CopyMemory lngPtrToIP, ByVal udtHostent.hAddrList, 4
        api_CopyMemory arrIpAddress(1), ByVal lngPtrToIP, 4
        api_CopyMemory lngRemoteHostAddress, ByVal lngPtrToIP, 4
        FreeMemory
        For Count = 1 To 4
            strIpAddress = strIpAddress & arrIpAddress(Count) & "."
        Next
        strIpAddress = Left$(strIpAddress, Len(strIpAddress) - 1)
        ConnectToIP lngRemoteHostAddress, 0
    Else
        FreeMemory
        ConnectToIP vbNull, lngErrorCode
    End If
End Sub
 
Private Sub PostSocket(ByVal lngEventID As Long, ByVal lngErrorCode As Long)
    Dim blnCancelDisplay As Boolean
    If lngErrorCode <> 0 Then
        m_enmState = sckError
        blnCancelDisplay = True
        RaiseEvent Error(lngErrorCode, GetErrorDescription(lngErrorCode), 0, "Socket.PostSocket", "", 0, blnCancelDisplay)
        If blnCancelDisplay = False Then MsgBox GetErrorDescription(lngErrorCode), vbOKOnly, "Socket.PostSocket"
        Exit Sub
    End If
    Dim udtSockAddr      As sockaddr_in
    Dim lngResult        As Long
    Dim lngBytesReceived As Long
    Select Case lngEventID
        Case FD_CONNECT
            If m_enmState <> sckConnecting Then
                Exit Sub
            End If
            GetLocalInfo m_lngSocketHandle, m_lngLocalPortBind, m_strLocalIP
            GetRemoteInfo m_lngSocketHandle, m_lngRemotePort, m_strRemoteHostIP, m_strRemoteHost
            m_enmState = sckConnected
            RaiseEvent Connect
 
        Case FD_WRITE
            If m_enmState <> sckConnected Then
                Exit Sub
            End If
            If Len(m_strSendBuffer) > 0 Then
                SendBufferedData
            End If
 
        Case FD_READ
            If m_enmProtocol = sckTCPProtocol Then
                If m_enmState <> sckConnected Then
                    Exit Sub
                End If
                lngBytesReceived = RecvDataToBuffer
                If lngBytesReceived > 0 Then
                    RaiseEvent DataArrival(Len(m_strRecvBuffer))
                End If
            Else
                If m_enmState <> sckOpen Then
                    Exit Sub
                End If
                lngBytesReceived = GetBufferLenUDP
                If lngBytesReceived > 0 Then
                    RaiseEvent DataArrival(lngBytesReceived)
                End If
                EmptyBuffer
            End If
 
        Case FD_ACCEPT
            If m_enmState <> sckListening Then
                Exit Sub
            End If
            lngResult = api_accept(m_lngSocketHandle, udtSockAddr, LenB(udtSockAddr))
            If lngResult = INVALID_SOCKET Then
                lngErrorCode = Err.LastDllError
                m_enmState = sckError
                blnCancelDisplay = True
                RaiseEvent Error(lngErrorCode, GetErrorDescription(lngErrorCode), 0, "Socket.PostSocket", "", 0, blnCancelDisplay)
                If blnCancelDisplay = False Then MsgBox GetErrorDescription(lngErrorCode), vbOKOnly, "Socket.PostSocket"
            Else
                modSocket.RegisterAccept lngResult
                Dim lngTempRP   As Long
                Dim strTempRHIP As String
                Dim strTempRH   As String
                lngTempRP = m_lngRemotePort
                strTempRHIP = m_strRemoteHostIP
                strTempRH = m_strRemoteHost
                m_strRemoteHost = ""
                GetRemoteInfo lngResult, m_lngRemotePort, m_strRemoteHostIP, m_strRemoteHost
                RaiseEvent ConnectionRequest(lngResult)
                If m_enmState = sckListening Then
                    m_lngRemotePort = lngTempRP
                    m_strRemoteHostIP = strTempRHIP
                    m_strRemoteHost = strTempRH
                End If
                If IsAcceptRegistered(lngResult) Then
                    api_closesocket lngResult
                    modSocket.UnregisterSocket lngResult
                    modSocket.UnregisterAccept lngResult
                End If
            End If
    
        Case FD_CLOSE
            If m_enmState <> sckConnected Then
                Exit Sub
            End If
            m_enmState = sckClosing
            RaiseEvent CloseSck
    End Select
End Sub
 
Private Sub ConnectToIP(ByVal lngRemoteHostAddress As Long, ByVal lngErrorCode As Long)
    Dim blnCancelDisplay As Boolean
    If lngErrorCode <> 0 Then
        m_enmState = sckError
        blnCancelDisplay = True
        RaiseEvent Error(lngErrorCode, GetErrorDescription(lngErrorCode), 0, "Socket.ConnectToIP", "", 0, blnCancelDisplay)
        If blnCancelDisplay = False Then MsgBox GetErrorDescription(lngErrorCode), vbOKOnly, "Socket.ConnectToIP"
        Exit Sub
    End If
    m_enmState = sckConnecting
    Dim udtSockAddr As sockaddr_in
    Dim lngResult   As Long
    With udtSockAddr
        .sin_addr = lngRemoteHostAddress
        .sin_family = AF_INET
        .sin_port = api_htons(modSocket.UnsignedToInteger(m_lngRemotePort))
    End With
    lngResult = api_connect(m_lngSocketHandle, udtSockAddr, LenB(udtSockAddr))
    If lngResult = Socket_Error Then
        lngErrorCode = Err.LastDllError
        If lngErrorCode <> WSAEWOULDBLOCK Then
            If lngErrorCode = WSAEADDRNOTAVAIL Then
                Exit Sub
            Else
                m_enmState = sckError
                blnCancelDisplay = True
                RaiseEvent Error(lngErrorCode, GetErrorDescription(lngErrorCode), 0, "Socket.ConnectToIP", "", 0, blnCancelDisplay)
                If blnCancelDisplay = False Then MsgBox GetErrorDescription(lngErrorCode), vbOKOnly, "Socket.ConnectToIP"
            End If
        End If
    End If
End Sub
 
Public Sub Bind(Optional LocalPort As Long, Optional LocalIP As Long)
    If m_enmState <> sckClosed Then
        Exit Sub
    End If
    If BindInternal(LocalPort, LocalIP) Then
        m_enmState = sckOpen
    End If
End Sub
 
Private Function BindInternal(Optional ByVal varLocalPort As Long, Optional ByVal varLocalIP As Long) As Boolean
    If m_enmState = sckOpen Then
        BindInternal = True
        Exit Function
    End If
    Dim lngLocalPortInternal As Long
    Dim strLocalHostInternal As String
    Dim strIP                As String
    Dim lngAddressInternal   As Long
    Dim lngResult            As Long
    Dim lngErrorCode         As Long
    BindInternal = False
    If Not IsMissing(varLocalPort) Then
        If IsNumeric(varLocalPort) Then
            If varLocalPort < 0 Or varLocalPort > 65535 Then
                BindInternal = False
                Exit Function
            Else
                lngLocalPortInternal = CLng(varLocalPort)
            End If
        Else
            BindInternal = False
            Exit Function
        End If
    Else
        lngLocalPortInternal = m_lngLocalPort
    End If
    If Not IsMissing(varLocalIP) Then
        If varLocalIP <> 0 Then
            strLocalHostInternal = CStr(varLocalIP)
        Else
            strLocalHostInternal = 0
        End If
    Else
        strLocalHostInternal = 0
    End If
    lngAddressInternal = ResolveIfHostnameSync(strLocalHostInternal, strIP, lngResult)
    If lngResult <> 0 Then
        Exit Function
    End If
    If Not SocketExists Then Exit Function
    Dim udtSockAddr As sockaddr_in
    With udtSockAddr
        .sin_addr = lngAddressInternal
        .sin_family = AF_INET
        .sin_port = api_htons(modSocket.UnsignedToInteger(lngLocalPortInternal))
    End With
    lngResult = api_bind(m_lngSocketHandle, udtSockAddr, LenB(udtSockAddr))
    If lngResult = Socket_Error Then
        lngErrorCode = Err.LastDllError
        Exit Function
    Else
        If lngLocalPortInternal <> 0 Then
            m_lngLocalPort = lngLocalPortInternal
        Else
            lngResult = GetLocalPort(m_lngSocketHandle)
            If lngResult = Socket_Error Then
                lngErrorCode = Err.LastDllError
                Exit Function
            Else
                m_lngLocalPortBind = lngResult
            End If
        End If
        BindInternal = True
    End If
End Function
 
Private Function AllocateMemory() As Long
    m_lngMemoryHandle = api_GlobalAlloc(GMEM_FIXED, MAXGETHOSTSTRUCT)
    If m_lngMemoryHandle <> 0 Then
        m_lngMemoryPointer = api_GlobalLock(m_lngMemoryHandle)
        If m_lngMemoryPointer <> 0 Then
            api_GlobalUnlock (m_lngMemoryHandle)
            AllocateMemory = m_lngMemoryPointer
        Else
            api_GlobalFree (m_lngMemoryHandle)
            AllocateMemory = m_lngMemoryPointer
        End If
    Else
        AllocateMemory = m_lngMemoryHandle
    End If
End Function
 
Private Sub FreeMemory()
    If m_lngMemoryHandle <> 0 Then
        m_lngMemoryPointer = 0
        api_GlobalFree m_lngMemoryHandle
        m_lngMemoryHandle = 0
    End If
End Sub
 
Private Function GetLocalHostName() As String
    Dim strHostNameBuf As String * LOCAL_HOST_BUFF
    Dim lngResult      As Long
    lngResult = api_gethostname(strHostNameBuf, LOCAL_HOST_BUFF)
    If lngResult = Socket_Error Then
        GetLocalHostName = vbNullString
        Dim lngErrorCode As Long
        lngErrorCode = Err.LastDllError
        Exit Function
    Else
        GetLocalHostName = Left$(strHostNameBuf, InStr(1, strHostNameBuf, vbNullChar) - 1)
    End If
End Function
 
Private Function GetLocalIP() As String
    Dim lngResult            As Long
    Dim lngPtrToIP           As Long
    Dim strLocalHost         As String
    Dim arrIpAddress(1 To 4) As Byte
    Dim Count                As Integer
    Dim udtHostent           As HOSTENT
    Dim strIpAddress         As String
    strLocalHost = GetLocalHostName
    lngResult = api_gethostbyname(strLocalHost)
    If lngResult = 0 Then
        GetLocalIP = vbNullString
        Dim lngErrorCode As Long
        lngErrorCode = Err.LastDllError
        Exit Function
    Else
        api_CopyMemory udtHostent, ByVal lngResult, LenB(udtHostent)
        api_CopyMemory lngPtrToIP, ByVal udtHostent.hAddrList, 4
        api_CopyMemory arrIpAddress(1), ByVal lngPtrToIP, 4
        For Count = 1 To 4
            strIpAddress = strIpAddress & arrIpAddress(Count) & "."
        Next
        strIpAddress = Left$(strIpAddress, Len(strIpAddress) - 1)
        GetLocalIP = strIpAddress
    End If
End Function
 
Private Function ResolveIfHostname(ByVal Host As String) As Long
    Dim lngAddress As Long
    lngAddress = api_inet_addr(Host)
    If lngAddress = INADDR_NONE Then
        ResolveIfHostname = vbNull
        m_enmState = sckResolvingHost
        If AllocateMemory Then
            Dim lngAsynHandle As Long
            lngAsynHandle = modSocket.ResolveHost(Host, m_lngMemoryPointer, ObjPtr(Me))
            If lngAsynHandle = 0 Then
                FreeMemory
                m_enmState = sckError
                Dim lngErrorCode As Long
                lngErrorCode = Err.LastDllError
                Dim blnCancelDisplay As Boolean
                blnCancelDisplay = True
                RaiseEvent Error(lngErrorCode, GetErrorDescription(lngErrorCode), 0, "Socket.ResolveIfHostname", "", 0, blnCancelDisplay)
                If blnCancelDisplay = False Then MsgBox GetErrorDescription(lngErrorCode), vbOKOnly, "Socket.ResolveIfHostname"
            Else
                m_colWaitingResolutions.Add lngAsynHandle, "R" & lngAsynHandle
            End If
        Else
            m_enmState = sckError
            Exit Function
        End If
    Else
        ResolveIfHostname = lngAddress
    End If
End Function
 
Private Function ResolveIfHostnameSync(ByVal Host As String, ByRef strHostIP As String, ByRef lngErrorCode As Long) As Long
    Dim lngPtrToHOSTENT      As Long
    Dim udtHostent           As HOSTENT
    Dim lngAddress           As Long
    Dim lngPtrToIP           As Long
    Dim arrIpAddress(1 To 4) As Byte
    Dim Count                As Integer
    lngAddress = api_inet_addr(Host)
    If lngAddress = INADDR_NONE Then
        lngPtrToHOSTENT = api_gethostbyname(Host)
        If lngPtrToHOSTENT = 0 Then
            lngErrorCode = Err.LastDllError
            strHostIP = vbNullString
            ResolveIfHostnameSync = vbNull
        Else
            api_CopyMemory udtHostent, ByVal lngPtrToHOSTENT, LenB(udtHostent)
            api_CopyMemory lngPtrToIP, ByVal udtHostent.hAddrList, 4
            api_CopyMemory arrIpAddress(1), ByVal lngPtrToIP, 4
            api_CopyMemory lngAddress, ByVal lngPtrToIP, 4
            For Count = 1 To 4
                strHostIP = strHostIP & arrIpAddress(Count) & "."
            Next
            strHostIP = Left$(strHostIP, Len(strHostIP) - 1)
            lngErrorCode = 0
            ResolveIfHostnameSync = lngAddress
        End If
    Else
        lngErrorCode = 0
        strHostIP = Host
        ResolveIfHostnameSync = lngAddress
    End If
End Function
 
Private Function GetLocalPort(ByVal lngSocket As Long) As Long
    Dim udtSockAddr As sockaddr_in
    Dim lngResult   As Long
    lngResult = api_getsockname(lngSocket, udtSockAddr, LenB(udtSockAddr))
    If lngResult = Socket_Error Then
        GetLocalPort = Socket_Error
    Else
        GetLocalPort = modSocket.IntegerToUnsigned(api_ntohs(udtSockAddr.sin_port))
    End If
End Function
 
Public Sub SendData(data As String)
    Dim arrData() As Byte
    If m_enmProtocol = sckTCPProtocol Then
        If m_enmState <> sckConnected Then
            Exit Sub
            Exit Sub
        End If
    Else
        If Not SocketExists Then Exit Sub
        If Not BindInternal Then Exit Sub
        m_enmState = sckOpen
    End If
    Select Case varType(data)
        Case vbString
            Dim strData As String
            strData = CStr(data)
            If Len(strData) = 0 Then Exit Sub
            ReDim arrData(Len(strData) - 1)
            arrData() = StrConv(strData, vbFromUnicode)

        Case vbArray + vbByte
            Dim strArray As String
            strArray = StrConv(data, vbUnicode)
            If Len(strArray) = 0 Then Exit Sub
            arrData() = StrConv(strArray, vbFromUnicode)

        Case vbBoolean
            Dim blnData As Boolean
            blnData = CBool(data)
            ReDim arrData(LenB(blnData) - 1)
            api_CopyMemory arrData(0), blnData, LenB(blnData)

        Case vbByte
            Dim bytData As Byte
            bytData = CByte(data)
            ReDim arrData(LenB(bytData) - 1)
            api_CopyMemory arrData(0), bytData, LenB(bytData)

        Case vbCurrency
            Dim curData As Currency
            curData = CCur(data)
            ReDim arrData(LenB(curData) - 1)
            api_CopyMemory arrData(0), curData, LenB(curData)

        Case vbDate
            Dim datData As Date
            datData = CDate(data)
            ReDim arrData(LenB(datData) - 1)
            api_CopyMemory arrData(0), datData, LenB(datData)

        Case vbDouble
            Dim dblData As Double
            dblData = CDbl(data)
            ReDim arrData(LenB(dblData) - 1)
            api_CopyMemory arrData(0), dblData, LenB(dblData)

        Case vbInteger
            Dim intData As Integer
            intData = CInt(data)
            ReDim arrData(LenB(intData) - 1)
            api_CopyMemory arrData(0), intData, LenB(intData)

        Case vbLong
            Dim lngData As Long
            lngData = CLng(data)
            ReDim arrData(LenB(lngData) - 1)
            api_CopyMemory arrData(0), lngData, LenB(lngData)

        Case vbSingle
            Dim sngData As Single
            sngData = CSng(data)
            ReDim arrData(LenB(sngData) - 1)
            api_CopyMemory arrData(0), sngData, LenB(sngData)

        Case Else
            Exit Sub
    End Select
    If Len(m_strSendBuffer) > 0 Then
        m_strSendBuffer = m_strSendBuffer + StrConv(arrData(), vbUnicode)
        Exit Sub
    Else
        m_strSendBuffer = m_strSendBuffer + StrConv(arrData(), vbUnicode)
    End If
    SendBufferedData
End Sub
 
Private Sub SendBufferedData()
    If m_enmProtocol = sckTCPProtocol Then
        SendBufferedDataTCP
    Else
        SendBufferedDataUDP
    End If
End Sub
 
Private Sub SendBufferedDataUDP()
    Dim lngAddress      As Long
    Dim udtSockAddr     As sockaddr_in
    Dim arrData()       As Byte
    Dim lngBufferLength As Long
    Dim lngResult       As Long
    Dim lngErrorCode    As Long
    Dim strTemp         As String
    lngAddress = ResolveIfHostnameSync(m_strRemoteHost, strTemp, lngErrorCode)
    If lngErrorCode <> 0 Then
        m_strSendBuffer = ""
        If lngErrorCode = WSAEAFNOSUPPORT Then
            Exit Sub
        Else
            Exit Sub
        End If
    End If
    With udtSockAddr
        .sin_addr = lngAddress
        .sin_family = AF_INET
        .sin_port = api_htons(modSocket.UnsignedToInteger(m_lngRemotePort))
    End With
    lngBufferLength = Len(m_strSendBuffer)
    arrData() = StrConv(m_strSendBuffer, vbFromUnicode)
    m_strSendBuffer = ""
    lngResult = api_sendto(m_lngSocketHandle, arrData(0), lngBufferLength, 0&, udtSockAddr, LenB(udtSockAddr))
    If lngResult = Socket_Error Then
        lngErrorCode = Err.LastDllError
        If lngErrorCode = WSAEWOULDBLOCK Then
            Exit Sub
        Else
            m_enmState = sckError
            Dim blnCancelDisplay As Boolean
            blnCancelDisplay = True
            RaiseEvent Error(lngErrorCode, GetErrorDescription(lngErrorCode), 0, "Socket.SendBufferedDataUDP", "", 0, blnCancelDisplay)
            If blnCancelDisplay = False Then MsgBox GetErrorDescription(lngErrorCode), vbOKOnly, "Socket.SendBufferedDataUDP"
        End If
    End If
End Sub
 
Private Sub SendBufferedDataTCP()
    Dim arrData()       As Byte
    Dim lngBufferLength As Long
    Dim lngResult       As Long
    Dim lngTotalSent    As Long
    Do Until lngResult = Socket_Error Or Len(m_strSendBuffer) = 0
        lngBufferLength = Len(m_strSendBuffer)
        If lngBufferLength > m_lngSendBufferLen Then
            lngBufferLength = m_lngSendBufferLen
            arrData() = StrConv(Left$(m_strSendBuffer, m_lngSendBufferLen), vbFromUnicode)
        Else
            arrData() = StrConv(m_strSendBuffer, vbFromUnicode)
        End If
        lngResult = api_send(m_lngSocketHandle, arrData(0), lngBufferLength, 0&)
        If lngResult = Socket_Error Then
            Dim lngErrorCode As Long
            lngErrorCode = Err.LastDllError
            If lngErrorCode = WSAEWOULDBLOCK Then
                If lngTotalSent > 0 Then RaiseEvent SendProgress(lngTotalSent, Len(m_strSendBuffer))
            Else
                m_enmState = sckError
                Dim blnCancelDisplay As Boolean
                blnCancelDisplay = True
                RaiseEvent Error(lngErrorCode, GetErrorDescription(lngErrorCode), 0, "Socket.SendBufferedData", "", 0, blnCancelDisplay)
                If blnCancelDisplay = False Then MsgBox GetErrorDescription(lngErrorCode), vbOKOnly, "Socket.SendBufferedData"
            End If
        Else
            lngTotalSent = lngTotalSent + lngResult
            If Len(m_strSendBuffer) > lngResult Then
                m_strSendBuffer = mid$(m_strSendBuffer, lngResult + 1)
            Else
                m_strSendBuffer = ""
                Dim lngTemp As Long
                lngTemp = lngTotalSent
                lngTotalSent = 0
                RaiseEvent SendProgress(lngTemp, 0)
                RaiseEvent SendComplete
            End If
        End If
    Loop
End Sub
 
Private Function RecvDataToBuffer() As Long
    Dim arrBuffer()      As Byte
    Dim lngBytesReceived As Long
    Dim strBuffTemporal  As String
    ReDim arrBuffer(m_lngRecvBufferLen - 1)
    lngBytesReceived = api_recv(m_lngSocketHandle, arrBuffer(0), m_lngRecvBufferLen, 0&)
    If lngBytesReceived = Socket_Error Then
        m_enmState = sckError
        Dim lngErrorCode As Long
        lngErrorCode = Err.LastDllError
        Exit Function
    ElseIf lngBytesReceived > 0 Then
        strBuffTemporal = StrConv(arrBuffer(), vbUnicode)
        m_strRecvBuffer = m_strRecvBuffer & Left$(strBuffTemporal, lngBytesReceived)
        RecvDataToBuffer = lngBytesReceived
    End If
End Function
 
Private Sub ProcessOptions()
    Dim lngResult    As Long
    Dim lngBuffer    As Long
    Dim lngErrorCode As Long
    If m_enmProtocol = sckTCPProtocol Then
        lngResult = api_getsockopt(m_lngSocketHandle, SOL_SOCKET, SO_RCVBUF, lngBuffer, LenB(lngBuffer))
        If lngResult = Socket_Error Then
            lngErrorCode = Err.LastDllError
            Exit Sub
        Else
            m_lngRecvBufferLen = lngBuffer
        End If
        lngResult = api_getsockopt(m_lngSocketHandle, SOL_SOCKET, SO_SNDBUF, lngBuffer, LenB(lngBuffer))
        If lngResult = Socket_Error Then
            lngErrorCode = Err.LastDllError
            Exit Sub
        Else
            m_lngSendBufferLen = lngBuffer
        End If
    Else
        lngBuffer = 1
        lngResult = api_setsockopt(m_lngSocketHandle, SOL_SOCKET, SO_BROADCAST, lngBuffer, LenB(lngBuffer))
        lngResult = api_getsockopt(m_lngSocketHandle, SOL_SOCKET, SO_MAX_MSG_SIZE, lngBuffer, LenB(lngBuffer))
        If lngResult = Socket_Error Then
            lngErrorCode = Err.LastDllError
            Exit Sub
        Else
            m_lngRecvBufferLen = lngBuffer
            m_lngSendBufferLen = lngBuffer
        End If
    End If
End Sub
 
Public Sub GetData(ByRef data As Variant, Optional varType As Variant, Optional maxLen As Variant)
    If m_enmProtocol = sckTCPProtocol Then
        If m_enmState <> sckConnected And Not m_blnAcceptClass Then
            Exit Sub
            Exit Sub
        End If
    Else
        If m_enmState <> sckOpen Then
            Exit Sub
            Exit Sub
        End If
        If GetBufferLenUDP = 0 Then Exit Sub
    End If
    If Not IsMissing(maxLen) Then
        If IsNumeric(maxLen) Then
            If CLng(maxLen) < 0 Then
                Exit Sub
            End If
        Else
            If m_enmProtocol = sckTCPProtocol Then
                maxLen = Len(m_strRecvBuffer)
            Else
                maxLen = GetBufferLenUDP
            End If
        End If
    End If
    Dim lngBytesRecibidos As Long
    lngBytesRecibidos = RecvData(data, False, varType, maxLen)
End Sub
 
Public Sub PeekData(ByRef data As Variant, Optional varType As Variant, Optional maxLen As Variant)
    If m_enmProtocol = sckTCPProtocol Then
        If m_enmState <> sckConnected Then
            Exit Sub
            Exit Sub
        End If
    Else
        If m_enmState <> sckOpen Then
            Exit Sub
            Exit Sub
        End If
        If GetBufferLenUDP = 0 Then Exit Sub
    End If
    If Not IsMissing(maxLen) Then
        If IsNumeric(maxLen) Then
            If CLng(maxLen) < 0 Then
                Exit Sub
            End If
        Else
            If m_enmProtocol = sckTCPProtocol Then
                maxLen = Len(m_strRecvBuffer)
            Else
                maxLen = GetBufferLenUDP
            End If
        End If
    End If
    Dim lngBytesRecibidos As Long
    lngBytesRecibidos = RecvData(data, True, varType, maxLen)
End Sub
 
Private Function RecvData(ByRef data As Variant, ByVal blnPeek As Boolean, Optional varClass As Variant, Optional maxLen As Variant) As Long
    Dim blnMaxLenMiss As Boolean
    Dim blnClassMiss  As Boolean
    Dim lngBufferLen  As Long
    Dim arrBuffer()   As Byte
    Dim lngErrorCode  As Long
    If m_enmProtocol = sckTCPProtocol Then
        lngBufferLen = Len(m_strRecvBuffer)
    Else
        lngBufferLen = GetBufferLenUDP
    End If
    blnMaxLenMiss = IsMissing(maxLen)
    blnClassMiss = IsMissing(varClass)
    If varType(data) = vbEmpty Then
        If blnClassMiss Then varClass = vbArray + vbByte
    Else
        varClass = varType(data)
    End If
    If varClass = vbString Or varClass = vbArray + vbByte Then
        If blnMaxLenMiss Then
            If lngBufferLen = 0 Then
                RecvData = 0
                arrBuffer = StrConv("", vbFromUnicode)
                data = arrBuffer
                Exit Function
            Else
                RecvData = lngBufferLen
                BuildArray lngBufferLen, blnPeek, lngErrorCode, arrBuffer
            End If
        Else
            If maxLen = 0 Or lngBufferLen = 0 Then
                RecvData = 0
                arrBuffer = StrConv("", vbFromUnicode)
                data = arrBuffer
                If m_enmProtocol = sckUDPProtocol Then
                    EmptyBuffer
                    Exit Function
                End If
                Exit Function
            ElseIf maxLen > lngBufferLen Then
                RecvData = lngBufferLen
                BuildArray lngBufferLen, blnPeek, lngErrorCode, arrBuffer
            Else
                RecvData = CLng(maxLen)
                BuildArray CLng(maxLen), blnPeek, lngErrorCode, arrBuffer
            End If
        End If
    End If
    Select Case varClass
        Case vbString
            Dim strData As String
            strData = StrConv(arrBuffer(), vbUnicode)
            data = strData

        Case vbArray + vbByte
            data = arrBuffer
            
        Case vbBoolean
            Dim blnData As Boolean
            If LenB(blnData) > lngBufferLen Then Exit Function
            BuildArray LenB(blnData), blnPeek, lngErrorCode, arrBuffer
            RecvData = LenB(blnData)
            api_CopyMemory blnData, arrBuffer(0), LenB(blnData)
            data = blnData

        Case vbByte
            Dim bytData As Byte
            If LenB(bytData) > lngBufferLen Then Exit Function
            BuildArray LenB(bytData), blnPeek, lngErrorCode, arrBuffer
            RecvData = LenB(bytData)
            api_CopyMemory bytData, arrBuffer(0), LenB(bytData)
            data = bytData

        Case vbCurrency
            Dim curData As Currency
            If LenB(curData) > lngBufferLen Then Exit Function
            BuildArray LenB(curData), blnPeek, lngErrorCode, arrBuffer
            RecvData = LenB(curData)
            api_CopyMemory curData, arrBuffer(0), LenB(curData)
            data = curData

        Case vbDate
            Dim datData As Date
            If LenB(datData) > lngBufferLen Then Exit Function
            BuildArray LenB(datData), blnPeek, lngErrorCode, arrBuffer
            RecvData = LenB(datData)
            api_CopyMemory datData, arrBuffer(0), LenB(datData)
            data = datData

        Case vbDouble
            Dim dblData As Double
            If LenB(dblData) > lngBufferLen Then Exit Function
            BuildArray LenB(dblData), blnPeek, lngErrorCode, arrBuffer
            RecvData = LenB(dblData)
            api_CopyMemory dblData, arrBuffer(0), LenB(dblData)
            data = dblData

        Case vbInteger
            Dim intData As Integer
            If LenB(intData) > lngBufferLen Then Exit Function
            BuildArray LenB(intData), blnPeek, lngErrorCode, arrBuffer
            RecvData = LenB(intData)
            api_CopyMemory intData, arrBuffer(0), LenB(intData)
            data = intData

        Case vbLong
            Dim lngData As Long
            If LenB(lngData) > lngBufferLen Then Exit Function
            BuildArray LenB(lngData), blnPeek, lngErrorCode, arrBuffer
            RecvData = LenB(lngData)
            api_CopyMemory lngData, arrBuffer(0), LenB(lngData)
            data = lngData

        Case vbSingle
            Dim sngData As Single
            If LenB(sngData) > lngBufferLen Then Exit Function
            BuildArray LenB(sngData), blnPeek, lngErrorCode, arrBuffer
            RecvData = LenB(sngData)
            api_CopyMemory sngData, arrBuffer(0), LenB(sngData)
            data = sngData

        Case Else
            Exit Function
    End Select
    If lngErrorCode <> 0 Then
        Exit Function
    End If
End Function
 
Private Sub BuildArray(ByVal Size As Long, ByVal blnPeek As Boolean, ByRef lngErrorCode As Long, ByRef bytArray() As Byte)
    Dim strData As String
    If m_enmProtocol = sckTCPProtocol Then
        strData = Left$(m_strRecvBuffer, CLng(Size))
        bytArray = StrConv(strData, vbFromUnicode)
        If Not blnPeek Then
            m_strRecvBuffer = mid$(m_strRecvBuffer, Size + 1)
        End If
    Else
        Dim arrBuffer() As Byte
        Dim lngResult   As Long
        Dim udtSockAddr As sockaddr_in
        Dim lngFlags    As Long
        If blnPeek Then lngFlags = MSG_PEEK
        ReDim arrBuffer(Size - 1)
        lngResult = api_recvfrom(m_lngSocketHandle, arrBuffer(0), Size, lngFlags, udtSockAddr, LenB(udtSockAddr))
        If lngResult = Socket_Error Then
            lngErrorCode = Err.LastDllError
        End If
        bytArray = arrBuffer
        GetRemoteInfoFromSI udtSockAddr, m_lngRemotePort, m_strRemoteHostIP, m_strRemoteHost
    End If
End Sub
 
Private Sub CleanResolutionSystem()
    Dim varAsynHandle As Variant
    Dim lngResult     As Long
    For Each varAsynHandle In m_colWaitingResolutions
        lngResult = api_WSACancelAsyncRequest(varAsynHandle)
        If lngResult = 0 Then
            modSocket.UnregisterResolution varAsynHandle
            Set m_colWaitingResolutions = Nothing
            Set m_colWaitingResolutions = New Collection
            FreeMemory
        End If
    Next
End Sub
 
Public Sub Listen()
    If m_enmState <> sckClosed And m_enmState <> sckOpen Then
        Exit Sub
    End If
    If Not SocketExists Then Exit Sub
    If Not BindInternal Then Exit Sub
    Dim lngResult As Long
    lngResult = api_listen(m_lngSocketHandle, SOMAXCONN)
    If lngResult = Socket_Error Then
        Dim lngErrorCode As Long
        lngErrorCode = Err.LastDllError
        Exit Sub
    Else
        m_enmState = sckListening
    End If
End Sub
 
Public Sub Accept(requestID As Long)
    If m_enmState <> sckClosed Then
        Exit Sub
    End If
    m_lngSocketHandle = requestID
    m_enmProtocol = sckTCPProtocol
    ProcessOptions
    If Not modSocket.IsAcceptRegistered(requestID) Then
        If IsSocketRegistered(requestID) Then
            m_lngSocketHandle = INVALID_SOCKET
            m_lngRecvBufferLen = 0
            m_lngSendBufferLen = 0
            Exit Sub
        Else
            m_blnAcceptClass = True
            m_enmState = sckConnected
            GetLocalInfo m_lngSocketHandle, m_lngLocalPortBind, m_strLocalIP
            modSocket.RegisterSocket m_lngSocketHandle, ObjPtr(Me), False
            Exit Sub
        End If
    End If
    Dim clsSocket As clsSocket
    Set clsSocket = GetAcceptClass(requestID)
    modSocket.UnregisterAccept requestID
    GetLocalInfo m_lngSocketHandle, m_lngLocalPortBind, m_strLocalIP
    GetRemoteInfo m_lngSocketHandle, m_lngRemotePort, m_strRemoteHostIP, m_strRemoteHost
    m_enmState = sckConnected
    If clsSocket.BytesReceived > 0 Then
        clsSocket.GetData m_strRecvBuffer
    End If
    modSocket.Subclass_ChangeOwner requestID, ObjPtr(Me)
    If Len(m_strRecvBuffer) > 0 Then RaiseEvent DataArrival(Len(m_strRecvBuffer))
    If clsSocket.State = sckClosing Then
        m_enmState = sckClosing
        RaiseEvent CloseSck
    End If
    Set clsSocket = Nothing
End Sub
 
Private Function GetLocalInfo(ByVal lngSocket As Long, ByRef lngLocalPort As Long, ByRef strLocalIP As String) As Boolean
    GetLocalInfo = False
    Dim lngResult   As Long
    Dim udtSockAddr As sockaddr_in
    lngResult = api_getsockname(lngSocket, udtSockAddr, LenB(udtSockAddr))
    If lngResult = Socket_Error Then
        lngLocalPort = 0
        strLocalIP = ""
    Else
        GetLocalInfo = True
        lngLocalPort = IntegerToUnsigned(api_ntohs(udtSockAddr.sin_port))
        strLocalIP = StringFromPointer(api_inet_ntoa(udtSockAddr.sin_addr))
    End If
End Function
 
Private Function GetRemoteInfo(ByVal lngSocket As Long, ByRef lngRemotePort As Long, ByRef strRemoteHostIP As String, ByRef strRemoteHost As String) As Boolean
    GetRemoteInfo = False
    Dim lngResult   As Long
    Dim udtSockAddr As sockaddr_in
    lngResult = api_getpeername(lngSocket, udtSockAddr, LenB(udtSockAddr))
    If lngResult = 0 Then
        GetRemoteInfo = True
        GetRemoteInfoFromSI udtSockAddr, lngRemotePort, strRemoteHostIP, strRemoteHost
    Else
        lngRemotePort = 0
        strRemoteHostIP = ""
        strRemoteHost = ""
    End If
End Function
 
Private Sub GetRemoteInfoFromSI(ByRef udtSockAddr As sockaddr_in, ByRef lngRemotePort As Long, ByRef strRemoteHostIP As String, ByRef strRemoteHost As String)
    lngRemotePort = IntegerToUnsigned(api_ntohs(udtSockAddr.sin_port))
    strRemoteHostIP = StringFromPointer(api_inet_ntoa(udtSockAddr.sin_addr))
End Sub
 
Private Function GetBufferLenUDP() As Long
    Dim lngResult As Long
    Dim lngBuffer As Long
    lngResult = api_ioctlsocket(m_lngSocketHandle, FIONREAD, lngBuffer)
    If lngResult = Socket_Error Then
        GetBufferLenUDP = 0
    Else
        GetBufferLenUDP = lngBuffer
    End If
End Function
 
Private Sub EmptyBuffer()
    Dim b As Byte
    api_recv m_lngSocketHandle, b, Len(b), 0&
End Sub
