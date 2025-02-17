Attribute VB_Name = "modSocket"
Option Explicit
Public Declare Sub api_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)
Public Declare Function api_GlobalAlloc Lib "kernel32" Alias "GlobalAlloc" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function api_GlobalFree Lib "kernel32" Alias "GlobalFree" (ByVal hMem As Long) As Long
Private Declare Function api_WSAStartup Lib "ws2_32.dll" Alias "WSAStartup" (ByVal wVersionRequired As Long, lpWSADATA As WSAData) As Long
Private Declare Function api_WSACleanup Lib "ws2_32.dll" Alias "WSACleanup" () As Long
Private Declare Function api_WSAAsyncGetHostByName Lib "ws2_32.dll" Alias "WSAAsyncGetHostByName" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal strHostName As String, buf As Any, ByVal buflen As Long) As Long
Private Declare Function api_WSAAsyncSelect Lib "ws2_32.dll" Alias "WSAAsyncSelect" (ByVal s As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Function api_CreateWindowEx Lib "user32" _
                Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
                                         ByVal lpClassName As String, _
                                         ByVal lpWindowName As String, _
                                         ByVal dwStyle As Long, _
                                         ByVal X As Long, _
                                         ByVal Y As Long, _
                                         ByVal nWidth As Long, _
                                         ByVal nHeight As Long, _
                                         ByVal hWndParent As Long, _
                                         ByVal hMenu As Long, _
                                         ByVal hInstance As Long, _
                                         lpParam As Any) As Long
Private Declare Function api_DestroyWindow Lib "user32" Alias "DestroyWindow" (ByVal hwnd As Long) As Long
Private Declare Function api_lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function api_lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function api_LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function api_SetTimer Lib "user32" Alias "SetTimer" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function api_KillTimer Lib "user32" Alias "KillTimer" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public Const Socket_Error        As Integer = -1
Public Const INVALID_SOCKET      As Integer = -1
Public Const INADDR_NONE         As Long = &HFFFF
Private Const WSADESCRIPTION_LEN As Integer = 257
Private Const WSASYS_STATUS_LEN  As Integer = 129

Private Enum WinsockVersion
    SOCKET_VERSION_11 = &H101
    SOCKET_VERSION_22 = &H202
End Enum

Public Const MAXGETHOSTSTRUCT   As Long = 1024
Public Const AF_INET            As Long = 2
Public Const SOCK_STREAM        As Long = 1
Public Const SOCK_DGRAM         As Long = 2
Public Const IPPROTO_TCP        As Long = 6
Public Const IPPROTO_UDP        As Long = 17
Public Const FD_READ            As Integer = &H1&
Public Const FD_WRITE           As Integer = &H2&
Public Const FD_ACCEPT          As Integer = &H8&
Public Const FD_CONNECT         As Integer = &H10&
Public Const FD_CLOSE           As Integer = &H20&
Private Const OFFSET_2          As Long = 65536
Private Const MAXINT_2          As Long = 32767
Public Const GMEM_FIXED         As Integer = &H0
Public Const LOCAL_HOST_BUFF    As Integer = 256
Public Const SOL_SOCKET         As Long = 65535
Public Const SO_SNDBUF          As Long = &H1001&
Public Const SO_RCVBUF          As Long = &H1002&
Public Const SO_MAX_MSG_SIZE    As Long = &H2003
Public Const SO_BROADCAST       As Long = &H20
Public Const FIONREAD           As Long = &H4004667F
Public Const WSABASEERR         As Long = 10000
Public Const WSAEINTR           As Long = (WSABASEERR + 4)
Public Const WSAEACCES          As Long = (WSABASEERR + 13)
Public Const WSAEFAULT          As Long = (WSABASEERR + 14)
Public Const WSAEINVAL          As Long = (WSABASEERR + 22)
Public Const WSAEMFILE          As Long = (WSABASEERR + 24)
Public Const WSAEWOULDBLOCK     As Long = (WSABASEERR + 35)
Public Const WSAEINPROGRESS     As Long = (WSABASEERR + 36)
Public Const WSAEALREADY        As Long = (WSABASEERR + 37)
Public Const WSAENOTSOCK        As Long = (WSABASEERR + 38)
Public Const WSAEDESTADDRREQ    As Long = (WSABASEERR + 39)
Public Const WSAEMSGSIZE        As Long = (WSABASEERR + 40)
Public Const WSAEPROTOTYPE      As Long = (WSABASEERR + 41)
Public Const WSAENOPROTOOPT     As Long = (WSABASEERR + 42)
Public Const WSAEPROTONOSUPPORT As Long = (WSABASEERR + 43)
Public Const WSAESOCKTNOSUPPORT As Long = (WSABASEERR + 44)
Public Const WSAEOPNOTSUPP      As Long = (WSABASEERR + 45)
Public Const WSAEPFNOSUPPORT    As Long = (WSABASEERR + 46)
Public Const WSAEAFNOSUPPORT    As Long = (WSABASEERR + 47)
Public Const WSAEADDRINUSE      As Long = (WSABASEERR + 48)
Public Const WSAEADDRNOTAVAIL   As Long = (WSABASEERR + 49)
Public Const WSAENETDOWN        As Long = (WSABASEERR + 50)
Public Const WSAENETUNREACH     As Long = (WSABASEERR + 51)
Public Const WSAENETRESET       As Long = (WSABASEERR + 52)
Public Const WSAECONNABORTED    As Long = (WSABASEERR + 53)
Public Const WSAECONNRESET      As Long = (WSABASEERR + 54)
Public Const WSAENOBUFS         As Long = (WSABASEERR + 55)
Public Const WSAEISCONN         As Long = (WSABASEERR + 56)
Public Const WSAENOTCONN        As Long = (WSABASEERR + 57)
Public Const WSAESHUTDOWN       As Long = (WSABASEERR + 58)
Public Const WSAETIMEDOUT       As Long = (WSABASEERR + 60)
Public Const WSAEHOSTUNREACH    As Long = (WSABASEERR + 65)
Public Const WSAECONNREFUSED    As Long = (WSABASEERR + 61)
Public Const WSAEPROCLIM        As Long = (WSABASEERR + 67)
Public Const WSASYSNOTREADY     As Long = (WSABASEERR + 91)
Public Const WSAVERNOTSUPPORTED As Long = (WSABASEERR + 92)
Public Const WSANOTINITIALISED  As Long = (WSABASEERR + 93)
Public Const WSAHOST_NOT_FOUND  As Long = (WSABASEERR + 1001)
Public Const WSATRY_AGAIN       As Long = (WSABASEERR + 1002)
Public Const WSANO_RECOVERY     As Long = (WSABASEERR + 1003)
Public Const WSANO_DATA         As Long = (WSABASEERR + 1004)
Public Const sckOutOfMemory     As Long = 7
Public Const sckBadState        As Long = 40006
Public Const sckInvalidArg      As Long = 40014
Public Const sckUnsupported     As Long = 40018
Public Const sckInvalidOp       As Long = 40020

Private Type WSAData
    wVersion       As Integer
    wHighVersion   As Integer
    szDescription  As String * WSADESCRIPTION_LEN
    szSystemStatus As String * WSASYS_STATUS_LEN
    iMaxSockets    As Integer
    iMaxUdpDg      As Integer
    lpVendorInfo   As Long
End Type
 
Public Type HOSTENT
    hName     As Long
    hAliases  As Long
    hAddrType As Integer
    hLength   As Integer
    hAddrList As Long
End Type
 
Public Type sockaddr_in
    sin_family       As Integer
    sin_port         As Integer
    sin_addr         As Long
    sin_zero(1 To 8) As Byte
End Type
  
Private m_blnInitiated     As Boolean
Private m_lngSocksQuantity As Long
Private m_colSocketsInst   As Collection
Private m_colAcceptList    As Collection
Private m_lngWindowHandle  As Long
Private Declare Function api_IsWindow Lib "user32" Alias "IsWindow" (ByVal hwnd As Long) As Long
Private Declare Function api_GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function api_SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function api_GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function api_GetProcAddress Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Const PATCH_09       As Long = 119
Private Const PATCH_0C       As Long = 150
Private Const GWL_WNDPROC    As Long = (-4)
Private Const WM_APP         As Long = 32768
Public Const RESOLVE_MESSAGE As Long = WM_APP
Public Const SOCKET_MESSAGE  As Long = WM_APP + 1
Private Const TIMER_TIMEOUT  As Long = 200
Private lngMsgCntA           As Long
Private lngMsgCntB           As Long
Private lngTableA1()         As Long
Private lngTableA2()         As Long
Private lngTableB1()         As Long
Private lngTableB2()         As Long
Private hWndSub              As Long
Private nAddrSubclass        As Long
Private nAddrOriginal        As Long
Private hTimer               As Long

Public Function InitiateProcesses() As Long
    InitiateProcesses = 0
    m_lngSocksQuantity = m_lngSocksQuantity + 1
    If Not m_blnInitiated Then
        Subclass_Initialize
        m_blnInitiated = True
        Dim lngResult As Long
        lngResult = InitiateService
        If lngResult <> 0 Then
            InitiateProcesses = lngResult
        End If
    End If
End Function
 
Private Function InitiateService() As Long
    Dim udtWSAData As WSAData
    Dim lngResult  As Long
    lngResult = api_WSAStartup(SOCKET_VERSION_22, udtWSAData)
    InitiateService = lngResult
End Function

Public Function FinalizeProcesses() As Long
    FinalizeProcesses = 0
    m_lngSocksQuantity = m_lngSocksQuantity - 1
    If m_blnInitiated And m_lngSocksQuantity = 0 Then
        If FinalizeService = Socket_Error Then
            Dim lngErrorCode As Long
            lngErrorCode = Err.LastDllError
            FinalizeProcesses = lngErrorCode
        End If
        Subclass_Terminate
        m_blnInitiated = False
    End If
End Function
 
Private Function FinalizeService() As Long
    Dim lngResultado As Long
    lngResultado = api_WSACleanup
    FinalizeService = lngResultado
End Function
 
Public Function GetErrorDescription(ByVal lngErrorCode As Long) As String
    Select Case lngErrorCode
        Case WSAEACCES
            GetErrorDescription = "Permission denied."

        Case WSAEADDRINUSE
            GetErrorDescription = "Address already in use."

        Case WSAEADDRNOTAVAIL
            GetErrorDescription = "Cannot assign requested address."

        Case WSAEAFNOSUPPORT
            GetErrorDescription = "Address family not supported by protocol family."

        Case WSAEALREADY
            GetErrorDescription = "Operation already in progress."

        Case WSAECONNABORTED
            GetErrorDescription = "Software caused connection abort."

        Case WSAECONNREFUSED
            GetErrorDescription = "Connection refused."

        Case WSAECONNRESET
            GetErrorDescription = "Connection reset by peer."

        Case WSAEDESTADDRREQ
            GetErrorDescription = "Destination address required."

        Case WSAEFAULT
            GetErrorDescription = "Bad address."

        Case WSAEHOSTUNREACH
            GetErrorDescription = "No route to host."

        Case WSAEINPROGRESS
            GetErrorDescription = "Operation now in progress."

        Case WSAEINTR
            GetErrorDescription = "Interrupted function call."

        Case WSAEINVAL
            GetErrorDescription = "Invalid argument."

        Case WSAEISCONN
            GetErrorDescription = "Socket is already connected."

        Case WSAEMFILE
            GetErrorDescription = "Too many open files."

        Case WSAEMSGSIZE
            GetErrorDescription = "Message too long."

        Case WSAENETDOWN
            GetErrorDescription = "Network is down."

        Case WSAENETRESET
            GetErrorDescription = "Network dropped connection on reset."

        Case WSAENETUNREACH
            GetErrorDescription = "Network is unreachable."

        Case WSAENOBUFS
            GetErrorDescription = "No buffer space available."

        Case WSAENOPROTOOPT
            GetErrorDescription = "Bad protocol option."

        Case WSAENOTCONN
            GetErrorDescription = "Socket is not connected."

        Case WSAENOTSOCK
            GetErrorDescription = "Socket operation on nonsocket."

        Case WSAEOPNOTSUPP
            GetErrorDescription = "Operation not supported."

        Case WSAEPFNOSUPPORT
            GetErrorDescription = "Protocol family not supported."

        Case WSAEPROCLIM
            GetErrorDescription = "Too many processes."

        Case WSAEPROTONOSUPPORT
            GetErrorDescription = "Protocol not supported."

        Case WSAEPROTOTYPE
            GetErrorDescription = "Protocol wrong type for socket."

        Case WSAESHUTDOWN
            GetErrorDescription = "Cannot send after socket shutdown."

        Case WSAESOCKTNOSUPPORT
            GetErrorDescription = "Socket type not supported."

        Case WSAETIMEDOUT
            GetErrorDescription = "Connection timed out."

        Case WSAEWOULDBLOCK
            GetErrorDescription = "Resource temporarily unavailable."

        Case WSAHOST_NOT_FOUND
            GetErrorDescription = "Host not found."

        Case WSANOTINITIALISED
            GetErrorDescription = "Successful WSAStartup not yet performed."

        Case WSANO_DATA
            GetErrorDescription = "Valid name, no data record of requested type."

        Case WSANO_RECOVERY
            GetErrorDescription = "This is a nonrecoverable error."

        Case WSASYSNOTREADY
            GetErrorDescription = "Network subsystem is unavailable."

        Case WSATRY_AGAIN
            GetErrorDescription = "Nonauthoritative host not found."

        Case WSAVERNOTSUPPORTED
            GetErrorDescription = "Winsock.dll version out of range."

        Case Else
            GetErrorDescription = "Unknown error."
    End Select
End Function
 
Private Function CreateWinsockMessageWindow() As Long
    m_lngWindowHandle = api_CreateWindowEx(0&, "STATIC", "SOCKET_WINDOW", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)
    If m_lngWindowHandle = 0 Then
        CreateWinsockMessageWindow = sckOutOfMemory
        Exit Function
    Else
        CreateWinsockMessageWindow = 0
    End If
End Function
 
Private Function DestroyWinsockMessageWindow() As Long
    DestroyWinsockMessageWindow = 0
    If m_lngWindowHandle = 0 Then
        Exit Function
    End If
    Dim lngResult As Long
    lngResult = api_DestroyWindow(m_lngWindowHandle)
    If lngResult = 0 Then
        DestroyWinsockMessageWindow = sckOutOfMemory
    Else
        m_lngWindowHandle = 0
    End If
End Function
 
Public Function ResolveHost(ByVal strHost As String, ByVal lngHOSTENBuf As Long, ByVal lngObjectPointer As Long) As Long
    Dim lngAsynHandle As Long
    lngAsynHandle = api_WSAAsyncGetHostByName(m_lngWindowHandle, RESOLVE_MESSAGE, strHost, ByVal lngHOSTENBuf, MAXGETHOSTSTRUCT)
    If lngAsynHandle <> 0 Then Subclass_AddResolveMessage lngAsynHandle, lngObjectPointer
    ResolveHost = lngAsynHandle
End Function
 
Public Function HiWord(lngValue As Long) As Long
    If (lngValue And &H80000000) = &H80000000 Then
        HiWord = ((lngValue And &H7FFF0000) \ &H10000) Or &H8000&
    Else
        HiWord = (lngValue And &HFFFF0000) \ &H10000
    End If
End Function
 
Public Function LoWord(lngValue As Long) As Long
    LoWord = (lngValue And &HFFFF&)
End Function
 
Public Function StringFromPointer(ByVal lPointer As Long) As String
    Dim strTemp As String
    Dim lRetVal As Long
    strTemp = String$(api_lstrlen(ByVal lPointer), 0)
    lRetVal = api_lstrcpy(ByVal strTemp, ByVal lPointer)
    If lRetVal Then StringFromPointer = strTemp
End Function
 
Public Function UnsignedToInteger(Value As Long) As Integer
    If Value < 0 Or Value >= OFFSET_2 Then Error 6
    If Value <= MAXINT_2 Then
        UnsignedToInteger = Value
    Else
        UnsignedToInteger = Value - OFFSET_2
    End If
End Function
 
Public Function IntegerToUnsigned(Value As Integer) As Long
    If Value < 0 Then
        IntegerToUnsigned = Value + OFFSET_2
    Else
        IntegerToUnsigned = Value
    End If
End Function
 
Public Function RegisterSocket(ByVal lngSocket As Long, ByVal lngObjectPointer As Long, ByVal blnEvents As Boolean) As Boolean
    If m_colSocketsInst Is Nothing Then
        Set m_colSocketsInst = New Collection
        If CreateWinsockMessageWindow <> 0 Then
        End If
        Subclass_Subclass (m_lngWindowHandle)
    End If
    Subclass_AddSocketMessage lngSocket, lngObjectPointer
    If blnEvents Then
        Dim lngEvents    As Long
        Dim lngResult    As Long
        Dim lngErrorCode As Long
        lngEvents = FD_READ Or FD_WRITE Or FD_ACCEPT Or FD_CONNECT Or FD_CLOSE
        lngResult = api_WSAAsyncSelect(lngSocket, m_lngWindowHandle, SOCKET_MESSAGE, lngEvents)
        If lngResult = Socket_Error Then
            lngErrorCode = Err.LastDllError
        End If
    End If
    m_colSocketsInst.Add lngObjectPointer, "S" & lngSocket
    RegisterSocket = True
End Function
 
Public Sub UnregisterSocket(ByVal lngSocket As Long)
    Subclass_DelSocketMessage lngSocket
    On Error Resume Next
    m_colSocketsInst.Remove "S" & lngSocket
    If m_colSocketsInst.Count = 0 Then
        Set m_colSocketsInst = Nothing
        Subclass_UnSubclass
        DestroyWinsockMessageWindow
    End If
End Sub
 
Public Function IsSocketRegistered(ByVal lngSocket As Long) As Boolean
    On Error GoTo ErrorHandler
    m_colSocketsInst.item ("S" & lngSocket)
    IsSocketRegistered = True
    Exit Function
ErrorHandler:
    IsSocketRegistered = False
End Function
 
Public Sub UnregisterResolution(ByVal lngAsynHandle As Long)
    Subclass_DelResolveMessage lngAsynHandle
End Sub
 
Public Sub RegisterAccept(ByVal lngSocket As Long)
    If m_colAcceptList Is Nothing Then
        Set m_colAcceptList = New Collection
    End If
    Dim Socket As clsSocket
    Set Socket = New clsSocket
    Socket.Accept lngSocket
    m_colAcceptList.Add Socket, "S" & lngSocket
End Sub
 
Public Function IsAcceptRegistered(ByVal lngSocket As Long) As Boolean
    On Error GoTo ErrorHandler
    m_colAcceptList.item ("S" & lngSocket)
    IsAcceptRegistered = True
    Exit Function
ErrorHandler:
    IsAcceptRegistered = False
End Function
 
Public Sub UnregisterAccept(ByVal lngSocket As Long)
    m_colAcceptList.Remove "S" & lngSocket
    If m_colAcceptList.Count = 0 Then
        Set m_colAcceptList = Nothing
    End If
End Sub

Public Function GetAcceptClass(ByVal lngSocket As Long) As clsSocket
    Set GetAcceptClass = m_colAcceptList("S" & lngSocket)
End Function
  
Private Sub Subclass_Initialize()
    Const PATCH_01 As Long = 16
    Const PATCH_03 As Long = 72
    Const PATCH_04 As Long = 77
    Const PATCH_06 As Long = 89
    Const PATCH_08 As Long = 113
    Const FUNC_EBM As String = "EbMode"
    Const FUNC_SWL As String = "SetWindowLongA"
    Const FUNC_CWP As String = "CallWindowProcA"
    Const FUNC_WCU As String = "WSACleanup"
    Const FUNC_KTM As String = "KillTimer"
    Const MOD_VBA5 As String = "vba5"
    Const MOD_VBA6 As String = "vba6"
    Const MOD_USER As String = "user32"
    Const MOD_WS   As String = "ws2_32"
    Dim i          As Long
    Dim nLen       As Long
    Dim sHex       As String
    Dim sCode      As String
    sHex = "5850505589E55753515231C0FCEB09E8xxxxx01x85C074258B45103D0080000074543D01800000746CE8310000005A595B5FC9C21400E824000000EBF168xxxxx02x6AFCFF750CE8xxxxx03xE8xxxxx04x68xxxxx05x6A00E8xxxxx06xEBCFFF7518FF7514FF7510FF750C68xxxxx07xE8xxxxx08xC3BBxxxxx09x8B4514BFxxxxx0Ax89D9F2AF75A529CB4B8B1C9Dxxxxx0BxEB1DBBxxxxx0Cx8B4514BFxxxxx0Dx89D9F2AF758629CB4B8B1C9Dxxxxx0Ex895D088B1B8B5B1C89D85A595B5FC9FFE0"
    nLen = Len(sHex)
    For i = 1 To nLen Step 2
        sCode = sCode & ChrB$(Val("&H" & mid$(sHex, i, 2)))
    Next i
    nLen = LenB(sCode)
    nAddrSubclass = api_GlobalAlloc(0, nLen)
    Call api_CopyMemory(ByVal nAddrSubclass, ByVal StrPtr(sCode), nLen)
    If Subclass_InIDE Then
        Call api_CopyMemory(ByVal nAddrSubclass + 13, &H9090, 2)
        i = Subclass_AddrFunc(MOD_VBA6, FUNC_EBM)
        If i = 0 Then
            i = Subclass_AddrFunc(MOD_VBA5, FUNC_EBM)
        End If
        Debug.Assert i
        Call Subclass_PatchRel(PATCH_01, i)
    End If
    Call api_LoadLibrary(MOD_WS)
    Call Subclass_PatchRel(PATCH_03, Subclass_AddrFunc(MOD_USER, FUNC_SWL))
    Call Subclass_PatchRel(PATCH_04, Subclass_AddrFunc(MOD_WS, FUNC_WCU))
    Call Subclass_PatchRel(PATCH_06, Subclass_AddrFunc(MOD_USER, FUNC_KTM))
    Call Subclass_PatchRel(PATCH_08, Subclass_AddrFunc(MOD_USER, FUNC_CWP))
End Sub
 
Private Sub Subclass_Terminate()
    Call Subclass_UnSubclass
    Call api_GlobalFree(nAddrSubclass)
    nAddrSubclass = 0
    ReDim lngTableA1(1 To 1)
    ReDim lngTableA2(1 To 1)
    ReDim lngTableB1(1 To 1)
    ReDim lngTableB2(1 To 1)
End Sub
 
Private Function Subclass_InIDE() As Boolean
    Debug.Assert Subclass_SetTrue(Subclass_InIDE)
End Function
 
Private Function Subclass_Subclass(ByVal hwnd As Long) As Boolean
    Const PATCH_02 As Long = 62
    Const PATCH_05 As Long = 82
    Const PATCH_07 As Long = 108
    If hWndSub = 0 Then
        Debug.Assert api_IsWindow(hwnd)
        hWndSub = hwnd
        nAddrOriginal = api_GetWindowLong(hwnd, GWL_WNDPROC)
        Call Subclass_PatchVal(PATCH_02, nAddrOriginal)
        Call Subclass_PatchVal(PATCH_07, nAddrOriginal)
        nAddrOriginal = api_SetWindowLong(hwnd, GWL_WNDPROC, nAddrSubclass)
        If nAddrOriginal <> 0 Then
            Subclass_Subclass = True
        End If
    End If
    If Subclass_InIDE Then
        hTimer = api_SetTimer(0, 0, TIMER_TIMEOUT, nAddrSubclass)
        Call Subclass_PatchVal(PATCH_05, hTimer)
    End If
    Debug.Assert Subclass_Subclass
End Function
 
Private Function Subclass_UnSubclass() As Boolean
    If hWndSub <> 0 Then
        lngMsgCntA = 0
        lngMsgCntB = 0
        Call Subclass_PatchVal(PATCH_09, lngMsgCntA)
        Call Subclass_PatchVal(PATCH_0C, lngMsgCntB)
        Call api_SetWindowLong(hWndSub, GWL_WNDPROC, nAddrOriginal)
        If hTimer <> 0 Then
            Call api_KillTimer(0&, hTimer)
            hTimer = 0
        End If
        hWndSub = 0
        Subclass_UnSubclass = True
    End If
End Function
 
Private Function Subclass_AddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    Subclass_AddrFunc = api_GetProcAddress(api_GetModuleHandle(sDLL), sProc)
End Function
 
Private Function Subclass_AddrMsgTbl(ByRef aMsgTbl() As Long) As Long
    On Error Resume Next
    Subclass_AddrMsgTbl = VarPtr(aMsgTbl(1))
    On Error GoTo 0
End Function
 
Private Sub Subclass_PatchRel(ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call api_CopyMemory(ByVal (nAddrSubclass + nOffset), nTargetAddr - nAddrSubclass - nOffset - 4, 4)
End Sub
 
Private Sub Subclass_PatchVal(ByVal nOffset As Long, ByVal nValue As Long)
    Call api_CopyMemory(ByVal (nAddrSubclass + nOffset), nValue, 4)
End Sub
 
Private Function Subclass_SetTrue(bValue As Boolean) As Boolean
    Subclass_SetTrue = True
    bValue = True
End Function
 
Private Sub Subclass_AddResolveMessage(ByVal lngAsync As Long, ByVal lngObjectPointer As Long)
    Dim Count As Long
    For Count = 1 To lngMsgCntA
        Select Case lngTableA1(Count)
            Case -1
                lngTableA1(Count) = lngAsync
                lngTableA2(Count) = lngObjectPointer
                Exit Sub

            Case lngAsync
                Exit Sub
        End Select
    Next Count
    lngMsgCntA = lngMsgCntA + 1
    ReDim Preserve lngTableA1(1 To lngMsgCntA)
    ReDim Preserve lngTableA2(1 To lngMsgCntA)
    lngTableA1(lngMsgCntA) = lngAsync
    lngTableA2(lngMsgCntA) = lngObjectPointer
    Subclass_PatchTableA
End Sub
 
Private Sub Subclass_AddSocketMessage(ByVal lngSocket As Long, ByVal lngObjectPointer As Long)
    Dim Count As Long
    For Count = 1 To lngMsgCntB
        Select Case lngTableB1(Count)
            Case -1
                lngTableB1(Count) = lngSocket
                lngTableB2(Count) = lngObjectPointer
                Exit Sub

            Case lngSocket
                Exit Sub
        End Select
    Next Count
    lngMsgCntB = lngMsgCntB + 1
    ReDim Preserve lngTableB1(1 To lngMsgCntB)
    ReDim Preserve lngTableB2(1 To lngMsgCntB)
    lngTableB1(lngMsgCntB) = lngSocket
    lngTableB2(lngMsgCntB) = lngObjectPointer
    Subclass_PatchTableB
End Sub
 
Private Sub Subclass_DelResolveMessage(ByVal lngAsync As Long)
    Dim Count As Long
    For Count = 1 To lngMsgCntA
        If lngTableA1(Count) = lngAsync Then
            lngTableA1(Count) = -1
            lngTableA2(Count) = -1
            Exit Sub
        End If
    Next Count
End Sub
 
Private Sub Subclass_DelSocketMessage(ByVal lngSocket As Long)
    Dim Count As Long
    For Count = 1 To lngMsgCntB
        If lngTableB1(Count) = lngSocket Then
            lngTableB1(Count) = -1
            lngTableB2(Count) = -1
            Exit Sub
        End If
    Next Count
End Sub
 
Private Sub Subclass_PatchTableA()
    Const PATCH_0A As Long = 127
    Const PATCH_0B As Long = 143
    Call Subclass_PatchVal(PATCH_09, lngMsgCntA)
    Call Subclass_PatchVal(PATCH_0A, Subclass_AddrMsgTbl(lngTableA1))
    Call Subclass_PatchVal(PATCH_0B, Subclass_AddrMsgTbl(lngTableA2))
End Sub
 
Private Sub Subclass_PatchTableB()
    Const PATCH_0D As Long = 158
    Const PATCH_0E As Long = 174
    Call Subclass_PatchVal(PATCH_0C, lngMsgCntB)
    Call Subclass_PatchVal(PATCH_0D, Subclass_AddrMsgTbl(lngTableB1))
    Call Subclass_PatchVal(PATCH_0E, Subclass_AddrMsgTbl(lngTableB2))
End Sub
 
Public Sub Subclass_ChangeOwner(ByVal lngSocket As Long, ByVal lngObjectPointer As Long)
    Dim Count As Long
    For Count = 1 To lngMsgCntB
        If lngTableB1(Count) = lngSocket Then
            lngTableB2(Count) = lngObjectPointer
            Exit Sub
        End If
    Next Count
End Sub
