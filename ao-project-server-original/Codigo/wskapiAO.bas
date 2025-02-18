Attribute VB_Name = "wskapiAO"
Option Explicit

#Const WSAPI_CREAR_LABEL = True
Private Const SD_BOTH As Long = &H2
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Const WS_CHILD = &H40000000
Public Const GWL_WNDPROC = (-4)
Private Const SIZE_RCVBUF As Long = 8192
Private Const SIZE_SNDBUF As Long = 8192

Public Type tSockCache
    Sock As Long
    Slot As Long
End Type

Public WSAPISock2Usr  As Collection
Public hWndMsg        As Long
Public SockListen     As Long
Public LastSockListen As Long

Public Sub IniciaWsApi(ByVal hwndParent As Long)
    Call LogApiSock("Iniciando Winsock2 API...")
    Debug.Print "Iniciando Winsock2 API..."
    #If WSAPI_CREAR_LABEL Then
        hWndMsg = CreateWindowEx(0, "STATIC", "AOMSG", WS_CHILD, 0, 0, 0, 0, hwndParent, 0, App.hInstance, ByVal 0&)
    #Else
        hWndMsg = hwndParent
    #End If
    Set frmMain.WinsockThread = New clsSubclass
    Call frmMain.WinsockThread.Hook(hWndMsg)
    Dim Desc As String
    Call StartWinsock(Desc)
End Sub

Public Sub LimpiaWsApi()
    Call LogApiSock("Limpiando informacion de Winsock2 API...")
    If WSAStartedUp Then
        Call EndWinsock
    End If
    Set frmMain.WinsockThread = Nothing
    #If WSAPI_CREAR_LABEL Then
        If hWndMsg <> 0 Then
            Call DestroyWindow(hWndMsg)
        End If
    #End If
End Sub

Public Function BuscaSlotSock(ByVal S As Long) As Long
    On Error GoTo ErrorHandler
    If WSAPISock2Usr.Count <> 0 Then ' GSZAO
        BuscaSlotSock = WSAPISock2Usr.Item(CStr(S))
    Else
        BuscaSlotSock = -1
    End If
    Exit Function
ErrorHandler:
    BuscaSlotSock = -1
End Function

Public Sub AgregaSlotSock(ByVal Sock As Long, ByVal Slot As Long)
    Debug.Print "AgregaSockSlot"
    If WSAPISock2Usr.Count > MaxUsers Then
        Call CloseSocket(Slot)
        Exit Sub
    End If
    Call WSAPISock2Usr.Add(CStr(Slot), CStr(Sock))
End Sub

Public Sub BorraSlotSock(ByVal Sock As Long)
    Dim cant As Long
        cant = WSAPISock2Usr.Count
    On Error Resume Next
    WSAPISock2Usr.Remove CStr(Sock)
    Debug.Print "BorraSockSlot " & cant & " -> " & WSAPISock2Usr.Count
End Sub

Public Function WsApiEnviar(ByVal Slot As Integer, ByRef Str As String) As Long
    Dim Ret     As String
    Dim Retorno As Long
    Dim data()  As Byte
    ReDim Preserve data(Len(Str) - 1) As Byte
    data = StrConv(Str, vbFromUnicode)
    Retorno = 0
    If UserList(Slot).ConnID <> -1 And UserList(Slot).ConnIDValida Then
        Ret = send(ByVal UserList(Slot).ConnID, data(0), ByVal UBound(data()) + 1, ByVal 0)
        If Ret < 0 Then
            Ret = Err.LastDllError
            If Ret = WSAEWOULDBLOCK Then
                Call UserList(Slot).outgoingData.WriteASCIIStringFixed(Str)
            End If
        End If
    ElseIf UserList(Slot).ConnID <> -1 And Not UserList(Slot).ConnIDValida Then
        If Not UserList(Slot).Counters.Saliendo Then
            Retorno = -1
        End If
    End If
    WsApiEnviar = Retorno
End Function

Public Sub LogApiSock(ByVal Str As String)
    On Error GoTo ErrorHandler
    Dim nfile As Integer
        nfile = FreeFile
    Open App.Path & "\logs\wsapi.log" For Append Shared As #nfile
        Print #nfile, Date & " " & time & " " & Str
    Close #nfile
    Exit Sub
ErrorHandler:

End Sub

Public Sub EventoSockAccept(ByVal SockID As Long)
    Dim NewIndex  As Integer
    Dim Ret       As Long
    Dim Tam       As Long
    Dim sa        As sockaddr
    Dim NuevoSock As Long
    Dim i         As Long
    Dim Str       As String
    Dim data()    As Byte
    Tam = sockaddr_size
    Ret = accept(SockID, sa, Tam)
    If Ret = INVALID_SOCKET Then
        i = Err.LastDllError
        Call LogCriticEvent("Error en Accept() API " & i & ": " & GetWSAErrorString(i))
        Exit Sub
    End If
    NuevoSock = Ret
    If setsockopt(NuevoSock, SOL_SOCKET, SO_LINGER, 0, 4) <> 0 Then
        i = Err.LastDllError
        Call LogCriticEvent("Error al setear lingers." & i & ": " & GetWSAErrorString(i))
    End If
    If Not SecurityIp.IpSecurityAceptarNuevaConexion(sa.sin_addr) Then
        Call WSApiCloseSocket(NuevoSock)
        Exit Sub
    End If
    If SecurityIp.IPSecuritySuperaLimiteConexiones(sa.sin_addr) Then
        Str = Protocol.PrepareMessageErrorMsg("Limite de conexiones para su IP alcanzado.")
        ReDim Preserve data(Len(Str) - 1) As Byte
        data = StrConv(Str, vbFromUnicode)
        Call send(ByVal NuevoSock, data(0), ByVal UBound(data()) + 1, ByVal 0)
        Call WSApiCloseSocket(NuevoSock)
        Exit Sub
    End If
    If setsockopt(NuevoSock, SOL_SOCKET, SO_RCVBUFFER, SIZE_RCVBUF, 4) <> 0 Then
        i = Err.LastDllError
        Call LogCriticEvent("Error al setear el tamano del buffer de entrada " & i & ": " & GetWSAErrorString(i))
    End If
    If setsockopt(NuevoSock, SOL_SOCKET, SO_SNDBUFFER, SIZE_SNDBUF, 4) <> 0 Then
        i = Err.LastDllError
        Call LogCriticEvent("Error al setear el tamano del buffer de salida " & i & ": " & GetWSAErrorString(i))
    End If
    NewIndex = NextOpenUser
    If NewIndex <= MaxUsers Then
        Call UserList(NewIndex).incomingData.ReadASCIIStringFixed(UserList(NewIndex).incomingData.Length)
        Call UserList(NewIndex).outgoingData.ReadASCIIStringFixed(UserList(NewIndex).outgoingData.Length)
        UserList(NewIndex).IP = GetAscIP(sa.sin_addr)
        For i = 1 To BanIps.Count
            If BanIps.Item(i) = UserList(NewIndex).IP Then
                Call WriteErrorMsg(NewIndex, "Su IP se encuentra bloqueada en este servidor.")
                Call SecurityIp.IpRestarConexion(sa.sin_addr)
                Call WSApiCloseSocket(NuevoSock)
                Exit Sub
            End If
        Next i
        If NewIndex > LastUser Then LastUser = NewIndex
        UserList(NewIndex).ConnID = NuevoSock
        UserList(NewIndex).ConnIDValida = True
        Call AgregaSlotSock(NuevoSock, NewIndex)
    Else
        Str = Protocol.PrepareMessageErrorMsg("El servidor se encuentra lleno en este momento. Disculpe las molestias ocasionadas.")
        ReDim Preserve data(Len(Str) - 1) As Byte
        data = StrConv(Str, vbFromUnicode)
        Call send(ByVal NuevoSock, data(0), ByVal UBound(data()) + 1, ByVal 0)
        Call WSApiCloseSocket(NuevoSock)
    End If
End Sub

Public Sub EventoSockRead(ByVal Slot As Integer, ByRef Datos() As Byte)
    With UserList(Slot)
        #If AntiExternos Then
            If UserList(Slot).flags.UserLogged Then
                Security.NAC_D_Byte Datos, UserList(Slot).Redundance
            Else
                Security.NAC_D_Byte Datos, 13
            End If
        #End If
        Call .incomingData.WriteBlock(Datos)
        If .ConnID <> -1 Then
            Do While HandleIncomingData(Slot) = True: Loop
        Else
            Exit Sub
        End If
    End With
End Sub

Public Sub EventoSockClose(ByVal Slot As Integer)
    If UserList(Slot).CentinelaUsuario.centinelaIndex <> 0 Then
        Call modCentinela.UsuarioInActivo(Slot)
    End If
    If UserList(Slot).flags.UserLogged Then
        Call CloseSocketSL(Slot)
        Call Cerrar_Usuario(Slot)
    Else
        Call CloseSocket(Slot)
    End If
End Sub

Public Sub WSApiReiniciarSockets()
    Dim i As Long
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    For i = 1 To MaxUsers
        If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida Then
            Call CloseSocket(i)
        End If
    Next i
    For i = 1 To MaxUsers
        Set UserList(i).incomingData = Nothing
        Set UserList(i).outgoingData = Nothing
    Next i
    ReDim UserList(1 To MaxUsers)
    For i = 1 To MaxUsers
        UserList(i).ConnID = -1
        UserList(i).ConnIDValida = False
        
        Set UserList(i).incomingData = New clsByteQueue
        Set UserList(i).outgoingData = New clsByteQueue
    Next i
    LastUser = 1
    NumUsers = 0
    Call LimpiaWsApi
    Call Sleep(100)
    Call IniciaWsApi(frmMain.hWnd)
    SockListen = ListenForConnect(Puerto, hWndMsg, vbNullString)
End Sub

Public Sub WSApiCloseSocket(ByVal Socket As Long)
    Call WSAAsyncSelect(Socket, hWndMsg, ByVal 1025, ByVal (FD_CLOSE))
    Call ShutDown(Socket, SD_BOTH)
End Sub

Public Function CondicionSocket(ByRef lpCallerId As WSABUF, ByRef lpCallerData As WSABUF, ByRef lpSQOS As FLOWSPEC, ByVal Reserved As Long, ByRef lpCalleeId As WSABUF, ByRef lpCalleeData As WSABUF, ByRef Group As Long, ByVal dwCallbackData As Long) As Long
    Dim sa As sockaddr
    If dwCallbackData = 1 Then
        CondicionSocket = CF_REJECT
        Exit Function
    End If
    Call CopyMemory(sa, ByVal lpCallerId.lpBuffer, lpCallerId.dwBufferLen)
    If Not SecurityIp.IpSecurityAceptarNuevaConexion(sa.sin_addr) Then
        CondicionSocket = CF_REJECT
        Exit Function
    End If
    CondicionSocket = CF_ACCEPT
End Function
