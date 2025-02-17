Attribute VB_Name = "mVerProcesos"
Option Explicit
 
Public Const TH32CS_SNAPPROCESS As Long = &H2
Public Const MAX_PATH As Integer = 260
Private Const GW_HWNDFIRST = 0&
Private Const GW_HWNDNEXT = 2&
Private Const GW_CHILD = 5&
Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Private CANTv As Byte
 
Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
 
Public Function ListarProcesosUsuario() As String
On Error Resume Next
    Dim hSnapShot As Long
    Dim uProcess As PROCESSENTRY32
    Dim r As Long
    ListarProcesosUsuario = ""
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapShot = 0 Then
        ListarProcesosUsuario = "ERROR"
        Exit Function
    End If
    uProcess.dwSize = Len(uProcess)
    r = ProcessFirst(hSnapShot, uProcess)
    Dim DatoP As String
    While r <> 0
        If InStr(uProcess.szExeFile, ".exe") <> 0 Then
            DatoP = ReadField(1, uProcess.szExeFile, Asc("."))
            ListarProcesosUsuario = ListarProcesosUsuario & "|" & DatoP
     
        End If
        r = ProcessNext(hSnapShot, uProcess)
    Wend
    Call CloseHandle(hSnapShot)
End Function

Public Function ListarCaptionsUsuario() As String
On Error Resume Next
    Dim buf As Long, handle As Long, titulo As String, lenT As Long, ret As Long
    handle = GetWindow(Screen.ActiveForm.hwnd, GW_HWNDFIRST)
    Do While handle <> 0
        If IsWindowVisible(handle) Then
            lenT = GetWindowTextLength(handle)
            If lenT > 0 Then
                titulo = String$(lenT, 0)
                ret = GetWindowText(handle, titulo, lenT + 1)
                titulo$ = Left$(titulo, ret)
                ListarCaptionsUsuario = titulo & "#" & ListarCaptionsUsuario
                CANTv = CANTv + 1
            End If
        End If
        handle = GetWindow(handle, GW_HWNDNEXT)
       Loop
End Function
