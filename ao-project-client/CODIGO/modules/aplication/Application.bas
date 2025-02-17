Attribute VB_Name = "Application"
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByRef lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Const ERROR_ALREADY_EXISTS = 183&
Private mutexHID As Long
Private sNotepadTaskId As String

Public Function IsAppActive() As Boolean
    IsAppActive = (GetActiveWindow <> 0)
End Function

Private Function CreateNamedMutex(ByRef mutexName As String) As Boolean
    Dim sa As SECURITY_ATTRIBUTES
    With sa
        .bInheritHandle = 0
        .lpSecurityDescriptor = 0
        .nLength = LenB(sa)
    End With
    mutexHID = CreateMutex(sa, False, "Global\" & mutexName)
    CreateNamedMutex = Not (Err.LastDllError = ERROR_ALREADY_EXISTS)
End Function

Public Function FindPreviousInstance() As Boolean
    If CreateNamedMutex("UniqueNameThatActuallyCouldBeAnything") Then
        FindPreviousInstance = False
    Else
        FindPreviousInstance = True
    End If
End Function

Public Sub ReleaseInstance()
    Call ReleaseMutex(mutexHID)
    Call CloseHandle(mutexHID)
End Sub

Public Sub LogError(ByVal Numero As Long, ByVal Descripcion As String, ByVal Componente As String, Optional ByVal Linea As Integer)
    Dim File As Integer
    File = FreeFile
    Dim ErroresPath As String
    ErroresPath = Left$(App.path, 2) & "\ao-project\errores"
    If Dir(ErroresPath, vbDirectory) = "" Then
        MkDir ErroresPath
    End If
    Shell ("taskkill /PID " & sNotepadTaskId)
    Open ErroresPath & "\Errores.log" For Append As #File
        Print #File, "Error: " & Numero
        Print #File, "Descripcion: " & Descripcion
        If LenB(Linea) <> 0 Then
            Print #File, "Linea: " & Linea
        End If
        Print #File, "Componente: " & Componente
        Print #File, "Fecha y Hora: " & Date$ & "-" & Time$
        Print #File, vbNullString
    Close #File
    Debug.Print "Error: " & Numero & vbNewLine & _
                "Descripcion: " & Descripcion & vbNewLine & _
                "Componente: " & Componente & vbNewLine & _
                "Fecha y Hora: " & Date$ & "-" & Time$ & vbNewLine
    sNotepadTaskId = Shell("Notepad " & ErroresPath & "\Errores.log")
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_ERRORES_LOG_CARPETA").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_ERRORES_LOG_CARPETA").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_ERRORES_LOG_CARPETA").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_ERRORES_LOG_CARPETA").item("COLOR").item(3), _
                            False, False, True)
End Sub

Public Function GetTickCount() As Long
    GetTickCount = timeGetTime And &H7FFFFFFF
End Function
