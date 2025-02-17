VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Const MESSAGE_COUNT As Byte = 10
Private Const MAX_MESSAGE_LENGTH As Byte = 120
Private Const INVALID_MESSAGE As Long = vbObjectError + 123
Private Const MESSAGE_FILE_PATH As String = "/init/messages.txt"
Private CustomMessages(MESSAGE_COUNT - 1) As String

Private Sub Class_Initialize()
On Error GoTo ErrorHandler
    Dim handle As Integer
    Dim i As Long
    handle = FreeFile()
    Open App.path & MESSAGE_FILE_PATH For Input As handle
    Do While (i < MESSAGE_COUNT) And (Not EOF(handle))
        Line Input #handle, CustomMessages(i)
        If Not MessageIsValid(CustomMessages(i)) Then
            CustomMessages(i) = vbNullString
        End If
        i = i + 1
    Loop
    Close handle
ErrorHandler:

End Sub

Private Sub Class_Terminate()
On Error Resume Next
    Dim handle As Integer
    Dim i As Long
    Dim directories() As String
    Dim Upper_directories As Long
    Dim Count_messages As Long
    Dim path As String
    directories = Split(MESSAGE_FILE_PATH, "/")
    path = App.path
    Upper_directories = UBound(directories()) - 1
    For i = 0 To Upper_directories
        path = path & "/" & directories(i)
        If LenB(Dir$(path, vbDirectory)) = 0 Then
            Call MkDir(path)
        End If
    Next i
    handle = FreeFile()
    Count_messages = MESSAGE_COUNT - 1
    Open App.path & MESSAGE_FILE_PATH For Output As handle
    For i = 0 To Count_messages
        Print #handle, CustomMessages(i)
    Next i
    Close handle
    Close handle
End Sub

Public Property Get Message(ByVal Index As Integer) As String
    If Index < 0 Or Index >= MESSAGE_COUNT Then
        Exit Property
    End If
    Message = CustomMessages(Index)
End Property

Public Property Let Message(ByVal Index As Integer, ByVal nValue As String)
    If Index < 0 Or Index >= MESSAGE_COUNT Then
        Exit Property
    End If
    If Not MessageIsValid(nValue) Then
        Call Err.Raise(INVALID_MESSAGE)
        Exit Property
    End If
    CustomMessages(Index) = nValue
End Property

Public Property Get InvalidMessageErrCode() As Long
    InvalidMessageErrCode = INVALID_MESSAGE
End Property

Private Function MessageIsValid(ByRef msg As String) As Boolean
    Dim i As Long
    Dim Char As String * 1
    Dim ascii As Byte
    Dim Len_msg As Long
    msg = Trim$(msg)
    If LenB(msg) = 0 Then
        MessageIsValid = True
        Exit Function
    End If
    Len_msg = Len(msg)
    For i = 1 To Len_msg
        Char = mid$(msg, i, 1)
        ascii = Asc(Char)
        If ascii < 31 Then Exit Function
        If ascii > 126 And Char <> "á" And Char <> "Á" And Char <> "ä" And Char <> "Ä" _
          And Char <> "â" And Char <> "Â" And Char <> "à" And Char <> "À" And Char <> "é" _
          And Char <> "É" And Char <> "ë" And Char <> "Ë" And Char <> "ê" And Char <> "Ê" _
          And Char <> "è" And Char <> "È" And Char <> "í" And Char <> "Í" And Char <> "ï" _
          And Char <> "Ï" And Char <> "î" And Char <> "Î" And Char <> "ì" And Char <> "Ì" _
          And Char <> "ó" And Char <> "Ó" And Char <> "ö" And Char <> "Ö" And Char <> "ô" _
          And Char <> "Ô" And Char <> "ò" And Char <> "Ò" And Char <> "ú" And Char <> "Ú" _
          And Char <> "ü" And Char <> "Ü" And Char <> "û" And Char <> "Û" And Char <> "ù" _
          And Char <> "Ù" And Char <> "ñ" And Char <> "Ñ" And Char <> "¡" And Char <> "¿" Then
                        Exit Function
        End If
    Next i
    If Len(msg) > MAX_MESSAGE_LENGTH Then Exit Function
    MessageIsValid = True
End Function
