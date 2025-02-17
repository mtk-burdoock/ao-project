VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDialogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Const MAX_DIALOGS As Byte = 100
Private Const MAX_LENGTH As Byte = 18
Private Const MS_PER_CHAR As Byte = 100
Private Const MS_ADD_EXTRA As Integer = 5000
Private dialogs(MAX_DIALOGS - 1) As dialog
Private dialogCount As Byte
Private usedFont As StdFont

Private Type dialog
    textLine()  As String
    X           As Integer
    Y           As Integer
    startTime   As Long
    lifeTime    As Long
    CharIndex   As Integer
    Color       As Long
    renderable  As Boolean
    Sube        As Integer
    startTickCount As Long
End Type

Private Sub Class_Initialize()
    dialogCount = 0
End Sub

Private Function BinarySearch(ByVal CharIndex As Integer) As Integer
    Dim min As Long
    Dim max As Long
    Dim mid As Long
    min = 0
    max = dialogCount - 1
    Do While min <= max
        mid = (min + max) \ 2
        If dialogs(mid).CharIndex < CharIndex Then
            min = mid + 1
        ElseIf dialogs(mid).CharIndex > CharIndex Then
            max = mid - 1
        Else
            BinarySearch = mid
            Exit Function
        End If
    Loop
    BinarySearch = Not min
End Function

Public Sub CreateDialog(ByRef chat As String, ByVal CharIndex As Integer, ByVal Color As Long)
    Dim slot As Integer
    slot = BinarySearch(CharIndex)
    If slot < 0 Then
        If dialogCount = MAX_DIALOGS Then Exit Sub
        Dim i As Long
        Dim Count_slot As Long
        slot = Not slot
        Count_slot = slot + 1
        For i = dialogCount To Count_slot Step -1
            dialogs(i) = dialogs(i - 1)
        Next i
        dialogCount = dialogCount + 1
    End If
    With dialogs(slot)
        .CharIndex = CharIndex
        .Color = ColorToDX8(Color)
        .renderable = False
        .startTime = GetTickCount()
        .textLine = FormatChat(chat)
        .lifeTime = MS_ADD_EXTRA + (MS_PER_CHAR * Len(chat))
        .Sube = 12
        .startTickCount = GetTickCount() + 12
    End With
End Sub

Private Function FormatChat(ByRef chat As String) As String()
    Dim word As String
    Dim curPos As Long
    Dim Length As Long
    Dim acumLength As Long
    Dim lineLength As Long
    Dim wordLength As Long
    Dim curLine As Long
    Dim chatLines() As String
    curLine = 0
    curPos = 1
    Length = Len(chat)
    acumLength = 0
    lineLength = -1
    ReDim chatLines(FieldCount(chat, 32)) As String
    Do While acumLength < Length
        word = ReadField(curPos, chat, 32)
        wordLength = Len(word)
        If lineLength = -1 Then
            chatLines(curLine) = word
            lineLength = wordLength
            acumLength = wordLength
        Else
            If lineLength + wordLength + 1 > MAX_LENGTH Then
                curLine = curLine + 1
                chatLines(curLine) = word
                lineLength = wordLength
            Else
                chatLines(curLine) = chatLines(curLine) & " " & word
                lineLength = lineLength + wordLength + 1
            End If
            acumLength = acumLength + wordLength + 1
        End If
        curPos = curPos + 1
    Loop
    If curLine = 0 And Length < MAX_LENGTH Then
        chatLines(curLine) = String$((MAX_LENGTH - Length) \ 2 + 1, " ") & chatLines(curLine)
    End If
    ReDim Preserve chatLines(curLine) As String
    FormatChat = chatLines
End Function

Public Sub UpdateDialogPos(ByVal X As Integer, ByVal Y As Integer, ByVal CharIndex As Integer)
    Dim slot As Integer
    slot = BinarySearch(CharIndex)
    If slot >= 0 Then
        With dialogs(slot)
            .X = X - (MAX_LENGTH \ 2) * 3
            .Y = Y - (UBound(.textLine()) * 3)
            .renderable = True
        End With
    End If
End Sub

Public Sub Render()
    Dim i As Long
    Dim J As Long
    Dim Offset As Integer
    Dim removedDialogs As Long
    Dim Count_dialog As Long
    Count_dialog = dialogCount - 1
    For i = 0 To Count_dialog
        With dialogs(i - removedDialogs)
            If (GetTickCount() - .startTime) >= .lifeTime Then
                Call RemoveDialog(.CharIndex)
                removedDialogs = removedDialogs + 1
            ElseIf .renderable Then
                Dim Upper_textLine As Long
                    Upper_textLine = UBound(.textLine())
                Offset = -(usedFont.Size + 2) * Upper_textLine
                    If GetTickCount() > .startTickCount Then
                        .Sube = .Sube - 1
                        .startTickCount = GetTickCount() + 12
                    End If
                    If .Sube > 0 Then .Y = .Y + .Sube
                For J = 0 To Upper_textLine
                    Call DrawText(.X, .Y + Offset, .textLine(J), .Color)
                    Offset = Offset + usedFont.Size + 5
                Next J
                .renderable = False
            End If
        End With
    Next i
End Sub

Public Sub RemoveDialog(ByVal CharIndex As Integer)
    Dim slot As Integer
    slot = BinarySearch(CharIndex)
    If slot < 0 Then Exit Sub
    Dim i As Long
    Dim Count_maxDialogs As Long
        Count_maxDialogs = MAX_DIALOGS - 2
    For i = slot To Count_maxDialogs
        dialogs(i) = dialogs(i + 1)
    Next i
    dialogCount = dialogCount - 1
End Sub

Public Sub RemoveAllDialogs()
    dialogCount = 0
End Sub

Public Property Get Font() As StdFont
    Font = usedFont
End Property

Public Property Let Font(ByRef nValue As StdFont)
    Set usedFont = nValue
End Property

Public Function NeedRender() As Boolean
    NeedRender = (dialogCount <> 0)
End Function
