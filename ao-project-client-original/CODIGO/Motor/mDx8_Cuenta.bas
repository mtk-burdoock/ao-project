Attribute VB_Name = "mDx8_Cuenta"
Option Explicit

Public Type Count
    min As Byte
    TickCount As Long
    DoIt As Boolean
End Type

Public DX_Count As Count

Public Sub RenderCount()
    If DX_Count.DoIt = False Then Exit Sub
    If DX_Count.min <> 0 Then
    Else
    
    End If
    Call CheckCount
End Sub

Public Sub CheckCount()
    If DX_Count.DoIt = False Then Exit Sub
        If GetTickCount - DX_Count.TickCount > 1000 Then
            If DX_Count.min > 0 Then
                DX_Count.min = DX_Count.min - 1
                DX_Count.TickCount = GetTickCount
            ElseIf DX_Count.min = 0 Then
                DX_Count.min = 0
                DX_Count.DoIt = False
            End If
        End If
End Sub

Public Sub InitCount(ByVal max As Byte)
    If DX_Count.DoIt = True Then Exit Sub
    With DX_Count
        .min = max
        .TickCount = GetTickCount
        .DoIt = True
    End With
End Sub
