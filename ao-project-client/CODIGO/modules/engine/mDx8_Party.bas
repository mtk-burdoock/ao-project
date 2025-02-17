Attribute VB_Name = "mDx8_Party"
Option Explicit

Public Type c_PartyMember
    Name As String
    Head As Integer
    Lvl As Byte
    ExpParty As Long
End Type

Public PartyMembers(1 To 5) As c_PartyMember

Public Sub Reset_Party()
    Dim i As Byte
        For i = 1 To 5
            PartyMembers(i).ExpParty = 0
            PartyMembers(i).Head = 0
            PartyMembers(i).Lvl = 0
            PartyMembers(i).Name = vbNullString
        Next i
End Sub

Public Sub Draw_Party_Members()
    Dim i As Byte, Count As Byte
    Count = 0
    For i = 1 To 5
        If Len(PartyMembers(i).Name) > 0 Then
            Count = Count + 1
            Call Engine_Draw_Box(410, 20 + (Count - 1) * 50 + 5, 120, 40, D3DColorARGB(100, 0, 0, 0))
            Call Draw_GrhIndex(HeadData(PartyMembers(i).Head).Head(3).GrhIndex, 410, 20 + (Count - 1) * 50 + 35, 1, Normal_RGBList(), 0, True)
        End If
    Next i
    If Count <> 0 Then
    End If
End Sub

Public Sub Set_PartyMember(ByVal Member As Byte, Name As String, ExpParty As Long, Lvl As Byte, Head As Integer)
    If Member < 1 Or Member > 5 Then Exit Sub
        With PartyMembers(Member)
            .Name = Name
            .ExpParty = ExpParty
            .Head = Head
            .Lvl = Lvl
        End With
End Sub

Public Sub Kick_PartyMember(ByVal Member As Byte)
    If Member < 1 Or Member > 5 Then Exit Sub
        With PartyMembers(Member)
            .Name = vbNullString
            .ExpParty = 0
            .Head = 0
            .Lvl = 0
        End With
End Sub
