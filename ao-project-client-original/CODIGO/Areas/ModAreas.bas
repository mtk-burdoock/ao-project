Attribute VB_Name = "Areas"
Option Explicit

Public Const TilesBuffer As Byte = 5
Private AreasX As Byte
Private AreasY As Byte
Private CurAreaX As Integer
Private CurAreaY As Integer

Public Sub CalcularAreas(HalfWindowTileWidth As Integer, HalfWindowTileHeight As Integer)
    AreasX = HalfWindowTileWidth + TileBufferSize
    AreasY = HalfWindowTileHeight + TileBufferSize
End Sub

Public Sub CambioDeArea(ByVal X As Byte, ByVal Y As Byte)
    CurAreaX = X \ AreasX
    CurAreaY = Y \ AreasY
    Dim loopX As Integer, loopY As Integer, CharIndex As Integer
    For loopX = 1 To 100
        For loopY = 1 To 100
            If Not EstaDentroDelArea(loopX, loopY) Then
                CharIndex = Char_MapPosExits(loopX, loopY)
                If (CharIndex > 0) Then
                    If (CharIndex <> UserCharIndex) Then
                        Call Char_Erase(CharIndex)
                    End If
                End If
                If (Map_PosExitsObject(loopX, loopY) > 0) Then
                    Call Map_DestroyObject(loopX, loopY)
                End If
            End If
        Next loopY
    Next loopX
End Sub

Public Function EstaDentroDelArea(ByVal X As Integer, ByVal Y As Integer) As Boolean
    EstaDentroDelArea = (Abs(CurAreaX - X \ AreasX) <= 1) And (Abs(CurAreaY - Y \ AreasY) <= 1)
End Function
