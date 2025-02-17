Attribute VB_Name = "mDx8_Utilities"
Option Explicit

Public Sub Engine_Draw_Line(x1 As Single, y1 As Single, x2 As Single, y2 As Single, Optional Color As Long = -1, Optional Color2 As Long = -1)
    On Error GoTo ErrorHandler
    Call Engine_Long_To_RGB_List(temp_rgb(), Color)
    Call SpriteBatch.SetTexture(Nothing)
    Call SpriteBatch.Draw(x1, y1, x2, y2, temp_rgb())
    Exit Sub
ErrorHandler:

End Sub

Public Sub Engine_Draw_Point(x1 As Single, y1 As Single, Optional Color As Long = -1)
    On Error GoTo ErrorHandler
    Call Engine_Long_To_RGB_List(temp_rgb(), Color)
    Call SpriteBatch.SetTexture(Nothing)
    Call SpriteBatch.Draw(x1, y1, 0, 1, temp_rgb(), 0, 0)
    Exit Sub
ErrorHandler:

End Sub

Public Function Engine_PixelPosX(ByVal X As Integer) As Integer
    Engine_PixelPosX = (X - 1) * 32
End Function

Public Function Engine_PixelPosY(ByVal Y As Integer) As Integer
    Engine_PixelPosY = (Y - 1) * 32
End Function

Public Function Engine_TPtoSPX(ByVal X As Byte) As Long
    Engine_TPtoSPX = Engine_PixelPosX(X - ((UserPos.X - HalfWindowTileWidth) - TileBufferSize)) + OffsetCounterX - 272 + ((10 - TileBufferSize) * 32)
End Function

Public Function Engine_TPtoSPY(ByVal Y As Byte) As Long
    Engine_TPtoSPY = Engine_PixelPosY(Y - ((UserPos.Y - HalfWindowTileHeight) - TileBufferSize)) + OffsetCounterY - 272 + ((10 - TileBufferSize) * 32)
End Function

Public Sub Engine_Draw_Box(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, Color As Long)
    Call Engine_Long_To_RGB_List(temp_rgb(), Color)
    Call SpriteBatch.SetTexture(Nothing)
    Call SpriteBatch.Draw(X, Y, Width, ByVal Height, temp_rgb())
End Sub

Public Sub Engine_D3DColor_To_RGB_List(rgb_list() As Long, Color As D3DCOLORVALUE)
    rgb_list(0) = D3DColorARGB(Color.a, Color.r, Color.g, Color.b)
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

Public Sub Engine_Long_To_RGB_List(rgb_list() As Long, long_color As Long)
    rgb_list(0) = long_color
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

Public Function SetARGB_Alpha(rgb_list() As Long, Alpha As Byte) As Long()
    Dim TempColor        As D3DCOLORVALUE
    Dim tempARGB(0 To 3) As Long
    Call ARGBtoD3DCOLORVALUE(rgb_list(1), TempColor)
    If Alpha > 255 Then Alpha = 255
    If Alpha < 0 Then Alpha = 0
    TempColor.a = Alpha
    Call Engine_D3DColor_To_RGB_List(tempARGB(), TempColor)
    SetARGB_Alpha = tempARGB()
End Function

Private Function Engine_Collision_Between(ByVal Value As Single, ByVal Bound1 As Single, ByVal Bound2 As Single) As Byte
    If Bound1 > Bound2 Then
        If Value >= Bound2 Then
            If Value <= Bound1 Then Engine_Collision_Between = 1
        End If
    Else
        If Value >= Bound1 Then
            If Value <= Bound2 Then Engine_Collision_Between = 1
        End If
    End If
End Function

Public Function Engine_Collision_Line(ByVal L1X1 As Long, ByVal L1Y1 As Long, ByVal L1X2 As Long, ByVal L1Y2 As Long, ByVal L2X1 As Long, ByVal L2Y1 As Long, ByVal L2X2 As Long, ByVal L2Y2 As Long) As Byte
    Dim m1 As Single
    Dim M2 As Single
    Dim b1 As Single
    Dim b2 As Single
    Dim IX As Single
    If L1X1 = L1X2 Then L1X1 = L1X1 + 1
    If L2X1 = L2X2 Then L2X1 = L2X1 + 1
    m1 = (L1Y2 - L1Y1) / (L1X2 - L1X1)
    b1 = L1Y2 - m1 * L1X2
    M2 = (L2Y2 - L2Y1) / (L2X2 - L2X1)
    b2 = L2Y2 - M2 * L2X2
    If M2 - m1 = 0 Then
        If b2 = b1 Then
            Engine_Collision_Line = 1
        Else
            Engine_Collision_Line = 0
        End If
    Else
        IX = ((b2 - b1) / (m1 - M2))
        If Engine_Collision_Between(IX, L1X1, L1X2) Then
            If Engine_Collision_Between(IX, L2X1, L2X2) Then Engine_Collision_Line = 1
        End If
    End If
End Function

Public Function Engine_Collision_LineRect(ByVal sX As Long, ByVal sY As Long, ByVal SW As Long, ByVal SH As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Byte
    If Engine_Collision_Line(sX, sY, sX + SW, sY, x1, y1, x2, y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If
    If Engine_Collision_Line(sX + SW, sY, sX + SW, sY + SH, x1, y1, x2, y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If
    If Engine_Collision_Line(sX, sY + SH, sX + SW, sY + SH, x1, y1, x2, y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If
    If Engine_Collision_Line(sX, sY, sX, sY + SW, x1, y1, x2, y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If
End Function

Function Engine_Collision_Rect(ByVal x1 As Integer, ByVal y1 As Integer, ByVal Width1 As Integer, ByVal Height1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByVal Width2 As Integer, ByVal Height2 As Integer) As Boolean
    If x1 + Width1 >= x2 Then
        If x1 <= x2 + Width2 Then
            If y1 + Height1 >= y2 Then
                If y1 <= y2 + Height2 Then
                    Engine_Collision_Rect = True
                End If
            End If
        End If
    End If
End Function

Public Sub Engine_ZoomIn()
    With MainScreenRect
        .Top = 0
        .Left = 0
        .Bottom = IIf(.Bottom - 1 <= 367, .Bottom, .Bottom - 1)
        .Right = IIf(.Right - 1 <= 491, .Right, .Right - 1)
    End With
End Sub

Public Sub Engine_ZoomOut()
    With MainScreenRect
        .Top = 0
        .Left = 0
        .Bottom = IIf(.Bottom + 1 >= 459, .Bottom, .Bottom + 1)
        .Right = IIf(.Right + 1 >= 583, .Right, .Right + 1)
    End With
End Sub

Public Sub Engine_ZoomNormal()
    With MainScreenRect
        .Top = 0
        .Left = 0
        .Bottom = ScreenHeight
        .Right = ScreenWidth
    End With
End Sub

Public Function ZoomOffset(ByVal Offset As Byte) As Single
    ZoomOffset = IIf((Offset = 1), (ScreenHeight - MainScreenRect.Bottom) / 2, (ScreenWidth - MainScreenRect.Right) / 2)
End Function

Function Engine_Distance(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer) As Long
    Engine_Distance = Abs(x1 - x2) + Abs(y1 - y2)
End Function

Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single
    Dim SideA As Single
    Dim SideC As Single
    On Error GoTo ErrorHandler
    If CenterY = TargetY Then
        If CenterX < TargetX Then
            Engine_GetAngle = 90
        Else
            Engine_GetAngle = 270
        End If
        Exit Function
    End If
    If CenterX = TargetX Then
        If CenterY > TargetY Then
            Engine_GetAngle = 360
        Else
            Engine_GetAngle = 180
        End If
        Exit Function
    End If
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583
    If TargetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle
    Exit Function
ErrorHandler:
    Engine_GetAngle = 0
    Exit Function
End Function
