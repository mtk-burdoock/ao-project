Attribute VB_Name = "Mod_TileEngine"
Option Explicit

Dim temp_verts(3) As TLVERTEX
Public OffsetCounterX As Single
Public OffsetCounterY As Single
Public WeatherFogX1 As Single
Public WeatherFogY1 As Single
Public WeatherFogX2 As Single
Public WeatherFogY2 As Single
Public WeatherFogCount As Byte
Public ParticleOffsetX As Long
Public ParticleOffsetY As Long
Public LastOffsetX As Integer
Public LastOffsetY As Integer
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1
Private Const GrhFogata As Long = 1521
Private Const INFINITE_LOOPS As Integer = -1
Public Const DegreeToRadian As Single = 0.01745329251994

Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type Position
    X As Long
    Y As Long
End Type

Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

Public Type GrhData
    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    Speed As Single
End Type

Public Type Grh
    GrhIndex As Long
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer
    angle As Single
End Type

Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

Public Type Char
    Escribiendot As Byte
    Escribiendo As Boolean
    Movement As Boolean
    active As Byte
    Heading As E_Heading
    Pos As Position
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    fX As Grh
    FxIndex As Integer
    Criminal As Byte
    Atacable As Byte
    Nombre As String
    Clan As String
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    priv As Byte
    attacking As Boolean
    Aura(1 To 4) As Aura
    ParticleIndex As Integer
    Particle_Count As Long
    Particle_Group() As Long
End Type

Public Type obj
    ObjIndex As Integer
    Amount As Integer
End Type

Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    Damage As DList
    NPCIndex As Integer
    OBJInfo As obj
    TileExit As WorldPos
    Blocked As Byte
    Trigger As Integer
    Engine_Light(0 To 3) As Long
    Particle_Group_Index As Long
    fX As Grh
    FxIndex As Integer
End Type

Public Type mapInfo
    Music As String
    Name As String
    Zona As String
    StartPos As WorldPos
    MapVersion As Integer
End Type

Public IniPath As String
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte
Public CurMap As Integer
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position
Public AddtoUserPos As Position
Public UserCharIndex As Integer
Public EngineRun As Boolean
Public FPS As Long
Public FramesPerSecCounter As Long
Public FPSLastCheck As Long
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer
Public HalfWindowTileWidth As Integer
Public HalfWindowTileHeight As Integer
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer
Dim timerElapsedTime As Single
Public timerTicksPerFrame As Single
Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Private MouseTileX As Byte
Private MouseTileY As Byte
Public GrhData() As GrhData
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public MapData() As MapBlock
Public mapInfo As mapInfo
Public Normal_RGBList(3) As Long
Public Color_Shadow(3) As Long
Public Color_Arbol(3) As Long
Public Color_Paralisis As Long
Public Color_Invisibilidad As Long
Public Color_Montura As Long
Public bRain As Boolean
Public bTecho       As Boolean
Public bFogata       As Boolean
Public charlist(1 To 10000) As Char

Private Type Size
    cx As Long
    cy As Long
End Type

Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef TX As Byte, ByRef TY As Byte)
    TX = UserPos.X + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
    TY = UserPos.Y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
On Error Resume Next
    If CharIndex > LastChar Then LastChar = CharIndex
    With charlist(CharIndex)
        If .active = 0 Then _
            NumChars = NumChars + 1
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        .Arma = WeaponAnimData(Arma)
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        .Heading = Heading
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        .Pos.X = X
        .Pos.Y = Y
        .attacking = False
        .active = 1
    End With
    MapData(X, Y).CharIndex = CharIndex
End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Long, Optional ByVal Started As Byte = 2)
    Grh.GrhIndex = GrhIndex
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed
End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
    Dim addx As Integer
    Dim addy As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim nX As Integer
    Dim nY As Integer
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        Select Case nHeading
            Case E_Heading.NORTH
                addy = -1
        
            Case E_Heading.EAST
                addx = 1
        
            Case E_Heading.SOUTH
                addy = 1
            
            Case E_Heading.WEST
                addx = -1
        End Select
        nX = X + addx
        nY = Y + addy
        If nX <= 0 Then nX = 1
        If nY <= 0 Then nY = 1
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
        If (X Or Y) = 0 Then Exit Sub
        MapData(X, Y).CharIndex = 0
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        .Moving = 1
        .Heading = nHeading
        .scrollDirectionX = addx
        .scrollDirectionY = addy
    End With
    If UserEstado = 0 Then Call DoPasosFx(CharIndex)
    If CharIndex <> UserCharIndex Then
        If Not EstaDentroDelArea(nX, nY) Then
            Call Char_Erase(CharIndex)
        End If
    End If
End Sub

Public Sub DoFogataFx()
    Dim Location As Position
    If bFogata Then
        bFogata = HayFogata(Location)
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata(Location)
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", Location.X, Location.Y, LoopStyle.Enabled)
    End If
End Sub

Public Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex).Pos
        EstaPCarea = .X > UserPos.X - MinXBorder And _
                     .X < UserPos.X + MinXBorder And _
                     .Y > UserPos.Y - MinYBorder And _
                     .Y < UserPos.Y + MinYBorder
    End With
End Function

Sub DoPasosFx(ByVal CharIndex As Integer)
    If Not UserNavegando Then
        With charlist(CharIndex)
            If Not .muerto And EstaPCarea(CharIndex) And (.priv = 0 Or .priv > 5) Then
                .pie = Not .pie
                If .pie Then
                    Call Audio.PlayWave(SND_PASOS1, .Pos.X, .Pos.Y)
                Else
                    Call Audio.PlayWave(SND_PASOS2, .Pos.X, .Pos.Y)
                End If
            End If
        End With
    Else
        Call Audio.PlayWave(SND_NAVEGANDO, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y)
    End If
End Sub

Private Function HayFogata(ByRef Location As Position) As Boolean
    Dim J As Long
    Dim k As Long
    For J = UserPos.X - 8 To UserPos.X + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6
            If InMapBounds(J, k) Then
                If MapData(J, k).ObjGrh.GrhIndex = GrhFogata Then
                    Location.X = J
                    Location.Y = k
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next k
    Next J
End Function

Function NextOpenChar() As Integer
    Dim LoopC As Long
    Dim Dale As Boolean
    LoopC = 1
    Do While charlist(LoopC).active And Dale
        LoopC = LoopC + 1
        Dale = (LoopC <= UBound(charlist))
    Loop
    NextOpenChar = LoopC
End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    InMapBounds = True
End Function

Function GetBitmapDimensions(ByVal BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
    Dim BMHeader As BITMAPFILEHEADER
    Dim BINFOHeader As BITMAPINFOHEADER
    Open BmpFile For Binary Access Read As #1
    Get #1, , BMHeader
    Get #1, , BINFOHeader
    Close #1
    bmWidth = BINFOHeader.biWidth
    bmHeight = BINFOHeader.biHeight
End Function

Public Sub DrawTransparentGrhtoHdc(ByVal dsthdc As Long, ByVal srchdc As Long, ByRef SourceRect As RECT, ByRef DestRect As RECT, ByVal TransparentColor As Long)
    Dim Color As Long
    Dim X As Long
    Dim Y As Long
    For X = SourceRect.Left To SourceRect.Right
        For Y = SourceRect.Top To SourceRect.Bottom
            Color = GetPixel(srchdc, X, Y)
            If Color <> TransparentColor Then
                Call SetPixel(dsthdc, DestRect.Left + (X - SourceRect.Left), DestRect.Top + (Y - SourceRect.Top), Color)
            End If
        Next Y
    Next X
End Sub

Public Sub DrawImageInPicture(ByRef PictureBox As PictureBox, ByRef Picture As StdPicture, ByVal x1 As Single, ByVal y1 As Single, Optional Width1, Optional Height1, Optional x2, Optional y2, Optional Width2, Optional Height2)
    Call PictureBox.PaintPicture(Picture, x1, y1, Width1, Height1, x2, y2, Width2, Height2)
End Sub

Sub RenderScreen(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
    On Error GoTo ErrorHandler
    Dim Y                As Long
    Dim X                As Long
    Dim screenminY       As Integer
    Dim screenmaxY       As Integer
    Dim screenminX       As Integer
    Dim screenmaxX       As Integer
    Dim minY             As Integer
    Dim maxY             As Integer
    Dim minX             As Integer
    Dim maxX             As Integer
    Dim ScreenX          As Integer
    Dim ScreenY          As Integer
    Dim minXOffset       As Integer
    Dim minYOffset       As Integer
    Dim PixelOffsetXTemp As Integer
    Dim PixelOffsetYTemp As Integer
    Dim ElapsedTime      As Single
    ElapsedTime = Engine_ElapsedTime()
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize * 2
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
    End If
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
    End If
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    ParticleOffsetX = (Engine_PixelPosX(screenminX) - PixelOffsetX)
    ParticleOffsetY = (Engine_PixelPosY(screenminY) - PixelOffsetY)
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
            PixelOffsetXTemp = (ScreenX - 1) * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = (ScreenY - 1) * TilePixelHeight + PixelOffsetY
            If MapData(X, Y).Graphic(1).GrhIndex <> 0 Then
                Call Draw_Grh(MapData(X, Y).Graphic(1), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, Y).Engine_Light(), 1)
            End If
            If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                Call Draw_Grh(MapData(X, Y).Graphic(2), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, Y).Engine_Light(), 1)
            End If
            ScreenX = ScreenX + 1
        Next
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
            If Map_InBounds(X, Y) Then
                PixelOffsetXTemp = ScreenX * TilePixelWidth + PixelOffsetX
                PixelOffsetYTemp = ScreenY * TilePixelHeight + PixelOffsetY
                With MapData(X, Y)
                    If .ObjGrh.GrhIndex <> 0 Then
                        Call Draw_Grh(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
                    End If
                    If .CharIndex <> 0 Then
                        Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)
                    End If
                    If .Graphic(3).GrhIndex <> 0 Then
                        If .Graphic(3).GrhIndex = 735 Or .Graphic(3).GrhIndex >= 6994 And .Graphic(3).GrhIndex <= 7002 Then
                                If Abs(UserPos.X - X) < 2 And (Abs(UserPos.Y - Y)) < 5 And (Abs(UserPos.Y) < Y) Then
                                Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, Color_Arbol(), 1)
                            Else
                                Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
                            End If
                        Else
                            Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
                        End If
                    End If
                    If .Damage.Activated Then
                        Call mDx8_Dibujado.Damage_Draw(X, Y, PixelOffsetXTemp, PixelOffsetYTemp - 20)
                    End If
                    If .Particle_Group_Index Then
                        If EstaDentroDelArea(X, Y) Then
                            Call mDx8_Particulas.Particle_Group_Render(.Particle_Group_Index, PixelOffsetXTemp + 16, PixelOffsetYTemp + 16)
                        End If
                    End If
                    If Not .FxIndex = 0 Then
                        Call Draw_Grh(.fX, PixelOffsetXTemp + FxData(MapData(X, Y).FxIndex).OffsetX, PixelOffsetYTemp + FxData(.FxIndex).OffsetY, 1, .Engine_Light(), 1, True)
                        If .fX.Started = 0 Then .FxIndex = 0
                    End If
                End With
            End If
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
            PixelOffsetXTemp = ScreenX * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = ScreenY * TilePixelHeight + PixelOffsetY
            If MapData(X, Y).Graphic(4).GrhIndex Then
                If bTecho Then
                    Call Draw_Grh(MapData(X, Y).Graphic(4), PixelOffsetXTemp, PixelOffsetYTemp, 1, temp_rgb(), 1)
                Else
                    If ColorTecho = 250 Then
                        Call Draw_Grh(MapData(X, Y).Graphic(4), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, Y).Engine_Light(), 1)
                    Else
                        Call Draw_Grh(MapData(X, Y).Graphic(4), PixelOffsetXTemp, PixelOffsetYTemp, 1, temp_rgb(), 1)
                    End If
                End If
            End If
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    If ClientSetup.ParticleEngine Then
        Call mDx8_Particulas.Engine_Weather_Update
    End If
    If ClientSetup.ProyectileEngine Then
        If LastProjectile > 0 Then
            Dim J As Long
            For J = 1 To LastProjectile
                If ProjectileList(J).Grh.GrhIndex Then
                    Dim angle As Single
                    angle = DegreeToRadian * Engine_GetAngle(ProjectileList(J).X, ProjectileList(J).Y, ProjectileList(J).TX, ProjectileList(J).TY)
                    ProjectileList(J).X = ProjectileList(J).X + (Sin(angle) * ElapsedTime * 0.8)
                    ProjectileList(J).Y = ProjectileList(J).Y - (Cos(angle) * ElapsedTime * 0.8)
                    If ProjectileList(J).RotateSpeed > 0 Then
                        ProjectileList(J).Rotate = ProjectileList(J).Rotate + (ProjectileList(J).RotateSpeed * ElapsedTime * 0.01)
                        Do While ProjectileList(J).Rotate > 360
                            ProjectileList(J).Rotate = ProjectileList(J).Rotate - 360
                        Loop
                    End If
                    X = ((-minX - 1) * 32) + ProjectileList(J).X + PixelOffsetX + ((10 - TileBufferSize) * 32) - 288 + ProjectileList(J).OffsetX
                    Y = ((-minY - 1) * 32) + ProjectileList(J).Y + PixelOffsetY + ((10 - TileBufferSize) * 32) - 288 + ProjectileList(J).OffsetY
                    If Y >= -32 Then
                        If Y <= (ScreenHeight + 32) Then
                            If X >= -32 Then
                                If X <= (ScreenWidth + 32) Then
                                    If ProjectileList(J).Rotate = 0 Then
                                        Call Draw_Grh(ProjectileList(J).Grh, X, Y, 0, MapData(50, 50).Engine_Light(), 0, True, ProjectileList(J).Rotate + 128)
                                    Else
                                        Call Draw_Grh(ProjectileList(J).Grh, X, Y, 0, MapData(50, 50).Engine_Light(), 0, True, ProjectileList(J).Rotate + 128)
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                End If
            Next J
            For J = 1 To LastProjectile
                If ProjectileList(J).Grh.GrhIndex Then
                    If Abs(ProjectileList(J).X - ProjectileList(J).TX) < 20 Then
                        If Abs(ProjectileList(J).Y - ProjectileList(J).TY) < 20 Then
                            Call Engine_Projectile_Erase(J)
                        End If
                    End If
                End If
            Next J
        End If
    End If
    If colorRender <> 240 Then
        Call DrawText(frmMain.MainViewPic.Width / 2, 50, renderText, render_msg(0), True, 2)
    End If
    LastOffsetX = ParticleOffsetX
    LastOffsetY = ParticleOffsetY
    If ClientSetup.PartyMembers Then Call Draw_Party_Members
    Call RenderCount
ErrorHandler:
    If Err.number Then
        Call LogError(Err.number, Err.Description, "Mod_TileEngine.RenderScreen")
    End If
End Sub

Public Function RenderSounds()
    Dim Location As Position
    If bRain And bLluvia(UserMap) Then
            If bTecho Then
                If frmMain.IsPlaying <> PlayLoop.plLluviain Then
                    If RainBufferIndex Then
                        Call Audio.StopWave(RainBufferIndex)
                    End If
                    RainBufferIndex = Audio.PlayWave("lluviain.wav", 0, 0, LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviain
                End If
            Else
                If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
                    If RainBufferIndex Then
                        Call Audio.StopWave(RainBufferIndex)
                    End If
                    RainBufferIndex = Audio.PlayWave("lluviaout.wav", 0, 0, LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviaout
                End If
            End If
    End If
    If bFogata Then
        bFogata = Map_CheckBonfire(Location)
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = Map_CheckBonfire(Location)
        If bFogata And FogataBufferIndex = 0 Then
            FogataBufferIndex = Audio.PlayWave("fuego.wav", Location.X, Location.Y, LoopStyle.Enabled)
        End If
    End If
End Function

Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Long) As Boolean
    If GrhIndex > 0 Then
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
                And charlist(UserCharIndex).Pos.Y <= Y
    End If
End Function

Public Sub InitTileEngine(ByVal setDisplayFormhWnd As Long, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer)
    On Error GoTo ErrorHandler:
    TileBufferSize = Areas.TilesBuffer
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = Round(frmMain.MainViewPic.Height / 32, 0)
    WindowTileWidth = Round(frmMain.MainViewPic.Width / 32, 0)
    IniPath = Game.path(INIT)
    HalfWindowTileHeight = WindowTileHeight \ 2
    HalfWindowTileWidth = WindowTileWidth \ 2
    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY
    Call CalcularAreas(HalfWindowTileWidth, HalfWindowTileHeight)
    #If UsarGraficosIni = 1 Then
        Call LoadGrhIni
    #Else
        Call LoadGrhInd
    #End If
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
    Call LoadGraphics
    Call CargarParticulas
    Exit Sub
ErrorHandler:
    Call LogError(Err.number, Err.Description, "Mod_TileEngine.InitTileEngine")
    Call CloseClient
End Sub

Public Sub LoadGraphics()
    Call SurfaceDB.Initialize(DirectD3D8, ClientSetup.byMemory)
End Sub

Sub ShowNextFrame(ByVal DisplayFormTop As Integer, ByVal DisplayFormLeft As Integer, ByVal MouseViewX As Integer, ByVal MouseViewY As Integer)
    On Error GoTo ErrorHandler:
    If EngineRun Then
        Call Engine_BeginScene
        Call DesvanecimientoTechos
        Call DesvanecimientoMsg
        If UserMoving Then
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
                End If
            End If
            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
            End If
        End If
        Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)
        If UserCiego Then
            Call DirectDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
        Else
            Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
        End If
        If Dialogos.NeedRender Then Call Dialogos.Render
        Call DibujarCartel
        If DialogosClanes.Activo Then Call DialogosClanes.Draw
        If UserParalizado And UserParalizadoSegundosRestantes > 0 Then
            Call DrawText(4, 10, UserParalizadoSegundosRestantes & " segundos restantes de Paralisis", Color_Paralisis)
        End If
        If UserInvisible And TiempoInvi > 0 Then
            Call DrawText(4, 25, TiempoInvi & " segundos restantes de Invisibilidad", Color_Invisibilidad)
        End If
        If TiempoDopas > 0 Then
            Call DrawText(4, 40, "Tus atributos perderan efecto en " & TiempoDopas & " segundos", Color_Invisibilidad)
        End If
        If Not UserEquitando And UserEquitandoSegundosRestantes > 0 Then
            Call DrawText(4, 55, UserEquitandoSegundosRestantes & " segundos restantes para volver a montarte", Color_Montura)
        End If
        Call Engine_Update_FPS
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * Engine_BaseSpeed
        Call Engine_EndScene(MainScreenRect, 0)
        Call Inventario.DrawDragAndDrop
    End If
ErrorHandler:
    If DirectDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        Call mDx8_Engine.Engine_DirectX8_Init
        Call LoadGraphics
    End If
End Sub

Private Function GetElapsedTime() As Single
    Dim Start_Time As Currency
    Static end_time As Currency
    Static timer_freq As Currency
    If timer_freq = 0 Then
        Call QueryPerformanceFrequency(timer_freq)
    End If
    Call QueryPerformanceCounter(Start_Time)
    GetElapsedTime = (Start_Time - end_time) / timer_freq * 1000
    Call QueryPerformanceCounter(end_time)
End Function

Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
    Dim moved As Boolean
    With charlist(CharIndex)
        If .Moving Then
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
                If .Body.Walk(.Heading).Speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                moved = True
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                If .Body.Walk(.Heading).Speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                moved = True
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        If .attacking And .Arma.WeaponWalk(.Heading).Started = 0 Then
            .Arma.WeaponWalk(.Heading).Started = 1
            .Arma.WeaponWalk(.Heading).FrameCounter = 1
        ElseIf .Arma.WeaponWalk(.Heading).FrameCounter > 4 And .attacking Then
            .attacking = False
        End If
        If Not moved Then
            If Not .Heading <> 0 Then .Heading = EAST
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 0
            If Not .Movement And Not .attacking Then
                .Arma.WeaponWalk(.Heading).Started = 0
                .Arma.WeaponWalk(.Heading).FrameCounter = 0
                .Escudo.ShieldWalk(.Heading).Started = 0
                .Escudo.ShieldWalk(.Heading).FrameCounter = 0
            End If
            .Moving = False
        End If
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        Dim ColorFinal(0 To 3) As Long
        Dim RenderSpell        As Boolean
        If Not .muerto Then
            If Abs(MouseTileX - .Pos.X) < 1 And (Abs(MouseTileY - .Pos.Y)) < 1 And CharIndex <> UserCharIndex And ClientSetup.TonalidadPJ Then
                If Len(.Nombre) > 0 Then
                    If .Criminal Then
                        Call Engine_Long_To_RGB_List(ColorFinal(), D3DColorXRGB(204, 100, 100))
                    Else
                        Call Engine_Long_To_RGB_List(ColorFinal(), D3DColorXRGB(100, 100, 255))
                    End If
                Else
                    ColorFinal(0) = MapData(.Pos.X, .Pos.Y).Engine_Light(0)
                    ColorFinal(1) = MapData(.Pos.X, .Pos.Y).Engine_Light(1)
                    ColorFinal(2) = MapData(.Pos.X, .Pos.Y).Engine_Light(2)
                    ColorFinal(3) = MapData(.Pos.X, .Pos.Y).Engine_Light(3)
                End If
                RenderSpell = True
            Else
                ColorFinal(0) = MapData(.Pos.X, .Pos.Y).Engine_Light(0)
                ColorFinal(1) = MapData(.Pos.X, .Pos.Y).Engine_Light(1)
                ColorFinal(2) = MapData(.Pos.X, .Pos.Y).Engine_Light(2)
                ColorFinal(3) = MapData(.Pos.X, .Pos.Y).Engine_Light(3)
            End If
        Else
            If EsGM(Val(CharIndex)) Then
                Call Engine_Long_To_RGB_List(ColorFinal(), D3DColorARGB(150, 200, 200, 0))
            Else
                If .Criminal Then
                    Call Engine_Long_To_RGB_List(ColorFinal(), D3DColorARGB(100, 255, 100, 100))
                Else
                    Call Engine_Long_To_RGB_List(ColorFinal(), D3DColorARGB(100, 128, 255, 255))
                End If
            End If
        End If
        If Not .invisible Then
            If ClientSetup.UsarSombras Then
                Call RenderSombras(CharIndex, PixelOffsetX, PixelOffsetY)
                Call RenderReflejos(CharIndex, PixelOffsetX, PixelOffsetY)
            End If
            Movement_Speed = 0.5
            If .Body.Walk(.Heading).GrhIndex Then
                Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1)
            End If
            If Len(.Nombre) > 0 Then
                If Nombres Then
                    If .iHead = 0 And .iBody > 0 Then
                        Call RenderName(CharIndex, PixelOffsetX, PixelOffsetY)
                    End If
                    If .Escribiendo Then
                        Call RenderIfCharIsInChatMode(CharIndex, PixelOffsetX, PixelOffsetY)
                    End If
                End If
            End If
            If .Head.Head(.Heading).GrhIndex Then
                Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, ColorFinal(), 0)
            End If
            If .Casco.Head(.Heading).GrhIndex Then
                Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X + 1, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, 1, ColorFinal(), 0)
            End If
            If .Arma.WeaponWalk(.Heading).GrhIndex Then
                Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1)
            End If
            If .Escudo.ShieldWalk(.Heading).GrhIndex Then
                Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1)
            End If
            If ClientSetup.ParticleEngine Then
                Call RenderCharParticles(CharIndex, PixelOffsetX, PixelOffsetY)
            End If
            If LenB(.Nombre) > 0 Then
                If Nombres Then
                    Call RenderName(CharIndex, PixelOffsetX, PixelOffsetY)
                    If .Escribiendo Then
                        Call RenderIfCharIsInChatMode(CharIndex, PixelOffsetX, PixelOffsetY)
                    End If
                End If
            End If
        ElseIf CharIndex = UserCharIndex Or (.Clan <> vbNullString And .Clan = charlist(UserCharIndex).Clan) Then
            If .Body.Walk(.Heading).GrhIndex Then
                Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, True)
            End If
            If .Head.Head(.Heading).GrhIndex Then
                Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, ColorFinal(), 0, True)
            End If
            If .Casco.Head(.Heading).GrhIndex Then
                Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X + 1, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, 1, ColorFinal(), 0, True)
            End If
            If .Arma.WeaponWalk(.Heading).GrhIndex Then
                Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, True)
            End If
            If .Escudo.ShieldWalk(.Heading).GrhIndex Then
                Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, True)
            End If
            If LenB(.Nombre) > 0 Then
                If Nombres Then
                    Call RenderName(CharIndex, PixelOffsetX, PixelOffsetY, True)
                    If .Escribiendo Then
                        Call RenderIfCharIsInChatMode(CharIndex, PixelOffsetX, PixelOffsetY)
                    End If
                End If
            End If
        End If
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, CharIndex)
        Movement_Speed = 1
        If .FxIndex <> 0 Then
            Call Draw_Grh(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, 1, SetARGB_Alpha(MapData(.Pos.X, .Pos.Y).Engine_Light(), 180), 1, True)
            If .fX.Started = 0 Then .FxIndex = 0
        End If
    End With
End Sub

Private Function Puntitos(ByVal CharIndex As Integer) As String
    Dim tActual As Long
    tActual = GetTickCount
    If Abs(tActual - lastTickEscribiendo) > 10 Then
        lastTickEscribiendo = tActual
        charlist(CharIndex).Escribiendot = charlist(CharIndex).Escribiendot + 1
    End If
    Select Case charlist(CharIndex).Escribiendot
        Case 1 To 20
            Puntitos = "."
        Case 20 To 40
            Puntitos = ".."
        Case 40 To 60
            Puntitos = "..."
        Case Else
            charlist(CharIndex).Escribiendot = 1
    End Select
End Function

Private Sub RenderIfCharIsInChatMode(ByVal CharIndex As Long, ByVal X As Integer, ByVal Y As Integer, Optional ByVal Invi As Boolean = False)
    Dim Color As Long
    Color = D3DColorARGB(255, 220, 220, 255)
    With charlist(CharIndex)
        Call DrawText(X + 25, Y - 33, Puntitos(CharIndex), Color, True)
    End With
End Sub

Private Sub RenderSombras(ByVal CharIndex As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
    With charlist(CharIndex)
        If (.iHead > 0) And (.iBody = 617 Or .iBody = 612 Or .iBody = 614 Or .iBody = 616) Then
            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX + 8, PixelOffsetY - 14, 1, Color_Shadow(), 0, False, 187, 1, 1.2)
            Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X + 12, PixelOffsetY + .Body.HeadOffset.Y - 13, 1, Color_Shadow(), 0, False, 187, 1, 1.2)
            Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X + 15, PixelOffsetY + .Body.HeadOffset.Y - 49, 1, Color_Shadow(), 0, False, 195, 1, 1.2)
        ElseIf ((.iHead = 0) And (HayAgua(.Pos.X, .Pos.Y + 1) Or HayAgua(.Pos.X + 1, .Pos.Y) Or HayAgua(.Pos.X, .Pos.Y - 1) Or HayAgua(.Pos.X - 1, .Pos.Y))) Then
            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX + 5, PixelOffsetY - 26, 1, Color_Shadow(), 0, False, 186, 1, 1.33)
        Else
            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX + 8, PixelOffsetY - 11, 1, Color_Shadow(), 0, False, 195, 1, 1.2)
            Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X + 12, PixelOffsetY + .Body.HeadOffset.Y - 10, 1, Color_Shadow(), 0, False, 195, 1, 1.2) ' Shadow Head
            Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X + 18, PixelOffsetY + .Body.HeadOffset.Y - 45, 1, Color_Shadow(), 0, False, 195, 1, 1.2) ' Shadow Helmet
        End If
        If .Arma.WeaponWalk(.Heading).GrhIndex Then
            Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX + 9, PixelOffsetY - 12, 1, Color_Shadow(), 0, False, 195, 1, 1.2)
        End If
        If .Escudo.ShieldWalk(.Heading).GrhIndex Then
            Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX + 9, PixelOffsetY - 12, 1, Color_Shadow(), 0, False, 195, 1, 1.2)
        End If
    End With
End Sub

Private Sub RenderCharParticles(ByVal CharIndex As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
    Dim i As Integer
    With charlist(CharIndex)
        If .Particle_Count > 0 Then
            For i = 1 To .Particle_Count
                If .Particle_Group(i) > 0 Then
                    Call mDx8_Particulas.Particle_Group_Render(.Particle_Group(i), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY)
                End If
            Next i
        End If
    End With
End Sub

Private Sub RenderReflejos(ByVal CharIndex As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
    With charlist(CharIndex)
        Movement_Speed = 0.5
        If HayAgua(.Pos.X, .Pos.Y + 1) Then
            Dim GetInverseHeading As Byte
            Dim ColorFinal(0 To 3) As Long
            Call Engine_Long_To_RGB_List(ColorFinal(), D3DColorARGB(100, 128, 128, 128))
            Select Case .Heading
                Case E_Heading.WEST
                    GetInverseHeading = E_Heading.EAST

                Case E_Heading.EAST
                    GetInverseHeading = E_Heading.WEST

                Case Else
                    GetInverseHeading = .Heading
            End Select
            If .Moving Then
                .Body.Walk(GetInverseHeading).Started = 1
                .Arma.WeaponWalk(GetInverseHeading).Started = 1
                .Escudo.ShieldWalk(GetInverseHeading).Started = 1
            Else
                .Body.Walk(GetInverseHeading).Started = 0
                .Escudo.ShieldWalk(GetInverseHeading).Started = 0
            End If
            If .attacking = False And .Moving = False Then
                .Arma.WeaponWalk(GetInverseHeading).Started = 0
            End If
            If .attacking And .Arma.WeaponWalk(GetInverseHeading).Started = 0 Then
                .Arma.WeaponWalk(GetInverseHeading).Started = 1
                .Arma.WeaponWalk(GetInverseHeading).FrameCounter = 1
                       
            ElseIf .Arma.WeaponWalk(GetInverseHeading).FrameCounter > 4 And .attacking Then
                .attacking = False
            End If
            If Not EsNPC(Val(CharIndex)) Then
                If ((.iHead = 0) Or (.iBody = eCabezas.FRAGATA_FANTASMAL)) Then
                    Call Draw_Grh(.Body.Walk(GetInverseHeading), PixelOffsetX, PixelOffsetY + 80, 1, ColorFinal(), 1, False, 360)
                ElseIf .iBody = 604 Or .iBody = 617 Or .iBody = 612 Or .iBody = 614 Or .iBody = 616 Then
                    If .Heading = E_Heading.SOUTH Or .Heading = E_Heading.NORTH Then
                        Call Draw_Grh(.Body.Walk(GetInverseHeading), PixelOffsetX, PixelOffsetY + 80, 1, ColorFinal(), 1, False, 360)
                        Call Draw_Grh(.Head.Head(GetInverseHeading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + 76, 1, ColorFinal(), 0, False, 360)
                        Call Draw_Grh(.Casco.Head(GetInverseHeading), PixelOffsetX + .Body.HeadOffset.X - 1, PixelOffsetY + .Body.HeadOffset.Y + 116, 1, ColorFinal(), 0, False, 360)
                    Else
                        Call Draw_Grh(.Body.Walk(GetInverseHeading), PixelOffsetX, PixelOffsetY + 70, 1, ColorFinal(), 1, False, 360)
                        Call Draw_Grh(.Head.Head(GetInverseHeading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + 76, 1, ColorFinal(), 0, False, 360)
                        Call Draw_Grh(.Casco.Head(GetInverseHeading), PixelOffsetX + .Body.HeadOffset.X - 1, PixelOffsetY + .Body.HeadOffset.Y + 116, 1, ColorFinal(), 0, False, 360)
                    End If
                Else
                    Call Draw_Grh(.Body.Walk(GetInverseHeading), PixelOffsetX, PixelOffsetY + 44, 1, ColorFinal(), 1, False, 360)
                    Call Draw_Grh(.Head.Head(GetInverseHeading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + 51, 1, ColorFinal(), 1, False, 360)
                    Call Draw_Grh(.Casco.Head(GetInverseHeading), PixelOffsetX + .Body.HeadOffset.X - 1, PixelOffsetY + .Body.HeadOffset.Y + 55, 1, ColorFinal(), 1, False, 360)
                    Call Draw_Grh(.Arma.WeaponWalk(GetInverseHeading), PixelOffsetX, PixelOffsetY + 44, 1, ColorFinal(), 1, False, 360)
                    Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY + 44, 1, ColorFinal(), 0, False, 360)
                End If
            End If
        End If
    End With
End Sub

Private Sub RenderName(ByVal CharIndex As Long, ByVal X As Integer, ByVal Y As Integer, Optional ByVal Invi As Boolean = False)
    Dim Pos   As Integer
    Dim line  As String
    Dim Color As Long
    With charlist(CharIndex)
        Pos = getTagPosition(.Nombre)
        If .priv = 0 Then
            If .muerto Then
                Color = D3DColorARGB(255, 220, 220, 255)
            Else
                If .Criminal Then
                    Color = ColoresPJ(50)
                Else
                    Color = ColoresPJ(49)
                End If
            End If
        Else
            Color = ColoresPJ(.priv)
        End If
    
        If Invi Then
            Color = D3DColorARGB(180, 150, 180, 220)
        End If
        line = Left$(.Nombre, Pos - 2)
        Call DrawText(X + 16, Y + 30, line, Color, True)
        line = mid$(.Nombre, Pos)
        Call DrawText(X + 16, Y + 45, line, Color, True)
    End With
End Sub

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
    With charlist(CharIndex)
        .FxIndex = fX
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)
            .fX.Loops = Loops
        End If
    End With
End Sub

Public Sub Device_Textured_Render(ByVal X As Single, ByVal Y As Single, _
                                  ByVal Width As Integer, ByVal Height As Integer, _
                                  ByVal sX As Integer, ByVal sY As Integer, _
                                  ByVal tex As Long, _
                                  ByRef Color() As Long, _
                                  Optional ByVal Alpha As Boolean = False, _
                                  Optional ByVal angle As Single = 0, _
                                  Optional ByVal ScaleX As Single = 1!, _
                                  Optional ByVal ScaleY As Single = 1!)
    Dim Texture As Direct3DTexture8
    Dim TextureWidth As Long, TextureHeight As Long
    Set Texture = SurfaceDB.GetTexture(tex, TextureWidth, TextureHeight)
    With SpriteBatch
        Call .SetTexture(Texture)
        Call .SetAlpha(Alpha)
        If TextureWidth <> 0 And TextureHeight <> 0 Then
            Call .Draw(X, Y, Width * ScaleX, Height * ScaleY, Color, sX / TextureWidth, sY / TextureHeight, (sX + Width) / TextureWidth, (sY + Height) / TextureHeight, angle)
        Else
            Call .Draw(X, Y, TextureWidth * ScaleX, TextureHeight * ScaleY, Color, , , , , angle)
        End If
    End With
End Sub

Public Sub RenderItem(ByVal hWndDest As Long, ByVal GrhIndex As Long)
    Dim DR As RECT
    With DR
        .Left = 0
        .Top = 0
        .Right = 32
        .Bottom = 32
    End With
    Call Engine_BeginScene
    Call Draw_GrhIndex(GrhIndex, 0, 0, 0, Normal_RGBList(), 0, False)
    Call Engine_EndScene(DR, hWndDest)
End Sub

Sub Draw_GrhIndex(ByVal GrhIndex As Long, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByRef Color_List() As Long, Optional ByVal angle As Single = 0, Optional ByVal Alpha As Boolean = False)
    Dim SourceRect As RECT
    With GrhData(GrhIndex)
        If Center Then
            If .TileWidth <> 1 Then
                X = X - (.pixelWidth - TilePixelWidth) \ 2
            End If
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        Call Device_Textured_Render(X, Y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color_List())
    End With
End Sub

Sub Draw_Grh(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByRef Color_List() As Long, ByVal Animate As Byte, Optional ByVal Alpha As Boolean = False, Optional ByVal angle As Single = 0, Optional ByVal ScaleX As Single = 1!, Optional ByVal ScaleY As Single = 1!)
    Dim CurrentGrhIndex As Long
    Dim FrameDuration As Single
    If Grh.GrhIndex = 0 Then Exit Sub
On Error GoTo ErrorHandler
    If Animate Then
        If Grh.Started = 1 Then
            FrameDuration = Grh.Speed / GrhData(Grh.GrhIndex).NumFrames
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime / FrameDuration) * Movement_Speed
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        ElseIf Grh.FrameCounter > 1 Then
            FrameDuration = Grh.Speed / GrhData(Grh.GrhIndex).NumFrames
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime / FrameDuration) * Movement_Speed
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = 1
            End If
        End If
    End If
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    With GrhData(CurrentGrhIndex)
        If Center Then
            If .TileWidth <> 1 Then
                X = X - (.pixelWidth * ScaleX - TilePixelWidth) \ 2
            End If
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        Call Device_Textured_Render(X, Y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color_List(), Alpha, angle, ScaleX, ScaleY)
    End With
Exit Sub
ErrorHandler:
    If Err.number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Error en el Engine Grafico, Por favor contacte a los adminsitradores enviandoles el archivo Errors.Log que se encuentra el la carpeta del cliente.", vbCritical
        Call CloseClient
    End If
End Sub

Public Function GrhCheck(ByVal GrhIndex As Long) As Boolean
    If GrhIndex > 0 And GrhIndex <= UBound(GrhData()) Then
        GrhCheck = GrhData(GrhIndex).NumFrames
    End If
End Function

Public Sub GrhUninitialize(Grh As Grh)
    With Grh
        .GrhIndex = 0
        .Started = False
        .Loops = 0
        .FrameCounter = 0
        .Speed = 0
    End With
End Sub

Public Sub DesvanecimientoTechos()
    If bTecho Then
        If Not Val(ColorTecho) = 150 Then ColorTecho = ColorTecho - 1
    Else
        If Not Val(ColorTecho) = 250 Then ColorTecho = ColorTecho + 1
    End If
    If Not Val(ColorTecho) = 250 Then
        Call Engine_Long_To_RGB_List(temp_rgb(), D3DColorARGB(ColorTecho, ColorTecho, ColorTecho, ColorTecho))
    End If
End Sub

Public Sub DesvanecimientoMsg()
    Static lastmovement As Long
    If GetTickCount - lastmovement > 1 Then
        lastmovement = GetTickCount
    Else
        Exit Sub
    End If
    If LenB(renderText) Then
        If Not Val(colorRender) = 0 Then colorRender = colorRender - 1
    ElseIf LenB(renderText) = 0 Then
        Exit Sub
    Else
        If Not Val(colorRender) = 240 Then colorRender = colorRender + 1
    End If
    If Not Val(colorRender) = 240 Then
        Call Engine_Long_To_RGB_List(render_msg(), ARGB(255, 255, 255, colorRender))
    End If
    If colorRender = 0 Then renderMsgReset
End Sub

Public Sub renderMsgReset()
    renderFont = 1
    renderText = vbNullString
    nameMap = vbNullString
End Sub
