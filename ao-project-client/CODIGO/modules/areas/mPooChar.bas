Attribute VB_Name = "mPooChar"
Option Explicit
 
Public Sub Char_Erase(ByVal CharIndex As Integer)
    With charlist(CharIndex)
        If (CharIndex = 0) Then Exit Sub
        If (CharIndex > LastChar) Then Exit Sub
        If Map_InBounds(.Pos.X, .Pos.Y) Then
            MapData(.Pos.X, .Pos.Y).CharIndex = 0
        End If
        If CharIndex = LastChar Then
            Do Until charlist(LastChar).Heading > 0
                LastChar = LastChar - 1
                If LastChar = 0 Then
                    NumChars = 0
                    Exit Sub
                End If
            Loop
        End If
        Call Char_ResetInfo(CharIndex)
        Call Dialogos.RemoveDialog(CharIndex)
        Call Char_Particle_Group_Remove_All(CharIndex)
        NumChars = NumChars - 1
        Exit Sub
    End With
End Sub
 
Private Sub Char_ResetInfo(ByVal CharIndex As Integer)
    With charlist(CharIndex)
        Call Delete_All_Auras(CharIndex)
        Call Char_Particle_Group_Remove_All(CharIndex)
        .active = 0
        .Criminal = 0
        .FxIndex = 0
        .invisible = False
        .Moving = 0
        .muerto = False
        .Nombre = vbNullString
        .Clan = vbNullString
        .pie = False
        .Pos.X = 0
        .Pos.Y = 0
        .UsandoArma = False
        .attacking = False
    End With
End Sub
 
Private Sub Char_MapPosGet(ByVal CharIndex As Long, ByRef X As Byte, ByRef Y As Byte)
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
    End With
End Sub
 
Public Sub Char_MapPosSet(ByVal X As Byte, ByVal Y As Byte)
    If (Map_InBounds(X, Y)) Then
        UserPos.X = X
        UserPos.Y = Y
        MapData(UserPos.X, UserPos.Y).CharIndex = UserCharIndex
        charlist(UserCharIndex).Pos = UserPos
        Exit Sub
    End If
End Sub
 
Public Function Char_Techo() As Boolean
    Char_Techo = False
    With charlist(UserCharIndex)
        If (Map_InBounds(.Pos.X, .Pos.Y)) Then
            If (MapData(.Pos.X, .Pos.Y).Trigger = eTrigger.BAJOTECHO Or MapData(.Pos.X, .Pos.Y).Trigger = eTrigger.CASA) Then
                Char_Techo = True
            End If
        End If
    End With
End Function
 
Public Function Char_MapPosExits(ByVal X As Byte, ByVal Y As Byte) As Integer
    If (Map_InBounds(X, Y)) Then
        Char_MapPosExits = MapData(X, Y).CharIndex
    Else
        Char_MapPosExits = 0
    End If
End Function
 
Public Sub Char_UserPos()
    Dim X As Byte
    Dim Y As Byte
    If Char_Check(UserCharIndex) Then
        Call Char_MapPosGet(UserCharIndex, X, Y)
        bTecho = Char_Techo
        frmMain.Coord.Caption = "Map:" & UserMap & " X:" & X & " Y:" & Y
        Call frmMain.ActualizarMiniMapa
        Exit Sub
    End If
End Sub
 
Public Sub Char_UserIndexSet(ByVal CharIndex As Integer)
    UserCharIndex = CharIndex
    With charlist(UserCharIndex)
        UserPos = .Pos
        Exit Sub
    End With
End Sub
 
Public Function Char_Check(ByVal CharIndex As Integer) As Boolean
    If CharIndex > 0 And CharIndex <= LastChar Then
        With charlist(CharIndex)
            Char_Check = (.Heading > 0)
        End With
    End If
End Function
 
Public Sub Char_SetInvisible(ByVal CharIndex As Integer, ByVal Value As Boolean)
    If Char_Check(CharIndex) Then
        With charlist(CharIndex)
            .invisible = Value
            Exit Sub
        End With
    End If
End Sub
 
Public Sub Char_SetBody(ByVal CharIndex As Integer, ByVal BodyIndex As Integer)
     If BodyIndex < LBound(BodyData()) Or BodyIndex > UBound(BodyData()) Then
        charlist(CharIndex).Body = BodyData(0)
        charlist(CharIndex).iBody = 0
        Exit Sub
    End If
    If Char_Check(CharIndex) Then
        With charlist(CharIndex)
            .Body = BodyData(BodyIndex)
            .iBody = BodyIndex
            Exit Sub
        End With
    End If
End Sub
 
Public Sub Char_SetHead(ByVal CharIndex As Integer, ByVal HeadIndex As Integer)
    If HeadIndex < LBound(HeadData()) Or HeadIndex > UBound(HeadData()) Then
        charlist(CharIndex).Head = HeadData(0)
        charlist(CharIndex).iHead = 0
        Exit Sub
    End If
    If Char_Check(CharIndex) Then
        With charlist(CharIndex)
            .Head = HeadData(HeadIndex)
            .iHead = HeadIndex
            .muerto = (HeadIndex = eCabezas.CASPER_HEAD)
            Exit Sub
        End With
    End If
End Sub
 
Public Sub Char_SetHeading(ByVal CharIndex As Long, ByVal Heading As Byte)
    If Char_Check(CharIndex) Then
         With charlist(CharIndex)
            .Heading = Heading
            Exit Sub
        End With
    End If
End Sub

Public Sub Char_SetName(ByVal CharIndex As Integer, ByVal Name As String)
    If (Len(Name) = 0) Then
        Exit Sub
    End If
    If Char_Check(CharIndex) Then
        With charlist(CharIndex)
            .Nombre = Name
            .Clan = mid$(.Nombre, getTagPosition(.Nombre))
            Exit Sub
        End With
    End If
End Sub
 
Public Sub Char_SetWeapon(ByVal CharIndex As Integer, ByVal WeaponIndex As Integer)
    If WeaponIndex > UBound(WeaponAnimData()) Or WeaponIndex < LBound(WeaponAnimData()) Then
        Exit Sub
    End If
    If Char_Check(CharIndex) Then
        With charlist(CharIndex)
            .Arma = WeaponAnimData(WeaponIndex)
            Exit Sub
        End With
    End If
End Sub
 
Public Sub Char_SetShield(ByVal CharIndex As Integer, ByVal ShieldIndex As Integer)
    If ShieldIndex > UBound(ShieldAnimData()) Or ShieldIndex < LBound(ShieldAnimData()) Then
        Exit Sub
    End If
    If Char_Check(CharIndex) Then
        With charlist(CharIndex)
            .Escudo = ShieldAnimData(ShieldIndex)
            Exit Sub
        End With
    End If
End Sub
 
Public Sub Char_SetCasco(ByVal CharIndex As Integer, ByVal CascoIndex As Integer)
    If CascoIndex > UBound(CascoAnimData()) Or CascoIndex < LBound(CascoAnimData()) Then
        Exit Sub
    End If
    If Char_Check(CharIndex) Then
        With charlist(CharIndex)
            .Casco = CascoAnimData(CascoIndex)
            Exit Sub
        End With
    End If
End Sub
 
Public Sub Char_SetFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
    If (Char_Check(CharIndex)) Then
        With charlist(CharIndex)
            .FxIndex = fX
            If .FxIndex > 0 Then
                Call InitGrh(.fX, FxData(fX).Animacion)
                .fX.Loops = Loops
            End If
        End With
    End If
End Sub
 
Public Sub Char_Make(ByVal CharIndex As Integer, _
                     ByVal Body As Integer, _
                     ByVal Head As Integer, _
                     ByVal Heading As Byte, _
                     ByVal X As Integer, _
                     ByVal Y As Integer, _
                     ByVal Arma As Integer, _
                     ByVal Escudo As Integer, _
                     ByVal Casco As Integer)
    If CharIndex > LastChar Then
        LastChar = CharIndex
    End If
    NumChars = NumChars + 1
    If Arma = 0 Then Arma = 2
    If Escudo = 0 Then Escudo = 2
    If Casco = 0 Then Casco = 2
    With charlist(CharIndex)
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
        .attacking = False
        .Pos.X = X
        .Pos.Y = Y
    End With
    MapData(X, Y).CharIndex = CharIndex
End Sub

Sub Char_MovebyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
    Dim X        As Integer
    Dim Y        As Integer
    Dim addx     As Integer
    Dim addy     As Integer
    Dim nHeading As E_Heading
    If (CharIndex <= 0) Then Exit Sub
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        If Not (Map_InBounds(X, Y)) Then Exit Sub
        MapData(X, Y).CharIndex = 0
        addx = nX - X
        addy = nY - Y
        If Sgn(addx) = 1 Then
            nHeading = E_Heading.EAST
        ElseIf Sgn(addx) = -1 Then
            nHeading = E_Heading.WEST
        ElseIf Sgn(addy) = -1 Then
            nHeading = E_Heading.NORTH
        ElseIf Sgn(addy) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        .Moving = 1
        .Heading = nHeading
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)
        If .FxIndex = FxMeditar.CHICO Or _
           .FxIndex = FxMeditar.GRANDE Or _
           .FxIndex = FxMeditar.MEDIANO Or _
           .FxIndex = FxMeditar.XGRANDE Or _
           .FxIndex = FxMeditar.XXGRANDE Then
            .FxIndex = 0
        End If
    End With
    If Not EstaPCarea(CharIndex) Then Call Dialogos.RemoveDialog(CharIndex)
    If Not EstaDentroDelArea(nX, nY) Then
        Call Char_Erase(CharIndex)
    End If
End Sub

Sub Char_MoveScreen(ByVal nHeading As E_Heading)
    Dim X  As Integer
    Dim Y  As Integer
    Dim TX As Integer
    Dim TY As Integer
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
        
        Case E_Heading.EAST
            X = 1
        
        Case E_Heading.SOUTH
            Y = 1
        
        Case E_Heading.WEST
            X = -1
    End Select
    TX = UserPos.X + X
    TY = UserPos.Y + Y
    If (TX < MinXBorder) Or (TX > MaxXBorder) Or (TY < MinYBorder) Or (TY > MaxYBorder) Then
        Exit Sub
    Else
        AddtoUserPos.X = X
        UserPos.X = TX
        AddtoUserPos.Y = Y
        UserPos.Y = TY
        UserMoving = 1
        bTecho = Char_Techo
    End If
End Sub

Sub Char_MovebyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
    Dim addx As Integer
    Dim addy As Integer
    Dim X    As Integer
    Dim Y    As Integer
    Dim nX   As Integer
    Dim nY   As Integer
    If (CharIndex <= 0) Then Exit Sub
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
        If Not (Map_InBounds(nX, nY)) Then Exit Sub
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
        MapData(X, Y).CharIndex = 0
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        .Moving = 1
        .Heading = nHeading
        .scrollDirectionX = addx
        .scrollDirectionY = addy
    End With
    If (UserEstado = 0) Then
        Call DoPasosFx(CharIndex)
    End If
    If CharIndex <> UserCharIndex Then
        If Not EstaDentroDelArea(nX, nY) Then
            Call Char_Erase(CharIndex)
        End If
    End If
End Sub

Sub Char_CleanAll()
    Dim X         As Long, Y As Long
    Dim CharIndex As Integer, obj As Integer
    For X = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
            CharIndex = Char_MapPosExits(CByte(X), CByte(Y))
            If (CharIndex > 0) Then
                Call Char_Erase(CharIndex)
            End If
            obj = Map_PosExitsObject(CByte(X), CByte(Y))
            If (obj > 0) Then
                Call Map_DestroyObject(CByte(X), CByte(Y))
            End If
        Next Y
    Next X
End Sub
