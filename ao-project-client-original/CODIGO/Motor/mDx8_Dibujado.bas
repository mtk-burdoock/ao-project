Attribute VB_Name = "mDx8_Dibujado"
Option Explicit

Private Const DAMAGE_TIME As Integer = 1000
Private Const DAMAGE_OFFSET As Integer = 20
Private Const DAMAGE_FONT_S As Byte = 12
 
Private Enum EDType
     edPunal = 1
     edNormal = 2
     edCritico = 3
     edFallo = 4
     edCurar = 5
     edTrabajo = 6
End Enum
 
Private DNormalFont    As New StdFont
 
Type DList
     DamageVal      As Integer
     ColorRGB       As Long
     DamageType     As EDType
     DamageFont     As New StdFont
     StartedTime    As Long
     Downloading    As Byte
     Activated      As Boolean
End Type

Private DrawBuffer As cDIBSection

Sub DrawGrhtoHdc(ByRef Pic As PictureBox, ByVal GrhIndex As Long, ByRef DestRect As RECT)
    DoEvents
    Pic.AutoRedraw = False
    Call Engine_BeginScene
    Call Draw_GrhIndex(GrhIndex, 0, 0, 0, Normal_RGBList())
    Call Engine_EndScene(DestRect, Pic.hwnd)
    Call DrawBuffer.LoadPictureBlt(Pic.hdc)
    Pic.AutoRedraw = True
    Call DrawBuffer.PaintPicture(Pic.hdc, 0, 0, Pic.Width, Pic.Height, 0, 0, vbSrcCopy)
    Pic.Picture = Pic.Image
End Sub

Public Sub PrepareDrawBuffer()
    Set DrawBuffer = New cDIBSection
    Call DrawBuffer.Create(1024, 1024)
End Sub

Public Sub CleanDrawBuffer()
    Set DrawBuffer = Nothing
End Sub

Public Sub CleanPJs()
    Dim LoopC As Long
    For LoopC = 1 To MAX_CHARACTERS
        frmPanelAccount.lblAccData(LoopC - 1).Caption = vbNullString
        frmPanelAccount.picChar(LoopC - 1).AutoRedraw = True
        frmPanelAccount.picChar(LoopC - 1).Refresh
        frmPanelAccount.picChar(LoopC - 1).AutoRedraw = False
        frmPanelAccount.picChar(LoopC - 1).Visible = False
    Next
    DoEvents
End Sub

Public Sub DrawPJ(ByVal Index As Byte)
    If LenB(cPJ(Index).Nombre) = 0 Then Exit Sub
    DoEvents
    frmPanelAccount.picChar(Index - 1).Visible = True
    Dim cColor       As Long
    Dim Head_OffSet  As Integer
    Dim PixelOffsetX As Integer
    Dim PixelOffsetY As Integer
    Dim RE           As RECT
    If cPJ(Index).GameMaster Then
        cColor = 2004510
    Else
        cColor = IIf(cPJ(Index).Criminal, 255, 16744448)
    End If
    With frmPanelAccount.lblAccData(Index)
        .Caption = cPJ(Index).Nombre
        .ForeColor = cColor
    End With
    With frmPanelAccount.picChar(Index - 1)
        RE.Left = 0
        RE.Top = 0
        RE.Bottom = .Height
        RE.Right = .Width
    End With
    PixelOffsetX = RE.Right \ 2 - 16
    PixelOffsetY = RE.Bottom \ 2
    Call Engine_BeginScene
    With cPJ(Index)
        If .Body <> 0 Then
            Call Draw_Grh(BodyData(.Body).Walk(3), PixelOffsetX, PixelOffsetY, 1, Normal_RGBList(), 0)
            If .Head <> 0 Then
                Call Draw_Grh(HeadData(.Head).Head(3), PixelOffsetX + BodyData(.Body).HeadOffset.X, PixelOffsetY + BodyData(.Body).HeadOffset.Y, 1, Normal_RGBList(), 0)
            End If
            If .helmet <> 0 Then
                Call Draw_Grh(CascoAnimData(.helmet).Head(3), PixelOffsetX + BodyData(.Body).HeadOffset.X, PixelOffsetY + BodyData(.Body).HeadOffset.Y + OFFSET_HEAD, 1, Normal_RGBList(), 0)
            End If
            If .weapon <> 0 Then
                Call Draw_Grh(WeaponAnimData(.weapon).WeaponWalk(3), PixelOffsetX, PixelOffsetY, 1, Normal_RGBList(), 0)
            End If
            If .shield <> 0 Then
                Call Draw_Grh(ShieldAnimData(.shield).ShieldWalk(3), PixelOffsetX, PixelOffsetY, 1, Normal_RGBList(), 0)
            End If
        End If
    End With
    Call Engine_EndScene(RE, frmPanelAccount.picChar(Index - 1).hwnd)
    Call DrawBuffer.LoadPictureBlt(frmPanelAccount.picChar(Index - 1).hdc)
    frmPanelAccount.picChar(Index - 1).AutoRedraw = True
    Call DrawBuffer.PaintPicture(frmPanelAccount.picChar(Index - 1).hdc, 0, 0, RE.Right, RE.Bottom, 0, 0, vbSrcCopy)
    frmPanelAccount.picChar(Index - 1).Picture = frmPanelAccount.picChar(Index - 1).Image
End Sub

Sub Damage_Initialize()
    With DNormalFont
        .Size = 20
        .italic = False
        .bold = False
        .Name = "Tahoma"
    End With
End Sub

Sub Damage_Create(ByVal X As Byte, ByVal Y As Byte, ByVal ColorRGB As Long, ByVal DamageValue As Integer, ByVal edMode As Byte)
    With MapData(X, Y).Damage
        .Activated = True
        .ColorRGB = ColorRGB
        .DamageType = edMode
        .DamageVal = DamageValue
        .StartedTime = GetTickCount
        .Downloading = 0
        Select Case .DamageType
            Case EDType.edPunal
                With .DamageFont
                    .Size = Val(DAMAGE_FONT_S)
                    .Name = "Tahoma"
                    .bold = False
                    Exit Sub
                End With
        End Select
        .DamageFont = DNormalFont
        .DamageFont.Size = 14
    End With
End Sub

Private Function EaseOutCubic(Time As Double)
    Time = Time - 1
    EaseOutCubic = Time * Time * Time + 1
End Function
 
Sub Damage_Draw(ByVal X As Byte, ByVal Y As Byte, ByVal PixelX As Integer, ByVal PixelY As Integer)
    With MapData(X, Y).Damage
        If (Not .Activated) Or (Not .DamageVal <> 0) Then Exit Sub
        Dim ElapsedTime As Long
        ElapsedTime = GetTickCount - .StartedTime
        If ElapsedTime < DAMAGE_TIME Then
            .Downloading = EaseOutCubic(ElapsedTime / DAMAGE_TIME) * DAMAGE_OFFSET
            .ColorRGB = Damage_ModifyColour(.DamageType)
            If .DamageType = EDType.edPunal Then
                .DamageFont.Size = Damage_NewSize(ElapsedTime)
            End If
            Select Case .DamageType
                Case EDType.edCritico
                    Call DrawText(PixelX, PixelY - .Downloading, .DamageVal & "!!", .ColorRGB)
                
                Case EDType.edCurar
                    Call DrawText(PixelX, PixelY - .Downloading, "+" & .DamageVal, .ColorRGB)
                
                Case EDType.edTrabajo
                    Call DrawText(PixelX, PixelY - .Downloading, "+" & .DamageVal, .ColorRGB)
                    
                Case EDType.edFallo
                    Call DrawText(PixelX, PixelY - .Downloading, "Fallo", .ColorRGB)
                    
                Case Else
                    Call DrawText(PixelX, PixelY - .Downloading, "-" & .DamageVal, .ColorRGB)
            End Select
        Else
            Damage_Clear X, Y
        End If
    End With
End Sub
 
Sub Damage_Clear(ByVal X As Byte, ByVal Y As Byte)
    With MapData(X, Y).Damage
        .Activated = False
        .ColorRGB = 0
        .DamageVal = 0
        .StartedTime = 0
    End With
End Sub
 
Function Damage_ModifyColour(ByVal DamageType As Byte) As Long
    Select Case DamageType
        Case EDType.edPunal
            Damage_ModifyColour = ColoresDano(52)
            
        Case EDType.edFallo
            Damage_ModifyColour = ColoresDano(54)
            
        Case EDType.edCurar
            Damage_ModifyColour = ColoresDano(55)
        
        Case EDType.edTrabajo
            Damage_ModifyColour = ColoresDano(56)
            
        Case Else
            Damage_ModifyColour = ColoresDano(51)
    End Select
End Function
 
Function Damage_NewSize(ByVal ElapsedTime As Long) As Byte
    Select Case ElapsedTime
        Case Is <= DAMAGE_TIME / 5
            Damage_NewSize = 14
       
        Case Is <= DAMAGE_TIME * 2 / 5
            Damage_NewSize = 13
           
        Case Is <= DAMAGE_TIME * 3 / 5
            Damage_NewSize = 12
           
        Case Else
            Damage_NewSize = 11
    End Select
End Function
