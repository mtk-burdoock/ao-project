Attribute VB_Name = "Mod_General"
Option Explicit

Public bFogata As Boolean

Public Type tRedditPost
    Title As String
    URL As String
End Type

Public Posts() As tRedditPost
Public bLluvia() As Byte
Private lFrameTimer As Long
Private keysMovementPressedQueue As clsArrayList

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    Randomize Timer
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Public Function GetRawName(ByRef sName As String) As String
    Dim Pos As Integer
    Pos = InStr(1, sName, "<")
    If Pos > 0 Then
        GetRawName = Trim$(Left$(sName, Pos - 1))
    Else
        GetRawName = sName
    End If
End Function

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, _
                    ByVal Text As String, _
                    Optional ByVal Red As Integer = -1, _
                    Optional ByVal Green As Integer, _
                    Optional ByVal Blue As Integer, _
                    Optional ByVal bold As Boolean = False, _
                    Optional ByVal italic As Boolean = False, _
                    Optional ByVal bCrLf As Boolean = True, _
                    Optional ByVal Alignment As Byte = rtfLeft)
    With RichTextBox
        If Len(.Text) > 1000 Then
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        .SelAlignment = Alignment
        If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        .SelText = Text
        If Not (RichTextBox = frmMain.RecTxt) Then
            RichTextBox.Refresh
        End If
    End With
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    Dim Len_cad As Long
    cad = LCase$(cad)
    Len_cad = Len(cad)
    For i = 1 To Len_cad
        car = Asc(mid$(cad, i, 1))
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    AsciiValidos = True
End Function

Function CheckUserData() As Boolean
    Dim LoopC As Long
    Dim CharAscii As Integer
    Dim Len_accountName As Long, Len_accountPassword As Long
    If LenB(AccountPassword) = 0 Then
        MsgBox JsonLanguage.item("VALIDACION_PASSWORD").item("TEXTO")
        Exit Function
    End If
    Len_accountPassword = Len(AccountPassword)
    For LoopC = 1 To Len_accountPassword
        CharAscii = Asc(mid$(AccountPassword, LoopC, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox Replace$(JsonLanguage.item("VALIDACION_BAD_PASSWORD").item("TEXTO").item(2), "VAR_CHAR_INVALIDO", Chr$(CharAscii))
            Exit Function
        End If
    Next LoopC
    If Len(AccountName) > 30 Then
        MsgBox JsonLanguage.item("VALIDACION_BAD_EMAIL").item("TEXTO").item(2)
        Exit Function
    End If
    Len_accountName = Len(AccountName)
    For LoopC = 1 To Len_accountName
        CharAscii = Asc(mid$(AccountName, LoopC, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox Replace$(JsonLanguage.item("VALIDACION_BAD_PASSWORD").item("TEXTO").item(4), "VAR_CHAR_INVALIDO", Chr$(CharAscii))
            Exit Function
        End If
    Next LoopC
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next
    Dim mifrm As Form
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    If KeyAscii > 126 Then
        Exit Function
    End If
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    LegalCharacter = True
End Function

Sub SetConnected()
    Connected = True
    Unload frmCrearPersonaje
    Unload frmConnect
    Unload frmPanelAccount
    keysMovementPressedQueue.Clear
    frmMain.lblName.Caption = UserName
    frmMain.Visible = True
    Call frmMain.ControlSM(eSMType.sResucitation, False)
    Call frmMain.ControlSM(eSMType.mWork, False)
    Call frmMain.ControlSM(eSMType.mSpells, False)
    Call frmMain.ControlSM(eSMType.sSafemode, False)
    frmMain.SendTxt.Visible = False
    Typing = False
    FPSFLAG = True
End Sub

Sub RandomMove()
    Call Map_MoveTo(RandomNumber(NORTH, WEST))
End Sub

Private Sub AddMovementToKeysMovementPressedQueue()
    If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyUp)) = False Then keysMovementPressedQueue.Add (CustomKeys.BindedKey(eKeyType.mKeyUp)) ' Agrega la tecla al arraylist
    Else
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyUp)) Then keysMovementPressedQueue.Remove (CustomKeys.BindedKey(eKeyType.mKeyUp)) ' Remueve la tecla que teniamos presionada
    End If
    If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyDown)) = False Then keysMovementPressedQueue.Add (CustomKeys.BindedKey(eKeyType.mKeyDown)) ' Agrega la tecla al arraylist
    Else
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyDown)) Then keysMovementPressedQueue.Remove (CustomKeys.BindedKey(eKeyType.mKeyDown)) ' Remueve la tecla que teniamos presionada
    End If
    If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyLeft)) = False Then keysMovementPressedQueue.Add (CustomKeys.BindedKey(eKeyType.mKeyLeft)) ' Agrega la tecla al arraylist
    Else
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyLeft)) Then keysMovementPressedQueue.Remove (CustomKeys.BindedKey(eKeyType.mKeyLeft)) ' Remueve la tecla que teniamos presionada
    End If
    If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyRight)) = False Then keysMovementPressedQueue.Add (CustomKeys.BindedKey(eKeyType.mKeyRight)) ' Agrega la tecla al arraylist
    Else
        If keysMovementPressedQueue.itemExist(CustomKeys.BindedKey(eKeyType.mKeyRight)) Then keysMovementPressedQueue.Remove (CustomKeys.BindedKey(eKeyType.mKeyRight)) ' Remueve la tecla que teniamos presionada
    End If
End Sub

Private Sub CheckKeys()
    Static lastmovement As Long
    Static lastmsg As Long
    Dim bCantMove As Boolean
    Dim intentemoverme As Boolean
    If Not Application.IsAppActive() Then Exit Sub
    If Comerciando Then Exit Sub
    If MirandoForo Then Exit Sub
    If pausa Then Exit Sub
    If Traveling Then bCantMove = True
    If EsGM(UserCharIndex) Then
        If frmMain.SendTxt.Visible Then Exit Sub
        If frmMain.SendCMSTXT.Visible Then Exit Sub
    End If
    If UserMoving = 0 Then
        If Not UserEstupido Then
            Call AddMovementToKeysMovementPressedQueue
            If keysMovementPressedQueue.GetLastItem() = CustomKeys.BindedKey(eKeyType.mKeyUp) Then
                If bCantMove Then
                    intentemoverme = True
                    GoTo CantMove
                End If
                Call Map_MoveTo(NORTH)
                Call Char_UserPos
                Exit Sub
            End If
            If keysMovementPressedQueue.GetLastItem() = CustomKeys.BindedKey(eKeyType.mKeyRight) Then
                If bCantMove Then
                    intentemoverme = True
                    GoTo CantMove
                End If
                Call Map_MoveTo(EAST)
                Call Char_UserPos
                Exit Sub
            End If
            If keysMovementPressedQueue.GetLastItem() = CustomKeys.BindedKey(eKeyType.mKeyDown) Then
                If bCantMove Then
                    intentemoverme = True
                    GoTo CantMove
                End If
                Call Map_MoveTo(SOUTH)
                Call Char_UserPos
                Exit Sub
            End If
            If keysMovementPressedQueue.GetLastItem() = CustomKeys.BindedKey(eKeyType.mKeyLeft) Then
                If bCantMove Then
                    intentemoverme = True
                    GoTo CantMove
                End If
                Call Map_MoveTo(WEST)
                Call Char_UserPos
                Exit Sub
            End If
            Call Audio.MoveListener(UserPos.X, UserPos.Y)
        Else
            If bCantMove Then
                intentemoverme = True
                GoTo CantMove
            End If
            Dim kp As Boolean
            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            If kp Then
                Call RandomMove
            Else
                Call Audio.MoveListener(UserPos.X, UserPos.Y)
            End If
            Call Char_UserPos
        End If
    End If
CantMove:
    If bCantMove And intentemoverme Then
        If (Abs(GetTickCount - lastmsg)) > 1000 Then
            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USUARIO_VIAJANDO_HOGAR").item("TEXTO"), 110, 220, 0)
            lastmsg = Abs(GetTickCount)
        End If
    End If
End Sub

Sub SwitchMap(ByVal Map As Integer)
    Dim Y        As Long
    Dim X        As Long
    Dim ByFlags  As Byte
    Dim handle   As Integer
    Dim fileBuff As clsByteBuffer
    Dim dData()  As Byte
    Dim dLen     As Long
    Set fileBuff = New clsByteBuffer
    Call Char_CleanAll
    Call Particle_Group_Remove_All
    dLen = FileLen(Game.path(Mapas) & "Mapa" & Map & ".map")
    ReDim dData(dLen - 1)
    handle = FreeFile()
    Open Game.path(Mapas) & "Mapa" & Map & ".map" For Binary As handle
        Get handle, , dData
    Close handle
    fileBuff.initializeReader dData
    mapInfo.MapVersion = fileBuff.getInteger
    
    With MiCabecera
        .Desc = fileBuff.getString(Len(.Desc))
        .CRC = fileBuff.getLong
        .MagicWord = fileBuff.getLong
    End With
    
    fileBuff.getDouble
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            ByFlags = fileBuff.getByte()
            With MapData(X, Y)
                .Blocked = (ByFlags And 1)
                .Graphic(1).GrhIndex = fileBuff.getLong()
                Call InitGrh(.Graphic(1), .Graphic(1).GrhIndex)
                If ByFlags And 2 Then
                    .Graphic(2).GrhIndex = fileBuff.getLong()
                    Call InitGrh(.Graphic(2), .Graphic(2).GrhIndex)
                Else
                    .Graphic(2).GrhIndex = 0
                End If
                If ByFlags And 4 Then
                    .Graphic(3).GrhIndex = fileBuff.getLong()
                    Call InitGrh(.Graphic(3), .Graphic(3).GrhIndex)
                Else
                    .Graphic(3).GrhIndex = 0
                End If
                If ByFlags And 8 Then
                    .Graphic(4).GrhIndex = fileBuff.getLong()
                    Call InitGrh(.Graphic(4), .Graphic(4).GrhIndex)
                Else
                    .Graphic(4).GrhIndex = 0
                End If
                If ByFlags And 16 Then
                    .Trigger = fileBuff.getInteger()
                Else
                    .Trigger = 0
                End If
                If ByFlags And 32 Then
                    Call General_Particle_Create(CLng(fileBuff.getInteger()), X, Y)
                Else
                    .Particle_Group_Index = 0
                End If
                If .CharIndex > 0 Then
                    .CharIndex = 0
                End If
                If .ObjGrh.GrhIndex > 0 Then
                    .ObjGrh.GrhIndex = 0
                End If
                Call Engine_D3DColor_To_RGB_List(.Engine_Light(), Estado_Actual)
            End With
        Next X
    Next Y
    Call LightRemoveAll
    Call mDx8_Particulas.RemoveWeatherParticles(eWeather.Rain)
    Set fileBuff = Nothing
    
    With mapInfo
        .Name = vbNullString
        .Music = vbNullString
    End With

    If FileExist(Game.path(Graficos) & "MiniMapa\" & Map & ".bmp", vbArchive) Then
        frmMain.MiniMapa.Picture = LoadPicture(Game.path(Graficos) & "MiniMapa\" & Map & ".bmp")
    Else
        frmMain.MiniMapa.Visible = False
        frmMain.RecTxt.Width = frmMain.RecTxt.Width + 100
    End If
    CurMap = Map
    Call Init_Ambient(Map)
    Call Load_Map_Particles(Map)
    renderText = nameMap
    renderFont = 2
    colorRender = 240
    frmMain.lblMapName.Caption = nameMap
End Sub

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    delimiter = Chr$(SepASCII)
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
    Dim Count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    If LenB(Text) = 0 Then Exit Function
    delimiter = Chr$(SepASCII)
    curPos = 0
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    FieldCount = Count
End Function

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(File, FileType) <> "")
End Function

Private Function GetCountryFromIp(ByVal Ip As String) As String
On Error Resume Next
    Dim URL As String
    Dim Endpoint As String
    Dim JsonObject As Object
    Dim Response As String
    Set Inet = New clsInet
    URL = GetVar(Game.path(INIT) & "Config.ini", "Parameters", "IpApiEndpoint")
    Endpoint = URL & Ip & "/json/"
    Response = Inet.OpenRequest(Endpoint, "GET")
    Response = Inet.Execute
    Response = Inet.GetResponseAsString
    Set JsonObject = JSON.parse(Response)
    GetCountryFromIp = JsonObject.item("country")
    Set Inet = Nothing
End Function

Sub Main()
    IPdelServidor = "10.1.74.145"
    PuertoDelServidor = 7666
    Static lastFlush As Long
    Call SetLanguageApplication
    Call Game.LeerConfiguracion
    Call modCompression.GenerateContra(vbNullString, 0)
    Call CargarHechizos
    Set Sonidos = New clsSoundMapas
    Call Sonidos.LoadSoundMapInfo
    Call LeerLineaComandos
    Call SaveSetting("ArgentumOnlineCliente", "Init", "Path", App.path & "\")
    ChDrive App.path
    ChDir App.path
    Call Resolution.SetResolution(1024, 768)
    Call LoadInitialConfig
    If GetVar(Game.path(INIT) & "Config.ini", "Parameters", "TestMode") <> 1 Then
        frmPres.Show vbModal
    End If
    frmConnect.Visible = True
    prgRun = True
    pausa = False
    LoadTimerIntervals
    Set DialogosClanes = New clsGuildDlg
    DialogosClanes.Activo = ClientSetup.bGldMsgConsole
    DialogosClanes.CantidadDialogos = ClientSetup.bCantMsgs
    DialogosClanes.Font = frmMain.Font
    Dialogos.Font = frmMain.Font
    lFrameTimer = GetTickCount
    Call Load(frmScreenshots)
    Do While prgRun
        If frmMain.WindowState <> 1 And frmMain.Visible Then
            Call ShowNextFrame(frmMain.Top, frmMain.Left, frmMain.MouseX, frmMain.MouseY)
            Call RenderSounds
            Call CheckKeys
        End If
        If GetTickCount - lFrameTimer >= 1000 Then
            If FPSFLAG Then frmMain.lblFPS.Caption = Mod_TileEngine.FPS
            lFrameTimer = GetTickCount
        End If
        If GetTickCount() >= lastFlush Then
            Call FlushBuffer
            lastFlush = GetTickCount() + 10
        End If
        DoEvents
    Loop
    Call CloseClient
End Sub

Public Function GetVersionOfTheGame() As String
    GetVersionOfTheGame = GetVar(Game.path(INIT) & "Config.ini", "Cliente", "VersionTagRelease")
End Function

Private Sub LoadInitialConfig()
    Dim AOLibreHelperFolder As String
    AOLibreHelperFolder = Left$(App.path, 2) & "\ao-project-config\"
    If Dir(AOLibreHelperFolder, vbDirectory) = "" Then
        MkDir AOLibreHelperFolder
    End If
    Set picMouseIcon = LoadPicture(Game.path(Graficos) & "MouseIcons\Baston.ico")
    Dim CursorAniDir As String
    Dim Cursor As Long
    CursorAniDir = Game.path(Graficos) & "MouseIcons\General.ani"
    hSwapCursor = SetClassLong(frmMain.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorAniDir))
    hSwapCursor = SetClassLong(frmMain.MainViewPic.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorAniDir))
    hSwapCursor = SetClassLong(frmMain.hlst.hwnd, GLC_HCURSOR, LoadCursorFromFile(CursorAniDir))
    frmCargando.Show
    frmCargando.Refresh
    frmConnect.version = GetVersionOfTheGame()
    Call AddtoRichTextBox(frmCargando.status, _
                            JsonLanguage.item("INICIA_CLASES").item("TEXTO"), _
                            JsonLanguage.item("INICIA_CLASES").item("COLOR").item(1), _
                            JsonLanguage.item("INICIA_CLASES").item("COLOR").item(2), _
                            JsonLanguage.item("INICIA_CLASES").item("COLOR").item(3), _
                            True, False, True, rtfCenter)
    Set Dialogos = New clsDialogs
    Set Audio = New clsAudio
    Set Inventario = New clsGraphicalInventory
    Set CustomKeys = New clsCustomKeys
    Set CustomMessages = New clsCustomMessages
    Set incomingData = New clsByteQueue
    Set outgoingData = New clsByteQueue
    Set MainTimer = New clsTimer
    Set clsForos = New clsForum
    Set frmMain.Client = New clsSocket
    Set keysMovementPressedQueue = New clsArrayList
    Call keysMovementPressedQueue.Initialize(1, 4)
    Call AddtoRichTextBox(frmCargando.status, _
                            "   " & JsonLanguage.item("HECHO").item("TEXTO"), _
                            JsonLanguage.item("HECHO").item("COLOR").item(1), _
                            JsonLanguage.item("HECHO").item("COLOR").item(2), _
                            JsonLanguage.item("HECHO").item("COLOR").item(3), _
                            True, False, False, rtfLeft)
    Call AddtoRichTextBox(frmCargando.status, _
                            JsonLanguage.item("INICIA_SONIDO").item("TEXTO"), _
                            JsonLanguage.item("INICIA_SONIDO").item("COLOR").item(1), _
                            JsonLanguage.item("INICIA_SONIDO").item("COLOR").item(2), _
                            JsonLanguage.item("INICIA_SONIDO").item("COLOR").item(3), _
                            True, False, True, rtfCenter)
    Call Audio.Initialize(DirectX, frmMain.hwnd, Game.path(Sounds), Game.path(Musica), Game.path(MusicaMp3))
    Audio.MusicActivated = ClientSetup.bMusic
    Audio.SoundActivated = ClientSetup.bSound
    Audio.SoundEffectsActivated = ClientSetup.bSoundEffects
    Audio.MusicVolume = ClientSetup.MusicVolume
    Audio.SoundVolume = ClientSetup.SoundVolume
    Call Audio.PlayBackgroundMusic("6", MusicTypes.Mp3)
    Call AddtoRichTextBox(frmCargando.status, _
                            "   " & JsonLanguage.item("HECHO").item("TEXTO"), _
                            JsonLanguage.item("HECHO").item("COLOR").item(1), _
                            JsonLanguage.item("HECHO").item("COLOR").item(2), _
                            JsonLanguage.item("HECHO").item("COLOR").item(3), _
                            True, False, False, rtfLeft)
    Call AddtoRichTextBox(frmCargando.status, _
                            JsonLanguage.item("INICIA_CONSTANTES").item("TEXTO"), _
                            JsonLanguage.item("INICIA_CONSTANTES").item("COLOR").item(1), _
                            JsonLanguage.item("INICIA_CONSTANTES").item("COLOR").item(2), _
                            JsonLanguage.item("INICIA_CONSTANTES").item("COLOR").item(3), _
                            True, False, True, rtfCenter)
    Call InicializarNombres
    Call Protocol.InitFonts
    UserMap = 1
    Call AddtoRichTextBox(frmCargando.status, _
                            "   " & JsonLanguage.item("HECHO").item("TEXTO"), _
                            JsonLanguage.item("HECHO").item("COLOR").item(1), _
                            JsonLanguage.item("HECHO").item("COLOR").item(2), _
                            JsonLanguage.item("HECHO").item("COLOR").item(3), _
                            True, False, False, rtfLeft)
    Call AddtoRichTextBox(frmCargando.status, _
                            JsonLanguage.item("INICIA_MOTOR_GRAFICO").item("TEXTO"), _
                            JsonLanguage.item("INICIA_MOTOR_GRAFICO").item("COLOR").item(1), _
                            JsonLanguage.item("INICIA_MOTOR_GRAFICO").item("COLOR").item(2), _
                            JsonLanguage.item("INICIA_MOTOR_GRAFICO").item("COLOR").item(3), _
                            True, False, True, rtfCenter)
    Call mDx8_Engine.Engine_DirectX8_Init
    Call Mod_TileEngine.InitTileEngine(frmMain.hwnd, 32, 32, 8, 8)
    Call AddtoRichTextBox(frmCargando.status, _
                            "   " & JsonLanguage.item("HECHO").item("TEXTO"), _
                            JsonLanguage.item("HECHO").item("COLOR").item(1), _
                            JsonLanguage.item("HECHO").item("COLOR").item(2), _
                            JsonLanguage.item("HECHO").item("COLOR").item(3), _
                            True, False, False, rtfLeft)
    Call AddtoRichTextBox(frmCargando.status, _
                            JsonLanguage.item("INICIA_FXS").item("TEXTO"), _
                            JsonLanguage.item("INICIA_FXS").item("COLOR").item(1), _
                            JsonLanguage.item("INICIA_FXS").item("COLOR").item(2), _
                            JsonLanguage.item("INICIA_FXS").item("COLOR").item(3), _
                            True, False, True, rtfCenter)
    Call CargarTips
    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarColores
    Call AddtoRichTextBox(frmCargando.status, _
                            "   " & JsonLanguage.item("HECHO").item("TEXTO"), _
                            JsonLanguage.item("HECHO").item("COLOR").item(1), _
                            JsonLanguage.item("HECHO").item("COLOR").item(2), _
                            JsonLanguage.item("HECHO").item("COLOR").item(3), _
                            True, False, False, rtfLeft)
    Call Inventario.Initialize(DirectD3D8, frmMain.PicInv, MAX_INVENTORY_SLOTS, , , , , , , , True)
    Call AddtoRichTextBox(frmCargando.status, _
                            JsonLanguage.item("BIENVENIDO").item("TEXTO"), _
                            JsonLanguage.item("BIENVENIDO").item("COLOR").item(1), _
                            JsonLanguage.item("BIENVENIDO").item("COLOR").item(2), _
                            JsonLanguage.item("BIENVENIDO").item("COLOR").item(3), _
                            True, False, True, rtfCenter)
    Unload frmCargando
End Sub

Private Sub LoadTimerIntervals()
    With MainTimer
        Call .SetInterval(TimersIndex.Attack, eIntervalos.INT_ATTACK)
        Call .SetInterval(TimersIndex.Work, eIntervalos.INT_WORK)
        Call .SetInterval(TimersIndex.UseItemWithU, eIntervalos.INT_USEITEMU)
        Call .SetInterval(TimersIndex.UseItemWithDblClick, eIntervalos.INT_USEITEMDCK)
        Call .SetInterval(TimersIndex.SendRPU, eIntervalos.INT_SENTRPU)
        Call .SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
        Call .SetInterval(TimersIndex.Arrows, eIntervalos.INT_ARROWS)
        Call .SetInterval(TimersIndex.CastAttack, eIntervalos.INT_CAST_ATTACK)
        Call .SetInterval(TimersIndex.ChangeHeading, eIntervalos.INT_CHANGE_HEADING)
        With frmMain.macrotrabajo
            .Interval = eIntervalos.INT_MACRO_TRABAJO
            .Enabled = False
        End With
        Call .Start(TimersIndex.Attack)
        Call .Start(TimersIndex.Work)
        Call .Start(TimersIndex.UseItemWithU)
        Call .Start(TimersIndex.UseItemWithDblClick)
        Call .Start(TimersIndex.SendRPU)
        Call .Start(TimersIndex.CastSpell)
        Call .Start(TimersIndex.Arrows)
        Call .Start(TimersIndex.CastAttack)
        Call .Start(TimersIndex.ChangeHeading)
    End With
End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
    writeprivateprofilestring Main, Var, Value, File
End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
    Dim sSpaces As String
    sSpaces = Space$(500)
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo ErrorHandler
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    Dim Len_sString As Long
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        Len_sString = Len(sString) - 1
        For lX = 0 To Len_sString
            If Not (lX = (lPos - 1)) Then
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        CheckMailString = True
    End If
ErrorHandler:
End Function

Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
    HayAgua = ((MapData(X, Y).Graphic(1).GrhIndex >= 1505 And MapData(X, Y).Graphic(1).GrhIndex <= 1520) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 5665 And MapData(X, Y).Graphic(1).GrhIndex <= 5680) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 13547 And MapData(X, Y).Graphic(1).GrhIndex <= 13562)) And _
                MapData(X, Y).Graphic(2).GrhIndex = 0
End Function

Public Sub LeerLineaComandos()
    Dim i As Long, t() As String, Upper_t As Long, Lower_t As Long
    t = Split(Command, " ")
    Lower_t = LBound(t)
    Upper_t = UBound(t)
    For i = Lower_t To Upper_t
        Select Case UCase$(t(i))
            Case "/NORES"
                NoRes = True
        End Select
    Next i
End Sub

Private Sub InicializarNombres()
    Ciudades(eCiudad.cUllathorpe) = "Ullathorpe"
    Ciudades(eCiudad.cNix) = "Nix"
    Ciudades(eCiudad.cBanderbill) = "Banderbill"
    Ciudades(eCiudad.cLindos) = "Lindos"
    Ciudades(eCiudad.cArghal) = "Arghal"
    ListaRazas(eRaza.Humano) = JsonLanguage.item("RAZAS").item("HUMANO")
    ListaRazas(eRaza.Elfo) = JsonLanguage.item("RAZAS").item("ELFO")
    ListaRazas(eRaza.ElfoOscuro) = JsonLanguage.item("RAZAS").item("ELFO_OSCURO")
    ListaRazas(eRaza.Gnomo) = JsonLanguage.item("RAZAS").item("GNOMO")
    ListaRazas(eRaza.Enano) = JsonLanguage.item("RAZAS").item("ENANO")
    ListaClases(eClass.Mage) = JsonLanguage.item("CLASES").item("MAGO")
    ListaClases(eClass.Cleric) = JsonLanguage.item("CLASES").item("CLERIGO")
    ListaClases(eClass.Warrior) = JsonLanguage.item("CLASES").item("GUERRERO")
    ListaClases(eClass.Assasin) = JsonLanguage.item("CLASES").item("ASESINO")
    ListaClases(eClass.Thief) = JsonLanguage.item("CLASES").item("LADRON")
    ListaClases(eClass.Bard) = JsonLanguage.item("CLASES").item("BARDO")
    ListaClases(eClass.Druid) = JsonLanguage.item("CLASES").item("DRUIDA")
    ListaClases(eClass.Bandit) = JsonLanguage.item("CLASES").item("BANDIDO")
    ListaClases(eClass.Paladin) = JsonLanguage.item("CLASES").item("PALADIN")
    ListaClases(eClass.Hunter) = JsonLanguage.item("CLASES").item("CAZADOR")
    ListaClases(eClass.Worker) = JsonLanguage.item("CLASES").item("TRABAJADOR")
    ListaClases(eClass.Pirate) = JsonLanguage.item("CLASES").item("PIRATA")
    SkillsNames(eSkill.Magia) = JsonLanguage.item("HABILIDADES").item("MAGIA").item("TEXTO")
    SkillsNames(eSkill.Robar) = JsonLanguage.item("HABILIDADES").item("ROBAR").item("TEXTO")
    SkillsNames(eSkill.Tacticas) = JsonLanguage.item("HABILIDADES").item("EVASION_EN_COMBATE").item("TEXTO")
    SkillsNames(eSkill.Armas) = JsonLanguage.item("HABILIDADES").item("COMBATE_CON_ARMAS").item("TEXTO")
    SkillsNames(eSkill.Meditar) = JsonLanguage.item("HABILIDADES").item("MEDITAR").item("TEXTO")
    SkillsNames(eSkill.Apunalar) = JsonLanguage.item("HABILIDADES").item("APUNALAR").item("TEXTO")
    SkillsNames(eSkill.Ocultarse) = JsonLanguage.item("HABILIDADES").item("OCULTARSE").item("TEXTO")
    SkillsNames(eSkill.Supervivencia) = JsonLanguage.item("HABILIDADES").item("SUPERVIVENCIA").item("TEXTO")
    SkillsNames(eSkill.Talar) = JsonLanguage.item("HABILIDADES").item("TALAR").item("TEXTO")
    SkillsNames(eSkill.Comerciar) = JsonLanguage.item("HABILIDADES").item("COMERCIO").item("TEXTO")
    SkillsNames(eSkill.Defensa) = JsonLanguage.item("HABILIDADES").item("DEFENSA_CON_ESCUDOS").item("TEXTO")
    SkillsNames(eSkill.Pesca) = JsonLanguage.item("HABILIDADES").item("PESCA").item("TEXTO")
    SkillsNames(eSkill.Mineria) = JsonLanguage.item("HABILIDADES").item("MINERIA").item("TEXTO")
    SkillsNames(eSkill.Carpinteria) = JsonLanguage.item("HABILIDADES").item("CARPINTERIA").item("TEXTO")
    SkillsNames(eSkill.Herreria) = JsonLanguage.item("HABILIDADES").item("HERRERIA").item("TEXTO")
    SkillsNames(eSkill.Liderazgo) = JsonLanguage.item("HABILIDADES").item("LIDERAZGO").item("TEXTO")
    SkillsNames(eSkill.Domar) = JsonLanguage.item("HABILIDADES").item("DOMAR_ANIMALES").item("TEXTO")
    SkillsNames(eSkill.Proyectiles) = JsonLanguage.item("HABILIDADES").item("COMBATE_A_DISTANCIA").item("TEXTO")
    SkillsNames(eSkill.Wrestling) = JsonLanguage.item("HABILIDADES").item("COMBATE_CUERPO_A_CUERPO").item("TEXTO")
    SkillsNames(eSkill.Navegacion) = JsonLanguage.item("HABILIDADES").item("NAVEGACION").item("TEXTO")
    AtributosNames(eAtributos.Fuerza) = JsonLanguage.item("ATRIBUTOS").item("FUERZA")
    AtributosNames(eAtributos.Agilidad) = JsonLanguage.item("ATRIBUTOS").item("AGILIDAD")
    AtributosNames(eAtributos.Inteligencia) = JsonLanguage.item("ATRIBUTOS").item("INTELIGENCIA")
    AtributosNames(eAtributos.Carisma) = JsonLanguage.item("ATRIBUTOS").item("CARISMA")
    AtributosNames(eAtributos.Constitucion) = JsonLanguage.item("ATRIBUTOS").item("CONSTITUCION")
End Sub

Public Sub CleanDialogs()
    frmMain.RecTxt.Text = vbNullString
    Call DialogosClanes.RemoveDialogs
    Call Dialogos.RemoveAllDialogs
End Sub

Public Sub CloseClient()
    Call Application.ReleaseInstance
    EngineRun = False
    If prgRun Then
        Call Game.GuardarConfiguracion
    End If
    frmMain.Client.CloseSck
    Call Engine_DirectX8_End
    Set CustomMessages = Nothing
    Set CustomKeys = Nothing
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    Set JsonLanguage = Nothing
    Set frmMain.Client = Nothing
    Call UnloadAllForms
    If ResolucionCambiada Then Resolution.ResetResolution
    End
End Sub

Public Function EsGM(ByVal CharIndex As Integer) As Boolean
    If charlist(CharIndex).priv >= 1 And charlist(CharIndex).priv <= 5 Or charlist(CharIndex).priv = 25 Then
        EsGM = True
        Exit Function
    End If
    EsGM = False
End Function

Public Function EsNPC(ByVal CharIndex As Integer) As Boolean
    If charlist(CharIndex).iHead = 0 Then
        EsNPC = True
        Exit Function
    End If
    EsNPC = False
End Function

Public Function getTagPosition(ByVal Nick As String) As Integer
    Dim buf As Integer
        buf = InStr(Nick, "<")
    If buf > 0 Then
        getTagPosition = buf
        Exit Function
    End If
    buf = InStr(Nick, "[")
    If buf > 0 Then
        getTagPosition = buf
        Exit Function
    End If
    getTagPosition = Len(Nick) + 2
End Function

Public Sub checkText(ByVal Text As String)
    Dim Nivel As Integer
    If Right$(Text, Len(JsonLanguage.item("MENSAJE_FRAGSHOOTER_TE_HA_MATADO").item("TEXTO"))) = JsonLanguage.item("MENSAJE_FRAGSHOOTER_TE_HA_MATADO").item("TEXTO") Then
        Call ScreenCapture(True)
        Exit Sub
    End If
    If Left$(Text, Len(JsonLanguage.item("MENSAJE_FRAGSHOOTER_HAS_MATADO").item("TEXTO"))) = JsonLanguage.item("MENSAJE_FRAGSHOOTER_HAS_MATADO").item("TEXTO") Then
        EsperandoLevel = True
        Exit Sub
    End If
    If EsperandoLevel Then
        If Right$(Text, Len(JsonLanguage.item("MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA").item("TEXTO"))) = JsonLanguage.item("MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA").item("TEXTO") Then
            If CInt(mid$(Text, Len(JsonLanguage.item("MENSAJE_FRAGSHOOTER_HAS_GANADO").item("TEXTO")), (Len(Text) - (Len(JsonLanguage.item("MENSAJE_FRAGSHOOTER_HAS_GANADO").item("TEXTO")))))) / 2 > ClientSetup.byMurderedLevel Then
                Call ScreenCapture(True)
            End If
        End If
    End If
    EsperandoLevel = False
End Sub

Public Function getStrenghtColor() As Long
    Dim m As Long
        m = 255 / MAXATRIBUTOS
    getStrenghtColor = RGB(255 - (m * UserFuerza), (m * UserFuerza), 0)
End Function
    
Public Function getDexterityColor() As Long
    Dim m As Long
    m = 255 / MAXATRIBUTOS
    getDexterityColor = RGB(255, m * UserAgilidad, 0)
End Function

Public Function getCharIndexByName(ByVal Name As String) As Integer
    Dim i As Long
    For i = 1 To LastChar
        If charlist(i).Nombre = Name Then
            getCharIndexByName = i
            Exit Function
        End If
    Next i
End Function

Public Function EsAnuncio(ByVal ForumType As Byte) As Boolean
    Select Case ForumType
        Case eForumMsgType.ieCAOS_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieGENERAL_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieREAL_STICKY
            EsAnuncio = True
    End Select
End Function

Public Function ForumAlignment(ByVal yForumType As Byte) As Byte
    Select Case yForumType
        Case eForumMsgType.ieCAOS, eForumMsgType.ieCAOS_STICKY
            ForumAlignment = eForumType.ieCAOS
            
        Case eForumMsgType.ieGeneral, eForumMsgType.ieGENERAL_STICKY
            ForumAlignment = eForumType.ieGeneral
            
        Case eForumMsgType.ieREAL, eForumMsgType.ieREAL_STICKY
            ForumAlignment = eForumType.ieREAL
    End Select
End Function

Public Sub ResetAllInfo(Optional ByVal UnloadForms As Boolean = True)
    frmMain.Second.Enabled = False
    frmMain.macrotrabajo.Enabled = False
    Connected = False
    Call frmMain.hlst.Clear
    If UnloadForms Then
        Dim frm As Form
        For Each frm In Forms
            If frm.Name <> frmMain.Name And _
               frm.Name <> frmConnect.Name And _
               frm.Name <> frmCrearPersonaje.Name Then
                Call Unload(frm)
            End If
        Next
    End If
    On Local Error GoTo 0
    If UnloadForms Then
        If Not frmCrearPersonaje.Visible Then frmConnect.Visible = True
        frmMain.Visible = False
    End If
    Call Audio.StopWave
    frmMain.IsPlaying = PlayLoop.plNone
    pausa = False
    UserMeditar = False
    UserEstupido = False
    UserCiego = False
    UserDescansar = False
    UserParalizado = False
    Traveling = False
    UserNavegando = False
    UserEvento = False
    bRain = False
    bFogata = False
    Comerciando = False
    bShowTutorial = False
    MirandoAsignarSkills = False
    MirandoCarpinteria = False
    MirandoEstadisticas = False
    MirandoForo = False
    MirandoHerreria = False
    MirandoParty = False
    Call CleanDialogs
    Dim i As Long
    For i = 1 To LastChar
        charlist(i).invisible = False
    Next i
    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = vbNullString
    SkillPoints = 0
    Alocados = 0
    UserEquitando = 0
    Call SetSpeedUsuario
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i
    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
    Inventario.ClearAllSlots
    Call Audio.PlayBackgroundMusic("2", MusicTypes.Mp3)
End Sub

Public Function DevolverNombreHechizo(ByVal Index As Byte) As String
Dim i As Long
    For i = 1 To NumHechizos
        If i = Index Then
            DevolverNombreHechizo = Hechizos(i).Nombre
            Exit Function
        End If
    Next i
End Function

Public Function DevolverIndexHechizo(ByVal Nombre As String) As Byte
Dim i As Long
    For i = 1 To NumHechizos
        If Hechizos(i).Nombre = Nombre Then
            DevolverIndexHechizo = i
            Exit Function
        End If
    Next i
End Function

Public Function ArrayInitialized(ByVal TheArray As Long) As Boolean
    ArrayInitialized = Not (TheArray = -1&)
End Function

Function ImgRequest(ByVal sFile As String) As String
    Dim RespondMsgBox As Byte
    If LenB(Dir(sFile, vbArchive)) = 0 Then
        RespondMsgBox = MsgBox("ERROR: Imagen no encontrada..." & vbCrLf & sFile, vbCritical + vbRetryCancel)
        If RespondMsgBox = vbRetry Then
            sFile = ImgRequest(sFile)
        Else
            Call MsgBox("ADVERTENCIA: El juego seguira funcionando sin alguna imagen!", vbInformation + vbOKOnly)
            sFile = Game.path(Interfaces) & "blank.bmp"
        End If
    End If
    ImgRequest = sFile
End Function

Public Sub LoadAOCustomControlsPictures(ByRef tForm As Form)
    Dim DirButtons As String
        DirButtons = Game.path(Graficos) & "\Botones\"
    Dim cControl As Control
    For Each cControl In tForm.Controls
        If TypeOf cControl Is uAOButton Then
            cControl.PictureEsquina = LoadPicture(ImgRequest(DirButtons & uAOButton_bEsquina))
            cControl.PictureFondo = LoadPicture(ImgRequest(DirButtons & uAOButton_bFondo))
            cControl.PictureHorizontal = LoadPicture(ImgRequest(DirButtons & uAOButton_bHorizontal))
            cControl.PictureVertical = LoadPicture(ImgRequest(DirButtons & uAOButton_bVertical))
        ElseIf TypeOf cControl Is uAOCheckbox Then
            cControl.Picture = LoadPicture(ImgRequest(DirButtons & uAOButton_cCheckboxSmall))
        End If
    Next
End Sub

Public Sub SetSpeedUsuario()
    If UserEquitando Then
        Engine_BaseSpeed = 0.024
    Else
        Engine_BaseSpeed = 0.018
    End If
End Sub

Public Function CheckIfIpIsNumeric(CurrentIp As String) As String
    If IsNumeric(mid$(CurrentIp, 1, 1)) Then
        CheckIfIpIsNumeric = True
    Else
        CheckIfIpIsNumeric = False
    End If
End Function

Public Function GetCountryCode(CurrentIp As String) As String
    Dim CountryCode As String
    CountryCode = GetCountryFromIp(CurrentIp)
    If LenB(CountryCode) > 0 Then
        GetCountryCode = CountryCode
    Else
        GetCountryCode = "??"
    End If
End Function
