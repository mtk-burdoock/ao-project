Attribute VB_Name = "Protocol"
Option Explicit

Private Const SEPARATOR As String * 1 = vbNullChar

Private Type tFont
    Red As Byte
    Green As Byte
    Blue As Byte
    bold As Boolean
    italic As Boolean
End Type

Private Enum ServerPacketID
    logged = 1
    RemoveDialogs = 2
    RemoveCharDialog = 3
    NavigateToggle = 4
    Disconnect = 5
    CommerceEnd = 6
    BankEnd = 7
    CommerceInit = 8
    BankInit = 9
    UserCommerceInit = 10
    UserCommerceEnd = 11
    UserOfferConfirm = 12
    CommerceChat = 13
    UpdateSta = 14
    UpdateMana = 15
    UpdateHP = 16
    UpdateGold = 17
    UpdateBankGold = 18
    UpdateExp = 19
    ChangeMap = 20
    PosUpdate = 21
    ChatOverHead = 22
    ConsoleMsg = 23
    GuildChat = 24
    ShowMessageBox = 25
    UserIndexInServer = 26
    UserCharIndexInServer = 27
    CharacterCreate = 28
    CharacterRemove = 29
    CharacterChangeNick = 30
    CharacterMove = 31
    ForceCharMove = 32
    CharacterChange = 33
    HeadingChange = 34
    ObjectCreate = 35
    ObjectDelete = 36
    BlockPosition = 37
    PlayMp3 = 38
    PlayMIDI = 39
    PlayWave = 40
    guildList = 41
    AreaChanged = 42
    PauseToggle = 43
    RainToggle = 44
    CreateFX = 45
    UpdateUserStats = 46
    ChangeInventorySlot = 47
    ChangeBankSlot = 48
    ChangeSpellSlot = 49
    Atributes = 50
    BlacksmithWeapons = 51
    BlacksmithArmors = 52
    InitCarpenting = 53
    RestOK = 54
    ErrorMsg = 55
    Blind = 56
    Dumb = 57
    ShowSignal = 58
    ChangeNPCInventorySlot = 59
    UpdateHungerAndThirst = 60
    Fame = 61
    MiniStats = 62
    LevelUp = 63
    AddForumMsg = 64
    ShowForumForm = 65
    SetInvisible = 66
    DiceRoll = 67
    MeditateToggle = 68
    BlindNoMore = 69
    DumbNoMore = 70
    SendSkills = 71
    TrainerCreatureList = 72
    guildNews = 73
    OfferDetails = 74
    AlianceProposalsList = 75
    PeaceProposalsList = 76
    CharacterInfo = 77
    GuildLeaderInfo = 78
    GuildMemberInfo = 79
    GuildDetails = 80
    ShowGuildFundationForm = 81
    ParalizeOK = 82
    ShowUserRequest = 83
    ChangeUserTradeSlot = 84
    SendNight = 85
    Pong = 86
    UpdateTagAndStatus = 87
    SpawnList = 88
    ShowSOSForm = 89
    ShowMOTDEditionForm = 90
    ShowGMPanelForm = 91
    UserNameList = 92
    ShowDenounces = 93
    RecordList = 94
    RecordDetails = 95
    ShowGuildAlign = 96
    ShowPartyForm = 97
    UpdateStrenghtAndDexterity = 98
    UpdateStrenght = 99
    UpdateDexterity = 100
    AddSlots = 101
    MultiMessage = 102
    StopWorking = 103
    CancelOfferItem = 104
    PalabrasMagicas = 105
    PlayAttackAnim = 106
    FXtoMap = 107
    AccountLogged = 108
    SearchList = 109
    QuestDetails = 110
    QuestListSend = 111
    CreateDamage = 112
    UserInEvent = 113
    renderMsg = 114
    DeletedChar = 115
    EquitandoToggle = 116
    EnviarDatosServer = 117
    InitCraftman = 118
    EnviarListDeAmigos = 119
    SeeInProcess = 120
    ShowProcess = 121
    Proyectil = 122
    PlayIsInChatMode = 123
End Enum

Private Enum ClientPacketID
    LoginExistingChar = 1
    ThrowDices = 2
    LoginNewChar = 3
    Talk = 4
    Yell = 5
    Whisper = 6
    Walk = 7
    RequestPositionUpdate = 8
    Attack = 9
    PickUp = 10
    SafeToggle = 11
    ResuscitationSafeToggle = 12
    RequestGuildLeaderInfo = 13
    RequestAtributes = 14
    RequestFame = 15
    RequestSkills = 16
    RequestMiniStats = 17
    CommerceEnd = 18
    UserCommerceEnd = 19
    UserCommerceConfirm = 20
    CommerceChat = 21
    BankEnd = 22
    UserCommerceOk = 23
    UserCommerceReject = 24
    Drop = 25
    CastSpell = 26
    LeftClick = 27
    DoubleClick = 28
    Work = 29
    UseSpellMacro = 30
    UseItem = 31
    CraftBlacksmith = 32
    CraftCarpenter = 33
    WorkLeftClick = 34
    CreateNewGuild = 35
    sadasdA = 36
    EquipItem = 37
    ChangeHeading = 38
    ModifySkills = 39
    Train = 40
    CommerceBuy = 41
    BankExtractItem = 42
    CommerceSell = 43
    BankDeposit = 44
    ForumPost = 45
    MoveSpell = 46
    MoveBank = 47
    ClanCodexUpdate = 48
    UserCommerceOffer = 49
    GuildAcceptPeace = 50
    GuildRejectAlliance = 51
    GuildRejectPeace = 52
    GuildAcceptAlliance = 53
    GuildOfferPeace = 54
    GuildOfferAlliance = 55
    GuildAllianceDetails = 56
    GuildPeaceDetails = 57
    GuildRequestJoinerInfo = 58
    GuildAlliancePropList = 59
    GuildPeacePropList = 60
    GuildDeclareWar = 61
    GuildNewWebsite = 62
    GuildAcceptNewMember = 63
    GuildRejectNewMember = 64
    GuildKickMember = 65
    GuildUpdateNews = 66
    GuildMemberInfo = 67
    GuildOpenElections = 68
    GuildRequestMembership = 69
    GuildRequestDetails = 70
    Online = 71
    Quit = 72
    GuildLeave = 73
    RequestAccountState = 74
    PetStand = 75
    PetFollow = 76
    ReleasePet = 77
    TrainList = 78
    Rest = 79
    Meditate = 80
    Resucitate = 81
    Heal = 82
    Help = 83
    RequestStats = 84
    CommerceStart = 85
    BankStart = 86
    Enlist = 87
    Information = 88
    Reward = 89
    RequestMOTD = 90
    UpTime = 91
    PartyLeave = 92
    PartyCreate = 93
    PartyJoin = 94
    Inquiry = 95
    GuildMessage = 96
    PartyMessage = 97
    GuildOnline = 98
    PartyOnline = 99
    CouncilMessage = 100
    RoleMasterRequest = 101
    GMRequest = 102
    bugReport = 103
    ChangeDescription = 104
    GuildVote = 105
    Punishments = 106
    ChangePassword = 107
    Gamble = 108
    InquiryVote = 109
    LeaveFaction = 110
    BankExtractGold = 111
    BankDepositGold = 112
    Denounce = 113
    GuildFundate = 114
    GuildFundation = 115
    PartyKick = 116
    PartySetLeader = 117
    PartyAcceptMember = 118
    Ping = 119
    RequestPartyForm = 120
    ItemUpgrade = 121
    GMCommands = 122
    InitCrafting = 123
    Home = 124
    ShowGuildNews = 125
    ShareNpc = 126
    StopSharingNpc = 127
    Consultation = 128
    moveItem = 129
    LoginExistingAccount = 130
    LoginNewAccount = 131
    CentinelReport = 132
    Ecvc = 133
    Acvc = 134
    IrCvc = 135
    DragAndDropHechizos = 136
    Quest = 137
    QuestAccept = 138
    QuestListRequest = 139
    QuestDetailsRequest = 140
    QuestAbandon = 141
    CambiarContrasena = 142
    FightSend = 143
    FightAccept = 144
    CloseGuild = 145
    Discord = 146
    DeleteChar = 147
    ObtenerDatosServer = 148
    CraftsmanCreate = 149
    AddAmigos = 150
    DelAmigos = 151
    OnAmigos = 152
    MsgAmigos = 153
    Lookprocess = 154
    SendProcessList = 155
    SendIfCharIsInChatMode = 156
End Enum

Public Enum FontTypeNames
    FONTTYPE_TALK = 0
    FONTTYPE_FIGHT = 1
    FONTTYPE_WARNING = 2
    FONTTYPE_INFO = 3
    FONTTYPE_INFOBOLD = 4
    FONTTYPE_EJECUCION = 5
    FONTTYPE_PARTY = 6
    FONTTYPE_VENENO = 7
    FONTTYPE_GUILD = 8
    FONTTYPE_SERVER = 9
    FONTTYPE_GUILDMSG = 10
    FONTTYPE_CONSEJO = 11
    FONTTYPE_CONSEJOCAOS = 12
    FONTTYPE_CONSEJOVesA = 13
    FONTTYPE_CONSEJOCAOSVesA = 14
    FONTTYPE_CENTINELA = 15
    FONTTYPE_GMMSG = 16
    FONTTYPE_GM = 17
    FONTTYPE_CITIZEN = 18
    FONTTYPE_CONSE = 19
    FONTTYPE_DIOS = 20
    FONTTYPE_CRIMINAL = 21
End Enum

Public FontTypes(21) As tFont

Public Sub Connect(ByVal Modo As E_MODO)
    frmConnect.btnConectarse.Enabled = False
    If frmMain.Client.State <> (sckClosed Or sckConnecting) Then
        frmMain.Client.CloseSck
        DoEvents
    End If
    EstadoLogin = Modo
    Call frmMain.Client.Connect(IPdelServidor, PuertoDelServidor)
    frmConnect.btnConectarse.Enabled = True
End Sub

Public Sub InitFonts()
    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        .Red = 204
        .Green = 255
        .Blue = 255
    End With
    With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
        .Red = 255
        .Green = 102
        .Blue = 102
        .bold = 1
        .italic = 0
    End With
    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
        .Red = 255
        .Green = 255
        .Blue = 102
        .bold = 1
        .italic = 0
    End With
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        .Red = 255
        .Green = 204
        .Blue = 153
    End With
    With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
        .Red = 255
        .Green = 204
        .Blue = 153
        .bold = 1
    End With
    With FontTypes(FontTypeNames.FONTTYPE_EJECUCION)
        .Red = 255
        .Green = 0
        .Blue = 127
        .bold = 1
    End With
    With FontTypes(FontTypeNames.FONTTYPE_PARTY)
        .Red = 252
        .Green = 203
        .Blue = 130
    End With
    With FontTypes(FontTypeNames.FONTTYPE_VENENO)
        .Red = 128
        .Green = 255
        .Blue = 0
        .bold = 1
    End With
    With FontTypes(FontTypeNames.FONTTYPE_GUILD)
        .Red = 205
        .Green = 101
        .Blue = 236
        .bold = 1
    End With
    With FontTypes(FontTypeNames.FONTTYPE_SERVER)
        .Red = 250
        .Green = 150
        .Blue = 237
        .bold = 1
    End With
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        .Red = 228
        .Green = 199
        .Blue = 27
    End With
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJO)
        .Red = 130
        .Green = 130
        .Blue = 255
        .bold = 1
    End With
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOS)
        .Red = 255
        .Green = 60
        .bold = 1
    End With
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA)
        .Green = 200
        .Blue = 255
        .bold = 1
    End With
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA)
        .Red = 255
        .Green = 50
        .bold = 1
    End With
    With FontTypes(FontTypeNames.FONTTYPE_CENTINELA)
        .Red = 240
        .Green = 230
        .Blue = 140
        .bold = 1
    End With
    With FontTypes(FontTypeNames.FONTTYPE_GMMSG)
        .Red = 255
        .Green = 255
        .Blue = 255
        .italic = 1
    End With
    With FontTypes(FontTypeNames.FONTTYPE_GM)
        .Red = 30
        .Green = 255
        .Blue = 30
        .bold = 1
    End With
    With FontTypes(FontTypeNames.FONTTYPE_CITIZEN)
        .Red = 78
        .Green = 78
        .Blue = 252
        .bold = 1
    End With
    With FontTypes(FontTypeNames.FONTTYPE_CONSE)
        .Red = 30
        .Green = 150
        .Blue = 30
        .bold = 1
    End With
    With FontTypes(FontTypeNames.FONTTYPE_DIOS)
        .Red = 250
        .Green = 250
        .Blue = 150
        .bold = 1
    End With
    With FontTypes(FontTypeNames.FONTTYPE_CRIMINAL)
        .Red = 224
        .Green = 52
        .Blue = 17
        .bold = 1
    End With
End Sub

Public Sub HandleIncomingData()
On Error Resume Next
    Dim Packet As Long: Packet = CLng(incomingData.PeekByte())
    Select Case Packet
        Case ServerPacketID.PlayIsInChatMode
            Call HandleSetTypingFlagToCharIndex
            
        Case ServerPacketID.logged
            Call HandleLogged
        
        Case ServerPacketID.RemoveDialogs
            Call HandleRemoveDialogs
        
        Case ServerPacketID.RemoveCharDialog
            Call HandleRemoveCharDialog
        
        Case ServerPacketID.NavigateToggle
            Call HandleNavigateToggle
        
        Case ServerPacketID.Disconnect
            Call HandleDisconnect
        
        Case ServerPacketID.CommerceEnd
            Call HandleCommerceEnd
            
        Case ServerPacketID.CommerceChat
            Call HandleCommerceChat
        
        Case ServerPacketID.BankEnd
            Call HandleBankEnd
        
        Case ServerPacketID.CommerceInit
            Call HandleCommerceInit
        
        Case ServerPacketID.BankInit
            Call HandleBankInit
        
        Case ServerPacketID.UserCommerceInit
            Call HandleUserCommerceInit
        
        Case ServerPacketID.UserCommerceEnd
            Call HandleUserCommerceEnd
            
        Case ServerPacketID.UserOfferConfirm
            Call HandleUserOfferConfirm

        Case ServerPacketID.UpdateSta
            Call HandleUpdateSta
        
        Case ServerPacketID.UpdateMana
            Call HandleUpdateMana
        
        Case ServerPacketID.UpdateHP
            Call HandleUpdateHP
        
        Case ServerPacketID.UpdateGold
            Call HandleUpdateGold
            
        Case ServerPacketID.UpdateBankGold
            Call HandleUpdateBankGold

        Case ServerPacketID.UpdateExp
            Call HandleUpdateExp
        
        Case ServerPacketID.ChangeMap
            Call HandleChangeMap
        
        Case ServerPacketID.PosUpdate
            Call HandlePosUpdate
        
        Case ServerPacketID.ChatOverHead
            Call HandleChatOverHead
        
        Case ServerPacketID.ConsoleMsg
            Call HandleConsoleMessage
        
        Case ServerPacketID.GuildChat
            Call HandleGuildChat
        
        Case ServerPacketID.ShowMessageBox
            Call HandleShowMessageBox
        
        Case ServerPacketID.UserIndexInServer
            Call HandleUserIndexInServer
        
        Case ServerPacketID.UserCharIndexInServer
            Call HandleUserCharIndexInServer
        
        Case ServerPacketID.CharacterCreate
            Call HandleCharacterCreate
        
        Case ServerPacketID.CharacterRemove
            Call HandleCharacterRemove
        
        Case ServerPacketID.CharacterChangeNick
            Call HandleCharacterChangeNick
            
        Case ServerPacketID.CharacterMove
            Call HandleCharacterMove
            
        Case ServerPacketID.ForceCharMove
            Call HandleForceCharMove
        
        Case ServerPacketID.CharacterChange
            Call HandleCharacterChange
            
        Case ServerPacketID.HeadingChange
            Call HandleHeadingChange
            
        Case ServerPacketID.ObjectCreate
            Call HandleObjectCreate
        
        Case ServerPacketID.ObjectDelete
            Call HandleObjectDelete
        
        Case ServerPacketID.BlockPosition
            Call HandleBlockPosition

        Case ServerPacketID.PlayMp3
            Call HandlePlayMP3
        
        Case ServerPacketID.PlayMIDI
            Call HandlePlayMIDI
        
        Case ServerPacketID.PlayWave
            Call HandlePlayWave
        
        Case ServerPacketID.guildList
            Call HandleGuildList
        
        Case ServerPacketID.AreaChanged
            Call HandleAreaChanged
        
        Case ServerPacketID.PauseToggle
            Call HandlePauseToggle
        
        Case ServerPacketID.RainToggle
            Call HandleRainToggle
        
        Case ServerPacketID.CreateFX
            Call HandleCreateFX
        
        Case ServerPacketID.UpdateUserStats
            Call HandleUpdateUserStats

        Case ServerPacketID.ChangeInventorySlot
            Call HandleChangeInventorySlot
        
        Case ServerPacketID.ChangeBankSlot
            Call HandleChangeBankSlot
        
        Case ServerPacketID.ChangeSpellSlot
            Call HandleChangeSpellSlot
        
        Case ServerPacketID.Atributes
            Call HandleAtributes
        
        Case ServerPacketID.BlacksmithWeapons
            Call HandleBlacksmithWeapons
        
        Case ServerPacketID.BlacksmithArmors
            Call HandleBlacksmithArmors
        
        Case ServerPacketID.InitCarpenting
            Call HandleInitCarpenting
            
        Case ServerPacketID.InitCraftman
            Call HandleInitCraftman
        
        Case ServerPacketID.RestOK
            Call HandleRestOK
        
        Case ServerPacketID.ErrorMsg
            Call HandleErrorMessage
        
        Case ServerPacketID.Blind
            Call HandleBlind
        
        Case ServerPacketID.Dumb
            Call HandleDumb
        
        Case ServerPacketID.ShowSignal
            Call HandleShowSignal
        
        Case ServerPacketID.ChangeNPCInventorySlot
            Call HandleChangeNPCInventorySlot
        
        Case ServerPacketID.UpdateHungerAndThirst
            Call HandleUpdateHungerAndThirst
        
        Case ServerPacketID.Fame
            Call HandleFame
        
        Case ServerPacketID.MiniStats
            Call HandleMiniStats
        
        Case ServerPacketID.LevelUp
            Call HandleLevelUp
        
        Case ServerPacketID.AddForumMsg
            Call HandleAddForumMessage
        
        Case ServerPacketID.ShowForumForm
            Call HandleShowForumForm
        
        Case ServerPacketID.SetInvisible
            Call HandleSetInvisible
        
        Case ServerPacketID.DiceRoll
            Call HandleDiceRoll
        
        Case ServerPacketID.MeditateToggle
            Call HandleMeditateToggle
        
        Case ServerPacketID.BlindNoMore
            Call HandleBlindNoMore
        
        Case ServerPacketID.DumbNoMore
            Call HandleDumbNoMore
        
        Case ServerPacketID.SendSkills
            Call HandleSendSkills
        
        Case ServerPacketID.TrainerCreatureList
            Call HandleTrainerCreatureList
        
        Case ServerPacketID.guildNews
            Call HandleGuildNews
        
        Case ServerPacketID.OfferDetails
            Call HandleOfferDetails
        
        Case ServerPacketID.AlianceProposalsList
            Call HandleAlianceProposalsList
        
        Case ServerPacketID.PeaceProposalsList
            Call HandlePeaceProposalsList
        
        Case ServerPacketID.CharacterInfo
            Call HandleCharacterInfo
        
        Case ServerPacketID.GuildLeaderInfo
            Call HandleGuildLeaderInfo
        
        Case ServerPacketID.GuildDetails
            Call HandleGuildDetails
        
        Case ServerPacketID.ShowGuildFundationForm
            Call HandleShowGuildFundationForm
        
        Case ServerPacketID.ParalizeOK
            Call HandleParalizeOK
        
        Case ServerPacketID.ShowUserRequest
            Call HandleShowUserRequest

        Case ServerPacketID.ChangeUserTradeSlot
            Call HandleChangeUserTradeSlot
            
        Case ServerPacketID.SendNight
            Call HandleSendNight
        
        Case ServerPacketID.Pong
            Call HandlePong
        
        Case ServerPacketID.UpdateTagAndStatus
            Call HandleUpdateTagAndStatus
        
        Case ServerPacketID.GuildMemberInfo
            Call HandleGuildMemberInfo
            
        Case ServerPacketID.PalabrasMagicas
            Call HandlePalabrasMagicas
            
        Case ServerPacketID.PlayAttackAnim
            Call HandleAttackAnim
            
        Case ServerPacketID.FXtoMap
            Call HandleFXtoMap
        
        Case ServerPacketID.AccountLogged
            Call HandleAccountLogged
            
        Case ServerPacketID.SearchList
            Call HandleSearchList

        Case ServerPacketID.QuestDetails
            Call HandleQuestDetails

        Case ServerPacketID.QuestListSend
            Call HandleQuestListSend

        Case ServerPacketID.SpawnList
            Call HandleSpawnList
        
        Case ServerPacketID.ShowSOSForm
            Call HandleShowSOSForm
            
        Case ServerPacketID.ShowDenounces
            Call HandleShowDenounces
            
        Case ServerPacketID.RecordDetails
            Call HandleRecordDetails
            
        Case ServerPacketID.RecordList
            Call HandleRecordList
            
        Case ServerPacketID.ShowMOTDEditionForm
            Call HandleShowMOTDEditionForm
        
        Case ServerPacketID.ShowGMPanelForm
            Call HandleShowGMPanelForm
        
        Case ServerPacketID.UserNameList
            Call HandleUserNameList
            
        Case ServerPacketID.ShowGuildAlign
            Call HandleShowGuildAlign
        
        Case ServerPacketID.ShowPartyForm
            Call HandleShowPartyForm
        
        Case ServerPacketID.UpdateStrenghtAndDexterity
            Call HandleUpdateStrenghtAndDexterity
            
        Case ServerPacketID.UpdateStrenght
            Call HandleUpdateStrenght
            
        Case ServerPacketID.UpdateDexterity
            Call HandleUpdateDexterity
            
        Case ServerPacketID.AddSlots
            Call HandleAddSlots

        Case ServerPacketID.MultiMessage
            Call HandleMultiMessage
        
        Case ServerPacketID.StopWorking
            Call HandleStopWorking
            
        Case ServerPacketID.CancelOfferItem
            Call HandleCancelOfferItem
            
        Case ServerPacketID.CreateDamage
            Call HandleCreateDamage
    
        Case ServerPacketID.UserInEvent
            Call HandleUserInEvent
            
        Case ServerPacketID.renderMsg
            Call HandleRenderMsg

        Case ServerPacketID.DeletedChar
            Call HandleDeletedChar

        Case ServerPacketID.EquitandoToggle
            Call HandleEquitandoToggle

        Case ServerPacketID.EnviarDatosServer
            Call HandleEnviarDatosServer
            
        Case ServerPacketID.EnviarListDeAmigos
            Call HandleEnviarListDeAmigos

        Case ServerPacketID.SeeInProcess
            Call HandleSeeInProcess
            
        Case ServerPacketID.ShowProcess
            Call HandleShowProcess
            
        Case ServerPacketID.Proyectil
            Call HandleProyectil

        Case Else
            Exit Sub
    End Select
    If incomingData.Length > 0 And Err.number <> incomingData.NotEnoughDataErrCode Then
        Call Err.Clear
        Call HandleIncomingData
    End If
End Sub

Public Sub HandleMultiMessage()
    Dim BodyPart As Byte
    Dim Dano As Integer
    Dim SpellIndex As Integer
    Dim Nombre     As String
    With incomingData
        Call .ReadByte
        Select Case .ReadByte
            Case eMessages.NPCSwing
                Call AddtoRichTextBox(frmMain.RecTxt, _
                        JsonLanguage.item("MENSAJE_CRIATURA_FALLA_GOLPE").item("TEXTO"), _
                        JsonLanguage.item("MENSAJE_CRIATURA_FALLA_GOLPE").item("COLOR").item(1), _
                        JsonLanguage.item("MENSAJE_CRIATURA_FALLA_GOLPE").item("COLOR").item(2), _
                        JsonLanguage.item("MENSAJE_CRIATURA_FALLA_GOLPE").item("COLOR").item(3), _
                        True, False, True)
        
            Case eMessages.NPCKillUser
                Call AddtoRichTextBox(frmMain.RecTxt, _
                    JsonLanguage.item("MENSAJE_CRIATURA_MATADO").item("TEXTO"), _
                    JsonLanguage.item("MENSAJE_CRIATURA_MATADO").item("COLOR").item(1), _
                    JsonLanguage.item("MENSAJE_CRIATURA_MATADO").item("COLOR").item(2), _
                    JsonLanguage.item("MENSAJE_CRIATURA_MATADO").item("COLOR").item(3), _
                    True, False, True)
        
            Case eMessages.BlockedWithShieldUser
                Call AddtoRichTextBox(frmMain.RecTxt, _
                    JsonLanguage.item("MENSAJE_RECHAZO_ATAQUE_ESCUDO").item("TEXTO"), _
                    JsonLanguage.item("MENSAJE_RECHAZO_ATAQUE_ESCUDO").item("COLOR").item(1), _
                    JsonLanguage.item("MENSAJE_RECHAZO_ATAQUE_ESCUDO").item("COLOR").item(2), _
                    JsonLanguage.item("MENSAJE_RECHAZO_ATAQUE_ESCUDO").item("COLOR").item(3), _
                    True, False, True)
        
            Case eMessages.BlockedWithShieldOther
                Call AddtoRichTextBox(frmMain.RecTxt, _
                    JsonLanguage.item("MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO").item("TEXTO"), _
                    JsonLanguage.item("MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO").item("COLOR").item(1), _
                    JsonLanguage.item("MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO").item("COLOR").item(2), _
                    JsonLanguage.item("MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO").item("COLOR").item(3), _
                    True, False, True)
        
            Case eMessages.UserSwing
                Call AddtoRichTextBox(frmMain.RecTxt, _
                    JsonLanguage.item("MENSAJE_FALLADO_GOLPE").item("TEXTO"), _
                    JsonLanguage.item("MENSAJE_FALLADO_GOLPE").item("COLOR").item(1), _
                    JsonLanguage.item("MENSAJE_FALLADO_GOLPE").item("COLOR").item(2), _
                    JsonLanguage.item("MENSAJE_FALLADO_GOLPE").item("COLOR").item(3), _
                    True, False, True)
        
            Case eMessages.SafeModeOn
                Call frmMain.ControlSM(eSMType.sSafemode, True)
        
            Case eMessages.SafeModeOff
                Call frmMain.ControlSM(eSMType.sSafemode, False)
        
            Case eMessages.ResuscitationSafeOff
                Call frmMain.ControlSM(eSMType.sResucitation, False)
         
            Case eMessages.ResuscitationSafeOn
                Call frmMain.ControlSM(eSMType.sResucitation, True)
        
            Case eMessages.NobilityLost
                Call AddtoRichTextBox(frmMain.RecTxt, _
                                        JsonLanguage.item("MENSAJE_PIERDE_NOBLEZA").item("TEXTO"), _
                                        JsonLanguage.item("MENSAJE_PIERDE_NOBLEZA").item("COLOR").item(1), _
                                        JsonLanguage.item("MENSAJE_PIERDE_NOBLEZA").item("COLOR").item(2), _
                                        JsonLanguage.item("MENSAJE_PIERDE_NOBLEZA").item("COLOR").item(3), _
                                        False, False, True)
        
            Case eMessages.CantUseWhileMeditating
                Call AddtoRichTextBox(frmMain.RecTxt, _
                                        JsonLanguage.item("MENSAJE_USAR_MEDITANDO").item("TEXTO"), _
                                        JsonLanguage.item("MENSAJE_USAR_MEDITANDO").item("COLOR").item(1), _
                                        JsonLanguage.item("MENSAJE_USAR_MEDITANDO").item("COLOR").item(2), _
                                        JsonLanguage.item("MENSAJE_USAR_MEDITANDO").item("COLOR").item(3), _
                                        False, False, True)
        
            Case eMessages.NPCHitUser
                Select Case incomingData.ReadByte()
                    Case ePartesCuerpo.bCabeza
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_GOLPE_CABEZA").item("TEXTO") & CStr(incomingData.ReadInteger()) & "!!", _
                            JsonLanguage.item("MENSAJE_GOLPE_CABEZA").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_GOLPE_CABEZA").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_GOLPE_CABEZA").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bBrazoIzquierdo
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_IZQ").item("TEXTO") & CStr(incomingData.ReadInteger()) & "!!", _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_IZQ").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_IZQ").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_IZQ").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bBrazoDerecho
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_DER").item("TEXTO") & CStr(incomingData.ReadInteger()) & "!!", _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_DER").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_DER").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_DER").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bPiernaIzquierda
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_IZQ").item("TEXTO") & CStr(incomingData.ReadInteger()) & "!!", _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_IZQ").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_IZQ").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_IZQ").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bPiernaDerecha
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_DER").item("TEXTO") & CStr(incomingData.ReadInteger()) & "!!", _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_DER").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_DER").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_DER").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bTorso
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_GOLPE_TORSO").item("TEXTO") & CStr(incomingData.ReadInteger() & "!!"), _
                            JsonLanguage.item("MENSAJE_GOLPE_TORSO").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_GOLPE_TORSO").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_GOLPE_TORSO").item("COLOR").item(3), _
                            True, False, True)
                End Select
        
            Case eMessages.UserHitNPC
                Dim MsgHitNpc As String
                    MsgHitNpc = JsonLanguage.item("MENSAJE_DAMAGE_NPC").item("TEXTO")
                    MsgHitNpc = Replace$(MsgHitNpc, "VAR_DANO", CStr(incomingData.ReadLong()))
                Call AddtoRichTextBox(frmMain.RecTxt, _
                                        MsgHitNpc, _
                                        JsonLanguage.item("MENSAJE_DAMAGE_NPC").item("COLOR").item(1), _
                                        JsonLanguage.item("MENSAJE_DAMAGE_NPC").item("COLOR").item(2), _
                                        JsonLanguage.item("MENSAJE_DAMAGE_NPC").item("COLOR").item(3), _
                                        True, False, True)
        
            Case eMessages.UserAttackedSwing
                Call AddtoRichTextBox(frmMain.RecTxt, _
                                        charlist(incomingData.ReadInteger()).Nombre & JsonLanguage.item("MENSAJE_ATAQUE_FALLO").item("TEXTO"), _
                                        JsonLanguage.item("MENSAJE_ATAQUE_FALLO").item("COLOR").item(1), _
                                        JsonLanguage.item("MENSAJE_ATAQUE_FALLO").item("COLOR").item(2), _
                                        JsonLanguage.item("MENSAJE_ATAQUE_FALLO").item("COLOR").item(3), _
                                        True, False, True)
        
            Case eMessages.UserHittedByUser
                Dim AttackerName As String
                AttackerName = GetRawName(charlist(incomingData.ReadInteger()).Nombre)
                BodyPart = incomingData.ReadByte()
                Dano = incomingData.ReadInteger()
            
                Select Case BodyPart
                    Case ePartesCuerpo.bCabeza
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            AttackerName & JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_CABEZA").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_CABEZA").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_CABEZA").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_CABEZA").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bBrazoIzquierdo
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                        AttackerName & JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ").item("TEXTO") & Dano & MENSAJE_2, _
                        JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ").item("COLOR").item(1), _
                        JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ").item("COLOR").item(2), _
                        JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ").item("COLOR").item(3), _
                        True, False, True)
                
                    Case ePartesCuerpo.bBrazoDerecho
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            AttackerName & JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_DER").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_DER").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_DER").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_DER").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bPiernaIzquierda
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            AttackerName & JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bPiernaDerecha
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            AttackerName & JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_DER").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_DER").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_DER").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_DER").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bTorso
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            AttackerName & JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_TORSO").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_TORSO").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_TORSO").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_TORSO").item("COLOR").item(3), _
                            True, False, True)
                End Select
        
            Case eMessages.UserHittedUser
                Dim VictimName As String
                VictimName = GetRawName(charlist(incomingData.ReadInteger()).Nombre)
                BodyPart = incomingData.ReadByte()
                Dano = incomingData.ReadInteger()
                Select Case BodyPart
                    Case ePartesCuerpo.bCabeza
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_1").item("TEXTO") & VictimName & JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_CABEZA").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_CABEZA").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_CABEZA").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_CABEZA").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bBrazoIzquierdo
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_1").item("TEXTO") & VictimName & JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bBrazoDerecho
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_1").item("TEXTO") & VictimName & JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_DER").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_DER").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_DER").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_DER").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bPiernaIzquierda
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_1").item("TEXTO") & VictimName & JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bPiernaDerecha
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_1").item("TEXTO") & VictimName & JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_DER").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_DER").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_DER").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_DER").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bTorso
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_1").item("TEXTO") & VictimName & JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_TORSO").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_TORSO").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_TORSO").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_TORSO").item("COLOR").item(3), _
                            True, False, True)
                End Select
                
            Case eMessages.WorkRequestTarget
                UsingSkill = incomingData.ReadByte()
                frmMain.MousePointer = 2
                Select Case UsingSkill
                    Case Magia
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_TRABAJO_MAGIA").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_TRABAJO_MAGIA").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_TRABAJO_MAGIA").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_TRABAJO_MAGIA").item("COLOR").item(3))
                
                    Case Pesca
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_TRABAJO_PESCA").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_TRABAJO_PESCA").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_TRABAJO_PESCA").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_TRABAJO_PESCA").item("COLOR").item(3))
                
                    Case Robar
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_TRABAJO_ROBAR").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_TRABAJO_ROBAR").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_TRABAJO_ROBAR").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_TRABAJO_ROBAR").item("COLOR").item(3))
                
                    Case Talar
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_TRABAJO_TALAR").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_TRABAJO_TALAR").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_TRABAJO_TALAR").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_TRABAJO_TALAR").item("COLOR").item(3))
                
                    Case Mineria
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_TRABAJO_MINERIA").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_TRABAJO_MINERIA").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_TRABAJO_MINERIA").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_TRABAJO_MINERIA").item("COLOR").item(3))
                
                    Case FundirMetal
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_TRABAJO_FUNDIRMETAL").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_TRABAJO_FUNDIRMETAL").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_TRABAJO_FUNDIRMETAL").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_TRABAJO_FUNDIRMETAL").item("COLOR").item(3))
                
                    Case Proyectiles
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_TRABAJO_PROYECTILES").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_TRABAJO_PROYECTILES").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_TRABAJO_PROYECTILES").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_TRABAJO_PROYECTILES").item("COLOR").item(3))
                End Select

            Case eMessages.HaveKilledUser
                Dim KilledUser As Integer
                Dim Exp        As Long
                Dim MensajeExp As String
                KilledUser = .ReadInteger
                Exp = .ReadLong
                Call AddtoRichTextBox(frmMain.RecTxt, _
                    JsonLanguage.item("MENSAJE_HAS_MATADO_A").item("TEXTO") & charlist(KilledUser).Nombre & MENSAJE_22, _
                    JsonLanguage.item("MENSAJE_HAS_MATADO_A").item("COLOR").item(1), _
                    JsonLanguage.item("MENSAJE_HAS_MATADO_A").item("COLOR").item(2), _
                    JsonLanguage.item("MENSAJE_HAS_MATADO_A").item("COLOR").item(3), _
                    True, False)
                MensajeExp = JsonLanguage.item("MENSAJE_HAS_GANADO_EXP").item("TEXTO")
                MensajeExp = Replace$(MensajeExp, "VAR_EXP_GANADA", Exp)
                Call AddtoRichTextBox(frmMain.RecTxt, _
                                    MensajeExp, _
                                    JsonLanguage.item("MENSAJE_HAS_GANADO_EXP").item("COLOR").item(1), _
                                    JsonLanguage.item("MENSAJE_HAS_GANADO_EXP").item("COLOR").item(2), _
                                    JsonLanguage.item("MENSAJE_HAS_GANADO_EXP").item("COLOR").item(3), _
                                    True, False)
                If ClientSetup.bKill And ClientSetup.bActive Then
                    If Exp \ 2 > ClientSetup.byMurderedLevel Then
                        FragShooterNickname = charlist(KilledUser).Nombre
                        FragShooterKilledSomeone = True
                        FragShooterCapturePending = True
                    End If
                End If
            
            Case eMessages.UserKill
                Dim KillerUser As Integer
                    KillerUser = .ReadInteger
                Call AddtoRichTextBox(frmMain.RecTxt, _
                                    charlist(KillerUser).Nombre & JsonLanguage.item("MENSAJE_TE_HA_MATADO").item("TEXTO"), _
                                    JsonLanguage.item("MENSAJE_TE_HA_MATADO").item("COLOR").item(1), _
                                    JsonLanguage.item("MENSAJE_TE_HA_MATADO").item("COLOR").item(2), _
                                    JsonLanguage.item("MENSAJE_TE_HA_MATADO").item("COLOR").item(3), _
                                    True, False)
                If ClientSetup.bDie And ClientSetup.bActive Then
                    FragShooterNickname = charlist(KillerUser).Nombre
                    FragShooterKilledSomeone = False
                    FragShooterCapturePending = True
                End If
            
            Case eMessages.NPCKill
                Call AddtoRichTextBox(frmMain.RecTxt, _
                                        JsonLanguage.item("NPC_KILL").item("TEXTO"), _
                                        JsonLanguage.item("NPC_KILL").item("COLOR").item(1), _
                                        JsonLanguage.item("NPC_KILL").item("COLOR").item(2), _
                                        JsonLanguage.item("NPC_KILL").item("COLOR").item(3), _
                                        True, False)
            
            Case eMessages.EarnExp
                Dim ExpObtenida As Long: ExpObtenida = .ReadLong()
                Dim MENSAJE_HAS_GANADO_EXP As String
                    MENSAJE_HAS_GANADO_EXP = JsonLanguage.item("MENSAJE_HAS_GANADO_EXP").item("TEXTO")
                    MENSAJE_HAS_GANADO_EXP = Replace$(MENSAJE_HAS_GANADO_EXP, "VAR_EXP_GANADA", ExpObtenida)
                Call AddtoRichTextBox(frmMain.RecTxt, _
                                        MENSAJE_HAS_GANADO_EXP, _
                                        JsonLanguage.item("MENSAJE_HAS_GANADO_EXP").item("COLOR").item(1), _
                                        JsonLanguage.item("MENSAJE_HAS_GANADO_EXP").item("COLOR").item(2), _
                                        JsonLanguage.item("MENSAJE_HAS_GANADO_EXP").item("COLOR").item(3), _
                                        True, False)
        
            Case eMessages.GoHome
                Dim Distance As Byte
                Dim Hogar    As String
                Dim tiempo   As Integer
                Dim msg      As String
                Dim msgGoHome As String
                Distance = .ReadByte
                tiempo = .ReadInteger
                Hogar = .ReadASCIIString
                If tiempo >= 60 Then
                    If tiempo Mod 60 = 0 Then
                        msg = tiempo / 60 & " " & JsonLanguage.item("MINUTOS").item("TEXTO") & "."
                    Else
                        msg = CInt(tiempo \ 60) & " " & JsonLanguage.item("MINUTOS").item("TEXTO") & " " & JsonLanguage.item("LETRA_Y").item("TEXTO") & " " & tiempo Mod 60 & " " & JsonLanguage.item("SEGUNDOS").item("TEXTO") & "."
                    End If
                Else
                    msg = tiempo & " " & JsonLanguage.item("SEGUNDOS").item("TEXTO") & "."
                End If
                msgGoHome = JsonLanguage.item("MENSAJE_ESTAS_A_MAPAS_DE_DURACION_VIAJE").item("TEXTO") & msg
                msgGoHome = Replace$(msgGoHome, "VAR_DISTANCIA_MAPAS", Distance)
                msgGoHome = Replace$(msgGoHome, "VAR_MAPA_DESTINO", Hogar)
                Call ShowConsoleMsg(msgGoHome, _
                                    JsonLanguage.item("MENSAJE_ESTAS_A_MAPAS_DE_DURACION_VIAJE").item("COLOR").item(1), _
                                    JsonLanguage.item("MENSAJE_ESTAS_A_MAPAS_DE_DURACION_VIAJE").item("COLOR").item(2), _
                                    JsonLanguage.item("MENSAJE_ESTAS_A_MAPAS_DE_DURACION_VIAJE").item("COLOR").item(3), _
                                    True)
                Traveling = True

            Case eMessages.CancelGoHome
                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_HOGAR_CANCEL").item("TEXTO"), _
                                    JsonLanguage.item("MENSAJE_HOGAR_CANCEL").item("COLOR").item(1), _
                                    JsonLanguage.item("MENSAJE_HOGAR_CANCEL").item("COLOR").item(2), _
                                    JsonLanguage.item("MENSAJE_HOGAR_CANCEL").item("COLOR").item(3), _
                                    True)
                Traveling = False
                   
            Case eMessages.FinishHome
                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_HOGAR").item("TEXTO"), _
                                    JsonLanguage.item("MENSAJE_HOGAR").item("COLOR").item(1), _
                                    JsonLanguage.item("MENSAJE_HOGAR").item("COLOR").item(2), _
                                    JsonLanguage.item("MENSAJE_HOGAR").item("COLOR").item(3))
                Traveling = False
            
            Case eMessages.UserMuerto
                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                    JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                    JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                    JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
        
            Case eMessages.NpcInmune
                Call AddtoRichTextBox(frmMain.RecTxt, _
                                    JsonLanguage.item("NPC_INMUNE").item("TEXTO"), _
                                    JsonLanguage.item("NPC_INMUNE").item("COLOR").item(1), _
                                    JsonLanguage.item("NPC_INMUNE").item("COLOR").item(2), _
                                    JsonLanguage.item("NPC_INMUNE").item("COLOR").item(3))
            
            Case eMessages.Hechizo_HechiceroMSG_NOMBRE
                SpellIndex = .ReadByte
                Nombre = .ReadASCIIString
                Call ShowConsoleMsg(Hechizos(SpellIndex).HechiceroMsg & " " & Nombre & ".", 210, 220, 220)
         
            Case eMessages.Hechizo_HechiceroMSG_ALGUIEN
                SpellIndex = .ReadByte
                Call ShowConsoleMsg(Hechizos(SpellIndex).HechiceroMsg & " " & JsonLanguage.item("ALGUIEN").item("TEXTO") & ".", 210, 220, 220)
         
            Case eMessages.Hechizo_HechiceroMSG_CRIATURA
                SpellIndex = .ReadByte
                Call ShowConsoleMsg(Hechizos(SpellIndex).HechiceroMsg & " la criatura.", 210, 220, 220)
         
            Case eMessages.Hechizo_PropioMSG
                SpellIndex = .ReadByte
                Call ShowConsoleMsg(Hechizos(SpellIndex).PropioMsg, 210, 220, 220)
         
            Case eMessages.Hechizo_TargetMSG
                SpellIndex = .ReadByte
                Nombre = .ReadASCIIString
                Call ShowConsoleMsg(Nombre & " " & Hechizos(SpellIndex).TargetMsg, 210, 220, 220)
        End Select
    End With
End Sub

Private Sub HandleDeletedChar()
    Call incomingData.ReadByte
    MsgBox ("El personaje se ha borrado correctamente. Por favor vuelve a iniciar sesion para ver el cambio")
    Call CloseConnectionAndResetAllInfo
End Sub

Private Sub HandleLogged()
    Call incomingData.ReadByte
    #If AntiExternos Then
        Security.Redundance = incomingData.ReadByte()
    #End If
    UserClase = incomingData.ReadByte
    IntervaloInvi = incomingData.ReadLong
    EngineRun = True
    Nombres = True
    bRain = False
    Call SetConnected
    If bShowTutorial Then
        Call frmTutorial.Show(vbModeless)
    End If
    If ClientSetup.MostrarTips = True Then
        frmtip.Visible = True
    End If
    If ClientSetup.MostrarBindKeysSelection = True Then
        Call frmKeysConfigurationSelect.Show(vbModeless, frmMain)
        Call frmKeysConfigurationSelect.SetFocus
    End If
End Sub

Private Sub HandleRemoveDialogs()
    Call incomingData.ReadByte
    Call Dialogos.RemoveAllDialogs
End Sub

Private Sub HandleRemoveCharDialog()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    Call Dialogos.RemoveDialog(incomingData.ReadInteger())
End Sub

Private Sub HandleNavigateToggle()
    Call incomingData.ReadByte
    UserNavegando = Not UserNavegando
End Sub

Private Sub HandleDisconnect()
    Call incomingData.ReadByte
    Call CloseConnectionAndResetAllInfo
End Sub

Private Sub CloseConnectionAndResetAllInfo()
    Call ResetAllInfo(False)
    If CheckUserData() Then
        frmMain.Visible = False
        Call Protocol.Connect(E_MODO.Normal)
    End If
End Sub

Private Sub HandleCommerceEnd()
    Call incomingData.ReadByte
    Set InvComUsu = Nothing
    Set InvComNpc = Nothing
    Unload frmComerciar
    Comerciando = False
End Sub

Private Sub HandleBankEnd()
    Call incomingData.ReadByte
    Set InvBanco(0) = Nothing
    Set InvBanco(1) = Nothing
    Unload frmBancoObj
    Comerciando = False
End Sub

Private Sub HandleCommerceInit()
    Dim i As Long
    Call incomingData.ReadByte
    Set InvComUsu = New clsGraphicalInventory
    Set InvComNpc = New clsGraphicalInventory
    Call InvComUsu.Initialize(DirectD3D8, frmComerciar.picInvUser, MAX_INVENTORY_SLOTS, , , , , , , , True)
    Call InvComNpc.Initialize(DirectD3D8, frmComerciar.picInvNpc, MAX_NPC_INVENTORY_SLOTS)
    For i = 1 To MAX_INVENTORY_SLOTS
        If Inventario.ObjIndex(i) <> 0 Then
            With Inventario
                Call InvComUsu.SetItem(i, .ObjIndex(i), _
                .Amount(i), .Equipped(i), .GrhIndex(i), _
                .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                .Valor(i), .ItemName(i), .Incompatible(i))
            End With
        End If
    Next i
    For i = 1 To MAX_NPC_INVENTORY_SLOTS
        If NPCInventory(i).ObjIndex <> 0 Then
            With NPCInventory(i)
                Call InvComNpc.SetItem(i, .ObjIndex, _
                .Amount, 0, .GrhIndex, _
                .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, _
                .Valor, .Name, .Incompatible)
            End With
        End If
    Next i
    frmComerciar.Show , frmMain
    Call Audio.PlayWave("comerciante" & RandomNumber(1, 9) & ".wav")
End Sub

Private Sub HandleBankInit()
    Dim i As Long
    Dim BankGold As Long
    Call incomingData.ReadByte
    Set InvBanco(0) = New clsGraphicalInventory
    Set InvBanco(1) = New clsGraphicalInventory
    BankGold = incomingData.ReadLong
    Call InvBanco(0).Initialize(DirectD3D8, frmBancoObj.PicBancoInv, MAX_BANCOINVENTORY_SLOTS)
    Call InvBanco(1).Initialize(DirectD3D8, frmBancoObj.PicInv, MAX_INVENTORY_SLOTS, , , , , , , , True)
    For i = 1 To MAX_INVENTORY_SLOTS
        With Inventario
            Call InvBanco(1).SetItem(i, .ObjIndex(i), _
                .Amount(i), .Equipped(i), .GrhIndex(i), _
                .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                .Valor(i), .ItemName(i), .Incompatible(i))
        End With
    Next i
    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        With UserBancoInventory(i)
            Call InvBanco(0).SetItem(i, .ObjIndex, _
                .Amount, .Equipped, .GrhIndex, _
                .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, _
                .Valor, .Name, .Incompatible)
        End With
    Next i
    Comerciando = True
    frmBancoObj.lblUserGld.Caption = BankGold
    frmBancoObj.Show , frmMain
    Call Audio.PlayWave("banquero" & RandomNumber(1, 7) & ".wav")
End Sub

Private Sub HandleUserCommerceInit()
    Dim i As Long
    Call incomingData.ReadByte
    TradingUserName = incomingData.ReadASCIIString
    Set InvComUsu = New clsGraphicalInventory
    Set InvOfferComUsu(0) = New clsGraphicalInventory
    Set InvOfferComUsu(1) = New clsGraphicalInventory
    Set InvOroComUsu(0) = New clsGraphicalInventory
    Set InvOroComUsu(1) = New clsGraphicalInventory
    Set InvOroComUsu(2) = New clsGraphicalInventory
    Call InvComUsu.Initialize(DirectD3D8, frmComerciarUsu.picInvComercio, MAX_INVENTORY_SLOTS)
    Call InvOfferComUsu(0).Initialize(DirectD3D8, frmComerciarUsu.picInvOfertaProp, INV_OFFER_SLOTS)
    Call InvOfferComUsu(1).Initialize(DirectD3D8, frmComerciarUsu.picInvOfertaOtro, INV_OFFER_SLOTS)
    Call InvOroComUsu(0).Initialize(DirectD3D8, frmComerciarUsu.picInvOroProp, INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)
    Call InvOroComUsu(1).Initialize(DirectD3D8, frmComerciarUsu.picInvOroOfertaProp, INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)
    Call InvOroComUsu(2).Initialize(DirectD3D8, frmComerciarUsu.picInvOroOfertaOtro, INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)
    For i = 1 To MAX_INVENTORY_SLOTS
        If Inventario.ObjIndex(i) <> 0 Then
            With Inventario
                Call InvComUsu.SetItem(i, .ObjIndex(i), _
                .Amount(i), .Equipped(i), .GrhIndex(i), _
                .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                .Valor(i), .ItemName(i), .Incompatible(i))
            End With
        End If
    Next i
    Call InvOroComUsu(0).SetItem(1, ORO_INDEX, UserGLD, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro")
    Call InvOroComUsu(1).SetItem(1, ORO_INDEX, 0, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro")
    Call InvOroComUsu(2).SetItem(1, ORO_INDEX, 0, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro")
    Comerciando = True
    Call frmComerciarUsu.Show(vbModeless, frmMain)
End Sub

Private Sub HandleUserCommerceEnd()
    Call incomingData.ReadByte
    Set InvComUsu = Nothing
    Set InvOroComUsu(0) = Nothing
    Set InvOroComUsu(1) = Nothing
    Set InvOroComUsu(2) = Nothing
    Set InvOfferComUsu(0) = Nothing
    Set InvOfferComUsu(1) = Nothing
    Unload frmComerciarUsu
    Comerciando = False
End Sub

Private Sub HandleUserOfferConfirm()
    Call incomingData.ReadByte
    With frmComerciarUsu
        .HabilitarAceptarRechazar True
        .PrintCommerceMsg TradingUserName & JsonLanguage.item("MENSAJE_COMM_OFERTA_ACEPTA").item("TEXTO"), FontTypeNames.FONTTYPE_CONSE
    End With
End Sub

Private Sub HandleUpdateSta()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    UserMinSTA = incomingData.ReadInteger()
    frmMain.lblEnergia = UserMinSTA & "/" & UserMaxSTA
    Dim bWidth As Byte
    bWidth = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 83)
    frmMain.shpEnergia.Width = 83 - bWidth
    frmMain.shpEnergia.Left = 797 + (83 - frmMain.shpEnergia.Width)
    frmMain.shpEnergia.Visible = (bWidth <> 83)
End Sub

Private Sub HandleUpdateMana()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    UserMinMAN = incomingData.ReadInteger()
    frmMain.lblMana = UserMinMAN & "/" & UserMaxMAN
    Dim bWidth As Byte
    If UserMaxMAN > 0 Then _
        bWidth = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 90)
    frmMain.shpMana.Width = 90 - bWidth
    frmMain.shpMana.Left = 902 + (90 - frmMain.shpMana.Width)
    frmMain.shpMana.Visible = (bWidth <> 90)
End Sub

Private Sub HandleUpdateHP()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    UserMinHP = incomingData.ReadInteger()
    frmMain.lblVida = UserMinHP & "/" & UserMaxHP
    Dim bWidth As Byte
    bWidth = (((UserMinHP / 100) / (UserMaxHP / 100)) * 90)
    frmMain.shpVida.Width = 90 - bWidth
    frmMain.shpVida.Left = 902 + (90 - frmMain.shpVida.Width)
    frmMain.shpVida.Visible = (bWidth <> 90)
    If UserMinHP = 0 Then
        UserEstado = 1
        If frmMain.macrotrabajo Then Call frmMain.DesactivarMacroTrabajo
        UserEquitando = 0
        Call SetSpeedUsuario
    Else
        UserEstado = 0
    End If
End Sub

Private Sub HandleUpdateGold()
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    UserGLD = incomingData.ReadLong()
    Call frmMain.SetGoldColor
    frmMain.GldLbl.Caption = UserGLD
End Sub

Private Sub HandleUpdateBankGold()
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    frmBancoObj.lblUserGld.Caption = incomingData.ReadLong
End Sub

Private Sub HandleUpdateExp()
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    UserExp = incomingData.ReadLong()
    frmMain.UpdateProgressExperienceLevelBar (UserExp)
End Sub

Private Sub HandleUpdateStrenghtAndDexterity()
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    UserFuerza = incomingData.ReadByte
    UserAgilidad = incomingData.ReadByte
    frmMain.lblStrg.Caption = UserFuerza
    frmMain.lblDext.Caption = UserAgilidad
    frmMain.lblStrg.ForeColor = getStrenghtColor()
    frmMain.lblDext.ForeColor = getDexterityColor()
    IntervaloDopas = incomingData.ReadLong
    TiempoDopas = IntervaloDopas * 0.04
End Sub

Private Sub HandleUpdateStrenght()
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    UserFuerza = incomingData.ReadByte
    frmMain.lblStrg.Caption = UserFuerza
    frmMain.lblStrg.ForeColor = getStrenghtColor()
    IntervaloDopas = incomingData.ReadLong
    TiempoDopas = IntervaloDopas * 0.04
End Sub

Private Sub HandleUpdateDexterity()
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    UserAgilidad = incomingData.ReadByte
    frmMain.lblDext.Caption = UserAgilidad
    frmMain.lblDext.ForeColor = getDexterityColor()
    IntervaloDopas = incomingData.ReadLong
    TiempoDopas = IntervaloDopas * 0.04
End Sub

Private Sub HandleChangeMap()
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    UserMap = incomingData.ReadInteger()
    nameMap = incomingData.ReadASCIIString
    mapInfo.Zona = incomingData.ReadASCIIString
    mapInfo.Zona = UCase(mapInfo.Zona)
    Call incomingData.ReadInteger
    If FileExist(Game.path(Mapas) & "Mapa" & UserMap & ".map", vbNormal) Then
        Call SwitchMap(UserMap)
        If bRain And bLluvia(UserMap) = 0 Then
                Call Audio.StopWave(RainBufferIndex)
                RainBufferIndex = 0
                frmMain.IsPlaying = PlayLoop.plNone
        End If
    Else
        MsgBox JsonLanguage.item("ERROR_MAPAS").item("TEXTO")
        Call CloseClient
    End If
End Sub

Private Sub HandlePosUpdate()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    Call Map_RemoveOldUser
    Call Char_MapPosSet(incomingData.ReadByte(), incomingData.ReadByte())
    Call Char_UserPos
End Sub

Private Sub WriteChatOverHeadInConsole(ByVal CharIndex As Integer, ByVal ChatText As String, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte)
    Dim NameRed As Byte
    Dim NameGreen As Byte
    Dim NameBlue As Byte
    With charlist(CharIndex)
        If .priv = 0 Then
            If .Atacable Then
                NameRed = 236
                NameGreen = 89
                NameBlue = 57
            Else
                If .Criminal Then
                    NameRed = 247
                    NameGreen = 44
                    NameBlue = 0
                Else
                    NameRed = 218
                    NameGreen = 131
                    NameBlue = 225
                End If
            End If
         Else
            NameRed = 222
            NameGreen = 221
            NameBlue = 211
        End If
        Dim Pos As Integer
        Pos = InStr(.Nombre, "<")
        If Pos = 0 Then Pos = LenB(.Nombre) + 2
        Dim Name As String
        Name = Left$(.Nombre, Pos - 2)
        ChatText = Trim$(ChatText)
        If LenB(.Nombre) <> 0 And LenB(ChatText) > 0 Then
            Call AddtoRichTextBox(frmMain.RecTxt, Name & "> ", NameRed, NameGreen, NameBlue, True, False, True, rtfLeft)
            Call AddtoRichTextBox(frmMain.RecTxt, ChatText, Red, Green, Blue, True, False, False, rtfLeft)
        End If
    End With
End Sub

Private Sub HandleChatOverHead()
    If incomingData.Length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim chat As String
    Dim CharIndex As Integer
    Dim Red As Byte
    Dim Green As Byte
    Dim Blue As Byte
    chat = Buffer.ReadASCIIString()
    CharIndex = Buffer.ReadInteger()
    Red = Buffer.ReadByte()
    Green = Buffer.ReadByte()
    Blue = Buffer.ReadByte()
    If Char_Check(CharIndex) Then
        Call Dialogos.CreateDialog(Trim$(chat), CharIndex, RGB(Red, Green, Blue))
        Call WriteChatOverHeadInConsole(CharIndex, chat, Red, Green, Blue)
    End If
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleConsoleMessage()
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim chat As String
    Dim FontIndex As Integer
    Dim str As String
    Dim Red As Byte
    Dim Green As Byte
    Dim Blue As Byte
    chat = Buffer.ReadASCIIString()
    FontIndex = Buffer.ReadByte()
    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)
            If Val(str) > 255 Then
                Red = 255
            Else
                Red = Val(str)
            End If
            str = ReadField(3, chat, 126)
            If Val(str) > 255 Then
                Green = 255
            Else
                Green = Val(str)
            End If
            str = ReadField(4, chat, 126)
            If Val(str) > 255 Then
                Blue = 255
            Else
                Blue = Val(str)
            End If
        Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), Red, Green, Blue, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    Else
        With FontTypes(FontIndex)
            Call AddtoRichTextBox(frmMain.RecTxt, chat, .Red, .Green, .Blue, .bold, .italic)
        End With
        If FontIndex = FontTypeNames.FONTTYPE_PARTY Then
            If MirandoParty Then frmParty.SendTxt.SetFocus
        End If
    End If
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleGuildChat()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim chat As String
    Dim str As String
    Dim Red As Byte
    Dim Green As Byte
    Dim Blue As Byte
    chat = Buffer.ReadASCIIString()
    If Not DialogosClanes.Activo Then
        If InStr(1, chat, "~") Then
            str = ReadField(2, chat, 126)
            If Val(str) > 255 Then
                Red = 255
            Else
                Red = Val(str)
            End If
            str = ReadField(3, chat, 126)
            If Val(str) > 255 Then
                Green = 255
            Else
                Green = Val(str)
            End If
            str = ReadField(4, chat, 126)
            If Val(str) > 255 Then
                Blue = 255
            Else
                Blue = Val(str)
            End If
            Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), Red, Green, Blue, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
        Else
            With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
                Call AddtoRichTextBox(frmMain.RecTxt, chat, .Red, .Green, .Blue, .bold, .italic)
            End With
        End If
    Else
        Call DialogosClanes.PushBackText(ReadField(1, chat, 126))
    End If
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleCommerceChat()
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim chat As String
    Dim FontIndex As Integer
    Dim str As String
    Dim Red As Byte
    Dim Green As Byte
    Dim Blue As Byte
    chat = Buffer.ReadASCIIString()
    FontIndex = Buffer.ReadByte()
    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)
            If Val(str) > 255 Then
                Red = 255
            Else
                Red = Val(str)
            End If
            str = ReadField(3, chat, 126)
            If Val(str) > 255 Then
                Green = 255
            Else
                Green = Val(str)
            End If
            str = ReadField(4, chat, 126)
            If Val(str) > 255 Then
                Blue = 255
            Else
                Blue = Val(str)
            End If
        Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, Left$(chat, InStr(1, chat, "~") - 1), Red, Green, Blue, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    Else
        With FontTypes(FontIndex)
            Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, chat, .Red, .Green, .Blue, .bold, .italic)
        End With
    End If
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleShowMessageBox()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    frmMensaje.msg.Caption = Buffer.ReadASCIIString()
    frmMensaje.Show
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleUserIndexInServer()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    UserIndex = incomingData.ReadInteger()
End Sub

Private Sub HandleUserCharIndexInServer()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    Call Char_UserIndexSet(incomingData.ReadInteger())
    Call Char_UserPos
End Sub

Private Sub HandleCharacterCreate()
    If incomingData.Length < 24 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim CharIndex As Integer
    Dim Body As Integer
    Dim Head As Integer
    Dim Heading As E_Heading
    Dim X As Byte
    Dim Y As Byte
    Dim weapon As Integer
    Dim shield As Integer
    Dim helmet As Integer
    Dim fX As Integer
    Dim FXLoops As Integer
    Dim Name As String
    Dim NickColor As Byte
    Dim Privileges As Integer
    CharIndex = Buffer.ReadInteger()
    Body = Buffer.ReadInteger()
    Head = Buffer.ReadInteger()
    Heading = Buffer.ReadByte()
    X = Buffer.ReadByte()
    Y = Buffer.ReadByte()
    weapon = Buffer.ReadInteger()
    shield = Buffer.ReadInteger()
    helmet = Buffer.ReadInteger()
    fX = Buffer.ReadInteger()
    FXLoops = Buffer.ReadInteger()
    Name = Buffer.ReadASCIIString()
    NickColor = Buffer.ReadByte()
    Privileges = Buffer.ReadByte()
    Call incomingData.CopyBuffer(Buffer)
    With charlist(CharIndex)
        Call Char_SetFx(CharIndex, fX, FXLoops)
        .Nombre = Name
        .Clan = mid$(.Nombre, getTagPosition(.Nombre))
        If (NickColor And eNickColor.ieCriminal) <> 0 Then
            .Criminal = 1
        Else
            .Criminal = 0
        End If
        .Atacable = (NickColor And eNickColor.ieAtacable) <> 0
        If Privileges <> 0 Then
            If (Privileges And PlayerType.ChaosCouncil) <> 0 And (Privileges And PlayerType.User) = 0 Then
                Privileges = Privileges Xor PlayerType.ChaosCouncil
            End If
            If (Privileges And PlayerType.RoyalCouncil) <> 0 And (Privileges And PlayerType.User) = 0 Then
                Privileges = Privileges Xor PlayerType.RoyalCouncil
            End If
            If Privileges And PlayerType.RoleMaster Then
                Privileges = PlayerType.RoleMaster
            End If
            .priv = Log(Privileges) / Log(2)
        Else
            .priv = 0
        End If
    End With
    Call Char_Make(CharIndex, Body, Head, Heading, X, Y, weapon, shield, helmet)
ErrorHandler:
    Dim Error As Long
        Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleCharacterChangeNick()
    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    Dim CharIndex As Integer
    CharIndex = incomingData.ReadInteger
    Call Char_SetName(CharIndex, incomingData.ReadASCIIString)
End Sub

Private Sub HandleCharacterRemove()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    Dim CharIndex As Integer
    CharIndex = incomingData.ReadInteger()
    Call Char_Erase(CharIndex)
End Sub

Private Sub HandleCharacterMove()
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    Dim CharIndex As Integer
    Dim X As Byte
    Dim Y As Byte
    CharIndex = incomingData.ReadInteger()
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    With charlist(CharIndex)
        If .FxIndex >= 40 And .FxIndex <= 49 Then
            .FxIndex = 0
        End If
        If .priv <> 1 And .priv <> 2 And .priv <> 3 And .priv <> 5 And .priv <> 25 Then
            Call DoPasosFx(CharIndex)
        End If
    End With
    Call Char_MovebyPos(CharIndex, X, Y)
End Sub

Private Sub HandleForceCharMove()
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    Dim Direccion As Byte
    Direccion = incomingData.ReadByte()
    Call Char_MovebyHead(UserCharIndex, Direccion)
    Call Char_MoveScreen(Direccion)
End Sub

Private Sub HandleCharacterChange()
    If incomingData.Length < 17 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    Dim CharIndex As Integer
    CharIndex = incomingData.ReadInteger()
    Call Char_SetBody(CharIndex, incomingData.ReadInteger())
    Call Char_SetHead(CharIndex, incomingData.ReadInteger)
    Call Char_SetWeapon(CharIndex, incomingData.ReadInteger())
    Call Char_SetShield(CharIndex, incomingData.ReadInteger())
    Call Char_SetCasco(CharIndex, incomingData.ReadInteger())
    Call Char_SetFx(CharIndex, incomingData.ReadInteger(), incomingData.ReadInteger())
End Sub

Private Sub HandleHeadingChange()
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    Dim CharIndex As Integer
    CharIndex = incomingData.ReadInteger()
    Call Char_SetHeading(CharIndex, incomingData.ReadByte())
End Sub

Private Sub HandleObjectCreate()
    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    Dim X        As Byte
    Dim Y        As Byte
    Dim GrhIndex As Long
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    GrhIndex = incomingData.ReadLong()
    Call Map_CreateObject(X, Y, GrhIndex)
End Sub

Private Sub HandleObjectDelete()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    Dim X   As Byte
    Dim Y   As Byte
    Dim obj As Integer
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    obj = Map_PosExitsObject(X, Y)
    If (obj > 0) Then
        Call Map_DestroyObject(X, Y)
    End If
End Sub

Private Sub HandleBlockPosition()
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    Dim X As Byte
    Dim Y As Byte
    Dim block As Boolean
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    block = incomingData.ReadBoolean()
    If block Then
        Map_SetBlocked X, Y, 1
    Else
        Map_SetBlocked X, Y, 0
    End If
End Sub

Private Sub HandlePlayMP3()
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Dim currentMp3 As Integer
    Dim Loops As Integer
    Call incomingData.ReadByte
    currentMp3 = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    Call Audio.PlayBackgroundMusic(CStr(currentMp3), MusicTypes.Mp3)
End Sub

Private Sub HandlePlayMIDI()
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Dim currentMidi As Integer
    Dim Loops As Integer
    Call incomingData.ReadByte
    currentMidi = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    Call Audio.PlayBackgroundMusic(CStr(currentMidi), MusicTypes.Midi, Loops)
End Sub

Private Sub HandlePlayWave()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    Dim wave As Byte
    Dim srcX As Byte
    Dim srcY As Byte
    wave = incomingData.ReadByte()
    srcX = incomingData.ReadByte()
    srcY = incomingData.ReadByte()
    Call Audio.PlayWave(CStr(wave) & ".wav", srcX, srcY)
End Sub

Private Sub HandleGuildList()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    With frmGuildAdm
        .guildslist.Clear
        GuildNames = Split(Buffer.ReadASCIIString(), SEPARATOR)
        Dim i As Long
        For i = 0 To UBound(GuildNames())
            If LenB(GuildNames(i)) <> 0 Then
                Call .guildslist.AddItem(GuildNames(i))
            End If
        Next i
        Call incomingData.CopyBuffer(Buffer)
        .Show vbModeless, frmMain
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleAreaChanged()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    Dim X As Byte
    Dim Y As Byte
    X = incomingData.ReadByte()
    Y = incomingData.ReadByte()
    Call CambioDeArea(X, Y)
End Sub

Private Sub HandlePauseToggle()
    Call incomingData.ReadByte
    pausa = Not pausa
End Sub

Private Sub HandleRainToggle()
    Call incomingData.ReadByte
    If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
    bTecho = (MapData(UserPos.X, UserPos.Y).Trigger = eTrigger.BAJOTECHO Or _
        MapData(UserPos.X, UserPos.Y).Trigger = eTrigger.CASA Or _
        MapData(UserPos.X, UserPos.Y).Trigger = eTrigger.ZONASEGURA)
    If bRain And bLluvia(UserMap) Then
        Call Audio.StopWave(RainBufferIndex)
        RainBufferIndex = 0
        If bTecho Then
            Call Audio.PlayWave("lluviainend.wav", 0, 0, LoopStyle.Disabled)
        Else
            Call Audio.PlayWave("lluviaoutend.wav", 0, 0, LoopStyle.Disabled)
        End If
        frmMain.IsPlaying = PlayLoop.plNone
    End If
    bRain = Not bRain
End Sub

Private Sub HandleCreateFX()
    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    Dim CharIndex As Integer
    Dim fX As Integer
    Dim Loops As Integer
    CharIndex = incomingData.ReadInteger()
    fX = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    Call Char_SetFx(CharIndex, fX, Loops)
End Sub

Private Sub HandleUpdateUserStats()
    If incomingData.Length < 26 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    UserMaxHP = incomingData.ReadInteger()
    UserMinHP = incomingData.ReadInteger()
    UserMaxMAN = incomingData.ReadInteger()
    UserMinMAN = incomingData.ReadInteger()
    UserMaxSTA = incomingData.ReadInteger()
    UserMinSTA = incomingData.ReadInteger()
    UserGLD = incomingData.ReadLong()
    UserLvl = incomingData.ReadByte()
    UserPasarNivel = incomingData.ReadLong()
    UserExp = incomingData.ReadLong()
    frmMain.UpdateProgressExperienceLevelBar (UserExp)
    frmMain.GldLbl.Caption = UserGLD
    frmMain.lblLvl.Caption = UserLvl
    frmMain.lblMana = UserMinMAN & "/" & UserMaxMAN
    frmMain.lblVida = UserMinHP & "/" & UserMaxHP
    frmMain.lblEnergia = UserMinSTA & "/" & UserMaxSTA
    Dim bWidth As Byte
    If UserMaxMAN > 0 Then _
        bWidth = (((UserMinMAN / 100) / (UserMaxMAN / 100)) * 90)
    frmMain.shpMana.Width = 90 - bWidth
    frmMain.shpMana.Left = 902 + (90 - frmMain.shpMana.Width)
    frmMain.shpMana.Visible = (bWidth <> 90)
    bWidth = (((UserMinHP / 100) / (UserMaxHP / 100)) * 90)
    frmMain.shpVida.Width = 90 - bWidth
    frmMain.shpVida.Left = 902 + (90 - frmMain.shpVida.Width)
    frmMain.shpVida.Visible = (bWidth <> 90)
    bWidth = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 83)
    frmMain.shpEnergia.Width = 83 - bWidth
    frmMain.shpEnergia.Left = 797 + (83 - frmMain.shpEnergia.Width)
    frmMain.shpEnergia.Visible = (bWidth <> 83)
    If UserMinHP = 0 Then
        UserEstado = 1
        If frmMain.macrotrabajo Then Call frmMain.DesactivarMacroTrabajo
    Else
        UserEstado = 0
    End If
    Call frmMain.SetGoldColor
End Sub

Private Sub HandleChangeInventorySlot()
    If incomingData.Length < 24 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim slot As Byte
    Dim ObjIndex As Integer
    Dim Name As String
    Dim Amount As Integer
    Dim Equipped As Boolean
    Dim GrhIndex As Long
    Dim OBJType As Byte
    Dim MaxHit As Integer
    Dim MinHit As Integer
    Dim MaxDef As Integer
    Dim MinDef As Integer
    Dim Value As Single
    Dim Incompatible As Boolean
    slot = Buffer.ReadByte()
    ObjIndex = Buffer.ReadInteger()
    Name = Buffer.ReadASCIIString()
    Amount = Buffer.ReadInteger()
    Equipped = Buffer.ReadBoolean()
    GrhIndex = Buffer.ReadLong()
    OBJType = Buffer.ReadByte()
    MaxHit = Buffer.ReadInteger()
    MinHit = Buffer.ReadInteger()
    MaxDef = Buffer.ReadInteger()
    MinDef = Buffer.ReadInteger()
    Value = Buffer.ReadSingle()
    Incompatible = Buffer.ReadBoolean()
    If Equipped Then
        Select Case OBJType
            Case eObjType.otWeapon
                frmMain.lblWeapon = MinHit & "/" & MaxHit
                UserWeaponEqpSlot = slot
                
            Case eObjType.otArmadura
                frmMain.lblArmor = MinDef & "/" & MaxDef
                UserArmourEqpSlot = slot
                
            Case eObjType.otescudo
                frmMain.lblShielder = MinDef & "/" & MaxDef
                UserHelmEqpSlot = slot
                
            Case eObjType.otcasco
                frmMain.lblHelm = MinDef & "/" & MaxDef
                UserShieldEqpSlot = slot
        End Select
    Else
        Select Case slot
            Case UserWeaponEqpSlot
                frmMain.lblWeapon = "0/0"
                UserWeaponEqpSlot = 0
                
            Case UserArmourEqpSlot
                frmMain.lblArmor = "0/0"
                UserArmourEqpSlot = 0
                
            Case UserHelmEqpSlot
                frmMain.lblShielder = "0/0"
                UserHelmEqpSlot = 0
                
            Case UserShieldEqpSlot
                frmMain.lblHelm = "0/0"
                UserShieldEqpSlot = 0
        End Select
    End If
    Call Inventario.SetItem(slot, ObjIndex, Amount, Equipped, GrhIndex, OBJType, MaxHit, MinHit, MaxDef, MinDef, Value, Name, Incompatible)
    If frmComerciar.Visible Then
        Call InvComUsu.SetItem(slot, ObjIndex, Amount, Equipped, GrhIndex, OBJType, MaxHit, MinHit, MaxDef, MinDef, Value, Name, Incompatible)
    End If
    If frmBancoObj.Visible Then
        Call InvBanco(1).SetItem(slot, ObjIndex, Amount, Equipped, GrhIndex, OBJType, MaxHit, MinHit, MaxDef, MinDef, Value, Name, Incompatible)
        frmBancoObj.NoPuedeMover = False
    End If
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleAddSlots()
    Call incomingData.ReadByte
    MaxInventorySlots = incomingData.ReadByte
    Call Inventario.DrawInventory
End Sub

Private Sub HandleStopWorking()
    Call incomingData.ReadByte
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_WORK_FINISHED"), .Red, .Green, .Blue, .bold, .italic)
    End With
    If frmMain.trainingMacro.Enabled Then Call frmMain.DesactivarMacroHechizos
    If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
End Sub

Private Sub HandleCancelOfferItem()
    Dim slot As Byte
    Dim Amount As Long
    Call incomingData.ReadByte
    slot = incomingData.ReadByte
    With InvOfferComUsu(0)
        Amount = .Amount(slot)
        If Amount <> 0 Then
            Call frmComerciarUsu.UpdateInvCom(.ObjIndex(slot), Amount)
            Call .SetItem(slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "")
        End If
    End With
    If Not frmComerciarUsu.HasAnyItem(InvOfferComUsu(0)) And Not frmComerciarUsu.HasAnyItem(InvOroComUsu(1)) Then
        Call frmComerciarUsu.HabilitarConfirmar(False)
    End If
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call frmComerciarUsu.PrintCommerceMsg(JsonLanguage.item("MENSAJE_NO_COMM_OBJETO").item("TEXTO"), FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

Private Sub HandleChangeBankSlot()
    If incomingData.Length < 23 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim slot As Byte
    slot = Buffer.ReadByte()
    With UserBancoInventory(slot)
        .ObjIndex = Buffer.ReadInteger()
        .Name = Buffer.ReadASCIIString()
        .Amount = Buffer.ReadInteger()
        .GrhIndex = Buffer.ReadLong()
        .OBJType = Buffer.ReadByte()
        .MaxHit = Buffer.ReadInteger()
        .MinHit = Buffer.ReadInteger()
        .MaxDef = Buffer.ReadInteger()
        .MinDef = Buffer.ReadInteger()
        .Valor = Buffer.ReadLong()
        .Incompatible = Buffer.ReadBoolean()
        If frmBancoObj.Visible Then
            Call InvBanco(0).SetItem(slot, .ObjIndex, .Amount, 0, .GrhIndex, .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, .Name, .Incompatible)
        End If
    End With
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleChangeSpellSlot()
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim slot As Byte
    slot = Buffer.ReadByte()
    Dim str As String
    UserHechizos(slot) = Buffer.ReadInteger()
    If slot <= frmMain.hlst.ListCount Then
         str = DevolverNombreHechizo(UserHechizos(slot))
        If str <> vbNullString Then
            frmMain.hlst.List(slot - 1) = str
        Else
            Call frmMain.hlst.AddItem(JsonLanguage.item("NADA").item("TEXTO"))
        End If
    Else
        str = DevolverNombreHechizo(UserHechizos(slot))
        If str <> vbNullString Then
            Call frmMain.hlst.AddItem(str)
        Else
            Call frmMain.hlst.AddItem(JsonLanguage.item("NADA").item("TEXTO"))
        End If
    End If
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleAtributes()
    If incomingData.Length < 1 + NUMATRIBUTES Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    Dim i As Long
    For i = 1 To NUMATRIBUTES
        UserAtributos(i) = incomingData.ReadByte()
    Next i
    If EstadoLogin = E_MODO.Dados Then
        With frmCrearPersonaje
            If .Visible Then
                For i = 1 To NUMATRIBUTES
                    .lblAtributos(i).Caption = UserAtributos(i)
                Next i
                .UpdateStats
            End If
        End With
    Else
        LlegaronAtrib = True
    End If
End Sub

Private Sub HandleBlacksmithWeapons()
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim Count As Integer
    Dim i As Long
    Dim J As Long
    Dim k As Long
    Count = Buffer.ReadInteger()
    ReDim ArmasHerrero(Count) As tItemsConstruibles
    ReDim HerreroMejorar(0) As tItemsConstruibles
    For i = 1 To Count
        With ArmasHerrero(i)
            .Name = Buffer.ReadASCIIString()
            .GrhIndex = Buffer.ReadLong()
            .LinH = Buffer.ReadInteger()
            .LinP = Buffer.ReadInteger()
            .LinO = Buffer.ReadInteger()
            .ObjIndex = Buffer.ReadInteger()
            .Upgrade = Buffer.ReadInteger()
        End With
    Next i
    Call incomingData.CopyBuffer(Buffer)
    If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
        Call WriteCraftBlacksmith(MacroBltIndex)
        Exit Sub
    Else
        Call frmHerrero.Show(vbModeless, frmMain)
        MirandoHerreria = True
    End If
    For i = 1 To MAX_LIST_ITEMS
        Set InvLingosHerreria(i) = New clsGraphicalInventory
    Next i
    With frmHerrero
        Call InvLingosHerreria(1).Initialize(DirectD3D8, .picLingotes0, 3, , , , , , False)
        Call InvLingosHerreria(2).Initialize(DirectD3D8, .picLingotes1, 3, , , , , , False)
        Call InvLingosHerreria(3).Initialize(DirectD3D8, .picLingotes2, 3, , , , , , False)
        Call InvLingosHerreria(4).Initialize(DirectD3D8, .picLingotes3, 3, , , , , , False)
        Call .HideExtraControls(Count)
        Call .RenderList(1, True)
    End With
    For i = 1 To Count
        With ArmasHerrero(i)
            If .Upgrade Then
                For k = 1 To Count
                    If .Upgrade = ArmasHerrero(k).ObjIndex Then
                        J = J + 1
                        ReDim Preserve HerreroMejorar(J) As tItemsConstruibles
                        HerreroMejorar(J).Name = .Name
                        HerreroMejorar(J).GrhIndex = .GrhIndex
                        HerreroMejorar(J).ObjIndex = .ObjIndex
                        HerreroMejorar(J).UpgradeName = ArmasHerrero(k).Name
                        HerreroMejorar(J).UpgradeGrhIndex = ArmasHerrero(k).GrhIndex
                        HerreroMejorar(J).LinH = ArmasHerrero(k).LinH - .LinH * 0.85
                        HerreroMejorar(J).LinP = ArmasHerrero(k).LinP - .LinP * 0.85
                        HerreroMejorar(J).LinO = ArmasHerrero(k).LinO - .LinO * 0.85
                        Exit For
                    End If
                Next k
            End If
        End With
    Next i
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleBlacksmithArmors()
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim Count As Integer
    Dim i As Long
    Dim J As Long
    Dim k As Long
    Count = Buffer.ReadInteger()
    ReDim ArmadurasHerrero(Count) As tItemsConstruibles
    For i = 1 To Count
        With ArmadurasHerrero(i)
            .Name = Buffer.ReadASCIIString()
            .GrhIndex = Buffer.ReadLong()
            .LinH = Buffer.ReadInteger()
            .LinP = Buffer.ReadInteger()
            .LinO = Buffer.ReadInteger()
            .ObjIndex = Buffer.ReadInteger()
            .Upgrade = Buffer.ReadInteger()
        End With
    Next i
    Call incomingData.CopyBuffer(Buffer)
    If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
        Call WriteCraftBlacksmith(MacroBltIndex)
                Exit Sub
    Else
        Call frmHerrero.Show(vbModeless, frmMain)
        MirandoHerreria = True
    End If
    J = UBound(HerreroMejorar)
    For i = 1 To Count
        With ArmadurasHerrero(i)
            If .Upgrade Then
                For k = 1 To Count
                    If .Upgrade = ArmadurasHerrero(k).ObjIndex Then
                        J = J + 1
                        ReDim Preserve HerreroMejorar(J) As tItemsConstruibles
                        HerreroMejorar(J).Name = .Name
                        HerreroMejorar(J).GrhIndex = .GrhIndex
                        HerreroMejorar(J).ObjIndex = .ObjIndex
                        HerreroMejorar(J).UpgradeName = ArmadurasHerrero(k).Name
                        HerreroMejorar(J).UpgradeGrhIndex = ArmadurasHerrero(k).GrhIndex
                        HerreroMejorar(J).LinH = ArmadurasHerrero(k).LinH - .LinH * 0.85
                        HerreroMejorar(J).LinP = ArmadurasHerrero(k).LinP - .LinP * 0.85
                        HerreroMejorar(J).LinO = ArmadurasHerrero(k).LinO - .LinO * 0.85
                        Exit For
                    End If
                Next k
            End If
        End With
    Next i
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleInitCarpenting()
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim Count As Integer
    Dim i As Long
    Dim J As Long
    Dim k As Long
    Count = Buffer.ReadInteger()
    ReDim ObjCarpintero(Count) As tItemsConstruibles
    ReDim CarpinteroMejorar(0) As tItemsConstruibles
    For i = 1 To Count
        With ObjCarpintero(i)
            .Name = Buffer.ReadASCIIString()
            .GrhIndex = Buffer.ReadLong()
            .Madera = Buffer.ReadInteger()
            .MaderaElfica = Buffer.ReadInteger()
            .ObjIndex = Buffer.ReadInteger()
            .Upgrade = Buffer.ReadInteger()
        End With
    Next i
    Call incomingData.CopyBuffer(Buffer)
    If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
        Call WriteCraftCarpenter(MacroBltIndex)
        Exit Sub
    Else
        Call frmCarpinteria.Show(vbModeless, frmMain)
        MirandoCarpinteria = True
    End If
    For i = 1 To MAX_LIST_ITEMS
        Set InvMaderasCarpinteria(i) = New clsGraphicalInventory
    Next i
    With frmCarpinteria
        Call InvMaderasCarpinteria(1).Initialize(DirectD3D8, .picMaderas0, 2, , , , , , False)
        Call InvMaderasCarpinteria(2).Initialize(DirectD3D8, .picMaderas1, 2, , , , , , False)
        Call InvMaderasCarpinteria(3).Initialize(DirectD3D8, .picMaderas2, 2, , , , , , False)
        Call InvMaderasCarpinteria(4).Initialize(DirectD3D8, .picMaderas3, 2, , , , , , False)
        Call .HideExtraControls(Count)
        Call .RenderList(1)
    End With
    For i = 1 To Count
        With ObjCarpintero(i)
            If .Upgrade Then
                For k = 1 To Count
                    If .Upgrade = ObjCarpintero(k).ObjIndex Then
                        J = J + 1
                        ReDim Preserve CarpinteroMejorar(J) As tItemsConstruibles
                        CarpinteroMejorar(J).Name = .Name
                        CarpinteroMejorar(J).GrhIndex = .GrhIndex
                        CarpinteroMejorar(J).ObjIndex = .ObjIndex
                        CarpinteroMejorar(J).UpgradeName = ObjCarpintero(k).Name
                        CarpinteroMejorar(J).UpgradeGrhIndex = ObjCarpintero(k).GrhIndex
                        CarpinteroMejorar(J).Madera = ObjCarpintero(k).Madera - .Madera * 0.85
                        CarpinteroMejorar(J).MaderaElfica = ObjCarpintero(k).MaderaElfica - .MaderaElfica * 0.85
                        Exit For
                    End If
                Next k
            End If
        End With
    Next i
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Public Sub HandleInitCraftman()
    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim CountObjs As Integer
    Dim CountCrafteo As Integer
    Dim i As Long
    Dim J As Long
    Dim k As Long
    frmArtesano.ArtesaniaCosto = Buffer.ReadLong()
    CountObjs = Buffer.ReadInteger()
    ReDim ObjArtesano(CountObjs) As tItemArtesano
    For i = 1 To CountObjs
        With ObjArtesano(i)
            .Name = Buffer.ReadASCIIString()
            .GrhIndex = Buffer.ReadLong()
            .ObjIndex = Buffer.ReadInteger()
            CountCrafteo = Buffer.ReadByte()
            ReDim .ItemsCrafteo(CountCrafteo) As tItemCrafteo
            For J = 1 To CountCrafteo
                .ItemsCrafteo(J).Name = Buffer.ReadASCIIString()
                .ItemsCrafteo(J).GrhIndex = Buffer.ReadLong()
                .ItemsCrafteo(J).ObjIndex = Buffer.ReadInteger()
                .ItemsCrafteo(J).Amount = Buffer.ReadInteger()
            Next J
        End With
    Next i
    Call incomingData.CopyBuffer(Buffer)
    Call frmArtesano.Show(vbModeless, frmMain)
    For i = 1 To MAX_LIST_ITEMS
        Set InvObjArtesano(i) = New clsGraphicalInventory
    Next i
    With frmArtesano
        Call InvObjArtesano(1).Initialize(DirectD3D8, .picObj0, MAX_ITEMS_CRAFTEO, , , , , , False)
        Call InvObjArtesano(2).Initialize(DirectD3D8, .picObj1, MAX_ITEMS_CRAFTEO, , , , , , False)
        Call InvObjArtesano(3).Initialize(DirectD3D8, .picObj2, MAX_ITEMS_CRAFTEO, , , , , , False)
        Call InvObjArtesano(4).Initialize(DirectD3D8, .picObj3, MAX_ITEMS_CRAFTEO, , , , , , False)
        Call .HideExtraControls(CountObjs)
        Call .RenderList(1)
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleRestOK()
    Call incomingData.ReadByte
    UserDescansar = Not UserDescansar
End Sub

Private Sub HandleErrorMessage()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Call MsgBox(Buffer.ReadASCIIString())
    If frmConnect.Visible And (Not frmCrearPersonaje.Visible) Then
        frmMain.Client.CloseSck
    End If
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleBlind()
    Call incomingData.ReadByte
    UserCiego = True
End Sub

Private Sub HandleDumb()
    Call incomingData.ReadByte
    UserEstupido = True
End Sub

Private Sub HandleShowSignal()
    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim tmp As String
    tmp = Buffer.ReadASCIIString()
    Call InitCartel(tmp, Buffer.ReadLong())
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleChangeNPCInventorySlot()
    If incomingData.Length < 23 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim slot As Byte
    slot = Buffer.ReadByte()
    With NPCInventory(slot)
        .Name = Buffer.ReadASCIIString()
        .Amount = Buffer.ReadInteger()
        .Valor = Buffer.ReadSingle()
        .GrhIndex = Buffer.ReadLong()
        .ObjIndex = Buffer.ReadInteger()
        .OBJType = Buffer.ReadByte()
        .MaxHit = Buffer.ReadInteger()
        .MinHit = Buffer.ReadInteger()
        .MaxDef = Buffer.ReadInteger()
        .MinDef = Buffer.ReadInteger()
        .Incompatible = Buffer.ReadBoolean()
        If frmComerciar.Visible Then
            Call InvComNpc.SetItem(slot, .ObjIndex, .Amount, 0, .GrhIndex, .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, .Valor, .Name, .Incompatible)
        End If
    End With
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleUpdateHungerAndThirst()
    If incomingData.Length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    UserMaxAGU = incomingData.ReadByte()
    UserMinAGU = incomingData.ReadByte()
    UserMaxHAM = incomingData.ReadByte()
    UserMinHAM = incomingData.ReadByte()
    frmMain.lblHambre = UserMinHAM & "/" & UserMaxHAM
    frmMain.lblSed = UserMinAGU & "/" & UserMaxAGU
    Dim bWidth As Byte
    bWidth = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 83)
    frmMain.shpHambre.Width = 83 - bWidth
    frmMain.shpHambre.Left = 797 + (83 - frmMain.shpHambre.Width)
    frmMain.shpHambre.Visible = (bWidth <> 83)
    bWidth = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 83)
    frmMain.shpSed.Width = 83 - bWidth
    frmMain.shpSed.Left = 797 + (83 - frmMain.shpSed.Width)
    frmMain.shpSed.Visible = (bWidth <> 83)
End Sub

Private Sub HandleFame()
    If incomingData.Length < 29 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    With UserReputacion
        .AsesinoRep = incomingData.ReadLong()
        .BandidoRep = incomingData.ReadLong()
        .BurguesRep = incomingData.ReadLong()
        .LadronesRep = incomingData.ReadLong()
        .NobleRep = incomingData.ReadLong()
        .PlebeRep = incomingData.ReadLong()
        .Promedio = incomingData.ReadLong()
    End With
    LlegoFama = True
End Sub

Private Sub HandleMiniStats()
    If incomingData.Length < 20 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    With UserEstadisticas
        .CiudadanosMatados = incomingData.ReadLong()
        .CriminalesMatados = incomingData.ReadLong()
        .UsuariosMatados = incomingData.ReadLong()
        .NpcsMatados = incomingData.ReadInteger()
        .Clase = ListaClases(incomingData.ReadByte())
        .PenaCarcel = incomingData.ReadLong()
    End With
End Sub

Private Sub HandleLevelUp()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    SkillPoints = SkillPoints + incomingData.ReadInteger()
    Call frmMain.LightSkillStar(True)
End Sub

Private Sub HandleAddForumMessage()
    If incomingData.Length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim ForumType As eForumMsgType
    Dim Title As String
    Dim Message As String
    Dim Author As String
    ForumType = Buffer.ReadByte
    Title = Buffer.ReadASCIIString()
    Author = Buffer.ReadASCIIString()
    Message = Buffer.ReadASCIIString()
    If Not frmForo.ForoLimpio Then
        clsForos.ClearForums
        frmForo.ForoLimpio = True
    End If
    Call clsForos.AddPost(ForumAlignment(ForumType), Title, Author, Message, EsAnuncio(ForumType))
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleShowForumForm()
    Call incomingData.ReadByte
    frmForo.Privilegios = incomingData.ReadByte
    frmForo.CanPostSticky = incomingData.ReadByte
    If Not MirandoForo Then
        frmForo.Show , frmMain
    End If
End Sub

Private Sub HandleSetInvisible()
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    Dim CharIndex As Integer
    Dim timeRemaining As Integer
    CharIndex = incomingData.ReadInteger()
    UserInvisible = incomingData.ReadBoolean()
    Call Char_SetInvisible(CharIndex, UserInvisible)
    If CharIndex = UserCharIndex Then
        If UserInvisible And TiempoInvi <= 0 Then
            TiempoInvi = (IntervaloInvi * 0.05) - 1
        Else
            TiempoInvi = 0
        End If
    End If
End Sub

Private Sub HandleDiceRoll()
    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    UserAtributos(eAtributos.Fuerza) = incomingData.ReadByte()
    UserAtributos(eAtributos.Agilidad) = incomingData.ReadByte()
    UserAtributos(eAtributos.Inteligencia) = incomingData.ReadByte()
    UserAtributos(eAtributos.Carisma) = incomingData.ReadByte()
    UserAtributos(eAtributos.Constitucion) = incomingData.ReadByte()
    With frmCrearPersonaje
        .lblAtributos(eAtributos.Fuerza) = UserAtributos(eAtributos.Fuerza)
        .lblAtributos(eAtributos.Agilidad) = UserAtributos(eAtributos.Agilidad)
        .lblAtributos(eAtributos.Inteligencia) = UserAtributos(eAtributos.Inteligencia)
        .lblAtributos(eAtributos.Carisma) = UserAtributos(eAtributos.Carisma)
        .lblAtributos(eAtributos.Constitucion) = UserAtributos(eAtributos.Constitucion)
        .UpdateStats
    End With
End Sub

Private Sub HandleMeditateToggle()
    Call incomingData.ReadByte
    UserMeditar = Not UserMeditar
End Sub

Private Sub HandleBlindNoMore()
    Call incomingData.ReadByte
    UserCiego = False
End Sub

Private Sub HandleDumbNoMore()
    Call incomingData.ReadByte
    UserEstupido = False
End Sub

Private Sub HandleSendSkills()
    If incomingData.Length < 2 + NUMSKILLS * 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    UserClase = incomingData.ReadByte
    Dim i As Long
    For i = 1 To NUMSKILLS
        UserSkills(i) = incomingData.ReadByte()
        PorcentajeSkills(i) = incomingData.ReadByte()
    Next i
    LlegaronSkills = True
End Sub

Private Sub HandleTrainerCreatureList()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim creatures() As String
    Dim i As Long
    Dim Upper_creatures As Long
    creatures = Split(Buffer.ReadASCIIString(), SEPARATOR)
    Upper_creatures = UBound(creatures())
    For i = 0 To Upper_creatures
        Call frmEntrenador.lstCriaturas.AddItem(creatures(i))
    Next i
    frmEntrenador.Show , frmMain
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleGuildNews()
    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim guildList() As String
    Dim Upper_guildList As Long
    Dim i As Long
    Dim sTemp As String
    frmGuildNews.news = Buffer.ReadASCIIString()
    guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    Upper_guildList = UBound(guildList)
    For i = 0 To Upper_guildList
        sTemp = frmGuildNews.txtClanesGuerra.Text
        frmGuildNews.txtClanesGuerra.Text = sTemp & guildList(i) & vbCrLf
    Next i
    guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    For i = 0 To Upper_guildList
        sTemp = frmGuildNews.txtClanesAliados.Text
        frmGuildNews.txtClanesAliados.Text = sTemp & guildList(i) & vbCrLf
    Next i
    If ClientSetup.bGuildNews Or bShowGuildNews Then frmGuildNews.Show vbModeless, frmMain
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleOfferDetails()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Call frmUserRequest.recievePeticion(Buffer.ReadASCIIString())
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleAlianceProposalsList()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim vsGuildList() As String, Upper_vsGuildList As Long
    Dim i As Long
    vsGuildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    Upper_vsGuildList = UBound(vsGuildList())
    Call frmPeaceProp.lista.Clear
    For i = 0 To Upper_vsGuildList
        Call frmPeaceProp.lista.AddItem(vsGuildList(i))
    Next i
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.ALIANZA
    Call frmPeaceProp.Show(vbModeless, frmMain)
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandlePeaceProposalsList()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim guildList()     As String
    Dim Upper_guildList As Long
    Dim i               As Long
    guildList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    With frmPeaceProp
        .lista.Clear
        Upper_guildList = UBound(guildList())
        For i = 0 To Upper_guildList
            .lista.AddItem (guildList(i))
        Next i
        .ProposalType = TIPO_PROPUESTA.PAZ
        .Show vbModeless, frmMain
    End With
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
    On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleCharacterInfo()
    If incomingData.Length < 35 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    With frmCharInfo
        If .frmType = CharInfoFrmType.frmMembers Then
            .imgRechazar.Visible = False
            .imgAceptar.Visible = False
            .imgEchar.Visible = True
            .imgPeticion.Visible = False
        Else
            .imgRechazar.Visible = True
            .imgAceptar.Visible = True
            .imgEchar.Visible = False
            .imgPeticion.Visible = True
        End If
        .Nombre.Caption = Buffer.ReadASCIIString()
        .Raza.Caption = ListaRazas(Buffer.ReadByte())
        .Clase.Caption = ListaClases(Buffer.ReadByte())
        If Buffer.ReadByte() = 1 Then
            .Genero.Caption = "Hombre"
        Else
            .Genero.Caption = "Mujer"
        End If
        .Nivel.Caption = Buffer.ReadByte()
        .Oro.Caption = Buffer.ReadLong()
        .Banco.Caption = Buffer.ReadLong()
        Dim reputation As Long
        reputation = Buffer.ReadLong()
        .reputacion.Caption = reputation
        .txtPeticiones.Text = Buffer.ReadASCIIString()
        .guildactual.Caption = Buffer.ReadASCIIString()
        .txtMiembro.Text = Buffer.ReadASCIIString()
        Dim armada As Boolean
        Dim caos As Boolean
        armada = Buffer.ReadBoolean()
        caos = Buffer.ReadBoolean()
        If armada Then
            .ejercito.Caption = JsonLanguage.item("ARMADA").item("TEXTO")
        ElseIf caos Then
            .ejercito.Caption = JsonLanguage.item("LEGION").item("TEXTO")
        End If
        .Ciudadanos.Caption = CStr(Buffer.ReadLong())
        .criminales.Caption = CStr(Buffer.ReadLong())
        If reputation > 0 Then
            .status.Caption = " " & JsonLanguage.item("CIUDADANO").item("TEXTO")
            .status.ForeColor = vbBlue
        Else
            .status.Caption = " " & JsonLanguage.item("CRIMINAL").item("TEXTO")
            .status.ForeColor = vbRed
        End If
        Call .Show(vbModeless, frmMain)
    End With
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleGuildLeaderInfo()
    If incomingData.Length < 9 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim i As Long
    Dim List() As String
    With frmGuildLeader
        GuildNames = Split(Buffer.ReadASCIIString(), SEPARATOR)
        Call .guildslist.Clear
        For i = 0 To UBound(GuildNames())
            If LenB(GuildNames(i)) <> 0 Then
                Call .guildslist.AddItem(GuildNames(i))
            End If
        Next i
        GuildMembers = Split(Buffer.ReadASCIIString(), SEPARATOR)
        .Miembros.Caption = CStr(UBound(GuildMembers()) + 1)
        Call .members.Clear
        For i = 0 To UBound(GuildMembers())
            Call .members.AddItem(GuildMembers(i))
        Next i
        .txtguildnews = Buffer.ReadASCIIString()
        List = Split(Buffer.ReadASCIIString(), SEPARATOR)
        Call .solicitudes.Clear
        For i = 0 To UBound(List())
            Call .solicitudes.AddItem(List(i))
        Next i
        .Show , frmMain
    End With
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleGuildDetails()
    If incomingData.Length < 26 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    With frmGuildBrief
        .imgDeclararGuerra.Visible = .EsLeader
        .imgOfrecerAlianza.Visible = .EsLeader
        .imgOfrecerPaz.Visible = .EsLeader
        .Nombre.Caption = Buffer.ReadASCIIString()
        .fundador.Caption = Buffer.ReadASCIIString()
        .creacion.Caption = Buffer.ReadASCIIString()
        .lider.Caption = Buffer.ReadASCIIString()
        .web.Caption = Buffer.ReadASCIIString()
        .Miembros.Caption = Buffer.ReadInteger()
        If Buffer.ReadBoolean() Then
            .eleccion.Caption = UCase$(JsonLanguage.item("ABIERTA").item("TEXTO"))
        Else
            .eleccion.Caption = UCase$(JsonLanguage.item("CERRADA").item("TEXTO"))
        End If
        .lblAlineacion.Caption = Buffer.ReadASCIIString()
        .Enemigos.Caption = Buffer.ReadInteger()
        .Aliados.Caption = Buffer.ReadInteger()
        .antifaccion.Caption = Buffer.ReadASCIIString()
        Dim codexStr() As String
        Dim i As Long
        codexStr = Split(Buffer.ReadASCIIString(), SEPARATOR)
        For i = 0 To 7
            .Codex(i).Caption = codexStr(i)
        Next i
        .Desc.Text = Buffer.ReadASCIIString()
    End With
    Call incomingData.CopyBuffer(Buffer)
    frmGuildBrief.Show vbModeless, frmMain
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleShowGuildAlign()
    Call incomingData.ReadByte
    frmEligeAlineacion.Show vbModeless, frmMain
End Sub

Private Sub HandleShowGuildFundationForm()
    Call incomingData.ReadByte
    CreandoClan = True
    frmGuildFoundation.Show , frmMain
End Sub

Private Sub HandleParalizeOK()
    Call incomingData.ReadByte
    Dim timeRemaining As Integer
    UserParalizado = Not UserParalizado
    timeRemaining = incomingData.ReadInteger()
    UserParalizadoSegundosRestantes = IIf(timeRemaining > 0, (timeRemaining * 0.04), 0)
    If UserParalizado And timeRemaining > 0 Then frmMain.timerPasarSegundo.Enabled = True
End Sub

Private Sub HandleShowUserRequest()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Call frmUserRequest.recievePeticion(Buffer.ReadASCIIString())
    Call frmUserRequest.Show(vbModeless, frmMain)
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleChangeUserTradeSlot()
    If incomingData.Length < 24 Then
        Call Err.Raise(incomingData.NotEnoughDataErrCode)
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    With Buffer
        Call Buffer.ReadByte
        Dim OfferSlot As Byte: OfferSlot = .ReadByte()
        Dim ObjIndex  As Integer: ObjIndex = .ReadInteger()
        Dim Amount    As Long: Amount = .ReadLong()
        Dim GrhIndex  As Long: GrhIndex = .ReadLong()
        Dim OBJType   As Byte: OBJType = .ReadByte()
        Dim MaxHit    As Integer: MaxHit = .ReadInteger()
        Dim MinHit    As Integer: MinHit = .ReadInteger()
        Dim MaxDef    As Integer: MaxDef = .ReadInteger()
        Dim MinDef    As Integer: MinDef = .ReadInteger()
        Dim SalePrice As Long: SalePrice = .ReadLong()
        Dim Name      As String: Name = .ReadASCIIString()
        Dim Incompatible As Boolean: Incompatible = .ReadBoolean()
    End With
    Call incomingData.CopyBuffer(Buffer)
    If OfferSlot = GOLD_OFFER_SLOT Then
        Call InvOroComUsu(2).SetItem(1, ObjIndex, Amount, 0, GrhIndex, OBJType, MaxHit, MinHit, MaxDef, MinDef, SalePrice, Name, Incompatible)
    Else
        Call InvOfferComUsu(1).SetItem(OfferSlot, ObjIndex, Amount, 0, GrhIndex, OBJType, MaxHit, MinHit, MaxDef, MinDef, SalePrice, Name, Incompatible)
    End If
    Call frmComerciarUsu.PrintCommerceMsg(TradingUserName & JsonLanguage.item("MENSAJE_COMM_OFERTA_CAMBIA").item("TEXTO"), FontTypeNames.FONTTYPE_VENENO)
ErrorHandler:
    Dim Error As Long
    On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then Call Err.Raise(Error)
        Call Err.Raise(Error)
End Sub

Private Sub HandleSendNight()
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    Dim tBool As Boolean
    tBool = incomingData.ReadBoolean()
End Sub

Private Sub HandleSpawnList()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim creatureList() As String
    Dim i As Long
    Dim Upper_creatureList As Long
    creatureList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    Upper_creatureList = UBound(creatureList())
    For i = 0 To Upper_creatureList
        Call frmSpawnList.lstCriaturas.AddItem(creatureList(i))
    Next i
    frmSpawnList.Show , frmMain
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleShowSOSForm()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim sosList() As String
    Dim i As Long
    Dim Upper_sosList As Long
    sosList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    Upper_sosList = UBound(sosList())
    For i = 0 To Upper_sosList
        Call frmMSG.List1.AddItem(sosList(i))
    Next i
    frmMSG.Show , frmMain
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleShowDenounces()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim DenounceList() As String
    Dim Upper_denounceList As Long
    Dim DenounceIndex As Long
    DenounceList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    Upper_denounceList = UBound(DenounceList())
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        For DenounceIndex = 0 To Upper_denounceList
            Call AddtoRichTextBox(frmMain.RecTxt, DenounceList(DenounceIndex), .Red, .Green, .Blue, .bold, .italic)
        Next DenounceIndex
    End With
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleShowPartyForm()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim members() As String
    Dim Upper_members As Long
    Dim i As Long
    EsPartyLeader = CBool(Buffer.ReadByte())
    members = Split(Buffer.ReadASCIIString(), SEPARATOR)
    Upper_members = UBound(members())
    For i = 0 To Upper_members
        Call frmParty.lstMembers.AddItem(members(i))
    Next i
    frmParty.lblTotalExp.Caption = Buffer.ReadLong
    frmParty.Show , frmMain
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleShowMOTDEditionForm()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    frmCambiaMotd.txtMotd.Text = Buffer.ReadASCIIString()
    frmCambiaMotd.Show , frmMain
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleShowGMPanelForm()
    Call incomingData.ReadByte
    frmPanelGm.Show vbModeless, frmMain
End Sub

Private Sub HandleUserNameList()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim userList() As String
    userList = Split(Buffer.ReadASCIIString(), SEPARATOR)
    If frmPanelGm.Visible Then
        frmPanelGm.cboListaUsus.Clear
        Dim i As Long
        Dim Upper_userlist As Long
            Upper_userlist = UBound(userList())
        For i = 0 To Upper_userlist
            Call frmPanelGm.cboListaUsus.AddItem(userList(i))
        Next i
        If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.ListIndex = 0
    End If
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Public Sub HandleRenderMsg()
    Call incomingData.ReadByte
    renderMsgReset
    renderText = incomingData.ReadASCIIString
    renderFont = incomingData.ReadInteger
    colorRender = 240
End Sub

Private Sub HandlePong()
    Call incomingData.ReadByte
    Dim MENSAJE_PING As String
        MENSAJE_PING = JsonLanguage.item("MENSAJE_PING").item("TEXTO")
        MENSAJE_PING = Replace$(MENSAJE_PING, "VAR_PING", (GetTickCount() - pingTime))
    Call AddtoRichTextBox(frmMain.RecTxt, _
                            MENSAJE_PING, _
                            JsonLanguage.item("MENSAJE_PING").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_PING").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_PING").item("COLOR").item(3), _
                            True, False, True)
    pingTime = 0
End Sub

Private Sub HandleGuildMemberInfo()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    With frmGuildMember
        .lstClanes.Clear
        GuildNames = Split(Buffer.ReadASCIIString(), SEPARATOR)
        Dim i As Long
        For i = 0 To UBound(GuildNames())
            If LenB(GuildNames(i)) <> 0 Then
                Call .lstClanes.AddItem(GuildNames(i))
            End If
        Next i
        GuildMembers = Split(Buffer.ReadASCIIString(), SEPARATOR)
        .lblCantMiembros.Caption = CStr(UBound(GuildMembers()) + 1)
        Call .lstMiembros.Clear
        For i = 0 To UBound(GuildMembers())
            Call .lstMiembros.AddItem(GuildMembers(i))
        Next i
        Call incomingData.CopyBuffer(Buffer)
        .Show vbModeless, frmMain
    End With
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleUpdateTagAndStatus()
    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim CharIndex As Integer
    Dim NickColor As Byte
    Dim UserTag As String
    CharIndex = Buffer.ReadInteger()
    NickColor = Buffer.ReadByte()
    UserTag = Buffer.ReadASCIIString()
    With charlist(CharIndex)
        If (NickColor And eNickColor.ieCriminal) <> 0 Then
            .Criminal = 1
        Else
            .Criminal = 0
        End If
        .Atacable = (NickColor And eNickColor.ieAtacable) <> 0
        .Nombre = UserTag
        .Clan = mid$(.Nombre, getTagPosition(.Nombre))
    End With
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Public Sub WriteLoginExistingAccount()
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginExistingAccount)
        Call .WriteASCIIString(AccountName)
        Call .WriteASCIIString(AccountPassword)
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
    End With
End Sub

Public Sub WriteDeleteChar()
    With outgoingData
        Call .WriteByte(ClientPacketID.DeleteChar)
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(AccountHash)
    End With
End Sub

Public Sub WriteLoginExistingChar()
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginExistingChar)
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(AccountHash)
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
    End With
End Sub

Public Sub WriteLoginNewAccount()
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginNewAccount)
        Call .WriteASCIIString(AccountName)
        Call .WriteASCIIString(AccountPassword)
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
    End With
End Sub

Public Sub WriteThrowDices()
    Call outgoingData.WriteByte(ClientPacketID.ThrowDices)
End Sub

Public Sub WriteLoginNewChar()
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginNewChar)
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(AccountHash)
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteByte(UserRaza)
        Call .WriteByte(UserSexo)
        Call .WriteByte(UserClase)
        Call .WriteInteger(UserHead)
        Call .WriteByte(UserHogar)
    End With
End Sub

Public Sub WriteTalk(ByVal chat As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.Talk)
        Call .WriteASCIIString(chat)
    End With
End Sub

Public Sub WriteYell(ByVal chat As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.Yell)
        Call .WriteASCIIString(chat)
    End With
End Sub

Public Sub WriteWhisper(ByVal CharName As String, ByVal chat As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.Whisper)
        Call .WriteASCIIString(CharName)
        Call .WriteASCIIString(chat)
    End With
End Sub

Public Sub WriteWalk(ByVal Heading As E_Heading)
    With outgoingData
        Call .WriteByte(ClientPacketID.Walk)
        Call .WriteByte(Heading)
    End With
End Sub

Public Sub WriteRequestPositionUpdate()
    Call outgoingData.WriteByte(ClientPacketID.RequestPositionUpdate)
End Sub

Public Sub WriteAttack()
    Call outgoingData.WriteByte(ClientPacketID.Attack)
    charlist(UserCharIndex).attacking = True
End Sub

Public Sub WritePickUp()
    Call outgoingData.WriteByte(ClientPacketID.PickUp)
End Sub

Public Sub WriteSafeToggle()
    Call outgoingData.WriteByte(ClientPacketID.SafeToggle)
End Sub

Public Sub WriteResuscitationToggle()
    Call outgoingData.WriteByte(ClientPacketID.ResuscitationSafeToggle)
End Sub

Public Sub WriteRequestGuildLeaderInfo()
    Call outgoingData.WriteByte(ClientPacketID.RequestGuildLeaderInfo)
End Sub

Public Sub WriteRequestPartyForm()
    Call outgoingData.WriteByte(ClientPacketID.RequestPartyForm)
End Sub

Public Sub WriteItemUpgrade(ByVal ItemIndex As Integer)
    Call outgoingData.WriteByte(ClientPacketID.ItemUpgrade)
    Call outgoingData.WriteInteger(ItemIndex)
End Sub

Public Sub WriteRequestAtributes()
    Call outgoingData.WriteByte(ClientPacketID.RequestAtributes)
End Sub

Public Sub WriteRequestFame()
    Call outgoingData.WriteByte(ClientPacketID.RequestFame)
End Sub

Public Sub WriteRequestSkills()
    Call outgoingData.WriteByte(ClientPacketID.RequestSkills)
End Sub

Public Sub WriteRequestMiniStats()
    Call outgoingData.WriteByte(ClientPacketID.RequestMiniStats)
End Sub

Public Sub WriteCommerceEnd()
    Call outgoingData.WriteByte(ClientPacketID.CommerceEnd)
End Sub

Public Sub WriteUserCommerceEnd()
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceEnd)
End Sub

Public Sub WriteUserCommerceConfirm()
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceConfirm)
End Sub

Public Sub WriteBankEnd()
    Call outgoingData.WriteByte(ClientPacketID.BankEnd)
End Sub

Public Sub WriteUserCommerceOk()
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceOk)
End Sub
Public Sub WriteUserCommerceReject()
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceReject)
End Sub

Public Sub WriteDrop(ByVal slot As Byte, ByVal Amount As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.Drop)
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

Public Sub WriteCastSpell(ByVal slot As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.CastSpell)
        Call .WriteByte(slot)
    End With
End Sub

Public Sub WriteLeftClick(ByVal X As Byte, ByVal Y As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.LeftClick)
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

Public Sub WriteDoubleClick(ByVal X As Byte, ByVal Y As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.DoubleClick)
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

Public Sub WriteWork(ByVal Skill As eSkill)
    With outgoingData
        Call .WriteByte(ClientPacketID.Work)
        Call .WriteByte(Skill)
    End With
End Sub

Public Sub WriteUseSpellMacro()
    Call outgoingData.WriteByte(ClientPacketID.UseSpellMacro)
End Sub

Public Sub WriteUseItem(ByVal slot As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.UseItem)
        Call .WriteByte(slot)
    End With
End Sub

Public Sub WriteCraftBlacksmith(ByVal item As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftBlacksmith)
        Call .WriteInteger(item)
    End With
End Sub

Public Sub WriteCraftCarpenter(ByVal item As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftCarpenter)
        Call .WriteInteger(item)
    End With
End Sub

Public Sub WriteCraftsmanCreate(ByVal item As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftsmanCreate)
        Call .WriteInteger(item)
    End With
End Sub

Public Sub WriteShowGuildNews()
     outgoingData.WriteByte (ClientPacketID.ShowGuildNews)
End Sub

Public Sub WriteWorkLeftClick(ByVal X As Byte, ByVal Y As Byte, ByVal Skill As eSkill)
    With outgoingData
        Call .WriteByte(ClientPacketID.WorkLeftClick)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteByte(Skill)
    End With
End Sub

Public Sub WriteCreateNewGuild(ByVal Desc As String, ByVal Name As String, ByVal Site As String, ByRef Codex() As String)
    Dim temp As String
    Dim i As Long
    Dim Lower_codex As Long, Upper_codex As Long
    With outgoingData
        Call .WriteByte(ClientPacketID.CreateNewGuild)
        Call .WriteASCIIString(Desc)
        Call .WriteASCIIString(Name)
        Call .WriteASCIIString(Site)
        Lower_codex = LBound(Codex())
        Upper_codex = UBound(Codex())
        For i = Lower_codex To Upper_codex
            temp = temp & Codex(i) & SEPARATOR
        Next i
        If Len(temp) Then _
            temp = Left$(temp, Len(temp) - 1)
        Call .WriteASCIIString(temp)
    End With
End Sub

Public Sub WriteEquipItem(ByVal slot As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.EquipItem)
        Call .WriteByte(slot)
    End With
End Sub

Public Sub WriteChangeHeading(ByVal Heading As E_Heading)
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeHeading)
        Call .WriteByte(Heading)
    End With
End Sub

Public Sub WriteModifySkills(ByRef skillEdt() As Byte)
    Dim i As Long
    With outgoingData
        Call .WriteByte(ClientPacketID.ModifySkills)
        For i = 1 To NUMSKILLS
            Call .WriteByte(skillEdt(i))
        Next i
    End With
End Sub

Public Sub WriteTrain(ByVal creature As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.Train)
        Call .WriteByte(creature)
    End With
End Sub

Public Sub WriteCommerceBuy(ByVal slot As Byte, ByVal Amount As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceBuy)
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

Public Sub WriteBankExtractItem(ByVal slot As Byte, ByVal Amount As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractItem)
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

Public Sub WriteCommerceSell(ByVal slot As Byte, ByVal Amount As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceSell)
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

Public Sub WriteBankDeposit(ByVal slot As Byte, ByVal Amount As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDeposit)
        Call .WriteByte(slot)
        Call .WriteInteger(Amount)
    End With
End Sub

Public Sub WriteForumPost(ByVal Title As String, ByVal Message As String, ByVal ForumMsgType As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.ForumPost)
        Call .WriteByte(ForumMsgType)
        Call .WriteASCIIString(Title)
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteMoveSpell(ByVal upwards As Boolean, ByVal slot As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveSpell)
        Call .WriteBoolean(upwards)
        Call .WriteByte(slot)
    End With
End Sub

Public Sub WriteMoveBank(ByVal upwards As Boolean, ByVal slot As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveBank)
        Call .WriteBoolean(upwards)
        Call .WriteByte(slot)
    End With
End Sub

Public Sub WriteClanCodexUpdate(ByVal Desc As String, ByRef Codex() As String)
    Dim temp As String
    Dim i As Long
    Dim Lower_codex As Long, Upper_codex As Long
    With outgoingData
        Call .WriteByte(ClientPacketID.ClanCodexUpdate)
        Call .WriteASCIIString(Desc)
        Lower_codex = LBound(Codex())
        Upper_codex = UBound(Codex())
        For i = Lower_codex To Upper_codex
            temp = temp & Codex(i) & SEPARATOR
        Next i
        If Len(temp) Then _
            temp = Left$(temp, Len(temp) - 1)
        Call .WriteASCIIString(temp)
    End With
End Sub

Public Sub WriteUserCommerceOffer(ByVal slot As Byte, ByVal Amount As Long, ByVal OfferSlot As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.UserCommerceOffer)
        Call .WriteByte(slot)
        Call .WriteLong(Amount)
        Call .WriteByte(OfferSlot)
    End With
End Sub

Public Sub WriteCommerceChat(ByVal chat As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceChat)
        Call .WriteASCIIString(chat)
    End With
End Sub

Public Sub WriteGuildAcceptPeace(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptPeace)
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteGuildRejectAlliance(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectAlliance)
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteGuildRejectPeace(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectPeace)
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteGuildAcceptAlliance(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptAlliance)
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteGuildOfferPeace(ByVal guild As String, ByVal proposal As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildOfferPeace)
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)
    End With
End Sub

Public Sub WriteGuildOfferAlliance(ByVal guild As String, ByVal proposal As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildOfferAlliance)
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)
    End With
End Sub

Public Sub WriteGuildAllianceDetails(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAllianceDetails)
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteGuildPeaceDetails(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildPeaceDetails)
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteGuildRequestJoinerInfo(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestJoinerInfo)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteGuildAlliancePropList()
    Call outgoingData.WriteByte(ClientPacketID.GuildAlliancePropList)
End Sub

Public Sub WriteGuildPeacePropList()
    Call outgoingData.WriteByte(ClientPacketID.GuildPeacePropList)
End Sub

Public Sub WriteGuildDeclareWar(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildDeclareWar)
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteGuildNewWebsite(ByVal URL As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildNewWebsite)
        Call .WriteASCIIString(URL)
    End With
End Sub

Public Sub WriteGuildAcceptNewMember(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptNewMember)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteGuildRejectNewMember(ByVal UserName As String, ByVal Reason As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectNewMember)
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(Reason)
    End With
End Sub

Public Sub WriteGuildKickMember(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildKickMember)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteGuildUpdateNews(ByVal news As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildUpdateNews)
        Call .WriteASCIIString(news)
    End With
End Sub

Public Sub WriteGuildMemberInfo(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMemberInfo)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteGuildOpenElections()
    Call outgoingData.WriteByte(ClientPacketID.GuildOpenElections)
End Sub

Public Sub WriteGuildRequestMembership(ByVal guild As String, ByVal Application As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestMembership)
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(Application)
    End With
End Sub

Public Sub WriteGuildRequestDetails(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestDetails)
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteOnline()
    Call outgoingData.WriteByte(ClientPacketID.Online)
End Sub

Public Sub WriteDiscord(ByVal chat As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.Discord)
        Call .WriteASCIIString(chat)
    End With
End Sub

Public Sub WriteQuit()
    Call outgoingData.WriteByte(ClientPacketID.Quit)
End Sub

Public Sub WriteGuildLeave()
    Call outgoingData.WriteByte(ClientPacketID.GuildLeave)
End Sub

Public Sub WriteRequestAccountState()
    Call outgoingData.WriteByte(ClientPacketID.RequestAccountState)
End Sub

Public Sub WritePetStand()
    Call outgoingData.WriteByte(ClientPacketID.PetStand)
End Sub

Public Sub WritePetFollow()
    Call outgoingData.WriteByte(ClientPacketID.PetFollow)
End Sub

Public Sub WriteReleasePet()
    Call outgoingData.WriteByte(ClientPacketID.ReleasePet)
End Sub

Public Sub WriteTrainList()
    Call outgoingData.WriteByte(ClientPacketID.TrainList)
End Sub

Public Sub WriteRest()
    Call outgoingData.WriteByte(ClientPacketID.Rest)
End Sub

Public Sub WriteMeditate()
    Call outgoingData.WriteByte(ClientPacketID.Meditate)
End Sub

Public Sub WriteResucitate()
    Call outgoingData.WriteByte(ClientPacketID.Resucitate)
End Sub

Public Sub WriteConsultation()
    Call outgoingData.WriteByte(ClientPacketID.Consultation)
End Sub

Public Sub WriteHeal()
    Call outgoingData.WriteByte(ClientPacketID.Heal)
End Sub

Public Sub WriteHelp()
    Call outgoingData.WriteByte(ClientPacketID.Help)
End Sub

Public Sub WriteRequestStats()
    Call outgoingData.WriteByte(ClientPacketID.RequestStats)
End Sub

Public Sub WriteCommerceStart()
    Call outgoingData.WriteByte(ClientPacketID.CommerceStart)
End Sub

Public Sub WriteBankStart()
    Call outgoingData.WriteByte(ClientPacketID.BankStart)
End Sub

Public Sub WriteEnlist()
    Call outgoingData.WriteByte(ClientPacketID.Enlist)
End Sub

Public Sub WriteInformation()
    Call outgoingData.WriteByte(ClientPacketID.Information)
End Sub

Public Sub WriteReward()
    Call outgoingData.WriteByte(ClientPacketID.Reward)
End Sub

Public Sub WriteRequestMOTD()
    Call outgoingData.WriteByte(ClientPacketID.RequestMOTD)
End Sub

Public Sub WriteUpTime()
    Call outgoingData.WriteByte(ClientPacketID.UpTime)
End Sub

Public Sub WritePartyLeave()
    Call outgoingData.WriteByte(ClientPacketID.PartyLeave)
End Sub

Public Sub WritePartyCreate()
    Call outgoingData.WriteByte(ClientPacketID.PartyCreate)
End Sub

Public Sub WritePartyJoin()
    Call outgoingData.WriteByte(ClientPacketID.PartyJoin)
End Sub

Public Sub WriteInquiry()
    Call outgoingData.WriteByte(ClientPacketID.Inquiry)
End Sub

Public Sub WriteGuildMessage(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMessage)
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WritePartyMessage(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyMessage)
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteCentinelReport(ByVal Clave As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.CentinelReport)
        Call .WriteASCIIString(Clave)
    End With
End Sub

Public Sub WriteGuildOnline()
    Call outgoingData.WriteByte(ClientPacketID.GuildOnline)
End Sub

Public Sub WritePartyOnline()
    Call outgoingData.WriteByte(ClientPacketID.PartyOnline)
End Sub

Public Sub WriteCouncilMessage(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.CouncilMessage)
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteRoleMasterRequest(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.RoleMasterRequest)
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteGMRequest()
    Call outgoingData.WriteByte(ClientPacketID.GMRequest)
End Sub

Public Sub WriteBugReport(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.bugReport)
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteChangeDescription(ByVal Desc As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeDescription)
        Call .WriteASCIIString(Desc)
    End With
End Sub

Public Sub WriteGuildVote(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildVote)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WritePunishments(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.Punishments)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteChangePassword(ByRef oldPass As String, ByRef newPass As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangePassword)
        Call .WriteASCIIString(oldPass)
        Call .WriteASCIIString(newPass)
    End With
End Sub

Public Sub WriteGamble(ByVal Amount As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.Gamble)
        Call .WriteInteger(Amount)
    End With
End Sub

Public Sub WriteInquiryVote(ByVal opt As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.InquiryVote)
        Call .WriteByte(opt)
    End With
End Sub

Public Sub WriteLeaveFaction()
    Call outgoingData.WriteByte(ClientPacketID.LeaveFaction)
End Sub

Public Sub WriteBankExtractGold(ByVal Amount As Long)
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractGold)
        Call .WriteLong(Amount)
    End With
End Sub

Public Sub WriteBankDepositGold(ByVal Amount As Long)
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDepositGold)
        Call .WriteLong(Amount)
    End With
End Sub

Public Sub WriteDenounce(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.Denounce)
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteGuildFundate()
    Call outgoingData.WriteByte(ClientPacketID.GuildFundate)
End Sub

Public Sub WriteGuildFundation(ByVal clanType As eClanType)
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildFundation)
        Call .WriteByte(clanType)
    End With
End Sub

Public Sub WritePartyKick(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyKick)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WritePartySetLeader(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.PartySetLeader)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WritePartyAcceptMember(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.PartyAcceptMember)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteGuildMemberList(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GuildMemberList)
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteInitCrafting(ByVal cantidad As Long, ByVal NroPorCiclo As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.InitCrafting)
        Call .WriteLong(cantidad)
        Call .WriteInteger(NroPorCiclo)
    End With
End Sub

Public Sub WriteHome()
    With outgoingData
        Call .WriteByte(ClientPacketID.Home)
    End With
End Sub

Public Sub WriteGMMessage(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GMMessage)
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteShowName()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.showName)
End Sub

Public Sub WriteOnlineRoyalArmy()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.OnlineRoyalArmy)
End Sub

Public Sub WriteOnlineChaosLegion()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.OnlineChaosLegion)
End Sub

Public Sub WriteGoNearby(ByVal UserName As String)
    With outgoingData
        Call outgoingData.WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GoNearby)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteComment(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.comment)
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteServerTime()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.serverTime)
End Sub

Public Sub WriteWhere(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Where)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteCreaturesInMap(ByVal Map As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreaturesInMap)
        Call .WriteInteger(Map)
    End With
End Sub

Public Sub WriteWarpMeToTarget()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.WarpMeToTarget)
End Sub

Public Sub WriteWarpChar(ByVal UserName As String, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)
    If X = UserPos.X And Y = UserPos.Y And Map = UserMap Then Exit Sub
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.WarpChar)
        Call .WriteASCIIString(UserName)
        Call .WriteInteger(Map)
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

Public Sub WriteSilence(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Silence)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteSOSShowList()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SOSShowList)
End Sub

Public Sub WriteSOSRemove(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SOSRemove)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteGoToChar(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GoToChar)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteInvisible()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.invisible)
End Sub

Public Sub WriteGMPanel()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.GMPanel)
End Sub

Public Sub WriteRequestUserList()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.RequestUserList)
End Sub

Public Sub WriteWorking()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Working)
End Sub

Public Sub WriteHiding()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Hiding)
End Sub

Public Sub WriteJail(ByVal UserName As String, ByVal Reason As String, ByVal Time As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Jail)
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(Reason)
        Call .WriteByte(Time)
    End With
End Sub

Public Sub WriteKillNPC()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KillNPC)
End Sub

Public Sub WriteWarnUser(ByVal UserName As String, ByVal Reason As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.WarnUser)
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(Reason)
    End With
End Sub

Public Sub WriteEditChar(ByVal UserName As String, ByVal EditOption As eEditOptions, ByVal arg1 As String, ByVal arg2 As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.EditChar)
        Call .WriteASCIIString(UserName)
        Call .WriteByte(EditOption)
        Call .WriteASCIIString(arg1)
        Call .WriteASCIIString(arg2)
    End With
End Sub

Public Sub WriteRequestCharInfo(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharInfo)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteRequestCharStats(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharStats)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteRequestCharGold(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharGold)
        Call .WriteASCIIString(UserName)
    End With
End Sub
    
Public Sub WriteRequestCharInventory(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharInventory)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteRequestCharBank(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharBank)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteRequestCharSkills(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharSkills)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteReviveChar(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ReviveChar)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteOnlineGM()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.OnlineGM)
End Sub

Public Sub WriteOnlineMap(ByVal Map As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.OnlineMap)
        Call .WriteInteger(Map)
    End With
End Sub

Public Sub WriteForgive(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Forgive)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteKick(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Kick)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteExecute(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.Execute)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteBanChar(ByVal UserName As String, ByVal Reason As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.BanChar)
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(Reason)
    End With
End Sub

Public Sub WriteUnbanChar(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.UnbanChar)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteNPCFollow()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.NPCFollow)
End Sub

Public Sub WriteSummonChar(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SummonChar)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteSpawnListRequest()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SpawnListRequest)
End Sub

Public Sub WriteSpawnCreature(ByVal creatureIndex As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SpawnCreature)
        Call .WriteInteger(creatureIndex)
    End With
End Sub

Public Sub WriteResetNPCInventory()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ResetNPCInventory)
End Sub

Public Sub WriteServerMessage(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ServerMessage)
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteMapMessage(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.MapMessage)
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteNickToIP(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.NickToIP)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteIPToNick(ByRef Ip() As Byte)
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub
    Dim i As Long
    Dim Upper_ip As Long, Lower_ip As Long
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.IPToNick)
        Lower_ip = LBound(Ip())
        Upper_ip = UBound(Ip())
        For i = Lower_ip To Upper_ip
            Call .WriteByte(Ip(i))
        Next i
    End With
End Sub

Public Sub WriteGuildOnlineMembers(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GuildOnlineMembers)
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteTeleportCreate(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal Radio As Byte = 0)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.TeleportCreate)
        Call .WriteInteger(Map)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteByte(Radio)
    End With
End Sub

Public Sub WriteTeleportDestroy()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.TeleportDestroy)
End Sub

Public Sub WriteExitDestroy()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ExitDestroy)
End Sub

Public Sub WriteRainToggle()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.RainToggle)
End Sub

Public Sub WriteSetCharDescription(ByVal Desc As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetCharDescription)
        Call .WriteASCIIString(Desc)
    End With
End Sub

Public Sub WriteForceMP3ToMap(ByVal Mp3Id As Byte, ByVal Map As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceMP3ToMap)
        Call .WriteByte(Mp3Id)
        Call .WriteInteger(Map)
    End With
End Sub

Public Sub WriteForceMIDIToMap(ByVal midiID As Byte, ByVal Map As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceMIDIToMap)
        Call .WriteByte(midiID)
        Call .WriteInteger(Map)
    End With
End Sub

Public Sub WriteForceWAVEToMap(ByVal waveID As Byte, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceWAVEToMap)
        Call .WriteByte(waveID)
        Call .WriteInteger(Map)
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

Public Sub WriteRoyalArmyMessage(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RoyalArmyMessage)
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteChaosLegionMessage(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChaosLegionMessage)
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteCitizenMessage(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CitizenMessage)
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteCriminalMessage(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CriminalMessage)
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteTalkAsNPC(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.TalkAsNPC)
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteDestroyAllItemsInArea()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DestroyAllItemsInArea)
End Sub

Public Sub WriteAcceptRoyalCouncilMember(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AcceptRoyalCouncilMember)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteAcceptChaosCouncilMember(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AcceptChaosCouncilMember)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteItemsInTheFloor()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ItemsInTheFloor)
End Sub

Public Sub WriteMakeDumb(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.MakeDumb)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteMakeDumbNoMore(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.MakeDumbNoMore)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteDumpIPTables()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DumpIPTables)
End Sub

Public Sub WriteCouncilKick(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CouncilKick)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteSetTrigger(ByVal Trigger As eTrigger)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetTrigger)
        Call .WriteByte(Trigger)
    End With
End Sub

Public Sub WriteAskTrigger()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.AskTrigger)
End Sub

Public Sub WriteBannedIPList()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.BannedIPList)
End Sub

Public Sub WriteBannedIPReload()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.BannedIPReload)
End Sub

Public Sub WriteGuildBan(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.GuildBan)
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteBanIP(ByVal byIp As Boolean, ByRef Ip() As Byte, ByVal Nick As String, ByVal Reason As String)
    If byIp And UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.BanIP)
        Call .WriteBoolean(byIp)
        If byIp Then
            Dim i As Long
            Dim Upper_ip As Long, Lower_ip As Long
            Lower_ip = LBound(Ip())
            Upper_ip = UBound(Ip())
            For i = Lower_ip To Upper_ip
                Call .WriteByte(Ip(i))
            Next i
        Else
            Call .WriteASCIIString(Nick)
        End If
        Call .WriteASCIIString(Reason)
    End With
End Sub

Public Sub WriteUnbanIP(ByRef Ip() As Byte)
    If UBound(Ip()) - LBound(Ip()) + 1 <> 4 Then Exit Sub
    Dim i As Long
    Dim Upper_ip As Long, Lower_ip As Long
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.UnbanIP)
        Lower_ip = LBound(Ip())
        Upper_ip = UBound(Ip())
        For i = Lower_ip To Upper_ip
            Call .WriteByte(Ip(i))
        Next i
    End With
End Sub

Public Sub WriteCreateItem(ByVal ItemIndex As Long, ByVal cantidad As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreateItem)
        Call .WriteInteger(ItemIndex)
        Call .WriteInteger(cantidad)
    End With
End Sub

Public Sub WriteDestroyItems()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DestroyItems)
End Sub

Public Sub WriteChaosLegionKick(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChaosLegionKick)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteRoyalArmyKick(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RoyalArmyKick)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteForceMP3All(ByVal Mp3Id As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceMP3All)
        Call .WriteByte(Mp3Id)
    End With
End Sub

Public Sub WriteForceMIDIAll(ByVal midiID As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceMIDIAll)
        Call .WriteByte(midiID)
    End With
End Sub

Public Sub WriteForceWAVEAll(ByVal waveID As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ForceWAVEAll)
        Call .WriteByte(waveID)
    End With
End Sub

Public Sub WriteRemovePunishment(ByVal UserName As String, ByVal punishment As Byte, ByVal NewText As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RemovePunishment)
        Call .WriteASCIIString(UserName)
        Call .WriteByte(punishment)
        Call .WriteASCIIString(NewText)
    End With
End Sub

Public Sub WriteTileBlockedToggle()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.TileBlockedToggle)
End Sub

Public Sub WriteKillNPCNoRespawn()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KillNPCNoRespawn)
End Sub

Public Sub WriteKillAllNearbyNPCs()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KillAllNearbyNPCs)
End Sub

Public Sub WriteLastIP(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.LastIP)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteChangeMOTD()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ChangeMOTD)
End Sub

Public Sub WriteSetMOTD(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetMOTD)
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteSystemMessage(ByVal Message As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SystemMessage)
        Call .WriteASCIIString(Message)
    End With
End Sub

Public Sub WriteCreateNPC(ByVal NPCIndex As Integer, ByVal WithRespawn As Boolean)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreateNPC)
        Call .WriteInteger(NPCIndex)
        Call .WriteBoolean(WithRespawn)
    End With
End Sub

Public Sub WriteImperialArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ImperialArmour)
        Call .WriteByte(armourIndex)
        Call .WriteInteger(objectIndex)
    End With
End Sub

Public Sub WriteChaosArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChaosArmour)
        Call .WriteByte(armourIndex)
        Call .WriteInteger(objectIndex)
    End With
End Sub

Public Sub WriteNavigateToggle()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.NavigateToggle)
End Sub

Public Sub WriteServerOpenToUsersToggle()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ServerOpenToUsersToggle)
End Sub

Public Sub WriteTurnOffServer()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.TurnOffServer)
End Sub

Public Sub WriteTurnCriminal(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.TurnCriminal)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteResetFactions(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ResetFactions)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteRemoveCharFromGuild(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RemoveCharFromGuild)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteRequestCharMail(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RequestCharMail)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteAlterPassword(ByVal UserName As String, ByVal CopyFrom As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AlterPassword)
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(CopyFrom)
    End With
End Sub

Public Sub WriteAlterMail(ByVal UserName As String, ByVal newMail As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AlterMail)
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(newMail)
    End With
End Sub

Public Sub WriteAlterName(ByVal UserName As String, ByVal newName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.AlterName)
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(newName)
    End With
End Sub

Public Sub WriteToggleCentinelActivated()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ToggleCentinelActivated)
End Sub

Public Sub WriteDoBackup()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.DoBackUp)
End Sub

Public Sub WriteShowGuildMessages(ByVal guild As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ShowGuildMessages)
        Call .WriteASCIIString(guild)
    End With
End Sub

Public Sub WriteSaveMap()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SaveMap)
End Sub

Public Sub WriteChangeMapInfoPK(ByVal isPK As Boolean)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoPK)
        Call .WriteBoolean(isPK)
    End With
End Sub

Public Sub WriteChangeMapInfoNoOcultar(ByVal PermitirOcultar As Boolean)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoOcultar)
        Call .WriteBoolean(PermitirOcultar)
    End With
End Sub

Public Sub WriteChangeMapInfoNoInvocar(ByVal PermitirInvocar As Boolean)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoInvocar)
        Call .WriteBoolean(PermitirInvocar)
    End With
End Sub

Public Sub WriteChangeMapInfoBackup(ByVal backup As Boolean)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoBackup)
        Call .WriteBoolean(backup)
    End With
End Sub

Public Sub WriteChangeMapInfoRestricted(ByVal restrict As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoRestricted)
        Call .WriteASCIIString(restrict)
    End With
End Sub

Public Sub WriteChangeMapInfoNoMagic(ByVal nomagic As Boolean)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoMagic)
        Call .WriteBoolean(nomagic)
    End With
End Sub

Public Sub WriteChangeMapInfoNoInvi(ByVal noinvi As Boolean)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoInvi)
        Call .WriteBoolean(noinvi)
    End With
End Sub
                            
Public Sub WriteChangeMapInfoNoResu(ByVal noresu As Boolean)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoNoResu)
        Call .WriteBoolean(noresu)
    End With
End Sub
                        
Public Sub WriteChangeMapInfoLand(ByVal land As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoLand)
        Call .WriteASCIIString(land)
    End With
End Sub
                        
Public Sub WriteChangeMapInfoZone(ByVal zone As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoZone)
        Call .WriteASCIIString(zone)
    End With
End Sub

Public Sub WriteChangeMapInfoStealNpc(ByVal forbid As Boolean)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChangeMapInfoStealNpc)
        Call .WriteBoolean(forbid)
    End With
End Sub

Public Sub WriteSaveChars()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.SaveChars)
End Sub

Public Sub WriteCleanSOS()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.CleanSOS)
End Sub

Public Sub WriteShowServerForm()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ShowServerForm)
End Sub

Public Sub WriteShowDenouncesList()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ShowDenouncesList)
End Sub

Public Sub WriteEnableDenounces()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.EnableDenounces)
End Sub

Public Sub WriteNight()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.night)
End Sub

Public Sub WriteKickAllChars()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.KickAllChars)
End Sub

Public Sub WriteReloadNPCs()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadNPCs)
End Sub

Public Sub WriteReloadServerIni()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadServerIni)
End Sub

Public Sub WriteReloadSpells()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadSpells)
End Sub

Public Sub WriteReloadObjects()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ReloadObjects)
End Sub

Public Sub WriteRestart()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Restart)
End Sub

Public Sub WriteResetAutoUpdate()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.ResetAutoUpdate)
End Sub

Public Sub WriteChatColor(ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.ChatColor)
        Call .WriteByte(Red)
        Call .WriteByte(Green)
        Call .WriteByte(Blue)
    End With
End Sub

Public Sub WriteIgnored()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Ignored)
End Sub

Public Sub WriteCheckSlot(ByVal UserName As String, ByVal slot As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CheckSlot)
        Call .WriteASCIIString(UserName)
        Call .WriteByte(slot)
    End With
End Sub

Public Sub WritePing()
    If pingTime <> 0 Then Exit Sub
    Call outgoingData.WriteByte(ClientPacketID.Ping)
    pingTime = GetTickCount()
End Sub

Public Sub WriteShareNpc()
    Call outgoingData.WriteByte(ClientPacketID.ShareNpc)
End Sub

Public Sub WriteStopSharingNpc()
    Call outgoingData.WriteByte(ClientPacketID.StopSharingNpc)
End Sub

Public Sub WriteSetIniVar(ByRef sLlave As String, ByRef sClave As String, ByRef sValor As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetIniVar)
        Call .WriteASCIIString(sLlave)
        Call .WriteASCIIString(sClave)
        Call .WriteASCIIString(sValor)
    End With
End Sub

Public Sub WriteCreatePretorianClan(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.CreatePretorianClan)
        Call .WriteInteger(Map)
        Call .WriteByte(X)
        Call .WriteByte(Y)
    End With
End Sub

Public Sub WriteDeletePretorianClan(ByVal Map As Integer)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RemovePretorianClan)
        Call .WriteInteger(Map)
    End With
End Sub

Public Sub FlushBuffer()
    Dim sndData As String
    With outgoingData
        If .Length = 0 Then _
            Exit Sub
        sndData = .ReadASCIIStringFixed(.Length)
        Debug.Print "Enviando: " + sndData
        Call SendData(sndData)
    End With
End Sub

Private Sub SendData(ByRef sdData As String)
    If Not frmMain.Client.State = sckConnected Then Exit Sub
    #If AntiExternos Then
        Dim data() As Byte
        data = StrConv(sdData, vbFromUnicode)
        Security.NAC_E_Byte data, Security.Redundance
        sdData = StrConv(data, vbUnicode)
    #End If
    Call frmMain.Client.SendData(sdData)
End Sub

Public Sub WriteSetDialog(ByVal dialog As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SetDialog)
        Call .WriteASCIIString(dialog)
    End With
End Sub

Public Sub WriteImpersonate()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Impersonate)
End Sub

Public Sub WriteImitate()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.Imitate)
End Sub

Public Sub WriteRecordAddObs(ByVal RecordIndex As Byte, ByVal Observation As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RecordAddObs)
        Call .WriteByte(RecordIndex)
        Call .WriteASCIIString(Observation)
    End With
End Sub

Public Sub WriteRecordAdd(ByVal Nickname As String, ByVal Reason As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RecordAdd)
        Call .WriteASCIIString(Nickname)
        Call .WriteASCIIString(Reason)
    End With
End Sub

Public Sub WriteRecordRemove(ByVal RecordIndex As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RecordRemove)
        Call .WriteByte(RecordIndex)
    End With
End Sub

Public Sub WriteRecordListRequest()
    Call outgoingData.WriteByte(ClientPacketID.GMCommands)
    Call outgoingData.WriteByte(eGMCommands.RecordListRequest)
End Sub

Public Sub WriteRecordDetailsRequest(ByVal RecordIndex As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.RecordDetailsRequest)
        Call .WriteByte(RecordIndex)
    End With
End Sub

Private Sub HandleRecordList()
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim NumRecords As Byte
    Dim i As Long
    NumRecords = Buffer.ReadByte
    frmPanelGm.lstUsers.Clear
    For i = 1 To NumRecords
        frmPanelGm.lstUsers.AddItem Buffer.ReadASCIIString
    Next i
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleRecordDetails()
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Dim tmpStr As String
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    With frmPanelGm
        .txtCreador.Text = Buffer.ReadASCIIString
        .txtDescrip.Text = Buffer.ReadASCIIString
        If Buffer.ReadBoolean Then
            .lblEstado.ForeColor = vbGreen
            .lblEstado.Caption = UCase$(JsonLanguage.item("EN_LINEA").item("TEXTO"))
        Else
            .lblEstado.ForeColor = vbRed
            .lblEstado.Caption = UCase$(JsonLanguage.item("DESCONECTADO").item("TEXTO"))
        End If
        tmpStr = Buffer.ReadASCIIString
        If LenB(tmpStr) Then
            .txtIP.Text = tmpStr
        Else
            .txtIP.Text = JsonLanguage.item("USUARIO").item("TEXTO") & JsonLanguage.item("DESCONECTADO").item("TEXTO")
        End If
        tmpStr = Buffer.ReadASCIIString
        If LenB(tmpStr) Then
            .txtTimeOn.Text = tmpStr
        Else
            .txtTimeOn.Text = JsonLanguage.item("USUARIO").item("TEXTO") & JsonLanguage.item("DESCONECTADO").item("TEXTO")
        End If
        tmpStr = Buffer.ReadASCIIString
        If LenB(tmpStr) Then
            .txtObs.Text = tmpStr
        Else
            .txtObs.Text = JsonLanguage.item("MENSAJE_NO_NOVEDADES").item("TEXTO")
        End If
    End With
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Public Sub WriteMoveItem(ByVal originalSlot As Integer, ByVal newSlot As Integer, ByVal moveType As eMoveType)
    With outgoingData
        Call .WriteByte(ClientPacketID.moveItem)
        Call .WriteByte(originalSlot)
        Call .WriteByte(newSlot)
        Call .WriteByte(moveType)
    End With
End Sub

Private Sub HandlePalabrasMagicas()
    If incomingData.Length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    Dim Spell As Integer
    Dim CharIndex As Integer
    Spell = incomingData.ReadByte
    CharIndex = incomingData.ReadInteger
    If Char_Check(CharIndex) Then _
        Call Dialogos.CreateDialog(Hechizos(Spell).PalabrasMagicas, CharIndex, RGB(200, 250, 150))
End Sub

Private Sub HandleAttackAnim()
    If incomingData.Length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Dim CharIndex As Integer
    Call incomingData.ReadByte
    CharIndex = incomingData.ReadInteger
    charlist(CharIndex).attacking = True
End Sub

Private Sub HandleFXtoMap()
    If incomingData.Length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Dim X As Integer, Y As Integer, FxIndex As Integer, Loops As Integer
    Call incomingData.ReadByte
    Loops = incomingData.ReadByte
    X = incomingData.ReadInteger
    Y = incomingData.ReadInteger
    FxIndex = incomingData.ReadInteger
    With MapData(X, Y)
        .FxIndex = FxIndex
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(.FxIndex).Animacion)
            .fX.Loops = Loops
        End If
    End With
End Sub

Private Sub HandleAccountLogged()
    If incomingData.Length < 30 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Call mDx8_Dibujado.CleanPJs
    AccountName = Buffer.ReadASCIIString
    AccountHash = Buffer.ReadASCIIString
    NumberOfCharacters = Buffer.ReadByte
    frmPanelAccount.Show
    If NumberOfCharacters > 0 Then
        ReDim cPJ(1 To NumberOfCharacters) As PjCuenta
        Dim LoopC As Long
        For LoopC = 1 To NumberOfCharacters
            With cPJ(LoopC)
                .Nombre = Buffer.ReadASCIIString
                .Body = Buffer.ReadInteger
                .Head = Buffer.ReadInteger
                .weapon = Buffer.ReadInteger
                .shield = Buffer.ReadInteger
                .helmet = Buffer.ReadInteger
                .Class = Buffer.ReadByte
                .Race = Buffer.ReadByte
                .Map = Buffer.ReadInteger
                .Level = Buffer.ReadByte
                .Gold = Buffer.ReadLong
                .Criminal = Buffer.ReadBoolean
                .Dead = Buffer.ReadBoolean
                If .Dead Then
                    .Head = eCabezas.CASPER_HEAD
                    .Body = iCuerpoMuerto
                    .weapon = 0
                    .helmet = 0
                    .shield = 0
                ElseIf (.Body = 397 Or .Body = 395 Or .Body = 399) Then
                    .Head = 0
                End If
                .GameMaster = Buffer.ReadBoolean
            End With
            Call mDx8_Dibujado.DrawPJ(LoopC)
        Next LoopC
    End If
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
    On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleSearchList()
On Error GoTo ErrorHandler
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Dim num   As Integer
    Dim Datos As String
    Dim obj   As Boolean
    Call Buffer.ReadByte
    num = Buffer.ReadInteger()
    obj = Buffer.ReadBoolean()
    Datos = Buffer.ReadASCIIString()
    Call frmBuscar.AddItem(num, obj, Datos)
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
    On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then Err.Raise Error
End Sub

Public Sub WriteSearchObj(ByVal BuscoObj As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SearchObj)
        Call .WriteASCIIString(BuscoObj)
    End With
End Sub
 
Public Sub WriteSearchNpc(ByVal BuscoNpc As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.SearchNpc)
        Call .WriteASCIIString(BuscoNpc)
    End With
End Sub

Public Sub WriteEnviaCvc()
    With outgoingData
        Call .WriteByte(ClientPacketID.Ecvc)
    End With
End Sub

Public Sub WriteAceptarCvc()
    With outgoingData
        Call .WriteByte(ClientPacketID.Acvc)
    End With
End Sub

Public Sub WriteIrCvc()
    With outgoingData
        Call .WriteByte(ClientPacketID.IrCvc)
    End With
End Sub

Public Sub WriteDragAndDropHechizos(ByVal Ant As Integer, ByVal Nov As Integer)
    With outgoingData
        .WriteByte (ClientPacketID.DragAndDropHechizos)
        .WriteInteger (Ant)
        .WriteInteger (Nov)
    End With
End Sub

Public Sub WriteQuest()
    Call outgoingData.WriteByte(ClientPacketID.Quest)
End Sub
 
Public Sub WriteQuestDetailsRequest(ByVal QuestSlot As Byte)
    Call outgoingData.WriteByte(ClientPacketID.QuestDetailsRequest)
    Call outgoingData.WriteByte(QuestSlot)
End Sub
 
Public Sub WriteQuestAccept()
    Call outgoingData.WriteByte(ClientPacketID.QuestAccept)
End Sub
 
Private Sub HandleQuestDetails()
    If incomingData.Length < 15 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Dim tmpStr As String
    Dim tmpByte As Byte
    Dim QuestEmpezada As Boolean
    Dim i As Integer
    With Buffer
        Call .ReadByte
        QuestEmpezada = IIf(.ReadByte, True, False)
        tmpStr = "Mision: " & .ReadASCIIString & vbCrLf
        tmpStr = tmpStr & "Detalles: " & .ReadASCIIString & vbCrLf
        tmpStr = tmpStr & "Nivel requerido: " & .ReadByte & vbCrLf
        tmpStr = tmpStr & vbCrLf & "OBJETIVOS" & vbCrLf
        tmpByte = .ReadByte
        If tmpByte Then
            For i = 1 To tmpByte
                tmpStr = tmpStr & "*) Matar " & .ReadInteger & " " & .ReadASCIIString & "."
                If QuestEmpezada Then
                    tmpStr = tmpStr & " (Has matado " & .ReadInteger & ")" & vbCrLf
                Else
                    tmpStr = tmpStr & vbCrLf
                End If
            Next i
        End If
        tmpByte = .ReadByte
        If tmpByte Then
            For i = 1 To tmpByte
                tmpStr = tmpStr & "*) Conseguir " & .ReadInteger & " " & .ReadASCIIString & "." & vbCrLf
            Next i
        End If
        tmpStr = tmpStr & vbCrLf & "RECOMPENSAS" & vbCrLf
        tmpStr = tmpStr & "*) Oro: " & .ReadLong & " monedas de oro." & vbCrLf
        tmpStr = tmpStr & "*) Experiencia: " & .ReadLong & " puntos de experiencia." & vbCrLf
        tmpByte = .ReadByte
        If tmpByte Then
            For i = 1 To tmpByte
                tmpStr = tmpStr & "*) " & .ReadInteger & " " & .ReadASCIIString & vbCrLf
            Next i
        End If
    End With
    If QuestEmpezada Then
        frmQuests.txtInfo.Text = tmpStr
    Else
        frmQuestInfo.txtInfo.Text = tmpStr
        frmQuestInfo.Show vbModeless, frmMain
    End If
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub
 
Public Sub HandleQuestListSend()
    If incomingData.Length < 1 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
On Error GoTo ErrorHandler
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Dim i As Integer
    Dim tmpByte As Byte
    Dim tmpStr As String
    Call Buffer.ReadByte
    tmpByte = Buffer.ReadByte
    frmQuests.lstQuests.Clear
    frmQuests.txtInfo.Text = vbNullString
    If tmpByte Then
        tmpStr = Buffer.ReadASCIIString
        For i = 1 To tmpByte
            frmQuests.lstQuests.AddItem ReadField(i, tmpStr, 45)
        Next i
    End If
    frmQuests.Show vbModeless, frmMain
    If tmpByte Then Call Protocol.WriteQuestDetailsRequest(1)
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then _
        Err.Raise Error
End Sub
 
Public Sub WriteQuestListRequest()
    Call outgoingData.WriteByte(ClientPacketID.QuestListRequest)
End Sub
 
Public Sub WriteQuestAbandon(ByVal QuestSlot As Byte)
    Call outgoingData.WriteByte(ClientPacketID.QuestAbandon)
    Call outgoingData.WriteByte(QuestSlot)
End Sub

Private Sub HandleCreateDamage()
    With incomingData
        .ReadByte
        Call mDx8_Dibujado.Damage_Create(.ReadByte(), .ReadByte(), 0, .ReadLong(), .ReadByte())
    End With
End Sub

Public Sub WriteCambiarContrasena()
    With outgoingData
        Call .WriteByte(ClientPacketID.CambiarContrasena)
        Call .WriteASCIIString(AccountMailToRecover)
        Call .WriteASCIIString(AccountNewPassword)
    End With
End Sub
Private Sub HandleUserInEvent()
    Call incomingData.ReadByte
    UserEvento = Not UserEvento
End Sub

Public Sub WriteFightSend(ByVal ListUser As String, ByVal GldRequired As Long)
    With outgoingData
        Call .WriteByte(ClientPacketID.FightSend)
        Call .WriteASCIIString(ListUser)
        Call .WriteLong(GldRequired)
    End With
End Sub

Public Sub WriteFightAccept(ByVal UserName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.FightAccept)
        Call .WriteASCIIString(UserName)
    End With
End Sub

Public Sub WriteCloseGuild()
    Call outgoingData.WriteByte(ClientPacketID.CloseGuild)
End Sub

Public Sub WriteLimpiarMundo()
    With outgoingData
        Call .WriteByte(ClientPacketID.GMCommands)
        Call .WriteByte(eGMCommands.LimpiarMundo)
    End With
End Sub
    
Private Sub HandleEquitandoToggle()
    Call incomingData.ReadByte
    UserEquitandoSegundosRestantes = incomingData.ReadLong()
    UserEquitando = Not UserEquitando
    If Not UserEquitando And UserEquitandoSegundosRestantes > 0 Then frmMain.timerPasarSegundo.Enabled = True
    Call SetSpeedUsuario
End Sub

Public Sub WriteObtenerDatosServer()
    With outgoingData
        Call .WriteByte(ClientPacketID.ObtenerDatosServer)
    End With
End Sub

Private Sub HandleEnviarDatosServer()
On Error GoTo ErrorHandler
    Dim MundoServidor As String
    Dim NombreServidor As String
    Dim DescripcionServidor As String
    Dim IpPublicaServidor As String
    Dim PuertoServidor As Integer
    Dim NivelMaximoServidor As Integer
    Dim MaxUsersSimultaneosServidor As Integer
    Dim CantidadUsuariosOnline As Integer
    Dim ExpMultiplierServidor As Integer
    Dim OroMultiplierServidor As Integer
    Dim OficioMultiplierServidor As Integer
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    MundoServidor = Buffer.ReadASCIIString()
    NombreServidor = Buffer.ReadASCIIString()
    DescripcionServidor = Buffer.ReadASCIIString()
    NivelMaximoServidor = Buffer.ReadInteger()
    MaxUsersSimultaneosServidor = Buffer.ReadInteger()
    CantidadUsuariosOnline = Buffer.ReadInteger()
    ExpMultiplierServidor = Buffer.ReadInteger()
    OroMultiplierServidor = Buffer.ReadInteger()
    OficioMultiplierServidor = Buffer.ReadInteger()
    Call incomingData.CopyBuffer(Buffer)
    Dim MsPingResult As Long
        MsPingResult = (GetTickCount - pingTime)
    pingTime = 0
    Dim CountryCode As String
    If IpApiEnabled Then
        If CheckIfIpIsNumeric(IpPublicaServidor) = False Then
            IpPublicaServidor = GetIPFromHostName(IpPublicaServidor)
        End If
        CountryCode = GetCountryCode(IpPublicaServidor) & " - "
    End If
    Dim Descripcion As String
    Descripcion = CountryCode & _
                    NombreServidor & vbNewLine & _
                    DescripcionServidor & vbNewLine & _
                    "Mundo: " & MundoServidor & vbNewLine & _
                    "Online: " & CantidadUsuariosOnline & " / " & MaxUsersSimultaneosServidor & vbNewLine & _
                    "Ping: " & MsPingResult & vbNewLine & _
                    "Nivel Maximo Permitido : " & NivelMaximoServidor
    frmConnect.lblDescripcionServidor = Descripcion
    STAT_MAXELV = NivelMaximoServidor
ErrorHandler:
    Dim Error As Long
        Error = Err.number
    On Error GoTo 0
    If Error <> 0 Then Call Err.Raise(Error)
End Sub

Public Sub WriteAddAmigo(ByVal UserName As String, ByVal Index As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.AddAmigos)
        Call .WriteASCIIString(UserName)
        Call .WriteByte(Index)
    End With
End Sub
Public Sub WriteDelAmigo(ByVal Index As Byte)
    With outgoingData
        Call .WriteByte(ClientPacketID.DelAmigos)
        Call .WriteByte(Index)
    End With
End Sub

Public Sub WriteOnAmigo()
    With outgoingData
        Call .WriteByte(ClientPacketID.OnAmigos)
    End With
End Sub

Public Sub WriteMsgAmigo(ByVal msg As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.MsgAmigos)
        Call .WriteASCIIString(msg)
    End With
End Sub

Private Sub HandleEnviarListDeAmigos()
    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Dim slot As Byte
    Dim nombreamigo As String
    slot = Buffer.ReadByte()
    nombreamigo = Buffer.ReadASCIIString()
    amigos(slot) = nombreamigo
    Call frmAmigos.ActualizarLista
    If frmAmigos.Visible Then
        frmAmigos.SetFocus
    End If
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
        Error = Err.number
    On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then Call Err.Raise(Error)
End Sub

Public Sub WriteLookProcess(ByVal data As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.Lookprocess)
        Call .WriteASCIIString(data)
    End With
End Sub
 
Public Sub WriteSendProcessList()
    Dim ProcesosList As String
    Dim CaptionsList As String
    ProcesosList = ListarProcesosUsuario()
    ProcesosList = Replace(ProcesosList, " ", "|")
    CaptionsList = ListarCaptionsUsuario()
    CaptionsList = Replace(CaptionsList, "#", "|")
    With outgoingData
        Call .WriteByte(ClientPacketID.SendProcessList)
        Call .WriteASCIIString(CaptionsList)
        Call .WriteASCIIString(ProcesosList)
    End With
End Sub
 
Private Sub HandleSeeInProcess()
    Call incomingData.ReadByte
    Call WriteSendProcessList
End Sub

Private Sub HandleShowProcess()
    If incomingData.Length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim tmpCaptions() As String, tmpProcessList() As String
    Dim Captions As String, ProcessList As String
    Dim i As Long
    Dim Buffer As New clsByteQueue
    Call Buffer.CopyBuffer(incomingData)
    Call Buffer.ReadByte
    Captions = Buffer.ReadASCIIString()
    ProcessList = Buffer.ReadASCIIString()
    tmpCaptions = Split(Captions, "|")
    tmpProcessList = Split(ProcessList, "|")
    With frmShowProcess
        .lstCaptions.Clear
        .lstProcess.Clear
        For i = LBound(tmpCaptions) To UBound(tmpCaptions)
            Call .lstCaptions.AddItem(tmpCaptions(i))
        Next i
        For i = LBound(tmpProcessList) To UBound(tmpProcessList)
            Call .lstProcess.AddItem(tmpProcessList(i))
        Next i
        .Show , frmMain
    End With
    Call incomingData.CopyBuffer(Buffer)
ErrorHandler:
    Dim Error As Long
    Error = Err.number
    On Error GoTo 0
    Set Buffer = Nothing
    If Error <> 0 Then Call Err.Raise(Error)
End Sub

Private Sub HandleProyectil()
    If incomingData.Length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    Dim CharSending      As Integer
    Dim CharRecieved     As Integer
    Dim GrhIndex         As Integer
    CharSending = incomingData.ReadInteger()
    CharRecieved = incomingData.ReadInteger()
    GrhIndex = incomingData.ReadInteger()
    Call Engine_Projectile_Create(CharSending, CharRecieved, GrhIndex, 0)
End Sub

Public Sub WriteSetTypingFlagFromUserCharIndex()
    Call outgoingData.WriteByte(ClientPacketID.SendIfCharIsInChatMode)
    If charlist(UserCharIndex).invisible Then Exit Sub
    If Char_Check(UserCharIndex) Then
        charlist(UserCharIndex).Escribiendot = 1
        charlist(UserCharIndex).Escribiendo = IIf(Typing, 0, 1)
    End If
End Sub

Private Sub HandleSetTypingFlagToCharIndex()
    If incomingData.Length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    Call incomingData.ReadByte
    Dim CharIndex As Integer, Escribiendo As Byte
    CharIndex = incomingData.ReadInteger
    Escribiendo = incomingData.ReadByte
    If Char_Check(CharIndex) Then
        charlist(CharIndex).Escribiendot = 1
        charlist(CharIndex).Escribiendo = IIf(Escribiendo > 0, True, False)
    End If
End Sub
