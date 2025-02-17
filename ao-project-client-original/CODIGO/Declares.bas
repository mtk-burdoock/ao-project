Attribute VB_Name = "Mod_Declaraciones"
Option Explicit

Public IntervaloDopas As Long
Public IntervaloInvi As Long
Public TiempoInvi As Long
Public TiempoDopas As Long
Public Const MAX_AMIGOS As Byte = 50
Public Const MAX_CHARACTERS As Byte = 10
Public amigos(1 To MAX_AMIGOS) As String
Public Typing As Boolean
Public lastTickEscribiendo As Long
Public Inet As clsInet
Public AccountMailToRecover As String
Public AccountNewPassword As String
Public ColorTecho As Byte
Public temp_rgb(3) As Long
Public renderText As String
Public renderFont As Integer
Public colorRender As Byte
Public render_msg(3) As Long
Public Sonidos As clsSoundMapas
Public Movement_Speed As Single
Public DialogosClanes As clsGuildDlg
Public Dialogos As clsDialogs
Public Audio As clsAudio
Public Inventario As clsGraphicalInventory
Public InvBanco(1) As clsGraphicalInventory
Public InvComUsu As clsGraphicalInventory
Public InvOroComUsu(2) As clsGraphicalInventory
Public InvOfferComUsu(1) As clsGraphicalInventory
Public InvComNpc As clsGraphicalInventory
Public Const MAX_LIST_ITEMS As Byte = 4
Public InvLingosHerreria(1 To MAX_LIST_ITEMS) As clsGraphicalInventory
Public InvMaderasCarpinteria(1 To MAX_LIST_ITEMS) As clsGraphicalInventory
Public InvObjArtesano(1 To MAX_LIST_ITEMS) As clsGraphicalInventory
Public Const MAX_ITEMS_CRAFTEO As Byte = 4
Public CustomKeys As clsCustomKeys
Public CustomMessages As clsCustomMessages
Public incomingData As clsByteQueue
Public outgoingData As clsByteQueue
Public MainTimer As clsTimer

Public Enum eSockError
   TOO_FAST = 24036
   REFUSED = 24061
   TIME_OUT = 24060
End Enum

Public Const SND_CLICK As String = "click.Wav"
Public Const SND_PASOS1 As String = "23.Wav"
Public Const SND_PASOS2 As String = "24.Wav"
Public Const SND_NAVEGANDO As String = "50.wav"
Public Const SND_DICE As String = "cupdice.Wav"

Public Enum eIntervalos
    INT_MACRO_HECHIS = 2000
    INT_MACRO_TRABAJO = 900
    INT_ATTACK = 1500
    INT_ARROWS = 1400
    INT_CAST_SPELL = 1400
    INT_CAST_ATTACK = 1000
    INT_WORK = 700
    INT_USEITEMU = 450
    INT_USEITEMDCK = 125
    INT_SENTRPU = 2000
    INT_CHANGE_HEADING = 300
End Enum

Public MacroBltIndex As Integer
Public Const NUMATRIBUTES As Byte = 5
Public Const iCuerpoMuerto As Integer = 8

Public Enum eCabezas
    CASPER_HEAD = 500
    FRAGATA_FANTASMAL = 87
    HUMANO_H_PRIMER_CABEZA = 1
    HUMANO_H_ULTIMA_CABEZA = 40
    HUMANO_H_CUERPO_DESNUDO = 21
    ELFO_H_PRIMER_CABEZA = 101
    ELFO_H_ULTIMA_CABEZA = 122
    ELFO_H_CUERPO_DESNUDO = 210
    DROW_H_PRIMER_CABEZA = 201
    DROW_H_ULTIMA_CABEZA = 221
    DROW_H_CUERPO_DESNUDO = 32
    ENANO_H_PRIMER_CABEZA = 301
    ENANO_H_ULTIMA_CABEZA = 319
    ENANO_H_CUERPO_DESNUDO = 53
    GNOMO_H_PRIMER_CABEZA = 401
    GNOMO_H_ULTIMA_CABEZA = 416
    GNOMO_H_CUERPO_DESNUDO = 222
    HUMANO_M_PRIMER_CABEZA = 70
    HUMANO_M_ULTIMA_CABEZA = 89
    HUMANO_M_CUERPO_DESNUDO = 39
    ELFO_M_PRIMER_CABEZA = 170
    ELFO_M_ULTIMA_CABEZA = 188
    ELFO_M_CUERPO_DESNUDO = 259
    DROW_M_PRIMER_CABEZA = 270
    DROW_M_ULTIMA_CABEZA = 288
    DROW_M_CUERPO_DESNUDO = 40
    ENANO_M_PRIMER_CABEZA = 370
    ENANO_M_ULTIMA_CABEZA = 384
    ENANO_M_CUERPO_DESNUDO = 60
    GNOMO_M_PRIMER_CABEZA = 470
    GNOMO_M_ULTIMA_CABEZA = 484
    GNOMO_M_CUERPO_DESNUDO = 260
End Enum

Public ColoresPJ(0 To 50) As Long
Public ColoresDano(51 To 56) As Long

Public Type tServerInfo
    Ip As String
    Puerto As Integer
    Desc As String
    Mundo As String
End Type

Public ServersLst() As tServerInfo
Public CurServer As Integer
Public CreandoClan As Boolean
Public ClanName As String
Public Site As String
Public UserCiego As Boolean
Public UserEstupido As Boolean
Public NoRes As Boolean
Public RainBufferIndex As Long
Public FogataBufferIndex As Long

Public Enum ePartesCuerpo
    bCabeza = 1
    bPiernaIzquierda = 2
    bPiernaDerecha = 3
    bBrazoDerecho = 4
    bBrazoIzquierdo = 5
    bTorso = 6
End Enum

Public NumEscudosAnims As Integer
Public ArmasHerrero() As tItemsConstruibles
Public ArmadurasHerrero() As tItemsConstruibles
Public ObjCarpintero() As tItemsConstruibles
Public CarpinteroMejorar() As tItemsConstruibles
Public HerreroMejorar() As tItemsConstruibles
Public ObjArtesano() As tItemArtesano
Public UsaMacro As Boolean
Public CnTd As Byte
Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory
Public TradingUserName As String
Public Tips() As String * 255

Public Enum E_Heading
    nada = 0
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

Public Const MAX_NORMAL_INVENTORY_SLOTS As Byte = 25
Public Const MAX_MOCHILA_CHICA_INVENTORY_SLOTS As Byte = 30
Public Const MAX_INVENTORY_SLOTS        As Byte = 35
Public Const MAX_INVENTORY_OBJS As Integer = 10000
Public Const MAX_NPC_INVENTORY_SLOTS As Byte = 50
Public Const MAXHECHI As Byte = 35
Public Const INV_OFFER_SLOTS As Byte = 20
Public Const INV_GOLD_SLOTS As Byte = 1
Public Const MAXSKILLPOINTS As Byte = 100
Public Const MAXATRIBUTOS As Byte = 40
Public Const FLAGORO As Integer = MAX_INVENTORY_SLOTS + 1
Public Const GOLD_OFFER_SLOT As Integer = INV_OFFER_SLOTS + 1

Public Enum eClass
    Mage = 1
    Cleric = 2
    Warrior = 3
    Assasin = 4
    Thief = 5
    Bard = 6
    Druid = 7
    Bandit = 8
    Paladin = 9
    Hunter = 10
    Worker = 11
    Pirate = 12
End Enum

Public Enum eCiudad
    cUllathorpe = 1
    cNix = 2
    cBanderbill = 3
    cLindos = 4
    cArghal = 5
End Enum

Enum eRaza
    Humano = 1
    Elfo = 2
    ElfoOscuro = 3
    Gnomo = 4
    Enano = 5
End Enum

Public Enum eSkill
    Magia = 1
    Robar = 2
    Tacticas = 3
    Armas = 4
    Meditar = 5
    Apunalar = 6
    Ocultarse = 7
    Supervivencia = 8
    Talar = 9
    Comerciar = 10
    Defensa = 11
    Pesca = 12
    Mineria = 13
    Carpinteria = 14
    Herreria = 15
    Liderazgo = 16
    Domar = 17
    Proyectiles = 18
    Wrestling = 19
    Navegacion = 20
End Enum

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5
End Enum

Enum eGenero
    Hombre = 1
    Mujer = 2
End Enum

Public Enum PlayerType
    User = &H1
    Consejero = &H2
    SemiDios = &H4
    Dios = &H8
    Admin = &H10
    RoleMaster = &H20
    ChaosCouncil = &H40
    RoyalCouncil = &H80
End Enum

Public Enum eObjType
    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otOro = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otForos = 10
    otPociones = 11
    otLibros = 12
    otBebidas = 13
    otLena = 14
    otFogata = 15
    otescudo = 16
    otcasco = 17
    otAnillo = 18
    otTeleport = 19
    otMuebles = 20
    otJoyas = 21
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otMonturas = 25
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otGemas = 29
    otFlores = 30
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otManuales = 35
    otArbolElfico = 36
    otMochilas = 37
    otYacimientoPez = 38
    otCualquiera = 1000
End Enum

Public MaxInventorySlots As Byte
Public Const GRH_SLOT_INVENTARIO_NEGRO As Integer = 26095
Public Const GRH_SLOT_INVENTARIO_ROJO As Integer = 26096
Public Const GRH_SLOT_INVENTARIO_VIOLETA As Integer = 6834
Public Const GRH_SLOT_INVENTARIO_DORADO As Integer = 6840
Public Const FundirMetal As Integer = 88

Public Enum eNickColor
    ieCriminal = &H1
    ieCiudadano = &H2
    ieAtacable = &H4
End Enum

Public Enum eGMCommands
    GMMessage = 1
    showName
    OnlineRoyalArmy
    OnlineChaosLegion
    GoNearby
    comment
    serverTime
    Where
    CreaturesInMap
    WarpMeToTarget
    WarpChar
    Silence
    SOSShowList
    SOSRemove
    GoToChar
    invisible
    GMPanel
    RequestUserList
    Working
    Hiding
    Jail
    KillNPC
    WarnUser
    EditChar
    RequestCharInfo
    RequestCharStats
    RequestCharGold
    RequestCharInventory
    RequestCharBank
    RequestCharSkills
    ReviveChar
    OnlineGM
    OnlineMap
    Forgive
    Kick
    Execute
    BanChar
    UnbanChar
    NPCFollow
    SummonChar
    SpawnListRequest
    SpawnCreature
    ResetNPCInventory
    ServerMessage
    NickToIP
    IPToNick
    GuildOnlineMembers
    TeleportCreate
    TeleportDestroy
    RainToggle
    SetCharDescription
    ForceMP3ToMap
    ForceMIDIToMap
    ForceWAVEToMap
    RoyalArmyMessage
    ChaosLegionMessage
    CitizenMessage
    CriminalMessage
    TalkAsNPC
    DestroyAllItemsInArea
    AcceptRoyalCouncilMember
    AcceptChaosCouncilMember
    ItemsInTheFloor
    MakeDumb
    MakeDumbNoMore
    DumpIPTables
    CouncilKick
    SetTrigger
    AskTrigger
    BannedIPList
    BannedIPReload
    GuildMemberList
    GuildBan
    BanIP
    UnbanIP
    CreateItem
    DestroyItems
    ChaosLegionKick
    RoyalArmyKick
    ForceMP3All
    ForceMIDIAll
    ForceWAVEAll
    RemovePunishment
    TileBlockedToggle
    KillNPCNoRespawn
    KillAllNearbyNPCs
    LastIP
    ChangeMOTD
    SetMOTD
    SystemMessage
    CreateNPC
    ImperialArmour
    ChaosArmour
    NavigateToggle
    ServerOpenToUsersToggle
    TurnOffServer
    TurnCriminal
    ResetFactions
    RemoveCharFromGuild
    RequestCharMail
    AlterPassword
    AlterMail
    AlterName
    DoBackUp
    ShowGuildMessages
    SaveMap
    ChangeMapInfoPK
    ChangeMapInfoBackup
    ChangeMapInfoRestricted
    ChangeMapInfoNoMagic
    ChangeMapInfoNoInvi
    ChangeMapInfoNoResu
    ChangeMapInfoLand
    ChangeMapInfoZone
    ChangeMapInfoStealNpc
    ChangeMapInfoNoOcultar
    ChangeMapInfoNoInvocar
    SaveChars
    CleanSOS
    ShowServerForm
    night
    KickAllChars
    ReloadNPCs
    ReloadServerIni
    ReloadSpells
    ReloadObjects
    Restart
    ResetAutoUpdate
    ChatColor
    Ignored
    CheckSlot
    SetIniVar
    CreatePretorianClan
    RemovePretorianClan
    EnableDenounces
    ShowDenouncesList
    MapMessage
    SetDialog
    Impersonate
    Imitate
    RecordAdd
    RecordRemove
    RecordAddObs
    RecordListRequest
    RecordDetailsRequest
    ExitDestroy
    ToggleCentinelActivated
    SearchNpc
    SearchObj
    LimpiarMundo
End Enum

Public Const MENSAJE_2 As String = "!!"
Public Const MENSAJE_22 As String = "!"

Public Enum eMessages
    NPCSwing
    NPCKillUser
    BlockedWithShieldUser
    BlockedWithShieldOther
    UserSwing
    SafeModeOn
    SafeModeOff
    ResuscitationSafeOff
    ResuscitationSafeOn
    NobilityLost
    CantUseWhileMeditating
    NPCHitUser
    UserHitNPC
    UserAttackedSwing
    UserHittedByUser
    UserHittedUser
    WorkRequestTarget
    HaveKilledUser
    UserKill
    NPCKill
    EarnExp
    GoHome
    CancelGoHome
    FinishHome
    UserMuerto
    NpcInmune
    Hechizo_HechiceroMSG_NOMBRE
    Hechizo_HechiceroMSG_ALGUIEN
    Hechizo_HechiceroMSG_CRIATURA
    Hechizo_PropioMSG
    Hechizo_TargetMSG
End Enum

Type Inventory
    ObjIndex As Integer
    Name As String
    GrhIndex As Long
    Amount As Long
    Equipped As Byte
    Valor As Single
    OBJType As Integer
    MaxDef As Integer
    MinDef As Integer
    MaxHit As Integer
    MinHit As Integer
    Incompatible As Boolean
End Type

Type NpCinV
    ObjIndex As Integer
    Name As String
    GrhIndex As Long
    Amount As Integer
    Valor As Single
    OBJType As Integer
    MaxDef As Integer
    MinDef As Integer
    MaxHit As Integer
    MinHit As Integer
    Incompatible As Boolean
    C1 As String
    C2 As String
    C3 As String
    C4 As String
    C5 As String
    C6 As String
    C7 As String
End Type

Type tReputacion
    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long
    Promedio As Long
End Type

Type tEstadisticasUsu
    CiudadanosMatados As Long
    CriminalesMatados As Long
    UsuariosMatados As Long
    NpcsMatados As Long
    Clase As String
    PenaCarcel As Long
End Type

Type tItemsConstruibles
    Name As String
    ObjIndex As Integer
    GrhIndex As Long
    LinH As Integer
    LinP As Integer
    LinO As Integer
    Madera As Integer
    MaderaElfica As Integer
    Upgrade As Integer
    UpgradeName As String
    UpgradeGrhIndex As Long
End Type

Type tItemCrafteo
    Name As String
    ObjIndex As Integer
    GrhIndex As Long
    Amount As Integer
End Type

Type tItemArtesano
    Name As String
    ObjIndex As Integer
    GrhIndex As Long
    ItemsCrafteo() As tItemCrafteo
End Type

Public Nombres As Boolean
Public UserHechizos(1 To MAXHECHI) As Integer

Public Type PjCuenta
    Nombre      As String
    Head        As Integer
    Body        As Integer
    shield      As Byte
    helmet      As Byte
    weapon      As Byte
    Mapa        As Integer
    Class       As Byte
    Race        As Byte
    Map         As Integer
    Level       As Byte
    Gold        As Long
    Criminal    As Boolean
    Dead        As Boolean
    GameMaster  As Boolean
End Type

Public cPJ() As PjCuenta

Public NPCInventory(1 To MAX_NPC_INVENTORY_SLOTS) As NpCinV
Public UserMeditar As Boolean
Public UserName As String
Public AccountName As String
Public AccountPassword As String
Public AccountHash As String
Public NumberOfCharacters As Byte
Public UserMaxHP As Integer
Public UserMinHP As Integer
Public UserMaxMAN As Integer
Public UserMinMAN As Integer
Public UserMaxSTA As Integer
Public UserMinSTA As Integer
Public UserMaxAGU As Byte
Public UserMinAGU As Byte
Public UserMaxHAM As Byte
Public UserMinHAM As Byte
Public UserGLD As Long
Public UserLvl As Integer
Public UserPort As Integer
Public UserEstado As Byte
Public UserPasarNivel As Long
Public UserExp As Long
Public UserReputacion As tReputacion
Public UserEstadisticas As tEstadisticasUsu
Public UserDescansar As Boolean
Public bShowTutorial As Boolean
Public FPSFLAG As Boolean
Public pausa As Boolean
Public UserParalizado As Boolean
Public UserInvisible As Boolean
Public UserNavegando As Boolean
Public UserEquitando As Boolean
Public UserEvento As Boolean
Public UserHogar As eCiudad
Public UserFuerza As Byte
Public UserAgilidad As Byte
Public UserWeaponEqpSlot As Byte
Public UserArmourEqpSlot As Byte
Public UserHelmEqpSlot As Byte
Public UserShieldEqpSlot As Byte
Public Comerciando As Boolean
Public MirandoForo As Boolean
Public MirandoAsignarSkills As Boolean
Public MirandoEstadisticas As Boolean
Public MirandoParty As Boolean
Public MirandoCarpinteria As Boolean
Public MirandoHerreria As Boolean
Public UserClase As eClass
Public UserSexo As eGenero
Public UserRaza As eRaza
Public UserEmail As String
Public Const NUMCIUDADES As Byte = 5
Public Const NUMSKILLS As Byte = 20
Public Const NUMATRIBUTOS As Byte = 5
Public Const NUMCLASES As Byte = 12
Public Const NUMRAZAS As Byte = 5
Public UserSkills(1 To NUMSKILLS) As Byte
Public PorcentajeSkills(1 To NUMSKILLS) As Byte
Public SkillsNames(1 To NUMSKILLS) As String
Public UserAtributos(1 To NUMATRIBUTOS) As Byte
Public AtributosNames(1 To NUMATRIBUTOS) As String
Public Ciudades(1 To NUMCIUDADES) As String
Public ListaRazas(1 To NUMRAZAS) As String
Public ListaClases(1 To NUMCLASES) As String
Public SkillPoints As Integer
Public Alocados As Integer
Public flags() As Integer
Public UsingSkill As Integer
Public pingTime As Long
Public EsPartyLeader As Boolean

Public Enum E_MODO
    Normal = 1
    CrearNuevoPj = 2
    Dados = 3
    CrearCuenta = 4
    CambiarContrasena = 5
    ObtenerDatosServer = 6
End Enum

Public EstadoLogin As E_MODO
   
Public Enum FxMeditar
    CHICO = 4
    MEDIANO = 5
    GRANDE = 6
    XGRANDE = 16
    XXGRANDE = 34
End Enum

Public Enum eClanType
    ct_RoyalArmy
    ct_Evil
    ct_Neutral
    ct_GM
    ct_Legal
    ct_Criminal
End Enum

Public Enum eEditOptions
    eo_Gold = 1
    eo_Experience = 2
    eo_Body = 3
    eo_Head = 4
    eo_CiticensKilled = 5
    eo_CriminalsKilled = 6
    eo_Level = 7
    eo_Class = 8
    eo_Skills = 9
    eo_SkillPointsLeft = 10
    eo_Nobleza = 11
    eo_Asesino = 12
    eo_Sex = 13
    eo_Raza = 14
    eo_addGold = 15
    eo_Vida = 16
    eo_Poss = 17
End Enum

Public Enum eTrigger
    nada = 0
    BAJOTECHO = 1
    CASA = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
End Enum

Public stxtbuffer As String
Public stxtbuffercmsg As String
Public Connected As Boolean
Public UserMap As Integer
Public nameMap As String
Public prgRun As Boolean
Public IPdelServidor As String
Public PuertoDelServidor As Integer
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GLC_HCURSOR = (-12)
Public hSwapCursor As Long
Public Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const SW_SHOWNORMAL As Long = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Type tIndiceCabeza
    Head(1 To 4) As Long
End Type

Public Type tIndiceCuerpo
    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Long
    OffsetX As Integer
    OffsetY As Integer
End Type

Public EsperandoLevel As Boolean

Public Enum eForumMsgType
    ieGeneral
    ieGENERAL_STICKY
    ieREAL
    ieREAL_STICKY
    ieCAOS
    ieCAOS_STICKY
End Enum

Public Enum eForumVisibility
    ieGENERAL_MEMBER = &H1
    ieREAL_MEMBER = &H2
    ieCAOS_MEMBER = &H4
End Enum

Public Enum eForumType
    ieGeneral
    ieREAL
    ieCAOS
End Enum

Public Const MAX_STICKY_POST As Byte = 5
Public Const MAX_GENERAL_POST As Byte = 30
Public Const STICKY_FORUM_OFFSET As Byte = 50

Public Type tForo
    StickyTitle(1 To MAX_STICKY_POST) As String
    StickyPost(1 To MAX_STICKY_POST) As String
    StickyAuthor(1 To MAX_STICKY_POST) As String
    GeneralTitle(1 To MAX_GENERAL_POST) As String
    GeneralPost(1 To MAX_GENERAL_POST) As String
    GeneralAuthor(1 To MAX_GENERAL_POST) As String
End Type

Public Foros(0 To 2) As tForo
Public clsForos As clsForum
Public FragShooterCapturePending As Boolean
Public FragShooterNickname As String
Public FragShooterKilledSomeone As Boolean
Public Traveling As Boolean
Public bShowGuildNews As Boolean
Public GuildNames() As String
Public GuildMembers() As String
Public Const OFFSET_HEAD As Integer = -34

Public Enum eSMType
    sResucitation
    sSafemode
    mSpells
    mWork
End Enum

Public Const SM_CANT As Byte = 4
Public SMStatus(SM_CANT) As Boolean
Public Const GRH_INI_SM As Long = 4978
Public Const ORO_INDEX As Long = 12
Public Const ORO_GRH As Long = 511
Public Const LH_GRH As Long = 724
Public Const LP_GRH As Long = 725
Public Const LO_GRH As Long = 723
Public Const MADERA_GRH As Long = 550
Public Const MADERA_ELFICA_GRH As Long = 1999
Public picMouseIcon As Picture

Public Enum eMoveType
    Inventory = 1
    Bank
End Enum

Public NumHechizos As Byte
Public Hechizos() As tHechizos
 
Public Type tHechizos
    Nombre As String
    Desc As String
    PalabrasMagicas As String
    ManaRequerida As Integer
    SkillRequerido As Byte
    EnergiaRequerida As Integer
    HechiceroMsg As String
    PropioMsg As String
    TargetMsg As String
End Type

Public MundoSeleccionado As String
Public Const uAOButton_bEsquina As String = "bEsquina.bmp"
Public Const uAOButton_bFondo As String = "bFondo.bmp"
Public Const uAOButton_bHorizontal As String = "bHorizontal.bmp"
Public Const uAOButton_bVertical As String = "bVertical.bmp"
Public Const uAOButton_cCheckbox As String = "cCheckbox.bmp"
Public Const uAOButton_cCheckboxSmall As String = "cCheckboxSmall.bmp"
Public JsonTips As Object
Public STAT_MAXELV As Byte
Public IntervaloParalizado As Integer
Public UserParalizadoSegundosRestantes As Integer
Public UserEquitandoSegundosRestantes As Long
Public QuantityServers As Integer
Public IpApiEnabled As Boolean

#If AntiExternos Then
    Public Security As New clsSecurity
#End If
