Attribute VB_Name = "Declaraciones"
Option Explicit

#If False Then
    Dim Map, X, Y, body, Clase, race, Email, obj, Length As Variant
#End If

Public Const DAMAGE_PUNAL    As Byte = 1
Public Const DAMAGE_NORMAL   As Byte = 2
Public Const DAMAGE_CRITICO  As Byte = 3
Public Const DAMAGE_FALLO    As Byte = 4
Public Const DAMAGE_CURAR    As Byte = 5
Public Const DAMAGE_TRABAJO  As Byte = 6
Public tickLimpieza          As Integer

Type CentinelaUser
    centinelaIndex     As Byte
    Codigo             As String
    CentinelaCheck     As Boolean
    Revisando          As Boolean
    UltimaRevision     As Long
End Type

Public aClon          As clsAntiMassClon
Public TrashCollector As Collection
Public Const MAXSPAWNATTEMPS = 60
Public Const INFINITE_LOOPS As Integer = -1
Public Const FXSANGRE = 14
Public Const MAXAMIGOS As Byte = 50
Public Const CHAT_COLOR_DEAD_CHAR As Long = &HC0C0C0
Public Const CHAT_COLOR_GM_YELL   As Long = &HF82FF
Public Const NO_3D_SOUND          As Byte = 0
Public Const iFragataFantasmal = 87
Public Const iFragataReal = 190
Public Const iFragataCaos = 189
Public Const iBarca = 84
Public Const iGalera = 85
Public Const iGaleon = 86
Public Const iBarcaCiuda = 395
Public Const iBarcaCiudaAtacable = 562
Public Const iGaleraCiuda = 397
Public Const iGaleraCiudaAtacable = 567
Public Const iGaleonCiuda = 399
Public Const iGaleonCiudaAtacable = 564
Public Const iBarcaReal = 560
Public Const iBarcaRealAtacable = 563
Public Const iGaleraReal = 565
Public Const iGaleraRealAtacable = 568
Public Const iGaleonReal = 572
Public Const iGaleonRealAtacable = 564
Public Const iBarcaPk = 396
Public Const iGaleraPk = 398
Public Const iGaleonPk = 400
Public Const iBarcaCaos = 561
Public Const iGaleraCaos = 566
Public Const iGaleonCaos = 573

Public Enum iMinerales
    HierroCrudo = 192
    PlataCruda = 193
    OroCrudo = 194
    LingoteDeHierro = 386
    LingoteDePlata = 387
    LingoteDeOro = 388
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

Public Enum ePrivileges
    Admin = 1
    Dios
    Especial
    SemiDios
    Consejero
    RoleMaster
End Enum

Public Enum eClass
    Mage = 1
    Cleric
    Warrior
    Assasin
    Thief
    Bard
    Druid
    Bandit
    Paladin
    Hunter
    Worker
    Pirat
End Enum

Public Enum eCiudad
    cUllathorpe = 1
    cNix = 2
    cBanderbill = 3
    cLindos = 4
    cArghal = 5
    cArkhein = 6
    cGotland = 7
    cPerdida = 8
    cTotem = 9
    cLastCity = 10
End Enum

Public Enum eRaza
    Humano = 1
    Elfo
    Drow
    Gnomo
    Enano
End Enum

Enum eGenero
    Hombre = 1
    Mujer
End Enum

Public Enum eClanType
    ct_RoyalArmy
    ct_Evil
    ct_Neutral
    ct_GM
    ct_Legal
    ct_Criminal
End Enum

Public Const LimiteNewbie As Byte = 12

Public Type tCabecera
    Desc As String * 255
    crc As Long
    MagicWord As Long
End Type

Public MiCabecera                    As tCabecera
Public Const TIEMPO_INICIOMEDITAR    As Integer = 2000
Public Const NingunEscudo            As Integer = 2
Public Const NingunCasco             As Integer = 2
Public Const NingunArma              As Integer = 2
Public Const EspadaMataDragonesIndex As Integer = 402
Public Const LAUDMAGICO              As Integer = 696
Public Const FLAUTAMAGICA            As Integer = 208
Public Const LAUDELFICO              As Integer = 1049
Public Const FLAUTAELFICA            As Integer = 1050
Public Const APOCALIPSIS_SPELL_INDEX As Integer = 25
Public Const DESCARGA_SPELL_INDEX    As Integer = 23
Public Const SLOTS_POR_FILA          As Byte = 5
Public Const PROB_ACUCHILLAR         As Byte = 20
Public Const DANO_ACUCHILLAR         As Single = 0.2
Public Const MAXMASCOTASENTRENADOR   As Byte = 7

Public Enum FXIDs
    FXWARP = 1
    FXMEDITARCHICO = 4
    FXMEDITARMEDIANO = 5
    FXMEDITARGRANDE = 6
    FXMEDITARXGRANDE = 16
    FXMEDITARXXGRANDE = 34
End Enum

Public Enum eTrigger
    NADA = 0
    BAJOTECHO = 1
    CASA = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
End Enum

Public Enum eTrigger6
    TRIGGER6_PERMITE = 1
    TRIGGER6_PROHIBE = 2
    TRIGGER6_AUSENTE = 3
End Enum

Public Const Bosque   As String = "BOSQUE"
Public Const Nieve    As String = "NIEVE"
Public Const Desierto As String = "DESIERTO"
Public Const Ciudad   As String = "CIUDAD"
Public Const Campo    As String = "CAMPO"
Public Const Dungeon  As String = "DUNGEON"

Public Enum eRestrict
    restrict_no = 0
    restrict_newbie = 1
    restrict_armada = 2
    restrict_caos = 3
    restrict_faccion = 4
End Enum

Public Enum TargetType
    uUsuarios = 1
    uNPC = 2
    uUsuariosYnpc = 3
    uTerreno = 4
End Enum

Public Enum TipoHechizo
    uPropiedades = 1
    uEstado = 2
    uMaterializa = 3
    uInvocacion = 4
End Enum

Public Const MAXUSERHECHIZOS               As Byte = 35
Public Const EsfuerzoTalarGeneral          As Byte = 4
Public Const EsfuerzoTalarLenador          As Byte = 2
Public Const EsfuerzoPescarPescador        As Byte = 1
Public Const EsfuerzoPescarGeneral         As Byte = 3
Public Const EsfuerzoExcavarMinero         As Byte = 2
Public Const EsfuerzoExcavarGeneral        As Byte = 5
Public Const FX_TELEPORT_INDEX             As Integer = 1
Public Const PORCENTAJE_MATERIALES_UPGRADE As Single = 0.85

Public Enum PartesCuerpo
    bCabeza = 1
    bPiernaIzquierda = 2
    bPiernaDerecha = 3
    bBrazoDerecho = 4
    bBrazoIzquierdo = 5
    bTorso = 6
End Enum

Public Const Guardias                       As Integer = 6
Public Const MAX_ORO_EDIT                   As Long = 5000000
Public Const MAX_VIDA_EDIT                  As Long = 30000
Public Const STANDARD_BOUNTY_HUNTER_MESSAGE As String = "Se te ha otorgado un premio por ayudar al proyecto reportando bugs, el mismo esta disponible en tu boveda."
Public Const TAG_USER_INVISIBLE             As String = "[INVISIBLE]"
Public Const TAG_CONSULT_MODE               As String = "[CONSULTA]"
Public Const MAXREP                         As Long = 6000000
Public Const MAXORO                         As Long = 200000000
Public Const MAXEXP                         As Long = 999999999
Public Const MAXUSERMATADOS                 As Long = 65000
Public Const MAXATRIBUTOS                   As Byte = 40
Public Const MINATRIBUTOS                   As Byte = 6
Public Const LingoteHierro                  As Integer = 386
Public Const LingotePlata                   As Integer = 387
Public Const LingoteOro                     As Integer = 388
Public Const Lena                           As Integer = 58
Public Const LenaElfica                     As Integer = 1006
Public Const MAXNPCS                        As Integer = 10000
Public Const MAXCHARS                       As Integer = 10000
Public Const HACHA_LENADOR                  As Integer = 127
Public Const HACHA_LENA_ELFICA              As Integer = 1005
Public Const PIQUETE_MINERO                 As Integer = 187
Public Const HACHA_LENADOR_NEWBIE           As Integer = 561
Public Const PIQUETE_MINERO_NEWBIE          As Integer = 562
Public Const CANA_PESCA_NEWBIE              As Integer = 563
Public Const SERRUCHO_CARPINTERO_NEWBIE     As Integer = 564
Public Const MARTILLO_HERRERO_NEWBIE        As Integer = 565
Public Const DAGA                           As Integer = 15
Public Const FOGATA_APAG                    As Integer = 136
Public Const FOGATA                         As Integer = 63
Public Const ORO_MINA                       As Integer = 194
Public Const PLATA_MINA                     As Integer = 193
Public Const HIERRO_MINA                    As Integer = 192
Public Const MARTILLO_HERRERO               As Integer = 389
Public Const SERRUCHO_CARPINTERO            As Integer = 198
Public Const ObjArboles                     As Integer = 4
Public Const RED_PESCA                      As Integer = 543
Public Const CANA_PESCA                     As Integer = 138
Public Const AMULETO_DEL_SILENCIO           As Integer = 1126

Public Enum eNPCType
    Comun = 0
    Revividor = 1
    GuardiaReal = 2
    Entrenador = 3
    Banquero = 4
    Noble = 5
    DRAGON = 6
    Timbero = 7
    Guardiascaos = 8
    Artesano = 9
    Pretoriano = 10
    Gobernador = 11
End Enum

Public Const MIN_APUNALAR   As Byte = 10
Public Const NUMSKILLS      As Byte = 20
Public Const NUMATRIBUTOS   As Byte = 5
Public Const NUMCLASES      As Byte = 12
Public Const NUMRAZAS       As Byte = 5
Public Const MAXSKILLPOINTS As Byte = 100
Public Const NUMCIUDADES    As Byte = 9

Public Enum eHeading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

Public Const MAXMASCOTAS   As Byte = 3
Public Const vlASALTO      As Integer = 100
Public Const vlASESINO     As Integer = 1000
Public Const vlCAZADOR     As Integer = 5
Public Const vlNoble       As Integer = 5
Public Const vlLadron      As Integer = 25
Public Const vlProleta     As Integer = 2
Public Const iCuerpoMuerto As Integer = 8
Public Const iCabezaMuerto As Integer = 500
Public Const iORO          As Byte = 12
Public Const Pescado       As Byte = 139

Public Enum PECES_POSIBLES
    PESCADO1 = 139
    PESCADO2 = 544
    PESCADO3 = 545
    PESCADO4 = 546
    PESCADO5 = 775
End Enum

Public Const NUM_PECES            As Integer = 5
Public ListaPeces(1 To NUM_PECES) As Integer

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
    pesca = 12
    Mineria = 13
    Carpinteria = 14
    Herreria = 15
    Liderazgo = 16
    Domar = 17
    Proyectiles = 18
    Wrestling = 19
    Navegacion = 20
    Equitacion = 21
End Enum

Public Enum eMochilas
    Mediana = 1
    Grande = 2
End Enum

Public Const FundirMetal = 88
Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5
End Enum

Public Const AdicionalHPGuerrero        As Byte = 2
Public Const AdicionalHPCazador         As Byte = 1
Public Const AumentoSTDef               As Byte = 15
Public Const AumentoStBandido           As Byte = AumentoSTDef + 3
Public Const AumentoSTLadron            As Byte = AumentoSTDef + 3
Public Const AumentoSTMago              As Byte = AumentoSTDef - 1
Public Const AumentoSTTrabajador        As Byte = AumentoSTDef + 25
Public Const SND_SWING                  As Byte = 2
Public Const SND_TALAR                  As Byte = 13
Public Const SND_PESCAR                 As Byte = 14
Public Const SND_MINERO                 As Byte = 15
Public Const SND_WARP                   As Byte = 3
Public Const SND_PUERTA                 As Byte = 5
Public Const SND_NIVEL                  As Byte = 6
Public Const SND_USERMUERTE             As Byte = 11
Public Const SND_IMPACTO                As Byte = 10
Public Const SND_IMPACTO2               As Byte = 12
Public Const SND_LENADOR                As Byte = 13
Public Const SND_FOGATA                 As Byte = 14
Public Const SND_AVE                    As Byte = 21
Public Const SND_AVE2                   As Byte = 22
Public Const SND_AVE3                   As Byte = 34
Public Const SND_GRILLO                 As Byte = 28
Public Const SND_GRILLO2                As Byte = 29
Public Const SND_SACARARMA              As Byte = 25
Public Const SND_ESCUDO                 As Byte = 37
Public Const SND_TRABAJO_HERRERO        As Byte = 41
Public Const SND_TRABAJO_CARPINTERO     As Byte = 42
Public Const SND_BEBER                  As Byte = 46
Public Const SND_RESUCITAR_SACERDOTE    As Byte = 213
Public Const SND_CURAR_SACERDOTE        As Byte = 214
Public Const MAX_INVENTORY_OBJS         As Integer = 10000
Public Const MAX_INVENTORY_SLOTS        As Byte = 35
Public Const MAX_NORMAL_INVENTORY_SLOTS As Byte = 25
Public Const SLOTS_PER_ROW_INVENTORY    As Byte = 5
Public Const FLAGORO                    As Integer = MAX_INVENTORY_SLOTS + 1

Public Enum eOBJType
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
    otEscudo = 16
    otCasco = 17
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

Public Enum ePocionType
    otAgilidad = 1
    otFuerza = 2
    otSalud = 3
    otMana = 4
    otCuraVeneno = 5
    otNegra = 6
End Enum

Public Enum eManualType
    otLiderazgo = 1
    otSupervivencia = 2
    otNavegacion = 3
    otInventSlots = 4
End Enum

Public STAT_MAXELV                    As Byte
Public Const STAT_MAXHP               As Integer = 999
Public Const STAT_MAXSTA              As Integer = 999
Public Const STAT_MAXMAN              As Integer = 9999
Public Const STAT_MAXHIT_UNDER36      As Byte = 99
Public Const STAT_MAXHIT_OVER36       As Integer = 999
Public Const STAT_MAXDEF              As Byte = 99
Public Const ELU_SKILL_INICIAL        As Byte = 200
Public Const EXP_ACIERTO_SKILL        As Byte = 50
Public Const EXP_FALLO_SKILL          As Byte = 20

Public Type tObservacion
    Creador As String
    Fecha As Date
    Detalles As String
End Type

Public Type tRecord
    Usuario As String
    Motivo As String
    Creador As String
    Fecha As Date
    NumObs As Byte
    Obs() As tObservacion
End Type

Public Type tHechizo
    Tipo As TipoHechizo
    WAV As Integer
    FXgrh As Integer
    loops As Byte
    SubeHP As Byte
    MinHp As Integer
    MaxHp As Integer
    SubeMana As Byte
    MiMana As Integer
    MaMana As Integer
    SubeSta As Byte
    MinSta As Integer
    MaxSta As Integer
    SubeHam As Byte
    MinHam As Integer
    MaxHam As Integer
    SubeSed As Byte
    MinSed As Integer
    MaxSed As Integer
    SubeAgilidad As Byte
    MinAgilidad As Integer
    MaxAgilidad As Integer
    SubeFuerza As Byte
    MinFuerza As Integer
    MaxFuerza As Integer
    SubeCarisma As Byte
    MinCarisma As Integer
    MaxCarisma As Integer
    Invisibilidad As Byte
    Paraliza As Byte
    Inmoviliza As Byte
    RemoverParalisis As Byte
    RemoverEstupidez As Byte
    CuraVeneno As Byte
    Envenena As Byte
    Maldicion As Byte
    RemoverMaldicion As Byte
    Bendicion As Byte
    Estupidez As Byte
    Ceguera As Byte
    Revivir As Byte
    Morph As Byte
    Mimetiza As Byte
    RemueveInvisibilidadParcial As Byte
    Warp As Byte
    Invoca As Byte
    NumNpc As Integer
    cant As Integer
    MinSkill As Integer
    ManaRequerido As Integer
    StaRequerido As Integer
    Target As TargetType
    NeedStaff As Integer
    StaffAffected As Boolean
End Type

Public Type LevelSkill
    LevelValue As Integer
End Type

Public Type UserObj
    ObjIndex As Integer
    Amount As Integer
    Equipped As Byte
End Type

Public Type Inventario
    Object(1 To MAX_INVENTORY_SLOTS) As UserObj
    WeaponEqpObjIndex As Integer
    WeaponEqpSlot As Byte
    ArmourEqpObjIndex As Integer
    ArmourEqpSlot As Byte
    EscudoEqpObjIndex As Integer
    EscudoEqpSlot As Byte
    CascoEqpObjIndex As Integer
    CascoEqpSlot As Byte
    MunicionEqpObjIndex As Integer
    MunicionEqpSlot As Byte
    AnilloEqpObjIndex As Integer
    AnilloEqpSlot As Byte
    BarcoObjIndex As Integer
    BarcoSlot As Byte
    MochilaEqpObjIndex As Integer
    MochilaEqpSlot As Byte
    MonturaObjIndex As Integer
    MonturaEqpSlot As Byte
    NroItems As Integer
End Type

Public Type tPartyData
    PIndex As Integer
    RemXP As Double
    TargetUser As Integer
End Type

Public Type Position
    X As Integer
    Y As Integer
End Type

Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

Public Type FXdata
    Nombre As String
    GrhIndex As Long
    Delay As Integer
End Type

Public Type Char
    Escribiendo As Byte
    CharIndex As Integer
    Head As Integer
    body As Integer
    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim As Integer
    FX As Integer
    loops As Integer
    heading As eHeading
End Type

Public Type CraftingItem
    ObjIndex As Integer
    Amount As Integer
End Type

Public Const MAX_ITEMS_CRAFTEO As Byte = 4

Public Type ObjData
    Name As String
    OBJType As eOBJType
    GrhIndex As Long
    GrhSecundario As Long
    MAXITEMS As Integer
    Conte As Inventario
    Apunala As Byte
    Acuchilla As Byte
    HechizoIndex As Integer
    ForoID As String
    MinHp As Integer
    MaxHp As Integer
    MineralIndex As Integer
    LingoteInex As Integer
    proyectil As Integer
    Municion As Integer
    Crucial As Byte
    Newbie As Integer
    MinSta As Integer
    TipoPocion As Byte
    MaxModificador As Integer
    MinModificador As Integer
    DuracionEfecto As Long
    LingoteIndex As Integer
    MinHIT As Integer
    MaxHIT As Integer
    MinHam As Integer
    MinSed As Integer
    def As Integer
    MinDef As Integer
    MaxDef As Integer
    Ropaje As Integer
    WeaponAnim As Integer
    WeaponRazaEnanaAnim As Integer
    ShieldAnim As Integer
    CascoAnim As Integer
    Valor As Long
    Cerrada As Integer
    Llave As Byte
    Clave As Long
    Radio As Integer
    MochilaType As Byte
    Guante As Byte
    IndexAbierta As Integer
    IndexCerrada As Integer
    IndexCerradaLlave As Integer
    RazaEnana As Byte
    RazaDrow As Byte
    RazaElfa As Byte
    RazaGnoma As Byte
    RazaHumana As Byte
    Mujer As Byte
    Hombre As Byte
    Envenena As Byte
    Paraliza As Byte
    Agarrable As Byte
    LingH As Integer
    LingO As Integer
    LingP As Integer
    Madera As Integer
    MaderaElfica As Integer
    SkHerreria As Integer
    SkCarpinteria As Integer
    ItemCrafteo() As CraftingItem
    MinSkill As Byte
    SkillRequerido As eSkill
    SkillCantidad As Byte
    MinLevel As Byte
    texto As String
    ClaseProhibida(1 To NUMCLASES) As eClass
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    Real As Integer
    Caos As Integer
    NoSeCae As Integer
    StaffPower As Integer
    StaffDamageBonus As Integer
    DefensaMagicaMax As Integer
    DefensaMagicaMin As Integer
    Refuerzo As Byte
    ImpideParalizar As Boolean
    ImpideAturdir As Boolean
    ImpideCegar As Boolean
    Log As Byte
    NoLog As Byte
    Upgrade As Integer
End Type

Public Type obj
    ObjIndex As Integer
    Amount As Integer
End Type

Public Type tQuestNpc
    NpcIndex As Integer
    Amount As Integer
End Type
 
Public Type tUserQuest
    NPCsKilled() As Integer
    QuestIndex As Integer
End Type
 
Public Type tQuestStats
    Quests(1 To MAXUSERQUESTS) As tUserQuest
    NumQuestsDone As Integer
    QuestsDone() As Integer
End Type

Public Type ModClase
    Evasion As Double
    AtaqueArmas As Double
    AtaqueProyectiles As Double
    AtaqueWrestling As Double
    DanoArmas As Double
    DanoProyectiles As Double
    DanoWrestling As Double
    Escudo As Double
End Type

Public Type ModRaza
    Fuerza As Single
    Agilidad As Single
    Inteligencia As Single
    Carisma As Single
    Constitucion As Single
End Type

Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40

Public Type BancoInventario
    Object(1 To MAX_BANCOINVENTORY_SLOTS) As UserObj
    NroItems As Integer
End Type

Public Enum eNickColor
    ieCriminal = &H1
    ieCiudadano = &H2
    ieAtacable = &H4
End Enum

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

Public Const MAX_STICKY_POST  As Byte = 10
Public Const MAX_GENERAL_POST As Byte = 35

Public Type tForo
    StickyTitle(1 To MAX_STICKY_POST) As String
    StickyPost(1 To MAX_STICKY_POST) As String
    GeneralTitle(1 To MAX_GENERAL_POST) As String
    GeneralPost(1 To MAX_GENERAL_POST) As String
End Type

Public Type tQuest
    Nombre As String
    Desc As String
    RequiredLevel As Byte
    RequiredOBJs As Byte
    RequiredOBJ() As obj
    RequiredNPCs As Byte
    RequiredNPC() As tQuestNpc
    RewardGLD As Long
    RewardEXP As Long
    RewardOBJs As Byte
    RewardOBJ() As obj
End Type

Public Type tReputacion
    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long
    Promedio As Long
End Type

Public Type UserStats
    Gld As Long
    Banco As Long
    MaxHp As Integer
    MinHp As Integer
    MaxSta As Integer
    MinSta As Integer
    MaxMAN As Integer
    MinMAN As Integer
    MaxHIT As Integer
    MinHIT As Integer
    MaxHam As Integer
    MinHam As Integer
    MaxAGU As Integer
    MinAGU As Integer
    def As Integer
    Exp As Double
    ELV As Byte
    ELU As Long
    UserSkills(1 To NUMSKILLS) As Byte
    UserAtributos(1 To NUMATRIBUTOS) As Byte
    UserAtributosBackUP(1 To NUMATRIBUTOS) As Byte
    UserHechizos(1 To MAXUSERHECHIZOS) As Integer
    UsuariosMatados As Long
    NPCsMuertos As Integer
    SkillPts As Integer
    ExpSkills(1 To NUMSKILLS) As Long
    EluSkills(1 To NUMSKILLS) As Long
End Type

Public Type UserFlags
    GMRequested As Integer
    SlotReto As Byte
    SlotRetoUser As Byte
    SlotCarcel As Integer
    Muerto As Byte
    Escondido As Byte
    Comerciando As Boolean
    UserLogged As Boolean
    Meditando As Boolean
    Descuento As String
    Hambre As Byte
    Sed As Byte
    PuedeMoverse As Byte
    TimerLanzarSpell As Long
    PuedeTrabajar As Byte
    Envenenado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    Estupidez As Byte
    Ceguera As Byte
    invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    Oculto As Byte
    Desnudo As Byte
    Descansar As Boolean
    Hechizo As Integer
    TomoPocion As Boolean
    TipoPocion As Byte
    NoPuedeSerAtacado As Boolean
    AtacablePor As Integer
    ShareNpcWith As Integer
    Vuela As Byte
    Navegando As Byte
    Equitando As Byte
    Seguro As Boolean
    SeguroResu As Boolean
    DuracionEfecto As Single
    TargetNPC As Integer
    TargetNpcTipo As eNPCType
    OwnedNpc As Integer
    NpcInv As Integer
    Ban As Byte
    AdministrativeBan As Byte
    TargetUser As Integer
    TargetObj As Integer
    TargetObjMap As Integer
    TargetObjX As Integer
    TargetObjY As Integer
    TargetMap As Integer
    TargetX As Integer
    TargetY As Integer
    TargetObjInvIndex As Integer
    TargetObjInvSlot As Integer
    AtacadoPorNpc As Integer
    AtacadoPorUser As Integer
    NPCAtacado As Integer
    Ignorado As Boolean
    EnConsulta As Boolean
    SendDenounces As Boolean
    StatsChanged As Byte
    Privilegios As PlayerType
    PrivEspecial As Boolean
    ValCoDe As Integer
    LastCrimMatado As String
    LastCiudMatado As String
    OldBody As Integer
    OldHead As Integer
    AdminInvisible As Byte
    AdminPerseguible As Boolean
    ChatColor As Long
    TimesWalk As Long
    StartWalk As Long
    CountSH As Long
    UltimoMensaje As Byte
    Silenciado As Byte
    Mimetizado As Byte
    lastMap As Integer
    Traveling As Byte
    ParalizedBy As String
    ParalizedByIndex As Integer
    ParalizedByNpcIndex As Integer
End Type

Public Type UserCounters
    TimeFight As Long
    IdleCount As Single
    AttackCounter As Integer
    HPCounter As Single
    STACounter As Single
    Frio As Single
    Lava As Single
    COMCounter As Single
    AGUACounter As Single
    Veneno As Single
    Paralisis As Single
    Ceguera As Single
    Estupidez As Single
    MonturaCounter As Single
    Invisibilidad As Single
    TiempoOculto As Single
    Mimetismo As Single
    PiqueteC As Long
    ContadorPiquete As Long
    Pena As Long
    SendMapCounter As WorldPos
    Saliendo As Boolean
    Salir As Integer
    tInicioMeditar As Long
    bPuedeMeditar As Boolean
    TimerLanzarSpell As Long
    TimerPuedeAtacar As Long
    TimerPuedeUsarArco As Long
    TimerPuedeTrabajar As Long
    TimerUsar As Long
    TimerMagiaGolpe As Long
    TimerGolpeMagia As Long
    TimerGolpeUsar As Long
    TimerPuedeSerAtacado As Long
    TimerPerteneceNpc As Long
    TimerEstadoAtacable As Long
    Trabajando As Long
    Ocultando As Long
    Cheat As TimeIntervalos
    failedUsageAttempts As Long
    goHome As Long
    AsignedSkills As Byte
    PacketsTick As Byte
End Type

Public Type tFacciones
    ArmadaReal As Byte
    FuerzasCaos As Byte
    CriminalesMatados As Long
    CiudadanosMatados As Long
    RecompensasReal As Long
    RecompensasCaos As Long
    RecibioExpInicialReal As Byte
    RecibioExpInicialCaos As Byte
    RecibioArmaduraReal As Byte
    RecibioArmaduraCaos As Byte
    Reenlistadas As Byte
    NivelIngreso As Integer
    FechaIngreso As String
    MatadosIngreso As Integer
    NextRecompensa As Integer
End Type

Public Type tCrafting
    Cantidad As Long
    PorCiclo As Integer
End Type

Public Type AccountUser
    Name As String
    body As Integer
    Head As Integer
    weapon As Integer
    shield As Integer
    helmet As Integer
    Class As Byte
    race As Byte
    Map As Integer
    level As Byte
    Gold As Long
    criminal As Boolean
    dead As Boolean
    gameMaster As Boolean
End Type

Public Type tUserRetoTemp
    Tipo As eTipoReto
    RequiredGld As Long
    Users() As String
    Accepts() As Byte
End Type

Public Type Amigos
    Nombre As String
    Ignorado As Byte
    index As Integer
End Type

Public Type User
    PosAnt As WorldPos
    RetoTemp As tUserRetoTemp
    Name As String
    ID As Long
    AccountHash As String
    showName As Boolean
    Amigos(1 To MAXAMIGOS) As Amigos
    Quien As String
    Char As Char
    CharMimetizado As Char
    OrigChar As Char
    Desc As String
    DescRM As String
    Clase As eClass
    raza As eRaza
    Genero As eGenero
    Email As String
    Hogar As eCiudad
    Invent As Inventario
    Pos As WorldPos
    ConnIDValida As Boolean
    ConnID As Long
    BancoInvent As BancoInventario
    Counters As UserCounters
    Construir As tCrafting
    MascotasIndex(1 To MAXMASCOTAS) As Integer
    MascotasType(1 To MAXMASCOTAS) As Integer
    NroMascotas As Integer
    Stats As UserStats
    flags As UserFlags
    Reputacion As tReputacion
    Faccion As tFacciones
    #If ConUpTime Then
        LogOnTime As Date
        UpTime As Long
    #End If
    IP As String
    ComUsu As tCOmercioUsuario
    GuildIndex As Integer
    FundandoGuildAlineacion As ALINEACION_GUILD
    EscucheClan As Integer
    PartyIndex As Integer
    PartySolicitud As Integer
    KeyCrypt As Integer
    AreasInfo As AreaInfo
    outgoingData As clsByteQueue
    incomingData As clsByteQueue
    CurrentInventorySlots As Byte
    CentinelaUsuario As CentinelaUser
    cvcUser As cvc_User
    QuestStats As tQuestStats
    Redundance As Byte
End Type

Public Type NPCStats
    Alineacion As Integer
    MaxHp As Long
    MinHp As Long
    MaxHIT As Integer
    MinHIT As Integer
    def As Integer
    defM As Integer
End Type

Public Type NpcCounters
    Paralisis As Integer
    TiempoExistencia As Single
    Ataque As Long
End Type

Public Type NPCFlags
    AfectaParalisis As Byte
    Domable As Integer
    Respawn As Byte
    NPCActive As Boolean
    Follow As Boolean
    Faccion As Byte
    AtacaDoble As Byte
    LanzaSpells As Byte
    ExpCount As Long
    OldMovement As TipoAI
    OldHostil As Byte
    AguaValida As Byte
    TierraInvalida As Byte
    Sound As Integer
    AttackedBy As String
    AttackedFirstBy As String
    BackUp As Byte
    RespawnOrigPos As Byte
    Envenenado As Byte
    Paralizado As Byte
    Inmovilizado As Byte
    invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    SiguiendoGm As Boolean
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
End Type

Public Type tCriaturasEntrenador
    NpcIndex As Integer
    NpcName As String
    tmpIndex As Integer
End Type

Public Type NpcPathFindingInfo
    Path() As tVertice
    Target As Position
    PathLenght As Integer
    CurPos As Integer
    TargetUser As Integer
    NoPath As Boolean
End Type

Public Type tDrops
    ObjIndex As Integer
    Amount As Long
End Type

Public Const MAX_NPC_DROPS As Byte = 5

Public Type npc
    Name As String
    Char As Char
    Desc As String
    NPCtype As eNPCType
    Numero As Integer
    InvReSpawn As Byte
    Comercia As Integer
    Target As Long
    TargetNPC As Long
    TipoItems As Integer
    Veneno As Byte
    Pos As WorldPos
    Orig As WorldPos
    SkillDomar As Integer
    Movement As TipoAI
    Attackable As Byte
    Hostile As Byte
    PoderAtaque As Long
    PoderEvasion As Long
    Owner As Integer
    GiveEXP As Long
    GiveGLD As Long
    Drop(1 To MAX_NPC_DROPS) As tDrops
    QuestNumber As Integer
    Stats As NPCStats
    flags As NPCFlags
    Contadores As NpcCounters
    Invent As Inventario
    NroExpresiones As Byte
    Expresiones() As String
    NroSpells As Byte
    Spells() As Integer
    NroCriaturas As Integer
    Criaturas() As tCriaturasEntrenador
    MaestroUser As Integer
    MaestroNpc As Integer
    Mascotas As Integer
    PFINFO As NpcPathFindingInfo
    AreasInfo As AreaInfo
    Ciudad As Byte
    ClanIndex As Integer
End Type

Public Type MapBlock
    Blocked As Byte
    Graphic(1 To 4) As Long
    Userindex As Integer
    NpcIndex As Integer
    ObjInfo As obj
    TileExit As WorldPos
    trigger As eTrigger
End Type

Type MapInfo
    NumUsers As Integer
    Music As String
    MusicMp3 As String
    Name As String
    StartPos As WorldPos
    OnDeathGoTo As WorldPos
    MapVersion As Integer
    Pk As Boolean
    MagiaSinEfecto As Byte
    NoEncriptarMP As Byte
    InviSinEfecto As Byte
    ResuSinEfecto As Byte
    OcultarSinEfecto As Byte
    InvocarSinEfecto As Byte
    RoboNpcsPermitido As Byte
    Terreno As String
    Zona As String
    Restringir As Byte
    BackUp As Byte
End Type

Public SERVERONLINE                       As Boolean
Public ULTIMAVERSION                      As String
Public BackUp                             As Boolean
Public ListaRazas(1 To NUMRAZAS)          As String
Public SkillsNames(1 To NUMSKILLS)        As String
Public ListaClases(1 To NUMCLASES)        As String
Public ListaAtributos(1 To NUMATRIBUTOS)  As String
Public RecordUsuariosOnline                     As Long

Public IniPath                            As String
Public CharPath                           As String
Public AccountPath                        As String
Public MapPath                            As String
Public DatPath                            As String
Public MinXBorder                         As Byte
Public MaxXBorder                         As Byte
Public MinYBorder                         As Byte
Public MaxYBorder                         As Byte
Public NumUsers                           As Integer
Public LastUser                           As Integer
Public LastChar                           As Integer
Public NumChars                           As Integer
Public LastNPC                            As Integer
Public NumNPCs                            As Integer
Public NumFX                              As Integer
Public NumMaps                            As Integer
Public NumObjDatas                        As Integer
Public NumeroHechizos                     As Integer
Public AllowMultiLogins                   As Boolean
Public IdleLimit                          As Integer
Public LimiteConexionesPorIp              As Byte
Public MaxUsers                           As Integer
Public HideMe                             As Boolean
Public LastBackup                         As String
Public Minutos                            As String
Public haciendoBK                         As Boolean
Public PuedeCrearPersonajes               As Integer
Public ServerSoloGMs                      As Integer
Public NumRecords                         As Integer
Public EnPausa                            As Boolean
Public EnTesting                          As Boolean
Public iniHappyHourActivado As Boolean
Public HappyHour As Single
Public HappyHourActivated As Boolean

Public Type tHappyHour
    Multi As Single
    Hour As Integer
End Type

Public HappyHourDays(1 To 7) As tHappyHour
Public UserList()                         As User
Public Npclist(1 To MAXNPCS)              As npc
Public MapData()                          As MapBlock
Public MapInfo()                          As MapInfo
Public Hechizos()                         As tHechizo
Public CharList(1 To MAXCHARS)            As Integer
Public ObjData()                          As ObjData
Public FX()                               As FXdata
Public SpawnList()                        As tCriaturasEntrenador
Public LevelSkill(1 To 50)                As LevelSkill
Public ForbidenNames()                    As String
Public ArmasHerrero()                     As Integer
Public ArmadurasHerrero()                 As Integer
Public ObjCarpintero()                    As Integer
Public ObjArtesano()                      As Integer
Public BanIps                             As Collection
Public Parties(1 To MAX_PARTIES)          As clsParty
Public ModClase(1 To NUMCLASES)           As ModClase
Public ModRaza(1 To NUMRAZAS)             As ModRaza
Public ModVida(1 To NUMCLASES)            As Double
Public DistribucionEnteraVida(1 To 5)     As Integer
Public DistribucionSemienteraVida(1 To 4) As Integer
Public Ciudades(1 To NUMCIUDADES)         As WorldPos
Public distanceToCities()                 As HomeDistance
Public QuestList()                        As tQuest
Public Records()                          As tRecord

Type HomeDistance
    distanceToCity(1 To NUMCIUDADES) As Integer
End Type

Public Nix             As WorldPos
Public Ullathorpe      As WorldPos
Public Banderbill      As WorldPos
Public Lindos          As WorldPos
Public Arghal          As WorldPos
Public Arkhein         As WorldPos
Public Nemahuak        As WorldPos
Public Prision         As WorldPos
Public Libertad        As WorldPos
Public Gotland         As WorldPos
Public Perdida         As WorldPos
Public Totem           As WorldPos
Public CustomSpawnMap  As WorldPos
Public Ayuda           As cCola
Public Denuncias       As cCola
Public ConsultaPopular As ConsultasPopulares

Public Declare Function writeprivateprofilestring _
               Lib "kernel32" _
               Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                   ByVal lpKeyname As Any, _
                                                   ByVal lpString As String, _
                                                   ByVal lpfilename As String) As Long

Public Declare Function GetPrivateProfileString _
               Lib "kernel32" _
               Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                 ByVal lpKeyname As Any, _
                                                 ByVal lpdefault As String, _
                                                 ByVal lpreturnedstring As String, _
                                                 ByVal nSize As Long, _
                                                 ByVal lpfilename As String) As Long

Public Declare Sub ZeroMemory _
               Lib "kernel32.dll" _
               Alias "RtlZeroMemory" (ByRef destination As Any, _
                                      ByVal Length As Long)

Public Enum e_ObjetosCriticos
    Manzana = 1
    Manzana2 = 2
    ManzanaNewbie = 467
End Enum

Public Enum eMessages
    NPCSwing
    NPCKillUser
    BlockedWithShieldUser
    BlockedWithShieldother
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
    Home
    CancelHome
    FinishHome
    UserMuerto
    NpcInmune
    Hechizo_HechiceroMSG_NOMBRE
    Hechizo_HechiceroMSG_ALGUIEN
    Hechizo_HechiceroMSG_CRIATURA
    Hechizo_PropioMSG
    Hechizo_TargetMSG
End Enum

Public Enum eGMCommands
    GMMessage = 1           '/GMSG
    showName                '/SHOWNAME
    OnlineRoyalArmy         '/ONLINEREAL
    OnlineChaosLegion       '/ONLINECAOS
    GoNearby                '/IRCERCA
    comment                 '/REM
    serverTime              '/HORA
    Where                   '/DONDE
    CreaturesInMap          '/NENE
    WarpMeToTarget          '/TELEPLOC
    WarpChar                '/TELEP
    Silence                 '/SILENCIAR
    SOSShowList             '/SHOW SOS
    SOSRemove               'SOSDONE
    GoToChar                '/IRA
    invisible               '/INVISIBLE
    GMPanel                 '/PANELGM
    RequestUserList         'LISTUSU
    Working                 '/TRABAJANDO
    Hiding                  '/OCULTANDO
    Jail                    '/CARCEL
    KillNPC                 '/RMATA
    WarnUser                '/ADVERTENCIA
    EditChar                '/MOD
    RequestCharInfo         '/INFO
    RequestCharStats        '/STAT
    RequestCharGold         '/BAL
    RequestCharInventory    '/INV
    RequestCharBank         '/BOV
    RequestCharSkills       '/SKILLS
    ReviveChar              '/REVIVIR
    OnlineGM                '/ONLINEGM
    OnlineMap               '/ONLINEMAP
    Forgive                 '/PERDON
    Kick                    '/ECHAR
    Execute                 '/EJECUTAR
    BanChar                 '/BAN
    UnbanChar               '/UNBAN
    NPCFollow               '/SEGUIR
    SummonChar              '/SUM
    SpawnListRequest        '/CC
    SpawnCreature           'SPA
    ResetNPCInventory       '/RESETINV
    ServerMessage           '/RMSG
    NickToIP                '/NICK2IP
    IPToNick                '/IP2NICK
    GuildOnlineMembers      '/ONCLAN
    TeleportCreate          '/CT
    TeleportDestroy         '/DT
    RainToggle              '/LLUVIA
    SetCharDescription      '/SETDESC
    ForceMP3ToMap          '/FORCEMP3MAP
    ForceMIDIToMap          '/FORCEMIDIMAP
    ForceWAVEToMap          '/FORCEWAVMAP
    RoyalArmyMessage        '/REALMSG
    ChaosLegionMessage      '/CAOSMSG
    CitizenMessage          '/CIUMSG
    CriminalMessage         '/CRIMSG
    TalkAsNPC               '/TALKAS
    DestroyAllItemsInArea   '/MASSDEST
    AcceptRoyalCouncilMember '/ACEPTCONSE
    AcceptChaosCouncilMember '/ACEPTCONSECAOS
    ItemsInTheFloor         '/PISO
    MakeDumb                '/ESTUPIDO
    MakeDumbNoMore          '/NOESTUPIDO
    DumpIPTables            '/DUMPSECURITY
    CouncilKick             '/KICKCONSE
    SetTrigger              '/TRIGGER
    AskTrigger              '/TRIGGER with no args
    BannedIPList            '/BANIPLIST
    BannedIPReload          '/BANIPRELOAD
    GuildMemberList         '/MIEMBROSCLAN
    GuildBan                '/BANCLAN
    BanIP                   '/BANIP
    UnbanIP                 '/UNBANIP
    CreateItem              '/CI
    DestroyItems            '/DEST
    ChaosLegionKick         '/NOCAOS
    RoyalArmyKick           '/NOREAL
    ForceMP3All             '/FORCEMP3
    ForceMIDIAll            '/FORCEMIDI
    ForceWAVEAll            '/FORCEWAV
    RemovePunishment        '/BORRARPENA
    TileBlockedToggle       '/BLOQ
    KillNPCNoRespawn        '/MATA
    KillAllNearbyNPCs       '/MASSKILL
    LastIP                  '/LASTIP
    ChangeMOTD              '/MOTDCAMBIA
    SetMOTD                 'ZMOTD
    SystemMessage           '/SMSG
    CreateNPC               '/ACC y /RACC
    ImperialArmour          '/AI1 - 4
    ChaosArmour             '/AC1 - 4
    NavigateToggle          '/NAVE
    ServerOpenToUsersToggle '/HABILITAR
    TurnOffServer           '/APAGAR
    TurnCriminal            '/CONDEN
    ResetFactions           '/RAJAR
    RemoveCharFromGuild     '/RAJARCLAN
    RequestCharMail         '/LASTEMAIL
    AlterPassword           '/APASS
    AlterMail               '/AEMAIL
    AlterName               '/ANAME
    DoBackUp                '/DOBACKUP
    ShowGuildMessages       '/SHOWCMSG
    SaveMap                 '/GUARDAMAPA
    ChangeMapInfoPK         '/MODMAPINFO PK
    ChangeMapInfoBackup     '/MODMAPINFO BACKUP
    ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
    ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
    ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
    ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
    ChangeMapInfoLand       '/MODMAPINFO TERRENO
    ChangeMapInfoZone       '/MODMAPINFO ZONA
    ChangeMapInfoStealNpc   '/MODMAPINFO ROBONPC
    ChangeMapInfoNoOcultar  '/MODMAPINFO OCULTARSINEFECTO
    ChangeMapInfoNoInvocar  '/MODMAPINFO INVOCARSINEFECTO
    SaveChars               '/GRABAR
    CleanSOS                '/BORRAR SOS
    ShowServerForm          '/SHOW INT
    night                   '/NOCHE
    KickAllChars            '/ECHARTODOSPJS
    ReloadNPCs              '/RELOADNPCS
    ReloadServerIni         '/RELOADSINI
    ReloadSpells            '/RELOADHECHIZOS
    ReloadObjects           '/RELOADOBJ
    Restart                 '/REINICIAR
    ResetAutoUpdate         '/AUTOUPDATE
    ChatColor               '/CHATCOLOR
    Ignored                 '/IGNORADO
    CheckSlot               '/SLOT
    SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
    CreatePretorianClan     '/CREARPRETORIANOS
    RemovePretorianClan     '/ELIMINARPRETORIANOS
    EnableDenounces         '/DENUNCIAS
    ShowDenouncesList       '/SHOW DENUNCIAS
    MapMessage              '/MAPMSG
    SetDialog               '/SETDIALOG
    Impersonate             '/IMPERSONAR
    Imitate                 '/MIMETIZAR
    RecordAdd
    RecordRemove
    RecordAddObs
    RecordListRequest
    RecordDetailsRequest
    ExitDestroy             '/DE
    ToggleCentinelActivated '/CENTINELAACTIVADO
    SearchNpc               '/BUSCAR
    SearchObj               '/BUSCAR
    LimpiarMundo            '/LIMPIARMUNDO
End Enum

Public Const MATRIX_INITIAL_MAP                     As Integer = 1
Public Const GOHOME_PENALTY                         As Integer = 5
Public Const GM_MAP                                 As Integer = 49
Public Const TELEP_OBJ_INDEX                        As Integer = 1012
Public Const HUMANO_H_PRIMER_CABEZA                 As Integer = 1
Public Const HUMANO_H_ULTIMA_CABEZA                 As Integer = 40
Public Const ELFO_H_PRIMER_CABEZA                   As Integer = 101
Public Const ELFO_H_ULTIMA_CABEZA                   As Integer = 122
Public Const DROW_H_PRIMER_CABEZA                   As Integer = 201
Public Const DROW_H_ULTIMA_CABEZA                   As Integer = 221
Public Const ENANO_H_PRIMER_CABEZA                  As Integer = 301
Public Const ENANO_H_ULTIMA_CABEZA                  As Integer = 319
Public Const GNOMO_H_PRIMER_CABEZA                  As Integer = 401
Public Const GNOMO_H_ULTIMA_CABEZA                  As Integer = 416
Public Const HUMANO_M_PRIMER_CABEZA                 As Integer = 70
Public Const HUMANO_M_ULTIMA_CABEZA                 As Integer = 89
Public Const ELFO_M_PRIMER_CABEZA                   As Integer = 170
Public Const ELFO_M_ULTIMA_CABEZA                   As Integer = 188
Public Const DROW_M_PRIMER_CABEZA                   As Integer = 270
Public Const DROW_M_ULTIMA_CABEZA                   As Integer = 288
Public Const ENANO_M_PRIMER_CABEZA                  As Integer = 370
Public Const ENANO_M_ULTIMA_CABEZA                  As Integer = 384
Public Const GNOMO_M_PRIMER_CABEZA                  As Integer = 470
Public Const GNOMO_M_ULTIMA_CABEZA                  As Integer = 484
Public Const GUANTE_HURTO                           As Integer = 873
Public Const ESPADA_VIKINGA                         As Integer = 123
Public ClanPretoriano()                             As clsClanPretoriano
Public Const MAX_DENOUNCES                          As Integer = 20
Public Const MENSAJE_REY_CAOS                       As String = "Esperabas pasar desapercibido, intruso? Los servidores del Demonio no son bienvenidos, Guardias, a el!"
Public Const MENSAJE_REY_CRIMINAL_NOENLISTABLE      As String = "Tus pecados son grandes, pero aun asi puedes redimirte. El pasado deja huellas, pero aun puedes limpiar tu alma."
Public Const MENSAJE_REY_CRIMINAL_ENLISTABLE        As String = "Limpia tu reputacion y paga por los delitos cometidos. Un miembro de la Armada Real debe tener un comportamiento ejemplar."
Public Const MENSAJE_DEMONIO_REAL                   As String = "Lacayo de Tancredo, ve y dile a tu gente que nadie pisara estas tierras si no se arrodilla ante mi."
Public Const MENSAJE_DEMONIO_CIUDADANO_NOENLISTABLE As String = "Tu indecision te ha condenado a una vida sin sentido, aun tienes eleccion... Pero ten mucho cuidado, mis hordas nunca descansan."
Public Const MENSAJE_DEMONIO_CIUDADANO_ENLISTABLE   As String = "Siento el miedo por tus venas. Deja de ser escoria y unete a mis filas, sabras que es el mejor camino."
Public Administradores                              As clsIniManager

Public Enum e_SoundIndex
    MUERTE_HOMBRE = 11
    MUERTE_MUJER = 74
    FLECHA_IMPACTO = 65
    CONVERSION_BARCO = 55
    MORFAR_MANZANA = 82
    SOUND_COMIDA = 7
    MUERTE_MUJER_AGUA = 211
    MUERTE_HOMBRE_AGUA = 212
End Enum

Public ExpMultiplier        As Integer
Public OroMultiplier        As Integer
Public OficioMultiplier     As Integer
Public DiceMinimum          As Integer
Public DiceMaximum          As Integer
Public DropItemsAlMorir     As Boolean
Public ArtesaniaCosto       As Long
Public ContadorAntiPiquete  As Integer
Public MinutosCarcelPiquete As Integer
Public InventarioUsarConfiguracionPersonalizada As Boolean
Public EstadisticasInicialesUsarConfiguracionPersonalizada As Boolean
Public UsarMundoPropio As Boolean
Public OroDirectoABille As Boolean
Public ConexionAPI As Boolean
Public ApiUrlServer As String
Public ApiPath As String
Public ApiNodeJsTaskId As Double
Public MundoSeleccionado As String
Public DescripcionServidor As String
Public NombreServidor As String
Public IpPublicaServidor As String
#If AntiExternos Then
    Public Security As New clsSecurity
#End If
