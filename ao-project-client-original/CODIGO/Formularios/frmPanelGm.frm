VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPanelGm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel GM"
   ClientHeight    =   8760
   ClientLeft      =   18300
   ClientTop       =   4590
   ClientWidth     =   4335
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   4335
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   7695
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   360
      Width           =   3855
      Begin VB.CommandButton Command7 
         Caption         =   "Ver Procesos"
         Height          =   315
         Left            =   0
         TabIndex        =   129
         Top             =   840
         Width           =   3735
      End
      Begin VB.CommandButton cmdConsulta 
         Caption         =   "Consulta"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   87
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdNOREAL 
         Caption         =   "Explulsar de la Armada"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   86
         Top             =   6360
         Width           =   1815
      End
      Begin VB.CommandButton cmdNOCAOS 
         Caption         =   "Expulsar de Caos"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         TabIndex        =   85
         Top             =   6360
         Width           =   1815
      End
      Begin VB.CommandButton cmdKICKCONSE 
         Caption         =   "Degradar Consejero"
         CausesValidation=   0   'False
         Height          =   915
         Left            =   2400
         TabIndex        =   84
         Top             =   6720
         Width           =   1335
      End
      Begin VB.CommandButton cmdACEPTCONSECAOS 
         Caption         =   "Ascender a consejero del Caos"
         CausesValidation=   0   'False
         Height          =   435
         Left            =   0
         TabIndex        =   83
         Top             =   7200
         Width           =   2295
      End
      Begin VB.CommandButton cmdACEPTCONSE 
         Caption         =   "Ascender a Consejero Real"
         CausesValidation=   0   'False
         Height          =   435
         Left            =   0
         TabIndex        =   82
         Top             =   6720
         Width           =   2295
      End
      Begin VB.ComboBox cboListaUsus 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   54
         Top             =   480
         Width           =   3675
      End
      Begin VB.CommandButton cmdActualiza 
         Caption         =   "Actualizar"
         CausesValidation=   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   120
         Width           =   3675
      End
      Begin VB.CommandButton cmdIRCERCA 
         Caption         =   "Ir Cerca"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         TabIndex        =   52
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdDONDE 
         Caption         =   "Ubicar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         TabIndex        =   51
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdPENAS 
         Caption         =   "Pena"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdTELEP 
         Caption         =   "Mandar User"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   49
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdSILENCIAR 
         Caption         =   "Silenciar"
         CausesValidation=   0   'False
         Height          =   435
         Left            =   1200
         TabIndex        =   48
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdIRA 
         Caption         =   "Ir al User"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   47
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdCARCEL 
         Caption         =   "Carcel"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   46
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdADVERTENCIA 
         Caption         =   "Advertencia"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   45
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton cmdINFO 
         Caption         =   "Informacion"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         TabIndex        =   44
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdSTAT 
         Caption         =   "Start"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   43
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdBAL 
         Caption         =   "Oro"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   42
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton cmdINV 
         Caption         =   "Inventario"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         TabIndex        =   41
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdBOV 
         Caption         =   "Boveda"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   40
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdSKILLS 
         Caption         =   "Skills"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   39
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton cmdREVIVIR 
         Caption         =   "Revivir User"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         TabIndex        =   38
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton cmdPERDON 
         Caption         =   "Perdonar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   37
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdECHAR 
         Caption         =   "Echar"
         CausesValidation=   0   'False
         Height          =   435
         Left            =   0
         TabIndex        =   36
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdEJECUTAR 
         Caption         =   "Ejecutar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   35
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton cmdBAN 
         Caption         =   "Bannear"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         TabIndex        =   34
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton cmdUNBAN 
         Caption         =   "Sacar Ban"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   33
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton cmdSUM 
         Caption         =   "Traer"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   32
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdNICK2IP 
         Caption         =   "Nick del IP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         TabIndex        =   31
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdESTUPIDO 
         Caption         =   "Estupidez al user"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         TabIndex        =   30
         Top             =   5760
         Width           =   1815
      End
      Begin VB.CommandButton cmdNOESTUPIDO 
         Caption         =   "Sacar la estupides"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   29
         Top             =   5760
         Width           =   1815
      End
      Begin VB.CommandButton cmdBORRARPENA 
         Caption         =   "Modificar condena"
         CausesValidation=   0   'False
         Height          =   435
         Left            =   2400
         TabIndex        =   28
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton cmdLASTIP 
         Caption         =   "Ultimo IP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   27
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdCONDEN 
         Caption         =   "Condenar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         TabIndex        =   26
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdRAJAR 
         Caption         =   "Sacar Faccion"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   25
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton cmdRAJARCLAN 
         Caption         =   "Dejar sin Clan"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   24
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton cmdLASTEMAIL 
         Caption         =   "Ultimo Mail"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   23
         Top             =   2880
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   8280
      Width           =   4095
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdGMSG 
         Caption         =   "Mensaje Consola"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdHORA 
         Caption         =   "Hora"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdRMSG 
         Caption         =   "     Mensaje      Rol Master"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   1320
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdREALMSG 
         Caption         =   "Mensajes a Reales"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdCAOSMSG 
         Caption         =   "Mensajes a Caos"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   1320
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdCIUMSG 
         Caption         =   "Mensajes a Ciudadanos"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   2640
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdTALKAS 
         Caption         =   "Hablar por NPC"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   2640
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdMOTDCAMBIA 
         Caption         =   "Carbiar Motd"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   1320
         TabIndex        =   4
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdSMSG 
         Caption         =   "Mensaje por Sistema"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   2640
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   1935
      Index           =   4
      Left            =   120
      TabIndex        =   22
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdSHOWNAME 
         Caption         =   "ShowName"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   63
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmdREM 
         Caption         =   "Dejae comentario"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   62
         Top             =   1560
         Width           =   3735
      End
      Begin VB.CommandButton cmdINVISIBLE 
         Caption         =   "Invisible"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   61
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdSETDESC 
         Caption         =   "Descripcion"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   60
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdNAVE 
         Caption         =   "Navegacion"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   59
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmdCHATCOLOR 
         Caption         =   "ChatColor"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   58
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdIGNORADO 
         Caption         =   "Ignorado"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   57
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   1815
      Index           =   6
      Left            =   120
      TabIndex        =   56
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdSHOWCMSG 
         Caption         =   "Escuchar a Clan"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   80
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdBANCLAN 
         Caption         =   "/Bannea al Clan"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   79
         Top             =   1440
         Width           =   3495
      End
      Begin VB.CommandButton cmdMIEMBROSCLAN 
         Caption         =   "Mienbros del Clan"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   78
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdBANIPRELOAD 
         Caption         =   "/BanIPreload"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   77
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton cmdBANIPLIST 
         Caption         =   "/BabIPlist"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   76
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdIP2NICK 
         Caption         =   "Ban X IP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   75
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdBANIP 
         Caption         =   "Ban IP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   74
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdUNBANIP 
         Caption         =   "Sacar Ban IP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   73
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   7695
      Index           =   5
      Left            =   120
      TabIndex        =   55
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton Command6 
         Caption         =   "Mapa sin Invocación"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   128
         Top             =   7200
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cambiar Triggers"
         Height          =   315
         Left            =   2040
         TabIndex        =   127
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton cndPK1 
         Caption         =   "Mapa Inseguro"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   126
         Top             =   5280
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Mapa prohibe Robar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   125
         Top             =   6720
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Mapa con Magia"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   124
         Top             =   6240
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Mapa sin Backup"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   123
         Top             =   5760
         Width           =   1695
      End
      Begin VB.CommandButton cmdBacup 
         Caption         =   "Mapa con BackUp"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   122
         Top             =   5760
         Width           =   1695
      End
      Begin VB.CommandButton cmdMagiaNO 
         Caption         =   "Maspa sin Magia"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   121
         Top             =   6240
         Width           =   1695
      End
      Begin VB.CommandButton cmdRoboSi 
         Caption         =   "Mapa perimite Robar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   120
         Top             =   6720
         Width           =   1695
      End
      Begin VB.CommandButton cmdInvocaSI 
         Caption         =   "Mapa con Invocación"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   119
         Top             =   7200
         Width           =   1695
      End
      Begin VB.CommandButton cmdPK0 
         Caption         =   "Mapa Seguro"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   117
         Top             =   5280
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Crear NPCs sin Respawn"
         Height          =   435
         Left            =   240
         TabIndex        =   116
         Top             =   4200
         Width           =   1695
      End
      Begin VB.CommandButton cmdNPCsConRespawn 
         Caption         =   "Crear NPC con Respawn"
         Height          =   435
         Left            =   240
         TabIndex        =   115
         Top             =   3720
         Width           =   1695
      End
      Begin VB.CommandButton cmdResetInv 
         Caption         =   "Resetear Inventario"
         Height          =   315
         Left            =   240
         TabIndex        =   114
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdMataconRepawn 
         Caption         =   "      Matar criatura       deja respawn"
         Height          =   435
         Left            =   2040
         TabIndex        =   113
         Top             =   3720
         Width           =   1695
      End
      Begin VB.CommandButton cmdMata 
         Caption         =   "Matar criatura"
         Height          =   435
         Left            =   2040
         TabIndex        =   112
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton cmddestBloq 
         Caption         =   "Quitar/Poner Bloqueo"
         Height          =   315
         Left            =   240
         TabIndex        =   111
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton cmdDE 
         Caption         =   "Destruir exit"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   110
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdCC 
         Caption         =   "Crear NPCs"
         CausesValidation=   0   'False
         Height          =   435
         Left            =   240
         TabIndex        =   72
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton cmdLIMPIAR 
         Caption         =   "Limpiar Mundo"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   71
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdCT 
         Caption         =   "Crear Telepor"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   70
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdDT 
         Caption         =   "Destruir Teleport"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   69
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdLLUVIA 
         Caption         =   "Lluvia - Si / No"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   68
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdMASSDEST 
         Caption         =   "Dest Item en Mapa"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   67
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton cmdPISO 
         Caption         =   "Informe del Piso"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   66
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton cmdCI 
         Caption         =   "Crear Item"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   65
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton cmdDEST 
         Caption         =   "Destruir item"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   64
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   3720
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label lblMapa 
         Caption         =   "Modificar opciones del Mapa"
         Height          =   375
         Left            =   840
         TabIndex        =   118
         Top             =   4920
         Width           =   2175
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   3600
         Y1              =   4800
         Y2              =   4800
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   7395
      Index           =   7
      Left            =   120
      TabIndex        =   88
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Actualizar"
         Height          =   495
         Left            =   2160
         TabIndex        =   109
         Top             =   2100
         Width           =   1695
      End
      Begin VB.TextBox txtNuevaDescrip 
         Height          =   765
         Left            =   120
         MaxLength       =   40
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   107
         Top             =   6120
         Width           =   3735
      End
      Begin VB.CommandButton cmdAddFollow 
         Caption         =   "Agregar Seguimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   105
         Top             =   6960
         Width           =   3735
      End
      Begin VB.TextBox txtNuevoUsuario 
         Height          =   285
         Left            =   120
         TabIndex        =   104
         Top             =   5580
         Width           =   3735
      End
      Begin VB.CommandButton cmdAddObs 
         Caption         =   "Agregar Observacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   102
         Top             =   4800
         Width           =   3735
      End
      Begin VB.TextBox txtObs 
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   101
         Top             =   3780
         Width           =   3735
      End
      Begin VB.TextBox txtDescrip 
         Height          =   675
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   40
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   99
         Top             =   2880
         Width           =   3735
      End
      Begin VB.TextBox txtCreador 
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   97
         Top             =   1620
         Width           =   1695
      End
      Begin VB.TextBox txtTimeOn 
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   95
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtIP 
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   93
         Top             =   540
         Width           =   1695
      End
      Begin VB.ListBox lstUsers 
         Height          =   2400
         Left            =   120
         TabIndex        =   89
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         TabIndex        =   108
         Top             =   60
         Width           =   660
      End
      Begin VB.Label Label9 
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   106
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   4200
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Label Label8 
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   103
         Top             =   5340
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   100
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   98
         Top             =   2700
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Creador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   96
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Logueado Hace:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   94
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "IP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   92
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblEstado 
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Online"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   2880
         TabIndex        =   91
         Top             =   60
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Usuarios Marcados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   2415
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdONLINEREAL 
         Caption         =   "Reales Online"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   21
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdONLINECAOS 
         Caption         =   "Caos Online"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   20
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdNENE 
         Caption         =   "NPCs en Mapa"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   19
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdOCULTANDO 
         Caption         =   "Ocultos"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   16
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdONLINEGM 
         Caption         =   "GMs Online"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   15
         Top             =   2040
         Width           =   3495
      End
      Begin VB.CommandButton cmdONLINEMAP 
         Caption         =   "User en Mapa"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   14
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdBORRAR_SOS 
         Caption         =   "Borrar S:O:S"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdTRABAJANDO 
         Caption         =   "Trabajando"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdSHOW_SOS 
         Caption         =   "Ver S.O.S"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   18
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSComctlLib.TabStrip TabStrip 
      CausesValidation=   0   'False
      Height          =   1935
      Left            =   120
      TabIndex        =   81
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3413
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   7
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Message"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Info"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Player"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Me"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "World"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Admin"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seguimientos"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuSeguimientos 
      Caption         =   "Seguimientos"
      Begin VB.Menu mnuIra 
         Caption         =   "Ir Cerca"
      End
      Begin VB.Menu mnuSum 
         Caption         =   "Sumonear"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Eliminar Seguimiento"
      End
   End
End
Attribute VB_Name = "frmPanelGm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboListaUsus_Validate(Cancel As Boolean)
    Cancel = True
End Sub

Private Sub cmdACEPTCONSE_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        If MsgBox("Seguro desea aceptar a " & Nick & " como consejero real?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteAcceptRoyalCouncilMember(Nick)
            frmMain.Show
End Sub

Private Sub cmdACEPTCONSECAOS_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        If MsgBox("Seguro desea aceptar a " & Nick & " como consejero del caos?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteAcceptChaosCouncilMember(Nick)
            frmMain.Show
End Sub

Private Sub cmdAddFollow_Click()
Dim i As Long
    For i = 0 To lstUsers.ListCount
        If UCase$(lstUsers.List(i)) = UCase$(txtNuevoUsuario.Text) Then
            Call MsgBox("El usuario ya esta en la lista!", vbOKOnly + vbExclamation)
            Exit Sub
        End If
    Next i
    If LenB(txtNuevoUsuario.Text) = 0 Then
        Call MsgBox("Escribe el nombre de un usuario!", vbOKOnly + vbExclamation)
        Exit Sub
    End If
    If LenB(txtNuevaDescrip.Text) = 0 Then
        Call MsgBox("Escribe el motivo del seguimiento!", vbOKOnly + vbExclamation)
        Exit Sub
    End If
    Call WriteRecordAdd(txtNuevoUsuario.Text, txtNuevaDescrip.Text)
    txtNuevoUsuario.Text = vbNullString
    txtNuevaDescrip.Text = vbNullString
End Sub

Private Sub cmdAddObs_Click()
    Dim Obs As String
    Obs = InputBox("Ingrese la observacion", "Nueva Observacion")
    If LenB(Obs) = 0 Then
        Call MsgBox("Escribe una observacion!", vbOKOnly + vbExclamation)
        Exit Sub
    End If
    If lstUsers.ListIndex = -1 Then
        Call MsgBox("Seleccione un seguimiento!", vbOKOnly + vbExclamation)
        Exit Sub
    End If
    Call WriteRecordAddObs(lstUsers.ListIndex + 1, Obs)
End Sub

Private Sub cmdADVERTENCIA_Click()
    Dim tStr As String
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Escriba el motivo de la advertencia.", "Advertir a " & Nick)
        If LenB(tStr) <> 0 Then
            Call ParseUserCommand("/ADVERTENCIA " & Nick & "@" & tStr)
        End If
    End If
    frmMain.Show
End Sub

Private Sub cmdBackUPNo_Click()
    If MsgBox("Seguro desea hacer que el Mapa NO tenga Backup hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then _
        Call ParseUserCommand("/MODMAPINFO BACKUP 0")
End Sub

Private Sub cmdBacup_Click()
    If MsgBox("Seguro desea hacer que el Mapa tenga Backup hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then _
        Call ParseUserCommand("/MODMAPINFO BACKUP 1")
        frmMain.Show
End Sub

Private Sub cmdBAL_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        Call WriteRequestCharGold(Nick)
        frmMain.Show
End Sub

Private Sub cmdBAN_Click()
    Dim tStr As String
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Escriba el motivo del ban.", "BAN a " & Nick)
        If LenB(tStr) <> 0 Then _
            If MsgBox("Seguro desea banear a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then _
                Call WriteBanChar(Nick, tStr)
    End If
    frmMain.Show
End Sub

Private Sub cmdBANCLAN_Click()
    Dim tStr As String
    tStr = InputBox("Escriba el nombre del clan.", "Banear clan")
    If LenB(tStr) <> 0 Then _
        If MsgBox("Seguro desea banear al clan " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteGuildBan(tStr)
            frmMain.Show
End Sub

Private Sub cmdBANIP_Click()
    Dim tStr As String
    Dim Reason As String
    tStr = InputBox("Escriba el ip o el nick del PJ.", "Banear IP")
    Reason = InputBox("Escriba el motivo del ban.", "Banear IP")
    If LenB(tStr) <> 0 Then _
        If MsgBox("Seguro desea banear la ip " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then _
            Call ParseUserCommand("/BANIP " & tStr & " " & Reason)
            frmMain.Show
End Sub

Private Sub cmdBANIPLIST_Click()
    Call WriteBannedIPList
    frmMain.Show
End Sub

Private Sub cmdBANIPRELOAD_Click()
    Call WriteBannedIPReload
    frmMain.Show
End Sub

Private Sub cmdBORRAR_SOS_Click()
    If MsgBox("Seguro desea borrar el SOS?", vbYesNo, "Atencion!") = vbYes Then _
        Call WriteCleanSOS
        frmMain.Show
End Sub

Private Sub cmdBORRARPENA_Click()
    Dim tStr As String
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Indique el numero de la pena a borrar.", "Borrar pena")
        If LenB(tStr) <> 0 Then _
            If MsgBox("Seguro desea borrar la pena " & tStr & " a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then _
                Call ParseUserCommand("/BORRARPENA " & Nick & "@" & tStr)
    End If
    frmMain.Show
End Sub

Private Sub cmdBOV_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        Call WriteRequestCharBank(Nick)
        frmMain.Show
End Sub

Private Sub cmdCAOSMSG_Click()
    Dim tStr As String
    tStr = InputBox(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"), "Mensaje por consola LegionOscura")
    If LenB(tStr) <> 0 Then _
        Call WriteChaosLegionMessage(tStr)
End Sub

Private Sub cmdCARCEL_Click()
    Dim tStr As String
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Escriba el motivo de la pena.", "Carcel a " & Nick)
        If LenB(tStr) <> 0 Then
            tStr = tStr & "@" & InputBox("Indique el tiempo de condena (entre 0 y 60 minutos).", "Carcel a " & Nick)
            Call ParseUserCommand("/CARCEL " & Nick & "@" & tStr)
        End If
    End If
    frmMain.Show
End Sub

Private Sub cmdCC_Click()
    Call WriteSpawnListRequest
    frmMain.Show
End Sub

Private Sub cmdCHATCOLOR_Click()
    Dim tStr As String
    tStr = InputBox("Defina el color (R G B). Deje en blanco para usar el default.", "Cambiar color del chat")
    Call ParseUserCommand("/CHATCOLOR " & tStr)
    frmMain.Show
End Sub

Private Sub cmdCI_Click()
    Dim tStr As String
    tStr = InputBox("Indique el numero del objeto a crear y la cantidad.", "Crear Objeto")
    If LenB(tStr) <> 0 Then _
        If MsgBox("Seguro desea crear el objeto " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then _
            Call ParseUserCommand("/CI " & tStr)
            frmMain.Show
End Sub

Private Sub cmdCIUMSG_Click()
    Dim tStr As String
    tStr = InputBox(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"), "Mensaje por consola Ciudadanos")
    If LenB(tStr) <> 0 Then _
        Call WriteCitizenMessage(tStr)
End Sub

Private Sub cmdCONDEN_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        If MsgBox("Seguro desea volver criminal a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteTurnCriminal(Nick)
            frmMain.Show
End Sub

Private Sub cmdConsulta_Click()
    WriteConsultation
    frmMain.Show
End Sub

Private Sub cmdCT_Click()
    Dim tStr As String
    tStr = InputBox("Indique la posicion donde lleva el portal (MAPA X Y).", "Crear Portal")
    If LenB(tStr) <> 0 Then _
        Call ParseUserCommand("/CT " & tStr)
    frmMain.Show
End Sub

Private Sub cmdDE_Click()
    If MsgBox("Seguro desea destruir el Tile Exit?", vbYesNo, "Atencion!") = vbYes Then _
        Call WriteExitDestroy
        frmMain.Show
End Sub

Private Sub cmdDEST_Click()
    If MsgBox("Seguro desea destruir el objeto sobre el que esta parado?", vbYesNo, "Atencion!") = vbYes Then _
        Call WriteDestroyItems
        frmMain.Show
End Sub

Private Sub cmddestBloq_Click()
    If MsgBox("Seguro desea el bloqueo en su ubicación ? ", vbYesNo, "Atencion!") = vbYes Then _
        Call WriteTileBlockedToggle
        frmMain.Show
End Sub

Private Sub cmdDONDE_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        Call WriteWhere(Nick)
        frmMain.Show
End Sub

Private Sub cmdDT_Click()
    If MsgBox("Seguro desea destruir el portal?", vbYesNo, "Atencion!") = vbYes Then _
        Call WriteTeleportDestroy
        Call WriteExitDestroy
        frmMain.Show
End Sub

Private Sub cmdECHAR_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        Call WriteKick(Nick)
        frmMain.Show
End Sub

Private Sub cmdEJECUTAR_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        If MsgBox("Seguro desea ejecutar a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteExecute(Nick)
            frmMain.Show
End Sub

Private Sub cmdESTUPIDO_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        Call WriteMakeDumb(Nick)
        frmMain.Show
End Sub

Private Sub cmdGMSG_Click()
    Dim tStr As String
    tStr = InputBox("Escriba el mensaje.", "Mensaje por consola de GM")
    If LenB(tStr) <> 0 Then _
        Call WriteGMMessage(tStr)
End Sub

Private Sub cmdHORA_Click()
    Call Protocol.WriteServerTime
End Sub

Private Sub cmdIGNORADO_Click()
    Call WriteIgnored
    frmMain.Show
End Sub

Private Sub cmdINFO_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        Call WriteRequestCharInfo(Nick)
        frmMain.Show
End Sub

Private Sub cmdINV_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        Call WriteRequestCharInventory(Nick)
        frmMain.Show
End Sub

Private Sub cmdINVISIBLE_Click()
    Call WriteInvisible
    frmMain.Show
End Sub

Private Sub cmdInvocaNO_Click()
    If MsgBox("Seguro desea hacer que el Mapa NO puedan Invocar hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then _
        Call ParseUserCommand("/MODMAPINFO INVOCARSINEFECTO 0")
End Sub

Private Sub cmdInvocaSI_Click()
    If MsgBox("Seguro desea hacer que el Mapa puedan Invocar hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then _
        Call ParseUserCommand("/MODMAPINFO INVOCARSINEFECTO 1")
        frmMain.Show
End Sub

Private Sub cmdIP2NICK_Click()
    Dim tStr As String
    tStr = InputBox("Escriba la ip.", "IP to Nick")
    If LenB(tStr) <> 0 Then _
        Call ParseUserCommand("/IP2NICK " & tStr)
        frmMain.Show
End Sub

Private Sub cmdIRA_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        Call WriteGoToChar(Nick)
        frmMain.Show
End Sub

Private Sub cmdIRCERCA_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        Call WriteGoNearby(Nick)
        frmMain.Show
End Sub

Private Sub cmdKICKCONSE_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        If MsgBox("Seguro desea destituir a " & Nick & " de su cargo de consejero?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteCouncilKick(Nick)
            frmMain.Show
End Sub

Private Sub cmdLASTEMAIL_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        Call WriteRequestCharMail(Nick)
        frmMain.Show
End Sub

Private Sub cmdLASTIP_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        Call WriteLastIP(Nick)
        frmMain.Show
End Sub

Private Sub cmdLIMPIAR_Click()
    Call WriteLimpiarMundo
    frmMain.Show
End Sub

Private Sub cmdLLUVIA_Click()
    Call WriteRainToggle
    frmMain.Show
End Sub

Private Sub cmdMagiaNO_Click()
    If MsgBox("Seguro desea hacer que el Mapa NO tenga Magia hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then _
        Call ParseUserCommand("/MODMAPINFO MAGIASINEFECTO 0")
        frmMain.Show
End Sub

Private Sub cmdMagiaSI_Click()
    If MsgBox("Seguro desea hacer que el Mapa tenga Magia hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then _
        Call ParseUserCommand("/MODMAPINFO MAGIASINEFECTO 1")
End Sub

Private Sub cmdMASSDEST_Click()
    If MsgBox("Seguro desea destruir todos los items a la vista?", vbYesNo, "Atencion!") = vbYes Then _
        Call WriteDestroyAllItemsInArea
        frmMain.Show
End Sub

Private Sub cmdMata_Click()
    Call WriteKillNPCNoRespawn
    frmMain.Show
End Sub

Private Sub cmdMataconRepawn_Click()
    Call WriteKillNPC
    frmMain.Show
End Sub

Private Sub cmdMIEMBROSCLAN_Click()
    Dim tStr As String
    tStr = InputBox("Escriba el nombre del clan.", "Lista de miembros del clan")
    If LenB(tStr) <> 0 Then _
        Call WriteGuildMemberList(tStr)
        frmMain.Show
End Sub

Private Sub cmdMOTDCAMBIA_Click()
    Call WriteChangeMOTD
End Sub

Private Sub cmdNAVE_Click()
    Call WriteNavigateToggle
    frmMain.Show
End Sub

Private Sub cmdNENE_Click()
    Dim tStr As String
    tStr = InputBox("Indique el mapa.", "Numero de NPCs enemigos.")
    If LenB(tStr) <> 0 Then _
        Call ParseUserCommand("/NENE " & tStr)
        frmMain.Show
End Sub

Private Sub cmdNICK2IP_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        Call WriteNickToIP(Nick)
        frmMain.Show
End Sub

Private Sub cmdNOCAOS_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        If MsgBox("Seguro desea expulsar a " & Nick & " de la legion oscura?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteChaosLegionKick(Nick)
            frmMain.Show
End Sub

Private Sub cmdNOESTUPIDO_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        Call WriteMakeDumbNoMore(Nick)
        frmMain.Show
End Sub

Private Sub cmdNOREAL_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        If MsgBox("Seguro desea expulsar a " & Nick & " de la armada real?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteRoyalArmyKick(Nick)
            frmMain.Show
End Sub

Private Sub cmdNPCsConRespawn_Click()
    Dim tStr As String
    tStr = InputBox("Indique el numero del NPC a crear.", "Crear NPC con Respawn")
    If LenB(tStr) <> 0 Then _
        If MsgBox("Seguro desea crear el NPC " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then _
            Call ParseUserCommand("/RACC " & tStr)
            frmMain.Show
End Sub

Private Sub cmdOCULTANDO_Click()
    Call WriteHiding
    frmMain.Show
End Sub

Private Sub cmdONLINECAOS_Click()
    Call WriteOnlineChaosLegion
    frmMain.Show
End Sub

Private Sub cmdONLINEGM_Click()
    Call WriteOnlineGM
    frmMain.Show
End Sub

Private Sub cmdONLINEMAP_Click()
    Call WriteOnlineMap(UserMap)
    frmMain.Show
End Sub

Private Sub cmdONLINEREAL_Click()
    Call WriteOnlineRoyalArmy
    frmMain.Show
End Sub

Private Sub cmdPENAS_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        Call WritePunishments(Nick)
        frmMain.Show
End Sub

Private Sub cmdPERDON_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        Call WriteForgive(Nick)
        frmMain.Show
End Sub

Private Sub cmdPISO_Click()
    Call WriteItemsInTheFloor
    frmMain.Show
End Sub

Private Sub cmdPK0_Click()
    If MsgBox("Seguro desea hacer el Mapa Seguro?", vbYesNo, "Atencion!") = vbYes Then _
        Call ParseUserCommand("/MODMAPINFO PK 1")
        frmMain.Show
End Sub

Private Sub cmdRAJAR_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        If MsgBox("Seguro desea resetear la faccion de " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteResetFactions(Nick)
            frmMain.Show
End Sub

Private Sub cmdRAJARCLAN_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        If MsgBox("Seguro desea expulsar a " & Nick & " de su clan?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteRemoveCharFromGuild(Nick)
            frmMain.Show
End Sub

Private Sub cmdREALMSG_Click()
    Dim tStr As String
    tStr = InputBox(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"), "Mensaje por consola ArmadaReal")
    If LenB(tStr) <> 0 Then _
        Call WriteRoyalArmyMessage(tStr)
End Sub

Private Sub cmdRefresh_Click()
    Call ClearRecordDetails
    Call WriteRecordListRequest
End Sub

Private Sub cmdREM_Click()
    Dim tStr As String
    tStr = InputBox("Escriba el comentario.", "Comentario en el logGM")
    If LenB(tStr) <> 0 Then _
        Call WriteComment(tStr)
        frmMain.Show
End Sub

Private Sub cmdResetInv_Click()
    Call WriteResetNPCInventory
    frmMain.Show
End Sub

Private Sub cmdREVIVIR_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        Call WriteReviveChar(Nick)
        frmMain.Show
End Sub

Private Sub cmdRMSG_Click()
    Dim tStr As String
    tStr = InputBox("Escriba el mensaje.", "Mensaje por consola de RoleMaster")
    If LenB(tStr) <> 0 Then _
        Call WriteServerMessage(tStr)
End Sub

Private Sub cmdRoboNO_Click()
    If MsgBox("Seguro desea hacer que el Mapa NO puedan Robar hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then _
        Call ParseUserCommand("/MODMAPINFO ROBONPC 0")
End Sub

Private Sub cmdRoboSi_Click()
    If MsgBox("Seguro desea hacer que el Mapa puedan Robar hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then _
        Call ParseUserCommand("/MODMAPINFO ROBONPC 1")
        frmMain.Show
End Sub

Private Sub cmdSETDESC_Click()
    Dim tStr As String
    tStr = InputBox("Escriba una DESC.", "Set Description")
    If LenB(tStr) <> 0 Then _
        Call WriteSetCharDescription(tStr)
        frmMain.Show
End Sub

Private Sub cmdSHOW_SOS_Click()
    Call WriteSOSShowList
    frmMain.Show
End Sub

Private Sub cmdSHOWCMSG_Click()
    Dim tStr As String
    tStr = InputBox("Escriba el nombre del clan que desea escuchar.", "Escuchar los mensajes del clan")
    If LenB(tStr) <> 0 Then _
        Call WriteShowGuildMessages(tStr)
        frmMain.Show
End Sub

Private Sub cmdSHOWNAME_Click()
    Call WriteShowName
    frmMain.Show
End Sub

Private Sub cmdSILENCIAR_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        Call WriteSilence(Nick)
        frmMain.Show
End Sub

Private Sub cmdSKILLS_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        Call WriteRequestCharSkills(Nick)
        frmMain.Show
End Sub

Private Sub cmdSMSG_Click()
    Dim tStr As String
    tStr = InputBox("Escriba el mensaje.", "Mensaje de sistema")
    If LenB(tStr) <> 0 Then _
        Call WriteSystemMessage(tStr)
End Sub

Private Sub cmdSTAT_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        Call WriteRequestCharStats(Nick)
        frmMain.Show
End Sub

Private Sub cmdSUM_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        Call WriteSummonChar(Nick)
        frmMain.Show
End Sub

Private Sub cmdTALKAS_Click()
    Dim tStr As String
    tStr = InputBox(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"), "Hablar por NPC")
    If LenB(tStr) <> 0 Then _
        Call WriteTalkAsNPC(tStr)
End Sub

Private Sub cmdTELEP_Click()
    Dim tStr As String
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Indique la posicion (MAPA X Y).", "Transportar a " & Nick)
        If LenB(tStr) <> 0 Then _
            Call ParseUserCommand("/TELEP " & Nick & " " & tStr) 'We use the Parser to control the command format
    End If
    frmMain.Show
End Sub

Private Sub cmdTRABAJANDO_Click()
    Call WriteWorking
    frmMain.Show
End Sub

Private Sub cmdUNBAN_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        If MsgBox("Seguro desea unbanear a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteUnbanChar(Nick)
            frmMain.Show
End Sub

Private Sub cmdUNBANIP_Click()
    Dim tStr As String
    tStr = InputBox("Escriba el ip.", "Unbanear IP")
    If LenB(tStr) <> 0 Then _
        If MsgBox("Seguro desea unbanear la ip " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then _
            Call ParseUserCommand("/UNBANIP " & tStr)
            frmMain.Show
End Sub


Private Sub cndPK1_Click()
    If MsgBox("Seguro desea hacer el Mapa Inseguro?", vbYesNo, "Atencion!") = vbYes Then _
        Call ParseUserCommand("/MODMAPINFO PK 0")
        frmMain.Show
End Sub

Private Sub Command1_Click()
    If MsgBox("Seguro desea hacer que el Mapa NO tenga Backup hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then _
        Call ParseUserCommand("/MODMAPINFO BACKUP 0")
        frmMain.Show
End Sub

Private Sub Command2_Click()
    Dim tStr As String
    tStr = InputBox("Indique el numero del NPC a crear.", "Crear NPC con Respawn")
    If LenB(tStr) <> 0 Then _
        If MsgBox("Seguro desea crear el NPC " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then _
            Call ParseUserCommand("/ACC " & tStr)
    frmMain.Show
End Sub



Private Sub Command3_Click()
    If MsgBox("Seguro desea hacer que el Mapa tenga Magia hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then _
        Call ParseUserCommand("/MODMAPINFO MAGIASINEFECTO 1")
        frmMain.Show
End Sub

Private Sub Command4_Click()
    If MsgBox("Seguro desea hacer que el Mapa NO puedan Robar hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then _
        Call ParseUserCommand("/MODMAPINFO ROBONPC 0")
        frmMain.Show
End Sub

Private Sub Command5_Click()
    Dim tStr As String
    
    tStr = InputBox("Indique el numero del Trigger´s a cambiar.", "Crear NPC con Respawn")
    If LenB(tStr) <> 0 Then _
        If MsgBox("Seguro desea cambiar a " & tStr & " el trigger donde esta parado?", vbYesNo, "Atencion!") = vbYes Then _
            Call ParseUserCommand("/trigger " & tStr)
    frmMain.Show
End Sub

Private Sub Command6_Click()
    If MsgBox("Seguro desea hacer que el Mapa NO puedan Invocar hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then _
        Call ParseUserCommand("/MODMAPINFO INVOCARSINEFECTO 0")
    frmMain.Show
End Sub

Private Sub Command7_Click()
    Dim Nick As String
    Nick = cboListaUsus.Text
    If LenB(Nick) <> 0 Then _
        Call WriteLookProcess(Nick)
        frmMain.Show
End Sub

Private Sub Form_Load()
    Call showTab(1)
    Call cmdActualiza_Click
    Call cmdRefresh_Click
    mnuSeguimientos.Visible = False
End Sub

Private Sub cmdActualiza_Click()
    Call WriteRequestUserList
    Call FlushBuffer
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub lstUsers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuSeguimientos
    Else
        If lstUsers.ListIndex <> -1 Then
            Call ClearRecordDetails
            Call WriteRecordDetailsRequest(lstUsers.ListIndex + 1)
        End If
    End If
End Sub

Private Sub ClearRecordDetails()
    txtIP.Text = vbNullString
    txtCreador.Text = vbNullString
    txtDescrip.Text = vbNullString
    txtObs.Text = vbNullString
    txtTimeOn.Text = vbNullString
    lblEstado.Caption = vbNullString
End Sub

Private Sub mnuDelete_Click()
    With lstUsers
        If .ListIndex = -1 Then
            Call MsgBox("Seleccione un usuario para remover el seguimiento!", vbOKOnly + vbExclamation)
            Exit Sub
        End If
        If MsgBox("Desea eliminar el seguimiento al personaje " & .List(.ListIndex) & "?", vbYesNo) = vbYes Then
            Call WriteRecordRemove(.ListIndex + 1)
            Call ClearRecordDetails
        End If
    End With
End Sub

Private Sub mnuIra_Click()
    With lstUsers
        If .ListIndex <> -1 Then
            Call WriteGoToChar(.List(.ListIndex))
        End If
    End With
End Sub

Private Sub mnuSum_Click()
    With lstUsers
        If .ListIndex <> -1 Then
            Call WriteSummonChar(.List(.ListIndex))
        End If
    End With
End Sub

Private Sub TabStrip_Click()
    Call showTab(TabStrip.SelectedItem.Index)
End Sub

Private Sub showTab(TabId As Byte)
    Dim i As Byte
    For i = 1 To Frame.UBound
        Frame(i).Visible = (i = TabId)
    Next i
    With Frame(TabId)
        frmPanelGm.Height = .Height + 1280
        TabStrip.Height = .Height + 480
        cmdCerrar.Top = .Height + 480
    End With
End Sub
