VERSION 5.00
Begin VB.Form frmSpawnList 
   BorderStyle     =   0  'None
   Caption         =   "Invocar"
   ClientHeight    =   3465
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   3300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   231
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCriaturas 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2370
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2475
   End
   Begin AOProjectClient.uAOButton imgSalir 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      TX              =   "Salir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmSpawnList.frx":0000
      PICF            =   "frmSpawnList.frx":001C
      PICH            =   "frmSpawnList.frx":0038
      PICV            =   "frmSpawnList.frx":0054
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOProjectClient.uAOButton imgInvocar 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      TX              =   "Invocar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmSpawnList.frx":0070
      PICF            =   "frmSpawnList.frx":008C
      PICH            =   "frmSpawnList.frx":00A8
      PICV            =   "frmSpawnList.frx":00C4
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione Criatura"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmSpawnList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private Sub Form_Load()
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaInvocar.jpg")
    Call LoadTextsForm
    Call LoadAOCustomControlsPictures(Me)
End Sub

Private Sub LoadTextsForm()
    imgSalir.Caption = JsonLanguage.item("FRM_SPAWN_LIST_IMGSALIR").item("TEXTO")
    imgInvocar.Caption = JsonLanguage.item("FRM_SPAWN_LIST_IMGINVOCAR").item("TEXTO")
    lblTitle.Caption = JsonLanguage.item("FRM_SPAWN_LIST_LBLTITLE").item("TEXTO")
End Sub

Private Sub imgInvocar_Click()
    Call WriteSpawnCreature(lstCriaturas.ListIndex + 1)
End Sub

Private Sub imgSalir_Click()
    Unload Me
End Sub
