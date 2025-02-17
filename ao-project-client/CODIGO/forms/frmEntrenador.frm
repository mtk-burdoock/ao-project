VERSION 5.00
Begin VB.Form frmEntrenador 
   BorderStyle     =   0  'None
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   245
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   281
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCriaturas 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   870
      TabIndex        =   0
      Top             =   675
      Width           =   2355
   End
   Begin AOProjectClient.uAOButton imgSalir 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TX              =   "Salir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmEntrenador.frx":0000
      PICF            =   "frmEntrenador.frx":001C
      PICH            =   "frmEntrenador.frx":0038
      PICV            =   "frmEntrenador.frx":0054
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOProjectClient.uAOButton imgLuchar 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TX              =   "Luchar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmEntrenador.frx":0070
      PICF            =   "frmEntrenador.frx":008C
      PICH            =   "frmEntrenador.frx":00A8
      PICV            =   "frmEntrenador.frx":00C4
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Con que criatura deseas combatir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmEntrenador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private Sub Form_Load()
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaEntrenador.jpg")
    Call LoadTextsForm
    Call LoadAOCustomControlsPictures(Me)
End Sub

Private Sub LoadTextsForm()
    lblTitle.Caption = JsonLanguage.item("FRM_ENTRENADOR_TITLE").item("TEXTO")
    imgLuchar.Caption = JsonLanguage.item("FRM_ENTRENADOR_LUCHAR").item("TEXTO")
    imgSalir.Caption = JsonLanguage.item("FRM_ENTRENADOR_SALIR").item("TEXTO")
End Sub

Private Sub imgLuchar_Click()
    Call WriteTrain(lstCriaturas.ListIndex + 1)
    Unload Me
End Sub

Private Sub imgSalir_Click()
    Unload Me
End Sub
