VERSION 5.00
Begin VB.Form frmUserRequest 
   BorderStyle     =   0  'None
   ClientHeight    =   2430
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4650
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
   ScaleHeight     =   162
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   1395
      Left            =   225
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   405
      Width           =   4185
   End
   Begin AOProjectClient.uAOButton imgCerrar 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1920
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      TX              =   "Cerrar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmUserRequest.frx":0000
      PICF            =   "frmUserRequest.frx":001C
      PICH            =   "frmUserRequest.frx":0038
      PICV            =   "frmUserRequest.frx":0054
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
      Caption         =   "Peticion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmUserRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Public Sub recievePeticion(ByVal p As String)
    Text1 = Replace$(p, "º", vbNewLine)
    Me.Show vbModeless, frmMain
End Sub

Private Sub Form_Load()
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaPeticion.jpg")
    Call LoadTextsForm
    Call LoadAOCustomControlsPictures(Me)
End Sub

Private Sub LoadTextsForm()
    imgCerrar.Caption = JsonLanguage.item("FRM_USER_REQUEST_IMGCERRAR").item("TEXTO")
    lblTitle.Caption = JsonLanguage.item("FRM_USER_REQUEST_LBLTITLE").item("TEXTO")
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub
