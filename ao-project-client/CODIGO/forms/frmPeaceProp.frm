VERSION 5.00
Begin VB.Form frmPeaceProp 
   BorderStyle     =   0  'None
   Caption         =   "Ofertas de paz"
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5070
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
   ScaleHeight     =   219
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lista 
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
      Height          =   1785
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   4620
   End
   Begin AOProjectClient.uAOButton imgCerrar 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      TX              =   "Cerrar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmPeaceProp.frx":0000
      PICF            =   "frmPeaceProp.frx":001C
      PICH            =   "frmPeaceProp.frx":0038
      PICV            =   "frmPeaceProp.frx":0054
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOProjectClient.uAOButton imgDetalle 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   2640
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      TX              =   "Detalle"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmPeaceProp.frx":0070
      PICF            =   "frmPeaceProp.frx":008C
      PICH            =   "frmPeaceProp.frx":00A8
      PICV            =   "frmPeaceProp.frx":00C4
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOProjectClient.uAOButton imgAceptar 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   2640
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      TX              =   "Aceptar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmPeaceProp.frx":00E0
      PICF            =   "frmPeaceProp.frx":00FC
      PICH            =   "frmPeaceProp.frx":0118
      PICV            =   "frmPeaceProp.frx":0134
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOProjectClient.uAOButton imgRechazar 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   2640
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      TX              =   "Rechazar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmPeaceProp.frx":0150
      PICF            =   "frmPeaceProp.frx":016C
      PICH            =   "frmPeaceProp.frx":0188
      PICV            =   "frmPeaceProp.frx":01A4
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmPeaceProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Private TipoProp As TIPO_PROPUESTA

Public Enum TIPO_PROPUESTA
    ALIANZA = 1
    PAZ = 2
End Enum

Private Sub Form_Load()
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaPeacePropForm.jpg")
    Call LoadTextsForm
    Call LoadAOCustomControlsPictures(Me)
End Sub

Private Sub LoadTextsForm()
    If TipoProp = TIPO_PROPUESTA.ALIANZA Then
        lblTitle.Caption = JsonLanguage.item("FRM_PEACE_PROP_LBLTITLE_ALIANZA").item("TEXTO")
    Else
        lblTitle.Caption = JsonLanguage.item("FRM_PEACE_PROP_LBLTITLE_PAZ").item("TEXTO")
    End If
    imgCerrar.Caption = JsonLanguage.item("FRM_PEACE_PROP_IMGCERRAR").item("TEXTO")
    imgAceptar.Caption = JsonLanguage.item("FRM_PEACE_PROP_IMGACEPTAR").item("TEXTO")
    imgDetalle.Caption = JsonLanguage.item("FRM_PEACE_PROP_IMGDETALLE").item("TEXTO")
    imgRechazar.Caption = JsonLanguage.item("FRM_PEACE_PROP_IMGRECHAZAR").item("TEXTO")
End Sub

Public Property Let ProposalType(ByVal nValue As TIPO_PROPUESTA)
    TipoProp = nValue
End Property

Private Sub imgAceptar_Click()
    If TipoProp = PAZ Then
        Call WriteGuildAcceptPeace(lista.List(lista.ListIndex))
    Else
        Call WriteGuildAcceptAlliance(lista.List(lista.ListIndex))
    End If
    Me.Hide
    Unload Me
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub imgDetalle_Click()
    If TipoProp = PAZ Then
        Call WriteGuildPeaceDetails(lista.List(lista.ListIndex))
    Else
        Call WriteGuildAllianceDetails(lista.List(lista.ListIndex))
    End If
End Sub

Private Sub imgRechazar_Click()
    If TipoProp = PAZ Then
        Call WriteGuildRejectPeace(lista.List(lista.ListIndex))
    Else
        Call WriteGuildRejectAlliance(lista.List(lista.ListIndex))
    End If
    Me.Hide
    
    Unload Me
End Sub
