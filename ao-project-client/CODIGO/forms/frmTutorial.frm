VERSION 5.00
Begin VB.Form frmTutorial 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   583
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin AOProjectClient.uAOButton imgSiguiente 
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   6960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TX              =   "Siguiente"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmTutorial.frx":0000
      PICF            =   "frmTutorial.frx":001C
      PICH            =   "frmTutorial.frx":0038
      PICV            =   "frmTutorial.frx":0054
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
   Begin AOProjectClient.uAOButton imgAnterior 
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   6960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TX              =   "Anterior"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmTutorial.frx":0070
      PICF            =   "frmTutorial.frx":008C
      PICH            =   "frmTutorial.frx":00A8
      PICV            =   "frmTutorial.frx":00C4
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
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   390
      Left            =   525
      TabIndex        =   4
      Top             =   435
      Width           =   7725
   End
   Begin VB.Label lblMensaje 
      BackStyle       =   0  'Transparent
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
      Height          =   5790
      Left            =   525
      TabIndex        =   3
      Top             =   840
      Width           =   7725
   End
   Begin VB.Label lblPagTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
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
      Height          =   255
      Left            =   7365
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblPagActual 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
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
      Height          =   255
      Left            =   6870
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblCerrar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   8430
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   75
      Width           =   255
   End
End
Attribute VB_Name = "frmTutorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private Type tTutorial
    sTitle As String
    sPage As String
End Type

Private Tutorial() As tTutorial
Private NumPages As Long
Private CurrentPage As Long

Private Sub Form_Load()
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaTutorial.jpg")
    Call LoadTextForms
    Call LoadAOCustomControlsPictures(Me)
    Call LoadButtons
    Call LoadTutorial
    CurrentPage = 1
    Call SelectPage(CurrentPage)
End Sub

Private Sub LoadTextForms()
    imgSiguiente.Caption = JsonLanguage.item("FRM_TUTORIAL_SIGUIENTE").item("TEXTO")
    imgAnterior.Caption = JsonLanguage.item("FRM_TUTORIAL_ANTERIOR").item("TEXTO")
End Sub

Private Sub LoadButtons()
    imgAnterior.Enabled = False
    lblCerrar.MouseIcon = picMouseIcon
End Sub

Private Sub imgAnterior_Click()
    If Not imgAnterior.Enabled Then Exit Sub
    CurrentPage = CurrentPage - 1
    If CurrentPage = 1 Then imgAnterior.Enabled = False
    If Not imgSiguiente.Enabled Then imgSiguiente.Enabled = True
    Call SelectPage(CurrentPage)
End Sub

Private Sub imgSiguiente_Click()
    If Not imgSiguiente.Enabled Then Exit Sub
    CurrentPage = CurrentPage + 1
    If CurrentPage = NumPages Then imgSiguiente.Enabled = False
    If Not imgAnterior.Enabled Then imgAnterior.Enabled = True
    Call SelectPage(CurrentPage)
End Sub

Private Sub lblCerrar_Click()
    bShowTutorial = False
    Unload Me
End Sub

Private Sub LoadTutorial()
    Dim TutorialPath As String
    Dim lPage As Long
    Dim NumLines As Long
    Dim lLine As Long
    Dim sLine As String
    TutorialPath = Game.path(Extras) & "Tutorial_" & Language & ".dat"
    NumPages = Val(GetVar(TutorialPath, "INIT", "NumPags"))
    If NumPages > 0 Then
        ReDim Tutorial(1 To NumPages)
        For lPage = 1 To NumPages
            NumLines = Val(GetVar(TutorialPath, "PAG" & lPage, "NumLines"))
            With Tutorial(lPage)
                .sTitle = GetVar(TutorialPath, "PAG" & lPage, "Title")
                For lLine = 1 To NumLines
                    sLine = GetVar(TutorialPath, "PAG" & lPage, "Line" & lLine)
                    .sPage = .sPage & sLine & vbNewLine
                Next lLine
            End With
        Next lPage
    End If
    lblPagTotal.Caption = NumPages
End Sub

Private Sub SelectPage(ByVal lPage As Long)
    lblTitulo.Caption = Tutorial(lPage).sTitle
    lblMensaje.Caption = Tutorial(lPage).sPage
    lblPagActual.Caption = lPage
End Sub
