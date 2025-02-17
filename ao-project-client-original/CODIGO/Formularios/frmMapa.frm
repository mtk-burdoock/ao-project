VERSION 5.00
Begin VB.Form frmMapa 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8595
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgToogleMap 
      Height          =   255
      Index           =   1
      Left            =   3840
      MousePointer    =   99  'Custom
      Top             =   120
      Width           =   975
   End
   Begin VB.Image imgToogleMap 
      Height          =   255
      Index           =   0
      Left            =   3960
      MousePointer    =   99  'Custom
      Top             =   7560
      Width           =   735
   End
   Begin VB.Image imgCerrar 
      Height          =   255
      Left            =   8040
      MousePointer    =   99  'Custom
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblTexto 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMapa.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   7920
      Width           =   8175
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private Enum eMaps
    ieGeneral
    ieDungeon
End Enum

Private picMaps(1) As Picture
Private CurrentMap As eMaps

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp
            ToggleImgMaps
        Case Else
            Unload Me
    End Select
    
End Sub

Private Sub ToggleImgMaps()
    imgToogleMap(CurrentMap).Visible = False
    If CurrentMap = eMaps.ieGeneral Then
        imgCerrar.Visible = False
        CurrentMap = eMaps.ieDungeon
    Else
        imgCerrar.Visible = True
        CurrentMap = eMaps.ieGeneral
    End If
    imgToogleMap(CurrentMap).Visible = True
    Me.Picture = picMaps(CurrentMap)
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    Call LoadTextsForm
    Set picMaps(eMaps.ieGeneral) = LoadPicture(Game.path(Interfaces) & "mapa1.jpg")
    Set picMaps(eMaps.ieDungeon) = LoadPicture(Game.path(Interfaces) & "mapa2.jpg")
    CurrentMap = eMaps.ieGeneral
    Me.Picture = picMaps(CurrentMap)
    imgCerrar.MouseIcon = picMouseIcon
    imgToogleMap(0).MouseIcon = picMouseIcon
    imgToogleMap(1).MouseIcon = picMouseIcon
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbInformation, JsonLanguage.item("ERROR").item("TEXTO") & ": " & Err.number
    Unload Me
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub imgToogleMap_Click(Index As Integer)
    ToggleImgMaps
End Sub

Private Sub LoadTextsForm()
    lblTexto.Caption = JsonLanguage.item("FRM_MAPA_TEXTO").item("TEXTO")
End Sub
