VERSION 5.00
Begin VB.Form frmGuildURL 
   BorderStyle     =   0  'None
   Caption         =   "Oficial Web Site"
   ClientHeight    =   1425
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6225
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
   ScaleHeight     =   95
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUrl 
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
      Height          =   225
      Left            =   210
      TabIndex        =   0
      Top             =   600
      Width           =   5805
   End
   Begin VB.Image imgAceptar 
      Height          =   255
      Left            =   165
      Tag             =   "1"
      Top             =   960
      Width           =   5880
   End
End
Attribute VB_Name = "frmGuildURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Private cBotonAceptar As clsGraphicalButton
Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    If Language = "spanish" Then
      Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaUrlClan_spanish.jpg")
    Else
      Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaUrlClan_english.jpg")
    End If
    Call LoadButtons
End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    GrhPath = Game.path(Interfaces)
    Set cBotonAceptar = New clsGraphicalButton
    Set LastButtonPressed = New clsGraphicalButton
    Call cBotonAceptar.Initialize(imgAceptar, GrhPath & "BotonAceptarUrl.jpg", _
                                    GrhPath & "BotonAceptaRolloverrUrl.jpg", _
                                    GrhPath & "BotonAceptarClickUrl.jpg", Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgAceptar_Click()
    If Len(txtUrl.Text) <> 0 Then _
        Call WriteGuildNewWebsite(txtUrl.Text)
    Unload Me
End Sub

Private Sub txtUrl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub
