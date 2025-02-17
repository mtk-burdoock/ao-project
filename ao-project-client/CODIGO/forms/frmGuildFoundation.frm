VERSION 5.00
Begin VB.Form frmGuildFoundation 
   BorderStyle     =   0  'None
   Caption         =   "Creacion de un Clan"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtClanName 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   1815
      Width           =   3345
   End
   Begin VB.TextBox txtWeb 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Width           =   3345
   End
   Begin VB.Image imgSiguiente 
      Height          =   375
      Left            =   2400
      Tag             =   "1"
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Image imgCancelar 
      Height          =   375
      Left            =   240
      Tag             =   "1"
      Top             =   3240
      Width           =   1335
   End
End
Attribute VB_Name = "frmGuildFoundation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Private cBotonSiguiente As clsGraphicalButton
Private cBotonCancelar As clsGraphicalButton
Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Deactivate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    If Language = "spanish" Then
      Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaNombreClan_spanish.jpg")
    Else
      Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaNombreClan_english.jpg")
    End If
    Call LoadButtons
    If Len(txtClanName.Text) <= 30 Then
        If Not AsciiValidos(txtClanName) Then
            MsgBox JsonLanguage.item("VALIDACION_BAD_NOMBRE_PJ").item("TEXTO").item(2)
            Exit Sub
        End If
    Else
        MsgBox JsonLanguage.item("VALIDACION_BAD_NOMBRE_PJ").item("TEXTO").item(3)
        Exit Sub
    End If
End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    GrhPath = Game.path(Interfaces)
    Set cBotonSiguiente = New clsGraphicalButton
    Set cBotonCancelar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    
    Call cBotonSiguiente.Initialize(imgSiguiente, GrhPath & "BotonSiguienteNombreClan.jpg", _
                                    GrhPath & "BotonSiguienteRolloverNombreClan.jpg", _
                                    GrhPath & "BotonSiguienteClickNombreClan.jpg", Me)

    Call cBotonCancelar.Initialize(imgCancelar, GrhPath & "BotonCancelarNombreClan.jpg", _
                                    GrhPath & "BotonCancelarRolloverNombreClan.jpg", _
                                    GrhPath & "BotonCancelarClickNombreClan.jpg", Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgCancelar_Click()
    Unload Me
End Sub

Private Sub imgSiguiente_Click()
    ClanName = txtClanName.Text
    Site = txtWeb.Text
    Unload Me
    frmGuildDetails.Show , frmMain
End Sub

Private Sub txtWeb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub
