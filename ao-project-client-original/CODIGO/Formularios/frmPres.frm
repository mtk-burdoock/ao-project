VERSION 5.00
Begin VB.Form frmPres 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "frmPres.frx":0000
   ScaleHeight     =   9030
   ScaleWidth      =   12045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3400
      Left            =   1125
      Top             =   1200
   End
   Begin VB.Label lblSubtitle 
      BackStyle       =   0  'Transparent
      Caption         =   "El titulo a elegir en la version puede ser tan largo como queramos hasta aca"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   735
      Left            =   3360
      TabIndex        =   1
      Top             =   8160
      Width           =   8535
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Capitulo 2"
      BeginProperty Font 
         Name            =   "MS Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   975
      Left            =   3240
      TabIndex        =   0
      Top             =   7200
      Width           =   3615
   End
End
Attribute VB_Name = "frmPres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Me.Picture = LoadPicture(Game.path(Interfaces) & "ImagenPresentacion.jpg")
    lblTitle.Caption = JsonLanguage.item("FRM_PRES_LBL_TITLE").item("TEXTO")
    lblSubtitle.Caption = JsonLanguage.item("FRM_PRES_LBL_SUBTITLE").item("TEXTO")
    Me.Width = 800 * Screen.TwipsPerPixelX
    Me.Height = 600 * Screen.TwipsPerPixelY
End Sub

Private Sub Timer1_Timer()
    Static ticks As Long
    Dim PresPath As String
    ticks = ticks + 1
    If ticks = 1 Then
    Else
        Unload Me
    End If
End Sub
