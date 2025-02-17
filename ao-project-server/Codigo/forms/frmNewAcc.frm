VERSION 5.00
Begin VB.Form frmNewAcc 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nueva Cuenta"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2775
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
   ScaleHeight     =   2745
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCrear 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Crear"
      Height          =   360
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   990
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00808080&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   270
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1110
      Width           =   2115
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H00808080&
      Height          =   375
      Left            =   270
      TabIndex        =   0
      Top             =   360
      Width           =   2115
   End
   Begin VB.Label lblEsperando 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Esperando..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   450
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   2430
   End
   Begin VB.Label lblContraseña 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   270
      TabIndex        =   3
      Top             =   900
      Width           =   1095
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   270
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmNewAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCrear_Click()
    If SaveDataNew(txtEmail.Text, txtPassword.Text) Then
        lblEsperando.Caption = "Cuenta creada)"
    Else
        lblEsperando.Caption = "ERROR: No se pudo crear la cuenta."
    End If
End Sub

