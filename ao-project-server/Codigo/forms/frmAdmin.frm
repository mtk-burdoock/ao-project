VERSION 5.00
Begin VB.Form frmAdmin 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administracion del servidor"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Personajes"
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
      Height          =   2535
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtPjInfo 
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         Height          =   495
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   720
         Width           =   3135
      End
      Begin VB.CommandButton cmdEcharTodos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Echar todos los PJS no privilegiados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1920
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
      Begin VB.ComboBox cboPjs 
         BackColor       =   &H00808080&
         Height          =   315
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   3135
      End
      Begin VB.CommandButton cmdEchar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Echar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1440
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboPjs_Change()
    Call ActualizaPjInfo
End Sub

Private Sub cboPjs_Click()
    Call ActualizaPjInfo
End Sub

Public Sub ActualizaListaPjs()
    Dim LoopC As Long
    With cboPjs
        .Clear
        For LoopC = 1 To LastUser
            If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
                If UserList(LoopC).flags.Privilegios And PlayerType.User Then
                    .AddItem UserList(LoopC).Name
                    .ItemData(.NewIndex) = LoopC
                End If
            End If
        Next LoopC
    End With
End Sub

Private Sub ActualizaPjInfo()
    Dim tIndex As Long
    tIndex = NameIndex(cboPjs.Text)
    If tIndex > 0 Then
        With UserList(tIndex)
            txtPjInfo.Text = .outgoingData.Length & " elementos en cola." & vbCrLf
        End With
    End If
End Sub

Private Sub cmdEchar_Click()
    Dim tIndex As Long
    tIndex = NameIndex(cboPjs.Text)
    If tIndex > 0 Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & UserList(tIndex).Name & " ha sido echado.", FontTypeNames.FONTTYPE_SERVER))
        Call CloseSocket(tIndex)
    End If
End Sub

Private Sub cmdEcharTodos_Click()
    Call EcharPjsNoPrivilegiados
End Sub
