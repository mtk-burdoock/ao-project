VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmCargando 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum"
   ClientHeight    =   3180
   ClientLeft      =   1410
   ClientTop       =   3000
   ClientWidth     =   6585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   267.49
   ScaleMode       =   0  'User
   ScaleWidth      =   439
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar cargar 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
      Min             =   1e-4
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2895
      ScaleWidth      =   6615
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   1440
         Top             =   1200
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cargando, por favor espere..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   3
         Left            =   2160
         TabIndex        =   1
         Top             =   2280
         Width           =   2535
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " aa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   6000
      TabIndex        =   3
      Top             =   2880
      Width           =   375
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private VersionNumberMaster As String
Private VersionNumberLocal As String

Private Sub Form_Load()
    Label1(2).Caption = GetVersionOfTheServer()
    Picture1.Picture = LoadPicture(App.Path & "\logo.jpg")
    Me.VerifyIfUsingLastVersion
End Sub

Function VerifyIfUsingLastVersion()
    On Error Resume Next
    If Not (CheckIfRunningLastVersion) Then
        If MsgBox("Tu version no es la actual, Deseas ejecutar el actualizador?. - Tu version: " & VersionNumberLocal & " Ultima version: " & VersionNumberMaster & " -- Your version is not up to date, open the launcher to update? ", vbYesNo) = vbYes Then
            Call ShellExecute(Me.hWnd, "open", App.Path & "\Autoupdate.exe", "", "", 1)
            End
        End If
    End If
End Function

Private Function CheckIfRunningLastVersion() As Boolean
    Dim responseGithub As String
    Dim JsonObject     As Object
    responseGithub = Inet1.OpenURL("https://api.github.com/repos/ao-libre/ao-server/releases/latest")
    If Len(responseGithub) = 0 Then Exit Function
    Set JsonObject = JSON.parse(responseGithub)
    VersionNumberMaster = JsonObject.Item("tag_name")
    VersionNumberLocal = GetVar(App.Path & "\Server.ini", "INIT", "VersionTagRelease")
    If VersionNumberMaster = VersionNumberLocal Then
        CheckIfRunningLastVersion = True
    Else
        CheckIfRunningLastVersion = False
    End If
End Function
