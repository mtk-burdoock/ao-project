VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online Libre"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   5.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmConnect.frx":000C
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin AOProjectClient.uAOButton btnSalir 
      Height          =   375
      Left            =   9960
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   8400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Salir"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":1DC21
      PICF            =   "frmConnect.frx":1E64B
      PICH            =   "frmConnect.frx":1F30D
      PICV            =   "frmConnect.frx":2029F
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOProjectClient.uAOButton btnRecuperar 
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Recuperar Pass"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":211A1
      PICF            =   "frmConnect.frx":21BCB
      PICH            =   "frmConnect.frx":2288D
      PICV            =   "frmConnect.frx":2381F
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOProjectClient.uAOButton btnCrearCuenta 
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Crear Cuenta"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":24721
      PICF            =   "frmConnect.frx":2514B
      PICH            =   "frmConnect.frx":25E0D
      PICV            =   "frmConnect.frx":26D9F
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
   Begin VB.Timer tEfectos 
      Left            =   1680
      Top             =   1080
   End
   Begin AOProjectClient.uAOButton btnTeclas 
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Teclas"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":27CA1
      PICF            =   "frmConnect.frx":286CB
      PICH            =   "frmConnect.frx":2938D
      PICV            =   "frmConnect.frx":2A31F
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOProjectClient.uAOButton btnConectarse 
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Conectarse"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmConnect.frx":2B221
      PICF            =   "frmConnect.frx":2BC4B
      PICH            =   "frmConnect.frx":2C90D
      PICV            =   "frmConnect.frx":2D89F
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOProjectClient.uAOCheckbox chkRecordar 
      Height          =   345
      Left            =   5280
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4680
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   609
      CHCK            =   0   'False
      ENAB            =   -1  'True
      PICC            =   "frmConnect.frx":2E7A1
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   4920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3720
      Width           =   2460
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   225
      Left            =   4905
      TabIndex        =   0
      Top             =   3210
      Width           =   2460
   End
   Begin VB.Label lblDescripcionServidor 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion Server ......."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1860
      Left            =   3720
      TabIndex        =   10
      Top             =   5520
      Width           =   4500
   End
   Begin VB.Image imgServArgentina 
      Height          =   795
      Left            =   360
      MousePointer    =   99  'Custom
      Top             =   9240
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Label version 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   555
   End
   Begin VB.Label lblRecordarme 
      BackStyle       =   0  'Transparent
      Caption         =   "Recordarme"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   4800
      Width           =   2055
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tAnimControl
    Activo As Boolean
    Velocidad As Double
    Top As Integer
End Type

Private AnimControl(1 To 11) As tAnimControl
Private Fuerza As Double
Private Lector As clsIniManager
Private Const AES_PASSWD As String = "tumamaentanga"

Private Sub btnConectarse_Click()
    AccountName = txtNombre.Text
    AccountPassword = txtPasswd.Text
    frmMain.hlst.Clear
    If Me.chkRecordar.Checked = False Then
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "Remember", "False")
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "UserName", vbNullString)
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "Password", vbNullString)
    Else
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "Remember", "True")
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "UserName", AccountName)
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "Password", Cripto.AesEncryptString(AccountPassword, AES_PASSWD))
    End If
    If CheckUserData() = True Then
        Call Protocol.Connect(E_MODO.Normal)
    End If
End Sub

Private Sub btnRecuperar_Click()
    Call Protocol.Connect(E_MODO.CambiarContrasena)
End Sub

Private Sub btnSalir_Click()
    Call CloseClient
End Sub

Private Sub btnTeclas_Click()
    Load frmKeypad
    frmKeypad.Show vbModal
    Unload frmKeypad
    txtPasswd.SetFocus
End Sub

Private Sub Form_Activate()
    If CBool(GetVar(Game.path(INIT) & "Config.ini", "LOGIN", "Remember")) = True Then
        Me.txtNombre = GetVar(Game.path(INIT) & "Config.ini", "LOGIN", "UserName")
        Me.txtPasswd = Cripto.AesDecryptString(GetVar(Game.path(INIT) & "Config.ini", "LOGIN", "Password"), AES_PASSWD)
        Me.chkRecordar.Checked = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Call CloseClient
    End If
End Sub

Private Sub Form_Load()
    EngineRun = False
    Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaConectar" & RandomNumber(1, 2) & ".jpg")
    Call LoadTextsForm
End Sub

Private Sub LoadTextsForm()
    btnConectarse.Caption = JsonLanguage.item("BTN_CONECTARSE").item("TEXTO")
    btnCrearCuenta.Caption = JsonLanguage.item("BTN_CREAR_CUENTA").item("TEXTO")
    btnRecuperar.Caption = JsonLanguage.item("BTN_RECUPERAR").item("TEXTO")
    lblRecordarme.Caption = JsonLanguage.item("LBL_RECORDARME").item("TEXTO")
    btnSalir.Caption = JsonLanguage.item("BTN_SALIR").item("TEXTO")
    btnTeclas.Caption = JsonLanguage.item("LBL_TECLAS").item("TEXTO")
End Sub

Private Sub lstRedditPosts_Click()
    Call ShellExecute(0, "Open", Posts(lstRedditPosts.ListIndex + 1).URL, "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then btnConectarse_Click
End Sub

Private Sub btnCrearCuenta_Click()
    Call Protocol.Connect(E_MODO.CrearCuenta)
End Sub
