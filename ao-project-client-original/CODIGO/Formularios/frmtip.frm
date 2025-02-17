VERSION 5.00
Begin VB.Form frmtip 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H80000004&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      MouseIcon       =   "frmtip.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1575
      Width           =   1185
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mostrar proxima vez"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2040
      MouseIcon       =   "frmtip.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1620
      Value           =   1  'Checked
      Width           =   2340
   End
   Begin VB.Label tip 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1260
      Left            =   120
      TabIndex        =   1
      Top             =   75
      Width           =   4305
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmtip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   If frmtip.Check1.Value = vbChecked Then
      ClientSetup.MostrarTips = True
   Else
      ClientSetup.MostrarTips = False
   End If
   Call WriteVar(Game.path(INIT) & "Config.ini", "OTHER", "MOSTRAR_TIPS", IIf(ClientSetup.MostrarTips, "True", "False"))
   Unload Me
End Sub

Private Sub Form_Deactivate()
   Me.SetFocus
End Sub

Private Sub CargarTip()
    Dim qtyTips As Integer
    qtyTips = JsonTips.Count
    Dim RandomNumberTip As Integer
    RandomNumberTip = RandomNumber(1, qtyTips)
    frmtip.tip.Caption = JsonTips.item("FRM_TIP_" & RandomNumberTip)
End Sub

Private Sub Form_Load()
    Call CargarTip
    With Me
        .Command1.Caption = JsonLanguage.item("TIP").item("TEXTO").item(1)
        .Check1.Caption = JsonLanguage.item("TIP").item("TEXTO").item(2)
    End With
End Sub
