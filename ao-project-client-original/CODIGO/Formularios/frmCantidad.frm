VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1470
   ClientLeft      =   1635
   ClientTop       =   4410
   ClientWidth     =   3240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   98
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   216
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCantidad 
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
      Height          =   315
      Left            =   450
      MaxLength       =   5
      TabIndex        =   0
      Top             =   450
      Width           =   2250
   End
   Begin AOProjectClient.uAOButton imgTirar 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Tirar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCantidad.frx":0000
      PICF            =   "frmCantidad.frx":001C
      PICH            =   "frmCantidad.frx":0038
      PICV            =   "frmCantidad.frx":0054
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
   Begin AOProjectClient.uAOButton imgTirarTodo 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      TX              =   "Tirar Todo"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCantidad.frx":0070
      PICF            =   "frmCantidad.frx":008C
      PICH            =   "frmCantidad.frx":00A8
      PICV            =   "frmCantidad.frx":00C4
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
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Escriba la cantidad"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   3199
   End
End
Attribute VB_Name = "frmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private Sub Form_Load()
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaTirarOro.jpg")
    Call LoadTextsForm
    Call LoadAOCustomControlsPictures(Me)
End Sub

Private Sub LoadTextsForm()
    imgTirar.Caption = JsonLanguage.item("FRM_CANTIDAD_TIRAR").item("TEXTO")
    imgTirarTodo.Caption = JsonLanguage.item("FRM_CANTIDAD_TIRAR_TODO").item("TEXTO")
    lblTitle.Caption = JsonLanguage.item("FRM_CANTIDAD_TITLE").item("TEXTO")
End Sub

Private Sub imgTirar_Click()
    If LenB(txtCantidad.Text) > 0 Then
        If Not IsNumeric(txtCantidad.Text) Then Exit Sub
        Call WriteDrop(Inventario.SelectedItem, frmCantidad.txtCantidad.Text)
        frmCantidad.txtCantidad.Text = vbNullString
    End If
    Unload Me
End Sub

Private Sub imgTirarTodo_Click()
    If Inventario.SelectedItem = 0 Then Exit Sub
    If Inventario.SelectedItem <> FLAGORO Then
        Call WriteDrop(Inventario.SelectedItem, Inventario.Amount(Inventario.SelectedItem))
        Unload Me
    Else
        If UserGLD > 10000 Then
            Call WriteDrop(Inventario.SelectedItem, 10000)
            Unload Me
        Else
            Call WriteDrop(Inventario.SelectedItem, UserGLD)
            Unload Me
        End If
    End If
    frmCantidad.txtCantidad.Text = vbNullString
End Sub

Private Sub txtCantidad_Change()
On Error GoTo ErrorHandler
    If Val(txtCantidad.Text) < 0 Then
        txtCantidad.Text = "1"
    End If
    If Val(txtCantidad.Text) > MAX_INVENTORY_OBJS Then
        txtCantidad.Text = "10000"
    End If
    Exit Sub
ErrorHandler:
    txtCantidad.Text = "1"
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub
