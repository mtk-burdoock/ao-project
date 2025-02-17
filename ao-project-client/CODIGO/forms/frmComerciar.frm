VERSION 5.00
Begin VB.Form frmComerciar 
   BorderStyle     =   0  'None
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox cantidad 
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
      Height          =   285
      Left            =   3150
      TabIndex        =   6
      Text            =   "1"
      Top             =   6570
      Width           =   630
   End
   Begin VB.PictureBox picInvUser 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   3945
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   5
      Top             =   1965
      Width           =   2400
   End
   Begin VB.PictureBox picInvNpc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   600
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   4
      Top             =   1965
      Width           =   2400
   End
   Begin AOProjectClient.uAOButton imgComprar 
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   6000
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      TX              =   "Comprar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmComerciar.frx":0000
      PICF            =   "frmComerciar.frx":001C
      PICH            =   "frmComerciar.frx":0038
      PICV            =   "frmComerciar.frx":0054
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOProjectClient.uAOButton imgVender 
      Height          =   495
      Left            =   3840
      TabIndex        =   8
      Top             =   6000
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      TX              =   "Vender"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmComerciar.frx":0070
      PICF            =   "frmComerciar.frx":008C
      PICH            =   "frmComerciar.frx":00A8
      PICV            =   "frmComerciar.frx":00C4
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgCross 
      Height          =   450
      Left            =   6075
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   360
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Index           =   2
      Left            =   3510
      TabIndex        =   3
      Top             =   1335
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Index           =   3
      Left            =   3510
      TabIndex        =   2
      Top             =   1050
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   1050
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   720
      Width           =   75
   End
End
Attribute VB_Name = "frmComerciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Public LasActionBuy As Boolean
Private ClickNpcInv As Boolean
Private cBotonCruz As clsGraphicalButton
Public LastButtonPressed As clsGraphicalButton

Private Sub cantidad_Change()
    If Val(cantidad.Text) < 1 Then
        cantidad.Text = 1
    End If
    If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
        cantidad.Text = MAX_INVENTORY_OBJS
    End If
    If ClickNpcInv Then
        If InvComNpc.SelectedItem <> 0 Then
            Label1(1).Caption = "$: " & CalculateSellPrice(NPCInventory(InvComNpc.SelectedItem).Valor, Val(cantidad.Text))
        End If
    Else
        If InvComUsu.SelectedItem <> 0 Then
            Label1(1).Caption = "$: " & CalculateBuyPrice(Inventario.Valor(InvComUsu.SelectedItem), Val(cantidad.Text))
        End If
    End If
End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Form_Load()
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaComercio.jpg")
    Call LoadButtons
    Call LoadTextsForm
    Call LoadAOCustomControlsPictures(Me)
End Sub

Private Sub Form_Activate()
On Error Resume Next
    Call InvComUsu.DrawInventory
    Call InvComNpc.DrawInventory
End Sub

Private Sub Form_GotFocus()
On Error Resume Next
    Call InvComUsu.DrawInventory
    Call InvComNpc.DrawInventory
End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    GrhPath = Game.path(Interfaces)
    Set LastButtonPressed = New clsGraphicalButton
    Set cBotonCruz = New clsGraphicalButton
    
    Call cBotonCruz.Initialize(imgCross, "", _
                                    GrhPath & "BotonCruzApretadaComercio.jpg", _
                                    GrhPath & "BotonCruzApretadaComercio.jpg", Me)
End Sub

Private Sub LoadTextsForm()
    imgComprar.Caption = JsonLanguage.item("FRMCOMERCIAR_COMPRAR").item("TEXTO")
    imgVender.Caption = JsonLanguage.item("FRMCOMERCIAR_VENDER").item("TEXTO")
End Sub

Private Function CalculateSellPrice(ByRef objValue As Single, ByVal objAmount As Long) As Long
    On Error GoTo ErrorHandler
    CalculateSellPrice = CCur(objValue * 1000000) / 1000000 * objAmount + 0.5
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Error: " & Err.number
End Function

Private Function CalculateBuyPrice(ByRef objValue As Single, ByVal objAmount As Long) As Long
    On Error GoTo ErrorHandler
    CalculateBuyPrice = Fix(CCur(objValue * 1000000) / 1000000 * objAmount)
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Error: " & Err.number
End Function

Private Sub imgComprar_Click()
    If InvComNpc.SelectedItem = 0 Then Exit Sub
    If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
    Call Audio.PlayWave(SND_CLICK)
    LasActionBuy = True
    If UserGLD >= CalculateSellPrice(NPCInventory(InvComNpc.SelectedItem).Valor, Val(cantidad.Text)) Then
        Call WriteCommerceBuy(InvComNpc.SelectedItem, Val(cantidad.Text))
    Else
        Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_SIN_ORO_SUFICIENTE").item("TEXTO"), 2, 51, 223, 1, 1)
        Exit Sub
    End If
    Call InvComUsu.DrawInventory
    Call InvComNpc.DrawInventory
End Sub

Private Sub imgCross_Click()
    Call WriteCommerceEnd
End Sub

Private Sub imgVender_Click()
    If InvComUsu.SelectedItem = 0 Then Exit Sub
    If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub
    Call Audio.PlayWave(SND_CLICK)
    LasActionBuy = False
    Call WriteCommerceSell(InvComUsu.SelectedItem, Val(cantidad.Text))
    Call InvComUsu.DrawInventory
    Call InvComNpc.DrawInventory
End Sub

Private Sub picInvNpc_Click()
    Dim ItemSlot As Byte
    ItemSlot = InvComNpc.SelectedItem
    If ItemSlot = 0 Then Exit Sub
    ClickNpcInv = True
    InvComUsu.DeselectItem
    Label1(0).Caption = NPCInventory(ItemSlot).Name
    Label1(1).Caption = "$: " & CalculateSellPrice(NPCInventory(ItemSlot).Valor, Val(cantidad.Text))
    If NPCInventory(ItemSlot).Amount <> 0 Then
        Select Case NPCInventory(ItemSlot).OBJType
            Case eObjType.otWeapon
                Label1(2).Caption = "Max " & JsonLanguage.item("GOLPE").item("TEXTO") & ":" & NPCInventory(ItemSlot).MaxHit
                Label1(3).Caption = "Min " & JsonLanguage.item("GOLPE").item("TEXTO") & ":" & NPCInventory(ItemSlot).MinHit
                Label1(2).Visible = True
                Label1(3).Visible = True
            Case eObjType.otArmadura, eObjType.otcasco, eObjType.otescudo
                Label1(2).Caption = "Max " & JsonLanguage.item("DEFENSA").item("TEXTO") & ":" & NPCInventory(ItemSlot).MaxDef
                Label1(3).Caption = "Min " & JsonLanguage.item("DEFENSA").item("TEXTO") & ":" & NPCInventory(ItemSlot).MinDef
                Label1(2).Visible = True
                Label1(3).Visible = True
            Case Else
                Label1(2).Visible = False
                Label1(3).Visible = False
        End Select
    Else
        Label1(2).Visible = False
        Label1(3).Visible = False
    End If
End Sub

Private Sub picInvUser_Click()
    Dim ItemSlot As Byte
    ItemSlot = InvComUsu.SelectedItem
    If ItemSlot = 0 Then Exit Sub
    ClickNpcInv = False
    InvComNpc.DeselectItem
    Label1(0).Caption = Inventario.ItemName(ItemSlot)
    Label1(1).Caption = "$: " & CalculateBuyPrice(Inventario.Valor(ItemSlot), Val(cantidad.Text))
    If Inventario.Amount(ItemSlot) <> 0 Then
        Select Case Inventario.OBJType(ItemSlot)
            Case eObjType.otWeapon, eObjType.otFlechas
                Label1(2).Caption = "Max " & JsonLanguage.item("GOLPE").item("TEXTO") & ":" & Inventario.MaxHit(ItemSlot)
                Label1(3).Caption = "Min " & JsonLanguage.item("GOLPE").item("TEXTO") & ":" & Inventario.MinHit(ItemSlot)
                Label1(2).Visible = True
                Label1(3).Visible = True
            Case eObjType.otArmadura, eObjType.otcasco, eObjType.otescudo
                Label1(2).Caption = "Max " & JsonLanguage.item("DEFENSA").item("TEXTO") & ":" & Inventario.MaxDef(ItemSlot)
                Label1(3).Caption = "Min " & JsonLanguage.item("DEFENSA").item("TEXTO") & ":" & Inventario.MinDef(ItemSlot)
                Label1(2).Visible = True
                Label1(3).Visible = True
            Case Else
                Label1(2).Visible = False
                Label1(3).Visible = False
        End Select
    Else
        Label1(2).Visible = False
        Label1(3).Visible = False
    End If
End Sub
