VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private InventoryOffset As Long
Private InvSelectedItem As Long
Private MouseDownSelectedItem As Long
Private ownTilePixelWidth As Integer
Private ownTilePixelHeight As Integer
Private startdX As Integer
Private startdY As Integer
Private prevSelItem As Long
Private ShowValue As Boolean
Private ValueFormated As Boolean
Private ImgContainer As Boolean
Private InventoryTickCount As Long
Private TempItem As Integer
Private TempSlot As Integer
Private MouseX As Integer
Private MouseY As Integer
Private DragItem As Byte
Private Drag As Byte
Private DrawBuffer As cDIBSection
Private isInventory As Boolean
Private Inventory() As Inventory
Private WithEvents InventoryWindow As PictureBox
Attribute InventoryWindow.VB_VarHelpID = -1
#If ConMenuesConextuales = 1 Then
    Dim ItemMenu As Menu
#End If

Private Sub Class_Initialize()
    Set DrawBuffer = Nothing
End Sub

Private Sub Class_Terminate()

End Sub

Public Sub ScrollInventory(ByVal Up As Boolean)
    If CanScroll(Up) Then
        InventoryOffset = InventoryOffset + IIf(Up, 1, -1)
        Call DrawInventory
    End If
End Sub

Public Function CanScroll(ByVal Up As Boolean)
    If Up Then
        CanScroll = (InventoryOffset + (InventoryWindow.ScaleHeight \ ownTilePixelHeight)) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) < UBound(Inventory)
    Else
        CanScroll = InventoryOffset > 0
    End If
End Function

Public Sub SelectGold()
    Dim prevSelectedItem As Long
    prevSelectedItem = InvSelectedItem
    InvSelectedItem = FLAGORO
    Call DrawInventory
End Sub

Public Sub DeselectItem()
    If InvSelectedItem = 0 Then Exit Sub
    Dim ClearSlot As Byte
        ClearSlot = InvSelectedItem
    InvSelectedItem = 0
    Call DrawInventory
End Sub

Public Sub ChangeSlotItemAmount(ByVal slot As Byte, ByVal Amount As Long)
    Inventory(slot).Amount = Amount
    Call DrawInventory
End Sub

Public Property Get SelectedItem() As Long
    SelectedItem = InvSelectedItem
End Property

Public Property Get MaxHit(ByVal slot As Byte) As Integer
    MaxHit = Inventory(slot).MaxHit
End Property

Public Property Get MinHit(ByVal slot As Byte) As Integer
    MinHit = Inventory(slot).MinHit
End Property

Public Property Get MaxDef(ByVal slot As Byte) As Integer
    MaxDef = Inventory(slot).MaxDef
End Property

Public Property Get MinDef(ByVal slot As Byte) As Integer
    MinDef = Inventory(slot).MinDef
End Property

Public Property Get GrhIndex(ByVal slot As Byte) As Long
    GrhIndex = Inventory(slot).GrhIndex
End Property

Public Property Get Valor(ByVal slot As Byte) As Single
    Valor = Inventory(slot).Valor
End Property

Public Property Get Amount(ByVal slot As Byte) As Long
    If slot = FLAGORO Then
        Amount = UserGLD
    ElseIf slot >= LBound(Inventory) And slot <= UBound(Inventory) Then
        Amount = Inventory(slot).Amount
    End If
End Property

Public Property Get ObjIndex(ByVal slot As Byte) As Integer
    ObjIndex = Inventory(slot).ObjIndex
End Property

Public Property Get OBJType(ByVal slot As Byte) As Integer
    OBJType = Inventory(slot).OBJType
End Property

Public Property Get ItemName(ByVal slot As Byte) As String
    ItemName = Inventory(slot).Name
End Property

Public Property Get Equipped(ByVal slot As Byte) As Boolean
    Equipped = Inventory(slot).Equipped
End Property

Public Property Get Incompatible(ByVal slot As Byte) As Boolean
    Incompatible = Inventory(slot).Incompatible
End Property

Public Property Get MaxObjs() As Byte
    MaxObjs = UBound(Inventory)
End Property

Public Sub SetMaxObjs(ByVal MaxObjs As Byte)
    If UBound(Inventory) = MaxObjs Then Exit Sub
    ReDim Preserve Inventory(1 To MaxObjs) As Inventory
End Sub

Public Sub ToggleShowValue(ByVal bShow As Boolean)
    ShowValue = bShow
    Dim slot As Integer
    Call DrawInventory
End Sub

Public Sub ClearAllSlots()
    Dim slot As Long
    For slot = 1 To Me.MaxObjs
        Call SetItem(slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, vbNullString)
    Next slot
End Sub

Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Long, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Long, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal MaxDef As Integer, ByVal MinDef As Integer, ByVal eValor As Single, ByVal eName As String, Optional ByVal eIncompatible As Boolean = False)
    If slot < 1 Or slot > UBound(Inventory) Then Exit Sub
    With Inventory(slot)
        .Amount = eAmount
        .MaxDef = MaxDef
        .MinDef = MinDef
        .Equipped = eEquipped
        .GrhIndex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .Name = eName
        .ObjIndex = eOBJIndex
        .OBJType = eObjType
        .Valor = eValor
        .Incompatible = eIncompatible
    End With
    Call DrawInventory
End Sub

Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long
    Dim TempItem As Long
    Dim temp_x   As Long
    Dim temp_y   As Long
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1
    TempSlot = TempItem
    If TempItem <= UBound(Inventory) Then
        If Inventory(TempItem).GrhIndex Then
            ClickItem = TempItem
        Else
            ClickItem = 0
        End If
    End If
End Function

Public Sub DrawInventory()
    Dim LoopC        As Long
    Dim X            As Integer
    Dim Y            As Integer
    Dim DestRect     As RECT
    Dim Inv_Color(3) As Long
    Dim Sel_Color(3) As Long
    With DestRect
        .Bottom = InventoryWindow.ScaleHeight
        .Right = InventoryWindow.ScaleWidth
    End With
    Call Engine_Long_To_RGB_List(Sel_Color(), D3DColorARGB(100, 255, 128, 0))
    Call Engine_BeginScene
    For LoopC = InventoryOffset * (InventoryWindow.ScaleWidth \ ownTilePixelWidth) + 1 To UBound(Inventory)
        startdX = ((LoopC - 1) Mod (frmMain.PicInv.Width / 32)) * 32
        startdY = ((LoopC - 1) \ (frmMain.PicInv.Width / 32)) * 32
        If isInventory Then
            If (LoopC > MAX_NORMAL_INVENTORY_SLOTS) And (LoopC <= MAX_MOCHILA_CHICA_INVENTORY_SLOTS) Then
                Call Draw_GrhIndex(GRH_SLOT_INVENTARIO_VIOLETA, startdX, startdY, 0, Normal_RGBList(), 0, False)
            ElseIf (LoopC > MAX_MOCHILA_CHICA_INVENTORY_SLOTS) And (LoopC <= MAX_INVENTORY_SLOTS) Then
                Call Draw_GrhIndex(GRH_SLOT_INVENTARIO_DORADO, startdX, startdY, 0, Normal_RGBList(), 0, False)
            Else
                Call Draw_GrhIndex(GRH_SLOT_INVENTARIO_NEGRO, startdX, startdY, 0, Normal_RGBList(), 0, False)
            End If
        Else
            Call Draw_GrhIndex(GRH_SLOT_INVENTARIO_NEGRO, startdX, startdY, 0, Normal_RGBList(), 0, False)
        End If
            
        If Inventory(LoopC).GrhIndex Then
            If LoopC <> DragItem Then
                If LoopC = SelectedItem Then
                    Call Draw_GrhIndex(GRH_SLOT_INVENTARIO_ROJO, startdX, startdY, 0, Normal_RGBList(), 0, False)
                End If
                Call Draw_GrhIndex(Inventory(LoopC).GrhIndex, startdX, startdY, 0, Normal_RGBList())
                If Inventory(LoopC).Equipped Then
                    Call DrawText(startdX + 20, startdY + 20, "E", -1)
                End If
                If Inventory(LoopC).Incompatible Then
                    Dim ColorText As Long
                    ColorText = D3DColorARGB(255, 200, 30, 30)
                    Call DrawText(startdX + 20, startdY + 20, "X", ColorText)
                End If
                Call DrawText(startdX, startdY, Inventory(LoopC).Amount, -1)
            End If
        End If
    Next LoopC
    If DragItem <> 0 And Drag = 1 And isInventory Then
        Call Draw_GrhIndex(Inventory(DragItem).GrhIndex, MouseX - 16, MouseY - 16, 0, Normal_RGBList())
        InventoryWindow.AutoRedraw = False
        Call Engine_EndScene(DestRect, InventoryWindow.hwnd)
    Else
        InventoryWindow.AutoRedraw = False
        Call Engine_EndScene(DestRect, InventoryWindow.hwnd)
        Call DrawBuffer.LoadPictureBlt(InventoryWindow.hdc)
        InventoryWindow.AutoRedraw = True
        Call DrawBuffer.PaintPicture(InventoryWindow.hdc, 0, 0, InventoryWindow.ScaleWidth, InventoryWindow.ScaleHeight, 0, 0, vbSrcCopy)
        InventoryWindow.Picture = InventoryWindow.Image
    End If
End Sub

Public Sub DrawDragAndDrop()
    If ((DragItem <> 0 And Drag = 1) And (GetTickCount() - InventoryTickCount > 40)) Then
        Call Inventario.DrawInventory
    End If
End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim temp_x   As Integer
    Dim temp_y   As Integer
    Dim TempItem As Integer
    If X < 16 Then X = 16
    If Y < 16 Then Y = 16
    If X > InventoryWindow.Width - 16 Then X = InventoryWindow.Width - 16
    If Y > InventoryWindow.Height - 16 Then Y = InventoryWindow.Height - 16
    temp_x = X \ ownTilePixelWidth
    temp_y = Y \ ownTilePixelHeight
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
    TempSlot = TempItem
    If TempItem <= UBound(Inventory) Then
        Select Case Inventory(TempItem).OBJType
            Case eObjType.otWeapon, eObjType.otFlechas
                InventoryWindow.ToolTipText = Inventory(TempItem).Name & ", Ataque: " & Inventory(TempItem).MaxHit & "/" & Inventory(TempItem).MinHit
                
            Case eObjType.otArmadura, eObjType.otcasco, eObjType.otescudo
                InventoryWindow.ToolTipText = Inventory(TempItem).Name & ", Defensa: " & Inventory(TempItem).MaxDef & "/" & Inventory(TempItem).MinDef
                
            Case Else
                InventoryWindow.ToolTipText = Inventory(TempItem).Name
        End Select
    End If
    MouseX = CInt(X)
    MouseY = CInt(Y)
End Sub

Private Sub InventoryWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And isInventory Then
        If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then Exit Sub
        Dim temp_x As Integer
        Dim temp_y As Integer
        temp_x = X \ 32
        temp_y = Y \ 32
        DragItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
        If Not (DragItem <= 25 And DragItem >= 1) Then
            DragItem = 0
            Exit Sub
        Else
            If Inventory(DragItem).GrhIndex = 0 Then
                DragItem = 0
                Exit Sub
            End If
        End If
        If isInventory Then
            Drag = 1
        End If
    End If
End Sub

Public Sub Initialize(ByRef DirectD3D As D3DX8, ByRef InvPic As PictureBox, ByVal MaxObjs As Byte, Optional ByVal FontSize As Integer = 7, Optional ByVal TileWidth As Integer = 32, Optional ByVal TileHeight As Integer = 32, Optional ByVal startX As Integer = 0, Optional ByVal startY As Integer = 0, Optional ByVal bImgContainer As Boolean = True, Optional ByVal bShowText As Boolean = True, Optional thisInventory As Boolean = False)
    Set InventoryWindow = InvPic
    isInventory = thisInventory
    InventoryWindow.AutoRedraw = True
    InventoryWindow.ScaleMode = 3
    Set DrawBuffer = New cDIBSection
    Call DrawBuffer.Create(InvPic.Width, InvPic.Height)
    ReDim Inventory(1 To MaxObjs) As Inventory
    If DirectD3D Is Nothing Then Exit Sub
    ownTilePixelHeight = TileHeight
    ownTilePixelWidth = TileWidth
    startdX = startX
    startdY = startY
    ImgContainer = bImgContainer
    ShowValue = bShowText
    If ImgContainer Then InvSelectedItem = ClickItem(1, 1)
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim prevSelItem As Long
    prevSelItem = InvSelectedItem
    InvSelectedItem = ClickItem(MouseX, MouseY)
    If Drag Then
        If Not DragItem = TempSlot And Not (TempSlot < 1 Or DragItem < 1) Then
            Call Protocol.WriteMoveItem(DragItem, TempSlot, eMoveType.Inventory)
            InvSelectedItem = TempSlot
        End If
        DragItem = 0
        Drag = 0
    End If
    Call DrawInventory
End Sub

Public Function GetPos(ByVal X As Integer, ByVal Y As Integer) As Byte
    Dim temp_x As Integer
    Dim temp_y As Integer
    temp_x = X \ 32
    temp_y = Y \ 32
    GetPos = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
End Function
