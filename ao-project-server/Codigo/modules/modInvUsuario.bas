Attribute VB_Name = "modInvUsuario"
#If False Then
    Dim errHandler, Userindex As Variant
#End If

Option Explicit

Public Function TieneObjetosRobables(ByVal Userindex As Integer) As Boolean
    On Error GoTo ErrorHandler
    Dim i        As Integer
    Dim ObjIndex As Integer
    For i = 1 To UserList(Userindex).CurrentInventorySlots
        ObjIndex = UserList(Userindex).Invent.Object(i).ObjIndex
        If ObjIndex > 0 Then
            If (ObjData(ObjIndex).OBJType <> eOBJType.otLlaves And ObjData(ObjIndex).OBJType <> eOBJType.otBarcos And Not ItemNewbie(ObjIndex)) Then
                TieneObjetosRobables = True
                Exit Function
            End If
        End If
    Next i
    Exit Function
ErrorHandler:
    Call LogError("Error en TieneObjetosRobables. Error: " & Err.Number & " - " & Err.description)
End Function

Function ClasePuedeUsarItem(ByVal Userindex As Integer, ByVal ObjIndex As Integer, Optional ByRef sMotivo As String) As Boolean
    On Error GoTo ErrorHandler
    If UserList(Userindex).flags.Privilegios And PlayerType.User Then
        If ObjData(ObjIndex).ClaseProhibida(1) <> 0 Then
            Dim i As Integer
            For i = 1 To NUMCLASES
                If ObjData(ObjIndex).ClaseProhibida(i) = UserList(Userindex).Clase Then
                    If ObjData(ObjIndex).OBJType = eOBJType.otPergaminos Then
                        sMotivo = "Tu clase no tiene la habilidad de aprender este hechizo."
                        ClasePuedeUsarItem = False
                        Exit Function
                    Else
                        sMotivo = "Tu clase no puede usar este objeto."
                        ClasePuedeUsarItem = False
                        Exit Function
                    End If
                End If
            Next i
        End If
    End If
    ClasePuedeUsarItem = True
    Exit Function
ErrorHandler:
    LogError ("Error en ClasePuedeUsarItem")
End Function

Public Function ItemIncompatibleConUser(ByVal Userindex As Integer, ByVal ObjIndex As Integer) As Boolean
    If ObjIndex = 0 Then
        ItemIncompatibleConUser = False
        Exit Function
    End If
    Select Case ObjData(ObjIndex).OBJType
        Case eOBJType.otWeapon, eOBJType.otAnillo, eOBJType.otFlechas, eOBJType.otEscudo
            If ClasePuedeUsarItem(Userindex, ObjIndex) And FaccionPuedeUsarItem(Userindex, ObjIndex) Then
                ItemIncompatibleConUser = False
            Else
                ItemIncompatibleConUser = True
            End If
            
        Case eOBJType.otArmadura
            If ClasePuedeUsarItem(Userindex, ObjIndex) And SexoPuedeUsarItem(Userindex, ObjIndex) And CheckRazaUsaRopa(Userindex, ObjIndex) And FaccionPuedeUsarItem(Userindex, ObjIndex) Then
                ItemIncompatibleConUser = False
            Else
                ItemIncompatibleConUser = True
            End If
            
        Case eOBJType.otCasco, eOBJType.otPergaminos
            If ClasePuedeUsarItem(Userindex, ObjIndex) Then
                ItemIncompatibleConUser = False
            Else
                ItemIncompatibleConUser = True
            End If
            
        Case Else
            ItemIncompatibleConUser = False
    End Select
End Function

Sub QuitarNewbieObj(ByVal Userindex As Integer)
    Dim j As Integer
    With UserList(Userindex)
        For j = 1 To UserList(Userindex).CurrentInventorySlots
            If .Invent.Object(j).ObjIndex > 0 Then
                If ObjData(.Invent.Object(j).ObjIndex).Newbie = 1 Then Call QuitarUserInvItem(Userindex, j, MAX_INVENTORY_OBJS)
                Call UpdateUserInv(False, Userindex, j)
            End If
        Next j
        If MapInfo(.Pos.Map).Restringir = eRestrict.restrict_newbie Then
            Dim DeDonde As tWorldPos
            Select Case .Hogar
                Case eCiudad.cLindos
                    DeDonde = Lindos

                Case eCiudad.cUllathorpe
                    DeDonde = Ullathorpe

                Case eCiudad.cBanderbill
                    DeDonde = Banderbill

                Case Else
                    DeDonde = Nix
            End Select
            Call WarpUserChar(Userindex, DeDonde.Map, DeDonde.X, DeDonde.Y, True)
        End If
    End With
End Sub

Sub LimpiarInventario(ByVal Userindex As Integer)
    Dim j As Integer
    With UserList(Userindex)
        For j = 1 To .CurrentInventorySlots
            .Invent.Object(j).ObjIndex = 0
            .Invent.Object(j).Amount = 0
            .Invent.Object(j).Equipped = 0
        Next j
        .Invent.NroItems = 0
        .Invent.ArmourEqpObjIndex = 0
        .Invent.ArmourEqpSlot = 0
        .Invent.WeaponEqpObjIndex = 0
        .Invent.WeaponEqpSlot = 0
        .Invent.CascoEqpObjIndex = 0
        .Invent.CascoEqpSlot = 0
        .Invent.EscudoEqpObjIndex = 0
        .Invent.EscudoEqpSlot = 0
        .Invent.AnilloEqpObjIndex = 0
        .Invent.AnilloEqpSlot = 0
        .Invent.MunicionEqpObjIndex = 0
        .Invent.MunicionEqpSlot = 0
        .Invent.BarcoObjIndex = 0
        .Invent.BarcoSlot = 0
        .Invent.MochilaEqpObjIndex = 0
        .Invent.MochilaEqpSlot = 0
        .Invent.MonturaObjIndex = 0
        .Invent.MonturaEqpSlot = 0
    End With
End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        If (Cantidad > 0) And (Cantidad <= .Stats.Gld) Then
            Dim MiObj As obj
            Dim loops As Integer
            If Cantidad > 50000 Then
                Dim j        As Integer
                Dim K        As Integer
                Dim M        As Integer
                Dim Cercanos As String
                M = .Pos.Map
                For j = .Pos.X - 10 To .Pos.X + 10
                    For K = .Pos.Y - 10 To .Pos.Y + 10
                        If InMapBounds(M, j, K) Then
                            If MapData(M, j, K).Userindex > 0 Then
                                Cercanos = Cercanos & UserList(MapData(M, j, K).Userindex).Name & ","
                            End If
                        End If
                    Next K
                Next j
                Call LogDesarrollo(.Name & " tira oro. Cercanos: " & Cercanos)
            End If
            Dim Extra    As Long
            Dim TeniaOro As Long
            TeniaOro = .Stats.Gld
            If Cantidad > 500000 Then
                Extra = Cantidad - 500000
                Cantidad = 500000
            End If
            Do While (Cantidad > 0)
                If Cantidad > MAX_INVENTORY_OBJS And .Stats.Gld > MAX_INVENTORY_OBJS Then
                    MiObj.Amount = MAX_INVENTORY_OBJS
                    Cantidad = Cantidad - MiObj.Amount
                Else
                    MiObj.Amount = Cantidad
                    Cantidad = Cantidad - MiObj.Amount
                End If
                MiObj.ObjIndex = iORO
                If EsGm(Userindex) Then Call LogGM(.Name, "Tiro cantidad:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name)
                Dim AuxPos As tWorldPos
                Dim EsGaleraOGaleon As Boolean
                EsGaleraOGaleon = False
                If .Invent.BarcoObjIndex <> 0 Then
                    If EsGalera(ObjData(.Invent.BarcoObjIndex)) Or EsGalera(ObjData(.Invent.BarcoObjIndex)) Then
                        EsGaleraOGaleon = True
                    End If
                End If
                If .Clase = eClass.Pirat And EsGaleraOGaleon Then
                    AuxPos = TirarItemAlPiso(.Pos, MiObj, False)
                Else
                    AuxPos = TirarItemAlPiso(.Pos, MiObj, True)
                End If
                If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                    .Stats.Gld = .Stats.Gld - MiObj.Amount
                End If
                loops = loops + 1
                If loops > 100 Then
                    LogError ("Error en tiraroro")
                    Exit Sub
                End If
            Loop
            If TeniaOro = .Stats.Gld Then Extra = 0
            If Extra > 0 Then
                .Stats.Gld = .Stats.Gld - Extra
            End If
        End If
    End With
    Exit Sub
ErrorHandler:
    Call LogError("Error en TirarOro. Error " & Err.Number & " : " & Err.description)
End Sub

Sub QuitarUserInvItem(ByVal Userindex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
    On Error GoTo ErrorHandler
    If Slot < 1 Or Slot > UserList(Userindex).CurrentInventorySlots Then Exit Sub
    With UserList(Userindex).Invent.Object(Slot)
        If .Amount <= Cantidad And .Equipped = 1 Then
            Call Desequipar(Userindex, Slot)
        End If
        .Amount = .Amount - Cantidad
        If .Amount <= 0 Then
            UserList(Userindex).Invent.NroItems = UserList(Userindex).Invent.NroItems - 1
            .ObjIndex = 0
            .Amount = 0
        End If
    End With
    Exit Sub
ErrorHandler:
    Call LogError("Error en QuitarUserInvItem. Error " & Err.Number & " : " & Err.description)
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal Userindex As Integer, ByVal Slot As Byte)
    On Error GoTo ErrorHandler
    Dim NullObj As UserObj
    Dim LoopC   As Long
    With UserList(Userindex)
        If Not UpdateAll Then
            If .Invent.Object(Slot).ObjIndex > 0 Then
                Call ChangeUserInv(Userindex, Slot, .Invent.Object(Slot))
            Else
                Call ChangeUserInv(Userindex, Slot, NullObj)
            End If
        Else
            For LoopC = 1 To .CurrentInventorySlots
                If .Invent.Object(LoopC).ObjIndex > 0 Then
                    Call ChangeUserInv(Userindex, LoopC, .Invent.Object(LoopC))
                Else
                    Call ChangeUserInv(Userindex, LoopC, NullObj)
                End If
            Next LoopC
        End If
        Exit Sub
    End With
ErrorHandler:
    Call LogError("Error en UpdateUserInv. Error " & Err.Number & " : " & Err.description)
End Sub

Sub DropObj(ByVal Userindex As Integer, ByVal Slot As Byte, ByVal Num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
    On Error GoTo ErrorHandler
    Dim DropObj As obj
    Dim MapObj  As obj
    With UserList(Userindex)
        If Num > 0 Then
            If .flags.Equitando = 1 And .Invent.MonturaEqpSlot = Slot Then
                Call WriteConsoleMsg(Userindex, "No podes tirar tu montura mientras la estas usando.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If .Invent.MochilaEqpSlot > 0 Then
                If .Invent.MochilaEqpSlot = Slot Then
                    Call WriteConsoleMsg(Userindex, "No puedes tirar tu alforja o mochila mientras la estes usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If
            DropObj.ObjIndex = .Invent.Object(Slot).ObjIndex
            If (ItemNewbie(DropObj.ObjIndex) And (.flags.Privilegios And PlayerType.User)) Then
                Call WriteConsoleMsg(Userindex, "No puedes tirar objetos newbie.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
            End If
            DropObj.Amount = MinimoInt(Num, .Invent.Object(Slot).Amount)
            MapObj.ObjIndex = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
            MapObj.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount
            If MapObj.ObjIndex = 0 Or MapObj.ObjIndex = DropObj.ObjIndex Then
                If MapObj.Amount = MAX_INVENTORY_OBJS Then
                    Call WriteConsoleMsg(Userindex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_WARNING)
                    Exit Sub
                End If
                If DropObj.Amount + MapObj.Amount > MAX_INVENTORY_OBJS Then
                    DropObj.Amount = MAX_INVENTORY_OBJS - MapObj.Amount
                End If
                Call QuitarUserInvItem(Userindex, Slot, DropObj.Amount)
                Call UpdateUserInv(False, Userindex, Slot)
                Call MakeObj(DropObj, Map, X, Y)
                If ObjData(DropObj.ObjIndex).OBJType = eOBJType.otBarcos Then
                    Call WriteConsoleMsg(Userindex, "ATENCION!! ACABAS DE TIRAR TU BARCA!", FontTypeNames.FONTTYPE_WARNING)
                End If
                If Not .flags.Privilegios And PlayerType.User Then Call LogGM(.Name, "Tiro cantidad:" & Num & " Objeto:" & ObjData(DropObj.ObjIndex).Name)
                If ObjData(DropObj.ObjIndex).Log = 1 Then
                    Call LogDesarrollo(.Name & " tiro al piso " & DropObj.Amount & " " & ObjData(DropObj.ObjIndex).Name & " Mapa: " & Map & " X: " & X & " Y: " & Y)
                ElseIf DropObj.Amount > 5000 Then
                    If ObjData(DropObj.ObjIndex).NoLog <> 1 Then
                        Call LogDesarrollo(.Name & " tiro al piso " & DropObj.Amount & " " & ObjData(DropObj.ObjIndex).Name & " Mapa: " & Map & " X: " & X & " Y: " & Y)
                    End If
                End If
            Else
                Call WriteConsoleMsg(Userindex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
    Exit Sub
ErrorHandler:
    Call LogError("Error en DropObj en " & Erl & " Nick: " & UserList(Userindex).Name & " (Map: " & UserList(Userindex).Pos.Map & "). Err: " & Err.Number & " " & Err.description)
End Sub

Sub EraseObj(ByVal Num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
    With MapData(Map, X, Y)
        .ObjInfo.Amount = .ObjInfo.Amount - Num
        If .ObjInfo.Amount <= 0 Then
            .ObjInfo.ObjIndex = 0
            .ObjInfo.Amount = 0
            Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectDelete(X, Y))
        End If
    End With
End Sub

Sub MakeObj(ByRef obj As obj, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
    On Error GoTo ErrorHandler
    If obj.ObjIndex > 0 And obj.ObjIndex <= UBound(ObjData) Then
        With MapData(Map, X, Y)
            If .ObjInfo.ObjIndex = obj.ObjIndex Then
                .ObjInfo.Amount = .ObjInfo.Amount + obj.Amount
            Else
                .ObjInfo = obj
                Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjData(obj.ObjIndex).GrhIndex, X, Y))
            End If
            Dim IsNotObjFogata As Boolean
            Dim IsNotObjTeleport As Boolean
            Dim IsNotFragua As Boolean
            Dim IsNotYacimientoPez As Boolean
            Dim IsNotYacimiento As Boolean
            Dim IsNotMueble As Boolean
            Dim IsNotArbolElfico As Boolean
            Dim IsNotArbol As Boolean
            Dim IsNotCartel As Boolean
            Dim IsNotBarco As Boolean
            Dim IsNotMontura As Boolean
            Dim IsNotYunque As Boolean
            Dim IsNotManual As Boolean
            Dim IsNotForo As Boolean
            Dim IsNotPuerta As Boolean
            Dim IsNotInstrumentos As Boolean
            Dim IsNotPergaminos As Boolean
            Dim IsNotGemas As Boolean
            Dim IsNotMochilas As Boolean
            Dim IsValidObjectToClean As Boolean
            IsNotObjFogata = ObjData(obj.ObjIndex).OBJType <> otFogata
            IsNotObjTeleport = ObjData(obj.ObjIndex).OBJType <> otTeleport
            IsNotFragua = ObjData(obj.ObjIndex).OBJType <> otFragua
            IsNotYacimientoPez = ObjData(obj.ObjIndex).OBJType <> otYacimientoPez
            IsNotYacimiento = ObjData(obj.ObjIndex).OBJType <> otYacimiento
            IsNotMueble = ObjData(obj.ObjIndex).OBJType <> otMuebles
            IsNotArbolElfico = ObjData(obj.ObjIndex).OBJType <> otArbolElfico
            IsNotArbol = ObjData(obj.ObjIndex).OBJType <> otArboles
            IsNotCartel = ObjData(obj.ObjIndex).OBJType <> otCarteles
            IsNotBarco = ObjData(obj.ObjIndex).OBJType <> otBarcos
            IsNotMontura = ObjData(obj.ObjIndex).OBJType <> otMonturas
            IsNotYunque = ObjData(obj.ObjIndex).OBJType <> otYunque
            IsNotManual = ObjData(obj.ObjIndex).OBJType <> otManuales
            IsNotForo = ObjData(obj.ObjIndex).OBJType <> otForos
            IsNotPuerta = ObjData(obj.ObjIndex).OBJType <> otPuertas
            IsNotInstrumentos = ObjData(obj.ObjIndex).OBJType <> otInstrumentos
            IsNotPergaminos = ObjData(obj.ObjIndex).OBJType <> otPergaminos
            IsNotGemas = ObjData(obj.ObjIndex).OBJType <> otGemas
            IsNotMochilas = ObjData(obj.ObjIndex).OBJType <> otMochilas
            If IsNotObjFogata And IsNotObjTeleport And IsNotFragua And IsNotYacimientoPez And IsNotYacimiento And IsNotMueble And IsNotArbolElfico And IsNotArbol And IsNotCartel And IsNotBarco And IsNotMontura And IsNotYunque And IsNotManual And IsNotForo And IsNotPuerta And IsNotInstrumentos And IsNotPergaminos And IsNotGemas And IsNotMochilas Then
                IsValidObjectToClean = True
            Else
                IsValidObjectToClean = False
            End If
            If IsValidObjectToClean And ItemNoEsDeMapa(obj.ObjIndex) Then
                Dim xPos As tWorldPos
                xPos.Map = Map
                xPos.X = X
                xPos.Y = Y
                Dim IsNotTileCasaTrigger As Boolean
                Dim IsNotTileBajoTecho As Boolean
                Dim IsNotTileBlocked As Boolean
                IsNotTileCasaTrigger = MapData(xPos.Map, xPos.X, xPos.Y).trigger <> eTrigger.CASA
                IsNotTileBajoTecho = MapData(xPos.Map, xPos.X, xPos.Y).trigger <> eTrigger.BAJOTECHO
                IsNotTileBlocked = MapData(xPos.Map, xPos.X, xPos.Y).Blocked <> 1
                If (IsNotTileCasaTrigger Or IsNotTileBajoTecho) And IsNotTileBlocked Then AgregarObjetoLimpieza xPos
            End If
        End With
    End If
    Exit Sub
ErrorHandler:
        Call LogError("Error en MakeObj en " & Erl & " Map: " & Map & "-" & X & "-" & Y & ". Err: " & Err.Number & " " & Err.description)
End Sub

Function MeterItemEnInventario(ByVal Userindex As Integer, ByRef MiObj As obj) As Boolean
    On Error GoTo ErrorHandler
    Dim Slot As Byte
    With UserList(Userindex)
        .CurrentInventorySlots = getMaxInventorySlots(Userindex)
        Slot = 1
        Do Until .Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And .Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
            Slot = Slot + 1
            If Slot > .CurrentInventorySlots Then
                Exit Do
            End If
        Loop
        If Slot > .CurrentInventorySlots Then
            Slot = 1
            Do Until .Invent.Object(Slot).ObjIndex = 0
                Slot = Slot + 1
                If Slot > .CurrentInventorySlots Then
                    Call WriteConsoleMsg(Userindex, "No puedes cargar mas objetos.", FontTypeNames.FONTTYPE_WARNING)
                    MeterItemEnInventario = False
                    Exit Function
                End If
            Loop
            .Invent.NroItems = .Invent.NroItems + 1
        End If
        If Slot > MAX_NORMAL_INVENTORY_SLOTS And Slot <= MAX_INVENTORY_SLOTS Then
            If Not ItemSeCae(MiObj.ObjIndex) Then
                Call WriteConsoleMsg(Userindex, "No puedes agarrar objetos especiales y ponerlos directamente en tu " & ObjData(.Invent.MochilaEqpObjIndex).Name & ". Puedes acomodar los items en tu inventario con drag and drop y asi si poder moverlos dentro de tu alforja. Recuerda que items que no se caen normalmente si estan dentro de la mochila caeran", FontTypeNames.FONTTYPE_WARNING)
                MeterItemEnInventario = False
                Exit Function
            End If
        End If
        If .Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
            .Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
            .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount + MiObj.Amount
        Else
            .Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
        End If
    End With
    MeterItemEnInventario = True
    Call UpdateUserInv(False, Userindex, Slot)
    Exit Function
ErrorHandler:
    Call LogError("Error en MeterItemEnInventario. Error " & Err.Number & " : " & Err.description)
End Function

Sub GetObj(ByVal Userindex As Integer)
    Dim obj    As ObjData
    Dim MiObj  As obj
    Dim ObjPos As String
    With UserList(Userindex)
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex > 0 Then
            If ObjData(MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then
                Dim X As Integer
                Dim Y As Integer
                X = .Pos.X
                Y = .Pos.Y
                obj = ObjData(MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex)
                MiObj.Amount = MapData(.Pos.Map, X, Y).ObjInfo.Amount
                MiObj.ObjIndex = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                If obj.OBJType = otOro Then
                    Dim RemainingAmountToMaximumGold As Long
                    RemainingAmountToMaximumGold = 2147483647 - .Stats.Gld
                    If Not .Stats.Gld > 2147483647 And RemainingAmountToMaximumGold >= MiObj.Amount Then
                        .Stats.Gld = .Stats.Gld + MiObj.Amount
                        Call EraseObj(MapData(.Pos.Map, X, Y).ObjInfo.Amount, .Pos.Map, .Pos.X, .Pos.Y)
                        Call WriteUpdateGold(Userindex)
                    Else
                        Call WriteConsoleMsg(Userindex, "No puedes juntar este oro por que tendrias mas del maximo disponible (2147483647)", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    If MeterItemEnInventario(Userindex, MiObj) Then
                        Call EraseObj(MapData(.Pos.Map, X, Y).ObjInfo.Amount, .Pos.Map, .Pos.X, .Pos.Y)
                        If Not .flags.Privilegios And PlayerType.User Then Call LogGM(.Name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name)
                        If ObjData(MiObj.ObjIndex).Log = 1 Then
                            ObjPos = " Mapa: " & .Pos.Map & " X: " & .Pos.X & " Y: " & .Pos.Y
                            Call LogDesarrollo(.Name & " junto del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & ObjPos)
                        ElseIf MiObj.Amount > 5000 Then
                            If ObjData(MiObj.ObjIndex).NoLog <> 1 Then
                                ObjPos = " Mapa: " & .Pos.Map & " X: " & .Pos.X & " Y: " & .Pos.Y
                                Call LogDesarrollo(.Name & " junto del piso " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name & ObjPos)
                            End If
                        End If
                    End If
                End If
            End If
        Else
            Call WriteConsoleMsg(Userindex, "No hay nada aqui.", FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Public Sub Desequipar(ByVal Userindex As Integer, ByVal Slot As Byte)
    On Error GoTo ErrorHandler
    Dim obj As ObjData
    With UserList(Userindex)
        With .Invent
            If (Slot < LBound(.Object)) Or (Slot > UBound(.Object)) Then
                Exit Sub
            ElseIf .Object(Slot).ObjIndex = 0 Then
                Exit Sub
            End If
            obj = ObjData(.Object(Slot).ObjIndex)
        End With
        Select Case obj.OBJType
            Case eOBJType.otWeapon
                With .Invent
                    .Object(Slot).Equipped = 0
                    .WeaponEqpObjIndex = 0
                    .WeaponEqpSlot = 0
                End With
                If Not .flags.Mimetizado = 1 Then
                    With .Char
                        .WeaponAnim = NingunArma
                        Call ChangeUserChar(Userindex, .body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
                    End With
                End If
            
            Case eOBJType.otFlechas
                With .Invent
                    .Object(Slot).Equipped = 0
                    .MunicionEqpObjIndex = 0
                    .MunicionEqpSlot = 0
                End With
            
            Case eOBJType.otAnillo
                With .Invent
                    .Object(Slot).Equipped = 0
                    .AnilloEqpObjIndex = 0
                    .AnilloEqpSlot = 0
                End With
            
            Case eOBJType.otArmadura
                With .Invent
                    .Object(Slot).Equipped = 0
                    .ArmourEqpObjIndex = 0
                    .ArmourEqpSlot = 0
                End With
                If Not .flags.Mimetizado = 1 And Not .flags.Navegando = 1 Then
                    Call DarCuerpoDesnudo(Userindex, .flags.Mimetizado = 1)
                    With .Char
                        Call ChangeUserChar(Userindex, .body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
                    End With
                End If
                 
            Case eOBJType.otCasco
                With .Invent
                    .Object(Slot).Equipped = 0
                    .CascoEqpObjIndex = 0
                    .CascoEqpSlot = 0
                End With
                If Not .flags.Mimetizado = 1 Or Not .flags.Navegando = 1 Then
                    With .Char
                        .CascoAnim = NingunCasco
                        Call ChangeUserChar(Userindex, .body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
                    End With
                End If
            
            Case eOBJType.otEscudo
                With .Invent
                    .Object(Slot).Equipped = 0
                    .EscudoEqpObjIndex = 0
                    .EscudoEqpSlot = 0
                End With
                If Not .flags.Mimetizado = 1 Then
                    With .Char
                        .ShieldAnim = NingunEscudo
                        Call ChangeUserChar(Userindex, .body, .Head, .heading, .WeaponAnim, .ShieldAnim, .CascoAnim)
                    End With
                End If
            
            Case eOBJType.otMochilas
                With .Invent
                    .Object(Slot).Equipped = 0
                    .MochilaEqpObjIndex = 0
                    .MochilaEqpSlot = 0
                End With
                Call modInvUsuario.TirarTodosLosItemsEnMochila(Userindex)
                .CurrentInventorySlots = MAX_NORMAL_INVENTORY_SLOTS
        End Select
    End With
    Call WriteUpdateUserStats(Userindex)
    Call UpdateUserInv(False, Userindex, Slot)
    Exit Sub
ErrorHandler:
    Call LogError("Error en Desquipar. Error " & Err.Number & " : " & Err.description)
End Sub

Function EsUsable(ByVal ObjIndex As Integer)
    Dim obj As ObjData
        obj = ObjData(ObjIndex)
    Select Case obj.OBJType
         Case eOBJType.otArbolElfico, _
              eOBJType.otArboles, _
              eOBJType.otCarteles, _
              eOBJType.otForos, _
              eOBJType.otFragua, _
              eOBJType.otMuebles, _
              eOBJType.otPuertas, _
              eOBJType.otTeleport, _
              eOBJType.otYacimiento, _
              eOBJType.otYacimientoPez, _
              eOBJType.otYunque
            EsUsable = False
        
        Case Else
            EsUsable = True
    End Select
End Function

Public Function EsBarca(ByRef Objeto As ObjData) As Boolean
    
    Select Case Objeto.Ropaje
        Case iFragataFantasmal, _
             iFragataReal, _
             iFragataCaos, _
             iBarca, _
             iBarcaCiuda, _
             iBarcaCiudaAtacable, _
             iBarcaReal, _
             iBarcaRealAtacable, _
             iBarcaPk, _
             iBarcaCaos
            EsBarca = True
            Exit Function
            
        Case EsBarca
            EsBarca = False
            Exit Function
    End Select
End Function

Public Function EsGalera(ByRef Objeto As ObjData) As Boolean
    
    Select Case Objeto.Ropaje
        Case iGalera, _
             iGaleraCiuda, _
             iGaleraCiudaAtacable, _
             iGaleraReal, _
             iGaleraRealAtacable, _
             iGaleraRealAtacable, _
             iGaleraPk, _
             iGaleraCaos
            EsGalera = True
            Exit Function
            
        Case Else
            EsGalera = False
            Exit Function
    End Select
End Function

Public Function EsGaleon(ByRef Objeto As ObjData) As Boolean
    Select Case Objeto.Ropaje
        Case iGaleon, iGaleonCiuda, iGaleonCiudaAtacable, iGaleonReal, iGaleonRealAtacable, iGaleonPk, iGaleonCaos
            EsGaleon = True
            Exit Function
            
        Case Else
            EsGaleon = False
            Exit Function
    End Select
End Function

Function SexoPuedeUsarItem(ByVal Userindex As Integer, ByVal ObjIndex As Integer, Optional ByRef sMotivo As String) As Boolean
    On Error GoTo ErrorHandler
    If ObjData(ObjIndex).Mujer = 1 Then
        SexoPuedeUsarItem = UserList(Userindex).Genero <> eGenero.Hombre
    ElseIf ObjData(ObjIndex).Hombre = 1 Then
        SexoPuedeUsarItem = UserList(Userindex).Genero <> eGenero.Mujer
    Else
        SexoPuedeUsarItem = True
    End If
    If Not SexoPuedeUsarItem Then sMotivo = "Tu genero no puede usar este objeto."
    Exit Function
ErrorHandler:
    Call LogError("SexoPuedeUsarItem")
End Function

Function FaccionPuedeUsarItem(ByVal Userindex As Integer, ByVal ObjIndex As Integer, Optional ByRef sMotivo As String) As Boolean
    If ObjData(ObjIndex).Real = 1 Then
        If Not criminal(Userindex) Then
            FaccionPuedeUsarItem = esArmada(Userindex)
        Else
            FaccionPuedeUsarItem = False
        End If
    ElseIf ObjData(ObjIndex).Caos = 1 Then
        If criminal(Userindex) Then
            FaccionPuedeUsarItem = esCaos(Userindex)
        Else
            FaccionPuedeUsarItem = False
        End If
    Else
        FaccionPuedeUsarItem = True
    End If
    If Not FaccionPuedeUsarItem Then sMotivo = "Tu alineacion no puede usar este objeto."
End Function

Sub EquiparInvItem(ByVal Userindex As Integer, ByVal Slot As Byte)
    On Error GoTo ErrorHandler
    Dim obj      As ObjData
    Dim ObjIndex As Integer
    Dim sMotivo  As String
    With UserList(Userindex)
        ObjIndex = .Invent.Object(Slot).ObjIndex
        obj = ObjData(ObjIndex)
        If Not EsUsable(ObjIndex) Then Exit Sub
        If .flags.Equitando = 1 Then
            Call WriteConsoleMsg(Userindex, "No puedes equiparte o desequiparte mientras estas en tu montura!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If obj.Newbie = 1 And Not EsNewbie(Userindex) Then
            Call WriteConsoleMsg(Userindex, "Solo los newbies pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .Stats.ELV < obj.MinLevel Then
            Call WriteConsoleMsg(Userindex, "Necesitas ser nivel " & obj.MinLevel & " para poder equipar este objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If obj.SkillRequerido Then
            If .Stats.UserSkills(obj.SkillRequerido) < obj.SkillCantidad Then
                Call WriteConsoleMsg(Userindex, "Necesitas " & obj.SkillCantidad & " puntos en " & SkillsNames(obj.SkillRequerido) & " para poder equipar este objeto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        Select Case obj.OBJType
            Case eOBJType.otWeapon
                If ClasePuedeUsarItem(Userindex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(Userindex, ObjIndex, sMotivo) Then
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(Userindex, Slot)
                        If .flags.Mimetizado = 1 Or .flags.Navegando = 1 Then
                            .CharMimetizado.WeaponAnim = NingunArma
                        Else
                            .Char.WeaponAnim = NingunArma
                            Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        End If
                        Exit Sub
                    End If
                    If .Invent.WeaponEqpObjIndex > 0 Then
                        Call Desequipar(Userindex, .Invent.WeaponEqpSlot)
                    End If
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.WeaponEqpObjIndex = ObjIndex
                    .Invent.WeaponEqpSlot = Slot
                    If Not (.flags.AdminInvisible = 1) Then Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_SACARARMA, .Pos.X, .Pos.Y))
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.WeaponAnim = GetWeaponAnim(Userindex, ObjIndex)
                    Else
                        .Char.WeaponAnim = GetWeaponAnim(Userindex, ObjIndex)
                        Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eOBJType.otAnillo
                If ClasePuedeUsarItem(Userindex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(Userindex, ObjIndex, sMotivo) Then
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(Userindex, Slot)
                        Exit Sub
                    End If
                    If .Invent.AnilloEqpObjIndex > 0 Then
                        Call Desequipar(Userindex, .Invent.AnilloEqpSlot)
                    End If
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.AnilloEqpObjIndex = ObjIndex
                    .Invent.AnilloEqpSlot = Slot
                Else
                    Call WriteConsoleMsg(Userindex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eOBJType.otFlechas
                If ClasePuedeUsarItem(Userindex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(Userindex, ObjIndex, sMotivo) Then
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(Userindex, Slot)
                        Exit Sub
                    End If
                    If .Invent.MunicionEqpObjIndex > 0 Then
                        Call Desequipar(Userindex, .Invent.MunicionEqpSlot)
                    End If
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.MunicionEqpObjIndex = ObjIndex
                    .Invent.MunicionEqpSlot = Slot
                Else
                    Call WriteConsoleMsg(Userindex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eOBJType.otArmadura
                If .flags.Equitando = 1 Then
                    Call WriteConsoleMsg(Userindex, "No podes equiparte o desequiparte vestimentas o armaduras mientras estas en tu montura.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If ClasePuedeUsarItem(Userindex, ObjIndex, sMotivo) And SexoPuedeUsarItem(Userindex, ObjIndex, sMotivo) And CheckRazaUsaRopa(Userindex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(Userindex, ObjIndex, sMotivo) Then
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(Userindex, Slot)
                        If .flags.Mimetizado = 0 And .flags.Navegando = 0 Then
                            Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        End If
                        Exit Sub
                    End If
                    If .Invent.ArmourEqpObjIndex > 0 Then
                        Call Desequipar(Userindex, .Invent.ArmourEqpSlot)
                    End If
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.ArmourEqpObjIndex = ObjIndex
                    .Invent.ArmourEqpSlot = Slot
                    If .flags.Mimetizado = 1 Or .flags.Navegando = 1 Then
                        .CharMimetizado.body = obj.Ropaje
                    Else
                        .Char.body = obj.Ropaje
                        Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                    .flags.Desnudo = 0
                Else
                    Call WriteConsoleMsg(Userindex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eOBJType.otCasco
                If ClasePuedeUsarItem(Userindex, ObjIndex, sMotivo) Then
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(Userindex, Slot)
                        If .flags.Mimetizado = 1 Or .flags.Navegando = 1 Then
                            .CharMimetizado.CascoAnim = NingunCasco
                        Else
                            .Char.CascoAnim = NingunCasco
                            Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        End If
                        Exit Sub
                    End If
                    If .Invent.CascoEqpObjIndex > 0 Then
                        Call Desequipar(Userindex, .Invent.CascoEqpSlot)
                    End If
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.CascoEqpObjIndex = ObjIndex
                    .Invent.CascoEqpSlot = Slot
                    If .flags.Mimetizado = 1 Or .flags.Navegando = 1 Then
                        .CharMimetizado.CascoAnim = obj.CascoAnim
                    Else
                        .Char.CascoAnim = obj.CascoAnim
                        Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                End If
                
            Case eOBJType.otEscudo
                If ClasePuedeUsarItem(Userindex, ObjIndex, sMotivo) And FaccionPuedeUsarItem(Userindex, ObjIndex, sMotivo) Then
                    If .Invent.Object(Slot).Equipped Then
                        Call Desequipar(Userindex, Slot)
                        If .flags.Mimetizado = 1 Or .flags.Navegando = 1 Then
                            .CharMimetizado.ShieldAnim = NingunEscudo
                        Else
                            .Char.ShieldAnim = NingunEscudo
                            Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        End If
                        Exit Sub
                    End If
                    If .Invent.EscudoEqpObjIndex > 0 Then
                        Call Desequipar(Userindex, .Invent.EscudoEqpSlot)
                    End If
                    .Invent.Object(Slot).Equipped = 1
                    .Invent.EscudoEqpObjIndex = ObjIndex
                    .Invent.EscudoEqpSlot = Slot
                    If .flags.Mimetizado = 1 Or .flags.Navegando = 1 Then
                        .CharMimetizado.ShieldAnim = obj.ShieldAnim
                    Else
                        .Char.ShieldAnim = obj.ShieldAnim
                        Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                End If
                 
            Case eOBJType.otMochilas
                If .flags.Muerto = 1 Then
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub
                End If
                If .Invent.Object(Slot).Equipped Then
                    Call Desequipar(Userindex, Slot)
                    Exit Sub
                End If
                If .Invent.MochilaEqpObjIndex > 0 Then
                    Call Desequipar(Userindex, .Invent.MochilaEqpSlot)
                End If
                .Invent.Object(Slot).Equipped = 1
                .Invent.MochilaEqpObjIndex = ObjIndex
                .Invent.MochilaEqpSlot = Slot
        End Select
    End With
    Call UpdateUserInv(False, Userindex, Slot)
    Exit Sub
ErrorHandler:
    Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & Err.Number & " - Error Description : " & Err.description)
End Sub

Private Function CheckRazaUsaRopa(ByVal Userindex As Integer, ItemIndex As Integer, Optional ByRef sMotivo As String) As Boolean
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        If .raza = eRaza.Humano Or .raza = eRaza.Elfo Or .raza = eRaza.Drow Then
            CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
        Else
            CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
        End If
        If (.raza <> eRaza.Drow) And ObjData(ItemIndex).RazaDrow Then
            CheckRazaUsaRopa = False
        End If
    End With
    If Not CheckRazaUsaRopa Then sMotivo = "Tu raza no puede usar este objeto."
    Exit Function
ErrorHandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)
End Function

Sub UseInvItem(ByVal Userindex As Integer, ByVal Slot As Byte)
    Dim obj      As ObjData
    Dim ObjIndex As Integer
    Dim TargObj  As ObjData
    Dim MiObj    As obj
    Dim sMotivo As String
    With UserList(Userindex)
        If .Invent.Object(Slot).Amount = 0 Then Exit Sub
        obj = ObjData(.Invent.Object(Slot).ObjIndex)
        If obj.Newbie = 1 And Not EsNewbie(Userindex) Then
            Call WriteConsoleMsg(Userindex, "Solo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If obj.OBJType = eOBJType.otWeapon Then
            If obj.proyectil = 1 Then
                If Not IntervaloPermiteUsar(Userindex, False) Then Exit Sub
            Else
                If Not IntervaloPermiteUsar(Userindex) Then Exit Sub
            End If
        Else
            If Not IntervaloPermiteUsar(Userindex) Then Exit Sub
        End If
        ObjIndex = .Invent.Object(Slot).ObjIndex
        .flags.TargetObjInvIndex = ObjIndex
        .flags.TargetObjInvSlot = Slot
        If Not EsUsable(ObjIndex) Then Exit Sub
        If .Stats.ELV < obj.MinLevel Then
            Call WriteConsoleMsg(Userindex, "Necesitas ser nivel " & obj.MinLevel & " para poder usar este objeto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Select Case obj.OBJType
            Case eOBJType.otUseOnce
                If .flags.Muerto = 1 Then
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub
                End If
                .Stats.MinHam = .Stats.MinHam + obj.MinHam
                If .Stats.MinHam > .Stats.MaxHam Then .Stats.MinHam = .Stats.MaxHam
                .flags.Hambre = 0
                Call WriteUpdateHungerAndThirst(Userindex)
                If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
                    Call ReproducirSonido(SendTarget.ToPCArea, Userindex, e_SoundIndex.MORFAR_MANZANA)
                Else
                    Call ReproducirSonido(SendTarget.ToPCArea, Userindex, e_SoundIndex.SOUND_COMIDA)
                End If
                Call QuitarUserInvItem(Userindex, Slot, 1)
                Call UpdateUserInv(False, Userindex, Slot)
        
            Case eOBJType.otOro
                If .flags.Muerto = 1 Then
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub
                End If
                .Stats.Gld = .Stats.Gld + .Invent.Object(Slot).Amount
                .Invent.Object(Slot).Amount = 0
                .Invent.Object(Slot).ObjIndex = 0
                .Invent.NroItems = .Invent.NroItems - 1
                Call UpdateUserInv(False, Userindex, Slot)
                Call WriteUpdateGold(Userindex)
                
            Case eOBJType.otWeapon
                If .flags.Equitando = 1 Then
                    Call WriteConsoleMsg(Userindex, "No puedes usar una herramienta mientras estas en tu montura!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If .flags.Muerto = 1 Then
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub
                End If
                If Not .Stats.MinSta > 0 Then
                    Call WriteConsoleMsg(Userindex, "Estas muy cansad" & IIf(.Genero = eGenero.Hombre, "o", "a") & ".", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If ObjData(ObjIndex).proyectil = 1 Then
                    If .Invent.Object(Slot).Equipped = 0 Then
                        Call WriteConsoleMsg(Userindex, "Antes de usar el arco deberias equipartelo.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    Call WriteMultiMessage(Userindex, eMessages.WorkRequestTarget, eSkill.Proyectiles)
                ElseIf .flags.TargetObj = Lena Then
                    If .Invent.Object(Slot).ObjIndex = DAGA Then
                        If .Invent.Object(Slot).Equipped = 0 Then
                            Call WriteConsoleMsg(Userindex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        Call TratarDeHacerFogata(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY, Userindex)
                    End If
                Else
                    Select Case ObjIndex
                        Case CANA_PESCA, RED_PESCA, CANA_PESCA_NEWBIE
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call WriteMultiMessage(Userindex, eMessages.WorkRequestTarget, eSkill.pesca)
                            Else
                                Call WriteConsoleMsg(Userindex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                            End If
                            
                        Case HACHA_LENADOR, HACHA_LENA_ELFICA, HACHA_LENADOR_NEWBIE
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call WriteMultiMessage(Userindex, eMessages.WorkRequestTarget, eSkill.Talar)
                            Else
                                Call WriteConsoleMsg(Userindex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                            End If
                            
                        Case PIQUETE_MINERO, PIQUETE_MINERO_NEWBIE
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call WriteMultiMessage(Userindex, eMessages.WorkRequestTarget, eSkill.Mineria)
                            Else
                                Call WriteConsoleMsg(Userindex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                            End If
                            
                        Case MARTILLO_HERRERO, MARTILLO_HERRERO_NEWBIE
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call WriteMultiMessage(Userindex, eMessages.WorkRequestTarget, eSkill.Herreria)
                            Else
                                Call WriteConsoleMsg(Userindex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                            End If
                            
                        Case SERRUCHO_CARPINTERO, SERRUCHO_CARPINTERO_NEWBIE
                            If .Invent.WeaponEqpObjIndex = ObjIndex Then
                                Call EnivarObjConstruibles(Userindex)
                            Else
                                Call WriteConsoleMsg(Userindex, "Debes tener equipada la herramienta para trabajar.", FontTypeNames.FONTTYPE_INFO)
                            End If
                        Case Else
                            If ObjData(ObjIndex).SkHerreria > 0 Then
                                Call WriteMultiMessage(Userindex, eMessages.WorkRequestTarget, FundirMetal)
                            End If
                    End Select
                End If
            
            Case eOBJType.otPociones
                If .flags.Muerto = 1 Then
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub
                End If
                If Not IntervaloPermiteGolpeUsar(Userindex, False) Then
                    Call WriteConsoleMsg(Userindex, "Debes esperar unos momentos para tomar otra pocion!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                .flags.TomoPocion = True
                .flags.TipoPocion = obj.TipoPocion
                Select Case .flags.TipoPocion
                    Case ePocionType.otAgilidad
                        .flags.DuracionEfecto = obj.DuracionEfecto
                        .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(obj.MinModificador, obj.MaxModificador)
                        If .Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then .Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
                        If .Stats.UserAtributos(eAtributos.Agilidad) > 2 * .Stats.UserAtributosBackUP(Agilidad) Then .Stats.UserAtributos(eAtributos.Agilidad) = 2 * .Stats.UserAtributosBackUP(Agilidad)
                        Call QuitarUserInvItem(Userindex, Slot, 1)
                        If .flags.AdminInvisible = 1 Then
                            Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        End If
                        Call WriteUpdateDexterity(Userindex)
                        
                    Case ePocionType.otFuerza
                        .flags.DuracionEfecto = obj.DuracionEfecto
                        .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(obj.MinModificador, obj.MaxModificador)
                        If .Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then .Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
                        If .Stats.UserAtributos(eAtributos.Fuerza) > 2 * .Stats.UserAtributosBackUP(Fuerza) Then .Stats.UserAtributos(eAtributos.Fuerza) = 2 * .Stats.UserAtributosBackUP(Fuerza)
                        Call QuitarUserInvItem(Userindex, Slot, 1)
                        If .flags.AdminInvisible = 1 Then
                            Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        End If
                        Call WriteUpdateStrenght(Userindex)
                        
                    Case ePocionType.otSalud
                        .Stats.MinHp = .Stats.MinHp + RandomNumber(obj.MinModificador, obj.MaxModificador)
                        If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
                        Call QuitarUserInvItem(Userindex, Slot, 1)
                        If .flags.AdminInvisible = 1 Then
                            Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        End If
                        
                    Case ePocionType.otMana
                        .Stats.MinMAN = .Stats.MinMAN + Porcentaje(.Stats.MaxMAN, 4) + .Stats.ELV \ 2 + 40 / .Stats.ELV
                        If .Stats.MinMAN > .Stats.MaxMAN Then .Stats.MinMAN = .Stats.MaxMAN
                        Call QuitarUserInvItem(Userindex, Slot, 1)
                        If .flags.AdminInvisible = 1 Then
                            Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        End If
                        
                    Case ePocionType.otCuraVeneno
                        If .flags.Envenenado = 1 Then
                            .flags.Envenenado = 0
                            Call WriteConsoleMsg(Userindex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
                        End If
                        Call QuitarUserInvItem(Userindex, Slot, 1)
                        If .flags.AdminInvisible = 1 Then
                            Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        Else
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                        End If
                        
                    Case ePocionType.otNegra
                        If .flags.SlotReto > 0 Then Exit Sub
                        If .flags.Privilegios And PlayerType.User Then
                            Call QuitarUserInvItem(Userindex, Slot, 1)
                            Call UserDie(Userindex)
                            Call WriteConsoleMsg(Userindex, "Sientes un gran mareo y pierdes el conocimiento.", FontTypeNames.FONTTYPE_FIGHT)
                        End If
                End Select
                Call WriteUpdateUserStats(Userindex)
                Call UpdateUserInv(False, Userindex, Slot)
                
            Case eOBJType.otBebidas
                If .flags.Muerto = 1 Then
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub
                End If
                .Stats.MinAGU = .Stats.MinAGU + obj.MinSed
                If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
                .flags.Sed = 0
                Call WriteUpdateHungerAndThirst(Userindex)
                Call QuitarUserInvItem(Userindex, Slot, 1)
                If .flags.AdminInvisible = 1 Then
                    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                End If
                Call UpdateUserInv(False, Userindex, Slot)
            
            Case eOBJType.otLlaves
                If .flags.Muerto = 1 Then
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub
                End If
                If .flags.TargetObj = 0 Then Exit Sub
                TargObj = ObjData(.flags.TargetObj)
                If TargObj.OBJType = eOBJType.otPuertas Then
                    If TargObj.Cerrada = 1 Then
                        If TargObj.Llave > 0 Then
                            If TargObj.Clave = obj.Clave Then
                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
                                .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                                Call WriteConsoleMsg(Userindex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            Else
                                Call WriteConsoleMsg(Userindex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                        Else
                            If TargObj.Clave = obj.Clave Then
                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
                                Call WriteConsoleMsg(Userindex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                                .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                                Exit Sub
                            Else
                                Call WriteConsoleMsg(Userindex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                        End If
                    Else
                        Call WriteConsoleMsg(Userindex, "No esta cerrada.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                End If
            
            Case eOBJType.otBotellaVacia
                If .flags.Muerto = 1 Then
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub
                End If
                If Not HayAgua(.Pos.Map, .flags.TargetX, .flags.TargetY) Then
                    Call WriteConsoleMsg(Userindex, "No hay agua alli.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                MiObj.Amount = 1
                MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexAbierta
                Call QuitarUserInvItem(Userindex, Slot, 1)
                If Not MeterItemEnInventario(Userindex, MiObj) Then
                    Call TirarItemAlPiso(.Pos, MiObj)
                End If
                Call UpdateUserInv(False, Userindex, Slot)
            
            Case eOBJType.otBotellaLlena
                If .flags.Muerto = 1 Then
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub
                End If
                .Stats.MinAGU = .Stats.MinAGU + obj.MinSed
                If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
                .flags.Sed = 0
                Call WriteUpdateHungerAndThirst(Userindex)
                MiObj.Amount = 1
                MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexCerrada
                Call QuitarUserInvItem(Userindex, Slot, 1)
                If Not MeterItemEnInventario(Userindex, MiObj) Then
                    Call TirarItemAlPiso(.Pos, MiObj)
                End If
                Call UpdateUserInv(False, Userindex, Slot)
                
            Case eOBJType.otPergaminos
                If .flags.Muerto = 1 Then
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub
                End If
                If .Stats.MaxMAN > 0 Then
                    If .flags.Hambre = 0 And .flags.Sed = 0 Then
                        If Not ClasePuedeUsarItem(Userindex, ObjIndex, sMotivo) Then
                            Call WriteConsoleMsg(Userindex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        Call AgregarHechizo(Userindex, Slot)
                        Call UpdateUserInv(False, Userindex, Slot)
                    Else
                        Call WriteConsoleMsg(Userindex, "Estas demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_INFO)
                End If

            Case eOBJType.otMinerales
                If .flags.Muerto = 1 Then
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub
                End If
                Call WriteMultiMessage(Userindex, eMessages.WorkRequestTarget, FundirMetal)
               
            Case eOBJType.otInstrumentos
                If .flags.Muerto = 1 Then
                    Call WriteMultiMessage(Userindex, eMessages.UserMuerto)
                    Exit Sub
                End If
                If obj.Real Then
                    If FaccionPuedeUsarItem(Userindex, ObjIndex) Then
                        If MapInfo(.Pos.Map).Pk = False Then
                            Call WriteConsoleMsg(Userindex, "No hay peligro aqui. Es zona segura.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        If .flags.AdminInvisible = 1 Then
                            Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                        Else
                            Call AlertarFaccionarios(Userindex)
                            Call SendData(SendTarget.toMap, .Pos.Map, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                        End If
                        Exit Sub
                    Else
                        Call WriteConsoleMsg(Userindex, "Solo miembros del ejercito real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                ElseIf obj.Caos Then
                    If FaccionPuedeUsarItem(Userindex, ObjIndex) Then
                        If MapInfo(.Pos.Map).Pk = False Then
                            Call WriteConsoleMsg(Userindex, "No hay peligro aqui. Es zona segura.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        If .flags.AdminInvisible = 1 Then
                            Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                        Else
                            Call AlertarFaccionarios(Userindex)
                            Call SendData(SendTarget.toMap, .Pos.Map, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                        End If
                        Exit Sub
                    Else
                        Call WriteConsoleMsg(Userindex, "Solo miembros de la legion oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                End If
                If .flags.AdminInvisible = 1 Then
                    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                Else
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                End If
                
            Case eOBJType.otBarcos
                If Not ClasePuedeUsarItem(Userindex, ObjIndex, sMotivo) Or Not FaccionPuedeUsarItem(Userindex, ObjIndex, sMotivo) Then
                    Call WriteConsoleMsg(Userindex, sMotivo, FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If .Stats.ELV < 25 Then
                    If .Clase <> eClass.Worker And .Clase <> eClass.Pirat Then
                        Call WriteConsoleMsg(Userindex, "Para recorrer los mares debes ser nivel 25 o superior.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    Else
                        If .Stats.ELV < 20 Then
                            If .Clase = eClass.Worker And .Stats.UserSkills(eSkill.pesca) <> 100 Then
                                Call WriteConsoleMsg(Userindex, "Para recorrer los mares debes ser nivel 20 y ademas tu skill en pesca debe ser 100.", FontTypeNames.FONTTYPE_INFO)
                            Else
                                Call WriteConsoleMsg(Userindex, "Para recorrer los mares debes ser nivel 20 o superior.", FontTypeNames.FONTTYPE_INFO)
                            End If
                            Exit Sub
                        Else
                            If .Clase = eClass.Worker Then
                                If .Stats.UserSkills(eSkill.pesca) <> 100 Then
                                    Call WriteConsoleMsg(Userindex, "Para recorrer los mares debes ser nivel 20 o superior y ademas tu skill en pesca debe ser 100.", FontTypeNames.FONTTYPE_INFO)
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
                If (LegalPos(.Pos.Map, .Pos.X - 1, .Pos.Y, True, False) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y - 1, True, False) Or LegalPos(.Pos.Map, .Pos.X + 1, .Pos.Y, True, False) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y + 1, True, False)) And .flags.Navegando = 0 Then
                    Call DoNavega(Userindex, obj, Slot)
                ElseIf (LegalPos(.Pos.Map, .Pos.X - 1, .Pos.Y, False, True) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y - 1, False, True) Or LegalPos(.Pos.Map, .Pos.X + 1, .Pos.Y, False, True) Or LegalPos(.Pos.Map, .Pos.X, .Pos.Y + 1, False, True)) And .flags.Navegando = 1 Then
                    Call DoNavega(Userindex, obj, Slot)
                Else
                    Call WriteConsoleMsg(Userindex, "Debes aproximarte al agua para usar un barco y a la tierra para desembarcar!", FontTypeNames.FONTTYPE_INFO)
                End If

            Case eOBJType.otMonturas
                If ClasePuedeUsarItem(Userindex, ObjIndex) Then
                    If .flags.invisible = 1 Then
                        Call WriteConsoleMsg(Userindex, "Estas invisible, no puedes montarte ni desmontarte en este estado!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(Userindex, "Estas muerto, no puedes montarte ni desmontarte en este estado!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    If .flags.Navegando = 1 Then
                        Call WriteConsoleMsg(Userindex, "Estas navegando, no puedes montarte ni desmontarte en este estado!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    Call DoEquita(Userindex, obj, Slot)
                Else
                    Call WriteConsoleMsg(Userindex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
                End If
                    
            Case eOBJType.otManuales
                Select Case ObjIndex
                    Case eManualType.otLiderazgo
                        If .Stats.UserSkills(eSkill.Liderazgo) < 100 Then
                            .Stats.UserSkills(eSkill.Liderazgo) = 100
                            Call WriteConsoleMsg(Userindex, "Has aprendido todo lo necesario para conformar un Clan!", FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(Userindex, "Este pergamino no tiene ning?n conocimiento que te sirva.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                    Case eManualType.otSupervivencia
                        If .Stats.UserSkills(eSkill.Supervivencia) < 100 Then
                            .Stats.UserSkills(eSkill.Supervivencia) = 100
                            Call WriteConsoleMsg(Userindex, "Te has vuelto un experto en el arte de la Supervivencia!", FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(Userindex, "Este pergamino no tiene ning?n conocimiento que te sirva.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                    Case eManualType.otNavegacion
                        If .Stats.UserSkills(eSkill.Navegacion) < 100 Then
                            .Stats.UserSkills(eSkill.Navegacion) = 100
                            Call WriteConsoleMsg(Userindex, "Ya estas listo para comandar todo tipo de embarcacion!", FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(Userindex, "Este pergamino no tiene ning?n conocimiento que te sirva.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                End Select
                Call QuitarUserInvItem(Userindex, Slot, 1)
                Call UpdateUserInv(False, Userindex, Slot)
            End Select
    End With
End Sub

Sub EnivarArmasConstruibles(ByVal Userindex As Integer)
    Call WriteBlacksmithWeapons(Userindex)
End Sub
 
Sub EnivarObjConstruibles(ByVal Userindex As Integer)
    Call WriteInitCarpenting(Userindex)
End Sub

Sub EnivarArmadurasConstruibles(ByVal Userindex As Integer)
    Call WriteBlacksmithArmors(Userindex)
End Sub

Sub TirarTodo(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    With UserList(Userindex)
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then Exit Sub
        Call TirarTodosLosItems(Userindex)
        Dim Cantidad As Long: Cantidad = .Stats.Gld - CLng(.Stats.ELV) * 10000
        If MapInfo(.Pos.Map).Pk Then
            If Cantidad > 0 Then
                Call TirarOro(Cantidad, Userindex)
            End If
        End If
    End With
    Exit Sub
ErrorHandler:
    Call LogError("Error en TirarTodo. Error: " & Err.Number & " - " & Err.description)
End Sub

Public Function ItemSeCae(ByVal index As Integer) As Boolean
    With ObjData(index)
        ItemSeCae = (.Real <> 1 Or .NoSeCae = 0) And (.Caos <> 1 Or .NoSeCae = 0) And .OBJType <> eOBJType.otLlaves And .OBJType <> eOBJType.otBarcos And .NoSeCae = 0
    End With
End Function

Sub TirarTodosLosItems(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Dim i         As Byte
    Dim NuevaPos  As tWorldPos
    Dim MiObj     As obj
    Dim ItemIndex As Integer
    Dim DropAgua  As Boolean
    With UserList(Userindex)
        For i = 1 To .CurrentInventorySlots
            ItemIndex = .Invent.Object(i).ObjIndex
            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.ObjIndex = ItemIndex
                    DropAgua = True
                    If .Clase = eClass.Pirat Then
                        If EsGalera(ObjData(.Invent.BarcoObjIndex)) Then
                            If .Stats.ELV <= 20 Then
                                DropAgua = False
                                Call WriteConsoleMsg(Userindex, "Por que sos Pirata y nivel menor o igual a 20 no se te caen las cosas con la Galera. Cuando llegues a nivel 21 perderas esta condicion.", FontTypeNames.FONTTYPE_WARNING)
                            End If
                        End If
                        If EsGaleon(ObjData(.Invent.BarcoObjIndex)) Then
                            If .Stats.ELV <= 25 Then
                                DropAgua = False
                                Call WriteConsoleMsg(Userindex, "Por que sos Pirata y nivel menor o igual a 25 no se te caen las cosas con el Galeon. Cuando llegues a nivel 26 perderas esta condicion.", FontTypeNames.FONTTYPE_WARNING)
                            End If
                        End If
                    End If
                    Call Tilelibre(.Pos, NuevaPos, MiObj, DropAgua, True)
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(Userindex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                    End If
                End If
            End If
        Next i
    End With
    Exit Sub
ErrorHandler:
    Call LogError("Error en TirarTodosLosItems en linea " & Erl & " - Nick:" & UserList(Userindex).Name & ". Error: " & Err.Number & " - " & Err.description)
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean
    If ItemIndex < 1 Or ItemIndex > UBound(ObjData) Then Exit Function
    ItemNewbie = ObjData(ItemIndex).Newbie = 1
End Function

Sub TirarTodosLosItemsNoNewbies(ByVal Userindex As Integer)
    Dim i         As Byte
    Dim NuevaPos  As tWorldPos
    Dim MiObj     As obj
    Dim ItemIndex As Integer
    With UserList(Userindex)
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then Exit Sub
        For i = 1 To UserList(Userindex).CurrentInventorySlots
            ItemIndex = .Invent.Object(i).ObjIndex
            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    MiObj.Amount = .Invent.Object(i).Amount
                    MiObj.ObjIndex = ItemIndex
                    Call Tilelibre(.Pos, NuevaPos, MiObj, True, True)
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        Call DropObj(Userindex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                    End If
                End If
            End If
        Next i
    End With
End Sub

Sub TirarTodosLosItemsEnMochila(ByVal Userindex As Integer)
    Dim i         As Byte
    Dim NuevaPos  As tWorldPos
    Dim MiObj     As obj
    Dim ItemIndex As Integer
    With UserList(Userindex)
        For i = MAX_NORMAL_INVENTORY_SLOTS + 1 To .CurrentInventorySlots
            ItemIndex = .Invent.Object(i).ObjIndex
            If ItemIndex > 0 Then
                If Not ItemSeCae(ItemIndex) Then
                    Call WriteConsoleMsg(Userindex, "Acabas de tirar un objeto que no se cae normalmente ya que lo tenias en tu mochila u alforja y la desequipaste o tiraste", FontTypeNames.FONTTYPE_WARNING)
                End If
                NuevaPos.X = 0
                NuevaPos.Y = 0
                MiObj.Amount = .Invent.Object(i).Amount
                MiObj.ObjIndex = ItemIndex
                Call Tilelibre(.Pos, NuevaPos, MiObj, True, True)
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    Call DropObj(Userindex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                End If
            End If
        Next i
    End With
End Sub

Public Function getObjType(ByVal ObjIndex As Integer) As eOBJType
    If ObjIndex > 0 Then
        getObjType = ObjData(ObjIndex).OBJType
    End If
End Function

Public Sub moveItem(ByVal Userindex As Integer, ByVal originalSlot As Integer, ByVal newSlot As Integer)
    Dim tmpObj      As UserObj
    Dim newObjIndex As Integer, originalObjIndex As Integer
    If (originalSlot <= 0) Or (newSlot <= 0) Then Exit Sub
    With UserList(Userindex)
        If (originalSlot > .CurrentInventorySlots) Or (newSlot > .CurrentInventorySlots) Then Exit Sub
        tmpObj = .Invent.Object(originalSlot)
        .Invent.Object(originalSlot) = .Invent.Object(newSlot)
        .Invent.Object(newSlot) = tmpObj
        If .Invent.AnilloEqpSlot = originalSlot Then
            .Invent.AnilloEqpSlot = newSlot
        ElseIf .Invent.AnilloEqpSlot = newSlot Then
            .Invent.AnilloEqpSlot = originalSlot
        End If
        If .Invent.ArmourEqpSlot = originalSlot Then
            .Invent.ArmourEqpSlot = newSlot
        ElseIf .Invent.ArmourEqpSlot = newSlot Then
            .Invent.ArmourEqpSlot = originalSlot
        End If
        If .Invent.BarcoSlot = originalSlot Then
            .Invent.BarcoSlot = newSlot
        ElseIf .Invent.BarcoSlot = newSlot Then
            .Invent.BarcoSlot = originalSlot
        End If
        If .Invent.CascoEqpSlot = originalSlot Then
            .Invent.CascoEqpSlot = newSlot
        ElseIf .Invent.CascoEqpSlot = newSlot Then
            .Invent.CascoEqpSlot = originalSlot
        End If
        If .Invent.EscudoEqpSlot = originalSlot Then
            .Invent.EscudoEqpSlot = newSlot
        ElseIf .Invent.EscudoEqpSlot = newSlot Then
            .Invent.EscudoEqpSlot = originalSlot
        End If
        If .Invent.MochilaEqpSlot = originalSlot Then
            .Invent.MochilaEqpSlot = newSlot
        ElseIf .Invent.MochilaEqpSlot = newSlot Then
            .Invent.MochilaEqpSlot = originalSlot
        End If
        If .Invent.MunicionEqpSlot = originalSlot Then
            .Invent.MunicionEqpSlot = newSlot
        ElseIf .Invent.MunicionEqpSlot = newSlot Then
            .Invent.MunicionEqpSlot = originalSlot
        End If
        If .Invent.WeaponEqpSlot = originalSlot Then
            .Invent.WeaponEqpSlot = newSlot
        ElseIf .Invent.WeaponEqpSlot = newSlot Then
            .Invent.WeaponEqpSlot = originalSlot
        End If
        Call UpdateUserInv(False, Userindex, originalSlot)
        Call UpdateUserInv(False, Userindex, newSlot)
    End With
End Sub
