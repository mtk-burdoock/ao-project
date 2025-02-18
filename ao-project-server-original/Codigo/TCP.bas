Attribute VB_Name = "TCP"
#If False Then
    Dim errHandler, Length, index As Variant
#End If

Option Explicit

#If False Then
    Dim X, Y, n, Mapa, Email, Length As Variant
#End If

Private MAX_OBJ_INICIAL As Byte
Private ItemsIniciales() As UserObj

Sub DarCuerpo(ByVal Userindex As Integer)
    Dim NewBody    As Integer
    Dim UserRaza   As Byte
    Dim UserGenero As Byte
    UserGenero = UserList(Userindex).Genero
    UserRaza = UserList(Userindex).raza
    Select Case UserGenero
        Case eGenero.Hombre
            Select Case UserRaza
                Case eRaza.Humano
                    NewBody = 1

                Case eRaza.Elfo
                    NewBody = 2

                Case eRaza.Drow
                    NewBody = 3

                Case eRaza.Enano
                    NewBody = 300

                Case eRaza.Gnomo
                    NewBody = 300
            End Select
            
        Case eGenero.Mujer
            Select Case UserRaza
                Case eRaza.Humano
                    NewBody = 1

                Case eRaza.Elfo
                    NewBody = 2

                Case eRaza.Drow
                    NewBody = 3

                Case eRaza.Gnomo
                    NewBody = 300

                Case eRaza.Enano
                    NewBody = 300
            End Select
    End Select
    UserList(Userindex).Char.body = NewBody
End Sub

Private Function ValidarCabeza(ByVal UserRaza As Byte, ByVal UserGenero As Byte, ByVal Head As Integer) As Boolean
    Select Case UserGenero
        Case eGenero.Hombre
            Select Case UserRaza
                Case eRaza.Humano
                    ValidarCabeza = (Head >= HUMANO_H_PRIMER_CABEZA And Head <= HUMANO_H_ULTIMA_CABEZA)

                Case eRaza.Elfo
                    ValidarCabeza = (Head >= ELFO_H_PRIMER_CABEZA And Head <= ELFO_H_ULTIMA_CABEZA)

                Case eRaza.Drow
                    ValidarCabeza = (Head >= DROW_H_PRIMER_CABEZA And Head <= DROW_H_ULTIMA_CABEZA)

                Case eRaza.Enano
                    ValidarCabeza = (Head >= ENANO_H_PRIMER_CABEZA And Head <= ENANO_H_ULTIMA_CABEZA)

                Case eRaza.Gnomo
                    ValidarCabeza = (Head >= GNOMO_H_PRIMER_CABEZA And Head <= GNOMO_H_ULTIMA_CABEZA)
            End Select
    
        Case eGenero.Mujer
            Select Case UserRaza
                Case eRaza.Humano
                    ValidarCabeza = (Head >= HUMANO_M_PRIMER_CABEZA And Head <= HUMANO_M_ULTIMA_CABEZA)

                Case eRaza.Elfo
                    ValidarCabeza = (Head >= ELFO_M_PRIMER_CABEZA And Head <= ELFO_M_ULTIMA_CABEZA)

                Case eRaza.Drow
                    ValidarCabeza = (Head >= DROW_M_PRIMER_CABEZA And Head <= DROW_M_ULTIMA_CABEZA)

                Case eRaza.Enano
                    ValidarCabeza = (Head >= ENANO_M_PRIMER_CABEZA And Head <= ENANO_M_ULTIMA_CABEZA)

                Case eRaza.Gnomo
                    ValidarCabeza = (Head >= GNOMO_M_PRIMER_CABEZA And Head <= GNOMO_M_ULTIMA_CABEZA)
            End Select
    End Select
End Function

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i   As Integer
    cad = LCase$(cad)
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
            AsciiValidos = False
            Exit Function
        End If
    Next i
    AsciiValidos = True
End Function

Public Function CheckMailString(ByVal sString As String) As Boolean
    On Error GoTo ErrorHandler
    Dim lPos As Long
    Dim lX   As Long
    Dim iAsc As Integer
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then Exit Function
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then Exit Function
            End If
        Next lX
        CheckMailString = True
    End If
ErrorHandler:

End Function

Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or (iAsc >= 65 And iAsc <= 90) Or (iAsc >= 97 And iAsc <= 122) Or (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

Function Numeric(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i   As Integer
    cad = LCase$(cad)
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        If (car < 48 Or car > 57) Then
            Numeric = False
            Exit Function
        End If
    Next i
    Numeric = True
End Function

Function NombrePermitido(ByVal Nombre As String) As Boolean
    Dim i As Integer
    For i = 1 To UBound(ForbidenNames)
        If InStr(Nombre, ForbidenNames(i)) Then
            NombrePermitido = False
            Exit Function
        End If
    Next i
    NombrePermitido = True
End Function

Function ValidateSkills(ByVal Userindex As Integer) As Boolean
    Dim LoopC As Integer
    For LoopC = 1 To NUMSKILLS
        If UserList(Userindex).Stats.UserSkills(LoopC) < 0 Then
            Exit Function
            If UserList(Userindex).Stats.UserSkills(LoopC) > 100 Then UserList(Userindex).Stats.UserSkills(LoopC) = 100
        End If
    Next LoopC
    ValidateSkills = True
End Function

Sub ConnectNewUser(ByVal Userindex As Integer, ByRef Name As String, ByRef AccountHash As String, ByVal UserRaza As eRaza, ByVal UserSexo As eGenero, ByVal UserClase As eClass, ByVal Hogar As eCiudad, ByVal Head As Integer)
    With UserList(Userindex)
        If Not AsciiValidos(Name) Or LenB(Name) = 0 Then
            Call WriteErrorMsg(Userindex, "Nombre invalido.")
            Exit Sub
        End If
        If UserList(Userindex).flags.UserLogged Then
            Call LogCheating("El usuario " & UserList(Userindex).Name & " ha intentado crear a " & Name & " desde la IP " & UserList(Userindex).IP)
            Call CloseSocketSL(Userindex)
            Call Cerrar_Usuario(Userindex)
            Exit Sub
        End If
        If PersonajeExiste(Name) Then
            Call WriteErrorMsg(Userindex, "Ya existe el personaje.")
            Exit Sub
        End If
        If .Stats.UserAtributos(eAtributos.Fuerza) = 0 Then
            Call WriteErrorMsg(Userindex, "Debe tirar los dados antes de poder crear un personaje.")
            Exit Sub
        End If
        If Not ValidarCabeza(UserRaza, UserSexo, Head) Then
            Call LogCheating("El usuario " & Name & " ha seleccionado la cabeza " & Head & " desde la IP " & .IP)
            Call WriteErrorMsg(Userindex, "Cabeza invalida, elija una cabeza seleccionable.")
            Exit Sub
        End If
        .flags.Muerto = 0
        .flags.Escondido = 0
        .Reputacion.AsesinoRep = 0
        .Reputacion.BandidoRep = 0
        .Reputacion.BurguesRep = 0
        .Reputacion.LadronesRep = 0
        .Reputacion.NobleRep = 1000
        .Reputacion.PlebeRep = 30
        .Reputacion.Promedio = 30 / 6
        .Name = Name
        .Clase = UserClase
        .raza = UserRaza
        .Genero = UserSexo
        .Hogar = Hogar
        .AccountHash = AccountHash
        If InventarioUsarConfiguracionPersonalizada Then
            Call AddItemsCustomToNewUser(Userindex)
        Else
            Call AddItemsToNewUser(Userindex, UserClase, UserRaza)
        End If
        Call SetAttributesToNewUser(Userindex, UserClase, UserRaza)
        If EstadisticasInicialesUsarConfiguracionPersonalizada Then
            Call SetAttributesCustomToNewUser(Userindex)
        End If
        Call DarCuerpo(Userindex)
        .Char.heading = eHeading.SOUTH
        .Char.Head = Head
        .OrigChar = .Char
        #If ConUpTime Then
            .LogOnTime = Now
            .UpTime = 0
        #End If
    End With
    Call ResetFacciones(Userindex)
    Call SaveUser(Userindex)
    If Not Database_Enabled Then
        Call SaveUserToAccountCharfile(Name, AccountHash)
    End If
    Call ConnectUser(Userindex, Name, AccountHash)
    If ConexionAPI Then
        Call ApiEndpointSendCreateNewCharacterMessageDiscord(Name)
    End If
End Sub

Private Sub SetAttributesCustomToNewUser(ByVal Userindex As Integer)
    With UserList(Userindex)
        .Stats.Gld = CLng(val(GetVar(IniPath & "Server.ini", "ESTADISTICASINICIALESPJ", "Oro")))
        .Stats.Banco = CLng(val(GetVar(IniPath & "Server.ini", "ESTADISTICASINICIALESPJ", "Banco")))
        Dim InitialLevel, Experiencia As Long
        InitialLevel = val(GetVar(IniPath & "Server.ini", "ESTADISTICASINICIALESPJ", "Nivel"))
        Dim i As Long
        For i = 1 To InitialLevel
            If i <> InitialLevel Then
                .Stats.Exp = .Stats.ELU
                Call CheckUserLevel(Userindex, False)
            End If
        Next i
        Dim SkillPointsIniciales As Long
        SkillPointsIniciales = val(GetVar(IniPath & "Server.ini", "ESTADISTICASINICIALESPJ", "SkillPoints"))
        For i = 1 To NUMSKILLS
            .Stats.UserSkills(i) = SkillPointsIniciales
        Next i
        .Stats.SkillPts = 0
    End With
End Sub

Private Sub SetAttributesToNewUser(ByVal Userindex As Integer, ByVal UserClase As eClass, ByVal UserRaza As eRaza)
    With UserList(Userindex)
        .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + ModRaza(UserRaza).Fuerza
        .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + ModRaza(UserRaza).Agilidad
        .Stats.UserAtributos(eAtributos.Inteligencia) = .Stats.UserAtributos(eAtributos.Inteligencia) + ModRaza(UserRaza).Inteligencia
        .Stats.UserAtributos(eAtributos.Carisma) = .Stats.UserAtributos(eAtributos.Carisma) + ModRaza(UserRaza).Carisma
        .Stats.UserAtributos(eAtributos.Constitucion) = .Stats.UserAtributos(eAtributos.Constitucion) + ModRaza(UserRaza).Constitucion
        Dim i As Long
        For i = 1 To NUMSKILLS
            .Stats.UserSkills(i) = 0
            Call CheckEluSkill(Userindex, i, True)
        Next i
        .Stats.SkillPts = 10
        Dim MiInt As Long
        MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Constitucion) \ 3)
        .Stats.MaxHp = 15 + MiInt
        .Stats.MinHp = 15 + MiInt
        MiInt = RandomNumber(1, .Stats.UserAtributos(eAtributos.Agilidad) \ 6)
        If MiInt = 1 Then MiInt = 2
        .Stats.MaxSta = 20 * MiInt
        .Stats.MinSta = 20 * MiInt
        .Stats.MaxAGU = 100
        .Stats.MinAGU = 100
        .Stats.MaxHam = 100
        .Stats.MinHam = 100
        If UserClase = eClass.Mage Then
            MiInt = .Stats.UserAtributos(eAtributos.Inteligencia) * 3
            .Stats.MaxMAN = MiInt
            .Stats.MinMAN = MiInt
        ElseIf UserClase = eClass.Cleric Or _
               UserClase = eClass.Druid Or _
               UserClase = eClass.Bard Or _
               UserClase = eClass.Assasin Or _
               UserClase = eClass.Bandit Or _
               UserClase = eClass.Paladin Then
            .Stats.MaxMAN = 50
            .Stats.MinMAN = 50
        Else
            .Stats.MaxMAN = 0
            .Stats.MinMAN = 0
        End If
        If UserClase = eClass.Cleric Or _
           UserClase = eClass.Druid Or _
           UserClase = eClass.Bard Or _
           UserClase = eClass.Assasin Or _
           UserClase = eClass.Bandit Or _
           UserClase = eClass.Paladin Or _
           UserClase = eClass.Mage Then
            .Stats.UserHechizos(1) = 2
            If UserClase = eClass.Druid Then .Stats.UserHechizos(2) = 46
        End If
        .Stats.MaxHIT = 2
        .Stats.MinHIT = 1
        .Stats.Gld = 0
        .Stats.Exp = 0
        .Stats.ELU = 300
        .Stats.ELV = 1
    End With
End Sub

Private Sub AddItemsToNewUser(ByVal Userindex As Integer, ByVal UserClase As eClass, ByVal UserRaza As eRaza)
    Dim Slot As Byte
    Dim IsPaladin As Boolean
    IsPaladin = UserClase = eClass.Paladin
    With UserList(Userindex)
        Slot = 1
        .Invent.Object(Slot).ObjIndex = 857
        .Invent.Object(Slot).Amount = 200
        If .Stats.MaxMAN > 0 Or IsPaladin Then
            Slot = Slot + 1
            .Invent.Object(Slot).ObjIndex = 856
            .Invent.Object(Slot).Amount = 200
        Else
            Slot = Slot + 1
            .Invent.Object(Slot).ObjIndex = 855
            .Invent.Object(Slot).Amount = 100
            Slot = Slot + 1
            .Invent.Object(Slot).ObjIndex = 858
            .Invent.Object(Slot).Amount = 50

        End If
        Slot = Slot + 1
        Select Case UserRaza
            Case eRaza.Humano
                .Invent.Object(Slot).ObjIndex = 463
                
            Case eRaza.Elfo
                .Invent.Object(Slot).ObjIndex = 464
                
            Case eRaza.Drow
                .Invent.Object(Slot).ObjIndex = 465
                
            Case eRaza.Enano, eRaza.Gnomo
                .Invent.Object(Slot).ObjIndex = 466
        End Select
        .Invent.Object(Slot).Amount = 1
        .Invent.Object(Slot).Equipped = 1
        .Invent.ArmourEqpSlot = Slot
        .Invent.ArmourEqpObjIndex = .Invent.Object(Slot).ObjIndex
        Slot = Slot + 1
        Select Case UserClase
            Case eClass.Hunter
                
                .Invent.Object(Slot).ObjIndex = 859
            Case eClass.Worker
                
                .Invent.Object(Slot).ObjIndex = RandomNumber(561, 565)
            Case Else
                
                .Invent.Object(Slot).ObjIndex = 460
        End Select
        .Invent.Object(Slot).Amount = 1
        .Invent.Object(Slot).Equipped = 1
        .Invent.WeaponEqpObjIndex = .Invent.Object(Slot).ObjIndex
        .Invent.WeaponEqpSlot = Slot
        .Char.WeaponAnim = GetWeaponAnim(Userindex, .Invent.WeaponEqpObjIndex)

        If UserClase = eClass.Hunter Then
            Slot = Slot + 1
            .Invent.Object(Slot).ObjIndex = 860
            .Invent.Object(Slot).Amount = 150
            .Invent.Object(Slot).Equipped = 1
            .Invent.MunicionEqpSlot = Slot
            .Invent.MunicionEqpObjIndex = 860
        End If
        Slot = Slot + 1
        .Invent.Object(Slot).ObjIndex = 467
        .Invent.Object(Slot).Amount = 100
        Slot = Slot + 1
        .Invent.Object(Slot).ObjIndex = 468
        .Invent.Object(Slot).Amount = 100
        .Char.ShieldAnim = NingunEscudo
        .Char.CascoAnim = NingunCasco
        .Invent.NroItems = Slot
        Dim i As Long
        For i = 1 To MAXAMIGOS
            .Amigos(i).Nombre = vbNullString
            .Amigos(i).Ignorado = 0
            .Amigos(i).index = 0
        Next i
     End With
End Sub

Private Sub AddItemsCustomToNewUser(ByVal Userindex As Integer)
    Dim CantidadItemsIniciales As Integer
    Dim Slot As Long
    Call CargarObjetosIniciales
    With UserList(Userindex)
        For Slot = 1 To MAX_OBJ_INICIAL
            .Invent.Object(Slot).ObjIndex = ItemsIniciales(Slot).ObjIndex
            .Invent.Object(Slot).Amount = ItemsIniciales(Slot).Amount
            .Invent.Object(Slot).Equipped = ItemsIniciales(Slot).Equipped
        Next Slot
    End With
End Sub

Private Sub CargarObjetosIniciales()
    Dim Leer As clsIniManager
    Set Leer = New clsIniManager
    Call Leer.Initialize(IniPath & "Server.ini")
    Dim Slot As Long, sTemp As String
    MAX_OBJ_INICIAL = val(Leer.GetValue("INVENTARIO", "CantidadItemsIniciales"))
    ReDim ItemsIniciales(1 To MAX_OBJ_INICIAL) As UserObj
    For Slot = 1 To MAX_OBJ_INICIAL
        sTemp = Leer.GetValue("INVENTARIO", "Item" & Slot)
        ItemsIniciales(Slot).ObjIndex = val(ReadField(1, sTemp, 45))
        ItemsIniciales(Slot).Amount = val(ReadField(2, sTemp, 45))
        ItemsIniciales(Slot).Equipped = val(ReadField(3, sTemp, 45))
    Next Slot
    Set Leer = Nothing
End Sub

Sub CreateNewAccount(ByVal Userindex As Integer, ByRef UserName As String, ByRef Password As String)
    Dim Salt    As String
    Dim oSHA256 As CSHA256
    Set oSHA256 = New CSHA256
    If Not CheckMailString(UserName) Or LenB(UserName) = 0 Then
        Call WriteErrorMsg(Userindex, "Nombre invalido.")
        Exit Sub
    End If
    If CuentaExiste(UserName) Then
        Call WriteErrorMsg(Userindex, "Ya existe la cuenta.")
        Exit Sub
    End If
    Salt = RandomString(10)
    Call SaveNewAccount(UserName, oSHA256.SHA256(Password & Salt), Salt)
    If ConexionAPI Then
        Call ApiEndpointSendWelcomeEmail(UserName, Password, UserName)
    End If
    Call ConnectAccount(Userindex, UserName, Password)
End Sub

Sub ConnectAccount(ByVal Userindex As Integer, ByRef UserName As String, ByRef Password As String)
    Dim oSHA256 As CSHA256
    Dim Salt    As String
    Set oSHA256 = New CSHA256
    If Not CheckMailString(UserName) Or LenB(UserName) = 0 Then
        Call WriteErrorMsg(Userindex, "Nombre invalido.")
        Exit Sub
    End If
    If Not CuentaExiste(UserName) Then
        Call WriteErrorMsg(Userindex, "No existe la cuenta.")
        Exit Sub
    End If
    Salt = GetAccountSalt(UserName)
    If oSHA256.SHA256(Password & Salt) <> GetAccountPassword(UserName) Then
        Call WriteErrorMsg(Userindex, "Password incorrecto.")
        Call CloseSocket(Userindex)
        Exit Sub
    End If
    If ConexionAPI Then
        Call ApiEndpointSendLoginAccountEmail(UserName, GetLastIpsAccount(UserName), UserList(Userindex).IP)
    End If
    If Not Database_Enabled Then
        Call LoginAccountCharfile(Userindex, UserName)
    Else
        Call SaveAccountLastLoginDatabase(UserName, UserList(Userindex).IP)
        Call LoginAccountDatabase(Userindex, UserName)
    End If
End Sub

Sub CloseSocket(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Call FlushBuffer(Userindex)
    With UserList(Userindex)
        Call SecurityIp.IpRestarConexion(GetLongIp(.IP))
        If .ConnID <> -1 Then
            Call CloseSocketSL(Userindex)
        End If
        If .CentinelaUsuario.centinelaIndex <> 0 Then
            Call modCentinela.UsuarioInActivo(Userindex)
        End If
        If .ComUsu.DestUsu > 0 Then
            If UserList(.ComUsu.DestUsu).flags.UserLogged Then
                If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = Userindex Then
                    Call WriteConsoleMsg(.ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_WARNING)
                    Call FinComerciarUsu(.ComUsu.DestUsu)
                End If
            End If
        End If
        If .flags.SlotReto > 0 Then
            Call Retos.UserDieFight(Userindex, 0, True)
        End If
        If .flags.Equitando = 1 Then
            Call UnmountMontura(Userindex)
        End If
        Call .incomingData.ReadASCIIStringFixed(.incomingData.Length)
        If .flags.UserLogged Then
            If NumUsers > 0 Then NumUsers = NumUsers - 1
            Call CloseUser(Userindex)
        Else
            Call ResetUserSlot(Userindex)
        End If
        Call LiberarSlot(Userindex)
    End With
    Exit Sub
ErrorHandler:
    Call ResetUserSlot(Userindex)
    Call LiberarSlot(Userindex)
    Call LogError("CloseSocket - Error = " & Err.Number & " - Descripcion = " & Err.description & " - UserIndex = " & Userindex)
End Sub

Sub CloseSocketSL(ByVal Userindex As Integer)
    If UserList(Userindex).ConnID <> -1 And UserList(Userindex).ConnIDValida Then
        Call BorraSlotSock(UserList(Userindex).ConnID)
        Call WSApiCloseSocket(UserList(Userindex).ConnID)
        UserList(Userindex).ConnIDValida = False
    End If
End Sub

Function EstaPCarea(index As Integer, Index2 As Integer) As Boolean
    Dim X As Integer, Y As Integer
    For Y = UserList(index).Pos.Y - MinYBorder + 1 To UserList(index).Pos.Y + MinYBorder - 1
        For X = UserList(index).Pos.X - MinXBorder + 1 To UserList(index).Pos.X + MinXBorder - 1
            If MapData(UserList(index).Pos.Map, X, Y).Userindex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        Next X
    Next Y
    EstaPCarea = False
End Function

Function HayPCarea(Pos As WorldPos) As Boolean
    Dim X As Integer, Y As Integer
    For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                If MapData(Pos.Map, X, Y).Userindex > 0 Then
                    HayPCarea = True
                    Exit Function
                End If
            End If
        Next X
    Next Y
    HayPCarea = False
End Function

Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean
    Dim X As Integer, Y As Integer
    For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If MapData(Pos.Map, X, Y).ObjInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function
            End If
        Next X
    Next Y
    HayOBJarea = False
End Function

Function ValidateChr(ByVal Userindex As Integer) As Boolean
    ValidateChr = UserList(Userindex).Char.Head <> 0 And UserList(Userindex).Char.body <> 0 And ValidateSkills(Userindex)
End Function

Sub ConnectUser(ByVal Userindex As Integer, ByRef Name As String, ByRef AccountHash As String)
    Dim n    As Integer
    Dim tStr As String
    With UserList(Userindex)
        If .flags.UserLogged Then
            Call LogCheating("El usuario " & .Name & " ha intentado loguear a " & Name & " desde la IP " & .IP)
            Call CloseSocketSL(Userindex)
            Call Cerrar_Usuario(Userindex)
            Exit Sub
        End If
        .flags.Escondido = 0
        .flags.TargetNPC = 0
        .flags.TargetNpcTipo = eNPCType.Comun
        .flags.TargetObj = 0
        .flags.TargetUser = 0
        .Char.FX = 0
        If NumUsers >= MaxUsers Then
            Call WriteErrorMsg(Userindex, "El servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
            Call CloseSocket(Userindex)
            Exit Sub
        End If
        If AllowMultiLogins = False Then
            If CheckForSameIP(Userindex, .IP) = True Then
                Call WriteErrorMsg(Userindex, "No es posible usar mas de un personaje al mismo tiempo.")
                Call CloseSocket(Userindex)
                Exit Sub
            End If
        End If
        If Not PersonajeExiste(Name) Then
            Call WriteErrorMsg(Userindex, "El personaje no existe.")
            Call CloseSocket(Userindex)
            Exit Sub
        End If
        If Not PersonajePerteneceCuenta(Name, AccountHash) Then
            Call WriteErrorMsg(Userindex, "Ha ocurrido un error, por favor inicie sesion nuevamente.")
            Call CloseSocket(Userindex)
            Exit Sub
        End If
        If CheckForSameName(Name) Then
            If UserList(NameIndex(Name)).Counters.Saliendo Then
                Call WriteErrorMsg(Userindex, "El usuario esta saliendo.")
            Else
                Call WriteErrorMsg(Userindex, "Un usuario con el mismo nombre esta conectado.")
                Call Cerrar_Usuario(NameIndex(Name))
            End If
            Exit Sub
        End If
        .flags.Privilegios = 0
        If EsAdmin(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Admin
            Call LogGM(Name, "Se conecto con ip:" & .IP)
        ElseIf EsDios(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Dios
            Call LogGM(Name, "Se conecto con ip:" & .IP)
        ElseIf EsSemiDios(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.SemiDios
            .flags.PrivEspecial = EsGmEspecial(Name)
            Call LogGM(Name, "Se conecto con ip:" & .IP)
        ElseIf EsConsejero(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.Consejero
            Call LogGM(Name, "Se conecto con ip:" & .IP)
        Else
            .flags.Privilegios = .flags.Privilegios Or PlayerType.User
            .flags.AdminPerseguible = True
        End If
        If EsRolesMaster(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.RoleMaster
        End If
        If ServerSoloGMs > 0 Then
            If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) = 0 Then
                Call WriteErrorMsg(Userindex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
                Call CloseSocket(Userindex)
                Exit Sub
            End If
        End If
        .Name = Name
        Call LoadUser(Userindex)
        If Not ValidateChr(Userindex) Then
            Call WriteErrorMsg(Userindex, "Error en el personaje.")
            Call CloseSocket(Userindex)
            Exit Sub
        End If
        If .Invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = NingunEscudo
        If .Invent.CascoEqpSlot = 0 Then .Char.CascoAnim = NingunCasco
        If .Invent.WeaponEqpSlot = 0 Then .Char.WeaponAnim = NingunArma
        .CurrentInventorySlots = getMaxInventorySlots(Userindex)
        If (.flags.Muerto = 0) Then
            .flags.SeguroResu = False
            Call WriteMultiMessage(Userindex, eMessages.ResuscitationSafeOff)
        Else
            .flags.SeguroResu = True
            Call WriteMultiMessage(Userindex, eMessages.ResuscitationSafeOn)
        End If
        Call UpdateUserInv(True, Userindex, 0)
        Call UpdateUserHechizos(True, Userindex, 0)
        Call ActualizarSlotAmigo(Userindex, 0, True)
        Call ObtenerIndexAmigos(Userindex, False)
        If .flags.Paralizado Then
            Call WriteParalizeOK(Userindex)
        End If
        Dim Mapa As Integer
        Mapa = .Pos.Map
        If Mapa = 0 Then
            If UsarMundoPropio Then
                .Pos = CustomSpawnMap
                Mapa = CustomSpawnMap.Map
            Else
                .Pos = Nemahuak
                Mapa = Nemahuak.Map
            End If
        Else
            If Not MapaValido(Mapa) Then
                Call WriteErrorMsg(Userindex, "El PJ se encuenta en un mapa invalido.")
                Call CloseSocket(Userindex)
                Exit Sub
            End If
            Dim StartMap As Integer
            StartMap = MapInfo(Mapa).StartPos.Map
            If StartMap <> 0 Then
                If MapaValido(StartMap) Then
                    .Pos = MapInfo(Mapa).StartPos
                    Mapa = StartMap
                End If
            End If
        End If
        If MapData(Mapa, .Pos.X, .Pos.Y).Userindex <> 0 Or MapData(Mapa, .Pos.X, .Pos.Y).NpcIndex <> 0 Then
            Dim FoundPlace As Boolean
            Dim esAgua     As Boolean
            Dim tX         As Long
            Dim tY         As Long
            FoundPlace = False
            esAgua = HayAgua(Mapa, .Pos.X, .Pos.Y)
            For tY = .Pos.Y - 1 To .Pos.Y + 1
                For tX = .Pos.X - 1 To .Pos.X + 1
                    If esAgua Then
                        If LegalPos(Mapa, tX, tY, True, False) Then
                            FoundPlace = True
                            Exit For
                        End If
                    Else
                        If LegalPos(Mapa, tX, tY, False, True) Then
                            FoundPlace = True
                            Exit For
                        End If
                    End If
                Next tX
                If FoundPlace Then Exit For
            Next tY
            If FoundPlace Then
                .Pos.X = tX
                .Pos.Y = tY
            Else
                If MapData(Mapa, .Pos.X, .Pos.Y).Userindex <> 0 Then
                    If UserList(MapData(Mapa, .Pos.X, .Pos.Y).Userindex).ComUsu.DestUsu > 0 Then
                        If UserList(UserList(MapData(Mapa, .Pos.X, .Pos.Y).Userindex).ComUsu.DestUsu).flags.UserLogged Then
                            Call FinComerciarUsu(UserList(MapData(Mapa, .Pos.X, .Pos.Y).Userindex).ComUsu.DestUsu)
                            Call WriteConsoleMsg(UserList(MapData(Mapa, .Pos.X, .Pos.Y).Userindex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_WARNING)
                        End If
                        If UserList(MapData(Mapa, .Pos.X, .Pos.Y).Userindex).flags.UserLogged Then
                            Call FinComerciarUsu(MapData(Mapa, .Pos.X, .Pos.Y).Userindex)
                            Call WriteErrorMsg(MapData(Mapa, .Pos.X, .Pos.Y).Userindex, "Alguien se ha conectado donde te encontrabas, por favor reconectate...")
                        End If
                    End If
                    Call CloseSocket(MapData(Mapa, .Pos.X, .Pos.Y).Userindex)
                End If
            End If
        End If
        .showName = True
        If .Invent.BarcoObjIndex > 0 And (HayAgua(Mapa, .Pos.X, .Pos.Y) Or BodyIsBoat(.Char.body)) Then
            .Char.Head = 0
            If .flags.Muerto = 0 Then
                Call ToggleBoatBody(Userindex)
            Else
                .Char.body = iFragataFantasmal
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco
            End If
            .flags.Navegando = 1
        End If
        Call WriteUserIndexInServer(Userindex)
        Call WriteChangeMap(Userindex, .Pos.Map, MapInfo(.Pos.Map).MapVersion)
        If MapInfo(.Pos.Map).MusicMp3 <> vbNullString Then
            Call WritePlayMp3(Userindex, MapInfo(.Pos.Map).MusicMp3)
        Else
            Call WritePlayMidi(Userindex, val(ReadField(1, MapInfo(.Pos.Map).Music, 45)))
        End If
        If .flags.Privilegios = PlayerType.Dios Then
            .flags.ChatColor = RGB(250, 250, 150)
        ElseIf .flags.Privilegios <> PlayerType.User And .flags.Privilegios <> (PlayerType.User Or PlayerType.ChaosCouncil) And .flags.Privilegios <> (PlayerType.User Or PlayerType.RoyalCouncil) Then
            .flags.ChatColor = RGB(0, 255, 0)
        ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.RoyalCouncil) Then
            .flags.ChatColor = RGB(0, 255, 255)
        ElseIf .flags.Privilegios = (PlayerType.User Or PlayerType.ChaosCouncil) Then
            .flags.ChatColor = RGB(255, 128, 64)
        Else
            .flags.ChatColor = vbWhite
        End If
        #If ConUpTime Then
            .LogOnTime = Now
        #End If
        If (.flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster)) = 0 Then
            Call DoAdminInvisible(Userindex)
            .flags.SendDenounces = True
        End If
        Call MakeUserChar(True, .Pos.Map, Userindex, .Pos.Map, .Pos.X, .Pos.Y)
        Call WriteUserCharIndexInServer(Userindex)
        Call DoTileEvents(Userindex, .Pos.Map, .Pos.X, .Pos.Y)
        Call CheckUserLevel(Userindex)
        Call WriteUpdateUserStats(Userindex)
        Call WriteUpdateHungerAndThirst(Userindex)
        Call WriteUpdateStrenghtAndDexterity(Userindex)
        Call SendMOTD(Userindex)
        If haciendoBK Then
            Call WritePauseToggle(Userindex)
            Call WriteConsoleMsg(Userindex, "Servidor> Por favor espera algunos segundos, el WorldSave esta ejecutandose.", FontTypeNames.FONTTYPE_SERVER)
        End If
        If EnPausa Then
            Call WritePauseToggle(Userindex)
            Call WriteConsoleMsg(Userindex, "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar mas tarde.", FontTypeNames.FONTTYPE_SERVER)
        End If
        If EnTesting And .Stats.ELV >= 18 Then
            Call WriteErrorMsg(Userindex, "Servidor en Testing por unos minutos, conectese con PJs de nivel menor a 18. No se conecte con Pjs que puedan resultar importantes por ahora pues pueden arruinarse.")
            Call CloseSocket(Userindex)
            Exit Sub
        End If
        NumUsers = NumUsers + 1
        .flags.UserLogged = True
        Call UpdateUserLogged(.Name, 1)
        MapInfo(.Pos.Map).NumUsers = MapInfo(.Pos.Map).NumUsers + 1
        If .Stats.SkillPts > 0 Then
            Call WriteSendSkills(Userindex)
            Call WriteLevelUp(Userindex, .Stats.SkillPts)
        End If
        If NumUsers > RecordUsuariosOnline Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Record de usuarios conectados simultaneamente. Hay " & NumUsers & " usuarios.", FontTypeNames.FONTTYPE_INFOBOLD))
            RecordUsuariosOnline = NumUsers
            Call WriteVar(IniPath & "Server.ini", "INIT", "RECORD", Str(RecordUsuariosOnline))
            frmMain.txtRecordOnline.Text = RecordUsuariosOnline
        End If
        If .NroMascotas > 0 And MapInfo(.Pos.Map).Pk Then
            Dim i As Integer
            For i = 1 To MAXMASCOTAS
                If .MascotasType(i) > 0 Then
                    .MascotasIndex(i) = SpawnNpc(.MascotasType(i), .Pos, True, True)
                    If .MascotasIndex(i) > 0 Then
                        Npclist(.MascotasIndex(i)).MaestroUser = Userindex
                        Call FollowAmo(.MascotasIndex(i))
                    Else
                        .MascotasIndex(i) = 0
                    End If
                End If
            Next i
        End If
        If .flags.Navegando = 1 Then
            Call WriteNavigateToggle(Userindex)
        End If
        If criminal(Userindex) Then
            Call WriteMultiMessage(Userindex, eMessages.SafeModeOff)
            .flags.Seguro = False
        Else
            .flags.Seguro = True
            Call WriteMultiMessage(Userindex, eMessages.SafeModeOn)
        End If
        If .GuildIndex > 0 Then
            If Not modGuilds.m_ConectarMiembroAClan(Userindex, .GuildIndex) Then
                Call WriteConsoleMsg(Userindex, "Tu estado no te permite entrar al clan.", FontTypeNames.FONTTYPE_GUILD)
            End If
        End If
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
        Call WriteLoggedMessage(Userindex)
        If (.flags.Muerto = 0) Then
            .flags.SeguroResu = False
            Call WriteMultiMessage(Userindex, eMessages.ResuscitationSafeOff)
        Else
            .flags.SeguroResu = True
            Call WriteMultiMessage(Userindex, eMessages.ResuscitationSafeOn)
        End If
        If criminal(Userindex) Then
            Call WriteMultiMessage(Userindex, eMessages.SafeModeOff)
            .flags.Seguro = False
        Else
            .flags.Seguro = True
            Call WriteMultiMessage(Userindex, eMessages.SafeModeOn)
        End If
        Call IntervaloPermiteSerAtacado(Userindex, True)
        If Lloviendo Then
            Call WriteRainToggle(Userindex)
        End If
        tStr = modGuilds.a_ObtenerRechazoDeChar(.Name)
        If LenB(tStr) <> 0 Then
            Call WriteShowMessageBox(Userindex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)
        End If
        Call Statistics.UserConnected(Userindex)
        Call MostrarNumUsers
        Call modGuilds.SendGuildNews(Userindex)
        If ConexionAPI Then
            Call ApiEndpointSendUserConnectedMessageDiscord(Name, .Desc, criminal(Userindex), ListaClases(.Clase))
        End If
        n = FreeFile
        Open App.Path & "\logs\numusers.log" For Output As n
        Print #n, NumUsers
        Close #n
        n = FreeFile
        Open App.Path & "\logs\Connect.log" For Append Shared As #n
        Print #n, .Name & " ha entrado al juego. UserIndex:" & Userindex & " " & time & " " & Date
        Close #n
    End With
End Sub

Sub SendMOTD(ByVal Userindex As Integer)
    Dim j As Long
    Call WriteGuildChat(Userindex, "Mensajes de entrada:")
    For j = 1 To MaxLines
        Call WriteGuildChat(Userindex, MOTD(j).texto)
    Next j
End Sub

Sub ResetFacciones(ByVal Userindex As Integer)
    With UserList(Userindex).Faccion
        .ArmadaReal = 0
        .CiudadanosMatados = 0
        .CriminalesMatados = 0
        .FuerzasCaos = 0
        .FechaIngreso = vbNullString
        .RecibioArmaduraCaos = 0
        .RecibioArmaduraReal = 0
        .RecibioExpInicialCaos = 0
        .RecibioExpInicialReal = 0
        .RecompensasCaos = 0
        .RecompensasReal = 0
        .Reenlistadas = 0
        .NivelIngreso = 0
        .MatadosIngreso = 0
        .NextRecompensa = 0
    End With
End Sub

Sub ResetContadores(ByVal Userindex As Integer)
    With UserList(Userindex).Counters
        .TimeFight = 0
        .AGUACounter = 0
        .AsignedSkills = 0
        .AttackCounter = 0
        .bPuedeMeditar = True
        .Ceguera = 0
        .COMCounter = 0
        .Estupidez = 0
        .failedUsageAttempts = 0
        .Frio = 0
        .goHome = 0
        .HPCounter = 0
        .IdleCount = 0
        .Invisibilidad = 0
        .Lava = 0
        .Mimetismo = 0
        .Ocultando = 0
        .Paralisis = 0
        .Pena = 0
        .PiqueteC = 0
        .Saliendo = False
        .Salir = 0
        .STACounter = 0
        .TiempoOculto = 0
        .TimerEstadoAtacable = 0
        .TimerGolpeMagia = 0
        .TimerGolpeUsar = 0
        .TimerLanzarSpell = 0
        .TimerMagiaGolpe = 0
        .TimerPerteneceNpc = 0
        .TimerPuedeAtacar = 0
        .TimerPuedeSerAtacado = 0
        .TimerPuedeTrabajar = 0
        .TimerPuedeUsarArco = 0
        .TimerUsar = 0
        .Trabajando = 0
        .Veneno = 0
    End With
    Call modAntiCheat.ResetAllCount(Userindex)
End Sub

Sub ResetCharInfo(ByVal Userindex As Integer)
    With UserList(Userindex).Char
        .Escribiendo = 0
        .body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .Head = 0
        .loops = 0
        .heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0

    End With

End Sub

Sub ResetBasicUserInfo(ByVal Userindex As Integer)
    With UserList(Userindex)
        .Name = vbNullString
        .ID = 0
        .AccountHash = vbNullString
        .Desc = vbNullString
        .DescRM = vbNullString
        .Pos.Map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .IP = vbNullString
        .Clase = 0
        .Email = vbNullString
        .Genero = 0
        .Hogar = 0
        .raza = 0
        .PartyIndex = 0
        .PartySolicitud = 0
        With .Stats
            .Banco = 0
            .ELV = 0
            .ELU = 0
            .Exp = 0
            .def = 0
            .NPCsMuertos = 0
            .UsuariosMatados = 0
            .SkillPts = 0
            .Gld = 0
            .UserAtributos(1) = 0
            .UserAtributos(2) = 0
            .UserAtributos(3) = 0
            .UserAtributos(4) = 0
            .UserAtributos(5) = 0
            .UserAtributosBackUP(1) = 0
            .UserAtributosBackUP(2) = 0
            .UserAtributosBackUP(3) = 0
            .UserAtributosBackUP(4) = 0
            .UserAtributosBackUP(5) = 0
        End With
    End With
End Sub

Sub ResetReputacion(ByVal Userindex As Integer)
    With UserList(Userindex).Reputacion
        .AsesinoRep = 0
        .BandidoRep = 0
        .BurguesRep = 0
        .LadronesRep = 0
        .NobleRep = 0
        .PlebeRep = 0
        .NobleRep = 0
        .Promedio = 0
    End With
End Sub

Sub ResetGuildInfo(ByVal Userindex As Integer)
    If UserList(Userindex).EscucheClan > 0 Then
        Call modGuilds.GMDejaDeEscucharClan(Userindex, UserList(Userindex).EscucheClan)
        UserList(Userindex).EscucheClan = 0
    End If
    If UserList(Userindex).GuildIndex > 0 Then
        Call modGuilds.m_DesconectarMiembroDelClan(Userindex, UserList(Userindex).GuildIndex)
    End If
    UserList(Userindex).GuildIndex = 0
End Sub

Sub ResetUserFlags(ByVal Userindex As Integer)
    With UserList(Userindex).flags
        .SlotReto = 0
        .SlotRetoUser = 255
        .Comerciando = False
        .SlotCarcel = 0
        .Ban = 0
        .Escondido = 0
        .DuracionEfecto = 0
        .NpcInv = 0
        .StatsChanged = 0
        .TargetNPC = 0
        .TargetNpcTipo = eNPCType.Comun
        .TargetObj = 0
        .TargetObjMap = 0
        .TargetObjX = 0
        .TargetObjY = 0
        .TargetUser = 0
        .TipoPocion = 0
        .TomoPocion = False
        .Descuento = vbNullString
        .Hambre = 0
        .Sed = 0
        .Descansar = False
        .Vuela = 0
        .Navegando = 0
        .Equitando = 0
        .Oculto = 0
        .Envenenado = 0
        .invisible = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Maldicion = 0
        .Bendicion = 0
        .Meditando = 0
        .Privilegios = 0
        .PrivEspecial = False
        .PuedeMoverse = 0
        .OldBody = 0
        .OldHead = 0
        .AdminInvisible = 0
        .ValCoDe = 0
        .Hechizo = 0
        .TimesWalk = 0
        .StartWalk = 0
        .CountSH = 0
        .Silenciado = 0
        .AdminPerseguible = False
        .lastMap = 0
        .Traveling = 0
        .AtacablePor = 0
        .AtacadoPorNpc = 0
        .AtacadoPorUser = 0
        .NoPuedeSerAtacado = False
        .ShareNpcWith = 0
        .EnConsulta = False
        .Ignorado = False
        .SendDenounces = False
        .ParalizedBy = vbNullString
        .ParalizedByIndex = 0
        .ParalizedByNpcIndex = 0
        If .OwnedNpc <> 0 Then
            Call PerdioNpc(Userindex)
        End If
    End With
End Sub

Sub ResetUserSpells(ByVal Userindex As Integer)
    Dim LoopC As Long
    For LoopC = 1 To MAXUSERHECHIZOS
        UserList(Userindex).Stats.UserHechizos(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserPets(ByVal Userindex As Integer)
    Dim LoopC As Long
    UserList(Userindex).NroMascotas = 0
    For LoopC = 1 To MAXMASCOTAS
        UserList(Userindex).MascotasIndex(LoopC) = 0
        UserList(Userindex).MascotasType(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserBanco(ByVal Userindex As Integer)
    Dim LoopC As Long
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
        UserList(Userindex).BancoInvent.Object(LoopC).Amount = 0
        UserList(Userindex).BancoInvent.Object(LoopC).Equipped = 0
        UserList(Userindex).BancoInvent.Object(LoopC).ObjIndex = 0
    Next LoopC
    UserList(Userindex).BancoInvent.NroItems = 0
End Sub

Public Sub LimpiarComercioSeguro(ByVal Userindex As Integer)
    With UserList(Userindex).ComUsu
        If .DestUsu > 0 Then
            Call FinComerciarUsu(.DestUsu)
            Call FinComerciarUsu(Userindex)
        End If
    End With
End Sub

Sub ResetUserSlot(ByVal Userindex As Integer)
    Dim i As Long
    UserList(Userindex).ConnIDValida = False
    UserList(Userindex).ConnID = -1
    Call LimpiarComercioSeguro(Userindex)
    Call ResetFacciones(Userindex)
    Call ResetContadores(Userindex)
    Call ResetGuildInfo(Userindex)
    Call ResetCharInfo(Userindex)
    Call ResetBasicUserInfo(Userindex)
    Call ResetReputacion(Userindex)
    Call ResetUserFlags(Userindex)
    Call LimpiarInventario(Userindex)
    Call ResetUserSpells(Userindex)
    Call ResetUserPets(Userindex)
    Call ResetUserBanco(Userindex)
    Call ResetQuestStats(Userindex)
    Call ResetUserExtras(Userindex)
    With UserList(Userindex).ComUsu
        .Acepto = False
        For i = 1 To MAX_OFFER_SLOTS
            .cant(i) = 0
            .Objeto(i) = 0
        Next i
        .GoldAmount = 0
        .DestNick = vbNullString
        .DestUsu = 0
    End With
End Sub

Sub CloseUser(ByVal Userindex As Integer)
    On Error GoTo ErrorHandler
    Dim n    As Integer
    Dim Map  As Integer
    Dim Name As String
    Dim i    As Integer
    Dim aN   As Integer
    With UserList(Userindex)
        aN = .flags.AtacadoPorNpc
        If aN > 0 Then
            Npclist(aN).Movement = Npclist(aN).flags.OldMovement
            Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
            Npclist(aN).flags.AttackedBy = vbNullString
        End If
        aN = .flags.NPCAtacado
        If aN > 0 Then
            If Npclist(aN).flags.AttackedFirstBy = .Name Then
                Npclist(aN).flags.AttackedFirstBy = vbNullString
            End If
        End If
        .flags.AtacadoPorNpc = 0
        .flags.NPCAtacado = 0
        Map = .Pos.Map
        Name = UCase$(.Name)
        .Char.FX = 0
        .Char.loops = 0
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
        .flags.UserLogged = False
        .Counters.Saliendo = False
        .flags.AdminInvisible = 0
        Call ObtenerIndexAmigos(Userindex, True)
        If .PartyIndex > 0 Then Call mdParty.SalirDeParty(Userindex)
        Call Statistics.UserDisconnected(Userindex)
        Call SaveUser(Userindex)
        Call UpdateUserLogged(.Name, 0)
        If MapInfo(Map).NumUsers > 0 Then
            Call SendData(SendTarget.ToPCAreaButIndex, Userindex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        End If
        If .Char.CharIndex > 0 Then
            Call EraseUserChar(Userindex, .flags.AdminInvisible = 1)
        End If
        For i = 1 To MAXMASCOTAS
            If .MascotasIndex(i) > 0 Then
                If Npclist(.MascotasIndex(i)).flags.NPCActive Then Call QuitarNPC(.MascotasIndex(i))
            End If
        Next i
        MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1
        If MapInfo(Map).NumUsers < 0 Then
            MapInfo(Map).NumUsers = 0
        End If
        If Ayuda.Existe(.Name) Then Call Ayuda.Quitar(.Name)
        Call ResetUserSlot(Userindex)
        Call MostrarNumUsers
        n = FreeFile(1)
        Open App.Path & "\logs\Connect.log" For Append Shared As #n
        Print #n, Name & " ha dejado el juego. " & "User Index:" & Userindex & " " & time & " " & Date
        Close #n
    End With
    Exit Sub
ErrorHandler:
    Call LogError("Error en CloseUser. Numero " & Err.Number & " Descripcion: " & Err.description)
End Sub

Sub ReloadSokcet()
    On Error GoTo ErrorHandler
    Call LogApiSock("ReloadSokcet() " & NumUsers & " " & LastUser & " " & MaxUsers)
    If NumUsers <= 0 Then
        Call WSApiReiniciarSockets
    End If
    Exit Sub
ErrorHandler:
    Call LogError("Error en CheckSocketState " & Err.Number & ": " & Err.description)
End Sub

Public Sub EnviarNoche(ByVal Userindex As Integer)
    Call WriteSendNight(Userindex, IIf(DeNoche And (MapInfo(UserList(Userindex).Pos.Map).Zona = Campo Or MapInfo(UserList(Userindex).Pos.Map).Zona = Ciudad), True, False))
    Call WriteSendNight(Userindex, IIf(DeNoche, True, False))
End Sub

Public Sub EcharPjsNoPrivilegiados()
    Dim LoopC As Long
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
            If UserList(LoopC).flags.Privilegios And PlayerType.User Then
                Call CloseSocket(LoopC)
            End If
        End If
    Next LoopC
End Sub

Function RandomString(cb As Integer) As String
    Randomize
    Dim rgch As String
    rgch = "abcdefghijklmnopqrstuvwxyz"
    rgch = rgch & UCase(rgch) & "0123456789" & "#@!~$()-_"
    Dim i As Long
    For i = 1 To cb
        RandomString = RandomString & mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
    Next
End Function

Public Sub ResetUserExtras(ByVal Userindex As Integer)
    Dim i As Long
    For i = 1 To MAXAMIGOS
        UserList(Userindex).Amigos(i).Nombre = vbNullString
        UserList(Userindex).Amigos(i).Ignorado = 0
        UserList(Userindex).Amigos(i).index = 0
    Next i
    UserList(Userindex).Quien = vbNullString
End Sub
