Attribute VB_Name = "modMainLoop"
Option Explicit

Public prgRun As Boolean
Public LastGameTick As Long
Private Const GAME_TIMER_INTERVAL = 40

Public Sub Auditoria()
    On Error GoTo ErrorHandler
    Call PasarSegundo
    Static centinelSecs As Byte
    centinelSecs = centinelSecs + 1
    If centinelSecs = 5 Then
        Call modCentinela.AvisarUsuarios
        centinelSecs = 0
    End If
    Exit Sub
ErrorHandler:
    Call LogError("Error en Timer Auditoria. Err: " & Err.description & " - " & Err.Number)
End Sub

Public Sub PacketResend()
    On Error GoTo ErrorHandler:
    Dim i As Long
    For i = 1 To LastUser
        If UserList(i).ConnIDValida Then Call FlushBuffer(i)
    Next i
    Exit Sub
ErrorHandler:
    Call LogError("Error en packetResend - Error: " & Err.Number & " - Desc: " & Err.description)
    Resume Next
End Sub

Public Sub TimerEnviarDatosServer()
    On Error GoTo ErrorHandler
    Call ApiEndpointSendServerDataToApiToShowOnlineUsers
ErrorHandler:
    If Err.Number <> 0 Then
        Call LogError("Error en TimerEnviarDatosServer - Error: " & Err.Number & " - Desc: " & Err.description)
    End If
End Sub

Public Sub TIMER_AI()
    On Error GoTo ErrorHandler
    Dim NpcIndex As Long
    Dim Mapa     As Integer
    Dim e_p      As Integer
    If Not haciendoBK And Not EnPausa Then
        For NpcIndex = 1 To LastNPC
            With Npclist(NpcIndex)
                If .flags.NPCActive Then
                    If .Owner > 0 Then Call ValidarPermanenciaNpc(NpcIndex)
                    If .flags.Paralizado = 1 Then
                        Call EfectoParalisisNpc(NpcIndex)
                    Else
                        If .NPCtype = eNPCType.Pretoriano Then
                            Call ClanPretoriano(.ClanIndex).PerformPretorianAI(NpcIndex)
                        Else
                            If .flags.Inmovilizado = 1 Then
                                Call EfectoParalisisNpc(NpcIndex)
                            End If
                            Mapa = .Pos.Map
                            If Mapa > 0 Then
                                If MapInfo(Mapa).NumUsers > 0 Then
                                    If .Movement <> TipoAI.ESTATICO Then
                                        Call NPCAI(NpcIndex)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End With
        Next NpcIndex
    End If
    Exit Sub
ErrorHandler:
    Call LogError("Error en TIMER_AI_Timer " & Npclist(NpcIndex).Name & " mapa:" & Npclist(NpcIndex).Pos.Map)
    Call MuereNpc(NpcIndex, 0)
End Sub

Public Sub GameTimer()
On Error GoTo ErrorHandler
    Dim iUserIndex   As Long
    Dim bEnviarStats As Boolean
    Dim bEnviarAyS   As Boolean
    Dim DeltaTick    As Single
    DeltaTick = (GetTickCount - LastGameTick) / GAME_TIMER_INTERVAL
    LastGameTick = GetTickCount
    For iUserIndex = 1 To LastUser
        With UserList(iUserIndex)
            If .ConnID <> -1 Then
                If .ConnIDValida And .flags.UserLogged Then
                    bEnviarStats = False
                    bEnviarAyS = False
                    If .flags.Paralizado = 1 Then Call EfectoParalisisUser(iUserIndex, DeltaTick)
                    If .flags.Ceguera = 1 Or .flags.Estupidez Then Call EfectoCegueEstu(iUserIndex, DeltaTick)
                    If .flags.Muerto = 0 Then
                        If (.flags.Privilegios And PlayerType.User) Then Call EfectoLava(iUserIndex, DeltaTick)
                        If .flags.Desnudo <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoFrio(iUserIndex, DeltaTick)
                        If .flags.Meditando Then Call DoMeditar(iUserIndex, DeltaTick)
                        If .flags.Envenenado <> 0 And (.flags.Privilegios And PlayerType.User) <> 0 Then Call EfectoVeneno(iUserIndex, DeltaTick)
                        If .flags.AdminInvisible <> 1 Then
                            If .flags.invisible = 1 Then Call EfectoInvisibilidad(iUserIndex, DeltaTick)
                            If .flags.Oculto = 1 Then Call DoPermanecerOculto(iUserIndex, DeltaTick)
                        End If
                        If .flags.Mimetizado = 1 Then Call EfectoMimetismo(iUserIndex, DeltaTick)
                        If .flags.AtacablePor <> 0 Then Call EfectoEstadoAtacable(iUserIndex)
                        Call DuracionPociones(iUserIndex, DeltaTick)
                        Call HambreYSed(iUserIndex, DeltaTick, bEnviarAyS)
                        If .flags.Hambre = 0 And .flags.Sed = 0 Then
                            If Lloviendo Then
                                If Not Intemperie(iUserIndex) Then
                                    If Not .flags.Descansar Then
                                        Call Sanar(iUserIndex, DeltaTick, bEnviarStats, SanaIntervaloSinDescansar)
                                        If bEnviarStats Then
                                            Call WriteUpdateHP(iUserIndex)
                                            bEnviarStats = False
                                        End If
                                        Call RecStamina(iUserIndex, DeltaTick, bEnviarStats, StaminaIntervaloSinDescansar)
                                        If bEnviarStats Then
                                            Call WriteUpdateSta(iUserIndex)
                                            bEnviarStats = False
                                        End If
                                    Else
                                        Call Sanar(iUserIndex, DeltaTick, bEnviarStats, SanaIntervaloDescansar)
                                        If bEnviarStats Then
                                            Call WriteUpdateHP(iUserIndex)
                                            bEnviarStats = False
                                        End If
                                        Call RecStamina(iUserIndex, DeltaTick, bEnviarStats, StaminaIntervaloDescansar)
                                        If bEnviarStats Then
                                            Call WriteUpdateSta(iUserIndex)
                                            bEnviarStats = False
                                        End If
                                        If .Stats.MaxHp = .Stats.MinHp And .Stats.MaxSta = .Stats.MinSta Then
                                            Call WriteRestOK(iUserIndex)
                                            Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                            .flags.Descansar = False
                                        End If
                                    End If
                                End If
                            Else
                                If Not .flags.Descansar Then
                                    Call Sanar(iUserIndex, DeltaTick, bEnviarStats, SanaIntervaloSinDescansar)
                                    If bEnviarStats Then
                                        Call WriteUpdateHP(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    Call RecStamina(iUserIndex, DeltaTick, bEnviarStats, StaminaIntervaloSinDescansar)
                                    If bEnviarStats Then
                                        Call WriteUpdateSta(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                Else
                                    Call Sanar(iUserIndex, DeltaTick, bEnviarStats, SanaIntervaloDescansar)
                                    If bEnviarStats Then
                                        Call WriteUpdateHP(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    Call RecStamina(iUserIndex, DeltaTick, bEnviarStats, StaminaIntervaloDescansar)
                                    If bEnviarStats Then
                                        Call WriteUpdateSta(iUserIndex)
                                        bEnviarStats = False
                                    End If
                                    If .Stats.MaxHp = .Stats.MinHp And .Stats.MaxSta = .Stats.MinSta Then
                                        Call WriteRestOK(iUserIndex)
                                        Call WriteConsoleMsg(iUserIndex, "Has terminado de descansar.", FontTypeNames.FONTTYPE_INFO)
                                        .flags.Descansar = False
                                    End If
                                End If
                            End If
                        End If
                        If bEnviarAyS Then Call WriteUpdateHungerAndThirst(iUserIndex)
                        If .NroMascotas > 0 Then Call TiempoInvocacion(iUserIndex, DeltaTick)
                    Else
                        If .flags.Traveling <> 0 Then Call TravelingEffect(iUserIndex)
                    End If

                Else
                    .Counters.IdleCount = .Counters.IdleCount + DeltaTick
                    If .Counters.IdleCount > IntervaloParaConexion Then
                        .Counters.IdleCount = 0
                        Call CloseSocket(iUserIndex)
                    End If
                End If
                .Counters.PacketsTick = 0
            End If
        End With
    Next iUserIndex
    Exit Sub
ErrorHandler:
    LogError ("Error en GameTimer: " & Err.description & " UserIndex = " & iUserIndex)
End Sub

Public Sub PasarSegundo()
    On Error GoTo ErrorHandler
    Dim i As Long
    If tickLimpieza > 0 Then
        tickLimpieza = tickLimpieza - 1
        Select Case tickLimpieza
            Case 300
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpieza del mundo en 5 Minuto. Atentos!!", FontTypeNames.FONTTYPE_SERVER))

            Case 240
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpieza del mundo en 4 Minutos. Atentos!!", FontTypeNames.FONTTYPE_SERVER))

            Case 180
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpieza del mundo en 3 Minutos. Atentos!!", FontTypeNames.FONTTYPE_SERVER))
            
            Case 120
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpieza del mundo en 2 Minutos. Atentos!!", FontTypeNames.FONTTYPE_SERVER))

            Case 60
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpieza del mundo en 1 Minuto. Atentos!!", FontTypeNames.FONTTYPE_SERVER))
   
            Case 15
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpieza del mundo en 15 segundos. Atentos!!", FontTypeNames.FONTTYPE_SERVER))
            
            Case 10
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpieza del mundo en 10 segundos. Atentos!!", FontTypeNames.FONTTYPE_SERVER))
                
            Case 5 To 1
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpieza del mundo en " & tickLimpieza & " segundos. Atentos!!", FontTypeNames.FONTTYPE_SERVER))
            
            Case 0
                Call BorrarObjetosLimpieza
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Limpieza del mundo finalizada.", FontTypeNames.FONTTYPE_SERVER))
        End Select
    End If
    For i = 1 To LastUser
        With UserList(i)
            If .flags.UserLogged Then
                If .Counters.Saliendo Then
                    .Counters.Salir = .Counters.Salir - 1
                    If .Counters.Salir <= 0 Then
                        Call WriteConsoleMsg(i, "Gracias por jugar Argentum Online", FontTypeNames.FONTTYPE_INFO)
                        Call WriteDisconnect(i)
                        Call FlushBuffer(i)
                        Call CloseSocket(i)
                    End If
                End If
                If .Counters.MonturaCounter > 0 Then .Counters.MonturaCounter = .Counters.MonturaCounter - 1
                If .Counters.TimeFight > 0 Then
                    .Counters.TimeFight = .Counters.TimeFight - 1
                    If .Counters.TimeFight = 0 Then
                        Call WriteConsoleMsg(i, "Cuenta -> YA!", FontTypeNames.FONTTYPE_FIGHT)
                        If .flags.SlotReto > 0 Then
                            Call WriteUserInEvent(i)
                        End If
                    Else
                        Call WriteConsoleMsg(i, "Cuenta -> " & .Counters.TimeFight, FontTypeNames.FONTTYPE_GUILD)
                    End If
                End If
                If .Counters.Pena > 0 Then
                    If .Counters.Pena > 0 Then
                        .Counters.Pena = .Counters.Pena - 1
                        If .Counters.Pena < 1 Then
                            .Counters.Pena = 0
                            Call WarpUserChar(i, Libertad.Map, Libertad.X, Libertad.Y, True)
                            Call WriteConsoleMsg(i, "Has sido liberado!", FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                End If
                If Lloviendo Then Call EfectoLluvia(i)
                If Not .Pos.Map = 0 Then
                    If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ANTIPIQUETE Then
                            If .flags.Muerto = 0 Then
                                .Counters.PiqueteC = .Counters.PiqueteC + 1
                                .Counters.ContadorPiquete = .Counters.ContadorPiquete + 1
                                If .Counters.ContadorPiquete = 6 Then
                                    Call WriteConsoleMsg(i, "Estas obstruyendo la via publica, muevete o seras encarcelado!!!", FontTypeNames.FONTTYPE_INFO)
                                    .Counters.ContadorPiquete = 0
                                End If
                                If .Counters.PiqueteC >= 30 Then
                                    .Counters.PiqueteC = 0
                                    .Counters.ContadorPiquete = 0
                                    Call Encarcelar(i, MinutosCarcelPiquete)
                                End If
                        Else
                            .Counters.PiqueteC = 0
                        End If
                    Else
                        .Counters.PiqueteC = 0
                    End If
                End If
            End If
        End With
    Next i
    Exit Sub
ErrorHandler:
    Call LogError("Error en PasarSegundo. Err: " & Err.description & " - " & Err.Number & " - UserIndex: " & i)
    Resume Next
End Sub
