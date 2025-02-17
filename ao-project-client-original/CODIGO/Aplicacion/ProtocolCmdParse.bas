Attribute VB_Name = "ProtocolCmdParse"
Option Explicit

Public Enum eNumber_Types
    ent_Byte
    ent_Integer
    ent_Long
    ent_Trigger
End Enum

Public Sub AuxWriteWhisper(ByVal UserName As String, ByVal Mensaje As String)
    If LenB(UserName) = 0 Then Exit Sub
    If (InStrB(UserName, "+") <> 0) Then
        UserName = Replace$(UserName, "+", " ")
    End If
    UserName = UCase$(UserName)
    Call WriteWhisper(UserName, Mensaje)
End Sub

Public Sub ParseUserCommand(ByVal RawCommand As String)
    Dim TmpArgos() As String
    Dim Comando As String
    Dim ArgumentosAll() As String
    Dim ArgumentosRaw As String
    Dim Argumentos2() As String
    Dim Argumentos3() As String
    Dim Argumentos4() As String
    Dim CantidadArgumentos As Long
    Dim notNullArguments As Boolean
    Dim tmpArr() As String
    Dim tmpInt As Integer
    TmpArgos = Split(RawCommand, " ", 2)
    Comando = Trim$(UCase$(TmpArgos(0)))
    If UBound(TmpArgos) > 0 Then
        ArgumentosRaw = TmpArgos(1)
        notNullArguments = LenB(Trim$(ArgumentosRaw))
        ArgumentosAll = Split(TmpArgos(1), " ")
        CantidadArgumentos = UBound(ArgumentosAll) + 1
        Argumentos2 = Split(TmpArgos(1), " ", 2)
        Argumentos3 = Split(TmpArgos(1), " ", 3)
        Argumentos4 = Split(TmpArgos(1), " ", 4)
    Else
        CantidadArgumentos = 0
    End If
    If LenB(Comando) = 0 Then Comando = " "
    If Left$(Comando, 1) = "/" Then
        Select Case Comando
            Case "/ONLINE"
                Call WriteOnline
                
            Case "/FADD"
                If notNullArguments Then
                    Call WriteAddAmigo(ArgumentosRaw, 2)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("FRM_LISTAAMIGOS_PARAMETRO").item("TEXTO") & " " & JsonLanguage.item("FRM_LISTAAMIGOS_FADD").item("TEXTO"))
                End If
 
            Case "/FMSG"
                If notNullArguments Then
                    Call WriteMsgAmigo(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("FRM_LISTAAMIGOS_PARAMETRO").item("TEXTO") & " " & JsonLanguage.item("FRM_LISTAAMIGOS_FMSG").item("TEXTO"))
                End If

            Case "/FON"
                Call WriteOnAmigo
  
            Case "/DISCORD"
                If CantidadArgumentos > 0 Then
                    Call WriteDiscord(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                End If
                
            Case "/SALIR"
                If UserParalizado Then
                    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_NO_SALIR").item("TEXTO"), .Red, .Green, .Blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                If frmMain.trainingMacro.Enabled Then Call frmMain.DesactivarMacroHechizos
                If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
                Call WriteQuit
                
            Case "/SALIRCLAN"
                Call WriteGuildLeave
                
            Case "/BALANCE"
                If UserEstado = 1 Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
                    End With
                    Exit Sub
                End If
                Call WriteRequestAccountState
                
            Case "/QUIETO"
                If UserEstado = 1 Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
                    End With
                    Exit Sub
                End If
                Call WritePetStand
                
            Case "/ACOMPANAR"
                If UserEstado = 1 Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
                    End With
                    Exit Sub
                End If
                Call WritePetFollow
                
            Case "/LIBERAR"
                If UserEstado = 1 Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
                    End With
                    Exit Sub
                End If
                Call WriteReleasePet
                
            Case "/ENTRENAR"
                If UserEstado = 1 Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
                    End With
                    Exit Sub
                End If
                Call WriteTrainList
                
            Case "/DESCANSAR"
                If UserEstado = 1 Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
                    End With
                    Exit Sub
                End If
                Call WriteRest
                
            Case "/MEDITAR"
                If UserMinMAN = UserMaxMAN Then Exit Sub
                
                If UserEstado = 1 Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
                    End With
                    Exit Sub
                End If
                Call WriteMeditate
        
            Case "/CONSULTA"
                Call WriteConsultation
            
            Case "/RESUCITAR"
                Call WriteResucitate
                
            Case "/CURAR"
                Call WriteHeal
                              
            Case "/EST"
                Call WriteRequestStats
            
            Case "/AYUDA"
                Call WriteHelp
                
            Case "/COMERCIAR"
                If UserEstado = 1 Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
                    End With
                    Exit Sub
                
                ElseIf Comerciando Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_COMERCIANDO").item("TEXTO"), .Red, .Green, .Blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WriteCommerceStart
                
            Case "/BOVEDA"
                If UserEstado = 1 Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
                    End With
                    Exit Sub
                End If
                Call WriteBankStart
                
            Case "/ENLISTAR"
                Call WriteEnlist
                    
            Case "/INFORMACION"
                Call WriteInformation
                
            Case "/RECOMPENSA"
                Call WriteReward
                
            Case "/MOTD"
                Call WriteRequestMOTD
                
            Case "/UPTIME"
                Call WriteUpTime
                
            Case "/SALIRPARTY"
                Call WritePartyLeave
                
            Case "/CREARPARTY"
                If UserEstado = 1 Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
                    End With
                    Exit Sub
                End If
                Call WritePartyCreate
                
            Case "/PARTY"
                If UserEstado = 1 Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
                    End With
                    Exit Sub
                End If
                Call WritePartyJoin
            
            Case "/COMPARTIRNPC"
                If UserEstado = 1 Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
                    End With
                    Exit Sub
                End If
                
                Call WriteShareNpc
                
            Case "/NOCOMPARTIRNPC"
                If UserEstado = 1 Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
                    End With
                    Exit Sub
                End If
                
                Call WriteStopSharingNpc
                
            Case "/ENCUESTA"
                If CantidadArgumentos = 0 Then
                    Call WriteInquiry
                Else
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Byte) Then
                        Call WriteInquiryVote(ArgumentosRaw)
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_ENCUESTA").item("TEXTO"))
                    End If
                End If
        
            Case "/CMSG"
                If CantidadArgumentos > 0 Then
                    Call WriteGuildMessage(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                End If
        
            Case "/PMSG"
                If CantidadArgumentos > 0 Then
                    Call WritePartyMessage(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                End If

            Case "/CENTINELA"
                If notNullArguments Then
                   Call WriteCentinelReport(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_CENTINELA").item("TEXTO"))
                End If
        
            Case "/ONLINECLAN"
                Call WriteGuildOnline
                
            Case "/ONLINEPARTY"
                Call WritePartyOnline
                
            Case "/BMSG"
                If notNullArguments Then
                    Call WriteCouncilMessage(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                End If
                
            Case "/ROL"
                If notNullArguments Then
                    Call WriteRoleMasterRequest(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_ASK").item("TEXTO"))
                End If
                
            Case "/GM"
                Call WriteGMRequest
                
            Case "/_BUG"
                If notNullArguments Then
                    Call WriteBugReport(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_BUG").item("TEXTO"))
                End If
            
            Case "/DESC"
                If UserEstado = 1 Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
                    End With
                    Exit Sub
                End If
                
                Call WriteChangeDescription(ArgumentosRaw)
            
            Case "/VOTO"
                If notNullArguments Then
                    Call WriteGuildVote(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /voto NICKNAME.")
                End If
               
            Case "/PENAS"
                If notNullArguments Then
                    Call WritePunishments(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /penas NICKNAME.")
                End If
                
            Case "/CONTRASENA"
                Call frmNewPassword.Show(vbModal, frmMain)
            
            Case "/APOSTAR"
                If UserEstado = 1 Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
                    End With
                    Exit Sub
                End If
                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Integer) Then
                        Call WriteGamble(ArgumentosRaw)
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_CANTIDAD_INCORRECTA").item("TEXTO") & " /apostar CANTIDAD.")
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /apostar CANTIDAD.")
                End If
                
            Case "/RETIRARFACCION"
                If UserEstado = 1 Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
                    End With
                    Exit Sub
                End If
                Call WriteLeaveFaction
                
            Case "/RETIRAR"
                If UserEstado = 1 Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
                    End With
                    Exit Sub
                End If
                If notNullArguments Then

                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Long) Then
                        Call WriteBankExtractGold(ArgumentosRaw)
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_CANTIDAD_INCORRECTA").item("TEXTO") & " /retirar CANTIDAD.")
                    End If
                End If

            Case "/DEPOSITAR"
                If UserEstado = 1 Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                            JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
                    End With
                    Exit Sub
                End If

                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Long) Then
                        Call WriteBankDepositGold(ArgumentosRaw)
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_CANTIDAD_INCORRECTA").item("TEXTO") & " /depositar CANTIDAD.")
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /depositar CANTIDAD.")
                End If
                
            Case "/DENUNCIAR"
                If notNullArguments Then
                    Call WriteDenounce(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg("Formule su denuncia.")
                End If
                
            Case "/FUNDARCLAN"
                If UserLvl >= 25 Then
                    Call WriteGuildFundate
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FUNDAR_CLAN").item("TEXTO"))
                End If
            
            Case "/FUNDARCLANGM"
                Call WriteGuildFundation(eClanType.ct_GM)
            
            Case "/ECHARPARTY"
                If notNullArguments Then
                    Call WritePartyKick(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /echarparty NICKNAME.")
                End If
                
            Case "/PARTYLIDER"
                If notNullArguments Then
                    Call WritePartySetLeader(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /partylider NICKNAME.")
                End If
                
            Case "/ACCEPTPARTY"
                If notNullArguments Then
                    Call WritePartyAcceptMember(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /acceptparty NICKNAME.")
                End If

            Case "/BUSCAR"
                frmBuscar.Show vbModeless, frmMain
            
            Case "/LIMPIARMUNDO"
                Call WriteLimpiarMundo
            
            Case "/GMSG"
                If notNullArguments Then
                    Call WriteGMMessage(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                End If
                
            Case "/SHOWNAME"
                Call WriteShowName
                
            Case "/ONLINEREAL"
                Call WriteOnlineRoyalArmy
                
            Case "/ONLINECAOS"
                Call WriteOnlineChaosLegion
                
            Case "/IRCERCA"
                If notNullArguments Then
                    Call WriteGoNearby(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /ircerca NICKNAME.")
                End If
                
            Case "/REM"
                If notNullArguments Then
                    Call WriteComment(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_COMENTARIO").item("TEXTO"))
                End If
            
            Case "/HORA"
                Call Protocol.WriteServerTime
            
            Case "/DONDE"
                If notNullArguments Then
                    Call WriteWhere(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /donde NICKNAME.")
                End If
                
            Case "/NENE"
                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Integer) Then
                        Call WriteCreaturesInMap(ArgumentosRaw)
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_MAPA_INCORRECTO").item("TEXTO") & " /nene MAPA.")
                    End If
                Else
                    Call WriteCreaturesInMap(UserMap)
                End If
                
            Case "/TELEPLOC"
                Call WriteWarpMeToTarget
                
            Case "/TELEP"
                If notNullArguments And CantidadArgumentos >= 4 Then
                    If ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(3), eNumber_Types.ent_Byte) Then
                        Call WriteWarpChar(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), ArgumentosAll(3))
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /telep NICKNAME MAPA X Y.")
                    End If
                ElseIf CantidadArgumentos = 3 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                        Call WriteWarpChar("YO", ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                    ElseIf ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                        Call WriteWarpChar(ArgumentosAll(0), UserMap, ArgumentosAll(1), ArgumentosAll(2))
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /telep NICKNAME MAPA X Y.")
                    End If
                ElseIf CantidadArgumentos = 2 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) Then
                        Call WriteWarpChar("YO", UserMap, ArgumentosAll(0), ArgumentosAll(1))
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /telep NICKNAME MAPA X Y.")
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /telep NICKNAME MAPA X Y.")
                End If
                
            Case "/SILENCIAR"
                If notNullArguments Then
                    Call WriteSilence(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /silenciar NICKNAME.")
                End If
                
            Case "/SHOW"
                If notNullArguments Then
                    Select Case UCase$(ArgumentosAll(0))
                        Case "SOS"
                            Call WriteSOSShowList
                            
                        Case "INT"
                            Call WriteShowServerForm
                        
                        Case "DENUNCIAS"
                            Call WriteShowDenouncesList
                    End Select
                End If
                
            Case "/DENUNCIAS"
                Call WriteEnableDenounces
                
            Case "/IRA"
                If notNullArguments Then
                    Call WriteGoToChar(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /ira NICKNAME.")
                End If
        
            Case "/INVISIBLE"
                Call WriteInvisible
                
            Case "/PANELGM"
                Call WriteGMPanel
                
            Case "/TRABAJANDO"
                Call WriteWorking
                
            Case "/OCULTANDO"
                Call WriteHiding
                
            Case "/CARCEL"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@")
                    If UBound(tmpArr) = 2 Then
                        If ValidNumber(tmpArr(2), eNumber_Types.ent_Byte) Then
                            Call WriteJail(tmpArr(0), tmpArr(1), tmpArr(2))
                        Else
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_TIEMPO_INCORRECTO").item("TEXTO") & " /carcel NICKNAME@MOTIVO@TIEMPO.")
                        End If
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /carcel NICKNAME@MOTIVO@TIEMPO.")
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /carcel NICKNAME@MOTIVO@TIEMPO.")
                End If
                
            Case "/RMATA"
                Call WriteKillNPC
                
            Case "/ADVERTENCIA"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)
                    If UBound(tmpArr) = 1 Then
                        Call WriteWarnUser(tmpArr(0), tmpArr(1))
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /advertencia NICKNAME@MOTIVO.")
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /advertencia NICKNAME@MOTIVO.")
                End If
                
            Case "/MOD"
                If notNullArguments And CantidadArgumentos >= 3 Then
                    Select Case UCase$(ArgumentosAll(1))
                        Case "BODY"
                            tmpInt = eEditOptions.eo_Body
                        
                        Case "HEAD"
                            tmpInt = eEditOptions.eo_Head
                        
                        Case "ORO"
                            tmpInt = eEditOptions.eo_Gold
                        
                        Case "LEVEL"
                            tmpInt = eEditOptions.eo_Level
                        
                        Case "SKILLS"
                            tmpInt = eEditOptions.eo_Skills
                        
                        Case "SKILLSLIBRES"
                            tmpInt = eEditOptions.eo_SkillPointsLeft
                        
                        Case "CLASE"
                            tmpInt = eEditOptions.eo_Class
                        
                        Case "EXP"
                            tmpInt = eEditOptions.eo_Experience
                        
                        Case "CRI"
                            tmpInt = eEditOptions.eo_CriminalsKilled
                        
                        Case "CIU"
                            tmpInt = eEditOptions.eo_CiticensKilled
                        
                        Case "NOB"
                            tmpInt = eEditOptions.eo_Nobleza
                        
                        Case "ASE"
                            tmpInt = eEditOptions.eo_Asesino
                        
                        Case "SEX"
                            tmpInt = eEditOptions.eo_Sex
                            
                        Case "RAZA"
                            tmpInt = eEditOptions.eo_Raza
                        
                        Case "AGREGAR"
                            tmpInt = eEditOptions.eo_addGold
                        
                        Case "VIDA"
                            tmpInt = eEditOptions.eo_Vida
                         
                        Case "POSS"
                            tmpInt = eEditOptions.eo_Poss
                         
                        Case Else
                            tmpInt = -1
                    End Select
                    If tmpInt > 0 Then
                        If CantidadArgumentos = 3 Then
                            Call WriteEditChar(ArgumentosAll(0), tmpInt, ArgumentosAll(2), vbNullString)
                        Else
                            Call WriteEditChar(ArgumentosAll(0), tmpInt, ArgumentosAll(2), ArgumentosAll(3))
                        End If
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_COMANDO_INCORRECTO").item("TEXTO"))
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO"))
                End If
            
            Case "/INFO"
                If notNullArguments Then
                    Call WriteRequestCharInfo(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /info NICKNAME.")
                End If
                
            Case "/STAT"
                If notNullArguments Then
                    Call WriteRequestCharStats(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /stat NICKNAME.")
                End If
                
            Case "/BAL"
                If notNullArguments Then
                    Call WriteRequestCharGold(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /bal NICKNAME.")
                End If
                
            Case "/INV"
                If notNullArguments Then
                    Call WriteRequestCharInventory(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /inv NICKNAME.")
                End If
                
            Case "/BOV"
                If notNullArguments Then
                    Call WriteRequestCharBank(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /bov NICKNAME.")
                End If
                
            Case "/SKILLS"
                If notNullArguments Then
                    Call WriteRequestCharSkills(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /skills NICKNAME.")
                End If
                
            Case "/REVIVIR"
                If notNullArguments Then
                    Call WriteReviveChar(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /revivir NICKNAME.")
                End If
                
            Case "/ONLINEGM"
                Call WriteOnlineGM
                
            Case "/ONLINEMAP"
                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                        Call WriteOnlineMap(ArgumentosAll(0))
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_MAPA_INCORRECTO").item("TEXTO") & " /ONLINEMAP")
                    End If
                Else
                    Call WriteOnlineMap(UserMap)
                End If
                
            Case "/PERDON"
                If notNullArguments Then
                    Call WriteForgive(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /perdon NICKNAME.")
                End If
                
            Case "/ECHAR"
                If notNullArguments Then
                    Call WriteKick(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /echar NICKNAME.")
                End If
                
            Case "/EJECUTAR"
                If notNullArguments Then
                    Call WriteExecute(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /ejecutar NICKNAME.")
                End If
                
            Case "/BAN"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)
                    If UBound(tmpArr) = 1 Then
                        Call WriteBanChar(tmpArr(0), tmpArr(1))
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /ban NICKNAME@MOTIVO.")
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /ban NICKNAME@MOTIVO.")
                End If
                
            Case "/UNBAN"
                If notNullArguments Then
                    Call WriteUnbanChar(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /unban NICKNAME.")
                End If
                
            Case "/SEGUIR"
                Call WriteNPCFollow
                
            Case "/SUM"
                If notNullArguments Then
                    Call WriteSummonChar(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /sum NICKNAME.")
                End If
                
            Case "/CC"
                Call WriteSpawnListRequest
                
            Case "/RESETINV"
                Call WriteResetNPCInventory
                
            Case "/RMSG"
                If notNullArguments Then
                    Call WriteServerMessage(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                End If
            
            Case "/MAPMSG"
                If notNullArguments Then
                    Call WriteMapMessage(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                End If
                
            Case "/NICK2IP"
                If notNullArguments Then
                    Call WriteNickToIP(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /nick2ip NICKNAME.")
                End If
                
            Case "/IP2NICK"
                If notNullArguments Then
                    If validipv4str(ArgumentosRaw) Then
                        Call WriteIPToNick(str2ipv4l(ArgumentosRaw))
                    Else
                        Call ShowConsoleMsg("IP incorrecta. Utilice /ip2nick IP.")
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /ip2nick IP.")
                End If
                
            Case "/ONCLAN"
                If notNullArguments Then
                    Call WriteGuildOnlineMembers(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_ONCLAN").item("TEXTO"))
                End If
                
            Case "/CT"
                If notNullArguments And CantidadArgumentos >= 3 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And _
                        ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                        
                        If CantidadArgumentos = 3 Then
                            Call WriteTeleportCreate(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                        Else
                            If ValidNumber(ArgumentosAll(3), eNumber_Types.ent_Byte) Then
                                Call WriteTeleportCreate(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), ArgumentosAll(3))
                            Else
                                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /ct MAPA X Y RADIO(Opcional).")
                            End If
                        End If
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /ct MAPA X Y RADIO(Opcional).")
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /ct MAPA X Y RADIO(Opcional).")
                End If
                
            Case "/DT"
                Call WriteTeleportDestroy
                
            Case "/DE"
                Call WriteExitDestroy
                
            Case "/LLUVIA"
                Call WriteRainToggle
                
            Case "/SETDESC"
                Call WriteSetCharDescription(ArgumentosRaw)

            Case "/FORCEMP3MAP"
                If notNullArguments Then
                    If CantidadArgumentos = 1 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                            Call WriteForceMP3ToMap(ArgumentosAll(0), 0)
                        Else
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /forcemp3map MP3 MAPA")
                        End If
                    Else
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                            Call WriteForceMP3ToMap(ArgumentosAll(0), ArgumentosAll(1))
                        Else
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /forcemp3map MP3 MAPA")
                        End If
                    End If
                Else
                    Call ShowConsoleMsg("Utilice /forcemp3map MP3 MAPA")
                End If
            
            Case "/FORCEMIDIMAP"
                If notNullArguments Then
                    If CantidadArgumentos = 1 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                            Call WriteForceMIDIToMap(ArgumentosAll(0), 0)
                        Else
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /forcemidimap MIDI MAPA")
                        End If
                    Else
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                            Call WriteForceMIDIToMap(ArgumentosAll(0), ArgumentosAll(1))
                        Else
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /forcemidimap MIDI MAPA")
                        End If
                    End If
                Else
                    Call ShowConsoleMsg("Utilice /forcemidimap MIDI MAPA")
                End If
                
            Case "/FORCEWAVMAP"
                If notNullArguments Then
                    If CantidadArgumentos = 1 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                            Call WriteForceWAVEToMap(ArgumentosAll(0), 0, 0, 0)
                        Else
                            Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los ultimos 3 opcionales.")
                        End If
                    ElseIf CantidadArgumentos = 4 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(3), eNumber_Types.ent_Byte) Then
                            Call WriteForceWAVEToMap(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), ArgumentosAll(3))
                        Else
                            Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los ultimos 3 opcionales.")
                        End If
                    Else
                        Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los ultimos 3 opcionales.")
                    End If
                Else
                    Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los ultimos 3 opcionales.")
                End If
                
            Case "/REALMSG"
                If notNullArguments Then
                    Call WriteRoyalArmyMessage(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                End If
                 
            Case "/CAOSMSG"
                If notNullArguments Then
                    Call WriteChaosLegionMessage(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                End If
                
            Case "/CIUMSG"
                If notNullArguments Then
                    Call WriteCitizenMessage(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                End If
            
            Case "/CRIMSG"
                If notNullArguments Then
                    Call WriteCriminalMessage(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                End If
            
            Case "/TALKAS"
                If notNullArguments Then
                    Call WriteTalkAsNPC(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                End If
        
            Case "/MASSDEST"
                Call WriteDestroyAllItemsInArea
    
            Case "/ACEPTCONSE"
                If notNullArguments Then
                    Call WriteAcceptRoyalCouncilMember(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /aceptconse NICKNAME.")
                End If
                
            Case "/ACEPTCONSECAOS"
                If notNullArguments Then
                    Call WriteAcceptChaosCouncilMember(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /aceptconsecaos NICKNAME.")
                End If
                
            Case "/PISO"
                Call WriteItemsInTheFloor
                
            Case "/ESTUPIDO"
                If notNullArguments Then
                    Call WriteMakeDumb(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /estupido NICKNAME.")
                End If
                
            Case "/NOESTUPIDO"
                If notNullArguments Then
                    Call WriteMakeDumbNoMore(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /noestupido NICKNAME.")
                End If
                
            Case "/DUMPSECURITY"
                Call WriteDumpIPTables
                
            Case "/KICKCONSE"
                If notNullArguments Then
                    Call WriteCouncilKick(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /kickconse NICKNAME.")
                End If
                
            Case "/TRIGGER"
                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Trigger) Then
                        Call WriteSetTrigger(ArgumentosRaw)
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /trigger NUMERO.")
                    End If
                Else
                    Call WriteAskTrigger
                End If
                
            Case "/BANIPLIST"
                Call WriteBannedIPList
                
            Case "/BANIPRELOAD"
                Call WriteBannedIPReload
                
            Case "/MIEMBROSCLAN"
                If notNullArguments Then
                    Call WriteGuildMemberList(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /miembrosclan GUILDNAME.")
                End If
                
            Case "/BANCLAN"
                If notNullArguments Then
                    Call WriteGuildBan(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /banclan GUILDNAME.")
                End If
                
            Case "/BANIP"
                If CantidadArgumentos >= 2 Then
                    If validipv4str(ArgumentosAll(0)) Then
                        Call WriteBanIP(True, str2ipv4l(ArgumentosAll(0)), vbNullString, Right$(ArgumentosRaw, Len(ArgumentosRaw) - Len(ArgumentosAll(0)) - 1))
                    Else
                        Call WriteBanIP(False, str2ipv4l("0.0.0.0"), ArgumentosAll(0), Right$(ArgumentosRaw, Len(ArgumentosRaw) - Len(ArgumentosAll(0)) - 1))
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /banip IP motivo o /banip nick motivo.")
                End If
                
            Case "/UNBANIP"
                If notNullArguments Then
                    If validipv4str(ArgumentosRaw) Then
                        Call WriteUnbanIP(str2ipv4l(ArgumentosRaw))
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /unbanip IP.")
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /unbanip IP.")
                End If
                
            Case "/CI"
                If notNullArguments And CantidadArgumentos = 2 Then
                    If IsNumeric(ArgumentosAll(0)) And IsNumeric(ArgumentosAll(1)) Then
                        Call WriteCreateItem(ArgumentosAll(0), ArgumentosAll(1))
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_OBJETO_INCORRECTO").item("TEXTO") & " /CI " & JsonLanguage.item("OBJETO").item("TEXTO") & " " & JsonLanguage.item("CANTIDAD").item("TEXTO"))
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /ci OBJETO CANTIDAD.")
                End If
                
            Case "/DEST"
                Call WriteDestroyItems
                
            Case "/NOCAOS"
                If notNullArguments Then
                    Call WriteChaosLegionKick(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /nocaos NICKNAME.")
                End If
    
            Case "/NOREAL"
                If notNullArguments Then
                    Call WriteRoyalArmyKick(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /noreal NICKNAME.")
                End If

            Case "/FORCEMP3"
                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                        Call WriteForceMP3All(ArgumentosAll(0))
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_MP3_INCORRECTO").item("TEXTO") & " /forcemp3 MP3.")
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /forcemp3 MP3.")
                End If
    
            Case "/FORCEMIDI"
                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                        Call WriteForceMIDIAll(ArgumentosAll(0))
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_MIDI_INCORRECTO").item("TEXTO") & " /forcemidi MIDI.")
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /forcemidi MIDI.")
                End If
    
            Case "/FORCEWAV"
                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                        Call WriteForceWAVEAll(ArgumentosAll(0))
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_WAV_INCORRECTO").item("TEXTO") & " /forcewav WAV.")
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /forcewav WAV.")
                End If
                
            Case "/MODIFICARPENA"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 3)
                    If UBound(tmpArr) = 2 Then
                        Call WriteRemovePunishment(tmpArr(0), tmpArr(1), tmpArr(2))
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /borrarpena NICK@PENA@NuevaPena.")
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /borrarpena NICK@PENA@NuevaPena.")
                End If
                
            Case "/BLOQ"
                Call WriteTileBlockedToggle
                
            Case "/MATA"
                Call WriteKillNPCNoRespawn
        
            Case "/MASSKILL"
                Call WriteKillAllNearbyNPCs
                
            Case "/LASTIP"
                If notNullArguments Then
                    Call WriteLastIP(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /lastip NICKNAME.")
                End If
    
            Case "/MOTDCAMBIA"
                Call WriteChangeMOTD
                
            Case "/SMSG"
                If notNullArguments Then
                    Call WriteSystemMessage(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                End If
                
            Case "/ACC"
                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                        Call WriteCreateNPC(ArgumentosAll(0), False)
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_NPC_INCORRECTO").item("TEXTO") & " /ACC NPC.")
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /ACC NPC.")
                End If
                
            Case "/RACC"
                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                        Call WriteCreateNPC(ArgumentosAll(0), True)
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_NPC_INCORRECTO").item("TEXTO") & " /RACC NPC.")
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /RACC NPC.")
                End If
        
            Case "/AI"
                If notNullArguments And CantidadArgumentos >= 2 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                        Call WriteImperialArmour(ArgumentosAll(0), ArgumentosAll(1))
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /ai ARMADURA OBJETO.")
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /ai ARMADURA OBJETO.")
                End If
                
            Case "/AC"
                If notNullArguments And CantidadArgumentos >= 2 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                        Call WriteChaosArmour(ArgumentosAll(0), ArgumentosAll(1))
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /ac ARMADURA OBJETO.")
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /ac ARMADURA OBJETO.")
                End If
                
            Case "/NAVE"
                Call WriteNavigateToggle
        
            Case "/HABILITAR"
                Call WriteServerOpenToUsersToggle
            
            Case "/APAGAR"
                Call WriteTurnOffServer
                
            Case "/CONDEN"
                If notNullArguments Then
                    Call WriteTurnCriminal(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /conden NICKNAME.")
                End If
                
            Case "/RAJAR"
                If notNullArguments Then
                    Call WriteResetFactions(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /rajar NICKNAME.")
                End If
                
            Case "/RAJARCLAN"
                If notNullArguments Then
                    Call WriteRemoveCharFromGuild(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /rajarclan NICKNAME.")
                End If
                
            Case "/LASTEMAIL"
                If notNullArguments Then
                    Call WriteRequestCharMail(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /lastemail NICKNAME.")
                End If
                
            Case "/APASS"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)
                    If UBound(tmpArr) = 1 Then
                        Call WriteAlterPassword(tmpArr(0), tmpArr(1))
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /apass PJSINPASS@PJCONPASS.")
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /apass PJSINPASS@PJCONPASS.")
                End If
                
            Case "/AEMAIL"
                If notNullArguments Then
                    tmpArr = AEMAILSplit(ArgumentosRaw)
                    If LenB(tmpArr(0)) = 0 Then
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /aemail NICKNAME-NUEVOMAIL.")
                    Else
                        Call WriteAlterMail(tmpArr(0), tmpArr(1))
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /aemail NICKNAME-NUEVOMAIL.")
                End If
                
            Case "/ANAME"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)
                    If UBound(tmpArr) = 1 Then
                        Call WriteAlterName(tmpArr(0), tmpArr(1))
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /aname ORIGEN@DESTINO.")
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /aname ORIGEN@DESTINO.")
                End If
                
            Case "/SLOT"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)
                    If UBound(tmpArr) = 1 Then
                        If ValidNumber(tmpArr(1), eNumber_Types.ent_Byte) Then
                            Call WriteCheckSlot(tmpArr(0), tmpArr(1))
                        Else
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /slot NICK@SLOT.")
                        End If
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /slot NICK@SLOT.")
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /slot NICK@SLOT.")
                End If

            Case "/CENTINELAACTIVADO"
                Call WriteToggleCentinelActivated
                
            Case "/CREARPRETORIANOS"
                If CantidadArgumentos = 3 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And _
                       ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And _
                       ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                        Call WriteCreatePretorianClan(Val(ArgumentosAll(0)), Val(ArgumentosAll(1)), Val(ArgumentosAll(2)))
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /CrearPretorianos MAPA X Y.")
                    End If
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /CrearPretorianos MAPA X Y.")
                End If
                
            Case "/ELIMINARPRETORIANOS"
                If CantidadArgumentos = 1 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                        Call WriteDeletePretorianClan(Val(ArgumentosAll(0)))
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /EliminarPretorianos MAPA.")
                    End If
                    
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /EliminarPretorianos MAPA.")
                End If
            
            Case "/DOBACKUP"
                Call WriteDoBackup
                
            Case "/SHOWCMSG"
                If notNullArguments Then
                    Call WriteShowGuildMessages(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /showcmsg GUILDNAME.")
                End If
                
            Case "/GUARDAMAPA"
                Call WriteSaveMap
                
            Case "/MODMAPINFO"
                If CantidadArgumentos > 1 Then
                    Select Case UCase$(ArgumentosAll(0))
                        Case "PK"
                            Call WriteChangeMapInfoPK(ArgumentosAll(1) = "1")
                        
                        Case "BACKUP"
                            Call WriteChangeMapInfoBackup(ArgumentosAll(1) = "1")
                        
                        Case "RESTRINGIR"
                            Call WriteChangeMapInfoRestricted(ArgumentosAll(1))
                        
                        Case "MAGIASINEFECTO"
                            Call WriteChangeMapInfoNoMagic(ArgumentosAll(1) = "1")
                        
                        Case "INVISINEFECTO"
                            Call WriteChangeMapInfoNoInvi(ArgumentosAll(1) = "1")
                        
                        Case "RESUSINEFECTO"
                            Call WriteChangeMapInfoNoResu(ArgumentosAll(1) = "1")
                        
                        Case "TERRENO"
                            Call WriteChangeMapInfoLand(ArgumentosAll(1))
                        
                        Case "ZONA"
                            Call WriteChangeMapInfoZone(ArgumentosAll(1))
                            
                        Case "ROBONPC"
                            Call WriteChangeMapInfoStealNpc(ArgumentosAll(1) = "1")
                            
                        Case "OCULTARSINEFECTO"
                            Call WriteChangeMapInfoNoOcultar(ArgumentosAll(1) = "1")
                            
                        Case "INVOCARSINEFECTO"
                            Call WriteChangeMapInfoNoInvocar(ArgumentosAll(1) = "1")
                    End Select
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " : PK, BACKUP, RESTRINGIR, MAGIASINEFECTO, INVISINEFECTO, RESUSINEFECTO, TERRENO, ZONA")
                End If
                
            Case "/GRABAR"
                Call WriteSaveChars
                
            Case "/BORRAR"
                If notNullArguments Then
                    Select Case UCase$(ArgumentosAll(0))
                        Case "SOS"
                            Call WriteCleanSOS
                    End Select
                End If
                
            Case "/NOCHE"
                Call WriteNight
                
            Case "/ECHARTODOSPJS"
                Call WriteKickAllChars
                
            Case "/RELOADNPCS"
                Call WriteReloadNPCs
                
            Case "/RELOADSINI"
                Call WriteReloadServerIni
                
            Case "/RELOADHECHIZOS"
                Call WriteReloadSpells
                
            Case "/RELOADOBJ"
                Call WriteReloadObjects
                 
            Case "/REINICIAR"
                Call WriteRestart
                
            Case "/AUTOUPDATE"
                Call WriteResetAutoUpdate
            
            Case "/CHATCOLOR"
                If notNullArguments And CantidadArgumentos >= 3 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                        Call WriteChatColor(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /chatcolor R G B.")
                    End If
                ElseIf Not notNullArguments Then
                    Call WriteChatColor(0, 255, 0)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /chatcolor R G B.")
                End If
            
            Case "/IGNORADO"
                Call WriteIgnored
            
            Case "/PING"
                Call WritePing
                
            Case "/RETOS"
                Call FrmRetos.Show(vbModeless, frmMain)
            
            Case "/CERRARCLAN"
                Call WriteCloseGuild
                
            Case "/ACEPTAR"
                If notNullArguments Then
                    Call WriteFightAccept(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /ACEPTAR NICKNAME.")
                End If
                
            Case "/QUEST"
                Call WriteQuest
 
            Case "/INFOQUEST"
                Call WriteQuestListRequest
                
            Case "/SETINIVAR"
                If CantidadArgumentos = 3 Then
                    ArgumentosAll(2) = Replace(ArgumentosAll(2), "+", " ")
                    Call WriteSetIniVar(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /SETINIVAR LLAVE CLAVE VALOR")
                End If
            
            Case "/CVC"
                Call WriteEnviaCvc

            Case "/ACVC"
                Call WriteAceptarCvc

            Case "/IRCVC"
                Call WriteIrCvc
            
            Case "/HOGAR"
                Call WriteHome

            Case "/SETDIALOG"
                If notNullArguments Then
                    Call WriteSetDialog(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /SETDIALOG DIALOGO.")
                End If
            
            Case "/IMPERSONAR"
                Call WriteImpersonate
                
            Case "/MIMETIZAR"
                Call WriteImitate

            Case "/VERPROCESOS"
                If notNullArguments Then
                    Call WriteLookProcess(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /VERPROCESOS NICKNAME.")
                End If
            
            Case Else
                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_COMANDO_INCORRECTO").item("TEXTO"))
        
        End Select
        
    ElseIf Left$(Comando, 1) = "\" Then
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                                    JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                                    JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                                    JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
            End With
            Exit Sub
        End If
        Call AuxWriteWhisper(mid$(Comando, 2), ArgumentosRaw)
    ElseIf Left$(Comando, 1) = "-" Then
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                                    JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                                    JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                                    JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
            End With
            Exit Sub
        End If
        Call WriteYell(mid$(RawCommand, 2))
    Else
        Call WriteTalk(RawCommand)
    End If
End Sub

Public Sub ShowConsoleMsg(ByVal Message As String, Optional ByVal Red As Integer = 255, Optional ByVal Green As Integer = 255, Optional ByVal Blue As Integer = 255, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False)
    Call AddtoRichTextBox(frmMain.RecTxt, Message, Red, Green, Blue, bold, italic)
End Sub

Public Function ValidNumber(ByVal Numero As String, ByVal TIPO As eNumber_Types) As Boolean
    Dim Minimo As Long
    Dim Maximo As Long
    If Not IsNumeric(Numero) Then _
        Exit Function
    Select Case TIPO
        Case eNumber_Types.ent_Byte
            Minimo = 0
            Maximo = 255

        Case eNumber_Types.ent_Integer
            Minimo = -32768
            Maximo = 32767

        Case eNumber_Types.ent_Long
            Minimo = -2147483648#
            Maximo = 2147483647
        
        Case eNumber_Types.ent_Trigger
            Minimo = 0
            Maximo = 6
    End Select
    If Val(Numero) >= Minimo And Val(Numero) <= Maximo Then _
        ValidNumber = True
End Function

Private Function validipv4str(ByVal Ip As String) As Boolean
    Dim tmpArr() As String
    tmpArr = Split(Ip, ".")
    If UBound(tmpArr) <> 3 Then _
        Exit Function
    If Not ValidNumber(tmpArr(0), eNumber_Types.ent_Byte) Or _
      Not ValidNumber(tmpArr(1), eNumber_Types.ent_Byte) Or _
      Not ValidNumber(tmpArr(2), eNumber_Types.ent_Byte) Or _
      Not ValidNumber(tmpArr(3), eNumber_Types.ent_Byte) Then _
        Exit Function
    validipv4str = True
End Function

Private Function str2ipv4l(ByVal Ip As String) As Byte()
    Dim tmpArr() As String
    Dim bArr(3) As Byte
    tmpArr = Split(Ip, ".")
    bArr(0) = CByte(tmpArr(0))
    bArr(1) = CByte(tmpArr(1))
    bArr(2) = CByte(tmpArr(2))
    bArr(3) = CByte(tmpArr(3))
    str2ipv4l = bArr
End Function

Private Function AEMAILSplit(ByRef Text As String) As String()
    Dim tmpArr(0 To 1) As String
    Dim Pos As Byte
    Pos = InStr(1, Text, "-")
    If Pos <> 0 Then
        tmpArr(0) = mid$(Text, 1, Pos - 1)
        tmpArr(1) = mid$(Text, Pos + 1)
    Else
        tmpArr(0) = vbNullString
    End If
    AEMAILSplit = tmpArr
End Function
