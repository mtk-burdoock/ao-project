Attribute VB_Name = "mPooMap"
Option Explicit

Private Const GrhFogata As Long = 1521

Public Sub Map_RemoveOldUser()
      With MapData(UserPos.X, UserPos.Y)
            If (.CharIndex = UserCharIndex) Then
                  .CharIndex = 0
            End If
      End With
End Sub

Public Sub Map_CreateObject(ByVal X As Byte, ByVal Y As Byte, ByVal GrhIndex As Long)
      If Not GrhCheck(GrhIndex) Then
            Exit Sub
      End If
      If (Map_InBounds(X, Y)) Then
            With MapData(X, Y)
                  Call InitGrh(.ObjGrh, GrhIndex)
            End With
      End If
End Sub

Public Sub Map_DestroyObject(ByVal X As Byte, ByVal Y As Byte)
      If (Map_InBounds(X, Y)) Then
            With MapData(X, Y)
                  .OBJInfo.ObjIndex = 0
                  .OBJInfo.Amount = 0
                  Call GrhUninitialize(.ObjGrh)
            End With
      End If
End Sub

Public Function Map_PosExitsObject(ByVal X As Byte, ByVal Y As Byte) As Integer
      If (Map_InBounds(X, Y)) Then
            Map_PosExitsObject = MapData(X, Y).ObjGrh.GrhIndex
      Else
            Map_PosExitsObject = 0
      End If
End Function

Public Function Map_GetBlocked(ByVal X As Integer, ByVal Y As Integer) As Boolean
      If (Map_InBounds(X, Y)) Then
            Map_GetBlocked = (MapData(X, Y).Blocked)
      End If
End Function

Public Sub Map_SetBlocked(ByVal X As Byte, ByVal Y As Byte, ByVal block As Byte)
      If (Map_InBounds(X, Y)) Then
            MapData(X, Y).Blocked = block
      End If
End Sub

Sub Map_MoveTo(ByVal Direccion As E_Heading)
      Dim LegalOk As Boolean
      Static lastmovement As Long
      If Cartel Then Cartel = False
      Select Case Direccion
            Case E_Heading.NORTH
                  LegalOk = Map_LegalPos(UserPos.X, UserPos.Y - 1)

            Case E_Heading.EAST
                  LegalOk = Map_LegalPos(UserPos.X + 1, UserPos.Y)

            Case E_Heading.SOUTH
                  LegalOk = Map_LegalPos(UserPos.X, UserPos.Y + 1)

            Case E_Heading.WEST
                  LegalOk = Map_LegalPos(UserPos.X - 1, UserPos.Y)
      End Select
      If LegalOk And Not UserParalizado And Not UserDescansar And Not UserMeditar Then
          Call WriteWalk(Direccion)
          Call frmMain.ActualizarMiniMapa
          Call Char_MovebyHead(UserCharIndex, Direccion)
          Call Char_MoveScreen(Direccion)
      Else
        If (charlist(UserCharIndex).Heading <> Direccion) Then
            If MainTimer.Check(TimersIndex.ChangeHeading) Then
                Call WriteChangeHeading(Direccion)
                Call Char_SetHeading(UserCharIndex, Direccion)
            End If
        End If
      End If
      If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
      If frmMain.trainingMacro.Enabled Then Call frmMain.DesactivarMacroHechizos
      Call Audio.MoveListener(UserPos.X, UserPos.Y)
      If UserMeditar Then
        UserMeditar = Not UserMeditar
      End If
      If UserDescansar Then
        UserDescansar = Not UserDescansar
      End If
End Sub

Function Map_LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
      Dim CharIndex As Integer
      If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
            Exit Function
      End If
      If (Map_GetBlocked(X, Y)) Then
            Exit Function
      End If
      CharIndex = (Char_MapPosExits(CByte(X), CByte(Y)))
      If (CharIndex > 0) Then
            If (Map_GetBlocked(UserPos.X, UserPos.Y)) Then
                  Exit Function
            End If
            With charlist(CharIndex)
                  If .iHead <> eCabezas.CASPER_HEAD And .iBody <> eCabezas.FRAGATA_FANTASMAL Then
                        Exit Function
                  Else
                        If (Map_CheckWater(UserPos.X, UserPos.Y)) Then
                              If Not (Map_CheckWater(X, Y)) Then
                                    Exit Function
                              End If
                        Else
                              If (Map_CheckWater(X, Y)) Then
                                    Exit Function
                              End If
                        End If
                        If (EsGM(UserCharIndex)) Then
                              If (charlist(UserCharIndex).invisible) Then
                                    Exit Function
                              End If
                        End If
                  End If
            End With
      End If
      If (UserNavegando <> Map_CheckWater(X, Y)) Then
            Exit Function
      End If
      If UserEquitando Then
            If MapData(X, Y).Trigger = eTrigger.BAJOTECHO Or _
               MapData(X, Y).Trigger = eTrigger.CASA Or _
               mapInfo.Zona = "DUNGEON" Then
                  If Not frmMain.MsgTimeadoOn Then
                        frmMain.MsgTimeadoOn = True
                        frmMain.MsgTimeado = JsonLanguage.item("MENSAJE_MONTURA_SALIR").item("TEXTO")
                  End If
                  Exit Function
            End If
      End If
      If UserEvento Then Exit Function
      Map_LegalPos = True
End Function

Function Map_InBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
      If (X < XMinMapSize) Or (X > XMaxMapSize) Or (Y < YMinMapSize) Or (Y > YMaxMapSize) Then
            Map_InBounds = False
            Exit Function
      End If
      Map_InBounds = True
End Function

Public Function Map_CheckBonfire(ByRef Location As Position) As Boolean
      Dim J As Long
      Dim k As Long
      For J = UserPos.X - 8 To UserPos.X + 8
            For k = UserPos.Y - 6 To UserPos.Y + 6
                  If Map_InBounds(J, k) Then
                        If MapData(J, k).ObjGrh.GrhIndex = GrhFogata Then
                              Location.X = J
                              Location.Y = k
                              Map_CheckBonfire = True
                              Exit Function
                        End If
                  End If
            Next k
      Next J
End Function

Function Map_CheckWater(ByVal X As Integer, ByVal Y As Integer) As Boolean
      If Map_InBounds(X, Y) Then
            With MapData(X, Y)
                  If ((.Graphic(1).GrhIndex >= 1505 And .Graphic(1).GrhIndex <= 1520) Or (.Graphic(1).GrhIndex >= 5665 And .Graphic(1).GrhIndex <= 5680) Or (.Graphic(1).GrhIndex >= 13547 And .Graphic(1).GrhIndex <= 13562)) And .Graphic(2).GrhIndex = 0 Then
                        Map_CheckWater = True
                  Else
                        Map_CheckWater = False
                  End If
            End With
      End If
End Function
