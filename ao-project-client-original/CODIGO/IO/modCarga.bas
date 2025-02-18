Attribute VB_Name = "Carga"
Option Explicit

Private FileManager As clsIniManager

Public Sub LoadGrhIni()
    On Error GoTo ErrorHandler
    Dim FileHandle     As Integer
    Dim Grh            As Long
    Dim Frame          As Long
    Dim SeparadorClave As String
    Dim SeparadorGrh   As String
    Dim CurrentLine    As String
    Dim Fields()       As String
    SeparadorClave = "="
    SeparadorGrh = "-"
    FileHandle = FreeFile()
    Open Game.path(INIT) & "Graficos.ini" For Input As FileHandle
    Do While Not EOF(FileHandle)
        Line Input #FileHandle, CurrentLine
        Fields = Split(CurrentLine, SeparadorClave)
        If Fields(0) = "NumGrh" Then
            ReDim GrhData(1 To Val(Fields(1))) As GrhData
            Exit Do
        End If
    Loop
    If UBound(GrhData) <= 0 Then GoTo ErrorHandler
    Do While Not EOF(FileHandle)
        Line Input #FileHandle, CurrentLine
        If UCase$(CurrentLine) = "[GRAPHICS]" Then
            Exit Do
        End If
    Loop
    Do While Not EOF(FileHandle)
        Line Input #FileHandle, CurrentLine
        If CurrentLine <> vbNullString Then
            Fields = Split(CurrentLine, SeparadorClave)
            Grh = Right(Fields(0), Len(Fields(0)) - 3)
            Fields = Split(Fields(1), SeparadorGrh)
            With GrhData(Grh)
                .NumFrames = Val(Fields(0))
                ReDim .Frames(1 To .NumFrames)
                If .NumFrames > 1 Then
                    For Frame = 1 To .NumFrames
                        .Frames(Frame) = Val(Fields(Frame))
                        If .Frames(Frame) <= LBound(GrhData) Or .Frames(Frame) > UBound(GrhData) Then GoTo ErrorHandler
                    Next
                    .Speed = Val(Fields(Frame))
                    If .Speed <= 0 Then GoTo ErrorHandler
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    If .TileWidth <= 0 Then GoTo ErrorHandler
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                    If .TileHeight <= 0 Then GoTo ErrorHandler
                ElseIf .NumFrames = 1 Then
                    .Frames(1) = Grh
                    .FileNum = Val(Fields(1))
                    If .FileNum <= 0 Then GoTo ErrorHandler
                    .sX = Val(Fields(2))
                    If .sX < 0 Then GoTo ErrorHandler
                    .sY = Val(Fields(3))
                    If .sY < 0 Then GoTo ErrorHandler
                    .pixelWidth = Val(Fields(4))
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    .pixelHeight = Val(Fields(5))
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    .TileWidth = .pixelWidth / TilePixelHeight
                    .TileHeight = .pixelHeight / TilePixelWidth
                Else
                    GoTo ErrorHandler
                End If
            End With
        End If
    Loop
ErrorHandler:
    Close FileHandle
    If Err.number <> 0 Then
        If Err.number = 53 Then
            Call MsgBox("El archivo Graficos.ini no existe. Por favor, reinstale el juego.", , "Argentum Online")
        ElseIf Grh > 0 Then
            Call MsgBox("Hay un error en Graficos.ini con el Grh" & Grh & ".", , "Argentum Online")
        Else
            Call MsgBox("Hay un error en Graficos.ini. Por favor, reinstale el juego.", , "Argentum Online")
        End If
        Call CloseClient
    End If
    Exit Sub
End Sub

Public Sub LoadGrhInd()
On Error GoTo ErrorHandler:
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
    handle = FreeFile()
    Open IniPath & "Graficos.ind" For Binary Access Read As handle
        Get handle, , fileVersion
        Get handle, , grhCount
        ReDim GrhData(0 To grhCount) As GrhData
        While Not EOF(handle)
            Get handle, , Grh
            With GrhData(Grh)
                Get handle, , .NumFrames
                If .NumFrames <= 0 Then GoTo ErrorHandler
                ReDim .Frames(1 To .NumFrames)
                If .NumFrames > 1 Then
                    For Frame = 1 To .NumFrames
                        Get handle, , .Frames(Frame)
                        If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then GoTo ErrorHandler
                    Next Frame
                    Get handle, , .Speed
                    If .Speed <= 0 Then GoTo ErrorHandler
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    If .TileWidth <= 0 Then GoTo ErrorHandler
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                    If .TileHeight <= 0 Then GoTo ErrorHandler
                Else
                    Get handle, , .FileNum
                    If .FileNum <= 0 Then GoTo ErrorHandler
                    Get handle, , GrhData(Grh).sX
                    If .sX < 0 Then GoTo ErrorHandler
                    Get handle, , .sY
                    If .sY < 0 Then GoTo ErrorHandler
                    Get handle, , .pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    Get handle, , .pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    .TileWidth = .pixelWidth / TilePixelHeight
                    .TileHeight = .pixelHeight / TilePixelWidth
                    .Frames(1) = Grh
                End If
            End With
        Wend
    Close handle
Exit Sub
ErrorHandler:
    If Err.number <> 0 Then
        If Err.number = 53 Then
            Call MsgBox("El archivo Graficos.ind no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
        End If
    End If
End Sub

Public Sub CargarCabezas()
On Error GoTo ErrorHandler:
    Dim N As Integer
    Dim i As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    N = FreeFile()
    Open Game.path(INIT) & "Cabezas.ind" For Binary Access Read As #N
    Get #N, , MiCabecera
    Get #N, , Numheads
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    Close #N
ErrorHandler:
    If Err.number <> 0 Then
        If Err.number = 53 Then
            Call MsgBox("El archivo Cabezas.ind no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
        End If
    End If
End Sub

Sub CargarCascos()
On Error GoTo ErrorHandler:
    Dim N As Integer
    Dim i As Long
    Dim NumCascos As Integer
    Dim Miscabezas() As tIndiceCabeza
    N = FreeFile()
    Open Game.path(INIT) & "Cascos.ind" For Binary Access Read As #N
    Get #N, , MiCabecera
    Get #N, , NumCascos
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    Close #N
ErrorHandler:
    If Err.number <> 0 Then
        If Err.number = 53 Then
            Call MsgBox("El archivo Cascos.ind no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
        End If
    End If
End Sub

Sub CargarCuerpos()
On Error GoTo ErrorHandler:
    Dim N As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    N = FreeFile()
    Open Game.path(INIT) & "Personajes.ind" For Binary Access Read As #N
    Get #N, , MiCabecera
    Get #N, , NumCuerpos
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        If MisCuerpos(i).Body(1) Then
            Call InitGrh(BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0)
            Call InitGrh(BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0)
            Call InitGrh(BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0)
            Call InitGrh(BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0)
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    Close #N
ErrorHandler:
    If Err.number <> 0 Then
        If Err.number = 53 Then
            Call MsgBox("El archivo Personajes.ind no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
        End If
    End If
End Sub

Sub CargarFxs()
On Error GoTo ErrorHandler:
    Dim i As Long
    Set FileManager = New clsIniManager
    Call FileManager.Initialize(Game.path(INIT) & "Fxs.ini")
    ReDim FxData(0 To FileManager.GetValue("INIT", "NumFxs")) As tIndiceFx
    For i = 1 To UBound(FxData())
        With FxData(i)
            .Animacion = Val(FileManager.GetValue("FX" & CStr(i), "Animacion"))
            .OffsetX = Val(FileManager.GetValue("FX" & CStr(i), "OffsetX"))
            .OffsetY = Val(FileManager.GetValue("FX" & CStr(i), "OffsetY"))
        End With
    Next
    Set FileManager = Nothing
ErrorHandler:
    If Err.number <> 0 Then
        If Err.number = 53 Then
            Call MsgBox("El archivo Fxs.ini no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
        End If
    End If
End Sub

Public Sub CargarTips()
On Error GoTo ErrorHandler:
    Dim TipFile As String
    TipFile = FileToString(Game.path(INIT) & "tips_" & Language & ".json")
    Set JsonTips = JSON.parse(TipFile)
ErrorHandler:
    If Err.number <> 0 Then
        If Err.number = 53 Then
            Call MsgBox("El archivo" & "tips_" & Language & ".json no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
        End If
    End If
End Sub

Sub CargarArrayLluvia()
On Error GoTo ErrorHandler:
    Dim N As Integer
    Dim i As Long
    Dim Nu As Integer
    N = FreeFile()
    Open Game.path(INIT) & "fk.ind" For Binary Access Read As #N
    Get #N, , MiCabecera
    Get #N, , Nu
    ReDim bLluvia(1 To Nu) As Byte
    For i = 1 To Nu
        Get #N, , bLluvia(i)
    Next i
    Close #N
ErrorHandler:
    If Err.number <> 0 Then
        If Err.number = 53 Then
            Call MsgBox("El archivo fk.ind no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
        End If
    End If
End Sub

Sub CargarAnimArmas()
On Error GoTo ErrorHandler:
    Dim LoopC As Long
    Set FileManager = New clsIniManager
    Call FileManager.Initialize(Game.path(INIT) & "armas.dat")
    NumWeaponAnims = Val(FileManager.GetValue("INIT", "NumArmas"))
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    For LoopC = 1 To NumWeaponAnims
        Call InitGrh(WeaponAnimData(LoopC).WeaponWalk(1), Val(FileManager.GetValue("ARMA" & LoopC, "Dir1")), 0)
        Call InitGrh(WeaponAnimData(LoopC).WeaponWalk(2), Val(FileManager.GetValue("ARMA" & LoopC, "Dir2")), 0)
        Call InitGrh(WeaponAnimData(LoopC).WeaponWalk(3), Val(FileManager.GetValue("ARMA" & LoopC, "Dir3")), 0)
        Call InitGrh(WeaponAnimData(LoopC).WeaponWalk(4), Val(FileManager.GetValue("ARMA" & LoopC, "Dir4")), 0)
    Next LoopC
    Set FileManager = Nothing
ErrorHandler:
    If Err.number <> 0 Then
        If Err.number = 53 Then
            Call MsgBox("El archivo armas.dat no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
        End If
    End If
End Sub

Public Sub CargarColores()
On Error GoTo ErrorHandler:
    Set FileManager = New clsIniManager
    Call FileManager.Initialize(Game.path(INIT) & "colores.dat")
    Dim i As Long
    For i = 0 To 47
        ColoresPJ(i) = D3DColorXRGB(FileManager.GetValue(CStr(i), "R"), FileManager.GetValue(CStr(i), "G"), FileManager.GetValue(CStr(i), "B"))
    Next i
    ColoresPJ(50) = D3DColorXRGB(FileManager.GetValue("CR", "R"), FileManager.GetValue("CR", "G"), FileManager.GetValue("CR", "B"))
    ColoresPJ(49) = D3DColorXRGB(FileManager.GetValue("CI", "R"), FileManager.GetValue("CI", "G"), FileManager.GetValue("CI", "B"))
    For i = 51 To 56
        ColoresDano(i) = D3DColorXRGB(FileManager.GetValue(CStr(i), "R"), FileManager.GetValue(CStr(i), "G"), FileManager.GetValue(CStr(i), "B"))
    Next i
    Set FileManager = Nothing
ErrorHandler:
    If Err.number <> 0 Then
        If Err.number = 53 Then
            Call MsgBox("El archivo colores.dat no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
        End If
    End If
End Sub

Sub CargarAnimEscudos()
On Error GoTo ErrorHandler:
    Dim LoopC As Long
    Dim NumEscudosAnims As Integer
    Set FileManager = New clsIniManager
    Call FileManager.Initialize(Game.path(INIT) & "escudos.dat")
    NumEscudosAnims = Val(FileManager.GetValue("INIT", "NumEscudos"))
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    For LoopC = 1 To NumEscudosAnims
        Call InitGrh(ShieldAnimData(LoopC).ShieldWalk(1), Val(FileManager.GetValue("ESC" & LoopC, "Dir1")), 0)
        Call InitGrh(ShieldAnimData(LoopC).ShieldWalk(2), Val(FileManager.GetValue("ESC" & LoopC, "Dir2")), 0)
        Call InitGrh(ShieldAnimData(LoopC).ShieldWalk(3), Val(FileManager.GetValue("ESC" & LoopC, "Dir3")), 0)
        Call InitGrh(ShieldAnimData(LoopC).ShieldWalk(4), Val(FileManager.GetValue("ESC" & LoopC, "Dir4")), 0)
    Next LoopC
    Set FileManager = Nothing
ErrorHandler:
    If Err.number <> 0 Then
        If Err.number = 53 Then
            Call MsgBox("El archivo escudos.dat no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
            Call CloseClient
        End If
    End If
End Sub

Public Sub CargarHechizos()
On Error GoTo ErrorHandler
    Dim J As Long
    Set FileManager = New clsIniManager
    Call FileManager.Initialize(Game.path(INIT) & "Hechizos.dat")
    NumHechizos = Val(FileManager.GetValue("INIT", "NumHechizos"))
    ReDim Hechizos(1 To NumHechizos) As tHechizos
    For J = 1 To NumHechizos
        With Hechizos(J)
            .Desc = FileManager.GetValue("HECHIZO" & J, "Desc")
            .PalabrasMagicas = FileManager.GetValue("HECHIZO" & J, "PalabrasMagicas")
            .Nombre = FileManager.GetValue("HECHIZO" & J, "Nombre")
            .SkillRequerido = Val(FileManager.GetValue("HECHIZO" & J, "MinSkill"))
            If J <> 38 And J <> 39 Then
                .EnergiaRequerida = Val(FileManager.GetValue("HECHIZO" & J, "StaRequerido"))
                .HechiceroMsg = FileManager.GetValue("HECHIZO" & J, "HechizeroMsg")
                .ManaRequerida = Val(FileManager.GetValue("HECHIZO" & J, "ManaRequerido"))
                .PropioMsg = FileManager.GetValue("HECHIZO" & J, "PropioMsg")
                .TargetMsg = FileManager.GetValue("HECHIZO" & J, "TargetMsg")
            End If
        End With
    Next J
    Set FileManager = Nothing
Exit Sub
ErrorHandler:
    If Err.number <> 0 Then
        Select Case Err.number
            Case 9
                Call MsgBox("Error cargando el archivo Hechizos.dat (Hechizo " & J & "). Por favor, avise a los administradores enviandoles el archivo Errores.log que se encuentra en la carpeta del cliente.", , "Argentum Online Libre")
                Call LogError(Err.number, Err.Description, "CargarHechizos")
            
            Case 53
                Call MsgBox("El archivo Hechizos.dat no existe. Por favor, reinstale el juego.", , "Argentum Online Libre")
        End Select
        Call CloseClient
    End If
End Sub
