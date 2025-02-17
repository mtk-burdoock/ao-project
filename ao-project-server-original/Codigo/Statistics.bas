Attribute VB_Name = "Statistics"
Option Explicit

Private Type trainingData
    startTick As Long
    trainingTime As Long
End Type

Private Type fragLvlRace
    matrix(1 To 50, 1 To 5) As Long
End Type

Private Type fragLvlLvl
    matrix(1 To 50, 1 To 50) As Long
End Type

Private trainingInfo()                        As trainingData
Private fragLvlRaceData(1 To 7)               As fragLvlRace
Private fragLvlLvlData(1 To 7)                As fragLvlLvl
Private fragAlignmentLvlData(1 To 50, 1 To 4) As Long
Private keyOcurrencies(255)                   As Currency

Public Sub Initialize()
    ReDim trainingInfo(1 To MaxUsers) As trainingData
End Sub

Public Sub UserConnected(ByVal Userindex As Integer)
    trainingInfo(Userindex).trainingTime = GetUserTrainingTime(UserList(Userindex).Name)
    trainingInfo(Userindex).startTick = (GetTickCount())
End Sub

Public Sub UserDisconnected(ByVal Userindex As Integer)
    With trainingInfo(Userindex)
        .trainingTime = .trainingTime + ((GetTickCount()) - .startTick) / 1000
        .startTick = (GetTickCount())
        Call SaveUserTrainingTime(UserList(Userindex).Name, .trainingTime)
    End With
End Sub

Public Sub UserLevelUp(ByVal Userindex As Integer)
    Dim handle As Integer
    handle = FreeFile()
    With trainingInfo(Userindex)
        Open App.Path & "\logs\statistics.log" For Append Shared As handle
        Print #handle, UCase$(UserList(Userindex).Name) & " completo el nivel " & CStr(UserList(Userindex).Stats.ELV) & " en " & CStr(.trainingTime + ((GetTickCount()) - .startTick) / 1000) & " segundos."
        Close handle
        .trainingTime = 0
        .startTick = (GetTickCount())
    End With
End Sub

Public Sub StoreFrag(ByVal killer As Integer, ByVal victim As Integer)
    Dim Clase     As Integer
    Dim raza      As Integer
    Dim alignment As Integer
    If UserList(victim).Stats.ELV > 50 Or UserList(killer).Stats.ELV > 50 Then Exit Sub
    Select Case UserList(killer).Clase
        Case eClass.Assasin
            Clase = 1
        
        Case eClass.Bard
            Clase = 2
        
        Case eClass.Mage
            Clase = 3
        
        Case eClass.Paladin
            Clase = 4
        
        Case eClass.Warrior
            Clase = 5
        
        Case eClass.Cleric
            Clase = 6
        
        Case eClass.Hunter
            Clase = 7
        
        Case Else
            Exit Sub
    End Select
    Select Case UserList(killer).raza
        Case eRaza.Elfo
            raza = 1
        
        Case eRaza.Drow
            raza = 2
        
        Case eRaza.Enano
            raza = 3
        
        Case eRaza.Gnomo
            raza = 4
        
        Case eRaza.Humano
            raza = 5
        
        Case Else
            Exit Sub
    End Select
    If criminal(killer) Then
        If esCaos(killer) Then
            alignment = 2
        Else
            alignment = 3
        End If
    Else
        If esArmada(killer) Then
            alignment = 1
        Else
            alignment = 4
        End If
    End If
    fragLvlRaceData(Clase).matrix(UserList(killer).Stats.ELV, raza) = fragLvlRaceData(Clase).matrix(UserList(killer).Stats.ELV, raza) + 1
    fragLvlLvlData(Clase).matrix(UserList(killer).Stats.ELV, UserList(victim).Stats.ELV) = fragLvlLvlData(Clase).matrix(UserList(killer).Stats.ELV, UserList(victim).Stats.ELV) + 1
    fragAlignmentLvlData(UserList(killer).Stats.ELV, alignment) = fragAlignmentLvlData(UserList(killer).Stats.ELV, alignment) + 1
End Sub

Public Sub DumpStatistics()
    Dim handle As Integer
    handle = FreeFile()
    Dim line As String
    Dim i    As Long
    Dim j    As Long
    Open App.Path & "\logs\frags.txt" For Output As handle
    Print #handle, "# name: fragLvlLvl_Ase"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    For j = 1 To 50
        For i = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(1).matrix(i, j))
        Next i
        Print #handle, line
        line = vbNullString
    Next j
    Print #handle, "# name: fragLvlLvl_Bar"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    For j = 1 To 50
        For i = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(2).matrix(i, j))
        Next i
        Print #handle, line
        line = vbNullString
    Next j
    Print #handle, "# name: fragLvlLvl_Mag"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    For j = 1 To 50
        For i = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(3).matrix(i, j))
        Next i
        Print #handle, line
        line = vbNullString
    Next j
    Print #handle, "# name: fragLvlLvl_Pal"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    For j = 1 To 50
        For i = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(4).matrix(i, j))
        Next i
        Print #handle, line
        line = vbNullString
    Next j
    Print #handle, "# name: fragLvlLvl_Gue"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    For j = 1 To 50
        For i = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(5).matrix(i, j))
        Next i
        Print #handle, line
        line = vbNullString
    Next j
    Print #handle, "# name: fragLvlLvl_Cle"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    For j = 1 To 50
        For i = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(6).matrix(i, j))
        Next i
        Print #handle, line
        line = vbNullString
    Next j
    Print #handle, "# name: fragLvlLvl_Caz"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 50"
    Print #handle, "# columns: 50"
    For j = 1 To 50
        For i = 1 To 50
            line = line & " " & CStr(fragLvlLvlData(7).matrix(i, j))
        Next i
        Print #handle, line
        line = vbNullString
    Next j
    Print #handle, "# name: fragLvlRace_Ase"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    For j = 1 To 5
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(1).matrix(i, j))
        Next i
        Print #handle, line
        line = vbNullString
    Next j
    Print #handle, "# name: fragLvlRace_Bar"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    For j = 1 To 5
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(2).matrix(i, j))
        Next i
        Print #handle, line
        line = vbNullString
    Next j
    Print #handle, "# name: fragLvlRace_Mag"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    For j = 1 To 5
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(3).matrix(i, j))
        Next i
        Print #handle, line
        line = vbNullString
    Next j
    Print #handle, "# name: fragLvlRace_Pal"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    For j = 1 To 5
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(4).matrix(i, j))
        Next i
        Print #handle, line
        line = vbNullString
    Next j
    Print #handle, "# name: fragLvlRace_Gue"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    For j = 1 To 5
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(5).matrix(i, j))
        Next i
        Print #handle, line
        line = vbNullString
    Next j
    Print #handle, "# name: fragLvlRace_Cle"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    For j = 1 To 5
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(6).matrix(i, j))
        Next i
        Print #handle, line
        line = vbNullString
    Next j
    Print #handle, "# name: fragLvlRace_Caz"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 5"
    Print #handle, "# columns: 50"
    For j = 1 To 5
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(7).matrix(i, j))
        Next i
        Print #handle, line
        line = vbNullString
    Next j
    Print #handle, "# name: fragLvlClass_Elf"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 7"
    Print #handle, "# columns: 50"
    For j = 1 To 7
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(j).matrix(i, 1))
        Next i
        Print #handle, line
        line = vbNullString
    Next j
    Print #handle, "# name: fragLvlClass_Dar"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 7"
    Print #handle, "# columns: 50"
    For j = 1 To 7
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(j).matrix(i, 2))
        Next i
        Print #handle, line
        line = vbNullString
    Next j
    Print #handle, "# name: fragLvlClass_Dwa"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 7"
    Print #handle, "# columns: 50"
    For j = 1 To 7
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(j).matrix(i, 3))
        Next i
        Print #handle, line
        line = vbNullString
    Next j
    Print #handle, "# name: fragLvlClass_Gno"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 7"
    Print #handle, "# columns: 50"
    For j = 1 To 7
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(j).matrix(i, 4))
        Next i
        Print #handle, line
        line = vbNullString
    Next j
    Print #handle, "# name: fragLvlClass_Hum"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 7"
    Print #handle, "# columns: 50"
    For j = 1 To 7
        For i = 1 To 50
            line = line & " " & CStr(fragLvlRaceData(j).matrix(i, 5))
        Next i
        Print #handle, line
        line = vbNullString
    Next j
    Print #handle, "# name: fragAlignmentLvl"
    Print #handle, "# type: matrix"
    Print #handle, "# rows: 4"
    Print #handle, "# columns: 50"
    For j = 1 To 4
        For i = 1 To 50
            line = line & " " & CStr(fragAlignmentLvlData(i, j))
        Next i
        Print #handle, line
        line = vbNullString
    Next j
    Close handle
    handle = FreeFile()
    Open App.Path & "\logs\huffman.log" For Output As handle
    Dim Total As Currency
    For i = 0 To 255
        Total = Total + keyOcurrencies(i)
    Next i
    If Total <> 0 Then
        For i = 0 To 255
            Print #handle, CStr(i) & "    " & CStr(Round(keyOcurrencies(i) / Total, 8))
        Next i
    End If
    Print #handle, "TOTAL =    " & CStr(Total)
    Close handle
End Sub

Public Sub ParseChat(ByRef S As String)
    Dim i   As Long
    Dim key As Integer
    For i = 1 To Len(S)
        key = Asc(mid$(S, i, 1))
        keyOcurrencies(key) = keyOcurrencies(key) + 1
    Next i
    keyOcurrencies(0) = keyOcurrencies(0) + 1
End Sub
