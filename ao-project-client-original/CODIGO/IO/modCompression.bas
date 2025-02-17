Attribute VB_Name = "modCompression"
Option Explicit

Public Const PNG_SOURCE_FILE_EXT As String = ".png"
Public Const BMP_SOURCE_FILE_EXT As String = ".bmp"
Public Const GRH_RESOURCE_FILE As String = "Graphics.AO"
Public Const GRH_PATCH_FILE As String = "Graficos.PATCH"
Public Const MAPS_SOURCE_FILE_EXT As String = ".map"
Public Const MAPS_RESOURCE_FILE As String = "Mapas.AO"
Public Const MAPS_PATCH_FILE As String = "Mapas.PATCH"
Public GrhDatContra() As Byte
Public GrhUsaContra As Boolean
Public MapsDatContra() As Byte
Public MapsUsaContra As Boolean

Public Type FILEHEADER
    lngNumFiles As Long
    lngFileSize As Long
    lngFileVersion As Long
End Type

Public Type INFOHEADER
    lngFileSize As Long
    lngFileStart As Long
    strFileName As String * 16
    lngFileSizeUncompressed As Long
End Type

Private Enum PatchInstruction
    Delete_File
    Create_File
    Modify_File
End Enum

Private Declare Function compress Lib "zlib.dll" (dest As Any, destlen As Any, src As Any, ByVal srclen As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destlen As Any, src As Any, ByVal srclen As Long) As Long

Public Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD
End Type

Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, FreeBytesToCaller As Currency, bytesTotal As Currency, FreeBytesTotal As Currency) As Long

Public Sub GenerateContra(ByVal Contra As String, Optional Modo As Byte = 0)
On Error Resume Next
    Dim LoopC As Byte
    Dim Upper_grhDatContra As Long, Upper_mapsDatContra As Long
    If Modo = 0 Then
        Erase GrhDatContra
    ElseIf Modo = 1 Then
        Erase MapsDatContra
    End If
    If LenB(Contra) <> 0 Then
        If Modo = 0 Then
            ReDim GrhDatContra(Len(Contra) - 1)
            Upper_grhDatContra = UBound(GrhDatContra)
            For LoopC = 0 To Upper_grhDatContra
                GrhDatContra(LoopC) = Asc(mid$(Contra, LoopC + 1, 1))
            Next LoopC
            GrhUsaContra = True
        ElseIf Modo = 1 Then
            ReDim MapsDatContra(Len(Contra) - 1)
            Upper_mapsDatContra = UBound(MapsDatContra)
            For LoopC = 0 To Upper_mapsDatContra
                MapsDatContra(LoopC) = Asc(mid$(Contra, LoopC + 1, 1))
            Next LoopC
            MapsUsaContra = True
        End If
    Else
        If Modo = 0 Then
            GrhUsaContra = False
        ElseIf Modo = 1 Then
            MapsUsaContra = False
        End If
    End If
End Sub

Private Function General_Drive_Get_Free_Bytes(ByVal DriveName As String) As Currency
    Dim retval As Long
    Dim FB As Currency
    Dim BT As Currency
    Dim FBT As Currency
    retval = GetDiskFreeSpace(Left$(DriveName, 2), FB, BT, FBT)
    General_Drive_Get_Free_Bytes = FB * 10000
End Function

Private Sub Sort_Info_Headers(ByRef InfoHead() As INFOHEADER, ByVal First As Long, ByVal Last As Long)
    Dim aux As INFOHEADER
    Dim min As Long
    Dim max As Long
    Dim comp As String
    min = First
    max = Last
    comp = InfoHead((min + max) \ 2).strFileName
    Do While min <= max
        Do While InfoHead(min).strFileName < comp And min < Last
            min = min + 1
        Loop
        Do While InfoHead(max).strFileName > comp And max > First
            max = max - 1
        Loop
        If min <= max Then
            aux = InfoHead(min)
            InfoHead(min) = InfoHead(max)
            InfoHead(max) = aux
            min = min + 1
            max = max - 1
        End If
    Loop
    If First < max Then Call Sort_Info_Headers(InfoHead, First, max)
    If min < Last Then Call Sort_Info_Headers(InfoHead, min, Last)
End Sub

Private Function BinarySearch(ByRef ResourceFile As Integer, ByRef InfoHead As INFOHEADER, ByVal FirstHead As Long, ByVal LastHead As Long, ByVal FileHeaderSize As Long, ByVal InfoHeaderSize As Long) As Boolean
    Dim ReadingHead As Long
    Dim ReadInfoHead As INFOHEADER
    Do Until FirstHead > LastHead
        ReadingHead = (FirstHead + LastHead) \ 2
        Get ResourceFile, FileHeaderSize + InfoHeaderSize * (ReadingHead - 1) + 1, ReadInfoHead
        If InfoHead.strFileName = ReadInfoHead.strFileName Then
            InfoHead = ReadInfoHead
            BinarySearch = True
            Exit Function
        Else
            If InfoHead.strFileName < ReadInfoHead.strFileName Then
                LastHead = ReadingHead - 1
            Else
                FirstHead = ReadingHead + 1
            End If
        End If
    Loop
End Function

Private Function Get_InfoHeader(ByRef ResourcePath As String, ByRef FileName As String, ByRef InfoHead As INFOHEADER, Optional Modo As Byte = 0) As Boolean
    Dim ResourceFile As Integer
    Dim ResourceFilePath As String
    Dim FileHead As FILEHEADER
    Dim ERROR_LEER_ARCHIVO As String
On Local Error GoTo ErrorHandler
    If Modo = 0 Then
        ResourceFilePath = ResourcePath & GRH_RESOURCE_FILE
    ElseIf Modo = 1 Then
        ResourceFilePath = ResourcePath & MAPS_RESOURCE_FILE
    End If
    InfoHead.strFileName = UCase$(FileName)
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        Get ResourceFile, 1, FileHead
        If LOF(ResourceFile) <> FileHead.lngFileSize Then
            MsgBox JsonLanguage.item("ERROR_ARCHIVO_CORRUPTO").item("TEXTO") & ": " & ResourceFilePath, , JsonLanguage.item("Error").item("TEXTO")
            Close ResourceFile
            Exit Function
        End If
        If BinarySearch(ResourceFile, InfoHead, 1, FileHead.lngNumFiles, Len(FileHead), Len(InfoHead)) Then
            Get_InfoHeader = True
        End If
    Close ResourceFile
Exit Function
ErrorHandler:
    Close ResourceFile
    ERROR_LEER_ARCHIVO = JsonLanguage.item("ERROR_LEER_ARCHIVO").item("TEXTO")
    ERROR_LEER_ARCHIVO = Replace$(ERROR_LEER_ARCHIVO, "VAR_ARCHIVO", ResourceFilePath)
    ERROR_LEER_ARCHIVO = Replace$(ERROR_LEER_ARCHIVO, "VAR_ERROR", Err.number & " : " & Err.Description)
    Call MsgBox(ERROR_LEER_ARCHIVO)
End Function

Private Sub Compress_Data(ByRef data() As Byte, Optional Modo As Byte = 0)
    Dim Dimensions As Long
    Dim DimBuffer As Long
    Dim BufTemp() As Byte
    Dim LoopC As Long
    Dim Upper_grhDatContra As Long, Upper_mapsDatContra As Long
    Dimensions = UBound(data) + 1
    DimBuffer = Dimensions * 1.06
    ReDim BufTemp(DimBuffer)
    Call compress(BufTemp(0), DimBuffer, data(0), Dimensions)
    Erase data
    ReDim data(DimBuffer - 1)
    ReDim Preserve BufTemp(DimBuffer - 1)
    data = BufTemp
    Erase BufTemp
    If Modo = 0 And GrhUsaContra = True Then
        If UBound(GrhDatContra) <= UBound(data) And UBound(GrhDatContra) <> 0 Then
            Upper_grhDatContra = UBound(GrhDatContra)
            For LoopC = 0 To Upper_grhDatContra
                data(LoopC) = data(LoopC) Xor GrhDatContra(LoopC)
            Next LoopC
        End If
    ElseIf Modo = 1 And MapsUsaContra = True Then
        If UBound(MapsDatContra) <= UBound(data) And UBound(MapsDatContra) <> 0 Then
            Upper_mapsDatContra = UBound(MapsDatContra)
            
            For LoopC = 0 To Upper_mapsDatContra
                data(LoopC) = data(LoopC) Xor MapsDatContra(LoopC)
            Next LoopC
        End If
    End If
End Sub

Private Sub Decompress_Data(ByRef data() As Byte, ByVal OrigSize As Long, Optional Modo As Byte = 0)
    Dim BufTemp() As Byte
    Dim LoopC As Integer
    Dim Upper_grhDatContra As Long, Upper_mapsDatContra As Long
    ReDim BufTemp(OrigSize - 1)
    If Modo = 0 And GrhUsaContra = True Then
        If UBound(GrhDatContra) <= UBound(data) And UBound(GrhDatContra) <> 0 Then
            Upper_grhDatContra = UBound(GrhDatContra)
            For LoopC = 0 To Upper_grhDatContra
                data(LoopC) = data(LoopC) Xor GrhDatContra(LoopC)
            Next LoopC
        End If
    ElseIf Modo = 1 And MapsUsaContra = True Then
        If UBound(MapsDatContra) <= UBound(data) And UBound(MapsDatContra) <> 0 Then
            Upper_mapsDatContra = UBound(MapsDatContra)
            For LoopC = 0 To Upper_mapsDatContra
                data(LoopC) = data(LoopC) Xor MapsDatContra(LoopC)
            Next LoopC
        End If
    End If
    Call uncompress(BufTemp(0), OrigSize, data(0), UBound(data) + 1)
    ReDim data(OrigSize - 1)
    data = BufTemp
    Erase BufTemp
End Sub

Public Function Compress_Files(ByRef SourcePath As String, ByRef OutputPath As String, ByVal version As Long, ByRef prgBar As ProgressBar, Optional Modo As Byte = 0) As Boolean
    Dim SourceFileName As String
    Dim OutputFilePath As String
    Dim SourceFile As Long
    Dim OutputFile As Long
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim LoopC As Long
    Dim ERROR_EXT_NO_ENCONTRADA As String
On Local Error GoTo ErrorHandler
    If Modo = 0 Then
        OutputFilePath = OutputPath & GRH_RESOURCE_FILE
            SourceFileName = Dir$(SourcePath & "*" & BMP_SOURCE_FILE_EXT, vbNormal)
    ElseIf Modo = 1 Then
        OutputFilePath = OutputPath & MAPS_RESOURCE_FILE
        SourceFileName = Dir$(SourcePath & "*" & MAPS_SOURCE_FILE_EXT, vbNormal)
    End If
    While LenB(SourceFileName) <> 0
        FileHead.lngNumFiles = FileHead.lngNumFiles + 1
        ReDim Preserve InfoHead(FileHead.lngNumFiles - 1)
        InfoHead(FileHead.lngNumFiles - 1).strFileName = UCase$(SourceFileName)
        SourceFileName = Dir$()
    Wend
        SourceFileName = Dir$(SourcePath & "*" & PNG_SOURCE_FILE_EXT, vbNormal)
        While LenB(SourceFileName) <> 0
            FileHead.lngNumFiles = FileHead.lngNumFiles + 1
            ReDim Preserve InfoHead(FileHead.lngNumFiles - 1)
            InfoHead(FileHead.lngNumFiles - 1).strFileName = UCase$(SourceFileName)
            SourceFileName = Dir$()
        Wend
        SourceFileName = Dir$(SourcePath & "*" & BMP_SOURCE_FILE_EXT, vbNormal)
        While LenB(SourceFileName) <> 0
            FileHead.lngNumFiles = FileHead.lngNumFiles + 1
            ReDim Preserve InfoHead(FileHead.lngNumFiles - 1)
            InfoHead(FileHead.lngNumFiles - 1).strFileName = UCase$(SourceFileName)
            SourceFileName = Dir$()
        Wend
    If FileHead.lngNumFiles = 0 Then
            ERROR_EXT_NO_ENCONTRADA = JsonLanguage.item("ERROR_EXT_NO_ENCONTRADA").item("TEXTO")
            ERROR_EXT_NO_ENCONTRADA = Replace$(ERROR_EXT_NO_ENCONTRADA, "VAR_EXT", BMP_SOURCE_FILE_EXT)
            ERROR_EXT_NO_ENCONTRADA = Replace$(ERROR_EXT_NO_ENCONTRADA, "VAR_PATH", SourcePath)
            MsgBox ERROR_EXT_NO_ENCONTRADA, , JsonLanguage.item("Error").item("TEXTO")
        Exit Function
    End If
    If Not prgBar Is Nothing Then
        prgBar.Value = 0
        prgBar.max = FileHead.lngNumFiles + 1
    End If
    If LenB(Dir$(OutputFilePath, vbNormal)) <> 0 Then
        Kill OutputFilePath
    End If
    FileHead.lngFileVersion = version
    FileHead.lngFileSize = Len(FileHead) + FileHead.lngNumFiles * Len(InfoHead(0))
    Call Sort_Info_Headers(InfoHead(), 0, FileHead.lngNumFiles - 1)
    OutputFile = FreeFile()
    Open OutputFilePath For Binary Access Read Write As OutputFile
        Seek OutputFile, FileHead.lngFileSize + 1
        For LoopC = 0 To FileHead.lngNumFiles - 1
            SourceFile = FreeFile()
            Open SourcePath & InfoHead(LoopC).strFileName For Binary Access Read Lock Write As SourceFile
                InfoHead(LoopC).lngFileSizeUncompressed = LOF(SourceFile)
                ReDim SourceData(LOF(SourceFile) - 1)
                Get SourceFile, , SourceData
                Call Compress_Data(SourceData, Modo)
                Put OutputFile, , SourceData
                With InfoHead(LoopC)
                    .lngFileSize = UBound(SourceData) + 1
                    .lngFileStart = FileHead.lngFileSize + 1
                    FileHead.lngFileSize = FileHead.lngFileSize + .lngFileSize
                End With
                Erase SourceData
            Close SourceFile
            If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1
            DoEvents
        Next LoopC
        Seek OutputFile, 1
        Put OutputFile, , FileHead
        Put OutputFile, , InfoHead
    Close OutputFile
    Erase InfoHead
    Erase SourceData
    Compress_Files = True
Exit Function
ErrorHandler:
    Erase SourceData
    Erase InfoHead
    Close OutputFile
    Call MsgBox(Replace$(JsonLanguage.item("ERROR_CREAR_BINARIO").item("TEXTO"), "VAR_ERROR", Err.number & " : " & Err.Description), vbOKOnly, JsonLanguage.item("Error").item("TEXTO"))
End Function

Public Function Get_File_RawData(ByRef ResourcePath As String, ByRef InfoHead As INFOHEADER, ByRef data() As Byte, Optional Modo As Byte = 0) As Boolean
    Dim ResourceFilePath As String
    Dim ResourceFile As Integer
On Local Error GoTo ErrorHandler
    If Modo = 0 Then
        ResourceFilePath = ResourcePath & GRH_RESOURCE_FILE
    ElseIf Modo = 1 Then
        ResourceFilePath = ResourcePath & MAPS_RESOURCE_FILE
    End If
    ReDim data(InfoHead.lngFileSize - 1)
    ResourceFile = FreeFile
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        Get ResourceFile, InfoHead.lngFileStart, data
    Close ResourceFile
    Get_File_RawData = True
Exit Function
ErrorHandler:
    Close ResourceFile
End Function

Public Function Extract_File(ByRef ResourcePath As String, ByRef InfoHead As INFOHEADER, ByRef data() As Byte, Optional Modo As Byte = 0) As Boolean
On Local Error GoTo ErrorHandler
    If Get_File_RawData(ResourcePath, InfoHead, data, Modo) Then
        Call Decompress_Data(data, InfoHead.lngFileSizeUncompressed, Modo)
        Extract_File = True
    End If
Exit Function
ErrorHandler:
    Call MsgBox(Replace$(JsonLanguage.item("ERROR_DECODE_RECURSOS").item("TEXTO"), "VAR_ERROR", Err.number & " : " & Err.Description), vbOKOnly, JsonLanguage.item("Error").item("TEXTO"))
End Function

Public Function Extract_Files(ByRef ResourcePath As String, ByRef OutputPath As String, ByRef prgBar As ProgressBar, Optional Modo As Byte = 0) As Boolean
    Dim LoopC As Long
    Dim ResourceFile As Integer
    Dim ResourceFilePath As String
    Dim OutputFile As Integer
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim Upper_infoHead As Long
    Dim RequiredSpace As Currency
On Local Error GoTo ErrorHandler
    If Modo = 0 Then
        ResourceFilePath = ResourcePath & GRH_RESOURCE_FILE
    ElseIf Modo = 1 Then
        ResourceFilePath = ResourcePath & MAPS_RESOURCE_FILE
    End If
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        Get ResourceFile, 1, FileHead
        If LOF(ResourceFile) <> FileHead.lngFileSize Then
            Call MsgBox(JsonLanguage.item("ERROR_ARCHIVO_CORRUPTO").item("TEXTO") & ": " & ResourceFilePath, , JsonLanguage.item("Error").item("TEXTO"))
            Close ResourceFile
            Exit Function
        End If
        ReDim InfoHead(FileHead.lngNumFiles - 1)
        Get ResourceFile, , InfoHead
        Upper_infoHead = UBound(InfoHead)
        For LoopC = 0 To Upper_infoHead
            RequiredSpace = RequiredSpace + InfoHead(LoopC).lngFileSizeUncompressed
        Next LoopC
        If RequiredSpace >= General_Drive_Get_Free_Bytes(Left$(App.path, 3)) Then
            Erase InfoHead
            Close ResourceFile
            Call MsgBox(JsonLanguage.item("ERROR_SIN_ESPACIO").item("TEXTO"), , JsonLanguage.item("Error").item("TEXTO"))
            Exit Function
        End If
    Close ResourceFile
    If Not prgBar Is Nothing Then
        prgBar.Value = 0
        prgBar.max = FileHead.lngNumFiles + 1
    End If
    Upper_infoHead = UBound(InfoHead)
    For LoopC = 0 To Upper_infoHead
        If Extract_File(ResourcePath, InfoHead(LoopC), SourceData) Then
            If FileExist(OutputPath & InfoHead(LoopC).strFileName, vbNormal) Then
                Call Kill(OutputPath & InfoHead(LoopC).strFileName)
            End If
            OutputFile = FreeFile()
            Open OutputPath & InfoHead(LoopC).strFileName For Binary As OutputFile
                Put OutputFile, , SourceData
            Close OutputFile
            Erase SourceData
        Else
            Erase SourceData
            Erase InfoHead
            Call MsgBox(JsonLanguage.item("ERROR_EXTRAER_ARCHIVO").item("TEXTO") & ": " & InfoHead(LoopC).strFileName, vbOKOnly, JsonLanguage.item("Error").item("TEXTO"))
            Exit Function
        End If
        If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1
        DoEvents
    Next LoopC
    Erase InfoHead
    Extract_Files = True
Exit Function
ErrorHandler:
    Close ResourceFile
    Erase SourceData
    Erase InfoHead
    Call MsgBox(Replace$(JsonLanguage.item("ERROR_EXTRAER_BINARIO").item("TEXTO"), "VAR_ERROR", Err.number & " : " & Err.Description), vbOKOnly, JsonLanguage.item("Error").item("TEXTO"))
End Function

Public Function Get_File_Data(ByRef ResourcePath As String, ByRef FileName As String, ByRef data() As Byte, Optional Modo As Byte = 0) As Boolean
    Dim InfoHead As INFOHEADER
    If Get_InfoHeader(ResourcePath, FileName, InfoHead, Modo) Then
        Get_File_Data = Extract_File(ResourcePath, InfoHead, data, Modo)
    Else
        Get_File_Data = False
    End If
End Function

Public Function Get_Image(ByRef ResourcePath As String, ByRef FileName As String, ByRef data() As Byte, Optional SoloBMP As Boolean = False) As Boolean
    Dim InfoHead As INFOHEADER
    Dim ExistFile As Boolean
    ExistFile = False
    If SoloBMP = True Then
        If Get_InfoHeader(ResourcePath, FileName & ".BMP", InfoHead, 0) Then
            FileName = FileName & ".BMP"
            ExistFile = True
        End If
    Else
        If Get_InfoHeader(ResourcePath, FileName & ".BMP", InfoHead, 0) Then
            FileName = FileName & ".BMP"
            ExistFile = True
        ElseIf Get_InfoHeader(ResourcePath, FileName & ".PNG", InfoHead, 0) Then
            FileName = FileName & ".PNG"
            ExistFile = True
        End If
    End If
    If ExistFile = True Then
        If Extract_File(ResourcePath, InfoHead, data, 0) Then Get_Image = True
    Else
        Call MsgBox(JsonLanguage("ERROR_404").item("TEXTO") & ": " & FileName)
    End If
End Function

Private Function Compare_Datas(ByRef data1() As Byte, ByRef data2() As Byte) As Boolean
    Dim Length As Long
    Dim act As Long
    Length = UBound(data1) + 1
    If (UBound(data2) + 1) = Length Then
        While act < Length
            If data1(act) Xor data2(act) Then Exit Function
            act = act + 1
        Wend
        Compare_Datas = True
    End If
End Function

Private Function ReadNext_InfoHead(ByRef ResourceFile As Integer, ByRef FileHead As FILEHEADER, ByRef InfoHead As INFOHEADER, ByRef ReadFiles As Long) As Boolean
    If ReadFiles < FileHead.lngNumFiles Then
        Get ResourceFile, Len(FileHead) + Len(InfoHead) * ReadFiles + 1, InfoHead
        ReadNext_InfoHead = True
    End If
    ReadFiles = ReadFiles + 1
End Function

Public Function GetNext_Bitmap(ByRef ResourcePath As String, ByRef ReadFiles As Long, ByRef bmpInfo As BITMAPINFO, ByRef data() As Byte, ByRef fileIndex As Long) As Boolean
On Error Resume Next
    Dim ResourceFile As Integer
    Dim FileHead As FILEHEADER
    Dim InfoHead As INFOHEADER
    Dim FileName As String
    ResourceFile = FreeFile
    Open ResourcePath & GRH_RESOURCE_FILE For Binary Access Read Lock Write As ResourceFile
    Get ResourceFile, 1, FileHead
    If ReadNext_InfoHead(ResourceFile, FileHead, InfoHead, ReadFiles) Then
        Call Get_Image(ResourcePath, InfoHead.strFileName, data())
        FileName = Trim$(InfoHead.strFileName)
        fileIndex = CLng(Left$(FileName, Len(FileName) - 4))
        GetNext_Bitmap = True
    End If
    Close ResourceFile
End Function

Public Function Make_Patch(ByRef NewResourcePath As String, ByRef OldResourcePath As String, ByRef OutputPath As String, ByRef prgBar As ProgressBar, Optional Modo As Byte = 0) As Boolean
    Dim NewResourceFile As Integer
    Dim NewResourceFilePath As String
    Dim NewFileHead As FILEHEADER
    Dim NewInfoHead As INFOHEADER
    Dim NewReadFiles As Long
    Dim NewReadNext As Boolean
    Dim OldResourceFile As Integer
    Dim OldResourceFilePath As String
    Dim OldFileHead As FILEHEADER
    Dim OldInfoHead As INFOHEADER
    Dim OldReadFiles As Long
    Dim OldReadNext As Boolean
    Dim OutputFile As Integer
    Dim OutputFilePath As String
    Dim data() As Byte
    Dim auxData() As Byte
    Dim Instruction As Byte
On Local Error GoTo ErrorHandler
    If Modo = 0 Then
        NewResourceFilePath = NewResourcePath & GRH_RESOURCE_FILE
        OldResourceFilePath = OldResourcePath & GRH_RESOURCE_FILE
        OutputFilePath = OutputPath & GRH_PATCH_FILE
    ElseIf Modo = 1 Then
        NewResourceFilePath = NewResourcePath & MAPS_RESOURCE_FILE
        OldResourceFilePath = OldResourcePath & MAPS_RESOURCE_FILE
        OutputFilePath = OutputPath & MAPS_PATCH_FILE
    End If
    OldResourceFile = FreeFile
    Open OldResourceFilePath For Binary Access Read Lock Write As OldResourceFile
        Get OldResourceFile, 1, OldFileHead
        If LOF(OldResourceFile) <> OldFileHead.lngFileSize Then
            Call MsgBox(JsonLanguage.item("ERROR_ARCHIVO_CORRUPTO").item("TEXTO") & ": " & OldResourceFilePath, , JsonLanguage.item("Error").item("TEXTO"))
            Close OldResourceFile
            Exit Function
        End If
        NewResourceFile = FreeFile()
        Open NewResourceFilePath For Binary Access Read Lock Write As NewResourceFile
            Get NewResourceFile, 1, NewFileHead
            If LOF(NewResourceFile) <> NewFileHead.lngFileSize Then
                Call MsgBox(JsonLanguage.item("ERROR_ARCHIVO_CORRUPTO").item("TEXTO") & ": " & NewResourceFilePath, , JsonLanguage.item("Error").item("TEXTO"))
                Close NewResourceFile
                Close OldResourceFile
                Exit Function
            End If
            If LenB(Dir$(OutputFilePath, vbNormal)) <> 0 Then Kill OutputFilePath
            OutputFile = FreeFile()
            Open OutputFilePath For Binary Access Read Write As OutputFile
                If Not prgBar Is Nothing Then
                    prgBar.Value = 0
                    prgBar.max = (OldFileHead.lngNumFiles + NewFileHead.lngNumFiles) + 1
                End If
                Put OutputFile, , OldFileHead.lngFileVersion
                Put OutputFile, , NewFileHead
                If ReadNext_InfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles) _
                  And ReadNext_InfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles) Then
                    prgBar.Value = prgBar.Value + 2
                    Do
                        If OldInfoHead.strFileName = NewInfoHead.strFileName Then
                            Call Get_File_RawData(OldResourcePath, OldInfoHead, auxData, Modo)
                            Call Get_File_RawData(NewResourcePath, NewInfoHead, data, Modo)
                            If Not Compare_Datas(data, auxData) Then
                                Instruction = PatchInstruction.Modify_File
                                Put OutputFile, , Instruction
                                Put OutputFile, , NewInfoHead
                                Put OutputFile, , data
                            End If
                            If Not ReadNext_InfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles) Then
                                Exit Do
                            End If
                            If Not ReadNext_InfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles) Then
                                OldReadFiles = OldReadFiles - 1
                                Exit Do
                            End If
                            If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 2
                        ElseIf OldInfoHead.strFileName < NewInfoHead.strFileName Then
                            Instruction = PatchInstruction.Delete_File
                            Put OutputFile, , Instruction
                            Put OutputFile, , OldInfoHead
                            If Not ReadNext_InfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles) Then
                                NewReadFiles = NewReadFiles - 1
                                Exit Do
                            End If
                            If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1
                        Else
                            Instruction = PatchInstruction.Create_File
                            Put OutputFile, , Instruction
                            Put OutputFile, , NewInfoHead
                            Call Get_File_RawData(NewResourcePath, NewInfoHead, data, Modo)
                            Put OutputFile, , data
                            If Not ReadNext_InfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles) Then
                                OldReadFiles = OldReadFiles - 1
                                Exit Do
                            End If
                            If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1
                        End If
                        DoEvents
                    Loop
                Else
                    OldReadFiles = 0
                    NewReadFiles = 0
                End If
                While ReadNext_InfoHead(OldResourceFile, OldFileHead, OldInfoHead, OldReadFiles)
                    Instruction = PatchInstruction.Delete_File
                    Put OutputFile, , Instruction
                    Put OutputFile, , OldInfoHead
                    If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1
                    DoEvents
                Wend
                While ReadNext_InfoHead(NewResourceFile, NewFileHead, NewInfoHead, NewReadFiles)
                    Instruction = PatchInstruction.Create_File
                    Put OutputFile, , Instruction
                    Put OutputFile, , NewInfoHead
                    Call Get_File_RawData(NewResourcePath, NewInfoHead, data, Modo)
                    Put OutputFile, , data
                    If Not prgBar Is Nothing Then prgBar.Value = prgBar.Value + 1
                    DoEvents
                Wend
            Close OutputFile
        Close NewResourceFile
    Close OldResourceFile
    Make_Patch = True
Exit Function
ErrorHandler:
    Close OutputFile
    Close NewResourceFile
    Close OldResourceFile
    Call MsgBox(Replace$(JsonLanguage.item("ERROR_CREAR_PARCHE").item("TEXTO"), "VAR_ERROR", Err.number & " : " & Err.Description), vbOKOnly, JsonLanguage.item("Error").item("TEXTO"))
End Function

Public Function Apply_Patch(ByRef ResourcePath As String, ByRef PatchPath As String, ByRef prgBar As ProgressBar, Optional Modo As Byte = 0) As Boolean
    Dim ResourceFile As Integer
    Dim ResourceFilePath As String
    Dim FileHead As FILEHEADER
    Dim InfoHead As INFOHEADER
    Dim ResourceReadFiles As Long
    Dim EOResource As Boolean
    Dim PatchFile As Integer
    Dim PatchFilePath As String
    Dim PatchFileHead As FILEHEADER
    Dim PatchInfoHead As INFOHEADER
    Dim Instruction As Byte
    Dim OldResourceVersion As Long
    Dim OutputFile As Integer
    Dim OutputFilePath As String
    Dim data() As Byte
    Dim WrittenFiles As Long
    Dim DataOutputPos As Long
On Local Error GoTo ErrorHandler
    If Modo = 0 Then
        ResourceFilePath = ResourcePath & GRH_RESOURCE_FILE
        PatchFilePath = PatchPath & GRH_PATCH_FILE
        OutputFilePath = ResourcePath & GRH_RESOURCE_FILE & "tmp"
    ElseIf Modo = 1 Then
        ResourceFilePath = ResourcePath & MAPS_RESOURCE_FILE
        PatchFilePath = PatchPath & MAPS_PATCH_FILE
        OutputFilePath = ResourcePath & MAPS_RESOURCE_FILE & "tmp"
    End If
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        Get ResourceFile, , FileHead
        If LOF(ResourceFile) <> FileHead.lngFileSize Then
            Call MsgBox(JsonLanguage.item("ERROR_ARCHIVO_CORRUPTO").item("TEXTO") & ": " & ResourceFilePath, , JsonLanguage.item("Error").item("TEXTO"))
            Close ResourceFile
            Exit Function
        End If
        PatchFile = FreeFile()
        Open PatchFilePath For Binary Access Read Lock Write As PatchFile
            Get PatchFile, , OldResourceVersion
            If OldResourceVersion <> FileHead.lngFileVersion Then
                Call MsgBox(JsonLanguage.item("ERROR_VERSIONES_RECURSOS").item("TEXTO"), , JsonLanguage.item("Error").item("TEXTO"))
                Close ResourceFile
                Close PatchFile
                Exit Function
            End If
            Get PatchFile, , PatchFileHead
            If FileExist(OutputFilePath, vbNormal) Then Call Kill(OutputFilePath)
            OutputFile = FreeFile()
            Open OutputFilePath For Binary Access Read Write As OutputFile
                Put OutputFile, , PatchFileHead
                If Not prgBar Is Nothing Then
                    prgBar.Value = 0
                    prgBar.max = PatchFileHead.lngNumFiles + 1
                End If
                DataOutputPos = Len(FileHead) + Len(InfoHead) * PatchFileHead.lngNumFiles + 1
                While Loc(PatchFile) < LOF(PatchFile)
                    Get PatchFile, , Instruction
                    Get PatchFile, , PatchInfoHead
                    Do
                        EOResource = Not ReadNext_InfoHead(ResourceFile, FileHead, InfoHead, ResourceReadFiles)
                        If Not EOResource And InfoHead.strFileName < PatchInfoHead.strFileName Then
                            Call Get_File_RawData(ResourcePath, InfoHead, data, Modo)
                            InfoHead.lngFileStart = DataOutputPos
                            Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, InfoHead
                            Put OutputFile, DataOutputPos, data
                            DataOutputPos = DataOutputPos + UBound(data) + 1
                            WrittenFiles = WrittenFiles + 1
                            If Not prgBar Is Nothing Then prgBar.Value = WrittenFiles
                        Else
                            Exit Do
                        End If
                    Loop
                    Select Case Instruction
                        Case PatchInstruction.Delete_File
                            If InfoHead.strFileName <> PatchInfoHead.strFileName Then
                                Err.Description = JsonLanguage.item("ERROR_VERSIONES_RECURSOS").item("TEXTO")
                                GoTo errhandler
                            End If
                        
                        Case PatchInstruction.Create_File
                            If (InfoHead.strFileName > PatchInfoHead.strFileName) Or EOResource Then
                                ReDim data(PatchInfoHead.lngFileSize - 1)
                                Get PatchFile, , data
                                Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, PatchInfoHead
                                Put OutputFile, DataOutputPos, data
                                EOResource = False
                                ResourceReadFiles = ResourceReadFiles - 1
                                DataOutputPos = DataOutputPos + UBound(data) + 1
                                WrittenFiles = WrittenFiles + 1
                                If Not prgBar Is Nothing Then prgBar.Value = WrittenFiles
                            Else
                                Err.Description = JsonLanguage.item("ERROR_VERSIONES_RECURSOS").item("TEXTO")
                                GoTo errhandler
                            End If
                        
                        Case PatchInstruction.Modify_File
                            If InfoHead.strFileName = PatchInfoHead.strFileName Then
                                ReDim data(PatchInfoHead.lngFileSize - 1)
                                Get PatchFile, , data
                                Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, PatchInfoHead
                                Put OutputFile, DataOutputPos, data
                                DataOutputPos = DataOutputPos + UBound(data) + 1
                                WrittenFiles = WrittenFiles + 1
                                If Not prgBar Is Nothing Then prgBar.Value = WrittenFiles
                            Else
                                Err.Description = JsonLanguage.item("ERROR_VERSIONES_RECURSOS").item("TEXTO")
                                GoTo errhandler
                            End If
                    End Select
                    DoEvents
                Wend
                While ReadNext_InfoHead(ResourceFile, FileHead, InfoHead, ResourceReadFiles)
                    Call Get_File_RawData(ResourcePath, InfoHead, data, Modo)
                    InfoHead.lngFileStart = DataOutputPos
                    Put OutputFile, Len(FileHead) + Len(InfoHead) * WrittenFiles + 1, InfoHead
                    Put OutputFile, DataOutputPos, data
                    DataOutputPos = DataOutputPos + UBound(data) + 1
                    WrittenFiles = WrittenFiles + 1
                    If Not prgBar Is Nothing Then prgBar.Value = WrittenFiles
                    DoEvents
                Wend
            Close OutputFile
        Close PatchFile
    Close ResourceFile
    If (PatchFileHead.lngNumFiles = WrittenFiles) Then
            Call Kill(ResourceFilePath)
            Name OutputFilePath As ResourceFilePath
    Else
        Err.Description = JsonLanguage.item("ERROR_LEER_PARCHE").item("TEXTO")
        GoTo errhandler
    End If
    Apply_Patch = True
Exit Function
ErrorHandler:
    Close OutputFile
    Close PatchFile
    Close ResourceFile
    If FileExist(OutputFilePath, vbNormal) Then Call Kill(OutputFilePath)
    Call MsgBox(Replace$(JsonLanguage.item("ERROR_CREAR_PARCHE").item("TEXTO"), "VAR_ERROR", Err.number & " : " & Err.Description), vbOKOnly, JsonLanguage.item("Error").item("TEXTO"))
End Function

Private Function AlignScan(ByVal inWidth As Long, ByVal inDepth As Integer) As Long
    AlignScan = (((inWidth * inDepth) + &H1F) And Not &H1F&) \ &H8
End Function

Public Function GetVersion(ByVal ResourceFilePath As String) As Long
    Dim ResourceFile As Integer
    Dim FileHead As FILEHEADER
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        Get ResourceFile, 1, FileHead
    Close ResourceFile
    GetVersion = FileHead.lngFileVersion
End Function
