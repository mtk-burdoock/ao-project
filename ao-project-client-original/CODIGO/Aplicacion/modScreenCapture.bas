Attribute VB_Name = "modScreenCapture"
Option Explicit

Private Enum IJLERR
  IJL_OK = 0
  IJL_INTERRUPT_OK = 1
  IJL_ROI_OK = 2
  IJL_EXCEPTION_DETECTED = -1
  IJL_INVALID_ENCODER = -2
  IJL_UNSUPPORTED_SUBSAMPLING = -3
  IJL_UNSUPPORTED_BYTES_PER_PIXEL = -4
  IJL_MEMORY_ERROR = -5
  IJL_BAD_HUFFMAN_TABLE = -6
  IJL_BAD_QUANT_TABLE = -7
  IJL_INVALID_JPEG_PROPERTIES = -8
  IJL_ERR_FILECLOSE = -9
  IJL_INVALID_FILENAME = -10
  IJL_ERROR_EOF = -11
  IJL_PROG_NOT_SUPPORTED = -12
  IJL_ERR_NOT_JPEG = -13
  IJL_ERR_COMP = -14
  IJL_ERR_SOF = -15
  IJL_ERR_DNL = -16
  IJL_ERR_NO_HUF = -17
  IJL_ERR_NO_QUAN = -18
  IJL_ERR_NO_FRAME = -19
  IJL_ERR_MULT_FRAME = -20
  IJL_ERR_DATA = -21
  IJL_ERR_NO_IMAGE = -22
  IJL_FILE_ERROR = -23
  IJL_INTERNAL_ERROR = -24
  IJL_BAD_RST_MARKER = -25
  IJL_THUMBNAIL_DIB_TOO_SMALL = -26
  IJL_THUMBNAIL_DIB_WRONG_COLOR = -27
  IJL_RESERVED = -99
End Enum

Private Enum IJLIOTYPE
  IJL_SETUP = -1&
  IJL_JFILE_READPARAMS = 0&
  IJL_JBUFF_READPARAMS = 1&
  IJL_JFILE_READWHOLEIMAGE = 2&
  IJL_JBUFF_READWHOLEIMAGE = 3&
  IJL_JFILE_READHEADER = 4&
  IJL_JBUFF_READHEADER = 5&
  IJL_JFILE_READENTROPY = 6&
  IJL_JBUFF_READENTROPY = 7&
  IJL_JFILE_WRITEWHOLEIMAGE = 8&
  IJL_JBUFF_WRITEWHOLEIMAGE = 9&
  IJL_JFILE_WRITEHEADER = 10&
  IJL_JBUFF_WRITEHEADER = 11&
  IJL_JFILE_WRITEENTROPY = 12&
  IJL_JBUFF_WRITEENTROPY = 13&
  IJL_JFILE_READONEHALF = 14&
  IJL_JBUFF_READONEHALF = 15&
  IJL_JFILE_READONEQUARTER = 16&
  IJL_JBUFF_READONEQUARTER = 17&
  IJL_JFILE_READONEEIGHTH = 18&
  IJL_JBUFF_READONEEIGHTH = 19&
  IJL_JFILE_READTHUMBNAIL = 20&
  IJL_JBUFF_READTHUMBNAIL = 21&
End Enum

Private Type JPEG_CORE_PROPERTIES_VB
  UseJPEGPROPERTIES As Long
  DIBBytes As Long
  DIBWidth As Long
  DIBHeight As Long
  DIBPadBytes As Long
  DIBChannels As Long
  DIBColor As Long
  DIBSubsampling As Long
  JPGFile As Long
  JPGBytes As Long
  JPGSizeBytes As Long
  JPGWidth As Long
  JPGHeight As Long
  JPGChannels As Long
  JPGColor As Long
  JPGSubsampling As Long
  JPGThumbWidth As Long
  JPGThumbHeight As Long
  cconversion_reqd As Long
  upsampling_reqd As Long
  jquality As Long
  jprops(0 To 19999) As Byte
End Type

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef source As Any, ByVal byteCount As Long)
Private Declare Function ijlInit Lib "ijl11.dll" (jcprops As Any) As Long
Private Declare Function ijlFree Lib "ijl11.dll" (jcprops As Any) As Long
Private Declare Function ijlRead Lib "ijl11.dll" (jcprops As Any, ByVal ioType As Long) As Long
Private Declare Function ijlWrite Lib "ijl11.dll" (jcprops As Any, ByVal ioType As Long) As Long
Private Const MAX_PATH = 260

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function lopen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Private Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Const OF_WRITE = &H1
Private Const OF_SHARE_DENY_WRITE = &H20
Private Const INVALID_HANDLE As Long = -1
Private Const SRCCOPY = &HCC0020
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Function LoadJPG(ByRef cDib As cDIBSection, ByVal sFile As String) As Boolean
    Dim tJ        As JPEG_CORE_PROPERTIES_VB
    Dim bFile()   As Byte
    Dim lR        As Long
    Dim lPtr      As Long
    Dim lJPGWidth As Long, lJPGHeight As Long
    lR = ijlInit(tJ)
    If lR = IJL_OK Then
        bFile = StrConv(sFile, vbFromUnicode)
        ReDim Preserve bFile(0 To UBound(bFile) + 1) As Byte
        bFile(UBound(bFile)) = 0
        lPtr = VarPtr(bFile(0))
        Call CopyMemory(tJ.JPGFile, lPtr, 4)
        lR = ijlRead(tJ, IJL_JFILE_READPARAMS)
        If lR <> IJL_OK Then
            Call MsgBox("Failed to read JPG", vbExclamation)
        Else
            If tJ.JPGChannels = 1 Then
                tJ.JPGColor = 4&
            Else
                tJ.JPGColor = 3&
            End If
            lJPGWidth = tJ.JPGWidth
            lJPGHeight = tJ.JPGHeight
            If cDib.Create(lJPGWidth, lJPGHeight) Then
                tJ.DIBWidth = lJPGWidth
                tJ.DIBPadBytes = cDib.BytesPerScanLine - lJPGWidth * 3
                tJ.DIBHeight = -lJPGHeight
                tJ.DIBChannels = 3&
                tJ.DIBBytes = cDib.DIBSectionBitsPtr
                lR = ijlRead(tJ, IJL_JFILE_READWHOLEIMAGE)
                If lR = IJL_OK Then
                    LoadJPG = True
                Else
                    Call MsgBox("Cannot read Image Data from file.", vbExclamation)
                End If
            Else
            End If
        End If
        Call ijlFree(tJ)
    Else
        Call MsgBox("Failed to initialise the IJL library: " & lR, vbExclamation)
    End If
End Function

Public Function LoadJPGFromPtr(ByRef cDib As cDIBSection, ByVal lPtr As Long, ByVal lSize As Long) As Boolean
    Dim tJ        As JPEG_CORE_PROPERTIES_VB
    Dim lR        As Long
    Dim lJPGWidth As Long, lJPGHeight As Long
    lR = ijlInit(tJ)
    If lR = IJL_OK Then
        tJ.JPGBytes = lPtr
        tJ.JPGSizeBytes = lSize
        lR = ijlRead(tJ, IJL_JBUFF_READPARAMS)
        If lR <> IJL_OK Then
            MsgBox "Failed to read JPG", vbExclamation
        Else
            If tJ.JPGChannels = 1 Then
                tJ.JPGColor = 4&
            Else
                tJ.JPGColor = 3&
            End If
            lJPGWidth = tJ.JPGWidth
            lJPGHeight = tJ.JPGHeight
            If cDib.Create(lJPGWidth, lJPGHeight) Then
                tJ.DIBWidth = lJPGWidth
                tJ.DIBPadBytes = cDib.BytesPerScanLine - lJPGWidth * 3
                tJ.DIBHeight = -lJPGHeight
                tJ.DIBChannels = 3&
                tJ.DIBBytes = cDib.DIBSectionBitsPtr
                lR = ijlRead(tJ, IJL_JBUFF_READWHOLEIMAGE)
                If lR = IJL_OK Then
                    LoadJPGFromPtr = True
                Else
                    Call MsgBox("Cannot read Image Data from file.", vbExclamation)
                End If
            Else
            End If
        End If
        Call ijlFree(tJ)
    Else
        Call MsgBox("Failed to initialise the IJL library: " & lR, vbExclamation)
    End If
End Function

Public Function SaveJPG(ByRef cDib As cDIBSection, ByVal sFile As String, Optional ByVal lQuality As Long = 90) As Boolean
    Dim tJ           As JPEG_CORE_PROPERTIES_VB
    Dim bFile()      As Byte
    Dim lPtr         As Long
    Dim lR           As Long
    Dim tFnd         As WIN32_FIND_DATA
    Dim hFile        As Long
    Dim bFileExisted As Boolean
    Dim lFileSize    As Long
    hFile = -1
    lR = ijlInit(tJ)
    If lR = IJL_OK Then
        bFileExisted = (FindFirstFile(sFile, tFnd) <> -1)
        If bFileExisted Then
            Kill sFile
        End If
        tJ.DIBWidth = cDib.Width
        tJ.DIBHeight = -cDib.Height
        tJ.DIBBytes = cDib.DIBSectionBitsPtr
        tJ.DIBPadBytes = cDib.BytesPerScanLine - cDib.Width * 3
        bFile = StrConv(sFile, vbFromUnicode)
        ReDim Preserve bFile(0 To UBound(bFile) + 1) As Byte
        bFile(UBound(bFile)) = 0
        lPtr = VarPtr(bFile(0))
        Call CopyMemory(tJ.JPGFile, lPtr, 4)
        tJ.JPGWidth = cDib.Width
        tJ.JPGHeight = cDib.Height
        tJ.jquality = lQuality
        lR = ijlWrite(tJ, IJL_JFILE_WRITEWHOLEIMAGE)
        If lR = IJL_OK Then
            If bFileExisted Then
                hFile = lopen(sFile, OF_WRITE Or OF_SHARE_DENY_WRITE)
                If hFile = 0 Then
                Else
                    Call SetFileTime(hFile, tFnd.ftCreationTime, tFnd.ftLastAccessTime, tFnd.ftLastWriteTime)
                    Call lclose(hFile)
                    Call SetFileAttributes(bFile, tFnd.dwFileAttributes)
                End If
            End If
            lFileSize = tJ.JPGSizeBytes - tJ.JPGBytes
            SaveJPG = True
        Else
            Call Err.Raise(26001, JsonLanguage.item("ERROR_GUARDAR_SCREENSHOT").item("TEXTO") & lR, vbExclamation)
        End If
        Call ijlFree(tJ)
    Else
        Call Err.Rais(26001, App.EXEName & ".mIntelJPEGLibrary", "No se pudo inicializar la Libreria " & lR)
    End If
End Function

Public Function SaveJPGToPtr(ByRef cDib As cDIBSection, ByVal lPtr As Long, ByRef lBufSize As Long, Optional ByVal lQuality As Long = 90) As Boolean
    Dim tJ    As JPEG_CORE_PROPERTIES_VB
    Dim lR    As Long
    Dim hFile As Long
    hFile = -1
    lR = ijlInit(tJ)
    If lR = IJL_OK Then
        tJ.DIBWidth = cDib.Width
        tJ.DIBHeight = -cDib.Height
        tJ.DIBBytes = cDib.DIBSectionBitsPtr
        tJ.DIBPadBytes = cDib.BytesPerScanLine - cDib.Width * 3
        tJ.JPGWidth = cDib.Width
        tJ.JPGHeight = cDib.Height
        tJ.jquality = lQuality
        tJ.JPGBytes = lPtr
        tJ.JPGSizeBytes = lBufSize
        lR = ijlWrite(tJ, IJL_JBUFF_WRITEWHOLEIMAGE)
        If lR = IJL_OK Then
            lBufSize = tJ.JPGSizeBytes
            SaveJPGToPtr = True
        Else
            Call Err.Raise(26001, App.EXEName & ".mIntelJPEGLibrary", "Failed to save to JPG " & lR, vbExclamation)
        End If
        Call ijlFree(tJ)
    Else
        Call Err.Raise(26001, App.EXEName & ".mIntelJPEGLibrary", "Failed to initialise the IJL library: " & lR)
    End If
End Function

Public Sub ScreenCapture(Optional ByVal Autofragshooter As Boolean = False)
    On Error GoTo ErrorHandler:
    Dim File As String
    Dim c    As cDIBSection
    Set c = New cDIBSection
    Dim hdcc    As Long
    Dim dirFile As String
    hdcc = GetDC(frmMain.hwnd)
    Dim FileName As String
        FileName = Format$(Now, "DD-MM-YYYY hh-mm-ss") & ".jpg"
    With frmScreenshots.Picture1
        .AutoRedraw = True
        .Width = frmMain.Width
        .Height = frmMain.Height
        Call BitBlt(.hdc, 0, 0, frmMain.Width, frmMain.Height, hdcc, 0, 0, SRCCOPY)
        Call ReleaseDC(frmMain.hwnd, hdcc)
        hdcc = INVALID_HANDLE
        dirFile = App.path & "\Screenshots"
        If Not FileExist(dirFile, vbDirectory) Then Call MkDir(dirFile)
        If Autofragshooter Then
            dirFile = dirFile & "\FragShooter"
            If Not FileExist(dirFile, vbDirectory) Then Call MkDir(dirFile)
            If FragShooterKilledSomeone Then
                dirFile = dirFile & "\Frags"
            Else
                dirFile = dirFile & "\Muertes"
            End If
            If Not FileExist(dirFile, vbDirectory) Then Call MkDir(dirFile)
            File = dirFile & "\" & FragShooterNickname & " -" & FileName
        Else
            File = dirFile & "\" & FileName
        End If
        Call .Refresh
        .Picture = .Image
        Call c.CreateFromPicture(.Picture)
        Call SaveJPG(c, File)
        Call AddtoRichTextBox(frmMain.RecTxt, "Screen Capturada!", 200, 200, 200, False, False, True)
        Exit Sub
    End With
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MOD_SCREENCAPTURE").item("TEXTO") & File, 100, 30, 20, False, False, True)
    Exit Sub
ErrorHandler:
    Call AddtoRichTextBox(frmMain.RecTxt, Err.number & "-" & Err.Description, 200, 200, 200, False, False, True)
    If hdcc <> INVALID_HANDLE Then Call ReleaseDC(frmMain.hwnd, hdcc)
End Sub

Public Function FullScreenCapture(ByVal File As String) As Boolean
    Dim c As cDIBSection
    Set c = New cDIBSection
    Dim hdcc   As Long
    Dim handle As Long
    hdcc = GetDC(handle)
    With frmScreenshots.Picture1
        .AutoRedraw = True
        If Not ResolucionCambiada Then
            .Width = Screen.Width
            .Height = Screen.Height
            Call BitBlt(frmScreenshots.Picture1.hdc, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, hdcc, 0, 0, SRCCOPY)
        Else
            .Width = frmMain.Width
            .Height = frmMain.Height
            Call BitBlt(.hdc, 0, 0, frmMain.Width, frmMain.Height, hdcc, 0, 0, SRCCOPY)
        End If
        Call ReleaseDC(handle, hdcc)
        hdcc = INVALID_HANDLE
        If Not FileExist(App.path & "\TEMP", vbDirectory) Then MkDir (App.path & "\TEMP")
        .Refresh
        .Picture = .Image
        Call c.CreateFromPicture(.Picture)
        Call SaveJPG(c, File)
        FullScreenCapture = True
    End With
End Function
