Attribute VB_Name = "Resolution"
Option Explicit

Public ResolucionCambiada As Boolean
Private Const CCDEVICENAME As Long = 32
Private Const CCFORMNAME As Long = 32
Private Const DM_BITSPERPEL As Long = &H40000
Private Const DM_PELSWIDTH As Long = &H80000
Private Const DM_PELSHEIGHT As Long = &H100000
Private Const DM_DISPLAYFREQUENCY As Long = &H400000
Private Const CDS_TEST As Long = &H4
Private Const ENUM_CURRENT_SETTINGS As Long = -1
Private MiDevM As typDevMODE
Private oldResHeight As Long
Private oldResWidth As Long
Private oldDepth As Integer
Private oldFrequency As Long
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long

Private Type typDevMODE
    dmDeviceName       As String * CCDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type

Public Sub SetResolution(ByRef newWidth As Integer, ByRef newHeight As Integer)
    Dim lRes As Long: lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, MiDevM)
    oldResWidth = Screen.Width \ Screen.TwipsPerPixelX
    oldResHeight = Screen.Height \ Screen.TwipsPerPixelY
    If oldResWidth <> newWidth Or oldResHeight <> newHeight Then
        If ClientSetup.bFullScreen Then
            With MiDevM
                .dmBitsPerPel = 32
                .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
                .dmPelsWidth = newWidth
                .dmPelsHeight = newHeight
                oldDepth = .dmBitsPerPel
                oldFrequency = .dmDisplayFrequency
            End With
            lRes = ChangeDisplaySettings(MiDevM, CDS_TEST)
            ResolucionCambiada = True
            frmMain.WindowState = vbMaximized
        Else
            ResolucionCambiada = False
            frmMain.WindowState = vbNormal
        End If
    End If
End Sub

Public Sub ResetResolution()
    Dim lRes As Long: lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, MiDevM)
    With MiDevM
        .dmFields = DM_PELSWIDTH And DM_PELSHEIGHT And DM_BITSPERPEL And DM_DISPLAYFREQUENCY
        .dmPelsWidth = oldResWidth
        .dmPelsHeight = oldResHeight
        .dmBitsPerPel = oldDepth
        .dmDisplayFrequency = oldFrequency
    End With
    lRes = ChangeDisplaySettings(MiDevM, CDS_TEST)
    ResolucionCambiada = False
End Sub
