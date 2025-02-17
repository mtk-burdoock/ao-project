VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const NUM_CUSTOM_KEYS As Byte = 27
Private Const NUM_CUSTOM_CONFIGS As Byte = 9
Private Const ERROR_NO_KEY_FILE As Long = vbObjectError + 16
Private Const KEY_CONFIG_FILE As String = "BindKeys.bin"
Private Const KEY_CONFIG_FILE_ALTERNATIVE As String = "BindAlternativeKeys.bin"
Private KeyboardConfigurationInUse As String
Private mappedKeys(0 To NUM_CUSTOM_CONFIGS, 1 To NUM_CUSTOM_KEYS) As Integer
Private ActiveConfig As Byte

Public Enum eKeyType
    mKeyUp = 1
    mKeyDown
    mKeyLeft
    mKeyRight
    mKeyToggleMusic
    mKeyToggleSound
    mKeyToggleFxs
    mKeyRequestRefresh
    mKeyToggleNames
    mKeyGetObject
    mKeyEquipObject
    mKeyTamAnimal
    mKeySteal
    mKeyToggleSafeMode
    mKeyToggleResuscitationSafe
    mKeyHide
    mKeyDropObject
    mKeyUseObject
    mKeyAttack
    mKeyTalk
    mKeyTalkWithGuild
    mKeyTakeScreenShot
    mKeyShowOptions
    mKeyMeditate
    mKeyCastSpellMacro
    mKeyWorkMacro
    mKeyExitGame
End Enum

Public Function ReadableName(ByVal KeyCode As Integer) As String
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyA To KeyCodeConstants.vbKeyZ
            ReadableName = Chr$(KeyCode)
        
        Case KeyCodeConstants.vbKeyNumpad0 To KeyCodeConstants.vbKeyNumpad9
            ReadableName = KeyCode - KeyCodeConstants.vbKeyNumpad0 & " (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyF1 To KeyCodeConstants.vbKeyF16
            ReadableName = "F" & (KeyCode - KeyCodeConstants.vbKeyF1 + 1)
        
        Case KeyCodeConstants.vbKeyMultiply
            ReadableName = "* (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyAdd
            ReadableName = "+ (teclado numerico)"
        
        Case KeyCodeConstants.vbKeySeparator
            ReadableName = "Enter (teclado numerico)"
        
        Case KeyCodeConstants.vbKeySubtract
            ReadableName = "- (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyDecimal
            ReadableName = ". (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyDivide
            ReadableName = "/ (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyShift
            ReadableName = "Shift"
        
        Case KeyCodeConstants.vbKeyControl
            ReadableName = "Control"
        
        Case KeyCodeConstants.vbKeyMenu
            ReadableName = "Alt"
        
        Case KeyCodeConstants.vbKeyPause
            ReadableName = "Pausa"
        
        Case KeyCodeConstants.vbKeyCapital
            ReadableName = "Blq Mayus"
        
        Case KeyCodeConstants.vbKeyEscape
            ReadableName = "Esc"
        
        Case KeyCodeConstants.vbKeyPageUp
            ReadableName = "Ret Pag"
        
        Case KeyCodeConstants.vbKeyPageDown
            ReadableName = "Av Pag"
        
        Case KeyCodeConstants.vbKeyEnd
            ReadableName = "Fin"
        
        Case KeyCodeConstants.vbKeyHome
            ReadableName = "Inicio"
        
        Case KeyCodeConstants.vbKeyLeft
            ReadableName = "Izquierda"
        
        Case KeyCodeConstants.vbKeyUp
            ReadableName = "Arriba"
        
        Case KeyCodeConstants.vbKeyRight
            ReadableName = "Derecha"
        
        Case KeyCodeConstants.vbKeyDown
            ReadableName = "Abajo"
        
        Case KeyCodeConstants.vbKeyInsert
            ReadableName = "Insertar"
        
        Case KeyCodeConstants.vbKeyNumlock
            ReadableName = "Blq Num"
        
        Case KeyCodeConstants.vbKeyScrollLock
            ReadableName = "Blq Despl"
        
        Case KeyCodeConstants.vbKeySpace
            ReadableName = "Barra Espaciadora"
        
        Case KeyCodeConstants.vbKeyBack
            ReadableName = "Borrar"
        
        Case KeyCodeConstants.vbKeyReturn
            ReadableName = "Enter"
        
        Case KeyCodeConstants.vbKeyDelete
            ReadableName = "Suprimir"
        
        Case Else
            ReadableName = Constants.vbNullString
    End Select
End Function

Public Sub LoadCustomKeys()
    Dim i As Long
    Dim J As Long
    Dim FileNum As Integer
    If Not FileExist(Game.path(INIT) & KeyboardConfigurationInUse, vbArchive) Then
        For i = 0 To NUM_CUSTOM_CONFIGS
            Call LoadDefaults(i)
        Next i
        ActiveConfig = 1
        Call SaveCustomKeys
    Else
        FileNum = FreeFile
        Open Game.path(INIT) & KeyboardConfigurationInUse For Binary Access Read Lock Read Write As FileNum
        Get FileNum, , ActiveConfig
        For i = 0 To NUM_CUSTOM_CONFIGS
            For J = 1 To NUM_CUSTOM_KEYS
                Get FileNum, , mappedKeys(i, J)
            Next J
        Next i
        Close FileNum
        Call SaveCustomKeys
    End If
End Sub

Public Sub SaveCustomKeys()
    Dim i As Long
    Dim J As Long
    Dim FileNum As Integer
    FileNum = FreeFile()
    Open Game.path(INIT) & KeyboardConfigurationInUse For Binary Access Write As FileNum
    Put FileNum, , ActiveConfig
    For i = 0 To NUM_CUSTOM_CONFIGS
        For J = 1 To NUM_CUSTOM_KEYS
            Put FileNum, , mappedKeys(i, J)
        Next J
    Next i
    Close FileNum
End Sub

Public Property Get BindedKey(ByVal Index As eKeyType) As Integer
    If Index < 1 Or Index > NUM_CUSTOM_KEYS Then Exit Property
    BindedKey = mappedKeys(ActiveConfig, Index)
End Property

Public Property Let BindedKey(ByVal Index As eKeyType, ByVal NewVal As Integer)
    If Index < 1 Or Index > NUM_CUSTOM_KEYS Then Exit Property
    If LenB(ReadableName(NewVal)) = 0 Then
        Exit Property
    End If
    If KeyAssigned(NewVal) Then
        Exit Property
    End If
    mappedKeys(ActiveConfig, Index) = NewVal
End Property

Public Sub LoadDefaults(ByVal KeyConfig As Byte)
    mappedKeys(KeyConfig, eKeyType.mKeyUp) = KeyCodeConstants.vbKeyUp
    mappedKeys(KeyConfig, eKeyType.mKeyDown) = KeyCodeConstants.vbKeyDown
    mappedKeys(KeyConfig, eKeyType.mKeyLeft) = KeyCodeConstants.vbKeyLeft
    mappedKeys(KeyConfig, eKeyType.mKeyRight) = KeyCodeConstants.vbKeyRight
    mappedKeys(KeyConfig, eKeyType.mKeyToggleMusic) = KeyCodeConstants.vbKeyM
    mappedKeys(KeyConfig, eKeyType.mKeyToggleSound) = KeyCodeConstants.vbKeyS
    mappedKeys(KeyConfig, eKeyType.mKeyToggleFxs) = KeyCodeConstants.vbKeyF
    mappedKeys(KeyConfig, eKeyType.mKeyRequestRefresh) = KeyCodeConstants.vbKeyL
    mappedKeys(KeyConfig, eKeyType.mKeyToggleNames) = KeyCodeConstants.vbKeyN
    mappedKeys(KeyConfig, eKeyType.mKeyGetObject) = KeyCodeConstants.vbKeyA
    mappedKeys(KeyConfig, eKeyType.mKeyEquipObject) = KeyCodeConstants.vbKeyE
    mappedKeys(KeyConfig, eKeyType.mKeyTamAnimal) = KeyCodeConstants.vbKeyD
    mappedKeys(KeyConfig, eKeyType.mKeySteal) = KeyCodeConstants.vbKeyR
    mappedKeys(KeyConfig, eKeyType.mKeyToggleSafeMode) = KeyCodeConstants.vbKeyC
    mappedKeys(KeyConfig, eKeyType.mKeyToggleResuscitationSafe) = KeyCodeConstants.vbKeyV
    mappedKeys(KeyConfig, eKeyType.mKeyHide) = KeyCodeConstants.vbKeyO
    mappedKeys(KeyConfig, eKeyType.mKeyDropObject) = KeyCodeConstants.vbKeyT
    mappedKeys(KeyConfig, eKeyType.mKeyUseObject) = KeyCodeConstants.vbKeyU
    mappedKeys(KeyConfig, eKeyType.mKeyAttack) = KeyCodeConstants.vbKeyControl
    mappedKeys(KeyConfig, eKeyType.mKeyTalk) = KeyCodeConstants.vbKeyReturn
    mappedKeys(KeyConfig, eKeyType.mKeyTalkWithGuild) = KeyCodeConstants.vbKeyDelete
    mappedKeys(KeyConfig, eKeyType.mKeyTakeScreenShot) = KeyCodeConstants.vbKeyF2
    mappedKeys(KeyConfig, eKeyType.mKeyShowOptions) = KeyCodeConstants.vbKeyF5
    mappedKeys(KeyConfig, eKeyType.mKeyMeditate) = KeyCodeConstants.vbKeyF6
    mappedKeys(KeyConfig, eKeyType.mKeyCastSpellMacro) = KeyCodeConstants.vbKeyF7
    mappedKeys(KeyConfig, eKeyType.mKeyWorkMacro) = KeyCodeConstants.vbKeyF8
    mappedKeys(KeyConfig, eKeyType.mKeyExitGame) = KeyCodeConstants.vbKeyF12
End Sub

Public Function KeyAssigned(ByVal KeyCode As Integer) As Boolean
    Dim Counter As Long
    KeyAssigned = False
    For Counter = 1 To NUM_CUSTOM_KEYS
        If mappedKeys(ActiveConfig, Counter) = KeyCode Then
            KeyAssigned = True
            Exit For
        End If
    Next Counter
End Function

Public Property Get KeyCount() As Byte
    KeyCount = NUM_CUSTOM_KEYS
End Property

Public Property Get ConfigCount() As Byte
    ConfigCount = NUM_CUSTOM_CONFIGS
End Property

Public Property Let CurrentConfig(ByVal NewVal As Byte)
    If NewVal < 0 Or NewVal > NUM_CUSTOM_CONFIGS Then Exit Property
    ActiveConfig = NewVal
End Property

Public Property Get CurrentConfig() As Byte
    CurrentConfig = ActiveConfig
End Property

Private Sub Class_Initialize()
    If ClientSetup.KeyboardBindKeysConfig = "Alternative" Then
        KeyboardConfigurationInUse = KEY_CONFIG_FILE_ALTERNATIVE
    Else
        KeyboardConfigurationInUse = KEY_CONFIG_FILE
    End If
    Call LoadCustomKeys
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    Call SaveCustomKeys
End Sub

Public Sub SetKeyConfigFileInUse(ByVal KeyConfigToUse As String)
    If KeyConfigToUse = "Alternative" Then
        KeyboardConfigurationInUse = KEY_CONFIG_FILE_ALTERNATIVE
    Else
        KeyboardConfigurationInUse = KEY_CONFIG_FILE
    End If
    Call LoadCustomKeys
    ClientSetup.KeyboardBindKeysConfig = KeyConfigToUse
    Call WriteVar(Game.path(INIT) & "Config.ini", "OTHER", "BIND_KEYS", KeyConfigToUse)
End Sub
