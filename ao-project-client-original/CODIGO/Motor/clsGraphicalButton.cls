VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents ButtonHolder As Image
Attribute ButtonHolder.VB_VarHelpID = -1
Private RollOverButtonImage As Picture
Private NormalButtonImage As Picture
Private ClickButtonImage As Picture
Private DisabledButtonImage As Picture
Private SoundClick As Boolean
Private FormActual As Form
Private OverButton As Boolean
Private Disabled As Boolean

Public Sub Initialize(ByRef imgButton As Image, ByVal imgNormalPath As String, ByVal imgRollOverPath As String, ByVal imgClickPath As String, ByRef Formulario As Form, _
            Optional ByVal imgDisabledPath As String, Optional ByVal StartDisabled As Boolean = False, Optional ByVal MakeClickSound As Boolean = True, Optional ByVal ShowHandPointer As Boolean = True)
    Set ButtonHolder = imgButton
    Set FormActual = Formulario
    Set RollOverButtonImage = LoadPicture(imgRollOverPath)
    Set NormalButtonImage = LoadPicture(imgNormalPath)
    Set ClickButtonImage = LoadPicture(imgClickPath)
    Set DisabledButtonImage = LoadPicture(imgDisabledPath)
    SoundClick = MakeClickSound
    With ButtonHolder
        If ShowHandPointer Then
            .MousePointer = vbCustom
            .MouseIcon = picMouseIcon
        End If
        If StartDisabled Then
            Disabled = True
            .Picture = DisabledButtonImage
        Else
            .Picture = NormalButtonImage
        End If
    End With
End Sub

Private Sub ButtonHolder_Click()
    If Disabled Then Exit Sub
    If Not SoundClick Then Exit Sub
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub ButtonHolder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Disabled Then Exit Sub
    ButtonHolder.Picture = ClickButtonImage
End Sub

Private Sub ButtonHolder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Disabled Then Exit Sub
    If Not OverButton Then
        FormActual.LastButtonPressed.ToggleToNormal
        ButtonHolder.Picture = RollOverButtonImage
        OverButton = True
        Set FormActual.LastButtonPressed = Me
    End If
End Sub

Public Sub ToggleToNormal()
    If Disabled Then Exit Sub
    If Not OverButton Then Exit Sub
    ButtonHolder.Picture = NormalButtonImage
    OverButton = False
End Sub

Private Sub ButtonHolder_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Disabled Then Exit Sub
    ButtonHolder.Picture = RollOverButtonImage
End Sub

Public Sub EnableButton(Optional ByVal ToogleEnable As Boolean = True)
    If Disabled = Not ToogleEnable Then Exit Sub
    OverButton = False
    Disabled = Not ToogleEnable
    If Disabled Then
        ButtonHolder.Picture = DisabledButtonImage
    Else
        ButtonHolder.Picture = NormalButtonImage
    End If
End Sub

Public Function IsEnabled() As Boolean
    IsEnabled = Not Disabled
End Function
