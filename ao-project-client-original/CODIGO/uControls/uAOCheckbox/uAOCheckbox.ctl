VERSION 5.00
Begin VB.UserControl uAOCheckbox 
   ClientHeight    =   1185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2700
   DrawStyle       =   5  'Transparent
   ScaleHeight     =   1185
   ScaleWidth      =   2700
   ToolboxBitmap   =   "uAOCheckbox.ctx":0000
   Begin VB.Timer MouseO 
      Interval        =   3
      Left            =   2160
      Top             =   600
   End
   Begin VB.PictureBox bCheckbox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   177
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
   Begin VB.Image iCHK 
      Height          =   15
      Left            =   0
      Picture         =   "uAOCheckbox.ctx":0312
      Top             =   1080
      Width           =   15
   End
End
Attribute VB_Name = "uAOCheckbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private szCheckBoxW    As Integer
Private szCheckBoxH    As Integer
Private iCheckBoxLoaded  As Boolean
Private iCheckBox    As IPictureDisp
Private isOver      As Boolean
Private lastStat    As Integer
Private lastHwnd    As Long
Private lastButton  As Byte
Private lastKeyDown As Byte
Private isFocus     As Boolean
Private IsEnabled   As Boolean
Private isChecked   As Boolean

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseOver()
Public Event MouseOut()

Private Sub bCheckbox_Click()
On Error Resume Next
    If IsEnabled Then
        isChecked = Not isChecked
        Call Redraw(True)
        RaiseEvent Click
    End If
End Sub

Private Sub bCheckbox_DblClick()
On Error Resume Next
    If IsEnabled Then
        Call bCheckbox_MouseDown(1, 0, 0, 0)
        RaiseEvent DblClick
    End If
End Sub

Private Sub bCheckbox_GotFocus()
On Error Resume Next
    If isOver = False And IsEnabled Then
        isFocus = True
        Call Redraw(True)
    End If
End Sub

Private Sub bCheckbox_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    RaiseEvent KeyDown(KeyCode, Shift)
    lastKeyDown = KeyCode
    Call Redraw(False)
End Sub

Private Sub bCheckbox_KeyPress(KeyAscii As Integer)
On Error Resume Next
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub bCheckbox_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    RaiseEvent KeyUp(KeyCode, Shift)
    If (KeyCode = 32) And (lastKeyDown = 32) Then
        Call Redraw(False)
        Call bCheckbox_Click
        UserControl.Refresh
        RaiseEvent Click
    End If
End Sub

Private Sub bCheckbox_LostFocus()
On Error Resume Next
    If isOver = False And IsEnabled Then
        isFocus = False
        Call Redraw(True)
    End If
End Sub

Private Sub bCheckbox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    RaiseEvent MouseDown(Button, Shift, X, Y)
    lastButton = Button
    If lastButton <> 2 And IsEnabled Then
        Call Redraw(False)
    End If
End Sub

Private Sub bCheckbox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If lastButton < 2 And IsEnabled Then
        lastHwnd = bCheckbox.hwnd
        If Not isMouseOver Then
            Call Redraw(False)
        Else
            If Button = 0 And Not isOver Then
                MouseO.Enabled = True
                isOver = True
                Call Redraw(False)
                RaiseEvent MouseOver
            End If
        End If
    End If
End Sub

Private Sub bCheckbox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If lastButton <> 2 And IsEnabled Then
        Call Redraw(False)
    End If
    lastButton = 0
End Sub

Private Sub MouseO_Timer()
On Error Resume Next
    If Not isMouseOver Then
        isOver = False
        isFocus = False
        Call Redraw(True)
        MouseO.Enabled = False
        RaiseEvent MouseOut
    End If
End Sub

Private Function isMouseOver() As Boolean
On Error Resume Next
    Dim pt As POINTAPI
    GetCursorPos pt
    isMouseOver = (WindowFromPoint(pt.X, pt.Y) = lastHwnd)
End Function


Private Sub ReloadTextures()
On Error Resume Next
    Set iCheckBox = iCHK.Picture
    If iCHK.Picture.Height <> iCHK.Picture.Width Then
        szCheckBoxW = iCHK.Width / 6
        szCheckBoxH = iCHK.Height
        iCheckBoxLoaded = True
    Else
        iCheckBoxLoaded = False
    End If
End Sub

Private Sub Redraw(Optional Force As Boolean = False)
On Error Resume Next
    Dim szCheckBoxX As Integer
    Dim rI          As Integer
    Dim rY          As Integer
    If IsEnabled = False Then
        If isChecked = True Then
            If iCheckBoxLoaded = False Then
                bCheckbox.BackColor = RGB(10, 64, 10)
            End If
            szCheckBoxX = szCheckBoxW * 5
        Else
            If iCheckBoxLoaded = False Then
                bCheckbox.BackColor = RGB(10, 10, 10)
            End If
            szCheckBoxX = szCheckBoxW * 4
        End If
        MouseO.Enabled = False
    Else
        If isChecked = True Then
            If isOver = False And isFocus = False Then
                If iCheckBoxLoaded = False Then
                    bCheckbox.BackColor = RGB(32, 128, 32)
                End If
                szCheckBoxX = szCheckBoxW * 2
            Else
                If iCheckBoxLoaded = False Then
                    bCheckbox.BackColor = RGB(64, 128, 64)
                End If
                szCheckBoxX = szCheckBoxW * 3
            End If
            If isFocus = False Then MouseO.Enabled = True
        Else
            If isOver = False And isFocus = False Then
                If iCheckBoxLoaded = False Then
                    bCheckbox.BackColor = RGB(32, 32, 32)
                End If
                szCheckBoxX = 0
                MouseO.Enabled = False
            Else
                If iCheckBoxLoaded = False Then
                    bCheckbox.BackColor = RGB(64, 64, 64)
                End If
                szCheckBoxX = szCheckBoxW
                If isFocus = False Then MouseO.Enabled = True
            End If
        End If
    End If
    If lastStat = szCheckBoxX And Force = False Then
        Exit Sub
    Else
    End If
    lastStat = szCheckBoxX
    bCheckbox.Cls
    If iCheckBoxLoaded = True Then
        bCheckbox.PaintPicture iCheckBox, 0, 0, szCheckBoxW, szCheckBoxH, szCheckBoxX, 0, szCheckBoxW, szCheckBoxH
    End If
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
On Error Resume Next
    lastButton = 1
    Call UserControl_Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
On Error Resume Next
    Call Redraw(True)
End Sub

Private Sub UserControl_Click()
On Error Resume Next
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
On Error Resume Next
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
On Error Resume Next
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
    Call ReloadTextures
    IsEnabled = True
End Sub

Private Sub UserControl_InitProperties()
On Error Resume Next
    lastStat = 0
    IsEnabled = True
    iCheckBoxLoaded = False
    bCheckbox.BackColor = RGB(32, 32, 32)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
On Error Resume Next
    Call Redraw(True)
    isOver = False
End Sub

Private Sub UserControl_Show()
On Error Resume Next
    If iCheckBoxLoaded = True Then
        UserControl.Width = szCheckBoxW
        UserControl.Height = szCheckBoxH
    Else
        If UserControl.Width > 340 Then
            UserControl.Width = 340
            UserControl.Height = 340
        End If
    End If
    Call Redraw(False)
    isOver = False
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    If iCheckBoxLoaded = True Then
        UserControl.Width = szCheckBoxW
        UserControl.Height = szCheckBoxH
    Else
        If UserControl.Width > 340 Then
            UserControl.Width = 340
            UserControl.Height = 340
        End If
    End If
    bCheckbox.Width = UserControl.Width
    bCheckbox.Height = UserControl.Height
    bCheckbox.ScaleWidth = bCheckbox.Width
    bCheckbox.ScaleHeight = bCheckbox.Height
    Call Redraw(True)
    isOver = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    With PropBag
        IsEnabled = .ReadProperty("ENAB", True)
        isChecked = .ReadProperty("CHCK", False)
        iCHK.Picture = .ReadProperty("PICC", iCHK.Picture)
    End With
    UserControl.Enabled = IsEnabled
    Call ReloadTextures
    Call Redraw(True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
    With PropBag
        Call .WriteProperty("CHCK", isChecked)
        Call .WriteProperty("ENAB", IsEnabled)
        Call .WriteProperty("PICC", iCHK.Picture)
    End With
End Sub

Public Property Get Enabled() As Boolean
On Error Resume Next
    Enabled = IsEnabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
On Error Resume Next
    IsEnabled = NewValue
    isOver = False
    Call Redraw(False)
    UserControl.Enabled = IsEnabled
    PropertyChanged "ENAB"
End Property


Public Property Get Checked() As Boolean
On Error Resume Next
    Checked = isChecked
End Property

Public Property Let Checked(ByVal NewValue As Boolean)
On Error Resume Next
    isChecked = NewValue
    Call Redraw(True)
    PropertyChanged "CHCK"
End Property

Public Property Get Picture() As StdPicture
On Error Resume Next
    Set Picture = iCHK.Picture
End Property

Public Property Set Picture(ByVal newPic As StdPicture)
On Error Resume Next
    iCHK.Picture = newPic
    Call ReloadTextures
    Call UserControl_Resize
    Call Redraw(True)
    PropertyChanged "PICC"
End Property
