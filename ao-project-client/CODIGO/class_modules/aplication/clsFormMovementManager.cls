VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormMovementManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents Formulario As Form
Attribute Formulario.VB_VarHelpID = -1
Private DragX As Single
Private Dragy As Single
Private LimitY As Integer

Public Sub Initialize(ByRef tForm As Form, Optional ByVal VerticalLimit As Integer = 0)
    Set Formulario = tForm
    LimitY = VerticalLimit
End Sub

Private Sub Formulario_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If LimitY > 0 Then
        If Y > LimitY Then Exit Sub
    End If
    DragX = X
    Dragy = Y
End Sub

Private Sub Formulario_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If LimitY > 0 Then
        If Y > LimitY Then Exit Sub
    End If
    If Button > 0 Then
        Formulario.Left = Formulario.Left + (X - DragX)
        Formulario.Top = Formulario.Top + (Y - Dragy)
    End If
End Sub
