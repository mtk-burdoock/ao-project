Attribute VB_Name = "Mod_TCP"
Option Explicit
Public LlegaronSkills As Boolean
Public LlegaronAtrib As Boolean
Public LlegoFama As Boolean

Sub Login()
    Select Case EstadoLogin
        Case E_MODO.Normal
            Call WriteLoginExistingAccount
        
        Case E_MODO.CrearNuevoPj
            Call WriteLoginNewChar
            
        Case E_MODO.CrearCuenta
            Call WriteLoginNewAccount
        
        Case E_MODO.CambiarContrasena
            Call WriteCambiarContrasena
    End Select
    DoEvents
    Call FlushBuffer
End Sub
