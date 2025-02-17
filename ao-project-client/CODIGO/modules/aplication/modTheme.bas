Attribute VB_Name = "modTheme"
Option Explicit

Public SkinSeleccionado As String

Public Sub LoadTheme()
    SkinSeleccionado = GetVar(Game.path(INIT) & "Config.ini", "Parameters", "SkinSelected")
    If Not SkinSeleccionado <> "" Then
        SkinSeleccionado = "Libre"
    End If
    Select Case SkinSeleccionado
        Case "Libre"
            frmMain.Picture = LoadPicture(Game.pathTheme(Libre) & "VentanaPrincipal.jpg")
            frmPanelAccount.Picture = LoadPicture(Game.pathTheme(Libre) & "frmPanelAccount.jpg")
            frmCrearCuenta.Picture = LoadPicture(Game.pathTheme(Libre) & "frmCrearCuenta.jpg")
            Set frmMain.picSkillStar = LoadPicture(Game.pathTheme(Libre) & "BotonAsignarSkills.bmp")
            
        Case "Psicodelico"
            frmMain.Picture = LoadPicture(Game.pathTheme(Psicodelico) & "VentanaPrincipal.jpg")
            frmPanelAccount.Picture = LoadPicture(Game.pathTheme(Psicodelico) & "frmPanelAccount.jpg")
            frmCrearCuenta.Picture = LoadPicture(Game.pathTheme(Psicodelico) & "frmCrearCuenta.jpg")
            Set frmMain.picSkillStar = LoadPicture(Game.pathTheme(Psicodelico) & "BotonAsignarSkills.bmp")
        Case "Oscuro"
            frmMain.Picture = LoadPicture(Game.pathTheme(Oscuro) & "VentanaPrincipal.jpg")
            frmPanelAccount.Picture = LoadPicture(Game.pathTheme(Oscuro) & "frmPanelAccount.jpg")
            frmCrearCuenta.Picture = LoadPicture(Game.pathTheme(Oscuro) & "frmCrearCuenta.jpg")
            Set frmMain.picSkillStar = LoadPicture(Game.pathTheme(Oscuro) & "BotonAsignarSkills.bmp")
    End Select
End Sub
