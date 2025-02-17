Attribute VB_Name = "ModLenguaje"
Option Explicit

Const LOCALE_USER_DEFAULT = &H400
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, _
                                        ByVal LCType As Long, _
                                        ByVal lpLCData As String, _
                                        ByVal cchData As Long) As Long
Public JsonLanguage As Object
Public Language As String

Public Function FileToString(strFileName As String) As String
    Dim IFile As Variant
    IFile = FreeFile
    Open strFileName For Input As #IFile
        FileToString = StrConv(InputB(LOF(IFile), IFile), vbUnicode)
    Close #IFile
End Function

Public Function ObtainOperativeSystemLanguage(ByVal lInfo As Long) As String
    Dim Buffer As String, ret As String
    Buffer = String$(256, 0)
    ret = GetLocaleInfo(LOCALE_USER_DEFAULT, lInfo, Buffer, Len(Buffer))
    If ret > 0 Then
        ObtainOperativeSystemLanguage = Left$(Buffer, ret - 1)
    Else
        ObtainOperativeSystemLanguage = "No se pudo obtener el idioma del sistema."
    End If
End Function

Public Sub SetLanguageApplication()
    Dim LangFile As String
    Language = GetVar(Game.path(INIT) & "Config.ini", "Parameters", "Language")
    If LenB(Language) = 0 Then
        If MsgBox("Deseas iniciar el juego en idioma Español? presiona Sí." + vbCrLf + vbCrLf + _
            "Start with Spanish (Yes), if you want the game in English press No.", vbYesNo + vbInformation, "Argentum Online Libre") = vbYes Then
            Language = "spanish"
        Else
            Language = "english"
        End If
        Call WriteVar(App.path & "\INIT\Config.ini", "Parameters", "Language", Language)
    End If
    LangFile = FileToString(Game.path(Lenguajes) & Language & ".json")
    Set JsonLanguage = JSON.parse(LangFile)
End Sub
