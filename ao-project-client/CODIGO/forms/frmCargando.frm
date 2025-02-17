VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3225
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCargando.frx":0000
   ScaleHeight     =   186
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox Status 
      Height          =   2070
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3651
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmCargando.frx":1D48A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NoInternetConnection As Boolean
Private VersionNumberMaster As String
Private VersionNumberLocal As String

Private Sub Form_Load()
    'Me.Analizar
End Sub

Function Analizar()
On Error Resume Next
    Dim binaryFileToOpen As String
    Dim isLastVersion As Boolean
    isLastVersion = CheckIfRunningLastVersion
    If NoInternetConnection Then
        MsgBox "No hay conexion a internet, verificar que tengas internet/No Internet connection, please verify"
        Exit Function
    End If
    If Not isLastVersion = True Then
        If MsgBox("Tu version no es la actual, Deseas ejecutar el actualizador?. - Tu version: " & VersionNumberLocal & " Ultima version: " & VersionNumberMaster & " -- Your version is not up to date, open the launcher to update? ", vbYesNo) = vbYes Then
            binaryFileToOpen = GetVar(Game.path(INIT) & "Config.ini", "Launcher", "fileToOpen")
            Call ShellExecute(Me.hwnd, "open", App.path & binaryFileToOpen, "", "", 1)
            End
        End If
    End If
End Function

Private Function CheckIfRunningLastVersion() As Boolean
On Error Resume Next
    Dim responseGithub As String
    Dim JsonObject     As Object
    Set Inet = New clsInet
    responseGithub = Inet.OpenRequest("https://api.github.com/repos/ao-libre/ao-cliente/releases/latest", "GET")
    responseGithub = Inet.Execute
    responseGithub = Inet.GetResponseAsString
    Set JsonObject = JSON.parse(responseGithub)
    VersionNumberMaster = JsonObject.item("tag_name")
    VersionNumberLocal = GetVar(Game.path(INIT) & "Config.ini", "Cliente", "VersionTagRelease")
    If VersionNumberMaster = VersionNumberLocal Then
        CheckIfRunningLastVersion = True
    Else
        CheckIfRunningLastVersion = False
    End If
End Function
