Attribute VB_Name = "mDx8_Engine"
#If False Then
    Dim hwnd, X, Y As Variant
#End If

Option Explicit

Public DirectX As New DirectX8
Public DirectD3D8 As D3DX8
Public DirectD3D As Direct3D8
Public DirectDevice As Direct3DDevice8
Public DispMode  As D3DDISPLAYMODE
Public D3DWindow As D3DPRESENT_PARAMETERS
Public SurfaceDB As New clsTextureManager
Public SpriteBatch As New clsBatch
Private Viewport As D3DVIEWPORT8
Private Projection As D3DMATRIX
Private View As D3DMATRIX
Public Engine_BaseSpeed As Single
Public TileBufferSize As Integer
Public ScreenWidth As Long
Public ScreenHeight As Long
Public MainScreenRect As RECT

Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    Color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type

Private EndTime As Long

Public Sub Engine_DirectX8_Init()
    On Error GoTo ErrorHandler:
    Set DirectX = New DirectX8
    Set DirectD3D = DirectX.Direct3DCreate
    Set DirectD3D8 = New D3DX8
    If ClientSetup.OverrideVertexProcess > 0 Then
        Select Case ClientSetup.OverrideVertexProcess
            Case 1:
               If Not Engine_Init_DirectDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then _
               Call MsgBox(JsonLanguage.item("ERROR_DIRECTX_INIT").item("TEXTO"))
            
            Case 2:
               If Not Engine_Init_DirectDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then _
               Call MsgBox(JsonLanguage.item("ERROR_DIRECTX_INIT").item("TEXTO"))
            
            Case 3:
               If Not Engine_Init_DirectDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then _
               Call MsgBox(JsonLanguage.item("ERROR_DIRECTX_INIT").item("TEXTO"))
        End Select
    Else
        If Not Engine_Init_DirectDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
            If Not Engine_Init_DirectDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
                If Not Engine_Init_DirectDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
                    Call MsgBox(JsonLanguage.item("ERROR_DIRECTX_INIT").item("TEXTO"))
                    End
                End If
            End If
        End If
    End If
    Call D3DXMatrixOrthoOffCenterLH(Projection, 0, ScreenWidth, ScreenHeight, 0, -1#, 1#)
    Call D3DXMatrixIdentity(View)
    Call DirectDevice.SetTransform(D3DTS_PROJECTION, Projection)
    Call DirectDevice.SetTransform(D3DTS_VIEW, View)
    Call Engine_Init_RenderStates
    Set SurfaceDB = New clsTextureManager
    Set SpriteBatch = New clsBatch
    Call SpriteBatch.Initialise(500)
    Call Engine_DirectX8_Aditional_Init
    EndTime = GetTickCount()
    Exit Sub
ErrorHandler:
    Call LogError(Err.number, Err.Description, "mDx8_Engine.Engine_DirectX8")
    Call CloseClient
End Sub

Private Function Engine_Init_DirectDevice(D3DCREATEFLAGS As CONST_D3DCREATEFLAGS) As Boolean
On Error GoTo ErrorHandler:
    ScreenWidth = frmMain.MainViewPic.ScaleWidth
    ScreenHeight = frmMain.MainViewPic.ScaleHeight
    Call DirectD3D.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, DispMode)
    With D3DWindow
        .Windowed = True
        .SwapEffect = D3DSWAPEFFECT_DISCARD
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = ScreenWidth
        .BackBufferHeight = ScreenHeight
        .hDeviceWindow = frmMain.MainViewPic.hwnd
    End With
    If Not DirectDevice Is Nothing Then
        Set DirectDevice = Nothing
    End If
    Set DirectDevice = DirectD3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, D3DWindow.hDeviceWindow, D3DCREATEFLAGS, D3DWindow)
    Select Case D3DCREATEFLAGS
        Case D3DCREATE_MIXED_VERTEXPROCESSING
            Debug.Print "Modo de Renderizado: MIXED"
        
        Case D3DCREATE_HARDWARE_VERTEXPROCESSING
            Debug.Print "Modo de Renderizado: HARDWARE"
            
        Case D3DCREATE_SOFTWARE_VERTEXPROCESSING
            Debug.Print "Modo de Renderizado: SOFTWARE"
    End Select
    Engine_Init_DirectDevice = True
    Exit Function
ErrorHandler:
    Set DirectDevice = Nothing
    Engine_Init_DirectDevice = False
End Function

Private Sub Engine_Init_RenderStates()
    With DirectDevice
        Call .SetVertexShader(D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
        Call .SetRenderState(D3DRS_LIGHTING, False)
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        Call .SetRenderState(D3DRS_ALPHABLENDENABLE, True)
        Call .SetRenderState(D3DRS_FILLMODE, D3DFILL_SOLID)
        Call .SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)
        Call .SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
    End With
End Sub

Public Sub Engine_DirectX8_End()
On Error Resume Next
    Dim i As Byte
    Call DeInit_LightEngine
    Call DeInit_Auras
    Call Particle_Group_Remove_All
    Call DirectDevice.SetTexture(0, Nothing)
    Call CleanDrawBuffer
    Erase MapData()
    Erase charlist()
    Set DirectD3D8 = Nothing
    Set DirectD3D = Nothing
    Set DirectX = Nothing
    Set DirectDevice = Nothing
    Set SpriteBatch = Nothing
End Sub

Public Sub Engine_DirectX8_Aditional_Init()
    FPS = 101
    FramesPerSecCounter = 101
    Engine_BaseSpeed = 0.018
    With MainScreenRect
        .Bottom = ScreenHeight
        .Right = ScreenWidth
    End With
    Call mDx8_Text.Engine_Init_FontTextures
    If Not prgRun Then
        ColorTecho = 250
        colorRender = 240
        Call Engine_Long_To_RGB_List(Normal_RGBList(), -1)
        Call Engine_Long_To_RGB_List(Color_Shadow(), D3DColorARGB(50, 0, 0, 0))
        Call Engine_Long_To_RGB_List(Color_Arbol(), D3DColorARGB(190, 100, 100, 100))
        Color_Paralisis = D3DColorARGB(180, 230, 230, 250)
        Color_Invisibilidad = D3DColorARGB(180, 236, 136, 66)
        Color_Montura = D3DColorARGB(180, 15, 230, 40)
        Call mDx8_Text.Engine_Init_FontSettings
        Call mDx8_Auras.Load_Auras
        Call mDx8_Clima.Init_MeteoEngine
        Call mDx8_Dibujado.Damage_Initialize
        Call PrepareDrawBuffer
    End If
End Sub

Public Sub Engine_BeginScene(Optional ByVal Color As Long = 0)
    Call DirectDevice.BeginScene
    Call DirectDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, Color, 1#, 0)
    Call SpriteBatch.Begin
End Sub

Public Sub Engine_EndScene(ByRef DestRect As RECT, Optional ByVal hWndDest As Long = 0)
On Error GoTo ErrorHandler:
    Call SpriteBatch.Flush
    Call DirectDevice.EndScene
    If hWndDest = 0 Then
        Call DirectDevice.Present(DestRect, ByVal 0&, ByVal 0&, ByVal 0&)
    Else
        Call DirectDevice.Present(DestRect, ByVal 0, hWndDest, ByVal 0)
    End If
    Exit Sub
ErrorHandler:
    If DirectDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        Call mDx8_Engine.Engine_DirectX8_Init
        Call LoadGraphics
    End If
End Sub

Public Sub Engine_Update_FPS()
    If FPSLastCheck + 1000 < GetTickCount() Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 1
        FPSLastCheck = GetTickCount()
    Else
        FramesPerSecCounter = FramesPerSecCounter + 1
    End If
End Sub

Public Function Engine_ElapsedTime() As Long
    Dim Start_Time As Long
    Start_Time = GetTickCount()
    Engine_ElapsedTime = Start_Time - EndTime
    EndTime = Start_Time
End Function
