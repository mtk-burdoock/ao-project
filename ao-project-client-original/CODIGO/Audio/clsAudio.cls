VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const NumSoundBuffers As Byte = 30
Private Const MAX_DISTANCE_TO_SOURCE As Integer = 150
Private Const DELTA_FQ As Integer = 75
Private Const MIN_MUSIC_VOLUME As Integer = -1500
Private Const INCREASE_MUSIC_VOLUME As Integer = 20
Private Const MIN_SOUND_VOLUME As Integer = -2000
Private Const INCREASE_SOUND_VOLUME As Integer = 20

Private Type SoundBuffer
    FileName As String
    looping As Boolean
    X As Byte
    Y As Byte
    normalFq As Long
    Buffer As DirectSoundSecondaryBuffer8
End Type

Public Enum LoopStyle
    Default = 0
    Disabled = 1
    Enabled = 2
End Enum

Public Enum MusicTypes
    Mp3
    Midi
End Enum

Dim inited As Boolean
Dim DirectSound As DirectSound8
Dim DSBuffers(1 To NumSoundBuffers) As SoundBuffer
Dim SoundPath As String
Dim SndVolume As Long
Dim Performance As DirectMusicPerformance8
Dim Segment As DirectMusicSegment8
Dim Loader As DirectMusicLoader8
Dim SegState As DirectMusicSegmentState8
Dim Music As Boolean
Dim MusicPath As String
Dim MusicMp3Path As String
Dim CurrentMusicFile As String
Dim CurrentMusicFileType As MusicTypes
Dim mActivated As Boolean
Dim sActivated As Boolean
Dim lastPosX As Integer
Dim lastPosY As Integer
Dim sEffectsActivated As Boolean
Dim IMC   As IMediaControl
Dim IBA   As IBasicAudio
Dim IME   As IMediaEvent
Dim IMPos As IMediaPosition

Private Sub Class_Terminate()
On Local Error Resume Next
    Dim LoopC As Long
    For LoopC = 1 To NumSoundBuffers
        If LenB(DSBuffers(LoopC).FileName) <> 0 Then
            Call DSBuffers(LoopC).Buffer.Stop
        End If
        Set DSBuffers(LoopC).Buffer = Nothing
    Next LoopC
    Call StopWave
    If Not Segment Is Nothing Then Call Performance.StopEx(Segment, 0, DMUS_SEGF_DEFAULT)
    Call Performance.CloseDown
    Set DirectSound = Nothing
    Set Segment = Nothing
    Set Performance = Nothing
    Set Loader = Nothing
    inited = False
End Sub

Public Function Initialize(ByRef dX As DirectX8, ByVal hwnd As Long, ByVal sound_path As String, ByVal music_path As String, ByVal music_mp3_path As String) As Boolean
On Error GoTo ErrorHandler
    If inited Then Exit Function
    If dX Is Nothing Then GoTo ErrorHandler
    Set DirectSound = dX.DirectSoundCreate("")
    Call DirectSound.SetCooperativeLevel(hwnd, DSSCL_PRIORITY)
    SoundPath = sound_path
    SndVolume = 0
    Music = InitializedirectMusic(dX, hwnd)
    MusicPath = music_path
    MusicMp3Path = music_mp3_path
    mActivated = True
    sActivated = True
    CurrentMusicFile = vbNullString
    MusicTempo = 0
    Call Performance.SetMasterVolume(0)
    inited = True
    Initialize = True
Exit Function
ErrorHandler:
    Call LogError(Err.number, Err.Description, "clsAudio.Initialize")
End Function

Private Function InitializedirectMusic(ByRef dX As DirectX8, ByVal hwnd As Long) As Boolean
On Error GoTo ErrorHandler
    Dim mus_Params As DMUS_AUDIOPARAMS
    Set Loader = DirectX.DirectMusicLoaderCreate()
    Set Performance = dX.DirectMusicPerformanceCreate()
    Call Performance.InitAudio(hwnd, DMUS_AUDIOF_ALL, mus_Params, Nothing, DMUS_APATH_DYNAMIC_STEREO, 128)
    Call Performance.SetMasterAutoDownload(True)
    Performance.SetMasterTempo 1
    Performance.SetMasterVolume 1
    InitializedirectMusic = True
Exit Function
ErrorHandler:
    Call LogError(Err.number, Err.Description, "clsAudio.InitializedirectMusic")
End Function

Public Function PlayWave(ByVal FileName As String, Optional ByVal srcX As Byte = 0, Optional ByVal srcY As Byte = 0, Optional ByVal LoopSound As LoopStyle = LoopStyle.Default) As Long
On Error GoTo ErrorHandler
    Dim bufferIndex As Long
    If Not SoundActivated Then Exit Function
    bufferIndex = LoadWave(FileName, LoopSound)
    If bufferIndex = 0 Then Exit Function
    With DSBuffers(bufferIndex)
        If .looping Then
            .Buffer.Play DSBPLAY_LOOPING
        Else
            .Buffer.Play DSBPLAY_DEFAULT
        End If
        Call .Buffer.SetVolume(SndVolume)
        .X = srcX
        .Y = srcY
        .normalFq = .Buffer.GetFrequency()
    End With
    PlayWave = bufferIndex
    If sEffectsActivated Then
        If srcX <> 0 And srcY <> 0 Then
            Call Update3DSound(bufferIndex, 0, 0)
        End If
    End If
Exit Function
ErrorHandler:
    Call LogError(Err.number, Err.Description, "clsAudio.PlayWave")
End Function

Private Sub ResetSoundValues()
    Dim i As Long
    For i = 1 To NumSoundBuffers
        If Not DSBuffers(i).Buffer Is Nothing Then
            With DSBuffers(i)
                Call .Buffer.SetFrequency(.normalFq)
                Call .Buffer.SetVolume(SndVolume)
                Call .Buffer.SetPan(0)
            End With
        End If
    Next i
End Sub

Private Sub Update3DSound(ByVal bufferIndex As Long, ByVal deltaX As Integer, ByVal deltaY As Integer)
    Dim linearDistanceOld As Single
    Dim linearDistanceNew As Single
    Dim deltaDistance As Single
    Dim distanceXOld As Long
    Dim distanceYOld As Long
    Dim distanceXNew As Long
    Dim distanceYNew As Long
    Dim perceivedFq As Long
    With DSBuffers(bufferIndex)
        distanceXOld = .X - lastPosX
        distanceYOld = .Y - lastPosY
        distanceXNew = distanceXOld + deltaX
        distanceYNew = distanceYOld + deltaY
        linearDistanceOld = Sqr(distanceXOld * distanceXOld + distanceYOld * distanceYOld)
        linearDistanceNew = Sqr(distanceXNew * distanceXNew + distanceYNew * distanceYNew)
        deltaDistance = linearDistanceNew - linearDistanceOld
        perceivedFq = .normalFq * (1 - deltaDistance / DELTA_FQ)
        If perceivedFq < DSBFREQUENCY_MIN Then perceivedFq = DSBFREQUENCY_MIN
        Call .Buffer.SetFrequency(perceivedFq)
        If linearDistanceNew > MAX_DISTANCE_TO_SOURCE Then
            Call .Buffer.SetVolume(SndVolume + (DSBVOLUME_MIN - SndVolume))
        Else
            Call .Buffer.SetVolume(SndVolume + (linearDistanceNew / MAX_DISTANCE_TO_SOURCE) * (DSBVOLUME_MIN - SndVolume))
        End If
        If linearDistanceNew = 0 Then linearDistanceNew = 1
        Call .Buffer.SetPan((distanceXNew / linearDistanceNew) * 2000)
    End With
End Sub

Public Sub MoveListener(ByVal X As Integer, ByVal Y As Integer)
    Dim i As Long
    Dim deltaX As Integer
    Dim deltaY As Integer
    If sEffectsActivated = False Then Exit Sub
    deltaX = X - lastPosX
    deltaY = Y - lastPosY
    For i = 1 To NumSoundBuffers
        If Not DSBuffers(i).Buffer Is Nothing Then
            If DSBuffers(i).Buffer.GetStatus And DSBSTATUS_PLAYING Then
                If DSBuffers(i).X <> 0 And DSBuffers(i).Y <> 0 Then
                    Call Update3DSound(i, deltaX, deltaY)
                End If
            End If
        End If
    Next i
    lastPosX = X
    lastPosY = Y
End Sub

Private Function LoadWave(ByVal FileName As String, ByVal looping As LoopStyle) As Long
On Error GoTo ErrorHandler
    Dim bufferDesc As DSBUFFERDESC
    Dim waveFormat As WAVEFORMATEX
    Dim i As Long
    FileName = UCase$(FileName)
    If Not FileExist(SoundPath & FileName, vbArchive) Then _
        Exit Function
    For i = 1 To NumSoundBuffers
        If DSBuffers(i).FileName = FileName Then
            If DSBuffers(i).Buffer.GetStatus = 0 Or DSBuffers(i).Buffer.GetStatus() = DSBSTATUS_BUFFERLOST Then
                LoadWave = i
                If looping <> LoopStyle.Default Then DSBuffers(i).looping = (looping = LoopStyle.Enabled)
                Call DSBuffers(i).Buffer.SetFrequency(DSBuffers(i).normalFq)
                Exit Function
            End If
        End If
    Next i
    For i = 1 To NumSoundBuffers
        If DSBuffers(i).Buffer Is Nothing Then
            Exit For
        End If
    Next i
    If i > NumSoundBuffers Then
        For i = 1 To NumSoundBuffers
            If DSBuffers(i).Buffer.GetStatus() = DSBSTATUS_BUFFERLOST Or DSBuffers(i).Buffer.GetStatus() = 0 Then
                Exit For
            End If
        Next i
    End If
    If i > NumSoundBuffers Then
        For i = 1 To NumSoundBuffers
            If Not (DSBuffers(i).Buffer.GetStatus() And DSBSTATUS_LOOPING) Then
                Call DSBuffers(i).Buffer.Stop
                Exit For
            End If
        Next i
    End If
    If i > NumSoundBuffers Then
        If Not looping Then Exit Function
        i = 1
    End If
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2
    waveFormat.lSamplesPerSec = 22050
    waveFormat.nBitsPerSample = 16
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    With DSBuffers(i)
        Set .Buffer = Nothing
        .FileName = FileName
        .looping = (looping = LoopStyle.Enabled)
        Set .Buffer = DirectSound.CreateSoundBufferFromFile(SoundPath & FileName, bufferDesc)
    End With
    LoadWave = i
Exit Function
ErrorHandler:
    Call LogError(Err.number, Err.Description, "clsAudio.LoadWave")
End Function

Public Sub StopWave(Optional ByVal bufferIndex As Long = 0)
    If bufferIndex > 0 And bufferIndex <= NumSoundBuffers Then
        If DSBuffers(bufferIndex).Buffer.GetStatus() And DSBSTATUS_PLAYING Then
            Call DSBuffers(bufferIndex).Buffer.Stop
        End If
    ElseIf bufferIndex = 0 Then
        Dim i As Long
        For i = 1 To NumSoundBuffers
            If Not DSBuffers(i).Buffer Is Nothing Then
                If DSBuffers(i).Buffer.GetStatus() <> 0 And DSBuffers(i).Buffer.GetStatus() <> DSBSTATUS_BUFFERLOST Then
                    Call DSBuffers(i).Buffer.Stop
                End If
            End If
        Next i
    End If
End Sub

Public Property Get PlayingSound() As Boolean
    Dim i As Long
    For i = 1 To NumSoundBuffers
        If Not DSBuffers(i).Buffer Is Nothing Then
            If DSBuffers(i).Buffer.GetStatus() And DSBSTATUS_PLAYING Then
                PlayingSound = True
                Exit Property
            End If
        End If
    Next i
End Property

Public Property Let SoundVolume(ByVal Volume As Byte)
    Dim i As Long
    If Volume < 0 Or Volume > 100 Then Exit Property
    SndVolume = Volume * INCREASE_SOUND_VOLUME + MIN_SOUND_VOLUME
    For i = 1 To NumSoundBuffers
        If Not DSBuffers(i).Buffer Is Nothing Then
            If DSBuffers(i).Buffer.GetStatus And DSBSTATUS_PLAYING Then
                Call DSBuffers(i).Buffer.SetVolume(SndVolume)
            End If
        End If
    Next i
End Property

Public Property Get SoundVolume() As Byte
    SoundVolume = CByte((SndVolume - MIN_SOUND_VOLUME) / INCREASE_SOUND_VOLUME)
End Property

Public Function PlayMIDI(Optional ByVal File As String = vbNullString, Optional ByVal Loops As Long = -1) As Boolean
On Error GoTo ErrorHandler
    If LenB(File) > 0 Then CurrentMusicFile = File
    CurrentMusicFileType = MusicTypes.Midi
    If Not MusicActivated Then Exit Function
    If Not Music Then Exit Function
    Call StopMidi
    Call StopMp3
    If LenB(File) > 0 Then
        If Not LoadMidi(File) Then Exit Function
    Else
        If Segment Is Nothing Then Exit Function
    End If
    Call Segment.SetRepeats(Loops)
    Set SegState = Performance.PlaySegmentEx(Segment, DMUS_SEGF_DEFAULT, 0)
    PlayMIDI = True
Exit Function
ErrorHandler:
    Call LogError(Err.number, Err.Description, "clsAudio.PlayMIDI")
End Function

Private Function LoadMidi(ByVal File As String) As Boolean
On Error GoTo ErrorHandler
    If Not FileExist(MusicPath & File, vbArchive) Then Exit Function
    Call StopMidi
    Set Segment = Nothing
    Set Segment = Loader.LoadSegment(MusicPath & File)
    If Segment Is Nothing Then
        Exit Function
    End If
    Call Segment.SetStandardMidiFile
    LoadMidi = True
Exit Function
ErrorHandler:
    Call LogError(Err.number, Err.Description, "clsAudio.LoadMidi")
End Function

Public Sub StopMidi()
On Error GoTo ErrorHandler
    If PlayingMusic Then
        Call Performance.StopEx(Segment, 0, DMUS_SEGF_DEFAULT)
    End If
    Exit Sub
ErrorHandler:
    Call LogError(Err.number, Err.Description, "clsAudio.StopMidi")
End Sub

Public Property Let MusicTempo(ByVal Tempo As Single)
    If Not Music Then Exit Property
    If Tempo < 0.25 Or Tempo > 2# Then Exit Property
    Call Performance.SetMasterTempo(Tempo)
End Property

Public Property Let MusicVolume(ByVal Volume As Long)
    If Not Music Then Exit Property
    If Volume < 0 Or Volume > 100 Then Exit Property
    Call Performance.SetMasterVolume(Volume * INCREASE_MUSIC_VOLUME + MIN_MUSIC_VOLUME)
    If Not IBA Is Nothing Then
        IBA.Volume = Volume * INCREASE_SOUND_VOLUME + MIN_SOUND_VOLUME
    End If
End Property

Public Property Get MusicTempo() As Single
    If Not Music Then Exit Property
    MusicTempo = Performance.GetMasterTempo()
End Property

Public Property Get MusicVolume() As Long
    If Not Music Then Exit Property
    MusicVolume = (Performance.GetMasterVolume() - MIN_MUSIC_VOLUME) / INCREASE_MUSIC_VOLUME
End Property

Public Property Get PlayingMusic() As Boolean
    If Not Music Or Segment Is Nothing Then Exit Property
    PlayingMusic = Performance.IsPlaying(Segment, SegState)
End Property

Public Property Get MusicLenght() As Long
    If Not Music Then Exit Property
    MusicLenght = Segment.GetLength()
End Property

Public Property Get MusicActivated() As Boolean
    MusicActivated = mActivated
End Property

Public Property Let MusicActivated(ByVal nValue As Boolean)
    If mActivated = nValue Then Exit Property
    mActivated = nValue
    If nValue Then
        Select Case CurrentMusicFileType
            Case MusicTypes.Midi
                mActivated = PlayMIDI(CurrentMusicFile)
                
            Case MusicTypes.Mp3
                mActivated = PlayMp3(CurrentMusicFile)
        End Select
    Else

        Select Case CurrentMusicFileType
            Case MusicTypes.Midi
                Call StopMidi
                
            Case MusicTypes.Mp3
                Call StopMp3
        End Select
    End If
End Property

Public Property Get SoundActivated() As Boolean
    SoundActivated = sActivated
End Property

Public Property Let SoundActivated(ByVal nValue As Boolean)
    If sActivated = nValue Then Exit Property
    sActivated = nValue
    If Not nValue Then
        Call StopWave
    End If
End Property

Public Property Let SoundEffectsActivated(ByVal nValue As Boolean)
    sEffectsActivated = nValue
    If Not nValue Then
        Call ResetSoundValues
    End If
End Property

Public Property Get SoundEffectsActivated() As Boolean
    SoundEffectsActivated = sEffectsActivated
End Property

Public Function PlayMp3(Optional ByVal File As String = vbNullString) As Boolean
    On Error GoTo ErrorHandler
    If LenB(File) > 0 Then CurrentMusicFile = File
    CurrentMusicFileType = MusicTypes.Mp3
    If Not MusicActivated Then Exit Function
    If Not Music Then Exit Function
    Call StopMp3
    Call StopMidi
    If LenB(File) > 0 Then
        If Not LoadMp3(File, MusicVolume * INCREASE_SOUND_VOLUME + MIN_SOUND_VOLUME) Then Exit Function
    Else
        If IMC Is Nothing Then Exit Function
    End If
    Call IMC.Run
    PlayMp3 = True
    Exit Function
ErrorHandler:
    PlayMp3 = False
    Call LogError(Err.number, Err.Description, "clsAudio.PlayMp3")
End Function

Public Function LoadMp3(ByVal FilePath As String, Optional ByVal Volume As Long = 0, Optional ByVal Balance As Long = 0) As Boolean
    On Error GoTo ErrorHandler
    If Not FileExist(MusicMp3Path & FilePath, vbArchive) Then Exit Function
    If Not MP3Empty() = True Then GoTo ErrorHandl
    Set IMC = New FilgraphManager
    Call IMC.RenderFile(MusicMp3Path & FilePath)
    Set IBA = IMC
    If Volume < -4000 Then Volume = -4000
    IBA.Volume = Volume
    If Balance > 5000 Then Balance = 5000
    If Balance < -5000 Then Balance = -5000
    IBA.Balance = Balance
    Set IME = IMC
    Set IMPos = IMC
    If ObjPtr(IMPos) Then IMPos.Rate = 1#
    IMPos.CurrentPosition = 0
    LoadMp3 = True
    Exit Function
ErrorHandler:
    LoadMp3 = False
    Call LogError(Err.number, Err.Description, "clsAudio.LoadMp3")
End Function

Public Function StopMp3() As Boolean
    On Error GoTo ErrorHandler
    If IMC Is Nothing Then Exit Function
    IMC.Stop
    IMPos.CurrentPosition = 0
    StopMp3 = True
    Exit Function
ErrorHandler:
    StopMp3 = False
    Call LogError(Err.number, Err.Description, "clsAudio.StopMp3")
End Function

Public Function MP3Empty() As Boolean
    On Error GoTo ErrorHandler
    If ObjPtr(IMC) > 0 Then
        Call StopMp3
    End If
    Set IBA = Nothing
    Set IME = Nothing
    Set IMPos = Nothing
    Set IMC = Nothing
    MP3Empty = True
    Exit Function
ErrorHandler:
    MP3Empty = False
    Call LogError(Err.number, Err.Description, "clsAudio.MP3Empty")
    End
End Function

Public Sub PlayBackgroundMusic(ByVal FileName As String, ByVal FileType As Byte, Optional ByVal Loops As Long = -1)
    If CurrentMusicFile = "6.mp3" And FileName = "2" Then Exit Sub
    Select Case FileType
        Case MusicTypes.Midi
            FileName = FileName & ".mid"
            If FileName = CurrentMusicFile Then Exit Sub
            Call PlayMIDI(FileName, Loops)
            
        Case MusicTypes.Mp3
            FileName = FileName & ".mp3"
            If FileName = CurrentMusicFile Then Exit Sub
            Call PlayMp3(FileName)
    End Select
End Sub
