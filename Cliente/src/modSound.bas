Attribute VB_Name = "modSound"
Option Explicit

'//General
Public EnableMusic As Boolean

'//Path
Public Const Music_Path As String = "\data\music\"
Public Const Sound_Path As String = "\data\sfx\"
Public Const Cries_Path As String = "\data\sfx\cries\"

'//list cache
Public musicCache() As String
Public soundCache() As String
Public criesCache() As String

'//index
Public SoundIndex As Long
Public MusicIndex As Long

'//Current background music
Public CurMusic As String

'//Volume Control
Public Const MAX_VOLUME As Byte = 7
Public VolumeRange(0 To MAX_VOLUME) As Single
Public CurMusicVolume As Byte
Public CurSoundVolume As Byte

Public Sub InitSound()
    Call ChDrive(App.Path)
    Call ChDir(App.Path)
    
    '//Set Volume Range
    VolumeRange(0) = 0
    VolumeRange(1) = 0.1
    VolumeRange(2) = 0.2
    VolumeRange(3) = 0.3
    VolumeRange(4) = 0.4
    VolumeRange(5) = 0.5
    VolumeRange(6) = 0.6
    VolumeRange(7) = 0.7
    
    '//set default volume
    If GameSetting.Background > MAX_VOLUME Then GameSetting.Background = MAX_VOLUME
    If GameSetting.Background < 0 Then GameSetting.Background = 0
    If GameSetting.SoundEffect > MAX_VOLUME Then GameSetting.SoundEffect = MAX_VOLUME
    If GameSetting.SoundEffect < 0 Then GameSetting.SoundEffect = 0
    CurMusicVolume = GameSetting.Background
    CurSoundVolume = GameSetting.SoundEffect

    '//Check if Bass version is the same with Bass input on source
    If (HiWord(BASS_GetVersion) <> BASSVERSION) Then
        Call MsgBox("An incorrect version of bass.dll was loaded.", vbCritical)
        End
    End If

    '//Check if it properly init bass
    If (BASS_Init(-1, BASS_FREQ, 0, frmMain.hwnd, 0) = 0) Then
        Call MsgBox("Failed to initialise the device.")
        End
    End If

    EnableMusic = True
    '//list all music/sound name
    PopulateLists
End Sub

Public Sub UnloadSound()
    '//clear every sound
    If EnableMusic = False Then Exit Sub
    StopMusic False
    StopMusic True
    Call BASS_Free
End Sub

Public Sub PopulateLists()
Dim strLoad As String, i As Long

    '//Cache music list
    strLoad = Dir(App.Path & Music_Path & "*.*")
    i = 1
    Do While strLoad > vbNullString
        ReDim Preserve musicCache(1 To i) As String
        musicCache(i) = strLoad
        strLoad = Dir
        i = i + 1
    Loop
    
    '//Cache sound list
    strLoad = Dir(App.Path & Sound_Path & "*.*")
    i = 1
    Do While strLoad > vbNullString
        ReDim Preserve soundCache(1 To i) As String
        soundCache(i) = strLoad
        strLoad = Dir
        i = i + 1
    Loop
    
    '//Cache cries list
    strLoad = Dir(App.Path & Cries_Path & "*.*")
    i = 1
    Do While strLoad > vbNullString
        ReDim Preserve criesCache(1 To i) As String
        criesCache(i) = strLoad
        strLoad = Dir
        i = i + 1
    Loop
End Sub

Public Sub StopMusic(ByVal BG As Boolean)
    '//check if sound system is working
    If EnableMusic = False Then Exit Sub
    
    If BG Then
        BASS_ChannelStop MusicIndex
        CurMusic = vbNullString
    Else
        BASS_ChannelStop SoundIndex
    End If
End Sub

Public Sub PlayMusic(ByVal FileName As String, ByVal ResetMusic As Boolean, ByVal CanLoop As Boolean)
    Dim sPlay As Boolean

    '//check if sound system is working
    If EnableMusic = False Then Exit Sub

    If CanLoop Then
        '//check if filename exist
        If Not FileExist(App.Path & Music_Path & FileName) Then Exit Sub
        '//Can loop means, it's a background music
        '//Check if current music is the same and check if it's not changing volume
        If CurMusic = FileName And Not ResetMusic Then Exit Sub

        '//stop current music
        StopMusic True

        sPlay = False
        If ResetMusic Then
            If MusicIndex = 0 Then Exit Sub
            sPlay = True
        Else
            MusicIndex = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & Music_Path & FileName), 0, 0, BASS_SAMPLE_LOOP)
            If MusicIndex = 0 Then Exit Sub
            sPlay = True
        End If

        If sPlay Then
            Call BASS_ChannelSetAttribute(MusicIndex, BASS_ATTRIB_VOL, VolumeRange(CurMusicVolume))
            Call BASS_ChannelPlay(MusicIndex, False)

            '//set current music
            CurMusic = FileName
        End If
    Else
        '//check if filename exist
        '//Sound
        If FileExist(App.Path & Sound_Path & FileName) Then
            SoundIndex = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & Sound_Path & FileName), 0, 0, 0)
            If SoundIndex = 0 Then Exit Sub

            Call BASS_ChannelSetAttribute(SoundIndex, BASS_ATTRIB_VOL, VolumeRange(CurSoundVolume))
            Call BASS_ChannelPlay(SoundIndex, False)
            '//Cries
        ElseIf FileExist(App.Path & Cries_Path & FileName) Then
            SoundIndex = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & Cries_Path & FileName), 0, 0, 0)
            If SoundIndex = 0 Then Exit Sub

            Call BASS_ChannelSetAttribute(SoundIndex, BASS_ATTRIB_VOL, VolumeRange(CurSoundVolume))
            Call BASS_ChannelPlay(SoundIndex, False)
        Else
            Exit Sub
        End If

    End If
End Sub

'//this change the volume if the current playing music/sound
Public Sub ChangeVolume(ByVal Volume As Byte, ByVal BG As Boolean)
    '//check if sound system is working
    If EnableMusic = False Then Exit Sub
    
    If Volume > MAX_VOLUME Then Volume = MAX_VOLUME
    If Volume < 0 Then Volume = 0
    
    If BG Then
        CurMusicVolume = Volume
        
        '//check if curmusic exist
        If Len(Trim$(CurMusic)) <= 0 Then Exit Sub
        '//check if filename exist
        If Not FileExist(App.Path & Music_Path & CurMusic) Then Exit Sub
        PlayMusic CurMusic, True, True
    Else
        CurSoundVolume = Volume
    End If
End Sub

