Attribute VB_Name = "modDS"
'The Main Direct Sound Object
Public dsMain As DirectSound

'The introbuffer and its description (Regis)
Public IntroBuffer As DirectSoundBuffer
Public IntroBufferDesc As DSBUFFERDESC

'The woohoobuffer and its description (Homer)
Public WooHooBuffer As DirectSoundBuffer
Public WooHooBufferDesc As DSBUFFERDESC

'The crapbuffer and its description (Krusty)
Public CrapBuffer As DirectSoundBuffer
Public CrapBufferDesc As DSBUFFERDESC

'The byebuffer and the seehell buffer and their descriptions (Apu)
Public ByeBuffer As DirectSoundBuffer
Public ByeBufferDesc As DSBUFFERDESC
Public SeeHellBuffer As DirectSoundBuffer
Public SeeHellBufferDesc As DSBUFFERDESC

'Used in the creation of the sound buffers
Public WavFormat As WAVEFORMATEX

'This creates the sounds from a file
Sub DS_CreateSoundBufFromFile(Buffer As DirectSoundBuffer, FileName As String, BufferDesc As DSBUFFERDESC, wFormat As WAVEFORMATEX)
    If dsMain Is Nothing Then Exit Sub

    Set Buffer = dsMain.CreateSoundBufferFromFile(FileName, BufferDesc, wFormat)
End Sub

'This is for easier transportation to the main initialization
Sub DS_LoadSounds()
    Call DS_CreateSoundBufFromFile(IntroBuffer, App.Path & "\intro.wav", IntroBufferDesc, WavFormat)
    Call DS_CreateSoundBufFromFile(WooHooBuffer, App.Path & "\woohoo.wav", WooHooBufferDesc, WavFormat)
    Call DS_CreateSoundBufFromFile(CrapBuffer, App.Path & "\ahcrap.wav", CrapBufferDesc, WavFormat)
    Call DS_CreateSoundBufFromFile(ByeBuffer, App.Path & "\bye.wav", ByeBufferDesc, WavFormat)
    Call DS_CreateSoundBufFromFile(SeeHellBuffer, App.Path & "\seeinhell.wav", SeeHellBufferDesc, WavFormat)
End Sub

'This plays the sound
Sub DS_Play(Buffer As DirectSoundBuffer, Looping As Boolean)
    Call Buffer.SetCurrentPosition(0)

    'If the sound is to loop
    If Looping = True Then
        'play it with a loop
        Call Buffer.Play(DSBPLAY_LOOPING)
    'If the sound isn't to loop
    Else
        'don't play it with a loop
        Call Buffer.Play(DSBPLAY_DEFAULT)
    End If
End Sub

'This stops the sound
Sub DS_Stop(Buffer As DirectSoundBuffer)
    Buffer.Stop
End Sub

'This stops all of the sounds
Sub DS_StopSounds()
    Call DS_Stop(IntroBuffer)
    Call DS_Stop(WooHooBuffer)
    Call DS_Stop(CrapBuffer)
    Call DS_Stop(ByeBuffer)
    Call DS_Stop(SeeHellBuffer)
End Sub

'This unloads all of the soundbuffers at the programs end
Sub DS_UnloadSounds()
    Set IntroBuffer = Nothing
    Set WooHooBuffer = Nothing
    Set CrapBuffer = Nothing
    Set ByeBuffer = Nothing
    Set SeeHellBuffer = Nothing
End Sub
