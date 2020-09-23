Attribute VB_Name = "modDD"
'The Main Direct Draw Object
Public ddMain As DirectDraw7

'The Primary Surface, The Backbuffer and their descriptions
Public PrimBuf As DirectDrawSurface7
Public BackBuf As DirectDrawSurface7
Public PrimBufDesc As DDSURFACEDESC2
Public BackBufDesc As DDSURFACEDESC2
'This is used in the creation of the backbuffer
Public BackBufCaps As DDSCAPS2

'These are all of the surfaces (pictures) used in the game
Public IntroSurf As DirectDrawSurface7
Public MrGuySurf As DirectDrawSurface7
Public GumSurf As DirectDrawSurface7
Public BadGumSurf As DirectDrawSurface7
Public BGSurf As DirectDrawSurface7
Public ScoreSurf As DirectDrawSurface7
Public ExitSurf As DirectDrawSurface7

'These are all of the surfaces descriptions
Public IntroSurfDesc As DDSURFACEDESC2
Public MrGuySurfDesc As DDSURFACEDESC2
Public GumSurfDesc As DDSURFACEDESC2
Public BadGumSurfDesc As DDSURFACEDESC2
Public BGSurfDesc As DDSURFACEDESC2
Public ScoreSurfDesc As DDSURFACEDESC2
Public ExitSurfDesc As DDSURFACEDESC2

'These are all of the surfaces RECT containers
Public rIntroSurf As RECT
Public rMrGuySurf As RECT
Public rGumSurf As RECT
Public rBadGumSurf As RECT
Public rBGSurf As RECT
Public rScoreSurf As RECT
Public rExitSurf As RECT

'This is for transparency
Public ColorKey As DDCOLORKEY

'This creates the primary and backbuffer
Sub DD_CreatePrimBackBuf()

    Set PrimBuf = Nothing
    Set BackBuf = Nothing

    'This is some stuff in the making of the primary surface
    PrimBufDesc.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    PrimBufDesc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    'This says that there will only be 1 backbuffer
    'put 2 if you want triple buffering!
    PrimBufDesc.lBackBufferCount = 1
    'This sets the primary surface to its description
    Set PrimBuf = ddMain.CreateSurface(PrimBufDesc)

    'This attaches the caps to the backbuffer
    BackBufCaps.lCaps = DDSCAPS_BACKBUFFER
    'This sets the backbuf surface as the backbuffer
    Set BackBuf = PrimBuf.GetAttachedSurface(BackBufCaps)
    Call BackBuf.GetSurfaceDesc(PrimBufDesc)

    'This sets the fonts color to white
    Call BackBuf.SetForeColor(RGB(255, 255, 255))
    'This sets the font to have a transparent background
    Call BackBuf.SetFontTransparency(True)

End Sub

'This creates the surfaces from files
Sub DD_CreateSurfFromFile(FileName As String, Surface As DirectDrawSurface7, SurfDesc As DDSURFACEDESC2, Width As Long, Height As Long)
On Error GoTo errNoFile

    SurfDesc.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    'This sets the surface as an offscreenplain
    SurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    'This is defining its height
    SurfDesc.lHeight = Height
    'This is defining its width
    SurfDesc.lWidth = Width

    'This sets the surface to the created file
    Set Surface = ddMain.CreateSurfaceFromFile(FileName, SurfDesc)

    'This is used for transparency
    Dim ColorKey As DDCOLORKEY
    ColorKey.high = 0
    ColorKey.low = 0
    Call Surface.SetColorKey(DDCKEY_SRCBLT, ColorKey)

errNoFile:
    Debug.Print "File Not Found: " & FileName
End Sub

'This loads up all of the graphics. This is mainly for easier
'transportation to the main initialization
Sub DD_LoadGraphics()

    DD_CreatePrimBackBuf

    Set MrGuySurf = Nothing
    Set GumSurf = Nothing

    Call DD_CreateSurfFromFile(App.Path & "\intro.bmp", IntroSurf, IntroSurfDesc, 640, 480)
    Call DD_CreateSurfFromFile(App.Path & "\mrguy.bmp", MrGuySurf, MrGuySurfDesc, 200, 100)
    Call DD_CreateSurfFromFile(App.Path & "\gum.bmp", GumSurf, GumSurfDesc, 50, 50)
    Call DD_CreateSurfFromFile(App.Path & "\badgum.bmp", BadGumSurf, BadGumSurfDesc, 50, 50)
    Call DD_CreateSurfFromFile(App.Path & "\bg.bmp", BGSurf, BGSurfDesc, 640, 480)
    Call DD_CreateSurfFromFile(App.Path & "\seescore.bmp", ScoreSurf, ScoreSurfDesc, 640, 480)
    Call DD_CreateSurfFromFile(App.Path & "\exit.bmp", ExitSurf, ExitSurfDesc, 640, 480)

End Sub

'This is for blitting the surfaces to the backbuffer
Sub DD_BltFast(rTop As Integer, rLeft As Integer, Width As Integer, Height As Integer, Surface As DirectDrawSurface7, srcRect As RECT, X As Integer, Y As Integer, Transparency As Boolean)

    DoUntilReady

    'This sets the RECT. This will be used for selecting different images
    'and that sort of thing
    srcRect.Top = rTop
    srcRect.Left = rLeft
    srcRect.Right = Width
    srcRect.Bottom = Height

    'if the image won't be transparent
    If Transparency = False Then
        'don't blit it with a transparency
        Call BackBuf.BltFast(X, Y, Surface, srcRect, DDBLTFAST_WAIT)
    'if the image is to be transparent
    Else
        'do blit it with a transparency
        Call BackBuf.BltFast(X, Y, Surface, srcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If

End Sub

'This is used for testing the Cooperative level and to see if the program
'is ready to proceed with operation
Sub DoUntilReady()
Dim bRest As Boolean

    bRest = False
    Do Until DirectXActive
        DoEvents
        bRest = True
    Loop

    DoEvents
    If bRest Then
        bRest = False
        ddMain.RestoreAllSurfaces
        DD_LoadGraphics
    End If
End Sub

'This is used in the sub from right above to test the cooperative level
Function DirectXActive() As Boolean
Dim TestCoopLevel As Long

    'This tests the cooperative level
    TestCoopLevel = ddMain.TestCooperativeLevel

    'If everything is okay then
    If (TestCoopLevel = DD_OK) Then
    'we're in directx mode
    DirectXActive = True
    'If everything isn't okay
    Else
    'we're not in directx mode
    DirectXActive = False
    End If

End Function

'this unloads all of the graphics at the programs shutdown
Sub DD_UnloadGraphics()
    Set IntroSurf = Nothing
    Set MrGuySurf = Nothing
    Set GumSurf = Nothing
    Set BadGumSurf = Nothing
    Set BGSurf = Nothing
    Set ScoreSurf = Nothing
    Set ExitSurf = Nothing
    
    Set PrimBuf = Nothing
    Set BackBuf = Nothing
End Sub
