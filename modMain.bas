Attribute VB_Name = "modMain"
'This is a great example of what can be done with DirectX and Visual Basic.
'I know that all of the bugs have not been worked out yet, but it is a
'pretty good game to play when bored. It uses Direct Sound, Direct Music,
'and Direct Draw but it doesn't use Direct Input. I didn't feel I needed
'to add in its functionality as the key input was already fast.
'
'If you wish to get the higher resolution pictures for this game then email
'me at basspler@aol.com. I originally intended on having them as downloads
'with the game but their file sizes were way to big.
'
'Thanks,
'Jason Shimkoski (basspler@aol.com)

'This is used for Showing and Hiding the cursor
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'This is to test if the computer is in directx mode or not
Public InDirectXMode As Boolean

'This is the MrGuys Type
Type tMrGuy
    Y As Integer
    X As Integer
    PlayerSpeed As Integer
End Type

'This is the Bad Gums and the Gums Type
Type tGum
    Y As Integer
    X As Integer
    Width As Integer
    Height As Integer
End Type

'Sets the things to their type
Public MrGuy As tMrGuy
Public Gum As tGum
Public BadGum As tGum

'Keeps track of score
Public Score As Integer
'used in character animation
Public CurrentFrame As Integer
'used in the characters speed
Public Speed As Integer

'checks to see if the user lost or not
Public UserLost As Boolean

'used to display score and pause screen
Public txtMsg As String
'used to display ranking
Public txtScoreChart As String

'Main Gaming Loop
Sub Main()

    'convert the jpgs to bmps
    frmMain.ConvertAllPics

    'Initializes directx
    DX_Init

    'sets the screens display
    Call DX_SetDisplay(640, 480, 16)

'Used if player chose to play again
PlayAgain:

    'Loads all of the sounds
    DS_LoadSounds

    'Loads all of the graphics
    DD_LoadGraphics

    'Stops all sounds playing
    Call DS_StopSounds
    Call DM_UnloadStopMidi

    'Plays intro sound
    Call DS_Play(IntroBuffer, False)

    'Goes to main intro screen
    Main_Intro

    'Stops all sounds
    Call DS_StopSounds
    'Plays music
    Call Main_PlayMusic

    'Goes to main game
    Main_LoadPlayGame

    'checks to see if the user wants to play again
    If Main_UserPlayAgain Then GoTo PlayAgain

    'Goes to the exit screen
    Main_ExitScreen

End Sub

'The Intro Screen
Sub Main_Intro()
    'Blits a black color fill to the background
    Call BackBuf.BltColorFill(rIntroSurf, RGB(0, 0, 0))
    'blits the introsurface to the backbuffer
    Call DD_BltFast(0, 0, 640, 480, IntroSurf, rIntroSurf, 0, 0, False)
    'flips the stuff on the backbuffer to the primary buffer
    Call PrimBuf.Flip(Nothing, DDFLIP_WAIT)

    'Used so screen doesn't automatically disappear
    frmMain.Key = 0

    'Sets MrGuys starting speed
    Speed = 1

    Do
    DoEvents
        Select Case frmMain.Key
            'If user presses Enter, then go to next faze of main gaming loop
            Case vbKeyReturn
            Exit Do
            'If user presses Escape, then exit the game
            Case vbKeyEscape
            Main_ExitScreen
            Exit Do
        End Select
    Loop

End Sub

'Sets Bad Gum and Gums initial placement to a random place on the screen
'and sets their height and width
Sub Main_SetUpGum(gGum As tGum)
    Randomize
    gGum.X = Int((590 * Rnd) + 1)
    gGum.Y = Int((430 * Rnd) + 1)

    gGum.Height = 50
    gGum.Width = 50
End Sub

'Sets Mr Guys initial placement to a random place on the screen
Sub Main_SetUpMrGuy(mGuy As tMrGuy)
    Randomize
    mGuy.X = Int((400 * Rnd) + 1)
    mGuy.Y = Int((400 * Rnd) + 1)
End Sub

Sub Main_LoadPlayGame()

    'Used so screen doesn't automatically disappear
    frmMain.Key = 0

    'Sets their placements
    Call Main_SetUpMrGuy(MrGuy)
    Call Main_SetUpGum(Gum)
    Call Main_SetUpGum(BadGum)

    'Sets the score to 0
    Score = 0
    'Sets the players speed to the 1
    MrGuy.PlayerSpeed = Speed
    'Says that the user hasn't lost yet
    UserLost = False

    'This is displayed if the line below is commented out
    txtMsg = "Press any of the Arrow Keys to begin!"

    frmMain.Key = vbKeyRight

    Do
        DoEvents
        'Blit a black to the background
        Call BackBuf.BltColorFill(rBGSurf, RGB(0, 0, 0))
        'blits the backgroundsurface to the backbuffer
        Call DD_BltFast(0, 0, 640, 480, BGSurf, rBGSurf, 0, 0, False)
        'draws MrGuy and the Gums
        Main_DrawMrGuyAndGum
        'Draws the Score on the backbuffer
        Call BackBuf.DrawText(5, 5, txtMsg, False)
        'Flips it all to the primary buffer
        Call PrimBuf.Flip(Nothing, DDFLIP_WAIT)

        'used in MrGuys animation
        CurrentFrame = CurrentFrame + 1
        'calls the sub for MrGuys Movement
        Main_MoveMrGuy
    'Loop until the user has lost
    Loop Until UserLost

End Sub

Sub Main_MoveMrGuy()

    'Set the players speed to the Speed Integer
    MrGuy.PlayerSpeed = Speed

    Select Case frmMain.Key
        'If they press up
        Case vbKeyUp
            'check if he ate the gum
            Main_CheckGumEat
            'Goes to Main_CheckBadGumEat for certain Scores Settings
            Main_CheckScoreForBadGumEat
            'Moves MrGuy Up
            MrGuy.Y = MrGuy.Y - MrGuy.PlayerSpeed
        Case vbKeyDown
            'check if he ate the gum
            Main_CheckGumEat
            'Goes to Main_CheckBadGumEat for certain Scores Settings
            Main_CheckScoreForBadGumEat
            'Moves MrGuy Down
            MrGuy.Y = MrGuy.Y + MrGuy.PlayerSpeed
        Case vbKeyLeft
            'check if he ate the gum
            Main_CheckGumEat
            'Goes to Main_CheckBadGumEat for certain Scores Settings
            Main_CheckScoreForBadGumEat
            'Moves MrGuy Left
            MrGuy.X = MrGuy.X - MrGuy.PlayerSpeed
        Case vbKeyRight
            'check if he ate the gum
            Main_CheckGumEat
            'Goes to Main_CheckBadGumEat for certain Scores Settings
            Main_CheckScoreForBadGumEat
            'Moves MrGuy Right
            MrGuy.X = MrGuy.X + MrGuy.PlayerSpeed
        Case vbKeyEscape
            'Goes to exit screen
            Main_ExitScreen
        Case Else
            'If any other key is pressed, pause the game
            txtMsg = "Press any of the Arrow Keys to unpause the game!"
    End Select

    'This is used for collision detection with the screen
    If MrGuy.Y < 0 Or MrGuy.Y > 429 Then
        UserLost = True
    ElseIf MrGuy.X < 0 Or MrGuy.X > 589 Then
        UserLost = True
    Else
        UserLost = False
    End If

End Sub

'This checks to see if the user ate the bad gum at certain scores
Sub Main_CheckScoreForBadGumEat()
    If Score = 4 Then
        Main_CheckBadGumEat
    ElseIf Score = 9 Then
        Main_CheckBadGumEat
    ElseIf Score = 14 Then
        Main_CheckBadGumEat
    End If
End Sub

'This writes the score
Sub Main_WriteScore()
    txtMsg = "Current Score: " & Score
End Sub

'This draws MrGuy with an animation and the gums
Sub Main_DrawMrGuyAndGum()
Dim i As Integer

    'This is used for MrGuys animation
    If CurrentFrame = 29 Then CurrentFrame = 0
    
    Select Case frmMain.Key
        Case vbKeyUp
            'Writes the score
            Main_WriteScore
            Select Case CurrentFrame
                'if the frame is at 0 through 14 then draw him with his mouth open
                Case 0 To 14
                Call DD_BltFast(50, 150, 200, 100, MrGuySurf, rMrGuySurf, MrGuy.X, MrGuy.Y, True)
                'if the frame is at 15 through 29 then draw him with his mouth closed
                Case 15 To 29
                Call DD_BltFast(50, 100, 150, 100, MrGuySurf, rMrGuySurf, MrGuy.X, MrGuy.Y, True)
            End Select

            'Draws the gum but sets no placement for it
            Main_DrawGumsNoPlacement
        Case vbKeyDown
            'Writes the score
            Main_WriteScore
            Select Case CurrentFrame
                'if the frame is at 0 through 14 then draw him with his mouth open
                Case 0 To 14
                Call DD_BltFast(50, 50, 100, 100, MrGuySurf, rMrGuySurf, MrGuy.X, MrGuy.Y, True)
                'if the frame is at 15 through 29 then draw him with his mouth closed
                Case 15 To 29
                Call DD_BltFast(50, 0, 50, 100, MrGuySurf, rMrGuySurf, MrGuy.X, MrGuy.Y, True)
            End Select

            'Draws the gum but sets no placement for it
            Main_DrawGumsNoPlacement
        Case vbKeyLeft
            'Writes the score
            Main_WriteScore
            Select Case CurrentFrame
                'if the frame is at 0 through 14 then draw him with his mouth open
                Case 0 To 14
                Call DD_BltFast(0, 100, 150, 50, MrGuySurf, rMrGuySurf, MrGuy.X, MrGuy.Y, True)
                'if the frame is at 15 through 29 then draw him with his mouth closed
                Case 15 To 29
                Call DD_BltFast(0, 150, 200, 50, MrGuySurf, rMrGuySurf, MrGuy.X, MrGuy.Y, True)
            End Select

            'Draws the gum but sets no placement for it
            Main_DrawGumsNoPlacement
        Case vbKeyRight
            'Writes the score
            Main_WriteScore
            Select Case CurrentFrame
                'if the frame is at 0 through 14 then draw him with his mouth open
                Case 0 To 14
                Call DD_BltFast(0, 50, 100, 50, MrGuySurf, rMrGuySurf, MrGuy.X, MrGuy.Y, True)
                'if the frame is at 15 through 29 then draw him with his mouth closed
                Case 15 To 29
                Call DD_BltFast(0, 0, 50, 50, MrGuySurf, rMrGuySurf, MrGuy.X, MrGuy.Y, True)
            End Select

            'Draws the gum but sets no placement for it
            Main_DrawGumsNoPlacement
    End Select

End Sub

'Checks to see if the gums has been eaten
Function Main_CheckGumEat() As Boolean

    'This is the collision detection for the gum and MrGuy
    Select Case MrGuy.X
        Case Gum.X - 35 To Gum.X + 40
        Select Case MrGuy.Y
            Case Gum.Y - 35 To Gum.Y + 35
            'Says that the gum was eaten
            Main_CheckGumEat = True
        End Select
    End Select

    'If the gum was eaten then
    If Main_CheckGumEat = True Then
        'play the woohoo sound
        Call DS_Play(WooHooBuffer, False)
        'increase the speed
        Speed = Speed + 1
        'increase the score
        Score = Score + 1
        'draws the gum in a new place
        Main_DrawGum
    Else
        'if the gum wasn't eaten then the speed stays the same
        Speed = Speed
    End If

End Function

'This checks to see if the bad gum was eaten
Function Main_CheckBadGumEat() As Boolean

    'This is the collision detection for the bad gum and MrGuy
    Select Case MrGuy.X
        Case BadGum.X - 35 To BadGum.X + 40
        Select Case MrGuy.Y
            Case BadGum.Y - 35 To BadGum.Y + 35
            'Says that the bad gum was eaten
            Main_CheckBadGumEat = True
        End Select
    End Select

    'If the bad gum was eaten then
    If Main_CheckBadGumEat = True Then
        'play the ah crap sound
        Call DS_Play(CrapBuffer, False)
        'increases the speed
        Speed = Speed + 1
        'decreases the score
        Score = Score - 5
        'draws the bad gum to a new place for future placement
        Main_DrawBadGum
    Else
        'if the bad gum wasn't eaten then the speed stays the same
        Speed = Speed
    End If

End Function

'Draws the gum with new placement
Sub Main_DrawGum()
    'randomly picks a place to put the gum
    Randomize
    Gum.X = Int((590 * Rnd) + 1)
    Gum.Y = Int((430 * Rnd) + 1)
    'draws it
    Call DD_BltFast(0, 0, Gum.Width, Gum.Height, GumSurf, rGumSurf, Gum.X, Gum.Y, True)
End Sub

'Draws the gum without new placement
Sub Main_DrawGumsNoPlacement()
    Call DD_BltFast(0, 0, Gum.Width, Gum.Height, GumSurf, rGumSurf, Gum.X, Gum.Y, True)

    'This will only draw it if the score is 4, 9, or 14
    If Score = 4 Then
        'draws it
        Call DD_BltFast(0, 0, BadGum.Width, BadGum.Height, BadGumSurf, rBadGumSurf, BadGum.X, BadGum.Y, True)
    ElseIf Score = 9 Then
        'draws it
        Call DD_BltFast(0, 0, BadGum.Width, BadGum.Height, BadGumSurf, rBadGumSurf, BadGum.X, BadGum.Y, True)
    ElseIf Score = 14 Then
       'draws it
        Call DD_BltFast(0, 0, BadGum.Width, BadGum.Height, BadGumSurf, rBadGumSurf, BadGum.X, BadGum.Y, True)
    End If

End Sub

'Draws the Bad Gum with new placement
Sub Main_DrawBadGum()

DrawBadGumAgain:
    'randomly picks a place to put the bad gum
    Randomize
    BadGum.X = Int((590 * Rnd) + 1)
    BadGum.Y = Int((430 * Rnd) + 1)

    'This will only draw it if the score is 4, 9, or 14
    If Score = 4 Then
        'draws it
        Call DD_BltFast(0, 0, BadGum.Width, BadGum.Height, BadGumSurf, rBadGumSurf, BadGum.X, BadGum.Y, True)
    ElseIf Score = 9 Then
        'draws it
        Call DD_BltFast(0, 0, BadGum.Width, BadGum.Height, BadGumSurf, rBadGumSurf, BadGum.X, BadGum.Y, True)
    ElseIf Score = 14 Then
        'draws it
        Call DD_BltFast(0, 0, BadGum.Width, BadGum.Height, BadGumSurf, rBadGumSurf, BadGum.X, BadGum.Y, True)
    Else
        'if the score isn't 4, 9, or 14 then just exit
        Exit Sub
    End If

    'checks to see if the bad gum is over the gum
    Select Case Gum.X
        Case BadGum.X - 35 To BadGum.X + 40
        Select Case Gum.Y
            Case BadGum.Y - 35 To BadGum.Y + 35
            'if it is then draw the gum over again
            GoTo DrawBadGumAgain
        End Select
    End Select

End Sub

Sub Main_ExitScreen()

    'this stops all sounds
    DM_UnloadStopMidi
    DS_StopSounds

    'this plays a randomly generated exit sound
    Main_PlayExitSound

    'this draws a black fill to the background
    Call BackBuf.BltColorFill(rIntroSurf, RGB(0, 0, 0))
    'this blits the exit surface to the screen
    Call DD_BltFast(0, 0, 640, 480, ExitSurf, rExitSurf, 0, 0, False)
    'this flips the backbuffers contents to the primary buffer
    Call PrimBuf.Flip(Nothing, DDFLIP_WAIT)

    'this is so the screen doesn't automatically disappear
    frmMain.Key = 0

    'shows this screen until a key is pressed
    Do
    DoEvents
    Loop Until frmMain.Key

    'once the key is pressed then it will go to the main exit routine
    Main_ExitRoutine

End Sub

'This randomly chooses an exit sound
Sub Main_PlayExitSound()
Dim i As Integer

    Randomize
    i = Int((2 * Rnd) + 1)

    If i = 1 Then
        Call DS_Play(ByeBuffer, False)
    Else
        Call DS_Play(SeeHellBuffer, False)
    End If

End Sub

'This randomly chooses music to play
Sub Main_PlayMusic()
Dim i As Integer

    Call DM_CreateLoaderPerformance(frmMain.hWnd)

    Randomize
    i = Int((2 * Rnd) + 1)

    If i = 1 Then
        Call DM_LoadPlayMidi("music.Mid")
    Else
        Call DM_LoadPlayMidi("music2.Mid")
    End If

End Sub

'this is the main exit routine
Sub Main_ExitRoutine()
    'Unloads all of the graphics
    Call DD_UnloadGraphics
    'Stops all currently playing sounds
    Call DS_StopSounds
    'Unloads all of the sounds
    Call DS_UnloadSounds
    'Unloads all of the music
    Call DM_UnloadStopMidi
    'Restores Direct Draws and Direct Sounds Cooperative Levels
    Call DX_RestoreCoopLevel(frmMain.hWnd)
    'Restores the users default display mode
    Call ddMain.RestoreDisplayMode
    'Shows the cursor
    Call ShowCursor(1)
    'ends the program
    End
End Sub

'This checks to see if the user wants to play again
Function Main_UserPlayAgain() As Boolean
'the first line of text to be shown on screen
Dim pScore1 As String
'the second line of text to be shown on screen
Dim pScore2 As String

    'used so the screen wouldn't automatically disappear
    frmMain.Key = 0

    'This gets the text that goes with users score
    Main_GetUserScore

    'This stops all playing sounds
    DM_UnloadStopMidi
    DS_StopSounds

    'This is the first line of text on the screen
    pScore1 = "Total Amount of Gum Eaten: " & Score
    'This is the second line of text on the screen
    pScore2 = "Your Ranking Is: " & txtScoreChart

    'draws a black fill to the background
    Call BackBuf.BltColorFill(rBGSurf, RGB(0, 0, 0))
    'blits the score screen to the backbuffer
    Call DD_BltFast(0, 0, 640, 480, ScoreSurf, rScoreSurf, 0, 0, False)
    'draws the first line of text on the backbuffer
    Call BackBuf.DrawText(250, 200, pScore1, False)
    'draws the second line of text on the backbuffer
    Call BackBuf.DrawText(250, 240, pScore2, False)
    'flips them from the backbuffer to the primary buffer
    Call PrimBuf.Flip(Nothing, DDFLIP_WAIT)

    Do
        DoEvents
        Select Case frmMain.Key
            'if the user presses Enter then start a new game
            Case vbKeyReturn
            Main_UserPlayAgain = True
            Exit Do
            'if the user presses Escape then go to the exit screen
            Case vbKeyEscape
            Main_UserPlayAgain = False
            Exit Do
        End Select
    Loop

End Function

'This basically sets a ranking to the users final score
Sub Main_GetUserScore()
    Select Case Score
        Case -100 To -1
        txtScoreChart = "Pathetic!"
        Case 0 To 4
        txtScoreChart = "Weinee!"
        Case 5 To 9
        txtScoreChart = "Monkey Boy!"
        Case 10 To 13
        txtScoreChart = "Good!"
        Case 14 To 17
        txtScoreChart = "Very Good!"
        Case 18 To 20
        txtScoreChart = "Holy Crap Your Good!"
        Case 21 To 25
        txtScoreChart = "You Rock Dude!"
        Case Else
        txtScoreChart = "You Are the Champion!"
    End Select
End Sub
