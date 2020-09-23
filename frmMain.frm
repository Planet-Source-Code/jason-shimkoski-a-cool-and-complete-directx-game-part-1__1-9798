VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Mr. Eat Gum Guy"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picLoadSave 
      Height          =   135
      Left            =   120
      ScaleHeight     =   75
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is used for the key presses in the modules
Public Key As Byte

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'This sets key to the keycode so the modules can use the forms keydown event
    Key = KeyCode
End Sub

Private Sub Form_Load()
    'This loads up the main gaming initializations
    Main
End Sub

'This converts the jpgs to bmps
Sub ConvertPic(oldFile As String, newFile As String)
Dim oldPathPic As String
Dim newPathPic As String

    'the jpg and its file path
    oldPathPic = App.Path & "\" & oldFile
    'the to be created bmp and its file path
    newPathPic = App.Path & "\" & newFile

    'loads the jpg into the picture box
    picLoadSave = LoadPicture(oldPathPic)
    'saves the jpg as a bmp
    SavePicture picLoadSave.Picture, newPathPic

End Sub

'This is for easier transportation to the main gaming loop
Sub ConvertAllPics()
    Call ConvertPic("intro.jpg", "intro.bmp")
    Call ConvertPic("bg.jpg", "bg.bmp")
    Call ConvertPic("seescore.jpg", "seescore.bmp")
    Call ConvertPic("exit.jpg", "exit.bmp")
End Sub
