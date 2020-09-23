VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHangman 
   BackColor       =   &H8000000E&
   Caption         =   "Hangman"
   ClientHeight    =   5010
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   7080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   12640511
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   661
      _Version        =   393216
      Rows            =   1
      Cols            =   26
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   16761024
   End
   Begin VB.Image imgRightLeg 
      Height          =   735
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgLeftLeg 
      Height          =   735
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgTorso 
      Height          =   855
      Left            =   5640
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image imgRightArm 
      Height          =   855
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgHead 
      Height          =   975
      Left            =   5640
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Answer:"
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Click On A Letter To Select It:"
      Height          =   255
      Left            =   4800
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.Line Line5 
      X1              =   6360
      X2              =   6360
      Y1              =   1680
      Y2              =   2040
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00004080&
      BorderWidth     =   5
      X1              =   4200
      X2              =   5160
      Y1              =   2640
      Y2              =   1680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404080&
      BorderWidth     =   10
      X1              =   6720
      X2              =   4200
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404080&
      BorderWidth     =   10
      X1              =   4200
      X2              =   4200
      Y1              =   6240
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404080&
      BorderWidth     =   10
      X1              =   4200
      X2              =   7440
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Image imgLeftArm 
      Height          =   975
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Menu mnuNewWord 
      Caption         =   "&New Word"
      Begin VB.Menu mnuPick 
         Caption         =   "Pick New Word"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
      Begin VB.Menu mnuEnd 
         Caption         =   "End Program"
      End
   End
End
Attribute VB_Name = "frmHangman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program:       Chapter 11 Project
'Name:          Richard Jackson
'Artwork:       Ann Jackson
'Sounds:        Richard and Ann Jackson
'Purpose:       This program allows the user to play a simple
'               game of hangman.  There are 20 answers stored
'               in a sequential file that are chosen from
'               randomly each time the game is reset.
'
'Variables:
'
'currentAns ...... Answer for the current game
'answer(1 to 20).. Stores all 20 answers in an array
'numOfWrong ...... Number of wrong guesses
'i and x ......... Indexes for For/Next loops
'rnum1 ........... random number to select which answer is used
'foundLetter ..... boolean that indicates if a letter has been found
'gameOver ........ boolean that indicates if game is over or not
'letter .......... stores current letter selected

Option Explicit

'set up API for playing media files

Private Declare Function mcisendstring Lib "winmm.dll" _
    Alias "mciSendStringA" _
    (ByVal lpstrCommand As String, _
     ByVal lpstrReturnstring As String, _
     ByVal uReturnLength As Long, _
     ByVal hwndCallback As Long) As Long

Dim currentAns As String
Dim answer(1 To 20) As String
Dim numOfWrong As Integer

Private Sub Form_Load()
   
    Dim i As Integer
    Dim x As Integer
    Dim nReturn As Long
    
    'open media files
    
    nReturn = mcisendstring("Open yes.wav ALIAS yes TYPE Waveaudio wait", "", 0, 0)
    nReturn = mcisendstring("Open won.wav ALIAS won TYPE Waveaudio wait", "", 0, 0)
    nReturn = mcisendstring("Open no.wav ALIAS no TYPE Waveaudio wait", "", 0, 0)
    nReturn = mcisendstring("Open scream.wav ALIAS scream TYPE Waveaudio wait", "", 0, 0)
       
    'setup the grid to hold all 26 letters
    
    imgHead.Picture = LoadPicture(App.Path & "\head.jpg")
    imgLeftArm.Picture = LoadPicture(App.Path & "\leftarm.jpg")
    imgRightArm.Picture = LoadPicture(App.Path & "\rightarm.jpg")
    imgLeftLeg.Picture = LoadPicture(App.Path & "\leftleg.jpg")
    imgTorso.Picture = LoadPicture(App.Path & "\torso.jpg")
    imgRightLeg.Picture = LoadPicture(App.Path & "\rightleg.jpg")
        
    For i = 0 To 25
        Grid1.ColWidth(i) = 415
        Grid1.Row = 0
        Grid1.Col = i
        Grid1.ColAlignment(i) = 2
        Grid1.Text = Chr$(65 + i)
    Next i
   
   'read answers and assign to array
   
    Open App.Path & "\answers.txt" For Input As #1
   
    For x = 1 To 20
        Input #1, answer(x)
    Next x
    
    Close #1
   
    'call sub to create answer grid
    
    Call createAnswerGrid
          
End Sub

Private Sub createAnswerGrid()

    Dim rnum1 As Integer
    Dim i As Integer
    
    'reset random algorythm
    
    Randomize

    'select random number between 1 and 20
    'assign the array with index rnum1 to currentAns
    
    rnum1 = Int(Rnd * 20) + 1
    currentAns = answer(rnum1)
    
    'set the number of columns to the length of currentAns
    
    Grid2.Cols = Len(currentAns)
    Grid2.Row = 0
    
    'create the answer grid
    
    For i = 1 To Len(currentAns)
        Grid2.ColAlignment(i - 1) = 2
        Grid2.ColWidth(i - 1) = 415
        Grid2.Col = i - 1
        Grid2.Text = ""
    Next i
    
    'make grid visible
    
    Grid2.Visible = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call closemedia
    
End Sub

Private Sub Grid1_Click()

    Dim foundLetter As Boolean
    Dim i As Integer
    Dim letter As String
    Dim gameOver As Boolean
    Dim nReturn As Long
            

    
    'assign the letter clicked on to the variable letter
    'blank the letter clicked on from the grid
    
    letter = Grid1.Text
    Grid1.Text = ""
        
    'if the grid cell clicked contains a letter
        
    If letter <> "" Then
    
        foundLetter = False
        
        'search the current answer for the letter selected,
        'set the correct cell/s of the answer grid with the
        'letter selected, and set foundLetter to true.
        
        For i = 1 To Len(currentAns)
            If letter = Mid(currentAns, i, 1) Then
                Grid2.Col = i - 1
                Grid2.Text = letter
                foundLetter = True
            End If
        Next i
        
        'if a letter was found set gameOver to true,
        'check the answer grid for any blanks and if any are
        'found set gameOver to false.
        
        If foundLetter = True Then
            
            'play wave file yes, which say "YES!"
            
            nReturn = mcisendstring("Play yes from 0 wait", "", 0, 0)
            
            gameOver = True
            For i = 1 To Len(currentAns)
                Grid2.Col = i - 1
                If Grid2.Text = "" Then
                    gameOver = False
                End If
            Next i
            
            'if gameOver is true display congratulations,
            'disable letter grid
            
            If gameOver = True Then
                
                'play wave file won, which says "You have won"
                
                nReturn = mcisendstring("Play won from 0", "", 0, 0)
                               
                MsgBox "You Won!", , "Congratulations"
                Grid1.Enabled = False
            End If
            
        Else
            
            'play wave file no, which says, "NOPE!"
            
            nReturn = mcisendstring("Play no from 0 wait", "", 0, 0)
                        
            numOfWrong = numOfWrong + 1
            
            'if the numOfWrong guesses >= 6 draw last leg and
            'display you lose and correct answer,
            'disable letter grid.
            
            If numOfWrong >= 6 Then
                
                
                imgRightLeg.Visible = True
                                
                'play wave file scream, which SCREAMS when the man ges hung!
                
                nReturn = mcisendstring("Play scream from 0", "", 0, 0)
                                
                currentAns = "Sorry, you lose.  The answer was " & currentAns & "."
                MsgBox currentAns, , "You Lose!"
                Grid1.Enabled = False
            Else
                
                'if numOfWrong < 6 then draw appropiate body part.
                
                Select Case numOfWrong
                    Case Is = 1
                        imgHead.Visible = True
                    Case Is = 2
                        imgTorso.Visible = True
                    Case Is = 3
                        imgLeftArm.Visible = True
                    Case Is = 4
                        imgRightArm.Visible = True
                    Case Is = 5
                        imgLeftLeg.Visible = True
                End Select
            End If
        End If
    End If
End Sub

Private Sub mnuEnd_Click()
    
    Call closemedia
    
    End
    
End Sub

Private Sub mnuPick_Click()

    Dim i As Integer
    
    'initialize numOfWrong
    'enable letter grid
    
    numOfWrong = 0
    Grid1.Enabled = True
    
    'erase body parts
    
    imgHead.Visible = False
    imgRightArm.Visible = False
    imgLeftArm.Visible = False
    imgRightLeg.Visible = False
    imgLeftLeg.Visible = False
    imgTorso.Visible = False
    
    'refill letter grid
    
    For i = 0 To 25
        Grid1.ColWidth(i) = 415
        Grid1.Row = 0
        Grid1.Col = i
        Grid1.Text = Chr$(65 + i)
    Next i
    
    Call createAnswerGrid
    
End Sub

Private Sub closemedia()

    Dim nReturn As Long
    
    'close all media files
    
    nReturn = mcisendstring("Close ALIAS yes", "", 0, 0)
    nReturn = mcisendstring("Close ALIAS won", "", 0, 0)
    nReturn = mcisendstring("Close ALIAS no", "", 0, 0)
    nReturn = mcisendstring("Close ALIAS scream", "", 0, 0)

End Sub
