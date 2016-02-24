VERSION 5.00
Begin VB.Form frmManual 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tetris - Manual"
   ClientHeight    =   8955
   ClientLeft      =   5700
   ClientTop       =   2160
   ClientWidth     =   9480
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Tetris_Manual_KTsuen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   9480
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Okay"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   8160
      Width           =   1935
   End
   Begin VB.Label lblManual 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   7695
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8895
   End
End
Attribute VB_Name = "frmManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    Unload Me

End Sub

Private Sub Form_Load()
        
    CenterObject Me, True
    CenterHorizontalObj cmdOK, False, Me
    
    Dim Msg As String
    
    Msg = Space(5) & "Tetris is a puzzle game where strategy is involved in the placement of blocks "
    Msg = Msg & " with the practice of hand-eye coordination skills to see how many lines and levels "
    Msg = Msg & " one can go pass." & vbCrLf & vbCrLf
    
    Msg = Msg & Space(5) & "As you are the player, you will begin with an empty playing field and pieces"
    Msg = Msg & " will appear from the top of the playing field and fall to the bottom or to the next"
    Msg = Msg & " available spot closest to the bottom of the playing field." & vbCrLf & vbCrLf
    
    Msg = Msg & Space(5) & " Once the piece has finished dropping as it is unable to drop any further"
    Msg = Msg & " down the playing field then it will stay as blocks in the playing field and a new piece"
    Msg = Msg & " will appear at the top." & vbCrLf & vbCrLf
    
    Msg = Msg & Space(5) & "This continues until you either fill up one row or multiple rows (max 4) and"
    Msg = Msg & " the row(s) will be cleared from the playing field."
    
    Msg = Msg & " If there were any blocks above the row(s), they will drop to the next row below it."
    
    Msg = Msg & " By clearing rows, you clear lines of blocks in the playing field and you also score"
    Msg = Msg & " points."
    
    Msg = Msg & " When the number of lines required to complete the level has been cleared, you will be"
    Msg = Msg & " able to go onto the next level where the number of lines to be cleared is increased"
    Msg = Msg & " and the falling speed of the active piece will increase." & vbCrLf & vbCrLf
    
    Msg = Msg & Space(5) & "During the gameplay, the player is able to use the arrow keys and the spacebar"
    Msg = Msg & " to trigger certain actions. Here is a list of the keys and their actions:" & vbCrLf & vbCrLf
    
    Msg = Msg & Space(10) & "Up Arrow Key" & Space(5) & "-" & Space(5) & "Rotate Piece Clockwise" & vbCrLf
    Msg = Msg & Space(10) & "Down Arrow Key" & Space(5) & "-" & Space(5) & "Quick Drop" & vbCrLf
    Msg = Msg & Space(10) & "Left Arrow Key" & Space(5) & "-" & Space(5) & "Move Piece Left" & vbCrLf
    Msg = Msg & Space(10) & "Right Arrow Key" & Space(5) & "-" & Space(5) & "Move Piece Right" & vbCrLf
    Msg = Msg & Space(10) & "Spacebar" & Space(5) & "-" & Space(5) & "Instant Drop" & vbCrLf
    Msg = Msg & Space(10) & "P" & Space(5) & "-" & Space(5) & "Pause Game" & vbCrLf & vbCrLf
    
    Msg = Msg & Space(5) & "The game is over when a new piece can ot be created in the top middle area of"
    Msg = Msg & " the playing field. If you have beat one of the top 10 scores, then you may add your name"
    Msg = Msg & " to the high scores list. After that, you may start a new game or exit the game." & vbCrLf & vbCrLf
    
    Msg = Msg & "Have fun, enjoy yourself with tetris!"
    
    lblManual.Caption = Msg
    
End Sub

