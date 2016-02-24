VERSION 5.00
Begin VB.Form frmGame 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tetris"
   ClientHeight    =   5580
   ClientLeft      =   12435
   ClientTop       =   4410
   ClientWidth     =   4785
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Tetris_GameInterface_KTsuen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   4785
   Begin VB.Frame fraGameOver 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   320
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label lblGameOver 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1455
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2175
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5120
      Left            =   240
      ScaleHeight     =   5115
      ScaleWidth      =   2565
      TabIndex        =   0
      Top             =   240
      Width           =   2570
   End
   Begin VB.Label lblLClear 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblLClearCap 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblLinesCap 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblScoreCap 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblLines 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblLvl 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuPlay 
         Caption         =   "&Play"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuPause 
         Caption         =   "P&ause"
         Shortcut        =   {F3}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "&End Game"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHighScores 
         Caption         =   "&Show High Scores"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuManual 
         Caption         =   "&Manual"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   {F8}
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Title:     Final Project - Tetris
'Author:    Kent Tsuen-Chy
'Date:      Monday, June 2, 2014
'Files:     Tetris_About_KTsuen.frm, Tetris_About_KTsuen.frx, Tetris_Boundries_KTsuen.bas,
'           Tetris_CreatePiece_KTsuen.bas, Tetris_Declarations_KTsuen.bas, Tetris_DrawGrid_KTsuen.bas,
'           Tetris_GameInteractions_KTsuen.bas, Tetris_GameInterface_KTsuen.frm,
'           Tetris_GameInterface_KTsuen.frx, Tetris_General_KTsuen.bas, Tetris_High_Scores_KTsuen.frm,
'           Tetris_High_Scores_KTsuen.frx, Tetris_Initialization_KTsuen.bas, Tetris_KTsuen.vbp,
'           Tetris_KTsuen.vbw, Tetris_Manual_KTsuen.frm, Tetris_Manual_KTsuen.frx, Tetris_Menu_KTsuen.frm,
'           Tetris_Menu_KTsuen.frx, Tetris_PieceInteractions_KTsuen.bas
'Purpose:   This application is a puzzle game that involves hand-eye coordination and the ability to think
'           quickly. This application is a version of the well known Tetris game. For more information on
'           this application, visit the manual of this application.

'This form contains the game controls and game interface of Tetris.

Const MAX_X = 10
Const MAX_Y = 20
Const NEXT_LEVEL_INCREMENT = 15

Dim PieceConstants As Pieces
Dim DrawingConstants As GridProperties
Dim PColorConstants(1 To 7) As ColorConstants

Dim GameOver As Boolean
Dim PauseGame As Boolean
Dim NewPiece As Boolean
        
Dim Speed As Single
Dim PrevSpeed As Single
Dim QuickDrop As Boolean
Dim InstantDropped As Boolean
Dim KeyDown As Boolean

Dim Scoring As GameScore
Dim CurLinesCleared As Integer

Dim DelayNewPiece As Boolean
Dim DelayEndGame As Boolean

Dim ActivePiece As GamePiece

Dim GameGrid(1 To MAX_X, 1 To MAX_Y) As GameBoard

Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim DMsg As String, DTitle As String

    DMsg = "Are you sure you want to exit the game?"
    DTitle = "Exit Game"
    
    If KeyCode = vbKeyF4 And Shift = vbAltMask Then
        If PromptDialog(DMsg, DTitle) Then GameOver = True: End
    ElseIf KeyCode = vbKeyP Then
            PauseGame = Not PauseGame
        mnuPause.Checked = Not mnuPause.Checked
    
        If PauseGame = True Then
            picGrid.Enabled = False
            Me.SetFocus
        Else
            picGrid.Enabled = True
            picGrid.SetFocus
        End If
    End If
    
End Sub

Private Sub Form_Load()
    
    Unload frmMenu
    
    CenterObject Me, True

    ResetGame ActivePiece, GameGrid(), picGrid, DrawingConstants, fraGameOver, lblGameOver, Me
    
    picGrid.Enabled = False
    
    With PieceConstants
        .Z = 1
        .S = 2
        .T = 3
        .O = 4
        .L = 5
        .J = 6
        .I = 7
    End With
    
    With DrawingConstants
        .MaxX = MAX_X
        .MaxY = MAX_Y
        .CellColor = &H202020
        .CellSize = 255
        .LineColor = &H404040
        .LineSize = 30
    End With
    
    PColorConstants(1) = RGB(255, 28, 28)   'Red
    PColorConstants(2) = RGB(28, 255, 28)   'Green
    PColorConstants(3) = RGB(128, 0, 255)   'Purple
    PColorConstants(4) = RGB(255, 255, 55)  'Yellow
    PColorConstants(5) = RGB(255, 255, 255) 'White
    PColorConstants(6) = RGB(255, 128, 0)   'Orange
    PColorConstants(7) = RGB(0, 128, 255)   'Blue
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim DMsg As String, DTitle As String

    DMsg = "Are you sure you want to exit the game?"
    DTitle = "Exit Game"
    
    
    If PromptDialog(DMsg, DTitle) Then
        GameOver = True
        End
    Else
        Cancel = True
    End If

End Sub

Private Sub mnuAbout_Click()
    
    PauseGame = True
    frmAbout.Show vbModal

End Sub

Private Sub mnuEnd_Click()
    
    Dim DMsg As String, DTitle As String
    
    DMsg = "Are you sure you want to end your current game?"
    DTitle = "End Current Game?"
    
    If PromptDialog(DMsg, DTitle) Then GameOver = True

End Sub

Private Sub mnuHighScores_Click()

    PauseGame = True
    frmHighScores.Show vbModeless

End Sub

Private Sub mnuManual_Click()

    PauseGame = True
    frmManual.Show vbModeless

End Sub

Private Sub mnuPause_Click()

    PauseGame = Not PauseGame
    mnuPause.Checked = Not mnuPause.Checked
    
    If PauseGame Then
        picGrid.Enabled = False
        Me.SetFocus
    Else
        picGrid.Enabled = True
        picGrid.SetFocus
    End If

End Sub

Public Sub mnuPlay_Click()

    mnuPause.Visible = True
    picGrid.Enabled = True
    picGrid.SetFocus
    picGrid.Cls
    lblGameOver.Caption = ""
    
    Dim K As Integer, Y As Integer
       
    ResetGame ActivePiece, GameGrid(), picGrid, DrawingConstants, fraGameOver, lblGameOver, Me
    
    picGrid.Enabled = False
    
    DisplayStatusMsg "3.2.1" & vbCrLf & "START", "3.2.1" & vbCrLf & "START", fraGameOver, lblGameOver
    Delay 0.5: DoEvents
    fraGameOver.Visible = False
    lblGameOver.Caption = ""
    picGrid.Enabled = True
    
    GameOver = False
    PauseGame = False
    NewPiece = True
    Speed = 0.7
    PrevSpeed = Speed
    KeyDown = False
    QuickDrop = False
    DelayNewPiece = False
    
    With Scoring
        .Level = 1
        .LinesToClear = NEXT_LEVEL_INCREMENT
        .Score = 0
        .NumLines = 0
    
        lblLvl.Caption = "Level " & .Level
        lblLClearCap.Caption = "Lines To Clear:"
        lblLClear.Caption = .LinesToClear
        lblLinesCap.Caption = "Lines Removed:"
        lblLines.Caption = .NumLines
        lblScoreCap.Caption = "Score:"
        lblScore.Caption = .Score
    End With
    
    Do
        If Not PauseGame Then
            If Not DelayNewPiece Then
                ChangeGrid GameGrid(), ActivePiece, DrawingConstants, PieceConstants, NewPiece
            End If
            
            DrawGrid GameGrid(), picGrid, DrawingConstants
            
            If InstantDropped Then InstantDropped = False
            If NewPiece And DelayNewPiece Then Delay 0.1: DoEvents
            
            DelayNewPiece = False
            
            If NewPiece Then
                CurLinesCleared = 0
                RemoveLine GameGrid(), DrawingConstants, PColorConstants(), picGrid, 5, CurLinesCleared
                
                ScoreCheck CurLinesCleared, Scoring, lblLines, lblScore, lblLClear, lblLvl
                LevelCheck Scoring, NEXT_LEVEL_INCREMENT, PrevSpeed, lblLClear, lblLvl
                
                IsGameOver GameGrid(), GameOver, DrawingConstants.CellColor
            End If
            
            Delay Speed: DoEvents
            If Not QuickDrop Then Speed = PrevSpeed
        End If
        DoEvents
    Loop While Not GameOver
    
    HighScore Scoring
    
    DisplayStatusMsg "GAME OVER", "GAME" & vbCrLf & "OVER", fraGameOver, lblGameOver
    
    mnuPause.Visible = False
    picGrid.Enabled = False
    
End Sub

Private Sub picGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim MotionEnd As Boolean
    
    MotionEnd = False
    
    Select Case KeyCode
        Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeySpace
            InputToGrid GameGrid(), ActivePiece, DrawingConstants.CellColor, True
            Select Case KeyCode
                Case vbKeyLeft, vbKeyRight
                    If NewPiece And Not InstantDropped Then
                        DelayNewPiece = True
                        HTranslatePiece GameGrid(), ActivePiece, KeyCode, DrawingConstants, NewPiece
                    ElseIf Not InstantDropped Then
                        KeyDown = True
                        HTranslatePiece GameGrid(), ActivePiece, KeyCode, DrawingConstants
                    End If
                Case vbKeyUp
                    If Not NewPiece Then
                        If ActivePiece.PShape <> PieceConstants.O Then
                            RotatePiece GameGrid(), ActivePiece, DrawingConstants, PieceConstants
                        End If
                    End If
                Case vbKeyDown
                    If Not NewPiece Then
                        Speed = 0.01
                        QuickDrop = True
                        ChangeGrid GameGrid(), ActivePiece, DrawingConstants, PieceConstants, NewPiece
                                              
                        If NewPiece Then
                            CurLinesCleared = 0
                            RemoveLine GameGrid(), DrawingConstants, PColorConstants(), picGrid, 5, _
                                CurLinesCleared
                                
                            ScoreCheck CurLinesCleared, Scoring, lblLines, lblScore, lblLClear, lblLvl
                            LevelCheck Scoring, NEXT_LEVEL_INCREMENT, Speed, lblLClear, lblLvl
                            
                            IsGameOver GameGrid(), GameOver, DrawingConstants.CellColor
                            
                            If NewPiece Then
                                ChangeGrid GameGrid(), ActivePiece, DrawingConstants, PieceConstants, _
                                    NewPiece
                            End If
                        End If
                    End If
                Case vbKeySpace
                    CheckBottom ActivePiece, MotionEnd, DrawingConstants, NewPiece

                    If Not MotionEnd Then
                        CheckBelowPiece GameGrid(), ActivePiece, MotionEnd, DrawingConstants.CellColor, _
                            NewPiece
                    End If
                    
                    If Not MotionEnd Then
                        InstantDropped = True
                        
                        InputToGrid GameGrid(), ActivePiece, DrawingConstants.CellColor, True
                        InstantDrop GameGrid(), ActivePiece, NewPiece, DrawingConstants
                        InputToGrid GameGrid(), ActivePiece, DrawingConstants.CellColor, False
                        If NewPiece And Not DelayNewPiece Then
                            CurLinesCleared = 0
                            RemoveLine GameGrid(), DrawingConstants, PColorConstants(), picGrid, 5, _
                                CurLinesCleared

                            ScoreCheck CurLinesCleared, Scoring, lblLines, lblScore, lblLClear, lblLvl
                            LevelCheck Scoring, NEXT_LEVEL_INCREMENT, Speed, lblLClear, lblLvl
                            
                            PrevSpeed = Speed
                            
                            IsGameOver GameGrid(), GameOver, DrawingConstants.CellColor
                            
                            If NewPiece Then
                                ChangeGrid GameGrid(), ActivePiece, DrawingConstants, PieceConstants, _
                                    NewPiece
                            End If
                        End If
                    End If
            End Select
            InputToGrid GameGrid(), ActivePiece, DrawingConstants.CellColor, False
            DrawGrid GameGrid(), picGrid, DrawingConstants
            If GameOver Then
                mnuPause.Visible = False
                picGrid.Enabled = False
            End If
        Case vbKeyP
            If Not GameOver Then
                PauseGame = Not PauseGame
                mnuPause.Checked = Not mnuPause.Checked
                
                If PauseGame Then
                    picGrid.Enabled = False
                Else
                    picGrid.Enabled = True
                    picGrid.SetFocus
                End If
            End If
    End Select
    
End Sub

Private Sub picGrid_KeyUp(KeyCode As Integer, Shift As Integer)
        
    If QuickDrop And KeyCode = vbKeyDown Then QuickDrop = False
    If KeyDown And (KeyCode = vbKeyLeft Or KeyCode = vbKeyRight) Then KeyDown = False
    
End Sub
