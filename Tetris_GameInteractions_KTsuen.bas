Attribute VB_Name = "mdlGameInteractiions"
Option Explicit

'This procedure is used to validate that the game is not yet over, so that gameplay may continue. Checks
'   the area made up by the 4th column to the 7th column and 1st row to 2nd row.

Public Sub IsGameOver(ByRef GGrid() As GameBoard, ByRef GameOver As Boolean, ByVal CColor As _
                      ColorConstants)

    Dim K As Integer, L As Integer
    
    K = 3
    L = 2
    
    Do
        K = K + 1
        
        If GGrid(K, L).GColor <> CColor Then GameOver = True
        
        If K = 7 Then
            K = 3
            L = L - 1
        End If
    Loop While K < 7 And L > 0 And Not GameOver

End Sub

'This procedure determines that a row or multiple rows have been filled then it removes the row or the
'   multiple rows and moves the blocks above it, down.

Public Sub RemoveLine(ByRef GGrid() As GameBoard, ByRef DConsts As GridProperties, ByRef ColorConsts() _
                      As ColorConstants, ByRef ctrlOut As PictureBox, ByRef nFlash As Integer, ByRef _
                      Lines As Integer)

    Dim I As Integer, J As Integer, K As Integer, L As Integer, n As Integer, UB As Integer, LB As Integer
    Dim LineRemove As Boolean

    UB = UBound(ColorConsts)
    LB = LBound(ColorConsts)
      
    If nFlash > 0 Then
        For I = 1 To DConsts.MaxY
            J = 0
            LineRemove = True
        
            Do
                J = J + 1
                
                If GGrid(J, I).GColor = DConsts.CellColor Then LineRemove = False
            Loop While J < DConsts.MaxX And LineRemove
            
            If LineRemove = True Then
                If nFlash - 1 = 0 Then
                    For K = I To 1 Step -1
                        For L = 1 To DConsts.MaxX
                            If K <> 1 Then
                                GGrid(L, K).GColor = GGrid(L, K - 1).GColor
                            Else
                                GGrid(L, K).GColor = DConsts.CellColor
                            End If
                        Next L
                    Next K
                    Lines = Lines + 1
                Else
                    For K = 1 To DConsts.MaxX
                        n = Int(Rnd * (UB - LB + 1)) + LB
                        GGrid(K, I).GColor = RGB(128, 64, 0) 'ColorConsts(n)
                    Next K
                End If
            End If
        Next I
        
        DrawGrid GGrid(), ctrlOut, DConsts
        
        Delay 0.1
        
        RemoveLine GGrid, DConsts, ColorConsts(), ctrlOut, nFlash - 1, Lines
    End If
    
End Sub

'This procedure checks the score and adds to the score according to the points calculated by the score
'   algorithm made up of level and the number of lines cleared multiplier. It also adds to the number of
'   lines cleared and decreases the number of lines to clear.

Public Sub ScoreCheck(ByVal LCleared As Integer, ByRef Scoring As GameScore, ByRef lblLines As Label, _
                      ByRef lblScore As Label, ByRef lblLClear As Label, ByRef lblLevel As Label)
                      
    Dim Multiplier As Integer
    
    If LCleared > 0 Then
        With Scoring
            .NumLines = .NumLines + LCleared
            .LinesToClear = .LinesToClear - LCleared
            
            Select Case LCleared
                Case 1: Multiplier = 10
                Case 2: Multiplier = 20
                Case 3: Multiplier = 40
                Case 4: Multiplier = 80
            End Select
            
            .Score = .Score + .Level * Multiplier
            
            lblLClear.Caption = .LinesToClear
            lblLines.Caption = .NumLines
            lblScore.Caption = .Score
        End With
    End If

End Sub

'This procedure determines if the player has cleared enough lines to go to the next level. The number of
'   lines to be cleared increases by the level multiplied by the increment value. The dropping speed of
'   the active piece will also increase as the level increases.

Public Sub LevelCheck(ByRef Scoring As GameScore, ByVal LvlInc As Integer, ByRef Spd As Single, ByRef _
                      lblLClear As Label, ByRef lblLvl As Label)

    With Scoring
        If .LinesToClear <= 0 Then
            .Level = .Level + 1
            Spd = Spd - (Spd * 0.2)
            
            .LinesToClear = LvlInc * .Level - .NumLines Mod LvlInc

            lblLvl.Caption = "Level " & .Level
            lblLClear.Caption = .LinesToClear
        End If
    End With
    
End Sub

'This procedure validates that the player's score is one of the top ten scores and then adds it into the
'   high score table.

Public Sub HighScore(ByRef Scoring As GameScore)

    Dim Scores(1 To 10) As HighScores
    Dim FileLoc As String, Player As String
    Dim NewHighScore As Boolean, InvalidName
    Dim K As Integer, L As Integer
    Dim Prompt As String
    
    FileLoc = App.Path & "/highscore.hs"
    NewHighScore = False
    InvalidName = False
    K = 1
    
    OpenHSFile FileLoc, Scores()
        
    Do
        With Scoring
            If (.Score > Scores(K).Score Or .Score = Scores(K).Score) And .Score > 0 Then
                Do
                    Prompt = IIf(Not InvalidName, "Name:", "Invalid Name! Please try again. Name:")
                    Player = InputBox$(Prompt, "New High Score")
                    InvalidName = IIf(Trim$(Player) = "", True, False)
                Loop While InvalidName
                
                For L = UBound(Scores) To K Step -1
                    If L + 1 <= UBound(Scores) Then
                        Scores(L + 1).PlayerName = Scores(L).PlayerName
                        Scores(L + 1).Score = Scores(L).Score
                    End If
                    If L = K Then
                        Scores(L).PlayerName = Trim$(Player)
                        Scores(L).Score = .Score
                    End If
                Next L
                
                NewHighScore = True
                
                WriteHSFile FileLoc, Scores()
                
                Unload frmHighScores
                frmHighScores.Show vbModeless
            End If
            
            K = K + 1
        End With
    Loop While Not NewHighScore And K <= 10
    
End Sub
