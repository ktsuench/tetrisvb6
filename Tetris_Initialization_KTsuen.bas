Attribute VB_Name = "mdlInitialization"
Option Explicit

'This procedure is used to reset the variables and controls used in the game.

Public Sub ResetGame(ByRef APiece As GamePiece, ByRef GGrid() As GameBoard, ByRef ctrlOut As PictureBox, _
                     ByRef DConsts As GridProperties, ByRef ctrlFrame As Frame, ByRef ctrlLabel As Label, _
                     ByRef frmOut As Form)
    
    ResetPiece APiece, DConsts.CellColor
    ResetGrid GGrid(), ctrlOut, DConsts
    
    ctrlFrame.Visible = False
    ctrlLabel.Caption = ""
    
End Sub

'This procedure is used to clear the active piece of previous information.

Public Sub ResetPiece(ByRef GPiece As GamePiece, ByVal CColor As ColorConstants)

    Dim K As Integer

    With GPiece
        .PShape = -1
        .PColor = CColor
        .PPosition = 1
        .PCenter.X = 0
        .PCenter.Y = 0
        
        For K = 1 To 3
            .PPiece(K).X = 0
            .PPiece(K).Y = 0
        Next K
    End With

End Sub

'This procedure is used to clear the grid of any existing blocks from a previous game.

Public Sub ResetGrid(ByRef GGrid() As GameBoard, ByRef ctrlOut As Control, ByRef DConsts As GridProperties)
    
    Dim K As Integer, L As Integer
    
    Erase GGrid
    
    With DConsts
        ctrlOut.Line (0, 0)-(ctrlOut.Width, ctrlOut.Height), .LineColor, BF
        
        For K = 1 To .MaxX
            For L = 1 To .MaxY
                GGrid(K, L).GColor = .CellColor
            Next L
        Next K
        
        DrawGrid GGrid(), ctrlOut, DConsts
    End With
    
End Sub
