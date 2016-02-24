Attribute VB_Name = "mdlDrawGrid"
Option Explicit

'This procedure is used to create a new active piece, move an active piece downward.

Public Sub ChangeGrid(ByRef GGrid() As GameBoard, ByRef APiece As GamePiece, ByRef DConsts As _
                      GridProperties, ByRef SConsts As Pieces, ByRef NewPiece As Boolean)
    
    If NewPiece = True Then
        NewPiece = False
        
        DeterminePiece APiece.PShape, APiece.PColor, SConsts
        CreatePiece APiece, SConsts
        InputToGrid GGrid(), APiece, DConsts.CellColor, False
    Else
        InputToGrid GGrid(), APiece, DConsts.CellColor, True
        DropPiece GGrid(), APiece, NewPiece, DConsts
        InputToGrid GGrid(), APiece, DConsts.CellColor, False
    End If
    
End Sub

'This procedure inputs the active piece to the grid once changes have been made to the active piece or it
'   removes the active piece from the grid so that changes can be made to the active piece.

Public Sub InputToGrid(ByRef GGrid() As GameBoard, ByRef APiece As GamePiece, ByVal CColor As _
                       ColorConstants, ByVal RemovePiece As Boolean)

    Dim X As Integer, Y As Integer, K As Integer
    Dim CellColor As ColorConstants
    
    For K = 1 To 4
        With APiece
            If K = 1 Then
                X = .PCenter.X
                Y = .PCenter.Y
            Else
                X = .PPiece(K - 1).X
                Y = .PPiece(K - 1).Y
            End If
            
            CellColor = IIf(RemovePiece = True, CColor, .PColor)
            
            GGrid(X, Y).GColor = CellColor
        End With
    Next K
    
End Sub

'This procedure draws the grid to the game interface once all changes have
'   been made to the grid.

Public Sub DrawGrid(ByRef GGrid() As GameBoard, ByRef ctrlOut As PictureBox, ByRef DConsts As _
                    GridProperties)

    Dim K As Integer, L As Integer
    Dim StartX As Integer, StartY As Integer
    Dim FinishX As Integer, FinishY As Integer
    Dim BgColor As ColorConstants

    With DConsts
        For K = 1 To .MaxX
            StartX = (K - 1) * .CellSize + .LineSize
            FinishX = K * .CellSize - .LineSize
            
            For L = 1 To .MaxY
                StartY = (L - 1) * .CellSize + .LineSize
                FinishY = L * .CellSize - .LineSize
                
                BgColor = GGrid(K, L).GColor
                
                ctrlOut.Line (StartX, StartY)-(FinishX, FinishY), BgColor, BF
            Next L
        Next K
    End With
    
End Sub

