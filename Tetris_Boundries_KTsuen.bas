Attribute VB_Name = "mdlBoundries"
Option Explicit

'This procedure validates that the active piece is not outside of the grid. Only used with the
'   RotatePiece procedure to validate that the rotated piece will be inside the grid.

Public Sub CheckOutOfBounds(ByRef APiece As GamePiece, ByVal PieceSide As Integer, ByRef IsOutOfBounds _
                            As Boolean, ByVal MaxX As Integer)
                            
    Dim X As Integer, K As Integer
    
    Do
        K = K + 1
        
        With APiece
            If K = 1 Then
                X = .PCenter.X
            Else
                X = .PPiece(K - 1).X
            End If
        End With
        
        If PieceSide = vbKeyLeft And X < 1 Then
            IsOutOfBounds = True
        ElseIf PieceSide = vbKeyRight And X > MaxX Then
            IsOutOfBounds = True
        End If
    Loop While K < 4 And Not IsOutOfBounds

End Sub

'This procedure validates that the space the active piece will occupy is available. Only used with the
'   RotatePiece procedure to validate that the space to be occupied by the rotated piece is available.

Public Sub CheckSpace(ByRef GGrid() As GameBoard, ByRef APiece As GamePiece, ByRef SpaceOccupied As _
                      Boolean, ByVal CColor As ColorConstants)

    Dim X As Integer, Y As Integer
    Dim K As Integer
    
    Do
        K = K + 1
        
        With APiece
            If K = 1 Then
                X = .PCenter.X
                Y = .PCenter.Y
            Else
                X = .PPiece(K - 1).X
                Y = .PPiece(K - 1).Y
            End If
        End With
        
        If GGrid(X, Y).GColor <> CColor Then SpaceOccupied = True
    Loop While K < 4 And SpaceOccupied = False

End Sub

'This procedure validates that the active piece is not at the bottom of the grid. Only used with the
'   DropPiece procedure to validate that the active piece is still in motion and with the HTranslatePiece
'   procedure to determine when the active piece is motionless so that horizontal movement is not
'   possible.

Public Sub CheckBottom(ByRef APiece As GamePiece, ByRef PastBottom As Boolean, ByRef DConsts As _
                       GridProperties, Optional ByRef NPiece As Boolean)

    Dim Y As Integer, K As Integer
    
    Do
        K = K + 1
        
        With APiece
            If K = 1 Then
                Y = .PCenter.Y + 1
            Else
                Y = .PPiece(K - 1).Y + 1
            End If
        End With
        
        If Y > DConsts.MaxY Then
            PastBottom = True
            NPiece = True
        End If
    Loop While K < 4 And PastBottom = False
    
End Sub

'This procedure validates that the active piece can move in one of the horizontal directions, which is
'   passed through the PieceSide parameter, and that the edge in that direction is not obstructing its
'   movement. Only used in the HTranslatePiece procedure to validate that the active piece can move in the
'   direction specified.

Public Sub CheckSide(ByRef APiece As GamePiece, ByRef PieceSide As Integer, ByRef PastSide As Boolean, _
                     ByVal MaxX As Integer)

    Dim X As Integer, K As Integer
    Dim PSide As Integer
    
    If PieceSide = vbKeyLeft Then
        PSide = -1
    ElseIf PieceSide = vbKeyRight Then
        PSide = 1
    End If
    
    Do
        K = K + 1
        
        With APiece
            If K = 1 Then
                X = .PCenter.X + PSide
            Else
                X = .PPiece(K - 1).X + PSide
            End If
        End With
        
        If PieceSide = vbKeyLeft And X < 1 Then
            PastSide = True
        ElseIf PieceSide = vbKeyRight And X > MaxX Then
            PastSide = True
        End If
    Loop While K < 4 And PastSide = False
    
End Sub

'This procedure validates that the active piece can still move down and that no existing blocks are
'   occupying the space that the active piece will move into. Only used with the DropPiece procedure
'   to validate that downward movement is still possible and with the HTranslatePiece to validate
'   that a new piece has not occured yet so movement is still possible.

Public Sub CheckBelowPiece(ByRef GGrid() As GameBoard, ByRef APiece As GamePiece, ByRef PieceBelow As _
                           Boolean, ByVal CColor As ColorConstants, Optional ByRef NPiece As Boolean)

    Dim X As Integer, Y As Integer, K As Integer
    
    Do
        K = K + 1
        
        With APiece
            If K = 1 Then
                X = .PCenter.X
                Y = .PCenter.Y
            Else
                X = .PPiece(K - 1).X
                Y = .PPiece(K - 1).Y
            End If
        End With
    
        If GGrid(X, Y + 1).GColor <> CColor Then
            PieceBelow = True
            NPiece = True
        Else
            NPiece = False
        End If
    Loop While K < 4 And PieceBelow = False

End Sub

'This procedure validates that the space that the active piece is about to move into is available and free
'   of any existing pieces. Only used with the HTranslatePiece procedure to validate that horizontal
'   movement in specified direction is possible.

Public Sub CheckBesidePiece(ByRef GGrid() As GameBoard, ByRef APiece As GamePiece, ByRef PieceSide As _
                            Integer, ByRef PieceBeside As Boolean, ByVal CColor As ColorConstants)
    
    Dim X As Integer, Y As Integer, K As Integer, PSide As Integer
    
    If PieceSide = vbKeyLeft Then
        PSide = -1
    ElseIf PieceSide = vbKeyRight Then
        PSide = 1
    End If
    
    Do
        K = K + 1
        
        With APiece
            If K = 1 Then
                X = .PCenter.X
                Y = .PCenter.Y
            Else
                X = .PPiece(K - 1).X
                Y = .PPiece(K - 1).Y
            End If
        End With
    
        If GGrid(X + PSide, Y).GColor <> CColor Then PieceBeside = True
    Loop While K < 4 And PieceBeside = False
    
End Sub

