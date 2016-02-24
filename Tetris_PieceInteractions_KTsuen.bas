Attribute VB_Name = "mdlPieceInteractions"
Option Explicit

'This procedure is determines and moves the active piece downward, when possible.

Public Sub DropPiece(ByRef GGrid() As GameBoard, ByRef APiece As GamePiece, ByRef NPiece As Boolean, _
                     ByRef DConsts As GridProperties)

    Dim Y As Integer, K As Integer
    Dim AtBoundary As Boolean
    
    AtBoundary = False

    CheckBottom APiece, AtBoundary, DConsts, NPiece
    
    If Not AtBoundary Then CheckBelowPiece GGrid(), APiece, AtBoundary, DConsts.CellColor, NPiece
    
    If Not AtBoundary Then
        With APiece
            .PCenter.Y = .PCenter.Y + 1
            
            For K = 1 To 3
                .PPiece(K).Y = .PPiece(K).Y + 1
            Next K
        End With
    End If
    
End Sub

'This procedure determines and moves the active piece in a specified horizontal direction, when possible.

Public Sub HTranslatePiece(ByRef GGrid() As GameBoard, ByRef APiece As GamePiece, PDirection As Integer, _
                           ByRef DConsts As GridProperties, Optional ByRef NPiece As Boolean = vbNull)
    
    Dim K As Integer, PSide As Integer
    Dim AtBoundary As Boolean, AtBottom As Boolean
    
    AtBoundary = False
    
    CheckSide APiece, PDirection, AtBoundary, DConsts.MaxX
    
    If Not AtBoundary Then CheckBesidePiece GGrid(), APiece, PDirection, AtBoundary, DConsts.CellColor
    
    If NPiece <> vbNull Then
        CheckBottom APiece, AtBottom, DConsts, NPiece

        If Not AtBottom Then CheckBelowPiece GGrid(), APiece, AtBottom, DConsts.CellColor, NPiece
    End If

    If Not AtBoundary Or (Not AtBoundary And AtBottom) Then
        If PDirection = vbKeyLeft Then
            PSide = -1
        ElseIf PDirection = vbKeyRight Then
            PSide = 1
        End If
        
        With APiece
            .PCenter.X = .PCenter.X + PSide
            
            For K = 1 To 3
                .PPiece(K).X = .PPiece(K).X + PSide
            Next K
        End With
    'Else
    '    If Not AtBottom Then CheckBelowPiece GGrid(), APiece, AtBottom, DConsts.CellColor, NPiece
    End If
    
End Sub

'This procedure determines and rotates the active piece, when possible.

Public Sub RotatePiece(ByRef GGrid() As GameBoard, ByRef APiece As GamePiece, ByRef DConsts As _
                       GridProperties, ByRef SConsts As Pieces)

    Dim RotatedPiece As GamePiece
    Dim IsOccupied As Boolean, OutOfBounds As Boolean
    Dim K As Integer

    With RotatedPiece
        .PShape = APiece.PShape
        .PPosition = APiece.PPosition
        
        .PCenter.X = APiece.PCenter.X
        .PCenter.Y = APiece.PCenter.Y
        
        For K = 1 To 3
            .PPiece(K).X = APiece.PPiece(K).X
            .PPiece(K).Y = APiece.PPiece(K).Y
        Next K
    End With

    If RotatedPiece.PCenter.Y < 3 Or RotatedPiece.PCenter.Y > DConsts.MaxY - 1 Then Exit Sub
    
    ChangePosition RotatedPiece, SConsts
    
    CheckOutOfBounds RotatedPiece, vbKeyLeft, OutOfBounds, DConsts.MaxX
    
    If Not OutOfBounds Then
        CheckOutOfBounds RotatedPiece, vbKeyRight, OutOfBounds, DConsts.MaxX
        
        If OutOfBounds Then
            With RotatedPiece
                If .PShape = SConsts.I And .PPosition = 3 Then
                    Do
                        .PCenter.X = .PCenter.X - 2
                    Loop While .PCenter.X > DConsts.MaxX
                Else
                    Do
                        .PCenter.X = .PCenter.X - 1
                    Loop While .PCenter.X > DConsts.MaxX
                End If
            End With
        End If
    Else
        With RotatedPiece
            If .PShape = SConsts.I And .PPosition = 1 Then
                Do
                    .PCenter.X = .PCenter.X + 1
                Loop While .PCenter.X < 3
            Else
                Do
                    .PCenter.X = .PCenter.X + 1
                Loop While .PCenter.X < 1
            End If
        End With
    End If
    
    If OutOfBounds Then
        RotatedPiece.PPosition = APiece.PPosition
        
        ChangePosition RotatedPiece, SConsts
    End If
    
    CheckSpace GGrid(), RotatedPiece, IsOccupied, DConsts.CellColor

    If Not IsOccupied Then
        With APiece
            .PPosition = RotatedPiece.PPosition
            
            .PCenter.X = RotatedPiece.PCenter.X
            .PCenter.Y = RotatedPiece.PCenter.Y
            
            For K = 1 To 3
                .PPiece(K).X = RotatedPiece.PPiece(K).X
                .PPiece(K).Y = RotatedPiece.PPiece(K).Y
            Next K
        End With
    End If

End Sub

'This procedure rotates the active piece by moving the blocks around.

Public Sub ChangePosition(ByRef APiece As GamePiece, ByRef SConsts As Pieces)

    With APiece
        Select Case .PShape
            Case SConsts.Z
                Select Case .PPosition
                    Case 1
                        .PPiece(1).X = .PCenter.X + 1
                        .PPiece(2).X = .PCenter.X + 1
                        .PPiece(3).X = .PCenter.X
                        .PPiece(1).Y = .PCenter.Y - 1
                        .PPiece(2).Y = .PCenter.Y
                        .PPiece(3).Y = .PCenter.Y + 1
                        
                        .PPosition = 2
                    Case 2
                        .PPiece(1).X = .PCenter.X + 1
                        .PPiece(2).X = .PCenter.X
                        .PPiece(3).X = .PCenter.X - 1
                        .PPiece(1).Y = .PCenter.Y + 1
                        .PPiece(2).Y = .PCenter.Y + 1
                        .PPiece(3).Y = .PCenter.Y
                        
                        .PPosition = 3
                    Case 3
                        .PPiece(1).X = .PCenter.X - 1
                        .PPiece(2).X = .PCenter.X - 1
                        .PPiece(3).X = .PCenter.X
                        .PPiece(1).Y = .PCenter.Y + 1
                        .PPiece(2).Y = .PCenter.Y
                        .PPiece(3).Y = .PCenter.Y - 1
                        
                        .PPosition = 4
                    Case 4
                        .PPiece(1).X = .PCenter.X - 1
                        .PPiece(2).X = .PCenter.X
                        .PPiece(3).X = .PCenter.X + 1
                        .PPiece(1).Y = .PCenter.Y - 1
                        .PPiece(2).Y = .PCenter.Y - 1
                        .PPiece(3).Y = .PCenter.Y
                        
                        .PPosition = 1
                End Select
            Case SConsts.S
                Select Case .PPosition
                    Case 1
                        .PPiece(1).X = .PCenter.X
                        .PPiece(2).X = .PCenter.X + 1
                        .PPiece(3).X = .PCenter.X + 1
                        .PPiece(1).Y = .PCenter.Y - 1
                        .PPiece(2).Y = .PCenter.Y
                        .PPiece(3).Y = .PCenter.Y + 1
                        
                        .PPosition = 2
                    Case 2
                        .PPiece(1).X = .PCenter.X + 1
                        .PPiece(2).X = .PCenter.X
                        .PPiece(3).X = .PCenter.X - 1
                        .PPiece(1).Y = .PCenter.Y
                        .PPiece(2).Y = .PCenter.Y + 1
                        .PPiece(3).Y = .PCenter.Y + 1
                        
                        .PPosition = 3
                    Case 3
                        .PPiece(1).X = .PCenter.X
                        .PPiece(2).X = .PCenter.X - 1
                        .PPiece(3).X = .PCenter.X - 1
                        .PPiece(1).Y = .PCenter.Y + 1
                        .PPiece(2).Y = .PCenter.Y
                        .PPiece(3).Y = .PCenter.Y - 1
                        
                        .PPosition = 4
                    Case 4
                        .PPiece(1).X = .PCenter.X - 1
                        .PPiece(2).X = .PCenter.X
                        .PPiece(3).X = .PCenter.X + 1
                        .PPiece(1).Y = .PCenter.Y
                        .PPiece(2).Y = .PCenter.Y - 1
                        .PPiece(3).Y = .PCenter.Y - 1
                        
                        .PPosition = 1
                End Select
            Case SConsts.T
                Select Case .PPosition
                    Case 1
                        .PPiece(1).X = .PCenter.X
                        .PPiece(2).X = .PCenter.X - 1
                        .PPiece(3).X = .PCenter.X
                        .PPiece(1).Y = .PCenter.Y - 1
                        .PPiece(2).Y = .PCenter.Y
                        .PPiece(3).Y = .PCenter.Y + 1
                        
                        .PPosition = 2
                    Case 2
                        .PPiece(1).X = .PCenter.X + 1
                        .PPiece(2).X = .PCenter.X
                        .PPiece(3).X = .PCenter.X - 1
                        .PPiece(1).Y = .PCenter.Y
                        .PPiece(2).Y = .PCenter.Y - 1
                        .PPiece(3).Y = .PCenter.Y
                        
                        .PPosition = 3
                    Case 3
                        .PPiece(1).X = .PCenter.X
                        .PPiece(2).X = .PCenter.X + 1
                        .PPiece(3).X = .PCenter.X
                        .PPiece(1).Y = .PCenter.Y + 1
                        .PPiece(2).Y = .PCenter.Y
                        .PPiece(3).Y = .PCenter.Y - 1
                        
                        .PPosition = 4
                    Case 4
                        .PPiece(1).X = .PCenter.X - 1
                        .PPiece(2).X = .PCenter.X
                        .PPiece(3).X = .PCenter.X + 1
                        .PPiece(1).Y = .PCenter.Y
                        .PPiece(2).Y = .PCenter.Y + 1
                        .PPiece(3).Y = .PCenter.Y
                        
                        .PPosition = 1
                End Select
            Case SConsts.L
                Select Case .PPosition
                    Case 1
                        .PPiece(1).X = .PCenter.X
                        .PPiece(2).X = .PCenter.X
                        .PPiece(3).X = .PCenter.X - 1
                        .PPiece(1).Y = .PCenter.Y - 1
                        .PPiece(2).Y = .PCenter.Y + 1
                        .PPiece(3).Y = .PCenter.Y - 1
                        
                        .PPosition = 2
                    Case 2
                        .PPiece(1).X = .PCenter.X + 1
                        .PPiece(2).X = .PCenter.X - 1
                        .PPiece(3).X = .PCenter.X + 1
                        .PPiece(1).Y = .PCenter.Y
                        .PPiece(2).Y = .PCenter.Y
                        .PPiece(3).Y = .PCenter.Y - 1
                        
                        .PPosition = 3
                    Case 3
                        .PPiece(1).X = .PCenter.X
                        .PPiece(2).X = .PCenter.X
                        .PPiece(3).X = .PCenter.X + 1
                        .PPiece(1).Y = .PCenter.Y + 1
                        .PPiece(2).Y = .PCenter.Y - 1
                        .PPiece(3).Y = .PCenter.Y + 1
                        
                        .PPosition = 4
                    Case 4
                        .PPiece(1).X = .PCenter.X - 1
                        .PPiece(2).X = .PCenter.X + 1
                        .PPiece(3).X = .PCenter.X - 1
                        .PPiece(1).Y = .PCenter.Y
                        .PPiece(2).Y = .PCenter.Y
                        .PPiece(3).Y = .PCenter.Y + 1
                        
                        .PPosition = 1
                End Select
            Case SConsts.J
                Select Case .PPosition
                    Case 1
                        .PPiece(1).X = .PCenter.X
                        .PPiece(2).X = .PCenter.X
                        .PPiece(3).X = .PCenter.X - 1
                        .PPiece(1).Y = .PCenter.Y - 1
                        .PPiece(2).Y = .PCenter.Y + 1
                        .PPiece(3).Y = .PCenter.Y + 1
                        
                        .PPosition = 2
                    Case 2
                        .PPiece(1).X = .PCenter.X + 1
                        .PPiece(2).X = .PCenter.X - 1
                        .PPiece(3).X = .PCenter.X - 1
                        .PPiece(1).Y = .PCenter.Y
                        .PPiece(2).Y = .PCenter.Y
                        .PPiece(3).Y = .PCenter.Y - 1
                        
                        .PPosition = 3
                    Case 3
                        .PPiece(1).X = .PCenter.X
                        .PPiece(2).X = .PCenter.X
                        .PPiece(3).X = .PCenter.X + 1
                        .PPiece(1).Y = .PCenter.Y + 1
                        .PPiece(2).Y = .PCenter.Y - 1
                        .PPiece(3).Y = .PCenter.Y - 1
                        
                        .PPosition = 4
                    Case 4
                        .PPiece(1).X = .PCenter.X - 1
                        .PPiece(2).X = .PCenter.X + 1
                        .PPiece(3).X = .PCenter.X + 1
                        .PPiece(1).Y = .PCenter.Y
                        .PPiece(2).Y = .PCenter.Y
                        .PPiece(3).Y = .PCenter.Y + 1
                        
                        .PPosition = 1
                End Select
            Case SConsts.I
                Select Case .PPosition
                    Case 1
                        .PPiece(1).X = .PCenter.X
                        .PPiece(2).X = .PCenter.X
                        .PPiece(3).X = .PCenter.X
                        .PPiece(1).Y = .PCenter.Y - 2
                        .PPiece(2).Y = .PCenter.Y - 1
                        .PPiece(3).Y = .PCenter.Y + 1
                        
                        .PPosition = 2
                    Case 2
                        .PPiece(1).X = .PCenter.X - 1
                        .PPiece(2).X = .PCenter.X + 1
                        .PPiece(3).X = .PCenter.X + 2
                        .PPiece(1).Y = .PCenter.Y
                        .PPiece(2).Y = .PCenter.Y
                        .PPiece(3).Y = .PCenter.Y
                    
                        .PPosition = 3
                    Case 3
                        .PPiece(1).X = .PCenter.X
                        .PPiece(2).X = .PCenter.X
                        .PPiece(3).X = .PCenter.X
                        .PPiece(1).Y = .PCenter.Y + 2
                        .PPiece(2).Y = .PCenter.Y + 1
                        .PPiece(3).Y = .PCenter.Y - 1
                    
                        .PPosition = 4
                    Case 4
                        .PPiece(1).X = .PCenter.X - 2
                        .PPiece(2).X = .PCenter.X - 1
                        .PPiece(3).X = .PCenter.X + 1
                        .PPiece(1).Y = .PCenter.Y
                        .PPiece(2).Y = .PCenter.Y
                        .PPiece(3).Y = .PCenter.Y
                        
                        .PPosition = 1
                End Select
        End Select
    End With
    
End Sub

Public Sub InstantDrop(ByRef GGrid() As GameBoard, ByRef APiece As GamePiece, ByRef NPiece As Boolean, _
                       ByRef DConsts As GridProperties)

    Dim TempPiece As GamePiece
    Dim MotionStopped As Boolean
    Dim K As Integer
    
    MotionStopped = False
    
    With APiece
        Do
            .PCenter.Y = .PCenter.Y + 1
            
            For K = 1 To 3
                .PPiece(K).Y = .PPiece(K).Y + 1
            Next K
            
            CheckBottom APiece, MotionStopped, DConsts
            
            If Not MotionStopped Then CheckBelowPiece GGrid(), APiece, MotionStopped, DConsts.CellColor
        Loop While MotionStopped = False
    End With
    
    NPiece = True
    
End Sub

