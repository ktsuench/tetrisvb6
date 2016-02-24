Attribute VB_Name = "mdlCreatePiece"
Option Explicit

'This procedure determines that next active piece to appear.

Public Function DeterminePiece(ByRef PShape As Integer, ByRef PColor As ColorConstants, ByRef SConsts As _
                               Pieces)

    Dim PieceShape(1 To 7) As Integer
    Dim NextPiece As Integer
    
    Randomize
        
    RandomizeArray PieceShape, UBound(PieceShape), UBound(PieceShape)
    
    NextPiece = Int(Rnd * UBound(PieceShape)) + 1
    
    With SConsts
        Select Case NextPiece
            Case PieceShape(1)
                PShape = .Z
                PColor = RGB(255, 28, 28)   'Red
            Case PieceShape(2)
                PShape = .S
                PColor = RGB(28, 255, 28)   'Green
            Case PieceShape(3)
                PShape = .T
                PColor = RGB(128, 0, 255)   'Purple
            Case PieceShape(4)
                PShape = .O
                PColor = RGB(255, 255, 55)  'Yellow
            Case PieceShape(5)
                PShape = .L
                PColor = RGB(255, 255, 255) 'White
            Case PieceShape(6)
                PShape = .J
                PColor = RGB(255, 128, 0)   'Orange
            Case PieceShape(7)
                PShape = .I
                PColor = RGB(0, 128, 255)   'Blue
        End Select
    End With
    
End Function

'This procedure creates the piece once its shape has been determined.

Public Sub CreatePiece(ByRef APiece As GamePiece, ByRef SConsts As Pieces)
    
    With APiece
        Select Case .PShape
            Case SConsts.S, SConsts.T, SConsts.L
                .PCenter.X = 5
            Case Else
                .PCenter.X = 6
        End Select
    
        Select Case .PShape
            Case SConsts.Z, SConsts.S, SConsts.O
                .PCenter.Y = 2
            Case Else
                .PCenter.Y = 1
        End Select
        
        Select Case .PShape
            Case SConsts.Z, SConsts.S, SConsts.T
                .PPiece(1).X = .PCenter.X - 1
                .PPiece(2).X = .PCenter.X
                .PPiece(3).X = .PCenter.X + 1
                
                If .PShape = SConsts.Z Then
                    .PPiece(1).Y = .PCenter.Y - 1
                    .PPiece(2).Y = .PCenter.Y - 1
                    .PPiece(3).Y = .PCenter.Y
                ElseIf .PShape = SConsts.S Then
                    .PPiece(1).Y = .PCenter.Y
                    .PPiece(2).Y = .PCenter.Y - 1
                    .PPiece(3).Y = .PCenter.Y - 1
                Else
                    .PPiece(1).Y = .PCenter.Y
                    .PPiece(2).Y = .PCenter.Y + 1
                    .PPiece(3).Y = .PCenter.Y
                End If
            Case SConsts.O
                .PPiece(1).X = .PCenter.X - 1
                .PPiece(2).X = .PCenter.X
                .PPiece(3).X = .PCenter.X - 1
                .PPiece(1).Y = .PCenter.Y - 1
                .PPiece(2).Y = .PCenter.Y - 1
                .PPiece(3).Y = .PCenter.Y
            Case SConsts.L
                .PPiece(1).X = .PCenter.X - 1
                .PPiece(2).X = .PCenter.X + 1
                .PPiece(3).X = .PCenter.X - 1
                .PPiece(1).Y = .PCenter.Y
                .PPiece(2).Y = .PCenter.Y
                .PPiece(3).Y = .PCenter.Y + 1
            Case SConsts.J
                .PPiece(1).X = .PCenter.X - 1
                .PPiece(2).X = .PCenter.X + 1
                .PPiece(3).X = .PCenter.X + 1
                .PPiece(1).Y = .PCenter.Y
                .PPiece(2).Y = .PCenter.Y
                .PPiece(3).Y = .PCenter.Y + 1
            Case SConsts.I
                .PPiece(1).X = .PCenter.X - 2
                .PPiece(2).X = .PCenter.X - 1
                .PPiece(3).X = .PCenter.X + 1
                .PPiece(1).Y = .PCenter.Y
                .PPiece(2).Y = .PCenter.Y
                .PPiece(3).Y = .PCenter.Y
        End Select
    End With
    
    APiece.PPosition = 1
    
End Sub
