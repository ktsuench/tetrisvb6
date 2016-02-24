Attribute VB_Name = "mdlDeclarations"
Public Type PointRec
    X As Integer
    Y As Integer
End Type

Public Type GameBoard
    GColor As ColorConstants
End Type

Public Type GamePiece
    PShape As Integer
    PColor As ColorConstants
    PPosition As Integer
    PCenter As PointRec
    PPiece(1 To 3) As PointRec
End Type

Public Type GameScore
    Level As Integer
    LinesToClear As Integer
    NumLines As Integer
    Score As Integer
End Type

Public Type HighScores
    PlayerName As String * 15
    Score As Long
End Type

Public Type Pieces
    Z As Integer
    S As Integer
    T As Integer
    O As Integer
    L As Integer
    J As Integer
    I As Integer
End Type

Public Type GridProperties
    MaxX As Integer
    MaxY As Integer
    CellColor As ColorConstants
    CellSize As Integer
    LineColor As ColorConstants
    LineSize As Integer
End Type

Option Explicit
