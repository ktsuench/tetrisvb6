VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHighScores 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tetris - High Scores"
   ClientHeight    =   5910
   ClientLeft      =   10410
   ClientTop       =   3030
   ClientWidth     =   4845
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Tetris_High_Scores_KTsuen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   4845
   Begin MSFlexGridLib.MSFlexGrid grdHighScores 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8705
      _Version        =   393216
      Rows            =   11
      FixedCols       =   0
      RowHeightMin    =   150
      BackColor       =   0
      ForeColor       =   16777215
      BackColorFixed  =   4210752
      ForeColorFixed  =   16777215
      BackColorSel    =   0
      ForeColorSel    =   0
      BackColorBkg    =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   0
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "High Scores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmHighScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    CenterObject Me, True
    
    Dim K As Integer, L As Integer
    Dim Scores(1 To 10) As HighScores
    
    OpenHSFile App.Path & "/highscore.hs", Scores()
    
    With grdHighScores
        'Stretch the width and height of the cells so they fit the table.
        .ColWidth(0) = .Width / 2
        .ColWidth(1) = .Width / 2
        For K = 0 To 10
            .RowHeight(K) = .Height / 11
        Next K
        
        'Fix the High Score Table display.
        .Width = .ColWidth(0) * 2 + 10
        .Height = .RowHeight(0) * 11 + 10
        
        'Titles of the High Score Table.
        .TextArray(0) = "Name"
        .TextArray(1) = "Score"
        
        'Player names and scores.
        L = 1
        
        For K = 2 To 21
            If K Mod 2 = 0 Then
                .TextArray(K) = Scores(L).PlayerName
            Else
                .TextArray(K) = Scores(L).Score
                L = L + 1
            End If
        Next K
    End With

End Sub
