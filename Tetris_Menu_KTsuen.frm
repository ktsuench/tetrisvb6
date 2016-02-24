VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tetris"
   ClientHeight    =   5340
   ClientLeft      =   5280
   ClientTop       =   3720
   ClientWidth     =   9270
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Tetris_Menu_KTsuen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   9270
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   615
      Left            =   6240
      TabIndex        =   2
      Top             =   3360
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2700
      Left            =   840
      Picture         =   "Tetris_Menu_KTsuen.frx":058A
      ScaleHeight     =   2700
      ScaleWidth      =   7575
      TabIndex        =   5
      Top             =   360
      Width           =   7575
   End
   Begin VB.CommandButton cmdManual 
      Caption         =   "&Manual"
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdHighScores 
      Caption         =   "&High Scores"
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   5040
      TabIndex        =   4
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Height          =   615
      Left            =   1200
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   3360
      Width           =   1935
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbout_Click()

    frmAbout.Show vbModal

End Sub

Private Sub cmdExit_Click()

    End
    
End Sub

Private Sub cmdHighScores_Click()

    frmHighScores.Show vbModeless
    
End Sub

Private Sub cmdManual_Click()

    frmManual.Show vbModeless
    
End Sub

Private Sub cmdPlay_Click()

    frmGame.Show vbModeless
    Call frmGame.mnuPlay_Click
    
End Sub

Private Sub Form_Load()

    CenterObject Me, True
    
End Sub
