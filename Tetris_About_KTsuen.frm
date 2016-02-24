VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tetris - About"
   ClientHeight    =   5700
   ClientLeft      =   11445
   ClientTop       =   4935
   ClientWidth     =   8535
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Tetris_About_KTsuen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8535
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   4680
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2700
      Left            =   480
      Picture         =   "Tetris_About_KTsuen.frx":058A
      ScaleHeight     =   2700
      ScaleWidth      =   7575
      TabIndex        =   0
      Top             =   360
      Width           =   7575
   End
   Begin VB.Label lblCopy 
      BackColor       =   &H00000000&
      Caption         =   "Copyright 2014"
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
      Left            =   480
      TabIndex        =   2
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label lblAuthor 
      BackColor       =   &H00000000&
      Caption         =   "By: Kent Tsuen 12 L"
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
      Left            =   480
      TabIndex        =   1
      Top             =   3240
      Width           =   4215
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()
    
    CenterObject Me, True
    CenterHorizontalObj lblAuthor, False, Me
    CenterHorizontalObj lblCopy, False, Me
    CenterHorizontalObj cmdBack, False, Me
    
End Sub
