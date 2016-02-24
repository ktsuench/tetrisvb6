Attribute VB_Name = "mdlGeneral"
Option Explicit

'This procedure is used to slow down events that are being executed.

Public Sub Delay(ByVal Interval As Single)

    Dim Start, Finish
    
    Start = Timer
    
    Do
        Finish = Timer
        DoEvents
    Loop While Finish - Start <= Interval
    
End Sub

'This function is used as a general dialog prompt to confirm with the user that they want to continue
'   their current course of action, by default it is exiting the program.

Public Function PromptDialog(Optional ByVal DMsg As String = "Are you sure you want to exit?", Optional _
                             ByVal DTitle As String = "Exit?") As Boolean
    
    Dim DType As Integer, Response As Integer
    Dim Answer As Boolean
    
    Answer = False

    DType = vbYesNo + vbQuestion
    
    Response = MsgBox(DMsg, DType, DTitle)
    
    If Response = vbYes Then Answer = True
    
    PromptDialog = Answer
    
End Function

'This procedure centers an object in a form or a form shown on the screen.

Public Sub CenterObject(ByRef objOut As Object, ByVal isForm As Boolean, Optional ByRef frmOut As Form)

    If isForm Then
        objOut.Left = (Screen.Width - objOut.Width) \ 2
        objOut.Top = (Screen.Height - objOut.Height) \ 2
    Else
        objOut.Left = (frmOut.ScaleWidth - objOut.Width) \ 2
        objOut.Top = (frmOut.ScaleHeight - objOut.Height) \ 2
    End If
    
End Sub

'This procedure centers an object horizontally in a form or a form shown on
'   the screen.

Public Sub CenterHorizontalObj(ByRef objOut As Object, ByVal isForm As _
                               Boolean, Optional ByRef frmOut As Form)

    If isForm Then
        objOut.Left = (Screen.Width - objOut.Width) \ 2
    Else
        objOut.Left = (frmOut.ScaleWidth - objOut.Width) \ 2
    End If
    
End Sub

'This procedure is used to display a status when the user activates specific events, such as game over.

Public Sub DisplayStatusMsg(ByVal Msg As String, ByVal FinalOutput As String, ByRef ctrlFrame As Frame, _
                            ByRef ctrlLabel As Label)

    Dim K As Integer
    Dim Char As String
    
    K = 0
    Char = ""
    
    ctrlFrame.Visible = True
    
    Do
        K = K + 1
        
        Char = Mid$(Msg, K, 1)
        
        If Char = " " Then Char = vbCrLf
        
        ctrlLabel.Caption = ctrlLabel.Caption & Char
        
        Delay 0.1
    Loop While K < Len(Msg) And ctrlLabel.Caption <> FinalOutput
End Sub

'This procedure is used to insert random numbers into an array without any repeated numbers.

Public Sub RandomizeArray(ByRef n() As Integer, ByVal M As Integer, ByVal High As Integer, Optional _
                          ByVal Low As Integer = 1)
    
    Dim K As Integer
    Dim IsRepeat As Boolean
    
    Randomize
    
    For K = 1 To M
        n(K) = Int(Rnd * (High - Low + 1)) + Low
        
        IsRepeat = IsRepeatNum(n(), K)
        If IsRepeat = True Then K = K - 1
    Next K
    
End Sub

'This function is used to validate that there are no repeated numbers in an array.

Public Function IsRepeatNum(ByRef n() As Integer, ByVal M As Integer) As Boolean

    Dim K As Integer
    Dim IsRepeat As Boolean
    
    IsRepeat = False
    
    Do
        K = K + 1
        If M <> K Then If n(M) = n(K) Then IsRepeat = True
    Loop While K < M - 1 And Not IsRepeat

    IsRepeatNum = IsRepeat
    
End Function

'This procedure opens up the highscore file.

Public Sub OpenHSFile(ByVal FileLoc As String, ByRef Scores() As HighScores)

On Error GoTo ErrorHandler:

    Dim K As Integer
    
    K = 1
    
    Open FileLoc For Random As #1
        Do While Not EOF(1)
            Get #1, K, Scores(K)
            K = K + 1
        Loop
    Close #1

    Exit Sub
    
ErrorHandler:
    Open FileLoc For Random As #1
        For K = 1 To 10
            Scores(K).PlayerName = "Player Name"
            Scores(K).Score = 0
            Put #1, K, Scores(K)
        Next K
    Close #1
    
End Sub

'This procedure writes to highscore file.

Public Sub WriteHSFile(ByVal FileLoc As String, ByRef Scores() As HighScores)

    Dim K As Integer
    
    Open FileLoc For Random As #1
        For K = 1 To UBound(Scores)
            Put #1, K, Scores(K)
        Next K
    Close #1

End Sub
