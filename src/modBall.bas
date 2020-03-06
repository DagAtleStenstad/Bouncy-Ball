Attribute VB_Name = "modBall"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds as Long) 'For 32 Bit Systems
#End If

Public Const gameHeight As Integer = 16
Public Const gameWidth As Integer = 15
Private gameRunning As Boolean
Private balls As New Collection

Sub StartBouncer()

Dim ball As New clsBall
balls.Add ball

gameRunning = True

Do While gameRunning
    DoEvents
    
    For Each ball In balls
        ball.drawBall (True)
    Next
    
    Sleep (100)
Loop

clearGameScreen

Set balls = Nothing

End Sub

Sub StopBouncer()

gameRunning = False

End Sub

Sub clearGameScreen()

Dim y, x As Integer

For y = 1 To gameHeight
    For x = 1 To gameWidth
        Cells(y, x).Interior.ColorIndex = 0
    Next x
Next y

End Sub
