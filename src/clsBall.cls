VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pBallX As Integer
Private pBallY As Integer
Private pBallXSpeed As Integer
Private pBallySpeed As Integer
Private pPreviousBallX As Integer
Private pPreviousBallY As Integer
Private pBallColour As Integer

Public Property Get ballX() As Integer
    ballX = pBallX
End Property

Public Property Let ballX(value As Integer)
    pBallX = value
End Property

Public Property Get ballY() As Integer
    ballY = pBallY
End Property

Public Property Let ballY(value As Integer)
    pBallY = value
End Property

Public Property Get ballXSpeed() As Integer
    ballXSpeed = pBallXSpeed
End Property

Public Property Let ballXSpeed(value As Integer)
    pBallXSpeed = value
End Property

Public Property Get ballYSpeed() As Integer
    ballYSpeed = pBallySpeed
End Property

Public Property Let ballYSpeed(value As Integer)
    pBallySpeed = value
End Property

Public Property Get previousBallX() As Integer
    previousBallX = pPreviousBallX
End Property

Public Property Let previousBallX(value As Integer)
    pPreviousBallX = value
End Property

Public Property Get previousBallY() As Integer
    previousBallY = pPreviousBallY
End Property

Public Property Let previousBallY(value As Integer)
    pPreviousBallY = value
End Property

Public Property Get ballColour() As Integer
    ballColour = pBallColour
End Property

Public Property Let ballColour(value As Integer)
    pBallColour = value
End Property

Private Sub Class_Initialize()
    Randomize
    
    ballX = Int((gameWidth * Rnd) + 1)
    ballY = Int((gameHeight * Rnd) + 1)
    ballXSpeed = Int((2 * Rnd) + 1)
    ballYSpeed = Int((2 * Rnd) + 1)
    ballColour = Int((56 * Rnd) + 1)
End Sub

Public Sub Initialise()

End Sub

Private Sub updateBallPossition()

previousBallY = ballY
previousBallX = ballX

ballX = ballX + ballXSpeed
ballY = ballY + ballYSpeed
 
If ballX > gameWidth Then
    ballX = gameWidth
    ballXSpeed = ballXSpeed * -1
End If
If ballX < 1 Then
    ballX = 1
    ballXSpeed = ballXSpeed * -1
End If
 
If ballY > gameHeight Then
    ballY = gameHeight
    ballYSpeed = ballYSpeed * -1
End If
If ballY < 1 Then
    ballY = 1
    ballYSpeed = ballYSpeed * -1
End If

End Sub

Public Sub drawBall(removeShadow As Boolean)

Cells(ballY, ballX).Interior.ColorIndex = ballColour

If (removeShadow And (previousBallY And previousBallX <> 0)) Then Cells(previousBallY, previousBallX).Interior.ColorIndex = 0

updateBallPossition

End Sub
