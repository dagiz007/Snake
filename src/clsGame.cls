VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public State As GameState
Public Score As Integer

Private Snake As clsSnake

Public Sub IncreaseScore(Optional ScoreValue As Integer = 1)
    Score = Score + ScoreValue
    ActiveSheet.Shapes("txtScore").TextFrame.Characters.Text = Score
End Sub

Private Sub Class_Initialize()
    Set Snake = New clsSnake
    IncreaseScore (0)
End Sub

Private Sub Class_Terminate()
    Set Snake = Nothing
End Sub

Public Sub Start()
    
    Dim i As Integer
    
    State = gsPlaying
    ClearGameboard
    AddFood
    
    Set Snake = New clsSnake
    Snake.Createsnake
    
    Do While Game.State = gsPlaying
    
        DoEvents
         
        For i = 1 To 4
            GetKeyboardInput
            Sleep ([SnakeSpeed] / 4)
        Next i
             
        Snake.Move
        
    Loop
    
    MsgBox "Game over!"
        
End Sub

Public Sub StopGame()
    Game.State = gsStopped
End Sub

Public Sub PauseGame()

    If Game.State = gsPlaying Then
        Game.State = gsPaused
    ElseIf Game.State = gsPaused Then
        Game.State = gsPlaying
    End If
    
     setPauseModus
            
End Sub

Private Sub setPauseModus()
    
    Do While Game.State = gsPaused
        DoEvents
        GetKeyboardInput
    Loop

End Sub

Public Sub AddFood()

    Dim fruitPosition As New clsPosition

    Randomize

    Do
        fruitPosition.Y = Int([GameboardStartPossitionY] + Rnd * ([GameboardHeight] - [GameboardStartPossitionY] + 1))
        fruitPosition.X = Int([GameboardStartPossitionX] + Rnd * ([GameboardWidth] - [GameboardStartPossitionX] + 1))
    Loop While getGameObject(fruitPosition) <> "Background"
    
    Call setGameObject(fruitPosition, "Fruit")
    
End Sub

Public Function getGameObject(pPosition As clsPosition) As String
    getGameObject = Cells(pPosition.Y, pPosition.X).style
End Function

Public Sub setGameObject(pPosition As clsPosition, gameObject As String)
    Cells(pPosition.Y, pPosition.X).style = gameObject
End Sub

Public Sub ClearGameboard()

    Range(Cells([GameboardStartPossitionY], [GameboardStartPossitionX]), _
          Cells([GameboardStartPossitionY] + [GameboardHeight], [GameboardStartPossitionX] + [GameboardWidth])).style = "Background"
  
End Sub

Private Sub GetKeyboardInput()
    
    If GetAsyncKeyState(VK_ESCAPE) <> 0 Then
        Game.StopGame
    
    ElseIf GetAsyncKeyState(VK_SPACE) <> 0 Then
        Game.PauseGame
        
    ElseIf GetAsyncKeyState(VK_LEFT) <> 0 Then
        Snake.Direction = "MoveLeft"
    
    ElseIf GetAsyncKeyState(VK_RIGHT) <> 0 Then
        Snake.Direction = "MoveRight"
    
    ElseIf GetAsyncKeyState(VK_UP) <> 0 Then
        Snake.Direction = "MoveUp"
    
    ElseIf GetAsyncKeyState(VK_DOWN) <> 0 Then
        Snake.Direction = "MoveDown"
    
    End If
    
End Sub
