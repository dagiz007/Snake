VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSnake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Direction As String
Public Length As Integer

Private pBodypieces As Collection

Property Get HeadPosition() As clsPosition
    Set HeadPosition = pBodypieces(pBodypieces.Count)
End Property

Property Get TailPosition() As clsPosition
    Set TailPosition = pBodypieces(1)
End Property

Public Sub IncreaseLength()
    Me.Length = Me.Length + [SnakeGrow]
End Sub

Private Sub Class_Initialize()
   
    Direction = "MoveDown"
    Length = 1
    
End Sub

Public Sub Createsnake()
    
    Set pBodypieces = New Collection

    Dim initialHeadPosition As New clsPosition
    initialHeadPosition.X = [SnakeheadInitialPositionX]
    initialHeadPosition.Y = [SnakeheadInitialPositionY]

    pBodypieces.Add initialHeadPosition

    Call Game.setGameObject(HeadPosition, "Snake")
    
End Sub

Public Sub Move()
 
    Dim newHeadPosition As New clsPosition
    newHeadPosition.X = HeadPosition.X
    newHeadPosition.Y = HeadPosition.Y

    Select Case Direction
    
        Case "MoveLeft"
            newHeadPosition.X = newHeadPosition.X - 1
        
        Case "MoveRight"
            newHeadPosition.X = newHeadPosition.X + 1
        
        Case "MoveUp"
            newHeadPosition.Y = newHeadPosition.Y - 1
        
        Case "MoveDown"
            newHeadPosition.Y = newHeadPosition.Y + 1
    
    End Select
    
    pBodypieces.Add newHeadPosition
    
    
    Dim gameObject As String
    gameObject = Game.getGameObject(HeadPosition)

    Select Case gameObject
    
        Case "Background"
            'Nothing happens
            
        Case "Fruit"
            IncreaseLength
            Game.IncreaseScore
            Game.AddFood
        
        Case Else
            Game.StopGame
            Exit Sub
    
    End Select
    
    Call Game.setGameObject(HeadPosition, "Snake")
    
    If pBodypieces.Count - 1 = Length Then
        Call Game.setGameObject(TailPosition, "Background")
        pBodypieces.Remove 1
    End If

End Sub
