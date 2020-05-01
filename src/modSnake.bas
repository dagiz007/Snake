Attribute VB_Name = "modSnake"
Option Explicit
Public Game As clsGame

Sub StartSnake()
    
    If Not Game Is Nothing Then Exit Sub
    
    Application.OnKey "{RIGHT}", ""
    Application.OnKey "{LEFT}", ""
    Application.OnKey "{UP}", ""
    Application.OnKey "{DOWN}", ""
    Application.OnKey "{ESC}", ""
    Application.OnKey "{TAB}", ""
      
    Application.Cursor = xlNorthwestArrow
    
    Set Game = New clsGame

    Game.Start
    
    Set Game = Nothing

    Application.Cursor = xlDefault
    
    Application.OnKey "{RIGHT}"
    Application.OnKey "{LEFT}"
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
    Application.OnKey "{ESC}"
    Application.OnKey "{TAB}"

End Sub

Sub PauseSnake()

    Game.PauseGame

End Sub

