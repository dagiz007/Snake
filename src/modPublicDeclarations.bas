Attribute VB_Name = "modPublicDeclarations"
Option Explicit

#If Win64 Then
    'Code is running in 64-bit Office
    Public Declare PtrSafe Sub Sleep Lib "kernel32" ( _
        ByVal dwMilliseconds As Long)
   
    Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" ( _
        ByVal vKey As Long) As Integer
       
#Else
    'Code is running in 32-bit Office
    Public Declare Sub Sleep Lib "kernel32" ( _
        ByVal dwMilliseconds As Long)

    Public Declare Function GetAsyncKeyState Lib "user32" ( _
        ByVal vKey As Long) As Integer

#End If

Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28

Enum GameState
    gsStopped
    gsPlaying
    gsPaused
End Enum
