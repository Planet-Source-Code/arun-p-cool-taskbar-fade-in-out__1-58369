Attribute VB_Name = "Module1"
'requires XP
'vb 6

'if u like please vote for me
'please do mail me
'arun_pbk@rediffmail.com

'this is my first submission

'''''''''''''from planet-source-code.com
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_EXSTYLE = -20
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_STYLE = (-16)
Public Const WS_VISIBLE = &H10000000
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Const LWA_ALPHA = &H2
'''''''''''

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Type POINTAPI
    X As Long
    Y As Long
End Type
    

Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

