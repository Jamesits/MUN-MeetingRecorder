Attribute VB_Name = "Common"
Option Explicit
Public Const minVer = 0
Public Const difSpeed = 50


Public Function getVersionString()
getVersionString = App.Major & "." & App.Minor & "." & minVer & " " & "Build " & Format$(App.Revision, "0000")
End Function

Public Sub Main()
If App.PrevInstance Then End
Load frmPreLoad
End Sub


Public Sub MovePort(ByRef Control, ByVal speed As Integer, ByVal x As Integer)
Dim i As Integer
If x < Control.Left Then speed = -speed
For i = Control.Left To x Step speed
Control.Left = i
DoEvents
Next i
End Sub
