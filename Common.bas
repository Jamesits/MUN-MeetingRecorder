Attribute VB_Name = "Common"
Option Explicit
Public Const minVer = 0



Public Function getVersionString()
getVersionString = App.Major & "." & App.Minor & "." & minVer & " " & "Build " & Format$(App.Revision, "0000")
End Function

Public Sub Main()
If App.PrevInstance Then End
Load frmPreLoad
End Sub
