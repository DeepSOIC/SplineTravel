Attribute VB_Name = "mdlCommon"
Option Explicit

Public Const Pi As Double = 3.14159265358979



Public Type typCurrentState
  Speed As Double
  pos As typVector3D
  Epos As Double 'extrusion accumulator
  MoveRelative As Boolean
  ExtrusionRelative As Boolean
End Type

Function vtStr(val As Double) As String
vtStr = Trim(Str(val))
End Function




