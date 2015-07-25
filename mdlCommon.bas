Attribute VB_Name = "mdlCommon"
Option Explicit

Public Const Pi As Double = 3.14159265358979

Public Type typCurrentState
  Speed As Double
  Epos As Double 'extrusion accumulator
  MoveRelative As Boolean
  ExtrusionRelative As Boolean
End Type

Public posDecimals As Integer
Public extrDecimals As Integer

Function vtStr(val As Double) As String
vtStr = Trim(Str(val))
End Function
