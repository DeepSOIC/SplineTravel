Attribute VB_Name = "mdlCommon"
Option Explicit

Public Const Pi As Double = 3.14159265358979


Public Type typVector3D
  X As Double
  Y As Double
  Z As Double
End Type

Public Type typCurrentState
  Speed As Double
  pos As typVector3D
  Epos As Double 'extrusion accumulator
  MoveRelative As Boolean
  ExtrusionRelative As Boolean
End Type

Public posDecimals As Integer
Public extrDecimals As Integer
Public speedDecimals As Integer

Function vtStr(val As Double) As String
vtStr = Trim(Str(val))
End Function


'some vector-type routines, for faster processing without creating vectors
Public Function dist(point1 As typVector3D, point2 As typVector3D) As Double
dist = (point2.X - point1.X) ^ 2 + (point2.Y - point1.Y) ^ 2 + (point2.Z - point1.Z) ^ 2
End Function

Public Function vectorLength(vec As typVector3D) As Double
vectorLength = vec.X ^ 2 + vec.Y ^ 2 + vec.Z ^ 2
End Function

Public Function makeClsVector(vec As typVector3D) As clsVector3D
Dim v As New clsVector3D
v.copyFromT vec
Set makeClsVector = v
End Function

Public Function diff(vec1 As typVector3D, vec2 As typVector3D) As typVector3D
Dim ret As typVector3D
ret.X = vec2.X - vec1.X
ret.Y = vec2.Y - vec1.Y
ret.Z = vec2.Z - vec1.Z
diff = ret
End Function
