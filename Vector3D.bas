Attribute VB_Name = "Vector3D"
Option Explicit

Public Type typVector3D
  X As Double
  Y As Double
  Z As Double
End Type

'some vector-type routines, for faster processing without creating vectors
Public Function Dist(point1 As typVector3D, point2 As typVector3D) As Double
Dist = Sqr((point2.X - point1.X) ^ 2 + (point2.Y - point1.Y) ^ 2 + (point2.Z - point1.Z) ^ 2)
End Function

Public Function Length(vec As typVector3D) As Double
Length = Sqr(vec.X ^ 2 + vec.Y ^ 2 + vec.Z ^ 2)
End Function

Public Function makeClsVector(vec As typVector3D) As clsVector3D
Dim v As New clsVector3D
v.copyFromT vec
Set makeClsVector = v
End Function

Public Function Subtracted(vec1 As typVector3D, vec2 As typVector3D) As typVector3D
Dim ret As typVector3D
ret.X = vec1.X - vec2.X
ret.Y = vec1.Y - vec2.Y
ret.Z = vec1.Z - vec2.Z
Subtracted = ret
End Function

Public Function Multed(vec As typVector3D, ByVal Multediplier As Double) As typVector3D
Multed.X = vec.X * Multediplier
Multed.Y = vec.Y * Multediplier
Multed.Z = vec.Z * Multediplier
End Function

Public Function Combi2( _
          vec1 As typVector3D, ByVal multiplier1 As Double, _
          vec2 As typVector3D, ByVal multiplier2 As Double _
          ) As typVector3D
Combi2.X = vec1.X * multiplier1 + vec2.X * multiplier2
Combi2.Y = vec1.Y * multiplier1 + vec2.Y * multiplier2
Combi2.Z = vec1.Z * multiplier1 + vec2.Z * multiplier2
End Function

Public Function Combi3( _
          vec1 As typVector3D, ByVal multiplier1 As Double, _
          vec2 As typVector3D, ByVal multiplier2 As Double, _
          vec3 As typVector3D, ByVal multiplier3 As Double _
          ) As typVector3D
Combi3.X = vec1.X * multiplier1 + vec2.X * multiplier2 + vec3.X * multiplier3
Combi3.Y = vec1.Y * multiplier1 + vec2.Y * multiplier2 + vec3.Y * multiplier3
Combi3.Z = vec1.Z * multiplier1 + vec2.Z * multiplier2 + vec3.Z * multiplier3
End Function


Public Function Combi4( _
          vec1 As typVector3D, ByVal multiplier1 As Double, _
          vec2 As typVector3D, ByVal multiplier2 As Double, _
          vec3 As typVector3D, ByVal multiplier3 As Double, _
          vec4 As typVector3D, ByVal multiplier4 As Double _
          ) As typVector3D
Combi4.X = vec1.X * multiplier1 + vec2.X * multiplier2 + vec3.X * multiplier3 + vec4.X * multiplier4
Combi4.Y = vec1.Y * multiplier1 + vec2.Y * multiplier2 + vec3.Y * multiplier3 + vec4.Y * multiplier4
Combi4.Z = vec1.Z * multiplier1 + vec2.Z * multiplier2 + vec3.Z * multiplier3 + vec4.Z * multiplier4
End Function

Public Function Normalized(vec As typVector3D) As typVector3D
Normalized = Multed(vec, 1 / Length(vec))
End Function

Public Function Dot(vec1 As typVector3D, vec2 As typVector3D) As Double
Dot = vec1.X * vec2.X + vec1.Y * vec2.Y + vec1.Z * vec2.Z
End Function
