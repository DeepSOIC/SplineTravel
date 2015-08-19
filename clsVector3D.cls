VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVector3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public X As Double
Public Y As Double
Public Z As Double

Public Function Length() As Double
Length = Sqr(X ^ 2 + Y ^ 2 + Z ^ 2)
End Function

Public Function Added(right As clsVector3D) As clsVector3D
Dim ret As New clsVector3D
ret.X = Me.X + right.X
ret.Y = Me.Y + right.Y
ret.Z = Me.Z + right.Z
Set Added = ret
End Function

Public Function Subtrd(right As clsVector3D) As clsVector3D
Dim ret As New clsVector3D
ret.X = Me.X - right.X
ret.Y = Me.Y - right.Y
ret.Z = Me.Z - right.Z
Set Subtrd = ret
End Function

Friend Function SubtrdT(right As typVector3D) As clsVector3D
Dim ret As New clsVector3D
ret.X = Me.X - right.X
ret.Y = Me.Y - right.Y
ret.Z = Me.Z - right.Z
Set SubtrdT = ret
End Function

Public Function Multed(multiplier As Double) As clsVector3D
Dim ret As New clsVector3D
ret.X = Me.X * multiplier
ret.Y = Me.Y * multiplier
ret.Z = Me.Z * multiplier
Set Multed = ret
End Function

Public Sub Mult(multiplier As Double)
Me.X = Me.X * multiplier
Me.Y = Me.Y * multiplier
Me.Z = Me.Z * multiplier
End Sub

Public Function Dot(other As clsVector3D) As Double
Dot = Me.X * other.X + Me.Y * other.Y + Me.Z * other.Z
End Function

Public Function Normalized() As clsVector3D
Dim l As Double
l = Me.Length
If l > 1E-100 Then
  Set Normalized = Me.Multed(1 / Me.Length)
Else
  Dim ret As New clsVector3D
  ret.X = 1
  ret.Y = 0
  ret.Z = 0
  Set Normalized = ret
End If
End Function

Public Sub Round()
Me.X = VBA.Round(Me.X, mdlPrecision.posDecimals)
Me.Y = VBA.Round(Me.Y, mdlPrecision.posDecimals)
Me.Z = VBA.Round(Me.Z, mdlPrecision.posDecimals)
End Sub

Public Function Copy() As clsVector3D
Dim ret As New clsVector3D
ret.copyFrom Me
Set Copy = ret
End Function

Public Sub copyFrom(vecFrom As clsVector3D)
Me.X = vecFrom.X
Me.Y = vecFrom.Y
Me.Z = vecFrom.Z
End Sub

Friend Sub copyFromT(vecFrom As typVector3D)
Me.X = vecFrom.X
Me.Y = vecFrom.Y
Me.Z = vecFrom.Z
End Sub

Friend Function asTypVector3D() As typVector3D
asTypVector3D.X = Me.X
asTypVector3D.Y = Me.Y
asTypVector3D.Z = Me.Z
End Function

Public Function repr() As String
Attribute repr.VB_UserMemId = 0
repr = "(" + Str(X) + ", " + Str(Y) + ", " + Str(Z) + ")"
End Function
