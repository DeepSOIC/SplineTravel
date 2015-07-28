VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public p1 As clsVector3D
Public p2 As clsVector3D
Public time As Double 'mm/s
Public Extrusion As Double 'in raw filament length, or whatever was printed to the G-code being processed. Positive for extrusion, negative for retraction.

Public PreviousMoveRef As clsGMove


Public Property Get traveldist() As Double
traveldist = p1.Subtrd(p2).Length
End Property

Public Property Get Speed() As Double
If Me.time > 1E-100 Then
  Speed = Me.traveldist / Me.time
Else
  Throw errZeroTimeMove, extraMessage:="can't get the speed"
End If
End Property

'changes the time, and consequently, extrusion speed
Public Property Let Speed(ByVal newSpeed As Double)
If Me.traveldist > 1E-100 Then
  Me.time = Me.traveldist / newSpeed
Else
  Throw errZeroTimeMove, extraMessage:="Can't set the speed for zero-distance move"
End If
End Property

Public Function isValid(Optional ByVal throwIfInvalid As Boolean) As Boolean
Dim ret As Boolean: ret = True
If time < 1E-100 Then
  ret = False
  If throwIfInvalid Then Throw errZeroTimeMove, extraMessage:="time is zero or negative, move is invalid"
End If
'for now, assume no move to be a valid move, it is a pause
'If Me.TravelDist < 1E-100 And Me.Extrusion < 1E-100 Then
'  ret = False
'
'End If
End Function

Function isPause() As Boolean
isValid throwIfInvalid:=True
isPause = Me.traveldist < 1E-100 And Abs(Me.Extrusion) < 1E-100
End Function

'EError is an error tracker for relative extrusion output. It outputs the error caused by rounding, as well as is an input (will be added to the extrusion).
Friend Function GenerateGCode(CurrentState As typCurrentState, Optional ByRef EError As Double) As String
Dim outStr As String
Me.isValid throwIfInvalid:=True
If Not Me.isPause Then
  'move
  
  'write position change
  Dim D As clsVector3D 'move vector (relative)
  Set D = Me.p2.Subtrd(Me.p1)
  Dim wr As clsVector3D 'coordinates to write
  If CurrentState.MoveRelative Then
    Set wr = D
  Else
    Set wr = Me.p2
  End If
  wr.Round
  If Abs(D.X) > 1E-100 Then
    outStr = outStr + " X" + vtStr(wr.X)
  End If
  If Abs(D.Y) > 1E-100 Then
    outStr = outStr + " Y" + vtStr(wr.Y)
  End If
  If Abs(D.Z) > 1E-100 Then
    outStr = outStr + " Z" + vtStr(wr.Z)
  End If
  
  If time > 0 Then
    If Abs(Me.Speed - CurrentState.Speed) > 0.001 Then
      outStr = outStr + " F" + vtStr(Round(Me.Speed * 60, mdlCommon.speedDecimals))
    End If
  End If
  'write extrusion change
  If Abs(Me.Extrusion) > 1E-100 Then
    Dim wrE As Double
    If CurrentState.ExtrusionRelative Then
      wrE = Me.Extrusion + EError
      EError = wrE - Round(wrE, mdlCommon.extrDecimals)
    Else
      'absolute E values
      wrE = CurrentState.Epos + Me.Extrusion
    End If
    wrE = Round(wrE, mdlCommon.extrDecimals)
    outStr = outStr + " E" + vtStr(wrE)
  End If
  
  If Len(outStr) > 0 Then
    GenerateGCode = "G1" + outStr
  End If
  
Else
  'pause
  GenerateGCode = "G4 P" + vtStr(Round(Me.time * 1000))
End If
End Function

Public Property Get ExtrusionSpeed() As Double
If time > 1E-100 Then
  ExtrusionSpeed = Abs(Extrusion) / time
Else
  Throw errZeroTimeMove, extraMessage:="can't get the speed"
End If
End Property

'note: this will actually change time of the move, so it will change move speed too
Public Property Let ExtrusionSpeed(ByVal newVal As Double)
If Abs(Extrusion) Then
  time = Extrusion / newVal
Else
  Throw errZeroTimeMove, extraMessage:="can't set the speed for zero-length extrusion"
End If
End Property

Private Sub Class_Initialize()
Set Me.p1 = New clsVector3D
Set Me.p2 = New clsVector3D
End Sub
