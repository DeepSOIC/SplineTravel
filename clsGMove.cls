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
  Set D = Me.p2.SubtrdT(CurrentState.pos) 'may differ from p1 because of rounding errors
  If Abs(D.X) > mdlPrecision.posConfusion Then
    outStr = outStr + " X" + vtStr(wr.X)
  End If
  If Abs(D.Y) > mdlPrecision.posConfusion Then
    outStr = outStr + " Y" + vtStr(wr.Y)
  End If
  If Abs(D.Z) > mdlPrecision.posConfusion Then
    outStr = outStr + " Z" + vtStr(wr.Z)
  End If
  
  If time > 0 Then
    If Abs(Me.FeedRate - CurrentState.Speed) > mdlPrecision.speedConfusion Then
      outStr = outStr + " F" + vtStr(Round(Me.FeedRate * 60, mdlPrecision.speedDecimals))
    End If
  End If
  'write extrusion change
  If Abs(Me.Extrusion) > mdlPrecision.extrConfusion Then
    Dim wrE As Double
    If CurrentState.ExtrusionRelative Then
      wrE = Me.Extrusion + EError
      EError = wrE - Round(wrE, mdlPrecision.extrDecimals)
    Else
      'absolute E values
      wrE = CurrentState.Epos + Me.Extrusion
    End If
    wrE = Round(wrE, mdlPrecision.extrDecimals)
    outStr = outStr + " E" + vtStr(wrE)
  End If
  
  If Len(outStr) > 0 Then
    GenerateGCode = "G1" + outStr
  End If
  
Else
  'pause
  GenerateGCode = "G4 P" + vtStr(Round(Me.time * 1000, Decimals:=0))
End If
End Function

Public Property Get FeedRate() As Double
If time > 1E-100 Then
  If Me.traveldist > 1E-100 Then
    FeedRate = Me.traveldist / time
  Else
    FeedRate = Abs(Extrusion) / time
  End If
Else
  Throw errZeroTimeMove, extraMessage:="can't get the speed"
End If
End Property

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
  time = Abs(Extrusion) / newVal
Else
  Throw errZeroTimeMove, extraMessage:="can't set the speed for zero-length extrusion"
End If
End Property

Private Sub Class_Initialize()
Set Me.p1 = New clsVector3D
Set Me.p2 = New clsVector3D
End Sub

Public Function split(ByVal timePoint As Double, ByRef Part1 As clsGMove, ByRef Part2 As clsGMove) As Boolean
If timePoint <= 1E-100 Then
  Set Part1 = Me
  split = False
ElseIf timePoint >= Me.time - 1E-100 Then
  Set Part2 = Me
  split = False
Else
  Set Part1 = New clsGMove
  Set Part2 = New clsGMove
  Part1.p1.copyFrom Me.p1
  Dim s As Double, t As Double
  t = timePoint / Me.time
  s = 1 - t
  Part1.p2.copyFrom Me.p1.Multed(s).Added(Me.p2.Multed(t))
  Part1.Extrusion = t * Me.Extrusion
  Part1.time = t * Me.time
  Part2.p1.copyFrom Part1.p2
  Part2.p2.copyFrom Me.p2
  Part2.Extrusion = s * Me.Extrusion
  Part2.time = s * Me.time
  split = True
End If
End Function
