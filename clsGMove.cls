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
Public Time As Double 'mm/s
Public Extrusion As Double 'in raw filament length, or whatever was printed to the G-code being processed. Positive for extrusion, negative for retraction.

Public PreviousMoveRef As clsGMove


Public Property Get TravelDist() As Double
TravelDist = p1.Subtrd(p2).Length
End Property

Public Property Get Speed() As Double
If Me.Time > 1E-100 Then
  Speed = Me.TravelDist / Me.Time
Else
  Throw errZeroTimeMove, extraMessage:="can't get the speed"
End If
End Property

'changes the time, and consequently, extrusion speed
Public Property Let Speed(ByVal newSpeed As Double)
If Me.TravelDist > 1E-100 Then
  Me.Time = Me.TravelDist / newSpeed
Else
  Throw errZeroTimeMove, extraMessage:="Can't set the speed for zero-distance move"
End If
End Property

Public Function isValid(Optional ByVal ThrowIfInvalid As Boolean) As Boolean
Dim ret As Boolean: ret = True
If Time < 1E-100 Then
  ret = False
  If ThrowIfInvalid Then Throw errZeroTimeMove, extraMessage:="time is zero or negative, move is invalid"
End If
'for now, assume no move to be a valid move, it is a pause
'If Me.TravelDist < 1E-100 And Me.Extrusion < 1E-100 Then
'  ret = False
'
'End If
End Function

Function isPause() As Boolean
isValid ThrowIfInvalid:=True
isPause = Me.TravelDist < 1E-100 And Abs(Me.Extrusion) < 1E-100
End Function


Friend Function GenerateGCode(CurrentState As typCurrentState) As String
Dim outStr As String
Me.isValid ThrowIfInvalid:=True
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
  
  'write extrusion change
  If Me.Extrusion > 1E-100 Then
    Dim wrE As Double
    If CurrentState.ExtrusionRelative Then
      wrE = Me.Extrusion
    Else
      'absolute E values
      wrE = CurrentState.Epos + Me.Extrusion
    End If
    wrE = Round(wrE, mdlCommon.extrDecimals)
    outStr = outStr + " E" + vtStr(wrE)
  End If
  
  If Len(outStr) > 0 Then
    GenerateGCode = "G1" + outStr + vbNewLine
  End If
  
Else
  'pause
  GenerateGCode = "G4 P" + vtStr(Round(Me.Time * 1000)) + vbNewLine
End If
End Function

Public Property Get ExtrusionSpeed() As Double
If Time > 1E-100 Then
  ExtrusionSpeed = Abs(Extrusion) / Time
Else
  Throw errZeroTimeMove, extraMessage:="can't get the speed"
End If
End Property

'note: this will actually change time of the move, so it will change move speed too
Public Property Let ExtrusionSpeed(ByVal newVal As Double)
If Abs(Extrusion) Then
  Time = Extrusion / newVal
Else
  Throw errZeroTimeMove, extraMessage:="can't set the speed for zero-length extrusion"
End If
End Property

Private Sub Class_Initialize()
Set Me.p1 = New clsVector3D
Set Me.p2 = New clsVector3D
End Sub
