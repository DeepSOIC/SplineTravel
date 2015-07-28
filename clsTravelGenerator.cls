VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTravelGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'move parameters
Public inSpeed As clsVector3D
Public outSpeed As clsVector3D
Public p1 As clsVector3D
Public p2 As clsVector3D

'settings
Public speedLimit As Double
Public ZJerk As Double 'Generally, this sets how fast Z is lifted
Public CurveJerk As Double 'The parameter that affects how good a curve is approximated with line segments
Public ZHopHeight As Double 'may be smaller than this value, this is simply the maximum
Public Retract As Double
Public RetractSpeed As Double
Public accelleration As Double

Public Function FitBezier(ByRef moveTimeResult As Double) As clsBezier
Const iterCount = 30
Dim bz As New clsBezier

Dim spd As Double
spd = maxD(Me.inSpeed.Length, Me.outSpeed.Length)
If spd < Me.CurveJerk Then
  'the spline is going to be a staight line, if both enter
  'speed and leave speed are very small.
  Throw errTooSlow
End If
'todo: test for all-same-direction (then, the spline will only slow down the move, so we shouldn't do it)

'currently, the bezier is obtained by a fitting iterative procedure:
'starting from an abrubt curve, it is smoothed out until the maximum
'accelleration is below me.accelleration. I think, it is possible to
'calculate everything analytically and speed up this process, but
'this is a todo.

Dim stopDist As Double 'the path length required to stop moving, we are going to use it for computing starting values
stopDist = Me.accelleration * (spd / Me.accelleration) ^ 2 / 2#

bz.Pole(0).copyFrom Me.p1
bz.Pole(3).copyFrom Me.p2

Dim Time As Double 'time it is to take to move through the whole spline
Dim timeOfSolved As Double

Time = stopDist / spd
Time = Time / 10 'start iteratons from very abrupt spline
Dim logFactor As Double
logFactor = 0.25 'sets, how fast to smooth out the spline (by increasing time) on each iteration
Dim i As Long
For i = 1 To iterCount
  bz.Pole(1).copyFrom Me.p1.Added(Me.inSpeed.Multed(Time / 3))
  bz.Pole(2).copyFrom Me.p2.Added(Me.outSpeed.Multed(-Time / 3))
  'test if the spline has enter and exit speeds correct (assert)
  Dim vin As clsVector3D
  Dim vout As clsVector3D
  bz.getDeriv 0, vin
  bz.getDeriv 1, vout
  vin.Mult 1 / Time
  vout.Mult 1 / Time
  Debug.Assert (vin.Subtrd(Me.inSpeed).Length < spd * 0.00000000001)
  Debug.Assert (vout.Subtrd(Me.outSpeed).Length < spd * 0.00000000001)
  
  'calculate max accelleration; test if this spline is smooth enough already
  Dim acc1 As clsVector3D
  Dim acc2 As clsVector3D
  bz.getDeriv2 0, acc1
  bz.getDeriv2 0, acc2
  acc1.Mult 1 / Time ^ 2
  acc2.Mult 1 / Time ^ 2
  If maxD(acc1.Length, acc2.Length) <= Me.accelleration Then
    'valid spline. roll back and try again with finer stepping
    timeOfSolved = Time
    Time = Time * Exp(-logFactor)
    logFactor = logFactor / 2
  Else
    'incease time, thus increasing smoothness
    Time = Time * Exp(logFactor)
  End If
Next i

'return results
moveTimeResult = timeOfSolved
Set FitBezier = bz
End Function


Private Function maxD(ByVal a As Double, ByVal b As Double) As Double
If a > b Then maxD = a Else maxD = b
End Function

Public Sub GenerateMoveTrainForBezier(ByRef arrMoves() As clsGMove, bz As clsBezier, ByVal TimeOfMove As Double)

'a few checks
If Me.CurveJerk <= 1E-100 Then Throw errClassNotInitialized, extraMessage:="CurveJerk value is zero, can't compute step"

Dim nMoves As Long
Dim nalloc As Long
Erase arrMoves
nalloc = 10
ReDim arrMoves(0 To nalloc - 1)
Dim curMove As clsGMove


Dim cur_t As Double
cur_t = 0
Do
  'compute current accelleration and use it to derive t-step
  Dim acc As clsVector3D
  Dim timestep As Double
  bz.getDeriv2 cur_t, acc
  acc.Mult 1 / TimeOfMove ^ 2
  timestep = Me.CurveJerk / acc.Length
  cur_t = cur_t + timestep / TimeOfMove
  If cur_t > 1 Then
    timestep = timestep - (cur_t - 1) * TimeOfMove
    cur_t = 1
  End If
  
  'compute new position, and generate the move
  'Dim pos As clsVector3D
  Dim prevPos As clsVector3D
  If prevPos Is Nothing Then
    Set prevPos = New clsVector3D
    prevPos.copyFrom bz.getPole(0)
  Else
    prevPos.copyFrom curMove.p2
  End If
  
  GoSub addMove
  curMove.p1.copyFrom prevPos
  bz.getValue cur_t, curMove.p2
  curMove.Time = timestep
  If curMove.Speed > Me.speedLimit Then curMove.Speed = speedLimit
Loop Until cur_t >= 1#

ReDim Preserve arrMoves(0 To nMoves - 1)

Exit Sub
'creates a new g-move, adds it to the array, and leaves it avalable for filling in curMove
addMove:
  Set curMove = New clsGMove
  nMoves = nMoves + 1
  If nMoves > nalloc Then
    nalloc = nalloc * 1.5
    ReDim Preserve arrMoves(0 To nalloc - 1)
  End If
  Set arrMoves(nMoves - 1) = curMove
Return
End Sub

Private Sub Class_Initialize()
Set Me.p1 = New clsVector3D
Set Me.p2 = New clsVector3D
Set Me.inSpeed = New clsVector3D
Set Me.outSpeed = New clsVector3D
End Sub
