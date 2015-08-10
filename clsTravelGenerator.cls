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
Public bRetract As Boolean, bUnretract As Boolean
Public RetractJerk As Double
Public RetractAccelleration As Double
Public accelleration As Double


Public Function FitBezier(ByRef moveTimeResult As Double) As clsBezier
Const iterCount = 30
Dim bz As New clsBezier

Dim spd As Double
spd = maxD(inSpeed.Length, outSpeed.Length)
If spd < Me.CurveJerk Then
  'the spline is going to be a staight line, if both enter
  'speed and leave speed are very small.
  Throw errTooSlow
End If
'todo: test for all-same-direction (then, the spline will only slow down the move, so we shouldn't do it)

'add Z component to speeds, to create a hop
Dim inSpeed_hop As New clsVector3D
Dim outSpeed_hop As New clsVector3D
inSpeed_hop.copyFrom Me.inSpeed
outSpeed_hop.copyFrom Me.outSpeed
inSpeed_hop.Z = inSpeed_hop.Z + Me.ZJerk
outSpeed_hop.Z = outSpeed_hop.Z - Me.ZJerk

'currently, the bezier is obtained by a fitting iterative procedure:
'starting from an abrubt curve, it is smoothed out until the maximum
'accelleration is below me.accelleration. I think, it is possible to
'calculate everything analytically and speed up this process, but
'this is a todo.

Dim stopDist As Double 'the path length required to stop moving, we are going to use it for computing starting values
stopDist = Me.accelleration * (spd / Me.accelleration) ^ 2 / 2#

bz.Pole(0) = Me.p1.asTypVector3D
bz.Pole(3) = Me.p2.asTypVector3D

Dim time As Double 'time it is to take to move through the whole spline
Dim timeOfSolved As Double

time = stopDist / spd
time = time / 10 'start iteratons from very abrupt spline
Dim logFactor As Double
logFactor = 0.25 'sets, how fast to smooth out the spline (by increasing time) on each iteration
Dim i As Long
For i = 1 To iterCount
  bz.Pole(1) = Me.p1.Added(inSpeed_hop.Multed(time / 3)).asTypVector3D
  bz.Pole(2) = Me.p2.Added(outSpeed_hop.Multed(-time / 3)).asTypVector3D
  'test if the spline has enter and exit speeds correct (assert)
'  Dim vin As clsVector3D
'  Dim vout As clsVector3D
'  bz.getDeriv 0, vin
'  bz.getDeriv 1, vout
'  vin.Mult 1 / Time
'  vout.Mult 1 / Time
'  Debug.Assert (vin.Subtrd(Me.inSpeed_hop).Length < spd * 0.00000000001)
'  Debug.Assert (vout.Subtrd(Me.outSpeed_hop).Length < spd * 0.00000000001)
  
  'calculate max accelleration; test if this spline is smooth enough already
  Dim acc1 As clsVector3D
  Dim acc2 As clsVector3D
  Set acc1 = Vector3D.makeClsVector(bz.getDeriv2(0))
  Set acc2 = Vector3D.makeClsVector(bz.getDeriv2(0))
  acc1.Mult 1 / time ^ 2
  acc2.Mult 1 / time ^ 2
  If maxD(acc1.Length, acc2.Length) <= Me.accelleration Then
    'valid spline. roll back and try again with finer stepping
    timeOfSolved = time
    time = time * Exp(-logFactor)
    logFactor = logFactor / 2
  Else
    'incease time, thus increasing smoothness
    time = time * Exp(logFactor)
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

'generate retract
Dim rtr As New clsRetractCurve
rtr.RetractLen = Me.Retract
rtr.RetractA = Me.RetractAccelleration
rtr.MoveTime = TimeOfMove
rtr.bRetract = Me.bRetract
rtr.bUnretract = Me.bUnretract

Dim nMoves As Long
Dim nalloc As Long
Erase arrMoves
nalloc = 10
ReDim arrMoves(0 To nalloc - 1)
Dim curMove As clsGMove

bz.DerivJerk = Me.CurveJerk * TimeOfMove
rtr.DerivJerk = Me.RetractJerk * TimeOfMove

Dim cur_t As Double
Dim prev_t As Double
Dim prevPos As typVector3D
Dim curPos As typVector3D
curPos = bz.getPole(0)
prevPos = curPos

cur_t = 0
Do
  'compute timestep
  prev_t = cur_t
  cur_t = 1
  bz.shrink_interval prev_t, cur_t
  rtr.shrink_interval prev_t, cur_t
  Dim timestep As Double
  timestep = TimeOfMove * (cur_t - prev_t)
  
  prevPos = curPos
  curPos = bz.getValue(cur_t)
  
  'test for jerk
  Dim moveJerk As Double
  moveJerk = Vector3D.Length(Vector3D.Combi2(bz.getDeriv(cur_t), 1 / TimeOfMove, bz.getDeriv(prev_t), -1 / TimeOfMove))
  Debug.Assert moveJerk < Me.CurveJerk * 1.1
  
  'Debug.Print "; ";
  Dim rtrJerk As Double
  rtrJerk = Abs(rtr.getDeriv(cur_t) * 1 / TimeOfMove - rtr.getDeriv(prev_t) * 1 / TimeOfMove)
  Debug.Assert rtrJerk < Me.RetractJerk * 1.1
  
  GoSub addMove
  curMove.p1.copyFromT prevPos
  curMove.p2.copyFromT curPos
  curMove.time = timestep
  If curMove.Speed > Me.speedLimit Then curMove.Speed = speedLimit
  curMove.Extrusion = -(rtr.getValue(cur_t) - rtr.getValue(prev_t))
Loop Until cur_t >= 1# - mdlPrecision.RelConfusion

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
Me.bRetract = True
Me.bUnretract = True
End Sub

Private Function minD(ByVal a As Double, ByVal b As Double) As Double
If a > b Then minD = b Else minD = a
End Function
