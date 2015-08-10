VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRetractCurve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type typThis
  MoveTime As Double 'time of the whole travel move
  RetractLenSetting As Double 'amount to retract (setting)
  RetractLen As Double 'actual possible retract length
  RetractASetting As Double 'accelleration - wanted
  RetractA As Double 'accelleration - actual (can be increased in some cases)
  
  bRetract As Boolean 'set to false to suppress retraction, but keep unretraction
  bUnretract As Boolean 'set to false to suppress unretraction, but keep retraction
  
  DerivJerk As Double 'defines tessellation; in E units (speed not by time, but by parameter)

  dt_acc As Double 'time it takes to accellerate to full retraction speed
  t_start1 As Double '=zero
  t_mid1 As Double 'moment of time when retract accelleration ends
  t_end1 As Double 'moment of time when deceleration ends
  t_start2 As Double
  t_mid2 As Double
  t_end2 As Double
  invalidated As Boolean
End Type
Dim this As typThis

Private Sub invalidate()
this.invalidated = True
End Sub

Public Property Get MoveTime() As Double
MoveTime = this.MoveTime
End Property

Public Property Let MoveTime(ByVal newVal As Double)
this.MoveTime = newVal
invalidate
End Property

Public Property Get RetractA() As Double
RetractA = this.RetractA
End Property

Public Property Let RetractA(ByVal newVal As Double)
this.RetractA = newVal
invalidate
End Property

Public Property Get RetractLen() As Double
RetractLen = this.RetractLenSetting
End Property

Public Property Let RetractLen(ByVal newVal As Double)
this.RetractLenSetting = newVal
invalidate
End Property





Private Sub recompute()
this.dt_acc = Sqr(Abs(2 * (this.RetractLenSetting / 2) / this.RetractA))
If this.bRetract And this.bUnretract Then
  If this.dt_acc * 4 > this.MoveTime Then
    this.dt_acc = MoveTime / 4 'not enough time for full retract-unretract, shrink
    this.RetractLen = this.RetractA * this.dt_acc ^ 2 '/2*2
  Else
    this.RetractLen = this.RetractLenSetting
  End If
ElseIf this.bRetract Or this.bUnretract Then
  If this.dt_acc * 2 > this.MoveTime Then
    this.dt_acc = MoveTime / 2 'not enough time to complete, increase accelleration.
    this.RetractA = (this.RetractLenSetting / 2) / (this.MoveTime / 2) ^ 2
    this.RetractLen = this.RetractLenSetting
  Else
    this.RetractLen = this.RetractLenSetting
  End If
Else
  'disable everything by setting retractlength to zero
  this.RetractLen = 0
  'the rest will be done by shifting out time points, as follows...
End If

'retract time moments
this.t_start1 = 0
If Not this.bRetract Then this.t_start1 = this.t_start1 - 10 * this.MoveTime 'suppress retract by simply moving it away
this.t_mid1 = this.t_start1 + this.dt_acc
this.t_end1 = this.t_mid1 + this.dt_acc

'unretract time moments
this.t_end2 = this.MoveTime
If Not this.bUnretract Then this.t_end2 = this.t_end2 + 10 * this.MoveTime 'suppress unretract by simply moving it away
this.t_mid2 = this.t_end2 - this.dt_acc
this.t_start2 = this.t_mid2 - this.dt_acc
this.invalidated = False
End Sub

Public Function getValue(ByVal t As Double) As Double
If this.invalidated Then recompute
Dim time  As Double: time = t * this.MoveTime
Dim v As Double 'output value
If time < this.t_start1 Then
  v = 0
ElseIf time < this.t_mid1 Then
  v = this.RetractA * (time - this.t_start1) ^ 2 / 2
ElseIf time < this.t_end1 Then
  If this.bRetract Then v = this.RetractLen - this.RetractA * (time - this.t_end1) ^ 2 / 2
ElseIf time < this.t_start2 Then
  v = this.RetractLen
ElseIf time < this.t_mid2 Then
  v = this.RetractLen - this.RetractA * (time - this.t_start2) ^ 2 / 2
ElseIf time < this.t_end2 Then
  v = this.RetractA * (time - this.t_end2) ^ 2 / 2
Else
  v = 0
End If
getValue = v
End Function


Public Function getDeriv(ByVal t As Double) As Double
If this.invalidated Then recompute
Dim time  As Double: time = t * this.MoveTime
Dim v As Double 'output value
If time < this.t_start1 Then
  v = 0
ElseIf time < this.t_mid1 Then
  v = this.RetractA * (time - this.t_start1)
ElseIf time < this.t_end1 Then
  v = -this.RetractA * (time - this.t_end1)
ElseIf time < this.t_start2 Then
  v = 0
ElseIf time < this.t_mid2 Then
  v = -this.RetractA * (time - this.t_start2)
ElseIf time < this.t_end2 Then
  v = this.RetractA * (time - this.t_end2)
Else
  v = 0
End If
getDeriv = v * this.MoveTime
End Function

Public Function getDeriv2(ByVal t As Double) As Double
If this.invalidated Then recompute
Dim time As Double: time = t * this.MoveTime
Dim v As Double 'output value
If time < this.t_start1 Then
  v = 0
ElseIf time < this.t_mid1 Then
  v = this.RetractA
ElseIf time < this.t_end1 Then
  v = -this.RetractA
ElseIf time < this.t_start2 Then
  v = 0
ElseIf time < this.t_mid2 Then
  v = -this.RetractA
ElseIf time < this.t_end2 Then
  v = this.RetractA
Else
  v = 0
End If
getDeriv2 = v * this.MoveTime ^ 2
End Function

'shrinks given interval to keep from skipping the important stuff
Public Function shrink_interval(ByVal prev_t As Double, ByRef cur_t As Double) As Boolean
If this.invalidated Then recompute
Dim ret As Boolean: ret = False

Dim t_breakpoint As Double
t_breakpoint = getNextBreakpoint(prev_t)

Dim tstep As Double

tstep = 100
Dim acc As Double
acc = Abs(Me.getDeriv2(prev_t))
If acc > 1E-100 Then
  tstep = this.DerivJerk / acc
End If

'collapse to a breakpoint
If prev_t + tstep > t_breakpoint Then
  tstep = t_breakpoint - prev_t
End If

'small final step prevention logic: if step to nearest breakpoint after this step is much smaller than this step, reduce this step
Dim tstepToEnd As Double
tstepToEnd = t_breakpoint - prev_t
If tstep < tstepToEnd - mdlPrecision.RelConfusion And tstepToEnd - tstep < tstepToEnd * 0.25 Then
  tstep = tstepToEnd * 0.3
End If

If cur_t > prev_t + tstep + mdlPrecision.RelConfusion Then
  cur_t = prev_t + tstep
  shrink_interval = True
End If
End Function

Private Function getNextBreakpoint(ByVal prev_t As Double) As Double
Dim ret As Double
ret = 1
Dim t_breakpoint As Double

t_breakpoint = this.t_start1 / this.MoveTime
If prev_t < t_breakpoint - mdlPrecision.RelConfusion And ret > t_breakpoint + mdlPrecision.RelConfusion Then
  ret = t_breakpoint
End If
  
t_breakpoint = this.t_start2 / this.MoveTime
If prev_t < t_breakpoint - mdlPrecision.RelConfusion And ret > t_breakpoint + mdlPrecision.RelConfusion Then
  ret = t_breakpoint
End If

getNextBreakpoint = ret
End Function

Private Sub Class_Initialize()
'fill up some defaults
this.bRetract = True
this.bUnretract = True
End Sub

Public Property Get bRetract() As Boolean
bRetract = this.bRetract
End Property

Public Property Let bRetract(ByVal newVal As Boolean)
this.bRetract = newVal
invalidate
End Property

Public Property Get bUnretract() As Boolean
bUnretract = this.bUnretract
End Property

Public Property Let bUnretract(ByVal newVal As Boolean)
this.bUnretract = newVal
invalidate
End Property

''returns actual length of retraction (will be reduced
'if full retract-unretract sequence doesn't fit into
'the MoveTime)
Public Property Get ActualRetract() As Double
If this.invalidated Then recompute
ActualRetract = this.RetractLen
End Property

Public Property Get DerivJerk() As Double
DerivJerk = this.DerivJerk
End Property

Public Property Let DerivJerk(ByVal newVal As Double)
this.DerivJerk = newVal
invalidate
End Property
