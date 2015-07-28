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
  RetractLen As Double 'amount to retract
  RetractA As Double 'accelleration
  RetractLenAct As Double 'actual possible retract length

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
RetractLen = this.RetractLen
End Property

Public Property Let RetractLen(ByVal newVal As Double)
this.RetractLen = newVal
invalidate
End Property

Private Sub recompute()
this.dt_acc = Sqr(Abs(2 * (this.RetractLen / 2) / this.RetractA))
If this.dt_acc * 4 > this.MoveTime Then
  this.dt_acc = MoveTime / 4 'not enough time for full retract-unretract, shrink
  this.RetractLenAct = this.RetractA * this.dt_acc ^ 2 '/2*2
Else
  this.RetractLenAct = this.RetractLen
End If
this.t_start1 = 0
this.t_mid1 = this.t_start1 + this.dt_acc
this.t_end1 = this.t_mid1 + this.dt_acc
this.t_end2 = this.MoveTime
this.t_mid2 = this.t_end2 - this.dt_acc
this.t_start2 = this.t_mid2 - this.dt_acc
Debug.Assert this.t_end1 <= this.t_start2 + 0.000000000001
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
  v = this.RetractLenAct - this.RetractA * (time - this.t_end1) ^ 2 / 2
ElseIf time < this.t_start2 Then
  v = this.RetractLenAct
ElseIf time < this.t_mid2 Then
  v = this.RetractLenAct - this.RetractA * (time - this.t_start2) ^ 2 / 2
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

t_breakpoint = this.t_start1 / this.MoveTime
If prev_t < t_breakpoint And cur_t > t_breakpoint Then
  cur_t = t_breakpoint
  ret = True
End If
  
t_breakpoint = this.t_start2 / this.MoveTime
If prev_t < t_breakpoint And cur_t > t_breakpoint Then
  cur_t = t_breakpoint
  ret = True
End If

shrink_interval = ret
End Function
