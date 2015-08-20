VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBezier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type typThis
  Poles(0 To 3) As typVector3D
  t_of_low_a As Double 't-point of lowest second derivative
  max_a As Double 'max second derivative in the curve ()
  DerivJerk As Double
  
  invalidated As Boolean
End Type
Dim this As typThis

Friend Function getPole(ByVal index As Double) As typVector3D
If index < 0 Or index > UBound(this.Poles) Then Throw errIndexOutOfRange, Source:="getPole"
getPole = this.Poles(index)
End Function

Friend Property Get Pole(ByVal index As Long) As typVector3D
If index < 0 Or index > UBound(this.Poles) Then Throw errIndexOutOfRange, Source:="getPole"
Pole = this.Poles(index)
End Property

Friend Property Let Pole(ByVal index As Long, newVal As typVector3D)
If index < 0 Or index > UBound(this.Poles) Then Throw errIndexOutOfRange, Source:="getPole"
this.Poles(index) = newVal
this.invalidated = True
End Property

Public Property Get DerivJerk() As Double
DerivJerk = this.DerivJerk
End Property

Public Property Let DerivJerk(ByVal newVal As Double)
this.DerivJerk = newVal
this.invalidated = True
End Property

Private Sub Class_Initialize()
this.invalidated = True
this.DerivJerk = 0.0000001
End Sub

't from 0 to 1
Friend Function getValue(ByVal t As Double) As typVector3D
Dim s As Double
s = 1 - t
getValue = Vector3D.Combi4( _
  this.Poles(0), s * s * s, _
  this.Poles(1), 3# * s * s * t, _
  this.Poles(2), 3# * s * t * t, _
  this.Poles(3), t * t * t)
End Function

Friend Function getDeriv(ByVal t As Double) As typVector3D
Dim s As Double
s = 1 - t
getDeriv = Vector3D.Combi4( _
  this.Poles(0), -3 * s * s, _
  this.Poles(1), -3 * 2 * s * t + 3 * s * s, _
  this.Poles(2), -3 * t * t + 3 * 2 * t * s, _
  this.Poles(3), 3 * t * t)
End Function

Friend Function getDeriv2(ByVal t As Double) As typVector3D
Dim s As Double
s = 1 - t
getDeriv2 = Vector3D.Combi4( _
  this.Poles(0), 6 * s, _
  this.Poles(1), 3 * 2 * t - 3 * 2 * s - 3 * 2 * s, _
  this.Poles(2), -3 * 2 * t - 3 * 2 * t + 3 * 2 * s, _
  this.Poles(3), 6 * t)
End Function

Public Function shrink_interval(ByVal prev_t As Double, ByRef cur_t As Double) As Boolean
If this.invalidated Then recompute

Dim t_breakpoint As Double
t_breakpoint = getNextBreakpoint(prev_t)

Dim tstep As Double
tstep = this.DerivJerk / this.max_a 'start off a t-step that is small enough for sure

Dim cnt As Long
Dim logFactor As Double
logFactor = 0.25 'sets, how fast to increase the step
Dim valid_tstep As Double
valid_tstep = tstep


'a few iterations to get the t-step that fits the jerk
Dim i As Long
For i = 1 To 7 'the more iteration, the closer the jerk will be to the required.
  tstep = tstep * Exp(logFactor)
  Dim jerk As Double
  jerk = Vector3D.Dist(Me.getDeriv(prev_t + tstep), Me.getDeriv(prev_t))
  If prev_t < this.t_of_low_a + mdlPrecision.RelConfusion And prev_t + tstep > this.t_of_low_a Then
    'we might have hopped over the t of lowest d2 / highest d1, let's check if we didn't miss a high jerk point
    Dim jerk2 As Double
    jerk2 = Vector3D.Dist(Me.getDeriv(this.t_of_low_a), Me.getDeriv(prev_t))
    If jerk2 > jerk Then jerk = jerk2
  End If
  If jerk < this.DerivJerk Then
    'timestep still valid
    'remember it and continue to increase
    valid_tstep = tstep
  Else
    'timestep invalid
    'roll back and increase by a smaller amount on next iter
    tstep = valid_tstep
    logFactor = logFactor / 2
  End If
Next i
Debug.Assert tstep = valid_tstep

'collapse to a breakpoint
If prev_t + tstep > t_breakpoint Then
  tstep = t_breakpoint - prev_t
End If

'small final step prevention logic
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

Public Function getNextBreakpoint(ByVal prev_t As Double) As Double
Dim ret As Double: ret = 1
Dim t_breakpoint As Double

't_breakpoint = this.t_of_low_a
'If prev_t < t_breakpoint - mdlPrecision.RelConfusion And ret > t_breakpoint + mdlPrecision.RelConfusion Then
'  ret = t_breakpoint
'End If

getNextBreakpoint = ret
End Function

Public Function recompute()
Dim a1 As typVector3D, a2 As typVector3D
a1 = Me.getDeriv2(0#)
a2 = Me.getDeriv2(1#)
Dim d As typVector3D
d = Vector3D.Subtracted(a2, a1)
If Vector3D.Length(d) < 1E-100 Then
  this.t_of_low_a = 1000 'doesn't matter, deriv2 is constant
Else
  Dim t_a As Double
  t_a = -Dot(a1, Vector3D.Normalized(d))
  this.t_of_low_a = t_a / Vector3D.Length(d)
End If
this.max_a = maxD(Vector3D.Length(a1), Vector3D.Length(a2))
this.invalidated = False
End Function


Private Function maxD(ByVal a As Double, ByVal b As Double) As Double
If a > b Then maxD = a Else maxD = b
End Function

