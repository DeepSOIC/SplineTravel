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
Dim myPoles(0 To 3) As clsVector3D

Public Function getPole(ByVal index As Double) As clsVector3D
If index < 0 Or index > UBound(myPoles) Then Throw errIndexOutOfRange, Source:="getPole"
Set getPole = myPoles(index).Copy
End Function

Public Property Get Pole(ByVal index As Long) As clsVector3D
If index < 0 Or index > UBound(myPoles) Then Throw errIndexOutOfRange, Source:="getPole"
Set Pole = myPoles(index)
End Property

Public Property Set Pole(ByVal index As Long, newRef As clsVector3D)
If index < 0 Or index > UBound(myPoles) Then Throw errIndexOutOfRange, Source:="getPole"
myPoles(index).copyFrom (newRef)
End Property

Private Sub Class_Initialize()
Dim i As Long
For i = 0 To UBound(myPoles)
  Set myPoles(i) = New clsVector3D
Next i
End Sub

't from 0 to 1
Public Sub getValue(ByVal t As Double, ByRef val As clsVector3D)
Dim s As Double
s = 1 - t
If val Is Nothing Then Set val = New clsVector3D
val.copyFrom (myPoles(0).Multed(t * t * t) _
    .Added(myPoles(1).Multed(3# * t * t * s)) _
    .Added(myPoles(2).Multed(3# * t * s * s)) _
    .Added(myPoles(3).Multed(s * s * s)))
End Sub

Public Sub getDeriv(ByVal t As Double, ByRef deriv As clsVector3D)
Dim s As Double
s = 1 - t
If deriv Is Nothing Then Set deriv = New clsVector3D
deriv.copyFrom (myPoles(0).Multed(3 * t * t) _
      .Added(myPoles(1).Multed(3 * 2# * t * s - 3# * t * t)) _
      .Added(myPoles(2).Multed(3 * s * s - 3 * t * 2 * s)) _
      .Added(myPoles(3).Multed(-3 * s * s)))
End Sub

Public Sub getDeriv2(ByVal t As Double, ByRef deriv2 As clsVector3D)
Dim s As Double
s = 1 - t
If deriv2 Is Nothing Then Set deriv2 = New clsVector3D
deriv2.copyFrom (myPoles(0).Multed(6 * t) _
      .Added(myPoles(1).Multed(3 * 2# * s - 3 * 2# * t - 6# * t)) _
      .Added(myPoles(2).Multed(-6 * s - 3 * 2 * s + 3 * t * 2)) _
      .Added(myPoles(3).Multed(6 * s)))
End Sub
