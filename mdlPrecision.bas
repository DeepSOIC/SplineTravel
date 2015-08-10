Attribute VB_Name = "mdlPrecision"
Option Explicit


Public posDecimals As Integer
Public extrDecimals As Integer
Public speedDecimals As Integer

Public posConfusion As Double
Public extrConfusion As Double
Public speedConfusion As Double

Public Const RelConfusion As Double = 0.000000000001

Public Sub InitModule()
posDecimals = 3
extrDecimals = 3
speedDecimals = -1
updateConfusions
End Sub

Public Sub updateConfusions()
posConfusion = 10 ^ (-posDecimals - 1)
extrConfusion = 10 ^ (-extrDecimals - 1)
speedConfusion = 10 ^ (-speedDecimals - 1)
End Sub

Public Function Round(ByVal Value As Double, ByVal Decimals As Integer) As Double
If Decimals < 0 Then Decimals = 0
Round = VBA.Round(Value, Decimals)
End Function
