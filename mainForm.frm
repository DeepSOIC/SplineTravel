VERSION 5.00
Begin VB.Form mainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SplineTravel"
   ClientHeight    =   4320
   ClientLeft      =   30
   ClientTop       =   370
   ClientWidth     =   9490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   9490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProcessFile 
      Caption         =   "Go"
      CausesValidation=   0   'False
      Height          =   840
      Left            =   3650
      TabIndex        =   13
      Top             =   3280
      Width           =   5150
   End
   Begin VB.TextBox txtSpeedLimit 
      Height          =   400
      Left            =   1440
      TabIndex        =   11
      Text            =   "200"
      Top             =   1520
      Width           =   1260
   End
   Begin VB.TextBox txtCurveJerk 
      Height          =   320
      Left            =   1400
      TabIndex        =   8
      Text            =   "2"
      Top             =   2730
      Width           =   1380
   End
   Begin VB.TextBox txtAccelleration 
      Height          =   400
      Left            =   1420
      TabIndex        =   5
      Text            =   "300"
      Top             =   2140
      Width           =   1260
   End
   Begin VB.TextBox txtFNOut 
      Height          =   410
      Left            =   1200
      TabIndex        =   3
      Text            =   "txtFNOut"
      Top             =   480
      Width           =   5590
   End
   Begin VB.TextBox txtFNIn 
      Height          =   360
      Left            =   1190
      TabIndex        =   0
      Text            =   "txtFNIn"
      Top             =   50
      Width           =   5610
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   390
      Left            =   140
      TabIndex        =   14
      Top             =   3420
      Width           =   1040
   End
   Begin VB.Label Label8 
      Caption         =   "mm/s2"
      Height          =   240
      Left            =   2830
      TabIndex        =   12
      Top             =   1550
      Width           =   920
   End
   Begin VB.Label Label7 
      Caption         =   "speed limit"
      Height          =   340
      Left            =   190
      TabIndex        =   10
      Top             =   1570
      Width           =   1210
   End
   Begin VB.Label Label6 
      Caption         =   "mm/s"
      Height          =   240
      Left            =   2920
      TabIndex        =   9
      Top             =   2750
      Width           =   760
   End
   Begin VB.Label Label5 
      Caption         =   "curve tesellation (jerk)"
      Height          =   610
      Left            =   150
      TabIndex        =   7
      Top             =   2560
      Width           =   1070
   End
   Begin VB.Label Label4 
      Caption         =   "mm/s2"
      Height          =   240
      Left            =   2810
      TabIndex        =   6
      Top             =   2170
      Width           =   920
   End
   Begin VB.Label Label3 
      Caption         =   "accelleration"
      Height          =   340
      Left            =   170
      TabIndex        =   4
      Top             =   2190
      Width           =   1210
   End
   Begin VB.Label Label2 
      Caption         =   "output"
      Height          =   380
      Left            =   50
      TabIndex        =   2
      Top             =   480
      Width           =   1090
   End
   Begin VB.Label Label1 
      Caption         =   "input"
      Height          =   270
      Left            =   40
      TabIndex        =   1
      Top             =   40
      Width           =   1020
   End
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type typTravelMoveRef
  prevBuildMove As clsGCommand
  nextBuildMove As clsGCommand
  firstTravelMove As clsGCommand
End Type

Private Sub cmdProcessFile_Click()
cmdProcessFile.Enabled = False
Dim f1 As Long
f1 = FreeFile
Dim chain As New clsChain
Dim iline As Long
Open Me.txtFNIn For Input As f1
  On Error GoTo eh
  Dim ln As String
  Do While Not EOF(f1)
    Line Input #(f1), ln
    chain.Add New clsGCommand
    chain.last.strLine = ln
    chain.last.ParseString
    chain.last.RecomputeStates
    iline = iline + 1
    If timeToDoEvents Then
      Me.cmdProcessFile.Caption = "reading line " + Str(iline)
      DoEvents
    End If
  Loop
Close f1
Dim nLines As Long
nLines = iline

Me.cmdProcessFile.Caption = "searching for travel moves"
DoEvents

Dim travelMoves() As typTravelMoveRef
ReDim travelMoves(0 To 10)
Dim nMoves As Long: nMoves = 0

Dim cmd As clsGCommand
Dim wasTravel As Boolean 'this var is used for filtering sequences of travel moves and to post just one from the sequence to the array
wasTravel = True 'should filter out the very first moves
Set cmd = chain.first
iline = 0
Do
  If cmd.isTravelMove Then
    If Not wasTravel Then
      'travel move after a build move - use it
      If nMoves + 1 > UBound(travelMoves) Then
        ReDim Preserve travelMoves(0 To nMoves * 1.5)
      End If
      Set travelMoves(nMoves).firstTravelMove = cmd
      nMoves = nMoves + 1
    Else
      'travel move after a travel move - skip it
    End If
    wasTravel = True
  ElseIf cmd.isBuildMove Then
    If wasTravel And nMoves > 0 Then
      Set travelMoves(nMoves - 1).nextBuildMove = cmd
    End If
    wasTravel = False
    Set travelMoves(nMoves).prevBuildMove = cmd
  End If
  iline = iline + 1
  If cmd Is chain.last Then Exit Do
  Set cmd = cmd.nextCommand
  If timeToDoEvents Then
    Me.cmdProcessFile.Caption = "line " + Str(iline) + " of " + Str(nLines)
    DoEvents
  End If
Loop

Me.cmdProcessFile.Caption = "generating splines"
DoEvents

'replace moves with splines
Dim imove As Long
For imove = 0 To nMoves - 1
  'experimental: delete everything whatsoever between the build moves
  Dim mv As typTravelMoveRef
  mv = travelMoves(imove)
  If mv.nextBuildMove Is Nothing Then Exit For 'fixes fail on the last travel move, where there is no nex build move
  chain.withdrawChain(mv.prevBuildMove.nextCommand, mv.nextBuildMove.prevCommand).delete
  Dim gen As clsTravelGenerator
  If gen Is Nothing Then
    Set gen = New clsTravelGenerator
    gen.accelleration = val(Me.txtAccelleration)
    gen.CurveJerk = val(Me.txtCurveJerk)
    gen.speedLimit = val(Me.txtSpeedLimit)
  End If
  gen.p1.copyFromT mv.prevBuildMove.CompleteStateAfter.pos
  gen.p2.copyFromT mv.nextBuildMove.CompleteStateBefore.pos
  Set gen.inSpeed = mv.prevBuildMove.getExitSpeed
  Set gen.outSpeed = mv.nextBuildMove.getEnterSpeed
  Dim arrSegments() As clsGMove
  Dim bz As clsBezier, moveTime As Double
  Set bz = gen.FitBezier(moveTime)
  gen.GenerateMoveTrainForBezier arrSegments, bz, moveTime
  Dim isegment As Long
  For isegment = 0 To UBound(arrSegments)
    Set cmd = New clsGCommand
    chain.Add cmd, Before:=mv.nextBuildMove
    cmd.strLine = arrSegments(isegment).GenerateGCode(cmd.prevCommand.CompleteStateAfter)
    cmd.ParseString throwIfInvalid:=True
    cmd.RecomputeStates
  Next isegment
  If timeToDoEvents Then
    Me.cmdProcessFile.Caption = "generating spline " + Str(imove) + " of " + Str(nMoves)
    DoEvents
  End If
Next imove

Me.cmdProcessFile.Caption = "writing file"
DoEvents

iline = 0
Open txtFNOut For Output As f1
  Set cmd = chain.first
  Do
    Print #(f1), cmd.strLine
    iline = iline + 1
    If cmd Is chain.last Then Exit Do
    Set cmd = cmd.nextCommand
  Loop
Close f1

Me.cmdProcessFile.Caption = "freeing memory"
DoEvents

chain.delete
Me.cmdProcessFile.Enabled = True
Exit Sub
eh:
  PushError
  Close f1
  PopError
  MsgError
  chain.delete
  Me.cmdProcessFile.Enabled = True
End Sub

Public Function timeToDoEvents()
Static lastDidTime As Double
If Abs(Timer - lastDidTime) > 0.3 Then
  timeToDoEvents = True
  lastDidTime = Timer
End If
End Function

Private Sub Form_Load()
mdlCommon.extrDecimals = 4
mdlCommon.posDecimals = 3
mdlCommon.speedDecimals = 3
End Sub
