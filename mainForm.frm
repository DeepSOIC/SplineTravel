VERSION 5.00
Begin VB.Form mainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SplineTravel"
   ClientHeight    =   5370
   ClientLeft      =   30
   ClientTop       =   370
   ClientWidth     =   9490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   9490
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEAccell 
      Height          =   370
      Left            =   5280
      TabIndex        =   23
      Text            =   "8000"
      Top             =   3350
      Width           =   1450
   End
   Begin VB.TextBox txtRetract 
      Height          =   370
      Left            =   5270
      TabIndex        =   20
      Text            =   "1.5"
      Top             =   2760
      Width           =   1450
   End
   Begin VB.TextBox txtEJerk 
      Height          =   370
      Left            =   5290
      TabIndex        =   17
      Text            =   "10"
      Top             =   2120
      Width           =   1450
   End
   Begin VB.TextBox txtZJerk 
      Height          =   370
      Left            =   5300
      TabIndex        =   14
      Text            =   "5"
      Top             =   1450
      Width           =   1450
   End
   Begin VB.CommandButton cmdProcessFile 
      Caption         =   "Go"
      CausesValidation=   0   'False
      Height          =   840
      Left            =   3870
      TabIndex        =   13
      Top             =   4050
      Width           =   5150
   End
   Begin VB.TextBox txtSpeedLimit 
      Height          =   370
      Left            =   1260
      TabIndex        =   11
      Text            =   "200"
      Top             =   1500
      Width           =   1450
   End
   Begin VB.TextBox txtCurveJerk 
      Height          =   370
      Left            =   1260
      TabIndex        =   8
      Text            =   "2"
      Top             =   2710
      Width           =   1450
   End
   Begin VB.TextBox txtAccelleration 
      Height          =   370
      Left            =   1260
      TabIndex        =   5
      Text            =   "300"
      Top             =   2120
      Width           =   1450
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
   Begin VB.Label Label16 
      Caption         =   "E accelleration"
      Height          =   610
      Left            =   4150
      TabIndex        =   25
      Top             =   3420
      Width           =   1070
   End
   Begin VB.Label Label15 
      Caption         =   "mm/s2"
      Height          =   240
      Left            =   6830
      TabIndex        =   24
      Top             =   3420
      Width           =   760
   End
   Begin VB.Label Label14 
      Caption         =   "retract"
      Height          =   610
      Left            =   4170
      TabIndex        =   22
      Top             =   2800
      Width           =   1070
   End
   Begin VB.Label Label13 
      Caption         =   "mm"
      Height          =   240
      Left            =   6820
      TabIndex        =   21
      Top             =   2830
      Width           =   760
   End
   Begin VB.Label Label12 
      Caption         =   "E jerk (for retraction)"
      Height          =   610
      Left            =   4180
      TabIndex        =   19
      Top             =   2090
      Width           =   1070
   End
   Begin VB.Label Label11 
      Caption         =   "mm/s"
      Height          =   240
      Left            =   6840
      TabIndex        =   18
      Top             =   2190
      Width           =   760
   End
   Begin VB.Label Label10 
      Caption         =   "Z jerk (for hopping)"
      Height          =   610
      Left            =   4180
      TabIndex        =   16
      Top             =   1360
      Width           =   1070
   End
   Begin VB.Label Label9 
      Caption         =   "mm/s"
      Height          =   240
      Left            =   6850
      TabIndex        =   15
      Top             =   1520
      Width           =   760
   End
   Begin VB.Label label8 
      Caption         =   "mm/s2"
      Height          =   240
      Left            =   2810
      TabIndex        =   12
      Top             =   1565
      Width           =   920
   End
   Begin VB.Label Label7 
      Caption         =   "speed limit"
      Height          =   340
      Left            =   190
      TabIndex        =   10
      Top             =   1515
      Width           =   1210
   End
   Begin VB.Label Label6 
      Caption         =   "mm/s"
      Height          =   240
      Left            =   2810
      TabIndex        =   9
      Top             =   2775
      Width           =   760
   End
   Begin VB.Label Label5 
      Caption         =   "curve tesellation (jerk)"
      Height          =   610
      Left            =   150
      TabIndex        =   7
      Top             =   2590
      Width           =   1070
   End
   Begin VB.Label Label4 
      Caption         =   "mm/s2"
      Height          =   240
      Left            =   2810
      TabIndex        =   6
      Top             =   2185
      Width           =   920
   End
   Begin VB.Label Label3 
      Caption         =   "accelleration"
      Height          =   340
      Left            =   170
      TabIndex        =   4
      Top             =   2135
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
    gen.Retract = val(Me.txtRetract)
    gen.RetractAccelleration = val(Me.txtEAccell)
    gen.RetractJerk = val(Me.txtEJerk)
    gen.ZJerk = val(Me.txtZJerk)
  End If
  gen.p1.copyFromT mv.prevBuildMove.CompleteStateAfter.pos
  gen.p2.copyFromT mv.nextBuildMove.CompleteStateBefore.pos
  Set gen.inSpeed = mv.prevBuildMove.getExitSpeed
  Set gen.outSpeed = mv.nextBuildMove.getEnterSpeed
  Dim arrSegments() As clsGMove
  Dim bz As clsBezier, MoveTime As Double
  Set bz = gen.FitBezier(MoveTime)
  gen.GenerateMoveTrainForBezier arrSegments, bz, MoveTime
  Dim isegment As Long
  For isegment = 0 To UBound(arrSegments)
    Set cmd = New clsGCommand
    chain.Add cmd, Before:=mv.nextBuildMove
    Dim EError As Double
    EError = 0
    cmd.strLine = arrSegments(isegment).GenerateGCode(cmd.prevCommand.CompleteStateAfter, EError)
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

