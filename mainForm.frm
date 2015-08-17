VERSION 5.00
Begin VB.Form mainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SplineTravel"
   ClientHeight    =   5620
   ClientLeft      =   30
   ClientTop       =   370
   ClientWidth     =   13540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5620
   ScaleWidth      =   13540
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "      seam concealment"
      Height          =   3010
      Left            =   7690
      TabIndex        =   32
      Top             =   880
      Width           =   3830
      Begin VB.TextBox txtLoopTol 
         Height          =   370
         Left            =   1330
         TabIndex        =   34
         Text            =   "0.3"
         Top             =   1250
         Width           =   1450
      End
      Begin VB.CheckBox chkSeamConceal 
         Height          =   370
         Left            =   130
         TabIndex        =   33
         Top             =   -50
         Value           =   1  'Checked
         Width           =   180
      End
      Begin VB.Label Label19 
         Caption         =   $"mainForm.frx":0000
         Height          =   630
         Left            =   110
         TabIndex        =   37
         Top             =   350
         Width           =   3600
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label18 
         Caption         =   "mm"
         Height          =   240
         Left            =   2900
         TabIndex        =   36
         Top             =   1320
         Width           =   760
      End
      Begin VB.Label Label17 
         Caption         =   "loop detection tolerance"
         Height          =   610
         Left            =   230
         TabIndex        =   35
         Top             =   1200
         Width           =   1070
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Smooth out travel moves"
      Height          =   3500
      Left            =   80
      TabIndex        =   10
      Top             =   890
      Width           =   7460
      Begin VB.TextBox txtAccelleration 
         Height          =   370
         Left            =   1120
         TabIndex        =   17
         Text            =   "800"
         Top             =   1610
         Width           =   1450
      End
      Begin VB.TextBox txtCurveJerk 
         Height          =   370
         Left            =   1120
         TabIndex        =   16
         Text            =   "2"
         Top             =   2200
         Width           =   1450
      End
      Begin VB.TextBox txtSpeedLimit 
         Height          =   370
         Left            =   1120
         TabIndex        =   15
         Text            =   "200"
         Top             =   990
         Width           =   1450
      End
      Begin VB.TextBox txtZJerk 
         Height          =   370
         Left            =   1120
         TabIndex        =   14
         Text            =   "0"
         Top             =   2860
         Width           =   1450
      End
      Begin VB.TextBox txtEJerk 
         Height          =   370
         Left            =   5080
         TabIndex        =   13
         Text            =   "15"
         Top             =   2220
         Width           =   1450
      End
      Begin VB.TextBox txtRetract 
         Height          =   370
         Left            =   5090
         TabIndex        =   12
         Text            =   "1.5"
         Top             =   930
         Width           =   1450
      End
      Begin VB.TextBox txtEAccell 
         Height          =   370
         Left            =   5070
         TabIndex        =   11
         Text            =   "1000"
         Top             =   1580
         Width           =   1450
      End
      Begin VB.Label Label20 
         Caption         =   $"mainForm.frx":008B
         Height          =   700
         Left            =   170
         TabIndex        =   38
         Top             =   300
         Width           =   7060
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "accelleration"
         Height          =   340
         Left            =   30
         TabIndex        =   31
         Top             =   1630
         Width           =   1210
      End
      Begin VB.Label Label4 
         Caption         =   "mm/s2"
         Height          =   240
         Left            =   2670
         TabIndex        =   30
         Top             =   1680
         Width           =   920
      End
      Begin VB.Label Label5 
         Caption         =   "curve tesellation (jerk)"
         Height          =   610
         Left            =   50
         TabIndex        =   29
         Top             =   2060
         Width           =   1070
      End
      Begin VB.Label Label6 
         Caption         =   "mm/s"
         Height          =   240
         Left            =   2670
         TabIndex        =   28
         Top             =   2270
         Width           =   760
      End
      Begin VB.Label Label7 
         Caption         =   "speed limit"
         Height          =   340
         Left            =   120
         TabIndex        =   27
         Top             =   1050
         Width           =   1210
      End
      Begin VB.Label label8 
         Caption         =   "mm/s"
         Height          =   240
         Left            =   2670
         TabIndex        =   26
         Top             =   1060
         Width           =   920
      End
      Begin VB.Label Label9 
         Caption         =   "mm/s"
         Height          =   240
         Left            =   2670
         TabIndex        =   25
         Top             =   2930
         Width           =   760
      End
      Begin VB.Label Label10 
         Caption         =   "Z jerk (for hopping)"
         Height          =   610
         Left            =   70
         TabIndex        =   24
         Top             =   2870
         Width           =   1070
      End
      Begin VB.Label Label11 
         Caption         =   "mm/s"
         Height          =   240
         Left            =   6630
         TabIndex        =   23
         Top             =   2290
         Width           =   760
      End
      Begin VB.Label Label12 
         Caption         =   "E jerk (for retraction)"
         Height          =   610
         Left            =   3910
         TabIndex        =   22
         Top             =   2180
         Width           =   1070
      End
      Begin VB.Label Label13 
         Caption         =   "mm"
         Height          =   240
         Left            =   6640
         TabIndex        =   21
         Top             =   1000
         Width           =   760
      End
      Begin VB.Label Label14 
         Caption         =   "retract"
         Height          =   610
         Left            =   3990
         TabIndex        =   20
         Top             =   970
         Width           =   1070
      End
      Begin VB.Label Label15 
         Caption         =   "mm/s2"
         Height          =   240
         Left            =   6620
         TabIndex        =   19
         Top             =   1650
         Width           =   760
      End
      Begin VB.Label Label16 
         Caption         =   "E accelleration"
         Height          =   610
         Left            =   3890
         TabIndex        =   18
         Top             =   1610
         Width           =   1070
      End
   End
   Begin VB.CommandButton cmdResetSettings 
      Caption         =   "reset to defaults"
      Height          =   430
      Left            =   11850
      TabIndex        =   9
      Top             =   330
      Width           =   1600
   End
   Begin VB.Frame Frame1 
      Caption         =   "Presets"
      Height          =   730
      Left            =   20
      TabIndex        =   5
      Top             =   40
      Width           =   8180
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         Height          =   390
         Left            =   5520
         TabIndex        =   8
         Top             =   250
         Width           =   720
      End
      Begin VB.ComboBox cmbPreset 
         Height          =   280
         ItemData        =   "mainForm.frx":015B
         Left            =   100
         List            =   "mainForm.frx":015D
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "cmbPeset"
         Top             =   310
         Width           =   4110
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save as..."
         Height          =   340
         Left            =   4320
         TabIndex        =   6
         Top             =   270
         Width           =   1060
      End
   End
   Begin VB.CommandButton cmdProcessFile 
      Caption         =   "Go"
      CausesValidation=   0   'False
      Height          =   840
      Left            =   7920
      TabIndex        =   4
      Top             =   4440
      Width           =   5150
   End
   Begin VB.TextBox txtFNOut 
      Height          =   410
      Left            =   1530
      TabIndex        =   3
      Tag             =   "!f"
      Top             =   4900
      Width           =   5590
   End
   Begin VB.TextBox txtFNIn 
      Height          =   360
      Left            =   1520
      TabIndex        =   0
      Tag             =   "!f"
      Top             =   4470
      Width           =   5610
   End
   Begin VB.Label Label2 
      Caption         =   "output"
      Height          =   380
      Left            =   340
      TabIndex        =   2
      Top             =   4930
      Width           =   1090
   End
   Begin VB.Label Label1 
      Caption         =   "input"
      Height          =   270
      Left            =   330
      TabIndex        =   1
      Top             =   4490
      Width           =   1020
   End
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum eChainType
  ectOther = 0
  ectBuildChain = 1
  ectTravelChain = 2
End Enum

Private Enum eRetractBlenderState
  rbsUnRetracting = 0
  rbsRetracting = 1
End Enum

Private Type typMoveChain
  chain As clsChain
  chType As eChainType
  firstMoveRef As clsGCommand 'link to the first command of relevant type in the group (note that group can contain any number of ectOther commands as well)
  lastMoveRef As clsGCommand
  
  ''flags relevant to build groups, indicating that retract and
  ''unretract were injected during processing. Usually, either
  ''none, or both
  retractInjected As Boolean
  unretractInjected As Boolean
End Type

Private Type typTravelMoveRef
  'prevBuildMoveBegin As clsGCommand
  prevBuildMoveEnd As clsGCommand
  'prevBuildLoopIsLoop As Boolean
  firstTravelMove As clsGCommand
  nextBuildMoveBegin As clsGCommand
  'nextBuildMoveEnd As clsGCommand
End Type

Private Sub cmdProcessFile_Click()
cmdProcessFile.Enabled = False
Dim f1 As Long
f1 = FreeFile
Dim chain As New clsChain
Dim iline As Long
On Error GoTo eh
Open Me.txtFNIn For Input As f1
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

Dim moveGroups() As typMoveChain
ReDim moveGroups(0 To 10)
Dim nMoveGroups As Long: nMoveGroups = 1 'first group is a dummy group, that will hold setup commands
moveGroups(0).chType = ectOther

Dim cmd As clsGCommand
Set cmd = chain.first
iline = 0
Do
  Dim curCmdType As eChainType
  If cmd.isBuildMove Then
    curCmdType = ectBuildChain
  ElseIf cmd.isTravelMove Or cmd.isExtruderMove Then
    curCmdType = ectTravelChain
  Else
    curCmdType = ectOther
  End If
  
  If curCmdType <> ectOther Then
    If moveGroups(nMoveGroups - 1).chType <> curCmdType Then
      'command type has changed - start a new group
      nMoveGroups = nMoveGroups + 1
      If nMoveGroups + 1 > UBound(moveGroups) Then
        ReDim Preserve moveGroups(0 To nMoveGroups * 1.5)
      End If
      
      Set moveGroups(nMoveGroups - 1).firstMoveRef = cmd
      Set moveGroups(nMoveGroups - 1).lastMoveRef = cmd 'initialize, essential if just one move in the group
      
      moveGroups(nMoveGroups - 1).chType = curCmdType
    Else
      'command type hasn't changed, update the link to last
      Set moveGroups(nMoveGroups - 1).lastMoveRef = cmd
    End If
  End If
  
  iline = iline + 1
  If cmd Is chain.last Then Exit Do
  Set cmd = cmd.nextCommand
  If timeToDoEvents Then
    Me.cmdProcessFile.Caption = "line " + Str(iline) + " of " + Str(nLines)
    DoEvents
  End If
Loop

Me.cmdProcessFile.Caption = "splitting"
DoEvents

'split the chain
Dim iGroup As Long
For iGroup = 0 To nMoveGroups - 1
  Dim c1 As clsGCommand, c2 As clsGCommand
  If iGroup = 0 Then
    Set c1 = chain.first
  Else
    Set c1 = moveGroups(iGroup - 1).chain.last.nextCommand
  End If
  If iGroup = nMoveGroups - 1 Then
    Set c2 = chain.last
  Else
    Set c2 = moveGroups(iGroup + 1).firstMoveRef.prevCommand
  End If
  Set moveGroups(iGroup).chain = chain.withdrawChain(c1, c2, preserveLinks:=True)
Next iGroup
Debug.Assert (chain.size = 0) 'chain should have been taken apart completely while splitting

If Me.chkSeamConceal.Value = vbChecked Then
  Me.cmdProcessFile.Caption = "blending seams"
  DoEvents
  
  Dim loopTol As Double
  loopTol = val(Me.txtLoopTol)
  Dim retractTime As Double
  retractTime = val(Me.txtRetract) / val(Me.txtEJerk)
  Dim retractSpeed As Double
  retractSpeed = val(Me.txtEJerk)
  
  For iGroup = 0 To nMoveGroups - 1
    If moveGroups(iGroup).chType = ectBuildChain Then
      Dim p1 As typVector3D, p2 As typVector3D
      p1 = moveGroups(iGroup).firstMoveRef.CompleteStateBefore.pos
      p2 = moveGroups(iGroup).lastMoveRef.CompleteStateAfter.pos
      If Vector3D.Dist(p1, p2) <= loopTol Then
        'generate unretract
        Dim t As Double
        Dim EError1 As Double: EError1 = 0
        Dim EError2 As Double: EError2 = 0
        t = retractTime
        Set cmd = moveGroups(iGroup).firstMoveRef
        Dim cmd2 As clsGCommand
        Dim chainRetract As clsChain
        Dim state As eRetractBlenderState
        state = rbsUnRetracting
        Do
          cmd.constructMove
          t = t - cmd.execTime
          Dim move As clsGMove
          Set move = cmd.getMove
          If state = rbsUnRetracting Then
            If Abs(t * retractSpeed) < 0.01 Or t > 0 Then
              'unretraction takes up this command as a whole (and may end with it)
                                        
              'add copy of the command to the end, for filling the empty piece created while unretracting
              Set cmd2 = New clsGCommand
              moveGroups(iGroup).chain.Add cmd2, After:=moveGroups(iGroup).lastMoveRef
              Set moveGroups(iGroup).lastMoveRef = cmd2 'note: this may potentially cause mulpiple passes of the loop, if unretraction is not possible within one loop. This may be buggy. Disabling it requires serious refactor (prediction of the situation and preventing retraction injection beforehand).
              cmd2.RecomputeStates
              cmd2.setMove move, EError2
              cmd2.RecomputeStates
              
              'modify the command, injecting unretraction
              move.Extrusion = retractSpeed * move.time
              cmd.setMove move, EError1
              cmd.RecomputeStates
              
              If Abs(t * retractSpeed) < 0.01 Then
                state = rbsRetracting 'slight under- or over-extrusion doesn't require a split
                t = retractTime
              End If
            Else
              cmd.split t + cmd.execTime, EError1
              
              cmd.constructMove 'need again, because split modified it
              Set move = cmd.getMove
              
              'add copy of the command, for retraction
              Set cmd2 = New clsGCommand
              moveGroups(iGroup).chain.Add cmd2, After:=moveGroups(iGroup).lastMoveRef
              Set moveGroups(iGroup).lastMoveRef = cmd2
              cmd2.RecomputeStates
              cmd2.setMove move, EError2
              cmd2.RecomputeStates
              
              'modify the first part of splitting, injecting unretraction
              move.Extrusion = retractSpeed * move.time
              cmd.setMove move, EError1
              cmd.RecomputeStates
              
              cmd.nextCommand.RecomputeStates 'recomputes the second part of split
              
              state = rbsRetracting
              t = retractTime
            End If
          ElseIf state = rbsRetracting Then
            'retracting
            If Abs(t * retractSpeed) < 0.01 Or t > 0 Then
              'retraction takes up this command as a whole (and may end with it)
                                        
              'add copy of the command to the end, changing extrusion to retraction
              Set cmd2 = New clsGCommand
              moveGroups(iGroup).chain.Add cmd2, After:=moveGroups(iGroup).lastMoveRef
              Set moveGroups(iGroup).lastMoveRef = cmd2 'note: this may potentially cause mulpiple passes of the loop, if unretraction is not possible within one loop. This may be buggy. Disabling it requires serious refactor (prediction of the situation and preventing retraction injection beforehand).
              move.Extrusion = -move.time * retractSpeed
              cmd2.RecomputeStates
              cmd2.setMove move, EError2
              cmd2.RecomputeStates
                            
              If Abs(t * retractSpeed) < 0.01 Then Exit Do 'slight under- or over-extrusion doesn't require a split
            Else
              'finalize retraction by generating a piece of current move to get the required amount
              Dim move2 As clsGMove, move3 As clsGMove
              move.split t + cmd.execTime, move2, move3
                            
              'add retract finalization command
              Set cmd2 = New clsGCommand
              moveGroups(iGroup).chain.Add cmd2, After:=moveGroups(iGroup).lastMoveRef
              Set moveGroups(iGroup).lastMoveRef = cmd2
              move2.Extrusion = -move2.time * retractSpeed
              cmd2.RecomputeStates
              cmd2.setMove move2, EError2
              cmd2.RecomputeStates
              Exit Do
            End If
          Else
            Debug.Assert False
          End If
          Set cmd = cmd.getNextMove
        Loop
        
        'recompute all states and generate new E-values for unaffected moves
        Set cmd = moveGroups(iGroup).chain.first
        Do
          cmd.RecomputeStates preserveDeltaE:=True
          Debug.Assert cmd.CompleteStateBefore.pos.X <> 0
          Debug.Assert cmd.CompleteStateAfter.pos.X <> 0
          cmd.regenerateString
          If cmd Is moveGroups(iGroup).chain.last Then Exit Do
          Set cmd = cmd.nextCommand
        Loop
        If iGroup < nMoveGroups - 1 Then
          'recreate inter-chain link that may have been lost when inserting
          'commands. This wasn't required at the time of writing this
          'comment, but =)
          chain.MakeLink cmd, moveGroups(iGroup + 1).chain.first
        End If
        moveGroups(iGroup).unretractInjected = True
        moveGroups(iGroup).retractInjected = True
        Debug.Assert moveGroups(iGroup).chain.size > 0
      End If
    End If
  Next iGroup
End If

Me.cmdProcessFile.Caption = "generating splines"
DoEvents

'replace moves with splines
For iGroup = 0 To nMoveGroups - 1
  If moveGroups(iGroup).chType = ectTravelChain Then
    Set chain = moveGroups(iGroup).chain
    'experimental: delete everything whatsoever between the build moves
    Dim mv As typTravelMoveRef
    Dim mvZero As typTravelMoveRef 'dummy variable used for clearing mv
    mv = mvZero
    
    Set mv.firstTravelMove = moveGroups(iGroup).firstMoveRef
    If moveGroups(iGroup - 1).chType = ectBuildChain Then
      Set mv.prevBuildMoveEnd = moveGroups(iGroup - 1).lastMoveRef
    Else
      Set mv.prevBuildMoveEnd = Nothing
    End If
    If iGroup < nMoveGroups - 1 Then
      If moveGroups(iGroup + 1).chType = ectBuildChain Then 'expected to be true if we got here
        Set mv.nextBuildMoveBegin = moveGroups(iGroup + 1).firstMoveRef
      End If
    End If

    If mv.nextBuildMoveBegin Is Nothing Then GoTo continue 'fixes fail on the last travel move, where there is no nex build move
    If mv.prevBuildMoveEnd Is Nothing Then GoTo continue
    moveGroups(iGroup).chain.delete
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
    gen.bRetract = Not moveGroups(iGroup - 1).retractInjected
    gen.bUnretract = Not moveGroups(iGroup + 1).unretractInjected
        
    gen.p1.copyFromT mv.prevBuildMoveEnd.CompleteStateAfter.pos
    gen.p2.copyFromT mv.nextBuildMoveBegin.CompleteStateBefore.pos
    Set gen.inSpeed = mv.prevBuildMoveEnd.getExitSpeed
    Set gen.outSpeed = mv.nextBuildMoveBegin.getEnterSpeed
    Dim arrSegments() As clsGMove
    Dim bz As clsBezier, MoveTime As Double
    Set bz = gen.FitBezier(MoveTime)
    gen.GenerateMoveTrainForBezier arrSegments, bz, MoveTime
    Dim isegment As Long
    For isegment = 0 To UBound(arrSegments)
      Set cmd = New clsGCommand
      chain.Add cmd
      If isegment = 0 Then
        'restore inter-chain connections
        chain.MakeLink moveGroups(iGroup - 1).chain.last, chain.first
        chain.MakeLink chain.last, moveGroups(iGroup + 1).chain.first
      End If
      Dim EError As Double
      EError = 0
      cmd.strLine = arrSegments(isegment).GenerateGCode(cmd.prevCommand.CompleteStateAfter, EError)
      cmd.ParseString throwIfInvalid:=True
      cmd.RecomputeStates
    Next isegment
    Debug.Assert chain.size > 0
    If timeToDoEvents Then
      Me.cmdProcessFile.Caption = "generating spline " + Str(iGroup) + " of " + Str(nMoveGroups)
      DoEvents
    End If
  End If
continue:
Next iGroup

Me.cmdProcessFile.Caption = "writing file"
DoEvents

iline = 0
Open txtFNOut For Output As f1
  For iGroup = 0 To nMoveGroups - 1
    Set chain = moveGroups(iGroup).chain
    If chain.size > 0 Then
      Set cmd = chain.first
      Do
        Print #(f1), cmd.strLine
        iline = iline + 1
        If cmd Is chain.last Then Exit Do
        Set cmd = cmd.nextCommand
      Loop
    End If
  Next iGroup
Close f1

Me.cmdProcessFile.Caption = "freeing memory"
DoEvents

For iGroup = 0 To nMoveGroups - 1
  moveGroups(iGroup).chain.delete
  If timeToDoEvents Then
    Me.cmdProcessFile.Caption = "freeing memory: move " + CStr(iGroup) + " of " + CStr(nMoveGroups)
    DoEvents
  End If
Next iGroup

Me.cmdProcessFile.Caption = "Done."

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

Private Sub cmdResetSettings_Click()
Me.ResetSettings includeFilenames:=False
End Sub

Private Sub Form_Load()
mdlPrecision.InitModule
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = QueryUnloadConstants.vbFormCode Then
  'a temporary form was created and is being destroyed, nothing special needs to be done
Else
  'explicitly terminate the application. This will prevent the app from keeping running if it is closed while processing a file
  End
End If
End Sub








Public Function GetConfigString(Optional ByVal includeFilenames As Boolean) As String
Dim ctr As Control
Dim configStr As New StringAccumulator
For Each ctr In Me
  Dim strLine As String
  strLine = ""
  
  If ctr.Tag = "!f" And Not includeFilenames Then GoTo continue
  
  If TypeOf ctr Is TextBox Then
    Dim txt As TextBox
    Set txt = ctr
    strLine = txt.Name + ".Text" + " = " + EscapeString(txt.Text)
  ElseIf TypeOf ctr Is CheckBox Then
    Dim chk As CheckBox
    Set chk = ctr
    strLine = chk.Name + ".Value" + " = " + Trim(Str(chk.Value))
  End If
  
  If Len(strLine) > 0 Then
    configStr.Append strLine + vbNewLine
  End If
continue:
Next
GetConfigString = configStr.Content
End Function

'fills the settings from the config string
Public Sub ApplyConfigStr(configStr As String, Optional ByVal suppressErrorMessages As Boolean = False, Optional ByVal includeFilenames As Boolean)
On Error GoTo eh
Dim lines() As String
lines = split(configStr, vbNewLine)
Dim i As Long
For i = 0 To UBound(lines)
  Dim strLine As String
  strLine = lines(i)
  If Not Len(strLine) > 0 Then GoTo continue
  
  Dim hs() As String
  hs = split(strLine, "=", limit:=2)
  If UBound(hs) <> 1 Then Throw errWrongConfigLine, "ApplyConfigStr", extraMessage:="Failed to split left-hand-side and right-hand-side in config line number " + Str(i)
  hs(0) = Trim(hs(0))
  'l-trim the value string, but only one space that is attached to " = "
  If Len(hs(1)) > 0 Then
    If Mid$(hs(1), 1, 1) = " " Then
      hs(1) = Mid$(hs(1), 2)
    End If
  End If
  
  Dim objprop() As String 'split object name and property name
  objprop = split(hs(0), ".", limit:=2)
  If UBound(objprop) <> 1 Then Throw errWrongConfigLine, "ApplyConfigStr", extraMessage:="Failed to split object and property in config line number " + Str(i)
  
  Dim objName As String: objName = Trim(objprop(0))
  Dim propName As String: propName = Trim(objprop(1))
  Dim obj As Control:  Set obj = Nothing
  Set obj = CallByName(Me, objName, VbGet)
  If Not obj Is Nothing Then
    If obj.Tag = "!f" And Not includeFilenames Then GoTo continue
    CallByName obj, propName, VbLet, unEscapeString(hs(1))
  End If
continue:
Next i

Exit Sub
eh:
  Dim answ As VbMsgBoxResult
  If Not suppressErrorMessages Then
    answ = MsgError(Style:=vbAbortRetryIgnore)
  Else
    answ = vbIgnore
    Debug.Print Err.source, Err.Description
  End If
  
  If answ = vbIgnore Then
    Resume continue
  ElseIf answ = vbRetry Then
    Resume
  Else
    Exit Sub
  End If
End Sub


Public Sub ResetSettings(Optional ByVal includeFilenames As Boolean = False)
Dim tmpForm As mainForm
Set tmpForm = New mainForm
Load tmpForm
On Error GoTo cleanup
Me.ApplyConfigStr tmpForm.GetConfigString(includeFilenames:=includeFilenames), includeFilenames:=includeFilenames
Unload tmpForm
Exit Sub
cleanup:
PushError
Unload tmpForm
PopError
Throw
End Sub
