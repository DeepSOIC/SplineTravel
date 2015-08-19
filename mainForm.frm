VERSION 5.00
Begin VB.Form mainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SplineTravel"
   ClientHeight    =   5460
   ClientLeft      =   30
   ClientTop       =   370
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   11730
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNotes 
      Height          =   700
      Left            =   7310
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   39
      Top             =   140
      Width           =   4220
   End
   Begin VB.Frame Frame3 
      Height          =   2330
      Left            =   7720
      TabIndex        =   31
      Top             =   1010
      Width           =   3830
      Begin VB.TextBox Text1 
         Height          =   370
         Left            =   1280
         TabIndex        =   40
         Text            =   "8"
         Top             =   1780
         Visible         =   0   'False
         Width           =   1450
      End
      Begin VB.TextBox txtLoopTol 
         Height          =   370
         Left            =   1330
         TabIndex        =   33
         Text            =   "0.3"
         Top             =   1250
         Width           =   1450
      End
      Begin VB.CheckBox chkSeamConceal 
         BackColor       =   &H00FF80FF&
         Caption         =   "seam concealment"
         Height          =   360
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   0
         Value           =   1  'Checked
         Width           =   1790
      End
      Begin VB.Label Label23 
         Caption         =   "speed of retraction"
         Height          =   470
         Left            =   240
         TabIndex        =   42
         Top             =   1710
         Visible         =   0   'False
         Width           =   1070
      End
      Begin VB.Label Label22 
         Caption         =   "mm/s"
         Height          =   240
         Left            =   2850
         TabIndex        =   41
         Top             =   1850
         Visible         =   0   'False
         Width           =   760
      End
      Begin VB.Label Label19 
         Caption         =   $"mainForm.frx":0000
         Height          =   630
         Left            =   110
         TabIndex        =   36
         Top             =   350
         Width           =   3600
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label18 
         Caption         =   "mm"
         Height          =   240
         Left            =   2900
         TabIndex        =   35
         Top             =   1320
         Width           =   760
      End
      Begin VB.Label Label17 
         Caption         =   "loop detection tolerance"
         Height          =   610
         Left            =   230
         TabIndex        =   34
         Top             =   1200
         Width           =   1070
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Smooth out travel moves"
      Height          =   3500
      Left            =   80
      TabIndex        =   9
      Top             =   890
      Width           =   7460
      Begin VB.TextBox txtAccelleration 
         Height          =   370
         Left            =   1120
         TabIndex        =   16
         Text            =   "800"
         Top             =   1610
         Width           =   1450
      End
      Begin VB.TextBox txtCurveJerk 
         Height          =   370
         Left            =   1120
         TabIndex        =   15
         Text            =   "2"
         Top             =   2200
         Width           =   1450
      End
      Begin VB.TextBox txtSpeedLimit 
         Height          =   370
         Left            =   1120
         TabIndex        =   14
         Text            =   "200"
         Top             =   990
         Width           =   1450
      End
      Begin VB.TextBox txtZJerk 
         Height          =   370
         Left            =   1120
         TabIndex        =   13
         Text            =   "0"
         Top             =   2860
         Width           =   1450
      End
      Begin VB.TextBox txtEJerk 
         Height          =   370
         Left            =   5080
         TabIndex        =   12
         Text            =   "8"
         Top             =   2220
         Width           =   1450
      End
      Begin VB.TextBox txtRetract 
         Height          =   370
         Left            =   5090
         TabIndex        =   11
         Text            =   "1.5"
         Top             =   930
         Width           =   1450
      End
      Begin VB.TextBox txtEAccell 
         Height          =   370
         Left            =   5070
         TabIndex        =   10
         Text            =   "1000"
         Top             =   1580
         Width           =   1450
      End
      Begin VB.Label Label20 
         Caption         =   $"mainForm.frx":008B
         Height          =   700
         Left            =   170
         TabIndex        =   37
         Top             =   300
         Width           =   7060
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "accelleration"
         Height          =   340
         Left            =   30
         TabIndex        =   30
         Top             =   1630
         Width           =   1210
      End
      Begin VB.Label Label4 
         Caption         =   "mm/s2"
         Height          =   240
         Left            =   2670
         TabIndex        =   29
         Top             =   1680
         Width           =   920
      End
      Begin VB.Label Label5 
         Caption         =   "curve tesellation (jerk)"
         Height          =   610
         Left            =   50
         TabIndex        =   28
         Top             =   2060
         Width           =   1070
      End
      Begin VB.Label Label6 
         Caption         =   "mm/s"
         Height          =   240
         Left            =   2670
         TabIndex        =   27
         Top             =   2270
         Width           =   760
      End
      Begin VB.Label Label7 
         Caption         =   "speed limit"
         Height          =   340
         Left            =   120
         TabIndex        =   26
         Top             =   1050
         Width           =   1210
      End
      Begin VB.Label label8 
         Caption         =   "mm/s"
         Height          =   240
         Left            =   2670
         TabIndex        =   25
         Top             =   1060
         Width           =   920
      End
      Begin VB.Label Label9 
         Caption         =   "mm/s"
         Height          =   240
         Left            =   2670
         TabIndex        =   24
         Top             =   2930
         Width           =   760
      End
      Begin VB.Label Label10 
         Caption         =   "Z jerk (for hopping)"
         Height          =   610
         Left            =   70
         TabIndex        =   23
         Top             =   2870
         Width           =   1070
      End
      Begin VB.Label Label11 
         Caption         =   "mm/s"
         Height          =   240
         Left            =   6630
         TabIndex        =   22
         Top             =   2290
         Width           =   760
      End
      Begin VB.Label Label12 
         Caption         =   "E jerk (for retraction)"
         Height          =   610
         Left            =   3910
         TabIndex        =   21
         Top             =   2180
         Width           =   1070
      End
      Begin VB.Label Label13 
         Caption         =   "mm"
         Height          =   240
         Left            =   6640
         TabIndex        =   20
         Top             =   1000
         Width           =   760
      End
      Begin VB.Label Label14 
         Caption         =   "retract"
         Height          =   610
         Left            =   3990
         TabIndex        =   19
         Top             =   970
         Width           =   1070
      End
      Begin VB.Label Label15 
         Caption         =   "mm/s2"
         Height          =   240
         Left            =   6620
         TabIndex        =   18
         Top             =   1650
         Width           =   760
      End
      Begin VB.Label Label16 
         Caption         =   "E accelleration"
         Height          =   610
         Left            =   3890
         TabIndex        =   17
         Top             =   1610
         Width           =   1070
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Presets"
      Height          =   730
      Left            =   20
      TabIndex        =   5
      Top             =   40
      Width           =   6290
      Begin VB.CommandButton cmdResetSettings 
         BackColor       =   &H008080FF&
         Caption         =   "*"
         Height          =   280
         Left            =   5890
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "load defaults"
         Top             =   390
         Visible         =   0   'False
         Width           =   340
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H008080FF&
         Caption         =   "Delete"
         Height          =   280
         Left            =   5450
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   270
         Width           =   720
      End
      Begin VB.ComboBox cmbPreset 
         Height          =   280
         ItemData        =   "mainForm.frx":015B
         Left            =   100
         List            =   "mainForm.frx":015D
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   270
         Width           =   4110
      End
      Begin VB.CommandButton cmdSaveAs 
         BackColor       =   &H0080FFFF&
         Caption         =   "Save as..."
         Height          =   280
         Left            =   4330
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   270
         Width           =   1060
      End
   End
   Begin VB.CommandButton cmdProcessFile 
      BackColor       =   &H0080FF80&
      Caption         =   "Go"
      CausesValidation=   0   'False
      Height          =   840
      Left            =   7930
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4470
      Width           =   3630
   End
   Begin VB.TextBox txtFNOut 
      Height          =   410
      Left            =   1540
      TabIndex        =   3
      Tag             =   "!f"
      Top             =   4930
      Width           =   5590
   End
   Begin VB.TextBox txtFNIn 
      Height          =   360
      Left            =   1530
      TabIndex        =   0
      Tag             =   "!f"
      Top             =   4500
      Width           =   5610
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      Caption         =   "notes"
      Height          =   230
      Left            =   6270
      TabIndex        =   38
      Top             =   330
      Width           =   940
   End
   Begin VB.Label Label2 
      Caption         =   "output"
      Height          =   380
      Left            =   350
      TabIndex        =   2
      Top             =   4960
      Width           =   1090
   End
   Begin VB.Label Label1 
      Caption         =   "input"
      Height          =   270
      Left            =   340
      TabIndex        =   1
      Top             =   4520
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

Private Type typPresetManagerGlobals
  filesInCombo() As String 'full paths
  block As New clsBlokada  'locks response to events, for when filling preset list
  curPresetIsModified As Boolean
  curPresetFN As String
End Type
Dim pm As typPresetManagerGlobals

Private Sub chkSeamConceal_Click()
ChangeWasMade
End Sub

Private Sub cmbPreset_Click()
If pm.block Then Exit Sub
On Error GoTo eh
If Me.getSelectedPreset = pm.curPresetFN Then Exit Sub  'no change
If pm.curPresetIsModified Then
  Dim answ As VbMsgBoxResult
  answ = MsgBox("Current preset was modified. The changes will be lost. Continue anyway?", vbYesNo Or vbDefaultButton2)
  If answ = vbNo Then
    Me.SelectPreset pm.curPresetFN
    Exit Sub
  End If
End If
Me.LoadPreset Me.getSelectedPreset
Exit Sub
eh:
MsgError
Me.SelectPreset pm.curPresetFN
End Sub

Private Sub cmdDelete_Click()
On Error GoTo eh
If Len(pm.curPresetFN) = 0 Then Throw errInvalidArgument, , "No preset is selected, can't delete"
Dim answ As VbMsgBoxResult
answ = MsgBox("You are about to delete preset " + getFileTitle(pm.curPresetFN) + ", stored in " + pm.curPresetFN + ". Continue?", vbYesNo Or vbDefaultButton2)
If answ = vbNo Then Throw errCancel
Kill pm.curPresetFN
MsgBox "Preset " + getFileTitle(pm.curPresetFN) + " was deleted."
Me.purgeModified
pm.curPresetFN = ""
Me.RefillPresets
Exit Sub
eh:
MsgError
End Sub

Private Sub cmdProcessFile_Click()
cmdProcessFile.Enabled = False
Dim f1 As Long
f1 = FreeFile
Dim chain As New clsChain
Dim iline As Long
On Error GoTo eh
SavePreset "(last used)"
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
      p1 = moveGroups(iGroup).firstMoveRef.CompleteStateBefore.Pos
      p2 = moveGroups(iGroup).lastMoveRef.CompleteStateAfter.Pos
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
          Debug.Assert cmd.CompleteStateBefore.Pos.X <> 0
          Debug.Assert cmd.CompleteStateAfter.Pos.X <> 0
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
        
    gen.p1.copyFromT mv.prevBuildMoveEnd.CompleteStateAfter.Pos
    gen.p2.copyFromT mv.nextBuildMoveBegin.CompleteStateBefore.Pos
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

Private Sub cmdSaveAs_Click()
On Error GoTo eh
Dim presetName As String
presetName = getFileTitle(pm.curPresetFN)
Dim newPresetName As String
again:
newPresetName = InputBox("Name the preset. Entering a new name will keep current preset intact.", presetName)
If Len(newPresetName) = 0 Then Exit Sub
verifyPresetName newPresetName
Dim newPresetFilename As String
If StrComp(newPresetName, presetName, vbTextCompare) = 0 Then
  newPresetFilename = pm.curPresetFN
Else
  Dim paths() As String
  paths = mdlFiles.PresetsPaths
  newPresetFilename = paths(UBound(paths)) + newPresetName + ".ini"
  If Me.FindInCombo(newPresetFilename) <> -1 Then
    If MsgBox("Overwriting another existing preset. Continue?", vbYesNo Or vbDefaultButton2) = vbNo Then
      Exit Sub
    End If
  End If
End If
WritePresetFile newPresetFilename, Me.GetConfigString
pm.curPresetFN = newPresetFilename
pm.curPresetIsModified = False
Me.RefillPresets
Exit Sub
eh:
MsgError

End Sub

Public Sub verifyPresetName(ByRef newName As String)
Dim i As Long
Dim ch As String
Const allowedCharacters As String = " ,.-+=!()'1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
newName = Trim(newName)
If Len(newName) = 0 Then Throw errInvalidArgument, extraMessage:="preset name is empty"
For i = 1 To Len(newName)
  ch = Mid$(newName, i, 1)
  If InStr(allowedCharacters, ch) = 0 Then
    Throw errInvalidArgument, extraMessage:="preset name contains an invalid character " + ch
  End If
Next i
End Sub

Private Sub Form_Activate()
Me.RefillPresets
End Sub

Private Sub Form_Load()
mdlPrecision.InitModule
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = QueryUnloadConstants.vbFormCode Then
  'a temporary form was created and is being destroyed, nothing special needs to be done
Else
  SavePreset "(since last close)"
  'explicitly terminate the application. This will prevent the app from keeping running if it is closed while processing a file
  End
End If
End Sub

'a shortcut. It doesn't update current preset, it merely writes a new preset file
Public Sub SavePreset(presetName As String)
Dim newPresetFilename As String
Dim paths() As String
paths = mdlFiles.PresetsPaths
newPresetFilename = paths(UBound(paths)) + presetName + ".ini"
WritePresetFile newPresetFilename, Me.GetConfigString
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
GetConfigString = configStr.content
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
    Debug.Print Err.Source, Err.Description
  End If
  
  If answ = vbIgnore Then
    Resume continue
  ElseIf answ = vbRetry Then
    Resume
  Else
    Throw errCancel
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

Public Sub RefillPresets()
Dim keeper As clsBlokada: Set keeper = pm.block.block

pm.filesInCombo = mdlFiles.getListOfFiles(mdlFiles.PresetsPaths, "*.ini")
cmbPreset.Clear
Dim i As Long
For i = 0 To ArrLen(pm.filesInCombo) - 1
  cmbPreset.AddItem mdlFiles.getFileTitle(pm.filesInCombo(i))
  cmbPreset.ItemData(cmbPreset.NewIndex) = i
  If pm.curPresetFN = pm.filesInCombo(i) And pm.curPresetIsModified Then
    cmbPreset.List(cmbPreset.NewIndex) = cmbPreset.List(cmbPreset.NewIndex) + "*"
  End If
Next i

Me.SelectPreset pm.curPresetFN

keeper.Unblock
End Sub

'returns -1 if item not found
Public Function FindInCombo(presetFilePath As String) As Long
FindInCombo = -1
Debug.Assert cmbPreset.ListCount = ArrLen(pm.filesInCombo)
Dim i As Long
For i = 0 To cmbPreset.ListCount - 1
  If StrComp(pm.filesInCombo(cmbPreset.ItemData(i)), presetFilePath, vbTextCompare) = 0 Then
    FindInCombo = i
  End If
Next i
End Function

Private Function ArrLen(arr As Variant) As Long
Dim ub As Long: ub = -1
On Error Resume Next
ub = UBound(arr)
ArrLen = ub + 1
End Function

Public Sub LoadPreset(FilePath As String)
Dim tmp As String
tmp = ReadPresetFile(FilePath)
Dim keeper As clsBlokada: Set keeper = pm.block.block
Me.ResetSettings
Me.ApplyConfigStr tmp
Me.purgeModified
Me.SelectPreset FilePath
pm.curPresetFN = FilePath
keeper.Unblock
End Sub

Public Function ReadPresetFile(FilePath As String) As String
Dim f As Long: f = FreeFile
Open FilePath For Input As f
On Error GoTo cleanup
  ReadPresetFile = Input$(LOF(f), #(f))
Close f
Exit Function
cleanup:
  PushError
  Close f
  PopError
  Throw
End Function

Public Sub WritePresetFile(FilePath As String, content As String)
Dim f As Long: f = FreeFile
Open FilePath For Output As f
On Error GoTo cleanup
  Print #(f), content;
Close f
Exit Sub
cleanup:
  PushError
  Close f
  PopError
  Throw
End Sub

Public Function getSelectedPreset(Optional ByVal index As Long = -2) As String
If index = -2 Then index = cmbPreset.ListIndex
If index = -1 Then getSelectedPreset = "": Exit Function
Debug.Assert ArrLen(pm.filesInCombo) = cmbPreset.ListCount
getSelectedPreset = pm.filesInCombo(cmbPreset.ItemData(index))
End Function

'only selects; no file is opened
Public Function SelectPreset(presetFilePath As String) As Long
Dim idx As Long
idx = Me.FindInCombo(presetFilePath)
If Len(presetFilePath) > 0 And idx = -1 Then
  'preset isn't in the list, add it...
  Me.addPresetToList presetFilePath
  idx = Me.FindInCombo(presetFilePath)
  Debug.Assert idx <> -1
End If
Dim keeper As clsBlokada: Set keeper = pm.block.block
cmbPreset.ListIndex = idx
keeper.Unblock
End Function

Public Function addPresetToList(presetFilePath As String, Optional ByVal checkIfExists As Boolean = False) As Long
If checkIfExists Then
  If Me.FindInCombo(presetFilePath) <> -1 Then
    'already in list!
    Exit Function
  End If
End If
If ArrLen(pm.filesInCombo) = 0 Then ReDim pm.filesInCombo(0 To 0) Else ReDim Preserve pm.filesInCombo(0 To UBound(pm.filesInCombo) + 1)
Dim i As Long: i = UBound(pm.filesInCombo)
pm.filesInCombo(i) = presetFilePath
cmbPreset.AddItem mdlFiles.getFileTitle(pm.filesInCombo(i))
cmbPreset.ItemData(cmbPreset.NewIndex) = i
End Function

Public Sub ChangeWasMade()
If pm.block Then Exit Sub
If cmbPreset.ListIndex = -1 Then Exit Sub
pm.curPresetIsModified = True
Dim txt As String
txt = cmbPreset.List(cmbPreset.ListIndex)
If right$(txt, 1) <> "*" Then
  txt = txt + "*"
  cmbPreset.List(cmbPreset.ListIndex) = txt
End If

End Sub

Private Sub txtAccelleration_Change()
ChangeWasMade
End Sub

Private Sub txtCurveJerk_Change()
ChangeWasMade
End Sub

Private Sub txtEAccell_Change()
ChangeWasMade
End Sub

Private Sub txtEJerk_Change()
ChangeWasMade
End Sub

Private Sub txtLoopTol_Change()
ChangeWasMade
End Sub

Private Sub txtRetract_Change()
ChangeWasMade
End Sub

Private Sub txtSpeedLimit_Change()
ChangeWasMade
End Sub

Private Sub txtZJerk_Change()
ChangeWasMade
End Sub

Public Sub purgeModified(Optional ByVal index As Long = -2)
If index = -2 Then
  index = Me.FindInCombo(pm.curPresetFN)
  pm.curPresetIsModified = False
End If
If index = -1 Then Exit Sub
Dim txt As String
txt = cmbPreset.List(index)
If right$(txt, 1) = "*" Then
  txt = Left$(txt, Len(txt) - 1)
  cmbPreset.List(index) = txt
End If
End Sub
