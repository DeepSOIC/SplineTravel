Attribute VB_Name = "mdlWorker"
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

Private Type typThis
  cfg As mainForm
End Type
Private this As typThis


Public Sub Process(FNIn As String, FNOut As String, cfg As mainForm)
On Error GoTo eh
Set this.cfg = cfg

'read file
Dim chain As clsChain
Set chain = ReadGCodeFile(FNIn)
Dim iline As Long, nLines As Long
nLines = chain.size

cfg.cmdProcessFile.Caption = "searching for travel moves"
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
    cfg.cmdProcessFile.Caption = "line " + Str(iline) + " of " + Str(nLines)
    DoEvents
  End If
Loop

cfg.cmdProcessFile.Caption = "splitting"
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

If cfg.chkSeamConceal.Value = vbChecked Then
  cfg.cmdProcessFile.Caption = "blending seams"
  DoEvents
  
  Dim loopTol As Double
  loopTol = val(cfg.txtLoopTol)
  Dim retractTime As Double
  retractTime = val(cfg.txtRetract) / val(cfg.txtEJerk)
  Dim retractSpeed As Double
  retractSpeed = val(cfg.txtRSpeedSC)
  
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

Dim bSplineTravel As Boolean: bSplineTravel = cfg.optTravelSpline.Value
Dim bStraightTravel As Boolean: bStraightTravel = cfg.optTravelStraight.Value

cfg.cmdProcessFile.Caption = "generating splines"
DoEvents

Dim retractLength As Double: retractLength = val(cfg.txtRetract)
Dim ZHop As Double: ZHop = val(cfg.txtZHop)

'replace moves with splines
For iGroup = 0 To nMoveGroups - 1
  If moveGroups(iGroup).chType = ectTravelChain Then
    Set chain = moveGroups(iGroup).chain
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
    If bSplineTravel Then
      'experimental: delete everything whatsoever between the build moves. This may also accidentally consume some important commands...
      moveGroups(iGroup).chain.delete
      Dim gen As clsTravelGenerator
      If gen Is Nothing Then
        Set gen = New clsTravelGenerator
        gen.acceleration = val(cfg.txtAcceleration)
        gen.CurveJerk = val(cfg.txtCurveJerk)
        gen.speedLimit = val(cfg.txtSpeedLimit)
        gen.Retract = val(cfg.txtRetract)
        gen.RetractAcceleration = val(cfg.txtEAccel)
        gen.RetractJerk = val(cfg.txtEJerk)
        gen.ZJerk = val(cfg.txtZJerk)
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
    
    ElseIf bStraightTravel Then
    
      'generate straight travel move
      
      'FIXME: keep old non-move commands
      chain.delete
      
      Dim bRetract As Boolean, bUnretract As Boolean
      bRetract = Not moveGroups(iGroup - 1).retractInjected
      bUnretract = Not moveGroups(iGroup + 1).unretractInjected
      If retractLength = 0 Then bRetract = False: bUnretract = False
      If bRetract Then
        Set cmd = New clsGCommand
        chain.Add cmd
        chain.MakeLink mv.prevBuildMoveEnd.inChain.last, chain.first
        cmd.RecomputeStates
        
        Set move = New clsGMove
        move.Extrusion = -retractLength
        move.ExtrusionSpeed = val(cfg.txtRSpeedStraight)
        move.p1.copyFromT cmd.CompleteStateBefore.Pos
        move.p2.copyFromT cmd.CompleteStateBefore.Pos
      
        cmd.setMove move
        cmd.RecomputeStates
      End If
      
      If ZHop Then
        Set cmd = New clsGCommand
        chain.Add cmd
        chain.MakeLink mv.prevBuildMoveEnd.inChain.last, chain.first
        cmd.RecomputeStates
        
        Set move = New clsGMove
        move.p1.copyFromT cmd.CompleteStateBefore.Pos
        move.p2.copyFromT cmd.CompleteStateBefore.Pos
        move.p2.Z = move.p1.Z + ZHop
        move.Speed = val(cfg.txtSpeedStraight)
        
        cmd.setMove move
        cmd.RecomputeStates
      End If
      
      Set cmd = New clsGCommand
      chain.Add cmd
      chain.MakeLink mv.prevBuildMoveEnd.inChain.last, chain.first
      cmd.RecomputeStates
      
      Set move = New clsGMove
      move.p1.copyFromT cmd.CompleteStateBefore.Pos
      move.p2.copyFromT mv.nextBuildMoveBegin.CompleteStateBefore.Pos
      move.p2.Z = move.p2.Z + ZHop
      move.Speed = val(cfg.txtSpeedStraight)
      
      cmd.setMove move
      cmd.RecomputeStates
      
      If ZHop Then
        Set cmd = New clsGCommand
        chain.Add cmd
        cmd.RecomputeStates
        
        Set move = New clsGMove
        move.p1.copyFromT cmd.CompleteStateBefore.Pos
        move.p2.copyFromT mv.nextBuildMoveBegin.CompleteStateBefore.Pos
        move.Speed = val(cfg.txtSpeedStraight)
        
        cmd.setMove move
        cmd.RecomputeStates
      End If
      
      If bUnretract Then
        Set cmd = New clsGCommand
        chain.Add cmd
        cmd.RecomputeStates
        
        Set move = New clsGMove
        move.Extrusion = retractLength
        move.ExtrusionSpeed = val(cfg.txtRSpeedStraight)
        move.p1.copyFromT cmd.CompleteStateBefore.Pos
        move.p2.copyFromT cmd.CompleteStateBefore.Pos
      
        cmd.setMove move
        cmd.RecomputeStates
      End If
      
      chain.MakeLink chain.last, mv.nextBuildMoveBegin.inChain.first
      
    End If
    Debug.Assert chain.size > 0
    If timeToDoEvents Then
      cfg.cmdProcessFile.Caption = "generating spline " + Str(iGroup) + " of " + Str(nMoveGroups)
      DoEvents
    End If
  End If
continue:
Next iGroup

cfg.cmdProcessFile.Caption = "writing file"
DoEvents

iline = 0
Dim f1 As Long: f1 = FreeFile
Open FNOut For Output As f1
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

GoSub cleanup

cfg.cmdProcessFile.Caption = "Done."

Exit Sub
eh:
  MsgError
  PushError
  Close f1
  GoSub cleanup
  PopError
  Throw
Exit Sub

cleanup:
  cfg.cmdProcessFile.Caption = "freeing memory"
  DoEvents
  
  For iGroup = 0 To nMoveGroups - 1
    moveGroups(iGroup).chain.delete
    If timeToDoEvents Then
      cfg.cmdProcessFile.Caption = "freeing memory: move " + CStr(iGroup) + " of " + CStr(nMoveGroups)
      DoEvents
    End If
  Next iGroup
Return

End Sub

Public Function timeToDoEvents()
Static lastDidTime As Double
If Abs(Timer - lastDidTime) > 0.3 Then
  timeToDoEvents = True
  lastDidTime = Timer
End If
End Function

Public Function ReadGCodeFile(path As String) As clsChain
Dim f1 As Long
f1 = FreeFile
Dim chain As New clsChain
Dim iline As Long
On Error GoTo cleanup
Open path For Input As f1
  Dim ln As String
  Do While Not EOF(f1)
    Line Input #(f1), ln
    chain.Add New clsGCommand
    chain.last.strLine = ln
    chain.last.ParseString
    chain.last.RecomputeStates
    iline = iline + 1
    If timeToDoEvents Then
      If Not this.cfg Is Nothing Then
        this.cfg.cmdProcessFile.Caption = "reading line " + Str(iline)
      End If
      DoEvents
    End If
  Loop
Close f1
Set ReadGCodeFile = chain
Exit Function
cleanup:
  PushError
  Close f1
  chain.delete
  PopError
  Throw
End Function
