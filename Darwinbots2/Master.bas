Attribute VB_Name = "Master"
Option Explicit
Public DynamicCountdown As Integer ' Used to countdown the cycles until we modify the dynamic costs
Public CostsWereZeroed As Boolean ' Flag used to indicate to the reinstatement threshodl that the costs were zeroed
Public PopulationLast10Cycles(10) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer 'Botsareus 12/11/2012 Pause Break Key to Pause code
'

Public energydif As Double 'Total last hide pred
Public energydifX As Double 'Avg last hide pred
Public energydifXP As Double 'The actual handycap

Public energydif2 As Double 'Total last hide pred
Public energydifX2 As Double 'Avg last hide pred
Public energydifXP2 As Double 'The actual handycap

Public stopflag As Boolean

Public savenow As Boolean

Public stagnent As Boolean

Public Sub UpdateSim()

  'core evo
  Dim avrnrgStart As Double
  Dim avrnrgEnd As Double

  Dim AmountOff As Single
  Dim UpperRange As Single
  Dim LowerRange As Single
  Dim CorrectionAmount As Single
  Dim CurrentPopulation As Integer
  Dim AllChlr As Long 'Panda 8/13/2013 The new way to figure out total number vegys
  Dim i As Integer
  Dim t As Integer
  
  Dim Base_count As Integer
  Dim Mutate_count As Integer

  'Botsareus 12/11/2012 Pause Break Key to Pause code
  If GetAsyncKeyState(vbKeyF12) Then
      DisplayActivations = False
      Form1.Active = False
      Form1.SecTimer.Enabled = False
      MDIForm1.unpause.Enabled = True
  End If
  
  ModeChangeCycles = ModeChangeCycles + 1
  SimOpts.TotRunCycle = SimOpts.TotRunCycle + 1
  
  'Botsareus 3/22/2014 Main hidepred logic (hide pred means hide base robot a.k.a. Predator)
  Dim usehidepred As Boolean
  '
  'Dim avgsize As Long
  Dim k As Long 'robots moved last attempt
  Dim k2 As Long 'robots moved total
  Dim ingdist As Single
  Dim pozdif As vector
  Dim newpoz As vector
  Dim posdif As vector
  '
  If usehidepred Then
    'Count species for end of evo
    Base_count = 0
    Mutate_count = 0
    For t = 1 To MaxRobs
      If rob(t).exist Then
          If rob(t).FName = "Base.txt" Then Base_count = Base_count + 1
          If rob(t).FName = "Mutate.txt" Then Mutate_count = Mutate_count + 1
      End If
    Next t
    If Base_count > Mutate_count Then stagnent = False 'Botsareus 10/20/2015 Base went above mutate, reset stagnent flag
    'See if end of evo
    If Mutate_count = 0 Then
        DisplayActivations = False
        Form1.Active = False
        Form1.SecTimer.Enabled = False
        stopflag = True 'Botsareus 9/2/2014 A bug fix from Spork22
    End If
    If Base_count = 0 And Not stopflag Then
        DisplayActivations = False
        Form1.Active = False
        Form1.SecTimer.Enabled = False
    End If
    'Botsareus 10/19/2015 Prevents simulation from running too long
    If SimOpts.TotRunCycle = 1000000 Then stagnent = True 'Set the stagnent flag now and see what happens
'    If SimOpts.TotRunCycle = 3000000 And stagnent Then 'Botsareus 1/9/2016 Rule no longer required as I am no longer evolving plants
'        DisplayActivations = False
'        Form1.Active = False
'        Form1.SecTimer.Enabled = False
'        UpdateWonEvo Form1.fittest
'    End If
    
Mode:
    If ModeChangeCycles > (hidePredCycl / 1.2 + hidePredOffset) Then
      'Botsareus 11/5/2015 If lfor max lower limit wait for mutate pop to match base pop
      If LFOR = 150 And Mutate_count < Base_count And hidepred Then
            ModeChangeCycles = ModeChangeCycles - 100
            GoTo Mode
      End If
      'calculate new energy handycap
      energydif2 = energydif2 + energydif / ModeChangeCycles 'inverse handycap
      If hidepred Then
        Dim holdXP As Double
        holdXP = (energydifX - (energydif / ModeChangeCycles)) / LFOR
        If holdXP < energydifXP Then energydifXP = holdXP Else energydifXP = (energydifXP * 9 + holdXP) / 10
        
        'inverse handycap
        energydifXP2 = (energydifX2 - energydif2) / LFOR
        If energydifXP2 > 0 Then energydifXP2 = 0
        If (energydifXP - energydifXP2) > 0.1 Then energydifXP2 = energydifXP - 0.1
        energydifX2 = energydif2
        energydif2 = 0
      End If
      energydifX = energydif / ModeChangeCycles
      energydif = 0
      'Botsareus 6/12/2016 An attempt to get rid of 'chasers' without using any reposition code:
      If hidepred Then
      
        'Erase offensive shots
        For t = 1 To maxshotarray
          With Shots(t)
            If .shottype = -1 Or .shottype = -6 Then
                .exist = False
                .flash = False
            End If
          End With
        Next t
        
        'Reposition robots the safe way
        k2 = 0
        Do
        k = 0
        For t = 1 To MaxRobs
             If rob(t).exist And rob(t).FName = "Base.txt" Then
                  For i = 1 To MaxRobs
                      If rob(i).exist And rob(i).FName = "Mutate.txt" Then
                      
                        'calculate ingagment distance
                        If rob(t).body > rob(i).body Then
                            If rob(t).body > 10 Then
                                ingdist = Log(rob(t).body) * 60 + 41
                            Else
                                ingdist = 40
                            End If
                        Else
                            If rob(i).body > 10 Then
                                ingdist = Log(rob(i).body) * 60 + 41
                            Else
                                ingdist = 40
                            End If
                        End If
                        ingdist = rob(t).radius + rob(i).radius + ingdist + 40 'both radii plus shot dist plus offset 1 shot travel dist
                            
                        posdif = VectorSub(rob(t).pos, rob(i).pos)
                        If VectorMagnitude(posdif) < ingdist Then
                        'if the distance between the robots is less then ingagment distance
                            With rob(i)
                                ingdist = ingdist - VectorMagnitude(posdif) 'ingdist becomes offset dist
                                newpoz = VectorSub(.pos, VectorScalar(VectorUnit(posdif), ingdist)) 'offset the multibot by ingagment distance
                                
                                pozdif.X = newpoz.X - .pos.X
                                pozdif.Y = newpoz.Y - .pos.Y
                                If .numties > 0 Then
                                    Dim clist(50) As Integer, tk As Integer
                                    clist(0) = i
                                    ListCells clist()
                                    'move multibot
                                    tk = 1
                                    While clist(tk) > 0
                                        'Botsareus 7/15/2016 Only own species
                                        If rob(clist(tk)).FName = "Mutate.txt" Then
                                            rob(clist(tk)).pos.X = rob(clist(tk)).pos.X + pozdif.X
                                            rob(clist(tk)).pos.Y = rob(clist(tk)).pos.Y + pozdif.Y
                                        End If
                                        tk = tk + 1
                                    Wend
                                End If
                                .pos.X = .pos.X + pozdif.X
                                .pos.Y = .pos.Y + pozdif.Y
                            End With
                            k = k + 1
                            k2 = k2 + 1
                        End If
                      End If
                  Next
             End If
        Next t
        Loop Until k = 0 Or k2 > (3200 + Mutate_count * 0.9) 'Scales as mutate_count scales
        
      End If
      'change hide pred
      hidepred = Not hidepred
      hidePredOffset = hidePredCycl / 3 * rndy
      ModeChangeCycles = 0
    End If
  End If
  
  'provides the mutation rates oscillation Botsareus 8/3/2013 moved to UpdateSim)
  Dim fullrange As Long
  If SimOpts.MutOscill Then 'Botsareus 10/8/2015 Yet another redo, sine wave optional
   With SimOpts
   
    If (.MutCycMax + .MutCycMin) > 0 Then
    
     If .MutOscillSine Then
     
        fullrange = .TotRunCycle Mod (.MutCycMax + .MutCycMin)
        If fullrange < .MutCycMax Then
         .MutCurrMult = 20 ^ Sin(fullrange / .MutCycMax * PI)
        Else
         .MutCurrMult = 20 ^ -Sin((fullrange - .MutCycMax) / .MutCycMin * PI)
        End If
     
     Else
     
        fullrange = .TotRunCycle Mod (.MutCycMax + .MutCycMin)
        If fullrange < .MutCycMax Then
         .MutCurrMult = 16
        Else
         .MutCurrMult = 1 / 16
        End If
    
     End If
    
    End If
    
   End With
  End If
  

  TotalSimEnergyDisplayed = TotalSimEnergy(CurrentEnergyCycle)
  CurrentEnergyCycle = SimOpts.TotRunCycle Mod 100
  TotalSimEnergy(CurrentEnergyCycle) = 0
  
  CurrentPopulation = totnvegsDisplayed
  If SimOpts.Costs(DYNAMICCOSTINCLUDEPLANTS) <> 0 Then
    CurrentPopulation = CurrentPopulation + totvegsDisplayed      'Include Plants in target population
  End If
  
  'If (SimOpts.TotRunCycle + 200) Mod 2000 = 0 Then MsgBox "sup" & SimOpts.TotRunCycle 'debug only
  
  If SimOpts.TotRunCycle Mod 10 = 0 Then
    For i = 10 To 2 Step -1
      PopulationLast10Cycles(i) = PopulationLast10Cycles(i - 1)
    Next i
    PopulationLast10Cycles(1) = CurrentPopulation
  End If
    
  If SimOpts.Costs(USEDYNAMICCOSTS) Then
    AmountOff = CurrentPopulation - SimOpts.Costs(DYNAMICCOSTTARGET)
         
    'If we are more than X% off of our target population either way AND the population isn't moving in the
    'the direction we want or hasn't moved at all in the past 10 cycles then adjust the cost multiplier
    UpperRange = TmpOpts.Costs(DYNAMICCOSTTARGETUPPERRANGE) * 0.01 * SimOpts.Costs(DYNAMICCOSTTARGET)
    LowerRange = TmpOpts.Costs(DYNAMICCOSTTARGETLOWERRANGE) * 0.01 * SimOpts.Costs(DYNAMICCOSTTARGET)
    If (CurrentPopulation = PopulationLast10Cycles(10)) Then
      DynamicCountdown = DynamicCountdown - 1
      If DynamicCountdown < -10 Then DynamicCountdown = -10
    Else
      DynamicCountdown = 10
    End If
    
    If (AmountOff > UpperRange And (PopulationLast10Cycles(10) < CurrentPopulation Or DynamicCountdown <= 0)) Or _
       (AmountOff < -LowerRange And (PopulationLast10Cycles(10) > CurrentPopulation Or DynamicCountdown <= 0)) Then
       
       If AmountOff > UpperRange Then
         CorrectionAmount = AmountOff - UpperRange
       Else
         CorrectionAmount = Abs(AmountOff) - LowerRange
       End If
       
      'Adjust the multiplier. The idea is to rachet this over time as bots evolve to be more effecient.
      'We don't muck with it if the bots are within X% of the target.  If they are outside the target, then
      'we adjust only if the populatiuon isn't heading towards the range and then we do it my an amount that is a function
      'of how far out of the range we are (not how far from the target itself) and the sensitivity set in the sim
      SimOpts.Costs(COSTMULTIPLIER) = SimOpts.Costs(COSTMULTIPLIER) + (0.0000001 * CorrectionAmount * Sgn(AmountOff) * SimOpts.Costs(DYNAMICCOSTSENSITIVITY))
      
      'Don't let the costs go negative if the user doesn't want them to
      If (SimOpts.Costs(ALLOWNEGATIVECOSTX) <> 1) Then
        If SimOpts.Costs(COSTMULTIPLIER) < 0 Then SimOpts.Costs(COSTMULTIPLIER) = 0
      End If
      DynamicCountdown = 10 ' Reset the countdown timer
    End If
 ' Else
 '   SimOpts.Costs(COSTMULTIPLIER) = 1
  End If
  
  If (CurrentPopulation < SimOpts.Costs(BOTNOCOSTLEVEL)) And (SimOpts.Costs(COSTMULTIPLIER) <> 0) Then
    CostsWereZeroed = True
    SimOpts.oldCostX = SimOpts.Costs(COSTMULTIPLIER)
    SimOpts.Costs(COSTMULTIPLIER) = 0 ' The population has fallen below the threshold to 0 all costs
  ElseIf (CurrentPopulation > SimOpts.Costs(COSTXREINSTATEMENTLEVEL)) And CostsWereZeroed Then
    CostsWereZeroed = False ' Set the flag so we don't do this again unless they get zeored again
    SimOpts.Costs(COSTMULTIPLIER) = SimOpts.oldCostX
  End If

  If usehidepred Then
  'Calculate average energy before sim update
  avrnrgStart = 0
  i = 0
  For t = 1 To MaxRobs
        If rob(t).FName = "Mutate.txt" And rob(t).exist Then
            If rob(t).LastMut > 0 Then '4/17/2014 New rule from Botsareus, only handycap fresh robots
                i = i + 1
                avrnrgStart = avrnrgStart + rob(t).nrg
            End If
        End If
  Next t
   If i > 0 Then
    avrnrgStart = avrnrgStart / i
   End If
  End If
  
  ExecRobs
  If datirob.Visible And datirob.ShowMemoryEarlyCycle Then
    With rob(robfocus)
      datirob.infoupdate robfocus, .nrg, .parent, .Mutations, .age, .SonNumber, 1, .FName, .genenum, .LastMut, .generation, .DnaLen, .LastOwner, .Waste, .body, .mass, .venom, .shell, .Slime, .chloroplasts
    End With
  End If
  
  'updateshots can write to bot sense, so we need to clear bot senses before updating shots
  For t = 1 To MaxRobs
    If rob(t).exist Then
      If (rob(t).DisableDNA = False) Then EraseSenses t
    End If
  Next t

  updateshots
  
    'Botsareus 6/22/2016 to figure out actual velocity of the bot incase there is a collision event
    For t = 1 To MaxRobs
        If rob(t).exist And Not (rob(t).FName = "Base.txt" And hidepred) Then
            rob(t).opos = rob(t).pos
        End If
    Next
  
  UpdateBots
  
    'to figure out actual velocity of the bot incase there is a collision event
    For t = 1 To MaxRobs
        If rob(t).exist And Not (rob(t).FName = "Base.txt" And hidepred) Then
            'Only if the robots position was already configured
            If Not (rob(t).opos.x = 0 And rob(t).opos.y = 0) Then rob(t).actvel = VectorSub(rob(t).pos, rob(t).opos)
        End If
    Next
  
  If numObstacles > 0 Then MoveObstacles
   
  For t = 1 To MaxRobs 'Panda 8/14/2013 to figure out how much vegys to repopulate across all robots
   If rob(t).exist And Not (rob(t).FName = "Base.txt" And hidepred) Then 'Botsareus 8/14/2013 We have to make sure the robot is alive first
    AllChlr = AllChlr + rob(t).chloroplasts
   End If
  Next t
  
  TotalChlr = AllChlr / 16000 'Panda 8/23/2013 Calculate total unit chloroplasts
  
  If TotalChlr < CLng(SimOpts.MinVegs) Then   'Panda 8/23/2013 Only repopulate vegs when total chlroplasts below value
    If totvegsDisplayed <> -1 Then VegsRepopulate  'Will be -1 first cycle after loading a sim.  Prevents spikes.
  End If
  
  feedvegs SimOpts.MaxEnergy
  
  If usehidepred Then
  'Calculate average energy after sim update
  avrnrgEnd = 0
  i = 0
  For t = 1 To MaxRobs
        If rob(t).FName = "Mutate.txt" And rob(t).exist Then
            If rob(t).LastMut > 0 Then '4/17/2014 New rule from Botsareus, only handycap fresh robots
                i = i + 1
                avrnrgEnd = avrnrgEnd + rob(t).nrg
            End If
        End If
  Next t
   If i > 0 Then
    avrnrgEnd = avrnrgEnd / i
    energydif = energydif - avrnrgStart + avrnrgEnd
   End If
  End If
  
  'okay, time to store some values for RGB monitor
  If MDIForm1.MonitorOn Then
    For t = 1 To MaxRobs
      If rob(t).exist Then
       With frmMonitorSet
        rob(t).monitor_r = rob(t).mem(.Monitor_mem_r)
        rob(t).monitor_g = rob(t).mem(.Monitor_mem_g)
        rob(t).monitor_b = rob(t).mem(.Monitor_mem_b)
       End With
      End If
    Next t
  End If
  
  'Kill some robots to prevent out of memory
  Dim totlen As Long
  totlen = 0
  For t = 1 To MaxRobs
    If rob(t).exist Then
        totlen = totlen + rob(t).DnaLen
'        On Error GoTo b:     'Botsareus 10/5/2015 Replaced with something better
'        For i = 0 To UBound(rob(t).delgenes) 'Botsareus 9/16/2014 More overflow prevention stuff
'         totlen = totlen + UBound(rob(t).delgenes(i).dna)
'        Next
'b:
    End If
  Next t
  If totlen > 4000000 Then
    Dim calcminenergy As Single
    Dim selectrobot As Integer
    Dim maxdel As Long

    maxdel = 1500 * (CLng(TotalRobotsDisplayed) * 425 / totlen)

    For i = 0 To maxdel
        calcminenergy = 320000 'only erase robots with lowest energy
        For t = 1 To MaxRobs
            If rob(t).exist Then
                If (rob(t).nrg + rob(t).body * 10) < calcminenergy Then
                    calcminenergy = (rob(t).nrg + rob(t).body * 10)
                    selectrobot = t
                End If
            End If
        Next t
        Call KillRobot(selectrobot)
    Next i
  End If
  If totlen > 3000000 Then
        For t = 1 To MaxRobs
            rob(t).LastMutDetail = ""
        Next t
  End If

  'Botsareus 5/6/2013 The safemode system
  If UseSafeMode Then 'special modes does not apply, may need to expended to other restart modes
    If SimOpts.TotRunCycle Mod 2000 = 0 And SimOpts.TotRunCycle > 0 Then  'Botsareus 10/19/2015 Safe mode uses different logic under use internet as randomizer
        SaveSimulation MDIForm1.MainDir + "\saves\lastautosave.sim"
        'Botsareus 5/13/2013 delete local copy
        If dir(MDIForm1.MainDir + "\saves\localcopy.sim") <> "" Then Kill (MDIForm1.MainDir + "\saves\localcopy.sim")
        Open App.path & "\autosaved.gset" For Output As #1
         Write #1, True
        Close #1
        savenow = False
    End If
  End If


Static totnrgnvegs As Double
Dim cmptotnrgnvegs As Double


End Sub
