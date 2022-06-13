Attribute VB_Name = "HDRoutines"
Option Explicit

'
'   D I S K    O P E R A T I O N S
'

Public Function FolderExists(sFullPath As String) As Boolean
  Dim myFSO As New FileSystemObject
  FolderExists = myFSO.FolderExists(sFullPath)
End Function

Public Function RecursiveMkDir(destDir As String) As Boolean
  Dim i As Long, prevDir As String
   
  On Error Resume Next
   
  For i = Len(destDir) To 1 Step -1
    If Mid(destDir, i, 1) = "\" Then
      prevDir = Left(destDir, i - 1)
      Exit For
    End If
  Next
   
  If prevDir = "" Then
    RecursiveMkDir = False
    Exit Function
  End If
  If Not Len(dir(prevDir & "\", vbDirectory)) > 0 And Not RecursiveMkDir(prevDir) Then
    RecursiveMkDir = False
    Exit Function
  End If
   
  On Error GoTo errDirMake
  MkDir destDir
  RecursiveMkDir = True
  Exit Function
   
errDirMake:
  RecursiveMkDir = False
End Function

' inserts organism file in the simulation
' remember that organisms could be made of more than one robot
Public Sub InsertOrganism(path As String)
  Dim x As Single, y As Single, n As Integer
  
  x = Random(60, SimOpts.FieldWidth - 60) 'Botsareus 2/24/2013 bug fix: robots location within screen limits
  y = Random(60, SimOpts.FieldHeight - 60)
  n = LoadOrganism(path, x, y)
End Sub

' saves organism file
Public Sub SaveOrganism(path As String, r As Integer)
  Dim clist(50) As Integer, k As Integer, cnum As Integer
  
  k = 0
  clist(0) = r
  ListCells clist()
  
  While clist(cnum) > 0
    cnum = cnum + 1
  Wend
  
  On Error GoTo problem
  
  Close #401
  
  Open path For Binary As #401
  Put #401, , cnum
  For k = 0 To cnum - 1
    SaveRobotBody 401, clist(k)
  Next
  
  Close #401
  Exit Sub
problem:
  Close #401
End Sub

'Adds a record to the species array when a bot with a new species is loaded or teleported in
Public Function AddSpecie(n As Integer) As Integer
  Dim k As Integer, fso As New FileSystemObject, robotFile As file
  
  If rob(n).Corpse Or rob(n).FName = "Corpse" Or rob(n).exist = False Then
    AddSpecie = 0
    Exit Function
  End If
  
  k = SimOpts.SpeciesNum
  If k < MAXNATIVESPECIES Then SimOpts.SpeciesNum = SimOpts.SpeciesNum + 1
   
  SimOpts.Specie(k).Name = rob(n).FName
  SimOpts.Specie(k).Veg = rob(n).Veg
  SimOpts.Specie(k).CantSee = rob(n).CantSee
  SimOpts.Specie(k).DisableMovementSysvars = rob(n).DisableMovementSysvars
  SimOpts.Specie(k).DisableDNA = rob(n).DisableDNA
  SimOpts.Specie(k).CantReproduce = rob(n).CantReproduce
  SimOpts.Specie(k).VirusImmune = rob(n).VirusImmune
  SimOpts.Specie(k).population = 1
  SimOpts.Specie(k).SubSpeciesCounter = 0
  SimOpts.Specie(k).color = rob(n).color
  SimOpts.Specie(k).Comment = "Species arrived from the Internet"
  SimOpts.Specie(k).Posrg = 1
  SimOpts.Specie(k).Posdn = 1
  SimOpts.Specie(k).Poslf = 0
  SimOpts.Specie(k).Postp = 0
  SetDefaultMutationRates SimOpts.Specie(k).Mutables
  SimOpts.Specie(k).Mutables.Mutations = rob(n).Mutables.Mutations
  SimOpts.Specie(k).qty = 5
  SimOpts.Specie(k).Stnrg = 3000
  SimOpts.Specie(k).path = MDIForm1.MainDir + "\robots"
  AddSpecie = k
End Function

' loads an organism file
Public Function LoadOrganism(path As String, x As Single, y As Single) As Integer
  Dim clist(50) As Integer
  Dim OList(50) As Integer
  Dim k As Integer, cnum As Integer
  Dim i As Integer
  Dim nuovo As Integer
  Dim foundSpecies As Boolean
  
  On Error GoTo problem
  Close #402
  Open path For Binary As #402
  Get #402, , cnum
  For k = 0 To cnum - 1
    nuovo = posto()
    clist(k) = nuovo
    LoadRobot 402, nuovo
    LoadOrganism = nuovo
    i = SimOpts.SpeciesNum
    foundSpecies = False
    While i > 0
      i = i - 1
      If rob(nuovo).FName = SimOpts.Specie(i).Name Then
        foundSpecies = True
        i = 0
      End If
    Wend
    If Not foundSpecies Then AddSpecie nuovo
  Next k
  Close #402
  If x > -1 And y > -1 Then
    PlaceOrganism clist(), x, y
  End If
  RemapTies clist(), OList, cnum

  Exit Function
problem:
  Close #402
  LoadOrganism = -1
  If nuovo > 0 Then
    rob(nuovo).exist = False
    UpdateBotBucket nuovo
  End If
End Function

' places an organism (made of robots listed in clist())
' in the specified x,y position
Public Sub PlaceOrganism(clist() As Integer, x As Single, y As Single)
  Dim k As Integer
  Dim dx As Single, dy As Single
  k = 0
  
  dx = x - rob(clist(0)).pos.x
  dy = y - rob(clist(0)).pos.y
  While clist(k) > 0
    rob(clist(k)).pos.x = rob(clist(k)).pos.x + dx
    rob(clist(k)).pos.y = rob(clist(k)).pos.y + dy
    rob(clist(k)).BucketPos.x = -2
    rob(clist(k)).BucketPos.y = -2
    UpdateBotBucket clist(k)
    k = k + 1
  Wend
End Sub

' remaps ties from the old index numbers - those the robots had
' when saved to disk- to the new indexes assigned in this simulation
Public Sub RemapTies(clist() As Integer, OList() As Integer, cnum As Integer)
  Dim t As Integer, ind As Integer, j As Integer, k As Integer
  Dim TiePointsToNode As Boolean
    
  For t = 0 To cnum - 1  ' Loop through each cell
    ind = rob(clist(t)).oldBotNum
    For k = 0 To cnum - 1  ' Loop through each cell
      j = 1
      While rob(clist(k)).Ties(j).pnt > 0  ' Loop through each tie
        If rob(clist(k)).Ties(j).pnt = ind Then
          rob(clist(k)).Ties(j).pnt = clist(t)
        End If
        j = j + 1
      Wend
    Next k
  Next t
  
  For k = 0 To cnum - 1 ' All cells
    j = 1
    While rob(clist(k)).Ties(j).pnt > 0 'All Ties
      TiePointsToNode = False
      For t = 0 To cnum - 1
        ind = clist(t)
        If rob(clist(k)).Ties(j).pnt = ind Then
          TiePointsToNode = True
        End If
      Next t
      If Not TiePointsToNode Then rob(clist(k)).Ties(j).pnt = 0
      j = j + 1
    Wend
  Next k
End Sub

Public Function RemapAllTies(numOfBots As Integer)
  Dim i As Integer, j As Integer, k As Integer
  
  For i = 1 To numOfBots
    j = 1
    While rob(i).Ties(j).pnt > 0  ' Loop through each tie
      For k = 1 To numOfBots
        If rob(i).Ties(j).pnt = rob(k).oldBotNum Then
          rob(i).Ties(j).pnt = k
          GoTo nexttie
        End If
      Next k
nexttie:
      j = j + 1
    Wend
  Next i
End Function

Public Function RemapAllShots(numOfShots As Long)
  Dim i As Long, j As Integer
  
  For i = 1 To numOfShots
    If Shots(i).Exists Then
      For j = 1 To MaxRobs
        If rob(j).exist And Shots(i).parent = rob(j).oldBotNum Then
          Shots(i).parent = j
          If Shots(i).stored Then rob(j).virusshot = i
          GoTo nextshot
        End If
      Next j
      Shots(i).stored = False ' Could not find parent.  Should probalby never happen but if it does, release the shot
    End If
nextshot:
  Next i
End Function

Public Function GetFilePath(FileName As String) As String
  Dim i As Long
  For i = Len(FileName) To 1 Step -1
    Select Case Mid$(FileName, i, 1)
      Case ":"
        ' colons are always included in the result
        GetFilePath = Left$(FileName, i)
        Exit For
      Case "\"
        ' backslash aren't included in the result
        GetFilePath = Left$(FileName, i - 1)
        Exit For
    End Select
  Next
End Function

' saves a whole simulation
Public Sub SaveSimulation(path As String)
  On Error GoTo tryagain
  Dim t As Integer, n As Integer, x As Integer, j As Long, s2 As String, temp As String, numOfExistingBots As Integer
  
  Form1.MousePointer = vbHourglass
  
  numOfExistingBots = 0
  
  For x = 1 To MaxRobs
    If rob(x).exist Then numOfExistingBots = numOfExistingBots + 1
  Next x
  
  Dim justPath As String
  justPath = GetFilePath(path)
  
  RecursiveMkDir (justPath)
  
  Close #1
  Open path For Binary As #1
  Put #1, , numOfExistingBots
  
  Form1.lblSaving.Visible = True 'Botsareus 1/14/2014 New code to display save status
  
  For t = 1 To MaxRobs
    If rob(t).exist Then
      SaveRobotBody 1, t
    End If
    If t Mod 20 = 0 Then
      Form1.lblSaving.Caption = "Saving... (" & Int(t / MaxRobs * 100) & "%)" 'Botsareus 1/14/2014
      DoEvents
    End If
  Next t
  
  Put #1, , SimOpts.BlockedVegs
  Put #1, , SimOpts.Costs(SHOTCOST)
  Put #1, , SimOpts.CostExecCond
  Put #1, , SimOpts.Costs(COSTSTORE)
  Put #1, , SimOpts.DeadRobotSnp
  Put #1, , SimOpts.SnpExcludeVegs
  Put #1, , SimOpts.DisableTies
  Put #1, , SimOpts.EnergyExType
  Put #1, , SimOpts.EnergyFix
  Put #1, , SimOpts.EnergyProp
  Put #1, , SimOpts.FieldHeight
  Put #1, , SimOpts.FieldSize
  Put #1, , SimOpts.FieldWidth
  Put #1, , SimOpts.KillDistVegs
  Put #1, , SimOpts.MaxEnergy
  Put #1, , SimOpts.MaxPopulation
  Put #1, , SimOpts.MinVegs
  Put #1, , SimOpts.MutCurrMult
  Put #1, , SimOpts.MutCycMax
  Put #1, , SimOpts.MutCycMin
  Put #1, , SimOpts.MutOscill
  Put #1, , SimOpts.PhysBrown
  Put #1, , SimOpts.Ygravity
  Put #1, , SimOpts.Zgravity
  Put #1, , SimOpts.PhysMoving
  Put #1, , SimOpts.PhysSwim
  Put #1, , SimOpts.PopLimMethod
  Put #1, , Len(SimOpts.SimName)
  Put #1, , SimOpts.SimName
  Put #1, , SimOpts.SpeciesNum
  Put #1, , SimOpts.Toroidal
  Put #1, , SimOpts.TotBorn
  Put #1, , SimOpts.TotRunCycle
  Put #1, , SimOpts.TotRunTime
  Put #1, , SimOpts.Pondmode
  Put #1, , SimOpts.LightIntensity
  Put #1, , SimOpts.CorpseEnabled
  Put #1, , SimOpts.Decay
  Put #1, , SimOpts.Gradient
  Put #1, , SimOpts.DayNight
  Put #1, , SimOpts.CycleLength
  Put #1, , SimOpts.Decaydelay
  Put #1, , SimOpts.DecayType
  Put #1, , SimOpts.Costs(MOVECOST)
  Put #1, , SimOpts.Restart
  Put #1, , SimOpts.Dxsxconnected
  Put #1, , SimOpts.Updnconnected
  Put #1, , SimOpts.RepopAmount
  Put #1, , SimOpts.RepopCooldown
  Put #1, , SimOpts.ZeroMomentum
  Put #1, , SimOpts.UserSeedNumber
  Put #1, , SimOpts.SpeciesNum
  
  Dim k As Integer
  For k = 0 To SimOpts.SpeciesNum - 1
    Put #1, , SimOpts.Specie(k).Colind
    Put #1, , SimOpts.Specie(k).color
    Put #1, , SimOpts.Specie(k).Fixed
    Put #1, , SimOpts.Specie(k).Mutables.mutarray
    Put #1, , SimOpts.Specie(k).Mutables.Mutations
    Put #1, , Len(SimOpts.Specie(k).Name)
    Put #1, , SimOpts.Specie(k).Name
    Put #1, , Len(SimOpts.Specie(k).path)
    Put #1, , SimOpts.Specie(k).path
    Put #1, , SimOpts.Specie(k).qty
    Put #1, , SimOpts.Specie(k).Skin
    Put #1, , SimOpts.Specie(k).Stnrg
    Put #1, , SimOpts.Specie(k).Veg
  Next k
  
  Put #1, , SimOpts.VegFeedingToBody
  Put #1, , SimOpts.CoefficientStatic
  Put #1, , SimOpts.CoefficientKinetic
  Put #1, , SimOpts.PlanetEaters
  Put #1, , SimOpts.PlanetEatersG
  Put #1, , SimOpts.Viscosity
  Put #1, , SimOpts.Density
  
  For k = 0 To SimOpts.SpeciesNum - 1
    Put #1, , SimOpts.Specie(k).Mutables.CopyErrorWhatToChange
    Put #1, , SimOpts.Specie(k).Mutables.PointWhatToChange
    
    Dim h As Integer
    
    For h = 0 To 20
      Put #1, , SimOpts.Specie(k).Mutables.Mean(h)
      Put #1, , SimOpts.Specie(k).Mutables.StdDev(h)
    Next h
  Next k
  
  For k = 0 To 70
    Put #1, , SimOpts.Costs(k)
  Next k
  
  'EricL 4/1/2006 Fixed bug below by added -1.  Loop was executing one too many times...
  For k = 0 To SimOpts.SpeciesNum - 1
    Put #1, , SimOpts.Specie(k).Poslf
    Put #1, , SimOpts.Specie(k).Posrg
    Put #1, , SimOpts.Specie(k).Postp
    Put #1, , SimOpts.Specie(k).Posdn
  Next k
  
  Put #1, , SimOpts.BadWastelevel    'EricL 4/1/2006 Added this
  Put #1, , SimOpts.chartingInterval 'EricL 4/1/2006 Added this
  Put #1, , SimOpts.CoefficientElasticity 'EricL 4/29/2006 Added this
  Put #1, , SimOpts.FluidSolidCustom ' EricL 5/7/2006
  Put #1, , SimOpts.CostRadioSetting ' EricL 5/7/2006
  Put #1, , SimOpts.MaxVelocity ' EricL 5/15/2006
  Put #1, , SimOpts.NoShotDecay ' EricL 6/8/2006
  Put #1, , SimOpts.SunUpThreshold 'EricL 6/8/2006 Added this
  Put #1, , SimOpts.SunUp 'EricL 6/8/2006 Added this
  Put #1, , SimOpts.SunDownThreshold 'EricL 6/8/2006 Added this
  Put #1, , SimOpts.SunDown 'EricL 6/8/2006 Added this
  Put #1, , SimOpts.FixedBotRadii
  Put #1, , SimOpts.DayNightCycleCounter
  Put #1, , SimOpts.Daytime
  Put #1, , SimOpts.SunThresholdMode
  Put #1, , numObstacles
  
  For x = 1 To numObstacles
    SaveObstacle 1, x
  Next x
  
  For k = 0 To SimOpts.SpeciesNum - 1
    Put #1, , SimOpts.Specie(k).CantSee
    Put #1, , SimOpts.Specie(k).DisableDNA
    Put #1, , SimOpts.Specie(k).DisableMovementSysvars
  Next k
  
  Put #1, , SimOpts.shapesAreVisable
  Put #1, , SimOpts.allowVerticalShapeDrift
  Put #1, , SimOpts.allowHorizontalShapeDrift
  Put #1, , SimOpts.shapesAreSeeThrough
  Put #1, , SimOpts.shapesAbsorbShots
  Put #1, , SimOpts.shapeDriftRate
  Put #1, , SimOpts.makeAllShapesTransparent
  Put #1, , SimOpts.makeAllShapesBlack
  
  For k = 0 To SimOpts.SpeciesNum - 1
    Put #1, , SimOpts.Specie(k).CantReproduce
  Next k
  
  Put #1, , maxshotarray
  
  For j = 1 To maxshotarray
    SaveShot 1, j
  Next j
  
  Put #1, , SimOpts.MaxAbsNum
    
  For k = 0 To SimOpts.SpeciesNum - 1
     Put #1, , SimOpts.Specie(k).VirusImmune
   Next k
   
  For k = 0 To SimOpts.SpeciesNum - 1
     Put #1, , SimOpts.Specie(k).population
     Put #1, , SimOpts.Specie(k).SubSpeciesCounter
  Next k
  
  Put #1, , SimOpts.oldCostX
  Put #1, , SimOpts.DisableMutations
  Put #1, , SimOpts.SimGUID
  Put #1, , SimOpts.SpeciationGenerationalDistance
  Put #1, , SimOpts.SpeciationGeneticDistance
  Put #1, , SimOpts.EnableAutoSpeciation
  Put #1, , SimOpts.SpeciationMinimumPopulation
  Put #1, , SimOpts.SpeciationForkInterval
  Put #1, , SimOpts.DisableTypArepro
  Put #1, , Len(strGraphQuery1)
  Put #1, , strGraphQuery1
  Put #1, , Len(strGraphQuery2)
  Put #1, , strGraphQuery2
  Put #1, , Len(strGraphQuery3)
  Put #1, , strGraphQuery3
  Put #1, , Len(strSimStart)
  Put #1, , strSimStart

  'the graphs themselfs
  For k = 1 To NUMGRAPHS
    Put #1, , graphfilecounter(k)
    Put #1, , graphvisible(k)
    Put #1, , graphleft(k)
    Put #1, , graphtop(k)
    Put #1, , graphsave(k)
  Next k
  
  Put #1, , SimOpts.NoWShotDecay 'Botsareus 9/28/2013
  
  'Botsareus 8/5/2014
  Put #1, , SimOpts.DisableFixing
  
  'Botsareus 8/16/2014
  Put #1, , SunPosition
  Put #1, , SunRange
  Put #1, , SunChange
  
  'Botsareus 10/13/2014
  Put #1, , SimOpts.Tides
  Put #1, , SimOpts.TidesOf
  
  'Botsareus 10/8/2015
  Put #1, , SimOpts.MutOscillSine
  
  'Botsareus 10/20/2015
  Put #1, , stagnent
       
  Form1.lblSaving.Visible = False 'Botsareus 1/14/2014
    
  Close #1
  Form1.MousePointer = vbArrow
  Exit Sub
tryagain:
  SaveSimulation path
End Sub

'Botsareus 3/15/2013 load global settings
Public Sub LoadGlobalSettings()
  Dim holdmaindir As String
  
  'defaults
  bodyfix = 32100
  chseedstartnew = True
  chseedloadsim = True
  GraphUp = False
  HideDB = False
  MDIForm1.MainDir = App.path
  UseSafeMode = True 'Botsareus 10/5/2015
  UseEpiGene = False 'Botsareus 10/8/2015
  intFindBestV2 = 100
  UseOldColor = True
  'mutations tab
  epiresetemp = 1.3
  epiresetOP = 17
  'Delta2
  Delta2 = False
  DeltaMainExp = 1
  DeltaMainLn = 0
  DeltaDevExp = 7
  DeltaDevLn = 1
  DeltaPM = 3000
  DeltaWTC = 15
  DeltaMainChance = 100
  DeltaDevChance = 30
  'Normailize mutation rates
  NormMut = False
  valNormMut = 1071
  valMaxNormMut = 1071

  'see if maindir overwrite exisits
  If dir(App.path & "\Maindir.gset") <> "" Then
    'load the new maindir
    Open App.path & "\Maindir.gset" For Input As #1
      Input #1, holdmaindir
    Close #1
    If dir(holdmaindir & "\", vbDirectory) <> "" Then
        MDIForm1.MainDir = holdmaindir
    End If
  End If

  'see if settings exsist
  If dir(MDIForm1.MainDir & "\Global.gset") <> "" Then
    'load all settings
    Open MDIForm1.MainDir & "\Global.gset" For Input As #1
    Input #1, screenratiofix
    If Not EOF(1) Then Input #1, bodyfix
    If Not EOF(1) Then Input #1, reprofix
    If Not EOF(1) Then Input #1, chseedstartnew
    If Not EOF(1) Then Input #1, chseedloadsim
    If Not EOF(1) Then Input #1, UseSafeMode
    If Not EOF(1) Then Input #1, intFindBestV2
    If Not EOF(1) Then Input #1, UseOldColor
    If Not EOF(1) Then Input #1, boylabldisp
    If Not EOF(1) Then Input #1, startnovid
    If Not EOF(1) Then Input #1, epireset
    If Not EOF(1) Then Input #1, epiresetemp
    If Not EOF(1) Then Input #1, epiresetOP
    If Not EOF(1) Then Input #1, sunbelt
    '
    If Not EOF(1) Then Input #1, Delta2
    If Not EOF(1) Then Input #1, DeltaMainExp
    If Not EOF(1) Then Input #1, DeltaMainLn
    If Not EOF(1) Then Input #1, DeltaDevExp
    If Not EOF(1) Then Input #1, DeltaDevLn
    If Not EOF(1) Then Input #1, DeltaPM
    '
    If Not EOF(1) Then Input #1, NormMut
    If Not EOF(1) Then Input #1, valNormMut
    If Not EOF(1) Then Input #1, valMaxNormMut
    If Not EOF(1) Then Input #1, DeltaWTC
    If Not EOF(1) Then Input #1, DeltaMainChance
    If Not EOF(1) Then Input #1, DeltaDevChance
    If Not EOF(1) Then Input #1, StartChlr
    If Not EOF(1) Then Input #1, GraphUp
    If Not EOF(1) Then Input #1, HideDB
    If Not EOF(1) Then Input #1, UseEpiGene
    Close #1
  End If

  'some global settings change during simulation (copy is here)
  loadboylabldisp = boylabldisp
  loadstartnovid = startnovid

  'see if safemode settings exisit
  If dir(App.path & "\Safemode.gset") <> "" Then
    'load all settings
    Open App.path & "\Safemode.gset" For Input As #1
    Input #1, simalreadyrunning
    Close #1
  End If

  'see if autosaved file exisit
  If dir(App.path & "\autosaved.gset") <> "" Then
    'load all settings
    Open App.path & "\autosaved.gset" For Input As #1
    Input #1, autosaved
    Close #1
  End If

  'If we are not using safe mode assume simulation is not runnin'
  If UseSafeMode = False Then simalreadyrunning = False
  If simalreadyrunning = False Then autosaved = False
End Sub


' loads a whole simulation
Public Sub LoadSimulation(path As String)
  Dim j As Long
  Dim k As Long
  Dim x As Integer
  Dim tempbool As Boolean
  Dim tempint As Integer
  Dim temp As String
  Dim s2 As String
  
  Form1.camfix = False 'Botsareus 2/23/2013 When simulation starts the screen is normailized
  Form1.MousePointer = vbHourglass
  
  Open path For Binary As 1
  Get #1, , MaxRobs
  
  'Round up to the next multiple of 500
  ReDim rob(MaxRobs + (500 - (MaxRobs Mod 500)))
  
  Form1.lblSaving.Visible = True 'Botsareus 1/14/2014 New code to display load status
  Form1.Visible = True
  
  For k = 1 To MaxRobs
   LoadRobot 1, k
    If k Mod 20 = 0 Then
      Form1.lblSaving.Caption = "Loading... (" & Int(k / MaxRobs * 100) & "%)" 'Botsareus 1/14/2014
      DoEvents
    End If
  Next k
  
  ' As of 2.42.8, the sim file is packed.  Every bot stored is guarenteed to exist, yet their bot numbers, when loaded, may be
  ' different from the sim they came from.  Thus, we remap all the ties from all the loaded bots.
  RemapAllTies MaxRobs

  Get #1, , SimOpts.BlockedVegs
  Get #1, , SimOpts.Costs(SHOTCOST)
  Get #1, , SimOpts.CostExecCond
  Get #1, , SimOpts.Costs(COSTSTORE)
  Get #1, , SimOpts.DeadRobotSnp
  Get #1, , SimOpts.SnpExcludeVegs
  Get #1, , SimOpts.DisableTies
  Get #1, , SimOpts.EnergyExType
  Get #1, , SimOpts.EnergyFix
  Get #1, , SimOpts.EnergyProp
  Get #1, , SimOpts.FieldHeight
  Get #1, , SimOpts.FieldSize
  Get #1, , SimOpts.FieldWidth
  Get #1, , SimOpts.KillDistVegs
  Get #1, , SimOpts.MaxEnergy
  Get #1, , SimOpts.MaxPopulation
  Get #1, , SimOpts.MinVegs
  Get #1, , SimOpts.MutCurrMult
  Get #1, , SimOpts.MutCycMax
  Get #1, , SimOpts.MutCycMin
  Get #1, , SimOpts.MutOscill
  Get #1, , SimOpts.PhysBrown
  Get #1, , SimOpts.Ygravity
  Get #1, , SimOpts.Zgravity
  Get #1, , SimOpts.PhysMoving
  Get #1, , SimOpts.PhysSwim
  Get #1, , SimOpts.PopLimMethod
  Get #1, , k: SimOpts.SimName = Space(Abs(k))
  Get #1, , SimOpts.SimName
  Get #1, , SimOpts.SpeciesNum
  Get #1, , SimOpts.Toroidal
  Get #1, , SimOpts.TotBorn
  Get #1, , SimOpts.TotRunCycle
  Get #1, , SimOpts.TotRunTime
  Get #1, , SimOpts.Pondmode
  Get #1, , SimOpts.LightIntensity
  Get #1, , SimOpts.CorpseEnabled
  Get #1, , SimOpts.Decay
  Get #1, , SimOpts.Gradient
  Get #1, , SimOpts.DayNight
  Get #1, , SimOpts.CycleLength
  Get #1, , SimOpts.Decaydelay
  Get #1, , SimOpts.DecayType
  Get #1, , SimOpts.Costs(MOVECOST)
  Get #1, , SimOpts.Restart
  Get #1, , SimOpts.Dxsxconnected
  Get #1, , SimOpts.Updnconnected
  Get #1, , SimOpts.RepopAmount
  Get #1, , SimOpts.RepopCooldown
  Get #1, , SimOpts.ZeroMomentum
  Get #1, , SimOpts.UserSeedNumber
  Get #1, , SimOpts.SpeciesNum
  
  For k = 0 To SimOpts.SpeciesNum - 1
    Get #1, , SimOpts.Specie(k).Colind
    Get #1, , SimOpts.Specie(k).color
    Get #1, , SimOpts.Specie(k).Fixed
    Get #1, , SimOpts.Specie(k).Mutables.mutarray
    Get #1, , SimOpts.Specie(k).Mutables.Mutations
    Get #1, , j: SimOpts.Specie(k).Name = Space(Abs(j))
    Get #1, , SimOpts.Specie(k).Name
    Get #1, , j: SimOpts.Specie(k).path = Space(j)
    Get #1, , SimOpts.Specie(k).path
    Get #1, , SimOpts.Specie(k).qty
    Get #1, , SimOpts.Specie(k).Skin
    Get #1, , SimOpts.Specie(k).Stnrg
    Get #1, , SimOpts.Specie(k).Veg
  Next k
  
  Get #1, , SimOpts.VegFeedingToBody
  Get #1, , SimOpts.CoefficientStatic
  Get #1, , SimOpts.CoefficientKinetic
  Get #1, , SimOpts.PlanetEaters
  Get #1, , SimOpts.PlanetEatersG
  Get #1, , SimOpts.Viscosity
  Get #1, , SimOpts.Density
  
  For k = 0 To SimOpts.SpeciesNum - 1
    Get #1, , SimOpts.Specie(k).Mutables.CopyErrorWhatToChange
    Get #1, , SimOpts.Specie(k).Mutables.PointWhatToChange
    
    For j = 0 To 20
      Get #1, , SimOpts.Specie(k).Mutables.Mean(j)
      Get #1, , SimOpts.Specie(k).Mutables.StdDev(j)
    Next j
  Next k
  
  For k = 0 To 70
    Get #1, , SimOpts.Costs(k)
  Next k
  
  For k = 0 To SimOpts.SpeciesNum - 1
    Get #1, , SimOpts.Specie(k).Poslf
    Get #1, , SimOpts.Specie(k).Posrg
    Get #1, , SimOpts.Specie(k).Postp
    Get #1, , SimOpts.Specie(k).Posdn
  Next k
  
  Get #1, , SimOpts.BadWastelevel
  Get #1, , SimOpts.chartingInterval 'EricL 4/1/2006 Added this
  Get #1, , SimOpts.CoefficientElasticity 'EricL 4/29/2006 Added this
  Get #1, , SimOpts.FluidSolidCustom 'EricL 5/7/2006 Added this for UI initialization
  Get #1, , SimOpts.CostRadioSetting 'EricL 5/7/2006 Added this for UI initialization
  Get #1, , SimOpts.MaxVelocity 'EricL 5/16/2006 Added this - was not saved before
  Get #1, , SimOpts.NoShotDecay 'EricL 6/8/2006 Added this
  Get #1, , SimOpts.SunUpThreshold 'EricL 6/8/2006 Added this
  Get #1, , SimOpts.SunUp 'EricL 6/8/2006 Added this
  Get #1, , SimOpts.SunDownThreshold 'EricL 6/8/2006 Added this
  Get #1, , SimOpts.SunDown 'EricL 6/8/2006 Added this
  Get #1, , SimOpts.FixedBotRadii
  Get #1, , SimOpts.DayNightCycleCounter
  Get #1, , SimOpts.Daytime
  Get #1, , SimOpts.SunThresholdMode
  Get #1, , numObstacles
         
  For x = 1 To numObstacles
    LoadObstacle 1, x
  Next x
  
  For k = 0 To SimOpts.SpeciesNum - 1
    Get #1, , SimOpts.Specie(k).CantSee
    Get #1, , SimOpts.Specie(k).DisableDNA
    Get #1, , SimOpts.Specie(k).DisableMovementSysvars
  Next k
  
  Get #1, , SimOpts.shapesAreVisable
  Get #1, , SimOpts.allowVerticalShapeDrift
  Get #1, , SimOpts.allowHorizontalShapeDrift
  Get #1, , SimOpts.shapesAreSeeThrough
  Get #1, , SimOpts.shapesAbsorbShots
  Get #1, , SimOpts.shapeDriftRate
  Get #1, , SimOpts.makeAllShapesTransparent
  Get #1, , SimOpts.makeAllShapesBlack
   
  For k = 0 To SimOpts.SpeciesNum - 1
    Get #1, , SimOpts.Specie(k).CantReproduce
  Next k
  
  Get #1, , maxshotarray
      
  ReDim Shots(maxshotarray)
             
  For j = 1 To maxshotarray
    LoadShot 1, j
  Next j
  RemapAllShots maxshotarray
  
  SimOpts.MaxAbsNum = MaxRobs
  Get #1, , SimOpts.MaxAbsNum
  
  For k = 0 To SimOpts.SpeciesNum - 1
    Get #1, , SimOpts.Specie(k).VirusImmune
  Next k
  
  For k = 0 To SimOpts.SpeciesNum - 1
    Get #1, , SimOpts.Specie(k).population
    Get #1, , SimOpts.Specie(k).SubSpeciesCounter
  Next k
      
  Get #1, , SimOpts.oldCostX
  Get #1, , SimOpts.DisableMutations
  Get #1, , SimOpts.SimGUID
  Get #1, , SimOpts.SpeciationGenerationalDistance
  Get #1, , SimOpts.SpeciationGeneticDistance
  Get #1, , SimOpts.EnableAutoSpeciation
  Get #1, , SimOpts.SpeciationMinimumPopulation
  Get #1, , SimOpts.SpeciationForkInterval
  Get #1, , SimOpts.DisableTypArepro
  Get #1, , j: strGraphQuery1 = Space(j)
  Get #1, , strGraphQuery1
  Get #1, , j: strGraphQuery2 = Space(j)
  Get #1, , strGraphQuery2
  Get #1, , j: strGraphQuery3 = Space(j)
  Get #1, , strGraphQuery3
  Get #1, , j: strSimStart = Space(j)
  Get #1, , strSimStart
  'the graphs themselfs
  For k = 1 To NUMGRAPHS
   Get #1, , graphfilecounter(k)
   Get #1, , graphvisible(k)
   Get #1, , graphleft(k)
   Get #1, , graphtop(k)
   Get #1, , graphsave(k)
   
   If graphvisible(k) Then
     Select Case k
      Case 1
        Form1.NewGraph POPULATION_GRAPH, "Populations"
      Case 2
        Form1.NewGraph MUTATIONS_GRAPH, "Average_Mutations"
      Case 3
        Form1.NewGraph AVGAGE_GRAPH, "Average_Age"
      Case 4
        Form1.NewGraph OFFSPRING_GRAPH, "Average_Offspring"
      Case 5
        Form1.NewGraph ENERGY_GRAPH, "Average_Energy"
      Case 6
        Form1.NewGraph DNALENGTH_GRAPH, "Average_DNA_length"
      Case 7
        Form1.NewGraph DNACOND_GRAPH, "Average_DNA_Cond_statements"
      Case 8
        Form1.NewGraph MUT_DNALENGTH_GRAPH, "Average_Mutations_per_DNA_length_x1000-"
      Case 9
        Form1.NewGraph ENERGY_SPECIES_GRAPH, "Total_Energy_per_Species_x1000-"
      Case 10
        Form1.NewGraph DYNAMICCOSTS_GRAPH, "Dynamic_Costs"
      Case 11
        Form1.NewGraph SPECIESDIVERSITY_GRAPH, "Species_Diversity"
      Case 12
        Form1.NewGraph AVGCHLR_GRAPH, "Average_Chloroplasts"
      Case 13
        Form1.NewGraph GENETIC_DIST_GRAPH, "Genetic_Distance_x1000-"
      Case 14
        Form1.NewGraph GENERATION_DIST_GRAPH, "Max_Generational_Distance"
      Case 15
        Form1.NewGraph GENETIC_SIMPLE_GRAPH, "Simple_Genetic_Distance_x1000-"
      Case 16
          Form1.NewGraph CUSTOM_1_GRAPH, "Customizable_Graph_1-"
      Case 17
          Form1.NewGraph CUSTOM_2_GRAPH, "Customizable_Graph_2-"
      Case 18
          Form1.NewGraph CUSTOM_3_GRAPH, "Customizable_Graph_3-"
    End Select
   End If
  Next k
  
  Get #1, , SimOpts.NoWShotDecay 'EricL 6/8/2006 Added this

  Get #1, , SimOpts.DisableFixing
   
  Get #1, , SunPosition
  Get #1, , SunRange
  Get #1, , SunChange
  
  'Botsareus 10/13/2014
  Get #1, , SimOpts.Tides
  Get #1, , SimOpts.TidesOf
  
  'Botsareus 10/8/2015
  Get #1, , SimOpts.MutOscillSine
  
  'Botsareus 10/20/2015
  Get #1, , stagnent
   
  Form1.lblSaving.Visible = False 'Botsareus 1/14/2014
    
  Close 1
  
  If SimOpts.Costs(DYNAMICCOSTSENSITIVITY) = 0 Then SimOpts.Costs(DYNAMICCOSTSENSITIVITY) = 500
   
  'EricL 3/28/2006 This line insures that all the simulation dialog options get set to match the loaded sim
  TmpOpts = SimOpts
  
  Form1.MousePointer = vbArrow
End Sub

' loads a single robot
Public Sub LoadRobot(fnum As Integer, ByVal n As Integer)
  LoadRobotBody fnum, n
  If rob(n).exist Then
    GiveAbsNum n
    insertsysvars n
    ScanUsedVars n
    makeoccurrlist n
    rob(n).DnaLen = DnaLen(rob(n).dna())
    rob(n).genenum = CountGenes(rob(n).dna())
    rob(n).mem(DnaLenSys) = rob(n).DnaLen
    rob(n).mem(GenesSys) = rob(n).genenum
  End If
End Sub

' assignes a robot his unique code
Public Sub GiveAbsNum(k As Integer)
  If rob(k).AbsNum = 0 Then
    SimOpts.MaxAbsNum = SimOpts.MaxAbsNum + 1
    rob(k).AbsNum = SimOpts.MaxAbsNum
  End If
End Sub

' loads the body of the robot
Private Sub LoadRobotBody(n As Integer, r As Integer)
  Dim t As Integer, k As Integer, Fe As Byte, L1 As Long, inttmp As Integer
  Dim MessedUpMutations As Boolean
  Dim longtmp As Long
  
  MessedUpMutations = False
  With rob(r)
    Get #n, , .Veg
    Get #n, , .wall
    Get #n, , .Fixed
    
    Get #n, , .pos.x
    Get #n, , .pos.y
    Get #n, , .vel.x
    Get #n, , .vel.y
    Get #n, , .aim
    Get #n, , .ma           'momento angolare
    Get #n, , .mt           'momento torcente
    
    .BucketPos.x = -2
    .BucketPos.y = -2
     
    'ties
    For t = 0 To MAXTIES
      Get #n, , .Ties(t).Port
      Get #n, , .Ties(t).pnt
      Get #n, , .Ties(t).ptt
      Get #n, , .Ties(t).ang
      Get #n, , .Ties(t).bend
      Get #n, , .Ties(t).angreg
      Get #n, , .Ties(t).ln
      Get #n, , .Ties(t).shrink
      Get #n, , .Ties(t).stat
      Get #n, , .Ties(t).last
      Get #n, , .Ties(t).mem
      Get #n, , .Ties(t).back
      Get #n, , .Ties(t).nrgused
      Get #n, , .Ties(t).infused
      Get #n, , .Ties(t).sharing
    Next t
    
    Get #n, , .nrg
    
    For t = 1 To 50
      Get #n, , k: .vars(t).Name = Space(k)
      Get #n, , .vars(t).Name   '|
      Get #n, , .vars(t).value
    Next t
    Get #n, , .vnum             '| variabili private
    
    ' macchina virtuale
    Get #n, , .mem()       ' memoria dati
    Get #n, , k: ReDim .dna(k)
    
    For t = 1 To k
      Get #n, , .dna(t).tipo
      Get #n, , .dna(t).value
    Next t
    
    'Force an end base pair to protect against DNA corruption
    .dna(k).tipo = 10
    .dna(k).value = 1
        
        
    'EricL Set reasonable default values to protect against corrupted sims that don't read these values
    SetDefaultMutationRates .Mutables, True
    
    For t = 0 To 20
      Get #n, , .Mutables.mutarray(t)
    Next t
    
    ' informative
    Get #n, , .SonNumber
    Get #n, , inttmp
    .Mutations = inttmp
    Get #n, , inttmp
    .LastMut = inttmp
    Get #n, , .parent
    Get #n, , .age
    Get #n, , .BirthCycle
    Get #n, , .genenum
    Get #n, , .generation
    Get #n, , .DnaLen
    
    ' aspetto
    Get #n, , .Skin()
    Get #n, , .color
    
    Get #n, , .body: .radius = FindRadius(r)
    Get #n, , .Bouyancy
    Get #n, , .Corpse
    Get #n, , .Pwaste
    Get #n, , .Waste
    Get #n, , .poison
    Get #n, , .venom
    Get #n, , .exist
    Get #n, , .Dead
    Get #n, , k: .FName = Space(k)
    Get #n, , .FName
            
    Get #n, , k: .LastOwner = Space(k)
    Get #n, , .LastOwner
    If .LastOwner = "" Then .LastOwner = "Local"
    
    If FileContinue(n) Then Get #n, , k
       
    'EricL 5/2/2006  This needs some explaining.  The length of the mutation details can exceed 2^15 -1 for bots with lots
    'of mutations.  If we are reading an old file, the length could be negative in which case we read what we can and then punt and skip the
    'rest of the bot.  We will miss some stuff, like the mutation settings, but at least the sim will load.
    'If the sim file was stored with 2.42.4 or later and this bot has a ton of mutation details, then an Int value of 1
    'indicates the actual length of the mutation details is stored as a Long in which case we read that and continue.
    If k < 0 Then
      'Its an old corrupted file with > 2^15 worth of mutation details.  Bail.
      .LastMutDetail = "Problem reading mutation details.  May be a very old sim.  Please tell the developers.  Mutation Details deleted."
      
      'EricL Set reasonable default values for everything read from this point on.
      .Mutables.Mutations = True

      SetDefaultMutationRates .Mutables, True
    
      .View = True
      .NewMove = False
      .oldBotNum = 0
      .CantSee = False
      .DisableDNA = False
      .DisableMovementSysvars = False
      .CantReproduce = False
      .VirusImmune = False
      .shell = 0
      .Slime = 0
                     
      GoTo OldFile
    End If
    If k = 1 Then
      'Its a new file with lots of mutations.  Read the actual length stored as a Long
      Get #n, , L1
    Else
      'Not that many mutations for this bot (It's possible its an old file with lots of mutations and the len wrapped.
      'If so, we just read the postiive len and keep going.  Everything following this will be wrong, but the sim should
      'still load.  It's a corner case.  The alternative is to try to parse the mutation details strings directly.  No thanks.
      L1 = CLng(k)
    End If
    
    If Form1.lblSaving.Visible Then 'Botsareus 4/18/2016 Bug fix to prevent string buffer overflow
        .LastMutDetail = Space(L1)
        If FileContinue(n) Then Get #n, , .LastMutDetail
    Else
        If L1 > (100000000 / TotalRobotsDisplayed) Then
            Seek #n, L1 + Seek(n)
        Else
            .LastMutDetail = Space(L1)
            If FileContinue(n) Then Get #n, , .LastMutDetail
        End If
    End If
    
    If FileContinue(n) Then Get #n, , .Mutables.Mutations
    
    For t = 0 To 20
      If FileContinue(n) Then Get #n, , .Mutables.Mean(t)
      If FileContinue(n) Then Get #n, , .Mutables.StdDev(t)
    Next t
    
    For t = 0 To 20
      If .Mutables.Mean(t) < 0 Or .Mutables.Mean(t) > 32000 Or .Mutables.StdDev(t) < 0 Or .Mutables.StdDev(t) > 32000 Then MessedUpMutations = True
    Next t
        
    If FileContinue(n) Then Get #n, , .Mutables.CopyErrorWhatToChange
    If FileContinue(n) Then Get #n, , .Mutables.PointWhatToChange
    
    If .Mutables.CopyErrorWhatToChange < 0 Or .Mutables.CopyErrorWhatToChange > 32000 Or .Mutables.PointWhatToChange < 0 Or .Mutables.PointWhatToChange > 32000 Then
      MessedUpMutations = True
    End If
    
     'If we read wacky values, the file was saved with an older version which messed these up.  Set the defaults.
    If MessedUpMutations Then
      SetDefaultMutationRates .Mutables, True
    End If
    
    If FileContinue(n) Then Get #n, , .View
    If FileContinue(n) Then Get #n, , .NewMove
    
    .oldBotNum = 0
    If FileContinue(n) Then Get #n, , .oldBotNum
      
    .CantSee = False
    If FileContinue(n) Then Get #n, , .CantSee
    If CInt(.CantSee) > 0 Or CInt(.CantSee) < -1 Then .CantSee = False ' Protection against corrpt sim files.
    
    .DisableDNA = False
    If FileContinue(n) Then Get #n, , .DisableDNA
    If CInt(.DisableDNA) > 0 Or CInt(.DisableDNA) < -1 Then .DisableDNA = False ' Protection against corrpt sim files.
    
    .DisableMovementSysvars = False
    If FileContinue(n) Then Get #n, , .DisableMovementSysvars
    If CInt(.DisableMovementSysvars) > 0 Or CInt(.DisableMovementSysvars) < -1 Then .DisableMovementSysvars = False    ' Protection against corrpt sim files.
    
    .CantReproduce = False
    If FileContinue(n) Then Get #n, , .CantReproduce
    If CInt(.CantReproduce) > 0 Or CInt(.CantReproduce) < -1 Then .CantReproduce = False ' Protection against corrpt sim files.
    
    .shell = 0
    If FileContinue(n) Then Get #n, , .shell
    
    If .shell > 32000 Then .shell = 32000
    If .shell < 0 Then .shell = 0
    
    .Slime = 0
    If FileContinue(n) Then Get #n, , .Slime
            
    If .Slime > 32000 Then .Slime = 32000
    If .Slime < 0 Then .Slime = 0
    
    .VirusImmune = False
    If FileContinue(n) Then Get #n, , .VirusImmune
    If CInt(.VirusImmune) > 0 Or CInt(.VirusImmune) < -1 Then .VirusImmune = False ' Protection against corrpt sim files.
    
    .SubSpecies = 0 ' For older sims saved before this was implemented, set the sup species to be the bot's number.  Every bot is a sub species.
    If FileContinue(n) Then Get #n, , .SubSpecies
    
    .spermDNAlen = 0
    If FileContinue(n) Then Get #n, , .spermDNAlen: ReDim .spermDNA(.spermDNAlen)
    For t = 1 To .spermDNAlen
      If FileContinue(n) Then Get #n, , .spermDNA(t).tipo
      If FileContinue(n) Then Get #n, , .spermDNA(t).value
    Next t
    
    .fertilized = -1
    If FileContinue(n) Then Get #n, , .fertilized
        
    .sim = 0
    If FileContinue(n) Then Get #n, , .sim
    If FileContinue(n) Then Get #n, , .AbsNum
    
    'Botsareus 2/23/2013 Rest of tie data
    If FileContinue(n) Then Get #n, , .Multibot
    For t = 0 To MAXTIES
        If FileContinue(n) Then Get #n, , .Ties(t).type
        If FileContinue(n) Then Get #n, , .Ties(t).b
        If FileContinue(n) Then Get #n, , .Ties(t).k
        If FileContinue(n) Then Get #n, , .Ties(t).NaturalLength
        'Botsareus 4/18/2016 Protection against currupt file
        If .Ties(t).NaturalLength < 0 Then .Ties(t).NaturalLength = 0
        If .Ties(t).NaturalLength > 1500 Then .Ties(t).NaturalLength = 1500
    Next
    
    'Botsareus 4/9/2013 For genetic distance graph
    If FileContinue(n) Then Get #n, , .OldGD
    .GenMut = .DnaLen / GeneticSensitivity
    
    'Panda 2013/08/11 chloroplasts
    If FileContinue(n) Then Get #n, , .chloroplasts
    'Botsareus 4/18/2016 Protection against currupt file
    If .chloroplasts < 0 Then .chloroplasts = 0
    If .chloroplasts > 32000 Then .chloroplasts = 32000
    
    'Botsareus 12/3/2013 Read epigenetic information
    
    For t = 0 To 14
        If FileContinue(n) Then Get #n, , .epimem(t)
    Next
    
    'Botsareus 1/28/2014 Read robot tag
    
    If FileContinue(n) Then Get #n, , .tag
    
    'Read if robot is using sunbelt
        
    Dim usesunbelt As Boolean 'sunbelt mutations
    
    If FileContinue(n) Then Get #n, , usesunbelt
    
    'Botsareus 3/28/2014 Read if disable chloroplasts
    
    If FileContinue(n) Then Get #n, , .NoChlr
    
    'Botsareus 3/28/2014 Read kill resrictions
    
    If FileContinue(n) Then Get #n, , .Chlr_Share_Delay
    If .Chlr_Share_Delay > 8 Then .Chlr_Share_Delay = 8 'Botsareus 4/18/2016 Protection against currupt file
    
    'Botsareus 10/8/2015 Keep track of mutations from old dna file
    If FileContinue(n) Then Get #n, , .OldMutations
    
    'Botsareus 6/22/2016 Actual velocity
    
    If FileContinue(n) Then Get #n, , .actvel.x
    If FileContinue(n) Then Get #n, , .actvel.y
    
    
    If .Veg Then
     If TotalChlr > SimOpts.MaxPopulation Then .Dead = True
    End If
    If .FName = "Corpse" Then .nrg = 0
    
    'read in any future data here
    
OldFile:
    'burn through any new data from a different version
    While FileContinue(n)
      Get #n, , Fe
    Wend
    
    'grab these three FE codes
    Get #n, , Fe
    Get #n, , Fe
    Get #n, , Fe
    
    'don't you dare put anything after this!
    'except some initialization stuff
    .Vtimer = 0
    .virusshot = 0
    
    'Botsareus 2/21/2014 Special case reset sunbelt mutations
    
    If Not usesunbelt Then
        .Mutables.mutarray(P2UP) = 0
        .Mutables.mutarray(CE2UP) = 0
        .Mutables.mutarray(AmplificationUP) = 0
        .Mutables.mutarray(TranslocationUP) = 0
    End If
  End With
End Sub

Private Function FileContinue(filenumber As Integer) As Boolean
  'three FE bytes (ie: 254) means we are at the end of the record
  
  Dim Fe As Byte
  Dim Position As Long
  Dim k As Integer
  
  FileContinue = False
  Position = Seek(filenumber)
   
  Do
    If Not EOF(filenumber) Then
      Get #filenumber, , Fe
    Else
      FileContinue = False
      Fe = 254
    End If
    
    k = k + 1
    
    If Fe <> 254 Then
      FileContinue = True
      'exit immediatly, we are done
    End If
  Loop While Not FileContinue And k < 3
  
  'reset position
  Get #filenumber, Position - 1, Fe
End Function
' saves the body of the robot
Private Sub SaveRobotBody(n As Integer, r As Integer)
  Dim t As Integer, k As Integer
  Dim s As String
  Dim s2 As String
  Dim temp As String
  Dim longtmp As Long
  
  Const Fe As Byte = 254
 ' Dim space As Integer
  
  s = "Mutation Details removed in last save."
  
  With rob(r)
    
    Put #n, , .Veg
    Put #n, , .wall
    Put #n, , .Fixed
    
    ' fisiche
    Put #n, , .pos.x
    Put #n, , .pos.y
    Put #n, , .vel.x
    Put #n, , .vel.y
    Put #n, , .aim
    Put #n, , .ma           'momento angolare
    Put #n, , .mt           'momento torcente
    
    For t = 0 To MAXTIES
      Put #n, , .Ties(t).Port
      Put #n, , .Ties(t).pnt
      Put #n, , .Ties(t).ptt
      Put #n, , .Ties(t).ang
      Put #n, , .Ties(t).bend
      Put #n, , .Ties(t).angreg
      Put #n, , .Ties(t).ln
      Put #n, , .Ties(t).shrink
      Put #n, , .Ties(t).stat
      Put #n, , .Ties(t).last
      Put #n, , .Ties(t).mem
      Put #n, , .Ties(t).back
      Put #n, , .Ties(t).nrgused
      Put #n, , .Ties(t).infused
      Put #n, , .Ties(t).sharing
    Next t
    
    ' biologiche
    Put #n, , .nrg
    
    'custom variables we're saving
    For t = 1 To 50
      Put #n, , CInt(Len(.vars(t).Name))
      Put #n, , .vars(t).Name         '|
      Put #n, , .vars(t).value
    Next t

    Put #n, , .vnum            '| variabili private
    
    ' macchina virtuale
    Put #n, , .mem()
    k = DnaLen(rob(r).dna()): Put #n, , k
    For t = 1 To k
      Put #n, , .dna(t).tipo
      Put #n, , .dna(t).value
    Next t
    
    For t = 0 To 20
      Put #n, , .Mutables.mutarray(t)
    Next t
    
    ' informative
    Put #n, , .SonNumber
    Put #n, , sint(.Mutations)
    Put #n, , sint(.LastMut)
    Put #n, , .parent
    Put #n, , .age
    Put #n, , .BirthCycle
    Put #n, , .genenum
    Put #n, , .generation
    Put #n, , .DnaLen
    
    ' aspetto

    Put #n, , .Skin()
    Put #n, , .color
    
    ' new features
    Put #n, , .body
    Put #n, , .Bouyancy
    Put #n, , .Corpse
    Put #n, , .Pwaste
    Put #n, , .Waste
    Put #n, , .poison
    Put #n, , .venom
    Put #n, , .exist
    Put #n, , .Dead
    
    Put #n, , CInt(Len(.FName))
    Put #n, , .FName
    
    Put #n, , CInt(Len(.LastOwner))
    Put #n, , .LastOwner
    
    'EricL 5/8/2006 New feature allows for saving sims without all the mutations details
    If MDIForm1.SaveWithoutMutations Then
      Put #n, , CInt(Len(s))
      Put #n, , s
    Else
      'EricL 5/3/2006  This needs some explaining.  It's all about backward compatability.  The length of the mutation details
      'was stored as an Integer in older sim file versions.  It can overflow and go negative or even wrap positive
      'again in sims with lots of mutations.  So, we test to see if it would have overflowed and it so, we write
      'the interger 1 there instead of the actual length.  Since the actual details, being string descriptions,
      'should never have length 1, this is a signal to the sim file read routine that the real length is a Long
      'stored right after the Int.
      If CLng(Len(.LastMutDetail)) > CLng((2 ^ 15 - 1)) Then
        ' Lots of mutations.  Tell the read routine that the real length is Long valued and coming up next.
        Put #n, , CInt(1)
        Put #n, , CLng(Len(.LastMutDetail)) ' The real length
      Else
        'Not so many mutation details.  Leave the length as an Int for backward compatability
        Put #n, , CInt(Len(.LastMutDetail))
      End If
      Put #n, , .LastMutDetail
    End If
    
    'EricL 3/30/2006 Added the following line.  Looks like it was just missing.  Mutations were turned off after loading save...
    Put #n, , .Mutables.Mutations
        
    For t = 0 To 20
      Put #n, , .Mutables.Mean(t)
      Put #n, , .Mutables.StdDev(t)
    Next t
    
    Put #n, , .Mutables.CopyErrorWhatToChange
    Put #n, , .Mutables.PointWhatToChange
    
    Put #n, , .View
    Put #n, , .NewMove
    Put #n, , r  'EricL  New for 2.42.8.  Save Robot number for use in re-mapping ties and shots when re-loaded

    Put #n, , .CantSee
    Put #n, , .DisableDNA
    Put #n, , .DisableMovementSysvars
    Put #n, , .CantReproduce
    Put #n, , .shell
    Put #n, , .Slime
    Put #n, , .VirusImmune
    Put #n, , .SubSpecies
    
    If .fertilized < 0 Then .spermDNAlen = 0
    
    Put #n, , .spermDNAlen
    For t = 1 To .spermDNAlen
      Put #n, , .spermDNA(t).tipo
      Put #n, , .spermDNA(t).value
    Next t
    Put #n, , .fertilized
        
    Put #n, , .sim
    Put #n, , .AbsNum
    
    'Botsareus 2/23/2013 Rest of tie data
    Put #n, , .Multibot
    For t = 0 To MAXTIES
        Put #n, , .Ties(t).type
        Put #n, , .Ties(t).b
        Put #n, , .Ties(t).k
        Put #n, , .Ties(t).NaturalLength
    Next
    
    'Botsareus 4/9/2013 For genetic distance graph
    Put #n, , .OldGD
    
    'Panda 8/13/2013 Write chloroplasts
    Put #n, , .chloroplasts
    
    'Botsareus 12/3/2013 Write epigenetic information
    For t = 0 To 14
        Put #n, , .epimem(t)
    Next
    
    'Botsareus 1/28/2014 Write robot tag
    
    Dim blank As String * 50
    
    Put #n, , .tag
    
    'Botsareus 1/28/2014 Write if robot is using sunbelt
    
    Put #n, , sunbelt
    
    'Botsareus 3/28/2014 Write if disable chloroplasts
    
    Put #n, , .NoChlr
    
    'Botsareus 3/28/2014 Read kill resrictions
    
    Put #n, , .Chlr_Share_Delay
    
    'Botsareus 10/8/2015 Keep track of mutations from old dna file
    
    Put #n, , .OldMutations
    
    'Botsareus 6/22/2016 Actual velocity
    
    Put #n, , .actvel.x
    Put #n, , .actvel.y
    
    'write any future data here
    
    Put #n, , Fe
    Put #n, , Fe
    Put #n, , Fe
  End With
End Sub

' saves a robot dna     !!!New routine from Carlo!!!
'Botsareus 10/8/2015 Code simplification
Sub salvarob(n As Integer, path As String)
  Dim hold As String, hashed As String, a As Integer, epigene As String
  
  Close #1
  Open path For Output As #1
  hold = SaveRobHeader(n)
  
  'Botsareus 10/8/2015 New code to save epigenetic memory as gene
  
  If UseEpiGene Then
    For a = 971 To 990
      If rob(n).mem(a) <> 0 Then epigene = epigene & rob(n).mem(a) & " " & a & " store" & vbCrLf
    Next
    
    If epigene <> "" Then
      epigene = "start" & vbCrLf & epigene & "*.thisgene .delgene store" & vbCrLf & "stop"
      hold = hold + epigene
    End If
  End If
    
  savingtofile = True 'Botsareus 2/28/2014 when saving to file the def sysvars should not save
  hold = hold + DetokenizeDNA(n, 0)
  savingtofile = False
  hashed = Hash(hold, 20)
  Print #1, hold
  Print #1, ""
  Print #1, "'#hash: " + hashed
  Dim blank As String * 50
  If Left(rob(n).tag, 45) <> Left(blank, 45) Then Print #1, "'#tag:" + Left(rob(n).tag, 45) + vbCrLf
  Close #1
  
  'Botsareus 12/11/2013 Save mrates file
  Save_mrates rob(n).Mutables, extractpath(path) & "\" & extractexactname(extractname(path)) & ".mrate"
  
  If MsgBox("Do you want to change robot's name to " + extractname(path) + " ?", vbYesNo, "Robot DNA saved") = vbYes Then
    rob(n).FName = extractname(path)
  End If
End Sub


' saves a Obstacle
Private Sub SaveObstacle(n As Integer, t As Integer)
    
  Const Fe As Byte = 254

  With Obstacles.Obstacles(t)
    Put #n, , .exist
    Put #n, , .pos
    Put #n, , .Width
    Put #n, , .Height
    Put #n, , .color
    Put #n, , .vel
    
    'write any future data here
    
    Put #n, , Fe
    Put #n, , Fe
    Put #n, , Fe
    'don't you dare put anything after this!
    
  End With
End Sub

' loads an Obstacle
Private Sub LoadObstacle(n As Integer, t As Integer)
  Dim k As Integer
  Dim Fe As Byte

  With Obstacles.Obstacles(t)
    Get #n, , .exist
    Get #n, , .pos
    Get #n, , .Width
    Get #n, , .Height
    Get #n, , .color
    Get #n, , .vel
    
    'burn through any new data from a different version
    While FileContinue(n)
      Get #n, , Fe
    Wend
    
    'grab these three FE codes
    Get #n, , Fe
    Get #n, , Fe
    Get #n, , Fe
    
    'don't you dare put anything after this!
  End With
End Sub

'Saves a Shot
'New routine by EricL
Private Sub SaveShot(n As Integer, t As Long)
  Dim k As Integer
  Dim x As Integer
  
  Const Fe As Byte = 254

  With Shots(t)
    Put #n, , .Exists       ' exists?
    Put #n, , .Position         ' position vector
    Put #n, , .OldPosition        ' old position vector
    Put #n, , .velocity    ' velocity vector
    Put #n, , .parent      ' who shot it?
    Put #n, , .age         ' shot age
    Put #n, , .Energy         ' energy carrier
    Put #n, , .Range       ' shot range (the maximum .nrg ever was)
    Put #n, , .value       ' power of shot for negative shots (or amt of shot, etc.), value to write for > 0
    Put #n, , .color       ' colour
    Put #n, , .shottype    ' carried location/value couple
    Put #n, , .fromveg     ' does shot come from veg?
    Put #n, , CInt(Len(.fromSpecies))
    Put #n, , .fromSpecies  ' Which species fired the shot
    Put #n, , .memoryLocation      ' Memory location for custom poison and venom
    Put #n, , .memoryValue      ' Value to insert into custom venom location
    
    ' Somewhere to store genetic code for a virus or sperm
    If (.shottype = -7 Or .shottype = -8) And .Exists And .DnaLen > 0 Then
      Put #n, , .DnaLen
      For x = 1 To .DnaLen
        Put #n, , .dna(x).tipo
        Put #n, , .dna(x).value
      Next x
    Else
      k = 0: Put #n, , k
    End If
    
    Put #n, , .genenum     ' which gene to copy in host bot
    Put #n, , .stored      ' for virus shots (and maybe future types) this shot is stored inside the bot until it's ready to be launched
  
    'write any future data here
    
    Put #n, , Fe
    Put #n, , Fe
    Put #n, , Fe
    'don't you dare put anything after this!
  End With
End Sub

'Loads a Shot
'New routine from EricL
Private Sub LoadShot(n As Integer, t As Long)
  Dim k As Integer
  Dim x As Integer
  Dim Fe As Byte

  With Shots(t)
    Get #n, , .exist       ' exists?
    Get #n, , .pos         ' position vector
    Get #n, , .opos        ' old position vector
    Get #n, , .velocity    ' velocity vector
    Get #n, , .parent      ' who shot it?
    Get #n, , .age         ' shot age
    Get #n, , .nrg         ' energy carrier
    Get #n, , .Range       ' shot range (the maximum .nrg ever was)
    Get #n, , .value       ' power of shot for negative shots (or amt of shot, etc.), value to write for > 0
    Get #n, , .color       ' colour
    Get #n, , .shottype    ' carried location/value couple
    Get #n, , .fromveg     ' does shot come from veg?
    
    Get #n, , k: .FromSpecie = Space(k)
    Get #n, , .FromSpecie  ' Which species fired the shot
    
    Get #n, , .memloc      ' Memory location for custom poison and venom
    Get #n, , .Memval      ' Value to insert into custom venom location

    
    ' Somewhere to store genetic code for a virus
    Get #n, , k
    If k > 0 Then
      ReDim .dna(k)
      For x = 1 To k
        Get #n, , .dna(x).tipo
        Get #n, , .dna(x).value
      Next x
    End If
    
    .DnaLen = k
    
    Get #n, , .genenum     ' which gene to copy in host bot
    Get #n, , .stored      ' for virus shots (and maybe future types) this shot is stored inside the bot until it's ready to be launched
    
   'burn through any new data from a different version
    While FileContinue(n)
      Get #n, , Fe
    Wend
    
    'grab these three FE codes
    Get #n, , Fe
    Get #n, , Fe
    Get #n, , Fe
    
    'don't you dare put anything after this!
  End With
End Sub

'generate mrates file
Sub Save_mrates(mut As mutationprobs, FName As String)
  Dim m As Byte
  Open FName For Output As #1
  With mut
    Write #1, .PointWhatToChange
    Write #1, .CopyErrorWhatToChange
    For m = 0 To 10
      Write #1, .mutarray(m)
      Write #1, .Mean(m)
      Write #1, .StdDev(m)
    Next
  End With
  Close #1
End Sub

'load mrates file
Public Function Load_mrates(FName As String) As mutationprobs
  Dim m As Byte
  Open FName For Input As #1
  With Load_mrates
    Input #1, .PointWhatToChange
    Input #1, .CopyErrorWhatToChange
    For m = 0 To 10
      Input #1, .mutarray(m)
      Input #1, .Mean(m)
      Input #1, .StdDev(m)
    Next
  End With
  Close #1
End Function

Private Function sint(ByVal lval As Long) As Integer
  sint = lval Mod 32000
End Function
