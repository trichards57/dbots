Attribute VB_Name = "Globals"
'A temporary module for everything without a home.
Option Explicit

'G L O B A L  S E T T I N G S Botsareus 3/15/2013
Public screenratiofix As Boolean
Public bodyfix As Integer
Public reprofix As Boolean
Public chseedstartnew As Boolean
Public chseedloadsim As Boolean
Public UseSafeMode As Boolean
Public UseEpiGene As Boolean
Public GraphUp As Boolean
Public HideDB As Boolean
Public intFindBestV2 As Integer
Public UseOldColor As Boolean
Public startnovid As Boolean
'Botsareus 11/29/2013 The mutations tab
Public epireset As Boolean
Public epiresetemp As Single
Public epiresetOP As Integer
Public sunbelt As Boolean
'Botsareus 12/16/2013 Delta2 on mutations tab
Public Delta2 As Boolean
Public DeltaMainExp As Single
Public DeltaMainLn As Single
Public DeltaDevExp As Single
Public DeltaDevLn As Single
Public DeltaPM As Integer
Public DeltaWTC As Byte
Public DeltaMainChance As Byte
Public DeltaDevChance As Byte
'Botsareus 12/16/2013 Normalize DNA length
Public NormMut As Boolean
Public valNormMut As Integer
Public valMaxNormMut As Integer

'some global settings change within simulation
Public loadboylabldisp As Boolean
Public loadstartnovid As Boolean

Public tmpseed As Long 'used only by "load simulation"
Public simalreadyrunning As Boolean
Public autosaved As Boolean

Public StartChlr As Integer 'Botsareus 2/12/2014 Start repopulating robots with chloroplasts

' var structure, to store the correspondance name<->value
Public Type var
  Name As String
  value As Integer
End Type

'Constants for the graphs, which are used all over the place unfortunately -Botsareus 8/3/2012 reimplemented
Public Const POPULATION_GRAPH As Integer = 1
Public Const MUTATIONS_GRAPH As Integer = 2
Public Const AVGAGE_GRAPH As Integer = 3
Public Const OFFSPRING_GRAPH As Integer = 4
Public Const ENERGY_GRAPH As Integer = 5
Public Const DNALENGTH_GRAPH As Integer = 6
Public Const DNACOND_GRAPH As Integer = 7
Public Const MUT_DNALENGTH_GRAPH As Integer = 8
Public Const ENERGY_SPECIES_GRAPH As Integer = 9
Public Const DYNAMICCOSTS_GRAPH As Integer = 10
Public Const SPECIESDIVERSITY_GRAPH As Integer = 11
Public Const AVGCHLR_GRAPH As Integer = 12 'Botsareus 8/31/2013 Average chloroplasts graph
Public Const GENETIC_DIST_GRAPH As Integer = 13
Public Const GENERATION_DIST_GRAPH As Integer = 14
Public Const GENETIC_SIMPLE_GRAPH As Integer = 15
'Botsareus 5/24/2013 Customizable graphs
Public Const CUSTOM_1_GRAPH As Integer = 16
Public Const CUSTOM_2_GRAPH As Integer = 17
Public Const CUSTOM_3_GRAPH As Integer = 18
Public strGraphQuery1 As String
Public strGraphQuery2 As String
Public strGraphQuery3 As String
'Botsareus 5/31/2013 Special graph info
Public strSimStart As String
Public Const NUMGRAPHS = 18 'Botsareus 5/25/2013 Two more graphs, moved to globals
Public graphfilecounter(NUMGRAPHS) As Long
Public graphleft(NUMGRAPHS) As Long
Public graphtop(NUMGRAPHS) As Long
Public graphvisible(NUMGRAPHS) As Boolean
Public graphsave(NUMGRAPHS) As Boolean

Public totnvegs As Integer          ' total non vegs in sim
Public totnvegsDisplayed As Integer   ' Toggle for display purposes, so the display doesn't catch half calculated value

Public TotalChlr As Long 'Panda 8/24/2013 total number of chlroroplasts
Public maxfieldsize As Long

' Not sure where to put this function, so it's going here
' makes poff. that is, creates that explosion effect with
' some fake shots...
Public Sub makepoff(n As Integer)
  Dim an As Integer
  Dim vs As Integer
  Dim vx As Integer
  Dim vy As Integer
  Dim x As Long
  Dim y As Long
  Dim t As Byte
  For t = 1 To 20
    an = (640 / 20) * t
    vs = Random(RobSize / 40, RobSize / 30)
    vx = robManager.GetVelocity(n).x + absx(an / 100, vs, 0, 0, 0)
    vy = robManager.GetVelocity(n).y + absy(an / 100, vs, 0, 0, 0)
    With rob(n)
    x = Random(robManager.GetRobotPosition(n).x - robManager.GetRadius(n), robManager.GetRobotPosition(n).x + robManager.GetRadius(n))
    y = Random(robManager.GetRobotPosition(n).y - robManager.GetRadius(n), robManager.GetRobotPosition(n).y + robManager.GetRadius(n))
    End With
    If Random(1, 2) = 1 Then
      createshot x, y, vx, vy, -100, 0, 0, RobSize * 2, rob(n).color
    Else
      createshot x, y, vx, vy, -100, 0, 0, RobSize * 2, DBrite(rob(n).color)
    End If
  Next t
End Sub

Private Function checkvegstatus(ByVal r As Integer) As Boolean
Dim t As Integer

Dim FName As String
Dim splitname() As String 'just incase original species is dead
Dim robname As String

FName = extractname(SimOpts.Specie(r).Name)

checkvegstatus = False

If SimOpts.Specie(r).Veg = True Then

    'see if any active robots have chloroplasts
      For t = 1 To MaxRobs
        With rob(t)
            If robManager.GetExists(t) And .chloroplasts > 0 Then
            
                'remove old nick name
                splitname = Split(.FName, ")")
                'if it is a nick name only
                If Left(splitname(0), 1) = "(" And IsNumeric(Right(splitname(0), Len(splitname(0)) - 1)) Then
                    robname = splitname(1)
                Else
                    robname = .FName
                End If
                
                If SimOpts.Specie(r).Name = robname Then
                
                    checkvegstatus = True
                    Exit Function
                    
                End If
                
            End If
        End With
      Next

    'If there is no robots at all with chlr then repop everything
    
    checkvegstatus = True
    
    For t = 1 To MaxRobs
            With rob(t)
                If robManager.GetExists(t) And .Veg And .age > 0 Then 'Botsareus 11/4/2015 age test makes sure all robots spawn
                        checkvegstatus = False
                        Exit Function
                End If
            End With
    Next

End If
End Function

' not sure where to put this function, so it's going here
' adds robots on the fly loading the script of specie(r)
' if r=-1 loads a vegetable (used for repopulation)
Public Sub aggiungirob(ByVal r As Integer, ByVal x As Single, ByVal y As Single) 'Botsareus 5/22/2014 Bugfix by adding byval
  Dim a As Integer
  Dim i As Integer
  Dim anyvegy As Boolean
  
  If r = -1 Then
    'run one loop to check vegy status
    For i = 0 To SimOpts.SpeciesNum - 1
        If checkvegstatus(i) Then
            anyvegy = True
            Exit For
        End If
    Next
    If Not anyvegy Then Exit Sub
  
    Do
    r = Random(0, SimOpts.SpeciesNum - 1)  ' start randomly in the list of species
    Loop Until checkvegstatus(r)
    
    x = fRnd(SimOpts.Specie(r).Poslf * (SimOpts.FieldWidth - 60), SimOpts.Specie(r).Posrg * (SimOpts.FieldWidth - 60))
    y = fRnd(SimOpts.Specie(r).Postp * (SimOpts.FieldHeight - 60), SimOpts.Specie(r).Posdn * (SimOpts.FieldHeight - 60))
  End If
  
  If SimOpts.Specie(r).Name <> "" And SimOpts.Specie(r).path <> "Invalid Path" Then
    a = RobScriptLoad(respath(SimOpts.Specie(r).path) + "\" + SimOpts.Specie(r).Name)
    
    'Check to see if we were able to load the bot.  If we can't, the path may be wrong, the sim may have
    'come from another machine with a different install path.  Set the species path to an empty string to
    'prevent endless looping of error dialogs.
    If Not robManager.GetExists(a) Then
      SimOpts.Specie(r).path = "Invalid Path"
      GoTo getout
    End If
    
    rob(a).Veg = SimOpts.Specie(r).Veg
    If rob(a).Veg Then rob(a).chloroplasts = StartChlr 'Botsareus 2/12/2014 Start a robot with chloroplasts
    'NewMove loaded via robscriptload
    rob(a).Fixed = SimOpts.Specie(r).Fixed
    rob(a).CantSee = SimOpts.Specie(r).CantSee
    rob(a).DisableDNA = SimOpts.Specie(r).DisableDNA
    rob(a).DisableMovementSysvars = SimOpts.Specie(r).DisableMovementSysvars
    rob(a).CantReproduce = SimOpts.Specie(r).CantReproduce
    rob(a).VirusImmune = SimOpts.Specie(r).VirusImmune
    rob(a).Corpse = False
    rob(a).Dead = False
    rob(a).body = 1000
    robManager.SetRadius a, FindRadius(a)
    rob(a).Mutations = 0
    rob(a).OldMutations = 0 'Botsareus 10/8/2015
    rob(a).LastMut = 0
    rob(a).generation = 0
    rob(a).SonNumber = 0
    rob(a).parent = 0
    Erase rob(a).mem
    If rob(a).Fixed Then rob(a).mem(216) = 1
    robManager.SetRobotPosition a, VectorSet(x, y)
    
    rob(a).aim = rndy * PI * 2 'Botsareus 5/30/2012 Added code to rotate the robot on placment
    rob(a).mem(SetAim) = rob(a).aim * 200
    
    UpdateBotBucket a
    rob(a).nrg = SimOpts.Specie(r).Stnrg
   ' EnergyAddedPerCycle = EnergyAddedPerCycle + rob(a).nrg
    rob(a).Mutables = SimOpts.Specie(r).Mutables
    
    rob(a).Vtimer = 0
    rob(a).virusshot = 0
    rob(a).genenum = CountGenes(rob(a).dna)
    
    
    rob(a).DnaLen = DnaLen(rob(a).dna())
    rob(a).GenMut = rob(a).DnaLen / GeneticSensitivity 'Botsareus 4/9/2013 automatically apply genetic to inserted robots
    
    
    rob(a).mem(DnaLenSys) = rob(a).DnaLen
    rob(a).mem(GenesSys) = rob(a).genenum
    
    'Botsareus 10/8/2015 New kill restrictions
    rob(a).NoChlr = SimOpts.Specie(r).NoChlr 'Botsareus 11/1/2015 Bug fix
    
    
    For i = 0 To 7 'Botsareus 5/20/2012 fix for skin engine
      rob(a).Skin(i) = SimOpts.Specie(r).Skin(i)
    Next i
    
    rob(a).color = SimOpts.Specie(r).color
    makeoccurrlist a
  End If
getout:
End Sub

