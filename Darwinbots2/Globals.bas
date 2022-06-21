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
    Next
End Sub

Private Function checkvegstatus(ByVal r As Integer) As Boolean
    Dim t As Integer
    Dim splitname() As String 'just incase original species is dead
    Dim robname As String

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
                If robManager.GetExists(t) And .Veg And .age > 0 Then
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
Public Sub aggiungirob(ByVal r As Integer, ByVal x As Single, ByVal y As Single)
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
            Exit Sub
        End If
        
        rob(a).Veg = SimOpts.Specie(r).Veg
        If rob(a).Veg Then rob(a).chloroplasts = StartChlr 'Botsareus 2/12/2014 Start a robot with chloroplasts
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
        rob(a).OldMutations = 0
        rob(a).LastMut = 0
        rob(a).generation = 0
        rob(a).SonNumber = 0
        rob(a).parent = 0
        Erase rob(a).mem
        If rob(a).Fixed Then rob(a).mem(216) = 1
        robManager.SetRobotPosition a, VectorSet(x, y)
        
        rob(a).aim = rndy * PI * 2
        rob(a).mem(SetAim) = rob(a).aim * 200
        
        UpdateBotBucket a
        rob(a).nrg = SimOpts.Specie(r).Stnrg
        rob(a).Mutables = SimOpts.Specie(r).Mutables
        
        rob(a).Vtimer = 0
        rob(a).virusshot = 0
        rob(a).genenum = CountGenes(rob(a).dna)
        
        rob(a).DnaLen = DnaLen(rob(a).dna())
        rob(a).GenMut = rob(a).DnaLen / GeneticSensitivity
        
        rob(a).mem(DnaLenSys) = rob(a).DnaLen
        rob(a).mem(GenesSys) = rob(a).genenum
        
        rob(a).NoChlr = SimOpts.Specie(r).NoChlr
        
        For i = 0 To 7
            rob(a).Skin(i) = SimOpts.Specie(r).Skin(i)
        Next i
         
        rob(a).color = SimOpts.Specie(r).color
        makeoccurrlist a
    End If
End Sub

