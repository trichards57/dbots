Attribute VB_Name = "Shots_Module"
Option Explicit

' shot structure definition
Public Type Old_Shot
 exist As Boolean       ' exists?
 
 pos As Vector          ' position vector
 opos As Vector         ' old position vector
 velocity As Vector     ' velocity vector
 
 parent As Integer      ' who shot it?
 
 age As Integer         ' rob age
 
 
 nrg As Single          ' energy carrier
 Range As Single        ' shot range (the maximum .nrg ever was)
 value As Integer       ' power of shot for negative shots (or amt of shot, etc.), value to write for > 0
 
 color As Long          ' colour
 shottype As Integer    ' carried location/value couple
 
 fromveg As Boolean     ' does shot come from veg?
 FromSpecie As String   ' Which species fired the shot
 
 memloc As Integer      ' Memory location for custom poison and venom
 Memval As Integer      ' Value to insert into custom venom location
 
 dna() As Block         ' Somewhere to store genetic code for a virus or sperm
 DnaLen As Integer      ' length of DNA  stored on this shot
 genenum As Integer     ' which gene to copy in host bot
 stored As Boolean      ' for virus shots (and maybe future types) this shot is stored
                        ' inside the bot until it's ready to be launched
 flash As Boolean       ' For showing shot impacts
End Type

Public Shots() As Shot  ' array of shots
Public shotpointer As Long     ' index into the Shots array used to find new slots for new shots
'Public maxshots As Integer
Public numshots As Long       'Counter for tracking number of shots in the sim
Public ShotsThisCycle As Long ' Shots this cycle.  Only updated at end of UpdateShots()
Public maxshotarray As Long
Const shotdecay As Integer = 40 'increase to have shots lose power slower
Const ShellEffectiveness As Integer = 20 'how strong each unit of shell is
Const SlimeEffectiveness As Single = 1 / 20 'how strong each unit of slime is against viruses 'Botsareus 10/5/2015 Virus more effective
Const VenumEffectivenessVSShell As Integer = 25 'Botsareus 3/15/2013 Multiply strength of venum agenst shell
Const MinBotRadius = 0.2 'A total hack.  Used to bypass checking the rest of the bots if the collision occurred during this
                           'intial fraction of the cycle.  We assume that no bot is small enough to possibly have been hit earlier
                           'in the cycle.  We risk not detecting collisions with tiny bots in the case where the shot hits it early
                           'in the cycle, but the perf benefit of skipping the rest of the bots is significant.
Public MaxBotShotSeperation As Single
Public FlashColor(10) As Long      ' array of colors to use for flashing bots when they get shot

Private manager As New ShotManager

'
'   S H O T S   M A N A G E M E N T
'

' calculates the half brightness of a colour
' for a vaguely shiny effect in particles burst
Private Function HBrite(ByVal col As Long) As Long
  Dim b As Integer, g As Integer, r As Integer
  b = Int(col / 65536)
  col = col - (b * 65536)
  g = Int(col / 256)
  r = col - (g * 256)
  b = b / 2
  g = g / 2
  r = r / 2
  HBrite = RGB(r, g, b)
End Function

' same, but doubles
Public Function DBrite(ByVal col As Long) As Long
  Dim b As Long, g As Long, r As Long
  b = Int(col / 65536)
  col = col - (b * 65536)
  g = Int(col / 256)
  r = col - (g * 256)
  b = b + (255 - b) / 2
  g = g + (255 - g) / 2
  r = r + (255 - r) / 2
  DBrite = RGB(r, g, b)
End Function

' creates a shot shooted by robot n, with couple location/value
' returns the shot num of the shot
Public Function newshot(n As Integer, ByVal shottype As Integer, ByVal val As Single, rngmultiplier As Single, Optional offset As Boolean = False) As Long
  Dim a As Long
  Dim ran As Single
  Dim angle As Vector
  Dim ShAngle As Single
  Dim x As Integer

  a = manager.createshot()
  
  If val > 32000 Then val = 32000 ' EricL March 16, 2006 This line moved here from below to catch val before assignment
  
  Dim s As Shot
  s = manager.GetShot(a)
  s.Exists = True
  s.age = 0
  s.parent = n
  s.FromSpecie = rob(n).FName
  s.fromveg = rob(n).Veg
  s.color = rob(n).color
  s.value = Int(val)

 
  If (shottype > 0) Or (shottype = -100) Then
    s.shottype = shottype
  Else
    s.shottype = -(Abs(shottype) Mod 8)  ' EricL 6/2006 essentially Mod 8 so as to increse probabiltiy that mutations do something interesting
    If s.shottype = 0 Then s.shottype = -8 ' want multiples of -8 to be -8
  End If
  If shottype = -2 Then s.color = vbWhite
  s.memloc = rob(n).mem(835)     'location for venom to target
  s.Memval = rob(n).mem(836)     'value to insert into venom target location
  
  ran = Random(-2, 2) / 20
  
  If rob(n).mem(backshot) = 0 Then
    ShAngle = rob(n).aim        'forward shots
  Else
    ShAngle = angnorm(rob(n).aim - PI) 'backward shots
    rob(n).mem(backshot) = 0
  End If
  
  If rob(n).mem(aimshoot) <> 0 Then '0 is the same as .shoot without any aiming
    rob(n).mem(aimshoot) = rob(n).mem(aimshoot) Mod 1256
    
    ShAngle = (rob(n).aim - rob(n).mem(aimshoot) / 200)
    rob(n).mem(aimshoot) = 0
  End If
  
  ShAngle = ShAngle + Random(-20, 20) / 200
  
  angle = VectorSet(Cos(ShAngle), -Sin(ShAngle))
  s.pos = VectorAdd(rob(n).pos, VectorScalar(angle, rob(n).radius))
  
  'Botsareus 6/23/2016 Takes care of shot position bug - so it matches the painted robot position
  If offset Then
    s.pos = VectorSub(s.pos, rob(n).vel)
    s.pos = VectorAdd(s.pos, rob(n).actvel)
  End If
  
  
  s.velocity = VectorAdd(rob(n).actvel, VectorScalar(angle, 40))
  
  s.opos = VectorSub(s.pos, s.velocity)
  
  If rob(n).vbody > 10 Then
    s.nrg = Log(Abs(rob(n).vbody)) * 60 * rngmultiplier
    Dim temp As Long
    temp = (s.nrg + 40 + 1) \ 40 'divides and rounds up
    s.Range = temp
    s.nrg = temp * 40
  Else
    s.Range = rngmultiplier
    s.nrg = 40 * rngmultiplier
  End If
  
  'return the new shot
  newshot = a
  
  If shottype = -7 Then
    s.color = vbCyan
    s.genenum = val
    s.stored = True
    If Not copygene(a, s.genenum) Then
      s.exist = False
      s.stored = False
      newshot = -1
    End If
  Else
    s.stored = False
  End If
  
  If shottype = -2 Then s.nrg = val
  
  ' sperm shot
  If shottype = -8 Then
    s.dna = rob(n).dna
    s.DnaLen = rob(n).DnaLen
  End If
      
End Function

' creates a generic particle with arbitrary x & y, vx & vy, etc
Public Sub createshot(ByVal x As Long, ByVal y As Long, ByVal vx As Integer, _
  ByVal vy As Integer, loc As Integer, par As Integer, val As Single, Range As Single, col As Long)
  Dim a As Long
  
  'If IsArrayBounded(Shots) = False Then
  '  ReDim Shots(300)
  '  maxshotarray = 300
  'End If
  
  a = FirstSlot
  If a > maxshotarray Then
    shotpointer = maxshotarray ' we know the array is full.  Set the pointer to the end so it will point to the free space
    maxshotarray = CLng(maxshotarray * 1.1) ' Increase the array by 10%
    ReDim Preserve Shots(maxshotarray)
  End If
  Shots(a).parent = par
  Shots(a).FromSpecie = rob(par).FName
  Shots(a).fromveg = rob(par).Veg
  
  Shots(a).pos.x = x '+ vx
  Shots(a).pos.y = y '+ vy
  Shots(a).velocity.x = vx
  Shots(a).velocity.y = vy
  Shots(a).opos = VectorSub(Shots(a).pos, Shots(a).velocity)
    
  Shots(a).age = 0
  Shots(a).color = col
  Shots(a).exist = True
  Shots(a).stored = False
  Shots(a).DnaLen = 0
  
    
  Dim temp As Long
  temp = (Range + 40 + 1) \ 40 'divides and rounds up ie: range / (Robsize/3)
  
  Shots(a).nrg = Range + 40 + 1
  If val > 32000 Then val = 32000 ' Overflow protection
  If loc = -2 Then Shots(a).nrg = val
  Shots(a).Range = temp
  Shots(a).value = CInt(val)
  If loc > 0 Or loc = -100 Then
    Shots(a).shottype = loc
  Else
    Shots(a).shottype = -(Abs(loc) Mod 8)  ' EricL 6/2006 essentially Mod 8 so as to increse probabiltiy that mutations do something interesting
    If Shots(a).shottype = 0 Then Shots(a).shottype = -8 ' want multiples of -8 to be -8
  End If
'  If rob(par).mem(834) <= 0 Then
'    Shots(a).memloc = 0
'  Else
'    Shots(a).memloc = rob(par).mem(834) Mod 1000
'    If Shots(a).memloc = 0 Then Shots(a).memloc = 1000
'  End If
  Shots(a).memloc = rob(par).mem(834) 'Botsareus 10/6/2015 Normalized on reseaving side
  
  If Shots(a).shottype = -5 Then Shots(a).Memval = rob(Shots(a).parent).mem(839)
  
End Sub

' searches some place to insert the new shot in the
' shots array.
Private Function FirstSlot() As Long
  Dim counter As Long
      
  counter = 1
  
  While Shots(shotpointer).exist
    counter = counter + 1
    shotpointer = shotpointer + 1
    If shotpointer > maxshotarray Then shotpointer = 1
    If counter > maxshotarray Then GoTo exitloop
  Wend
exitloop:
  
  If counter > maxshotarray Then
    'maxshots = counter
    'Ran off the end of the array.  Return the array size + 1 to indicate it needs needs to be redimed.
    FirstSlot = counter
  Else
    FirstSlot = shotpointer
  End If
End Function

' calculates next shots position
Public Sub updateshots()
'moves shot then checks for collision
  Dim a As Integer
  Dim t As Long
  Dim h As Integer
  Dim dx As Integer
  Dim sx As Integer
  Dim rp As Integer
  Dim jj As Integer
  Dim ti As Single
  Dim x As Long
  Dim y As Long
  Dim tempnum As Single
  
 ' shotpointer = 1
   
  numshots = 0
  For t = 1 To maxshotarray
    'This is one of the most CPU intensive routines.  We need to make the UI responsive.
    If t Mod 250 = 0 Then DoEvents
    If t <= maxshotarray Then 'Botsareus 4/5/2016 Bug fix
    
    If Shots(t).flash Then
        Shots(t).exist = False
        Shots(t).flash = False
        Shots(t).DnaLen = 0
    End If
      If Shots(t).exist Then
        numshots = numshots + 1 ' Counts the number of existing shots each cycle for display purposes
      
        'Add the energy in the shot to the total sim energy if it is an energy shot
        If Shots(t).shottype = -2 Then TotalSimEnergy(CurrentEnergyCycle) = TotalSimEnergy(CurrentEnergyCycle) + Shots(t).nrg
              
        If (Shots(t).shottype = -100) Or (Shots(t).stored = True) Then
          h = 0  ' It's purely an ornimental shot like a poff or it's a virus shot that hasn't been fired yet
        Else
          h = NewShotCollision(t) ' go off and check for collisions with bots.
        End If
        
        'babies born into a stream of shots from its parent shouldn't die
        'from those shots.  I can't imagine this temporary imunity can be
        'exploited, so it should be safe
        If h > 0 And Not (Shots(t).parent = rob(h).parent And rob(h).age <= 1) Then
        
          'EricL 4/19/2006 Divide by zero protection for cases where the shot range is zero
          If Shots(t).Range = 0 Then
            tempnum = Shots(t).age + 1 ' / (.range + 1)
          Else
            tempnum = Shots(t).age / Shots(t).Range
          End If
          
          'this below is horribly complicated:  allow me to explain:
          'nrg dissipates in a non-linear fashion.  Very little nrg disappears until you
          'get near the last 10% of the journey or so.
          'Don't dissipate nrg if nrg shots last forever.
         If Not SimOpts.NoShotDecay Or Shots(t).shottype <> -2 Then
          If Not (Shots(t).shottype = -4 And SimOpts.NoWShotDecay) Then 'Botsareus 9/29/2013 Do not decay waste shots
           Shots(t).nrg = Shots(t).nrg * (Atn(tempnum * shotdecay - shotdecay)) / Atn(-shotdecay)
          End If
         End If
          
        
          If Shots(t).shottype > 0 Then
          'Botsareus 10/6/2015 Minor bug fixing and redundent code removal
            Shots(t).shottype = (Shots(t).shottype - 1) Mod 1000 + 1 ' EricL 6/2006 Mod 1000 so as to increse probabiltiy that mutations do something interesting
            
            If Shots(t).shottype <> DelgeneSys Then
    
              If (Shots(t).nrg / 2 > rob(h).poison) Or (rob(h).poison = 0) Then
                rob(h).mem(Shots(t).shottype) = Shots(t).value
              Else
                createshot Shots(t).pos.x, Shots(t).pos.y, -Shots(t).velocity.x, -Shots(t).velocity.y, -5, h, Shots(t).nrg / 2, Shots(t).Range * 40, vbYellow
                rob(h).poison = rob(h).poison - (Shots(t).nrg / 2) * 0.9
                rob(h).Waste = rob(h).Waste + (Shots(t).nrg / 2) * 0.1
                If rob(h).poison < 0 Then rob(h).poison = 0
                rob(h).mem(poison) = rob(h).poison
              End If
              
            End If
          Else
            ' Shots(t).shottype = -(Abs(Shots(t).shottype) Mod 8)  ' EricL 6/2006 essentially Mod 8 so as to increse probabiltiy that mutations do something interesting
            ' If Shots(t).shottype = 0 Then Shots(t).shottype = -8 ' want multiples of -8 to be -8
            Select Case Shots(t).shottype
            'Problem with this: returning nrg shots appear where the shot would have been instead of where
            'it hit the bot - EricL 5/20/2006 - Not anymore as of 2.42.5!
              Case -1: releasenrg h, t
              Case -2: takenrg h, t
              Case -3: takeven h, t
              Case -4: takewaste h, t
              Case -5: takepoison h, t
              Case -6: releasebod h, t
              Case -7: addgene h, t
              Case -8: takesperm h, t ' bot hit by a sperm shot for sexual reproduction
             End Select
          End If
          taste h, Shots(t).opos.x, Shots(t).opos.y, Shots(t).shottype
          Shots(t).flash = True
                
        End If
        If numObstacles > 0 Then DoShotObstacleCollisions t
        Shots(t).opos = Shots(t).pos
        Shots(t).pos = VectorAdd(Shots(t).pos, Shots(t).velocity) 'Euler integration
        
        'Age shots unless we are not decaying them.  At some point, we may want to see how old shots are, so
        'this may need to be changed at some point but for now, it lets shots never die by never growing old.
        'Always age Poff shots
        If (SimOpts.NoShotDecay And Shots(t).shottype = -2) Or (Shots(t).stored) Then
        Else
         If Shots(t).shottype = -4 And SimOpts.NoWShotDecay Then
         Else
          Shots(t).age = Shots(t).age + 1
         End If
        End If
        
        If Shots(t).age > Shots(t).Range And Not Shots(t).flash Then 'Botsareus 9/12/2016 Bug fix
          Shots(t).exist = False ' Kill shots once they reach maturity
          Shots(t).DnaLen = 0
        End If
          
      End If
      
    End If
    Next t
    
    ' Here we test for sparsity of the shots array.  If the number of shots is less than 70% of the array size, then we
    ' compact the array and reset maxshotarray
    If (numshots < (maxshotarray * 0.7)) And (maxshotarray > 100) Then
      CompactShots
      If numshots < 90 Then
        maxshotarray = CLng(100)
      Else
        maxshotarray = CLng(numshots * 1.2)
      End If
      shotpointer = numshots ' set the shot pointer at the beginning of the free space in the newly shrunk array
      ReDim Preserve Shots(maxshotarray)
    End If
  ShotsThisCycle = numshots
End Sub
Public Sub CompactShots()
  Dim i As Long
  Dim j As Long
  Dim x As Integer
  
  j = 1
  For i = 1 To maxshotarray
    If Shots(i).exist Then
      If Shots(i).stored Then
        If rob(Shots(i).parent).exist Then
          rob(Shots(i).parent).virusshot = j
        Else
          Shots(i).exist = False
          Shots(i).stored = False
          Shots(i).DnaLen = 0
        End If
      End If
      If i <> j Then
        If (Shots(j).shottype = -8 Or Shots(j).shottype = -7) And Shots(i).DnaLen > 0 Then
          ReDim Shots(j).dna(Shots(i).DnaLen)
        End If
        Shots(j) = Shots(i)
        Shots(i).exist = False
        Shots(i).stored = False
        Shots(i).DnaLen = 0
        'ReDim Shots(i).DNA(1) ' 1 so as to not hit the bounded routine exception everytime
      End If
      j = j + 1
    End If
  Next i
End Sub
Public Sub Decay(n As Integer) 'corpse decaying as waste shot, energy shot or no shot
  Dim SH As Integer
  Dim va As Single
  rob(n).DecayTimer = rob(n).DecayTimer + 1
  If rob(n).DecayTimer >= SimOpts.Decaydelay Then
    rob(n).DecayTimer = 0
    
    rob(n).aim = rndy * 2 * PI
    rob(n).aimvector = VectorSet(Cos(rob(n).aim), Sin(rob(n).aim))
    
    If rob(n).body > SimOpts.Decay / 10 Then
        va = SimOpts.Decay
      ElseIf rob(n).body > 0 Then
        va = rob(n).body
      Else: va = 0
      End If
      
    If SimOpts.DecayType = 2 And va <> 0 Then
      SH = -4
      newshot n, SH, va, 1
    End If
    
    If SimOpts.DecayType = 3 And va <> 0 Then
      SH = -2
      newshot n, SH, va, 1
    End If
    

    rob(n).body = rob(n).body - SimOpts.Decay / 10
    rob(n).radius = FindRadius(n)
  End If
End Sub
Public Sub defacate(n As Integer) 'only used to get rid of massive amounts of waste
  Dim SH As Integer
  Dim va As Single
  SH = -4
  va = 200
  
  If va > rob(n).Waste Then va = rob(n).Waste
  If rob(n).Waste > 32000 Then rob(n).Waste = 31500: va = 500
  
  rob(n).Waste = rob(n).Waste - va
  rob(n).nrg = rob(n).nrg - (SimOpts.Costs(SHOTCOST) * SimOpts.Costs(COSTMULTIPLIER)) / (IIf(rob(n).numties < 0, 0, rob(n).numties) + 1)
  newshot n, SH, va, 1, True
  rob(n).Pwaste = rob(n).Pwaste + va / 1000
End Sub

' robot n, hit by shot t, releases energy
Public Sub releasenrg(ByVal n As Integer, ByVal t As Long)
  'n=robot number
  't=shot number
  Dim vel As Vector
  
  Dim vs As Integer
  Dim vr As Single
  Dim power As Single
  Dim Range As Single
  Dim scalingfactor As Single
  Dim Newangle As Single
  Dim startingPos As Vector
  Dim incoming As Vector
  Dim offcenter As Single
  Dim shotNow As Vector
  Dim x As Single
  Dim y As Single
  Dim angle As Single
  Dim relVel As Vector
  Dim EnergyLost As Single
  Dim a As Long
  
  a = FirstSlot
    
  If rob(n).nrg <= 0.5 Then
   ' rob(n).Dead = True ' Don't kill them here so they can be corpses.  Still might have body.
    GoTo getout
  End If

  vel = VectorSub(rob(n).actvel, Shots(t).velocity) 'negative relative velocity of shot hitting bot 'Botsareus 6/22/2016 Now based on robots actual velocity
                                                             'the shot to the hit bot
  vel = VectorAdd(vel, VectorScalar(rob(n).actvel, 0.5)) 'then add in half the velocity of hit robot
 
  If SimOpts.EnergyExType Then
    If Shots(t).Range = 0 Then ' Divide by zero protection
      power = 0#
    Else
      power = CSng(Shots(t).value) * Shots(t).nrg / CSng((Shots(t).Range * (RobSize / 3))) * SimOpts.EnergyProp
    End If
    If Shots(t).nrg < 0 Then GoTo getout
  Else
    power = SimOpts.EnergyFix
  End If
    
 'If power > rob(n).nrg + rob(n).poison And rob(n).nrg > 0 Then
 '  power = rob(n).nrg + rob(n).poison
 'End If
  
  If rob(n).Corpse Then power = power * 0.5 'half power against corpses.  Most of your shot is wasted 'The only way I can see this happening is if something tie injected energy into corpse
  
  Range = Shots(t).Range * 2 'returned energy shots have twice the range as the shot that it came from (but half the velocity)
  
  If rob(n).poison > power Then 'create poison shot
    createshot Shots(t).pos.x, Shots(t).pos.y, vel.x, vel.y, -5, n, power, Range * (RobSize / 3), vbYellow
  '  rob(n).Waste = rob(n).Waste + (power * 0.1)
    rob(n).poison = rob(n).poison - (power * 0.9)
    If rob(n).poison < 0 Then rob(n).poison = 0
    rob(n).mem(poison) = rob(n).poison
  Else ' create energy shot
       
    EnergyLost = power * 0.9
    If EnergyLost > rob(n).nrg Then
   '   EnergyLostPerCycle = EnergyLostPerCycle - rob(n).nrg
      power = rob(n).nrg
      rob(n).nrg = 0
    Else
      rob(n).nrg = rob(n).nrg - EnergyLost  'some of shot comes from nrg
    '  EnergyLostPerCycle = EnergyLostPerCycle - EnergyLost
    End If
  
    EnergyLost = power * 0.01
    If EnergyLost > rob(n).body Then
   '   EnergyLostPerCycle = EnergyLostPerCycle - (rob(n).body * 10)
      rob(n).body = 0
    Else
      rob(n).body = rob(n).body - EnergyLost 'some of shot comes from body
     ' EnergyLostPerCycle = EnergyLostPerCycle - EnergyLost * 10
    End If
    
    ' pass local vars to createshot so that no Shots array elements are on the stack in case the Shots array gets redimmed
    x = Shots(t).pos.x
    y = Shots(t).pos.y
    
    createshot x, y, vel.x, vel.y, -2, n, power, Range * (RobSize / 3), vbWhite
    rob(n).radius = FindRadius(n)
  End If
  
  If rob(n).body <= 0.5 Or rob(n).nrg <= 0.5 Then
    rob(n).Dead = True
    rob(Shots(t).parent).Kills = rob(Shots(t).parent).Kills + 1
    rob(Shots(t).parent).mem(220) = rob(Shots(t).parent).Kills
  End If
getout:
End Sub
Private Sub releasebod(ByVal n As Integer, ByVal t As Long) 'a robot is shot by a -6 shot and releases energy directly from body points
  'much more effective against a corpse
  Dim vel As Vector
  Dim Range As Single
  Dim power As Single
  Dim shell As Single
  Dim EnergyLost As Single
  
  
  'If rob(n).body <= 0 Or rob(n).wall Then goto getout
  If rob(n).body <= 0 Then GoTo getout
  
  
  
  vel = VectorSub(rob(n).actvel, Shots(t).velocity) 'negative relative velocity of shot hitting bot 'Botsareus 6/22/2016 Now based on robots actual velocity
                                                 'the shot to the hit bot
  vel = VectorAdd(vel, VectorScalar(rob(n).actvel, 0.5)) 'then add in half the velocity of hit robot
 ' vel = VectorScalar(VectorSub(rob(n).vel, Shots(t).velocity), 0.5) 'half the relative velocity of
                                                                     'the shot to the hit bot
  'vel = VectorAdd(vel, rob(n).vel) 'then add in the velocity of hit robot
  
  If SimOpts.EnergyExType Then
    If Shots(t).Range = 0 Then ' Divide by zero protection
      power = 0
    Else
      power = CSng(Shots(t).value) * Shots(t).nrg / CSng((Shots(t).Range * (RobSize / 3))) * SimOpts.EnergyProp
    End If
  Else
    power = SimOpts.EnergyFix
  End If
   
  If power > 32000 Then power = 32000
  
  shell = rob(n).shell * CSng(ShellEffectiveness)
  
  If power > ((rob(n).body * 10) / 0.8 + shell) Then
    power = (rob(n).body * 10) / 0.8 + shell
  End If
  
  If power < shell Then
    rob(n).shell = rob(n).shell - power / ShellEffectiveness
    If rob(n).shell < 0 Then rob(n).shell = 0
    rob(n).mem(823) = rob(n).shell
    GoTo getout
  Else
    rob(n).shell = rob(n).shell - power / ShellEffectiveness
    If rob(n).shell < 0 Then rob(n).shell = 0
    rob(n).mem(823) = rob(n).shell
    power = power - shell
  End If
  
  If power <= 0 Then GoTo getout
  
  Range = Shots(t).Range * 2   'new range formula based on range of incoming shot
  
  ' create energy shot
  If rob(n).Corpse = True Then
    power = power * 4 'So effective against corpses it makes me siiiiiinnnnnggg
    If power > rob(n).body * 10 Then power = rob(n).body * 10
    rob(n).body = rob(n).body - power / 10      'all energy comes from body
  '  EnergyLostPerCycle = EnergyLostPerCycle - power
    rob(n).radius = FindRadius(n)
  Else
    Dim leftover As Single
    
    leftover = 0
    EnergyLost = power * 0.2
    If EnergyLost > rob(n).nrg Then
   '   EnergyLostPerCycle = EnergyLostPerCycle - rob(n).nrg
      leftover = EnergyLost - rob(n).nrg
      rob(n).nrg = 0
    Else
      rob(n).nrg = rob(n).nrg - EnergyLost  'some of shot comes from nrg
   '   EnergyLostPerCycle = EnergyLostPerCycle - EnergyLost
    End If
    
    EnergyLost = power * 0.08
    If EnergyLost > rob(n).body Then
   '   EnergyLostPerCycle = EnergyLostPerCycle - (rob(n).body * 10)
      leftover = leftover + EnergyLost - rob(n).body * 10
      rob(n).body = 0
    Else
      rob(n).body = rob(n).body - EnergyLost 'some of shot comes from body
   '   EnergyLostPerCycle = EnergyLostPerCycle - EnergyLost * 10
    End If

    With rob(n)
    If leftover > 0 Then
      If .nrg > 0 And .nrg > leftover Then
        .nrg = .nrg - leftover
      '  EnergyLostPerCycle = EnergyLostPerCycle - leftover
        leftover = 0
      ElseIf .nrg > 0 And .nrg < leftover Then
        leftover = leftover - .nrg
      '  EnergyLostPerCycle = EnergyLostPerCycle - rob(n).nrg
        .nrg = 0
      End If

      If .body > 0 And .body * 10 > leftover Then
        .body = .body - leftover * 0.1
     '   EnergyLostPerCycle = EnergyLostPerCycle - leftover
        leftover = 0
      ElseIf rob(n).body > 0 And rob(n).body * 10 < leftover Then
     '   EnergyLostPerCycle = EnergyLostPerCycle - (rob(n).body * 10)
        .body = 0
      End If
    End If
    End With
    rob(n).radius = FindRadius(n)
  End If
  
  If rob(n).body <= 0.5 Or rob(n).nrg <= 0.5 Then
    rob(n).Dead = True
    rob(Shots(t).parent).Kills = rob(Shots(t).parent).Kills + 1
    rob(Shots(t).parent).mem(220) = rob(Shots(t).parent).Kills
  End If
  
  createshot Shots(t).pos.x, Shots(t).pos.y, vel.x, vel.y, -2, n, power, Range * (RobSize / 3), vbWhite
getout:
End Sub

' robot n takes the energy carried by shot t
Private Sub takenrg(ByVal n As Integer, ByVal t As Long)
  Dim partial As Single
  Dim overflow As Single
     
  'If rob(n).Corpse Or rob(n).wall Then goto getout
  If rob(n).Corpse Then GoTo getout
  
  If Shots(t).Range < 0.00001 Then
    partial = 0
  Else
   ' If SimOpts.NoShotDecay Then
      partial = Shots(t).nrg
   ' Else
   '   partial = CSng(Shots(t).nrg / CSng(Shots(t).Range * (RobSize / 3)) * Shots(t).value)
   ' End If
  End If

  If (rob(n).nrg + partial * 0.95) > 32000 Then
   overflow = rob(n).nrg + (partial * 0.95) - 32000
   rob(n).nrg = 32000
  Else
    rob(n).nrg = rob(n).nrg + partial * 0.95       '95% of energy goes to nrg
  End If
  
  If (rob(n).body + partial * 0.004) + (overflow * 0.1) > 32000 Then
    rob(n).body = 32000
  Else
    rob(n).body = rob(n).body + (partial * 0.004) + (overflow * 0.1)      '4% goes to body
  End If
  
  rob(n).Waste = rob(n).Waste + partial * 0.01  '1% goes to waste
 
  'Shots(t).Exist = False
  rob(n).radius = FindRadius(n)
getout:
End Sub
'  robot takes a venomous shot and becomes seriously messed up
Private Sub takeven(ByVal n As Integer, ByVal t As Long)
  Dim power As Single
  Dim temp As Single
   
  'If rob(n).Corpse Or rob(n).wall Then goto getout
  If rob(n).Corpse Then GoTo getout
  
  power = CSng(Shots(t).nrg / CSng((Shots(t).Range * (RobSize / 3))) * Shots(t).value)
  
  If power < 1 Then GoTo getout
  
  If Shots(t).FromSpecie = rob(n).FName Then   'Robot is immune to venom from his own species
    rob(n).venom = rob(n).venom + power   'Robot absorbs venom fired by conspec
    
    'EricL 4/10/2006 This line prevents an overflow when power is too large
    If ((rob(n).venom) > 32000) Then rob(n).venom = 32000
    
    rob(n).mem(825) = rob(n).venom
  Else
    power = power * VenumEffectivenessVSShell  'Botsareus 3/6/2013 max power for venum is capped at 100 so I multiply to get an average
    If power < rob(n).shell * ShellEffectiveness Then
      rob(n).shell = rob(n).shell - power / ShellEffectiveness
      rob(n).mem(823) = rob(n).shell
      GoTo getout 'Botsareus 3/6/2013 Exit sub if enough shell
    Else
      temp = power
      power = power - rob(n).shell * ShellEffectiveness
      rob(n).shell = rob(n).shell - temp / ShellEffectiveness
      If rob(n).shell < 0 Then rob(n).shell = 0
      rob(n).mem(823) = rob(n).shell
    End If
    power = power / VenumEffectivenessVSShell 'Botsareus 3/6/2013 after shell conversion devide
    
    If power < 1 Then GoTo getout
    
    rob(n).Paralyzed = True
    
    'EricL - Following lines added March 15, 2006 to avoid Paracount being overflowed.
    If ((rob(n).Paracount + power) > 32000) Then
      rob(n).Paracount = 32000
    Else
      rob(n).Paracount = rob(n).Paracount + power
    End If
        
    If Shots(t).memloc > 0 Then 'Botsareus 10/6/2015 Minor bug fixing
      rob(n).Vloc = (Shots(t).memloc - 1) Mod 1000 + 1
      If rob(n).Vloc = 340 Then rob(n).Vloc = 0 'protection from delgene attacks Botsareus 10/7/2015 Moved here after mod
    Else
      Do
       rob(n).Vloc = Random(1, 1000)
      Loop Until rob(n).Vloc <> 340
    End If
    
    rob(n).Vval = Shots(t).Memval
  End If
  'Shots(t).Exist = False
getout:
End Sub
'  Robot n takes shot t and adds its value to his waste reservoir
Private Sub takewaste(ByVal n As Integer, ByVal t As Long)
  Dim power As Single
  
'  If rob(n).wall Then goto getout
  
  power = Shots(t).nrg / (Shots(t).Range * (RobSize / 3)) * Shots(t).value
  If power < 0 Then GoTo getout
  rob(n).Waste = rob(n).Waste + power
 ' Shots(t).Exist = False
getout:
End Sub
' Robot receives poison shot and becomes disorientated
Private Sub takepoison(ByVal n As Integer, ByVal t As Long)
  Dim power As Single
   
  'If rob(n).Corpse Or rob(n).wall Then goto getout
  If rob(n).Corpse Then GoTo getout
  
  power = CSng(Shots(t).nrg / CSng((Shots(t).Range * (RobSize / 3))) * Shots(t).value)
  
  If power < 1 Then GoTo getout
  
  If Shots(t).FromSpecie = rob(n).FName Then    'Robot is immune to poison from his own species
    rob(n).poison = rob(n).poison + power 'Robot absorbs poison fired by conspecs
    If rob(n).poison > 32000 Then rob(n).poison = 32000
    rob(n).mem(827) = rob(n).poison
  Else
    rob(n).Poisoned = True
    rob(n).Poisoncount = rob(n).Poisoncount + power / 1.5 'Botsareus 6/24/3016 Div by 1.5 to make poison shots more proportenal to venom shots
    If rob(n).Poisoncount > 32000 Then rob(n).Poisoncount = 32000
    If Shots(t).memloc > 0 Then 'Botsareus 10/6/2015 Minor bug fixing
      rob(n).Ploc = (Shots(t).memloc - 1) Mod 1000 + 1
      If rob(n).Ploc = 340 Then rob(n).Ploc = 0 'protection from delgene attacks Botsareus 10/7/2015 Moved here after mod
    Else
      Do
       rob(n).Ploc = Random(1, 1000)
      Loop Until rob(n).Ploc <> 340
    End If
    rob(n).Pval = Shots(t).Memval
  End If
'  Shots(t).Exist = False
getout:
End Sub

'Robot is hit by sperm shot and becomes fertilized for potential sexual reproduction
Private Sub takesperm(ByVal n As Integer, ByVal t As Long)
If rob(n).fertilized < -10 Then Exit Sub 'block sex repro when necessary

  Dim x As Integer
  
  If Shots(t).DnaLen = 0 Then GoTo getout
  rob(n).fertilized = 10                      ' bots stay fertilized for 10 cycles currently
  rob(n).mem(SYSFERTILIZED) = 10
  ReDim rob(n).spermDNA(Shots(t).DnaLen)      ' copy the male's DNA to the female's bot structure
  rob(n).spermDNA = Shots(t).dna
  rob(n).spermDNAlen = Shots(t).DnaLen
getout:
End Sub

'EricL 5/16/2006 Checks for collisions between shots and bots.  Takes into consideration
'motion of target bot as well as shots which potentially pass through the target bot during the cycle
'Argument: The shot number to check
'Returns: bot number of the hit bot if a collison occurred, 0 otherwise
'Side Effect: On a hit, changes the shot position to be the point of impact with the bot
Private Function NewShotCollision(shotnum As Long) As Integer
  Dim robnum As Integer
  Dim B0 As Vector 'Position of bot at time 0
  Dim b As Vector 'Position of bot at time 0 < t < 1
  Dim S0 As Vector 'Position of shot at time 0
  Dim S1 As Vector 'Position of shot at time 1
  Dim s As Vector 'Position of shot at time 0 < t < 1
  Dim vs As Vector 'Velocity of the shot
  Dim vb As Vector 'Velocity of the bot
  Dim d As Vector 'Vector from bot center to shot at time 0
  Dim D2 As Single
  Dim r As Single 'Bot radius
  Dim t As Single 'Loop counter
  Dim hitTime As Single ' time in the cycle when collision occurred.
  Dim earliestCollision As Single 'Used to find which bot was hit earliest in the cycle.
                                  'The time in the cycle at which the earliest collision with the shot occurred.
  Dim time0 As Single
  Dim time1 As Single
  Dim p As Vector 'Position Vector - Realtive positions of bot and shot over time
  Dim L1 As Single
  Dim P2 As Single
  Dim x As Single
  Dim y As Single
  Dim DdotP As Single
  Dim usetime0 As Boolean
  Dim usetime1 As Boolean
  
  ' Check for collisions with the field edges
  With Shots(shotnum)
    If SimOpts.Updnconnected = True Then
      If .pos.y > SimOpts.FieldHeight Then
        .pos.y = .pos.y - SimOpts.FieldHeight
      ElseIf .pos.y < 0 Then
        .pos.y = .pos.y + SimOpts.FieldHeight
      End If
    Else
      If .pos.y > SimOpts.FieldHeight Then
        .pos.y = SimOpts.FieldHeight
        .velocity.y = -1 * Abs(.velocity.y)
      ElseIf .pos.y < 0 Then
        .pos.y = 0
        .velocity.y = Abs(.velocity.y)
      End If
    End If
     
    If SimOpts.Dxsxconnected = True Then
      If .pos.x > SimOpts.FieldWidth Then
        .pos.x = .pos.x - SimOpts.FieldWidth
      ElseIf .pos.x < 0 Then
        .pos.x = .pos.x + SimOpts.FieldWidth
      End If
    Else
      If .pos.x > SimOpts.FieldWidth Then
        .pos.x = SimOpts.FieldWidth
        .velocity.x = -1 * Abs(.velocity.x)
      ElseIf .pos.x < 0 Then
        .pos.x = 0
        .velocity.x = Abs(.velocity.x)
      End If
    End If
  End With
  
    
  'Initialize the return value in case no collision is found.
  NewShotCollision = 0
   
  'Initialize that the earliest collision to 100 to indicate no collision has been detected
  earliestCollision = 100
  
  S0 = Shots(shotnum).pos
  vs = Shots(shotnum).velocity
  
  For robnum = 1 To MaxRobs ' Walk through all the bots
  
    'Make sure the bot is eligable to be hit by the shot.  It has to exist, it can't have been the one who
    'fired the shot, it can't be a wall bot and it has to be close enough that an impact is possible.  Note that for perf reasons we
    'ignore edge cases here where the field is a torus and a shot wraps around so it's possible to miss collisons in such cases.
    If rob(robnum).exist And (Shots(shotnum).parent <> robnum) And _
     (Abs(Shots(shotnum).opos.x - rob(robnum).pos.x) < MaxBotShotSeperation And Abs(Shots(shotnum).opos.y - rob(robnum).pos.y) < MaxBotShotSeperation) Then
        
        r = rob(robnum).radius ' + 5 ' Tweak the bot radius up a bit to handle the issue with bots appearing a little larger than then are
       
        
        'Note that this routine is called before the position for both the bot and the shot is updated this cycle.  This means
        'we are looking forward in time, from the current positions to where they will be at the end of this cycle.  This is why
        'we can use .pos and not .opos
        B0 = rob(robnum).pos
        
        'Botsareus 6/22/2016 The robots actual velocity and non collision velocity can be different - correct here
        B0 = VectorSub(B0, rob(robnum).vel)
        B0 = VectorAdd(B0, rob(robnum).actvel)
        
        p = VectorSub(S0, B0)
        
        If VectorMagnitude(p) < r Then ' shot is inside the target at Time 0.  Did we miss the entry last cycle?  How?
          hitTime = 0
          earliestCollision = 0
          NewShotCollision = robnum
          GoTo FinialCollisionDetected
        End If
        
        vb = rob(robnum).actvel
        d = VectorSub(vs, vb) ' Vector of velocity change from both bot and shot over time t
        P2 = VectorMagnitudeSquare(p) ' |P|^2
          
        D2 = VectorMagnitudeSquare(d) ' |D|^2
        If D2 = 0 Then GoTo CheckRestOfBots
        DdotP = Dot(d, p)
        x = -DdotP
        y = DdotP ^ 2 - D2 * (P2 - r ^ 2)
        
        If y < 0 Then GoTo CheckRestOfBots ' No collision
        
        y = Sqr(y)
                
        time0 = (x - y) / D2
        time1 = (x + y) / D2
        
        usetime0 = False
        usetime1 = False
      
        If Not (time0 <= 0 Or time0 >= 1) Then usetime0 = True
        If Not (time1 <= 0 Or time1 >= 1) Then usetime1 = True
        If (Not usetime0) And (Not usetime1) Then
          GoTo CheckRestOfBots
        ElseIf usetime0 And usetime1 Then
          hitTime = Min(time0, time1)
          NewShotCollision = robnum
        ElseIf usetime0 Then
          hitTime = time0
          NewShotCollision = robnum
        Else
          hitTime = time1
          NewShotCollision = robnum
        End If
                
        If hitTime < earliestCollision Then earliestCollision = hitTime
                 
        'If the collision occurred early enough in the cycle, we can assume no other bot could have been hit ealier and we can
        'skip checking the rest of the bots.  This is all about perf.
        If earliestCollision <= MinBotRadius Then
          GoTo FinialCollisionDetected
        Else
          GoTo CheckRestOfBots
        End If
    End If
'We jump here if we found a collision with the current bot, but it was late enough in the cycle that another
'bot could have been hit earlier in the cycle, so we keep checking the rest of the bots
'Or if we have ruled out a possibile collision between this shot and the current bot.
CheckRestOfBots:
  Next robnum
'We jump here if we are confident that the collision occurred early enough in the cycle that no other bot could have been
'hit before this one.  Note that this is sensitive to shot speed and minumum bot radius
FinialCollisionDetected:
  If earliestCollision <= 1 Then
    'This is a total hack, but if we found a collision, any collision, then we set the position of the shot to be the point of the earliest
    'collision so that in the case where a return shot is generated, that return shot starts from the point of impact and not
    'from wherever the shot would have ended up at the end of the cycle had it not collided (which it did!)
    Shots(shotnum).pos = VectorAdd(VectorScalar(vs, earliestCollision), S0)
  End If
End Function

'Botsareus 10/5/2015 Bug fix for negative values in virus
Public Sub Vshoot(n As Integer, thisshot As Long)
'here we shoot a virus
  
Dim tempa As Single
Dim ShAngle As Single

  If Not Shots(thisshot).exist Then GoTo getout
  If Not Shots(thisshot).stored Then GoTo getout
  
  If rob(n).mem(VshootSys) < 0 Then rob(n).mem(VshootSys) = 1
  
  tempa = CSng(rob(n).mem(VshootSys)) * 20# '.vshoot * 20
  If tempa > 32000 Then tempa = 32000
  If tempa < 0 Then tempa = 0
    
  Shots(thisshot).nrg = tempa
  rob(n).nrg = rob(n).nrg - (tempa / 20#) - (SimOpts.Costs(SHOTCOST) * SimOpts.Costs(COSTMULTIPLIER))
  
  Shots(thisshot).Range = 11 + CInt((rob(n).mem(VshootSys)) / 2)
  rob(n).nrg = rob(n).nrg - CSng(rob(n).mem(VshootSys)) - (SimOpts.Costs(SHOTCOST) * SimOpts.Costs(COSTMULTIPLIER))
    
  With Shots(thisshot)
    ShAngle = Random(1, 1256) / 200
    .stored = False
    .pos.x = (rob(n).pos.x + Cos(ShAngle) * rob(n).radius)
    .pos.y = (rob(n).pos.y - Sin(ShAngle) * rob(n).radius)
  
    .velocity.x = absx(ShAngle, RobSize / 3, 0, 0, 0) ' set shot speed x seems to not work well at high bot speeds
    .velocity.y = absy(ShAngle, RobSize / 3, 0, 0, 0) ' set shot speed y
  
    .velocity.x = .velocity.x + rob(n).actvel.x
    .velocity.y = .velocity.y + rob(n).actvel.y
    
    .opos.x = .pos.x - .velocity.x
    .opos.y = .pos.y - .velocity.y
  End With
getout:
End Sub

Public Function MakeVirus(robn As Integer, ByVal gene As Integer) As Boolean
  rob(robn).virusshot = newshot(robn, -7, Int(gene), 1)
  If rob(robn).virusshot > 0 Then
    MakeVirus = True
  Else
    MakeVirus = False
  End If
End Function

' copy gene number p from robot that fired shot n into shot n dna (virus)
Public Function copygene(n As Long, ByVal p As Integer) As Boolean
  Dim t As Integer
  Dim parent As Integer
  Dim genelen As Integer
  Dim GeneStart As Long
  Dim GeneEnding As Integer
  
  parent = Shots(n).parent
  
  If (p > rob(parent).genenum) Or p < 1 Then
    ' target gene is beyond the DNA bounds
    copygene = False
    GoTo getout
  End If
  
  GeneStart = genepos(rob(parent).dna, p)
  GeneEnding = GeneEnd(rob(parent).dna, GeneStart)
  genelen = GeneEnding - GeneStart + 1
  
  If genelen < 1 Then
    copygene = False
    GoTo getout
  End If
  
  ReDim Shots(n).dna(genelen)
  
  ' Put an end on it just in case...
 ' Shots(n).DNA(genelen).tipo = 10
  'Shots(n).DNA(genelen).value = 1
  
  For t = 0 To genelen - 1
    Shots(n).dna(t) = rob(parent).dna(GeneStart + t)
  Next t
  
  Shots(n).DnaLen = genelen
  
  copygene = True
getout:
End Function
' adds gene from shot p to robot n dna
Public Function addgene(ByVal n As Integer, ByVal p As Long) As Integer
  Dim t As Long
  Dim Insert As Long
  Dim vlen As Long   'length of the DNA code of the virus
  Dim Position As Integer   'gene position to insert the virus
  Dim power As Single
  
  'Dead bodies and virus immune bots can't catch a virus
  If rob(n).Corpse Or (rob(n).VirusImmune) Then GoTo getout
  
  power = Shots(p).nrg / (Shots(p).Range * RobSize / 3) * Shots(p).value
  
  If power < rob(n).Slime * SlimeEffectiveness Then
    rob(n).Slime = rob(n).Slime - power / SlimeEffectiveness
    GoTo getout
  Else
    rob(n).Slime = rob(n).Slime - power / SlimeEffectiveness
    power = power - rob(n).Slime * SlimeEffectiveness
    If rob(n).Slime < 0.5 Then rob(n).Slime = 0
  End If
  
  Position = Random(0, rob(n).genenum)                  'randomize the gene number
  If Position = 0 Then
    Insert = 0
  Else
    Insert = GeneEnd(rob(n).dna, genepos(rob(n).dna, Position))
    If Insert = (rob(n).DnaLen) Then
      Insert = rob(n).DnaLen
    End If
  End If
  
'  vlen = DnaLen(Shots(P).DNA())
  vlen = Shots(p).DnaLen
  
  If MakeSpace(rob(n).dna, Insert, vlen) Then 'Moves genes back to make space
    For t = Insert To Insert + vlen - 1
      rob(n).dna(t + 1) = Shots(p).dna(t - Insert)
    Next t
  End If
  
  makeoccurrlist n
  rob(n).DnaLen = DnaLen(rob(n).dna())
  rob(n).genenum = CountGenes(rob(n).dna)
  rob(n).mem(DnaLenSys) = rob(n).DnaLen
  rob(n).mem(GenesSys) = rob(n).genenum
  
  rob(n).SubSpecies = NewSubSpecies(n) ' Infection with a virus counts as a new subspecies
  logmutation n, "Infected with virus of length " + Str(vlen) + " during cycle " + Str(SimOpts.TotRunCycle) + " at pos " + Str(Insert)
  rob(n).Mutations = rob(n).Mutations + 1
  rob(n).LastMut = rob(n).LastMut + 1
getout:
End Function

Private Function IsArrayBounded(ByRef ArrayIn() As Shot) As Boolean
On Error GoTo getout
 
  IsArrayBounded = (UBound(ArrayIn) >= LBound(ArrayIn))
  Exit Function
  
getout:
  IsArrayBounded = False
  'Resume Next

End Function
