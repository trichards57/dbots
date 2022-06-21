Attribute VB_Name = "Shots_Module"
Option Explicit

Public numshots As Long       'Counter for tracking number of shots in the sim
Public ShotsThisCycle As Long ' Shots this cycle.  Only updated at end of UpdateShots()
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

Public ShotManager As New ShotManager

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
    
    a = ShotManager.createshot()
    
    If val > 32000 Then val = 32000 ' EricL March 16, 2006 This line moved here from below to catch val before assignment
    
    Dim s As Shot
    s = ShotManager.GetShot(a)
    s.Exists = True
    s.age = 0
    s.parent = n
    s.FromSpecies = rob(n).FName
    s.FromVeg = rob(n).Veg
    s.color = rob(n).color
    s.value = Int(val)
    
    If (shottype > 0) Or (shottype = -100) Then
        s.shottype = shottype
    Else
        s.shottype = -(Abs(shottype) Mod 8)  ' EricL 6/2006 essentially Mod 8 so as to increse probabiltiy that mutations do something interesting
        If s.shottype = 0 Then s.shottype = -8 ' want multiples of -8 to be -8
    End If
    If shottype = -2 Then s.color = vbWhite
    s.MemoryLocation = rob(n).mem(835)     'location for venom to target
    s.MemoryValue = rob(n).mem(836)     'value to insert into venom target location
    
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
    Dim pos As Vector
    pos = robManager.GetRobotPosition(n)
    s.Position = VectorAdd(pos, VectorScalar(angle, robManager.GetRadius(n)))
    
    'Botsareus 6/23/2016 Takes care of shot position bug - so it matches the painted robot position
    If offset Then
        s.Position = VectorSub(s.Position, robManager.GetVelocity(n))
        s.Position = VectorAdd(s.Position, robManager.GetActualVelocity(n))
    End If
    
    s.Velocity = VectorAdd(robManager.GetActualVelocity(n), VectorScalar(angle, 40))
    
    s.OldPosition = VectorSub(s.Position, s.Velocity)
    
    If rob(n).vbody > 10 Then
        s.Energy = Log(Abs(rob(n).vbody)) * 60 * rngmultiplier
        Dim temp As Long
        temp = (s.Energy + 40 + 1) \ 40 'divides and rounds up
        s.Range = temp
        s.Energy = temp * 40
    Else
        s.Range = rngmultiplier
        s.Energy = 40 * rngmultiplier
    End If
    
    'return the new shot
    newshot = a
    
    If shottype = -7 Then
        s.color = vbCyan
        s.GeneNumber = val
        s.Stored = True
        If Not copygene(s, s.GeneNumber) Then
            s.Exists = False
            s.Stored = False
            newshot = -1
        End If
    Else
        s.Stored = False
    End If
    
    If shottype = -2 Then s.Energy = val
    
    ' sperm shot
    If shottype = -8 Then
        s.dna = rob(n).dna
        s.DNALength = rob(n).DnaLen
    End If
            
    ShotManager.SetShot a, s
End Function

' creates a generic particle with arbitrary x & y, vx & vy, etc
Public Sub createshot(ByVal x As Long, ByVal y As Long, ByVal vx As Integer, ByVal vy As Integer, loc As Integer, par As Integer, val As Single, Range As Single, col As Long)
    Dim a As Long
    
    a = ShotManager.createshot()
    Dim s As Shot
    s = ShotManager.GetShot(a)
    
    s.parent = par
    s.FromSpecies = rob(par).FName
    s.FromVeg = rob(par).Veg
    
    s.Position.x = x
    s.Position.y = y
    s.Velocity.x = vx
    s.Velocity.y = vy
    s.OldPosition = VectorSub(s.Position, s.Velocity)
      
    s.age = 0
    s.color = col
    s.Exists = True
    s.Stored = False
    s.DNALength = 0
    
    Dim temp As Long
    temp = (Range + 40 + 1) \ 40 'divides and rounds up ie: range / (Robsize/3)
    
    s.Energy = Range + 40 + 1
    If val > 32000 Then val = 32000 ' Overflow protection
    If loc = -2 Then s.Energy = val
    s.Range = temp
    s.value = CInt(val)
    If loc > 0 Or loc = -100 Then
        s.shottype = loc
    Else
        s.shottype = -(Abs(loc) Mod 8)  ' EricL 6/2006 essentially Mod 8 so as to increse probabiltiy that mutations do something interesting
        If s.shottype = 0 Then s.shottype = -8 ' want multiples of -8 to be -8
    End If
    s.MemoryLocation = rob(par).mem(834) 'Botsareus 10/6/2015 Normalized on reseaving side
     
    If s.shottype = -5 Then s.MemoryValue = rob(s.parent).mem(839)
     
    ShotManager.SetShot a, s
End Sub

' calculates next shots position
Public Sub updateshots()
'moves shot then checks for collision
     'Dim a As Integer
     Dim t As Long
     Dim h As Integer
     'Dim dx As Integer
     'Dim sx As Integer
     'Dim rp As Integer
     'Dim jj As Integer
     'Dim ti As Single
     'Dim x As Long
     'Dim y As Long
     Dim tempnum As Single
     Dim s As Shot
     Dim result As UpdateResult
     
     result = ShotManager.UpdateShotsCollisions(MaxBotShotSeperation, MinBotRadius, SimOpts.Updnconnected, SimOpts.Dxsxconnected, VectorSet(SimOpts.FieldWidth, SimOpts.FieldHeight))
     
     For t = 0 To UBound(result.Collisions)
        h = result.Collisions(t).bot
        s = ShotManager.GetShot(result.Collisions(t).Shot)
            
        If Not (s.parent = rob(h).parent And rob(h).age <= 1) Then
       
            If s.Range = 0 Then
                tempnum = s.age + 1
            Else
                tempnum = s.age / s.Range
            End If
          
            'this below is horribly complicated:  allow me to explain:
            'nrg dissipates in a non-linear fashion.  Very little nrg disappears until you
            'get near the last 10% of the journey or so.
            'Don't dissipate nrg if nrg shots last forever.
            If Not SimOpts.NoShotDecay Or s.shottype <> -2 Then
                If Not (s.shottype = -4 And SimOpts.NoWShotDecay) Then 'Botsareus 9/29/2013 Do not decay waste shots
                    s.Energy = s.Energy * (Atn(tempnum * shotdecay - shotdecay)) / Atn(-shotdecay)
                End If
            End If
          
        
            If s.shottype > 0 Then
                s.shottype = (s.shottype - 1) Mod 1000 + 1 ' EricL 6/2006 Mod 1000 so as to increse probabiltiy that mutations do something interesting
            
                If s.shottype <> DelgeneSys Then
    
                    If (s.Energy / 2 > rob(h).poison) Or (rob(h).poison = 0) Then
                        rob(h).mem(s.shottype) = s.value
                    Else
                        createshot s.Position.x, s.Position.y, -s.Velocity.x, -s.Velocity.y, -5, h, s.Energy / 2, s.Range * 40, vbYellow
                        rob(h).poison = rob(h).poison - (s.Energy / 2) * 0.9
                        rob(h).Waste = rob(h).Waste + (s.Energy / 2) * 0.1
                        If rob(h).poison < 0 Then rob(h).poison = 0
                        rob(h).mem(poison) = rob(h).poison
                    End If
                End If
            Else
                Select Case s.shottype
                    Case -1: releasenrg h, s
                    Case -2: takenrg h, s
                    Case -3: takeven h, s
                    Case -4: takewaste h, s
                    Case -5: takepoison h, s
                    Case -6: releasebod h, s
                    Case -7: addgene h, s
                    Case -8: takesperm h, s ' bot hit by a sperm shot for sexual reproduction
                 End Select
            End If
            taste h, s.OldPosition.x, s.OldPosition.y, s.shottype
            s.flash = True
            
            ShotManager.SetShot result.Collisions(t).Shot, s
        End If
     Next
     
     ShotsThisCycle = result.numshots
     TotalSimEnergy(CurrentEnergyCycle) = TotalSimEnergy(CurrentEnergyCycle) + result.TotalEnergy
     
     ShotManager.UpdateShotsPosition SimOpts.NoShotDecay, SimOpts.NoWShotDecay
End Sub

Public Sub Decay(n As Integer) 'corpse decaying as waste shot, energy shot or no shot
    Dim sh As Integer
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
        Else
            va = 0
        End If
      
        If SimOpts.DecayType = 2 And va <> 0 Then
            sh = -4
            newshot n, sh, va, 1
        End If
    
        If SimOpts.DecayType = 3 And va <> 0 Then
            sh = -2
            newshot n, sh, va, 1
        End If
    
        rob(n).body = rob(n).body - SimOpts.Decay / 10
        robManager.SetRadius n, FindRadius(n)
    End If
End Sub

Public Sub defacate(n As Integer) 'only used to get rid of massive amounts of waste
    Dim sh As Integer
    Dim va As Single
    sh = -4
    va = 200
    
    If va > rob(n).Waste Then va = rob(n).Waste
    If rob(n).Waste > 32000 Then rob(n).Waste = 31500: va = 500
    
    rob(n).Waste = rob(n).Waste - va
    rob(n).nrg = rob(n).nrg - (SimOpts.Costs(SHOTCOST) * SimOpts.Costs(COSTMULTIPLIER)) / (IIf(rob(n).numties < 0, 0, rob(n).numties) + 1)
    newshot n, sh, va, 1, True
    rob(n).Pwaste = rob(n).Pwaste + va / 1000
End Sub

' robot n, hit by shot t, releases energy
Public Sub releasenrg(ByVal n As Integer, ByRef s As Shot)
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
      
    If rob(n).nrg <= 0.5 Then Exit Sub
    
    vel = VectorSub(robManager.GetActualVelocity(n), s.Velocity) 'negative relative velocity of shot hitting bot 'Botsareus 6/22/2016 Now based on robots actual velocity
                                                               'the shot to the hit bot
    vel = VectorAdd(vel, VectorScalar(robManager.GetActualVelocity(n), 0.5)) 'then add in half the velocity of hit robot
    
    If SimOpts.EnergyExType Then
        If s.Range = 0 Then ' Divide by zero protection
            power = 0#
        Else
            power = CSng(s.value) * s.Energy / CSng((s.Range * (RobSize / 3))) * SimOpts.EnergyProp
        End If
        If s.Energy < 0 Then Exit Sub
    Else
        power = SimOpts.EnergyFix
    End If
        
    If rob(n).Corpse Then power = power * 0.5 'half power against corpses.  Most of your shot is wasted 'The only way I can see this happening is if something tie injected energy into corpse
    
    Range = s.Range * 2 'returned energy shots have twice the range as the shot that it came from (but half the velocity)
    
    If rob(n).poison > power Then 'create poison shot
        createshot s.Position.x, s.Position.y, vel.x, vel.y, -5, n, power, Range * (RobSize / 3), vbYellow
        rob(n).poison = rob(n).poison - (power * 0.9)
        If rob(n).poison < 0 Then rob(n).poison = 0
        rob(n).mem(poison) = rob(n).poison
    Else ' create energy shot
         
        EnergyLost = power * 0.9
        If EnergyLost > rob(n).nrg Then
            power = rob(n).nrg
            rob(n).nrg = 0
        Else
            rob(n).nrg = rob(n).nrg - EnergyLost  'some of shot comes from nrg
        End If
    
        EnergyLost = power * 0.01
        If EnergyLost > rob(n).body Then
            rob(n).body = 0
        Else
            rob(n).body = rob(n).body - EnergyLost 'some of shot comes from body
        End If
      
        ' pass local vars to createshot so that no Shots array elements are on the stack in case the Shots array gets redimmed
        x = s.Position.x
        y = s.Position.y
      
        createshot x, y, vel.x, vel.y, -2, n, power, Range * (RobSize / 3), vbWhite
        robManager.SetRadius n, FindRadius(n)
    End If
    
    If rob(n).body <= 0.5 Or rob(n).nrg <= 0.5 Then
        rob(n).Dead = True
        rob(s.parent).Kills = rob(s.parent).Kills + 1
        rob(s.parent).mem(220) = rob(s.parent).Kills
    End If
End Sub
Private Sub releasebod(ByVal n As Integer, ByRef s As Shot) 'a robot is shot by a -6 shot and releases energy directly from body points
    'much more effective against a corpse
    Dim vel As Vector
    Dim Range As Single
    Dim power As Single
    Dim shell As Single
    Dim EnergyLost As Single
    
    If rob(n).body <= 0 Then Exit Sub
    
    vel = VectorSub(robManager.GetActualVelocity(n), s.Velocity) 'negative relative velocity of shot hitting bot 'Botsareus 6/22/2016 Now based on robots actual velocity
                                                   'the shot to the hit bot
    vel = VectorAdd(vel, VectorScalar(robManager.GetActualVelocity(n), 0.5)) 'then add in half the velocity of hit robot
    
    If SimOpts.EnergyExType Then
        If s.Range = 0 Then ' Divide by zero protection
            power = 0
        Else
            power = CSng(s.value) * s.Energy / CSng((s.Range * (RobSize / 3))) * SimOpts.EnergyProp
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
        Exit Sub
    Else
        rob(n).shell = rob(n).shell - power / ShellEffectiveness
        If rob(n).shell < 0 Then rob(n).shell = 0
        rob(n).mem(823) = rob(n).shell
        power = power - shell
    End If
    
    If power <= 0 Then Exit Sub
    
    Range = s.Range * 2   'new range formula based on range of incoming shot
    
    ' create energy shot
    If rob(n).Corpse = True Then
        power = power * 4 'So effective against corpses it makes me siiiiiinnnnnggg
        If power > rob(n).body * 10 Then power = rob(n).body * 10
        rob(n).body = rob(n).body - power / 10      'all energy comes from body
        robManager.SetRadius n, FindRadius(n)
    Else
        Dim leftover As Single
        
        leftover = 0
        EnergyLost = power * 0.2
        If EnergyLost > rob(n).nrg Then
            leftover = EnergyLost - rob(n).nrg
            rob(n).nrg = 0
        Else
            rob(n).nrg = rob(n).nrg - EnergyLost  'some of shot comes from nrg
        End If
        
        EnergyLost = power * 0.08
        If EnergyLost > rob(n).body Then
            leftover = leftover + EnergyLost - rob(n).body * 10
            rob(n).body = 0
        Else
            rob(n).body = rob(n).body - EnergyLost 'some of shot comes from body
        End If
        
        With rob(n)
            If leftover > 0 Then
                If .nrg > 0 And .nrg > leftover Then
                    .nrg = .nrg - leftover
                    leftover = 0
                ElseIf .nrg > 0 And .nrg < leftover Then
                    leftover = leftover - .nrg
                    .nrg = 0
                End If
            
                If .body > 0 And .body * 10 > leftover Then
                    .body = .body - leftover * 0.1
                    leftover = 0
                ElseIf rob(n).body > 0 And rob(n).body * 10 < leftover Then
                    .body = 0
                End If
            End If
        End With
        robManager.SetRadius n, FindRadius(n)
    End If
    
    If rob(n).body <= 0.5 Or rob(n).nrg <= 0.5 Then
        rob(n).Dead = True
        rob(s.parent).Kills = rob(s.parent).Kills + 1
        rob(s.parent).mem(220) = rob(s.parent).Kills
    End If
    
    createshot s.Position.x, s.Position.y, vel.x, vel.y, -2, n, power, Range * (RobSize / 3), vbWhite
End Sub

' robot n takes the energy carried by shot t
Private Sub takenrg(ByVal n As Integer, ByRef s As Shot)
    Dim partial As Single
    Dim overflow As Single
                
    If rob(n).Corpse Then Exit Sub
    
    If s.Range < 0.00001 Then
        partial = 0
    Else
        partial = s.Energy
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
    
    robManager.SetRadius n, FindRadius(n)
End Sub
'  robot takes a venomous shot and becomes seriously messed up
Private Sub takeven(ByVal n As Integer, ByRef s As Shot)
    Dim power As Single
    Dim temp As Single
      
    If rob(n).Corpse Then Exit Sub
    
    power = CSng(s.Energy / CSng((s.Range * (RobSize / 3))) * s.value)
  
    If power < 1 Then Exit Sub
  
    If s.FromSpecies = rob(n).FName Then   'Robot is immune to venom from his own species
        rob(n).venom = rob(n).venom + power   'Robot absorbs venom fired by conspec
    
        If ((rob(n).venom) > 32000) Then rob(n).venom = 32000
    
        rob(n).mem(825) = rob(n).venom
    Else
        power = power * VenumEffectivenessVSShell  'Botsareus 3/6/2013 max power for venum is capped at 100 so I multiply to get an average
        If power < rob(n).shell * ShellEffectiveness Then
            rob(n).shell = rob(n).shell - power / ShellEffectiveness
            rob(n).mem(823) = rob(n).shell
            Exit Sub
        Else
            temp = power
            power = power - rob(n).shell * ShellEffectiveness
            rob(n).shell = rob(n).shell - temp / ShellEffectiveness
            If rob(n).shell < 0 Then rob(n).shell = 0
            rob(n).mem(823) = rob(n).shell
        End If
        power = power / VenumEffectivenessVSShell
    
        If power < 1 Then Exit Sub
    
        rob(n).Paralyzed = True
    
        If ((rob(n).Paracount + power) > 32000) Then
            rob(n).Paracount = 32000
        Else
            rob(n).Paracount = rob(n).Paracount + power
        End If
        
        If s.MemoryLocation > 0 Then
            rob(n).Vloc = (s.MemoryLocation - 1) Mod 1000 + 1
            If rob(n).Vloc = 340 Then rob(n).Vloc = 0 'protection from delgene attacks
        Else
            Do
                rob(n).Vloc = Random(1, 1000)
            Loop Until rob(n).Vloc <> 340
        End If
    
        rob(n).Vval = s.MemoryValue
    End If
End Sub

'  Robot n takes shot t and adds its value to his waste reservoir
Private Sub takewaste(ByVal n As Integer, ByRef s As Shot)
    Dim power As Single
  
    power = s.Energy / (s.Range * (RobSize / 3)) * s.value
    If power >= 0 Then rob(n).Waste = rob(n).Waste + power
End Sub
' Robot receives poison shot and becomes disorientated
Private Sub takepoison(ByVal n As Integer, ByRef s As Shot)
    Dim power As Single

    If rob(n).Corpse Then Exit Sub
    
    power = CSng(s.Energy / CSng((s.Range * (RobSize / 3))) * s.value)
    
    If power < 1 Then Exit Sub
    
    If s.FromSpecies = rob(n).FName Then    'Robot is immune to poison from his own species
        rob(n).poison = rob(n).poison + power 'Robot absorbs poison fired by conspecs
        If rob(n).poison > 32000 Then rob(n).poison = 32000
        rob(n).mem(827) = rob(n).poison
    Else
        rob(n).Poisoned = True
        rob(n).Poisoncount = rob(n).Poisoncount + power / 1.5
        If rob(n).Poisoncount > 32000 Then rob(n).Poisoncount = 32000
        If s.MemoryLocation > 0 Then 'Botsareus 10/6/2015 Minor bug fixing
            rob(n).Ploc = (s.MemoryLocation - 1) Mod 1000 + 1
            If rob(n).Ploc = 340 Then rob(n).Ploc = 0 'protection from delgene attacks Botsareus 10/7/2015 Moved here after mod
        Else
            Do
                rob(n).Ploc = Random(1, 1000)
            Loop Until rob(n).Ploc <> 340
        End If
        rob(n).Pval = s.MemoryValue
    End If
End Sub

'Robot is hit by sperm shot and becomes fertilized for potential sexual reproduction
Private Sub takesperm(ByVal n As Integer, ByRef s As Shot)
    If rob(n).fertilized < -10 Then Exit Sub 'block sex repro when necessary

    Dim x As Integer

    If s.DNALength = 0 Then Exit Sub
    rob(n).fertilized = 10                      ' bots stay fertilized for 10 cycles currently
    rob(n).mem(SYSFERTILIZED) = 10
    ReDim rob(n).spermDNA(s.DNALength)      ' copy the male's DNA to the female's bot structure
    rob(n).spermDNA = s.dna
    rob(n).spermDNAlen = s.DNALength
End Sub

'Botsareus 10/5/2015 Bug fix for negative values in virus
Public Sub Vshoot(n As Integer, ByRef s As Shot)
'here we shoot a virus
    Dim tempa As Single
    Dim ShAngle As Single

    If Not s.Exists Or Not s.Stored Then Exit Sub
  
    If rob(n).mem(VshootSys) < 0 Then rob(n).mem(VshootSys) = 1
  
    tempa = CSng(rob(n).mem(VshootSys)) * 20# '.vshoot * 20
    If tempa > 32000 Then tempa = 32000
    If tempa < 0 Then tempa = 0
    
    s.Energy = tempa
    rob(n).nrg = rob(n).nrg - (tempa / 20#) - (SimOpts.Costs(SHOTCOST) * SimOpts.Costs(COSTMULTIPLIER))
  
    s.Range = 11 + CInt((rob(n).mem(VshootSys)) / 2)
    rob(n).nrg = rob(n).nrg - CSng(rob(n).mem(VshootSys)) - (SimOpts.Costs(SHOTCOST) * SimOpts.Costs(COSTMULTIPLIER))
    
    With s
        ShAngle = Random(1, 1256) / 200
        .Stored = False
        .Position.x = (robManager.GetRobotPosition(n).x + Cos(ShAngle) * robManager.GetRadius(n))
        .Position.y = (robManager.GetRobotPosition(n).y - Sin(ShAngle) * robManager.GetRadius(n))
  
        .Velocity.x = absx(ShAngle, RobSize / 3, 0, 0, 0) ' set shot speed x seems to not work well at high bot speeds
        .Velocity.y = absy(ShAngle, RobSize / 3, 0, 0, 0) ' set shot speed y
  
        .Velocity.x = .Velocity.x + robManager.GetActualVelocity(n).x
        .Velocity.y = .Velocity.y + robManager.GetActualVelocity(n).y
    
        .OldPosition.x = .Position.x - .Velocity.x
        .OldPosition.y = .Position.y - .Velocity.y
    End With
End Sub

Public Function MakeVirus(robn As Integer, ByVal gene As Integer) As Boolean
    rob(robn).virusshot = newshot(robn, -7, Int(gene), 1)
    MakeVirus = rob(robn).virusshot > 0
End Function

' copy gene number p from robot that fired shot n into shot n dna (virus)
Public Function copygene(ByRef s As Shot, ByVal p As Integer) As Boolean
    Dim t As Integer
    Dim parent As Integer
    Dim genelen As Integer
    Dim GeneStart As Long
    Dim GeneEnding As Integer
    
    parent = s.parent
    
    If (p > rob(parent).genenum) Or p < 1 Then
      ' target gene is beyond the DNA bounds
        copygene = False
        Exit Function
    End If
    
    GeneStart = genepos(rob(parent).dna, p)
    GeneEnding = GeneEnd(rob(parent).dna, GeneStart)
    genelen = GeneEnding - GeneStart + 1
    
    If genelen < 1 Then
        copygene = False
        Exit Function
    End If
    
    ReDim s.dna(genelen)
        
    For t = 0 To genelen - 1
        s.dna(t) = rob(parent).dna(GeneStart + t)
    Next t
    
    s.DNALength = genelen
    
    copygene = True
End Function

' adds gene from shot p to robot n dna
Public Function addgene(ByVal n As Integer, ByRef s As Shot) As Integer
    Dim t As Long
    Dim Insert As Long
    Dim vlen As Long   'length of the DNA code of the virus
    Dim Position As Integer   'gene position to insert the virus
    Dim power As Single
    
  'Dead bodies and virus immune bots can't catch a virus
    If rob(n).Corpse Or (rob(n).VirusImmune) Then Exit Function
  
    power = s.Energy / (s.Range * RobSize / 3) * s.value
  
    If power < rob(n).Slime * SlimeEffectiveness Then
        rob(n).Slime = rob(n).Slime - power / SlimeEffectiveness
        Exit Function
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
        If Insert = (rob(n).DnaLen) Then Insert = rob(n).DnaLen
    End If
  
    vlen = s.DNALength
  
    If MakeSpace(rob(n).dna, Insert, vlen) Then 'Moves genes back to make space
        For t = Insert To Insert + vlen - 1
            rob(n).dna(t + 1) = s.dna(t - Insert)
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
End Function

