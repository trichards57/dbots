Attribute VB_Name = "Ties"
Option Explicit

' tie structure, used to represent robot ties
'booleans = integers in space taken up.
Public Type tie
  Port As Integer         ' the tie port, i.e. the value one must give to .tienum var to access the tie
  pnt As Integer          ' the robot the tie points to
  ptt As Integer          ' the back tie of the pointed robot
  ang As Single           ' current tie angle (relative to aim)
  bend As Single          ' angle bend value
  angreg As Boolean       ' the angle is fixed?
  ln As Long              ' tie len
  shrink As Long          ' tie shrink value
  stat As Boolean         ' apparently unused
  last As Integer         ' tie timing: <-1: cycles to angle fixing; >1: cycles to tie destruction
  mem As Integer          ' Reads nrg of tied robot
  back As Boolean         ' is it a back tie?
  nrgused As Boolean      ' for colouring the tie red in case of energy transfer
  infused As Boolean      ' for colouring white in case of information transfer
  sharing As Boolean       ' for coloring yellow in case of sharing
  
  'New:
  'Spring Force = -k*displacement - velocty * b
  'b and k are constants between 0 and 1
  
  Fx As Single 'for storing what forces have been applied to us.  Erase when Finished!
  Fy As Single
  
  NaturalLength As Single
  k As Single
  b As Single
  type As Byte
    '0 = damped spring, lower b and k values, birth ties here
    '1 = string (only applies force if longer than
      '"natural length", not if too short) b and k values high
    '2 = bone (very high b and k values).  (Or perhaps something better?)
    '3 = anti-rope - only applies force if shorter than
      '"natural length", not if too long) b and k values high
End Type

Public Const MAXTIES As Integer = 10


'This routine deals with information transfer only. Added in to fix a major bug
'in which older robots could transfer information to younger bots OK but
'young bots could not transfer information to older bots in time for the information
'to do any good
Public Sub tieportcom(t As Integer)
  Dim k As Integer
  Dim tn As Integer
  'Dim tp As Integer
  'tp = tieport1
  
  With rob(t)
    If Not (.mem(455) <> 0 And .numties > 0 And .mem(tieloc) > 0) Then GoTo getout
    tn = .mem(TIENUM)
    k = 1
      If .mem(tieloc) > 0 And .mem(tieloc) < 1001 Then  '.tieloc value
        While .Ties(k).pnt > 0
          If .Ties(k).Port = tn Then
            rob(.Ties(k).pnt).mem(.mem(tieloc)) = .mem(tieval) 'stores a value in tied robot memory location (.tieloc) specified in .tieval
            If Not .Ties(k).back Then   'forward information transfer
              .Ties(k).infused = True   'draws tie white
            Else                        'backward information transfer
              rob(.Ties(k).pnt).Ties(.Ties(k).ptt).infused = True 'draws tie white
            End If
           ' .mem(tienum) = 0 ' EricL 4/24/2006 Commented out
            .mem(tieval) = 0
            .mem(tieloc) = 0
          End If
          k = k + 1
        Wend
      End If
getout:
  End With
End Sub

Public Sub UpdateTieAngles(t As Integer)
  Dim k As Integer
  Dim n As Integer
  Dim tieAngle As Single
  Dim dist As Single
  Dim whichTie As Integer
    
  ' Zero these out incase no ties or tienum is non-zero, but does not refer to a tieport, etc.
  rob(t).mem(TIEANG) = 0
  rob(t).mem(TIELEN) = 0
  
  'No point in setting the length and angle if no ties!
  If rob(t).numties <= 0 Then GoTo getout:
    
  'Figure out if .tienum has a value.  If it's zero, use .tiepres
  If rob(t).mem(TIENUM) <> 0 Then
    whichTie = rob(t).mem(TIENUM)
  Else
    whichTie = rob(t).mem(TIEPRES)
  End If
  
  If whichTie = 0 Then Exit Sub
  
  'Now find the tie that corrosponds to either .tienum or .tiepres and set .tieang and .tielen accordingly
  'We count down through the ties to find the most recent tie with the specified tieport since more than one tie
  'can potentially have the same tieport and we want the most recent, which will be the one with the highest k.
  k = rob(t).numties
  While k > 0
    If rob(t).Ties(k).Port = whichTie Then
       n = rob(t).Ties(k).pnt  ' The bot number of the robot on the other end of the tie
       tieAngle = angle(robManager.GetRobotPosition(t).x, robManager.GetRobotPosition(t).y, robManager.GetRobotPosition(n).x, robManager.GetRobotPosition(n).y)
       dist = Sqr((robManager.GetRobotPosition(t).x - robManager.GetRobotPosition(n).x) ^ 2 + (robManager.GetRobotPosition(t).y - robManager.GetRobotPosition(n).y) ^ 2)
       'Overflow prevention.  Very long ties can happen for one cycle when bots wrap in torridal fields
       If dist > 32000 Then dist = 32000
       rob(t).mem(TIEANG) = -CInt(AngDiff(angnorm(tieAngle), angnorm(rob(t).aim)) * 200)
       rob(t).mem(TIELEN) = CInt(dist - rob(t).radius - rob(n).radius)
       GoTo getout
    End If
    k = k - 1
  Wend
getout:
End Sub

' this procedure takes care of parsing and addressing
' ties commands: bend, shrink, communications
Public Sub Update_Ties(t As Integer)
  Dim tp As Integer
  Dim tn As Integer
  Dim k As Integer
  Dim l As Single
  Dim ptag As Single
  Dim length As Long

  With rob(t)
    tp = tieport1
    tn = .mem(TIENUM)   '.tienum value
    
    'this routine addresses all ties. not just ones that match .tienum
    k = 1
    .vbody = .body
    Dim atleast1tie As Boolean
    atleast1tie = False
    While .Ties(k).pnt > 0 'while there is a tie that points to another robot that this bot created.
      If .Multibot Then
        If Not .Ties(k).back Then
          If .mem(830) > 0 Then
            sharenrg t, k
            .Ties(k).sharing = True   'yellow ties
          End If
          If .mem(831) > 0 Then
            sharewaste t, k
            .Ties(k).sharing = True   'yellow ties
          End If
          If .mem(832) > 0 Then
            shareshell t, k
            .Ties(k).sharing = True   'yellow ties
          End If
          If .mem(833) > 0 Then
            shareslime t, k
            .Ties(k).sharing = True    'yellow ties
          End If
          If .mem(sharechlr) > 0 And .Chlr_Share_Delay = 0 And Not rob(t).NoChlr Then    'Panda 8/31/2013 code to share chloroplasts 'Botsareus 8/16/2014 chloroplast sharing disable
            sharechloroplasts t, k
            .Ties(k).sharing = True   'yellow ties
          End If
        End If
        .vbody = .vbody + rob(.Ties(k).pnt).body
        If .FName = rob(.Ties(k).pnt).FName Then atleast1tie = True
      End If
      k = k + 1
    Wend
    
    ' Zero out the sharing sysvars
    .mem(830) = 0
    .mem(831) = 0
    .mem(832) = 0
    .mem(833) = 0
    .mem(sharechlr) = 0
        
    .numties = k - 1  ' Set the number of ties.
    .mem(numties) = .numties   'places a value in the memory cell .numties for number of ties attached to a robot
    
    If .numties = 0 Then
      .Multibot = False
      .mem(multi) = 0
      GoTo getout:
    End If
           
    ' Handle the deltie sysvar.  Bot is trying to delete one or more ties
    If .mem(DELTIE) <> 0 Then
      For k = 1 To .numties
        If .Ties(k).pnt > 0 And .Ties(k).Port = .mem(tp + 17) Then DeleteTie t, .Ties(k).pnt
      Next k
      .mem(DELTIE) = 0 'resets .deltie command
    End If
    
    If tn = 0 Then tn = .mem(TIEPRES)
    If tn = 0 Then GoTo getout
   ' If tn Then  'routines only carried out if .tienum has a value
      
      k = 1
      While .Ties(k).pnt > 0 And k < MAXTIES
        If .Multibot And .Ties(k).type = 3 Then ' Has to be a multibot and tie has to have hardened
         
          'FixAng - fixes tie angle
          'Positive values fix the tie angle
          'Negative values allow the tie to pivot freely
          If .mem(FIXANG) <> 32000 And .Ties(k).Port = tn Then
            
            If .mem(FIXANG) >= 0 Then
              .Ties(k).ang = (.mem(FIXANG) Mod 1256) / 200
              .Ties(k).angreg = True 'EricL 4/24/2006
              'If .mem(FIXANG) > 628 Then .mem(FIXANG) = -627
              'If .mem(FIXANG) < -628 Then .mem(FIXANG) = 627
            Else
              .Ties(k).angreg = False 'EricL 4/24/2006
            End If
          End If
                    
          'TieLen Section
          If .mem(FIXLEN) <> 0 And .Ties(k).Port = tn Then 'fixes tie length
           'length = Abs(.mem(FIXLEN))
            length = Abs(.mem(FIXLEN)) + .radius + rob(.Ties(k).pnt).radius ' include the radius of the tied bots in the length
            If length > 32000 Then length = 32000 ' Can happen for very big bots with very long ties.
            .Ties(k).NaturalLength = CInt(length) 'for first robot
            rob(.Ties(k).pnt).Ties(srctie((.Ties(k).pnt), t)).NaturalLength = CInt(length) 'for second robot. What a messed up formula
          End If
    
          'EricL 5/7/2006 Added Stifftie section.  This never made it into the 2.4 code
          If .mem(stifftie) <> 0 And .Ties(k).Port = tn Then
            .mem(stifftie) = .mem(stifftie) Mod 100
            If .mem(stifftie) = 0 Then .mem(stifftie) = 100
            If .mem(stifftie) < 0 Then .mem(stifftie) = 1
            .Ties(k).b = 0.005 * .mem(stifftie) ' May need to tweak the multiplier here vares from 0.0025 to .1
            rob(.Ties(k).pnt).Ties(srctie((.Ties(k).pnt), t)).b = 0.005 * .mem(stifftie) ' May need to tweak the multiplier here
            .Ties(k).k = 0.0025 * .mem(stifftie) 'May need to tweak the multiplier here:  varies from 0.00125 to 0.05
            rob(.Ties(k).pnt).Ties(srctie((.Ties(k).pnt), t)).k = 0.0025 * .mem(stifftie) ' May need to tweak the multiplier here: varies from 0.00125 to 0.05
          End If
        End If
        k = k + 1
      Wend
      
      .mem(FIXANG) = 32000
      .mem(FIXLEN) = 0
      .mem(stifftie) = 0
      k = 1
      
      
'Botsareus 3/22/2013 Complete fix for tielen...tieang 1...4
      If .Multibot Then 'check for multibot first
      
       For k = 1 To 4
       If .Ties(k).pnt > 0 And .Ties(k).type = 3 Then
        'input
        If .TieLenOverwrite(k - 1) Then
         length = .mem(483 + k) + .radius + rob(.Ties(k).pnt).radius ' include the radius of the tied bots in the length
         If length > 32000 Then length = 32000 ' Can happen for very big bots with very long ties.
         .Ties(k).NaturalLength = CInt(length) 'for first robot
         rob(.Ties(k).pnt).Ties(srctie((.Ties(k).pnt), t)).NaturalLength = CInt(length) 'for second robot. What a messed up formula
        End If
        If .TieAngOverwrite(k - 1) Then
         .Ties(k).ang = angnorm(.mem(479 + k) / 200)
         .Ties(k).angreg = True 'EricL 4/24/2006
        End If
        'clear input
        .TieAngOverwrite(k - 1) = False
        .TieLenOverwrite(k - 1) = False
        'output
        Dim n As Integer
        Dim dist As Single
        Dim tieAngle As Single
        n = .Ties(k).pnt
        tieAngle = angle(robManager.GetRobotPosition(t).x, robManager.GetRobotPosition(t).y, robManager.GetRobotPosition(n).x, robManager.GetRobotPosition(n).y)
        dist = Sqr((robManager.GetRobotPosition(t).x - robManager.GetRobotPosition(n).x) ^ 2 + (robManager.GetRobotPosition(t).y - robManager.GetRobotPosition(n).y) ^ 2)
        If dist > 32000 Then dist = 32000 'Botsareus 1/24/2014 Bug fix here
        .mem(483 + k) = CInt(dist - .radius - rob(n).radius)
        .mem(479 + k) = angnorm(angnorm(tieAngle) - angnorm(.aim)) * 200
       End If
       Next
      
      End If
      
      
      k = 1
          
'      If .mem(tp) Then  '.tieang value
'        k = 1
'        While .Ties(k).pnt > 0
'          If .Ties(k).Port = tn Then bend t, k, .mem(tp) 'bend a tie
'          k = k + 1
'        Wend
'      End If
'
'      If .mem(tp + 1) Then  '.tielen value
'        k = 1
'        While .Ties(k).pnt > 0
'          If .Ties(k).Port = tn Then shrink t, k, .mem(FIXLEN) 'set tie length to value specified in mem location 451 (tp+1)
'          k = k + 1
'        Wend
'      End If
      
     
      
        'Botsareus 7/22/2015 Code more coherent
        If .mem(tp + 2) < 0 Then  'we are checking values that are negative such as -1 or -6
        
            If .mem(tp + 2) = -1 And .mem(tp + 3) <> 0 Then   'tieloc = -1 and .tieval not zero
                  
                l = .mem(tp + 3) ' l is amount of energy to exchange, positive to give nrg away, negative to take it
                                
                'Limits on Tie feeding as a function of body attempting to do the feeding (or sharing)
                If .body < 0 Then l = 0                        ' If your body has gone negative, you can't take or give nrg.
                If .nrg < 0 Then l = 0                         ' If you nrg has gone negative, you can't take or give nrg.
                If .age = 0 Then l = 0                         ' The just born can't trasnfer nrg
                If l > 1000 Then l = 1000                      ' Upper limt on sharing
                If l < -3000 Then l = -3000                    ' Upper limit on tie feeding
                  
                k = 1
                While .Ties(k).pnt > 0        'tie actually points at something
                    If .Ties(k).Port = tn Then  'tienum indicates this tie
                       
                        'Giving nrg away
                        If l > 0 Then
                        
                            If l > .nrg Then l = .nrg ' Can't give away more nrg than you have
                        
                            rob(.Ties(k).pnt).nrg = rob(.Ties(k).pnt).nrg + l * 0.7               'tied robot receives energy
                            If rob(.Ties(k).pnt).nrg > 32000 Then rob(.Ties(k).pnt).nrg = 32000
                            rob(.Ties(k).pnt).body = rob(.Ties(k).pnt).body + l * 0.029           'tied robot stores some fat
                            If rob(.Ties(k).pnt).body > 32000 Then rob(.Ties(k).pnt).body = 32000
                            rob(.Ties(k).pnt).Waste = rob(.Ties(k).pnt).Waste + l * 0.01          'tied robot receives waste
                            rob(.Ties(k).pnt).radius = FindRadius(.Ties(k).pnt)
                            
                            .nrg = .nrg - l                                                       'tying robot gives up energy
                        End If
                        
                        'Taking nrg
                        If l < 0 Then
                                            
                            If l < -rob(.Ties(k).pnt).nrg Then l = -rob(.Ties(k).pnt).nrg ' Can't give away more nrg than you have
                                           
                            'Poison
                            ptag = Abs(l / 4)
                            If rob(.Ties(k).pnt).poison > ptag Then 'target robot has poison
                                If rob(.Ties(k).pnt).FName <> .FName Then 'can't poison your brother
                                
                                    .Poisoned = True
                                    .Poisoncount = .Poisoncount + ptag
                                    If .Poisoncount > 32000 Then .Poisoncount = 32000
                                    
                                    l = 0
                                    
                                    rob(.Ties(k).pnt).poison = rob(.Ties(k).pnt).poison - ptag
                                    rob(.Ties(k).pnt).mem(827) = rob(.Ties(k).pnt).poison
                                    
                                    If rob(.Ties(k).pnt).mem(834) > 0 Then
                                     .Ploc = ((rob(.Ties(k).pnt).mem(834) - 1) Mod 1000) + 1  'sets .Ploc to targets .mem(ploc) EricL - 3/29/2006 Added Mod to fix overflow
                                     If .Ploc = 340 Then .Ploc = 0
                                    Else
                                     Do
                                      .Ploc = Random(1, 1000)
                                     Loop Until .Ploc <> 340
                                    End If
        
                                    .Pval = rob(.Ties(k).pnt).mem(839)
        
                                End If
                            End If
                                                
                            .nrg = .nrg - l * 0.7                'tying robot receives energy
                            If .nrg > 32000 Then .nrg = 32000
                            .body = .body - l * 0.029            'tying robot stores some fat
                            If .body > 32000 Then .body = 32000
                            .Waste = .Waste - l * 0.01      'tying robot adds waste
                            .radius = FindRadius(t)
                            
                            rob(.Ties(k).pnt).nrg = rob(.Ties(k).pnt).nrg + l 'Take the nrg
                            
                            If rob(.Ties(k).pnt).nrg <= 0 And rob(.Ties(k).pnt).Dead = False Then 'Botsareus 3/11/2014 Tie feeding kills
                                If Not rob(.Ties(k).pnt).Corpse Then 'Botsareus 7/17/2016 Bug fix to prevent logging infinate kills against a corpse
                                    .Kills = .Kills + 1
                                    If .Kills > 32000 Then .Kills = 32000
                                    .mem(220) = .Kills
                                End If
                            End If
                        End If
                        
                        If Not .Ties(k).back Then   'forward ties
                            .Ties(k).nrgused = True   'red ties
                        Else                        'backward ties
                            rob(.Ties(k).pnt).Ties(.Ties(k).ptt).nrgused = True 'red ties
                        End If
                        
                    End If
                        
                k = k + 1
                Wend
            End If
                
            If .mem(tp + 2) = -3 And .mem(tp + 3) <> 0 Then  'inject or steal venom
            
                l = .mem(tp + 3) 'amount of venom to take or inject
            
                'limits on injecting or taking venum
                If l > 100 Then l = 100
                If l < -100 Then l = -100
                  
                k = 1
                While .Ties(k).pnt > 0
                    If .Ties(k).Port = tn Then
                    
                        If l > .venom Then l = .venom
                        
                        If l > 0 Then 'works the same as a venom injection
                                                
                            rob(.Ties(k).pnt).Paracount = rob(.Ties(k).pnt).Paracount + l 'paralysis counter set
                            If rob(.Ties(k).pnt).Paracount > 32000 Then rob(.Ties(k).pnt).Paracount = 32000
                            rob(.Ties(k).pnt).Paralyzed = True         'robot paralyzed
                            
                            If .mem(835) > 0 Then
                             rob(.Ties(k).pnt).Vloc = ((.mem(835) - 1) Mod 1000) + 1
                             If rob(.Ties(k).pnt).Vloc = 340 Then rob(.Ties(k).pnt).Vloc = 0
                            Else
                             Do
                              rob(.Ties(k).pnt).Vloc = Random(1, 1000)
                             Loop Until rob(.Ties(k).pnt).Vloc <> 340
                            End If
                            
                            rob(.Ties(k).pnt).Vval = .mem(836)
                            .venom = .venom - l
                            .mem(825) = .venom
                        
                        End If
                      
                        If l < 0 Then 'Taking venom
                    
                            If l < -rob(.Ties(k).pnt).venom Then l = -rob(.Ties(k).pnt).venom ' Can't give away more venom than you have
                            
                            'robot steals venom from tied target
                            rob(.Ties(k).pnt).venom = rob(.Ties(k).pnt).venom + l
                            .venom = .venom - l
                            If .venom > 32000 Then .venom = 32000
                            .mem(825) = .venom
                        
                        End If
                      
                        If Not .Ties(k).back Then   'forward ties
                            .Ties(k).nrgused = True   'red ties
                        Else                        'backward ties
                            rob(.Ties(k).pnt).Ties(.Ties(k).ptt).nrgused = True 'red ties
                        End If
                      
                    End If
                    
                k = k + 1
                Wend
            End If
            
            If .mem(tp + 2) = -4 And .mem(tp + 3) <> 0 Then 'trade waste via ties
            
                l = .mem(tp + 3) ' l is amount of waste to exchange, positive to give waste away, negative to take it
            
                'limits on giving or taking waste
                If l > 1000 Then l = 1000
                If l < -1000 Then l = -1000
                
                k = 1
                While .Ties(k).pnt > 0
                    If .Ties(k).Port = tn Then
                    
                        'giving waste away
                        If l > 0 Then
                        
                            If l > .Waste Then l = .Waste
                            
                            rob(.Ties(k).pnt).Waste = rob(.Ties(k).pnt).Waste + l * 0.99
                            .Waste = .Waste - l
                            .Pwaste = .Pwaste + l * 0.01 'some waste is converted into perminent waste rather then given away
                        
                        End If
                        
                        'taking waste
                        If l < 0 Then
                        
                            If l < -rob(.Ties(k).pnt).Waste Then l = -rob(.Ties(k).pnt).Waste
                        
                            .Waste = .Waste - l * 0.99 'robot reseaves waste from opponent
                            rob(.Ties(k).pnt).Waste = rob(.Ties(k).pnt).Waste + l 'opponent losing some waste
                            rob(.Ties(k).pnt).Pwaste = rob(.Ties(k).pnt).Pwaste - l * 0.01 'some waste is converted into perminent waste rather then given away
            
                        End If
                        
                        If Not .Ties(k).back Then   'forward ties
                            .Ties(k).nrgused = True   'red ties
                        Else                        'backward ties
                            rob(.Ties(k).pnt).Ties(.Ties(k).ptt).nrgused = True 'red ties
                        End If
            
                    End If
            
                k = k + 1
                Wend
            End If
                
            If .mem(tp + 2) = -6 And .mem(tp + 3) <> 0 Then   'tieloc = -6 and .tieval not zero
                  
                l = .mem(tp + 3) ' l is amount of body to exchange, positive to give body away, negative to take it
                                
                'Limits on Tie feeding as a function of body attempting to do the feeding (or sharing)
                If .body < 0 Then l = 0                        ' If your body has gone negative, you can't take or give body.
                If .nrg < 0 Then l = 0                         ' If you nrg has gone negative, you can't take or give body.
                If .age = 0 Then l = 0                         ' The just born can't trasnfer body
                If l > 100 Then l = 100                      ' Upper limt on sharing
                If l < -300 Then l = -300                    ' Upper limit on tie feeding
                  
                k = 1
                While .Ties(k).pnt > 0        'tie actually points at something
                    If .Ties(k).Port = tn Then  'tienum indicates this tie
                       
                        'Giving body away
                        If l > 0 Then
                        
                            If l > .body Then l = .body ' Can't give away more body than you have
                        
                            rob(.Ties(k).pnt).nrg = rob(.Ties(k).pnt).nrg + l * 0.03              'tied robot receives energy
                            If rob(.Ties(k).pnt).nrg > 32000 Then rob(.Ties(k).pnt).nrg = 32000
                            rob(.Ties(k).pnt).body = rob(.Ties(k).pnt).body + l * 0.987         'tied robot stores some fat 'Botsareus 3/23/2016 Bugfix
                            If rob(.Ties(k).pnt).body > 32000 Then rob(.Ties(k).pnt).body = 32000
                            rob(.Ties(k).pnt).Waste = rob(.Ties(k).pnt).Waste + l * 0.01          'tied robot receives waste
                            rob(.Ties(k).pnt).radius = FindRadius(.Ties(k).pnt)
                            
                            .body = .body - l                                                       'tying robot gives up energy
                        End If
                        
                        'Taking body
                        If l < 0 Then
                                            
                            If l < -rob(.Ties(k).pnt).body Then l = -rob(.Ties(k).pnt).body ' Can't give away more body than you have
                                           
                            'Poison (Yes tiefeeding body is a reason enough to get poisoned)
                            ptag = Abs(l / 4)
                            If rob(.Ties(k).pnt).poison > ptag Then 'target robot has poison
                                If rob(.Ties(k).pnt).FName <> .FName Then 'can't poison your brother
                                
                                    .Poisoned = True
                                    .Poisoncount = .Poisoncount + ptag
                                    If .Poisoncount > 32000 Then .Poisoncount = 32000
                                    
                                    l = 0
                                    
                                    rob(.Ties(k).pnt).poison = rob(.Ties(k).pnt).poison - ptag
                                    rob(.Ties(k).pnt).mem(827) = rob(.Ties(k).pnt).poison
                                    
                                    If rob(.Ties(k).pnt).mem(834) > 0 Then
                                     .Ploc = ((rob(.Ties(k).pnt).mem(834) - 1) Mod 1000) + 1  'sets .Ploc to targets .mem(ploc) EricL - 3/29/2006 Added Mod to fix overflow
                                     If .Ploc = 340 Then .Ploc = 0
                                    Else
                                     Do
                                      .Ploc = Random(1, 1000)
                                     Loop Until .Ploc <> 340
                                    End If
        
                                    .Pval = rob(.Ties(k).pnt).mem(839)
                                    
                                End If
                            End If
                                                
                            .nrg = .nrg - l * 0.03             'tying robot receives energy
                            If .nrg > 32000 Then .nrg = 32000
                            .body = .body - l * 0.987          'tying robot stores some fat 'Botsareus 3/23/2016 Bugfix
                            If .body > 32000 Then .body = 32000
                            .Waste = .Waste - l * 0.01      'tying robot adds waste
                            .radius = FindRadius(t)
                            
                            rob(.Ties(k).pnt).body = rob(.Ties(k).pnt).body + l 'Take the body
                            
                            If rob(.Ties(k).pnt).body <= 0 And rob(.Ties(k).pnt).Dead = False Then 'Botsareus 3/11/2014 Tie feeding kills
                                .Kills = .Kills + 1
                                If .Kills > 32000 Then .Kills = 32000
                                .mem(220) = .Kills
                            End If
                        End If
                        
                        If Not .Ties(k).back Then   'forward ties
                            .Ties(k).nrgused = True   'red ties
                        Else                        'backward ties
                            rob(.Ties(k).pnt).Ties(.Ties(k).ptt).nrgused = True 'red ties
                        End If
                        
                    End If
                        
                k = k + 1
                Wend
            End If
                
            .mem(tp + 2) = 0
            .mem(tp + 3) = 0
        End If

      
    .mem(tp + 5) = 0 ' .tienum should be reset every cycle
getout:
  End With
End Sub
Public Sub EraseTRefVars(t As Integer)
Dim counter As Integer

  With rob(t)
   ' Zero out the trefvars as all ties have gone.  Perf -> Could set a flag to not do this everytime
      For counter = 456 To 465
         .mem(counter) = 0
      Next counter
      .mem(trefbody) = 0   'trefbody
      .mem(475) = 0        'tmemval
     ' .mem(476) = 0       'tmemloc EricL 4/20/2006 Commented out.  Should persist even when tie goes away or no tienum is specified
      .mem(478) = 0        'treffixed
      .mem(479) = 0        'trefaim
      For counter = 0 To 10 'For trefvelX functions.
        .mem(trefxpos + counter) = 0
      Next counter
       
      'These are .tin trefvars
      For counter = 420 To 429
        .mem(counter) = 0
      Next counter
  End With
End Sub


Public Sub readtie(ByVal t As Integer) 'Botsareus 2/11/2014 Bug fix

'reads all of the tref variables from a given tie number
  Dim k As Integer
  Dim tn As Integer
  Dim counter
 
  With rob(t)
    If rob(t).newage < 2 Then Exit Sub 'Botsareus 3/6/2013 Bug fix: Robot must be fully loaded before checking ties
  
    If .numties = 0 Then
      EraseTRefVars (t)
      GoTo getout
    Else
      ' If there is a value in .readtie then get the trefvars from that tie else read the trefvars from the last tie created
      If .mem(471) <> 0 Then
        tn = .mem(471) ' .readtie
      Else
        tn = .mem(454) ' .tiepres
      End If
      k = 1
      While .Ties(k).pnt > 0
        If .Ties(k).Port = tn Then
          ReadTRefVars t, k
          GoTo getout
        End If
        k = k + 1
      Wend
        'If we got here, no tie exists with this number.
        EraseTRefVars (t) ' Zero the trefvars.  The bot might be checking if the tie still exists.
    End If
getout:
  End With
End Sub

'EricL 4/20/2006  This was the heart of readtie.  Seperated it out so Trefvars can be loaded when tie is formed.
'Reads the Tie Refvars for tie k into the mem of bot t
Public Sub ReadTRefVars(t As Integer, k As Integer)
  Dim l As Integer ' just a loop counter

  With rob(t)
    If rob(.Ties(k).pnt).nrg < 32000 And rob(.Ties(k).pnt).nrg > -32000 Then
      .mem(464) = CInt(rob(.Ties(k).pnt).nrg) 'copies tied robot's energy into memory cell *trefnrg
    End If
    If rob(.Ties(k).pnt).age < 32000 Then
      .mem(465) = rob(.Ties(k).pnt).age + 1 'copies age of tied robot into *refvar
    Else
      .mem(465) = 32000
    End If
    If rob(.Ties(k).pnt).body < 32000 And rob(.Ties(k).pnt).body > -32000 Then
      .mem(trefbody) = CInt(rob(.Ties(k).pnt).body)  'copies tied robot's body value
    Else
      .mem(trefbody) = 32000
    End If
    
    For l = 1 To 8    'copies all ref vars from tied robot
      .mem(455 + l) = rob(.Ties(k).pnt).occurr(l)
    Next l
    
    If .mem(476) > 0 And .mem(476) <= 1000 Then   'tmemval and tmemloc couple used to read a specific memory value from tied robot.
      .mem(475) = rob(.Ties(k).pnt).mem(.mem(476))
      If .mem(479) > EyeStart And .mem(479) < EyeEnd Then
        rob(.Ties(k).pnt).View = True
      End If
    End If
    If rob(.Ties(k).pnt).Fixed Then
      .mem(478) = 1
    Else
      .mem(478) = 0
    End If
    
    .mem(479) = rob(.Ties(k).pnt).mem(AimSys)
        
    .mem(trefxpos) = rob(.Ties(k).pnt).mem(219)
    .mem(trefypos) = rob(.Ties(k).pnt).mem(217)
    .mem(trefvelyourup) = rob(.Ties(k).pnt).mem(velup)
    .mem(trefvelyourdn) = rob(.Ties(k).pnt).mem(veldn)
    .mem(trefvelyoursx) = rob(.Ties(k).pnt).mem(velsx)
    .mem(trefvelyourdx) = rob(.Ties(k).pnt).mem(veldx)
                
    'Botsareus 9/27/2014 I was thinking long and hard where to place this bug fix, probebly best to place it at the source
    If Abs(rob(.Ties(k).pnt).vel.y) > 16000 Then rob(.Ties(k).pnt).vel.y = 16000 * Sgn(rob(.Ties(k).pnt).vel.y)
    If Abs(rob(.Ties(k).pnt).vel.x) > 16000 Then rob(.Ties(k).pnt).vel.x = 16000 * Sgn(rob(.Ties(k).pnt).vel.x)
    
    .mem(trefvelmyup) = rob(.Ties(k).pnt).vel.x * Cos(.aim) + Sin(.aim) * rob(.Ties(k).pnt).vel.y * -1 - .mem(velup) 'gives velocity from mybots frame of reference
    .mem(trefvelmydn) = .mem(trefvelmyup) * -1
    .mem(trefvelmydx) = rob(.Ties(k).pnt).vel.y * Cos(.aim) + Sin(.aim) * rob(.Ties(k).pnt).vel.x - .mem(veldx)
    .mem(trefvelmysx) = .mem(trefvelmydx) * -1
    .mem(trefvelscalar) = rob(.Ties(k).pnt).mem(velscalar)
   ' .mem(trefbody) = rob(.Ties(k).pnt).body
    .mem(trefshell) = rob(.Ties(k).pnt).shell
    
    'These are the tie in/out pairs
    For l = 410 To 419
      .mem(l + 10) = rob(.Ties(k).pnt).mem(l)
    Next l
        
  End With
End Sub

' deletes all ties of a robot a
Public Sub delallties(a As Integer)
  Dim i As Integer
  i = 1
  While rob(a).Ties(1).pnt <> 0 And i <= MAXTIES
    DeleteTie a, rob(a).Ties(1).pnt
    i = i + 1
  Wend
End Sub

' deletes a tie between robots a and b
Public Sub DeleteTie(ByVal a As Integer, ByVal b As Integer)
  Dim k As Integer
  Dim j As Integer
  Dim t As Integer
  
  'Quick tests to rule out whether a tie can't exist between the bots.
  If (Not rob(a).exist) Or (Not rob(b).exist) Then GoTo getout
  If rob(a).numties = 0 Or rob(b).numties = 0 Then GoTo getout
  
  k = 1
  j = 1
  
  'Only allows 9 ties at present.  Change this?
  
  While rob(a).Ties(k).pnt <> b And k < MAXTIES
    k = k + 1
  Wend
    
  While rob(b).Ties(j).pnt <> a And j < MAXTIES
    j = j + 1
  Wend
  
  If k < MAXTIES Then
    rob(a).numties = rob(a).numties - 1 ' Decrement numties
    rob(a).mem(numties) = rob(a).numties
    If rob(a).mem(TIEPRES) = rob(a).Ties(k).Port Then ' we are deleting the last tie created.  Have to update .tiepres.
      If k > 1 Then
        rob(a).mem(TIEPRES) = rob(a).Ties(k - 1).Port
      Else
        rob(a).mem(TIEPRES) = 0 ' no more ties
      End If
    End If
  End If
  
  If j < MAXTIES Then
    rob(b).numties = rob(b).numties - 1 ' Decrement numties
    rob(b).mem(numties) = rob(b).numties
    If rob(b).mem(TIEPRES) = rob(b).Ties(j).Port Then ' we are deleting the last tie created.  Have to update .tiepres.
      If j > 1 Then
        rob(b).mem(TIEPRES) = rob(b).Ties(j - 1).Port
      Else
        rob(b).mem(TIEPRES) = 0 ' no more ties
      End If
    End If
  End If
    
    
  For t = k To MAXTIES - 1
    rob(a).Ties(t) = rob(a).Ties(t + 1)
  Next t
      
  rob(a).Ties(MAXTIES).pnt = 0
  
  For t = j To MAXTIES - 1
    rob(b).Ties(t) = rob(b).Ties(t + 1)
  Next t
      
  rob(b).Ties(MAXTIES).pnt = 0
getout:
End Sub

'
' T I E S
'

' creates a tie between rob a and b,of len c, lasting last cycles
' (or waiting -last cycles before consolidating)
' tie is addressed with index mem (putting mem in .tienum)
Public Function maketie(ByVal a As Integer, ByVal b As Integer, c As Long, last As Integer, mem As Integer) As Boolean
'returns true on success
'Ties and slime need to be reworked at some point
  Dim k As Integer
  Dim j As Integer
  Dim OK As Boolean
  Dim Max As Integer
  Dim deflect As Integer
  Dim length As Long
  Dim deletedtie As Boolean
  
  maketie = False
  
  If rob(a).exist = False Then GoTo getout
 
  deflect = Random(2, 92) 'random number which allows for the effect of slime on the target robot. If slime is greater then no tie is formed
  Max = MAXTIES
  OK = True
  k = 1
  j = 1
  
  length = VectorMagnitude(VectorSub(robManager.GetRobotPosition(a), robManager.GetRobotPosition(b)))
    
  If length <= c * 1.5 Then 'And deflect > rob(b).slime Then
    If deflect < rob(b).Slime Then OK = False  'should stop ties forming when slime is high
    
    If OK = True Then DeleteTie a, b
    
    While rob(a).Ties(k).pnt > 0 And k <= Max And OK
      k = k + 1
    Wend
    While rob(b).Ties(j).pnt > 0 And j <= Max And OK
      j = j + 1
    Wend
    
    If k < Max And j < Max And OK Then
      rob(a).Ties(k).pnt = b
      rob(a).Ties(k).ptt = j
      rob(a).Ties(k).NaturalLength = length
      rob(a).Ties(k).stat = False
      rob(a).Ties(k).last = last
      rob(a).Ties(k).Port = mem
      rob(a).Ties(k).back = False
      rob(a).numties = k
      rob(a).mem(466) = rob(a).numties   'EricL 3/22/2006 Increment numties in the bot's memory
      rob(a).mem(TIEPRES) = mem
      ReadTRefVars a, k ' EricL 4./20/2006  Load up the trefvars for the bot that created the tie.
      
      'EricL 5/7/2006 All ties are springs when first created
      rob(a).Ties(k).b = 0.02
      rob(a).Ties(k).k = 0.01
      rob(a).Ties(k).type = 0
          
      rob(b).Ties(j).pnt = a
      rob(b).Ties(j).ptt = k
      rob(b).Ties(j).NaturalLength = length
      rob(b).Ties(j).stat = False
      rob(b).Ties(j).last = last
      rob(b).Ties(j).back = True
      rob(b).numties = j
      rob(b).Ties(j).Port = rob(b).numties ' The port of the tie from the point of view of the tied bot
      rob(b).mem(466) = rob(b).numties     'EricL 3/22/2006 Increment numties in the bot's memory
      rob(b).mem(TIEPRES) = j
      
      'EricL 5/7/2006 All ties are springs when first created
      rob(b).Ties(j).b = 0.02
      rob(b).Ties(j).k = 0.01
      rob(b).Ties(j).type = 0
    End If
  End If
  
  If rob(b).Slime > 0 Then rob(b).Slime = rob(b).Slime - 20
  If rob(b).Slime < 0 Then rob(b).Slime = 0 'cost to slime layer when attacked
  rob(a).nrg = rob(a).nrg - (SimOpts.Costs(TIECOST) * SimOpts.Costs(COSTMULTIPLIER)) / (IIf(rob(a).numties < 0, 0, rob(a).numties) + 1) 'Tie cost to form tie
getout:
End Function

' searches a tie in rob t pointing to rob p
' returns tie number (j) of the tie pointing to the specified robot
Public Function srctie(t As Integer, p As Integer) As Integer
  Dim j As Integer
  j = 1
  srctie = 0
  With rob(t)
  While .Ties(j).pnt > 0 And srctie = 0
    If .Ties(j).pnt = p And .Ties(j).last < 1 Then
      srctie = j
    End If
    j = j + 1
  Wend
  End With
End Function

'fixes tie angle and length at whatever it currently is
Public Sub regang(t As Integer, j As Integer)
  Dim n As Integer
  Dim angl As Single
  Dim dist As Single
  With rob(t)
      .Multibot = True: .mem(multi) = 1
      .Ties(j).b = 0.1
      .Ties(j).k = 0.05
      .Ties(j).type = 3
      n = .Ties(j).pnt
      angl = angle(robManager.GetRobotPosition(t).x, robManager.GetRobotPosition(t).y, robManager.GetRobotPosition(n).x, robManager.GetRobotPosition(n).y)
      dist = Sqr((robManager.GetRobotPosition(t).x - robManager.GetRobotPosition(n).x) ^ 2 + (robManager.GetRobotPosition(t).y - robManager.GetRobotPosition(n).y) ^ 2)
      If .Ties(j).back = False Then
        .Ties(j).ang = AngDiff(angnorm(angl), angnorm(rob(t).aim)) ' only fix the angle of the bot that created the tie
        .Ties(j).angreg = True
      End If
      .Ties(j).NaturalLength = dist
  End With
End Sub

' bends a tie
Public Sub bend(t As Integer, lnk As Integer, ang As Integer)
  Dim an As Single
  If Abs(ang) > 100 Then ang = 100 * Sgn(ang)
  an = ang / 100
  With rob(t).Ties(lnk)
    .bend = an
    rob(.pnt).Ties(.ptt).bend = -an
  End With
  ang = 0
End Sub

' shrinks a tie
Public Sub shrink(t As Integer, lnk As Integer, ln As Integer)
  If Abs(ln) > 100 Then ln = 1000 * Sgn(ln) ' EricL 5/7/2006 Changed from 100 to 1000
  With rob(t).Ties(lnk)
    .shrink = ln
    rob(.pnt).Ties(.ptt).shrink = ln
  End With
  ln = 0
End Sub
