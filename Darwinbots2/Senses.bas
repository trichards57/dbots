Attribute VB_Name = "Senses"
'
'     S E N S E S
'

'This module is the most processor intensive.

Option Explicit

' Sets .sun to 1 if robot.aim is within 0.18 radians of 1.57 (Basically up)
' new version with less clutter.
Public Sub LandMark(ByVal iRobID As Integer)

rob(iRobID).mem(LandM) = 0
If rob(iRobID).aim > 1.39 And rob(iRobID).aim < 1.75 Then rob(iRobID).mem(LandM) = 1

End Sub

' touch: tells a robot whether it has been hit by another one
' and where: up, dn dx, sx
Public Sub touch(ByVal a As Long, ByVal x As Long, ByVal y As Long)
  Dim xc As Single
  Dim yc As Single
  Dim dx As Single
  Dim dy As Single
  Dim tn As Single
  Dim ang As Single
  Dim aim As Single
  Dim dang As Single
  aim = 6.28 - rob(a).aim
  xc = robManager.GetRobotPosition(a).x
  yc = robManager.GetRobotPosition(a).y
  dx = x - xc
  dy = y - yc
  
  If dx <> 0 Then
    tn = dy / dx
    ang = Atn(tn)
    If dx < 0 Then ang = ang - 3.14
  Else
    ang = 1.57 * Sgn(dy)
  End If
  
  dang = ang - aim
  While dang < 0
    dang = dang + 6.28
  Wend
  While dang > 6.28
    dang = dang - 6.28
  Wend
  If dang > 5.49 Or dang < 0.78 Then rob(a).mem(hitup) = 1
  If dang > 2.36 And dang < 3.92 Then rob(a).mem(hitdn) = 1
  If dang > 0.78 And dang < 2.36 Then rob(a).mem(hitdx) = 1
  If dang > 3.92 And dang < 5.49 Then rob(a).mem(hitsx) = 1
  rob(a).mem(hit) = 1
End Sub

' taste: same as for touch, but for shots, and gives back
' also the flavour of the shot, that is, its shottype
' value
Public Sub taste(a As Integer, ByVal x As Single, ByVal y As Single, value As Integer)
  Dim xc As Single
  Dim yc As Single
  Dim dx As Single
  Dim dy As Single
  Dim tn As Single
  Dim ang As Single
  Dim aim As Single
  Dim dang As Single
  aim = 6.28 - rob(a).aim
  xc = robManager.GetRobotPosition(a).x
  yc = robManager.GetRobotPosition(a).y
  dx = x - xc
  dy = y - yc
  If dx <> 0 Then
    tn = dy / dx
    ang = Atn(tn)
    If dx < 0 Then ang = ang - 3.14
  Else
    ang = 1.57 * Sgn(dy)
  End If
  dang = ang - aim
  While dang < 0
    dang = dang + 6.28
  Wend
  While dang > 6.28
    dang = dang - 6.28
  Wend
  If dang > 5.49 Or dang < 0.78 Then rob(a).mem(shup) = value
  If dang > 2.36 And dang < 3.92 Then rob(a).mem(shdn) = value
  If dang > 0.78 And dang < 2.36 Then rob(a).mem(shdx) = value
  If dang > 3.92 And dang < 5.49 Then rob(a).mem(shsx) = value
  rob(a).mem(209) = dang * 200  'sysvar = .shang just returns the angle of the shot without the flavor
  rob(a).mem(shflav) = value    'sysvar = .shflav returns the flavor without the angle
End Sub

' erases some senses
Public Sub EraseSenses(n As Integer)
  Dim l As Integer
  With rob(n)
    .lasttch = 0 'Botsareus 11/26/2013 Erase lasttch here
    .mem(hitup) = 0
    .mem(hitdn) = 0
    .mem(hitdx) = 0
    .mem(hitsx) = 0
    .mem(hit) = 0
    .mem(shflav) = 0
    .mem(209) = 0 '.shang
    .mem(shup) = 0
    .mem(shdn) = 0
    .mem(shdx) = 0
    .mem(shsx) = 0
    .mem(214) = 0   'edge collision detection
    EraseLookOccurr (n)
    
   'EricL - *trefvars now persist across cycles
   ' For l = 1 To 10 ' resets *trefvars
   '   .mem(455 + l) = 0
   ' Next l
   ' For l = 0 To 10     'resets
   '     .mem(trefxpos + l) = 0
   ' Next l
   ' .mem(472) = 0
  End With
End Sub

'Public Function BasicProximity(n As Integer, Optional force As Boolean = False) As Integer 'returns .lastopp
'  Dim counter As Integer
'  Dim u As vector
'  Dim dotty As Long, crossy As Long
'  Dim x As Integer
'
'  'until I get some better data structures, this will ahve to do
'
'  rob(n).lastopp = 0
'  rob(n).lastopptype = 0 ' set the default type of object seen to a bot.
'  rob(n).mem(EYEF) = 0
'  For x = EyeStart + 1 To EyeEnd - 1
'    rob(n).mem(x) = 0
'  Next x
'
'  'We have to populate eyes for every bot, even for those without .eye sysvars
'  'since they could evolve indirect addressing of the eye sysvars.
'  For counter = 1 To MaxRobs
'    If n <> counter And rob(counter).exist Then
'       CompareRobots3 n, counter
'    End If
'  Next counter
'
'  If SimOpts.shapesAreVisable And rob(n).exist Then CompareShapes n, 12
'
'  BasicProximity = rob(n).lastopp ' return the index of the last viewed object
'End Function

'Returns the index into the Specie array to which a given bot conforms
Public Function SpeciesFromBot(n As Integer) As Integer
Dim i As Integer

i = 0
While SimOpts.Specie(i).Name <> rob(n).FName And i < SimOpts.SpeciesNum
   i = i + 1
Wend
SpeciesFromBot = i
End Function


' writes some senses: view, .ref* vars, absvel
' pain, pleas, nrg
Public Sub WriteSenses(ByVal n As Integer)
  Dim t As Integer
  Dim i As Integer
  Dim temp As Single
  
  LandMark n
  With rob(n)
   
    .mem(TotalBots) = TotalRobots
    .mem(TOTALMYSPECIES) = SimOpts.Specie(SpeciesFromBot(n)).population
    
    If Not .CantSee And Not .Corpse Then
      If BucketsProximity(n) > 0 Then
      'If BasicProximity(n) > 0 Then
        'There is somethign visable in the focus eye
        If .lastopptype = 0 Then lookoccurr n, .lastopp ' It's a bot.  Populate the refvar sysvars
      End If
    End If

    If .nrg > 32000 Then .nrg = 32000
    If .onrg < 0 Then .onrg = 0
    If .obody < 0 Then .obody = 0
    If .nrg < 0 Then .nrg = 0

    .mem(pain) = CInt(.onrg - .nrg)
    .mem(pleas) = CInt(.nrg - .onrg)
    .mem(bodloss) = CInt(.obody - .body)
    .mem(bodgain) = CInt(.body - .obody)
    
    .onrg = .nrg
    .obody = .body
    .mem(Energy) = CInt(.nrg)
    If .age = 0 And .mem(body) = 0 Then .mem(body) = .body 'to stop an odd bug in birth.  Don't ask
    If .Fixed Then .mem(215) = 1 Else .mem(215) = 0
    Dim pos As Vector
    pos = robManager.GetRobotPosition(n)
    If pos.y < 0 Then pos.y = 0
    temp = Int((pos.y / Form1.yDivisor) / 32000#)
    temp = (pos.y / Form1.yDivisor) - (temp * 32000#)
    .mem(217) = CInt(temp Mod 32000)
    If pos.x < 0 Then pos.x = 0
    temp = Int((pos.x / Form1.xDivisor) / 32000#)
    temp = (pos.x / Form1.xDivisor) - (temp * 32000#)
    .mem(219) = CInt(temp Mod 32000)
    
    robManager.SetRobotPosition n, pos
  End With
End Sub

' copies the occurr array of a viewed robot
' in the ref* vars of the viewing one
Public Sub lookoccurr(ByVal n As Integer, ByVal o As Integer)
  If rob(n).Corpse Then GoTo getout
  Dim t As Byte
  Dim x As Single
  Dim y As Single
  
  rob(n).mem(REFTYPE) = 0
  
  For t = 1 To 8
    rob(n).mem(occurrstart + t) = rob(o).occurr(t)
  Next t
  
  If rob(o).nrg < 0 Then
     rob(n).mem(occurrstart + 9) = 0
  ElseIf rob(o).nrg < 32001 Then
    rob(n).mem(occurrstart + 9) = rob(o).nrg
  Else
    rob(n).mem(occurrstart + 9) = 32000
  End If
  'EricL 4/13/2006 Added If Then now that age can exceed 32000
  If rob(o).age < 32001 Then
    rob(n).mem(occurrstart + 10) = rob(o).age '.refage
  Else
    rob(n).mem(occurrstart + 10) = 32000
  End If
  
  rob(n).mem(in1) = rob(o).mem(out1)
  rob(n).mem(in2) = rob(o).mem(out2)
  rob(n).mem(in3) = rob(o).mem(out3)
  rob(n).mem(in4) = rob(o).mem(out4)
  rob(n).mem(in5) = rob(o).mem(out5)
  rob(n).mem(in6) = rob(o).mem(out6)
  rob(n).mem(in7) = rob(o).mem(out7)
  rob(n).mem(in8) = rob(o).mem(out8)
  rob(n).mem(in9) = rob(o).mem(out9)
  rob(n).mem(in10) = rob(o).mem(out10)
  rob(n).mem(711) = rob(o).mem(18)      'refaim
  rob(n).mem(712) = rob(o).occurr(9)    'reftie
  rob(n).mem(refshell) = rob(o).shell
  rob(n).mem(refbody) = rob(o).body
  rob(n).mem(refypos) = rob(o).mem(217)
  rob(n).mem(refxpos) = rob(o).mem(219)
  'give reference variables from the bots frame of reference
  x = (rob(o).vel.x * Cos(rob(n).aim) + rob(o).vel.y * Sin(rob(n).aim) * -1) - rob(n).mem(velup)
  y = (rob(o).vel.y * Cos(rob(n).aim) + rob(o).vel.x * Sin(rob(n).aim)) - rob(n).mem(veldx)
  If x > 32000 Then x = 32000
  If x < -32000 Then x = -32000
  If y > 32000 Then y = 32000
  If y < -32000 Then y = -32000
  
  rob(n).mem(refvelup) = x
  rob(n).mem(refveldn) = rob(n).mem(refvelup) * -1
  rob(n).mem(refveldx) = y
  rob(n).mem(refvelsx) = rob(n).mem(refvelsx) * -1
  Dim temp As Single
  temp = Sqr(CLng(rob(n).mem(refvelup) ^ 2) + CLng(rob(n).mem(refveldx) ^ 2))  ' how fast is this robot moving compared to me?
  If temp > 32000 Then temp = 32000
  
  rob(n).mem(refvelscalar) = temp
  rob(n).mem(713) = rob(o).mem(827)     'refpoison. current value of poison. not poison commands
  rob(n).mem(714) = rob(o).mem(825)     'refvenom (as with poison)
  rob(n).mem(715) = rob(o).Kills        'refkills
   If rob(o).Multibot = True Then
    rob(n).mem(refmulti) = 1
  Else
    rob(n).mem(refmulti) = 0
  End If
  If rob(n).mem(474) > 0 And rob(n).mem(474) <= 1000 Then 'readmem and memloc couple used to read a specified memory location of the target robot
    rob(n).mem(473) = rob(o).mem(rob(n).mem(474))
    If rob(n).mem(474) > EyeStart And rob(n).mem(474) < EyeEnd Then
        rob(o).View = True
    End If
  End If
  If rob(o).Fixed Then                  'reffixed. Tells if a viewed robot is fixed by .fixpos.
    rob(n).mem(477) = 1
  Else
    rob(n).mem(477) = 0
  End If
 ' rob(n).mem(825) = Int(rob(o).venom)
  'rob(n).mem(827) = Int(rob(o).poison)
getout:
End Sub

' Erases the occurr array
Public Sub EraseLookOccurr(ByVal n As Integer)
  
  If rob(n).Corpse Then GoTo getout
  
  Dim t As Byte
  
  rob(n).mem(REFTYPE) = 0
  
  For t = 1 To 10
    rob(n).mem(occurrstart + t) = 0
  Next t
  
  rob(n).mem(in1) = 0
  rob(n).mem(in2) = 0
  rob(n).mem(in3) = 0
  rob(n).mem(in4) = 0
  rob(n).mem(in5) = 0
  rob(n).mem(in6) = 0
  rob(n).mem(in7) = 0
  rob(n).mem(in8) = 0
  rob(n).mem(in9) = 0
  rob(n).mem(in10) = 0
  
  rob(n).mem(711) = 0      'refaim
  rob(n).mem(712) = 0    'reftie
  rob(n).mem(refshell) = 0
  rob(n).mem(refbody) = 0
  rob(n).mem(refypos) = 0
  rob(n).mem(refxpos) = 0
  rob(n).mem(refvelup) = 0
  rob(n).mem(refveldn) = 0
  rob(n).mem(refveldx) = 0
  rob(n).mem(refvelsx) = 0
  rob(n).mem(refvelscalar) = 0
  rob(n).mem(713) = 0        'refpoison. current value of poison. not poison commands
  rob(n).mem(714) = 0        'refvenom (as with poison)
  rob(n).mem(715) = 0        'refkills
  rob(n).mem(refmulti) = 0
  rob(n).mem(473) = 0
  rob(n).mem(477) = 0
getout:
End Sub

' creates the array which is copied to the ref* variables
' of any opponent looking at us
Public Sub makeoccurrlist(n As Integer)
  Dim t As Long
  Dim k As Integer
  With rob(n)
  
    
    For t = 1 To 12
      .occurr(t) = 0
    Next t
    t = 1
    k = 1
    While Not (.dna(t).type = 10 And .dna(t).value = 1) And t <= 32000 And t < UBound(.dna) 'Botsareus 6/16/2012 Added code to check upper bounds
      
      If .dna(t).type = 0 Then 'number
      
        If .dna(t + 1).type = 7 Then 'DNA is going to store to this value, so it's probably a sysvar
          
          If .dna(t).value < 8 And .dna(t).value > 0 Then 'if we are dealing with one of the first 8 sysvars
            .occurr(.dna(t).value) = .occurr(.dna(t).value) + 1 'then the occur listing for this fxn is incremented
          End If
          
          If .dna(t).value = 826 Then 'referencing .strpoison
            .occurr(10) = .occurr(10) + 1
          End If
        
          If .dna(t).value = 824 Then  'refencing .strvenom
            .occurr(11) = .occurr(11) + 1
          End If
              
        End If
        
        If .dna(t).value = 330 Then 'the bot is referencing .tie 'Botsareus 11/29/2013 Moved to "." list
          .occurr(9) = .occurr(9) + 1 'ties
        End If
        
      End If
      
      If .dna(t).type = 1 Then '*number
        If .dna(t).value > 500 And .dna(t).value < 510 Then 'the bot is referencing an eye
          .occurr(8) = .occurr(8) + 1 'eyes
        End If
      End If
      
      t = t + 1
    Wend

    'creates the "ownvars" our own readbacks as versions of the refvars seen by others
    For t = 1 To 8
      .mem(720 + t) = .occurr(t)
    Next t
    .mem(728) = .occurr(8)
    .mem(729) = .occurr(9)
    .mem(730) = .occurr(10)
    .mem(731) = .occurr(11)
  End With
End Sub

