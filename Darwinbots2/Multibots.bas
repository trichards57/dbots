Attribute VB_Name = "Multibots"
Option Explicit

'
' M U L T I C E L L U L A R   R O U T I N E S
'

' moves the organism of which robot n is part to the position x,y
Public Sub ReSpawn(n As Integer, x As Single, y As Single)
  Dim clist(50) As Integer 'changed from 20 to 50
  Dim Min As Single, nmin As Integer
  Dim t As Integer, dx As Single, dy As Single
  Dim radiidif As Single
  clist(0) = n
  ListCells clist()
  Min = 999999999999#
  t = 0
  While clist(t) > 0
    If ((robManager.GetRobotPosition(clist(t)).x - x) ^ 2 + (robManager.GetRobotPosition(clist(t)).y - y) ^ 2) <= Min Then
      Min = (robManager.GetRobotPosition(clist(t)).x - x) ^ 2 + (robManager.GetRobotPosition(clist(t)).y - y) ^ 2
      nmin = clist(t)
    End If
    t = t + 1
    If t > 50 Then GoTo getout
  Wend
  dx = x - robManager.GetRobotPosition(nmin).x
  dy = y - robManager.GetRobotPosition(nmin).y
  
  'Botsareus 7/15/2016 Bug fix: corrects by radii difference between the two robots
  radiidif = rob(n).radius - rob(nmin).radius
  
  dx = dx - 1 * Sgn(dx) + Sgn(dx) * radiidif
  dy = dy - 1 * Sgn(dy) + Sgn(dy) * radiidif
  
  t = 0
  While clist(t) > 0
  Dim pos As Vector
  
    pos.x = robManager.GetRobotPosition(clist(t)).x + dx
    pos.y = robManager.GetRobotPosition(clist(t)).y + dy
    robManager.SetRobotPosition clist(t), pos
    'Botsareus 7/6/2016 Make sure to resolve actvel
    rob(clist(t)).opos.x = pos.x
    rob(clist(t)).opos.y = pos.y
    UpdateBotBucket clist(t)
    t = t + 1
  Wend
getout:
End Sub

' kill organism
Public Sub KillOrganism(n As Integer)
  Dim clist(50) As Integer, t As Integer    'changed from 20 to 50
  Dim temp As Boolean
  clist(0) = n
  ListCells clist()
  temp = MDIForm1.nopoff
  MDIForm1.nopoff = True
  While clist(t) > 0
    KillRobot clist(t)
    t = t + 1
  Wend
  MDIForm1.nopoff = temp
End Sub

' selects the whole organism
Public Sub FreezeOrganism(n As Integer)
  Dim clist(50) As Integer, t As Integer    'changed from 20 to 50
  clist(0) = n
  ListCells clist()
  While clist(t) > 0
    rob(clist(t)).highlight = True
    t = t + 1
  Wend
End Sub

' lists all the cells of an organism, starting from any one
' in position lst(0). Leaves the result in array lst()
Public Sub ListCells(lst() As Integer)
  Dim k As Integer
  Dim j As Integer
  Dim w As Integer
  Dim pres As Boolean
  Dim n As Long
  w = 0
  n = lst(w)
  
  While n > 0
    With rob(n)
    If Not rob(n).Multibot Then GoTo skipties  ' If the bot isn't a multibot, then ignore connected cells
    k = 1
    While .Ties(k).pnt > 0
      pres = False
      j = 0
      While lst(j) > 0
        If lst(j) = .Ties(k).pnt Then pres = True
        j = j + 1
        If j = 50 Then lst(j) = 0
      Wend
      If Not pres Then lst(j) = .Ties(k).pnt
      k = k + 1
    Wend
skipties:
   End With
    w = w + 1
    If w > 50 Then
      w = 50   'don't know what effect this will have. Should stop overflows
      lst(w) = 0 'EricL - added June 2006 to prevent overflows
      GoTo getout
    End If
    n = lst(w)
  Wend
getout:
End Sub

'Made obsolete by TieHooke
'Public Sub MB_Transfer_Accelerations(n As Integer)
''calculates accelerations of a robot that is part of an MB
''and applies a fraction of the acceleration to any other robot
''joined to it by a tie
'Dim pt As Integer
'  Dim j As Integer
'  Dim L As Long
'  Dim k As Integer
'  Dim tvel As Single
'  Dim ivel As Single
'  Dim cost As Single
'  Dim Absaccel As Single
'  Dim NewAccelx As Single
'  Dim NewAccely As Single
'  Dim Reduce As Single
'  Dim up As Integer, dn As Integer, dx As Integer, sx As Integer
'
'  With rob(n)
'  If .Exist = False Then Exit Sub
'    .mass = (.body / 1000) + (.Shell / 200)   'set value for mass
'    If .mass = 0 Then .mass = 0.001
'    Absaccel = 0                'reset acceleration
'    .absvel = Cos(.aim) * .vel.x * -1 + Sin(.aim) * .vel.y 'formula changed to give velocity in the direction robot is facing rather than always a positive number. Make *.vel work properly.
'    '.mem(vel) = .absvel * -1
'
'    up = .mem(dirup)
'    dn = .mem(dirdn)
'    dx = .mem(dirdx)
'    sx = .mem(dirsx)
'
'    NewAccelx = absx(.aim, up, dn, sx, dx) * SimOpts.PhysMoving
'    .ax = .ax + NewAccelx
'    NewAccely = absy(.aim, up, dn, sx, dx) * SimOpts.PhysMoving
'    .ay = .ay + NewAccely
'    Absaccel = Sqr(.ax ^ 2 + .ay ^ 2)
'    .ax = .ax / .mass       'having large mass doesn't cost more. You just lose acceleration
'    .ay = .ay / .mass
'    ivel = .absvel
'    tvel = .absvel + Sqr(.ax ^ 2 + .ay ^ 2)
'    If tvel > .MaxSpeed Then       'limits speed to maxspeed
'      Reduce = tvel / .MaxSpeed
'      .ax = .ax / Reduce
'      .ay = .ay / Reduce
'      tvel = .MaxSpeed
'    End If
'    'transfer acceleration to other parts of the MB
'    k = 1
'    While .Ties(k).pnt <> 0
'      rob(.Ties(k).pnt).ax = rob(.Ties(k).pnt).ax + NewAccelx
'      rob(.Ties(k).pnt).ay = rob(.Ties(k).pnt).ay + NewAccely
'      k = k + 1
'    Wend
'  End With
'End Sub
