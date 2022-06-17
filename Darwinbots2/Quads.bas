Attribute VB_Name = "Buckets_Module"
Option Explicit

'Using a bucket size of 4000.  3348 plus twice radius of the largest possible bot is the farthest possible a bot can see.  4000 is a
'nice round number.
Public Const BucketSize As Long = 4000
Dim NumXBuckets As Integer ' Field Width divided by bucket size
Dim NumYBuckets As Integer ' Field height divided by bucket size

'This is the buckets Array
Dim Buckets() As BucketType

Private Type BucketType
  arr() As Integer
  size As Integer 'number of bots in the bucket i.e. highest array element with a bot
  adjBucket(8) As Vector ' List of buckets adjoining this one.  Interior buckets will have 8.  Edge buckets 5.  Corners 3.
End Type

Public eyeDistance(10) As Single  ' used for exact distances to viewed objects for displaying the eye viewer for the focus bot

'also erases array elements to retrieve memory
Public Sub Init_Buckets()
  Dim x As Integer: Dim y As Integer: Dim z As Integer
  
  'Determine the nubmer of buckets.
  NumXBuckets = Int(SimOpts.FieldWidth / BucketSize)
  NumYBuckets = Int(SimOpts.FieldHeight / BucketSize)
  
  ReDim Buckets(NumXBuckets, NumYBuckets)
    
  'Buckets count along rows, top row, then next...
  For y = 0 To NumYBuckets - 1
    For x = 0 To NumXBuckets - 1
      ReDim Buckets(x, y).arr(0)
      Buckets(x, y).size = 0
      
      'Zero out the list of adjacent buckets
      For z = 1 To 8
        Buckets(x, y).adjBucket(z).x = -1
      Next z
      
      z = 1
      'Set the list of adjacent buckets for this bucket
      'We take the time to do this here to save the time it would take to compute these every cycle.
      If x > 0 Then Buckets(x, y).adjBucket(z).x = x - 1: Buckets(x, y).adjBucket(z).y = y: z = z + 1 ' Bucket to the Left
      If x < NumXBuckets - 1 Then Buckets(x, y).adjBucket(z).x = x + 1: Buckets(x, y).adjBucket(z).y = y: z = z + 1 ' Bucket to the Right
      If y > 0 Then Buckets(x, y).adjBucket(z).y = y - 1: Buckets(x, y).adjBucket(z).x = x: z = z + 1 ' Bucket on top
      If y < NumYBuckets - 1 Then Buckets(x, y).adjBucket(z).y = y + 1: Buckets(x, y).adjBucket(z).x = x: z = z + 1 ' Bucket below
      If x > 0 And y > 0 Then Buckets(x, y).adjBucket(z).x = x - 1: Buckets(x, y).adjBucket(z).y = y - 1: z = z + 1 ' Bucket to the Left and Up
      If x > 0 And y < NumYBuckets - 1 Then Buckets(x, y).adjBucket(z).x = x - 1: Buckets(x, y).adjBucket(z).y = y + 1: z = z + 1 ' Bucket to the Left and Down
      If x < NumXBuckets - 1 And y > 0 Then Buckets(x, y).adjBucket(z).x = x + 1: Buckets(x, y).adjBucket(z).y = y - 1: z = z + 1 ' Bucket to the Right and Up
      If x < NumXBuckets - 1 And y < NumYBuckets - 1 Then Buckets(x, y).adjBucket(z).x = x + 1: Buckets(x, y).adjBucket(z).y = y + 1: z = z + 1  ' Bucket to the Right and Down
    Next x
  Next y
  
  For x = 1 To MaxRobs
    If rob(x).exist Then
      rob(x).BucketPos.x = -2
      rob(x).BucketPos.y = -2
      UpdateBotBucket x
    End If
  Next x
End Sub

Public Sub UpdateBotBucket(n As Integer)
  'makes calls to Add_Bot and Delete_Bot
  'if we move out of our bucket
  'call this from outside the function
    
  Dim currbucket As Single, newbucket As Vector, changed As Boolean
  
  If Not rob(n).exist Then
    Delete_Bot n, rob(n).BucketPos
    GoTo getout
  End If
   
  
    newbucket = rob(n).BucketPos
    currbucket = Int(robManager.GetRobotPosition(n).x / BucketSize)
    If currbucket < 0 Then currbucket = 0 ' Possible bot is off the field
    If currbucket >= NumXBuckets Then currbucket = NumXBuckets - 1 ' Possible bot is off the field
  
    If rob(n).BucketPos.x <> currbucket Then
      'we've moved off the bucket, update bucket
      newbucket.x = currbucket
      changed = True
    End If
  
    currbucket = Int(robManager.GetRobotPosition(n).y / BucketSize)
    If currbucket < 0 Then currbucket = 0 ' Possible bot is off the field
    If currbucket >= NumYBuckets Then currbucket = NumYBuckets - 1 ' Possible bot is off the field
  
    If rob(n).BucketPos.y <> currbucket Then
      newbucket.y = currbucket
      changed = True
    End If
  
    If changed Then
      Delete_Bot n, rob(n).BucketPos
      Add_Bot n, newbucket
      rob(n).BucketPos = newbucket
    End If

getout:
End Sub

Public Sub Add_Bot(n As Integer, pos As Vector)
  Dim a As Integer

'Will grow the bucket's array if necessary
'.size is always the total length of the array.
'Array is packed.  Code can assume no more bots exist in the bucket if a -1 is encounterred

    
  With Buckets(pos.x, pos.y)

    For a = 1 To .size
      If .arr(a) = -1 Then
        .arr(a) = n
        GoTo getout
      End If
    Next a

    'we have to add it to the end somewhere
    ReDim Preserve .arr(.size + 5) 'faster to redim 5 at a time
    .arr(.size + 1) = n
    .arr(.size + 2) = -1
    .arr(.size + 3) = -1
    .arr(.size + 4) = -1
    .arr(.size + 5) = -1
    .size = .size + 5
getout:
  End With
End Sub

Public Sub Delete_Bot(n As Integer, pos As Vector)
  Dim a As Integer, b As Integer, c As Integer

'Removes a bot fro a bucket
'Keeps the array packed
'Redimension the array to recover memory if warrented

  If pos.x < 0 Or pos.y < 0 Then GoTo getout ' Can happen for new bots.  They arn't in any buckets.
  If pos.x > NumXBuckets - 1 Or pos.y > NumYBuckets - 1 Then GoTo getout ' Can happen when field is resized
  
    For a = 1 To Buckets(pos.x, pos.y).size
      If Buckets(pos.x, pos.y).arr(a) = n Then 'we've found the bot
                 
        'Slide all the bots in the bucket down one slot, effectivly deleting the specific bot
        While Buckets(pos.x, pos.y).arr(a) <> -1 And a < Buckets(pos.x, pos.y).size
          Buckets(pos.x, pos.y).arr(a) = Buckets(pos.x, pos.y).arr(a + 1)
          a = a + 1
        Wend
        
        'a now points to either .size or the last -1 slot.  If .size, we need to set the location to -1
        'Either way, it doesn't hurt to stomp on it.
        Buckets(pos.x, pos.y).arr(a) = -1
                
        'The array is now compact and a now points to the first -1 slot in the array.
        'We should reclaim memory if there is a lot to reclaim, up to a point.
        'For now, we only reclaim memory if more than 50 open slots
        If Buckets(pos.x, pos.y).size - a > 50 And Buckets(pos.x, pos.y).size > 55 Then 'last bot in array, collapse array
          ReDim Preserve Buckets(pos.x, pos.y).arr(Buckets(pos.x, pos.y).size - 50)
          Buckets(pos.x, pos.y).size = Buckets(pos.x, pos.y).size - 50
        End If
              
        GoTo getout
      End If
    Next a
getout:

End Sub

Public Function BucketsProximity(n As Integer) As Integer
  'mirror of proximity function.  Checks all the bots in the same bucket and surrounding buckets
  Dim x As Long, y As Long
  Dim BucketPos As Vector
  Dim adjBucket As Vector

  BucketPos = rob(n).BucketPos
  rob(n).lastopp = 0
  rob(n).lastopptype = 0 ' set the default type of object seen to a bot.
  rob(n).mem(EYEF) = 0
  For x = EyeStart + 1 To EyeEnd - 1
    rob(n).mem(x) = 0
  Next x
  
  'Check the bucket the bot is in
  CheckBotBucketForVision n, BucketPos
  
  'Check all the adjacent buckets
  For x = 1 To 8
    adjBucket = Buckets(BucketPos.x, BucketPos.y).adjBucket(x)
    If adjBucket.x <> -1 Then
      CheckBotBucketForVision n, adjBucket
    Else
      GoTo done
    End If
  Next x
done:
              
  BucketsProximity = rob(n).lastopp ' return the index of the last viewed object
End Function

Private Sub CheckBotBucketForVision(n As Integer, pos As Vector)
  Dim a As Integer, robnumber As Integer
  
  With Buckets(pos.x, pos.y)
  If .size = 0 Then GoTo getout
  a = 1
  While .arr(a) <> -1
    robnumber = .arr(a)
    If robnumber <> n Then CompareRobots3 n, robnumber
    If a = .size Then GoTo getout
    a = a + 1
  Wend
getout:
  End With
End Sub

Public Sub BucketsCollision(n As Integer)
  'mirror of proximity function.  Checks all the bots in the same bucket and surrounding buckets
  Dim x As Long, y As Long
  Dim BucketPos As Vector
  Dim adjBucket As Vector

  BucketPos = rob(n).BucketPos

  'Check the bucket the bot is in
  CheckBotBucketForCollision n, BucketPos
  
  'Check all the adjacent buckets
  For x = 1 To 8
    adjBucket = Buckets(BucketPos.x, BucketPos.y).adjBucket(x)
    If adjBucket.x <> -1 Then
      CheckBotBucketForCollision n, adjBucket
    Else
      GoTo done
    End If
  Next x
done:

End Sub


Private Sub CheckBotBucketForCollision(n As Integer, pos As Vector)
  Dim a As Integer, robnumber As Integer
  Dim k As Integer
  Dim distvector As Vector
  Dim dist As Single
  'If pos.x = -2 Or pos.Y = -2 Then goto getout

  If Buckets(pos.x, pos.y).size = 0 Then GoTo getout
  a = 1
    While Buckets(pos.x, pos.y).arr(a) <> -1
      robnumber = Buckets(pos.x, pos.y).arr(a)
      If robnumber > n Then ' only have to check bots higher than n otherwise we do it twice for each bot pair
        distvector = VectorSub(robManager.GetRobotPosition(n), robManager.GetRobotPosition(robnumber))
        dist = rob(n).radius + rob(robnumber).radius
        If VectorMagnitudeSquare(distvector) < (dist * dist) Then Repel3 n, robnumber
      End If
      If a = Buckets(pos.x, pos.y).size Then GoTo getout
      a = a + 1
    Wend
getout:

End Sub

'Returns the absolute width of an eye
Public Function AbsoluteEyeWidth(Width As Integer) As Integer
  If Width = 0 Then
    AbsoluteEyeWidth = 35
  Else
    AbsoluteEyeWidth = (Width Mod 1256) + 35
    If AbsoluteEyeWidth <= 0 Then AbsoluteEyeWidth = 1256 + AbsoluteEyeWidth
  End If
End Function


'Returns the absolute width of the narrowest eye of bot n
Public Function NarrowestEye(n As Integer) As Integer
Dim i As Integer
Dim Width As Integer

  NarrowestEye = 1221
  For i = 0 To 8
    Width = AbsoluteEyeWidth(rob(n).mem(EYE1WIDTH + i))
    If Width < NarrowestEye Then NarrowestEye = Width
  Next i
End Function

'Returns the distance an eye of absolute width w can see.
'Eye sight distance S varies as a function of eye width according to:  S =  1 - ln(w)/4
'where w is the absolute eyewidth as a multiple of the standard Pi/18 eyewidths
Public Function EyeSightDistance(w As Integer, n1 As Integer) As Single 'Botsareus 2/3/2013 modified to except robot id
  If w = 35 Then
    EyeSightDistance = 1440 * eyestrength(n1)
  Else
    EyeSightDistance = 1440 * (1 - (Log(w / 35) / 4)) * eyestrength(n1)
  End If
End Function

Private Function eyestrength(n1 As Integer) As Single 'Botsareus 2/3/2013 eye strength mod
Const EyeEffectiveness As Byte = 3  'Botsareus 3/26/2013 For eye strength formula

If SimOpts.Pondmode And robManager.GetRobotPosition(n1).y > 1 Then 'Botsareus 3/26/2013 Bug fix if robot Y pos is almost zero
  eyestrength = (EyeEffectiveness / (robManager.GetRobotPosition(n1).y / 2000) ^ SimOpts.Gradient) ^ (6828 / SimOpts.FieldHeight)  'Botsareus 3/26/2013 Robots only effected by density, not light intensity
Else
  eyestrength = 1
End If


If Not SimOpts.Daytime Then eyestrength = eyestrength * 0.8

If eyestrength > 1 Then eyestrength = 1

End Function

'New compare routine from EricL
'Takes into consideration movable eyes and eyes of variable width
Public Sub CompareRobots3(n1 As Integer, n2 As Integer)
      Dim ab As Vector, ac As Vector, ad As Vector 'vector from n1 to n2
      Dim invdist As Single, sightdist As Single, eyedist As Single, distsquared As Single
      Dim edgetoedgedist As Single, percentdist As Single
      Dim a As Integer
      Dim eyevalue As Single
      Dim eyeaim As Single
      Dim eyeaimleft As Single
      Dim eyeaimright As Single
      Dim beta As Single
      Dim theta As Single
      Dim halfeyewidth As Single
      Dim botspanszero As Boolean
      Dim eyespanszero As Boolean
      Dim eyesum As Long
             
      ab = VectorSub(robManager.GetRobotPosition(n2), robManager.GetRobotPosition(n1))
      edgetoedgedist = VectorMagnitude(ab) - rob(n1).radius - rob(n2).radius
      
      'Here we compute the maximum possible distance bot N1 can see.  Sight distance is a function of
      'eye width.  Narrower eyes can see farther, wider eyes not so much.  So, we find the narrowest eye
      'and use that to determine the max distance the bot can see.  But first we check the special case
      'where the bot has not changed any of it's eye widths.  Sims generally have lots of veggies which
      'don't bother to do this, so this is worth it.
      eyesum = CLng(rob(n1).mem(531)) + _
               CLng(rob(n1).mem(532)) + _
               CLng(rob(n1).mem(533)) + _
               CLng(rob(n1).mem(534)) + _
               CLng(rob(n1).mem(535)) + _
               CLng(rob(n1).mem(536)) + _
               CLng(rob(n1).mem(537)) + _
               CLng(rob(n1).mem(538)) + _
               CLng(rob(n1).mem(539))
      If eyesum = 0 Then
         sightdist = 1440 * eyestrength(n1)
      Else
        sightdist = EyeSightDistance(NarrowestEye(n1), n1)
      End If
            
      'Now we check the maximum possible distance bot N1 can see against how far away bot N2 is.
      If edgetoedgedist > sightdist Then GoTo getout ' Bot too far away to see
      
      invdist = VectorInvMagnitude(ab)
        
      'ac and ad are to either end of the bots, while ab is to the center
      
      ac = VectorScalar(ab, invdist)
      'ac is now unit vector
      
      ad = VectorSet(ac.y, -ac.x)
      ad = VectorScalar(ad, rob(n2).radius)
      ad = VectorAdd(ab, ad)
      
      ac = VectorSet(-ac.y, ac.x)
      ac = VectorScalar(ac, rob(n2).radius)
      ac = VectorAdd(ab, ac)
            

      'Coordinates are in the 4th quadrant, so make the y values negative so the math works
      ad.y = -ad.y
      ac.y = -ac.y
      
      ' theta is the angle to the left edge of the viewed bot
      ' beta is the andgle to the right edge of the viewed bot
      
      If ad.x = 0 Then ' Divide by zero protection
        If ad.y > 0 Then
          theta = PI / 2 ' left edge of viewed bot is at 90 degrees
        Else
          theta = 3 * PI / 2 ' left edge of viewed bot is at 270 degrees
        End If
      Else
        If ad.x > 0 Then
          theta = Atn(ad.y / ad.x)
        Else
          theta = Atn(ad.y / ad.x) + PI
        End If
      End If
      
      If ac.x = 0 Then
        If ac.y > 0 Then
          beta = PI / 2
        Else
          beta = 3 * PI / 2
        End If
      Else
        If ac.x > 0 Then
          beta = Atn(ac.y / ac.x)
        Else
          beta = Atn(ac.y / ac.x) + PI
        End If
      End If
      
      'lets be sure to just deal with postive angles
      If theta < 0 Then theta = theta + 2 * PI
      If beta < 0 Then beta = beta + 2 * PI
      
      If beta > theta Then
        botspanszero = True
      Else
        botspanszero = False
      End If
      
      'For each eye
      For a = 0 To 8
                  
        'Now we check to see if the sight distance for this specific eye is far enough to see bot N2
        If rob(n1).mem(EYE1WIDTH + a) = 0 Then
          eyedist = 1440 * eyestrength(n1)
        Else
          eyedist = EyeSightDistance(AbsoluteEyeWidth(rob(n1).mem(EYE1WIDTH + a)), n1)
        End If
        If edgetoedgedist <= eyedist Then
      
          'Check to see if the bot is viewable in this eye
          'First, figure out the direction in radians in which the eye is pointed relative to .aim
          'We have to mod the value and divide by 200 to get radians
          'then since the eyedir values are offsets from their defaults, eye 1 is off from .aim by 4 eye field widths,
          'three for eye2, and so on.
          eyeaim = (rob(n1).mem(EYE1DIR + a) Mod 1256) / 200 - ((PI / 18) * a) + (PI / 18) * 4 + rob(n1).aim
        
          'It's possible we wrapped 0 so check
          While eyeaim > 2 * PI: eyeaim = eyeaim - 2 * PI: Wend
          While eyeaim < 0: eyeaim = eyeaim + 2 * PI: Wend
        
          'These are the left and right sides of the field of view for the eye
          halfeyewidth = ((rob(n1).mem(EYE1WIDTH + a)) Mod 1256) / 400
          While halfeyewidth > PI - PI / 36: halfeyewidth = halfeyewidth - PI: Wend
          While halfeyewidth < -PI / 36: halfeyewidth = halfeyewidth + PI: Wend
          eyeaimleft = eyeaim + halfeyewidth + PI / 36
          eyeaimright = eyeaim - halfeyewidth - PI / 36
        
          'Check the case where the eye field of view spans 0
          If eyeaimright < 0 Then eyeaimright = 2 * PI + eyeaimright
          If eyeaimleft > 2 * PI Then eyeaimleft = eyeaimleft - 2 * PI
          If eyeaimleft < eyeaimright Then
            eyespanszero = True
          Else
            eyespanszero = False
          End If
            
               
          ' Bot is visiable if either left edge is in eye or right edge is in eye or whole bot spans eye
          'If leftside of bot is in eye or
          '   rightside of bot is in eye or
          '   bot spans eye
          If ((eyeaimleft) >= (theta)) And ((theta) >= (eyeaimright)) And Not eyespanszero Or _
             ((eyeaimleft) >= (theta)) And eyespanszero Or _
             ((eyeaimright) <= (theta)) And eyespanszero Or _
             ((eyeaimleft) >= (beta)) And ((beta) >= (eyeaimright)) And Not eyespanszero Or _
             ((eyeaimleft) >= (beta)) And eyespanszero Or _
             ((eyeaimright) <= (beta)) And eyespanszero Or _
             ((eyeaimleft) <= (theta)) And ((beta) <= (eyeaimright)) And Not eyespanszero And Not botspanszero Or _
             ((eyeaimleft) <= (theta)) And Not eyespanszero And botspanszero Or _
             ((eyeaimright) >= (beta)) And Not eyespanszero And botspanszero Or _
             ((eyeaimleft) <= (theta)) And (eyeaimright >= beta) And eyespanszero And botspanszero Then
            'The bot is viewable in this eye.
                      
            'Calculate the eyevalue
            If edgetoedgedist <= 0 Then ' bots overlap
               eyevalue = 32000
            Else
               percentdist = (edgetoedgedist + 10) / eyedist
               eyevalue = 1 / (percentdist * percentdist)
               If eyevalue > 32000 Then eyevalue = 32000
            End If
            
            'Check to see if it is closer than other bots we may have seen
            If rob(n1).mem(EyeStart + 1 + a) < eyevalue Then
              'It is closer than other bots we may have seen.
              'Check to see if this eye has the focus
              If a = Abs(rob(n1).mem(FOCUSEYE) + 4) Mod 9 Then
                'This eye does have the focus
                'Set the EYEF value and also lastopp so the lookoccur list will get populated later
                rob(n1).lastopp = n2
                rob(n1).mem(EYEF) = eyevalue
              End If
              'Set the distance for the eye
              rob(n1).mem(EyeStart + 1 + a) = eyevalue
             ' If n1 = robfocus Then eyeDistance(a + 1) = edgetoedgedist + rob(n1).radius
            End If
          End If
        End If
      Next a
getout:
    End Sub

'Returns the percent along vector P0 + sDO where it interects vector P1 + tD1.
'Returns 0 if there is no interestion
Public Function SegmentSegmentIntersect(P0 As Vector, D0 As Vector, P1 As Vector, D1 As Vector) As Single
Dim dotPerp As Single
Dim Delta As Vector
Dim s As Single
Dim t As Single

  SegmentSegmentIntersect = 0
  dotPerp = D0.x * D1.y - D1.x * D0.y ' Test for intersection
        
  If dotPerp <> 0 Then
    Delta = VectorSub(P1, P0)
    s = Dot(Delta, VectorSet(D1.y, -D1.x)) / dotPerp
    t = Dot(Delta, VectorSet(D0.y, -D0.x)) / dotPerp
    If s >= 0 And s <= 1 And t >= 0 And t <= 1 Then SegmentSegmentIntersect = s
  End If
        
End Function
