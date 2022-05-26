Attribute VB_Name = "NeoMutations"
Option Explicit

'1-(perbot+1)^(1/DNALength) = per unit
'1-(1-perunit)^DNALength = perbot

Public Const PointUP As Integer = 0 'expressed as 1 chance in X per kilocycle per bp
Public Const MinorDeletionUP As Integer = 1
Public Const ReversalUP As Integer = 2
Public Const InsertionUP As Integer = 3
Public Const AmplificationUP As Integer = 4
Public Const MajorDeletionUP As Integer = 5
Public Const CopyErrorUP As Integer = 6
Public Const DeltaUP As Integer = 7
Public Const TranslocationUP As Integer = 8
Public Const P2UP As Integer = 9 'Botsareus 12/10/2013 new mutation rates
Public Const CE2UP As Integer = 10

Private Const overtime As Double = 30 'Time correction across all mutations

Sub logmutation(ByVal n As Integer, ByVal strmut As String) 'Botsareus 10/8/2015 Wrap mutations to prevent crash
If SimOpts.TotRunCycle = 0 Then Exit Sub 'Botsareus 4/28/2016 Prevents div/0
 With rob(n)
  If Len(.LastMutDetail) > (100000000 / TotalRobotsDisplayed) Then .LastMutDetail = "" 'Botsareus 4/11/2016 Bug fix - Use the total string length across all robots
  .LastMutDetail = strmut + vbCrLf + .LastMutDetail
 End With
End Sub

Private Function MutationType(thing As Integer) As String
  MutationType = ""
  Select Case thing
    Case 0
      MutationType = "Point Mutation"
    Case 1
      MutationType = "Minor Deletion"
    Case 2
      MutationType = "Reversal"
    Case 3
      MutationType = "Insertion"
    Case 4
      MutationType = "Amplification"
    Case 5
      MutationType = "Major Deletion"
    Case 6
      MutationType = "Copy Error"
    Case 7
      MutationType = "Delta Mutation"
  End Select
  
End Function


'NEVER allow anything after end, which must be = DNALen
'ALWAYS assume that DNA is sized right
'ALWAYS size DNA correctly when mutating

Private Function EraseUnit(ByRef unit As block)
  unit.tipo = -1
  unit.value = -1
End Function

Public Function MakeSpace(ByRef dna() As block, ByVal beginning As Long, ByVal length As Long, Optional DNALength As Integer = -1) As Boolean
  'add length elements after beginning.  Beginning doesn't move places
  'returns true if the space was created,
  'false otherwise

  Dim t As Integer
  If DNALength < 0 Then DNALength = DnaLen(dna)
  If length < 1 Or beginning < 0 Or beginning > DNALength - 1 Or (DNALength + length > 32000) Then
    MakeSpace = False
    GoTo getout
  End If
  
  MakeSpace = True

  ReDim Preserve dna(DNALength + length)
  
  'Botsareus 6/25/2016 Bugfix - erase dna that was just created to insure all units get erased
  For t = DNALength + 1 To DNALength + length Step 1
    EraseUnit dna(t)
  Next

  For t = DNALength To beginning + 1 Step -1
    dna(t + length) = dna(t)
    EraseUnit dna(t)
  Next t
getout:
End Function

Public Sub Delete(ByRef dna() As block, ByRef beginning As Long, ByRef elements As Long, Optional DNALength As Integer = -1)
  'delete elements starting at beginning
  Dim t As Integer
  If DNALength < 0 Then DNALength = DnaLen(dna)
  If elements < 1 Or beginning < 1 Or beginning > DNALength - 1 Then GoTo getout
 ' If elements + beginning > DNALength - 1 Then elements = DNALength - 1 - beginning

  For t = beginning + elements To DNALength
    On Error GoTo step2 'small error mod
    dna(t - elements) = dna(t)
  Next t

step2:
  DNALength = DnaLen(dna)
  ReDim Preserve dna(DNALength)
getout:
End Sub

Public Function NewSubSpecies(n As Integer) As Integer
Dim i As Integer

  i = SpeciesFromBot(n)  ' Get the index into the species array for this bot
  SimOpts.Specie(i).SubSpeciesCounter = SimOpts.Specie(i).SubSpeciesCounter + 1 ' increment the counter
  If SimOpts.Specie(i).SubSpeciesCounter > 32000 Then SimOpts.Specie(i).SubSpeciesCounter = -32000 'wrap the counter if necessary
  NewSubSpecies = SimOpts.Specie(i).SubSpeciesCounter

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub mutate(ByVal robn As Integer, Optional reproducing As Boolean = False) 'Botsareus 12/17/2013
  Dim Delta As Long

  With rob(robn)
    If Not .Mutables.Mutations Or SimOpts.DisableMutations Then GoTo getout
    Delta = CLng(.LastMut)
    
    
    ismutating = True 'Botsareus 2/2/2013 Tells the parseor to ignore debugint and debugbool while the robot is mutating
    If Not reproducing Then
        
      If .Mutables.mutarray(PointUP) > 0 Then PointMutation robn
      If .Mutables.mutarray(DeltaUP) > 0 And Not Delta2 Then DeltaMut robn
      If .Mutables.mutarray(P2UP) > 0 And sunbelt Then PointMutation2 robn
      
      'special case update epigenetic reset
      If CLng(.LastMut) - Delta > 0 And epireset Then .MutEpiReset = .MutEpiReset + (CLng(.LastMut) - Delta) ^ epiresetemp
      
      'Delta2 point mutation change
      If Delta2 And DeltaPM > 0 Then
       If .age Mod DeltaPM = 0 And .age > 0 Then
        Dim MratesMax As Long
        MratesMax = IIf(NormMut, CLng(.DnaLen) * CLng(valMaxNormMut), 2000000000)
        Dim t As Byte
        For t = 0 To 9 Step 9 'Point and Point2
        If .Mutables.mutarray(t) < 1 Then GoTo skip 'Botsareus 1/3/2014 if mutation off then skip it
         If rndy < DeltaMainChance / 100 Then
          If DeltaMainExp <> 0 Then .Mutables.mutarray(t) = .Mutables.mutarray(t) * 10 ^ ((rndy * 2 - 1) / DeltaMainExp)
          .Mutables.mutarray(t) = .Mutables.mutarray(t) + (rndy * 2 - 1) * DeltaMainLn
          If .Mutables.mutarray(t) < 1 Then .Mutables.mutarray(t) = 1
          If .Mutables.mutarray(t) > MratesMax Then .Mutables.mutarray(t) = MratesMax
         End If
         If rndy < DeltaDevChance / 100 Then
          If DeltaDevExp <> 0 Then .Mutables.StdDev(t) = .Mutables.StdDev(t) * 10 ^ ((rndy * 2 - 1) / DeltaDevExp)
          .Mutables.StdDev(t) = .Mutables.StdDev(t) + (rndy * 2 - 1) * DeltaDevLn
          If DeltaDevExp <> 0 Then .Mutables.Mean(t) = .Mutables.Mean(t) * 10 ^ ((rndy * 2 - 1) / DeltaDevExp)
          .Mutables.Mean(t) = .Mutables.Mean(t) + (rndy * 2 - 1) * DeltaDevLn
          'Max range is always 0 to 800
          If .Mutables.StdDev(t) < 0 Then .Mutables.StdDev(t) = 0
          If .Mutables.StdDev(t) > 200 Then .Mutables.StdDev(t) = 200
          If .Mutables.Mean(t) < 1 Then .Mutables.Mean(t) = 1
          If .Mutables.Mean(t) > 400 Then .Mutables.Mean(t) = 400
         End If
skip:
        Next
        .Mutables.PointWhatToChange = .Mutables.PointWhatToChange + (rndy * 2 - 1) * DeltaWTC
        If .Mutables.PointWhatToChange < 0 Then .Mutables.PointWhatToChange = 0
        If .Mutables.PointWhatToChange > 100 Then .Mutables.PointWhatToChange = 100
        .Point2MutCycle = 0
        .PointMutCycle = 0
       End If
      End If
      
    Else
    
      If .Mutables.mutarray(CopyErrorUP) > 0 Then CopyError robn
      If .Mutables.mutarray(CE2UP) > 0 And sunbelt Then CopyError2 robn
      If .Mutables.mutarray(InsertionUP) > 0 Then Insertion robn
      If .Mutables.mutarray(ReversalUP) > 0 Then Reversal robn
      If .Mutables.mutarray(TranslocationUP) > 0 And sunbelt Then Translocation robn 'Botsareus Translocation and Amplification still bugy, but I want them.
      If .Mutables.mutarray(AmplificationUP) > 0 And sunbelt Then Amplification robn
      If .Mutables.mutarray(MajorDeletionUP) > 0 Then MajorDeletion robn
      If .Mutables.mutarray(MinorDeletionUP) > 0 Then MinorDeletion robn
      
    End If

    ismutating = False 'Botsareus 2/2/2013 Tells the parseor to ignore debugint and debugbool while the robot is mutating
        
    Delta = CLng(.LastMut) - Delta 'Botsareus 9/4/2012 Moved delta check before overflow reset to fix an error where robot info is not being updated
  
    'auto forking
    If SimOpts.EnableAutoSpeciation Then
        If CDbl(.Mutations) > CDbl(.DnaLen) * CDbl(SimOpts.SpeciationGeneticDistance / 100) Then
                Dim robname As String
                Dim splitname() As String
                'generate new specie name
                SimOpts.SpeciationForkInterval = SimOpts.SpeciationForkInterval + 1
                'remove old nick name
                splitname = Split(.FName, ")")
                'if it is a nick name only
                If Left(splitname(0), 1) = "(" And IsNumeric(Right(splitname(0), Len(splitname(0)) - 1)) Then
                    robname = splitname(1)
                Else
                    robname = .FName
                End If
                    robname = "(" & SimOpts.SpeciationForkInterval & ")" & robname
                'do we have room for new specie?
                If SimOpts.SpeciesNum < 49 Then
                    .FName = robname
                    .Mutations = 0
                    AddSpecie robn, False
                Else
                    SimOpts.SpeciationForkInterval = SimOpts.SpeciationForkInterval - 1
                End If
        End If
    End If
  
    If .Mutations > 32000 Then .Mutations = 32000  'Botsareus 5/31/2012 Prevents mutations overflow
    If .LastMut > 32000 Then .LastMut = 32000
  
    If (Delta > 0) Then  'The bot has mutated.
    
       .GenMut = .GenMut - .LastMut
       If .GenMut < 0 Then .GenMut = 0
      
      mutatecolors robn, Delta
      .SubSpecies = NewSubSpecies(robn)
      .genenum = CountGenes(rob(robn).dna())
      .DnaLen = DnaLen(rob(robn).dna())
      .mem(DnaLenSys) = .DnaLen
      .mem(GenesSys) = .genenum
    End If
getout:
  End With
End Sub

Private Sub Amplification(robn As Integer) 'Botsareus 12/10/2013
On Error GoTo getout:
  '1. pick a spot (1 to .dnalen - 1)
  '2. Run a length, copied to a temporary location
  '3. Pick a new spot (1 to .dnalen - 1)
  '4. Insert copied DNA
      
  Dim t As Long
  Dim length As Long
  With rob(robn)
  
 Dim floor As Double
 floor = CDbl(.DnaLen) * CDbl(.Mutables.Mean(AmplificationUP) + .Mutables.StdDev(AmplificationUP)) / (1200 * overtime)
 floor = floor * SimOpts.MutCurrMult
 If .Mutables.mutarray(AmplificationUP) < floor Then .Mutables.mutarray(AmplificationUP) = floor 'Botsareus 10/5/2015 Prevent freezing
  
  Dim tempDNA() As block
  Dim start As Long
  Dim second As Long
  Dim counter As Long
  t = 1
  Do
  t = t + 1
    If rndy < 1 / (.Mutables.mutarray(AmplificationUP) / SimOpts.MutCurrMult) Then
      length = Gauss(.Mutables.StdDev(AmplificationUP), .Mutables.Mean(AmplificationUP))
      length = length Mod UBound(.dna)
      If length < 1 Then length = 1
      
      length = length - 1
      length = length \ 2
      If t - length < 1 Then GoTo skip
      If t + length > .DnaLen - 1 Then GoTo skip
      If UBound(.dna) + CLng(length) * 2 > 32000 Then GoTo skip 'Botsareus 10/5/2015 Size limit is calculated here
      
      If length > 0 Then
      
        ReDim tempDNA(length * 2)
        
        second = 0
        For counter = t - length To t + length
          tempDNA(second) = .dna(counter)
          second = second + 1
        Next counter
        'we now have the appropriate length of DNA in the temporary array.

        start = Random(1, UBound(.dna) - 2)
        MakeSpace .dna(), start, UBound(tempDNA) + 1

        For counter = start + 1 To start + UBound(tempDNA) + 1
         .dna(counter) = tempDNA(counter - start - 1)
        Next counter
             
        'BOTSAREUSIFIED
        .Mutations = .Mutations + 1
        .LastMut = .LastMut + 1
        logmutation robn, "Amplification copied a series at" + Str(t) + Str(length * 2 + 1) + "bps long to " + Str(start) + " during cycle" + _
          Str(SimOpts.TotRunCycle)
          
      End If
    End If
skip:
  Loop Until t >= UBound(.dna) - 1
  
        'add "end" to end of the DNA
        .dna(UBound(.dna)).tipo = 10
        .dna(UBound(.dna)).value = 1
  
  End With
getout:
rob(robn).DnaLen = DnaLen(rob(robn).dna()) 'Botsareus 10/6/2015 Calculate new dna size
End Sub


Private Sub Translocation(robn As Integer) 'Botsareus 12/10/2013
On Error GoTo getout:
  '1. pick a spot (1 to .dnalen - 1)
  '2. Run a length, copied to a temporary location
  '3.  Pick a new spot (1 to .dnalen - 1)
  '4. Insert copied DNA
  
  Dim t As Long
  Dim length As Long
  With rob(robn)
  
 Dim floor As Double
 floor = CDbl(.DnaLen) * CDbl(.Mutables.Mean(TranslocationUP) + .Mutables.StdDev(TranslocationUP)) / (360 * overtime)
 floor = floor * SimOpts.MutCurrMult
 If .Mutables.mutarray(TranslocationUP) < floor Then .Mutables.mutarray(TranslocationUP) = floor  'Botsareus 10/5/2015 Prevent freezing
  
  Dim tempDNA() As block
  Dim start As Long
  Dim second As Long
  Dim counter As Long
  
  For t = 1 To UBound(.dna) - 1
    If rndy < 1 / (.Mutables.mutarray(TranslocationUP) / SimOpts.MutCurrMult) Then

      length = Gauss(.Mutables.StdDev(TranslocationUP), .Mutables.Mean(TranslocationUP))
      length = length Mod UBound(.dna)
      If length < 1 Then length = 1
      
      length = length - 1
      length = length \ 2
      If t - length < 1 Then GoTo skip
      If t + length > UBound(.dna) - 1 Then GoTo skip
      
      If length > 0 Then
      
        ReDim tempDNA(length * 2)
        
        second = 0
        For counter = t - length To t + length
          tempDNA(second) = .dna(counter)
          second = second + 1
        Next counter
        'we now have the appropriate length of DNA in the temporary array.
        
        'delete fragment
        Delete .dna, t - length, length * 2 + 1 'Botsareus 12/11/2015 Bug fix

        'open up a hole
        start = Random(1, UBound(.dna) - 2)
        MakeSpace .dna(), start, UBound(tempDNA) + 1 'Botsareus 12/11/2015 Bug fix

        For counter = start + 1 To start + UBound(tempDNA) + 1
         .dna(counter) = tempDNA(counter - start - 1)
        Next counter
             
        'BOTSAREUSIFIED
        .Mutations = .Mutations + 1
        .LastMut = .LastMut + 1
        logmutation robn, "Translocation moved a series at" + Str(t) + Str(length * 2 + 1) + "bps long to " + Str(start) + " during cycle" + _
          Str(SimOpts.TotRunCycle)
          
          
       
      End If
    End If
    
skip:
  Next t
  
        'add "end" to end of the DNA
        .dna(UBound(.dna)).tipo = 10
        .dna(UBound(.dna)).value = 1
  
  End With
getout:
End Sub

Private Sub CopyError2(robn As Integer) 'Just like Copyerror but effects only special chars
Dim DNAsize As Integer
Dim e As Integer 'counter
Dim e2 As Integer 'update generator (our position)

With rob(robn)

 Dim floor As Double
 floor = CDbl(.DnaLen) * CDbl(.Mutables.Mean(CopyErrorUP) + .Mutables.StdDev(CopyErrorUP)) / (5 * overtime) 'Botsareus 3/22/2016 works like p2 now
 floor = floor * SimOpts.MutCurrMult
 If .Mutables.mutarray(CE2UP) < floor Then .Mutables.mutarray(CE2UP) = floor 'Botsareus 10/5/2015 Prevent freezing

    DNAsize = DnaLen(.dna) - 1 'get aprox length
    
    Dim datahit() As Boolean 'operation repeat prevention
    ReDim datahit(DNAsize)
    For e = 1 To DNAsize 'Botsareus 3/22/2016 Bugfix
    
    'Botsareus 3/22/2016 keeps CopyError2 lengths the same as CopyError
    Dim calc_gauss As Double
    calc_gauss = Gauss(.Mutables.StdDev(CopyErrorUP), .Mutables.Mean(CopyErrorUP))
    If calc_gauss < 1 Then calc_gauss = 1
    
        If rndy < (0.75 / (.Mutables.mutarray(CE2UP) / (SimOpts.MutCurrMult * calc_gauss))) Then   'chance 'Botsareus 3/22/2016 works like p2 now
            Do
                e2 = Int(rndy * DNAsize) + 1 'Botsareus 3/22/2016 Bugfix
            Loop Until datahit(e2) = False
            datahit(e2) = True
            
            ChangeDNA2 robn, e2, DNAsize  'Botsareus 4/10/2016 Less boilerplate code
            
        End If
    Next
End With
End Sub

Private Sub PointMutation2(robn As Integer) 'Botsareus 12/10/2013
  'assume the bot has a positive (>0) mutarray value for this
  
  Dim DNAsize As Integer
  Dim randompos As Integer
   
  With rob(robn)
  
 Dim floor As Double
 floor = CDbl(.DnaLen) * CDbl(.Mutables.Mean(PointUP) + .Mutables.StdDev(PointUP)) / (400 * overtime)
 floor = floor * SimOpts.MutCurrMult
 If .Mutables.mutarray(P2UP) < floor Then .Mutables.mutarray(P2UP) = floor 'Botsareus 10/5/2015 Prevent freezing

  
    If .age = 0 Or .Point2MutCycle < .age Then Point2MutWhen rndy, robn
    
    'Do it again in case we get two point mutations in a single cycle
    While .age = .Point2MutCycle And .age > 0 And .DnaLen > 1 ' Avoid endless loop when .age = 0 and/or .DNALen = 1
        
        'sysvar mutation
        DNAsize = DnaLen(.dna) - 1 'get aprox length
        randompos = Int(rndy * DNAsize) + 1 'Botsareus 3/22/2016 Bug fix
            
        ChangeDNA2 robn, randompos, DNAsize, True  'Botsareus 4/10/2016 Less boilerplate code
        
      Point2MutWhen rndy, robn
    Wend
  End With
End Sub

Private Sub PointMutation(robn As Integer)
  'assume the bot has a positive (>0) mutarray value for this
  
  Dim temp As Single
  Dim temp2 As Long
 
  With rob(robn)
  
 Dim floor As Double
 floor = CDbl(.DnaLen) * CDbl(.Mutables.Mean(PointUP) + .Mutables.StdDev(PointUP)) / (400 * overtime)
 floor = floor * SimOpts.MutCurrMult
 If .Mutables.mutarray(PointUP) < floor Then .Mutables.mutarray(PointUP) = floor 'Botsareus 10/5/2015 Prevent freezing


  
    If .age = 0 Or .PointMutCycle < .age Then PointMutWhereAndWhen rndy, robn, .PointMutBP
    
    'Do it again in case we get two point mutations in a single cycle
    While .age = .PointMutCycle And .age > 0 And .DnaLen > 1 ' Avoid endless loop when .age = 0 and/or .DNALen = 1
      temp = Gauss(.Mutables.StdDev(PointUP), .Mutables.Mean(PointUP))
      temp2 = Int(temp) Mod 32000        '<- Overflow was here when huge single is assigned to a Long
      ChangeDNA robn, .PointMutBP, temp2, .Mutables.PointWhatToChange
      PointMutWhereAndWhen rndy, robn, .PointMutBP
    Wend
  End With
End Sub


Private Sub Point2MutWhen(randval As Single, robn As Integer)
  Dim result As Single
  
  Dim mutation_rate As Single
 
  'If randval = 0 Then randval = 0.0001
  With rob(robn)
    If .DnaLen = 1 Then GoTo getout ' avoid divide by 0 below
    
    mutation_rate = .Mutables.mutarray(P2UP) / SimOpts.MutCurrMult
    
    'keeps Point2 lengths the same as Point Botsareus 1/14/2014 Checking to make sure value is >= 1
    Dim calc_gauss As Double
    calc_gauss = Gauss(.Mutables.StdDev(PointUP), .Mutables.Mean(PointUP))
    If calc_gauss < 1 Then calc_gauss = 1
    
    mutation_rate = mutation_rate / calc_gauss
    
    mutation_rate = mutation_rate * 1.33 'Botsareus 4/19/2016 Adjust because changedna2 may write 2 commands
    
    'Here we test to make sure the probability of a point mutation isn't crazy high.
    'A value of 1 is the probability of mutating every base pair every 1000 cycles
    'Lets not let it get lower than 1 shall we?
    If mutation_rate < 1 And mutation_rate > 0 Then
      mutation_rate = 1
    End If
  
    'result = offset + Fix(Log(randval) / Log(1 - 1 / (1000 * .Mutables.mutarray(PointUP))))
    result = Log(1 - randval) / Log(1 - 1 / (1000 * mutation_rate))
    While result > 1800000000: result = result - 1800000000: Wend 'Botsareus 3/15/2013 overflow fix
    .Point2MutCycle = .age + result / (.DnaLen - 1)
getout:
  End With
End Sub


Private Sub PointMutWhereAndWhen(randval As Single, robn As Integer, Optional offset As Long = 0)
  Dim result As Single
  
  Dim mutation_rate As Single
  
  'If randval = 0 Then randval = 0.0001
  With rob(robn)
    If .DnaLen = 1 Then GoTo getout ' avoid divide by 0 below
    
    mutation_rate = .Mutables.mutarray(PointUP) / SimOpts.MutCurrMult
    
    'Here we test to make sure the probability of a point mutation isn't crazy high.
    'A value of 1 is the probability of mutating every base pair every 1000 cycles
    'Lets not let it get lower than 1 shall we?
    If mutation_rate < 1 And mutation_rate > 0 Then
      mutation_rate = 1
    End If
  
    'result = offset + Fix(Log(randval) / Log(1 - 1 / (1000 * .Mutables.mutarray(PointUP))))
    result = Log(1 - randval) / Log(1 - 1 / (1000 * mutation_rate))
    While result > 1800000000: result = result - 1800000000: Wend 'Botsareus 3/15/2013 overflow fix
    .PointMutBP = (result Mod (.DnaLen - 1)) + 1 'note that DNA(DNALen) = end.
    'We don't mutate end.  Also note that DNA does NOT start at 0th element
    .PointMutCycle = .age + result / (.DnaLen - 1)
getout:
  End With
End Sub

Private Sub DeltaMut(robn As Integer)
  Dim temp As Integer
  Dim newval As Single ' EricL Made newval Single instead of Long.
    
  With rob(robn)

  If rndy > 1 - 1 / (100 * .Mutables.mutarray(DeltaUP) / SimOpts.MutCurrMult) Then
    If .Mutables.StdDev(DeltaUP) = 0 Then .Mutables.Mean(DeltaUP) = 50
    If .Mutables.Mean(DeltaUP) = 0 Then .Mutables.Mean(DeltaUP) = 25
    
    'temp = Random(0, 20)
    Do
      temp = Random(0, 10) 'Botsareus 12/14/2013 Added new mutations
    Loop While .Mutables.mutarray(temp) <= 0
    
    Do
      newval = Gauss(.Mutables.Mean(DeltaUP), .Mutables.mutarray(temp))
    Loop While .Mutables.mutarray(temp) = newval Or newval <= 0
    
    logmutation robn, "Delta mutations changed " + MutationType(temp) + " from 1 in" + Str(.Mutables.mutarray(temp)) + _
      " to 1 in" + Str(newval)
    .Mutations = .Mutations + 1
    .LastMut = .LastMut + 1
    .Mutables.mutarray(temp) = newval
  End If
  
  End With
End Sub

Private Sub CopyError(robn As Integer)
  Dim t As Long
  Dim length As Long
  
  With rob(robn)
  
 Dim floor As Double
 floor = CDbl(.DnaLen) * CDbl(.Mutables.Mean(CopyErrorUP) + .Mutables.StdDev(CopyErrorUP)) / (25 * overtime)
 floor = floor * SimOpts.MutCurrMult
 If .Mutables.mutarray(CopyErrorUP) < floor Then .Mutables.mutarray(CopyErrorUP) = floor 'Botsareus 10/5/2015 Prevent freezing
  
  For t = 1 To (.DnaLen - 1) 'note that DNA(.dnalen) = end, and we DON'T mutate that.
   
    If rndy < 1 / (rob(robn).Mutables.mutarray(CopyErrorUP) / SimOpts.MutCurrMult) Then
      length = Gauss(rob(robn).Mutables.StdDev(CopyErrorUP), _
        rob(robn).Mutables.Mean(CopyErrorUP)) 'length
        
      ChangeDNA robn, t, length, rob(robn).Mutables.CopyErrorWhatToChange, _
        CopyErrorUP
    End If
  Next t
  
  End With
End Sub

Private Sub ChangeDNA2(robn As Integer, ByVal nth As Integer, ByVal DNAsize As Integer, Optional IsPoint As Boolean = False)
Dim randomsysvar As Integer
Dim holddetail As String
Dim special As Boolean
With rob(robn)

    'for .tieloc, .shoot, and functional
    Do
        randomsysvar = Int(rndy * 256)
    Loop Until sysvarOUT(randomsysvar).Name <> ""

    special = False
    'special cases
    If nth < DNAsize - 2 Then
        'for .shoot store
        If .dna(nth + 1).tipo = 0 And .dna(nth + 1).value = shoot _
        And .dna(nth + 2).tipo = 7 And .dna(nth + 2).value = 1 Then
            .dna(nth).value = Choose(Int(rndy * 7) + 1, -1, -2, -3, -4, -6, -8, sysvar(randomsysvar).value) 'Botsareus 10/6/2015 Better values
            .dna(nth).tipo = 0
            holddetail = " changed dna location " & nth & " to " & .dna(nth).value
            special = True
        End If
        'for .focuseye store
        If .dna(nth + 1).tipo = 0 And .dna(nth + 1).value = FOCUSEYE _
        And .dna(nth + 2).tipo = 7 And .dna(nth + 2).value = 1 Then
            .dna(nth).value = Int(rndy * 9) - 4
            .dna(nth).tipo = 0
            holddetail = " changed dna location " & nth & " to " & .dna(nth).value
            special = True
        End If
        'for .tieloc store
        If .dna(nth + 1).tipo = 0 And .dna(nth + 1).value = tieloc _
        And .dna(nth + 2).tipo = 7 And .dna(nth + 2).value = 1 Then
            .dna(nth).value = Choose(Int(rndy * 5) + 1, -1, -3, -4, -6, sysvar(randomsysvar).value) 'Botsareus 10/6/2015 Better values 'Botsareus 3/22/2016 Better values
            .dna(nth).tipo = 0
            holddetail = " changed dna location " & nth & " to " & .dna(nth).value
            special = True
        End If
    End If
    
    If special Then
            logmutation robn, IIf(IsPoint, "Point Mutation 2", "Copy Error 2") & holddetail & " during cycle" & Str(SimOpts.TotRunCycle)
            .Mutations = .Mutations + 1
            .LastMut = .LastMut + 1
    Else 'other cases
        If nth < DNAsize - 1 And Int(rndy * 3) = 0 Then '1/3 chance functional
        
            .dna(nth).tipo = 0
            .dna(nth).value = sysvarOUT(randomsysvar).value
            holddetail = " changed dna location " & nth & " to number ." & sysvarOUT(randomsysvar).Name
            logmutation robn, IIf(IsPoint, "Point Mutation 2", "Copy Error 2") & holddetail & " during cycle" & Str(SimOpts.TotRunCycle)
            .Mutations = .Mutations + 1
            .LastMut = .LastMut + 1
            
            .dna(nth + 1).tipo = 7
            .dna(nth + 1).value = 1
            holddetail = " changed dna location " & (nth + 1) & " to store"
            logmutation robn, IIf(IsPoint, "Point Mutation 2", "Copy Error 2") & holddetail & " during cycle" & Str(SimOpts.TotRunCycle)
            .Mutations = .Mutations + 1
            .LastMut = .LastMut + 1
            
        Else '2/3 chance informational
            If Int(rndy * 5) = 0 Then '1/5 chance large number (but still use a sysvar, if anything the parse will mod it)
                Do
                    randomsysvar = Int(rndy * 1000)
                Loop Until sysvar(randomsysvar).Name <> ""
                .dna(nth).tipo = 0
                .dna(nth).value = sysvar(randomsysvar).value + Int(rndy * 32) * 1000
                holddetail = " changed dna location " & nth & " to number " & .dna(nth).value
            Else
                Do
                    randomsysvar = Int(rndy * 256)
                Loop Until sysvarIN(randomsysvar).Name <> ""
                .dna(nth).tipo = 1
                .dna(nth).value = sysvarIN(randomsysvar).value
                holddetail = " changed dna location " & nth & " to *number *." & sysvarIN(randomsysvar).Name
            End If
            logmutation robn, IIf(IsPoint, "Point Mutation 2", "Copy Error 2") & holddetail & " during cycle" & Str(SimOpts.TotRunCycle)
            .Mutations = .Mutations + 1
            .LastMut = .LastMut + 1
        End If
    End If
End With
End Sub

Private Sub ChangeDNA(robn As Integer, ByVal nth As Long, Optional ByVal length As Long = 1, Optional ByVal PointWhatToChange As Integer = 50, Optional Mtype As Integer = PointUP)

  'we need to rework .lastmutdetail
  Dim Max As Long
  Dim temp As String
  Dim bp As block
  Dim tempbp As block
  Dim Name As String
  Dim oldname As String
  Dim t As Long
  Dim old As Long
  
  With rob(robn)
     
  For t = nth To (nth + length - 1) 'if length is 1, it's only one bp we're mutating, remember?
    If t >= .DnaLen Then GoTo getout 'don't mutate end either
    If .dna(t).tipo = 10 Then GoTo getout 'mutations can't cross control barriers
    
    If Random(0, 99) < PointWhatToChange Then
      '''''''''''''''''''''''''''''''''''''''''
      'Mutate VALUE
      '''''''''''''''''''''''''''''''''''''''''
      If .dna(t).value And Mtype = InsertionUP Then
        'Insertion mutations should get a good range for value.
        'Don't worry, this will get "mod"ed for non number commands.
        'This doesn't count as a mutation, since the whole:
        ' -- Add an element, set it's tipo and value to random stuff -- is a SINGLE mutation
        'we'll increment mutation counters and .lastmutdetail later.
        .dna(t).value = Gauss(500, 0) 'generates values roughly between -1000 and 1000
      End If
      

      old = .dna(t).value
      If .dna(t).tipo = 0 Or .dna(t).tipo = 1 Then '(number or *number)
        Do
            If Abs(old) <= 1000 Then   'Botsareus 3/19/2016 Simplified
                If Int(rndy * 2) = 0 Then  '1/2 chance the mutation is large
                    .dna(t).value = Gauss(94, .dna(t).value)
                Else
                    .dna(t).value = Gauss(7, .dna(t).value)
                End If
            Else
                .dna(t).value = Gauss(old / 10, .dna(t).value) 'for very large numbers scale gauss
            End If
        Loop While .dna(t).value = old

        .Mutations = .Mutations + 1
        .LastMut = .LastMut + 1
        
        logmutation robn, MutationType(Mtype) + " changed " + TipoDetok(.dna(t).tipo) + " from" + Str(old) + " to" + Str(.dna(t).value) + " at position" + Str(t) + " during cycle" + Str(SimOpts.TotRunCycle)
        
      Else
        'find max legit value
        'this should really be done a better way
        bp.tipo = .dna(t).tipo
        Max = 0
        Do
          temp = ""
          Max = Max + 1
          bp.value = Max
          Parse temp, bp
        Loop While temp <> ""
        Max = Max - 1
        If Max <= 1 Then GoTo getout 'failsafe in case its an invalid type or there's no way to mutate it
        
        Do
          .dna(t).value = Random(1, Max)
        Loop While .dna(t).value = old

        bp.tipo = .dna(t).tipo
        bp.value = old
        
        tempbp = .dna(t)
        
        Name = ""
        oldname = ""
        Parse Name, tempbp    ' Have to use a temp var because Parse() can change the arguments
        Parse oldname, bp
        
        .Mutations = .Mutations + 1
        .LastMut = .LastMut + 1
        
        logmutation robn, MutationType(Mtype) + " changed value of " + TipoDetok(.dna(t).tipo) + " from " + _
          oldname + " to " + Name + " at position" + Str(t) + " during cycle" + _
          Str(SimOpts.TotRunCycle)
      End If
    Else
      bp.tipo = .dna(t).tipo
      bp.value = .dna(t).value
      Do
        .dna(t).tipo = Random(0, 20)
      Loop While .dna(t).tipo = bp.tipo Or TipoDetok(.dna(t).tipo) = ""
      Max = 0
      If .dna(t).tipo >= 2 Then
        Do
          temp = ""
          Max = Max + 1
          .dna(t).value = Max
          Parse temp, .dna(t)
        Loop While temp <> ""
        Max = Max - 1
        If Max <= 1 Then GoTo getout 'failsafe in case its an invalid type or there's no way to mutate it
        .dna(t).value = ((Abs(bp.value) - 1) Mod Max) + 1 'put values in range 'Botsareus 4/10/2016 Bug fix
        If .dna(t).value = 0 Then .dna(t).value = 1
      Else
        'we do nothing, it has to be in range
      End If
       tempbp = .dna(t)
       
       Name = ""
       oldname = ""
       Parse Name, tempbp ' Have to use a temp var because Parse() can change the arguments
       Parse oldname, bp
      .Mutations = .Mutations + 1
      .LastMut = .LastMut + 1
      
      logmutation robn, MutationType(Mtype) + " changed the " + TipoDetok(bp.tipo) + ": " + _
          oldname + " to the " + TipoDetok(.dna(t).tipo) + ": " + Name + " at position" + Str(t) + " during cycle" + _
          Str(SimOpts.TotRunCycle)
      
    End If
  Next t
getout:
  End With
End Sub

Private Sub Insertion(robn As Integer)
  Dim location As Integer
  Dim length As Integer
  Dim accum As Long
  Dim t As Long
  
  With rob(robn)
  
    Dim floor As Double
    floor = CDbl(.DnaLen) * CDbl(.Mutables.Mean(InsertionUP) + .Mutables.StdDev(InsertionUP)) / (5 * overtime)
    floor = floor * SimOpts.MutCurrMult
    If .Mutables.mutarray(InsertionUP) < floor Then .Mutables.mutarray(InsertionUP) = floor  'Botsareus 10/5/2015 Prevent freezing
    
    For t = 1 To (.DnaLen - 1)
      If rndy < 1 / (.Mutables.mutarray(InsertionUP) / SimOpts.MutCurrMult) Then
        If .Mutables.Mean(InsertionUP) = 0 Then .Mutables.Mean(InsertionUP) = 1
        Do
          length = Gauss(.Mutables.StdDev(InsertionUP), .Mutables.Mean(InsertionUP))
        Loop While length <= 0
        
        If CLng(rob(robn).DnaLen) + CLng(length) > 32000 Then Exit Sub
        
        
        MakeSpace .dna(), t + accum, length, .DnaLen
        rob(robn).DnaLen = rob(robn).DnaLen + length
        ChangeDNA robn, t + 1 + accum, length, 0, InsertionUP 'change the type first so that the mutated value is within the space of the new type
        ChangeDNA robn, t + 1 + accum, length, 100, InsertionUP 'set a good value up
        accum = length + accum 'Botsareus 3/22/2016 Bugfix Since DNA expended move index down
      End If
    Next t
      
  End With
End Sub

Private Sub Reversal(robn As Integer)
  'reverses a length of DNA
  Dim length As Long
  Dim counter As Long
  Dim location As Long
  Dim low As Long
  Dim high As Long
  Dim templong As Long
  Dim tempblock As block
  Dim t As Long
  Dim second As Long
  
  With rob(robn)
  
 Dim floor As Double
 floor = CDbl(.DnaLen) * CDbl(.Mutables.Mean(ReversalUP) + .Mutables.StdDev(ReversalUP)) / (105 * overtime)
 floor = floor * SimOpts.MutCurrMult
 If .Mutables.mutarray(ReversalUP) < floor Then .Mutables.mutarray(ReversalUP) = floor  'Botsareus 10/5/2015 Prevent freezing

  
    For t = 1 To (.DnaLen - 1)
      If rndy < 1 / (.Mutables.mutarray(ReversalUP) / SimOpts.MutCurrMult) Then
        If .Mutables.Mean(ReversalUP) < 2 Then .Mutables.Mean(ReversalUP) = 2
        
        Do
          length = Gauss(.Mutables.StdDev(ReversalUP), .Mutables.Mean(ReversalUP))
        Loop While length <= 0
        
        length = length \ 2 'be sure we go an even amount to either side
        
        If t - length < 1 Then length = t - 1
        If t + length > .DnaLen - 1 Then length = .DnaLen - 1 - t
        If length > 0 Then
        
          second = 0
          For counter = t - length To t - 1
            tempblock = .dna(counter)
            .dna(counter) = .dna(t + length - second)
            .dna(t + length - second) = tempblock
            second = second + 1
          Next counter
          
          .Mutations = .Mutations + 1
          .LastMut = .LastMut + 1
          
          logmutation robn, "Reversal of" + Str(length * 2 + 1) + "bps centered at " + Str(t) + " during cycle" + _
            Str(SimOpts.TotRunCycle)
         
        End If
      End If
    Next t
  End With
End Sub

Private Sub MinorDeletion(robn As Integer)
  Dim length As Long, t As Long
  With rob(robn)
  
 Dim floor As Double
 floor = CDbl(.DnaLen) / (2.5 * overtime)
 floor = floor * SimOpts.MutCurrMult
 If .Mutables.mutarray(MinorDeletionUP) < floor Then .Mutables.mutarray(MinorDeletionUP) = floor  'Botsareus 10/5/2015 Prevent freezing
  

    If .Mutables.Mean(MinorDeletionUP) < 1 Then .Mutables.Mean(MinorDeletionUP) = 1
    For t = 1 To (.DnaLen - 1)
      If rndy < 1 / (.Mutables.mutarray(MinorDeletionUP) / SimOpts.MutCurrMult) Then
        Do
          length = Gauss(.Mutables.StdDev(MinorDeletionUP), .Mutables.Mean(MinorDeletionUP))
        Loop While length <= 0
        
        If t + length > .DnaLen Then length = .DnaLen - t 'Botsareus 3/22/2016 Bug fix
        If length <= 0 Then Exit Sub  'Botsareus 3/22/2016 Bugfix
        
        Delete .dna, t, length, .DnaLen
        
        .DnaLen = DnaLen(.dna())
        
        .Mutations = .Mutations + 1
        .LastMut = .LastMut + 1
        logmutation robn, "Minor Deletion deleted a run of" + _
          Str(length) + " bps at position" + Str(t) + " during cycle" + _
          Str(SimOpts.TotRunCycle)
        
      End If
    Next t
  End With
End Sub

Private Sub MajorDeletion(robn As Integer)
  Dim length As Long, t As Long
  With rob(robn)
  
 Dim floor As Double
 floor = CDbl(.DnaLen) / (2.5 * overtime)
 floor = floor * SimOpts.MutCurrMult
 If .Mutables.mutarray(MajorDeletionUP) < floor Then .Mutables.mutarray(MajorDeletionUP) = floor  'Botsareus 10/5/2015 Prevent freezing

    If .Mutables.Mean(MajorDeletionUP) < 1 Then .Mutables.Mean(MajorDeletionUP) = 1
    For t = 1 To (.DnaLen - 1)
      If rndy < 1 / (.Mutables.mutarray(MajorDeletionUP) / SimOpts.MutCurrMult) Then
        Do
          length = Gauss(.Mutables.StdDev(MajorDeletionUP), .Mutables.Mean(MajorDeletionUP))
        Loop While length <= 0
        
        If t + length > .DnaLen Then length = .DnaLen - t 'Botsareus 3/22/2016 Bugfix
        If length <= 0 Then Exit Sub  'Botsareus 3/22/2016 Bugfix
        
        Delete .dna, t, length, .DnaLen
        
        .DnaLen = DnaLen(.dna())
        
        .Mutations = .Mutations + 1
        .LastMut = .LastMut + 1
        logmutation robn, "Major Deletion deleted a run of" + _
          Str(length) + " bps at position" + Str(t) + " during cycle" + _
          Str(SimOpts.TotRunCycle)
        
      End If
    Next t
  End With
End Sub

' mutates robot colour in robot n a times
Private Sub mutatecolors(n As Integer, a As Long)
  Dim color As Long
  Dim r As Long, g As Long, b As Long
  Dim counter As Long
  
  color = rob(n).color
  
  b = color \ (65536)
  g = color \ 256 - b * 256
  r = color - b * 65536 - g * 256
  
  For counter = 1 To a
    Select Case (Random(1, 3))
      Case 1
        b = b + (Random(0, 1) * 2 - 1) * 20
      Case 2
        g = g + (Random(0, 1) * 2 - 1) * 20
      Case 3
        r = r + (Random(0, 1) * 2 - 1) * 20
    End Select
    
    If r > 255 Then r = 255
    If r < 0 Then r = 0
    
    If g > 255 Then g = 255
    If g < 0 Then g = 0
    
    If b > 255 Then b = 255
    If b < 0 Then b = 0
  Next counter
  
  rob(n).color = b * 65536 + g * 256 + r
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function delgene(n As Integer, g As Integer) As Boolean
  Dim k As Integer, t As Integer
  k = rob(n).genenum
  If g > 0 And g <= k Then
    DeleteSpecificGene rob(n).dna, g
    delgene = True
    rob(n).DnaLen = DnaLen(rob(n).dna)
    rob(n).genenum = CountGenes(rob(n).dna)
    rob(n).mem(DnaLenSys) = rob(n).DnaLen
    rob(n).mem(GenesSys) = rob(n).genenum
    makeoccurrlist n
  End If
End Function

Private Sub DeleteSpecificGene(ByRef dna() As block, ByVal k As Integer)
  Dim i As Long, f As Long
  
  i = genepos(dna, k)
  If i < 0 Then GoTo getout
  f = GeneEnd(dna, i)
  Delete dna, i, f - i + 1   ' EricL Added +1
getout:
End Sub

Public Sub SetDefaultMutationRates(ByRef changeme As mutationprobs, Optional skipNorm As Boolean = False)
'Botsareus 12/17/2013 Figure out dna length
Dim length As Integer
Dim path As String
If NormMut And Not skipNorm Then
    If optionsform.CurrSpec = 50 Or optionsform.CurrSpec = -1 Then    'only if current spec is selected
        length = rob(robfocus).DnaLen
    Else 'load dna length
        If MaxRobs = 0 Then ReDim rob(0)
        path = TmpOpts.Specie(optionsform.CurrSpec).path & "\" & TmpOpts.Specie(optionsform.CurrSpec).Name
        path = Replace(path, "&#", MDIForm1.MainDir)
        If dir(path) = "" Then path = MDIForm1.MainDir & "\Robots\" & TmpOpts.Specie(optionsform.CurrSpec).Name
        If LoadDNA(path, 0) Then
            length = DnaLen(rob(0).dna)
        End If
    End If
End If

  Dim a As Long
  With (changeme)
  
  For a = 0 To 20
    .mutarray(a) = IIf(NormMut And Not skipNorm, length * CLng(valNormMut), 5000)
    .Mean(a) = 1
    .StdDev(a) = 0
  Next a
  If skipNorm Then .mutarray(P2UP) = 0 'Botsareus 2/21/2014 Might as well disable p2 mutations if loading from the net
  
  SetDefaultLengths changeme
  End With
End Sub

Public Sub SetDefaultLengths(ByRef changeme As mutationprobs)
  With (changeme)
  
  .Mean(PointUP) = 3
  .StdDev(PointUP) = 1
  
  .Mean(DeltaUP) = 500
  .StdDev(DeltaUP) = 150
  
  .Mean(MinorDeletionUP) = 1
  .StdDev(MinorDeletionUP) = 0
  
  .Mean(InsertionUP) = 1
  .StdDev(InsertionUP) = 0
  
  .Mean(CopyErrorUP) = 1
  .StdDev(CopyErrorUP) = 0
  
  .Mean(MajorDeletionUP) = 3
  .StdDev(MajorDeletionUP) = 1
  
  .Mean(ReversalUP) = 3
  .StdDev(ReversalUP) = 1
  
  .CopyErrorWhatToChange = 80
  .PointWhatToChange = 80
  
  .Mean(AmplificationUP) = 250
  .StdDev(AmplificationUP) = 75
  
  .Mean(TranslocationUP) = 250
  .StdDev(TranslocationUP) = 75
  End With

End Sub
