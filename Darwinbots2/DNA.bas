Attribute VB_Name = "DNAExecution"
Option Explicit

'boolstack structure used for conditionals
Private Type boolstack
  val(100) As Boolean
  pos As Integer
End Type

Dim CurrentFlow As Byte
Const CLEAR As Byte = 0
Const COND As Byte = 1
Const body As Byte = 2
Const ELSEBODY As Byte = 3

Dim CurrentCondFlag As Boolean
Const NEXTBODY As Boolean = True 'both of these two are subsets of the clear flag technically
Const NEXTELSE As Boolean = False

Public sysvar(1000) As var    ' array of system variables

Public sysvarIN(255) As var    ' array of system variables informational
Public sysvarOUT(255) As var    ' array of system variables functional

Public Const stacklim As Integer = 100

' stack structure, used for robots' stack
Private Type Stack
  val(stacklim) As Long
  pos As Integer
End Type

Private Type Queue
  memloc As Integer
  Memval As Integer
End Type

Public IntStack As Stack
Public Condst As boolstack 'for the conditions stack
Dim CommandQueue() As Queue 'apply stores at end of cycle

Dim currbot As Long
Dim currgene As Long 'for *.thisgene
Public DisplayActivations As Boolean 'EricL - Toggle for displaying activations in the consol
                                     'Indicates whether the cycle was executed from a console
Public ingene As Boolean             ' Flag for current gene counting.
''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''

Private Sub ExecuteDNA(ByVal n As Integer)
  Dim a As Integer
  Dim b As Integer
  Dim tipo As Long
  Dim i As Integer
  
  currbot = n
  currgene = 0
  CurrentCondFlag = NEXTBODY 'execute body statements if no cond is found
  ingene = False
  
  'New bot.  clear the stacks
  ClearIntStack
  ClearBoolStack
  
  'EricL - March 15, 2006 This section initializes the robot's ga() array to all False so that it can
  'be populated below for those genes that activate this cycle.  Used for displaying
  'Gene Activations.  Only initialized and populated for the robot with the focus or if the bot's console
  'is open.
  If (n = robfocus) Or Not (rob(n).console Is Nothing) Then
  '  rob(n).genenum = CountGenes(rob(n).DNA) ' EricL 4/6/2006 This keeps the gene number up to date
    ReDim rob(n).ga(rob(n).genenum)
    For i = 0 To rob(n).genenum
      rob(n).ga(i) = False
    Next i
  End If
      
  With rob(n)
  a = 1
  rob(n).condnum = 0 ' EricL 4/6/2006 reset the COND statement counter to 0
  While Not (.dna(a).type = 10 And .dna(a).value = 1) And a <= 32000 And a < UBound(.dna) 'Botsareus 6/16/2012 Added upper bounds check (This seems like overkill but I had situations where 'end' command did not exisit)
    tipo = .dna(a).type
    Select Case tipo
      Case 0 'number
        If CurrentFlow <> CLEAR Then
          PushIntStack .dna(a).value
          rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(NUMCOST) * SimOpts.Costs(COSTMULTIPLIER))
        End If
      Case 1 '*number
        If CurrentFlow <> CLEAR Then 'And .DNA(a).value <= 1000 And .DNA(a).value > 0 Then
        
          b = .dna(a).value
          If b > MaxMem Or b < 1 Then
              b = Abs(.dna(a).value) Mod MaxMem
              If b = 0 Then b = 1000 ' Special case that multiples of 1000 should store to location 1000
    
              '2/28/2014 New code from Botsareus if it is a real sysvar then put it into range 'Disabled 3/20/2016 Replaced with Point2 and Copyerror2
              'If Not IsNumeric(SysvarDetok(b, n)) Then .dna(a).value = b
          End If
          
          PushIntStack .mem(b)
          rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(DOTNUMCOST) * SimOpts.Costs(COSTMULTIPLIER))
         ' If .DNA(a).value > EyeStart And .DNA(a).value <= EyeEnd Then ' Can mutations make robots blind?
         '    rob(n).View = True
         ' End If
        End If
      Case 2 'commands (add, sub, etc.)
        If CurrentFlow <> CLEAR Then
          ExecuteBasicCommand .dna(a).value
        End If
      Case 3 'advanced commands
        If CurrentFlow <> CLEAR Then
          ExecuteAdvancedCommand .dna(a).value, a
        End If
      Case 4 'bitwise commands
        If CurrentFlow <> CLEAR Then
          ExecuteBitwiseCommand .dna(a).value
        End If
      Case 5 'conditions
        'EricL  11/2007 New execution paradym.  Conditions can now be executeed anywhere in the gene
        If CurrentFlow = COND Or CurrentFlow = body Or CurrentFlow = ELSEBODY Then
          ExecuteConditions .dna(a).value
        End If
      Case 6 'logic commands (and, or, etc.)
        'EricL  11/2007 New execution paradym.  Conditions can now be executeed anywhere in the gene
        If CurrentFlow = COND Or CurrentFlow = body Or CurrentFlow = ELSEBODY Then
          ExecuteLogic .dna(a).value
        End If
      Case 7 'store, inc and dec
      
          '2/28/2014 New code from Botsareus if it is a real sysvar then put it into range 'Disabled 3/20/2016 Replaced with Point2 and Copyerror2
'          If a > 0 Then
'              b = .dna(a - 1).value
'              If (b > MaxMem Or b < 1) And .dna(a - 1).tipo = 0 Then
'                b = Abs(.dna(a - 1).value) Mod MaxMem
'                If b = 0 Then b = 1000 ' Special case that multiples of 1000 should store to location 1000
'
'                '2/28/2014 New code from Botsareus if it is a real sysvar then put it into range
'                If Not IsNumeric(SysvarDetok(b, n)) Then .dna(a - 1).value = b
'              End If
'          End If

        If CurrentFlow = body Or CurrentFlow = ELSEBODY Then
          If CondStateIsTrue Then  ' Check the Bool stack.  If empty or True on top, do the stores.  Don't if False.
            ExecuteStores .dna(a).value
            If n = robfocus Or Not (rob(n).console Is Nothing) Then rob(n).ga(currgene) = True   'EricL  This gene fired this cycle!  Populate ga()
          End If
        End If
      Case 8 'reserved for a future type
      Case 9 'flow commands
      
        ' EricL 4/6/2006 Added If statement.  This counts the number of COND statements in each bot.
        If Not ExecuteFlowCommands(.dna(a).value, n) Then
          rob(n).condnum = rob(n).condnum + 1
        End If
        
        'If .VirusArray(currgene) > 1 Then 'next gene is busy, so clear flag
        '  CurrentFlow = CLEAR
        'End If
        
        .mem(thisgene) = currgene
      Case 10 'Master flow, such as end, chromostart, etc.
        'ExecuteMasterFlow .dna(a).value
    End Select
    a = a + 1
  Wend
  End With
  CurrentFlow = CLEAR ' EricL 4/15/2006 Do this so next bot doesn't inherit the flow control
End Sub

''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''

Private Sub ExecuteBasicCommand(n As Integer)
Dim i As Long
  '& denotes commands that can be constructed from other commands, but
  'are still basic enough to be listed here
 
 rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(BCCMDCOST) * SimOpts.Costs(COSTMULTIPLIER))
  
  Select Case n
    Case 1 'add
      DNAadd
    Case 2 'sub (negative add) &
      DNASub
    Case 3 'mult
      DNAmult
    Case 4 'div
      DNAdiv
    Case 5 'rnd
      DNArnd
    Case 6 'dereference AKA *
      DNAderef
    Case 7 'mod
      DNAmod
    Case 8  'sgn
      DNAsgn
    Case 9 'absolute value &
      DNAabs
    Case 10 'dup or dupint
      DNAdup
    Case 11 'dropint - Drops the top value on the Int stack
      i = PopIntStack
    Case 12 'clearint - Clears the Int stack
      ClearIntStack
    Case 13 'swapint - Swaps the top two values on the Int stack
      SwapIntStack
    Case 14 'overint - a b -> a b a  Dups the second value on the Int stack
      OverIntStack
  End Select
End Sub

Private Sub DNAadd()
  Dim a As Single
  Dim b As Single
  Dim c As Double
  b = PopIntStack
  a = PopIntStack
  
  a = a Mod 2000000000
  b = b Mod 2000000000
  
  c = a + b
  
  If Abs(c) > 2000000000 Then c = c - Sgn(c) * 2000000000
  PushIntStack c
End Sub

Private Sub DNASub() 'Botsareus 5/20/2012 new code to stop overflow
  Dim a As Single
  Dim b As Single
  Dim c As Double
  b = PopIntStack
  a = PopIntStack
  
  
  a = a Mod 2000000000
  b = b Mod 2000000000
  
  c = a - b
  
  If Abs(c) > 2000000000 Then c = c - Sgn(c) * 2000000000
  PushIntStack c
End Sub

Private Sub DNAmult()
  Dim a As Long
  Dim b As Long
  Dim c As Double
  b = PopIntStack
  a = PopIntStack
  c = CDbl(a) * CDbl(b)
  If Abs(c) > 2000000000 Then c = Sgn(c) * 2000000000
  PushIntStack CLng(c)
End Sub

Private Sub DNAdiv()
  Dim a As Long
  Dim b As Long
  b = PopIntStack
  a = PopIntStack
  If b <> 0 Then
    PushIntStack a / b
  Else
    PushIntStack 0
  End If
End Sub

Private Sub DNArnd()
  PushIntStack Random(0, PopIntStack)
End Sub

Private Sub DNAderef()
  Dim b As Long
  
  b = PopIntStack
  
'  If b > EyeStart And b < EyeEnd Then
'    rob(currbot).View = True
'  End If
  b = Abs(b) Mod MaxMem
  If b = 0 Then b = 1000 ' Special case that multiples of 1000 should store to location 1000
'  If b <= 1000 And b >= 1 Then
  PushIntStack rob(currbot).mem(b)
 ' Else
  '  PushIntStack 0
  'End If
End Sub

Private Sub DNAmod()
  Dim b As Long
  
  b = PopIntStack
  If b = 0 Then
    PopIntStack
    PushIntStack 0
  Else
    PushIntStack PopIntStack Mod b
  End If
End Sub

Private Sub DNAsgn()
  PushIntStack Sgn(PopIntStack)
End Sub

Private Sub DNAabs()
  PushIntStack Abs(PopIntStack)
End Sub

Private Sub DNAdup()
  Dim b As Long
  
  b = PopIntStack
  PushIntStack b
  PushIntStack b
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub ExecuteAdvancedCommand(n As Integer, at_position As Integer)

If n < 13 Then rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(ADCMDCOST) * SimOpts.Costs(COSTMULTIPLIER))

  Select Case n
    Case 1 'findang
      findang
    Case 2 'finddist
      finddist
    Case 3 'ceil
      DNAceil
    Case 4 'floor
      DNAfloor
    Case 5 ' sqr
      DNASqr
    Case 6 ' power
      DNApow
    Case 7 ' pyth
      DNApyth
    Case 8
      DNAanglecmp
    Case 9
      DNAroot  ' a ^ (1/b)
    Case 10
      DNAlogx  ' log(a) / Log(b)
    Case 11
      DNAsin
    Case 12
      DNAcos
  End Select
End Sub

Private Sub DNAanglecmp() 'Allowes a robot to quickly calculate the difference between two angles
    Dim a As Double
    Dim b As Double
    Dim c As Double
    
    b = PopIntStack
    a = PopIntStack
    
    'Botsareus 10/5/2015 Value normalization
    b = b Mod 1256
    If b < 0 Then b = b + 1256
    
    a = a Mod 1256
    If a < 0 Then a = a + 1256
    
    c = AngDiff(a / 200, b / 200) * 200
        
    PushIntStack c
End Sub

Private Sub findang()
  Dim a As Single  'target xpos
  Dim b As Single  'target ypos
  Dim c As Single  'robot's xpos
  Dim d As Single  'robot's ypos
  Dim e As Single  'angle to target
  b = PopIntStack
  a = PopIntStack
  
  c = robManager.GetRobotPosition(currbot).x / Form1.xDivisor
  d = robManager.GetRobotPosition(currbot).y / Form1.yDivisor
  e = angnorm(angle(c, d, a, b)) * 200
  PushIntStack e
End Sub


Private Sub finddist()
  Dim a As Single  'target xpos
  Dim b As Single  'target ypos
  Dim c As Single  'robot's xpos
  Dim d As Single  'robot's ypos
  Dim e As Single  'distance to target
  b = PopIntStack * Form1.yDivisor
  a = PopIntStack * Form1.xDivisor
  c = robManager.GetRobotPosition(currbot).x
  d = robManager.GetRobotPosition(currbot).y
  e = Sqr(((c - a) ^ 2 + (d - b) ^ 2))
  If Abs(e) > 2000000000# Then
    e = Sgn(e) * 2000000000#
  End If
  PushIntStack CLng(e)
End Sub

'applies a ceiling to a value on the stack.
'Usage: val ceilingvalue ceil.
Private Sub DNAceil()
  Dim a As Single
  Dim b As Single
  
  b = PopIntStack
  a = PopIntStack

  PushIntStack IIf(a > b, b, a)
End Sub

'similar to ceil but with a floor instead
Private Sub DNAfloor()
  Dim a As Long
  Dim b As Long
  
  b = PopIntStack
  a = PopIntStack

  PushIntStack IIf(a < b, b, a)
End Sub

'Returns square root of a positive number. Can't think of a specific use but it is valid.
Private Sub DNASqr()
    Dim a As Single
    a = PopIntStack
    Dim b As Single
    
    If a > 0 Then
      b = Sqr(a)
    Else
      b = 0
    End If
    
    PushIntStack b
End Sub

Private Sub DNAsin()
    Dim a As Single
    a = PopIntStack
    Dim b As Single
    
      b = Sin(a / 200) * 32000
    PushIntStack b
End Sub

Private Sub DNAcos()
    Dim a As Single
    a = PopIntStack
    Dim b As Single
    
    b = Cos(a / 200) * 32000
    
    PushIntStack b
End Sub

'returns a power number. Raises a (top number) to the power of b (second number)
'Seems kind of pointless to me
Private Sub DNApow()
    Dim a As Double
    Dim b As Double
    Dim c As Double
    b = PopIntStack
    a = PopIntStack
    
    If Abs(b) > 10 Then b = 10 * Sgn(b)
    
    If a = 0 Then
      c = 0
    Else
      c = a ^ b
    End If
    If Abs(c) > 2000000000 Then c = Sgn(c) * 2000000000
    PushIntStack c
End Sub

'Botsareus 9/7/2013 more power commands
Private Sub DNAroot()
    Dim a As Double
    Dim b As Double
    Dim c As Double
    b = Abs(PopIntStack)
    a = Abs(PopIntStack)
    
    If b = 0 Then
      c = 0
    Else
      c = a ^ (1 / b)
    End If
    PushIntStack c
End Sub

Private Sub DNAlogx()
    Dim a As Double
    Dim b As Double
    Dim c As Double
    b = Abs(PopIntStack)
    a = Abs(PopIntStack)
    
    If b < 2 Or a = 0 Then 'Botsareus 9/15/2013 b is changed to 2 to avoid div/0
      c = 0
    Else
      c = Log(a) / Log(b)
    End If
    PushIntStack c
End Sub

Private Sub DNApyth()
  Dim a As Single
  Dim b As Single
  
  b = PopIntStack
  a = PopIntStack
  Dim c As Single
  
  c = Sqr(a * a + b * b)
  If Abs(c) > 2000000000 Then c = Sgn(c) * 2000000000
  
  PushIntStack c
End Sub



'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''

'unimplemented as yet
Private Sub ExecuteBitwiseCommand(n As Integer)
 
 rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(BTCMDCOST) * SimOpts.Costs(COSTMULTIPLIER))

  Select Case n
    Case 1 'Compliment ~ (tilde)
      DNABitwiseCompliment
    Case 2 '& And
      DNABitwiseAND
    Case 3 '| or
      DNABitwiseOR
    Case 4 ' XOR, ^ (I need another representation)
      DNABitwiseXOR
    Case 5 'bitinc ++
      DNABitwiseINC
    Case 6 'bitdec --
      DNABitwiseDEC
    Case 7 'negate
      PushIntStack -PopIntStack
    Case 8 ' <<
      DNABitwiseShiftLeft
    Case 9 ' >>
      DNABitwiseShiftRight
  End Select
End Sub

Private Sub DNABitwiseCompliment()
  Dim value As Long
  Dim bits As DoubleWord
  
  value = PopIntStack
  bits = NumberToBit(value)
  InvertBits bits
  PushIntStack BitToNumber(bits)
End Sub

Private Sub DNABitwiseAND()
  Dim valueA As Long
  Dim valueB As Long
  Dim bitsA As DoubleWord
  Dim bitsB As DoubleWord
  
  valueB = PopIntStack
  valueA = PopIntStack
  
  bitsA = NumberToBit(valueA)
  bitsB = NumberToBit(valueB)
  
  bitsA = BitAND(bitsA, bitsB)
  PushIntStack BitToNumber(bitsA)
End Sub

Private Sub DNABitwiseOR()
  Dim valueA As Long
  Dim valueB As Long
  Dim bitsA As DoubleWord
  Dim bitsB As DoubleWord
  
  valueB = PopIntStack
  valueA = PopIntStack
  
  bitsA = NumberToBit(valueA)
  bitsB = NumberToBit(valueB)
  
  bitsA = BitOR(bitsA, bitsB)
  PushIntStack BitToNumber(bitsA)
End Sub

Private Sub DNABitwiseXOR()
  Dim valueA As Long
  Dim valueB As Long
  Dim bitsA As DoubleWord
  Dim bitsB As DoubleWord
  
  valueB = PopIntStack
  valueA = PopIntStack
  
  bitsA = NumberToBit(valueA)
  bitsB = NumberToBit(valueB)
  
  bitsA = BitXOR(bitsA, bitsB)
  PushIntStack BitToNumber(bitsA)
End Sub

Private Sub DNABitwiseINC()
  Dim value As Long
  Dim bits As DoubleWord
  
  value = PopIntStack
  bits = NumberToBit(value)
  IncBits bits
  PushIntStack BitToNumber(bits)
End Sub

Private Sub DNABitwiseDEC()
  Dim value As Long
  Dim bits As DoubleWord
  
  value = PopIntStack
  bits = NumberToBit(value)
  DecBits bits
  PushIntStack BitToNumber(bits)
End Sub

Private Sub DNABitwiseShiftLeft()
  Dim value As Long
  Dim bits As DoubleWord
  
  value = PopIntStack
  bits = NumberToBit(value)
  BitShiftLeft bits
  PushIntStack BitToNumber(bits)
End Sub

Private Sub DNABitwiseShiftRight()
  Dim value As Long
  Dim bits As DoubleWord
  
  value = PopIntStack
  bits = NumberToBit(value)
  BitShiftRight bits
  PushIntStack BitToNumber(bits)
End Sub

'''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''

Private Sub ExecuteConditions(n As Integer)
  rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(CONDCOST) * SimOpts.Costs(COSTMULTIPLIER))
  
  Select Case n
    Case 1 '<
      Min
    Case 2 '>
      magg
    Case 3 '=
      equa
    Case 4 ' <>, !=
      diff
    Case 5 ' %=
      cequa
    Case 6 '!%=
      cdiff
    Case 7 '~=
      customcequa
    Case 8 '!~=
      customcdiff
    Case 9  '>=
      maggequal
    Case 10 '<=
      minequal
  End Select
End Sub

Private Function Min() As Boolean
  PushBoolStack (PopIntStack > PopIntStack)
End Function

Private Function magg() As Boolean
  PushBoolStack (PopIntStack < PopIntStack)
End Function

Private Function equa() As Boolean
  PushBoolStack (PopIntStack = PopIntStack)
End Function

Private Function diff() As Boolean
  PushBoolStack (PopIntStack <> PopIntStack)
End Function

Private Function cequa() As Boolean
  Dim a As Single
  Dim b As Single
  Dim c As Single
  b = PopIntStack
  a = PopIntStack
  c = a / 10
  PushBoolStack ((a - c <= b) And (a + c >= b))
End Function

Private Function cdiff() As Boolean
  Dim a As Single
  Dim b As Single
  Dim c As Single
  b = PopIntStack
  a = PopIntStack
  c = a / 10
  PushBoolStack (Not ((a + c >= b) And (a - c <= b)))
End Function
Private Function customcequa() As Boolean
'usage: 10 20 30 ~= are 10 and 20 within 30 percent of each other?
  Dim a As Long
  Dim b As Long
  Dim c As Single
  Dim d As Long
  
  d = PopIntStack
  b = PopIntStack
  a = PopIntStack
  c = a / 100 * d
  PushBoolStack ((a - c <= b) And (a + c >= b))
End Function

Private Function customcdiff() As Boolean
  Dim a As Long
  Dim b As Long
  Dim c As Single
  Dim d As Long
    
  d = PopIntStack
  b = PopIntStack
  a = PopIntStack
  c = a / 100 * d
  If Abs(c) > 2000000000 Then c = Sgn(c) * 2000000000
  PushBoolStack (Not ((a + c >= b) And (a - c <= b)))
End Function

Private Function minequal() As Boolean
  PushBoolStack (PopIntStack >= PopIntStack)
End Function

Private Function maggequal() As Boolean
  PushBoolStack (PopIntStack <= PopIntStack)
End Function

'''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''

Private Sub ExecuteLogic(n As Integer)

rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(LOGICCOST) * SimOpts.Costs(COSTMULTIPLIER))

  Dim a As Integer, b As Integer

    Select Case n
      Case 1 'and
        b = PopBoolStack
        If b = -5 Then b = True
        a = PopBoolStack
        If a <> -5 Then
          PushBoolStack a And b
        Else
          PushBoolStack b
        End If
      Case 2 'or
        b = PopBoolStack
        If b = -5 Then b = True
        a = PopBoolStack
        If a <> -5 Then
          PushBoolStack a Or b
        Else
          PushBoolStack True
        End If
      Case 3 'xor
        b = PopBoolStack
        If b = -5 Then b = True
        a = PopBoolStack
        If a <> -5 Then
          PushBoolStack a Xor b
        Else
          PushBoolStack Not b
        End If
      Case 4 'not
        b = PopBoolStack
        If b = -5 Then b = True
        PushBoolStack Not b
      Case 5 ' true
        PushBoolStack True
      Case 6 ' false
        PushBoolStack False
      Case 7 ' dropbool
        b = PopBoolStack
      Case 8 ' clearbool
        ClearBoolStack
      Case 9 ' dupbool
        DupBoolStack
      Case 10 ' swapbool
        SwapBoolStack
      Case 11 ' overbool
        OverBoolStack
    End Select

End Sub

'''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''

Private Sub ExecuteStores(n As Integer)
   Select Case n
     Case 1 'store
       DNAstore
     Case 2 'inc
       DNAinc
     Case 3 'dec
       DNAdec
     Case 4 '+= 'Botsareus 9/7/2013 New commannds
        DNAaddstore
     Case 5 '-=
        DNAsubstore
     Case 6 '*=
        DNAmultstore
     Case 7 '/=
        DNAdivstore
     Case 8
        DNAceilstore
     Case 9
        DNAfloorstore
     Case 10
        DNArndstore
     Case 11
        DNAsgnstore
     Case 12
        DNAabsstore
     Case 13
        DNAsqrstore
     Case 14
        DNAnegstore
   End Select
End Sub

Private Sub DNAstore()
   Dim b As Long
   Dim a As Long
   b = PopIntStack          ' Pop the stack and get the mem location to store to
   If b <> 0 Then           ' Stores to 0 are allowed, but do nothing and cost nothing
     b = Abs(b) Mod MaxMem  ' Make sure the location hits the bot's memory to increase the chance of mutations hitting sysvars.
     If b = 0 Then b = 1000 ' Special case that multiples of 1000 should store to location 1000
     a = PopIntStack
     
     'Botsareus 3/22/2013 handle tieang...tielen 1...4 overwrites
     Dim k As Byte
     For k = 0 To 3
      If b = 480 + k Then rob(currbot).TieAngOverwrite(k) = True
      If b = 484 + k Then rob(currbot).TieLenOverwrite(k) = True
     Next
     
     rob(currbot).mem(b) = mod32000(a)
     rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(COSTSTORE) * SimOpts.Costs(COSTMULTIPLIER))
   End If
End Sub

Private Sub DNAinc()
   Dim a As Long, b As Long
   a = PopIntStack
   If a <> 0 Then
     a = Abs(a) Mod MaxMem
     If a = 0 Then a = 1000
     b = rob(currbot).mem(a) + 1
     rob(currbot).mem(a) = mod32000(b)
     rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(COSTSTORE) * SimOpts.Costs(COSTMULTIPLIER)) / 10
   End If
End Sub

Private Sub DNAdec()
   Dim a As Long, b As Long
   a = PopIntStack
    If a <> 0 Then
     a = Abs(a) Mod MaxMem
     If a = 0 Then a = 1000
     b = rob(currbot).mem(a) - 1
     rob(currbot).mem(a) = mod32000(b)
     rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(COSTSTORE) * SimOpts.Costs(COSTMULTIPLIER)) / 10
   End If
End Sub

Private Sub DNAaddstore()
   Dim b As Long
   Dim a As Long
   b = PopIntStack          ' Pop the stack and get the mem location to store to
   If b <> 0 Then           ' Stores to 0 are allowed, but do nothing and cost nothing
     b = Abs(b) Mod MaxMem  ' Make sure the location hits the bot's memory to increase the chance of mutations hitting sysvars.
     If b = 0 Then b = 1000 ' Special case that multiples of 1000 should store to location 1000
     a = PopIntStack + rob(currbot).mem(b)
     
     'Botsareus 3/22/2013 handle tieang...tielen 1...4 overwrites
     Dim k As Byte
     For k = 0 To 3
      If b = 480 + k Then rob(currbot).TieAngOverwrite(k) = True
      If b = 484 + k Then rob(currbot).TieLenOverwrite(k) = True
     Next
     
     rob(currbot).mem(b) = mod32000(a)
     rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(COSTSTORE) * SimOpts.Costs(COSTMULTIPLIER)) / 5
   End If
End Sub

Private Sub DNAsubstore()
   Dim b As Long
   Dim a As Long
   b = PopIntStack          ' Pop the stack and get the mem location to store to
   If b <> 0 Then           ' Stores to 0 are allowed, but do nothing and cost nothing
     b = Abs(b) Mod MaxMem  ' Make sure the location hits the bot's memory to increase the chance of mutations hitting sysvars.
     If b = 0 Then b = 1000 ' Special case that multiples of 1000 should store to location 1000
     a = rob(currbot).mem(b) - PopIntStack
     
     'Botsareus 3/22/2013 handle tieang...tielen 1...4 overwrites
     Dim k As Byte
     For k = 0 To 3
      If b = 480 + k Then rob(currbot).TieAngOverwrite(k) = True
      If b = 484 + k Then rob(currbot).TieLenOverwrite(k) = True
     Next
     
     rob(currbot).mem(b) = mod32000(a)
     rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(COSTSTORE) * SimOpts.Costs(COSTMULTIPLIER)) / 5
   End If
End Sub

Private Sub DNAmultstore()
   Dim b As Long
   Dim a As Long
   b = PopIntStack          ' Pop the stack and get the mem location to store to
   If b <> 0 Then           ' Stores to 0 are allowed, but do nothing and cost nothing
     b = Abs(b) Mod MaxMem  ' Make sure the location hits the bot's memory to increase the chance of mutations hitting sysvars.
     If b = 0 Then b = 1000 ' Special case that multiples of 1000 should store to location 1000
     
     'Botsareus 11/30/2013 Small bugfix to prevent overflow
     Dim c As Long
     c = PopIntStack
     c = mod32000(c)
     
     a = rob(currbot).mem(b) * c
     
     'Botsareus 3/22/2013 handle tieang...tielen 1...4 overwrites
     Dim k As Byte
     For k = 0 To 3
      If b = 480 + k Then rob(currbot).TieAngOverwrite(k) = True
      If b = 484 + k Then rob(currbot).TieLenOverwrite(k) = True
     Next
     
     rob(currbot).mem(b) = mod32000(a)
     rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(COSTSTORE) * SimOpts.Costs(COSTMULTIPLIER)) / 5
   End If
End Sub

Private Sub DNAdivstore()
   Dim b As Long
   Dim a As Long
   Dim c As Long
   b = PopIntStack          ' Pop the stack and get the mem location to store to
   c = PopIntStack
   If b <> 0 Then           ' Stores to 0 are allowed, but do nothing and cost nothing
     b = Abs(b) Mod MaxMem  ' Make sure the location hits the bot's memory to increase the chance of mutations hitting sysvars.
     If b = 0 Then b = 1000 ' Special case that multiples of 1000 should store to location 1000
     If c = 0 Then
      a = 0
     Else
      a = rob(currbot).mem(b) / c
     End If
     
     'Botsareus 3/22/2013 handle tieang...tielen 1...4 overwrites
     Dim k As Byte
     For k = 0 To 3
      If b = 480 + k Then rob(currbot).TieAngOverwrite(k) = True
      If b = 484 + k Then rob(currbot).TieLenOverwrite(k) = True
     Next
     
     rob(currbot).mem(b) = a
     rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(COSTSTORE) * SimOpts.Costs(COSTMULTIPLIER)) / 5
   End If
End Sub

Private Sub DNAceilstore()
   Dim b As Long
   Dim a As Long
   Dim c As Long
   b = PopIntStack          ' Pop the stack and get the mem location to store to
   c = PopIntStack
   If b <> 0 Then           ' Stores to 0 are allowed, but do nothing and cost nothing
     b = Abs(b) Mod MaxMem  ' Make sure the location hits the bot's memory to increase the chance of mutations hitting sysvars.
     If b = 0 Then b = 1000 ' Special case that multiples of 1000 should store to location 1000
     a = IIf(rob(currbot).mem(b) > c, c, rob(currbot).mem(b))
     
     'Botsareus 3/22/2013 handle tieang...tielen 1...4 overwrites
     Dim k As Byte
     For k = 0 To 3
      If b = 480 + k Then rob(currbot).TieAngOverwrite(k) = True
      If b = 484 + k Then rob(currbot).TieLenOverwrite(k) = True
     Next
          
     rob(currbot).mem(b) = mod32000(a)
     rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(COSTSTORE) * SimOpts.Costs(COSTMULTIPLIER)) / 5
   End If
End Sub

Private Sub DNAfloorstore()
   Dim b As Long
   Dim a As Long
   Dim c As Long
   b = PopIntStack          ' Pop the stack and get the mem location to store to
   c = PopIntStack
   If b <> 0 Then           ' Stores to 0 are allowed, but do nothing and cost nothing
     b = Abs(b) Mod MaxMem  ' Make sure the location hits the bot's memory to increase the chance of mutations hitting sysvars.
     If b = 0 Then b = 1000 ' Special case that multiples of 1000 should store to location 1000
     a = IIf(rob(currbot).mem(b) < c, c, rob(currbot).mem(b))
     
     'Botsareus 3/22/2013 handle tieang...tielen 1...4 overwrites
     Dim k As Byte
     For k = 0 To 3
      If b = 480 + k Then rob(currbot).TieAngOverwrite(k) = True
      If b = 484 + k Then rob(currbot).TieLenOverwrite(k) = True
     Next
     rob(currbot).mem(b) = mod32000(a)
     rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(COSTSTORE) * SimOpts.Costs(COSTMULTIPLIER)) / 5
   End If
End Sub


Private Function mod32000(ByVal a As Long) As Integer
'Botsareus 10/6/2015 Fix for out of range

     If a > 0 Then
       a = a Mod 32000
       If a = 0 Then a = 32000  ' Special case 32000
     ElseIf a < 0 Then
       a = a Mod 32000
       If a = 0 Then a = -32000 ' special case -32000
     End If

mod32000 = a
End Function

Private Sub DNArndstore()
   Dim a As Long, b As Long
   a = PopIntStack
   If a <> 0 Then
     a = Abs(a) Mod MaxMem
     If a = 0 Then a = 1000
     b = Random(0, Abs(rob(currbot).mem(a))) * Sgn(rob(currbot).mem(a))
     rob(currbot).mem(a) = b
     rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(COSTSTORE) * SimOpts.Costs(COSTMULTIPLIER)) / 7
   End If
End Sub

Private Sub DNAsgnstore()
   Dim a As Long, b As Long
   a = PopIntStack
   If a <> 0 Then
     a = Abs(a) Mod MaxMem
     If a = 0 Then a = 1000
     b = Sgn(rob(currbot).mem(a))
     rob(currbot).mem(a) = b
     rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(COSTSTORE) * SimOpts.Costs(COSTMULTIPLIER)) / 7
   End If
End Sub

Private Sub DNAabsstore()
   Dim a As Long, b As Long
   a = PopIntStack
   If a <> 0 Then
     a = Abs(a) Mod MaxMem
     If a = 0 Then a = 1000
     b = Abs(rob(currbot).mem(a))
     rob(currbot).mem(a) = b
     rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(COSTSTORE) * SimOpts.Costs(COSTMULTIPLIER)) / 8
   End If
End Sub

Private Sub DNAsqrstore()
   Dim a As Long, b As Long
   a = PopIntStack
   If a <> 0 Then
     a = Abs(a) Mod MaxMem
     If a = 0 Then a = 1000
     If rob(currbot).mem(a) > 0 Then
        b = Sqr(rob(currbot).mem(a))
     Else
        b = 0
     End If
     rob(currbot).mem(a) = b
     rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(COSTSTORE) * SimOpts.Costs(COSTMULTIPLIER)) / 7
   End If
End Sub

Private Sub DNAnegstore()
   Dim a As Long, b As Long
   a = PopIntStack
   If a <> 0 Then
     a = Abs(a) Mod MaxMem
     If a = 0 Then a = 1000
     b = -rob(currbot).mem(a)
     rob(currbot).mem(a) = b
     rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(COSTSTORE) * SimOpts.Costs(COSTMULTIPLIER)) / 8
   End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''
Private Function ExecuteFlowCommands(n As Integer, bot As Integer) As Boolean

rob(currbot).nrg = rob(currbot).nrg - (SimOpts.Costs(FLOWCOST) * SimOpts.Costs(COSTMULTIPLIER))

'returns true if a stop command was found (start, stop, or else)
'returns false if cond was found
  ExecuteFlowCommands = False
  Select Case n
    Case 1 'cond
      CurrentFlow = COND
      currgene = currgene + 1
      ClearBoolStack
      ingene = True
      GoTo getout
    Case 2, 3, 4 'assume a stop command, or it really is a stop command
      'this is supposed to come before case 2 and 3, since these commands
      'must be executed before start and else have a chance to go
      ExecuteFlowCommands = True
      If CurrentFlow = COND Then CurrentCondFlag = AddupCond
      If Not ingene Then CurrentCondFlag = NEXTBODY
      If CurrentCondFlag And (CurrentFlow = ELSEBODY Or CurrentFlow = body) Then 'Botsareus 3/24/2012 Fixed a bug where: any else gene was showing activation
        ' Need to check this for the case where the gene body doesn't have any stores to trigger the activation dialog
        If bot = robfocus Or Not (rob(bot).console Is Nothing) Then rob(bot).ga(currgene) = True   'EricL  This gene fired this cycle!  Populate ga()
      End If
      CurrentFlow = CLEAR
      Select Case n
        Case 2 'start
          If Not ingene Then ' the first start or else after a cond is not a new gene but the rest are
            currgene = currgene + 1
          End If
          ingene = False
          If CurrentCondFlag = NEXTBODY Then CurrentFlow = body
        Case 3 'else
          If CurrentCondFlag = NEXTELSE Then CurrentFlow = ELSEBODY
          If Not ingene Then
            currgene = currgene + 1
          End If
          ingene = False
        Case 4 ' stop
          ingene = False
          CurrentFlow = CLEAR
      End Select
    End Select
getout:
End Function

Private Function AddupCond() As Boolean
  'AND together all conditions on the boolstack
  Dim a As Integer
  
  AddupCond = True
  
  a = PopBoolStack
  While a <> -5
    AddupCond = AddupCond And a
    a = PopBoolStack
  Wend
End Function

' EricL 11/2007 - New execution paradym.  Returns true if the bool stack is empty or has true on the top.
Private Function CondStateIsTrue() As Boolean

Dim a As Integer

  CondStateIsTrue = True
  
  a = PopBoolStack
  If a = -5 Then GoTo getout
  PushBoolStack CBool(a)           ' If we popped something off the stack, push it back on
    
  If a = False Then CondStateIsTrue = False ' Return True unless False is on the top of the stack
getout:

End Function

'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''

'
'   L O A D I N G   A N D   P A R S I N G
'

' root of the dna execution part
' takes each robot and passes it to the interpreter
' with some variations for console debug and genes activation
Public Sub ExecRobs()
  Dim t As Integer
  Dim k As Integer

  For t = 1 To MaxRobs
    If t Mod 250 = 0 Then DoEvents
    If robManager.GetExists(t) And Not rob(t).Corpse And Not rob(t).DisableDNA Then
      ExecuteDNA t
      If Not (rob(t).console Is Nothing) And DisplayActivations Then
         rob(t).console.textout ""
         rob(t).console.textout "***ROBOT GENES EXECUTION***" 'Botsareus 3/24/2012 looks a little better now
         For k = 1 To rob(t).genenum
          If rob(t).ga(k) Then
            rob(t).console.textout CStr(k) & " executed"
          Else
            rob(t).console.textout CStr(k) & " not executed" 'Botsareus 3/24/2012 looks a little better now
          End If
        Next k
      End If
      If t = robfocus And ActivForm.Visible Then
          exechighlight t
      End If
    End If
  Next t
  
End Sub
