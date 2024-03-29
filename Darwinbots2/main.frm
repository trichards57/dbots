VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8436
   ClientLeft      =   36
   ClientTop       =   96
   ClientWidth     =   12036
   FillColor       =   &H00C00000&
   ForeColor       =   &H00511206&
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8436
   ScaleWidth      =   12036
   Begin VB.Timer SecTimer 
      Interval        =   1000
      Left            =   2040
      Top             =   120
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblSaving 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   4440
      TabIndex        =   5
      Top             =   4680
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label lblSafeMode 
      BackStyle       =   0  'Transparent
      Caption         =   "Safe Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label GraphLab 
      BackStyle       =   0  'Transparent
      Caption         =   "Updating Graph: 0%"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label BoyLabl 
      BackStyle       =   0  'Transparent
      Caption         =   $"main.frx":08CA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   9855
   End
   Begin VB.Image InternetModePopFileProblem 
      Height          =   204
      Left            =   5280
      Picture         =   "main.frx":0959
      Top             =   5400
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image InternetModeBackPressure 
      Height          =   204
      Left            =   5280
      Picture         =   "main.frx":0CCB
      Top             =   5040
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image InternetModeStart 
      Height          =   204
      Left            =   5280
      Picture         =   "main.frx":103D
      Top             =   4680
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image InternetModeOff 
      Height          =   204
      Left            =   5280
      Picture         =   "main.frx":13AF
      Top             =   4320
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image ServerGood 
      Height          =   204
      Left            =   5280
      Picture         =   "main.frx":1721
      Top             =   3600
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image ServerBad 
      Height          =   204
      Left            =   5280
      Picture         =   "main.frx":1A93
      Top             =   3960
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Label InternetMode 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Internet Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   3225
   End
   Begin VB.Image TeleporterMask 
      Height          =   576
      Left            =   240
      Picture         =   "main.frx":1E05
      Top             =   7560
      Visible         =   0   'False
      Width           =   9216
   End
   Begin VB.Image Teleporter 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   240
      Picture         =   "main.frx":1CE47
      Stretch         =   -1  'True
      Top             =   6600
      Visible         =   0   'False
      Width           =   11520
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Output video interrotto... premi sul pulsante nella barra in alto per ripristinarlo."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   600
      TabIndex        =   0
      Tag             =   "30000"
      Top             =   1440
      Visible         =   0   'False
      Width           =   10665
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const PIOVER4 = PI / 4

Private Type tmppostyp
  n As Integer
  x As Single
  y As Single
End Type

Public camfix As Boolean 'Botsareus 2/23/2013 normalizes screen
Public pausefix As Boolean 'Botsareus 3/6/2013 Figure out if simulation must start paused
Public TotalOffspring As Long 'Botsareus 5/22/2013 For find best
Public cyc As Integer          ' cycles/second
Public dispskin As Boolean  ' skin drawing enabled?
Public Active As Boolean    ' sim running?
Public visiblew As Single     ' field visible portion (for zoom)
Public visibleh As Long
Public DNAMaxConds As Integer   ' max conditions per gene allowed by mutation
Public twipWidth As Single
Public TwipHeight As Single
Public FortyEightOverTwipWidth As Single
Public FortyEightOverTwipHeight As Single
Public xDivisor As Single
Public yDivisor As Single
Private tmprob_c As Byte
Private tmppos(50) As tmppostyp
Private p_reclev As Integer 'Botsareus 8/3/2012 for generational distance
Private MouseClickX As Long     ' mouse pos when clicked
Private MouseClickY As Long
Private MouseClicked As Boolean
Private ZoomFlag As Boolean        ' EricL True while mouse button four is held down - for zooming
Private DraggingBot As Boolean     ' EricL True while mouse is down dragging bot around

Private Sub BoyLabl_Click()
    BoyLabl.Visible = False
End Sub

Private Sub Form_Load()
    Dim i As Integer
     
    strings Me
    
    LoadSysVars
    Form1.Top = 0
    Form1.Left = 0
    Form1.Width = MDIForm1.ScaleWidth
    Form1.Height = MDIForm1.ScaleHeight
    visiblew = SimOpts.FieldWidth
    visibleh = SimOpts.FieldHeight
    MDIForm1.visualize = True
    MaxMem = 1000
    maxfieldsize = SimOpts.FieldWidth * 2
    TotalRobots = 0
    robfocus = 0
    MDIForm1.DisableRobotsMenu
    dispskin = True
    
    FlashColor(1) = vbBlack         ' Hit with memory shot
    FlashColor(-1 + 10) = vbRed     ' Hit with Nrg feeding shot
    FlashColor(-2 + 10) = vbWhite   ' Hit with Nrg Shot
    FlashColor(-3 + 10) = vbBlue    ' Hit with venom shot
    FlashColor(-4 + 10) = vbGreen   ' Hit with waste shot
    FlashColor(-5 + 10) = vbYellow  ' Hit with poison Shot
    FlashColor(-6 + 10) = vbMagenta ' Hit with body feeding shot
    FlashColor(-7 + 10) = vbCyan    ' Hit with virus shot
End Sub

'
'              D R A W I N G
'

' redraws screen
Public Sub Redraw()
    Dim t As Integer
    Dim pos As Vector
    'Botsareus 6/23/2016 Need to offset the robots by (actual velocity minus velocity) before drawing them
    'It is a hacky way of doing it, but should be a bit faster since the computation is only preformed twice, other than preforming it in each subsection.
    For t = 1 To MaxRobs
        If robManager.GetExists(t) Then
            pos.x = robManager.GetRobotPosition(t).x - (robManager.GetVelocity(t).x - robManager.GetActualVelocity(t).x)
            pos.y = robManager.GetRobotPosition(t).y - (robManager.GetVelocity(t).y - robManager.GetActualVelocity(t).y)
            robManager.SetRobotPosition t, pos
        End If
    Next

    Dim count As Long
  
    Cls
    
    DrawArena
    DrawAllTies
    DrawAllRobs
    DrawShots
    Me.AutoRedraw = True
  
    'Place robots back
    For t = 1 To MaxRobs
        If robManager.GetExists(t) Then
            
            pos.x = robManager.GetRobotPosition(t).x + (robManager.GetVelocity(t).x - robManager.GetActualVelocity(t).x)
            pos.y = robManager.GetRobotPosition(t).y + (robManager.GetVelocity(t).y - robManager.GetActualVelocity(t).y)
            
            robManager.SetRobotPosition t, pos
        End If
    Next
End Sub

' calculates the pixel/twip ratio, since some graphic methods
' need pixel values
Function GetTwipWidth() As Single
    Dim scw As Long, sch As Long, scm As Integer
    Dim sct As Long, scl As Long
    scm = Form1.ScaleMode
    scw = Form1.ScaleWidth
    sch = Form1.ScaleHeight
    sct = Form1.ScaleTop
    scl = Form1.ScaleLeft
    Form1.ScaleMode = vbPixels
    GetTwipWidth = Form1.ScaleWidth / scw
    Form1.ScaleMode = scm
    Form1.ScaleWidth = scw
    Form1.ScaleHeight = sch
    Form1.ScaleTop = sct
    Form1.ScaleLeft = scl
End Function

' calculates the pixel/twip ratio, since some graphic methods
' need pixel values
Function GetTwipHeight() As Single
    Dim scw As Long, sch As Long, scm As Integer
    Dim sct As Long, scl As Long
    scm = Form1.ScaleMode
    scw = Form1.ScaleWidth
    sch = Form1.ScaleHeight
    sct = Form1.ScaleTop
    scl = Form1.ScaleLeft
    Form1.ScaleMode = vbPixels
    GetTwipHeight = Form1.ScaleHeight / sch
    Form1.ScaleMode = scm
    Form1.ScaleWidth = scw
    Form1.ScaleHeight = sch
    Form1.ScaleTop = sct
    Form1.ScaleLeft = scl
End Function

Private Sub DrawArena()
    Dim sunstart As Long
    Dim sunstop As Long
    Dim sunadd As Long
    Dim colr As Long
    colr = vbYellow
    If TmpOpts.Tides > 0 Then colr = RGB(255 - 255 * BouyancyScaling, 255 - 255 * BouyancyScaling, 0)
    If SimOpts.Daytime Then
        'Range calculation 0.25 + ~.75 pow 3
        sunstart = (SunPosition - (0.25 + (SunRange ^ 3) * 0.75) / 2) * SimOpts.FieldWidth
        sunstop = (SunPosition + (0.25 + (SunRange ^ 3) * 0.75) / 2) * SimOpts.FieldWidth
        If sunstart < 0 Then
            sunadd = SimOpts.FieldWidth + sunstart
            Line (sunadd, 0)-(SimOpts.FieldWidth, SimOpts.FieldHeight / 100), colr, BF
            Line (sunadd, 0)-(sunadd, SimOpts.FieldHeight), colr
            sunstart = 0
        End If
        If sunstop > SimOpts.FieldWidth Then
            sunadd = sunstop - SimOpts.FieldWidth
            Line (sunadd, 0)-(0, SimOpts.FieldHeight / 100), colr, BF
            Line (sunadd, 0)-(sunadd, SimOpts.FieldHeight), colr
            sunstop = SimOpts.FieldWidth
        End If
        Line (sunstart, 0)-(sunstop, SimOpts.FieldHeight / 100), colr, BF
        Line (sunstart, 0)-(sunstart, SimOpts.FieldHeight), colr
        Line (sunstop, 0)-(sunstop, SimOpts.FieldHeight), colr
    End If

    If MDIForm1.ZoomLock = 0 Then Exit Sub 'no need to draw boundaries if we aren't going to see them
    Line (0, 0)-(0, 0 + SimOpts.FieldHeight), vbWhite
    Line -(SimOpts.FieldWidth - 0, SimOpts.FieldHeight), vbWhite
    Line -(0 + SimOpts.FieldWidth, 0), vbWhite
    Line -(0, -0), vbWhite
End Sub

' draws memory monitor
Private Sub DrawMonitor(n As Integer)
    Dim rangered As Double
    Dim rangestart_r As Double
    Dim valred As Double
    Dim rangegreen As Double
    Dim rangestart_g As Double
    Dim valgreen As Double
    Dim rangeblue As Double
    Dim rangestart_b As Double
    Dim aspectmod As Double
    
    With frmMonitorSet
        rangered = (.Monitor_ceil_r - .Monitor_floor_r) / 255
        rangestart_r = rob(n).monitor_r - .Monitor_floor_r
        valred = rangestart_r / rangered
        If valred > 255 Then valred = 255
        If valred < 0 Then valred = 0
        
        rangegreen = (.Monitor_ceil_g - .Monitor_floor_g) / 255
        rangestart_g = rob(n).monitor_g - .Monitor_floor_g
        valgreen = rangestart_g / rangegreen
        If valgreen > 255 Then valgreen = 255
        If valgreen < 0 Then valgreen = 0
        
        rangeblue = (.Monitor_ceil_b - .Monitor_floor_b) / 255
        rangestart_b = rob(n).monitor_b - .Monitor_floor_b
        Dim valblue As Double
        valblue = rangestart_b / rangeblue
        If valblue > 255 Then valblue = 255
        If valblue < 0 Then valblue = 0
        
        aspectmod = TwipHeight / twipWidth
        Line (robManager.GetRobotPosition(n).x - robManager.GetRadius(n) * 1.1, robManager.GetRobotPosition(n).y - robManager.GetRadius(n) * 1.1 / aspectmod)-(robManager.GetRobotPosition(n).x + robManager.GetRadius(n) * 1.1, robManager.GetRobotPosition(n).y + robManager.GetRadius(n) * 1.1 / aspectmod), RGB(valred, valgreen, valblue), B
    End With
End Sub

' draws rob perimeter
Private Sub DrawRobPer(n As Integer)
  Dim Sides As Integer
  Dim t As Single
  Dim Sdlen As Single
  Dim CentreX As Long
  Dim CentreY As Long
  Dim realX As Long
  Dim realY As Long
  Dim xc As Long
  Dim yc As Long
  Dim radius As Single
  Dim Percent As Single
  Dim botDirection As Integer
  Dim Diameter As Single
  Dim topX As Single
  Dim topY As Single
  
  CentreX = robManager.GetRobotPosition(n).x
  CentreY = robManager.GetRobotPosition(n).y
  radius = robManager.GetRadius(n)
   
  If rob(n).highlight Then Circle (CentreX, CentreY), radius * 1.2, vbYellow
  If n = robfocus Then Circle (CentreX, CentreY), radius * 1.2, vbWhite
 
    FillColor = BackColor
  
      Circle (CentreX, CentreY), robManager.GetRadius(n), rob(n).color    'new line
      
    If MDIForm1.displayResourceGuagesToggle = True Then

      If rob(n).nrg > 0.5 Then
        If rob(n).nrg < 32000 Then
          Percent = rob(n).nrg / 32000
        Else
          Percent = 1
        End If
        If Percent > 0.99 Then Percent = 0.99 ' Do this because the cirlce won't draw if it goes all the way around for some reason...
        Circle (CentreX, CentreY), robManager.GetRadius(n) * 0.95, vbWhite, 0, (Percent * PI * 2#)
      End If
      
      If rob(n).body > 0.5 Then
        If rob(n).body < 32000 Then
          Percent = rob(n).body / 32000
        Else
          Percent = 1
        End If
        If Percent > 0.99 Then Percent = 0.99 ' Do this because the cirlce won't draw if it goes all the way around for some reason...
        Circle (CentreX, CentreY), robManager.GetRadius(n) * 0.9, vbMagenta, 0, (Percent * PI * 2#)
      End If
      
      If rob(n).Waste > 0.5 Then
        If rob(n).Waste < 32000 Then
          Percent = rob(n).Waste / 32000
        Else
          Percent = 1
        End If
        If Percent > 0.99 Then Percent = 0.99 ' Do this because the cirlce won't draw if it goes all the way around for some reason...
        Circle (CentreX, CentreY), robManager.GetRadius(n) * 0.85, vbGreen, 0, (Percent * PI * 2#)
      End If
      
      If rob(n).venom > 0.5 Then
        If rob(n).venom < 32000 Then
          Percent = rob(n).venom / 32000
        Else
          Percent = 1
        End If
        If Percent > 0.99 Then Percent = 0.99 ' Do this because the cirlce won't draw if it goes all the way around for some reason...
        Circle (CentreX, CentreY), robManager.GetRadius(n) * 0.8, vbBlue, 0, (Percent * PI * 2#)
      End If
      
      If rob(n).shell > 0.5 Then
        If rob(n).shell < 32000 Then
          Percent = rob(n).shell / 32000
        Else
          Percent = 1
        End If
        If Percent > 0.99 Then Percent = 0.99 ' Do this because the cirlce won't draw if it goes all the way around for some reason...
        Circle (CentreX, CentreY), robManager.GetRadius(n) * 0.75, vbRed, 0, (Percent * PI * 2#)
      End If
      
      If rob(n).Slime > 0.5 Then
        If rob(n).Slime < 32000 Then
          Percent = rob(n).Slime / 32000
        Else
          Percent = 1
        End If
        If Percent > 0.99 Then Percent = 0.99 ' Do this because the cirlce won't draw if it goes all the way around for some reason...
        Circle (CentreX, CentreY), robManager.GetRadius(n) * 0.7, vbBlack, 0, (Percent * PI * 2#)
      End If
      
       If rob(n).poison > 0.5 Then
        If rob(n).poison < 32000 Then
          Percent = rob(n).poison / 32000
        Else
          Percent = 1
        End If
        If Percent > 0.99 Then Percent = 0.99 ' Do this because the cirlce won't draw if it goes all the way around for some reason...
        Circle (CentreX, CentreY), robManager.GetRadius(n) * 0.65, vbYellow, 0, (Percent * PI * 2#)
      End If
      
      If rob(n).Vtimer > 0 Then
        Percent = rob(n).Vtimer / 100
        If Percent > 0.99 Then Percent = 0.99 ' Do this because the cirlce won't draw if it goes all the way around for some reason...
        Circle (CentreX, CentreY), robManager.GetRadius(n) * 0.6, vbCyan, 0, (Percent * PI * 2#)
      End If
      
      If rob(n).chloroplasts > 0 Then 'Panda 8/13/2013 Show how much chloroplasts a robot has
        Percent = rob(n).chloroplasts / 32000
        If Percent > 0.98 Then Percent = 0.98
        Circle (CentreX, CentreY), robManager.GetRadius(n) * 0.55, vbGreen, 0, (Percent * PI * 2#)
      End If
    End If
End Sub

' draws rob perimeter if in distance
Private Sub DrawRobDistPer(n As Integer)
  Dim CentreX As Long, CentreY As Long
  
  Dim nrgPercent As Single
  Dim bodyPercent As Single
  
  CentreX = robManager.GetRobotPosition(n).x
  CentreY = robManager.GetRobotPosition(n).y
   
  If rob(n).highlight Then Circle (CentreX, CentreY), RobSize * 2, vbYellow 'new line
  If n = robfocus Then Circle (CentreX, CentreY), RobSize * 2, vbWhite
  
  Form1.FillColor = rob(n).color

  Circle (CentreX, CentreY), robManager.GetRadius(n), rob(n).color
End Sub

' draws rob aim
Private Sub DrawRobAim(n As Integer)
  Dim x As Long, y As Long
  Dim pos As Vector
  Dim pos2 As Vector
  Dim vol As Vector
  Dim arrow1 As Vector
  Dim arrow2 As Vector
  Dim arrow3 As Vector
  Dim temp As Vector
  
  If Not rob(n).Corpse Then
    With rob(n)
  
    'We have to remember that the upper left corner is (0,0)
    pos.x = .aimvector.x
    pos.y = -.aimvector.y
       
    pos2 = VectorAdd(robManager.GetRobotPosition(n), VectorScalar(VectorUnit(pos), robManager.GetRadius(n)))
    PSet (pos2.x, pos2.y), vbWhite
    
    If MDIForm1.displayMovementVectorsToggle Then
      'Draw the voluntary movement vectors
      If .lastup <> 0 Then
        If .lastup < -1000 Then .lastup = -1000
        If .lastup > 1000 Then .lastup = 1000
        vol = VectorAdd(pos2, VectorScalar(pos, CSng(.lastup)))
        Line (pos2.x, pos2.y)-(vol.x, vol.y), .color
        
        arrow3 = VectorAdd(vol, VectorScalar(pos, 15)) ' point of the arrowhead
        temp = VectorSet(Cos(.aim - PI / 2), Sin(.aim - PI / 2))
        temp.y = -temp.y
        pos2 = VectorScalar(temp, 10)
        arrow1 = VectorAdd(vol, pos2) ' left side of arrowhead
        arrow2 = VectorSub(vol, pos2) ' right side of arrowhead
        Line (arrow1.x, arrow1.y)-(arrow3.x, arrow3.y), .color
        Line (arrow2.x, arrow2.y)-(arrow3.x, arrow3.y), .color
        Line (arrow1.x, arrow1.y)-(arrow2.x, arrow2.y), .color
      End If
      If .lastdown <> 0 Then
        If .lastdown < -1000 Then .lastdown = -1000
        If .lastdown > 1000 Then .lastdown = 1000
        pos2 = VectorSub(robManager.GetRobotPosition(n), VectorScalar(pos, robManager.GetRadius(n)))
        vol = VectorSub(pos2, VectorScalar(pos, CSng(.lastdown)))
        Line (pos2.x, pos2.y)-(vol.x, vol.y), .color
        
        arrow3 = VectorAdd(vol, VectorScalar(pos, -15)) ' point of the arrowhead
        temp = VectorSet(Cos(.aim - PI / 2), Sin(.aim - PI / 2))
        temp.y = -temp.y
        pos2 = VectorScalar(temp, 10)
        arrow1 = VectorAdd(vol, pos2) ' left side of arrowhead
        arrow2 = VectorSub(vol, pos2) ' right side of arrowhead
        Line (arrow1.x, arrow1.y)-(arrow3.x, arrow3.y), .color
        Line (arrow2.x, arrow2.y)-(arrow3.x, arrow3.y), .color
        Line (arrow1.x, arrow1.y)-(arrow2.x, arrow2.y), .color
      End If
      If .lastleft <> 0 Then
        If .lastleft < -1000 Then .lastleft = -1000
        If .lastleft > 1000 Then .lastleft = 1000
        pos = VectorSet(Cos(.aim - PI / 2), Sin(.aim - PI / 2))
        pos.y = -pos.y
        pos2 = VectorAdd(robManager.GetRobotPosition(n), VectorScalar(pos, robManager.GetRadius(n)))
        vol = VectorAdd(pos2, VectorScalar(pos, CSng(.lastleft)))
        Line (pos2.x, pos2.y)-(vol.x, vol.y), .color
        
        arrow3 = VectorAdd(vol, VectorScalar(pos, 15)) ' point of the arrowhead
        temp = .aimvector
        temp.y = -temp.y
        pos2 = VectorScalar(temp, 10)
        arrow1 = VectorAdd(vol, pos2) ' left side of arrowhead
        arrow2 = VectorSub(vol, pos2) ' right side of arrowhead
        Line (arrow1.x, arrow1.y)-(arrow3.x, arrow3.y), .color
        Line (arrow2.x, arrow2.y)-(arrow3.x, arrow3.y), .color
        Line (arrow1.x, arrow1.y)-(arrow2.x, arrow2.y), .color
      End If
      If .lastright <> 0 Then
        If .lastright < -1000 Then .lastright = -1000
        If .lastright > 1000 Then .lastright = 1000
        pos = VectorSet(Cos(.aim + PI / 2), Sin(.aim + PI / 2))
        pos.y = -pos.y
        pos2 = VectorAdd(robManager.GetRobotPosition(n), VectorScalar(pos, robManager.GetRadius(n)))
        vol = VectorAdd(pos2, VectorScalar(pos, CSng(.lastright)))
        Line (pos2.x, pos2.y)-(vol.x, vol.y), .color
        
        arrow3 = VectorAdd(vol, VectorScalar(pos, 15)) ' point of the arrowhead
        temp = .aimvector
        temp.y = -temp.y
        pos2 = VectorScalar(temp, 10)
        arrow1 = VectorAdd(vol, pos2) ' left side of arrowhead
        arrow2 = VectorSub(vol, pos2) ' right side of arrowhead
        Line (arrow1.x, arrow1.y)-(arrow3.x, arrow3.y), .color
        Line (arrow2.x, arrow2.y)-(arrow3.x, arrow3.y), .color
        Line (arrow1.x, arrow1.y)-(arrow2.x, arrow2.y), .color
      End If
    End If
    End With
  End If
End Sub

' draws skin
Private Sub DrawRobSkin(n As Integer)
  Dim x1 As Integer
  Dim x2 As Integer
  Dim y1 As Integer
  Dim y2 As Integer
  
  If rob(n).Corpse Then Exit Sub
    If rob(n).oaim <> rob(n).aim Then
      Dim t As Integer
      With rob(n)
        .OSkin(0) = (Cos(.Skin(1) / 100 - .aim) * .Skin(0)) * robManager.GetRadius(n) / 60
        .OSkin(1) = (Sin(.Skin(1) / 100 - .aim) * .Skin(0)) * robManager.GetRadius(n) / 60
        PSet (.OSkin(0) + robManager.GetRobotPosition(n).x, .OSkin(1) + robManager.GetRobotPosition(n).y)
        For t = 2 To 6 Step 2
          .OSkin(t) = (Cos(.Skin(t + 1) / 100 - .aim) * .Skin(t)) * robManager.GetRadius(n) / 60
          .OSkin(t + 1) = (Sin(.Skin(t + 1) / 100 - .aim) * .Skin(t)) * robManager.GetRadius(n) / 60
          Line -(.OSkin(t) + robManager.GetRobotPosition(n).x, .OSkin(t + 1) + robManager.GetRobotPosition(n).y), .color
        Next t
        .oaim = .aim
      End With
    Else
      With rob(n)
        PSet (.OSkin(0) + robManager.GetRobotPosition(n).x, .OSkin(1) + robManager.GetRobotPosition(n).y)
        For t = 2 To 6 Step 2
          Line -(.OSkin(t) + robManager.GetRobotPosition(n).x, .OSkin(t + 1) + robManager.GetRobotPosition(n).y), .color
        Next t
      End With
    End If
End Sub

' draws ties
Private Sub DrawRobTies(t As Integer, w As Integer, ByVal s As Integer)
  Dim k As Byte
  Dim rp As Integer
  Dim drawsmall As Integer
  Dim CentreX As Single
  Dim CentreY As Single
  Dim CentreX1 As Single
  Dim CentreY1 As Single
  
  drawsmall = w / 4
  If drawsmall = 0 Then drawsmall = 1
  
  k = 1
  With rob(t)
  CentreX = robManager.GetRobotPosition(t).x
  CentreY = robManager.GetRobotPosition(t).y
  While .Ties(k).pnt > 0
    If Not .Ties(k).back Then
      rp = .Ties(k).pnt
      CentreX1 = robManager.GetRobotPosition(rp).x
      CentreY1 = robManager.GetRobotPosition(rp).y
      DrawWidth = drawsmall
      If .Ties(k).last > 0 Then
        If w > 2 Then
          DrawWidth = w
        Else
          DrawWidth = 2
        End If
        Line (CentreX, CentreY)-(CentreX1, CentreY1), .color
      End If
    End If
    k = k + 1
  Wend
  End With
End Sub

' draws ties if ties colouring used
Private Sub DrawRobTiesCol(t As Integer, w As Integer, ByVal s As Integer)
  Dim k As Byte
  Dim col As Long
  Dim rp As Integer
  Dim drawsmall As Integer
  Dim CentreX As Single
  Dim CentreY As Single
  Dim CentreX1 As Single
  Dim CentreY1 As Single
  
  drawsmall = w / 4
  If drawsmall = 0 Then drawsmall = 1
  k = 1
  With rob(t)
  CentreX = robManager.GetRobotPosition(t).x
  CentreY = robManager.GetRobotPosition(t).y
  While .Ties(k).pnt > 0
    If Not .Ties(k).back Then
      rp = .Ties(k).pnt
      CentreX1 = robManager.GetRobotPosition(rp).x
      CentreY1 = robManager.GetRobotPosition(rp).y
      DrawWidth = drawsmall
      col = .color
      If w < 2 Then w = 2
      If .Ties(k).last > 0 Then DrawWidth = w / 2 'size of birth ties
      If .Ties(k).infused Then
        col = vbWhite
        .Ties(k).infused = False
      End If
      If .Ties(k).nrgused Then
        col = vbRed
        .Ties(k).nrgused = False
      End If
      If .Ties(k).sharing Then
        col = vbYellow
        .Ties(k).sharing = False
      End If
      
      Line (CentreX, CentreY)-(CentreX1, CentreY1), col
    End If
    k = k + 1
  Wend
  End With
End Sub

' shots...
Public Sub DrawShots()
  DrawWidth = Int(50 / (ScaleWidth / RobSize) + 1)
  If DrawWidth > 2 Then DrawWidth = 2
  Dim t As Long
  Dim s As Shot
  FillStyle = 0
  For t = 1 To ShotManager.GetMaxShot()
    s = ShotManager.GetShot(t)
    If s.flash And MDIForm1.displayShotImpactsToggle Then
       If s.shottype < 0 And s.shottype >= -7 Then
        FillColor = FlashColor(s.shottype + 10)
        Form1.Circle (s.OldPosition.x, s.OldPosition.y), 20, FlashColor(s.shottype + 10)
      Else
        FillColor = vbBlack
        Form1.Circle (s.OldPosition.x, s.OldPosition.y), 20, vbBlack
      End If
    ElseIf s.Exists And s.Stored = False Then
      PSet (s.Position.x, s.Position.y), s.color
    End If
  Next t
  FillColor = BackColor
End Sub

' main drawing procedure
Public Sub DrawAllRobs()
  Dim w As Integer
  Dim t As Integer
  Dim s As Integer
  Dim noeyeskin As Boolean
  Dim PixelsPerTwip As Single
  Dim PixRobSize As Integer
  Dim PixBorderSize As Integer
  Dim offset As Single
  Dim halfeyewidth As Single
  Dim visibleLeft As Long
  Dim visibleRight As Long
  Dim visibleTop As Long
  Dim visibleBottom As Long
  Dim low As Single
  Dim highest As Single
  Dim hi As Single
  Dim length As Single
  Dim a As Integer
  Dim r As Single

  visibleLeft = Form1.ScaleLeft
  visibleRight = Form1.ScaleLeft + Form1.ScaleWidth
  visibleTop = Form1.ScaleTop
  visibleBottom = Form1.ScaleTop + Form1.ScaleHeight
  
  twipWidth = GetTwipWidth
  TwipHeight = GetTwipHeight
  
  FortyEightOverTwipWidth = 48 / twipWidth
  FortyEightOverTwipHeight = 48 / TwipHeight
   
  noeyeskin = False
  w = Int(30 / (Form1.visiblew / RobSize) + 1)
  If (Form1.visiblew / RobSize) > 500 Then noeyeskin = True
  DrawMode = 13
  DrawStyle = 0
  DrawWidth = w
  
  If robfocus > 0 And MDIForm1.showVisionGridToggle Then
    
    length = RobSize * 12
    highest = rob(robfocus).aim + PI / 4
         
    For a = 0 To 8
            
      halfeyewidth = ((rob(robfocus).mem(EYE1WIDTH + a)) Mod 1256) / 400 ' half the eyewidth, thus /400 not /200
      While halfeyewidth > PI - PI / 36: halfeyewidth = halfeyewidth - PI: Wend
      While halfeyewidth < -PI / 36: halfeyewidth = halfeyewidth + PI: Wend
      hi = highest - (PI / 18) * a + ((rob(robfocus).mem(EYE1DIR + a) Mod 1256) / 200) + halfeyewidth  ' Display where the eye is looking
      low = hi - PI / 18 - (2 * halfeyewidth)
      
      While hi > PI * 2: hi = hi - PI * 2: Wend
      While low > PI * 2: low = low - PI * 2: Wend
      While low < 0: low = low + PI * 2: Wend
      While hi < 0: hi = hi + PI * 2: Wend
      
      If rob(robfocus).mem(EyeStart + a + 1) > 0 Then
        DrawMode = vbNotMergePen
        length = (1 / Sqr(rob(robfocus).mem(EyeStart + a + 1))) * (EyeSightDistance(AbsoluteEyeWidth(rob(robfocus).mem(EYE1WIDTH + a)), robfocus) + robManager.GetRadius(robfocus)) + robManager.GetRadius(robfocus)
        If length < 0 Then length = 0
      Else
        DrawMode = vbCopyPen
        length = EyeSightDistance(AbsoluteEyeWidth(rob(robfocus).mem(EYE1WIDTH + a)), robfocus) + robManager.GetRadius(robfocus) + robManager.GetRadius(robfocus)
      End If
            
      Circle (robManager.GetRobotPosition(robfocus).x, robManager.GetRobotPosition(robfocus).y), length, vbCyan, -low, -hi
      
      If (a = Abs(rob(robfocus).mem(FOCUSEYE) + 4) Mod 9) Then
        Circle (robManager.GetRobotPosition(robfocus).x, robManager.GetRobotPosition(robfocus).y), length, vbRed, low, hi
      End If
       
    Next
  End If

  
  DrawMode = vbCopyPen
  DrawStyle = 0
  
  
  If noeyeskin Then
    For a = 1 To MaxRobs
      If robManager.GetExists(a) Then
        r = robManager.GetRadius(a)
        If robManager.GetRobotPosition(a).x + r > visibleLeft And robManager.GetRobotPosition(a).x - r < visibleRight And robManager.GetRobotPosition(a).y + r > visibleTop And robManager.GetRobotPosition(a).y - r < visibleBottom Then
           DrawRobDistPer a
        End If
      End If
    Next a
  Else
    FillColor = BackColor
    For a = 1 To MaxRobs
      If robManager.GetExists(a) Then
         r = robManager.GetRadius(a)
         If robManager.GetRobotPosition(a).x + r > visibleLeft And robManager.GetRobotPosition(a).x - r < visibleRight And _
            robManager.GetRobotPosition(a).y + r > visibleTop And robManager.GetRobotPosition(a).y - r < visibleBottom Then
            DrawRobPer a
         End If
      End If
    Next a
  End If
  
  DrawStyle = 0
  If dispskin And Not noeyeskin Then
    For a = 1 To MaxRobs
      If robManager.GetExists(a) Then
        If robManager.GetRobotPosition(a).x + r > visibleLeft And robManager.GetRobotPosition(a).x - r < visibleRight And _
           robManager.GetRobotPosition(a).y + r > visibleTop And robManager.GetRobotPosition(a).y - r < visibleBottom Then
           DrawRobSkin a
        End If
      End If
    Next a
  End If
  
  DrawWidth = w * 2
  
  If Not noeyeskin Then
    For a = 1 To MaxRobs
     If robManager.GetExists(a) Then
       If robManager.GetRobotPosition(a).x + r > visibleLeft And robManager.GetRobotPosition(a).x - r < visibleRight And _
          robManager.GetRobotPosition(a).y + r > visibleTop And robManager.GetRobotPosition(a).y - r < visibleBottom Then
          DrawRobAim a
       End If
     End If
    Next a
  End If
  
  
  FillStyle = 1
  DrawWidth = 1
  'draw memory monitor
  If MDIForm1.MonitorOn.Checked Then
    For a = 1 To MaxRobs
      If robManager.GetExists(a) Then
         r = robManager.GetRadius(a)
         If robManager.GetRobotPosition(a).x + r > visibleLeft And robManager.GetRobotPosition(a).x - r < visibleRight And _
            robManager.GetRobotPosition(a).y + r > visibleTop And robManager.GetRobotPosition(a).y - r < visibleBottom Then
            DrawMonitor a
         End If
      End If
    Next a
  End If
  FillStyle = 0
  
End Sub

Public Sub DrawAllTies()
  Dim t As Integer
  Dim PixelsPerTwip As Single
  Dim PixRobSize As Integer
  Dim visibleLeft As Long
  Dim visibleRight As Long
  Dim visibleTop As Long
  Dim visibleBottom As Long
  
  visibleLeft = Form1.ScaleLeft
  visibleRight = Form1.ScaleLeft + Form1.ScaleWidth
  visibleTop = Form1.ScaleTop
  visibleBottom = Form1.ScaleTop + Form1.ScaleHeight
  
  PixelsPerTwip = GetTwipWidth
  PixRobSize = PixelsPerTwip * RobSize
  
  For t = 1 To MaxRobs
    If robManager.GetExists(t) Then
      If robManager.GetRobotPosition(t).x > visibleLeft And robManager.GetRobotPosition(t).x < visibleRight And _
         robManager.GetRobotPosition(t).y > visibleTop And robManager.GetRobotPosition(t).y < visibleBottom Then
         DrawRobTiesCol t, PixelsPerTwip * robManager.GetRadius(t) * 2, robManager.GetRadius(t)
      End If
    End If
  Next t
End Sub

' changes robot colour
Sub changerobcol()
  ColorForm.setcolor rob(robfocus).color
  rob(robfocus).color = ColorForm.color
End Sub

' initializes a simulation.
Sub StartSimul()
   optionsform.savesett MDIForm1.MainDir + "\settings\lastran.set" 'Botsareus 5/3/2013 Save the lastran setting

  'lets reset the autosafe data
    Open App.path & "\autosaved.gset" For Output As #1
      Write #1, False
    Close #1

Form1.camfix = False 'Botsareus 2/23/2013 When simulation starts the scren is normailized

MDIForm1.visualize = True 'Botsareus 8/31/2012 reset vedio tuggle button
MDIForm1.menuupdate

    Rnd -1
    Randomize SimOpts.UserSeedNumber / 100


    'Botsareus 5/5/2013 Update the system that the sim is running
  
    Open App.path & "\Safemode.gset" For Output As #1
      Write #1, True
    Close #1
  
    'Botsareus 4/27/2013 Create Simulation's skin
    Dim tmphsl As H_S_L
    tmphsl.h = Int(Rnd * 240)
    tmphsl.s = Int(Rnd * 60) + 180
    tmphsl.l = 222 + Int(Rnd * 2) * 6
    Dim tmprgb As R_G_B
    tmprgb = hsltorgb(tmphsl)
    chartcolor = RGB(tmprgb.r, tmprgb.g, tmprgb.b)
    tmphsl.l = tmphsl.l - 195
    tmprgb = hsltorgb(tmphsl)
    backgcolor = RGB(tmprgb.r, tmprgb.g, tmprgb.b)
    'Botsareus 6/8/2013 Overwrite skin as nessisary
    If UseOldColor Then
        chartcolor = vbWhite
        backgcolor = &H400000
    End If
  
'Botsareus 8/16/2014 Lets initate the sun

    SunPosition = 0.5
    SunRange = 1


  SimOpts.SimGUID = CLng(Rnd)
    
  Form1.Show
  Form1.ScaleWidth = SimOpts.FieldWidth
  Form1.ScaleHeight = SimOpts.FieldHeight
  Form1.visiblew = SimOpts.FieldWidth
  Form1.visibleh = SimOpts.FieldHeight
  xDivisor = 1
  yDivisor = 1
  
  If SimOpts.FieldWidth > 32000 Then xDivisor = SimOpts.FieldWidth / 32000
  If SimOpts.FieldHeight > 32000 Then yDivisor = SimOpts.FieldHeight / 32000
  
  
  MDIForm1.DontDecayNrgShots.Checked = SimOpts.NoShotDecay
  MDIForm1.DontDecayWstShots.Checked = SimOpts.NoWShotDecay
  
  MDIForm1.DisableTies.Checked = SimOpts.DisableTies
  MDIForm1.DisableArep.Checked = SimOpts.DisableTypArepro
  MDIForm1.DisableFixing.Checked = SimOpts.DisableFixing
  
  'Botsareus 4/18/2016 recording menu
  MDIForm1.SnpDeadEnable.Checked = SimOpts.DeadRobotSnp
  MDIForm1.SnpDeadExRep.Checked = SimOpts.SnpExcludeVegs
  
  SimOpts.TotBorn = 0
  MaxMem = 1000
  maxfieldsize = SimOpts.FieldWidth * 2
  robfocus = 0
  MDIForm1.DisableRobotsMenu
  
  'EricL - This is used in Shot collision as a fast way to weed out bots that could not possibily have collided with a shot
  'this cycle.  It is the maximum possible distance a bot center can be from a shot and still have had the shot impact.
  'This is the case where the bot and shot are traveling at maximum velocity in opposite directions and the shot just
  'grazes the edge of the bot.  If the shot was just about to hit the bot at the end of the last cycle, then it's distance at
  'the end of this cycle will be the hypotinose (sp?) of a right triangle C where side A is the miximum possible bot radius, and
  'side B is the sum of the maximum bot velocity and the maximum shot velocity, the latter of which can be robsize/3 + the bot
  'max velocity since bot velocity is added to shot velocity.
  MaxBotShotSeperation = Sqr((FindRadius(0, -1) ^ 2) + ((SimOpts.MaxVelocity * 2 + RobSize / 3) ^ 2))
  
  Dim t As Integer
  
  ReDim rob(500)
  
  For t = 1 To 500
    robManager.SetExists t, False
    rob(t).virusshot = 0
  Next t
  MaxRobs = 0
  Init_Buckets
  
  ShotManager.CLEAR
         
  MaxRobs = 0
  loadrobs
  If Form1.Active Then SecTimer.Enabled = True
  SimOpts.TotRunTime = 0
  If MDIForm1.visualize Then DrawAllRobs
  MDIForm1.enablesim
  


  If SimOpts.MaxEnergy > 5000 Then
    If MsgBox("Your nrg allotment is set to" + Str(SimOpts.MaxEnergy) + ".  A correct value " + _
              "is in the neighborhood of about 10 or so.  Do you want to change your energy allotment " + _
              "to 10?", vbYesNo, "Energy allotment suspicously high.") = vbYes Then
        SimOpts.MaxEnergy = 10
    End If
  End If
 
  'sim running

  main
  
End Sub

' same, but for a loaded sim
Sub startloaded()
 Dim t As Integer
 
  If tmpseed <> 0 Then
   SimOpts.UserSeedNumber = tmpseed
   TmpOpts.UserSeedNumber = tmpseed
   'Botsareus 5/8/2013 save the safemode for 'load sim'
   optionsform.savesett MDIForm1.MainDir + "\settings\lastran.set" 'Botsareus 5/3/2013 Save the lastran setting
  End If
  
    'lets reset the autosafe data
    Open App.path & "\autosaved.gset" For Output As #1
      Write #1, False
    Close #1
  
    Rnd -1
    Randomize SimOpts.UserSeedNumber / 100
  
    'Botsareus 5/5/2013 Update the system that the sim is running
  
    Open App.path & "\Safemode.gset" For Output As #1
      Write #1, True
    Close #1
  
    'Botsareus 4/27/2013 Create Simulation's skin
    Dim tmphsl As H_S_L
    tmphsl.h = Int(Rnd * 240)
    tmphsl.s = Int(Rnd * 60) + 180
    tmphsl.l = 222 + Int(Rnd * 2) * 6
    Dim tmprgb As R_G_B
    tmprgb = hsltorgb(tmphsl)
    chartcolor = RGB(tmprgb.r, tmprgb.g, tmprgb.b)
    tmphsl.l = tmphsl.l - 195
    tmprgb = hsltorgb(tmphsl)
    backgcolor = RGB(tmprgb.r, tmprgb.g, tmprgb.b)
    'Botsareus 6/8/2013 Overwrite skin as nessisary
    If UseOldColor Then
        chartcolor = vbWhite
        backgcolor = &H400000
    End If
  
  Init_Buckets
  
  Form1.ScaleWidth = SimOpts.FieldWidth
  Form1.ScaleHeight = SimOpts.FieldHeight
  Form1.visiblew = SimOpts.FieldWidth
  Form1.visibleh = SimOpts.FieldHeight
  
  xDivisor = 1
  yDivisor = 1
  If SimOpts.FieldWidth > 32000 Then xDivisor = SimOpts.FieldWidth / 32000
  If SimOpts.FieldHeight > 32000 Then yDivisor = SimOpts.FieldHeight / 32000
  
  MDIForm1.visualize = True
  Active = True
  MaxMem = 1000
  maxfieldsize = SimOpts.FieldWidth * 2
  robfocus = 0
  MDIForm1.DisableRobotsMenu
  
  'EricL - This is used in Shot collision as a fast way to weed out bots that could not possibily have collided with a shot
  'this cycle.  It is the maximum possible distance a bot center can be from a shot and still have had the shot impact.
  'This is the case where a bot with 32000 body and a shot are traveling at maximum velocity in opposite directions and the shot just
  'grazes the edge of the bot.  If the shot was just about to hit the bot at the end of the last cycle, then it's distance at
  'the end of this cycle will be the hypotinoose (sp?) of a right triangle ABC where side A is the maximum possible bot radius, and
  'side B is the sum of the maximum bot velocity and the maximum shot velocity, the latter of which can be robsize/3 + the bot
  'max velocity since bot velocity is added to shot velocity.
  MaxBotShotSeperation = Sqr((FindRadius(0, -1) ^ 2) + ((SimOpts.MaxVelocity * 2 + RobSize / 3) ^ 2))
    
  MDIForm1.DontDecayNrgShots.Checked = SimOpts.NoShotDecay
  MDIForm1.DontDecayWstShots.Checked = SimOpts.NoWShotDecay
  
  MDIForm1.DisableTies.Checked = SimOpts.DisableTies
  MDIForm1.DisableArep.Checked = SimOpts.DisableTypArepro
  MDIForm1.DisableFixing.Checked = SimOpts.DisableFixing
  
  'Botsareus 4/18/2016 recording menu
  MDIForm1.SnpDeadEnable.Checked = SimOpts.DeadRobotSnp
  MDIForm1.SnpDeadExRep.Checked = SimOpts.SnpExcludeVegs
  
  MDIForm1.AutoFork.Checked = SimOpts.EnableAutoSpeciation
    
  SecTimer.Enabled = True
  'setfeed
  If MDIForm1.visualize Then DrawAllRobs
  MDIForm1.enablesim
  Me.Visible = True
    
  Vegs.cooldown = -SimOpts.RepopCooldown
  totnvegsDisplayed = -1 ' Just set this to -1 for the first cycle so the cost low water mark doesn't trigger.
  totvegs = -1 ' Set to -1 to avoid veggy reproduction on first cycle
  
  main
End Sub

' loads robot DNA at the beginning of a simulation
Private Sub loadrobs()
  Dim k As Integer
  Dim a As Integer
  Dim i As Integer
  Dim cc As Integer, t As Integer
  k = 0
  For cc = 1 To SimOpts.SpeciesNum
    For t = 1 To SimOpts.Specie(k).qty
      a = RobScriptLoad(respath(SimOpts.Specie(k).path) + "\" + SimOpts.Specie(k).Name)
      If a < 0 Then
        t = SimOpts.Specie(k).qty
        GoTo bypassThisSpecies
      End If
      rob(a).Veg = SimOpts.Specie(k).Veg
      rob(a).NoChlr = SimOpts.Specie(k).NoChlr
      rob(a).Fixed = SimOpts.Specie(k).Fixed
      If rob(a).Fixed Then rob(a).mem(216) = 1
      Dim pos As Vector
      pos.x = Random(SimOpts.Specie(k).Poslf * CSng(SimOpts.FieldWidth - 60#), SimOpts.Specie(k).Posrg * CSng(SimOpts.FieldWidth - 60#))
      pos.y = Random(SimOpts.Specie(k).Postp * CSng(SimOpts.FieldHeight - 60#), SimOpts.Specie(k).Posdn * CSng(SimOpts.FieldHeight - 60#))
      robManager.SetRobotPosition a, pos
      
      rob(a).nrg = SimOpts.Specie(k).Stnrg
      rob(a).body = 1000
      
      robManager.SetRadius a, FindRadius(a)
      
      rob(a).mem(SetAim) = rob(a).aim * 200
      If rob(a).Veg Then rob(a).chloroplasts = StartChlr 'Botsareus 2/12/2014 Start a robot with chloroplasts
      rob(a).Dead = False
            
      rob(a).Mutables = SimOpts.Specie(k).Mutables
      
      For i = 0 To 7 'Botsareus 5/20/2012 fix for skin engine
        rob(a).Skin(i) = SimOpts.Specie(k).Skin(i)
      Next i
      
      rob(a).color = SimOpts.Specie(k).color
      rob(a).mem(timersys) = Random(-32000, 32000)
      rob(a).CantSee = SimOpts.Specie(k).CantSee
      rob(a).DisableDNA = SimOpts.Specie(k).DisableDNA
      rob(a).DisableMovementSysvars = SimOpts.Specie(k).DisableMovementSysvars
      rob(a).CantReproduce = SimOpts.Specie(k).CantReproduce
      rob(a).VirusImmune = SimOpts.Specie(k).VirusImmune
      rob(a).virusshot = 0
      rob(a).Vtimer = 0
      rob(a).genenum = CountGenes(rob(a).dna)
      
      rob(a).DnaLen = DnaLen(rob(a).dna())
      rob(a).GenMut = rob(a).DnaLen / GeneticSensitivity 'Botsareus 4/9/2013 automatically apply genetic to inserted robots
      
      rob(a).mem(DnaLenSys) = rob(a).DnaLen
      rob(a).mem(GenesSys) = rob(a).genenum
    Next t
bypassThisSpecies:
    k = k + 1
    MDIForm1.Caption = "Loading... " & Int((cc - 1) * 100 / SimOpts.SpeciesNum) & "% Please wait..."
  Next cc
  MDIForm1.Caption = MDIForm1.BaseCaption
  
End Sub

' calls main form status bar update
Public Sub cyccaption(ByVal num As Single)
  MDIForm1.infos num, TotalRobotsDisplayed, totnvegsDisplayed, TotalChlr, SimOpts.TotBorn, SimOpts.TotRunCycle, SimOpts.TotRunTime  'Botsareus 8/25/2013 Mod to send TotalChlr
End Sub

' which rob has been clicked?
Private Function whichrob(x As Single, y As Single) As Integer
  Dim dist As Double, pist As Double
  Dim t As Integer
  
'Botsareus 6/23/2016 Need to offset the robots by (actual velocity minus velocity) before drawing them
Dim pos As Vector
For t = 1 To MaxRobs
    If robManager.GetExists(t) Then
        pos = robManager.GetRobotPosition(t)
        pos = VectorAdd(VectorSub(pos, robManager.GetVelocity(t)), robManager.GetActualVelocity(t))
        robManager.SetRobotPosition t, pos
    End If
Next
  
  whichrob = 0
  dist = 10000
  For t = 1 To MaxRobs
    If robManager.GetExists(t) Then
      pist = Abs(robManager.GetRobotPosition(t).x - x) ^ 2 + Abs(robManager.GetRobotPosition(t).y - y) ^ 2
      If Abs(robManager.GetRobotPosition(t).x - x) < robManager.GetRadius(t) And Abs(robManager.GetRobotPosition(t).y - y) < robManager.GetRadius(t) And pist < dist And robManager.GetExists(t) Then
        whichrob = t
        dist = pist
      End If
    End If
  Next t

For t = 1 To MaxRobs
    If robManager.GetExists(t) Then
        pos = robManager.GetRobotPosition(t)
        pos = VectorSub(VectorAdd(pos, robManager.GetVelocity(t)), robManager.GetActualVelocity(t))
        robManager.SetRobotPosition t, pos
    End If
Next
End Function

' stuff for clicking, dragging, etc
' move+click: drags robot if one selected, else drags screen
Private Sub Form_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
'Botsareus 7/2/2014 Added exitsub to awoid bugs
  Dim st As Long
  Dim sl As Long
  Dim vsv As Long
  Dim hsv As Long
  Dim t As Byte
  Dim vel As Vector
  
  visibleh = Int(Form1.ScaleHeight)
  If button = 0 Then
    MouseClickX = x
    MouseClickY = y
  End If
  
  If button = 1 And Not MDIForm1.insrob Then
    If MouseClicked Then
      st = ScaleTop + MouseClickY - y
      sl = ScaleLeft + MouseClickX - x
      If st < 0 And MDIForm1.ZoomLock.value = 0 Then
        st = 0
        MouseClickY = y
      End If
      If sl < 0 And MDIForm1.ZoomLock.value = 0 Then
        sl = 0
        MouseClickX = x
      End If
      If st > SimOpts.FieldHeight - visibleh And MDIForm1.ZoomLock.value = 0 Then
        st = SimOpts.FieldHeight - visibleh
      End If
      If sl > SimOpts.FieldWidth - visiblew And MDIForm1.ZoomLock.value = 0 Then
        sl = SimOpts.FieldWidth - visiblew
      End If
      Form1.ScaleTop = st
      Form1.ScaleLeft = sl
      Form1.Refresh
      Redraw
      Form1.Refresh
      Exit Sub
    End If
  End If
  
  If button = 1 And robfocus > 0 And DraggingBot Then
    vel = VectorSub(robManager.GetRobotPosition(robfocus), VectorSet(x, y))
    robManager.SetRobotPosition robfocus, VectorSet(x, y)
    robManager.SetVelocity robfocus, VectorSet(0, 0)
    robManager.SetActualVelocity robfocus, VectorSet(0, 0)
    Dim a As Byte
    For a = 1 To tmprob_c
        robManager.SetRobotPosition tmppos(a).n, VectorSet(x - tmppos(a).x, y - tmppos(a).y)
        robManager.SetVelocity tmppos(a).n, VectorSet(0, 0)
        robManager.SetActualVelocity tmppos(a).n, VectorSet(0, 0)
    Next
    If Not Active Then Redraw
    Exit Sub
  End If
End Sub

' it seems that there's no simple way to know the mouse status
' outside of a Form event. So I've used the event to switch
' on and off some global vars
Private Sub Form_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
  If lblSafeMode.Visible Then Exit Sub 'Botsareus 5/13/2013 Safemode restrictions
  
  MouseClicked = False
  ZoomFlag = False ' EricL - stop zooming in!
  DraggingBot = False
End Sub

Private Sub Form_Resize()
  If Form1.WindowState = 2 Then
    SimOpts.FieldWidth = Form1.ScaleWidth
    SimOpts.FieldHeight = Form1.ScaleHeight
    maxfieldsize = SimOpts.FieldWidth * 2
  End If
End Sub

' double clicking on a rob pops up the info window
' elsewhere closes it
Private Sub Form_DblClick()
'elcrasho = 50 'Botsareudnotdone debug only

  Dim n As Integer
  Dim m As Integer
  n = whichrob(CSng(MouseClickX), CSng(MouseClickY))

  If n > 0 Then
    robfocus = n
    MDIForm1.EnableRobotsMenu
    If Not rob(n).highlight Then deletemark
    datirob.Visible = True
    datirob.RefreshDna
    DoEvents 'fixo?
    datirob.infoupdate n, rob(n).nrg, rob(n).parent, rob(n).Mutations, rob(n).age, rob(n).SonNumber, 1, rob(n).FName, rob(n).genenum, rob(n).LastMut, rob(n).generation, rob(n).DnaLen, rob(n).LastOwner, rob(n).Waste, rob(n).body, rob(n).mass, rob(n).venom, rob(n).shell, rob(n).Slime, rob(n).chloroplasts
  Else
    datirob.Form_Unload 0
    robfocus = 0
    MDIForm1.DisableRobotsMenu
    If ActivForm.Visible Then ActivForm.NoFocus
  End If
End Sub

' clicking outside robots closes info window
' mind the walls tool
Private Sub Form_Click()
If lblSafeMode.Visible Then Exit Sub 'Botsareus 5/13/2013 Safemode restrictions

  Dim n As Integer
   
  'LogForm.Visible = False
  n = whichrob(CSng(MouseClickX), CSng(MouseClickY))
  
  If n = 0 Then
    robfocus = 0
    MDIForm1.DisableRobotsMenu
    If ActivForm.Visible Then ActivForm.NoFocus
  End If
End Sub


'Botsareus 11/29/2013 You can now move whole organism
Private Sub gen_tmp_pos_lst()
Dim a As Integer
Dim rst As Boolean
rst = True
tmprob_c = 0
    For a = 1 To MaxRobs
      If robManager.GetExists(a) And rob(a).highlight Then
        If a = robfocus Then rst = False
        tmprob_c = tmprob_c + 1
        If tmprob_c < 51 Then
            tmppos(tmprob_c).n = a
            tmppos(tmprob_c).x = robManager.GetRobotPosition(robfocus).x - robManager.GetRobotPosition(a).x
            tmppos(tmprob_c).y = robManager.GetRobotPosition(robfocus).y - robManager.GetRobotPosition(a).y
        Else
            tmprob_c = 50
        End If
      End If
    Next a
If rst Then tmprob_c = 0
End Sub

' clicking (well, half-clicking) on a robot selects it
' clicking outside can add a robot if we're in robot insertion
' mode.
Private Sub Form_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
If lblSafeMode.Visible Then Exit Sub 'Botsareus 5/13/2013 Safemode restrictions

  Dim n As Integer
  Dim k As Integer
  Dim m As Integer
  
  'EricL - Zooms in on the position of the mouse when the scroll button is held.
  If button = 4 Then
    ZoomFlag = True
    While ZoomFlag = True
      If Form1.visiblew > 100 And Form1.visibleh > 100 Then
        Form1.visiblew = Form1.visiblew / 1.05
        Form1.visibleh = Form1.visibleh / 1.05
        Form1.ScaleHeight = Form1.visibleh
        Form1.ScaleWidth = Form1.visiblew
        Form1.ScaleTop = y - Form1.ScaleHeight / 2
        Form1.ScaleLeft = x - Form1.ScaleWidth / 2
      End If
        Form1.Redraw
        DoEvents
     Wend
  End If
    
  n = whichrob(x, y)
  
  If n = 0 Then
    DraggingBot = False
    If Not Form1.SecTimer.Enabled Then datirob.Visible = False 'Botsareus 1/5/2013 Small fix to do with wrong data displayed in robot info, auto hide the window
  Else
    DraggingBot = True
    gen_tmp_pos_lst
  End If
  
  If button = 2 Then
    robfocus = n
    If n > 0 Then
      MDIForm1.PopupMenu MDIForm1.popup
      MDIForm1.EnableRobotsMenu
    End If
  End If
  
  If n > 0 Then
    robfocus = n
    MDIForm1.EnableRobotsMenu
    If Not rob(n).highlight Then deletemark
  End If
  
  If n = 0 And button = 1 And MDIForm1.insrob Then
   If Not lblSaving.Visible Then 'Botsareus 9/6/2014 Bug fix
    k = 0
    While SimOpts.Specie(k).Name <> "" And SimOpts.Specie(k).Name <> MDIForm1.Combo1.Text
      k = k + 1
    Wend
    
    If SimOpts.Specie(k).path = "Invalid Path" Then
      MsgBox ("The path for this robot is invalid.")
    'ElseIf Not SimOpts.Specie(k).Native Then
    '   MsgBox ("Sorry, but you can't insert a species which did not originate in this simulation.")
    Else
      aggiungirob k, x, y
    End If
   End If
  End If
  
  If button = 1 And Not MDIForm1.insrob And n = 0 Then
    MouseClicked = True
  End If
  
  Redraw
End Sub

' deletes the yellow highlight mark around robots
Private Sub deletemark()
  Dim t As Integer
  For t = 1 To MaxRobs
    rob(t).highlight = False
  Next t
End Sub

' seconds timer, used to periodically check _cycles_ counters
' expecially for internet transfers
' tricky!
Private Sub SecTimer_Timer()

  If Enabled = False Then Exit Sub '9/7/2013 Do not count the timer if form is disabled

  Static TenSecondsAgo(10) As Long
  Static LastDownload As Long
  Static LastMut As Long
  Static MutPhase As Integer
  Static LastShuffle As Long
  Dim TitLog As String
  Dim t As Integer
  Dim i As Integer
  
  SimOpts.TotRunTime = SimOpts.TotRunTime + 1

  ' reset counters if simulation restarted
  If SimOpts.TotRunTime = 1 Then
    For i = 0 To 9
      TenSecondsAgo(i) = SimOpts.TotRunCycle
    Next i
    ' reset the counter for horiz/vertical shuffle
    LastShuffle = 0
  End If
  
  ' same as above, but checking totruncycle<lastcycle instead
  If SimOpts.TotRunCycle < TenSecondsAgo((SimOpts.TotRunTime + 9) Mod 10) Then
    For i = 0 To 9
      TenSecondsAgo(i) = SimOpts.TotRunCycle
    Next i
    LastShuffle = 0
  End If
  
  ' if we've had 5000 cycles in a second, probably we've
  ' loaded a saved sim. So we need to reset some counters
  If SimOpts.TotRunCycle - TenSecondsAgo((SimOpts.TotRunTime + 9) Mod 10) > 5000 Then
    For i = 0 To 9
      TenSecondsAgo(i) = SimOpts.TotRunCycle
    Next i
    ' facciamo avvenire uno shuffle fra 50000 cicli
    LastShuffle = SimOpts.TotRunCycle - 50000
  End If
  
  ' update status bar in MDI formMod
  
  TenSecondsAgo(SimOpts.TotRunTime Mod 10) = SimOpts.TotRunCycle
  SimOpts.CycSec = CSng(CSng(CSng(SimOpts.TotRunCycle) - CSng(TenSecondsAgo((SimOpts.TotRunTime + 1) Mod 10))) * 0.1)
    
  cyccaption SimOpts.CycSec

  '(provides the mutation rates oscillation Botsareus 8/3/2013 moved to UpdateSim)
    
  'Botsareus 7/13/2012 calls update function for main icon menu
  MDIForm1.menuupdate
End Sub

' main procedure. Oh yes!
Private Sub main()
  Dim clocks As Long
  Dim i As Integer
  Dim b As Integer
  
  'clocks = GetTickCount
  Do
    If Active Then
      
      If MDIForm1.ignoreerror = True Then
        On Error Resume Next
      Else
        'On Error GoTo SaveError  'EricL - Uncomment this line in compiled version! error.sim
      End If
      
      Form1.SecTimer.Enabled = True
           
      UpdateSim
      MDIForm1.Follow ' 11/29/2013 zoom follow selected robot
          
      ' redraws all:
      If MDIForm1.visualize Then
        Form1.Label1.Visible = False
        If Not MDIForm1.oneonten Then
          Redraw
        Else
          If SimOpts.TotRunCycle Mod 10 = 0 Then Redraw
        End If
      End If
    
      If datirob.Visible And Not datirob.ShowMemoryEarlyCycle Then
        With rob(robfocus)
          datirob.infoupdate robfocus, .nrg, .parent, .Mutations, .age, .SonNumber, 1, .FName, .genenum, .LastMut, .generation, .DnaLen, .LastOwner, .Waste, .body, .mass, .venom, .shell, .Slime, .chloroplasts
        End With
      End If
            
    End If
    DoEvents
    If MDIForm1.limitgraphics = True Then
      clocks = GetTickCount
      If (GetTickCount - clocks) < 67 Then
        While (GetTickCount - clocks < 67)
            
        Wend
      End If
    End If
    
    If Not camfix Then
        MDIForm1.fixcam 'Botsareus 2/23/2013 normalizes screen
        camfix = True
        Form1.Active = Not pausefix 'Botsareus 3/6/2013 allowes starting a simulation paused
        Form1.SecTimer.Enabled = Not pausefix
    End If
  Loop
  Exit Sub
  
SaveError:
  MsgBox "Error. " + Err.Description + ".  Saving sim in saves directory as error.sim"
  
tryagain:
  On Error GoTo missingdir
  SaveSimulation MDIForm1.MainDir + "\saves\error.sim"
missingdir:
 If Err.Number = 76 Then
    b = MsgBox("Cannot find the Saves Directory to save error.sim.  " + vbCrLf + _
            "Would you like me to create one?   " + vbCrLf + vbCrLf + _
            "If this is a new install, choose OK.", vbOKCancel + vbQuestion)
    If b = vbOK Then
       MkDir (MDIForm1.MainDir + "\saves")
       GoTo tryagain
    Else
      MsgBox ("Exiting without saving error.sim.")
    End If
  End If
  
  Erase rob
  End
End Sub

'
' M A N A G E M E N T   A N D   R E P R O D U C T I O N
'

'should remove highlight on a robot whenever the simulation is restarted after a pause
Public Sub unfocus()
  Dim t As Integer
  For t = 1 To MaxRobs
    rob(t).highlight = False
  Next t
End Sub


'
'   S E R V I C E S
'

' recursively calculates the number of alive descendants of a rob
Public Function discendenti(t As Integer, disce As Integer) As Integer
  Dim k As Integer
  Dim n As Integer
  Dim rid As Long
  rid = rob(t).AbsNum
  If disce < 1000 Then
    For n = 1 To MaxRobs
      If robManager.GetExists(n) Then
        If rob(n).parent = rid Then
          discendenti = discendenti + 1
          disce = disce + 1
          discendenti = discendenti + discendenti(n, disce)
        End If
      End If
    Next n
  End If
End Function

' returns the fittest robot (selected through the score function)
' altered from the bot with the most generations
' to the bot with the most invested energy in itself and children
Function fittest() As Integer
'Botsareus 5/22/2013 Lets figure out what we are searching for
Dim sPopulation As Double
Dim sEnergy As Double
sEnergy = (IIf(intFindBestV2 > 100, 100, intFindBestV2)) / 100
sPopulation = (IIf(intFindBestV2 < 100, 100, 200 - intFindBestV2)) / 100

  Dim t As Integer
  Dim s As Double
  Dim Mx As Double
  Mx = 0
  For t = 1 To MaxRobs
    If robManager.GetExists(t) And Not rob(t).Veg And Not rob(t).FName = "Corpse" Then
      TotalOffspring = 1
      s = score(t, 1, 10, 0) + rob(t).nrg + rob(t).body * 10 'Botsareus 5/22/2013 Advanced fit test
      If s < 0 Then s = 0 'Botsareus 9/23/2016 Bug fix
      s = (TotalOffspring ^ sPopulation) * (s ^ sEnergy)

      If s >= Mx Then
        Mx = s
        fittest = t
      End If
    End If
  Next t
  
    'Z E R O B O T
    'Pass result of fittest back to evo

End Function

' does various things: with
' tipo=0 returns the number of descendants for maxrec generations
' tipo=1 highlights the descendants
' tipo=2 searches up the tree for eldest ancestor, then down again
' tipo=3 draws the lines showing kinship relations
Function score(ByVal r As Integer, ByVal reclev As Integer, maxrec As Integer, tipo As Integer) As Double
  Dim al As Integer
  Dim dx As Single
  Dim dy As Single
  Dim cr As Long
  Dim ct As Long
  Dim t As Integer
  If tipo = 2 Then plines (r)
  score = 0
  For t = 1 To MaxRobs
    If robManager.GetExists(t) Then
      If rob(t).parent = rob(r).AbsNum Then
        If reclev < maxrec Then score = score + score(t, reclev + 1, maxrec, tipo)
        If tipo = 0 Then score = score + InvestedEnergy(t) 'Botsareus 8/3/2012 generational distance code
        If tipo = 4 And reclev > p_reclev Then p_reclev = reclev
        If tipo = 1 Then rob(t).highlight = True
        If tipo = 3 Then
          dx = (robManager.GetRobotPosition(r).x - robManager.GetRobotPosition(t).x) / 2
          dy = (robManager.GetRobotPosition(r).y - robManager.GetRobotPosition(t).y) / 2
          cr = RGB(128, 128, 128)
          ct = vbWhite
          If rob(r).AbsNum > rob(t).AbsNum Then
            cr = vbWhite
            ct = RGB(128, 128, 128)
          End If
          Line (robManager.GetRobotPosition(t).x, robManager.GetRobotPosition(t).y)-Step(dx, dy), ct
          Line -(robManager.GetRobotPosition(r).x, robManager.GetRobotPosition(r).y), cr
        End If
      End If
    End If
  Next t
  If tipo = 1 Then
    Form1.Cls
    DrawAllRobs
  End If
End Function

Function InvestedEnergy(t As Integer) As Double 'Botsareus 5/22/2013 Calculate both population and energy
  InvestedEnergy = rob(t).nrg + rob(t).body * 10 'botschange fittest
  TotalOffspring = TotalOffspring + 1 'botschange fittest
End Function

' goes up the tree searching for eldest ancestor
Private Sub plines(ByVal t As Integer)
  Dim p As Integer
  p = parent(t)
  While p > 0
    t = p
    p = parent(t)
  Wend
  If p = 0 Then p = t
  t = score(p, 1, 1000, 3)
End Sub

' returns the robot's parent
Function parent(r As Integer) As Integer
  Dim t As Integer
  parent = 0
  For t = 1 To MaxRobs
    If rob(t).AbsNum = rob(r).parent And robManager.GetExists(t) Then parent = t
  Next t
End Function
