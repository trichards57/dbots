VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGset 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Global Settings"
   ClientHeight    =   5700
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   10764
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   10764
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   8160
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   9480
      TabIndex        =   0
      Top             =   5280
      Width           =   1215
   End
   Begin TabDlg.SSTab tb 
      Height          =   5175
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10755
      _ExtentX        =   18965
      _ExtentY        =   9123
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Main settings"
      TabPicture(0)   =   "frmGset.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chkIntRnd"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkEpiGene"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ffmInitChlr"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ffmFBSBO"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkSafeMode"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ffmMainDir"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ffmCheatin"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ffmUI"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Mutations"
      TabPicture(1)   =   "frmGset.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ffmEpiReset"
      Tab(1).Control(1)=   "ffmSunMut"
      Tab(1).ControlCount=   2
      Begin VB.Frame ffmUI 
         Caption         =   "UI Settings"
         Height          =   2000
         Left            =   120
         TabIndex        =   53
         Top             =   360
         Width           =   4695
         Begin VB.CheckBox chkOldColor 
            Caption         =   "Use old simulation colors"
            Height          =   375
            Left            =   240
            TabIndex        =   59
            Top             =   480
            Width           =   3075
         End
         Begin VB.CheckBox chkNoBoyMsg 
            Caption         =   "Don't display Buoyancy Warning"
            Height          =   255
            Left            =   240
            TabIndex        =   58
            Top             =   810
            Width           =   2775
         End
         Begin VB.CheckBox chkNoVid 
            Caption         =   "Turn off Video when simulation starts"
            Height          =   255
            Left            =   240
            TabIndex        =   57
            Top             =   1080
            Width           =   3495
         End
         Begin VB.CheckBox chkGraphUp 
            Caption         =   "Automatically assume to keep updating graphs"
            Height          =   255
            Left            =   240
            TabIndex        =   56
            Top             =   1365
            Width           =   3975
         End
         Begin VB.CheckBox chkScreenRatio 
            Caption         =   "Fix Screen Ratio when simulation starts"
            Height          =   255
            Left            =   240
            TabIndex        =   55
            Top             =   270
            Width           =   3075
         End
         Begin VB.CheckBox chkHide 
            Caption         =   "Hide Darwinbots on restart mode in system tray"
            Height          =   255
            Left            =   240
            TabIndex        =   54
            Top             =   1650
            Width           =   3975
         End
      End
      Begin VB.Frame ffmCheatin 
         Caption         =   "Cheating Prevention"
         Height          =   1575
         Left            =   120
         TabIndex        =   49
         Top             =   2375
         Width           =   4695
         Begin VB.TextBox txtBodyFix 
            Height          =   375
            Left            =   1200
            TabIndex        =   51
            Text            =   "32100"
            Top             =   480
            Width           =   1455
         End
         Begin VB.CheckBox chkGreedy 
            Caption         =   "Kill robots that are excessively greedy to there kids, using them to dump there energy."
            Height          =   375
            Left            =   240
            TabIndex        =   50
            Top             =   960
            Width           =   3615
         End
         Begin VB.Label CheatinLab 
            Caption         =   "Kill robots that have more then this amound of body to prevent BigBerthas:"
            Height          =   615
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Randomization"
         Height          =   1095
         Left            =   120
         TabIndex        =   46
         Top             =   3960
         Width           =   4695
         Begin VB.CheckBox chkchseedstartnew 
            Caption         =   "Generate new seed when you click 'Start new'"
            Height          =   375
            Left            =   240
            TabIndex        =   48
            Top             =   240
            Width           =   4275
         End
         Begin VB.CheckBox chkchseedloadsim 
            Caption         =   "Generate new seed when you click 'Load Simulation'"
            Height          =   375
            Left            =   240
            TabIndex        =   47
            Top             =   600
            Width           =   4275
         End
      End
      Begin VB.Frame ffmMainDir 
         Caption         =   "Main Directory"
         Height          =   1215
         Left            =   4920
         TabIndex        =   43
         Top             =   480
         Width           =   5655
         Begin VB.CheckBox chkUseCD 
            Caption         =   "Change Directory"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtCD 
            Height          =   375
            Left            =   120
            TabIndex        =   44
            Text            =   "Text1"
            Top             =   720
            Width           =   5415
         End
      End
      Begin VB.CheckBox chkSafeMode 
         Caption         =   "Use Safe Mode"
         Height          =   255
         Left            =   5160
         TabIndex        =   42
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Frame ffmFBSBO 
         Caption         =   "Find Best Settings base on:"
         Height          =   915
         Left            =   4920
         TabIndex        =   38
         Top             =   2280
         Width           =   5655
         Begin MSComctlLib.Slider sldFindBest 
            Height          =   570
            Left            =   1320
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   240
            Width           =   2955
            _ExtentX        =   5207
            _ExtentY        =   1016
            _Version        =   393216
            LargeChange     =   40
            Max             =   200
            TickStyle       =   2
            TickFrequency   =   10
         End
         Begin VB.Label lblIE 
            Caption         =   "Invested Energy"
            Height          =   195
            Left            =   4300
            TabIndex        =   41
            Top             =   420
            Width           =   1200
         End
         Begin VB.Label lblTP 
            Caption         =   "Total Population"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   420
            Width           =   1215
         End
      End
      Begin VB.Frame ffmEpiReset 
         Caption         =   "Epigenetic Reset"
         Height          =   735
         Left            =   -73440
         TabIndex        =   32
         Top             =   460
         Width           =   7575
         Begin VB.CheckBox chkEpiReset 
            Caption         =   "Periodically reset Epigenetic memory"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   300
            Width           =   2895
         End
         Begin VB.TextBox txtMEmp 
            Height          =   375
            Left            =   4920
            TabIndex        =   34
            Text            =   "1.3"
            ToolTipText     =   $"frmGset.frx":0038
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtOP 
            Height          =   375
            Left            =   6840
            TabIndex        =   33
            Text            =   "17"
            ToolTipText     =   "This is how much exponential amplified mutations a robot must have in order to trigger an epigenetic memory reset."
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblMEmp 
            Caption         =   "mutation amplification:"
            Height          =   255
            Left            =   3240
            TabIndex        =   37
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblOP 
            Caption         =   "overload point:"
            Height          =   255
            Left            =   5640
            TabIndex        =   36
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame ffmSunMut 
         Caption         =   "Sunline Mutations"
         Height          =   3775
         Left            =   -74160
         TabIndex        =   9
         Top             =   1280
         Width           =   9135
         Begin VB.CheckBox chkSunbelt 
            Caption         =   "Enable Point2, CopyError2, Amplification, and Translocation Mutations"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   300
            Width           =   5295
         End
         Begin VB.PictureBox picIcon 
            AutoSize        =   -1  'True
            ClipControls    =   0   'False
            Height          =   432
            Left            =   120
            Picture         =   "frmGset.frx":00F0
            ScaleHeight     =   263.118
            ScaleMode       =   0  'User
            ScaleWidth      =   263.118
            TabIndex        =   21
            Top             =   3120
            Width           =   432
         End
         Begin VB.CheckBox chkDelta2 
            Caption         =   "Enable Delta2 Mutations"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtMainExp 
            Height          =   285
            Left            =   4995
            TabIndex        =   19
            Top             =   960
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtMainLn 
            Height          =   285
            Left            =   6240
            TabIndex        =   18
            Top             =   960
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtDevExp 
            Height          =   285
            Left            =   5280
            TabIndex        =   17
            Top             =   1440
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtDevLn 
            Height          =   285
            Left            =   6480
            TabIndex        =   16
            Top             =   1440
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtPMinter 
            Height          =   285
            Left            =   4320
            TabIndex        =   15
            Top             =   1920
            Width           =   615
         End
         Begin VB.CheckBox chkNorm 
            Caption         =   "Normalize default mutation rates and slowest possible rate based on DNA length"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   2400
            Width           =   6015
         End
         Begin VB.TextBox txtDnalen 
            Height          =   285
            Left            =   3600
            TabIndex        =   13
            Top             =   2760
            Width           =   615
         End
         Begin VB.TextBox txtMxDnalen 
            Height          =   285
            Left            =   6240
            TabIndex        =   12
            Top             =   2760
            Width           =   615
         End
         Begin VB.TextBox txtWTC 
            Height          =   285
            Left            =   6960
            TabIndex        =   11
            ToolTipText     =   "Oscillation between type and value"
            Top             =   1920
            Width           =   615
         End
         Begin MSComctlLib.Slider sldMain 
            Height          =   285
            Left            =   7080
            TabIndex        =   10
            ToolTipText     =   "Probability Delta2 Chance"
            Top             =   960
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3408
            _ExtentY        =   508
            _Version        =   393216
            Max             =   100
            TickFrequency   =   10
         End
         Begin MSComctlLib.Slider sldDev 
            Height          =   285
            Left            =   7080
            TabIndex        =   23
            ToolTipText     =   "Mean/Stddev Delta2 Chance"
            Top             =   1440
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3408
            _ExtentY        =   508
            _Version        =   393216
            Max             =   100
            TickFrequency   =   10
         End
         Begin VB.Label lblExplMut 
            Caption         =   $"frmGset.frx":0D32
            Height          =   600
            Left            =   720
            TabIndex        =   31
            Top             =   3100
            Width           =   8175
         End
         Begin VB.Label lblMmain 
            Caption         =   "Probablity:    Multiply by (10 ^ ± 1/   XXXX   ) ± Liner"
            Height          =   375
            Left            =   2520
            TabIndex        =   30
            Top             =   960
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.Label lblMmean 
            Caption         =   "Mean/Stddev:    Multiply by (10 ^ ± 1/   XXXX   ) ± Liner"
            Height          =   375
            Left            =   2520
            TabIndex        =   29
            Top             =   1440
            Visible         =   0   'False
            Width           =   4335
         End
         Begin VB.Label lblPMDelta2 
            Caption         =   "Delta2 cycle interval for Point mutations:"
            Height          =   495
            Left            =   2520
            TabIndex        =   28
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label lblDnalen 
            Caption         =   "DNA length *    XXXX     Slowest rate DNA length * "
            Height          =   255
            Left            =   2520
            TabIndex        =   27
            Top             =   2760
            Width           =   4455
         End
         Begin VB.Label lblWTC 
            Caption         =   "Delta2 for value(s) of what to change:     ±"
            Height          =   495
            Left            =   5280
            TabIndex        =   26
            ToolTipText     =   "Oscillation between type and value"
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label lblChance 
            Caption         =   "Chance of mutation %"
            Height          =   255
            Left            =   7200
            TabIndex        =   25
            Top             =   720
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label lblD2 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Click here for some really advanced delta2 settings."
            Height          =   255
            Left            =   2640
            TabIndex        =   24
            Top             =   1200
            Width           =   6135
         End
      End
      Begin VB.Frame ffmInitChlr 
         Caption         =   "Advanced Chloroplast Options"
         Height          =   975
         Left            =   4920
         TabIndex        =   6
         Top             =   3480
         Width           =   5655
         Begin VB.TextBox txtStartChlr 
            Height          =   375
            Left            =   2400
            TabIndex        =   7
            Text            =   "0"
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblStartChlr 
            Caption         =   "Start Repopulating Robots with  XXXXXXXXXXXXX  Chloroplasts"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   435
            Width           =   5295
         End
      End
      Begin VB.CheckBox chkEpiGene 
         Caption         =   "Save epigenetic memory as a temporary gene when saving robot DNA to file.        Warning: This uses delgene."
         Height          =   375
         Left            =   5520
         TabIndex        =   5
         Top             =   4560
         Width           =   4815
      End
      Begin VB.CheckBox chkIntRnd 
         Caption         =   "Use Internet Pictures to seed Randomizer"
         Height          =   255
         Left            =   7080
         TabIndex        =   4
         Top             =   1920
         Width           =   3375
      End
   End
   Begin VB.Label lbl 
      Caption         =   "Note: To reset all values delete global.gset file from your main directory."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   7455
   End
End
Attribute VB_Name = "frmGset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Botsareus 3/15/2013 The global settings form
Dim robpath As String

Private Sub btnCancel_Click()
'cancel has been pressed
Unload Me
End Sub

Private Sub btnOK_Click()
  'Botsareus 5/10/2013 Make sure txtCD points to a valid directory
  If Not FolderExists(txtCD.Text) And chkUseCD.value = 1 Then
    MsgBox "Please use a valid directory for the main directory.", vbCritical
    Exit Sub
  End If

  If chkUseCD.value = 0 Then
   'delete the maindir setting if no longer used
    If dir(App.path & "\Maindir.gset") <> "" Then Kill (App.path & "\Maindir.gset")
  Else
   'write current path to setting
    Open App.path & "\Maindir.gset" For Output As #1
    Write #1, txtCD.Text
    Close #1
  End If

  Dim specielist As String
  specielist = ""
  Dim i As Integer

  'prompt that settings will take place when you restart db
  MsgBox "Global settings will take effect the next time DarwinBots starts.", vbInformation

  'save all settings
  Open MDIForm1.MainDir & "\Global.gset" For Output As #1
  Write #1, chkScreenRatio = 1
  Write #1, val(txtBodyFix)
  Write #1, chkGreedy = 1
  Write #1, chkchseedstartnew = 1
  Write #1, chkchseedloadsim = 1
  Write #1, chkSafeMode = 1
  Write #1, sldFindBest.value
  Write #1, chkOldColor = 1
  Write #1, chkNoBoyMsg = 1
  Write #1, chkNoVid = 1
  Write #1, chkEpiReset = 1
  Write #1, val(txtMEmp)
  Write #1, val(txtOP)
  Write #1, chkSunbelt = 1
  Write #1, chkDelta2 = 1
  Write #1, val(txtMainExp)
  Write #1, val(txtMainLn)
  Write #1, val(txtDevExp)
  Write #1, val(txtDevLn)
  Write #1, val(txtPMinter)
  Write #1, chkNorm = 1
  Write #1, val(txtDnalen)
  Write #1, val(txtMxDnalen)
  Write #1, val(txtWTC)
  Write #1, val(sldMain)
  Write #1, val(sldDev)
  Write #1, False
  
  Dim tmpopt As OptionButton
  
  Write #1, val(txtStartChlr.Text)
  
  'Restrictions
  
  Write #1, chkGraphUp.value = 1
  Write #1, chkHide.value = 1
  
  'New way to preserve epigenetic memory
  
  Write #1, chkEpiGene.value = 1
  
  'Use internet as randomizer
  
  Write #1, chkIntRnd.value = 1
    
  Close #1

  'unload
  Unload Me
End Sub


Private Sub chkDelta2_Click()
  If chkDelta2.value = 1 And Visible Then
    If MsgBox("Enabling Delta2 mutations may slow down your simulation considerably as it may optimize for extreme mutation rates. Are you sure?", vbExclamation + vbYesNo, "Global Darwinbots Settings") = vbNo Then chkDelta2.value = 0 'Botsareus 9/13/2014 Warnings for Shvarz
  End If
  lblD2.Visible = chkDelta2.value = 1
  lblPMDelta2.Visible = chkDelta2.value = 1
  txtPMinter.Visible = chkDelta2.value = 1
  lblWTC.Visible = chkDelta2.value = 1
  txtWTC.Visible = chkDelta2.value = 1
'
  If chkDelta2.value = 0 Then
    lblMmain.Visible = False
    lblMmean.Visible = False
    txtMainExp.Visible = False
    txtMainLn.Visible = False
    txtDevExp.Visible = False
    txtDevLn.Visible = False
    lblChance.Visible = False
    sldMain.Visible = False
    sldDev.Visible = False
  End If
End Sub

Private Sub chkEpiReset_Click()
  txtMEmp.Enabled = chkEpiReset.value = 1
  txtOP.Enabled = chkEpiReset.value = 1
End Sub


Private Sub chkIntRnd_Click()
  If chkIntRnd.value = 1 And Visible Then
     MsgBox "The pictures are loaded from urls listed in the " & App.path & "\web.gset file.", vbInformation
  End If
End Sub

Private Sub chkNorm_Click()
  txtDnalen.Visible = chkNorm.value = 1
  lblDnalen.Visible = chkNorm.value = 1
  txtMxDnalen.Visible = chkNorm.value = 1
End Sub

Private Sub chkUseCD_Click()
  If chkUseCD.value = 1 Then
    If Visible Then MsgBox "If you are running parallel simulations on a single computer, make sure you disable this setting or make the path is unique for each instance. Also, don't forget to have each Darwin.exe in a separate directory"
    txtCD.Enabled = True
  Else
    txtCD.Enabled = False
  End If
End Sub

Private Sub Form_Load()
  'load all global settings into controls
  chkScreenRatio.value = IIf(screenratiofix, 1, 0)
  txtBodyFix = bodyfix
  chkGreedy = IIf(reprofix, 1, 0)
  chkchseedstartnew.value = IIf(chseedstartnew, 1, 0)
  chkchseedloadsim.value = IIf(chseedloadsim, 1, 0)
  chkNoBoyMsg.value = IIf(loadboylabldisp, 1, 0) 'some global settings change within simulation
  chkNoVid.value = IIf(loadstartnovid, 1, 0) 'some global settings change within simulation
  txtCD = MDIForm1.MainDir
  'only eanable txtCD and chkUseCD if maindir.gset exisits
  If dir(App.path & "\Maindir.gset") <> "" Then
    chkUseCD.value = 1
    txtCD.Enabled = True
  Else
    chkUseCD.value = 0
    txtCD.Enabled = False
  End If
  'are we using safemode
  chkSafeMode = IIf(UseSafeMode, 1, 0)
  'find best
  sldFindBest.value = intFindBestV2
  'use old color
  chkOldColor = IIf(UseOldColor, 1, 0)
  'epigenetic reset
  chkEpiReset = IIf(epireset, 1, 0)
  txtMEmp = epiresetemp
  txtOP = epiresetOP
  txtMEmp.Enabled = chkEpiReset.value = 1
  txtOP.Enabled = chkEpiReset.value = 1
  'Eclipse mutations
  chkSunbelt.value = IIf(sunbelt, 1, 0)
  'Delta2
  chkDelta2.value = IIf(Delta2, 1, 0)
  txtMainExp = DeltaMainExp
  txtMainLn = DeltaMainLn
  txtDevExp = DeltaDevExp
  txtDevLn = DeltaDevLn
  txtPMinter = DeltaPM
  txtWTC = DeltaWTC
  sldMain = DeltaMainChance
  sldDev = DeltaDevChance
  'Norm Mut
  chkNorm = IIf(NormMut, 1, 0)
  txtDnalen = valNormMut
  txtMxDnalen = valMaxNormMut
  'Set values Delta2 and Norm mut
  lblD2.Visible = chkDelta2.value = 1
  lblPMDelta2.Visible = chkDelta2.value = 1
  txtPMinter.Visible = chkDelta2.value = 1
  lblWTC.Visible = chkDelta2.value = 1
  txtWTC.Visible = chkDelta2.value = 1
  '
  txtDnalen.Visible = chkNorm.value = 1
  lblDnalen.Visible = chkNorm.value = 1
  txtMxDnalen.Visible = chkNorm.value = 1
  '
  
  '
  txtStartChlr.Text = StartChlr
  '
  chkGraphUp.value = IIf(GraphUp, 1, 0)
  chkHide.value = IIf(HideDB, 1, 0)
  '
  chkEpiGene = IIf(UseEpiGene, 1, 0)
End Sub

Private Sub lblD2_Click()
  lblMmain.Visible = True
  lblMmean.Visible = True
  txtMainExp.Visible = True
  txtMainLn.Visible = True
  txtDevExp.Visible = True
  txtDevLn.Visible = True
  lblChance.Visible = True
  sldMain.Visible = True
  sldDev.Visible = True
  lblD2.Visible = False
End Sub

Private Sub txtBodyFix_LostFocus()
  'make sure the value is sane
  txtBodyFix = Abs(val(txtBodyFix))
  If txtBodyFix > 32100 Then txtBodyFix = 32100
  If txtBodyFix < 2500 Then
    If MsgBox("It is not recommended to set 'Cheating Prevention' below 2500 because it may result in your robots never getting enough body to survive. Do you want to set to 2500 instead?", vbExclamation + vbYesNo, "Global Darwinbots Settings") = vbYes Then txtBodyFix = 2500 'Botsareus 9/13/2014 Warnings for Shvarz
  End If
  If txtBodyFix < 1000 Then txtBodyFix = 1000
End Sub

Private Sub txtDevLn_LostFocus()
  'make sure the value is sane
  txtDevLn = Abs(val(txtDevLn))
  If txtDevLn > 5000 Then txtDevLn = 3000
End Sub

Private Sub txtDnalen_LostFocus()
  'make sure the value is sane
  txtDnalen = Abs(val(txtDnalen))
  If txtDnalen < 1 Then txtDnalen = 1
  If txtDnalen > 2000 Then txtDnalen = 2000
End Sub

Private Sub txtMainExp_LostFocus()
  'make sure the value is sane
  txtMainExp = Abs(val(txtMainExp))
  If txtMainExp = 0 Then Exit Sub
  If txtMainExp < 0.4 Then txtMainExp = 0.4
  If txtMainExp > 25 Then txtMainExp = 25
End Sub

Private Sub txtDevExp_LostFocus()
  'make sure the value is sane
  txtDevExp = Abs(val(txtDevExp))
  If txtDevExp = 0 Then Exit Sub
  If txtDevExp < 0.4 Then txtDevExp = 0.4
  If txtDevExp > 25 Then txtDevExp = 25
End Sub

Private Sub txtMainLn_LostFocus()
  'make sure the value is sane
  txtMainLn = Round(Abs(val(txtMainLn)))
  If txtMainLn > 5000 Then txtMainLn = 3000
End Sub

Private Sub txtMEmp_LostFocus()
  'make sure the value is sane
  txtMEmp = Abs(val(txtMEmp))
  If txtMEmp > 5 Then txtMEmp = 5
End Sub

Private Sub txtMxDnalen_LostFocus()
  'make sure the value is sane
  txtMxDnalen = Abs(val(txtMxDnalen))
  If txtMxDnalen < 1 Then txtMxDnalen = 1
  If txtMxDnalen > 32000 Then txtMxDnalen = 32000
End Sub

Private Sub txtOP_LostFocus()
  'make sure the value is sane
  txtOP = Abs(val(txtOP))
  If txtOP > 32000 Then txtOP = 32000
End Sub

Private Sub txtPMinter_LostFocus()
  'make sure the value is sane
  txtPMinter = Round(Abs(val(txtPMinter)))
  If txtPMinter > 32000 Then txtPMinter = 32000
End Sub

Private Sub txtStartChlr_LostFocus()
  'make sure the value is sane
  txtStartChlr = Abs(val(txtStartChlr))
  If txtStartChlr > 32000 Then txtStartChlr = 32000
End Sub

Private Sub txtWTC_Change()
  'make sure the value is sane
  txtWTC = Abs(val(txtWTC))
  If txtWTC > 100 Then txtWTC = 100
End Sub
