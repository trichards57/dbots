VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.MDIForm MDIForm1 
   AutoShowChildren=   0   'False
   BackColor       =   &H00400000&
   Caption         =   "DarwinBots"
   ClientHeight    =   6504
   ClientLeft      =   3636
   ClientTop       =   2652
   ClientWidth     =   14292
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":08CA
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1000"
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2520
      Top             =   2160
      _ExtentX        =   995
      _ExtentY        =   995
      _Version        =   393216
   End
   Begin VB.Timer unpause 
      Interval        =   200
      Left            =   4200
      Top             =   1560
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   1440
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":095C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0EF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1490
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1A2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1FC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":255E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2AF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3092
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":362C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3BC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4160
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":46FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4C94
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":556E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":63C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":695A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":6AB4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   396
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14292
      _ExtentX        =   25210
      _ExtentY        =   699
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   24
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "newsim"
            Object.ToolTipText     =   "Start a new simulation"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "loadsim"
            Object.ToolTipText     =   "Load a previously saved simulation"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "savesim"
            Object.ToolTipText     =   "Save the current simulation"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   500
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "play"
            Object.ToolTipText     =   "Play simulation"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cycle"
            Object.ToolTipText     =   "Calculate a single cycle"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "stop"
            Object.ToolTipText     =   "Pause simulation"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   300
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "limit"
            Object.ToolTipText     =   "Limits speed to 15 cycles/sec"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "fast"
            Object.ToolTipText     =   "Toggle fast mode"
            ImageIndex      =   7
            Style           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Flickermode"
            Object.ToolTipText     =   "Toggle Flickermode. Faster but flickery."
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "noskin"
            ImageIndex      =   17
            Style           =   1
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nopoff"
            Object.ToolTipText     =   "Turns off bot explosion poffs"
            ImageIndex      =   13
            Style           =   1
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Novideo"
            Object.ToolTipText     =   "Turns off graphical display for the utmost speed."
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   300
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "insert"
            Object.ToolTipText     =   "Toggle robot insertion mode"
            ImageIndex      =   8
            Style           =   1
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1500
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "best"
            Object.ToolTipText     =   "Find the most successful robot"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "graph"
            Object.ToolTipText     =   "Select graph to display"
            ImageIndex      =   11
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   23
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "pop"
                  Text            =   "Population graph"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "avgmut"
                  Text            =   "Average mutations"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "avgmutlen"
                  Text            =   "Average mutations/DNA len"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "avgage"
                  Text            =   "Average Age"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "avgsons"
                  Text            =   "Average Descendents"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "avgnrg"
                  Text            =   "Average Energy"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "avglen"
                  Text            =   "Average DNA length"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "avgcond"
                  Text            =   "Average DNA conditions"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "simnrg"
                  Text            =   "Total Energy"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "autocost"
                  Text            =   "Dynamic Cost Stats"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "speciesdiversity"
                  Text            =   "Species Diversity"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "avgchlr"
                  Text            =   "Average Chloroplasts"
               EndProperty
               BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu14 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "maxgeneticdistance"
                  Text            =   "(Selective) Genetic Distance"
               EndProperty
               BeginProperty ButtonMenu15 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "maxgenerationaldistance"
                  Text            =   "Generational Distance"
               EndProperty
               BeginProperty ButtonMenu16 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "simplegeneticdistance"
                  Text            =   "(Slow) Genetic Distance"
               EndProperty
               BeginProperty ButtonMenu17 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu18 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CG1"
                  Text            =   "Customizable Graph 1"
               EndProperty
               BeginProperty ButtonMenu19 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CG2"
                  Text            =   "Customizable Graph 2"
               EndProperty
               BeginProperty ButtonMenu20 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CG3"
                  Text            =   "Customizable Graph 3"
               EndProperty
               BeginProperty ButtonMenu21 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu22 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "resgraph"
                  Text            =   "Reset all graphs"
               EndProperty
               BeginProperty ButtonMenu23 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "listgraphs"
                  Text            =   "List all running graphs"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "snapshot"
            Object.ToolTipText     =   "capture data on all living robots"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   300
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stealth"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ignore"
            Object.ToolTipText     =   "The program will attempt to ignore errors"
            ImageIndex      =   10
         EndProperty
      EndProperty
      Begin VB.CheckBox SunButton 
         DownPicture     =   "MDIForm1.frx":6C0E
         Height          =   375
         Left            =   10200
         Picture         =   "MDIForm1.frx":6FB6
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Toggles the Sun"
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox F1InternetButton 
         BackColor       =   &H00E0E0E0&
         DownPicture     =   "MDIForm1.frx":734D
         Height          =   375
         Left            =   10680
         Picture         =   "MDIForm1.frx":76BF
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Toggles Internet Mode"
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox ZoomLock 
         DownPicture     =   "MDIForm1.frx":7A31
         Height          =   375
         Left            =   9000
         Picture         =   "MDIForm1.frx":7D73
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Locks and unlocks being able to lock at areas outside of the arena."
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox F1Piccy 
         Height          =   375
         Left            =   11160
         Picture         =   "MDIForm1.frx":7E55
         ScaleHeight     =   324
         ScaleWidth      =   324
         TabIndex        =   5
         Top             =   0
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         IntegralHeight  =   0   'False
         ItemData        =   "MDIForm1.frx":8225
         Left            =   4800
         List            =   "MDIForm1.frx":8227
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Select the species for robot insertion"
         Top             =   25
         Width           =   1530
      End
      Begin VB.CommandButton czo 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   9720
         Picture         =   "MDIForm1.frx":8229
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Zoom out"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton czin 
         Height          =   375
         Left            =   9360
         Picture         =   "MDIForm1.frx":87B3
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Zoom in"
         Top             =   0
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   6225
      Width           =   14295
      _ExtentX        =   25210
      _ExtentY        =   466
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   13
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1418
            MinWidth        =   1411
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1418
            MinWidth        =   1411
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1418
            MinWidth        =   1411
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1418
            MinWidth        =   1411
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1418
            MinWidth        =   1411
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1418
            MinWidth        =   1411
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1418
            MinWidth        =   1411
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1418
            MinWidth        =   1411
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1418
            MinWidth        =   1411
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1418
            MinWidth        =   1411
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1418
            MinWidth        =   1411
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1418
            MinWidth        =   1411
            Object.ToolTipText     =   "The ratio of the energy this cycle to the average"
         EndProperty
         BeginProperty Panel13 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1778
            MinWidth        =   1764
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Index           =   1
      NegotiatePosition=   3  'Right
      Begin VB.Menu newsim 
         Caption         =   "New Simulation"
         Index           =   2
         Shortcut        =   {F1}
      End
      Begin VB.Menu loadsim 
         Caption         =   "Load Simulation"
         Index           =   3
         Shortcut        =   {F2}
      End
      Begin VB.Menu SaveSim 
         Caption         =   "Save Simulation"
         Shortcut        =   {F3}
      End
      Begin VB.Menu SaveSimWithoutMutations 
         Caption         =   "Save Sim Without Mutations"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu pause 
         Caption         =   "Pause Sim (F12)"
      End
      Begin VB.Menu p 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu quit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Begin VB.Menu inssp 
         Caption         =   "Insert Organism..."
         Shortcut        =   ^I
      End
      Begin VB.Menu sep11 
         Caption         =   "-"
      End
      Begin VB.Menu Species 
         Caption         =   "Species..."
         Shortcut        =   ^X
      End
      Begin VB.Menu fisica 
         Caption         =   "General Settings..."
         Shortcut        =   ^P
      End
      Begin VB.Menu costi 
         Caption         =   "Physics and Costs..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu moltiplicatore 
         Caption         =   "Global Mutation Rates..."
         Shortcut        =   ^G
      End
      Begin VB.Menu Leagues 
         Caption         =   "Restart and Leagues..."
         Shortcut        =   ^F
      End
      Begin VB.Menu intOptionsOpen 
         Caption         =   "Internet Options..."
      End
      Begin VB.Menu sep90 
         Caption         =   "-"
      End
      Begin VB.Menu DisableFixing 
         Caption         =   "Disallow robot from fixing or unfixing"
      End
      Begin VB.Menu DisableArep 
         Caption         =   "Disable asexual reproduction for non-repopulating robots"
      End
      Begin VB.Menu sspp 
         Caption         =   "-"
      End
      Begin VB.Menu Autotag 
         Caption         =   "Automatically tag by name..."
      End
      Begin VB.Menu AutoFork 
         Caption         =   "Enable Automatic Forking"
      End
      Begin VB.Menu SepEYE 
         Caption         =   "-"
      End
      Begin VB.Menu showEyeDesign 
         Caption         =   "Go to eye designer..."
      End
      Begin VB.Menu Sep81 
         Caption         =   "-"
      End
      Begin VB.Menu RESOver 
         Caption         =   "Restriction Overwrites"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu Backgrounds 
      Caption         =   "View"
      Begin VB.Menu LoadPiccy 
         Caption         =   "Import Background Picture"
      End
      Begin VB.Menu removepiccy 
         Caption         =   "Remove Background Picture"
      End
      Begin VB.Menu sep987 
         Caption         =   "-"
      End
      Begin VB.Menu ShowVisionGrid 
         Caption         =   "Show Vision Grid"
         Checked         =   -1  'True
      End
      Begin VB.Menu DisplayMovementVectors 
         Caption         =   "Display Movement Vectors"
         Checked         =   -1  'True
      End
      Begin VB.Menu DisplayShotImpacts 
         Caption         =   "Display Shot Impacts"
         Checked         =   -1  'True
      End
      Begin VB.Menu DisplayResourceGuages 
         Caption         =   "Display Resource Guages"
         Checked         =   -1  'True
      End
      Begin VB.Menu sep98 
         Caption         =   "-"
      End
      Begin VB.Menu MonitorSettings 
         Caption         =   "Settings for RGB Memory Monitor..."
      End
      Begin VB.Menu MonitorOn 
         Caption         =   "RGB Memory Monitor"
      End
   End
   Begin VB.Menu popup 
      Caption         =   "Robot"
      Begin VB.Menu robinf 
         Caption         =   "Show Robot Info"
         Shortcut        =   ^R
      End
      Begin VB.Menu par 
         Caption         =   "Show Philogeny"
         Shortcut        =   ^T
      End
      Begin VB.Menu mutrat 
         Caption         =   "Mutation Rates"
         Shortcut        =   ^M
      End
      Begin VB.Menu col 
         Caption         =   "Change Color"
         Shortcut        =   ^C
      End
      Begin VB.Menu Sep 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu cons 
         Caption         =   "Open Console"
         Shortcut        =   ^Y
      End
      Begin VB.Menu genact 
         Caption         =   "View Gene Activations"
         Shortcut        =   ^Z
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu ucci 
         Caption         =   "Kill Robot"
         Shortcut        =   ^K
      End
      Begin VB.Menu sdna 
         Caption         =   "Save Robot's DNA"
         Shortcut        =   ^S
      End
      Begin VB.Menu makenewspecies 
         Caption         =   "Make New Species"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu selorg 
         Caption         =   "Select Entire Organism"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu saveorg 
         Caption         =   "Save Entire Organism"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu killorg 
         Caption         =   "Kill Entire Organism"
         Shortcut        =   ^E
      End
      Begin VB.Menu sep99 
         Caption         =   "-"
      End
      Begin VB.Menu fittest 
         Caption         =   "Find Best"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu Objects 
      Caption         =   "Objects"
      Begin VB.Menu ShotsMenu 
         Caption         =   "Shots"
         Begin VB.Menu DontDecayNrgShots 
            Caption         =   "Don't decay nrg shots"
            Checked         =   -1  'True
         End
         Begin VB.Menu DontDecayWstShots 
            Caption         =   "Don't decay waste shots"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu TiesMenu 
         Caption         =   "Ties"
         Begin VB.Menu DisableTies 
            Caption         =   "Disable Ties"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu TeleporterMenu 
         Caption         =   "Teleporters"
         Begin VB.Menu NewTeleportMenu 
            Caption         =   "New Teleporter..."
         End
         Begin VB.Menu HighLightTeleportersMenu 
            Caption         =   "Hilight Teleporters"
            Checked         =   -1  'True
         End
         Begin VB.Menu DeleteTeleporterMenu 
            Caption         =   "Delete Teleporter"
         End
         Begin VB.Menu DeleteTeleportersMenu 
            Caption         =   "Delete All Teleporters"
         End
      End
      Begin VB.Menu sep47 
         Caption         =   "Shapes"
         Begin VB.Menu shapes 
            Caption         =   "New Shape..."
         End
         Begin VB.Menu AddTenObstacles 
            Caption         =   "Add Ten Random Shapes"
         End
         Begin VB.Menu DeleteTenObstacles 
            Caption         =   "Delete 10 Random Shapes"
         End
         Begin VB.Menu DeleteShape 
            Caption         =   "Delete Shape"
         End
         Begin VB.Menu DeleteAllShapes 
            Caption         =   "Delete All Shapes"
         End
      End
      Begin VB.Menu MakeMaze 
         Caption         =   "Mazes"
         Begin VB.Menu HorizontalMaze 
            Caption         =   "Simple Horizontal Maze"
         End
         Begin VB.Menu VerticaMaze 
            Caption         =   "Simple Vertical Maze"
         End
         Begin VB.Menu DrawSpiral 
            Caption         =   "Sprial"
         End
         Begin VB.Menu CheckerMaze 
            Caption         =   "Checkerboard"
         End
         Begin VB.Menu PolarIce 
            Caption         =   "Polar Ice"
         End
         Begin VB.Menu TrashCompactor 
            Caption         =   "Trash Compactor"
         End
      End
   End
   Begin VB.Menu EGrid 
      Caption         =   "E-Grid"
      Index           =   45
      Visible         =   0   'False
      Begin VB.Menu EGridEnabled 
         Caption         =   "Enable"
         Checked         =   -1  'True
      End
      Begin VB.Menu GridSize 
         Caption         =   "Grid Size"
         Begin VB.Menu EGridLarge 
            Caption         =   "Large"
            Checked         =   -1  'True
         End
         Begin VB.Menu EGridMedium 
            Caption         =   "Medium"
            Checked         =   -1  'True
         End
         Begin VB.Menu EGridSmall 
            Caption         =   "Small"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu interent 
      Caption         =   "Internet"
      Begin VB.Menu F1Internet 
         Caption         =   "Enabled"
      End
      Begin VB.Menu EditIntTeleporter 
         Caption         =   "Internet Teleporter..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu recording 
      Caption         =   "Recording"
      Begin VB.Menu SafemodeBkp 
         Caption         =   "Safemode Backup"
      End
      Begin VB.Menu SnpLiving 
         Caption         =   "Snapshot of the living"
      End
      Begin VB.Menu SnpDead 
         Caption         =   "Snapshot of the dead"
         Begin VB.Menu SnpDeadEnable 
            Caption         =   "Enable recording"
         End
         Begin VB.Menu SnpDeadExRep 
            Caption         =   "Exclude repopulating robots"
         End
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Index           =   5
      NegotiatePosition=   3  'Right
      Begin VB.Menu DNAexp 
         Caption         =   "DNA Help"
         Shortcut        =   ^D
      End
      Begin VB.Menu sep102 
         Caption         =   "-"
      End
      Begin VB.Menu RobTagInfo 
         Caption         =   "Robot Tag Information"
      End
      Begin VB.Menu y_info 
         Caption         =   "Survival Info"
         Visible         =   0   'False
      End
      Begin VB.Menu sep101 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "About"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu TrayIconPopup 
      Caption         =   "TrayIconPopup"
      Visible         =   0   'False
      Begin VB.Menu ShowInfo 
         Caption         =   "Show Information"
         Begin VB.Menu CyclesNumber 
            Caption         =   "CyclesNumber"
         End
         Begin VB.Menu MutationValue 
            Caption         =   "MutationValue"
         End
         Begin VB.Menu TotalBots 
            Caption         =   "TotalBots"
         End
         Begin VB.Menu NumberBots 
            Caption         =   "NumberBots"
         End
         Begin VB.Menu NumberVeg 
            Caption         =   "NumberVeg"
         End
         Begin VB.Menu BotsBorn 
            Caption         =   "BotsBorn"
         End
      End
      Begin VB.Menu ShowDB 
         Caption         =   "Show DarwinBots"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' DarwinBots - copyright 2003 Carlo Comis
' Modifications by Purple Youko and Numsgil - 2004, 2005
' Post V2.42 modifications copyright (c) 2006, 2007 Eric Lockard  eric@sulaadventures.com
'
' Post V2.45 modifications copyright (c) 2012, 2013, 2014, 2015, 2016 Paul Kononov
'a.k.a
'______________________________________________1$$$___108033_____$$______________________________
'____1$$$$$$$3________________011_______________$$__$$$$$$$$$$8_1$_________1$$$1__8$$$1_______3$$
'____3$$811$$$0______________1$$3_______________0___$$$$__1$$$8____________0$$$__1$$$$______0$$8_
'____1$$__1$$$1______________$$$_______880_________3$$$3__8$$$3____________0$$0__1$$$0____1$$$1__
'____3$$11$$$0____8$$$8___8$$$$$$$3__38$___________0$$$$_$$$$$_____________$$$8__8$$$1____$$$8___
'____1$$$$$$1____$$$$$$$3__$$$$88___$$8____________8$$$$$$$$0______________8$$$__$$$$_____3$$$1__
'____1$$$$$$$0__$$$$_18$$___8$$___1$$$_____________$$$$$$$$8_______________0$$8_0$$$$_______8$$__
'____3$$$88$$$$_$$$___8$$__3$$8____0$$$___________1$$$$1$$$$_______________8$$$$$$$$3_______$$$1_
'____1$$____3$$_8$$__8$$8__8$$______0$$1__________3$$$0_1$$$1______________8$$$$$$$$_______$$$0__
'____3$$____0$$__$$$$$$$1__$$8______$$0___________$$$$___8$$0______________1$$$$$$$3_____1$$$0___
'_____$$011$$$$__8$$$$$1__1$$1____3$$1___________3$$$3___3$$$_______________8$$$$$3____1$$$0_____
'____3$$$$$$81____3$80____$$81__188______________8$$$_____8$$0_______________0$$81____8003_______
'
' All rights reserved.
'
'Redistribution and use in source and binary forms, with or without
'modification, are permitted provided that:
'
'(1) source code distributions retain the above copyright notice and this
'    paragraph in its entirety,
'(2) distributions including binary code include the above copyright notice and
'    this paragraph in its entirety in the documentation or other materials
'    provided with the distribution, and
'(3) Without the agreement of the author redistribution of this product is only allowed
'    in non commercial terms and non profit distributions.
'
'THIS SOFTWARE IS PROVIDED ``AS IS'' AND WITHOUT ANY EXPRESS OR IMPLIED
'WARRANTIES, INCLUDING, WITHOUT LIMITATION, THE IMPLIED WARRANTIES OF
'MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE.

Option Explicit
'Botsareus 5/19/2012 Update to the way the tool menu looks; removed unnecessary pictures from collection
'Botsareus 3/15/2013 got rid of screen save code (was broken)
'Botsareus 4/17/2013 Added a bunch of new components

Private lockswitch As Byte
Public MainDir As String
Public BaseCaption As String
Public insrob As Boolean
Dim AspettaFlag As Boolean
Public visualize As Boolean   ' video output on/off
Public oneonten As Boolean    ' fast mode on/off
Public nopoff As Boolean      ' don't anim deaths with particles
Public limitgraphics As Boolean
Public ignoreerror As Boolean
Public xc As Long
Public yc As Long
Public showVisionGridToggle As Boolean
Public displayShotImpactsToggle As Boolean
Public displayResourceGuagesToggle As Boolean
Public displayMovementVectorsToggle As Boolean
Public SaveWithoutMutations As Boolean

Sub wait(n As Byte)
  Dim e As Long
  e = Timer
  Do
    DoEvents
  Loop Until ((e + n) Mod 86400) < Timer And IIf((e + n) > 86400, Timer < 100, True)
End Sub

Private Sub AutoFork_Click() 'Botsareus 3/23/2014 auto forking
  On Error GoTo b:
  AutoFork.Checked = Not AutoFork.Checked
  If AutoFork.Checked Then SimOpts.SpeciationGeneticDistance = InputBox("Enter % of mutations to DNA length that constitutes forking", "Automatic Forking", SimOpts.SpeciationGeneticDistance)
  SimOpts.EnableAutoSpeciation = AutoFork.Checked
  Exit Sub
b:
  AutoFork.Checked = False
End Sub

Private Sub AutoTag_Click()
    Dim t As Integer
    Dim robname As String
    Dim robtag As String
  
    robname = InputBox("Please enter the exact name of the robot you wish append with a new tag.")
    If robname = "" Then
        MsgBox "Cancel or blank entry", vbCritical
        Exit Sub
    End If
  
    robtag = InputBox("Please enter the new tag. It can not be more then 45 characters long.")
    robtag = Left(replacechars(robtag), 45)
  
    For t = 1 To MaxRobs
        If robManager.GetExists(t) And rob(t).FName = robname Then
            rob(t).tag = robtag
        End If
    Next t
End Sub


Private Sub DisableArep_Click() 'Botsareus 4/17/2013 The new disable asexrepro button
  DisableArep.Checked = Not DisableArep.Checked
  SimOpts.DisableTypArepro = DisableArep.Checked
  TmpOpts.DisableTypArepro = DisableArep.Checked
End Sub

Private Sub DisableFixing_Click()
  DisableFixing.Checked = Not DisableFixing.Checked
  SimOpts.DisableFixing = DisableFixing.Checked
  TmpOpts.DisableFixing = DisableFixing.Checked
End Sub

Private Sub DisableTies_Click()
  DisableTies.Checked = Not DisableTies.Checked
  SimOpts.DisableTies = DisableTies.Checked
  TmpOpts.DisableTies = DisableTies.Checked
End Sub

Private Sub DisplayMovementVectors_Click()
  DisplayMovementVectors.Checked = Not DisplayMovementVectors.Checked
  displayMovementVectorsToggle = DisplayMovementVectors.Checked
End Sub

Private Sub DisplayResourceGuages_Click()
  DisplayResourceGuages.Checked = Not DisplayResourceGuages.Checked
  displayResourceGuagesToggle = DisplayResourceGuages.Checked
End Sub

Private Sub DisplayShotImpacts_Click()
  DisplayShotImpacts.Checked = Not DisplayShotImpacts.Checked
  displayShotImpactsToggle = DisplayShotImpacts.Checked
End Sub

Private Sub DNAexp_Click() 'Botsareus 8/7/2012 help loads a while, added a hourglass
  MousePointer = vbHourglass
  DNA_Help.Show
  MousePointer = vbDefault
End Sub

Private Sub DontDecayNrgShots_Click()
  DontDecayNrgShots.Checked = Not DontDecayNrgShots.Checked
  SimOpts.NoShotDecay = DontDecayNrgShots.Checked
  TmpOpts.NoShotDecay = DontDecayNrgShots.Checked
End Sub

Private Sub DontDecayWstShots_Click() 'Botsareus 9/28/2013 Don't decay waste shots
  DontDecayWstShots.Checked = Not DontDecayWstShots.Checked
  SimOpts.NoWShotDecay = DontDecayWstShots.Checked
  TmpOpts.NoWShotDecay = DontDecayWstShots.Checked
End Sub


Private Sub fisica_Click()
  optionsform.SSTab1.Tab = 1
  optionsform.Show vbModal
End Sub

Private Sub fittest_Click()
  robfocus = Form1.fittest
End Sub

Private Sub genact_Click()
  ActivForm.Show
End Sub


Private Sub intOptionsOpen_Click()
  optionsform.SSTab1.Tab = 5
  optionsform.Show vbModal
End Sub

Private Sub Leagues_Click()
  optionsform.SSTab1.Tab = 4
  optionsform.Show vbModal
End Sub

Private Sub loadpiccy_Click()
  On Error GoTo fine
  optionsform.Visible = False
  CommonDialog1.DialogTitle = "Load a Background picture file"
  CommonDialog1.InitDir = App.path
  CommonDialog1.FileName = ""
  CommonDialog1.Filter = "Pictures (*.bmp;*.jpg)|*.bmp;*.jpg"
  CommonDialog1.ShowOpen
  If CommonDialog1.FileName <> "" Then Form1.BackPic = CommonDialog1.FileName
  Form1.PiccyMode = True
  Form1.Newpic = True
fine:
End Sub


Private Sub makenewspecies_Click()
  If robfocus > 0 Then
    If rob(robfocus).Corpse Then
      MsgBox ("Sorry, but you cannot make a new species from a corpse.")
    ElseIf MsgBox("Start new species using this robot?", vbYesNo) = vbYes Then
      'Change the species of the bot with the focus
      MakeNewSpeciesFromBot robfocus
    End If
  End If
End Sub

Public Function MakeNewSpeciesFromBot(n As Integer)
Dim i As Integer
Dim OldSpeciesName As String

  If Not robManager.GetExists(n) Or rob(n).Corpse Then Exit Function
  'Change the species of bot n
  OldSpeciesName = rob(n).FName
  If Right(rob(n).FName, 4) = ".txt" Then
    rob(n).FName = Left(rob(n).FName, Len(rob(n).FName) - 4)
  End If
 
  rob(n).FName = Left(rob(n).FName, 28) + Str(Random(1, 10000))
    
  AddSpecie n
  ChangeNameOfAllChildren n, OldSpeciesName
End Function

'Recursivly changes the name of all extant descendants of bot n to be the same as bot n
'Also changes the name of any other bots that have a subspecies number > bot n
'Used when forking a species
Private Sub ChangeNameOfAllChildren(n As Integer, OldSpeciesName As String)
    Dim t As Integer
    If rob(n).SonNumber = 0 Then Exit Sub
    For t = 1 To MaxRobs
        If robManager.GetExists(t) And Not rob(t).Corpse And t <> n And rob(t).parent = rob(n).AbsNum Then
            rob(t).FName = rob(n).FName
            ChangeNameOfAllChildren t, OldSpeciesName
        End If
    Next t
End Sub

Private Sub MonitorOn_Click()
  If frmMonitorSet.overwrite Then MonitorOn.Checked = Not MonitorOn.Checked Else MsgBox "Please configure monitor settings first.", vbInformation
End Sub

Private Sub MonitorSettings_Click()
  frmMonitorSet.Show vbModal
End Sub

Private Sub pause_Click()
  DisplayActivations = False
  Form1.Active = False
  Form1.SecTimer.Enabled = False
  MDIForm1.unpause.Enabled = True
End Sub


Private Sub removepiccy_Click() 'Botsareus 3/24/2012 Added code that deletes the background picture
  Form1.PiccyMode = False
  Form1.Picture = Nothing
End Sub

Private Sub RESOver_Click()
  frmRestriOps.res_state = 3
  frmRestriOps.Show vbModal
End Sub

Private Sub robinf_Click()
  Dim n As Integer
  n = robfocus
  datirob.Show
  datirob.infoupdate n, rob(n).nrg, rob(n).parent, rob(n).Mutations, rob(n).age, rob(n).SonNumber, 1, rob(n).FName, rob(n).genenum, rob(n).LastMut, rob(n).generation, rob(n).DnaLen, rob(n).LastOwner, rob(n).Waste, rob(n).body, rob(n).mass, rob(n).venom, rob(n).shell, rob(n).Slime, rob(n).chloroplasts
End Sub

Private Sub RobTagInfo_Click() 'Botsareus & Peter 9/1/2014 Simple idea to list tag information
Dim all_str() As String
Dim blank As String * 50
ReDim all_str(0)
all_str(0) = "Tag:FileName:User" & vbCrLf & "~~~" & vbCrLf
Dim t As Integer
Dim rob_str As String
Dim i As Integer
Dim datahit As Boolean
For t = 1 To MaxRobs
 If robManager.GetExists(t) Then
  If Left(rob(t).tag, 45) = Left(blank, 45) Then
   rob_str = String(45, " ") & ":" & rob(t).FName & ":" & rob(t).LastOwner
  Else
   rob_str = Left(rob(t).tag, 45) & ":" & rob(t).FName & ":" & rob(t).LastOwner
  End If
  datahit = False
  For i = 0 To UBound(all_str)
   If all_str(i) = rob_str Then
    datahit = True
    Exit For
   End If
  Next
  If Not datahit Then
   ReDim Preserve all_str(UBound(all_str) + 1)
   all_str(UBound(all_str)) = rob_str
  End If
 End If
Next

Clipboard.CLEAR
Clipboard.SetText Join(all_str, vbCrLf)

MsgBox "Data is now copyable from clipboard", vbInformation
End Sub

Private Sub SafemodeBkp_Click()
On Error GoTo b:
shell MainDir & "\SafeModeBackup.exe"
MsgBox "Safemode backup backs up lastautosave.sim every 6 hours.", vbInformation
Exit Sub
b:
MsgBox "Can not find SafeModeBackup.exe. Did you forget to move it into your " & MainDir & " folder?", vbCritical
End Sub

Private Sub SaveSimWithoutMutations_Click()
  SaveWithoutMutations = True
  simsave
End Sub

Private Sub showEyeDesign_Click()
On Error Resume Next
frmEYE.Show
Dim i As Byte
For i = 0 To 8
 frmEYE.txtDir(i).text = rob(robfocus).mem(i + EYE1DIR)
 frmEYE.txtWth(i).text = rob(robfocus).mem(i + EYE1WIDTH)
Next
End Sub

Private Sub ShowVisionGrid_Click()
  ShowVisionGrid.Checked = Not ShowVisionGrid.Checked
  showVisionGridToggle = ShowVisionGrid.Checked
End Sub

Private Sub SnpDeadEnable_Click()
SnpDeadEnable.Checked = Not SnpDeadEnable.Checked
If SnpDeadEnable.Checked Then MsgBox "Snapshot of the dead writes to the DeadRobots.snp and DeadRobots_Mutations.txt in your " & MainDir & "\Autosave folder. " & _
"You can delete this files to reset the snapshot. WARNING: This feature consumes disk space quickly.", vbInformation
SimOpts.DeadRobotSnp = SnpDeadEnable.Checked
TmpOpts.DeadRobotSnp = SnpDeadEnable.Checked
End Sub

Private Sub SnpDeadExRep_Click()
SnpDeadExRep.Checked = Not SnpDeadExRep.Checked
SimOpts.SnpExcludeVegs = SnpDeadExRep.Checked
TmpOpts.SnpExcludeVegs = SnpDeadExRep.Checked
End Sub

Private Sub SnpLiving_Click()
Snapshot
End Sub

Private Sub Species_Click()
  If Not optionsform Is Nothing Then
    optionsform.SSTab1.Tab = 0
    optionsform.Show vbModal
  End If
End Sub


Private Sub SunButton_Click()
  SimOpts.Daytime = Not (SunButton.value * True)
End Sub

Public Sub menuupdate() 'Botsareus 7/13/2012 The menu handler
If limitgraphics Then Toolbar1.buttons(9).value = tbrPressed Else Toolbar1.buttons(9).value = tbrUnpressed
If oneonten Then Toolbar1.buttons(10).value = tbrPressed Else Toolbar1.buttons(10).value = tbrUnpressed
If Not Form1.dispskin Then Toolbar1.buttons(12).value = tbrPressed Else Toolbar1.buttons(12).value = tbrUnpressed
If nopoff Then Toolbar1.buttons(13).value = tbrPressed Else Toolbar1.buttons(13).value = tbrUnpressed
If Not visualize Then Toolbar1.buttons(14).value = tbrPressed Else Toolbar1.buttons(14).value = tbrUnpressed
Toolbar1.Refresh 'Botsareus 1/11/2013 Force toolbar to refresh
End Sub

Sub fixcam() 'Botsareus 2/23/2013 When simulation starts the screen is normailized

'Botsareus 9/12/2014 Simulation now always starts with ignoreerror on
If UseSafeMode = False Then
 ignoreerror = True
 Toolbar1.buttons(24).value = tbrPressed
End If
'Botsareus 2/9/2014 Based on collected data we need to figure out fudging here

showEyeDesign.Enabled = True
inssp.Enabled = True

Form1.BackColor = backgcolor 'Botsareus 4/27/2013 Set back ground skin color
If startnovid Then 'turn off vedio as requested
     visualize = False
     Form1.Label1.Visible = True
     startnovid = False
End If
'
If MDIForm1.WindowState <> 2 Then Exit Sub
If screenratiofix = False Then Exit Sub

'the bloody screen ratio fix took me 4ever - Bots Sometime in June 2014
lockswitch = 0
Form1.visiblew = Form1.Width / Form1.Height * 4 / 3 * Form1.visibleh
ZoomLock.value = 0
ZoomOut
ZoomLock.value = 1
ZoomOut
ZoomIn
End Sub

Private Sub Toolbar1_ButtonClick(ByVal button As MSComctlLib.button)
  Dim a As String
 
  Select Case button.key
    Case "newsim"
      If Form1.GraphLab.Visible Then Exit Sub
      optionsform.SSTab1.Tab = 0
      optionsform.Show vbModal
      If Not optionsform.Canc Then
        Form1.Show
      End If
    Case "loadsim"
      simload
    Case "savesim"
      simsave
    Case "play"
      DisplayActivations = False
      Form1.Active = True
      Form1.SecTimer.Enabled = True
      Form1.unfocus
      Form1.pausefix = False  'Botsareus 3/15/2013 Figure out if simulation must start paused
    Case "stop"
      DisplayActivations = False
      Form1.Active = False
      Form1.SecTimer.Enabled = False
      Form1.pausefix = True 'Botsareus 3/15/2013 Figure out if simulation must start paused
    Case "cycle"
      DisplayActivations = False
      Consoleform.cycle 1
    Case "limit"
      limitgraphics = Not limitgraphics 'Botsareus 7/13/2012 moved icon update to a seporate procedure
      menuupdate
    Case "fast"
      oneonten = Not oneonten 'Botsareus 7/13/2012 added icon update
      menuupdate
    Case "best"
      robfocus = Form1.fittest
    Case "mutfreq"
      optionsform.SSTab1.Tab = 3
      optionsform.Show vbModal
    Case "physics"
      optionsform.SSTab1.Tab = 2
      optionsform.Show vbModal
    Case "costs"
      optionsform.SSTab1.Tab = 2
      optionsform.Show vbModal
    Case "noskin"
      Form1.dispskin = Not Form1.dispskin 'Botsareus 7/13/2012 added icon update
      menuupdate
    Case "nopoff"
      nopoff = Not nopoff 'Botsareus 7/13/2012 added icon update
      menuupdate
    Case "Novideo"
      visualize = Not visualize 'Botsareus 7/13/2012 moved icon update to a seporate procedure
      If visualize Then
        Form1.Label1.Visible = False
      Else
        Form1.Label1.Visible = True
      End If
      menuupdate
    Case "insert"
      If Not insrob Then
        Form1.MousePointer = vbCrosshair
      Else
        Form1.MousePointer = vbArrow
      End If
      insrob = Not insrob
    Case "snapshot"
      Snapshot
    Case "Ignore"
      'ignores errors when it encounters them with the hope that they'll fix themselves
      ignoreerror = Not ignoreerror
      If Not ignoreerror Then
        button.value = tbrUnpressed
      Else
        button.value = tbrPressed
      End If
    
      
  End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu) 'Botsareus 8/3/2012 graph id mod, looks a little better now
'Botsareus 5/26/2013 We now support customizable graphs
  Dim queryhold As String
  Dim queryhelp As String
  queryhelp = vbCrLf & vbCrLf & _
  "Supported variables:" & vbCrLf & _
  "pop= Populations" & vbCrLf & _
  "avgmut= Average Mutations" & vbCrLf & _
  "avgage= Average Age" & vbCrLf & _
  "avgsons= Average Offspring" & vbCrLf & _
  "avgnrg= Average Energy" & vbCrLf & _
  "avglen= Average DNA length" & vbCrLf & _
  "avgcond= Average DNA Cond statements" & vbCrLf & _
  "simnrg= Total Energy_per Species" & vbCrLf & _
  "specidiv= Species Diversity" & vbCrLf & _
  "maxgd= Max Generational Distance" & vbCrLf & _
  "simpgenetic= Selective Genetic Distance" & vbCrLf & vbCrLf & _
  "Supported operators: add sub div mult pow" & vbCrLf & "Please use reverse polish notation."

  Select Case ButtonMenu.key
    Case "pop"
      Form1.NewGraph POPULATION_GRAPH, "Populations"
    Case "avgmut"
      Form1.NewGraph MUTATIONS_GRAPH, "Average_Mutations"
    Case "avgage"
      Form1.NewGraph AVGAGE_GRAPH, "Average_Age"
    Case "avgsons"
      Form1.NewGraph OFFSPRING_GRAPH, "Average_Offspring"
    Case "avgnrg"
      Form1.NewGraph ENERGY_GRAPH, "Average_Energy"
    Case "avglen"
      Form1.NewGraph DNALENGTH_GRAPH, "Average_DNA_length"
    Case "avgcond"
      Form1.NewGraph DNACOND_GRAPH, "Average_DNA_Cond_statements"
    Case "avgmutlen"
      Form1.NewGraph MUT_DNALENGTH_GRAPH, "Average_Mutations_per_DNA_length_x1000-"
    Case "simnrg"
      Form1.NewGraph ENERGY_SPECIES_GRAPH, "Total_Energy_per_Species_x1000-"
    Case "autocost"
      Form1.NewGraph DYNAMICCOSTS_GRAPH, "Dynamic_Costs"
    Case "speciesdiversity"
      Form1.NewGraph SPECIESDIVERSITY_GRAPH, "Species_Diversity"
    Case "avgchlr"
      Form1.NewGraph AVGCHLR_GRAPH, "Average_Chloroplasts"
    Case "maxgeneticdistance"
      Form1.NewGraph GENETIC_DIST_GRAPH, "Genetic_Distance_x1000-"
    Case "maxgenerationaldistance"
      Form1.NewGraph GENERATION_DIST_GRAPH, "Max_Generational_Distance"
    Case "simplegeneticdistance"
      Form1.NewGraph GENETIC_SIMPLE_GRAPH, "Simple_Genetic_Distance_x1000-"
    Case "CG1"
      queryhold = InputBox("Enter query for Customizable Graph 1:" & queryhelp, , strGraphQuery1)
      If queryhold <> "" Then
        strGraphQuery1 = queryhold
        Form1.NewGraph CUSTOM_1_GRAPH, "Customizable_Graph_1-"
      End If
    Case "CG2"
      queryhold = InputBox("Enter query for Customizable Graph 2:" & queryhelp, , strGraphQuery2)
      If queryhold <> "" Then
        strGraphQuery2 = queryhold
        Form1.NewGraph CUSTOM_2_GRAPH, "Customizable_Graph_2-"
      End If
    Case "CG3"
      queryhold = InputBox("Enter query for Customizable Graph 3:" & queryhelp, , strGraphQuery3)
      If queryhold <> "" Then
        strGraphQuery3 = queryhold
        Form1.NewGraph CUSTOM_3_GRAPH, "Customizable_Graph_3-"
      End If
    Case "resgraph"
      If MsgBox("Are you sure you want to reset all graphs?", vbOKCancel) = vbOK Then
        Form1.ResetGraphs (0)
        Form1.FeedGraph (0) ' EricL 4/7/2006 Update the graphs right now instead of waiting until the next update
      End If
    Case "listgraphs"
        Dim lg As String
        lg = "List of all running graphs:" & vbCrLf & Form1.calc_graphs
        MsgBox lg, vbInformation
  End Select
End Sub

Private Sub about_Click()
    frmAbout.Show
End Sub

Private Sub col_Click()
  Form1.changerobcol
End Sub

Private Sub cons_Click()
  Consoleform.openconsole
End Sub

Private Sub costi_Click()
  optionsform.SSTab1.Tab = 2
  optionsform.Show vbModal
End Sub

Private Sub czin_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
  AspettaFlag = True
  ZoomInPremuto
End Sub

Private Sub czin_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
  AspettaFlag = False
End Sub

Public Sub ZoomIn()
  If Form1.visiblew > RobSize * 4 Then
    If robfocus > 0 Then
      xc = robManager.GetPosition(robfocus).x
      yc = robManager.GetPosition(robfocus).y
    Else
      xc = Form1.visiblew / 2 + Form1.ScaleLeft
      yc = Form1.visibleh / 2 + Form1.ScaleTop
    End If
    Form1.visiblew = Form1.visiblew / 1.05
    Form1.visibleh = Form1.visibleh / 1.05
    Form1.ScaleHeight = Form1.visibleh
    Form1.ScaleWidth = Form1.visiblew
    Form1.ScaleTop = yc - Form1.ScaleHeight / 2
    Form1.ScaleLeft = xc - Form1.ScaleWidth / 2
    Form1.Redraw
  End If
  
 If ZoomLock.value = 1 Then
    If lockswitch = 0 Then
        Dim ratio As Double
        ratio = Form1.TwipHeight / Form1.twipWidth
        Dim expectedscreenratio  As Double
        expectedscreenratio = 9645 / 15300
        Dim actualscreenratio As Double
        actualscreenratio = Form1.Height / Form1.Width
        Form1.visiblew = Form1.visiblew * ratio * expectedscreenratio / actualscreenratio * 1.065
        lockswitch = 1
    End If
  End If
End Sub

Public Sub Follow() 'Botsareus 11/29/2013 Zoom follow selected robot
    If robfocus > 0 And Form1.visiblew < 6000 And visualize Then
      xc = robManager.GetPosition(robfocus).x
      yc = robManager.GetPosition(robfocus).y
      Form1.ScaleTop = yc - Form1.ScaleHeight / 2
      Form1.ScaleLeft = xc - Form1.ScaleWidth / 2
  End If
End Sub

Private Sub ZoomInPremuto()
  While AspettaFlag = True
    ZoomIn
    DoEvents
  Wend
End Sub

Private Sub ZoomOutPremuto()
  While AspettaFlag = True
    ZoomOut
    DoEvents
  Wend
End Sub

Private Sub czo_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
  AspettaFlag = True
  ZoomOutPremuto
End Sub

Public Sub ZoomOut()
  Dim tvv As Long
  Dim thv As Long
  
  'EricL Prevents zooming out too far from causing an overflow
  If Form1.visiblew >= 10000000 Then Exit Sub
  If Form1.visibleh >= 10000000 Then Exit Sub

  
  xc = Form1.visiblew / 2 + Form1.ScaleLeft
  yc = Form1.visibleh / 2 + Form1.ScaleTop
  
  Form1.visiblew = Form1.visiblew / 0.95
  Form1.visibleh = Form1.visibleh / 0.95
  
  If Form1.visiblew > SimOpts.FieldWidth And ZoomLock.value = 0 Then
    Form1.visiblew = SimOpts.FieldWidth
    Form1.visibleh = SimOpts.FieldHeight
  End If
  
  If Form1.visibleh > SimOpts.FieldHeight And ZoomLock.value = 0 Then
    Form1.visiblew = SimOpts.FieldWidth
    Form1.visibleh = SimOpts.FieldHeight
  End If
  
  Form1.ScaleTop = yc - Form1.visibleh / 2
  Form1.ScaleLeft = xc - Form1.visiblew / 2
  
  If Form1.visiblew + Form1.ScaleLeft > SimOpts.FieldWidth And ZoomLock.value = 0 Then
    Form1.ScaleLeft = SimOpts.FieldWidth - Form1.visiblew
    Form1.ScaleTop = SimOpts.FieldHeight - Form1.visibleh
  End If
  
  If Form1.ScaleLeft < 0 And ZoomLock.value = 0 Then
    Form1.ScaleLeft = 0
  End If
  
  If Form1.ScaleTop < 0 And ZoomLock.value = 0 Then
    Form1.ScaleTop = 0
  End If
  
  Form1.ScaleHeight = Form1.visibleh
  Form1.ScaleWidth = Form1.visiblew
  
  If ZoomLock.value = 1 Then
    If lockswitch = 0 Then
        Dim ratio As Double
        ratio = Form1.TwipHeight / Form1.twipWidth
        Dim expectedscreenratio  As Double
        expectedscreenratio = 9645 / 15300
        Dim actualscreenratio As Double
       actualscreenratio = Form1.Height / Form1.Width
        Form1.visiblew = Form1.visiblew * ratio * expectedscreenratio / actualscreenratio * 1.065
        lockswitch = 1
    End If
  End If
  Form1.Redraw
End Sub

Private Sub czo_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
 AspettaFlag = False
End Sub

Private Sub inssp_Click()
  On Error GoTo fine
  CommonDialog1.DialogTitle = "Load organism"
  CommonDialog1.FileName = ""
  CommonDialog1.Filter = "Organism file(*.dbo)|*.dbo"
  CommonDialog1.InitDir = MainDir + "\robots"
  CommonDialog1.ShowOpen
  If CommonDialog1.FileName <> "" Then InsertOrganism CommonDialog1.FileName
  Exit Sub
fine:
  MsgBox "Organism not inserted"
End Sub

Private Sub killorg_Click()
  KillOrganism robfocus
End Sub

Private Sub loadsim_Click(Index As Integer)
If Form1.GraphLab.Visible Then Exit Sub

If chseedloadsim Then SimOpts.UserSeedNumber = Timer * 100 'Botsareus 5/3/2013 Change seed on load sim
tmpseed = SimOpts.UserSeedNumber 'Botsareus 5/3/2013 temporarly holds seed for load sim
simload

End Sub

Private Sub simload(Optional path As String)
  Dim i As Integer
  Dim path2 As String
   
  On Error GoTo fine ' Uncomment this line in compiled version error.sim

  If path = "" Then
'    optionsform.Visible = False
    CommonDialog1.DialogTitle = MBLoadSim
    CommonDialog1.InitDir = MDIForm1.MainDir + "\saves"
    CommonDialog1.FileName = ""
    CommonDialog1.Filter = "Simulation(*.sim)|*.sim"
    CommonDialog1.ShowOpen
    If CommonDialog1.CancelError Then
      If Err.Number = 32755 Then
        Exit Sub ' The user pressed the cancel button
      End If
    End If
    If CommonDialog1.FileName = "" Then
      Exit Sub
    Else
      path2 = CommonDialog1.FileName
    End If
    'Botsareus 5/14/2013 Create our local copy
    'Botsareus 6/9/2013 Make saves dir if not found
    RecursiveMkDir MDIForm1.MainDir + "\saves\"
    If path2 <> (MDIForm1.MainDir + "\saves\localcopy.sim") Then FileCopy path2, MDIForm1.MainDir + "\saves\localcopy.sim"
  Else
    path2 = path
    'Botsareus 5/13/2013 Show the safemode lab. and no internet
    If autosaved Then
        Form1.lblSafeMode.Visible = True
        MDIForm1.Objects.Enabled = False
        MDIForm1.inssp.Enabled = False
        MDIForm1.DisableArep.Enabled = False
        MDIForm1.AutoFork.Enabled = False
    End If
  End If
  
  MDIForm1.Caption = MDIForm1.BaseCaption + " " + path2
  
  LoadSimulation path2
    
 'Populate the Add Species dropdown combo when sims loaded
  For i = 0 To SimOpts.SpeciesNum - 1
    If i > MAXNATIVESPECIES Then
      MsgBox "Exceeded number of native species."
    Else
      MDIForm1.Combo1.additem SimOpts.Specie(i).Name
    End If
  Next i
  
  DisplayActivations = False ' EricL - Initialize the flag that controls displaying activations in the console
  
  Form1.startloaded

fine:

  If Err.Number <> 32755 Then
    MsgBox "An Error Occurred.  Darwinbots cannot continue.  Sorry.  " + Err.Description + " " + Err.Source + " " + Str$(Err.Number) + " " + Str$(Err.LastDllError) + ".", vbOKOnly
  Else
    Exit Sub
  End If
End Sub

Private Sub moltiplicatore_Click()
  optionsform.SSTab1.Tab = 3
  optionsform.Show vbModal
End Sub

Private Sub mutrat_Click()
  robmutchange
End Sub

' changes a robot's mutation rates
Public Sub robmutchange()
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Dim t As Integer
  Dim Specie As datispecie
  
  t = optionsform.CurrSpec
  Specie = TmpOpts.Specie(50)
  
  TmpOpts.Specie(50).Mutables = rob(robfocus).Mutables
  optionsform.CurrSpec = 50
  
  MutationsProbability.Show vbModal
  
  rob(robfocus).Mutables = TmpOpts.Specie(50).Mutables
  TmpOpts.Specie(50) = Specie
  optionsform.CurrSpec = t
End Sub

Private Sub par_Click()
  parentele.mostra
End Sub


Private Sub saveorg_Click()
  On Error GoTo fine
  CommonDialog1.DialogTitle = "Save organism"
  CommonDialog1.FileName = ""
  CommonDialog1.Filter = "Organism file(*.dbo)|*.dbo"
  CommonDialog1.InitDir = MainDir + "\robots"
  CommonDialog1.ShowSave
  If CommonDialog1.FileName <> "" Then SaveOrganism CommonDialog1.FileName, robfocus
  Exit Sub
fine:
  MsgBox "Organism not saved"
End Sub

Private Sub savesim_Click()
  SaveWithoutMutations = False
  simsave
End Sub

Private Sub simsave()
  On Error GoTo fine
  CommonDialog1.DialogTitle = MBSaveSim
  CommonDialog1.FileName = ""
  CommonDialog1.Filter = "Simulation(*.sim)|*.sim"
  CommonDialog1.InitDir = MDIForm1.MainDir + "\saves"
  CommonDialog1.ShowSave
  SaveSimulation CommonDialog1.FileName
  Exit Sub
fine:
MsgBox "Saving sim failed.  " + Err.Description, vbOKOnly
End Sub

Private Sub load_league_res()
'Botsareus 7/30/214 Load restrictions
Dim lastmod As Byte
Dim holdother As Byte
Dim i As Byte
'evo restrictions
For i = 0 To UBound(TmpOpts.Specie)
 If TmpOpts.Specie(i).Veg Then
  '
        holdother = 0
  '
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).Fixed = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).CantSee = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).DisableDNA = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).CantReproduce = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).VirusImmune = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).DisableMovementSysvars = lastmod * True
 Else
  '
        holdother = 0
  '
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).NoChlr = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).Fixed = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).CantSee = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).DisableDNA = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).CantReproduce = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).VirusImmune = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).DisableMovementSysvars = lastmod * True
 End If
Next
End Sub
Private Sub load_evo_res()
'Botsareus 7/30/214 Load restrictions
Dim lastmod As Byte
Dim holdother As Byte
Dim i As Byte
'evo restrictions
For i = 0 To UBound(TmpOpts.Specie)
 If TmpOpts.Specie(i).Veg Then
  '
        holdother = 0
  '
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).Fixed = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).CantSee = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).DisableDNA = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).CantReproduce = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).VirusImmune = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).DisableMovementSysvars = lastmod * True
 Else

  '
        holdother = 0
  '
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).NoChlr = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).Fixed = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).CantSee = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).DisableDNA = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).CantReproduce = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).VirusImmune = lastmod * True
        holdother = (holdother - lastmod) / 2
        lastmod = holdother Mod 2
  TmpOpts.Specie(i).DisableMovementSysvars = lastmod * True
 End If
Next
End Sub

Private Sub MDIForm_Load()
'Botsareus 6/16/2014 Starting positions for graphs so they are less annoying
Dim k As Byte
For k = 1 To NUMGRAPHS
   graphleft(k) = Screen.Width - 8400
   graphtop(k) = Screen.Height - 4800
Next k
'Botsareus 7/5/2014 Intialize array for player bot mode
ReDim PB_keys(0)

calc_dnamatrix
'Botsareus 5/8/2013 Safemode strings are declared here (sorry, no Italian version)
Dim strMsgSendData As String
Dim strMsgEnterDiagMode As String

LoadGlobalSettings 'Botsareus 3/15/2013 lets try to load global settings first

'Botsareus 5/8/2013 If program did crash and no autosave then it is time to give all data to the user
strMsgSendData = "Please go to " & MDIForm1.MainDir & " and give the administrator the following files:" & vbCrLf & vbCrLf & _
"Global.gset" & vbCrLf & _
"settings\lastran.set" & vbCrLf & _
"saves\localcopy.sim" & vbCrLf & _
"saves\lastautosave.sim" & vbCrLf & vbCrLf & _
"If you don't see any or all of these file(s) let the administrator know they are missing." & vbCrLf & vbCrLf & _
"If you where running a league please give the following files if they exsist:" & vbCrLf & vbCrLf & _
"league\Test.txt" & vbCrLf & _
"league\robotA.txt" & vbCrLf & _
"league\robotB.txt" & vbCrLf & vbCrLf & _
"If you where running evolution please give the administrator the \evolution\ folder (subfolders not required)." & vbCrLf & vbCrLf

'Botsareus 5/8/2013 If the program did crash and autosave prompt to enter safemode 'Botsareus 4/5/2016 Spelling fix
strMsgEnterDiagMode = "Warning: Diagnostic mode does not check for errors by user generated events. If the error happened immediately after you manipulated the simulation. Please press NO and tell what you did to the administrator. Otherwise, it is recommended that you run diagnostic mode." & vbCrLf & vbCrLf & _
"Do you want to run diagnostic mode?"


If simalreadyrunning And Not autosaved Then
    MsgBox strMsgSendData
    Open App.path & "\Safemode.gset" For Output As #1
     Write #1, False
    Close #1
    Open App.path & "\autosaved.gset" For Output As #1
     Write #1, False
    Close #1
    End
End If

Dim path As String
Dim fso As New FileSystemObject
Dim lastSim As file
Dim revision As String

Form1.Active = True 'Botsareus 2/21/2013 moved active here to enable to pause initial simulation

  globstrings
  strings Me
  MDIForm1.WindowState = 2
  
  MDIForm1.BaseCaption = "DarwinBots " + CStr(App.Major) + "." + CStr(App.Minor) + "." + Format(App.revision, "00")
  MDIForm1.Caption = MDIForm1.BaseCaption
  
  Me.Show
  
  'These are all defaults that might get overridden by the settings loaded below
  F1Internet.Checked = False
  
  ShowVisionGrid.Checked = True
  showVisionGridToggle = True
  displayShotImpactsToggle = True
  displayResourceGuagesToggle = True
  displayMovementVectorsToggle = True
  TmpOpts.allowHorizontalShapeDrift = False
  TmpOpts.allowVerticalShapeDrift = False
  DeleteShape.Enabled = False
  TmpOpts.shapesAreSeeThrough = False
  HighLightTeleportersMenu.Checked = True
  DontDecayNrgShots.Checked = False
  DontDecayWstShots.Checked = False
  DisableTies.Checked = False
  DisableArep.Checked = False
  TmpOpts.DisableTies = False
  TmpOpts.DisableTypArepro = False
  TmpOpts.DisableFixing = False
  TmpOpts.NoShotDecay = False
  TmpOpts.NoWShotDecay = False
  TmpOpts.chartingInterval = 200
  TmpOpts.FieldWidth = 16000
  TmpOpts.FieldHeight = 12000
  TmpOpts.MaxVelocity = 60
  TmpOpts.Costs(DYNAMICCOSTSENSITIVITY) = 50
  TmpOpts.Costs(BOTNOCOSTLEVEL) = -1 'Botsareus 5/11/2012 Sets BotNoCostThreshold to -1 to fix a bug when running a veg only sim.
  TmpOpts.Costs(COSTMULTIPLIER) = 1 'Botsareus 1/5/2013 default for cost multiply
  TmpOpts.VegFeedingToBody = 0.75 'Botsareus 1/5/2013 Vegy feed distribution intialized at energy 25% body 75%
  TmpOpts.Gradient = 1.02 'Botsareus 12/12/2012 Default for Gradient
  TmpOpts.DayNightCycleCounter = 0
  TmpOpts.Daytime = True
  TmpOpts.BadWastelevel = -1
  TmpOpts.FluidSolidCustom = 2 ' Default to custom for older settings files
  TmpOpts.CostRadioSetting = 2 ' Default to custom for older settings files
  TmpOpts.CoefficientElasticity = 0 ' Default for older settings files
  TmpOpts.NoShotDecay = False ' Default for older settings files
  TmpOpts.SunUpThreshold = 500000   'Set to a reasonable default value
  TmpOpts.SunUp = False   'Set to a reasonable default value
  TmpOpts.SunDownThreshold = 1000000   'Set to a reasonable default value
  TmpOpts.SunDown = False   'Set to a reasonable default value
  TmpOpts.FixedBotRadii = False
  TmpOpts.SunThresholdMode = 0
  TmpOpts.PhysMoving = 0.66
  TmpOpts.EnergyExType = True
  TmpOpts.EnergyFix = 200
  TmpOpts.EnergyProp = 1
  TmpOpts.MaxEnergy = 100
  TmpOpts.MaxPopulation = 100
  TmpOpts.MinVegs = 50
  TmpOpts.RepopAmount = 10
  TmpOpts.RepopCooldown = 10
  TmpOpts.PhysBrown = 0.5
  TmpOpts.FieldSize = 2


  
  EnableRobotsMenu
    
  optionsform.ReadSett MDIForm1.MainDir + IIf(simalreadyrunning, "\settings\lastran.set", "\settings\lastexit.set")
  
  'From now on all league and special evolution modes use the restart system.
  'I have decided to get rid of Eric's attempt at the stepladder league primarly because I
  'do not trust the randomizer and DBs current logic incase of a crash. I also wanted the
  'file system to keep track of the league instead of the internal logic of the program for
  'the same reason. Search "R E S" to find the new components. -Bots
  'Botsareus 1/31/2014 R E S T A R T  L O A D
                  Dim files As Collection
                  Dim seeded As Collection
                  Dim i As Byte
                  Dim ecocount As Byte

skipsetup:
    
  'optionsform.datatolist
  'TmpOpts.Daytime = True ' Ericl March 15, 2006
 
 ' Unload optionsform  ' We do this here becuase reading in the settings above loads the Options dialog.
                      ' We want it unloaded so that when the user loads it the next time, it gets properly
                      ' populated by the form's load routine
  SimOpts = TmpOpts
  path = Command
  
  If path = "" Then
 
    On Error GoTo bypass
    Set lastSim = fso.GetFile(MDIForm1.MainDir + IIf(autosaved, "\Saves\lastautosave.sim", "\Saves\lastexit.sim"))
    If lastSim.size > 0 Then
      'Botsareus 5/8/2013 Change of message for diag mode
      If MsgBox(IIf(autosaved, strMsgEnterDiagMode, "Continue the last simulation?"), vbYesNo + vbExclamation, MBwarning) = vbYes Then
         simload MDIForm1.MainDir + IIf(autosaved, "\Saves\lastautosave.sim", "\Saves\lastexit.sim")
      End If
    End If
     
  Else
    If InStr(Command, """") <> 0 Then path = Replace(Command, """", "")
    If InStr(Command, "\") <> 0 Then
      simload path
    Else
      simload MDIForm1.MainDir + "\saves\" + path
    End If
  End If
bypass:

'Botsareus 5/3/2013 Randomize seed here
SimOpts.UserSeedNumber = Timer * 100
End Sub

Public Function EnableRobotsMenu()

  MDIForm1.robinf.Enabled = True
  MDIForm1.par.Enabled = True
  MDIForm1.mutrat.Enabled = True
  MDIForm1.col.Enabled = True
  MDIForm1.cons.Enabled = True
  MDIForm1.genact.Enabled = True
  MDIForm1.ucci.Enabled = True
  MDIForm1.sdna.Enabled = True
  MDIForm1.selorg.Enabled = True
  MDIForm1.saveorg.Enabled = True
  MDIForm1.killorg.Enabled = True
End Function

Public Function DisableRobotsMenu()

  MDIForm1.robinf.Enabled = False
  MDIForm1.par.Enabled = False
  MDIForm1.mutrat.Enabled = False
  MDIForm1.col.Enabled = False
  MDIForm1.cons.Enabled = False
  MDIForm1.genact.Enabled = False
  MDIForm1.ucci.Enabled = False
  MDIForm1.sdna.Enabled = False
  MDIForm1.selorg.Enabled = False
  MDIForm1.saveorg.Enabled = False
  MDIForm1.killorg.Enabled = False
       
End Function

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Caption Like "Moving files*" Then
MsgBox "Please wait until files are calculated"
Cancel = True
Exit Sub
End If

If Form1.lblSaving.Visible Then 'Botsareus 2/7/2014 small bug fix for autosave
    Cancel = 1
    Exit Sub
End If

Form1.hide_graphs

'Botsareus 5/5/2013 Replaced MBsure with a better message. (Sorry, no Italian version)
'Botsareus 5/10/2013 Only prompt to overwrite setting if lastexit.set already exisits.
If dir(MDIForm1.MainDir + "\settings\lastexit.set") <> "" Then

  Select Case MsgBox("Would you like to save changes to the settings? Press CANCEL to return to the program.", vbYesNoCancel + vbExclamation, MBwarning)
  Case vbYes
    Form1.Form_Unload 0
    datirob.Form_Unload 1
          
        'moved savesett here
        If optionsform.Visible = False Then
          TmpOpts = SimOpts
        End If
        optionsform.savesett MDIForm1.MainDir + "\settings\lastexit.set" 'save last settings
        
    Form1.Form_Unload 0
    datirob.Form_Unload 1
    MDIForm_Unload 0
  Case vbNo
    Form1.Form_Unload 0
    datirob.Form_Unload 1
    MDIForm_Unload 0
  Case vbCancel
    Cancel = 1
  End Select
  
Else

  If MsgBox(MBsure, vbYesNo + vbExclamation, MBwarning) = vbYes Then
  
        'copyed savesett here
        If optionsform.Visible = False Then
          TmpOpts = SimOpts
        End If
        optionsform.savesett MDIForm1.MainDir + "\settings\lastexit.set" 'save last settings
  
    Form1.Form_Unload 0
    datirob.Form_Unload 1
    MDIForm_Unload 0
  Else
    Cancel = 1
  End If

End If

Form1.show_graphs
End Sub


Private Sub MDIForm_Resize()
'  Form1.dimensioni
  'InfoForm.ZOrder
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
   
SaveSimulation MDIForm1.MainDir + "\saves\lastexit.sim"  'save last settings
  
  'Botsareus 5/5/2013 Update the system that the program closed
  
    Open App.path & "\Safemode.gset" For Output As #1
      Write #1, False
    Close #1
  
  End
End Sub

Sub infos(ByVal cyc As Single, tot As Integer, tnv As Integer, tv As Long, brn As Long, totcyc As Long, tottim As Long)  'Botsareus 8/25/2013 Mod to except totalchlr
  Dim sec As Long
  Dim Min As Long
  Dim h As Long
  Dim i As Integer
  Dim k As Integer
 ' Dim AvgSimEnergyLastHundredCycles As Long
  Dim AvgSimEnergyLastTenCycles As Long
  Dim Delta As Double
  
  If tot = 0 Then Exit Sub
  StatusBar1.Panels(1).text = SBcycsec + Str$(Round(cyc, 3)) + " "
  StatusBar1.Panels(2).text = "Tot " + Str$(tot) + " "
  Me.TotalBots.Caption = "Tot " + Str$(tot)
  StatusBar1.Panels(3).text = "Bots " + Str$(tnv) + " "
  Me.NumberBots.Caption = "Bots " + Str$(tnv)
  StatusBar1.Panels(4).text = "Chlr " + Str$(tv) + " "
  Me.NumberVeg.Caption = "Chlr " + Str$(tv)
  StatusBar1.Panels(5).text = SBborn + Str$(brn) + " "
  Me.BotsBorn.Caption = SBborn + Str$(brn)
  StatusBar1.Panels(6).text = "Cycles" + Str$(totcyc) + " "
  Me.CyclesNumber.Caption = "Cycles" + Str$(totcyc)
  sec = tottim
  Min = Fix(sec / 60)
  sec = sec Mod 60
  h = Fix(Min / 60)
  Min = Min Mod 60
  StatusBar1.Panels(7).text = Str$(h) + "h" + Str$(Min) + "m" + Str$(sec) + "s  "
  StatusBar1.Panels(8).text = "Mut " + Str$(SimOpts.MutCurrMult) + "x "
  Me.MutationValue.Caption = "Mut " + Str$(SimOpts.MutCurrMult)
  StatusBar1.Panels(10).text = "Shots " + Str$(Shots_Module.ShotsThisCycle) + " "
  
  'AvgSimEnergyLastTenCycles = 0
  'This delibertly counts the 10 cycles *before* this one to avoid cases where the timer invokes
  'this routine before the calculations for the current energy cycle have completed.
  'For i = 99 To 90 Step -1
  '  k = (CurrentEnergyCycle + i) Mod 100
  '  AvgSimEnergyLastTenCycles = AvgSimEnergyLastTenCycles + (TotalSimEnergy(k) * 0.1)
  'Next i
  
' AvgSimEnergyLastHundredCycles = AvgSimEnergyLastTenCycles
'    k = (CurrentEnergyCycle + 100 - i) Mod 100
'    AvgSimEnergyLastHundredCycles = AvgSimEnergyLastHundredCycles + TotalSimEnergy(k)
 ' Next i
 ' AvgSimEnergyLastTenCycles = AvgSimEnergyLastTenCycles * 0.1
 ' AvgSimEnergyLastHundredCycles = AvgSimEnergyLastHundredCycles * 0.01

  'If AvgSimEnergyLastTenCycles <> 0 Then delta = TotalSimEnergyDisplayed - AvgSimEnergyLastTenCycles
  k = (CurrentEnergyCycle + 98) Mod 100
  Delta = TotalSimEnergyDisplayed - TotalSimEnergy(k)
  
  StatusBar1.Panels(11).text = "Nrg " + Str$(TotalSimEnergyDisplayed) + " "
  StatusBar1.Panels(12).text = "Delta " + Str$(Round(Delta, 5)) + " "
  StatusBar1.Panels(13).text = "CostX " + Str$(Round(SimOpts.Costs(COSTMULTIPLIER), 5)) + " "
End Sub

Private Sub newsim_Click(Index As Integer)
If Form1.GraphLab.Visible Then Exit Sub
optionsform.SSTab1.Tab = 0
optionsform.Show vbModal
If Not optionsform.Canc Then
  Form1.Show
End If
End Sub

Private Sub quit_Click() 'Botsareus 2/7/2014 Simple quit code
  MDIForm_QueryUnload 0, 0
End Sub

Private Sub sdna_Click()
  robsave
End Sub

Private Sub robsave()
  On Error GoTo fine
  CommonDialog1.DialogTitle = MBSaveDNA
  CommonDialog1.FileName = ""
  CommonDialog1.Filter = "DNA file(*.txt)|*.txt"
  CommonDialog1.InitDir = MainDir + "\robots"
  CommonDialog1.ShowSave
  If CommonDialog1.FileName <> "" Then salvarob robfocus, CommonDialog1.FileName
  Exit Sub
fine:
  MsgBox MBDNANotSaved
End Sub

Private Sub selorg_Click()
  FreezeOrganism robfocus
End Sub


Private Sub ucci_Click()
  KillRobot robfocus 'Botsareus 6/12/2016 Bugfix
End Sub

Private Sub unpause_Timer()
If GetAsyncKeyState(vbKeyF12) Then
      DisplayActivations = False
      Form1.Active = True
      Form1.SecTimer.Enabled = True
      Form1.unfocus
      unpause.Enabled = False
End If
End Sub

Public Sub enablesim()
  Edit.Enabled = True
  popup.Enabled = True
  czin.Enabled = True
  czo.Enabled = True
End Sub

Private Sub y_info_Click()
MsgBox "Target DNA size: N/A" & vbCrLf & "Next DNA size change: N/A" & vbCrLf, vbInformation, "Survival information"
End Sub

Private Sub ZoomLock_Click()
  If Not MDIForm1.ZoomLock Then
     Form1.visiblew = Screen.Width / Screen.Height * 4 / 3 * Form1.visibleh
  Else
     Form1.visiblew = 0.75 * Form1.visibleh
  End If
End Sub
