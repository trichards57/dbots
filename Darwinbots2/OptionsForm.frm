VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form optionsform 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simulation Settings"
   ClientHeight    =   6768
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   10560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6768
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   645
      Left            =   240
      TabIndex        =   0
      Top             =   5760
      Width           =   9660
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5760
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton gset 
         Caption         =   "Global settings"
         Height          =   330
         Left            =   4440
         TabIndex        =   161
         Top             =   225
         Width           =   1305
      End
      Begin VB.CommandButton OKButton 
         Caption         =   "Change"
         Height          =   330
         Left            =   8520
         TabIndex        =   4
         Top             =   225
         Width           =   975
      End
      Begin VB.CommandButton StartNew 
         Caption         =   "Start new"
         Height          =   330
         Left            =   7440
         TabIndex        =   3
         Top             =   225
         Width           =   960
      End
      Begin VB.CommandButton SaveSettings 
         Caption         =   "Save settings"
         Height          =   330
         Left            =   1560
         TabIndex        =   2
         Top             =   225
         Width           =   1305
      End
      Begin VB.CommandButton LoadSettings 
         Caption         =   "Load settings"
         Height          =   330
         Left            =   210
         TabIndex        =   1
         Top             =   225
         Width           =   1305
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6690
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   11811
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   10
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   12632256
      TabCaption(0)   =   "Species"
      TabPicture(0)   =   "OptionsForm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SpeciesLabel"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label36"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CommentBox"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "RenameButton"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DelSpec"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "AddSpec"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "SpecList"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "DuplicaButt"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "NativeSpeciesButton"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "General"
      TabPicture(1)   =   "OptionsForm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GenPropFrame"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Physics and Costs"
      TabPicture(2)   =   "OptionsForm.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame21"
      Tab(2).Control(1)=   "Frame20"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Mutations"
      TabPicture(3)   =   "OptionsForm.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame13"
      Tab(3).Control(1)=   "DisableMutationsCheck"
      Tab(3).ControlCount=   2
      Begin VB.CommandButton NativeSpeciesButton 
         Caption         =   "List Non-Native Species "
         Height          =   375
         Left            =   1440
         TabIndex        =   159
         Tag             =   "0"
         ToolTipText     =   "Add a new robot type to the simulation"
         Top             =   3240
         Width           =   2295
      End
      Begin VB.CheckBox DisableMutationsCheck 
         Caption         =   "Disable All Mutations"
         Height          =   375
         Left            =   -74640
         TabIndex        =   158
         Top             =   840
         Width           =   1815
      End
      Begin VB.Frame Frame20 
         Caption         =   "Physics"
         Height          =   5235
         Left            =   -74820
         TabIndex        =   123
         Top             =   420
         Width           =   4635
         Begin VB.CommandButton ToPhysics 
            Caption         =   "Custom Physics"
            Height          =   375
            Left            =   1800
            TabIndex        =   131
            Top             =   1680
            Width           =   1950
         End
         Begin VB.ComboBox EfficiencyCombo 
            Height          =   315
            ItemData        =   "OptionsForm.frx":0070
            Left            =   2880
            List            =   "OptionsForm.frx":007D
            TabIndex        =   130
            Text            =   "Set Efficiency"
            Top             =   2400
            Width           =   1395
         End
         Begin VB.Frame Frame26 
            Caption         =   "The Big Blue Screen Acts Like A "
            Height          =   1935
            Left            =   120
            TabIndex        =   127
            Top             =   300
            Width           =   4395
            Begin VB.OptionButton FluidSolidRadio 
               Caption         =   "Custom"
               Height          =   195
               Index           =   2
               Left            =   360
               TabIndex        =   141
               Top             =   1440
               Width           =   3015
            End
            Begin VB.OptionButton FluidSolidRadio 
               Caption         =   "Solid"
               Height          =   195
               Index           =   1
               Left            =   360
               TabIndex        =   140
               Top             =   960
               Width           =   975
            End
            Begin VB.OptionButton FluidSolidRadio 
               Caption         =   "Fluid"
               Height          =   195
               Index           =   0
               Left            =   360
               TabIndex        =   139
               Top             =   480
               Width           =   975
            End
            Begin VB.ComboBox DragCombo 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "OptionsForm.frx":00A0
               Left            =   1680
               List            =   "OptionsForm.frx":00B0
               TabIndex        =   129
               Text            =   "Set Drag"
               Top             =   360
               Width           =   1995
            End
            Begin VB.ComboBox FrictionCombo 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "OptionsForm.frx":00EE
               Left            =   1680
               List            =   "OptionsForm.frx":00FE
               TabIndex        =   128
               Text            =   "Set Friction"
               Top             =   840
               Width           =   1995
            End
         End
         Begin VB.ComboBox BrownianCombo 
            Height          =   315
            ItemData        =   "OptionsForm.frx":0131
            Left            =   2880
            List            =   "OptionsForm.frx":013E
            TabIndex        =   126
            Text            =   "Set Brownian"
            Top             =   2760
            Width           =   1395
         End
         Begin VB.ComboBox GravityCombo 
            Height          =   315
            ItemData        =   "OptionsForm.frx":015F
            Left            =   2880
            List            =   "OptionsForm.frx":0172
            TabIndex        =   125
            Text            =   "Set Gravity"
            Top             =   3120
            Width           =   1395
         End
         Begin MSComctlLib.Slider MaxVelSlider 
            Height          =   495
            Left            =   240
            TabIndex        =   124
            ToolTipText     =   "Maximum bot velocity"
            Top             =   3840
            Width           =   4215
            _ExtentX        =   7430
            _ExtentY        =   868
            _Version        =   393216
            LargeChange     =   36
            SmallChange     =   5
            Max             =   180
            TickStyle       =   1
            TickFrequency   =   10
         End
         Begin MSComctlLib.Slider Elasticity 
            Height          =   495
            Left            =   240
            TabIndex        =   145
            ToolTipText     =   "Controls how elastic bots act during collisions."
            Top             =   4680
            Width           =   4215
            _ExtentX        =   7430
            _ExtentY        =   868
            _Version        =   393216
            LargeChange     =   1
            Min             =   -10
            TickStyle       =   1
         End
         Begin VB.Label Label47 
            Caption         =   "Brownian Movment and Vertical Gravity currently overwritten by Tidal Mode"
            Height          =   615
            Left            =   480
            TabIndex        =   166
            Top             =   2760
            Width           =   2175
         End
         Begin VB.Label Label38 
            Caption         =   "Fast"
            Height          =   255
            Left            =   4080
            TabIndex        =   150
            Top             =   3600
            Width           =   375
         End
         Begin VB.Label Label37 
            Caption         =   "Slow"
            Height          =   255
            Left            =   240
            TabIndex        =   149
            Top             =   3600
            Width           =   615
         End
         Begin VB.Label Label33 
            Caption         =   "Marbles"
            Height          =   255
            Left            =   3840
            TabIndex        =   148
            Top             =   4440
            Width           =   615
         End
         Begin VB.Label Label18 
            Caption         =   "Ghosts"
            Height          =   255
            Left            =   240
            TabIndex        =   147
            Top             =   4440
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "Collision Elasticity"
            Height          =   255
            Left            =   1440
            TabIndex        =   146
            Top             =   4440
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Movement Efficiency"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   135
            Top             =   2400
            Width           =   1875
         End
         Begin VB.Label Label7 
            Caption         =   "Brownian Movement"
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   134
            Top             =   2760
            Width           =   1875
         End
         Begin VB.Label Label7 
            Caption         =   "Vertical Gravity"
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   133
            Top             =   3120
            Width           =   1875
         End
         Begin VB.Label Label12 
            Caption         =   "Max Velocity"
            Height          =   255
            Left            =   1920
            TabIndex        =   132
            Top             =   3600
            Width           =   975
         End
      End
      Begin VB.Frame GenPropFrame 
         Caption         =   "General Properties"
         Height          =   5295
         Left            =   -74880
         TabIndex        =   58
         Tag             =   "2020"
         Top             =   360
         Width           =   9885
         Begin VB.Frame Frame22 
            Caption         =   "Wrap Around"
            Height          =   795
            Left            =   120
            TabIndex        =   119
            Top             =   2160
            Width           =   2055
            Begin VB.CheckBox TopDownCheck 
               Alignment       =   1  'Right Justify
               Caption         =   "Top / Down Wrap"
               Height          =   255
               Left            =   120
               TabIndex        =   121
               Top             =   240
               Width           =   1605
            End
            Begin VB.CheckBox RightLeftCheck 
               Alignment       =   1  'Right Justify
               Caption         =   "Left / Right Wrap"
               Height          =   255
               Left            =   120
               TabIndex        =   120
               Top             =   480
               Width           =   1605
            End
         End
         Begin VB.Frame Frame23 
            Caption         =   "Misc. Controls"
            Height          =   975
            Left            =   120
            TabIndex        =   118
            Top             =   4200
            Width           =   2055
            Begin VB.CheckBox FixBotRadius 
               Alignment       =   1  'Right Justify
               Caption         =   "Fix bot radii"
               Height          =   255
               Left            =   120
               TabIndex        =   151
               ToolTipText     =   "Shots will live forever unitl they impact a bot."
               Top             =   360
               Width           =   1695
            End
         End
         Begin VB.Frame Frame24 
            Caption         =   "Field Controls"
            Height          =   795
            Left            =   120
            TabIndex        =   111
            Top             =   240
            Width           =   4035
            Begin MSComctlLib.Slider FieldSizeSlide 
               Height          =   210
               Left            =   60
               TabIndex        =   112
               Top             =   480
               Width           =   2415
               _ExtentX        =   4255
               _ExtentY        =   381
               _Version        =   393216
               LargeChange     =   1
               Min             =   1
               Max             =   25
               SelStart        =   2
               Value           =   2
            End
            Begin VB.Label Label28 
               Caption         =   "Width:"
               Height          =   240
               Index           =   0
               Left            =   2640
               TabIndex        =   117
               Tag             =   "0"
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label32 
               Caption         =   "Height:"
               Height          =   240
               Left            =   2640
               TabIndex        =   116
               Tag             =   "0"
               Top             =   480
               Width           =   525
            End
            Begin VB.Label FWidthLab 
               Caption         =   "XXXX"
               Height          =   195
               Left            =   3270
               TabIndex        =   115
               Top             =   240
               Width           =   690
            End
            Begin VB.Label FHeightLab 
               Caption         =   "XXXX"
               Height          =   240
               Left            =   3270
               TabIndex        =   114
               Top             =   480
               Width           =   690
            End
            Begin VB.Label FieldSizeLab 
               Caption         =   "Size"
               Height          =   255
               Left            =   1080
               TabIndex        =   113
               Tag             =   "0"
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.Frame Frame25 
            Caption         =   "Waste"
            Height          =   615
            Left            =   120
            TabIndex        =   107
            Top             =   3120
            Width           =   2055
            Begin VB.TextBox CustomWaste 
               Height          =   285
               Left            =   1020
               TabIndex        =   109
               Text            =   "400"
               Top             =   240
               Width           =   555
            End
            Begin ComCtl2.UpDown WasteThresholdUpDown 
               Height          =   285
               Left            =   1680
               TabIndex        =   108
               Top             =   240
               Width           =   255
               _ExtentX        =   445
               _ExtentY        =   508
               _Version        =   327681
               BuddyControl    =   "CustomWaste"
               BuddyDispid     =   196650
               OrigLeft        =   780
               OrigTop         =   480
               OrigRight       =   1020
               OrigBottom      =   795
               Increment       =   50
               Max             =   30000
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label Label21 
               Caption         =   "Threshold"
               Height          =   195
               Left            =   180
               TabIndex        =   110
               Top             =   300
               Width           =   735
            End
         End
         Begin VB.Frame Frame27 
            Caption         =   "Random Numbers"
            Height          =   975
            Left            =   120
            TabIndex        =   104
            Top             =   1200
            Width           =   2055
            Begin VB.CommandButton Newseed 
               Caption         =   "Change seed"
               Height          =   255
               Left            =   120
               TabIndex        =   162
               Top             =   240
               Width           =   1815
            End
            Begin VB.TextBox UserSeedText 
               Height          =   285
               Left            =   1140
               TabIndex        =   105
               Text            =   "1234"
               Top             =   540
               Width           =   795
            End
            Begin VB.Label Label58 
               Caption         =   "Seed Value"
               Height          =   195
               Left            =   165
               TabIndex        =   106
               Top             =   600
               Width           =   915
            End
         End
         Begin VB.Frame VeggyControls 
            Caption         =   "Veggy Controls"
            Height          =   4995
            Left            =   4200
            TabIndex        =   72
            Tag             =   "2200"
            Top             =   180
            Width           =   5610
            Begin VB.CommandButton Energy 
               Caption         =   "Energy Management"
               Height          =   570
               Left            =   3960
               TabIndex        =   165
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Frame Frame16 
               Caption         =   "Pond Mode"
               Height          =   1515
               Left            =   240
               TabIndex        =   93
               Top             =   240
               Width           =   3375
               Begin VB.Frame Frame17 
                  Caption         =   "Light Demo"
                  Height          =   1275
                  Left            =   2100
                  TabIndex        =   97
                  Top             =   180
                  Width           =   1215
                  Begin VB.Timer tmrLight 
                     Enabled         =   0   'False
                     Interval        =   100
                     Left            =   840
                     Top             =   360
                  End
                  Begin VB.TextBox EnergyScalingFactor 
                     Height          =   285
                     Left            =   540
                     TabIndex        =   98
                     Text            =   "42"
                     ToolTipText     =   "Scale the brightness of the graph to the left.  Value sets the energy gain per cycle above which the brightness is set to maximum."
                     Top             =   900
                     Visible         =   0   'False
                     Width           =   375
                  End
                  Begin VB.Label Label23 
                     Caption         =   "Energy Scaling Factor"
                     Height          =   615
                     Left            =   540
                     TabIndex        =   99
                     Top             =   240
                     Width           =   555
                  End
                  Begin VB.Line LightStrata 
                     BorderStyle     =   6  'Inside Solid
                     Index           =   15
                     X1              =   90
                     X2              =   510
                     Y1              =   250
                     Y2              =   250
                  End
                  Begin VB.Line LightStrata 
                     BorderStyle     =   6  'Inside Solid
                     BorderWidth     =   4
                     Index           =   14
                     X1              =   120
                     X2              =   480
                     Y1              =   1140
                     Y2              =   1140
                  End
                  Begin VB.Line LightStrata 
                     BorderStyle     =   6  'Inside Solid
                     BorderWidth     =   4
                     Index           =   13
                     X1              =   120
                     X2              =   480
                     Y1              =   1080
                     Y2              =   1080
                  End
                  Begin VB.Line LightStrata 
                     BorderStyle     =   6  'Inside Solid
                     BorderWidth     =   4
                     Index           =   12
                     X1              =   120
                     X2              =   480
                     Y1              =   1020
                     Y2              =   1020
                  End
                  Begin VB.Line LightStrata 
                     BorderStyle     =   6  'Inside Solid
                     BorderWidth     =   4
                     Index           =   11
                     X1              =   120
                     X2              =   480
                     Y1              =   960
                     Y2              =   960
                  End
                  Begin VB.Line LightStrata 
                     BorderStyle     =   6  'Inside Solid
                     BorderWidth     =   4
                     Index           =   10
                     X1              =   120
                     X2              =   480
                     Y1              =   900
                     Y2              =   900
                  End
                  Begin VB.Line LightStrata 
                     BorderStyle     =   6  'Inside Solid
                     BorderWidth     =   4
                     Index           =   9
                     X1              =   120
                     X2              =   480
                     Y1              =   840
                     Y2              =   840
                  End
                  Begin VB.Line LightStrata 
                     BorderStyle     =   6  'Inside Solid
                     BorderWidth     =   4
                     Index           =   8
                     X1              =   120
                     X2              =   480
                     Y1              =   780
                     Y2              =   780
                  End
                  Begin VB.Line LightStrata 
                     BorderStyle     =   6  'Inside Solid
                     BorderWidth     =   4
                     Index           =   7
                     X1              =   120
                     X2              =   480
                     Y1              =   720
                     Y2              =   720
                  End
                  Begin VB.Line LightStrata 
                     BorderStyle     =   6  'Inside Solid
                     BorderWidth     =   4
                     Index           =   6
                     X1              =   120
                     X2              =   480
                     Y1              =   660
                     Y2              =   660
                  End
                  Begin VB.Line LightStrata 
                     BorderStyle     =   6  'Inside Solid
                     BorderWidth     =   4
                     Index           =   5
                     X1              =   120
                     X2              =   480
                     Y1              =   600
                     Y2              =   600
                  End
                  Begin VB.Line LightStrata 
                     BorderStyle     =   6  'Inside Solid
                     BorderWidth     =   4
                     Index           =   4
                     X1              =   120
                     X2              =   480
                     Y1              =   540
                     Y2              =   540
                  End
                  Begin VB.Line LightStrata 
                     BorderStyle     =   6  'Inside Solid
                     BorderWidth     =   4
                     Index           =   3
                     X1              =   120
                     X2              =   480
                     Y1              =   480
                     Y2              =   480
                  End
                  Begin VB.Line LightStrata 
                     BorderStyle     =   6  'Inside Solid
                     BorderWidth     =   4
                     Index           =   2
                     X1              =   120
                     X2              =   480
                     Y1              =   420
                     Y2              =   420
                  End
                  Begin VB.Line LightStrata 
                     BorderStyle     =   6  'Inside Solid
                     BorderWidth     =   4
                     Index           =   1
                     X1              =   120
                     X2              =   480
                     Y1              =   360
                     Y2              =   360
                  End
                  Begin VB.Line LightStrata 
                     BorderStyle     =   6  'Inside Solid
                     BorderWidth     =   4
                     Index           =   0
                     X1              =   120
                     X2              =   480
                     Y1              =   300
                     Y2              =   300
                  End
                  Begin VB.Shape Shape4 
                     BackColor       =   &H00FFFF00&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H00FFFF00&
                     Height          =   915
                     Left            =   90
                     Top             =   255
                     Width           =   425
                  End
               End
               Begin VB.TextBox Gradient 
                  Height          =   285
                  Left            =   1350
                  TabIndex        =   96
                  Text            =   "0"
                  ToolTipText     =   "Set the gradient for light transmission through the water. A value of zero means no light reduction at any depth."
                  Top             =   780
                  Width           =   420
               End
               Begin VB.TextBox LightText 
                  Height          =   285
                  Left            =   1350
                  TabIndex        =   95
                  Text            =   "100"
                  ToolTipText     =   "Set the light intensity to feed your veggies"
                  Top             =   480
                  Width           =   420
               End
               Begin VB.CheckBox Pondcheck 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Pondmode"
                  Height          =   255
                  Left            =   60
                  TabIndex        =   94
                  ToolTipText     =   $"OptionsForm.frx":0198
                  Top             =   240
                  Width           =   1980
               End
               Begin ComCtl2.UpDown GradientUpDn 
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   100
                  ToolTipText     =   "Set the gradient for light transmission through the water. A value of zero means no light reduction at any depth."
                  Top             =   780
                  Width           =   255
                  _ExtentX        =   445
                  _ExtentY        =   508
                  _Version        =   327681
                  Value           =   1
                  BuddyControl    =   "Gradient"
                  BuddyDispid     =   196665
                  OrigLeft        =   1800
                  OrigTop         =   1140
                  OrigRight       =   2055
                  OrigBottom      =   1425
                  Increment       =   2
                  Max             =   1000
                  SyncBuddy       =   -1  'True
                  BuddyProperty   =   65569
                  Enabled         =   -1  'True
               End
               Begin ComCtl2.UpDown LightUpDn 
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   101
                  ToolTipText     =   "Set the light intensity to feed your veggies"
                  Top             =   480
                  Width           =   255
                  _ExtentX        =   445
                  _ExtentY        =   508
                  _Version        =   327681
                  BuddyControl    =   "LightText"
                  BuddyDispid     =   196666
                  OrigLeft        =   5040
                  OrigTop         =   1560
                  OrigRight       =   5280
                  OrigBottom      =   1845
                  Increment       =   2
                  Max             =   1000
                  SyncBuddy       =   -1  'True
                  BuddyProperty   =   65547
                  Enabled         =   -1  'True
               End
               Begin VB.Label LightLabel 
                  Caption         =   "Light intensity"
                  Height          =   240
                  Left            =   90
                  TabIndex        =   103
                  ToolTipText     =   "Set the light intensity to feed your veggies"
                  Top             =   540
                  Width           =   975
               End
               Begin VB.Label GradientLabel 
                  Caption         =   "Sediment Level"
                  Height          =   240
                  Left            =   90
                  TabIndex        =   102
                  ToolTipText     =   "Set the gradient for light transmission through the water. A value of zero means no light reduction at any depth."
                  Top             =   840
                  Width           =   1155
               End
            End
            Begin VB.Frame Frame31 
               Caption         =   "Veg Body/NRG Distribution"
               Height          =   915
               Left            =   1920
               TabIndex        =   89
               Top             =   4020
               Width           =   3630
               Begin MSComctlLib.Slider BodyNrgDist 
                  Height          =   570
                  Left            =   480
                  TabIndex        =   90
                  TabStop         =   0   'False
                  ToolTipText     =   "When a veg gets NRG, what percentage should go into body points?"
                  Top             =   240
                  Width           =   2355
                  _ExtentX        =   4149
                  _ExtentY        =   1016
                  _Version        =   393216
                  LargeChange     =   10
                  Max             =   100
                  TickStyle       =   2
                  TickFrequency   =   10
               End
               Begin VB.Label LabelZZ 
                  Caption         =   "Body"
                  Height          =   195
                  Index           =   18
                  Left            =   2880
                  TabIndex        =   92
                  Top             =   420
                  Width           =   375
               End
               Begin VB.Label LabelZZ 
                  Caption         =   "NRG"
                  Height          =   195
                  Index           =   17
                  Left            =   120
                  TabIndex        =   91
                  Top             =   420
                  Width           =   375
               End
            End
            Begin VB.Frame Frame29 
               Caption         =   "Rates for Repopulation"
               Height          =   2175
               Left            =   2040
               TabIndex        =   76
               Top             =   1800
               Width           =   3375
               Begin VB.CheckBox KillDistVegsCheck 
                  Caption         =   "Force robots to sustain their own chloroplasts"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   160
                  ToolTipText     =   $"OptionsForm.frx":022C
                  Top             =   1720
                  Visible         =   0   'False
                  Width           =   3135
               End
               Begin VB.TextBox RepopCooldownText 
                  Height          =   285
                  Left            =   120
                  TabIndex        =   80
                  Text            =   "1"
                  ToolTipText     =   "The minimum number of cycles that must elapse between repopulation events."
                  Top             =   1380
                  Width           =   480
               End
               Begin VB.TextBox RepopAmountText 
                  Height          =   285
                  Left            =   120
                  TabIndex        =   79
                  Text            =   "10"
                  ToolTipText     =   "How many robots to drop in the sim per repopulation events."
                  Top             =   1040
                  Width           =   480
               End
               Begin VB.TextBox MinVegText 
                  Height          =   285
                  Left            =   120
                  TabIndex        =   78
                  Text            =   "10"
                  ToolTipText     =   $"OptionsForm.frx":02F6
                  Top             =   660
                  Width           =   495
               End
               Begin VB.TextBox MaxPopText 
                  Height          =   285
                  Left            =   120
                  TabIndex        =   77
                  Text            =   "1000"
                  Top             =   280
                  Width           =   480
               End
               Begin ComCtl2.UpDown RepopAmountUpDown 
                  Height          =   285
                  Index           =   0
                  Left            =   600
                  TabIndex        =   81
                  Top             =   1040
                  Width           =   255
                  _ExtentX        =   445
                  _ExtentY        =   487
                  _Version        =   327681
                  BuddyControl    =   "RepopAmountText"
                  BuddyDispid     =   196675
                  OrigLeft        =   660
                  OrigTop         =   1020
                  OrigRight       =   900
                  OrigBottom      =   1305
                  SyncBuddy       =   -1  'True
                  BuddyProperty   =   65547
                  Enabled         =   -1  'True
               End
               Begin ComCtl2.UpDown MaxPopUpDn 
                  Height          =   285
                  Left            =   600
                  TabIndex        =   82
                  Top             =   280
                  Width           =   255
                  _ExtentX        =   445
                  _ExtentY        =   508
                  _Version        =   327681
                  Value           =   1
                  BuddyControl    =   "MaxPopText"
                  BuddyDispid     =   196677
                  OrigLeft        =   660
                  OrigTop         =   300
                  OrigRight       =   900
                  OrigBottom      =   585
                  Increment       =   25
                  Max             =   5000
                  Min             =   1
                  SyncBuddy       =   -1  'True
                  BuddyProperty   =   65547
                  Enabled         =   -1  'True
               End
               Begin ComCtl2.UpDown MinVegUpDn 
                  Height          =   285
                  Left            =   600
                  TabIndex        =   83
                  Top             =   660
                  Width           =   255
                  _ExtentX        =   445
                  _ExtentY        =   508
                  _Version        =   327681
                  Value           =   100
                  BuddyControl    =   "MinVegText"
                  BuddyDispid     =   196676
                  OrigLeft        =   2730
                  OrigTop         =   2295
                  OrigRight       =   2970
                  OrigBottom      =   2580
                  Increment       =   10
                  Max             =   950
                  Min             =   20
                  SyncBuddy       =   -1  'True
                  BuddyProperty   =   65547
                  Enabled         =   -1  'True
               End
               Begin ComCtl2.UpDown RepopCooldownUpDown 
                  Height          =   285
                  Index           =   1
                  Left            =   600
                  TabIndex        =   84
                  Top             =   1380
                  Width           =   255
                  _ExtentX        =   445
                  _ExtentY        =   508
                  _Version        =   327681
                  BuddyControl    =   "RepopCooldownText"
                  BuddyDispid     =   196674
                  OrigLeft        =   660
                  OrigTop         =   1380
                  OrigRight       =   900
                  OrigBottom      =   1665
                  SyncBuddy       =   -1  'True
                  BuddyProperty   =   65547
                  Enabled         =   -1  'True
               End
               Begin VB.Label LabelZZ 
                  Caption         =   "* 16K Maximum Chloroplasts"
                  Height          =   195
                  Index           =   0
                  Left            =   960
                  TabIndex        =   88
                  ToolTipText     =   "An average Vegy has 16000 Chloroplasts."
                  Top             =   315
                  Width           =   2295
               End
               Begin VB.Label LabelZZ 
                  Caption         =   "* 16K Repopulation threshold of Chloroplasts"
                  Height          =   435
                  Index           =   1
                  Left            =   960
                  TabIndex        =   87
                  ToolTipText     =   "An average Vegy has 16000 Chloroplasts."
                  Top             =   600
                  Width           =   2055
               End
               Begin VB.Label LabelZZ 
                  Caption         =   "Robots per repopulation event"
                  Height          =   195
                  Index           =   2
                  Left            =   960
                  TabIndex        =   86
                  Top             =   1100
                  Width           =   2235
               End
               Begin VB.Label LabelZZ 
                  Caption         =   "Repopulation cooldown period"
                  Height          =   195
                  Index           =   3
                  Left            =   960
                  TabIndex        =   85
                  Top             =   1440
                  Width           =   2235
               End
            End
            Begin VB.Frame Frame30 
               Caption         =   "Initial light energy"
               Height          =   735
               Left            =   3720
               TabIndex        =   73
               Top             =   240
               Width           =   1695
               Begin VB.TextBox MaxNRGText 
                  Height          =   285
                  Left            =   180
                  TabIndex        =   74
                  Text            =   "40"
                  ToolTipText     =   "Amount of energy to give each veggy each cycle.  Is overrided by Pond Mode settings if that's on."
                  Top             =   300
                  Width           =   735
               End
               Begin VB.Label VegNRG 
                  Caption         =   "NRG"
                  Height          =   195
                  Left            =   960
                  TabIndex        =   75
                  Top             =   360
                  Width           =   690
               End
            End
         End
         Begin VB.Frame Frame18 
            Caption         =   "Corpse Mode"
            Height          =   2955
            Left            =   2280
            TabIndex        =   59
            Top             =   1200
            Width           =   1815
            Begin VB.CheckBox CorpseCheck 
               Alignment       =   1  'Right Justify
               Caption         =   "Enable"
               Height          =   255
               Left            =   360
               TabIndex        =   71
               ToolTipText     =   "Enable corpse mode.  Corpses are dead robots who still have carcasses that can be eaten by other robots."
               Top             =   240
               Width           =   1185
            End
            Begin VB.Frame Frame3 
               Caption         =   "Decay Rate"
               Height          =   1005
               Left            =   240
               TabIndex        =   64
               Top             =   1920
               Width           =   1455
               Begin VB.TextBox FrequencyText 
                  Height          =   285
                  Left            =   600
                  TabIndex        =   66
                  Text            =   "1"
                  ToolTipText     =   "How many cycles per shot?"
                  Top             =   600
                  Width           =   435
               End
               Begin VB.TextBox DecayText 
                  Height          =   285
                  Left            =   600
                  TabIndex        =   65
                  Text            =   "2"
                  ToolTipText     =   "How large is the decay shot?"
                  Top             =   240
                  Width           =   435
               End
               Begin ComCtl2.UpDown DecayUpDn 
                  Height          =   285
                  Index           =   0
                  Left            =   1080
                  TabIndex        =   67
                  ToolTipText     =   "Set how fast you want your corpses to decay away."
                  Top             =   240
                  Width           =   255
                  _ExtentX        =   445
                  _ExtentY        =   508
                  _Version        =   327681
                  BuddyControl    =   "DecayText"
                  BuddyDispid     =   196685
                  OrigLeft        =   5040
                  OrigTop         =   2280
                  OrigRight       =   5280
                  OrigBottom      =   2565
                  Max             =   100
                  SyncBuddy       =   -1  'True
                  BuddyProperty   =   0
                  Enabled         =   -1  'True
               End
               Begin ComCtl2.UpDown DecayUpDn 
                  Height          =   285
                  Index           =   1
                  Left            =   1080
                  TabIndex        =   68
                  ToolTipText     =   "Set how fast you want your corpses to decay away."
                  Top             =   600
                  Width           =   255
                  _ExtentX        =   445
                  _ExtentY        =   508
                  _Version        =   327681
                  BuddyControl    =   "FrequencyText"
                  BuddyDispid     =   196684
                  OrigLeft        =   5040
                  OrigTop         =   2280
                  OrigRight       =   5280
                  OrigBottom      =   2565
                  Max             =   100
                  SyncBuddy       =   -1  'True
                  BuddyProperty   =   0
                  Enabled         =   -1  'True
               End
               Begin VB.Label FrequencyLabel1 
                  Caption         =   "Period"
                  Height          =   255
                  Left            =   90
                  TabIndex        =   70
                  Top             =   660
                  Width           =   810
               End
               Begin VB.Label Label59 
                  Caption         =   "Size"
                  Height          =   255
                  Left            =   90
                  TabIndex        =   69
                  Top             =   300
                  Width           =   630
               End
            End
            Begin VB.Frame Frame9 
               Caption         =   "Decay Type"
               Height          =   1215
               Left            =   240
               TabIndex        =   60
               Top             =   600
               Width           =   1455
               Begin VB.OptionButton DecayOption 
                  Caption         =   "None"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   63
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.OptionButton DecayOption 
                  Caption         =   "Waste"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   62
                  Top             =   540
                  Width           =   1095
               End
               Begin VB.OptionButton DecayOption 
                  Caption         =   "NRG"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   61
                  Top             =   840
                  Width           =   1095
               End
            End
         End
         Begin VB.Label Label27 
            Caption         =   "Energy per veggie per cycle"
            Height          =   210
            Left            =   1260
            TabIndex        =   122
            Tag             =   "2201"
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin VB.CommandButton DuplicaButt 
         Caption         =   "Duplicate"
         Height          =   375
         Left            =   1440
         TabIndex        =   56
         Tag             =   "0"
         ToolTipText     =   "Add a new robot type to the simulation"
         Top             =   3720
         Width           =   1095
      End
      Begin VB.ListBox SpecList 
         Height          =   2352
         ItemData        =   "OptionsForm.frx":038B
         Left            =   240
         List            =   "OptionsForm.frx":038D
         TabIndex        =   55
         Top             =   720
         Width           =   3465
      End
      Begin VB.CommandButton AddSpec 
         Caption         =   "Add"
         Height          =   375
         Left            =   240
         TabIndex        =   54
         Tag             =   "0"
         ToolTipText     =   "Add a new robot type to the simulation"
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton DelSpec 
         Caption         =   "Delete"
         Height          =   375
         Left            =   240
         TabIndex        =   53
         Tag             =   "0"
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton RenameButton 
         Caption         =   "Rename"
         Height          =   375
         Left            =   2640
         TabIndex        =   52
         Tag             =   "0"
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Frame Frame13 
         Caption         =   "Current Multiplier"
         Height          =   2115
         Left            =   -74700
         TabIndex        =   18
         Top             =   1320
         Width           =   4755
         Begin MSComctlLib.Slider MutSlide 
            Height          =   225
            Left            =   360
            TabIndex        =   19
            Top             =   420
            Width           =   4215
            _ExtentX        =   7430
            _ExtentY        =   402
            _Version        =   393216
            LargeChange     =   1
            Min             =   -5
            Max             =   5
         End
         Begin VB.Label Label31 
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"OptionsForm.frx":038F
            Height          =   975
            Left            =   300
            TabIndex        =   24
            Tag             =   "5005"
            Top             =   1020
            Width           =   4350
         End
         Begin VB.Label Label30 
            Caption         =   "1"
            Height          =   255
            Left            =   2400
            TabIndex        =   23
            Top             =   720
            Width           =   255
         End
         Begin VB.Label MutLab 
            Caption         =   "1 X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2280
            TabIndex        =   22
            Top             =   150
            Width           =   975
         End
         Begin VB.Label Label15 
            Caption         =   "alta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4305
            TabIndex        =   21
            Tag             =   "5004"
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label14 
            Caption         =   "bassa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   225
            TabIndex        =   20
            Tag             =   "5003"
            Top             =   720
            Width           =   495
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "Costs and Returned Shots"
         Height          =   5235
         Left            =   -70080
         TabIndex        =   6
         Top             =   420
         Width           =   4815
         Begin ComCtl2.UpDown PropUpDn 
            Height          =   285
            Left            =   3960
            TabIndex        =   14
            Top             =   3000
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   445
            _ExtentY        =   508
            _Version        =   327681
            BuddyControl    =   "Prop"
            BuddyDispid     =   196739
            OrigLeft        =   4260
            OrigTop         =   2760
            OrigRight       =   4500
            OrigBottom      =   3090
            Increment       =   2
            Max             =   10000
            Enabled         =   -1  'True
         End
         Begin VB.OptionButton ExchangeProp 
            Caption         =   "Proportional"
            Height          =   225
            Left            =   2160
            TabIndex        =   12
            Top             =   3000
            Width           =   1185
         End
         Begin VB.OptionButton ExchangeFix 
            Caption         =   "Fixed Nrg"
            Height          =   225
            Left            =   2160
            TabIndex        =   11
            Top             =   2520
            Width           =   1125
         End
         Begin VB.TextBox Fixed 
            Height          =   285
            Left            =   3480
            TabIndex        =   10
            Text            =   "0"
            Top             =   2520
            Width           =   510
         End
         Begin VB.TextBox Prop 
            Height          =   285
            Left            =   3480
            TabIndex        =   9
            Text            =   "0"
            Top             =   3000
            Width           =   510
         End
         Begin ComCtl2.UpDown FixUpDn 
            Height          =   285
            Left            =   3960
            TabIndex        =   13
            Top             =   2520
            Width           =   255
            _ExtentX        =   445
            _ExtentY        =   508
            _Version        =   327681
            BuddyControl    =   "Fixed"
            BuddyDispid     =   196738
            OrigLeft        =   4245
            OrigTop         =   2340
            OrigRight       =   4485
            OrigBottom      =   2670
            Increment       =   10
            Max             =   10000
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Frame Frame28 
            Caption         =   "Costs"
            Height          =   1935
            Left            =   120
            TabIndex        =   7
            Top             =   300
            Width           =   4575
            Begin VB.OptionButton CostRadio 
               Caption         =   "Custom"
               Height          =   495
               Index           =   2
               Left            =   600
               TabIndex        =   144
               Top             =   1200
               Width           =   1215
            End
            Begin VB.OptionButton CostRadio 
               Caption         =   "F1 Default"
               Height          =   495
               Index           =   1
               Left            =   600
               TabIndex        =   143
               Top             =   720
               Width           =   1215
            End
            Begin VB.OptionButton CostRadio 
               Caption         =   "No Costs"
               Height          =   495
               Index           =   0
               Left            =   600
               TabIndex        =   142
               Top             =   240
               Width           =   1215
            End
            Begin VB.CommandButton ToCosts 
               Caption         =   "Custom Simulation Costs"
               Height          =   375
               Left            =   2040
               TabIndex        =   8
               Top             =   1320
               Width           =   1875
            End
         End
         Begin VB.Label Label8 
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"OptionsForm.frx":044A
            Height          =   945
            Left            =   225
            TabIndex        =   17
            Tag             =   "6009"
            Top             =   3840
            Width           =   4335
         End
         Begin VB.Label Label16 
            Caption         =   "Shot Energy Exchange Method"
            Height          =   495
            Left            =   225
            TabIndex        =   16
            Top             =   2640
            Width           =   1905
         End
         Begin VB.Label Label10 
            Caption         =   "%"
            Height          =   195
            Left            =   4320
            TabIndex        =   15
            Top             =   3050
            Width           =   135
         End
      End
      Begin RichTextLib.RichTextBox CommentBox 
         Height          =   975
         Left            =   240
         TabIndex        =   57
         Top             =   4680
         Width           =   3495
         _ExtentX        =   6160
         _ExtentY        =   1715
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"OptionsForm.frx":050F
      End
      Begin VB.Frame Frame1 
         Caption         =   "Species Properties"
         Height          =   5085
         Left            =   3840
         TabIndex        =   25
         Tag             =   "2010"
         Top             =   600
         Width           =   6075
         Begin VB.CommandButton btnFatalRES 
            Caption         =   "Fatal Restrictions"
            Height          =   360
            Left            =   3260
            TabIndex        =   164
            Top             =   4480
            Width           =   1400
         End
         Begin VB.CheckBox chkNoChlr 
            Alignment       =   1  'Right Justify
            Caption         =   "Disable Chloroplasts"
            Height          =   210
            Left            =   3360
            TabIndex        =   163
            ToolTipText     =   "Feed automatically this type (only for vegetables)"
            Top             =   2400
            Width           =   2430
         End
         Begin VB.CheckBox MutEnabledCheck 
            Alignment       =   1  'Right Justify
            Caption         =   "Disable Mutations"
            Height          =   330
            Left            =   3360
            TabIndex        =   157
            ToolTipText     =   "Disables voluntary movement for this species"
            Top             =   4080
            Width           =   2430
         End
         Begin VB.CheckBox VirusImmuneCheck 
            Alignment       =   1  'Right Justify
            Caption         =   "Virus Immune"
            Height          =   330
            Left            =   3360
            TabIndex        =   156
            ToolTipText     =   "Disables voluntary movement for this species"
            Top             =   3840
            Width           =   2430
         End
         Begin VB.CheckBox DisableReproductionCheck 
            Alignment       =   1  'Right Justify
            Caption         =   "Disable Reproduction"
            Height          =   330
            Left            =   3360
            TabIndex        =   155
            ToolTipText     =   "Disables voluntary movement for this species"
            Top             =   3600
            Width           =   2430
         End
         Begin VB.CheckBox DisableMovementSysvarsCheck 
            Alignment       =   1  'Right Justify
            Caption         =   "Disable Voluntary Movement"
            Height          =   330
            Left            =   3360
            TabIndex        =   154
            ToolTipText     =   "Disables voluntary movement for this species"
            Top             =   3360
            Width           =   2430
         End
         Begin VB.CheckBox DisableDNACheck 
            Alignment       =   1  'Right Justify
            Caption         =   "Disable DNA Execution"
            Height          =   330
            Left            =   3360
            TabIndex        =   153
            ToolTipText     =   "Speeds up the simulation by turning off DNA execution for this species"
            Top             =   3120
            Width           =   2430
         End
         Begin VB.CheckBox DisableVisionCheck 
            Alignment       =   1  'Right Justify
            Caption         =   "Disable Vision"
            Height          =   330
            Left            =   3360
            TabIndex        =   152
            ToolTipText     =   "Speeds up the simulation by turning off vision for this species"
            Top             =   2880
            Width           =   2430
         End
         Begin VB.CommandButton InNrg 
            Caption         =   "30000"
            Height          =   255
            Index           =   3
            Left            =   5280
            TabIndex        =   45
            Top             =   1800
            Width           =   615
         End
         Begin VB.CommandButton InNrg 
            Caption         =   "5000"
            Height          =   255
            Index           =   2
            Left            =   4680
            TabIndex        =   44
            Top             =   1800
            Width           =   615
         End
         Begin VB.CommandButton InNrg 
            Caption         =   "3000"
            Height          =   255
            Index           =   0
            Left            =   4080
            TabIndex        =   43
            Top             =   1800
            Width           =   615
         End
         Begin VB.CommandButton IndNum 
            Caption         =   "30"
            Height          =   255
            Index           =   3
            Left            =   5400
            TabIndex        =   42
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton IndNum 
            Caption         =   "15"
            Height          =   255
            Index           =   2
            Left            =   4920
            TabIndex        =   41
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton IndNum 
            Caption         =   "5"
            Height          =   255
            Index           =   0
            Left            =   4440
            TabIndex        =   40
            Top             =   840
            Width           =   495
         End
         Begin VB.CheckBox BlockSpec 
            Alignment       =   1  'Right Justify
            Caption         =   "Fixed in place"
            Height          =   330
            Left            =   3360
            TabIndex        =   39
            Top             =   2640
            Width           =   2430
         End
         Begin VB.CheckBox SpecVeg 
            Alignment       =   1  'Right Justify
            Caption         =   "Repopulate"
            Height          =   210
            Left            =   3360
            TabIndex        =   36
            ToolTipText     =   "Feed automatically this type (only for vegetables)"
            Top             =   2160
            Width           =   2430
         End
         Begin VB.TextBox SpecQty 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5040
            TabIndex        =   35
            Text            =   "0"
            Top             =   480
            Width           =   585
         End
         Begin VB.ComboBox SpecCol 
            Height          =   315
            ItemData        =   "OptionsForm.frx":0591
            Left            =   1350
            List            =   "OptionsForm.frx":05B3
            Style           =   2  'Dropdown List
            TabIndex        =   34
            ToolTipText     =   "Initial colour"
            Top             =   1770
            Width           =   1380
         End
         Begin VB.TextBox SpecNrg 
            Alignment       =   1  'Right Justify
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   5040
            TabIndex        =   33
            Text            =   "0"
            Top             =   1440
            Width           =   585
         End
         Begin VB.CommandButton MutRatesBut 
            Caption         =   "Mutation Rates"
            Height          =   360
            Left            =   4700
            TabIndex        =   32
            Tag             =   "2016"
            ToolTipText     =   "Sets the initial mutation rates for this type"
            Top             =   4480
            Width           =   1215
         End
         Begin VB.CommandButton IndNum 
            Caption         =   "0"
            Height          =   255
            Index           =   4
            Left            =   3960
            TabIndex        =   31
            Top             =   840
            Width           =   495
         End
         Begin VB.CommandButton PosReset 
            Caption         =   "Reset"
            Height          =   375
            Left            =   1680
            TabIndex        =   30
            Top             =   2280
            Width           =   1095
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   285
            Left            =   5640
            TabIndex        =   46
            ToolTipText     =   "Set the initial number of copies for this robot type"
            Top             =   480
            Width           =   240
            _ExtentX        =   445
            _ExtentY        =   508
            _Version        =   327681
            BuddyControl    =   "SpecQty"
            BuddyDispid     =   196759
            OrigLeft        =   2310
            OrigTop         =   525
            OrigRight       =   2550
            OrigBottom      =   810
            Max             =   200
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown3 
            Height          =   285
            Left            =   5640
            TabIndex        =   47
            ToolTipText     =   "Initial energy assigned to this type"
            Top             =   1440
            Width           =   240
            _ExtentX        =   445
            _ExtentY        =   508
            _Version        =   327681
            Value           =   1000
            BuddyControl    =   "SpecNrg"
            BuddyDispid     =   196761
            OrigLeft        =   1560
            OrigTop         =   240
            OrigRight       =   1800
            OrigBottom      =   615
            Increment       =   500
            Max             =   32000
            Min             =   10
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Frame Frame6 
            Caption         =   "Skin"
            Height          =   1200
            Left            =   195
            TabIndex        =   37
            Tag             =   "2121"
            Top             =   390
            Width           =   2565
            Begin VB.CommandButton SkinChange 
               Caption         =   "Change"
               Height          =   330
               Left            =   285
               TabIndex        =   38
               Tag             =   "2120"
               Top             =   450
               Width           =   855
            End
            Begin VB.Shape Cerchio 
               BorderColor     =   &H8000000C&
               BorderWidth     =   4
               FillColor       =   &H008080FF&
               Height          =   645
               Left            =   1515
               Shape           =   3  'Circle
               Top             =   315
               Width           =   645
            End
            Begin VB.Shape Shape3 
               FillColor       =   &H00FFFFFF&
               FillStyle       =   0  'Solid
               Height          =   165
               Left            =   2115
               Shape           =   3  'Circle
               Top             =   555
               Width           =   150
            End
            Begin VB.Line Line9 
               BorderColor     =   &H00000000&
               BorderWidth     =   2
               X1              =   1635
               X2              =   2055
               Y1              =   690
               Y2              =   690
            End
            Begin VB.Line Line8 
               BorderColor     =   &H00000000&
               BorderWidth     =   2
               X1              =   1605
               X2              =   2100
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Line Line7 
               BorderColor     =   &H00000000&
               BorderWidth     =   2
               X1              =   1620
               X2              =   2070
               Y1              =   510
               Y2              =   510
            End
            Begin VB.Shape Shape2 
               FillStyle       =   0  'Solid
               Height          =   885
               Left            =   1410
               Top             =   210
               Width           =   900
            End
         End
         Begin VB.PictureBox IPB 
            AutoRedraw      =   -1  'True
            Height          =   1695
            Left            =   120
            ScaleHeight     =   1644
            ScaleWidth      =   2724
            TabIndex        =   26
            Top             =   2760
            Width           =   2775
            Begin VB.PictureBox Initial_Position 
               BackColor       =   &H00FF0000&
               FillColor       =   &H000000FF&
               FillStyle       =   0  'Solid
               Height          =   1455
               Left            =   120
               ScaleHeight     =   1404
               ScaleWidth      =   2484
               TabIndex        =   29
               Top             =   120
               Width           =   2535
            End
            Begin VB.PictureBox RobPlacLine 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               DrawStyle       =   5  'Transparent
               ForeColor       =   &H80000008&
               Height          =   135
               Index           =   0
               Left            =   600
               ScaleHeight     =   132
               ScaleWidth      =   1212
               TabIndex        =   28
               Top             =   960
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.PictureBox picHandle 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000D&
               BorderStyle     =   0  'None
               CausesValidation=   0   'False
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   1200
               ScaleHeight     =   252
               ScaleWidth      =   252
               TabIndex        =   27
               Top             =   960
               Visible         =   0   'False
               Width           =   255
            End
         End
         Begin VB.Line Line4 
            BorderColor     =   &H8000000C&
            X1              =   3120
            X2              =   3120
            Y1              =   180
            Y2              =   4920
         End
         Begin VB.Label Label2 
            Caption         =   "Individui"
            Height          =   255
            Left            =   3360
            TabIndex        =   51
            Tag             =   "2011"
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Posizione iniziale"
            Height          =   255
            Left            =   165
            TabIndex        =   50
            Tag             =   "2014"
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Colore"
            Height          =   225
            Left            =   360
            TabIndex        =   49
            Tag             =   "2015"
            Top             =   1830
            Width           =   750
         End
         Begin VB.Label Label6 
            Caption         =   "Energia iniziale"
            Height          =   255
            Left            =   3360
            TabIndex        =   48
            Tag             =   "2012"
            Top             =   1440
            Width           =   1455
         End
      End
      Begin VB.Label Label36 
         Caption         =   "Commenti sulla specie:"
         Height          =   255
         Left            =   240
         TabIndex        =   137
         Tag             =   "2100"
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label SpeciesLabel 
         Caption         =   "Native Species:"
         Height          =   255
         Left            =   240
         TabIndex        =   136
         Tag             =   "0"
         Top             =   480
         Width           =   3375
      End
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   495
      Left            =   4440
      TabIndex        =   138
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "optionsform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Botsareus 5/8/2012 Selected the first tab to be the default when window first opens (vb generated code above may have changed)

' boring stuff. this is the sim options module, so
' there's much graphical interface (and a big mess).
' simulation options are saved in a structure
' called SimOptions. When the options window is opened
' the current settings SimOpts are copied to TmpOpts
' which is modified. On exit, if ok has been clicked
' TmpOpts is copied to SimOpts again
' SimOpts struct is defined in SimOptModule

Option Explicit

Private validate As Boolean

Dim pass As Boolean

Dim follow1 As Boolean
Dim follow2 As Boolean
Dim follow3 As Boolean
Dim follow4 As Boolean

'Dim SpeciesToggle As Boolean 'Botsareus 1/21/2013 no more need for speices toggle

Dim lastsettings As String
Dim contrmethod As Integer
Public CurrSpec As Integer

Public col1 As Long
Public Canc As Boolean

Public IPBWidth As Long
Public IPBHeight As Long

Dim multx As Long
Dim multy As Long

'Windows declarations
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const NULL_BRUSH = 5
Private Const PS_SOLID = 0
Private Const R2_NOT = 6

Enum ControlState
    StateNothing = 0
    StateDragging
    StateSizing
End Enum

Private m_CurrCtl As Control
Private m_DragState As ControlState
Private m_DragHandle As Integer
Private m_DragRect As New CRect
Private m_DragPoint As POINTAPI
Private speclistchecked As Boolean



Private Sub btnFatalRES_Click()
'Botsareus 7/29/2014 Show restrictions
  If optionsform.CurrSpec < 0 Then
   MsgBox ("Please select a species.")
  Else
    frmRestriOps.res_state = IIf(TmpOpts.Specie(optionsform.CurrSpec).Veg, 0, 1)
    frmRestriOps.Show vbModal
  End If
End Sub


Private Sub chkNoChlr_Click()
  If CurrSpec > -1 Then
    TmpOpts.Specie(CurrSpec).NoChlr = False
    If chkNoChlr.value = 1 Then
        TmpOpts.Specie(CurrSpec).NoChlr = True
        TmpOpts.Specie(CurrSpec).Veg = False
        SpecVeg.value = 0
    End If
  End If
End Sub

Private Sub Gradient_GotFocus()
tmrLight.Enabled = True
End Sub

Private Sub LightText_GotFocus()
tmrLight.Enabled = True
End Sub

Private Sub tmrLight_Timer()
Static dir As Boolean
Dim n As Integer
If dir Then
    n = EnergyScalingFactor + 5
    If n > 100 Then
        dir = False
        n = 100
    End If
Else
    n = EnergyScalingFactor - 5
    If n < 2 Then
        dir = True
        n = 2
    End If
End If
EnergyScalingFactor = n
End Sub


Private Sub CorpseCheck_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
If CorpseCheck.value = 1 Then 'Botsareus 1/17/2013 set default values
    TmpOpts.DecayType = 3
    TmpOpts.Decay = 75
    TmpOpts.Decaydelay = 3
    DispSettings
  End If
End Sub

Private Sub DisableDNACheck_Click()
If CurrSpec > -1 Then
    TmpOpts.Specie(CurrSpec).DisableDNA = False
    If DisableDNACheck.value = 1 Then TmpOpts.Specie(CurrSpec).DisableDNA = True
  End If
End Sub

Private Sub DisableMovementSysvarsCheck_Click()
  If CurrSpec > -1 Then
    TmpOpts.Specie(CurrSpec).DisableMovementSysvars = False
    If DisableMovementSysvarsCheck.value = 1 Then TmpOpts.Specie(CurrSpec).DisableMovementSysvars = True
  End If
End Sub

Private Sub DisableMutationsCheck_Click()
  TmpOpts.DisableMutations = DisableMutationsCheck.value * True
End Sub

Private Sub DisableReproductionCheck_Click()
  If CurrSpec > -1 Then
    TmpOpts.Specie(CurrSpec).CantReproduce = False
    If DisableReproductionCheck.value = 1 Then TmpOpts.Specie(CurrSpec).CantReproduce = True
  End If
End Sub

Private Sub DisableVisionCheck_Click()
  If CurrSpec > -1 Then
    TmpOpts.Specie(CurrSpec).CantSee = False
    If DisableVisionCheck.value = 1 Then TmpOpts.Specie(CurrSpec).CantSee = True
  End If
End Sub


Private Sub Elasticity_Change()
  TmpOpts.CoefficientElasticity = Elasticity.value / 10
  Elasticity.Text = Elasticity.value / 10
 ' Elasticity.Refresh
End Sub

Private Sub Energy_Click()
  EnergyForm.Show vbModal
  Update
End Sub

'Botsareus 9/1/2014 Update main presets
Private Sub Update()
BrownianCombo.Visible = TmpOpts.Tides = 0
GravityCombo.Visible = TmpOpts.Tides = 0
Label7(2).Visible = TmpOpts.Tides = 0
Label7(3).Visible = TmpOpts.Tides = 0
Label47.Visible = TmpOpts.Tides > 0
  
If TmpOpts.PhysMoving < 0.33 Then EfficiencyCombo.Text = EfficiencyCombo.List(2)
If TmpOpts.PhysMoving > 0.33 And TmpOpts.PhysMoving < 0.66 Then EfficiencyCombo.Text = EfficiencyCombo.List(1)
If TmpOpts.PhysMoving > 0.66 Then EfficiencyCombo.Text = EfficiencyCombo.List(0)

If TmpOpts.PhysMoving = 0.33 Then EfficiencyCombo.Text = EfficiencyCombo.List(2)
If TmpOpts.PhysMoving = 0.66 Then EfficiencyCombo.Text = EfficiencyCombo.List(1)
If TmpOpts.PhysMoving = 1 Then EfficiencyCombo.Text = EfficiencyCombo.List(0)
  
'

If TmpOpts.PhysBrown > 7 Then BrownianCombo.Text = BrownianCombo.List(0)
If TmpOpts.PhysBrown > 0.5 And TmpOpts.PhysBrown < 7 Then BrownianCombo.Text = BrownianCombo.List(1)
If TmpOpts.PhysBrown < 0.5 Then BrownianCombo.Text = BrownianCombo.List(2)

If TmpOpts.PhysBrown = 7 Then BrownianCombo.Text = BrownianCombo.List(0)
If TmpOpts.PhysBrown = 0.5 Then BrownianCombo.Text = BrownianCombo.List(1)
If TmpOpts.PhysBrown = 0 Then BrownianCombo.Text = BrownianCombo.List(2)

'

If TmpOpts.Ygravity < 0.1 Then GravityCombo.Text = GravityCombo.List(0)
If TmpOpts.Ygravity > 0.1 And TmpOpts.Ygravity < 0.3 Then GravityCombo.Text = GravityCombo.List(1)
If TmpOpts.Ygravity > 0.3 And TmpOpts.Ygravity < 0.9 Then GravityCombo.Text = GravityCombo.List(2)
If TmpOpts.Ygravity > 0.9 And TmpOpts.Ygravity < 9 Then GravityCombo.Text = GravityCombo.List(3)
If TmpOpts.Ygravity > 6 Then GravityCombo.Text = GravityCombo.List(4)

If TmpOpts.Ygravity = 0 Then GravityCombo.Text = GravityCombo.List(0)
If TmpOpts.Ygravity = 0.1 Then GravityCombo.Text = GravityCombo.List(1)
If TmpOpts.Ygravity = 0.3 Then GravityCombo.Text = GravityCombo.List(2)
If TmpOpts.Ygravity = 0.9 Then GravityCombo.Text = GravityCombo.List(3)
If TmpOpts.Ygravity = 6 Then GravityCombo.Text = GravityCombo.List(4)

End Sub

Private Sub FixBotRadius_Click()
  TmpOpts.FixedBotRadii = FixBotRadius.value * True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) 'Botsareus 1/5/2013 moved cancel code here
  Form1.camfix = False 'Botsareus 2/23/2013 re-normalize screen
  Canc = True
End Sub

Private Sub Gradient_LostFocus()
'Botsareus 3/26/2013 applyed caps to gradient
Gradient.Text = val(Gradient.Text)
If Gradient.Text > 200 Then Gradient.Text = 200
If Gradient.Text < 0 Then Gradient.Text = 0
tmrLight.Enabled = False
End Sub

Private Sub gset_Click()
frmGset.tb.Tab = 0
frmGset.Show vbModal, Me
End Sub

Private Sub MutEnabledCheck_Click()
 If CurrSpec > -1 Then
 

     MutEnabledCheck.Enabled = True

    If MutEnabledCheck.value = 1 Then
      TmpOpts.Specie(CurrSpec).Mutables.Mutations = False
      MutRatesBut.Enabled = False
    Else
      TmpOpts.Specie(CurrSpec).Mutables.Mutations = True
      MutRatesBut.Enabled = True
    End If
  End If
End Sub

Public Sub PopulateSpeciesList()
Dim i As Integer

 SortSpecies
    SpecList.CLEAR
    For i = 0 To TmpOpts.SpeciesNum - 1
      SpecList.additem (TmpOpts.Specie(i).Name)
    Next i
 
  SpecList.Refresh
End Sub

Private Sub Newseed_Click()
    UserSeedText = Timer * 100
End Sub

Private Sub RenameButton_Click()
Dim ind As Integer
 
  ind = optionsform.SpecList.ListIndex
  If ind >= 0 Then
    'RenameForm.Show vbModal
  Else
    MsgBox ("No Species Selected.")
  End If
  
  PopulateSpeciesList
End Sub


Private Sub ToCosts_Click()
  CostsForm.Show vbModal
  Unload CostsForm
End Sub


Private Sub CostRadio_Click(Index As Integer)
Dim k As Integer
Select Case Index
  Case 0: 'No Costs
    ToCosts.Enabled = False
    For k = 0 To 70
      TmpOpts.Costs(k) = 0
    Next k
  Case 1: 'F1 Costs
    ToCosts.Enabled = False
    For k = 0 To 70
      TmpOpts.Costs(k) = 0
    Next k
    TmpOpts.Costs(5) = 0.004
    TmpOpts.Costs(7) = 0.04
    TmpOpts.Costs(20) = 0.05
    TmpOpts.Costs(22) = 2
    TmpOpts.Costs(23) = 2
    TmpOpts.Costs(26) = 0.01
    TmpOpts.Costs(27) = 0.01
    TmpOpts.Costs(28) = 0.1
    TmpOpts.Costs(29) = 0.1
    TmpOpts.Costs(BODYUPKEEP) = 0.00001
    TmpOpts.Costs(AGECOST) = 0.01
    TmpOpts.Costs(COSTMULTIPLIER) = 1
  Case 2: 'Custom
    ToCosts.Enabled = True
  End Select
  TmpOpts.CostRadioSetting = Index
End Sub

Private Sub FluidSolidRadio_Click(Index As Integer)
  Select Case Index
  Case 0: 'Fluid
    FrictionCombo.Text = FrictionCombo.List(3)
    FrictionCombo_Click
    FrictionCombo.Enabled = False
    DragCombo.Enabled = True
    ToPhysics.Enabled = False
    DragCombo_Click
 Case 1: 'Solid
    DragCombo.Text = DragCombo.List(3)
    DragCombo_Click
    FrictionCombo.Enabled = True
    DragCombo.Enabled = False
    ToPhysics.Enabled = False
    FrictionCombo_Click
  Case 2: 'Custom
    FrictionCombo.Enabled = False
    DragCombo.Enabled = False
    ToPhysics.Enabled = True
  End Select
  TmpOpts.FluidSolidCustom = Index
End Sub

Private Sub MaxVelSlider_Change()
  TmpOpts.MaxVelocity = MaxVelSlider.value
End Sub


''''''''''''''''''''''''''''''''

'Species Tab

''''''''''''''''''''''''''''''''

Private Sub SpecList_Click()
  Dim k As Integer
  Dim cmaxx As Long
  Dim cminx As Long
  Dim cmaxy As Long
  Dim cminy As Long
  Dim w As Long
  Dim h As Long
  
  w = IPB.Width
  h = IPB.Height
  
  
  enprop
  DragEnd
  k = SpecList.ListIndex
  CurrSpec = k
  ShowSkin k
  speclistchecked = True
  SpecCol.ListIndex = TmpOpts.Specie(k).Colind 'EricL 3/21/2006 - Set the Color Drop Down to match the selected species
  speclistchecked = False
  Frame1.Caption = WSproperties + TmpOpts.Specie(k).Name
  SpecQty.Text = Str$(TmpOpts.Specie(k).qty)
  SpecNrg.Text = Str$(TmpOpts.Specie(k).Stnrg)
  Cerchio.FillColor = TmpOpts.Specie(k).color
  Cerchio.BorderColor = TmpOpts.Specie(k).color
  Line7.BorderColor = TmpOpts.Specie(k).color
  Line8.BorderColor = TmpOpts.Specie(k).color
  Line9.BorderColor = TmpOpts.Specie(k).color
  If TmpOpts.Specie(k).Veg Then
    SpecVeg.value = 1
  Else
    SpecVeg.value = 0
  End If
  If TmpOpts.Specie(k).NoChlr Then  'Botsareus 3/28/2014 Disable chloroplasts
    chkNoChlr.value = 1
  Else
    chkNoChlr.value = 0
  End If
  
  If TmpOpts.Specie(k).CantSee Then
    DisableVisionCheck.value = 1
  Else
    DisableVisionCheck.value = 0
  End If
  
  If TmpOpts.Specie(k).DisableDNA Then
    DisableDNACheck.value = 1
  Else
    DisableDNACheck.value = 0
  End If
  
  If TmpOpts.Specie(k).CantReproduce Then
    DisableReproductionCheck.value = 1
  Else
    DisableReproductionCheck.value = 0
  End If
    
  If TmpOpts.Specie(k).DisableMovementSysvars Then
    DisableMovementSysvarsCheck.value = 1
  Else
    DisableMovementSysvarsCheck.value = 0
  End If
  
 If TmpOpts.Specie(k).VirusImmune Then
    VirusImmuneCheck.value = 1
  Else
    VirusImmuneCheck.value = 0
  End If
  
  If TmpOpts.Specie(k).Mutables.Mutations Then
    MutEnabledCheck.value = 0
  Else
    MutEnabledCheck.value = 1
  End If
  
  If TmpOpts.Specie(k).Fixed Then
    BlockSpec.value = 1
  Else
    BlockSpec.value = 0
  End If
  
  
  CommentBox.Text = TmpOpts.Specie(k).Comment
  
  If TmpOpts.Specie(k).Poslf < 0 Then TmpOpts.Specie(k).Poslf = 0
  If TmpOpts.Specie(k).Postp < 0 Then TmpOpts.Specie(k).Postp = 0
  
  If TmpOpts.Specie(k).Poslf > 1 Then TmpOpts.Specie(k).Poslf = 0
  If TmpOpts.Specie(k).Postp > 1 Then TmpOpts.Specie(k).Postp = 0
  
  If TmpOpts.Specie(k).Posrg > 1 Then TmpOpts.Specie(k).Posrg = 1
  If TmpOpts.Specie(k).Posdn > 1 Then TmpOpts.Specie(k).Posdn = 1
  
  If TmpOpts.Specie(k).Posrg < 0 Then TmpOpts.Specie(k).Posrg = 1
  If TmpOpts.Specie(k).Posdn < 0 Then TmpOpts.Specie(k).Posdn = 1
  
  Initial_Position.Left = TmpOpts.Specie(k).Poslf * w
  Initial_Position.Top = TmpOpts.Specie(k).Postp * h
  
  Initial_Position.Width = (TmpOpts.Specie(k).Posrg * w) - Initial_Position.Left
  Initial_Position.Height = (TmpOpts.Specie(k).Posdn * h) - Initial_Position.Top
  Initial_Position.Visible = True
  
  Frame1.Refresh
  
End Sub

Private Sub AddSpec_Click()
  On Error GoTo fine
  CommonDialog1.FileName = ""
  CommonDialog1.Filter = "Dna file(*.txt)|*.txt" 'Botsareus 1/11/2013 DNA only
  CommonDialog1.InitDir = MDIForm1.MainDir + "\robots"
  CommonDialog1.DialogTitle = WSchoosedna
  CommonDialog1.ShowOpen
  If CommonDialog1.FileName <> "" Then 'Botsareus 1/11/2013 Do not insert robot if filename is blank
  additem CommonDialog1.FileName
 ' TmpOpts.SpeciesNum = TmpOpts.SpeciesNum + 1
  PopulateSpeciesList
  End If
fine:
End Sub

Private Sub DuplicaButt_Click() 'Botsareus 4/30/2013 Fix for the duplicator
  On Error GoTo fine
  Dim ind As Integer
  ind = SpecList.ListIndex
  If ind >= 0 Then
    TmpOpts.SpeciesNum = TmpOpts.SpeciesNum + 1
    TmpOpts.Specie(TmpOpts.SpeciesNum - 1) = TmpOpts.Specie(ind)
    DispSettings
  Else
    MsgBox ("Sorry, but you can only duplicate bots that originated in this simulation.")
  End If
  PopulateSpeciesList
fine:
End Sub
Sub DelSpec_Click()
  Dim ind As Integer, k As Long, t As Long, l As Long
  ind = SpecList.ListIndex
  If ind >= 0 Then
    SpecList.RemoveItem ind
    k = ind
    'l = SpecList.ListCount + 1
    For t = ind To SpecList.ListCount 'Listcount now has one fewer than it did before!!!
      TmpOpts.Specie(t) = TmpOpts.Specie(t + 1)
    Next t
    TmpOpts.SpeciesNum = TmpOpts.SpeciesNum - 1
    disprop
  'Else
   ' MsgBox ("Sorry, but you can only delete bots that originated in this simulation.")
    'Check to make sure our current slected index is still valid
    If ind >= SpecList.ListCount Then
      ind = ind - 1
    End If
    If SpecList.ListCount > 0 Then
      SpecList.ListIndex = ind
    End If
  End If
End Sub



Private Sub clearall()
  Dim k As Integer
  k = 0
'  SpecList.CLEAR
  While TmpOpts.Specie(k).Name <> ""
    TmpOpts.Specie(k).Name = ""
    k = k + 1
  Wend
  
  TmpOpts.Tides = 0
  TmpOpts.TidesOf = 0
End Sub

Sub additem(path As String)
  Dim k As Integer, t As Long
  SpecList.additem extractname(path)
  'k = SpecList.ListCount - 1
  
  k = TmpOpts.SpeciesNum
  TmpOpts.SpeciesNum = TmpOpts.SpeciesNum + 1
  
  TmpOpts.Specie(k).Colind = 8 'Botsareus 5/8/2012 Changed color CamboBox default value to Random
  TmpOpts.Specie(k).Posrg = 1
  TmpOpts.Specie(k).Posdn = 1
  TmpOpts.Specie(k).Poslf = 0
  TmpOpts.Specie(k).Postp = 0
  TmpOpts.Specie(k).Name = extractname(path)
  TmpOpts.Specie(k).path = extractpath(path)
  ExtractComment path, k
  TmpOpts.Specie(k).path = relpath(TmpOpts.Specie(k).path)
  TmpOpts.Specie(k).Veg = False
  TmpOpts.Specie(k).CantSee = False
  TmpOpts.Specie(k).DisableMovementSysvars = False
  TmpOpts.Specie(k).DisableDNA = False
  TmpOpts.Specie(k).CantReproduce = False
  TmpOpts.Specie(k).VirusImmune = False
    
  Randomize 'Botsareus 4/27/2013 Added randomize here so we have interesting colors
  TmpOpts.Specie(k).color = IIf(UseOldColor, RGB(Rnd * 200 + 55, Rnd * 200 + 55, Rnd * 255), RGB(Rnd * 255, Rnd * 255, Rnd * 255))
  
  'Special overwrites for unique robot names
  If TmpOpts.Specie(k).Name = "Base.txt" Then TmpOpts.Specie(k).color = vbBlue
  If TmpOpts.Specie(k).Name = "Mutate.txt" Then TmpOpts.Specie(k).color = vbRed
  If TmpOpts.Specie(k).Name = "robotA.txt" Then TmpOpts.Specie(k).color = RGB(255, 128, 0)
  If TmpOpts.Specie(k).Name = "robotB.txt" Then TmpOpts.Specie(k).color = RGB(0, 128, 255)
  If TmpOpts.Specie(k).Name = "Test.txt" Then TmpOpts.Specie(k).color = vbRed
  
  Cerchio.FillColor = TmpOpts.Specie(k).color 'Botsareus 4/27/2013 Update ze color on load speicies
  Cerchio.BorderColor = TmpOpts.Specie(k).color
  Line7.BorderColor = TmpOpts.Specie(k).color
  Line8.BorderColor = TmpOpts.Specie(k).color
  Line9.BorderColor = TmpOpts.Specie(k).color
  
  CurrSpec = k
  SetDefaultMutationRates TmpOpts.Specie(k).Mutables
  'Botsareus 12/11/2013 Do we have an .mrate file assoicated with this robot?
  Dim mfname As String: mfname = TmpOpts.Specie(k).path & "\" & extractexactname(TmpOpts.Specie(k).Name) & ".mrate"
  mfname = Replace(mfname, "&#", MDIForm1.MainDir) 'Botsareus 1/18/2014 Bugfix
  If dir(mfname) <> "" Then TmpOpts.Specie(k).Mutables = Load_mrates(mfname)
  
  TmpOpts.Specie(k).Mutables.Mutations = True
  TmpOpts.Specie(k).qty = 5
  TmpOpts.Specie(k).Stnrg = 3000
  
  AssignSkin k, path
  ShowSkin k
  

End Sub

Sub duplitem(b As Integer, a As Integer)
  Dim t As Integer
  
  TmpOpts.Specie(a).Posrg = TmpOpts.Specie(b).Posrg
  TmpOpts.Specie(a).Posdn = TmpOpts.Specie(b).Posdn
  TmpOpts.Specie(a).Poslf = TmpOpts.Specie(b).Poslf
  TmpOpts.Specie(a).Postp = TmpOpts.Specie(b).Postp
  TmpOpts.Specie(a).Veg = TmpOpts.Specie(b).Veg
  TmpOpts.Specie(a).Stnrg = TmpOpts.Specie(b).Stnrg
  TmpOpts.Specie(a).qty = TmpOpts.Specie(b).qty
  TmpOpts.Specie(a).Fixed = TmpOpts.Specie(b).Fixed
  TmpOpts.Specie(a).Colind = TmpOpts.Specie(b).Colind
  TmpOpts.Specie(a).CantSee = TmpOpts.Specie(b).CantSee
  TmpOpts.Specie(a).DisableMovementSysvars = TmpOpts.Specie(b).DisableMovementSysvars
  TmpOpts.Specie(a).DisableDNA = TmpOpts.Specie(b).DisableDNA
  TmpOpts.Specie(a).CantReproduce = TmpOpts.Specie(b).CantReproduce
  TmpOpts.Specie(a).VirusImmune = TmpOpts.Specie(b).VirusImmune
  TmpOpts.Specie(a).Mutables = TmpOpts.Specie(b).Mutables
End Sub

Private Sub ExtractComment(path As String, k As Integer)
  On Error GoTo fine
  Dim commend As Boolean
  Dim a As String
  path = stringops.respath(path)
  TmpOpts.Specie(k).Comment = ""
  Open path For Input As 1
    While Not EOF(1) And Not commend
      Line Input #1, a
      'Debug.Print a
      If Left(a, 1) = "'" Or Left(a, 1) = "/" Then
        TmpOpts.Specie(k).Comment = TmpOpts.Specie(k).Comment + Mid(a, 2) + vbCrLf
      Else
        commend = True
      End If
    Wend
  Close 1
  Exit Sub
fine:
  TmpOpts.Specie(k).Comment = "ATTENTION!  NOT A VALID BOT FILE."
End Sub

Private Sub SkinChange_Click()
  If SpecList.ListIndex >= 0 Then
    TmpOpts.Specie(SpecList.ListIndex).Skin(6) = (TmpOpts.Specie(SpecList.ListIndex).Skin(6) + Random(0, half) * 2) / 3
    ShowSkin SpecList.ListIndex
  End If
End Sub

Private Sub ShowSkin(k As Integer)
  Dim t As Integer
  Dim x As Long
  Dim y As Long
  x = Cerchio.Left + Cerchio.Width / 2
  y = Cerchio.Top + Cerchio.Height / 2
  multx = Cerchio.Width / 120
  multy = Cerchio.Height / 120
  Me.AutoRedraw = True
  Line7.x1 = TmpOpts.Specie(k).Skin(0) * multx * Cos(TmpOpts.Specie(k).Skin(1) / 100) + x
  Line7.y1 = TmpOpts.Specie(k).Skin(0) * multy * Sin(TmpOpts.Specie(k).Skin(1) / 100) + y
  Line7.x2 = TmpOpts.Specie(k).Skin(2) * multx * Cos(TmpOpts.Specie(k).Skin(3) / 100) + x
  Line7.y2 = TmpOpts.Specie(k).Skin(2) * multy * Sin(TmpOpts.Specie(k).Skin(3) / 100) + y
  Line8.x1 = Line7.x2
  Line8.y1 = Line7.y2
  Line8.x2 = TmpOpts.Specie(k).Skin(4) * multx * Cos(TmpOpts.Specie(k).Skin(5) / 100) + x
  Line8.y2 = TmpOpts.Specie(k).Skin(4) * multy * Sin(TmpOpts.Specie(k).Skin(5) / 100) + y
  Line9.x1 = Line8.x2
  Line9.y1 = Line8.y2
  Line9.x2 = TmpOpts.Specie(k).Skin(6) * multx * Cos(TmpOpts.Specie(k).Skin(7) / 100) + x
  Line9.y2 = TmpOpts.Specie(k).Skin(6) * multy * Sin(TmpOpts.Specie(k).Skin(7) / 100) + y
End Sub

Sub AssignSkin(k As Integer, path As String) 'The new skin engine requires path
'Botsareus 4/27/2013 The new skin engine


Randomize 0

Dim robname As String
robname = Replace(TmpOpts.Specie(k).Name, ".txt", "")

Dim newR As Double
Dim nextR As Double
Dim nameR As Double
Dim x As Long

Dim dbls() As Double

ReDim dbls(Len(robname) - 1)
For x = 1 To Len(robname)
dbls(x - 1) = Rnd(-Asc(Mid(robname, x, 1)))
Next 'pre seeds

For x = 1 To Len(robname)
newR = dbls(x - 1)
nextR = Rnd(-(angle(0, 0, nextR - 0.5, newR - 0.5)))
Next 'randomize by name

nameR = nextR
newR = 0
nextR = 0

  If MaxRobs = 0 Then ReDim rob(0)
  If LoadDNA(path, 0) Then

    Randomize 0
    
    ReDim dbls(UBound(rob(0).dna))
    For x = 0 To UBound(rob(0).dna)
    dbls(x) = Rnd(-(angle(0, 0, Rnd(-rob(0).dna(x).value) - 0.5, Rnd(-rob(0).dna(x).type) - 0.5)))
    Next 'pre seeds
    
    For x = 0 To UBound(rob(0).dna)
    newR = dbls(x)
    nextR = Rnd(-(angle(0, 0, nextR - 0.5, newR - 0.5)))
    Next 'randomize by dna
    
    If MaxRobs = 0 Then ReDim rob(0)
  
  End If

Randomize nextR * 1000

  Dim i As Integer
  If k > -1 Then
    For i = 0 To 7 Step 2
      TmpOpts.Specie(k).Skin(i) = Int(Rnd * (half + 1))
      If i = 4 Then Randomize nameR * 1000
      TmpOpts.Specie(k).Skin(i + 1) = Int(Rnd * 629)
    Next i
    Randomize
    TmpOpts.Specie(k).Skin(6) = (TmpOpts.Specie(k).Skin(6) + Int(Rnd * (half + 1)) * 2) / 3
  End If
End Sub

Private Sub MutRatesBut_Click()
  'EricL 4/9/2006 Catches crash when no species is selected
  If optionsform.CurrSpec < 0 Then
   MsgBox ("Please select a species.")
  Else
    MutationsProbability.Show vbModal
  End If
End Sub

Public Sub SwapSpecies(a As Integer, b As Integer)
Dim c As datispecie

  c = TmpOpts.Specie(a)
  TmpOpts.Specie(a) = TmpOpts.Specie(b)
  TmpOpts.Specie(b) = c

End Sub

Public Sub SortSpecies()
Dim i As Integer
Dim j As Integer

  For i = 0 To TmpOpts.SpeciesNum - 2
    For j = i + 1 To TmpOpts.SpeciesNum - 1
      If UCase(TmpOpts.Specie(i).Name) > UCase(TmpOpts.Specie(j).Name) Then
        SwapSpecies i, j
      End If
    Next j
  Next i

j = TmpOpts.SpeciesNum - 1

End Sub

Public Sub datatolist() 'datatolist
  Dim i As Integer
  SpecList.CLEAR
  
  SortSpecies
  
  For i = 0 To TmpOpts.SpeciesNum - 1
      SpecList.additem (TmpOpts.Specie(i).Name)
      ExtractComment TmpOpts.Specie(i).path + "\" + TmpOpts.Specie(i).Name, i
  Next i
   
   
End Sub

Private Sub enprop()
  Frame1.Enabled = True
End Sub

Private Sub disprop()
  Frame1.Enabled = False
End Sub

Private Sub IndNum_Click(Index As Integer)
  Dim qty As Integer
  Select Case Index
    Case 0
      qty = 5
    Case 1
      qty = 10
    Case 2
      qty = 15
    Case 3
      qty = 30
    Case 4
      qty = 0
  End Select
  SpecQty.Text = qty
End Sub

Private Sub SpecQty_Change()
  If CurrSpec > -1 Then
    TmpOpts.Specie(CurrSpec).qty = val(SpecQty.Text)
  End If
End Sub

Private Sub InNrg_Click(Index As Integer)
  Dim qty As Integer
  Select Case Index
    Case 0
      qty = 3000
    Case 1
      qty = 4000
    Case 2
      qty = 5000
    Case 3
      qty = 30000
  End Select
  SpecNrg.Text = qty
End Sub

Private Sub SpecNrg_Change()
  If CurrSpec >= 0 Then TmpOpts.Specie(CurrSpec).Stnrg = val(SpecNrg.Text) Mod 32000
End Sub

Private Sub SpecVeg_Click()
  If CurrSpec > -1 Then
    TmpOpts.Specie(CurrSpec).Veg = False
    If SpecVeg.value = 1 Then
        TmpOpts.Specie(CurrSpec).Veg = True
        TmpOpts.Specie(CurrSpec).NoChlr = False
        chkNoChlr.value = 0
    End If
  End If
End Sub

Private Sub BlockSpec_Click()
  If CurrSpec > -1 Then
    TmpOpts.Specie(CurrSpec).Fixed = False
    If BlockSpec.value = 1 Then TmpOpts.Specie(CurrSpec).Fixed = True
  End If
End Sub

Private Sub SpecCol_click()
If CurrSpec = -1 Then Exit Sub 'Botsareus 2/3/2013 bug fix when no robot selected

  Dim c As String
  Dim r As Single
  Dim g As Single
  Dim b As Single
  Dim k As Integer
  Dim col As Long
  
  c = SpecCol.Text
  k = CurrSpec
  If speclistchecked Then GoTo bypass
  Select Case c
    Case "Red"
      r = 200
      g = 60
      b = 60
    Case "Blue"
      r = 60
      g = 60
      b = 200
    Case "Green"
      r = 60
      g = 200
      b = 60
    Case "Yellow"
      r = 220
      g = 220
      b = 60
    Case "Pink"
      r = 210
      g = 160
      b = 210
    Case "Brown"
      r = 140
      g = 110
      b = 70
    Case "Purple"
      r = 150
      g = 56
      b = 180
    Case "Orange"
      r = 220
      g = 150 ' EricL
      b = 60
    Case "Cyan"
      r = 58
      g = 207
      b = 228
    Case "Random" 'Botsareus 4/27/2013 This was very broken
      Randomize
      r = IIf(UseOldColor, Rnd * 200 + 55, Rnd * 255)
      g = IIf(UseOldColor, Rnd * 200 + 55, Rnd * 255)
      b = Rnd * 255
    Case "Custom"
      col = TmpOpts.Specie(k).color
      MakeColor col, respath(TmpOpts.Specie(k).path) + "\" + TmpOpts.Specie(k).Name
      If ColorForm.SelectColor Then
        TmpOpts.Specie(k).color = ColorForm.color
      Else
        TmpOpts.Specie(k).color = ColorForm.OldColor
      End If
      TmpOpts.Specie(k).Colind = SpecCol.ListIndex
      GoTo bypass
  End Select
  TmpOpts.Specie(k).color = (65536 * b) + (256 * g) + r
  TmpOpts.Specie(k).Colind = SpecCol.ListIndex
bypass:
  Cerchio.FillColor = TmpOpts.Specie(k).color
  Cerchio.BorderColor = TmpOpts.Specie(k).color
  Line7.BorderColor = TmpOpts.Specie(k).color
  Line8.BorderColor = TmpOpts.Specie(k).color
  Line9.BorderColor = TmpOpts.Specie(k).color
End Sub
Private Sub MakeColor(col As Long, path As String)
  ColorForm.color = col
  ColorForm.SelectColor = False
  ColorForm.path = path
  ColorForm.Show vbModal
End Sub

''''''''''''''''''''''''''''''''''''
' Position control '''''''''''''''''
''''''''''''''''''''''''''''''''''''
'=========================== Sample controls ===========================
'To drag a control, simply call the DragBegin function with
'the control to be dragged
'=======================================================================

Private Sub PosReset_Click()
  DragEnd
  
  Initial_Position.Left = 0
  Initial_Position.Top = 0
  
  Initial_Position.Width = IPB.Width
  Initial_Position.Height = IPB.Height
  
  'EricL Hitting the reset button is sufficient to set the bots starting position
  'without having to click on the Initial Position control
  'These are percentages of the field width and height
  If CurrSpec > -1 Then
    TmpOpts.Specie(CurrSpec).Posrg = 1
    TmpOpts.Specie(CurrSpec).Posdn = 1
    TmpOpts.Specie(CurrSpec).Poslf = 0
    TmpOpts.Specie(CurrSpec).Postp = 0
  End If

End Sub

Private Sub Initial_Position_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
  If button = vbLeftButton Then
    DragBegin Initial_Position
  End If
End Sub

Private Sub IPB_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
  If button = vbLeftButton Then
    DragBegin Initial_Position
  End If
End Sub

'========================== Dragging Code ================================

'Initialization -- Do not call more than once
Public Sub DragInit()
    Dim i As Integer, xHandle As Single, yHandle As Single

    'Use black Picture box controls for 8 sizing handles
    'Calculate size of each handle
    xHandle = 8 * Screen.TwipsPerPixelX
    yHandle = 8 * Screen.TwipsPerPixelY
    'Load array of handles until we have 8
    For i = 0 To 7
        If i <> 0 Then
            Load picHandle(i)
        End If
        picHandle(i).Width = xHandle
        picHandle(i).Height = yHandle
        'Must be in front of other controls
        picHandle(i).ZOrder
    Next i
    
    For i = 0 To 3
      If i <> 0 Then
          Load RobPlacLine(i)
          RobPlacLine(i) = RobPlacLine(0)
      End If
      'Must be in front of other controls
      RobPlacLine(i).ZOrder
    Next i
    'Set mousepointers for each sizing handle
    picHandle(0).MousePointer = vbSizeNWSE
    picHandle(2).MousePointer = vbSizeNESW
    picHandle(4).MousePointer = vbSizeNWSE
    picHandle(6).MousePointer = vbSizeNESW
    
    RobPlacLine(0).MousePointer = vbSizeNS
    RobPlacLine(1).MousePointer = vbSizeNS
    
    RobPlacLine(2).MousePointer = vbSizeWE
    RobPlacLine(3).MousePointer = vbSizeWE
    
    'Initialize current control
    Set m_CurrCtl = Nothing
End Sub

'Drags the specified control
Public Sub DragBegin(ctl As Control)
    Dim rc As RECT

    'Hide any visible handles
    ShowHandles False
    
    'Save reference to control being dragged
    Set m_CurrCtl = ctl
    
    'Store initial mouse position
    GetCursorPos m_DragPoint
    
    'Save control position (in screen coordinates)
    'Note: control might not have a window handle
    m_DragRect.SetRectToCtrl m_CurrCtl, IPB, Frame1, SSTab1
    m_DragRect.TwipsToScreen m_CurrCtl
    
    'Make initial mouse position relative to control
    m_DragPoint.x = m_DragPoint.x - m_DragRect.Left
    m_DragPoint.y = m_DragPoint.y - m_DragRect.Top
    
    'Force redraw of form without sizing handles
    'before drawing dragging rectangle
    Refresh
    
    'Show dragging rectangle
    DrawDragRect
    
    'Indicate dragging under way
    m_DragState = StateDragging
    
    'In order to detect mouse movement over any part of the form,
    'we set the mouse capture to the form and will process mouse
    'movement from the applicable form events
    ReleaseCapture  'This appears needed before calling SetCapture
    SetCapture hwnd
    
    'Limit cursor movement within form
    GetWindowRect hwnd, rc
    ClipCursor rc
End Sub

'Clears any current drag mode and hides sizing handles
Public Sub DragEnd()
    Set m_CurrCtl = Nothing
    ShowHandles False
    m_DragState = StateNothing
End Sub

'To handle all mouse message anywhere on the form, we set the mouse
'capture to the form. Mouse movement is processed here
Private Sub Form_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
If CurrSpec = -1 Then Exit Sub 'Botsareus 2/3/2013 bug fix when no robot selected

    Dim nWidth As Single, nHeight As Single
    Dim pt As POINTAPI

    If m_DragState = StateDragging Then
        'Save dimensions before modifying rectangle
        nWidth = m_DragRect.Right - m_DragRect.Left
        nHeight = m_DragRect.Bottom - m_DragRect.Top
        'Get current mouse position in screen coordinates
        GetCursorPos pt
        'Hide existing rectangle
        DrawDragRect
        'Update drag rectangle coordinates
        m_DragRect.Left = pt.x - m_DragPoint.x
        m_DragRect.Top = pt.y - m_DragPoint.y
        m_DragRect.Right = m_DragRect.Left + nWidth
        m_DragRect.Bottom = m_DragRect.Top + nHeight
        'Draw new rectangle
        DrawDragRect
    ElseIf m_DragState = StateSizing Then
        'Get current mouse position in screen coordinates
        GetCursorPos pt
        'Hide existing rectangle
        DrawDragRect
        'Action depends on handle being dragged
        Select Case m_DragHandle
            Case 0
                m_DragRect.Left = pt.x
                m_DragRect.Top = pt.y
            Case 2
                m_DragRect.Right = pt.x
                m_DragRect.Top = pt.y
            Case 4
                m_DragRect.Right = pt.x
                m_DragRect.Bottom = pt.y
            Case 6
                m_DragRect.Left = pt.x
                m_DragRect.Bottom = pt.y
            Case 9
                m_DragRect.Top = pt.y
            Case 10
                m_DragRect.Bottom = pt.y
            Case 11
                m_DragRect.Left = pt.x
            Case 12
                m_DragRect.Right = pt.x
        End Select
        'Draw new rectangle
        DrawDragRect
    End If
End Sub

'To handle all mouse message anywhere on the form, we set the mouse
'capture to the form. Mouse up is processed here
Private Sub Form_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
If CurrSpec = -1 Then Exit Sub 'Botsareus 2/3/2013 bug fix when no robot selected

    If button = vbLeftButton Then
        If m_DragState = StateDragging Or m_DragState = StateSizing Then
            'Hide drag rectangle
            DrawDragRect
            'Move control to new location
            m_DragRect.ScreenToTwips m_CurrCtl
            m_DragRect.SetCtrlToRect m_CurrCtl, IPB, Frame1, SSTab1
            'Restore sizing handles
            ShowHandles True
            'Free mouse movement
            ClipCursor ByVal 0&
            'Release mouse capture
            ReleaseCapture
            'Reset drag state
            m_DragState = StateNothing
            
            Dim w As Long, h As Long
    
            w = IPB.Width
            h = IPB.Height
    
            TmpOpts.Specie(CurrSpec).Posrg = (Initial_Position.Left + Initial_Position.Width) / w
            TmpOpts.Specie(CurrSpec).Posdn = (Initial_Position.Top + Initial_Position.Height) / h
            TmpOpts.Specie(CurrSpec).Poslf = (Initial_Position.Left) / w
            TmpOpts.Specie(CurrSpec).Postp = (Initial_Position.Top) / h
            
            If (TmpOpts.Specie(CurrSpec).Posrg > 1) Then TmpOpts.Specie(CurrSpec).Posrg = 1
            If (TmpOpts.Specie(CurrSpec).Posrg < 0) Then TmpOpts.Specie(CurrSpec).Posrg = 0
            
            If (TmpOpts.Specie(CurrSpec).Posdn > 1) Then TmpOpts.Specie(CurrSpec).Posdn = 1
            If (TmpOpts.Specie(CurrSpec).Posdn < 0) Then TmpOpts.Specie(CurrSpec).Posdn = 0
            
            If (TmpOpts.Specie(CurrSpec).Poslf > 1) Then TmpOpts.Specie(CurrSpec).Poslf = 1
            If (TmpOpts.Specie(CurrSpec).Poslf < 0) Then TmpOpts.Specie(CurrSpec).Poslf = 0
            
            If (TmpOpts.Specie(CurrSpec).Postp > 1) Then TmpOpts.Specie(CurrSpec).Postp = 1
            If (TmpOpts.Specie(CurrSpec).Postp < 0) Then TmpOpts.Specie(CurrSpec).Postp = 0
        End If
    End If
End Sub

'Process MouseDown over handles
Private Sub picHandle_MouseDown(Index As Integer, button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    Dim rc As RECT

    'Handles should only be visible when a control is selected
    If picHandle(Index).Visible = False Then Exit Sub
    
    'NOTE: m_DragPoint not used for sizing
    'Save control position in screen coordinates
    m_DragRect.SetRectToCtrl m_CurrCtl, IPB, Frame1, SSTab1
    m_DragRect.TwipsToScreen m_CurrCtl
    'Track index handle
    m_DragHandle = Index
    'Hide sizing handles
    ShowHandles False
    'We need to force handles to hide themselves before drawing drag rectangle
    Refresh
    'Indicate sizing is under way
    m_DragState = StateSizing
    'Show sizing rectangle
    DrawDragRect
    'In order to detect mouse movement over any part of the form,
    'we set the mouse capture to the form and will process mouse
    'movement from the applicable form events
    SetCapture hwnd
    'Limit cursor movement within form
    GetWindowRect hwnd, rc
    ClipCursor rc
End Sub

Private Sub Robplacline_MouseDown(Index As Integer, button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    Dim rc As RECT

    'Handles should only be visible when a control is selected
    If RobPlacLine(Index).Visible = False Then Exit Sub
    
    'NOTE: m_DragPoint not used for sizing
    'Save control position in screen coordinates
    m_DragRect.SetRectToCtrl m_CurrCtl, IPB, Frame1, SSTab1
    m_DragRect.TwipsToScreen m_CurrCtl
    'Track index handle
    m_DragHandle = Index + 9
    'Hide sizing handles
    ShowHandles False
    'We need to force handles to hide themselves before drawing drag rectangle
    Refresh
    'Indicate sizing is under way
    m_DragState = StateSizing
    'Show sizing rectangle
    DrawDragRect
    'In order to detect mouse movement over any part of the form,
    'we set the mouse capture to the form and will process mouse
    'movement from the applicable form events
    SetCapture hwnd
    'Limit cursor movement within form
    GetWindowRect hwnd, rc
    ClipCursor rc
End Sub

'Display or hide the sizing handles and arrange them for the current rectangld
Private Sub ShowHandles(Optional bShowHandles As Boolean = True)
    Dim i As Integer
    Dim xFudge As Long, yFudge As Long
    
    Dim ShowHandlesBool As Boolean
    
    ShowHandlesBool = True
    
    If bShowHandles And Not m_CurrCtl Is Nothing Then
        With m_DragRect
            If Not (.Width < 250 Or .Height < 250) Then
            
            'Save some calculations in variables for speed
            xFudge = (2 * Screen.TwipsPerPixelX)
            yFudge = (2 * Screen.TwipsPerPixelY)
            
            'Top Left
            picHandle(0).Move (.Left - IPB.Left - Frame1.Left - SSTab1.Left) + xFudge, _
              (.Top - IPB.Top - Frame1.Top) + yFudge - SSTab1.Top
            'Top right
            picHandle(2).Move .Left - IPB.Left - Frame1.Left - SSTab1.Left + .Width - picHandle(0).Width - xFudge, _
              .Top - IPB.Top - Frame1.Top + yFudge - SSTab1.Top
            'Bottom left
            picHandle(6).Move .Left - IPB.Left - Frame1.Left - SSTab1.Left + xFudge, _
              .Top + .Height - picHandle(0).Height - yFudge - IPB.Top - Frame1.Top - SSTab1.Top
            'Bottom right
            picHandle(4).Move (.Left - IPB.Left - Frame1.Left - SSTab1.Left + .Width) - picHandle(0).Width - xFudge, _
              .Top + .Height - picHandle(0).Height - yFudge - IPB.Top - Frame1.Top - SSTab1.Top
            Else
              ShowHandlesBool = False
            End If
              
            RobPlacLine(0).Move .Left - IPB.Left - Frame1.Left - SSTab1.Left, .Top - IPB.Top - Frame1.Top - SSTab1.Top, .Width, 60
            RobPlacLine(1).Move .Left - IPB.Left - Frame1.Left - SSTab1.Left, .Top + .Height - IPB.Top - Frame1.Top - SSTab1.Top - 60, .Width, 60
            
            RobPlacLine(2).Move .Left - IPB.Left - Frame1.Left - SSTab1.Left, .Top - IPB.Top - Frame1.Top - SSTab1.Top, 60, .Height
            RobPlacLine(3).Move .Left + .Width - IPB.Left - Frame1.Left - SSTab1.Left - 60, .Top - IPB.Top - Frame1.Top - SSTab1.Top, 60, .Height
        End With
    End If
    'Show or hide each handle
    picHandle(0).Visible = bShowHandles And ShowHandlesBool
    picHandle(2).Visible = bShowHandles And ShowHandlesBool
    picHandle(6).Visible = bShowHandles And ShowHandlesBool
    picHandle(4).Visible = bShowHandles And ShowHandlesBool
    
    RobPlacLine(0).Visible = bShowHandles
    RobPlacLine(1).Visible = bShowHandles
    RobPlacLine(2).Visible = bShowHandles
    RobPlacLine(3).Visible = bShowHandles
End Sub
'Draw drag rectangle. The API is used for efficiency and also
'because drag rectangle must be drawn on the screen DC in
'order to appear on top of all controls
Private Sub DrawDragRect()
    Dim hPen As Long, hOldPen As Long
    Dim hBrush As Long, hOldBrush As Long
    Dim hScreenDC As Long, nDrawMode As Long

    'Get DC of entire screen in order to
    'draw on top of all controls
    hScreenDC = GetDC(0)
    'Select GDI object
    hPen = CreatePen(PS_SOLID, 2, 0)
    hOldPen = SelectObject(hScreenDC, hPen)
    hBrush = GetStockObject(NULL_BRUSH)
    hOldBrush = SelectObject(hScreenDC, hBrush)
    nDrawMode = SetROP2(hScreenDC, R2_NOT)
    'Draw rectangle
    Rectangle hScreenDC, m_DragRect.Left, m_DragRect.Top, _
        m_DragRect.Right, m_DragRect.Bottom
    'Restore DC
    SetROP2 hScreenDC, nDrawMode
    SelectObject hScreenDC, hOldBrush
    SelectObject hScreenDC, hOldPen
    ReleaseDC 0, hScreenDC
    'Delete GDI objects
    DeleteObject hPen
End Sub


''''''''''''''''''''''''''''''''''''''''''''

'''General Panel'''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''

'''''Start dimensions control

Private Sub FieldSizeSlide_Scroll()
  Dim t As Integer
  Dim oldsw As Single, oldsh As Single
  
  TmpOpts.FieldSize = FieldSizeSlide.value
  
  oldsw = TmpOpts.FieldWidth
  oldsh = TmpOpts.FieldHeight
  
  If TmpOpts.FieldSize = 1 Then 'F1 mode
    TmpOpts.FieldWidth = 9237
    TmpOpts.FieldHeight = 6928
  Else
    TmpOpts.FieldWidth = 8000
    TmpOpts.FieldHeight = 6000
    
    If TmpOpts.FieldSize <= 12 Then
      TmpOpts.FieldWidth = TmpOpts.FieldWidth * TmpOpts.FieldSize
      TmpOpts.FieldHeight = TmpOpts.FieldHeight * TmpOpts.FieldSize
    Else ' Field sizes larger than size 12 get big fast...
      TmpOpts.FieldWidth = (TmpOpts.FieldWidth * 12) * (TmpOpts.FieldSize - 12) * 2
      TmpOpts.FieldHeight = (TmpOpts.FieldHeight * 12) * (TmpOpts.FieldSize - 12) * 2
    End If
  End If
  
  FWidthLab.Caption = TmpOpts.FieldWidth
  FHeightLab.Caption = TmpOpts.FieldHeight
End Sub


'''''End Dimensions Control

'''''Start Seeded Random Control

Private Sub UserSeedText_Change()
  TmpOpts.UserSeedNumber = val(UserSeedText.Text)
  SimOpts.UserSeedNumber = val(UserSeedText.Text)
End Sub

''''''End Seeded Random Control

''''''Start Torroidal, etc.

Private Sub TopDownCheck_Click()
  TmpOpts.Updnconnected = TopDownCheck.value * True
  TmpOpts.Toroidal = TopDownCheck.value * True And RightLeftCheck.value * True
End Sub

Private Sub RightLeftCheck_Click()
  TmpOpts.Dxsxconnected = RightLeftCheck.value * True
  TmpOpts.Toroidal = TopDownCheck.value * True And RightLeftCheck.value * True
End Sub

'''''End Torroidal, etc.

''''''Waste Control
Private Sub CustomWaste_Change()
  If val(CustomWaste.Text) = 0 Then
    TmpOpts.BadWastelevel = -1
  Else
    TmpOpts.BadWastelevel = val(CustomWaste.Text)
  End If
End Sub

Private Sub VirusImmuneCheck_Click()
 If CurrSpec > -1 Then
    TmpOpts.Specie(CurrSpec).VirusImmune = False
    If VirusImmuneCheck.value = 1 Then TmpOpts.Specie(CurrSpec).VirusImmune = True
  End If
End Sub

'''''''End Waste Control

'''''''Start Corpse Controls

Private Sub CorpseCheck_Click()
  TmpOpts.CorpseEnabled = CorpseCheck.value * True
  DecayText.Enabled = CorpseCheck.value * True
  DecayUpDn(0).Enabled = CorpseCheck.value * True
  DecayUpDn(1).Enabled = CorpseCheck.value * True
  FrequencyText.Enabled = CorpseCheck.value * True
  DecayOption(0).Enabled = CorpseCheck.value * True
  DecayOption(1).Enabled = CorpseCheck.value * True
  DecayOption(2).Enabled = CorpseCheck.value * True
End Sub

Private Sub DecayOption_Click(Index As Integer)
  TmpOpts.DecayType = Index + 1
End Sub

Private Sub DecayText_Change()
  TmpOpts.Decay = val(DecayText.Text)
End Sub

Private Sub FrequencyText_Change()
  TmpOpts.Decaydelay = val(FrequencyText.Text)
End Sub

Private Sub DecayUpDn_Change(Index As Integer)
  'Buddied to DecayText amd Frequency Text, so do nothing.
End Sub

'''''End Corpse Controls

'''''Start Veg Controls

'''StartPond Mode
Private Sub Pondcheck_Click()

TmpOpts.Pondmode = Pondcheck.value * True
LightText.Enabled = Pondcheck.value * True
LightUpDn.Enabled = Pondcheck.value * True
Gradient.Enabled = Pondcheck.value * True
GradientUpDn.Enabled = Pondcheck.value * True
GradientLabel.Enabled = Pondcheck.value * True
Frame30.Visible = Not Pondcheck.value * True 'Botsareus 8/23/2014 A little mod here to make UI less confusing

If validate Then Exit Sub

If Pondcheck.value = 1 Then
validate = True

 If MsgBox("Turning on Pond Mode will greatly alter the physics of your simulation. Most importantly, in Pond Mode the light providing energy to bots with chloroplasts comes from the top and gradually disappears toward the bottom. Also, Pond Mode includes Earth-like gravity and top/down sides of the screen will not wrap-around. Are you sure?", vbExclamation + vbYesNo, "Darwinbots Settings") = vbNo Then
    Pondcheck.value = 0
    GoTo fin
 End If


TmpOpts.Updnconnected = False
TmpOpts.Ygravity = 6.2
TmpOpts.LightIntensity = 100
TmpOpts.Gradient = 2
DispSettings

fin:

validate = False

End If


End Sub

Private Sub LightText_Change()
  'Botsareus 5/18/2013 Overflow prevention
  If val(LightText.Text) > 1000 Then LightText.Text = 1000

  TmpOpts.LightIntensity = val(LightText.Text)
  color_lightlines
End Sub


Private Sub Gradient_Change()
  'Botsareus 5/18/2013 Overflow prevention
  If val(Gradient.Text) > 200 Then Gradient.Text = 200
  
  TmpOpts.Gradient = val(Gradient.Text) / 10 + 1
  color_lightlines
End Sub

Private Sub GradientUpDn_Change()
  Dim a As Single
  a = GradientUpDn.value
  Gradient.Text = (a / 5)
End Sub

Private Sub EnergyScalingFactor_Change()
  'this just controls the little graph to the left of this control,
  'doesn't effect the simulation
  If val(EnergyScalingFactor.Text) = 0 Then EnergyScalingFactor.Text = "1"
  Gradient_Change 'Botsareus 12/12/2012 Update the graph by updating the gradient
End Sub

Private Sub color_lightlines()
  Dim Index As Integer
  Dim fxn As Single
  Dim depth As Single, depth2 As Single
  Dim greyform As Single
  Dim toplevel As Single
  Dim Gradient As Single
  Dim color As Long
  
  Gradient = TmpOpts.Gradient
  
  depth = 1.1
  If EnergyScalingFactor = 0 Then EnergyScalingFactor = 1
  toplevel = 2 * (val(EnergyScalingFactor.Text) / (depth ^ Gradient))
   
  depth = TmpOpts.FieldHeight / 16 / 2000 + 1
  fxn = 2 * (TmpOpts.LightIntensity / (depth ^ Gradient)) / toplevel
  greyform = Convert_PercentageColor(Ceil(fxn, 1))
  greyform = Abs(greyform)
  color = (65536 * Ceil(greyform, 255)) + (256 * Ceil(greyform, 255))
  LightStrata(Index).BorderColor = color
  
  depth = 1
  fxn = 2 * (TmpOpts.LightIntensity / (depth ^ Gradient)) / toplevel
  greyform = Convert_PercentageColor(fxn)
  greyform = Abs(greyform)
  color = (65536 * Ceil(greyform, 255)) + (256 * Ceil(greyform, 255))
  LightStrata(15).BorderColor = color
  
  For Index = 1 To 14
    depth = (Index * TmpOpts.FieldHeight / 16) / 2000 + 1
    depth2 = ((Index + 1) * TmpOpts.FieldHeight / 16 / 2000) + 1
    fxn = 2 * (TmpOpts.LightIntensity / (depth ^ Gradient)) / toplevel
    fxn = fxn + 2 * (TmpOpts.LightIntensity / (depth2 ^ Gradient)) / toplevel
    fxn = Ceil(fxn / 2, 1)
    greyform = Convert_PercentageColor(fxn)
    greyform = Abs(greyform)
    LightStrata(Index).BorderColor = (65536 * greyform) + (256 * greyform)
  Next Index
  Frame17.Visible = False
  Frame17.Visible = True
End Sub

Private Function Convert_PercentageColor(sgl As Single) As Integer
  Dim temp As Single
  
  temp = 255 * sgl
  If temp > 255 Then temp = 255
  
  Convert_PercentageColor = temp
End Function

Private Function Ceil(ByVal a As Single, ByVal b As Single) As Single
  If a > b Then a = b
  Ceil = a
End Function

'''''''''''''''''''''''''''''''''''''''

'''End Pond Mode

'''Start Rest of Veg controls

'''''''''''''''''''''''''''''''''''''''

Private Sub MaxPopText_Change()
  TmpOpts.MaxPopulation = val(MaxPopText.Text) Mod 32000
End Sub

Private Sub MaxPopUpDn_Change()
  'Buddied to MaxPopText, and synchronized, so we
  'do nothing
End Sub

Private Sub MinVegText_Change()
  TmpOpts.MinVegs = val(MinVegText.Text) Mod 32000
End Sub

Private Sub MinVegUpDn_Change()
  'Buddied to MaxPopText, and synchronized, so we
  'do nothing
End Sub

Private Sub RepopAmountText_Change()
  TmpOpts.RepopAmount = val(RepopAmountText.Text) Mod 32000
End Sub

Private Sub RepopAmountUpDown_Change(Index As Integer)
  'Buddied to RepopAmountText, and synchronized, so we
  'do nothing
End Sub

Private Sub RepopCooldownText_Change()
  TmpOpts.RepopCooldown = val(RepopCooldownText.Text) Mod 32000
End Sub

Private Sub RepopCooldownUpDown_Change(Index As Integer)
  'Buddied to RepopCooldownText, and synchronized, so we
  'do nothing
End Sub

Private Sub KillDistVegsCheck_Click()
  TmpOpts.KillDistVegs = KillDistVegsCheck.value * True
End Sub

Private Sub MaxNRGText_Change()
  TmpOpts.MaxEnergy = val(MaxNRGText.Text) Mod 32000
End Sub

Private Sub BodyNrgDist_change()
  TmpOpts.VegFeedingToBody = BodyNrgDist.value / 100
End Sub

''''''''''''''''''''''''''''''''''

'''''End Veg Controls and END GENERAL

''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''

''''Physics Panel'''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''

'needs work
Private Sub DragCombo_Click()
  Select Case DragCombo.Text
    Case DragCombo.List(0)
      TmpOpts.Viscosity = 0.01
      TmpOpts.Density = 0.0000001
    Case DragCombo.List(1)
      TmpOpts.Viscosity = 0.0005
      TmpOpts.Density = 0.0000001
    Case DragCombo.List(2)
      TmpOpts.Viscosity = 0.000025
      TmpOpts.Density = 0.0000001
    Case DragCombo.List(3)
      TmpOpts.Viscosity = 0#
      TmpOpts.Density = 0#
  End Select
End Sub

Private Sub EfficiencyCombo_Click()
  Select Case EfficiencyCombo.Text
    Case EfficiencyCombo.List(0)
      TmpOpts.PhysMoving = 1
    Case EfficiencyCombo.List(1)
      TmpOpts.PhysMoving = 0.66
    Case EfficiencyCombo.List(2)
      TmpOpts.PhysMoving = 0.33
  End Select
End Sub

Private Sub FrictionCombo_Click()
  Select Case FrictionCombo.Text
   Case FrictionCombo.List(0)
    TmpOpts.Zgravity = 4
    TmpOpts.CoefficientStatic = 0.9
    TmpOpts.CoefficientKinetic = 0.75
   Case FrictionCombo.List(1)
    TmpOpts.Zgravity = 2
    TmpOpts.CoefficientStatic = 0.6
    TmpOpts.CoefficientKinetic = 0.4
   Case FrictionCombo.List(2)
    TmpOpts.Zgravity = 1
    TmpOpts.CoefficientStatic = 0.05
    TmpOpts.CoefficientKinetic = 0.05
   Case FrictionCombo.List(3)
    TmpOpts.Zgravity = 0
    TmpOpts.CoefficientStatic = 0#
    TmpOpts.CoefficientKinetic = 0#
  End Select
End Sub

Private Sub GravityCombo_Click()
  Select Case GravityCombo.Text
    Case GravityCombo.List(0)
      TmpOpts.Ygravity = 0
    Case GravityCombo.List(1)
      TmpOpts.Ygravity = 0.1
    Case GravityCombo.List(2)
      TmpOpts.Ygravity = 0.3
    Case GravityCombo.List(3)
      TmpOpts.Ygravity = 0.9
    Case GravityCombo.List(4)
      TmpOpts.Ygravity = 6
  End Select
End Sub

Private Sub BrownianCombo_Click()
  Select Case BrownianCombo.Text
    Case BrownianCombo.List(0)
      TmpOpts.PhysBrown = 7
    Case BrownianCombo.List(1)
      TmpOpts.PhysBrown = 0.5
    Case BrownianCombo.List(2)
      TmpOpts.PhysBrown = 0
  End Select
End Sub

'''''''''''''''''''''''''''''''''''''''''''

Private Sub ToPhysics_Click()
  PhysicsOptions.Show vbModal
  Update
End Sub

'
'  All settings
'

Private Sub Form_Activate()

'Botsareus 12/12/2012 Hide always on top forms for easy readability
datirob.Visible = False
ActivForm.Visible = False
Shape2.FillColor = IIf(UseOldColor, &H511206, vbBlack)

End Sub

Private Sub Form_Load()

TmpOpts = SimOpts
  
 CurrSpec = -1 ' EricL 4/1/2006 Initialize that no species is selected
 DragInit    'Initialize drag code
 strings Me

 Form1.ScaleWidth = SimOpts.FieldWidth
 Form1.ScaleHeight = SimOpts.FieldHeight
 If TmpOpts.FieldWidth = 0 Then TmpOpts.FieldWidth = 16000
 If TmpOpts.FieldHeight = 0 Then TmpOpts.FieldHeight = 12000
 datatolist
 validate = True
 DispSettings
 validate = False
End Sub

Private Sub LightText_lostfocus()
  Dim a As Single
  a = val(LightText.Text)
  If a < LightUpDn.Min Then a = LightUpDn.Min
  If a > LightUpDn.Max Then a = LightUpDn.Max
  LightUpDn.value = a
  TmpOpts.LightIntensity = a
  tmrLight.Enabled = False
End Sub

Private Sub OKButton_Click()
'Botsareus 5/13/2013 Safemode restrictions
If Form1.lblSafeMode.Visible Then
    MsgBox "Can not change settings during safemode"
    Exit Sub
End If

savesett MDIForm1.MainDir + "\settings\lastran.set"


    SunPosition = 0.5
    SunRange = 1


Form1.camfix = False 'Botsareus 2/23/2013 When simulation starts the screen is normailized

  Dim i As Integer
  Dim k As Integer
  Dim a As Integer
  Dim fso As FileSystemObject
  Set fso = New FileSystemObject
  
  DragEnd
  
  MDIForm1.Combo1.CLEAR
  MDIForm1.Combo1.additem WSnone
  
  For i = 0 To TmpOpts.SpeciesNum - 1
    If i > MAXNATIVESPECIES Then
      MsgBox "Exceeded number of native species."
    Else
       MDIForm1.Combo1.additem TmpOpts.Specie(i).Name
    End If
  Next
    
  Canc = False
    
  'These change values while a sim is running.
  'Copy them so the correct value will be put back into SimOpts
  TmpOpts.TotRunCycle = SimOpts.TotRunCycle
  TmpOpts.TotBorn = SimOpts.TotBorn
  TmpOpts.TotRunTime = SimOpts.TotRunTime
  TmpOpts.DayNightCycleCounter = SimOpts.DayNightCycleCounter
  TmpOpts.Daytime = SimOpts.Daytime
  If TmpOpts.DayNight = False Then TmpOpts.Daytime = True
  
  
  SimOpts = TmpOpts
  
  Form1.ScaleWidth = SimOpts.FieldWidth
  Form1.ScaleHeight = SimOpts.FieldHeight
  Form1.visiblew = SimOpts.FieldWidth
  Form1.visibleh = SimOpts.FieldHeight
  Form1.xDivisor = 1
  Form1.yDivisor = 1
  If SimOpts.FieldWidth > 32000 Then Form1.xDivisor = SimOpts.FieldWidth / 32000
  If SimOpts.FieldHeight > 32000 Then Form1.yDivisor = SimOpts.FieldHeight / 32000
  Form1.SecTimer.Enabled = True
  Form1.Active = True
  
        'Botsareus 1/5/2013 more fix for n-vedio button
         MDIForm1.visualize = True 'Botsareus 1/11/2013 set to true
  
  'Me.Hide
  Unload Me
End Sub

Sub StartNew_Click() 'startnew

  If Form1.Visible Then
    If MsgBox("Are you sure?", vbYesNo, "About to start a new simulation") = vbNo Then
        Exit Sub
    End If
  End If

If chseedstartnew Then TmpOpts.UserSeedNumber = Timer * 100 'Botsareus 5/3/2013 Change seed on start new

  'Botsareus 5/7/2013 Safemode component
  Form1.lblSafeMode.Visible = False
  
    MDIForm1.Objects.Enabled = True
    MDIForm1.inssp.Enabled = True
    MDIForm1.DisableArep.Enabled = True
  
  MDIForm1.AutoFork.Enabled = True
  
  
  If dir(MDIForm1.MainDir + "\saves\localcopy.sim") <> "" Then Kill (MDIForm1.MainDir + "\saves\localcopy.sim")
  If dir(MDIForm1.MainDir + "\saves\lastautosave.sim") <> "" Then Kill (MDIForm1.MainDir + "\saves\lastautosave.sim")

  Dim i As Integer
  Dim k As Integer
  Dim t As Integer
  
  DragEnd
  
  ' EricL Moved here from StartSimul() so that cycles counter isn't reset when sim settings
  ' are changed.  Only reset when a new simulation is started.
  TmpOpts.TotRunCycle = -1
  
  MDIForm1.visualize = True
  Form1.Label1.Visible = False
  
  MDIForm1.Combo1.CLEAR
  MDIForm1.Combo1.additem WSnone
  
  For i = 0 To TmpOpts.SpeciesNum - 1
    If i > MAXNATIVESPECIES Then
      MsgBox "Exceeded number of native species."
    Else
       MDIForm1.Combo1.additem TmpOpts.Specie(i).Name
    End If
  Next i
  
  For t = 0 To TmpOpts.SpeciesNum - 1
    TmpOpts.Specie(t).population = TmpOpts.Specie(t).qty
    TmpOpts.Specie(t).SubSpeciesCounter = 0
  Next t
  
    MDIForm1.F1Piccy.Visible = False
    
  TmpOpts.SpeciesNum = SpecList.ListCount
  Canc = False
  'Me.Hide
  
      'Botsareus 1/5/2013 more fix for n-vedio button
      MDIForm1.visualize = False
      MDIForm1.menuupdate
  
  Unload Me
  
  SimOpts = TmpOpts
  
  If Form1.Active Then Form1.SecTimer.Enabled = True
  
    Form1.StartSimul
    SimOpts.UserSeedNumber = Rnd * 2147483647 'Botsareus 6/11/2013 Randomize seed on restart, moved to after first sim

End Sub

Private Sub DispSettings()
'Botsareus 5/3/2013 No need to change seed if it is fully randomized in any condition
If chseedloadsim And chseedstartnew Then
    Newseed.Enabled = False
    UserSeedText.Enabled = False
End If

  Dim t As Integer
  
  
  'PauseButton.Caption = IIf(Form1.Active, "Unpaused", "Paused") 'Botsareus 1/5/2013 Can not control pause from settings
  
 
  FieldSizeSlide.value = TmpOpts.FieldSize
  MaxPopText.Text = TmpOpts.MaxPopulation
  MinVegText.Text = TmpOpts.MinVegs
  RepopAmountText.Text = TmpOpts.RepopAmount
  RepopCooldownText.Text = TmpOpts.RepopCooldown
  DecayText.Text = TmpOpts.Decay
  LightText.Text = TmpOpts.LightIntensity
  Gradient.Text = (TmpOpts.Gradient - 1) * 10 'Botsareus 12/12/2012 Fix for Gradient
  GradientUpDn.value = Gradient.Text * 5
'  DNLength.text = TmpOpts.CycleLength
  UserSeedText.Text = TmpOpts.UserSeedNumber
  FrequencyText.Text = TmpOpts.Decaydelay
  
  FWidthLab.Caption = TmpOpts.FieldWidth
  FHeightLab.Caption = TmpOpts.FieldHeight
  
  'DisableTiesCheck.value = TmpOpts.DisableTies * True
  TopDownCheck.value = TmpOpts.Updnconnected * True
  RightLeftCheck.value = TmpOpts.Dxsxconnected * True
  
  Pondcheck.value = TmpOpts.Pondmode * True
  Gradient.Enabled = TmpOpts.Pondmode * True
  LightText.Enabled = TmpOpts.Pondmode * True
  LightUpDn.Enabled = TmpOpts.Pondmode * True
  GradientUpDn.Enabled = TmpOpts.Pondmode * True
  
  CorpseCheck.value = TmpOpts.CorpseEnabled * True
  DecayText.Enabled = TmpOpts.CorpseEnabled * True
  DecayUpDn(0).Enabled = TmpOpts.CorpseEnabled * True
  DecayUpDn(1).Enabled = TmpOpts.CorpseEnabled * True
  
  'EricL 4/11/2006 Added these to initialize the control
  DecayOption(0).Enabled = TmpOpts.CorpseEnabled * True
  DecayOption(1).Enabled = TmpOpts.CorpseEnabled * True
  DecayOption(2).Enabled = TmpOpts.CorpseEnabled * True
  DecayOption(0).value = False
  DecayOption(1).value = False
  DecayOption(2).value = False
  
  'EricL 4/11/2006 Don't ask me why, but historically, DecayType was stored as DecayOption Index +1
  'So this IF statement is here to gaurd against the case where DecayType has never been set and has value 0 by default.
  If (TmpOpts.DecayType > 0) Then
   DecayOption(TmpOpts.DecayType - 1) = True
  Else
    DecayOption(0) = True
  End If
    
 ' DNLength.Enabled = TmpOpts.DayNight * True
'  DNCycleUpDn.Enabled = TmpOpts.DayNight * True
'  DNCheck.value = TmpOpts.DayNight * True
  
'Botsareus 5/3/2013 Replaced by safe mode
'  UserSeed.value = TmpOpts.UserSeedToggle * True
'  UserSeedText.Enabled = TmpOpts.UserSeedToggle * True
  
  FrequencyText.Text = TmpOpts.Decaydelay
  
  Prop.Text = Str(TmpOpts.EnergyProp * 100)
  Fixed.Text = (TmpOpts.EnergyFix)
  FixUpDn.value = TmpOpts.EnergyFix
  ExchangeProp.value = TmpOpts.EnergyExType
  ExchangeFix.value = Not TmpOpts.EnergyExType
  
  If TmpOpts.MutCurrMult <= 0 Then TmpOpts.MutCurrMult = 1
  MutSlide.value = Log(TmpOpts.MutCurrMult) / Log(2)
 
  If TmpOpts.MutCurrMult > 1 Then
    MutLab.Caption = CStr(TmpOpts.MutCurrMult) + " X"
  Else
    MutLab.Caption = "1/" + Str(2 ^ -MutSlide.value) + " X"
  End If
     
  DisableMutationsCheck.value = TmpOpts.DisableMutations * True
    
  MaxNRGText.Text = TmpOpts.MaxEnergy
  
  'EricL 5/7/2006 Initialize new UI
  FluidSolidRadio(TmpOpts.FluidSolidCustom).value = True
  CostRadio(TmpOpts.CostRadioSetting).value = True

    
  If TmpOpts.CoefficientKinetic = 0.75 And _
    TmpOpts.CoefficientStatic = 0.9 And _
    TmpOpts.Zgravity = 4 Then
    FrictionCombo.Text = FrictionCombo.List(0)
  ElseIf TmpOpts.CoefficientKinetic = 0.4 And _
    TmpOpts.CoefficientStatic = 0.6 And _
    TmpOpts.Zgravity = 2 Then
    FrictionCombo.Text = FrictionCombo.List(1)
  ElseIf TmpOpts.CoefficientStatic = 0.05 And _
    TmpOpts.CoefficientKinetic = 0.05 And _
    TmpOpts.Zgravity = 1 Then
    FrictionCombo.Text = FrictionCombo.List(2)
  ElseIf TmpOpts.CoefficientStatic = 0# And _
    TmpOpts.CoefficientKinetic = 0# And _
    TmpOpts.Zgravity = 0 Then
    FrictionCombo.Text = FrictionCombo.List(3)
  Else
    FrictionCombo.Text = "Custom"
  End If
  
  'needs work
  If TmpOpts.Viscosity = 0.01 And _
     TmpOpts.Density = 0.0000001 Then
     DragCombo.Text = DragCombo.List(0)
  ElseIf TmpOpts.Viscosity = 0.0005 And _
     TmpOpts.Density = 0.0000001 Then
     DragCombo.Text = DragCombo.List(1)
  ElseIf TmpOpts.Viscosity = 0.000025 And _
     TmpOpts.Density = 0.0000001 Then
     DragCombo.Text = DragCombo.List(2)
  ElseIf TmpOpts.Viscosity = 0# And _
     TmpOpts.Density = 0# Then
     DragCombo.Text = DragCombo.List(3)
  Else
    DragCombo.Text = "Custom"
  End If
    
  MaxVelSlider.value = TmpOpts.MaxVelocity
  BodyNrgDist.value = TmpOpts.VegFeedingToBody * 100
    CustomWaste.Text = TmpOpts.BadWastelevel
  
  Elasticity.value = TmpOpts.CoefficientElasticity * 10
  Elasticity.Text = TmpOpts.CoefficientElasticity
   
  FixBotRadius.value = TmpOpts.FixedBotRadii * True
  
  ' Botsareus 12/12/2012 SimOpts should never overwrite TmpOpts during display settings
  'Do this so the right CostX gets put back into SimOpts even when no Cost changes are made
 'TmpOpts.Costs(COSTMULTIPLIER) = SimOpts.Costs(COSTMULTIPLIER)
  
  ' EricL Initialize that no species is selected
  CurrSpec = -1
  SpecVeg.value = 0
  BlockSpec.value = 0
  DisableVisionCheck.value = 0
  DisableMovementSysvarsCheck.value = 0
  DisableReproductionCheck.value = 0
  DisableDNACheck.value = 0
  VirusImmuneCheck.value = 0
  
' Botsareus 12/12/2012 SimOpts should never overwrite TmpOpts during display settings
'  'So the right value gets put back in when the dialog is closed and tmpopts is copied back into simopts...
'  'TmpOpts.SpeciesNum = SimOpts.SpeciesNum
'  For t = 0 To TmpOpts.SpeciesNum - 1
'    TmpOpts.Specie(t) = SimOpts.Specie(t) ' Population
'    'TmpOpts.Specie(t).SubSpeciesCounter = SimOpts.Specie(t).SubSpeciesCounter
'    'TmpOpts.Specie(t).Native = SimOpts.Specie(t).Native
'  Next t
       
  'display the scriptlist
  SpecList.Refresh
  Update
  
    
End Sub


Sub LoadSettings_Click() 'opensettings
'On Error GoTo fine
  CommonDialog1.FileName = ""
  CommonDialog1.InitDir = MDIForm1.MainDir + "\settings"
  CommonDialog1.Filter = "Settings file(*.set)|*.set"
  CommonDialog1.ShowOpen
  If CommonDialog1.FileName <> "" Then
    ReadSett CommonDialog1.FileName
    datatolist
    Dim i As Long
    
    SpecList.CLEAR
    For i = 0 To TmpOpts.SpeciesNum - 1
        SpecList.additem (TmpOpts.Specie(i).Name)
    Next i
    
    validate = True
    DispSettings
    validate = False
    
  End If
  Exit Sub
fine:
  MsgBox "Error loading"
End Sub

Private Sub SaveSettings_Click() 'savesettings
'On Error GoTo fine
  CommonDialog1.InitDir = MDIForm1.MainDir + "\settings"
  CommonDialog1.FileName = ""
  CommonDialog1.Filter = "Settings file(*.set)|*.set"
  CommonDialog1.ShowSave
  If CommonDialog1.FileName <> "" Then savesett CommonDialog1.FileName
fine:
End Sub

Public Sub savesett(path As String)
On Error GoTo fine
  Dim t As Integer
  Dim k As Integer
  Dim numSpecies As Integer
  Open path For Output As #1
  
  'EricL 4/13/2006 using this for version information
  '-2 is pre 2.42.2 versions
  '-1 is 2.42.2 and beyond VB fork
  '-3 is the C++ version
  t = -1
  Write #1, t
  
  numSpecies = 0
  For t = 0 To TmpOpts.SpeciesNum - 1
    numSpecies = numSpecies + 1
  Next t
  
  
  Write #1, numSpecies - 1  ' done this way becuase of the busted way the read routine loops through species...
  For t = 0 To TmpOpts.SpeciesNum - 1
    
    Write #1, TmpOpts.FieldWidth
    Write #1, TmpOpts.FieldHeight
    Write #1, 0
    Write #1, 0
    
    Write #1, TmpOpts.Specie(t).Mutables.mutarray(0)
    Write #1, TmpOpts.Specie(t).path
    Write #1, TmpOpts.Specie(t).qty
    Write #1, TmpOpts.Specie(t).Name
    Write #1, TmpOpts.Specie(t).Veg
    Write #1, TmpOpts.Specie(t).Fixed
    Write #1, TmpOpts.Specie(t).color
    Write #1, TmpOpts.Specie(t).Colind
    Write #1, TmpOpts.Specie(t).Stnrg
    For k = 0 To 13
      Write #1, TmpOpts.Specie(t).Mutables.mutarray(k)
    Next k
    For k = 0 To 12
      Write #1, TmpOpts.Specie(t).Skin(k)
    Next k
skipthisspecie:
  Next t
  Write #1, TmpOpts.SimName
  'TmpOpts.SpeciesNum = SpecList.ListCount
  
  'generali
  Write #1, numSpecies 'Botsareus 1/21/2013 No more non-native species -bug fixed
  Write #1, TmpOpts.FieldSize
  Write #1, TmpOpts.FieldWidth
  Write #1, TmpOpts.FieldHeight
  Write #1, TmpOpts.MaxPopulation
  Write #1, TmpOpts.BlockedVegs
  Write #1, TmpOpts.DisableTies
  Write #1, TmpOpts.PopLimMethod
  Write #1, TmpOpts.Toroidal
  Write #1, TmpOpts.MaxEnergy
  Write #1, TmpOpts.MinVegs
  
  'costi
  Write #1, TmpOpts.CostExecCond
  Write #1, TmpOpts.Costs(COSTSTORE)
  Write #1, TmpOpts.Costs(SHOTCOST)
  Write #1, TmpOpts.EnergyProp
  Write #1, TmpOpts.EnergyFix
  Write #1, TmpOpts.EnergyExType
  
  'fisica
  Write #1, TmpOpts.Ygravity
  Write #1, TmpOpts.Zgravity
  Write #1, TmpOpts.PhysBrown
  Write #1, TmpOpts.PhysMoving
  Write #1, TmpOpts.PhysSwim
  
  Write #1, "null"
  Write #1, 0
  Write #1, "null"
  Write #1, 0
  
  Write #1, TmpOpts.MutCurrMult
  
  Write #1, "null"
  Write #1, TmpOpts.DeadRobotSnp
  Write #1, TmpOpts.SnpExcludeVegs
  
  Write #1, TmpOpts.Pondmode
  Write #1, False
  Write #1, TmpOpts.LightIntensity
  Write #1, TmpOpts.CorpseEnabled
  Write #1, TmpOpts.Decay
  Write #1, TmpOpts.Gradient
  Write #1, TmpOpts.DayNight
  Write #1, TmpOpts.CycleLength
  
  Write #1, TmpOpts.DecayType
  Write #1, TmpOpts.Decaydelay
  
  'obsolete
  Write #1, TmpOpts.Costs(MOVECOST)
  
  Write #1, False
  Write #1, TmpOpts.Restart
  
  SaveScripts
  
  Write #1, TmpOpts.Dxsxconnected
  Write #1, TmpOpts.Updnconnected
  
  Write #1, TmpOpts.RepopAmount
  Write #1, TmpOpts.RepopCooldown
  
  Write #1, TmpOpts.ZeroMomentum
  Write #1, True 'Botsareus 5/3/2013 Replaced by safemode
  Write #1, TmpOpts.UserSeedNumber
  
  For t = 0 To TmpOpts.SpeciesNum - 1
    For k = 14 To 20
      Write #1, TmpOpts.Specie(t).Mutables.mutarray(k)
    Next k
    Write #1, TmpOpts.Specie(t).Mutables.Mutations
skipthisspecie2:
  Next t
  
  Write #1, CInt(0)
  Write #1, TmpOpts.VegFeedingToBody
  
  'New for 2.4:
  Write #1, TmpOpts.CoefficientStatic
  Write #1, TmpOpts.CoefficientKinetic
  Write #1, TmpOpts.PlanetEaters
  Write #1, TmpOpts.PlanetEatersG
  Write #1, TmpOpts.Viscosity
  Write #1, TmpOpts.Density
  
  
  For k = 0 To TmpOpts.SpeciesNum - 1
    Write #1, TmpOpts.Specie(k).Mutables.CopyErrorWhatToChange
    Write #1, TmpOpts.Specie(k).Mutables.PointWhatToChange
    
    Dim h As Integer
    
    For h = 0 To 20
      Write #1, TmpOpts.Specie(k).Mutables.Mean(h)
      Write #1, TmpOpts.Specie(k).Mutables.StdDev(h)
    Next h
skipthisspecie3:
  Next k
  
  For k = 0 To 70
    Write #1, TmpOpts.Costs(k)
  Next k
  
  For k = 0 To TmpOpts.SpeciesNum - 1
    Write #1, TmpOpts.Specie(k).Poslf
    Write #1, TmpOpts.Specie(k).Posrg
    Write #1, TmpOpts.Specie(k).Postp
    Write #1, TmpOpts.Specie(k).Posdn
skipthisspecie4:
  Next k
  
  Write #1, TmpOpts.MaxVelocity
  Write #1, TmpOpts.BadWastelevel 'EricL 4/1/2006 Added this
  Write #1, TmpOpts.chartingInterval 'EricL 4/1/2006 Added this
  Write #1, TmpOpts.FluidSolidCustom 'EricL 5/7/2006
  Write #1, TmpOpts.CostRadioSetting ' EricL 5/7/2006
  Write #1, TmpOpts.CoefficientElasticity ' EricL 5/7/2006
  Write #1, TmpOpts.MaxVelocity ' EricL 5/15/2006
  Write #1, TmpOpts.NoShotDecay ' EricL 6/8/2006
  Write #1, TmpOpts.SunUpThreshold 'EricL 6/8/2006 Added this
  Write #1, TmpOpts.SunUp 'EricL 6/8/2006 Added this
  Write #1, TmpOpts.SunDownThreshold 'EricL 6/8/2006 Added this
  Write #1, TmpOpts.SunDown 'EricL 6/8/2006 Added this
  Write #1, False
  Write #1, False
  Write #1, TmpOpts.FixedBotRadii
  Write #1, TmpOpts.SunThresholdMode
    
  'Botsareus 4/17/2013
  Write #1, TmpOpts.DisableTypArepro
  
  'Botsareus 4/27/2013 Rest of species data
  For t = 0 To numSpecies - 1
    Write #1, TmpOpts.Specie(t).CantSee
    Write #1, TmpOpts.Specie(t).DisableDNA
    Write #1, TmpOpts.Specie(t).DisableMovementSysvars
    Write #1, TmpOpts.Specie(t).CantReproduce
    Write #1, TmpOpts.Specie(t).VirusImmune
  Next
  
  Write #1, TmpOpts.NoWShotDecay 'Botsareus 9/28/2013
  
  'Botsareus 1/5/2014 Save obstacle data
  Write #1, TmpOpts.makeAllShapesTransparent
  Write #1, TmpOpts.makeAllShapesBlack
  Write #1, TmpOpts.shapeDriftRate
  Write #1, TmpOpts.allowHorizontalShapeDrift
  Write #1, TmpOpts.allowVerticalShapeDrift
  Write #1, TmpOpts.shapesAreSeeThrough
  Write #1, TmpOpts.shapesAbsorbShots
  Write #1, TmpOpts.shapesAreVisable
  
  Dim o As Integer
  
  Write #1, 0
  Write #1, 0
  Write #1, 0
  Write #1, 0
  
  'Botsareus 3/28/2014 Some more species data
  For t = 0 To numSpecies - 1
    Write #1, TmpOpts.Specie(t).NoChlr
  Next

  'Botsareus 7/30/2014 Some UI exposed to settings
  Write #1, MDIForm1.displayResourceGuagesToggle
  Write #1, MDIForm1.displayMovementVectorsToggle
  Write #1, MDIForm1.displayShotImpactsToggle
  Write #1, MDIForm1.showVisionGridToggle
  
  'Botsareus 8/5/2014 Disable fixing
  Write #1, TmpOpts.DisableFixing
  
  'Botsareus 8/23/2014 Tides
  Write #1, TmpOpts.Tides
  Write #1, TmpOpts.TidesOf
  
  Close 1
  Exit Sub
fine:
  MsgBox ("Unable to save settings: some error occurred")
End Sub

Public Sub ReadSett(path As String)
On Error GoTo aiuto
  Dim t As Integer
  Dim col As Single
  Dim m As Integer
  Dim k As Integer
  Dim maxs As Integer
  Dim longv As Long
  Dim intv As Integer
  Dim singv As Single
  Dim boolv As Boolean
  Dim b As Integer
  

carica:
  On Error GoTo aiuto
  Open path For Input As #1
  Input #1, maxs 'we can actually use this for version info
  
  'EricL 4/13/2006 Check for older settings files.  2.42.1 fixed bugs that introduced incomptabilities...
  If maxs <> -1 Then
    If Right(path, 12) = "lastexit.set" Then
      MsgBox ("The settings from your last exit are incomptable with this version.  Last exit settings not loaded.  When you exit the program, a new lastexit settings file will be created automatically.")
    Else
      If Right(path, 11) = "default.set" Then
        MsgBox ("The default settings file is incomptable with this version.  You can save your settings to default.set to create a new one.  Settings not loaded.")
      Else
        MsgBox ("The settings file is incomptable with this version.  Settings not loaded.")
      End If
    End If
  Else
    clearall
    ReadSettFromFile
  End If
  Close 1
  'EricL 3/21/2006 Added following three lines to work around problem with default settings file
  If (TmpOpts.MaxEnergy = 500000) And (Right(path, 11) = "default.set") Then
    TmpOpts.MaxEnergy = 50
  End If
  lastsettings = path
  Exit Sub
aiuto:
  Close 1
  
  'EricL 3/22/2006 Added following section to work around case where there is no settings directory
  If Err.Number = 76 Then
    b = MsgBox("Cannot find the Settings Directory.  " + vbCrLf + _
            "Would you like me to create one?   " + vbCrLf + vbCrLf + _
            "If this is a new install, choose OK.", vbOKCancel + vbQuestion)
    If b = vbOK Then
       RecursiveMkDir (MDIForm1.MainDir + "\settings")
    Else
      MsgBox ("Darwinbots cannot continue.  Program will exit.")
      End 'Botsareus 7/12/2012 force DB to exit
    End If
  ElseIf Err.Number = 53 And Right(path, 12) = "lastexit.set" Then
       MsgBox ("Cannot find the settings file from your last exit.  " + vbCrLf + _
               "Using the internal default settings. " + vbCrLf + vbCrLf + _
               "If this is a new install, this is normal.")
  Else
    MsgBox MBcannotfindI, , MBwarning
    CommonDialog1.FileName = path
    CommonDialog1.ShowOpen
    path = CommonDialog1.FileName
    If path <> "" Then GoTo carica
  End If
  
End Sub



Public Sub ReadSettFromFile()
  Dim maxsp As Integer
  Dim strvar As String
  Dim check As Boolean
  Dim sinvar As Single
  Dim t As Integer
  Dim k As Integer
  Dim obsoleteLong As Long ' EricL 3/28/2006 Added to read obsolete long values from the saved settings file
  Dim obsoleteInt As Integer
  Dim obsoleteString As String
  Dim obsoleteBool As Boolean
  Input #1, maxsp
  
  For t = 0 To maxsp
    TmpOpts.Specie(t).Posrg = 1
    TmpOpts.Specie(t).Posdn = 1
    TmpOpts.Specie(t).Poslf = 0
    TmpOpts.Specie(t).Postp = 0
    
    'Obsolete
    Input #1, obsoleteLong 'EricL 3/28/2006 Changed to Long from k to fix bug with reloading settings with large fields
    Input #1, obsoleteLong 'EricL 3/28/2006 Changed to Long from k to fix bug with reloading settings with large fields
    Input #1, k
    Input #1, k
    
    Input #1, TmpOpts.Specie(t).Mutables.mutarray(0)
    Input #1, TmpOpts.Specie(t).path
    Input #1, TmpOpts.Specie(t).qty
    Input #1, TmpOpts.Specie(t).Name
    Input #1, TmpOpts.Specie(t).Veg
    Input #1, TmpOpts.Specie(t).Fixed
    Input #1, TmpOpts.Specie(t).color
    Input #1, TmpOpts.Specie(t).Colind
    Input #1, TmpOpts.Specie(t).Stnrg
    For k = 0 To 13
      Input #1, TmpOpts.Specie(t).Mutables.mutarray(k)
    Next k
    For k = 0 To 12
      Input #1, TmpOpts.Specie(t).Skin(k)
    Next k
  Next t
  
  'generali
  Input #1, TmpOpts.SimName
  Input #1, TmpOpts.SpeciesNum
  Input #1, TmpOpts.FieldSize
  Input #1, TmpOpts.FieldWidth
  Input #1, TmpOpts.FieldHeight
  Input #1, TmpOpts.MaxPopulation
  Input #1, TmpOpts.BlockedVegs
  Input #1, TmpOpts.DisableTies
  Input #1, TmpOpts.PopLimMethod
  Input #1, TmpOpts.Toroidal
  Input #1, TmpOpts.MaxEnergy
  Input #1, TmpOpts.MinVegs
  
  'costi
  Input #1, TmpOpts.CostExecCond
  Input #1, TmpOpts.Costs(COSTSTORE)
  Input #1, TmpOpts.Costs(SHOTCOST)
  Input #1, TmpOpts.EnergyProp
  Input #1, TmpOpts.EnergyFix
  Input #1, TmpOpts.EnergyExType
  
  'fisica
  Input #1, TmpOpts.Ygravity
  Input #1, TmpOpts.Zgravity
  Input #1, TmpOpts.PhysBrown
  Input #1, TmpOpts.PhysMoving
  Input #1, TmpOpts.PhysSwim
  
  Input #1, obsoleteString
  Input #1, obsoleteInt
  Input #1, obsoleteString
  Input #1, obsoleteInt
  Input #1, TmpOpts.MutCurrMult
  Input #1, obsoleteString
  Input #1, TmpOpts.DeadRobotSnp
  Input #1, TmpOpts.SnpExcludeVegs
  
  If Not EOF(1) Then Input #1, TmpOpts.Pondmode
  If Not EOF(1) Then Input #1, TmpOpts.CorpseEnabled 'dummy variable
  If Not EOF(1) Then Input #1, TmpOpts.LightIntensity
  If Not EOF(1) Then Input #1, TmpOpts.CorpseEnabled
  If Not EOF(1) Then Input #1, TmpOpts.Decay
  If Not EOF(1) Then Input #1, TmpOpts.Gradient
  If Not EOF(1) Then Input #1, TmpOpts.DayNight
  If Not EOF(1) Then Input #1, TmpOpts.CycleLength
  If Not EOF(1) Then Input #1, TmpOpts.DecayType
  If Not EOF(1) Then Input #1, TmpOpts.Decaydelay
  
  'obsolete
  If Not EOF(1) Then Input #1, TmpOpts.Costs(MOVECOST)
  
  Dim discard As Boolean
  
  If Not EOF(1) Then Input #1, discard
  If Not EOF(1) Then Input #1, TmpOpts.Restart
  
  LoadScripts 'load up the scripts. Only available in form1. Can't access them from here.
  
  'even even newer newer stuff
  If Not EOF(1) Then Input #1, TmpOpts.Dxsxconnected
  If Not EOF(1) Then Input #1, TmpOpts.Updnconnected
  If Not EOF(1) Then Input #1, TmpOpts.RepopAmount
  If Not EOF(1) Then Input #1, TmpOpts.RepopCooldown
  If Not EOF(1) Then Input #1, TmpOpts.ZeroMomentum
  If Not EOF(1) Then Input #1, check 'Botsareus 5/3/2013 Replaced by safemode
  If Not EOF(1) Then Input #1, TmpOpts.UserSeedNumber
  
  For t = 0 To maxsp
    For k = 14 To 20
      If Not EOF(1) Then Input #1, TmpOpts.Specie(t).Mutables.mutarray(k)
    Next k
    
    If Not EOF(1) Then Input #1, TmpOpts.Specie(t).Mutables.Mutations
  Next t
  
  If Not EOF(1) Then Input #1, obsoleteInt
  If Not EOF(1) Then: Input #1, TmpOpts.VegFeedingToBody: Else: TmpOpts.VegFeedingToBody = 0.1
  
  'New for 2.4
  If Not EOF(1) Then Input #1, TmpOpts.CoefficientStatic
  If Not EOF(1) Then Input #1, TmpOpts.CoefficientKinetic
  If Not EOF(1) Then Input #1, TmpOpts.PlanetEaters
  If Not EOF(1) Then Input #1, TmpOpts.PlanetEatersG
  If Not EOF(1) Then Input #1, TmpOpts.Viscosity
  If Not EOF(1) Then Input #1, TmpOpts.Density
  
  For k = 0 To maxsp
    If Not EOF(1) Then Input #1, TmpOpts.Specie(k).Mutables.CopyErrorWhatToChange
    If Not EOF(1) Then Input #1, TmpOpts.Specie(k).Mutables.PointWhatToChange
    
    Dim h As Integer
    
    For h = 0 To 20
      If Not EOF(1) Then Input #1, TmpOpts.Specie(k).Mutables.Mean(h)
      If Not EOF(1) Then Input #1, TmpOpts.Specie(k).Mutables.StdDev(h)
    Next h
  Next k
  
  For k = 0 To 70
    Dim temp As Single
    If Not EOF(1) Then Input #1, temp
    TmpOpts.Costs(k) = temp
  Next k
  
  For k = 0 To maxsp
    If Not EOF(1) Then Input #1, TmpOpts.Specie(k).Poslf
    If Not EOF(1) Then Input #1, TmpOpts.Specie(k).Posrg
    If Not EOF(1) Then Input #1, TmpOpts.Specie(k).Postp
    If Not EOF(1) Then Input #1, TmpOpts.Specie(k).Posdn
  Next k
  
  If Not EOF(1) Then Input #1, TmpOpts.MaxVelocity
  
  TmpOpts.BadWastelevel = -1 ' Default
  If Not EOF(1) Then Input #1, TmpOpts.BadWastelevel 'EricL 4/1/2006 Added this
  
  TmpOpts.chartingInterval = 200 ' Default
  If Not EOF(1) Then Input #1, TmpOpts.chartingInterval 'EricL 4/1/2006 Added this
  
  TmpOpts.FluidSolidCustom = 2 ' Default to custom for older settings files
  If Not EOF(1) Then Input #1, TmpOpts.FluidSolidCustom 'EricL 5/7/2006 Added this
  
  TmpOpts.CostRadioSetting = 2 ' Default to custom for older settings files
  If Not EOF(1) Then Input #1, TmpOpts.CostRadioSetting 'EricL 5/7/2006 Added this
  
  TmpOpts.CoefficientElasticity = 0 ' Default for older settings files
  If Not EOF(1) Then Input #1, TmpOpts.CoefficientElasticity 'EricL 5/7/2006 Added this
  
  TmpOpts.MaxVelocity = 40 ' Default for older settings files
  If Not EOF(1) Then Input #1, TmpOpts.MaxVelocity 'EricL 5/15/2006 Added this
  
  TmpOpts.NoShotDecay = False ' Default for older settings files
  If Not EOF(1) Then Input #1, TmpOpts.NoShotDecay 'EricL 6/8/2006 Added this
    
  TmpOpts.SunUpThreshold = 500000   'Set to a reasonable default value
  If Not EOF(1) Then Input #1, TmpOpts.SunUpThreshold 'EricL 6/8/2006 Added this
    
  TmpOpts.SunUp = False   'Set to a reasonable default value
  If Not EOF(1) Then Input #1, TmpOpts.SunUp 'EricL 6/8/2006 Added this
    
  TmpOpts.SunDownThreshold = 1000000   'Set to a reasonable default value
  If Not EOF(1) Then Input #1, TmpOpts.SunDownThreshold 'EricL 6/8/2006 Added this
    
  TmpOpts.SunDown = False   'Set to a reasonable default value
  If Not EOF(1) Then Input #1, TmpOpts.SunDown 'EricL 6/8/2006 Added this
    
  If Not EOF(1) Then Input #1, obsoleteBool
  If Not EOF(1) Then Input #1, obsoleteBool
  
  TmpOpts.FixedBotRadii = False
  If Not EOF(1) Then Input #1, TmpOpts.FixedBotRadii
  
  TmpOpts.SunThresholdMode = 0
  If Not EOF(1) Then Input #1, TmpOpts.SunThresholdMode
  
  TmpOpts.DisableTypArepro = 0
  If Not EOF(1) Then Input #1, TmpOpts.DisableTypArepro
  
  'Botsareus 4/37/2013 Rest of species data
  For t = 0 To maxsp
    If Not EOF(1) Then Input #1, TmpOpts.Specie(t).CantSee
    If Not EOF(1) Then Input #1, TmpOpts.Specie(t).DisableDNA
    If Not EOF(1) Then Input #1, TmpOpts.Specie(t).DisableMovementSysvars
    If Not EOF(1) Then Input #1, TmpOpts.Specie(t).CantReproduce
    If Not EOF(1) Then Input #1, TmpOpts.Specie(t).VirusImmune
  Next
  
  'Botsareus 9/28/2013 Do not decay waste shots
  TmpOpts.NoWShotDecay = False
  If Not EOF(1) Then Input #1, TmpOpts.NoWShotDecay
  
  'Botsareus 1/5/2014 Obstecle settings
  If Not EOF(1) Then Input #1, TmpOpts.makeAllShapesTransparent
  If Not EOF(1) Then Input #1, TmpOpts.makeAllShapesBlack
  If Not EOF(1) Then Input #1, TmpOpts.shapeDriftRate
  If Not EOF(1) Then Input #1, TmpOpts.allowHorizontalShapeDrift
  If Not EOF(1) Then Input #1, TmpOpts.allowVerticalShapeDrift
  If Not EOF(1) Then Input #1, TmpOpts.shapesAreSeeThrough
  If Not EOF(1) Then Input #1, TmpOpts.shapesAbsorbShots
  If Not EOF(1) Then Input #1, TmpOpts.shapesAreVisable
    
  Dim disc As Integer
  
  If Not EOF(1) Then Input #1, disc
  If Not EOF(1) Then Input #1, disc
  If Not EOF(1) Then Input #1, disc
  If Not EOF(1) Then Input #1, disc
  
  'Botsareus 3/28/2014 Some more species data
  For t = 0 To maxsp
    If Not EOF(1) Then Input #1, TmpOpts.Specie(t).NoChlr
  Next
  
  'Botsareus 7/18/2014 Some more settings exposed for Testlund
  If Not EOF(1) Then Input #1, check
  MDIForm1.displayResourceGuagesToggle = check
  MDIForm1.DisplayResourceGuages.Checked = check
  If Not EOF(1) Then Input #1, check
  MDIForm1.displayMovementVectorsToggle = check
  MDIForm1.DisplayMovementVectors.Checked = check
  If Not EOF(1) Then Input #1, check
  MDIForm1.displayShotImpactsToggle = check
  MDIForm1.DisplayShotImpacts.Checked = check
  If Not EOF(1) Then Input #1, check
  MDIForm1.showVisionGridToggle = check
  MDIForm1.ShowVisionGrid.Checked = check
  
  'Botsareus 8/5/2014 Disable fixing
  TmpOpts.DisableFixing = 0
  If Not EOF(1) Then Input #1, TmpOpts.DisableFixing
  
  'Botsareus 8/23/2014 Tides
  If Not EOF(1) Then Input #1, TmpOpts.Tides
  If Not EOF(1) Then Input #1, TmpOpts.TidesOf
  
  'Botsareus 10/8/2015
  If (Not EOF(1)) Then MsgBox "This settings file is a newer version than this version can read.  " + vbCrLf + _
                              "Not all the information it contains can be " + vbCrLf + _
                              "transfered."
  
  If TmpOpts.FieldWidth = 0 Then TmpOpts.FieldWidth = 16000
  If TmpOpts.FieldHeight = 0 Then TmpOpts.FieldHeight = 12000
  If TmpOpts.MaxVelocity = 0 Then TmpOpts.MaxVelocity = 60
  
  
  'EricL 4/13/2006 divide by zero protection for older settings files.
  If TmpOpts.chartingInterval = 0 Then
    TmpOpts.chartingInterval = 200
  End If
  
  TmpOpts.DayNightCycleCounter = 0 ' When you load settings, you don't get the state from the last sim
  TmpOpts.Daytime = True ' EricL 3/21/2006 - this is a bettter place for this than in MDIForm_Load
  
End Sub

Private Sub Prop_Lostfocus()
  TmpOpts.EnergyProp = val(Prop.Text) / 100
End Sub

Private Sub PropUpDn_Change()
  Dim a As Single
  a = PropUpDn.value
  Prop.Text = Str$(a / 100#)
  TmpOpts.EnergyProp = a / 100#
End Sub

Sub Fixed_Lostfocus()
  Dim a As Single
  a = val(Fixed.Text)
  If a < FixUpDn.Min Then a = FixUpDn.Min
  If a > FixUpDn.Max Then a = FixUpDn.Max
  TmpOpts.EnergyFix = a
End Sub

Private Sub FixUpDn_Change()
  Dim a As Single
  a = FixUpDn.value
  TmpOpts.EnergyFix = a
End Sub

Private Sub ExchangeProp_Click()
  TmpOpts.EnergyExType = ExchangeProp.value
End Sub

Private Sub ExchangeFix_Click()
  TmpOpts.EnergyExType = Not ExchangeFix.value
End Sub

'
'  M U T R A T E   O P T I O N S
'
'

Private Sub MutSlide_Scroll()
  TmpOpts.MutCurrMult = 2 ^ MutSlide.value
  If TmpOpts.MutCurrMult > 1 Then
    MutLab.Caption = Str$(TmpOpts.MutCurrMult) + " X"
  Else
    MutLab.Caption = "1/" + Str$(2 ^ -MutSlide.value) + " X"
  End If
End Sub
