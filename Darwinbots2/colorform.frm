VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ColorForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Color"
   ClientHeight    =   2232
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3948
   Icon            =   "colorform.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2232
   ScaleWidth      =   3948
   StartUpPosition =   3  'Windows Default
   Tag             =   "9001"
   Begin VB.CommandButton btnWrite 
      Caption         =   "Write color as data into DNA file"
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton UseColor 
      Caption         =   "Use This Color"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   1680
      Width           =   2415
   End
   Begin MSComctlLib.Slider SliderR 
      Height          =   225
      Left            =   630
      TabIndex        =   4
      Top             =   420
      Width           =   2850
      _ExtentX        =   5017
      _ExtentY        =   402
      _Version        =   393216
      Max             =   255
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider SliderG 
      Height          =   225
      Left            =   630
      TabIndex        =   5
      Top             =   630
      Width           =   2850
      _ExtentX        =   5017
      _ExtentY        =   381
      _Version        =   393216
      Max             =   255
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider SliderB 
      Height          =   225
      Left            =   630
      TabIndex        =   6
      Top             =   840
      Width           =   2850
      _ExtentX        =   5017
      _ExtentY        =   402
      _Version        =   393216
      Max             =   255
      TickStyle       =   3
   End
   Begin VB.Label LabelB 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   195
      Left            =   15
      TabIndex        =   3
      Top             =   780
      Width           =   450
   End
   Begin VB.Label LabelG 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   195
      Left            =   15
      TabIndex        =   2
      Top             =   585
      Width           =   450
   End
   Begin VB.Label LabelR 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   195
      Left            =   15
      TabIndex        =   1
      Top             =   390
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   "R G  B"
      Height          =   615
      Left            =   3615
      TabIndex        =   0
      Top             =   405
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      Height          =   690
      Left            =   120
      Shape           =   1  'Square
      Top             =   1320
      Width           =   615
   End
End
Attribute VB_Name = "ColorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rval As Long
Private gval As Long
Private bval As Long
Public OldColor As Long
Public UseThisColor As Boolean
Public color As Long
Public SelectColor As Boolean
Public path As String

Private Sub btnWrite_Click() 'Writes color as data into DNA file
  'Step 1, where is the dna file
  If dir(path) = "" Then
    Dim splt() As String
    splt = Split(path, "\")
    Dim namepart As String
    namepart = splt(UBound(splt))
    path = MDIForm1.MainDir + "\Robots\" & namepart
    If dir(path) = "" Then
      MsgBox "Robot not found!", vbCritical
      Exit Sub
    End If
  End If
  
  'Step 2, load Dna (ignore lines that def red, green , or blue) (initial lines that have ' will be moved)
  Dim dtl As String 'Data line
  Dim robot As String 'Whole robot
  
  Dim cmtrob As String
  
  Dim endofcmt As Boolean
  endofcmt = False
  
  Open path For Input As #1
  While Not EOF(1)
    Line Input #1, dtl
    
    If Trim(dtl) = "" Or Trim(dtl) Like "'*" And Not endofcmt Then 'initial comments move to top
      cmtrob = cmtrob & dtl & vbCrLf
    Else
      endofcmt = True
    
      If Not Trim(dtl) Like "def red*" And Not Trim(dtl) Like "def green*" And Not Trim(dtl) Like "def blue*" And dtl <> "@" Then
        robot = robot & dtl & vbCrLf
      End If
    End If
  Wend
  Close #1
  
  robot = Left(robot, Len(robot) - 2)
  If cmtrob <> "" Then cmtrob = Left(cmtrob, Len(cmtrob) - 2) 'trim back comments only if there where comments
  
  'Step 3 add back new values for red, green, and blue, and comments
  robot = "def blue " & bval & vbNewLine & robot
  robot = "def green " & gval & vbNewLine & robot
  robot = "def red " & rval & vbNewLine & robot
  robot = "@" & vbNewLine & robot 'Botsareus 11/29/2013 bug fix
  robot = cmtrob & vbNewLine & robot
  
  'Step 4 write back to dna file
  Open path For Output As #1
    Print #1, robot
  Close #1
  
  'Step5 use the color
  UseThisColor = True
  SelectColor = True
  Me.Hide
End Sub

Private Sub Form_Load()
  strings Me
  SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
  OldColor = color
  setcolor color
End Sub

Private Sub Form_Terminate()
  UseThisColor = True
  Me.Hide
End Sub

Private Sub SliderG_Scroll()
  gval = SliderG.value
  LabelG.Caption = Str$(gval)
  dispcolor
End Sub

Private Sub SliderR_Scroll()
  rval = SliderR.value
  LabelR.Caption = Str$(rval)
  dispcolor
End Sub

Private Sub SliderB_Scroll()
  bval = SliderB.value
  LabelB.Caption = Str$(bval)
  dispcolor
End Sub

Private Sub dispcolor()
  color = bval * 65536 + gval * 256 + rval
  Shape1.BackColor = color
End Sub

Sub setcolor(col As Long)
  bval = Int(col / 65536)
  col = col - bval * 65536
  gval = Int(col / 256)
  rval = col - gval * 256
  SliderB.value = bval
  SliderR.value = rval
  SliderG.value = gval
  LabelG.Caption = Str$(gval)
  LabelB.Caption = Str$(bval)
  LabelR.Caption = Str$(rval)
  dispcolor
End Sub

Private Sub UseColor_Click()
  UseThisColor = True
  SelectColor = True
  Me.Hide
End Sub
