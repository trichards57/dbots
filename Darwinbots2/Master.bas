Attribute VB_Name = "Master"
Option Explicit

Public PopulationLast10Cycles(10) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer 'Botsareus 12/11/2012 Pause Break Key to Pause code

Public Sub UpdateSim()
    Dim CurrentPopulation As Integer
    Dim AllChlr As Long
    Dim i As Integer
    Dim t As Integer
    Dim pos As Vector
    Dim totlen As Long

    If GetAsyncKeyState(vbKeyF12) Then
        DisplayActivations = False
        Form1.Active = False
        Form1.SecTimer.Enabled = False
        MDIForm1.unpause.Enabled = True
    End If
  
    SimOpts.TotRunCycle = SimOpts.TotRunCycle + 1
    TotalSimEnergyDisplayed = TotalSimEnergy(CurrentEnergyCycle)
    CurrentEnergyCycle = SimOpts.TotRunCycle Mod 100
    TotalSimEnergy(CurrentEnergyCycle) = 0
  
    CurrentPopulation = totnvegsDisplayed + totvegsDisplayed
  
    If SimOpts.TotRunCycle Mod 10 = 0 Then
        For i = 10 To 2 Step -1
            PopulationLast10Cycles(i) = PopulationLast10Cycles(i - 1)
        Next
        PopulationLast10Cycles(1) = CurrentPopulation
    End If
  
    ExecRobs
    
    If datirob.Visible And datirob.ShowMemoryEarlyCycle Then
        With rob(robfocus)
            datirob.infoupdate robfocus, .nrg, .parent, .Mutations, .age, .SonNumber, 1, .FName, .genenum, .LastMut, .generation, .DnaLen, .LastOwner, .Waste, .body, .mass, .venom, .shell, .Slime, .chloroplasts
        End With
    End If
  
    'updateshots can write to bot sense, so we need to clear bot senses before updating shots
    For t = 1 To MaxRobs
        If robManager.GetExists(t) And Not rob(t).DisableDNA Then EraseSenses t
    Next

    updateshots
  
    'Botsareus 6/22/2016 to figure out actual velocity of the bot incase there is a collision event
    For t = 1 To MaxRobs
        If robManager.GetExists(t) Then
            rob(t).opos = robManager.GetRobotPosition(t)
        End If
    Next
  
    UpdateBots
  
    'to figure out actual velocity of the bot incase there is a collision event
    For t = 1 To MaxRobs
        If robManager.GetExists(t) Then
            'Only if the robots position was already configured
            pos = robManager.GetRobotPosition(t)
            If Not (rob(t).opos.x = 0 And rob(t).opos.y = 0) Then robManager.SetActualVelocity t, VectorSub(pos, rob(t).opos)
        End If
    Next
  
    For t = 1 To MaxRobs 'Panda 8/14/2013 to figure out how much vegys to repopulate across all robots
        If robManager.GetExists(t) Then  'Botsareus 8/14/2013 We have to make sure the robot is alive first
            AllChlr = AllChlr + rob(t).chloroplasts
        End If
    Next
  
    TotalChlr = AllChlr / 16000 'Panda 8/23/2013 Calculate total unit chloroplasts
  
    If TotalChlr < CLng(SimOpts.MinVegs) Then   'Panda 8/23/2013 Only repopulate vegs when total chlroplasts below value
        If totvegsDisplayed <> -1 Then VegsRepopulate  'Will be -1 first cycle after loading a sim.  Prevents spikes.
    End If
  
    feedvegs SimOpts.MaxEnergy
  
    'okay, time to store some values for RGB monitor
    If MDIForm1.MonitorOn Then
        For t = 1 To MaxRobs
            If robManager.GetExists(t) Then
            With frmMonitorSet
                rob(t).monitor_r = rob(t).mem(.Monitor_mem_r)
                rob(t).monitor_g = rob(t).mem(.Monitor_mem_g)
                rob(t).monitor_b = rob(t).mem(.Monitor_mem_b)
            End With
        End If
        Next
    End If
  
    'Kill some robots to prevent out of memory
    For t = 1 To MaxRobs
        If robManager.GetExists(t) Then
            totlen = totlen + rob(t).DnaLen
        End If
    Next
    If totlen > 4000000 Then
        Dim calcminenergy As Single
        Dim selectrobot As Integer
        Dim maxdel As Long

        maxdel = 1500 * (CLng(TotalRobotsDisplayed) * 425 / totlen)

        For i = 0 To maxdel
            calcminenergy = 320000 'only erase robots with lowest energy
            For t = 1 To MaxRobs
                If robManager.GetExists(t) Then
                    If (rob(t).nrg + rob(t).body * 10) < calcminenergy Then
                        calcminenergy = (rob(t).nrg + rob(t).body * 10)
                        selectrobot = t
                    End If
                End If
            Next
            KillRobot selectrobot
        Next
    End If
    If totlen > 3000000 Then
        For t = 1 To MaxRobs
            rob(t).LastMutDetail = ""
        Next t
    End If

    If UseSafeMode Then 'special modes does not apply, may need to expended to other restart modes
        If SimOpts.TotRunCycle Mod 2000 = 0 And SimOpts.TotRunCycle > 0 Then
            SaveSimulation MDIForm1.MainDir + "\saves\lastautosave.sim"
            If dir(MDIForm1.MainDir + "\saves\localcopy.sim") <> "" Then Kill (MDIForm1.MainDir + "\saves\localcopy.sim")
            Open App.path & "\autosaved.gset" For Output As #1
            Write #1, True
            Close #1
        End If
    End If
End Sub
