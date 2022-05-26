Attribute VB_Name = "varspecie"
Public Type mutationprobs
  'should the mutation even be triggered?
  Mutations As Boolean
  mutarray(20) As Single 'probs of the different mutations
  
  Mean(20) As Single
  StdDev(20) As Single
  
  'Extras
  PointWhatToChange As Integer
  CopyErrorWhatToChange As Integer
End Type

Public Type datispecie
  Skin(13) As Integer
  path As String
  Name As String
  Stnrg As Integer
  Veg As Boolean
  NoChlr As Boolean 'Botsareus 3/28/2014 Prevents robots from evolving chloroplasts
  Fixed As Boolean
  color As Long
  Colind As Integer
  Postp As Single
  Poslf As Single
  Posdn As Single
  Posrg As Single
  qty As Integer
  Comment As String
  Mutables As mutationprobs
  CantSee As Boolean                ' Flag indicating eyes should be turned off for this species
  DisableDNA As Boolean             ' Flag indicating DNA should not execute for this species
  DisableMovementSysvars As Boolean ' Flag indicating movement sysvars should be disabled for this species
  CantReproduce As Boolean          ' Flag indicating whether reproduction has been disabled for this species.
  VirusImmune As Boolean            ' Flag indicating whether members of this species are suceptable to viruses.
  population As Integer             ' Number of this species in the sim.  Updated each cycle.
  SubSpeciesCounter As Integer      ' Using to increment the per-bot subspecies.
  DisplayImage As Image
End Type

