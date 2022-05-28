Attribute VB_Name = "Flex"
Private Function Exists(ByRef key As String, ByRef Keys) As Boolean
  Dim k As Integer
  Dim u As Integer
  u = UBound(Keys)
  k = 1
  While Keys(k) <> key And Keys(k) <> ""
    k = k + 1
  Wend
  Exists = False
  If Keys(k) = key Then Exists = True
End Function

Private Function ConsPos(ByRef key As String, ByRef Keys) As Integer
  Dim k As Integer
  Dim u As Integer
  u = UBound(Keys)
  k = 1
  While Keys(k) <> key And Keys(k) <> ""
    k = k + 1
  Wend
  ConsPos = 0
  If Keys(k) = key Then
    ConsPos = k
  End If
End Function

Public Function Position(ByRef key As String, ByRef Keys) As Integer
  Dim k As Integer
  Dim u As Integer
  u = UBound(Keys)
  k = 1
  While Keys(k) <> key And Keys(k) <> "" And k < u
    k = k + 1
  Wend
  Position = 0
  If Keys(k) = key Or k < u Then
    Position = k
    Keys(k) = key
  End If
End Function

Private Sub Delete(ByRef key As String, ByRef Keys, ByRef data)
  Dim k As Integer, p As Integer
  Dim uk As Integer, ud1 As Integer, ud2 As Integer
  p = ConsPos(key, Keys)
  uk = UBound(Keys)
  ud1 = UBound(data, 1)
  ud2 = UBound(data, 2)
  If p > 0 Then
    For k = p To ud1 - 1
      For j = 0 To ud2
        data(k, j) = data(k + 1, j)
      Next j
    Next k
    For k = p To uk - 1
      Keys(k) = Keys(k + 1)
    Next k
    Keys(uk) = ""
  End If
End Sub

Public Function last(ByRef Keys) As Integer
  Dim k As Integer
  k = 1
  While Keys(k) <> ""
    k = k + 1
    If k > MAXNATIVESPECIES Then
      last = k - 1
      Exit Function
    End If
  Wend
  last = k - 1
End Function
