Attribute VB_Name = "RndMod"
Public RndMas(1 To 48)

Public Sub Random()

Static a As Integer

For i = 1 To UBound(RndMas)
upp:
Randomize
a = (Rnd * 47) + 1

For e = 1 To (i - 1)
If RndMas(e) = a Then GoTo upp
Next e

RndMas(i) = a
Next i

End Sub
