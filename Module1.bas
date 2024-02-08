Attribute VB_Name = "Module1"
Public Function HASH(ByVal Text As String) As String
A = 1
For i = 1 To Len(Text)
    A = Sqr(A * i * Asc(Mid(Text, i, 1))) 'Numeric Hash
Next i
Rnd (-1)
Randomize A 'seed PRNG

For i = 1 To 16
    HASH = HASH & Chr(Int(Rnd * 256))
Next i
End Function

