VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSHA256"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Visual Basic SHA-256 Implementation
' Phil Fresle and David Midkiff (mznull@earthlink.net)
'
' Standard SHA256 implementation optimised for the Visual Basic environment.
' Conforms to all vectors and can be used in digital signature or password
' protection related schemes.

Private m_lOnBits(30)   As Long
Private m_l2Power(30)   As Long
Private k(63)           As Long

Private Const BITS_TO_A_BYTE  As Long = 8
Private Const BYTES_TO_A_WORD As Long = 4
Private Const BITS_TO_A_WORD  As Long = BYTES_TO_A_WORD * BITS_TO_A_BYTE
Private Const MODULUS_BITS As Long = 512
Private Const CONGRUENT_BITS As Long = 448
Private Sub Class_Initialize()
    m_lOnBits(0) = 1
    m_lOnBits(1) = 3
    m_lOnBits(2) = 7
    m_lOnBits(3) = 15
    m_lOnBits(4) = 31
    m_lOnBits(5) = 63
    m_lOnBits(6) = 127
    m_lOnBits(7) = 255
    m_lOnBits(8) = 511
    m_lOnBits(9) = 1023
    m_lOnBits(10) = 2047
    m_lOnBits(11) = 4095
    m_lOnBits(12) = 8191
    m_lOnBits(13) = 16383
    m_lOnBits(14) = 32767
    m_lOnBits(15) = 65535
    m_lOnBits(16) = 131071
    m_lOnBits(17) = 262143
    m_lOnBits(18) = 524287
    m_lOnBits(19) = 1048575
    m_lOnBits(20) = 2097151
    m_lOnBits(21) = 4194303
    m_lOnBits(22) = 8388607
    m_lOnBits(23) = 16777215
    m_lOnBits(24) = 33554431
    m_lOnBits(25) = 67108863
    m_lOnBits(26) = 134217727
    m_lOnBits(27) = 268435455
    m_lOnBits(28) = 536870911
    m_lOnBits(29) = 1073741823
    m_lOnBits(30) = 2147483647
    m_l2Power(0) = 1
    m_l2Power(1) = 2
    m_l2Power(2) = 4
    m_l2Power(3) = 8
    m_l2Power(4) = 16
    m_l2Power(5) = 32
    m_l2Power(6) = 64
    m_l2Power(7) = 128
    m_l2Power(8) = 256
    m_l2Power(9) = 512
    m_l2Power(10) = 1024
    m_l2Power(11) = 2048
    m_l2Power(12) = 4096
    m_l2Power(13) = 8192
    m_l2Power(14) = 16384
    m_l2Power(15) = 32768
    m_l2Power(16) = 65536
    m_l2Power(17) = 131072
    m_l2Power(18) = 262144
    m_l2Power(19) = 524288
    m_l2Power(20) = 1048576
    m_l2Power(21) = 2097152
    m_l2Power(22) = 4194304
    m_l2Power(23) = 8388608
    m_l2Power(24) = 16777216
    m_l2Power(25) = 33554432
    m_l2Power(26) = 67108864
    m_l2Power(27) = 134217728
    m_l2Power(28) = 268435456
    m_l2Power(29) = 536870912
    m_l2Power(30) = 1073741824
    k(0) = &H428A2F98
    k(1) = &H71374491
    k(2) = &HB5C0FBCF
    k(3) = &HE9B5DBA5
    k(4) = &H3956C25B
    k(5) = &H59F111F1
    k(6) = &H923F82A4
    k(7) = &HAB1C5ED5
    k(8) = &HD807AA98
    k(9) = &H12835B01
    k(10) = &H243185BE
    k(11) = &H550C7DC3
    k(12) = &H72BE5D74
    k(13) = &H80DEB1FE
    k(14) = &H9BDC06A7
    k(15) = &HC19BF174
    k(16) = &HE49B69C1
    k(17) = &HEFBE4786
    k(18) = &HFC19DC6
    k(19) = &H240CA1CC
    k(20) = &H2DE92C6F
    k(21) = &H4A7484AA
    k(22) = &H5CB0A9DC
    k(23) = &H76F988DA
    k(24) = &H983E5152
    k(25) = &HA831C66D
    k(26) = &HB00327C8
    k(27) = &HBF597FC7
    k(28) = &HC6E00BF3
    k(29) = &HD5A79147
    k(30) = &H6CA6351
    k(31) = &H14292967
    k(32) = &H27B70A85
    k(33) = &H2E1B2138
    k(34) = &H4D2C6DFC
    k(35) = &H53380D13
    k(36) = &H650A7354
    k(37) = &H766A0ABB
    k(38) = &H81C2C92E
    k(39) = &H92722C85
    k(40) = &HA2BFE8A1
    k(41) = &HA81A664B
    k(42) = &HC24B8B70
    k(43) = &HC76C51A3
    k(44) = &HD192E819
    k(45) = &HD6990624
    k(46) = &HF40E3585
    k(47) = &H106AA070
    k(48) = &H19A4C116
    k(49) = &H1E376C08
    k(50) = &H2748774C
    k(51) = &H34B0BCB5
    k(52) = &H391C0CB3
    k(53) = &H4ED8AA4A
    k(54) = &H5B9CCA4F
    k(55) = &H682E6FF3
    k(56) = &H748F82EE
    k(57) = &H78A5636F
    k(58) = &H84C87814
    k(59) = &H8CC70208
    k(60) = &H90BEFFFA
    k(61) = &HA4506CEB
    k(62) = &HBEF9A3F7
    k(63) = &HC67178F2
End Sub

Private Function LShift(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then LShift = &H80000000 Else LShift = 0
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    If (lValue And m_l2Power(31 - iShiftBits)) Then LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000 Else LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
End Function
Private Function RShift(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then RShift = 1 Else RShift = 0
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
    If (lValue And &H80000000) Then RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
End Function

Private Function AddUnsigned(ByVal lX As Long, ByVal lY As Long) As Long
    Dim lX4 As Long, lY4 As Long, lX8 As Long, lY8 As Long, lResult As Long
    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000
    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8 Else lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If
    AddUnsigned = lResult
End Function
Private Function Ch(ByVal x As Long, ByVal Y As Long, ByVal z As Long) As Long
    Ch = ((x And Y) Xor ((Not x) And z))
End Function
Private Function Maj(ByVal x As Long, ByVal Y As Long, ByVal z As Long) As Long
    Maj = ((x And Y) Xor (x And z) Xor (Y And z))
End Function

Private Function S(ByVal x As Long, ByVal n As Long) As Long
    S = (RShift(x, (n And m_lOnBits(4))) Or LShift(x, (32 - (n And m_lOnBits(4)))))
End Function
Private Function R(ByVal x As Long, ByVal n As Long) As Long
    R = RShift(x, CInt(n And m_lOnBits(4)))
End Function
Private Function Sigma0(ByVal x As Long) As Long
    Sigma0 = (S(x, 2) Xor S(x, 13) Xor S(x, 22))
End Function

Private Function Sigma1(ByVal x As Long) As Long
    Sigma1 = (S(x, 6) Xor S(x, 11) Xor S(x, 25))
End Function

Private Function Gamma0(ByVal x As Long) As Long
    Gamma0 = (S(x, 7) Xor S(x, 18) Xor R(x, 3))
End Function

Private Function Gamma1(ByVal x As Long) As Long
    Gamma1 = (S(x, 17) Xor S(x, 19) Xor R(x, 10))
End Function

Private Function ConvertToWordArray(sMessage As String) As Long()
    Dim lMessageLength As Long, lNumberOfWords As Long, lWordArray() As Long, lBytePosition As Long, lByteCount As Long, lWordCount As Long, lByte As Long
    lMessageLength = Len(sMessage)
    lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
    ReDim lWordArray(lNumberOfWords - 1)
    Do Until lByteCount >= lMessageLength
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
        lByte = AscB(Mid$(sMessage, lByteCount + 1, 1))
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(lByte, lBytePosition)
        lByteCount = lByteCount + 1
    Loop
    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)
    lWordArray(lNumberOfWords - 1) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 2) = RShift(lMessageLength, 29)
    ConvertToWordArray = lWordArray
End Function
Public Function SHA256(sMessage As String) As String
    Dim HASH(7) As Long, M() As Long, w(63) As Long, A As Long, B As Long
    Dim C As Long, D As Long, E As Long, F As Long, G As Long, H As Long
    Dim i As Long, j As Long, T1 As Long, T2 As Long
    
    HASH(0) = &H6A09E667: HASH(1) = &HBB67AE85: HASH(2) = &H3C6EF372
    HASH(3) = &HA54FF53A: HASH(4) = &H510E527F: HASH(5) = &H9B05688C
    HASH(6) = &H1F83D9AB: HASH(7) = &H5BE0CD19
    
    M = ConvertToWordArray(sMessage)
    
    For i = 0 To UBound(M) Step 16
        A = HASH(0): B = HASH(1): C = HASH(2): D = HASH(3)
        E = HASH(4): F = HASH(5): G = HASH(6): H = HASH(7)
        
        For j = 0 To 63
            If j < 16 Then w(j) = M(j + i) Else w(j) = AddUnsigned(AddUnsigned(AddUnsigned(Gamma1(w(j - 2)), w(j - 7)), Gamma0(w(j - 15))), w(j - 16))
            T1 = AddUnsigned(AddUnsigned(AddUnsigned(AddUnsigned(H, Sigma1(E)), Ch(E, F, G)), k(j)), w(j))
            T2 = AddUnsigned(Sigma0(A), Maj(A, B, C))
            H = G: G = F: F = E
            E = AddUnsigned(D, T1)
            D = C: C = B: B = A
            A = AddUnsigned(T1, T2)
        Next
        
        HASH(0) = AddUnsigned(A, HASH(0))
        HASH(1) = AddUnsigned(B, HASH(1))
        HASH(2) = AddUnsigned(C, HASH(2))
        HASH(3) = AddUnsigned(D, HASH(3))
        HASH(4) = AddUnsigned(E, HASH(4))
        HASH(5) = AddUnsigned(F, HASH(5))
        HASH(6) = AddUnsigned(G, HASH(6))
        HASH(7) = AddUnsigned(H, HASH(7))
    Next
    SHA256 = LCase$(Right$("00000000" & Hex(HASH(0)), 8) & Right$("00000000" & Hex(HASH(1)), 8) & Right$("00000000" & Hex(HASH(2)), 8) & Right$("00000000" & Hex(HASH(3)), 8) & Right$("00000000" & Hex(HASH(4)), 8) & Right$("00000000" & Hex(HASH(5)), 8) & Right$("00000000" & Hex(HASH(6)), 8) & Right$("00000000" & Hex(HASH(7)), 8))
    sMessage = ""
End Function

