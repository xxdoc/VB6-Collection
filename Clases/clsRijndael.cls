VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRijndael"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'   Credits :      Phil Fresle
Option Explicit
Private m_lOnBits(30)    As Long
Private m_l2Power(30)    As Long
Private m_bytOnBits(7)   As Byte
Private m_byt2Power(7)   As Byte
Private m_InCo(3)        As Byte
Private m_fbsub(255)     As Byte
Private m_rbsub(255)     As Byte
Private m_ptab(255)      As Byte
Private m_ltab(255)      As Byte
Private m_ftable(255)    As Long
Private m_rtable(255)    As Long
Private m_rco(29)        As Long
Private m_Nk             As Long
Private m_Nb             As Long
Private m_Nr             As Long
Private m_fi(23)         As Byte
Private m_ri(23)         As Byte
Private m_fkey(119)      As Long
Private m_rkey(119)      As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Any, ByVal Source As Any, ByVal Length As Long)
Private Sub Class_Initialize()
    m_InCo(0) = &HB:    m_InCo(1) = &HD
    m_InCo(2) = &H9:    m_InCo(3) = &HE
    m_bytOnBits(0) = 1          ' 00000001
    m_bytOnBits(1) = 3          ' 00000011
    m_bytOnBits(2) = 7          ' 00000111
    m_bytOnBits(3) = 15         ' 00001111
    m_bytOnBits(4) = 31         ' 00011111
    m_bytOnBits(5) = 63         ' 00111111
    m_bytOnBits(6) = 127        ' 01111111
    m_bytOnBits(7) = 255        ' 11111111
    m_byt2Power(0) = 1          ' 00000001
    m_byt2Power(1) = 2          ' 00000010
    m_byt2Power(2) = 4          ' 00000100
    m_byt2Power(3) = 8          ' 00001000
    m_byt2Power(4) = 16         ' 00010000
    m_byt2Power(5) = 32         ' 00100000
    m_byt2Power(6) = 64         ' 01000000
    m_byt2Power(7) = 128        ' 10000000
    m_lOnBits(0) = 1            ' 00000000000000000000000000000001
    m_lOnBits(1) = 3            ' 00000000000000000000000000000011
    m_lOnBits(2) = 7            ' 00000000000000000000000000000111
    m_lOnBits(3) = 15           ' 00000000000000000000000000001111
    m_lOnBits(4) = 31           ' 00000000000000000000000000011111
    m_lOnBits(5) = 63           ' 00000000000000000000000000111111
    m_lOnBits(6) = 127          ' 00000000000000000000000001111111
    m_lOnBits(7) = 255          ' 00000000000000000000000011111111
    m_lOnBits(8) = 511          ' 00000000000000000000000111111111
    m_lOnBits(9) = 1023         ' 00000000000000000000001111111111
    m_lOnBits(10) = 2047        ' 00000000000000000000011111111111
    m_lOnBits(11) = 4095        ' 00000000000000000000111111111111
    m_lOnBits(12) = 8191        ' 00000000000000000001111111111111
    m_lOnBits(13) = 16383       ' 00000000000000000011111111111111
    m_lOnBits(14) = 32767       ' 00000000000000000111111111111111
    m_lOnBits(15) = 65535       ' 00000000000000001111111111111111
    m_lOnBits(16) = 131071      ' 00000000000000011111111111111111
    m_lOnBits(17) = 262143      ' 00000000000000111111111111111111
    m_lOnBits(18) = 524287      ' 00000000000001111111111111111111
    m_lOnBits(19) = 1048575     ' 00000000000011111111111111111111
    m_lOnBits(20) = 2097151     ' 00000000000111111111111111111111
    m_lOnBits(21) = 4194303     ' 00000000001111111111111111111111
    m_lOnBits(22) = 8388607     ' 00000000011111111111111111111111
    m_lOnBits(23) = 16777215    ' 00000000111111111111111111111111
    m_lOnBits(24) = 33554431    ' 00000001111111111111111111111111
    m_lOnBits(25) = 67108863    ' 00000011111111111111111111111111
    m_lOnBits(26) = 134217727   ' 00000111111111111111111111111111
    m_lOnBits(27) = 268435455   ' 00001111111111111111111111111111
    m_lOnBits(28) = 536870911   ' 00011111111111111111111111111111
    m_lOnBits(29) = 1073741823  ' 00111111111111111111111111111111
    m_lOnBits(30) = 2147483647  ' 01111111111111111111111111111111
    m_l2Power(0) = 1            ' 00000000000000000000000000000001
    m_l2Power(1) = 2            ' 00000000000000000000000000000010
    m_l2Power(2) = 4            ' 00000000000000000000000000000100
    m_l2Power(3) = 8            ' 00000000000000000000000000001000
    m_l2Power(4) = 16           ' 00000000000000000000000000010000
    m_l2Power(5) = 32           ' 00000000000000000000000000100000
    m_l2Power(6) = 64           ' 00000000000000000000000001000000
    m_l2Power(7) = 128          ' 00000000000000000000000010000000
    m_l2Power(8) = 256          ' 00000000000000000000000100000000
    m_l2Power(9) = 512          ' 00000000000000000000001000000000
    m_l2Power(10) = 1024        ' 00000000000000000000010000000000
    m_l2Power(11) = 2048        ' 00000000000000000000100000000000
    m_l2Power(12) = 4096        ' 00000000000000000001000000000000
    m_l2Power(13) = 8192        ' 00000000000000000010000000000000
    m_l2Power(14) = 16384       ' 00000000000000000100000000000000
    m_l2Power(15) = 32768       ' 00000000000000001000000000000000
    m_l2Power(16) = 65536       ' 00000000000000010000000000000000
    m_l2Power(17) = 131072      ' 00000000000000100000000000000000
    m_l2Power(18) = 262144      ' 00000000000001000000000000000000
    m_l2Power(19) = 524288      ' 00000000000010000000000000000000
    m_l2Power(20) = 1048576     ' 00000000000100000000000000000000
    m_l2Power(21) = 2097152     ' 00000000001000000000000000000000
    m_l2Power(22) = 4194304     ' 00000000010000000000000000000000
    m_l2Power(23) = 8388608     ' 00000000100000000000000000000000
    m_l2Power(24) = 16777216    ' 00000001000000000000000000000000
    m_l2Power(25) = 33554432    ' 00000010000000000000000000000000
    m_l2Power(26) = 67108864    ' 00000100000000000000000000000000
    m_l2Power(27) = 134217728   ' 00001000000000000000000000000000
    m_l2Power(28) = 268435456   ' 00010000000000000000000000000000
    m_l2Power(29) = 536870912   ' 00100000000000000000000000000000
    m_l2Power(30) = 1073741824  ' 01000000000000000000000000000000
End Sub
Private Function LShift(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
    If iShiftBits = 0 Then
        LShift = lValue:        Exit Function
    ElseIf iShiftBits = 31 Then
        LShift = IIf(lValue And 1, &H80000000, 0)
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    If (lValue And m_l2Power(31 - iShiftBits)) Then
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
    Else
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
    End If
End Function
Private Function RShift(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
    If iShiftBits = 0 Then
        RShift = lValue:        Exit Function
    ElseIf iShiftBits = 31 Then
        RShift = IIf(lValue And &H80000000, 1, 0)
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
    If (lValue And &H80000000) Then RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
End Function
Private Function LShiftByte(ByVal bytValue As Byte, ByVal bytShiftBits As Byte) As Byte
    If bytShiftBits = 0 Then
        LShiftByte = bytValue:        Exit Function
    ElseIf bytShiftBits = 7 Then
        LShiftByte = IIf(bytValue And 1, &H80, 0)
        Exit Function
    ElseIf bytShiftBits < 0 Or bytShiftBits > 7 Then
        Err.Raise 6
    End If
    LShiftByte = ((bytValue And m_bytOnBits(7 - bytShiftBits)) * m_byt2Power(bytShiftBits))
End Function
Private Function RShiftByte(ByVal bytValue As Byte, ByVal bytShiftBits As Byte) As Byte
    If bytShiftBits = 0 Then
        RShiftByte = bytValue:        Exit Function
    ElseIf bytShiftBits = 7 Then
        RShiftByte = IIf(bytValue And &H80, 1, 0)
        Exit Function
    ElseIf bytShiftBits < 0 Or bytShiftBits > 7 Then
        Err.Raise 6
    End If
    RShiftByte = bytValue \ m_byt2Power(bytShiftBits)
End Function
Private Function RotateLeft(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long
    RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function
Private Function RotateLeftByte(ByVal bytValue As Byte, ByVal bytShiftBits As Byte) As Byte
    RotateLeftByte = LShiftByte(bytValue, bytShiftBits) Or RShiftByte(bytValue, (8 - bytShiftBits))
End Function
Private Function Pack(B() As Byte) As Long
Dim lCount As Long, lTemp  As Long
    For lCount = 0 To 3
        lTemp = B(lCount): Pack = Pack Or LShift(lTemp, (lCount * 8))
    Next lCount
End Function
Private Function PackFrom(B() As Byte, ByVal k As Long) As Long
Dim lCount As Long, lTemp  As Long
    For lCount = 0 To 3
        lTemp = B(lCount + k): PackFrom = PackFrom Or LShift(lTemp, (lCount * 8))
    Next lCount
End Function
Private Sub Unpack(ByVal a As Long, B() As Byte)
    B(0) = a And m_lOnBits(7)
    B(1) = RShift(a, 8) And m_lOnBits(7)
    B(2) = RShift(a, 16) And m_lOnBits(7)
    B(3) = RShift(a, 24) And m_lOnBits(7)
End Sub
Private Sub UnpackFrom(ByVal a As Long, B() As Byte, ByVal k As Long)
    B(0 + k) = a And m_lOnBits(7)
    B(1 + k) = RShift(a, 8) And m_lOnBits(7)
    B(2 + k) = RShift(a, 16) And m_lOnBits(7)
    B(3 + k) = RShift(a, 24) And m_lOnBits(7)
End Sub
Private Function Xtime(ByVal a As Byte) As Byte
Dim B As Byte:    B = IIf(a And &H80, &H1B, 0)
    a = LShiftByte(a, 1):   a = a Xor B:   Xtime = a
End Function
Private Function Bmul(ByVal X As Byte, Y As Byte) As Byte
    If X <> 0 And Y <> 0 Then
        Bmul = m_ptab((CLng(m_ltab(X)) + CLng(m_ltab(Y))) Mod 255)
    Else
        Bmul = 0
    End If
End Function
Private Function SubByte(ByVal a As Long) As Long
Dim B(3) As Byte
    Unpack a, B
    B(0) = m_fbsub(B(0)):    B(1) = m_fbsub(B(1))
    B(2) = m_fbsub(B(2)):    B(3) = m_fbsub(B(3))
    SubByte = Pack(B)
End Function
Private Function Product(ByVal X As Long, ByVal Y As Long) As Long
Dim xb(3) As Byte, yb(3) As Byte
    Unpack X, xb:    Unpack Y, yb
    Product = Bmul(xb(0), yb(0)) Xor Bmul(xb(1), yb(1)) Xor Bmul(xb(2), yb(2)) Xor Bmul(xb(3), yb(3))
End Function
Private Function InvMixCol(ByVal X As Long) As Long
Dim Y       As Long, M       As Long, B(3)    As Byte
    M = Pack(m_InCo)
    B(3) = Product(M, X):    M = RotateLeft(M, 24)
    B(2) = Product(M, X):    M = RotateLeft(M, 24)
    B(1) = Product(M, X):    M = RotateLeft(M, 24)
    B(0) = Product(M, X):    Y = Pack(B)
    InvMixCol = Y
End Function
Private Function ByteSub(ByVal X As Byte) As Byte
Dim Y As Byte
    Y = m_ptab(255 - m_ltab(X)):    X = Y
    X = RotateLeftByte(X, 1):       Y = Y Xor X
    X = RotateLeftByte(X, 1):       Y = Y Xor X
    X = RotateLeftByte(X, 1):       Y = Y Xor X
    X = RotateLeftByte(X, 1):       Y = Y Xor X
    Y = Y Xor &H63:              ByteSub = Y
End Function
Private Sub gentables()
Dim i    As Long, Y As Byte
Dim B(3) As Byte, ib As Byte
    m_ltab(0) = 0:    m_ptab(0) = 1
    m_ltab(1) = 0:    m_ptab(1) = 3
    m_ltab(3) = 1
    For i = 2 To 255
        m_ptab(i) = m_ptab(i - 1) Xor Xtime(m_ptab(i - 1))
        m_ltab(m_ptab(i)) = i
    Next i
    m_fbsub(0) = &H63:    m_rbsub(&H63) = 0
    For i = 1 To 255
        ib = i:        Y = ByteSub(ib)
        m_fbsub(i) = Y:        m_rbsub(Y) = i
    Next i
    Y = 1
    For i = 0 To 29
        m_rco(i) = Y:        Y = Xtime(Y)
    Next i
    For i = 0 To 255
        Y = m_fbsub(i):        B(3) = Y Xor Xtime(Y)
        B(2) = Y:               B(1) = Y
        B(0) = Xtime(Y):        m_ftable(i) = Pack(B)
        Y = m_rbsub(i):        B(3) = Bmul(m_InCo(0), Y)
        B(2) = Bmul(m_InCo(1), Y): B(1) = Bmul(m_InCo(2), Y)
        B(0) = Bmul(m_InCo(3), Y): m_rtable(i) = Pack(B)
    Next i
End Sub
Private Sub gkey(ByVal nb As Long, ByVal nk As Long, Key() As Byte)
Dim i            As Long, j As Long, k As Long, M As Long
Dim n            As Long, C1 As Long, C2 As Long, C3 As Long
Dim CipherKey(7) As Long
    m_Nb = nb:    m_Nk = nk
    If m_Nb >= m_Nk Then
        m_Nr = 6 + m_Nb
    Else
        m_Nr = 6 + m_Nk
    End If
    C1 = 1
    If m_Nb < 8 Then
        C2 = 2:        C3 = 3
    Else
        C2 = 3:        C3 = 4
    End If
    For j = 0 To nb - 1
        M = j * 3
        m_fi(M) = (j + C1) Mod nb
        m_fi(M + 1) = (j + C2) Mod nb
        m_fi(M + 2) = (j + C3) Mod nb
        m_ri(M) = (nb + j - C1) Mod nb
        m_ri(M + 1) = (nb + j - C2) Mod nb
        m_ri(M + 2) = (nb + j - C3) Mod nb
    Next j
    n = m_Nb * (m_Nr + 1)
    For i = 0 To m_Nk - 1
        j = i * 4:        CipherKey(i) = PackFrom(Key, j)
    Next i
    For i = 0 To m_Nk - 1
        m_fkey(i) = CipherKey(i)
    Next i
    j = m_Nk:    k = 0
    Do While j < n
        m_fkey(j) = m_fkey(j - m_Nk) Xor SubByte(RotateLeft(m_fkey(j - 1), 24)) Xor m_rco(k)
        If m_Nk <= 6 Then
            i = 1
            Do While i < m_Nk And (i + j) < n
                m_fkey(i + j) = m_fkey(i + j - m_Nk) Xor m_fkey(i + j - 1): i = i + 1
            Loop
        Else
            i = 1
            Do While i < 4 And (i + j) < n
                m_fkey(i + j) = m_fkey(i + j - m_Nk) Xor m_fkey(i + j - 1): i = i + 1
            Loop
            If j + 4 < n Then
                m_fkey(j + 4) = m_fkey(j + 4 - m_Nk) Xor SubByte(m_fkey(j + 3))
            End If
            i = 5
            Do While i < m_Nk And (i + j) < n
                m_fkey(i + j) = m_fkey(i + j - m_Nk) Xor m_fkey(i + j - 1): i = i + 1
            Loop
        End If
        j = j + m_Nk:        k = k + 1
    Loop
    For j = 0 To m_Nb - 1
        m_rkey(j + n - nb) = m_fkey(j)
    Next j
    i = m_Nb
    Do While i < n - m_Nb
        k = n - m_Nb - i
        For j = 0 To m_Nb - 1
            m_rkey(k + j) = InvMixCol(m_fkey(i + j))
        Next j
        i = i + m_Nb
    Loop
    j = n - m_Nb
    Do While j < n
        m_rkey(j - n + m_Nb) = m_fkey(j):        j = j + 1
    Loop
End Sub
Private Sub Encrypt(Buff() As Byte)
Dim i    As Long, j As Long, k As Long, M As Long
Dim a(7) As Long, B(7) As Long, X() As Long
Dim Y()  As Long, t() As Long
    For i = 0 To m_Nb - 1
        j = i * 4: a(i) = PackFrom(Buff, j): a(i) = a(i) Xor m_fkey(i)
    Next i
    k = m_Nb:    X = a:    Y = B
    For i = 1 To m_Nr - 1
        For j = 0 To m_Nb - 1
            M = j * 3
            Y(j) = m_fkey(k) Xor m_ftable(X(j) And m_lOnBits(7)) Xor RotateLeft(m_ftable(RShift(X(m_fi(M)), 8) And m_lOnBits(7)), 8) Xor RotateLeft(m_ftable(RShift(X(m_fi(M + 1)), 16) And m_lOnBits(7)), 16) Xor RotateLeft(m_ftable(RShift(X(m_fi(M + 2)), 24) And m_lOnBits(7)), 24)
            k = k + 1
        Next j
        t = X:        X = Y:        Y = t
    Next i
    For j = 0 To m_Nb - 1
        M = j * 3
        Y(j) = m_fkey(k) Xor m_fbsub(X(j) And m_lOnBits(7)) Xor RotateLeft(m_fbsub(RShift(X(m_fi(M)), 8) And m_lOnBits(7)), 8) Xor RotateLeft(m_fbsub(RShift(X(m_fi(M + 1)), 16) And m_lOnBits(7)), 16) Xor RotateLeft(m_fbsub(RShift(X(m_fi(M + 2)), 24) And m_lOnBits(7)), 24)
        k = k + 1
    Next j
    For i = 0 To m_Nb - 1
        j = i * 4: UnpackFrom Y(i), Buff, j: X(i) = 0: Y(i) = 0
    Next i
End Sub
Private Sub Decrypt(Buff() As Byte)
Dim i    As Long, j As Long, k As Long, M As Long
Dim a(7) As Long, B(7) As Long, X() As Long
Dim Y()  As Long, t() As Long
    For i = 0 To m_Nb - 1
        j = i * 4: a(i) = PackFrom(Buff, j): a(i) = a(i) Xor m_rkey(i)
    Next i
    k = m_Nb:    X = a:    Y = B
    For i = 1 To m_Nr - 1
        For j = 0 To m_Nb - 1
            M = j * 3
            Y(j) = m_rkey(k) Xor m_rtable(X(j) And m_lOnBits(7)) Xor RotateLeft(m_rtable(RShift(X(m_ri(M)), 8) And m_lOnBits(7)), 8) Xor RotateLeft(m_rtable(RShift(X(m_ri(M + 1)), 16) And m_lOnBits(7)), 16) Xor RotateLeft(m_rtable(RShift(X(m_ri(M + 2)), 24) And m_lOnBits(7)), 24)
            k = k + 1
        Next j
        t = X:        X = Y:        Y = t
    Next i
    For j = 0 To m_Nb - 1
        M = j * 3
        Y(j) = m_rkey(k) Xor m_rbsub(X(j) And m_lOnBits(7)) Xor RotateLeft(m_rbsub(RShift(X(m_ri(M)), 8) And m_lOnBits(7)), 8) Xor RotateLeft(m_rbsub(RShift(X(m_ri(M + 1)), 16) And m_lOnBits(7)), 16) Xor RotateLeft(m_rbsub(RShift(X(m_ri(M + 2)), 24) And m_lOnBits(7)), 24)
        k = k + 1
    Next j
    For i = 0 To m_Nb - 1
        j = i * 4: UnpackFrom Y(i), Buff, j: X(i) = 0: Y(i) = 0
    Next i
End Sub
Private Function IsInitialized(ByRef vArray As Variant) As Boolean
    On Local Error Resume Next
    IsInitialized = IsNumeric(UBound(vArray))
End Function
Public Function EncryptData(bytMessage() As Byte, bytPassword() As Byte) As Byte()
On Error Resume Next
Dim bytKey(31)     As Byte, bytIn() As Byte, bytOut() As Byte
Dim bytTemp(31)    As Byte, lCount As Long, lLength As Long
Dim lEncodedLength As Long, bytLen(3) As Byte, lPosition As Long
    If Not IsInitialized(bytMessage) Then Exit Function
    If Not IsInitialized(bytPassword) Then Exit Function
    For lCount = 0 To UBound(bytPassword)
        bytKey(lCount) = bytPassword(lCount): If lCount = 31 Then Exit For
    Next lCount
    gentables
    gkey 8, 8, bytKey
    lLength = UBound(bytMessage) + 1:    lEncodedLength = lLength + 4
    If lEncodedLength Mod 32 <> 0 Then lEncodedLength = lEncodedLength + 32 - (lEncodedLength Mod 32)
    ReDim bytIn(lEncodedLength - 1):    ReDim bytOut(lEncodedLength - 1)
    CopyMemory VarPtr(bytIn(0)), VarPtr(lLength), 4
    CopyMemory VarPtr(bytIn(4)), VarPtr(bytMessage(0)), lLength
    For lCount = 0 To lEncodedLength - 1 Step 32
        CopyMemory VarPtr(bytTemp(0)), VarPtr(bytIn(lCount)), 32
        Encrypt bytTemp
        CopyMemory VarPtr(bytOut(lCount)), VarPtr(bytTemp(0)), 32
    Next lCount
    EncryptData = bytOut
End Function
Public Function DecryptData(bytIn() As Byte, bytPassword() As Byte) As Byte()
On Error Resume Next
Dim bytMessage()   As Byte, bytKey(31) As Byte, bytOut() As Byte
Dim bytTemp(31)    As Byte, lCount As Long, lLength As Long
Dim lEncodedLength As Long, bytLen(3) As Byte, lPosition As Long
    If Not IsInitialized(bytIn) Then Exit Function
    If Not IsInitialized(bytPassword) Then Exit Function
    lEncodedLength = UBound(bytIn) + 1
    If lEncodedLength Mod 32 <> 0 Then Exit Function
    For lCount = 0 To UBound(bytPassword)
        bytKey(lCount) = bytPassword(lCount):  If lCount = 31 Then Exit For
    Next lCount
    gentables
    gkey 8, 8, bytKey
    ReDim bytOut(lEncodedLength - 1)
    For lCount = 0 To lEncodedLength - 1 Step 32
        CopyMemory VarPtr(bytTemp(0)), VarPtr(bytIn(lCount)), 32
        Decrypt bytTemp
        CopyMemory VarPtr(bytOut(lCount)), VarPtr(bytTemp(0)), 32
    Next lCount
    CopyMemory VarPtr(lLength), VarPtr(bytOut(0)), 4
    If lLength > lEncodedLength - 4 Then Exit Function
    ReDim bytMessage(lLength - 1)
    CopyMemory VarPtr(bytMessage(0)), VarPtr(bytOut(4)), lLength
    DecryptData = bytMessage
End Function

