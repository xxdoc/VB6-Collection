VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBase64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const clOneMask = 16515072          '000000 111111 111111 111111
Private Const clTwoMask = 258048            '111111 000000 111111 111111
Private Const clThreeMask = 4032            '111111 111111 000000 111111
Private Const clFourMask = 63               '111111 111111 111111 000000

Private Const clHighMask = 16711680         '11111111 00000000 00000000
Private Const clMidMask = 65280             '00000000 11111111 00000000
Private Const clLowMask = 255               '00000000 00000000 11111111

Private Const cl2Exp18 = 262144             '2 to the 18th power
Private Const cl2Exp12 = 4096               '2 to the 12th
Private Const cl2Exp6 = 64                  '2 to the 6th
Private Const cl2Exp8 = 256                 '2 to the 8th
Private Const cl2Exp16 = 65536              '2 to the 16th

Private cbTransTo(63) As Byte
Private cbTransFrom(255) As Byte
Private clPowers8(255) As Long
Private clPowers16(255) As Long
Private clPowers6(63) As Long
Private clPowers12(63) As Long
Private clPowers18(63) As Long

Private Sub Class_Initialize()

    Dim lTemp As Long

    For lTemp = 0 To 63                             'Fill the translation table.
        Select Case lTemp
            Case 0 To 25
                cbTransTo(lTemp) = 65 + lTemp       'A - Z
            Case 26 To 51
                cbTransTo(lTemp) = 71 + lTemp       'a - z
            Case 52 To 61
                cbTransTo(lTemp) = lTemp - 4        '1 - 0
            Case 62
                cbTransTo(lTemp) = 43               'Chr(43) = "+"
            Case 63
                cbTransTo(lTemp) = 47               'Chr(47) = "/"
        End Select
    Next lTemp

    For lTemp = 0 To 255                            'Fill the lookup tables.
        clPowers8(lTemp) = lTemp * cl2Exp8
        clPowers16(lTemp) = lTemp * cl2Exp16
    Next lTemp
    
    For lTemp = 0 To 63
        clPowers6(lTemp) = lTemp * cl2Exp6
        clPowers12(lTemp) = lTemp * cl2Exp12
        clPowers18(lTemp) = lTemp * cl2Exp18
    Next lTemp

    For lTemp = 0 To 255                            'Fill the translation table.
        Select Case lTemp
            Case 65 To 90
                cbTransFrom(lTemp) = lTemp - 65     'A - Z
            Case 97 To 122
                cbTransFrom(lTemp) = lTemp - 71     'a - z
            Case 48 To 57
                cbTransFrom(lTemp) = lTemp + 4      '1 - 0
            Case 43
                cbTransFrom(lTemp) = 62             'Chr(43) = "+"
            Case 47
                cbTransFrom(lTemp) = 63             'Chr(47) = "/"
        End Select
    Next lTemp

End Sub

Public Function Encode(sString As String) As String

    Dim bTrans(63) As Byte, bOut() As Byte, bIn() As Byte, lOutSize As Long
    Dim lChar As Long, lTrip As Long, iPad As Integer, lLen As Long, lTemp As Long, lPos As Long


    iPad = Len(sString) Mod 3                           'See if the length is divisible by 3
    If iPad Then                                        'If not, figure out the end pad and resize the input.
        iPad = 3 - iPad
        sString = sString & String(iPad, Chr(0))
    End If
    
    bIn = StrConv(sString, vbFromUnicode)               'Load the input string.
    lLen = ((UBound(bIn) + 1) \ 3) * 4                  'Length of resulting string.
    lTemp = lLen \ 72                                   'Added space for vbCrLfs.
    lOutSize = ((lTemp * 2) + lLen) - 1                 'Calculate the size of the output buffer.
    ReDim bOut(lOutSize)                                'Make the output buffer.
    
    lLen = 0                                            'Reusing this one, so reset it.
    
    For lChar = LBound(bIn) To UBound(bIn) Step 3
        lTrip = clPowers16(bIn(lChar)) + clPowers8(bIn(lChar + 1)) + bIn(lChar + 2)    'Combine the 3 bytes
        lTemp = lTrip And clOneMask                     'Mask for the first 6 bits
        bOut(lPos) = cbTransTo(lTemp \ cl2Exp18)        'Shift it down to the low 6 bits and get the value
        lTemp = lTrip And clTwoMask                     'Mask for the second set.
        bOut(lPos + 1) = cbTransTo(lTemp \ cl2Exp12)    'Shift it down and translate.
        lTemp = lTrip And clThreeMask                   'Mask for the third set.
        bOut(lPos + 2) = cbTransTo(lTemp \ cl2Exp6)     'Shift it down and translate.
        bOut(lPos + 3) = cbTransTo(lTrip And clFourMask) 'Mask for the low set.
        If lLen = 68 Then                               'Ready for a newline
            bOut(lPos + 4) = 13                         'Chr(13) = vbCr
            bOut(lPos + 5) = 10                         'Chr(10) = vbLf
            lLen = 0                                    'Reset the counter
            lPos = lPos + 6
        Else
            lLen = lLen + 4
            lPos = lPos + 4
        End If
    Next lChar
    
    If bOut(lOutSize) = 10 Then lOutSize = lOutSize - 2 'Shift the padding chars down if it ends with CrLf.
    
    If iPad = 1 Then                                    'Add the padding chars if any.
        bOut(lOutSize) = 61                             'Chr(61) = "="
    ElseIf iPad = 2 Then
        bOut(lOutSize) = 61
        bOut(lOutSize - 1) = 61
    End If
    
    Encode = StrConv(bOut, vbUnicode)                   'Convert back to a string and return it.
    
End Function

Public Function Decode(sString As String) As String

    Dim bOut() As Byte, bIn() As Byte, lQuad As Long, iPad As Integer, lChar As Long, lPos As Long, sOut As String
    Dim lTemp As Long

    sString = Replace(sString, vbCr, vbNullString)      'Get rid of the vbCrLfs.  These could be in...
    sString = Replace(sString, vbLf, vbNullString)      'either order.

    lTemp = Len(sString) Mod 4                          'Test for valid input.
    If lTemp Then
        Call Err.Raise(vbObjectError, "MyDecode", "Input string is not valid Base64.")
    End If
    
    If InStrRev(sString, "==") Then                     'InStrRev is faster when you know it's at the end.
        iPad = 2                                        'Note:  These translate to 0, so you can leave them...
    ElseIf InStrRev(sString, "=") Then                  'in the string and just resize the output.
        iPad = 1
    End If

    bIn = StrConv(sString, vbFromUnicode)               'Load the input byte array.
    ReDim bOut((((UBound(bIn) + 1) \ 4) * 3) - 1)       'Prepare the output buffer.
    
    For lChar = 0 To UBound(bIn) Step 4
        lQuad = clPowers18(cbTransFrom(bIn(lChar))) + clPowers12(cbTransFrom(bIn(lChar + 1))) + _
                clPowers6(cbTransFrom(bIn(lChar + 2))) + cbTransFrom(bIn(lChar + 3))           'Rebuild the bits.
        lTemp = lQuad And clHighMask                    'Mask for the first byte
        bOut(lPos) = lTemp \ cl2Exp16                   'Shift it down
        lTemp = lQuad And clMidMask                     'Mask for the second byte
        bOut(lPos + 1) = lTemp \ cl2Exp8                'Shift it down
        bOut(lPos + 2) = lQuad And clLowMask            'Mask for the third byte
        lPos = lPos + 3
    Next lChar

    sOut = StrConv(bOut, vbUnicode)                     'Convert back to a string.
    If iPad Then sOut = Left$(sOut, Len(sOut) - iPad)   'Chop off any extra bytes.
    Decode = sOut

End Function


