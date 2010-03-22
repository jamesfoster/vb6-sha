Attribute VB_Name = "modCommon"
Option Explicit

Public m_lOnBits(31)   As Long
Public m_l2Power(31)   As Long

Public Const BITS_TO_A_BYTE  As Long = 8
Public Const BYTES_TO_A_WORD As Long = 4
Public Const BITS_TO_A_WORD  As Long = BYTES_TO_A_WORD * BITS_TO_A_BYTE

Private Initialized As Boolean

Public Sub InitializeConstants()
    If Initialized Then Exit Sub

    m_lOnBits(0) = &H1&         ' 00000000000000000000000000000001
    m_lOnBits(1) = &H3&         ' 00000000000000000000000000000011
    m_lOnBits(2) = &H7&         ' 00000000000000000000000000000111
    m_lOnBits(3) = &HF&         ' 00000000000000000000000000001111
    m_lOnBits(4) = &H1F&        ' 00000000000000000000000000011111
    m_lOnBits(5) = &H3F&        ' 00000000000000000000000000111111
    m_lOnBits(6) = &H7F&        ' 00000000000000000000000001111111
    m_lOnBits(7) = &HFF&        ' 00000000000000000000000011111111
    m_lOnBits(8) = &H1FF&       ' 00000000000000000000000111111111
    m_lOnBits(9) = &H3FF&       ' 00000000000000000000001111111111
    m_lOnBits(10) = &H7FF&      ' 00000000000000000000011111111111
    m_lOnBits(11) = &HFFF&      ' 00000000000000000000111111111111
    m_lOnBits(12) = &H1FFF&     ' 00000000000000000001111111111111
    m_lOnBits(13) = &H3FFF&     ' 00000000000000000011111111111111
    m_lOnBits(14) = &H7FFF&     ' 00000000000000000111111111111111
    m_lOnBits(15) = &HFFFF&     ' 00000000000000001111111111111111
    m_lOnBits(16) = &H1FFFF     ' 00000000000000011111111111111111
    m_lOnBits(17) = &H3FFFF     ' 00000000000000111111111111111111
    m_lOnBits(18) = &H7FFFF     ' 00000000000001111111111111111111
    m_lOnBits(19) = &HFFFFF     ' 00000000000011111111111111111111
    m_lOnBits(20) = &H1FFFFF    ' 00000000000111111111111111111111
    m_lOnBits(21) = &H3FFFFF    ' 00000000001111111111111111111111
    m_lOnBits(22) = &H7FFFFF    ' 00000000011111111111111111111111
    m_lOnBits(23) = &HFFFFFF    ' 00000000111111111111111111111111
    m_lOnBits(24) = &H1FFFFFF   ' 00000001111111111111111111111111
    m_lOnBits(25) = &H3FFFFFF   ' 00000011111111111111111111111111
    m_lOnBits(26) = &H7FFFFFF   ' 00000111111111111111111111111111
    m_lOnBits(27) = &HFFFFFFF   ' 00001111111111111111111111111111
    m_lOnBits(28) = &H1FFFFFFF  ' 00011111111111111111111111111111
    m_lOnBits(29) = &H3FFFFFFF  ' 00111111111111111111111111111111
    m_lOnBits(30) = &H7FFFFFFF  ' 01111111111111111111111111111111
    m_lOnBits(31) = &HFFFFFFFF  ' 11111111111111111111111111111111

    m_l2Power(0) = &H1&         ' 00000000000000000000000000000001
    m_l2Power(1) = &H2&         ' 00000000000000000000000000000010
    m_l2Power(2) = &H4&         ' 00000000000000000000000000000100
    m_l2Power(3) = &H8&         ' 00000000000000000000000000001000
    m_l2Power(4) = &H10&        ' 00000000000000000000000000010000
    m_l2Power(5) = &H20&        ' 00000000000000000000000000100000
    m_l2Power(6) = &H40&        ' 00000000000000000000000001000000
    m_l2Power(7) = &H80&        ' 00000000000000000000000010000000
    m_l2Power(8) = &H100&       ' 00000000000000000000000100000000
    m_l2Power(9) = &H200&       ' 00000000000000000000001000000000
    m_l2Power(10) = &H400&      ' 00000000000000000000010000000000
    m_l2Power(11) = &H800&      ' 00000000000000000000100000000000
    m_l2Power(12) = &H1000&     ' 00000000000000000001000000000000
    m_l2Power(13) = &H2000&     ' 00000000000000000010000000000000
    m_l2Power(14) = &H4000&     ' 00000000000000000100000000000000
    m_l2Power(15) = &H8000&     ' 00000000000000001000000000000000
    m_l2Power(16) = &H10000     ' 00000000000000010000000000000000
    m_l2Power(17) = &H20000     ' 00000000000000100000000000000000
    m_l2Power(18) = &H40000     ' 00000000000001000000000000000000
    m_l2Power(19) = &H80000     ' 00000000000010000000000000000000
    m_l2Power(20) = &H100000    ' 00000000000100000000000000000000
    m_l2Power(21) = &H200000    ' 00000000001000000000000000000000
    m_l2Power(22) = &H400000    ' 00000000010000000000000000000000
    m_l2Power(23) = &H800000    ' 00000000100000000000000000000000
    m_l2Power(24) = &H1000000   ' 00000001000000000000000000000000
    m_l2Power(25) = &H2000000   ' 00000010000000000000000000000000
    m_l2Power(26) = &H4000000   ' 00000100000000000000000000000000
    m_l2Power(27) = &H8000000   ' 00001000000000000000000000000000
    m_l2Power(28) = &H10000000  ' 00010000000000000000000000000000
    m_l2Power(29) = &H20000000  ' 00100000000000000000000000000000
    m_l2Power(30) = &H40000000  ' 01000000000000000000000000000000
    m_l2Power(31) = &H80000000  ' 10000000000000000000000000000000

    Initialized = True
End Sub

' Print a binary number. for debug purposes
Public Function DecToBin(DeciValue As Long) As String
  Dim i As Integer

  For i = 0 To 31
      DecToBin = IIf((DeciValue And m_l2Power(i)) = 0, "0", "1") & DecToBin
  Next i
End Function


Public Function LShift(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long

    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function

    ElseIf iShiftBits = 32 Then
        LShift = 0
        Exit Function

    ' A shift of 31 will result in the right most bit becoming the left most
    ' bit and all other bits being cleared
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If
        Exit Function

    ' A shift of less than zero or more than 32 is undefined
    ElseIf iShiftBits < 0 Or iShiftBits > 32 Then
        Err.Raise 6
    End If

    ' If the left most bit that remains will end up in the negative bit
    ' position (&H80000000) we would end up with an overflow if we took the
    ' standard route. We need to strip the left most bit and add it back
    ' afterwards.
    If (lValue And m_l2Power(31 - iShiftBits)) Then

        ' (Value And OnBits(31 - (Shift + 1))) chops off the left most bits that
        ' we are shifting into, but also the left most bit we still want as this
        ' is going to end up in the negative bit marker position (&H80000000).
        ' After the multiplication/shift we Or the result with &H80000000 to
        ' turn the negative bit on.
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000

    Else

        ' (Value And OnBits(31-Shift)) chops off the left most bits that we are
        ' shifting into so we do not get an overflow error when we do the
        ' multiplication/shift
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))

    End If
End Function


Public Function RShift(ByVal lValue As Long, ByVal iShiftBits As Integer) As Long

    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function

    ElseIf iShiftBits = 32 Then
        RShift = 0
        Exit Function

    ' A shift of 31 will clear all bits and move the left most bit to the right
    ' most bit position
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If
        Exit Function

    ' A shift of less than zero or more than 32 is undefined
    ElseIf iShiftBits < 0 Or iShiftBits > 32 Then
        Err.Raise 6
    End If

    ' ingore the sign bit (&H80000000) and perform integer division
    RShift = (lValue And &H7FFFFFFF) \ m_l2Power(iShiftBits)

    ' If the sign bit (&H80000000) was set we need to add it back correctly shifted
    If (lValue And &H80000000) Then
        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
    End If
End Function


Public Function LeftRotate32(ByVal x As Long, ByVal n As Long) As Long

    LeftRotate32 = RightRotate32(x, 32 - n)

End Function

Public Function RightRotate32(ByVal x As Long, ByVal n As Long) As Long

    RightRotate32 = RShift(x, (n And m_lOnBits(4))) Or LShift(x, 32 - (n And m_lOnBits(4)))

End Function


Public Function RightShift32(ByVal x As Long, ByVal n As Long) As Long

    RightShift32 = RShift(x, CInt(n And m_lOnBits(4)))

End Function

' Adds two 32bit unsigned numbers without overflowing
Public Function Add32(ByVal lX As Long, ByVal lY As Long) As Long
    Dim lX4     As Long
    Dim lY4     As Long
    Dim lX8     As Long
    Dim lY8     As Long
    Dim lResult As Long

    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000

    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)

    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If

    Add32 = lResult
End Function
