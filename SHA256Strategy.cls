VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SHA256Strategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISHAStrategy

Private k(63)   As Long

Private Sub Class_Initialize()
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


Private Function Ch(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Long
    Ch = ((x And y) Xor ((Not x) And z))
End Function


Private Function Maj(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Long
    Maj = ((x And y) Xor (x And z) Xor (y And z))
End Function


Private Function Sigma0(ByVal x As Long) As Long
    Sigma0 = (RightRotate32(x, 2) Xor RightRotate32(x, 13) Xor RightRotate32(x, 22))
End Function


Private Function Sigma1(ByVal x As Long) As Long
    Sigma1 = (RightRotate32(x, 6) Xor RightRotate32(x, 11) Xor RightRotate32(x, 25))
End Function


Private Function Gamma0(ByVal x As Long) As Long
    Gamma0 = (RightRotate32(x, 7) Xor RightRotate32(x, 18) Xor RightShift32(x, 3))
End Function


Private Function Gamma1(ByVal x As Long) As Long
    Gamma1 = (RightRotate32(x, 17) Xor RightRotate32(x, 19) Xor RightShift32(x, 10))
End Function

Private Function ISHAStrategy_InitializeHash() As Long()
    Dim Hash(7) As Long

    Hash(0) = &H6A09E667
    Hash(1) = &HBB67AE85
    Hash(2) = &H3C6EF372
    Hash(3) = &HA54FF53A
    Hash(4) = &H510E527F
    Hash(5) = &H9B05688C
    Hash(6) = &H1F83D9AB
    Hash(7) = &H5BE0CD19

    ISHAStrategy_InitializeHash = Hash

End Function

Private Function ISHAStrategy_Expand(Words() As Long, Index As Long) As Long

    ISHAStrategy_Expand = Add32(Add32(Add32(Gamma1(Words(Index - 2)), Words(Index - 7)), Gamma0(Words(Index - 15))), Words(Index - 16))

End Function

Private Property Get ISHAStrategy_NumRounds() As Long

    ISHAStrategy_NumRounds = 64

End Property

Private Sub ISHAStrategy_Round(lRound As Long, Hash() As Long, Value As Long)
    Dim T1 As Long
    Dim T2 As Long

    T1 = Add32(Add32(Add32(Add32(Hash(7), Sigma1(Hash(4))), Ch(Hash(4), Hash(5), Hash(6))), k(lRound)), Value)
    T2 = Add32(Sigma0(Hash(0)), Maj(Hash(0), Hash(1), Hash(2)))

    Hash(7) = Hash(6)
    Hash(6) = Hash(5)
    Hash(5) = Hash(4)
    Hash(4) = Add32(Hash(3), T1)
    Hash(3) = Hash(2)
    Hash(2) = Hash(1)
    Hash(1) = Hash(0)
    Hash(0) = Add32(T1, T2)

End Sub

Private Function ISHAStrategy_Output(Hash() As Long) As String

    ISHAStrategy_Output = LCase(Right("00000000" & Hex(Hash(0)), 8) & _
                                Right("00000000" & Hex(Hash(1)), 8) & _
                                Right("00000000" & Hex(Hash(2)), 8) & _
                                Right("00000000" & Hex(Hash(3)), 8) & _
                                Right("00000000" & Hex(Hash(4)), 8) & _
                                Right("00000000" & Hex(Hash(5)), 8) & _
                                Right("00000000" & Hex(Hash(6)), 8) & _
                                Right("00000000" & Hex(Hash(7)), 8))

End Function


