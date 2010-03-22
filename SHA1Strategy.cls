VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SHA1Strategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISHAStrategy

Private Function ISHAStrategy_InitializeHash() As Long()
    Dim Hash(4) As Long

    Hash(0) = &H67452301
    Hash(1) = &HEFCDAB89
    Hash(2) = &H98BADCFE
    Hash(3) = &H10325476
    Hash(4) = &HC3D2E1F0

    ISHAStrategy_InitializeHash = Hash

End Function

Private Function ISHAStrategy_Expand(Words() As Long, Index As Long) As Long

    ISHAStrategy_Expand = LeftRotate32(Words(Index - 3) Xor Words(Index - 8) Xor Words(Index - 14) Xor Words(Index - 16), 1)

End Function

Private Property Get ISHAStrategy_NumRounds() As Long

    ISHAStrategy_NumRounds = 80

End Property

Private Sub ISHAStrategy_Round(lRound As Long, Hash() As Long, Value As Long)
    Dim k    As Long
    Dim f    As Long
    Dim temp As Long

    If lRound < 20 Then
        f = (Hash(1) And Hash(2)) Or ((Not Hash(1)) And Hash(3))
        k = &H5A827999
    ElseIf lRound < 40 Then
        f = Hash(1) Xor Hash(2) Xor Hash(3)
        k = &H6ED9EBA1
    ElseIf lRound < 60 Then
        f = (Hash(1) And Hash(2)) Or (Hash(1) And Hash(3)) Or (Hash(2) And Hash(3))
        k = &H8F1BBCDC
    ElseIf lRound < 80 Then
        f = Hash(1) Xor Hash(2) Xor Hash(3)
        k = &HCA62C1D6
    End If

    temp = Add32(Add32(Add32(Add32(LeftRotate32(Hash(0), 5), f), Hash(4)), k), Value)
    Hash(4) = Hash(3)
    Hash(3) = Hash(2)
    Hash(2) = LeftRotate32(Hash(1), 30)
    Hash(1) = Hash(0)
    Hash(0) = temp

End Sub

Private Function ISHAStrategy_Output(Hash() As Long) As String

    ISHAStrategy_Output = LCase(Right("00000000" & Hex(Hash(0)), 8) & _
                                Right("00000000" & Hex(Hash(1)), 8) & _
                                Right("00000000" & Hex(Hash(2)), 8) & _
                                Right("00000000" & Hex(Hash(3)), 8) & _
                                Right("00000000" & Hex(Hash(4)), 8))

End Function
