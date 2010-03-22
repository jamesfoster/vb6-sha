VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SHA224Strategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISHAStrategy

Private inner As ISHAStrategy

Private Sub Class_Initialize()

    ' The only differences between SHA-224 and SHA-256 are the initial
    ' hash values and the output
    Set inner = New SHA256Strategy

End Sub


Private Function ISHAStrategy_InitializeHash() As Long()
    Dim Hash(7) As Long

    Hash(0) = &HC1059ED8
    Hash(1) = &H367CD507
    Hash(2) = &H3070DD17
    Hash(3) = &HF70E5939
    Hash(4) = &HFFC00B31
    Hash(5) = &H68581511
    Hash(6) = &H64F98FA7
    Hash(7) = &HBEFA4FA4

    ISHAStrategy_InitializeHash = Hash

End Function

Private Function ISHAStrategy_Expand(Words() As Long, Index As Long) As Long

    ISHAStrategy_Expand = inner.Expand(Words, Index)

End Function

Private Property Get ISHAStrategy_NumRounds() As Long

    ISHAStrategy_NumRounds = inner.NumRounds

End Property

Private Sub ISHAStrategy_Round(lRound As Long, Hash() As Long, Value As Long)

    inner.Round lRound, Hash, Value

End Sub

Private Function ISHAStrategy_Output(Hash() As Long) As String

    ISHAStrategy_Output = LCase(Right("00000000" & Hex(Hash(0)), 8) & _
                                Right("00000000" & Hex(Hash(1)), 8) & _
                                Right("00000000" & Hex(Hash(2)), 8) & _
                                Right("00000000" & Hex(Hash(3)), 8) & _
                                Right("00000000" & Hex(Hash(4)), 8) & _
                                Right("00000000" & Hex(Hash(5)), 8) & _
                                Right("00000000" & Hex(Hash(6)), 8))

End Function
