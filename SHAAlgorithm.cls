VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SHAAlgorithm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mSHA1 As SHA1Strategy
Private mSHA256 As SHA256Strategy
Private mSHA224 As SHA224Strategy

Private Sub Class_Initialize()
    InitializeConstants
End Sub

Public Function SHA1(Message() As Byte) As String

    If mSHA1 Is Nothing Then
        Set mSHA1 = New SHA1Strategy
    End If

    SHA1 = Run(Message, mSHA1)

End Function

Public Function SHA1FromString(Message As String) As String

    SHA1FromString = SHA1(StrConv(Message, vbFromUnicode))

End Function

Public Function SHA256(Message() As Byte) As String

    If mSHA256 Is Nothing Then
        Set mSHA256 = New SHA256Strategy
    End If

    SHA256 = Run(Message, mSHA256)

End Function

Public Function SHA256FromString(Message As String) As String

    SHA256FromString = SHA256(StrConv(Message, vbFromUnicode))

End Function

Public Function SHA224(Message() As Byte) As String

    If mSHA224 Is Nothing Then
        Set mSHA224 = New SHA224Strategy
    End If

    SHA224 = Run(Message, mSHA224)

End Function

Public Function SHA224FromString(Message As String) As String

    SHA224FromString = SHA224(StrConv(Message, vbFromUnicode))

End Function

Public Function Run(Message() As Byte, strategy As ISHAStrategy) As String
    Dim RoundHash() As Long
    Dim Hash()      As Long
    Dim Msg()       As Long
    Dim Words()     As Long
    Dim lChunk      As Long
    Dim lRound      As Long
    Dim i           As Long
    Dim NumRounds   As Long
    Dim HashSize    As Long

    Hash = strategy.InitializeHash()
    HashSize = UBound(Hash)
    ReDim RoundHash(HashSize)

    Msg = ConvertToWordArray(Message)

    NumRounds = strategy.NumRounds
    ReDim Words(NumRounds - 1)

    For lChunk = 0 To UBound(Msg) Step 16
        For i = 0 To HashSize
            RoundHash(i) = Hash(i)
        Next

        For lRound = 0 To NumRounds - 1
            If lRound < 16 Then
                Words(lRound) = Msg(lRound + lChunk)
            Else
                Words(lRound) = strategy.Expand(Words, lRound)
            End If

            strategy.Round lRound, RoundHash, Words(lRound)
        Next

        For i = 0 To HashSize
            Hash(i) = Add32(Hash(i), RoundHash(i))
        Next
    Next

    Run = strategy.Output(Hash)

End Function

Public Function ConvertToWordArray(sMessage() As Byte) As Long()

    Dim lMessageLength  As Long
    Dim lNumberOfWords  As Long
    Dim lWordArray()    As Long
    Dim lBytePosition   As Long
    Dim lByteCount      As Long
    Dim lWordCount      As Long
    Dim lByte           As Long

    Const MODULUS_BITS      As Long = 512
    Const CONGRUENT_BITS    As Long = 448

    Const DIFF_BYTES        As Long = (MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE
    Const MODULUS_BYTES     As Long = MODULUS_BITS \ BITS_TO_A_BYTE
    Const MODULUS_WORDS     As Long = MODULUS_BITS \ BITS_TO_A_WORD

    lMessageLength = UBound(sMessage) + 1

    ' Get padded number of words. Message needs to be congruent to 448 bits,
    ' modulo 512 bits. If it is exactly congruent to 448 bits, modulo 512 bits
    ' it must still have another 512 bits added. 512 bits = 64 bytes
    ' (or 16 * 4 byte words), 448 bits = 56 bytes. This means lNumberOfWords must
    ' be a multiple of 16 (i.e. 16 * 4 (bytes) * 8 (bits))

    lNumberOfWords = (((lMessageLength + DIFF_BYTES) \ MODULUS_BYTES) + 1) * MODULUS_WORDS
    ReDim lWordArray(lNumberOfWords - 1)

    ' Combine each block of 4 bytes (ascii code of character) into one long
    ' value and store in the message. The high-order (most significant) bit of
    ' each byte is listed first. However, unlike MD5 we put the high-order
    ' (most significant) byte first in each word.
    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= lMessageLength
        ' Each word is 4 bytes
        lWordCount = lByteCount \ BYTES_TO_A_WORD

        lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE

        lByte = sMessage(lByteCount)

        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(lByte, lBytePosition)
        lByteCount = lByteCount + 1
    Loop

    ' Terminate according to SHA-256 rules with a 1 bit, zeros and the length in
    ' bits stored in the last two words
    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE

    ' Add a terminating 1 bit, all the rest of the bits to the end of the
    ' word array will default to zero
    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

    ' We put the length of the message in bits into the last two words, to get
    ' the length in bits we need to multiply by 8 (or left shift 3). This left
    ' shifted value is put in the last word. Any bits shifted off the left edge
    ' need to be put in the penultimate word, we can work out which bits by shifting
    ' right the length by 29 bits.
    lWordArray(lNumberOfWords - 1) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 2) = RShift(lMessageLength, 29)

    ConvertToWordArray = lWordArray
End Function

