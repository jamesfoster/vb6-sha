VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTestSHA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SHA Test"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmnDlg 
      Left            =   8760
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdTestSHA224 
      Caption         =   "Test SHA-224"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdBenchmark 
      Caption         =   "Benchmark"
      Height          =   495
      Left            =   5760
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdTestSHA1 
      Caption         =   "Test SHA-1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdTestSHA256 
      Caption         =   "Test SHA-256"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblOutput 
      Height          =   5655
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   9015
   End
End
Attribute VB_Name = "frmTestSHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oSHA As New SHAAlgorithm

Private Sub Output(ByVal str As String)

    lblOutput.Caption = lblOutput.Caption & str & vbCrLf
    DoEvents

End Sub

Private Sub ClearOutput()

    lblOutput.Caption = ""
    DoEvents

End Sub

Private Sub TestSHA256(Expected As String, Value As String)
    Dim Actual As String

    Actual = oSHA.SHA256FromString(Value)

    Output "SHA256(""" & Value & """)"
    Output "Expected: " & Expected
    Output "Actual:   " & Actual
    Output IIf(Actual = Expected, "Success", "Fail")
    Output ""

End Sub

Private Sub TestSHA224(Expected As String, Value As String)
    Dim Actual As String

    Actual = oSHA.SHA224FromString(Value)

    Output "SHA224(""" & Value & """)"
    Output "Expected: " & Expected
    Output "Actual:   " & Actual
    Output IIf(Actual = Expected, "Success", "Fail")
    Output ""

End Sub

Private Sub TestSHA1(Expected As String, Value As String)
    Dim Actual As String

    Actual = oSHA.SHA1FromString(Value)

    Output "SHA1(""" & Value & """)"
    Output "Expected: " & Expected
    Output "Actual:   " & Actual
    Output IIf(Actual = Expected, "Success", "Fail")
    Output ""

End Sub

Private Sub cmdTestSHA256_Click()

    ClearOutput

    TestSHA256 "d7a8fbb307d7809469ca9abcb0082e4f8d5651e46d3cdb762d02d0bf37c9e592", "The quick brown fox jumps over the lazy dog"
    TestSHA256 "9f86d081884c7d659a2feaa0c55ad015a3bf4f1b2b0b822cd15d6c15b0f00a08", "test"
    TestSHA256 "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", "abc"

End Sub

Private Sub cmdTestSHA1_Click()

    ClearOutput

    TestSHA1 "2fd4e1c67a2d28fced849ee1bb76e7391b93eb12", "The quick brown fox jumps over the lazy dog"
    TestSHA1 "a94a8fe5ccb19ba61c4c0873d391e987982fbbd3", "test"
    TestSHA1 "d0be2dc421be4fcd0172e5afceea3970e2f3d940", "apple"

End Sub

Private Sub cmdTestSHA224_Click()

    ClearOutput

    TestSHA224 "23097d223405d8228642a477bda255b32aadbce4bda0b3f7e36c9da7", "abc"

End Sub

Private Sub cmdBenchmark_Click()
    Dim starttime   As Double
    Dim length      As Double
    Dim total       As Double
    Dim average     As Double
    Dim fileName    As String
    Dim fileLength  As Long
    Dim bytes()     As Byte
    Dim i           As Integer
    Dim SHA         As New SHAAlgorithm

    On Error Resume Next
    cmnDlg.ShowOpen
    If Err.Number = cdlCancel Then
        ' The user canceled.
        Exit Sub
    ElseIf Err.Number <> 0 Then
        ' Unknown error.
        MsgBox "Error " & Format$(Err.Number) & " selecting file." & vbCrLf & Err.Description
        Exit Sub
    End If
    On Error GoTo 0

    fileName = cmnDlg.fileName
    fileLength = FileLen(fileName)
    ClearOutput

    Output "Input: " & fileName
    Output ""

    ReDim bytes(fileLength - 1)

    Open fileName For Binary As #1
        Get #1, 1, bytes
    Close #1


    total = 0
    Output "SHA1"
    For i = 1 To 5
        starttime = Timer

        SHA.SHA1 bytes

        length = Timer - starttime

        Output i & ": " & length & "s - " & (fileLength / (length * 1024)) & "Kb/s"

        total = total + length
    Next

    average = total / 5

    Output "average: " & average & "s - " & (fileLength / (average * 1024)) & "Kb/s"
    Output ""

    total = 0
    Output "SHA256"
    For i = 1 To 5
        starttime = Timer

        SHA.SHA256 bytes

        length = Timer - starttime

        Output i & ": " & length & "s - " & (fileLength / (length * 1024)) & "Kb/s"

        total = total + length
    Next

    average = total / 5

    Output "average: " & average & "s - " & (fileLength / (average * 1024)) & "Kb/s"
    Output ""

    total = 0
    Output "SHA224"
    For i = 1 To 5
        starttime = Timer

        SHA.SHA224 bytes

        length = Timer - starttime

        Output i & ": " & length & "s - " & (fileLength / (length * 1024)) & "Kb/s"

        total = total + length
    Next

    average = total / 5

    Output "average: " & average & "s - " & (fileLength / (average * 1024)) & "Kb/s"
    Output ""
    Output "Finished"

End Sub

Private Sub Form_Load()
    cmnDlg.DialogTitle = "Open File"
    cmnDlg.InitDir = App.Path
    cmnDlg.Filter = "All Files (*.*)|*.*"
    cmnDlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNExplorer
    cmnDlg.CancelError = True
End Sub
