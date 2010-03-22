VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISHAStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function InitializeHash() As Long()
End Function

Public Function Expand(Words() As Long, Index As Long) As Long
End Function

Public Sub Round(lRound As Long, Hash() As Long, Value As Long)
End Sub

Public Function Output(Hash() As Long) As String
End Function

Public Property Get NumRounds() As Long
End Property
