VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestStreamWriterWithMem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestStreamWriterWithMem
'
Option Explicit

Private sw As StreamWriter
Private ms As MemoryStream

Public Sub Setup()
    Set ms = New MemoryStream
    Set sw = cor.NewStreamWriter(ms)
End Sub

Public Sub TestWriteLetter()
    sw.WriteValue "a"
    Assert.Equals 0@, ms.Length
    sw.Flush
    Assert.Equals 1@, ms.Length
    ms.Position = 0
    Assert.Equals 97, ms.ReadByte
End Sub
