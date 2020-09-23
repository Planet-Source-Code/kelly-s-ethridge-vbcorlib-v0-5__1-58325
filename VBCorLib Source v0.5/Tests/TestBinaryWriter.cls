VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestBinaryWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestBinaryWriter
'
Option Explicit

Private bw As BinaryWriter
Private ms As MemoryStream

Public Sub Setup()
    Set ms = New MemoryStream
    Set bw = cor.NewBinaryWriter(ms)
End Sub

Public Sub TestByte()
    bw.WriteValue CByte(129)
    Assert.Equals 1@, ms.Length
    ms.Position = 0
    Assert.Equals 129, ms.ReadByte
End Sub

Public Sub Test10Bytes()
    Dim i As Long
    
    For i = 1 To 10
        bw.WriteValue CByte(i)
    Next i
    Assert.Equals 10@, ms.Length
    ms.Position = 0
    For i = 1 To 10
        Assert.Equals i, ms.ReadByte
    Next i
End Sub

Public Sub TestLong()
    bw.WriteValue 1&
    Assert.Equals 4@, ms.Length
    ms.Position = 0
    Assert.Equals 1, ms.ReadByte
    Assert.Equals 0, ms.ReadByte
    Assert.Equals 0, ms.ReadByte
    Assert.Equals 0, ms.ReadByte
End Sub

Public Sub Test10Longs()
    Dim i As Long
    For i = 1 To 10
        bw.WriteValue i
    Next i
    Assert.Equals 40@, ms.Length
    ms.Position = 0
    For i = 0 To 39
        If (i Mod 4) = 0 Then
            Assert.Equals i \ 4 + 1, ms.ReadByte
        Else
            Assert.Equals 0, ms.ReadByte
        End If
    Next i
End Sub

Public Sub TestString()
    Dim s As String
    
    s = "hello"
    bw.WriteValue s
    Assert.Equals 6@, ms.Length
    ms.Position = 0
    Assert.Equals 5, ms.ReadByte
    Assert.Equals 104, ms.ReadByte
    Assert.Equals 101, ms.ReadByte
    Assert.Equals 108, ms.ReadByte
    Assert.Equals 108, ms.ReadByte
    Assert.Equals 111, ms.ReadByte
End Sub

Public Sub Test200CharString()
    Dim s As String
    s = String$(200, "a")
    bw.WriteValue s
    Assert.Equals 202@, ms.Length
    ms.Position = 0
    
    Dim i As Long
    Assert.Equals 200, ms.ReadByte
    Assert.Equals 1, ms.ReadByte
    For i = 1 To 200
        Assert.Equals 97, ms.ReadByte
    Next i
End Sub


