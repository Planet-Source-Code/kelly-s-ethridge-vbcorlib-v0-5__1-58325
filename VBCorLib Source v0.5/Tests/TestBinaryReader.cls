VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestBinaryReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   TestBinaryReader
'
Option Explicit

Private ms As MemoryStream
Private br As BinaryReader

Public Sub Setup()
    Set ms = New MemoryStream
    Set br = Cor.NewBinaryReader(ms)
End Sub

Public Sub TestReadByte()
    ms.WriteByte 129
    ms.Position = 0
    
    Assert.Equals 129, br.ReadByte
End Sub

Public Sub TestRead10Bytes()
    Dim i As Long
    For i = 1 To 10
        ms.WriteByte i
    Next i
    ms.Position = 0
    
    For i = 1 To 10
        Assert.Equals i, br.ReadByte
    Next i
End Sub

Public Sub TestReadLong()
    ms.WriteByte &HFF
    ms.WriteByte &H0
    ms.WriteByte &H80
    ms.WriteByte &H80
    ms.Position = 0
    
    Assert.Equals &H808000FF, br.ReadLong
End Sub

Public Sub TestReadString()
    Dim bw As BinaryWriter
    Set bw = Cor.NewBinaryWriter(ms)
    bw.WriteValue "hello" & ChrW$(&H2030)
    ms.Position = 0
    
    Assert.Equals "hello" & ChrW$(&H2030), br.ReadString
End Sub

Public Sub TestReadIntegerStringDouble()
    Dim bw As BinaryWriter
    Set bw = Cor.NewBinaryWriter(ms)
    bw.WriteValue 67
    bw.WriteValue "This is a string!"
    bw.WriteValue 12345.6789
    ms.Position = 0
    
    Assert.Equals 67, br.ReadInteger
    Assert.Equals "This is a string!", br.ReadString
    Assert.Equals 12345.6789, br.ReadDouble
    
End Sub

Public Sub TestReadChar()
    Dim bw As BinaryWriter
    Set bw = Cor.NewBinaryWriter(ms)
    bw.WriteValue ChrW$(&H2030)
    ms.Position = 1 ' move passed the 1 byte string length descriptor
    
    Assert.Equals &H2030, br.Read
    Assert.Equals -1, br.Read
End Sub

Public Sub TestPeekChar()
    ms.WriteByte 97
    ms.Position = 0
    Assert.Equals 97, br.PeekChar
    Assert.Equals 97, br.Read
    Assert.Equals -1, br.Read
End Sub

Public Sub TestReadBigString()
    Dim i As Long
    Dim s As String
    Dim bw As BinaryWriter
    
    For i = 1 To 1000
        s = s & ChrW$(i)
    Next i
    
    Set bw = Cor.NewBinaryWriter(ms)
    bw.WriteValue s
    
    ms.Position = 0
    
    Assert.Equals s, br.ReadString
End Sub
