VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestFileStreamSmallBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestFileStreamSmallBuffer
'
Option Explicit

Private fs As FileStream

Public Sub Setup()
    Set fs = newfilestream(App.Path & "\100bytes.txt", OpenExisting, buffersize:=8)
End Sub

Public Sub Teardown()
    fs.CloseStream
End Sub

Public Sub TestLength()
    Assert.Equals 100@, fs.Length, "Length should be 100 bytes."
End Sub

Public Sub TestSeekToChangeBuffer()
    Assert.Equals 0@, fs.Position
    Assert.Equals 49, fs.ReadByte
    Assert.Equals 11@, fs.SeekPosition(11, FromBeginning)
    Assert.Equals 50, fs.ReadByte
End Sub

Public Sub TestReadLarge()
    Dim b() As Byte
    Dim i As Long
    
    ReDim b(24)
    fs.Position = 12
    Assert.Equals 25, fs.ReadBlock(b, 0, 25)
    For i = 0 To 24
        Assert.Equals 48 + ((i + 3) Mod 10), b(i)
    Next i
    Assert.Equals 56, fs.ReadByte
End Sub

Public Sub TestReadAll()
    Dim b() As Byte
    Dim i As Long
    
    ReDim b(100)
    Assert.Equals 100, fs.ReadBlock(b, 0, 100)
    For i = 0 To 99
        Assert.Equals 48 + ((i + 1) Mod 10), b(i)
    Next i
    Assert.Equals 0, fs.ReadBlock(b, 0, 1)
    Assert.Equals -1, fs.ReadByte
End Sub

Public Sub ReadInconsistent()
    Dim b() As Byte
    Dim i As Long
    
    ReDim b(100)
    fs.ReadBlock b, 0, 6
    fs.ReadByte
    Assert.Equals 5, fs.ReadBlock(b, 0, 5)
    
    For i = 0 To 4
        Assert.Equals 48 + ((8 + i) Mod 10), b(i)
    Next i
End Sub

Public Sub TestReadOneByteAtATime()
    Dim i As Long
    Dim b As Long
    Dim j As Long
    
    b = fs.ReadByte
    Do While b >= 0
        j = j + 1
        i = (i + 1) Mod 10
        Assert.Equals 48 + i, b
        b = fs.ReadByte
    Loop
    
    Assert.Equals 100, j
End Sub
