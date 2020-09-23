VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestFileStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestFileStream
'
Option Explicit

Private fs As FileStream

Public Sub Setup()
    Set fs = newfilestream(App.Path & "\hello.txt", FileMode.OpenExisting)
End Sub

Public Sub Teardown()
    fs.CloseStream
End Sub

Public Sub TestLength()
    Assert.Equals 5@, fs.Length
    fs.CloseStream
End Sub

Public Sub TestReadByte()
    Dim b As Byte
    
    b = fs.ReadByte
    Assert.Equals Asc("h"), b
    Assert.Equals 1@, fs.Position
End Sub

Public Sub TestReadBytePassedEnd()
    Dim i As Long
    
    For i = 1 To 5
        fs.ReadByte
    Next i
    Assert.Equals 5@, fs.Position
    Assert.Equals -1, fs.ReadByte
    Assert.Equals 5@, fs.Position
End Sub

Public Sub TestRead()
    Dim b() As Byte
    Dim i As Long
    
    ReDim b(400)
    Assert.Equals 0@, fs.Position
    Assert.Equals 5, fs.ReadBlock(b, 0, 50)
    For i = 0 To 4
        Assert.Equals Choose(i + 1, 104, 101, 108, 108, 111), b(i)
    Next i
    Assert.Equals 5@, fs.Position
End Sub

Public Sub TestSeekBeforeFile()
    On Error GoTo errTrap
    fs.SeekPosition -1, FromBeginning
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim ex As Exception
    Assert.IsTrue catch(ex), "An exception should be caught."
    Assert.IsTrue TypeOf ex Is ArgumentException, "Wrong exception type caught."
End Sub

Public Sub TestReadOneByteAtATime()
    Dim b As Long
    Dim s As String
    
    b = fs.ReadByte
    Do While b >= 0
        s = s & Chr$(b)
        b = fs.ReadByte
    Loop
    
    Assert.Equals "hello", s
End Sub

Public Sub TestWriteSeekReadRead()
    Dim f As String
    Dim i As Long
    
    fs.CloseStream
    f = App.Path & "\WriteSeekReadRead.txt"
    
    Set fs = newfilestream(f, FileMode.OpenOrCreate)
    fs.SeekPosition 0, FromBeginning
    For i = 48 To 57
        fs.WriteByte i
    Next i
    fs.SeekPosition -5, FromCurrent
    Assert.Equals 53, fs.ReadByte
    fs.SeekPosition 2, FromCurrent
    Assert.Equals 8@, fs.Position, "Position should be 8."
    Assert.Equals 56, fs.ReadByte
    
End Sub
