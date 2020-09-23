VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestFileStreamWrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestFileStreamWrite
'
Option Explicit

Private fs As FileStream

Public Sub Setup()
    Set fs = newfilestream(App.Path & "\testwrite.txt", Create)
End Sub

Public Sub Teardown()
    fs.CloseStream
End Sub

Public Sub TestWriteByte()
    fs.WriteByte 48
    Assert.Equals 1@, fs.Length
    fs.SeekPosition 0, FromBeginning
    Assert.Equals 48, fs.ReadByte
End Sub

Public Sub TestWriteAlphabet()
    Dim i As Long
    
    For i = 65 To 90
        fs.WriteByte i
    Next i
    Assert.Equals 26@, fs.Length
    
    fs.SeekPosition 0, FromBeginning
    For i = 65 To 90
        Assert.Equals i, fs.ReadByte
    Next i
End Sub

Public Sub TestOverWrite()
    Dim i As Long
    
    For i = 65 To 90
        fs.WriteByte i
    Next i
    
    fs.SeekPosition -13, FromCurrent
    For i = 110 To 122
        fs.WriteByte i
    Next i
    Assert.Equals 26@, fs.Length
    
    fs.Position = 0
    For i = 65 To 77
        Assert.Equals i, fs.ReadByte
    Next i
    For i = 110 To 122
        Assert.Equals i, fs.ReadByte
    Next i
End Sub

Public Sub TestSetLength()
    Assert.Equals 0@, fs.Length
    fs.SetLength 100
    Assert.Equals 100@, fs.Length
End Sub

Public Sub TestAppend()
    Dim i As Long
    
    Assert.Equals 0@, fs.Length
    For i = 0 To 9
        fs.WriteByte 48 + i
    Next i
    
    fs.CloseStream
    Set fs = newfilestream(App.Path & "\testwrite.txt", Append)
    
    For i = 65 To 90
        fs.WriteByte i
    Next i
    Assert.Equals 36@, fs.Length
    
    fs.CloseStream
    Set fs = newfilestream(App.Path & "\testwrite.txt", OpenExisting)
    
    For i = 0 To 9
        Assert.Equals 48 + i, fs.ReadByte
    Next i
    For i = 65 To 90
        Assert.Equals i, fs.ReadByte
    Next i
End Sub

Public Sub TestWriteLarge()
    Dim b() As Byte
    Dim i As Long
    Dim ff As Integer
    
    ReDim b(100000)
    For i = 0 To 100000
        b(i) = i And &HFF
    Next i
    
    fs.WriteBlock b, 0, 100001
    fs.CloseStream
    
    ff = FreeFile
    Open App.Path & "\testwrite.txt" For Binary As #ff
    ReDim b(0 To LOF(ff) - 1)
    Get #ff, , b
    Close #ff
    Assert.Equals 100001, carray.GetLength(b)
    
    For i = 0 To 100000
        Assert.Equals i And &HFF, b(i)
    Next i
    
    Kill App.Path & "\testwrite.txt"
End Sub
