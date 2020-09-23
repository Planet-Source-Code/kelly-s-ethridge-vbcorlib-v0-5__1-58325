VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestStringReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   TestStringReader
'
Option Explicit

Private sr As StringReader



Public Sub TestRead()
    Dim ch As Long
    Dim i As Long
    
    Set sr = newstringreader("abcdefghijk")
    i = Asc("a")
    ch = sr.Read
    Do While ch > 0
        Assert.Equals i, ch
        i = i + 1
        ch = sr.Read
    Loop
    Assert.Equals Asc("l"), i
    Assert.Equals -1, sr.Read
End Sub

Public Sub TestPeek()
    Set sr = newstringreader("a")
    Assert.Equals Asc("a"), sr.Peek
    Assert.Equals Asc("a"), sr.Read
    Assert.Equals -1, sr.Read
End Sub

Public Sub TestReadBlock()
    Dim ch(100) As Integer
    Dim i As Long
    
    Set sr = newstringreader("abcdefghijk")
    
    Assert.Equals 11, sr.ReadBlock(ch, 0, 100)
    For i = 0 To 10
        Assert.Equals Asc("a") + i, ch(i)
    Next i
End Sub

Public Sub TestReadLine()
    Set sr = newstringreader("abc" & vbCr & vbLf & "def" & vbLf & vbCr)
    Assert.Equals "abc", sr.ReadLine
    Assert.Equals "def", sr.ReadLine
    Assert.IsNotNull sr.ReadLine
    Assert.IsNull sr.ReadLine
End Sub

Public Sub TestReadToEnd()
    Set sr = newstringreader("abcdefghi")
    sr.Read
    Assert.Equals "bcdefghi", sr.ReadToEnd
End Sub
