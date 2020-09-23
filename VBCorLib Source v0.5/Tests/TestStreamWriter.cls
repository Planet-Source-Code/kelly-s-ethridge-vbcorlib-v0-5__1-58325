VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestStreamWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestStreamWriter
'
Option Explicit

Public Sub TestWriteChars0to64000()
    Dim sw As StreamWriter
    Dim i As Long
    Dim f As String
    
    f = App.Path & "\UTF8_Chars0To64000.txt"
    If Len(Dir$(f)) > 0 Then Kill f
    
    Set sw = newstreamwriter(f, newutf8encoding(True))
    For i = 0 To 64000
        sw.WriteValue ChrW$(i)
    Next i
    sw.CloseWriter
    
    Dim en As Encoding
    Dim b() As Byte
    Dim ff As Integer
    Dim ch() As Integer
    
    Set en = sw.Encoding
    
    ff = FreeFile
    Open f For Binary As #ff
    ReDim b(0 To LOF(ff) - 1)
    Get #ff, , b
    Close #ff
    Kill f
    
    Assert.Equals 189828, carray.GetLength(b)
    ' test for UTF8 Byte Order Mark
    Assert.Equals &HEF, b(0)
    Assert.Equals &HBB, b(1)
    Assert.Equals &HBF, b(2)
    
    ch = en.GetChars(b, 3, carray.GetLength(b) - 3)
    Assert.Equals 64001, carray.GetLength(ch)
    
    Dim j As Long
    For i = 0 To 64000
        AsWord(j) = ch(i)   ' we need to convert negative values to position without extending the sign
        Assert.Equals i, j, "Wrong character as index " & i
    Next i
    
End Sub

Public Sub TestWriteHugeStringAtOnce()
    Dim sw As StreamWriter
    Dim s As String
    Dim f As String
    Dim en As Encoding
    Dim ff As Integer
    Dim b() As Byte
    Dim i As Long
    
    f = App.Path & "\100003bytes.txt"
    If Len(Dir$(f)) > 0 Then Kill f
    
    Set sw = newstreamwriter(f, newutf8encoding(True))
    
    s = String$(100000, "1")
    sw.WriteValue s
    sw.CloseWriter
    Set en = sw.Encoding
    
    ff = FreeFile
    Open f For Binary As #ff
    ReDim b(0 To LOF(ff) - 1)
    Get #ff, , b
    Close #ff
    Kill f
    Assert.Equals 100003, carray.GetLength(b), "File length should be 100003 (includes Byte Order Mark)."
    
    For i = 3 To UBound(b)
        Assert.Equals 49, b(i)
    Next i
    
End Sub

