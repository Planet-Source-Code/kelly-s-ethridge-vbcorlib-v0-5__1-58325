VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestUTF8Encoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestUTF8Encoding
'
Option Explicit

Private e As UTF8Encoding
Private s As String
Private b() As Byte

Public Sub Setup()
    Set e = NewUTF8Encoding(True)
End Sub

Public Sub TestGetPreamble()
    Dim b() As Byte
    
    b = e.GetPreamble
    Assert.IsFalse carray.IsNull(b)
    Assert.Equals 3, carray.GetLength(b)
    '0xEF, 0xBB, 0xBF
    Assert.Equals &HEF, b(0)
    Assert.Equals &HBB, b(1)
    Assert.Equals &HBF, b(2)
End Sub

Public Sub TestGetBytesFrom1CharString()
    s = Chr$(&H7F)
    b = e.GetBytes(s)
    Assert.Equals 1, carray.GetLength(b)
    Assert.Equals &H7F, b(0)
End Sub

Public Sub TestGetBytesFrom100CharString()
    s = String$(100, &H7F)
    b = e.GetBytes(s)
    Assert.Equals 100, carray.GetLength(b)
    
    Dim i As Long
    For i = 0 To 99
        Assert.Equals &H7F, b(i)
    Next i
End Sub

Public Sub TestGetBytesFromSurrogatePair()
    Dim ch() As Integer
    
    ch = carray.NewArray(ciInteger, &HD800, &HDC00)
    b = e.GetBytes(ch)
    Assert.Equals 4, carray.GetLength(b)
    Assert.Equals 240, b(0)
    Assert.Equals 144, b(1)
    Assert.Equals 128, b(2)
    Assert.Equals 128, b(3)
End Sub

Public Sub TestGetBytesFromHighSurrogateOnly()
    Dim ch() As Integer
    
    ch = carray.NewArray(ciInteger, &HD801)
    b = e.GetBytes(ch)
    Assert.Equals 3, carray.GetLength(b)
    Assert.Equals 237, b(0)
    Assert.Equals 160, b(1)
    Assert.Equals 129, b(2)
End Sub

Public Sub TestGetBytesFromLowSurrogateOnly()
    b = e.GetBytes(ChrW$(&HDC01))
    Assert.Equals 3, carray.GetLength(b)
    Assert.Equals 237, b(0)
    Assert.Equals 176, b(1)
    Assert.Equals 129, b(2)
End Sub

Public Sub TestGetBytesExFromHighSurrogateOnly()
    Dim ch() As Integer
    
    ch = carray.NewArray(ciInteger, &HD801)
    ReDim b(0 To e.GetByteCount(ch) - 1)
    Assert.Equals 3, e.GetBytesEx(ch, 0, 1, b, 0)
    Assert.Equals 3, carray.GetLength(b)
    Assert.Equals 237, b(0)
    Assert.Equals 160, b(1)
    Assert.Equals 129, b(2)
End Sub

Public Sub TestGetBytesExFromLowSurrogateOnly()
    ReDim b(0 To e.GetByteCount(ChrW$(&HDC01)) - 1)
    Assert.Equals 3, e.GetBytesEx(ChrW$(&HDC01), 0, 1, b, 0)
    Assert.Equals 3, carray.GetLength(b)
    Assert.Equals 237, b(0)
    Assert.Equals 176, b(1)
    Assert.Equals 129, b(2)
End Sub

Public Sub TestGetBytesFromAll()
    b = e.GetBytes(carray.NewArray(ciInteger, 97, &HD801, &HDC01, &H600))
    Assert.Equals 7, carray.GetLength(b)
    Assert.Equals 97, b(0)
    Assert.Equals 240, b(1)
    Assert.Equals 144, b(2)
    Assert.Equals 144, b(3)
    Assert.Equals 129, b(4)
    Assert.Equals 216, b(5)
    Assert.Equals 128, b(6)
End Sub

Public Sub TestGetBytesFromSubArray()
    b = e.GetBytes(carray.NewArray(ciInteger, 97, &HD801, &HDC01, &H600), 1, 2)
    Assert.Equals 4, carray.GetLength(b)
    Assert.Equals 240, b(0)
    Assert.Equals 144, b(1)
    Assert.Equals 144, b(2)
    Assert.Equals 129, b(3)
End Sub

Public Sub TestGetBytesEx()
    ReDim b(6)
    Assert.Equals 7, e.GetBytesEx(carray.NewArray(ciInteger, 97, &HD801, &HDC01, &H600), 0, 4, b, 0)
    Assert.Equals 97, b(0)
    Assert.Equals 240, b(1)
    Assert.Equals 144, b(2)
    Assert.Equals 144, b(3)
    Assert.Equals 129, b(4)
    Assert.Equals 216, b(5)
    Assert.Equals 128, b(6)
End Sub

Public Sub TestGetBytesExBytesTooSmall()
    ReDim b(5)
    
    On Error GoTo errTrap
    e.GetBytesEx carray.NewArray(ciInteger, 97, &HD801, &HDC01, &H600), 0, 4, b, 0
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim ex As Exception
    Assert.IsTrue catch(ex), "An exception should be caught."
End Sub

Public Sub TestEncoder()
    Dim en As Encoder
    ReDim b(3)
    Dim ch(0) As Integer
    
    Set en = e.GetEncoder
    
    ch(0) = &HD801
    Assert.Equals 0, en.GetBytes(ch, 0, 1, b, 0, False)
    ch(0) = &HDC02
    Assert.Equals 4, en.GetBytes(ch, 0, 1, b, 0, False)
    Assert.Equals 240, b(0)
    Assert.Equals 144, b(1)
    Assert.Equals 144, b(2)
    Assert.Equals 130, b(3)
End Sub

Public Sub TestDecoder()
    Dim de As Decoder
    Dim ch(1) As Integer
    
    Set de = e.GetDecoder
    b = e.GetBytes(ChrW$(&HD801) & ChrW$(&HDC02))
    Assert.Equals 0, de.GetChars(b, 0, 1, ch, 0)
    Assert.Equals 2, de.GetChars(b, 1, 3, ch, 0)
    Assert.Equals &HD801, ch(0)
    Assert.Equals &HDC02, ch(1)
End Sub

Public Sub TestEncodeChars0To1000()
    Dim ch() As Integer
    Dim i As Long
    Dim b() As Byte
    Dim fileBytes() As Byte
    Dim ff As Integer
    
    ReDim ch(1000)
    For i = 0 To 1000
        ch(i) = i
    Next i
    b = e.GetBytes(ch)
    
    Assert.Equals 1874, carray.GetLength(b)
    ff = FreeFile
    Open App.Path & "\UTF8_Chars_0to1000.txt" For Binary As #ff
    ReDim fileBytes(0 To LOF(ff) - 1)
    Get #ff, , fileBytes
    Close #ff
    
    For i = 0 To carray.GetLength(b) - 1
        Assert.Equals fileBytes(i), b(i)
    Next i
End Sub

Public Sub TestGetBytesFromHugeString()
    Dim s As String
    Dim b() As Byte
    Dim i As Long
    
    s = String$(100000, "1")
    Assert.Equals 100000, e.GetByteCount(s)
    b = e.GetBytes(s)
    Assert.Equals 100000, carray.GetLength(b)
    For i = 0 To UBound(b)
        Assert.Equals 49, b(i)
    Next i
End Sub
