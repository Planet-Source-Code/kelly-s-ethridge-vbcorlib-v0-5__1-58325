VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestUTF8GetChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestUTF8GetChars
'
Option Explicit

Private e As UTF8Encoding
Private b() As Byte
Private ch() As Integer

Public Sub Setup()
    Set e = NewUTF8Encoding(False)
End Sub

Public Sub Teardown()
    Erase b
    Erase ch
End Sub

Public Sub TestFrom1ByteAnsii()
    Dim ch() As Integer
    ch = e.GetChars(carray.NewArray(ciByte, 65))
    Assert.Equals 1, carray.GetLength(ch)
    Assert.Equals 65, ch(0)
End Sub

Public Sub TestFrom10BytesAnsii()
    Dim ch() As Integer
    Dim i As Long
    
    b = e.GetBytes(String$(10, &H7F))
    ch = e.GetChars(b)
    Assert.Equals 10, carray.GetLength(ch)
    For i = 0 To 9
        Assert.Equals &H7F, ch(i)
    Next i
End Sub

Public Sub TestSurrogatePair()
    b = e.GetBytes(carray.NewArray(ciInteger, &HDBFF, &HDC00))
    ch = e.GetChars(b)
    Assert.Equals 2, carray.GetLength(ch)
    Assert.Equals &HDBFF, ch(0)
    Assert.Equals &HDC00, ch(1)
End Sub

Public Sub TestNonAscii()
    b = e.GetBytes(String$(10, 167))
    ch = e.GetChars(b)
    Assert.Equals 10, carray.GetLength(ch)
    
    Assert.Equals String$(10, 167), cstring.FromCharArray(ch)
End Sub

Public Sub TestComboChars()
    Dim asciiChars As String
    Dim nonAsciiChars As String
    Dim surrogateChars As String
    Dim i As Long
    
    asciiChars = String$(10, "A")
    nonAsciiChars = String$(10, ChrW$(167))
    For i = 1 To 10
        surrogateChars = surrogateChars & ChrW$(&HD800) & ChrW$(&HDC00)
    Next i
    Dim s As String
    s = nonAsciiChars & surrogateChars & asciiChars
    
    b = e.GetBytes(s)
    ch = e.GetChars(b)
    Assert.Equals 40, carray.GetLength(ch)
    Assert.Equals s, cstring.FromCharArray(ch)
End Sub

Public Sub TestGetStringComboChars()
    Dim asciiChars As String
    Dim nonAsciiChars As String
    Dim surrogateChars As String
    Dim i As Long
    
    asciiChars = String$(10, "A")
    nonAsciiChars = String$(10, ChrW$(167))
    For i = 1 To 10
        surrogateChars = surrogateChars & ChrW$(&HD801) & ChrW$(&HDC01)
    Next i
    Dim s As String
    s = nonAsciiChars & surrogateChars & asciiChars
    
    b = e.GetBytes(s)
    Dim s2 As String
    s2 = e.GetString(b)
    Assert.Equals 40, Len(s)
    Assert.Equals s, s
End Sub

Public Sub TestCharBufferTooSmall()
    Dim ch() As Integer
    b = e.GetBytes("Hello")
    ReDim ch(3)
    
    On Error GoTo errTrap
    e.GetCharsex b, 0, 5, ch, 0
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim ex As Exception
    Assert.IsTrue catch(ex), "An exception should be caught."
    Assert.IsTrue TypeOf ex Is ArgumentException, "Incorrect exception type."
End Sub

Public Sub TestShortestFrom()
    Dim ch() As Integer
    
    b = e.GetBytes(ChrW$(8191))
    Assert.Equals 3, carray.GetLength(b)
    ch = e.GetChars(b)
    Assert.Equals 1, carray.GetLength(ch)
    Assert.Equals 8191, ch(0)
End Sub

Public Sub TestLongBuffer()
    b = e.GetBytes("hello")
    ReDim Preserve b(0 To 100)
    Assert.Equals "hello", e.GetString(b, 0, 5)
End Sub

Public Sub TestGet1001CharsFromfile()
    Dim ff As Long
    Dim b() As Byte
    Dim ch() As Integer
    Dim i As Long
    Dim en As UTF8Encoding
    
    ff = FreeFile
    Open App.Path & "\UTF8_Chars_0to1000.txt" For Binary As #ff
    ReDim b(0 To LOF(ff) - 1)
    Get #ff, , b
    Close #ff
    
    Set en = New UTF8Encoding
    ch = en.GetChars(b)
    Assert.Equals 1001, carray.GetLength(ch)
    
    For i = 0 To 1000
        Assert.Equals i, ch(i)
    Next i
End Sub

Public Sub TestBadSurrogatePair()
    Dim ch() As Integer
    
    b = e.GetBytes(ChrW$(&HD800) & ChrW$(&HD801))
    ch = e.GetChars(b)
    Assert.Equals &HD800, ch(0)
    Assert.Equals &HD801, ch(1)
End Sub
