VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestUTF7GetChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestUTF7GetChars
'
Option Explicit

Private en As UTF7Encoding
Private b() As Byte
Private ch() As Integer

Public Sub Setup()
    Set en = New UTF7Encoding
End Sub

Public Sub TestPlus()
    ReDim ch(10)
    
    b = carray.NewArray(ciByte, 43)
    Assert.Equals 0, en.GetCharsex(b, 0, 1, ch, 0)
    Assert.Equals 0, ch(0)
End Sub

Public Sub TestOneNonDirect()
    ReDim ch(100)
    b = carray.NewArray(ciByte, 33)
    Assert.Equals 1, en.GetCharsex(b, 0, 1, ch, 0)
    Assert.Equals 33, ch(0)
End Sub

Public Sub TestPlusMinus()
    ReDim ch(100)
    b = carray.NewArray(ciByte, 43, 45)
    Assert.Equals 1, en.GetCharsex(b, 0, 2, ch, 0)
    Assert.Equals 43, ch(0)
    Assert.Equals 0, ch(1)
End Sub

Public Sub TestSimple()
    ReDim ch(100)
    b = en.GetBytes("hello")
    Assert.Equals 5, en.GetCharsex(b, 0, UBound(b) + 1, ch, 0)
    Assert.Equals "hello", cstring.FromCharArray(ch, 0, 5)
    Assert.Equals 0, ch(5)
End Sub

Public Sub TestJapaneseWord_nihongo()
    Dim ch() As Integer
    ch = carray.NewArray(ciInteger, &H65E5, &H672C, &H8A9E)
    b = en.GetBytes(ch)
    ReDim ch(100)
    Assert.Equals 3, en.GetCharsex(b, 0, UBound(b) + 1, ch, 0)
    Assert.Equals &H65E5, ch(0)
    Assert.Equals &H672C, ch(1)
    Assert.Equals &H8A9E, ch(2)
    Assert.Equals 0, ch(3)
End Sub

Public Sub TestGetString()
    b = en.GetBytes("Hello there.")
    Assert.Equals "Hello there.", en.GetString(b)
End Sub

Public Sub TestDecoder()
    Dim c As Decoder
    Dim i As Long
    
    b = en.GetBytes("a" & ChrW$(&H65E5) & "a")
    ReDim ch(10)
    
    Set c = en.GetDecoder
    For i = 0 To 6
        Assert.Equals Choose(i + 1, 1, 0, 0, 0, 1, 0, 1), c.GetChars(b, i, 1, ch, 0), "Failed on index " & i & "."
    Next i
    
    Set c = en.GetDecoder
    For i = 0 To 6
        Assert.Equals Choose(i + 1, 1, 1, 1, 1, 2, 2, 3), c.GetChars(b, 0, i + 1, ch, 0), "Failed on index " & i & "."
    Next i
End Sub

Public Sub TestDecoderWith1001CharsFromFile()
    Dim ff As Long
    Dim b() As Byte
    Dim ch() As Integer
    Dim i As Long
    Dim en As UTF7Encoding
    Dim c As Decoder
    Dim Char As Long
    
    ff = FreeFile
    Open App.Path & "\UTF7_Chars_0to1000.txt" For Binary As #ff
    ReDim b(0 To LOF(ff) - 1)
    Get #ff, , b
    Close #ff
    
    Set en = New UTF7Encoding
    Set c = en.GetDecoder
    
    ReDim ch(0)
    For i = 0 To UBound(b)
        If c.GetChars(b, i, 1, ch, 0) = 1 Then
            Assert.Equals Char, ch(0)
            Char = Char + 1
        End If
    Next i
End Sub

Public Sub TestGet1001CharsFromfile()
    Dim ff As Long
    Dim b() As Byte
    Dim ch() As Integer
    Dim i As Long
    Dim en As UTF7Encoding
    
    ff = FreeFile
    Open App.Path & "\UTF7_Chars_0to1000.txt" For Binary As #ff
    ReDim b(0 To LOF(ff) - 1)
    Get #ff, , b
    Close #ff
    
    Set en = New UTF7Encoding
    
    Assert.Equals 1001, en.GetCharCount(b), "GetCharCount should return 1001."
    
    ch = en.GetChars(b)
    Assert.Equals 1001, carray.GetLength(ch), "Length should be 1001."
    
    For i = 0 To 1000
        Assert.Equals i, ch(i), "Wrong char at index " & i & "."
    Next i
    
End Sub

