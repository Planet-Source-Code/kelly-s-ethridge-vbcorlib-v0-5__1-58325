VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestUTF7GetBytes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestUTF7GetBytes
'
Option Explicit

Private e As UTF7Encoding
Private b() As Byte
Private ch() As Integer

Public Sub Setup()
    Set e = New UTF7Encoding
End Sub

Public Sub TestDirectlyEncodable()
    ch = carray.NewArray(ciInteger, 65)
    b = e.GetBytes(ch)
    Assert.Equals 1, carray.GetLength(b)
    Assert.Equals 65, b(0)
End Sub

Public Sub TestEncodeSinglePlus()
    ch = carray.NewArray(ciInteger, 43)
    b = e.GetBytes(ch)
    Assert.Equals 2, carray.GetLength(b)
    Assert.Equals 43, b(0)
    Assert.Equals 45, b(1)
End Sub

Public Sub TestOneNonDirectlyEncodable()
    ch = carray.NewArray(ciInteger, 33)
    b = e.GetBytes(ch)
    Assert.Equals 5, carray.GetLength(b)
    Assert.Equals 43, b(0)
    Assert.Equals 65, b(1)
    Assert.Equals 67, b(2)
    Assert.Equals 69, b(3)
    Assert.Equals 45, b(4)
End Sub

Public Sub TestTwoNonDirectlyEncodable()
    ch = carray.NewArray(ciInteger, 33, 33)
    b = e.GetBytes(ch)
    
    Assert.Equals 8, carray.GetLength(b)
    
    Dim i As Long
    For i = 0 To 7
        Assert.Equals Choose(i + 1, 43, 65, 67, 69, 65, 73, 81, 45), b(i)
    Next i
End Sub

Public Sub TestSimpleCombo()
    ch = carray.NewArray(ciInteger, &H41, &H2262, &H391, &H2E)
    b = e.GetBytes(ch)
    
    Assert.Equals 10, carray.GetLength(b)
    Dim i As Long
    For i = 0 To 9
        Assert.Equals Choose(i + 1, 65, 43, 73, 109, 73, 68, 107, 81, 45, 46), b(i)
    Next i
End Sub

Public Sub TestSimpleCombo2()
    ch = carray.NewArray(ciInteger, &H48, &H69, &H20, &H4D, &H6F, &H6D, &H20, &H2D, &H263A, &H2D, &H21)
    b = e.GetBytes(ch)
    
    Assert.Equals 19, carray.GetLength(b)
    Dim i As Long
    For i = 0 To 18
        Assert.Equals Choose(i + 1, 72, 105, 32, 77, 111, 109, 32, 45, 43, 74, 106, 111, 45, 45, 43, 65, 67, 69, 45), b(i)
    Next i
End Sub

Public Sub TestJapaneseWord_nihongo()
    ch = carray.NewArray(ciInteger, &H65E5, &H672C, &H8A9E)
    b = e.GetBytes(ch)
    Assert.Equals 10, carray.GetLength(b)
    Dim i As Long
    For i = 0 To 9
        Assert.Equals Choose(i + 1, 43, 90, 101, 86, 110, 76, 73, 113, 101, 45), b(i)
    Next i
End Sub

Public Sub TestBase64AndDirect()
    ch = carray.NewArray(ciInteger, &H65E5, 65)
    b = e.GetBytes(ch)
    Assert.Equals 6, carray.GetLength(b)
    Dim i As Long
    For i = 0 To 5
        Assert.Equals Choose(i + 1, 43, 90, 101, 85, 45, 65), b(i)
    Next i
End Sub

Public Sub TestNonDirectAndDirect()
    ch = carray.NewArray(ciInteger, 33, 65)
    b = e.GetBytes(ch)
    Assert.Equals 6, carray.GetLength(b)
    Dim i As Long
    For i = 0 To 5
        Assert.Equals Choose(i + 1, 43, 65, 67, 69, 45, 65), b(i)
    Next i
End Sub

Public Sub TestComplexCombo()
    b = e.GetBytes(ChrW$(&H65E5) & "hi!" & ChrW$(&H8A9E))
    Dim i As Long
    For i = 0 To 14
        Assert.Equals Choose(i + 1, 43, 90, 101, 85, 45, 104, 105, 43, 65, 67, 71, 75, 110, 103, 45), b(i)
    Next i
End Sub

Public Sub TestEncoderWithFlush()
    Dim c As Encoder
    Dim i As Long
    
    Set c = e.GetEncoder
    ch = carray.NewArray(ciInteger, &H65E5)
    ReDim b(10)
    
    Assert.Equals 5, c.GetBytes(ch, 0, 1, b, 0, True)
    For i = 0 To 4
        Assert.Equals Choose(i + 1, 43, 90, 101, 85, 45), b(i)
    Next i
End Sub

Public Sub TestEncoderWithoutFlush()
    Dim c As Encoder
    Dim i As Long
    
    Set c = e.GetEncoder
    ReDim ch(0)
    ch(0) = &H65E5
    ReDim b(3)
    
    Assert.Equals 3, c.GetBytes(ch, 0, 1, b, 0, False)
    For i = 0 To 2
        Assert.Equals Choose(i + 1, 43, 90, 101), b(i)
    Next i
    
    ch(0) = 97
    Assert.Equals 3, c.GetBytes(ch, 0, 1, b, 0, True)
    For i = 0 To 2
        Assert.Equals Choose(i + 1, 85, 45, 97), b(i)
    Next i
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
    
    Assert.Equals 2565, carray.GetLength(b)

    ff = FreeFile
    Open App.Path & "\UTF7_Chars_0to1000.txt" For Binary As #ff
    ReDim fileBytes(0 To LOF(ff) - 1)
    Get #ff, , fileBytes
    Close #ff

    For i = 0 To carray.GetLength(b) - 1
        Assert.Equals fileBytes(i), b(i)
    Next i
End Sub

