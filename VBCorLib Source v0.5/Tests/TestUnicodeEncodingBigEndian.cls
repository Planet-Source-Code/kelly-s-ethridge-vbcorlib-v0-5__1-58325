VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestUnicodeEncodingBigEndian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestUnicodeEncodingBigEndian
'
Option Explicit

Private en As UnicodeEncoding


Public Sub Setup()
    Set en = NewUnicodeEncoding(True, True)
End Sub

Public Sub TestByteOrderMark()
    Dim b() As Byte
    
    b = en.GetPreamble
    Assert.Equals 2, carray.GetLength(b)
    Assert.Equals &HFF, b(0)
    Assert.Equals &HFE, b(1)
End Sub

Public Sub TestGetByteCount()
    Assert.Equals 100, en.GetByteCount(String$(50, " "))
End Sub

Public Sub TestGetCharCount()
    Dim b() As Byte
    
    b = carray.CreateInstance(ciByte, 100)
    Assert.Equals 50, en.GetCharCount(b)
End Sub

Public Sub TestGetBytes()
    Dim b() As Byte
    Dim i As Long
    
    b = en.GetBytes(String$(50, " "))
    Assert.Equals 100, carray.GetLength(b)
    
    For i = 0 To UBound(b)
        If i And 1 Then
            Assert.Equals 32, b(i)
        Else
            Assert.Equals 0, b(i)
        End If
    Next i
End Sub

Public Sub TestGetChars()
    Dim ch() As Integer
    Dim b() As Byte
    Dim i As Long
    
    b = en.GetBytes(String$(50, " "))
    ch = en.GetChars(b)
    Assert.Equals 50, carray.GetLength(ch)
    
    For i = 0 To UBound(ch)
        Assert.Equals 32, ch(i)
    Next i
End Sub

Public Sub TestDecoderWithEvenBytes()
    Dim c As Decoder
    Dim b(99) As Byte
    Dim ch(100) As Integer
    Dim i As Long
    
    For i = 0 To 99
        If (i And 1) = 1 Then
            b(i) = 32
        End If
    Next i
    
    Set c = en.GetDecoder
    Assert.Equals 50, c.GetCharCount(b, 0, 100)
    Assert.Equals 50, c.GetChars(b, 0, 100, ch, 0)
    
    For i = 0 To 49
        Assert.Equals 32, ch(i)
    Next i
End Sub

Public Sub TestDecoderWithOddBytes()
    Dim c As Decoder
    Dim b(99) As Byte
    Dim ch(100) As Integer
    Dim i As Long
    
    For i = 0 To 99
        If (i And 1) = 1 Then
            b(i) = 32
        End If
    Next i
    
    Set c = en.GetDecoder
    Assert.Equals 18, c.GetCharCount(b, 0, 37)
    Assert.Equals 18, c.GetChars(b, 0, 37, ch, 0)
    
    For i = 0 To 17
        Assert.Equals 32, ch(i)
    Next i
    
    ch(0) = 0
    Assert.Equals 1, c.GetCharCount(b, 37, 1)
    Assert.Equals 1, c.GetChars(b, 37, 1, ch, 0)
    Assert.Equals 32, ch(0)
End Sub

Public Sub TestGetString()
    Dim b() As Byte
    
    b = en.GetBytes("Hello, World!")
    Assert.Equals "Hello, World!", en.GetString(b)
End Sub

