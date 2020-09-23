VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestcString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub TestFormat()
    Assert.Equals "hel1234lo", cstring.Format("hel{0}lo", 1234)
End Sub

Public Sub TestToCharArray()
    Dim s As String
    Dim ch() As Integer
    
    s = "Hello"
    ch = cstring.ToCharArray(s)
    Assert.Equals 5, carray.GetLength(ch, 1), "Array length should be 5."
    Assert.Equals Asc("H"), ch(0)
    Assert.Equals Asc("e"), ch(1)
    Assert.Equals Asc("l"), ch(2)
    Assert.Equals Asc("l"), ch(3)
    Assert.Equals Asc("o"), ch(4)
End Sub

Public Sub TestEmptyToCharArray()
    Dim s As String
    Dim ch() As Integer
    
    ch = cstring.ToCharArray(s)
    Assert.Equals 0, carray.GetLength(ch, 1), "Array should be empty."
End Sub

Public Sub TestPadLeft()
    Assert.Equals "   abc", cstring.PadLeft("abc", 6)
End Sub

Public Sub TestPadLeftWithCharCode()
    Assert.Equals "000abc", cstring.PadLeft("abc", 6, 48)
End Sub

Public Sub TestPadLeftTooShort()
    Assert.Equals "abc", cstring.PadLeft("abc", 2)
End Sub

Public Sub TestPadRight()
    Assert.Equals "abc   ", cstring.PadRight("abc", 6)
End Sub

Public Sub TestPadRightWithCharCode()
    Assert.Equals "abc000", cstring.PadRight("abc", 6, 48)
End Sub

Public Sub TestPadRightTooShort()
    Assert.Equals "abc", cstring.PadRight("abc", 2)
End Sub

Public Sub TestTrimStart()
    Dim chars() As Integer
    Dim s As String
    
    chars = cstring.ToCharArray("abc")
    s = "cabhi"
    s = cstring.TrimStart(s, chars)
    Assert.Equals "hi", s
End Sub

Public Sub TestTrimStartNoTrim()
    Dim chars(0) As Integer
    Dim s As String
    
    chars(0) = 48
    s = "hello"
    s = cstring.TrimStart(s, chars)
    Assert.Equals "hello", s
End Sub

Public Sub TestTrimStartAllTrimmed()
    Dim chars(0) As Integer
    Dim s As String
    
    chars(0) = 48
    s = "00000"
    s = cstring.TrimStart(s, chars)
    Assert.Equals "", s
End Sub

Public Sub TestTrimEnd()
    Dim chars() As Integer
    Dim s As String
    
    chars = cstring.ToCharArray("abc")
    s = "hicab"
    s = cstring.TrimEnd(s, chars)
    Assert.Equals "hi", s
End Sub

Public Sub TestTrimEndNoTrim()
    Dim chars(0) As Integer
    Dim s As String
    
    chars(0) = 48
    s = "hello"
    s = cstring.TrimEnd(s, chars)
    Assert.Equals "hello", s
End Sub

Public Sub TestTrimEndAllTrimmed()
    Dim chars(0) As Integer
    Dim s As String
    
    chars(0) = 48
    s = "00000"
    s = cstring.TrimEnd(s, chars)
    Assert.Equals "", s
End Sub

Public Sub TestTrim()
    Dim s As String
    
    s = "*;|@123***456@|;*"
    s = cstring.Trim(s, "*;|@")
    Assert.Equals "123***456", s
End Sub

Public Sub TestTrimWhiteSpace()
    Dim s As String
    
    s = vbCr & vbCr & " hi " & vbLf & vbLf
    s = cstring.Trim(s)
    Assert.Equals "hi", s
End Sub

Public Sub TestTrimSpaces()
    Dim s As String
    
    s = "     "
    s = cstring.Trim(s)
    Assert.Equals "", s
End Sub

Public Sub TestTrimEmptyString()
    Dim s As String
    s = cstring.Trim(s)
    Assert.Equals "", s
End Sub
