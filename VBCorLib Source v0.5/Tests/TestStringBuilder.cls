VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sb As StringBuilder

Public Sub Setup()
    Set sb = New StringBuilder
End Sub

Public Sub TestAppend()
    Assert.Equals "HelloBye", sb.Append("Hello").Append("Bye").ToString
End Sub

Public Sub TestAppendString()
    sb.Append "Hello"
    sb.Append "Bye"
    Assert.Equals "HelloBye", sb.ToString
End Sub

Public Sub TestAppendChar()
    sb.AppendChar 48, 50
    Assert.Equals String$(50, "0"), sb.ToString
End Sub

Public Sub TestChars()
    sb.Append "1111111111"
    sb(2) = 50
    sb(3) = 50
    
    Assert.Equals 49, sb(1)
    Assert.Equals "1122111111", sb.ToString
End Sub

Public Sub TestSetLengthPastCapacity()
    sb.capacity = 10
    sb.Append "12345"
    sb.Length = 20
    
    Assert.Equals "12345               ", sb.ToString
    
End Sub

Public Sub TestLength()
    sb.Append "1234567890"
    sb.Length = 5
    sb.Append "abcde"
    Assert.Equals "12345abcde", sb.ToString
End Sub

Public Sub TestAppendSubString()
    sb.AppendString "Hello", 2, 3
    Assert.Equals "llo", sb.ToString
End Sub

Public Sub TestAppendChars()
    Dim chars(9) As Integer
    Dim i As Long
    For i = 0 To 9
        chars(i) = 48 + i
    Next i
    
    sb.AppendChars chars
    Assert.Equals "0123456789", sb.ToString
End Sub

Public Sub TestAppendSubChars()
    Dim chars(9) As Integer
    Dim i As Long
    For i = 0 To 9
        chars(i) = 48 + i
    Next i
    
    sb.AppendChars chars, 2, 3
    Assert.Equals "234", sb.ToString
End Sub

Public Sub TestInsertString()
    sb.Append "1234567890"
    sb.Insert 2, "aaa"
    Assert.Equals "12aaa34567890", sb.ToString
End Sub

Public Sub InsertBeginning()
    sb.Insert 0, Space$(10), 10
    Assert.Equals Space$(100), sb.ToString
End Sub

Public Sub TestInsertAtEnd()
    sb.Append "Hello"
    sb.Insert 5, "Bye"
    Assert.Equals "HelloBye", sb.ToString
End Sub

Public Sub TestInsertStringCount()
    sb.Append "1234567890"
    sb.Insert 2, "aaa", 2
    Assert.Equals "12aaaaaa34567890", sb.ToString
End Sub

Public Sub TestInsertChars()
    sb.Append "1234567890"
    Dim chars() As Integer
    chars = carray.NewArray(ciInteger, 97, 97, 97)
    sb.Insertchars 2, chars
    Assert.Equals "12aaa34567890", sb.ToString
End Sub

Public Sub TestInsertCharsAtEnd()
    sb.Append "Hello"
    Dim chars() As Integer
    chars = carray.NewArray(ciInteger, 97, 97, 97)
    sb.Insertchars 5, chars
    Assert.Equals "Helloaaa", sb.ToString
End Sub

Public Sub TestInsertSubChars()
    sb.Append "1234567890"
    Dim chars() As Integer
    chars = carray.NewArray(ciInteger, 97, 98, 99, 100)
    sb.Insertchars 2, chars, 1, 3
    Assert.Equals "12bcd34567890", sb.ToString
End Sub

Public Sub TestRemoveFromStart()
    sb.Append "1234567890"
    sb.Remove 0, 5
    Assert.Equals "67890", sb.ToString
End Sub

Public Sub TestRemoveFromEnd()
    sb.Append "1234567890"
    sb.Remove 5, 5
    Assert.Equals "12345", sb.ToString
End Sub

Public Sub TestReplace()
    sb.Append "1234567890"
    sb.Replace "45", ""
    Assert.Equals "12367890", sb.ToString
End Sub

Public Sub TestReplaceMultiple()
    sb.Insert 0, "Hello", 2
    sb.Replace "ll", "mmmm"
    
    Dim s As String
    Dim i As Long
    For i = 1 To 2
        s = s & "Hemmmmo"
    Next i
    Assert.Equals s, sb.ToString
End Sub

Public Sub TestReplaceWithSubString()
    sb.Append "1234567890"
    sb.Replace "45", "Hello", 1, 3
    Assert.Equals "123ell67890", sb.ToString
End Sub

Public Sub XBeatTest()
    Dim rep As String
    Dim find As String
    Dim src As String
    
    find = "ha"
    rep = "blahblah"
    src = "hahahahahahahahahaha"
    sb.Append src
    sb.Replace find, rep
    Assert.Equals Replace$(src, find, rep), sb.ToString
End Sub

Public Sub XBeatTest2()
    sb.Insert 0, "xx" & Space$(10000) & "xx", 10
    sb.Replace "x", Space$(10000)
    Assert.Equals Space$(500000), sb.ToString
End Sub

Public Sub TestReplaceLastChar()
    sb.Insert 0, "abc", 10
    sb.Append "d"
    sb.Replace "cd", "1234"
    Assert.Equals "abcabcabcabcabcabcabcabcabcab1234", sb.ToString
End Sub

Public Sub TestAppendFormat()
    sb.AppendFormat "Hello {0}", "World"
    Assert.Equals "Hello World", sb.ToString
End Sub

Public Sub TestAppendFormatLeftAdjust()
    sb.AppendFormat "Hello {0, -5 }", "1"
    Assert.Equals "Hello 1    ", sb.ToString
End Sub

Public Sub TestAppendFormatRightAdjust()
    sb.AppendFormat "Hello {0, 10 } world!", "kelly"
    Assert.Equals "Hello      kelly world!", sb.ToString
End Sub

Public Sub TestAppendFormatMultiArgs()
    sb.AppendFormat "1={0}, 2={1}, 1={0}, Name={2}", 1, 2, "Kelly"
    Assert.Equals "1=1, 2=2, 1=1, Name=Kelly", sb.ToString
End Sub

Public Sub TestAppendFormatFormattable()
    Dim i1 As New Int32Formattable
    Dim i2 As New Int32Formattable
    i1.mValue = 10
    i2.mValue = 20
    sb.AppendFormat "10={0}, 20={1}", i1, i2
    Assert.Equals "10=10, 20=20", sb.ToString
End Sub

Public Sub TestcObjectBased()
    Dim i1 As New Int32
    Dim i2 As New Int32
    i1.mValue = 10
    i2.mValue = 20
    sb.AppendFormat "10={0}, 20={1}", i1, i2
    Assert.Equals "10=10, 20=20", sb.ToString
End Sub

Public Sub TestFormatInfo()
    On Error Resume Next
    Dim i1 As New Int32Formattable
    i1.mValue = 1
    sb.AppendFormat "{0:9}", i1
    Assert.Equals "000000001", sb.ToString
End Sub

Public Sub TestAppendFormatNoArgs()
    sb.AppendFormat "Hello"
    Assert.Equals "Hello", sb.ToString
End Sub

Public Sub TestAppendFormatInt32()
    sb.AppendFormat "Hello {0:G3}", 1235
    Assert.Equals "Hello 1.24E+03", sb.ToString
End Sub

Public Sub TestAppendFormatHex()
    sb.AppendFormat "{0:x}", &HFFFF
    Assert.Equals "ffff", sb.ToString
    
    sb.Length = 0
    sb.AppendFormat "{0:x}", CByte(255)
    Assert.Equals "ff", sb.ToString
End Sub

Public Sub TestAppendFormatDate()
    Dim dt As cDateTime
    Dim s As String
    
    Set dt = cDateTime.FromOADate(#1/1/2004 8:30:40 AM#)
    sb.AppendFormat "{0:dddd, dd MMMM yyyy HH:mm:ss}", dt
    s = sb.ToString
    Assert.Equals "Thursday, 01 January 2004 08:30:40", s
End Sub
