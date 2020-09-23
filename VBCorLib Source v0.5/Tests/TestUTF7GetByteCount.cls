VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestUTF7GetByteCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestUTF7GetByteCount
'
Option Explicit

Private e As UTF7Encoding

Public Sub Setup()
    Set e = New UTF7Encoding
End Sub

Public Sub TestDirectlyEncodedCharacters()
    Assert.Equals 1, e.GetByteCount("a")
    Assert.Equals 26, e.GetByteCount(String$(26, "1"))
End Sub

Public Sub TestNonDirectUnicode()
    Assert.Equals 5, e.GetByteCount("!")
    Assert.Equals 16, e.GetByteCount(String$(5, "!"))
    Assert.Equals 6, e.GetByteCount("!" & vbCr)
    Assert.Equals 10, e.GetByteCount("+!+")
    Assert.Equals 4, e.GetByteCount("++")
End Sub

Public Sub TestMixture()
    Assert.Equals 13, e.GetByteCount("a!'=2")
End Sub

Public Sub TestPlusMinus()
    Assert.Equals 3, e.GetByteCount("+-")
End Sub

Public Sub TestSimpleCombo2()
    Dim ch() As Integer
    ch = carray.NewArray(ciInteger, &H48, &H69, &H20, &H4D, &H6F, &H6D, &H20, &H2D, &H263A, &H2D, &H21)
    Assert.Equals 19, e.GetByteCount(ch)
End Sub

Public Sub TestJapaneseWord_nihongo()
    Dim ch() As Integer
    ch = carray.NewArray(ciInteger, &H65E5, &H672C, &H8A9E)
    Assert.Equals 10, e.GetByteCount(ch)
End Sub

Public Sub TestComplexCombo()
    Assert.Equals 15, e.GetByteCount(ChrW$(&H65E5) & "hi!" & ChrW$(&H8A9E))
End Sub

Public Sub TestEncoder()
    Dim c As Encoder
    Dim ch(100) As Integer
    Dim i As Long
    
    For i = 0 To 100
        ch(i) = &H65E5
    Next i
    Set c = e.GetEncoder
    
    Assert.Equals 270, c.GetByteCount(ch, 0, 101, False)
    Assert.Equals 272, c.GetByteCount(ch, 0, 101, True)
    
    For i = 0 To 100
        ch(i) = IIf(i < 50, &H65E5, 97)
    Next i
    
    Assert.Equals 187, c.GetByteCount(ch, 0, 101, False), "Should return 187 when Flush = False."
    Assert.Equals 187, c.GetByteCount(ch, 0, 101, True), "Should return 187 when Flush = True."
End Sub

