VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestUTF8GetByteCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestUTF8GetByteCount
'
Option Explicit

Private e As UTF8Encoding

Public Sub Setup()
    Set e = NewUTF8Encoding(False, True)
End Sub

Public Sub TestFirstSequenceForLength()
    Assert.Equals 1, e.GetByteCount(ChrW$(0))
    Assert.Equals 2, e.GetByteCount(ChrW$(&H80))
    Assert.Equals 3, e.GetByteCount(ChrW$(&H800))
    Assert.Equals 4, e.GetByteCount(ChrW$(&HD800) & ChrW$(&HDC00))
End Sub

Public Sub TestLastSequenceForLength()
    Assert.Equals 1, e.GetByteCount(ChrW$(&H7F))
    Assert.Equals 2, e.GetByteCount(ChrW$(&H7FF))
    Assert.Equals 3, e.GetByteCount(ChrW$(&HFFFF))
End Sub

Public Sub TestFirstSequenceArray()
    Assert.Equals 100, e.GetByteCount(String$(100, ChrW$(0)))
    Assert.Equals 200, e.GetByteCount(String$(100, ChrW$(&H80)))
    Assert.Equals 300, e.GetByteCount(String$(100, ChrW$(&H800)))
End Sub

Public Sub TestLastSequenceArray()
    Assert.Equals 100, e.GetByteCount(String$(100, ChrW$(&H7F)))
    Assert.Equals 200, e.GetByteCount(String$(100, ChrW$(&H7FF)))
    Assert.Equals 300, e.GetByteCount(String$(100, ChrW$(&HFFFF)))
End Sub

Public Sub TestFromChars()
    Assert.Equals 100, e.GetByteCount(GetChars(0))
    Assert.Equals 200, e.GetByteCount(GetChars(&H80))
    Assert.Equals 300, e.GetByteCount(GetChars(&H800))
End Sub

Public Sub TestMalformedSurrogate()
    Dim ex As Exception
    
    On Error GoTo errTrap
    e.GetByteCount ChrW$(&HD800) & Chr$(0)
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Assert.IsTrue catch(ex), "An exception should be caught."
End Sub

Public Sub TestMissingLastSurrogateWithError()
    Dim ex As Exception
    
    On Error GoTo errTrap
    e.GetByteCount ChrW$(&HD800)
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Assert.IsTrue catch(ex), "An exception should be caught."
End Sub

Public Sub TestMissingLastSurrogate()
    Dim e As New UTF8Encoding
    
    Assert.Equals 3, e.GetByteCount(ChrW$(&HD801))
End Sub


Private Function GetChars(ByVal Char As Integer) As Integer()
    Dim ret() As Integer
    Dim i As Long
    
    ReDim ret(99)
    For i = 0 To 99
        ret(i) = Char
    Next i
    GetChars = ret
End Function
