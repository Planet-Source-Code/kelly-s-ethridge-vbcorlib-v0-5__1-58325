VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestArrayReverse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub TestAllLongs()
    Dim a(100) As Long
    Dim i As Long
    
    For i = 0 To 100
        a(i) = i
    Next i
    carray.Reverse a
    For i = 0 To 100
        Assert.Equals 100 - i, a(i)
    Next i
End Sub

Public Sub TestAllStrings()
    Dim a(100) As String
    Dim i As Long
    
    For i = 0 To 100
        a(i) = i
    Next i
    carray.Reverse a
    For i = 0 To 100
        Assert.Equals CStr(100 - i), a(i)
    Next i
End Sub

Public Sub TestSubArrayLongs()
    Dim a(100) As Long
    Dim i As Long
    
    For i = 0 To 100
        a(i) = i
    Next i
    carray.Reverse a, 26, 24
    For i = 0 To 100
        If i <= 25 Or i >= 50 Then
            Assert.Equals i, a(i), "Reversed outside subarray."
        Else
            Assert.Equals 100 - i - 25, a(i), "Subarray not reversed."
        End If
    Next i
End Sub

Public Sub TestLengthTooLong()
    Dim a(100)
    
    On Error GoTo errTrap
    carray.Reverse a, 0, 102
    On Error GoTo 0
    Assert.Fail "An exception should have been thrown."
errTrap:
    Dim ex As Exception
    Assert.IsTrue catch(ex), "An exception should have been caught."
    Assert.IsTrue TypeOf ex Is ArgumentException, "Wrong exception type caught."
End Sub
