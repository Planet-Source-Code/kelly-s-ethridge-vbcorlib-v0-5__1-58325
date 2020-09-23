VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestArrayBinarySearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub TestSearchSubArray()
    Dim a() As Long
    Dim i As Long
    
    ReDim a(0 To 100)
    For i = 0 To 100
        If i >= 26 And i <= 49 Then a(i) = i
    Next i
    
    i = carray.BinarySearch(a, 27&, 26, 24)
    Assert.Equals 27, i
    
End Sub

Public Sub TestNotInSubArray()
    Dim a() As Long
    Dim i As Long
    
    ReDim a(0 To 100)
    For i = 0 To 100
        If i >= 26 And i <= 49 Then a(i) = i
    Next i
    
    i = carray.BinarySearch(a, 1&, 26, 24, comparer.Default)
    Assert.Equals -27, i
    
End Sub

Public Sub TestObjectBinarySearch()
    Dim a() As Int32
    Dim i As Long
    
    ReDim a(0 To 100)
    For i = 0 To 100
        Set a(i) = New Int32
        a(i).mValue = i
    Next i
    Dim oInt As New Int32
    oInt.mValue = 5
    i = carray.BinarySearch(a, oInt)
    Assert.Equals 5, i, "Should return 5."
End Sub

Public Sub TestStringBinarySearch()
    Dim a() As String
    Dim i As Long
    
    ReDim a(0 To 100)
    For i = 0 To 100
        a(i) = "Item " & Chr$(i)
    Next i
    
    i = carray.BinarySearch(a, "Item " & Chr$(5))
    Assert.Equals 5, i, "Should return 5."
End Sub

Public Sub TestVariantBinarySearch()
    Dim a() As Variant
    Dim i As Long
    
    ReDim a(0 To 100)
    For i = 0 To 100
        a(i) = "Item " & Chr$(i)
    Next i
    a(0) = Null
    a(1) = Empty
    Assert.Equals 0, carray.BinarySearch(a, Null), "Should return 0."
    Assert.Equals 1, carray.BinarySearch(a, Empty), "Should return 1."
End Sub

'Public Sub TestFailedSystemComparer()
'    Dim a(0) As Variant
'
'    On Error GoTo errTrap
'    carray.BinarySearch a, "0"
'    On Error GoTo 0
'    Assert.Fail "An error should have been raised."
'
'errTrap:
'    Dim ex As Exception
'    Assert.IsTrue catch(ex, Err), "An exception should have been caught."
'    Assert.IsTrue TypeOf ex Is InvalidOperationException Or TypeOf ex Is InvalidCastException, "Wrong exception type was caught."
'End Sub

Public Sub TestLengthTooLong()
    Dim a(10) As Long
    
    On Error GoTo errTrap
    carray.BinarySearch a, 1, 0, 12
    On Error GoTo 0
    Assert.Fail "An exception should have been raised."
    
errTrap:
    Dim ex As Exception
    Assert.IsTrue catch(ex), "An exception should have been caught."
    Assert.IsTrue TypeOf ex Is ArgumentException, "Wrong exception type caught."
End Sub

'#If STRICT Then
'Public Sub TestStrictOn()
'    Dim a(0) As Long
'    On Error GoTo errTrap
'    carray.BinarySearch a, "1"
'    On Error GoTo 0
'    Assert.Fail "An error should have been raised."
'errTrap:
'    Dim ex As Exception
'    Assert.IsTrue catch(ex), "An exception should have been thrown."
'    Assert.IsTrue TypeOf ex Is InvalidCastException, "Wrong exception caught."
'End Sub
'#Else
'Public Sub TestStrictOff()
'    Dim a(100) As Long
'    Dim i As Long
'    For i = 0 To 100
'        a(i) = i
'    Next i
'    Assert.Equals 5, carray.BinarySearch(a, "5")
'End Sub
'#End If
