VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestArraySort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub TestSortLongs()
    Dim a() As Long
    FillArray a, 100, 1000
    carray.Sort a
    AssertValueTypes a
End Sub

Public Sub TestSortStrings()
    Dim a() As String
    FillArray a, 100, 1000
    carray.Sort a
    AssertValueTypes a
End Sub

Public Sub TestSortObjects()
    Dim a() As Int32
    Dim i As Long
    
    ReDim a(0 To 100)
    Rnd -13
    For i = 0 To UBound(a)
        Set a(i) = New Int32
        a(i).mValue = Int(Rnd * 100)
    Next i
    
    carray.Sort a
    
    For i = 0 To UBound(a) - 1
        Assert.IsTrue a(i).mValue <= a(i + 1).mValue, "Elements out of order at " & i & " and " & i + 1 & "."
    Next i
    
End Sub

Public Sub TestSortInteger()
    Dim a() As Integer
    FillArray a, 100, 1000
    carray.Sort a
    AssertValueTypes a
End Sub

Public Sub TestSortBytes()
    Dim a() As Byte
    FillArray a, 100, 255
    carray.Sort a
    AssertValueTypes a
End Sub

Public Sub TestSortKeysAndItems()
    Dim k() As Long
    Dim i() As Byte

    FillArray k, 100, 1000
    FillArray i, 100, 255
    carray.SortKey k, i
    AssertValueTypes k
    AssertValueTypes i
End Sub

Public Sub TestWithComparer()
    Dim k() As Long
    Dim i() As Byte

    FillArray k, 100, 1000
    FillArray i, 100, 255
    carray.SortKey k, i, comparer.Default
    AssertValueTypes k
    AssertValueTypes i
End Sub

Public Sub TestWithStringComparer()
    Dim s(100) As String
    Dim i As Long
    
    Rnd -13
    For i = 0 To 100
        s(i) = "Value " & Chr$(Int(Rnd * 255))
    Next i
    carray.Sort s, comparer.Default
    AssertValueTypes s
End Sub

Public Sub TestSortVariants()
    Dim a() As Variant
    FillArray a, 100, 1000
    carray.Sort a
    AssertValueTypes a
End Sub

Public Sub TestGeneralSort()
    Dim a() As Long
    FillArray a, 100, 1000
    carray.Sort a, comparer.Default
    AssertValueTypes a
End Sub

Public Sub TestNonComparableElements()
    Dim a() As Variant
    Dim b() As Long
    FillArray a, 100, 1000
    a(39) = b
    On Error GoTo errTrap
    carray.Sort a
    On Error GoTo 0
    Assert.Fail "Should have thrown an error."
errTrap:
    Dim ex As Exception
    Assert.IsTrue catch(ex), "Should have caught an exception."
    Assert.IsTrue TypeOf ex Is InvalidOperationException, "Wrong exception type caught."
End Sub

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

Public Sub TestSortSubArray()
    Dim a() As Long
    FillArray a, 100, 1000
    carray.SortEx a, 26, 24
    
    Dim i As Long
    Dim j As Long
    Rnd -13
    For i = 0 To 98
        j = Int(Rnd * 1000)
        If i <= 25 Or i >= 50 Then
            Assert.Equals j, a(i), "Sorted outside sub-array."
        Else
            ' we don't want to bleed into the unsorted list with a(i + 1).
            If i < 49 Then Assert.IsTrue a(i) <= a(i + 1), "sub-array not sorted."
        End If
    Next i
End Sub


Private Sub FillArray(ByRef arr As Variant, ByVal size As Long, ByVal maxValue As Long)
    Dim i As Long

    ReDim arr(0 To size - 1)
    Rnd -13
    For i = LBound(arr) To UBound(arr)
        arr(i) = Int(Rnd * maxValue)
    Next i
End Sub

Private Sub AssertValueTypes(ByRef arr As Variant)
    Dim i As Long
    
    For i = LBound(arr) To UBound(arr) - 1
        Assert.IsTrue arr(i) <= arr(i + 1), "Elementout of order at " & i & "."
    Next i
End Sub


