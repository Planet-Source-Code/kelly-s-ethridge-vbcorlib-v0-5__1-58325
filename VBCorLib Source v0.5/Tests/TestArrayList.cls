VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private list As ArrayList

Public Sub Setup()
    Set list = New ArrayList
End Sub

Public Sub TestEmptyList()
    Assert.Equals 0, list.Count
End Sub

Public Sub TestAfterAddingOne()
    list.Add 1
    Assert.Equals 1, list.Count
End Sub

Public Sub TestAfterAddingTwo()
    list.Add 1
    list.Add 2
    Assert.Equals 2, list.Count
End Sub

Public Sub TestRetrieveAfterAdding()
    list.Add 1
    Assert.Equals 1, list(0)
End Sub

Public Sub TestAddRangeFromCollection()
    Dim col As New Collection
    col.Add 1
    col.Add 2
    col.Add 3
    
    list.AddRange col
    Assert.Equals 3, list.Count
    Assert.Equals 1, list(0)
    Assert.Equals 2, list(1)
    Assert.Equals 3, list(2)
End Sub

Public Sub TestAddRangeFromArray()
    Dim a() As Long
    a = carray.NewArray(ciLong, 1, 2, 3)
    
    list.AddRange a
    Assert.Equals 3, list.Count
    Assert.Equals 1, list(0)
    Assert.Equals 2, list(1)
    Assert.Equals 3, list(2)
End Sub

Public Sub TestAddRangeFromICollection()
    Dim a As New ArrayList
    a.AddRange carray.NewArray(ciLong, 1, 2, 3)
    
    list.AddRange a
    Assert.Equals 3, list.Count
    Assert.Equals 1, list(0)
    Assert.Equals 2, list(1)
    Assert.Equals 3, list(2)
End Sub

Public Sub TestSetItem()
    list.AddRange carray.NewArray(ciLong, 1, 2, 3, 4, 5)
    list(2) = "Hello"
    Assert.Equals 5, list.Count
    Assert.Equals 1, list(0)
    Assert.Equals 2, list(1)
    Assert.Equals "Hello", list(2)
    Assert.Equals 4, list(3)
    Assert.Equals 5, list(4)
End Sub

Public Sub TestClear()
    list.Add 1
    list.Clear
    Assert.Equals 0, list.Count
End Sub

Public Sub TestCopyTo()
    Dim a(2) As String
    list.Add "Item1"
    list.Add "Item2"
    list.Add "Item3"
    list.CopyTo a
    Assert.Equals "Item1", a(0)
    Assert.Equals "Item2", a(1)
    Assert.Equals "Item3", a(2)
End Sub

Public Sub TestCopyToIndexOffset()
    Dim a(5) As String
    list.Add "Item1"
    list.Add "Item2"
    list.Add "Item3"
    list.CopyTo a, 2
    Assert.Equals "Item1", a(2)
    Assert.Equals "Item2", a(3)
    Assert.Equals "Item3", a(4)
End Sub

Public Sub TestEnumerator()
    list.AddRange carray.NewArray(ciString, "Item1", "Item2", "Item3")
    
    Dim v As Variant
    For Each v In list
        Assert.Equals "Item", Left$(v, 4)
    Next v
End Sub

Public Sub TestInsertRangeArray()
    list.AddRange carray.NewArray(ciLong, 1, 2, 3, 4)
    list.Insertrange 2, carray.NewArray(ciString, "Item1", "Item2", "Item3")
    AssertInsert
End Sub

Public Sub TestInsertRangecollection()
    list.AddRange carray.NewArray(ciLong, 1, 2, 3, 4)
    Dim col As New Collection
    col.Add "Item1"
    col.Add "Item2"
    col.Add "Item3"
    list.Insertrange 2, col
    AssertInsert
End Sub

Public Sub TestContains()
    list.AddRange carray.NewArray(ciLong, 1, 2, 3)
    Assert.IsTrue list.Contains(1&), "list should contain 1."
End Sub

Public Sub TestNotContains()
    list.AddRange carray.NewArray(ciLong, 1, 2, 3)
    Assert.IsFalse list.Contains(4), "list should not contain 4."
End Sub



Private Sub AssertInsert()
    Assert.Equals 7, list.Count, "Count should return 7."
    Assert.Equals 1, list(0), "list should return '1'."
    Assert.Equals 2, list(1), "list should return '2'."
    Assert.Equals "Item1", list(2), "list should return 'Item1'."
    Assert.Equals "Item2", list(3), "list should return 'Item2'."
    Assert.Equals "Item3", list(4), "list should return 'Item3'."
    Assert.Equals 3, list(5), "list should return '3'."
    Assert.Equals 4, list(6), "list should return '4'."
End Sub


