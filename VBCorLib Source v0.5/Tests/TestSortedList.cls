VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSortedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private list As SortedList

Public Sub Setup()
    Set list = New SortedList
End Sub

Public Sub TestEmptyList()
    Assert.Equals 0, list.Count, "Count should be 0."
End Sub

Public Sub TestCountAfterAddingOne()
    list.Add 1, "hi"
    Assert.Equals 1, list.Count, "Count should be 1."
End Sub

Public Sub TestCountAfterAddingTwo()
    list.Add 1, "hi"
    list.Add 2, "Hello"
    Assert.Equals 2, list.Count, "Count should be 2."
End Sub

Public Sub TestAddingDuplicateKey()
    list.Add 1, "hi"
    On Error GoTo errTrap
    list.Add 1, "bye"
    On Error GoTo 0
    Assert.Fail "An error should be raised."
errTrap:
    Dim ex As Exception
    Assert.IsTrue catch(ex), "A exception should be caught."
    Assert.IsTrue TypeOf ex Is ArgumentException, "Wrong exception type caught."
End Sub

Public Sub TestAddingDifferentKeyTypes()
    list.Add 1, "hi"
    On Error GoTo errTrap
    list.Add "1", "byte"
    On Error GoTo 0
    Assert.Fail "Error should be raised."
errTrap:
    Dim ex As Exception
    Assert.IsTrue catch(ex), "Exception should be caught."
    Assert.IsTrue TypeOf ex Is InvalidOperationException, "Wrong exception type."
End Sub

Public Sub TestContents()
    list.Add "1", "hi"
    list.Add "2", "Hello"
    list.Add "key", "Me"

    Assert.Equals "Me", list("key"), "List should return 'Me'."
End Sub

Public Sub TestGetByIndex()
    list.Add 3, "Me"
    list.Add 1, "hi"
    list.Add 2, "Hello"

    Assert.Equals "hi", list.GetByIndex(0)
    Assert.Equals "Hello", list.GetByIndex(1)
    Assert.Equals "Me", list.GetByIndex(2)
End Sub

Public Sub TestClone()
    Dim i As Long
    For i = 1 To 10
        list.Add i, i
    Next i
    Dim list2 As SortedList
    Set list2 = list.Clone
    For i = 1 To 10
        Assert.Equals i, list2(i)
    Next i
End Sub

Public Sub TestContains()
    Dim i As Long
    For i = 1 To 10
        list.Add i, i
    Next i
    Assert.IsTrue list.Contains(3&)
End Sub

Public Sub TestContainsValue()
    Dim i As Long
    For i = 1 To 10
        list.Add i, "Item" & i
    Next i
    Assert.IsTrue list.ContainsValue("Item3")
End Sub

Public Sub TestCopyTo()
    Dim i As Long
    For i = 1 To 10
        list.Add i, "Item" & i
    Next i
    Dim a() As DictionaryEntry
    ReDim a(9)
    list.CopyTo a, 0
    For i = 0 To 9
        Assert.Equals i + 1, a(i).Key
        Assert.Equals "Item" & i + 1, a(i).value
    Next i
End Sub

Public Sub TestGetKey()
    Dim i As Long
    For i = 1 To 10
        list.Add i, i
    Next i
    Assert.Equals 6, list.GetKey(5), "Should return 6."
End Sub

Public Sub TestIndexOfKey()
    Dim i As Long
    For i = 1 To 10
        list.Add i, i
    Next i
    Assert.Equals 5, list.IndexOfKey(6&), "Should return 5."
End Sub

Public Sub TestIndexOfValue()
    Dim i As Long
    For i = 1 To 10
        list.Add i, i
    Next i
    Assert.Equals 5, list.IndexOfValue(6&), "Should return 5."
End Sub

Public Sub TestRemove()
    Dim i As Long
    For i = 1 To 10
        list.Add i, i
    Next i
    list.Remove 5&
    Assert.Equals 9, list.Count, "Count should be 9."
    For i = 1 To 9
        Assert.Equals IIf(i < 5, i, i + 1), list.GetKey(i - 1)
    Next i
End Sub

Public Sub TestRemoveAt()
    Dim i As Long
    For i = 1 To 10
        list.Add i, i
    Next i
    list.RemoveAt 4
    Assert.Equals 9, list.Count, "Count should be 9."
    For i = 1 To 9
        Assert.Equals IIf(i < 5, i, i + 1), list.GetKey(i - 1)
    Next i
End Sub

Public Sub TestSetByIndex()
    Dim i As Long
    For i = 1 To 10
        list.Add i, i
    Next i
    list.SetByIndex 4, "Hello"
    Assert.Equals 10, list.Count, "Count should not change."
    For i = 0 To 9
        Assert.Equals IIf(i = 4, "Hello", i + 1), list.GetByIndex(i)
    Next i
End Sub

Public Sub TestTrimToSize()
    Dim i As Long
    For i = 1 To 5
        list.Add i, i
    Next i
    list.TrimToSize
    Assert.Equals 5, list.Capacity, "Capacity should be 5."
End Sub

Public Sub TestEnumerator()
    Dim i As Long
    For i = 1 To 10
        list.Add i, i
    Next i
    Dim entry As DictionaryEntry
    i = 1
    For Each entry In list
        Assert.Equals i, entry.Key
        Assert.Equals i, entry.value
        i = i + 1
    Next entry
End Sub

Public Sub TestCreateWithList()
    Dim l2 As SortedList
    Dim i As Long

    For i = 1 To 10
        list.Add i, i
    Next i
    Set l2 = newsortedlist(list)
    Assert.Equals 10, l2.Capacity, "Capacity should be 10."
    For i = 1 To 10
        Assert.Equals i, l2.GetKey(i - 1)
    Next i
End Sub

Public Sub TestPersitence()
    Dim l2 As SortedList
    Dim i As Long
    For i = 1 To 100
        list.Add i, i
    Next i
    Dim pp1 As New PropertyBag
    Dim pp2 As New PropertyBag
    pp1.WriteProperty "list", list
    pp2.Contents = pp1.Contents
    Set l2 = pp2.ReadProperty("list")

    Assert.Equals 100, l2.Count
    For i = 1 To 10
        Assert.Equals i, l2(i)
    Next i
End Sub

Public Sub TestValueAndKeyList()
    list.Add 1, "hi"
    Assert.Equals 1, list.GetKeyList(0)
    Assert.Equals "hi", list.GetValueList(0)
End Sub

Public Sub TestEnumerateKeysAndValues()
    list.Add 1, "value0"
    list.Add 2, "value1"
    
    Dim v As Variant
    Dim i As Long
    For Each v In list.Keys
        Assert.Equals list.GetKey(i), v
        i = i + 1
    Next v
    
    i = 0
    For Each v In list.Values
        Assert.Equals "value" & i, v
        i = i + 1
    Next v
End Sub


