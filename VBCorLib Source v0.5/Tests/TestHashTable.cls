VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestHashTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestHashTable
'
Option Explicit
Private t As Hashtable

Public Sub Setup()
    Set t = New Hashtable
End Sub

Public Sub TestEmptyTable()
    Assert.Equals 0, t.Count, "Count should return 0."
End Sub

Public Sub TestCountAfterAddingOne()
    t.Add 1, 1
    Assert.Equals 1, t.Count, "Count should return 1."
End Sub

Public Sub TestContents()
    t.Add 1, 1
    Assert.Equals 1, t(1), "table should return '1'."
End Sub

Public Sub TestContentsAfterAddingTwo()
    t.Add 1, 1
    t.Add 2, 2
    Assert.Equals 1, t(1)
    Assert.Equals 2, t(2)
End Sub

Public Sub Test1000Items()
    Add1000Strings
    
    Dim i As Long
    For i = 1 To 1000
        Assert.Equals "Value " & i, t("Key " & i)
    Next i
End Sub

Public Sub TestChangeExisting()
    t.Add "Key", "first"
    Assert.Equals "first", t("Key")
    t("Key") = "second"
    Assert.Equals "second", t("Key")
End Sub

Public Sub TestSetsNew()
    Assert.Equals Empty, t("Key")
    t("Key") = "first"
    Assert.Equals "first", t("Key")
End Sub

Public Sub TestContainsKey()
    Assert.IsFalse t.ContainsKey("Key")
    t("Key") = "first"
    Assert.IsTrue t.ContainsKey("Key")
End Sub

Public Sub TestRemove()
    Add1000Longs
    
    Dim i As Long
    For i = 20 To 60
        t.Remove i
    Next i
    
    For i = 1 To 1000
        If i >= 20 And i <= 60 Then
            Assert.Equals Empty, t(i)
        Else
            Assert.Equals i, t(i)
        End If
    Next i
End Sub

Public Sub TestContainsValue()
    Add1000Strings
    
    Dim i As Long
    Dim sw As StopWatch
    Set sw = New StopWatch
    sw.Reset
    For i = 1 To 1000
        Assert.IsTrue t.ContainsValue("Value " & i)
    Next i
    Assert.IsFalse t.ContainsValue("Not here.")
End Sub

Public Sub TestClone()
    Add1000Strings
    
    Dim i As Long
    Dim t2 As Hashtable
    Set t2 = t.Clone
    For i = 1 To 1000
        Assert.Equals "Value " & i, t2("Key " & i)
    Next i
End Sub

Public Sub TestEnumerator()
    Add1000Longs
    Dim de As DictionaryEntry
    Dim col As New Collection
    For Each de In t.GetEnumerator
        col.Add de.value, CStr(de.Key)
    Next de
    Assert.Equals 1000, col.Count
End Sub

Public Sub TestKeysCollection()
    Add1000Longs
    Assert.Equals 1000, t.Keys.Count
    Dim v As Variant
    Dim col As New Collection
    For Each v In t.Keys
        col.Add v, CStr(v)
    Next v
    Assert.Equals 1000, col.Count
End Sub

Public Sub TestValuesCollection()
    Add1000Longs
    Assert.Equals 1000, t.Values.Count
    Dim v As Variant
    Dim col As New Collection
    For Each v In t.Values
        col.Add v, CStr(v)
    Next v
    Assert.Equals 1000, col.Count
End Sub

Public Sub TestClear()
    Add1000Longs
    t.Clear
    Assert.Equals 0, t.Count, "Count should return 0."
End Sub




Private Sub Add1000Longs()
    Dim i As Long
    For i = 1 To 1000
        t.Add i, i
    Next i
End Sub
Private Sub Add1000Strings()
    Dim i As Long
    For i = 1 To 1000
        t.Add "Key " & i, "Value " & i
    Next i
End Sub
