VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Q As Queue

Public Sub Setup()
    Set Q = New Queue
End Sub

Public Sub TestEmptyQueue()
    Assert.Equals 0, Q.Count, "Count should return 0."
End Sub

Public Sub TestCountAfterAddingOne()
    Q.Enqueue 1
    Assert.Equals 1, Q.Count, "Count should return 1."
End Sub

Public Sub TestContents()
    Q.Enqueue 1
    Assert.Equals 1, Q.Dequeue, "Deque should return 1."
    Assert.Equals 0, Q.Count, "Count should return 0."
End Sub

Public Sub TestQueueMany()
    Dim i As Long
    For i = 1 To 100
        Q.Enqueue i
    Next i
    
    For i = 1 To 100
        Assert.Equals i, Q.Dequeue
    Next i
End Sub

Public Sub TestClear()
    Dim i As Long
    For i = 1 To 100
        Q.Enqueue i
    Next i
    Q.Clear
    Assert.Equals 0, Q.Count, "Count should return 0."
    For i = 101 To 200
        Q.Enqueue i
    Next i
    For i = 101 To 200
        Assert.Equals i, Q.Dequeue
    Next i
End Sub

Public Sub TestClone()
    Dim i As Long
    For i = 1 To 10
        Q.Enqueue i
    Next i
    
    Dim Clone As Queue
    Set Clone = Q.Clone
    For i = 1 To 10
        Assert.Equals i, Clone.Dequeue
    Next i
End Sub

Public Sub TestContains()
    Dim i As Long
    For i = 1 To 100
        Q.Enqueue i
    Next i
    
    Assert.IsTrue Q.Contains(99&), "q should contain 99."
End Sub

Public Sub TestContainsWithComparer()
    Dim i As Long
    For i = 1 To 100
        Q.Enqueue i
    Next i
    
    Assert.IsTrue Q.Contains(99&, comparer.Default), "q should contain 99."
End Sub

Public Sub TestCopyTo()
    Dim i As Long
    For i = 1 To 100
        Q.Enqueue i
    Next i
    
    Dim a(99) As Long
    Q.CopyTo a, 0
    For i = 1 To 100
        Assert.Equals i, a(i - 1)
    Next i
End Sub

Public Sub TestCopyToWrapping()
    ' required INDEBUG=1 VBCorLib compiler constant.
    Dim i As Long
    For i = 1 To 8
        Q.Enqueue i
        Q.Dequeue
    Next i
    For i = 1 To 14
        Q.Enqueue i
    Next i
    
    Dim a(-5 To 8) As Long
    Q.CopyTo a, -5
    For i = -5 To 8
        Assert.Equals i + 6, a(i)
    Next i
End Sub

Public Sub TestEnumerator()
    Dim i As Long
    For i = 1 To 100
        Q.Enqueue i
    Next i
    Dim v As Variant
    i = 1
    For Each v In Q
        Assert.Equals i, v
        i = i + 1
    Next v
End Sub

Public Sub TestPeek()
    Q.Enqueue 1
    Assert.Equals 1, Q.Peek, "Peek should return 1."
    Assert.Equals 1, Q.Count, "Count should not change."
End Sub

Public Sub TestToArray()
    Dim i As Long
    For i = 1 To 10
        Q.Enqueue 1
        Q.Dequeue
    Next i
    For i = 11 To 20
        Q.Enqueue i
    Next i
    Dim a() As Variant
    a = Q.ToArray
    Assert.Equals Q.Count, carray.GetLength(a), "Wrong array length."
    
    For i = 11 To 20
        Assert.Equals i, a(i - 11)
    Next i
End Sub

Public Sub TestEmptyToArray()
    Dim a() As Variant
    a = Q.ToArray
    Assert.IsFalse carray.IsNull(a), "Array should not be null."
    Assert.Equals 0, carray.GetLength(a), "Array should be zero length."
End Sub

Public Sub TestTrimToSize()
    ' requires INDEBUG=1 VBCorLib compiler constant.
    Dim i As Long
    For i = 1 To 8
        Q.Enqueue i
    Next i
    Q.Dequeue
    Q.TrimToSize
    
    For i = 2 To 8
        Assert.Equals i, Q.Dequeue, "Queue contents should not change."
    Next i
End Sub

Public Sub TestTrimToSize2()
    Dim i As Long
    For i = 1 To 16
        Q.Enqueue i
    Next i
    For i = 1 To 5
        Q.Dequeue
    Next i
    For i = 17 To 21
        Q.Enqueue i
    Next i
    
    Q.TrimToSize
    
    For i = 6 To 21
        Assert.Equals i, Q.Dequeue
    Next i
End Sub

Public Sub TestPersistence()
    Dim i As Long
    For i = 1 To 100
        Q.Enqueue i
    Next i
    Dim pp1 As New PropertyBag
    Dim pp2 As New PropertyBag
    Dim q2 As Queue
    
    pp1.WriteProperty "q", Q
    pp2.Contents = pp1.Contents
    Set q2 = pp2.ReadProperty("q")
    
    For i = 1 To 100
        Assert.Equals i, q2.Dequeue
    Next i
End Sub
