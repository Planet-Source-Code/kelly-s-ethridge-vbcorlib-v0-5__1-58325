VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestcArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub TestLengthOf1DArray()
    Dim arr(1 To 10) As Long
    Assert.Equals 10, carray.GetLength(arr), "Length should be 10."
End Sub

Public Sub TestLengthof2DArray()
    Dim arr(1 To 10, 1 To 15) As Long
    Assert.Equals 150, carray.GetLength(arr), "Length should be 100."
End Sub

Public Sub TestSpecificDimension()
    Dim arr(1 To 10, 1 To 15)
    Assert.Equals 15, carray.GetLength(arr, 2)
End Sub

Public Sub TestGetLengthInVariantByRef()
    Dim arr As Variant
    ReDim arr(1 To 10) As Long
    Assert.Equals 10, carray.GetLength(arr), "Lenght should be 10."
End Sub

Public Sub TestGetLengthRankError()
    Dim arr(1 To 10, 1 To 15)
    On Error GoTo errTrap
    carray.GetLength arr, 3
    On Error GoTo 0
    Assert.Fail "An exception should have been thrown."
    
errTrap:
    Dim ex As Exception
    Assert.IsTrue catch(ex), "An exception should have been caught."
    Assert.IsTrue TypeOf ex Is IndexOutOfRangeException, "Incorrect exception was caught."
End Sub

Public Sub TestIsArrayError()
    Dim NotAnArray As Long
    Dim ex As Exception
    
    On Error GoTo errTrap
    carray.GetLength NotAnArray
    On Error GoTo 0
    Assert.Fail "An exception should have been thrown."

errTrap:
    Assert.IsTrue catch(ex), "An exception should have been caught."
    Assert.IsTrue TypeOf ex Is ArgumentException, "Incorrect exception caught."
End Sub

Public Sub TestGetLengthNullArray()
    Dim ex As Exception
    Dim NullArray() As Long
    
    On Error GoTo errTrap
    carray.GetLength NullArray
    On Error GoTo 0
    Assert.Fail "An exception should have been thrown."
    
errTrap:
    Assert.IsTrue catch(ex), "An exception should have been caught."
    Assert.IsTrue TypeOf ex Is ArgumentNullException, "Incorrect exception caught."
End Sub

Public Sub TestIsNullArrayTrue()
    Dim NullArray() As Long
    Assert.IsTrue carray.IsNull(NullArray), "IsNull should return True for 'NullArray'."
    
    Dim NullObjArray() As Object
    Assert.IsTrue carray.IsNull(NullObjArray), "IsNull should return True for 'NullObjArray'."
    
    ReDim NullArray(0)
    Assert.IsFalse carray.IsNull(NullArray), "IsNull should return false for 'NullArray'."
End Sub

Public Sub TestGetRankFor1DArray()
    Dim arr(0) As Long
    Assert.Equals 1, carray.GetRank(arr), "Rank should be 1."
End Sub

Public Sub TestGetRankFor2DArray()
    Dim arr(0, 0) As Long
    Assert.Equals 2, carray.GetRank(arr), "Rank should be 2."
End Sub

Public Sub TestGetRankFromVariantByRef()
    Dim arr As Variant
    ReDim arr(0, 0, 0) As Long
    Assert.Equals 3, carray.GetRank(arr), "Rank should be 3."
End Sub

Public Sub TestGetRankNullArray()
    Dim ex As Exception
    Dim NullArray() As Long
    
    On Error GoTo errTrap
    carray.GetRank NullArray
    On Error GoTo 0
    Assert.Fail "An exception should have been thrown."
    
errTrap:
    Assert.IsTrue catch(ex), "An exception should have been caught."
    Assert.IsTrue TypeOf ex Is ArgumentNullException, "Incorrect exception caught."
End Sub

Public Sub TestGetRankNotAnArray()
    Dim NotAnArray As Long
    Dim ex As Exception
    On Error GoTo errTrap
    carray.GetRank NotAnArray
    On Error GoTo 0
    Assert.Fail "An exception should have been thrown."
    
errTrap:
    Assert.IsTrue catch(ex), "An exception should have been caught."
    Assert.IsTrue TypeOf ex Is ArgumentException, "Incorrect exception caught."
End Sub



Public Sub TestClearAllLongs()
    Dim a() As Long
    ReDim a(100)
    Dim i As Long
    
    For i = 0 To 100
        a(i) = i
    Next i
    carray.Clear a, 0, 101
    For i = 0 To 100
        Assert.Equals 0, a(i)
    Next i
End Sub

Public Sub TestClearSomeLongs()
    Dim a() As Long
    ReDim a(100)
    Dim i As Long
    
    For i = 0 To 100
        a(i) = i
    Next i
    carray.Clear a, 25, 30
    For i = 0 To 100
        If i < 25 Or i > 54 Then
            Assert.Equals i, a(i)
        Else
            Assert.Equals 0, a(i)
        End If
    Next i
End Sub

'Public Sub TestClearSomeUDTs()
'    Dim a() As TestType
'    ReDim a(-1 To 1)
'    Dim i As Long
'
'    For i = -1 To 1
'        Set a(i).objValue = New Int32
'    Next i
'
'    carray.Clear a, 0, 1
'    Assert.IsNotNothing a(-1).objValue
'    Assert.IsNothing a(0).objValue
'    Assert.IsNotNothing a(1).objValue
'End Sub

