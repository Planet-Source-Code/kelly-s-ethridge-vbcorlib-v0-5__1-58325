VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestArrayCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestArrayCopy
'
Option Explicit

Public Sub TestCopyLongs()
    Dim src() As Long
    Dim dst() As Long
    Dim i As Long
    
    ReDim src(0 To 1000000)
    ReDim dst(0 To UBound(src))
    For i = 0 To UBound(src)
        src(i) = i
    Next i
    
    carray.Copy src, dst, UBound(src) + 1
    
    For i = 0 To 100
        Assert.Equals i, dst(i)
    Next i
End Sub

Public Sub TestCopyStrings()
    Dim src() As String
    Dim dst() As String
    Dim i As Long
    
    ReDim src(0 To 100)
    ReDim dst(0 To UBound(src))
    For i = 0 To UBound(src)
        src(i) = CStr(i)
    Next i
    
    carray.Copy src, dst, UBound(src) + 1
    
    For i = 0 To 100
        Assert.Equals CStr(i), dst(i)
    Next i
End Sub

Public Sub TestCopySubArray()
    Dim src() As Long
    Dim dst() As Long
    Dim i As Long
    
    ReDim src(0 To 100)
    ReDim dst(0 To 100)
    For i = 0 To 100
        src(i) = i
    Next i
    
    carray.CopyEx src, 20, dst, 0, 20
    For i = 0 To 100
        If i >= 20 Then
            Assert.Equals 0, dst(i)
        Else
            Assert.Equals i + 20, dst(i)
        End If
    Next i

End Sub

Public Sub TestCopyObjects()
    Dim src() As Int32
    Dim dst() As Int32
    Dim i As Long
    
    ReDim src(0 To 100)
    ReDim dst(0 To 100)
    For i = 0 To 100
        Set src(i) = New Int32
        src(i).mValue = i
    Next i
    
    carray.Copy src, dst, 101
    For i = 0 To 100
        Assert.Equals i, dst(i).mValue
    Next i
End Sub

Public Sub TestObjectsToSelf()
    Dim src() As Int32
    Dim dst() As Int32
    Dim i As Long
    
    ReDim src(0 To 20)
    ReDim dst(0 To 20)
    For i = 0 To 20
        Set src(i) = New Int32
        src(i).mValue = i
    Next i
    
    carray.CopyEx src, 0, src, 10, 10
    
    For i = 0 To 9
        Assert.Equals src(i), src(i + 10)
    Next i
End Sub

Public Sub TestLongToString()
    Dim src() As Long
    Dim dst() As String
    Dim i As Long
    
    ReDim src(0 To 100)
    ReDim dst(0 To 100)
    For i = 0 To 100
        src(i) = i
    Next i
    
    carray.Copy src, dst, 101
    For i = 0 To 100
        Assert.Equals CStr(i), dst(i)
    Next i
End Sub

Public Sub TestIndexes()
    Dim src() As String
    Dim dst() As String
    
    ReDim src(0 To 4)
    ReDim dst(5 To 15)
    
    Dim i As Long
    For i = 0 To 4
        src(i) = i
    Next i
    
    carray.CopyEx src, 0, dst, 5, 5
    
    For i = 5 To 9
        Assert.Equals CStr(i - 5), dst(i)
    Next i
End Sub

Public Sub TestVariantToObjects()
    Dim src() As Variant
    Dim dst() As Int32
    
    ReDim src(0 To 5)
    ReDim dst(0 To 5)
    Dim i As Long
    For i = 0 To 5
        Set src(i) = New Int32
        Set dst(i) = New Int32
    Next i
    
    carray.Copy src, dst, 6
    
    For i = 0 To 5
        Assert.Equals src(i), dst(i)
    Next i
End Sub

Public Sub TestArrayTypeMismatch()
    Dim src(0) As Double
    Dim dst(0) As Long
    
    On Error GoTo errTrap
    carray.Copy src, dst, 1
    On Error GoTo 0
    Assert.Fail "Error should have been raised."
    
errTrap:
    Dim ex As Exception
    Assert.IsTrue catch(ex), "An exception should have been caught."
    Assert.IsTrue TypeOf ex Is ArrayTypeMismatchException, "Wrong exception type caught."
    
End Sub

'Public Sub TestUserDefined()
'    Dim src(10) As TestType
'    Dim dst(10) As TestType
'    Dim i As Long
'
'    For i = 0 To 10
'        With src(i)
'            .value = i
'            .strValue = "Item " & i
'            Set .objValue = New Int32
'            .objValue.mValue = i
'        End With
'    Next i
'
'    carray.Copy src, dst, 11
'
'    For i = 0 To 10
'        With dst(i)
'            Assert.Equals i, .value
'            Assert.Equals "Item " & i, .strValue
'            Assert.IsNotNothing .objValue
'        End With
'    Next i
'
'End Sub

