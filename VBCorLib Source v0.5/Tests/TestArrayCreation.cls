VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestArrayCreation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub TestCreateLong()
    Dim a() As Long
    a = cArray.NewArray(ciLong, 1, 2, 3, 4, 5)
    Dim i As Long
    For i = 0 To 4
        Assert.Equals i + 1, a(i)
    Next i
End Sub

Public Sub TestCreateObject()
    Dim a() As Int32
    a = cArray.NewArray(ciObject, New Int32, New Int32, New Int32)
    Dim i As Long
    For i = 0 To 2
        Assert.IsNotNothing a(i)
        Assert.IsTrue TypeOf a(i) Is Int32
    Next i
End Sub

Public Sub TestEmptyArray()
    Dim a() As Long
    a = cArray.CreateInstance(ciLong)
    Assert.Equals -1, UBound(a)
End Sub

Public Sub TestCreateLong2()
    Dim a() As Long
    Dim i As Long
    
    a = cArray.NewArray(ciLong, 89, 191, 409, 907, 1973, 4327, 9461, 20743, 45481, 99607, 218287)
    Assert.Equals 10, UBound(a), "NewArray should return 11 elements."
    
    For i = 0 To UBound(a)
        Assert.Equals Choose(i + 1, 89, 191, 409, 907, 1973, 4327, 9461, 20743, 45481, 99607, 218287), a(i)
    Next i
End Sub

Public Sub TestCreateString()
    Dim a() As String
    Dim i As Long
    
    a = cArray.NewArray(ciString, "hi", "bye", "hi again")
    Assert.Equals 3, cArray.GetLength(a)
    
    For i = 0 To UBound(a)
        Assert.Equals Choose(i + 1, "hi", "bye", "hi again"), a(i)
    Next i
End Sub
