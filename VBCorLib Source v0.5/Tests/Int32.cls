VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Int32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   Int32
'
'   This is a simple class to represent an object version of a 32bit number.
'
Option Explicit
Implements IComparable
Implements cObject

Public mValue As Long


Private Sub Class_Terminate()
    'Debug.Print "Terminated"
End Sub

Private Function cObject_Equals(value As Variant) As Boolean
    If IsObject(value) Then
        cObject_Equals = (value Is Me)
    End If
End Function

Private Function cObject_GetHashCode() As Long
    cObject_GetHashCode = ObjPtr(Me)
End Function

Private Function cObject_ToString() As String
    cObject_ToString = mValue
End Function

Private Function IComparable_CompareTo(value As Variant) As Long
    Dim l As Long
    If IsObject(value) Then
        Dim o As Int32
        Set o = value
        l = o.mValue
    Else
        l = value
    End If
    
    If mValue < l Then IComparable_CompareTo = -1: Exit Function
    If mValue > l Then IComparable_CompareTo = 1
End Function

