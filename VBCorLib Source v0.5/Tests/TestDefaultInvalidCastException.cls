VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestDefInvalidCast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const EXPECTED_MESSAGE As String = "The specified cast is not valid."

Private ex As InvalidCastException

Public Sub Setup()
    Set ex = New InvalidCastException
End Sub

Public Sub TestMessage()
    Assert.Equals EXPECTED_MESSAGE, ex.Message, "Message should return 'An exception has occured."
End Sub

Public Sub TestNullInnerException()
    Assert.IsNothing ex.InnerException, "InnerException should be Nothing."
End Sub

Public Sub TestGetBaseExceptionReturnsSelf()
    Assert.Equals ex, ex.GetBaseException, "GetBaseException should return itself."
End Sub

Public Sub TestToString()
    Assert.Equals "VBCorLib.InvalidCastException: " & EXPECTED_MESSAGE, ex.ToString, "ToString return incorrect value."
End Sub

Public Sub TestEqualsSelf()
    Assert.IsTrue ex.Equals(ex), "Equals should return True."
End Sub

Public Sub TestNotEqualsNonObject()
    Assert.IsFalse ex.Equals("Not object value."), "Equals should return False."
End Sub

Public Sub TestNotEqualsNothing()
    Assert.IsFalse ex.Equals(Nothing), "Equals should return False."
End Sub

Public Sub TestNotEqualsOther()
    Dim ex2 As New Exception
    Assert.IsFalse ex.Equals(ex2), "Equals should return False."
End Sub

Public Sub TestGetHashCode()
    Dim ex2 As New Exception
    Assert.NotEqual ex2.GetHashCode, ex.GetHashCode, "GetHashCode should return different values."
End Sub



