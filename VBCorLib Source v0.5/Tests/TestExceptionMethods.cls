VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestExceptionMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub TestThrowCatch()
    Dim original As New ArgumentException
    On Error GoTo errTrap
    Throw original
    On Error GoTo 0
    Assert.Fail "An exception should have been thrown."
    
errTrap:
    Dim ex As SystemException
    Assert.IsTrue Catch(ex), "An exception should have been caught."
    Assert.Equals original, ex, "Should have caught original instance."
End Sub

Public Sub TestNothingToCatch()
    Dim ex As Exception
    Assert.IsFalse Catch(ex), "No exception should have been caught."
End Sub

Public Sub TestCatchErrRaise()
    Dim ex As Exception
    
    On Error GoTo errTrap
    Err.Raise 383, "TestCatchErrRaise", "Test Catch."
    On Error GoTo 0
    Assert.Fail "An error should have been thrown."
errTrap:
    Assert.IsTrue Catch(ex, Err), "An error should have been caught."
    Assert.Equals 383, ex.HResult, "HResult should be 383."
    Assert.Equals "TestCatchErrRaise", ex.Source, "Source should be 'TestCatchErrRaise'."
    Assert.Equals "Test Catch.", ex.Message, "Message should be 'Test Catch.'"
End Sub
