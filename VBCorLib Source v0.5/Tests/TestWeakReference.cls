VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestWeakReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestWeakReference
'
Option Explicit


Public Sub TestWeakness()
    Dim strong As StrongRef
    Dim weak As WeakRef
    
    Set strong = New StrongRef
    Set weak = New WeakRef
    Set strong.Reference = weak
    Set weak.Reference = strong
    
    Assert.IsNotNothing strong.Reference, "Target should not be Nothing."
    strong.AssertAlive
    Set weak = Nothing
    strong.AssertDead
    
End Sub
