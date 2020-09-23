VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestMappedFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestMappedFile
'
Option Explicit

'***
'*** This test required the vbcultures.nlp file to be where
'*** ever App.Path is located.
'***

Private file As MappedFile

Public Sub Setup()
    Set file = newmappedfile(App.Path & "\mappedfile.txt")
End Sub

Public Sub Teardown()
    file.CloseFile
End Sub

Public Sub TestMapping()
    Assert.IsTrue file.IsOpen, "Mapped file should be open."
    Assert.IsTrue file.BaseAddress <> 0, "Mapped file should have memory set."
End Sub

Public Sub TestGetCultureCount()
    Const CHAR_A As Byte = 65
    Const CHAR_B As Byte = 66
    
    Dim i As Long
    i = MemByte(file.BaseAddress)
    Assert.Equals CHAR_A, i, "Memory should contain 'A'."
    i = MemByte(file.BaseAddress + 1)
    Assert.Equals CHAR_B, i, "Memory should contain 'B'."
End Sub

