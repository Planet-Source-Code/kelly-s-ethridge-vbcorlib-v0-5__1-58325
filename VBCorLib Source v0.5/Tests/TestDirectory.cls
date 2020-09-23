VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestDirectory
'
Option Explicit

Public Sub TestExists()
    Assert.IsTrue directory.Exists(App.Path & "\Dummy Test Files")
    Assert.IsFalse directory.Exists(App.Path & "\TestDirectory.cls")
End Sub

Public Sub TestCreateDeleteDirectory()
    Dim f As String
    
    f = App.Path & "\dummy test files\Sub1\Sub2"
    If directory.Exists(App.Path & "\dummy test files\Sub1") Then directory.Delete App.Path & "\dummy test files\Sub1", True
    
    directory.CreateDirectory f
    Assert.IsTrue directory.Exists(f)
    directory.Delete App.Path & "\dummy test files\sub1", True
    Assert.IsFalse directory.Exists(f)
End Sub

Public Sub TestGetDirectories()
    Dim s() As String
    Dim i As Long
    
    s = directory.GetDirectories(App.Path)
    Assert.Equals 2, carray.GetLength(s)
    For i = 0 To 1
        Assert.IsIn s(i), "Dummy Test Files", "Folder"
    Next i
End Sub

Public Sub TestGetDirectoriesWithPattern()
    Dim s() As String
    
    s = directory.GetDirectories(App.Path, "fol*")
    Assert.Equals 1, carray.GetLength(s)
    Assert.Equals "Folder", s(0)
End Sub

Public Sub TestGetDirectoryRoot()
    Assert.Equals "c:\", directory.GetDirectoryRoot("c:\somepath\")
    Assert.Equals "\\server\share", directory.GetDirectoryRoot("\\server\share\some more folders\")
    Assert.Equals "C:\", directory.GetDirectoryRoot("something\more somthing")
End Sub

Public Sub TestGetFiles()
    Dim files() As String
    Dim i As Long
    
    files = directory.GetFiles("c:\")
    Console.WriteLine "Count: " & carray.GetLength(files)
    For i = 0 To UBound(files)
        Console.WriteLine "File: " & files(i)
    Next i
End Sub

Public Sub TestGetFileSystemEntries()
    Dim files() As String
    Dim i As Long
    
    files = directory.GetFileSystemEntries("c:\")
    Console.WriteLine "--- VISUAL INSPECTION ---"
    Console.WriteLine "TestGetFileSystemEntries"
    Console.WriteLine "Count: " & carray.GetLength(files)
    For i = 0 To UBound(files)
        Console.WriteLine "File: " & files(i)
    Next i
End Sub
