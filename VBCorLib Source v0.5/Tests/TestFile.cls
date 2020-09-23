VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestFile
'
Option Explicit

Public Sub TestExists()
    Dim f As String
    
    f = App.Path & "\dummy test files\Exists.txt"
    Assert.IsTrue file.Exists(f), "Exists.txt should exist."
End Sub

Public Sub TestDirectory()
    Dim f As String
    
    f = App.Path
    Assert.IsFalse file.Exists(f), "A directory should never return true."
End Sub

Public Sub TestCopySourceNotFound()
    On Error GoTo errTrap
    file.Copy App.Path & "\dummy test files\notfound.txt", "\dummy test files\cannotbe.txt"
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."

errTrap:
    Dim ex As Exception
    Assert.IsTrue catch(ex), "An exception should be caught."
    Assert.IsTrue TypeOf ex Is FileNotFoundException, "Wrong exception type caught."
End Sub

Public Sub TestCopyNoOverwrite()
    On Error GoTo errTrap
    file.Copy App.Path & "\dummy test files\source.txt", App.Path & "\dummy test files\Exists.txt"
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim ex As Exception
    Assert.IsTrue catch(ex), "An exception should be caught."
    Assert.IsTrue TypeOf ex Is IOException, "Wrong exception type caught."
End Sub

Public Sub TestCopyOverExisting()
    TestExists
    
    On Error GoTo errTrap
    file.Copy App.Path & "\dummy test files\Source.txt", App.Path & "\dummy test files\Exists.txt", True
    Exit Sub
    
errTrap:
    Assert.Fail "No exception should be thrown."
End Sub

Public Sub TestCopyOverFolderNoOverwrite()
    On Error GoTo errTrap
    file.Copy App.Path & "\dummy test files\Source.txt", "c:\ATI"
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim ex As Exception
    Assert.IsTrue catch(ex), "An exception should be caught."
    Assert.IsTrue TypeOf ex Is IOException, "Wrong exception type caught."
End Sub

Public Sub TestGetCreationTime()
    Console.WriteLine vbCrLf & "--- VISUAL INSPECTION ---"
    Console.WriteLine "TestGetCreationTime: " & file.GetCreationTime(App.Path & "\dummy test files\source.txt").ToString
End Sub

Public Sub TestSetGetFileTime()
    Dim d As Date
    Dim f As String
    
    f = App.Path & "\dummy test files\source.txt"
    d = Now
    
    file.SetCreationTime f, d
    file.SetLastAccessTime f, d
    file.SetLastWriteTime f, d

    Assert.Equals d, file.GetCreationTime(f).ToOADate, "Wrong Creation time."
    Assert.Equals d, file.GetLastAccessTime(f).ToOADate, "Wrong LastAccess time."
    Assert.Equals d, file.GetLastWriteTime(f).ToOADate, "Wrong LastWrite time."
End Sub

Public Sub TestSetFileTimeUtc()
    Dim d As Date
    Dim f As String
    
    f = App.Path & "\dummy test files\source.txt"
    d = Now
    
    file.SetCreationTimeutc f, d
    file.SetLastAccessTimeutc f, d
    file.SetLastWriteTimeutc f, d

    Assert.Equals d, file.GetCreationTimeutc(f).ToOADate, "Wrong Creation time."
    Assert.Equals d, file.GetLastAccessTimeutc(f).ToOADate, "Wrong LastAccess time."
    Assert.Equals d, file.GetLastWriteTimeutc(f).ToOADate, "Wrong LastWrite time."
End Sub

Public Sub TestDelete()
    Dim f As String
    
    f = App.Path & "\dummy test files\deleteme.txt"
    If Not file.Exists(f) Then
        Open f For Binary As #1
        Close #1
    End If
    
    Assert.IsTrue file.Exists(f), "The deletable file does not exist."
    file.Delete f
    Assert.IsFalse file.Exists(f), "The deletable file should be deleted."
End Sub


