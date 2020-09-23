VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestPath
'
Option Explicit

Public Sub TestFixtureSetup()
    Console.WriteLine ""
    Console.WriteLine "***** PATH CLASS *****"
End Sub

Public Sub TestChangeExtension()
    Assert.Equals "Filename.rec", Path.ChangeExtension("Filename.txt", "rec")
    Assert.Equals "c.:hello.", Path.ChangeExtension("c.:hello", "")
    Assert.Equals "hello", Path.ChangeExtension("hello.txt", Null)
    Assert.Equals "hello.d.d", Path.ChangeExtension("hello", "d.d")
    Assert.Equals "h.d.d", Path.ChangeExtension("h.ello", "d.d")
    Assert.Equals "h", Path.ChangeExtension("h.ello", Null)
    Assert.Equals "hello.txt", Path.ChangeExtension("hello", ".txt")
    Assert.Equals "hello.txt", Path.ChangeExtension("hello", "txt")
    Assert.Equals "", Path.ChangeExtension("", "")
End Sub

Public Sub TestIsPathRooted()
    Assert.IsTrue Path.IsPathRooted("c:\"), "'c:\' should be rooted."
    Assert.IsFalse Path.IsPathRooted("Abc\file"), "'Abc\fil' should not be rooted."
    Assert.IsTrue Path.IsPathRooted("\folder"), "'\folder' should be rooted."
    Assert.IsFalse Path.IsPathRooted(":\folder"), "':\folder' should not be rooted."
    Assert.IsFalse Path.IsPathRooted(""), "empty path should not be rooted."
    Assert.IsTrue Path.IsPathRooted("\\share"), "'\\share' should be rooted."
End Sub

Public Sub TestHasExtension()
    Assert.IsTrue Path.HasExtension("file.txt"), "'file.txt' should have an extension."
    Assert.IsTrue Path.HasExtension(".a"), "'.a' should have an extension."
    Assert.IsFalse Path.HasExtension("."), "'.' should not have an extension."
    Assert.IsFalse Path.HasExtension(".\file"), "'.\file' should not have an extension."
    Assert.IsFalse Path.HasExtension("file."), "'file.' should not have an extension."
    Assert.IsFalse Path.HasExtension("file"), "'file' should not have an extension."
End Sub

Public Sub TestCombine()
    Assert.Equals "a\b", Path.Combine("a", "b")
    Assert.Equals "\b", Path.Combine("a", "\b")
    Assert.Equals "a/b", Path.Combine("a/", "b")
    Assert.Equals "a:b", Path.Combine("a:", "b")
    Assert.Equals "c:\*.txt", Path.Combine("c:\", "*.txt")
    Assert.Equals "*.txt", Path.Combine("", "*.txt")
    Assert.Equals "", Path.Combine("", "")
    Assert.Equals "a", Path.Combine("a", "")
    Assert.Equals "b", Path.Combine("", "b")
End Sub

Public Sub TestGetDirectoryName()
    Assert.Equals "c:\", Path.GetDirectoryName("c:\abc")
    Assert.Equals "c:\abc", Path.GetDirectoryName("c:\abc\")
    Assert.IsNull Path.GetDirectoryName("c:\"), "'c:\' should return a null."
    Assert.Equals "", Path.GetDirectoryName("abc"), "'abc' should return an empty string."
    Assert.Equals "abc", Path.GetDirectoryName("abc\")
    Assert.IsNull Path.GetDirectoryName("\\s\a")
    Assert.Equals "\\s\e", Path.GetDirectoryName("\\s//e\")
End Sub

Public Sub TestGetExtension()
    Assert.Equals ".txt", Path.GetExtension("file.txt")
    Assert.Equals "", Path.GetExtension("file")
    Assert.Equals ".file", Path.GetExtension(".file")
    Assert.Equals "", Path.GetExtension(".")
    Assert.Equals "", Path.GetExtension("..")
    Assert.Equals "", Path.GetExtension("c.:txt")
End Sub

Public Sub TestGetFileName()
    Assert.Equals "file.txt", Path.GetFileName("file.txt")
    Assert.Equals "file.txt", Path.GetFileName("c:\file.txt")
    Assert.Equals "", Path.GetFileName("c:\")
    Assert.Equals "file.txt", Path.GetFileName("c:file.txt")
End Sub

Public Sub TestGetFileNameWithoutExtension()
    Assert.Equals "file", Path.GetFileNameWithoutExtension("file.txt")
    Assert.Equals "file", Path.GetFileNameWithoutExtension("c:\file.txt")
    Assert.Equals "", Path.GetFileNameWithoutExtension("c:\")
    Assert.Equals "file", Path.GetFileNameWithoutExtension("c:file.txt")
    Assert.Equals "file", Path.GetFileNameWithoutExtension("file")
    Assert.Equals "file", Path.GetFileNameWithoutExtension("c:file")
    Assert.Equals "", Path.GetFileNameWithoutExtension("c:.file")
End Sub

Public Sub TestGetFullPath()
    ' The output is machine specific. The result will prepend the
    ' current directory to the file.
    Console.WriteLine ""
    Console.WriteLine "--- VISUAL INSPECTION ---"
    Console.WriteLine "GetFullPath (file.txt): " & Path.GetFullPath("file.txt")
    
    Dim ex As Exception
    On Error Resume Next
    Path.GetFullPath "\\server"
    On Error GoTo 0
    Assert.IsTrue catch(ex), "'\\server' should throw an exception."
    
    On Error Resume Next
    Path.GetFullPath "\\server\"
    On Error GoTo 0
    Assert.IsTrue catch(ex), "'\\server\' should throw an exception."
    
    Assert.Equals "\\server\share", Path.GetFullPath("//server\share")
    Assert.Equals "\\server\share\", Path.GetFullPath("\\server\share\")
End Sub

Public Sub TestGetPathRoot()
    Assert.Equals "\", Path.GetPathRoot("\mydir\")
    Assert.Equals "c:\", Path.GetPathRoot("c:")
    Assert.Equals "", Path.GetPathRoot("file.txt")
    Assert.Equals "\\server", Path.GetPathRoot("\\server")
    Assert.Equals "\\server\", Path.GetPathRoot("\\server\")
    Assert.Equals "\\server\share", Path.GetPathRoot("\\server\share")
    Assert.Equals "\\server\share", Path.GetPathRoot("\\server\share\dir")
End Sub

Public Sub TestGetTempFileName()
    Dim s As String
    
    s = Path.GetTempFileName
    Assert.IsTrue Len(s) > 0
    Assert.Equals 0, FileLen(s)
    Console.WriteLine ""
    Console.WriteLine "--- VISUAL INSPECTION ---"
    Console.WriteLine "GetTempFileName: " & s
    Kill s
End Sub

    ' the output is user specific, so visual inspection is suggested.
Public Sub TestGetTempPath()
    Console.WriteLine ""
    Console.WriteLine "--- VISUAL INSPECTION ---"
    Console.WriteLine "GetTempPath: " & Path.GetTempPath
End Sub
