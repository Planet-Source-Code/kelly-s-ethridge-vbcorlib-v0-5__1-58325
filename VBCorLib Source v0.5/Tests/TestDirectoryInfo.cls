VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestDirectoryInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestDirectoryInfo
'
Option Explicit

Private info As DirectoryInfo

Public Sub TestFixtureSetup()
    Console.WriteLine ""
    Console.WriteLine "***** DIRECTORYINFO *****"
End Sub

Public Sub Setup()
    Set info = NewDirectoryInfo(App.Path)
End Sub

Public Sub TestAttributes()
    Assert.Equals 8208, info.Attributes
End Sub

Public Sub TestCreationTime()
    Console.WriteLine ""
    Console.WriteLine "--- VISUAL INSPECTION ---"
    Console.WriteLine "CreationTime: " & info.CreationTime.ToString
End Sub

Public Sub TestCreationTimeLocalVsUTC()
    Dim loc As cDateTime
    Dim utc As cDateTime
    Dim offset As TimeSpan
    
    Set loc = info.CreationTime
    Set utc = info.CreationTimeUtc
    Set offset = TimeZone.CurrentTimeZone.GetUtcOffset(loc)
    
    Console.WriteLine cString.Format("Local: {0}", loc)
    Console.WriteLine cString.Format("UTC: {0}", utc)
    Console.WriteLine cString.Format("Offset: {0}", offset)
    
    Assert.Equals offset.TotalHours, loc.Subtract(utc).TotalHours
End Sub

Public Sub TestName()
    Assert.Equals "Tests", info.Name
End Sub

Public Sub TestParentName()
    Dim p As DirectoryInfo
    
    Set p = info.Parent
    Assert.IsNotNothing p
    Assert.Equals "VBCorLib Source v0.5", p.Name
End Sub

Public Sub TestRoot()
    Dim r As DirectoryInfo
    
    Set r = info.Root
    Assert.IsNotNothing r, "No root information in path."
    Assert.Equals "c:\", LCase$(r.Name)
End Sub

Public Sub TestGetDirectories()
    Dim di() As DirectoryInfo
    Dim i As Long
    
    di = info.GetDirectories
    Assert.IsFalse cArray.IsNull(di), "Should not be a null array."
    
    Console.WriteLine ""
    Console.WriteLine "--- VISUAL INSPECTION ---"
    Console.WriteLine "GetDirectories"
    Console.WriteLine "Count: " & cArray.GetLength(di)
    For i = 0 To UBound(di)
        Console.WriteLine di(i).Name
    Next i
End Sub

Public Sub TestNoParent()
    Set info = Cor.NewDirectoryInfo("c:\")
    Assert.IsNothing info.Parent
    
    Set info = Cor.NewDirectoryInfo("\\server\share")
    Assert.IsNothing info.Parent
End Sub

Public Sub TestGetDirectoriesNoMatch()
    Dim di() As DirectoryInfo
    
    di = info.GetDirectories("abc123")
    Assert.IsFalse cArray.IsNull(di)
    Assert.Equals 1, cArray.GetLength(di)
    Assert.Equals Path.GetPathRoot(Path.GetFullPath("abc123")), di(0).ToString
End Sub

Public Sub TestGetFilesNoMatch()
    Dim fi() As FileInfo
    
    fi = info.GetFiles("abc123")
    Assert.IsFalse cArray.IsNull(fi)
    Assert.Equals 0, cArray.GetLength(fi)
End Sub
